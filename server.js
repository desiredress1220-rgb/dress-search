const express = require('express');
const multer = require('multer');
const cookieParser = require('cookie-parser');
const jwt = require('jsonwebtoken');
const crypto = require('crypto');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// ============================================================
// Config
// ============================================================
const PORT = process.env.PORT || 3000;
const APP_PASSWORD = process.env.APP_PASSWORD || 'bestwishes';
const FEISHU_APP_ID = process.env.FEISHU_APP_ID || '';
const FEISHU_APP_SECRET = process.env.FEISHU_APP_SECRET || '';
const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID || '';
const VERTEX_PROJECT = 'desire-dress-search';
const VERTEX_LOCATION = 'us-central1';
const VERTEX_MODEL = 'multimodalembedding@001';

let GCP_CREDENTIALS = {};
try {
  GCP_CREDENTIALS = JSON.parse(process.env.GCP_CREDENTIALS || '{}');
} catch (e) {
  console.error('Failed to parse GCP_CREDENTIALS:', e.message);
}

const BITABLE_APP_TOKEN = 'Uj4ubusyrast2ds8H2ZcacqqnVh';
const BITABLE_TABLE_ID = 'tblo0edlld8OgL4q';

// ============================================================
// State
// ============================================================
let searchReady = false;
let loadingProgress = '等待启动...';
let loadError = null;
let styleEmbeddings = {};
let styleMetadata = {};
let metadataList = [];         // full metadata array
let embDim = 1408;
let thumbnailsFolderId = null; // Drive folder ID for thumbnails
const thumbCache = {};         // index -> Buffer cache

// ============================================================
// GCP Auth
// ============================================================
let cachedToken = null;
let tokenExpiry = 0;

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiry - 60000) return cachedToken;
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    iss: GCP_CREDENTIALS.client_email,
    scope: 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/cloud-platform',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  };
  const assertion = jwt.sign(payload, GCP_CREDENTIALS.private_key, { algorithm: 'RS256' });
  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${assertion}`
  });
  if (!resp.ok) throw new Error(`Token exchange failed: ${resp.status}`);
  const data = await resp.json();
  cachedToken = data.access_token;
  tokenExpiry = Date.now() + data.expires_in * 1000;
  return cachedToken;
}

// ============================================================
// Google Drive
// ============================================================
async function driveListFolder(folderId, pageSize = 100) {
  const token = await getAccessToken();
  const q = encodeURIComponent(`'${folderId}' in parents and trashed=false`);
  const resp = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,size,mimeType)&pageSize=${pageSize}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!resp.ok) throw new Error(`Drive list failed: ${resp.status} ${await resp.text()}`);
  return (await resp.json()).files || [];
}

async function driveDownload(fileId) {
  const token = await getAccessToken();
  const resp = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!resp.ok) throw new Error(`Drive download failed: ${resp.status}`);
  return resp;
}

async function driveSearchFile(folderId, fileName) {
  const token = await getAccessToken();
  const q = encodeURIComponent(`'${folderId}' in parents and name='${fileName}' and trashed=false`);
  const resp = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&pageSize=1`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!resp.ok) return null;
  const data = await resp.json();
  return data.files?.[0] || null;
}

async function loadDataFromDrive() {
  loadingProgress = '正在列出 Google Drive 文件...';
  console.log(loadingProgress);
  const files = await driveListFolder(DRIVE_FOLDER_ID);
  console.log('Found files:', files.map(f => `${f.name} (${f.size || f.mimeType})`).join(', '));

  const find = (name) => files.find(f => f.name === name);
  const metaFile = find('metadata.json');
  const embFile = find('embeddings.bin');
  const dimsFile = find('embeddings_dims.json');
  const thumbFolder = files.find(f => f.name === 'thumbnails' && f.mimeType === 'application/vnd.google-apps.folder');

  if (!metaFile) throw new Error('metadata.json not found in Drive folder');
  if (!embFile) throw new Error('embeddings.bin not found in Drive folder');

  if (thumbFolder) {
    thumbnailsFolderId = thumbFolder.id;
    console.log('Thumbnails folder found:', thumbnailsFolderId);
  } else {
    console.log('No thumbnails folder found in Drive');
  }

  loadingProgress = '正在下载 metadata.json...';
  console.log(loadingProgress);
  const metaResp = await driveDownload(metaFile.id);
  const metadata = await metaResp.json();
  console.log(`Loaded ${metadata.length} image records`);

  let dims = { dimensions: 1408 };
  if (dimsFile) {
    const dimsResp = await driveDownload(dimsFile.id);
    dims = await dimsResp.json();
  }
  embDim = dims.dimensions || 1408;

  loadingProgress = `正在下载 embeddings.bin (${Math.round(embFile.size / 1024 / 1024)}MB)...`;
  console.log(loadingProgress);
  const embResp = await driveDownload(embFile.id);
  const embBuffer = Buffer.from(await embResp.arrayBuffer());
  const embeddings = new Float32Array(embBuffer.buffer, embBuffer.byteOffset, embBuffer.byteLength / 4);
  console.log(`Loaded ${embeddings.length} float values (${embeddings.length / embDim} vectors)`);

  return { metadata, embeddings };
}

// ============================================================
// Compute style averages
// ============================================================
function computeStyleAverages(metadata, embeddings) {
  loadingProgress = '正在计算款式平均向量...';
  console.log(loadingProgress);
  metadataList = metadata;

  const groups = {};
  for (let i = 0; i < metadata.length; i++) {
    const img = metadata[i];
    const style = img.style || img.style_number || 'unknown';
    if (!groups[style]) {
      groups[style] = { indices: [], series: img.series || '', images: [] };
    }
    groups[style].indices.push(i);
    if (groups[style].images.length < 3) {
      groups[style].images.push({ index: i, filename: img.filename || img.name || '' });
    }
  }

  let validCount = 0;
  for (const [style, group] of Object.entries(groups)) {
    const avg = new Float32Array(embDim);
    let hasData = false;
    for (const idx of group.indices) {
      const offset = idx * embDim;
      if (offset + embDim > embeddings.length) continue;
      hasData = true;
      for (let d = 0; d < embDim; d++) avg[d] += embeddings[offset + d];
    }
    if (!hasData) continue;
    const n = group.indices.length;
    for (let d = 0; d < embDim; d++) avg[d] /= n;
    let norm = 0;
    for (let d = 0; d < embDim; d++) norm += avg[d] * avg[d];
    norm = Math.sqrt(norm);
    if (norm > 0) for (let d = 0; d < embDim; d++) avg[d] /= norm;

    styleEmbeddings[style] = avg;
    styleMetadata[style] = {
      count: n,
      series: group.series,
      sampleImages: group.images,
      thumbIndex: group.indices[0],  // first image is usually the main/front view
      thumbIndices: group.indices.slice(0, 6)  // first 6 images for thumbnail row
    };
    validCount++;
  }
  console.log(`Computed averages for ${validCount} styles`);
}

// ============================================================
// Vertex AI
// ============================================================
async function getQueryEmbedding(imageBuffer) {
  const token = await getAccessToken();
  const base64 = imageBuffer.toString('base64');
  const url = `https://${VERTEX_LOCATION}-aiplatform.googleapis.com/v1/projects/${VERTEX_PROJECT}/locations/${VERTEX_LOCATION}/publishers/google/models/${VERTEX_MODEL}:predict`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      instances: [{
        image: { bytesBase64Encoded: base64 },
        text: 'evening gown formal dress close-up product photo'
      }]
    })
  });
  if (!resp.ok) throw new Error(`Vertex AI error ${resp.status}: ${await resp.text()}`);
  const data = await resp.json();
  if (!data.predictions?.[0]?.imageEmbedding) throw new Error('No embedding in response');
  return new Float32Array(data.predictions[0].imageEmbedding);
}

// ============================================================
// Search
// ============================================================
function cosineSim(a, b) {
  let dot = 0, nA = 0, nB = 0;
  for (let i = 0; i < a.length; i++) { dot += a[i]*b[i]; nA += a[i]*a[i]; nB += b[i]*b[i]; }
  return dot / (Math.sqrt(nA) * Math.sqrt(nB));
}

function searchStyles(queryEmb, topK = 5) {
  const results = [];
  for (const [style, emb] of Object.entries(styleEmbeddings)) {
    results.push({ style, score: cosineSim(queryEmb, emb), ...styleMetadata[style] });
  }
  results.sort((a, b) => b.score - a.score);
  return results.slice(0, topK);
}

// ============================================================
// Feishu Bitable
// ============================================================
let feishuToken = null;
let feishuTokenExpiry = 0;

async function getFeishuToken() {
  if (feishuToken && Date.now() < feishuTokenExpiry) return feishuToken;
  const resp = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ app_id: FEISHU_APP_ID, app_secret: FEISHU_APP_SECRET })
  });
  const data = await resp.json();
  if (data.code !== 0) throw new Error(`Feishu auth: ${data.msg}`);
  feishuToken = data.tenant_access_token;
  feishuTokenExpiry = Date.now() + (data.expire - 300) * 1000;
  return feishuToken;
}

async function lookupPrice(styleNumber) {
  if (!FEISHU_APP_SECRET) return null;
  try {
    const token = await getFeishuToken();
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${BITABLE_APP_TOKEN}/tables/${BITABLE_TABLE_ID}/records/search`;
    const resp = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filter: {
          conjunction: 'and',
          conditions: [{ field_name: 'ITEM NO', operator: 'is', value: [styleNumber] }]
        },
        page_size: 1
      })
    });
    const data = await resp.json();
    console.log(`Feishu lookup [${styleNumber}]:`, JSON.stringify(data.data?.items?.[0]?.fields || data).substring(0, 300));
    if (data.data?.items?.length > 0) {
      const f = data.data.items[0].fields;
      return {
        wholesale: f['WHOLESALE PRICE USD'] ?? f['批发价'] ?? null,
        retail: f['RETAILER PRICE USD'] ?? f['零售价'] ?? null
      };
    }
    return null;
  } catch (e) {
    console.error('Feishu error:', e.message);
    return null;
  }
}

// ============================================================
// Express
// ============================================================
app.use(express.json());
app.use(cookieParser());

const AUTH_SECRET = 'dress-search-v1';
function makeAuthToken(pw) { return crypto.createHmac('sha256', AUTH_SECRET).update(pw).digest('hex'); }

function authCheck(req, res, next) {
  if (req.path === '/api/login') return next();
  const token = req.cookies?.auth;
  const expected = makeAuthToken(APP_PASSWORD);
  if (token === expected) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: '请先登录' });
  next();
}

app.use(authCheck);
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/login', (req, res) => {
  if (req.body.password === APP_PASSWORD) {
    res.cookie('auth', makeAuthToken(APP_PASSWORD), { httpOnly: true, sameSite: 'lax', maxAge: 90*24*60*60*1000 });
    return res.json({ success: true });
  }
  res.status(401).json({ error: '密码错误' });
});

app.post('/api/logout', (req, res) => { res.clearCookie('auth'); res.json({ success: true }); });

app.get('/api/status', (req, res) => {
  const authed = req.cookies?.auth === makeAuthToken(APP_PASSWORD);
  res.json({ authenticated: authed, ready: searchReady, progress: loadingProgress, error: loadError, styles: Object.keys(styleEmbeddings).length, hasThumbnails: !!thumbnailsFolderId });
});

// Thumbnail proxy - downloads from Drive on demand with caching
app.get('/api/thumb/:index', async (req, res) => {
  if (!thumbnailsFolderId) return res.status(404).send('No thumbnails');
  const idx = req.params.index;

  // Check cache
  if (thumbCache[idx]) {
    res.set('Content-Type', 'image/jpeg');
    res.set('Cache-Control', 'public, max-age=86400');
    return res.send(thumbCache[idx]);
  }

  try {
    const file = await driveSearchFile(thumbnailsFolderId, `${idx}.jpg`);
    if (!file) return res.status(404).send('Thumbnail not found');
    const dlResp = await driveDownload(file.id);
    const buf = Buffer.from(await dlResp.arrayBuffer());
    // Cache (limit to 500 thumbnails in memory ~25MB)
    if (Object.keys(thumbCache).length < 500) thumbCache[idx] = buf;
    res.set('Content-Type', 'image/jpeg');
    res.set('Cache-Control', 'public, max-age=86400');
    res.send(buf);
  } catch (e) {
    console.error('Thumb error:', e.message);
    res.status(500).send('Error');
  }
});

// Search
app.post('/api/search', upload.single('image'), async (req, res) => {
  if (req.cookies?.auth !== makeAuthToken(APP_PASSWORD)) return res.status(401).json({ error: '请先登录' });
  if (!searchReady) return res.status(503).json({ error: '数据加载中，请稍候...' });
  if (!req.file) return res.status(400).json({ error: '请上传图片' });

  try {
    console.log(`Search: ${req.file.originalname} (${Math.round(req.file.size/1024)}KB)`);
    const queryEmb = await getQueryEmbedding(req.file.buffer);
    const results = searchStyles(queryEmb, 5);

    // Lookup prices in parallel for speed
    const pricePromises = results.map(r => lookupPrice(r.style));
    const prices = await Promise.all(pricePromises);
    results.forEach((r, i) => {
      r.matchPercent = Math.round(r.score * 100);
      r.thumbIndex = r.thumbIndex ?? (r.sampleImages?.[0]?.index ?? null);
      r.thumbIndices = r.thumbIndices || [];
      if (prices[i]) {
        r.wholesalePrice = prices[i].wholesale;
        r.retailPrice = prices[i].retail;
      }
    });

    console.log(`Results: ${results.map(r => `${r.style}(${r.matchPercent}%)`).join(', ')}`);
    res.json({ results, hasThumbnails: !!thumbnailsFolderId });
  } catch (e) {
    console.error('Search error:', e);
    res.status(500).json({ error: `搜索失败: ${e.message}` });
  }
});

// Reload endpoint for n8n
app.post('/api/reload', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  searchReady = false; styleEmbeddings = {}; styleMetadata = {}; loadError = null;
  try { await loadAndInit(); res.json({ success: true, styles: Object.keys(styleEmbeddings).length }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ============================================================
// Startup
// ============================================================
async function loadAndInit() {
  const { metadata, embeddings } = await loadDataFromDrive();
  computeStyleAverages(metadata, embeddings);
  searchReady = true;
  loadingProgress = '就绪';
  console.log(`✅ Search engine ready — ${Object.keys(styleEmbeddings).length} styles loaded`);
}

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  if (!GCP_CREDENTIALS.private_key) { loadError = 'Missing GCP_CREDENTIALS'; console.error(loadError); return; }
  if (!DRIVE_FOLDER_ID) { loadError = 'Missing DRIVE_FOLDER_ID'; console.error(loadError); return; }
  loadAndInit().catch(e => { loadError = e.message; console.error('Init failed:', e); });
});
