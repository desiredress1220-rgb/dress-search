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
const PRICE_CACHE_TTL_MS = 15 * 60 * 1000;

// ============================================================
// State
// ============================================================
let searchReady = false;
let loadingProgress = '等待启动...';
let loadError = null;
let styleEmbeddings = {};
let styleMetadata = {};
let metadataList = [];
let imageEmbeddings = null; // Int8 quantized embeddings kept in memory for per-image search
let imageNorms = [];
let embDim = 1408;
let thumbnailsFolderId = null; // Drive folder ID for thumbnails
const thumbCache = {};         // index -> Buffer cache
const indexFileIds = { metadata: null, embeddings: null, dims: null };
const indexJobs = new Map();

function textFieldValue(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value);
  if (Array.isArray(value)) return value.map(textFieldValue).join('');
  if (typeof value === 'object') return value.text || value.name || value.value || '';
  return String(value);
}

function normalizeStyleId(value) {
  return textFieldValue(value)
    .replace(/\.[^.]+$/, '')
    .replace(/\s*\([^)]*\)\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function metadataStyleId(img) {
  return normalizeStyleId(img.style || img.style_number || img.item_no || img.name || 'unknown');
}

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
    scope: 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/cloud-platform',
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

async function driveUpdateFile(fileId, bodyBuffer, mimeType) {
  const token = await getAccessToken();
  const resp = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': mimeType },
    body: bodyBuffer
  });
  if (!resp.ok) throw new Error(`Drive update failed: ${resp.status} ${await resp.text()}`);
  return resp.json();
}

async function driveUploadToFolder(folderId, fileName, bodyBuffer, mimeType) {
  const token = await getAccessToken();
  const metadata = { name: fileName, parents: [folderId] };
  const boundary = `dress-search-${Date.now()}`;
  const delimiter = `--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;
  const multipartBody = Buffer.concat([
    Buffer.from(delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) + '\r\n'),
    Buffer.from(delimiter + `Content-Type: ${mimeType}\r\n\r\n`),
    bodyBuffer,
    Buffer.from(closeDelimiter)
  ]);
  const resp = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': `multipart/related; boundary=${boundary}` },
    body: multipartBody
  });
  if (!resp.ok) throw new Error(`Drive upload failed: ${resp.status} ${await resp.text()}`);
  return resp.json();
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
  indexFileIds.metadata = metaFile.id;
  indexFileIds.embeddings = embFile.id;
  indexFileIds.dims = dimsFile?.id || null;

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
  let embBuffer = Buffer.from(await embResp.arrayBuffer());
  let usableBytes = embBuffer.byteLength - (embBuffer.byteLength % 4);
  let floatCount = usableBytes / 4;
  const vectorCount = Math.floor(floatCount / embDim);
  const usableFloatCount = vectorCount * embDim;
  usableBytes = usableFloatCount * 4;
  if (usableBytes !== embBuffer.byteLength) {
    console.warn(`Trimming embeddings.bin from ${embBuffer.byteLength} to ${usableBytes} bytes`);
    embBuffer = embBuffer.subarray(0, usableBytes);
  }
  const embeddings = new Float32Array(embBuffer.buffer, embBuffer.byteOffset, usableFloatCount);
  const usableMetadata = metadata.slice(0, Math.min(metadata.length, vectorCount));
  if (usableMetadata.length !== metadata.length || usableFloatCount !== floatCount) {
    console.warn(`Index size mismatch: metadata=${metadata.length}, completeVectors=${vectorCount}, usingMetadata=${usableMetadata.length}`);
  }
  console.log(`Loaded ${embeddings.length} float values (${vectorCount} complete vectors)`);

  return { metadata: usableMetadata, embeddings };
}

// ============================================================
// Compute style averages
// ============================================================
function computeStyleAverages(metadata, embeddings) {
  loadingProgress = '正在计算款式平均向量...';
  console.log(loadingProgress);
  metadataList = metadata;
  imageEmbeddings = new Int8Array(embeddings.length);
  imageNorms = new Float32Array(metadata.length);

  const groups = {};
  for (let i = 0; i < metadata.length; i++) {
    const img = metadata[i];
    const style = metadataStyleId(img);
    if (!groups[style]) {
      groups[style] = { indices: [], series: img.series || '' };
    }
    groups[style].indices.push(i);

    const offset = i * embDim;
    let norm = 0;
    if (offset + embDim <= embeddings.length) {
      for (let d = 0; d < embDim; d++) {
        const value = embeddings[offset + d];
        const quantized = Math.max(-127, Math.min(127, Math.round(value * 127)));
        imageEmbeddings[offset + d] = quantized;
        const restored = quantized / 127;
        norm += restored * restored;
      }
    }
    imageNorms[i] = Math.sqrt(norm);
  }

  let validCount = 0;
  for (const [style, group] of Object.entries(groups)) {
    const avg = new Float32Array(embDim);
    let validImages = 0;
    for (const idx of group.indices) {
      const offset = idx * embDim;
      if (offset + embDim > embeddings.length) continue;
      validImages++;
      for (let d = 0; d < embDim; d++) avg[d] += embeddings[offset + d];
    }
    if (!validImages) continue;
    const n = validImages;
    for (let d = 0; d < embDim; d++) avg[d] /= n;
    let norm = 0;
    for (let d = 0; d < embDim; d++) norm += avg[d] * avg[d];
    norm = Math.sqrt(norm);
    if (norm > 0) for (let d = 0; d < embDim; d++) avg[d] /= norm;

    styleEmbeddings[style] = avg;
    styleMetadata[style] = {
      count: n,
      series: group.series,
      thumbIndex: group.indices[0],
      thumbIndices: group.indices.slice(0, 6)
    };
    validCount++;
  }
  console.log(`Computed averages for ${validCount} styles`);
}

// ============================================================
// Gemini — Analyze dress features before embedding
// ============================================================
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const USE_GEMINI_ANALYSIS = process.env.USE_GEMINI_ANALYSIS === 'true';

async function analyzeDressWithGemini(imageBuffer) {
  if (!USE_GEMINI_ANALYSIS || !GEMINI_API_KEY) return 'evening gown formal dress product photo';
  
  try {
    const base64 = imageBuffer.toString('base64');
    const resp = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [
              { inlineData: { mimeType: 'image/jpeg', data: base64 } },
              { text: 'Describe this dress in ONE short sentence focusing only on: silhouette (mermaid/A-line/ballgown/sheath), color, neckline, fabric texture, and any embellishments (beading/sequins/lace/feathers). Do NOT mention the model, background, or setting. Reply with only the description, nothing else.' }
            ]
          }],
          generationConfig: { maxOutputTokens: 80, temperature: 0.1 }
        })
      }
    );
    const data = await resp.json();
    const desc = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
    console.log('Gemini dress analysis:', desc);
    return desc || 'evening gown formal dress product photo';
  } catch (e) {
    console.error('Gemini analysis error:', e.message);
    return 'evening gown formal dress product photo';
  }
}

// ============================================================
// Vertex AI
// ============================================================
async function getQueryEmbedding(imageBuffer) {
  const token = await getAccessToken();
  const base64 = imageBuffer.toString('base64');
  const dressText = await analyzeDressWithGemini(imageBuffer);
  const url = `https://${VERTEX_LOCATION}-aiplatform.googleapis.com/v1/projects/${VERTEX_PROJECT}/locations/${VERTEX_LOCATION}/publishers/google/models/${VERTEX_MODEL}:predict`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      instances: [{
        image: { bytesBase64Encoded: base64 },
        text: `${dressText}. Match this dress to mannequin product photos. Ignore the model, face, pose, background, lighting, and scene. Focus on silhouette, neckline, straps, slit, beading, sequins, lace pattern, waist detail, train, and color.`
      }]
    })
  });
  if (!resp.ok) throw new Error(`Vertex AI error ${resp.status}: ${await resp.text()}`);
  const data = await resp.json();
  if (!data.predictions?.[0]?.imageEmbedding) throw new Error('No embedding in response');
  return new Float32Array(data.predictions[0].imageEmbedding);
}

async function getIndexEmbedding(imageBuffer) {
  const token = await getAccessToken();
  const base64 = imageBuffer.toString('base64');
  const url = `https://${VERTEX_LOCATION}-aiplatform.googleapis.com/v1/projects/${VERTEX_PROJECT}/locations/${VERTEX_LOCATION}/publishers/google/models/${VERTEX_MODEL}:predict`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      instances: [{
        image: { bytesBase64Encoded: base64 },
        text: 'mannequin product photo of an evening gown dress, clear front view, dress details'
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
function searchStyles(queryEmb, topK = 3) {
  if (imageEmbeddings && imageEmbeddings.length && metadataList.length) {
    return searchStylesByImages(queryEmb, topK);
  }

  const results = [];
  for (const [style, emb] of Object.entries(styleEmbeddings)) {
    let dot = 0, nA = 0, nB = 0;
    for (let i = 0; i < embDim; i++) {
      dot += queryEmb[i] * emb[i];
      nA += queryEmb[i] * queryEmb[i];
      nB += emb[i] * emb[i];
    }
    const score = dot / (Math.sqrt(nA) * Math.sqrt(nB));
    const sm = styleMetadata[style];
    results.push({ style, score, count: sm.count, series: sm.series, thumbIndex: sm.thumbIndex, thumbIndices: sm.thumbIndices });
  }
  results.sort((a, b) => b.score - a.score);
  return results.slice(0, topK);
}

function searchStylesByImages(queryEmb, topK = 5) {
  let queryNorm = 0;
  for (let d = 0; d < embDim; d++) queryNorm += queryEmb[d] * queryEmb[d];
  queryNorm = Math.sqrt(queryNorm);
  if (!queryNorm) return [];

  const styleScores = new Map();
  for (let idx = 0; idx < metadataList.length; idx++) {
    const offset = idx * embDim;
    const imageNorm = imageNorms[idx];
    if (!imageNorm || offset + embDim > imageEmbeddings.length) continue;

    let dot = 0;
    for (let d = 0; d < embDim; d++) dot += queryEmb[d] * (imageEmbeddings[offset + d] / 127);
    const score = dot / (queryNorm * imageNorm);
    const img = metadataList[idx];
    const style = metadataStyleId(img);
    const current = styleScores.get(style) || {
      style,
      series: img.series || '',
      count: styleMetadata[style]?.count || 0,
      imageCount: 0,
      topScores: [],
      topMatches: [],
      thumbIndex: idx
    };

    current.imageCount++;
    insertTopScore(current.topScores, score, 3);
    insertTopMatch(current.topMatches, { idx, score }, 6);
    if (score > (current.bestScore ?? -Infinity)) {
      current.bestScore = score;
      current.thumbIndex = idx;
      if (img.series) current.series = img.series;
    }
    styleScores.set(style, current);
  }

  const results = [];
  for (const item of styleScores.values()) {
    const topAverage = item.topScores.reduce((sum, score) => sum + score, 0) / item.topScores.length;
    const styleAverageScore = cosine(queryEmb, styleEmbeddings[item.style]);
    const score = 0.62 * item.bestScore + 0.28 * topAverage + 0.10 * styleAverageScore;
    results.push({
      style: item.style,
      score,
      count: item.count || item.imageCount,
      series: item.series,
      thumbIndex: item.thumbIndex,
      thumbIndices: item.topMatches.map(match => match.idx)
    });
  }

  results.sort((a, b) => b.score - a.score);
  return results.slice(0, topK);
}

function insertTopScore(scores, score, limit) {
  let inserted = false;
  for (let i = 0; i < scores.length; i++) {
    if (score > scores[i]) {
      scores.splice(i, 0, score);
      inserted = true;
      break;
    }
  }
  if (!inserted) scores.push(score);
  if (scores.length > limit) scores.length = limit;
}

function insertTopMatch(matches, match, limit) {
  let inserted = false;
  for (let i = 0; i < matches.length; i++) {
    if (match.score > matches[i].score) {
      matches.splice(i, 0, match);
      inserted = true;
      break;
    }
  }
  if (!inserted) matches.push(match);
  if (matches.length > limit) matches.length = limit;
}

function cosine(a, b) {
  if (!a || !b) return 0;
  let dot = 0, nA = 0, nB = 0;
  for (let i = 0; i < embDim; i++) {
    dot += a[i] * b[i];
    nA += a[i] * a[i];
    nB += b[i] * b[i];
  }
  return nA && nB ? dot / (Math.sqrt(nA) * Math.sqrt(nB)) : 0;
}

function bestThumbIndices(matches) {
  return matches
    .sort((a, b) => b.score - a.score)
    .slice(0, 6)
    .map(item => item.idx);
}

function appendImageToMemory(record, embedding) {
  const idx = metadataList.length;
  const quantized = new Int8Array(embDim);
  let norm = 0;
  for (let d = 0; d < embDim; d++) {
    const value = embedding[d];
    const q = Math.max(-127, Math.min(127, Math.round(value * 127)));
    quantized[d] = q;
    const restored = q / 127;
    norm += restored * restored;
  }

  const nextImageEmbeddings = new Int8Array(imageEmbeddings.length + embDim);
  nextImageEmbeddings.set(imageEmbeddings);
  nextImageEmbeddings.set(quantized, imageEmbeddings.length);
  imageEmbeddings = nextImageEmbeddings;

  const nextImageNorms = new Float32Array(imageNorms.length + 1);
  nextImageNorms.set(imageNorms);
  nextImageNorms[idx] = Math.sqrt(norm);
  imageNorms = nextImageNorms;

  metadataList = metadataList.concat(record);

  const existingMeta = styleMetadata[record.style];
  if (existingMeta) {
    existingMeta.count += 1;
    existingMeta.thumbIndices = [idx].concat(existingMeta.thumbIndices || []).slice(0, 6);
  } else {
    styleMetadata[record.style] = {
      count: 1,
      series: record.series || '',
      thumbIndex: idx,
      thumbIndices: [idx]
    };
    styleEmbeddings[record.style] = embedding;
  }
}

async function addImageToIndex({ imageBuffer, fileName, style, series }) {
  if (!searchReady || !imageEmbeddings) throw new Error('Search index is not ready');
  if (!indexFileIds.metadata || !indexFileIds.embeddings) throw new Error('Index file ids are not loaded');

  const normalizedStyle = normalizeStyleId(style || fileName);
  if (!normalizedStyle) throw new Error('Missing style number');

  const existing = metadataList.find(img =>
    normalizeStyleId(img.name || img.fileName || '') === normalizeStyleId(fileName) ||
    (img.driveName && normalizeStyleId(img.driveName) === normalizeStyleId(fileName))
  );
  if (existing) return { added: false, reason: 'already_exists', style: normalizedStyle };

  const embedding = await getIndexEmbedding(imageBuffer);
  if (embedding.length !== embDim) throw new Error(`Unexpected embedding dimension ${embedding.length}, expected ${embDim}`);

  const idx = metadataList.length;
  const record = {
    style: normalizedStyle,
    style_number: normalizedStyle,
    series: series || '',
    name: fileName,
    fileName,
    addedAt: new Date().toISOString()
  };

  const nextMetadata = metadataList.concat(record);

  const embResp = await driveDownload(indexFileIds.embeddings);
  const embBuffer = Buffer.from(await embResp.arrayBuffer());
  const embeddingBuffer = Buffer.from(embedding.buffer, embedding.byteOffset, embedding.byteLength);
  const nextEmbeddings = Buffer.concat([embBuffer, embeddingBuffer]);

  await driveUpdateFile(indexFileIds.embeddings, nextEmbeddings, 'application/octet-stream');
  await driveUpdateFile(indexFileIds.metadata, Buffer.from(JSON.stringify(nextMetadata)), 'application/json');

  appendImageToMemory(record, embedding);

  if (thumbnailsFolderId) {
    try {
      await driveUploadToFolder(thumbnailsFolderId, `${idx}.jpg`, imageBuffer, 'image/jpeg');
      thumbCache[idx] = imageBuffer;
    } catch (e) {
      console.error('Thumbnail upload error:', e.message);
    }
  }

  return { added: true, style: normalizedStyle, index: idx, images: metadataList.length };
}

// ============================================================
// Feishu Bitable
// ============================================================
let feishuToken = null;
let feishuTokenExpiry = 0;
let priceCache = new Map();
let priceCacheLoadedAt = 0;

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

async function lookupPriceDirect(styleNumber) {
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

async function lookupPrice(styleNumber) {
  if (!FEISHU_APP_SECRET) return null;
  try {
    await refreshPriceCacheIfNeeded();
    return priceCache.get(normalizeStyleId(styleNumber)) || null;
  } catch (e) {
    console.error('Feishu error:', e.message);
    return null;
  }
}

async function refreshPriceCacheIfNeeded(force = false) {
  if (!FEISHU_APP_SECRET) return;
  if (!force && priceCache.size && Date.now() - priceCacheLoadedAt < PRICE_CACHE_TTL_MS) return;

  const token = await getFeishuToken();
  const nextCache = new Map();
  let pageToken = '';
  let loaded = 0;
  let hasMore = false;

  do {
    const url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${BITABLE_APP_TOKEN}/tables/${BITABLE_TABLE_ID}/records/search${pageToken ? `?page_token=${encodeURIComponent(pageToken)}` : ''}`;
    const resp = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        field_names: ['ITEM NO', 'WHOLESALE PRICE USD', 'SHIPPING COST USD', 'TOTAL AMOUNT USD', 'RETAILER PRICE USD', 'COLOR', '系列'],
        page_size: 500
      })
    });
    const data = await resp.json();
    if (data.code !== 0) throw new Error(`Feishu records: ${data.msg || data.code}`);

    for (const item of data.data?.items || []) {
      const f = item.fields || {};
      const key = normalizeStyleId(f['ITEM NO']);
      if (!key) continue;
      if (!nextCache.has(key)) {
        nextCache.set(key, {
          wholesale: f['WHOLESALE PRICE USD'] ?? null,
          retail: f['RETAILER PRICE USD'] ?? null,
          shipping: f['SHIPPING COST USD'] ?? null,
          total: f['TOTAL AMOUNT USD'] ?? null,
          color: textFieldValue(f.COLOR).trim(),
          series: textFieldValue(f['系列']).trim()
        });
      }
      loaded++;
    }

    hasMore = !!data.data?.has_more;
    pageToken = data.data?.page_token || '';
  } while (hasMore && pageToken);

  priceCache = nextCache;
  priceCacheLoadedAt = Date.now();
  console.log(`Loaded ${loaded} Feishu price rows (${priceCache.size} styles)`);
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
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if ((req.path === '/api/reload' || req.path === '/api/index/add-image' || req.path.startsWith('/api/index/job/')) && secret === APP_PASSWORD) return next();
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
  res.json({
    authenticated: authed,
    ready: searchReady,
    progress: loadingProgress,
    error: loadError,
    styles: Object.keys(styleMetadata).length,
    images: metadataList.length,
    hasThumbnails: !!thumbnailsFolderId,
    priceRows: priceCache.size,
    priceCacheLoadedAt
  });
});

app.get('/api/style/:style', (req, res) => {
  const style = normalizeStyleId(req.params.style);
  const meta = styleMetadata[style];
  if (!meta) return res.status(404).json({ found: false, style });
  res.json({ found: true, style, ...meta });
});

app.post('/api/index/add-image', upload.single('image'), async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  if (!req.file) return res.status(400).json({ error: 'Missing image file' });

  const input = {
    imageBuffer: req.file.buffer,
    fileName: req.body.fileName || req.file.originalname,
    style: req.body.style || req.body.styleNumber || req.file.originalname,
    series: req.body.series || ''
  };

  if (req.query.async === '1') {
    const jobId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    indexJobs.set(jobId, { status: 'running', startedAt: new Date().toISOString() });
    addImageToIndex(input)
      .then(result => indexJobs.set(jobId, { status: 'done', finishedAt: new Date().toISOString(), result }))
      .catch(error => indexJobs.set(jobId, { status: 'error', finishedAt: new Date().toISOString(), error: error.message }));
    return res.status(202).json({ accepted: true, jobId });
  }

  try {
    const result = await addImageToIndex(input);
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('Add image index error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/index/job/:id', (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  const job = indexJobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: 'Job not found' });
  res.json(job);
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
    const timings = {};
    const startedAt = Date.now();
    const queryEmb = await getQueryEmbedding(req.file.buffer);
    timings.embeddingMs = Date.now() - startedAt;
    const topK = Math.max(1, Math.min(50, parseInt(req.query.topK || req.body.topK || '5', 10) || 5));
    const searchStartedAt = Date.now();
    const results = searchStyles(queryEmb, topK);
    timings.rankMs = Date.now() - searchStartedAt;

    // Lookup prices in parallel for speed
    const priceStartedAt = Date.now();
    const pricePromises = results.map(r => lookupPrice(r.style));
    const prices = await Promise.all(pricePromises);
    timings.priceMs = Date.now() - priceStartedAt;
    results.forEach((r, i) => {
      r.matchPercent = Math.round(r.score * 100);
      if (prices[i]) {
        r.wholesalePrice = prices[i].wholesale;
        r.retailPrice = prices[i].retail;
      }
    });

    console.log(`Results: ${results.map(r => `${r.style}(${r.matchPercent}%)`).join(', ')}`);
    timings.totalMs = Date.now() - startedAt;
    res.json({ results, hasThumbnails: !!thumbnailsFolderId, timings });
  } catch (e) {
    console.error('Search error:', e);
    res.status(500).json({ error: `搜索失败: ${e.message}` });
  }
});

// Reload endpoint for n8n
app.post('/api/reload', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  searchReady = false; styleEmbeddings = {}; styleMetadata = {}; imageEmbeddings = null; imageNorms = []; loadError = null;
  try { await loadAndInit(); res.json({ success: true, styles: Object.keys(styleMetadata).length }); }
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
  console.log(`✅ Search engine ready — ${Object.keys(styleMetadata).length} styles, ${metadataList.length} images loaded`);
}

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  if (!GCP_CREDENTIALS.private_key) { loadError = 'Missing GCP_CREDENTIALS'; console.error(loadError); return; }
  if (!DRIVE_FOLDER_ID) { loadError = 'Missing DRIVE_FOLDER_ID'; console.error(loadError); return; }
  loadAndInit().catch(e => { loadError = e.message; console.error('Init failed:', e); });
});
