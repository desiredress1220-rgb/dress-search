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
const indexFileIds = { metadata: null, embeddings: null, dims: null, deltaMetadata: null, deltaEmbeddings: null, tombstones: null, oneDriveState: null };
const indexJobs = new Map();
const HIDDEN_STYLE_PREFIXES = ['MD'];
const DELTA_METADATA_FILE = 'delta_metadata.json';
const DELTA_EMBEDDINGS_FILE = 'delta_embeddings.bin';
const TOMBSTONES_FILE = 'index_tombstones.json';
const ONEDRIVE_STATE_FILE = 'onedrive_delta_state.json';
const ONEDRIVE_DELTA_ROOT_URL = 'https://graph.microsoft.com/v1.0/me/drive/root:/%E5%85%AC%E5%8F%B8%E4%BA%A7%E5%93%81%E6%AC%BE%E5%BC%8F%E5%9B%BE%E7%89%87/%E5%85%AC%E4%BB%94/%E5%85%AC%E4%BB%94%E5%9B%BE:/delta';
const MAX_ONEDRIVE_ADDS_PER_RUN = 20;

function textFieldValue(value) {
  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number') return String(value);
  if (Array.isArray(value)) return value.map(textFieldValue).join('');
  if (typeof value === 'object') return textFieldValue(value.text ?? value.name ?? value.value ?? value.formatted_value ?? '');
  return String(value);
}

function priceFieldValue(value) {
  const raw = String(textFieldValue(value)).trim();
  if (!raw) return null;
  const match = raw.match(/-?\d+(?:\.\d+)?/);
  if (!match) return null;
  const num = Number(match[0]);
  if (!Number.isFinite(num)) return null;
  return Number.isInteger(num) ? String(num) : String(num).replace(/\.?0+$/, '');
}

function normalizeStyleId(value) {
  return textFieldValue(value)
    .replace(/\.[^.]+$/, '')
    .replace(/\s*\([^)]*\)\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function extractStyleId(value) {
  const normalized = normalizeStyleId(value);
  const match = normalized.match(/[A-Z]{1,5}\d{3,6}(?:-[A-Z0-9]+)?/);
  return match ? match[0] : '';
}

function displaySeriesForStyle(style, fallback = '') {
  const prefix = extractStyleId(style).match(/^[A-Z]+/)?.[0];
  if (prefix) return `${prefix}系列`;
  return fallback || '';
}

function resolvedStyleId({ style, style_number, item_no, name, fileName, driveName } = {}) {
  return extractStyleId(name || fileName || driveName) ||
    extractStyleId(style || style_number || item_no) ||
    normalizeStyleId(style || style_number || item_no || name || fileName || driveName || 'unknown');
}

function metadataStyleId(img) {
  return resolvedStyleId(img);
}

function isHiddenStyle(style) {
  const normalized = normalizeStyleId(style);
  return HIDDEN_STYLE_PREFIXES.some(prefix => normalized.startsWith(prefix));
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

async function readJsonDriveFile(fileId, fallback) {
  if (!fileId) return fallback;
  const resp = await driveDownload(fileId);
  try {
    return await resp.json();
  } catch (e) {
    console.warn(`Failed to parse Drive JSON file ${fileId}:`, e.message);
    return fallback;
  }
}

async function ensureDriveFile(fileName, initialBuffer, mimeType) {
  let file = await driveSearchFile(DRIVE_FOLDER_ID, fileName);
  if (!file) {
    file = await driveUploadToFolder(DRIVE_FOLDER_ID, fileName, initialBuffer, mimeType);
  }
  return file.id;
}

function normalizedRecordKeys(record) {
  const keys = [];
  const id = normalizeStyleId(record.driveId || record.oneDriveId || record.id || '');
  const name = normalizeStyleId(record.driveName || record.name || record.fileName || '');
  const style = metadataStyleId(record);
  if (id) keys.push(`id:${id}`);
  if (name) keys.push(`name:${name}`);
  if (style) keys.push(`style:${style}`);
  return keys;
}

function buildTombstoneSet(tombstones) {
  const set = new Set();
  for (const item of Array.isArray(tombstones) ? tombstones : []) {
    const id = normalizeStyleId(item.driveId || item.oneDriveId || item.id || '');
    const name = normalizeStyleId(item.driveName || item.name || item.fileName || '');
    const style = normalizeStyleId(item.style || item.styleNumber || '');
    if (id) set.add(`id:${id}`);
    if (name) set.add(`name:${name}`);
    if (style && item.deleteStyle === true) set.add(`style:${style}`);
  }
  return set;
}

function applyTombstones(metadata, embeddings, tombstones) {
  const tombstoneSet = buildTombstoneSet(tombstones);
  if (!tombstoneSet.size) return { metadata, embeddings };

  const keptMetadata = [];
  const keptVectors = [];
  for (let i = 0; i < metadata.length; i++) {
    const record = metadata[i];
    const deleted = normalizedRecordKeys(record).some(key => tombstoneSet.has(key));
    if (!deleted) {
      keptMetadata.push(record);
      keptVectors.push(i);
    }
  }

  if (keptVectors.length === metadata.length) return { metadata, embeddings };

  const nextEmbeddings = new Float32Array(keptVectors.length * embDim);
  keptVectors.forEach((sourceIdx, targetIdx) => {
    const sourceOffset = sourceIdx * embDim;
    const targetOffset = targetIdx * embDim;
    nextEmbeddings.set(embeddings.subarray(sourceOffset, sourceOffset + embDim), targetOffset);
  });

  console.log(`Applied tombstones: ${metadata.length - keptMetadata.length} image records hidden`);
  return { metadata: keptMetadata, embeddings: nextEmbeddings };
}

function concatFloatEmbeddings(baseEmbeddings, deltaEmbeddings) {
  if (!deltaEmbeddings || !deltaEmbeddings.length) return baseEmbeddings;
  const combined = new Float32Array(baseEmbeddings.length + deltaEmbeddings.length);
  combined.set(baseEmbeddings);
  combined.set(deltaEmbeddings, baseEmbeddings.length);
  return combined;
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
  const deltaMetaFile = find(DELTA_METADATA_FILE);
  const deltaEmbFile = find(DELTA_EMBEDDINGS_FILE);
  const tombstonesFile = find(TOMBSTONES_FILE);
  const oneDriveStateFile = find(ONEDRIVE_STATE_FILE);
  const thumbFolder = files.find(f => f.name === 'thumbnails' && f.mimeType === 'application/vnd.google-apps.folder');

  if (!metaFile) throw new Error('metadata.json not found in Drive folder');
  if (!embFile) throw new Error('embeddings.bin not found in Drive folder');
  indexFileIds.metadata = metaFile.id;
  indexFileIds.embeddings = embFile.id;
  indexFileIds.dims = dimsFile?.id || null;
  indexFileIds.deltaMetadata = deltaMetaFile?.id || null;
  indexFileIds.deltaEmbeddings = deltaEmbFile?.id || null;
  indexFileIds.tombstones = tombstonesFile?.id || null;
  indexFileIds.oneDriveState = oneDriveStateFile?.id || null;

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
  let embeddings = new Float32Array(embBuffer.buffer, embBuffer.byteOffset, usableFloatCount);
  let usableMetadata = metadata.slice(0, Math.min(metadata.length, vectorCount));
  if (usableMetadata.length !== metadata.length || usableFloatCount !== floatCount) {
    console.warn(`Index size mismatch: metadata=${metadata.length}, completeVectors=${vectorCount}, usingMetadata=${usableMetadata.length}`);
  }
  console.log(`Loaded ${embeddings.length} float values (${vectorCount} complete vectors)`);

  if (deltaMetaFile && deltaEmbFile) {
    const deltaMetadata = await readJsonDriveFile(deltaMetaFile.id, []);
    const deltaEmbResp = await driveDownload(deltaEmbFile.id);
    let deltaEmbBuffer = Buffer.from(await deltaEmbResp.arrayBuffer());
    const deltaFloatCount = Math.floor(deltaEmbBuffer.byteLength / 4 / embDim) * embDim;
    const deltaVectorCount = Math.floor(deltaFloatCount / embDim);
    if (deltaVectorCount) {
      const deltaEmbeddings = new Float32Array(deltaEmbBuffer.buffer, deltaEmbBuffer.byteOffset, deltaFloatCount);
      usableMetadata = usableMetadata.concat(deltaMetadata.slice(0, deltaVectorCount));
      embeddings = concatFloatEmbeddings(embeddings, deltaEmbeddings);
      console.log(`Loaded delta index: ${deltaVectorCount} image records`);
    }
  }

  const tombstones = tombstonesFile ? await readJsonDriveFile(tombstonesFile.id, []) : [];
  ({ metadata: usableMetadata, embeddings } = applyTombstones(usableMetadata, embeddings, tombstones));

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
    const n = group.indices.length;
    if (!n) continue;
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
    signal: AbortSignal.timeout(15000),
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
    signal: AbortSignal.timeout(15000),
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
    if (isHiddenStyle(style)) continue;
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
  return selectDisplayResults(results, topK);
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
    if (isHiddenStyle(style)) continue;
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
    const styleAverageScore = styleEmbeddings[item.style] ? cosine(queryEmb, styleEmbeddings[item.style]) : topAverage;
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
  return selectDisplayResults(results, topK);
}

function selectDisplayResults(results, topK) {
  if (topK > 10) return results.slice(0, topK);

  const selected = [];
  const overflow = [];
  const seriesCounts = new Map();
  const maxPerSeries = 2;

  for (const result of results) {
    const seriesKey = result.series || result.style.replace(/\d.*$/, '');
    const count = seriesCounts.get(seriesKey) || 0;
    if (count < maxPerSeries) {
      selected.push(result);
      seriesCounts.set(seriesKey, count + 1);
    } else {
      overflow.push(result);
    }
    if (selected.length >= topK) break;
  }

  if (selected.length < topK) selected.push(...overflow.slice(0, topK - selected.length));
  return selected;
}

function filterResultsByPriceCatalog(results) {
  if (!priceCacheLoadedAt || !priceCache.size) return results;
  const filtered = results.filter(result => priceCache.has(normalizeStyleId(result.style)));
  return filtered.length ? filtered : results;
}

function findSimilarStyles(style, limit = 20) {
  const target = normalizeStyleId(style);
  if (!target) return [];
  const compactTarget = target.replace(/[^A-Z0-9]/g, '');
  const matches = [];
  for (const [key, meta] of Object.entries(styleMetadata)) {
    const compactKey = normalizeStyleId(key).replace(/[^A-Z0-9]/g, '');
    if (
      key === target ||
      key.includes(target) ||
      target.includes(key) ||
      compactKey.includes(compactTarget) ||
      compactTarget.includes(compactKey)
    ) {
      matches.push({
        style: key,
        count: meta.count || 0,
        series: meta.series || '',
        thumbIndex: meta.thumbIndex
      });
    }
    if (matches.length >= limit) break;
  }
  return matches;
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

async function addImageToIndex({ imageBuffer, fileName, style, series, driveId, parentPath }) {
  if (!searchReady || !imageEmbeddings) throw new Error('Search index is not ready');
  if (!indexFileIds.metadata || !indexFileIds.embeddings) throw new Error('Index file ids are not loaded');

  const normalizedStyle = resolvedStyleId({ name: fileName, fileName, driveName: fileName, style });
  if (!normalizedStyle) throw new Error('Missing style number');
  const resolvedSeries = displaySeriesForStyle(normalizedStyle, series);

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
    series: resolvedSeries,
    name: fileName,
    fileName,
    driveName: fileName,
    driveId: driveId || '',
    parentPath: parentPath || '',
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

async function addImageToDeltaIndex({ imageBuffer, fileName, style, series, driveId, parentPath }) {
  if (!searchReady) throw new Error('Search index is not ready');

  const normalizedStyle = resolvedStyleId({ name: fileName, fileName, driveName: fileName, style });
  const normalizedName = normalizeStyleId(fileName);
  if (!normalizedStyle) throw new Error('Missing style number');
  const resolvedSeries = displaySeriesForStyle(normalizedStyle, series);

  const existing = metadataList.find(img => {
    const existingName = normalizeStyleId(img.driveName || img.name || img.fileName || '');
    return normalizedName && existingName === normalizedName;
  });
  if (existing) return { added: false, reason: 'already_exists', style: normalizedStyle };

  let deltaMetadataId;
  let deltaEmbeddingsId;
  try {
    deltaMetadataId = await ensureDriveFile(DELTA_METADATA_FILE, Buffer.from('[]'), 'application/json');
    deltaEmbeddingsId = await ensureDriveFile(DELTA_EMBEDDINGS_FILE, Buffer.alloc(0), 'application/octet-stream');
  } catch (e) {
    throw new Error(`Delta index files are missing and cannot be created by the service account. Create ${DELTA_METADATA_FILE} and ${DELTA_EMBEDDINGS_FILE} in the Drive index folder first. ${e.message}`);
  }
  indexFileIds.deltaMetadata = deltaMetadataId;
  indexFileIds.deltaEmbeddings = deltaEmbeddingsId;

  const deltaMetadata = await readJsonDriveFile(deltaMetadataId, []);
  const deltaExisting = deltaMetadata.find(img => {
    const existingName = normalizeStyleId(img.driveName || img.name || img.fileName || '');
    return normalizedName && existingName === normalizedName;
  });
  if (deltaExisting) return { added: false, reason: 'already_exists_in_delta', style: normalizedStyle };

  const embedding = await getIndexEmbedding(imageBuffer);
  if (embedding.length !== embDim) throw new Error(`Unexpected embedding dimension ${embedding.length}, expected ${embDim}`);

  const record = {
    style: normalizedStyle,
    style_number: normalizedStyle,
    series: resolvedSeries,
    name: fileName,
    fileName,
    driveName: fileName,
    driveId: driveId || '',
    parentPath: parentPath || '',
    addedAt: new Date().toISOString(),
    source: 'onedrive-delta'
  };

  const embResp = await driveDownload(deltaEmbeddingsId);
  const embBuffer = Buffer.from(await embResp.arrayBuffer());
  const embeddingBuffer = Buffer.from(embedding.buffer, embedding.byteOffset, embedding.byteLength);
  await driveUpdateFile(deltaEmbeddingsId, Buffer.concat([embBuffer, embeddingBuffer]), 'application/octet-stream');
  await driveUpdateFile(deltaMetadataId, Buffer.from(JSON.stringify(deltaMetadata.concat(record))), 'application/json');

  return { added: true, style: normalizedStyle, deltaImages: deltaMetadata.length + 1 };
}

function applyMetadataUpdates(records, items) {
  const updated = [];
  const missing = [];
  const seenIndexes = new Set();

  for (const item of Array.isArray(items) ? items : []) {
    const fileName = item.fileName || item.name || item.driveName || '';
    const driveId = normalizeStyleId(item.driveId || item.id || item.oneDriveId || '');
    const normalizedName = normalizeStyleId(fileName);
    const normalizedStyle = resolvedStyleId({ name: fileName, fileName, driveName: fileName, style: item.style || item.styleNumber });
    if (!normalizedStyle && !driveId && !normalizedName) continue;

    const idx = records.findIndex((record, i) => {
      if (seenIndexes.has(i)) return false;
      const recordDriveId = normalizeStyleId(record.driveId || record.oneDriveId || record.id || '');
      const recordName = normalizeStyleId(record.driveName || record.name || record.fileName || '');
      return (driveId && recordDriveId && recordDriveId === driveId) ||
        (normalizedName && recordName && recordName === normalizedName);
    });

    if (idx < 0) {
      missing.push({ driveId: item.driveId || item.id || '', name: fileName, style: normalizedStyle });
      continue;
    }

    seenIndexes.add(idx);
    const next = {
      ...records[idx],
      style: normalizedStyle,
      style_number: normalizedStyle,
      item_no: normalizedStyle,
      series: displaySeriesForStyle(normalizedStyle, item.series || records[idx].series || ''),
      name: fileName || records[idx].name,
      fileName: fileName || records[idx].fileName,
      driveName: fileName || records[idx].driveName,
      driveId: item.driveId || item.id || records[idx].driveId || '',
      oneDriveId: item.driveId || item.id || records[idx].oneDriveId || '',
      parentPath: item.parentPath || records[idx].parentPath || '',
      updatedAt: new Date().toISOString(),
      source: records[idx].source || 'onedrive-delta'
    };
    records[idx] = next;
    updated.push({ index: idx, style: normalizedStyle, name: next.name, driveId: next.driveId || next.oneDriveId || '' });
  }

  return { updated, missing };
}

async function updateIndexMetadata(items) {
  if (!Array.isArray(items)) throw new Error('items must be an array');
  if (!items.length) return { updated: 0, missing: 0, details: [] };
  if (!indexFileIds.metadata) throw new Error('Metadata file id is not loaded');

  const metaResp = await driveDownload(indexFileIds.metadata);
  const mainMetadata = await metaResp.json();
  const mainResult = applyMetadataUpdates(mainMetadata, items);
  if (mainResult.updated.length) {
    await driveUpdateFile(indexFileIds.metadata, Buffer.from(JSON.stringify(mainMetadata)), 'application/json');
  }

  let deltaResult = { updated: [], missing: [] };
  if (indexFileIds.deltaMetadata) {
    const deltaMetadata = await readJsonDriveFile(indexFileIds.deltaMetadata, []);
    deltaResult = applyMetadataUpdates(deltaMetadata, items);
    if (deltaResult.updated.length) {
      await driveUpdateFile(indexFileIds.deltaMetadata, Buffer.from(JSON.stringify(deltaMetadata)), 'application/json');
    }
  }

  const memoryResult = applyMetadataUpdates(metadataList, items);
  return {
    updated: mainResult.updated.length + deltaResult.updated.length,
    missing: mainResult.missing.length,
    memoryUpdated: memoryResult.updated.length,
    details: mainResult.updated.concat(deltaResult.updated).slice(0, 20),
    missingDetails: mainResult.missing.slice(0, 20)
  };
}

async function ensureOneDriveStateFile() {
  const id = await ensureDriveFile(ONEDRIVE_STATE_FILE, Buffer.from('{}'), 'application/json');
  indexFileIds.oneDriveState = id;
  return id;
}

async function readOneDriveState() {
  const id = indexFileIds.oneDriveState;
  if (!id) {
    const tombstones = await readJsonDriveFile(indexFileIds.tombstones, []);
    const state = Array.isArray(tombstones) ? tombstones.find(item => item && item.type === 'onedrive_delta_state') : null;
    if (state && typeof state === 'object') return state;
    const dims = await readJsonDriveFile(indexFileIds.dims, {});
    if (dims?.oneDriveDeltaState && typeof dims.oneDriveDeltaState === 'object') return dims.oneDriveDeltaState;
    const metadata = await readJsonDriveFile(indexFileIds.metadata, []);
    const metadataState = Array.isArray(metadata) ? metadata.find(item => item && item.type === 'onedrive_delta_state') : null;
    if (metadataState && typeof metadataState === 'object') return metadataState;
    return state && typeof state === 'object' ? state : {};
  }
  const state = await readJsonDriveFile(id, {});
  return state && typeof state === 'object' ? state : {};
}

async function writeOneDriveState(nextState) {
  const id = indexFileIds.oneDriveState;
  if (!id) {
    if (indexFileIds.tombstones) {
      const tombstones = await readJsonDriveFile(indexFileIds.tombstones, []);
      const nextTombstones = Array.isArray(tombstones) ? tombstones.filter(item => !item || item.type !== 'onedrive_delta_state') : [];
      nextTombstones.push({ ...nextState, type: 'onedrive_delta_state' });
      await driveUpdateFile(indexFileIds.tombstones, Buffer.from(JSON.stringify(nextTombstones)), 'application/json');
      return { saved: true, storage: TOMBSTONES_FILE };
    }
    if (indexFileIds.dims) {
      const dims = await readJsonDriveFile(indexFileIds.dims, { dimensions: embDim });
      await driveUpdateFile(indexFileIds.dims, Buffer.from(JSON.stringify({ ...dims, oneDriveDeltaState: nextState })), 'application/json');
      return { saved: true, storage: 'dims.json' };
    }
    if (indexFileIds.metadata) {
      const metadata = await readJsonDriveFile(indexFileIds.metadata, []);
      const nextMetadata = Array.isArray(metadata) ? metadata.filter(item => !item || item.type !== 'onedrive_delta_state') : [];
      nextMetadata.push({ ...nextState, type: 'onedrive_delta_state' });
      await driveUpdateFile(indexFileIds.metadata, Buffer.from(JSON.stringify(nextMetadata)), 'application/json');
      return { saved: true, storage: 'metadata.json' };
    }
    return { saved: false, reason: `${ONEDRIVE_STATE_FILE}, ${TOMBSTONES_FILE}, dims.json, and metadata.json are missing` };
  }
  await driveUpdateFile(id, Buffer.from(JSON.stringify(nextState)), 'application/json');
  return { saved: true, storage: ONEDRIVE_STATE_FILE };
}

async function getOneDriveDeltaUrl(mode = '') {
  if (mode === 'reconcile') return ONEDRIVE_DELTA_ROOT_URL;
  const state = await readOneDriveState();
  return state.deltaLink || `${ONEDRIVE_DELTA_ROOT_URL}?token=latest`;
}

function seriesFromOneDrivePath(parentPath, style = '') {
  const text = decodeURIComponent(String(parentPath || ''));
  const stylePrefix = extractStyleId(style).match(/^[A-Z]+/)?.[0] || '';
  const seriesMatch = text.match(/([A-Z]{1,5})\s*系列/i);
  const prefix = (seriesMatch?.[1] || stylePrefix || '').toUpperCase();
  return prefix ? `${prefix}系列` : '其他';
}

function isOneDriveImageFile(item) {
  const name = String(item?.name || '').toLowerCase();
  const mime = String(item?.file?.mimeType || '').toLowerCase();
  return mime.startsWith('image/') || /\.(jpe?g|png|webp)$/i.test(name);
}

function oneDriveImageRecord(item) {
  const name = item.name || '';
  const parentPath = item.parentReference?.path || '';
  const style = extractStyleId(name) || normalizeStyleId(name);
  return {
    id: item.id || '',
    driveId: item.id || '',
    name,
    fileName: name,
    driveName: name,
    style,
    series: seriesFromOneDrivePath(parentPath, style),
    parentPath,
    size: item.size || 0,
    lastModified: item.lastModifiedDateTime || '',
    downloadUrl: item['@microsoft.graph.downloadUrl'] || item['@content.downloadUrl'] || ''
  };
}

async function processOneDriveDeltaPayload(payload, { mode = '', addLimit = MAX_ONEDRIVE_ADDS_PER_RUN } = {}) {
  const changes = Array.isArray(payload?.value) ? payload.value : [];
  const nextLink = payload?.['@odata.nextLink'] || '';
  const finalDeltaLink = payload?.['@odata.deltaLink'] || '';
  const deleted = [];
  const changedImages = [];

  for (const item of changes) {
    if (item.deleted) {
      deleted.push({
        id: item.id || '',
        driveId: item.id || '',
        name: item.name || '',
        reason: 'onedrive-delete'
      });
      continue;
    }
    if (item.folder || !isOneDriveImageFile(item)) continue;
    changedImages.push(oneDriveImageRecord(item));
  }

  const sync = { tombstone: null, updated: [], added: [], skipped: [], errors: [] };
  if (changedImages.length) {
    try {
      const result = await updateIndexMetadata(changedImages);
      sync.updated = result.details || [];
      const updatedKeys = new Set(sync.updated.flatMap(item => [
        item.driveId ? `id:${normalizeStyleId(item.driveId)}` : '',
        item.name ? `name:${normalizeStyleId(item.name)}` : ''
      ]).filter(Boolean));

      let addsStarted = 0;
      const missing = result.missingDetails || [];
      for (const file of missing) {
        const source = changedImages.find(item =>
          (file.driveId && normalizeStyleId(file.driveId) === normalizeStyleId(item.driveId)) ||
          (file.name && normalizeStyleId(file.name) === normalizeStyleId(item.name))
        );
        if (!source) continue;
        if (updatedKeys.has(`id:${normalizeStyleId(source.driveId)}`) || updatedKeys.has(`name:${normalizeStyleId(source.name)}`)) continue;
        if (!source.downloadUrl) {
          sync.skipped.push({ name: source.name, reason: 'no_download_url' });
          continue;
        }
        if (addsStarted >= addLimit) {
          sync.skipped.push({ name: source.name, reason: 'add_limit' });
          continue;
        }
        addsStarted += 1;
        try {
          const resp = await fetch(source.downloadUrl);
          if (!resp.ok) {
            sync.skipped.push({ name: source.name, reason: `download_${resp.status}` });
            continue;
          }
          const imageBuffer = Buffer.from(await resp.arrayBuffer());
          sync.added.push(await addImageToDeltaIndex({
            imageBuffer,
            fileName: source.name,
            style: source.style,
            series: source.series,
            driveId: source.driveId,
            parentPath: source.parentPath
          }));
        } catch (e) {
          sync.errors.push({ name: source.name, error: e.message });
        }
      }
    } catch (e) {
      sync.errors.push({ op: 'update-metadata', error: e.message });
    }
  }

  if (deleted.length) {
    try {
      sync.tombstone = await addIndexTombstones(deleted);
    } catch (e) {
      sync.errors.push({ op: 'tombstone', error: e.message });
    }
  }

  if ((sync.updated.length || sync.added.length || deleted.length) && !sync.errors.length) {
    await loadIndex();
  }

  if (finalDeltaLink && !nextLink && !sync.errors.length && mode !== 'dry-run') {
    const stateResult = await writeOneDriveState({
      deltaLink: finalDeltaLink,
      updatedAt: new Date().toISOString()
    });
    sync.state = stateResult;
  }

  return {
    success: sync.errors.length === 0,
    mode,
    hasMoreData: !!nextLink,
    url: nextLink,
    deltaLinkSaved: !!sync.state?.saved,
    summary: {
      files: changes.length,
      changedImages: changedImages.length,
      deleted: deleted.length,
      metadataUpdated: sync.updated.length,
      added: sync.added.length,
      skipped: sync.skipped.length,
      errors: sync.errors.length
    },
    sync: {
      updated: sync.updated.slice(0, 10),
      added: sync.added.slice(0, 10),
      skipped: sync.skipped.slice(0, 20),
      errors: sync.errors.slice(0, 20),
      state: sync.state || null
    }
  };
}

async function addIndexTombstones(items) {
  if (!Array.isArray(items)) throw new Error('items must be an array');
  if (!items.length) return { added: 0, total: 0 };

  let tombstonesId;
  try {
    tombstonesId = await ensureDriveFile(TOMBSTONES_FILE, Buffer.from('[]'), 'application/json');
  } catch (e) {
    throw new Error(`Tombstone file is missing and cannot be created by the service account. Create ${TOMBSTONES_FILE} in the Drive index folder first. ${e.message}`);
  }
  indexFileIds.tombstones = tombstonesId;

  const existing = await readJsonDriveFile(tombstonesId, []);
  const seen = new Set(existing.flatMap(normalizedRecordKeys));
  const added = [];

  for (const item of items) {
    const normalized = {
      id: item.id || item.driveId || item.oneDriveId || '',
      driveId: item.driveId || item.id || '',
      name: item.name || item.driveName || item.fileName || '',
      fileName: item.fileName || item.name || '',
      style: item.style || item.styleNumber || '',
      reason: item.reason || 'onedrive-delta',
      deletedAt: new Date().toISOString()
    };
    const keys = normalizedRecordKeys(normalized);
    if (!keys.length || keys.every(key => seen.has(key))) continue;
    keys.forEach(key => seen.add(key));
    added.push(normalized);
  }

  if (added.length) {
    await driveUpdateFile(tombstonesId, Buffer.from(JSON.stringify(existing.concat(added))), 'application/json');
  }

  return { added: added.length, total: existing.length + added.length };
}

async function compactMainIndexWithTombstones(items) {
  if (!indexFileIds.metadata || !indexFileIds.embeddings) throw new Error('Index file ids are not loaded');
  const tombstoneSet = buildTombstoneSet(items);
  if (!tombstoneSet.size) return { removed: 0, total: metadataList.length, mode: 'main-index-compact' };

  const metaResp = await driveDownload(indexFileIds.metadata);
  const metadata = await metaResp.json();
  const embResp = await driveDownload(indexFileIds.embeddings);
  let embBuffer = Buffer.from(await embResp.arrayBuffer());
  let usableBytes = embBuffer.byteLength - (embBuffer.byteLength % 4);
  let floatCount = usableBytes / 4;
  const vectorCount = Math.floor(floatCount / embDim);
  const usableFloatCount = vectorCount * embDim;
  embBuffer = embBuffer.subarray(0, usableFloatCount * 4);
  const embeddings = new Float32Array(embBuffer.buffer, embBuffer.byteOffset, usableFloatCount);

  const usableMetadata = metadata.slice(0, Math.min(metadata.length, vectorCount));
  const keptMetadata = [];
  const keptVectors = [];
  for (let i = 0; i < usableMetadata.length; i++) {
    const record = usableMetadata[i];
    const deleted = normalizedRecordKeys(record).some(key => tombstoneSet.has(key));
    if (!deleted) {
      keptMetadata.push(record);
      keptVectors.push(i);
    }
  }

  const removed = usableMetadata.length - keptMetadata.length;
  if (!removed) return { removed: 0, total: usableMetadata.length, mode: 'main-index-compact' };

  const nextEmbeddings = Buffer.alloc(keptVectors.length * embDim * 4);
  keptVectors.forEach((sourceIdx, targetIdx) => {
    const sourceStart = sourceIdx * embDim;
    const targetStart = targetIdx * embDim;
    for (let d = 0; d < embDim; d++) {
      nextEmbeddings.writeFloatLE(embeddings[sourceStart + d], (targetStart + d) * 4);
    }
  });

  await driveUpdateFile(indexFileIds.embeddings, nextEmbeddings, 'application/octet-stream');
  await driveUpdateFile(indexFileIds.metadata, Buffer.from(JSON.stringify(keptMetadata)), 'application/json');

  return { removed, total: keptMetadata.length, mode: 'main-index-compact' };
}

// ============================================================
// Feishu Bitable
// ============================================================
let feishuToken = null;
let feishuTokenExpiry = 0;
let priceCache = new Map();
let priceCacheLoadedAt = 0;
let priceCacheRefreshPromise = null;
let priceCacheError = null;

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
        wholesale: priceFieldValue(f['WHOLESALE PRICE USD'] ?? f['批发价']),
        retail: priceFieldValue(f['RETAILER PRICE USD'] ?? f['零售价'])
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
  schedulePriceCacheRefreshIfNeeded();
  return priceCache.get(normalizeStyleId(styleNumber)) || null;
}

function schedulePriceCacheRefreshIfNeeded(force = false) {
  if (!FEISHU_APP_SECRET) return;
  if (!force && priceCache.size && Date.now() - priceCacheLoadedAt < PRICE_CACHE_TTL_MS) return;
  if (priceCacheRefreshPromise) return;

  priceCacheRefreshPromise = refreshPriceCacheIfNeeded(force)
    .catch(e => {
      priceCacheError = e.message;
      console.error('Feishu cache refresh error:', e.message);
    })
    .finally(() => { priceCacheRefreshPromise = null; });
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
        field_names: ['ITEM NO', 'WHOLESALE PRICE USD', 'RETAILER PRICE USD'],
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
          wholesale: priceFieldValue(f['WHOLESALE PRICE USD']),
          retail: priceFieldValue(f['RETAILER PRICE USD']),
          shipping: priceFieldValue(f['SHIPPING COST USD']),
          total: priceFieldValue(f['TOTAL AMOUNT USD']),
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
  priceCacheError = null;
  console.log(`Loaded ${loaded} Feishu price rows (${priceCache.size} styles)`);
}

// ============================================================
// Express
// ============================================================
app.use(express.json({ limit: '10mb' }));
app.use(cookieParser());

const AUTH_SECRET = 'dress-search-v1';
function makeAuthToken(pw) { return crypto.createHmac('sha256', AUTH_SECRET).update(pw).digest('hex'); }

function authCheck(req, res, next) {
  if (req.path === '/api/login') return next();
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if ((req.path === '/api/reload' || req.path === '/api/prices/reload' || req.path === '/api/onedrive/next-delta-url' || req.path === '/api/onedrive/process-delta' || req.path === '/api/index/add-image' || req.path === '/api/index/update-metadata' || req.path === '/api/index/tombstone' || req.path.startsWith('/api/index/job/') || req.path.startsWith('/api/style/') || req.path.startsWith('/api/styles/find/')) && secret === APP_PASSWORD) return next();
  const token = req.cookies?.auth;
  const expected = makeAuthToken(APP_PASSWORD);
  if (token === expected) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: '请先登录' });
  next();
}

app.use(authCheck);
app.use((req, res, next) => {
  if (!req.path.startsWith('/api/')) res.set('Cache-Control', 'no-store');
  next();
});
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
    priceCacheLoadedAt,
    priceCacheError
  });
});

app.post('/api/prices/reload', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  try {
    await refreshPriceCacheIfNeeded(true);
    res.json({ success: true, priceRows: priceCache.size, priceCacheLoadedAt, priceCacheError });
  } catch (e) {
    priceCacheError = e.message;
    res.status(500).json({ success: false, error: e.message });
  }
});

app.get('/api/style/:style', (req, res) => {
  const style = normalizeStyleId(req.params.style);
  const meta = styleMetadata[style];
  if (!meta) return res.status(404).json({ found: false, style });
  res.json({ found: true, style, ...meta });
});

app.get('/api/styles/find/:style', (req, res) => {
  const style = normalizeStyleId(req.params.style);
  const exact = styleMetadata[style] || null;
  res.json({
    found: !!exact,
    style,
    exact,
    similar: exact ? [] : findSimilarStyles(style)
  });
});

app.get('/api/onedrive/next-delta-url', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  try {
    const mode = String(req.query.mode || '');
    res.json({ success: true, url: await getOneDriveDeltaUrl(mode), mode });
  } catch (e) {
    console.error('OneDrive next delta URL error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/onedrive/process-delta', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  try {
    const mode = String(req.query.mode || req.body.mode || '');
    const addLimit = Math.max(0, Math.min(Number(req.query.addLimit || MAX_ONEDRIVE_ADDS_PER_RUN), MAX_ONEDRIVE_ADDS_PER_RUN));
    res.json(await processOneDriveDeltaPayload(req.body, { mode, addLimit }));
  } catch (e) {
    console.error('OneDrive process delta error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/index/update-metadata', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  try {
    const items = req.body.items || req.body.files || req.body.changes || [];
    const result = await updateIndexMetadata(items);
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('Update index metadata error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/index/add-image', upload.single('image'), async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  if (!req.file) return res.status(400).json({ error: 'Missing image file' });

  const input = {
    imageBuffer: req.file.buffer,
    fileName: req.body.fileName || req.file.originalname,
    style: req.body.style || req.body.styleNumber || req.file.originalname,
    series: req.body.series || '',
    driveId: req.body.driveId || req.body.id || '',
    parentPath: req.body.parentPath || ''
  };

  if (req.query.async === '1') {
    const jobId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    indexJobs.set(jobId, { status: 'running', startedAt: new Date().toISOString() });
    addImageToDeltaIndex(input)
      .then(result => indexJobs.set(jobId, { status: 'done', finishedAt: new Date().toISOString(), result }))
      .catch(error => indexJobs.set(jobId, { status: 'error', finishedAt: new Date().toISOString(), error: error.message }));
    return res.status(202).json({ accepted: true, jobId });
  }

  try {
    const result = await addImageToDeltaIndex(input);
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('Add image index error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/index/tombstone', async (req, res) => {
  const secret = req.headers['x-reload-secret'] || req.query.secret;
  if (secret !== APP_PASSWORD) return res.status(403).json({ error: 'Forbidden' });
  try {
    const result = await addIndexTombstones(req.body.items || req.body.deleted || []);
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('Tombstone index error:', e);
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
  if (!searchReady) return res.status(503).json({ error: '数据加载中，请稍候...', retryAfterMs: 3000 });
  if (!req.file) return res.status(400).json({ error: '请上传图片' });

  try {
    console.log(`Search: ${req.file.originalname} (${Math.round(req.file.size/1024)}KB)`);
    const timings = {};
    const startedAt = Date.now();
    const queryEmb = await getQueryEmbedding(req.file.buffer);
    timings.embeddingMs = Date.now() - startedAt;
    const topK = Math.max(1, Math.min(50, parseInt(req.query.topK || req.body.topK || '5', 10) || 5));
    const searchStartedAt = Date.now();
    const candidateK = Math.min(50, Math.max(topK * 4, topK + 10));
    let results = searchStyles(queryEmb, candidateK);
    results = filterResultsByPriceCatalog(results);
    results = results.slice(0, topK);
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
  searchReady = false;
  styleEmbeddings = {};
  styleMetadata = {};
  metadataList = [];
  imageEmbeddings = null;
  imageNorms = [];
  loadError = null;
  loadingProgress = 'Reload queued...';
  res.status(202).json({ accepted: true });

  setImmediate(() => {
    loadAndInit().catch(e => {
      loadError = e.message;
      searchReady = false;
      console.error('Reload failed:', e);
    });
  });
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

setInterval(() => {
  if (searchReady) schedulePriceCacheRefreshIfNeeded();
}, 60 * 1000);

setTimeout(() => {
  if (searchReady) schedulePriceCacheRefreshIfNeeded(true);
}, 5 * 1000);
