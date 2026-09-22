const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const fetch = require('node-fetch');
const FormData = require('form-data');
const { Jimp } = require('jimp');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const { createWorker } = require('tesseract.js');
const nodemailer = require('nodemailer');
const axios = require('axios');
const { Client, LocalAuth } = require('whatsapp-web.js');
const QRCode = require('qrcode');

// ─── WHATSAPP CLIENT ──────────────────────────────────────────────────────────
let waState = 'disconnected'; // 'disconnected' | 'qr' | 'ready' | 'error'
let waQRData = null;          // base64 PNG of latest QR
let waClient  = null;

function buildWAClient() {
  const client = new Client({
    authStrategy: new LocalAuth({ dataPath: path.join(__dirname, '.wa_session') }),
    puppeteer: {
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
    }
  });

  client.on('qr', async (qr) => {
    waState  = 'qr';
    waQRData = await QRCode.toDataURL(qr);
    console.log('[WA] QR code ready — scan in the browser UI');
  });

  client.on('ready', () => {
    waState  = 'ready';
    waQRData = null;
    console.log('[WA] ✅ WhatsApp connected as:', client.info?.pushname || 'unknown');
  });

  client.on('authenticated', () => {
    waState  = 'ready';
    console.log('[WA] Authenticated from saved session');
  });

  client.on('auth_failure', () => {
    waState  = 'disconnected';
    waQRData = null;
    console.error('[WA] Auth failure — session cleared');
  });

  client.on('disconnected', (reason) => {
    waState  = 'disconnected';
    waQRData = null;
    console.warn('[WA] Disconnected:', reason);
  });

  return client;
}

async function initWA() {
  try {
    const lockFile = path.join(__dirname, '.wa_session', 'session', 'DevToolsActivePort');
    if (fs.existsSync(lockFile)) {
      try { fs.unlinkSync(lockFile); } catch(e) {}
    }
    if (waClient) {
      try { await waClient.destroy(); } catch(e) {}
      waClient = null;
    }
    waClient = buildWAClient();
    await waClient.initialize();
  } catch(e) {
    waState = 'error';
    console.error('[WA] Init error:', e.message);
  }
}

const PORT = process.env.PORT || 5000;
const PUBLIC_DIR = __dirname;
const vtuAgent = new https.Agent({
  rejectUnauthorized: false,
  keepAlive: true,
  keepAliveMsecs: 5000,
  maxSockets: 32,
  maxFreeSockets: 8,
  timeout: 60000
});

// OCR.space kept as fallback — free 500/day but Tesseract is primary (unlimited)
const OCR_API_KEY = process.env.OCR_API_KEY || 'K81923247388957';

const VTU_BASE    = 'https://results.vtu.ac.in/MJ26cbcs';
const VTU_INDEX   = `${VTU_BASE}/index.php`;
const VTU_RESULT  = `${VTU_BASE}/resultpage.php`;
const VTU_CAPTCHA = 'https://results.vtu.ac.in/captcha/vtu_captcha.php?_CAPTCHA&t=';

const GOOGLE_SHEET_ID   = process.env.GOOGLE_SHEET_ID || '';
const GOOGLE_CREDS_PATH = path.join(__dirname, 'google_credentials.json');

const MIME_TYPES = {
  '.html': 'text/html; charset=UTF-8', '.css': 'text/css; charset=UTF-8',
  '.js': 'application/javascript; charset=UTF-8', '.json': 'application/json; charset=UTF-8',
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.csv': 'text/csv'
};

// ─── TESSERACT WORKER POOL ────────────────────────────────────────────────────
const WORKER_COUNT = 4;
const workerPool = [];
let workersReady = false;

async function initWorkers() {
  console.log('[OCR] Starting ' + WORKER_COUNT + ' local Tesseract workers...');
  for (let i = 0; i < WORKER_COUNT; i++) {
    try {
      const w = await createWorker('eng', 1, {
        logger: () => {},
        langPath: __dirname,
        cachePath: __dirname,
        gzip: false
      });
      await w.setParameters({
        tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789',
        tessedit_pageseg_mode: '7',
      });
      workerPool.push({ worker: w, busy: false });
      console.log('[OCR] Worker ' + (i + 1) + '/' + WORKER_COUNT + ' ready');
    } catch(e) {
      console.error('[OCR] Worker ' + (i + 1) + ' failed to start:', e.message);
    }
  }
  workersReady = workerPool.length > 0;
  if (workersReady) {
    console.log('[OCR] ✅ ' + workerPool.length + ' Tesseract workers ready (local, no rate limit)');
  } else {
    console.warn('[OCR] ⚠ No Tesseract workers started — will use OCR.space fallback only');
  }
}

async function withWorker(fn) {
  for (;;) {
    const slot = workerPool.find(w => !w.busy);
    if (slot) {
      slot.busy = true;
      try { return await fn(slot.worker); }
      finally { slot.busy = false; }
    }
    await sleep(50);
  }
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function vtuHeaders(extra) {
  return Object.assign({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept': 'text/html,application/xhtml+xml,*/*;q=0.8',
    'Referer': VTU_INDEX
  }, extra || {});
}

async function fetchWithTimeout(url, options, ms = 30000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try { return await fetch(url, { ...options, signal: ctrl.signal }); }
  finally { clearTimeout(t); }
}

function getCookieString(res) {
  let cookies = [];
  if (typeof res.headers.raw === 'function') {
    const raw = res.headers.raw()['set-cookie'];
    if (Array.isArray(raw)) {
      cookies = raw;
    }
  }
  if (cookies.length === 0) {
    const single = res.headers.get('set-cookie');
    if (single) {
      cookies = single.split(',');
    }
  }
  
  const pairs = cookies.map(c => {
    const part = c.split(';')[0].trim();
    return part;
  }).filter(p => p && !/^(path|expires|domain|secure|httponly|samesite)=/i.test(p));
  
  return [...new Set(pairs)].join('; ');
}

async function getSession() {
  const res = await fetchWithTimeout(VTU_INDEX, { headers: vtuHeaders(), redirect: 'follow', agent: vtuAgent });
  if (!res.ok) throw new Error('VTU index: ' + res.status);
  const html = await res.text();
  const cookie = getCookieString(res);
  
  let token = '1fee1a3d7f68fb90616ece8a12cddbbd2fb7c2ca'; // fallback
  const tokenMatch = html.match(/name=["']Token["']\s+value=["']([^"']+)["']/i) || 
                     html.match(/value=["']([^"']+)["']\s+name=["']Token["']/i) ||
                     html.match(/Token["']?\s*:\s*["']([^"']+)["']/i);
  if (tokenMatch) {
    token = tokenMatch[1];
  }
  
  return { cookie, token };
}

async function fetchCaptchaBuffer(cookie) {
  const res = await fetchWithTimeout(VTU_CAPTCHA + Math.random() + '+' + Date.now(), {
    headers: vtuHeaders({ Accept: 'image/*', Cookie: cookie }), agent: vtuAgent
  });
  if (!res.ok) throw new Error('Captcha fetch: ' + res.status);
  const contentType = res.headers.get('content-type') || '';
  if (contentType && !contentType.toLowerCase().includes('image')) {
    const text = await res.text();
    if (text.includes('IP address') || text.includes('blocked') || text.includes('forbidden') || text.includes('Access Denied')) {
      throw new Error('IP blocked or rate-limited by VTU');
    }
    throw new Error('Captcha fetch returned non-image: ' + contentType);
  }
  return res.buffer();
}

// ─── CAPTCHA SESSIONS ─────────────────────────────────────────────────────────
const captchaSessions = new Map();
function storeCaptchaSession(data) {
  const id = crypto.randomBytes(16).toString('hex');
  captchaSessions.set(id, { ...data, ts: Date.now() });
  for (const [k, v] of captchaSessions) if (Date.now() - v.ts > 300000) captchaSessions.delete(k);
  return id;
}
function getCaptchaSession(id) {
  const s = captchaSessions.get(id);
  return (!s || Date.now() - s.ts > 300000) ? null : s;
}
async function createCaptchaSession() {
  const { cookie, token } = await getSession();
  const capBuf = await fetchCaptchaBuffer(cookie);
  return { sessionId: storeCaptchaSession({ cookie, token }), captchaBase64: capBuf.toString('base64') };
}

// ─── IMAGE PREPROCESSING — only uses Jimp v1 methods that actually exist ──────
// Available: new Jimp({width,height,color}), Jimp.fromBuffer(buf),
//            img.width, img.height, img.getPixelColor(x,y), img.setPixelColor(color,x,y),
//            img.clone(), img.getBuffer(mime)
async function makeVariants(buf) {
  const out = [];
  let base;
  try { base = await Jimp.fromBuffer(buf); }
  catch(e) { console.error('[OCR] Cannot read captcha buffer:', e.message); return out; }

  const srcW = base.width;
  const srcH = base.height;
  const SCALE = 4;

  const thresholds = [90, 110, 130, 150, 170];

  for (const thresh of thresholds) {
    try {
      const img = base.clone();
      
      // Grayscale and Binarize
      for (let sy = 0; sy < srcH; sy++) {
        for (let sx = 0; sx < srcW; sx++) {
          const c = base.getPixelColor(sx, sy);
          const r = (c >>> 24) & 0xFF;
          const g = (c >>> 16) & 0xFF;
          const b = (c >>> 8)  & 0xFF;
          const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
          const v = gray < thresh ? 0x000000FF : 0xFFFFFFFF;
          img.setPixelColor(v, sx, sy);
        }
      }
      
      // Clean isolated single-pixel noise (remove lone black pixels)
      const cleaned = img.clone();
      for (let sy = 1; sy < srcH - 1; sy++) {
        for (let sx = 1; sx < srcW - 1; sx++) {
          if (img.getPixelColor(sx, sy) === 0x000000FF) {
            let blackNeighbors = 0;
            for (let dy = -1; dy <= 1; dy++) {
              for (let dx = -1; dx <= 1; dx++) {
                if ((dx !== 0 || dy !== 0) && img.getPixelColor(sx + dx, sy + dy) === 0x000000FF) {
                  blackNeighbors++;
                }
              }
            }
            if (blackNeighbors < 2) {
              cleaned.setPixelColor(0xFFFFFFFF, sx, sy);
            }
          }
        }
      }

      cleaned.resize({ w: srcW * SCALE, h: srcH * SCALE });
      out.push(await cleaned.getBuffer('image/png'));
    } catch(e) { /* skip bad variant */ }
  }

  return out;
}

// ─── OCR HELPERS ──────────────────────────────────────────────────────────────
function extractCaptchaCodes(text) {
  const codes = new Set();
  const clean = (text || '').replace(/[^a-zA-Z0-9]/g, '');
  if (clean.length === 6) codes.add(clean);
  else if (clean.length > 6) {
    codes.add(clean.slice(0, 6));
    for (let i = 0; i <= clean.length - 6; i++) codes.add(clean.slice(i, i + 6));
  } else if (clean.length >= 4) codes.add(clean);
  return [...codes];
}

// Primary: Tesseract (local, unlimited)
async function ocrTesseract(imgBuf) {
  if (!workersReady) throw new Error('Tesseract not ready');
  return withWorker(async (w) => {
    const codes = new Set();
    // Default PSM 7 is pre-configured at worker level
    try {
      const { data: { text } } = await w.recognize(imgBuf);
      extractCaptchaCodes(text).forEach(c => codes.add(c));
    } catch(e) {}

    // Fallback to PSM 8 only if PSM 7 found nothing
    if (codes.size === 0) {
      try {
        await w.setParameters({ tessedit_pageseg_mode: '8' });
        const { data: { text } } = await w.recognize(imgBuf);
        extractCaptchaCodes(text).forEach(c => codes.add(c));
      } catch(e) {}
      // Restore default PSM 7 parameter on this worker for future requests
      try { await w.setParameters({ tessedit_pageseg_mode: '7' }); } catch(e) {}
    }

    if (!codes.size) throw new Error('Tesseract: no text');
    return [...codes];
  });
}

// Fallback: OCR.space API (500/day free but better at distorted captchas)
async function ocrSpace(imgBuf) {
  const form = new FormData();
  form.append('apikey', OCR_API_KEY);
  form.append('language', 'eng');
  form.append('isOverlayRequired', 'false');
  form.append('OCREngine', '2');
  form.append('scale', 'true');
  form.append('detectOrientation', 'false');
  form.append('file', imgBuf, { filename: 'c.png', contentType: 'image/png' });
  const res = await fetchWithTimeout('https://api.ocr.space/parse/image', {
    method: 'POST', body: form, headers: form.getHeaders()
  }, 20000);
  if (!res.ok) throw new Error('OCR.space HTTP ' + res.status);
  const data = await res.json();
  if (data.IsErroredOnProcessing) throw new Error('OCR.space error');
  const parsed = (data.ParsedResults || [])[0];
  if (!parsed) throw new Error('OCR.space no result');
  const codes = extractCaptchaCodes(parsed.ParsedText || '');
  if (!codes.length) throw new Error('OCR.space: no code');
  return codes;
}

// ─── VTU FORM ─────────────────────────────────────────────────────────────────
async function submitForm(usn, captchaCode, token, cookie) {
  const params = new URLSearchParams({
    js_token: Buffer.from('student_access_' + new Date().getFullYear()).toString('base64'),
    Token: token, lns: usn.toUpperCase(), captchacode: captchaCode
  });
  const res = await fetchWithTimeout(VTU_RESULT, {
    method: 'POST', body: params.toString(), agent: vtuAgent,
    headers: vtuHeaders({ 'Content-Type': 'application/x-www-form-urlencoded', Cookie: cookie }),
    redirect: 'follow'
  });
  if (!res.ok) throw new Error('VTU submit: ' + res.status);
  return res.text();
}

function isValid(html) {
  if (!html || typeof html !== 'string' || html.length < 500) return false;
  if (/invalid captcha|wrong captcha/i.test(html)) return false;
  return html.includes('University Seat Number');
}

async function fetchResultManual(usn, captchaCode, sessionId) {
  const sess = getCaptchaSession(sessionId);
  if (!sess) throw new Error('Captcha session expired. Refresh and try again.');
  
  // Invalidate the session immediately as the VTU server will consume the captcha anyway
  captchaSessions.delete(sessionId);
  
  const html = await submitForm(usn, captchaCode.trim(), sess.token, sess.cookie);
  if (!isValid(html)) throw new Error('Invalid captcha. Check the code and try again.');
  return { html, parsed: parseResult(html) };
}

// ─── PARSE VTU HTML ───────────────────────────────────────────────────────────
function parseResult(html) {
  try {
    if (!html || typeof html !== 'string') return emptyResult('ERROR');
    const usnMatch  = html.match(/University Seat Number[\s\S]*?divTableCell[^>]*>\s*([A-Z0-9]{10})\s*</i);
    const nameMatch = html.match(/Student Name[\s\S]*?divTableCell[^>]*>\s*([^<]{2,60}?)\s*<\/div>/i);
    const semMatch  = html.match(/Semester\s*:\s*(\d+)/i);
    const usn      = usnMatch  ? usnMatch[1].trim()  : '';
    const name     = nameMatch ? nameMatch[1].trim()  : '';
    const semester = semMatch  ? semMatch[1]          : '';
    const subjects = [];
    const rowPat = /class="divTableRow">([\s\S]*?)(?=class="divTableRow"|divTableBody[\s\S]*?<\/div>)/g;
    let rm;
    while ((rm = rowPat.exec(html)) !== null) {
      const cells = [];
      const cp = /class="divTableCell"[^>]*>([\s\S]*?)<\/div>/g;
      let cm;
      while ((cm = cp.exec(rm[1])) !== null)
        cells.push(cm[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim());
      if (cells.length >= 6 && /^[A-Z0-9]{5,12}$/.test(cells[0]) && cells[0] !== 'Subject Code')
        subjects.push({ code: cells[0], name: cells[1]||'', internal: cells[2]||'', external: cells[3]||'', total: cells[4]||'', result: cells[5]||'' });
    }
    const overallResult = subjects.length > 0 && !subjects.some(s => ['F','A','X'].includes(s.result)) ? 'PASS' : 'FAIL';
    return { usn, name, semester, subjects, overallResult };
  } catch(e) {
    console.error('[parseResult]', e.message);
    return emptyResult('ERROR');
  }
}
function emptyResult(status) { return { usn: '', name: '', semester: '', subjects: [], overallResult: status }; }

function generateMockResult(usn) {
  const semesters = ['3', '4', '5', '6'];
  const sem = semesters[Math.floor(Math.random() * semesters.length)];
  const nameSuffixes = ['Kumar', 'Sharma', 'Patil', 'Gowda', 'Reddy', 'Singh', 'Devi'];
  const namePrefixes = ['Rohan', 'Amit', 'Priya', 'Sneha', 'Rahul', 'Anjali', 'Vijay'];
  const name = namePrefixes[Math.floor(Math.random() * namePrefixes.length)] + ' ' + nameSuffixes[Math.floor(Math.random() * nameSuffixes.length)];
  
  const subjects = [
    { code: '21CS' + sem + '1', name: 'Computer Networks', internal: String(18 + Math.floor(Math.random() * 22)), external: String(20 + Math.floor(Math.random() * 60)) },
    { code: '21CS' + sem + '2', name: 'Database Management Systems', internal: String(18 + Math.floor(Math.random() * 22)), external: String(20 + Math.floor(Math.random() * 60)) },
    { code: '21CS' + sem + '3', name: 'Software Engineering', internal: String(18 + Math.floor(Math.random() * 22)), external: String(20 + Math.floor(Math.random() * 60)) },
    { code: '21CS' + sem + '4', name: 'Design and Analysis of Algorithms', internal: String(18 + Math.floor(Math.random() * 22)), external: String(20 + Math.floor(Math.random() * 60)) }
  ];

  subjects.forEach(s => {
    const int = parseInt(s.internal);
    const ext = parseInt(s.external);
    const tot = int + ext;
    s.total = String(tot);
    s.result = (ext >= 35 && tot >= 40) ? 'P' : 'F';
  });

  const overallResult = subjects.some(s => s.result === 'F') ? 'FAIL' : 'PASS';
  const sgpa = (5.5 + Math.random() * 4.0).toFixed(2);

  return { usn, name, semester: sem, subjects, overallResult, sgpa };
}

// ─── FETCH SINGLE USN WITH RETRY ─────────────────────────────────────────────
async function fetchResult(usn) {
  const MAX_ROUNDS = 16;
  for (let round = 1; round <= MAX_ROUNDS; round++) {
    try {
      const { cookie, token } = await getSession();
      const capBuf  = await fetchCaptchaBuffer(cookie);
      const variants = await makeVariants(capBuf);

      const codes = new Set();

      // Try Tesseract on the original and variants sequentially to save worker capacity
      if (workersReady) {
        try {
          const orig = await ocrTesseract(capBuf);
          orig.forEach(c => { if (c.length >= 4) codes.add(c); });
        } catch(e) {}

        if (![...codes].some(c => c.length === 6)) {
          for (const v of variants) {
            try {
              const res = await ocrTesseract(v);
              res.forEach(c => { if (c.length >= 4) codes.add(c); });
              if ([...codes].some(c => c.length === 6)) break;
            } catch(e) {}
          }
        }
      }

      // If Tesseract found nothing, try OCR.space on the original captcha as fallback
      if (codes.size === 0) {
        try {
          const spaceCodes = await ocrSpace(capBuf);
          spaceCodes.forEach(c => { if (c.length >= 4) codes.add(c); });
        } catch(e) { /* OCR.space also failed */ }
      }

      if (codes.size === 0) {
        await sleep(150 * Math.min(round, 4));
        continue;
      }

      // Try 6-char codes first
      const sorted = [...codes].sort((a, b) =>
        (b.length === 6 ? 1 : 0) - (a.length === 6 ? 1 : 0)
      );

      for (const code of sorted) {
        try {
          const html = await submitForm(usn, code, token, cookie);
          if (html && html.length >= 500) {
            // Check for javascript alerts first, which denote wrong captcha or invalid USN
            const alertMatch = html.match(/alert\((["'])(.*?)\1\)/i);
            if (alertMatch) {
              const alertMsg = alertMatch[2].trim();
              console.log(`[VTU] ⚠️ Server alert for ${usn}: "${alertMsg}"`);
              
              const isCaptchaError = /captcha/i.test(alertMsg);
              if (!isCaptchaError) {
                // Not a captcha error -> captcha was correct, but USN has an issue (e.g. invalid USN)
                console.log(`[VTU] ℹ️ ${usn} results not available (alert: "${alertMsg}")`);
                const err = new Error(alertMsg || 'USN not found or results not available');
                err.abort = true;
                throw err;
              }
            } else {
              // No alert, check if it's a valid result page
              if (html.includes('University Seat Number')) {
                console.log('[VTU] ✅ ' + usn + ' (round ' + round + ', code: ' + code + ')');
                return { html, parsed: parseResult(html) };
              }
            }
          }
        } catch(e) {
          if (e.abort) {
            throw e;
          }
          /* try next code */
        }
      }

      await sleep(100 * Math.min(round, 4));
    } catch(e) {
      console.error('[VTU] Round ' + round + ' error for ' + usn + ':', e.message);
      if (e.abort) {
        throw e;
      }
      if (round === MAX_ROUNDS) throw new Error('Failed after ' + MAX_ROUNDS + ' attempts: ' + e.message);
      await sleep(200 * Math.min(round, 4));
    }
  }
  throw new Error('Failed after ' + MAX_ROUNDS + ' attempts for ' + usn);
}

// ─── CONCURRENCY ──────────────────────────────────────────────────────────────
async function runWithConcurrency(tasks, limit) {
  const results = new Array(tasks.length);
  let idx = 0;
  async function worker() {
    while (idx < tasks.length) { const i = idx++; results[i] = await tasks[i](); }
  }
  await Promise.all(Array.from({ length: Math.min(limit, tasks.length) }, () => worker()));
  return results;
}

// ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────
async function getGoogleSheetsClient() {
  if (!fs.existsSync(GOOGLE_CREDS_PATH)) return null;
  try {
    const creds = JSON.parse(fs.readFileSync(GOOGLE_CREDS_PATH, 'utf8'));
    const auth  = new google.auth.GoogleAuth({ credentials: creds, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    return google.sheets({ version: 'v4', auth });
  } catch(e) { console.error('[Sheets] Creds error:', e.message); return null; }
}

async function writeToGoogleSheets(sheetId, rows, sheetName) {
  const sheets = await getGoogleSheetsClient();
  if (!sheets) return { ok: false, error: 'No google_credentials.json found.' };
  try {
    const header = ['USN', 'Student Name', 'Semester', 'Overall Result',
      ...Array.from({ length: 9 }, (_, i) => [`Sub${i+1} Code`,`Sub${i+1} Name`,`Sub${i+1} Int`,`Sub${i+1} Ext`,`Sub${i+1} Total`,`Sub${i+1} Result`]).flat()
    ];
    const dataRows = rows.map(r => {
      const base = [r.usn||'', r.name||'', r.semester||'', r.overallResult||''];
      const subs = [];
      for (let i = 0; i < 9; i++) {
        const s = (r.subjects||[])[i];
        subs.push(...(s ? [s.code, s.name, s.internal, s.external, s.total, s.result] : ['','','','','','']));
      }
      return [...base, ...subs];
    });
    const sn = sheetName || 'Results';
    await sheets.spreadsheets.values.clear({ spreadsheetId: sheetId, range: sn + '!A:AZ' });
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId, range: sn + '!A1', valueInputOption: 'RAW',
      resource: { values: [header, ...dataRows] }
    });
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

function normalizeUSN(s) {
  if (!s || s.length !== 10) return s;
  let clean = s.toUpperCase().trim();
  // Correct C1 to CI typo
  clean = clean.replace(/^([1-9][A-Z]{2}[0-9]{2})C1([0-9]{3})$/, '$1CI$2');
  // Correct 1S to IS typo
  clean = clean.replace(/^([1-9][A-Z]{2}[0-9]{2})1S([0-9]{3})$/, '$1IS$2');
  // Correct E1 to EI typo
  clean = clean.replace(/^([1-9][A-Z]{2}[0-9]{2})E1([0-9]{3})$/, '$1EI$2');
  return clean;
}

// ─── PARSE UPLOADED FILE ──────────────────────────────────────────────────────
// FIX: Strict VTU USN regex — starts with digit, not just any 10 alphanumeric chars
// This prevents student names like ROHANPATIL being mistaken for USNs
function isVtuUSN(s) {
  if (!s || s.length !== 10) return false;
  // VTU USN format: digit + 2 letters + 2 digits + 2-3 letters + 2-3 digits
  // e.g. 1RF25CI002, 4SF23CS001, 1SI22CS001
  if (/^[1-9][A-Z]{2}[0-9]{2}[A-Z]{2,3}[0-9]{2,3}$/.test(s)) return true;
  // Broader fallback: must START with a digit (eliminates pure names like ROHANPATIL)
  if (/^[0-9][A-Z0-9]{9}$/.test(s)) return true;
  return false;
}

function parseUploadedFile(buffer, filename) {
  const ext = path.extname(filename || '').toLowerCase();
  const usns = [];
  // contacts: { USN -> { email, phone } }
  const contacts = {};

  function looksLikeEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim()); }
  function normalizePhone(s) {
    if (!s) return '';
    let d = String(s).replace(/\D/g, '');
    if (d.length === 12 && d.startsWith('91')) d = d.slice(2);
    if (d.length === 11 && d.startsWith('0')) d = d.slice(1);
    return d;
  }
  function looksLikePhone(s) {
    const d = normalizePhone(s);
    return /^[6-9]\d{9}$/.test(d);
  }

  try {
    let rows = [];
    if (ext === '.csv') {
      rows = buffer.toString('utf8').split(/\r?\n/).map(line => line.split(',').map(c => c.trim().replace(/^"|"$/g, '')));
    } else {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    }

    // Detect header row to find email/phone column indices
    let emailCol = -1, phoneCol = -1;
    const headerRow = rows[0] || [];
    for (let c = 0; c < headerRow.length; c++) {
      const h = String(headerRow[c] || '').toLowerCase().trim();
      if (emailCol === -1 && (h.includes('email') || h.includes('mail') || h.includes('gmail'))) emailCol = c;
      if (phoneCol === -1 && (h.includes('phone') || h.includes('mobile') || h.includes('contact') || h.includes('whatsapp') || h.includes('wa') || h.includes('number'))) phoneCol = c;
    }

    for (const row of rows) {
      let rowUSN = '';
      for (let c = 0; c < (row || []).length; c++) {
        let cell = String(row[c] || '').trim().replace(/[^A-Za-z0-9]/g, '').toUpperCase();
        cell = normalizeUSN(cell);
        if (isVtuUSN(cell)) {
          if (!rowUSN) rowUSN = cell;
          if (!usns.includes(cell)) usns.push(cell);
        }
      }
      // If a USN was found in this row, try to pick up email/phone from designated columns
      if (rowUSN) {
        const email = emailCol >= 0 ? String(row[emailCol] || '').trim() : '';
        const phone = phoneCol >= 0 ? String(row[phoneCol] || '').trim().replace(/\D/g, '') : '';
        // Also scan entire row for email/phone if columns not found by header
        let rowEmail = looksLikeEmail(email) ? email : '';
        let rowPhone = looksLikePhone(phone) ? normalizePhone(phone) : '';
        if (!rowEmail || !rowPhone) {
          for (const cell of (row || [])) {
            const s = String(cell || '').trim();
            if (!rowEmail && looksLikeEmail(s)) rowEmail = s;
            if (!rowPhone && looksLikePhone(s)) rowPhone = normalizePhone(s);
          }
        }
        if (rowEmail || rowPhone) contacts[rowUSN] = { email: rowEmail, phone: rowPhone };
      }
    }
  } catch(e) { console.error('[parseFile]', e.message); }

  const unique = [...new Set(usns)];
  const contactCount = Object.keys(contacts).length;
  console.log('[Upload] Found', unique.length, 'valid VTU USNs,', contactCount, 'with parent contacts');
  return { usns: unique, contacts };
}

// ─── HTTP SERVER ──────────────────────────────────────────────────────────────
const server = http.createServer(async function(req, res) {
  function jsonRes(status, data) {
    res.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
    res.end(JSON.stringify(data));
  }

  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST,GET', 'Access-Control-Allow-Headers': 'Content-Type' });
    res.end(); return;
  }

  if (req.method === 'GET' && req.url === '/api/captcha') {
    createCaptchaSession()
      .then(d => jsonRes(200, { success: true, sessionId: d.sessionId, captcha: d.captchaBase64 }))
      .catch(e => jsonRes(500, { success: false, error: e.message }));
    return;
  }

  if (req.method === 'GET' && req.url === '/api/vtu-announcements') {
    fetchWithTimeout('https://results.vtu.ac.in/index.php', { headers: vtuHeaders(), agent: vtuAgent }, 15000)
      .then(async res => {
        if (!res.ok) throw new Error('VTU index fetch failed: ' + res.status);
        const html = await res.text();
        const updatedMatch = html.match(/Results updated on\s*:\s*([^<]+)/i);
        const updatedOn = updatedMatch ? updatedMatch[1].trim() : 'Unknown';
        
        const announcements = [];
        const itemRegex = /onclick="window\.open\('([^']+)'[^>]*>[\s\S]*?<b>\s*([\s\S]*?)\s*<\/b>/gi;
        let match;
        while ((match = itemRegex.exec(html)) !== null) {
          let link = match[1];
          let title = match[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
          if (!link.startsWith('http')) {
            link = 'https://results.vtu.ac.in/' + link;
          }
          announcements.push({ title, link });
        }
        jsonRes(200, { success: true, updatedOn, announcements });
      })
      .catch(e => jsonRes(500, { success: false, error: e.message }));
    return;
  }

  if (req.method === 'POST' && req.url === '/api/vtu-result') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      let p = {};
      try { p = JSON.parse(body); } catch(e) {}
      let usn = (p.usn || '').trim().toUpperCase();
      usn = normalizeUSN(usn);
      if (!usn || !/^[A-Za-z0-9]{10}$/.test(usn))
        return jsonRes(400, { success: false, error: 'USN must be exactly 10 alphanumeric characters.' });
      
      if (p.mock === true) {
        const student = generateMockResult(usn);
        // Generate mock HTML that parseResult can parse successfully
        const rows = student.subjects.map(s => `
          <div class="divTableRow">
            <div class="divTableCell">${s.code}</div>
            <div class="divTableCell">${s.name}</div>
            <div class="divTableCell">${s.internal}</div>
            <div class="divTableCell">${s.external}</div>
            <div class="divTableCell">${s.total}</div>
            <div class="divTableCell">${s.result}</div>
          </div>
        `).join('');
        const html = `
          <html>
            <body>
              <table>
                <tr>
                  <td class="divTableCell">University Seat Number</td>
                  <td class="divTableCell">${usn}</td>
                </tr>
                <tr>
                  <td class="divTableCell">Student Name</td>
                  <td class="divTableCell">${student.name}</div>
                </tr>
              </table>
              <div>Semester : ${student.semester}</div>
              <div class="divTableBody">${rows}</div>
            </body>
          </html>
        `;
        return jsonRes(200, { success: true, html, parsed: student });
      }

      const run = (p.captcha && p.sessionId)
        ? fetchResultManual(usn, p.captcha, p.sessionId)
        : fetchResult(usn);
      run.then(({ html, parsed }) => jsonRes(200, { success: true, html, parsed }))
         .catch(e => jsonRes(500, { success: false, error: e.message }));
    });
    return;
  }

  if (req.method === 'POST' && req.url === '/api/upload-sheet') {
    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', async () => {
      try {
        const data = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        const { usns, contacts } = parseUploadedFile(Buffer.from(data.fileData, 'base64'), data.fileName);
        if (!usns.length) return jsonRes(400, { success: false, error: 'No valid VTU USNs found. Make sure USNs start with a digit (e.g. 1RF25CI001).' });
        jsonRes(200, { success: true, usns, contacts, count: usns.length, contactCount: Object.keys(contacts).length });
      } catch(e) { jsonRes(500, { success: false, error: e.message }); }
    });
    return;
  }

  if (req.method === 'GET' && req.url.startsWith('/api/bulk-stream')) {
    const urlObj  = new URL(req.url, 'http://localhost');
    const usnList = (urlObj.searchParams.get('usns') || '')
      .split(',').map(u => normalizeUSN(u.trim().toUpperCase())).filter(u => /^[A-Z0-9]{10}$/.test(u));
    const sheetId = urlObj.searchParams.get('sheetId') || GOOGLE_SHEET_ID;
    const useMock = urlObj.searchParams.get('mock') === 'true';
    const PARALLEL = useMock ? 10 : 2; // allow faster mock loading

    res.writeHead(200, {
      'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache',
      'Connection': 'keep-alive', 'Access-Control-Allow-Origin': '*'
    });

    let closed = false;
    req.on('close', () => { closed = true; });

    function send(event, data) {
      if (!closed) res.write('event: ' + event + '\ndata: ' + JSON.stringify(data) + '\n\n');
    }

    send('start', { total: usnList.length, usns: usnList });
    let done = 0, passCount = 0, failCount = 0, errCount = 0;

    (async () => {
      try {
        const tasks = usnList.map(usn => async () => {
          if (closed) return emptyResult('ERROR');
          send('progress', { usn, status: 'fetching', done, total: usnList.length });
          try {
            let parsed;
            if (useMock) {
              await sleep(150); // simulate fast processing time
              parsed = generateMockResult(usn);
            } else {
              const res = await fetchResult(usn);
              parsed = res.parsed;
            }
            done++;
            parsed.overallResult === 'PASS' ? passCount++ : failCount++;
            send('result', { usn, status: 'done', data: parsed, done, total: usnList.length, passCount, failCount, errCount });
            return parsed;
          } catch(e) {
            done++; errCount++;
            const err = { ...emptyResult('ERROR'), usn, error: e.message };
            send('result', { usn, status: 'error', data: err, error: e.message, done, total: usnList.length, passCount, failCount, errCount });
            console.error('[BULK] ❌ ' + usn + ':', e.message);
            return err;
          }
        });

        const allParsed = await runWithConcurrency(tasks, PARALLEL);
        allParsed.sort((a, b) => {
          if (!a || !a.usn) return 1;
          if (!b || !b.usn) return -1;
          const aMatch = a.usn.match(/\d+$/);
          const bMatch = b.usn.match(/\d+$/);
          const aNum = aMatch ? parseInt(aMatch[0], 10) : 0;
          const bNum = bMatch ? parseInt(bMatch[0], 10) : 0;
          if (aNum === bNum) {
            return a.usn.localeCompare(b.usn);
          }
          return aNum - bNum;
        });

        let sheetsResult = { ok: false, error: 'Not configured' };
        if (sheetId && !closed) {
          send('sheets', { status: 'writing' });
          sheetsResult = await writeToGoogleSheets(sheetId, allParsed.filter(Boolean), 'VTU Results');
        }
        send('complete', { total: usnList.length, done: allParsed.length, sheetsResult, results: allParsed });
        res.end();
      } catch(e) {
        console.error('[BULK] Fatal:', e.message);
        send('error', { error: e.message });
        res.end();
      }
    })();
    return;
  }

  if (req.method === 'POST' && req.url === '/api/save-to-sheets') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const { results, sheetId, sheetName } = JSON.parse(body);
        const id = sheetId || GOOGLE_SHEET_ID;
        if (!id) return jsonRes(400, { success: false, error: 'No Google Sheet ID provided.' });
        const r = await writeToGoogleSheets(id, results || [], sheetName || 'VTU Results');
        jsonRes(200, { success: r.ok, error: r.error });
      } catch(e) { jsonRes(500, { success: false, error: e.message }); }
    });
    return;
  }

  // ─── NOTIFICATION HELPERS ──────────────────────────────────────────────────

  function buildEmailHtml(student) {
    const { usn, name, semester, subjects, overallResult, sgpa } = student;
    const resultColor = overallResult === 'PASS' ? '#10b981' : '#ef4444';
    const resultIcon  = overallResult === 'PASS' ? '✅' : '❌';
    const subRows = (subjects || []).map(s => {
      const color = s.result === 'P' ? '#10b981' : '#ef4444';
      const icon  = s.result === 'P' ? '✅' : '❌';
      return `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;color:#c8cadc">${s.code}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;color:#e8eaf6">${s.name}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;color:#8b93b8;text-align:center">${s.internal}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;color:#8b93b8;text-align:center">${s.external}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;font-weight:700;text-align:center;color:${color}">${s.total}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #2a2d4a;font-size:13px;text-align:center">${icon}</td>
      </tr>`;
    }).join('');

    return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"/></head>
<body style="margin:0;padding:0;background:#0f0f1a;font-family:'Segoe UI',Arial,sans-serif">
<div style="max-width:600px;margin:24px auto;background:#16182a;border-radius:16px;overflow:hidden;border:1px solid #2a2d4a">
  <div style="background:linear-gradient(135deg,#6c63ff,#a78bfa);padding:28px 32px">
    <h1 style="margin:0;color:#fff;font-size:22px;font-weight:700">VTU Result Notification</h1>
    <p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:14px">Semester ${semester} — MJ 2026</p>
  </div>
  <div style="padding:28px 32px">
    <p style="color:#8b93b8;font-size:14px;margin:0 0 20px">Dear Parent / Guardian,</p>
    <p style="color:#e8eaf6;font-size:14px;line-height:1.6;margin:0 0 24px">
      The result for your ward <strong style="color:#a78bfa">${name}</strong>
      (USN: <code style="background:#1c1f35;padding:2px 8px;border-radius:4px;color:#a78bfa">${usn}</code>)
      for <strong>Semester ${semester}</strong> has been declared.
    </p>
    <div style="display:flex;gap:16px;margin-bottom:24px;flex-wrap:wrap">
      <div style="flex:1;min-width:140px;background:#1c1f35;border-radius:12px;padding:16px;border:1px solid #2a2d4a;text-align:center">
        <div style="font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Overall Result</div>
        <div style="font-size:22px;font-weight:700;color:${resultColor}">${resultIcon} ${overallResult}</div>
      </div>
      ${sgpa !== null && sgpa !== undefined ? `<div style="flex:1;min-width:140px;background:#1c1f35;border-radius:12px;padding:16px;border:1px solid #2a2d4a;text-align:center">
        <div style="font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">SGPA</div>
        <div style="font-size:22px;font-weight:700;color:#a78bfa">${sgpa}</div>
      </div>` : ''}
    </div>
    <h3 style="color:#e8eaf6;font-size:14px;margin:0 0 12px">Subject-wise Performance</h3>
    <table style="width:100%;border-collapse:collapse;background:#1c1f35;border-radius:10px;overflow:hidden;border:1px solid #2a2d4a">
      <thead>
        <tr style="background:#20234a">
          <th style="padding:10px 12px;text-align:left;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Code</th>
          <th style="padding:10px 12px;text-align:left;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Subject</th>
          <th style="padding:10px 12px;text-align:center;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Int</th>
          <th style="padding:10px 12px;text-align:center;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Ext</th>
          <th style="padding:10px 12px;text-align:center;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Total</th>
          <th style="padding:10px 12px;text-align:center;font-size:11px;color:#8b93b8;text-transform:uppercase;letter-spacing:.06em">Pass</th>
        </tr>
      </thead>
      <tbody>${subRows}</tbody>
    </table>
    <p style="color:#8b93b8;font-size:12px;margin-top:24px;padding-top:16px;border-top:1px solid #2a2d4a">
      This is an automated notification from VTU Result Fetcher.<br/>
      Results are provisional. Visit <a href="https://results.vtu.ac.in" style="color:#a78bfa">results.vtu.ac.in</a> for official records.
    </p>
  </div>
</div>
</body></html>`;
  }

  function buildSMSText(student) {
    const { usn, name, semester, subjects, overallResult, sgpa } = student;
    const subSummary = (subjects || []).slice(0, 6).map(s =>
      s.code.slice(-4) + ':' + s.total + '(' + (s.result === 'P' ? 'P' : 'F') + ')'
    ).join(' ');
    let msg = `VTU Result: ${name} (${usn}) Sem${semester} - ${overallResult}`;
    if (sgpa !== null && sgpa !== undefined) msg += ` SGPA:${sgpa}`;
    if (subSummary) msg += ` | ${subSummary}`;
    msg += ' -VTU Fetcher';
    return msg.slice(0, 320); // Fast2SMS limit
  }

  async function sendEmail(transporter, to, student) {
    const info = await transporter.sendMail({
      from: transporter.options.auth.user,
      to,
      subject: `VTU Result — ${student.name} | Semester ${student.semester} | ${student.overallResult}`,
      html: buildEmailHtml(student)
    });
    return info.messageId;
  }

  async function sendSMS(apiKey, phone, student) {
    const message = buildSMSText(student);
    const resp = await axios.post('https://www.fast2sms.com/dev/bulkV2', {
      route: 'q',
      message,
      language: 'english',
      flash: 0,
      numbers: phone
    }, {
      headers: { authorization: apiKey },
      timeout: 15000
    });
    if (!resp.data.return) throw new Error(resp.data.message || 'Fast2SMS error');
    return resp.data.request_id;
  }

  // ─── /api/send-notifications ──────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/send-notifications') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const { students, gmailUser, gmailPass, fast2smsKey } = JSON.parse(body);
        if (!students || !students.length) return jsonRes(400, { success: false, error: 'No students provided.' });

        // Build Gmail transporter once if credentials given
        let transporter = null;
        if (gmailUser && gmailPass) {
          transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { user: gmailUser, pass: gmailPass }
          });
        }

        const results = [];
        for (const student of students) {
          const r = { usn: student.usn, name: student.name, emailSent: false, smsSent: false, emailError: null, smsError: null };
          // Send email
          if (transporter && student.parentEmail) {
            try { await sendEmail(transporter, student.parentEmail, student); r.emailSent = true; }
            catch(e) { r.emailError = e.message; }
          }
          // Send SMS
          if (fast2smsKey && student.parentPhone) {
            try { await sendSMS(fast2smsKey, student.parentPhone, student); r.smsSent = true; }
            catch(e) { r.smsError = e.message; }
          }
          results.push(r);
          await sleep(200); // avoid rate limits
        }

        const emailCount = results.filter(r => r.emailSent).length;
        const smsCount   = results.filter(r => r.smsSent).length;
        console.log(`[Notify] Sent ${emailCount} emails, ${smsCount} SMS`);
        jsonRes(200, { success: true, results, emailCount, smsCount });
      } catch(e) { jsonRes(500, { success: false, error: e.message }); }
    });
    return;
  }

  // ─── /api/test-notification ───────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/test-notification') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const { gmailUser, gmailPass, fast2smsKey, testEmail, testPhone } = JSON.parse(body);
        const testStudent = {
          usn: 'TEST12345X', name: 'Test Student', semester: '5',
          overallResult: 'PASS', sgpa: 8.5,
          subjects: [
            { code: 'CS501', name: 'Data Structures', internal: '25', external: '55', total: '80', result: 'P' },
            { code: 'CS502', name: 'Computer Networks', internal: '20', external: '48', total: '68', result: 'P' },
            { code: 'CS503', name: 'Operating Systems', internal: '18', external: '35', total: '53', result: 'P' },
          ]
        };

        const out = { emailSent: false, smsSent: false, emailError: null, smsError: null };

        if (gmailUser && gmailPass && testEmail) {
          try {
            const t = nodemailer.createTransport({ service: 'gmail', auth: { user: gmailUser, pass: gmailPass } });
            await sendEmail(t, testEmail, testStudent);
            out.emailSent = true;
          } catch(e) { out.emailError = e.message; }
        }

        if (fast2smsKey && testPhone) {
          try { await sendSMS(fast2smsKey, testPhone, testStudent); out.smsSent = true; }
          catch(e) { out.smsError = e.message; }
        }

        jsonRes(200, { success: true, ...out });
      } catch(e) { jsonRes(500, { success: false, error: e.message }); }
    });
    return;
  }

  // ─── /api/wa-status ───────────────────────────────────────────────────────
  if (req.method === 'GET' && req.url === '/api/wa-status') {
    jsonRes(200, {
      state: waState,
      qr: waQRData || null,
      name: (waState === 'ready' && waClient?.info?.pushname) || null
    });
    return;
  }

  // ─── /api/wa-connect ──────────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/wa-connect') {
    if (waState === 'ready') { jsonRes(200, { success: true, state: 'ready' }); return; }
    if (waState === 'qr')    { jsonRes(200, { success: true, state: 'qr', qr: waQRData }); return; }
    // Start fresh connection
    waState = 'connecting';
    initWA().catch(e => { waState = 'error'; console.error('[WA] Connect error:', e.message); });
    jsonRes(200, { success: true, state: 'connecting' });
    return;
  }

  // ─── /api/wa-disconnect ───────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/wa-disconnect') {
    try {
      if (waClient) await waClient.destroy();
    } catch(e) { /* ignore */ }
    waClient  = null;
    waState   = 'disconnected';
    waQRData  = null;
    jsonRes(200, { success: true });
    return;
  }

  // ─── /api/test-whatsapp ───────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/test-whatsapp') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        if (waState !== 'ready') return jsonRes(400, { success: false, error: 'WhatsApp not connected. Please scan the QR code first in Notify Parents tab.' });
        const { testPhone } = JSON.parse(body);
        if (!testPhone) return jsonRes(400, { success: false, error: 'Please enter a valid test phone number.' });

        let digits = String(testPhone).replace(/\D/g, '');
        if (digits.length === 12 && digits.startsWith('91')) digits = digits.slice(2);
        if (digits.length === 11 && digits.startsWith('0')) digits = digits.slice(1);

        if (!/^[6-9]\d{9}$/.test(digits)) {
          return jsonRes(400, { success: false, error: 'Invalid 10-digit phone number. Please enter a valid Indian mobile number.' });
        }

        const waId = '91' + digits + '@c.us';
        const sampleStudent = {
          usn: '1VT21CS001',
          name: 'Sample Student',
          semester: '5',
          overallResult: 'PASS',
          sgpa: '8.75',
          subjects: [
            { code: '21CS51', name: 'Software Engineering', internal: '45', external: '42', total: '87', result: 'P' },
            { code: '21CS52', name: 'Computer Networks', internal: '42', external: '38', total: '80', result: 'P' },
            { code: '21CS53', name: 'Database Management', internal: '48', external: '45', total: '93', result: 'P' },
            { code: '21CSL57', name: 'DBMS Laboratory', internal: '49', external: '48', total: '97', result: 'P' }
          ]
        };

        const msg = buildWAMessage(sampleStudent);
        await waClient.sendMessage(waId, msg);
        console.log('[WA] Test message sent successfully to', waId);
        jsonRes(200, { success: true, message: 'Test WhatsApp message sent successfully to ' + digits + '!' });
      } catch(e) {
        console.error('[WA] Test message error:', e.message);
        jsonRes(500, { success: false, error: e.message });
      }
    });
    return;
  }

  // ─── /api/send-whatsapp ───────────────────────────────────────────────────
  if (req.method === 'POST' && req.url === '/api/send-whatsapp') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        if (waState !== 'ready') return jsonRes(400, { success: false, error: 'WhatsApp not connected. Scan the QR code in Notify Parents tab first.' });
        const { students } = JSON.parse(body);
        if (!students || !students.length) return jsonRes(400, { success: false, error: 'No students provided.' });

        const results = [];
        for (const student of students) {
          const r = { usn: student.usn, name: student.name, phone: student.parentPhone || '', sent: false, error: null };
          if (!student.parentPhone) { r.error = 'No parent phone number provided'; results.push(r); continue; }
          try {
            let digits = String(student.parentPhone).replace(/\D/g, '');
            if (digits.length === 12 && digits.startsWith('91')) digits = digits.slice(2);
            if (digits.length === 11 && digits.startsWith('0')) digits = digits.slice(1);

            if (!/^[6-9]\d{9}$/.test(digits)) {
              r.error = 'Invalid 10-digit mobile number (' + student.parentPhone + ')';
              results.push(r);
              continue;
            }

            const waId = '91' + digits + '@c.us';
            const msg  = buildWAMessage(student);
            await waClient.sendMessage(waId, msg);
            r.sent = true;
            await sleep(500); // delay between messages to respect WA rate limits
          } catch(e) { r.error = e.message || 'Failed to send message'; }
          results.push(r);
        }
        const sentCount = results.filter(r => r.sent).length;
        const failedCount = results.filter(r => !r.sent).length;
        console.log('[WA] Sent ' + sentCount + ' / ' + results.length + ' WhatsApp messages (' + failedCount + ' failed)');
        jsonRes(200, { success: true, results, sentCount, failedCount });
      } catch(e) { jsonRes(500, { success: false, error: e.message }); }
    });
    return;
  }

  // Static files

  let urlPath = req.url.split('?')[0];
  if (urlPath === '/') urlPath = '/bulk.html';
  const filePath = path.normalize(path.join(PUBLIC_DIR, urlPath));
  if (!filePath.startsWith(PUBLIC_DIR)) { res.writeHead(403); res.end('Forbidden'); return; }
  const ext = path.extname(filePath).toLowerCase();
  fs.readFile(filePath, (err, content) => {
    if (err) { res.writeHead(err.code === 'ENOENT' ? 404 : 500, { 'Content-Type': 'text/plain' }); res.end(err.code === 'ENOENT' ? '404 Not Found' : 'Error'); }
    else { res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] || 'application/octet-stream', 'Cache-Control': 'no-cache', 'Access-Control-Allow-Origin': '*' }); res.end(content); }
  });
});

// ─── WhatsApp message builder ─────────────────────────────────────────────────
function buildWAMessage(student) {
  const { usn, name, semester, subjects, overallResult, sgpa } = student;
  const isPass = (overallResult || '').toUpperCase() === 'PASS';
  const resultEmoji = isPass ? '✅' : '❌';
  let msg = `🎓 *VTU RESULT NOTIFICATION*\n`;
  if (semester) msg += `*Semester:* ${semester}\n`;
  msg += `\nDear Parent / Guardian,\n`;
  msg += `The semester examination result for your ward has been announced:\n\n`;
  msg += `👤 *Student Name:* ${name || 'Student'}\n`;
  msg += `🆔 *USN:* ${usn}\n`;
  msg += `📊 *Overall Result:* ${resultEmoji} *${overallResult || 'N/A'}*\n`;
  if (sgpa !== null && sgpa !== undefined && sgpa !== '' && sgpa !== 'N/A') {
    msg += `⭐ *SGPA:* *${sgpa}*\n`;
  }
  
  if (subjects && subjects.length > 0) {
    msg += `\n📚 *Subject-wise Performance:*\n`;
    msg += `─────────────────────────\n`;
    for (const s of subjects) {
      const p = (s.result === 'P' || (s.result || '').toUpperCase() === 'PASS');
      const icon = p ? '✅' : '❌';
      const code = s.code || '';
      const sName = s.name ? ` (${s.name})` : '';
      const marks = `*${s.total || 0}* [Int: ${s.internal || 0}, Ext: ${s.external || 0}]`;
      msg += `${icon} *${code}*${sName}\n   Marks: ${marks} | Result: *${s.result || (p ? 'P' : 'F')}*\n`;
    }
    msg += `─────────────────────────\n`;
  }
  msg += `\n_Note: This is a provisional result notification. For official records, please visit https://results.vtu.ac.in_`;
  return msg;
}

server.on('error', e => {
  if (e.code === 'EADDRINUSE') console.error('❌ Port ' + PORT + ' already in use.');
  else console.error('❌ Server error:', e.message);
});

initWorkers().then(() => {
  server.listen(PORT, '0.0.0.0', () => {
    console.log('\n==================================================');
    console.log('   VTU RESULT CHECKER — SERVER LIVE              ');
    console.log('   OCR: Local Tesseract + OCR.space fallback     ');
    console.log('==================================================');
    console.log('  Bulk: http://localhost:' + PORT + '/bulk.html');
    console.log('==================================================\n');
  });
  // Init WhatsApp in background (non-blocking)
  initWA().catch(e => console.error('[WA] Background init error:', e.message));
}).catch(e => {
  console.error('❌ Failed to init Tesseract:', e.message);
  process.exit(1);
});
