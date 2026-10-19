/**
 * server.js - Máy chủ phát triển HTTPS cục bộ cho Office Add-in
 * Tích hợp AI Engine Controller & Proxy cho Qwen2.5-VL-3B OCR Local
 */

const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn, exec } = require('child_process');
const { detectAndSyncMachine, loadProfiles } = require('./scripts/detect-machine');

const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;

// Cấu hình đường dẫn QuetOCR & AI Models
const QUETOCR_DIR = 'G:\\My Drive\\Github\\QuetOCR';
const QUETOCR_MODELS_DIR = path.join(QUETOCR_DIR, 'models');
const QUETOCR_SERVER_DIR = path.join(QUETOCR_DIR, 'server');
const LLAMA_EXE = path.join(QUETOCR_SERVER_DIR, 'llama-server.exe');

const OCR_MODELS = {
  qwen3b: {
    name: 'Qwen2.5-VL-3B-Instruct',
    label: 'Qwen2.5-VL-3B (Khuyên dùng - Nhanh, Tiết kiệm RAM)',
    model: path.join(QUETOCR_MODELS_DIR, 'Qwen2.5-VL-3B-Instruct-Q4_K_M.gguf'),
    mmproj: path.join(QUETOCR_MODELS_DIR, 'mmproj-Qwen2.5-VL-3B-Instruct-Q8_0.gguf')
  },
  qwen: {
    name: 'Qwen2.5-VL-7B-Instruct',
    label: 'Qwen2.5-VL-7B (Toàn Năng - Bảng Biểu & Đa Trang)',
    model: path.join(QUETOCR_MODELS_DIR, 'Qwen2.5-VL-7B-Instruct-Q4_K_M.gguf'),
    mmproj: path.join(QUETOCR_MODELS_DIR, 'mmproj-Qwen2.5-VL-7B-Instruct-Q8_0.gguf')
  }
};

let activeModelKey = 'qwen3b';
let isStartingServer = false;

// Bảng ánh xạ MIME Types chuẩn cho Office Add-in
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.wasm': 'application/wasm'
};

function checkLlamaHealth() {
  return new Promise((resolve) => {
    const req = http.get('http://127.0.0.1:8080/health', { timeout: 1500 }, (res) => {
      resolve({
        ready: res.statusCode === 200,
        loading: res.statusCode === 503
      });
    });
    req.on('error', () => resolve({ ready: false, loading: false }));
    req.on('timeout', () => { req.destroy(); resolve({ ready: false, loading: false }); });
  });
}

function killLlamaProcess() {
  return new Promise((resolve) => {
    exec('taskkill /F /IM llama-server.exe', () => resolve());
  });
}

async function handleApiOcrStatus(req, res) {
  const health = await checkLlamaHealth();
  const modelsStatus = {};
  for (const [k, v] of Object.entries(OCR_MODELS)) {
    modelsStatus[k] = {
      name: v.name,
      label: v.label,
      installed: fs.existsSync(v.model) && fs.existsSync(v.mmproj)
    };
  }

  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({
    running: health.ready,
    isStarting: isStartingServer || health.loading,
    currentModel: activeModelKey,
    modelLabel: OCR_MODELS[activeModelKey]?.label || activeModelKey,
    models: modelsStatus
  }));
}

async function handleApiOcrStart(req, res, body) {
  let modelKey = 'qwen3b';
  try {
    const parsed = JSON.parse(body || '{}');
    if (parsed.model && OCR_MODELS[parsed.model]) {
      modelKey = parsed.model;
    }
  } catch (_) {}

  // Kiểm tra file model, fallback nếu thiếu
  const selectedConfig = OCR_MODELS[modelKey];
  if (!fs.existsSync(selectedConfig.model) || !fs.existsSync(selectedConfig.mmproj)) {
    if (modelKey === 'qwen3b' && fs.existsSync(OCR_MODELS.qwen.model)) {
      modelKey = 'qwen';
    }
  }

  const cfg = OCR_MODELS[modelKey];

  // Nếu server đang chạy sẵn sàng và đúng model, không cần khởi động lại!
  const alreadyHealthy = await checkLlamaHealth();
  if (alreadyHealthy.ready && activeModelKey === modelKey) {
    console.log(`[OK] AI Server đã chạy sẵn sàng với model [${cfg.name}]. Không cần khởi động lại.`);
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      status: 'ok',
      running: true,
      currentModel: modelKey,
      label: cfg.label
    }));
    return;
  }

  activeModelKey = modelKey;
  await killLlamaProcess();
  isStartingServer = true;

  console.log(`[*] Đang khởi chạy AI Server với model [${cfg.name}]...`);
  const child = spawn(LLAMA_EXE, [
    '-m', cfg.model,
    '--mmproj', cfg.mmproj,
    '--port', '8080',
    '-c', '8192',
    '-ngl', '99',
    '--host', '127.0.0.1'
  ], {
    cwd: QUETOCR_SERVER_DIR,
    detached: true,
    stdio: 'ignore',
    env: {
      ...process.env,
      PATH: `${QUETOCR_SERVER_DIR};${process.env.PATH || ''}`
    }
  });
  child.unref();

  // Chờ server phản hồi sẵn sàng tối đa 60 giây (nạp model 3B từ disk mất khoảng 20-35s)
  for (let i = 0; i < 60; i++) {
    await new Promise(r => setTimeout(r, 1000));
    const healthy = await checkLlamaHealth();
    if (healthy.ready) {
      isStartingServer = false;
      console.log(`[OK] AI Server đã nạp xong [${cfg.name}] và sẵn sàng!`);
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({
        status: 'ok',
        running: true,
        currentModel: modelKey,
        label: cfg.label
      }));
      return;
    }
  }

  isStartingServer = false;
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({
    status: 'starting',
    running: false,
    currentModel: modelKey,
    label: cfg.label
  }));
}

async function handleApiOcrStop(req, res) {
  await killLlamaProcess();
  isStartingServer = false;
  console.log('[*] Đã tắt AI Server để giải phóng RAM/VRAM.');
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({ status: 'stopped', running: false }));
}

function handleApiOcrChat(req, res, body, retryCount = 0) {
  const options = {
    hostname: '127.0.0.1',
    port: 8080,
    path: '/v1/chat/completions',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body)
    }
  };

  const proxyReq = http.request(options, (proxyRes) => {
    // Nếu server trả về 503 (vẫn đang nạp model những giây cuối), tự động thử lại tối đa 15 lần
    if (proxyRes.statusCode === 503 && retryCount < 15) {
      console.log(`[Retry] AI Server đang nạp model (503). Đang đợi và thử lại (${retryCount + 1}/15)...`);
      proxyRes.resume();
      setTimeout(() => {
        handleApiOcrChat(req, res, body, retryCount + 1);
      }, 2000);
      return;
    }

    res.writeHead(proxyRes.statusCode, {
      ...proxyRes.headers,
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-cache'
    });
    proxyRes.pipe(res);
  });

  proxyReq.on('error', (err) => {
    if (retryCount < 15) {
      console.log(`[Retry] Chưa kết nối được tới AI Server. Đang thử lại (${retryCount + 1}/15)...`);
      setTimeout(() => {
        handleApiOcrChat(req, res, body, retryCount + 1);
      }, 2000);
      return;
    }
    res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ error: { message: `Không thể kết nối tới AI Server 8080: ${err.message}` } }));
  });

  proxyReq.write(body);
  proxyReq.end();
}

/**
 * Cấu hình kết nối Supabase cho Server Local
 */
const SUPABASE_URL = (process.env.SUPABASE_URL || 'https://mabnaygyrscdfsqownhy.supabase.co').replace(/\/$/, '');
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';

function fetchUrlContent(targetUrl, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error('Quá nhiều lần chuyển hướng liên tiếp'));
    const parsed = new URL(targetUrl);
    const client = parsed.protocol === 'https:' ? https : http;

    const req = client.get(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'vi,en;q=0.9'
      },
      timeout: 30000
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        let nextUrl = res.headers.location;
        if (nextUrl.startsWith('/')) {
          nextUrl = `${parsed.protocol}//${parsed.host}${nextUrl}`;
        }
        res.resume();
        return resolve(fetchUrlContent(nextUrl, maxRedirects - 1));
      }

      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        resolve({
          statusCode: res.statusCode,
          html: Buffer.concat(chunks).toString('utf8')
        });
      });
      res.on('error', reject);
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Hết thời gian chờ kết nối (Timeout 30s)'));
    });
  });
}

function extractCleanTextFromHtml(html) {
  if (!html) return '';
  let clean = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/<aside\b[^<]*(?:(?!<\/aside>)<[^<]*)*<\/aside>/gi, '')
    .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, '')
    .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, '')
    .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, '')
    .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '');

  const mainMatch = clean.match(/<main[^>]*class="[^"]*page[^"]*"[^>]*>([\s\S]*?)<\/main>/i) ||
                    clean.match(/<div[^>]*class="[^"]*theme-default-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i) ||
                    clean.match(/<div[^>]*class="[^"]*content[^"]*"[^>]*>([\s\S]*?)<\/div>/i) ||
                    clean.match(/<div[^>]*class="[^"]*detail-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i) ||
                    clean.match(/<article[^>]*>([\s\S]*?)<\/article>/i) ||
                    clean.match(/<main[^>]*>([\s\S]*?)<\/main>/i);
  if (mainMatch) clean = mainMatch[1];

  clean = clean
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/(p|div|tr|li)>/gi, '\n')
    .replace(/<br\s*[\/]?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'");

  const lines = clean.split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0);

  return lines.join('\n\n');
}

function patchSupabaseDoc(id, payload) {
  return new Promise((resolve) => {
    const postData = JSON.stringify(payload);
    const parsed = new URL(`${SUPABASE_URL}/rest/v1/van_ban_phap_luat?id=eq.${id}`);
    const req = https.request({
      hostname: parsed.hostname,
      port: 443,
      path: parsed.pathname + parsed.search,
      method: 'PATCH',
      headers: {
        'apikey': SUPABASE_SERVICE_KEY,
        'Authorization': `Bearer ${SUPABASE_SERVICE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation',
        'Content-Length': Buffer.byteLength(postData)
      }
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        try {
          resolve({ status: res.statusCode, data: JSON.parse(text), raw: text });
        } catch (_) {
          resolve({ status: res.statusCode, raw: text });
        }
      });
    });
    req.on('error', (e) => resolve({ status: 500, raw: e.message }));
    req.write(postData);
    req.end();
  });
}

async function handleApiExtractExternalLaw(req, res, body) {
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  try {
    const { id, soHieu, url } = JSON.parse(body || '{}');
    if (!url || !url.startsWith('http')) {
      return res.end(JSON.stringify({ success: false, error: 'Đường dẫn URL không hợp lệ (phải bắt đầu bằng http:// hoặc https://)' }));
    }

    console.log(`[Trích xuất Link ngoài] Đang tải: ${soHieu || id} -> ${url}`);
    const fetchRes = await fetchUrlContent(url);
    if (fetchRes.statusCode !== 200) {
      return res.end(JSON.stringify({ success: false, error: `Trang web phản hồi mã lỗi HTTP ${fetchRes.statusCode}` }));
    }

    const cleanText = extractCleanTextFromHtml(fetchRes.html);
    if (!cleanText || cleanText.length < 50) {
      return res.end(JSON.stringify({ success: false, error: 'Không trích xuất được khối văn bản hợp lệ từ URL này (nội dung chữ quá ngắn)' }));
    }

    // 1. Thử cập nhật kèm is_external_managed
    let patchPayload = {
      duong_dan_gxd: url,
      noi_dung_toan_van: cleanText,
      co_toan_van_html: true,
      is_external_managed: true,
      updated_at: new Date().toISOString()
    };

    let patchRes = await patchSupabaseDoc(id, patchPayload);
    // 2. Nếu schema Supabase chưa có cột is_external_managed, cập nhật mà không gửi cột này
    if (patchRes.status === 400 && patchRes.raw && patchRes.raw.includes('is_external_managed')) {
      delete patchPayload.is_external_managed;
      patchRes = await patchSupabaseDoc(id, patchPayload);
    }

    if (patchRes.status >= 200 && patchRes.status < 300) {
      console.log(`[Trích xuất Link ngoài] ✅ Thành công: ${soHieu} (${cleanText.length} ký tự)`);
      res.end(JSON.stringify({
        success: true,
        id,
        soHieu,
        url,
        textLength: cleanText.length,
        snippet: cleanText.slice(0, 300)
      }));
    } else {
      res.end(JSON.stringify({
        success: false,
        error: `Lỗi cập nhật Supabase (HTTP ${patchRes.status}): ${patchRes.raw || ''}`
      }));
    }
  } catch (err) {
    console.error('[Trích xuất Link ngoài] Thất bại:', err.message);
    res.end(JSON.stringify({ success: false, error: err.message }));
  }
}

/**
 * Xử lý yêu cầu HTTP/HTTPS
 */
function handleRequest(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const parsedUrl = new URL(req.url, `https://${req.headers.host || 'localhost'}`);
  let pathname = decodeURIComponent(parsedUrl.pathname);

  // API Thông số hệ thống & Nhận diện máy tính
  if (pathname === '/api/system/machine-info' && req.method === 'GET') {
    const profiles = loadProfiles();
    const activeId = profiles.lastActiveMachineId;
    const currentMachine = activeId ? profiles.machines[activeId] : null;
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      currentMachine,
      allMachines: profiles.machines
    }));
    return;
  }

  // Xử lý các API OCR
  if (pathname === '/api/ocr/status' && req.method === 'GET') {
    handleApiOcrStatus(req, res);
    return;
  }

  if (pathname === '/api/ocr/stop' && req.method === 'POST') {
    handleApiOcrStop(req, res);
    return;
  }

  if ((pathname === '/api/ocr/start' || pathname === '/api/ocr/chat') && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      if (pathname === '/api/ocr/start') {
        handleApiOcrStart(req, res, body);
      } else {
        handleApiOcrChat(req, res, body);
      }
    });
    return;
  }

  // Xử lý trích xuất văn bản từ link ngoài tùy chỉnh
  if (pathname === '/api/extract-external-law' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      handleApiExtractExternalLaw(req, res, body);
    });
    return;
  }

  // Phục vụ tệp tĩnh (Static files)
  if (pathname === '/' || pathname === '') {
    pathname = '/taskpane.html';
  }

  const safePath = path.normalize(pathname).replace(/^(\.\.[\/\\])+/, '');
  const filePath = path.join(ROOT_DIR, safePath);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(`404 Not Found: ${pathname}`);
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    const stream = fs.createReadStream(filePath);
    stream.pipe(res);
  });
}

/**
 * Khởi động máy chủ HTTPS
 */
async function startServer() {
  try {
    const devCerts = require('office-addin-dev-certs');
    const httpsOptions = await devCerts.getHttpsServerOptions();

    const server = https.createServer(httpsOptions, handleRequest);
    server.listen(PORT, () => {
      console.log('================================================================');
      console.log(`🚀 MÁY CHỦ HTTPS ĐÃ KHỞI CHẠY THÀNH CÔNG TẠI CỔNG ${PORT}`);
      console.log(`🔗 Taskpane URL   : https://localhost:${PORT}/taskpane.html`);
      console.log(`🔍 OCR Dialog URL : https://localhost:${PORT}/ocr-dialog.html`);
      console.log(`📄 Manifest URL   : https://localhost:${PORT}/manifest.xml`);
      console.log('================================================================');

      // Tự động nhận diện cấu hình phần cứng khi khởi động
      try {
        detectAndSyncMachine();
      } catch (e) {
        console.warn('⚠️ Lỗi nhận diện phần cứng:', e.message);
      }
    });
  } catch (err) {
    console.error('Lỗi khởi động máy chủ:', err);
  }
}

startServer();
