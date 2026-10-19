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
    const req = http.get('http://127.0.0.1:8080/health', { timeout: 1200 }, (res) => {
      resolve(res.statusCode === 200);
    });
    req.on('error', () => resolve(false));
    req.on('timeout', () => { req.destroy(); resolve(false); });
  });
}

function killLlamaProcess() {
  return new Promise((resolve) => {
    exec('taskkill /F /IM llama-server.exe', () => resolve());
  });
}

async function handleApiOcrStatus(req, res) {
  const isHealthy = await checkLlamaHealth();
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
    running: isHealthy,
    isStarting: isStartingServer,
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
  if (alreadyHealthy && activeModelKey === modelKey) {
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
    if (healthy) {
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
