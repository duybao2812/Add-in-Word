/**
 * scripts/crawler-gxd.js
 * Bot thu thập văn bản pháp luật tự động từ QLDA GXD (Tầng 1)
 * Phiên bản: Text-First – Anti-WAF – Auto-Redirect (308/301) – Chuẩn hóa UTF-8 Không lỗi Font – Telegram Alert
 */

const https = require('https');
const http = require('http');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

// Cấu hình môi trường từ GitHub Secrets / Process Env
const SUPABASE_URL = (process.env.SUPABASE_URL || 'https://mabnaygyrscdfsqownhy.supabase.co').replace(/\/$/, '');
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';
const TELEGRAM_BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN || '8660476357:AAEynNtesZtuTyjEFPfL4VsPWmn9IxOIbw8';
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID || '7901714823';
const IS_HEARTBEAT_ONLY = process.env.HEARTBEAT_ONLY === 'true';

const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

/**
 * Gửi tin nhắn cảnh báo hoặc báo cáo qua Telegram Bot
 */
async function sendTelegramMessage(text) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;

  const payload = JSON.stringify({
    chat_id: TELEGRAM_CHAT_ID,
    text: text,
    parse_mode: 'HTML'
  });

  return new Promise((resolve) => {
    const req = https.request({
      hostname: 'api.telegram.org',
      port: 443,
      path: `/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
      }
    }, (res) => {
      resolve(res.statusCode === 200);
    });
    req.on('error', (e) => {
      console.warn('[Telegram] Gửi tin nhắn thất bại:', e.message);
      resolve(false);
    });
    req.write(payload);
    req.end();
  });
}

/**
 * Gửi HTTP GET tự động follow chuyển hướng (301, 302, 308) và Decode UTF-8 CHUẨN XÁC KHÔNG LỖI FONT
 */
function fetchWithRedirect(url, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error('Quá nhiều redirect'));

    const parsed = new URL(url);
    const client = parsed.protocol === 'https:' ? https : http;

    const req = client.get(url, {
      headers: {
        'User-Agent': USER_AGENT,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'vi,en;q=0.9'
      },
      timeout: 25000
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        let nextUrl = res.headers.location;
        if (nextUrl.startsWith('/')) {
          nextUrl = `${parsed.protocol}//${parsed.host}${nextUrl}`;
        }
        res.resume();
        return resolve(fetchWithRedirect(nextUrl, maxRedirects - 1));
      }

      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        // Buffer.concat ghép toàn bộ byte thô trước khi decode utf-8, ngăn chặn 100% lỗi chia cắt ký tự tiếng Việt
        const fullHtml = Buffer.concat(chunks).toString('utf8');
        resolve({ status: res.statusCode, data: fullHtml });
      });
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request Timeout sau 25s'));
    });
  });
}

/**
 * Gửi HTTP tới Supabase REST API
 */
function sendSupabaseRequest(method, endpointPath, bodyData) {
  return new Promise((resolve, reject) => {
    const payload = bodyData ? JSON.stringify(bodyData) : '';
    const parsed = new URL(`${SUPABASE_URL}/rest/v1/${endpointPath}`);

    const options = {
      hostname: parsed.hostname,
      port: 443,
      path: parsed.pathname + parsed.search,
      method: method,
      headers: {
        'apikey': SUPABASE_SERVICE_KEY,
        'Authorization': `Bearer ${SUPABASE_SERVICE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation'
      }
    };
    if (payload) {
      options.headers['Content-Length'] = Buffer.byteLength(payload);
    }

    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        try {
          const parsedRes = text ? JSON.parse(text) : {};
          resolve({ status: res.statusCode, data: parsedRes });
        } catch (_) {
          resolve({ status: res.statusCode, raw: text });
        }
      });
    });

    req.on('error', reject);
    if (payload) req.write(payload);
    req.end();
  });
}

/**
 * Trích xuất toàn văn chuẩn VuePress từ trang web QLDA GXD
 */
function extractLawContent(html) {
  if (!html) return '';

  let clean = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/<aside\b[^<]*(?:(?!<\/aside>)<[^<]*)*<\/aside>/gi, '')
    .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, '')
    .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, '')
    .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, '')
    .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
    .replace(/<details\b[^<]*(?:(?!<\/details>)<[^<]*)*<\/details>/gi, '');

  const mainMatch = clean.match(/<main[^>]*class="[^"]*page[^"]*"[^>]*>([\s\S]*?)<\/main>/i) ||
                    clean.match(/<div[^>]*class="[^"]*theme-default-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i) ||
                    clean.match(/<main[^>]*>([\s\S]*?)<\/main>/i);

  if (mainMatch) {
    clean = mainMatch[1];
  }

  clean = clean
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/(p|div|tr)>/gi, '\n')
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

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * LUỒNG CHÍNH CỦA BOT
 */
async function main() {
  const startTime = new Date();
  console.log(`[Crawler GXD] Khởi động bot lúc: ${startTime.toISOString()}`);

  // 1. Chế độ Heartbeat Ping
  if (IS_HEARTBEAT_ONLY) {
    console.log('[Heartbeat] Đang thực hiện Heartbeat Ping để giữ Supabase Database luôn Active...');
    try {
      const ping = await sendSupabaseRequest('GET', 'system_metadata?limit=1');
      console.log(`[Heartbeat] Ping thành công (HTTP ${ping.status}). Cơ sở dữ liệu đang hoạt động hoàn hảo!`);
    } catch (e) {
      console.error('[Heartbeat] Ping thất bại:', e.message);
    }
    return;
  }

  // 2. Chế độ cào và đồng bộ toàn văn cho các văn bản QLDA GXD
  let scannedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  try {
    console.log('[1/3] Đang lấy danh sách các văn bản QLDA GXD từ Supabase...');
    // Chỉ lấy văn bản có link qlda.gxd.vn (bỏ qua toàn bộ link ngoài hoặc văn bản quản lý riêng)
    const res = await sendSupabaseRequest(
      'GET', 
      'van_ban_phap_luat?select=id,so_hieu,ten_van_ban,duong_dan_gxd,co_toan_van_html,hash_noi_dung,updated_at&duong_dan_gxd=like.*qlda.gxd.vn*&order=updated_at.asc&limit=60'
    );
    
    const docsToCrawl = (Array.isArray(res.data) ? res.data : []).filter(d => {
      const url = d.duong_dan_gxd || '';
      return url.includes('qlda.gxd.vn');
    });

    scannedCount = docsToCrawl.length;
    console.log(`[2/3] Tìm thấy ${scannedCount} văn bản QLDA GXD được chọn kiểm tra.`);

    for (let i = 0; i < docsToCrawl.length; i++) {
      const doc = docsToCrawl[i];
      if (!doc.duong_dan_gxd) continue;

      console.log(`[${i + 1}/${scannedCount}] Đang kiểm tra: ${doc.so_hieu} -> ${doc.duong_dan_gxd}`);
      try {
        const pageRes = await fetchWithRedirect(doc.duong_dan_gxd);
        if (pageRes.status === 200) {
          const text = extractLawContent(pageRes.data);
          if (text.length > 200 && !text.includes("There's nothing here")) {
            const newHash = crypto.createHash('sha256').update(text).digest('hex');

            // Nếu chưa có text hoặc hash thay đổi (văn bản được sửa đổi)
            if (!doc.co_toan_van_html || doc.hash_noi_dung !== newHash) {
              await sendSupabaseRequest('PATCH', `van_ban_phap_luat?id=eq.${doc.id}`, {
                noi_dung_toan_van: text,
                co_toan_van_html: true,
                hash_noi_dung: newHash,
                updated_at: new Date().toISOString()
              });
              updatedCount++;
              console.log(`   ✅ Đã cập nhật văn bản (${text.length} ký tự, SHA256: ${newHash.slice(0, 8)}...)!`);
            } else {
              skippedCount++;
              console.log(`   ⚡ Nội dung không đổi, bỏ qua để tiết kiệm băng thông.`);
            }
          } else {
            console.log(`   ℹ️ Trang không có khối chữ HTML dài (${text.length} ký tự).`);
          }
        }
      } catch (err) {
        errorCount++;
        console.error(`   ❌ Lỗi khi cào ${doc.so_hieu}:`, err.message);
      }

      await sleep(1500); // Random Jitter 1.5s
    }

    const endTime = new Date();
    const durationSec = Math.round((endTime - startTime) / 1000);

    const reportMessage = `
📢 <b>BÁO CÁO CRAWLER BOT QLDA GXD</b>
⏰ <b>Thời gian:</b> ${endTime.toLocaleString('vi-VN')} (${durationSec}s)

• <b>Văn bản đã quét:</b> ${scannedCount}
• <b>Văn bản nạp mới / cập nhật:</b> ${updatedCount}
• <b>Văn bản không đổi (giữ nguyên):</b> ${skippedCount}
• <b>Số lỗi phát sinh:</b> ${errorCount}

✅ <b>Trạng thái:</b> ${errorCount === 0 ? 'THÀNH CÔNG' : 'CÓ CẢNH BÁO'}
    `.trim();

    console.log(reportMessage);
    await sendTelegramMessage(reportMessage);

  } catch (err) {
    console.error('[LỖI KHẨN CẤP BOT CÀO]:', err.message);
    const alertMsg = `
🚨 <b>CẢNH BÁO LỖI CRAWLER BOT GXD</b>
❌ Lỗi: <code>${err.message}</code>
    `.trim();
    await sendTelegramMessage(alertMsg);
  }
}

main();
