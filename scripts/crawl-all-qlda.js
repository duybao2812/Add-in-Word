/**
 * scripts/crawl-all-qlda.js
 * Script cào toàn văn tự động 1 lần duy nhất cho toàn bộ các văn bản thuộc qlda.gxd.vn
 * Bóc tách UTF-8 chuẩn xác, cập nhật trực tiếp vào Supabase Cloud Database
 */

const https = require('https');
const http = require('http');

const SUPABASE_URL = (process.env.SUPABASE_URL || 'https://mabnaygyrscdfsqownhy.supabase.co').replace(/\/$/, '');
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';

const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

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
        const fullHtml = Buffer.concat(chunks).toString('utf8');
        resolve({ status: res.statusCode, data: fullHtml });
      });
    });

    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Timeout sau 25s'));
    });
  });
}

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
                    clean.match(/<div[^>]*class="[^"]*content[^"]*"[^>]*>([\s\S]*?)<\/div>/i) ||
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
          resolve({ status: res.statusCode, data: text ? JSON.parse(text) : {} });
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

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function run() {
  console.log('====================================================================');
  console.log('🚀 BẮT ĐẦU CÀO TOÀN BỘ VĂN BẢN TỪ QLDA GXD VÀO SUPABASE');
  console.log('====================================================================');

  // Lấy toàn bộ các văn bản qlda.gxd.vn chưa có toàn văn
  const res = await sendSupabaseRequest(
    'GET',
    'van_ban_phap_luat?select=id,so_hieu,ten_van_ban,duong_dan_gxd,co_toan_van_html&co_toan_van_html=eq.false&limit=1000&order=id.asc'
  );

  const allDocs = Array.isArray(res.data) ? res.data : [];
  const qldaDocs = allDocs.filter(d => (d.duong_dan_gxd || '').includes('qlda.gxd.vn'));

  console.log(`📌 Tìm thấy tổng cộng: ${qldaDocs.length} văn bản QLDA GXD cần cào toàn văn.`);

  let successCount = 0;
  let emptyCount = 0;
  let errorCount = 0;

  // Chạy lần lượt với độ trễ nhẹ 400ms để không bị Cloudflare / Server chặn
  for (let i = 0; i < qldaDocs.length; i++) {
    const doc = qldaDocs[i];
    const url = doc.duong_dan_gxd;
    process.stdout.write(`[${i + 1}/${qldaDocs.length}] ${doc.so_hieu} ... `);

    try {
      const pageRes = await fetchWithRedirect(url);
      if (pageRes.status === 200) {
        const text = extractLawContent(pageRes.data);
        if (text && text.length > 200 && !text.includes("There's nothing here")) {
          // Lưu vào Supabase
          const patchRes = await sendSupabaseRequest('PATCH', `van_ban_phap_luat?id=eq.${doc.id}`, {
            noi_dung_toan_van: text,
            co_toan_van_html: true,
            updated_at: new Date().toISOString()
          });

          if (patchRes.status >= 200 && patchRes.status < 300) {
            successCount++;
            console.log(`✅ ${text.length.toLocaleString()} ký tự`);
          } else {
            console.log(`❌ Lỗi lưu DB: HTTP ${patchRes.status}`);
            errorCount++;
          }
        } else {
          emptyCount++;
          console.log(`ℹ️ Không có text HTML (sẽ quản lý bằng Link ngoài)`);
        }
      } else {
        emptyCount++;
        console.log(`⚠️ HTTP ${pageRes.status}`);
      }
    } catch (err) {
      errorCount++;
      console.log(`❌ Lỗi mạng: ${err.message}`);
    }

    await sleep(400);
  }

  // Tăng revision trên Supabase
  try {
    await sendSupabaseRequest('PATCH', 'system_metadata?key=eq.global_sync_revision', {
      value: String(Date.now()),
      updated_at: new Date().toISOString()
    });
  } catch (_) {}

  console.log('====================================================================');
  console.log(`🏁 HOÀN TẤT CÀO DỮ LIỆU:`);
  console.log(`   • Đã nạp thành công toàn văn: ${successCount} văn bản`);
  console.log(`   • Văn bản không có HTML (file đính kèm / 404): ${emptyCount} văn bản`);
  console.log(`   • Lỗi mạng / Thất bại: ${errorCount} văn bản`);
  console.log('====================================================================');
}

run();
