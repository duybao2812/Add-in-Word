/**
 * scripts/update-sample-laws.js
 * Cào toàn văn chữ thực tế cho các văn bản quan trọng hàng đầu và cập nhật thẳng lên Supabase
 */

const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

const SERVICE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';
const SUPABASE_HOST = 'mabnaygyrscdfsqownhy.supabase.co';

function fetchWithRedirect(url, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error('Quá nhiều redirect'));

    const parsed = new URL(url);
    const client = parsed.protocol === 'https:' ? https : http;

    const req = client.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'vi,en;q=0.9'
      },
      timeout: 15000
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
      reject(new Error('Timeout'));
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

function updateSupabaseRecord(soHieu, data) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify(data);
    const options = {
      hostname: SUPABASE_HOST,
      port: 443,
      path: `/rest/v1/van_ban_phap_luat?so_hieu=eq.${encodeURIComponent(soHieu)}`,
      method: 'PATCH',
      headers: {
        'apikey': SERVICE_KEY,
        'Authorization': `Bearer ${SERVICE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation',
        'Content-Length': Buffer.byteLength(payload)
      }
    };

    const req = https.request(options, (res) => {
      let resData = '';
      res.on('data', chunk => resData += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(true);
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${resData}`));
        }
      });
    });

    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function run() {
  const content = fs.readFileSync(path.join(__dirname, '..', 'law-data.js'), 'utf8');
  const sandbox = {};
  new Function('window', content)(sandbox);
  const laws = sandbox.LAW_DATABASE || [];

  // Danh sách các văn bản quan trọng hàng đầu cần cào toàn văn trước
  const targetCodes = [
    'Luật Xây dựng 135/2025',
    'NĐ 217/2026/NĐ-CP',
    'NĐ 212/2026/NĐ-CP',
    'NĐ 206/2026/NĐ-CP',
    'NĐ 210/2026/NĐ-CP',
    'NĐ 207/2026/NĐ-CP',
    'NĐ 193/2026/NĐ-CP',
    'NĐ 175/2026/NĐ-CP',
    'NĐ 140/2025/NĐ-CP',
    'NĐ 10/2021/NĐ-CP',
    'NĐ 15/2021/NĐ-CP',
    'NĐ 06/2021/NĐ-CP',
    'NĐ 50/2021/NĐ-CP',
    'NĐ 37/2015/NĐ-CP',
    'NĐ 35/2023/NĐ-CP'
  ];

  const targetLaws = laws.filter(x => targetCodes.includes(x.code) && x.linkGoc);

  console.log(`Bắt đầu cào toàn văn cho ${targetLaws.length} văn bản trọng điểm...`);

  for (let i = 0; i < targetLaws.length; i++) {
    const doc = targetLaws[i];
    console.log(`[${i + 1}/${targetLaws.length}] Đang cào: ${doc.code} -> ${doc.linkGoc}`);

    try {
      const res = await fetchWithRedirect(doc.linkGoc);
      if (res.status === 200) {
        const text = extractLawContent(res.data);
        console.log(`   -> Trích xuất được ${text.length} ký tự text.`);

        if (text.length > 200) {
          await updateSupabaseRecord(doc.code, {
            noi_dung_toan_van: text,
            co_toan_van_html: true,
            is_deleted: false
          });
          console.log(`   ✅ Đã cập nhật thành công lên Supabase!`);
        } else {
          console.log(`   ⚠️ Nội dung text quá ngắn (${text.length} ký tự).`);
        }
      } else {
        console.warn(`   ❌ Lỗi HTTP: ${res.status}`);
      }
    } catch (e) {
      console.error(`   ❌ Lỗi cào văn bản ${doc.code}:`, e.message);
    }

    await sleep(1500); // Giãn cách nhẹ nhàng
  }

  console.log('🎉 HOÀN THÀNH CẬP NHẬT TOÀN VĂN LÊN SUPABASE!');
}

run();
