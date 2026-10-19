/**
 * scripts/deduplicate-laws.js
 * Tìm và dọn dẹp toàn bộ các bản ghi văn bản bị trùng lặp trên Supabase
 */

const https = require('https');

const SUPABASE_URL = (process.env.SUPABASE_URL || 'https://mabnaygyrscdfsqownhy.supabase.co').replace(/\/$/, '');
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';

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

function normalizeCode(code) {
  if (!code) return '';
  return code
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/\//g, '/')
    .replace(/[-_.]/g, '');
}

async function run() {
  console.log('🔍 Đang kiểm tra trùng lặp trên Supabase...');
  const res = await sendSupabaseRequest('GET', 'van_ban_phap_luat?select=id,so_hieu,ten_van_ban,duong_dan_gxd,co_toan_van_html,noi_dung_toan_van&limit=1000');
  const all = Array.isArray(res.data) ? res.data : [];
  console.log(`Tổng số bản ghi ban đầu: ${all.length}`);

  const map = new Map();
  for (const doc of all) {
    const key = normalizeCode(doc.so_hieu);
    if (!key) continue;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(doc);
  }

  const toDelete = [];
  for (const [key, list] of map.entries()) {
    if (list.length > 1) {
      console.log(`\nPhát hiện trùng: [${list[0].so_hieu}] (${list.length} bản ghi)`);
      // Sắp xếp: Ưu tiên có text dài nhất, link hợp lệ nhất
      list.sort((a, b) => {
        const lenA = (a.noi_dung_toan_van || '').length;
        const lenB = (b.noi_dung_toan_van || '').length;
        if (lenA !== lenB) return lenB - lenA; // dài hơn đứng trước
        const isAHasText = a.co_toan_van_html ? 1 : 0;
        const isBHasText = b.co_toan_van_html ? 1 : 0;
        if (isAHasText !== isBHasText) return isBHasText - isAHasText;
        return 0;
      });

      const best = list[0];
      const dups = list.slice(1);
      console.log(`   -> GIỮ LẠI: ID ${best.id} (${(best.noi_dung_toan_van || '').length} ký tự, ${best.duong_dan_gxd})`);
      for (const d of dups) {
        console.log(`   -> XÓA BẢN TRÙNG: ID ${d.id} (${(d.noi_dung_toan_van || '').length} ký tự, ${d.duong_dan_gxd})`);
        toDelete.push(d.id);
      }
    }
  }

  console.log(`\n======================================================`);
  console.log(`Tổng số bản ghi trùng lặp cần xóa: ${toDelete.length}`);
  console.log(`======================================================`);

  if (toDelete.length > 0) {
    let deletedCount = 0;
    for (const id of toDelete) {
      const delRes = await sendSupabaseRequest('DELETE', `van_ban_phap_luat?id=eq.${id}`);
      if (delRes.status >= 200 && delRes.status < 300) {
        deletedCount++;
      } else {
        console.error(`Lỗi khi xóa ID ${id}: HTTP ${delRes.status}`);
      }
    }
    console.log(`✅ Đã xóa thành công ${deletedCount}/${toDelete.length} bản ghi trùng lặp trên Supabase!`);
  }

  // Tăng revision để các client đồng bộ lại
  await sendSupabaseRequest('PATCH', 'system_metadata?key=eq.global_sync_revision', {
    value: String(Date.now()),
    updated_at: new Date().toISOString()
  });
  console.log('🔄 Đã bump sync revision.');
}

run().catch(console.error);
