const fs = require('fs');
const path = require('path');
const https = require('https');

const SUPABASE_URL = 'https://mabnaygyrscdfsqownhy.supabase.co';
const SUPABASE_SERVICE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';

function fetchAllLaws() {
  return new Promise((resolve, reject) => {
    const url = `${SUPABASE_URL}/rest/v1/van_ban_phap_luat?select=*&order=id.asc`;
    https.get(url, {
      headers: {
        'apikey': SUPABASE_SERVICE_KEY,
        'Authorization': `Bearer ${SUPABASE_SERVICE_KEY}`
      }
    }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        try {
          const list = JSON.parse(Buffer.concat(chunks).toString('utf8'));
          resolve(list);
        } catch (e) {
          reject(e);
        }
      });
    }).on('error', reject);
  });
}

async function run() {
  console.log('Đang tải danh sách văn bản từ Supabase...');
  const rows = await fetchAllLaws();
  console.log(`Đã tải ${rows.length} văn bản từ Supabase.`);

  const mapped = rows.map((r, idx) => ({
    id: idx + 1,
    uuid: r.id,
    code: r.so_hieu,
    name: r.ten_van_ban,
    type: r.loai_van_ban || 'Văn bản QPPL',
    category: r.chu_de_id || 'chung',
    issueDate: r.ngay_ban_hanh || '',
    effectiveDate: r.ngay_hieu_luc || '',
    status: r.tinh_trang_hieu_luc || 'Còn hiệu lực',
    issuer: r.co_quan_ban_hanh || 'Cơ quan nhà nước',
    linkGoc: r.duong_dan_gxd || '',
    duong_dan_gxd: r.duong_dan_gxd || '',
    noi_dung_toan_van: r.noi_dung_toan_van || '',
    co_toan_van_html: !!(r.co_toan_van_html && r.noi_dung_toan_van && r.noi_dung_toan_van.length > 50),
    is_external_managed: !!r.is_external_managed,
    citation: `- Căn cứ ${r.ten_van_ban};`
  }));

  const fileContent = `/**
 * Dữ liệu văn bản pháp luật xây dựng (Offline Bundle)
 * Tự động đồng bộ từ Supabase Cloud Database (${new Date().toISOString()})
 */
window.LAW_DATABASE = ${JSON.stringify(mapped, null, 2)};
`;

  const targetPath = path.resolve(__dirname, '../law-data.js');
  fs.writeFileSync(targetPath, fileContent, 'utf8');
  console.log(`✅ Đã cập nhật ${mapped.length} văn bản vào ${targetPath}`);
}

run().catch(console.error);
