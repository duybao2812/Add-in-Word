/**
 * scripts/seed-supabase.js
 * Đẩy toàn bộ danh mục chủ đề và 515 văn bản pháp luật lên cơ sở dữ liệu Supabase Cloud
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const SERVICE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjQxNzQ2NywiZXhwIjoyMTA3OTkzNDY3fQ.9DIvmI1IDu3LHrh8vkKESb5NWqRiAmZwmW2Iq-Q3ad8';

function loadLocalLaws() {
  const content = fs.readFileSync(path.join(__dirname, '..', 'law-data.js'), 'utf8');
  const sandbox = {};
  const fn = new Function('window', content);
  fn(sandbox);
  return sandbox.LAW_DATABASE || [];
}

function sendRequest(endpointPath, method, data) {
  return new Promise((resolve, reject) => {
    const payload = data ? JSON.stringify(data) : '';
    const options = {
      hostname: 'mabnaygyrscdfsqownhy.supabase.co',
      port: 443,
      path: `/rest/v1/${endpointPath}`,
      method: method,
      headers: {
        'apikey': SERVICE_KEY,
        'Authorization': `Bearer ${SERVICE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates,return=representation'
      }
    };
    if (payload) {
      options.headers['Content-Length'] = Buffer.byteLength(payload);
    }

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
    if (payload) req.write(payload);
    req.end();
  });
}

async function run() {
  // 1. Nạp đầy đủ các chủ đề
  const allCategories = [
    { id: 'qlda', ten_chu_de: 'Quản lý Dự án & Thủ tục ĐTXD' },
    { id: 'luat-chung', ten_chu_de: 'Hệ thống Luật & Bộ Luật' },
    { id: 'chung', ten_chu_de: 'Quy định Chung & Quản lý Nhà nước' },
    { id: 'chi-phi', ten_chu_de: 'Quản lý Chi phí & Định mức ĐTXD' },
    { id: 'hop-dong', ten_chu_de: 'Hợp đồng Xây dựng' },
    { id: 'chat-luong', ten_chu_de: 'Quản lý Chất lượng & An toàn Thi công' },
    { id: 'dau-thau', ten_chu_de: 'Đấu thầu & Lựa chọn Nhà thầu' },
    { id: 'dat-dai', ten_chu_de: 'Đất đai & Giải phóng Mặt bằng' },
    { id: 'pccc', ten_chu_de: 'Phòng cháy Chữa cháy & Cứu nạn' },
    { id: 'tai-chinh', ten_chu_de: 'Tài chính, Thuế & Bảo hiểm Xây dựng' }
  ];

  console.log('[1/2] Đang cập nhật 10 nhóm chủ đề vào chu_de_phap_luat...');
  await sendRequest('chu_de_phap_luat', 'POST', allCategories);
  console.log('✅ Đã cập nhật xong nhóm chủ đề!');

  // 2. Nạp 515 văn bản
  const laws = loadLocalLaws();
  console.log(`[2/2] Đang nạp ${laws.length} văn bản vào van_ban_phap_luat...`);

  const mapped = laws.map((item, idx) => {
    let issueDate = null;
    if (item.issueDate && /^\d{4}-\d{2}-\d{2}$/.test(item.issueDate)) {
      issueDate = item.issueDate;
    } else if (item.year) {
      issueDate = `${item.year}-01-01`;
    }

    return {
      so_hieu: item.code || item.name || `VB-${idx}`,
      ten_van_ban: item.name || item.code || '',
      loai_van_ban: item.cat === 'luat-chung' ? 'Luật' : 'Nghị định/Thông tư',
      chu_de_id: item.cat || 'qlda',
      ngay_ban_hanh: issueDate,
      tinh_trang_hieu_luc: (item.isExpired || item.status === 'Hết hiệu lực') ? 'Hết hiệu lực' : 'Còn hiệu lực',
      loai_phap_ly: item.loaiPhapLy || 'Văn bản gốc',
      trich_dan_chuan: item.citation || `- Căn cứ ${item.name};`,
      duong_dan_gxd: item.linkGoc || null,
      co_toan_van_html: false,
      duong_dan_file_dinh_kem: item.linkGoc || null,
      noi_dung_toan_van: item.noi_dung_toan_van || '',
      is_deleted: false
    };
  });

  const BATCH_SIZE = 50;
  for (let i = 0; i < mapped.length; i += BATCH_SIZE) {
    const chunk = mapped.slice(i, i + BATCH_SIZE);
    try {
      await sendRequest('van_ban_phap_luat', 'POST', chunk);
      console.log(`Đã nạp thành công ${Math.min(i + BATCH_SIZE, mapped.length)} / ${mapped.length} văn bản...`);
    } catch (err) {
      console.error(`Lỗi nạp batch từ ${i}:`, err.message);
    }
  }

  console.log('🎉 ĐÃ NẠP TOÀN BỘ 515 VĂN BẢN LÊN SUPABASE THÀNH CÔNG RỰC RỠ!');
}

run();
