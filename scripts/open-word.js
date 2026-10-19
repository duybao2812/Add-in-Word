/**
 * scripts/open-word.js
 * Mở tệp tài liệu Word đã được sideload Add-in một cách chính xác
 */

const cp = require('child_process');
const path = require('path');
const os = require('os');
const fs = require('fs');

const docxName = 'Word add-in 67e1a3bc-1478-4395-885e-fb854cf71391.docx';
const docxPath = path.join(os.tmpdir(), docxName);

console.log('--------------------------------------------------');
console.log('Đang kiểm tra tệp tài liệu Word sideload:');
console.log(docxPath);

if (!fs.existsSync(docxPath)) {
  console.error('❌ Không tìm thấy tệp tài liệu Word sideload.');
  console.log('Vui lòng chạy lại lệnh: npm start để hệ thống tái tạo file.');
  process.exit(1);
}

console.log('✅ Đã tìm thấy tệp tài liệu. Đang khởi chạy Microsoft Word...');

// Tìm đường dẫn WINWORD.EXE
const possiblePaths = [
  'C:\\Program Files\\Microsoft Office\\Root\\Office16\\WINWORD.EXE',
  'C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\WINWORD.EXE',
  'C:\\Program Files\\Microsoft Office\\Office16\\WINWORD.EXE',
  'C:\\Program Files (x86)\\Microsoft Office\\Office16\\WINWORD.EXE'
];

let winwordExe = null;
for (const p of possiblePaths) {
  if (fs.existsSync(p)) {
    winwordExe = p;
    break;
  }
}

if (winwordExe) {
  console.log(`Đang mở qua: ${winwordExe}`);
  const child = cp.spawn(winwordExe, [docxPath], {
    detached: true,
    stdio: 'ignore'
  });
  child.unref();
} else {
  console.log('Đang mở qua lệnh Windows Start mặc định...');
  const child = cp.spawn('cmd.exe', ['/c', 'start', '""', docxPath], {
    detached: true,
    stdio: 'ignore'
  });
  child.unref();
}

console.log('🚀 Đã gửi lệnh mở Microsoft Word thành công!');
console.log('--------------------------------------------------');
