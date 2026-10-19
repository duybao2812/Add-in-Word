/**
 * scripts/detect-machine.js
 * Tự động phát hiện, thống kê và nhận diện phần cứng máy tính khi dự án được mở hoặc di chuyển qua máy khác.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const PROFILE_FILE = path.join(PROJECT_ROOT, 'hardware-profiles.json');
const MARKDOWN_FILE = path.join(PROJECT_ROOT, 'HARDWARE_SPECS.md');
const POWERSHELL_SCRIPT = path.join(__dirname, 'get-hardware-info.ps1');

/**
 * Thu thập thông tin phần cứng chi tiết của máy hiện tại
 */
function gatherCurrentHardware() {
  // Thử dùng PowerShell script để lấy thông số phần cứng cấp sâu trên Windows
  if (process.platform === 'win32') {
    try {
      const psCommand = `powershell -NoProfile -ExecutionPolicy Bypass -File "${POWERSHELL_SCRIPT}"`;
      const output = execSync(psCommand, { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] });
      if (output && output.trim().startsWith('{')) {
        const parsed = JSON.parse(output.trim());
        parsed.NodeVersion = process.version;
        parsed.Platform = process.platform;
        return parsed;
      }
    } catch (e) {
      // Fallback nếu PowerShell gặp lỗi
    }
  }

  // Fallback thông qua Node.js core os module
  const cpus = os.cpus() || [];
  const totalMemGB = Math.round((os.totalmem() / (1024 * 1024 * 1024)) * 100) / 100;

  return {
    MachineId: `${os.hostname()}_${os.userInfo().username}`,
    ComputerName: os.hostname(),
    UserName: os.userInfo().username,
    SystemUUID: '',
    Timestamp: new Date().toISOString().replace('T', ' ').substring(0, 19),
    OperatingSystem: {
      Caption: `${os.type()} ${os.release()}`,
      Version: os.release(),
      BuildNumber: '',
      Architecture: os.arch()
    },
    Motherboard: {
      Manufacturer: 'Generic',
      Product: 'Motherboard',
      SerialNumber: ''
    },
    BIOS: {
      Manufacturer: '',
      Version: '',
      SerialNumber: ''
    },
    Processor: {
      Name: cpus.length > 0 ? cpus[0].model.trim() : 'Unknown CPU',
      Manufacturer: '',
      Architecture: os.arch(),
      NumberOfCores: cpus.length,
      NumberOfLogicalProcessors: cpus.length,
      MaxClockSpeedMHz: cpus.length > 0 ? cpus[0].speed : 0,
      SocketDesignation: ''
    },
    Memory: {
      TotalPhysicalRAM_GB: totalMemGB,
      SlotCount: 1,
      Modules: []
    },
    Graphics: [],
    PhysicalDisks: [],
    LogicalVolumes: [],
    NetworkAdapters: [],
    NodeVersion: process.version,
    Platform: process.platform
  };
}

/**
 * Đọc dữ liệu hồ sơ phần cứng hiện có
 */
function loadProfiles() {
  if (fs.existsSync(PROFILE_FILE)) {
    try {
      const data = fs.readFileSync(PROFILE_FILE, 'utf8');
      return JSON.parse(data);
    } catch (e) {
      console.warn('⚠️ Không thể đọc file hardware-profiles.json cũ, đang tạo mới...');
    }
  }
  return {
    $schemaVersion: '1.0',
    description: 'Danh sách hồ sơ phần cứng các máy tính từng chạy dự án Add-in Word',
    lastActiveMachineId: null,
    lastDetectedAt: null,
    machines: {}
  };
}

/**
 * Lưu dữ liệu hồ sơ
 */
function saveProfiles(profiles) {
  fs.writeFileSync(PROFILE_FILE, JSON.stringify(profiles, null, 2), 'utf8');
}

/**
 * Sinh file Markdown tóm tắt trực quan
 */
function generateMarkdownSummary(profiles, currentMachine, isNewMachine, previousMachine) {
  const c = currentMachine;
  const osInfo = c.OperatingSystem || {};
  const cpu = c.Processor || {};
  const mem = c.Memory || {};
  const gpus = c.Graphics || [];
  const disks = c.PhysicalDisks || [];
  const volumes = c.LogicalVolumes || [];
  const nets = c.NetworkAdapters || [];

  let md = `# 🖥️ Thống Kê Thông Số Phần Cứng & Nhận Diện Máy Tính\n\n`;
  md += `> **Thời gian cập nhật:** \`${c.Timestamp}\`  \n`;
  md += `> **Máy tính hiện tại:** **\`${c.ComputerName}\`** (${c.Motherboard?.Product || 'PC'})  \n`;
  md += `> **Trạng thái:** ${isNewMachine ? '✨ **MÁY MỚI PHÁT HIỆN!**' : '✅ **ĐÃ NHẬN DIỆN CHÍNH XÁC**'}  \n\n`;

  if (isNewMachine && previousMachine) {
    md += `## 🔄 Nhật Ký Di Chuyển Máy (Chuyển đổi thiết bị)\n\n`;
    md += `Dự án vừa được chuyển từ **\`${previousMachine.ComputerName}\`** sang **\`${c.ComputerName}\`**.\n\n`;
    md += `| Thông số | Máy trước đó (\`${previousMachine.ComputerName}\`) | Máy hiện tại (\`${c.ComputerName}\`) |\n`;
    md += `| :--- | :--- | :--- |\n`;
    md += `| **CPU** | ${previousMachine.Processor?.Name || 'N/A'} | **${cpu.Name || 'N/A'}** |\n`;
    md += `| **RAM** | ${previousMachine.Memory?.TotalPhysicalRAM_GB || 'N/A'} GB | **${mem.TotalPhysicalRAM_GB || 'N/A'} GB** |\n`;
    md += `| **GPU** | ${(previousMachine.Graphics || []).map(g => g.Name).join(', ') || 'N/A'} | **${gpus.map(g => g.Name).join(', ') || 'N/A'}** |\n`;
    md += `| **Hệ Điều Hành** | ${previousMachine.OperatingSystem?.Caption || 'N/A'} | **${osInfo.Caption || 'N/A'}** |\n\n`;
  }

  md += `## 1. Thông Số Tổng Quan\n\n`;
  md += `| Thông Tin | Chi Tiết |\n`;
  md += `| :--- | :--- |\n`;
  md += `| **Computer Name** | \`${c.ComputerName}\` |\n`;
  md += `| **User Name** | \`${c.UserName}\` |\n`;
  md += `| **Mã Định Danh (UUID)** | \`${c.SystemUUID || c.MachineId}\` |\n`;
  md += `| **Mainboard / Model** | ${c.Motherboard?.Manufacturer || ''} ${c.Motherboard?.Product || ''} *(S/N: ${c.Motherboard?.SerialNumber || 'N/A'})* |\n`;
  md += `| **BIOS Version** | ${c.BIOS?.Manufacturer || ''} ${c.BIOS?.Version || ''} |\n`;
  md += `| **Hệ Điều Hành** | ${osInfo.Caption || ''} (${osInfo.Architecture || ''}) - Build ${osInfo.BuildNumber || ''} |\n`;
  md += `| **Môi trường Node.js** | ${c.NodeVersion || process.version} (${process.platform}) |\n\n`;

  md += `## 2. Bộ Xử Lý (CPU)\n\n`;
  md += `- **Model CPU:** **${cpu.Name || 'Unknown'}**\n`;
  md += `- **Số nhân vật lý (Cores):** \`${cpu.NumberOfCores || 'N/A'}\` nhân\n`;
  md += `- **Số luồng xử lý (Threads):** \`${cpu.NumberOfLogicalProcessors || 'N/A'}\` luồng\n`;
  md += `- **Xung nhịp tối đa:** \`${cpu.MaxClockSpeedMHz || 'N/A'} MHz\`\n`;
  md += `- **Socket:** \`${cpu.SocketDesignation || 'N/A'}\`\n\n`;

  md += `## 3. Bộ Nhớ Trong (RAM)\n\n`;
  md += `- **Tổng dung lượng RAM:** **\`${mem.TotalPhysicalRAM_GB || 0} GB\`**\n`;
  md += `- **Số khe cắm đang dùng:** \`${mem.SlotCount || (mem.Modules ? mem.Modules.length : 0)}\`\n\n`;

  if (mem.Modules && mem.Modules.length > 0) {
    md += `### Chi tiết từng thanh RAM:\n`;
    md += `| Vị trí (Slot) | Dung lượng | Tốc độ (Bus) | Nhà sản xuất | Mã linh kiện (Part Number) |\n`;
    md += `| :--- | :--- | :--- | :--- | :--- |\n`;
    mem.Modules.forEach((m, idx) => {
      md += `| Khe ${idx + 1} (${m.DeviceLocator || m.BankLabel}) | **${m.CapacityGB} GB** | ${m.ConfiguredClockSpeedMHz || m.SpeedMHz} MHz | ${m.Manufacturer} | \`${m.PartNumber}\` |\n`;
    });
    md += `\n`;
  }

  md += `## 4. Card Đồ Họa (GPU & Hiển thị)\n\n`;
  if (gpus.length > 0) {
    md += `| Tên Card Đồ Họa | Bộ nhớ đồ họa (VRAM) | Driver Version | Độ phân giải hiển thị |\n`;
    md += `| :--- | :--- | :--- | :--- |\n`;
    gpus.forEach(g => {
      md += `| **${g.Name}** | ${g.AdapterRAM_GB ? g.AdapterRAM_GB + ' GB' : 'Onboard/Shared'} | \`${g.DriverVersion}\` | ${g.CurrentResolution} |\n`;
    });
    md += `\n`;
  } else {
    md += `*Không phát hiện card đồ họa rời.*\n\n`;
  }

  md += `## 5. Ổ Cứng & Không Gian Lưu Trữ\n\n`;
  if (disks.length > 0) {
    md += `### Ổ cứng vật lý:\n`;
    md += `| Model Ổ Cứng | Chuẩn kết nối | Dung lượng | Serial Number |\n`;
    md += `| :--- | :--- | :--- | :--- |\n`;
    disks.forEach(d => {
      md += `| **${d.Model}** | ${d.InterfaceType} | **${d.SizeGB} GB** | \`${d.SerialNumber}\` |\n`;
    });
    md += `\n`;
  }

  if (volumes.length > 0) {
    md += `### Phân vùng ổ đĩa (Logical Drives):\n`;
    md += `| Phân vùng | Tên nhãn | Định dạng | Tổng dung lượng | Còn trống | Tỷ lệ trống |\n`;
    md += `| :--- | :--- | :--- | :--- | :--- | :--- |\n`;
    volumes.forEach(v => {
      md += `| **${v.DriveLetter}** | ${v.VolumeName || '(Không tên)'} | ${v.FileSystem} | ${v.TotalSizeGB} GB | **${v.FreeSpaceGB} GB** | ${v.FreePercent}% |\n`;
    });
    md += `\n`;
  }

  md += `## 6. Kết Nối Mạng (Network Interfaces)\n\n`;
  if (nets.length > 0) {
    md += `| Card mạng | Địa chỉ MAC | Địa chỉ IP |\n`;
    md += `| :--- | :--- | :--- |\n`;
    nets.forEach(n => {
      md += `| ${n.Description} | \`${n.MACAddress || 'N/A'}\` | \`${n.IPAddress}\` |\n`;
    });
    md += `\n`;
  }

  md += `## 7. Danh Sách Các Máy Từng Chạy Dự Án Này\n\n`;
  md += `| Tên máy | ID Hồ Sơ | CPU | RAM | Lần đầu ghi nhận | Lần cuối hoạt động |\n`;
  md += `| :--- | :--- | :--- | :--- | :--- | :--- |\n`;
  Object.keys(profiles.machines).forEach(mid => {
    const m = profiles.machines[mid];
    const isCurrent = (mid === c.MachineId || m.SystemUUID === c.SystemUUID);
    const badge = isCurrent ? ' **(Máy này)**' : '';
    md += `| **${m.ComputerName}**${badge} | \`${mid.substring(0, 16)}...\` | ${m.Processor?.Name || 'N/A'} | ${m.Memory?.TotalPhysicalRAM_GB || 'N/A'} GB | ${m.firstSeen || m.Timestamp} | ${m.lastSeen || m.Timestamp} |\n`;
  });
  md += `\n---\n*File được tạo và cập nhật tự động bởi hệ thống nhận diện cấu hình máy tính Add-in Word.*\n`;

  fs.writeFileSync(MARKDOWN_FILE, md, 'utf8');
}

/**
 * Hàm kiểm tra và nhận diện máy
 */
function detectAndSyncMachine() {
  console.log('🔍 Đang kiểm tra và quét thông số phần cứng máy tính hiện tại...');
  const currentHardware = gatherCurrentHardware();
  const profiles = loadProfiles();

  // Xác định xem máy hiện tại đã có trong danh sách chưa
  let matchedKey = null;

  // Ưu tiên so khớp: 1) SystemUUID, 2) MachineId, 3) Motherboard Serial, 4) ComputerName + Primary MAC
  for (const [key, profile] of Object.entries(profiles.machines)) {
    if (currentHardware.SystemUUID && profile.SystemUUID && currentHardware.SystemUUID === profile.SystemUUID) {
      matchedKey = key;
      break;
    }
    if (key === currentHardware.MachineId) {
      matchedKey = key;
      break;
    }
    if (currentHardware.Motherboard?.SerialNumber && profile.Motherboard?.SerialNumber &&
        currentHardware.Motherboard.SerialNumber !== 'Unknown' &&
        currentHardware.Motherboard.SerialNumber === profile.Motherboard.SerialNumber) {
      matchedKey = key;
      break;
    }
    if (profile.ComputerName === currentHardware.ComputerName) {
      matchedKey = key;
      break;
    }
  }

  const isNewMachine = !matchedKey;
  const previousMachine = profiles.lastActiveMachineId ? profiles.machines[profiles.lastActiveMachineId] : null;
  const nowStr = new Date().toISOString().replace('T', ' ').substring(0, 19);

  const machineKey = matchedKey || currentHardware.MachineId;

  if (isNewMachine) {
    console.log('\n================================================================');
    console.log('🚀 [PHÁT HIỆN CHUYỂN SANG MÁY TÍNH MỚI] 🚀');
    console.log(`➡️  Tên máy mới: ${currentHardware.ComputerName}`);
    console.log(`➡️  CPU: ${currentHardware.Processor?.Name}`);
    console.log(`➡️  RAM: ${currentHardware.Memory?.TotalPhysicalRAM_GB} GB`);
    if (previousMachine) {
      console.log(`ℹ️  Máy trước đó: ${previousMachine.ComputerName} (${previousMachine.Processor?.Name} - ${previousMachine.Memory?.TotalPhysicalRAM_GB} GB RAM)`);
    }
    console.log('💾 Đang tự động lưu cấu hình máy mới vào "hardware-profiles.json"...');
    console.log('================================================================\n');

    currentHardware.firstSeen = nowStr;
    currentHardware.lastSeen = nowStr;
    profiles.machines[machineKey] = currentHardware;
  } else {
    console.log('\n================================================================');
    console.log(`✅ [ĐÃ NHẬN DIỆN MÁY TÍNH QUEN THUỘC]`);
    console.log(`💻 Tên máy: ${currentHardware.ComputerName} (${currentHardware.Motherboard?.Product || 'PC'})`);
    console.log(`⚙️  CPU: ${currentHardware.Processor?.Name} (${currentHardware.Processor?.NumberOfCores} Cores, ${currentHardware.Processor?.NumberOfLogicalProcessors} Threads)`);
    console.log(`🧠 RAM: ${currentHardware.Memory?.TotalPhysicalRAM_GB} GB`);
    if (currentHardware.Graphics && currentHardware.Graphics.length > 0) {
      console.log(`🎮 GPU: ${currentHardware.Graphics.map(g => g.Name).join(', ')}`);
    }
    console.log(`🕒 Lần nhận diện trước: ${profiles.machines[machineKey].lastSeen || 'N/A'}`);
    console.log('================================================================\n');

    // Cập nhật thông số mới nhất (ví dụ dung lượng ổ cứng trống, ngày giờ)
    const originalFirstSeen = profiles.machines[machineKey].firstSeen || profiles.machines[machineKey].Timestamp;
    profiles.machines[machineKey] = {
      ...currentHardware,
      firstSeen: originalFirstSeen,
      lastSeen: nowStr
    };
  }

  profiles.lastActiveMachineId = machineKey;
  profiles.lastDetectedAt = nowStr;

  // Lưu file JSON
  saveProfiles(profiles);

  // Sinh file Markdown
  generateMarkdownSummary(profiles, profiles.machines[machineKey], isNewMachine, previousMachine);

  console.log(`📁 Đã lưu thông số chi tiết vào file:`);
  console.log(`   - JSON (để code đọc): ${path.relative(PROJECT_ROOT, PROFILE_FILE)}`);
  console.log(`   - Markdown (để xem trực quan): ${path.relative(PROJECT_ROOT, MARKDOWN_FILE)}\n`);

  return profiles.machines[machineKey];
}

// Nếu chạy trực tiếp từ dòng lệnh: node scripts/detect-machine.js
if (require.main === module) {
  detectAndSyncMachine();
}

module.exports = {
  detectAndSyncMachine,
  gatherCurrentHardware,
  loadProfiles
};
