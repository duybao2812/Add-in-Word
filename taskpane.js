/**
 * taskpane.js - Xử lý logic Office.js chuẩn hóa thể thức theo Nghị định 30/2020/NĐ-CP
 * Microsoft Word JavaScript API Taskpane Add-in (WinForms Classic Theme)
 * Hỗ trợ đa dạng thể thức: Công văn hành chính, Hợp đồng kinh tế/thi công, Biên bản làm việc, Tờ trình
 */

/* ==========================================================================
   1. CÁC HÀM TIỆN ÍCH HỆ THỐNG & GIAO DIỆN
   ========================================================================== */

/**
 * Cập nhật chuỗi thông báo trên thanh StatusStrip WinForms ở đáy
 */
function setStatus(message, isHost = false) {
  const targetId = isHost ? "hostInfo" : "statusMessage";
  const el = document.getElementById(targetId);
  if (el) {
    el.textContent = message;
    el.setAttribute("title", message);
  }
  console.log(`[StatusStrip] ${message}`);
}

/**
 * Hiển thị thông báo an toàn không ném exception trong môi trường Office WebView2
 */
function safeAlert(msg) {
  try {
    if (typeof window !== "undefined" && typeof window.alert === "function") {
      window.alert(msg);
      return;
    }
  } catch (e) {
    console.warn("window.alert không được hỗ trợ trong môi trường này:", e);
  }
  setStatus(typeof msg === "string" ? msg.split("\n")[0] : String(msg));
}

// Thay thế window.alert an toàn cho toàn bộ Taskpane
window.alert = safeAlert;

// Biến chống gọi trùng lặp (debounce/dedup) các lệnh từ dialog
let lastProcessedActionId = "";
let lastProcessedActionTime = 0;

/**
 * Loại bỏ dấu tiếng Việt để tạo ký hiệu văn bản Latin chuẩn
 */
function removeVietnameseTones(str) {
  if (!str) return "";
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
}

/**
 * Tạo chuỗi ngày tháng năm chuẩn NĐ 30 theo ngày hiện tại
 */
function generateCurrentDateString(locationName = "TP. Hồ Chí Minh") {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();
  const loc = (locationName || "").trim() || "TP. Hồ Chí Minh";
  return `${loc}, ngày ${day} tháng ${month} năm ${year}`;
}

/**
 * Kiểm tra xem chuỗi có phải là chuỗi mẫu giả định / placeholder không
 */
function isPlaceholderText(text) {
  if (!text) return true;
  const t = text.trim();
  return (
    /^(?:TÊN\s*CƠ\s*QUAN|Số\s*[:\.]\s*\.\.\.|\.\.\.|Địa\s*danh.*ngày)/i.test(t) ||
    /\.\.\./.test(t) ||
    /ngày\s+\.\.\.\s+tháng/i.test(t) ||
    /ngày\s+.*tháng\s+.*năm\s+20\.\./i.test(t) ||
    /TÊN\s*VIẾT\s*TẮT/i.test(t) ||
    /TÊN\s*CƠ\s*QUAN\s*CHỦ\s*QUẢN/i.test(t) ||
    /TÊN\s*CƠ\s*QUAN\s*BAN\s*HÀNH/i.test(t) ||
    /TÊN\s*CƠ\s*QUAN,\s*ĐƠN\s*VỊ/i.test(t)
  );
}

/**
 * Trích xuất viết tắt thông minh từ tên cơ quan / doanh nghiệp:
 * - Bỏ tiền tố doanh nghiệp: Công ty, Cty, TNHH, MTV, Cổ phần, CP, Tập đoàn, Tổng công ty, Doanh nghiệp...
 * - Bỏ ngành nghề chung nếu có tên riêng phía sau: Thương mại dịch vụ (TMDV), Xây dựng, Đầu tư, Sản xuất, Tư vấn...
 * - Bỏ liên từ: và, &, của, thuộc
 * - Loại bỏ dấu tiếng Việt để thành ký hiệu chuẩn Latin (Á -> A, Đ -> D...)
 */
function getSmartAgencyAcronym(agencyName) {
  if (!agencyName) return "CQ";
  let name = agencyName.trim();
  if (!name || isPlaceholderText(name)) return "CQ";

  const isEnterprise =
    /^(?:Công\s*ty|Cty|Tập\s*đoàn|Tổng\s*công\s*ty|TCT|Doanh\s*nghiệp|DNTN|Hợp\s*tác\s*xã|HTX)/i.test(name) ||
    /(?:TNHH|Cổ\s*phần|CP|MTV|Một\s*thành\s*viên)/i.test(name);

  if (isEnterprise) {
    name = name.replace(
      /^(?:Công\s*ty(?:\s*cổ\s*phần|\s*TNHH)?|Cty|Tập\s*đoàn|Tổng\s*công\s*ty|TCT|Doanh\s*nghiệp(?:\s*tư\s*nhân)?|DNTN|Hợp\s*tác\s*xã|HTX)\s+/gi,
      ""
    );
    name = name.replace(
      /^(?:TNHH\s+MTV|TNHH\s+Một\s+thành\s+viên|TNHH\s+1\s*TV|TNHH|MTV|Cổ\s*phần|CP)\s+/gi,
      ""
    );

    const industryKeywords = [
      "Thương\\s*mại\\s*(?:và|&|\\/)?\\s*Dịch\\s*vụ",
      "TMDV",
      "Dịch\\s*vụ\\s*(?:và|&|\\/)?\\s*Thương\\s*mại",
      "Dịch\\s*vụ",
      "Xây\\s*dựng\\s*(?:và|&|\\/)?\\s*Thương\\s*mại",
      "Xây\\s*dựng",
      "XD",
      "Đầu\\s*tư\\s*(?:và|&|\\/)?\\s*Phát\\s*triển",
      "Đầu\\s*tư",
      "Sản\\s*xuất\\s*(?:và|&|\\/)?\\s*Thương\\s*mại",
      "Sản\\s*xuất",
      "SX",
      "Tư\\s*vấn\\s*(?:và|&|\\/)?\\s*Thiết\\s*kế",
      "Tư\\s*vấn",
      "Thiết\\s*kế",
      "Kỹ\\s*thuật\\s*(?:và|&|\\/)?\\s*Công\\s*nghệ",
      "Kỹ\\s*thuật",
      "Công\\s*nghệ\\s*Thông\\s*tin",
      "Công\\s*nghệ",
      "CNTT",
      "Phần\\s*mềm",
      "Truyền\\s*thông\\s*(?:và|&|\\/)?\\s*Giải\\s*trí",
      "Truyền\\s*thông",
      "Quảng\\s*cáo",
      "Bất\\s*động\\s*sản",
      "Địa\\s*ốc",
      "Vận\\s*tải\\s*(?:và|&|\\/)?\\s*Du\\s*lịch",
      "Vận\\s*tải",
      "Logistics",
      "Xuất\\s*nhập\\s*khẩu",
      "XNK",
      "Dược\\s*phẩm",
      "Y\\s*tế",
      "Nông\\s*nghiệp",
      "Môi\\s*trường"
    ];

    let loop = true;
    while (loop) {
      loop = false;
      name = name.replace(/^(?:và|&|\-)\s+/i, "").trim();
      for (const kw of industryKeywords) {
        const reg = new RegExp("^" + kw + "\\s+", "i");
        if (reg.test(name)) {
          const remainder = name.replace(reg, "").trim();
          if (remainder.length > 0) {
            name = remainder;
            loop = true;
            break;
          }
        }
      }
    }
  }

  const clean = name.replace(/[\(\)\[\]\{\}\.,:;\/\\_\-–—]/g, " ").trim();
  const rawWords = clean.split(/\s+/).filter(w => w.length > 0);
  const words = rawWords.filter(w => !/^(?:và|&|của|thuộc)$/i.test(w));
  const effectiveWords = words.length > 0 ? words : rawWords;
  if (effectiveWords.length === 0) return "CQ";

  let acronym = "";
  for (const w of effectiveWords) {
    acronym += w[0].toUpperCase();
  }
  return removeVietnameseTones(acronym) || "CQ";
}

/**
 * Định dạng số thứ tự tự động thêm số 0 phía trước:
 * 1 số -> thêm 2 số 0 (1 -> 001, 9 -> 009)
 * 2 số -> thêm 1 số 0 (10 -> 010, 99 -> 099)
 * 3 số trở lên -> giữ nguyên (100 -> 100, 125 -> 125)
 */
function formatSequenceNumber(rawVal) {
  if (!rawVal) return "001";
  const num = parseInt(String(rawVal).replace(/\D/g, ""), 10);
  if (isNaN(num) || num <= 0) return "001";
  if (num < 10) return `00${num}`;
  if (num < 100) return `0${num}`;
  return String(num);
}

/**
 * Lấy mã ký hiệu viết tắt loại hợp đồng
 */
function getContractCode(contractType) {
  switch (contractType) {
    case "hd-thi-cong": return "HĐTC";
    case "hd-kinh-te": return "HĐKT";
    case "hd-nguyen-tac": return "HĐNT";
    case "hd-dich-vu": return "HĐDV";
    case "hd-mua-ban": return "HĐMB";
    case "hd-tu-van": return "HĐTV";
    default: return "HĐ";
  }
}

/**
 * Cập nhật tự động trường Ký hiệu văn bản theo từng Loại văn bản
 */
function updateDocSymbolString() {
  const txtPub = document.getElementById("txtPublishingOrg");
  const cboType = document.getElementById("cboDocType");
  const txtSymbol = document.getElementById("txtDocSymbol");

  if (!txtSymbol) return;

  const docType = cboType ? cboType.value : "cong-van";
  const agency = txtPub ? txtPub.value.trim() : "";
  const acronym = getSmartAgencyAcronym(agency);
  const now = new Date();
  const year = now.getFullYear();

  if (docType === "hop-dong") {
    const cboContract = document.getElementById("cboContractType");
    const contractCode = getContractCode(cboContract ? cboContract.value : "hd-thi-cong");
    txtSymbol.value = `/${year}/${contractCode}-${acronym}`;
  } else if (docType === "bien-ban") {
    const cboMinute = document.getElementById("cboMinuteType");
    const minCode = (cboMinute && cboMinute.value === "bb-nghiem-thu") ? "BBNT" : (cboMinute && cboMinute.value === "bb-ban-giao") ? "BBBG" : "BB";
    txtSymbol.value = `/${minCode}-${acronym}`;
  } else if (docType === "to-trinh") {
    const cboReport = document.getElementById("cboReportType");
    const repCode = (cboReport && cboReport.value.startsWith("bc-")) ? "BC" : "TTr";
    txtSymbol.value = `/${repCode}-${acronym}`;
  } else {
    txtSymbol.value = `/CV-${acronym}`;
  }
}

function updateDocNumberString() {
  updateDocSymbolString();
}

/**
 * Cập nhật chuỗi ngày tháng năm chuẩn NĐ 30 kết hợp Địa danh và Lịch chọn ngày
 */
function updateFullDateString(optDay, optMonth, optYear) {
  const txtLocation = document.getElementById("txtLocation");
  const txtDocDate = document.getElementById("txtDocDate");
  const dtpDocDate = document.getElementById("dtpDocDate");

  let d, m, y;
  if (optDay && optMonth && optYear) {
    d = String(optDay).padStart(2, "0");
    m = String(optMonth).padStart(2, "0");
    y = String(optYear);
  } else if (dtpDocDate && dtpDocDate.value) {
    const parts = dtpDocDate.value.split("-");
    if (parts.length === 3) {
      y = parts[0];
      m = parts[1];
      d = parts[2];
    }
  }

  if (!d || !m || !y) {
    const now = new Date();
    d = String(now.getDate()).padStart(2, "0");
    m = String(now.getMonth() + 1).padStart(2, "0");
    y = String(now.getFullYear());
  }

  const loc = (txtLocation && txtLocation.value ? txtLocation.value.trim() : "TP. Hồ Chí Minh");
  if (txtDocDate) {
    txtDocDate.value = `${loc}, ngày ${d} tháng ${m} năm ${y}`;
  }
}

/**
 * Xử lý khi người dùng nhập tay ngày tháng theo dạng dd/mm/yyyy
 */
function handleDateManualInput(e) {
  const val = e.target.value.trim();
  const match = val.match(/(?:^|[^\d])(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  if (match) {
    const d = String(match[1]).padStart(2, "0");
    const m = String(match[2]).padStart(2, "0");
    const y = match[3];
    const dtp = document.getElementById("dtpDocDate");
    if (dtp) {
      dtp.value = `${y}-${m}-${d}`;
    }
    const loc = (document.getElementById("txtLocation").value || "TP. Hồ Chí Minh").trim();
    e.target.value = `${loc}, ngày ${d} tháng ${m} năm ${y}`;
    setStatus(`Đã chuyển đổi ngày: ngày ${d} tháng ${m} năm ${y}`);
  }
}

/**
 * Chuẩn hóa chuỗi trích yếu của Công văn luôn có tiền tố "Về việc: "
 */
function normalizeVeViec(rawText) {
  if (!rawText) return "Về việc: tham gia hội nghị";
  let t = rawText.trim();
  t = t.replace(/^(?:Về\s*việc|V\/v|Vv)\s*[:\.]?\s*/i, "");
  return `Về việc: ${t || "phối hợp triển khai công tác"}`;
}

/* ==========================================================================
   2. THUẬT TOÁN THIÊN VĂN LỊCH ÂM VIỆT NAM (UTC+7 HỒ NGỌC ĐỨC) & ĐỒNG HỒ 24H
   ========================================================================== */

function jdFromDate(dd, mm, yy) {
  let a = Math.floor((14 - mm) / 12);
  let y = yy + 4800 - a;
  let m = mm + 12 * a - 3;
  let jd = dd + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) - 32045;
  if (jd < 2299161) {
    jd = dd + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - 32083;
  }
  return jd;
}

function getNewMoonDay(k, timeZone = 7) {
  let T = k / 1236.85;
  let T2 = T * T;
  let T3 = T2 * T;
  let dr = Math.PI / 180;
  let Jd1 = 2415020.75933 + 29.53058868 * k + 0.0001178 * T2 - 0.000000155 * T3;
  Jd1 += 0.00033 * Math.sin((166.56 + 132.87 * T - 0.009173 * T2) * dr);
  let M = 359.2242 + 29.10535608 * k - 0.0000333 * T2 - 0.00000347 * T3;
  let Mpr = 306.0253 + 385.81691806 * k + 0.0107306 * T2 + 0.00001236 * T3;
  let F = 21.2964 + 390.67050646 * k - 0.0016528 * T2 - 0.00000239 * T3;
  let C1 = (0.1734 - 0.000393 * T) * Math.sin(M * dr) + 0.0021 * Math.sin(2 * dr * M);
  C1 -= 0.4068 * Math.sin(Mpr * dr) + 0.0161 * Math.sin(2 * dr * Mpr);
  C1 -= 0.0004 * Math.sin(3 * dr * Mpr);
  C1 += 0.0104 * Math.sin(2 * dr * F) - 0.0051 * Math.sin((M + Mpr) * dr);
  C1 -= 0.0074 * Math.sin((M - Mpr) * dr) + 0.0004 * Math.sin((2 * F + M) * dr);
  C1 -= 0.0004 * Math.sin((2 * F - M) * dr) - 0.0006 * Math.sin((2 * F + Mpr) * dr);
  C1 += 0.001 * Math.sin((2 * F - Mpr) * dr) + 0.0005 * Math.sin((2 * Mpr + M) * dr);
  let deltat = (T < -11) ? (0.001 + 0.000839 * T + 0.0002261 * T2 - 0.00000845 * T3 - 0.000000081 * T * T3) : (-0.00005 + 0.000297 * T + 0.000045 * T2);
  let JdNew = Jd1 + C1 - deltat;
  return Math.floor(JdNew + 0.5 + timeZone / 24);
}

function getSunLongitude(jdn, timeZone = 7) {
  let T = (jdn - 2451545.0 + 0.5 - timeZone / 24) / 36525;
  let T2 = T * T;
  let dr = Math.PI / 180;
  let L0 = 280.46645 + 36000.76983 * T + 0.0003032 * T2;
  let M = 357.5291 + 35999.05029 * T - 0.0001536 * T2;
  let C = (1.9146 - 0.004817 * T - 0.000014 * T2) * Math.sin(M * dr);
  C += (0.019993 - 0.000101 * T) * Math.sin(2 * M * dr) + 0.00029 * Math.sin(3 * M * dr);
  let theta = L0 + C;
  theta = theta - 360 * Math.floor(theta / 360);
  return Math.floor(theta / 30);
}

function getLunarMonth11(yy, timeZone = 7) {
  let off = jdFromDate(31, 12, yy) - 2415021;
  let k = Math.floor(off / 29.530588853);
  let nm = getNewMoonDay(k, timeZone);
  let sunLong = getSunLongitude(nm, timeZone);
  if (sunLong >= 9) {
    nm = getNewMoonDay(k - 1, timeZone);
  }
  return nm;
}

function getLeapMonthOffset(a11, timeZone = 7) {
  let k = Math.floor((a11 - 2415021.076998695) / 29.530588853 + 0.5);
  let last = 0;
  let i = 1;
  let arc = getSunLongitude(getNewMoonDay(k + i, timeZone), timeZone);
  do {
    last = arc;
    i++;
    arc = getSunLongitude(getNewMoonDay(k + i, timeZone), timeZone);
  } while (arc !== last && i < 14);
  return i - 1;
}

function getLunarDate(dd, mm, yy, timeZone = 7) {
  let dayNumber = jdFromDate(dd, mm, yy);
  let k = Math.floor((dayNumber - 2415021.076998695) / 29.530588853);
  let monthStart = getNewMoonDay(k + 1, timeZone);
  if (monthStart > dayNumber) {
    monthStart = getNewMoonDay(k, timeZone);
  }
  let a11 = getLunarMonth11(yy, timeZone);
  let b11 = a11;
  let lunarYear;
  if (a11 >= monthStart) {
    lunarYear = yy;
    a11 = getLunarMonth11(yy - 1, timeZone);
  } else {
    lunarYear = yy + 1;
    b11 = getLunarMonth11(yy + 1, timeZone);
  }
  let lunarDay = dayNumber - monthStart + 1;
  let diff = Math.floor((monthStart - a11) / 29);
  let lunarLeap = 0;
  let lunarMonth = diff + 11;
  if (b11 - a11 > 365) {
    let leapMonthDiff = getLeapMonthOffset(a11, timeZone);
    if (diff >= leapMonthDiff) {
      lunarMonth = diff + 10;
      if (diff === leapMonthDiff) {
        lunarLeap = 1;
      }
    }
  }
  if (lunarMonth > 12) {
    lunarMonth = lunarMonth - 12;
  }
  if (lunarMonth >= 11 && diff < 4) {
    lunarYear -= 1;
  }

  const CAN = ['Giáp', 'Ất', 'Bính', 'Đinh', 'Mậu', 'Kỷ', 'Canh', 'Tân', 'Nhâm', 'Quý'];
  const CHI = ['Tý', 'Sửu', 'Dần', 'Mão', 'Thìn', 'Tỵ', 'Ngọ', 'Mùi', 'Thân', 'Dậu', 'Tuất', 'Hợi'];

  let yearCanChi = CAN[(lunarYear + 6) % 10] + ' ' + CHI[(lunarYear + 8) % 12];
  let dayCanChi = CAN[(dayNumber + 9) % 10] + ' ' + CHI[(dayNumber + 1) % 12];

  return {
    day: lunarDay,
    month: lunarMonth,
    year: lunarYear,
    leap: lunarLeap,
    yearCanChi,
    dayCanChi
  };
}

/**
 * Cập nhật hiển thị đồng hồ 24h và lịch âm dương thời gian thực
 */
function updateClockAndLunar() {
  const now = new Date();
  const h = String(now.getHours()).padStart(2, '0');
  const m = String(now.getMinutes()).padStart(2, '0');
  const s = String(now.getSeconds()).padStart(2, '0');

  const elClock = document.getElementById('clock24h');
  if (elClock) elClock.textContent = `${h}:${m}:${s}`;

  const daysOfWeek = ['Chủ Nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
  const dayName = daysOfWeek[now.getDay()];
  const dd = now.getDate();
  const mm = now.getMonth() + 1;
  const yyyy = now.getFullYear();

  const elDate = document.getElementById('clockDate');
  if (elDate) {
    elDate.textContent = `${dayName}, ngày ${String(dd).padStart(2, '0')} tháng ${String(mm).padStart(2, '0')} năm ${yyyy}`;
  }

  const elLunar = document.getElementById('lblLunarText');
  if (elLunar) {
    try {
      const lunar = getLunarDate(dd, mm, yyyy, 7);
      const lDay = String(lunar.day).padStart(2, '0');
      const lMonth = String(lunar.month).padStart(2, '0');
      const leapStr = lunar.leap ? ' (Nhuận)' : '';
      elLunar.textContent = `Âm lịch: Ngày ${lDay} tháng ${lMonth}${leapStr} năm ${lunar.yearCanChi} (${lunar.dayCanChi})`;
    } catch (err) {
      elLunar.textContent = `Âm lịch: Ngày ${dd} tháng ${mm}`;
    }
  }
}

/**
 * Điều hướng giữa Menu chính và Form Quy chuẩn NĐ 30
 * BẢO ĐẢM 100%: Dữ liệu form không hề bị xóa hay nạp lại vì toàn bộ DOM vẫn nguyên vẹn
 */
function switchToView(viewName) {
  const viewMainMenu = document.getElementById('viewMainMenu');
  const viewStandardize = document.getElementById('viewStandardize');
  if (!viewMainMenu || !viewStandardize) return;

  if (viewName === 'standardize') {
    viewMainMenu.classList.add('hidden');
    viewStandardize.classList.remove('hidden');
    setStatus('Đang ở Form Quy chuẩn Thể thức NĐ 30/2020/NĐ-CP | Nhập liệu & Chuẩn hóa');
  } else {
    viewStandardize.classList.add('hidden');
    viewMainMenu.classList.remove('hidden');
    setStatus('Đang ở Menu chính | Giữ nguyên 100% dữ liệu đã nhập');
  }
}

/**
 * Hộp thoại trợ giúp thông tin quy chuẩn thể thức NĐ 30/2020/NĐ-CP
 */
function showND30Help() {
  alert(
    'HỆ THỐNG CHUẨN HÓA VĂN BẢN ĐA THỂ THỨC (NĐ 30/2020/NĐ-CP):\n\n' +
    '1. Công văn hành chính: Có Bảng Quốc hiệu & Cơ quan 2 cột không viền, Trích yếu có "Về việc:", Kính gửi thụt lề, Thân bài căn đều hai lề (Justify), Bảng Nơi nhận & Chữ ký 2 cột.\n\n' +
    '2. Hợp đồng kinh tế / Thi công: Quốc hiệu & Tiêu ngữ căn giữa, Tiêu đề Hợp đồng in hoa căn giữa, Căn cứ pháp lý, Thông tin dự án/gói thầu/giá trị, Bên A & Bên B, Điều khoản thực hiện, Bảng Chữ ký 2 bên.\n\n' +
    '3. Biên bản làm việc: Tiêu đề BIÊN BẢN, Thời gian, Địa điểm, Thành phần tham dự, Nội dung diễn biến làm việc, Kết luận, Thư ký & Chủ tọa ký tên.\n\n' +
    '4. Tờ trình / Báo cáo: Tiêu đề TỜ TRÌNH, Trích yếu V/v, Kính gửi cấp phê duyệt, Căn cứ đề xuất, Nội dung kiến nghị, Chữ ký thủ trưởng cơ quan.\n\n' +
    '5. Tính toàn vẹn: Dữ liệu bạn nhập trên Form sẽ được lưu giữ nguyên vẹn 100% khi bạn bấm nút "Quay về Menu" và quay lại.'
  );
}

/**
 * Khởi tạo ứng dụng sau khi toàn bộ DOM và Office.js đã nạp xong
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setStatus("Sẵn sàng | NĐ 30/2020/NĐ-CP Ready");
    setStatus(`Word (${info.platform || "Desktop"})`, true);
    console.log("Office.js đã nạp thành công trên Microsoft Word!");
  } else {
    setStatus("Sẵn sàng (Chế độ Trình duyệt Test UI)");
    setStatus("Browser Preview", true);
  }

  // Khởi động đồng hồ 24h và lịch âm dương
  updateClockAndLunar();
  setInterval(updateClockAndLunar, 1000);

  // Mặc định luôn mở Menu chính đầu tiên theo yêu cầu người dùng
  switchToView("menu");

  // Khởi tạo các sự kiện giao diện
  initUIEvents();
  initPriceDialogEvents();
  initOcrBroadcastReceiver();
});

/**
 * Hiển thị/Ẩn panel tương ứng với Loại văn bản được chọn
 */
function switchDocTypePanels(docType) {
  const pCV = document.getElementById("panelCongVan");
  const pHD = document.getElementById("panelHopDong");
  const pBB = document.getElementById("panelBienBan");
  const pTT = document.getElementById("panelToTrinh");

  if (pCV) pCV.classList.add("hidden");
  if (pHD) pHD.classList.add("hidden");
  if (pBB) pBB.classList.add("hidden");
  if (pTT) pTT.classList.add("hidden");

  switch (docType) {
    case "cong-van":
      if (pCV) pCV.classList.remove("hidden");
      setStatus("Đã chuyển giao diện: Công văn hành chính (CV).");
      break;
    case "hop-dong":
      if (pHD) pHD.classList.remove("hidden");
      setStatus("Đã chuyển giao diện: Hợp đồng kinh tế / Thi công / Dịch vụ (HĐ).");
      break;
    case "bien-ban":
      if (pBB) pBB.classList.remove("hidden");
      setStatus("Đã chuyển giao diện: Biên bản làm việc / Nghiệm thu (BB).");
      break;
    case "to-trinh":
      if (pTT) pTT.classList.remove("hidden");
      setStatus("Đã chuyển giao diện: Tờ trình / Báo cáo (TTr).");
      break;
  }
}

/**
 * Đăng ký tất cả các sự kiện tương tác trên Taskpane
 */
function initUIEvents() {
  const txtLocation = document.getElementById("txtLocation");
  const txtDocDate = document.getElementById("txtDocDate");
  const dtpDocDate = document.getElementById("dtpDocDate");
  const txtPublishingOrg = document.getElementById("txtPublishingOrg");
  const txtDocSeqNumber = document.getElementById("txtDocSeqNumber");
  const cboDocType = document.getElementById("cboDocType");

  // Thiết lập mặc định:
  if (txtLocation && !txtLocation.value) {
    txtLocation.value = "TP. Hồ Chí Minh";
  }

  const now = new Date();
  const nowY = now.getFullYear();
  const nowM = String(now.getMonth() + 1).padStart(2, "0");
  const nowD = String(now.getDate()).padStart(2, "0");

  if (dtpDocDate && !dtpDocDate.value) {
    dtpDocDate.value = `${nowY}-${nowM}-${nowD}`;
  }
  updateFullDateString(nowD, nowM, nowY);

  if (txtDocSeqNumber) {
    if (!txtDocSeqNumber.value) {
      txtDocSeqNumber.value = "001";
    } else {
      txtDocSeqNumber.value = formatSequenceNumber(txtDocSeqNumber.value);
    }
    txtDocSeqNumber.addEventListener("blur", () => {
      txtDocSeqNumber.value = formatSequenceNumber(txtDocSeqNumber.value);
    });
    txtDocSeqNumber.addEventListener("change", () => {
      txtDocSeqNumber.value = formatSequenceNumber(txtDocSeqNumber.value);
    });
  }

  // Khởi tạo hiển thị panel theo loại văn bản ban đầu
  if (cboDocType) {
    switchDocTypePanels(cboDocType.value);
  }
  updateDocSymbolString();

  // Sự kiện thay đổi Loại văn bản chính
  if (cboDocType) {
    cboDocType.addEventListener("change", () => {
      switchDocTypePanels(cboDocType.value);
      updateDocSymbolString();
    });
  }

  // Sự kiện thay đổi Loại hợp đồng phụ
  const cboContractType = document.getElementById("cboContractType");
  const txtContractTitle = document.getElementById("txtContractTitle");
  if (cboContractType) {
    cboContractType.addEventListener("change", () => {
      updateDocSymbolString();
      if (txtContractTitle) {
        switch (cboContractType.value) {
          case "hd-thi-cong":
            txtContractTitle.value = "HỢP ĐỒNG THI CÔNG XÂY DỰNG CÔNG TRÌNH";
            break;
          case "hd-kinh-te":
            txtContractTitle.value = "HỢP ĐỒNG KINH TẾ";
            break;
          case "hd-nguyen-tac":
            txtContractTitle.value = "HỢP ĐỒNG NGUYÊN TẮC";
            break;
          case "hd-dich-vu":
            txtContractTitle.value = "HỢP ĐỒNG CUNG CẤP DỊCH VỤ";
            break;
          case "hd-mua-ban":
            txtContractTitle.value = "HỢP ĐỒNG MUA BÁN HÀNG HÓA";
            break;
          case "hd-tu-van":
            txtContractTitle.value = "HỢP ĐỒNG TƯ VẤN THIẾT KẾ VÀ GIÁM SÁT";
            break;
        }
      }
    });
  }

  // Sự kiện thay đổi Loại biên bản phụ
  const cboMinuteType = document.getElementById("cboMinuteType");
  const txtMinuteTitle = document.getElementById("txtMinuteTitle");
  if (cboMinuteType) {
    cboMinuteType.addEventListener("change", () => {
      updateDocSymbolString();
      if (txtMinuteTitle) {
        switch (cboMinuteType.value) {
          case "bb-nghiem-thu":
            txtMinuteTitle.value = "BIÊN BẢN NGHIỆM THU VÀ BÀN GIAO HẠNG MỤC THI CÔNG";
            break;
          case "bb-lam-viec":
            txtMinuteTitle.value = "BIÊN BẢN LÀM VIỆC";
            break;
          case "bb-ban-giao":
            txtMinuteTitle.value = "BIÊN BẢN BÀN GIAO THIẾT BỊ VÀ TÀI SẢN";
            break;
          case "bb-hop-giao-ban":
            txtMinuteTitle.value = "BIÊN BẢN HỌP GIAO BAN ĐỊNH KỲ";
            break;
        }
      }
    });
  }

  // Sự kiện thay đổi Loại tờ trình phụ
  const cboReportType = document.getElementById("cboReportType");
  if (cboReportType) {
    cboReportType.addEventListener("change", () => {
      updateDocSymbolString();
    });
  }

  // Gõ tên cơ quan ban hành -> cập nhật ký hiệu và Bên B
  if (txtPublishingOrg) {
    txtPublishingOrg.addEventListener("input", () => {
      updateDocSymbolString();
      const txtPartyB = document.getElementById("txtPartyBName");
      if (txtPartyB && !txtPartyB.value) {
        txtPartyB.placeholder = txtPublishingOrg.value || "Tên đơn vị ban hành";
      }
    });
  }

  // Sự kiện địa danh và lịch ngày
  if (txtLocation) {
    txtLocation.addEventListener("input", () => updateFullDateString());
    txtLocation.addEventListener("change", () => updateFullDateString());
  }

  if (dtpDocDate) {
    dtpDocDate.addEventListener("change", () => {
      updateFullDateString();
      setStatus("Đã cập nhật ngày ban hành từ lịch.");
    });
  }

  if (txtDocDate) {
    txtDocDate.addEventListener("input", handleDateManualInput);
    txtDocDate.addEventListener("change", handleDateManualInput);
  }

  // ToolStrip: Nạp lại
  const btnReload = document.getElementById("btnReload");
  if (btnReload) {
    btnReload.addEventListener("click", () => {
      window.location.reload();
    });
  }

  // ToolStrip: Trợ giúp NĐ 30
  const btnHelp = document.getElementById("btnHelp");
  if (btnHelp) {
    btnHelp.addEventListener("click", showND30Help);
  }

  // --- ĐIỀU HƯỚNG CÁC NÚT MENU CHÍNH: MỞ CỬA SỔ CHỨC NĂNG TƯƠNG ỨNG ---
  const btnNavStandardize = document.getElementById("btnNavStandardize");
  if (btnNavStandardize) {
    btnNavStandardize.addEventListener("click", () => {
      openStandardizeDialogWindow();
    });
  }

  const btnBackToMenu = document.getElementById("btnBackToMenu");
  if (btnBackToMenu) {
    btnBackToMenu.addEventListener("click", () => {
      switchToView("menu");
    });
  }

  const btnReloadMenu = document.getElementById("btnReloadMenu");
  if (btnReloadMenu) {
    btnReloadMenu.addEventListener("click", () => {
      window.location.reload();
    });
  }

  const btnMenuHelp = document.getElementById("btnMenuHelp");
  if (btnMenuHelp) {
    btnMenuHelp.addEventListener("click", () => {
      openHelpDialogWindow();
    });
  }

  const btnNavHelpND30 = document.getElementById("btnNavHelpND30");
  if (btnNavHelpND30) {
    btnNavHelpND30.addEventListener("click", () => {
      openHelpDialogWindow();
    });
  }

  const btnMenuOpenDialog = document.getElementById("btnMenuOpenDialog");
  if (btnMenuOpenDialog) {
    btnMenuOpenDialog.addEventListener("click", () => {
      openStandaloneDialogWindow();
    });
  }

  const btnNavPriceDialog = document.getElementById("btnNavPriceDialog");
  if (btnNavPriceDialog) {
    btnNavPriceDialog.addEventListener("click", () => {
      openStandaloneDialogWindow();
    });
  }

  const btnNavOcrDialog = document.getElementById("btnNavOcrDialog");
  if (btnNavOcrDialog) {
    btnNavOcrDialog.addEventListener("click", () => {
      openOcrDialogWindow();
    });
  }

  const btnNavInsertTemplate = document.getElementById("btnNavInsertTemplate");
  if (btnNavInsertTemplate) {
    btnNavInsertTemplate.addEventListener("click", () => {
      openTemplateDialogWindow();
    });
  }

  // Nút Xóa trắng Form
  const btnClearForm = document.getElementById("btnClearForm");
  if (btnClearForm) {
    btnClearForm.addEventListener("click", () => {
      document.getElementById("txtSuperiorOrg").value = "";
      document.getElementById("txtPublishingOrg").value = "";
      document.getElementById("txtDocSeqNumber").value = "001";
      document.getElementById("txtLocation").value = "TP. Hồ Chí Minh";
      const now = new Date();
      const nowY = now.getFullYear();
      const nowM = String(now.getMonth() + 1).padStart(2, "0");
      const nowD = String(now.getDate()).padStart(2, "0");
      dtpDocDate.value = `${nowY}-${nowM}-${nowD}`;
      updateFullDateString(nowD, nowM, nowY);
      updateDocSymbolString();
      setStatus("Đã đặt lại form về trạng thái mặc định.");
    });
  }

  // Nút hành động chính
  const btnStandardizeAll = document.getElementById("btnStandardizeAll");
  if (btnStandardizeAll) {
    btnStandardizeAll.addEventListener("click", async () => {
      await standardizeDocument();
    });
  }

  // Nút chèn mẫu văn bản mới
  const btnInsertTemplate = document.getElementById("btnInsertTemplate");
  if (btnInsertTemplate) {
    btnInsertTemplate.addEventListener("click", async () => {
      await insertFullTemplate();
    });
  }
}

/* ==========================================================================
   2. BỘ TRÍCH XUẤT NỘI DUNG TỪ WORD (PARSER)
   ========================================================================== */

async function extractDocumentRawLines(context) {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("text");

  const tables = body.tables;
  tables.load("items");

  await context.sync();

  const lines = [];

  for (let i = 0; i < paragraphs.items.length; i++) {
    const text = paragraphs.items[i].text;
    if (!text) continue;
    const splitLines = text.split(/[\r\n\u000b\x0c]+/);
    for (const raw of splitLines) {
      const trimmed = raw.trim();
      if (trimmed && !/^[\s\-_—–=]{3,}$/.test(trimmed)) {
        lines.push(trimmed);
      }
    }
  }

  if (tables.items.length > 0) {
    for (const tbl of tables.items) {
      tbl.rows.load("items");
    }
    await context.sync();

    for (const tbl of tables.items) {
      for (const row of tbl.rows.items) {
        row.cells.load("items");
      }
    }
    await context.sync();

    for (const tbl of tables.items) {
      for (const row of tbl.rows.items) {
        for (const cell of row.cells.items) {
          cell.body.paragraphs.load("text");
        }
      }
    }
    await context.sync();

    for (const tbl of tables.items) {
      for (const row of tbl.rows.items) {
        for (const cell of row.cells.items) {
          for (const p of cell.body.paragraphs.items) {
            const pText = p.text ? p.text.trim() : "";
            if (pText && !/^[\s\-_—–=]{3,}$/.test(pText)) {
              if (!lines.includes(pText)) {
                lines.push(pText);
              }
            }
          }
        }
      }
    }
  }

  return lines;
}

function parseAdministrativeDoc(lines) {
  let superiorOrg = "";
  let publishingOrg = "";
  let docNumber = "";
  let docDate = "";
  let docTitle = "";
  let dear = "";
  let bodyLines = [];
  let signRole = "";
  let signNote = "";
  let signName = "";
  let recipients = [];

  const headerIndices = new Set();
  const footerIndices = new Set();
  let inRecipients = false;

  let idxCountry = -1;
  let idxNumber = -1;
  let idxTitle = -1;
  let idxDear = -1;
  let idxSignRole = -1;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/CỘNG\s*HÒA\s*XÃ\s*HỘI\s*CHỦ\s*NGHĨA\s*VIỆT\s*NAM/i.test(l) || /CỘNG\s*HÒA\s*XÃ\s*HỘI/i.test(l)) {
      idxCountry = i;
      headerIndices.add(i);
    } else if (/Độc\s*lập.*Tự\s*do.*Hạnh\s*phúc/i.test(l)) {
      headerIndices.add(i);
    } else if (/^\s*Số\s*[:\.]/i.test(l)) {
      idxNumber = i;
      docNumber = l;
      headerIndices.add(i);
    } else if (/(?:ngày\s+.*tháng\s+.*năm|ngày\s+\d+.*tháng|Địa\s*danh.*ngày)/i.test(l)) {
      docDate = l;
      headerIndices.add(i);
    } else if (/^\s*(Về\s*việc|V\/v)\s*[:\.]?\s*/i.test(l)) {
      if (idxTitle === -1) {
        idxTitle = i;
        docTitle = l;
      }
    } else if (/^\s*Kính\s*gửi\s*[:\.]?\s*/i.test(l)) {
      if (idxDear === -1) {
        idxDear = i;
        dear = l;
      }
    } else if (/^\s*(QUYỀN\s*HẠN.*CHỨC\s*VỤ|GIÁM\s*ĐỐC|TỔNG\s*GIÁM\s*ĐỐC|CHỦ\s*TỊCH|HIỆU\s*TRƯỞNG|TRƯỞNG\s*PHÒNG|THỦ\s*TRƯỞNG|TM\.|KT\.|TL\.|TUQ\.|ĐẠI\s*DIỆN\s*BÊN)/i.test(l)) {
      if (idxSignRole === -1 && i > Math.floor(lines.length / 3)) {
        idxSignRole = i;
        signRole = l;
        footerIndices.add(i);

        if (i + 1 < lines.length && /^\(/.test(lines[i + 1])) {
          signNote = lines[i + 1];
          footerIndices.add(i + 1);

          if (i + 2 < lines.length && !/Nơi\s*nhận/i.test(lines[i + 2]) && !/^[-+*•]/.test(lines[i + 2])) {
            signName = lines[i + 2];
            footerIndices.add(i + 2);
          }
        }
      }
    } else if (/^\s*Nơi\s*nhận\s*[:\.]?/i.test(l)) {
      inRecipients = true;
      footerIndices.add(i);
    } else if (inRecipients) {
      if (/^[-+*•]/.test(l) || /^\d+\./.test(l) || /^Lưu\s*[:\.]/i.test(l)) {
        recipients.push(l);
        footerIndices.add(i);
      } else {
        inRecipients = false;
      }
    }
  }

  // Bóc tách tên cơ quan
  if (idxNumber > 0) {
    let prev1 = idxNumber - 1;
    while (prev1 >= 0 && (headerIndices.has(prev1) || /^[\s\-_—–=]{3,}$/.test(lines[prev1]))) {
      prev1--;
    }
    if (prev1 >= 0 && !headerIndices.has(prev1)) {
      publishingOrg = lines[prev1];
      headerIndices.add(prev1);

      let prev2 = prev1 - 1;
      while (prev2 >= 0 && (headerIndices.has(prev2) || /^[\s\-_—–=]{3,}$/.test(lines[prev2]))) {
        prev2--;
      }
      if (prev2 >= 0 && !headerIndices.has(prev2)) {
        const cand = lines[prev2];
        if (/CHỦ\s*QUẢN|BỘ|SỞ|UBND|TỔNG\s*CÔNG\s*TY|ỦY\s*BAN|TẬP\s*ĐOÀN|VIỆN|HỌC\s*VIỆN/i.test(cand) || cand === cand.toUpperCase()) {
          superiorOrg = cand;
          headerIndices.add(prev2);
        }
      }
    }
  }

  // Thu thập các đoạn thân bài
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (
      !headerIndices.has(i) &&
      !footerIndices.has(i) &&
      i !== idxTitle &&
      i !== idxDear &&
      !/^[\s\-_—–=]{3,}$/.test(l)
    ) {
      bodyLines.push(l);
    }
  }

  return {
    superiorOrg,
    publishingOrg,
    docNumber,
    docDate,
    docTitle,
    dear,
    signRole,
    signNote,
    signName,
    recipients,
    bodyLines
  };
}

/* ==========================================================================
   3. HELPER CHÈN BẢNG TIÊU ĐỀ 2 CỘT (CÔNG VĂN, BIÊN BẢN, TỜ TRÌNH)
   ========================================================================== */

async function insertHeaderTable(context, body, superiorOrg, publishingOrg, docNumber, docDate, insertLocation) {
  const table = body.insertTable(1, 2, insertLocation, [["", ""]]);
  table.styleBuiltIn = "PlainTable1";
  table.alignment = Word.Alignment.centered;

  try {
    table.getBorder("All").type = "None";
  } catch (e) {}

  try {
    const columns = table.columns;
    columns.load("items");
    await context.sync();

    if (columns.items.length >= 2) {
      columns.items[0].setWidth(175, "None");
      columns.items[1].setWidth(288, "None");
    }
  } catch (colErr) {
    console.warn("Không thể gán column width:", colErr.message);
  }

  // CỘT TRÁI
  const cellLeft = table.getCell(0, 0);
  cellLeft.body.clear();

  if (superiorOrg) {
    const pSup = cellLeft.body.insertParagraph(superiorOrg.toUpperCase(), Word.InsertLocation.end);
    pSup.font.name = "Times New Roman";
    pSup.font.size = 12;
    pSup.font.bold = false;
    pSup.alignment = Word.Alignment.centered;
    pSup.firstLineIndent = 0;
    pSup.spaceBefore = 0;
    pSup.spaceAfter = 1;
  }

  const pOrg = cellLeft.body.insertParagraph((publishingOrg || "TÊN CƠ QUAN, ĐƠN VỊ").toUpperCase(), Word.InsertLocation.end);
  pOrg.font.name = "Times New Roman";
  pOrg.font.size = 12;
  pOrg.font.bold = true;
  pOrg.alignment = Word.Alignment.centered;
  pOrg.firstLineIndent = 0;
  pOrg.spaceBefore = 0;
  pOrg.spaceAfter = 1;

  const pOrgLine = cellLeft.body.insertParagraph("————————", Word.InsertLocation.end);
  pOrgLine.font.name = "Times New Roman";
  pOrgLine.font.size = 11;
  pOrgLine.font.bold = true;
  pOrgLine.alignment = Word.Alignment.centered;
  pOrgLine.firstLineIndent = 0;
  pOrgLine.spaceBefore = 0;
  pOrgLine.spaceAfter = 3;

  const pNum = cellLeft.body.insertParagraph(docNumber || "Số: 001/CV-...", Word.InsertLocation.end);
  pNum.font.name = "Times New Roman";
  pNum.font.size = 13;
  pNum.font.bold = false;
  pNum.alignment = Word.Alignment.centered;
  pNum.firstLineIndent = 0;
  pNum.spaceBefore = 0;
  pNum.spaceAfter = 0;

  // CỘT PHẢI
  const cellRight = table.getCell(0, 1);
  cellRight.body.clear();

  const pCountry = cellRight.body.insertParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", Word.InsertLocation.end);
  pCountry.font.name = "Times New Roman";
  pCountry.font.size = 12;
  pCountry.font.bold = true;
  pCountry.alignment = Word.Alignment.centered;
  pCountry.firstLineIndent = 0;
  pCountry.spaceBefore = 0;
  pCountry.spaceAfter = 2;

  const pMotto = cellRight.body.insertParagraph("Độc lập - Tự do - Hạnh phúc", Word.InsertLocation.end);
  pMotto.font.name = "Times New Roman";
  pMotto.font.size = 13;
  pMotto.font.bold = true;
  pMotto.alignment = Word.Alignment.centered;
  pMotto.firstLineIndent = 0;
  pMotto.spaceBefore = 0;
  pMotto.spaceAfter = 1;

  const pMottoLine = cellRight.body.insertParagraph("—————————————", Word.InsertLocation.end);
  pMottoLine.font.name = "Times New Roman";
  pMottoLine.font.size = 11;
  pMottoLine.font.bold = true;
  pMottoLine.alignment = Word.Alignment.centered;
  pMottoLine.firstLineIndent = 0;
  pMottoLine.spaceBefore = 0;
  pMottoLine.spaceAfter = 3;

  const pDate = cellRight.body.insertParagraph(docDate || generateCurrentDateString(), Word.InsertLocation.end);
  pDate.font.name = "Times New Roman";
  pDate.font.size = 13;
  pDate.font.bold = false;
  pDate.font.italic = true;
  pDate.alignment = Word.Alignment.centered;
  pDate.firstLineIndent = 0;
  pDate.spaceBefore = 0;
  pDate.spaceAfter = 0;

  return table;
}

/* ==========================================================================
   4. LOGIC CHUẨN HÓA THEO TỪNG LOẠI VĂN BẢN (ROUTING CHÍNH)
   ========================================================================== */

async function standardizeDocument(isInsertNewTemplate = false) {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Chỉ chạy được trong Microsoft Word Desktop.");
    alert("Chức năng chỉ hoạt động bên trong Microsoft Word!");
    return;
  }

  const cboDocType = document.getElementById("cboDocType");
  const docType = cboDocType ? cboDocType.value : "cong-van";

  const formSup = (document.getElementById("txtSuperiorOrg")?.value || "").trim();
  const formPub = (document.getElementById("txtPublishingOrg")?.value || "").trim();
  const formLoc = (document.getElementById("txtLocation")?.value || "").trim() || "TP. Hồ Chí Minh";
  const formDate = (document.getElementById("txtDocDate")?.value || "").trim();
  const formSeq = formatSequenceNumber(document.getElementById("txtDocSeqNumber")?.value || "001");
  let formSymbol = (document.getElementById("txtDocSymbol")?.value || "").trim();

  if (document.getElementById("txtDocSeqNumber")) {
    document.getElementById("txtDocSeqNumber").value = formSeq;
  }

  const actionName = isInsertNewTemplate ? "Chèn mẫu văn bản mới" : "Chuẩn hóa thể thức";
  setStatus(`Đang ${actionName}: ${docType.toUpperCase()}...`);

  try {
    await Word.run(async (context) => {
      const wordDoc = context.document;
      const body = wordDoc.body;

      let rawLines = [];
      let parsed = { bodyLines: [] };

      // Nếu không phải ép chèn mẫu mới thì quét nội dung văn bản hiện tại
      if (!isInsertNewTemplate) {
        rawLines = await extractDocumentRawLines(context);
        if (rawLines.length > 0) {
          parsed = parseAdministrativeDoc(rawLines);
        }
      }

      // Xác định tên cơ quan và ngày tháng cuối cùng
      let finalPub = formPub && !isPlaceholderText(formPub) ? formPub : (parsed.publishingOrg && !isPlaceholderText(parsed.publishingOrg) ? parsed.publishingOrg : (formPub || "TÊN CƠ QUAN, ĐƠN VỊ BAN HÀNH"));
      let finalSup = formSup && !isPlaceholderText(formSup) ? formSup : (parsed.superiorOrg && !isPlaceholderText(parsed.superiorOrg) ? parsed.superiorOrg : formSup);
      let finalDate = formDate && !isPlaceholderText(formDate) ? formDate : (parsed.docDate && !isPlaceholderText(parsed.docDate) ? parsed.docDate : generateCurrentDateString(formLoc));

      if (document.getElementById("txtPublishingOrg") && !document.getElementById("txtPublishingOrg").value) {
        document.getElementById("txtPublishingOrg").value = finalPub;
      }
      if (document.getElementById("txtDocDate")) {
        document.getElementById("txtDocDate").value = finalDate;
      }

      // Ký hiệu văn bản
      if (!formSymbol) {
        updateDocSymbolString();
        formSymbol = (document.getElementById("txtDocSymbol")?.value || "").trim();
      }
      if (!formSymbol.startsWith("/")) formSymbol = "/" + formSymbol;
      const finalNum = `Số: ${formSeq}${formSymbol}`;

      // 3. Cấu hình khổ giấy A4
      if (document.getElementById("chkMarginA4")?.checked) {
        try {
          const sections = wordDoc.sections;
          sections.load("items");
          await context.sync();
          if (sections.items.length > 0 && sections.items[0].pageSetup) {
            sections.items[0].pageSetup.topMargin = 56.7;    // 20mm
            sections.items[0].pageSetup.bottomMargin = 56.7; // 20mm
            sections.items[0].pageSetup.leftMargin = 85.05;  // 30mm
            sections.items[0].pageSetup.rightMargin = 42.52; // 15mm
          }
        } catch (e) {}
      }

      // 4. Xóa sạch nội dung cũ để tái cấu trúc
      body.clear();
      body.font.name = "Times New Roman";
      body.font.size = 13;
      body.font.color = "#000000";

      const cleanedBody = (parsed.bodyLines || []).filter(line => line.trim().length > 0 && !/^[\s\-_—–=]{3,}$/.test(line.trim()));

      // 5. Chuẩn hóa theo từng loại văn bản được chọn
      switch (docType) {
        case "hop-dong":
          await standardizeHopDong(context, body, { finalPub, finalNum, finalDate, cleanedBody });
          break;
        case "bien-ban":
          await standardizeBienBan(context, body, { finalPub, finalNum, finalDate, cleanedBody });
          break;
        case "to-trinh":
          await standardizeToTrinh(context, body, { finalSup, finalPub, finalNum, finalDate, cleanedBody });
          break;
        case "cong-van":
        default:
          await standardizeCongVan(context, body, { finalSup, finalPub, finalNum, finalDate, cleanedBody, parsed });
          break;
      }

      await context.sync();
      const doneMsg = isInsertNewTemplate ? `Đã chèn mẫu văn bản [${docType.toUpperCase()}] vào Word thành công!` : `Chuẩn hóa hoàn tất cho thể thức: ${docType.toUpperCase()}`;
      setStatus(doneMsg);
      alert(doneMsg);
    });
  } catch (error) {
    console.error("Lỗi thực thi Word API:", error);
    setStatus(`Lỗi xử lý: ${error.message}`);
    alert(`Lỗi thực thi Word API: ${error.message}`);
  }
}

/* ==========================================================================
   4a. CHUẨN HÓA CÔNG VĂN HÀNH CHÍNH (CV)
   ========================================================================== */

async function standardizeCongVan(context, body, opt) {
  const isHeaderTable = document.getElementById("chkHeaderTable").checked;
  const isFooterTable = document.getElementById("chkFooterTable").checked;
  const isIndentSpacing = document.getElementById("chkIndentSpacing").checked;

  // 1. Bảng Tiêu đề (Quốc hiệu & Cơ quan) 2 cột
  if (isHeaderTable) {
    await insertHeaderTable(context, body, opt.finalSup, opt.finalPub, opt.finalNum, opt.finalDate, Word.InsertLocation.start);
  }

  const pSpacer = body.insertParagraph("", Word.InsertLocation.end);
  pSpacer.spaceBefore = 0;
  pSpacer.spaceAfter = 6;
  pSpacer.firstLineIndent = 0;

  // 2. Trích yếu: "Về việc: ..."
  let rawTitle = document.getElementById("txtDocTitle") ? document.getElementById("txtDocTitle").value.trim() : (opt.parsed.docTitle || "");
  const titleText = normalizeVeViec(rawTitle);

  const pTitle = body.insertParagraph(titleText, Word.InsertLocation.end);
  pTitle.font.name = "Times New Roman";
  pTitle.font.size = 12;
  pTitle.font.bold = false;
  pTitle.alignment = Word.Alignment.centered;
  pTitle.firstLineIndent = 0;
  pTitle.spaceBefore = 4;
  pTitle.spaceAfter = 12;

  try {
    const sTitle = pTitle.search("Về việc:", { matchCase: false });
    sTitle.load("items");
    await context.sync();
    if (sTitle.items.length > 0) sTitle.items[0].font.bold = true;
  } catch (e) {}

  // 3. Kính gửi
  const dearVal = document.getElementById("txtDearCongVan")?.value.trim() || opt.parsed.dear || "Kính gửi: Tên cơ quan, tổ chức hoặc cá nhân nhận công văn";
  const pDear = body.insertParagraph(dearVal, Word.InsertLocation.end);
  pDear.font.name = "Times New Roman";
  pDear.font.size = 13;
  pDear.firstLineIndent = 36;
  pDear.spaceBefore = 0;
  pDear.spaceAfter = 8;

  try {
    const matchDear = pDear.search("Kính gửi:", { matchCase: false });
    matchDear.load("items");
    await context.sync();
    if (matchDear.items.length > 0) matchDear.items[0].font.bold = true;
  } catch (e) {}

  // 4. Các đoạn thân bài
  const bodyLines = opt.cleanedBody.length > 0 ? opt.cleanedBody : [
    "Lời mở đầu: Ghi lý do dẫn đến việc viết công văn (Dựa trên nhu cầu thực tế, kế hoạch hoạt động, hoặc theo yêu cầu công tác...).",
    "Nội dung chính: Ghi rõ các vấn đề cần giải quyết, đề xuất, yêu cầu hoặc thông báo thông tin chi tiết đến đơn vị nhận.",
    "Lời kết thúc: Khẳng định lại mục đích của công văn và mong muốn nhận được sự phối hợp, phản hồi từ phía cơ quan nhận."
  ];

  const prefixSearches = [];
  for (const line of bodyLines) {
    const p = body.insertParagraph(line, Word.InsertLocation.end);
    p.font.name = "Times New Roman";
    p.font.size = 13;
    p.font.bold = false;
    p.alignment = isIndentSpacing ? (Word.Alignment.justified || "Justified") : (Word.Alignment.left || "Left");
    p.firstLineIndent = isIndentSpacing ? 36 : 0;
    p.lineSpacing = 16.25;
    p.spaceBefore = 0;
    p.spaceAfter = 6;

    const colonMatch = line.match(/^([^:]{2,35}:)/);
    if (colonMatch) {
      const searchRes = p.search(colonMatch[1], { matchCase: true });
      searchRes.load("items");
      prefixSearches.push(searchRes);
    }
  }

  if (prefixSearches.length > 0) {
    await context.sync();
    for (const s of prefixSearches) {
      if (s.items && s.items.length > 0) s.items[0].font.bold = true;
    }
  }

  // 5. Bảng Nơi nhận & Chữ ký
  if (isFooterTable) {
    const signTable = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
    signTable.styleBuiltIn = "PlainTable1";
    signTable.alignment = Word.Alignment.centered;
    try { signTable.getBorder("All").type = "None"; } catch (e) {}

    try {
      const sCols = signTable.columns;
      sCols.load("items");
      await context.sync();
      if (sCols.items.length >= 2) {
        sCols.items[0].setWidth(200, "None");
        sCols.items[1].setWidth(260, "None");
      }
    } catch (e) {}

    const signLeft = signTable.getCell(0, 0);
    signLeft.body.clear();
    const pRecTitle = signLeft.body.insertParagraph("Nơi nhận:", Word.InsertLocation.end);
    pRecTitle.font.name = "Times New Roman";
    pRecTitle.font.bold = true;
    pRecTitle.font.italic = true;
    pRecTitle.font.size = 12;
    pRecTitle.firstLineIndent = 0;
    pRecTitle.spaceAfter = 2;

    const recLines = ["- Như trên;", "- Ban Giám đốc (để b/c);", "- Lưu: VT, HS."];
    for (const r of recLines) {
      const pR = signLeft.body.insertParagraph(r, Word.InsertLocation.end);
      pR.font.size = 11;
      pR.firstLineIndent = 0;
      pR.spaceAfter = 1;
    }

    const signRight = signTable.getCell(0, 1);
    signRight.body.clear();
    const pSignRole = signRight.body.insertParagraph((opt.parsed.signRole || "GIÁM ĐỐC").toUpperCase(), Word.InsertLocation.end);
    pSignRole.font.name = "Times New Roman";
    pSignRole.font.bold = true;
    pSignRole.font.size = 13;
    pSignRole.alignment = Word.Alignment.centered;
    pSignRole.firstLineIndent = 0;
    pSignRole.spaceAfter = 2;

    const pSignNote = signRight.body.insertParagraph("(Chữ ký, dấu của cơ quan)", Word.InsertLocation.end);
    pSignNote.font.name = "Times New Roman";
    pSignNote.font.italic = true;
    pSignNote.font.bold = false;
    pSignNote.font.size = 12;
    pSignNote.alignment = Word.Alignment.centered;
    pSignNote.firstLineIndent = 0;
    pSignNote.spaceAfter = 45;

    const pSignName = signRight.body.insertParagraph((opt.parsed.signName || "Họ và tên người ký").toUpperCase(), Word.InsertLocation.end);
    pSignName.font.name = "Times New Roman";
    pSignName.font.bold = true;
    pSignName.font.size = 13;
    pSignName.alignment = Word.Alignment.centered;
    pSignName.firstLineIndent = 0;
  }
}

/**
 * Trích xuất cụm "ngày ... tháng ... năm ..." từ chuỗi địa danh / ngày tháng
 */
function extractOnlyDate(dateStr) {
  if (!dateStr || typeof dateStr !== "string") {
    const today = new Date();
    return `ngày ${String(today.getDate()).padStart(2, "0")} tháng ${String(today.getMonth() + 1).padStart(2, "0")} năm ${today.getFullYear()}`;
  }
  const match = dateStr.match(/ngày\s+\d+.*tháng\s+\d+.*năm\s+\d+/i) || dateStr.match(/ngày\s+.*tháng\s+.*năm.*/i);
  return match ? match[0] : dateStr;
}

/* ==========================================================================
   4b. CHUẨN HÓA HỢP ĐỒNG KINH TẾ / DỊCH VỤ / THI CÔNG (HĐ)
   ========================================================================== */

async function standardizeHopDong(context, body, opt) {
  const isIndentSpacing = document.getElementById("chkIndentSpacing").checked;
  const contractTitle = document.getElementById("txtContractTitle")?.value.trim() || "HỢP ĐỒNG THI CÔNG XÂY DỰNG CÔNG TRÌNH";
  const projectName = document.getElementById("txtProjectName")?.value.trim();
  const biddingPackage = document.getElementById("txtBiddingPackage")?.value.trim();
  const projectLocation = document.getElementById("txtProjectLocation")?.value.trim();
  const contractValue = document.getElementById("txtContractValue")?.value.trim();

  const partyAName = document.getElementById("txtPartyAName")?.value.trim() || "CÔNG TY CỔ PHẦN ĐẦU TƯ BẤT ĐỘNG SẢN AN GIA";
  const partyARep = document.getElementById("txtPartyARep")?.value.trim() || "Ông Nguyễn Văn A - Chức vụ: Tổng Giám đốc";
  const partyBName = document.getElementById("txtPartyBName")?.value.trim() || opt.finalPub || "CÔNG TY TNHH XÂY DỰNG HUỲNH BẢO";
  const partyBRep = document.getElementById("txtPartyBRep")?.value.trim() || "Ông Huỳnh Bảo - Chức vụ: Giám đốc";

  // 1. Quốc hiệu & Tiêu ngữ căn giữa
  const pCountry = body.insertParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", Word.InsertLocation.end);
  pCountry.font.name = "Times New Roman";
  pCountry.font.size = 12;
  pCountry.font.bold = true;
  pCountry.alignment = Word.Alignment.centered;
  pCountry.firstLineIndent = 0;
  pCountry.spaceBefore = 0;
  pCountry.spaceAfter = 2;

  const pMotto = body.insertParagraph("Độc lập - Tự do - Hạnh phúc", Word.InsertLocation.end);
  pMotto.font.name = "Times New Roman";
  pMotto.font.size = 13;
  pMotto.font.bold = true;
  pMotto.alignment = Word.Alignment.centered;
  pMotto.firstLineIndent = 0;
  pMotto.spaceBefore = 0;
  pMotto.spaceAfter = 2;

  const pMottoLine = body.insertParagraph("—————————————", Word.InsertLocation.end);
  pMottoLine.font.name = "Times New Roman";
  pMottoLine.font.size = 11;
  pMottoLine.font.bold = true;
  pMottoLine.alignment = Word.Alignment.centered;
  pMottoLine.firstLineIndent = 0;
  pMottoLine.spaceBefore = 0;
  pMottoLine.spaceAfter = 14;

  // 2. Tiêu đề Hợp đồng (In hoa, 14pt, Bold, Centered)
  const pTitle = body.insertParagraph(contractTitle.toUpperCase(), Word.InsertLocation.end);
  pTitle.font.name = "Times New Roman";
  pTitle.font.size = 14;
  pTitle.font.bold = true;
  pTitle.alignment = Word.Alignment.centered;
  pTitle.firstLineIndent = 0;
  pTitle.spaceBefore = 4;
  pTitle.spaceAfter = 3;

  // Dòng Số hợp đồng
  const pNumHD = body.insertParagraph(opt.finalNum, Word.InsertLocation.end);
  pNumHD.font.name = "Times New Roman";
  pNumHD.font.size = 12;
  pNumHD.font.italic = true;
  pNumHD.alignment = Word.Alignment.centered;
  pNumHD.firstLineIndent = 0;
  pNumHD.spaceAfter = 10;

  // 3. Căn cứ pháp lý
  const grounds = [
    "- Căn cứ Bộ luật Dân sự số 91/2015/QH13 ngày 24/11/2015 của Quốc hội;",
    "- Căn cứ Luật Thương mại số 36/2005/QH11 ngày 14/06/2005 của Quốc hội;",
    (projectName || /thi\s*công|xây\s*dựng/i.test(contractTitle))
      ? "- Căn cứ Luật Xây dựng số 50/2014/QH13 và Luật sửa đổi số 62/2020/QH14 của Quốc hội;"
      : "- Căn cứ các văn bản quy phạm pháp luật có liên quan hiện hành;",
    "- Căn cứ vào nhu cầu và năng lực thực tế của hai bên."
  ];

  for (const g of grounds) {
    const pG = body.insertParagraph(g, Word.InsertLocation.end);
    pG.font.name = "Times New Roman";
    pG.font.size = 12;
    pG.font.italic = true;
    pG.firstLineIndent = 36;
    pG.spaceAfter = 2;
  }

  // 4. Thông tin dự án / công trình (nếu có)
  if (projectName || biddingPackage || projectLocation || contractValue) {
    const pProjBox = body.insertParagraph("", Word.InsertLocation.end);
    pProjBox.spaceBefore = 4;

    if (projectName) {
      const pP = body.insertParagraph(`- Tên công trình / Dự án: ${projectName}`, Word.InsertLocation.end);
      pP.font.size = 13; pP.firstLineIndent = 36; pP.spaceAfter = 2;
    }
    if (biddingPackage) {
      const pB = body.insertParagraph(`- Gói thầu: ${biddingPackage}`, Word.InsertLocation.end);
      pB.font.size = 13; pB.firstLineIndent = 36; pB.spaceAfter = 2;
    }
    if (projectLocation) {
      const pL = body.insertParagraph(`- Địa điểm thực hiện: ${projectLocation}`, Word.InsertLocation.end);
      pL.font.size = 13; pL.firstLineIndent = 36; pL.spaceAfter = 2;
    }
    if (contractValue) {
      const pV = body.insertParagraph(`- Giá trị hợp đồng: ${contractValue}`, Word.InsertLocation.end);
      pV.font.size = 13; pV.font.bold = true; pV.firstLineIndent = 36; pV.spaceAfter = 4;
    }
  }

  // 5. Thời gian mở đầu
  const pIntro = body.insertParagraph(`Hôm nay, ${extractOnlyDate(opt.finalDate)}, tại trụ sở Bên A, chúng tôi gồm các bên:`, Word.InsertLocation.end);
  pIntro.font.size = 13;
  pIntro.firstLineIndent = 36;
  pIntro.spaceBefore = 6;
  pIntro.spaceAfter = 6;

  // 6. Thông tin Bên A
  const pPA = body.insertParagraph("BÊN A (CHỦ ĐẦU TƯ / BÊN SỬ DỤNG DỊCH VỤ):", Word.InsertLocation.end);
  pPA.font.bold = true; pPA.font.size = 13; pPA.firstLineIndent = 18; pPA.spaceAfter = 2;
  const linesA = [`- Tên đơn vị: ${partyAName}`, `- Đại diện: ${partyARep}`];
  for (const l of linesA) {
    const p = body.insertParagraph(l, Word.InsertLocation.end);
    p.font.size = 13; p.firstLineIndent = 18; p.spaceAfter = 2;
  }

  // 7. Thông tin Bên B
  const pPB = body.insertParagraph("BÊN B (NHÀ THẦU / BÊN CUNG CẤP DỊCH VỤ):", Word.InsertLocation.end);
  pPB.font.bold = true; pPB.font.size = 13; pPB.firstLineIndent = 18; pPB.spaceBefore = 6; pPB.spaceAfter = 2;
  const linesB = [`- Tên công ty: ${partyBName}`, `- Đại diện: ${partyBRep}`];
  for (const l of linesB) {
    const p = body.insertParagraph(l, Word.InsertLocation.end);
    p.font.size = 13; p.firstLineIndent = 18; p.spaceAfter = 2;
  }

  const pLead = body.insertParagraph("Hai bên cùng thỏa thuận và nhất trí ký kết hợp đồng với các điều khoản sau đây:", Word.InsertLocation.end);
  pLead.font.italic = true; pLead.font.size = 13; pLead.firstLineIndent = 36; pLead.spaceBefore = 6; pLead.spaceAfter = 8;

  // 8. Nội dung điều khoản
  const contractBodyLines = opt.cleanedBody.length > 0 ? opt.cleanedBody : [
    "ĐIỀU 1: PHẠM VI CÔNG VIỆC VÀ YÊU CẦU KỸ THUẬT\nBên B cam kết thực hiện đúng thiết kế, tiêu chuẩn kỹ thuật và tiến độ công trình theo yêu cầu của Bên A.",
    "ĐIỀU 2: GIÁ TRỊ HỢP ĐỒNG VÀ PHƯƠNG THỨC THANH TOÁN\nGiá trị hợp đồng là giá trọn gói theo thỏa thuận. Bên A tiến hành thanh toán cho Bên B theo từng đợt nghiệm thu thực tế.",
    "ĐIỀU 3: QUYỀN VÀ NGHĨA VỤ CỦA CÁC BÊN\nCác bên cam kết thực hiện đầy đủ nghĩa vụ đã cam kết và phối hợp chặt chẽ trong suốt quá trình triển khai hợp đồng.",
    "ĐIỀU 4: ĐIỀU KHOẢN THI HÀNH\nHợp đồng có hiệu lực kể từ ngày ký. Được lập thành 04 bản có giá trị pháp lý như nhau, mỗi bên giữ 02 bản."
  ];

  const prefixSearches = [];
  for (const line of contractBodyLines) {
    const p = body.insertParagraph(line, Word.InsertLocation.end);
    p.font.name = "Times New Roman";
    p.font.size = 13;
    p.font.bold = false;
    p.alignment = isIndentSpacing ? (Word.Alignment.justified || "Justified") : (Word.Alignment.left || "Left");
    p.firstLineIndent = isIndentSpacing ? 36 : 0;
    p.lineSpacing = 16.25;
    p.spaceBefore = 0;
    p.spaceAfter = 6;

    const matchHead = line.match(/^(ĐIỀU\s*\d+[^:\n]*[:\.]?|Khoản\s*\d+[^:\n]*[:\.]?)/i);
    if (matchHead) {
      const searchRes = p.search(matchHead[1], { matchCase: false });
      searchRes.load("items");
      prefixSearches.push(searchRes);
    }
  }

  if (prefixSearches.length > 0) {
    await context.sync();
    for (const s of prefixSearches) {
      if (s.items && s.items.length > 0) s.items[0].font.bold = true;
    }
  }

  // 9. Bảng chữ ký 2 bên
  const signTable = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
  signTable.styleBuiltIn = "PlainTable1";
  signTable.alignment = Word.Alignment.centered;
  try { signTable.getBorder("All").type = "None"; } catch (e) {}

  try {
    const sCols = signTable.columns;
    sCols.load("items");
    await context.sync();
    if (sCols.items.length >= 2) {
      sCols.items[0].setWidth(230, "None");
      sCols.items[1].setWidth(230, "None");
    }
  } catch (e) {}

  // Bên A
  const cellA = signTable.getCell(0, 0);
  cellA.body.clear();
  const pSignA = cellA.body.insertParagraph("ĐẠI DIỆN BÊN A", Word.InsertLocation.end);
  pSignA.font.bold = true; pSignA.font.size = 13; pSignA.alignment = Word.Alignment.centered; pSignA.spaceAfter = 2;
  const pSignANote = cellA.body.insertParagraph("(Ký, ghi rõ họ tên và đóng dấu)", Word.InsertLocation.end);
  pSignANote.font.italic = true; pSignANote.font.size = 12; pSignANote.alignment = Word.Alignment.centered; pSignANote.spaceAfter = 45;
  const pSignAName = cellA.body.insertParagraph(partyARep.split("-")[0].replace("Ông|Bà", "").trim(), Word.InsertLocation.end);
  pSignAName.font.bold = true; pSignAName.font.size = 13; pSignAName.alignment = Word.Alignment.centered;

  // Bên B
  const cellB = signTable.getCell(0, 1);
  cellB.body.clear();
  const pSignB = cellB.body.insertParagraph("ĐẠI DIỆN BÊN B", Word.InsertLocation.end);
  pSignB.font.bold = true; pSignB.font.size = 13; pSignB.alignment = Word.Alignment.centered; pSignB.spaceAfter = 2;
  const pSignBNote = cellB.body.insertParagraph("(Ký, ghi rõ họ tên và đóng dấu)", Word.InsertLocation.end);
  pSignBNote.font.italic = true; pSignBNote.font.size = 12; pSignBNote.alignment = Word.Alignment.centered; pSignBNote.spaceAfter = 45;
  const pSignBName = cellB.body.insertParagraph(partyBRep.split("-")[0].replace("Ông|Bà", "").trim(), Word.InsertLocation.end);
  pSignBName.font.bold = true; pSignBName.font.size = 13; pSignBName.alignment = Word.Alignment.centered;
}

/* ==========================================================================
   4c. CHUẨN HÓA BIÊN BẢN LÀM VIỆC / BÀN GIAO / NGHIỆM THU (BB)
   ========================================================================== */

async function standardizeBienBan(context, body, opt) {
  const isIndentSpacing = document.getElementById("chkIndentSpacing").checked;
  const minuteTitle = document.getElementById("txtMinuteTitle")?.value.trim() || "BIÊN BẢN NGHIỆM THU VÀ BÀN GIAO CÔNG VIỆC";
  const minuteSubject = document.getElementById("txtMinuteSubject")?.value.trim() || "V/v: Nghiệm thu hoàn thành hạng mục thi công và bàn giao đưa vào sử dụng";
  const minuteTime = document.getElementById("txtMinuteTime")?.value.trim() || "09 giờ 00 phút, ngày 18 tháng 10 năm 2026";
  const minuteLocation = document.getElementById("txtMinuteLocation")?.value.trim() || "Tại văn phòng Ban Chỉ huy công trường";
  const minuteChair = document.getElementById("txtMinuteChair")?.value.trim() || "Ông Nguyễn Văn A - Chỉ huy trưởng";
  const minuteSecretary = document.getElementById("txtMinuteSecretary")?.value.trim() || "Bà Trần Thị B - Cán bộ kỹ thuật";
  const minuteAttendees = document.getElementById("txtMinuteAttendees")?.value.trim() || "Đại diện Chủ đầu tư, Tư vấn giám sát và Nhà thầu thi công";

  // 1. Header Table 2 cột
  await insertHeaderTable(context, body, "", opt.finalPub, opt.finalNum, opt.finalDate, Word.InsertLocation.start);

  const pSpacer = body.insertParagraph("", Word.InsertLocation.end);
  pSpacer.spaceBefore = 0; pSpacer.spaceAfter = 6;

  // 2. Tiêu đề Biên bản
  const pTitle = body.insertParagraph(minuteTitle.toUpperCase(), Word.InsertLocation.end);
  pTitle.font.name = "Times New Roman";
  pTitle.font.size = 14;
  pTitle.font.bold = true;
  pTitle.alignment = Word.Alignment.centered;
  pTitle.spaceBefore = 4;
  pTitle.spaceAfter = 3;

  const pSubj = body.insertParagraph(minuteSubject, Word.InsertLocation.end);
  pSubj.font.italic = true;
  pSubj.font.size = 12;
  pSubj.alignment = Word.Alignment.centered;
  pSubj.spaceAfter = 12;

  // 3. I. THỜI GIAN VÀ ĐỊA ĐIỂM
  const pI = body.insertParagraph("I. THỜI GIAN VÀ ĐỊA ĐIỂM", Word.InsertLocation.end);
  pI.font.bold = true; pI.font.size = 13; pI.firstLineIndent = 18; pI.spaceAfter = 3;
  const pTime = body.insertParagraph(`- Thời gian bắt đầu: ${minuteTime}`, Word.InsertLocation.end);
  pTime.firstLineIndent = 36; pTime.spaceAfter = 2;
  const pLoc = body.insertParagraph(`- Địa điểm: ${minuteLocation}`, Word.InsertLocation.end);
  pLoc.firstLineIndent = 36; pLoc.spaceAfter = 8;

  // 4. II. THÀNH PHẦN THAM DỰ
  const pII = body.insertParagraph("II. THÀNH PHẦN THAM DỰ", Word.InsertLocation.end);
  pII.font.bold = true; pII.font.size = 13; pII.firstLineIndent = 18; pII.spaceAfter = 3;
  const pChair = body.insertParagraph(`- Chủ trì / Chủ tọa: ${minuteChair}`, Word.InsertLocation.end);
  pChair.firstLineIndent = 36; pChair.spaceAfter = 2;
  const pSec = body.insertParagraph(`- Thư ký cuộc họp: ${minuteSecretary}`, Word.InsertLocation.end);
  pSec.firstLineIndent = 36; pSec.spaceAfter = 2;
  const pAtt = body.insertParagraph(`- Thành phần tham gia khác: ${minuteAttendees}`, Word.InsertLocation.end);
  pAtt.firstLineIndent = 36; pAtt.spaceAfter = 8;

  // 5. III. NỘI DUNG VÀ DIỄN BIẾN LÀM VIỆC
  const pIII = body.insertParagraph("III. NỘI DUNG VÀ DIỄN BIẾN LÀM VIỆC", Word.InsertLocation.end);
  pIII.font.bold = true; pIII.font.size = 13; pIII.firstLineIndent = 18; pIII.spaceAfter = 4;

  const minBodyLines = opt.cleanedBody.length > 0 ? opt.cleanedBody : [
    "Các bên đã tiến hành kiểm tra thực tế hiện trường và đối chiếu hồ sơ hoàn công các hạng mục thi công.",
    "Ý kiến của các bên tham gia: Các bên thống nhất đánh giá hạng mục thi công đạt yêu cầu kỹ thuật theo hồ sơ thiết kế đã được phê duyệt.",
    "Các tồn tại cần khắc phục (nếu có): Nhà thầu cam kết dọn dẹp vệ sinh mặt bằng trước ngày bàn giao chính thức."
  ];

  for (const line of minBodyLines) {
    const p = body.insertParagraph(line, Word.InsertLocation.end);
    p.font.size = 13;
    p.alignment = isIndentSpacing ? (Word.Alignment.justified || "Justified") : (Word.Alignment.left || "Left");
    p.firstLineIndent = isIndentSpacing ? 36 : 0;
    p.lineSpacing = 16.25;
    p.spaceAfter = 6;
  }

  // 6. IV. KẾT LUẬN
  const pIV = body.insertParagraph("IV. KẾT LUẬN", Word.InsertLocation.end);
  pIV.font.bold = true; pIV.font.size = 13; pIV.firstLineIndent = 18; pIV.spaceBefore = 4; pIV.spaceAfter = 3;
  const pEnd = body.insertParagraph(
    "Biên bản được đọc lại cho toàn thể các bên tham gia cùng nghe và nhất trí ký tên xác nhận. Biên bản kết thúc cùng ngày và được lập thành 02 bản có giá trị pháp lý như nhau.",
    Word.InsertLocation.end
  );
  pEnd.firstLineIndent = 36; pEnd.alignment = Word.Alignment.justified || "Justified"; pEnd.lineSpacing = 16.25; pEnd.spaceAfter = 14;

  // 7. Bảng Chữ ký Thư ký & Chủ tọa
  const signTable = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
  signTable.styleBuiltIn = "PlainTable1";
  signTable.alignment = Word.Alignment.centered;
  try { signTable.getBorder("All").type = "None"; } catch (e) {}

  try {
    const sCols = signTable.columns;
    sCols.load("items");
    await context.sync();
    if (sCols.items.length >= 2) {
      sCols.items[0].setWidth(230, "None");
      sCols.items[1].setWidth(230, "None");
    }
  } catch (e) {}

  const cellSec = signTable.getCell(0, 0);
  cellSec.body.clear();
  const pSignSec = cellSec.body.insertParagraph("THƯ KÝ CUỘC HỌP", Word.InsertLocation.end);
  pSignSec.font.bold = true; pSignSec.font.size = 13; pSignSec.alignment = Word.Alignment.centered; pSignSec.spaceAfter = 2;
  const pSignSecNote = cellSec.body.insertParagraph("(Ký và ghi rõ họ tên)", Word.InsertLocation.end);
  pSignSecNote.font.italic = true; pSignSecNote.font.size = 12; pSignSecNote.alignment = Word.Alignment.centered; pSignSecNote.spaceAfter = 45;
  const pSignSecName = cellSec.body.insertParagraph(minuteSecretary.split("-")[0].replace("Ông|Bà", "").trim(), Word.InsertLocation.end);
  pSignSecName.font.bold = true; pSignSecName.font.size = 13; pSignSecName.alignment = Word.Alignment.centered;

  const cellChair = signTable.getCell(0, 1);
  cellChair.body.clear();
  const pSignChair = cellChair.body.insertParagraph("CHỦ TỌA / CHỦ TRÌ", Word.InsertLocation.end);
  pSignChair.font.bold = true; pSignChair.font.size = 13; pSignChair.alignment = Word.Alignment.centered; pSignChair.spaceAfter = 2;
  const pSignChairNote = cellChair.body.insertParagraph("(Ký, ghi rõ họ tên và đóng dấu)", Word.InsertLocation.end);
  pSignChairNote.font.italic = true; pSignChairNote.font.size = 12; pSignChairNote.alignment = Word.Alignment.centered; pSignChairNote.spaceAfter = 45;
  const pSignChairName = cellChair.body.insertParagraph(minuteChair.split("-")[0].replace("Ông|Bà", "").trim(), Word.InsertLocation.end);
  pSignChairName.font.bold = true; pSignChairName.font.size = 13; pSignChairName.alignment = Word.Alignment.centered;
}

/* ==========================================================================
   4d. CHUẨN HÓA TỜ TRÌNH / BÁO CÁO (TTr)
   ========================================================================== */

async function standardizeToTrinh(context, body, opt) {
  const isIndentSpacing = document.getElementById("chkIndentSpacing").checked;
  const reportTitle = document.getElementById("txtReportTitle")?.value.trim() || "V/v xin phê duyệt kế hoạch thi công và dự toán kinh phí năm 2026";
  const reportDear = document.getElementById("txtReportDear")?.value.trim() || "Kính gửi: Ban Giám đốc Công ty";
  const reportBasis = document.getElementById("txtReportBasis")?.value.trim() || "Căn cứ Kế hoạch năm 2026; Căn cứ năng lực và nhu cầu thực tế của đơn vị.";

  // 1. Header Table 2 cột
  await insertHeaderTable(context, body, opt.finalSup, opt.finalPub, opt.finalNum, opt.finalDate, Word.InsertLocation.start);

  const pSpacer = body.insertParagraph("", Word.InsertLocation.end);
  pSpacer.spaceBefore = 0; pSpacer.spaceAfter = 6;

  // 2. Tiêu đề Tờ trình
  const pTitle = body.insertParagraph("TỜ TRÌNH", Word.InsertLocation.end);
  pTitle.font.name = "Times New Roman";
  pTitle.font.size = 14;
  pTitle.font.bold = true;
  pTitle.alignment = Word.Alignment.centered;
  pTitle.spaceBefore = 4;
  pTitle.spaceAfter = 3;

  const pSubj = body.insertParagraph(reportTitle, Word.InsertLocation.end);
  pSubj.font.bold = true;
  pSubj.font.size = 12;
  pSubj.alignment = Word.Alignment.centered;
  pSubj.spaceAfter = 12;

  // 3. Kính gửi
  const pDear = body.insertParagraph(reportDear, Word.InsertLocation.end);
  pDear.font.size = 13;
  pDear.firstLineIndent = 36;
  pDear.spaceAfter = 6;
  try {
    const sDear = pDear.search("Kính gửi:", { matchCase: false });
    sDear.load("items");
    await context.sync();
    if (sDear.items.length > 0) sDear.items[0].font.bold = true;
  } catch (e) {}

  // 4. Căn cứ
  if (reportBasis) {
    const pBasis = body.insertParagraph(reportBasis, Word.InsertLocation.end);
    pBasis.font.italic = true;
    pBasis.font.size = 12;
    pBasis.firstLineIndent = 36;
    pBasis.spaceAfter = 8;
  }

  // 5. Nội dung tờ trình
  const reportBodyLines = opt.cleanedBody.length > 0 ? opt.cleanedBody : [
    "1. Sự cần thiết / Lý do trình duyệt:\nNhằm đảm bảo tiến độ triển khai công tác sản xuất kinh doanh và đáp ứng yêu cầu chất lượng công trình năm 2026.",
    "2. Nội dung chi tiết đề xuất:\nKính đề nghị cấp có thẩm quyền xem xét thông qua danh mục hạng mục công việc và phương án tổ chức thi công chi tiết.",
    "3. Dự kiến kinh phí và nguồn vốn thực hiện:\nTổng mức kinh phí dự kiến đã được rà soát phù hợp với định mức đơn giá hiện hành."
  ];

  const prefixSearches = [];
  for (const line of reportBodyLines) {
    const p = body.insertParagraph(line, Word.InsertLocation.end);
    p.font.size = 13;
    p.alignment = isIndentSpacing ? (Word.Alignment.justified || "Justified") : (Word.Alignment.left || "Left");
    p.firstLineIndent = isIndentSpacing ? 36 : 0;
    p.lineSpacing = 16.25;
    p.spaceAfter = 6;

    const matchHead = line.match(/^(\d+\.[^:\n]*[:\.]?)/i);
    if (matchHead) {
      const searchRes = p.search(matchHead[1], { matchCase: false });
      searchRes.load("items");
      prefixSearches.push(searchRes);
    }
  }

  if (prefixSearches.length > 0) {
    await context.sync();
    for (const s of prefixSearches) {
      if (s.items && s.items.length > 0) s.items[0].font.bold = true;
    }
  }

  // 6. Lời kết
  const pClose = body.insertParagraph("Kính trình Ban Giám đốc xem xét, phê duyệt để đơn vị có cơ sở tổ chức triển khai thực hiện.", Word.InsertLocation.end);
  pClose.firstLineIndent = 36;
  pClose.alignment = Word.Alignment.justified || "Justified";
  pClose.lineSpacing = 16.25;
  pClose.spaceAfter = 14;

  // 7. Bảng Nơi nhận & Chữ ký
  const signTable = body.insertTable(1, 2, Word.InsertLocation.end, [["", ""]]);
  signTable.styleBuiltIn = "PlainTable1";
  signTable.alignment = Word.Alignment.centered;
  try { signTable.getBorder("All").type = "None"; } catch (e) {}

  try {
    const sCols = signTable.columns;
    sCols.load("items");
    await context.sync();
    if (sCols.items.length >= 2) {
      sCols.items[0].setWidth(200, "None");
      sCols.items[1].setWidth(260, "None");
    }
  } catch (e) {}

  const signLeft = signTable.getCell(0, 0);
  signLeft.body.clear();
  const pRec = signLeft.body.insertParagraph("Nơi nhận:", Word.InsertLocation.end);
  pRec.font.bold = true; pRec.font.italic = true; pRec.font.size = 12; pRec.spaceAfter = 2;
  const pR1 = signLeft.body.insertParagraph("- Như trên;", Word.InsertLocation.end);
  pR1.font.size = 11; pR1.spaceAfter = 1;
  const pR2 = signLeft.body.insertParagraph("- Lưu: VT, ĐV.", Word.InsertLocation.end);
  pR2.font.size = 11; pR2.spaceAfter = 1;

  const signRight = signTable.getCell(0, 1);
  signRight.body.clear();
  const pRole = signRight.body.insertParagraph("THỦ TRƯỞNG ĐƠN VỊ TRÌNH", Word.InsertLocation.end);
  pRole.font.bold = true; pRole.font.size = 13; pRole.alignment = Word.Alignment.centered; pRole.spaceAfter = 2;
  const pNote = signRight.body.insertParagraph("(Ký và ghi rõ họ tên)", Word.InsertLocation.end);
  pNote.font.italic = true; pNote.font.size = 12; pNote.alignment = Word.Alignment.centered; pNote.spaceAfter = 45;
  const pName = signRight.body.insertParagraph("Họ và tên người trình", Word.InsertLocation.end);
  pName.font.bold = true; pName.font.size = 13; pName.alignment = Word.Alignment.centered;
}

/* ==========================================================================
   5. CHÈN MẪU VĂN BẢN MỚI
   ========================================================================== */

async function insertFullTemplate() {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Chỉ chạy được trong Microsoft Word Desktop.");
    alert("Chức năng chỉ hoạt động bên trong Microsoft Word!");
    return;
  }

  // Gọi chuẩn hóa ở chế độ chèn mẫu văn bản mới
  await standardizeDocument(true);
}

/* ==========================================================================
   6. WINFORMS DEMO MODAL DIALOG: BẢNG GIÁ KÝ HỢP ĐỒNG (TEST CỬA SỔ)
   ========================================================================== */

/**
 * Định dạng tiền tệ kiểu Việt Nam (VD: 1.850.000)
 */
function formatCurrency(num) {
  if (isNaN(num)) return "0";
  return Math.round(num).toLocaleString("vi-VN");
}

/**
 * Định dạng số có dấu chấm hàng nghìn nhảy theo thời gian thực (real-time)
 */
function formatDots(val) {
  if (val === null || val === undefined || val === "") return "";
  const num = typeof val === "number" ? Math.round(val) : parseInt(String(val).replace(/[^\d]/g, ""), 10);
  if (isNaN(num)) return "";
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

/**
 * Phân tích chuỗi số có dấu chấm thành số nguyên thuần
 */
function parseDots(val) {
  if (!val) return 0;
  const clean = String(val).replace(/[^\d]/g, "");
  return parseInt(clean, 10) || 0;
}

/**
 * Đọc số tiền thành chữ tiếng Việt chuẩn xác
 */
function convertNumberToVietnameseWords(n) {
  if (isNaN(n) || n <= 0) return "Không đồng";
  const units = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];
  const digits = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];

  function readGroup(group, isLeading) {
    const c = Math.floor(group / 100);
    const t = Math.floor((group % 100) / 10);
    const d = group % 10;
    let res = "";

    if (c > 0 || !isLeading) res += digits[c] + " trăm ";
    if (t > 1) {
      res += digits[t] + " mươi ";
      if (d === 1) res += "mốt ";
      else if (d === 5) res += "lăm ";
      else if (d > 0) res += digits[d] + " ";
    } else if (t === 1) {
      res += "mười ";
      if (d === 5) res += "lăm ";
      else if (d > 0) res += digits[d] + " ";
    } else if (d > 0) {
      if (c > 0 || !isLeading) res += "lẻ ";
      res += digits[d] + " ";
    }
    return res.trim();
  }

  let num = Math.round(n);
  const groups = [];
  while (num > 0) {
    groups.push(num % 1000);
    num = Math.floor(num / 1000);
  }

  let text = "";
  for (let i = groups.length - 1; i >= 0; i--) {
    const g = groups[i];
    if (g > 0) {
      const gText = readGroup(g, i === groups.length - 1);
      text += gText + " " + units[i] + " ";
    }
  }

  text = text.trim();
  if (!text) return "Không đồng";
  text = text.charAt(0).toUpperCase() + text.slice(1) + " đồng chẵn.";
  return text.replace(/\s+/g, " ");
}

// Bộ dữ liệu mẫu cho từng loại bảng giá
const DEMO_PRICE_DATA = {
  "thi-cong": [
    { name: "Thi công đào móng, vận chuyển đất đổ thải", unit: "m3", qty: 250, price: 120000, note: "Máy + thủ công" },
    { name: "Gia công lắp dựng cốt thép móng, dầm, cột", unit: "tấn", qty: 8.5, price: 24500000, note: "Thép CB300" },
    { name: "Bê tông thương phẩm mác 300 móng và dầm sàn", unit: "m3", qty: 180, price: 1450000, note: "Độ sụt 12±2" },
    { name: "Xây tường gạch ống dày 100-200 vữa xi măng mác 75", unit: "m2", qty: 650, price: 210000, note: "Gạch Tuynel" },
    { name: "Trát tường trong và ngoài nhà vữa XM mác 75", unit: "m2", qty: 1300, price: 110000, note: "Dày 1.5cm" },
    { name: "Lắp đặt hệ thống cấp thoát nước và thiết bị vệ sinh", unit: "gói", qty: 1, price: 185000000, note: "Trọn gói" }
  ],
  "mua-ban": [
    { name: "Xi măng Portland hỗn hợp PCB40 bao 50kg", unit: "tấn", qty: 45, price: 1750000, note: "Hà Tiên" },
    { name: "Thép vằn xây dựng phi 10 - 25 Hòa Phát", unit: "tấn", qty: 12, price: 17200000, note: "Tiêu chuẩn ASTM" },
    { name: "Cát vàng xây tô hạt vừa sàng lọc sạch", unit: "m3", qty: 85, price: 360000, note: "Cát sông Lô" },
    { name: "Gạch men lát nền 600x600 Viglacera cao cấp", unit: "m2", qty: 220, price: 195000, note: "Loại 1" },
    { name: "Dây cáp điện Cadivi 2x4.0 mm2 cuộn 100m", unit: "cuộn", qty: 15, price: 1250000, note: "Chính hãng" },
    { name: "Sơn nước nội thất Dulux EasyClean lau chùi hiệu quả", unit: "thùng", qty: 25, price: 1650000, note: "Thùng 18L" }
  ],
  "dich-vu": [
    { name: "Dịch vụ bảo trì, bảo dưỡng hệ thống điều hòa không khí", unit: "hệ thống", qty: 1, price: 35000000, note: "Định kỳ 6 tháng" },
    { name: "Dịch vụ kiểm định an toàn thiết bị nâng và PCCC", unit: "lần", qty: 2, price: 18000000, note: "Cấp chứng nhận" },
    { name: "Dịch vụ hỗ trợ kỹ thuật và giám sát thi công chuyên nghiệp", unit: "tháng", qty: 6, price: 25000000, note: "Kỹ sư túc trực" }
  ]
};

let currentPriceItems = [];
let selectedGridRowIndex = -1;

/**
 * Mở Cửa sổ phần mềm độc lập chạy song song với Microsoft Word
 */
function openStandaloneDialogWindow() {
  const projName = document.getElementById("txtProjectName")?.value.trim() || "";
  const bidPackage = document.getElementById("txtBiddingPackage")?.value.trim() || "";
  const projLocation = document.getElementById("txtProjectLocation")?.value.trim() || "";

  const queryParams = new URLSearchParams({
    v: "2.7.0",
    project: projName,
    package: bidPackage,
    location: projLocation
  });
  const dialogUrl = new URL(`dialog.html?${queryParams.toString()}`, window.location.href).href;
  const features = "width=1040,height=720,left=120,top=60,resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no";

  const syncToModalInputs = () => {
    const dlgProj = document.getElementById("txtDlgProjectName");
    const dlgPack = document.getElementById("txtDlgBiddingPackage");
    const dlgLoc = document.getElementById("txtDlgProjectLocation");
    if (dlgProj) dlgProj.value = projName;
    if (dlgPack) dlgPack.value = bidPackage;
    if (dlgLoc) dlgLoc.value = projLocation;
  };

  let standaloneWin = null;
  try {
    standaloneWin = window.open(dialogUrl, "WinForms_FormBaoGia", features);
  } catch (e) {
    console.warn("window.open thất bại:", e);
  }

  if (!standaloneWin || standaloneWin.closed || typeof standaloneWin.closed === "undefined") {
    // Nếu trình duyệt hoặc WebView chặn popup, thử dùng Office.context.ui.displayDialogAsync
    if (Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 75, width: 65, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const dlgOverlay = document.getElementById("dlgPriceScheduleOverlay");
          if (dlgOverlay) {
            dlgOverlay.classList.remove("hidden");
            syncToModalInputs();
            loadPriceTemplateData("thi-cong");
          }
          setStatus("Cảnh báo: Không thể mở cửa sổ độc lập (Popup blocked). Đã mở cửa sổ nội bộ.");
        } else {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            try {
              const data = JSON.parse(arg.message);
              await handleDialogMessage(data);
            } catch (err) {}
          });
          setStatus("Đã mở cửa sổ độc lập qua Office Dialog API.");
        }
      });
    } else {
      const dlgOverlay = document.getElementById("dlgPriceScheduleOverlay");
      if (dlgOverlay) {
        dlgOverlay.classList.remove("hidden");
        syncToModalInputs();
        loadPriceTemplateData("thi-cong");
      }
    }
  } else {
    standaloneWin.focus();
    setStatus("Đã mở cửa sổ phần mềm độc lập: Bảng giá ký hợp đồng.");
  }
}

/**
 * Mở Cửa sổ Trích xuất OCR AI Qwen2.5-VL-3B Local song song với Word
 */
function openOcrDialogWindow() {
  const dialogUrl = new URL("ocr-dialog.html?v=" + Date.now(), window.location.href).href;
  const width = 1180;
  const height = 820;
  const left = Math.max(0, Math.round((window.screen.availWidth - width) / 2));
  const top = Math.max(0, Math.round((window.screen.availHeight - height) / 2));
  const features = `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no`;

  let ocrWin = null;
  try {
    ocrWin = window.open(dialogUrl, "WinForms_FormBocTachOCR", features);
  } catch (e) {
    console.warn("window.open OCR dialog thất bại:", e);
  }

  if (!ocrWin || ocrWin.closed || typeof ocrWin.closed === "undefined") {
    if (Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 85, width: 80, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          alert("Trình duyệt chặn mở pop-up. Vui lòng cho phép mở cửa sổ độc lập!");
        } else {
          setStatus("Đã mở cửa sổ OCR qua Office Dialog API.");
        }
      });
    } else {
      window.open(dialogUrl, "_blank");
    }
  } else {
    ocrWin.focus();
    setStatus("Đã mở Cửa sổ Bóc tách OCR AI Qwen2.5-VL-3B Local.");
  }
}

/**
 * Mở Cửa sổ Chuẩn hóa Thể thức Văn bản (NĐ 30/2020/NĐ-CP) song song với Word
 */
function openStandardizeDialogWindow() {
  const docType = document.getElementById("cboDocType")?.value || "cong-van";
  const supOrg = document.getElementById("txtSuperiorOrg")?.value.trim() || "";
  const pubOrg = document.getElementById("txtPublishingOrg")?.value.trim() || "";
  const seqNum = document.getElementById("txtDocSeqNumber")?.value.trim() || "001";
  const proj = document.getElementById("txtProjectName")?.value.trim() || "";
  const pack = document.getElementById("txtBiddingPackage")?.value.trim() || "";
  const loc = document.getElementById("txtProjectLocation")?.value.trim() || "";

  const queryParams = new URLSearchParams({
    v: "2.7.0",
    docType: docType,
    superiorOrg: supOrg,
    publishingOrg: pubOrg,
    docSeqNumber: seqNum,
    project: proj,
    package: pack,
    location: loc
  });
  const dialogUrl = new URL(`standardize-dialog.html?${queryParams.toString()}`, window.location.href).href;
  const width = 1060;
  const height = 760;
  const left = Math.max(0, Math.round((window.screen.availWidth - width) / 2));
  const top = Math.max(0, Math.round((window.screen.availHeight - height) / 2));
  const features = `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no`;

  let win = null;
  try {
    win = window.open(dialogUrl, "WinForms_FormStandardizeND30", features);
  } catch (e) {
    console.warn("window.open standardize dialog thất bại:", e);
  }

  if (!win || win.closed || typeof win.closed === "undefined") {
    if (Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 85, width: 75, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          switchToView("standardize");
          setStatus("Đã chuyển sang Form chuẩn hóa nội bộ trong Taskpane.");
        } else {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            try {
              const data = JSON.parse(arg.message);
              await handleDialogMessage(data);
            } catch (err) {}
          });
          setStatus("Đã mở cửa sổ Chuẩn hóa NĐ 30 qua Office Dialog API.");
        }
      });
    } else {
      switchToView("standardize");
    }
  } else {
    win.focus();
    setStatus("Đã mở Cửa sổ Chuẩn hóa Thể thức Văn bản (NĐ 30).");
  }
}

/**
 * Mở Cửa sổ Thư viện Mẫu Văn bản Chuẩn NĐ 30 song song với Word
 */
function openTemplateDialogWindow() {
  const dialogUrl = new URL("template-dialog.html?v=" + Date.now(), window.location.href).href;
  const width = 1040;
  const height = 740;
  const left = Math.max(0, Math.round((window.screen.availWidth - width) / 2));
  const top = Math.max(0, Math.round((window.screen.availHeight - height) / 2));
  const features = `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no`;

  let win = null;
  try {
    win = window.open(dialogUrl, "WinForms_FormTemplateLibrary", features);
  } catch (e) {
    console.warn("window.open template dialog thất bại:", e);
  }

  if (!win || win.closed || typeof win.closed === "undefined") {
    if (Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 80, width: 75, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          switchToView("standardize");
          const btnInsert = document.getElementById("btnInsertTemplate");
          if (btnInsert) btnInsert.click();
        } else {
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            try {
              const data = JSON.parse(arg.message);
              await handleDialogMessage(data);
            } catch (err) {}
          });
          setStatus("Đã mở Thư viện Mẫu Văn bản qua Office Dialog API.");
        }
      });
    } else {
      switchToView("standardize");
      const btnInsert = document.getElementById("btnInsertTemplate");
      if (btnInsert) btnInsert.click();
    }
  } else {
    win.focus();
    setStatus("Đã mở Cửa sổ Thư viện Mẫu Văn bản Chuẩn NĐ 30.");
  }
}

/**
 * Mở Cửa sổ Sổ tay Quy chuẩn Thể thức NĐ 30 song song với Word
 */
function openHelpDialogWindow() {
  const dialogUrl = new URL("help-dialog.html?v=" + Date.now(), window.location.href).href;
  const width = 1020;
  const height = 700;
  const left = Math.max(0, Math.round((window.screen.availWidth - width) / 2));
  const top = Math.max(0, Math.round((window.screen.availHeight - height) / 2));
  const features = `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no`;

  let win = null;
  try {
    win = window.open(dialogUrl, "WinForms_FormHelpND30", features);
  } catch (e) {
    console.warn("window.open help dialog thất bại:", e);
  }

  if (!win || win.closed || typeof win.closed === "undefined") {
    if (Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(dialogUrl, { height: 75, width: 70, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          showND30Help();
        } else {
          setStatus("Đã mở Sổ tay Quy chuẩn NĐ 30 qua Office Dialog API.");
        }
      });
    } else {
      showND30Help();
    }
  } else {
    win.focus();
    setStatus("Đã mở Cửa sổ Sổ tay Quy chuẩn Thể thức NĐ 30.");
  }
}

/**
 * Khởi tạo kênh lắng nghe BroadcastChannel từ Cửa sổ OCR
 */
function initOcrBroadcastReceiver() {
  try {
    const channel = new BroadcastChannel('WordAddin_OCR_Channel');
    channel.onmessage = async (event) => {
      const { type, blocks } = event.data || {};
      if (type === 'INSERT_OCR_BLOCKS' && Array.isArray(blocks)) {
        console.log('[BroadcastChannel] Nhận yêu cầu chèn dữ liệu OCR vào Word:', blocks);
        await insertOcrBlocksToWord(blocks);
      }
    };
  } catch (err) {
    console.warn('Không thể khởi tạo BroadcastChannel cho OCR:', err);
  }
}

/**
 * Chèn các khối dữ liệu OCR và bảng Markdown vào tài liệu Word đang mở
 */
async function insertOcrBlocksToWord(blocks) {
  if (typeof Word === 'undefined' || !Word.run) return;
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      let isFirstPage = true;
      for (const block of blocks) {
        if (block.type === 'page_break') {
          if (!isFirstPage) {
            body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
          }
          isFirstPage = false;
        } else if (block.type === 'heading') {
          const p = body.insertParagraph(block.text, Word.InsertLocation.end);
          p.font.name = 'Times New Roman';
          p.font.size = 13;
          p.font.bold = true;
          p.alignment = Word.Alignment.center;
          p.spaceBefore = 8;
          p.spaceAfter = 4;
        } else if (block.type === 'paragraph') {
          const p = body.insertParagraph(block.text, Word.InsertLocation.end);
          p.font.name = 'Times New Roman';
          p.font.size = 13;
          p.lineSpacing = 1.15 * 12;
          p.spaceAfter = 4;
          p.alignment = Word.Alignment.justified;
        } else if (block.type === 'table') {
          const rows = block.data;
          if (rows.length > 0) {
            const rowCount = rows.length;
            const colCount = Math.max(...rows.map(r => r.length));
            const values = rows.map(r => {
              const rowCopy = [...r];
              while (rowCopy.length < colCount) rowCopy.push('');
              return rowCopy;
            });

            const tbl = body.insertTable(rowCount, colCount, Word.InsertLocation.end, values);
            tbl.font.name = 'Times New Roman';
            tbl.font.size = 11;
            tbl.alignment = Word.Alignment.center;
            tbl.headerRowCount = 1;
            tbl.style = 'Table Grid';

            const headerRow = tbl.rows.getFirst();
            headerRow.font.bold = true;
          }
        }
      }

      await context.sync();
      setStatus("Đã chèn toàn bộ dữ liệu OCR vào Word thành công!");
    });
  } catch (err) {
    console.error('Lỗi khi chèn dữ liệu OCR vào Word:', err);
    setStatus(`Lỗi chèn Word: ${err.message}`);
  }
}

/**
 * Đồng bộ dữ liệu nhận từ cửa sổ Chuẩn hóa sang Form Taskpane
 */
function syncDataToTaskpaneForm(d) {
  if (!d) return;

  const setVal = (id, val) => {
    if (val === undefined || val === null) return;
    const el = document.getElementById(id);
    if (el) el.value = val;
  };

  const setChk = (id, val) => {
    if (val === undefined || val === null) return;
    const el = document.getElementById(id);
    if (el) el.checked = Boolean(val);
  };

  if (d.docType) {
    setVal("cboDocType", d.docType);
    switchDocTypePanels(d.docType);
  }

  setVal("txtSuperiorOrg", d.superiorOrg);
  setVal("txtPublishingOrg", d.publishingOrg);
  setVal("txtDocSeqNumber", d.docSeqNumber);
  setVal("txtDocSymbol", d.docSymbol);
  setVal("txtLocation", d.location);
  setVal("txtDocDate", d.docDate);

  // Công văn
  setVal("txtDocTitle", d.docTitle);
  setVal("txtDearCongVan", d.dearCongVan);

  // Hợp đồng
  if (d.contractType) setVal("cboContractType", d.contractType);
  setVal("txtContractTitle", d.contractTitle);
  setVal("txtProjectName", d.projectName);
  setVal("txtBiddingPackage", d.biddingPackage);
  setVal("txtProjectLocation", d.projectLocation);
  setVal("txtContractValue", d.contractValue);
  setVal("txtPartyAName", d.partyAName);
  setVal("txtPartyARep", d.partyARep);
  setVal("txtPartyBName", d.partyBName);
  setVal("txtPartyBRep", d.partyBRep);

  // Biên bản
  if (d.minuteType) setVal("cboMinuteType", d.minuteType);
  setVal("txtMinuteTitle", d.minuteTitle);
  setVal("txtMinuteSubject", d.minuteSubject);
  setVal("txtMinuteTime", d.minuteTime);
  setVal("txtMinuteLocation", d.minuteLocation);
  setVal("txtMinuteChair", d.minuteChair);
  setVal("txtMinuteSecretary", d.minuteSecretary);
  setVal("txtMinuteAttendees", d.minuteAttendees);

  // Tờ trình
  if (d.reportType) setVal("cboReportType", d.reportType);
  setVal("txtReportTitle", d.reportTitle);
  setVal("txtReportDear", d.reportDear);
  setVal("txtReportBasis", d.reportBasis);

  // Checkboxes
  setChk("chkSmartParse", d.chkSmartParse);
  setChk("chkHeaderTable", d.chkHeaderTable);
  setChk("chkFooterTable", d.chkFooterTable);
  setChk("chkMarginA4", d.chkMarginA4);
  setChk("chkIndentSpacing", d.chkIndentSpacing);
  setChk("chkFontTimes", d.chkFontTimes);
}

/**
 * Xử lý yêu cầu chèn mẫu được chọn từ Thư viện mẫu vào Word
 */
async function handleInsertSelectedTemplate(data) {
  const cbo = document.getElementById("cboDocType");
  if (data.docType && cbo) {
    cbo.value = data.docType;
    switchDocTypePanels(data.docType);
    updateDocSymbolString();
  }
  await insertFullTemplate();
  setStatus(`Đã chèn mẫu [${data.title}] vào Microsoft Word thành công.`);
}

/**
 * Áp dụng căn lề trang chuẩn A4 theo Nghị định 30/2020/NĐ-CP vào Word
 */
async function applyA4MarginsToWord() {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Microsoft Word API không sẵn sàng.");
    return;
  }
  try {
    await Word.run(async (context) => {
      const section = context.document.sections.getFirst();
      // 1 mm = 72 / 25.4 pt = 2.8346 pt
      section.pageSetup.topMargin = 56.7;    // 20 mm
      section.pageSetup.bottomMargin = 56.7; // 20 mm
      section.pageSetup.leftMargin = 85.05;  // 30 mm
      section.pageSetup.rightMargin = 42.52; // 15 mm
      await context.sync();
      setStatus("Đã thiết lập căn lề A4 NĐ 30: Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm thành công.");
      alert("Đã thiết lập căn lề chuẩn A4 theo Nghị định 30 vào Word thành công!\n(Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm)");
    });
  } catch (err) {
    console.error("Lỗi căn lề:", err);
    setStatus("Lỗi căn lề: " + err.message);
  }
}

/**
 * Chuẩn hóa Font chữ toàn bộ văn bản về Times New Roman size 13
 */
async function applyTimesFontToWord() {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Microsoft Word API không sẵn sàng.");
    return;
  }
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.font.name = "Times New Roman";
      body.font.size = 13;
      await context.sync();
      setStatus("Đã chuẩn hóa toàn bộ văn bản về font Times New Roman size 13.");
      alert("Đã chuẩn hóa toàn bộ văn bản về Font Times New Roman (size 13) thành công!");
    });
  } catch (err) {
    console.error("Lỗi chuẩn hóa font:", err);
    setStatus("Lỗi font: " + err.message);
  }
}

/**
 * Xử lý dữ liệu nhận được từ tất cả các Cửa sổ phần mềm độc lập
 */
async function handleDialogMessage(data) {
  if (!data || !data.action) return;

  // Ngăn chặn gọi lặp cùng 1 hành động trong vòng 1000ms (do BroadcastChannel + window.opener đồng thời bắn)
  const now = Date.now();
  const actionSignature = data.action + "_" + (data.templateKey || "") + "_" + (data.data ? JSON.stringify(data.data.docType) : "");
  if (actionSignature === lastProcessedActionId && (now - lastProcessedActionTime) < 1200) {
    console.log(`[Dedup] Bỏ qua thông điệp trùng lặp: ${actionSignature}`);
    return;
  }
  lastProcessedActionId = actionSignature;
  lastProcessedActionTime = now;

  if (data.action === "SYNC_TOTAL") {
    const txtVal = document.getElementById("txtContractValue");
    if (txtVal && data.formattedString) {
      txtVal.value = data.formattedString;
    }
    const txtProj = document.getElementById("txtProjectName");
    if (txtProj && data.projectName) {
      txtProj.value = data.projectName;
    }
    const txtPack = document.getElementById("txtBiddingPackage");
    if (txtPack && data.biddingPackage) {
      txtPack.value = data.biddingPackage;
    }
    const txtLoc = document.getElementById("txtProjectLocation");
    if (txtLoc && data.projectLocation) {
      txtLoc.value = data.projectLocation;
    }
    setStatus(`Đã nhận tổng giá trị từ Bảng giá: ${data.formattedString}`);
  } else if (data.action === "INSERT_WORD_TABLE") {
    await insertPriceTableToWordFromData(data);
  } else if (data.action === "STANDARDIZE_WORD") {
    if (data.data) syncDataToTaskpaneForm(data.data);
    await standardizeDocument();
  } else if (data.action === "INSERT_TEMPLATE_WORD") {
    if (data.data) syncDataToTaskpaneForm(data.data);
    await insertFullTemplate();
  } else if (data.action === "SYNC_STANDARDIZE_FORM") {
    if (data.data) syncDataToTaskpaneForm(data.data);
    setStatus("Đã đồng bộ thông tin thể thức từ Cửa sổ độc lập về Add-in Word.");
  } else if (data.action === "INSERT_SELECTED_TEMPLATE") {
    await handleInsertSelectedTemplate(data);
  } else if (data.action === "APPLY_A4_MARGINS") {
    await applyA4MarginsToWord();
  } else if (data.action === "APPLY_TIMES_FONT") {
    await applyTimesFontToWord();
  }
}

// Các hàm toàn cục cho gọi trực tiếp qua window.opener
window.onReceiveSyncFromDialog = (payload) => { handleDialogMessage(payload); };
window.onReceiveWordTableFromDialog = async (payload) => { await handleDialogMessage(payload); };
window.onReceiveStandardizeFromDialog = async (payload) => { await handleDialogMessage(payload); };
window.onReceiveInsertTemplateFromDialog = async (payload) => { await handleDialogMessage(payload); };
window.onReceiveSyncStandardizeFromDialog = (payload) => { handleDialogMessage(payload); };
window.onReceiveTemplateFromDialog = async (payload) => { await handleDialogMessage(payload); };
window.onReceiveApplyMarginsFromDialog = async (payload) => { await handleDialogMessage(payload); };
window.onReceiveApplyFontFromDialog = async (payload) => { await handleDialogMessage(payload); };

// Kênh BroadcastChannel nhận thông điệp qua các cửa sổ song song
if (typeof BroadcastChannel !== "undefined") {
  const channel = new BroadcastChannel("WordAddin_WinForms_Channel");
  channel.onmessage = async (event) => {
    await handleDialogMessage(event.data);
  };
}

// Lắng nghe sự kiện postMessage
window.addEventListener("message", async (event) => {
  if (event.data && event.data.action) {
    await handleDialogMessage(event.data);
  }
});

/**
 * Chèn bảng giá vào Microsoft Word từ dữ liệu gửi qua cửa sổ độc lập
 */
async function insertPriceTableToWordFromData(payload) {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Microsoft Word API không sẵn sàng.");
    return;
  }

  const items = payload.items || [];
  if (items.length === 0) return;

  const projectName = payload.projectName || "";
  const biddingPackage = payload.biddingPackage || "";
  const projectLocation = payload.projectLocation || "";
  const templateName = payload.templateName || "Bảng giá hợp đồng";
  const vatRate = payload.vatRate || 10;
  const subtotal = payload.subtotal || 0;
  const vatAmount = payload.vatAmount || 0;
  const grandTotal = payload.grandTotal || 0;
  const inWords = payload.inWords || convertNumberToVietnameseWords(grandTotal);

  let subtitleParts = [templateName];
  if (projectName) subtitleParts.push(`Công trình: ${projectName}`);
  if (biddingPackage) subtitleParts.push(`Gói thầu: ${biddingPackage}`);
  if (projectLocation) subtitleParts.push(`Địa điểm: ${projectLocation}`);
  const subtitleStr = `(${subtitleParts.join(" - ")})`;

  setStatus("Đang chèn bảng giá từ cửa sổ độc lập vào Word...");

  try {
    await Word.run(async (context) => {
      const wordDoc = context.document;
      const body = wordDoc.body;

      // 1. Tiêu đề bảng
      const pTitle = body.insertParagraph("BẢNG TỔNG HỢP DỰ TOÁN / GIÁ TRỊ KÝ HỢP ĐỒNG", Word.InsertLocation.end);
      pTitle.font.name = "Times New Roman";
      pTitle.font.size = 13;
      pTitle.font.bold = true;
      pTitle.alignment = Word.Alignment.centered;
      pTitle.spaceBefore = 14;
      pTitle.spaceAfter = 2;

      const pSub = body.insertParagraph(subtitleStr, Word.InsertLocation.end);
      pSub.font.name = "Times New Roman";
      pSub.font.size = 12;
      pSub.font.italic = true;
      pSub.alignment = Word.Alignment.centered;
      pSub.spaceAfter = 8;

      // 2. Tạo ma trận dữ liệu: Header row (1) + Item rows (N) + Summary rows (4)
      const rowsCount = 1 + items.length + 4;
      const colsCount = 6;
      const matrix = [];

      // Dòng 1: Header
      matrix.push(["STT", "Nội dung công việc / Hàng hóa vật tư", "ĐVT", "Khối lượng", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)"]);

      // Các dòng dữ liệu
      items.forEach((item, idx) => {
        const itemAmount = (item.qty || 0) * (item.price || 0);
        matrix.push([
          String(idx + 1),
          item.name || "",
          item.unit || "",
          String(item.qty || 0),
          formatCurrency(item.price || 0),
          formatCurrency(itemAmount)
        ]);
      });

      // Các dòng tổng kết
      matrix.push(["", "Cộng giá trị trước thuế:", "", "", "", formatCurrency(subtotal)]);
      matrix.push(["", `Thuế giá trị gia tăng GTGT (${vatRate}%):`, "", "", "", formatCurrency(vatAmount)]);
      matrix.push(["", "TỔNG CỘNG THANH TOÁN:", "", "", "", formatCurrency(grandTotal)]);
      matrix.push(["", `(Bằng chữ: ${inWords})`, "", "", "", ""]);

      // Chèn bảng
      const table = body.insertTable(rowsCount, colsCount, Word.InsertLocation.end, matrix);
      table.styleBuiltIn = "TableGrid";
      table.alignment = Word.Alignment.centered;

      // Định dạng độ rộng cột
      try {
        const columns = table.columns;
        columns.load("items");
        await context.sync();
        if (columns.items.length >= 6) {
          columns.items[0].setWidth(30, "None");  // STT
          columns.items[1].setWidth(185, "None"); // Tên
          columns.items[2].setWidth(40, "None");  // ĐVT
          columns.items[3].setWidth(55, "None");  // KL
          columns.items[4].setWidth(75, "None");  // Đơn giá
          columns.items[5].setWidth(85, "None");  // Thành tiền
        }
      } catch (e) {}

      // Định dạng font và căn lề
      try {
        const rows = table.rows;
        rows.load("items");
        await context.sync();

        if (rows.items.length > 0) {
          rows.items[0].font.bold = true;
          rows.items[0].font.name = "Times New Roman";
          rows.items[0].font.size = 11;

          const rSub = rows.items[rows.items.length - 4];
          rSub.font.bold = true;

          const rVat = rows.items[rows.items.length - 3];
          rVat.font.bold = true;

          const rGrand = rows.items[rows.items.length - 2];
          rGrand.font.bold = true;
          rGrand.font.size = 12;

          const rWords = rows.items[rows.items.length - 1];
          rWords.font.italic = true;
        }
      } catch (e) {}

      const pSpacer = body.insertParagraph("", Word.InsertLocation.end);
      pSpacer.spaceAfter = 12;

      await context.sync();
      setStatus(`Đã chèn bảng giá (${items.length} hạng mục) từ cửa sổ độc lập vào Microsoft Word thành công.`);
    });
  } catch (error) {
    console.error("Lỗi chèn bảng giá từ cửa sổ độc lập:", error);
    setStatus(`Lỗi chèn bảng: ${error.message}`);
  }
}

/**
 * Khởi tạo dữ liệu và sự kiện cho Cửa sổ Demo Bảng giá (WinForms Dialog)
 */
function initPriceDialogEvents() {
  const dlgOverlay = document.getElementById("dlgPriceScheduleOverlay");
  const btnOpen = document.getElementById("btnOpenDemoDialog");
  const btnClose = document.getElementById("btnCloseDialog");
  const btnCloseBottom = document.getElementById("btnCloseDialogBottom");
  const cboType = document.getElementById("cboPriceTemplateType");
  const cboVat = document.getElementById("cboDlgVatRate");
  const btnAdd = document.getElementById("btnAddGridRow");
  const btnDel = document.getElementById("btnDeleteGridRow");
  const btnReset = document.getElementById("btnResetGridRows");
  const btnInsertToWord = document.getElementById("btnInsertTableToWord");
  const btnApplyToForm = document.getElementById("btnApplyTotalToForm");

  // Mở cửa sổ demo độc lập song song với Word
  if (btnOpen) {
    btnOpen.addEventListener("click", () => {
      openStandaloneDialogWindow();
    });
  }

  // Đóng cửa sổ nội bộ (nếu dùng)
  const closeDlg = () => {
    if (dlgOverlay) dlgOverlay.classList.add("hidden");
    setStatus("Đã đóng cửa sổ demo.");
  };

  if (btnClose) btnClose.addEventListener("click", closeDlg);
  if (btnCloseBottom) btnCloseBottom.addEventListener("click", closeDlg);

  // Nhấn ESC để đóng dialog
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && dlgOverlay && !dlgOverlay.classList.contains("hidden")) {
      closeDlg();
    }
  });

  // Thay đổi loại mẫu bảng giá
  if (cboType) {
    cboType.addEventListener("change", () => {
      loadPriceTemplateData(cboType.value);
    });
  }

  // Thay đổi tỷ lệ VAT
  if (cboVat) {
    cboVat.addEventListener("change", () => {
      recalcPriceTotals();
    });
  }

  // Thêm dòng mới vào grid
  if (btnAdd) {
    btnAdd.addEventListener("click", () => {
      currentPriceItems.push({
        name: "Hạng mục công việc / Vật tư mới",
        unit: "gói",
        qty: 1,
        price: 10000000,
        note: ""
      });
      renderPriceGrid();
      selectedGridRowIndex = currentPriceItems.length - 1;
      highlightSelectedRow();
    });
  }

  // Xóa dòng đang chọn
  if (btnDel) {
    btnDel.addEventListener("click", () => {
      if (currentPriceItems.length === 0) return;
      const targetIdx = selectedGridRowIndex >= 0 ? selectedGridRowIndex : currentPriceItems.length - 1;
      currentPriceItems.splice(targetIdx, 1);
      selectedGridRowIndex = Math.min(targetIdx, currentPriceItems.length - 1);
      renderPriceGrid();
    });
  }

  // Nạp lại dữ liệu mẫu
  if (btnReset) {
    btnReset.addEventListener("click", () => {
      loadPriceTemplateData(cboType ? cboType.value : "thi-cong");
    });
  }

  // Chèn bảng giá vào Microsoft Word
  if (btnInsertToWord) {
    btnInsertToWord.addEventListener("click", async () => {
      await insertPriceTableToWord();
    });
  }

  // Lấy giá trị về Form chính
  if (btnApplyToForm) {
    btnApplyToForm.addEventListener("click", () => {
      applyPriceTotalToForm();
    });
  }
}

/**
 * Nạp dữ liệu mẫu vào mảng và vẽ lại bảng
 */
function loadPriceTemplateData(typeKey) {
  const source = DEMO_PRICE_DATA[typeKey] || DEMO_PRICE_DATA["thi-cong"];
  currentPriceItems = source.map(item => ({ ...item }));
  selectedGridRowIndex = -1;
  renderPriceGrid();
}

/**
 * Vẽ lại toàn bộ hàng của DataGridView
 */
function renderPriceGrid() {
  const tbody = document.getElementById("tblPriceGridBody");
  if (!tbody) return;
  tbody.innerHTML = "";

  currentPriceItems.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = index;

    const amount = (item.qty || 0) * (item.price || 0);

    tr.innerHTML = `
      <td class="text-center" style="font-weight: 600; color: #555555;">${index + 1}</td>
      <td>
        <input type="text" class="win-grid-input" data-field="name" value="${item.name || ""}" />
      </td>
      <td class="text-center">
        <input type="text" list="unitOptionsListTaskpane" class="win-grid-input text-center" data-field="unit" value="${item.unit || "gói"}" placeholder="Chọn/nhập" />
      </td>
      <td class="text-right">
        <input type="number" step="any" class="win-grid-input text-right" data-field="qty" value="${item.qty || 0}" />
      </td>
      <td class="text-right">
        <input type="text" class="win-grid-input text-right" data-field="price" value="${formatDots(item.price || 0)}" placeholder="0" />
      </td>
      <td class="text-right" style="font-weight: 600; font-family: Consolas, monospace; color: #003399;">
        ${formatCurrency(amount)}
      </td>
      <td>
        <input type="text" class="win-grid-input" data-field="note" value="${item.note || ""}" />
      </td>
    `;

    // Click chọn dòng
    tr.addEventListener("click", () => {
      selectedGridRowIndex = index;
      highlightSelectedRow();
    });

    // Lắng nghe thay đổi dữ liệu trên ô input
    tr.addEventListener("input", (e) => {
      const field = e.target.dataset.field;
      if (!field) return;

      if (field === "price") {
        const rawDigits = e.target.value.replace(/[^\d]/g, "");
        const numVal = parseInt(rawDigits, 10) || 0;
        e.target.value = formatDots(rawDigits);
        currentPriceItems[index].price = numVal;
      } else if (field === "qty") {
        currentPriceItems[index].qty = parseFloat(e.target.value) || 0;
      } else {
        currentPriceItems[index][field] = e.target.value;
      }

      // Cập nhật ngay ô Thành tiền của dòng
      const updatedAmount = (currentPriceItems[index].qty || 0) * (currentPriceItems[index].price || 0);
      const cellAmount = tr.children[5];
      if (cellAmount) cellAmount.textContent = formatCurrency(updatedAmount);

      recalcPriceTotals();
    });

    tbody.appendChild(tr);
  });

  highlightSelectedRow();
  recalcPriceTotals();
}

/**
 * Đánh dấu dòng đang được chọn trong DataGridView
 */
function highlightSelectedRow() {
  const tbody = document.getElementById("tblPriceGridBody");
  if (!tbody) return;
  const rows = tbody.querySelectorAll("tr");
  rows.forEach((r, idx) => {
    if (idx === selectedGridRowIndex) {
      r.classList.add("selected");
    } else {
      r.classList.remove("selected");
    }
  });
}

/**
 * Tính toán lại tổng tiền trước thuế, VAT và tổng cộng thanh toán
 */
function recalcPriceTotals() {
  let subtotal = 0;
  currentPriceItems.forEach(item => {
    subtotal += (item.qty || 0) * (item.price || 0);
  });

  const cboVat = document.getElementById("cboDlgVatRate");
  const vatRate = cboVat ? parseInt(cboVat.value, 10) : 10;
  const vatAmount = subtotal * (vatRate / 100);
  const grandTotal = subtotal + vatAmount;

  const lblSub = document.getElementById("lblSubTotal");
  const lblVatText = document.getElementById("lblVatText");
  const lblVat = document.getElementById("lblVatAmount");
  const lblGrand = document.getElementById("lblGrandTotal");
  const lblWords = document.getElementById("lblTotalInWords");

  if (lblSub) lblSub.textContent = `${formatCurrency(subtotal)} VNĐ`;
  if (lblVatText) lblVatText.textContent = `Tiền thuế GTGT (${vatRate}%):`;
  if (lblVat) lblVat.textContent = `${formatCurrency(vatAmount)} VNĐ`;
  if (lblGrand) lblGrand.textContent = `${formatCurrency(grandTotal)} VNĐ`;
  if (lblWords) lblWords.textContent = convertNumberToVietnameseWords(grandTotal);
}

/**
 * Cập nhật tổng giá trị tính được vào ô Giá trị hợp đồng trên Form chính
 */
function applyPriceTotalToForm() {
  let subtotal = 0;
  currentPriceItems.forEach(item => {
    subtotal += (item.qty || 0) * (item.price || 0);
  });

  const cboVat = document.getElementById("cboDlgVatRate");
  const vatRate = cboVat ? parseInt(cboVat.value, 10) : 10;
  const grandTotal = subtotal + subtotal * (vatRate / 100);

  const formattedStr = `${formatCurrency(grandTotal)} VNĐ (${convertNumberToVietnameseWords(grandTotal)})`;

  const txtVal = document.getElementById("txtContractValue");
  if (txtVal) {
    txtVal.value = formattedStr;
  }

  // Đồng bộ 3 ô thông tin công trình về Form chính
  const dlgProj = document.getElementById("txtDlgProjectName")?.value.trim() || "";
  const dlgPack = document.getElementById("txtDlgBiddingPackage")?.value.trim() || "";
  const dlgLoc = document.getElementById("txtDlgProjectLocation")?.value.trim() || "";

  const txtProj = document.getElementById("txtProjectName");
  if (txtProj && dlgProj) txtProj.value = dlgProj;
  const txtPack = document.getElementById("txtBiddingPackage");
  if (txtPack && dlgPack) txtPack.value = dlgPack;
  const txtLoc = document.getElementById("txtProjectLocation");
  if (txtLoc && dlgLoc) txtLoc.value = dlgLoc;

  // Đóng dialog và thông báo
  const dlgOverlay = document.getElementById("dlgPriceScheduleOverlay");
  if (dlgOverlay) dlgOverlay.classList.add("hidden");

  setStatus(`Đã cập nhật giá trị: ${formatCurrency(grandTotal)} VNĐ vào Form hợp đồng.`);
  alert(`Đã cập nhật giá trị vào Form thành công:\n${formattedStr}`);
}

/**
 * Chèn bảng giá dự toán hoàn chỉnh vào tài liệu Microsoft Word
 */
async function insertPriceTableToWord() {
  if (!window.Word || !Office.context || !Office.context.document) {
    setStatus("Cảnh báo: Chỉ chạy được trong Microsoft Word Desktop.");
    alert("Chức năng chỉ hoạt động bên trong Microsoft Word!");
    return;
  }

  if (currentPriceItems.length === 0) {
    alert("Bảng giá hiện đang trống! Hãy thêm ít nhất 1 hạng mục.");
    return;
  }

  const projectName = document.getElementById("txtDlgProjectName")?.value.trim() || "";
  const biddingPackage = document.getElementById("txtDlgBiddingPackage")?.value.trim() || "";
  const projectLocation = document.getElementById("txtDlgProjectLocation")?.value.trim() || "";
  const cboType = document.getElementById("cboPriceTemplateType");
  const typeText = cboType ? cboType.options[cboType.selectedIndex].text : "Bảng giá hợp đồng";

  let subtitleParts = [typeText];
  if (projectName) subtitleParts.push(`Công trình: ${projectName}`);
  if (biddingPackage) subtitleParts.push(`Gói thầu: ${biddingPackage}`);
  if (projectLocation) subtitleParts.push(`Địa điểm: ${projectLocation}`);
  const subtitleStr = `(${subtitleParts.join(" - ")})`;

  let subtotal = 0;
  currentPriceItems.forEach(item => {
    subtotal += (item.qty || 0) * (item.price || 0);
  });

  const cboVat = document.getElementById("cboDlgVatRate");
  const vatRate = cboVat ? parseInt(cboVat.value, 10) : 10;
  const vatAmount = subtotal * (vatRate / 100);
  const grandTotal = subtotal + vatAmount;
  const inWords = convertNumberToVietnameseWords(grandTotal);

  setStatus("Đang chèn bảng giá vào Microsoft Word...");

  try {
    await Word.run(async (context) => {
      const wordDoc = context.document;
      const body = wordDoc.body;

      // Tiêu đề bảng
      const pTitle = body.insertParagraph("BẢNG TỔNG HỢP DỰ TOÁN / GIÁ TRỊ KÝ HỢP ĐỒNG", Word.InsertLocation.end);
      pTitle.font.name = "Times New Roman";
      pTitle.font.size = 13;
      pTitle.font.bold = true;
      pTitle.alignment = Word.Alignment.centered;
      pTitle.spaceBefore = 14;
      pTitle.spaceAfter = 2;

      const pSub = body.insertParagraph(subtitleStr, Word.InsertLocation.end);
      pSub.font.name = "Times New Roman";
      pSub.font.size = 12;
      pSub.font.italic = true;
      pSub.alignment = Word.Alignment.centered;
      pSub.spaceAfter = 8;

      // Tạo ma trận dữ liệu: Header row (1) + Item rows (N) + Summary rows (4)
      const rowsCount = 1 + currentPriceItems.length + 4;
      const colsCount = 6;
      const matrix = [];

      // Dòng 1: Header
      matrix.push(["STT", "Nội dung công việc / Hàng hóa vật tư", "ĐVT", "Khối lượng", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)"]);

      // Các dòng dữ liệu
      currentPriceItems.forEach((item, idx) => {
        const itemAmount = (item.qty || 0) * (item.price || 0);
        matrix.push([
          String(idx + 1),
          item.name || "",
          item.unit || "",
          String(item.qty || 0),
          formatCurrency(item.price || 0),
          formatCurrency(itemAmount)
        ]);
      });

      // Các dòng tổng kết
      matrix.push(["", "Cộng giá trị trước thuế:", "", "", "", formatCurrency(subtotal)]);
      matrix.push(["", `Thuế giá trị gia tăng GTGT (${vatRate}%):`, "", "", "", formatCurrency(vatAmount)]);
      matrix.push(["", "TỔNG CỘNG THANH TOÁN:", "", "", "", formatCurrency(grandTotal)]);
      matrix.push(["", `(Bằng chữ: ${inWords})`, "", "", "", ""]);

      // Chèn bảng
      const table = body.insertTable(rowsCount, colsCount, Word.InsertLocation.end, matrix);
      table.styleBuiltIn = "TableGrid";
      table.alignment = Word.Alignment.centered;

      // Định dạng độ rộng cột
      try {
        const columns = table.columns;
        columns.load("items");
        await context.sync();
        if (columns.items.length >= 6) {
          columns.items[0].setWidth(30, "None");  // STT
          columns.items[1].setWidth(185, "None"); // Tên
          columns.items[2].setWidth(40, "None");  // ĐVT
          columns.items[3].setWidth(55, "None");  // KL
          columns.items[4].setWidth(75, "None");  // Đơn giá
          columns.items[5].setWidth(85, "None");  // Thành tiền
        }
      } catch (e) {}

      // Định dạng font và căn lề cho dòng tiêu đề
      try {
        const rows = table.rows;
        rows.load("items");
        await context.sync();

        if (rows.items.length > 0) {
          // Row 0: Header
          rows.items[0].font.bold = true;
          rows.items[0].font.name = "Times New Roman";
          rows.items[0].font.size = 11;

          // Các dòng tổng kết cuối
          const rSub = rows.items[rows.items.length - 4];
          rSub.font.bold = true;

          const rVat = rows.items[rows.items.length - 3];
          rVat.font.bold = true;

          const rGrand = rows.items[rows.items.length - 2];
          rGrand.font.bold = true;
          rGrand.font.size = 12;

          const rWords = rows.items[rows.items.length - 1];
          rWords.font.italic = true;
        }
      } catch (e) {}

      const pSpacer = body.insertParagraph("", Word.InsertLocation.end);
      pSpacer.spaceAfter = 12;

      await context.sync();
      setStatus(`Đã chèn bảng giá (${currentPriceItems.length} hạng mục) vào Microsoft Word thành công.`);
      alert("Đã chèn bảng giá hợp đồng vào tài liệu Word thành công!");
    });
  } catch (error) {
    console.error("Lỗi chèn bảng giá vào Word:", error);
    setStatus(`Lỗi chèn bảng: ${error.message}`);
    alert(`Lỗi Word API: ${error.message}`);
  }
}

