/**
 * ocr-dialog.js - Xử lý Bóc tách OCR AI Qwen2.5-VL-3B Local & Chuyển dữ liệu sang Word
 * Kết nối song song với Word Taskpane Add-in qua BroadcastChannel & Office.js
 */

const SYSTEM_PROMPT_OCR = `Bạn là hệ thống OCR quang học chuyên nghiệp, trích xuất văn bản và bảng biểu tiếng Việt chính xác 100% nguyên văn (High-Precision Verbatim OCR).

NHIỆM VỤ TỐI THƯỢNG:
Chuyển đổi CHÍNH XÁC NGUYÊN VĂN 100% toàn bộ nội dung trong hình ảnh thành văn bản, giữ nguyên từng câu, từng chữ, từng con số từ trên xuống dưới.

QUY TẮC BẮT BUỘC:
1. TRÍCH XUẤT ĐẦY ĐỦ 100% NGUYÊN VĂN ĐẾN HẾT TRANG (KHÔNG TÓM TẮT, KHÔNG RÚT GỌN):
   - Đọc và ghi lại toàn bộ nội dung văn bản theo thứ tự từ trên xuống dưới, không được dừng dở giữa chừng:
     + Quốc hiệu, Tiêu ngữ (CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM / Độc lập - Tự do - Hạnh phúc).
     + Tên cơ quan, đơn vị ban hành văn bản.
     + Số và ký hiệu văn bản, địa danh, ngày tháng năm.
     + Tiêu đề văn bản, phần "Về việc...".
     + Phần "Kính gửi: ..." (ghi đầy đủ tất cả các cơ quan, đơn vị, công ty được kính gửi).
     + Toàn bộ các đoạn "Căn cứ...": Căn cứ số..., ngày..., của... (ghi đầy đủ nguyên văn từng căn cứ, không cắt ngắn bất kỳ chữ nào).
     + Toàn bộ nội dung lời dẫn, các mục được đánh số (1., 2., 3., a), b), Điều 1, Điều 2...).
     + Toàn bộ các số tiền bằng số và bằng chữ (ví dụ: 2.312.614.000 đồng - Bằng chữ: Hai tỷ ba trăm mười hai triệu sáu trăm mười bốn nghìn đồng).
     + Nơi nhận: ghi đầy đủ danh sách nơi nhận ở cuối văn bản (nếu có).
     + Chức vụ, họ tên người ký, ghi chú về con dấu (ví dụ: GIÁM ĐỐC / Nguyễn Hữu Trí [Đã ký và đóng dấu]).

2. BẢNG BIỂU SỐ LIỆU (NẾU CÓ):
   - Chuyển thành BẢNG MARKDOWN chuẩn (| Cột 1 | Cột 2 | Cột 3 |).
   - Bắt buộc trích xuất 100% TẤT CẢ CÁC HÀNG VÀ CÁC CỘT, không bỏ sót dòng nào.

3. TUYỆT ĐỐI CẤM:
   - TUYỆT ĐỐI KHÔNG DỪNG GIỮA CHỪNG KHI CHƯA HẾT TRANG. Phải quét trọn vẹn đến dòng chữ cuối cùng của ảnh.
   - TUYỆT ĐỐI KHÔNG tóm tắt nội dung.
   - TUYỆT ĐỐI KHÔNG tự ý cắt ngắn gọn các đoạn văn bản.
   - TUYỆT ĐỐI KHÔNG gom nhóm thành các mục tóm tắt như "[THÔNG TIN CHUNG]" hay "[CĂN CỨ]" nếu tài liệu gốc không có chữ đó.
   - Giữ nguyên cấu trúc dòng và dấu câu của văn bản gốc.`;

// Cấu hình PDF.js Offline
if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
}

// Khởi tạo kênh liên lạc nội bộ với Taskpane
const ocrChannel = new BroadcastChannel('WordAddin_OCR_Channel');

// State quản lý toàn bộ Dialog
const state = {
  file: null,
  pdfDoc: null,
  currentPage: 1,
  totalPages: 1,
  currentImageBase64: null,
  zoom: 1.0,
  isExtracting: false,
  rawResultText: '',
  selectedModel: 'qwen3b',
  isServerOnline: false
};

// DOM Elements
const btnSelectFile = document.getElementById('btnSelectFile');
const fileInput = document.getElementById('fileInput');
const btnStartOcr = document.getElementById('btnStartOcr');
const cboExtractScope = document.getElementById('cboExtractScope');
const btnTransferToWord = document.getElementById('btnTransferToWord');
const btnCopyResult = document.getElementById('btnCopyResult');
const btnToggleAi = document.getElementById('btnToggleAi');
const btnReload = document.getElementById('btnReload');
const cboModelSelect = document.getElementById('cboModelSelect');
const aiStatusBadge = document.getElementById('aiStatusBadge');
const aiStatusDot = document.getElementById('aiStatusDot');
const aiStatusText = document.getElementById('aiStatusText');

const dropzone = document.getElementById('dropzone');
const viewerContainer = document.getElementById('viewerContainer');
const viewerScrollBox = document.getElementById('viewerScrollBox');
const pdfPagination = document.getElementById('pdfPagination');
const btnPrevPage = document.getElementById('btnPrevPage');
const btnNextPage = document.getElementById('btnNextPage');
const lblPageIndicator = document.getElementById('lblPageIndicator');
const lblFileMeta = document.getElementById('lblFileMeta');
const btnZoomIn = document.getElementById('btnZoomIn');
const btnZoomOut = document.getElementById('btnZoomOut');
const btnZoomReset = document.getElementById('btnZoomReset');
const lblZoomLevel = document.getElementById('lblZoomLevel');
const btnRemoveFile = document.getElementById('btnRemoveFile');
const pdfCanvas = document.getElementById('pdfCanvas');
const imgPreview = document.getElementById('imgPreview');

const logConsole = document.getElementById('logConsole');
const lblResultStatus = document.getElementById('lblResultStatus');
const tabVisual = document.getElementById('tabVisual');
const tabRaw = document.getElementById('tabRaw');
const viewVisual = document.getElementById('viewVisual');
const viewRaw = document.getElementById('viewRaw');
const txtRawResult = document.getElementById('txtRawResult');
const resultEmptyState = document.getElementById('resultEmptyState');

const statusMessage = document.getElementById('statusMessage');
const modelInfo = document.getElementById('modelInfo');
const btnMinWindow = document.getElementById('btnMinWindow');
const btnCloseWindow = document.getElementById('btnCloseWindow');

/**
 * Ghi log ra màn hình console của dialog
 */
function appendLog(message, type = 'info') {
  const line = document.createElement('div');
  line.className = `log-line ${type}`;
  const now = new Date();
  const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
  line.textContent = `[${timeStr}] ${message}`;
  logConsole.appendChild(line);
  logConsole.scrollTop = logConsole.scrollHeight;
}

function setStatus(msg) {
  if (statusMessage) statusMessage.textContent = msg;
}

/**
 * Kiểm tra trạng thái AI Server
 */
async function checkAiStatus() {
  try {
    const res = await fetch('/api/ocr/status');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    state.isServerOnline = data.running;

    if (data.running) {
      aiStatusDot.className = 'status-dot online';
      aiStatusText.textContent = `Trực tuyến (${data.currentModel === 'qwen3b' ? '3B' : '7B'})`;
      if (modelInfo) modelInfo.textContent = `AI: ${data.modelLabel || 'Qwen 3B'}`;
      if (btnToggleAi) {
        btnToggleAi.className = 'win-btn win-btn-danger';
        btnToggleAi.textContent = '2. Tắt AI';
        btnToggleAi.title = 'Tắt tiến trình AI';
      }
    } else if (data.isStarting) {
      aiStatusDot.className = 'status-dot starting';
      aiStatusText.textContent = 'Đang nạp...';
      if (btnToggleAi) {
        btnToggleAi.className = 'win-btn';
        btnToggleAi.textContent = '2. Đang nạp...';
        btnToggleAi.title = 'Đang nạp mô hình vào RAM...';
      }
    } else {
      aiStatusDot.className = 'status-dot offline';
      aiStatusText.textContent = 'Chế độ chờ';
      if (btnToggleAi) {
        btnToggleAi.className = 'win-btn';
        btnToggleAi.textContent = '2. Bật AI';
        btnToggleAi.title = 'Bật tiến trình AI';
      }
    }

    return data;
  } catch (err) {
    aiStatusDot.className = 'status-dot offline';
    aiStatusText.textContent = 'Ngoại tuyến';
    return { running: false };
  }
}

/**
 * Khởi động AI Server theo model đã chọn
 */
async function startAiServer(modelKey) {
  appendLog(`Đang khởi chạy mô hình AI [${modelKey}]...`, 'info');
  setStatus('Đang nạp mô hình AI vào bộ nhớ...');
  aiStatusDot.className = 'status-dot starting';
  aiStatusText.textContent = 'Đang nạp mô hình...';

  try {
    const res = await fetch('/api/ocr/start', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: modelKey })
    });
    const data = await res.json();
    if (data.running || data.status === 'ok') {
      appendLog(`Mô hình ${data.label || modelKey} đã sẵn sàng 100%!`, 'success');
      state.isServerOnline = true;
      checkAiStatus();
      return true;
    }

    // Nếu server chưa nạp xong, tiếp tục đợi và thăm dò /api/ocr/status
    appendLog('Đang nạp mô hình vào RAM/VRAM, vui lòng đợi trong giây lát...', 'warn');
    for (let i = 0; i < 30; i++) {
      await new Promise(r => setTimeout(r, 2000));
      const st = await checkAiStatus();
      if (st.running) {
        appendLog(`Mô hình ${st.modelLabel || modelKey} đã sẵn sàng 100%!`, 'success');
        state.isServerOnline = true;
        return true;
      }
      appendLog(`... đang nạp mô hình (${(i + 1) * 2}s)`, 'info');
    }

    appendLog('Thời gian nạp mô hình lâu hơn dự kiến. Hãy thử lại.', 'error');
    return false;
  } catch (err) {
    appendLog(`Lỗi khởi động AI Server: ${err.message}`, 'error');
    return false;
  }
}

/**
 * Tắt AI Server để giải phóng tài nguyên
 */
async function stopAiServer() {
  appendLog('Đang tắt AI Server để giải phóng RAM/VRAM...', 'warn');
  try {
    await fetch('/api/ocr/stop', { method: 'POST' });
    state.isServerOnline = false;
    appendLog('Đã tắt hoàn toàn tiến trình AI. Bộ nhớ RAM đã được giải phóng.', 'success');
    setStatus('AI đã dừng (Tiết kiệm RAM)');
    checkAiStatus();
  } catch (err) {
    appendLog(`Lỗi tắt server: ${err.message}`, 'error');
  }
}

/**
 * Xử lý khi nạp file vào ứng dụng
 */
async function handleFileLoaded(file) {
  if (!file) return;
  state.file = file;
  lblFileMeta.textContent = `${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
  appendLog(`Đã nạp file: ${file.name}`, 'info');

  dropzone.classList.add('hidden');
  viewerContainer.classList.remove('hidden');
  btnStartOcr.disabled = false;
  setStatus(`Đã nạp file: ${file.name} | Sẵn sàng trích xuất`);

  if (file.type === 'application/pdf') {
    pdfPagination.classList.remove('hidden');
    imgPreview.classList.add('hidden');
    pdfCanvas.classList.remove('hidden');
    await loadPdfDocument(file);
  } else {
    pdfPagination.classList.add('hidden');
    pdfCanvas.classList.add('hidden');
    imgPreview.classList.remove('hidden');
    loadImageDocument(file);
  }
}

/**
 * Nạp và hiển thị file ảnh
 */
function loadImageDocument(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    state.currentImageBase64 = e.target.result;
    imgPreview.src = state.currentImageBase64;
    state.currentPage = 1;
    state.totalPages = 1;
    state.zoom = 1.0;
    applyZoom();
    if (viewerScrollBox) {
      viewerScrollBox.scrollTop = 0;
      viewerScrollBox.scrollLeft = 0;
    }
    appendLog('Đã hiển thị bản xem trước ảnh tài liệu.', 'info');
  };
  reader.readAsDataURL(file);
}

/**
 * Nạp và hiển thị file PDF
 */
async function loadPdfDocument(file) {
  try {
    appendLog('Đang đọc cấu trúc file PDF...', 'info');
    const arrayBuffer = await file.arrayBuffer();
    const loadingTask = window.pdfjsLib.getDocument({ data: arrayBuffer });
    state.pdfDoc = await loadingTask.promise;
    state.totalPages = state.pdfDoc.numPages;
    state.currentPage = 1;
    state.zoom = 1.0;
    applyZoom();
    appendLog(`Tài liệu PDF gồm ${state.totalPages} trang.`, 'info');
    updatePdfPageUi();
    await renderPdfPage(1);
  } catch (err) {
    appendLog(`Lỗi nạp file PDF: ${err.message}`, 'error');
    alert(`Không thể đọc file PDF: ${err.message}`);
  }
}

/**
 * Render 1 trang PDF ra Canvas với độ phân giải chuẩn xác
 */
async function renderPdfPage(pageNum) {
  if (!state.pdfDoc) return;
  try {
    const page = await state.pdfDoc.getPage(pageNum);
    const unscaled = page.getViewport({ scale: 1.0 });
    const maxDim = Math.max(unscaled.width, unscaled.height);
    let targetScale = 1600 / maxDim;
    if (targetScale < 1.4) targetScale = 1.4;
    if (targetScale > 2.2) targetScale = 2.2;

    const viewport = page.getViewport({ scale: targetScale });
    pdfCanvas.height = viewport.height;
    pdfCanvas.width = viewport.width;

    const ctx = pdfCanvas.getContext('2d');
    ctx.clearRect(0, 0, pdfCanvas.width, pdfCanvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;

    state.currentImageBase64 = pdfCanvas.toDataURL('image/jpeg', 0.88);

    // Luôn reset cuộn chuột về đỉnh trên cùng để không bị mất lề/cạnh trên
    if (viewerScrollBox) {
      viewerScrollBox.scrollTop = 0;
      viewerScrollBox.scrollLeft = 0;
    }
    appendLog(`Đã render trang ${pageNum}/${state.totalPages} chuẩn bị bóc tách.`, 'info');
  } catch (err) {
    appendLog(`Lỗi render trang ${pageNum}: ${err.message}`, 'error');
  }
}

function updatePdfPageUi() {
  lblPageIndicator.textContent = `Trang ${state.currentPage}/${state.totalPages}`;
  btnPrevPage.disabled = state.currentPage <= 1;
  btnNextPage.disabled = state.currentPage >= state.totalPages;
}

/**
 * Cập nhật mức thu phóng hiển thị ảnh/PDF
 */
function applyZoom() {
  if (lblZoomLevel) {
    lblZoomLevel.textContent = `${Math.round(state.zoom * 100)}%`;
  }
  if (pdfCanvas) {
    pdfCanvas.style.transform = 'none';
    if (state.zoom === 1.0) {
      pdfCanvas.style.width = '';
      pdfCanvas.style.maxWidth = '100%';
    } else {
      const zoomPct = Math.round(state.zoom * 100);
      pdfCanvas.style.width = `${zoomPct}%`;
      pdfCanvas.style.maxWidth = 'none';
    }
  }
  if (imgPreview) {
    imgPreview.style.transform = 'none';
    if (state.zoom === 1.0) {
      imgPreview.style.width = '';
      imgPreview.style.maxWidth = '100%';
    } else {
      const zoomPct = Math.round(state.zoom * 100);
      imgPreview.style.width = `${zoomPct}%`;
      imgPreview.style.maxWidth = 'none';
    }
  }
}

/**
 * Trích xuất nguyên văn 1 ảnh/trang bằng AI Qwen2.5-VL qua SSE Stream
 */
async function extractSinglePageImage(imageBase64, pageTitle, isLastPage = false) {
  let userPrompt = `Đọc và bóc tách TOÀN BỘ văn bản từ trên xuống dưới trong trang tài liệu này.`;
  if (pageTitle) {
    userPrompt += ` (${pageTitle}).`;
  }
  userPrompt += `\n- BẮT BUỘC ĐỌC VÀ TRÍCH XUẤT 100% NGUYÊN VĂN TỪ ĐẦU TRANG ĐẾN DÒNG CUỐI CÙNG CỦA TRANG.`;
  userPrompt += `\n- TUYỆT ĐỐI KHÔNG DỪNG GIỮA CHỪNG. Phải đọc trọn vẹn từng câu chữ, từng căn cứ pháp lý, các điều khoản, các mục 1, 2, 3, số tiền bằng số và bằng chữ.`;
  if (isLastPage) {
    userPrompt += `\n- ĐẶC BIỆT: Đây là trang cuối, hãy trích xuất đầy đủ danh sách Nơi nhận, chức vụ, họ tên người ký và tình trạng con dấu ở cuối trang.`;
  }
  userPrompt += `\n- Nếu có bảng biểu, hãy chuyển thành bảng Markdown hoàn chỉnh tất cả các hàng và các cột.`;

  const payload = {
    messages: [
      { role: 'system', content: SYSTEM_PROMPT_OCR },
      {
        role: 'user',
        content: [
          { type: 'text', text: userPrompt },
          { type: 'image_url', image_url: { url: imageBase64 } }
        ]
      }
    ],
    temperature: 0.1,
    max_tokens: 4096,
    stream: true
  };

  let response = null;
  for (let attempt = 0; attempt < 12; attempt++) {
    response = await fetch('/api/ocr/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (response.status === 503) {
      appendLog(`AI đang hoàn tất nạp trọng số mô hình (503). Đang đợi và thử lại (${attempt + 1}/12)...`, 'warn');
      await new Promise(r => setTimeout(r, 2000));
      continue;
    }
    break;
  }

  if (!response || !response.ok) {
    throw new Error(`Server báo lỗi HTTP ${response ? response.status : 'không phản hồi'}`);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let buffer = '';
  let pageContent = '';

  if (pageTitle) {
    const pageHeader = (state.rawResultText ? '\n\n' : '') + `### --- ${pageTitle} ---\n\n`;
    state.rawResultText += pageHeader;
    txtRawResult.value = state.rawResultText;
  }

  let lastRenderTime = Date.now();

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop();

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith(':')) continue;
      if (trimmed === 'data: [DONE]') continue;

      if (trimmed.startsWith('data: ')) {
        try {
          const json = JSON.parse(trimmed.slice(6));
          const delta = json.choices?.[0]?.delta?.content || '';
          if (delta) {
            pageContent += delta;
            state.rawResultText += delta;
            txtRawResult.value = state.rawResultText;
            txtRawResult.scrollTop = txtRawResult.scrollHeight;

            // Cập nhật tab trực quan định kỳ mỗi 500ms
            if (Date.now() - lastRenderTime > 500) {
              renderFormattedResult(state.rawResultText);
              lastRenderTime = Date.now();
            }
          }
        } catch (_) {}
      }
    }
  }

  return pageContent;
}

/**
 * Thực hiện trích xuất dữ liệu bằng AI Qwen2.5-VL (Hỗ trợ đa trang PDF tuần tự)
 */
async function performOcrExtraction() {
  if (!state.currentImageBase64) {
    alert('Vui lòng chọn hoặc nạp một file tài liệu trước khi bấm trích xuất.');
    return;
  }

  state.isExtracting = true;
  btnStartOcr.disabled = true;
  btnSelectFile.disabled = true;
  btnTransferToWord.disabled = true;
  btnCopyResult.disabled = true;

  lblResultStatus.textContent = 'Đang trích xuất...';
  setStatus('Đang kết nối AI Qwen2.5-VL...');
  appendLog('=== BẮT ĐẦU QUÁ TRÌNH BÓC TÁCH NGUYÊN VĂN 100% BẰNG AI ===', 'info');

  // Kiểm tra và khởi động server nếu chưa chạy
  const status = await checkAiStatus();
  if (!status.running) {
    appendLog(`AI Server chưa bật. Tiến hành nạp mô hình [${state.selectedModel}]...`, 'warn');
    const started = await startAiServer(state.selectedModel);
    if (!started) {
      appendLog('Không thể khởi chạy AI Server. Vui lòng kiểm tra lại.', 'error');
      state.isExtracting = false;
      btnStartOcr.disabled = false;
      btnSelectFile.disabled = false;
      return;
    }
  }

  // Chuẩn bị giao diện kết quả
  resultEmptyState.classList.add('hidden');
  viewVisual.classList.remove('hidden');
  viewVisual.innerHTML = '<div style="color: #0078D7; font-weight: 600; padding: 10px;">⏳ Mô hình AI đang đọc và bóc tách nguyên văn toàn bộ nội dung từng token...</div>';
  txtRawResult.value = '';
  state.rawResultText = '';

  const scope = (cboExtractScope && cboExtractScope.value) ? cboExtractScope.value : 'all';

  try {
    if (state.pdfDoc && state.totalPages > 1 && scope === 'all') {
      appendLog(`Tiến hành bóc tách toàn bộ ${state.totalPages} trang của tài liệu PDF...`, 'info');
      for (let p = 1; p <= state.totalPages; p++) {
        state.currentPage = p;
        updatePdfPageUi();
        setStatus(`Đang kết xuất và bóc tách Trang ${p}/${state.totalPages}...`);
        appendLog(`--- Bắt đầu xử lý Trang ${p}/${state.totalPages} ---`, 'info');

        await renderPdfPage(p);
        await extractSinglePageImage(state.currentImageBase64, `TRANG ${p} / ${state.totalPages}`, p === state.totalPages);
        renderFormattedResult(state.rawResultText);
        appendLog(`Đã hoàn tất trích xuất Trang ${p}/${state.totalPages}.`, 'success');
      }
    } else {
      const pageTitle = (state.pdfDoc && state.totalPages > 1) ? `TRANG ${state.currentPage} / ${state.totalPages}` : '';
      setStatus('Đang bóc tách dữ liệu nguyên văn...');
      appendLog(`Đang gửi dữ liệu ảnh sang AI Qwen2.5-VL (${state.selectedModel.toUpperCase()})...`, 'info');
      await extractSinglePageImage(state.currentImageBase64, pageTitle, true);
      renderFormattedResult(state.rawResultText);
    }

    appendLog('=== BÓC TÁCH HOÀN TẤT THÀNH CÔNG 100% NGUYÊN VĂN ===', 'success');
    lblResultStatus.textContent = 'Hoàn tất';
    setStatus('Trích xuất thành công 100%! Bạn có thể xem kết quả hoặc bấm "Chuyển sang Word".');
    renderFormattedResult(state.rawResultText);

    btnTransferToWord.disabled = false;
    btnCopyResult.disabled = false;
  } catch (err) {
    appendLog(`Lỗi trong quá trình trích xuất: ${err.message}`, 'error');
    lblResultStatus.textContent = 'Lỗi trích xuất';
    setStatus(`Lỗi: ${err.message}`);
    alert(`Lỗi bóc tách OCR: ${err.message}`);
  } finally {
    state.isExtracting = false;
    btnStartOcr.disabled = false;
    btnSelectFile.disabled = false;
  }
}

/**
 * Chuyển đổi chuỗi Markdown kết quả thành HTML trực quan (bảng, tiêu đề, thẻ)
 */
function renderFormattedResult(rawMarkdown) {
  if (!rawMarkdown) return;
  const lines = rawMarkdown.split('\n');
  let html = '';
  let inTable = false;
  let tableRows = [];

  function flushTable() {
    if (tableRows.length === 0) return '';
    let tblHtml = '<table class="ocr-markdown-table"><thead>';
    // Hàng đầu làm header
    const headers = tableRows[0];
    tblHtml += '<tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr></thead><tbody>';

    for (let i = 1; i < tableRows.length; i++) {
      tblHtml += '<tr>' + tableRows[i].map(c => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>';
    }
    tblHtml += '</tbody></table>';
    tableRows = [];
    return tblHtml;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Dòng bảng Markdown: | Cột 1 | Cột 2 |
    if (line.startsWith('|') && line.endsWith('|')) {
      // Bỏ qua dòng phân cách |---|---|
      if (/^\|[\s\-:]+(\|[\s\-:]+)+\|$/.test(line)) {
        continue;
      }
      const cells = line.split('|').slice(1, -1).map(c => c.trim());
      inTable = true;
      tableRows.push(cells);
      continue;
    } else {
      if (inTable) {
        html += flushTable();
        inTable = false;
      }
    }

    if (!line) {
      continue;
    }

    if (line.startsWith('### --- ') && line.endsWith(' ---')) {
      const title = line.replace(/^###\s*---\s*|\s*---\s*$/g, '');
      html += `<div style="background-color: #E8EEF4; color: #003366; font-weight: bold; font-size: 13px; padding: 6px 12px; margin: 18px 0 10px 0; border-left: 4px solid #0078D7; border-radius: 2px;">📑 ${escapeHtml(title)}</div>`;
    } else if (line.startsWith('### ')) {
      html += `<h3>${escapeHtml(line.slice(4))}</h3>`;
    } else if (line.startsWith('## ')) {
      html += `<h2>${escapeHtml(line.slice(3))}</h2>`;
    } else if (line.startsWith('# ')) {
      html += `<h1>${escapeHtml(line.slice(2))}</h1>`;
    } else if (line.startsWith('- ')) {
      html += `<div class="ocr-meta-row" style="margin-left: 12px;">• ${escapeHtml(line.slice(2))}</div>`;
    } else {
      html += `<p style="margin: 4px 0; line-height: 1.5; font-size: 13px; text-align: justify;">${escapeHtml(line)}</p>`;
    }
  }

  if (inTable) {
    html += flushTable();
  }

  viewVisual.innerHTML = html;
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Phân tích Markdown tables và text để chèn sang Word
 */
function parseMarkdownForWord(markdown) {
  const lines = markdown.split('\n');
  const blocks = [];
  let currentTable = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Dấu phân tách trang
    if (line.startsWith('### --- TRANG') || line.startsWith('## --- TRANG') || line.startsWith('--- TRANG')) {
      if (currentTable) {
        blocks.push({ type: 'table', data: currentTable });
        currentTable = null;
      }
      blocks.push({ type: 'page_break', text: line.replace(/^#+\s*---\s*|\s*---\s*$/g, '') });
      continue;
    }

    if (line.startsWith('|') && line.endsWith('|')) {
      if (/^\|[\s\-:]+(\|[\s\-:]+)+\|$/.test(line)) {
        continue;
      }
      const cells = line.split('|').slice(1, -1).map(c => c.trim());
      if (!currentTable) {
        currentTable = [];
      }
      currentTable.push(cells);
    } else {
      if (currentTable) {
        blocks.push({ type: 'table', data: currentTable });
        currentTable = null;
      }
      if (line) {
        if (line.startsWith('#')) {
          blocks.push({ type: 'heading', text: line.replace(/^#+\s*/, '') });
        } else {
          blocks.push({ type: 'paragraph', text: line });
        }
      }
    }
  }

  if (currentTable) {
    blocks.push({ type: 'table', data: currentTable });
  }

  return blocks;
}

/**
 * Thực thi chèn danh sách blocks vào Word
 */
async function insertBlocksIntoWord(wordApp, blocks) {
  await wordApp.run(async (context) => {
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
        p.spaceAfter = 3;
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
  });
}

/**
 * Thực thi chuyển dữ liệu OCR sang Word (Chỉ gọi 1 kênh duy nhất để tránh chèn trùng lặp)
 */
async function transferResultToWord() {
  if (!state.rawResultText) {
    alert('Chưa có dữ liệu trích xuất để chuyển sang Word.');
    return;
  }

  appendLog('Đang chuyển dữ liệu và định dạng Word tự động...', 'info');
  setStatus('Đang chèn dữ liệu vào Microsoft Word...');

  const blocks = parseMarkdownForWord(state.rawResultText);

  // 1. Nếu mở từ Word Taskpane qua window.opener
  if (window.opener && window.opener.Word && window.opener.Word.run) {
    try {
      await insertBlocksIntoWord(window.opener.Word, blocks);
      appendLog('Đã chèn toàn bộ dữ liệu vào Word thành công!', 'success');
      setStatus('Đã chuyển sang Word thành công!');
      alert('Đã chuyển dữ liệu OCR vào Word thành công!');
      return;
    } catch (err) {
      console.warn('Lỗi chèn qua window.opener:', err);
    }
  }

  // 2. Nếu cửa sổ này có sẵn quyền Office trực tiếp
  if (typeof Word !== 'undefined' && Word.run) {
    try {
      await insertBlocksIntoWord(Word, blocks);
      appendLog('Đã chèn toàn bộ dữ liệu vào Word thành công!', 'success');
      setStatus('Đã chuyển sang Word thành công!');
      alert('Đã chuyển dữ liệu OCR vào Word thành công!');
      return;
    } catch (err) {
      console.warn('Lỗi chèn trực tiếp:', err);
    }
  }

  // 3. Fallback: Phát qua BroadcastChannel nếu không gọi trực tiếp được
  ocrChannel.postMessage({
    type: 'INSERT_OCR_BLOCKS',
    blocks: blocks,
    rawText: state.rawResultText
  });

  appendLog('Đã phát tín hiệu sang Taskpane Add-in để chèn vào Word.', 'success');
  setStatus('Đã gửi lệnh chèn dữ liệu sang Word!');
  alert('Đã gửi dữ liệu sang Word! Hãy kiểm tra tài liệu Word đang mở.');
}

/**
 * Đăng ký sự kiện giao diện
 */
function initEvents() {
  // Chọn file bằng nút
  btnSelectFile.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFileLoaded(e.target.files[0]);
    }
  });

  // Dropzone kéo thả
  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileLoaded(e.dataTransfer.files[0]);
    }
  });

  // Dán ảnh từ clipboard (Ctrl + V)
  window.addEventListener('paste', (e) => {
    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    for (const item of items) {
      if (item.type.indexOf('image') !== -1) {
        const blob = item.getAsFile();
        handleFileLoaded(blob);
        appendLog('Đã nhận hình ảnh dán từ Clipboard (Ctrl + V).', 'info');
        break;
      }
    }
  });

  // Điều hướng trang PDF
  btnPrevPage.addEventListener('click', async () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      updatePdfPageUi();
      await renderPdfPage(state.currentPage);
    }
  });

  btnNextPage.addEventListener('click', async () => {
    if (state.currentPage < state.totalPages) {
      state.currentPage++;
      updatePdfPageUi();
      await renderPdfPage(state.currentPage);
    }
  });

  // Thu phóng ảnh/PDF
  btnZoomIn.addEventListener('click', () => {
    state.zoom += 0.15;
    applyZoom();
  });
  btnZoomOut.addEventListener('click', () => {
    if (state.zoom > 0.3) {
      state.zoom -= 0.15;
      applyZoom();
    }
  });
  btnZoomReset.addEventListener('click', () => {
    state.zoom = 1.0;
    applyZoom();
    if (viewerScrollBox) {
      viewerScrollBox.scrollTop = 0;
      viewerScrollBox.scrollLeft = 0;
    }
  });

  // Đổi file
  btnRemoveFile.addEventListener('click', () => {
    state.file = null;
    state.pdfDoc = null;
    state.currentImageBase64 = null;
    fileInput.value = '';
    dropzone.classList.remove('hidden');
    viewerContainer.classList.add('hidden');
    btnStartOcr.disabled = true;
    lblFileMeta.textContent = 'Chưa có file';
    appendLog('Đã gỡ file, sẵn sàng chọn file mới.', 'info');
  });

  // Chuyển tab Trực quan / Raw
  tabVisual.addEventListener('click', () => {
    tabVisual.classList.add('active');
    tabRaw.classList.remove('active');
    viewVisual.classList.remove('hidden');
    viewRaw.classList.add('hidden');
  });

  tabRaw.addEventListener('click', () => {
    tabRaw.classList.add('active');
    tabVisual.classList.remove('active');
    viewRaw.classList.remove('hidden');
    viewVisual.classList.add('hidden');
  });

  // Chọn mô hình
  cboModelSelect.addEventListener('change', () => {
    state.selectedModel = cboModelSelect.value;
    appendLog(`Đã chọn mô hình: ${state.selectedModel.toUpperCase()}`, 'info');
  });

  // Nút bắt đầu OCR
  btnStartOcr.addEventListener('click', performOcrExtraction);

  // Nút chuyển sang Word
  btnTransferToWord.addEventListener('click', transferResultToWord);

  // Nút sao chép
  btnCopyResult.addEventListener('click', () => {
    if (state.rawResultText) {
      navigator.clipboard.writeText(state.rawResultText).then(() => {
        appendLog('Đã sao chép toàn bộ kết quả vào bộ nhớ tạm (Clipboard).', 'success');
        alert('Đã sao chép kết quả vào bộ nhớ tạm!');
      });
    }
  });

  // Nút Bật / Tắt AI
  if (btnToggleAi) {
    btnToggleAi.addEventListener('click', async () => {
      if (state.isServerOnline) {
        await stopAiServer();
      } else {
        await startAiServer(cboModelSelect.value);
      }
    });
  }

  // Nhấn vào Huy hiệu trạng thái cũng hỗ trợ Bật / Tắt AI nhanh
  if (aiStatusBadge) {
    aiStatusBadge.addEventListener('click', async () => {
      if (state.isServerOnline) {
        await stopAiServer();
      } else {
        await startAiServer(cboModelSelect.value);
      }
    });
  }

  // Nút nạp lại
  btnReload.addEventListener('click', () => window.location.reload());

  // Nút trợ giúp
  const btnOcrHelp = document.getElementById('btnOcrHelp');
  if (btnOcrHelp) {
    btnOcrHelp.addEventListener('click', () => {
      alert(
        "HƯỚNG DẪN TÍNH NĂNG CÁC NÚT TRÊN CỬA SỔ BÓC TÁCH OCR AI:\n\n" +
        "1. Nút [1. Chọn file]:\n" +
        "   - Mở hộp thoại để chọn tệp tin PDF, hình ảnh (PNG, JPG, WEBP) tài liệu, hợp đồng cần trích xuất.\n\n" +
        "2. Nút [2. Bật AI / Tắt AI]:\n" +
        "   - Bật hoặc Tắt tiến trình AI Local (Qwen2.5-VL-3B) trên máy tính để tiết kiệm tài nguyên RAM/VRAM khi không dùng.\n\n" +
        "3. Nút [3. Trích xuất]:\n" +
        "   - AI tiến hành đọc và quét tài liệu, trích xuất chính xác văn bản, chữ ký, bảng biểu sang định dạng chuẩn.\n\n" +
        "4. Nút [4. Chèn vào Word]:\n" +
        "   - Tự động chuyển các khối văn bản và bảng biểu đã bóc tách sang tài liệu Microsoft Word đang mở.\n\n" +
        "5. Nút [5. Sao chép]:\n" +
        "   - Sao chép toàn bộ kết quả bóc tách dạng Markdown/Text vào Clipboard.\n\n" +
        "6. Nút [6. Làm mới]:\n" +
        "   - Nạp lại giao diện và xóa dữ liệu phiên làm việc hiện tại.\n\n" +
        "7. Nút [7. Trợ giúp]:\n" +
        "   - Mở bảng hướng dẫn tính năng chi tiết này."
      );
    });
  }

  // Nút đóng cửa sổ
  btnCloseWindow.addEventListener('click', () => window.close());
  btnMinWindow.addEventListener('click', () => {
    appendLog('Đã thu nhỏ (hoặc nhấn Alt+Tab để chuyển cửa sổ).', 'info');
  });
}

// Khởi chạy khi tài liệu nạp xong
document.addEventListener('DOMContentLoaded', () => {
  initEvents();
  checkAiStatus();
  setInterval(checkAiStatus, 5000);
});
