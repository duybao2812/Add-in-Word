/**
 * law-search-dialog.js
 * Logic giao diện điều khiển Cửa sổ Tra cứu Văn bản Pháp luật (Nút chức năng số 6)
 */

(function () {
  'use strict';

  let currentSelectedId = null;
  let broadcastChannel = null;

  // Khởi tạo BroadcastChannel để giao tiếp với Taskpane Add-in
  try {
    broadcastChannel = new BroadcastChannel('WordAddin_Law_Channel');
  } catch (e) {
    console.warn('BroadcastChannel không khả dụng:', e);
  }

  document.addEventListener('DOMContentLoaded', () => {
    initUI();
    renderDocumentList();
    updateStatusStrip();
  });

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => {
      console.log('Office.js sẵn sàng trong Cửa sổ Tra cứu Pháp luật.');
    });
  }

  /**
   * Khởi tạo các sự kiện giao diện
   */
  function initUI() {
    const txtSearch = document.getElementById('txtSearchKeyword');
    const btnClearSearch = document.getElementById('btnClearSearch');
    const cboCategory = document.getElementById('cboCategoryFilter');
    const cboStatus = document.getElementById('cboStatusFilter');
    const btnSyncCloud = document.getElementById('btnSyncCloud');
    const btnConfigApi = document.getElementById('btnConfigApi');
    const btnReload = document.getElementById('btnReloadData');

    // Tìm kiếm tức thì khi gõ
    let debounceTimer = null;
    txtSearch.addEventListener('input', () => {
      btnClearSearch.style.display = txtSearch.value ? 'block' : 'none';
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        renderDocumentList();
      }, 250);
    });

    btnClearSearch.addEventListener('click', () => {
      txtSearch.value = '';
      btnClearSearch.style.display = 'none';
      renderDocumentList();
      txtSearch.focus();
    });

    // Sự kiện lọc
    cboCategory.addEventListener('change', () => renderDocumentList());
    cboStatus.addEventListener('change', () => renderDocumentList());

    // Nạp lại
    btnReload.addEventListener('click', () => {
      window.LawSyncEngine.initLocalData();
      renderDocumentList();
      updateStatusStrip();
    });

    // Đồng bộ đám mây
    btnSyncCloud.addEventListener('click', handleSyncCloud);

    // Cấu hình API Modal
    btnConfigApi.addEventListener('click', openConfigModal);
    document.getElementById('btnModalClose').addEventListener('click', closeConfigModal);
    document.getElementById('btnCancelConfig').addEventListener('click', closeConfigModal);
    document.getElementById('btnSaveConfig').addEventListener('click', saveConfigModal);
    document.getElementById('btnTestConnection').addEventListener('click', testSupabaseConnection);

    // Nút hành động chi tiết
    document.getElementById('btnInsertCitation').addEventListener('click', handleInsertToWord);
    document.getElementById('btnCopyCitation').addEventListener('click', handleCopyCitation);
    document.getElementById('btnOpenSourceLink').addEventListener('click', handleOpenSourceLink);

    // Phím tắt ESC để đóng cửa sổ
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        const modal = document.getElementById('modalConfig');
        if (modal.classList.contains('show')) {
          closeConfigModal();
        } else {
          window.close();
        }
      }
    });
  }

  /**
   * Cập nhật thanh trạng thái dưới đáy
   */
  function updateStatusStrip() {
    const engine = window.LawSyncEngine;
    const lblRevision = document.getElementById('lblSyncRevision');
    const lblLastSync = document.getElementById('lblLastSync');
    const lblStatus = document.getElementById('lblConnectionStatus');
    const indicator = document.getElementById('statusIndicator');

    if (lblRevision) lblRevision.textContent = `Revision: ${engine.localRevision}`;

    if (engine.config && engine.config.supabaseUrl) {
      if (lblStatus) lblStatus.textContent = 'Đã kết nối Cloud (Supabase Active)';
      if (indicator) indicator.style.backgroundColor = '#4CD964';
      if (engine.config.lastSyncTime) {
        const d = new Date(engine.config.lastSyncTime);
        if (lblLastSync) lblLastSync.textContent = `Đồng bộ gần nhất: ${d.toLocaleTimeString('vi-VN')} ${d.toLocaleDateString('vi-VN')}`;
      } else {
        if (lblLastSync) lblLastSync.textContent = 'Đồng bộ gần nhất: Chưa đồng bộ';
      }
    } else {
      if (lblStatus) lblStatus.textContent = 'Chế độ Ngoại tuyến (Offline Ready)';
      if (indicator) indicator.style.backgroundColor = '#FFA500';
      if (lblLastSync) lblLastSync.textContent = 'Đồng bộ gần nhất: Dữ liệu cơ sở 515 văn bản';
    }
  }

  /**
   * Render danh sách văn bản sang cột trái
   */
  function renderDocumentList() {
    const txtSearch = document.getElementById('txtSearchKeyword');
    const cboCategory = document.getElementById('cboCategoryFilter');
    const cboStatus = document.getElementById('cboStatusFilter');
    const container = document.getElementById('docListContainer');
    const lblStats = document.getElementById('lblListStats');

    const keyword = txtSearch ? txtSearch.value : '';
    const category = cboCategory ? cboCategory.value : 'all';
    const status = cboStatus ? cboStatus.value : 'all';

    const results = window.LawSyncEngine.searchLaws(keyword, category, status);

    if (lblStats) {
      lblStats.textContent = `Hiển thị: ${results.length} văn bản`;
    }

    if (!results || results.length === 0) {
      container.innerHTML = `
        <div style="padding: 30px 15px; text-align: center; color: #777;">
          Không tìm thấy văn bản pháp luật nào phù hợp với bộ lọc.
        </div>
      `;
      return;
    }

    let html = '';
    results.forEach((item) => {
      const isSelected = item.id === currentSelectedId;
      let badgeHtml = '';

      if (item.isExpired || item.status === 'Hết hiệu lực') {
        badgeHtml = '<span class="badge badge-expired">Hết hiệu lực</span>';
      } else if (item.loaiPhapLy === 'Sửa đổi bổ sung' || item.loaiPhapLy === 'Văn bản sửa đổi') {
        badgeHtml = '<span class="badge badge-modified">Sửa đổi BS</span>';
      } else {
        badgeHtml = '<span class="badge badge-active">Còn hiệu lực</span>';
      }

      const textTag = item.co_toan_van_html 
        ? '<span class="badge badge-tag" title="Có toàn văn chữ số hóa">📄 Text</span>'
        : '<span class="badge badge-tag" title="Chỉ có file đính kèm gốc">📎 File</span>';

      html += `
        <div class="doc-item ${isSelected ? 'active' : ''}" data-id="${item.id}">
          <div class="doc-item-title-row">
            <span class="doc-item-code">${escapeHtml(item.code || item.name)}</span>
            <div style="display: flex; gap: 4px;">
              ${badgeHtml}
              ${textTag}
            </div>
          </div>
          <div class="doc-item-name">${escapeHtml(item.name)}</div>
          <div class="doc-item-meta">
            <span>📅 ${escapeHtml(item.issueDate || 'Năm ' + (item.year || ''))}</span>
            <span>•</span>
            <span>📂 ${escapeHtml(item.nhomChuDe || '')}</span>
          </div>
        </div>
      `;
    });

    container.innerHTML = html;

    // Gắn sự kiện click
    container.querySelectorAll('.doc-item').forEach(el => {
      el.addEventListener('click', () => {
        const docId = el.getAttribute('data-id');
        selectDocument(docId);
      });
    });

    // Nếu chưa chọn văn bản nào hoặc văn bản hiện tại vẫn còn trong kết quả, chọn nó
    if (results.length > 0) {
      if (!currentSelectedId || !results.some(x => x.id === currentSelectedId)) {
        selectDocument(results[0].id);
      }
    }
  }

  /**
   * Chọn và xem chi tiết một văn bản
   */
  function selectDocument(docId) {
    currentSelectedId = docId;

    // Highlight item bên trái
    const container = document.getElementById('docListContainer');
    container.querySelectorAll('.doc-item').forEach(el => {
      if (el.getAttribute('data-id') === docId) {
        el.classList.add('active');
        el.scrollIntoView({ block: 'nearest' });
      } else {
        el.classList.remove('active');
      }
    });

    const doc = window.LawSyncEngine.getLawById(docId);
    if (!doc) return;

    renderDocumentDetail(doc);
    document.getElementById('detailActionBar').style.display = 'flex';
  }

  /**
   * Render chi tiết văn bản sang cột phải
   */
  function renderDocumentDetail(doc) {
    const detailBox = document.getElementById('docDetailContainer');

    // Khối cảnh báo quan hệ pháp lý
    let alertHtml = '';
    if (doc.isExpired || doc.status === 'Hết hiệu lực') {
      if (doc.replacedBy) {
        alertHtml = `
          <div class="relation-alert-box danger">
            ⚠️ <b>CẢNH BÁO HIỆU LỰC:</b> Văn bản này đã <b>HẾT HIỆU LỰC</b> và được thay thế bởi:
            <br/>➡️ <b style="color: #003366; cursor: pointer; text-decoration: underline;" id="btnGoToReplaced" data-id="${doc.replacedBy.id}">
              ${escapeHtml(doc.replacedBy.code || doc.replacedBy.name)}
            </b> (${doc.replacedBy.issueDate || doc.replacedBy.year}).
          </div>
        `;
      } else {
        alertHtml = `
          <div class="relation-alert-box danger">
            ⚠️ <b>CẢNH BÁO HIỆU LỰC:</b> Văn bản này hiện ghi nhận đã <b>HẾT HIỆU LỰC</b> hoặc bị bãi bỏ. Cân nhắc kỹ trước khi đưa vào căn cứ pháp lý.
          </div>
        `;
      }
    } else if (doc.loaiPhapLy === 'Sửa đổi bổ sung' || doc.loaiPhapLy === 'Văn bản sửa đổi') {
      alertHtml = `
        <div class="relation-alert-box">
          ℹ️ <b>VĂN BẢN SỬA ĐỔI / BỔ SUNG:</b> Đây là văn bản sửa đổi, bổ sung một phần cho các văn bản quy phạm gốc cùng lĩnh vực.
        </div>
      `;
    }

    // Nội dung toàn văn hoặc banner file gốc
    let contentHtml = '';
    if (doc.co_toan_van_html && doc.noi_dung_toan_van && doc.noi_dung_toan_van.trim().length > 50) {
      contentHtml = `
        <div class="fulltext-section-title">NỘI DUNG VĂN BẢN TRÍCH XUẤT</div>
        <div class="fulltext-content">${escapeHtml(doc.noi_dung_toan_van)}</div>
      `;
    } else {
      const targetUrl = doc.duong_dan_file_dinh_kem || doc.linkGoc || 'https://qlda.gxd.vn/';
      contentHtml = `
        <div class="no-text-card">
          <h4>📄 Văn bản đính kèm dạng File gốc</h4>
          <p style="margin-bottom: 12px; font-size: 12px;">Văn bản này được phát hành dưới dạng tệp đính kèm gốc (PDF/Word). Để đảm bảo dung lượng siêu nhẹ cho hệ thống, văn bản không lưu file cục bộ.</p>
          <a href="${escapeHtml(targetUrl)}" target="_blank" class="win-btn win-btn-primary" style="display: inline-flex; text-decoration: none;">
            🌐 Xem hoặc tải tệp gốc trên QLDA GXD
          </a>
        </div>
      `;
    }

    const citationText = doc.citation || `- Căn cứ ${doc.name};`;

    detailBox.innerHTML = `
      <div class="doc-headline">${escapeHtml(doc.name)}</div>
      
      ${alertHtml}

      <table class="doc-meta-table">
        <tbody>
          <tr>
            <td class="label-cell">Số hiệu văn bản:</td>
            <td style="font-weight: 700; color: #003366;">${escapeHtml(doc.code || doc.name)}</td>
            <td class="label-cell">Tình trạng hiệu lực:</td>
            <td>
              ${doc.isExpired ? '<b style="color: #C5221F;">Hết hiệu lực</b>' : '<b style="color: #137333;">Còn hiệu lực</b>'} 
              (${escapeHtml(doc.loaiPhapLy || 'Văn bản gốc')})
            </td>
          </tr>
          <tr>
            <td class="label-cell">Nhóm chủ đề:</td>
            <td>${escapeHtml(doc.nhomChuDe || 'Chung')}</td>
            <td class="label-cell">Ngày ban hành:</td>
            <td>${escapeHtml(doc.issueDate || (doc.year ? 'Năm ' + doc.year : 'Chưa rõ'))}</td>
          </tr>
          <tr>
            <td class="label-cell">Đường dẫn bài viết:</td>
            <td colspan="3">
              <a href="${escapeHtml(doc.linkGoc || '#')}" target="_blank" style="color: #0066CC; text-decoration: underline;">
                ${escapeHtml(doc.linkGoc || 'Chưa cập nhật')}
              </a>
            </td>
          </tr>
        </tbody>
      </table>

      <!-- HỘP TRÍCH DẪN CĂN CỨ PHÁP LÝ CHUẨN NGHỊ ĐỊNH 30 -->
      <div class="citation-box">
        <div class="citation-box-header">
          <span>ĐỊNH DẠNG CĂN CỨ PHÁP LÝ (CHUẨN NGHỊ ĐỊNH 30):</span>
          <span style="font-size: 10px; color: #64748B;">Sẵn sàng chèn vào Word</span>
        </div>
        <div class="citation-text" id="citationBoxText">${escapeHtml(citationText)}</div>
      </div>

      <!-- VÙNG NỘI DUNG TOÀN VĂN -->
      ${contentHtml}
    `;

    // Nút chuyển đến văn bản thay thế nếu có
    const btnGo = document.getElementById('btnGoToReplaced');
    if (btnGo) {
      btnGo.addEventListener('click', () => {
        const targetId = btnGo.getAttribute('data-id');
        selectDocument(targetId);
      });
    }
  }

  /**
   * Xử lý Chèn Căn cứ pháp lý vào Word
   */
  async function handleInsertToWord() {
    if (!currentSelectedId) return;
    const doc = window.LawSyncEngine.getLawById(currentSelectedId);
    if (!doc) return;

    const citation = doc.citation || `- Căn cứ ${doc.name};`;

    // 1. Gửi qua BroadcastChannel
    if (broadcastChannel) {
      broadcastChannel.postMessage({
        type: 'INSERT_LAW_CITATION',
        docId: doc.id,
        citation: citation,
        code: doc.code,
        name: doc.name
      });
    }

    // 2. Gửi qua Office.context.ui.messageParent (nếu mở bằng displayDialogAsync)
    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.messageParent) {
      try {
        Office.context.ui.messageParent(JSON.stringify({
          type: 'INSERT_LAW_CITATION',
          docId: doc.id,
          citation: citation,
          code: doc.code,
          name: doc.name
        }));
      } catch (e) {}
    }

    // 3. Fallback qua window.opener
    if (window.opener && window.opener !== window) {
      try {
        window.opener.postMessage({
          type: 'INSERT_LAW_CITATION',
          docId: doc.id,
          citation: citation,
          code: doc.code,
          name: doc.name
        }, '*');
      } catch (e) {}
    }

    // Thông báo nhanh
    const btn = document.getElementById('btnInsertCitation');
    const oldText = btn.textContent;
    btn.textContent = '✅ Đã gửi lệnh chèn!';
    btn.style.background = '#28a745';
    setTimeout(() => {
      btn.textContent = oldText;
      btn.style.background = '';
    }, 1500);
  }

  /**
   * Sao chép trích dẫn vào clipboard
   */
  function handleCopyCitation() {
    if (!currentSelectedId) return;
    const doc = window.LawSyncEngine.getLawById(currentSelectedId);
    if (!doc) return;

    const citation = doc.citation || `- Căn cứ ${doc.name};`;
    navigator.clipboard.writeText(citation).then(() => {
      const btn = document.getElementById('btnCopyCitation');
      const oldText = btn.textContent;
      btn.textContent = '✅ Đã chép!';
      setTimeout(() => {
        btn.textContent = oldText;
      }, 1500);
    }).catch(err => {
      alert('Không thể sao chép tự động: ' + err.message);
    });
  }

  /**
   * Mở link bài viết gốc trên QLDA GXD
   */
  function handleOpenSourceLink() {
    if (!currentSelectedId) return;
    const doc = window.LawSyncEngine.getLawById(currentSelectedId);
    if (!doc || !doc.linkGoc) return;
    window.open(doc.linkGoc, '_blank');
  }

  /**
   * Kích hoạt Delta Sync từ Supabase
   */
  async function handleSyncCloud() {
    const btn = document.getElementById('btnSyncCloud');
    btn.disabled = true;
    btn.textContent = '⏳ Đang đồng bộ...';

    const result = await window.LawSyncEngine.syncFromCloud();
    btn.disabled = false;
    btn.textContent = '⚡ Đồng bộ Đám mây';

    if (result.success) {
      renderDocumentList();
      updateStatusStrip();
      alert(`[ĐỒNG BỘ THÀNH CÔNG]\n\n${result.message}`);
    } else {
      if (confirm(result.message + '\n\nBạn có muốn mở cửa sổ Cấu hình API để nhập thông tin Supabase không?')) {
        openConfigModal();
      }
    }
  }

  /**
   * Modal Cấu hình Supabase
   */
  function openConfigModal() {
    const engine = window.LawSyncEngine;
    document.getElementById('txtSupabaseUrl').value = engine.config.supabaseUrl || '';
    document.getElementById('txtSupabaseAnonKey').value = engine.config.anonKey || '';
    document.getElementById('modalConfig').classList.add('show');
  }

  function closeConfigModal() {
    document.getElementById('modalConfig').classList.remove('show');
  }

  function saveConfigModal() {
    const url = document.getElementById('txtSupabaseUrl').value.trim();
    const key = document.getElementById('txtSupabaseAnonKey').value.trim();

    window.LawSyncEngine.saveConfig({
      supabaseUrl: url,
      anonKey: key
    });

    closeConfigModal();
    updateStatusStrip();
    alert('Đã lưu cấu hình Supabase Cloud thành công!');
  }

  async function testSupabaseConnection() {
    const url = document.getElementById('txtSupabaseUrl').value.trim().replace(/\/$/, '');
    const key = document.getElementById('txtSupabaseAnonKey').value.trim();

    if (!url || !key) {
      alert('Vui lòng nhập đầy đủ Supabase URL và Anon Public Key để kiểm tra.');
      return;
    }

    const btn = document.getElementById('btnTestConnection');
    btn.disabled = true;
    btn.textContent = 'Đang kiểm tra...';

    try {
      const res = await fetch(`${url}/rest/v1/system_metadata?select=*&limit=1`, {
        method: 'GET',
        headers: {
          'apikey': key,
          'Authorization': `Bearer ${key}`
        }
      });

      if (res.ok) {
        alert('✅ Kết nối thành công tới Supabase Cloud Database!');
      } else {
        alert(`❌ Kết nối thất bại (Mã lỗi ${res.status}): Vui lòng kiểm tra lại URL hoặc Anon Key.`);
      }
    } catch (err) {
      alert(`❌ Lỗi mạng khi kết nối: ${err.message}`);
    } finally {
      btn.disabled = false;
      btn.textContent = '🔍 Kiểm tra kết nối';
    }
  }

  function escapeHtml(str) {
    if (!str || typeof str !== 'string') return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

})();
