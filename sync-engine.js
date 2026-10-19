/**
 * sync-engine.js
 * Bộ điều phối đồng bộ dữ liệu Pháp luật: Offline-First + Delta Sync + Supabase Cloud
 * Đảm bảo ứng dụng chạy tức thì không cần mạng, đồng thời hỗ trợ đồng bộ gia tăng khi có kết nối Cloud.
 */

(function (global) {
  'use strict';

  const STORAGE_KEY_LAWS = 'GXD_LOCAL_LAWS_DB';
  const STORAGE_KEY_REVISION = 'GXD_LOCAL_SYNC_REVISION';
  const STORAGE_KEY_CONFIG = 'GXD_SUPABASE_CONFIG';

  class LawSyncEngine {
    constructor() {
      this.localRevision = parseInt(localStorage.getItem(STORAGE_KEY_REVISION) || '1', 10);
      this.laws = [];
      this.config = this.loadConfig();
      this.initLocalData();
    }

    /**
     * Tải cấu hình kết nối Supabase từ LocalStorage
     */
    loadConfig() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY_CONFIG);
        if (raw) return JSON.parse(raw);
      } catch (e) {}
      return {
        supabaseUrl: 'https://mabnaygyrscdfsqownhy.supabase.co',
        anonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1hYm5heWd5cnNjZGZzcW93bmh5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI0MTc0NjcsImV4cCI6MjEwNzk5MzQ2N30.iRkh626KK81gWoaDe7scHR1SwfxkuuPyUMXrXz2kULo',
        autoSync: true,
        lastSyncTime: null
      };
    }

    /**
     * Lưu cấu hình kết nối Supabase
     */
    saveConfig(cfg) {
      this.config = { ...this.config, ...cfg };
      localStorage.setItem(STORAGE_KEY_CONFIG, JSON.stringify(this.config));
    }

    /**
     * Khởi tạo dữ liệu cục bộ (Offline-First)
     * Nếu LocalStorage chưa có, tự động nạp 515 văn bản từ law-data.js (bundled)
     */
    initLocalData() {
      try {
        const cached = localStorage.getItem(STORAGE_KEY_LAWS);
        if (cached) {
          this.laws = this.deduplicateLaws(JSON.parse(cached));
        }
      } catch (e) {
        console.warn('Lỗi đọc cache văn bản cục bộ:', e);
      }

      // Nếu chưa có cache hoặc rỗng, lấy từ window.LAW_DATABASE (law-data.js)
      if ((!this.laws || this.laws.length === 0) && global.LAW_DATABASE && Array.isArray(global.LAW_DATABASE)) {
        this.laws = JSON.parse(JSON.stringify(global.LAW_DATABASE));
        // Chuẩn hóa cờ bổ sung
        this.laws.forEach(item => {
          if (item.co_toan_van_html === undefined) {
            item.co_toan_van_html = Boolean(item.noi_dung_toan_van && item.noi_dung_toan_van.length > 100);
          }
          if (item.is_deleted === undefined) item.is_deleted = false;
        });
        this.saveLocalData();
      }
    }

    /**
     * Lưu dữ liệu vào LocalStorage an toàn
     */
    saveLocalData() {
      try {
        localStorage.setItem(STORAGE_KEY_LAWS, JSON.stringify(this.laws));
        localStorage.setItem(STORAGE_KEY_REVISION, String(this.localRevision));
      } catch (e) {
        console.warn('LocalStorage đầy hoặc không khả dụng:', e);
      }
    }

    /**
     * Khử trùng lặp danh sách văn bản theo số hiệu chuẩn hóa
     */
    deduplicateLaws(list) {
      if (!list || !Array.isArray(list)) return [];
      const map = new Map();
      for (const item of list) {
        const rawCode = item.code || item.so_hieu || '';
        const key = rawCode.trim().toLowerCase().replace(/[-_.\s/]/g, '') || String(item.id);
        if (!map.has(key)) {
          map.set(key, item);
        } else {
          const existing = map.get(key);
          const lenExist = (existing.noi_dung_toan_van || '').length;
          const lenCur = (item.noi_dung_toan_van || '').length;
          if (lenCur > lenExist || (!existing.co_toan_van_html && item.co_toan_van_html)) {
            map.set(key, item);
          }
        }
      }
      return Array.from(map.values());
    }

    /**
     * Lấy toàn bộ danh sách văn bản hiện hành (bỏ qua bản ghi is_deleted = true và khử trùng)
     */
    getAllActiveLaws() {
      if (!this.laws || !Array.isArray(this.laws)) return [];
      const active = this.laws.filter(doc => !doc.is_deleted);
      return this.deduplicateLaws(active);
    }

    getAllDocuments() {
      return this.getAllActiveLaws();
    }

    saveDocumentLocally(doc) {
      if (doc && doc.id) {
        const idx = this.laws.findIndex(x => x.id === doc.id);
        if (idx !== -1) {
          this.laws[idx] = { ...this.laws[idx], ...doc };
        }
      }
      this.saveLocalData();
    }

    /**
     * Tìm kiếm văn bản theo từ khóa (Số hiệu, tên, trích yếu, nội dung)
     */
    searchLaws(keyword, category = 'all', statusFilter = 'all') {
      let list = this.getAllActiveLaws();

      if (category && category !== 'all') {
        list = list.filter(item => item.cat === category || item.nhomChuDe === category);
      }

      if (statusFilter && statusFilter !== 'all') {
        if (statusFilter === 'con-hieu-luc') {
          list = list.filter(item => !item.isExpired && item.status !== 'Hết hiệu lực');
        } else if (statusFilter === 'het-hieu-luc') {
          list = list.filter(item => item.isExpired || item.status === 'Hết hiệu lực');
        } else if (statusFilter === 'sua-doi') {
          list = list.filter(item => item.loaiPhapLy === 'Sửa đổi bổ sung' || item.loaiPhapLy === 'Văn bản sửa đổi');
        }
      }

      if (!keyword || !keyword.trim()) return list;

      const q = keyword.trim().toLowerCase();
      return list.filter(item => {
        const code = (item.code || '').toLowerCase();
        const name = (item.name || '').toLowerCase();
        const cit = (item.citation || '').toLowerCase();
        const text = (item.noi_dung_toan_van || '').toLowerCase();
        return code.includes(q) || name.includes(q) || cit.includes(q) || text.includes(q);
      });
    }

    /**
     * Lấy chi tiết 1 văn bản kèm thông tin quan hệ pháp lý
     */
    getLawById(id) {
      const doc = this.laws.find(item => item.id === id);
      if (!doc) return null;

      // Phân tích văn bản thay thế nếu có
      let replacedBy = null;
      if (doc.isExpired || doc.status === 'Hết hiệu lực') {
        // Tìm văn bản mới nhất trong cùng phân nhóm
        const candidates = this.laws.filter(other => 
          other.id !== doc.id && 
          other.cat === doc.cat && 
          !other.isExpired && 
          (other.year || 0) >= (doc.year || 0)
        );
        if (candidates.length > 0) {
          // Lấy văn bản có năm ban hành lớn nhất
          candidates.sort((a, b) => (b.year || 0) - (a.year || 0));
          replacedBy = candidates[0];
        }
      }

      return {
        ...doc,
        replacedBy
      };
    }

    /**
     * Đồng bộ gia tăng (Delta Sync) từ Supabase
     */
    async syncFromCloud() {
      if (!this.config.supabaseUrl || !this.config.anonKey) {
        return {
          success: false,
          message: 'Chưa cấu hình Supabase URL hoặc Anon Key. Vui lòng bấm Cấu hình API để kết nối Cloud.'
        };
      }

      try {
        const baseUrl = this.config.supabaseUrl.replace(/\/$/, '');
        const endpoint = `${baseUrl}/rest/v1/van_ban_phap_luat?select=*&sync_revision=gt.${this.localRevision}&order=sync_revision.asc`;

        const res = await fetch(endpoint, {
          method: 'GET',
          headers: {
            'apikey': this.config.anonKey,
            'Authorization': `Bearer ${this.config.anonKey}`,
            'Content-Type': 'application/json'
          }
        });

        if (!res.ok) {
          throw new Error(`Lỗi HTTP ${res.status}: ${res.statusText}`);
        }

        const deltaRecords = await res.json();

        if (!Array.isArray(deltaRecords) || deltaRecords.length === 0) {
          this.config.lastSyncTime = new Date().toISOString();
          this.saveConfig(this.config);
          return {
            success: true,
            updatedCount: 0,
            message: 'Dữ liệu đã ở phiên bản mới nhất! Không có bản ghi mới.'
          };
        }

        let updated = 0;
        let deleted = 0;
        let maxRev = this.localRevision;

        deltaRecords.forEach(rec => {
          if (rec.sync_revision && rec.sync_revision > maxRev) {
            maxRev = rec.sync_revision;
          }

          const existingIdx = this.laws.findIndex(x => x.id === rec.id || (rec.so_hieu && x.code === rec.so_hieu));

          if (rec.is_deleted) {
            // Xóa khỏi cache
            if (existingIdx !== -1) {
              this.laws.splice(existingIdx, 1);
              deleted++;
            }
          } else {
            const mappedItem = {
              id: rec.id,
              cat: rec.chu_de_id || 'qlda',
              nhomChuDe: rec.ten_chu_de || 'Quản lý Dự án & Thủ tục ĐTXD',
              code: rec.so_hieu,
              name: rec.ten_van_ban,
              issueDate: rec.ngay_ban_hanh,
              effectDate: rec.ngay_hieu_luc,
              status: rec.tinh_trang_hieu_luc || 'Còn hiệu lực',
              citation: rec.trich_dan_chuan || `- Căn cứ ${rec.ten_van_ban};`,
              linkGoc: rec.duong_dan_gxd,
              duong_dan_file_dinh_kem: rec.duong_dan_file_dinh_kem,
              co_toan_van_html: rec.co_toan_van_html !== false,
              noi_dung_toan_van: rec.noi_dung_toan_van || '',
              year: rec.ngay_ban_hanh ? parseInt(rec.ngay_ban_hanh.substring(0, 4), 10) : new Date().getFullYear(),
              isExpired: rec.tinh_trang_hieu_luc === 'Hết hiệu lực',
              loaiPhapLy: rec.loai_phap_ly || 'Văn bản gốc'
            };

            if (existingIdx !== -1) {
              this.laws[existingIdx] = mappedItem;
            } else {
              this.laws.push(mappedItem);
            }
            updated++;
          }
        });

        this.localRevision = maxRev;
        this.config.lastSyncTime = new Date().toISOString();
        this.saveConfig(this.config);
        this.saveLocalData();

        return {
          success: true,
          updatedCount: updated,
          deletedCount: deleted,
          newRevision: maxRev,
          message: `Đồng bộ thành công! Cập nhật: ${updated} văn bản, Dọn dẹp: ${deleted} văn bản.`
        };
      } catch (err) {
        console.error('Lỗi Delta Sync:', err);
        return {
          success: false,
          message: `Lỗi kết nối Supabase: ${err.message}`
        };
      }
    }
  }

  global.LawSyncEngine = new LawSyncEngine();
})(typeof window !== 'undefined' ? window : this);
