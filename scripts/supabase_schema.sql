-- ==============================================================================
-- SUPABASE POSTGRESQL SCHEMA MIGRATION: HỆ THỐNG TRA CỨU VĂN BẢN PHÁP LUẬT XÂY DỰNG
-- Phiên bản: Text-First – Offline-First – Production-Ready
-- Chức năng: Tạo cấu trúc bảng, RLS Policies, Indexes và Triggers Delta Sync
-- ==============================================================================

-- 1. BẬT TIỆN ÍCH EXTENSION UUID
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 2. BẢNG 1: CHỦ ĐỀ PHÁP LUẬT (chu_de_phap_luat)
CREATE TABLE IF NOT EXISTS public.chu_de_phap_luat (
    id VARCHAR(50) PRIMARY KEY,
    ten_chu_de VARCHAR(255) NOT NULL,
    mo_ta TEXT,
    stt_hien_thi INT DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Seed danh mục chủ đề ban đầu
INSERT INTO public.chu_de_phap_luat (id, ten_chu_de, mo_ta, stt_hien_thi) VALUES
('qlda', 'Quản lý Dự án & Thủ tục ĐTXD', 'Các quy định về thủ tục lập, thẩm định, phê duyệt dự án đầu tư xây dựng', 1),
('chi-phi', 'Quản lý Chi phí & Định mức ĐTXD', 'Tổng mức đầu tư, dự toán xây dựng công trình, định mức, đơn giá', 2),
('hop-dong', 'Hợp đồng Xây dựng', 'Hợp đồng thi công, tư vấn, EPC, thanh quyết toán hợp đồng', 3),
('dau-thau', 'Đấu thầu & Lựa chọn Nhà thầu', 'Luật đấu thầu, hồ sơ mời thầu, đánh giá hồ sơ dự thầu', 4),
('chat-luong', 'Quản lý Chất lượng & An toàn Thi công', 'Giám sát thi công, nghiệm thu bàn giao, bảo trì công trình', 5),
('luat-chung', 'Hệ thống Luật & Bộ Luật', 'Luật Xây dựng, Luật Đất đai, Luật Nhà ở, Luật Đầu tư...', 6)
ON CONFLICT (id) DO NOTHING;

-- 3. BẢNG 2: VĂN BẢN PHÁP LUẬT (van_ban_phap_luat)
CREATE TABLE IF NOT EXISTS public.van_ban_phap_luat (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    so_hieu VARCHAR(150) NOT NULL,
    ten_van_ban TEXT NOT NULL,
    loai_van_ban VARCHAR(50) DEFAULT 'Văn bản quy phạm',
    chu_de_id VARCHAR(50) REFERENCES public.chu_de_phap_luat(id) ON DELETE SET NULL,
    ngay_ban_hanh DATE,
    ngay_hieu_luc DATE,
    co_quan_ban_hanh VARCHAR(150),
    tinh_trang_hieu_luc VARCHAR(50) DEFAULT 'Còn hiệu lực', -- 'Còn hiệu lực', 'Hết hiệu lực', 'Bị sửa đổi một phần'
    loai_phap_ly VARCHAR(50) DEFAULT 'Văn bản gốc',         -- 'Văn bản gốc', 'Sửa đổi bổ sung', 'Thay thế'
    trich_dan_chuan TEXT,
    noi_dung_toan_van TEXT,
    co_toan_van_html BOOLEAN DEFAULT TRUE,                  -- Cờ đánh dấu có text HTML tra cứu hay chỉ có file
    duong_dan_file_dinh_kem TEXT,                           -- Link file PDF/Word gốc trên web nếu trang chỉ có file
    duong_dan_gxd TEXT,
    hash_noi_dung VARCHAR(64),                              -- Mã băm SHA-256 để chống trùng nội dung
    sync_revision BIGINT DEFAULT 1,                         -- Số phiên bản tăng dần cho Delta Sync
    is_deleted BOOLEAN DEFAULT FALSE,                       -- Soft Delete để Client dọn rác
    deleted_at TIMESTAMPTZ,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Chỉ mục tối ưu tốc độ tìm kiếm
CREATE INDEX IF NOT EXISTS idx_van_ban_so_hieu ON public.van_ban_phap_luat(so_hieu);
CREATE INDEX IF NOT EXISTS idx_van_ban_chu_de ON public.van_ban_phap_luat(chu_de_id);
CREATE INDEX IF NOT EXISTS idx_van_ban_sync_rev ON public.van_ban_phap_luat(sync_revision);
CREATE INDEX IF NOT EXISTS idx_van_ban_active ON public.van_ban_phap_luat(is_deleted);

-- 4. BẢNG 3: QUAN HỆ VĂN BẢN PHÁP LUẬT (quan_he_van_ban)
CREATE TABLE IF NOT EXISTS public.quan_he_van_ban (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    van_ban_nguon_id UUID REFERENCES public.van_ban_phap_luat(id) ON DELETE CASCADE,
    van_ban_dich_id UUID REFERENCES public.van_ban_phap_luat(id) ON DELETE CASCADE,
    loai_quan_he VARCHAR(50) NOT NULL,                      -- 'Sửa đổi', 'Bổ sung', 'Thay thế', 'Bãi bỏ', 'Hướng dẫn'
    pham_vi_tac_dong VARCHAR(100) DEFAULT 'Toàn bộ',        -- 'Toàn bộ' hoặc 'Một phần'
    ngay_hieu_luc_quan_he DATE,
    can_cu_xac_minh TEXT,                                   -- Trích dẫn điều khoản bãi bỏ/thay thế
    url_can_cu TEXT,
    trang_thai_xac_minh VARCHAR(50) DEFAULT 'Đã xác minh',  -- 'Đã xác minh' / 'Chờ kiểm tra'
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. BẢNG 4: LỊCH SỬ PHIÊN BẢN VĂN BẢN (lich_su_phien_ban_van_ban)
CREATE TABLE IF NOT EXISTS public.lich_su_phien_ban_van_ban (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    van_ban_id UUID REFERENCES public.van_ban_phap_luat(id) ON DELETE CASCADE,
    phien_ban INT NOT NULL DEFAULT 1,
    noi_dung_toan_van TEXT,
    hash_noi_dung VARCHAR(64),
    thoi_diem_thu_thap TIMESTAMPTZ DEFAULT NOW(),
    nguon_cap_nhat VARCHAR(100) DEFAULT 'Crawler Bot',
    ly_do_cap_nhat TEXT
);

-- 6. BẢNG 5: BIỂU MẪU VĂN BẢN (mau_van_ban)
CREATE TABLE IF NOT EXISTS public.mau_van_ban (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    tieu_de VARCHAR(255) NOT NULL,
    loai_van_ban VARCHAR(50) NOT NULL,                      -- 'cong-van', 'hop-dong', 'bien-ban', 'to-trinh'
    can_cu_mac_dinh JSONB,
    cau_truc_du_lieu JSONB,
    phien_ban INT DEFAULT 1,
    sync_revision BIGINT DEFAULT 1,
    is_deleted BOOLEAN DEFAULT FALSE,
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 7. BẢNG 6: SIÊU DỮ LIỆU ĐỒNG BỘ HỆ THỐNG (system_metadata)
CREATE TABLE IF NOT EXISTS public.system_metadata (
    key VARCHAR(50) PRIMARY KEY,
    sync_revision BIGINT DEFAULT 1,
    version_hash VARCHAR(64),
    total_records INT DEFAULT 0,
    last_updated TIMESTAMPTZ DEFAULT NOW()
);

-- Khởi tạo metadata ban đầu
INSERT INTO public.system_metadata (key, sync_revision, version_hash, total_records, last_updated)
VALUES ('current_sync_state', 1, 'init_hash', 0, NOW())
ON CONFLICT (key) DO NOTHING;

-- 8. BẢNG 7: NHẬT KÝ THU THẬP BOT (crawler_runs)
CREATE TABLE IF NOT EXISTS public.crawler_runs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    start_time TIMESTAMPTZ DEFAULT NOW(),
    end_time TIMESTAMPTZ,
    records_scanned INT DEFAULT 0,
    records_inserted INT DEFAULT 0,
    records_updated INT DEFAULT 0,
    error_count INT DEFAULT 0,
    status VARCHAR(50) DEFAULT 'RUNNING',                  -- 'RUNNING', 'SUCCESS', 'PARTIAL', 'FAILED'
    log_details JSONB
);

-- ==============================================================================
-- 9. POSTGRESQL TRIGGER: TỰ ĐỘNG TĂNG SYNC_REVISION VÀ METADATA KHI CÓ THAY ĐỔI
-- ==============================================================================

CREATE OR REPLACE FUNCTION public.fn_bump_sync_revision()
RETURNS TRIGGER 
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
    next_rev BIGINT;
BEGIN
    -- Lấy revision hiện tại + 1
    SELECT COALESCE(sync_revision, 0) + 1 INTO next_rev 
    FROM public.system_metadata 
    WHERE key = 'current_sync_state';

    -- Gán revision vào bản ghi văn bản
    NEW.sync_revision := next_rev;
    NEW.updated_at := NOW();

    -- Cập nhật lại system_metadata
    UPDATE public.system_metadata
    SET sync_revision = next_rev,
        total_records = (SELECT COUNT(*) FROM public.van_ban_phap_luat WHERE is_deleted = FALSE),
        last_updated = NOW()
    WHERE key = 'current_sync_state';

    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_bump_law_revision ON public.van_ban_phap_luat;
CREATE TRIGGER trg_bump_law_revision
BEFORE INSERT OR UPDATE ON public.van_ban_phap_luat
FOR EACH ROW EXECUTE FUNCTION public.fn_bump_sync_revision();

-- ==============================================================================
-- 10. THIẾT LẬP BẢO MẬT ROW LEVEL SECURITY (RLS)
-- ==============================================================================

ALTER TABLE public.chu_de_phap_luat ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.van_ban_phap_luat ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.quan_he_van_ban ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.lich_su_phien_ban_van_ban ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.mau_van_ban ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.system_metadata ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.crawler_runs ENABLE ROW LEVEL SECURITY;

-- Quyền cho Anon Key (Word Add-in Client): CHỈ ĐƯỢC ĐỌC (SELECT)
CREATE POLICY "Anon can view chu_de" ON public.chu_de_phap_luat FOR SELECT TO anon USING (true);
CREATE POLICY "Anon can view van_ban" ON public.van_ban_phap_luat FOR SELECT TO anon USING (true);
CREATE POLICY "Anon can view quan_he" ON public.quan_he_van_ban FOR SELECT TO anon USING (true);
CREATE POLICY "Anon can view mau_van_ban" ON public.mau_van_ban FOR SELECT TO anon USING (true);
CREATE POLICY "Anon can view system_metadata" ON public.system_metadata FOR SELECT TO anon USING (true);

-- Quyền toàn quyền cho Service Role (Crawler Bot trên GitHub Actions)
CREATE POLICY "Service role full access chu_de" ON public.chu_de_phap_luat FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access van_ban" ON public.van_ban_phap_luat FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access quan_he" ON public.quan_he_van_ban FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access phien_ban" ON public.lich_su_phien_ban_van_ban FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access mau_van_ban" ON public.mau_van_ban FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access system_metadata" ON public.system_metadata FOR ALL TO service_role USING (true);
CREATE POLICY "Service role full access crawler_runs" ON public.crawler_runs FOR ALL TO service_role USING (true);
