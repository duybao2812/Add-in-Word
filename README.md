# Microsoft Word Taskpane Add-in (WinForms Classic UI)

Dự án Add-in cho Microsoft Word Desktop xây dựng trên nền tảng **Office.js** với phong cách giao diện cổ điển **Windows Forms (WinForms Classic)**: phẳng, không bo góc (0px radius), phông chữ sắc nét và bố cục co giãn linh hoạt (Fluid Responsive Layout).

---

## 📁 Cấu trúc dự án

```text
Add-in-Word/
├── assets/                  # Biểu tượng của Add-in (16x16, 32x32, 64x64, 80x80)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── manifest.xml             # Tệp cấu hình Add-in chính thức cho Word Desktop
├── taskpane.html            # Khung giao diện HTML5 chuẩn 3 GroupBox WinForms
├── taskpane.css             # Định dạng phong cách WinForms (0px radius, bảng màu hệ thống)
├── taskpane.js              # Xử lý khởi tạo Office.js và logic giao diện
├── server.js                # Máy chủ HTTPS cục bộ hỗ trợ CORS & Dev-Certs
├── package.json             # Danh mục thư viện phụ thuộc và scripts terminal
└── README.md                # Tài liệu hướng dẫn sử dụng
```

---

## 🎨 Đặc điểm UI/UX WinForms Classic

1. **Tuyệt đối không bo góc**:
   - Toàn bộ các phần tử (`button`, `input`, `select`, `textarea`, `fieldset`, `statusstrip`) đều tuân thủ nguyên tắc `border-radius: 0px !important`.
2. **Hệ thống phông chữ & Màu sắc WinForms**:
   - Phông chữ: `Segoe UI`, `Tahoma`, `Arial, sans-serif` kích thước 11px - 13px.
   - Nền Form chính: `#F0F0F0` (Màu Control tiêu chuẩn của hệ thống Windows).
   - Nút bấm (Button): Nền `#E1E1E1`, viền `1px solid #707070`, khi hover chuyển sang `#E5F1FB` viền `#0078D7`, khi nhấn chìm xuống `#CCE4F7` viền `#005499`.
   - Khung nhóm (GroupBox): Thẻ `<fieldset>` với viền `1px solid #A0A0A0`, tiêu đề `<legend>` màu xanh `#003399` in đậm.
   - Thanh trạng thái (StatusStrip): Cố định dưới đáy, phân chia các panel nhỏ viền chìm 3D hiển thị trạng thái, nền tảng host và phiên bản.
3. **Bố cục co giãn linh hoạt (Fluid & Responsive)**:
   - Tương thích tốt từ chiều ngang hẹp (280px) đến rộng (600px+).
   - Sử dụng CSS Grid tự động phân bổ cột (`repeat(auto-fit, minmax(130px, 1fr))`), khi kéo rộng các nút tự dàn đều, khi thu hẹp tự động rớt dòng ngay ngắn.

---

## 🛠 Hướng dẫn Khởi chạy và Thử nghiệm

### 1. Cài đặt các gói phụ thuộc
```bash
npm install
```

### 2. Cài đặt chứng chỉ SSL tự ký cho localhost (Nếu chạy lần đầu)
Microsoft Word Desktop yêu cầu kết nối HTTPS với chứng chỉ được tin cậy trên máy:
```bash
npm run certs
```
*(Nếu Windows hiện hộp thoại cảnh báo bảo mật, hãy chọn **Yes** để đồng ý thêm chứng chỉ phát triển vào Trusted Root).*

### 3. Khởi động Máy chủ HTTPS
```bash
npm run serve
```
Máy chủ sẽ lắng nghe tại:
- **Taskpane UI**: `https://localhost:3000/taskpane.html`
- **Manifest File**: `https://localhost:3000/manifest.xml`

*(Bạn có thể mở trực tiếp link trên bằng trình duyệt Chrome/Edge để xem trước giao diện và kiểm tra độ co giãn).*

### 4. Kiểm tra tính hợp lệ của Manifest
```bash
npm run validate
```

### 5. Tự động Sideload Add-in vào Word Desktop
Chạy lệnh sau để tự động khởi chạy Microsoft Word Desktop và nhúng Add-in vào thanh Ribbon:
```bash
npm start
```
Để dừng và gỡ Add-in ra khỏi Word:
```bash
npm run stop
```
