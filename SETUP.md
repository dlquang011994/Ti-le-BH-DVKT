# Cài đặt 1 lần cho Dashboard Tỉ lệ BH

> Chỉ làm các bước dưới đây **1 lần duy nhất**. Sau đó mỗi tháng chỉ cần
> bỏ 2 file mới vào Drive rồi bấm 1 nút là xong (xem Bước 5).

- Kho chứa dashboard: https://github.com/dlquang011994/Ti-le-BH-DVKT
- Trang dashboard: https://dlquang011994.github.io/Ti-le-BH-DVKT/

> ⚠️ **Lưu ý quan trọng:** File JSON chứa "chìa khóa" (khóa bí mật) chỉ được
> dán vào đúng 1 chỗ là mục Secret của kho GitHub (Bước 3). Không gửi file
> này cho ai, không dán lên mạng, không lưu vào file nào trong kho.

---

## Bước 1. Tạo tài khoản dịch vụ (service account) trên Google Cloud

Tài khoản này là "người máy" giúp hệ thống tự đọc file Excel trong Drive
và tự ghi số liệu vào dashboard. Làm như sau:

1. Mở https://console.cloud.google.com/ và đăng nhập bằng tài khoản Google
   của bạn.
2. Bấm vào ô chọn dự án ở góc trên bên trái (chữ "Select a project" hoặc
   tên dự án hiện tại) → bấm **NEW PROJECT** → đặt tên, ví dụ
   `Ti le BH DVKT` → bấm **CREATE**. Đợi vài giây cho Google tạo xong,
   rồi chọn dự án vừa tạo.
3. Bật 2 tính năng cần dùng:
   - Vào menu trái → **APIs & Services** → **Library**.
   - Gõ `Google Drive API` vào ô tìm kiếm → bấm vào kết quả → bấm **ENABLE**.
   - Quay lại Library, gõ `Google Sheets API` → bấm vào kết quả → bấm **ENABLE**.
4. Tạo tài khoản dịch vụ:
   - Vào menu trái → **IAM & Admin** → **Service Accounts** → bấm
     **+ CREATE SERVICE ACCOUNT**.
   - Đặt tên, ví dụ `bh-dashboard` → bấm **CREATE AND CONTINUE** →
     bấm **CONTINUE** → bấm **DONE** (không cần chọn thêm quyền gì).
5. Tạo chìa khóa JSON:
   - Bấm vào tài khoản dịch vụ vừa tạo → tab **KEYS** → **ADD KEY** →
     **Create new key** → chọn **JSON** → bấm **CREATE**.
   - Google tải về 1 file `.json`. Giữ file này cẩn thận, lát nữa dùng
     ở Bước 3.
6. Ghi lại địa chỉ email của tài khoản dịch vụ (dạng
   `ten-gi-do@ten-du-an.iam.gserviceaccount.com`, thấy ở trang
   Service Accounts). Lát nữa dùng ở Bước 2 và Bước 4.

## Bước 2. Chia sẻ thư mục Drive cho tài khoản dịch vụ

1. Mở Google Drive, tìm thư mục **Tỉ lệ BH DVKT - 2026**.
2. Chuột phải vào thư mục → **Share** (Chia sẻ).
3. Dán địa chỉ email của tài khoản dịch vụ (đã ghi ở Bước 1, bước 6)
   vào ô "Add people".
4. Chọn quyền **Editor** (Người chỉnh sửa) — cần quyền này vì hệ thống
   ghi nhóm sản phẩm vào cột A của các file tháng.
5. Bấm **Send** (hoặc **Share**). Không cần tích gửi thông báo.

## Bước 3. Lưu chìa khóa JSON vào kho GitHub (Secret)

1. Mở file `.json` đã tải về ở Bước 1 bằng Notepad (Windows) hoặc
   TextEdit (Mac). Bấm Ctrl+A (hoặc Cmd+A) rồi Ctrl+C (Cmd+C) để
   copy **toàn bộ** nội dung file.
2. Mở kho https://github.com/dlquang011994/Ti-le-BH-DVKT →
   tab **Settings** → menu trái **Secrets and variables** → **Actions**.
3. Bấm **New repository secret**.
4. Name (tên): `GOOGLE_SERVICE_ACCOUNT_JSON`
5. Secret (nội dung): dán toàn bộ nội dung file JSON vừa copy.
6. Bấm **Add secret**.

## Bước 4. Tạo bảng Google Sheet để duyệt mã sản phẩm mới

1. Mở https://sheets.google.com/ → bấm **+** (Blank spreadsheet) để
   tạo bảng mới. Đặt tên, ví dụ: **Bảng tra mã - Tỉ lệ BH**.
2. Đổi tên tab đầu tiên thành **Bảng tra**, rồi gõ đúng 4 tiêu đề này
   vào hàng 1 (mỗi tiêu đề 1 ô, theo đúng thứ tự):
   `Mã sản phẩm` | `Nhóm` | `Hãng` | `Tên sản phẩm`
3. Bấm dấu **+** ở dưới cùng để thêm tab mới, đặt tên **Mã mới chờ gán**,
   rồi gõ đúng 5 tiêu đề này vào hàng 1:
   `Mã sản phẩm` | `Tên sản phẩm` | `Hãng` | `Số máy bán` | `Nhóm`
4. Chia sẻ bảng này cho tài khoản dịch vụ: bấm nút **Share** ở góc
   trên bên phải → dán địa chỉ email tài khoản dịch vụ → quyền
   **Editor** → bấm **Send**.
5. Lấy mã của bảng (Sheet ID): nhìn lên thanh địa chỉ của trình duyệt,
   địa chỉ có dạng
   `https://docs.google.com/spreadsheets/d/XXXXXXX/edit` — phần `XXXXXXX`
   ở giữa chính là Sheet ID. Copy phần đó.
6. Lưu Sheet ID vào kho GitHub: mở kho → tab **Settings** →
   **Secrets and variables** → **Actions** → **New repository secret**:
   - Name: `REVIEW_SHEET_ID`
   - Secret: dán Sheet ID vừa copy.
   - Bấm **Add secret**.

## Bước 5. Cập nhật số liệu mỗi tháng (làm hằng tháng)

1. Bỏ 2 file Excel của tháng mới vào đúng thư mục trong Drive
   **Tỉ lệ BH DVKT - 2026**: file bán vào `Bán/<năm>/<tháng>/`,
   file bảo hành vào `BH/<năm>/<tháng>/`.
2. Mở kho https://github.com/dlquang011994/Ti-le-BH-DVKT →
   tab **Actions** → chọn **Refresh dashboard data** ở menu trái →
   bấm nút **Run workflow** → bấm **Run workflow**.
3. Đợi vài phút (tháng nhiều dữ liệu có thể mất 10–20 phút). Khi dòng
   chạy hiện dấu tick xanh là xong, dashboard tự có số mới.
4. Hệ thống cũng tự chạy vào **ngày 5 mỗi tháng lúc 9 giờ sáng**
   (giờ Việt Nam), nên nếu đã bỏ file vào Drive trước ngày 5 thì
   không cần bấm tay.

### Nếu có mã sản phẩm mới

Khi tháng mới có mã sản phẩm chưa từng thấy, hệ thống liệt kê các mã đó
vào tab **Mã mới chờ gán** trong bảng Google Sheet (Bước 4) và dừng lại
chờ bạn gán nhóm. Cách xử lý:

1. Mở bảng Google Sheet, vào tab **Mã mới chờ gán**.
2. Điền cột **Nhóm** cho từng mã mới: 1 trong 7 nhóm (MT, Apple, ĐTDĐ,
   Samsung, Smartwatch, MTB, Máy In) hoặc `Loại` nếu mã đó không thuộc
   nhóm nào.
3. Chạy lại workflow như Bước 5 (mục 2). Hệ thống sẽ nhớ các mã này,
   tháng sau không hỏi lại nữa.

## Kiểm tra đã xong chưa

- [ ] Mở https://dlquang011994.github.io/Ti-le-BH-DVKT/ — trang hiện số liệu,
      chọn tháng nào cũng có số.
- [ ] Trong kho GitHub, tab **Actions**, lần chạy **Refresh dashboard data**
      gần nhất hiện dấu tick xanh.
- [ ] Tab **Mã mới chờ gán** trong Google Sheet không còn mã nào chưa
      điền nhóm.
