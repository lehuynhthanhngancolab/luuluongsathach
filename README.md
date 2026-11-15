# Phần mềm Xuất Lưu Lượng Sát Hạch Lái Xe Hàng Tháng (PDF → Excel)

Ứng dụng web cục bộ giúp **tự động trích xuất lịch sát hạch lái xe hàng tháng** từ file PDF (hoặc Excel/Word) và xuất ra file Excel gồm 3 sheet tổng hợp, dùng cho công tác thống kê lưu lượng sát hạch của từng cơ sở đào tạo và trung tâm sát hạch lái xe.

Toàn bộ quá trình xử lý diễn ra **hoàn toàn cục bộ trên máy tính** — không upload dữ liệu lên bất kỳ máy chủ nào.

## Tính năng chính

### 1. Nhập liệu linh hoạt

- Kéo thả hoặc chọn file **PDF** lịch sát hạch hàng tháng.
- Hỗ trợ thêm file **Excel** (`.xlsx`) và **Word** (`.docx`) làm nguồn dữ liệu thay thế.
- Tự động nhận diện **tháng/năm** từ nội dung file đầu vào.

### 2. Bóc tách dữ liệu PDF thông minh

- Bóc tách theo **lưới ô** (cell) của PDF, loại bỏ các đường kẻ gây nhiễu (`extractPageWithoutGridLines`).
- Gom cụm tọa độ text để khôi phục đúng ô chứa nội dung (`clusterCoords`, `joinLineItems`).
- Tự động **fallback sang OCR** bằng Tesseract khi PDF là bản scan (ảnh chụp/scan), không có lớp text. Lưu ý: dữ liệu ngôn ngữ `vie.traineddata` được Tesseract tải từ CDN ở lần chạy OCR đầu tiên nên bước này **cần có mạng**; sau đó được cache lại.
- Xem lại **file PDF gốc** song song với dữ liệu đã trích xuất để đối chiếu.

### 3. Chuẩn hóa & tính toán lịch

- Chuẩn hóa tên **Cơ sở đào tạo (CSĐT)** và **Trung tâm sát hạch (TTSH)** để ghép dòng chính xác (`normalizeCsdtName`, `normalizeTtshName`, `shortTtshName`).
- Ghép các dòng bị vỡ trong PDF, tách hạng xe/số lượng (`parseCsdtLine`, `extractHangSuffix`).
- Phân bổ lịch theo **hạng B, C** (`allocateRowCsdts`) và mở rộng dòng lịch theo từng ngày (`expandScheduleRows`).
- Tính lại toàn bộ ma trận tổng hợp khi có chỉnh sửa (`recalculateAll`).

### 4. Xem trước & chỉnh sửa trực tiếp

- Giao diện 3 tab xem trước, **nhấp vào ô bất kỳ để sửa** trước khi xuất Excel.
- Thêm / xoá dòng lịch, đánh số lại STT tự động.
- Thanh thống kê tổng quan (số lượt sát hạch, số CSĐT, số TTSH, tổng số lượng).

### 5. Xuất Excel 3 sheet

| Sheet | Tên | Nội dung |
| --- | --- | --- |
| 1 | `Hạng B,C` | Danh sách sát hạch chi tiết: Số TT, Ngày sát hạch, Cơ sở đào tạo, Trung tâm sát hạch lái xe, Số lượng / Hạng xe |
| 2 | `TONG_HOP_MAU_MOI` | Ma trận tổng hợp theo **CSĐT × Trung tâm SH** |
| 3 | `TONG_HOP_MAU_MOI_2` | Ma trận tổng hợp theo **Trung tâm SH × CSĐT** |

Định dạng xuất: font Times New Roman, viền lưới, cột ngày theo định dạng `dd/mm/yyyy`, chiều cao hàng tự co giãn theo số dòng gộp.

## Yêu cầu

- **Node.js >= 18** (kiểm tra bằng `node -v`)
- Trình duyệt web hiện đại: Chrome, Edge, Cốc Cốc, Firefox

> Không cần cài thêm package: toàn bộ thư viện (PDF.js, ExcelJS, Mammoth, Tesseract) đã được vendor sẵn trong thư mục `libs/`.

## Cách chạy

### Cách 1 — Windows (một cú bấm)

Double-click vào file [`Chạy Phần Mềm.bat`](Chạy%20Phần%20Mềm.bat). File này sẽ tự mở trình duyệt tại `http://localhost:8080` và khởi động server.

### Cách 2 — Dòng lệnh

```bash
npm start
```

Sau đó mở trình duyệt tại **http://localhost:8080**

Server tĩnh nằm ở cổng `8080`. Để đổi cổng, sửa hằng `PORT` trong [`server.js`](server.js).

### Cách 3 — Bản deploy trên GitHub Pages

Ứng dụng đã được deploy tại:

**https://lehuynhthanhngancolab.github.io/luuluongsathach/**

Mỗi khi có commit lên nhánh `main`, GitHub Actions tự động build và publish lại site.

## Cấu trúc thư mục

```
luuluongsathach/
├── server.js                              # Static server (Node.js, cổng 8080)
├── package.json                           # Metadata + script "start"
├── index.html                             # Toàn bộ giao diện + logic xử lý
├── Chạy Phần Mềm.bat                     # Launcher 1-click cho Windows
├── libs/                                  # Thư viện vendor (chạy offline, không cần npm install)
│   ├── pdf.min.js, pdf.worker.min.js      # PDF.js - đọc & render PDF
│   ├── exceljs.min.js                     # ExcelJS - tạo file .xlsx
│   ├── mammoth.browser.min.js             # Mammoth - đọc file .docx
│   ├── tesseract.min.js                   # Tesseract - OCR cho PDF scan
│   └── presets.js                         # Bộ lịch mẫu T2→T11 (tự nạp khi mở app, dùng để thử nhanh)
```

## Quy trình sử dụng

1. Mở ứng dụng tại `http://localhost:8080`.
2. Kéo thả file PDF lịch sát hạch vào khung nạp file (hoặc chọn file Excel/Word).
3. Chờ ứng dụng bóc tách xong; nếu PDF scan sẽ tự chạy OCR.
4. Kiểm tra ở các tab **Sheet 1 / Sheet 2 / Sheet 3**, nhấp vào ô để sửa nếu cần.
5. Nhấn **Xuất Excel** để tải file `.xlsx` gồm 3 sheet.

## Ghi chú

- Ứng dụng **không đọc file lịch sát hạch nào từ thư mục gốc** — mọi dữ liệu đầu vào đều do người dùng tự nạp lên (kéo thả PDF/Excel/Word) hoặc lấy từ bộ lịch mẫu tích hợp sẵn trong `libs/presets.js`.
- Không có bước cài đặt dependency nào (`npm install` không cần thiết vì `libs/` đã chứa sẵn bản bundle).
- Server chỉ phục vụ file tĩnh trong thư mục dự án và chặn truy cập ra ngoài (`path traversal` guard).
- Repo không chứa file dữ liệu lịch sát hạch (`.xlsx`) nhằm tránh công khai thông tin thực tế. Bộ lịch mẫu trong `libs/presets.js` là dữ liệu đã chuẩn hóa dùng để thử tính năng.
