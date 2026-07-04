# 📗 Hướng dẫn sử dụng (dành cho người không rành máy tính)

Công cụ chạy **ngay trên máy của bạn** (dữ liệu không gửi đi đâu cả), hiện lên
dưới dạng một trang web trong trình duyệt.

---

## 🖥️ Cài đặt lần đầu (chỉ làm 1 lần)

### Bước 1 — Cài Python
1. Vào trang: https://www.python.org/downloads/
2. Bấm nút vàng **Download Python** để tải về.
3. Mở file vừa tải. **QUAN TRỌNG:** ở màn hình đầu tiên, tích vào ô
   **“Add Python to PATH”** rồi bấm **Install Now**.
4. Cài xong bấm **Close**.

### Bước 2 — Tải công cụ về máy
1. Vào trang GitHub của dự án.
2. Bấm nút xanh **Code** → **Download ZIP**.
3. Giải nén file ZIP ra một thư mục (ví dụ ngoài Màn hình / Desktop).

---

## ▶️ Chạy công cụ (mỗi lần dùng)

- **Máy Windows:** vào thư mục vừa giải nén, **bấm đúp** vào file **`run.bat`**.
- **Máy Mac:** mở Terminal trong thư mục đó và gõ `bash run.sh`.

Lần chạy đầu sẽ hơi lâu (đang tự cài các thành phần). Sau đó trình duyệt sẽ tự
mở ra trang công cụ. Nếu không tự mở, vào trình duyệt gõ địa chỉ:
**http://localhost:8501**

> Muốn tắt: đóng cửa sổ màu đen (Windows) hoặc nhấn `Ctrl + C` (Mac).

---

## 🧾 Các bước dùng trên giao diện

1. **Tải file lên:**
   - *File dữ liệu chính* (bảng kê hóa đơn cần kiểm tra) — **bắt buộc**.
   - *File TMS*, *File DS rủi ro*, *File hóa đơn điện tử* — có thì tải thêm.
2. (Nếu file của bạn khác bố cục chuẩn) mở phần **⚙️ Cấu hình cột** bên trái để
   chỉnh lại số cột / tên cột.
3. Bấm **▶️ Bắt đầu kiểm tra**.
4. Xem kết quả và bấm **⬇️ Tải file Excel kết quả** — file gồm nhiều sheet:
   dữ liệu tổng hợp và từng loại cảnh báo (hóa đơn trùng, mặt hàng nghi ngờ,
   hóa đơn sai trạng thái, chênh lệch thuế…).

---

## 🧪 Muốn thử trước bằng dữ liệu mẫu?

Trong thư mục có sẵn folder **`samples/`** chứa 4 file mẫu. Cứ tải chúng lên
để xem công cụ hoạt động thế nào trước khi dùng dữ liệu thật.

---

## ❓ Gặp lỗi thường gặp

| Hiện tượng | Cách xử lý |
|-----------|-----------|
| Báo "Chưa cài đặt Python" | Làm lại Bước 1, nhớ tích **Add Python to PATH** |
| Không tìm thấy cột (vd cột 42) | Mở **⚙️ Cấu hình cột**, sửa lại số/tên cột cho đúng file của bạn |
| Kết quả sai vị trí | File thật khác bố cục mẫu — chỉnh lại ở phần cấu hình cột |
