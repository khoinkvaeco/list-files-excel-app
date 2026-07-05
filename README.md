# Công cụ hỗ trợ kiểm tra thuế (Excel)

Ứng dụng web (Streamlit) giúp tự động rà soát bảng kê hóa đơn phục vụ công tác
kiểm tra thuế: đối chiếu trạng thái người nộp thuế, danh sách doanh nghiệp rủi ro,
hóa đơn điện tử, phát hiện hóa đơn trùng, mặt hàng không phục vụ SXKD, chênh lệch
tiền thuế, v.v.

## 12 tác vụ được hỗ trợ

| # | Tác vụ | Module |
|---|--------|--------|
| 1 | Xử lý / chuẩn hóa MST (giữ số 0 đầu, bù MST 10 số, tách mã đơn vị phụ thuộc) | `tax_audit/mst.py` |
| 2 | Copy MST ra 1 sheet riêng để tra cứu TMS | `tax_audit/mst.py` |
| 3 | Lấy **trạng thái người nộp thuế** từ TMS (cột 42) | `tax_audit/taxpayer.py` |
| 4 | Lấy **ngày đóng trạng thái tổ chức** từ TMS (cột 51) | `tax_audit/taxpayer.py` |
| 5 | Cảnh báo hóa đơn xuất **sau ngày đóng trạng thái** | `tax_audit/taxpayer.py` |
| 6 | Dò MST với **danh sách DN rủi ro**, lấy văn bản (cột 5) | `tax_audit/risk.py` |
| 7 | Tìm mặt hàng **không phục vụ SXKD** (golf, quà, biếu, rượu, wine…) | `tax_audit/goods.py` |
| 8 | **Xóa số 0 đầu** của số hóa đơn | `tax_audit/invoice.py` |
| 9 | Tìm **hóa đơn kê khai trùng** (số HĐ + ngày + MST + giá trị chưa thuế) | `tax_audit/invoice.py` |
| 10 | Tra **file hóa đơn điện tử** → trạng thái + tổng tiền thuế | `tax_audit/einvoice.py` |
| 11 | Phân loại HĐ **bị thay thế / xóa bỏ / không tìm thấy** | `tax_audit/einvoice.py` |
| 12 | Tính **chênh lệch tiền thuế** = Tổng tiền thuế − Thuế GTGT | `tax_audit/einvoice.py` |
| + | **Chuyển hóa đơn điện tử XML → Excel** (đọc XML chuẩn TCT, xuất bảng hóa đơn + chi tiết hàng hóa) | `tax_audit/xml_invoice.py` |

Giao diện có 2 tab: **🔍 Kiểm tra bảng kê** và **🔄 XML hóa đơn → Excel**. Tab XML
cho phép tải nhiều file XML hóa đơn điện tử cùng lúc, trích xuất thành Excel (bảng
hóa đơn + chi tiết hàng hóa); file này dùng luôn được làm *File hóa đơn điện tử* ở
tab kiểm tra.

> **Lưu ý về tra cứu online:** Công cụ **không đăng nhập trực tiếp** vào hệ thống
> TMS hay hệ thống hóa đơn điện tử của cơ quan thuế (việc đó cần tài khoản/chữ ký
> số của bạn). Thay vào đó, bạn **tải file đã export sẵn** từ các hệ thống này lên,
> công cụ sẽ đối chiếu theo MST / số hóa đơn.

## Cài đặt

```bash
pip install -r requirements.txt
```

## Chạy ứng dụng web

```bash
streamlit run app.py
```

Trình duyệt sẽ mở giao diện. Các bước sử dụng:

1. **Tải file lên** — bắt buộc *File dữ liệu chính* (bảng kê hóa đơn cần kiểm tra);
   tùy chọn thêm *File TMS*, *File DS rủi ro*, *File hóa đơn điện tử*.
2. (Tùy chọn) Chỉnh **vị trí cột** ở thanh bên trái nếu file có bố cục khác mặc định.
   Có thể nhập theo **số thứ tự cột** (đếm từ 1) hoặc theo **tên cột**.
3. Nhấn **▶️ Bắt đầu kiểm tra**.
4. Xem các chỉ tiêu tổng hợp, xem trước từng sheet và **tải file Excel kết quả**
   (gồm nhiều sheet: dữ liệu tổng hợp + từng loại cảnh báo).

## Cấu hình cột (`config.py`)

Vị trí cột mặc định được khai báo trong `config.py`. Quy ước:

- **Số nguyên** → vị trí cột đếm từ 1 (vd `9` = cột thứ 9).
- **Chuỗi** → khớp theo tên tiêu đề cột (không phân biệt hoa/thường, khớp gần đúng).

Mặc định đã căn theo bộ file thực tế:

**File dữ liệu chính (Bảng kê hóa đơn mua vào):** tiêu đề ở dòng 9–10, dữ liệu từ dòng 11 (`header_row = 10`).

| Trường | Cột | Nội dung |
|--------|-----|----------|
| MST người bán | 9 | Mã số thuế người bán |
| Số hóa đơn | 5 | Số HĐ (bản có số 0 đầu) |
| Ngày hóa đơn | 7 | Ngày, tháng, năm |
| Giá trị chưa thuế | 14 | Giá trị HHDV mua vào chưa có thuế GTGT |
| Thuế GTGT | 17 | Tiền thuế GTGT |
| Tên hàng hóa | 10 | Tên hàng hóa, dịch vụ |

**File DS rủi ro:** hỗ trợ **nhiều sheet** khác cấu trúc; mỗi sheet khai báo `name`,
`header_row`, `mst_col`, `doc_col` (`doc_col` là **số** → lấy văn bản theo dòng; là
**chữ** → gán nhãn văn bản cố định cho cả sheet, dùng khi văn bản nằm ở tên/tiêu đề sheet).

**File HĐĐT:** khớp theo **tên cột** (export chuẩn `hoadondientu.gdt.gov.vn`). Nếu file
export bị rỗng (không có dòng dữ liệu), công cụ sẽ tự bỏ qua các bước tra HĐĐT và báo cảnh báo.

Tất cả cấu hình này cũng chỉnh được trực tiếp trên giao diện web (thanh bên trái)
mà không cần sửa code.

> **Lưu ý tác vụ tìm HĐ trùng:** khóa trùng gồm *số HĐ + ngày + MST + giá trị chưa thuế*
> đúng theo yêu cầu. Kết quả là **danh sách cần rà soát** (có thể gồm hóa đơn nhiều dòng
> có cùng giá trị), người kiểm tra xác nhận lại trước khi kết luận.

> **Lưu ý tác vụ mặt hàng nghi ngờ:** so khớp **giữ nguyên dấu tiếng Việt** để tránh
> nhầm "tặng" (quà tặng) với "tầng/tăng". Nếu dữ liệu gõ không dấu, thêm biến thể không
> dấu vào danh sách từ khóa trên giao diện.

## File mẫu để thử nghiệm

```bash
python scripts/make_samples.py     # tạo file mẫu trong thư mục samples/
```

Sau đó tải các file trong `samples/` lên ứng dụng để xem toàn bộ luồng hoạt động.

## Cấu trúc dự án

```
├── app.py                 # Giao diện web Streamlit
├── pipeline.py            # Điều phối 12 tác vụ
├── config.py              # Cấu hình vị trí cột & từ khóa
├── tax_audit/             # Thư viện lõi (mỗi tác vụ 1 module)
│   ├── utils.py           # Đọc/ghi Excel, chuẩn hóa MST/số HĐ/ngày/tiền
│   ├── mst.py             # Tác vụ 1, 2
│   ├── taxpayer.py        # Tác vụ 3, 4, 5
│   ├── risk.py            # Tác vụ 6
│   ├── goods.py           # Tác vụ 7
│   ├── invoice.py         # Tác vụ 8, 9
│   └── einvoice.py        # Tác vụ 10, 11, 12
└── scripts/make_samples.py
```
