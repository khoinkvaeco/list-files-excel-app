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

### File Excel kết quả gồm các sheet
- **KQ rà soát** — bảng tổng hợp theo từng loại dấu hiệu, tách **theo từng năm + dòng "Tổng kỳ"** (cột: **Chỉ tiêu · Năm · Giá trị chưa thuế · Thuế GTGT · Số hóa đơn đã bỏ trùng · Giải thích**). Gồm: tổng bảng kê; theo từng trạng thái NNT (TMS); HĐ của DN rủi ro; HĐ kê khai trùng; mặt hàng "Không phục vụ SXKD" / "Quà tặng"; HĐ bị thay thế/điều chỉnh/xóa bỏ/không tìm thấy (HĐĐT); rà soát thuế suất 10%/8%. Số hóa đơn đếm **theo hóa đơn (đã khử trùng dòng)**.
- **Dữ liệu chỉ tiêu** — các dòng chi tiết dính ≥1 chỉ tiêu, gắn nhãn "Các chỉ tiêu dính" + năm + số HĐ + MST + tên hàng + giá trị/thuế + trạng thái, để đối chiếu/kiểm tra ngược số liệu ở sheet KQ.
- **Bảng kê đã xử lý** — cột A–T gốc + 9 cột kết quả **U–AC** tự điền: Loại hàng hóa (nghi ngờ), Trạng thái NNT, Ngày liên quan, Hóa đơn rủi ro, Kê khai trùng, Tra HĐĐT, Tiền thuế trên HĐĐT, **Chênh lệch với BK (AB = AA − S)**, **Check 10%/8%** (đối chiếu thuế suất thực tế với chính sách giảm 8% theo **ngày hóa đơn** — NĐ 15/2022, 44/2023, 94/2023, 72/2024, 174/2024 — kết hợp **nhóm hàng không được giảm** theo Phụ lục I/II/III NĐ 15/2022: viễn thông, tài chính/ngân hàng, chứng khoán, bảo hiểm, bất động sản, kim loại, khai khoáng, than cốc/dầu tinh chế, hóa chất; hàng chịu thuế TTĐB như rượu/bia/thuốc lá/ô tô/xăng/golf/casino…; công nghệ thông tin). Danh mục loại trừ khớp theo **từ khóa mô tả** nên mang tính trợ giúp — nhãn "rà lại" cần người kiểm tra xác nhận theo mã HS/ngành).
- **Dữ liệu tổng hợp** và các sheet cảnh báo riêng (mặt hàng nghi ngờ, HĐ trùng, HĐ sai trạng thái, Đối chiếu thuế theo HĐ, Chênh lệch tiền thuế…).

> Công cụ tự bỏ dòng "Tổng cộng"/dòng trống cuối bảng kê (không có MST và số HĐ) để tổng tiền thuế khớp đúng số liệu gốc.

Giao diện có 3 tab:

- **🔍 Kiểm tra bảng kê** — quy trình 12 tác vụ.
- **🔄 XML hóa đơn → Excel** — tải nhiều file XML hóa đơn điện tử, trích xuất thành Excel (bảng hóa đơn + chi tiết hàng hóa); dùng luôn được làm *File hóa đơn điện tử* ở tab kiểm tra.
- **📁 Gộp file Excel** — gộp nhiều file (mỗi file lấy sheet đầu) hoặc gộp các sheet trong 1 file thành một bảng tổng hợp, kèm tùy chọn: bỏ dòng trống, dùng chung tiêu đề, thêm cột "Nguồn", khử trùng lặp.

**Từ khóa mặt hàng nghi ngờ** được tách 2 nhóm: *Không phục vụ SXKD* (golf, rượu, bia, thuốc lá…) và *Quà tặng* (quà, tặng, biếu). Cột "Loại hàng hóa" (U) trong kết quả hiển thị tên nhóm khớp.

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

Trên giao diện có lựa chọn **"Kiểu chỉ định cột"**:

- **Chữ cái cột Excel** → `"A"`, `"B"`, `"L"`, `"AA"`… đúng như chữ cái cột trên Excel; hoặc số thứ tự (đếm từ 1).
- **Theo tên tiêu đề** → nhập tên cột (khớp gần đúng "chứa"), ví dụ `"Ngày hóa đơn"`, `"Mã số thuế người bán"`, `"chưa có thuế"`. Nhớ đặt đúng **Dòng tiêu đề** của file. Ở chế độ này, chuỗi ngắn không bị hiểu nhầm thành chữ cái cột.

Cả hai kiểu đều **lưu được** (nút 💾) và tự dùng lại lần sau.

Mặc định đã căn theo bộ file thực tế:

**File dữ liệu chính (Bảng kê 01-1/HT):** dữ liệu ở sheet `BKMV`, tiêu đề dòng 10–11, dữ liệu từ dòng 12 (`header_row = 11`).

| Trường | Cột Excel | Nội dung |
|--------|-----|----------|
| MST người bán | L | Mã số thuế người bán |
| Số hóa đơn | F | Số hóa đơn |
| Ngày hóa đơn | G | Ngày, tháng, năm |
| Giá trị chưa thuế | Q | Giá trị HHDV mua vào chưa có thuế GTGT |
| Thuế GTGT | S | Tiền thuế GTGT |
| Tên hàng hóa | M | Tên hàng hóa, dịch vụ |

**File TMS (Danh bạ NNT):** SpreadsheetML, sheet `NNT`, `header_row = 2`; MST cột **D**, Trạng thái tổ chức cột **AS**, Ngày đóng trạng thái cột **BB**.

**File HĐĐT:** sheet `DanhSach`, `header_row = 4`; MST người bán **G**, Số HĐ **D**, Ngày lập **E**, Trạng thái **S**, Tổng tiền thuế **M**.

**File DS rủi ro:** hỗ trợ **nhiều sheet** khác cấu trúc; mỗi sheet khai báo `name`,
`header_row`, `mst_col`, `doc_col` (`doc_col` là **chữ cái cột** → lấy văn bản theo dòng;
là **chuỗi khác** (có dấu cách/số) → gán nhãn văn bản cố định cho cả sheet).

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
