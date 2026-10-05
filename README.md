# Tool tổng hợp CRM – NTS Hanoi Corp.

**`Tool tổng hợp CRM.html`** là một file HTML duy nhất, mở bằng trình duyệt (Chrome / Edge / Firefox), **chạy hoàn toàn offline**, không gửi dữ liệu đi đâu. Có thể gửi file cho người khác dùng.

**Dùng online (GitHub Pages):** https://pm-ntshn.github.io/Tool_CRM_v2/ – dữ liệu vẫn chỉ xử lý trong trình duyệt, không tải lên máy chủ.

## Cách dùng
1. Mở `Tool tổng hợp CRM.html`, kéo-thả (hoặc bấm chọn) file Excel export Deals từ CRM.
   - File raw: có sheet `1 - Bảng danh mục sản phẩm dự án` → tool tự convert sang template chuẩn (logic giống hệt `reference/convert_deals.py`).
   - Cũng nhận file template đã convert (sheet `Data`).
2. **Ngày chốt số liệu** lấy từ tên file (`..._dd_mm_yy.xlsx` hoặc `export.deals.report.HH.MM.dd.mm.yy.xlsx`), nếu không có thì là hôm nay; sửa được ở thanh tiêu đề.
3. Tìm kiếm / lọc / xem biểu đồ / bảng chi tiết.
4. Tải về:
   - **Tải file template** → `CRM_deals_template_dd_mm_yy.xlsx` (sheet Data, Huong_dan, Danh_muc – giống hệt file của script Python).
   - **Xuất Excel tổng hợp** → `Tổng hợp CRM_ Tháng <tháng>.<năm>.xlsx`, tính trên kết quả đang lọc.

## Tính năng chính
- Ô tìm kiếm không phân biệt dấu, gợi ý theo khách hàng / sale / hãng / reseller / nhóm KH / thương vụ khi đang gõ (phím `/` để mở nhanh, ↑ ↓ Enter).
- Bộ lọc kết hợp: Trạng thái (Active/Won/Lost), Giai đoạn, Sale, Hãng, Reseller, Nhóm KH, Mức giá trị (< 1 tỷ, 1–2, 2–5, 5–10, > 10 tỷ), Loại HĐ (New/Renew), Thời gian (lịch chọn ngày + mốc nhanh; theo Close date / Ngày tạo / Cập nhật cuối), và Thêm: Cờ nhắc, Tình trạng timeline, Loại deal, Chất lượng dữ liệu, Khách hàng. Mỗi lựa chọn hiện số thương vụ khớp.
- Lọc nhanh dạng chip, chip bộ lọc đang áp dụng (bỏ từng cái / xóa tất cả).
- Cờ nhắc theo *Cập nhật lần cuối* so với ngày chốt: Mức 1 > 1 tháng, Mức 2 > 3 tháng, Mức 3 > 6 tháng, Mức 4 > 1 năm (mặc định chỉ tính thương vụ Active, bỏ chọn được). Soạn nội dung nhắc theo từng sale (sao chép / tải .txt).
- KPI, điểm nhấn tự động, 10 biểu đồ tự cập nhật theo bộ lọc (bấm vào cột để lọc; chuyển Số thương vụ ↔ Giá trị).
- Bảng chi tiết: sắp xếp theo cột, ẩn/hiện 25 cột, phân trang, mật độ dòng, xuất CSV; bấm dòng để xem chi tiết (các hãng, vấn đề dữ liệu, các xử lý khi convert).
- Giao diện sáng / tối, dùng được trên điện thoại.
- Mặc định chỉ phân tích *Cơ hội* (ẩn *Lead chiến dịch* theo quy tắc của file template) – bấm “Hiện cả Lead” để xem tất cả.

## File Excel tổng hợp gồm
Tổng quan · Danh sách thương vụ (1 dòng/thương vụ, cột màu cam là cột bổ sung) · Theo Sale · Theo Hãng · Theo Reseller · Theo Nhóm KH · Theo Khách hàng · Nhắc cập nhật · Chi tiết theo hãng (dữ liệu template + 4 cột bổ sung) · Hướng dẫn.

## Cấu trúc mã nguồn
| Đường dẫn | Nội dung |
|---|---|
| `src/converter.js` | Bản JavaScript của `convert_deals.py` + ghi file template |
| `src/app.js`, `src/template.html` | Giao diện |
| `src/vendor/` | SheetJS (xlsx-js-style, có vá nhỏ để ghi *freeze panes*) và Chart.js, nhúng sẵn để chạy offline |
| `build.js` | Gộp mọi thứ thành `Tool tổng hợp CRM.html`: `node build.js` |
| `tests/compare_with_python.js` | Chạy converter JS từ dòng lệnh để đối chiếu với kết quả Python |
| `reference/convert_deals.py` | Script gốc |

Sửa danh mục sale, quy tắc phân nhóm KH, từ khóa hãng… trong `src/converter.js` (giữ đồng bộ với script Python), rồi chạy `node build.js`.
