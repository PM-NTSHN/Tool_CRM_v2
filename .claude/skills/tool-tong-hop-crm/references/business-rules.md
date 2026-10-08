# Quy tắc nghiệp vụ – Tool tổng hợp CRM

## 1. Dữ liệu đầu vào
- File raw: export Deals từ CRM, sheet bắt đầu bằng `1 - Bảng danh mục` (tên bị Excel cắt 31 ký tự). Dòng có ID = deal + dòng BOM đầu; dòng ID trống = BOM nối tiếp. BOM chỉ lấy từ 9 cột "Bảng danh mục sản phẩm dự án - …".
- Tool cũng nhận file template đã convert (sheet `Data` có cột ID, Tên thương vụ, Loại deal, Hãng, Cờ chất lượng); khi đó nút "Tải file template" trả lại đúng file gốc.
- Deal test (tên chứa "test", "ko sale", "khong sale") bị loại.

## 2. Template chuẩn (22 cột, sheet Data)
ID · Tên thương vụ · Loại deal · Tên khách hàng · Phân nhóm khách hàng · Sale · Reseller · Giai đoạn · Trạng thái · Giá trị thương vụ · Hãng · Nguồn hãng · Giá trị BOM theo hãng · Thông tin BOM · Phân loại HĐ · Phân loại deal size · Close date · Số ngày quá hạn forecast · Lý do Failed · Ngày tạo · Cập nhật lần cuối · Cờ chất lượng.
- Thương vụ nhiều hãng tách thành nhiều dòng cùng ID; giá trị thương vụ chỉ ghi ở dòng hãng chính (dòng phụ nền xám). Đếm thương vụ = đếm ID duy nhất.
- Hãng chính = hãng có giá trị BOM lớn nhất ("Khác", "Dịch vụ" chỉ là chính khi không còn hãng nào khác).
- Lead chiến dịch = Active + "Đăng kí cơ hội" + không BOM + giá trị 0 + không Close date + không hoạt động.
- Cờ chất lượng: `[VẤN ĐỀ] a; b | [ĐÃ XỬ LÝ] c; d` – tool tách ra để hiển thị và đếm.

## 3. Giá trị theo hãng (phương án A, chốt 10/2026)
Lý do chọn: khớp con số người dùng thấy ở BOM / drawer và ý nghĩa cột "Giá trị BOM theo hãng".
- 1 hãng → cả giá trị thương vụ.
- Nhiều hãng → Giá trị BOM của hãng (hãng không có giá BOM = 0).
- Nhiều hãng, BOM không có giá nào → cả giá trị cho hãng chính.
- Hệ quả đã giải thích với người dùng: cộng tất cả hãng ≠ tổng giá trị thương vụ (dữ liệu 05/10/2026: 1.716,66 tỷ vs 1.648,32 tỷ) vì 66 thương vụ có giá trị CRM lệch tổng BOM > 5%.
- Phương án B (phân bổ giá trị thương vụ theo tỷ trọng BOM) đã bị loại vì số theo hãng khác số BOM người dùng thấy.
- Code: `vendorValues(d)` → `d.vendorValue` (Map hãng → giá trị); `val(d)` trả phần giá trị của các hãng đang lọc, hoặc `d.value` khi không lọc hãng. Mọi tổng giá trị phải đi qua `val`. Sheet "Theo Hãng" luôn dùng `d.vendorValue` (kể cả khi không lọc).
- Khi lọc hãng: dòng kết quả ghi rõ đang tính phần của hãng nào + tổng cả thương vụ; bảng tự hiện cột "Giá trị hãng đang lọc"; drawer có cột "Tính cho hãng"; Excel có cột "Giá trị hãng đang lọc" (Danh sách) và "Giá trị tính cho hãng" (Chi tiết theo hãng).

## 4. Phân loại & cờ
- Mức giá trị (lọc): < 1 tỷ, 1–2, 2–5, 5–10, > 10 tỷ, Chưa có giá trị – theo giá trị cả thương vụ. (Cột "Phân loại deal size" của template giữ 4 mức cũ của Python.)
- Cờ nhắc: so *Cập nhật lần cuối* (nếu trống dùng Ngày tạo) với ngày chốt bằng tháng dương lịch: M1 > 1, M2 > 3, M3 > 6, M4 > 12 tháng. Mặc định chỉ Active (tùy chọn "Chỉ nhắc thương vụ Active"); Won/Lost ghi "Không áp dụng".
- Tình trạng timeline (chỉ Active): Đã qua timeline · Trong 30 ngày tới · 31–90 ngày tới · Sau 90 ngày · Chưa có timeline; Won/Lost = "Đã đóng".
- Tỷ lệ thắng = Won / (Won + Lost) theo số lượng.

## 5. File Excel tổng hợp (theo bộ lọc hiện tại)
Sheets: Tổng quan · Danh sách thương vụ (1 dòng/thương vụ, cột bổ sung tiêu đề cam) · Theo Sale · Theo Hãng · Theo Reseller · Theo Nhóm KH · Theo Khách hàng · Nhắc cập nhật · Chi tiết theo hãng · Hướng dẫn. Font Arial 10, header xanh `1F4E78`, freeze pane + autofilter.

## 6. Giao diện (yêu cầu gốc của người dùng)
Logo NTS góc trái + favicon; tông xanh/cam theo website ntshanoi.com.vn; tìm kiếm có gợi ý khi gõ; bộ lọc dạng nút popover nhiều lựa chọn có số đếm; lịch chọn ngày; chip bộ lọc đang áp dụng; biểu đồ bấm-để-lọc; bảng ẩn/hiện cột; soạn nội dung nhắc theo sale; sáng/tối; dùng được trên điện thoại.
