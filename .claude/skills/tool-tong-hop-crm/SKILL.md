---
name: tool-tong-hop-crm
description: Phát triển, sửa lỗi và mở rộng "Tool tổng hợp CRM" của NTS Hanoi Corp (repo PM-NTSHN/Tool_CRM_v2) – tool HTML offline 1 file nạp file export Deals từ CRM, tự convert sang template chuẩn (logic convert_deals.py), lọc/biểu đồ/nhắc cập nhật và xuất Excel "Tổng hợp CRM_ Tháng M.YYYY.xlsx". Dùng skill này bất cứ khi nào người dùng nhắc tới tool tổng hợp CRM, file export.deals.report, CRM_deals_template, convert_deals.py, thương vụ/deal/pipeline theo sale, hãng, reseller, cờ nhắc cập nhật, giá trị theo hãng, hoặc muốn thêm tính năng / sửa số liệu / thêm sale, hãng, quy tắc phân nhóm khách hàng cho tool này – kể cả khi họ chỉ gửi ảnh chụp màn hình tool và nói "số này sai".
---

# Tool tổng hợp CRM – NTS Hanoi Corp.

Tool là **một file HTML duy nhất chạy offline** (không gửi dữ liệu ra ngoài), dùng bởi phòng Quản lý sản phẩm để tổng hợp thương vụ sale đăng ký trên CRM. Bản online: https://pm-ntshn.github.io/Tool_CRM_v2/ (GitHub Pages, nhánh mặc định `claude/funny-hypatia-8rfq4p`, thư mục root).

## Cấu trúc repo

| Đường dẫn | Vai trò |
|---|---|
| `src/converter.js` | Bản JS của `reference/convert_deals.py` (đọc raw → template) + ghi file template. Chạy được cả trong trình duyệt và Node. |
| `src/app.js` | Toàn bộ giao diện: lọc, gợi ý tìm kiếm, KPI, biểu đồ (Chart.js), bảng, drawer chi tiết, soạn nhắc, xuất Excel (xlsx-js-style). |
| `src/template.html` | Khung HTML + CSS (token màu sáng/tối, nhận diện NTS xanh `#0b4ea2` / cam `#f26522`). |
| `src/vendor/` | `xlsx-js-style.min.js` (đã **vá** để ghi freeze pane qua `ws['!freeze']`), `chart.umd.min.js`. |
| `build.js` | `node build.js` → nhúng mọi thứ (kể cả logo base64) thành `Tool tổng hợp CRM.html` **và** `index.html` (trang chủ Pages). |
| `tests/compare_with_python.js` | Chạy converter JS từ dòng lệnh để đối chiếu với Python. |
| `.nojekyll` | Bắt buộc giữ – không có thì Pages dựng README thay vì mở tool. |

Không commit file Excel dữ liệu thật (`.gitignore` chặn `*.xlsx`) – có tên khách hàng.

## Quy trình khi sửa tool

1. **Phân tích trước khi sửa** khi người dùng báo số liệu sai: tìm mọi chỗ tính chỉ số đó (`grep` trong `src/app.js`), đo mức ảnh hưởng trên dữ liệu thật bằng script Node (xem `references/dev-and-testing.md`), trình bày nguyên nhân + bảng trước/sau. Nếu cách sửa phụ thuộc **quy tắc nghiệp vụ** (VD chia giá trị theo hãng) thì đưa ra các phương án kèm số minh họa và để người dùng chọn – đừng tự quyết.
2. Sửa trong `src/`, không sửa trực tiếp file HTML đã build.
3. `node build.js`, rồi kiểm tra cú pháp và chạy thử bằng Playwright (nạp file raw, lọc, xuất Excel, kiểm tra console không lỗi, chụp màn hình xem bố cục).
4. Nếu đụng `converter.js`: bắt buộc chạy so khớp với Python ở ≥ 2 ngày chốt; 3 sheet Data/Huong_dan/Danh_muc phải `equals` 100%.
5. Commit (message tiếng Việt, mô tả trước/sau bằng số), push lên nhánh mặc định; Pages tự deploy sau 1–2 phút (người dùng Ctrl+F5).

## Quy tắc nghiệp vụ đã chốt với người dùng

Chi tiết đầy đủ: `references/business-rules.md`. Tóm tắt những điểm hay bị hỏi:

- **Convert phải giống hệt `convert_deals.py`** (tách BOM, suy hãng, chuẩn hóa tên KH, phân nhóm 12 khối, cờ chất lượng). Thay đổi quy tắc convert → sửa cả Python và JS cho đồng bộ.
- **Giá trị theo hãng (phương án A – người dùng đã chọn):** thương vụ 1 hãng = cả giá trị; nhiều hãng = "Giá trị BOM theo hãng" của hãng đó (không có giá = 0); nhiều hãng mà BOM không có giá nào = cả giá trị cho hãng chính. Khi đang lọc Hãng, *mọi* con số giá trị (KPI, panel, biểu đồ, điểm nhấn, Excel) chỉ lấy phần của hãng đã chọn qua hàm `val(d)`; cột "Giá trị thương vụ" và bộ lọc "Mức giá trị" vẫn theo cả thương vụ.
- **Mặc định chỉ phân tích "Cơ hội"** (ẩn Lead chiến dịch) – có nút "Hiện cả Lead".
- **Cờ nhắc cập nhật** theo *Cập nhật lần cuối* so ngày chốt, tháng dương lịch: Mức 1 > 1 tháng, 2 > 3 tháng, 3 > 6 tháng, 4 > 1 năm; mặc định chỉ tính thương vụ Active.
- **Mức giá trị lọc:** < 1 tỷ · 1–2 · 2–5 · 5–10 · > 10 tỷ (cận dưới thuộc mức trên).
- **Ngày chốt** lấy từ tên file (`_dd_mm_yy.xlsx` hoặc `export.deals.report.HH.MM.dd.mm.yy.xlsx`), không có thì hôm nay; sửa được trên header.
- **Tên file xuất:** `Tổng hợp CRM_ Tháng <tháng>.<năm>.xlsx` (theo tháng của ngày chốt); file template `CRM_deals_template_dd_mm_yy.xlsx`.
- Chân trang cố định: "Dữ liệu được tổng hợp nội bộ từ phòng Quản lý sản phẩm - Công ty Nam Trường Sơn Hà Nội".

## Bẫy kỹ thuật đã gặp (đọc trước khi sửa)

Xem `references/dev-and-testing.md` – gồm: pandas `na_values` ("NA" là ô trống), `\b` Unicode trong regex, làm tròn nửa-chẵn kiểu Python, ngày lưu dạng ms "giờ tường" `Date.UTC`, không đặt `display:-webkit-box` trực tiếp lên `<td>`, `cpexcel.js` khi chạy Node, tên file tiếng Việt khi tải về trong Chromium headless cần `LANG=C.UTF-8`.

## Thêm sale / hãng / quy tắc phân nhóm

- Sale mới: thêm vào `SALE_NAMES` trong `src/converter.js` **và** `reference/convert_deals.py`.
- Hãng mới / tên viết khác: `VENDOR_MAP` + `VENDOR_KEYWORDS` (cả hai file).
- Phân nhóm KH: `SEGMENT_RULES` (theo thứ tự, khớp đầu tiên thì dừng) hoặc `SEGMENT_OVERRIDES` cho KH cụ thể.
Sau đó build + chạy so khớp Python.
