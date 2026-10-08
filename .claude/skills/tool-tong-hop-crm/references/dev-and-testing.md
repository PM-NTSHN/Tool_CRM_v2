# Phát triển & kiểm thử

## Kiểm tra converter JS == Python
```bash
pip install pandas openpyxl
python reference/convert_deals.py raw.xlsx py_out.xlsx --ref-date 05/10/2026
node tests/compare_with_python.js raw.xlsx js_out.xlsx 05/10/2026
python3 -c "import pandas as pd
for sh in ['Data','Huong_dan','Danh_muc']: print(sh, pd.read_excel('py_out.xlsx',sheet_name=sh,header=None).equals(pd.read_excel('js_out.xlsx',sheet_name=sh,header=None)))"
```
Lặp lại với một ngày chốt khác (VD 01/03/2027) vì số ngày quá hạn phụ thuộc ngày chốt.

## Đo số liệu trên dữ liệu thật (Node)
```js
const Module = require('module'); const o = Module._load;
Module._load = function (r, ...a) { return /cpexcel\.js$/.test(r) ? {} : o.call(this, r, ...a); }; // bản vendor không kèm cpexcel.js
const XLSX = require('./src/vendor/xlsx-js-style.min.js'), C = require('./src/converter.js');
const res = C.convert(XLSX, XLSX.read(require('fs').readFileSync('raw.xlsx'), { cellDates: true }), C.parseDate('05/10/2026'));
// res.rows = các dòng template; gom theo ID để ra thương vụ
```

## Chạy thử giao diện (Playwright, Chromium có sẵn)
- `chromium.launch({ executablePath: '/opt/pw-browsers/chromium' })`, mở `file:///…/Tool tổng hợp CRM.html`, `setInputFiles('#fileInput', 'raw.xlsx')`, chờ `#app:not(.hidden)`.
- Trạng thái app có ở `window.__crm` (deals, view, filters…) – tiện để set bộ lọc và đọc số.
- Bắt `pageerror` + console error; chụp màn hình full page, dark mode (`#btnTheme`), mobile 390px (kiểm tra `scrollWidth == innerWidth`).
- Tải file: chạy Node với `LANG=C.UTF-8 LC_ALL=C.UTF-8`, nếu không Chromium headless đặt tên file tiếng Việt thành "download" (trình duyệt thật không bị).
- Đọc lại file xuất bằng openpyxl để kiểm tra sheet, freeze pane, tổng cộng.

## Bẫy đã gặp
- **pandas na_values:** `read_excel` coi "NA", "N/A", "null", "None", "#N/A"… là ô trống. Converter JS có danh sách `PD_NA` để khớp – thiếu là lệch cờ "Thiếu thông tin: Reseller".
- **`\b` trong regex:** Python là Unicode, JS chỉ ASCII → dùng hàm `rx()` (lookaround `\p{L}\p{N}_`, cờ `u`).
- **Làm tròn:** Python `format` làm tròn nửa-chẵn → `roundHalfEven()`; nhóm nghìn bằng `.` cho VND, `,` cho USD/tham số.
- **Ngày giờ:** lưu ms "giờ tường" bằng `Date.UTC(...)`, ghi Excel bằng serial tự tính (`toSerial`) → không lệch múi giờ.
- **Bảng:** không đặt `display:-webkit-box` lên `<td>` (vỡ cột) – bọc nội dung trong `<div class="w2">`.
- **Chart.js:** plugin `valueLabels` chỉ register một lần; đổi theme thì `renderCharts(true)` để vẽ lại cả legend HTML.
- **Pages:** cần `.nojekyll` + `index.html` (build.js tự tạo) – nếu không Pages hiển thị README.
- **Freeze pane:** thư viện gốc không hỗ trợ; bản vendor đã vá `Yt("sheetView",c["!freeze"]||null,o)`. Nếu cập nhật thư viện phải vá lại.
