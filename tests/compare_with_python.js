// Kiểm tra converter.js cho kết quả giống convert_deals.py.
// Cách dùng:
//   python reference/convert_deals.py raw.xlsx py_out.xlsx --ref-date 05/10/2026
//   node tests/compare_with_python.js raw.xlsx js_out.xlsx 05/10/2026
//   -> so sánh py_out.xlsx với js_out.xlsx (VD bằng pandas: read_excel(...).equals(...)) cho 3 sheet Data, Huong_dan, Danh_muc
const fs = require('fs'), path = require('path'), Module = require('module');
// Bản vendor không kèm cpexcel.js (bảng mã cũ, không cần cho .xlsx) -> trả về module rỗng khi được require
const origLoad = Module._load;
Module._load = function (req, ...rest) { return /cpexcel\.js$/.test(req) ? {} : origLoad.call(this, req, ...rest); };
const XLSX = require('../src/vendor/xlsx-js-style.min.js');
const C = require('../src/converter.js');
const [src, dst, ref = null] = process.argv.slice(2);
const wb = XLSX.read(fs.readFileSync(src), { type: 'buffer', cellDates: true });
const refDate = ref ? C.parseDate(ref) : C.refFromName(path.basename(src), Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()));
const res = C.convert(XLSX, wb, refDate);
console.log('Ngày chốt:', C.fmtDate(refDate), '| Thống kê:', JSON.stringify(res.stats));
res.warnings.forEach((w) => console.log(w));
XLSX.writeFile(C.buildTemplateWorkbook(XLSX, res), dst);
console.log('Đã ghi', dst);
