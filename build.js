// Gộp src/* thành 1 file HTML chạy offline: node build.js
const fs = require('fs'), path = require('path');
const R = (p) => fs.readFileSync(path.join(__dirname, 'src', p), 'utf8');
const safe = (js) => js.replace(/<\/script/gi, '<\\/script');
const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(__dirname, 'src', 'logo-nts.png')).toString('base64');
let html = R('template.html');
const parts = {
  __XLSX__: safe(R('vendor/xlsx-js-style.min.js')),
  __CHART__: safe(R('vendor/chart.umd.min.js')),
  __CONVERTER__: safe(R('converter.js')),
  __APP__: safe(R('app.js')),
};
html = html.split('__LOGO__').join(logo);
for (const [k, v] of Object.entries(parts)) html = html.split(k).join(v);
const out = path.join(__dirname, 'Tool tổng hợp CRM.html');
fs.writeFileSync(out, html);
console.log('Đã tạo:', out, (html.length / 1024).toFixed(0) + ' KB');
