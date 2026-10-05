/* app.js – Giao diện Tool tổng hợp CRM (chạy offline, mọi xử lý trong trình duyệt). */
(function () {
  'use strict';
  var C = window.CRMConvert, X = window.XLSX;
  var DAY = 86400000;

  // ============================== HẰNG SỐ ==============================
  var STAGES = ['Đăng kí cơ hội', 'Tư vấn', 'Phương án kinh doanh', 'Xây dựng hồ sơ', 'Thực hiện hợp đồng'];
  var STATUSES = ['Active', 'Won', 'Lost'];
  var BUCKETS = ['< 1 tỷ', '1 – 2 tỷ', '2 – 5 tỷ', '5 – 10 tỷ', '> 10 tỷ', 'Chưa có giá trị'];
  var FLAGS = ['Bình thường (≤ 1 tháng)', 'Mức 1 · > 1 tháng', 'Mức 2 · > 3 tháng', 'Mức 3 · > 6 tháng', 'Mức 4 · > 1 năm', 'Không áp dụng (đã Won/Lost)', 'Không có ngày cập nhật'];
  var TIMELINES = ['Đã qua timeline', 'Trong 30 ngày tới', '31 – 90 ngày tới', 'Sau 90 ngày', 'Chưa có timeline', 'Đã đóng (Won/Lost)'];
  var QUALITY = ['Có vấn đề cần kiểm tra', 'Không có vấn đề'];
  var NO_RESELLER = '(Không có reseller)', NO_VENDOR = '(Chưa xác định hãng)', NO_CONTRACT = '(Chưa phân loại)', NO_SALE = '(Chưa có sale)';

  // ============================== TIỆN ÍCH ==============================
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  function esc(s) { return s === null || s === undefined ? '' : String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function css(v) { return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
  function alpha(hex, a) {
    hex = hex.replace('#', ''); if (hex.length === 3) hex = hex.split('').map(function (c) { return c + c; }).join('');
    var n = parseInt(hex, 16); return 'rgba(' + (n >> 16 & 255) + ',' + (n >> 8 & 255) + ',' + (n & 255) + ',' + a + ')';
  }
  var nf = new Intl.NumberFormat('vi-VN');
  var nf1 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 1 });
  var nf2 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 });
  function fmtN(n) { return nf.format(n || 0); }
  function fmtVND(v) { return v === null || v === undefined ? '' : nf.format(Math.round(v)); }
  function fmtShort(v) {
    if (v === null || v === undefined) return '–';
    var a = Math.abs(v);
    if (a >= 1e9) return nf2.format(v / 1e9) + ' tỷ';
    if (a >= 1e6) return nf1.format(v / 1e6) + ' tr';
    return nf.format(Math.round(v));
  }
  function pct(a, b) { return b ? nf1.format(a * 100 / b) + '%' : '–'; }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function fmtD(t) { if (t === null || t === undefined) return ''; var d = new Date(t); return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear(); }
  function fmtDT(t) { if (t === null || t === undefined) return ''; var d = new Date(t); return fmtD(t) + ' ' + pad2(d.getUTCHours()) + ':' + pad2(d.getUTCMinutes()); }
  function isoD(t) { if (t === null || t === undefined) return ''; var d = new Date(t); return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()); }
  function fromIso(s) { if (!s) return null; var p = s.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]); }
  function addMonths(t, m) {
    var d = new Date(t), y = d.getUTCFullYear(), mo = d.getUTCMonth() + m, day = d.getUTCDate();
    var last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
    return Date.UTC(y, mo, Math.min(day, last));
  }
  function todayUTC() { var d = new Date(); return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()); }
  function fold(s) { return C.fold(s === null || s === undefined ? '' : s); }
  function debounce(f, ms) { var t; return function () { var a = arguments, th = this; clearTimeout(t); t = setTimeout(function () { f.apply(th, a); }, ms); }; }
  function store(k, v) { try { if (v === undefined) { var s = localStorage.getItem('ntscrm.' + k); return s === null ? undefined : JSON.parse(s); } localStorage.setItem('ntscrm.' + k, JSON.stringify(v)); } catch (e) { return undefined; } }
  function toast(msg, err) {
    var t = document.createElement('div'); t.className = 'toast' + (err ? ' err' : '');
    t.innerHTML = '<svg class="i"><use href="#i-' + (err ? 'info' : 'check') + '"/></svg>' + esc(msg);
    document.body.appendChild(t); setTimeout(function () { t.remove(); }, err ? 6500 : 2600);
  }
  function median(a) { if (!a.length) return null; a = a.slice().sort(function (x, y) { return x - y; }); var m = a.length >> 1; return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2; }
  function sum(a, f) { return a.reduce(function (s, x) { var v = f(x); return s + (v || 0); }, 0); }
  function sortVi(a, b) { return String(a).localeCompare(String(b), 'vi'); }

  // ============================== TRẠNG THÁI ==============================
  var S = {
    file: null, fileBytes: null, wb: null, res: null, ref: null, deals: [], view: [],
    filters: {}, q: '', date: { field: 'close', from: null, to: null },
    onlyActiveRemind: store('onlyActive') !== false,
    metric: store('metric') || 'count',
    sort: { k: 'value', dir: -1 }, page: 1, pageSize: store('pageSize') || 50,
    cols: null, dense: !!store('dense'), charts: {}
  };

  // ============================== MÔ HÌNH DỮ LIỆU ==============================
  function bucketOf(v) {
    if (v === null || v === undefined || !(v > 0)) return 'Chưa có giá trị';
    if (v < 1e9) return '< 1 tỷ'; if (v < 2e9) return '1 – 2 tỷ'; if (v < 5e9) return '2 – 5 tỷ'; if (v < 10e9) return '5 – 10 tỷ';
    return '> 10 tỷ';
  }
  function parseFlags(s) {
    var issues = [], done = [];
    (s || '').split(' | ').forEach(function (p) {
      if (p.indexOf('[VẤN ĐỀ] ') === 0) issues = p.slice(9).split('; ');
      else if (p.indexOf('[ĐÃ XỬ LÝ] ') === 0) done = p.slice(11).split('; ');
    });
    return { issues: issues, done: done };
  }
  function buildDeals(res) {
    var map = new Map(), deals = [];
    res.rows.forEach(function (r) {
      var d = map.get(r.ID);
      if (!d) {
        var fl = parseFlags(r['Cờ chất lượng']);
        d = { id: r.ID, name: r['Tên thương vụ'] || '', type: r['Loại deal'], customer: r['Tên khách hàng'] || '', segment: r['Phân nhóm khách hàng'] || 'Khác',
          sale: r['Sale'] || NO_SALE, reseller: r['Reseller'], stage: r['Giai đoạn'], status: r['Trạng thái'], value: r['Giá trị thương vụ'],
          bomText: r[C.BOM_TXT], contract: r['Phân loại HĐ'], dealSize: r['Phân loại deal size'], close: r['Close date'], failed: r['Lý do Failed'],
          created: r['Ngày tạo'], updated: r['Cập nhật lần cuối'], quality: r['Cờ chất lượng'], issues: fl.issues, done: fl.done, vendors: [], rows: [] };
        map.set(r.ID, d); deals.push(d);
      }
      d.vendors.push({ v: r['Hãng'], src: r['Nguồn hãng'], bom: r['Giá trị BOM theo hãng'] });
      d.rows.push(r);
    });
    deals.forEach(function (d) {
      d.vendorNames = d.vendors.map(function (x) { return x.v || NO_VENDOR; }).filter(function (v, i, a) { return a.indexOf(v) === i; });
      d.mainVendor = d.vendors[0].v || NO_VENDOR;
      d.bucket = bucketOf(d.value);
      d.qualityLbl = d.issues.length ? QUALITY[0] : QUALITY[1];
      d.text = fold([d.id, d.name, d.customer, d.sale, d.reseller, d.vendorNames.join(' '), d.segment, d.stage, d.status, d.contract, d.bomText, d.failed].join(' | '));
    });
    return deals;
  }
  function derive() {
    var ref = S.ref, t1 = addMonths(ref, -1), t3 = addMonths(ref, -3), t6 = addMonths(ref, -6), t12 = addMonths(ref, -12);
    S.deals.forEach(function (d) {
      var u = d.updated !== null && d.updated !== undefined ? d.updated : d.created;
      d.staleDays = u === null || u === undefined ? null : Math.max(0, Math.floor((ref - u) / DAY));
      d.level = u === null || u === undefined ? null : u < t12 ? 4 : u < t6 ? 3 : u < t3 ? 2 : u < t1 ? 1 : 0;
      var applies = d.status === 'Active' || !S.onlyActiveRemind;
      d.remind = applies && d.level >= 1;
      d.flag = !applies ? FLAGS[5] : d.level === null ? FLAGS[6] : FLAGS[d.level];
      d.ageDays = d.created === null || d.created === undefined ? null : Math.max(0, Math.floor((ref - d.created) / DAY));
      d.overdue = d.close !== null && d.close !== undefined && d.status === 'Active' && d.close < ref ? Math.floor((ref - d.close) / DAY) : null;
      if (d.status !== 'Active') d.timeline = TIMELINES[5];
      else if (d.close === null || d.close === undefined) d.timeline = TIMELINES[4];
      else if (d.close < ref) d.timeline = TIMELINES[0];
      else { var dd = Math.floor((d.close - ref) / DAY); d.timeline = dd <= 30 ? TIMELINES[1] : dd <= 90 ? TIMELINES[2] : TIMELINES[3]; }
    });
  }

  // ============================== BỘ LỌC ==============================
  var FILTERS = [
    { k: 'status', label: 'Trạng thái', get: function (d) { return d.status; }, order: STATUSES, btn: true },
    { k: 'stage', label: 'Giai đoạn', get: function (d) { return d.stage; }, order: STAGES, btn: true },
    { k: 'sale', label: 'Sale', get: function (d) { return d.sale; }, search: true, btn: true },
    { k: 'vendor', label: 'Hãng', multi: function (d) { return d.vendorNames; }, search: true, btn: true },
    { k: 'reseller', label: 'Reseller', get: function (d) { return d.reseller || NO_RESELLER; }, search: true, btn: true },
    { k: 'segment', label: 'Nhóm KH', get: function (d) { return d.segment; }, order: C.SEGMENTS, btn: true },
    { k: 'bucket', label: 'Giá trị', get: function (d) { return d.bucket; }, order: BUCKETS, btn: true },
    { k: 'contract', label: 'Loại HĐ', get: function (d) { return d.contract || NO_CONTRACT; }, btn: true },
    { k: 'flag', label: 'Cờ nhắc', get: function (d) { return d.flag; }, order: FLAGS, more: true },
    { k: 'timeline', label: 'Tình trạng timeline', get: function (d) { return d.timeline; }, order: TIMELINES, more: true },
    { k: 'type', label: 'Loại deal', get: function (d) { return d.type; }, more: true },
    { k: 'quality', label: 'Chất lượng dữ liệu', get: function (d) { return d.qualityLbl; }, order: QUALITY, more: true },
    { k: 'customer', label: 'Khách hàng', get: function (d) { return d.customer; }, search: true, more: true }
  ];
  var FBY = {}; FILTERS.forEach(function (f) { FBY[f.k] = f; });
  var DATE_FIELDS = { close: 'Close date (Timeline dự án)', created: 'Ngày tạo', updated: 'Cập nhật lần cuối' };

  function valuesOf(f, d) { return f.multi ? f.multi(d) : [f.get(d)]; }
  function passFilter(f, d) {
    var set = S.filters[f.k]; if (!set || !set.size) return true;
    var vs = valuesOf(f, d); for (var i = 0; i < vs.length; i++) if (set.has(vs[i])) return true; return false;
  }
  function qTokens() { return fold(S.q).split(/\s+/).filter(Boolean); }
  function passSearch(d, toks) { for (var i = 0; i < toks.length; i++) if (d.text.indexOf(toks[i]) < 0) return false; return true; }
  function passDate(d) {
    var D = S.date; if (D.from === null && D.to === null) return true;
    var t = d[D.field]; if (t === null || t === undefined) return false;
    if (D.from !== null && t < D.from) return false;
    if (D.to !== null && t >= D.to + DAY) return false;
    return true;
  }
  function filtered(exceptKey) {
    var toks = qTokens();
    return S.deals.filter(function (d) {
      if (toks.length && !passSearch(d, toks)) return false;
      if (exceptKey !== 'date' && !passDate(d)) return false;
      for (var i = 0; i < FILTERS.length; i++) if (FILTERS[i].k !== exceptKey && !passFilter(FILTERS[i], d)) return false;
      return true;
    });
  }
  function setFilter(k, vals) { if (!vals || !vals.length) delete S.filters[k]; else S.filters[k] = new Set(vals); }
  function toggleValue(k, v) {
    var set = S.filters[k] || new Set();
    if (set.has(v)) set.delete(v); else set.add(v);
    if (set.size) S.filters[k] = set; else delete S.filters[k];
    S.page = 1; refresh();
  }
  function activeCount() {
    var n = Object.keys(S.filters).length + (S.q ? 1 : 0) + (S.date.from !== null || S.date.to !== null ? 1 : 0);
    return n;
  }
  function clearAll() { S.filters = {}; S.q = ''; $('#q').value = ''; $('#searchBox').classList.remove('has'); S.date = { field: S.date.field, from: null, to: null }; S.page = 1; refresh(); }

  // ============================== LOAD FILE ==============================
  function showLoading(on, msg) {
    var el = $('#loadingEl');
    if (on) {
      if (!el) { el = document.createElement('div'); el.id = 'loadingEl'; el.className = 'loading'; document.body.appendChild(el); }
      el.innerHTML = '<div class="card-l"><div class="spin"></div>' + esc(msg || 'Đang xử lý…') + '</div>';
    } else if (el) el.remove();
  }
  function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { toast('Vui lòng chọn file Excel (.xlsx)', true); return; }
    showLoading(true, 'Đang đọc và convert “' + file.name + '”…');
    var rd = new FileReader();
    rd.onload = function (e) {
      setTimeout(function () {
        try {
          var bytes = new Uint8Array(e.target.result);
          var wb = X.read(bytes, { type: 'array', cellDates: true });
          S.file = file.name; S.fileBytes = bytes; S.wb = wb;
          var ref = C.refFromName(file.name, todayUTC());
          if (C.findSourceSheet(wb)) { S.mode = 'raw'; S.res = C.convert(X, wb, ref); }
          else if (C.isTemplateWorkbook(X, wb)) { S.mode = 'template'; S.res = C.loadTemplate(X, wb, ref); }
          else throw new Error('Không nhận dạng được file. Cần file export CRM có sheet “1 - Bảng danh mục sản phẩm dự án” hoặc file template chuẩn có sheet “Data”. Các sheet có: ' + wb.SheetNames.join(', '));
          S.ref = S.res.refDate;
          S.deals = buildDeals(S.res);
          S.filters = {}; S.q = ''; S.page = 1; S.date = { field: 'close', from: null, to: null };
          // Mặc định chỉ phân tích "Cơ hội" (theo quy tắc của file template) nếu có Lead chiến dịch
          if (S.deals.some(function (d) { return d.type === 'Lead chiến dịch'; }) && S.deals.some(function (d) { return d.type === 'Cơ hội'; })) setFilter('type', ['Cơ hội']);
          derive();
          enterApp();
          toast(S.mode === 'raw' ? 'Đã convert ' + fmtN(S.res.stats.deals) + ' thương vụ từ file raw' : 'Đã nạp file template: ' + fmtN(S.res.stats.deals) + ' thương vụ');
        } catch (err) {
          console.error(err);
          toast(err.message || String(err), true);
        } finally { showLoading(false); }
      }, 30);
    };
    rd.onerror = function () { showLoading(false); toast('Không đọc được file', true); };
    rd.readAsArrayBuffer(file);
  }
  function changeRef(t) {
    if (t === null || isNaN(t)) return;
    if (S.mode === 'raw') {
      try { S.res = C.convert(X, S.wb, t); S.deals = buildDeals(S.res); } catch (e) { toast(e.message, true); return; }
    } else S.res.refDate = t;
    S.ref = t; derive(); renderMeta(); refresh();
    toast('Đã tính lại theo ngày chốt ' + fmtD(t));
  }

  // ============================== KHUNG ỨNG DỤNG ==============================
  function enterApp() {
    $('#landing').classList.add('hidden'); $('#app').classList.remove('hidden');
    ['#btnNew', '#btnTpl', '#btnExport'].forEach(function (s) { $(s).classList.remove('hidden'); });
    if (!S.cols) S.cols = store('cols') || COLS.filter(function (c) { return c.def; }).map(function (c) { return c.k; });
    renderMeta(); renderFilterButtons(); renderCharts(true); renderRecon(); refresh();
    $('#q').value = '';
  }
  function renderMeta() {
    var st = S.res.stats;
    $('#metaLine').innerHTML =
      '<span class="pill" title="' + esc(S.file) + '"><svg class="i sm"><use href="#i-file"/></svg>' + esc(S.file.length > 46 ? S.file.slice(0, 43) + '…' : S.file) + '</span>' +
      '<span class="pill">' + (S.mode === 'raw' ? 'Đã convert từ raw' : 'File template') + ' · ' + fmtN(st.deals) + ' thương vụ</span>' +
      '<label class="pill" title="Ngày chốt số liệu: dùng tính cờ nhắc, quá hạn timeline"><svg class="i sm"><use href="#i-cal"/></svg>Ngày chốt: <input type="date" id="refDate" value="' + isoD(S.ref) + '"></label>';
    $('#refDate').addEventListener('change', function (e) { changeRef(fromIso(e.target.value)); });
  }

  function refresh() {
    S.view = filtered();
    renderChips(); renderQuick(); renderResultLine(); renderAlert(); renderDuo(); renderKpis(); renderInsights(); updateCharts(); renderTable();
    $$('.fb[data-k]').forEach(function (el) {
      var k = el.dataset.k, n;
      if (k === 'date') n = (S.date.from !== null || S.date.to !== null) ? 1 : 0;
      else if (k === 'more') n = FILTERS.filter(function (f) { return f.more && S.filters[f.k]; }).reduce(function (s, f) { return s + S.filters[f.k].size; }, 0);
      else n = S.filters[k] ? S.filters[k].size : 0;
      el.classList.toggle('on', n > 0);
      var c = el.querySelector('.cnt'); c.textContent = n; c.classList.toggle('hidden', !n);
    });
  }

  // ---------- nút bộ lọc + popover ----------
  var openPop = null;
  function closePop() { if (openPop) { openPop.remove(); openPop = null; } }
  function renderFilterButtons() {
    var html = FILTERS.filter(function (f) { return f.btn; }).map(function (f) {
      return '<div class="fb" data-k="' + f.k + '"><button type="button">' + esc(f.label) + '<span class="cnt hidden">0</span><svg class="i car"><use href="#i-chev"/></svg></button></div>';
    }).join('');
    html += '<div class="fb" data-k="date"><button type="button"><svg class="i sm"><use href="#i-cal"/></svg>Thời gian<span class="cnt hidden">0</span><svg class="i car"><use href="#i-chev"/></svg></button></div>';
    html += '<div class="fb" data-k="more"><button type="button">Thêm bộ lọc<span class="cnt hidden">0</span><svg class="i car"><use href="#i-chev"/></svg></button></div>';
    $('#fbtns').innerHTML = html;
    $$('#fbtns .fb > button').forEach(function (b) {
      b.addEventListener('click', function (e) {
        e.stopPropagation();
        var fb = b.parentNode, k = fb.dataset.k;
        if (openPop && openPop.parentNode === fb) { closePop(); return; }
        closePop();
        if (k === 'date') openDatePop(fb); else if (k === 'more') openMorePop(fb); else openListPop(fb, [FBY[k]]);
      });
    });
  }
  function optionCounts(f) {
    var base = filtered(f.k), cnt = new Map();
    base.forEach(function (d) { valuesOf(f, d).forEach(function (v) { cnt.set(v, (cnt.get(v) || 0) + 1); }); });
    var all = new Set(); S.deals.forEach(function (d) { valuesOf(f, d).forEach(function (v) { all.add(v); }); });
    var vals = Array.from(all);
    if (f.order) vals.sort(function (a, b) { var ia = f.order.indexOf(a), ib = f.order.indexOf(b); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || sortVi(a, b); });
    else vals.sort(function (a, b) { return (cnt.get(b) || 0) - (cnt.get(a) || 0) || sortVi(a, b); });
    return { vals: vals, cnt: cnt };
  }
  function dotColor(k, v) {
    if (k === 'status') return css({ Active: '--s1', Won: '--good', Lost: '--critical' }[v] || '--ink-3');
    if (k === 'flag') return flagColor(FLAGS.indexOf(v));
    return null;
  }
  function listHtml(f, filterText) {
    var oc = optionCounts(f), set = S.filters[f.k] || new Set(), ft = fold(filterText || '');
    return oc.vals.filter(function (v) { return !ft || fold(v).indexOf(ft) >= 0; }).map(function (v) {
      var n = oc.cnt.get(v) || 0, dc = dotColor(f.k, v);
      return '<label class="opt' + (n ? '' : ' zero') + '"><input type="checkbox" data-k="' + f.k + '" value="' + esc(v) + '"' + (set.has(v) ? ' checked' : '') + '>' +
        (dc ? '<span class="dot" style="background:' + dc + '"></span>' : '') + '<span class="l" title="' + esc(v) + '">' + esc(v) + '</span><span class="c">' + fmtN(n) + '</span></label>';
    }).join('') || '<div class="sg-empty">Không có lựa chọn phù hợp</div>';
  }
  function openListPop(fb, fs, right) {
    var pop = document.createElement('div'); pop.className = 'pop' + (right ? ' right' : '');
    var single = fs.length === 1, f0 = fs[0];
    var head = single ? '<div class="pop-h">' + (f0.search ? '<input type="search" placeholder="Tìm ' + esc(f0.label.toLowerCase()) + '…">' : '<b>' + esc(f0.label) + '</b>') +
      '<div class="row"><button data-act="all">Chọn tất cả đang hiện</button><button data-act="none">Bỏ chọn</button></div></div>' : '';
    var body = fs.map(function (f) {
      return (single ? '' : '<div class="sg-h">' + esc(f.label) + '</div>') + '<div class="lst" data-k="' + f.k + '">' + listHtml(f) + '</div>';
    }).join('');
    pop.innerHTML = head + '<div class="pop-b">' + body + '</div><div class="pop-f"><span class="muted" style="font-size:12px;align-self:center">Số bên phải = số thương vụ khớp</span><button class="btn sm" data-act="close">Xong</button></div>';
    fb.appendChild(pop); openPop = pop;
    var r = pop.getBoundingClientRect(); if (r.right > window.innerWidth - 8) pop.classList.add('right');
    pop.addEventListener('click', function (e) { e.stopPropagation(); });
    pop.addEventListener('change', function (e) {
      var cb = e.target; if (!cb.dataset.k) return;
      var set = S.filters[cb.dataset.k] || new Set();
      if (cb.checked) set.add(cb.value); else set.delete(cb.value);
      if (set.size) S.filters[cb.dataset.k] = set; else delete S.filters[cb.dataset.k];
      S.page = 1; refresh(); refreshPopLists();
    });
    var inp = pop.querySelector('input[type=search]');
    if (inp) { inp.focus(); inp.addEventListener('input', function () { refreshPopLists(); }); }
    function refreshPopLists() {
      $$('.lst', pop).forEach(function (l) { l.innerHTML = listHtml(FBY[l.dataset.k], inp ? inp.value : ''); });
    }
    pop.addEventListener('click', function (e) {
      var a = e.target.closest('[data-act]'); if (!a) return;
      if (a.dataset.act === 'close') { closePop(); return; }
      if (a.dataset.act === 'all') { setFilter(f0.k, $$('input[type=checkbox]', pop).map(function (c) { return c.value; })); }
      if (a.dataset.act === 'none') delete S.filters[f0.k];
      S.page = 1; refresh(); refreshPopLists();
    });
  }
  function openMorePop(fb) { openListPop(fb, FILTERS.filter(function (f) { return f.more && f.k !== 'customer'; }).concat(S.filters.customer ? [FBY.customer] : []), true); }

  function datePresets() {
    var ref = S.ref, d = new Date(ref), y = d.getUTCFullYear(), m = d.getUTCMonth(), q = Math.floor(m / 3);
    return [
      ['Tháng này', Date.UTC(y, m, 1), Date.UTC(y, m + 1, 0)],
      ['Tháng tới', Date.UTC(y, m + 1, 1), Date.UTC(y, m + 2, 0)],
      ['Quý này', Date.UTC(y, q * 3, 1), Date.UTC(y, q * 3 + 3, 0)],
      ['Quý tới', Date.UTC(y, q * 3 + 3, 1), Date.UTC(y, q * 3 + 6, 0)],
      ['Năm nay', Date.UTC(y, 0, 1), Date.UTC(y, 11, 31)],
      ['Năm sau', Date.UTC(y + 1, 0, 1), Date.UTC(y + 1, 11, 31)],
      ['30 ngày tới', ref, ref + 30 * DAY],
      ['90 ngày tới', ref, ref + 90 * DAY],
      ['30 ngày qua', ref - 30 * DAY, ref],
      ['Năm trước', Date.UTC(y - 1, 0, 1), Date.UTC(y - 1, 11, 31)]
    ];
  }
  function openDatePop(fb) {
    var pop = document.createElement('div'); pop.className = 'pop date';
    var D = S.date;
    pop.innerHTML = '<div class="pop-h"><b>Lọc theo thời gian</b></div>' +
      '<div class="dgrid"><label style="grid-column:1/-1">Trường ngày<select id="dField">' + Object.keys(DATE_FIELDS).map(function (k) { return '<option value="' + k + '"' + (D.field === k ? ' selected' : '') + '>' + DATE_FIELDS[k] + '</option>'; }).join('') + '</select></label>' +
      '<label>Từ ngày<input type="date" id="dFrom" value="' + isoD(D.from) + '"></label><label>Đến ngày<input type="date" id="dTo" value="' + isoD(D.to) + '"></label></div>' +
      '<div class="presets">' + datePresets().map(function (p, i) { return '<button data-p="' + i + '">' + p[0] + '</button>'; }).join('') + '</div>' +
      '<div class="pop-f"><button class="btn sm ghost" data-act="clear">Xóa lọc thời gian</button><button class="btn sm" data-act="close">Xong</button></div>';
    fb.appendChild(pop); openPop = pop;
    var r = pop.getBoundingClientRect(); if (r.right > window.innerWidth - 8) pop.classList.add('right');
    pop.addEventListener('click', function (e) {
      e.stopPropagation();
      var p = e.target.closest('[data-p]'), a = e.target.closest('[data-act]');
      if (p) { var pr = datePresets()[+p.dataset.p]; S.date.from = pr[1]; S.date.to = pr[2]; $('#dFrom', pop).value = isoD(pr[1]); $('#dTo', pop).value = isoD(pr[2]); S.page = 1; refresh(); }
      if (a && a.dataset.act === 'clear') { S.date.from = S.date.to = null; $('#dFrom', pop).value = ''; $('#dTo', pop).value = ''; S.page = 1; refresh(); }
      if (a && a.dataset.act === 'close') closePop();
    });
    pop.addEventListener('change', function () {
      S.date.field = $('#dField', pop).value; S.date.from = fromIso($('#dFrom', pop).value); S.date.to = fromIso($('#dTo', pop).value);
      S.page = 1; refresh();
    });
    setTimeout(function () { var i = $('#dFrom', pop); if (i && i.showPicker && !S.date.from) { try { i.focus(); } catch (e) { /* bỏ qua */ } } }, 30);
  }
  document.addEventListener('click', function (e) { if (openPop && !openPop.parentNode.contains(e.target)) closePop(); if (!$('#searchBox').contains(e.target)) hideSuggest(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closePop(); hideSuggest(); closeOverlay(); }
    if (e.key === '/' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA' && !$('#app').classList.contains('hidden')) { e.preventDefault(); $('#q').focus(); }
  });

  // ---------- chips ----------
  function renderChips() {
    var parts = [];
    if (S.q) parts.push(['q', null, 'Tìm', S.q]);
    FILTERS.forEach(function (f) { var s = S.filters[f.k]; if (s) s.forEach(function (v) { parts.push(['f', f.k, f.label, v]); }); });
    if (S.date.from !== null || S.date.to !== null) parts.push(['date', null, DATE_FIELDS[S.date.field].replace(/ \(.*\)/, ''), (S.date.from !== null ? fmtD(S.date.from) : '…') + ' → ' + (S.date.to !== null ? fmtD(S.date.to) : '…')]);
    $('#chips').innerHTML = parts.map(function (p, i) {
      return '<span class="chip"><b>' + esc(p[2]) + ':</b><span title="' + esc(p[3]) + '">' + esc(p[3]) + '</span><button data-i="' + i + '" title="Bỏ lọc"><svg class="i sm"><use href="#i-x"/></svg></button></span>';
    }).join('') + (parts.length > 1 ? '<button class="chip-clear" id="clearAll">Xóa tất cả bộ lọc</button>' : '');
    $$('#chips .chip button').forEach(function (b) {
      b.addEventListener('click', function () {
        var p = parts[+b.dataset.i];
        if (p[0] === 'q') { S.q = ''; $('#q').value = ''; $('#searchBox').classList.remove('has'); }
        else if (p[0] === 'date') { S.date.from = S.date.to = null; }
        else { S.filters[p[1]].delete(p[3]); if (!S.filters[p[1]].size) delete S.filters[p[1]]; }
        S.page = 1; refresh();
      });
    });
    var ca = $('#clearAll'); if (ca) ca.addEventListener('click', clearAll);
  }

  // ---------- lọc nhanh ----------
  var QUICK = [
    { l: 'Active', k: 'status', v: ['Active'], c: '--s1' },
    { l: 'Won', k: 'status', v: ['Won'], c: '--good' },
    { l: 'Lost', k: 'status', v: ['Lost'], c: '--critical' },
    { l: 'Cần nhắc (≥ Mức 1)', k: 'flag', v: FLAGS.slice(1, 5), c: '--warn' },
    { l: 'Cờ ≥ Mức 2', k: 'flag', v: FLAGS.slice(2, 5), c: '--serious' },
    { l: 'Đã qua timeline', k: 'timeline', v: [TIMELINES[0]], c: '--critical' },
    { l: 'Close ≤ 30 ngày', k: 'timeline', v: [TIMELINES[1]], c: '--s3' },
    { l: 'Deal ≥ 5 tỷ', k: 'bucket', v: ['5 – 10 tỷ', '> 10 tỷ'], c: '--s7' },
    { l: 'Renew', k: 'contract', v: ['Renew'], c: '--s4' },
    { l: 'Có vấn đề dữ liệu', k: 'quality', v: [QUALITY[0]], c: '--s2' }
  ];
  function quickOn(q) { var s = S.filters[q.k]; return !!s && s.size === q.v.length && q.v.every(function (v) { return s.has(v); }); }
  function renderQuick() {
    $('#quick').innerHTML = '<span class="ttl">Lọc nhanh</span>' + QUICK.map(function (q, i) {
      var base = filtered(q.k), n = base.filter(function (d) { var vs = valuesOf(FBY[q.k], d); return vs.some(function (v) { return q.v.indexOf(v) >= 0; }); }).length;
      return '<button class="qc' + (quickOn(q) ? ' on' : '') + '" data-i="' + i + '"><span class="d" style="background:' + css(q.c) + '"></span>' + esc(q.l) + ' <span class="n">' + fmtN(n) + '</span></button>';
    }).join('');
    $$('#quick .qc').forEach(function (b) {
      b.addEventListener('click', function () { var q = QUICK[+b.dataset.i]; setFilter(q.k, quickOn(q) ? [] : q.v); S.page = 1; refresh(); });
    });
  }
  function renderResultLine() {
    var v = S.view, total = S.deals.length, hiddenLead = 0;
    if (S.filters.type && S.filters.type.size === 1 && S.filters.type.has('Cơ hội')) {
      var s = S.filters.type; delete S.filters.type; hiddenLead = filtered().filter(function (d) { return d.type === 'Lead chiến dịch'; }).length; S.filters.type = s;
    }
    $('#resultLine').innerHTML = 'Đang hiển thị <strong>' + fmtN(v.length) + ' / ' + fmtN(total) + '</strong> thương vụ · tổng giá trị <strong>' + fmtShort(sum(v, function (d) { return d.value; })) + '</strong>' +
      (activeCount() ? '' : ' · <span class="muted">chưa áp dụng bộ lọc nào</span>') +
      (hiddenLead ? ' · <span class="muted">đang ẩn ' + fmtN(hiddenLead) + ' Lead chiến dịch (chưa có thông tin cơ hội)</span> <button class="linkbtn" id="showLead">Hiện cả Lead</button>' : '');
    var sl = $('#showLead'); if (sl) sl.addEventListener('click', function () { delete S.filters.type; refresh(); });
  }

  // ---------- cảnh báo cờ nhắc ----------
  function flagColor(i) { return css(['--good', '--warn', '--serious', '--critical', '--severe', '--ink-3', '--axis'][i] || '--ink-3'); }
  function renderAlert() {
    var v = S.view, rem = v.filter(function (d) { return d.remind; });
    var lv = [1, 2, 3, 4].map(function (l) { return rem.filter(function (d) { return d.level === l; }).length; });
    var oldest = rem.slice().sort(function (a, b) { return b.staleDays - a.staleDays; })[0];
    var box = $('#alertBox');
    box.innerHTML = '<div class="alert"><div class="ic"><svg class="i"><use href="#i-alarm"/></svg></div><div class="txt">' +
      (rem.length ? '<div class="big"><b>' + fmtN(rem.length) + '</b> thương vụ đã <b style="font-size:inherit">hơn 1 tháng</b> chưa được cập nhật (tính đến ' + fmtD(S.ref) + ').' +
        (oldest ? ' Lâu nhất: <a href="#" data-open="' + oldest.id + '">#' + oldest.id + ' ' + esc(oldest.name.length > 50 ? oldest.name.slice(0, 48) + '…' : oldest.name) + '</a> — <b style="font-size:inherit">' + fmtN(oldest.staleDays) + ' ngày</b>.' : '') + '</div>'
        : '<div class="big">Không có thương vụ nào quá 1 tháng chưa cập nhật trong kết quả lọc.</div>') +
      '<div class="lvls">' + [1, 2, 3, 4].map(function (l) {
        var on = S.filters.flag && S.filters.flag.size === 1 && S.filters.flag.has(FLAGS[l]);
        return '<span class="lvl' + (on ? ' on' : '') + '" data-l="' + l + '"><i style="background:' + flagColor(l) + '"></i>' + esc(FLAGS[l]) + ' <b>' + fmtN(lv[l - 1]) + '</b></span>';
      }).join('') + '</div></div>' +
      '<div class="acts"><label class="tg"><input type="checkbox" id="onlyAct"' + (S.onlyActiveRemind ? ' checked' : '') + '> Chỉ nhắc thương vụ Active</label>' +
      '<button class="btn blue sm" id="btnSeeList"><svg class="i sm"><use href="#i-list"/></svg>Xem danh sách</button>' +
      '<button class="btn primary sm" id="btnCompose"><svg class="i sm"><use href="#i-mail"/></svg>Soạn nội dung nhắc</button></div></div>';
    $$('.lvl', box).forEach(function (el) { el.addEventListener('click', function () { var f = FLAGS[+el.dataset.l]; var on = el.classList.contains('on'); setFilter('flag', on ? [] : [f]); S.page = 1; refresh(); }); });
    $('#onlyAct').addEventListener('change', function (e) { S.onlyActiveRemind = e.target.checked; store('onlyActive', S.onlyActiveRemind); derive(); refresh(); });
    $('#btnSeeList').addEventListener('click', function () { setFilter('flag', FLAGS.slice(1, 5)); S.sort = { k: 'staleDays', dir: -1 }; S.page = 1; refresh(); $('#tableCard').scrollIntoView({ behavior: 'smooth' }); });
    $('#btnCompose').addEventListener('click', openCompose);
  }

  // ---------- 2 panel tổng quan ----------
  var duoCharts = [];
  function ringChart(canvas, labels, data, colors, onClick) {
    return new Chart(canvas, {
      type: 'doughnut',
      data: { labels: labels, datasets: [{ data: data, backgroundColor: colors, borderColor: 'rgba(6,42,99,1)', borderWidth: 2, hoverOffset: 4 }] },
      options: { cutout: '70%', responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: function (c) { var t = c.dataset.data.reduce(function (a, b) { return a + b; }, 0); return ' ' + c.label + ': ' + fmtN(c.raw) + ' (' + pct(c.raw, t) + ')'; } } } },
        onClick: function (e, els) { if (els.length && onClick) onClick(els[0].index); } }
    });
  }
  function renderDuo() {
    duoCharts.forEach(function (c) { c.destroy(); }); duoCharts = [];
    var v = S.view, n = v.length;
    var stC = STATUSES.map(function (s) { return v.filter(function (d) { return d.status === s; }).length; });
    var stV = STATUSES.map(function (s) { return sum(v.filter(function (d) { return d.status === s; }), function (d) { return d.value; }); });
    var stCols = [css('--s1'), css('--good'), css('--critical')];
    var act = v.filter(function (d) { return d.status === 'Active'; });
    var stg = STAGES.map(function (s) { return [s, act.filter(function (d) { return d.stage === s; }).length]; });
    var appl = v.filter(function (d) { return d.status === 'Active' || !S.onlyActiveRemind; });
    var lvC = [0, 1, 2, 3, 4].map(function (l) { return appl.filter(function (d) { return d.level === l; }).length; });
    var remN = lvC[1] + lvC[2] + lvC[3] + lvC[4];
    var tl = TIMELINES.slice(0, 5).map(function (t) { return [t, act.filter(function (d) { return d.timeline === t; }).length]; });
    var tlCol = [css('--critical'), css('--warn'), css('--s1'), css('--s3'), css('--axis')];
    var rampC = [css('--ramp-1'), css('--ramp-2'), css('--ramp-3'), css('--ramp-4'), css('--ramp-5')];
    $('#duo').innerHTML =
      '<div class="panel"><div class="ring"><canvas id="ring1"></canvas><div class="c"><div><b>' + fmtN(n) + '</b><span>thương vụ</span></div></div></div><div><h4>Kết quả thương vụ</h4><div class="lg">' +
      STATUSES.map(function (s, i) { return '<div class="r" data-k="status" data-v="' + s + '"><i style="background:' + stCols[i] + '"></i><span>' + s + ' <span style="opacity:.7">· ' + fmtShort(stV[i]) + '</span></span><b class="num">' + fmtN(stC[i]) + '</b><span class="p">' + pct(stC[i], n) + '</span></div>'; }).join('') +
      '<hr><h4 style="margin:2px 0 4px">Thương vụ Active theo giai đoạn</h4>' +
      stg.map(function (x, i) { return '<div class="r" data-k="stage" data-v="' + esc(x[0]) + '"><i style="background:' + rampC[i] + '"></i><span>' + esc(x[0]) + '</span><b class="num">' + fmtN(x[1]) + '</b><span class="p">' + pct(x[1], act.length) + '</span></div>'; }).join('') +
      '</div></div></div>' +
      '<div class="panel alt"><div class="ring"><canvas id="ring2"></canvas><div class="c"><div><b>' + fmtN(remN) + '</b><span>cần nhắc</span></div></div></div><div><h4>Tình trạng cập nhật (cờ nhắc)</h4><div class="lg">' +
      [0, 1, 2, 3, 4].map(function (l) { return '<div class="r" data-k="flag" data-v="' + esc(FLAGS[l]) + '"><i style="background:' + flagColor(l) + '"></i><span>' + esc(FLAGS[l]) + '</span><b class="num">' + fmtN(lvC[l]) + '</b><span class="p">' + pct(lvC[l], appl.length) + '</span></div>'; }).join('') +
      '<hr><h4 style="margin:2px 0 4px">Timeline dự án (thương vụ Active)</h4>' +
      tl.map(function (x, i) { return '<div class="r" data-k="timeline" data-v="' + esc(x[0]) + '"><i style="background:' + tlCol[i] + '"></i><span>' + esc(x[0]) + '</span><b class="num">' + fmtN(x[1]) + '</b><span class="p">' + pct(x[1], act.length) + '</span></div>'; }).join('') +
      '</div></div></div>';
    duoCharts.push(ringChart($('#ring1'), STATUSES, stC, stCols, function (i) { toggleValue('status', STATUSES[i]); }));
    duoCharts.push(ringChart($('#ring2'), FLAGS.slice(0, 5), lvC, [0, 1, 2, 3, 4].map(flagColor), function (i) { toggleValue('flag', FLAGS[i]); }));
    $$('#duo .r').forEach(function (r) { r.addEventListener('click', function () { toggleValue(r.dataset.k, r.dataset.v); }); });
  }

  // ---------- KPI ----------
  function renderKpis() {
    var v = S.view, act = v.filter(function (d) { return d.status === 'Active'; }), won = v.filter(function (d) { return d.status === 'Won'; }), lost = v.filter(function (d) { return d.status === 'Lost'; });
    var tot = sum(v, function (d) { return d.value; }), withVal = v.filter(function (d) { return d.value > 0; });
    var rem = v.filter(function (d) { return d.remind; }), od = act.filter(function (d) { return d.overdue !== null; });
    var sales = new Set(v.map(function (d) { return d.sale; })), vend = new Set(), cust = new Set(v.map(function (d) { return d.customer; }));
    v.forEach(function (d) { d.vendorNames.forEach(function (x) { if (x !== NO_VENDOR) vend.add(x); }); });
    var wr = won.length + lost.length ? won.length / (won.length + lost.length) : null;
    var K = [
      { l: 'Tổng thương vụ', v: fmtN(v.length), s: fmtN(sales.size) + ' sale · ' + fmtN(vend.size) + ' hãng · ' + fmtN(cust.size) + ' KH', c: '--brand-600' },
      { l: 'Tổng giá trị', v: fmtShort(tot), s: 'TB ' + fmtShort(withVal.length ? tot / withVal.length : 0) + ' / thương vụ có giá trị', c: '--teal' },
      { l: 'Pipeline Active', v: fmtShort(sum(act, function (d) { return d.value; })), s: fmtN(act.length) + ' thương vụ đang xử lý', c: '--s1', f: ['status', ['Active']] },
      { l: 'Won', v: fmtShort(sum(won, function (d) { return d.value; })), s: fmtN(won.length) + ' thương vụ thắng', c: '--good', f: ['status', ['Won']] },
      { l: 'Lost', v: fmtN(lost.length), s: fmtShort(sum(lost, function (d) { return d.value; })) + ' giá trị mất', c: '--critical', f: ['status', ['Lost']] },
      { l: 'Tỷ lệ thắng', v: wr === null ? '–' : nf1.format(wr * 100) + '<small>%</small>', s: 'Won / (Won + Lost) theo số lượng', c: '--s3' },
      { l: 'Cần nhắc update', v: fmtN(rem.length), s: pct(rem.length, v.length) + ' · > 1 tháng không cập nhật', c: '--accent', f: ['flag', FLAGS.slice(1, 5)] },
      { l: 'Đã qua timeline', v: fmtN(od.length), s: fmtShort(sum(od, function (d) { return d.value; })) + ' · Active quá Close date', c: '--severe', f: ['timeline', [TIMELINES[0]]] }
    ];
    $('#kpis').innerHTML = K.map(function (k, i) {
      return '<div class="kpi' + (k.f ? ' click' : '') + '" data-i="' + i + '" style="--k:' + css(k.c) + '"' + (k.f ? ' title="Bấm để lọc"' : '') + '><div class="lb">' + esc(k.l) + '</div><div class="v num">' + k.v + '</div><div class="s">' + esc(k.s) + '</div></div>';
    }).join('');
    $$('#kpis .kpi.click').forEach(function (el) { el.addEventListener('click', function () { var f = K[+el.dataset.i].f; setFilter(f[0], f[1]); S.page = 1; refresh(); }); });
  }

  // ---------- điểm nhấn ----------
  function topBy(list, key, val) {
    var m = new Map();
    list.forEach(function (d) { var k = key(d); if (k === null || k === undefined) return; var e = m.get(k) || [0, 0]; e[0]++; e[1] += val(d) || 0; m.set(k, e); });
    return Array.from(m.entries()).sort(function (a, b) { return b[1][1] - a[1][1] || b[1][0] - a[1][0]; });
  }
  function issueCat(s) { return s.replace(/\(.*?\)/g, '').replace(/"[^"]*"/g, '…').replace(/\d[\d.,/]*/g, 'N').replace(/\s+/g, ' ').trim(); }
  function renderInsights() {
    var v = S.view, L = [], link = function (d) { return '<a data-open="' + d.id + '">#' + d.id + ' ' + esc(d.name.length > 60 ? d.name.slice(0, 58) + '…' : d.name) + '</a>'; };
    if (!v.length) { $('#insights').innerHTML = '<li>Không có thương vụ nào khớp bộ lọc.</li>'; return; }
    var act = v.filter(function (d) { return d.status === 'Active'; }), actV = sum(act, function (d) { return d.value; });
    var rem = v.filter(function (d) { return d.remind; }).sort(function (a, b) { return b.staleDays - a.staleDays; });
    if (rem.length) L.push('<span class="hot">' + fmtN(rem.length) + '</span> thương vụ hơn 1 tháng chưa cập nhật; lâu nhất ' + link(rem[0]) + ' — <b>' + fmtN(rem[0].staleDays) + ' ngày</b> (' + esc(rem[0].sale) + ').');
    var ts = topBy(act, function (d) { return d.sale; }, function (d) { return d.value; });
    if (ts.length && actV) L.push('Sale dẫn đầu pipeline Active: <b>' + esc(ts[0][0]) + '</b> với ' + fmtShort(ts[0][1][1]) + ' (' + fmtN(ts[0][1][0]) + ' thương vụ, ' + pct(ts[0][1][1], actV) + ' pipeline).');
    var tv = topBy(v, function (d) { return d.mainVendor === NO_VENDOR ? null : d.mainVendor; }, function (d) { return d.value; });
    if (tv.length) {
      var byCnt = tv.slice().sort(function (a, b) { return b[1][0] - a[1][0]; })[0];
      L.push('Hãng chính có giá trị lớn nhất: <b>' + esc(tv[0][0]) + '</b> (' + fmtShort(tv[0][1][1]) + '); nhiều thương vụ nhất: <b>' + esc(byCnt[0]) + '</b> (' + fmtN(byCnt[1][0]) + ').');
    }
    var od = act.filter(function (d) { return d.overdue !== null; }).sort(function (a, b) { return b.overdue - a.overdue; });
    if (od.length) L.push('<span class="hot">' + fmtN(od.length) + '</span> thương vụ Active đã qua timeline (tổng ' + fmtShort(sum(od, function (d) { return d.value; })) + '); quá hạn lâu nhất ' + link(od[0]) + ' — ' + fmtN(od[0].overdue) + ' ngày.');
    var soon = act.filter(function (d) { return d.timeline === TIMELINES[1]; });
    if (soon.length) L.push('<b>' + fmtN(soon.length) + '</b> thương vụ có Close date trong 30 ngày tới, giá trị ' + fmtShort(sum(soon, function (d) { return d.value; })) + ' — cần ưu tiên chốt.');
    var big = act.filter(function (d) { return d.value > 0; }).sort(function (a, b) { return b.value - a.value; })[0];
    if (big) L.push('Thương vụ Active lớn nhất: ' + link(big) + ' — <b>' + fmtShort(big.value) + '</b> (' + esc(big.sale) + ', ' + esc(big.stage) + ').');
    var tc = topBy(v, function (d) { return d.customer; }, function (d) { return d.value; }), totV = sum(v, function (d) { return d.value; });
    if (tc.length >= 5 && totV) L.push('Top 5 khách hàng chiếm <b>' + pct(tc.slice(0, 5).reduce(function (s, x) { return s + x[1][1]; }, 0), totV) + '</b> tổng giá trị; lớn nhất <b>' + esc(tc[0][0]) + '</b> (' + fmtShort(tc[0][1][1]) + ').');
    var seg = topBy(v, function (d) { return d.segment; }, function (d) { return d.value; });
    if (seg.length && totV) L.push('Nhóm KH đóng góp giá trị lớn nhất: <b>' + esc(seg[0][0]) + '</b> (' + pct(seg[0][1][1], totV) + ', ' + fmtN(seg[0][1][0]) + ' thương vụ).');
    var rs = topBy(v, function (d) { return d.reseller || null; }, function () { return 1; });
    var noRs = v.filter(function (d) { return !d.reseller; }).length;
    if (rs.length) L.push('Reseller nhiều thương vụ nhất: <b>' + esc(rs[0][0]) + '</b> (' + fmtN(rs[0][1][0]) + '); ' + fmtN(rs.length) + ' reseller, ' + fmtN(noRs) + ' thương vụ chưa ghi reseller.');
    var lost = v.filter(function (d) { return d.status === 'Lost' && d.failed; });
    if (lost.length) { var lr = topBy(lost, function (d) { return d.failed.split(' – ')[0]; }, function () { return 1; }); L.push('Lý do thất bại phổ biến: <b>' + esc(lr[0][0]) + '</b> (' + fmtN(lr[0][1][0]) + '/' + fmtN(lost.length) + ' thương vụ Lost có lý do).'); }
    var iss = v.filter(function (d) { return d.issues.length && d.type !== 'Lead chiến dịch'; });
    if (iss.length) {
      var cats = new Map(); iss.forEach(function (d) { d.issues.forEach(function (x) { var c = issueCat(x); cats.set(c, (cats.get(c) || 0) + 1); }); });
      var top = Array.from(cats.entries()).sort(function (a, b) { return b[1] - a[1]; })[0];
      L.push('<b>' + fmtN(iss.length) + '</b> thương vụ có cờ dữ liệu cần kiểm tra; phổ biến nhất: “' + esc(top[0]) + '” (' + fmtN(top[1]) + ').');
    }
    var cts = v.filter(function (d) { return d.contract; }), renew = cts.filter(function (d) { return d.contract === 'Renew'; }).length;
    if (cts.length) L.push('Phân loại HĐ: <b>' + fmtN(cts.length - renew) + '</b> New/khác · <b>' + fmtN(renew) + '</b> Renew; ' + fmtN(v.length - cts.length) + ' thương vụ chưa phân loại.');
    $('#insights').innerHTML = L.map(function (x) { return '<li>' + x + '</li>'; }).join('');
  }

  // ============================== BIỂU ĐỒ ==============================
  var CHARTS = [
    { id: 'sale', title: 'Theo Sale (chủ sở hữu thương vụ)', sub: 'chia theo trạng thái · bấm cột để lọc', kind: 'stack', dim: 'sale', w: 'w8', tall: true },
    { id: 'stage', title: 'Pipeline theo giai đoạn', sub: 'bấm thanh để lọc', kind: 'hbar', dim: 'stage', order: STAGES, color: 'ramp', w: 'w4', tall: true },
    { id: 'vendor', title: 'Theo Hãng', sub: 'top 15 · chế độ giá trị dùng Giá trị BOM theo hãng', kind: 'hbar', dim: 'vendor', top: 15, tall: true },
    { id: 'reseller', title: 'Theo Reseller', sub: 'top 15 · bấm thanh để lọc', kind: 'hbar', dim: 'reseller', top: 15, tall: true },
    { id: 'segment', title: 'Theo phân nhóm khách hàng', sub: 'bấm thanh để lọc', kind: 'hbar', dim: 'segment', tall: true },
    { id: 'bucket', title: 'Theo mức giá trị thương vụ', sub: 'bấm cột để lọc', kind: 'vbar', dim: 'bucket', order: BUCKETS, color: 'ramp' },
    { id: 'month', title: 'Close date (timeline dự án) theo tháng', sub: 'chia theo trạng thái · bấm cột để lọc theo tháng', kind: 'month', w: 'w8' },
    { id: 'contract', title: 'Phân loại hợp đồng', sub: 'New / Renew', kind: 'donut', dim: 'contract', w: 'w4' },
    { id: 'customer', title: 'Top 10 khách hàng', sub: 'theo chỉ số đang chọn', kind: 'hbar', dim: 'customer', top: 10 },
    { id: 'timeline', title: 'Tình trạng timeline – thương vụ Active', sub: 'so với ngày chốt', kind: 'hbar', dim: 'timeline', order: TIMELINES.slice(0, 5), activeOnly: true }
  ];
  var valueLabels = {
    id: 'valueLabels',
    afterDatasetsDraw: function (chart, args, opts) {
      if (!opts || !opts.on) return;
      var ctx = chart.ctx, horiz = chart.options.indexAxis === 'y', metas = chart.data.datasets.map(function (d, i) { return chart.getDatasetMeta(i); }).filter(function (m) { return !m.hidden; });
      if (!metas.length) return;
      ctx.save(); ctx.font = '600 11px ' + css('--font'); ctx.fillStyle = css('--ink-2');
      var n = chart.data.labels.length;
      for (var j = 0; j < n; j++) {
        var tot = 0, end = null;
        metas.forEach(function (m) { var val = chart.data.datasets[m.index].data[j] || 0; tot += val; var el = m.data[j]; if (el && val) end = horiz ? Math.max(end === null ? -1e9 : end, el.x) : Math.min(end === null ? 1e9 : end, el.y); });
        if (!tot || end === null) continue;
        var txt = opts.fmt(tot), el0 = metas[0].data[j];
        if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, end + 6, el0.y); }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, el0.x, end - 4); }
      }
      ctx.restore();
    }
  };
  function renderCharts(first) {
    if (first) {
      if (!valueLabels._reg) { Chart.register(valueLabels); valueLabels._reg = true; }
      $('#charts').innerHTML = CHARTS.map(function (c) {
        return '<div class="card ' + (c.w || '') + '"><div class="card-h"><h3>' + esc(c.title) + '</h3><span class="sub">' + esc(c.sub) + '</span></div>' +
          (c.kind === 'stack' || c.kind === 'month' ? '<div class="legend">' + STATUSES.map(function (s, i) { return '<span><i style="background:' + css(['--s1', '--good', '--critical'][i]) + '"></i>' + s + '</span>'; }).join('') + '</div>' : '') +
          '<div class="card-b"><div class="cbox' + (c.tall ? ' tall' : '') + '"><canvas id="ch-' + c.id + '"></canvas><div class="empty-chart hidden">Không có dữ liệu</div></div></div></div>';
      }).join('');
    }
    Object.keys(S.charts).forEach(function (k) { S.charts[k].destroy(); }); S.charts = {};
    Chart.defaults.font.family = css('--font'); Chart.defaults.color = css('--ink-2'); Chart.defaults.font.size = 12;
    CHARTS.forEach(function (c) { S.charts[c.id] = makeChart(c); });
    updateCharts();
  }
  function metricVal(d, key, dim) {
    if (S.metric === 'count') return 1;
    if (dim === 'vendor') {
      var m = d.vendors.filter(function (x) { return (x.v || NO_VENDOR) === key; });
      var b = m.reduce(function (s, x) { return s + (x.bom || 0); }, 0);
      if (b) return b;
      return key === d.mainVendor ? (d.value || 0) : 0;
    }
    return d.value || 0;
  }
  function aggregate(c, list) {
    var f = FBY[c.dim], m = new Map();
    list.forEach(function (d) {
      valuesOf(f, d).forEach(function (k) { m.set(k, (m.get(k) || 0) + metricVal(d, k, c.dim)); });
    });
    var ent = Array.from(m.entries());
    if (c.order) { ent = c.order.map(function (k) { return [k, m.get(k) || 0]; }); }
    else ent.sort(function (a, b) { return b[1] - a[1]; });
    if (c.top) ent = ent.filter(function (e) { return e[1] > 0; }).slice(0, c.top);
    return ent;
  }
  function axisFmt(v) { return S.metric === 'count' ? fmtN(v) : fmtShort(v); }
  function baseOpts(horiz) {
    var grid = css('--grid'), axis = css('--axis');
    var valAxis = { beginAtZero: true, grid: { color: grid, drawTicks: false }, border: { display: false }, ticks: { padding: 6, callback: function (v) { return axisFmt(v); }, maxTicksLimit: 6 } };
    var catAxis = { grid: { display: false }, border: { color: axis }, ticks: { autoSkip: false, padding: 4, callback: function (v) { var l = this.getLabelForValue(v); return l.length > 24 ? l.slice(0, 22) + '…' : l; } } };
    return {
      responsive: true, maintainAspectRatio: false, animation: { duration: 350 }, indexAxis: horiz ? 'y' : 'x',
      layout: { padding: horiz ? { right: 56 } : { top: 18 } },
      scales: horiz ? { x: valAxis, y: catAxis } : { x: catAxis, y: valAxis },
      plugins: {
        legend: { display: false },
        valueLabels: { on: true, fmt: axisFmt },
        tooltip: { backgroundColor: 'rgba(4,29,71,.95)', padding: 10, cornerRadius: 8, titleFont: { weight: '700' },
          callbacks: { label: function (ctx) { return ' ' + (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + (S.metric === 'count' ? fmtN(ctx.raw) + ' thương vụ' : fmtVND(ctx.raw) + ' VND (' + fmtShort(ctx.raw) + ')'); } } }
      },
      onHover: function (e, els) { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
    };
  }
  function makeChart(c) {
    var cv = $('#ch-' + c.id), horiz = c.kind === 'hbar';
    if (c.kind === 'donut') {
      return new Chart(cv, { type: 'doughnut', data: { labels: [], datasets: [{ data: [], borderColor: css('--surface'), borderWidth: 2 }] },
        options: { responsive: true, maintainAspectRatio: false, cutout: '62%', animation: { duration: 350 },
          plugins: { legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'rectRounded', padding: 14, generateLabels: function (ch) {
            var ds = ch.data.datasets[0], t = ds.data.reduce(function (a, b) { return a + b; }, 0);
            return ch.data.labels.map(function (l, i) { return { text: l + '  ' + (S.metric === 'count' ? fmtN(ds.data[i]) : fmtShort(ds.data[i])) + ' (' + pct(ds.data[i], t) + ')', fillStyle: ds.backgroundColor[i], strokeStyle: ds.backgroundColor[i], fontColor: css('--ink-2'), index: i }; });
          } }, onClick: function (e, item) { toggleValue(c.dim, c._labels[item.index]); } },
          tooltip: { callbacks: { label: function (ctx) { return ' ' + ctx.label + ': ' + (S.metric === 'count' ? fmtN(ctx.raw) + ' thương vụ' : fmtShort(ctx.raw)); } } } },
          onClick: function (e, els) { if (els.length) toggleValue(c.dim, c._labels[els[0].index]); } } });
    }
    var o = baseOpts(horiz);
    if (c.kind === 'stack' || c.kind === 'month') { o.scales.x.stacked = true; o.scales.y.stacked = true; o.plugins.tooltip.mode = 'index'; o.plugins.tooltip.callbacks.footer = function (items) { var t = items.reduce(function (s, i) { return s + i.raw; }, 0); return 'Tổng: ' + (S.metric === 'count' ? fmtN(t) + ' thương vụ' : fmtShort(t)); }; }
    o.onClick = function (e, els) {
      if (!els.length) return;
      var lab = c._labels[els[0].index];
      if (c.kind === 'month') {
        if (!c._ranges || !c._ranges[els[0].index]) return;
        var r = c._ranges[els[0].index]; S.date = { field: 'close', from: r[0], to: r[1] }; S.page = 1; refresh(); return;
      }
      toggleValue(c.dim, lab);
    };
    return new Chart(cv, { type: 'bar', data: { labels: [], datasets: [] }, options: o });
  }
  function colorsFor(c, labels) {
    var sel = c.dim && S.filters[c.dim], ramp = ['--ramp-1', '--ramp-2', '--ramp-3', '--ramp-4', '--ramp-5'].map(css);
    return labels.map(function (l, i) {
      var col;
      if (c.color === 'ramp') col = c.dim === 'bucket' ? (l === 'Chưa có giá trị' ? css('--axis') : ramp[Math.min(i, 4)]) : ramp[Math.min(i, 4)];
      else if (c.dim === 'timeline') col = [css('--critical'), css('--warn'), css('--s1'), css('--s3'), css('--axis')][i];
      else col = css('--s1');
      return sel && sel.size && !sel.has(l) ? alpha(col, 0.28) : col;
    });
  }
  function updateCharts() {
    if (!Object.keys(S.charts).length) return;
    var v = S.view, stCols = [css('--s1'), css('--good'), css('--critical')];
    CHARTS.forEach(function (c) {
      var ch = S.charts[c.id], list = c.activeOnly ? v.filter(function (d) { return d.status === 'Active'; }) : v, labels, empty;
      if (c.kind === 'stack') {
        var ent = aggregate(c, list).filter(function (e) { return e[1] > 0; }).slice(0, 25);
        labels = ent.map(function (e) { return e[0]; });
        var sel = S.filters.sale;
        ch.data.labels = labels;
        ch.data.datasets = STATUSES.map(function (s, si) {
          return { label: s, data: labels.map(function (l) { return sum(list.filter(function (d) { return d.sale === l && d.status === s; }), function (d) { return metricVal(d, l, 'sale'); }); }),
            backgroundColor: labels.map(function (l) { return sel && !sel.has(l) ? alpha(stCols[si], 0.28) : stCols[si]; }), borderColor: css('--surface'), borderWidth: { top: 2 }, borderRadius: 3, maxBarThickness: 34, borderSkipped: false };
        });
        empty = !labels.length;
      } else if (c.kind === 'month') {
        var ref = new Date(S.ref), y0 = ref.getUTCFullYear(), m0 = ref.getUTCMonth();
        var withClose = list.filter(function (d) { return d.close !== null && d.close !== undefined; });
        var keys = [], ranges = [], idx = {};
        var lo = Date.UTC(y0, m0 - 6, 1), hi = Date.UTC(y0, m0 + 13, 1);
        var hasBefore = withClose.some(function (d) { return d.close < lo; }), hasAfter = withClose.some(function (d) { return d.close >= hi; });
        if (hasBefore) { keys.push('Trước ' + pad2(new Date(lo).getUTCMonth() + 1) + '/' + new Date(lo).getUTCFullYear()); ranges.push([null, lo - DAY]); }
        for (var i = -6; i < 13; i++) { var t = Date.UTC(y0, m0 + i, 1), dt = new Date(t); idx[dt.getUTCFullYear() * 12 + dt.getUTCMonth()] = keys.length; keys.push('T' + (dt.getUTCMonth() + 1) + '/' + String(dt.getUTCFullYear()).slice(2)); ranges.push([t, Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)]); }
        if (hasAfter) { keys.push('Sau ' + pad2(new Date(hi - DAY).getUTCMonth() + 1) + '/' + new Date(hi - DAY).getUTCFullYear()); ranges.push([hi, null]); }
        var pos = function (d) { if (d.close < lo) return 0; if (d.close >= hi) return keys.length - 1; var dt = new Date(d.close); return idx[dt.getUTCFullYear() * 12 + dt.getUTCMonth()]; };
        labels = keys; c._ranges = ranges;
        ch.data.labels = keys;
        var refIdx = idx[y0 * 12 + m0];
        ch.data.datasets = STATUSES.map(function (s, si) {
          var data = keys.map(function () { return 0; });
          withClose.forEach(function (d) { if (d.status === s) data[pos(d)] += metricVal(d, null, 'month'); });
          return { label: s, data: data, backgroundColor: keys.map(function (k, i) { return i < refIdx ? alpha(stCols[si], 0.75) : stCols[si]; }), borderColor: css('--surface'), borderWidth: { top: 2 }, borderRadius: 3, maxBarThickness: 30, borderSkipped: false };
        });
        empty = !withClose.length;
      } else if (c.kind === 'donut') {
        var e2 = aggregate(c, list).filter(function (e) { return e[1] > 0; });
        labels = e2.map(function (e) { return e[0]; });
        var pal = ['--s1', '--s2', '--s3', '--s4', '--s7'].map(css), sel2 = S.filters[c.dim];
        ch.data.labels = labels;
        ch.data.datasets[0].data = e2.map(function (e) { return e[1]; });
        ch.data.datasets[0].backgroundColor = labels.map(function (l, i) { var col = l === NO_CONTRACT ? css('--axis') : pal[i % pal.length]; return sel2 && !sel2.has(l) ? alpha(col, 0.28) : col; });
        ch.data.datasets[0].borderColor = css('--surface');
        empty = !labels.length;
      } else {
        var e3 = aggregate(c, list);
        labels = e3.map(function (e) { return e[0]; });
        ch.data.labels = labels;
        ch.data.datasets = [{ label: '', data: e3.map(function (e) { return e[1]; }), backgroundColor: colorsFor(c, labels), borderRadius: 4, borderSkipped: 'start', maxBarThickness: c.kind === 'hbar' ? 22 : 46, categoryPercentage: 0.78 }];
        empty = !e3.some(function (e) { return e[1] > 0; });
      }
      c._labels = labels;
      ch.canvas.parentNode.querySelector('.empty-chart').classList.toggle('hidden', !empty);
      ch.update();
    });
  }

  // ============================== BẢNG ==============================
  function badgeStatus(s) { var c = s === 'Won' ? 'b-won' : s === 'Lost' ? 'b-lost' : 'b-active'; return '<span class="badge ' + c + '"><i style="background:currentColor"></i>' + esc(s) + '</span>'; }
  function badgeFlag(d) {
    var i = FLAGS.indexOf(d.flag);
    if (i >= 0 && i <= 4) return '<span class="badge b-l' + i + '">' + (i === 0 ? 'Bình thường' : 'Mức ' + i) + '</span>';
    return '<span class="badge b-grey">' + (i === 5 ? 'Đã đóng' : 'Không rõ') + '</span>';
  }
  var COLS = [
    { k: 'id', l: 'ID', v: function (d) { return d.id; }, cls: 'idc', num: true, def: true },
    { k: 'name', l: 'Tên thương vụ', v: function (d) { return d.name; }, cls: 'wrap2', w: 280, def: true },
    { k: 'customer', l: 'Khách hàng', v: function (d) { return d.customer; }, cls: 'wrap2', w: 220, def: true },
    { k: 'segment', l: 'Nhóm KH', v: function (d) { return d.segment; }, cls: 'clip', w: 180 },
    { k: 'sale', l: 'Sale', v: function (d) { return d.sale; }, cls: 'clip', def: true },
    { k: 'reseller', l: 'Reseller', v: function (d) { return d.reseller || ''; }, cls: 'clip', def: true },
    { k: 'stage', l: 'Giai đoạn', v: function (d) { return d.stage; }, cls: 'clip', def: true, sortV: function (d) { return STAGES.indexOf(d.stage); } },
    { k: 'status', l: 'Trạng thái', v: function (d) { return d.status; }, html: function (d) { return badgeStatus(d.status); }, def: true },
    { k: 'value', l: 'Giá trị (VND)', v: function (d) { return d.value; }, fmt: fmtVND, num: true, r: true, def: true },
    { k: 'bucket', l: 'Mức giá trị', v: function (d) { return d.bucket; }, sortV: function (d) { return BUCKETS.indexOf(d.bucket); } },
    { k: 'vendors', l: 'Hãng', v: function (d) { return d.vendorNames.join(', '); }, html: function (d) { return d.vendorNames.map(function (x, i) { return '<span class="vend' + (i === 0 ? ' main' : '') + '">' + esc(x) + '</span>'; }).join(''); }, def: true },
    { k: 'contract', l: 'Loại HĐ', v: function (d) { return d.contract || ''; } },
    { k: 'close', l: 'Close date', v: function (d) { return d.close; }, fmt: fmtD, num: true, def: true },
    { k: 'timeline', l: 'Tình trạng timeline', v: function (d) { return d.timeline; }, cls: 'clip' },
    { k: 'overdue', l: 'Quá hạn (ngày)', v: function (d) { return d.overdue; }, num: true, r: true },
    { k: 'updated', l: 'Cập nhật cuối', v: function (d) { return d.updated; }, fmt: fmtDT, num: true, def: true },
    { k: 'staleDays', l: 'Ngày chưa cập nhật', v: function (d) { return d.staleDays; }, num: true, r: true },
    { k: 'flag', l: 'Cờ nhắc', v: function (d) { return d.flag; }, html: badgeFlag, def: true, sortV: function (d) { var i = FLAGS.indexOf(d.flag); return i > 4 ? -1 : i; } },
    { k: 'created', l: 'Ngày tạo', v: function (d) { return d.created; }, fmt: fmtD, num: true },
    { k: 'ageDays', l: 'Tuổi thương vụ (ngày)', v: function (d) { return d.ageDays; }, num: true, r: true },
    { k: 'type', l: 'Loại deal', v: function (d) { return d.type; } },
    { k: 'dealSize', l: 'Deal size', v: function (d) { return d.dealSize || ''; } },
    { k: 'bomText', l: 'Thông tin BOM', v: function (d) { return d.bomText || ''; }, cls: 'wrap2', w: 300 },
    { k: 'failed', l: 'Lý do Failed', v: function (d) { return d.failed || ''; }, cls: 'wrap2', w: 240 },
    { k: 'quality', l: 'Cờ chất lượng', v: function (d) { return d.issues.join('; '); }, cls: 'wrap2', w: 360 }
  ];
  var CBY = {}; COLS.forEach(function (c) { CBY[c.k] = c; });
  function sortedView() {
    var c = CBY[S.sort.k], dir = S.sort.dir, get = c.sortV || c.v;
    return S.view.slice().sort(function (a, b) {
      var x = get(a), y = get(b), xn = x === null || x === undefined || x === '', yn = y === null || y === undefined || y === '';
      if (xn && yn) return 0; if (xn) return 1; if (yn) return -1;
      if (typeof x === 'number' && typeof y === 'number') return (x - y) * dir;
      return sortVi(x, y) * dir;
    });
  }
  function renderTable() {
    var cols = COLS.filter(function (c) { return S.cols.indexOf(c.k) >= 0; });
    var rows = sortedView(), ps = S.pageSize === 'all' ? rows.length || 1 : S.pageSize, pages = Math.max(1, Math.ceil(rows.length / ps));
    if (S.page > pages) S.page = pages;
    var page = rows.slice((S.page - 1) * ps, S.page * ps);
    $('#tblCount').textContent = fmtN(rows.length) + ' thương vụ · ' + cols.length + '/' + COLS.length + ' cột';
    $('#dt thead').innerHTML = '<tr>' + cols.map(function (c, i) {
      var on = S.sort.k === c.k;
      return '<th data-k="' + c.k + '" class="' + (c.r ? 'r ' : '') + (on ? 'sorted ' : '') + (i === 0 ? 'first' : '') + '"' + (c.w ? ' style="min-width:' + c.w + 'px"' : '') + '>' + esc(c.l) + '<span class="ar">' + (on ? (S.sort.dir > 0 ? '▲' : '▼') : '↕') + '</span></th>';
    }).join('') + '</tr>';
    var pad = S.dense ? 'padding:5px 12px' : '';
    $('#dt tbody').innerHTML = page.map(function (d) {
      return '<tr data-id="' + d.id + '">' + cols.map(function (c, i) {
        var val = c.v(d), h = c.html ? c.html(d) : esc(c.fmt ? c.fmt(val) : (val === null || val === undefined ? '' : (c.num && typeof val === 'number' && c.k !== 'id' ? fmtN(val) : val)));
        var w2 = c.cls === 'wrap2';
        return '<td class="' + (w2 ? '' : (c.cls || '')) + (c.r ? ' r num' : '') + (i === 0 ? ' first' : '') + '"' + (pad ? ' style="' + pad + '"' : '') + (w2 || c.cls === 'clip' ? ' title="' + esc(val) + '"' : '') + '>' + (w2 ? '<div class="w2">' + h + '</div>' : h) + '</td>';
      }).join('') + '</tr>';
    }).join('') || '<tr><td colspan="' + cols.length + '" style="text-align:center;padding:30px;color:var(--ink-3)">Không có thương vụ nào khớp bộ lọc</td></tr>';
    // pager
    var btns = [], add = function (p) { btns.push('<button data-p="' + p + '" class="' + (p === S.page ? 'on' : '') + '">' + p + '</button>'); };
    var lo = Math.max(1, S.page - 2), hi = Math.min(pages, S.page + 2);
    if (lo > 1) { add(1); if (lo > 2) btns.push('<span>…</span>'); }
    for (var p = lo; p <= hi; p++) add(p);
    if (hi < pages) { if (hi < pages - 1) btns.push('<span>…</span>'); add(pages); }
    $('#pager').innerHTML = '<span>Hiển thị ' + fmtN(rows.length ? (S.page - 1) * ps + 1 : 0) + '–' + fmtN(Math.min(S.page * ps, rows.length)) + ' / ' + fmtN(rows.length) + '</span>' +
      '<label>· Số dòng/trang <select id="psSel">' + [25, 50, 100, 200, 'all'].map(function (n) { return '<option value="' + n + '"' + (String(S.pageSize) === String(n) ? ' selected' : '') + '>' + (n === 'all' ? 'Tất cả' : n) + '</option>'; }).join('') + '</select></label>' +
      '<span class="sp"></span><button data-p="' + (S.page - 1) + '"' + (S.page <= 1 ? ' disabled' : '') + '>‹</button>' + btns.join('') + '<button data-p="' + (S.page + 1) + '"' + (S.page >= pages ? ' disabled' : '') + '>›</button>';
    $('#psSel').addEventListener('change', function (e) { S.pageSize = e.target.value === 'all' ? 'all' : +e.target.value; store('pageSize', S.pageSize); S.page = 1; renderTable(); });
    $$('#pager button[data-p]').forEach(function (b) { b.addEventListener('click', function () { S.page = +b.dataset.p; renderTable(); $('#tableCard').scrollIntoView({ block: 'start' }); }); });
  }
  $('#dt').addEventListener('click', function (e) {
    var th = e.target.closest('th');
    if (th) { var k = th.dataset.k; if (S.sort.k === k) S.sort.dir *= -1; else S.sort = { k: k, dir: CBY[k].num || k === 'value' ? -1 : 1 }; renderTable(); return; }
    var tr = e.target.closest('tr[data-id]'); if (tr) openDrawer(+tr.dataset.id);
  });
  $('#btnCols').addEventListener('click', function (e) {
    e.stopPropagation(); var fb = $('#colFb');
    if (openPop && openPop.parentNode === fb) { closePop(); return; }
    closePop();
    var pop = document.createElement('div'); pop.className = 'pop right colpop';
    var draw = function () {
      pop.innerHTML = '<div class="pop-h"><b>Ẩn / hiện cột</b><div class="row"><button data-act="all">Hiện tất cả</button><button data-act="def">Mặc định</button></div></div><div class="pop-b">' +
        COLS.map(function (c) { return '<label class="opt"><input type="checkbox" value="' + c.k + '"' + (S.cols.indexOf(c.k) >= 0 ? ' checked' : '') + (c.k === 'id' ? ' disabled' : '') + '><span class="l">' + esc(c.l) + '</span></label>'; }).join('') + '</div>';
    };
    draw(); fb.appendChild(pop); openPop = pop;
    pop.addEventListener('click', function (ev) {
      ev.stopPropagation(); var a = ev.target.closest('[data-act]'); if (!a) return;
      S.cols = a.dataset.act === 'all' ? COLS.map(function (c) { return c.k; }) : COLS.filter(function (c) { return c.def; }).map(function (c) { return c.k; });
      store('cols', S.cols); draw(); renderTable();
    });
    pop.addEventListener('change', function () { S.cols = COLS.filter(function (c) { var cb = pop.querySelector('input[value="' + c.k + '"]'); return cb && cb.checked; }).map(function (c) { return c.k; }); store('cols', S.cols); renderTable(); });
  });
  $('#btnDensity').addEventListener('click', function () { S.dense = !S.dense; store('dense', S.dense); renderTable(); });
  $('#btnCsv').addEventListener('click', function () {
    var cols = COLS.filter(function (c) { return S.cols.indexOf(c.k) >= 0; });
    var q = function (s) { s = s === null || s === undefined ? '' : String(s); return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    var lines = [cols.map(function (c) { return q(c.l); }).join(',')].concat(sortedView().map(function (d) {
      return cols.map(function (c) { var v = c.v(d); return q(c.fmt && c.k !== 'value' ? c.fmt(v) : v); }).join(',');
    }));
    download(new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'Danh sach thuong vu ' + fmtD(S.ref).replace(/\//g, '.') + '.csv');
  });

  // ============================== DRAWER CHI TIẾT ==============================
  function closeOverlay() { $$('.scrim,.drawer,.modal').forEach(function (e) { e.remove(); }); }
  function openDrawer(id) {
    var d = S.deals.find(function (x) { return x.id === id; }); if (!d) return;
    closeOverlay();
    var kv = [['Khách hàng', esc(d.customer)], ['Phân nhóm KH', esc(d.segment)], ['Sale', esc(d.sale)], ['Reseller', esc(d.reseller || '–')], ['Giai đoạn', esc(d.stage)], ['Trạng thái', badgeStatus(d.status)],
      ['Giá trị thương vụ', d.value ? '<b>' + fmtVND(d.value) + ' VND</b> <span class="muted">(' + fmtShort(d.value) + ' · ' + esc(d.bucket) + ')</span>' : '<span class="muted">Chưa có</span>'],
      ['Loại deal', esc(d.type)], ['Phân loại HĐ', esc(d.contract || '–')], ['Close date', d.close ? fmtD(d.close) + ' <span class="muted">· ' + esc(d.timeline) + (d.overdue !== null ? ' ' + fmtN(d.overdue) + ' ngày' : '') + '</span>' : '–'],
      ['Ngày tạo', fmtDT(d.created) + (d.ageDays !== null ? ' <span class="muted">· ' + fmtN(d.ageDays) + ' ngày trước</span>' : '')],
      ['Cập nhật lần cuối', fmtDT(d.updated) + (d.staleDays !== null ? ' <span class="muted">· ' + fmtN(d.staleDays) + ' ngày</span> ' : ' ') + badgeFlag(d)],
      ['Lý do Failed', esc(d.failed || '–')], ['Thông tin BOM', esc(d.bomText || '–')]];
    var sc = document.createElement('div'); sc.className = 'scrim'; sc.addEventListener('click', closeOverlay);
    var dr = document.createElement('aside'); dr.className = 'drawer';
    dr.innerHTML = '<div class="drawer-h"><button class="btn glass icon sm x" title="Đóng (Esc)"><svg class="i"><use href="#i-x"/></svg></button><div class="idl">THƯƠNG VỤ #' + d.id + '</div><h3>' + esc(d.name) + '</h3><div>' + badgeStatus(d.status) + ' ' + badgeFlag(d) + '</div></div>' +
      '<div class="drawer-b"><dl class="kv">' + kv.map(function (x) { return '<dt>' + x[0] + '</dt><dd>' + x[1] + '</dd>'; }).join('') + '</dl>' +
      '<div class="dsec"><h5>Hãng trong thương vụ</h5><table class="mini"><thead><tr><th>Hãng</th><th>Nguồn hãng</th><th style="text-align:right">Giá trị BOM</th></tr></thead><tbody>' +
      d.vendors.map(function (x, i) { return '<tr><td>' + (i === 0 ? '<b>' + esc(x.v || NO_VENDOR) + '</b> <span class="muted">(chính)</span>' : esc(x.v || NO_VENDOR)) + '</td><td>' + esc(x.src || '') + '</td><td style="text-align:right" class="num">' + (x.bom !== null && x.bom !== undefined ? fmtVND(x.bom) : '–') + '</td></tr>'; }).join('') + '</tbody></table></div>' +
      (d.issues.length ? '<div class="dsec"><h5>Vấn đề cần kiểm tra</h5><ul class="flags">' + d.issues.map(function (x) { return '<li class="issue">' + esc(x) + '</li>'; }).join('') + '</ul></div>' : '') +
      (d.done.length ? '<div class="dsec"><h5>Đã xử lý khi convert</h5><ul class="flags">' + d.done.map(function (x) { return '<li class="done">' + esc(x) + '</li>'; }).join('') + '</ul></div>' : '') + '</div>';
    document.body.appendChild(sc); document.body.appendChild(dr);
    dr.querySelector('.x').addEventListener('click', closeOverlay);
  }
  document.addEventListener('click', function (e) { var a = e.target.closest('[data-open]'); if (a) { e.preventDefault(); openDrawer(+a.dataset.open); } });

  // ============================== SOẠN NỘI DUNG NHẮC ==============================
  function openCompose() {
    closeOverlay();
    var sc = document.createElement('div'); sc.className = 'scrim'; sc.addEventListener('click', closeOverlay);
    var md = document.createElement('div'); md.className = 'modal';
    md.innerHTML = '<div class="modal-h"><h3>Soạn nội dung nhắc cập nhật CRM</h3><label class="muted" style="font-size:13px">Ngưỡng <select id="cmpLv" style="height:32px;border-radius:8px;border:1px solid var(--line-2);background:var(--surface)">' +
      [1, 2, 3, 4].map(function (l) { return '<option value="' + l + '">≥ ' + esc(FLAGS[l]) + '</option>'; }).join('') + '</select></label>' +
      '<label class="muted" style="font-size:13px">Hạn cập nhật <input type="date" id="cmpDue" value="' + isoD(S.ref + 7 * DAY) + '" style="height:32px;border-radius:8px;border:1px solid var(--line-2);background:var(--surface);padding:0 6px"></label>' +
      '<button class="btn ghost icon sm" data-x title="Đóng"><svg class="i"><use href="#i-x"/></svg></button></div>' +
      '<div class="modal-b"><p class="muted" style="margin:0 0 10px;font-size:13px">Nội dung được nhóm theo từng sale, dựa trên kết quả lọc hiện tại (' + fmtN(S.view.length) + ' thương vụ). Có thể chỉnh sửa trực tiếp trước khi sao chép.</p><textarea id="cmpTxt"></textarea></div>' +
      '<div class="modal-f"><button class="btn" id="cmpTxtDl"><svg class="i sm"><use href="#i-download"/></svg>Tải .txt</button><button class="btn primary" id="cmpCopy"><svg class="i sm"><use href="#i-copy"/></svg>Sao chép nội dung</button></div>';
    document.body.appendChild(sc); document.body.appendChild(md);
    var gen = function () {
      var lv = +$('#cmpLv').value, due = fromIso($('#cmpDue').value);
      var list = S.view.filter(function (d) { return d.remind && d.level >= lv; }).sort(function (a, b) { return b.staleDays - a.staleDays; });
      var bySale = new Map(); list.forEach(function (d) { if (!bySale.has(d.sale)) bySale.set(d.sale, []); bySale.get(d.sale).push(d); });
      var out = ['NHẮC CẬP NHẬT THƯƠNG VỤ TRÊN CRM', 'Số liệu tính đến ngày ' + fmtD(S.ref) + ' · Ngưỡng: ' + FLAGS[lv] + ' không cập nhật · Tổng: ' + list.length + ' thương vụ / ' + bySale.size + ' sale', ''];
      Array.from(bySale.entries()).sort(function (a, b) { return b[1].length - a[1].length; }).forEach(function (e) {
        out.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        out.push('Kính gửi anh/chị ' + e[0] + ',');
        out.push('Hiện có ' + e[1].length + ' thương vụ do anh/chị phụ trách chưa được cập nhật trên CRM (' + FLAGS[lv].replace(/^Mức \d · /, '') + '):');
        e[1].forEach(function (d, i) {
          out.push('  ' + (i + 1) + '. [Mức ' + d.level + ' · ' + d.staleDays + ' ngày] #' + d.id + ' – ' + d.name);
          out.push('     KH: ' + d.customer + ' · Giai đoạn: ' + d.stage + ' · Giá trị: ' + (d.value ? fmtShort(d.value) : 'chưa có') + ' · Close date: ' + (d.close ? fmtD(d.close) : 'chưa có') + ' · Cập nhật cuối: ' + (d.updated ? fmtD(d.updated) : '–'));
        });
        out.push('Đề nghị anh/chị cập nhật tình trạng, giai đoạn, timeline (Close date) và giá trị của các thương vụ trên' + (due ? ' trước ngày ' + fmtD(due) : '') + '. Thương vụ đã dừng vui lòng chuyển trạng thái Lost kèm lý do.');
        out.push('Trân trọng,');
        out.push('Phòng Quản lý sản phẩm');
        out.push('');
      });
      if (!list.length) out.push('Không có thương vụ nào vượt ngưỡng trong kết quả lọc hiện tại.');
      $('#cmpTxt').value = out.join('\n');
    };
    gen();
    $('#cmpLv').addEventListener('change', gen); $('#cmpDue').addEventListener('change', gen);
    md.querySelector('[data-x]').addEventListener('click', closeOverlay);
    $('#cmpCopy').addEventListener('click', function () {
      var t = $('#cmpTxt');
      var ok = function () { toast('Đã sao chép nội dung nhắc'); };
      if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t.value).then(ok, function () { t.select(); document.execCommand('copy'); ok(); });
      else { t.select(); document.execCommand('copy'); ok(); }
    });
    $('#cmpTxtDl').addEventListener('click', function () { download(new Blob([$('#cmpTxt').value], { type: 'text/plain;charset=utf-8' }), 'Nhac cap nhat CRM ' + fmtD(S.ref).replace(/\//g, '.') + '.txt'); });
  }

  // ============================== ĐỐI SOÁT ==============================
  function renderRecon() {
    var s = S.res.stats, w = S.res.warnings || [];
    var items = S.mode === 'raw' ? [
      [s.raw_rows, 'dòng trong sheet nguồn'], [s.deals_raw, 'deal raw'], [s.lines_raw, 'dòng BOM raw'], [s.tests, 'deal test đã loại'],
      [s.deals, 'thương vụ sau convert'], [s.co_hoi, 'Cơ hội'], [s.lead, 'Lead chiến dịch'], [s.rows, 'dòng template (tách hãng)'],
      [s.inferred, 'dòng BOM trống hãng (đã suy)'], [s.value_err, 'dòng BOM lỗi #VALUE!'], [fmtShort(s.value_total), 'tổng giá trị thương vụ'], [fmtShort(s.bom_total), 'tổng giá trị BOM theo hãng']
    ] : [[s.deals, 'thương vụ'], [s.co_hoi, 'Cơ hội'], [s.lead, 'Lead chiến dịch'], [s.rows, 'dòng template']];
    $('#recon').innerHTML = '<summary><svg class="i sm chev"><use href="#i-right"/></svg>Đối soát dữ liệu convert <span class="muted" style="font-weight:400">· ' + (S.mode === 'raw' ? 'sheet “' + esc(s.sheet) + '”' : 'file template') + ' · ngày chốt ' + fmtD(S.ref) + (w.length ? ' · <span style="color:var(--warn-ink)">' + w.length + ' cảnh báo</span>' : '') + '</span></summary>' +
      '<div class="card-b"><div class="recon">' + items.map(function (x) { return '<div><b class="num">' + (typeof x[0] === 'number' ? fmtN(x[0]) : x[0]) + '</b><span>' + x[1] + '</span></div>'; }).join('') + '</div>' +
      (S.mode === 'raw' && s.test_names.length ? '<p class="muted" style="font-size:12.5px;margin:10px 0 0">Deal test đã loại: ' + s.test_names.map(esc).join(' · ') + '</p>' : '') +
      (w.length ? '<ul class="warnlist">' + w.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>' : '') + '</div>';
  }

  // ============================== XUẤT FILE ==============================
  function download(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function writeWb(wb, name) {
    var out = X.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
  }
  function tplName() { var d = new Date(S.ref); return 'CRM_deals_template_' + pad2(d.getUTCDate()) + '_' + pad2(d.getUTCMonth() + 1) + '_' + String(d.getUTCFullYear()).slice(2) + '.xlsx'; }
  $('#btnTpl').addEventListener('click', function () {
    try {
      if (S.mode === 'raw') writeWb(C.buildTemplateWorkbook(X, S.res), tplName());
      else download(new Blob([S.fileBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), S.file);
      toast('Đã tải file template chuẩn');
    } catch (e) { console.error(e); toast('Lỗi tạo file: ' + e.message, true); }
  });

  // --- workbook tổng hợp ---
  var XF = 'Arial', thin = { style: 'thin', color: { rgb: 'D9DEE7' } }, BD = { left: thin, right: thin, top: thin, bottom: thin };
  var ST = {
    title: { font: { name: XF, sz: 16, bold: true, color: { rgb: '0B4EA2' } } },
    sub: { font: { name: XF, sz: 10, italic: true, color: { rgb: '667085' } } },
    h2: { font: { name: XF, sz: 12, bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '0B4EA2' } } },
    hdr: { font: { name: XF, sz: 10, bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '1F4E78' } }, border: BD, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } },
    hdrNew: { font: { name: XF, sz: 10, bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: 'F26522' } }, border: BD, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } },
    cell: { font: { name: XF, sz: 10 }, border: BD, alignment: { vertical: 'top' } },
    cellWrap: { font: { name: XF, sz: 10 }, border: BD, alignment: { vertical: 'top', wrapText: true } },
    grey: { font: { name: XF, sz: 10 }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: 'F2F2F2' } }, alignment: { vertical: 'top' } },
    total: { font: { name: XF, sz: 10, bold: true }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: 'E3EFFC' } } },
    kL: { font: { name: XF, sz: 10, bold: true, color: { rgb: '475467' } }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: 'F2F7FE' } } },
    kV: { font: { name: XF, sz: 11, bold: true, color: { rgb: '0F1B2D' } }, border: BD },
    label: { font: { name: XF, sz: 10, bold: true } },
    text: { font: { name: XF, sz: 10 }, alignment: { wrapText: true, vertical: 'top' } }
  };
  var lvlFill = ['E8F6E8', 'FFF4D6', 'FDE4D8', 'F9DADA', 'E8C4C4'];
  function put(ws, r, c, v, st, z) { ws[C.colLetter(c) + (r + 1)] = C.cellOf(v, z, st); }
  function finish(ws, nr, nc, widths, opts) {
    ws['!ref'] = 'A1:' + C.colLetter(Math.max(nc, 1) - 1) + Math.max(nr, 1);
    ws['!cols'] = widths.map(function (w) { return { wch: w }; });
    if (opts && opts.freeze) ws['!freeze'] = '<pane xSplit="' + (opts.freeze[0] || 0) + '" ySplit="' + opts.freeze[1] + '" topLeftCell="' + C.colLetter(opts.freeze[0] || 0) + (opts.freeze[1] + 1) + '" activePane="bottomRight" state="frozen"/>';
    if (opts && opts.filter) ws['!autofilter'] = { ref: opts.filter };
    return ws;
  }
  function tableSheet(header, rows, spec, title, subtitle) {
    // spec: [{w, z, wrap, isNew, date}]
    var ws = {}, r0 = 0;
    if (title) { put(ws, 0, 0, title, ST.title); put(ws, 1, 0, subtitle || '', ST.sub); r0 = 3; }
    header.forEach(function (h, j) { put(ws, r0, j, h, spec[j] && spec[j].isNew ? ST.hdrNew : ST.hdr); });
    rows.forEach(function (row, i) {
      var st0 = row._grey ? ST.grey : null;
      row.forEach(function (v, j) {
        var sp = spec[j] || {}, val = v;
        if (sp.date && v !== null && v !== undefined) val = C.toSerial(v);
        var st = row._style && row._style[j] ? row._style[j] : st0 || (sp.wrap ? ST.cellWrap : ST.cell);
        put(ws, r0 + 1 + i, j, val, st, sp.z);
      });
    });
    var last = C.colLetter(header.length - 1) + (r0 + rows.length + 1);
    ws['!rows'] = []; ws['!rows'][r0] = { hpt: 30 };
    return finish(ws, r0 + rows.length + 1, header.length, spec.map(function (s) { return s.w || 14; }), { freeze: [spec._freezeCols || 0, r0 + 1], filter: C.colLetter(0) + (r0 + 1) + ':' + last });
  }
  function pivot(list, keyFn, order) {
    var m = new Map();
    list.forEach(function (d) {
      keyFn(d).forEach(function (k) {
        var e = m.get(k); if (!e) { e = { n: 0, a: 0, w: 0, l: 0, v: 0, va: 0, vw: 0, rem: 0, od: 0, big: 0 }; m.set(k, e); }
        e.n++; e.v += d.value || 0;
        if (d.status === 'Active') { e.a++; e.va += d.value || 0; } else if (d.status === 'Won') { e.w++; e.vw += d.value || 0; } else if (d.status === 'Lost') e.l++;
        if (d.remind) e.rem++; if (d.overdue !== null) e.od++; if (d.value >= 5e9) e.big++;
      });
    });
    var ent = Array.from(m.entries());
    if (order) ent.sort(function (a, b) { var ia = order.indexOf(a[0]), ib = order.indexOf(b[0]); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib); });
    else ent.sort(function (a, b) { return b[1].v - a[1].v || b[1].n - a[1].n; });
    return ent;
  }
  var PIV_H = ['Số thương vụ', 'Active', 'Won', 'Lost', 'Tổng giá trị (VND)', 'Pipeline Active (VND)', 'Giá trị Won (VND)', 'Tỷ lệ thắng', 'Tỷ trọng giá trị', 'Cần nhắc (≥ Mức 1)', 'Đã qua timeline', 'Deal ≥ 5 tỷ'];
  function pivotRows(ent, total) {
    return ent.map(function (e) {
      var x = e[1];
      return [e[0], x.n, x.a, x.w, x.l, x.v, x.va, x.vw, x.w + x.l ? x.w / (x.w + x.l) : null, total ? x.v / total : null, x.rem, x.od, x.big];
    });
  }
  var PIV_SPEC = [{ w: 34 }, { w: 11, z: '#,##0' }, { w: 9, z: '#,##0' }, { w: 9, z: '#,##0' }, { w: 9, z: '#,##0' }, { w: 20, z: '#,##0' }, { w: 20, z: '#,##0' }, { w: 20, z: '#,##0' }, { w: 11, z: '0.0%' }, { w: 11, z: '0.0%' }, { w: 12, z: '#,##0' }, { w: 12, z: '#,##0' }, { w: 10, z: '#,##0' }];
  function pivotSheet(name, ent, list, sub) {
    var total = sum(list, function (d) { return d.value; });
    var rows = pivotRows(ent, total);
    var t = ['TỔNG CỘNG', list.length, list.filter(function (d) { return d.status === 'Active'; }).length, list.filter(function (d) { return d.status === 'Won'; }).length, list.filter(function (d) { return d.status === 'Lost'; }).length, total,
      sum(list.filter(function (d) { return d.status === 'Active'; }), function (d) { return d.value; }), sum(list.filter(function (d) { return d.status === 'Won'; }), function (d) { return d.value; }), null, total ? 1 : null,
      list.filter(function (d) { return d.remind; }).length, list.filter(function (d) { return d.overdue !== null; }).length, list.filter(function (d) { return d.value >= 5e9; }).length];
    t._style = t.map(function () { return ST.total; });
    rows.push(t);
    return tableSheet([name].concat(PIV_H), rows, PIV_SPEC, 'Tổng hợp theo ' + name.toLowerCase(), sub);
  }
  function filterDesc() {
    var parts = [];
    if (S.q) parts.push('Tìm: "' + S.q + '"');
    FILTERS.forEach(function (f) { if (S.filters[f.k]) parts.push(f.label + ': ' + Array.from(S.filters[f.k]).join(', ')); });
    if (S.date.from !== null || S.date.to !== null) parts.push(DATE_FIELDS[S.date.field] + ': ' + (S.date.from !== null ? fmtD(S.date.from) : '…') + ' → ' + (S.date.to !== null ? fmtD(S.date.to) : '…'));
    return parts.length ? parts.join(' | ') : 'Không áp dụng bộ lọc (toàn bộ dữ liệu)';
  }
  function exportSummary() {
    var v = sortedView(), wb = X.utils.book_new(), ref = new Date(S.ref);
    var sub = 'Ngày chốt số liệu: ' + fmtD(S.ref) + ' · Nguồn: ' + S.file + ' · Bộ lọc: ' + filterDesc();
    // ---- Tổng quan
    var ws = {}, r = 0;
    put(ws, r++, 0, 'TỔNG HỢP CRM – THƯƠNG VỤ & DỰ ÁN · THÁNG ' + (ref.getUTCMonth() + 1) + '.' + ref.getUTCFullYear(), ST.title);
    put(ws, r++, 0, 'Công ty Nam Trường Sơn Hà Nội (NTS Hanoi Corp.) – Phòng Quản lý sản phẩm', ST.sub);
    r++;
    [['Ngày chốt số liệu', fmtD(S.ref)], ['File nguồn', S.file + (S.mode === 'raw' ? ' (convert từ raw)' : ' (file template)')], ['Bộ lọc áp dụng', filterDesc()], ['Số thương vụ trong báo cáo', v.length + ' / ' + S.deals.length]].forEach(function (x) {
      put(ws, r, 0, x[0], ST.kL); put(ws, r, 1, x[1], ST.kV); r++;
    });
    r++;
    var act = v.filter(function (d) { return d.status === 'Active'; }), won = v.filter(function (d) { return d.status === 'Won'; }), lost = v.filter(function (d) { return d.status === 'Lost'; });
    put(ws, r++, 0, 'CHỈ SỐ CHÍNH', ST.h2); put(ws, r - 1, 1, '', ST.h2); put(ws, r - 1, 2, '', ST.h2);
    [['Tổng thương vụ', v.length, '#,##0'], ['Tổng giá trị (VND)', sum(v, function (d) { return d.value; }), '#,##0'], ['Pipeline Active (VND)', sum(act, function (d) { return d.value; }), '#,##0'], ['Số thương vụ Active', act.length, '#,##0'],
      ['Giá trị Won (VND)', sum(won, function (d) { return d.value; }), '#,##0'], ['Số thương vụ Won', won.length, '#,##0'], ['Số thương vụ Lost', lost.length, '#,##0'],
      ['Tỷ lệ thắng (Won/(Won+Lost))', won.length + lost.length ? won.length / (won.length + lost.length) : null, '0.0%'], ['Cần nhắc cập nhật (≥ Mức 1)', v.filter(function (d) { return d.remind; }).length, '#,##0'],
      ['Active đã qua timeline', act.filter(function (d) { return d.overdue !== null; }).length, '#,##0'], ['Close date trong 30 ngày tới', act.filter(function (d) { return d.timeline === TIMELINES[1]; }).length, '#,##0'],
      ['Thương vụ có cờ dữ liệu cần kiểm tra', v.filter(function (d) { return d.issues.length && d.type !== 'Lead chiến dịch'; }).length, '#,##0']].forEach(function (x) {
      put(ws, r, 0, x[0], ST.kL); put(ws, r, 1, x[1], ST.kV, x[2]); r++;
    });
    var block = function (title, ent, h) {
      r++; put(ws, r, 0, title, ST.h2); for (var j = 1; j < 5; j++) put(ws, r, j, '', ST.h2); r++;
      [h, 'Số thương vụ', 'Tỷ trọng SL', 'Giá trị (VND)', 'Tỷ trọng GT'].forEach(function (x, j) { put(ws, r, j, x, ST.hdr); }); r++;
      var tn = v.length, tv = sum(v, function (d) { return d.value; });
      ent.forEach(function (e, i) {
        var st = e.st || ST.cell;
        put(ws, r, 0, e[0], st); put(ws, r, 1, e[1], st, '#,##0'); put(ws, r, 2, tn ? e[1] / tn : null, st, '0.0%'); put(ws, r, 3, e[2], st, '#,##0'); put(ws, r, 4, tv ? e[2] / tv : null, st, '0.0%'); r++;
      });
    };
    var grp = function (order, fn) { return order.map(function (k) { var l = v.filter(function (d) { return fn(d) === k; }); return [k, l.length, sum(l, function (d) { return d.value; })]; }); };
    block('THEO TRẠNG THÁI', grp(STATUSES, function (d) { return d.status; }), 'Trạng thái');
    block('THEO GIAI ĐOẠN', grp(STAGES, function (d) { return d.stage; }), 'Giai đoạn');
    block('THEO MỨC GIÁ TRỊ', grp(BUCKETS, function (d) { return d.bucket; }), 'Mức giá trị');
    var fl = grp(FLAGS, function (d) { return d.flag; }).filter(function (e) { return e[1] || FLAGS.indexOf(e[0]) <= 4; });
    fl.forEach(function (e) { var i = FLAGS.indexOf(e[0]); if (i <= 4) e.st = { font: { name: XF, sz: 10 }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: lvlFill[i] } } }; });
    block('CỜ NHẮC CẬP NHẬT (theo Cập nhật lần cuối)', fl, 'Cờ nhắc');
    block('TÌNH TRẠNG TIMELINE', grp(TIMELINES, function (d) { return d.timeline; }), 'Timeline');
    var ctv = Array.from(new Set(v.map(function (d) { return d.contract || NO_CONTRACT; })));
    block('PHÂN LOẠI HỢP ĐỒNG', grp(ctv, function (d) { return d.contract || NO_CONTRACT; }), 'Loại HĐ');
    block('THEO PHÂN NHÓM KHÁCH HÀNG', grp(C.SEGMENTS, function (d) { return d.segment; }).filter(function (e) { return e[1]; }), 'Nhóm KH');
    finish(ws, r, 5, [40, 26, 14, 22, 14]);
    X.utils.book_append_sheet(wb, ws, 'Tổng quan');

    // ---- Danh sách thương vụ (1 dòng / thương vụ)
    var H = ['ID', 'Tên thương vụ', 'Loại deal', 'Tên khách hàng', 'Phân nhóm khách hàng', 'Sale', 'Reseller', 'Giai đoạn', 'Trạng thái', 'Giá trị thương vụ', 'Mức giá trị', 'Hãng chính', 'Các hãng', 'Số hãng',
      'Phân loại HĐ', 'Phân loại deal size', 'Close date', 'Tháng/Quý Close date', 'Tình trạng timeline', 'Số ngày quá hạn forecast', 'Ngày tạo', 'Tuổi thương vụ (ngày)', 'Cập nhật lần cuối', 'Số ngày chưa cập nhật', 'Cờ nhắc cập nhật',
      'Số vấn đề dữ liệu', 'Vấn đề cần kiểm tra', 'Lý do Failed', C.BOM_TXT];
    var NEW = ['Mức giá trị', 'Hãng chính', 'Các hãng', 'Số hãng', 'Tháng/Quý Close date', 'Tình trạng timeline', 'Tuổi thương vụ (ngày)', 'Số ngày chưa cập nhật', 'Cờ nhắc cập nhật', 'Số vấn đề dữ liệu', 'Vấn đề cần kiểm tra'];
    var W = { 'ID': 10, 'Tên thương vụ': 44, 'Tên khách hàng': 36, 'Phân nhóm khách hàng': 26, 'Sale': 20, 'Reseller': 16, 'Giai đoạn': 20, 'Trạng thái': 10, 'Giá trị thương vụ': 18, 'Các hãng': 24, 'Close date': 12, 'Ngày tạo': 16, 'Cập nhật lần cuối': 16, 'Cờ nhắc cập nhật': 22, 'Vấn đề cần kiểm tra': 70, 'Lý do Failed': 36, 'Tình trạng timeline': 20, 'Tháng/Quý Close date': 14 };
    W[C.BOM_TXT] = 50;
    var Z = { 'ID': '0', 'Giá trị thương vụ': '#,##0', 'Close date': 'dd/mm/yyyy', 'Ngày tạo': 'dd/mm/yyyy hh:mm', 'Cập nhật lần cuối': 'dd/mm/yyyy hh:mm' };
    var spec = H.map(function (h) { return { w: W[h] || 13, z: Z[h], isNew: NEW.indexOf(h) >= 0, date: ['Close date', 'Ngày tạo', 'Cập nhật lần cuối'].indexOf(h) >= 0, wrap: ['Tên thương vụ', 'Vấn đề cần kiểm tra', C.BOM_TXT, 'Lý do Failed'].indexOf(h) >= 0 }; });
    spec._freezeCols = 2;
    var rows = v.map(function (d) {
      var cq = d.close ? (function () { var x = new Date(d.close); return 'T' + (x.getUTCMonth() + 1) + '/' + x.getUTCFullYear() + ' · Q' + (Math.floor(x.getUTCMonth() / 3) + 1); })() : null;
      var row = [d.id, d.name, d.type, d.customer, d.segment, d.sale, d.reseller, d.stage, d.status, d.value, d.bucket, d.mainVendor, d.vendorNames.join(', '), d.vendorNames.length,
        d.contract, d.dealSize, d.close, cq, d.timeline, d.overdue, d.created, d.ageDays, d.updated, d.staleDays, d.flag, d.issues.length, d.issues.join('; ') || null, d.failed, d.bomText];
      var li = FLAGS.indexOf(d.flag);
      if (li >= 1 && li <= 4) { row._style = []; row._style[24] = { font: { name: XF, sz: 10, bold: true }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: lvlFill[li] } } }; }
      return row;
    });
    X.utils.book_append_sheet(wb, tableSheet(H, rows, spec, 'Danh sách thương vụ (1 dòng / thương vụ)', sub), 'Danh sách thương vụ');

    // ---- Pivot sheets
    X.utils.book_append_sheet(wb, pivotSheet('Sale', pivot(v, function (d) { return [d.sale]; }), v, sub), 'Theo Sale');
    X.utils.book_append_sheet(wb, pivotSheet('Hãng', pivot(v, function (d) { return d.vendorNames; }), v, sub + ' · Một thương vụ nhiều hãng được tính cho từng hãng; giá trị = giá trị thương vụ.'), 'Theo Hãng');
    X.utils.book_append_sheet(wb, pivotSheet('Reseller', pivot(v, function (d) { return [d.reseller || NO_RESELLER]; }), v, sub), 'Theo Reseller');
    X.utils.book_append_sheet(wb, pivotSheet('Nhóm khách hàng', pivot(v, function (d) { return [d.segment]; }, C.SEGMENTS), v, sub), 'Theo Nhóm KH');
    X.utils.book_append_sheet(wb, pivotSheet('Khách hàng', pivot(v, function (d) { return [d.customer]; }), v, sub), 'Theo Khách hàng');

    // ---- Nhắc cập nhật
    var rem = v.filter(function (d) { return d.remind; }).sort(function (a, b) { return a.sale.localeCompare(b.sale, 'vi') || b.staleDays - a.staleDays; });
    var RH = ['Sale', 'Cờ nhắc', 'Số ngày chưa cập nhật', 'Cập nhật lần cuối', 'ID', 'Tên thương vụ', 'Tên khách hàng', 'Giai đoạn', 'Trạng thái', 'Giá trị thương vụ', 'Close date', 'Tình trạng timeline'];
    var rspec = [{ w: 22 }, { w: 22 }, { w: 12, z: '#,##0' }, { w: 16, z: 'dd/mm/yyyy hh:mm', date: true }, { w: 10, z: '0' }, { w: 46, wrap: true }, { w: 34 }, { w: 20 }, { w: 10 }, { w: 18, z: '#,##0' }, { w: 12, z: 'dd/mm/yyyy', date: true }, { w: 20 }];
    var rrows = rem.map(function (d) {
      var row = [d.sale, d.flag, d.staleDays, d.updated, d.id, d.name, d.customer, d.stage, d.status, d.value, d.close, d.timeline];
      row._style = []; row._style[1] = { font: { name: XF, sz: 10, bold: true }, border: BD, fill: { patternType: 'solid', fgColor: { rgb: lvlFill[d.level] } } };
      return row;
    });
    X.utils.book_append_sheet(wb, tableSheet(RH, rrows, rspec, 'Danh sách thương vụ cần nhắc cập nhật (> 1 tháng)', sub + (S.onlyActiveRemind ? ' · Chỉ tính thương vụ Active' : '')), 'Nhắc cập nhật');

    // ---- Chi tiết theo hãng (dòng template, có cột bổ sung)
    var ids = new Set(v.map(function (d) { return d.id; })), byId = new Map(v.map(function (d) { return [d.id, d]; }));
    var TC = C.TEMPLATE_COLS.concat(['Mức giá trị', 'Số ngày chưa cập nhật', 'Cờ nhắc cập nhật', 'Tình trạng timeline']);
    var TZ = { 'ID': '0', 'Giá trị thương vụ': '#,##0', 'Giá trị BOM theo hãng': '#,##0', 'Close date': 'dd/mm/yyyy', 'Ngày tạo': 'dd/mm/yyyy hh:mm', 'Cập nhật lần cuối': 'dd/mm/yyyy hh:mm', 'Số ngày quá hạn forecast': '0' };
    var TW = { 'ID': 10, 'Tên thương vụ': 42, 'Tên khách hàng': 36, 'Phân nhóm khách hàng': 26, 'Sale': 20, 'Giai đoạn': 20, 'Giá trị thương vụ': 18, 'Nguồn hãng': 20, 'Giá trị BOM theo hãng': 18, 'Lý do Failed': 36, 'Ngày tạo': 16, 'Cập nhật lần cuối': 16, 'Cờ chất lượng': 80, 'Cờ nhắc cập nhật': 22, 'Tình trạng timeline': 20 };
    TW[C.BOM_TXT] = 50;
    var tspec = TC.map(function (h, j) { return { w: TW[h] || 14, z: TZ[h], isNew: j >= C.TEMPLATE_COLS.length, date: C.DATE_COLS.indexOf(h) >= 0 }; }); tspec._freezeCols = 2;
    var trows = [];
    S.res.rows.forEach(function (r0, i) {
      if (!ids.has(r0.ID)) return; var d = byId.get(r0.ID);
      var row = C.TEMPLATE_COLS.map(function (h) { return r0[h]; }).concat([d.bucket, d.staleDays, d.flag, d.timeline]);
      if (S.res.splitFlags[i]) row._grey = true;
      trows.push(row);
    });
    X.utils.book_append_sheet(wb, tableSheet(TC, trows, tspec, 'Chi tiết theo hãng (dữ liệu template chuẩn, 1 dòng / hãng)', sub + ' · Dòng nền xám = dòng tách hãng phụ, giá trị thương vụ chỉ ghi ở dòng hãng chính.'), 'Chi tiết theo hãng');

    // ---- Hướng dẫn
    var g = {}, gl = [
      ['HƯỚNG DẪN ĐỌC FILE TỔNG HỢP', null],
      ['Ngày chốt số liệu', fmtD(S.ref)],
      ['Phạm vi', 'Toàn bộ các sheet tính trên tập thương vụ đã lọc trong tool tại thời điểm xuất. Bộ lọc: ' + filterDesc()],
      ['Tổng quan', 'Chỉ số chính và bảng phân bố theo trạng thái, giai đoạn, mức giá trị, cờ nhắc, timeline, loại HĐ, nhóm KH.'],
      ['Danh sách thương vụ', 'Mỗi thương vụ 1 dòng (gộp các dòng tách hãng). Cột tiêu đề màu cam là cột bổ sung để phân tích.'],
      ['Theo Sale / Hãng / Reseller / Nhóm KH / Khách hàng', 'Bảng tổng hợp: số thương vụ, Active/Won/Lost, tổng giá trị, pipeline Active, giá trị Won, tỷ lệ thắng, tỷ trọng giá trị, số cần nhắc, số đã qua timeline, số deal ≥ 5 tỷ.'],
      ['Nhắc cập nhật', 'Danh sách thương vụ có Cập nhật lần cuối cách ngày chốt hơn 1 tháng, sắp theo sale và số ngày chưa cập nhật.'],
      ['Chi tiết theo hãng', 'Dữ liệu template chuẩn (giữ nguyên cột của file convert, 1 dòng / hãng) kèm 4 cột bổ sung.'],
      ['Mức giá trị', '< 1 tỷ · 1 – 2 tỷ · 2 – 5 tỷ · 5 – 10 tỷ · > 10 tỷ (cận dưới tính vào mức trên, VD đúng 2 tỷ thuộc 2 – 5 tỷ) · Chưa có giá trị.'],
      ['Cờ nhắc cập nhật', 'So Cập nhật lần cuối với ngày chốt: Mức 1 > 1 tháng · Mức 2 > 3 tháng · Mức 3 > 6 tháng · Mức 4 > 1 năm (theo tháng dương lịch). ' + (S.onlyActiveRemind ? 'Chỉ áp dụng cho thương vụ Active; thương vụ Won/Lost ghi "Không áp dụng".' : 'Áp dụng cho mọi trạng thái.')],
      ['Tình trạng timeline', 'Thương vụ Active so Close date với ngày chốt: Đã qua timeline · Trong 30 ngày tới · 31 – 90 ngày tới · Sau 90 ngày · Chưa có timeline. Won/Lost ghi "Đã đóng".'],
      ['Tuổi thương vụ', 'Số ngày từ Ngày tạo đến ngày chốt.'],
      ['Tỷ lệ thắng', 'Won / (Won + Lost) theo số lượng thương vụ.'],
      ['Vấn đề cần kiểm tra', 'Lấy từ phần [VẤN ĐỀ] của cột Cờ chất lượng trong file template (do bước convert sinh ra).']
    ];
    gl.forEach(function (x, i) { put(g, i, 0, x[0], i === 0 ? ST.title : ST.label); put(g, i, 1, x[1], ST.text); });
    finish(g, gl.length, 2, [34, 120]);
    X.utils.book_append_sheet(wb, g, 'Hướng dẫn');

    writeWb(wb, 'Tổng hợp CRM_ Tháng ' + (ref.getUTCMonth() + 1) + '.' + ref.getUTCFullYear() + '.xlsx');
  }
  $('#btnExport').addEventListener('click', function () {
    showLoading(true, 'Đang tạo file Excel tổng hợp…');
    setTimeout(function () {
      try { exportSummary(); toast('Đã xuất file Excel tổng hợp (' + fmtN(S.view.length) + ' thương vụ)'); }
      catch (e) { console.error(e); toast('Lỗi xuất file: ' + e.message, true); }
      finally { showLoading(false); }
    }, 30);
  });

  // ============================== TÌM KIẾM + GỢI Ý ==============================
  var sugItems = [], sugIdx = -1;
  function hideSuggest() { $('#suggest').classList.add('hidden'); sugIdx = -1; }
  function hl(text, toks) {
    var f = fold(text), marks = [];
    toks.forEach(function (t) { var i = f.indexOf(t); if (i >= 0) marks.push([i, i + t.length]); });
    if (!marks.length || f.length !== text.length) return esc(text);
    marks.sort(function (a, b) { return a[0] - b[0]; });
    var out = '', p = 0;
    marks.forEach(function (m) { if (m[0] < p) return; out += esc(text.slice(p, m[0])) + '<mark>' + esc(text.slice(m[0], m[1])) + '</mark>'; p = m[1]; });
    return out + esc(text.slice(p));
  }
  function score(text, toks) {
    var f = fold(text), s = 0;
    for (var i = 0; i < toks.length; i++) { var j = f.indexOf(toks[i]); if (j < 0) return -1; s += j === 0 ? 3 : /[^a-z0-9]/.test(f[j - 1]) ? 2 : 1; }
    return s;
  }
  function buildSuggest() {
    var raw = $('#q').value, toks = fold(raw).split(/\s+/).filter(Boolean), box = $('#suggest');
    if (!toks.length) { hideSuggest(); return; }
    var groups = [
      { t: 'Khách hàng', k: 'customer', vals: function (d) { return [d.customer]; } },
      { t: 'Sale', k: 'sale', vals: function (d) { return [d.sale]; } },
      { t: 'Hãng', k: 'vendor', vals: function (d) { return d.vendorNames; } },
      { t: 'Reseller', k: 'reseller', vals: function (d) { return d.reseller ? [d.reseller] : []; } },
      { t: 'Nhóm KH', k: 'segment', vals: function (d) { return [d.segment]; } }
    ];
    sugItems = []; var html = '';
    groups.forEach(function (g) {
      var m = new Map();
      S.deals.forEach(function (d) { g.vals(d).forEach(function (x) { if (!x) return; var e = m.get(x); if (e) e.n++; else m.set(x, { n: 1, s: score(x, toks) }); }); });
      var hits = Array.from(m.entries()).filter(function (e) { return e[1].s >= 0; }).sort(function (a, b) { return b[1].s - a[1].s || b[1].n - a[1].n; }).slice(0, g.k === 'customer' ? 6 : 4);
      if (!hits.length) return;
      html += '<div class="sg-h">' + g.t + '</div>';
      hits.forEach(function (h) { sugItems.push({ type: 'filter', k: g.k, v: h[0] }); html += '<div class="sg" data-i="' + (sugItems.length - 1) + '"><span class="tag">' + g.t + '</span><span class="t">' + hl(h[0], toks) + '</span><span class="c">' + fmtN(h[1].n) + ' thương vụ</span></div>'; });
    });
    var deals = S.deals.map(function (d) { return [d, Math.max(score(d.name, toks), score(String(d.id), toks))]; }).filter(function (x) { return x[1] >= 0; })
      .sort(function (a, b) { return b[1] - a[1] || (b[0].value || 0) - (a[0].value || 0); }).slice(0, 6);
    if (deals.length) {
      html += '<div class="sg-h">Thương vụ</div>';
      deals.forEach(function (x) { var d = x[0]; sugItems.push({ type: 'deal', id: d.id }); html += '<div class="sg" data-i="' + (sugItems.length - 1) + '"><span class="tag">#' + d.id + '</span><span class="t">' + hl(d.name, toks) + '<small>' + esc(d.customer) + '</small></span><span class="c">' + esc(d.status) + '</span></div>'; });
    }
    var all = S.deals.filter(function (d) { return passSearch(d, toks); }).length;
    sugItems.push({ type: 'text' });
    html = '<div class="sg" data-i="' + (sugItems.length - 1) + '"><span class="tag" style="background:var(--brand-600);color:#fff">Enter</span><span class="t">Lọc mọi thương vụ chứa “<b>' + esc(raw.trim()) + '</b>”</span><span class="c">' + fmtN(all) + ' kết quả</span></div>' + html;
    sugItems.unshift(sugItems.pop());
    // đánh lại chỉ số sau khi đưa dòng "Enter" lên đầu
    var idx = 0; html = html.replace(/data-i="\d+"/g, function () { return 'data-i="' + (idx++) + '"'; });
    if (!all) html += '<div class="sg-empty">Không tìm thấy kết quả phù hợp</div>';
    html += '<div class="sg-foot">↑ ↓ để chọn · Enter để áp dụng · Esc để đóng · tìm không phân biệt dấu</div>';
    box.innerHTML = html; box.classList.remove('hidden'); sugIdx = 0; markSug();
  }
  function markSug() { $$('#suggest .sg').forEach(function (el) { el.classList.toggle('act', +el.dataset.i === sugIdx); if (+el.dataset.i === sugIdx) el.scrollIntoView({ block: 'nearest' }); }); }
  function pickSug(i) {
    var it = sugItems[i]; if (!it) return;
    if (it.type === 'text') { S.q = $('#q').value.trim(); }
    else if (it.type === 'deal') { hideSuggest(); openDrawer(it.id); return; }
    else { var s = S.filters[it.k] || new Set(); s.add(it.v); S.filters[it.k] = s; S.q = ''; $('#q').value = ''; }
    $('#searchBox').classList.toggle('has', !!$('#q').value);
    hideSuggest(); S.page = 1; refresh();
  }
  var liveSearch = debounce(function () { S.q = $('#q').value.trim(); S.page = 1; refresh(); }, 250);
  $('#q').addEventListener('input', function () { $('#searchBox').classList.toggle('has', !!this.value); buildSuggest(); liveSearch(); });
  $('#q').addEventListener('focus', function () { if (this.value) buildSuggest(); });
  $('#q').addEventListener('keydown', function (e) {
    var n = $$('#suggest .sg').length, open = !$('#suggest').classList.contains('hidden');
    if (e.key === 'ArrowDown' && open) { e.preventDefault(); sugIdx = (sugIdx + 1) % n; markSug(); }
    else if (e.key === 'ArrowUp' && open) { e.preventDefault(); sugIdx = (sugIdx - 1 + n) % n; markSug(); }
    else if (e.key === 'Enter') { e.preventDefault(); if (open && sugIdx >= 0) pickSug(sugIdx); else { S.q = this.value.trim(); refresh(); } }
  });
  $('#suggest').addEventListener('mousedown', function (e) { var el = e.target.closest('.sg'); if (el) { e.preventDefault(); pickSug(+el.dataset.i); } });
  $('#qClear').addEventListener('click', function () { $('#q').value = ''; S.q = ''; $('#searchBox').classList.remove('has'); hideSuggest(); refresh(); });

  // ============================== KHỞI TẠO ==============================
  $$('#metricSeg button').forEach(function (b) {
    b.classList.toggle('on', b.dataset.m === S.metric);
    b.addEventListener('click', function () { S.metric = b.dataset.m; store('metric', S.metric); $$('#metricSeg button').forEach(function (x) { x.classList.toggle('on', x === b); }); updateCharts(); });
  });
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    $('#btnTheme').innerHTML = '<svg class="i"><use href="#i-' + (t === 'dark' ? 'sun' : 'moon') + '"/></svg>';
    if (S.deals.length) { renderCharts(true); renderDuo(); renderQuick(); renderAlert(); renderKpis(); }
  }
  var th = store('theme') || (window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  applyTheme(th);
  $('#btnTheme').addEventListener('click', function () { th = th === 'dark' ? 'light' : 'dark'; store('theme', th); applyTheme(th); });

  var dz = $('#dropZone'), fi = $('#fileInput');
  dz.addEventListener('click', function () { fi.click(); });
  dz.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  $('#btnNew').addEventListener('click', function () { fi.click(); });
  fi.addEventListener('change', function () { handleFile(fi.files[0]); fi.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); if (ev === 'drop' || !e.relatedTarget) dz.classList.remove('over'); }); });
  document.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });

  window.__crm = S; // hỗ trợ kiểm tra
})();
