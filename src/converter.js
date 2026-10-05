/*
 * converter.js – Bản JavaScript của convert_deals.py (v2).
 * Chuyển file export Deals từ CRM (raw) sang dữ liệu template phân tích, chạy hoàn toàn trong trình duyệt.
 * Giữ nguyên quy tắc, thứ tự kiểm tra và câu chữ cờ chất lượng của script Python gốc (reference/convert_deals.py).
 *
 * Quy ước ngày: mọi ngày giờ được lưu dưới dạng số mili-giây "giờ tường" (ghép thành phần ngày giờ bằng Date.UTC),
 * nên không bị lệch múi giờ khi tính số ngày và khi ghi ra Excel.
 */
(function (root) {
  'use strict';

  // ============================== THAM SỐ ==============================
  var SHEET_PREFIX = '1 - Bảng danh mục';
  var USD_MISLABEL_THRESHOLD = 1000000;
  var USD_RATE = 26000;
  var BOM_TOLERANCE = 0.05;
  var DEAL_SIZE = [[1e9, '< 1B'], [2e9, '1B - 2B'], [5e9, '2B - 5B']];
  var LEAD_STAGE = 'Đăng kí cơ hội';

  var B = 'Bảng danh mục sản phẩm dự án - ';
  var BOM_COLS = {};
  ['Hãng', 'Tên giải pháp', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Thời hạn bản quyền', 'Ghi chú', 'Sales user', 'Link deal']
    .forEach(function (k) { BOM_COLS[k] = B + k; });
  var BOM_CORE = ['Hãng', 'Tên giải pháp', 'Số lượng', 'Đơn giá', 'Thành tiền'].map(function (k) { return BOM_COLS[k]; });
  var ACTIVITY = ['Số email', 'Số cuộc họp', 'Số cuộc gọi', 'Số ghi chú', 'Số ghi nhanh'];
  var BOM_TXT = 'Thông tin BOM (model, số lượng, năm support...)';

  var REQUIRED = ['ID', 'Tên thương vụ', 'Giá trị thương vụ', 'Tiền tệ thương vụ', 'Chủ sở hữu thương vụ',
    'Giai đoạn', 'Trạng thái thương vụ', 'Tên công ty', 'Phân loại HĐ', 'Reseller', 'Tên hãng',
    BOM_TXT, 'Tên dự án', 'Timeline dự án',
    'Ngày ký hợp đồng', 'Phân loại nhóm deal', 'Lý do thất bại', 'Failed additional comment',
    'Từ lúc', 'Cập nhật lần cuối'].concat(ACTIVITY, Object.keys(BOM_COLS).map(function (k) { return BOM_COLS[k]; }));

  var VENDOR_MAP = {
    'KASPERSKY': 'Kaspersky', 'KAPERSKY': 'Kaspersky', 'GTB': 'GTB', 'SOPHOS': 'Sophos', 'OPSWAT': 'OPSWAT',
    'NETSCOUT': 'NetScout', 'DELINEA': 'Delinea', 'CLOUDFLARE': 'Cloudflare', 'PROGRESS': 'Progress',
    'INFOEXPRESS': 'Infoexpress', 'OPENTEXT': 'OpenText', 'BARRACUDA': 'Barracuda', 'RADWARE': 'Radware',
    'SAFETICA': 'Safetica', 'SAFEBREACH': 'SafeBreach', 'HCL SOFTWARE': 'HCL Software', 'HCL': 'HCL Software',
    'ZECURION': 'Zecurion', 'ACRONIS': 'Acronis', 'QUALYS': 'Qualys', 'STELLAR CYBER': 'Stellar Cyber',
    'PENTA SECURITY': 'Penta Security', 'ARCON': 'Arcon', 'PENTERA': 'Pentera', 'PENTARA': 'Pentera',
    'TRUSTWAVE': 'Trustwave', 'NETGEAR': 'Netgear', 'THALES': 'Thales', 'FORCEPOINT': 'Forcepoint',
    'SECPOD': 'SecPod', 'PROOFPOINT': 'Proofpoint', 'CYBLE': 'Cyble', 'FORTRA': 'Fortra', 'VIAVI': 'Viavi',
    'TRELLIX': 'Trellix', 'EFFICIENTIP': 'EfficientIP', 'HÃNG KHÁC': 'Khác', 'SẢN PHẨM KHÁC': 'Khác', 'KHÁC': 'Khác'
  };

  // Python \b là ranh giới từ Unicode; JS \b chỉ ASCII -> đổi sang lookaround Unicode để kết quả giống hệt.
  var W = '[\\p{L}\\p{N}_]';
  function rx(src, flags) {
    var b = '(?:(?<=' + W + ')(?!' + W + ')|(?<!' + W + ')(?=' + W + '))';
    return new RegExp(src.replace(/\\b/g, b), (flags || '') + 'u');
  }

  var VENDOR_KEYWORDS = [
    ['Kaspersky', 'kaspersky|kapersky|\\bkas\\b'], ['Sophos', 'sophos'], ['GTB', '\\bgtb\\b'],
    ['OPSWAT', 'opswat|metadefender'], ['NetScout', 'netscout|arbor'], ['Delinea', 'delinea'],
    ['Cloudflare', 'cloudflare'], ['Progress', 'progress|whatsup|\\bwug\\b|loadmaster|kemp'],
    ['Infoexpress', 'infoexpress|easynac'], ['OpenText', 'opentext|arcsight|fortify|\\bsmax\\b|\\bopds\\b|voltage|encase'],
    ['Barracuda', 'barracuda'], ['Radware', 'radware'], ['Safetica', 'safetica'], ['SafeBreach', 'safebreach'],
    ['HCL Software', '\\bhcl\\b|appscan|\\bhlc\\b'], ['Zecurion', 'zecurion'], ['Acronis', 'acronis'],
    ['Qualys', 'qualys'], ['Stellar Cyber', 'stellar'], ['Penta Security', 'penta'], ['Arcon', 'arcon'],
    ['Pentera', 'pentera|pentara'], ['Trellix', 'trellix'], ['SecPod', 'secpod'], ['Thales', 'thales'],
    ['Proofpoint', 'proofpoint'], ['Forcepoint', 'forcepoint'], ['Trustwave', 'trustwave'],
    ['Netgear', 'netgear'], ['Viavi', 'viavi'], ['EfficientIP', 'efficient ?ip'],
    ['Cyble', 'cyble'], ['Fortra', 'fortra']
  ].map(function (p) { return [p[0], rx(p[1])]; });

  var SOLUTION_KW = new Set(['waf', 'ddos', 'dr', 'fw', 'firewall', 'sophos', 'kaspersky', 'kapersky', 'kas', 'gtb', 'safetica',
    'safebreach', 'smax', 'loadmaster', 'opswat', 'qualys', 'progress', 'pam', 'secpod', 'renew',
    'foundations', 'optimum', 'appscan', 'hlc', 'endpoint', 'dlp', 'nac', 'easynac', 'trellix', 'delinea',
    'netscout', 'opds', 'cloudflare', 'opentext', 'barracuda', 'radware', 'infoexpress', 'acronis',
    'zecurion', 'stellar', 'pentera', 'fortify', 'arcsight', 'wug', 'whatsup', 'hcl', 'kemp', 'edr',
    'siem', 'mfa', 'ztna', 'appsec', 'db', 'security']);
  var FILLER = new Set(['va', 'and', 'plus', 'cap', 'do', 'bo', 'sung']);
  var PROJECT_PREFIX = rx('^(mua sam|giai phap|du an|trien khai|xay dung|nang cap|thay the)');
  var TEST_NAMES = rx('(^|[^a-z])test([^a-z]|$)|\\bko sale\\b|\\bkhong sale\\b');

  // ============================== HÀM TIỆN ÍCH ==============================
  function isStr(x) { return typeof x === 'string'; }
  function isNA(x) { return x === null || x === undefined || (typeof x === 'number' && isNaN(x)); }
  function notna(x) { return !isNA(x); }

  function fold(s) {
    return String(s).toLowerCase().normalize('NFD').replace(/đ/g, 'd').replace(/\p{Mn}/gu, '');
  }

  var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ndash: '–', mdash: '—', hellip: '…',
    lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', middot: '·', copy: '©', reg: '®', trade: '™',
    laquo: '«', raquo: '»', deg: '°', times: '×', divide: '÷', plusmn: '±', euro: '€', shy: '­', zwnj: '‌', zwj: '‍' };
  var _ta = null;
  function unescapeHtml(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);/g, function (m, e) {
      if (e[0] === '#') {
        var cp = (e[1] === 'x' || e[1] === 'X') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        try { return String.fromCodePoint(cp); } catch (err) { return '�'; }
      }
      if (Object.prototype.hasOwnProperty.call(ENT, e)) return ENT[e];
      if (typeof document !== 'undefined') {
        _ta = _ta || document.createElement('textarea');
        _ta.innerHTML = m; return _ta.value;
      }
      return m;
    });
  }

  function cleanText(x) {
    if (!isStr(x)) return x;
    x = unescapeHtml(x);
    x = x.replace(/<[^>]+>/g, ' ');
    x = x.replace(/\s+/g, ' ').trim();
    return x || null;
  }

  // Trả về số ms "giờ tường" hoặc null
  function parseDate(s) {
    if (s instanceof Date) {
      if (isNaN(s.getTime())) return null;
      return Date.UTC(s.getFullYear(), s.getMonth(), s.getDate(), s.getHours(), s.getMinutes(), s.getSeconds());
    }
    if (!isStr(s) || !s.trim()) return null;
    s = s.trim().replace(/-/g, '/');
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{1,2}))?$/.exec(s);
    if (!m) return null;
    var d = +m[1], mo = +m[2], y = +m[3], h = m[4] ? +m[4] : 0, mi = m[5] ? +m[5] : 0;
    if (mo < 1 || mo > 12 || d < 1 || h > 23 || mi > 59) return null;
    var t = Date.UTC(y, mo - 1, d, h, mi);
    if (new Date(t).getUTCDate() !== d) return null;
    return t;
  }

  // Làm tròn kiểu Python (nửa chẵn) cho số nguyên
  function roundHalfEven(y) {
    var f = Math.floor(y), diff = y - f;
    if (diff > 0.5) return f + 1;
    if (diff < 0.5) return f;
    return (f % 2 === 0) ? f : f + 1;
  }
  function group(n, sep) {
    var neg = n < 0 || Object.is(n, -0);
    var s = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, sep);
    return (neg && n !== 0 ? '-' : '') + s;
  }
  function vnd(x) { return group(roundHalfEven(x), '.'); }          // f"{x:,.0f}".replace(',', '.')
  function commaInt(x) { return group(roundHalfEven(x), ','); }     // f"{x:,.0f}"
  function pctSigned(r) {                                           // f"{r:+.0%}"
    var y = r * 100, neg = y < 0 || Object.is(y, -0);
    return (neg ? '-' : '+') + roundHalfEven(Math.abs(y)) + '%';
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(t) { var d = new Date(t); return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear(); }

  function pyTitle(s) {
    var out = '', prevCased = false;
    for (var ch of s) {
      var lo = ch.toLowerCase(), up = ch.toUpperCase(), cased = lo !== up;
      out += cased ? (prevCased ? lo : up) : ch;
      prevCased = cased;
    }
    return out;
  }
  function pyIsUpper(s) { return s.toUpperCase() === s && s.toLowerCase() !== s; }

  function canonVendor(v) {
    if (!v) return null;
    var k = v.trim().toUpperCase();
    return Object.prototype.hasOwnProperty.call(VENDOR_MAP, k) ? VENDOR_MAP[k] : pyTitle(v.trim());
  }

  function isSolution(seg) {
    var toks = fold(seg).split(/[^a-z0-9]+/).filter(Boolean);
    if (!toks.length) return false;
    var hit = toks.filter(function (t) { return SOLUTION_KW.has(t) || FILLER.has(t) || /^\d+$/.test(t); }).length;
    return hit >= Math.max(1, toks.length * 0.6);
  }

  var RX_CHO = rx('\\bcho\\s+(.+)$');
  var RX_DASH = /^(.*\S)\s*-\s*([^-]+)$/;
  function customerFromDealName(name) {
    var n = name.trim().replace(/;+$/, '').trim();
    var f = fold(n);
    if (TEST_NAMES.test(f)) return [n, 'test'];
    if (/^C12-BCA 175 - /.test(n)) return ['C12-BCA 175', 'tach'];
    var mc = RX_CHO.exec(n);
    if (PROJECT_PREFIX.test(f) && mc) return [mc[1].trim(), 'tach'];
    if (n.indexOf('_') >= 0) {
      var parts = n.split('_').map(function (p) { return p.trim(); }).filter(Boolean);
      var keep = parts.filter(function (p) { return !isSolution(p); });
      if (keep.length && keep.length < parts.length) return [keep.join(' - '), 'tach'];
      if (keep.length === parts.length) return [parts.join(' - '), 'giu'];
      return [n, 'kiem tra'];
    }
    var m = RX_DASH.exec(n);
    if (m && isSolution(m[2])) return [m[1].trim(), 'tach'];
    if (PROJECT_PREFIX.test(f)) return [n, 'kiem tra'];
    return [n, 'giu'];
  }

  function dealSize(v) {
    if (isNA(v) || v <= 0) return null;
    for (var i = 0; i < DEAL_SIZE.length; i++) if (v < DEAL_SIZE[i][0]) return DEAL_SIZE[i][1];
    return '> 5B';
  }

  var SERVICE_RX = rx('^\\s*dich vu|trien khai|cai dat|dao tao');
  var OTHER_RX = rx('giai phap khac|san pham khac|hang khac');
  var NON_MAIN = ['Khác', 'Dịch vụ'];
  var SRC_RANK = { 'BOM': 0, 'Tên giải pháp/Ghi chú': 1, 'Tên hãng (CRM)': 2, 'Từ khóa': 3, 'Không xác định': 4 };

  var RX_FLOAT = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
  function num(v) {
    if (isNA(v)) return [null, false];
    if (typeof v === 'number') return [v, false];
    var s = String(v).trim();
    if (s[0] === '#') return [null, true];
    s = s.replace(/,/g, '');
    if (RX_FLOAT.test(s)) return [parseFloat(s), false];
    return [null, true];
  }

  function keywordVendors(text) {
    var f = fold(text), found = [];
    VENDOR_KEYWORDS.forEach(function (p) { var m = p[1].exec(f); if (m) found.push([m.index, p[0]]); });
    found.sort(function (a, b) { return a[0] - b[0] || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0); });
    var out = [];
    found.forEach(function (x) { if (out.indexOf(x[1]) < 0) out.push(x[1]); });
    return out;
  }

  function dealLevelVendors(r) {
    var col = isStr(r['Tên hãng']) ? canonVendor(r['Tên hãng']) : null;
    if (col && col !== 'Khác') return [[col], 'Tên hãng (CRM)'];
    var text = [BOM_TXT, 'Tên thương vụ', 'Tên dự án'].filter(function (c) { return isStr(r[c]); })
      .map(function (c) { return r[c]; }).join(' ');
    var kw = keywordVendors(text);
    if (kw.length) return [kw, 'Từ khóa'];
    if (col === 'Khác') return [['Khác'], 'Tên hãng (CRM)'];
    return [[], 'Không xác định'];
  }

  function lineVendor(line, dealRow) {
    var h = line[BOM_COLS['Hãng']];
    if (isStr(h) && h.trim()) return [canonVendor(h), 'BOM'];
    var text = ['Tên giải pháp', 'Ghi chú'].filter(function (c) { return isStr(line[BOM_COLS[c]]); })
      .map(function (c) { return line[BOM_COLS[c]]; }).join(' ');
    if (text) {
      var kw = keywordVendors(text);
      if (kw.length) return [kw[0], 'Tên giải pháp/Ghi chú'];
      if (SERVICE_RX.test(fold(text))) return ['Dịch vụ', 'Tên giải pháp/Ghi chú'];
      if (OTHER_RX.test(fold(text))) return ['Khác', 'Tên giải pháp/Ghi chú'];
    }
    var dv = dealLevelVendors(dealRow);
    return dv[0].length ? [dv[0][0], dv[1]] : [null, 'Không xác định'];
  }

  // ============================== DANH MỤC SALE ==============================
  var SALE_NAMES = {
    'thuydinh': 'Đinh Văn Thủy', 'hangnguyen': 'Nguyễn Thu Hằng', 'trinhluong': 'Lương Thị Tuyết Trinh',
    'phamquyet': 'Phạm Văn Quyết', 'theduong': 'Dương Văn Thế', 'quyvy': 'Vy Công Quý',
    'kienkhong': 'Khổng Đức Kiên', 'nhungnguyen': 'Nguyễn Hồng Nhung', 'hungle': 'Lê Huy Hùng',
    'anhle': 'Lê Tuấn Anh', 'thangnguyen': 'Nguyễn Duy Thắng', 'nghiemphan': 'Phan Quế Nghiệm',
    'hungpham': 'Phạm Quốc Hùng', 'tupham': 'Phạm Cao Anh Tú', 'hieudinh': 'Đinh Văn Hiệu',
    'hiendoan': 'Đoàn Thị Hiền', 'hungtran': 'Trần Quang Hưng',
    'truongdang': 'Đặng Đình Trường'
  };
  var SALE_TAGS = null;
  function saleTags() {
    var tags = new Set();
    Object.keys(SALE_NAMES).forEach(function (user) {
      var p = fold(SALE_NAMES[user]).split(/\s+/).filter(Boolean);
      var fam = p[0], mid = p.slice(1, -1), giv = p[p.length - 1];
      [user, fam + giv, giv + fam, giv, p.join(''), mid.slice(-1).join('') + giv].forEach(function (t) { tags.add(t); });
      if (p.length >= 2) tags.add(p.slice(-2).join(''));
    });
    return new Set(Array.from(tags).filter(function (t) { return t.length >= 4; }));
  }
  function stripSaleTag(name) {
    var m = /^(.*\S)\s*_\s*([^_]+)$/.exec(name);
    SALE_TAGS = SALE_TAGS || saleTags();
    if (m && SALE_TAGS.has(fold(m[2]).replace(/[^a-z]/g, ''))) return [m[1].trim(), m[2].trim()];
    return [name, null];
  }

  // ============================== PHÂN NHÓM KHÁCH HÀNG ==============================
  var SEGMENTS = ['Cơ quan Đảng, Nhà nước', 'Bộ, Ngành, Cơ quan TW', 'UBND / Sở ban ngành địa phương',
    'An ninh - Quốc phòng', 'Ngân hàng - Tài chính - Bảo hiểm', 'Viễn thông - CNTT',
    'Năng lượng - Điện - Dầu khí', 'Y tế - Giáo dục', 'Sản xuất - Thương mại - Dịch vụ',
    'Giao thông - Vận tải - Logistics', 'DNNN khác', 'Khác'];
  var SEGMENT_RULES = [
    ['Viễn thông - CNTT', 'viettel|\\bvtg\\b|natcom'],
    ['An ninh - Quốc phòng', 'cong an|\\bbca\\b|\\bbo ca\\b|^ca |\\bca (tp|tinh|thanh pho)\\b|\\bpa0\\d|' +
      '\\b(a0[1-9]|c0[1-9]|c1[0-2]|c500|h0[1-9]|v0[1-9]|b01|x01)\\b|quoc phong|\\bbqp\\b|quan doi|' +
      'quan khu|bo tu lenh|binh chung|bien phong|canh sat|quan y|hai quan|military|' +
      'tong cuc (2|ii|tinh bao|ky thuat|chinh tri)|co yeu|si quan|tong tham muu|tac chien|' +
      'an ninh nhan dan|bo tham muu|bo doi'],
    ['Cơ quan Đảng, Nhà nước', 'tinh uy|thanh uy|quan uy|huyen uy|\\bvptu\\b|trung uong dang|tw dang|vptw|' +
      'tap chi cong san|phu chu tich|van phong chu tich nuoc|quoc hoi|toa an|vien kiem sat|' +
      'kiem toan nha nuoc'],
    ['UBND / Sở ban ngành địa phương', '\\bubnd\\b|uy ban nhan dan|^so |\\bso (khoa|kh|thong tin|tttt|4t|xay dung|' +
      'dan toc|tai chinh|y te|giao duc|noi vu|tu phap|ke hoach|cong thuong|van hoa)\\b|' +
      'bqlda|ban quan ly khu kinh te|ban duy tu|chuyen doi so|' +
      'phuc vu hanh chinh cong|dai phat thanh|bao va phat thanh'],
    ['Bộ, Ngành, Cơ quan TW', '^bo |\\bbo (noi vu|tai chinh|tai nguyen|ngoai giao|dan toc|xay dung|khoa hoc|' +
      'cong thuong|y te|giao duc|tu phap|ke hoach|nong nghiep|giao thong|van hoa|thong tin)\\b|' +
      '\\bcuc\\b|tong cuc|kho bac|bao hiem xa hoi|\\bbhxh\\b|ngan hang nha nuoc|' +
      '(uy ban|ub) chung khoan|uy ban tieu chuan|quy phat trien|thong tin tin dung|' +
      '\\bcic\\b|vnnic|ttxvn|thong tan xa|vien han lam'],
    ['Năng lượng - Điện - Dầu khí', '\\bevn|dien luc|thuy dien|nhiet dien|truyen tai dien|he thong dien|' +
      'thi truong dien|dieu do|\\bnpt\\b|\\bptc ?\\d|^pc |genco|\\bpvn\\b|pv ?gas|petro|' +
      'dau khi|khi viet nam|^than |\\btkv\\b|vimico|khoang san|nhien lieu|xang dau|nsrp|' +
      'npcit|\\bnpc\\b|\\bspc\\b'],
    ['Y tế - Giáo dục', 'benh vien|^bv|\\bbvdk\\b|y te|huyet hoc|^truong |dai hoc|hoc vien|giao duc|educa|' +
      '\\bttyt\\b|hospital'],
    ['Ngân hàng - Tài chính - Bảo hiểm', 'ngan hang|bank|chung khoan|securities|bao hiem|insurance|tai chinh vi mo|' +
      'finance|napas|vnpay|thanh toan|\\b(bidv|vietcombank|vietinbank|agribank|' +
      'mbbank|tpbank|vib|scb|shb|shbfc|lpb|lpbs|exim|eximbank|tcbs|hsc|kbsv|' +
      'vndirect|mbs|vps|shs|dnse|dsc|agriseco|opes|pjico|pvi|insmart|mfinance|' +
      'timo|oceanbank|publicbank)\\b'],
    ['Giao thông - Vận tải - Logistics', 'hang khong|airline|airport|san bay|\\bacv\\b|vaeco|viags|sasco|\\bvna\\b|' +
      'truc thang|\\bvnh\\b|quan ly bay|vnaic|hang hai|giao hang|\\bghtk\\b|' +
      'logistics|van tai|duong sat|\\bvetc\\b|\\bitl\\b|vnpost|buu chinh'],
    ['Viễn thông - CNTT', 'vnpt|mobifone|\\bcmc\\b|telecom|vien thong|\\bfpt\\b|\\bfis\\b|\\bctin\\b|cong nghe|technology|' +
      'tecapro|gosu|vtvcab|software|vietbay|tntech'],
    ['DNNN khác', 'cap nuoc|sawaco|vietlott|xo so|trac dia|dong tau'],
    ['Sản xuất - Thương mại - Dịch vụ', 'cong ty|\\bcty\\b|\\bctcp\\b|tnhh|tap doan|group|corporation|\\bcorp\\b|' +
      'limited|\\bjsc\\b|holdings']
  ].map(function (p) { return [p[0], rx(p[1])]; });
  var SX = 'Sản xuất - Thương mại - Dịch vụ';
  var SEGMENT_OVERRIDES = {
    'tong cong ty cp buu chinh viettel': 'Giao thông - Vận tải - Logistics',
    'viettel post': 'Giao thông - Vận tải - Logistics',
    'cong ty thuong mai & xuat nhap khau/ thuong mai & xnk viettel (viettel commerce)': SX,
    'cong ty lien doanh thap ngan hang dau tu va phat trien viet nam': SX,
    'vietinbank gold & jewellery': 'Ngân hàng - Tài chính - Bảo hiểm',
    'pvi': 'Ngân hàng - Tài chính - Bảo hiểm',
    'cong ty tnhh mot thanh vien dong tau hong ha': 'DNNN khác',
    'cong ty tnhh mtv trac dia ban do': 'DNNN khác',
    'tap doan phenikaa': SX,
    'access trade': SX, 'career viet': SX, 'diana unicharm': SX, 'everland van don': SX,
    'fujikin': SX, 'freshmart (c.p)': SX, 'fujimart': SX, 'jarllytec': SX, 'maison': SX, 'masterise': SX,
    'messer viet nam': SX, 'nutifood': SX, 'panasonic r&d': SX, 'rang dong': SX, 'shemar power vn': SX,
    'sun group': SX, 'yakult viet nam': SX, 'yokohama tyre viet nam': SX, 'berjaya gia thinh': SX,
    'viet duc co so 3': 'Y tế - Giáo dục',
    'cong ty tnhh mtv van hanh htd va ttd quoc gia': 'Năng lượng - Điện - Dầu khí',
    'gtb endpoint protector': 'Khác', 'giai phap attt cap do 3': 'Khác'
  };
  var SEGMENT_UNSURE = new Set(['hal vn', 'mcst', 'itdb', 'lpex', 'plc', 'intech',
    'trung tam nghien cuu va ung dung cong nghe truyen thong (r&d)']);

  function classifyCustomer(name) {
    var f = fold(name).replace(/\s+/g, ' ').trim();
    if (SEGMENT_UNSURE.has(f)) return [(f.indexOf('cong nghe') >= 0 ? 'Viễn thông - CNTT' : 'Khác'), false];
    if (Object.prototype.hasOwnProperty.call(SEGMENT_OVERRIDES, f)) return [SEGMENT_OVERRIDES[f], true];
    for (var i = 0; i < SEGMENT_RULES.length; i++) {
      var seg = SEGMENT_RULES[i][0];
      if (SEGMENT_RULES[i][1].test(f)) {
        if (seg === 'Viễn thông - CNTT' && /viettel/.test(f)) {
          if (/buu chinh|post/.test(f)) return ['Giao thông - Vận tải - Logistics', true];
          if (/commerce|thuong mai/.test(f)) return [SX, true];
        }
        return [seg, true];
      }
    }
    return ['Khác', false];
  }

  // ============================== ĐỌC FILE NGUỒN ==============================
  function ConvertError(msg) { var e = new Error(msg); e.name = 'ConvertError'; return e; }

  // Chuỗi mà pandas.read_excel mặc định coi là ô trống (na_values) – giữ để kết quả giống script Python
  var PD_NA = new Set(['', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan', '1.#IND', '1.#QNAN',
    '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'None', 'n/a', 'nan', 'null']);
  function sheetRows(XLSX, ws) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    while (aoa.length && aoa[aoa.length - 1].every(isNA)) aoa.pop();
    if (!aoa.length) return { header: [], rows: [] };
    var seen = {};
    var header = aoa[0].map(function (h, j) {
      var name = isNA(h) ? 'Unnamed: ' + j : String(h);
      if (seen[name] !== undefined) { seen[name]++; name = name + '.' + seen[name]; } else seen[name] = 0;
      return name;
    });
    var rows = aoa.slice(1).map(function (a) {
      var o = {};
      header.forEach(function (h, j) {
        var v = a[j] === undefined ? null : a[j];
        o[h] = (isStr(v) && PD_NA.has(v)) ? null : v;
      });
      return o;
    });
    return { header: header, rows: rows };
  }

  function findSourceSheet(wb) {
    var names = wb.SheetNames.filter(function (s) { return s.indexOf(SHEET_PREFIX) === 0; });
    return names.length ? names[0] : null;
  }

  // ============================== XỬ LÝ CHÍNH ==============================
  function dealBom(lines, dealRow) {
    return lines.map(function (ln) {
      var a = num(ln[BOM_COLS['Thành tiền']]), amt = a[0], err = a[1], computed = false;
      if (amt === null) {
        var q = num(ln[BOM_COLS['Số lượng']])[0], p = num(ln[BOM_COLS['Đơn giá']])[0];
        if (q !== null && p !== null) { amt = q * p; computed = true; }
      }
      var lv = lineVendor(ln, dealRow);
      var link = ln[BOM_COLS['Link deal']];
      var m = isStr(link) ? /\/(\d+)\s*$/.exec(link) : null;
      var su = ln[BOM_COLS['Sales user']];
      return { vendor: lv[0], src: lv[1], amount: amt, err: err, computed: computed,
        sales: isStr(su) && su.trim() ? su.trim().toLowerCase() : null,
        link: m ? parseInt(m[1], 10) : null };
    });
  }

  function vendorRows(items, dealRow) {
    if (!items.length) {
      var dv = dealLevelVendors(dealRow), vs = dv[0];
      var rows0 = vs.map(function (v) { return [v, dv[1], null]; });
      return [rows0.length ? rows0 : [[null, 'Không xác định', null]], vs.length ? vs[0] : null, null];
    }
    var order = [], amt = new Map(), cnt = new Map(), srcs = new Map();
    items.forEach(function (it) {
      var v = it.vendor;
      if (order.indexOf(v) < 0) order.push(v);
      cnt.set(v, (cnt.get(v) || 0) + 1);
      if (it.amount !== null) amt.set(v, (amt.has(v) ? amt.get(v) : 0) + it.amount);
      if (!srcs.has(v)) srcs.set(v, []);
      srcs.get(v).push(it.src);
    });
    var A = function (v) { return amt.has(v) ? amt.get(v) : 0; };
    var known = order.filter(function (v) { return v !== null; });
    var pool = known.filter(function (v) { return NON_MAIN.indexOf(v) < 0; });
    if (!pool.length) pool = known;
    if (!pool.length) return [[[null, 'Không xác định', amt.has(null) ? amt.get(null) : null]], null, null];
    var main = pool[0];
    pool.forEach(function (v) {
      if (A(v) > A(main) || (A(v) === A(main) && cnt.get(v) > cnt.get(main))) main = v;
    });
    var how = A(main) > 0 ? 'giá trị BOM lớn nhất' : 'nhiều dòng BOM nhất (BOM chưa có giá)';
    var rest = known.filter(function (v) { return v !== main; }).sort(function (a, b) { return -A(a) - -A(b); });
    if (order.indexOf(null) >= 0) rest.push(null);
    var rows = [main].concat(rest).map(function (v) {
      var best = srcs.get(v).reduce(function (x, y) { return SRC_RANK[y] < SRC_RANK[x] ? y : x; });
      return [v, best, amt.has(v) ? amt.get(v) : null];
    });
    return [rows, main, how];
  }

  function isLead(r, items) {
    var acts = ACTIVITY.reduce(function (s, c) { return s + (notna(r[c]) ? parseFloat(r[c]) : 0); }, 0);
    var v0 = r['Giá trị thương vụ'];
    return r['Trạng thái thương vụ'] === 'Active' && r['Giai đoạn'] === LEAD_STAGE && !items.length &&
      (isNA(v0) || v0 === 0) && r['Timeline dự án'] === null && acts === 0;
  }

  function toNumber(v) {
    if (isNA(v)) return null;
    if (typeof v === 'number') return v;
    var n = num(v)[0];
    return n === null ? null : n;
  }

  /**
   * convert(XLSX, workbook, refDate) -> { rows, splitFlags, stats, warnings }
   * refDate: ms "giờ tường" (Date.UTC) của ngày chốt số liệu.
   */
  function convert(XLSX, wb, refDate) {
    var warnings = [];
    var sheet = findSourceSheet(wb);
    if (!sheet) throw ConvertError('Không tìm thấy sheet bắt đầu bằng "' + SHEET_PREFIX + '". Các sheet có: ' + wb.SheetNames.join(', '));
    var sr = sheetRows(XLSX, wb.Sheets[sheet]);
    var missing = REQUIRED.filter(function (c) { return sr.header.indexOf(c) < 0; });
    if (missing.length) throw ConvertError('Sheet "' + sheet + '" thiếu cột: ' + missing.join(', ') + '. Kiểm tra lại cấu hình export CRM.');
    var raw = sr.rows;
    raw.forEach(function (r, i) {
      for (var k in r) if (isStr(r[k])) r[k] = cleanText(r[k]);
      r._row = i + 2;
    });

    // ---- tách deal / dòng BOM
    var cur = null, bomVals = new Set(BOM_CORE), dealCols = sr.header.filter(function (c) { return !Object.values(BOM_COLS).includes(c); });
    var deals = [], lines = [], leaked = [];
    raw.forEach(function (r) {
      if (notna(r['ID'])) cur = r['ID'];
      r._ID = cur;
      if (notna(r['ID'])) deals.push(r);
      else if (dealCols.some(function (c) { return notna(r[c]); })) leaked.push(r._row);
    });
    if (leaked.length) warnings.push('CẢNH BÁO: ' + leaked.length + ' dòng BOM nối tiếp (ID trống) có dữ liệu ở cột thông tin deal – kiểm tra dòng Excel ' + leaked.slice(0, 10).join(', '));
    if (raw.some(function (r) { return r._ID === null; })) throw ConvertError('Sheet nguồn có dòng BOM nằm trước dòng deal đầu tiên (không xác định được ID) – dừng để kiểm tra.');
    raw.forEach(function (r) {
      if (BOM_CORE.some(function (c) { return notna(r[c]); })) lines.push(r);
    });
    bomVals = null;
    var idSeen = new Set(), dup = [];
    deals.forEach(function (r) {
      var id = Math.trunc(Number(r['ID']));
      if (isNaN(id)) throw ConvertError('ID deal không phải số: "' + r['ID'] + '" (dòng Excel ' + r._row + ')');
      if (idSeen.has(id)) dup.push(id);
      idSeen.add(id);
      r['ID'] = id;
    });
    if (dup.length) throw ConvertError('Sheet nguồn có ID deal trùng: ' + dup.slice(0, 10).join(', ') + ' – dừng để kiểm tra.');
    lines.forEach(function (r) { r._ID = Math.trunc(Number(r._ID)); });
    var nRawRows = raw.length, nDealsRaw = deals.length, nLinesRaw = lines.length;

    deals.forEach(function (r) {
      r['Giá trị thương vụ'] = toNumber(r['Giá trị thương vụ']);
    });
    ['Timeline dự án', 'Ngày ký hợp đồng', 'Từ lúc', 'Cập nhật lần cuối'].forEach(function (c) {
      var bad = [];
      deals.forEach(function (r) {
        var p = parseDate(r[c]);
        if (notna(r[c]) && p === null) bad.push(r[c]);
        r[c] = p;
      });
      if (bad.length) warnings.push('CẢNH BÁO: ' + bad.length + ' giá trị ngày không đọc được ở cột "' + c + '": ' + bad.slice(0, 5).join(', '));
    });

    // ---- Bước 1: Tên khách hàng + loại deal test
    var testIds = new Set(), testNames = [];
    deals.forEach(function (r) {
      var nm = String(r['Tên thương vụ'] === null ? '' : r['Tên thương vụ']);
      if (TEST_NAMES.test(fold(nm))) { testIds.add(r['ID']); testNames.push(nm); }
      var c, how, note;
      if (isStr(r['Tên công ty']) && r['Tên công ty']) { c = r['Tên công ty']; note = null; }
      else {
        var x = customerFromDealName(nm); c = x[0]; how = x[1];
        if (how === 'giu' && isSolution(c)) how = 'kiem tra';
        if (how === 'test' || how === 'giu') note = how === 'test' ? null : 'Tên KH lấy từ tên thương vụ (CRM trống "Tên công ty")';
        else if (how === 'tach') note = 'Tên KH lấy từ tên thương vụ (CRM trống "Tên công ty"), đã bỏ phần hãng/giải pháp/reseller';
        else note = 'Tên KH lấy nguyên tên thương vụ nhưng tên giống tên dự án/giải pháp – cần kiểm tra lại';
      }
      var st = stripSaleTag(c);
      r._cust = st[0]; r._sale_tag = st[1]; r._cust_note = note;
    });
    var df = deals.filter(function (r) { return !testIds.has(r['ID']); });
    lines = lines.filter(function (r) { return !testIds.has(r._ID); });

    // chuẩn hóa tên KH theo khóa (bỏ dấu, chỉ chữ số)
    var groups = new Map();
    df.forEach(function (r) {
      r._key = fold(r._cust).replace(/[^a-z0-9]+/g, ' ').trim();
      if (!groups.has(r._key)) groups.set(r._key, new Map());
      var g = groups.get(r._key); g.set(r._cust, (g.get(r._cust) || 0) + 1);
    });
    var stdName = new Map();
    groups.forEach(function (vc, key) {
      var cands = Array.from(vc.keys());
      cands.sort(function (a, b) {
        var ka = [fold(a) === a.toLowerCase() ? 1 : 0, pyIsUpper(a) ? 1 : 0, -vc.get(a)];
        var kb = [fold(b) === b.toLowerCase() ? 1 : 0, pyIsUpper(b) ? 1 : 0, -vc.get(b)];
        for (var i = 0; i < 3; i++) if (ka[i] !== kb[i]) return ka[i] - kb[i];
        return a < b ? -1 : a > b ? 1 : 0;
      });
      stdName.set(key, cands[0]);
    });
    df.forEach(function (r) { r._cust_std = stdName.get(r._key); });

    var byDeal = new Map();
    lines.forEach(function (ln) { if (!byDeal.has(ln._ID)) byDeal.set(ln._ID, []); byDeal.get(ln._ID).push(ln); });
    var nameCounts = new Map();
    df.forEach(function (r) {
      var k = r['Tên thương vụ'];
      if (!nameCounts.has(k)) nameCounts.set(k, []);
      nameCounts.get(k).push(r['ID']);
    });
    var out = [], splitFlags = [];
    var stCnt = { lines: 0, value_err: 0, value_fixed: 0, unresolved: 0, inferred: 0 };

    df.forEach(function (r) {
      var items = dealBom(byDeal.get(r['ID']) || [], r);
      var lead = isLead(r, items);
      var vr = vendorRows(items, r), rows = vr[0], main = vr[1], how = vr[2];
      var priced = items.filter(function (it) { return it.amount !== null; }).map(function (it) { return it.amount; });
      var bom = priced.reduce(function (s, x) { return s + x; }, 0);
      stCnt.lines += items.length;
      var issues = [], notes = [], leadKeep = [];

      // ---- giá trị
      var v0 = r['Giá trị thương vụ'], curr = r['Tiền tệ thương vụ'], val;
      if ((isNA(v0) || v0 === 0) && bom > 0) { val = bom; notes.push('Giá trị lấy từ tổng BOM (' + vnd(bom) + ') vì giá trị CRM = 0'); }
      else if (isNA(v0) || v0 === 0) { val = null; issues.push('Thiếu giá trị thương vụ (CRM = 0, BOM không có giá)'); }
      else if (curr === 'USD' && v0 < USD_MISLABEL_THRESHOLD) { val = v0 * USD_RATE; notes.push('Quy đổi ' + commaInt(v0) + ' USD x ' + commaInt(USD_RATE) + ' (tỷ giá giả định)'); }
      else {
        val = v0;
        if (curr === 'USD') notes.push('CRM ghi tiền tệ USD nhưng số tiền là VND – giữ số gốc, coi là VND');
      }
      if (val && bom > 0 && Math.abs(val - bom) > BOM_TOLERANCE * bom)
        issues.push('Giá trị ' + vnd(val) + ' lệch tổng BOM ' + vnd(bom) + ' (' + pctSigned((val - bom) / bom) + ')');

      // ---- chất lượng dòng BOM
      var nErr = items.filter(function (it) { return it.err; }).length;
      var nFix = items.filter(function (it) { return it.computed; }).length;
      var nUnpriced = items.filter(function (it) { return it.amount === null; }).length;
      stCnt.value_err += nErr; stCnt.value_fixed += nFix;
      if (nErr) issues.push('BOM có ' + nErr + ' dòng Thành tiền lỗi công thức CRM (#VALUE!)');
      if (nFix) notes.push(nFix + ' dòng BOM thiếu Thành tiền – đã tính = Số lượng x Đơn giá');
      if (nUnpriced) issues.push('BOM có ' + nUnpriced + ' dòng chưa có giá');
      var blank = items.filter(function (it) { return it.src !== 'BOM'; });
      if (blank.length) {
        stCnt.inferred += blank.length;
        issues.push('BOM có ' + blank.length + ' dòng trống cột Hãng');
        var res = blank.filter(function (it) { return it.vendor; });
        var unres = blank.length - res.length;
        stCnt.unresolved += unres;
        if (res.length) {
          var desc = Array.from(new Set(res.map(function (it) { return it.vendor + ' (suy từ ' + it.src + ')'; }))).sort().join(', ');
          notes.push('Hãng của ' + res.length + ' dòng BOM trống đã suy: ' + desc);
        }
        if (unres) issues.push(unres + ' dòng BOM không xác định được hãng');
      }
      if (!items.length && !lead) {
        issues.push('Deal không có dòng BOM');
        if (main) notes.push('Hãng suy từ ' + rows[0][1]);
      }

      // ---- sale, link deal
      var u = r['Chủ sở hữu thương vụ'];
      var uk = isStr(u) ? u.trim().toLowerCase() : null;
      var sale;
      if (uk !== null && Object.prototype.hasOwnProperty.call(SALE_NAMES, uk)) sale = SALE_NAMES[uk];
      else {
        sale = isStr(u) ? u : null;
        if (isStr(u)) leadKeep.push('Sale "' + u + '" chưa có trong bảng tên sale – giữ tên tài khoản CRM');
      }
      var su = Array.from(new Set(items.filter(function (it) { return it.sales && it.sales !== uk; }).map(function (it) { return it.sales; }))).sort();
      if (su.length) issues.push('Sales user ở dòng BOM (' + su.map(function (s) { return SALE_NAMES[s] || s; }).join(', ') + ') khác chủ deal (' + (sale === null ? 'None' : sale) + ')');
      var links = Array.from(new Set(items.filter(function (it) { return it.link && it.link !== r['ID']; }).map(function (it) { return it.link; }))).sort(function (a, b) { return a - b; });
      if (links.length) issues.push('Link deal ở dòng BOM trỏ sang deal khác (ID ' + links.join(', ') + ') – BOM có thể copy từ deal khác');

      // ---- trạng thái / giai đoạn / close date
      var st = r['Trạng thái thương vụ'], gd = r['Giai đoạn'];
      if (st === 'Active' && gd === 'Thực hiện hợp đồng') issues.push('Trạng thái Active nhưng giai đoạn CRM là "Thực hiện hợp đồng" (có thể đã Won)');
      if (st === 'Won' && gd !== 'Thực hiện hợp đồng') issues.push('Trạng thái Won nhưng giai đoạn CRM là "' + (gd === null ? 'None' : gd) + '"');
      if (st === 'Won' && r['Ngày ký hợp đồng'] === null) issues.push('Won nhưng thiếu ngày ký hợp đồng');
      if (st === 'Lost' && !isStr(r['Lý do thất bại'])) issues.push('Lost nhưng thiếu lý do thất bại');
      var cd = r['Timeline dự án'];
      var overdue = (cd !== null && st === 'Active' && cd < refDate) ? Math.floor((refDate - cd) / 86400000) : null;
      if (overdue !== null) issues.push('Close date ' + fmtDate(cd) + ' đã qua ' + overdue + ' ngày nhưng deal vẫn Active');
      var same = nameCounts.get(r['Tên thương vụ']).filter(function (i) { return i !== r['ID']; });
      if (same.length) issues.push('Trùng tên thương vụ với ' + same.length + ' deal khác (ID ' + same.slice(0, 5).join(', ') + (same.length > 5 ? ' …' : '') + ')');
      if (!isStr(u)) issues.push('Thiếu Sale (owner) trong CRM');
      var miss = [['Close date', cd !== null], ['Reseller', isStr(r['Reseller'])], ['Phân loại HĐ', isStr(r['Phân loại HĐ'])]]
        .filter(function (x) { return !x[1]; }).map(function (x) { return x[0]; });
      if (miss.length) issues.push('Thiếu thông tin: ' + miss.join(', '));
      if (main === null && !lead) issues.push('Chưa xác định được hãng (BOM, Tên giải pháp/Ghi chú, Tên hãng và từ khóa đều không có)');
      if (how && rows.length > 1) notes.push('Hãng chính = ' + main + ' (theo ' + how + ')');

      // ---- tên KH, deal size, phân nhóm, failed
      if (isStr(r._cust_note)) (r._cust_note.indexOf('cần kiểm tra') >= 0 ? leadKeep : notes).push(r._cust_note);
      if (isStr(r._sale_tag)) notes.push('Đã bỏ hậu tố tên sale "_' + r._sale_tag + '" trong tên KH');
      if (r._cust_std !== r._cust) notes.push('Tên KH chuẩn hóa từ "' + r._cust + '"');
      var ds = dealSize(val), ds0 = r['Phân loại nhóm deal'];
      if (ds && !isStr(ds0)) notes.push('Deal size tính từ giá trị (CRM trống)');
      else if (ds && ds0 !== ds) notes.push('Deal size CRM "' + ds0 + '" không khớp giá trị – đã tính lại');
      var sc = classifyCustomer(r._cust_std), seg = sc[0];
      if (!sc[1]) leadKeep.push('Phân nhóm KH "' + seg + '" chưa chắc chắn – cần kiểm tra');
      var reasons = [r['Lý do thất bại'], r['Failed additional comment']].filter(isStr);
      var failed = reasons.length ? reasons.join(' – ') : null;

      if (lead) issues = ['Lead chiến dịch – chưa có thông tin cơ hội (chưa có BOM, giá trị, Close date, hoạt động)'].concat(leadKeep);
      else issues = issues.concat(leadKeep);

      var n = rows.length;
      rows.forEach(function (row, k) {
        var extra = [];
        if (n > 1) {
          extra = k === 0
            ? ['Deal có ' + n + ' dòng hãng: ' + rows.map(function (x) { return x[0] || '(trống)'; }).join(', ') + ' – giá trị thương vụ ghi ở dòng này']
            : ['Dòng tách hãng ' + (k + 1) + '/' + n + ' của deal ID ' + r['ID'] + '; giá trị thương vụ chỉ ghi ở dòng hãng chính (' + rows[0][0] + ')'];
        }
        var text = [];
        if (issues.length) text.push('[VẤN ĐỀ] ' + issues.join('; '));
        if (notes.length || extra.length) text.push('[ĐÃ XỬ LÝ] ' + extra.concat(notes).join('; '));
        out.push({
          'ID': r['ID'], 'Tên thương vụ': r['Tên thương vụ'], 'Loại deal': lead ? 'Lead chiến dịch' : 'Cơ hội',
          'Tên khách hàng': r._cust_std, 'Phân nhóm khách hàng': seg, 'Sale': sale, 'Reseller': r['Reseller'],
          'Giai đoạn': gd, 'Trạng thái': st, 'Giá trị thương vụ': k === 0 ? val : null,
          'Hãng': row[0], 'Nguồn hãng': row[1], 'Giá trị BOM theo hãng': row[2],
          'Thông tin BOM (model, số lượng, năm support...)': r[BOM_TXT],
          'Phân loại HĐ': r['Phân loại HĐ'], 'Phân loại deal size': ds, 'Close date': cd,
          'Số ngày quá hạn forecast': overdue, 'Lý do Failed': failed, 'Ngày tạo': r['Từ lúc'],
          'Cập nhật lần cuối': r['Cập nhật lần cuối'], 'Cờ chất lượng': text.join(' | ') || null
        });
        splitFlags.push(k > 0);
      });
    });

    // ---- Đối soát
    var firstSeen = new Set(), first = [];
    out.forEach(function (o) { if (!firstSeen.has(o.ID)) { firstSeen.add(o.ID); first.push(o); } });
    var sum = function (arr, f) { return arr.reduce(function (s, o) { var v = f(o); return s + (isNA(v) ? 0 : v); }, 0); };
    var stats = {
      sheet: sheet, raw_rows: nRawRows, deals_raw: nDealsRaw, lines_raw: nLinesRaw,
      tests: testIds.size, test_names: testNames, deals: firstSeen.size, rows: out.length,
      lead: first.filter(function (o) { return o['Loại deal'] === 'Lead chiến dịch'; }).length,
      co_hoi: first.filter(function (o) { return o['Loại deal'] === 'Cơ hội'; }).length,
      lines_used: stCnt.lines, lines_after_test: lines.length,
      value_total: sum(first, function (o) { return o['Giá trị thương vụ']; }),
      value_cohoi: sum(first.filter(function (o) { return o['Loại deal'] === 'Cơ hội'; }), function (o) { return o['Giá trị thương vụ']; }),
      bom_total: sum(out, function (o) { return o['Giá trị BOM theo hãng']; })
    };
    Object.keys(stCnt).forEach(function (k) { stats[k] = stCnt[k]; });
    if (stats.lines_used !== stats.lines_after_test) throw ConvertError('Số dòng BOM đã đọc khác số dòng BOM trong file nguồn');
    if (stats.deals !== nDealsRaw - testIds.size) throw ConvertError('Số deal output khác số deal raw trừ deal test');
    return { rows: out, splitFlags: splitFlags, stats: stats, warnings: warnings, refDate: refDate };
  }

  // ============================== GHI FILE TEMPLATE ==============================
  var TEMPLATE_COLS = ['ID', 'Tên thương vụ', 'Loại deal', 'Tên khách hàng', 'Phân nhóm khách hàng', 'Sale', 'Reseller',
    'Giai đoạn', 'Trạng thái', 'Giá trị thương vụ', 'Hãng', 'Nguồn hãng', 'Giá trị BOM theo hãng', BOM_TXT,
    'Phân loại HĐ', 'Phân loại deal size', 'Close date', 'Số ngày quá hạn forecast', 'Lý do Failed', 'Ngày tạo',
    'Cập nhật lần cuối', 'Cờ chất lượng'];
  var DATE_COLS = ['Close date', 'Ngày tạo', 'Cập nhật lần cuối'];

  function toSerial(t) { return t / 86400000 + 25569; }
  function fromSerial(s) { return Math.round((s - 25569) * 86400000 / 60000) * 60000; }

  function colLetter(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  function cellOf(v, z, style) {
    var c;
    if (v === null || v === undefined || (typeof v === 'number' && isNaN(v))) c = { t: 'z', v: undefined };
    else if (typeof v === 'number') c = { t: 'n', v: v };
    else if (typeof v === 'boolean') c = { t: 'b', v: v };
    else c = { t: 's', v: String(v) };
    if (z) c.z = z;
    if (style) c.s = style;
    return c;
  }

  function buildTemplateWorkbook(XLSX, res) {
    var F = 'Arial';
    var thin = { style: 'thin', color: { rgb: 'D9D9D9' } };
    var bd = { left: thin, right: thin, top: thin, bottom: thin };
    var widths = { 'ID': 10, 'Tên thương vụ': 42, 'Loại deal': 15, 'Tên khách hàng': 38, 'Phân nhóm khách hàng': 26, 'Sale': 20,
      'Reseller': 18, 'Giai đoạn': 20, 'Trạng thái': 11, 'Giá trị thương vụ': 18, 'Hãng': 15, 'Nguồn hãng': 20,
      'Giá trị BOM theo hãng': 18, 'Thông tin BOM (model, số lượng, năm support...)': 50, 'Phân loại HĐ': 13,
      'Phân loại deal size': 12, 'Close date': 12, 'Số ngày quá hạn forecast': 12, 'Lý do Failed': 40,
      'Ngày tạo': 16, 'Cập nhật lần cuối': 16, 'Cờ chất lượng': 90 };
    var newCols = ['Loại deal', 'Nguồn hãng', 'Giá trị BOM theo hãng', 'Số ngày quá hạn forecast'];
    var fmt = { 'Giá trị thương vụ': '#,##0', 'Giá trị BOM theo hãng': '#,##0', 'Close date': 'dd/mm/yyyy',
      'Ngày tạo': 'dd/mm/yyyy hh:mm', 'Cập nhật lần cuối': 'dd/mm/yyyy hh:mm', 'ID': '0', 'Số ngày quá hạn forecast': '0' };
    var cols = TEMPLATE_COLS, ws = {};
    cols.forEach(function (h, j) {
      ws[colLetter(j) + '1'] = cellOf(h, null, {
        font: { name: F, bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
        fill: { patternType: 'solid', fgColor: { rgb: newCols.indexOf(h) >= 0 ? '2E75B6' : '1F4E78' } }, border: bd,
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true } });
    });
    var base = { font: { name: F, sz: 10 }, border: bd };
    var grey = { font: { name: F, sz: 10 }, border: bd, fill: { patternType: 'solid', fgColor: { rgb: 'F2F2F2' } } };
    res.rows.forEach(function (rec, i) {
      var st = res.splitFlags[i] ? grey : base;
      cols.forEach(function (h, j) {
        var v = rec[h];
        if (DATE_COLS.indexOf(h) >= 0 && v !== null && v !== undefined) v = toSerial(v);
        ws[colLetter(j) + (i + 2)] = cellOf(v, fmt[h], st);
      });
    });
    var last = colLetter(cols.length - 1) + (res.rows.length + 1);
    ws['!ref'] = 'A1:' + last;
    ws['!cols'] = cols.map(function (h) { return { wch: widths[h] || 15 }; });
    ws['!rows'] = [{ hpt: 32 }];
    ws['!autofilter'] = { ref: 'A1:' + last };
    ws['!freeze'] = '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>';

    var s = res.stats;
    var g = {}, glines = [
      ['HƯỚNG DẪN ĐỌC FILE', null],
      ['Ngày chốt số liệu', fmtDate(res.refDate)],
      ['Nguồn dữ liệu', 'Sheet "' + s.sheet + '" của file export CRM. BOM chỉ lấy từ 9 cột "Bảng danh mục sản phẩm dự án - …"; ' +
        'không dùng cột text tổng hợp "Bảng danh mục sản phẩm dự án".'],
      ['Đối soát', 'Raw: ' + s.raw_rows + ' dòng = ' + s.deals_raw + ' deal, ' + s.lines_raw + ' dòng BOM. Loại ' + s.tests + ' deal test. ' +
        'Output: ' + s.deals + ' deal (' + s.co_hoi + ' Cơ hội, ' + s.lead + ' Lead chiến dịch), ' + s.rows + ' dòng; ' +
        'đã đọc ' + s.lines_used + ' dòng BOM.'],
      ['Quy tắc đếm', 'Một deal nhiều hãng được tách thành nhiều dòng cùng ID. Đếm số deal phải dùng ID duy nhất. ' +
        'Giá trị thương vụ chỉ ghi ở dòng hãng chính (dòng tách tô nền xám, giá trị để trống), nên cộng ' +
        'cột Giá trị thương vụ không bị đếm trùng.'],
      ['Loại deal', 'Lead chiến dịch = deal Active, giai đoạn "Đăng kí cơ hội", không có dòng BOM, giá trị = 0, không có Close date ' +
        'và không có hoạt động (email, họp, gọi, ghi chú, ghi nhanh = 0). Còn lại là Cơ hội. Mặc định phân tích chỉ tính ' +
        'Cơ hội. Deal Lead tự chuyển sang Cơ hội khi sale bổ sung một trong các thông tin trên hoặc chuyển giai đoạn.'],
      ['Giá trị thương vụ', 'Đơn vị VND. Lấy giá trị CRM; nếu CRM = 0 thì lấy tổng Thành tiền BOM. Deal gắn USD có giá trị ' +
        '>= ' + commaInt(USD_MISLABEL_THRESHOLD) + ' coi là VND nhập nhầm nhãn. Để trống = chưa có dữ liệu (không phải 0).'],
      ['Hãng', 'Lấy từ cột "Bảng danh mục sản phẩm dự án - Hãng". Nếu trống thì suy lần lượt từ: Tên giải pháp/Ghi chú của dòng ' +
        'đó (dòng dịch vụ gán "Dịch vụ") → cột "Tên hãng" → từ khóa trong Thông tin BOM / Tên thương vụ / Tên dự án. ' +
        'Dòng đầu của mỗi deal là hãng chính (hãng có giá trị BOM lớn nhất; "Khác", "Dịch vụ" chỉ là hãng chính khi ' +
        'deal không có hãng nào khác). Để trống = chưa xác định được.'],
      ['Nguồn hãng', 'BOM = lấy trực tiếp từ cột Hãng của BOM · Tên giải pháp/Ghi chú = suy từ dòng BOM · Tên hãng (CRM) = cột ' +
        '"Tên hãng" của deal · Từ khóa = tìm trong Thông tin BOM / tên thương vụ / tên dự án · Không xác định.'],
      ['Giá trị BOM theo hãng', 'Tổng Thành tiền các dòng BOM của hãng đó (dòng thiếu Thành tiền nhưng có Số lượng và Đơn giá ' +
        'thì tính = SL x ĐG; dòng lỗi #VALUE! coi là chưa có giá). Dùng để xem pipeline theo hãng.'],
      ['Phân loại deal size', '< 1B: dưới 1 tỷ · 1B - 2B: 1 đến dưới 2 tỷ · 2B - 5B: 2 đến dưới 5 tỷ · > 5B: từ 5 tỷ. ' +
        'Tính lại từ Giá trị thương vụ, không dùng phân loại CRM vì CRM ghi không nhất quán.'],
      ['Sale', 'Đổi tài khoản CRM (cột "Chủ sở hữu thương vụ") sang họ tên theo bảng ở sheet Danh_muc. ' +
        'Tài khoản chưa có trong bảng thì giữ nguyên và gắn cờ. Sales user ở dòng BOM khác chủ deal được gắn cờ.'],
      ['Phân nhóm khách hàng', 'Phân tự động theo tên KH với 12 khối ở sheet Danh_muc. Cơ quan nhà nước phân theo cấp ' +
        '(Đảng/NN, Bộ ngành TW, địa phương); mọi đơn vị thuộc Bộ Công an, Bộ Quốc phòng vào An ninh - ' +
        'Quốc phòng; doanh nghiệp phân theo ngành kinh doanh chính, công ty con theo ngành của tập đoàn. ' +
        'KH chưa phân được hoặc chưa chắc chắn có cờ cần kiểm tra.'],
      ['Giai đoạn / Trạng thái', 'Giữ nguyên giá trị CRM ("Giai đoạn", "Trạng thái thương vụ"); mâu thuẫn giữa hai cột được ' +
        'ghi ở Cờ chất lượng.'],
      ['Close date', 'Lấy từ cột "Timeline dự án" của CRM = ngày dự kiến Won (forecast).'],
      ['Số ngày quá hạn forecast', 'Ngày chốt số liệu − Close date; chỉ tính cho deal Active có Close date đã qua.'],
      ['Lý do Failed', 'Gộp "Lý do thất bại" và "Failed additional comment": <lý do> – <ghi chú>.'],
      ['Cờ chất lượng', '[VẤN ĐỀ] = dữ liệu thiếu/mâu thuẫn cần người phụ trách kiểm tra (luôn ghi lỗi gốc kể cả khi đã suy ' +
        'thông tin). [ĐÃ XỬ LÝ] = thay đổi so với CRM (tên KH, hãng, giá trị, deal size). Deal Lead chỉ có cờ ' +
        '"Lead chiến dịch" và các cờ cần kiểm tra tên KH / sale / phân nhóm.'],
      ['Tham số', 'Ngưỡng USD nhập nhầm: ' + commaInt(USD_MISLABEL_THRESHOLD) + '; tỷ giá giả định: ' + commaInt(USD_RATE) + '; ' +
        'ngưỡng lệch BOM: ' + Math.round(BOM_TOLERANCE * 100) + '%.']
    ];
    glines.forEach(function (ab, i) {
      g['A' + (i + 1)] = cellOf(ab[0], null, { font: { name: F, sz: i === 0 ? 12 : 10, bold: true }, alignment: { vertical: 'top' } });
      g['B' + (i + 1)] = cellOf(ab[1], null, { font: { name: F, sz: 10 }, alignment: { wrapText: true, vertical: 'top' } });
    });
    g['!ref'] = 'A1:B' + glines.length;
    g['!cols'] = [{ wch: 26 }, { wch: 110 }];

    var dm = {}, hdr = { font: { name: F, bold: true, color: { rgb: 'FFFFFF' }, sz: 10 }, fill: { patternType: 'solid', fgColor: { rgb: '1F4E78' } } };
    var plain = { font: { name: F, sz: 10 } };
    ['Tài khoản CRM', 'Tên Sale', '', 'Phân nhóm khách hàng', 'Số deal Cơ hội', 'Ví dụ khách hàng'].forEach(function (h, j) {
      if (h) dm[colLetter(j) + '1'] = cellOf(h, null, hdr);
    });
    var users = Object.keys(SALE_NAMES);
    users.forEach(function (u, i) {
      dm['A' + (i + 2)] = cellOf(u, null, plain); dm['B' + (i + 2)] = cellOf(SALE_NAMES[u], null, plain);
    });
    var seen = new Set(), co = [];
    res.rows.forEach(function (o) { if (!seen.has(o.ID)) { seen.add(o.ID); if (o['Loại deal'] === 'Cơ hội') co.push(o); } });
    SEGMENTS.forEach(function (seg, i) {
      var sub = co.filter(function (o) { return o['Phân nhóm khách hàng'] === seg; });
      var vc = new Map();
      sub.forEach(function (o) { vc.set(o['Tên khách hàng'], (vc.get(o['Tên khách hàng']) || 0) + 1); });
      var ex = Array.from(vc.keys()).sort(function (a, b) { return vc.get(b) - vc.get(a); }).slice(0, 4).join(', ');
      dm['D' + (i + 2)] = cellOf(seg, null, plain); dm['E' + (i + 2)] = cellOf(sub.length, null, plain); dm['F' + (i + 2)] = cellOf(ex, null, plain);
    });
    dm['!ref'] = 'A1:F' + (Math.max(users.length, SEGMENTS.length) + 1);
    dm['!cols'] = [16, 24, 3, 32, 12, 90].map(function (w) { return { wch: w }; });

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.utils.book_append_sheet(wb, g, 'Huong_dan');
    XLSX.utils.book_append_sheet(wb, dm, 'Danh_muc');
    return wb;
  }

  // ============================== ĐỌC LẠI FILE TEMPLATE ĐÃ CONVERT ==============================
  function isTemplateWorkbook(wb) {
    if (wb.SheetNames.indexOf('Data') < 0) return false;
    var hdr = (XLSXref().utils.sheet_to_json(wb.Sheets['Data'], { header: 1, range: 0, defval: null })[0] || []).map(String);
    return ['ID', 'Tên thương vụ', 'Loại deal', 'Hãng', 'Cờ chất lượng'].every(function (c) { return hdr.indexOf(c) >= 0; });
  }
  var _X = null;
  function XLSXref() { return _X || root.XLSX; }

  function loadTemplate(XLSX, wb, refDate) {
    _X = XLSX;
    var sr = sheetRows(XLSX, wb.Sheets['Data']);
    var rows = sr.rows.filter(function (r) { return notna(r['ID']); }).map(function (r) {
      var o = {};
      TEMPLATE_COLS.forEach(function (c) {
        var v = r[c] === undefined ? null : r[c];
        if (DATE_COLS.indexOf(c) >= 0) {
          if (typeof v === 'number') v = fromSerial(v);
          else if (v instanceof Date || isStr(v)) v = parseDate(v);
          else v = null;
        } else if (isStr(v)) v = v === '' ? null : v;
        o[c] = v;
      });
      o.ID = Math.trunc(Number(o.ID));
      return o;
    });
    var seen = new Set(), flags = rows.map(function (o) { var f = seen.has(o.ID); seen.add(o.ID); return f; });
    var first = rows.filter(function (o, i) { return !flags[i]; });
    // Ngày chốt: đọc từ sheet Huong_dan nếu có
    var ref = refDate;
    if (wb.Sheets['Huong_dan']) {
      var gd = sheetRows(XLSX, wb.Sheets['Huong_dan']);
      var all = [gd.header].concat(gd.rows.map(function (r) { return gd.header.map(function (h) { return r[h]; }); }));
      all.forEach(function (a) { if (a[0] === 'Ngày chốt số liệu') { var p = a[1] instanceof Date ? parseDate(a[1]) : parseDate(String(a[1])); if (p !== null) ref = p; } });
    }
    return {
      rows: rows, splitFlags: flags, warnings: [], refDate: ref, fromTemplate: true,
      stats: { sheet: 'Data', deals: first.length, rows: rows.length,
        lead: first.filter(function (o) { return o['Loại deal'] === 'Lead chiến dịch'; }).length,
        co_hoi: first.filter(function (o) { return o['Loại deal'] === 'Cơ hội'; }).length }
    };
  }

  function refFromName(name, today) {
    var m = /_(\d{2})_(\d{2})_(\d{2})\.xlsx$/i.exec(name) || /\.\d{2}\.\d{2}\.(\d{2})\.(\d{2})\.(\d{2})\.xlsx$/i.exec(name);
    if (m) {
      var t = parseDate(m[1] + '/' + m[2] + '/20' + m[3]);
      if (t !== null) return t;
    }
    return today;
  }

  var API = {
    convert: convert, buildTemplateWorkbook: buildTemplateWorkbook, loadTemplate: loadTemplate,
    isTemplateWorkbook: function (XLSX, wb) { _X = XLSX; return isTemplateWorkbook(wb); },
    findSourceSheet: findSourceSheet, refFromName: refFromName, parseDate: parseDate, fmtDate: fmtDate, fold: fold,
    toSerial: toSerial, colLetter: colLetter, cellOf: cellOf,
    TEMPLATE_COLS: TEMPLATE_COLS, SEGMENTS: SEGMENTS, SALE_NAMES: SALE_NAMES, DATE_COLS: DATE_COLS, BOM_TXT: BOM_TXT
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.CRMConvert = API;
})(typeof window !== 'undefined' ? window : this);
