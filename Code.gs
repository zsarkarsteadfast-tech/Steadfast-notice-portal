/**
 * Steadfast Update Hub - Google Apps Script backend
 * Reads your existing tabs (SFC Notice, QA Follow-up, IB Product Knowledge, Police Station)
 * and adds 4 helper tabs when you run setup(): Users, Reports, Ticker, Links.
 *
 * 1) Paste this whole file into Extensions > Apps Script (in the Google Sheet)
 * 2) Run setup() once, then Deploy > New deployment > Web app
 *    (Execute as: Me, Who has access: Anyone). The site uses password login and server-validated sessions; FIRST_ADMIN below is auto-approved as admin the first time setup() runs.
 */
var FIRST_ADMIN = 'zsarkar.steadfast@gmail.com';   // always seeded as admin by setup()
var SHEET_ID  = '1X-QZIGes4h2S5LyW0zPmGNRUF7M5su64B9MXqWHdlCQ';
// Any valid email may request access; an admin must approve it.
var TZ        = 'Asia/Dhaka';
var NOTICE_TABS = [{tab: 'SFC Notice', group: 'cs'}, {tab: 'QA Follow-up', group: 'qa'}];   // cs = CS Official Notice, qa = QA Follow-up
var FAQ_TAB = 'IB Product Knowledge';
var POLICE_TAB = 'Police Station';

/* ------------------------------ entry points ------------------------------ */
function doGet() { return out_({ok: true, name: 'Steadfast Update Hub API'}); }

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    return out_(handle_(req));
  } catch (err) {
    return out_({error: String(err.message || err), code: err.code || 'error'});
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(code, msg) { var e = new Error(msg || code); e.code = code; throw e; }

function handle_(req) {
  if (req.action === 'login') return lock_(function () { return login_(req); });
  if (req.action === 'forgotPassword') return forgotPassword_(req);
  var user = sessionUser_(req.token);
  if (req.action === 'logout') { removeSession_(req.token); return {ok: true}; }
  if (req.action === 'changePassword') return lock_(function () { return changePassword_(req, user); });

  switch (req.action) {
    case 'load': return load_(user);
    case 'report': return lock_(function () { return report_(req, user); });
  }
  if (user.role === 'viewer') fail_('forbidden', 'Editors only');
  switch (req.action) {
    case 'setExpired': return lock_(function () { return setExpired_(req, user); });
    case 'dismiss': return lock_(function () { delReports_(Number(req.id)); return {ok: true}; });
    case 'saveTicker': return lock_(function () { return saveTicker_(req.messages); });
    case 'saveFaq': return lock_(function () { return saveFaq_(req, user); });
  }
  if (user.role !== 'admin') fail_('forbidden', 'Admins only');
  switch (req.action) {
    case 'saveNotice': return lock_(function () { return saveNotice_(req, user); });
    case 'createNotice': return lock_(function () { req.id = 0; return saveNotice_(req, user); });
    case 'deleteNotice': return lock_(function () { return deleteNotice_(req, user); });
    case 'resetPassword': return lock_(function () { return resetPassword_(req); });
    case 'saveUser': return lock_(function () { return saveUser_(req, user); });
    case 'deleteUser': return lock_(function () { return deleteUser_(req, user); });
  }
  fail_('bad', 'Unknown action');
}

function lock_(fn) {
  var l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ------------------------------ sheet helpers ------------------------------ */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function tab_(name) { var s = ss_().getSheetByName(name); if (!s) fail_('server', 'Missing tab: ' + name); return s; }

function readTab_(sh) {
  var vals = sh.getDataRange().getValues();
  var cols = {}, rows = [];
  if (!vals.length) return {sh: sh, cols: cols, rows: rows};
  vals[0].forEach(function (h, i) {
    var k = String(h).trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    if (k && cols[k] === undefined) cols[k] = i;
  });
  for (var r = 1; r < vals.length; r++) {
    var o = {_row: r + 1};
    for (var k in cols) o[k] = vals[r][cols[k]];
    rows.push(o);
  }
  return {sh: sh, cols: cols, rows: rows};
}

function s_(v) { return v === null || v === undefined ? '' : String(v); }
function nz_(v) { var t = s_(v).trim(); return /^(_|-|n\/a|na|null)$/i.test(t) ? '' : t; }
function yes_(v) { return v === true || /^(yes|y|true|1|expired|urgent)$/i.test(s_(v).trim()); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm"); }
function h32_(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return h >>> 0; }

var MONTHS_ = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
function parseDate_(v, y, m) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var t = s_(v).trim(), a;
  if ((a = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/))) {          // M/D/YYYY (as in your sheet)
    var yr = +a[3]; if (yr < 100) yr += 2000;
    if (+a[1] >= 1 && +a[1] <= 12 && +a[2] >= 1 && +a[2] <= 31) return yr + '-' + pad_(+a[1]) + '-' + pad_(+a[2]);
  }
  if ((a = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) return a[0];
  var mi = MONTHS_.indexOf(s_(m).trim().toLowerCase().slice(0, 3));
  if (Number(y) && mi >= 0) return Number(y) + '-' + pad_(mi + 1) + '-01';
  return '';
}

/* ------------------------------ notices ------------------------------ */
var GENERIC_ = /^(notice|নোটিশ|ঘোষণা|ঘোষণা|বিজ্ঞপ্তি|announcement|official notice|অফিসিয়াল নোটিশ|জরুরি ঘোষণা|জরুরি নোটিশ|গুরুত্বপূর্ণ বিজ্ঞপ্তি|জরুরি আপডেট|hello team)\s*[!:।]?$/i;
function lineClean_(l) {
  return l.replace(/[\u200b\u200c\u200d\ufeff]/g, '').trim()
    .replace(/^[^\p{L}\p{M}\p{N}“"(\[]+/u, '')
    .replace(/[^\p{L}\p{M}\p{N}”")\]?!.।]+$/u, '').trim();
}
function cut_(s, n) {
  if (s.length <= n) return s;
  var c = s.slice(0, n), i = c.lastIndexOf(' ');
  return (i > n * 0.5 ? c.slice(0, i) : c) + '…';
}
function splitNotice_(text) {
  var lines = s_(text).replace(/\r/g, '').split('\n');
  var urgent = /জরুরি|urgent/i.test(lines.slice(0, 4).join(' '));
  for (var i = 0; i < lines.length; i++) {
    var c = lineClean_(lines[i]);
    if (!c || GENERIC_.test(c)) continue;
    c = c.replace(/^(বিষয়|subject|নোটিশ|notice)\s*[:：]\s*/i, '').trim();
    var rest = lines.slice(i + 1).join('\n').trim();
    if (!rest) return {title: cut_(c, 80), body: c.length > 80 ? c : '', urgent: urgent};
    return {title: cut_(c, 120), body: rest, urgent: urgent};
  }
  return {title: cut_(lineClean_(s_(text)) || s_(text).trim(), 80), body: '', urgent: urgent};
}

function noticeId_(tabName, r, seen) {
  if (Number(r.notice_id) > 0) return Number(r.notice_id);
  var key = tabName + '|' + r.date_key + '|' + s_(r.notice).trim().slice(0, 80);
  seen[key] = (seen[key] || 0) + 1;
  return h32_(key + '|' + seen[key]);
}

function noticesOf_(nt) {
  var sh = ss_().getSheetByName(nt.tab);
  if (!sh) return [];
  var d = readTab_(sh), seen = {}, out = [], today = today_();
  d.rows.forEach(function (r) {
    var txt = s_(r.notice).trim();
    if (!txt) return;
    var date = parseDate_(r.date, r.year, r.month);
    r.date_key = date;
    var id = noticeId_(nt.tab, r, seen);
    var sp = splitNotice_(txt);
    var title = nz_(r.title) || sp.title;
    var exp = yes_(r.expired) || /^expired$/i.test(s_(r.status).trim());
    if (!exp && r.expires) { var ex = parseDate_(r.expires, '', ''); if (ex && ex < today) exp = true; }
    var urgent = d.cols.urgent !== undefined && s_(r.urgent).trim() !== '' ? yes_(r.urgent) : sp.urgent;
    out.push({
      id: id, tab: nt.tab, src: nt.group, date: date,
      cat: s_(r.type).replace(/\s+/g, ' ').trim() || 'General',
      title: title, body: d.cols.body !== undefined && s_(r.body) ? s_(r.body) : sp.body, expires: r.expires ? parseDate_(r.expires, "", "") : "", revision: digest_(JSON.stringify(r)), urgent: urgent ? 1 : 0, expired: exp ? 1 : 0,
      ann: nz_(r.announced_by), appr: nz_(r.approved_by),
      upd: {by: nz_(r.sheet_update_by), at: ''}
    });
  });
  return out;
}
function notices_() {
  var all = [];
  NOTICE_TABS.forEach(function (nt) { all = all.concat(noticesOf_(nt)); });
  return all;
}

/* ------------------------------ FAQ (IB Product Knowledge) ------------------------------ */
var NUMLINE_ = /^[^\p{L}\p{N}]*\d+\s*[.)]\s*/u;
function lastQuestion_(s) {
  var lines = s_(s).replace(/\r/g, '').split('\n'), idx = -1;
  for (var i = 0; i < lines.length; i++) if (NUMLINE_.test(lines[i].trim())) idx = i;
  var q = idx >= 0 ? lines.slice(idx).join(' ') : lines.filter(function (x) { return x.trim(); }).slice(-1).join(' ');
  return q.replace(NUMLINE_, '').replace(/\s+/g, ' ').replace(/[^\p{L}\p{M}\p{N}?)\]”"।.]+$/u, '').trim();
}
function trimHeaders_(a) {
  var lines = a.split('\n');
  while (lines.length > 1) {
    var l = lines[lines.length - 1].trim();
    if (l === '' || (/^[^\p{L}\p{N}]/u.test(l) && l.length < 70)) lines.pop(); else break;
  }
  return lines.join('\n').trim();
}
function parseQA_(text) {
  var t = s_(text).replace(/\r/g, '').trim();
  var parts = t.split(/Answer\s*[:：]/i);
  if (parts.length < 2) return [];
  var out = [], q = lastQuestion_(parts[0]);
  for (var k = 1; k < parts.length; k++) {
    var seg = parts[k], a = seg, next = '';
    if (k < parts.length - 1) {
      var m = seg.search(/\n[ \t]*[^\p{L}\p{N}\n]*\d+\s*[.)]\s/u);
      if (m >= 0) { a = seg.slice(0, m); next = seg.slice(m + 1); }
    }
    a = trimHeaders_(a.trim());
    if (q && a) out.push({q: q, a: a});
    q = lastQuestion_(next);
  }
  return out;
}
function faqs_() {
  var sh = ss_().getSheetByName(FAQ_TAB);
  if (!sh) return [];
  var d = readTab_(sh), seen = {}, out = [];
  d.rows.forEach(function (r) {
    var cat = s_(r.type).replace(/\s+/g, ' ').trim() || 'General';
    var date = parseDate_(r.date, r.year, r.month);
    parseQA_(r.service).forEach(function (p) {
      var key = 'faq|' + cat + '|' + p.q.slice(0, 80);
      seen[key] = (seen[key] || 0) + 1;
      out.push({id: h32_(key + '|' + seen[key]), cat: cat, q: p.q, a: p.a, upd: {by: nz_(r.added_by), at: ''}, date: date});
    });
  });
  // Keep edits in a separate helper tab so the original knowledge sheet is never overwritten.
  var edits = ss_().getSheetByName('FAQ Edits');
  if (edits) {
    var ed = readTab_(edits), byId = {};
    ed.rows.forEach(function (r) { if (s_(r.faq_id).trim()) byId[s_(r.faq_id).trim()] = r; });
    out.forEach(function (f) {
      var r = byId[String(f.id)];
      if (r) { f.q = s_(r.q) || f.q; f.a = s_(r.a); f.upd = {by: s_(r.updated_by), at: s_(r.updated_at)}; }
    });
  }
  return out;
}

function saveFaq_(req, user) {
  var f = req.faq || {}, id = s_(f.id).trim();
  var q = s_(f.q).trim(), a = s_(f.a).trim(), cat = s_(f.cat).trim();
  if (!/^\d+$/.test(id) || !q || !a || q.length > 5000 || a.length > 20000 || cat.length > 150)
    fail_('bad', 'Invalid FAQ content');
  var existing = faqs_().filter(function (x) { return String(x.id) === id; })[0];
  if (!existing) fail_('conflict', 'FAQ no longer exists');
  // Prevent changing an FAQ into a different record by tampering with its identity.
  if (cat && cat !== existing.cat) fail_('conflict', 'FAQ category changed; reload and try again');
  var sh = ensureTab_(ss_(), 'FAQ Edits', ['faq_id', 'cat', 'q', 'a', 'updated_by', 'updated_at']);
  var d = readTab_(sh), hit = null;
  d.rows.forEach(function (r) { if (s_(r.faq_id).trim() === id) hit = r; });
  var values = [Number(id), existing.cat, q, a, user.name || user.email, now_()];
  if (hit) sh.getRange(hit._row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  audit_(user, 'EDIT_FAQ', 'faq', id, q, 'FAQ question and answer updated', FAQ_TAB);
  return load_(user);
}

/* ------------------------------ police station lookup ------------------------------ */
function police_() {
  var sh = ss_().getSheetByName(POLICE_TAB);
  if (!sh) return [];
  return readTab_(sh).rows.map(function (r) {
    return {hub: s_(r.hub).replace(/\s+/g, ' ').trim(), ps: s_(r.police_station).replace(/\s+/g, ' ').trim(), area: s_(r.area).replace(/\s+/g, ' ').trim()};
  }).filter(function (x) { return x.hub || x.ps; });
}

/* ------------------------------ users, links, reports, ticker ------------------------------ */
function ids_(v) {
  var t = s_(v).trim().toLowerCase();
  if (t === 'all' || t === '*') return 'all';
  return t.split(/[,\s;]+/).filter(Boolean).map(Number).filter(function (n) { return isFinite(n); });
}
function userOf_(r) {
  var role = s_(r.role).trim().toLowerCase();
  if (role !== 'admin' && role !== 'editor') role = 'viewer';
  var st = s_(r.status).trim().toLowerCase();
  if (st !== 'pending' && st !== 'blocked') st = 'active';
  return {
    email: s_(r.email).trim().toLowerCase(), name: s_(r.name).trim(), role: role, status: st,
    perms: {n: yes_(r.notice) ? 1 : 0, f: yes_(r.faq) ? 1 : 0, s: yes_(r.sheets) ? 1 : 0, ids: ids_(r.sheet_ids)},
    req: s_(r.requested_at)
  };
}
function users_() { return readTab_(tab_('Users')).rows.map(userOf_).filter(function (u) { return u.email; }); }
function findUser_(email) {
  var all = users_();
  for (var i = 0; i < all.length; i++) if (all[i].email === email) {
    if (!all[i].name) all[i].name = email.split('@')[0].split(/[._-]+/).map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
    return all[i];
  }
  return null;
}
function requestAccess_(email) {
  var d = readTab_(tab_('Users'));
  for (var i = 0; i < d.rows.length; i++) if (s_(d.rows[i].email).trim().toLowerCase() === email) return;
  d.sh.appendRow([email, 'viewer', '', 'pending', 'no', 'no', 'no', '', now_()]);
}
function links_() {
  return readTab_(tab_('Links')).rows.filter(function (r) { return s_(r.title).trim(); }).map(function (r) {
    var type = /police/i.test(s_(r.type)) ? 'police' : 'link', url = s_(r.url).trim();
    if (type === 'link' && !/^https?:\/\//i.test(url)) url = '';
    return {id: Number(r.id) || h32_(s_(r.title)), cat: s_(r.category).trim() || 'IB', title: s_(r.title).trim(), desc: s_(r.description).trim(), url: url, type: type, upd: {by: '', at: ''}};
  });
}
function reports_(user) {
  var rows = readTab_(tab_('Reports')).rows.map(function (r) {
    return {id: Number(r.id), nid: Number(r.notice_id), by: s_(r.reported_by).toLowerCase(), note: s_(r.note), date: parseDate_(r.date, '', '')};
  });
  return user.role === 'viewer' ? rows.filter(function (r) { return r.by === user.email; }) : rows;
}
function ticker_() {
  return readTab_(tab_('Ticker')).rows.map(function (r) { return s_(r.message).trim(); }).filter(Boolean);
}

function audit_(user, action, targetType, targetId, title, details, tabName) {
  var sh = ensureTab_(ss_(), 'Edit History', ['event_at', 'actor_email', 'actor_name', 'action', 'target_type', 'target_id', 'title', 'details', 'tab']);
  sh.appendRow([now_(), s_(user.email), s_(user.name), s_(action), s_(targetType), s_(targetId), s_(title).slice(0, 1000), s_(details).slice(0, 5000), s_(tabName)]);
}
function history_() {
  var sh = ss_().getSheetByName('Edit History');
  if (!sh) return [];
  return readTab_(sh).rows.slice(-500).reverse().map(function (r) {
    return {at: s_(r.event_at), email: s_(r.actor_email), name: s_(r.actor_name), action: s_(r.action), target_type: s_(r.target_type), id: s_(r.target_id), title: s_(r.title), details: s_(r.details), tab: s_(r.tab)};
  });
}

function load_(user) {
  var can = function (k) { return user.role !== 'viewer' || user.perms[k]; };
  var d = {user: user, notices: [], faqs: [], sheets: [], police: null, reports: [], ticker: ticker_(), users: null, history: []};
  if (can('n')) d.notices = notices_();
  if (can('f')) d.faqs = faqs_();
  if (can('s')) {
    d.sheets = links_().filter(function (l) {
      return user.role !== 'viewer' || user.perms.ids === 'all' || user.perms.ids.indexOf(l.id) > -1;
    });
    if (d.sheets.some(function (l) { return l.type === 'police'; })) d.police = police_();
  }
  d.reports = reports_(user);
  if (user.role === 'admin') {
    var props = PropertiesService.getScriptProperties();
    d.users = users_().map(function (u) { u.resetRequested = props.getProperty('reset:' + u.email) || ''; return u; });
    d.history = history_();
  }
  return d;
}

/* ------------------------------ writes ------------------------------ */
function delReports_(nid) {
  var d = readTab_(tab_('Reports')), c = d.cols.notice_id;
  if (c === undefined) return;
  for (var i = d.rows.length - 1; i >= 0; i--) if (Number(d.rows[i].notice_id) === nid) d.sh.deleteRow(d.rows[i]._row);
}
function report_(req, user) {
  var nid = Number(req.id);
  var exists = notices_().some(function (n) { return n.id === nid && !n.expired; });
  if (!exists) fail_('bad', 'Notice not found');
  var d = readTab_(tab_('Reports'));
  for (var i = 0; i < d.rows.length; i++)
    if (Number(d.rows[i].notice_id) === nid && s_(d.rows[i].reported_by).toLowerCase() === user.email) return {ok: true, dup: true};
  d.sh.appendRow([Date.now(), nid, user.email, s_(req.note).slice(0, 500), today_()]);
  return {ok: true};
}
function setExpired_(req, user) {
  var nt = NOTICE_TABS.filter(function (x) { return x.tab === req.tab; })[0];
  if (!nt) fail_('bad', 'Unknown tab');
  var sh = tab_(nt.tab), d = readTab_(sh), seen = {}, hit = null, id = Number(req.id);
  d.rows.forEach(function (r) {
    if (!s_(r.notice).trim()) return;
    r.date_key = parseDate_(r.date, r.year, r.month);
    if (noticeId_(nt.tab, r, seen) === id) hit = r;
  });
  if (!hit) fail_('bad', 'Notice not found');
  var col = d.cols.expired;
  if (col === undefined) { col = sh.getLastColumn(); sh.getRange(1, col + 1).setValue('Expired'); }
  sh.getRange(hit._row, col + 1).setValue(req.expired ? 'Yes' : '');
  if (d.cols.sheet_update_by !== undefined) sh.getRange(hit._row, d.cols.sheet_update_by + 1).setValue(user.name);
  if (req.expired) {
    audit_(user, 'MARK_EXPIRED', 'notice', id, s_(hit.title) || splitNotice_(hit.notice).title, 'Notice marked as expired', nt.tab);
    delReports_(id);
  } else {
    audit_(user, 'RESTORE_NOTICE', 'notice', id, s_(hit.title) || splitNotice_(hit.notice).title, 'Notice restored to active status', nt.tab);
  }
  return load_(user);
}
function saveTicker_(msgs) {
  var sh = tab_('Ticker'), list = (msgs || []).map(function (m) { return s_(m).trim().slice(0, 300); }).filter(Boolean).slice(0, 20);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 1).clearContent();
  if (list.length) sh.getRange(2, 1, list.length, 1).setNumberFormat('@').setValues(list.map(function (m) { return [m]; }));
  return {ok: true};
}
function saveUser_(req, me) {
  var email = s_(req.email).trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail_('bad', 'Bad email');
  var role = ['viewer', 'editor', 'admin'].indexOf(req.role) > -1 ? req.role : 'viewer';
  var st = ['active', 'blocked'].indexOf(req.status) > -1 ? req.status : 'active';
  if (email === me.email && (role !== 'admin' || st !== 'active')) fail_('self', 'You cannot block or demote yourself');
  var p = req.perms || {}, ids = p.ids === 'all' ? 'all' : (p.ids || []).map(Number).filter(isFinite).join(',');
  var yn = function (b) { return b ? 'yes' : 'no'; };
  var d = readTab_(tab_('Users')), row = null;
  d.rows.forEach(function (r) { if (s_(r.email).trim().toLowerCase() === email) row = r; });
  var keep = row ? s_(row.name) : '';
  var vals = [email, role, keep, st, yn(p.n), yn(p.f), yn(p.s), ids, ''];
  if (row) { vals[8] = s_(row.requested_at); d.sh.getRange(row._row, 1, 1, vals.length).setNumberFormat('@').setValues([vals]); }
  else d.sh.appendRow(vals);
  return {ok: true};
}
function deleteUser_(req, me) {
  var email = s_(req.email).trim().toLowerCase();
  if (email === me.email) fail_('self', 'You cannot remove yourself');
  var d = readTab_(tab_('Users'));
  for (var i = d.rows.length - 1; i >= 0; i--) if (s_(d.rows[i].email).trim().toLowerCase() === email) d.sh.deleteRow(d.rows[i]._row);
  return {ok: true};
}

/* ------------------------------ one-time setup (run from the editor) ------------------------------ */
function ensureTab_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function setup() {
  var ss = ss_();
  var users = ensureTab_(ss, 'Users', ['email', 'role', 'name', 'status', 'notice', 'faq', 'sheets', 'sheet_ids', 'requested_at']);
  ensureTab_(ss, 'Reports', ['id', 'notice_id', 'reported_by', 'note', 'date']);
  ensureTab_(ss, 'FAQ Edits', ['faq_id', 'cat', 'q', 'a', 'updated_by', 'updated_at']);
  ensureTab_(ss, 'Edit History', ['event_at', 'actor_email', 'actor_name', 'action', 'target_type', 'target_id', 'title', 'details', 'tab']);
  var tk = ensureTab_(ss, 'Ticker', ['message']);
  var links = ensureTab_(ss, 'Links', ['id', 'category', 'title', 'description', 'url', 'type']);
  var hasAdmin = users.getLastRow() > 1 && readTab_(users).rows.some(function (r) { return s_(r.email).trim().toLowerCase() === FIRST_ADMIN; });
  if (!hasAdmin) users.appendRow([FIRST_ADMIN, 'admin', '', 'active', 'yes', 'yes', 'yes', 'all', '']);
  if (tk.getLastRow() < 2) tk.getRange(2, 1, 2, 1).setValues([['📢 প্রতিদিন কাজ শুরুর আগে নতুন নোটিশ দেখে নিন'], ['🚩 পুরনো বা ভুল নোটিশ দেখলে ফ্ল্যাগ আইকন দিয়ে জানান']]);
  if (links.getLastRow() < 2) {
    var base = ss.getUrl().replace(/\/edit.*$/, '/edit'), gid = function (n) { var s = ss.getSheetByName(n); return s ? base + '#gid=' + s.getSheetId() : base; };
    links.getRange(2, 1, 3, 6).setValues([
      [1, 'IB', 'IB Product Knowledge', 'Questions and answers for IB agents', gid(FAQ_TAB), 'link'],
      [2, 'IB', 'Police Station Finder', 'Search an area to find its hub and police station', '', 'police'],
      [3, 'QA', 'QA Follow-up', 'QA notices and follow-ups', gid('QA Follow-up'), 'link']
    ]);
  }
  Logger.log('Setup done. Admin email: ' + users.getRange(2, 1).getValue());
}

var INITIAL_PASSWORD_ = 'sfcnotice123';
function authKey_(email) { return 'password:' + email; }
function digest_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function passwordHash_(password, salt) { var v = salt + ':' + password; for (var i = 0; i < 1000; i++) v = digest_(v); return v; }
function credentials_(email) { var value = PropertiesService.getScriptProperties().getProperty(authKey_(email)); return value ? JSON.parse(value) : {version: 'initial'}; }
function matchesPassword_(email, password) { var c = credentials_(email); return c.hash ? passwordHash_(String(password || ''), c.salt) === c.hash : String(password || '') === INITIAL_PASSWORD_; }
function setPassword_(email, password) { var salt = Utilities.getUuid(), version = Utilities.getUuid(); PropertiesService.getScriptProperties().setProperty(authKey_(email), JSON.stringify({salt: salt, hash: passwordHash_(password, salt), version: version})); return version; }
function issueSession_(email) { var token = Utilities.getUuid() + Utilities.getUuid(); CacheService.getScriptCache().put('session:' + digest_(token), JSON.stringify({email: email, version: credentials_(email).version}), 21600); return token; }
function sessionUser_(token) { var raw = CacheService.getScriptCache().get('session:' + digest_(String(token || ''))); if (!raw) fail_('auth', 'Please sign in again'); var s = JSON.parse(raw), user = findUser_(s.email); if (!user || user.status !== 'active' || s.version !== credentials_(s.email).version) fail_('auth', 'Please sign in again'); return user; }
function checkRate_(email) { var c = CacheService.getScriptCache(), key = 'attempt:' + digest_(email), n = Number(c.get(key) || 0); if (n >= 10) fail_('rate', 'Too many attempts. Try again in 15 minutes'); c.put(key, String(n + 1), 900); }
function login_(req) { var email = String(req.email || '').trim().toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) fail_('credentials', 'Enter a valid email and password'); checkRate_(email); if (!matchesPassword_(email, req.password)) fail_('credentials', 'Incorrect email or password'); var user = findUser_(email); if (!user) { requestAccess_(email); fail_('pending', 'Access requested. Wait for admin approval'); } if (user.status === 'pending') fail_('pending', 'Waiting for approval'); if (user.status !== 'active') fail_('denied', 'Access denied'); CacheService.getScriptCache().remove('attempt:' + digest_(email)); var d = load_(user); d.token = issueSession_(email); return d; }
function forgotPassword_(req) { var email = String(req.email || '').trim().toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) fail_('bad', 'Enter a valid email'); var cache = CacheService.getScriptCache(), key = 'reset-request:' + digest_(email); if (!cache.get(key)) { cache.put(key, '1', 900); if (findUser_(email)) PropertiesService.getScriptProperties().setProperty('reset:' + email, now_()); } return {ok: true}; }
function changePassword_(req, user) { checkRate_(user.email); if (!matchesPassword_(user.email, req.currentPassword)) fail_('credentials', 'Current password is incorrect'); var password = String(req.newPassword || ''); if (password.length < 8 || password.length > 128) fail_('bad', 'Use 8 to 128 characters'); setPassword_(user.email, password); PropertiesService.getScriptProperties().deleteProperty('reset:' + user.email); CacheService.getScriptCache().remove('attempt:' + digest_(user.email)); return {ok: true, token: issueSession_(user.email)}; }
function resetPassword_(req) { var email = String(req.email || '').trim().toLowerCase(); if (!findUser_(email)) fail_('bad', 'User not found'); setPassword_(email, INITIAL_PASSWORD_); PropertiesService.getScriptProperties().deleteProperty('reset:' + email); CacheService.getScriptCache().remove('attempt:' + digest_(email)); return {ok: true}; }

function removeSession_(token) { var key = 'session:' + digest_(String(token || '')); PropertiesService.getScriptProperties().deleteProperty(key); CacheService.getScriptCache().remove(key); }
function issueSession_(email) { var token = Utilities.getUuid() + Utilities.getUuid(), props = PropertiesService.getScriptProperties(), now = Date.now(); var all = props.getProperties(); Object.keys(all).forEach(function (key) { if (key.indexOf('session:') === 0) { try { if (JSON.parse(all[key]).expires < now) props.deleteProperty(key); } catch (e) { props.deleteProperty(key); } } }); props.setProperty('session:' + digest_(token), JSON.stringify({email:email,version:credentials_(email).version,expires:now+7*86400000})); return token; }
function sessionUser_(token) { var key = 'session:' + digest_(String(token || '')), props = PropertiesService.getScriptProperties(); var raw = props.getProperty(key) || CacheService.getScriptCache().get(key); if (!raw) fail_('auth','Please sign in again'); var session=JSON.parse(raw),user=findUser_(session.email),now=Date.now(); if (!user || user.status !== 'active' || session.version !== credentials_(session.email).version || (session.expires && session.expires < now)) { removeSession_(token); fail_('auth','Please sign in again'); } if (!session.expires || session.expires-now < 6*86400000) { session.expires=now+7*86400000;props.setProperty(key,JSON.stringify(session)); } return user; }
function saveNotice_(req, user) {
  var creating = !(Number(req.id) > 0);
  var target = NOTICE_TABS.filter(function (x) { return x.group === req.src; })[0];
  var nt = creating ? target : NOTICE_TABS.filter(function (x) { return x.tab === req.tab; })[0];
  if (!nt || !target) fail_('bad', 'Unknown notice group');
  if (creating) nt = target;
  var title = s_(req.title).trim(), body = s_(req.body).trim(), category = s_(req.cat).trim();
  if (!title || !category || title.length > 1000 || body.length > 40000 || category.length > 150) fail_('bad', 'Invalid notice content');
  var validDate = function (v) {
    if (!v) return true;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
    var d = new Date(v + 'T00:00:00Z');
    return !isNaN(d) && Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd') === v;
  };
  var date = s_(req.date), expires = s_(req.expires);
  if (!validDate(date) || !validDate(expires)) fail_('bad', 'Invalid date');
  var src = tab_(nt.tab), data = readTab_(src), id = creating ? Date.now() : Number(req.id), hit = null, seen = {};
  if (data.cols.notice_id === undefined) {
    data.cols.notice_id = src.getLastColumn();
    src.getRange(1, data.cols.notice_id + 1).setValue('Notice ID');
    data = readTab_(src);
  }
  if (!creating) {
    data.rows.forEach(function (r) {
      if (!s_(r.notice).trim()) return;
      r.date_key = parseDate_(r.date, r.year, r.month);
      if (noticeId_(nt.tab, r, seen) === id) hit = r;
    });
    if (!hit) {
      // Recover the edited row using the original snapshot when its legacy/generated ID changed.
      if (s_(req.revision).trim()) {
        for (var ri = 0; ri < data.rows.length; ri++) {
          var rr = data.rows[ri];
          if (!s_(rr.notice).trim()) continue;
          rr.date_key = parseDate_(rr.date, rr.year, rr.month);
          if (digest_(JSON.stringify(rr)) === s_(req.revision).trim()) { hit = rr; break; }
        }
      }
      if (!hit && s_(req.originalTitle).trim()) {
        var originalTitle = s_(req.originalTitle).trim(), originalBody = s_(req.originalBody).trim(), originalDate = s_(req.originalDate).trim();
        var candidates = data.rows.filter(function (rr) {
          if (!s_(rr.notice).trim()) return false;
          var rd = parseDate_(rr.date, rr.year, rr.month);
          var parts = splitNotice_(rr.notice);
          var rt = s_(rr.title).trim() || parts.title;
          var rb = (data.cols.body !== undefined && s_(rr.body)) ? s_(rr.body) : parts.body;
          return rt === originalTitle && (!originalDate || rd === originalDate) && (!originalBody || s_(rb).trim() === originalBody);
        });
        if (candidates.length === 1) hit = candidates[0];
      }
      if (!hit) fail_('notfound', 'The notice could not be located. Refresh the page and reopen it.');
    }
    // Persist the selected notice's stable ID only. Do not rewrite IDs across the entire sheet.
    // Bulk ID rewrites can cause stale-ID conflicts when legacy rows have generated IDs.
    if (data.cols.notice_id !== undefined && !(Number(hit.notice_id) > 0)) {
      src.getRange(hit._row, data.cols.notice_id + 1).setValue(id);
      hit.notice_id = id;
    }
  }
  var dest = tab_(target.tab), destData = target.tab === nt.tab ? data : readTab_(dest);
  var fields = {
    notice_id: id, title: title, body: body, notice: title + (body ? '\n\n' + body : ''),
    type: category, urgent: req.urgent ? 'Yes' : 'No', expired: req.expired ? 'Yes' : '',
    status: req.expired ? 'Expired' : 'Active', date: date, year: date ? date.slice(0, 4) : '',
    month: date ? MONTHS_[Number(date.slice(5, 7)) - 1] : '', expires: expires,
    announced_by: s_(req.ann).trim(), approved_by: s_(req.appr).trim(), sheet_update_by: user.name || user.email
  };
  var headerNames = {notice_id:'Notice ID', title:'Title', body:'Body', notice:'Notice', type:'Type', urgent:'Urgent', expired:'Expired', status:'Status', date:'Date', year:'Year', month:'Month', expires:'Expires', announced_by:'Announced By', approved_by:'Approved By', sheet_update_by:'Sheet Update By'};
  Object.keys(fields).forEach(function (key) {
    if (destData.cols[key] === undefined) {
      var col = dest.getLastColumn();
      dest.getRange(1, col + 1).setValue(headerNames[key]);
      destData.cols[key] = col;
    }
  });
  if (creating) {
    var values = new Array(dest.getLastColumn()).fill('');
    Object.keys(fields).forEach(function (key) { values[destData.cols[key]] = fields[key]; });
    dest.appendRow(values.map(function (v) { return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v; }));
    audit_(user, 'CREATE_NOTICE', 'notice', id, title, 'Notice created', target.tab);
  } else if (target.tab === nt.tab) {
    Object.keys(fields).forEach(function (key) {
      var value = fields[key];
      dest.getRange(hit._row, destData.cols[key] + 1).setValue(typeof value === 'string' && value.charAt(0) === '=' ? "'" + value : value);
    });
    audit_(user, 'EDIT_NOTICE', 'notice', id, title, 'Notice details updated', target.tab);
  } else {
    var values = new Array(dest.getLastColumn()).fill('');
    Object.keys(destData.cols).forEach(function (key) { if (hit[key] !== undefined) values[destData.cols[key]] = hit[key]; });
    Object.keys(fields).forEach(function (key) { values[destData.cols[key]] = fields[key]; });
    dest.appendRow(values.map(function (v) { return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v; }));
    src.deleteRow(hit._row);
    audit_(user, 'EDIT_NOTICE', 'notice', id, title, 'Notice updated and moved to ' + target.tab, target.tab);
  }
  if (req.expired) delReports_(id);
  return load_(user);
}
function deleteNotice_(req, user) {
  var nt = NOTICE_TABS.filter(function (x) { return x.tab === req.tab; })[0];
  if (!nt) fail_('bad', 'Unknown notice group');
  var sh = tab_(nt.tab), d = readTab_(sh), seen = {}, hit = null, id = Number(req.id);
  d.rows.forEach(function (r) {
    if (!s_(r.notice).trim()) return;
    r.date_key = parseDate_(r.date, r.year, r.month);
    if (noticeId_(nt.tab, r, seen) === id) hit = r;
  });
  if (!hit) fail_('bad', 'Notice not found');
  var title = s_(hit.title) || splitNotice_(hit.notice).title || 'Notice';
  sh.deleteRow(hit._row);
  delReports_(id);
  audit_(user, 'DELETE_NOTICE', 'notice', id, title, 'Notice permanently deleted', nt.tab);
  return load_(user);
}