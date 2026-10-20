/**
 * 高蛋白 CP 值計算器：共用清單後端（Google Apps Script）
 * 從 Google 試算表的「擴充功能 → Apps Script」打開，貼上這份程式，再部署成網頁應用程式。
 *
 * 試算表裡會自動建立兩個工作表：
 *   清單     — 大家看到的共用清單，也可以直接在表格裡改
 *   刪除紀錄 — 網頁上被刪掉的項目會先搬到這裡，刪錯可以找回來
 * 如果「清單」的欄位是舊格式，會自動改名成「清單(舊)」，再建立新的「清單」。
 */

var SHEET = '清單';
var TRASH = '刪除紀錄';
var HEAD = ['id', '新增日期', '購買位置', '品名', '分類', '總份數', '每份蛋白質(g)',
            '原價(元)', '折扣券(元)', '折扣後價格(元)', '原價 g/元', '折扣後 g/元', '評價(折扣後)', '更新時間'];
var C = { id: 0, date: 1, place: 2, name: 3, cat: 4, n: 5, p: 6, price: 7, coupon: 8 };
var CATS = ['豆漿', '濃縮乳清', '分離乳清', '水解乳清'];
var NAMES = ['推薦購買', '可購買', '考慮購買', '猶豫'];
var LEVELS = { '豆漿': [1.2, 1.05, 0.9, 0.75], '濃縮乳清': [0.9, 0.8, 0.7, 0.6],
               '分離乳清': [0.75, 0.65, 0.55, 0.45], '水解乳清': [0.55, 0.45, 0.4, 0.35] }; // 評價(折扣後)欄用的預設標準
var TZ = 'Asia/Taipei';

function doGet() {
  return json_({ ok: true, items: list_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'put') put_(body.id, body.rec);
    else if (body.action === 'del') del_(body.id);
    else if (body.action === 'batch') (body.recs || []).slice(0, 200).forEach(function (r) { put_(r.id, r); });
    else throw new Error('未知的動作');
    return json_({ ok: true, items: list_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), items: list_() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 內部函式 ---------- */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var head = name === TRASH ? HEAD.concat(['刪除時間']) : HEAD;
  var sh = ss.getSheetByName(name);
  if (sh && sh.getLastColumn() > 0) {
    var cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join('|');
    if (cur !== head.join('|')) { // 舊格式：改名保留，另建新的
      sh.setName(name + '(舊)' + Utilities.formatDate(new Date(), TZ, 'MMdd-HHmm'));
      sh = null;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 日期當文字存，保持 2026/10/20 格式
  }
  return sh;
}

function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy/MM/dd');
  return String(v || '');
}

function list_() {
  var sh = sheet_(SHEET);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  var out = [];
  rows.forEach(function (r, i) {
    var n = Number(r[C.n]), p = Number(r[C.p]), price = Number(r[C.price]), coupon = Number(r[C.coupon]) || 0;
    if (!String(r[C.name]).trim() || !(n > 0 && p > 0 && price > 0) || coupon < 0 || coupon >= price) return;
    var id = String(r[C.id] || '');
    if (!id) { // 直接在表格裡新增的列，補上 id
      id = 's' + Utilities.getUuid().slice(0, 8);
      sh.getRange(i + 2, 1).setValue(id);
    }
    out.push({ id: id, date: dateText_(r[C.date]), place: String(r[C.place] || ''), name: String(r[C.name]).trim(),
               cat: String(r[C.cat] || ''), n: n, p: p, price: price, coupon: coupon });
  });
  return out;
}

function clean_(rec) {
  rec = rec || {};
  var name = String(rec.name || '').trim().slice(0, 60);
  var place = String(rec.place || '').trim().slice(0, 20);
  var n = Number(rec.n), p = Number(rec.p), price = Number(rec.price), coupon = Number(rec.coupon) || 0;
  if (!name) throw new Error('缺少品名');
  if (!(n > 0 && n <= 1000 && p > 0 && p <= 200 && price > 0 && price <= 100000)) throw new Error('數字不合理');
  if (coupon < 0 || coupon >= price) throw new Error('折扣券不合理');
  var date = /^\d{4}\/\d{2}\/\d{2}$/.test(rec.date || '') ? rec.date : Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');
  var cat = CATS.indexOf(rec.cat) >= 0 ? rec.cat : '';
  return { name: name, place: place, cat: cat, n: n, p: p, price: price, coupon: coupon, date: date };
}

function row_(id, r) {
  var tot = r.n * r.p, paid = r.price - r.coupon, g = tot / paid, lv = '不太推薦';
  var t = LEVELS[r.cat] || LEVELS['濃縮乳清'];
  for (var i = 0; i < 4; i++) if (g >= t[i]) { lv = NAMES[i]; break; }
  var round = function (x) { return Math.round(x * 100) / 100; };
  return [id, r.date, r.place, r.name, r.cat, r.n, r.p, r.price, r.coupon, paid, round(tot / r.price), round(g), lv, new Date()];
}

function find_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function put_(id, rec) {
  id = String(id || '').slice(0, 40);
  if (!id) throw new Error('缺少 id');
  var r = clean_(rec), sh = sheet_(SHEET), at = find_(sh, id);
  if (at > 0) {
    var old = dateText_(sh.getRange(at, C.date + 1).getValue());
    if (old) r.date = old; // 修改時保留原本的新增日期
    sh.getRange(at, 1, 1, HEAD.length).setValues([row_(id, r)]);
  } else {
    sh.appendRow(row_(id, r));
  }
}

function del_(id) {
  var sh = sheet_(SHEET), at = find_(sh, id);
  if (at < 0) return;
  var row = sh.getRange(at, 1, 1, HEAD.length).getValues()[0];
  sheet_(TRASH).appendRow(row.concat([new Date()]));
  sh.deleteRow(at);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
