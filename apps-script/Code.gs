/**
 * หลังบ้านหวย (Google Apps Script + Google Sheet)
 *
 * ติดตั้ง:
 *  1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางไฟล์นี้ทั้งไฟล์
 *  2) เลือกฟังก์ชัน setup แล้วกด Run (ครั้งแรกจะขออนุญาต กดยอมรับ)
 *     จะได้ชีต "หวย" "ตั้งค่า" "โพย"
 *  3) Project Settings > Script properties > เพิ่ม ADMIN_KEY = รหัสแอดมิน
 *  4) Deploy > New deployment > Web app
 *       Execute as: Me   |   Who has access: Anyone
 *     คัดลอก URL ที่ลงท้าย /exec ไปใส่ใน config.js
 *
 * แก้รายการหวยในชีต "หวย" และบัญชีโอนในชีต "ตั้งค่า" ได้ตลอด ไม่ต้อง deploy ใหม่
 * (แก้โค้ดในไฟล์นี้เมื่อไหร่ ต้อง Deploy > Manage deployments > New version)
 */

var TZ = 'Asia/Bangkok';
var AFTER_MIDNIGHT_CUTOFF_HOUR = 6;
var SH_LOTTO = 'หวย';
var SH_SETTINGS = 'ตั้งค่า';
var SH_ORDERS = 'โพย';
var ORDER_HEADERS = ['เวลา', 'เลขที่โพย', 'ชื่อ', 'เบอร์', 'หวย', 'งวด', 'ยอดรวม', 'รายการ'];

/* ===================== SETUP ===================== */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var lotto = ss.getSheetByName(SH_LOTTO) || ss.insertSheet(SH_LOTTO);
  if (lotto.getLastRow() === 0) {
    lotto.appendRow(['ชื่อหวย', 'เวลาปิด (HH:mm)', 'เปิดรับ', 'ชื่อเรียกอื่น (คั่นด้วยช่องว่าง)']);
    lotto.getRange('B:B').setNumberFormat('@');
    lotto.getRange(2, 1, 4, 3).setValues([
      ['หวยฮานอย', '18:10', true],
      ['หวยลาว', '20:15', true],
      ['หวยรัฐบาล', '15:00', false],
      ['ลาวดึก (ปิดหลังเที่ยงคืน)', '00:10', false]
    ]);
    lotto.getRange('C2:C').insertCheckboxes();
    lotto.setFrozenRows(1);
  }

  var st = ss.getSheetByName(SH_SETTINGS) || ss.insertSheet(SH_SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange('B:B').setNumberFormat('@');
    st.getRange(1, 1, 4, 2).setValues([
      ['หัวข้อ', 'ค่า'],
      ['ธนาคาร', ''],
      ['เลขบัญชี', ''],
      ['ชื่อบัญชี', '']
    ]);
    st.setFrozenRows(1);
  }

  var od = ss.getSheetByName(SH_ORDERS) || ss.insertSheet(SH_ORDERS);
  if (od.getLastRow() === 0) {
    od.appendRow(ORDER_HEADERS);
    od.getRange('B:D').setNumberFormat('@');
    od.getRange('F:F').setNumberFormat('@');
    od.setFrozenRows(1);
  }

  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

/* ===================== HTTP ===================== */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  try {
    if (action === 'boot') return json_({ ok: true, lottos: openLottos_(), bank: bank_() });
    return json_({ ok: true, message: 'lotto api' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'createOrder') return json_(createOrder_(body));
    if (body.action === 'adminOrders') {
      requireAdmin_(body.key);
      return json_({ ok: true, orders: listOrders_() });
    }
    return json_({ ok: false, error: 'ไม่รู้จัก action' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

/* ===================== LOTTOS ===================== */
function readLottos_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LOTTO);
  if (!sh || sh.getLastRow() < 2) return [];
  // คอลัมน์ D (ไม่บังคับ) = ชื่อเรียกอื่น คั่นด้วยช่องว่าง เช่น "ลาวex เอ็กตร้า"
  var cols = Math.max(3, Math.min(4, sh.getLastColumn()));
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues()
    .map(function (r) {
      return {
        name: String(r[0]).trim(),
        closeTime: normTime_(r[1]),
        open: r[2] === true || String(r[2]).toUpperCase() === 'TRUE',
        aliases: String(r[3] || '').trim().split(/\s+/).filter(Boolean)
      };
    })
    .filter(function (l) { return l.name && l.closeTime; });
}

function normTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  var m = String(v || '').match(/(\d{1,2})[:.](\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

/** งวดปัจจุบันของหวย + เวลาปิดจริง (หวยปิด 00:00-05:59 นับเป็นงวดของเมื่อวาน) */
function roundOf_(lotto, now) {
  now = now || new Date();
  var hm = lotto.closeTime.split(':');
  var closeH = Number(hm[0]), closeM = Number(hm[1]);
  var nowH = Number(Utilities.formatDate(now, TZ, 'H'));
  var nowM = Number(Utilities.formatDate(now, TZ, 'm'));
  var round = new Date(now.getTime());
  var after = closeH < AFTER_MIDNIGHT_CUTOFF_HOUR;

  if (after && (nowH * 60 + nowM) < (closeH * 60 + closeM)) {
    round = new Date(now.getTime() - 24 * 3600 * 1000);
  }
  var roundYmd = Utilities.formatDate(round, TZ, 'yyyy-MM-dd');
  var close = Utilities.parseDate(roundYmd + ' ' + lotto.closeTime, TZ, 'yyyy-MM-dd HH:mm');
  if (after) close = new Date(close.getTime() + 24 * 3600 * 1000);

  return {
    drawDate: Utilities.formatDate(round, TZ, 'dd-MM-yyyy'),
    close: close
  };
}

function openLottos_() {
  var now = new Date();
  return readLottos_()
    .filter(function (l) { return l.open; })
    .map(function (l) {
      var r = roundOf_(l, now);
      return { name: l.name, closeTime: l.closeTime, drawDate: r.drawDate, aliases: l.aliases };
    });
}

function bank_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SETTINGS);
  if (!sh || sh.getLastRow() < 2) return null;
  var map = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { map[String(r[0]).trim()] = String(r[1]).trim(); });
  if (!map['เลขบัญชี']) return null;
  return { bank: map['ธนาคาร'] || '', number: map['เลขบัญชี'], name: map['ชื่อบัญชี'] || '' };
}

/* ===================== ORDERS ===================== */
var LEN_BY_MODE = { two_top: 2, two_bottom: 2, three_top: 3, three_tod: 3, three_return: 3 };

function createOrder_(body) {
  var name = String(body.customerName || '').trim().slice(0, 60);
  var phone = String(body.customerPhone || '').replace(/[^\d+]/g, '').slice(0, 20);
  if (!name) throw new Error('กรุณากรอกชื่อผู้แทง');

  var lotto = readLottos_().filter(function (l) { return l.name === body.lottoName; })[0];
  if (!lotto || !lotto.open) throw new Error('หวยนี้ไม่เปิดรับ');
  var round = roundOf_(lotto);
  if (round.close.getTime() <= Date.now()) throw new Error('หวยนี้ปิดรับแล้ว');

  var items = Array.isArray(body.items) ? body.items : [];
  if (!items.length || items.length > 300) throw new Error('รายการไม่ถูกต้อง');

  // คิดเงินใหม่ฝั่งเซิร์ฟเวอร์ ไม่เชื่อยอดจากหน้าเว็บ
  var total = 0, clean = [], nCount = 0;
  items.forEach(function (it) {
    var len = LEN_BY_MODE[it.modeKey];
    if (!len) throw new Error('ประเภทไม่ถูกต้อง');
    var nums = String(it.numbers || '').trim().split(/\s+/).filter(Boolean);
    nums.forEach(function (n) { if (!/^\d+$/.test(n) || n.length !== len) throw new Error('เลขไม่ถูกต้อง: ' + n); });
    var top = Math.max(0, Math.floor(Number(it.top) || 0));
    var bottom = Math.max(0, Math.floor(Number(it.bottom) || 0));
    if (!nums.length || top + bottom <= 0) throw new Error('รายการไม่ถูกต้อง');
    nCount += nums.length;
    var subtotal = nums.length * (top + bottom);
    total += subtotal;
    clean.push({ label: String(it.label || '').slice(0, 40), modeKey: it.modeKey, numbers: nums.join(' '), top: top, bottom: bottom, count: nums.length, subtotal: subtotal });
  });
  if (nCount > 3000) throw new Error('จำนวนเลขมากเกินไป');

  var now = new Date();
  var ref = Utilities.formatDate(now, TZ, 'ddMM-HHmm') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ORDERS).appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), ref, name, phone,
      lotto.name, round.drawDate, total, JSON.stringify(clean)
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, ref: ref, total: total, drawDate: round.drawDate };
}

function listOrders_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ORDERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - 4999);
  return sh.getRange(start, 1, last - start + 1, ORDER_HEADERS.length).getValues().map(function (r) {
    var items = [];
    try { items = JSON.parse(r[7] || '[]'); } catch (e) {}
    return {
      time: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
      ref: String(r[1]), name: String(r[2]), phone: String(r[3]),
      lotto: String(r[4]), drawDate: String(r[5]), total: Number(r[6]) || 0, items: items
    };
  });
}

/* ===================== UTIL ===================== */
function requireAdmin_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real) throw new Error('ยังไม่ได้ตั้ง ADMIN_KEY ใน Script properties');
  if (String(key || '') !== real) throw new Error('รหัสไม่ถูกต้อง');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function msg_(err) { return String(err && err.message || err); }
