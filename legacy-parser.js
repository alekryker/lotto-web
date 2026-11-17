/* =====================================================================
   ตัวอ่านโพยจากบอทเดิม (legacy-bot_V1117_RUDMODE_NATURAL.gs)
   สร้างอัตโนมัติด้วย tools/build-legacy-parser.js — อย่าแก้ไฟล์นี้ด้วยมือ
   ข้อมูลลับ (API key, SHEET_ID, บัญชี) ถูกลบออกแล้ว
   ใช้: LegacyParser.parse(ข้อความ, [{name, closeTime, aliases}]) → { ok, lotto, groups, total, error }
===================================================================== */
(function (window) {

  var stub = new Proxy(function () {}, {
    get: function (t, k) { return k === Symbol.toPrimitive ? function () { return ""; } : stub; },
    apply: function () { return stub; },
    construct: function () { return stub; }
  });

  /* Utilities.formatDate แบบใช้เวลาจริง (isClosed ใช้) */
  var Utilities = {
    formatDate: function (date, tz, fmt) {
      var p = {};
      new Intl.DateTimeFormat("en-GB", { timeZone: tz || "Asia/Bangkok", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" })
        .formatToParts(date).forEach(function (x) { p[x.type] = x.value; });
      return String(fmt).replace("yyyy", p.year).replace("MM", p.month).replace("dd", p.day)
        .replace("HH", p.hour).replace("mm", p.minute).replace("ss", p.second);
    },
    sleep: function () {}
  };
  var quiet = { log: function () {}, info: function () {}, warn: function () {}, error: function () {} };
  var sandbox = {};

  var L = (function (SpreadsheetApp, PropertiesService, CacheService, Session, LockService, UrlFetchApp, HtmlService, ContentService, ScriptApp, DriveApp, MailApp, Logger, CardService, DocumentApp, GmailApp, Utilities, console, globalThis) {
// ============================================================================
// LEGACY BOT PATCH VERSION
// ============================================================================
const LEGACY_BOT_RAW_PATCH_VERSION_ = "V1117-RUDMODE-NATURAL";
console.log("LEGACY_BOT_RAW_PATCH_VERSION", LEGACY_BOT_RAW_PATCH_VERSION_);

// ============================================================================
// LEGACY BOT PATCH VERSION
// ============================================================================
const LEGACY_BOT_PATCH_VERSION_ = "V1115-FIX304-PREPARSE";
console.log("LEGACY_BOT_PATCH_VERSION", LEGACY_BOT_PATCH_VERSION_);

function getLegacyBotPatchVersion_() {
  return LEGACY_BOT_PATCH_VERSION_;
}
globalThis.getLegacyBotPatchVersion_ = getLegacyBotPatchVersion_;

// ================== CONFIG ==================
const CHANNEL_ACCESS_TOKEN = "";
const SHEET_ID = "";

const DOOR_SHEET = "ชีต2";
const LOTTO_SHEET = "ชีต3";
const KEY_SHEET = "ชีต4";

const TZ = "Asia/Bangkok";

const BANK_ACCOUNT_NUMBER = "";
const BANK_ACCOUNT_BANK = "";
const BANK_ACCOUNT_NAME = "";

const THUNDER_API_KEY = "";
const THUNDER_VERIFY_BANK_URL = "https://api.thunder.in.th/v2/verify/bank";
const SLIP_VERIFY_CHECK_DUPLICATE = true;
const SLIP_VERIFY_MATCH_ACCOUNT = true;

// Thunder: จำกัดเวลาต่อครั้งให้สั้น แล้ว retry อัตโนมัติ
// รวมสูงสุดประมาณ 45 วินาที เพื่อไม่ให้ LINE รอนานเกินไป
const THUNDER_VERIFY_MAX_ATTEMPTS = 3;
const THUNDER_VERIFY_TIMEOUT_SECONDS = 15;
const THUNDER_VERIFY_CONNECT_TIMEOUT_SECONDS = 5;

const SLIP_IMAGE_MAX_BYTES = 4 * 1024 * 1024;
const DATA_CACHE_TTL_SECONDS = 120;

// รวมข้อความโพยที่ลูกค้าส่งแยกหลายข้อความติดกัน
const BET_BUFFER_TTL_SECONDS = 45;
const BET_BUFFER_MAX_LENGTH = 6000;
// รอให้ webhook หลายข้อความที่เข้ามาเกือบพร้อมกันเขียน Buffer ให้ครบ
// แล้วให้เฉพาะข้อความล่าสุดเป็นตัวประมวลผลโพยรวม
const BET_BUFFER_DEBOUNCE_MS = 800;

// ราคาเริ่มต้นสำหรับ 3 ตัว 6กลับ เมื่อระบุคำว่า 6กลับ แต่ไม่ได้ใส่ราคา
// ตัวอย่าง: 878-839-783 6กลับ => ใช้ราคา 20 บาท/ชุดกลับ
const DEFAULT_6_RETURN_PRICE = 20;

const HAM_NUMBERS = [
  "010","020","030","040","050","060","070","080","090",
  "101","121","131","141","151","161","171","181","191",
  "202","212","232","242","252","262","272","282","292",
  "303","313","323","343","353","363","373","383","393",
  "404","414","424","434","454","464","474","484","494",
  "505","515","525","535","545","565","575","585","595",
  "606","616","626","636","646","656","676","686","696",
  "707","717","727","737","747","757","767","787","797",
  "808","818","828","838","848","858","868","878","898",
  "909","919","929","939","949","959","969","979","989"
];

// ================== MAIN ==================
function doGet(e) {
  const __act = (e && e.parameter && e.parameter.action) || "";
  const __page = (e && e.parameter && e.parameter.page) || "";

  if (__act === "boot") return handleLiffBoot_();

  // หน้าแอดมินกรอกผลหวย
  if (__page === "result") {
    return HtmlService
      .createHtmlOutputFromFile("ResultAdmin")
      .setTitle("ตรวจผลหวย")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover"
      );
  }

  // หน้าแยกดึงผล ExpHuay
  if (__page === "exphuay") {
    return HtmlService
      .createHtmlOutputFromFile("ExphuayResult")
      .setTitle("ดึงผลหวย ExpHuay")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover"
      );
  }
// หน้า Dashboard สรุปยอดหวย
if (__page === "dashboard") {
  return HtmlService
    .createHtmlOutputFromFile("Dashboard")
    .setTitle("Dashboard หวย")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover"
    );
}
  // หน้าผลหวยสด — อัปเดตอัตโนมัติ
  if (
    __page === "live-result" ||
    __page === "live"
  ) {
    return HtmlService
      .createHtmlOutputFromFile(
        "LiveResult"
      )
      .setTitle(
        "ผลหวยเรียลไทม์"
      )
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover"
      );
  }

  // หน้าแทงหวยเดิม
  return HtmlService
    .createHtmlOutput(getLiffOrderHtml_())
    .setTitle("ซื้อหวย")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover");
}

function getLiffId_() {
  return PropertiesService.getScriptProperties().getProperty("LIFF_ID") ||
    PropertiesService.getScriptProperties().getProperty("2001131713-NBQyOwvg") ||
    "2001131713-NBQyOwvg";
}

function getLiffBootData() {
  const lottoData = loadLotto();
  return {
    liffId: getLiffId_(),
    lottos: getOpenLottoCandidates(lottoData)
  };
}

function createLiffOrder(payload) {
  payload = payload || {};
  const lottoName = String(payload.lottoName || "").trim();
  const kind = String(payload.kind || "").trim();
  const numbers = normalizeLiffNumbers_(payload.numbers, kind);
  const top = Number(payload.top || 0);
  const bottom = Number(payload.bottom || 0);

  if (!lottoName) throw new Error("กรุณาเลือกหวย");
  if (!kind) throw new Error("กรุณาเลือกประเภทการซื้อ");
  if (numbers.length === 0) throw new Error("กรุณาใส่เลข");
  if (top <= 0 && bottom <= 0) throw new Error("กรุณาใส่ราคา");

  const lottoData = loadLotto();
  const closeTime = lottoData.map[lottoName];
  if (!closeTime) throw new Error("ไม่พบหวยที่เลือก");

  const liffRound = getLiffLottoRoundInfo_(lottoName, closeTime, new Date());
  if (!liffRound.open) throw new Error(lottoName + " ปิดรับแล้ว (" + closeTime + ")");
  const roundDate = liffRound.roundDate;

  const keywords = loadKeywords();
  const doorMap = loadDoors();
  const cmd = buildLiffCommand_(kind, numbers, top, bottom);
  const result = parseMessage(cmd, keywords, doorMap);
  if (!result.success) throw new Error(result.error || "คำนวณโพยไม่สำเร็จ");

  // เก็บโพยจาก LIFF ก่อนส่ง Flex เสมอ
  // ถ้ายังไม่มี LINE userId ห้ามถือว่าโพยสำเร็จ เพราะจะตรวจผลย้อนหลังไม่ได้
  const liffUserId = String(payload.userId || "").trim();
  if (!liffUserId) {
    throw new Error("ยังไม่พบ LINE userId กรุณารอให้ LIFF โหลดเสร็จแล้วส่งใหม่");
  }

  saveParsedBetByUserId_(
    liffUserId,
    lottoName,
    closeTime,
    lottoName + "\n" + cmd,
    result.groups,
    result.totalAmount,
    result.totalDoors,
    "",
    roundDate
  );

  return {
    flex: buildFlex({
      name: lottoName,
      closeTime: closeTime,
      roundDate: roundDate,
      groups: result.groups,
      totalDoors: result.totalDoors,
      totalAmount: result.totalAmount
    }),
    text: lottoName + "\n" + cmd,
    totalAmount: result.totalAmount,
    totalDoors: result.totalDoors
  };
}

function normalizeLiffNumbers_(raw, kind) {
  const expectedLen = (kind === "two") ? 2 : 3;
  return String(raw || "")
    .replace(/[,，]+/g, " ")
    .replace(/[\/\-\_]+/g, " ")
    .split(/\s+/)
    .map(n => n.trim())
    .filter(Boolean)
    .map(n => (expectedLen === 2 && n.length === 1) ? "0" + n : n)
    .filter(n => new RegExp("^\\d{" + expectedLen + "}$").test(n));
}

function buildLiffCommand_(kind, numbers, top, bottom) {
  const nums = numbers.join(" ");
  if (kind === "three") {
    const price = bottom > 0 ? top + "*" + bottom : String(top);
    return nums + " ตรงโต๊ด " + price;
  }
  if (kind === "two") {
    const price = bottom > 0 ? top + "*" + bottom : String(top);
    return nums + " บลก " + price;
  }
  if (kind === "three_return") {
    return nums + " 6กลับ " + top;
  }
  throw new Error("ประเภทการซื้อไม่ถูกต้อง");
}



// ============================================================================
// HIGH PRIORITY RAW NORMALIZER — ก่อน Buffer / Lotto routing / legacy parser
//
// แก้เคสที่ parser เก่าเอาราคาบรรทัดท้ายไหลย้อนมาทับเลขก่อนหน้า
// เช่น:
//   02/19/36/45
//   37/81/01
//   8บลก
//   11=40บล
// ->
//   02 19 36 45 37 81 01 บลก 8*8
//   11 บล 40*40
//
// และ:
//   รูด7/5
//   5*5
// -> รูด 7/5 บล 5*5
// ============================================================================
function normalizeRawHighPriorityBetSyntax_(input) {
  const lines = String(input || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map(function(line) {
      return String(line || "")
        .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
        .replace(/\u00A0/g, " ")
        .trim();
    });

  function parse2List_(line) {
    const s = String(line || "").trim();
    if (!/^\d{2}(?:\s*[,\/\\._-]?\s*\d{2})*$/.test(s)) return null;
    const nums = s
      .replace(/[,\/\\._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .filter(Boolean);
    if (!nums.length || !nums.every(function(n) { return /^\d{2}$/.test(n); })) return null;
    return nums;
  }

  function canonMode_(mode) {
    const m = String(mode || "").replace(/\s+/g, "");
    if (m === "บกล" || m === "บนล่างกลับ") return "บลก";
    if (m === "บนล่าง") return "บล";
    if (m === "บ") return "บน";
    if (m === "ล") return "ล่าง";
    return m;
  }

  const out = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // ======================================================
    // PATCH 2026-11-17 — ชื่อหวย + บลก แล้วเลขรูด=ราคา
    //
    // ตัวอย่าง:
    //   กาชาด บลก
    //   4/8=5*5
    //
    // canonical:
    //   กาชาด
    //   รูด 4/8 บลก 5*5
    // ======================================================
    const lottoModeLine = line.match(
      /^(.+?)\s+(บลก|บกล|บน\s*ล่าง\s*กลับ|บนล่างกลับ)\s*$/i
    );

    if (
      lottoModeLine &&
      i + 1 < lines.length
    ) {
      const roodAssigned = String(lines[i + 1] || "")
        .replace(/\s+/g, "")
        .match(
          /^([0-9](?:[\/,\-][0-9]){1,6})=(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)$/
        );

      if (roodAssigned) {
        const lottoName =
          String(lottoModeLine[1] || "").trim();

        const mode =
          canonMode_(lottoModeLine[2]);

        const digits =
          String(roodAssigned[1] || "")
            .replace(/[,\-]+/g, "/");

        out.push(lottoName);
        out.push(
          "รูด " +
          digits +
          " " +
          mode +
          " " +
          roodAssigned[2] +
          "*" +
          roodAssigned[3]
        );

        i += 2;
        continue;
      }
    }

    // ======================================================
    // PATCH 2026-11-17 — ภาษาพูด:
    //
    //   รูด2/3ตัวล่ะ5บาท
    //   บน/ล่าง/กลับตัวล่ะ5บาท
    //   30/31/36/37/38/39
    //
    // canonical:
    //   รูด 2/3 บลก 5*5
    //   30 31 36 37 38 39 บลก 5*5
    //
    // รองรับ "ตัวละ" และ "ตัวล่ะ"
    // ======================================================
    const naturalRood = line.match(
      /^(?:รูด|รุด|ลูด)\s*([0-9](?:\s*[\/,\-]\s*[0-9]){1,6}|[0-9]{2,7})\s*(?:ตัวละ|ตัวล่ะ|ตัว)\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?)?\s*$/i
    );

    if (
      naturalRood &&
      i + 2 < lines.length
    ) {
      const modePriceLine =
        String(lines[i + 1] || "")
          .trim()
          .match(
            /^(?:บน\s*[\/\-]\s*ล่าง\s*[\/\-]\s*กลับ|บนล่างกลับ|บลก|บกล)\s*(?:ตัวละ|ตัวล่ะ|ตัว)\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?)?\s*$/i
          );

      const nums =
        parse2List_(lines[i + 2]);

      if (
        modePriceLine &&
        nums &&
        nums.length > 0
      ) {
        let digits =
          String(naturalRood[1] || "")
            .replace(/\s+/g, "");

        if (
          /^\d{2,7}$/.test(digits) &&
          !/[\/,\-]/.test(digits)
        ) {
          digits =
            digits.split("").join("/");
        } else {
          digits =
            digits.replace(/[,\-]+/g, "/");
        }

        const roodPrice =
          String(naturalRood[2] || "");

        const listPrice =
          String(modePriceLine[1] || "");

        out.push(
          "รูด " +
          digits +
          " บลก " +
          roodPrice +
          "*" +
          roodPrice
        );

        out.push(
          nums.join(" ") +
          " บลก " +
          listPrice +
          "*" +
          listPrice
        );

        i += 3;
        continue;
      }
    }

    // A) รูด/รุด/ลูด + ราคาแยกบรรทัด
    const rm = line.match(
      /^(?:รูด|รุด|ลูด)\s*([0-9](?:\s*[,\/-]\s*[0-9]){0,6}|[0-9]{2,7})\s*$/
    );

    if (rm && i + 1 < lines.length) {
      const pm = String(lines[i + 1] || "")
        .replace(/\s+/g, "")
        .match(/^(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)$/);

      if (pm) {
        let digits = String(rm[1] || "").replace(/\s+/g, "");
        if (/^\d{2,7}$/.test(digits) && !/[\/,\-]/.test(digits)) {
          digits = digits.split("").join("/");
        } else {
          digits = digits.replace(/[,\-]+/g, "/");
        }

        out.push("รูด " + digits + " บล " + pm[1] + "*" + pm[2]);
        i += 2;
        continue;
      }
    }

    // B) สะสมเลข 2 ตัวหลายบรรทัด แล้วเจอ ราคา+mode เช่น 8บลก
    const first = parse2List_(line);
    if (first) {
      let nums = first.slice();
      let j = i + 1;
      while (j < lines.length) {
        const more = parse2List_(lines[j]);
        if (!more) break;
        nums.push.apply(nums, more);
        j++;
      }

      if (j < lines.length) {
        const priceMode = String(lines[j] || "").match(
          /^(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
        );

        if (priceMode) {
          const p = String(priceMode[1]);
          const mode = canonMode_(priceMode[2]);
          out.push(nums.join(" ") + " " + mode + " " + p + "*" + p);
          i = j + 1;
          continue;
        }
      }
    }

    // C) เลขเดี่ยว=ราคา+mode เช่น 11=40บล
    const eq = line.match(
      /^(\d{2})\s*=\s*(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
    );

    if (eq) {
      const mode = canonMode_(eq[3]);
      const p = String(eq[2]);
      out.push(eq[1] + " " + mode + " " + p + "*" + p);
      i++;
      continue;
    }

    out.push(line);
    i++;
  }

  return out.join("\n");
}

function doPost(e) {
  try {
    const webhookStartMs = Date.now();

    if (!e || !e.postData || !e.postData.contents) {
      logBotTiming("manualRun", 0, {
        message: "doPost called without webhook payload"
      });

      return ContentService.createTextOutput("OK");
    }

    const data = JSON.parse(e.postData.contents);

    // =====================================================
    // LIFF ORDER API
    // =====================================================
    if (data && data.action === "createLiffOrder") {
      return jsonResponse_(
        createLiffOrder(
          data.payload || data
        )
      );
    }

    if (data && data.action === "liffOrder") {
      return handleLiffOrderRequest_(data);
    }

    if (!data.events) {
      return ContentService.createTextOutput("OK");
    }


    // =====================================================
    // LINE EVENTS
    // =====================================================
    data.events.forEach(function(ev, eventIndex) {

      const eventStartMs = Date.now();

      const eventMeta = {
        eventIndex: eventIndex,
        eventType: ev && ev.type,
        messageType:
          ev &&
          ev.message &&
          ev.message.type,

        isRedelivery: !!(
          ev &&
          ev.deliveryContext &&
          ev.deliveryContext.isRedelivery
        )
      };


      try {

        // =====================================================
        // LINE ส่ง webhook ซ้ำ
        // =====================================================
        if (
          ev.deliveryContext &&
          ev.deliveryContext.isRedelivery
        ) {
          logBotTiming(
            "skipRedelivery",
            0,
            eventMeta
          );

          return;
        }


        // =====================================================
        // อนุมัติผลหวย
        //
        // สำคัญ:
        // เรียก handler นี้เฉพาะ POSTBACK เท่านั้น
        // เดิมถูกเรียกกับทุกข้อความ ถ้าระบบอนุมัติผลสะดุด
        // จะทำให้ข้อความปกติและคำสั่ง "เทส" ตายไปด้วย
        // =====================================================
        if (
          ev.type === "postback" &&
          typeof handleLotteryApprovalWebhookEvent_ === "function"
        ) {
          try {
            if (handleLotteryApprovalWebhookEvent_(ev)) {
              eventMeta.lotteryApprovalHandled = true;
              return;
            }
          } catch (approvalErr) {
            console.error(
              "LOTTERY_APPROVAL_HANDLER_FAILED " +
              String(
                approvalErr && approvalErr.stack
                  ? approvalErr.stack
                  : approvalErr
              )
            );
            eventMeta.lotteryApprovalError = String(
              approvalErr && approvalErr.message
                ? approvalErr.message
                : approvalErr
            );
          }
        }


        // =====================================================
        // รับเฉพาะ MESSAGE
        // =====================================================
        if (
          ev.type !== "message" ||
          !ev.message
        ) {
          return;
        }


        const replyToken = ev.replyToken;


        // =====================================================
        // รูปสลิป
        // =====================================================
        if (ev.message.type === "image") {

          // PATCH 2026-09-21:
          // อ่านโพยจากรูปก่อน ถ้าไม่ใช่รูปโพยค่อยส่งไปตรวจสลิป
          handleIncomingLineImage_(
            replyToken,
            ev.message.id,
            ev
          );

          return;
        }


        // =====================================================
        // รับเฉพาะข้อความ TEXT
        // =====================================================
        if (ev.message.type !== "text") {
          return;
        }


        // =====================================================
        // ทำความสะอาดข้อความจาก LINE
        // =====================================================
        let text =
          cleanIncomingLineText_(
            ev.message.text
          );

        if (!text) {
          return;
        }

        // =====================================================
        // DEBUG VERSION — เช็กว่า LINE วิ่งเข้าไฟล์ใหม่นี้จริง
        // =====================================================
        if (/^(?:เวอร์ชันบอท|เช็คเวอร์ชัน|version)$/i.test(text)) {
          clearBetBuffer_(ev);
          replyText(
            replyToken,
            "LEGACY " + LEGACY_BOT_RAW_PATCH_VERSION_
          );
          return;
        }

        // =====================================================
        // HIGH PRIORITY RAW NORMALIZE
        // ต้องทำก่อน Buffer / หาเลขหวย / parser ทุกชนิด
        // =====================================================
        text = normalizeRawHighPriorityBetSyntax_(text);

        eventMeta.textHead =
          text.substring(0, 20);


        // =====================================================
        // HEALTH CHECK ของบอท
        //
        // ต้องอยู่ก่อน handler อื่นทั้งหมด
        // ถ้า webhook เข้ามาถึง doPost คำว่า "เทส" ต้องตอบได้
        // แม้ระบบเมนู / ชีต / parser / ระบบผลหวยบางส่วนมีปัญหา
        // =====================================================
        if (isTestCommand(text)) {
          clearBetBuffer_(ev);

          const testReply = replyText(
            replyToken,
            "ใช้งานได้"
          );

          eventMeta.healthCheck = true;
          eventMeta.healthReplyStatus =
            testReply && testReply.status
              ? testReply.status
              : 0;

          return;
        }


        // =====================================================
        // OCR รูปโพย -> รอข้อความราคา
        //
        // Flow ใหม่:
        //   1) ลูกค้าส่งรูป
        //   2) OCR เก็บเลขไว้ใน pending
        //   3) ลูกค้าพิมพ์ราคา เช่น "กับ10*10"
        //   4) จึงค่อยขึ้น Quick Reply ให้เลือกชื่อหวย
        // =====================================================
        const imageBetUserKey =
          getLineUserKey(ev);

        if (
          imageBetUserKey &&
          tryHandlePendingImageBetPrice_(
            replyToken,
            imageBetUserKey,
            text,
            ev
          )
        ) {
          eventMeta.imageBetPriceHandled = true;
          return;
        }


        // =====================================================
        // คำสั่งสรุปผลหวย
        //
        // สร
        //   = สรุปผลหวยของวันนี้
        //
        // สร 03-09-2026
        // สร 03/09/2026
        // สร03-09-2569
        //   = สรุปผลหวยของวันที่ระบุ
        // =====================================================
        if (isLotteryResultSummaryCommand_(text)) {
          clearBetBuffer_(ev);

          try {
            replyLotteryResultSummary_(
              replyToken,
              text
            );

            eventMeta.resultSummary = true;
          } catch (summaryErr) {
            console.error(
              "RESULT_SUMMARY_FAILED " +
              String(
                summaryErr && summaryErr.stack
                  ? summaryErr.stack
                  : summaryErr
              )
            );

            replyText(
              replyToken,
              "⚠️ ไม่สามารถสรุปผลหวยได้ในขณะนี้ กรุณาลองใหม่อีกครั้ง"
            );

            eventMeta.resultSummaryError =
              String(
                summaryErr && summaryErr.message
                  ? summaryErr.message
                  : summaryErr
              );
          }

          return;
        }


        // =====================================================
        // เมนู "ทรัพย์มั่งมี"
        //
        // อ่านข้อมูลจากชีต:
        // รวมลิ้งค์
        //
        // A = ชื่อเมนู
        // B = ลิงก์
        //
        // สำคัญ:
        // ต้องอยู่ก่อนระบบ Buffer / Parser หวย
        // =====================================================
        let sapMungMeeHandled = false;

        try {
          if (
            typeof handleSapMungMeeMenuCommand_ === "function"
          ) {
            sapMungMeeHandled =
              !!handleSapMungMeeMenuCommand_(
                replyToken,
                text
              );
          }
        } catch (menuErr) {
          console.error(
            "SAP_MUNG_MEE_MENU_FAILED " +
            String(
              menuErr && menuErr.stack
                ? menuErr.stack
                : menuErr
            )
          );
          eventMeta.sapMungMeeMenuError = String(
            menuErr && menuErr.message
              ? menuErr.message
              : menuErr
          );
        }

        if (sapMungMeeHandled) {

          // ล้างโพยค้าง
          // ป้องกันคำว่า "ทรัพย์มั่งมี"
          // ไปต่อกับโพยก่อนหน้า
          clearBetBuffer_(ev);


          const menuUserKey =
            getLineUserKey(ev);


          if (menuUserKey) {
            clearPendingBet(
              menuUserKey
            );
          }


          eventMeta.sapMungMeeMenu =
            true;

          return;
        }


        // =====================================================
        // ข้อความแจ้งถูก / แจ้งยอด
        //
        // ลูกค้าไม่ได้กำลังแทงหวย
        // ให้บอทเงียบ
        //
        // เช่น
        // ถูก ลาว 100
        // ลาวถูก 0=100
        // ลาวสตาร์ ถูก 5=500
        // ฮานอยถูก=200
        // =====================================================
        if (
          isCustomerResultReport_(
            text
          )
        ) {
          return;
        }


        // =====================================================
        // AUTO REPLY จากชีต4
        // A = คีย์เวิร์ด
        // B = ข้อความตอบกลับ
        // ต้องอยู่ก่อนระบบ Buffer / Parser หวย
        // =====================================================
        try {
          if (
            typeof tryReplyKeywordFromSheet4_ === "function" &&
            tryReplyKeywordFromSheet4_(replyToken, text, ev)
          ) {
            eventMeta.keywordAutoReply = true;
            return;
          }
        } catch (keywordErr) {
          console.error(
            "KEYWORD_AUTO_REPLY_FAILED: " +
            String(keywordErr && keywordErr.stack ? keywordErr.stack : keywordErr)
          );
        }

        const userKey =
          getLineUserKey(ev);

        const lottoData =
          loadLotto();


        // =====================================================
        // MULTI-MESSAGE BUFFER
        //
        // รองรับลูกค้าส่งโพยเป็นหลายข้อความ
        //
        // 1) เลข + บนล่างกลับ
        // 2) 5x5
        // 3) นิวี
        // =====================================================
        const hadBuffer =
          hasBetBuffer_(ev);

        const currentHasLotto =
          hasAnyLottoReference_(
            text,
            lottoData
          );

        const currentLooksLikePart =
          looksLikeBetConversation_(
            text
          ) ||
          currentHasLotto;


        const currentIsCompleteBet =
          looksLikeStandaloneCompleteBet_(
            text,
            lottoData
          );


        if (hadBuffer) {

          const bufferedText =
            loadBetBuffer_(ev);


          const normalizedCurrentForBuffer =
            cleanIncomingLineText_(
              text
            );


          const normalizedBufferedForBuffer =
            cleanIncomingLineText_(
              bufferedText
            );


          // =====================================================
          // ลูกค้าส่งโพยเต็มเดิมซ้ำ
          // =====================================================
          const repeatedWholeMessage =

            normalizedCurrentForBuffer &&

            normalizedBufferedForBuffer &&

            normalizedCurrentForBuffer ===
              normalizedBufferedForBuffer &&

            (
              currentHasLotto ||
              normalizedCurrentForBuffer.length >= 30
            );


          if (repeatedWholeMessage) {

            clearBetBuffer_(ev);

            text =
              normalizedCurrentForBuffer;

            eventMeta.bufferDuplicateSuppressed =
              true;


          } else if (
            currentIsCompleteBet
          ) {

            // โพยใหม่เต็มก้อน
            clearBetBuffer_(ev);


          } else if (
            currentLooksLikePart
          ) {

            text =
              appendBetBuffer_(
                ev,
                text
              );

            eventMeta.buffered =
              true;


          } else {

            clearBetBuffer_(ev);
          }


        } else if (
          !currentIsCompleteBet &&
          shouldStartBetBuffer_(
            text,
            lottoData
          )
        ) {

          appendBetBuffer_(
            ev,
            text
          );

          eventMeta.bufferStarted =
            true;

          // LINE อาจยิงหลาย webhook พร้อมกันจนทุกคำขอเห็นว่า Buffer ยังว่าง
          // รอสั้น ๆ ให้ทุก fragment ถูกเติมก่อน แล้วให้เฉพาะ event ล่าสุดทำงานต่อ
          Utilities.sleep(
            BET_BUFFER_DEBOUNCE_MS
          );

          if (
            !isLatestBetBufferEvent_(ev)
          ) {
            return;
          }

          text = loadBetBuffer_(ev);

          eventMeta.bufferDebounced =
            true;
        }


        // =====================================================
        // ชื่อหวยอย่างเดียว = เงียบ
        //
        // เช่น
        // จีน
        // นิวี
        // ลาวพัฒนา
        // =====================================================
        if (
          hasAnyLottoReference_(
            text,
            lottoData
          ) &&
          !hasRealBetPayload_(
            text,
            lottoData
          )
        ) {

          if (
            !hasBetBuffer_(ev)
          ) {

            appendBetBuffer_(
              ev,
              text
            );
          }


          eventMeta.silentLottoNameOnly =
            true;

          return;
        }


        // =====================================================
        // MULTI LOTTO BET
        //
        // เล่นหลายหวยในโพยเดียว
        // =====================================================
        const multiLottoMatches =
          findMultipleLottosInText_(
            text,
            lottoData
          );


        if (
          multiLottoMatches.length >= 2
        ) {

          clearBetBuffer_(ev);


          if (userKey) {
            clearPendingBet(
              userKey
            );
          }


          // =====================================================
          // หวยที่ยังเปิด
          // =====================================================
          const openMulti =
            multiLottoMatches.filter(
              function(m) {

                return (
                  m &&
                  m.name &&
                  m.closeTime &&
                  !isClosed(
                    m.closeTime
                  )
                );
              }
            );


          // =====================================================
          // หวยที่ปิดแล้ว
          // =====================================================
          const closedMulti =
            multiLottoMatches.filter(
              function(m) {

                return (
                  m &&
                  m.name &&
                  m.closeTime &&
                  isClosed(
                    m.closeTime
                  )
                );
              }
            );


          // =====================================================
          // ตัดชื่อหวยออกก่อน parse
          // =====================================================
          let multiCmd =
            stripMultipleLottoNamesFromText_(
              text,
              multiLottoMatches,
              lottoData
            );


          // =====================================================
          // รูปแบบกำหนดราคาทีละบรรทัด
          // =====================================================
          const assignedMulti =
            normalizePerLineAssignedBetFormat_(
              multiCmd,
              ""
            );


          if (assignedMulti) {

            multiCmd =
              assignedMulti;


          } else {

            const legacyMulti =
              normalizeLegacyNaturalBetFormat_(
                multiCmd,
                ""
              );


            if (legacyMulti) {
              multiCmd =
                legacyMulti;
            }
          }


          // =====================================================
          // ถ้ายังต้องถามว่าจะกลับเลขไหม
          // =====================================================
          if (
            shouldAskReturnChoice(
              multiCmd
            )
          ) {

            replyQuickReplyReturnChoice(
              replyToken,
              text
            );

            return;
          }


          const multiKeywords =
            loadKeywords();


          const multiDoorMap =
            loadDoors();


          const multiResult =
            parseMessage(
              multiCmd,
              multiKeywords,
              multiDoorMap
            );


          if (
            !multiResult.success
          ) {

            return;
          }


          const multiGroups =
            multiResult.groups || [];


          let multiTotalAmount = 0;
          let multiTotalBetCount = 0;


          multiGroups.forEach(
            function(g) {

              multiTotalAmount +=

                (
                  typeof g.subtotal ===
                  "number"
                )

                  ? g.subtotal

                  : (
                      g.numbers || []
                    ).length *

                    (
                      (
                        Number(g.top) ||
                        0
                      ) +

                      (
                        Number(g.bottom) ||
                        0
                      )
                    );


              multiTotalBetCount +=

                (
                  typeof g.betCount ===
                  "number"
                )

                  ? g.betCount

                  : (
                      g.numbers || []
                    ).length;
            }
          );


          const multiFlexes = [];

          const multiSaveFailed =
            [];


          // =====================================================
          // บันทึกหวยแต่ละตัวลง Sheet
          // =====================================================
          openMulti.forEach(
            function(m) {

              try {

                saveParsedBetFromLine_(
                  ev,
                  m.name,
                  m.closeTime,
                  text,
                  multiGroups,
                  multiTotalAmount,
                  multiTotalBetCount
                );


                multiFlexes.push(
                  buildFlex({

                    name:
                      m.name,

                    closeTime:
                      m.closeTime,

                    groups:
                      multiGroups,

                    totalDoors:
                      multiTotalBetCount,

                    totalAmount:
                      multiTotalAmount
                  })
                );


              } catch (saveErr) {

                multiSaveFailed.push(
                  m.name
                );


                console.error(
                  "BET_SAVE_FATAL " +
                  JSON.stringify({
                    scope: "multi_lotto",
                    lottoName: String(m && m.name ? m.name : ""),
                    userId: String(ev && ev.source && ev.source.userId ? ev.source.userId : ""),
                    webhookEventId: String(ev && ev.webhookEventId ? ev.webhookEventId : ""),
                    message: String(saveErr && saveErr.message ? saveErr.message : saveErr),
                    stack: String(saveErr && saveErr.stack ? saveErr.stack : "")
                  })
                );
              }
            }
          );


          // =====================================================
          // ส่ง Flex
          // =====================================================
          if (
            multiFlexes.length > 0
          ) {

            replyFlexMany_(
              replyToken,
              multiFlexes
            );


          } else if (
            multiSaveFailed.length > 0
          ) {

            replyText(
              replyToken,
              "⚠️ ระบบบันทึกโพยขัดข้อง กรุณาส่งโพยนี้ใหม่อีกครั้ง"
            );


          } else if (
            closedMulti.length > 0
          ) {

            replyText(
              replyToken,
              "❌ หวยที่เลือกปิดรับแล้วทั้งหมด"
            );
          }


          eventMeta.multiLotto =
            openMulti.map(
              function(m) {
                return m.name;
              }
            );


          eventMeta.multiLottoClosed =
            closedMulti.map(
              function(m) {
                return m.name;
              }
            );


          return;
        }


        // =====================================================
        // SAFE LOTTO MATCH
        // =====================================================
        let matchInText =
          findLottoSafe_(
            text,
            lottoData
          );


        // =====================================================
        // alias กำกวม แต่เหลือเปิดหวยเดียว
        // =====================================================
        if (
          matchInText &&
          matchInText.ambiguous
        ) {

          const onlyOpenCandidates =

            (
              matchInText.candidates ||
              []
            )

            .filter(
              function(c) {

                return (
                  c &&
                  c.name &&
                  c.closeTime &&
                  !isClosed(
                    c.closeTime
                  )
                );
              }
            );


          if (
            onlyOpenCandidates.length === 1
          ) {

            const only =
              onlyOpenCandidates[0];


            matchInText = {

              name:
                only.name,

              closeTime:
                only.closeTime,

              matchedToken:

                (
                  matchInText.matchedTokens &&
                  matchInText.matchedTokens[0]
                ) ||

                only.name,

              autoResolvedSingleOpen:
                true
            };


            eventMeta.autoResolvedSingleOpen =
              true;
          }
        }


        // =====================================================
        // ยังไม่เจอชื่อหวย
        // =====================================================
        if (!matchInText) {

          if (
            hasBetBuffer_(ev)
          ) {
            return;
          }


          const pending =
            userKey
              ? loadPendingBet(
                  userKey
                )
              : null;


          if (
            pending &&
            (
              looksLikeBet(
                text
              ) ||
              looksLikePriceOnly(
                text
              )
            )
          ) {

            const closeTime =

              lottoData.map[
                pending.name
              ] ||

              pending.closeTime;


            if (!closeTime) {

              clearPendingBet(
                userKey
              );


            } else if (
              isClosed(
                closeTime
              )
            ) {

              clearPendingBet(
                userKey
              );

              clearBetBuffer_(ev);


              replyText(
                replyToken,
                "❌ " +
                pending.name +
                " ปิดรับแล้ว (" +
                closeTime +
                ")"
              );

              return;


            } else {

              const keywords =
                loadKeywords();


              const doorMap =
                loadDoors();


              const pendingText =

                (
                  pending.cmd ||
                  ""
                ) +

                "\n" +

                text;


              const result =
                parseMessage(
                  pendingText,
                  keywords,
                  doorMap
                );


              if (
                result.success
              ) {

                clearPendingBet(
                  userKey
                );

                clearBetBuffer_(ev);


                try {

                  saveParsedBetFromLine_(
                    ev,
                    pending.name,
                    closeTime,
                    pendingText,
                    result.groups,
                    result.totalAmount,
                    result.totalDoors
                  );


                } catch (saveErr) {

                  console.error(
                  "BET_SAVE_FATAL " +
                  JSON.stringify({
                    scope: "pending",
                    lottoName: String(pending && pending.name ? pending.name : ""),
                    userId: String(ev && ev.source && ev.source.userId ? ev.source.userId : ""),
                    webhookEventId: String(ev && ev.webhookEventId ? ev.webhookEventId : ""),
                    message: String(saveErr && saveErr.message ? saveErr.message : saveErr),
                    stack: String(saveErr && saveErr.stack ? saveErr.stack : "")
                  })
                );


                  replyText(
                    replyToken,
                    "⚠️ ระบบบันทึกโพยขัดข้อง กรุณาส่งโพยนี้ใหม่อีกครั้ง"
                  );

                  return;
                }


                replyBetFlex_(
                  replyToken,

                  buildFlex({

                    name:
                      pending.name,

                    closeTime:
                      closeTime,

                    groups:
                      result.groups,

                    totalDoors:
                      result.totalDoors,

                    totalAmount:
                      result.totalAmount
                  })
                );


                return;
              }


              appendBetBuffer_(
                ev,
                pendingText
              );

              return;
            }
          }


          // =====================================================
          // โพยไม่มีชื่อหวย
          // =====================================================
          if (
            shouldOfferLottoChoiceForBet_(
              text
            )
          ) {

            replyQuickReplyOpenLottoChoice(
              replyToken,
              lottoData,
              text
            );
          }


          return;
        }


        // =====================================================
        // ชื่อหวยกำกวม
        // =====================================================
        if (
          matchInText.ambiguous
        ) {

          if (
            matchInText.allClosed
          ) {

            clearBetBuffer_(ev);


            replyText(
              replyToken,
              "❌ หวยรายการนี้ปิดรับแล้วทั้งหมด"
            );


            return;
          }


          clearBetBuffer_(ev);


          replyQuickReplyLottoChoice(

            replyToken,

            matchInText.candidates,

            text,

            matchInText.matchedTokens ||
              []
          );


          return;
        }


        const name =
          matchInText.name;


        const closeTime =
          matchInText.closeTime;


        const matchedToken =
          matchInText.matchedToken;


        // =====================================================
        // หวยปิดรับ
        // =====================================================
        if (
          isClosed(
            closeTime
          )
        ) {

          clearBetBuffer_(ev);


          if (userKey) {

            clearPendingBet(
              userKey
            );
          }


          replyText(
            replyToken,

            "❌ " +
            name +
            " ปิดรับแล้ว (" +
            closeTime +
            ")"
          );


          return;
        }


        // =====================================================
        // PRE-PARSE CUSTOM FORMAT (ต้องมาก่อน assigned/legacy)
        // =====================================================
        const preParseCustomText =
          normalizeMainPreParseCustomBet_(
            text,
            name
          );

        if (preParseCustomText) {
          text = preParseCustomText;

        } else {
          // =====================================================
          // แปลงโพยกำหนดราคาเป็นรายเลข
          // =====================================================
          const assignedNormalizedText =
            normalizePerLineAssignedBetFormat_(
              text,
              name
            );

          if (assignedNormalizedText) {
            text = assignedNormalizedText;

          } else {
            const legacyNormalizedText =
              normalizeLegacyNaturalBetFormat_(
                text,
                name
              );

            if (legacyNormalizedText) {
              text = legacyNormalizedText;
            }
          }
        }


        // =====================================================
        // ตัดชื่อหวยออกจากโพย
        // =====================================================
        let cmd =
          stripLottoNameFromText(

            text,

            name,

            lottoData,

            matchedToken
          );


        cmd = cmd

          .split(/\n/)

          .map(
            function(line) {

              return line
                .replace(
                  /[ \t]+/g,
                  " "
                )
                .trim();
            }
          )

          .filter(Boolean)

          .join("\n");


        // =====================================================
        // ถามกลับเลข
        // =====================================================
        if (
          shouldAskReturnChoice(
            cmd
          )
        ) {

          clearBetBuffer_(ev);


          replyQuickReplyReturnChoice(
            replyToken,
            text
          );


          return;
        }


        // =====================================================
        // PARSE โพย
        // =====================================================
        const keywords =
          loadKeywords();


        const doorMap =
          loadDoors();


        const result =
          parseMessage(
            cmd,
            keywords,
            doorMap
          );


        // =====================================================
        // PARSE ไม่สำเร็จ
        // =====================================================
        if (
          !result.success
        ) {

          if (
            isNoPriceError(
              result.error
            )
          ) {

            const completeSingleMessage =
              looksLikeStandaloneCompleteBet_(
                text,
                lottoData
              );


            if (
              !completeSingleMessage
            ) {

              if (
                !hasBetBuffer_(ev)
              ) {

                appendBetBuffer_(
                  ev,
                  text
                );
              }


              if (userKey) {

                savePendingBet(
                  userKey,
                  {
                    name:
                      name,

                    closeTime:
                      closeTime,

                    cmd:
                      cmd
                  }
                );
              }


            } else {

              clearBetBuffer_(ev);


              if (userKey) {

                clearPendingBet(
                  userKey
                );
              }
            }
          }


          return;
        }


        // =====================================================
        // PARSE สำเร็จ
        // =====================================================
        if (userKey) {

          clearPendingBet(
            userKey
          );
        }


        clearBetBuffer_(ev);


        const groups =
          result.groups || [];


        let totalAmount = 0;
        let totalBetCount = 0;


        groups.forEach(
          function(g) {

            totalAmount +=

              (
                typeof g.subtotal ===
                "number"
              )

                ? g.subtotal

                : (
                    g.numbers || []
                  ).length *

                  (
                    (
                      Number(
                        g.top
                      ) ||
                      0
                    ) +

                    (
                      Number(
                        g.bottom
                      ) ||
                      0
                    )
                  );


            totalBetCount +=

              (
                typeof g.betCount ===
                  "number"
              )

                ? g.betCount

                : (
                    g.numbers || []
                  ).length;
          }
        );


        // =====================================================
        // บันทึกโพยลง Sheet
        // =====================================================
        try {

          saveParsedBetFromLine_(

            ev,

            name,

            closeTime,

            text,

            groups,

            totalAmount,

            totalBetCount
          );


        } catch (saveErr) {

          console.error(
          "BET_SAVE_FATAL " +
          JSON.stringify({
            scope: "single_lotto",
            lottoName: String(name || ""),
            userId: String(ev && ev.source && ev.source.userId ? ev.source.userId : ""),
            webhookEventId: String(ev && ev.webhookEventId ? ev.webhookEventId : ""),
            message: String(saveErr && saveErr.message ? saveErr.message : saveErr),
            stack: String(saveErr && saveErr.stack ? saveErr.stack : "")
          })
        );


          replyText(
            replyToken,

            "⚠️ ระบบบันทึกโพยขัดข้อง กรุณาส่งโพยนี้ใหม่อีกครั้ง"
          );


          return;
        }


        // =====================================================
        // ส่ง Flex โพย
        // =====================================================
        replyBetFlex_(

          replyToken,

          buildFlex({

            name:
              name,

            closeTime:
              closeTime,

            groups:
              groups,

            totalDoors:
              totalBetCount,

            totalAmount:
              totalAmount
          })
        );


      } catch (eventErr) {

        eventMeta.error = String(
          eventErr && eventErr.message
            ? eventErr.message
            : eventErr
        );

        eventMeta.stack = String(
          eventErr && eventErr.stack
            ? eventErr.stack
            : ""
        );

        console.error(
          "LINE_EVENT_FATAL " +
          JSON.stringify({
            webhookEventId: String(
              ev && ev.webhookEventId
                ? ev.webhookEventId
                : ""
            ),
            eventType: String(
              ev && ev.type
                ? ev.type
                : ""
            ),
            messageType: String(
              ev && ev.message && ev.message.type
                ? ev.message.type
                : ""
            ),
            textHead: String(
              eventMeta.textHead || ""
            ),
            error: eventMeta.error,
            stack: eventMeta.stack
          })
        );

      } finally {

        logBotTiming(
          "event",

          Date.now() -
            eventStartMs,

          eventMeta
        );
      }
    });


    // =====================================================
    // BOT TIMING
    // =====================================================
    logBotTiming(
      "webhook",

      Date.now() -
        webhookStartMs,

      {
        eventCount:
          (
            data.events ||
            []
          ).length
      }
    );


    return ContentService.createTextOutput(
      "OK"
    );


  } catch (err) {

    logBotTiming(
      "error",
      0,
      {
        message:
          String(
            err &&
            err.message
              ? err.message
              : err
          )
      }
    );


    // ตอบ OK เสมอ
    // ป้องกัน LINE redelivery ซ้ำ
    return ContentService.createTextOutput(
      "OK"
    );
  }
}

function isCustomerResultReport_(text) {
  const t = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!t) return false;

  // มีคำว่า "ถูก" ที่ไหนในข้อความก็ถือว่าเป็นการแจ้งผล
  if (/ถูก/i.test(t)) {
    return true;
  }

  return false;
}
// ================== LEGACY / NATURAL BET FORMAT ==================
// รองรับโพยภาษาพูดที่ลูกค้าพิมพ์ประมาณ:
// ลาวพัฒนา3มาตัวบนหก
// กับ=365/368/.../081=1บ/บนล่าง
// กับ
// ด้วย=56/95/.../83/=3บ
//
// ความหมาย:
// - "3มาตัวบนหก" + "กับ" (พิมพ์ตกคำว่า ล) = 3 ตัว 6กลับ
// - "บนล่าง กับ ด้วย" = 2 ตัว บนล่างกลับ
// - เครื่องหมาย = ตัวแรกหลัง "กับ/ด้วย" เป็นตัวเปิดรายการ ไม่ใช่ราคา
// - ราคาจริงคือ =1บ / =3บ ที่ท้ายรายการ
//
// หมายเหตุ:
// รูปแบบ ".../83/=3บ" ตามตัวอย่างอ้างอิงของผู้ใช้
// จะถือว่าเลขหน้าสุดท้ายที่ติด "/=" เป็นเศษก่อนราคาและไม่นำมาคิด
// เพื่อให้ผลตรงบิลอ้างอิง (34 ตัว / 204 บาท)
function normalizeLegacyNaturalBetFormat_(text, lottoName) {
  let t = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯]/g, "*");

  // ทำให้คำที่พิมพ์ตก "ล" อ่านเป็น "กลับ" เฉพาะในบริบท mode
  // "ตัวบนหก กับ" / "ตัวบน6กับ" -> 6กลับ
  const hasSixReturnHint =
    /(?:3|สาม)\s*(?:มา)?\s*ตัวบน\s*(?:หก|6)\s*(?:กลับ|กับ)?/i.test(t) ||
    /ตัวบน\s*(?:หก|6)\s*(?:กลับ|กับ)/i.test(t);

  // "บนล่าง กับ ด้วย" / "บนล่างกลับด้วย"
  const hasTwoReturnHint =
    /บน\s*ล่าง\s*(?:กลับ|กับ)/i.test(t) ||
    /บนล่าง[\s\/]*\n?\s*กับ\s*\n?\s*ด้วย\s*=/i.test(t) ||
    /บนล่าง\s*\/?\s*กับ[\s\S]{0,20}?ด้วย\s*=/i.test(t);

  // ต้องเป็น syntax แบบ กับ=...=ราคา หรือ ด้วย=...=ราคา จริง ๆ
  const blockRe = /(กับ|ด้วย)\s*=\s*([0-9\s\/\\,._-]+?)\s*(\/?)\s*=\s*(\d+(?:\s*\*\s*\d+)?)\s*(?:บ|บาท)?/g;
  const found = [];
  let m;

  while ((m = blockRe.exec(t)) !== null) {
    const label = m[1];
    const rawNumbers = String(m[2] || "");
    const slashBeforePrice = String(m[3] || "") === "/";
    const price = String(m[4] || "").replace(/\s+/g, "");

    let nums = rawNumbers
      .split(/[\/\\,\s._-]+/)
      .map(function(n) { return String(n || "").trim(); })
      .filter(function(n) { return /^\d+$/.test(n); });

    if (nums.length === 0) continue;

    const all3 = nums.every(function(n) { return n.length === 3; });
    const all2 = nums.every(function(n) { return n.length === 2; });

    // ทำตามบิลอ้างอิง: รูปแบบ "ด้วย=.../83/=3บ"
    // ตัวก่อน "/=" เป็นเศษท้ายรายการ จึงตัดออก 1 ตัว
    if (label === "ด้วย" && slashBeforePrice && all2 && nums.length > 1) {
      nums = nums.slice(0, -1);
    }

    if (all3 && hasSixReturnHint) {
      found.push({
        order: m.index,
        text: nums.join(" ") + " 6กลับ " + price
      });
      continue;
    }

    if (all2 && hasTwoReturnHint) {
      // ราคาเดี่ยว 3 = บน 3 + ล่าง 3 ตามกติกา parseDirect2Block เดิม
      found.push({
        order: m.index,
        text: nums.join(" ") + " บลก " + price
      });
      continue;
    }
  }

  // ใช้ตัวแปลงนี้เฉพาะเมื่อจับได้อย่างน้อยหนึ่งบล็อกที่ชัดเจน
  if (found.length === 0) return "";

  found.sort(function(a, b) { return a.order - b.order; });

  const out = [];
  if (lottoName) out.push(String(lottoName).trim());
  found.forEach(function(item) {
    out.push(item.text);
  });

  return out.filter(Boolean).join("\n");
}


// ============================================================================
// PATCH V1115-FIX304-PREPARSE
// ดัก syntax พิเศษ "ก่อน" normalizePerLineAssignedBetFormat_ และ parseMessage()
// เพื่อป้องกัน parser เดิมเอาราคาของบรรทัดท้ายไปทับเลขชุดก่อนหน้า
// ============================================================================
function normalizeMainPreParseCustomBet_(text, lottoName) {
  const raw = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯xX]/g, "*");

  const lines = raw
    .split(/\n/)
    .map(function(line) {
      return String(line || "")
        .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
        .replace(/\u00A0/g, " ")
        .trim();
    })
    .filter(Boolean);

  function parse2List(line) {
    const s = String(line || "").trim();
    if (!/^\d{2}(?:\s*[,\/\\._-]?\s*\d{2})*$/.test(s)) return null;
    const nums = s
      .replace(/[,\/\\._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .filter(Boolean);
    if (!nums.length || !nums.every(function(n){ return /^\d{2}$/.test(n); })) return null;
    return nums;
  }

  function modeCanon(mode) {
    const m = String(mode || "").replace(/\s+/g, "");
    if (m === "บกล" || m === "บนล่างกลับ") return "บลก";
    if (m === "บนล่าง") return "บล";
    if (m === "บ") return "บน";
    if (m === "ล") return "ล่าง";
    return m;
  }

  // ------------------------------------------------------------------------
  // A) รูด/รุด/ลูด + ราคาแยกบรรทัด
  // ------------------------------------------------------------------------
  for (let i = 0; i < lines.length - 1; i++) {
    const rm = lines[i].match(
      /^(?:รูด|รุด|ลูด)\s*([0-9](?:\s*[,\/-]\s*[0-9]){0,6}|[0-9]{2,7})\s*$/
    );
    if (!rm) continue;

    const pm = lines[i + 1]
      .replace(/\s+/g, "")
      .match(/^(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)$/);
    if (!pm) continue;

    let ds = String(rm[1] || "").replace(/\s+/g, "");
    if (/^\d{2,7}$/.test(ds) && !/[\/,\-]/.test(ds)) {
      ds = ds.split("").join("/");
    } else {
      ds = ds.replace(/[,\-]+/g, "/");
    }

    const out = [];
    if (lottoName) out.push(String(lottoName).trim());
    out.push("รูด " + ds + " บล " + pm[1] + "*" + pm[2]);
    return out.join("\n");
  }

  // ------------------------------------------------------------------------
  // B) หลายบรรทัดเลข -> 8บลก -> 11=40บล
  // แปลงเป็น canonical ก่อน parser หลัก
  // ------------------------------------------------------------------------
  const canonical = [];
  let recognized = 0;

  for (let i = 0; i < lines.length; i++) {
    const first = parse2List(lines[i]);
    if (!first) continue;

    let nums = first.slice();
    let j = i + 1;
    while (j < lines.length) {
      const more = parse2List(lines[j]);
      if (!more) break;
      nums.push.apply(nums, more);
      j++;
    }

    if (j >= lines.length) continue;

    const priceMode = lines[j].match(
      /^(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
    );
    if (!priceMode) continue;

    const p = String(priceMode[1]);
    const m = modeCanon(priceMode[2]);
    canonical.push(nums.join(" ") + " " + m + " " + p + "*" + p);
    recognized++;
    i = j;
  }

  // เลขเดี่ยว=ราคา+mode เช่น 11=40บล
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(
      /^(\d{2})\s*=\s*(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
    );
    if (!m) continue;
    const mode = modeCanon(m[3]);
    const p = String(m[2]);
    canonical.push(m[1] + " " + mode + " " + p + "*" + p);
    recognized++;
  }

  if (!recognized) return "";

  const out = [];
  if (lottoName) out.push(String(lottoName).trim());
  out.push.apply(out, canonical);
  return out.filter(Boolean).join("\n");
}

// ================== PER-LINE ASSIGNED PRICE / RETURN ==================
// รองรับโพยที่เขียนราคาและชนิดกลับต่อท้ายเลขทีละบรรทัด เช่น:
//
// บน
// 124=1/6กับ
// 853=1/6กับ
// 803=1/6กับ
// 880=1/3กับ
// บนล่างกับ
// 12=5/5
// 24=10/10
// 14=10/10
//
// ความหมาย:
// - 124=1/6กับ  => 124 ราคา 1 แบบ 6กลับ
// - 880=1/3กับ  => อยู่ในกลุ่ม 6กลับ แต่มีเลขไม่ซ้ำจริง 3 แบบ
// - ภายใต้ "บนล่างกับ/บนล่างกลับ": 12=5/5
//   => กลับเลข 12/21, บน 5 + ล่าง 5
//
// คำว่า "กับ" ถือเป็นการพิมพ์ตกตัว "ล" เฉพาะบริบท /6กับ, /3กับ
// และ header บนล่างกับ เท่านั้น ไม่แก้คำว่า "กับ" ทั้งข้อความแบบกว้าง ๆ
function normalizePerLineAssignedBetFormat_(text, lottoName) {
  const raw = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯]/g, "*");

  const lines = raw
    .split(/\n/)
    .map(function(line) {
      return String(line || "")
        .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
        .replace(/\u00A0/g, " ")
        .trim();
    })
    .filter(Boolean);

  if (lines.length === 0) return "";

  const canonical = [];
  let recognized = 0;
  let twoDigitReturnSection = false;
  let sawAssignedSyntax = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // section header: "บนล่างกับ" = "บนล่างกลับ"
    if (/^(?:บน\s*ล่าง|บนล่าง|บล)\s*(?:กลับ|กับ)$/i.test(line)) {
      twoDigitReturnSection = true;
      continue;
    }

    // header "บน" ก่อนรายการ 3 ตัว ไม่ต้องส่งเข้า parser ต่อ
    if (/^บน$/i.test(line)) {
      twoDigitReturnSection = false;
      continue;
    }

    // 3 ตัว: 124=1/6กับ, 880=1/3กลับ
    // group 1 = เลข, group 2 = ราคา, group 3 = จำนวนกลับ
    let m = line.match(/^(\d{3})\s*=\s*(\d+)\s*[\/\\]\s*([36])\s*(?:กลับ|กับ)\s*$/i);
    if (m) {
      canonical.push(m[1] + " 6กลับ " + m[2]);
      recognized++;
      sawAssignedSyntax = true;
      twoDigitReturnSection = false;
      continue;
    }

    // 2 ตัว: 12=5/5 ภายใต้ "บนล่างกับ"
    // หรืออนุญาต suffix กลับ/กับ ติดท้ายบรรทัดโดยตรง
    m = line.match(/^(\d{2})\s*=\s*(\d+)\s*[\/\\]\s*(\d+)\s*(กลับ|กับ)?\s*$/i);
    if (m && (twoDigitReturnSection || m[4])) {
      canonical.push(m[1] + " บลก " + m[2] + "*" + m[3]);
      recognized++;
      sawAssignedSyntax = true;
      continue;
    }

    // บรรทัดชื่อหวย/ธง/คำบรรยายที่ไม่มีเลข ปล่อยผ่านได้
    if (!/\d/.test(line)) {
      continue;
    }

    // ถ้าเริ่มพบ syntax ชนิดนี้แล้ว แต่มีบรรทัดตัวเลขที่เราไม่รู้จัก
    // ห้ามแปลงเพียงบางส่วน เพราะจะทำให้เลขลูกค้าหาย
    if (sawAssignedSyntax || /=/.test(line)) {
      return "";
    }
  }

  if (recognized === 0) return "";

  const out = [];
  if (lottoName) out.push(String(lottoName).trim());
  out.push.apply(out, canonical);
  return out.filter(Boolean).join("\n");
}

// ===== ตัดชื่อหวย + alias ทุกตัวออกจากข้อความ =====
// ป้องกันเศษคำ เช่น "ฮานอยทีวี" → ถ้าตัดเฉพาะ "ฮานอย" จะเหลือ "ทีวี" ค้าง
function stripLottoNameFromText(text, name, lottoData, matchedToken) {
  let t = String(text || "");

  // =====================================================
  // สำคัญมาก:
  // ต้องลบ "ชื่อหวย + alias ที่รู้จัก" ก่อน matchedToken จาก fuzzy
  //
  // ตัวอย่าง:
  // ลาวกาชาดรูด4,7=1*1บลก
  //
  // fuzzy อาจคืน matchedToken = "ลาวกาชาดรูด"
  // ถ้าลบ matchedToken ก่อน จะทำให้คำว่า "รูด" หาย
  // แล้วระบบจะอ่านเป็น 04 40 07 70 แทน
  //
  // วิธีใหม่:
  // 1) ลบชื่อจริง/alias "ลาวกาชาด" ก่อน -> เหลือ "รูด4,7=1*1บลก"
  // 2) ค่อยลบ fuzzy token เฉพาะกรณีที่ยังเหลืออยู่จริง
  // =====================================================

  const knownTerms = [];

  if (name) knownTerms.push(String(name).trim());

  const aliases = (lottoData && lottoData.aliasMap && lottoData.aliasMap[name]) || [];
  for (const a of aliases) {
    const alias = String(a || "").trim();
    if (alias) knownTerms.push(alias);
  }

  // กันซ้ำ + ยาวก่อน
  const uniqueKnownTerms = [...new Set(knownTerms.filter(Boolean))]
    .sort((a, b) => b.length - a.length);

  // STEP 1: ลบชื่อจริงและ alias ก่อนเสมอ
  for (const w of uniqueKnownTerms) {
    const esc = escapeRegExp(w);
    t = t.replace(new RegExp(esc, "gi"), " ");
  }

  // STEP 2: matchedToken ใช้เฉพาะ fuzzy typo ที่ชื่อ/alias ลบไม่ได้จริง ๆ
  // ถ้าชื่อถูกลบไปแล้ว matchedToken เต็มก้อนจะไม่ match อีก
  if (matchedToken) {
    const fuzzyToken = String(matchedToken).trim();

    if (fuzzyToken) {
      // ป้องกันอีกชั้น: matchedToken ที่มีคำสั่งแทงติดอยู่
      // ห้ามลบทั้งก้อน เพราะจะกินคำสั่ง เช่น รูด/วิน/บลก/6กลับ
      const hasBetCommand = /(?:รูด|วิน|บลก|บกล|บล|บนล่าง|บน|ล่าง|กลับ|6\s*กลับ|3\s*กลับ|ตรง|โต๊ด|เบิ้ล|ตอง|หาม)/i.test(fuzzyToken);

      if (!hasBetCommand) {
        const esc = escapeRegExp(fuzzyToken);
        t = t.replace(new RegExp(esc, "gi"), " ");
      }
    }
  }

  // =====================================================
  // ลบบรรทัดเศษคำที่ไม่ใช่เลขและไม่ใช่คำสั่งแทง
  // =====================================================
  const MODE_KEYS_FOR_HEADER = [
    "บนล่างกลับ","บนล่าง","ล่าง","กลับ","บลก","บกล","บล","บน","บ","ล",
    "ตรงโต๊ด","บนตรงโต๊ด","บต","6กลับ","3กลับ","โต๊ด","ตรง",
    "รูดเบิ้ล","เบิ้ลหน้า","เบิ้ลหลัง","เบิ้ล",
    "รูดบนล่าง","รูดบน","รูดล่าง","รูด",
    "วินสอง","วิน",
    "ตอง","หาม"
  ];

  const lines = t.split(/\n/);
  const kept = [];

  for (const line of lines) {
    const trimmed = line.trim();

    if (!trimmed) {
      kept.push(line);
      continue;
    }

    // มีตัวเลข -> เก็บ
    if (/\d/.test(trimmed)) {
      kept.push(line);
      continue;
    }

    // เป็นคำสั่งแทง -> เก็บ
    const normalizedHeader = normalizeBetKeywordTypos(trimmed);
    const isMode = MODE_KEYS_FOR_HEADER.some(function(k) {
      return normalizedHeader === k || normalizedHeader.includes(k);
    });

    if (isMode) {
      kept.push(line);
    }
  }

  return kept.join("\n");
}


// ทดสอบเคสชื่อหวยติดกับคำสั่งรูด
// ลาวกาชาดรูด4,7=1*1บลก -> รูด 4/7 = 38 รายการ = 76 บาท
function testLottoNameAttachedRood_() {
  const lottoName = "🇱🇦 ลาวกาชาด";
  const lottoData = {
    map: {},
    aliasMap: {}
  };
  lottoData.map[lottoName] = "23:25";
  lottoData.aliasMap[lottoName] = ["ลาวกาชาด"];

  const raw = "ลาวกาชาดรูด4,7=1*1บลก";
  const match = findLottoSafe_(raw, lottoData);
  if (!match) throw new Error("FAIL: หา ลาวกาชาด ไม่เจอ");

  const cmd = stripLottoNameFromText(raw, match.name, lottoData, match.matchedToken);
  if (!/รูด/.test(cmd)) {
    throw new Error("FAIL: คำว่า รูด ถูกลบหาย -> " + cmd);
  }

  const doorMap = {};
  for (let d = 0; d <= 9; d++) {
    doorMap[String(d)] = buildRoodDoorsForDigit_(String(d), {});
  }

  const result = parseMessage(cmd, [], doorMap);
  if (!result.success) {
    throw new Error("FAIL parser: " + result.error + " / cmd=" + cmd);
  }

  if (result.grossDoors !== 38 || result.totalAmount !== 76) {
    throw new Error(
      "FAIL: grossDoors=" + result.grossDoors +
      " total=" + result.totalAmount +
      " / cmd=" + cmd
    );
  }

  return "PASS: ชื่อหวยติดรูด -> 38 ตัว / 76 บาท / cmd=" + cmd;
}

function isTestCommand(text) {
  const t = String(text || "").trim();
  return /^(เทส|ทเส|เมส)$/i.test(t);
}

function getLineUserKey(ev) {
  const src = ev && ev.source ? ev.source : {};
  return src.userId || src.groupId || src.roomId || "";
}

function pendingBetKey(userKey) {
  return "pending_bet_" + userKey;
}

function savePendingBet(userKey, data) {
  if (!userKey) return;
  CacheService.getScriptCache().put(pendingBetKey(userKey), JSON.stringify(data), 10 * 60);
}

function loadPendingBet(userKey) {
  if (!userKey) return null;
  const raw = CacheService.getScriptCache().get(pendingBetKey(userKey));
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    clearPendingBet(userKey);
    return null;
  }
}

function clearPendingBet(userKey) {
  if (!userKey) return;
  CacheService.getScriptCache().remove(pendingBetKey(userKey));
}

function isNoPriceError(errorMsg) {
  return /ไม่พบราคา/.test(errorMsg || "");
}


// ================== MULTI-MESSAGE BET BUFFER ==================
// ใช้รวมโพยที่ลูกค้าส่งเป็นหลาย bubble / หลายข้อความติดกัน
// แยก Buffer ตามห้อง + user เพื่อไม่ให้โพยคนอื่นปนกัน
function cleanIncomingLineText_(text) {
  return String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    // รองรับสัญลักษณ์เงินบาทที่ลูกค้าพิมพ์ติดราคา เช่น 50฿ / 20 ฿
    // แปลงเป็นคำว่า "บาท" เพื่อให้ parser และระบบตรวจราคาใช้กฎเดิมร่วมกัน
    .replace(/฿/g, " บาท ")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function getBetBufferKey_(ev) {
  const src = (ev && ev.source) || {};
  const room = src.groupId || src.roomId || "private";
  const user = src.userId || "unknown";

  // CacheService key มีข้อจำกัดด้านความยาว จึงตัดเผื่อไว้
  return ("bet_buffer_" + room + "_" + user).slice(0, 240);
}

function readBetBufferData_(ev) {
  const key = getBetBufferKey_(ev);
  const raw = CacheService.getScriptCache().get(key);

  if (!raw) return null;

  try {
    const data = JSON.parse(raw);
    if (!data || !Array.isArray(data.fragments) || data.fragments.length === 0) {
      CacheService.getScriptCache().remove(key);
      return null;
    }

    const updatedAt = Number(data.updatedAt || 0);
    if (
      updatedAt &&
      Date.now() - updatedAt > BET_BUFFER_TTL_SECONDS * 1000
    ) {
      CacheService.getScriptCache().remove(key);
      return null;
    }

    return data;
  } catch (err) {
    CacheService.getScriptCache().remove(key);
    return null;
  }
}

function hasBetBuffer_(ev) {
  return !!readBetBufferData_(ev);
}

function loadBetBuffer_(ev) {
  const data = readBetBufferData_(ev);
  if (!data) return "";

  return data.fragments
    .slice()
    .sort(function(a, b) {
      return Number(a.timestamp || 0) - Number(b.timestamp || 0);
    })
    .map(function(item) {
      return String(item.text || "").trim();
    })
    .filter(Boolean)
    .join("\n")
    .trim();
}

// ใช้หลัง debounce เพื่อป้องกัน webhook หลายตัวประมวลผลโพยก้อนเดียวกันซ้ำ
function isLatestBetBufferEvent_(ev) {
  const data = readBetBufferData_(ev);
  if (!data || !data.fragments || data.fragments.length === 0) return false;

  const currentMessageId = String(
    ev && ev.message && ev.message.id
      ? ev.message.id
      : ""
  );

  // ถ้า LINE ไม่ส่ง message id มา ให้ event นี้ทำงานต่อได้เพื่อไม่ทำโพยหาย
  if (!currentMessageId) return true;

  const latest = data.fragments
    .slice()
    .sort(function(a, b) {
      const timeDiff = Number(a.timestamp || 0) - Number(b.timestamp || 0);
      if (timeDiff !== 0) return timeDiff;
      return String(a.messageId || "").localeCompare(String(b.messageId || ""));
    })
    .pop();

  return !!latest && String(latest.messageId || "") === currentMessageId;
}

function appendBetBuffer_(ev, text) {
  text = cleanIncomingLineText_(text);
  if (!text) return loadBetBuffer_(ev);

  let lock = null;
  let locked = false;

  try {
    // กันข้อความ LINE ที่เข้ามาแทบพร้อมกันแล้วแย่งเขียน Cache
    lock = LockService.getScriptLock();
    locked = lock.tryLock(2500);
  } catch (err) {
    locked = false;
  }

  try {
    const key = getBetBufferKey_(ev);
    let data = readBetBufferData_(ev);

    if (!data) {
      data = {
        fragments: [],
        updatedAt: Date.now()
      };
    }

    const messageId = String(
      ev && ev.message && ev.message.id
        ? ev.message.id
        : ""
    );

    const timestamp = Number(
      ev && ev.timestamp
        ? ev.timestamp
        : Date.now()
    );

    // กันข้อความเดียวกันถูกเติมซ้ำ
    const alreadyExists = messageId && data.fragments.some(function(item) {
      return String(item.messageId || "") === messageId;
    });

    if (!alreadyExists) {
      data.fragments.push({
        messageId: messageId,
        timestamp: timestamp,
        text: text
      });
    }

    data.fragments.sort(function(a, b) {
      return Number(a.timestamp || 0) - Number(b.timestamp || 0);
    });

    // กัน Buffer โตผิดปกติ: ตัดข้อความเก่าสุดออกก่อน
    while (data.fragments.length > 1) {
      const combined = data.fragments
        .map(function(item) { return String(item.text || ""); })
        .join("\n");

      if (combined.length <= BET_BUFFER_MAX_LENGTH) break;
      data.fragments.shift();
    }

    if (data.fragments.length === 1) {
      const only = String(data.fragments[0].text || "");
      if (only.length > BET_BUFFER_MAX_LENGTH) {
        data.fragments[0].text = only.slice(-BET_BUFFER_MAX_LENGTH);
      }
    }

    data.updatedAt = Date.now();

    CacheService.getScriptCache().put(
      key,
      JSON.stringify(data),
      BET_BUFFER_TTL_SECONDS
    );

    return data.fragments
      .map(function(item) { return String(item.text || "").trim(); })
      .filter(Boolean)
      .join("\n")
      .trim();

  } finally {
    if (locked && lock) {
      try {
        lock.releaseLock();
      } catch (err) {}
    }
  }
}

function clearBetBuffer_(ev) {
  try {
    CacheService.getScriptCache().remove(getBetBufferKey_(ev));
  } catch (err) {}
}

function looksLikeBetConversation_(text) {
  const t = cleanIncomingLineText_(text);
  if (!t) return false;

  // มีตัวเลข = อาจเป็นเลขแทงหรือราคา
  if (/\d/.test(t)) return true;

  // mode ที่ลูกค้าอาจส่งเป็นข้อความเดี่ยว
  return /(บน|ล่าง|กลับ|บลก|บกล|บล|ตรง|โต๊ด|6กลับ|3กลับ|วิน|รูด|เบิ้ล|ตอง|หาม)/.test(t);
}

function hasAnyLottoReference_(text, lottoData) {
  if (!text || !lottoData) return false;
  return !!findLottoSafe_(text, lottoData);
}



// ============================================================================
// MULTI LOTTO MATCHER
// รองรับ: ลาวพัฒนา + ลาวสามัคคี
//         นกช+สิงวี+นอยพิ+นอยปก+นอยวี
// separator: + - / \ | , & และ
// หมายเหตุ: ตรวจเฉพาะ token ที่มีตัวอักษร เพื่อไม่ชนกับเลข 58/78 หรือ 5*5
// ============================================================================
function findMultipleLottosInText_(text, data) {
  const raw = cleanHiddenText_(text);
  const lines = raw.split(/\n/);
  const found = [];
  const seen = {};

  function addMatch(match) {
    if (!match || match.ambiguous || !match.name || seen[match.name]) return;
    seen[match.name] = true;
    found.push({
      name: match.name,
      closeTime: match.closeTime,
      matchedToken: match.matchedToken || match.name
    });
  }

  lines.forEach(function(line) {
    // ต้องเป็นบรรทัดชื่อหวยหลายตัวจริง ไม่ใช่บรรทัดเลข
    if (!/[ก-๙a-zA-Z]/.test(line)) return;
    if (!/[+\-\/\\|,&]/.test(line) && !/\sและ\s/.test(line)) return;

    const parts = line
      .split(/\s*(?:\+|\-|\/|\\|\||,|&|\sและ\s)\s*/g)
      .map(function(x) { return String(x || "").trim(); })
      .filter(function(x) { return /[ก-๙a-zA-Z]/.test(x); });

    if (parts.length < 2) return;

    parts.forEach(function(part) {
      let m = findLottoSafe_(part, data);

      // alias กำกวม แต่เหลือเปิดรับจริงตัวเดียว => เลือกตัวนั้น
      if (m && m.ambiguous) {
        const open = (m.candidates || []).filter(function(c) {
          return c && c.name && c.closeTime && !isClosed(c.closeTime);
        });
        if (open.length === 1) {
          m = {
            name: open[0].name,
            closeTime: open[0].closeTime,
            matchedToken: part
          };
        }
      }

      addMatch(m);
    });
  });

  return found;
}

function stripMultipleLottoNamesFromText_(text, matches, lottoData) {
  let cmd = String(text || "");

  (matches || []).forEach(function(m) {
    cmd = stripLottoNameFromText(
      cmd,
      m.name,
      lottoData,
      m.matchedToken || m.name
    );
  });

  return cmd
    .split(/\n/)
    .map(function(line) {
      // คำบอกว่าเล่นหลายหวยเป็นคำอธิบาย ไม่ใช่คำสั่งแทง
      return line
        .replace(/เล่น\s*หลาย\s*หวย/gi, " ")
        .replace(/[ \t]+/g, " ")
        .trim();
    })
    .filter(Boolean)
    .join("\n");
}

function buildBankAccountTextMessage_() {
  return {
    type: "text",
    text: [
      "0000000000 SCB",
      "ไทยพาณิชย์",
      "น.ส. ชื่อบัญชี นามสกุลบัญชี"
    ].join("\n")
  };
}

function replyBetFlex_(token, msg) {
  const started = Date.now();
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    headers: { Authorization: "Bearer " + getChannelAccessToken_() },
    contentType: "application/json",
    payload: JSON.stringify({
      replyToken: token,
      messages: [
        msg,
        buildBankAccountTextMessage_()
      ]
    })
  });

  logBotTiming("replyBetFlex", Date.now() - started, {
    status: response.getResponseCode()
  });
}

function replyFlexMany_(token, flexMessages) {
  const msgs = (flexMessages || []).filter(Boolean);
  if (msgs.length === 0) return;

  // LINE reply ได้สูงสุด 5 messages ต่อครั้ง
  // กัน 1 ช่องไว้สำหรับข้อความเลขบัญชี
  let messages;
  if (msgs.length <= 4) {
    messages = msgs.slice();
  } else {
    messages = [{
      type: "flex",
      altText: "โพยหลายหวย " + msgs.length + " รายการ",
      contents: {
        type: "carousel",
        contents: msgs.map(function(m) {
          return m && m.contents ? m.contents : null;
        }).filter(Boolean).slice(0, 12)
      }
    }];
  }

  // ส่งเลขบัญชีเป็นข้อความคู่กับ Flex ทุกครั้ง
  messages.push(buildBankAccountTextMessage_());

  const started = Date.now();
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    headers: { Authorization: "Bearer " + getChannelAccessToken_() },
    contentType: "application/json",
    payload: JSON.stringify({
      replyToken: token,
      messages: messages
    })
  });

  logBotTiming("replyFlexMany", Date.now() - started, {
    status: response.getResponseCode(),
    lottoCount: msgs.length
  });
}

// ============================================================================
// SAFE LOTTO MATCHER
// ป้องกัน alias ซ้ำ / คีย์เวิร์ดชนกันแล้วเลือกหวยผิด
//
// หลักการ:
// 1) รวบรวมชื่อจริง + alias ทุกตัวก่อน ไม่ return ตามลำดับแถวในชีต
// 2) เลือกคำที่ match "ยาวที่สุด" ก่อน เช่น ฮานอยเอ็ต ต้องชนะ ฮานอย
// 3) ถ้าคำเดียวกันผูกกับหลายหวย => ambiguous ให้ Quick Reply เลือก
// 4) เช้า/บ่าย/ปก/VIP ที่ผู้ใช้ระบุชัด ใช้กรอง candidate
// 5) fuzzy ใช้เฉพาะคำยาว >= 4 และ threshold สูง เพื่อลดการเดาผิด
// ============================================================================
function lottoCloseSortMinutes_(closeTime) {
  const normalized = normalize(closeTime);
  if (!normalized) return 999999;

  let minutes = toMin(normalized);

  // เวลา 00:00-02:00 เป็นรอบดึกต่อเนื่องจากวันก่อน
  // จัดให้อยู่หลัง 23:59 เพื่อเรียงเวลาปิดได้ถูกต้อง
  if (minutes <= 120) {
    minutes += 24 * 60;
  }

  return minutes;
}

function findLottoSafe_(text, data) {
  const map = (data && data.map) || {};
  const aliasMap = (data && data.aliasMap) || {};
  const rawText = cleanHiddenText_(text);
  const textNorm = normalizeLottoMatch_(rawText);

  if (!textNorm) return null;

  const names = Object.keys(map);
  const entries = [];
  const entrySeen = {};

  function addEntry(term, lottoName, kind) {
    const rawTerm = cleanHiddenText_(term);
    const key = normalizeLottoMatch_(rawTerm);
    if (!rawTerm || !key || key.length < 2) return;

    const uniqueKey = lottoName + "|" + key + "|" + kind;
    if (entrySeen[uniqueKey]) return;
    entrySeen[uniqueKey] = true;

    entries.push({
      term: rawTerm,
      key: key,
      lottoName: lottoName,
      kind: kind
    });
  }

  names.forEach(function(name) {
    addEntry(name, name, "name");

    (aliasMap[name] || []).forEach(function(alias) {
      addEntry(alias, name, "alias");
    });
  });

  // =====================================================
  // EXACT / NORMALIZED TERM MATCH
  // สำคัญ: หา "ทุกตัว" ก่อน แล้วค่อยเลือกคำที่ยาวที่สุด
  // ไม่ให้ alias ที่อยู่แถวบนในชีตชนะโดยอัตโนมัติ
  // =====================================================
  let hits = entries.filter(function(entry) {
    if (entry.key.length <= 3) {
      return hasShortLottoTermHit_(rawText, entry.term);
    }

    return textNorm.indexOf(entry.key) !== -1;
  });

  if (hits.length > 0) {
    hits.sort(function(a, b) {
      if (b.key.length !== a.key.length) {
        return b.key.length - a.key.length;
      }
      // ความยาวเท่ากัน ให้ชื่อจริงมาก่อน alias แต่ยังไม่เลือกทันที
      if (a.kind === b.kind) return 0;
      return a.kind === "name" ? -1 : 1;
    });

    const strongestLength = hits[0].key.length;
    const strongestHits = hits.filter(function(hit) {
      return hit.key.length === strongestLength;
    });

    let candidateNames = [];
    strongestHits.forEach(function(hit) {
      if (candidateNames.indexOf(hit.lottoName) === -1) {
        candidateNames.push(hit.lottoName);
      }
    });

    candidateNames = filterLottoCandidatesByHint_(candidateNames, rawText);

    if (candidateNames.length === 1) {
      const name = candidateNames[0];
      const matchedHit = strongestHits.find(function(hit) {
        return hit.lottoName === name;
      }) || strongestHits[0];

      return {
        name: name,
        closeTime: map[name],
        matchedToken: matchedHit ? matchedHit.term : name,
        safeMatch: true,
        matchType: matchedHit ? matchedHit.kind : "exact"
      };
    }

    if (candidateNames.length > 1) {
      const matchedTokens = [];

      strongestHits.forEach(function(hit) {
        if (matchedTokens.indexOf(hit.term) === -1) {
          matchedTokens.push(hit.term);
        }
      });

      // =====================================================
      // ALIAS ซ้ำหลายหวย:
      // เลือก "หวยที่ยังเปิดและปิดรับก่อน" อัตโนมัติ
      //
      // ตัวอย่าง:
      // นิปก -> นิเคอิเช้า ก่อนนิเคอิบ่าย
      // จีนวี -> จีนเช้า VIP ก่อนจีนบ่าย VIP
      //
      // ถ้ารอบเช้าปิดแล้ว จะข้ามไปเลือกรอบบ่ายที่ยังเปิด
      // ชื่อหวยที่ผู้ใช้พิมพ์ชัดเจนยังถูกเลือกก่อนมาถึงจุดนี้อยู่แล้ว
      // =====================================================
      const openCandidates = candidateNames
        .map(function(name) {
          return {
            name: name,
            closeTime: map[name]
          };
        })
        .filter(function(c) {
          return c.closeTime && !isClosed(c.closeTime);
        })
        .sort(function(a, b) {
          return lottoCloseSortMinutes_(a.closeTime) -
                 lottoCloseSortMinutes_(b.closeTime);
        });

      if (openCandidates.length > 0) {
        const selected = openCandidates[0];
        const selectedHit = strongestHits.find(function(hit) {
          return hit.lottoName === selected.name;
        });

        return {
          name: selected.name,
          closeTime: selected.closeTime,
          matchedToken: selectedHit ? selectedHit.term : matchedTokens[0],
          safeMatch: true,
          matchType: "alias-close-first",
          aliasCollisionResolved: true
        };
      }

      return {
        ambiguous: true,
        candidates: candidateNames.map(function(name) {
          return {
            name: name,
            closeTime: map[name]
          };
        }),
        matchedTokens: matchedTokens,
        allClosed: true,
        safeMatch: true,
        aliasCollision: true
      };
    }
  }

  // =====================================================
  // FUZZY SAFE
  // ใช้เป็นทางเลือกสุดท้ายเท่านั้น
  // =====================================================
  return findLottoFuzzySafe_(rawText, data, entries);
}

function normalizeLottoMatch_(text) {
  if (!text) return "";

  let t = cleanHiddenText_(text)
    .toLowerCase()
    .replace(/วี\s*ไอ\s*พี/g, "vip")
    .replace(/\s+/g, "");

  // รองรับ alias เช่น นิv / เกาv ให้เทียบกับ vip ได้
  t = t.replace(/v(?!ip)/g, "vip");

  // เก็บเฉพาะอักษรไทย/อังกฤษ/ตัวเลข ตัดธง emoji และเครื่องหมาย
  return t.replace(/[^ก-๙a-z0-9]/gi, "");
}

// alias สั้น 2-3 ตัวอันตรายที่สุด เช่น นพ / ฮช / จช / จบ
// จึงไม่ใช้ includes() กว้าง ๆ แต่ต้องอยู่ต้น/ท้ายคำ หรือชิดคำสั่งแทง/ตัวเลข
function hasShortLottoTermHit_(rawText, term) {
  const hay = cleanHiddenText_(rawText).toLowerCase();
  const needle = cleanHiddenText_(term)
    .toLowerCase()
    .replace(/[^ก-๙a-z0-9]/gi, "");

  if (!needle) return false;

  let startAt = 0;
  while (startAt <= hay.length) {
    const idx = hay.indexOf(needle, startAt);
    if (idx === -1) break;

    const end = idx + needle.length;
    const before = idx > 0 ? hay.charAt(idx - 1) : "";
    const after = hay.substring(end);
    const afterFirst = after.charAt(0);

    // หน้า alias ต้องเป็นขอบข้อความ/space/emoji/เครื่องหมาย/ตัวเลข
    // ไม่ใช่อยู่กลางคำไทยหรืออังกฤษอีกคำ
    const beforeOk =
      idx === 0 ||
      !/[ก-๙a-z]/i.test(before);

    // หลัง alias ยอมให้เป็นขอบข้อความ/เครื่องหมาย/ตัวเลข
    // หรือคำสั่งแทงที่พิมพ์ติดชื่อหวย เช่น จีนรูด6/9...
    const afterOk =
      end === hay.length ||
      !/[ก-๙a-z]/i.test(afterFirst) ||
      /^(?:รูด|วิน(?:สอง|2)?|บลก|บกล|บล|บนล่าง|บน|ล่าง|กลับ|ตรง|โต๊ด|6\s*กลับ|3\s*กลับ|เบิ้ล|ตอง|หาม)/i.test(after);

    if (beforeOk && afterOk) {
      return true;
    }

    startAt = idx + 1;
  }

  return false;
}

function filterLottoCandidatesByHint_(names, rawText) {
  const original = [...new Set(names || [])];
  let result = original.slice();
  const text = cleanHiddenText_(rawText).toLowerCase();
  const textNorm = normalizeLottoMatch_(text);

  function applyFilter(fn) {
    const filtered = result.filter(fn);
    if (filtered.length > 0) {
      result = filtered;
    }
  }

  // ระบุรอบชัดเจน
  if (/เช้า/.test(text)) {
    applyFilter(function(name) {
      return /เช้า/.test(name);
    });
  }

  if (/บ่าย/.test(text)) {
    applyFilter(function(name) {
      return /บ่าย/.test(name);
    });
  }

  // VIP ต้องมีคำว่า VIP / วีไอพี / v ที่ normalize เป็น vip ชัดเจน
  // ไม่ใช้แค่ตัวอักษร "วี" เพราะคำว่า ทีวี ก็มี "วี"
  if (/vip/.test(textNorm)) {
    applyFilter(function(name) {
      return /vip/i.test(name);
    });
  }

  // ปก / ปกติ => ตัด VIP
  if (/ปกติ|ปก/.test(text)) {
    applyFilter(function(name) {
      return !/vip/i.test(name);
    });
  }

  return result.length > 0 ? result : original;
}

function findLottoFuzzySafe_(text, data, prebuiltEntries) {
  const map = (data && data.map) || {};
  const aliasMap = (data && data.aliasMap) || {};
  const entries = prebuiltEntries ? prebuiltEntries.slice() : [];

  if (!prebuiltEntries) {
    Object.keys(map).forEach(function(name) {
      const nameKey = normalizeLottoMatch_(name);
      if (nameKey.length >= 4) {
        entries.push({
          term: name,
          key: nameKey,
          lottoName: name,
          kind: "name"
        });
      }

      (aliasMap[name] || []).forEach(function(alias) {
        const aliasKey = normalizeLottoMatch_(alias);
        if (aliasKey.length >= 4) {
          entries.push({
            term: alias,
            key: aliasKey,
            lottoName: name,
            kind: "alias"
          });
        }
      });
    });
  }

  const candidates = entries.filter(function(entry) {
    return entry.key && entry.key.length >= 4;
  });

  // ตัดเลขออกก่อน เพื่อเอาเฉพาะคำที่น่าจะเป็นชื่อหวย
  const cleaned = cleanHiddenText_(text).replace(/[0-9]+/g, " ");
  const tokens = cleaned
    .split(/[\s\n,.\-_\/=*]+/)
    .map(function(token) {
      return normalizeLottoMatch_(token);
    })
    .filter(function(token) {
      return token.length >= 4;
    });

  if (tokens.length === 0) return null;

  const scoreMap = {};
  const matchedTokenMap = {};

  tokens.forEach(function(token) {
    candidates.forEach(function(cand) {
      const score = fuzzyScore(token, cand.key);

      // ลดจากการเดากว้าง ๆ: ต้องคล้ายกันอย่างน้อย 90%
      if (score < 0.90) return;

      if (
        scoreMap[cand.lottoName] === undefined ||
        score > scoreMap[cand.lottoName]
      ) {
        scoreMap[cand.lottoName] = score;
        matchedTokenMap[cand.lottoName] = cand.term;
      }
    });
  });

  let names = Object.keys(scoreMap);
  if (names.length === 0) return null;

  names = filterLottoCandidatesByHint_(names, text);
  names.sort(function(a, b) {
    return scoreMap[b] - scoreMap[a];
  });

  if (names.length === 1) {
    const name = names[0];

    return {
      name: name,
      closeTime: map[name],
      matchedToken: matchedTokenMap[name],
      safeMatch: true,
      fuzzyMatch: true,
      fuzzyScore: scoreMap[name]
    };
  }

  const first = names[0];
  const second = names[1];

  // fuzzy จะเลือกอัตโนมัติเฉพาะกรณีชนะชัดมากจริง ๆ
  if (
    scoreMap[first] >= 0.94 &&
    scoreMap[first] - scoreMap[second] >= 0.15
  ) {
    return {
      name: first,
      closeTime: map[first],
      matchedToken: matchedTokenMap[first],
      safeMatch: true,
      fuzzyMatch: true,
      fuzzyScore: scoreMap[first]
    };
  }

  // ไม่แน่ใจ = ห้ามเดา ให้เลือก
  return {
    ambiguous: true,
    candidates: names.slice(0, 8).map(function(name) {
      return {
        name: name,
        closeTime: map[name]
      };
    }),
    matchedTokens: names.slice(0, 8)
      .map(function(name) {
        return matchedTokenMap[name];
      })
      .filter(Boolean),
    allClosed: names.every(function(name) {
      return isClosed(map[name]);
    }),
    safeMatch: true,
    fuzzyAmbiguous: true
  };
}

// =====================================================
// ตรวจ alias ซ้ำในชีต3
// กด Run ฟังก์ชันนี้ใน Apps Script แล้วดู Execution log
// จะเห็นคำที่ผูกกับหวยมากกว่า 1 ตัว
// =====================================================
function auditDuplicateLottoAliases_() {
  const data = loadLotto();
  const aliasMap = (data && data.aliasMap) || {};
  const reverse = {};

  Object.keys(aliasMap).forEach(function(name) {
    (aliasMap[name] || []).forEach(function(alias) {
      const key = normalizeLottoMatch_(alias);
      if (!key) return;

      if (!reverse[key]) {
        reverse[key] = {
          aliases: [],
          names: []
        };
      }

      if (reverse[key].aliases.indexOf(alias) === -1) {
        reverse[key].aliases.push(alias);
      }
      if (reverse[key].names.indexOf(name) === -1) {
        reverse[key].names.push(name);
      }
    });
  });

  const collisions = Object.keys(reverse)
    .filter(function(key) {
      return reverse[key].names.length > 1;
    })
    .sort()
    .map(function(key) {
      return {
        keyword: reverse[key].aliases.join(" / "),
        lottos: reverse[key].names.join(" | ")
      };
    });

  console.log(JSON.stringify(collisions, null, 2));
  return collisions;
}

// Regression test ไม่ต้องใช้ Spreadsheet จริง
function testSafeLottoMatcher_() {
  const data = {
    map: {
      "🇯🇵 นิเคอิเช้า VIP": "10:00",
      "🇯🇵 นิเคอิบ่าย-VIP": "13:20",
      "🇨🇳 จีนเช้า": "10:00",
      "🇨🇳 จีนบ่าย": "14:00",
      "🇻🇳 ฮานอย": "18:00",
      "🇻🇳 ฮานอยEXTAR": "19:00",
      "🇱🇦 ลาวกาชาด": "23:25"
    },
    aliasMap: {
      "🇯🇵 นิเคอิเช้า VIP": ["นิวี"],
      "🇯🇵 นิเคอิบ่าย-VIP": ["นิวี"],
      "🇨🇳 จีนเช้า": ["จีน"],
      "🇨🇳 จีนบ่าย": ["จีน"],
      "🇻🇳 ฮานอย": ["นอย"],
      "🇻🇳 ฮานอยEXTAR": ["ฮานอยเอ็ต", "นอยเอก"],
      "🇱🇦 ลาวกาชาด": ["ลาวกา", "ลกช"]
    }
  };

  const a = findLottoSafe_("นิวี\nรูด6/9บล1", data);
  if (!a || !a.ambiguous || a.candidates.length !== 2) {
    throw new Error("FAIL collision นิวี");
  }

  const b = findLottoSafe_("🇯🇵 นิเคอิบ่าย-VIPรูด6/9บล1", data);
  if (!b || b.name !== "🇯🇵 นิเคอิบ่าย-VIP") {
    throw new Error("FAIL exact นิเคอิบ่าย-VIP");
  }

  const c = findLottoSafe_("ฮานอยเอ็ต\n123 โต๊ด 1", data);
  if (!c || c.name !== "🇻🇳 ฮานอยEXTAR") {
    throw new Error("FAIL longest term ฮานอยเอ็ต");
  }

  const d = findLottoSafe_("ลาวกาชาดรูด4,7=1*1บลก", data);
  if (!d || d.name !== "🇱🇦 ลาวกาชาด") {
    throw new Error("FAIL ชื่อหวยติดคำสั่งรูด");
  }

  return "PASS: safe lotto matcher";
}

// =====================================================
// ตรวจว่าข้อความมี "รายการแทงจริง" นอกเหนือจากชื่อหวยหรือไม่
// ใช้สำหรับกันกรณีลูกค้าพิมพ์แค่ชื่อหวย แล้วบอทตอบ Quick Reply
// หรือแจ้งหวยปิดรับโดยไม่จำเป็น
// =====================================================

// ============================================================================
// PATCH 2026-09-16 FIX — ตรวจรูปแบบวินย่อ
//
// 2917504 3*3/2*6
// นอย 2917504 3*3/2*6
//
// ใช้ทั้งระบบตรวจว่าเป็นโพยจริง + complete bet + quick reply
// ============================================================================
function isCompactWinTwoThreeSyntax_(text) {
  const t = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯xX]/g, "*");

  return /\d{3,10}\s+\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?\s*\/\s*\d+(?:\.\d+)?\s*\*\s*6(?:\s|$)/i.test(t);
}



// ============================================================================
// PATCH 2026-09-21 — รูปแบบวินย่อเพิ่มเติม
//
// 7045239 2*6
//   = วิน 3 ตัว จาก 7045239 ราคา 2 บาท แบบ 6กลับ
//
// 780314 บลก 1*1
//   = วิน 2 ตัว จาก 780314 บน/ล่าง/กลับ 1*1
//
// 780314 บต1*1
//   = วิน 3 ตัว จาก 780314 ตรง/โต๊ด 1*1
// ============================================================================
function isCompactWinShortcutSyntax_(text) {
  const t = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯xX]/g, "*");

  return (
    // 7045239 2*6
    /(?:^|\s)\d{4,10}\s+\d+(?:\.\d+)?\s*\*\s*6(?:\s|$)/i.test(t) ||

    // 780314 บลก 1*1
    /(?:^|\s)\d{4,10}\s+(?:บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล)\s*=?\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?(?:\s|$)/i.test(t) ||

    // 780314 บต1*1 / ตรงโต๊ด1*1
    /(?:^|\s)\d{4,10}\s+(?:บต|ตรงโต๊ด|ตรง\s*โต๊ด|บนตรงโต๊ด)\s*=?\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?(?:\s|$)/i.test(t)
  );
}


function isAnyCompactWinSyntax_(text) {
  return (
    isCompactWinTwoThreeSyntax_(text) ||
    isCompactWinShortcutSyntax_(text)
  );
}


function hasRealBetPayload_(text, lottoData) {
  let t = cleanIncomingLineText_(text);
  if (!t) return false;

  // 2917504 3*3/2*6 เป็นโพยจริง แม้เลขต้นทางจะยาวกว่า 3 หลัก
  if (isAnyCompactWinSyntax_(t)) return true;

  // ลบชื่อหวยจริง + alias ทุกตัวก่อนตรวจ
  // ป้องกันเลข/คำที่เป็นส่วนหนึ่งของชื่อหวยถูกมองเป็นโพย
  const toRemove = [];
  const map = (lottoData && lottoData.map) || {};
  const aliasMap = (lottoData && lottoData.aliasMap) || {};

  Object.keys(map).forEach(function(name) {
    if (name) toRemove.push(name);
    (aliasMap[name] || []).forEach(function(alias) {
      if (alias) toRemove.push(alias);
    });
  });

  // ยาวก่อน เพื่อไม่ให้ alias สั้นไปตัดชื่อยาวก่อน
  toRemove
    .filter(Boolean)
    .sort(function(a, b) { return b.length - a.length; })
    .forEach(function(word) {
      t = t.replace(new RegExp(escapeRegExp(word), "gi"), " ");
    });

  t = cleanIncomingLineText_(t);
  if (!t) return false;

  t = normalizeBetKeywordTypos(normalizePriceSeparators(t));

  // มีเลขแทง 2 หรือ 3 ตัวจริง
  const normalizedNumbers = t
    .replace(/[\/\-_.,=]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (/(^|\s)\d{2,3}(?=\s|$)/.test(normalizedNumbers)) {
    return true;
  }

  // คำสั่งสร้างเลขเองที่อาจมีเพียงราคา 1 หลัก
  // เช่น รูด6/9บล1, เบิ้ลหน้า 1, เบิ้ลหลัง 1, หาม 1, ตอง 5
  if (
    /(รูด|เบิ้ลหน้า|เบิ้ลหลัง|เบิ้ล|หาม|ตอง|วิน|6\s*กลับ|6กลับ|3\s*กลับ|3กลับ|โต๊ด|ตรง|บลก|บกล|บนล่าง|กลับ)/.test(t) &&
    /\d/.test(t)
  ) {
    return true;
  }

  // ราคาแบบชัดเจนร่วมกับข้อมูลเลข/โหมด
  if (/\d+\s*[xX*]\s*\d+/.test(t) || /=\s*\d+/.test(t)) {
    return /\d{2,3}|รูด|เบิ้ล|หาม|ตอง|วิน|กลับ|โต๊ด|ตรง|บล/.test(t);
  }

  return false;
}

function hasStrongBetPriceMarker_(text) {
  let t = cleanIncomingLineText_(text);
  if (!t) return false;

  // ทำราคาแบบภาษาพูดให้เป็นรูปมาตรฐานก่อนตรวจว่าโพยครบหรือยัง
  // เช่น:
  // บนล่าง ตัวละ50฿        -> บนล่าง = 50 บาท
  // บนล่าง กลับ ตัวล่ะ20฿  -> บนล่าง กลับ = 20 บาท
  // สำคัญมาก: ถ้าไม่แปลงตรงนี้ ระบบ Buffer จะคิดว่ายัง "ไม่มีราคา"
  // แล้วบอทจะเงียบทั้งที่โพยครบแล้ว
  t = normalizeBetKeywordTypos(t);
  t = normalizePriceSeparators(t);
  t = stripPriceWords(t);

  // คง newline ไว้ เพราะลูกค้ามักใส่ราคาในแต่ละบรรทัด
  t = t
    .replace(/[×✕✖⨯]/g, "*")
    .replace(/[ \t]+/g, " ")
    .trim();

  // =2 / =2*2 อยู่ตรงไหนของข้อความก็ถือว่ามีราคา
  if (/=\s*\d+(?:\s*\*\s*\d+)?\s*(?:บาท|บ\.?|฿)?/i.test(t)) return true;

  // 5*5 / 5x5
  if (/\d+\s*[xX*]\s*\d+\s*(?:บาท|บ\.?|฿)?/i.test(t)) return true;

  const modePriceRe = /(บนล่างกลับ|บน\s*ล่าง\s*กลับ|บนล่าง|บน\s*ล่าง|บลก|บกล|บล|กลับ|บน|ล่าง|ตรงโต๊ด|ตรง\s*โต๊ด|บต|โต๊ด(?:อย่างเดียว)?|ตรง|6\s*กลับ|6กลับ|3\s*กลับ|3กลับ|วินสอง|วิน|รูดเบิ้ล|เบิ้ลหน้า|เบิ้ลหลัง|เบิ้ล|ตอง|หาม)\s*=?\s*\d+(?:\s*\*\s*\d+)?\s*(?:บาท|บ\.?|฿)?\s*$/i;

  // ตรวจทีละบรรทัด เพื่อรองรับรูปพิมพ์ติด เช่น รูด6/9บล1
  const lines = t.split(/\n/)
    .map(function(line) { return line.trim(); })
    .filter(Boolean);

  for (const line of lines) {
    if (modePriceRe.test(line)) return true;
  }

  return false;
}
function looksLikeStandaloneCompleteBet_(text, lottoData) {
  const t = cleanIncomingLineText_(text);
  if (!t) return false;
  if (!hasAnyLottoReference_(t, lottoData)) return false;

  // ชื่อหวย + 2917504 3*3/2*6 = โพยครบในข้อความเดียว
  // ห้ามส่งเข้า Buffer แล้วเงียบ
  if (isAnyCompactWinSyntax_(t)) return true;

  if (!hasStrongBetPriceMarker_(t)) return false;

  const normalized = normalizeBetKeywordTypos(normalizePriceSeparators(t))
    .replace(/[\/\-_.,=]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // มีเลขแทง 2/3 ตัวจริงในข้อความ
  if (/(^|\s)\d{2,3}(?=\s|$)/.test(normalized)) return true;

  // หรือเป็นคำสั่งที่สร้างเลขเอง เช่น รูด/เบิ้ล/ตอง/หาม
  if (/(รูด(?:บนล่าง|บน|ล่าง)?\s*\d|รูดเบิ้ล|เบิ้ลหน้า|เบิ้ลหลัง|(^|\s)เบิ้ล(?=\s|$)|ตอง|หาม)/.test(normalized)) {
    return true;
  }

  return false;
}

function isStandalonePriceFragment_(text) {
  let t = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯]/g, "*")
    .trim();

  return /^=?\s*\d+\s*(?:\*\s*\d+)?\s*(?:บาท|บ\.?|฿)?\s*(?:กลับ)?\s*$/i.test(t) &&
    (/=|\*/.test(t) || /^\d{1,2}\s*(?:บาท|บ\.?|฿)$/i.test(t));
}

function shouldStartBetBuffer_(text, lottoData) {
  const t = cleanIncomingLineText_(text);
  if (!t) return false;

  // ถ้าข้อความก้อนเดียวมีชื่อหวย + เลข/คำสั่ง + ราคาแล้ว
  // ถือว่าเป็นโพยสมบูรณ์ ห้ามเก็บเข้า Buffer
  // ไม่เช่นนั้นลูกค้าต้องส่งซ้ำถึงจะทำงาน และโพยอาจถูกต่อซ้ำ
  if (looksLikeStandaloneCompleteBet_(t, lottoData)) {
    return false;
  }

  const hasLotto = hasAnyLottoReference_(t, lottoData);
  const looksPart = looksLikeBetConversation_(t);

  // ชื่อหวยอย่างเดียว: เก็บเงียบ ๆ ได้เพื่อรองรับการส่งหลายข้อความ
  // แต่ doPost จะไม่ตอบ Quick Reply / หวยปิด จนกว่าจะมีรายการแทงจริง
  if (hasLotto && !hasRealBetPayload_(t, lottoData)) {
    return true;
  }

  // ชื่อหวย + โพยที่ยังไม่มีราคา
  if (hasLotto && !hasStrongBetPriceMarker_(t)) {
    return true;
  }

  // ราคาเดี่ยว ๆ เช่น 5x5 อาจเป็นข้อความที่ 2 ของโพย
  if (isStandalonePriceFragment_(t)) {
    return true;
  }

  // เลข/mode มาแล้ว แต่ยังไม่มีราคาชัดเจน -> รอข้อความต่อไป
  if (looksPart && !hasStrongBetPriceMarker_(t)) {
    return true;
  }

  return false;
}

// ================== BET DETECTION ==================
function looksLikeBet(text) {
  const t = cleanIncomingLineText_(text);
  if (!t) return false;

  if (isAnyCompactWinSyntax_(t)) return true;

  const KEYS = [
    "บลก","บกล","บล","บน","ล่าง","กลับ",
    "ตรงโต๊ด","บต","6กลับ","3กลับ","โต๊ด","ตรง",
    "วินสอง","วิน",
    "รูดเบิ้ล","เบิ้ลหน้า","เบิ้ลหลัง","เบิ้ล","รูด","หาม","ตอง"
  ];

  for (const k of KEYS) {
    if (t.includes(k)) return true;
  }

  // คำย่อ "ร" หมายถึงรูด เฉพาะเมื่ออยู่หน้าเลขเท่านั้น
  // เช่น ร7 / ร 7 / ร6/9
  // ห้ามใช้ text.includes("ร") เพราะข้อความทั่วไป เช่น "รบกวน..."
  // จะถูกมองว่าเป็นโพยหวยทันที
  if (/(^|[^ก-๙a-zA-Z])ร\s*\d/.test(t)) return true;

  if (/\d+\s*[\*xX×]\s*\d+/.test(t)) return true;
  if (/\d+\s*=\s*\d+/.test(t)) return true;
  if (/(^|\s|\n)[36]\s*[\*xX×ก]\s*\d+/.test(t)) return true;

  const nums = t.match(/\b\d{2,3}\b/g);
  if (nums && nums.length >= 3) return true;

  return false;
}

// =====================================================
// ใช้ตัดสินใจว่า "ควรขึ้น Quick Reply ให้เลือกหวยหรือไม่"
// ต้องเป็นข้อความแทงหวยจริงเท่านั้น
// ข้อความสนทนาทั่วไปต้องเงียบ
// =====================================================
function shouldOfferLottoChoiceForBet_(text) {
  let t = cleanIncomingLineText_(text);
  if (!t) return false;

  // ส่ง 2917504 3*3/2*6 โดยไม่ใส่ชื่อหวย -> ให้เลือกหวย
  if (isAnyCompactWinSyntax_(t)) return true;

  t = normalizeBetKeywordTypos(normalizePriceSeparators(t));

  // ต้องมีตัวเลขอย่างน้อย 1 ตัวก่อนเสมอ
  // ป้องกันคำทั่วไปที่บังเอิญมีคำว่า บน/ตรง/กลับ ฯลฯ
  if (!/\d/.test(t)) return false;

  // มี mode แทงชัดเจน + มีเลข
  if (/(รูด|วิน(?:สอง|2)?|บลก|บกล|บนล่างกลับ|บนล่าง|บล|บน|ล่าง|กลับ|ตรงโต๊ด|บต|โต๊ด(?:อย่างเดียว)?|ตรง|6\s*กลับ|6กลับ|3\s*กลับ|3กลับ|เบิ้ลหน้า|เบิ้ลหลัง|เบิ้ล|หาม|ตอง)/.test(t)) {
    return true;
  }

  // รูดแบบย่อ ร7 / ร6/9
  if (/(^|[^ก-๙a-zA-Z])ร\s*\d/.test(t)) return true;

  // มีราคาแบบชัดเจน และมีเลข 2/3 ตัวอย่างน้อยหนึ่งเลข
  const hasPrice =
    /\d+\s*[\*xX×]\s*\d+/.test(t) ||
    /=\s*\d+(?:\s*[\*xX×]\s*\d+)?/.test(t);

  const normalizedNums = t
    .replace(/[\/\-_.,=]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (hasPrice && /(^|\s)\d{2,3}(?=\s|$)/.test(normalizedNums)) {
    return true;
  }

  // เลข 2/3 ตัวหลายตัวติดกัน มีลักษณะเป็นโพย แม้ยังไม่ใส่ mode
  const nums = normalizedNums.match(/(^|\s)\d{2,3}(?=\s|$)/g) || [];
  if (nums.length >= 3) return true;

  return false;
}

function testNormalChatMustBeSilent_() {
  const normal = "รบกวนแอดโอนมาหั้ยหน่อยค่ะ";
  if (looksLikeBet(normal)) {
    throw new Error("FAIL: ข้อความทั่วไปถูกมองว่าเป็นโพย");
  }
  if (shouldOfferLottoChoiceForBet_(normal)) {
    throw new Error("FAIL: ข้อความทั่วไปจะขึ้น Quick Reply");
  }

  const roodShort = "ร6/9บล1";
  if (!looksLikeBet(roodShort) || !shouldOfferLottoChoiceForBet_(roodShort)) {
    throw new Error("FAIL: ร6/9บล1 ต้องยังอ่านเป็นโพยรูด");
  }

  const bet = "รูด6/9บล1";
  if (!shouldOfferLottoChoiceForBet_(bet)) {
    throw new Error("FAIL: รูด6/9บล1 ต้องขึ้นเลือกหวยเมื่อไม่มีชื่อหวย");
  }

  return "PASS: ข้อความทั่วไปเงียบ / โพยจริงยังทำงาน";
}

function looksLikePriceOnly(text) {
  return /^\s*\d+(?:\s*[\*xX×]\s*\d+)?\s*(?:บาท|บ\.?|฿)?\s*$/.test(cleanIncomingLineText_(text));
}

// =====================================================
// เรียงเวลาปิดรับสำหรับแสดง Quick Reply
//
// 23:50 -> 1430
// 00:10 -> 1450
// 00:30 -> 1470
//
// ทำให้หวยหลังเที่ยงคืนไปอยู่ท้ายรายการ
// =====================================================
function lottoDisplaySortMin_(time) {
  const min = toMin(time);

  // หวยที่ปิดหลังเที่ยงคืนถึง 02:00
  // ถือว่าเป็นวันต่อเนื่องจากวันนี้
  if (min <= 120) {
    return min + (24 * 60);
  }

  return min;
}
function getOpenLottoNames(lottoData) {
  const map = (lottoData && lottoData.map) || {};

  return Object.keys(map)
    .filter(function(name) {
      return !isClosed(map[name]);
    })
    .sort(function(a, b) {
      return lottoDisplaySortMin_(map[a]) -
             lottoDisplaySortMin_(map[b]);
    })
    .map(function(name) {
      return name + " (ปิด " + map[name] + ")";
    });
}


// =====================================================
// LIFF ROUND / หวยเช้า
// - หวยทั่วไป: ใช้กติกาเปิด/ปิดเดิม
// - หวยที่ชื่อมีคำว่า "เช้า": เมื่อเลยเวลาปิดของวันนี้
//   ให้หน้าเว็บเปิดรับเป็น "รอบวันพรุ่งนี้" อัตโนมัติ
// - ส่ง roundDate กลับไปให้หน้าเว็บและใช้วันเดียวกันตอนบันทึกโพย
// =====================================================
function getLiffLottoRoundInfo_(lottoName, closeTime, nowObj) {
  const now = new Date(nowObj || new Date());
  const normalizedClose = normalize(closeTime);

  if (!normalizedClose) {
    return {
      open: false,
      roundDate: Utilities.formatDate(now, TZ, "yyyy-MM-dd")
    };
  }

  const nowMin = toMin(Utilities.formatDate(now, TZ, "HH:mm"));
  const closeMin = toMin(normalizedClose);
  const isMorning = /เช้า/i.test(String(lottoName || ""));

  // หวยเช้า: หลังเวลาปิดวันนี้ ให้ถือว่าเว็บกำลังรับรอบพรุ่งนี้
  if (isMorning && nowMin > closeMin) {
    const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    return {
      open: true,
      roundDate: Utilities.formatDate(tomorrow, TZ, "yyyy-MM-dd"),
      rolledToTomorrow: true
    };
  }

  return {
    open: !isClosed(normalizedClose),
    roundDate: rsGetRoundDate_(now, normalizedClose),
    rolledToTomorrow: false
  };
}

function getOpenLottoCandidates(lottoData) {
  const map = (lottoData && lottoData.map) || {};
  const now = new Date();

  return Object.keys(map)
    .map(function(name) {
      const closeTime = map[name];
      const roundInfo = getLiffLottoRoundInfo_(name, closeTime, now);

      return {
        name: name,
        closeTime: closeTime,
        roundDate: roundInfo.roundDate,
        rolledToTomorrow: !!roundInfo.rolledToTomorrow,
        open: !!roundInfo.open
      };
    })
    .filter(function(item) {
      return item.open;
    })
    .sort(function(a, b) {
      // รอบวันนี้ขึ้นก่อนรอบพรุ่งนี้ แล้วค่อยเรียงตามเวลาปิด
      if (a.roundDate !== b.roundDate) {
        return a.roundDate < b.roundDate ? -1 : 1;
      }
      return lottoDisplaySortMin_(a.closeTime) -
             lottoDisplaySortMin_(b.closeTime);
    })
    .map(function(item) {
      return {
        name: item.name,
        closeTime: item.closeTime,
        roundDate: item.roundDate,
        rolledToTomorrow: item.rolledToTomorrow
      };
    });
}

// ================== PRE-NORMALIZE WHOLE MESSAGE ==================
function stripPriceWords(text) {
  const STRIP_WORDS = [
    "ตัวล่ะ","ตัวละ","บาทล่ะ","บาทละ","ราคา","ราคาตัวล่ะ","ราคาตัวละ",
    "ตัวหล่ะ","บาทหล่ะ","ราคาตัวหล่ะ",
    "ละ","ล่ะ","หล่ะ","อันละ","อันล่ะ","อันหล่ะ","ชุดละ","ชุดล่ะ","ชุดหล่ะ"
  ].sort((a,b)=>b.length-a.length);

  let t = String(text || "")
    // รองรับ 50฿, 50 ฿ และกรณี parseMessage ถูกเรียกตรงโดยไม่ผ่าน cleanIncomingLineText_()
    .replace(/฿/g, " บาท ");

  for (const w of STRIP_WORDS) {
    const esc = w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(esc + "\\s*(?=\\d)", "g");
    t = t.replace(re, " = ");
  }

  // ตัดหน่วยราคาออกหลังจากแปลง "ตัวละ/ตัวล่ะ" เป็น =
  // 50 บาท -> 50
  // 50 บ.   -> 50
  t = t
    .replace(/(\d)\s*(?:บาท|บ\.)(?=\s|$|\n|[.])/g, "$1 ")
    .replace(/[ \t]+/g, " ");

  return t;
}
function normalizeStandaloneGapAs3Return_(text) {
  const lines = String(text || "").split(/\n/);
  let hasPending3Digit = false;

  return lines.map(function(rawLine) {
    let line = String(rawLine || "");

    // จำว่าก่อนหน้านี้มีชุดเลข 3 ตัวรอ mode อยู่
    if (/(^|[^0-9])\d{3}(?=$|[^0-9])/.test(line)) {
      hasPending3Digit = true;
    }

    // "กับ" เป็นคำพิมพ์ตกของ "กลับ"
    // ถ้าอยู่ในบล็อกเลข 3 ตัว ให้หมายถึง 3 ตัวกลับ
    // ใช้ 6กลับ เป็นโหมดกลาง เพราะ perms3() จะคืน 3 ตัวเองเมื่อเลขซ้ำ
    if (
      hasPending3Digit &&
      /(^|[^ก-๙a-zA-Z])กับ(?=$|[^ก-๙a-zA-Z])/.test(line) &&
      !/(บน\s*ล่าง|บล|2\s*ตัว)/.test(line)
    ) {
      line = line.replace(
        /(^|[^ก-๙a-zA-Z])กับ(?=$|[^ก-๙a-zA-Z])/g,
        "$16กลับ"
      );
      hasPending3Digit = false;
    }

    // เจอโหมด 2 ตัวชัดเจน ให้จบสถานะ 3 ตัวค้าง
    if (/(บลก|บกล|บน\s*ล่าง|บน|ล่าง)/.test(line)) {
      hasPending3Digit = false;
    }

    return line;
  }).join("\n");
}

function normalizeBetKeywordTypos(text) {
  // ลบอักขระซ่อนที่มักติดมากับ LINE / copy-paste
  // เช่น "โต๊ด\u200Bอย่างเดียว2" -> "โต๊ดอย่างเดียว2"
  text = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ");

  text = normalizeStandaloneGapAs3Return_(text);

  // คำพิมพ์ใกล้เคียงของ "รูด"
  // รองรับทั้งแบบเว้นวรรคและติดเลข เช่น รุด0, ลูด 1/7
  text = text
    .replace(/(^|[^ก-๙a-zA-Z])รุด(?=$|[^ก-๙a-zA-Z])/g, "$1รูด")
    .replace(/(^|[^ก-๙a-zA-Z])ลูด(?=$|[^ก-๙a-zA-Z])/g, "$1รูด");

  const replacements = {
    "รุด": "รูด",
    "ลูด": "รูด",

    "บลค": "บลก",
    "บลฅ": "บลก",
    "บรก": "บลก",
    "บงก": "บลก",
    "บลว": "บลก",
    "บคก": "บลก",
    "บลกบ": "บลก",
    "บกล": "บลก",
    "บนล่างกัลบ": "บนล่างกลับ",
    "บนล่างกบับ": "บนล่างกลับ",
    "กัลบ": "กลับ",
    "กบับ": "กลับ",
    "รูดบนล่ง": "รูดบนล่าง",
    "รูดบนร่าง": "รูดบนล่าง"
  };

  let t = text;
  t = t
    // =====================================================
    // ราคาอยู่หน้าคำสั่ง 3 ตัวกลับ
    //
    // 1บ6ก       => 6กลับ 1
    // 1บ6กลับ    => 6กลับ 1
    // 1บาท6ก     => 6กลับ 1
    // 2บ3ก       => 3กลับ 2
    //
    // สำคัญ: ต้องแปลงก่อน shorthand อื่น เพื่อไม่ให้ "บ"
    // ถูกตีความเป็นคำสั่ง "บน"
    // =====================================================
    .replace(/(^|[^0-9])([0-9]+(?:\.[0-9]+)?)\s*(?:บาท|บ\.?)\s*6\s*(?:ก|กลับ)(?=$|[^ก-๙a-zA-Z0-9])/g, "$16กลับ $2")
    .replace(/(^|[^0-9])([0-9]+(?:\.[0-9]+)?)\s*(?:บาท|บ\.?)\s*3\s*(?:ก|กลับ)(?=$|[^ก-๙a-zA-Z0-9])/g, "$13กลับ $2")

    // รองรับแบบติดกันท้ายบรรทัด เช่น 1บ6ก / 1บ6กลับ
    .replace(/^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:บาท|บ\.?)\s*6\s*(?:ก|กลับ)\s*$/g, "6กลับ $1")
    .replace(/^\s*([0-9]+(?:\.[0-9]+)?)\s*(?:บาท|บ\.?)\s*3\s*(?:ก|กลับ)\s*$/g, "3กลับ $1")

    // =====================================================
    // คำเรียก 6กลับ / 3กลับ แบบภาษาพูดหรือพิมพ์คลาดเคลื่อน
    // หกกลับ / หก กลับ / หกกับ / หก กับ -> 6กลับ
    // สามกลับ / สาม กลับ / สามกับ / สาม กับ -> 3กลับ
    // =====================================================
    .replace(/หก\s*กลับ/g, "6กลับ")
    .replace(/หก\s*กับ/g, "6กลับ")
    .replace(/สาม\s*กลับ/g, "3กลับ")
    .replace(/สาม\s*กับ/g, "3กลับ")
    // shorthand ตัวเลข: 6ก / 3ก
    .replace(/(^|[^0-9])6\s*ก(?=$|[^ก-๙a-zA-Z0-9])/g, "$16กลับ")
    .replace(/(^|[^0-9])3\s*ก(?=$|[^ก-๙a-zA-Z0-9])/g, "$13กลับ")

    // =====================================================
    // บน-ล่าง shorthand
    // บ-ล / บ/ล / บ ล / บ.ล. / บน-ล่าง / บน/ล่าง
    // ให้ถือเป็น "บนล่าง" เสมอ
    // สำคัญสำหรับโพยที่เขียนราคา 10*10 = บน 10 + ล่าง 10
    // =====================================================
    .replace(/(^|[^ก-๙a-zA-Z])บ\s*[-\/\.\\]\s*ล(?=$|[^ก-๙a-zA-Z])/g, "$1บนล่าง")
    .replace(/(^|[^ก-๙a-zA-Z])บ\s+ล(?=$|[^ก-๙a-zA-Z])/g, "$1บนล่าง")
    .replace(/บน\s*[-\/]\s*ล่าง/g, "บนล่าง")

    // shorthand รูด เช่น
    // รูดเบิ้ล/บล1  -> รูดเบิ้ล บล 1
    // รูด0บล1       -> รูด 0 บล 1
    // รูด0/บล1      -> รูด 0 บล 1
    .replace(/รูด\s*เบิ้ล\s*[\/\\]?\s*(บนล่างกลับ|บนล่าง|บลก|บกล|บล|บน|ล่าง|บ|ล)\s*=?\s*(\d+(?:\s*\*\s*\d+)?)/g, "รูดเบิ้ล $1 $2")
    .replace(/รูด\s*(\d)\s*[\/\\]?\s*(บนล่างกลับ|บนล่าง|บลก|บกล|บล|บน|ล่าง|บ|ล)\s*=?\s*(\d+(?:\s*\*\s*\d+)?)/g, "รูด $1 $2 $3")
    // "โต๊ดอย่างเดียว2" / "โต๊ดอย่างเดียว 2" / "โต๊ดอย่างเดียว=2"
    // ให้มีความหมายเท่ากับ "โต๊ด 2" / "โต๊ด =2"
    .replace(/โต๊ด\s*(?:อย่างเดียว|อย่างเดี่ยว|อย่างเดี๋ยว)/g, "โต๊ด ")
    .replace(/(^|\\s)(บนล่างกลับ|บลก|บกล|บนล่าง|บล|บน|ล่าง|บ|ล)\\s*=?\\s*(\\d+(?:\\s*\\*\\s*\\d+)?)\\s*[\\n\\r ]*วิน\\s*([0-9]{2,10})(?=\\s|$)/g, "$1วิน $4 $2 $3")
    .replace(/วิน\\s*(?:สอง|2)/g, "วิน")
    .replace(/บ\s*\.\s*ล\s*\.\s*ก\.?/g, "บลก")
    .replace(/บน\s*[-\/]\s*ล่าง/g, "บนล่าง")
    .replace(/(^|[^ก-๙])ร\s*(?=\d)/g, "$1รูด ");
  Object.keys(replacements)
    .sort((a, b) => b.length - a.length)
    .forEach(w => {
      const esc = w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      t = t.replace(new RegExp(esc, "g"), replacements[w]);
    });
  return t;
}

function splitInlineModeBlocks(text) {
  return String(text || "")
    .split(/\n/)
    .map(line => line
      .replace(/\s+(?=รูดเบิ้ล\s*=)/g, "\n")
      .replace(/\s+(?=เบิ้ล(?:หน้า|หลัง)?(?:\s+\d+)?\s*(?:บลก|บกล|บล|บนล่างกลับ|บนล่าง|บน|ล่าง|บ|ล)?\s*=)/g, "\n")
      .replace(/\s+(?=เบิ้ล(?:หน้า|หลัง)?\s+\d+\s*(?:บลก|บกล|บล|บนล่างกลับ|บนล่าง|บน|ล่าง|บ|ล)(?:\s|$))/g, "\n")
      .replace(/\s+(?=ตอง\s*=)/g, "\n")
      .replace(/\s+(?=รูด\s+\d\s*=)/g, "\n")
      .replace(/\s+(?=รูดบน\s+\d\s*=)/g, "\n")
      .replace(/\s+(?=รูดล่าง\s+\d\s*=)/g, "\n")
      .replace(/\s+(?=รูดบนล่าง\s+\d\s*=)/g, "\n")
    )
    .join("\n");
}

function hasAlreadyReturnedTwoDigitNumbers_(text) {
  const lines = String(text || "").split(/\n/);
  const betNums = [];

  lines.forEach(function(rawLine) {
    // เอาเฉพาะฝั่งซ้ายของราคา เช่น
    // 20 02 = 20*20  => ตรวจแค่ 20,02 ไม่เอา 20*20 มาปน
    let line = String(rawLine || "");
    if (line.indexOf("=") !== -1) line = line.split("=")[0];

    line = line
      .replace(/[,]+/g, " ")
      .replace(/(\d)\s*[\-\_\/]+\s*(\d)/g, "$1 $2");

    const nums = line.match(/(^|\D)(\d{1,2})(?=\D|$)/g) || [];
    nums.forEach(function(part) {
      const m = String(part).match(/(\d{1,2})/);
      if (!m) return;
      betNums.push(m[1].padStart(2, "0"));
    });
  });

  if (betNums.length < 2) return false;

  const set = {};
  betNums.forEach(function(n) { set[n] = true; });

  // ถือว่า "ลูกค้ากลับเลขมาแล้ว" เมื่อเลขที่ไม่ใช่เบิ้ลทุกตัว
  // มีเลขกลับของตัวเองอยู่ในโพย เช่น 20/02, 17/71, 39/93
  return betNums.every(function(n) {
    if (n[0] === n[1]) return true;
    const rev = n[1] + n[0];
    return !!set[rev];
  });
}

function shouldAskReturnChoice(text) {
  let t = normalizeBetKeywordTypos(text);
  t = normalizePriceSeparators(t);
  t = stripPriceWords(t);

  // มีคำสั่งโหมดชัดเจนอยู่แล้ว ไม่ต้องถาม
  if (/(บลก|บกล|บล|บนล่าง|บน|ล่าง|กลับ|รูด|ร|เบิ้ล|หาม|ตอง|ตรง|โต๊ด|6กลับ|3กลับ)/.test(t)) return false;

  // ลูกค้าส่งเลขกลับมาเองครบแล้ว เช่น 20 02 = 20*20
  // ห้ามถาม "ต้องการกลับเลขไหม?" ซ้ำ
  if (hasAlreadyReturnedTwoDigitNumbers_(t)) return false;

  if (!/\d+\s*=\s*\d+(?:\s*\*\s*\d+)?/.test(t)) return false;

  const normalized = t
    .replace(/[,]+/g, " ")
    .replace(/(\d)\s*[\-\_\/]+\s*(\d)/g, "$1 $2");
  const tokens = normalized.split(/[\s\n=*]+/).filter(x => x);
  const nums = tokens.filter(x => /^\d+$/.test(x));
  const has2 = nums.some(x => x.length === 2 || x.length === 1);
  const has3 = nums.some(x => x.length === 3);
  return has2 && !has3;
}


// ============================================================
// PATCH 2026-11-15 — รูด/รุด + ราคาแยกบรรทัด
//
// รองรับ:
//   รุด6/9
//   5*5
//
//   รูด 6/9
//   5×5
//
//   ลูด6,9
//   5*5
//
// canonical:
//   รูด 6/9 บล 5*5
//
// หมายเหตุ:
// - ราคา 5*5 = บน 5 + ล่าง 5
// - รองรับเลขรูดรวมสูงสุด 7 ตัว ตามกฎเดิม
// ============================================================
function normalizeRoodSeparatePriceLines_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const out = [];

  for (let i = 0; i < lines.length; i++) {
    const raw = String(lines[i] || "").trim();

    const m = raw.match(
      /^(?:รูด|รุด|ลูด)\s*([0-9](?:\s*[,\/\-]\s*[0-9]){0,6}|[0-9]{2,7})\s*$/
    );

    if (
      m &&
      i + 1 < lines.length
    ) {
      const priceLine =
        String(lines[i + 1] || "")
          .trim()
          .replace(/\s+/g, "");

      const p = priceLine.match(
        /^(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)$/
      );

      if (p) {
        let digits =
          String(m[1] || "")
            .replace(/\s+/g, "");

        // ถ้าพิมพ์ติดกัน เช่น รูด69 ให้แตกเป็น 6/9
        if (
          /^\d{2,7}$/.test(digits) &&
          !/[\/,\-]/.test(digits)
        ) {
          digits =
            digits.split("").join("/");
        } else {
          digits =
            digits.replace(/[,\\-]+/g, "/");
        }

        out.push(
          "รูด " +
          digits +
          " บล " +
          p[1] +
          "*" +
          p[2]
        );

        i++;
        continue;
      }
    }

    out.push(lines[i]);
  }

  return out.join("\n");
}


// ============================================================
// NORMALIZE รูปแบบภาษาพูด/ลำดับคำสั่งของเลข 2 ตัวและรูด
// ============================================================
function normalizeNaturalTwoDigitAndRoodForms_(text) {
  return String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .split(/\n/)
    .map(function(rawLine) {
      let line = String(rawLine || "").trim();
      if (!line) return line;

      // ----------------------------------------------------
      // PATCH 2026-09-23
      //
      // 34 10*กลับ บ-ล
      // 68 10* กลับ บ-ล
      //
      // = บนล่างกลับ ราคา 10 บาท
      // = 34/43 บน 10 ล่าง 10
      //
      // แปลงเป็น canonical:
      //   34 บลก 10*10
      //
      // รองรับ:
      //   บ-ล / บ/ล / บ ล / บล / บนล่าง
      //   กลับ / กับ
      //   และสลับท้ายเป็น "บ-ล กลับ" ได้ด้วย
      // ----------------------------------------------------
      let returnBothMatch = line.match(
        /^([\d\s,\/\\._-]+?)\s+(\d+(?:\.\d+)?)\s*\*\s*(?:กลับ|กับ)\s*(?:บ\s*[-\/\.\\]?\s*ล|บล|บน\s*[-\/\.\\]?\s*ล่าง|บนล่าง)\s*(?:บาท|บ\.?)?\s*$/i
      );

      if (!returnBothMatch) {
        returnBothMatch = line.match(
          /^([\d\s,\/\\._-]+?)\s+(\d+(?:\.\d+)?)\s*\*\s*(?:บ\s*[-\/\.\\]?\s*ล|บล|บน\s*[-\/\.\\]?\s*ล่าง|บนล่าง)\s*(?:กลับ|กับ)\s*(?:บาท|บ\.?)?\s*$/i
        );
      }

      if (returnBothMatch) {
        const nums = String(returnBothMatch[1] || "")
          .replace(/[,\/\\._-]+/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .split(/\s+/)
          .filter(Boolean);

        // จับเฉพาะเลข 2 ตัว เพื่อไม่ให้ชน parser 3 ตัว
        if (
          nums.length > 0 &&
          nums.every(function(n) {
            return /^\d{2}$/.test(n);
          })
        ) {
          const price =
            String(returnBothMatch[2] || "").trim();

          return (
            nums.join(" ") +
            " บลก " +
            price +
            "*" +
            price
          );
        }
      }

      // ----------------------------------------------------
      // PATCH 2026-09-23
      //
      // รูปแบบ:
      //   บ/ล/ก 13,14,15,34,35,45 ตัว 20
      //
      // ความหมาย:
      //   บนล่างกลับ ตัวละ 20 บาท
      //
      // canonical:
      //   13 14 15 34 35 45 บลก 20*20
      //
      // รองรับ:
      //   บ/ล/ก
      //   บ-ล-ก
      //   บลก
      //   บกล
      //   บนล่างกลับ
      //   คั่นเลขด้วย comma / slash / space
      //   ตัว20 / ตัว 20 / ตัวละ20 / ตัวละ 20
      // ----------------------------------------------------
      let compactBothReturnMatch = line.match(
        /^(?:บ\s*[\/\-\.]\s*ล\s*[\/\-\.]\s*ก|บลก|บกล|บน\s*ล่าง\s*กลับ|บนล่างกลับ)\s+([\d\s,\/\\._-]+?)\s+(?:ตัวละ|ตัว)\s*=?\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?)?\s*$/i
      );

      if (compactBothReturnMatch) {
        const nums = String(compactBothReturnMatch[1] || "")
          .replace(/[,\/\\._-]+/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .split(/\s+/)
          .filter(Boolean);

        if (
          nums.length > 0 &&
          nums.every(function(n) {
            return /^\d{2}$/.test(n);
          })
        ) {
          const price =
            String(compactBothReturnMatch[2] || "").trim();

          return (
            nums.join(" ") +
            " บลก " +
            price +
            "*" +
            price
          );
        }
      }

      // รูด0=5×5 / รูด 0=5*5
      // ทำให้เป็นรูป canonical ที่ rood header parser อ่านได้แน่นอน
      line = line.replace(
        /^รูด\s*(\d)\s*=\s*(\d+)\s*\*\s*(\d+)\s*$/i,
        "รูด $1=$2*$3"
      );

      // ----------------------------------------------------
      // "รูด5กับ2ตัวละ5บาท" = รูด 5 + รูด 2 ตัวละ 5 บาท
      // ไม่ได้หมายถึง "เลข 2 ตัว"
      //
      // รองรับ:
      // รูด5กับ2ตัวละ5บาท
      // รูด 5 กับ 2 ตัวละ 5 บาท
      // รูด5กับรูด2ตัวละ5บาท
      // รูด 5 กับ รูด 2 ตัวละ 5
      //
      // ผลลัพธ์กลาง:
      // รูด 5/2 บล 5*5
      // ----------------------------------------------------
      line = line.replace(
        /รูด\s*(\d)\s*กับ\s*(?:รูด\s*)?(\d)\s*ตัว\s*(?:ละ|ล่ะ|หล่ะ)?\s*(\d+)\s*(?:บาท|บ\.?)?/gi,
        "รูด $1/$2 บล $3*$3"
      );

      // ----------------------------------------------------
      // ไม่มีคำว่า "รูด" แต่เป็นเลขหลักเดียว 2 ตัว + ราคา 2 ฝั่ง
      // 6-7 20*20  => รูด 6/7 บล 20*20
      // 6/7=20*20 => รูด 6/7 บล 20*20
      // 6,7 20*20 => รูด 6/7 บล 20*20
      // จำกัดเป็นเลขหลักเดียวเท่านั้น จึงไม่ชน 66-99
      // ----------------------------------------------------
      let m = line.match(
        /^\s*(\d)\s*[-\/,]\s*(\d)\s*(?:=\s*)?(\d+)\s*\*\s*(\d+)\s*(?:บลก|บกล|บนล่างกลับ|บล|บนล่าง)?\s*$/i
      );
      if (m) {
        const tailMode = /(?:บลก|บกล|บนล่างกลับ)/i.test(line) ? "บลก" : "บล";
        return "รูด " + m[1] + "/" + m[2] + " " + tailMode + " " + m[3] + "*" + m[4];
      }

      // ----------------------------------------------------
      // รูดหลายหลัก: รูด5,6=2*2บลก
      // ให้เรียง mode ก่อนราคา เพื่อให้ splitBlocks แยกบล็อกแน่นอน
      // ----------------------------------------------------
      line = line.replace(
        /รูด\s*([0-9](?:\s*[,\/\-]\s*[0-9]){1,3})\s*=\s*(\d+)\s*\*\s*(\d+)\s*(บลก|บกล|บนล่างกลับ|บล|บนล่าง)?/gi,
        function(_, digits, a, b, mode) {
          let md = String(mode || "บล").replace(/\s+/g, "");
          if (md === "บกล" || md === "บนล่างกลับ") md = "บลก";
          if (md === "บนล่าง") md = "บล";
          return "รูด " + digits.replace(/\s+/g, "") + " " + md + " " + a + "*" + b;
        }
      );

      // ----------------------------------------------------
      // เลข + ราคา + mode อยู่ท้าย
      // 52,51,62,61 2*2 บลก => 52 51 62 61 บลก 2*2
      // รองรับทั้ง บลก และ บล และไม่ต้องมีเครื่องหมาย =
      // ----------------------------------------------------
      m = line.match(
        /^([\d\s,\/\\._-]+?)\s+(\d+)\s*\*\s*(\d+)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล)\s*$/i
      );
      if (m) {
        const nums = String(m[1] || "")
          .replace(/[,\/\\._-]+/g, " ")
          .replace(/\s+/g, " ")
          .trim();
        let md = String(m[4] || "").replace(/\s+/g, "");
        if (md === "บกล" || md === "บนล่างกลับ") md = "บลก";
        if (md === "บนล่าง") md = "บล";
        return nums + " " + md + " " + m[2] + "*" + m[3];
      }

      // ----------------------------------------------------
      // บล 20 บาท / บนล่าง 20 บาท = บน 20 + ล่าง 20
      // ไม่ต้องบังคับลูกค้าเขียน 20*20
      // ----------------------------------------------------
      line = line.replace(
        /^(\s*)(บล|บน\s*ล่าง|บนล่าง)\s*(\d+)\s*(?:บาท|บ\.?)?\s*$/i,
        function(_, lead, mode, price) {
          return lead + "บล " + price + "*" + price;
        }
      );

      return line;
    })
    .join("\n");
}

// ============================================================
// 3 ตัว + ราคา*6 = 6กลับ
//
// ตัวอย่าง:
// 264
// 415
// 851
// 5*6
//
// ความหมาย = 264/415/851 เล่น 6กลับ ชุดละ 5 บาท
// ยอด = 3 ชุด x 6 กลับ x 5 = 90 บาท
//
// สำคัญ:
// - ใช้กฎนี้เฉพาะเมื่อก่อนหน้าเป็นเลข 3 ตัวเท่านั้น
// - 2 ตัว 12 21 5*6 ยังใช้ความหมายราคาบน 5 / ล่าง 6 เดิม
// - รองรับทั้ง 5*6, 5x6, 5×6
// ============================================================

// ============================================================
// 3 ตัวแบบกำหนดราคาต่อเลข = ตรงอย่างเดียว
//
// ตัวอย่าง:
// 027=5 บาท   => 027 ตรง 5 บาท
// 072=5 บาท   => 072 ตรง 5 บาท
// 207=10 บาท  => 207 ตรง 10 บาท
//
// กฎนี้จับเฉพาะ:
// - เลข 3 หลักพอดี
// - ตามด้วย = ราคา
// - ไม่มี * จึงไม่ชนกับ 5*6 / ราคาแบบตรงโต๊ดเดิม
// - ไม่กระทบเลข 2 ตัว
// ============================================================
function normalizeAssignedThreeDigitStraight_(text) {
  return String(text || "")
    .replace(/[×✕✖⨯]/g, "*")
    .split(/\n/)
    .map(function(rawLine) {
      const line = String(rawLine || "").trim();
      if (!line) return line;

      const m = line.match(
        /^(\d{3})\s*=\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?|฿)?\s*$/i
      );

      if (!m) return line;

      // ใส่คำว่า "ตรง" ให้ parser ล็อกเป็น 3 ตัวตรงอย่างเดียว
      return m[1] + " ตรง = " + m[2];
    })
    .join("\n");
}

// Regression test:
// เลข 3 ตัว = ราคา ต้องเป็นตรงอย่างเดียว และรวม 50 บาท
function testAssignedThreeDigitStraight_() {
  const input = [
    "027=5 บาท",
    "072=5 บาท",
    "207=10 บาท",
    "270=10 บาท",
    "702=10 บาท",
    "720=10 บาท"
  ].join("\n");

  const normalized = normalizeAssignedThreeDigitStraight_(input);

  const expected = [
    "027 ตรง = 5",
    "072 ตรง = 5",
    "207 ตรง = 10",
    "270 ตรง = 10",
    "702 ตรง = 10",
    "720 ตรง = 10"
  ].join("\n");

  if (normalized !== expected) {
    throw new Error(
      "FAIL normalizeAssignedThreeDigitStraight_: " + normalized
    );
  }

  const result = parseMessage(input, [], {});

  if (!result || !result.success) {
    throw new Error(
      "FAIL parseAssignedThreeDigitStraight_: " +
      (result && result.error ? result.error : "parse failed")
    );
  }

  if (Number(result.totalAmount) !== 50) {
    throw new Error(
      "FAIL total ต้องเป็น 50 แต่ได้ " + result.totalAmount
    );
  }

  const groups = result.groups || [];
  if (groups.length !== 6) {
    throw new Error(
      "FAIL groups ต้องมี 6 กลุ่ม แต่ได้ " + groups.length
    );
  }

  groups.forEach(function(g, i) {
    if (String(g.mode || "") !== "ตรง") {
      throw new Error(
        "FAIL group " + (i + 1) + " ต้องเป็นโหมดตรง แต่ได้ " + g.mode
      );
    }
    if (Number(g.bottom || 0) !== 0) {
      throw new Error(
        "FAIL group " + (i + 1) + " ต้องไม่มีราคาโต๊ด"
      );
    }
  });

  // กัน regression: 3 ตัว 5*6 ต้องยังเป็น 6กลับตามกฎ V3
  const sixReturn = parseMessage("264\n415\n851\n5*6", [], {});
  if (
    !sixReturn ||
    !sixReturn.success ||
    Number(sixReturn.totalAmount) !== 90 ||
    !sixReturn.groups ||
    !sixReturn.groups[0] ||
    sixReturn.groups[0].mode !== "6กลับ"
  ) {
    throw new Error("FAIL: กฎ 5*6 = 6กลับ ถูกกระทบ");
  }

  return "PASS: 027=5 ฯลฯ = ตรงอย่างเดียว / รวม 50 บาท";
}

function normalizeThreeDigitPriceStarSixReturn_(text) {
  let raw = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*");

  // ----------------------------------------------------------
  // แบบบรรทัดเดียว:
  // 264 415 851 5*6
  // 264/415/851 5*6
  // ->
  // 264 415 851
  // 6กลับ 5
  // ----------------------------------------------------------
  raw = raw
    .split(/\n/)
    .map(function(line) {
      const s = String(line || "").trim();
      if (!s) return line;

      const m = s.match(
        /^((?:\d{3}(?:\s*[,\/\\._-]\s*|\s+))+\d{3})\s+(\d+(?:\.\d+)?)\s*\*\s*6\s*(?:บาท|บ\.?)?$/i
      );

      if (!m) return line;

      const nums = String(m[1] || "")
        .split(/[\s,\/\\._-]+/)
        .filter(Boolean);

      if (
        nums.length > 0 &&
        nums.every(function(n) { return /^\d{3}$/.test(n); })
      ) {
        return nums.join(" ") + "\n6กลับ " + m[2];
      }

      return line;
    })
    .join("\n");

  const lines = raw.split(/\n/);

  function getThreeDigitNumbersOnly_(line) {
    const s = String(line || "").trim();
    if (!s) return [];

    // ต้องมีแค่ตัวเลข 3 ตัว + separator เท่านั้น
    // ห้ามมี mode/ราคา/คำอื่นปน
    if (!/^[0-9\s,\/\\._-]+$/.test(s)) return [];

    const nums = s
      .split(/[\s,\/\\._-]+/)
      .filter(Boolean);

    if (!nums.length) return [];
    if (!nums.every(function(n) { return /^\d{3}$/.test(n); })) return [];

    return nums;
  }

  for (let i = 0; i < lines.length; i++) {
    const priceLine = String(lines[i] || "").trim();

    // 5*6 / 10*6 / 20*6 บาท
    const pm = priceLine.match(
      /^(\d+(?:\.\d+)?)\s*\*\s*6\s*(?:บาท|บ\.?)?$/i
    );

    if (!pm) continue;

    // ย้อนดูบรรทัดก่อนหน้า ต้องเป็นกลุ่มเลข 3 ตัวอย่างน้อย 1 ตัว
    let sawThreeDigit = false;

    for (let j = i - 1; j >= 0; j--) {
      const prev = String(lines[j] || "").trim();

      // อนุญาตบรรทัดว่างในชุดเดียวกัน
      if (!prev) continue;

      const nums3 = getThreeDigitNumbersOnly_(prev);

      if (nums3.length > 0) {
        sawThreeDigit = true;
        continue;
      }

      // เจอข้อความ/โหมด/เลขชนิดอื่น = จบการย้อนหา
      break;
    }

    if (sawThreeDigit) {
      lines[i] = "6กลับ " + pm[1];
    }
  }

  return lines.join("\n");
}

// TEST: 5*6 หลังเลข 3 ตัว = ราคา 5 บาท 6กลับ
function testThreeDigitPriceStarSixReturn_() {
  const input = [
    "264",
    "415",
    "851",
    "5*6"
  ].join("\n");

  const normalized = normalizeThreeDigitPriceStarSixReturn_(input);

  if (!/6กลับ\s+5/.test(normalized)) {
    throw new Error("FAIL normalize 5*6 => 6กลับ 5 / " + normalized);
  }

  const result = parseMessage(input, [], {});

  if (!result || !result.success) {
    throw new Error("FAIL parse: " + (result && result.error));
  }

  if (Number(result.totalAmount) !== 90) {
    throw new Error(
      "FAIL total ต้องเป็น 90 แต่ได้ " + result.totalAmount
    );
  }

  if (
    !result.groups ||
    result.groups.length !== 1 ||
    result.groups[0].mode !== "6กลับ"
  ) {
    throw new Error("FAIL mode ต้องเป็น 6กลับ");
  }

  // กัน regression: เลข 2 ตัว + 5*6 ต้องยังไม่ถูกแปลงเป็น 6กลับ
  const twoDigit = normalizeThreeDigitPriceStarSixReturn_(
    "12\n21\n5*6"
  );

  if (/6กลับ/.test(twoDigit)) {
    throw new Error("FAIL: 2 ตัว 5*6 ถูกแปลงเป็น 6กลับผิด");
  }

  return "PASS: 264/415/851 5*6 = 6กลับ 5 บาท / รวม 90 บาท";
}


// =====================================================
// วินสอง + วินสาม แบบอ้างเลขต้นทางจากบรรทัดก่อน
//
// รองรับ:
// 7048วินสองบลก2
// วินสามหกกับ1
//
// และ:
// 7048 วิน 2 บลก 2
// วิน 3 6กลับ 1
//
// ผลวินสามจาก 7048:
// 704 708 748 048
// แล้วแต่ละชุดนำไป 6กลับตามราคาที่ระบุ
// =====================================================
function normalizeLinkedWinTwoWinThree_(text) {
  const lines = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  let lastWinSource = "";

  const out = lines.map(function(rawLine) {
    let line = String(rawLine || "").trim();

    if (!line) return line;

    // จำเลขต้นทางของ "วินสอง"
    // เช่น 7048วินสองบลก2 / 7048 วิน2 บลก2
    const sourceMatch = line.match(
      /(\d{3,10})\s*วิน\s*(?:สอง|2)(?=\s|บลก|บกล|บล|บน|ล่าง|กลับ|$)/i
    );

    if (sourceMatch) {
      lastWinSource = String(sourceMatch[1] || "");
      return line;
    }

    // "วินสาม..." ไม่มีเลขในบรรทัดตัวเอง
    // ให้ใช้เลขต้นทางจากวินสองบรรทัดก่อนหน้า
    const win3Match = line.match(
      /^วิน\s*(?:สาม|3)\s*(?:(?:หก|6)\s*(?:กลับ|กับ)|6กลับ)\s*=?\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?|฿)?\s*$/i
    );

    if (!win3Match || !lastWinSource) {
      return line;
    }

    const price = String(win3Match[1] || "").trim();

    // ตัดเลขต้นทางซ้ำเหมือนระบบวินสอง
    const digits = [];
    String(lastWinSource).split("").forEach(function(d) {
      if (digits.indexOf(d) === -1) digits.push(d);
    });

    if (digits.length < 3) {
      return line;
    }

    // เลือก 3 ตัวจากเลขต้นทางโดยรักษาลำดับเดิม
    // 7048 => 704, 708, 748, 048
    const triples = [];

    for (let i = 0; i < digits.length - 2; i++) {
      for (let j = i + 1; j < digits.length - 1; j++) {
        for (let k = j + 1; k < digits.length; k++) {
          triples.push(
            digits[i] + digits[j] + digits[k]
          );
        }
      }
    }

    if (triples.length === 0) {
      return line;
    }

    return triples.join(" ") + " 6กลับ " + price;
  });

  return out.join("\n");
}


// ================== TEST: วินสอง + วินสาม ==================
function testLinkedWinTwoWinThree_() {
  const result = parseMessage(
    [
      "7048วินสองบลก2",
      "วินสามหกกับ1"
    ].join("\n"),
    loadKeywords(),
    loadDoors()
  );

  if (!result || !result.success) {
    throw new Error(
      "FAIL: " +
      String(
        result && result.error
          ? result.error
          : "parse failed"
      )
    );
  }

  if (Number(result.totalAmount) !== 72) {
    throw new Error(
      "FAIL: ยอดควรเป็น 72 แต่ได้ " +
      String(result.totalAmount)
    );
  }

  const win2 = (result.groups || []).find(function(g) {
    return !!g.isWin;
  });

  const win3 = (result.groups || []).find(function(g) {
    return g.mode === "6กลับ";
  });

  if (!win2 || (win2.numbers || []).length !== 12) {
    throw new Error(
      "FAIL: วินสองต้องได้ 12 เลข"
    );
  }

  if (!win3 || Number(win3.subtotal) !== 24) {
    throw new Error(
      "FAIL: วินสาม 6กลับ ต้องได้ 24 บาท"
    );
  }

  return {
    success: true,
    totalAmount: result.totalAmount,
    win2Count: win2.numbers.length,
    win3Source: ["704", "708", "748", "048"],
    win3BetCount: win3.betCount,
    groups: result.groups
  };
}



// ============================================================================
// PATCH 2026-09-09 — รูปแบบโพยใหม่
//
// รองรับ:
//   951
//   325
//   783
//   6กับ 2
//
// และ:
//   บลก5*5
//   วิน 87920
//
// "กับ" หลัง 6/3 ใช้ความหมายเดียวกับ "กลับ"
// ============================================================================
function normalizeV0909NewBetFormats_(text) {
  let t = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[×✕✖⨯xX]/g, "*");

  // 6กับ / หกกับ / 3กับ / สามกับ
  t = t
    .replace(/(^|[\n\s])(?:6|หก)\s*กับ\s*=?\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?|฿)?(?=\s*(?:\n|$))/gi,
      "$1 6กลับ $2")
    .replace(/(^|[\n\s])(?:3|สาม)\s*กับ\s*=?\s*(\d+(?:\.\d+)?)\s*(?:บาท|บ\.?|฿)?(?=\s*(?:\n|$))/gi,
      "$1 3กลับ $2");

  const lines = t.split("\n");
  const out = [];

  for (let i = 0; i < lines.length; i++) {
    const current = String(lines[i] || "").trim();
    const next = i + 1 < lines.length
      ? String(lines[i + 1] || "").trim()
      : "";

    // ราคา/โหมดอยู่ก่อน "วิน"
    // บลก5*5
    // วิน 87920
    const modePrice = current.match(
      /^(บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล|บนล่าง|บน\s*ล่าง|บล|กลับ|บน|ล่าง|บ|ล)\s*=?\s*(\d+(?:\.\d+)?(?:\s*\*\s*\d+(?:\.\d+)?)?)\s*(?:บาท|บ\.?|฿)?$/i
    );

    const winNext = next.match(
      /^วิน\s*([0-9]{2,10})\s*$/i
    );

    if (modePrice && winNext) {
      let mode = String(modePrice[1] || "").replace(/\s+/g, "");

      if (mode === "บกล" || mode === "บนล่างกลับ") mode = "บลก";
      if (mode === "บนล่าง") mode = "บล";
      if (mode === "กลับ") mode = "ก";

      const price = String(modePrice[2] || "").replace(/\s+/g, "");
      const source = String(winNext[1] || "");

      out.push("วิน " + source + " " + mode + " " + price);
      i++;
      continue;
    }

    // แบบบรรทัดเดียว: บลก5*5 วิน87920
    const sameLine = current.match(
      /^(บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล|บนล่าง|บน\s*ล่าง|บล|กลับ|บน|ล่าง|บ|ล)\s*=?\s*(\d+(?:\.\d+)?(?:\s*\*\s*\d+(?:\.\d+)?)?)\s+วิน\s*([0-9]{2,10})\s*$/i
    );

    if (sameLine) {
      let mode = String(sameLine[1] || "").replace(/\s+/g, "");

      if (mode === "บกล" || mode === "บนล่างกลับ") mode = "บลก";
      if (mode === "บนล่าง") mode = "บล";
      if (mode === "กลับ") mode = "ก";

      const price = String(sameLine[2] || "").replace(/\s+/g, "");
      const source = String(sameLine[3] || "");

      out.push("วิน " + source + " " + mode + " " + price);
      continue;
    }

    out.push(lines[i]);
  }

  return out.join("\n");
}


// ============================================================================
// PATCH 2026-09-16 — วิน 2 ตัว + วิน 3 ตัว แบบย่อในบรรทัดเดียว
//
// ตัวอย่าง:
//   2917504 3*3/2*6
//
// ความหมาย:
//   - วิน 2 ตัว จาก 2917504 บน 3 / ล่าง 3
//   - วิน 3 ตัว จาก 2917504 ราคา 2 บาท แบบ 6กลับ
//
// ถ้ามีชื่อหวยอยู่ด้านหน้า ระบบหลักจะตัดชื่อหวยออกก่อนเข้า parseMessage()
// เช่น:
//   นอย 2917504 3*3/2*6
// ============================================================================
function extractCompactWinTwoThreeBets_(text) {
  const lines = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const remaining = [];
  const groups = [];

  for (const rawLine of lines) {
    const line = String(rawLine || "").trim();

    if (!line) {
      remaining.push(rawLine);
      continue;
    }

    const m = line.match(
      /^(?:[ก-๙A-Za-z._\-]+\s+)?(\d{3,10})\s+(\d+(?:\.\d+)?)\s*[*xX×✕✖⨯]\s*(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)\s*[*xX×✕✖⨯]\s*6\s*$/i
    );

    if (!m) {
      remaining.push(rawLine);
      continue;
    }

    const source = String(m[1] || "");
    const win2Top = Number(m[2]) || 0;
    const win2Bottom = Number(m[3]) || 0;
    const win3Price = Number(m[4]) || 0;

    if (win2Top <= 0 && win2Bottom <= 0) {
      return {
        success: false,
        error: "❌ ราคาวิน 2 ตัวไม่ถูกต้อง",
        groups: [],
        remainingText: String(text || "")
      };
    }

    if (win3Price <= 0) {
      return {
        success: false,
        error: "❌ ราคาวิน 3 ตัวไม่ถูกต้อง",
        groups: [],
        remainingText: String(text || "")
      };
    }

    // วิน 2 ตัว: generateWin2() สร้างคู่เรียงลำดับทั้งสองทางอยู่แล้ว
    const win2Numbers = generateWin2(source);

    if (win2Numbers.length === 0) {
      return {
        success: false,
        error: "❌ ไม่สามารถวินเลข 2 ตัวได้",
        groups: [],
        remainingText: String(text || "")
      };
    }

    groups.push({
      mode: "บลก",
      top: win2Top,
      bottom: win2Bottom,
      numbers: win2Numbers,
      isWin: true,
      winSource: [source],
      betCount: win2Numbers.length,
      preserveDuplicates: false,
      subtotal: win2Numbers.length * (win2Top + win2Bottom)
    });

    // วิน 3 ตัว: ตัดเลขต้นทางซ้ำก่อน แล้วเลือกทีละ 3 ตัว
    const digits = [];
    source.split("").forEach(function(d) {
      if (digits.indexOf(d) === -1) digits.push(d);
    });

    if (digits.length < 3) {
      return {
        success: false,
        error: "❌ วิน 3 ตัวต้องมีเลขต้นทางอย่างน้อย 3 ตัวไม่ซ้ำ",
        groups: [],
        remainingText: String(text || "")
      };
    }

    const triples = [];
    const win3Numbers = [];
    let win3Subtotal = 0;

    for (let i = 0; i < digits.length - 2; i++) {
      for (let j = i + 1; j < digits.length - 1; j++) {
        for (let k = j + 1; k < digits.length; k++) {
          const triple = digits[i] + digits[j] + digits[k];
          triples.push(triple);

          const variants = perms3(triple);
          win3Numbers.push.apply(win3Numbers, variants);
          win3Subtotal += variants.length * win3Price;
        }
      }
    }

    groups.push({
      mode: "6กลับ",
      top: win3Price,
      bottom: 0,
      numbers: win3Numbers,
      isWin: true,
      winSource: [source],
      win3Source: triples,
      betCount: win3Numbers.length,
      preserveDuplicates: false,
      subtotal: win3Subtotal
    });
  }

  return {
    success: true,
    groups: groups,
    remainingText: remaining.join("\n")
  };
}



// ============================================================================
// PATCH 2026-09-21 — ดึงวินย่อใหม่ก่อนเข้า splitBlocks()
// ============================================================================
function buildWin3CombinationSources_(source) {
  const digits = [];

  String(source || "").replace(/\D/g, "").split("").forEach(function(d) {
    if (digits.indexOf(d) === -1) digits.push(d);
  });

  const triples = [];

  for (let i = 0; i < digits.length - 2; i++) {
    for (let j = i + 1; j < digits.length - 1; j++) {
      for (let k = j + 1; k < digits.length; k++) {
        triples.push(digits[i] + digits[j] + digits[k]);
      }
    }
  }

  return triples;
}


function extractCompactWinShortcutBets_(text) {
  const lines = String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/[×✕✖⨯xX]/g, "*")
    .split("\n");

  const remaining = [];
  const groups = [];

  for (const rawLine of lines) {
    const line = String(rawLine || "").trim();

    if (!line) {
      remaining.push(rawLine);
      continue;
    }

    // ------------------------------------------------------
    // A) 7045239 2*6
    // = วิน 3 ตัว ราคา 2 บาท 6กลับ
    // ------------------------------------------------------
    let m = line.match(
      /^(?:[ก-๙A-Za-z._\-]+\s+)?(\d{4,10})\s+(\d+(?:\.\d+)?)\s*\*\s*6\s*$/i
    );

    if (m) {
      const source = String(m[1] || "");
      const price = Number(m[2]) || 0;
      const triples = buildWin3CombinationSources_(source);

      if (price <= 0) {
        return {
          success: false,
          error: "❌ ราคาวิน 3 ตัวไม่ถูกต้อง",
          groups: [],
          remainingText: String(text || "")
        };
      }

      if (triples.length === 0) {
        return {
          success: false,
          error: "❌ วิน 3 ตัวต้องมีเลขต้นทางอย่างน้อย 3 ตัวไม่ซ้ำ",
          groups: [],
          remainingText: String(text || "")
        };
      }

      const numbers = [];
      let subtotal = 0;

      triples.forEach(function(triple) {
        const variants = perms3(triple);
        numbers.push.apply(numbers, variants);
        subtotal += variants.length * price;
      });

      groups.push({
        mode: "6กลับ",
        top: price,
        bottom: 0,
        numbers: numbers,
        isWin: true,
        winSource: [source],
        win3Source: triples,
        betCount: numbers.length,
        preserveDuplicates: false,
        subtotal: subtotal
      });

      continue;
    }

    // ------------------------------------------------------
    // B) 780314 บลก 1*1
    // = วิน 2 ตัว บน/ล่าง/กลับ 1*1
    // generateWin2() มีทั้ง 78 และ 87 อยู่แล้ว
    // ------------------------------------------------------
    m = line.match(
      /^(?:[ก-๙A-Za-z._\-]+\s+)?(\d{4,10})\s+(?:บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล)\s*=?\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*$/i
    );

    if (m) {
      const source = String(m[1] || "");
      const top = Number(m[2]) || 0;
      const bottom = Number(m[3]) || 0;
      const numbers = generateWin2(source);

      if (numbers.length === 0) {
        return {
          success: false,
          error: "❌ ไม่สามารถวินเลข 2 ตัวได้",
          groups: [],
          remainingText: String(text || "")
        };
      }

      groups.push({
        mode: "บลก",
        top: top,
        bottom: bottom,
        numbers: numbers,
        isWin: true,
        winSource: [source],
        betCount: numbers.length,
        preserveDuplicates: false,
        subtotal: numbers.length * (top + bottom)
      });

      continue;
    }

    // ------------------------------------------------------
    // C) 780314 บต1*1
    // = วิน 3 ตัว ตรง + โต๊ด 1*1
    //
    // สร้างเลข 3 ตัวแบบ combination จากเลขต้นทาง
    // เช่น 6 ตัวไม่ซ้ำ => C(6,3) = 20 ชุด
    // แต่ละชุดคิด ตรง 1 + โต๊ด 1
    // ------------------------------------------------------
    m = line.match(
      /^(?:[ก-๙A-Za-z._\-]+\s+)?(\d{4,10})\s+(?:บต|ตรงโต๊ด|ตรง\s*โต๊ด|บนตรงโต๊ด)\s*=?\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*$/i
    );

    if (m) {
      const source = String(m[1] || "");
      const directPrice = Number(m[2]) || 0;
      const todPrice = Number(m[3]) || 0;
      const triples = buildWin3CombinationSources_(source);

      if (triples.length === 0) {
        return {
          success: false,
          error: "❌ วินตรงโต๊ดต้องมีเลขต้นทางอย่างน้อย 3 ตัวไม่ซ้ำ",
          groups: [],
          remainingText: String(text || "")
        };
      }

      groups.push({
        mode: "ตท",
        top: directPrice,
        bottom: todPrice,
        numbers: triples,
        isWin: true,
        winSource: [source],
        win3Source: triples,
        betCount: triples.length,
        preserveDuplicates: false,
        subtotal: triples.length * (directPrice + todPrice)
      });

      continue;
    }

    remaining.push(rawLine);
  }

  return {
    success: true,
    groups: groups,
    remainingText: remaining.join("\n")
  };
}




// ============================================================================
// PATCH 2026-09-24 — เลขชุด = ราคา บลก
//
// ตัวอย่าง:
//   16,17,19,15,67,69,65,79,75,95=2×2บลก
//
// ความหมาย:
//   เลขทุกตัวเล่น บน/ล่าง/กลับ
//   ราคา บน 2 + ล่าง 2
//
// สำคัญ:
// - เก็บเลขซ้ำทุก occurrence
// - บลก / บกล / บนล่างกลับ = กลับเลข
// - บล / บนล่าง = ไม่กลับเลข
// ============================================================================
function extractExplicitTwoDigitEqualsModeBets_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const groups = [];
  const remaining = [];

  for (const rawLine of lines) {
    const line = String(rawLine || "").trim();

    if (!line) {
      remaining.push(rawLine);
      continue;
    }

    const m = line.match(
      /^((?:\d{2}\s*[,\/\\._-]\s*)+\d{2}|\d{2}(?:\s+\d{2})+)\s*=\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล)\s*$/i
    );

    if (!m) {
      remaining.push(rawLine);
      continue;
    }

    const sourceNums = String(m[1] || "")
      .replace(/[,\/\\._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(/\s+/)
      .filter(Boolean);

    if (
      sourceNums.length === 0 ||
      !sourceNums.every(function(n) {
        return /^\d{2}$/.test(n);
      })
    ) {
      remaining.push(rawLine);
      continue;
    }

    const top = Number(m[2]) || 0;
    const bottom = Number(m[3]) || 0;

    let mode = String(m[4] || "")
      .replace(/\s+/g, "");

    if (
      mode === "บกล" ||
      mode === "บนล่างกลับ"
    ) {
      mode = "บลก";
    }

    if (mode === "บนล่าง") {
      mode = "บล";
    }

    const numbers = [];

    if (mode === "บลก") {
      sourceNums.forEach(function(n) {
        numbers.push(n);

        const rev =
          n.charAt(1) + n.charAt(0);

        if (rev !== n) {
          numbers.push(rev);
        }
      });
    } else {
      sourceNums.forEach(function(n) {
        numbers.push(n);
      });
    }

    groups.push({
      mode: mode,
      top: top,
      bottom: bottom,
      numbers: numbers,
      isWin: false,
      betCount: numbers.length,
      preserveDuplicates: true,
      subtotal:
        numbers.length *
        (top + bottom)
    });
  }

  return {
    success: true,
    groups: groups,
    remainingText:
      remaining.join("\n")
  };
}






// ============================================================================
// PATCH 2026-11-11 — เลขหลายบรรทัด + ราคาอยู่หน้าคำสั่ง
//
// รองรับ:
//   02/19/36/45
//   37/81/01
//   8บลก
//   11=40บล
//
// ความหมาย:
// - เลขทั้งหมดก่อน "8บลก" ใช้ บน/ล่าง/กลับ 8*8
// - "11=40บล" = 11 บน/ล่าง 40*40 (ไม่กลับ)
// ============================================================================
function extractTrailingPriceModeTwoDigitBets_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const groups = [];
  const consumed = new Set();

  function parseNumberList_(line) {
    const raw = String(line || "").trim();
    if (!raw) return null;

    if (!/^\d{2}(?:\s*[,\/\\._-]?\s*\d{2})*$/.test(raw)) {
      return null;
    }

    const nums = raw
      .replace(/[,\/\\._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .filter(Boolean);

    if (!nums.length) return null;
    if (!nums.every(function(n) { return /^\d{2}$/.test(n); })) {
      return null;
    }

    return nums;
  }

  function normalizeMode_(mode) {
    let m = String(mode || "")
      .replace(/\s+/g, "")
      .toLowerCase();

    if (m === "บกล" || m === "บนล่างกลับ") return "บลก";
    if (m === "บนล่าง") return "บล";
    if (m === "บ") return "บน";
    if (m === "ล") return "ล่าง";
    return m;
  }

  function makeGroup_(nums, mode, price) {
    const m = normalizeMode_(mode);
    const p = Number(price) || 0;

    let top = 0;
    let bottom = 0;
    let numbers = nums.slice();

    if (m === "บลก") {
      top = p;
      bottom = p;

      const withReturn = [];
      nums.forEach(function(n) {
        withReturn.push(n);

        const rev = n.charAt(1) + n.charAt(0);
        if (rev !== n) {
          withReturn.push(rev);
        }
      });

      numbers = withReturn;
    } else if (m === "บล") {
      top = p;
      bottom = p;
    } else if (m === "บน") {
      top = p;
      bottom = 0;
    } else if (m === "ล่าง") {
      top = 0;
      bottom = p;
    } else {
      return null;
    }

    return {
      mode: m,
      top: top,
      bottom: bottom,
      numbers: numbers,
      isWin: false,
      betCount: numbers.length,
      preserveDuplicates: true,
      subtotal: numbers.length * (top + bottom)
    };
  }

  // ------------------------------------------------------------
  // A) เลขหลายบรรทัด แล้วปิดท้ายด้วย ราคา+mode
  //    02/19/36/45
  //    37/81/01
  //    8บลก
  // ------------------------------------------------------------
  let i = 0;

  while (i < lines.length) {
    if (consumed.has(i)) {
      i++;
      continue;
    }

    const firstNums = parseNumberList_(lines[i]);
    if (!firstNums) {
      i++;
      continue;
    }

    let nums = firstNums.slice();
    const numberLineIndexes = [i];
    let j = i + 1;

    while (j < lines.length) {
      const more = parseNumberList_(lines[j]);
      if (!more) break;

      nums.push.apply(nums, more);
      numberLineIndexes.push(j);
      j++;
    }

    if (j < lines.length) {
      const priceMode = String(lines[j] || "")
        .trim()
        .match(
          /^(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
        );

      if (priceMode) {
        // สำคัญ: ราคาในบรรทัดนี้ใช้กับเลขที่สะสมมาก่อนหน้าเท่านั้น
        // เช่น 8บลก ต้องจบกลุ่ม 02/19/.../01 ที่ราคา 8*8 ทันที
        // ห้ามให้ราคาในบรรทัดถัดไป เช่น 11=40บล ไหลย้อนกลับมาทับกลุ่มนี้
        const group = makeGroup_(
          nums,
          priceMode[2],
          priceMode[1]
        );

        if (group) {
          groups.push(group);

          numberLineIndexes.forEach(function(idx) {
            consumed.add(idx);
          });
          consumed.add(j);

          i = j + 1;
          continue;
        }
      }
    }

    i++;
  }

  // ------------------------------------------------------------
  // B) เลข=ราคา+mode
  //    11=40บล
  //    12=5บลก
  // ------------------------------------------------------------
  for (let k = 0; k < lines.length; k++) {
    if (consumed.has(k)) continue;

    const line = String(lines[k] || "").trim();

    const m = line.match(
      /^(\d{2})\s*=\s*(\d+(?:\.\d+)?)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล|บน|ล่าง|บ|ล)\s*$/
    );

    if (!m) continue;

    const group = makeGroup_(
      [m[1]],
      m[3],
      m[2]
    );

    if (!group) continue;

    groups.push(group);
    consumed.add(k);
  }

  const remaining = lines.filter(function(_, idx) {
    return !consumed.has(idx);
  });

  return {
    success: true,
    groups: groups,
    remainingText: remaining.join("\n")
  };
}



// ============================================================================
// PATCH 2026-11-11 — ราคาแบบต่อเนื่อง (sticky price)
//
// ตัวอย่าง:
//   82=5×5
//   84
//   81
//   85
//   86
//   87
//   83
//   38=3×3
//   78
//   68
//   58
//   18
//   48
//   28
//   บนล่าง
//
// ความหมาย:
// - 82 84 81 85 86 87 83 ใช้ราคา 5×5
// - 38 78 68 58 18 48 28 ใช้ราคา 3×3
// - ราคาล่าสุดมีผลกับเลขถัดไป จนกว่าจะเจอบรรทัดเลข=ราคาใหม่
// - คำว่า "บนล่าง" ท้ายชุดใช้กับทุกกลุ่มราคา
// ============================================================================
function extractStickyTwoDigitPriceGroups_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  function parsePricedNumber_(line) {
    const m = String(line || "")
      .trim()
      .match(
        /^(\d{2})\s*(?:=|@|:)\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*$/
      );

    if (!m) return null;

    return {
      number: m[1],
      top: Number(m[2]) || 0,
      bottom: Number(m[3]) || 0
    };
  }

  function parseBareNumbers_(line) {
    const raw = String(line || "").trim();
    if (!raw) return null;

    // รองรับทั้งเลขตัวเดียวต่อบรรทัด และหลายเลขคั่น space/comma/slash
    if (!/^\d{2}(?:\s*[,\/._-]?\s*\d{2})*$/.test(raw)) {
      return null;
    }

    const nums = raw
      .replace(/[,\/._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .filter(Boolean);

    if (!nums.length || !nums.every(function(n) {
      return /^\d{2}$/.test(n);
    })) {
      return null;
    }

    return nums;
  }

  function parseTerminalMode_(line) {
    const mode = String(line || "")
      .replace(/\s+/g, "")
      .trim();

    if (/^(?:บนล่าง|บล)$/.test(mode)) return "บล";
    if (/^บน$/.test(mode)) return "บน";
    if (/^ล่าง$/.test(mode)) return "ล่าง";
    return null;
  }

  const consumed = new Set();
  const rawGroups = [];

  let i = 0;
  while (i < lines.length) {
    const first = parsePricedNumber_(lines[i]);
    if (!first) {
      i++;
      continue;
    }

    const start = i;
    const candidateGroups = [];
    let current = {
      top: first.top,
      bottom: first.bottom,
      numbers: [first.number]
    };

    i++;
    let terminalMode = null;
    let end = -1;

    while (i < lines.length) {
      const raw = String(lines[i] || "").trim();

      if (!raw) {
        i++;
        continue;
      }

      const nextPriced = parsePricedNumber_(raw);
      if (nextPriced) {
        candidateGroups.push(current);
        current = {
          top: nextPriced.top,
          bottom: nextPriced.bottom,
          numbers: [nextPriced.number]
        };
        i++;
        continue;
      }

      const nums = parseBareNumbers_(raw);
      if (nums) {
        current.numbers.push.apply(current.numbers, nums);
        i++;
        continue;
      }

      terminalMode = parseTerminalMode_(raw);
      if (terminalMode) {
        candidateGroups.push(current);
        end = i;
        break;
      }

      // เจอข้อความอื่นก่อน mode ท้ายชุด = ไม่ถือว่าเป็นรูปแบบ sticky นี้
      break;
    }

    if (terminalMode && end >= start) {
      for (let x = start; x <= end; x++) {
        consumed.add(x);
      }

      candidateGroups.forEach(function(g) {
        let top = 0;
        let bottom = 0;

        if (terminalMode === "บล") {
          top = g.top;
          bottom = g.bottom;
        } else if (terminalMode === "บน") {
          top = g.top;
          bottom = 0;
        } else {
          top = 0;
          bottom = g.bottom;
        }

        rawGroups.push({
          mode: terminalMode,
          top: top,
          bottom: bottom,
          numbers: g.numbers.slice(),
          isWin: false,
          betCount: g.numbers.length,
          preserveDuplicates: true,
          subtotal: g.numbers.length * (top + bottom)
        });
      });

      i = end + 1;
      continue;
    }

    // ไม่ match จริง ให้เดินต่อจากบรรทัดเริ่ม เพื่อไม่กลืนข้อความเดิม
    i = start + 1;
  }

  const remaining = lines.filter(function(_, idx) {
    return !consumed.has(idx);
  });

  return {
    success: true,
    groups: rawGroups,
    remainingText: remaining.join("\n")
  };
}


// ============================================================================
// PATCH 2026-11-11 — ราคาเลขตรง/เลขกลับแยกกัน
//
// ตัวอย่าง:
//   80 81 83 84 85 86 87 89
//   บนล่าง@5×5
//   กลับ@3×3
//
// ความหมาย:
// - เลขต้นฉบับ เล่นบน/ล่าง 5×5
// - เฉพาะเลขกลับ เล่นบน/ล่าง 3×3
// - เลขเบิ้ล เช่น 11 ไม่มีเลขกลับเพิ่ม เพราะกลับแล้วยังเป็น 11
// - เก็บเลขซ้ำตามจำนวนครั้งที่ลูกค้าพิมพ์
// ============================================================================
function extractDirectAndReturnSeparatePriceBets_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const groups = [];
  const remaining = [];

  function parseNumberLine_(line) {
    const raw = String(line || "").trim();
    if (!raw) return null;

    // ต้องเป็นบรรทัดเลข 2 ตัวล้วน ๆ คั่นด้วย space/comma/slash/dash/dot
    if (!/^\d{2}(?:\s*[,\/\\._-]?\s*\d{2})+$/.test(raw)) {
      return null;
    }

    const nums = raw
      .replace(/[,\/\\._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .filter(Boolean);

    if (!nums.length || !nums.every(function(n) { return /^\d{2}$/.test(n); })) {
      return null;
    }

    return nums;
  }

  function parseSidePrice_(line, type) {
    const raw = String(line || "")
      .replace(/\s+/g, " ")
      .trim();

    let re;
    if (type === "direct") {
      re = /^(?:บน\s*ล่าง|บนล่าง|บล)\s*(?:@|=|:)?\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*$/i;
    } else {
      re = /^(?:กลับ|ก)\s*(?:@|=|:)?\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*$/i;
    }

    const m = raw.match(re);
    if (!m) return null;

    return {
      top: Number(m[1]) || 0,
      bottom: Number(m[2]) || 0
    };
  }

  let i = 0;
  while (i < lines.length) {
    const nums = parseNumberLine_(lines[i]);

    if (nums && i + 2 < lines.length) {
      const directPrice = parseSidePrice_(lines[i + 1], "direct");
      const returnPrice = parseSidePrice_(lines[i + 2], "return");

      if (directPrice && returnPrice) {
        // กลุ่มตรง: ใช้เลขที่ลูกค้าพิมพ์เท่านั้น
        groups.push({
          mode: "บล",
          top: directPrice.top,
          bottom: directPrice.bottom,
          numbers: nums.slice(),
          isWin: false,
          betCount: nums.length,
          preserveDuplicates: true,
          subtotal: nums.length * (directPrice.top + directPrice.bottom)
        });

        // กลุ่มกลับ: ใช้เฉพาะเลขที่กลับแล้ว ไม่รวมเลขต้นฉบับ
        const reversed = [];
        nums.forEach(function(n) {
          const rev = n.charAt(1) + n.charAt(0);
          if (rev !== n) reversed.push(rev);
        });

        if (reversed.length > 0) {
          groups.push({
            mode: "กลับ",
            top: returnPrice.top,
            bottom: returnPrice.bottom,
            numbers: reversed,
            isWin: false,
            betCount: reversed.length,
            preserveDuplicates: true,
            subtotal: reversed.length * (returnPrice.top + returnPrice.bottom)
          });
        }

        i += 3;
        continue;
      }
    }

    remaining.push(lines[i]);
    i++;
  }

  return {
    success: true,
    groups: groups,
    remainingText: remaining.join("\n")
  };
}


// ============================================================================
// PATCH 2026-11-15 — วินรวมเบิ้ล
//
// รองรับ:
//   วิน3857602
//   5*5
//   บลก
//   วินรวมเบิ้ล
//
// รวมทั้ง typo ที่พบบ่อย: วินรวมเบิ้น
//
// ความหมาย:
// - ใช้เลขต้นทางแบบไม่ซ้ำ ตามลำดับที่ลูกค้าพิมพ์
// - สร้างคู่แบบ combination with replacement (i <= j)
// - จึงได้ทั้งเลขเบิ้ล 33 88 ... และคู่ 38 35 ...
// - ถ้าโหมดเป็น บลก ให้เพิ่มเลขกลับของทุกตัวที่ไม่ใช่เลขเบิ้ล
// - ตัวเบิ้ล เช่น 33, 88 ไม่สร้างซ้ำ
// - ราคา 5*5 = บน 5 + ล่าง 5
// ============================================================================
function generateWinCombinedDouble2_(source) {
  const digits = [];
  String(source || "")
    .replace(/\D/g, "")
    .split("")
    .forEach(function(d) {
      if (digits.indexOf(d) === -1) digits.push(d);
    });

  if (digits.length < 1) return [];

  const out = [];
  for (let i = 0; i < digits.length; i++) {
    for (let j = i; j < digits.length; j++) {
      out.push(digits[i] + digits[j]);
    }
  }

  return out;
}

function extractWinCombinedDoubleBets_(text) {
  const lines = String(text || "")
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\r\n?/g, "\n")
    .split("\n");

  const groups = [];
  const consumed = new Set();

  function cleanLine_(v) {
    return String(v || "").trim();
  }

  for (let i = 0; i < lines.length; i++) {
    if (consumed.has(i)) continue;

    const first = cleanLine_(lines[i]);
    const mWin = first.match(/^วิน\s*([0-9]{1,10})\s*$/i);
    if (!mWin) continue;

    // ข้ามบรรทัดว่างได้ แต่ต้องเรียง: ราคา -> บลก/บล -> วินรวมเบิ้ล
    let j = i + 1;
    while (j < lines.length && !cleanLine_(lines[j])) j++;
    if (j >= lines.length) continue;

    const priceLine = cleanLine_(lines[j]).replace(/\s+/g, "");
    const mPrice = priceLine.match(/^(\d+(?:\.\d+)?)\*(\d+(?:\.\d+)?)$/);
    if (!mPrice) continue;

    let k = j + 1;
    while (k < lines.length && !cleanLine_(lines[k])) k++;
    if (k >= lines.length) continue;

    const modeLine = cleanLine_(lines[k]).replace(/\s+/g, "");
    if (!/^(?:บลก|บกล|บนล่างกลับ|บล|บนล่าง)$/i.test(modeLine)) continue;

    let q = k + 1;
    while (q < lines.length && !cleanLine_(lines[q])) q++;
    if (q >= lines.length) continue;

    const cmd = cleanLine_(lines[q]).replace(/\s+/g, "");
    if (!/^(?:วิน)?รวมเบิ(?:้ล|้น|ล)$/i.test(cmd)) continue;

    const source = mWin[1];
    const baseNumbers = generateWinCombinedDouble2_(source);
    if (!baseNumbers.length) continue;

    const normalizedMode = /^(?:บลก|บกล|บนล่างกลับ)$/i.test(modeLine)
      ? "บลก"
      : "บล";

    const numbers = [];
    const seen = {};

    function pushUnique_(n) {
      const key = String(n || "");
      if (!/^\d{2}$/.test(key)) return;
      if (seen[key]) return;
      seen[key] = true;
      numbers.push(key);
    }

    // ชุดหลัก: combination with replacement
    baseNumbers.forEach(function(n) {
      pushUnique_(n);
    });

    // ถ้ามี บลก ให้เพิ่มตัวกลับเฉพาะเลขที่ไม่ใช่เบิ้ล
    if (normalizedMode === "บลก") {
      baseNumbers.forEach(function(n) {
        const s = String(n || "");
        if (!/^\d{2}$/.test(s)) return;
        if (s.charAt(0) === s.charAt(1)) return;

        const rev = s.charAt(1) + s.charAt(0);
        pushUnique_(rev);
      });
    }

    const top = Number(mPrice[1]) || 0;
    const bottom = Number(mPrice[2]) || 0;

    groups.push({
      mode: normalizedMode,
      top: top,
      bottom: bottom,
      numbers: numbers,
      isWin: true,
      winSource: [source],
      winCombinedDouble: true,
      winCombinedDoubleWithReverse: normalizedMode === "บลก",
      betCount: numbers.length,
      preserveDuplicates: false,
      subtotal: numbers.length * (top + bottom)
    });

    consumed.add(i);
    consumed.add(j);
    consumed.add(k);
    consumed.add(q);
  }

  return {
    success: true,
    groups: groups,
    remainingText: lines.filter(function(_, idx) {
      return !consumed.has(idx);
    }).join("\n")
  };
}


// ================== PARSE MESSAGE (multi-block + stack) ==================
function parseMessage(text, keywords, doorMap) {

  // วินรวมเบิ้ล เช่น:
  // วิน3857602 / 5*5 / บลก / วินรวมเบิ้ล
  const winCombinedDoubleResult =
    extractWinCombinedDoubleBets_(text);

  if (!winCombinedDoubleResult.success) {
    return {
      success: false,
      error: winCombinedDoubleResult.error
    };
  }

  const winCombinedDoubleGroups =
    winCombinedDoubleResult.groups || [];

  text = winCombinedDoubleResult.remainingText;

  // เลขหลายบรรทัด + ราคาอยู่หน้าคำสั่ง
  // เช่น:
  // 02/19/36/45
  // 37/81/01
  // 8บลก
  // 11=40บล
  const trailingPriceModeResult =
    extractTrailingPriceModeTwoDigitBets_(text);

  if (!trailingPriceModeResult.success) {
    return {
      success: false,
      error: trailingPriceModeResult.error
    };
  }

  const trailingPriceModeGroups =
    trailingPriceModeResult.groups || [];

  text = trailingPriceModeResult.remainingText;

  // ราคาแบบต่อเนื่อง เช่น:
  // 82=5×5
  // 84
  // ...
  // 38=3×3
  // 78
  // ...
  // บนล่าง
  const stickyPriceResult =
    extractStickyTwoDigitPriceGroups_(text);

  if (!stickyPriceResult.success) {
    return {
      success: false,
      error: stickyPriceResult.error
    };
  }

  const stickyPriceGroups =
    stickyPriceResult.groups || [];

  text = stickyPriceResult.remainingText;

  // ราคาเลขตรง/เลขกลับแยกกัน
  // เช่น:
  // 80 81 83 84 85 86 87 89
  // บนล่าง@5×5
  // กลับ@3×3
  const directReturnResult =
    extractDirectAndReturnSeparatePriceBets_(text);

  if (!directReturnResult.success) {
    return {
      success: false,
      error: directReturnResult.error
    };
  }

  const directReturnGroups =
    directReturnResult.groups || [];

  text = directReturnResult.remainingText;

  // อ่านรูปแบบย่อ 2917504 3*3/2*6 ก่อน splitBlocks()
  // เพราะ parser เดิมอาจตีความ 3*3 เป็น 3กลับ
  const compactWinResult = extractCompactWinTwoThreeBets_(text);

  if (!compactWinResult.success) {
    return {
      success: false,
      error: compactWinResult.error
    };
  }

  const compactWinGroups = compactWinResult.groups || [];
  text = compactWinResult.remainingText;

  // รูปแบบย่อเพิ่มเติม:
  // 7045239 2*6
  // 780314 บลก 1*1
  // 780314 บต1*1
  const compactShortcutResult = extractCompactWinShortcutBets_(text);

  if (!compactShortcutResult.success) {
    return {
      success: false,
      error: compactShortcutResult.error
    };
  }

  const compactShortcutGroups = compactShortcutResult.groups || [];
  text = compactShortcutResult.remainingText;

  // เลขชุด = ราคา บลก
  // เช่น 16,17,...,95=2×2บลก
  const explicitTwoDigitResult =
    extractExplicitTwoDigitEqualsModeBets_(
      text
    );

  if (!explicitTwoDigitResult.success) {
    return {
      success: false,
      error: explicitTwoDigitResult.error
    };
  }

  const explicitTwoDigitGroups =
    explicitTwoDigitResult.groups || [];

  text =
    explicitTwoDigitResult.remainingText;

  text = normalizeV0909NewBetFormats_(text);

// =====================================================
  // วินสอง + วินสาม ใช้เลขต้นทางชุดเดียวกัน
  //
  // ตัวอย่าง:
  // 7048วินสองบลก2
  // วินสามหกกับ1
  //
  // วินสอง = 12 เลข x (บน2 + ล่าง2) = 48 บาท
  // วินสาม = 704 708 748 048 อย่างละ 6กลับ x 1 = 24 บาท
  // รวม = 72 บาท
  // =====================================================
  text = normalizeLinkedWinTwoWinThree_(text);

  // รูด/รุด + ราคาแยกบรรทัด
  // เช่น รุด6/9 \n 5*5 -> รูด 6/9 บล 5*5
  text = normalizeRoodSeparatePriceLines_(text);

  // แปลงรูปภาษาพูด/คำสั่งกลับด้านให้เป็นรูปมาตรฐานก่อน
  text = normalizeNaturalTwoDigitAndRoodForms_(text);

  // เลข 3 ตัวแบบ 027=5 บาท = ตรงอย่างเดียว ราคา 5 บาท
  // ต้องทำก่อนกฎ 5*6 เพื่อแยก syntax ให้ชัดเจน
  text = normalizeAssignedThreeDigitStraight_(text);

  // เลข 3 ตัว + 5*6 = ราคา 5 บาท แบบ 6กลับ
  // เช่น 264 / 415 / 851 / 5*6 => 3 ชุด x 6 x 5 = 90 บาท
  text = normalizeThreeDigitPriceStarSixReturn_(text);

  // รองรับรูปที่ลูกค้าวางราคาไว้หน้าคำสั่ง เช่น
  // 23,20,28=3*3บลก  => 23 20 28 บลก 3*3
  // 22,33,00=2*2บล   => 22 33 00 บล 2*2
  text = normalizePostfixTwoDigitModePrice_(text);

  // บลก ราคาเดียว = บน/ล่างราคาเท่ากันเสมอ
  // ทำตั้งแต่ต้นเพื่อให้ทุกเส้นทาง parser ได้กฎเดียวกัน
  text = normalizeTwoSideReturnSinglePrice_(text);
  text = normalizeBetKeywordTypos(text);
  text = normalizePriceSeparators(text);
  text = stripPriceWords(text);
  text = splitInlineModeBlocks(text);

  const { roodLines, otherLines } = extractRoodHeaderLines(text);

  let roodResult = null;
  if (roodLines.length > 0) {
    const roodText = roodLines.join("\n");
    roodResult = parseRoodHeaderMessage(roodText, doorMap);
    if (roodResult && !roodResult.success) return roodResult;
  }

  const allGroups = [];

  // วินรวมเบิ้ล
  if (winCombinedDoubleGroups.length > 0) {
    allGroups.push(...winCombinedDoubleGroups);
  }

  // กลุ่มเลขหลายบรรทัด + ราคาอยู่หน้าคำสั่ง
  if (trailingPriceModeGroups.length > 0) {
    allGroups.push(...trailingPriceModeGroups);
  }

  // กลุ่มราคาแบบต่อเนื่อง เลข=ราคา แล้วเลขถัดไปใช้ราคาล่าสุด
  if (stickyPriceGroups.length > 0) {
    allGroups.push(...stickyPriceGroups);
  }

  // กลุ่มเลขตรง/เลขกลับที่แยกราคา
  if (directReturnGroups.length > 0) {
    allGroups.push(...directReturnGroups);
  }

  // กลุ่มจาก syntax ย่อ 2917504 3*3/2*6
  if (compactWinGroups.length > 0) {
    allGroups.push(...compactWinGroups);
  }

  // กลุ่มจาก syntax ย่อใหม่
  if (compactShortcutGroups.length > 0) {
    allGroups.push(...compactShortcutGroups);
  }

  // เลขชุด = ราคา บลก
  if (explicitTwoDigitGroups.length > 0) {
    allGroups.push(
      ...explicitTwoDigitGroups
    );
  }

  if (roodResult && roodResult.success && roodResult.groups) {
    allGroups.push(...roodResult.groups);
  }

  if (otherLines.length > 0) {
    const normalText = otherLines.join("\n");
    const blocks = splitBlocks(normalText);
    for (const b of blocks) {
      // SKIP เศษบล็อกที่ไม่มีอะไรเป็นแก่นสาร: ไม่มีราคา + ไม่มีเลข
      const bodyStr = Array.isArray(b.body) ? b.body.join(" ") : b.body;
      const hasNumber = /\d/.test(bodyStr);
      if (!b.priceStr && !hasNumber) continue;

      if (!b.priceStr) {
        return { success: false, error: "❌ ไม่พบราคาในบล็อก" };
      }
      const p = parseBlock(bodyStr, b.priceStr, keywords, doorMap, b.header, b.isHeaderOnly);
      if (!p.success) return p;
      const parsedGroups = p.groups || [p];
      parsedGroups.forEach(pg => {
        allGroups.push({
          mode: pg.mode,
          top: pg.top,
          bottom: pg.bottom,
          numbers: pg.numbers,
          isWin: !!pg.isWin,
          winSource: pg.winSource || [],
          betCount: (typeof pg.betCount === "number") ? pg.betCount : (pg.numbers || []).length,
          preserveDuplicates: !!pg.preserveDuplicates,
          subtotal: (typeof pg.subtotal3 === "number")
            ? pg.subtotal3
            : (typeof pg.subtotal2 === "number")
              ? pg.subtotal2
              : pg.numbers.length * (pg.top + pg.bottom)
        });
      });
    }
  }

  if (allGroups.length === 0) {
    return { success: false, error: "❌ ไม่พบราคา (เช่น = 10*10)" };
  }

  let totalAmount = 0;
  let totalBetCount = 0;

  allGroups.forEach(function(g) {
    totalAmount += g.subtotal;
    totalBetCount += (typeof g.betCount === "number")
      ? g.betCount
      : (g.numbers || []).length;
  });

  return {
    success: true,
    groups: allGroups,
    // ตัว = จำนวนรายการแทงจริง ไม่ใช่จำนวนเลข unique
    totalDoors: totalBetCount,
    grossDoors: totalBetCount,
    totalAmount: totalAmount
  };
}


// ================== ROOD HEADER PARSER ==================
function parseRoodHeaderMessage(text, doorMap) {
  text = text
    .replace(/(\d)\s*[xX]\s*(\d)/g, "$1*$2")
    .replace(/[ \t]*=[ \t]*/g, "=");

  const lines = text.split(/\n/).map(l=>l.trim()).filter(l=>l);
  const ROOD_HEADERS = [
    { text: "รูดบนล่าง", mode: "รูดบล" },
    { text: "รูดบน ล่าง", mode: "รูดบล" },
    { text: "รูด บนล่าง", mode: "รูดบล" },
    { text: "รูด บน ล่าง", mode: "รูดบล" },
    { text: "รูดบน",     mode: "รูดบ"  },
    { text: "รูด บน",    mode: "รูดบ"  },
    { text: "รูดล่าง",    mode: "รูดล"  },
    { text: "รูด ล่าง",   mode: "รูดล"  },
    { text: "รูด",       mode: "รูดบล" }
  ];

  function detectHeader(line) {
    for (const h of ROOD_HEADERS) {
      if (line === h.text) return h;
      if (line.startsWith(h.text + " ")) {
        const rest = line.substring(h.text.length).trim();
        if (h.text === "รูด") {
          const tokens = rest.split(/\s+/);
          const firstToken = tokens[0];
          const isHeaderTokenForm = /^\d=\d+(\*\d+)?$/.test(firstToken);
          const isPairedSpaceForm =
            tokens.length >= 2 &&
            tokens.length % 2 === 0 &&
            /^\d$/.test(tokens[0]) &&
            /^\d+(\*\d+)?$/.test(tokens[1]);
          if (!isHeaderTokenForm && !isPairedSpaceForm) return null;
        }
        return h;
      }
    }
    return null;
  }

  let hasHeader = false;
  for (const line of lines) {
    if (detectHeader(line)) { hasHeader=true; break; }
  }
  if (!hasHeader) return null;

  const blocks = [];
  let currentMode = null;

  function parseLineMaybeMulti(line, mode) {
    const t = line.replace(/(\d)\s*[xX]\s*(\d)/g, "$1*$2").trim();
    if (!t) return [];
    const tokens = t.split(/\s+/);
    const found = [];
    let i = 0;
    while (i < tokens.length) {
      const tk = tokens[i];
      let m = tk.match(/^(\d)\s*=\s*(\d+(?:\*\d+)?)$/);
      if (m) { found.push(makeRoodBlock(m[1], m[2], mode)); i++; continue; }
      if (/^\d$/.test(tk) && i + 1 < tokens.length) {
        const next = tokens[i + 1];
        if (/^\d+(\*\d+)?$/.test(next)) {
          found.push(makeRoodBlock(tk, next, mode)); i += 2; continue;
        }
      }
      return null;
    }
    return found;
  }

  for (const line of lines) {
    const header = detectHeader(line);
    if (header) {
      currentMode = header.mode;
      const rest = line.substring(header.text.length).trim();
      if (rest) {
        const bs = parseLineMaybeMulti(rest, currentMode);
        if (bs === null) return { success: false, error: `❌ ไม่เข้าใจ "${rest}"` };
        bs.forEach(b => blocks.push(b));
      }
      continue;
    }
    if (currentMode) {
      const bs = parseLineMaybeMulti(line, currentMode);
      if (bs === null) return { success: false, error: `❌ ไม่เข้าใจบรรทัด "${line}"` };
      bs.forEach(b => blocks.push(b));
    }
  }

  if (blocks.length === 0) return null;

  // จำกัดเลขรูดรวมสูงสุด 7 ตัวในรูปแบบ header ด้วย
  const uniqueRoodDigits = [...new Set(
    blocks.map(function(b) {
      return String(b.digit || "");
    })
  )];

  if (uniqueRoodDigits.length > 7) {
    return {
      success: false,
      error: "❌ รูดได้ไม่เกิน 6 ตัว"
    };
  }

  const groups = [];
  let totalAmount = 0;
  const uniqueSet = new Set();

  for (const rb of blocks) {
    // ใช้ข้อมูลจากชีตก่อน ถ้าไม่มี/ไม่ครบให้สร้าง 19 ประตูเอง
    // ทำให้ "รูด0=5*5" ทำงานได้แม้ doorMap ไม่มีเลข 0
    const doors = buildRoodDoorsForDigit_(rb.digit, doorMap);

    if (!doors || doors.length === 0) {
      return { success: false, error: `❌ ไม่พบประตูของเลข ${rb.digit}` };
    }

    const numbers = [...new Set(doors)];
    const subtotal = numbers.length * (rb.top + rb.bottom);

    groups.push({
      mode: rb.mode + " " + rb.digit,
      top: rb.top,
      bottom: rb.bottom,
      numbers,
      betCount: numbers.length,
      preserveDuplicates: true,
      subtotal
    });

    totalAmount += subtotal;
    numbers.forEach(n => uniqueSet.add(n));
  }

  return {
    success: true,
    groups,
    totalDoors: uniqueSet.size,
    grossDoors: groups.reduce(function(sum, g) {
      return sum + Number(g.betCount || 0);
    }, 0),
    totalAmount
  };
}

function makeRoodBlock(digit, priceStr, mode) {
  let top, bottom;
  if (priceStr.includes("*")) {
    const s = priceStr.split("*"); top = +s[0]; bottom = +s[1];
  } else {
    top = bottom = +priceStr;
  }

  if (mode === "รูดบ") bottom = 0;
  else if (mode === "รูดล") top = 0;

  return { digit, mode, top, bottom };
}

function extractRoodHeaderLines(text) {
  const normText = text
    .replace(/(\d)\s*[xX]\s*(\d)/g, "$1*$2")
    .replace(/[ \t]*=[ \t]*/g, "=");

  const lines = normText.split(/\n/).map(l => l.trim());
  const roodLines = [];
  const otherLines = [];

  const dataPattern = /^\d=\d+(\*\d+)?$|^\d\s+\d+(\*\d+)?$/;
  const roodHeaders = [
    "รูดบนล่าง","รูดบน ล่าง","รูด บนล่าง","รูด บน ล่าง",
    "รูดบน","รูด บน","รูดล่าง","รูด ล่าง","รูด"
  ];

  function isRoodHeaderLine(line) {
    if (/รูดเบิ้ล/.test(line)) return false;
    for (const h of roodHeaders) {
      if (line === h) return true;
      if (line.startsWith(h + " ") || line.startsWith(h)) {
        const rest = line.substring(h.length).trim();
        if (!rest) return true;
        const tokens = rest.split(/\s+/);
        let i = 0;
        let allMatch = true;
        while (i < tokens.length) {
          const tk = tokens[i];
          if (/^\d=\d+(\*\d+)?$/.test(tk)) { i++; continue; }
          if (/^\d$/.test(tk) && i+1 < tokens.length && /^\d+(\*\d+)?$/.test(tokens[i+1])) {
            i += 2; continue;
          }
          allMatch = false; break;
        }
        return allMatch && tokens.length > 0;
      }
    }
    return false;
  }

  let inRoodSection = false;
  for (const rawLine of lines) {
    if (!rawLine) {
      if (inRoodSection) roodLines.push("");
      else otherLines.push("");
      continue;
    }
    if (isRoodHeaderLine(rawLine)) {
      roodLines.push(rawLine);
      inRoodSection = true;
      continue;
    }
    if (inRoodSection && dataPattern.test(rawLine)) {
      roodLines.push(rawLine);
      continue;
    }
    inRoodSection = false;
    otherLines.push(rawLine);
  }

  return { roodLines, otherLines };
}


// ============================================================
// ราคาอยู่หน้าคำสั่ง 2 ตัว
// รองรับ:
//   23,20,28=3×3บลก  => 23 20 28 บลก 3*3
//   22,33,00=2×2บล   => 22 33 00 บล 2*2
//   12/21=5*5บนล่างกลับ
//
// สำคัญ: แปลงทีละบรรทัด เพื่อให้ราคาของแต่ละชุดปิดกลุ่มของตัวเอง
// บลก/บกล/บนล่างกลับ = กลับเลข
// บล/บนล่าง = ไม่กลับเลข
// ============================================================
function normalizePostfixTwoDigitModePrice_(text) {
  return String(text || "")
    .replace(/[×✕✖⨯]/g, "*")
    .split(/\n/)
    .map(function(rawLine) {
      const line = String(rawLine || "").trim();
      if (!line) return line;

      // จับเฉพาะบรรทัดที่ลงท้ายด้วย mode ชัดเจน
      // group1 = ชุดเลขด้านหน้า
      // group2/3 = ราคาบน/ล่าง
      // group4 = mode
      const m = line.match(
        /^([\d\s,\/\\._-]+?)\s*=\s*(\d+)\s*\*\s*(\d+)\s*(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล|บน\s*ล่าง|บนล่าง|บล)\s*$/i
      );

      if (!m) return line;

      const nums = String(m[1] || "")
        .replace(/[,\/\\._-]+/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      let mode = String(m[4] || "").replace(/\s+/g, "");
      if (mode === "บกล") mode = "บลก";
      if (mode === "บนล่างกลับ") mode = "บลก";
      if (mode === "บนล่าง") mode = "บล";

      return nums + " " + mode + " " + m[2] + "*" + m[3];
    })
    .join("\n");
}

function testPostfixTwoDigitModePrice_() {
  const cases = [
    [
      "23,20,28,26,30,38,36,08,06,86=3×3บลก",
      "23 20 28 26 30 38 36 08 06 86 บลก 3*3"
    ],
    [
      "22,33,00,88,66=2×2บล",
      "22 33 00 88 66 บล 2*2"
    ],
    [
      "12/21=5*5บนล่างกลับ",
      "12 21 บลก 5*5"
    ]
  ];

  cases.forEach(function(c) {
    const got = normalizePostfixTwoDigitModePrice_(c[0]);
    if (got !== c[1]) {
      throw new Error("FAIL: " + c[0] + " => " + got + " | expected: " + c[1]);
    }
  });

  Logger.log("PASS testPostfixTwoDigitModePrice_: " + cases.length + " cases");
}

// ============================================================
// บลก / บกล / บนล่างกลับ + ราคาเดียว
// ราคาเดียวหมายถึง "บนเท่ากับล่าง" เสมอ
// เช่น บลก15 => บลก15*15, บลก 10 => บลก 10*10
// ทำก่อน splitBlocks เพื่อให้ราคานั้นปิดกลุ่มทันทีเหมือนลูกค้าพิมพ์ * มาเอง
// ============================================================
function normalizeTwoSideReturnSinglePrice_(text) {
  return String(text || "")
    .replace(
      // สำคัญ:
      // ต้องจับ "ราคาเต็มทั้งจำนวน" เท่านั้น
      // เดิม (\\d+)(?!\\s*\\*) สามารถ backtrack จาก 15 เหลือ 1
      // ทำให้ บลก15 กลายเป็น 1*1 แล้วเหลือเลข 5 ค้างอยู่
      //
      // กฎใหม่:
      // บลก15    => บลก 15*15
      // บลก 20   => บลก 20*20
      // บลก15*15 => คงเดิม ไม่แก้ซ้ำ
      /(บน\s*ล่าง\s*กลับ|บนล่างกลับ|บลก|บกล)\s*=?\s*(\d+)(?!\d)(?!\s*\*)/gi,
      function(_, mode, price) {
        return mode + " " + price + "*" + price;
      }
    );
}

// ทดสอบกฎราคา บลก ราคาเดียว
// รันฟังก์ชันนี้จาก Apps Script ได้ โดยต้องไม่ throw Error
function testTwoSideReturnSinglePrice_() {
  const cases = [
    ["47/69/87 บลก15", "47/69/87 บลก 15*15"],
    ["10/30/40/50/60/90/80 บลก20", "10/30/40/50/60/90/80 บลก 20*20"],
    ["23 24 25 บลก 10", "23 24 25 บลก 10*10"],
    ["23 24 25 บกล10", "23 24 25 บกล 10*10"],
    ["23 24 25 บนล่างกลับ 25", "23 24 25 บนล่างกลับ 25*25"],
    ["23 24 25 บลก15*15", "23 24 25 บลก15*15"]
  ];

  cases.forEach(function(c) {
    const got = normalizeTwoSideReturnSinglePrice_(c[0]);
    if (got !== c[1]) {
      throw new Error(
        "FAIL: " + c[0] + " => " + got + " | expected: " + c[1]
      );
    }
  });

  Logger.log("PASS testTwoSideReturnSinglePrice_: " + cases.length + " cases");
}


function testNaturalTwoDigitAndRoodForms_() {
  const cases = [
    ["รูด5,6=2×2บลก", "รูด 5,6 บลก 2*2"],
    ["52,51,62,61 2*2 บลก", "52 51 62 61 บลก 2*2"],
    ["บล 20 บาท", "บล 20*20"],
    ["รูด5กับ2ตัวละ5บาท", "รูด 5/2 บล 5*5"],
    ["รูด5กับรูด2ตัวละ5บาท", "รูด 5/2 บล 5*5"],
    ["รูด 5 กับ รูด 2 ตัวละ 5 บาท", "รูด 5/2 บล 5*5"],
    ["6-7 20*20", "รูด 6/7 บล 20*20"],
    ["66-99", "66-99"]
  ];

  cases.forEach(function(c) {
    const got = normalizeNaturalTwoDigitAndRoodForms_(c[0]);
    if (got !== c[1]) {
      throw new Error("FAIL: " + c[0] + " => " + got + " | expected " + c[1]);
    }
  });

  Logger.log("PASS testNaturalTwoDigitAndRoodForms_: " + cases.length + " cases");
}

// ================== SPLIT BLOCKS ==================
function splitBlocks(text) {
  text = normalizeTwoSideReturnSinglePrice_(text);

  let pre = text
    // safety-net: ถ้า normalize รอบแรกหลุด ให้ บ-ล / บ/ล / บ ล = บนล่าง
    .replace(/(^|[^ก-๙a-zA-Z])บ\s*[-\/\.\\]\s*ล(?=$|[^ก-๙a-zA-Z])/g, "$1บนล่าง")
    .replace(/(^|[^ก-๙a-zA-Z])บ\s+ล(?=$|[^ก-๙a-zA-Z])/g, "$1บนล่าง")
    .replace(/[,]+/g, " ")
    .replace(/บาท/g, " ")
    .replace(/[฿$¥€£]/g, " ")
    .replace(/[():]/g, " ")
    .replace(/\./g, " ")
    .replace(/ตัวตรง/g, "ตรง")
    .replace(/([36])ก(\d)/g, "$1*$2")
    .replace(/บน\s*[-\/]\s*ล่าง/g, "บนล่าง")
    .replace(/วิน\s*(?:สอง|2)/g, "วิน")
    .replace(/(วินสอง|วิน)(บนล่างกลับ|บนล่าง|บลก|บกล|บล|กลับ|บน|ล่าง|บ|ล)/g, "$1 $2")
    .replace(/(รูดเบิ้ล|เบิ้ลหน้า|เบิ้ลหลัง|เบิ้ล|ตอง|รูดบนล่าง|รูดบน|รูดล่าง|รูด)(บนล่างกลับ|บนล่าง|บลก|บกล|บล|บน|ล่าง|ล|บ)/g, "$1 $2")
    .replace(/([ก-๙])(\d)/g, "$1 $2")
    .replace(/(\d)([ก-๙])/g, "$1 $2");

  const lines = pre.split(/\n/).map(l => l.replace(/\s+/g," ").trim());

  const blocks = [];
  let cur = null;
  let activeHeader = null;

  function flush() {
    if (cur && (cur.body.length > 0 || cur.header)) blocks.push(cur);
    cur = null;
  }

  function discardEmptyHeader() {
    if (cur && cur.header && !cur.priceStr && cur.body.length === 0) {
      cur = null;
      return true;
    }
    return false;
  }

  function normalizeLineSeparators(line) {
    const trimmed = String(line || "").trim();
    // รูป 66-99 คือ "เลข 66 และ 99" ไม่ใช่ 66 = ราคา 99
    // คง special case เดิมไว้เฉพาะเลขหลักเดียวเท่านั้น
    if (/^\d\s*[-\/]\s*\d+(?:\*\d+)?$/.test(trimmed)) {
      return trimmed.replace(/^\s*(\d)\s*[-\/]\s*(\d+(?:\*\d+)?)\s*$/, "$1 = $2");
    }
    let out = trimmed.replace(/(\d)\s*[-\/]+\s*(\d)/g, "$1 $2");
    let prev;
    do { prev = out; out = out.replace(/(\d)[\-\_\/]+(\d)/g, "$1 $2"); } while (out !== prev);
    return out;
  }

  for (let rawLine of lines) {
    if (!rawLine) {
      flush();
      continue;
    }

    let line = normalizeLineSeparators(rawLine)
      .replace(/(\d)\s*[xX]\s*(\d)/g, "$1*$2")
      .replace(/(\d)=(\d)/g, "$1 = $2")
      .replace(/\s+/g, " ")
      .trim();

    const shortReturn = detectReturnShorthand(line);
    if (shortReturn) {
      if (!discardEmptyHeader()) flush();
      cur = {
        header: shortReturn.header,
        body: [shortReturn.cleaned],
        priceStr: shortReturn.priceStr
      };
      flush();
      continue;
    }

    const modeInfo = findModeInLine(line);
    const priceMatch = extractPrice(line, modeInfo);

    const mode = modeInfo ? modeInfo.keyword : null;

if (mode && priceMatch) {

  const cleanedHasNumbers = /\d/.test(
    priceMatch.cleaned.replace(
      new RegExp(escapeRegExp(mode), "g"),
      ""
    )
  );


  // =====================================================
  // MODE + PRICE แต่ไม่มีเลขในบรรทัดเดียวกัน
  //
  // รองรับทั้ง 2 แบบ:
  // 1) เลขมาก่อน แล้วปิดท้ายด้วยราคา/โหมด
  //    35-37-39
  //    =1*1กลับ
  //
  // 2) โหมด+ราคามาก่อน แล้วเลขอยู่บรรทัดถัดไป
  //    หกกลับ 20 บาท
  //    878-839-783
  //    บลก 30 บาท
  //    87-88-37-38-39-89-79
  // =====================================================
  if (!cleanedHasNumbers) {

    // มีเลขค้างอยู่และยังไม่มีราคา -> บรรทัดนี้คือราคาปิดบล็อกเดิม
    if (cur && !cur.priceStr && cur.body && cur.body.length > 0) {
      cur.header = mode;
      cur.priceStr = priceMatch.priceStr;
      cur.isHeaderOnly = true;
      cur.explicitLinePrice = !!priceMatch.explicitLinePrice;

      // HARD LOCK โหมด 3 ตัวลงใน body ด้วย
      // เช่นเลขหลายตัว แล้วบรรทัดถัดไปเป็น "โต๊ด 1"
      // ต้องให้ body มีคำว่า "โต๊ด" เพื่อกัน parser เก่าตกไป default ตรง+โต๊ด
      const hardMode = canonicalMode(mode);
      if (["โต๊ด", "ตรง", "ตท", "6กลับ", "3กลับ"].includes(hardMode)) {
        const modeTextForBody = hardMode === "ตท" ? "ตรงโต๊ด" : hardMode;
        const joinedBody = cur.body.join(" ");
        if (joinedBody.indexOf(modeTextForBody) === -1) {
          cur.body.push(modeTextForBody);
        }
      }

      flush();
      activeHeader = null;
      continue;
    }

    // บล็อกก่อนหน้าจบแล้ว ให้เริ่ม header ใหม่รอเลขบรรทัดถัดไป
    if (cur) flush();

    cur = {
      header: mode,
      body: [priceMatch.cleaned || mode],
      priceStr: priceMatch.priceStr,
      isHeaderOnly: true,
      explicitLinePrice: !!priceMatch.explicitLinePrice
    };

    activeHeader = mode;
    continue;
  }


  // =====================================================
  // MODE + ตัวเลข + ราคา
  //
  // รองรับเลขที่ต่อมาจากบรรทัดก่อนหน้า เช่น
  // 35-37-39-57-59-
  // 79=1*1กลับ
  //
  // บรรทัดบนยังไม่มีราคา จึงต้องรวมกับบรรทัดนี้เป็นบล็อกเดียว
  // ไม่เช่นนั้นจะเกิดบล็อกเลขที่ไม่มีราคาและ parser จะหยุดอ่าน
  // =====================================================
  if (cur && !cur.priceStr && cur.body && cur.body.length > 0 && isTrailingMode(mode)) {
    cur.header = mode;
    cur.body.push(priceMatch.cleaned);
    cur.priceStr = priceMatch.priceStr;
    cur.isHeaderOnly = false;
    cur.explicitLinePrice = !!priceMatch.explicitLinePrice;
    activeHeader = null;
    continue;
  }

  // โหมดประเภทอื่น เช่น รูด 1/7=1 ให้จบบล็อกของตัวเองตามเดิม
  flush();

  cur = {
    header: mode,
    body: [priceMatch.cleaned],
    priceStr: priceMatch.priceStr,
    isHeaderOnly: false,
    explicitLinePrice: !!priceMatch.explicitLinePrice
  };

  continue;
}

    if (mode && !priceMatch) {
      const canonical = canonicalMode(mode);
      const lineWithoutMode = line
        .replace(new RegExp(escapeRegExp(mode), "g"), " ")
        .replace(/\s+/g, " ")
        .trim();
      const nums3 = lineWithoutMode.match(/\b\d{3}\b/g) || [];

      // =====================================================
      // 3 ตัว 6กลับ แต่ไม่ได้ใส่ราคาในบรรทัด
      // เช่น: 878-839-783 6กลับ
      // ให้ใช้ราคาเริ่มต้น 20 บาท และปิดบล็อกทันที
      // ป้องกันราคา 30*30 ของบล็อก 2 ตัวถัดไปไหลย้อนมาใช้กับ 6กลับ
      // =====================================================
      if (canonical === "6กลับ" && nums3.length > 0) {
        flush();

        blocks.push({
          header: "6กลับ",
          body: [nums3.join(" ")],
          priceStr: String(DEFAULT_6_RETURN_PRICE),
          isHeaderOnly: false,
          explicitLinePrice: false
        });

        activeHeader = null;
        continue;
      }

      const lineHasNumbers = hasBetNumberInLine(lineWithoutMode);

      // =====================================================
      // MODE เดี่ยวหลังบล็อกที่มีราคาแล้ว = modifier ของบล็อกก่อนหน้า
      // เช่น
      // 81-89-87-85-83
      // 19-17-15-13=2*2
      // กลับ
      //
      // ต้องหมายถึงเลขทั้งสองบรรทัด "บนล่างกลับ 2*2"
      // ไม่ใช่เอาคำว่า กลับ ไปเป็น header ของบล็อกถัดไป
      // =====================================================
      if (!lineHasNumbers && cur && cur.priceStr && isTrailingMode(mode)) {
        cur.header = mode;
        cur.isHeaderOnly = true;
        flush();
        activeHeader = null;
        continue;
      }

      if (!lineHasNumbers) activeHeader = mode;

      if (!lineHasNumbers && cur) {
        // mode ก่อนราคาหรือก่อนเลขชุดถัดไป = header ของบล็อกปัจจุบัน/ถัดไป
        if (cur.priceStr) {
          flush();
          cur = {
            header: mode,
            body: [],
            priceStr: null,
            isHeaderOnly: true
          };
        } else {
          cur.header = mode;
          cur.isHeaderOnly = true;
        }
        continue;
      }
      if (!lineHasNumbers && !cur && blocks.length > 0 && isTrailingMode(mode)) {
        blocks[blocks.length - 1].header = mode;
        blocks[blocks.length - 1].isHeaderOnly = true;
        continue;
      }
      if (!cur) cur = { header: mode, body: lineHasNumbers ? [line] : [], priceStr: null, isHeaderOnly: !lineHasNumbers };
      else {
        if (!lineHasNumbers) {
          flush();
          cur = { header: mode, body: [], priceStr: null, isHeaderOnly: true };
        } else {
          cur.body.push(line);
        }
      }
      continue;
    }

    if (!mode && priceMatch) {
      const cleanedPriceBody = String(priceMatch.cleaned || "").trim();
      const priceLineHasNumbers = hasBetNumberInLine(cleanedPriceBody);
      const hasNewMode = !!findModeInLine(cleanedPriceBody);

      // =====================================================
      // ราคาอยู่ "บรรทัดเดี่ยว" หลังชุดเลข = ปิดกลุ่มทันที
      //
      // ตัวอย่าง:
      // 19
      // 91
      // 50*50
      // 17
      // 71
      // 15*15
      //
      // ของเดิมตั้งราคาให้ cur แต่ยังไม่ flush ทำให้เลขชุดถัดไป
      // ไหลเข้าไปรวมกับราคา 50*50 จนยอดผิด
      // =====================================================
      if (
        cur &&
        !cur.priceStr &&
        cur.body &&
        cur.body.length > 0 &&
        !priceLineHasNumbers &&
        !hasNewMode
      ) {
        cur.priceStr = priceMatch.priceStr;
        cur.explicitLinePrice = true;
        flush();
        activeHeader = null;
        continue;
      }

      if (cur) {
        if (cur.priceStr && (priceLineHasNumbers || hasNewMode)) {
          flush();
          cur = {
            header: activeHeader,
            body: cleanedPriceBody ? [cleanedPriceBody] : [],
            priceStr: priceMatch.priceStr,
            isHeaderOnly: false,
            explicitLinePrice: !!priceMatch.explicitLinePrice
          };
        } else {
          if (cleanedPriceBody) cur.body.push(cleanedPriceBody);
          if (!cur.priceStr) cur.priceStr = priceMatch.priceStr;
          if (priceMatch.explicitLinePrice) cur.explicitLinePrice = true;
        }
      } else {
        cur = {
          header: activeHeader,
          body: cleanedPriceBody ? [cleanedPriceBody] : [],
          priceStr: priceMatch.priceStr,
          isHeaderOnly: false,
          explicitLinePrice: !!priceMatch.explicitLinePrice
        };
      }
      continue;
    }

    if (cur && cur.explicitLinePrice) flush();
    if (!cur) cur = { header: activeHeader, body: [line], priceStr: null, isHeaderOnly: false };
    else cur.body.push(line);
  }
  flush();

  return blocks;
}

function hasBetNumberInLine(line) {
  const nums = String(line || "").match(/\b\d+\b/g) || [];
  return nums.some(n => n.length >= 2);
}

function detectReturnShorthand(line) {
  let m = line.match(/^(.+?)\*\s*([36])\s*กลับ\s*(\d+)\s*$/);
  if (m) {
    const numsPart = m[1].trim();
    const kind = m[2];
    const price = m[3];
    if (/\d{3}/.test(numsPart)) {
      return { header: kind + "กลับ", cleaned: numsPart, priceStr: price };
    }
  }

  m = line.match(/^(.+?)\*\s*([36])\s*=\s*(\d+)(?:\*\d+)*\s*$/);
  if (m) {
    const numsPart = m[1].trim();
    const kind = m[2];
    const price = m[3];
    if (/\d{3}/.test(numsPart)) {
      return { header: kind + "กลับ", cleaned: numsPart, priceStr: price };
    }
  }
  m = line.match(/^(.+?)=\s*(\d+)(?:\*(\d+)){2,}\s*$/);
  if (m) {
    const numsPart = m[1].trim();
    const price = m[2];
    if (/\d{3}/.test(numsPart)) {
      return { header: "6กลับ", cleaned: numsPart, priceStr: price };
    }
  }
  m = line.match(/^(.+?)\s+([36])\s*[\*xXก]\s*(\d+)\s*$/);
  if (m) {
    const numsPart = m[1].trim();
    const kind = m[2];
    const price = m[3];
    if (/\d{3}/.test(numsPart) && numsPart.indexOf("=") === -1) {
      return { header: kind + "กลับ", cleaned: numsPart, priceStr: price };
    }
  }
  return null;
}

function isTrailingMode(mode) {
  const m = canonicalMode(mode);
  return ["บลก", "ก", "บล", "บ", "ล", "6กลับ", "3กลับ"].includes(m);
}

function canonicalMode(mode) {
  if (!mode) return "";
  let compact = String(mode).replace(/\s+/g, "");

  // shorthand บ-ล / บ/ล / บ.ล. = บนล่าง
  if (/^บ[-\/\.]ล$/.test(compact)) compact = "บล";

  if (compact === "หาม") return "หาม";
  if (compact === "บกล") return "บลก";
  if (compact === "บนล่างกลับ") return "บลก";
  if (compact === "บนล่าง") return "บล";
  if (compact === "กลับ") return "ก";
  if (compact === "บน") return "บ";
  if (compact === "ล่าง") return "ล";
  if (compact === "6กลับ") return "6กลับ";
  if (compact === "3กลับ") return "3กลับ";

  // =====================================================
  // 3 ตัว: ล็อกโหมดให้ชัดเจน
  // "โต๊ด" / "โต๊ดอย่างเดียว" = โต๊ดอย่างเดียว
  // ห้ามตกกลับไป default "ตรง+โต๊ด"
  // =====================================================
  if (
    compact === "โต๊ด" ||
    compact === "โต๊ดอย่างเดียว" ||
    compact === "โต๊ดอย่างเดี่ยว" ||
    compact === "โต๊ดอย่างเดี๋ยว"
  ) return "โต๊ด";

  if (compact === "ตรง") return "ตรง";
  if (compact === "ตท" || compact === "ตรงโต๊ด" || compact === "บนตรงโต๊ด" || compact === "บต") return "ตท";

  return compact;
}

function findModeInLine(line) {
  const MK = [
    "บนล่างกลับ","บน ล่าง กลับ","บนล่าง กลับ","บนล่าง","บน ล่าง","บลก","บกล","บล","ล่าง","กลับ","บน","บ","ล",
    "ตรงโต๊ด","ตรง โต๊ด","บนตรงโต๊ด","บต","6กลับ","6 กลับ","3กลับ","3 กลับ","โต๊ด","ตรง",
    "วินสอง","วิน",
    "รูดเบิ้ล","เบิ้ลหน้า","เบิ้ลหลัง","เบิ้ล","ตอง",
    "รูดบนล่าง","รูดบน","รูดล่าง","รูด","หาม"
  ].sort((a,b)=>b.length-a.length);

  for (const k of MK) {
    const esc = k.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
    const re = new RegExp("(^|[^ก-๙])" + esc + "(?![ก-๙])");
    const m = line.match(re);
    if (m) {
      const start = m.index + (m[1] ? m[1].length : 0);
      return { keyword: k, index: start, length: k.length };
    }
  }
  return null;
}

function extractPrice(line, modeInfo) {
  let t = line.replace(/\s+/g, " ").trim();

  // ถ้ามี "=" หลายตัว เช่น
  // กับ=365/368/.../081=1บ
  // "=" ตัวแรกเป็นตัวเปิดรายการ ไม่ใช่ราคา
  // ให้ใช้ "=ราคา" ตัวสุดท้ายเป็นราคาเสมอ
  const eqRe = /=\s*(\d+(?:\*\d+)?)/g;
  let eqMatch = null;
  let scan;
  while ((scan = eqRe.exec(t)) !== null) {
    eqMatch = scan;
  }

  if (eqMatch) {
    const before = t.substring(0, eqMatch.index);
    const after = t.substring(eqMatch.index + eqMatch[0].length);
    const cleaned = (before + " " + after)
      .replace(/\s+/g, " ")
      .trim();

    return {
      priceStr: eqMatch[1],
      cleaned: cleaned,
      explicitLinePrice: true
    };
  }

  let m = t.match(/(\d+\*\d+)/);
  if (m) {
    const cleaned = t.replace(m[1], " ").replace(/\s+/g, " ").trim();
    return { priceStr: m[1], cleaned, explicitLinePrice: false };
  }

  if (modeInfo) {
    const afterMode = t.substring(modeInfo.index + modeInfo.length);
    const modeText = t.substring(modeInfo.index, modeInfo.index + modeInfo.length);
    if (/^(รูดเบิ้ล|เบิ้ล|ตอง)$/.test(modeText)) {
      // รองรับ "รูดเบิ้ล บล 1" / "รูดเบิ้ล/บล1" หลัง normalize
      const sideThenPrice = afterMode.match(/^\s*(บลก|บกล|บล|บนล่างกลับ|บนล่าง|บน|ล่าง|บ|ล)\s*=?\s*(\d+(?:\*\d+)?)\s*$/);
      if (sideThenPrice) {
        const sideMode = sideThenPrice[1];
        const priceVal = sideThenPrice[2];
        const cleaned = (
          t.substring(0, modeInfo.index + modeInfo.length) + " " + sideMode
        ).replace(/\s+/g, " ").trim();
        return { priceStr: priceVal, cleaned, explicitLinePrice: false };
      }

      const middlePrice = afterMode.match(/^\s*(\d+(?:\*\d+)?)\s*(บลก|บกล|บล|บนล่างกลับ|บนล่าง|บน|ล่าง|บ|ล)?\s*$/);
      if (middlePrice) {
        const priceVal = middlePrice[1];
        const cleaned = (t.substring(0, modeInfo.index + modeInfo.length) + " " + afterMode.replace(priceVal, " ")).replace(/\s+/g, " ").trim();
        return { priceStr: priceVal, cleaned, explicitLinePrice: false };
      }
    }
    const tightMatch = afterMode.match(/^\s*(\d+)\s*$/);
    if (tightMatch) {
      const priceVal = tightMatch[1];
      const cleaned = (t.substring(0, modeInfo.index + modeInfo.length) + " " + afterMode.replace(priceVal, " ")).replace(/\s+/g, " ").trim();
      return { priceStr: priceVal, cleaned, explicitLinePrice: false };
    }
    const tokens = afterMode.trim().split(/\s+/).filter(x=>x);
    if (tokens.length > 0) {
      const last = tokens[tokens.length-1];
      if (/^\d+$/.test(last)) {
        const cleaned = t.replace(new RegExp("\\s*" + last + "\\s*$"), "").trim();
        return { priceStr: last, cleaned, explicitLinePrice: false };
      }
    }
  }

  return null;
}

// ================== PARSE BLOCK ==================
function parseBlock(body, priceStr, keywords, doorMap, blockHeader, isHeaderOnly) {
  body = body.trim();
  blockHeader = canonicalMode(blockHeader);
  if (!body) {
    return { success: false, error: "❌ บล็อกว่าง" };
  }

  let top, bottom;
  const starParts = priceStr.split("*");
  if (starParts.length === 1) {
    top = bottom = +starParts[0];
  } else if (starParts.length === 2) {
    top = +starParts[0]; bottom = +starParts[1];
  } else {
    top = bottom = +starParts[0];
  }

  const sidePrice = applyInlineSideMode(body, top, bottom);
  top = sidePrice.top;
  bottom = sidePrice.bottom;

  if (blockHeader === "ตอง" || /(^|[^ก-๙])ตอง(?![ก-๙])/.test(body)) {
    return parseTongBlock(top);
  }

  if (blockHeader === "6กลับ" || blockHeader === "3กลับ") {
    return parseReturn3Block(body, blockHeader, top);
  }

  if (blockHeader === "หาม" || /(^|[^ก-๙])หาม(?![ก-๙])/.test(body)) {
    return parseHamBlock(top);
  }

  // =====================================================
  // ชุดเลข 3 ตัวสำเร็จ: แทง "บนตรง" เท่านั้น
  // เบิ้ลหน้า 1 = 90 ตัว x 1 = 90 บาท
  // เบิ้ลหลัง 1 = 90 ตัว x 1 = 90 บาท
  // =====================================================
  if (blockHeader === "เบิ้ลหน้า" || /เบิ้ลหน้า/.test(body)) {
    return parseBeilNa(body, top);
  }
  if (blockHeader === "เบิ้ลหลัง" || /เบิ้ลหลัง/.test(body)) {
    return parseBeilLang(body, top);
  }

  // "เบิ้ล" เฉย ๆ / รูดเบิ้ล ยังคงเป็นเลข 2 ตัวตามระบบเดิม
  if (/รูดเบิ้ล/.test(body) || /(^|[^ก-๙])เบิ้ล(?!หน้า|หลัง|[ก-๙])/.test(body)) {
    return parseRoodBeil(body, top, bottom);
  }

  const roodKey = findKeyword(body, keywords) || (/รูด/.test(body) ? "รูด" : null);
  if (roodKey) return parseRoodBlock(body, roodKey, top, bottom, doorMap);

  const singleRoodDigit = detectSingleDigitRood(body, blockHeader);
  if (singleRoodDigit) {
    return parseRoodBlock("รูด " + singleRoodDigit, "รูด", top, bottom, doorMap);
  }

  // วินเป็นเลข 2 ตัวที่สร้างจากชุดตัวเลขต้นทาง
  // ต้องเช็กก่อน has3DigitNumber() ไม่เช่นนั้น "493 วิน ..." จะถูกส่งไป parser 3 ตัว
  if (hasWin2Keyword(body)) {
    return parseDirect2Block(body, top, bottom, blockHeader, priceStr, isHeaderOnly);
  }

  if (has3DigitNumber(body) && has2DigitNumber(body)) {
    return parseMixedNumberBlock(body, top, bottom, blockHeader, priceStr, isHeaderOnly);
  }

  if (has3DigitNumber(body)) {
    return parseDirect3Block(body, top, bottom, priceStr, blockHeader);
  }

  return parseDirect2Block(body, top, bottom, blockHeader, priceStr, isHeaderOnly);
}

function applyInlineSideMode(body, top, bottom) {
  const t = " " + String(body || "").replace(/\s+/g, " ").trim() + " ";
  if (/( บนล่างกลับ | บนล่าง | บลก | บกล | บล )/.test(t)) {
    return { top, bottom };
  }
  if (/( บน | บ )/.test(t)) {
    return { top, bottom: 0 };
  }
  if (/( ล่าง | ล )/.test(t)) {
    return { top: 0, bottom };
  }
  return { top, bottom };
}

function hasWin2Keyword(body) {
  const t = String(body || "")
    // วินสอง / วิน 2 / วิน2 = วินเลข 2 ตัวเหมือนคำว่า "วิน"
    .replace(/วิน\s*(?:สอง|2)/g, "วิน")
    // กันกรณีพิมพ์ติดกับเลขหรือโหมด เช่น 45967วินบลก1
    .replace(/(\d)(วิน)/g, "$1 $2")
    .replace(/(วิน)(?=บลก|บกล|บล|บนล่างกลับ|บนล่าง|กลับ|บน|ล่าง|บ|ล)/g, "$1 ")
    .replace(/\s+/g, " ")
    .trim();

  return /(^|\s)วิน(?=\s|$)/.test(t);
}

function has3DigitNumber(body) {
  const t = body.replace(/[\-\_\/\.\=,]+/g, " ");
  const tokens = t.split(/\s+/);
  for (const tk of tokens) {
    if (/^\d{3}$/.test(tk)) return true;
  }
  return false;
}

function has2DigitNumber(body) {
  const t = body.replace(/[\-\_\/\.\=,]+/g, " ");
  const tokens = t.split(/\s+/);
  for (const tk of tokens) {
    if (/^\d{2}$/.test(tk)) return true;
  }
  return false;
}

function detectSingleDigitRood(body, blockHeader) {
  if (blockHeader) return null;
  const cleaned = body
    .replace(/[\-\_\/\.\=,]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return /^\d$/.test(cleaned) ? cleaned : null;
}

function parseMixedNumberBlock(body, top, bottom, blockHeader, priceStr, isHeaderOnly) {
  const t = body.replace(/[\-\_\/\.\=,]+/g, " ").replace(/\s+/g, " ").trim();
  const tokens = t.split(/\s+/).filter(x => x);
  const nums3 = [];
  const nums2 = [];

  for (const tk of tokens) {
    if (/^\d{3}$/.test(tk)) nums3.push(tk);
    else if (/^\d{2}$/.test(tk)) nums2.push(tk);
  }

  const groups = [];
  if (nums3.length > 0) {
    const p3 = parseDirect3Block(nums3.join(" ") + extract3DigitModeText(body), top, bottom, priceStr, blockHeader);
    if (!p3.success) return p3;
    groups.push(p3);
  }

  if (nums2.length > 0) {
    const p2 = parseDirect2Block(nums2.join(" ") + extract2DigitModeText(body), top, bottom, blockHeader, priceStr, isHeaderOnly);
    if (!p2.success) return p2;
    groups.push(p2);
  }

  return { success: true, groups };
}

function extract3DigitModeText(body) {
  const keys = ["6กลับ", "6 กลับ", "3กลับ", "3 กลับ", "ตรงโต๊ด", "ตรง โต๊ด", "บนตรงโต๊ด", "บต", "โต๊ด", "ตรง"];
  for (const k of keys) {
    if (body.indexOf(k) !== -1) return " " + k;
  }
  return "";
}

function extract2DigitModeText(body) {
  const keys = ["บนล่างกลับ", "บน ล่าง กลับ", "บนล่าง กลับ", "บลก", "บกล", "กลับ", "บนล่าง", "บน ล่าง", "ล่าง", "บน", "บล", "บ", "ล"];
  for (const k of keys) {
    if (body.indexOf(k) !== -1) return " " + k;
  }
  return "";
}

// สร้างรูด 19 ประตูต่อ 1 หลักแบบมาตรฐาน
// เช่น รูด 6 = 60-69 จำนวน 10 ตัว + 06,16,26,36,46,56,76,86,96 จำนวน 9 ตัว
// รวม 19 รายการเสมอ
function buildRoodDoorsForDigit_(digit, doorMap) {
  const d = String(digit || "");
  if (!/^\d$/.test(d)) return [];

  // ถ้าชีตมีข้อมูลถูกต้อง 19 ตัว ใช้ข้อมูลจากชีตก่อน
  const fromSheet = ((doorMap && doorMap[d]) || [])
    .map(function(n) { return String(n || "").trim(); })
    .filter(function(n) { return /^\d{2}$/.test(n); });

  const sheetUnique = [...new Set(fromSheet)];
  if (sheetUnique.length === 19) {
    return sheetUnique;
  }

  // fallback สร้าง 19 ประตูเอง ป้องกันข้อมูลชีตซ้ำ/ขาด
  const result = [];

  for (let x = 0; x <= 9; x++) {
    result.push(d + String(x));
  }

  for (let x = 0; x <= 9; x++) {
    const xStr = String(x);
    if (xStr === d) continue; // dd มีแล้วด้านบน
    result.push(xStr + d);
  }

  return result;
}

function parseRoodBlock(body, roodKey, top, bottom, doorMap) {
  let t = String(body || "");

  // ลบคำว่า "รูด" แล้วอ่านเฉพาะ selector ที่อยู่หน้าราคา/โหมด
  // เช่น รูด6/9บล1 -> selector = 6,9 เท่านั้น
  const keyEsc = escapeRegExp(roodKey || "รูด");
  t = t.replace(new RegExp(keyEsc), " ").trim();

  t = t.replace(/^(?:บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล|บนล่าง|บน\s*ล่าง|บล|บน|ล่าง|บ|ล)\s*/, "");

  const stopRe = /(?:บนล่างกลับ|บน\s*ล่าง\s*กลับ|บลก|บกล|บนล่าง|บน\s*ล่าง|บล|กลับ|บน|ล่าง|บ|ล)/;
  const stopMatch = t.match(stopRe);
  if (stopMatch && typeof stopMatch.index === "number") {
    t = t.substring(0, stopMatch.index).trim();
  }

  if (t.indexOf("=") !== -1) {
    t = t.substring(0, t.indexOf("=")).trim();
  }

  const numberChunks = t.match(/\d+/g) || [];
  let digits = [];

  numberChunks.forEach(function(chunk) {
    if (chunk.length === 1) {
      digits.push(chunk);
    } else if (chunk.length <= 4) {
      // รูด69 / รูด691
      digits = digits.concat(chunk.split(""));
    }
  });

  const uniq = [...new Set(digits)].filter(function(d) {
    return /^\d$/.test(d);
  });

  if (uniq.length === 0) {
    return { success: false, error: "❌ รูด: ไม่มีเลข" };
  }

  // PATCH 2026-09-23:
  // รองรับรูดสูงสุด 7 ตัว
  // ตามข้อความที่กำหนด ถ้าเกินให้แจ้ง "รูดได้ไม่เกิน 6 ตัว"
  if (uniq.length > 7) {
    return { success: false, error: "❌ รูดได้ไม่เกิน 6 ตัว" };
  }

  // สำคัญ:
  // ห้าม dedupe ข้ามเลขรูดแต่ละหลัก
  // รูด 6/9 = 19 + 19 = 38 รายการ
  // แม้ 69 และ 96 จะอยู่ทั้งสองชุด ก็ต้องคิดเดิมพันทั้งสองครั้ง
  let all = [];

  uniq.forEach(function(d) {
    const oneDoorSet = buildRoodDoorsForDigit_(d, doorMap);
    all = all.concat(oneDoorSet);
  });

  if (all.length === 0) {
    return { success: false, error: "❌ ไม่พบเลขประตู" };
  }

  const betCount = all.length;
  const subtotal = betCount * ((Number(top) || 0) + (Number(bottom) || 0));

  return {
    success: true,
    numbers: all,
    top: Number(top) || 0,
    bottom: Number(bottom) || 0,
    mode: "รูด " + uniq.join("-"),

    // จำนวนรายการแทงจริง ไม่ใช่จำนวนเลข unique
    betCount: betCount,
    preserveDuplicates: true,
    subtotal2: subtotal
  };
}

function parseReturn3Block(body, kind, price) {
  let t = body.replace(/[\-\_\/\.\=,]+/g, " ").replace(/\s+/g, " ").trim();
  const rawNums = t.split(/\s+/).filter(x => /^\d{3}$/.test(x));
  if (rawNums.length === 0) return { success: false, error: "❌ ไม่พบเลข 3 ตัว" };

  let finalNumbers = [];
  let subtotal = 0;
  rawNums.forEach(num => {
    const variants = perms3(num);
    finalNumbers = finalNumbers.concat(variants);
    subtotal += variants.length * price;
  });
  finalNumbers = [...new Set(finalNumbers)];

  return {
    success: true,
    numbers: finalNumbers,
    top: price, bottom: 0,
    mode: kind,
    subtotal3: subtotal
  };
}

function parseHamBlock(price) {
  return {
    success: true,
    numbers: [...HAM_NUMBERS],
    top: price,
    bottom: 0,
    mode: "หาม",
    subtotal3: HAM_NUMBERS.length * price
  };
}

function parseTongBlock(price) {
  const numbers = ["111","222","333","444","555","666","777","888","999","000"];
  return {
    success: true,
    numbers,
    top: price,
    bottom: 0,
    mode: "ตอง",
    subtotal3: numbers.length * price
  };
}

function parseDirect3Block(body, top, bottom, origPriceStr, blockHeader) {
  let t = body.replace(/[\_\/\.\=,]+/g, " ").replace(/\s+/g, " ").trim();

  const MK3 = [
    { text: "ตรงโต๊ด",  mode: "ตท" },
    { text: "ตรง โต๊ด", mode: "ตท" },
    { text: "บนตรงโต๊ด", mode: "ตท" },
    { text: "บต",       mode: "ตท" },
    { text: "6กลับ",    mode: "6กลับ" },
    { text: "6 กลับ",   mode: "6กลับ" },
    { text: "3กลับ",    mode: "3กลับ" },
    { text: "3 กลับ",   mode: "3กลับ" },
    { text: "โต๊ด",      mode: "โต๊ด" },
    { text: "ตรง",      mode: "ตรง" }
  ];

  // =====================================================
  // ถ้า splitBlocks() ระบุ header มาแล้ว ให้ header มีสิทธิ์สูงสุด
  // ตัวอย่าง:
  // 932/930/.../384
  // โต๊ด 1
  // => blockHeader = "โต๊ด" จึงต้องเป็นโต๊ดอย่างเดียวแน่นอน
  // =====================================================
  const headerMode = canonicalMode(blockHeader);
  const hasForcedHeaderMode = ["ตรง", "โต๊ด", "ตท", "6กลับ", "3กลับ"].includes(headerMode);

  let mode = hasForcedHeaderMode ? headerMode : null;
  let modeIdx = -1, modeLen = 0;

  // ถ้าไม่มี header ชัดเจน ค่อยหา mode ที่อยู่ใน body
  if (!mode) {
    for (const mk of MK3) {
      const esc = mk.text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const re = new RegExp("(^|\\s)" + esc + "(\\s|$)");
      const m = t.match(re);
      if (m) {
        mode = mk.mode;
        modeIdx = m.index + (m[1] ? m[1].length : 0);
        modeLen = mk.text.length;
        break;
      }
    }
  }

  let numbersText;

  // ถ้า mode ถูกเจอจาก body ให้ตัดคำ mode ออกจากชุดเลข
  if (modeIdx >= 0) {
    numbersText = (t.substring(0, modeIdx) + " " + t.substring(modeIdx + modeLen)).trim();
  } else {
    numbersText = t;
  }

  // ไม่มี mode ทั้ง header และ body จึงค่อยใช้ fallback เดิม
  if (!mode) {
    if (origPriceStr && origPriceStr.split("*").length >= 3) {
      mode = "6กลับ";
    } else {
      mode = "ตท";
    }
  }

  const rawNums = numbersText.split(/\s+/).filter(x => /^\d{3}$/.test(x));
  if (rawNums.length === 0) return { success: false, error: "❌ ไม่พบเลข 3 ตัว" };

  // 3 ตัวตรง+โต๊ด ถ้าระบุราคาเดียว ให้ราคา "ตรง" และ "โต๊ด" เท่ากัน
  // เช่น 687 817 ... =1  => ตรง 1 + โต๊ด 1
  // ป้องกันกรณี parser ก่อนหน้าเหลือ bottom = 0 แต่ subtotal ถูกคิดเป็น 2 ฝั่ง
  if (mode === "ตท" && Number(top) > 0 && Number(bottom) <= 0) {
    bottom = Number(top);
  }

  let finalNumbers = [];
  let subtotal = 0;

  rawNums.forEach(num => {
    if (mode === "ตรง") {
      finalNumbers.push(num);
      subtotal += top;
    } else if (mode === "โต๊ด") {
      finalNumbers.push(num);
      subtotal += top;
    } else if (mode === "ตท") {
      finalNumbers.push(num);
      subtotal += top + bottom;
    } else if (mode === "6กลับ" || mode === "3กลับ") {
      const variants = perms3(num);
      finalNumbers = finalNumbers.concat(variants);
      subtotal += variants.length * top;
    }
  });

  // PATCH 2026-09-21:
  // ห้ามตัดเลข 3 ตัวซ้ำ ลูกค้าพิมพ์ซ้ำกี่ครั้งให้นับทุกครั้ง
  // เช่น 523 826 763 465 826 ต้องเก็บ 826 ทั้ง 2 รายการ
  return {
    success: true,
    numbers: finalNumbers,
    top: top,
    // สำหรับ "ตรงโต๊ด" ต้องเก็บราคาโต๊ดไว้ เพื่อใช้ตรวจรางวัลภายหลัง
    bottom: mode === "ตท" ? bottom : 0,
    mode: mode,
    betCount: finalNumbers.length,
    preserveDuplicates: true,
    subtotal3: subtotal
  };
}

function perms3(s) {
  const [a, b, c] = s;
  return [...new Set([
    a+b+c, a+c+b,
    b+a+c, b+c+a,
    c+a+b, c+b+a
  ])];
}

function parseRoodBeil(body, top, bottom) {
  const numbers = ["00","11","22","33","44","55","66","77","88","99"];
  return {
    success: true,
    numbers,
    top, bottom,
    mode: "รูดเบิ้ล"
  };
}

function parseBeilNa(body, price) {
  const numbers = [];

  // เบิ้ลหน้า: สองหลักหน้าเหมือนกัน และหลักท้ายต้องต่าง
  // 001..009, 110/112..119, ... , 990..998 = 90 ตัว
  for (let d = 0; d <= 9; d++) {
    for (let last = 0; last <= 9; last++) {
      if (last === d) continue; // ตัด 000,111,...,999
      numbers.push("" + d + d + last);
    }
  }

  return {
    success: true,
    numbers: numbers,
    top: price,
    bottom: 0,
    mode: "เบิ้ลหน้า",
    subtotal3: numbers.length * price
  };
}

function parseBeilLang(body, price) {
  const numbers = [];

  // เบิ้ลหลัง: สองหลักหลังเหมือนกัน และหลักหน้าต้องต่าง
  // 100/200..900, 011/211..911, ... , 099..899 = 90 ตัว
  for (let d = 0; d <= 9; d++) {
    for (let first = 0; first <= 9; first++) {
      if (first === d) continue; // ตัด 000,111,...,999
      numbers.push("" + first + d + d);
    }
  }

  return {
    success: true,
    numbers: numbers,
    top: price,
    bottom: 0,
    mode: "เบิ้ลหลัง",
    subtotal3: numbers.length * price
  };
}

// ================== TEST 3-DIGIT SETS ==================
function testGenerated3DigitSets_() {
  const front = parseBeilNa("เบิ้ลหน้า", 1);
  const back = parseBeilLang("เบิ้ลหลัง", 1);
  const ham = parseHamBlock(1);

  if (front.numbers.length !== 90 || front.subtotal3 !== 90) {
    throw new Error("เบิ้ลหน้าไม่ถูกต้อง: " + front.numbers.length + " ตัว / " + front.subtotal3 + " บาท");
  }
  if (back.numbers.length !== 90 || back.subtotal3 !== 90) {
    throw new Error("เบิ้ลหลังไม่ถูกต้อง: " + back.numbers.length + " ตัว / " + back.subtotal3 + " บาท");
  }
  if (ham.numbers.length !== 90 || ham.subtotal3 !== 90) {
    throw new Error("หามไม่ถูกต้อง: " + ham.numbers.length + " ตัว / " + ham.subtotal3 + " บาท");
  }

  return "PASS: เบิ้ลหน้า 90 / เบิ้ลหลัง 90 / หาม 90";
}

// ================== WIN 2 DIGIT ==================
// ตัวอย่าง 4935:
// 49 43 45
// 94 93 95
// 34 39 35
// 54 59 53
function generateWin2(source) {
  source = String(source || "").replace(/\D/g, "");

  // ตัดเลขต้นทางซ้ำ เช่น 49335 => 4935
  const digits = [...new Set(source.split(""))];
  if (digits.length < 2) return [];

  const result = [];
  for (let i = 0; i < digits.length; i++) {
    for (let j = 0; j < digits.length; j++) {
      if (i === j) continue;
      result.push(digits[i] + digits[j]);
    }
  }

  return [...new Set(result)];
}

function parseDirect2Block(body, top, bottom, blockHeader, priceStr, isHeaderOnly) {
  let t = String(body || "")
    // วินสอง / วิน 2 / วิน2 ให้เป็น canonical "วิน"
    .replace(/วิน\s*(?:สอง|2)/g, "วิน")
    // กันคำสั่งติดเลข/โหมด แม้ splitBlocks จะไม่ได้แยกช่องว่างให้
    .replace(/(\d)(วิน)/g, "$1 $2")
    .replace(/(วิน)(?=บลก|บกล|บล|บนล่างกลับ|บนล่าง|กลับ|บน|ล่าง|บ|ล)/g, "$1 ")
    .replace(/[-_/.=,]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  blockHeader = canonicalMode(blockHeader);

  // รองรับ: วิน / วินสอง / วิน 2 / วิน2
  const hasWin = /(^|\s)วิน(?=\s|$)/.test(t);
  if (hasWin) {
    t = t
      .replace(/(^|\s)วิน(?=\s|$)/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const MODE_KEYWORDS = [
    { text: "บน ล่าง กลับ", mode: "บลก" },
    { text: "บนล่าง กลับ",  mode: "บลก" },
    { text: "บนล่างกลับ",   mode: "บลก" },
    { text: "บน ล่าง",      mode: "บล"  },
    { text: "บนล่าง",       mode: "บล"  },
    { text: "บลก", mode: "บลก" },
    { text: "บกล", mode: "บลก" },
    { text: "กลับ", mode: "ก" },
    { text: "ล่าง", mode: "ล" },
    { text: "บน", mode: "บ" },
    { text: "บล", mode: "บล" },
    { text: "บ", mode: "บ" },
    { text: "ล", mode: "ล" }
  ];

  let mode = null, modeIdx = -1, modeLen = 0;

  for (const mk of MODE_KEYWORDS) {
    const escaped = escapeRegExp(mk.text);
    const re = new RegExp("(^|\\s)" + escaped + "(?=\\s|$)");
    const m = t.match(re);

    if (m) {
      mode = canonicalMode(mk.mode);
      modeIdx = m.index + (m[1] ? m[1].length : 0);
      modeLen = mk.text.length;
      break;
    }
  }

  let numbersText;
  let useTop = Number(top) || 0;
  let useBottom = Number(bottom) || 0;

  if (mode) {
    numbersText = (
      t.substring(0, modeIdx) + " " + t.substring(modeIdx + modeLen)
    ).replace(/\s+/g, " ").trim();

    if (mode === "บ") useBottom = 0;
    else if (mode === "ล") useTop = 0;
  } else {
    numbersText = t;

    // ถ้า blockHeader มาจากคำว่า "วิน" ให้ default เป็น บนล่าง
    const headerMode = canonicalMode(blockHeader);
    mode = /^วิน/.test(String(headerMode || "")) ? "บล" : (headerMode || "บล");

    if (mode === "บ") useBottom = 0;
    else if (mode === "ล") useTop = 0;
  }

  const hasStarPrice = priceStr && String(priceStr).includes("*");
  if (hasStarPrice && isHeaderOnly) {
    if (blockHeader === "ล") {
      if (mode === "ล" || mode === "บล") {
        useTop = 0;
        useBottom = (Number(top) || 0) + (Number(bottom) || 0);
        mode = "ล";
      }
    } else if (blockHeader === "บ") {
      if (mode === "บ" || mode === "บล") {
        useTop = (Number(top) || 0) + (Number(bottom) || 0);
        useBottom = 0;
        mode = "บ";
      }
    }
  }

  if (mode === "บลก" && blockHeader === "ล") {
    useTop = 0;
  } else if (mode === "บลก" && blockHeader === "บ") {
    useBottom = 0;
  }

  // ====================================================
  // วินเลข
  // ====================================================
  if (hasWin) {
    const sources = numbersText
      .split(/\s+/)
      .filter(x => /^\d{2,}$/.test(x));

    if (sources.length === 0) {
      return { success: false, error: "❌ ไม่พบเลขสำหรับวิน" };
    }

    let finalNums = [];
    for (const source of sources) {
      finalNums.push(...generateWin2(source));
    }
    finalNums = [...new Set(finalNums)];

    if (finalNums.length === 0) {
      return { success: false, error: "❌ ไม่สามารถวินเลขได้" };
    }

    // generateWin2() สร้างทั้ง 49 และ 94 มาแล้ว
    // จึงไม่ต้องกลับซ้ำอีกครั้ง
    if (mode === "ก") mode = "บลก";

    const subtotal2 = finalNums.length * (useTop + useBottom);

    return {
      success: true,
      numbers: finalNums,
      top: useTop,
      bottom: useBottom,
      mode: mode,
      isWin: true,
      winSource: sources,
      subtotal2: subtotal2
    };
  }

  // ====================================================
  // เลข 2 ตัวปกติ
  // ====================================================
  const rawNums = numbersText
    .split(/\s+/)
    .filter(x => /^\d+$/.test(x));

  if (rawNums.length === 0) {
    return { success: false, error: "❌ ไม่พบเลข 2 ตัว" };
  }

  const nums2 = [];
  for (let n of rawNums) {
    if (n.length === 1) n = "0" + n;

    if (n.length !== 2) {
      return {
        success: false,
        error: `❌ เลข "${n}" ไม่ใช่เลข 2 ตัว`
      };
    }

    nums2.push(n);
  }

  // PATCH 2026-09-09:
  // เก็บเลขทุก occurrence ตามที่ลูกค้าพิมพ์ ห้าม unique
  const sourceNums2 = nums2.slice();
  let finalNums = [];

  if (mode === "บลก" || mode === "บกล" || mode === "ก") {
    for (const n of sourceNums2) {
      finalNums.push(n);
      const rev = n.charAt(1) + n.charAt(0);
      if (rev !== n) finalNums.push(rev);
    }
    if (mode === "ก") mode = "บลก";
  } else {
    finalNums = sourceNums2.slice();
  }
  return {
    success: true,
    numbers: finalNums,
    top: useTop,
    bottom: useBottom,
    mode: mode,
    isWin: false,

    // PATCH 2026-09-09: จำนวนรายการจริงรวมเลขซ้ำ
    betCount: finalNums.length,
    preserveDuplicates: true,

    subtotal2: finalNums.length * (useTop + useBottom)
  };
}

// ================== TEST: วินสอง ==================
function testWinSongBet_() {
  const keywords = loadKeywords();
  const doorMap = loadDoors();
  const cmd = [
    "45967วินสองบลก1",
    "24/49/29บลก1",
    "รูด2/4บล1"
  ].join("\n");

  const result = parseMessage(cmd, keywords, doorMap);
  if (!result || !result.success) {
    throw new Error("FAIL วินสอง: " + (result && result.error ? result.error : "parse failed"));
  }

  const winGroup = (result.groups || []).find(function(g) { return !!g.isWin; });
  if (!winGroup) throw new Error("FAIL: ไม่พบกลุ่มวินสอง");

  // 45967 มีเลขไม่ซ้ำ 5 ตัว => ordered pairs 5*4 = 20 ตัว
  if ((winGroup.numbers || []).length !== 20) {
    throw new Error("FAIL: วินสองควรได้ 20 ตัว แต่ได้ " + (winGroup.numbers || []).length);
  }

  return "PASS: 45967วินสองบลก1 อ่านได้ " + winGroup.numbers.length + " ตัว";
}

function escapeRegExp(text) {
  return String(text)
    .replace(/[.*+?^${}()|[\]\\\/\-_=:;,!@#%&~]/g, "\\$&");
}

// ================== LOTTO ==================
function loadLotto() {
  const cacheKey = "lotto_data_v3";
  const cached = getCachedJson(cacheKey);
  if (cached) {
    logBotTiming("loadLotto", 0, { cache: "hit" });
    return cached;
  }
  const started = Date.now();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(LOTTO_SHEET);
  const v = sh.getDataRange().getValues();
  const map = {}, aliasMap = {};
  for (let i = 1; i < v.length; i++) {
    const name = String(v[i][0] || "").trim();
    const alias = String(v[i][1] || "").trim();
    const time = normalize(v[i][2]);
    if (!name || !time) continue;
    map[name] = time;
    aliasMap[name] = alias ? alias.split(/\s+/) : [];
  }
  const result = { map, aliasMap };
  putCachedJson(cacheKey, result, DATA_CACHE_TTL_SECONDS);
  logBotTiming("loadLotto", Date.now() - started, { cache: "miss", rows: v.length });
  return result;
}
// =====================================================
// SESSION LOTTO : ระบบหวย เช้า / บ่าย แบบทั่วไป
// ใช้ได้กับหวยทุกตัว ไม่ hard-code ชื่อหวย
// =====================================================

function cleanHiddenText_(text) {
  return String(text || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .trim();
}


// ตรวจว่าหวยรายการนี้เป็น เช้า หรือ บ่าย
function getLottoSession_(name, aliases) {

  const text = cleanHiddenText_(
    [name].concat(aliases || []).join(" ")
  );

  if (/บ่าย/.test(text)) {
    return "บ่าย";
  }

  if (/เช้า/.test(text)) {
    return "เช้า";
  }

  return "";
}


// เอาคำว่า เช้า / บ่าย / รอบเช้า / รอบบ่าย ออกจากชื่อ
//
// นิเคอิเช้า  => นิเคอิ
// นิเคอิบ่าย  => นิเคอิ
// ฮั่งเส็งเช้า VIP => ฮั่งเส็ง VIP
//
function removeSessionWord_(text) {

  return cleanHiddenText_(text)
    .replace(/รอบ\s*เช้า/g, "")
    .replace(/รอบ\s*บ่าย/g, "")
    .replace(/เช้า/g, "")
    .replace(/บ่าย/g, "")
    .replace(/\s+/g, " ")
    .trim();
}


// =====================================================
// ตรวจหวยที่มี เช้า / บ่าย
// ต้องเรียกก่อน findLotto()
// =====================================================
function findSessionLottoChoice(text, lottoData) {

  const map =
    (lottoData && lottoData.map) || {};

  const aliasMap =
    (lottoData && lottoData.aliasMap) || {};


  const rawText =
    cleanHiddenText_(text);

  const normalizedText =
    normalizeLottoText(rawText);


  // ===================================================
  // ถ้าข้อความมีชื่อหวยรอบที่เลือกแบบเต็ม ๆ อยู่แล้ว
  // เช่น "🇨🇳 จีนบ่าย" หรือ alias ที่มีคำว่า บ่าย/เช้า
  // ให้ล็อกตัวนี้ทันที ห้ามเอาไปเทียบชื่อกลางซ้ำ
  // ป้องกัน Quick Reply เลือกแล้วถามวนไปเรื่อย ๆ
  // ===================================================
  const explicitSessionTerms = [];

  Object.keys(map).forEach(function(name) {
    const aliases = aliasMap[name] || [];
    const session = getLottoSession_(name, aliases);
    if (!session) return;

    [name].concat(aliases).forEach(function(term) {
      term = cleanHiddenText_(term);
      if (!term) return;

      // ต้องเป็นชื่อที่ระบุรอบชัดเจนเท่านั้น
      if (!/(เช้า|บ่าย)/.test(term)) return;

      explicitSessionTerms.push({
        term: term,
        name: name,
        closeTime: map[name]
      });
    });
  });

  explicitSessionTerms.sort(function(a, b) {
    return b.term.length - a.term.length;
  });

  for (const item of explicitSessionTerms) {
    const esc = escapeRegExp(item.term);
    const re = new RegExp(
      "(^|[\\s\\n])" + esc + "(?=[\\s\\n]|$)",
      "i"
    );

    if (re.test(rawText)) {
      return {
        name: item.name,
        closeTime: item.closeTime,
        matchedToken: item.term,
        exactSessionChoice: true
      };
    }
  }


  // ===================================================
  // ลูกค้าระบุรอบมาหรือยัง
  // ===================================================

  let requestedSession = "";

  if (/บ่าย/.test(rawText)) {

    requestedSession = "บ่าย";

  } else if (/เช้า/.test(rawText)) {

    requestedSession = "เช้า";
  }


  // ===================================================
  // หา lottery candidates
  // ===================================================

  const matches = [];


  Object.keys(map).forEach(function(name) {

    const aliases =
      aliasMap[name] || [];

    const session =
      getLottoSession_(name, aliases);


    // ไม่มีเช้า/บ่าย ไม่ต้องเข้าระบบนี้
    if (!session) {
      return;
    }


    const searchTerms = [];

    // ---------------------------------------------
    // ชื่อจริง
    // ---------------------------------------------
    searchTerms.push({
      raw: name,
      normalized: normalizeLottoText(name)
    });


    // ชื่อที่ตัด เช้า/บ่าย
    const baseName =
      removeSessionWord_(name);

    if (baseName) {

      searchTerms.push({
        raw: baseName,
        normalized:
          normalizeLottoText(baseName)
      });
    }


    // ---------------------------------------------
    // alias
    // ---------------------------------------------
    aliases.forEach(function(alias) {

      if (!alias) return;


      searchTerms.push({
        raw: alias,
        normalized:
          normalizeLottoText(alias)
      });


      const baseAlias =
        removeSessionWord_(alias);

      if (baseAlias) {

        searchTerms.push({
          raw: baseAlias,
          normalized:
            normalizeLottoText(baseAlias)
        });
      }

    });


    // ยาวก่อน
    searchTerms.sort(function(a, b) {
      return b.normalized.length -
        a.normalized.length;
    });


    let matchedTerm = null;


    for (const term of searchTerms) {

      if (!term.normalized) {
        continue;
      }


      // ป้องกันคำสั้นเกินไป
      if (term.normalized.length < 2) {
        continue;
      }


      if (
        normalizedText.includes(
          term.normalized
        )
      ) {

        matchedTerm = term;
        break;
      }
    }


    if (!matchedTerm) {
      return;
    }


    matches.push({

      name: name,

      closeTime:
        map[name],

      session: session,

      matchedToken:
        matchedTerm.raw
    });

  });


  if (matches.length === 0) {
    return null;
  }


  // ===================================================
  // ถ้าลูกค้าระบุ เช้า / บ่าย มาแล้ว
  // ===================================================

  if (requestedSession) {

    const sameSession =
      matches.filter(function(item) {

        return item.session ===
          requestedSession;

      });


    // เจอรอบเดียว
    if (sameSession.length === 1) {

      const found =
        sameSession[0];


      return {

        name:
          found.name,

        closeTime:
          found.closeTime,

        matchedToken:
          found.matchedToken
      };
    }


    // ถ้ารอบเดียวกันยังมีหลายตัว
    if (sameSession.length > 1) {

      const open =
        sameSession.filter(function(item) {

          return !isClosed(
            item.closeTime
          );

        });

      // เหลือเปิดเพียงตัวเดียว เลือกให้อัตโนมัติ
      if (open.length === 1) {
        return {
          name: open[0].name,
          closeTime: open[0].closeTime,
          matchedToken: open[0].matchedToken,
          autoResolvedSingleOpen: true
        };
      }


      return {

        ambiguous: true,

        candidates:
          open.map(function(item) {

            return {
              name: item.name,
              closeTime:
                item.closeTime
            };

          }),

        matchedTokens:
          [
            ...new Set(
              sameSession
                .map(function(item) {
                  return item.matchedToken;
                })
                .filter(Boolean)
            )
          ],

        sessionChoice: true,

        allClosed:
          open.length === 0
      };
    }


    return null;
  }


  // ===================================================
  // ลูกค้าไม่ได้พิมพ์ เช้า / บ่าย
  //
  // ต้องเช็กว่ามีทั้งสองรอบจริงหรือไม่
  // ===================================================

  const sessions =
    new Set(
      matches.map(function(item) {
        return item.session;
      })
    );


  const hasMorning =
    sessions.has("เช้า");

  const hasAfternoon =
    sessions.has("บ่าย");


  // ถ้าไม่ได้มีทั้ง เช้า + บ่าย
  // ปล่อยให้ findLotto เดิมทำงาน
  if (
    !hasMorning ||
    !hasAfternoon
  ) {

    return null;
  }


  // ===================================================
  // เอาเฉพาะรายการที่ยังเปิด
  // ===================================================

  const openMatches =
    matches.filter(function(item) {

      return !isClosed(
        item.closeTime
      );

    });


  // กันชื่อซ้ำ
  const unique = [];
  const seen = {};


  openMatches.forEach(function(item) {

    if (seen[item.name]) {
      return;
    }

    seen[item.name] = true;

    unique.push(item);
  });


  // ===================================================
  // ถ้าหวยเช้า/บ่ายที่ชื่อย่อนี้อ้างถึง เหลือเปิดจริงเพียงตัวเดียว
  // เลือกให้อัตโนมัติ ไม่ต้องถาม Quick Reply
  // ===================================================
  if (unique.length === 1) {
    return {
      name: unique[0].name,
      closeTime: unique[0].closeTime,
      matchedToken: unique[0].matchedToken,
      autoResolvedSingleOpen: true
    };
  }


  // ===================================================
  // หา token ที่ลูกค้าพิมพ์
  // เอาไว้ลบชื่อกลางตอนกด Quick Reply
  // ===================================================

  const matchedTokens = [
    ...new Set(
      matches
        .map(function(item) {

          return removeSessionWord_(
            item.matchedToken
          );

        })
        .filter(Boolean)
    )
  ];


  return {

    ambiguous: true,

    candidates:
      unique.map(function(item) {

        return {
          name: item.name,
          closeTime:
            item.closeTime
        };

      }),

    matchedTokens:
      matchedTokens,

    sessionChoice: true,

    allClosed:
      unique.length === 0
  };
}

function findLotto(text, data) {
  const { map, aliasMap } = data;
  const names = Object.keys(map).sort((a, b) => b.length - a.length);

  const rawText = String(text || "").trim();

  // =====================================================
  // STEP 0: ชื่อหวยจริงต้องมาก่อนเสมอ
  // =====================================================
  for (const name of names) {
    const esc = escapeRegExp(name);

    const re = new RegExp(
      "(^|[\\s\\n])" + esc + "(?=[\\s\\n]|$)",
      "i"
    );

    if (re.test(rawText)) {
      return {
        name: name,
        closeTime: map[name],
        matchedToken: name
      };
    }
  }


  // =====================================================
  // STEP 1: alias แบบตรงตัว
  // =====================================================
  for (const name of names) {

    const aliases = aliasMap[name] || [];

    // alias ยาวก่อน
    aliases.sort((a, b) => b.length - a.length);

    for (const alias of aliases) {

      if (!alias) continue;

      const esc = escapeRegExp(alias);

      const re = new RegExp(
        "(^|[\\s\\n])" + esc + "(?=[\\s\\n]|$)",
        "i"
      );

      if (re.test(rawText)) {
        return {
          name: name,
          closeTime: map[name],
          matchedToken: alias
        };
      }
    }
  }


  // =====================================================
  // STEP 2: normalized match
  // =====================================================
  const textNorm = normalizeLottoText(rawText);

  for (const name of names) {

    const nameNorm = normalizeLottoText(name);

    if (
      nameNorm &&
      textNorm.includes(nameNorm)
    ) {
      return {
        name: name,
        closeTime: map[name],
        matchedToken: name
      };
    }
  }


  // =====================================================
  // STEP 3: fuzzy เป็นทางเลือกสุดท้าย
  // =====================================================
  const candidates = [];

  for (const name of names) {

    candidates.push({
      keyword: name,
      lottoName: name
    });

    for (const alias of (aliasMap[name] || [])) {

      if (!alias) continue;

      candidates.push({
        keyword: alias,
        lottoName: name
      });
    }
  }


  const cleaned =
    rawText.replace(/[0-9]+/g, " ");

  const tokens =
    cleaned
      .split(/[\s\n,.\-_\/=*]+/)
      .map(t => t.trim())
      .filter(t =>
        t.length >= 2 &&
        /[ก-๙a-zA-Z]/.test(t)
      );


  if (tokens.length === 0) {
    return null;
  }


  const scoreMap = {};
  const matchedTokenMap = {};


  for (const token of tokens) {

    for (const cand of candidates) {

      const score =
        fuzzyScore(
          token,
          cand.keyword
        );


      if (score >= 0.6) {

        if (
          !scoreMap[cand.lottoName] ||
          score > scoreMap[cand.lottoName]
        ) {

          scoreMap[cand.lottoName] = score;
          matchedTokenMap[cand.lottoName] = token;
        }
      }
    }
  }


  const matched =
    Object.keys(scoreMap);


  if (matched.length === 0) {
    return null;
  }


  if (matched.length === 1) {

    const name = matched[0];

    return {
      name: name,
      closeTime: map[name],
      matchedToken: matchedTokenMap[name]
    };
  }


  matched.sort(
    (a, b) =>
      scoreMap[b] - scoreMap[a]
  );


  const top = scoreMap[matched[0]];
  const second = scoreMap[matched[1]];


  if (top - second >= 0.25) {

    const name = matched[0];

    return {
      name: name,
      closeTime: map[name],
      matchedToken: matchedTokenMap[name]
    };
  }


  // เก็บเฉพาะคำที่ทำให้ระบบจับชื่อหวยได้
const matchedTokens = [
  ...new Set(
    matched
      .map(n => matchedTokenMap[n])
      .filter(Boolean)
  )
];

return {

  ambiguous: true,

  candidates: matched.map(n => ({
    name: n,
    closeTime: map[n]
  })),

  matchedTokens: matchedTokens
};
}

function normalizeLottoText(s) {
  if (!s) return "";
  let t = String(s).toLowerCase().replace(/\s+/g, "");
  t = t.replace(/v(?!ip)/g, "vip");
  return t;
}

function fuzzyScore(token, keyword) {
  if (!token || !keyword) return 0;
  if (token === keyword) return 1.0;
  if (keyword.includes(token)) {
    return Math.max(0.85, token.length / keyword.length);
  }
  if (token.includes(keyword)) return 0.9;
  if (isSubsequence(token, keyword)) {
    const ratio = token.length / keyword.length;
    return Math.max(0.65, ratio + 0.2);
  }
  const dist = levenshtein(token, keyword);
  const maxLen = Math.max(token.length, keyword.length);
  const sim = 1 - dist / maxLen;
  if (sim >= 0.7) return sim;
  return 0;
}

function isSubsequence(a, b) {
  let i = 0, j = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) i++;
    j++;
  }
  return i === a.length;
}

function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const dp = [];
  for (let i = 0; i <= a.length; i++) dp[i] = [i];
  for (let j = 0; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[a.length][b.length];
}

function replyQuickReplyLottoChoice(
  token,
  candidates,
  originalText,
  matchedTokens
) {
  // Safety gate: ถ้าไม่มีรายการแทงจริง ไม่ต้องให้เลือกหวย
  if (!shouldOfferLottoChoiceForBet_(originalText)) {
    return;
  }

  const started = Date.now();

  // =====================================================
  // เอาเฉพาะหวยที่ "ยังไม่ปิดรับ"
  // =====================================================
  const openCandidates = (candidates || [])
    .filter(function(c) {
      return c &&
        c.name &&
        c.closeTime &&
        !isClosed(c.closeTime);
    })
  .sort(function(a, b) {
  return lottoDisplaySortMin_(a.closeTime) -
         lottoDisplaySortMin_(b.closeTime);
});


  // =====================================================
  // ถ้าไม่มีหวยที่ยังเปิด
  // =====================================================
  if (openCandidates.length === 0) {
    replyText(token, "❌ หวยรายการนี้ปิดรับแล้วทั้งหมด");
    return;
  }


  // =====================================================
  // เก็บโพยเดิม
  // =====================================================
  let cleanText = String(originalText || "");


  // =====================================================
  // ลบคำค้นเดิมก่อนสร้างข้อความของ Quick Reply
  // สำคัญ: ถ้าเดิมพิมพ์ "จีน" แล้วเลือก "จีนบ่าย"
  // ข้อความใหม่ต้องเหลือแค่ "จีนบ่าย\nโพย..."
  // ห้ามกลายเป็น "จีนบ่าย\nจีน\nโพย..." เพราะจะถามเลือกซ้ำ
  // =====================================================
  const removeTokens = [];

  (matchedTokens || []).forEach(function(tk) {
    if (tk) removeTokens.push(String(tk));
  });

  // เพิ่มชื่อกลางที่ตัดคำว่า เช้า/บ่าย ออกจาก candidate
  // เช่น 🇨🇳 จีนบ่าย -> 🇨🇳 จีน -> จีน
  (candidates || []).forEach(function(c) {
    if (!c || !c.name) return;

    const base = removeSessionWord_(c.name);
    if (base) {
      removeTokens.push(base);

      const plainBase = String(base)
        .replace(/^[^ก-๙a-zA-Z0-9]+/, "")
        .replace(/[^ก-๙a-zA-Z0-9]+$/, "")
        .trim();

      if (plainBase) removeTokens.push(plainBase);
    }
  });

  [...new Set(removeTokens)]
    .filter(Boolean)
    .sort(function(a, b) {
      return b.length - a.length;
    })
    .forEach(function(tk) {
      const esc = escapeRegExp(tk);
      const re = new RegExp(
        "(^|[^ก-๙a-zA-Z0-9])" + esc + "(?![ก-๙a-zA-Z0-9])",
        "gi"
      );
      cleanText = cleanText.replace(re, "$1");
    });


  // =====================================================
  // จัดข้อความ
  // =====================================================
  cleanText = cleanText
    .split(/\n/)
    .map(function(line) {
      return line
        .replace(/[ \t]+/g, " ")
        .trim();
    })
    .filter(Boolean)
    .join("\n");


  // =====================================================
  // QUICK REPLY
  // LINE รองรับสูงสุด 13 ปุ่ม
  // =====================================================
  const items = openCandidates
    .slice(0, 13)
    .map(function(c) {

      return {
        type: "action",

        action: {
          type: "message",

          label: c.name.substring(0, 20),

          text: (
            c.name +
            "\n" +
            cleanText
          ).trim()
        }
      };
    });


  const msg = {
    type: "text",

    text:
      "หวยที่ยังเปิดรับ กดที่ชื่อหวยที่ต้องการจะแทง\n*** กดที่ชื่อเท่านั้นนะครับ  :",

    quickReply: {
      items: items
    }
  };


  const response = UrlFetchApp.fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "post",

      headers: {
        Authorization:
          "Bearer " + getChannelAccessToken_()
      },

      contentType: "application/json",

      payload: JSON.stringify({
        replyToken: token,
        messages: [msg]
      })
    }
  );


  logBotTiming(
    "replyQuickReplyLottoChoice",
    Date.now() - started,
    {
      status: response.getResponseCode(),
      openCandidates: openCandidates.length
    }
  );
}

function replyQuickReplyOpenLottoChoice(token, lottoData, originalText) {
  // Safety gate: ห้ามขึ้น Quick Reply จากข้อความสนทนาทั่วไป
  if (!shouldOfferLottoChoiceForBet_(originalText)) {
    return;
  }

  const candidates = getOpenLottoCandidates(lottoData);
  if (candidates.length === 0) {
    replyText(token, "❌ ตอนนี้หวยปิดรับแล้วทั้งหมด");
    return;
  }

 const cleanText = String(originalText || "").trim();
  const items = candidates.slice(0, 13).map(c => ({
    type: "action",
    action: {
      type: "message",
      label: c.name.substring(0, 20),
text: (
  c.name +
  (cleanText ? "\n" + cleanText : "")
).trim()
    }
  }));

  const msg = {
    type: "text",
    text: "กดที่ชื่อหวยที่ต้องการจะแทง\n*** กดที่ชื่อเท่านั้นนะครับ  :",
    quickReply: { items }
  };

  const started = Date.now();
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    headers: { Authorization: "Bearer " + getChannelAccessToken_() },
    contentType: "application/json",
    payload: JSON.stringify({ replyToken: token, messages: [msg] })
  });
  logBotTiming("replyQuickReplyOpenLottoChoice", Date.now() - started, {
    status: response.getResponseCode()
  });
}

function replyQuickReplyReturnChoice(token, originalText) {
  const started = Date.now();
  const cleanText = String(originalText || "").trim();
  const msg = {
    type: "text",
    text: "ต้องการกลับเลขไหม?",
    quickReply: {
      items: [
        {
          type: "action",
          action: {
            type: "message",
            label: "กลับเลข",
            text: `${cleanText}\nกลับ`
          }
        },
        {
          type: "action",
          action: {
            type: "message",
            label: "ไม่กลับ",
            text: `${cleanText}\nบนล่าง`
          }
        }
      ]
    }
  };

  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    headers: { Authorization: "Bearer " + getChannelAccessToken_() },
    contentType: "application/json",
    payload: JSON.stringify({ replyToken: token, messages: [msg] })
  });
  logBotTiming("replyQuickReplyReturnChoice", Date.now() - started, {
    status: response.getResponseCode()
  });
}

// ================== KEYWORDS ==================
function loadKeywords() {
  const cacheKey = "keywords_v3";
  const cached = getCachedJson(cacheKey);
  if (cached) {
    logBotTiming("loadKeywords", 0, { cache: "hit" });
    return cached;
  }
  const started = Date.now();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(KEY_SHEET);
  const values = sh.getDataRange().getValues();
  const result = values.flat()
    .map(x => String(x).trim()).filter(x => x);
  putCachedJson(cacheKey, result, DATA_CACHE_TTL_SECONDS);
  logBotTiming("loadKeywords", Date.now() - started, { cache: "miss", rows: values.length });
  return result;
}

function findKeyword(text, keys) {
  const sorted = [...keys].sort((a,b)=>b.length-a.length);
  for (const k of sorted) if (text.includes(k)) return k;
  return null;
}

// ================== DOORS ==================
function loadDoors() {
  const cacheKey = "doors_v3";
  const cached = getCachedJson(cacheKey);
  if (cached) {
    logBotTiming("loadDoors", 0, { cache: "hit" });
    return cached;
  }
  const started = Date.now();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(DOOR_SHEET);
  const v = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < v.length; i++) {
    const d = String(v[i][0]).trim();
    const arr = String(v[i][1]).trim().split(/\s+/);
    map[d] = arr;
  }
  putCachedJson(cacheKey, map, DATA_CACHE_TTL_SECONDS);
  logBotTiming("loadDoors", Date.now() - started, { cache: "miss", rows: v.length });
  return map;
}

function getCachedJson(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    logBotTiming("cacheReadError", 0, { key, message: String(err && err.message ? err.message : err) });
    return null;
  }
}

function putCachedJson(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSeconds);
  } catch (err) {
    logBotTiming("cacheWriteError", 0, { key, message: String(err && err.message ? err.message : err) });
  }
}

function logBotTiming(name, ms, meta) {
  try {
    const payload = Object.assign({
      tag: "bot_timing",
      name,
      ms
    }, meta || {});
    console.log(JSON.stringify(payload));
  } catch (err) {
    // Logging must never block bot replies.
  }
}

// ================== TIME ==================
function normalize(t) {
  t = String(t).replace(/ปิด|น\.?/g,"").replace(/\s/g,"").replace(".",":");
  if (/^\d{1,2}:\d{1,2}$/.test(t)) {
    let [h,m]=t.split(":");
    return h.padStart(2,"0")+":"+m.padStart(2,"0");
  }
  return null;
}

function isClosed(c) {
  const closeMin = toMin(c);
  let nowMin = toMin(Utilities.formatDate(new Date(), TZ, "HH:mm"));

  // หวยที่ปิดหลังเที่ยงคืน (00:00-02:00) ให้นับเป็นรอบของวันก่อนหน้า
  if (closeMin <= 120 && nowMin >= 720) {
    nowMin -= 24 * 60;
  }

  return nowMin > closeMin;
}

function toMin(t){
  const [h,m]=t.split(":");
  return +h*60 + +m;
}

// ทดสอบกติกาเว็บ: หวยเช้าที่ปิดไปแล้วต้องเลื่อนไปรอบวันพรุ่งนี้
function testLiffMorningRollToTomorrow_() {
  const fakeNow = new Date("2026-08-19T16:59:00+07:00");
  const info = getLiffLottoRoundInfo_("🇯🇵 นิเคอิเช้า", "09:10", fakeNow);

  if (!info.open) throw new Error("FAIL: หวยเช้าต้องเปิดรับรอบพรุ่งนี้");
  if (info.roundDate !== "2026-08-20") {
    throw new Error("FAIL: roundDate ต้องเป็น 2026-08-20 แต่ได้ " + info.roundDate);
  }

  const evening = getLiffLottoRoundInfo_("หุ้นไทยเย็น", "16:30", fakeNow);
  if (evening.open) throw new Error("FAIL: หวยเย็นที่ปิดแล้วต้องไม่ถูกเลื่อนไปรอบพรุ่งนี้");

  Logger.log("PASS testLiffMorningRollToTomorrow_: " + JSON.stringify(info));
  return info;
}

function testThreeDigitStraightTodPrice_() {
  const parsed = parseDirect3Block("687 817 ตรงโต๊ด", 1, 0, "1", "ตท");
  if (!parsed.success) throw new Error("parse 3 ตัวไม่สำเร็จ");
  if (Number(parsed.top) !== 1 || Number(parsed.bottom) !== 1) {
    throw new Error("ราคา ตรง+โต๊ด ผิด: " + parsed.top + "/" + parsed.bottom);
  }
  if (Number(parsed.subtotal3) !== 4) {
    throw new Error("ยอด 2 เลข x (ตรง1+โต๊ด1) ต้องเป็น 4 แต่ได้ " + parsed.subtotal3);
  }

  const label = getFlexPriceLabel({
    mode: "ตท", top: 1, bottom: 0, betCount: 31, subtotal: 62, numbers: []
  });
  if (label !== "ตรง 1 โต๊ด 1") {
    throw new Error("Flex label ผิด: " + label);
  }

  Logger.log("PASS testThreeDigitStraightTodPrice_");
}

// ================== FLEX ==================
function buildFlex(d) {
  const groupBlocks = [];
  const displayGroups = mergeFlexGroups(d.groups || []);

  // วันที่ของรอบหวย (หวยปิดหลังเที่ยงคืนจะยังนับเป็นรอบวันก่อน)
  const flexRoundDate = /^\d{4}-\d{2}-\d{2}$/.test(String(d.roundDate || ""))
    ? String(d.roundDate)
    : rsGetRoundDate_(new Date(), d.closeTime);
  const flexDate = rsDisplayThaiDate_(flexRoundDate);

  displayGroups.forEach((g, gi) => {
    const displayMode = getFlexModeLabel(g);
    const numbersText = buildFlexNumbersText(g);

    groupBlocks.push({
      type: "box",
      layout: "horizontal",
      margin: gi === 0 ? "none" : "lg",
      contents: [
        {
          type: "text",
          text: displayMode,
          color: "#DC2626",
          size: "sm",
          weight: "bold",
          flex: 0
        },
        {
          type: "text",
          text: getFlexPriceLabel(g),
          color: "#9CA3AF",
          size: "xs",
          align: "end"
        }
      ]
    });

    groupBlocks.push({
      type: "box",
      layout: "horizontal",
      margin: "xs",
      contents: [
        {
          type: "text",
          text: `${(typeof g.betCount === "number" ? g.betCount : g.numbers.length)} ตัว`,
          color: "#9CA3AF",
          size: "xs",
          flex: 0
        },
        {
          type: "text",
          text: `= ${numberWithCommas(g.subtotal)} บาท`,
          color: "#6B7280",
          size: "xs",
          align: "end",
          weight: "bold"
        }
      ]
    });

    groupBlocks.push({
      type: "box",
      layout: "vertical",
      margin: "sm",
      paddingAll: "10px",
      backgroundColor: "#FEF2F2",
      cornerRadius: "8px",
      contents: [
        {
          type: "text",
          text: numbersText,
          color: "#7F1D1D",
          weight: "bold",
          size: "sm",
          wrap: true
        }
      ]
    });
  });

  const header = {
    type: "box",
    layout: "vertical",
    paddingAll: "20px",
    backgroundColor: "#B91C1C",
    contents: [
      {
        type: "text",
        text: d.name,
        color: "#FFFFFF",
        weight: "bold",
        size: "xl"
      },
      {
        type: "text",
        text: "📅 วันที่ " + flexDate,
        color: "#FEE2E2",
        size: "sm",
        margin: "sm"
      },
      {
        type: "box",
        layout: "horizontal",
        margin: "xs",
        contents: [
          {
            type: "text",
            text: "🕐 ปิด " + d.closeTime,
            color: "#FEE2E2",
            size: "sm",
            flex: 0
          },
          {
            type: "text",
            text: `• ${d.totalDoors} ตัว`,
            color: "#FEE2E2",
            size: "sm",
            margin: "sm",
            flex: 0
          }
        ]
      }
    ]
  };

  const gradientStrip = {
    type: "box",
    layout: "vertical",
    contents: [
      {
        type: "box",
        layout: "vertical",
        contents: [{ type: "filler" }],
        height: "3px",
        backgroundColor: "#DC2626"
      },
      {
        type: "box",
        layout: "vertical",
        contents: [{ type: "filler" }],
        height: "3px",
        backgroundColor: "#EF4444"
      },
      {
        type: "box",
        layout: "vertical",
        contents: [{ type: "filler" }],
        height: "3px",
        backgroundColor: "#FCA5A5"
      }
    ]
  };

  const totalBox = {
    type: "box",
    layout: "vertical",
    paddingAll: "16px",
    backgroundColor: "#FEF2F2",
    cornerRadius: "10px",
    contents: [
      {
        type: "text",
        text: "ยอดที่ต้องโอน",
        color: "#DC2626",
        size: "sm",
        align: "center",
        weight: "bold"
      },
      {
        type: "box",
        layout: "baseline",
        margin: "xs",
        justifyContent: "center",
        contents: [
          {
            type: "text",
            text: numberWithCommas(d.totalAmount),
            color: "#B91C1C",
            size: "3xl",
            weight: "bold",
            flex: 0
          },
          {
            type: "text",
            text: " บาท",
            color: "#DC2626",
            size: "md",
            weight: "bold",
            flex: 0
          }
        ]
      }
    ]
  };

  const verifyBeforePay = {
    type: "text",
    text: "โปรดตรวจสอบรายการก่อนโอนเงิน",
    color: "#9CA3AF",
    size: "sm",
    align: "center",
    margin: "md",
    wrap: true
  };

  const autoBuyButton = {
    type: "button",
    style: "primary",
    color: "#B91C1C",
    height: "sm",
    margin: "lg",
    action: {
      type: "uri",
      label: "ซื้อหวยออโต้",
      uri: "https://supmangmee.com"
    }
  };

  const body = {
    type: "box",
    layout: "vertical",
    paddingAll: "20px",
    backgroundColor: "#FFFFFF",
    contents: [
      ...groupBlocks,
      {
        type: "separator",
        margin: "lg",
        color: "#FCA5A5"
      },
      totalBox,
      verifyBeforePay,
      autoBuyButton
    ]
  };

  return {
    type: "flex",
    altText: `${d.name} รวม ${d.totalAmount} บาท`,
    contents: {
      type: "bubble",
      size: "kilo",
      header: header,
      hero: gradientStrip,
      body: body
    }
  };
}

function numberWithCommas(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function getFlexPriceLabel(g) {
  const mode = String((g && g.mode) || "");
  const top = Number((g && g.top) || 0);
  let bottom = Number((g && g.bottom) || 0);

  // กรณีข้อมูลเก่าหรือกลุ่มที่ subtotal ถูกต้อง แต่ bottom หลุดเป็น 0
  // ให้คำนวณราคาโต๊ดคืนจากยอดจริง เพื่อให้ Flex แสดงตรงกับยอดโอน
  if (mode === "ตท" && bottom <= 0 && top > 0) {
    const count = Number((g && g.betCount) || ((g && g.numbers) || []).length || 0);
    const subtotal = Number((g && g.subtotal) || 0);
    if (count > 0 && subtotal > 0) {
      const inferred = (subtotal / count) - top;
      if (inferred > 0) bottom = inferred;
    }
    if (bottom <= 0) bottom = top;
  }

  // 3 ตัวต้องใช้คำว่า ตรง / โต๊ด ไม่ใช้ บน / ล่าง
  if (mode === "ตท") return `ตรง ${top} โต๊ด ${bottom}`;
  if (mode === "ตรง") return `ตรง ${top}`;
  if (mode === "โต๊ด") return `โต๊ด ${top}`;
  if (mode === "6กลับ" || mode === "3กลับ" || mode === "ตอง" ||
      mode === "หาม" || mode === "เบิ้ลหน้า" || mode === "เบิ้ลหลัง") {
    return `ราคา ${top}`;
  }

  // เลข 2 ตัวคงรูปแบบเดิม
  return `บน ${top} ล่าง ${bottom}`;
}

function getFlexModeLabel(g) {
  const mode = String((g && g.mode) || "");
  let label = mode;

  if (mode === "บล") label = "บนล่าง";
  else if (mode === "บลก") label = "บนล่างกลับ";
  else if (mode === "บ") label = "บน";
  else if (mode === "ล") label = "ล่าง";
  else if (mode === "ตรง") label = "3 ตรง";
  else if (mode === "ตท") label = "3 ตรง+โต๊ด";
  else if (mode === "โต๊ด") label = "3 โต๊ด";
  else if (mode === "ตอง") label = "เลขตอง";
  else if (mode === "เบิ้ลหน้า") label = "เบิ้ลหน้า";
  else if (mode === "เบิ้ลหลัง") label = "เบิ้ลหลัง";
  else if (mode === "หาม") label = "หาม";

  if (g && g.winCombinedDouble) return "วินรวมเบิ้ล";
  return (g && g.isWin) ? "วิน " + label : label;
}
function buildFlexNumbersText(g) {
  const nums = ((g && g.numbers) || []).map(n => String(n));

  // วินชุดเดียว: จัดบรรทัดตามเลขต้นทาง
  // 4935 => บรรทัดละ 3 ตัว
  if (g && g.isWin && Array.isArray(g.winSource) && g.winSource.length === 1) {
    const sourceDigits = [...new Set(String(g.winSource[0] || "").replace(/\D/g, "").split(""))];
    const perRow = Math.max(1, sourceDigits.length - 1);
    const rows = [];

    for (let i = 0; i < nums.length; i += perRow) {
      rows.push(nums.slice(i, i + perRow).join(" "));
    }
    return rows.join("\n");
  }

  return nums.join(" ");
}
function mergeFlexGroups(groups) {
  const merged = [];
  const indexByKey = {};

  (groups || []).forEach(function(g) {
    const key = [
      g.mode,
      g.top,
      g.bottom,
      g.isWin ? "win" : "normal",
      g.preserveDuplicates ? "keepdup" : "dedupe"
    ].join("|");

    const sourceCount = (typeof g.betCount === "number")
      ? g.betCount
      : (g.numbers || []).length;

    if (indexByKey[key] === undefined) {
      indexByKey[key] = merged.length;

      merged.push({
        mode: g.mode,
        top: g.top,
        bottom: g.bottom,
        numbers: [...(g.numbers || [])],
        isWin: !!g.isWin,
        winSource: [...(g.winSource || [])],
        winCombinedDouble: !!g.winCombinedDouble,
        preserveDuplicates: !!g.preserveDuplicates,
        betCount: sourceCount,
        subtotal: typeof g.subtotal === "number"
          ? g.subtotal
          : sourceCount * (g.top + g.bottom)
      });
      return;
    }

    const target = merged[indexByKey[key]];
    target.numbers = target.numbers.concat(g.numbers || []);

    // รูดหลายหลักต้องเก็บเลขซ้ำข้ามชุดไว้ เช่น 69/96 ของรูด 6 และรูด 9
    if (!target.preserveDuplicates) {
      target.numbers = [...new Set(target.numbers)];
    }

    if (g.isWin && Array.isArray(g.winSource)) {
      target.winSource = [...new Set(target.winSource.concat(g.winSource))];
    }

    target.betCount += sourceCount;

    target.subtotal += typeof g.subtotal === "number"
      ? g.subtotal
      : sourceCount * (g.top + g.bottom);
  });

  return merged;
}
function normalizePriceSeparators(text) {
  let t = text;
  t = t.replace(/[×✕✖⨯]/g, "*");
  t = t.replace(/(\d)\s*\+\s*(\d)/g, "$1*$2");
  return t;
}


// ============================================================================
// IMAGE BET OCR 2026-09-21
// อ่านโพยจากรูปด้วย Google Cloud Vision OCR
//
// Flow:
// 1) โหลดรูปจาก LINE
// 2) OCR ด้วย Google Cloud Vision
// 3) ถ้าเจอชุดเลขแบบโพย -> สร้างโพย 2 ตัว/3 ตัว ราคา 10*10
// 4) ขึ้น Quick Reply ให้ลูกค้าเลือกชื่อหวยที่ยังเปิด
// 5) ถ้าไม่เหมือนโพย -> fallback ไปตรวจสลิป Thunder ตามระบบเดิม
// ============================================================================
const IMAGE_BET_TWO_TOP_ = 10;
const IMAGE_BET_TWO_BOTTOM_ = 10;
const IMAGE_BET_THREE_DIRECT_ = 10;
const IMAGE_BET_THREE_TOD_ = 10;


// ============================================================================
// BANGKOK BANK SLIP OCR 2026-09-22
//
// สลิป "ต้นทางธนาคารกรุงเทพ" ไม่เรียก Thunder
// ใช้ข้อความ OCR จากรูปโดยตรง:
//   - ชื่อผู้โอน ต้องอ่านได้
//   - ชื่อผู้รับ ต้องอ่านได้และตรงกับบัญชีร้าน
//   - ยอดเงิน ต้องอ่านได้และมากกว่า 0
//
// ธนาคารอื่นยังใช้ Thunder ตามเดิม
// ============================================================================
function isBangkokBankOriginSlipOcr_(rawText) {
  const lines = String(rawText || "")
    .split(/\r?\n/)
    .map(function(line) {
      return String(line || "").trim();
    })
    .filter(Boolean);

  if (lines.length === 0) return false;

  // ใช้เฉพาะส่วนต้นของสลิปเป็นหลัก
  // ป้องกันกรณี "ธนาคารกรุงเทพ" ปรากฏแค่ตรงบัญชีผู้รับ
  const head = lines
    .slice(0, 10)
    .join(" ");

  if (
    /bangkok\s*bank|bualuang|บัวหลวง/i.test(head)
  ) {
    return true;
  }

  if (
    /ธนาคารกรุงเทพ/i.test(head) &&
    /โอนเงินสำเร็จ|โอนสำเร็จ|ทำรายการสำเร็จ|transfer/i.test(
      String(rawText || "")
    )
  ) {
    return true;
  }

  return false;
}


function cleanBangkokSlipName_(value) {
  let s = String(value || "")
    .replace(
      /^(?:จาก|ผู้โอน|ชื่อผู้โอน|ไปที่|ไปยัง|ถึง|ผู้รับ|ชื่อผู้รับ|from|to)\s*[:：\-–—]?\s*/i,
      ""
    )
    .replace(/\s+/g, " ")
    .trim();

  // Bangkok Bank OCR บางครั้งรวม:
  // "น.ส. ทัศนีย์ 624-0-xxx863 ธนาคารกรุงเทพ"
  // ให้นำเฉพาะชื่อบุคคลช่วงต้นออกมา
  const honorificName = s.match(
    /^(?:(?:นาย|นางสาว|นาง|น\.?\s*ส\.?|ด\.?\s*ช\.?|ด\.?\s*ญ\.?)\s*)[ก-๙A-Za-z]+(?:\s+[ก-๙A-Za-z]+){0,4}/i
  );

  if (honorificName) {
    return String(honorificName[0] || "")
      .replace(/\s+/g, " ")
      .trim();
  }

  // ถ้าไม่มีคำนำหน้า ให้ตัดก่อนเลขบัญชี/ชื่อธนาคาร
  s = s
    .replace(
      /\s+(?:\d{3}[-\s]?\d[-\sxX*]*\d{2,}|ธนาคาร|bank|SCB|KBANK|KTB|BBL|BAY|ttb).*$/i,
      ""
    )
    .trim();

  return s;
}

function looksLikeBangkokSlipPersonName_(value) {
  const s =
    cleanBangkokSlipName_(value);

  if (!s || s.length < 3) return false;

  // ต้องมีตัวอักษร ไม่ใช่เลขล้วน
  if (!/[ก-๙A-Za-z]/.test(s)) return false;

  // ตัด label / ธนาคาร / เลขบัญชี / จำนวนเงิน / metadata
  if (
    /ธนาคาร|bank|บัญชี|account|พร้อมเพย์|promptpay|จำนวนเงิน|ยอดเงิน|บาท|ค่าธรรมเนียม|เลขที่รายการ|เลขอ้างอิง|reference|transaction|วันที่|เวลา|โอนเงินสำเร็จ|ทำรายการสำเร็จ|QR/i.test(
      s
    )
  ) {
    return false;
  }

  // ไม่รับชื่อธนาคาร/คำภาษาอังกฤษที่ OCR แยกเหลือคำเดียว
  if (
    /^(?:bangkok|scb|kbank|ktb|bbl|bay|ttb|thb|bank)$/i.test(s)
  ) {
    return false;
  }

  // ไม่รับบรรทัดที่เป็นตัวเลขยาว
  if (/\d{6,}/.test(s)) return false;

  return true;
}


function findBangkokSlipPartyName_(
  lines,
  labelRegex
) {
  for (let i = 0; i < lines.length; i++) {
    const line =
      String(lines[i] || "")
        .replace(/\s+/g, " ")
        .trim();

    if (!labelRegex.test(line)) {
      continue;
    }

    // 1) ชื่ออยู่บรรทัดเดียวกับ label
    const inline =
      cleanBangkokSlipName_(
        line.replace(
          labelRegex,
          ""
        )
      );

    if (
      looksLikeBangkokSlipPersonName_(
        inline
      )
    ) {
      return inline;
    }

    // 2) ชื่ออยู่บรรทัดถัดไป
    // สลิป Bangkok Bank ยาว บางครั้ง Vision แทรกบรรทัดโลโก้/บัญชี
    // จึงไล่หาได้ถึง 7 บรรทัด
    for (
      let j = i + 1;
      j < Math.min(lines.length, i + 8);
      j++
    ) {
      const candidate =
        cleanBangkokSlipName_(
          lines[j]
        );

      if (
        looksLikeBangkokSlipPersonName_(
          candidate
        )
      ) {
        return candidate;
      }
    }
  }

  return "";
}

function extractBangkokSlipAmount_(rawText) {
  const text =
    String(rawText || "")
      .replace(/[，]/g, ",")
      .replace(/[฿]/g, "฿")
      .replace(/\s+/g, " ");

  const patterns = [
    /(?:จำนวนเงิน|ยอดเงิน|ยอดโอน|amount)\s*[:：\-–—]?\s*(?:THB|฿)?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i,
    /(?:THB|฿)\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/i,
    /([0-9][0-9,]*(?:\.[0-9]{2}))\s*บาท/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const m =
      text.match(patterns[i]);

    if (!m) continue;

    const amount =
      Number(
        String(m[1] || "")
          .replace(/,/g, "")
      );

    if (
      Number.isFinite(amount) &&
      amount > 0
    ) {
      return amount;
    }
  }

  return null;
}


function extractBangkokSlipDateText_(rawText) {
  const text =
    String(rawText || "");

  const m =
    text.match(
      /(\d{1,2}\s*(?:ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.|\-|\/)\s*\d{1,2}(?:\s*(?:ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.|\-|\/)\s*\d{2,4})?(?:\s*[-–—]?\s*\d{1,2}:\d{2})?)/i
    );

  return m
    ? String(m[1] || "").trim()
    : "";
}


function extractBangkokSlipReference_(rawText) {
  const lines =
    String(rawText || "")
      .split(/\r?\n/)
      .map(function(line) {
        return String(line || "").trim();
      })
      .filter(Boolean);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const m =
      line.match(
        /(?:เลขที่รายการ|เลขอ้างอิง|รหัสอ้างอิง|หมายเลขรายการ|reference|transaction)\s*[:：\-–—]?\s*([A-Za-z0-9\-]{6,})/i
      );

    if (m) {
      return String(m[1] || "");
    }
  }

  return "";
}


function getBangkokSlipReceiverBankText_(rawText) {
  const text =
    String(rawText || "");

  if (
    /ไทยพาณิชย์|SCB/i.test(text)
  ) {
    return "SCB";
  }

  if (
    /กสิกร|KBANK|KBank|K PLUS/i.test(text)
  ) {
    return "KBANK";
  }

  if (
    /กรุงไทย|KTB/i.test(text)
  ) {
    return "KTB";
  }

  if (
    /กรุงศรี|BAY/i.test(text)
  ) {
    return "BAY";
  }

  if (
    /ธนาคารกรุงเทพ|Bangkok Bank|BBL/i.test(
      text
    )
  ) {
    return "BBL";
  }

  return "-";
}



// ---------------------------------------------------------------------------
// Bangkok Bank OCR robust name helpers
// ---------------------------------------------------------------------------
function hasExpectedBangkokSlipReceiverName_(rawText) {
  const t =
    normalizeSlipText(
      String(rawText || "")
    );

  // ชื่อผู้รับของร้านอาจ OCR เป็น
  // "นางสาว ชื่อบัญชี นามสกุลบัญชี" หรือ "น.ส. ชื่อบัญชี นามสกุลบัญชี"
  return (
    t.indexOf("ชื่อบัญชี") !== -1 &&
    t.indexOf("นามสกุลบัญชี") !== -1
  );
}


function extractThaiPersonNamesFromBangkokSlip_(rawText) {
  const text =
    String(rawText || "")
      .replace(/\r\n?/g, "\n");

  const found = [];

  const re =
    /(?:นาย|นางสาว|นาง|น\.?\s*ส\.?|ด\.?\s*ช\.?|ด\.?\s*ญ\.?)\s*[ก-๙A-Za-z]+(?:\s+[ก-๙A-Za-z]+){0,3}/gi;

  let m;

  while ((m = re.exec(text)) !== null) {
    const name =
      cleanBangkokSlipName_(
        String(m[0] || "")
      );

    if (
      looksLikeBangkokSlipPersonName_(name) &&
      found.indexOf(name) === -1
    ) {
      found.push(name);
    }
  }

  return found;
}


function pickBangkokSlipSenderName_(
  rawText,
  labelSenderName,
  receiverName
) {
  // ถ้าชื่อจาก label ดูเป็นชื่อคนไทยจริง ใช้ก่อน
  if (
    labelSenderName &&
    looksLikeBangkokSlipPersonName_(
      labelSenderName
    ) &&
    !isExpectedAccountName(
      labelSenderName
    )
  ) {
    return labelSenderName;
  }

  const names =
    extractThaiPersonNamesFromBangkokSlip_(
      rawText
    );

  for (let i = 0; i < names.length; i++) {
    const name = names[i];

    if (
      receiverName &&
      normalizeSlipText(name) ===
        normalizeSlipText(receiverName)
    ) {
      continue;
    }

    if (
      isExpectedAccountName(name)
    ) {
      continue;
    }

    return name;
  }

  return "";
}


function parseBangkokBankSlipFromOcr_(
  rawText
) {
  const lines =
    String(rawText || "")
      .split(/\r?\n/)
      .map(function(line) {
        return String(line || "")
          .replace(/\s+/g, " ")
          .trim();
      })
      .filter(Boolean);

  const labelSenderName =
    findBangkokSlipPartyName_(
      lines,
      /^(?:จาก|ผู้โอน|ชื่อผู้โอน|from)(?:\s|[:：\-–—]|$)*/i
    );

  const labelReceiverName =
    findBangkokSlipPartyName_(
      lines,
      /^(?:ไปที่|ไปยัง|ถึง|ผู้รับ|ชื่อผู้รับ|to)(?:\s|[:：\-–—]|$)*/i
    );

  // ผู้รับ:
  // ถ้า OCR ทั้งภาพมีชื่อ "ชื่อบัญชี นามสกุลบัญชี" ให้ใช้ชื่อบัญชีร้านโดยตรง
  // ป้องกัน Vision เรียงบรรทัดผิดแล้วจับ "Bangkok" เป็นชื่อผู้รับ
  const receiverName =
    hasExpectedBangkokSlipReceiverName_(rawText)
      ? BANK_ACCOUNT_NAME
      : (
          isExpectedAccountName(labelReceiverName)
            ? labelReceiverName
            : labelReceiverName
        );

  const senderName =
    pickBangkokSlipSenderName_(
      rawText,
      labelSenderName,
      receiverName
    );

  const amount =
    extractBangkokSlipAmount_(
      rawText
    );

  const dateText =
    extractBangkokSlipDateText_(
      rawText
    );

  const reference =
    extractBangkokSlipReference_(
      rawText
    );

  const receiverOk =
    (
      !!receiverName &&
      isExpectedAccountName(
        receiverName
      )
    ) ||
    hasExpectedBangkokSlipReceiverName_(
      rawText
    );

  const senderOk =
    !!senderName;

  const amountOk =
    typeof amount === "number" &&
    isFinite(amount) &&
    amount > 0;

  return {
    success:
      receiverOk &&
      senderOk &&
      amountOk,

    receiverOk: receiverOk,
    senderOk: senderOk,
    amountOk: amountOk,

    senderName: senderName,
    receiverName: receiverName,
    amount: amount,
    dateText: dateText,
    reference: reference,

    // รูปแบบให้ Flex เดิมอ่านได้
    data: {
      ocrVerified: true,
      verificationSource:
        "BANGKOK_BANK_OCR",
      amountInSlip: amount,
      rawSlip: {
        date: dateText || "",
        reference: reference || "",
        amount: {
          amount: amount
        },
        sender: {
          bank: {
            short: "BBL",
            nameTh:
              "ธนาคารกรุงเทพ"
          },
          account: {
            name: {
              th: senderName || ""
            },
            bank: {
              account: "-"
            }
          }
        },
        receiver: {
          bank: {
            short: "SCB",
            nameTh: "ธนาคารไทยพาณิชย์"
          },
          account: {
            name: {
              th: receiverName || ""
            },
            bank: {
              account:
                BANK_ACCOUNT_NUMBER
            }
          }
        }
      }
    }
  };
}


function handleBangkokBankSlipOcr_(
  replyToken,
  rawText,
  messageId
) {
  const parsed =
    parseBangkokBankSlipFromOcr_(
      rawText
    );

  console.log(
    "BBL_SLIP_OCR_RESULT",
    JSON.stringify({
      messageId:
        String(messageId || ""),
      senderOk:
        !!parsed.senderOk,
      receiverOk:
        !!parsed.receiverOk,
      amountOk:
        !!parsed.amountOk,
      amount:
        parsed.amount,
      hasReference:
        !!parsed.reference
    })
  );

  // ผู้รับไม่ตรงบัญชีร้าน
  if (
    parsed.receiverName &&
    !parsed.receiverOk
  ) {
    replyFlex(
      replyToken,
      buildInvalidSlipReceiverFlex(
        parsed.data
      )
    );

    return true;
  }

  // OCR อ่านข้อมูลสำคัญไม่ครบ
  if (!parsed.success) {
    const missing = [];

    if (!parsed.senderOk) {
      missing.push("ชื่อผู้โอน");
    }

    if (!parsed.receiverName) {
      missing.push("ชื่อผู้รับ");
    }

    if (!parsed.amountOk) {
      missing.push("ยอดเงิน");
    }

    replyText(
      replyToken,
      "⚠️ อ่านสลิปธนาคารกรุงเทพไม่ครบ (" +
      missing.join(", ") +
      ")\nกรุณาส่งรูปสลิปที่ชัดขึ้นอีกครั้ง"
    );

    return true;
  }

  // อ่านครบ + ผู้รับตรงบัญชีร้าน -> สำเร็จ
  replySlipVerifiedMessages(
    replyToken,
    parsed.data
  );

  return true;
}


function handleIncomingLineImage_(replyToken, messageId, ev) {
  const started = Date.now();
  let imageBlob = null;

  try {
    imageBlob = fetchLineImageBlob(messageId);

    try {
      const ocrResult = readLotteryBetImage_(imageBlob);

      if (
        ocrResult &&
        ocrResult.isBankSlip
      ) {
        const fullSlipOcrText =
          String(
            ocrResult.ocrText || ""
          );

        // =================================================
        // ธนาคารกรุงเทพ:
        // ไม่ตรวจชื่อผู้โอน / ชื่อผู้รับ / ยอดเงิน
        // ไม่เรียก Thunder API
        // ถ้าระบุได้ว่าเป็นสลิปต้นทาง Bangkok Bank
        // ให้ Flex "รับยอดเรียบร้อยครับ" ทันที
        // =================================================
        if (
          isBangkokBankOriginSlipOcr_(
            fullSlipOcrText
          )
        ) {
          console.log(
            "BBL_SLIP_AUTO_ACCEPT",
            JSON.stringify({
              messageId:
                String(messageId || ""),
              mode:
                "NO_VERIFY_AUTO_ACCEPT"
            })
          );

          replyFlex(
            replyToken,
            buildSlipThankYouFlex()
          );

          return;
        }

        // ธนาคารอื่น -> Thunder ตามเดิม
        console.log(
          "IMAGE_ROUTE_TO_SLIP_VERIFY",
          JSON.stringify({
            messageId:
              String(messageId || ""),
            reason:
              "OCR_BANK_SLIP_NON_BBL"
          })
        );
      }

      if (ocrResult && ocrResult.isLotteryBet) {
        const userKey = getLineUserKey(ev);

        console.log(
          "IMAGE_BET_OCR_OK",
          JSON.stringify({
            twoCount: (ocrResult.twoDigits || []).length,
            threeCount: (ocrResult.threeDigits || []).length,
            messageId: String(messageId || ""),
            hasUserKey: !!userKey
          })
        );

        if (!userKey) {
          replyText(
            replyToken,
            "⚠️ อ่านเลขจากรูปได้ แต่ไม่สามารถผูกโพยกับผู้ส่งได้ กรุณาส่งรูปใหม่อีกครั้ง"
          );
          return;
        }

        // เก็บเฉพาะข้อมูลที่จำเป็น ไม่เก็บ image/base64
        savePendingImageBet_(userKey, {
          messageId: String(messageId || ""),
          twoDigits: (ocrResult.twoDigits || []).slice(),
          threeDigits: (ocrResult.threeDigits || []).slice(),
          createdAt: Date.now()
        });

        // รูปอย่างเดียว "ยังไม่สร้างราคา และยังไม่ขึ้น Quick Reply"
        replyImageBetAwaitPrice_(
          replyToken,
          ocrResult
        );

        return;
      }
    } catch (ocrErr) {
      // OCR พังต้องไม่ทำให้ระบบตรวจสลิปเดิมพังตาม
      console.error(
        "IMAGE_BET_OCR_FAILED",
        JSON.stringify({
          messageId: String(messageId || ""),
          code: safeImageOcrErrorCode_(ocrErr)
        })
      );
    }

    // ไม่ใช่รูปโพย -> ตรวจสลิปตามระบบเดิม
    handleSlipImageMessage(
      replyToken,
      messageId,
      imageBlob
    );

  } catch (err) {
    console.error(
      "IMAGE_HANDLER_FATAL",
      JSON.stringify({
        messageId: String(messageId || ""),
        code: safeImageOcrErrorCode_(err)
      })
    );

    replyText(
      replyToken,
      "⚠️ อ่านรูปไม่สำเร็จ กรุณาลองส่งรูปใหม่อีกครั้ง"
    );

  } finally {
    logBotTiming(
      "handleIncomingLineImage_",
      Date.now() - started,
      {}
    );
  }
}

function safeImageOcrErrorCode_(err) {
  const msg = String(err && err.message ? err.message : err || "");

  if (/VISION_API_DISABLED/i.test(msg)) return "VISION_API_DISABLED";
  if (/VISION_HTTP_403/i.test(msg)) return "VISION_HTTP_403";
  if (/VISION_HTTP_429/i.test(msg)) return "VISION_HTTP_429";
  if (/VISION_HTTP_5/i.test(msg)) return "VISION_HTTP_5XX";
  if (/VISION_/i.test(msg)) return "VISION_ERROR";

  return "IMAGE_OCR_ERROR";
}


function getCloudPlatformAccessToken_() {
  const tokenUrl =
    "http://metadata.google.internal/computeMetadata/v1/instance/" +
    "service-accounts/default/token?scopes=" +
    encodeURIComponent("https://www.googleapis.com/auth/cloud-platform");

  const response = UrlFetchApp.fetch(tokenUrl, {
    method: "get",
    headers: {
      "Metadata-Flavor": "Google"
    },
    muteHttpExceptions: true,
    connectTimeoutSeconds: 2,
    timeoutSeconds: 8
  });

  const code = Number(response.getResponseCode() || 0);

  if (code < 200 || code >= 300) {
    throw new Error("VISION_TOKEN_HTTP_" + code);
  }

  let data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (_) {
    throw new Error("VISION_TOKEN_JSON");
  }

  if (!data || !data.access_token) {
    throw new Error("VISION_TOKEN_MISSING");
  }

  return String(data.access_token);
}


function runGoogleVisionTextOcr_(imageBlob) {
  if (!imageBlob) {
    throw new Error("VISION_IMAGE_MISSING");
  }

  if (
    !Utilities ||
    typeof Utilities.base64Encode !== "function"
  ) {
    throw new Error("VISION_BASE64_UNAVAILABLE");
  }

  const content = Utilities.base64Encode(
    imageBlob.getBytes()
  );

  const response = UrlFetchApp.fetch(
    "https://vision.googleapis.com/v1/images:annotate",
    {
      method: "post",
      headers: {
        Authorization:
          "Bearer " + getCloudPlatformAccessToken_()
      },
      contentType: "application/json",
      payload: JSON.stringify({
        requests: [
          {
            image: {
              content: content
            },
            features: [
              {
                type: "DOCUMENT_TEXT_DETECTION",
                maxResults: 1
              }
            ],
            imageContext: {
              languageHints: ["th", "en"]
            }
          }
        ]
      }),
      muteHttpExceptions: true,
      connectTimeoutSeconds: 5,
      timeoutSeconds: 25
    }
  );

  const code = Number(response.getResponseCode() || 0);
  const body = String(response.getContentText() || "");

  if (code === 403) {
    if (/has not been used|disabled|SERVICE_DISABLED/i.test(body)) {
      throw new Error("VISION_API_DISABLED");
    }
    throw new Error("VISION_HTTP_403");
  }

  if (code === 429) {
    throw new Error("VISION_HTTP_429");
  }

  if (code >= 500) {
    throw new Error("VISION_HTTP_5XX");
  }

  if (code < 200 || code >= 300) {
    throw new Error("VISION_HTTP_" + code);
  }

  let json;
  try {
    json = JSON.parse(body);
  } catch (_) {
    throw new Error("VISION_BAD_JSON");
  }

  const first =
    json &&
    json.responses &&
    json.responses[0]
      ? json.responses[0]
      : null;

  if (!first) {
    return {
      text: "",
      lines: []
    };
  }

  if (first.error) {
    throw new Error(
      "VISION_RESPONSE_ERROR_" +
      String(first.error.code || "UNKNOWN")
    );
  }

  const fullText = String(
    first.fullTextAnnotation && first.fullTextAnnotation.text
      ? first.fullTextAnnotation.text
      : (
          first.textAnnotations &&
          first.textAnnotations[0] &&
          first.textAnnotations[0].description
            ? first.textAnnotations[0].description
            : ""
        )
  );

  // เก็บ paragraph + bounding box ไว้ด้วย
  // เพื่อเลือกอ่านเฉพาะ "กรอบโพย" กลางภาพ ไม่อ่าน thumbnail/UI รอบนอก
  const paragraphs = [];

  const pages =
    first.fullTextAnnotation &&
    first.fullTextAnnotation.pages
      ? first.fullTextAnnotation.pages
      : [];

  pages.forEach(function(page) {
    const pageWidth = Number(page.width || 1);
    const pageHeight = Number(page.height || 1);

    (page.blocks || []).forEach(function(block) {
      (block.paragraphs || []).forEach(function(paragraph) {
        let paragraphText = "";

        (paragraph.words || []).forEach(function(word) {
          (word.symbols || []).forEach(function(symbol) {
            paragraphText += String(symbol.text || "");

            const breakType =
              symbol.property &&
              symbol.property.detectedBreak
                ? String(
                    symbol.property.detectedBreak.type || ""
                  )
                : "";

            if (
              breakType === "SPACE" ||
              breakType === "SURE_SPACE"
            ) {
              paragraphText += " ";
            } else if (
              breakType === "EOL_SURE_SPACE" ||
              breakType === "LINE_BREAK"
            ) {
              paragraphText += "\n";
            }
          });
        });

        paragraphText =
          String(paragraphText || "").trim();

        if (!paragraphText) return;

        const box = paragraph.boundingBox || {};
        const normalized =
          box.normalizedVertices &&
          box.normalizedVertices.length
            ? box.normalizedVertices
            : null;

        const rawVertices =
          normalized || box.vertices || [];

        if (!rawVertices.length) return;

        const xs = [];
        const ys = [];

        rawVertices.forEach(function(v) {
          let x = Number(v.x || 0);
          let y = Number(v.y || 0);

          if (!normalized) {
            x = x / pageWidth;
            y = y / pageHeight;
          }

          xs.push(x);
          ys.push(y);
        });

        paragraphs.push({
          text: paragraphText,
          xMin: Math.min.apply(null, xs),
          yMin: Math.min.apply(null, ys),
          xMax: Math.max.apply(null, xs),
          yMax: Math.max.apply(null, ys)
        });
      });
    });
  });

  return {
    text: fullText,
    lines: fullText
      .split(/\r?\n/)
      .map(function(line) {
        return String(line || "").trim();
      })
      .filter(Boolean),
    paragraphs: paragraphs
  };
}


function normalizeOcrDigits_(value) {
  const thai = "๐๑๒๓๔๕๖๗๘๙";

  return String(value || "")
    .replace(/[๐-๙]/g, function(ch) {
      return String(thai.indexOf(ch));
    })
    .replace(/[０-９]/g, function(ch) {
      return String(
        ch.charCodeAt(0) - "０".charCodeAt(0)
      );
    })
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}



// ============================================================================
// OCR PATCH 2026-09-21 — อ่านชุดเลขคั่นด้วย / และ :
//
// รองรับรูปแบบ:
//   17
//   10/13/15/16/
//   18/19
//   70/73/75/76/
//   78/79
//   11:14:74:77:72
//
// ผล:
//   17 10 13 15 16 18 19 70 73 75 76 78 79 11 14 74 77 72
// ============================================================================

// ============================================================================
// OCR STRICT BET FILTER 2026-09-21
//
// สำคัญ:
// อ่านเฉพาะ "เลขโพย" เท่านั้น
// วันที่ / เวลา / ปี / จำนวนสมาชิก / เบอร์โทร / LINE ID / เลข UI
// ต้องไม่ถูกนำไปนับรวมกับโพย
// ============================================================================
function isOcrDateLine_(rawLine) {
  const line = normalizeOcrDigits_(rawLine)
    .replace(/[／]/g, "/")
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .trim();

  if (!line) return false;

  // ISO: 2026-09-21 / 2026/09/21
  let m = line.match(
    /(?:^|\D)(\d{4})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]\s*(\d{1,2})(?:\D|$)/
  );

  if (m) {
    const month = Number(m[2]);
    const day = Number(m[3]);

    if (
      month >= 1 && month <= 12 &&
      day >= 1 && day <= 31
    ) {
      return true;
    }
  }

  // ไทย/ทั่วไป: 21/09/69, 21/9/2569, 21-09-69
  m = line.match(
    /(?:^|\D)(\d{1,2})\s*[\/.-]\s*(\d{1,2})\s*[\/.-]\s*(\d{2,4})(?:\D|$)/
  );

  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]);
    const yearText = String(m[3] || "");

    if (
      day >= 1 && day <= 31 &&
      month >= 1 && month <= 12 &&
      (yearText.length === 2 || yearText.length === 4)
    ) {
      return true;
    }
  }

  // Vision บางครั้งคืนวันที่แบบเว้นวรรค:
  // 2026 09 21 12:28 น.
  m = line.match(
    /(?:^|\D)(20\d{2}|25\d{2})\s+(\d{1,2})\s+(\d{1,2})(?:\D|$)/
  );

  if (m) {
    const month = Number(m[2]);
    const day = Number(m[3]);

    if (
      month >= 1 && month <= 12 &&
      day >= 1 && day <= 31
    ) {
      return true;
    }
  }

  return false;
}


function isOcrTimeLine_(rawLine) {
  const line = normalizeOcrDigits_(rawLine)
    .replace(/[：]/g, ":")
    .replace(/\s+/g, " ")
    .trim();

  if (!line) return false;

  // รับเฉพาะเวลา "ทั้งบรรทัด" เช่น
  // 14:23
  // 16:31 น.
  // เวลา 12:28 น.
  //
  // ไม่ตัดโพยแบบ 11:14:74:77:72 เพราะมี colon มากกว่าเวลา
  const m = line.match(
    /^(?:เวลา\s*)?(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2}))?\s*(?:น\.?)?$/
  );

  if (!m) return false;

  const hour = Number(m[1]);
  const minute = Number(m[2]);
  const second =
    m[3] == null ? 0 : Number(m[3]);

  return (
    hour >= 0 && hour <= 23 &&
    minute >= 0 && minute <= 59 &&
    second >= 0 && second <= 59
  );
}


function isOcrNonBetMetadataLine_(rawLine) {
  const line = normalizeOcrDigits_(rawLine)
    .replace(/\s+/g, " ")
    .trim();

  if (!line) return true;

  if (isOcrDateLine_(line)) return true;
  if (isOcrTimeLine_(line)) return true;

  // จำนวนสมาชิก/ตัวเลข UI เช่น (183), 99+
  if (/^\(?\s*\d{1,4}\s*\)?\s*\+?$/.test(line)) {
    // เลข 2 ตัวเดี่ยวอาจเป็นโพย เช่น 17
    // จึงตัดเฉพาะรูปแบบที่มีวงเล็บหรือ +
    if (/[()+]/.test(line)) return true;
  }

  // เบอร์โทร / เลขอ้างอิงยาว
  if (/(?:^|\D)\d{8,}(?:\D|$)/.test(line)) {
    return true;
  }

  // ข้อความ metadata/UI ที่ไม่ใช่โพย
  if (
    /(?:วันที่|เวลา|วันนี้|สมาชิก|ผู้ติดตาม|ครั้ง|วิว|views?|โทร|เบอร์|บัญชี|เลขที่รายการ|transaction|reference|ref\.?|line\s*id)/i.test(line)
  ) {
    return true;
  }

  return false;
}


function getOcrTwoDigitListInfo_(rawLine) {
  let line = normalizeOcrDigits_(rawLine)
    .replace(/[|｜]/g, "/")
    .replace(/[：]/g, ":")
    .replace(/[／]/g, "/")
    .replace(/\s+/g, " ")
    .trim();

  // วันที่/เวลา/เลข UI ห้ามเข้าระบบเลขโพย
  if (isOcrNonBetMetadataLine_(line)) {
    return {
      tokens: [],
      isSingle: false,
      isList: false,
      separator: "",
      ignoredMetadata: true
    };
  }

  if (!line) {
    return {
      tokens: [],
      isSingle: false,
      isList: false,
      separator: ""
    };
  }

  // OCR บางครั้งอ่าน 17 เป็น "1 7"
  // หรือ 10/13 เป็น "1 0 / 1 3"
  for (let pass = 0; pass < 3; pass++) {
    line = line.replace(
      /(^|[^\d])(\d)\s+(\d)(?=$|[^\d])/g,
      "$1$2$3"
    );
  }

  // หา list แบบ slash แม้มีข้อความ/emoji ติดหัวท้าย
  const slashMatch = line.match(
    /(\d{2}(?:\s*\/\s*\d{2})+\s*\/?)/
  );

  if (slashMatch) {
    const tokens = slashMatch[1].match(/\d{2}/g) || [];
    return {
      tokens: tokens,
      isSingle: false,
      isList: tokens.length >= 2,
      separator: "/"
    };
  }

  // colon อย่างน้อย 3 เลข ป้องกันเวลา 12:28
  const colonMatch = line.match(
    /(\d{2}(?:\s*:\s*\d{2}){2,}\s*:?)/
  );

  if (colonMatch) {
    const tokens = colonMatch[1].match(/\d{2}/g) || [];
    return {
      tokens: tokens,
      isSingle: false,
      isList: tokens.length >= 3,
      separator: ":"
    };
  }

  // แบบเว้นวรรคหลายเลข
  const spaceMatch = line.match(
    /(?:^|[^\d])((?:\d{2}\s+){1,}\d{2})(?:$|[^\d])/
  );

  if (spaceMatch) {
    const tokens = spaceMatch[1].match(/\d{2}/g) || [];
    return {
      tokens: tokens,
      isSingle: false,
      isList: tokens.length >= 2,
      separator: " "
    };
  }

  // เลขเดี่ยว เช่น 17; ยอมให้มี noise สั้น ๆ หัวท้าย
  const singleMatch = line.match(
    /^\D{0,3}(\d{2})\D{0,3}$/
  );

  if (singleMatch) {
    return {
      tokens: [singleMatch[1]],
      isSingle: true,
      isList: false,
      separator: ""
    };
  }

  return {
    tokens: [],
    isSingle: false,
    isList: false,
    separator: ""
  };
}


function extractLotteryBetNumbersFromOcrLines_(lines, fullText) {
  const twoDigits = [];
  const threeDigits = [];
  const evidenceLines = [];

  const bankText = normalizeOcrDigits_(fullText);
  const bankMarkers = [
    /โอนเงินสำเร็จ/i,
    /ผู้โอน/i,
    /ผู้รับ/i,
    /ธนาคาร/i,
    /เลขที่รายการ/i,
    /ค่าธรรมเนียม/i,
    /transaction/i,
    /transfer/i
  ].filter(function(re) {
    return re.test(bankText);
  }).length;

  // รูปสลิปให้ Thunder จัดการ ไม่เอาไปตีเป็นโพยหวย
  if (bankMarkers >= 2) {
    return {
      isLotteryBet: false,
      twoDigits: [],
      threeDigits: [],
      evidenceLines: []
    };
  }

  // ----------------------------------------------------------
  // PRE-PASS: อ่านชุดเลข 2 ตัวที่ OCR แบ่งหลายบรรทัด
  //
  // 17
  // 10/13/15/16/
  // 18/19
  // ...
  // 11:14:74:77:72
  //
  // เลขเดี่ยว 2 หลักจะรับเฉพาะเมื่ออยู่ติดกับแถวชุดเลข
  // เพื่อไม่ให้เวลา/ตัวเลขเดี่ยวใน UI ถูกเก็บมั่ว
  // ----------------------------------------------------------
  const normalizedOcrLines = (lines || []).map(function(rawLine) {
    return normalizeOcrDigits_(rawLine);
  });

  const twoListInfo = normalizedOcrLines.map(function(line) {
    return getOcrTwoDigitListInfo_(line);
  });

  const consumedTwoListRows = {};

  twoListInfo.forEach(function(info, index) {
    if (!info || !info.tokens || info.tokens.length === 0) return;

    let accept = !!info.isList;

    if (info.isSingle) {
      // เลขเดี่ยว เช่น 17 ให้รับถ้าอยู่ใกล้ชุด / หรือ :
      // รองรับ OCR ที่แทรกหัวข้อ/ข้อความ 1 บรรทัด
      const near = [];

      for (let d = 1; d <= 2; d++) {
        if (index - d >= 0) near.push(twoListInfo[index - d]);
        if (index + d < twoListInfo.length) near.push(twoListInfo[index + d]);
      }

      accept = near.some(function(item) {
        return !!(item && item.isList);
      });
    }

    if (!accept) return;

    info.tokens.forEach(function(n) {
      twoDigits.push(n);
    });

    evidenceLines.push(normalizedOcrLines[index]);
    consumedTwoListRows[index] = true;
  });

  // ถ้าเจอชุดคั่น / หรือ : ชัดเจน ให้ถือเป็นรายการหลักของรูป
  // และไม่ดูดเลขจาก thumbnail/ตารางสถิติด้านบน
  const strongSeparatedListCount = twoListInfo.reduce(function(sum, info) {
    if (!info || !info.isList) return sum;
    if (info.separator !== "/" && info.separator !== ":") return sum;
    return sum + (info.tokens || []).length;
  }, 0);

  const hasStrongSeparatedTwoList =
    strongSeparatedListCount >= 5 &&
    twoDigits.length >= 5;

  (lines || []).forEach(function(rawLine, lineIndex) {
    // แถวนี้ถูกอ่านใน pre-pass แล้ว ไม่อ่านซ้ำอีกรอบ
    if (consumedTwoListRows[lineIndex]) return;
    const line = normalizeOcrDigits_(rawLine);
    if (!line) return;

    // สำคัญ: วันที่ / เวลา / เบอร์ / เลข UI / metadata
    // ห้ามนำไปนับรวมกับโพย
    if (isOcrNonBetMetadataLine_(line)) {
      return;
    }

    // มี list หลักชัดเจนแล้ว: ห้ามเก็บเลขจาก thumbnail/ตารางด้านบนเพิ่ม
    if (hasStrongSeparatedTwoList) {
      return;
    }

    // เม็ดเดียว / ผลท้ายภาพ เช่น 826 - 26
    const resultPairs = [];
    line.replace(
      /(^|\D)(\d{3})\s*-\s*(\d{2})(?!\d)/g,
      function(_, lead, three, two) {
        resultPairs.push({
          three: three,
          two: two
        });
        return _;
      }
    );

    if (resultPairs.length > 0) {
      resultPairs.forEach(function(pair) {
        threeDigits.push(pair.three);
        twoDigits.push(pair.two);
      });

      evidenceLines.push(line);
      return;
    }

    // ดึงเฉพาะ token ตัวเลข 2 หรือ 3 หลัก
    // ต้องมีหลายตัวอยู่ในบรรทัดเดียวกัน จึงถือเป็นแถวเลขโพย
    const tokens = line.match(/(^|\D)(\d{2,3})(?!\d)/g) || [];

    const nums = tokens
      .map(function(token) {
        const m = String(token).match(/(\d{2,3})/);
        return m ? m[1] : "";
      })
      .filter(Boolean);

    const twos = nums.filter(function(n) {
      return /^\d{2}$/.test(n);
    });

    const threes = nums.filter(function(n) {
      return /^\d{3}$/.test(n);
    });

    if (twos.length >= 3) {
      twos.forEach(function(n) {
        twoDigits.push(n);
      });
      evidenceLines.push(line);
    }

    if (threes.length >= 3) {
      threes.forEach(function(n) {
        threeDigits.push(n);
      });
      evidenceLines.push(line);
    }
  });

  // ต้องมีหลักฐานเป็น "แถวตัวเลข" อย่างน้อยหนึ่งฝั่ง
  // ป้องกันเลขวันที่/บัญชี/ข้อความทั่วไปถูกตีเป็นโพย
  const isLotteryBet =
    twoDigits.length >= 3 ||
    threeDigits.length >= 3;

  return {
    isLotteryBet: isLotteryBet,
    twoDigits: twoDigits,
    threeDigits: threeDigits,
    evidenceLines: evidenceLines
  };
}



// ============================================================================
// IMAGE BET PENDING PRICE 2026-09-21
//
// รูป -> เก็บเลข -> รอราคา -> Quick Reply เลือกหวย
// ============================================================================
const IMAGE_BET_PENDING_TTL_SECONDS_ = 10 * 60;
const IMAGE_BET_PENDING_SHEET_ = "OCR_PENDING";


function pendingImageBetKey_(userKey) {
  return "pending_image_bet_" + String(userKey || "");
}


function getPendingImageBetSheet_() {
  const ss =
    SpreadsheetApp.openById(SHEET_ID);

  let sh =
    ss.getSheetByName(
      IMAGE_BET_PENDING_SHEET_
    );

  if (!sh) {
    sh =
      ss.insertSheet(
        IMAGE_BET_PENDING_SHEET_
      );

    sh.appendRow([
      "userKey",
      "payloadJson",
      "updatedAtMs"
    ]);
  }

  if (sh.getLastRow() < 1) {
    sh.appendRow([
      "userKey",
      "payloadJson",
      "updatedAtMs"
    ]);
  }

  return sh;
}


function savePendingImageBet_(userKey, data) {
  if (!userKey) return;

  const payload = Object.assign(
    {},
    data || {},
    {
      updatedAt: Date.now()
    }
  );

  const raw =
    JSON.stringify(payload);

  // Cache = เร็ว
  try {
    CacheService
      .getScriptCache()
      .put(
        pendingImageBetKey_(userKey),
        raw,
        IMAGE_BET_PENDING_TTL_SECONDS_
      );
  } catch (_) {}

  // Sheet = durable ข้าม request / ข้าม Cloud Run instance
  try {
    const sh =
      getPendingImageBetSheet_();

    const lastRow =
      sh.getLastRow();

    let foundRow = 0;

    if (lastRow >= 2) {
      const rows =
        sh.getRange(
          2,
          1,
          lastRow - 1,
          3
        ).getDisplayValues();

      for (let i = rows.length - 1; i >= 0; i--) {
        if (
          String(rows[i][0] || "") ===
          String(userKey)
        ) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow) {
      sh.getRange(foundRow, 2).setValue(raw);
      sh.getRange(foundRow, 3).setValue(
        String(Date.now())
      );
    } else {
      sh.appendRow([
        String(userKey),
        raw,
        String(Date.now())
      ]);
    }

    console.log(
      "IMAGE_BET_PENDING_SAVED",
      JSON.stringify({
        userKeyTail:
          String(userKey).slice(-8),
        twoCount:
          (payload.twoDigits || []).length,
        threeCount:
          (payload.threeDigits || []).length
      })
    );

  } catch (err) {
    console.error(
      "IMAGE_BET_PENDING_SAVE_SHEET_FAILED",
      JSON.stringify({
        code: String(
          err && err.message
            ? err.message
            : err
        ).slice(0, 120)
      })
    );
  }
}


function loadPendingImageBet_(userKey) {
  if (!userKey) return null;

  // 1) ลอง Cache ก่อน
  try {
    const raw =
      CacheService
        .getScriptCache()
        .get(
          pendingImageBetKey_(userKey)
        );

    if (raw) {
      const parsed =
        JSON.parse(raw);

      const age =
        Date.now() -
        Number(
          parsed.updatedAt ||
          parsed.createdAt ||
          0
        );

      if (
        age >= 0 &&
        age <=
          IMAGE_BET_PENDING_TTL_SECONDS_ *
          1000
      ) {
        return parsed;
      }
    }
  } catch (_) {}

  // 2) Cache ไม่มี -> อ่านจาก Sheet
  try {
    const sh =
      getPendingImageBetSheet_();

    const lastRow =
      sh.getLastRow();

    if (lastRow < 2) {
      return null;
    }

    const rows =
      sh.getRange(
        2,
        1,
        lastRow - 1,
        3
      ).getDisplayValues();

    for (let i = rows.length - 1; i >= 0; i--) {
      const rowUserKey =
        String(rows[i][0] || "");

      if (
        rowUserKey !==
        String(userKey)
      ) {
        continue;
      }

      const rowNo = i + 2;
      const raw =
        String(rows[i][1] || "");

      let parsed = null;

      try {
        parsed =
          JSON.parse(raw);
      } catch (_) {
        sh.deleteRow(rowNo);
        return null;
      }

      const updatedAt =
        Number(
          rows[i][2] ||
          parsed.updatedAt ||
          parsed.createdAt ||
          0
        );

      const age =
        Date.now() - updatedAt;

      if (
        !updatedAt ||
        age >
          IMAGE_BET_PENDING_TTL_SECONDS_ *
          1000
      ) {
        sh.deleteRow(rowNo);
        return null;
      }

      // เติมกลับ Cache เพื่อให้ครั้งถัดไปเร็ว
      try {
        CacheService
          .getScriptCache()
          .put(
            pendingImageBetKey_(userKey),
            raw,
            IMAGE_BET_PENDING_TTL_SECONDS_
          );
      } catch (_) {}

      console.log(
        "IMAGE_BET_PENDING_LOADED_FROM_SHEET",
        JSON.stringify({
          userKeyTail:
            String(userKey).slice(-8),
          twoCount:
            (parsed.twoDigits || []).length,
          threeCount:
            (parsed.threeDigits || []).length
        })
      );

      return parsed;
    }

  } catch (err) {
    console.error(
      "IMAGE_BET_PENDING_LOAD_SHEET_FAILED",
      JSON.stringify({
        code: String(
          err && err.message
            ? err.message
            : err
        ).slice(0, 120)
      })
    );
  }

  return null;
}


function clearPendingImageBet_(userKey) {
  if (!userKey) return;

  try {
    CacheService
      .getScriptCache()
      .remove(
        pendingImageBetKey_(userKey)
      );
  } catch (_) {}

  try {
    const sh =
      getPendingImageBetSheet_();

    const lastRow =
      sh.getLastRow();

    if (lastRow < 2) return;

    const rows =
      sh.getRange(
        2,
        1,
        lastRow - 1,
        1
      ).getDisplayValues();

    // ลบจากล่างขึ้นบน เผื่อมีซ้ำจากเวอร์ชันเก่า
    for (let i = rows.length - 1; i >= 0; i--) {
      if (
        String(rows[i][0] || "") ===
        String(userKey)
      ) {
        sh.deleteRow(i + 2);
      }
    }

  } catch (err) {
    console.error(
      "IMAGE_BET_PENDING_CLEAR_SHEET_FAILED",
      JSON.stringify({
        code: String(
          err && err.message
            ? err.message
            : err
        ).slice(0, 120)
      })
    );
  }
}


// รองรับ:
//   10*10
//   กับ10*10
//   กัน10*10
//   กลับ10*10
//   บลก10*10
//   บต10*10
//   10x10
//   10×10
//
// ความหมายสำหรับ OCR:
//   2 ตัว -> บลก ราคา ซ้าย*ขวา
//   3 ตัว -> บต  ราคา ซ้าย*ขวา
//
// ถ้ารูปมีทั้ง 2 ตัวและ 3 ตัว ราคาเดียวกันจะใช้กับทั้งสองกลุ่ม

// ============================================================================
// OCR PRICE + LOTTO 2026-09-21
//
// หลังส่งรูป ลูกค้าพิมพ์ได้ เช่น:
//   10*10
//   กับ10*10
//   กัน10*10
//   บลก10*10
//   บต10*10
//   10×10
//
// และใส่ชื่อหวยในข้อความเดียวกันได้:
//   ฮานอยปกติ กับ10*10
//   กับ10*10 ฮานอยปกติ
//   ฮานอยปกติ 10×10
//
// ถ้าชื่อหวยตรง/resolve ได้ในระบบ -> ส่ง Flex ทันที
// ถ้าไม่มีชื่อหรือชื่อกำกวม/ไม่ตรง -> Quick Reply ให้เลือกหวย
// ============================================================================
function parseImageBetPriceAndLotto_(text) {
  let raw = cleanIncomingLineText_(text)
    .replace(/[×✕✖⨯xX]/g, "*")
    .replace(/\s+/g, " ")
    .trim();

  if (!raw) return null;

  // รองรับ:
  // 10*10
  // กับ10*10
  // กัน10*10
  // บลก10*10
  // บต10*10
  // 10×10
  //
  // และใส่ชื่อหวยหน้า/หลังราคา:
  // ฮานอย 10*10
  // ฮานอย กับ10*10
  // 10*10 ฮานอย
  //
  // ใช้ regex literal ตรง ๆ เพื่อไม่ให้เกิดปัญหา escape ใน RegExp constructor
  const priceRe =
    /(?:ตรงโต๊ดกลับ|บนล่างกลับ|ตรงโต๊ด|ตัวละ|ราคา|บลก|บกล|บนล่าง|บต|กับ|กัน|กลับ)?\s*=?\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)/i;

  const m = priceRe.exec(raw);

  if (!m) return null;

  const first = Number(m[1]) || 0;
  const second = Number(m[2]) || 0;

  if (first <= 0 && second <= 0) {
    return null;
  }

  const before =
    raw.slice(0, Number(m.index || 0));

  const after =
    raw.slice(
      Number(m.index || 0) +
      String(m[0] || "").length
    );

  // เหลือข้อความที่คาดว่าเป็นชื่อหวย
  let lottoText =
    (before + " " + after)
      .replace(/^[\s:=\-–—|/]+|[\s:=\-–—|/]+$/g, "")
      .replace(/\s+/g, " ")
      .trim();

  return {
    first: first,
    second: second,
    raw: String(text || "").trim(),
    lottoText: lottoText,
    matchedPriceText: String(m[0] || "")
  };
}

function parseImageBetPriceText_(text) {
  const parsed =
    parseImageBetPriceAndLotto_(text);

  if (!parsed) return null;

  return {
    first: parsed.first,
    second: parsed.second,
    raw: parsed.raw
  };
}


function pendingImageBetToOcrResult_(pending) {
  return {
    isLotteryBet: true,
    twoDigits: (pending && pending.twoDigits) || [],
    threeDigits: (pending && pending.threeDigits) || [],
    evidenceLines: []
  };
}


function buildLotteryBetTextFromImageWithPrice_(ocrResult, price) {
  const lines = [];

  const first =
    Number(price && price.first) || 0;

  const second =
    Number(price && price.second) || 0;

  const two = (ocrResult.twoDigits || [])
    .filter(function(n) {
      return /^\d{2}$/.test(String(n));
    });

  const three = (ocrResult.threeDigits || [])
    .filter(function(n) {
      return /^\d{3}$/.test(String(n));
    });

  if (two.length > 0) {
    lines.push(
      two.join(" ") +
      " บลก " +
      first +
      "*" +
      second
    );
  }

  if (three.length > 0) {
    lines.push(
      three.join(" ") +
      " บต " +
      first +
      "*" +
      second
    );
  }

  return lines.join("\n");
}


function replyImageBetAwaitPrice_(token, ocrResult) {
  const two = ocrResult.twoDigits || [];
  const three = ocrResult.threeDigits || [];
  const lines = [];

  if (two.length > 0) {
    lines.push("2 ตัว:");
    lines.push(two.join(" "));
  }

  if (three.length > 0) {
    if (lines.length > 0) {
      lines.push("");
    }

    lines.push("3 ตัว:");
    lines.push(three.join(" "));
  }

  lines.push("");
  lines.push("กรุณาพิมพ์ชื่อหวย และราคา");
  lines.push("เช่น ฮานอย บลก 10*10");

  replyText(
    token,
    lines.join("\n")
  );
}



function resolveImageBetLottoFromText_(lottoText, lottoData) {
  const q = String(lottoText || "").trim();

  if (!q) {
    return {
      matched: false,
      reason: "NO_NAME"
    };
  }

  let match =
    findLottoSafe_(
      q,
      lottoData
    );

  if (!match) {
    return {
      matched: false,
      reason: "NOT_FOUND"
    };
  }

  // alias กำกวม: ถ้าเหลือหวยเปิดเพียงรายการเดียว ให้เลือกได้
  if (match.ambiguous) {
    const openCandidates =
      (match.candidates || [])
        .filter(function(c) {
          return (
            c &&
            c.name &&
            c.closeTime &&
            !isClosed(c.closeTime)
          );
        });

    if (openCandidates.length === 1) {
      return {
        matched: true,
        name: openCandidates[0].name,
        closeTime: openCandidates[0].closeTime,
        matchType: "single-open"
      };
    }

    return {
      matched: false,
      reason: "AMBIGUOUS"
    };
  }

  if (!match.name || !match.closeTime) {
    return {
      matched: false,
      reason: "NOT_FOUND"
    };
  }

  if (isClosed(match.closeTime)) {
    return {
      matched: false,
      reason: "CLOSED",
      name: match.name,
      closeTime: match.closeTime
    };
  }

  return {
    matched: true,
    name: match.name,
    closeTime: match.closeTime,
    matchType: match.matchType || "safe"
  };
}


function sendPendingImageBetFlexDirect_(
  replyToken,
  ev,
  lottoName,
  closeTime,
  betText
) {
  const keywords = loadKeywords();
  const doorMap = loadDoors();

  const result =
    parseMessage(
      betText,
      keywords,
      doorMap
    );

  if (!result || !result.success) {
    replyText(
      replyToken,
      String(
        result && result.error
          ? result.error
          : "❌ ไม่สามารถอ่านโพยได้"
      )
    );

    return false;
  }

  try {
    saveParsedBetFromLine_(
      ev,
      lottoName,
      closeTime,
      (
        String(lottoName || "") +
        "\n" +
        String(betText || "")
      ).trim(),
      result.groups,
      result.totalAmount,
      result.totalDoors
    );
  } catch (saveErr) {
    console.error(
      "IMAGE_BET_SAVE_FATAL " +
      JSON.stringify({
        lottoName: String(lottoName || ""),
        userId: String(
          ev &&
          ev.source &&
          ev.source.userId
            ? ev.source.userId
            : ""
        ),
        message: String(
          saveErr && saveErr.message
            ? saveErr.message
            : saveErr
        )
      })
    );

    replyText(
      replyToken,
      "⚠️ ระบบบันทึกโพยขัดข้อง กรุณาลองส่งราคาและชื่อหวยใหม่อีกครั้ง"
    );

    return false;
  }

  replyBetFlex_(
    replyToken,
    buildFlex({
      name: lottoName,
      closeTime: closeTime,
      groups: result.groups,
      totalDoors: result.totalDoors,
      totalAmount: result.totalAmount
    })
  );

  return true;
}


function tryHandlePendingImageBetPrice_(
  replyToken,
  userKey,
  text,
  ev
) {
  const pending =
    loadPendingImageBet_(userKey);

  if (!pending) {
    const rawText =
      cleanIncomingLineText_(text)
        .replace(/[×✕✖⨯xX]/g, "*")
        .trim();

    // แจ้ง pending หาย "เฉพาะกรณีที่ข้อความเป็นราคาอย่างเดียว"
    // เช่น 10*10 / กับ10*10 / บลก10*10
    //
    // แต่ถ้าเป็นข้อความจาก Quick Reply:
    //   สิงคโปร์
    //   07 10 ... บลก 10*10
    // ต้องปล่อยให้ parser โพยปกติทำงานต่อ เพื่อส่ง Flex
    const isPriceOnly =
      /^(?:ตรงโต๊ดกลับ|บนล่างกลับ|ตรงโต๊ด|ตัวละ|ราคา|บลก|บกล|บนล่าง|บต|กับ|กัน|กลับ)?\s*=?\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?$/i.test(
        rawText
      );

    if (isPriceOnly) {
      console.warn(
        "IMAGE_BET_PENDING_MISSING",
        JSON.stringify({
          userKeyTail:
            String(userKey || "").slice(-8),
          textHead:
            rawText.slice(0, 40)
        })
      );

      replyText(
        replyToken,
        "⚠️ ไม่พบรูปโพยที่รอราคาอยู่ กรุณาส่งรูปใหม่ แล้วพิมพ์ราคาอีกครั้ง"
      );

      return true;
    }

    // มีชื่อหวย/เลขโพย/ข้อความอื่นติดมาด้วย
    // เช่น payload จาก Quick Reply -> ให้ flow ปกติอ่านโพยต่อ
    return false;
  }

  console.log(
    "IMAGE_BET_PENDING_TEXT",
    JSON.stringify({
      hasPending: true,
      textHead: String(text || "").slice(0, 60),
      twoCount: (pending.twoDigits || []).length,
      threeCount: (pending.threeDigits || []).length
    })
  );

  // ยกเลิก pending รูปได้
  if (
    /^(?:ยกเลิก|cancel)$/i.test(
      String(text || "").trim()
    )
  ) {
    clearPendingImageBet_(userKey);

    replyText(
      replyToken,
      "ยกเลิกโพยจากรูปแล้วครับ"
    );

    return true;
  }

  const parsedInput =
    parseImageBetPriceAndLotto_(text);

  // ถ้ามีโพยจากรูปค้างอยู่ และข้อความดูเหมือนพยายามส่งราคา
  // ห้ามเงียบเด็ดขาด — แจ้งรูปแบบที่รองรับ
  if (!parsedInput) {
    const looksPriceLike =
      /\d+\s*[×✕✖⨯xX*]\s*\d+/.test(
        String(text || "")
      );

    if (looksPriceLike) {
      replyText(
        replyToken,
        [
          "⚠️ อ่านราคาไม่สำเร็จ",
          "กรุณาพิมพ์ราคา เช่น",
          "10*10",
          "กับ10*10",
          "ฮานอย 10*10"
        ].join("\n")
      );

      return true;
    }

    return false;
  }

  const price = {
    first: parsedInput.first,
    second: parsedInput.second,
    raw: parsedInput.raw
  };

  console.log(
    "IMAGE_BET_PRICE_PARSED",
    JSON.stringify({
      first: price.first,
      second: price.second,
      lottoText:
        String(
          parsedInput.lottoText || ""
        ).slice(0, 60)
    })
  );

  const ocrResult =
    pendingImageBetToOcrResult_(pending);

  const betText =
    buildLotteryBetTextFromImageWithPrice_(
      ocrResult,
      price
    );

  if (!betText) {
    clearPendingImageBet_(userKey);

    replyText(
      replyToken,
      "⚠️ ไม่พบเลขโพยจากรูป กรุณาส่งรูปใหม่อีกครั้ง"
    );

    return true;
  }

  const lottoData =
    loadLotto();

  const resolved =
    resolveImageBetLottoFromText_(
      parsedInput.lottoText,
      lottoData
    );

  // =====================================================
  // มีชื่อหวยและตรงกับระบบ -> ส่ง Flex ทันที
  // =====================================================
  if (resolved.matched) {
    const sent =
      sendPendingImageBetFlexDirect_(
        replyToken,
        ev,
        resolved.name,
        resolved.closeTime,
        betText
      );

    if (sent) {
      clearPendingImageBet_(userKey);
    }

    return true;
  }

  // =====================================================
  // ไม่มีชื่อ / ชื่อไม่ตรง / ชื่อกำกวม / หวยที่พิมพ์ปิดแล้ว
  // -> ให้ลูกค้าเลือกหวยจาก Quick Reply
  //
  // ข้อความ:
  // "กรุณาตรวจเลข แล้วกดเลือกชื่อหวยด้านล่าง"
  // =====================================================
  replyImageBetLottoChoice_(
    replyToken,
    lottoData,
    ocrResult,
    betText,
    price
  );

  clearPendingImageBet_(userKey);
  return true;
}



// ============================================================================
// OCR ROI 2026-09-21
//
// อ่านเฉพาะกรอบโพยหลักตรงกลางภาพ
// ตามตัวอย่างผู้ใช้:
//
//        X = 8%  ถึง 84% ของความกว้าง
//        Y = 24% ถึง 81% ของความสูง
//
// ใช้ bounding box จาก Google Vision จึงไม่ต้อง crop ไฟล์จริง
// ข้อความจาก thumbnail ด้านบน / เวลา / เมนู LINE ด้านนอกจะไม่ถูกนำไป parse
// ============================================================================
const IMAGE_BET_ROI_ = {
  xMin: 0.08,
  xMax: 0.84,
  yMin: 0.24,
  yMax: 0.81
};


function selectLotteryOcrRegion_(ocr) {
  const paragraphs =
    ocr && ocr.paragraphs
      ? ocr.paragraphs
      : [];

  if (!paragraphs.length) {
    return {
      lines: [],
      text: "",
      usedGeometry: false
    };
  }

  const selected = paragraphs.filter(function(p) {
    const cx =
      (Number(p.xMin || 0) + Number(p.xMax || 0)) / 2;

    const cy =
      (Number(p.yMin || 0) + Number(p.yMax || 0)) / 2;

    return (
      cx >= IMAGE_BET_ROI_.xMin &&
      cx <= IMAGE_BET_ROI_.xMax &&
      cy >= IMAGE_BET_ROI_.yMin &&
      cy <= IMAGE_BET_ROI_.yMax
    );
  });

  // เรียงบนลงล่าง แล้วซ้ายไปขวา
  selected.sort(function(a, b) {
    const ay = Number(a.yMin || 0);
    const by = Number(b.yMin || 0);

    if (Math.abs(ay - by) > 0.01) {
      return ay - by;
    }

    return Number(a.xMin || 0) - Number(b.xMin || 0);
  });

  const lines = [];

  selected.forEach(function(p) {
    String(p.text || "")
      .split(/\r?\n/)
      .map(function(line) {
        return String(line || "").trim();
      })
      .filter(Boolean)
      .forEach(function(line) {
        lines.push(line);
      });
  });

  return {
    lines: lines,
    text: lines.join("\n"),
    usedGeometry: true,
    selectedCount: selected.length
  };
}



// ============================================================================
// IMAGE CLASSIFIER 2026-09-21
// ตรวจสลิปโอนเงินจาก OCR ทั้งภาพก่อนอ่านเป็นโพย
// ============================================================================
function classifyBankSlipFromOcrText_(rawText) {
  const text = String(rawText || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) {
    return {
      isBankSlip: false,
      score: 0,
      markers: []
    };
  }

  const markers = [];
  let score = 0;

  function hit_(label, regex, weight) {
    if (regex.test(text)) {
      markers.push(label);
      score += Number(weight || 1);
    }
  }

  hit_(
    "transfer_success",
    /โอนเงินสำเร็จ|โอนสำเร็จ|ทำรายการสำเร็จ|รายการสำเร็จ|ชำระเงินสำเร็จ/i,
    5
  );

  hit_(
    "amount",
    /จำนวนเงิน|ยอดเงิน|ยอดโอน|จำนวน\s*[0-9,.]+\s*บาท/i,
    2
  );

  hit_(
    "from",
    /(?:^|\s)จาก(?:\s|$)|ผู้โอน|บัญชีผู้โอน/i,
    2
  );

  hit_(
    "to",
    /ไปที่|ไปยัง|ผู้รับ|บัญชีผู้รับ|รับเงิน/i,
    2
  );

  hit_(
    "reference",
    /เลขที่รายการ|เลขอ้างอิง|รหัสอ้างอิง|หมายเลขรายการ|reference|transaction/i,
    2
  );

  hit_("fee", /ค่าธรรมเนียม/i, 1);

  hit_(
    "account",
    /เลขที่บัญชี|บัญชีเลขที่|xxx+[-x0-9]*|บัญชี/i,
    1
  );

  hit_("qr", /QR|คิวอาร์/i, 1);

  hit_(
    "bank",
    /SCB|KBank|KBANK|K PLUS|Krungthai|KTB|กรุงไทย|กสิกร|ไทยพาณิชย์|กรุงเทพ|BBL|กรุงศรี|BAY|ttb|ทีเอ็มบีธนชาต/i,
    2
  );

  hit_(
    "date_time",
    /\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}|\d{1,2}:\d{2}/,
    1
  );

  const hasStrongTransferWord =
    markers.indexOf("transfer_success") !== -1;

  const hasSlipStructure =
    markers.some(function(m) {
      return (
        m === "bank" ||
        m === "amount" ||
        m === "from" ||
        m === "to" ||
        m === "reference"
      );
    });

  return {
    isBankSlip:
      hasStrongTransferWord ||
      (score >= 5 && hasSlipStructure),
    score: score,
    markers: markers
  };
}


function readLotteryBetImage_(imageBlob) {
  const ocr =
    runGoogleVisionTextOcr_(imageBlob);

  // 1) ตรวจจาก OCR ทั้งภาพก่อนว่าเป็นสลิปหรือไม่
  const slipClass =
    classifyBankSlipFromOcrText_(
      ocr && ocr.text
        ? ocr.text
        : ""
    );

  if (slipClass.isBankSlip) {
    console.log(
      "IMAGE_CLASSIFIED_BANK_SLIP",
      JSON.stringify({
        score: slipClass.score,
        markers: slipClass.markers
      })
    );

    return {
      isLotteryBet: false,
      isBankSlip: true,
      twoDigits: [],
      threeDigits: [],
      evidenceLines: [],
      ocrText:
        String(
          ocr && ocr.text
            ? ocr.text
            : ""
        ).slice(0, 1500)
    };
  }

  // 2) ไม่ใช่สลิป ค่อยอ่านกรอบโพย
  const region =
    selectLotteryOcrRegion_(ocr);

  if (
    !region.usedGeometry ||
    !region.lines ||
    region.lines.length === 0
  ) {
    console.log(
      "IMAGE_BET_OCR_ROI_EMPTY",
      JSON.stringify({
        paragraphCount:
          (ocr.paragraphs || []).length
      })
    );

    return {
      isLotteryBet: false,
      isBankSlip: false,
      twoDigits: [],
      threeDigits: [],
      evidenceLines: [],
      ocrText: ""
    };
  }

  const strictBetLines =
    (region.lines || []).filter(function(line) {
      return !isOcrNonBetMetadataLine_(line);
    });

  const strictBetText =
    strictBetLines.join("\n");

  const parsed =
    extractLotteryBetNumbersFromOcrLines_(
      strictBetLines,
      strictBetText
    );

  console.log(
    "IMAGE_BET_OCR_DEBUG",
    JSON.stringify({
      isLotteryBet: !!parsed.isLotteryBet,
      isBankSlip: false,
      twoCount:
        (parsed.twoDigits || []).length,
      threeCount:
        (parsed.threeDigits || []).length,
      roi: IMAGE_BET_ROI_,
      selectedParagraphs:
        Number(region.selectedCount || 0),
      regionLines:
        strictBetLines.slice(0, 40),
      regionTextPreview:
        String(strictBetText || "").slice(0, 1500)
    })
  );

  parsed.isBankSlip = false;
  parsed.ocrText = strictBetText;
  return parsed;
}

function buildLotteryBetTextFromImage_(ocrResult) {
  const lines = [];

  const two = (ocrResult.twoDigits || [])
    .filter(function(n) {
      return /^\d{2}$/.test(String(n));
    });

  const three = (ocrResult.threeDigits || [])
    .filter(function(n) {
      return /^\d{3}$/.test(String(n));
    });

  if (two.length > 0) {
    lines.push(
      two.join(" ") +
      " บลก " +
      IMAGE_BET_TWO_TOP_ +
      "*" +
      IMAGE_BET_TWO_BOTTOM_
    );
  }

  if (three.length > 0) {
    // ใช้ บต แบบชัดเจน เพื่อให้ parser อ่านเป็น ตรง+โต๊ด
    lines.push(
      three.join(" ") +
      " บต " +
      IMAGE_BET_THREE_DIRECT_ +
      "*" +
      IMAGE_BET_THREE_TOD_
    );
  }

  return lines.join("\n");
}


function replyImageBetLottoChoice_(token, lottoData, ocrResult, betText, price) {
  const candidates = getOpenLottoCandidates(lottoData);

  if (!candidates || candidates.length === 0) {
    replyText(token, "❌ ตอนนี้หวยปิดรับแล้วทั้งหมด");
    return;
  }

  const two = ocrResult.twoDigits || [];
  const three = ocrResult.threeDigits || [];

  const preview = [];
  preview.push("📷 อ่านโพยจากรูปได้แล้ว");

  const first =
    Number(price && price.first) || IMAGE_BET_TWO_TOP_;

  const second =
    Number(price && price.second) || IMAGE_BET_TWO_BOTTOM_;

  if (two.length > 0) {
    preview.push(
      "2 ตัว: " +
      two.join(" ") +
      " บลก " +
      first +
      "*" +
      second
    );
  }

  if (three.length > 0) {
    preview.push(
      "3 ตัว: " +
      three.join(" ") +
      " บต " +
      first +
      "*" +
      second
    );
  }

  preview.push("");
  preview.push("✅ รับราคาแล้ว");
  preview.push("กรุณาตรวจเลข แล้วกดเลือกชื่อหวยด้านล่าง");

  const items = candidates
    .slice(0, 13)
    .map(function(c) {
      return {
        type: "action",
        action: {
          type: "message",
          label: String(c.name || "").substring(0, 20),
          text: (
            String(c.name || "") +
            "\n" +
            betText
          ).trim()
        }
      };
    });

  const response = UrlFetchApp.fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "post",
      headers: {
        Authorization:
          "Bearer " + getChannelAccessToken_()
      },
      contentType: "application/json",
      payload: JSON.stringify({
        replyToken: token,
        messages: [
          {
            type: "text",
            text: preview.join("\n"),
            quickReply: {
              items: items
            }
          }
        ]
      }),
      muteHttpExceptions: true
    }
  );

  const code = Number(response.getResponseCode() || 0);

  if (code < 200 || code >= 300) {
    throw new Error("IMAGE_BET_LINE_REPLY_" + code);
  }
}


// Regression test สำหรับภาพตัวอย่าง 21/9/69
function testLotteryImageNumberExtraction_() {
  const sampleLines = [
    "2026 09 21 12:28 น.",
    "เกาหลี+VIP",
    "2 - 6",
    "23 25 26 28",
    "61 63 65 68",
    "523 826 763 465",
    "เม็ดเดียว",
    "826 - 26",
    "LINE @332cfcnj"
  ];

  const parsed = extractLotteryBetNumbersFromOcrLines_(
    sampleLines,
    sampleLines.join("\n")
  );

  const expected2 =
    "23 25 26 28 61 63 65 68 26";

  const expected3 =
    "523 826 763 465 826";

  if (parsed.twoDigits.join(" ") !== expected2) {
    throw new Error(
      "IMAGE OCR 2 ตัวผิด: " +
      parsed.twoDigits.join(" ")
    );
  }

  if (parsed.threeDigits.join(" ") !== expected3) {
    throw new Error(
      "IMAGE OCR 3 ตัวผิด: " +
      parsed.threeDigits.join(" ")
    );
  }

  const betText = buildLotteryBetTextFromImage_(parsed);

  return {
    success: true,
    twoDigits: parsed.twoDigits,
    threeDigits: parsed.threeDigits,
    betText: betText
  };
}


// ================== SLIP VERIFY ==================
function handleSlipImageMessage(replyToken, messageId, prefetchedImageBlob) {
  const started = Date.now();

  try {
    console.log("SLIP_START", JSON.stringify({
      messageId: String(messageId || "")
    }));

    const imageBlob = prefetchedImageBlob || fetchLineImageBlob(messageId);

    console.log("SLIP_LINE_IMAGE_OK", JSON.stringify({
      messageId: String(messageId || ""),
      bytes: imageBlob.getBytes().length
    }));

    const result = verifyThunderBankSlip(imageBlob, messageId);

    console.log("SLIP_THUNDER_OK", JSON.stringify({
      duplicate: !!(result && result.isDuplicate),
      receiverOk: !!(result && isExpectedSlipReceiver(result))
    }));

    if (!result) {
      throw new Error("Thunder ไม่ส่งข้อมูลสลิปกลับมา");
    }

    // สลิปซ้ำ
    if (result.isDuplicate) {
      replyFlex(
        replyToken,
        buildDuplicateSlipFlex(result)
      );
      return;
    }

    // บัญชีปลายทางไม่ตรง
    if (!isExpectedSlipReceiver(result)) {
      replyFlex(
        replyToken,
        buildInvalidSlipReceiverFlex(result)
      );
      return;
    }

    // สลิปถูกต้อง
    replySlipVerifiedMessages(
      replyToken,
      result
    );

  } catch (err) {
    console.error(
      "SLIP_ERROR",
      JSON.stringify({
        messageId: String(messageId || ""),
        code: getSafeSlipErrorCode_(err)
      })
    );

    if (isSilentNonSlipImageError(err)) {
      return;
    }

    const errorMessage =
      buildSlipVerifyErrorMessage(err);

    if (errorMessage) {
      replyText(
        replyToken,
        errorMessage
      );
    }

  } finally {
    console.log(
      "SLIP_FINISH",
      JSON.stringify({
        messageId: String(messageId || ""),
        ms: Date.now() - started
      })
    );

    logBotTiming(
      "handleSlipImageMessage",
      Date.now() - started,
      {}
    );
  }
}


// =====================================================
// SLIP VERIFY - Cloud Run compatibility helpers
// =====================================================

function replySlipVerifiedMessages(replyToken, data) {
  const messages = [
    buildSlipVerifyFlex(data),
    buildSlipThankYouFlex()
  ];

  const response = UrlFetchApp.fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + getChannelAccessToken_()
      },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: messages
      }),
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();

  console.log("SLIP_REPLY", JSON.stringify({
    status: code,
    messages: messages.length
  }));

  if (code < 200 || code >= 300) {
    throw new Error(
      "LINE_REPLY_ERROR " + code + ": " +
      response.getContentText()
    );
  }
}

function buildCleanSlipRow(label, value, valueColor) {
  return {
    type: "box",
    layout: "horizontal",
    alignItems: "center",
    margin: "xs",
    contents: [
      {
        type: "text",
        text: label,
        size: "xs",
        color: "#7B8190",
        flex: 2
      },
      {
        type: "text",
        text: safeText(value, "-"),
        size: "xs",
        color: valueColor,
        weight: "bold",
        align: "end",
        wrap: true,
        flex: 6
      }
    ]
  };
}

function buildCleanPartyRow(label, party, valueColor) {
  const name = getSlipPartyName(party);
  const bank = getSlipPartyBankText(party);
  const account = getSlipPartyAccount(party);

  return {
    type: "box",
    layout: "horizontal",
    margin: "sm",
    contents: [
      {
        type: "text",
        text: label,
        size: "xs",
        color: "#7B8190",
        flex: 2,
        gravity: "center"
      },
      {
        type: "box",
        layout: "vertical",
        flex: 6,
        contents: [
          {
            type: "text",
            text: name,
            size: "sm",
            color: valueColor,
            weight: "bold",
            align: "end",
            wrap: true
          },
          {
            type: "text",
            text: bank + " · " + account,
            size: "xxs",
            color: "#A0A5B0",
            align: "end",
            margin: "xs",
            wrap: true
          }
        ]
      }
    ]
  };
}

function buildSlipThankYouFlex() {
  return {
    type: "flex",
    altText: "รับยอดเรียบร้อย ขอบคุณครับ",
    contents: {
      type: "bubble",
      size: "kilo",
      body: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#FFFFFF",
        paddingAll: "16px",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: "#E9FBF5",
            cornerRadius: "12px",
            paddingAll: "13px",
            alignItems: "center",
            contents: [
              {
                type: "image",
                url: "https://cdn-icons-png.flaticon.com/512/190/190411.png",
                size: "22px",
                aspectMode: "fit",
                flex: 0
              },
              {
                type: "box",
                layout: "vertical",
                margin: "md",
                flex: 1,
                contents: [
                  {
                    type: "text",
                    text: "รับยอดเรียบร้อยครับ",
                    size: "sm",
                    weight: "bold",
                    color: "#087F5B"
                  },
                  {
                    type: "text",
                    text: "ขอบคุณที่มาอุดหนุนนะครับ",
                    size: "xxs",
                    color: "#4D7D6E",
                    margin: "xs"
                  }
                ]
              }
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: "#FFF8F1",
            cornerRadius: "12px",
            paddingAll: "13px",
            contents: [
              {
                type: "image",
                url: "https://cdn-icons-png.flaticon.com/512/595/595067.png",
                size: "20px",
                aspectMode: "fit",
                flex: 0
              },
              {
                type: "box",
                layout: "vertical",
                margin: "md",
                flex: 1,
                contents: [
                  {
                    type: "text",
                    text: "กรุณาตรวจสอบโพยทุกครั้ง",
                    size: "xs",
                    weight: "bold",
                    color: "#A45125",
                    wrap: true
                  },
                  {
                    type: "text",
                    text: "หากพบรายการไม่ถูกต้อง กรุณาแจ้งแอดมินทันทีครับ",
                    size: "xxs",
                    color: "#9B6A50",
                    margin: "xs",
                    wrap: true
                  }
                ]
              }
            ]
          }
        ]
      }
    }
  };
}

function buildDuplicateSlipFlex(data) {
  const raw = data.rawSlip || {};
  const sender = raw.sender || {};
  const receiver = raw.receiver || {};
  const amount = getSlipAmount(data);
  const red = "#EF4444";
  const deepRed = "#991B1B";
  const lightRed = "#FEF2F2";

  return {
    type: "flex",
    altText: "❌ สลิปนี้ถูกใช้แล้ว",
    contents: {
      type: "bubble",
      size: "kilo",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: red,
        paddingAll: "14px",
        contents: [
          {
            type: "text",
            text: "สลิปนี้ถูกใช้แล้ว",
            color: "#FFFFFF",
            size: "md",
            weight: "bold",
            align: "center"
          },
          {
            type: "text",
            text: "รายการนี้เคยถูกตรวจสอบแล้ว",
            color: "#FFF1F1",
            size: "xxs",
            margin: "xs",
            align: "center"
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "14px",
        spacing: "sm",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: lightRed,
            cornerRadius: "14px",
            paddingAll: "15px",
            contents: [
              { type: "text", text: "ยอดโอน", color: deepRed, size: "xs", weight: "bold", flex: 1 },
              { type: "text", text: "฿" + formatBahtNumber(amount), color: deepRed, size: "xl", weight: "bold", align: "end", flex: 2 }
            ]
          },
          buildCleanSlipRow("เมื่อ", formatSlipDate(raw.date), deepRed),
          { type: "separator", color: "#FECACA", margin: "xs" },
          buildCleanPartyRow("จาก", sender, deepRed),
          buildCleanPartyRow("ไปยัง", receiver, deepRed),
          {
            type: "box",
            layout: "vertical",
            backgroundColor: lightRed,
            cornerRadius: "10px",
            paddingAll: "10px",
            margin: "sm",
            contents: [
              {
                type: "text",
                text: "❗ รายการนี้เคยถูกตรวจแล้ว กรุณาอย่ารับยอดซ้ำ",
                color: deepRed,
                size: "xs",
                weight: "bold",
                align: "center",
                wrap: true
              }
            ]
          }
        ]
      }
    }
  };
}

function buildInvalidSlipReceiverFlex(data) {
  const raw = data.rawSlip || {};
  const receiver = raw.receiver || {};
  const amount = getSlipAmount(data);
  const red = "#EF4444";
  const deepRed = "#991B1B";
  const lightRed = "#FEF2F2";

  return {
    type: "flex",
    altText: "❌ สลิปไม่ถูกต้อง บัญชีปลายทางไม่ตรง",
    contents: {
      type: "bubble",
      size: "kilo",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: red,
        paddingAll: "14px",
        contents: [
          {
            type: "text",
            text: "สลิปไม่ถูกต้อง",
            color: "#FFFFFF",
            size: "md",
            weight: "bold",
            align: "center"
          },
          {
            type: "text",
            text: "บัญชีปลายทางไม่ตรง",
            color: "#FFF1F1",
            size: "xxs",
            margin: "xs",
            align: "center"
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "14px",
        spacing: "sm",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: lightRed,
            cornerRadius: "10px",
            paddingAll: "13px",
            contents: [
              { type: "text", text: "ยอดโอน", color: deepRed, size: "xs", weight: "bold", flex: 1 },
              { type: "text", text: "฿" + formatBahtNumber(amount), color: deepRed, size: "xl", weight: "bold", align: "end", flex: 2 }
            ]
          },
          buildCleanSlipRow("เมื่อ", formatSlipDate(raw.date), deepRed),
          { type: "separator", color: "#FECACA", margin: "xs" },
          buildCleanPartyRow("ไปยัง", receiver, deepRed),
          {
            type: "box",
            layout: "vertical",
            backgroundColor: lightRed,
            cornerRadius: "10px",
            paddingAll: "10px",
            spacing: "xs",
            margin: "sm",
            contents: [
              { type: "text", text: "บัญชีที่ต้องโอนเข้า", color: "#6B7280", size: "xxs", weight: "bold" },
              { type: "text", text: BANK_ACCOUNT_NAME, color: deepRed, size: "sm", weight: "bold", wrap: true },
              { type: "text", text: BANK_ACCOUNT_BANK + " · " + BANK_ACCOUNT_NUMBER, color: "#6B7280", size: "xxs", wrap: true }
            ]
          }
        ]
      }
    }
  };
}

function getSafeSlipErrorCode_(err) {
  const msg = String(err && err.message ? err.message : err);

  if (/THUNDER_TIMEOUT/i.test(msg)) return "THUNDER_TIMEOUT";
  if (/THUNDER_NETWORK/i.test(msg)) return "THUNDER_NETWORK";
  if (/THUNDER_RETRYABLE_HTTP/i.test(msg)) return "THUNDER_RETRYABLE_HTTP";
  if (/THUNDER_NON_JSON/i.test(msg)) return "THUNDER_NON_JSON";
  if (/SLIP_PENDING/i.test(msg)) return "SLIP_PENDING";
  if (/SLIP_NOT_FOUND/i.test(msg)) return "SLIP_NOT_FOUND";
  if (/VALIDATION_ERROR|QR_NOT_FOUND|NO_QR/i.test(msg)) return "NOT_A_SLIP";
  if (/THUNDER_ERROR/i.test(msg)) return "THUNDER_ERROR";

  return "SLIP_VERIFY_ERROR";
}


function buildSlipVerifyErrorMessage(err) {
  const code = getSafeSlipErrorCode_(err);

  if (code === "THUNDER_TIMEOUT") {
    return "⚠️ ระบบตรวจสลิปตอบสนองช้า\nกรุณารอสักครู่แล้วส่งสลิปใหม่อีกครั้ง";
  }

  // กรณี Thunder ตอบ 502 / non-JSON / upstream ชั่วคราว
  // ไม่ส่งข้อความใด ๆ กลับหาลูกค้า
  if (
    code === "THUNDER_RETRYABLE_HTTP" ||
    code === "THUNDER_NON_JSON"
  ) {
    return "";
  }

  if (code === "THUNDER_NETWORK") {
    return "⚠️ ระบบตรวจสลิปขัดข้องชั่วคราว\nกรุณาลองส่งสลิปใหม่อีกครั้ง";
  }

  if (code === "SLIP_PENDING") {
    return "⏳ สลิปนี้ยังรอดำเนินการจากธนาคาร\nกรุณารอสักครู่แล้วส่งใหม่อีกครั้ง";
  }

  if (code === "SLIP_NOT_FOUND" || code === "NOT_A_SLIP") {
    return "";
  }

  // ห้ามส่ง err.message ดิบ เพราะอาจมี curl command / Authorization token / API key
  return "❌ ตรวจสลิปไม่สำเร็จ\nกรุณาลองส่งสลิปใหม่อีกครั้ง";
}


function isSilentNonSlipImageError(err) {
  const msg = String(err && err.message ? err.message : err);

  return (
    msg.indexOf("SLIP_NOT_FOUND") !== -1 ||
    msg.indexOf("VALIDATION_ERROR") !== -1 ||
    msg.indexOf("QR_NOT_FOUND") !== -1 ||
    msg.indexOf("NO_QR") !== -1 ||
    msg.indexOf("Please provide either") !== -1 ||
    msg.indexOf("ไม่พบ QR") !== -1 ||
    msg.indexOf("ไม่พบคิวอาร์") !== -1
  );
}


function fetchLineImageBlob(messageId) {
  if (!messageId) throw new Error("ไม่พบรหัสรูปภาพจาก LINE");

  const started = Date.now();
  const response = UrlFetchApp.fetch(
    "https://api-data.line.me/v2/bot/message/" + encodeURIComponent(messageId) + "/content",
    {
      method: "get",
      headers: { Authorization: "Bearer " + getChannelAccessToken_() },
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  logBotTiming("fetchLineImageBlob", Date.now() - started, { status: code });
  if (code < 200 || code >= 300) {
    throw new Error("ดึงรูปจาก LINE ไม่สำเร็จ (" + code + ")");
  }

  const blob = response.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > SLIP_IMAGE_MAX_BYTES) {
    throw new Error("รูปใหญ่เกิน 4 MB กรุณาส่งรูปสลิปที่เล็กลง");
  }

  return blob.setName("line-slip-" + messageId + ".jpg");
}

function isThunderRetryableFetchError_(err) {
  const msg = String(err && err.message ? err.message : err);

  return (
    /curl:\s*\((5|6|7|28|35|52|56)\)/i.test(msg) ||
    /timed out|timeout|could not resolve|failed to connect|connection reset|empty reply/i.test(msg)
  );
}


function verifyThunderBankSlipOnce_(imageBlob, messageId) {
  let response;

  try {
    response = UrlFetchApp.fetch(
      THUNDER_VERIFY_BANK_URL,
      {
        method: "post",
        headers: {
          Authorization: "Bearer " + THUNDER_API_KEY
        },
        payload: {
          image: imageBlob,
          checkDuplicate: String(SLIP_VERIFY_CHECK_DUPLICATE),
          matchAccount: String(SLIP_VERIFY_MATCH_ACCOUNT),
          remark: "LINE message " + messageId
        },
        muteHttpExceptions: true,

        // gas-runtime.js ตัวใหม่รองรับ 2 option นี้
        timeoutSeconds: THUNDER_VERIFY_TIMEOUT_SECONDS,
        connectTimeoutSeconds: THUNDER_VERIFY_CONNECT_TIMEOUT_SECONDS
      }
    );
  } catch (err) {
    if (isThunderRetryableFetchError_(err)) {
      const msg = String(err && err.message ? err.message : err);

      if (/curl:\s*\(28\)|timed out|timeout/i.test(msg)) {
        throw new Error("THUNDER_TIMEOUT");
      }

      throw new Error("THUNDER_NETWORK");
    }

    // ไม่ส่งข้อความดิบขึ้นไปเด็ดขาด
    throw new Error("THUNDER_NETWORK");
  }

  const code = Number(response.getResponseCode() || 0);
  const body = response.getContentText();

  if (code === 429 || code >= 500) {
    throw new Error("THUNDER_RETRYABLE_HTTP:" + code);
  }

  let result;

  try {
    result = JSON.parse(body);
  } catch (_) {
    throw new Error("THUNDER_NON_JSON");
  }

  if (code < 200 || code >= 300 || !result.success) {
    const apiError =
      result && result.error
        ? result.error
        : {};

    // เก็บเฉพาะ code ที่ปลอดภัย ไม่แนบ message ดิบจาก upstream
    const safeCode =
      String(apiError.code || "THUNDER_ERROR")
        .replace(/[^A-Z0-9_\-]/gi, "")
        .slice(0, 80) || "THUNDER_ERROR";

    throw new Error(safeCode);
  }

  return result.data;
}


function verifyThunderBankSlip(imageBlob, messageId) {
  const started = Date.now();
  let lastErr = null;

  for (
    let attempt = 1;
    attempt <= THUNDER_VERIFY_MAX_ATTEMPTS;
    attempt++
  ) {
    const attemptStarted = Date.now();

    try {
      const data =
        verifyThunderBankSlipOnce_(
          imageBlob,
          messageId
        );

      console.log(
        "SLIP_THUNDER_ATTEMPT",
        JSON.stringify({
          attempt: attempt,
          ok: true,
          ms: Date.now() - attemptStarted
        })
      );

      logBotTiming(
        "verifyThunderBankSlip",
        Date.now() - started,
        {
          attempts: attempt,
          ok: true
        }
      );

      return data;

    } catch (err) {
      lastErr = err;

      const safeCode =
        getSafeSlipErrorCode_(err);

      console.warn(
        "SLIP_THUNDER_ATTEMPT",
        JSON.stringify({
          attempt: attempt,
          ok: false,
          code: safeCode,
          ms: Date.now() - attemptStarted
        })
      );

      const retryable =
        safeCode === "THUNDER_TIMEOUT" ||
        safeCode === "THUNDER_NETWORK" ||
        safeCode === "THUNDER_RETRYABLE_HTTP" ||
        safeCode === "THUNDER_NON_JSON";

      if (
        !retryable ||
        attempt >= THUNDER_VERIFY_MAX_ATTEMPTS
      ) {
        break;
      }
    }
  }

  logBotTiming(
    "verifyThunderBankSlip",
    Date.now() - started,
    {
      attempts: THUNDER_VERIFY_MAX_ATTEMPTS,
      ok: false,
      code: getSafeSlipErrorCode_(lastErr)
    }
  );

  throw lastErr || new Error("THUNDER_NETWORK");
}


function isExpectedSlipReceiver(data) {
  const raw = data.rawSlip || {};
  const receiver = raw.receiver || {};
  const matchedAccount = data.matchedAccount || null;
  let matchedOk = false;

  if (matchedAccount) {
    matchedOk = isExpectedBank(matchedAccount.bank) &&
      isExpectedAccountName(matchedAccount.nameTh || matchedAccount.nameEn) &&
      isExpectedAccountNumber(matchedAccount.bankNumber);
  }

  const receiverBank = receiver.bank || {};
  const receiverAccount = receiver.account || {};
  const receiverName = receiverAccount.name || {};
  const receiverBankAccount = receiverAccount.bank || receiverAccount.proxy || {};
  const rawOk = isExpectedBank(receiverBank) &&
    isExpectedAccountName(receiverName.th || receiverName.en) &&
    isExpectedAccountNumber(receiverBankAccount.account);

  return matchedOk || rawOk;
}

function isExpectedBank(bank) {
  const text = normalizeSlipText([
    bank && bank.id,
    bank && bank.code,
    bank && bank.short,
    bank && bank.shortCode,
    bank && bank.name,
    bank && bank.nameTh,
    bank && bank.nameEn
  ].filter(x => x).join(" "));
  return text.indexOf("scb") !== -1 ||
    text.indexOf("014") !== -1 ||
    text.indexOf("ไทยพาณิชย์") !== -1;
}

function isExpectedAccountName(name) {
  const actual = normalizeSlipText(name);
  const expected = normalizeSlipText(BANK_ACCOUNT_NAME);
  if (!actual) return false;
  return actual === expected ||
    actual.indexOf(expected) !== -1 ||
    expected.indexOf(actual) !== -1 ||
    (actual.indexOf("ชื่อบัญชี") !== -1 && actual.indexOf("นามสกุลบัญชี") !== -1);
}

function isExpectedAccountNumber(accountNumber) {
  const actual = onlyDigits(accountNumber);
  const expected = onlyDigits(BANK_ACCOUNT_NUMBER);
  if (!actual) return false;
  if (actual === expected) return true;
  if (actual.length >= 6 && expected.indexOf(actual) !== -1) return true;
  if (actual.length >= 4 && expected.slice(-4) === actual.slice(-4)) return true;
  return false;
}

function normalizeSlipText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[.\-\_–—_()]/g, "")
    .trim();
}

function onlyDigits(value) {
  return String(value || "").replace(/\D/g, "");
}

function buildSlipVerifyFlex(data) {
  const raw = data.rawSlip || {};
  const sender = raw.sender || {};
  const receiver = raw.receiver || {};
  const amount = getSlipAmount(data);

  if (data.isDuplicate || data.duplicate) {
    return buildDuplicateSlipFlex(data);
  }

  const green = "#12B886";
  const deepGreen = "#087F5B";
  const softGreen = "#E9FBF5";
  const line = "#E7ECEA";
  const muted = "#8B949E";

  return {
    type: "flex",
    altText: "สลิปถูกต้อง " + formatBaht(amount),
    contents: {
      type: "bubble",
      size: "kilo",
      header: {
        type: "box",
        layout: "horizontal",
        backgroundColor: green,
        paddingAll: "18px",
        alignItems: "center",
        contents: [
          {
            type: "box",
            layout: "vertical",
            width: "46px",
            height: "46px",
            backgroundColor: "#FFFFFF",
            cornerRadius: "23px",
            justifyContent: "center",
            alignItems: "center",
            contents: [
              {
                type: "image",
                url: "https://cdn-icons-png.flaticon.com/512/190/190411.png",
                size: "25px",
                aspectMode: "fit"
              }
            ]
          },
          {
            type: "box",
            layout: "vertical",
            flex: 1,
            margin: "md",
            contents: [
              {
                type: "text",
                text: "ตรวจสอบสลิปสำเร็จ",
                color: "#FFFFFF",
                size: "lg",
                weight: "bold"
              },
              {
                type: "text",
                text: "รายการได้รับการยืนยันแล้ว",
                color: "#DFFFF5",
                size: "xxs",
                margin: "xs"
              }
            ]
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#FFFFFF",
        paddingAll: "18px",
        spacing: "md",
        contents: [
          {
            type: "box",
            layout: "vertical",
            backgroundColor: softGreen,
            cornerRadius: "14px",
            paddingAll: "16px",
            contents: [
              {
                type: "text",
                text: "ยอดโอน",
                color: deepGreen,
                size: "xs",
                weight: "bold",
                align: "center"
              },
              {
                type: "text",
                text: "฿" + formatBahtNumber(amount),
                color: deepGreen,
                size: "3xl",
                weight: "bold",
                align: "center",
                margin: "sm"
              },
              {
                type: "text",
                text: formatSlipDate(raw.date),
                color: muted,
                size: "xxs",
                align: "center",
                margin: "sm"
              }
            ]
          },
          {
            type: "separator",
            color: line
          },
          {
            type: "box",
            layout: "vertical",
            spacing: "md",
            contents: [
              buildPremiumSlipParty("จาก", sender, deepGreen),
              {
                type: "text",
                text: "↓",
                color: "#B8C2BE",
                size: "lg",
                align: "center",
                margin: "none"
              },
              buildPremiumSlipParty("ไปยัง", receiver, deepGreen)
            ]
          },
          {
            type: "box",
            layout: "horizontal",
            backgroundColor: "#F7FAF9",
            cornerRadius: "10px",
            paddingAll: "10px",
            alignItems: "center",
            contents: [
              {
                type: "image",
                url: "https://cdn-icons-png.flaticon.com/512/190/190411.png",
                size: "16px",
                aspectMode: "fit",
                flex: 0
              },
              {
                type: "text",
                text: "ตรวจสอบรายการเรียบร้อย",
                color: deepGreen,
                size: "xxs",
                weight: "bold",
                margin: "sm",
                flex: 1
              }
            ]
          }
        ]
      }
    }
  };
}

function buildPremiumSlipParty(label, party, valueColor) {
  const name = getSlipPartyName(party);
  const bank = getSlipPartyBankText(party);
  const account = getSlipPartyAccount(party);

  return {
    type: "box",
    layout: "vertical",
    contents: [
      {
        type: "text",
        text: label,
        color: "#9AA2A9",
        size: "xxs",
        weight: "bold"
      },
      {
        type: "text",
        text: name,
        color: valueColor,
        size: "md",
        weight: "bold",
        margin: "xs",
        wrap: true
      },
      {
        type: "text",
        text: bank + "  •  " + account,
        color: "#8B949E",
        size: "xxs",
        margin: "xs",
        wrap: true
      }
    ]
  };
}

function buildSlipHero(title, subtitle, accentColor, badgeText) {
  return {
    type: "box",
    layout: "vertical",
    paddingAll: "20px",
    backgroundColor: accentColor,
    contents: [
      {
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "box",
            layout: "vertical",
            width: "44px",
            height: "44px",
            cornerRadius: "22px",
            backgroundColor: "#FFFFFF",
            justifyContent: "center",
            contents: [
              { type: "text", text: accentColor === "#10B981" ? "✓" : "!", color: accentColor, size: "xxl", weight: "bold", align: "center" }
            ]
          },
          {
            type: "box",
            layout: "vertical",
            margin: "md",
            flex: 1,
            contents: [
              { type: "text", text: title, color: "#FFFFFF", weight: "bold", size: "xl", wrap: true },
              { type: "text", text: subtitle, color: "#FFFFFF", size: "xs", margin: "xs", wrap: true }
            ]
          }
        ]
      },
      {
        type: "box",
        layout: "vertical",
        margin: "md",
        paddingAll: "6px",
        cornerRadius: "16px",
        backgroundColor: "#FFFFFF",
        contents: [
          { type: "text", text: badgeText, color: accentColor, size: "xxs", weight: "bold", align: "center" }
        ]
      }
    ]
  };
}

function buildSlipAmountPanel(amount, textColor, backgroundColor) {
  return {
    type: "box",
    layout: "vertical",
    paddingAll: "20px",
    backgroundColor,
    cornerRadius: "12px",
    contents: [
      { type: "text", text: "ยอดโอน", color: textColor, size: "xs", weight: "bold", align: "center" },
      {
        type: "box",
        layout: "baseline",
        margin: "sm",
        justifyContent: "center",
        contents: [
          { type: "text", text: "฿", color: textColor, size: "xl", weight: "bold", flex: 0 },
          { type: "text", text: formatBahtNumber(amount), color: textColor, size: "3xl", weight: "bold", flex: 0 }
        ]
      }
    ]
  };
}

function buildSlipMetaRow(label, value, color) {
  return {
    type: "box",
    layout: "horizontal",
    margin: "md",
    contents: [
      { type: "text", text: label, color: "#6B7280", size: "xs", flex: 0 },
      { type: "text", text: safeText(value, "-"), color, size: "xs", weight: "bold", align: "end", wrap: true }
    ]
  };
}

function buildSlipPartyCard(label, party, color, backgroundColor) {
  return {
    type: "box",
    layout: "horizontal",
    margin: "lg",
    contents: [
      { type: "text", text: label, color: "#6B7280", size: "sm", flex: 0 },
      {
        type: "box",
        layout: "vertical",
        margin: "md",
        paddingAll: "12px",
        backgroundColor,
        cornerRadius: "10px",
        flex: 1,
        contents: [
          { type: "text", text: getSlipPartyName(party), color, size: "sm", weight: "bold", wrap: true },
          { type: "text", text: getSlipPartyAccount(party), color: "#6B7280", size: "xs", margin: "xs", wrap: true },
          { type: "text", text: getSlipPartyBankText(party), color: "#9CA3AF", size: "xxs", margin: "xs", wrap: true }
        ]
      }
    ]
  };
}

function buildExpectedAccountCard(color, backgroundColor) {
  return {
    type: "box",
    layout: "vertical",
    margin: "lg",
    paddingAll: "12px",
    backgroundColor,
    cornerRadius: "10px",
    contents: [
      { type: "text", text: "บัญชีที่กำหนด", color, size: "xs", weight: "bold" },
      { type: "text", text: BANK_ACCOUNT_NAME, color, size: "sm", weight: "bold", margin: "xs", wrap: true },
      { type: "text", text: BANK_ACCOUNT_BANK + " / " + BANK_ACCOUNT_NUMBER, color: "#6B7280", size: "xs", margin: "xs", wrap: true }
    ]
  };
}

function buildSlipBrandFooter(text, color) {
  return {
    type: "box",
    layout: "horizontal",
    margin: "xl",
    contents: [
      { type: "text", text: "THUNDER", color, size: "xs", weight: "bold", flex: 0 },
      { type: "text", text, color: "#6B7280", size: "xxs", margin: "md", wrap: true }
    ]
  };
}

function getSlipPartyName(party) {
  const account = (party || {}).account || {};
  const name = account.name || {};
  return safeText(name.th || name.en, "-");
}

function getSlipPartyAccount(party) {
  const account = (party || {}).account || {};
  const bankAccount = account.bank || account.proxy || {};
  return safeText(bankAccount.account, "-");
}

function getSlipPartyBankText(party) {
  const bank = (party || {}).bank || {};
  return safeText(bank.short || bank.name || bank.nameTh || bank.nameEn, "-");
}

function formatBahtNumber(value) {
  if (typeof value !== "number" || isNaN(value)) return "-";
  return value.toLocaleString("th-TH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function getSlipAmount(data) {
  const raw = data.rawSlip || {};
  if (typeof data.amountInSlip === "number") return data.amountInSlip;
  if (raw.amount && typeof raw.amount.amount === "number") return raw.amount.amount;
  return null;
}

function formatBaht(value) {
  if (typeof value !== "number" || isNaN(value)) return "-";
  return Utilities.formatString("%s บาท", value.toLocaleString("th-TH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }));
}

function formatSlipDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, TZ, "dd/MM/yyyy HH:mm:ss");
}

function safeText(value, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  return String(value);
}

// ================== LINE TOKEN ==================
function getChannelAccessToken_() {
  const fromProperty = PropertiesService
    .getScriptProperties()
    .getProperty("CHANNEL_ACCESS_TOKEN");

  if (fromProperty) return fromProperty;

  // รองรับโปรเจกต์เดิมที่ประกาศ CHANNEL_ACCESS_TOKEN ไว้ในไฟล์ .gs อื่น
  if (typeof CHANNEL_ACCESS_TOKEN !== "undefined" && CHANNEL_ACCESS_TOKEN) {
    return CHANNEL_ACCESS_TOKEN;
  }

  throw new Error("ยังไม่ได้ตั้งค่า CHANNEL_ACCESS_TOKEN");
}

// =====================================================
// ตรวจระบบหลักโดยไม่ต้องรอ LINE ส่งข้อความ
// กด Run: diagnoseBotNow_
// =====================================================
function diagnoseBotNow_() {
  const out = {
    time: Utilities.formatDate(
      new Date(),
      TZ,
      "yyyy-MM-dd HH:mm:ss"
    ),
    webAppUrl: "",
    lineBotInfo: null,
    spreadsheet: null
  };

  try {
    out.webAppUrl =
      String(
        ScriptApp.getService().getUrl() || ""
      );
  } catch (err) {
    out.webAppUrlError =
      String(err && err.message ? err.message : err);
  }

  try {
    const res = UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/info",
      {
        method: "get",
        headers: {
          Authorization:
            "Bearer " +
            getChannelAccessToken_()
        },
        muteHttpExceptions: true
      }
    );

    out.lineBotInfo = {
      status: res.getResponseCode(),
      body: String(
        res.getContentText() || ""
      ).substring(0, 1000)
    };
  } catch (err) {
    out.lineBotInfo = {
      status: 0,
      error:
        String(
          err && err.stack
            ? err.stack
            : err
        )
    };
  }

  try {
    const ss =
      SpreadsheetApp.openById(
        SHEET_ID
      );

    out.spreadsheet = {
      ok: true,
      name: ss.getName()
    };
  } catch (err) {
    out.spreadsheet = {
      ok: false,
      error:
        String(
          err && err.stack
            ? err.stack
            : err
        )
    };
  }

  console.log(
    "BOT_NOW_DIAGNOSTIC " +
    JSON.stringify(
      out,
      null,
      2
    )
  );

  return out;
}


// ================== REPLY ==================
function replyFlex(token,msg){
  const started = Date.now();
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply",{
    method:"post",
    headers:{Authorization:"Bearer "+getChannelAccessToken_()},
    contentType:"application/json",
    payload:JSON.stringify({replyToken:token,messages:[msg]})
  });
  logBotTiming("replyFlex", Date.now() - started, {
    status: response.getResponseCode()
  });
}

function replyText(token,msg){
  const started = Date.now();

  const response = UrlFetchApp.fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "post",
      headers: {
        Authorization: "Bearer " + getChannelAccessToken_()
      },
      contentType: "application/json",
      payload: JSON.stringify({
        replyToken: token,
        messages: [
          {
            type: "text",
            text: String(msg || "")
          }
        ]
      }),
      muteHttpExceptions: true
    }
  );

  const status = Number(
    response.getResponseCode() || 0
  );

  const body = String(
    response.getContentText() || ""
  );

  logBotTiming(
    "replyText",
    Date.now() - started,
    {
      status: status,
      responseBody:
        status >= 200 && status < 300
          ? ""
          : body.substring(0, 1000)
    }
  );

  if (status < 200 || status >= 300) {
    console.error(
      "LINE_REPLY_FAILED " +
      JSON.stringify({
        status: status,
        body: body.substring(0, 1000)
      })
    );
  }

  return {
    ok: status >= 200 && status < 300,
    status: status,
    body: body
  };
}

// ============================================================================
// LIFF API (Netlify hosted page) — boot & order endpoints
// Added by Cowork to support the new minimal LIFF UI on Netlify.
// ============================================================================

function handleLiffBoot_() {
  try {
    const lottoData = loadLotto();
    const lottos = getOpenLottoCandidates(lottoData);
    return jsonResponse_({
      ok: true,
      now: Utilities.formatDate(new Date(), TZ, "HH:mm"),
      lottos: lottos,
      bank: {
        number: BANK_ACCOUNT_NUMBER,
        bank: BANK_ACCOUNT_BANK,
        name: BANK_ACCOUNT_NAME
      }
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) });
  }
}

function handleLiffOrderRequest_(data) {
  try {
    const userId = String((data && data.userId) || '').trim();
    const lottoName = String((data && data.lottoName) || '').trim();
    const rawGroups = (data && data.groups) || [];

    if (!userId) {
      return jsonResponse_({ ok: false, error: 'missing userId (เปิดผ่าน LINE)' });
    }
    if (!lottoName) {
      return jsonResponse_({ ok: false, error: 'ยังไม่ได้เลือกหวย' });
    }

    const lottoData = loadLotto();
    const closeTime = (lottoData && lottoData.map && lottoData.map[lottoName]) || '';
    if (!closeTime) {
      return jsonResponse_({ ok: false, error: 'ไม่พบหวย "' + lottoName + '"' });
    }

    const liffRound = getLiffLottoRoundInfo_(lottoName, closeTime, new Date());
    if (!liffRound.open) {
      return jsonResponse_({ ok: false, error: lottoName + ' ปิดรับแล้ว (' + closeTime + ')' });
    }
    const roundDate = liffRound.roundDate;

    const cleanGroups = [];
    rawGroups.forEach(function(g) {
      if (!g) return;
      const mode = String(g.mode || 'บล');
      const top = Number(g.top || 0) || 0;
      const bottom = Number(g.bottom || 0) || 0;
      const numbers = ((g.numbers) || [])
        .map(function(n) { return String(n).trim(); })
        .filter(function(n) { return n && /^\d+$/.test(n); });
      if (numbers.length === 0) return;
      if (top + bottom <= 0) return;
      const subtotal = numbers.length * (top + bottom);
      cleanGroups.push({ mode: mode, top: top, bottom: bottom, numbers: numbers, subtotal: subtotal });
    });

    if (cleanGroups.length === 0) {
      return jsonResponse_({ ok: false, error: 'ไม่มีรายการที่ถูกต้อง' });
    }

    let totalAmount = 0;
    const uniqueSet = {};
    let uniqueCount = 0;
    cleanGroups.forEach(function(g) {
      totalAmount += g.subtotal;
      g.numbers.forEach(function(n) {
        if (!uniqueSet[n]) { uniqueSet[n] = 1; uniqueCount++; }
      });
    });

    const flex = buildFlex({
      name: lottoName,
      closeTime: closeTime,
      roundDate: roundDate,
      groups: cleanGroups,
      totalDoors: uniqueCount,
      totalAmount: totalAmount
    });

    // บันทึกให้สำเร็จก่อนค่อย Push Flex
    // ถ้าบันทึกไม่ได้ จะเข้า catch ด้านล่างและไม่ออกบิลที่ไม่มีในชีต
    saveParsedBetByUserId_(
      userId,
      lottoName,
      closeTime,
      "LIFF/Netlify",
      cleanGroups,
      totalAmount,
      uniqueCount,
      "",
      roundDate
    );

    linePushFlex_(userId, flex);

    return jsonResponse_({
      ok: true,
      totalAmount: totalAmount,
      totalDoors: uniqueCount,
      groups: cleanGroups.length
    });

  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) });
  }
}

function linePushFlex_(userId, flexMessage) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + getChannelAccessToken_() },
    contentType: 'application/json',
    payload: JSON.stringify({ to: userId, messages: [flexMessage] }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('LINE push failed (' + code + '): ' + res.getContentText());
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ================== SELF TEST : 3 ตัวโต๊ดอย่างเดียว ==================
function testTodOnly20_() {
  const text = [
    "932/930/938/324/308/304/920/928/942/208/204/482/098/084/904/984/230/328/943/384",
    "โต๊ด 1"
  ].join("\n");

  const blocks = splitBlocks(text);
  const result = parseMessage(text, [], {});

  console.log(JSON.stringify({ blocks: blocks, result: result }, null, 2));

  if (!result || !result.success) throw new Error("TEST FAIL: parse ไม่สำเร็จ");
  if (result.totalAmount !== 20) throw new Error("TEST FAIL: ยอดต้อง 20 แต่ได้ " + result.totalAmount);
  if (!result.groups || !result.groups[0] || result.groups[0].mode !== "โต๊ด") {
    throw new Error("TEST FAIL: mode ต้องเป็น โต๊ด");
  }

  return "PASS: โต๊ด 20 ตัว x 1 = 20 บาท";
}

// ============================================================================
// TEST: เลขเบิ้ล 2 ตัว
// รองรับ:
//   เบิ้ล 10
//   เบิ้ล10*10
//   รูดเบิ้ล 10
//   รูดเบิ้ล 10*10
// ทุกแบบต้องได้ 00 11 22 33 44 55 66 77 88 99 และยอด 200 บาท
// ============================================================================
function testDouble2DigitBet_() {
  const cases = [
    "เบิ้ล 10",
    "เบิ้ล10*10",
    "รูดเบิ้ล 10",
    "รูดเบิ้ล 10*10"
  ];

  const expected = ["00","11","22","33","44","55","66","77","88","99"];

  cases.forEach(function(text) {
    const result = parseMessage(text, ["รูด"], {});

    if (!result || !result.success) {
      throw new Error("อ่านไม่ได้: " + text + " / " + (result && result.error));
    }

    if (!result.groups || result.groups.length !== 1) {
      throw new Error("จำนวนกลุ่มผิด: " + text);
    }

    const g = result.groups[0];
    const actual = (g.numbers || []).map(String);

    if (actual.length !== 10 || actual.join(",") !== expected.join(",")) {
      throw new Error("ชุดเลขเบิ้ลผิด: " + text + " => " + actual.join(" "));
    }

    if (Number(g.top) !== 10 || Number(g.bottom) !== 10) {
      throw new Error("ราคาผิด: " + text + " => บน " + g.top + " ล่าง " + g.bottom);
    }

    if (Number(result.totalAmount) !== 200) {
      throw new Error("ยอดผิด: " + text + " => " + result.totalAmount);
    }
  });

  return "PASS: เลขเบิ้ล 10 ตัว / บน10 ล่าง10 / ยอด 200 บาท ทั้ง 4 รูปแบบ";
}



// ============================================================================
// RESULT / WINNER SYSTEM - รวมไว้ในโค้ดหลักไฟล์เดียว
// ============================================================================

// ============================================================================
// LINE COMMAND: สร = สรุปผลหวยรายวัน
// ============================================================================

function isLotteryResultSummaryCommand_(text) {
  return /^สร(?:\s*\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})?\s*$/i
    .test(String(text || "").trim());
}


function parseLotteryResultSummaryDate_(text) {
  const raw = String(text || "").trim();

  const m = raw.match(
    /^สร(?:\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4}))?\s*$/i
  );

  if (!m) {
    throw new Error("รูปแบบคำสั่งไม่ถูกต้อง");
  }

  if (!m[1]) {
    return Utilities.formatDate(
      new Date(),
      TZ,
      "yyyy-MM-dd"
    );
  }

  const day = Number(m[1]);
  const month = Number(m[2]);
  let year = Number(m[3]);

  if (year > 2400) {
    year -= 543;
  }

  const d = new Date(
    year,
    month - 1,
    day,
    12,
    0,
    0
  );

  if (
    isNaN(d.getTime()) ||
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    throw new Error("วันที่ไม่ถูกต้อง");
  }

  return (
    String(year).padStart(4, "0") +
    "-" +
    String(month).padStart(2, "0") +
    "-" +
    String(day).padStart(2, "0")
  );
}


function getLotteryResultSummaryByDate_(dateKey) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("ผลหวย");

  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  // อ่านตรงตามโครงสร้างชีตจริง:
  // B = roundDate
  // C = lottoName
  // D = threeTop
  // E = twoTop
  // F = twoBottom
  //
  // ใช้ getDisplayValues() เพื่อให้ roundDate ออกมาเป็นข้อความ
  // แบบเดียวกับที่เห็นในชีต เช่น 2026-09-04
  const lastRow = sh.getLastRow();
  const rows = sh
    .getRange(2, 2, lastRow - 1, 5)
    .getDisplayValues();

  const wantedDate = String(dateKey || "").trim();
  const matched = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const roundDate = normalizeSummaryRoundDate_(row[0]);
    if (roundDate !== wantedDate) {
      continue;
    }

    const lottoName = String(row[1] || "").trim();
    const threeTop = String(row[2] || "").trim();
    const twoTop = String(row[3] || "").trim();
    const twoBottom = String(row[4] || "").trim();

    if (!lottoName) {
      continue;
    }

    // ต้องมีอย่างน้อย 3 ตัว หรือ 2 ตัวล่าง ถึงถือว่ามีผล
    if (!threeTop && !twoBottom) {
      continue;
    }

    matched.push({
      row: i + 2,
      roundDate: roundDate,
      lottoName: lottoName,
      threeTop: threeTop,
      twoTop: twoTop,
      twoBottom: twoBottom
    });
  }

  console.log(
    "RESULT_SUMMARY_READ " +
    JSON.stringify({
      wantedDate: wantedDate,
      lastRow: lastRow,
      rowsRead: rows.length,
      matched: matched.length,
      firstDates: rows.slice(0, 5).map(function(r) {
        return String(r[0] || "");
      })
    })
  );

  return matched;
}


function normalizeSummaryRoundDate_(value) {
  const s = String(value == null ? "" : value).trim();

  if (!s) {
    return "";
  }

  // yyyy-MM-dd หรือ yyyy-MM-dd HH:mm:ss
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

  if (m) {
    let year = Number(m[1]);

    if (year >= 2400) {
      year -= 543;
    }

    return (
      String(year).padStart(4, "0") +
      "-" +
      String(Number(m[2])).padStart(2, "0") +
      "-" +
      String(Number(m[3])).padStart(2, "0")
    );
  }

  // dd/MM/yyyy, dd-MM-yyyy และปี พ.ศ.
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);

  if (m) {
    let year = Number(m[3]);

    if (year >= 2400) {
      year -= 543;
    }

    return (
      String(year).padStart(4, "0") +
      "-" +
      String(Number(m[2])).padStart(2, "0") +
      "-" +
      String(Number(m[1])).padStart(2, "0")
    );
  }

  return "";
}


function formatSummaryResultNumber_(value, digits) {
  let s = String(value == null ? "" : value).trim();

  if (!s) {
    return "-";
  }

  // กรณี Google Sheet ตัดเลข 0 ด้านหน้า เช่น
  // 12 -> 012 สำหรับ threeTop
  // 4  -> 04  สำหรับ twoBottom
  if (/^\d+$/.test(s)) {
    return s.padStart(Number(digits) || 0, "0");
  }

  return s;
}


function buildLotteryResultSummaryCopyText_(dateKey, results) {
  const displayDate = formatLotterySummaryThaiLongDate_(dateKey);

  const lines = [
    "สรุปผลหวยประจำวันที่ " + displayDate
  ];

  results.forEach(function(item) {
    lines.push(
      formatSummaryResultNumber_(item.threeTop, 3) +
      "-" +
      formatSummaryResultNumber_(item.twoBottom, 2) +
      "   " +
      decorateLotteryNameWithFlag_(item.lottoName)
    );
  });

  return lines.join("\n");
}


function buildLotteryResultSummaryFlex_(dateKey, results) {
  const displayDate = formatLotterySummaryThaiLongDate_(dateKey);
  const copyText = buildLotteryResultSummaryCopyText_(dateKey, results);
  const body = results.map(function(item) {
    return {
      type: "box", layout: "horizontal", margin: "md", spacing: "sm",
      contents: [
        { type: "text", text: formatSummaryResultNumber_(item.threeTop, 3) + "-" + formatSummaryResultNumber_(item.twoBottom, 2),
          size: "md", weight: "bold", color: "#991B1B", flex: 2 },
        { type: "text", text: decorateLotteryNameWithFlag_(item.lottoName),
          size: "sm", color: "#1F2937", wrap: true, flex: 4 }
      ]
    };
  });

  return {
    type: "flex",
    altText: "สรุปผลหวยประจำวันที่ " + displayDate,
    contents: {
      type: "bubble", size: "kilo",
      header: {
        type: "box", layout: "vertical", backgroundColor: "#B91C1C", paddingAll: "18px",
        contents: [
          { type: "text", text: "สรุปผลหวย", size: "xl", weight: "bold", color: "#FFFFFF" },
          { type: "text", text: "ประจำวันที่ " + displayDate, size: "sm", color: "#FEE2E2", margin: "sm" }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "16px",
        contents: body
      },
      footer: {
        type: "box",
        layout: "vertical",
        paddingAll: "12px",
        contents: [
          {
            type: "button",
            style: "primary",
            color: "#B91C1C",
            height: "sm",
            action: {
              type: "clipboard",
              label: "คัดลอกผลทั้งหมด",
              clipboardText: copyText
            }
          }
        ]
      }
    }
  };
}

function formatLotterySummaryThaiLongDate_(dateKey) {
  const p = String(dateKey || "").split("-");
  const months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                  "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  if (p.length !== 3) return String(dateKey || "");
  return Number(p[2]) + " " + months[Number(p[1])-1] + " " + (Number(p[0]) + 543);
}

function decorateLotteryNameWithFlag_(name) {
  const n = String(name || "").trim();
  let f = "";
  if (/ลาว/i.test(n)) f="🇱🇦";
  else if (/ฮานอย|เวียดนาม/i.test(n)) f="🇻🇳";
  else if (/นิคเคอิ|ญี่ปุ่น/i.test(n)) f="🇯🇵";
  else if (/จีน|เซินเจิ้น/i.test(n)) f="🇨🇳";
  else if (/ฮั่งเส็ง|ฮ่องกง/i.test(n)) f="🇭🇰";
  else if (/ไต้หวัน/i.test(n)) f="🇹🇼";
  else if (/เกาหลี/i.test(n)) f="🇰🇷";
  else if (/สิงคโปร์/i.test(n)) f="🇸🇬";
  else if (/อังกฤษ/i.test(n)) f="🇬🇧";
  else if (/เยอรมัน/i.test(n)) f="🇩🇪";
  else if (/รัสเซีย/i.test(n)) f="🇷🇺";
  else if (/ดาวโจนส์|อเมริกา/i.test(n)) f="🇺🇸";
  else if (/ไทย|รัฐบาล/i.test(n)) f="🇹🇭";
  return f ? f + " " + n : n;
}

function replyLotteryResultSummary_(
  replyToken,
  commandText
) {
  const dateKey =
    parseLotteryResultSummaryDate_(
      commandText
    );

  const results =
    getLotteryResultSummaryByDate_(
      dateKey
    );

  if (!results.length) {
    return replyText(
      replyToken,
      "ยังไม่มีผลหวยวันที่ " +
      rsDisplayThaiDate_(
        dateKey
      )
    );
  }

  const flex =
    buildLotteryResultSummaryFlex_(
      dateKey,
      results
    );

  const response =
    UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/message/reply",
      {
        method: "post",
        headers: {
          Authorization:
            "Bearer " +
            getChannelAccessToken_()
        },
        contentType:
          "application/json",
        payload:
          JSON.stringify({
            replyToken:
              replyToken,
            messages: [
              flex
            ]
          }),
        muteHttpExceptions:
          true
      }
    );

  const status =
    Number(
      response.getResponseCode() || 0
    );

  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      "LINE summary reply failed (" +
      status +
      "): " +
      response.getContentText()
    );
  }

  return {
    ok: true,
    status: status,
    date: dateKey,
    count: results.length
  };
}


// ============================================================================
// ResultSystem.gs
// ระบบเก็บโพยจาก LINE + กรอกผล + ตรวจผู้ถูกรางวัล + แจ้งผู้ถูกรางวัล
// ใช้ร่วมกับไฟล์บอทหวยเดิม โดยอาศัย:
//   SHEET_ID
//   TZ
//   loadLotto()
//   getChannelAccessToken_()
// ============================================================================

const RS_BET_SHEET = "โพยหวย";
const RS_RESULT_SHEET = "ผลหวย";
const RS_WINNER_SHEET = "ผู้ถูกรางวัล";
const RS_RESULT_MATCH_SHEET = "ชีต5";

const RS_RATE_3_STRAIGHT = 800;
const RS_RATE_3_TOD = 125;
const RS_RATE_2_TOP = 100;
const RS_RATE_2_BOTTOM = 100;

const RS_BET_HEADERS = [
  "billId",
  "roundDate",
  "betAt",
  "lottoName",
  "closeTime",
  "lineUserId",
  "lineName",
  "sourceType",
  "sourceId",
  "rawText",
  "totalAmount",
  "totalDoors",
  "groupsJson",
  "webhookEventId",
  "status"
];

const RS_RESULT_HEADERS = [
  "resultId",
  "roundDate",
  "lottoName",
  "threeTop",
  "twoTop",
  "twoBottom",
  "checkedAt",
  "totalBills",
  "totalWinners",
  "totalPayout",
  "notifiedAt"
];

const RS_WINNER_HEADERS = [
  "resultId",
  "roundDate",
  "lottoName",
  "lineUserId",
  "lineName",
  "billId",
  "type",
  "number",
  "betAmount",
  "rate",
  "payout"
];


// ============================================================================
// SETUP
// กด Run ฟังก์ชันนี้ 1 ครั้งหลังติดตั้ง
// ============================================================================
function setupLotteryResultSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  rsEnsureSheet_(ss, RS_BET_SHEET, RS_BET_HEADERS);
  rsEnsureSheet_(ss, RS_RESULT_SHEET, RS_RESULT_HEADERS);
  rsEnsureSheet_(ss, RS_WINNER_SHEET, RS_WINNER_HEADERS);

  return "พร้อมใช้งาน: " +
    RS_BET_SHEET + ", " +
    RS_RESULT_SHEET + ", " +
    RS_WINNER_SHEET;
}


function rsEnsureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(
      sh.getMaxColumns(),
      headers.length - sh.getMaxColumns()
    );
  }

  if (sh.getLastRow() === 0 || !String(sh.getRange(1, 1).getValue() || "").trim()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }

  return sh;
}


// ============================================================================
// ADMIN AUTH
// Script Properties:
// RESULT_ADMIN_PIN = รหัสที่ต้องการ เช่น 1234
//
// ถ้าไม่ตั้ง RESULT_ADMIN_PIN ระบบจะไม่บังคับรหัส
// ============================================================================
function rsAssertAdmin_(pin) {
  const expected = String(
    PropertiesService.getScriptProperties()
      .getProperty("RESULT_ADMIN_PIN") || ""
  ).trim();

  if (!expected) return true;

  if (String(pin || "").trim() !== expected) {
    throw new Error("รหัสแอดมินไม่ถูกต้อง");
  }

  return true;
}


// ============================================================================
// หน้า ResultAdmin เรียกฟังก์ชันนี้
// ============================================================================
function getResultAdminBootData(pin) {
  rsAssertAdmin_(pin);
  setupLotteryResultSystem();

  const lottoData = loadLotto();
  const names = Object.keys((lottoData && lottoData.map) || {});

  // รวมชื่อหวยที่เคยมีโพยไว้ด้วย เผื่อหวยนั้นถูกลบจากชีต3ภายหลัง
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const betSh = ss.getSheetByName(RS_BET_SHEET);

  if (betSh && betSh.getLastRow() >= 2) {
    const oldNames = betSh
      .getRange(2, 4, betSh.getLastRow() - 1, 1)
      .getDisplayValues()
      .flat();

    oldNames.forEach(function(name) {
      name = String(name || "").trim();
      if (name && names.indexOf(name) === -1) names.push(name);
    });
  }

  names.sort();

  return {
    ok: true,
    today: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"),
    lottos: names,
    rates: {
      threeStraight: RS_RATE_3_STRAIGHT,
      threeTod: RS_RATE_3_TOD,
      twoTop: RS_RATE_2_TOP,
      twoBottom: RS_RATE_2_BOTTOM
    }
  };
}


// ============================================================================
// SAVE BET FROM LINE WEBHOOK
//
// เรียกหลัง parseMessage() สำเร็จ และก่อน replyFlex()
// ============================================================================
function saveParsedBetFromLine_(
  ev,
  lottoName,
  closeTime,
  rawText,
  groups,
  totalAmount,
  totalDoors
) {
  const src = (ev && ev.source) || {};
  const userId = String(src.userId || "").trim();

  if (!userId) {
    // ไม่มีเจ้าของโพย = ห้ามบันทึกเป็นโพยจริง
    throw new Error("ไม่พบ LINE userId ของเจ้าของโพย");
  }

  const sourceType = String(src.type || "user");
  const sourceId = String(src.groupId || src.roomId || "");

  // การดึงชื่อ LINE เป็นข้อมูลประกอบเท่านั้น
  // ถ้า LINE Profile API มีปัญหา ต้องไม่ทำให้การบันทึกโพยหาย
  let lineName = "";
  try {
    lineName = rsGetLineDisplayName_(userId, sourceType, sourceId);
  } catch (err) {
    lineName = "LINE " + userId.substring(Math.max(0, userId.length - 6));
  }

  const eventId = String(
    (ev && ev.webhookEventId) ||
    (ev && ev.message && ev.message.id) ||
    ""
  );

  return rsSaveBetRecord_({
    lottoName: lottoName,
    closeTime: closeTime,
    userId: userId,
    lineName: lineName,
    sourceType: sourceType,
    sourceId: sourceId,
    rawText: rawText,
    groups: groups,
    totalAmount: totalAmount,
    totalDoors: totalDoors,
    eventId: eventId
  });
}


// ============================================================================
// SAVE BET BY USER ID
// ใช้กรณีหน้า LIFF / API มี userId แต่ไม่มี webhook event
// ============================================================================
function saveParsedBetByUserId_(
  userId,
  lottoName,
  closeTime,
  rawText,
  groups,
  totalAmount,
  totalDoors,
  lineNameFromLiff,
  roundDateFromLiff
) {
  userId = String(userId || "").trim();
  if (!userId) {
    throw new Error("ไม่พบ LINE userId ของเจ้าของโพย");
  }

  let lineName = String(lineNameFromLiff || "").trim();

  if (!lineName) {
    try {
      lineName = rsGetLineDisplayName_(userId, "user", "");
    } catch (err) {
      lineName = "LINE " + userId.substring(Math.max(0, userId.length - 6));
    }
  }

  return rsSaveBetRecord_({
    lottoName: lottoName,
    closeTime: closeTime,
    userId: userId,
    lineName: lineName,
    sourceType: "liff",
    sourceId: "",
    rawText: rawText || "LIFF",
    groups: groups,
    totalAmount: totalAmount,
    totalDoors: totalDoors,
    eventId: "",
    roundDate: String(roundDateFromLiff || "").trim()
  });
}


// ============================================================================
// SAVE BET RECORD — FAST / ROBUST VERSION
//
// แก้ปัญหา "ระบบบันทึกโพยขัดข้อง" ที่เกิดบ่อย:
// 1) ไม่ scan webhookEventId ทั้งชีตทุกครั้ง
// 2) ตรวจ duplicate เฉพาะข้อมูลช่วงท้าย
// 3) ลด Spreadsheet call ตอนอยู่ใน Lock
// 4) ไม่ flush + read-back ทุกโพย
// 5) retry เมื่อ Google Sheets สะดุด
// 6) ก่อน retry ตรวจ billId ป้องกันการเขียนซ้ำ
// ============================================================================
const RS_BET_RECENT_SCAN_ROWS_ = 500;
const RS_BET_VERIFY_SCAN_ROWS_ = 250;
const RS_BET_SAVE_MAX_RETRY_ = 4;
const RS_BET_LOCK_WAIT_MS_ = 1500;

function rsSaveBetRecord_(data) {
  data = data || {};

  const started = Date.now();
  const now = new Date();

  const requestedRoundDate = String(data.roundDate || "").trim();
  const roundDate = /^\d{4}-\d{2}-\d{2}$/.test(requestedRoundDate)
    ? requestedRoundDate
    : rsGetRoundDate_(now, data.closeTime);

  const billId = Utilities.getUuid();
  const cleanGroups = rsNormalizeGroups_(data.groups || []);

  if (cleanGroups.length === 0) {
    throw new Error("ไม่มีรายการแทงสำหรับบันทึก");
  }

  const groupsJson = JSON.stringify(cleanGroups);
  if (groupsJson.length > 49000) {
    throw new Error("ข้อมูลโพยมีขนาดใหญ่เกินไป groupsJson=" + groupsJson.length);
  }

  let rawText = String(data.rawText || "");
  if (rawText.length > 45000) {
    rawText = rawText.substring(0, 45000);
  }

  const eventId = String(data.eventId || "").trim();
  const lottoName = String(data.lottoName || "").trim();

  const row = [
    billId,                              // A billId
    roundDate,                           // B roundDate
    now,                                 // C betAt
    lottoName,                           // D lottoName
    String(data.closeTime || ""),        // E closeTime
    String(data.userId || ""),           // F lineUserId
    String(data.lineName || ""),         // G lineName
    String(data.sourceType || ""),       // H sourceType
    String(data.sourceId || ""),         // I sourceId
    rawText,                              // J rawText
    Number(data.totalAmount || 0),        // K totalAmount
    Number(data.totalDoors || 0),         // L totalDoors
    groupsJson,                           // M groupsJson
    eventId,                              // N webhookEventId
    "ACTIVE"                              // O status
  ];

  let lastError = null;

  // ==========================================================
  // PRIMARY SAVE
  //
  // retry ครอบทั้ง open spreadsheet / get sheet / duplicate
  // check / append ไม่ใช่ retry เฉพาะ appendRow อย่างเดียว
  // ==========================================================
  for (let attempt = 1; attempt <= RS_BET_SAVE_MAX_RETRY_; attempt++) {
    let lock = null;
    let locked = false;

    try {
      // Lock เป็นเพียงตัวช่วยกันชน ไม่ใช่เงื่อนไขว่าต้องได้ Lock
      // หากมี execution อื่นถือ Lock อยู่ ห้ามทำให้โพยลูกค้าตก
      try {
        lock = LockService.getScriptLock();
        locked = lock.tryLock(RS_BET_LOCK_WAIT_MS_);
      } catch (lockErr) {
        locked = false;
        console.warn(
          "BET_SAVE_LOCK_ERROR " +
          JSON.stringify({
            attempt: attempt,
            billId: billId,
            error: String(lockErr && lockErr.message ? lockErr.message : lockErr)
          })
        );
      }

      if (!locked) {
        console.warn(
          "BET_SAVE_LOCK_BYPASS " +
          JSON.stringify({
            attempt: attempt,
            billId: billId,
            lottoName: lottoName
          })
        );
      }

      // เปิด Spreadsheet ใหม่ในทุก retry
      // ป้องกัน service object เก่าค้างหลัง Google Sheets สะดุด
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sh = ss.getSheetByName(RS_BET_SHEET);

      if (!sh) {
        sh = ss.insertSheet(RS_BET_SHEET);
        sh.getRange(1, 1, 1, RS_BET_HEADERS.length)
          .setValues([RS_BET_HEADERS]);
        sh.setFrozenRows(1);
      }

      if (sh.getMaxColumns() < RS_BET_HEADERS.length) {
        sh.insertColumnsAfter(
          sh.getMaxColumns(),
          RS_BET_HEADERS.length - sh.getMaxColumns()
        );
      }

      // ------------------------------------------------------
      // Duplicate check เป็น best-effort
      // ถ้าการอ่าน Sheet สะดุด ต้องไม่ทำให้การบันทึกโพยล้ม
      // LINE redelivery ถูกกันไว้ใน doPost อยู่แล้ว
      // ------------------------------------------------------
      if (eventId) {
        try {
          if (rsBetEventExists_(sh, eventId, lottoName)) {
            console.log(
              "BET_DUPLICATE_SKIP " +
              JSON.stringify({
                eventId: eventId,
                lottoName: lottoName,
                attempt: attempt
              })
            );

            return {
              duplicate: true,
              billId: "",
              roundDate: roundDate,
              lineName: String(data.lineName || "")
            };
          }
        } catch (dupErr) {
          console.warn(
            "BET_DUPLICATE_CHECK_SKIPPED " +
            JSON.stringify({
              eventId: eventId,
              lottoName: lottoName,
              attempt: attempt,
              error: String(dupErr && dupErr.message ? dupErr.message : dupErr)
            })
          );
        }
      }

      // ------------------------------------------------------
      // เขียนจริง
      // appendRow เหมาะกับ concurrent append มากกว่า
      // getLastRow()+1 แล้ว setValues ที่มีโอกาสชนแถวเดียวกัน
      // ------------------------------------------------------
      sh.appendRow(row);

      console.log(
        "BET_SAVED " +
        JSON.stringify({
          billId: billId,
          eventId: eventId,
          lottoName: lottoName,
          userId: String(data.userId || ""),
          totalAmount: Number(data.totalAmount || 0),
          totalDoors: Number(data.totalDoors || 0),
          attempt: attempt,
          lock: locked ? "locked" : "bypass",
          ms: Date.now() - started
        })
      );

      return {
        billId: billId,
        roundDate: roundDate,
        lineName: String(data.lineName || "")
      };

    } catch (err) {
      lastError = err;

      console.error(
        "BET_SAVE_RETRY " +
        JSON.stringify({
          attempt: attempt,
          billId: billId,
          eventId: eventId,
          lottoName: lottoName,
          error: String(err && err.message ? err.message : err)
        })
      );

      // บางครั้ง append สำเร็จ แต่ Spreadsheet service ส่ง error กลับมา
      // เช็ก billId แบบ best-effort ก่อน retry เพื่อไม่เขียนซ้ำ
      try {
        const verifySs = SpreadsheetApp.openById(SHEET_ID);
        const verifySh = verifySs.getSheetByName(RS_BET_SHEET);

        if (verifySh && rsFindBetBillId_(verifySh, billId)) {
          console.log(
            "BET_SAVE_FOUND_AFTER_ERROR " +
            JSON.stringify({
              billId: billId,
              lottoName: lottoName,
              attempt: attempt,
              ms: Date.now() - started
            })
          );

          return {
            billId: billId,
            roundDate: roundDate,
            lineName: String(data.lineName || "")
          };
        }
      } catch (verifyErr) {
        console.warn(
          "BET_VERIFY_SKIPPED " +
          JSON.stringify({
            billId: billId,
            attempt: attempt,
            error: String(
              verifyErr && verifyErr.message
                ? verifyErr.message
                : verifyErr
            )
          })
        );
      }

    } finally {
      if (locked && lock) {
        try {
          lock.releaseLock();
        } catch (releaseErr) {}
      }
    }

    if (attempt < RS_BET_SAVE_MAX_RETRY_) {
      Utilities.sleep(250 * attempt);
    }
  }

  // ==========================================================
  // FALLBACK: GOOGLE SHEETS REST API
  //
  // ถ้า SpreadsheetApp สะดุดทุกครั้ง ใช้ Sheets values.append
  // เป็นทางสำรองอีกชั้นหนึ่ง
  // ==========================================================
  try {
    rsAppendBetViaSheetsApi_(row);

    console.log(
      "BET_SAVED_VIA_SHEETS_API " +
      JSON.stringify({
        billId: billId,
        eventId: eventId,
        lottoName: lottoName,
        totalAmount: Number(data.totalAmount || 0),
        ms: Date.now() - started
      })
    );

    return {
      billId: billId,
      roundDate: roundDate,
      lineName: String(data.lineName || "")
    };

  } catch (apiErr) {
    console.error(
      "BET_SAVE_API_FALLBACK_FAILED " +
      JSON.stringify({
        billId: billId,
        eventId: eventId,
        lottoName: lottoName,
        spreadsheetError: String(
          lastError && lastError.message
            ? lastError.message
            : lastError
        ),
        apiError: String(
          apiErr && apiErr.message
            ? apiErr.message
            : apiErr
        )
      })
    );
  }

  throw new Error(
    "บันทึกโพยไม่สำเร็จ: " +
    String(
      lastError && lastError.message
        ? lastError.message
        : lastError || "ไม่ทราบสาเหตุ"
    )
  );
}


// ============================================================================
// FALLBACK SAVE VIA GOOGLE SHEETS REST API
// ไม่ต้องเปิด Advanced Google Service
// ใช้ OAuth token ของ Apps Script โดยตรง
// ============================================================================
function rsAppendBetViaSheetsApi_(row) {
  const sheetRange = "'" +
    String(RS_BET_SHEET || "โพยหวย").replace(/'/g, "''") +
    "'!A:O";

  const apiRow = (row || []).map(function(value) {
    if (Object.prototype.toString.call(value) === "[object Date]") {
      return Utilities.formatDate(value, TZ, "yyyy-MM-dd HH:mm:ss");
    }
    return value;
  });

  const url =
    "https://sheets.googleapis.com/v4/spreadsheets/" +
    encodeURIComponent(SHEET_ID) +
    "/values/" +
    encodeURIComponent(sheetRange) +
    ":append?valueInputOption=RAW&insertDataOption=INSERT_ROWS";

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    contentType: "application/json",
    payload: JSON.stringify({
      majorDimension: "ROWS",
      values: [apiRow]
    }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(
      "Sheets API HTTP " +
      code +
      ": " +
      response.getContentText().substring(0, 600)
    );
  }

  return true;
}


// ตรวจ billId เฉพาะช่วงท้ายของชีต
// ใช้ตอน retry เพื่อกัน append ซ้ำ
function rsFindBetBillId_(sh, billId) {
  billId = String(billId || "").trim();

  if (!sh || !billId) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const startRow = Math.max(2, lastRow - RS_BET_VERIFY_SCAN_ROWS_ + 1);
  const count = lastRow - startRow + 1;

  if (count <= 0) return false;

  const values = sh
    .getRange(startRow, 1, count, 1)
    .getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || "").trim() === billId) {
      return true;
    }
  }

  return false;
}


function rsNormalizeGroups_(groups) {
  return (groups || [])
    .filter(function(g) {
      return g && Array.isArray(g.numbers) && g.numbers.length > 0;
    })
    .map(function(g) {
      return {
        mode: String(g.mode || ""),
        numbers: (g.numbers || []).map(function(n) {
          return String(n);
        }),
        top: Number(g.top || 0),
        bottom: Number(g.bottom || 0),
        isWin: !!g.isWin,
        winSource: (g.winSource || []).map(function(n) {
          return String(n);
        }),
        betCount:
          typeof g.betCount === "number"
            ? Number(g.betCount)
            : (g.numbers || []).length,
        preserveDuplicates: !!g.preserveDuplicates,
        subtotal:
          typeof g.subtotal === "number"
            ? Number(g.subtotal)
            : (g.numbers || []).length *
              (Number(g.top || 0) + Number(g.bottom || 0))
      };
    });
}


// ตรวจ eventId + lottoName เฉพาะรายการช่วงท้าย
// ไม่ scan คอลัมน์ N ตั้งแต่แถว 2 ถึงท้ายชีตทุกโพยอีกแล้ว
function rsBetEventExists_(sh, eventId, lottoName) {
  eventId = String(eventId || "").trim();
  lottoName = String(lottoName || "").trim();

  if (!sh || !eventId) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const startRow = Math.max(2, lastRow - RS_BET_RECENT_SCAN_ROWS_ + 1);
  const count = lastRow - startRow + 1;

  if (count <= 0) return false;

  // D:N = 11 คอลัมน์
  // D = lottoName, N = webhookEventId
  const values = sh
    .getRange(startRow, 4, count, 11)
    .getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const savedLotto = String(values[i][0] || "").trim();
    const savedEventId = String(values[i][10] || "").trim();

    if (savedEventId !== eventId) continue;

    if (!lottoName) return true;

    if (savedLotto === lottoName) {
      return true;
    }
  }

  return false;
}


// ============================================================================
// LINE DISPLAY NAME
// ============================================================================
function rsGetLineDisplayName_(userId, sourceType, sourceId) {
  userId = String(userId || "").trim();
  if (!userId) return "";

  const cache = CacheService.getScriptCache();
  const cacheKey = (
    "rs_line_name_" +
    sourceType + "_" +
    sourceId + "_" +
    userId
  ).slice(0, 240);

  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const token = getChannelAccessToken_();

  const urls = [];

  if (sourceType === "group" && sourceId) {
    urls.push(
      "https://api.line.me/v2/bot/group/" +
      encodeURIComponent(sourceId) +
      "/member/" +
      encodeURIComponent(userId)
    );
  }

  if (sourceType === "room" && sourceId) {
    urls.push(
      "https://api.line.me/v2/bot/room/" +
      encodeURIComponent(sourceId) +
      "/member/" +
      encodeURIComponent(userId)
    );
  }

  urls.push(
    "https://api.line.me/v2/bot/profile/" +
    encodeURIComponent(userId)
  );

  for (let i = 0; i < urls.length; i++) {
    try {
      const res = UrlFetchApp.fetch(urls[i], {
        method: "get",
        headers: {
          Authorization: "Bearer " + token
        },
        muteHttpExceptions: true
      });

      const code = res.getResponseCode();

      if (code >= 200 && code < 300) {
        const json = JSON.parse(res.getContentText() || "{}");
        const name = String(json.displayName || "").trim();

        if (name) {
          cache.put(cacheKey, name, 6 * 60 * 60);
          return name;
        }
      }
    } catch (err) {
      // ลอง URL ถัดไป
    }
  }

  const fallback =
    "LINE " +
    userId.substring(
      Math.max(0, userId.length - 6)
    );

  cache.put(cacheKey, fallback, 60 * 60);

  return fallback;
}


// ============================================================================
// ROUND DATE
//
// หวยที่ปิดหลังเที่ยงคืนไม่เกิน 02:00
// ถ้ามีโพยเข้าช่วง 00:00 - 02:00 ให้ถือเป็นรอบของ "วันก่อน"
// เพื่อให้โพยก่อนและหลังเที่ยงคืนอยู่รอบเดียวกัน
// ============================================================================
function rsGetRoundDate_(dateObj, closeTime) {
  let d = new Date(dateObj || new Date());

  const nowHour = Number(
    Utilities.formatDate(d, TZ, "H")
  );

  const nowMinute = Number(
    Utilities.formatDate(d, TZ, "m")
  );

  const nowMin = nowHour * 60 + nowMinute;

  const m = String(closeTime || "")
    .match(/^(\d{1,2}):(\d{2})$/);

  if (m) {
    const closeMin =
      Number(m[1]) * 60 +
      Number(m[2]);

    if (
      closeMin <= 120 &&
      nowMin <= 120
    ) {
      d = new Date(
        d.getTime() -
        24 * 60 * 60 * 1000
      );
    }
  }

  return Utilities.formatDate(
    d,
    TZ,
    "yyyy-MM-dd"
  );
}



// ============================================================================
// NORMALIZE DATE FROM SHEET
// รองรับทั้ง String "2026-08-10" และ Date ที่ Google Sheets แปลงให้อัตโนมัติ
// ============================================================================
function rsDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, "yyyy-MM-dd");
  }

  const s = String(value == null ? "" : value).trim();
  if (!s) return "";

  // yyyy-MM-dd
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    return (
      String(m[1]).padStart(4, "0") + "-" +
      String(m[2]).padStart(2, "0") + "-" +
      String(m[3]).padStart(2, "0")
    );
  }

  // dd/MM/yyyy or dd/MM/yyyy พ.ศ.
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    let year = Number(m[3]);
    if (year > 2400) year -= 543;

    return (
      String(year).padStart(4, "0") + "-" +
      String(m[2]).padStart(2, "0") + "-" +
      String(m[1]).padStart(2, "0")
    );
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
  }

  return s;
}



// ============================================================================
// ADMIN: CHECK SAVED BET LIST
// ใช้กับเมนู "เช็ครายการโพย" ใน ResultAdmin.html
// ============================================================================
function getAdminBetList(filter, pin) {
  rsAssertAdmin_(pin);
  setupLotteryResultSystem();

  filter = filter || {};

  const roundDate =
    String(
      filter.date || ""
    ).trim();

  const lottoName =
    String(
      filter.lottoName || ""
    ).trim();

  const keyword =
    String(
      filter.keyword || ""
    )
      .trim()
      .toLowerCase();

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(
      roundDate
    )
  ) {
    throw new Error(
      "วันที่ไม่ถูกต้อง"
    );
  }

  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const sh =
    ss.getSheetByName(
      RS_BET_SHEET
    );

  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return {
      ok: true,
      date: roundDate,
      displayDate:
        rsDisplayThaiDate_(
          roundDate
        ),
      lottoName:
        lottoName,
      truncated: false,
      summary: {
        totalBills: 0,
        totalCustomers: 0,
        totalAmount: 0
      },
      bills: []
    };
  }

  const schema =
    rsGetBetSheetSchema_(
      sh
    );

  const lastRow =
    sh.getLastRow();

  const rows =
    sh.getRange(
      2,
      1,
      lastRow - 1,
      schema.columnCount
    )
    .getValues();

  const matched = [];

  rows.forEach(function(row, index) {
    const bill =
      rsBetRowToObject_(
        row,
        schema,
        index + 2
      );

    if (
      bill.status &&
      bill.status !== "ACTIVE"
    ) {
      return;
    }

    if (
      bill.roundDate !==
      roundDate
    ) {
      return;
    }

    if (
      lottoName &&
      !rsLottoNamesEquivalent_(
        bill.lottoName,
        lottoName
      )
    ) {
      return;
    }

    if (keyword) {
      const haystack = [
        bill.lineName,
        bill.userId,
        bill.lottoName,
        bill.rawText,
        bill.billId
      ]
        .join(" ")
        .toLowerCase();

      if (
        haystack.indexOf(
          keyword
        ) === -1
      ) {
        return;
      }
    }

    matched.push({
      rowNumber:
        bill.rowNumber,
      billId:
        bill.billId,
      date:
        bill.roundDate,
      time:
        bill.betAt instanceof Date
          ? Utilities.formatDate(
              bill.betAt,
              TZ,
              "HH:mm:ss"
            )
          : String(
              bill.betAt || ""
            ),
      lottoName:
        bill.lottoName,
      lineName:
        bill.lineName,
      lineUserId:
        bill.userId,
      rawText:
        bill.rawText,
      totalAmount:
        bill.totalAmount,
      totalDoors:
        bill.totalDoors,
      groups:
        bill.groups
    });
  });

  matched.sort(function(a, b) {
    return (
      Number(b.rowNumber || 0) -
      Number(a.rowNumber || 0)
    );
  });

  const limited =
    matched.slice(
      0,
      300
    );

  const customers = {};

  let totalAmount = 0;

  matched.forEach(function(bill) {
    totalAmount +=
      Number(
        bill.totalAmount || 0
      );

    const key =
      bill.lineUserId ||
      bill.lineName ||
      bill.billId;

    if (key) {
      customers[key] = true;
    }
  });

  return {
    ok: true,
    date:
      roundDate,
    displayDate:
      rsDisplayThaiDate_(
        roundDate
      ),
    lottoName:
      lottoName,
    truncated:
      matched.length >
      limited.length,
    schemaHeaders:
      schema.headers,
    summary: {
      totalBills:
        matched.length,
      totalCustomers:
        Object.keys(
          customers
        ).length,
      totalAmount:
        totalAmount
    },
    bills:
      limited
  };
}


// ============================================================================
// SAVE RESULT + CHECK ALL BETS
// ============================================================================
function saveLotteryResultAndCheck(payload, pin) {
  rsAssertAdmin_(pin);
  setupLotteryResultSystem();

  payload = payload || {};

  const roundDate =
    String(payload.date || "").trim();

  const lottoName =
    String(payload.lottoName || "").trim();

  const threeTop =
    String(payload.threeTop || "")
      .replace(/\D/g, "");

  const twoBottom =
    String(payload.twoBottom || "")
      .replace(/\D/g, "");

  if (!/^\d{4}-\d{2}-\d{2}$/.test(roundDate)) {
    throw new Error("วันที่ไม่ถูกต้อง");
  }

  if (!lottoName) {
    throw new Error("กรุณาเลือกหวย");
  }

  if (!/^\d{3}$/.test(threeTop)) {
    throw new Error("3 ตัวบนต้องมี 3 หลัก");
  }

  if (!/^\d{2}$/.test(twoBottom)) {
    throw new Error("2 ตัวล่างต้องมี 2 หลัก");
  }

  const twoTop = threeTop.slice(-2);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const betSh = ss.getSheetByName(RS_BET_SHEET);
  const resultSh = ss.getSheetByName(RS_RESULT_SHEET);
  const winnerSh = ss.getSheetByName(RS_WINNER_SHEET);

  const existing =
    rsFindResultByDateLotto_(
      resultSh,
      roundDate,
      lottoName
    );

  const resultId =
    existing
      ? String(existing.values[0])
      : Utilities.getUuid();

  const oldThree =
    existing
      ? String(existing.values[3] || "")
      : "";

  const oldBottom =
    existing
      ? String(existing.values[5] || "")
      : "";

  const resultChanged =
    !existing ||
    oldThree !== threeTop ||
    oldBottom !== twoBottom;

  const previousNotifiedAt =
    existing
      ? existing.values[10]
      : "";

  const resultData = {
    resultId: resultId,
    roundDate: roundDate,
    lottoName: lottoName,
    threeTop: threeTop,
    twoTop: twoTop,
    twoBottom: twoBottom
  };

  const bills =
    rsReadBillsForRound_(
      betSh,
      roundDate,
      lottoName
    );

  const rawAwards = [];

  bills.forEach(function(bill) {
    (bill.groups || []).forEach(function(group) {
      rsEvaluateGroup_(
        group,
        resultData,
        bill,
        rawAwards
      );
    });
  });

  const winners =
    rsAggregateWinners_(rawAwards);

  let totalPayout = 0;

  winners.forEach(function(winner) {
    totalPayout += Number(
      winner.totalPayout || 0
    );
  });

  // ลบรายการคนถูกเก่าของผลรอบนี้ก่อน แล้วเขียนใหม่
  rsDeleteWinnerRows_(
    winnerSh,
    resultId
  );

  if (rawAwards.length > 0) {
    const rows = rawAwards.map(function(a) {
      return [
        resultId,
        roundDate,
        lottoName,
        a.userId,
        a.name,
        a.billId,
        a.type,
        a.number,
        a.betAmount,
        a.rate,
        a.payout
      ];
    });

    winnerSh
      .getRange(
        winnerSh.getLastRow() + 1,
        1,
        rows.length,
        RS_WINNER_HEADERS.length
      )
      .setValues(rows);
  }

  const checkedAt = new Date();

  const notifiedAt =
    resultChanged
      ? ""
      : previousNotifiedAt;

  const resultRow = [
    resultId,
    roundDate,
    lottoName,
    threeTop,
    twoTop,
    twoBottom,
    checkedAt,
    bills.length,
    winners.length,
    totalPayout,
    notifiedAt
  ];

  if (existing) {
    resultSh
      .getRange(
        existing.row,
        1,
        1,
        RS_RESULT_HEADERS.length
      )
      .setValues([resultRow]);
  } else {
    resultSh.appendRow(resultRow);
  }

  return {
    ok: true,
    resultId: resultId,
    roundDate: roundDate,
    displayDate: rsDisplayThaiDate_(roundDate),
    lottoName: lottoName,
    threeTop: threeTop,
    twoTop: twoTop,
    twoBottom: twoBottom,
    notified: !!notifiedAt,
    summary: {
      totalBills: bills.length,
      totalWinners: winners.length,
      totalPayout: totalPayout,
      billsWithoutGroups:
        bills.filter(function(b) {
          return !Array.isArray(b.groups) ||
            b.groups.length === 0;
        }).length,
      billsWithoutUserId:
        bills.filter(function(b) {
          return !String(b.userId || "").trim();
        }).length,
      matchedLottoNames:
        Array.from(
          new Set(
            bills.map(function(b) {
              return String(b.lottoName || "");
            }).filter(Boolean)
          )
        )
    },
    winners: winners
  };
}



// ============================================================================
// MATCH RESULT LOTTO NAME WITH SAVED BET LOTTO NAME
//
// ใช้ชีต5เป็นสะพานเชื่อมชื่อผลหวยกับชื่อหวยที่ถูกบันทึกตอนลูกค้าแทง
//
// รูปแบบชีต5ที่รองรับ:
// A = ชื่อที่แสดง/ชื่อจากเว็บ
// B = ชื่อหวยในบอท หรือ alias
//
// ถ้า B มีหลายชื่อ ให้คั่นด้วย | , ; หรือขึ้นบรรทัดใหม่
// ตัวอย่าง:
// A: ฮานอยปกติ
// B: 🇻🇳 ฮานอย | ฮานอย
// ============================================================================
function rsNormalizeLottoForResult_(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/[\u{1F1E6}-\u{1F1FF}]/gu, "")
    .replace(/หวย/g, "")
    .replace(/หุ้น/g, "")
    .replace(/วีไอพี/g, "vip")
    .replace(/\s+/g, "")
    .replace(/[^0-9a-z\u0E00-\u0E7F]/g, "")
    .trim();
}


function rsSplitResultAliases_(text) {
  const raw = String(text || "").trim();

  if (!raw) return [];

  // ถ้ามี delimiter ชัดเจน ให้แยก
  if (/[|,;\n\r]/.test(raw)) {
    return raw
      .split(/[|,;\n\r]+/)
      .map(function(x) {
        return String(x || "").trim();
      })
      .filter(Boolean);
  }

  // ถ้าไม่มี delimiter ให้ถือว่าทั้งเซลล์เป็นชื่อเดียว
  // เพื่อรักษาชื่อหลายคำ เช่น "หุ้นนิคเคอิ VIP เช้า"
  return [raw];
}


function rsGetEquivalentLottoNames_(lottoName) {
  const result = [];
  const seen = {};

  function add_(name) {
    name =
      String(name || "")
        .trim();

    if (!name) return;

    const key =
      rsNormalizeLottoForResult_(
        name
      );

    if (
      !key ||
      seen[key]
    ) {
      return;
    }

    seen[key] = true;
    result.push(name);
  }

  add_(lottoName);

  const targetNorm =
    rsNormalizeLottoForResult_(
      lottoName
    );

  if (!targetNorm) {
    return result;
  }

  try {
    const ss =
      SpreadsheetApp.openById(
        SHEET_ID
      );

    const sh =
      ss.getSheetByName(
        RS_RESULT_MATCH_SHEET
      );

    if (
      !sh ||
      sh.getLastRow() < 1
    ) {
      return result;
    }

    const values =
      sh.getRange(
        1,
        1,
        sh.getLastRow(),
        Math.min(
          3,
          sh.getMaxColumns()
        )
      )
      .getDisplayValues();


    // ======================================================
    // PASS 1:
    // ถ้าชื่อผลตรงคอลัมน์ A ให้ใช้เฉพาะ row นั้น
    // ======================================================
    for (
      let i = 0;
      i < values.length;
      i++
    ) {
      const row =
        values[i];

      const colA =
        String(row[0] || "")
          .trim();

      if (!colA) continue;

      if (
        rsNormalizeLottoForResult_(
          colA
        ) !== targetNorm
      ) {
        continue;
      }

      add_(colA);

      rsSplitResultAliases_(
        row[1] || ""
      ).forEach(add_);

      rsSplitResultAliases_(
        row[2] || ""
      ).forEach(add_);

      // เจอ primary exact แล้วหยุดทันที
      return result;
    }


    // ======================================================
    // PASS 2:
    // ถ้าไม่ตรง A ให้หา exact alias
    // แต่ต้องมีเพียง row เดียวเท่านั้น
    // ======================================================
    const aliasRows = [];

    values.forEach(
      function(row) {
        const colA =
          String(row[0] || "")
            .trim();

        if (!colA) return;

        const aliases = [];

        rsSplitResultAliases_(
          row[1] || ""
        ).forEach(
          function(x) {
            aliases.push(x);
          }
        );

        rsSplitResultAliases_(
          row[2] || ""
        ).forEach(
          function(x) {
            aliases.push(x);
          }
        );

        const matched =
          aliases.some(
            function(alias) {
              return (
                rsNormalizeLottoForResult_(
                  alias
                ) === targetNorm
              );
            }
          );

        if (matched) {
          aliasRows.push(row);
        }
      }
    );


    // Alias ชี้ไปหวยเดียว = ใช้ได้
    if (
      aliasRows.length === 1
    ) {
      const row =
        aliasRows[0];

      add_(row[0]);

      rsSplitResultAliases_(
        row[1] || ""
      ).forEach(add_);

      rsSplitResultAliases_(
        row[2] || ""
      ).forEach(add_);
    }

    // ถ้า alias ซ้ำหลาย row:
    // ไม่รวมชื่อใดเพิ่ม เพื่อกันตรวจโพยผิดหวย

  } catch (err) {
    console.error(
      "อ่านชีต5สำหรับเทียบชื่อหวยไม่สำเร็จ: " +
      err
    );
  }

  return result;
}


function rsLottoNamesEquivalent_(savedName, resultName) {
  const savedNorm =
    rsNormalizeLottoForResult_(savedName);

  const resultNorm =
    rsNormalizeLottoForResult_(resultName);

  if (!savedNorm || !resultNorm) {
    return false;
  }

  // ตรงกันหลังตัดธง/คำว่า หวย/หุ้น/ช่องว่าง
  if (savedNorm === resultNorm) {
    return true;
  }

  const equivalents =
    rsGetEquivalentLottoNames_(resultName);

  for (let i = 0; i < equivalents.length; i++) {
    if (
      savedNorm ===
      rsNormalizeLottoForResult_(
        equivalents[i]
      )
    ) {
      return true;
    }
  }

  return false;
}



// ============================================================================
// ROBUST READER FOR SHEET "โพยหวย"
//
// ไม่บังคับว่าคอลัมน์ต้องอยู่ตำแหน่งเดิม
// ระบบจะอ่านจากชื่อหัวตารางก่อน และ fallback ไปตำแหน่งมาตรฐานเดิม
// ============================================================================
function rsNormalizeHeader_(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\s+/g, "")
    .replace(/[_\-./\\()[\]{}:]+/g, "")
    .trim();
}


function rsFindHeaderIndex_(headers, names, fallbackIndex) {
  const normalizedHeaders =
    (headers || []).map(function(h) {
      return rsNormalizeHeader_(h);
    });

  for (let i = 0; i < names.length; i++) {
    const key =
      rsNormalizeHeader_(
        names[i]
      );

    const idx =
      normalizedHeaders.indexOf(
        key
      );

    if (idx !== -1) {
      return idx;
    }
  }

  return (
    typeof fallbackIndex === "number"
      ? fallbackIndex
      : -1
  );
}


function rsGetBetSheetSchema_(sh) {
  if (!sh) {
    throw new Error(
      'ไม่พบชีต "' +
      RS_BET_SHEET +
      '"'
    );
  }

  const lastColumn =
    Math.max(
      sh.getLastColumn(),
      RS_BET_HEADERS.length
    );

  const headers =
    sh.getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  return {
    columnCount: lastColumn,

    billId:
      rsFindHeaderIndex_(
        headers,
        [
          "billId",
          "bill id",
          "เลขบิล",
          "รหัสบิล",
          "บิล"
        ],
        0
      ),

    roundDate:
      rsFindHeaderIndex_(
        headers,
        [
          "roundDate",
          "round date",
          "วันที่รอบ",
          "รอบวันที่",
          "วันที่",
          "date"
        ],
        1
      ),

    betAt:
      rsFindHeaderIndex_(
        headers,
        [
          "betAt",
          "bet at",
          "เวลาแทง",
          "เวลาส่ง",
          "timestamp",
          "เวลา"
        ],
        2
      ),

    lottoName:
      rsFindHeaderIndex_(
        headers,
        [
          "lottoName",
          "lotto name",
          "ชื่อหวย",
          "หวย",
          "lotto"
        ],
        3
      ),

    closeTime:
      rsFindHeaderIndex_(
        headers,
        [
          "closeTime",
          "close time",
          "เวลาปิด",
          "ปิดรับ"
        ],
        4
      ),

    lineUserId:
      rsFindHeaderIndex_(
        headers,
        [
          "lineUserId",
          "line user id",
          "userId",
          "user id",
          "lineuid",
          "uid",
          "ไลน์userid",
          "ไลน์uid"
        ],
        5
      ),

    lineName:
      rsFindHeaderIndex_(
        headers,
        [
          "lineName",
          "line name",
          "ชื่อLINE",
          "ชื่อไลน์",
          "ชื่อลูกค้า",
          "ลูกค้า",
          "name"
        ],
        6
      ),

    sourceType:
      rsFindHeaderIndex_(
        headers,
        [
          "sourceType",
          "source type"
        ],
        7
      ),

    sourceId:
      rsFindHeaderIndex_(
        headers,
        [
          "sourceId",
          "source id"
        ],
        8
      ),

    rawText:
      rsFindHeaderIndex_(
        headers,
        [
          "rawText",
          "raw text",
          "ข้อความโพย",
          "ข้อความ",
          "โพย",
          "รายการแทง",
          "betText"
        ],
        9
      ),

    totalAmount:
      rsFindHeaderIndex_(
        headers,
        [
          "totalAmount",
          "total amount",
          "ยอดรวม",
          "ยอด",
          "amount"
        ],
        10
      ),

    totalDoors:
      rsFindHeaderIndex_(
        headers,
        [
          "totalDoors",
          "total doors",
          "จำนวนประตู",
          "จำนวนรายการ",
          "ประตู"
        ],
        11
      ),

    groupsJson:
      rsFindHeaderIndex_(
        headers,
        [
          "groupsJson",
          "groups json",
          "groups",
          "รายการjson",
          "json"
        ],
        12
      ),

    webhookEventId:
      rsFindHeaderIndex_(
        headers,
        [
          "webhookEventId",
          "webhook event id",
          "eventId",
          "event id"
        ],
        13
      ),

    status:
      rsFindHeaderIndex_(
        headers,
        [
          "status",
          "สถานะ"
        ],
        14
      ),

    headers: headers
  };
}


function rsCellByIndex_(row, index) {
  if (
    !Array.isArray(row) ||
    typeof index !== "number" ||
    index < 0 ||
    index >= row.length
  ) {
    return "";
  }

  return row[index];
}


// ============================================================================
// ถ้าไม่มี groupsJson แต่มีข้อความโพยเดิม ให้ลอง parse กลับ
// ============================================================================
function rsParseGroupsFromRawText_(rawText) {
  rawText =
    String(rawText || "")
      .trim();

  if (!rawText) {
    return [];
  }

  try {
    const keywords =
      typeof loadKeywords === "function"
        ? loadKeywords()
        : [];

    const doorMap =
      typeof loadDoors === "function"
        ? loadDoors()
        : {};

    const parsed =
      parseMessage(
        rawText,
        keywords,
        doorMap
      );

    if (
      parsed &&
      parsed.success &&
      Array.isArray(parsed.groups)
    ) {
      return parsed.groups;
    }

  } catch (err) {
    console.error(
      "parse โพยจาก rawText ไม่สำเร็จ: " +
      err
    );
  }

  return [];
}


// ============================================================================
// แปลง 1 แถวในชีต -> โครงสร้างโพย
// ============================================================================
function rsBetRowToObject_(
  row,
  schema,
  rowNumber
) {
  schema =
    schema || {};

  let groups = [];

  const groupsRaw =
    rsCellByIndex_(
      row,
      schema.groupsJson
    );

  if (
    groupsRaw !== "" &&
    groupsRaw != null
  ) {
    try {
      const parsedGroups =
        typeof groupsRaw === "string"
          ? JSON.parse(groupsRaw)
          : groupsRaw;

      if (
        Array.isArray(parsedGroups)
      ) {
        groups =
          parsedGroups;
      }
    } catch (err) {
      groups = [];
    }
  }

  const rawText =
    String(
      rsCellByIndex_(
        row,
        schema.rawText
      ) || ""
    );

  if (!groups.length && rawText) {
    groups =
      rsParseGroupsFromRawText_(
        rawText
      );
  }

  const closeTime =
    String(
      rsCellByIndex_(
        row,
        schema.closeTime
      ) || ""
    );

  const betAt =
    rsCellByIndex_(
      row,
      schema.betAt
    );

  let rowDate =
    rsDateKey_(
      rsCellByIndex_(
        row,
        schema.roundDate
      )
    );

  // ถ้าไม่มี roundDate ให้ใช้ betAt แทน
  if (!rowDate && betAt) {
    if (
      betAt instanceof Date &&
      !isNaN(betAt.getTime())
    ) {
      rowDate =
        rsGetRoundDate_(
          betAt,
          closeTime
        );
    } else {
      rowDate =
        rsDateKey_(
          betAt
        );
    }
  }

  return {
    rowNumber:
      rowNumber || 0,

    billId:
      String(
        rsCellByIndex_(
          row,
          schema.billId
        ) || ""
      ),

    roundDate:
      rowDate,

    betAt:
      betAt,

    lottoName:
      String(
        rsCellByIndex_(
          row,
          schema.lottoName
        ) || ""
      ).trim(),

    closeTime:
      closeTime,

    userId:
      String(
        rsCellByIndex_(
          row,
          schema.lineUserId
        ) || ""
      ).trim(),

    lineName:
      String(
        rsCellByIndex_(
          row,
          schema.lineName
        ) || ""
      ).trim(),

    rawText:
      rawText,

    totalAmount:
      Number(
        rsCellByIndex_(
          row,
          schema.totalAmount
        ) || 0
      ),

    totalDoors:
      Number(
        rsCellByIndex_(
          row,
          schema.totalDoors
        ) || 0
      ),

    status:
      String(
        rsCellByIndex_(
          row,
          schema.status
        ) || "ACTIVE"
      ).trim(),

    groups:
      Array.isArray(groups)
        ? groups
        : []
  };
}


// ============================================================================
// DEBUG — ดูว่าระบบอ่านชีต "โพยหวย" เป็นอะไร
//
// Run:
// debugBetSheetReader_("2026-08-10", "ลาว Extra")
//
// หรือ:
// debugBetSheetReader_("", "")
// ============================================================================
function debugBetSheetReader_(
  roundDate,
  lottoName
) {
  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const sh =
    ss.getSheetByName(
      RS_BET_SHEET
    );

  if (!sh) {
    throw new Error(
      'ไม่พบชีต "' +
      RS_BET_SHEET +
      '"'
    );
  }

  const schema =
    rsGetBetSheetSchema_(
      sh
    );

  const lastRow =
    sh.getLastRow();

  const rows =
    lastRow >= 2
      ? sh.getRange(
          2,
          1,
          lastRow - 1,
          schema.columnCount
        ).getValues()
      : [];

  const objects =
    rows.map(
      function(row, i) {
        return rsBetRowToObject_(
          row,
          schema,
          i + 2
        );
      }
    );

  const active =
    objects.filter(function(b) {
      return (
        !b.status ||
        b.status === "ACTIVE"
      );
    });

  const byDate =
    roundDate
      ? active.filter(function(b) {
          return (
            b.roundDate ===
            roundDate
          );
        })
      : active;

  const matched =
    lottoName
      ? byDate.filter(function(b) {
          return rsLottoNamesEquivalent_(
            b.lottoName,
            lottoName
          );
        })
      : byDate;

  const result = {
    sheet:
      RS_BET_SHEET,

    totalRows:
      rows.length,

    activeRows:
      active.length,

    matchedRows:
      matched.length,

    requestedDate:
      roundDate || "",

    requestedLotto:
      lottoName || "",

    schema:
      schema,

    availableDates:
      Array.from(
        new Set(
          active
            .map(function(b) {
              return b.roundDate;
            })
            .filter(Boolean)
        )
      ).slice(-20),

    availableLottos:
      Array.from(
        new Set(
          active
            .map(function(b) {
              return b.lottoName;
            })
            .filter(Boolean)
        )
      ).slice(0, 100),

    matched:
      matched.slice(0, 20)
        .map(function(b) {
          return {
            rowNumber:
              b.rowNumber,
            billId:
              b.billId,
            roundDate:
              b.roundDate,
            lottoName:
              b.lottoName,
            lineName:
              b.lineName,
            userId:
              b.userId,
            totalAmount:
              b.totalAmount,
            groupCount:
              b.groups.length,
            rawText:
              b.rawText
          };
        })
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


// ============================================================================
// READ BILLS
// ============================================================================
function rsReadBillsForRound_(
  sh,
  roundDate,
  lottoName
) {
  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return [];
  }

  const schema =
    rsGetBetSheetSchema_(
      sh
    );

  const lastRow =
    sh.getLastRow();

  const rows =
    sh.getRange(
      2,
      1,
      lastRow - 1,
      schema.columnCount
    )
    .getValues();

  const result = [];

  rows.forEach(function(row, index) {
    const bill =
      rsBetRowToObject_(
        row,
        schema,
        index + 2
      );

    // สถานะว่างถือว่า ACTIVE
    if (
      bill.status &&
      bill.status !== "ACTIVE"
    ) {
      return;
    }

    if (
      bill.roundDate !==
      roundDate
    ) {
      return;
    }

    if (
      !rsLottoNamesEquivalent_(
        bill.lottoName,
        lottoName
      )
    ) {
      return;
    }

    // เจอโพยแม้ groups จะว่าง
    // เพื่อให้ summary totalBills บอกได้ว่าเจอแถวแล้ว
    result.push({
      rowNumber:
        bill.rowNumber,
      billId:
        bill.billId,
      roundDate:
        bill.roundDate,
      betAt:
        bill.betAt,
      lottoName:
        bill.lottoName,
      closeTime:
        bill.closeTime,
      userId:
        bill.userId,
      lineName:
        bill.lineName,
      rawText:
        bill.rawText,
      totalAmount:
        bill.totalAmount,
      totalDoors:
        bill.totalDoors,
      groups:
        bill.groups
    });
  });

  console.log(
    "RESULT_BET_SCAN date=" +
    roundDate +
    " lotto=" +
    lottoName +
    " rows=" +
    rows.length +
    " matched=" +
    result.length
  );

  return result;
}


// ============================================================================
// CHECK WINNERS
// ============================================================================
function rsEvaluateGroup_(
  g,
  result,
  bill,
  out
) {
  if (!g) return;

  const mode =
    String(g.mode || "");

  const numbers =
    (g.numbers || [])
      .map(function(n) {
        return String(n);
      });

  const top =
    Number(g.top || 0);

  const bottom =
    Number(g.bottom || 0);


  // --------------------------------------------------------------------------
  // 3 ตัว
  // --------------------------------------------------------------------------
  numbers.forEach(function(number) {
    if (!/^\d{3}$/.test(number)) {
      return;
    }

    // โต๊ดอย่างเดียว
    if (mode === "โต๊ด") {
      if (
        top > 0 &&
        rsSameTod3_(
          number,
          result.threeTop
        )
      ) {
        rsAddAward_(
          out,
          bill,
          "3 ตัวโต๊ด",
          number,
          top,
          RS_RATE_3_TOD
        );
      }

      return;
    }


    // ตรง + โต๊ด
    if (mode === "ตท") {
      if (
        top > 0 &&
        number === result.threeTop
      ) {
        rsAddAward_(
          out,
          bill,
          "3 ตัวตรง",
          number,
          top,
          RS_RATE_3_STRAIGHT
        );
      }

      if (
        bottom > 0 &&
        rsSameTod3_(
          number,
          result.threeTop
        )
      ) {
        rsAddAward_(
          out,
          bill,
          "3 ตัวโต๊ด",
          number,
          bottom,
          RS_RATE_3_TOD
        );
      }

      return;
    }


    // ตรง / 6กลับ / 3กลับ / ตอง / หาม / เบิ้ลหน้า / เบิ้ลหลัง
    // parser สร้างเลขที่จะซื้อจริงไว้ใน numbers แล้ว
    if (
      top > 0 &&
      number === result.threeTop
    ) {
      rsAddAward_(
        out,
        bill,
        "3 ตัวตรง",
        number,
        top,
        RS_RATE_3_STRAIGHT
      );
    }
  });


  // --------------------------------------------------------------------------
  // 2 ตัว
  // --------------------------------------------------------------------------
  numbers.forEach(function(number) {
    if (!/^\d{2}$/.test(number)) {
      return;
    }

    if (
      top > 0 &&
      number === result.twoTop
    ) {
      rsAddAward_(
        out,
        bill,
        "2 ตัวบน",
        number,
        top,
        RS_RATE_2_TOP
      );
    }

    if (
      bottom > 0 &&
      number === result.twoBottom
    ) {
      rsAddAward_(
        out,
        bill,
        "2 ตัวล่าง",
        number,
        bottom,
        RS_RATE_2_BOTTOM
      );
    }
  });
}


function rsSameTod3_(a, b) {
  return String(a || "")
    .split("")
    .sort()
    .join("") ===
    String(b || "")
      .split("")
      .sort()
      .join("");
}


function rsAddAward_(
  out,
  bill,
  type,
  number,
  betAmount,
  rate
) {
  betAmount = Number(betAmount || 0);
  rate = Number(rate || 0);

  if (betAmount <= 0 || rate <= 0) {
    return;
  }

  out.push({
    userId: String(bill.userId || ""),
    name: String(
      bill.lineName ||
      bill.userId ||
      "ไม่ทราบชื่อ"
    ),
    billId: String(bill.billId || ""),
    type: type,
    number: String(number || ""),
    betAmount: betAmount,
    rate: rate,
    payout: betAmount * rate
  });
}


// ============================================================================
// AGGREGATE WINNERS
// รวมรายการของลูกค้าคนเดียวกัน
// ============================================================================
function rsAggregateWinners_(rawAwards) {
  const users = {};

  (rawAwards || []).forEach(function(a) {
    const userKey =
      a.userId
        ? "U:" + a.userId
        : "N:" + a.name;

    if (!users[userKey]) {
      users[userKey] = {
        userId: a.userId,
        name: a.name,
        totalPayout: 0,
        itemsMap: {}
      };
    }

    const itemKey = [
      a.type,
      a.number,
      a.rate
    ].join("|");

    if (!users[userKey].itemsMap[itemKey]) {
      users[userKey].itemsMap[itemKey] = {
        type: a.type,
        number: a.number,
        betAmount: 0,
        rate: a.rate,
        payout: 0
      };
    }

    const item =
      users[userKey].itemsMap[itemKey];

    item.betAmount +=
      Number(a.betAmount || 0);

    item.payout +=
      Number(a.payout || 0);

    users[userKey].totalPayout +=
      Number(a.payout || 0);
  });


  return Object.keys(users)
    .map(function(key) {
      const u = users[key];

      return {
        userId: u.userId,
        name: u.name,
        totalPayout: u.totalPayout,
        items: Object.keys(u.itemsMap)
          .map(function(itemKey) {
            return u.itemsMap[itemKey];
          })
      };
    })
    .sort(function(a, b) {
      return b.totalPayout - a.totalPayout;
    });
}


// ============================================================================
// RESULT SHEET HELPERS
// ============================================================================
function rsFindResultByDateLotto_(
  sh,
  roundDate,
  lottoName
) {
  if (!sh || sh.getLastRow() < 2) {
    return null;
  }

  const rows = sh
    .getRange(
      2,
      1,
      sh.getLastRow() - 1,
      RS_RESULT_HEADERS.length
    )
    .getValues();

  for (let i = 0; i < rows.length; i++) {
    if (
      rsDateKey_(rows[i][1]) === roundDate &&
      String(rows[i][2] || "") === lottoName
    ) {
      return {
        row: i + 2,
        values: rows[i]
      };
    }
  }

  return null;
}


function rsFindResultById_(
  sh,
  resultId
) {
  if (!sh || sh.getLastRow() < 2) {
    return null;
  }

  const rows = sh
    .getRange(
      2,
      1,
      sh.getLastRow() - 1,
      RS_RESULT_HEADERS.length
    )
    .getValues();

  for (let i = 0; i < rows.length; i++) {
    if (
      String(rows[i][0] || "") ===
      String(resultId || "")
    ) {
      return {
        row: i + 2,
        values: rows[i]
      };
    }
  }

  return null;
}


function rsDeleteWinnerRows_(
  sh,
  resultId
) {
  if (!sh || sh.getLastRow() < 2) {
    return;
  }

  const ids = sh
    .getRange(
      2,
      1,
      sh.getLastRow() - 1,
      1
    )
    .getDisplayValues();

  for (let i = ids.length - 1; i >= 0; i--) {
    if (
      String(ids[i][0] || "") ===
      String(resultId || "")
    ) {
      sh.deleteRow(i + 2);
    }
  }
}




// ============================================================================
// WINNER NOTIFICATION SAFETY
//
// กฎ:
// 1) ผลย้อนหลังห้ามส่ง
// 2) resultId ต้องถูก mark ว่าเป็น "ผลใหม่" จากระบบ AUTO ก่อน
// 3) userId เดิม + resultId เดิม ส่งสำเร็จได้เพียงครั้งเดียว
// ============================================================================

const RS_NOTIFY_ELIGIBLE_KEY_ =
  "RS_NOTIFY_ELIGIBLE_RESULTS_V3";

const RS_NOTIFY_LEDGER_KEY_ =
  "RS_WINNER_NOTIFY_LEDGER_V3";


function rsLoadJsonProperty_(key) {
  try {
    const raw =
      PropertiesService
        .getScriptProperties()
        .getProperty(key);

    if (!raw) return {};

    const parsed =
      JSON.parse(raw);

    return (
      parsed &&
      typeof parsed === "object"
    )
      ? parsed
      : {};

  } catch (err) {
    return {};
  }
}


function rsSaveJsonProperty_(key, value) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      key,
      JSON.stringify(
        value || {}
      )
    );
}


function rsPruneTimestampMap_(
  map,
  maxItems
) {
  map = map || {};

  const entries =
    Object.keys(map)
      .map(function(key) {
        const item =
          map[key];

        const ts =
          typeof item === "number"
            ? item
            : Number(
                item &&
                item.ts
                  ? item.ts
                  : 0
              );

        return {
          key: key,
          ts: ts
        };
      })
      .sort(function(a, b) {
        return b.ts - a.ts;
      });

  const keep =
    entries.slice(
      0,
      maxItems || 1000
    );

  const result = {};

  keep.forEach(function(item) {
    result[item.key] =
      map[item.key];
  });

  return result;
}


function rsMarkResultNotificationEligible_(
  resultId
) {
  resultId =
    String(resultId || "")
      .trim();

  if (!resultId) return;

  let map =
    rsLoadJsonProperty_(
      RS_NOTIFY_ELIGIBLE_KEY_
    );

  map[resultId] = {
    ts: Date.now()
  };

  map =
    rsPruneTimestampMap_(
      map,
      500
    );

  rsSaveJsonProperty_(
    RS_NOTIFY_ELIGIBLE_KEY_,
    map
  );
}


function rsIsResultNotificationEligible_(
  resultId
) {
  resultId =
    String(resultId || "")
      .trim();

  if (!resultId) {
    return false;
  }

  const map =
    rsLoadJsonProperty_(
      RS_NOTIFY_ELIGIBLE_KEY_
    );

  return !!map[resultId];
}


function rsWinnerNotifyKey_(
  resultId,
  userId
) {
  return (
    String(resultId || "") +
    "|" +
    String(userId || "")
  );
}


function rsWasWinnerAlreadyNotified_(
  resultId,
  userId
) {
  const key =
    rsWinnerNotifyKey_(
      resultId,
      userId
    );

  const ledger =
    rsLoadJsonProperty_(
      RS_NOTIFY_LEDGER_KEY_
    );

  return !!ledger[key];
}


function rsMarkWinnerNotified_(
  resultId,
  userId
) {
  const key =
    rsWinnerNotifyKey_(
      resultId,
      userId
    );

  let ledger =
    rsLoadJsonProperty_(
      RS_NOTIFY_LEDGER_KEY_
    );

  ledger[key] = {
    ts: Date.now()
  };

  ledger =
    rsPruneTimestampMap_(
      ledger,
      1500
    );

  rsSaveJsonProperty_(
    RS_NOTIFY_LEDGER_KEY_,
    ledger
  );
}


// ============================================================================
// NOTIFY ONE WINNER
// ใช้ปุ่มแจ้งถูกเฉพาะลูกค้าคนนั้น
// ============================================================================
function notifySingleLotteryWinner(
  resultId,
  userId,
  pin
) {
  rsAssertAdmin_(pin);
  setupLotteryResultSystem();

  resultId =
    String(resultId || "")
      .trim();

  userId =
    String(userId || "")
      .trim();

  if (!resultId) {
    throw new Error(
      "ไม่พบ resultId"
    );
  }

  if (!userId) {
    throw new Error(
      "ไม่พบ LINE userId ของลูกค้า"
    );
  }

  // ผลย้อนหลังห้ามส่ง
  if (
    !rsIsResultNotificationEligible_(
      resultId
    )
  ) {
    return {
      ok: true,
      blockedHistorical: true,
      sent: 0,
      message:
        "ผลนี้เป็นผลย้อนหลัง ระบบไม่ส่งแจ้งลูกค้า"
    };
  }

  // คนนี้เคยส่งแล้ว ห้ามส่งซ้ำ
  if (
    rsWasWinnerAlreadyNotified_(
      resultId,
      userId
    )
  ) {
    return {
      ok: true,
      alreadyNotified: true,
      sent: 0,
      message:
        "ลูกค้าคนนี้เคยได้รับแจ้งผลนี้แล้ว"
    };
  }

  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const resultSh =
    ss.getSheetByName(
      RS_RESULT_SHEET
    );

  const winnerSh =
    ss.getSheetByName(
      RS_WINNER_SHEET
    );

  const resultRow =
    rsFindResultById_(
      resultSh,
      resultId
    );

  if (!resultRow) {
    throw new Error(
      "ไม่พบผลหวยรายการนี้"
    );
  }

  // ถ้ารอบนี้เคยถูก mark แจ้งแล้วจากระบบเก่า
  // ถือว่าแจ้งไปแล้ว ห้ามยิงซ้ำ
  if (resultRow.values[10]) {
    return {
      ok: true,
      alreadyNotified: true,
      sent: 0,
      message:
        "ผลรอบนี้เคยแจ้งคนถูกแล้ว"
    };
  }

  const rows =
    winnerSh.getLastRow() >= 2
      ? winnerSh
          .getRange(
            2,
            1,
            winnerSh.getLastRow() - 1,
            RS_WINNER_HEADERS.length
          )
          .getValues()
          .filter(function(r) {
            return (
              String(r[0] || "") ===
                resultId &&
              String(r[3] || "") ===
                userId
            );
          })
      : [];

  if (!rows.length) {
    throw new Error(
      "ไม่พบรายการถูกรางวัลของลูกค้าคนนี้"
    );
  }

  const rawAwards =
    rows.map(function(r) {
      return {
        userId:
          String(r[3] || ""),
        name:
          String(r[4] || ""),
        billId:
          String(r[5] || ""),
        type:
          String(r[6] || ""),
        number:
          String(r[7] || ""),
        betAmount:
          Number(r[8] || 0),
        rate:
          Number(r[9] || 0),
        payout:
          Number(r[10] || 0)
      };
    });

  const winners =
    rsAggregateWinners_(
      rawAwards
    );

  if (!winners.length) {
    throw new Error(
      "ไม่พบผู้ถูกรางวัล"
    );
  }

  const winner =
    winners[0];

  const flex =
    rsBuildWinnerFlex_({
      lottoName:
        String(
          resultRow.values[2] || ""
        ),

      roundDate:
        rsDateKey_(
          resultRow.values[1]
        ),

      winner:
        winner
    });

  rsPushFlex_(
    userId,
    flex
  );

  // บันทึกทันทีหลัง LINE ส่งสำเร็จ
  // ป้องกันการส่งซ้ำถ้าระบบรันอีก
  rsMarkWinnerNotified_(
    resultId,
    userId
  );

  return {
    ok: true,
    sent: 1,
    name:
      winner.name,
    totalPayout:
      winner.totalPayout,
    message:
      "แจ้งถูกให้ " +
      winner.name +
      " แล้ว"
  };
}


// ============================================================================
// NOTIFY WINNERS THROUGH LINE
// ============================================================================
function notifyLotteryWinners(
  resultId,
  pin
) {
  rsAssertAdmin_(pin);
  setupLotteryResultSystem();

  resultId =
    String(resultId || "")
      .trim();

  if (!resultId) {
    throw new Error(
      "ไม่พบ resultId"
    );
  }

  // ========================================================
  // ห้ามแจ้งผลย้อนหลัง
  // resultId ต้องถูก mark จาก AUTO ว่าเป็นผลที่เพิ่งออกหลัง baseline
  // ========================================================
  if (
    !rsIsResultNotificationEligible_(
      resultId
    )
  ) {
    return {
      ok: true,
      blockedHistorical: true,
      sent: 0,
      skipped: 0,
      duplicateSkipped: 0,
      message:
        "ผลนี้เป็นผลย้อนหลัง ระบบไม่ส่งแจ้งลูกค้า"
    };
  }

  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const resultSh =
    ss.getSheetByName(
      RS_RESULT_SHEET
    );

  const winnerSh =
    ss.getSheetByName(
      RS_WINNER_SHEET
    );

  const resultRow =
    rsFindResultById_(
      resultSh,
      resultId
    );

  if (!resultRow) {
    throw new Error(
      "ไม่พบผลหวยรายการนี้"
    );
  }

  // ระบบเดิมเคยแจ้งรอบนี้แล้ว
  if (resultRow.values[10]) {
    return {
      ok: true,
      alreadyNotified: true,
      sent: 0,
      skipped: 0,
      duplicateSkipped: 0,
      message:
        "ผลรอบนี้เคยแจ้งคนถูกแล้ว"
    };
  }

  const roundDate =
    String(
      resultRow.values[1] || ""
    );

  const lottoName =
    String(
      resultRow.values[2] || ""
    );

  let rows = [];

  if (
    winnerSh.getLastRow() >= 2
  ) {
    rows =
      winnerSh
        .getRange(
          2,
          1,
          winnerSh.getLastRow() - 1,
          RS_WINNER_HEADERS.length
        )
        .getValues()
        .filter(function(r) {
          return (
            String(r[0] || "") ===
            resultId
          );
        });
  }

  if (!rows.length) {
    return {
      ok: true,
      sent: 0,
      skipped: 0,
      duplicateSkipped: 0,
      message:
        "ไม่มีผู้ถูกรางวัล"
    };
  }

  const rawAwards =
    rows.map(function(r) {
      return {
        userId:
          String(r[3] || ""),
        name:
          String(r[4] || ""),
        billId:
          String(r[5] || ""),
        type:
          String(r[6] || ""),
        number:
          String(r[7] || ""),
        betAmount:
          Number(r[8] || 0),
        rate:
          Number(r[9] || 0),
        payout:
          Number(r[10] || 0)
      };
    });

  const winners =
    rsAggregateWinners_(
      rawAwards
    );

  let sent = 0;
  let skipped = 0;
  let duplicateSkipped = 0;
  let sendable = 0;

  winners.forEach(function(winner) {
    if (!winner.userId) {
      skipped++;
      return;
    }

    sendable++;

    // เคยส่งให้ userId นี้ใน resultId นี้แล้ว
    if (
      rsWasWinnerAlreadyNotified_(
        resultId,
        winner.userId
      )
    ) {
      duplicateSkipped++;
      return;
    }

    const flex =
      rsBuildWinnerFlex_({
        lottoName:
          lottoName,
        roundDate:
          roundDate,
        winner:
          winner
      });

    try {
      rsPushFlex_(
        winner.userId,
        flex
      );

      // บันทึกทันทีหลัง push สำเร็จ
      // ถ้าคนถัดไป error รอบหน้าไม่ยิงคนนี้ซ้ำ
      rsMarkWinnerNotified_(
        resultId,
        winner.userId
      );

      sent++;

    } catch (err) {
      console.error(
        "แจ้งคนถูกไม่สำเร็จ " +
        winner.name +
        ": " +
        err
      );

      skipped++;
    }
  });

  // เช็กว่าทุก user ที่ส่งได้ถูกบันทึกใน ledger แล้วหรือยัง
  const allSendableDone =
    winners
      .filter(function(w) {
        return !!w.userId;
      })
      .every(function(w) {
        return rsWasWinnerAlreadyNotified_(
          resultId,
          w.userId
        );
      });

  // mark notifiedAt เฉพาะเมื่อส่งครบทุกคนที่มี userId
  if (
    sendable > 0 &&
    allSendableDone
  ) {
    resultSh
      .getRange(
        resultRow.row,
        11
      )
      .setValue(
        new Date()
      );
  }

  return {
    ok: true,
    sent:
      sent,
    skipped:
      skipped,
    duplicateSkipped:
      duplicateSkipped,
    complete:
      allSendableDone,
    message:
      sent > 0
        ? (
            "แจ้งผู้ถูกรางวัลใหม่แล้ว " +
            sent +
            " คน" +
            (
              duplicateSkipped
                ? " • ข้ามที่เคยแจ้ง " +
                  duplicateSkipped +
                  " คน"
                : ""
            )
          )
        : (
            duplicateSkipped > 0
              ? "ไม่มีการส่งซ้ำ ทุกคนเคยได้รับแจ้งแล้ว"
              : "ยังไม่มีผู้รับที่ส่งสำเร็จ"
          )
  };
}


function rsPushFlex_(
  userId,
  flex
) {
  const res =
    UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/message/push",
      {
        method: "post",
        headers: {
          Authorization:
            "Bearer " +
            getChannelAccessToken_()
        },
        contentType: "application/json",
        payload: JSON.stringify({
          to: userId,
          messages: [flex]
        }),
        muteHttpExceptions: true
      }
    );

  const code =
    res.getResponseCode();

  if (
    code < 200 ||
    code >= 300
  ) {
    throw new Error(
      "LINE push failed (" +
      code +
      "): " +
      res.getContentText()
    );
  }
}


function rsBuildWinnerFlex_(data) {
  const winner = data.winner;
  const items = winner.items || [];

  const contents = [];

  items.slice(0, 20).forEach(function(item) {
    contents.push({
      type: "box",
      layout: "horizontal",
      margin: "md",
      contents: [
        {
          type: "box",
          layout: "vertical",
          flex: 1,
          contents: [
            {
              type: "text",
              text:
                item.type +
                "  " +
                item.number,
              size: "sm",
              weight: "bold",
              color: "#111827"
            },
            {
              type: "text",
              text:
                rsMoney_(item.betAmount) +
                " × " +
                rsMoney_(item.rate),
              size: "xs",
              color: "#6B7280",
              margin: "xs"
            }
          ]
        },
        {
          type: "text",
          text:
            rsMoney_(item.payout) +
            " บาท",
          size: "sm",
          weight: "bold",
          color: "#16A34A",
          align: "end",
          flex: 0
        }
      ]
    });
  });

  if (items.length > 20) {
    contents.push({
      type: "text",
      text:
        "และอีก " +
        (items.length - 20) +
        " รายการ",
      size: "xs",
      color: "#6B7280",
      margin: "md"
    });
  }

  contents.push({
    type: "separator",
    margin: "lg",
    color: "#D1FAE5"
  });

  contents.push({
    type: "box",
    layout: "horizontal",
    margin: "lg",
    contents: [
      {
        type: "text",
        text: "รวมถูกรางวัล",
        size: "md",
        weight: "bold",
        color: "#111827",
        flex: 1
      },
      {
        type: "text",
        text:
          rsMoney_(
            winner.totalPayout
          ) +
          " บาท",
        size: "lg",
        weight: "bold",
        color: "#15803D",
        align: "end",
        flex: 0
      }
    ]
  });

  return {
    type: "flex",
    altText:
      "🎉 ถูกรางวัล " +
      data.lottoName +
      " " +
      rsMoney_(
        winner.totalPayout
      ) +
      " บาท",
    contents: {
      type: "bubble",
      size: "kilo",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#16A34A",
        paddingAll: "18px",
        contents: [
          {
            type: "text",
            text: "🎉 ถูกรางวัล",
            size: "xl",
            weight: "bold",
            color: "#FFFFFF"
          },
          {
            type: "text",
            text: data.lottoName,
            size: "sm",
            color: "#DCFCE7",
            margin: "sm",
            wrap: true
          },
          {
            type: "text",
            text:
              "วันที่ " +
              rsDisplayThaiDate_(
                data.roundDate
              ),
            size: "xs",
            color: "#BBF7D0",
            margin: "xs"
          }
        ]
      },
      body: {
        type: "box",
        layout: "vertical",
        paddingAll: "18px",
        contents: contents
      }
    }
  };
}


// ============================================================================
// FORMAT
// ============================================================================
function rsMoney_(value) {
  value = Number(value || 0);

  return String(
    Math.round(
      value * 100
    ) / 100
  ).replace(
    /\B(?=(\d{3})+(?!\d))/g,
    ","
  );
}


function rsDisplayThaiDate_(
  yyyyMmDd
) {
  const m =
    String(yyyyMmDd || "")
      .match(
        /^(\d{4})-(\d{2})-(\d{2})$/
      );

  if (!m) {
    return String(
      yyyyMmDd || ""
    );
  }

  return (
    m[3] +
    "/" +
    m[2] +
    "/" +
    (Number(m[1]) + 543)
  );
}



// ============================================================================
// DIAGNOSTIC
// กด Run ฟังก์ชันนี้เพื่อเช็กว่าชีตโพยมีข้อมูลหรือไม่
// ============================================================================
function testSavedBetData_() {
  setupLotteryResultSystem();

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(RS_BET_SHEET);

  if (!sh) {
    throw new Error("ไม่พบชีต " + RS_BET_SHEET);
  }

  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    console.log("ชีตโพยหวยยังไม่มีข้อมูลโพย มีเฉพาะหัวตาราง");
    return {
      ok: true,
      rows: 0,
      message: "ยังไม่มีโพยถูกบันทึก"
    };
  }

  const last = sh
    .getRange(
      Math.max(2, lastRow - 4),
      1,
      Math.min(5, lastRow - 1),
      RS_BET_HEADERS.length
    )
    .getDisplayValues();

  console.log(JSON.stringify(last));

  return {
    ok: true,
    rows: lastRow - 1,
    latest: last
  };
}


// ============================================================================
// KeywordReply.gs — AUTO REPLY จากชีต4
// คอลัมน์ A = คีย์เวิร์ด
// คอลัมน์ B = ข้อความตอบกลับ
// ============================================================================
function tryReplyKeywordFromSheet4_(replyToken, text, ev) {
  const item = findSheet4AutoReply_(text);
  if (!item) return false;

  try {
    if (typeof clearBetBuffer_ === "function") clearBetBuffer_(ev);
  } catch (err) {
    console.warn("clearBetBuffer_ failed: " + err);
  }

  try {
    if (typeof getLineUserKey === "function" && typeof clearPendingBet === "function") {
      const userKey = getLineUserKey(ev);
      if (userKey) clearPendingBet(userKey);
    }
  } catch (err) {
    console.warn("clearPendingBet failed: " + err);
  }

  if (typeof replyText !== "function") {
    throw new Error("ไม่พบฟังก์ชัน replyText()");
  }

  replyText(replyToken, item.reply);
  return true;
}

function findSheet4AutoReply_(text) {
  const message = normalizeSheet4KeywordText_(text);
  if (!message) return null;

  const items = loadSheet4AutoReplies_();
  const sorted = items.slice().sort(function(a, b) {
    return b.keywordNormalized.length - a.keywordNormalized.length;
  });

  for (let i = 0; i < sorted.length; i++) {
    const item = sorted[i];
    if (item.keywordNormalized && message.indexOf(item.keywordNormalized) !== -1) {
      return item;
    }
  }

  return null;
}

function loadSheet4AutoReplies_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = "sheet4_auto_reply_v2";

  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  const sheetName =
    (typeof KEY_SHEET !== "undefined" && KEY_SHEET)
      ? KEY_SHEET
      : "ชีต4";

  if (typeof SHEET_ID === "undefined" || !SHEET_ID) {
    throw new Error("ไม่พบ SHEET_ID");
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('ไม่พบชีต "' + sheetName + '"');

  const lastRow = sh.getLastRow();
  if (lastRow < 1) return [];

  const values = sh.getRange(1, 1, lastRow, 2).getDisplayValues();
  const result = [];

  for (let i = 0; i < values.length; i++) {
    const keyword = String(values[i][0] || "").trim();
    const reply = String(values[i][1] || "").trim();
    if (!keyword || !reply) continue;

    if (
      i === 0 &&
      /^(คีย์เวิร์ด|keyword)$/i.test(keyword) &&
      /^(คำตอบ|คำตอบกลับ|ตอบกลับ|reply|response)$/i.test(reply)
    ) {
      continue;
    }

    const keywordNormalized = normalizeSheet4KeywordText_(keyword);
    if (!keywordNormalized) continue;

    result.push({
      keyword: keyword,
      keywordNormalized: keywordNormalized,
      reply: reply,
      row: i + 1
    });
  }

  try {
    cache.put(cacheKey, JSON.stringify(result), 120);
  } catch (err) {}

  return result;
}

function normalizeSheet4KeywordText_(value) {
  return String(value || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\u00A0/g, " ")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function clearSheet4AutoReplyCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove("sheet4_auto_reply_v1");
  cache.remove("sheet4_auto_reply_v2");
}

function testSheet4AutoReply_() {
  clearSheet4AutoReplyCache_();
  const items = loadSheet4AutoReplies_();

  console.log(JSON.stringify({
    count: items.length,
    items: items.slice(0, 20)
  }, null, 2));

  if (!items.length) {
    throw new Error(
      'ไม่พบคีย์เวิร์ดในชีต "' +
      ((typeof KEY_SHEET !== "undefined" && KEY_SHEET) ? KEY_SHEET : "ชีต4") +
      '" คอลัมน์ A/B'
    );
  }

  return "PASS: โหลด Auto Reply ได้ " + items.length + " รายการ";
}


// ============================================================================
// TEST PATCH 2026-09-09
// ============================================================================
function testBetFormatsV0909_() {
  const a = parseMessage(
    ["951", "325", "783", "6กับ 2"].join("\n"),
    [],
    {}
  );

  if (!a || !a.success) {
    throw new Error("6กับ อ่านไม่สำเร็จ: " + (a && a.error));
  }

  if (Number(a.totalAmount) !== 36) {
    throw new Error("6กับ ยอดต้อง 36 แต่ได้ " + a.totalAmount);
  }

  const b = parseMessage(
    ["บลก5*5", "วิน 87920"].join("\n"),
    [],
    {}
  );

  if (!b || !b.success) {
    throw new Error("วิน 87920 อ่านไม่สำเร็จ: " + (b && b.error));
  }

  if (Number(b.totalAmount) !== 200) {
    throw new Error("วิน 87920 ยอดต้อง 200 แต่ได้ " + b.totalAmount);
  }

  const c = parseMessage(
    [
      "75 76 79 69 27 26 20 บลก 10*10",
      "70 78 77 บลก 20*20",
      "57 76 86 78 73 29 บลก 10*10",
      "786 578 870 970 978 10*10"
    ].join("\n"),
    [],
    {}
  );

  if (!c || !c.success) {
    throw new Error("โพยเลขซ้ำอ่านไม่สำเร็จ: " + (c && c.error));
  }

  const merged = (typeof mergeFlexGroups === "function")
    ? mergeFlexGroups(c.groups || [])
    : (c.groups || []);

  const g10 = merged.find(function(g) {
    return (
      String(g.mode || "") === "บลก" &&
      Number(g.top || 0) === 10 &&
      Number(g.bottom || 0) === 10
    );
  });

  if (!g10) {
    throw new Error("ไม่พบกลุ่ม บลก 10*10");
  }

  const count76 = (g10.numbers || []).filter(function(n) {
    return String(n) === "76";
  }).length;

  const count67 = (g10.numbers || []).filter(function(n) {
    return String(n) === "67";
  }).length;

  if (count76 < 2 || count67 < 2) {
    throw new Error(
      "เลขซ้ำถูกตัด: 76=" + count76 + " / 67=" + count67
    );
  }

  return {
    success: true,
    sixReturnTotal: a.totalAmount,
    win87920Total: b.totalAmount,
    duplicate76: count76,
    duplicate67: count67,
    mixedTotal: c.totalAmount,
    mixedDoors: c.totalDoors
  };
}



// ============================================================================
// TEST 2026-09-16 — 2917504 3*3/2*6
// ============================================================================
function testCompactWinTwoThreeV0916_() {
  const samples = [
    "2917504 3*3/2*6",
    "นอย 2917504 3*3/2*6"
  ];

  samples.forEach(function(sample) {
    const result = parseMessage(sample, [], {});

    if (!result || !result.success) {
      throw new Error(
        "FAIL: " + sample + " / " +
        String(result && result.error ? result.error : "parse failed")
      );
    }

    const groups = result.groups || [];

    const win2 = groups.find(function(g) {
      return g.mode === "บลก" &&
        Number(g.top) === 3 &&
        Number(g.bottom) === 3;
    });

    const win3 = groups.find(function(g) {
      return g.mode === "6กลับ" &&
        Number(g.top) === 2;
    });

    if (!win2 || (win2.numbers || []).length !== 42) {
      throw new Error("FAIL วิน 2 ตัว: " + sample);
    }

    if (Number(win2.subtotal) !== 252) {
      throw new Error("FAIL วิน 2 ตัวยอดไม่ใช่ 252: " + sample);
    }

    if (!win3 || (win3.win3Source || []).length !== 35) {
      throw new Error("FAIL วิน 3 ตัว: " + sample);
    }

    if (Number(win3.subtotal) !== 420) {
      throw new Error("FAIL วิน 3 ตัวยอดไม่ใช่ 420: " + sample);
    }

    if (Number(result.totalAmount) !== 672) {
      throw new Error(
        "FAIL ยอดรวม: " + sample +
        " ได้ " + String(result.totalAmount)
      );
    }
  });

  return {
    success: true,
    compact: "2917504 3*3/2*6",
    prefixed: "นอย 2917504 3*3/2*6",
    totalAmount: 672
  };
}


// ============================================================================
// TEST 2026-09-21 — รูปแบบวินย่อใหม่
// ============================================================================
function testCompactWinShortcutsV0921_() {
  const a = parseMessage("7045239 2*6", [], {});
  if (!a || !a.success || Number(a.totalAmount) !== 420) {
    throw new Error(
      "FAIL 7045239 2*6: " +
      String(a && a.totalAmount)
    );
  }

  const b = parseMessage("780314 บลก 1*1", [], {});
  if (!b || !b.success || Number(b.totalAmount) !== 60) {
    throw new Error(
      "FAIL 780314 บลก 1*1: " +
      String(b && b.totalAmount)
    );
  }

  const c = parseMessage("780314 บต1*1", [], {});
  if (!c || !c.success || Number(c.totalAmount) !== 40) {
    throw new Error(
      "FAIL 780314 บต1*1: " +
      String(c && c.totalAmount)
    );
  }

  const d = parseMessage("รูด0=5*5", [], {});
  if (!d || !d.success || Number(d.totalAmount) !== 190) {
    throw new Error(
      "FAIL รูด0=5*5: " +
      String(d && d.totalAmount)
    );
  }

  return {
    success: true,
    win3SixReturn: a.totalAmount,
    win2BLG: b.totalAmount,
    win3DirectTod: c.totalAmount,
    rood0: d.totalAmount
  };
}


// ============================================================================
// TEST OCR 2026-09-21 — เลขคั่น / และ :
// ============================================================================
function testOcrSlashColonTwoDigitListV0921_() {
  const lines = [
    "ฮานอยรวมมิตร",
    "17",
    "10/13/15/16/",
    "18/19",
    "70/73/75/76/",
    "78/79",
    "กัน",
    "11:14:74:77:72"
  ];

  const result = extractLotteryBetNumbersFromOcrLines_(
    lines,
    lines.join("\\n")
  );

  const expected = [
    "17",
    "10", "13", "15", "16",
    "18", "19",
    "70", "73", "75", "76",
    "78", "79",
    "11", "14", "74", "77", "72"
  ];

  const got = result.twoDigits || [];

  if (got.join(",") !== expected.join(",")) {
    throw new Error(
      "FAIL OCR slash/colon\\n" +
      "got: " + got.join(" ") + "\\n" +
      "expected: " + expected.join(" ")
    );
  }

  if (!result.isLotteryBet) {
    throw new Error("FAIL: ต้องตรวจเป็นรูปโพยหวย");
  }

  return {
    success: true,
    twoDigits: got,
    count: got.length,
    betText: buildLotteryBetTextFromImage_(result)
  };
}


// ============================================================================
// TEST OCR — รูปฮานอยมัดรวม
// ============================================================================
function testOcrHanoiMudRuamV0921_() {
  const lines = [
    "16:37",
    "แนวทา... (183)",
    "ฮานอยมัดรวม",
    "1 7",
    "10 / 13 / 15 / 16 /",
    "18/19",
    "70/73/75/76/",
    "78/79",
    "กัน",
    "11 : 14 : 74 : 77 : 72",
    "@เทพเดินดงค์",
    "16:32"
  ];

  const parsed = extractLotteryBetNumbersFromOcrLines_(
    lines,
    lines.join("\\n")
  );

  const expected = [
    "17",
    "10", "13", "15", "16",
    "18", "19",
    "70", "73", "75", "76",
    "78", "79",
    "11", "14", "74", "77", "72"
  ];

  if ((parsed.twoDigits || []).join(",") !== expected.join(",")) {
    throw new Error(
      "FAIL OCR ฮานอยมัดรวม: " +
      (parsed.twoDigits || []).join(" ")
    );
  }

  if ((parsed.threeDigits || []).length !== 0) {
    throw new Error("FAIL: ดูดเลข 3 ตัวจาก thumbnail");
  }

  return {
    success: true,
    count: parsed.twoDigits.length,
    numbers: parsed.twoDigits,
    betText: buildLotteryBetTextFromImage_(parsed)
  };
}


// ============================================================================
// TEST — รูปก่อน แล้วพิมพ์ราคาทีหลัง
// ============================================================================
function testImageBetPriceAfterImageV0921_() {
  const p1 = parseImageBetPriceText_("กับ10*10");
  const p2 = parseImageBetPriceText_("10×20");
  const p3 = parseImageBetPriceText_("บลก 5*5");

  if (!p1 || p1.first !== 10 || p1.second !== 10) {
    throw new Error("FAIL ราคา กับ10*10");
  }

  if (!p2 || p2.first !== 10 || p2.second !== 20) {
    throw new Error("FAIL ราคา 10×20");
  }

  if (!p3 || p3.first !== 5 || p3.second !== 5) {
    throw new Error("FAIL ราคา บลก 5*5");
  }

  const betText =
    buildLotteryBetTextFromImageWithPrice_(
      {
        twoDigits: ["17", "10", "13", "15", "16"],
        threeDigits: ["523", "826"]
      },
      p1
    );

  const expected =
    "17 10 13 15 16 บลก 10*10\\n" +
    "523 826 บต 10*10";

  if (betText !== expected) {
    throw new Error(
      "FAIL betText\\n" +
      betText
    );
  }

  return {
    success: true,
    price: p1,
    betText: betText
  };
}


// ============================================================================
// TEST ROI — thumbnail ด้านบนต้องไม่ถูกเลือก
// ============================================================================
function testImageBetRoiV0921_() {
  const ocr = {
    paragraphs: [
      {
        text: "4416 5526 4822 416-26",
        xMin: 0.10,
        xMax: 0.72,
        yMin: 0.08,
        yMax: 0.17
      },
      {
        text: "ฮานอยมัดรวม",
        xMin: 0.12,
        xMax: 0.58,
        yMin: 0.29,
        yMax: 0.34
      },
      {
        text: "17\n10/13/15/16/\n18/19\n70/73/75/76/\n78/79",
        xMin: 0.13,
        xMax: 0.67,
        yMin: 0.35,
        yMax: 0.63
      },
      {
        text: "กัน\n11:14:74:77:72",
        xMin: 0.13,
        xMax: 0.69,
        yMin: 0.64,
        yMax: 0.72
      },
      {
        text: "16:32",
        xMin: 0.86,
        xMax: 0.97,
        yMin: 0.80,
        yMax: 0.84
      }
    ]
  };

  const region =
    selectLotteryOcrRegion_(ocr);

  if (
    region.text.indexOf(
      "4416 5526 4822 416-26"
    ) !== -1
  ) {
    throw new Error(
      "FAIL: ROI ยังอ่าน thumbnail ด้านบน"
    );
  }

  if (
    region.text.indexOf(
      "10/13/15/16/"
    ) === -1
  ) {
    throw new Error(
      "FAIL: ROI ไม่ได้อ่านโพยหลัก"
    );
  }

  return {
    success: true,
    text: region.text
  };
}


// ============================================================================
// TEST — อ่านเฉพาะเลขโพย ห้ามวันที่/เวลา/เลข UI
// ============================================================================
function testOcrStrictBetOnlyV0921_() {
  const lines = [
    "14:23",
    "21/09/69",
    "2026 09 21 12:28 น.",
    "(183)",
    "ฮานอยมัดรวม",
    "17",
    "10/13/15/16/",
    "18/19",
    "70/73/75/76/",
    "78/79",
    "กัน",
    "11:14:74:77:72",
    "16:31 น."
  ];

  const filtered =
    lines.filter(function(line) {
      return !isOcrNonBetMetadataLine_(line);
    });

  if (filtered.indexOf("14:23") !== -1) {
    throw new Error("FAIL: เวลา 14:23 หลุดเข้ามา");
  }

  if (filtered.indexOf("21/09/69") !== -1) {
    throw new Error("FAIL: วันที่ 21/09/69 หลุดเข้ามา");
  }

  if (filtered.indexOf("(183)") !== -1) {
    throw new Error("FAIL: เลข UI (183) หลุดเข้ามา");
  }

  if (filtered.indexOf("11:14:74:77:72") === -1) {
    throw new Error("FAIL: โพย colon ถูกกรองทิ้ง");
  }

  const parsed =
    extractLotteryBetNumbersFromOcrLines_(
      filtered,
      filtered.join("\\n")
    );

  const expected = [
    "17",
    "10", "13", "15", "16",
    "18", "19",
    "70", "73", "75", "76",
    "78", "79",
    "11", "14", "74", "77", "72"
  ];

  if (
    (parsed.twoDigits || []).join(",") !==
    expected.join(",")
  ) {
    throw new Error(
      "FAIL เลขโพย\\nได้: " +
      (parsed.twoDigits || []).join(" ") +
      "\\nควรได้: " +
      expected.join(" ")
    );
  }

  return {
    success: true,
    twoDigits: parsed.twoDigits,
    ignored: [
      "14:23",
      "21/09/69",
      "2026 09 21 12:28 น.",
      "(183)",
      "16:31 น."
    ]
  };
}


// ============================================================================
// TEST — ราคา OCR + ชื่อหวย
// ============================================================================



// ============================================================================
// TEST — ราคา OCR + ชื่อหวย (FIX ไม่ตอบ)
// ============================================================================
function testImageBetPriceAndLottoV0921_() {
  const cases = [
    ["10*10", 10, 10, ""],
    ["กับ10*10", 10, 10, ""],
    ["กัน10*10", 10, 10, ""],
    ["บลก10*10", 10, 10, ""],
    ["บต10*10", 10, 10, ""],
    ["10×10", 10, 10, ""],
    ["ฮานอย 10*10", 10, 10, "ฮานอย"],
    ["ฮานอย กับ10*10", 10, 10, "ฮานอย"],
    ["ฮานอย บลก 10*10", 10, 10, "ฮานอย"],
    ["10*10 ฮานอย", 10, 10, "ฮานอย"]
  ];

  cases.forEach(function(item) {
    const p =
      parseImageBetPriceAndLotto_(item[0]);

    if (
      !p ||
      Number(p.first) !== item[1] ||
      Number(p.second) !== item[2] ||
      String(p.lottoText || "") !== item[3]
    ) {
      throw new Error(
        "FAIL: " +
        item[0] +
        " => " +
        JSON.stringify(p)
      );
    }
  });

  return {
    success: true,
    tested: cases.length
  };
}


// ============================================================================
// TEST — pending OCR ต้องมี durable fallback
// ============================================================================
function testImageBetPendingStorageV0921_() {
  const source =
    String(savePendingImageBet_);

  if (
    source.indexOf(
      "IMAGE_BET_PENDING_SHEET_"
    ) === -1
  ) {
    throw new Error(
      "FAIL: save pending ยังไม่มี Sheet fallback"
    );
  }

  const loadSource =
    String(loadPendingImageBet_);

  if (
    loadSource.indexOf(
      "IMAGE_BET_PENDING_LOADED_FROM_SHEET"
    ) === -1
  ) {
    throw new Error(
      "FAIL: load pending ยังไม่มี Sheet fallback"
    );
  }

  return {
    success: true,
    sheet:
      IMAGE_BET_PENDING_SHEET_
  };
}


// ============================================================================
// TEST — หลังเลือกราคาจากรูปแล้ว กด Quick Reply ชื่อหวย
// ต้องไม่ถูกตีเป็น "pending missing"
// ============================================================================
function testQuickReplyAfterImagePriceV0921_() {
  function isPriceOnlyForTest_(value) {
    const rawText =
      String(value || "")
        .replace(/[×✕✖⨯xX]/g, "*")
        .trim();

    return /^(?:ตรงโต๊ดกลับ|บนล่างกลับ|ตรงโต๊ด|ตัวละ|ราคา|บลก|บกล|บนล่าง|บต|กับ|กัน|กลับ)?\s*=?\s*\d+(?:\.\d+)?\s*\*\s*\d+(?:\.\d+)?$/i.test(
      rawText
    );
  }

  const priceOnly = [
    "10*10",
    "กับ10*10",
    "บลก10*10",
    "10×10"
  ];

  priceOnly.forEach(function(v) {
    if (!isPriceOnlyForTest_(v)) {
      throw new Error(
        "FAIL price-only: " + v
      );
    }
  });

  const quickReplyPayload =
    "สิงคโปร์\n" +
    "07 10 13 15 16 18 19 70 73 75 76 78 79 11 14 74 77 72 บลก 10*10";

  if (isPriceOnlyForTest_(quickReplyPayload)) {
    throw new Error(
      "FAIL: Quick Reply payload ถูกตีเป็นราคาอย่างเดียว"
    );
  }

  return {
    success: true,
    quickReplyPassesToNormalParser: true
  };
}


// ============================================================================
// TEST — สลิปต้องไม่อ่านเป็นโพย
// ============================================================================
function testBankSlipMustNotBecomeBetV0921_() {
  const scbText = [
    "SCB",
    "โอนเงินสำเร็จ",
    "21 ก.ย. 2569 - 19:57",
    "จาก นาย ก",
    "ไปยัง นางสาว ชื่อบัญชี",
    "จำนวนเงิน 180.00 บาท",
    "เลขที่รายการ 1234567890"
  ].join("\n");

  const slip =
    classifyBankSlipFromOcrText_(
      scbText
    );

  if (!slip.isBankSlip) {
    throw new Error(
      "FAIL: SCB slip ไม่ถูกจัดเป็นสลิป"
    );
  }

  const betText = [
    "ฮานอยมัดรวม",
    "17",
    "10/13/15/16/",
    "18/19",
    "70/73/75/76/",
    "78/79",
    "กัน",
    "11:14:74:77:72"
  ].join("\n");

  const bet =
    classifyBankSlipFromOcrText_(
      betText
    );

  if (bet.isBankSlip) {
    throw new Error(
      "FAIL: รูปโพยถูกจัดเป็นสลิป"
    );
  }

  return {
    success: true,
    slipScore: slip.score,
    slipMarkers: slip.markers
  };
}


// ============================================================================
// TEST — Bangkok Bank OCR ไม่ผ่าน Thunder
// ============================================================================
function testBangkokBankSlipOcrV0922_() {
  const sample = [
    "Bangkok Bank",
    "โอนเงินสำเร็จ",
    "22 ก.ย. 2569 - 18:00",
    "จาก",
    "นาย สมชาย ใจดี",
    "XXX-X-X123-X",
    "ไปยัง",
    "น.ส. ชื่อบัญชี นามสกุลบัญชี",
    "SCB",
    "XXX-X-X114-X",
    "จำนวนเงิน",
    "180.00 บาท",
    "เลขอ้างอิง ABC123456789"
  ].join("\\n");

  if (
    !isBangkokBankOriginSlipOcr_(
      sample
    )
  ) {
    throw new Error(
      "FAIL: ไม่ตรวจเป็นสลิปต้นทางธนาคารกรุงเทพ"
    );
  }

  const parsed =
    parseBangkokBankSlipFromOcr_(
      sample
    );

  if (!parsed.senderOk) {
    throw new Error(
      "FAIL: อ่านชื่อผู้โอนไม่ได้"
    );
  }

  if (!parsed.receiverOk) {
    throw new Error(
      "FAIL: ผู้รับไม่ตรงบัญชีร้าน"
    );
  }

  if (
    Number(parsed.amount) !== 180
  ) {
    throw new Error(
      "FAIL: ยอดต้อง 180 แต่ได้ " +
      String(parsed.amount)
    );
  }

  if (!parsed.success) {
    throw new Error(
      "FAIL: BBL OCR ต้องผ่าน"
    );
  }

  return {
    success: true,
    sender:
      parsed.senderName,
    receiver:
      parsed.receiverName,
    amount:
      parsed.amount,
    source:
      parsed.data.verificationSource
  };
}


// ============================================================================
// TEST 2026-09-23
// รูปแบบ: 34 10*กลับ บ-ล = บนล่างกลับ 10 บาท
// ============================================================================
function testPriceStarReturnUpperLowerV0923_() {
  const cases = [
    ["34 10*กลับ บ-ล", "34 บลก 10*10"],
    ["68 10* กลับ บ-ล", "68 บลก 10*10"],
    ["34 10*กลับ บ/ล", "34 บลก 10*10"],
    ["34 10* บ-ล กลับ", "34 บลก 10*10"],
    ["34 68 10*กลับ บ-ล", "34 68 บลก 10*10"]
  ];

  cases.forEach(function(item) {
    const got =
      normalizeNaturalTwoDigitAndRoodForms_(
        item[0]
      );

    if (got !== item[1]) {
      throw new Error(
        "FAIL: " +
        item[0] +
        " => " +
        got +
        " | expected " +
        item[1]
      );
    }
  });

  // ตัวอย่างเต็มตามที่ลูกค้าส่ง
  const parsed =
    parseMessage(
      [
        "34 10*กลับ บ-ล",
        "68 10* กลับ บ-ล"
      ].join("\n"),
      [],
      {}
    );

  if (!parsed || !parsed.success) {
    throw new Error(
      "FAIL parse: " +
      String(
        parsed && parsed.error
          ? parsed.error
          : "unknown"
      )
    );
  }

  // 34/43 และ 68/86
  // แต่ละเลข บน 10 + ล่าง 10
  // 4 เลข x 20 = 80 บาท
  if (Number(parsed.totalAmount) !== 80) {
    throw new Error(
      "FAIL total: expected 80, got " +
      String(parsed.totalAmount)
    );
  }

  return {
    success: true,
    totalAmount: parsed.totalAmount,
    groups: parsed.groups
  };
}


// ============================================================================
// TEST 2026-09-23
// บ/ล/ก 13,14,15,34,35,45 ตัว 20
// ============================================================================
function testCompactBothReturnPerNumberV0923_() {
  const raw =
    "บ/ล/ก 13,14,15,34,35,45 ตัว 20";

  const normalized =
    normalizeNaturalTwoDigitAndRoodForms_(raw);

  const expected =
    "13 14 15 34 35 45 บลก 20*20";

  if (normalized !== expected) {
    throw new Error(
      "FAIL normalize: " +
      normalized +
      " | expected " +
      expected
    );
  }

  const parsed =
    parseMessage(
      raw,
      [],
      {}
    );

  if (!parsed || !parsed.success) {
    throw new Error(
      "FAIL parse: " +
      String(
        parsed && parsed.error
          ? parsed.error
          : "unknown"
      )
    );
  }

  // 6 เลขต้นทาง -> กลับเป็น 12 เลข
  // บน 20 + ล่าง 20 = 40 ต่อเลข
  // 12 * 40 = 480 บาท
  if (Number(parsed.totalAmount) !== 480) {
    throw new Error(
      "FAIL total: expected 480, got " +
      String(parsed.totalAmount)
    );
  }

  return {
    success: true,
    normalized: normalized,
    totalAmount: parsed.totalAmount,
    groups: parsed.groups
  };
}


// ============================================================================
// TEST 2026-09-23 — Bangkok Bank long slip
// ============================================================================
function testBangkokLongSlipV0923_() {
  const sample = [
    "Bangkok Bank",
    "รายการสำเร็จ",
    "23 ก.ย. 69, 19:22",
    "จำนวนเงิน",
    "80.00 THB",
    "จาก",
    "น.ส. ทัศนีย์",
    "624-0-xxx863",
    "ธนาคารกรุงเทพ",
    "ไปที่",
    "นางสาว ชื่อบัญชี นามสกุลบัญชี",
    "505-4-xxx114",
    "ธนาคารไทยพาณิชย์",
    "ค่าธรรมเนียม",
    "0.00 THB",
    "หมายเลขอ้างอิง",
    "624741",
    "เลขที่อ้างอิง",
    "20260923192212240080620"
  ].join("\n");

  const parsed =
    parseBangkokBankSlipFromOcr_(
      sample
    );

  if (
    parsed.senderName !==
    "น.ส. ทัศนีย์"
  ) {
    throw new Error(
      "FAIL sender: " +
      String(parsed.senderName)
    );
  }

  if (
    parsed.receiverName !==
    "นางสาว ชื่อบัญชี นามสกุลบัญชี"
  ) {
    throw new Error(
      "FAIL receiver: " +
      String(parsed.receiverName)
    );
  }

  if (
    Number(parsed.amount) !== 80
  ) {
    throw new Error(
      "FAIL amount: " +
      String(parsed.amount)
    );
  }

  if (!parsed.senderOk) {
    throw new Error(
      "FAIL senderOk"
    );
  }

  if (!parsed.receiverOk) {
    throw new Error(
      "FAIL receiverOk"
    );
  }

  if (!parsed.success) {
    throw new Error(
      "FAIL success"
    );
  }

  return {
    success: true,
    senderName: parsed.senderName,
    receiverName: parsed.receiverName,
    amount: parsed.amount
  };
}


// ============================================================================
// TEST 2026-09-23 — รูด 7 ตัว
// ============================================================================
function testRoodSevenDigitsV0923_() {
  const doorMap = {};

  const ok =
    parseRoodBlock(
      "รูด1234567",
      "รูด",
      1,
      1,
      doorMap
    );

  if (!ok || !ok.success) {
    throw new Error(
      "FAIL: รูด 7 ตัวต้องผ่าน"
    );
  }

  const tooMany =
    parseRoodBlock(
      "รูด12345678",
      "รูด",
      1,
      1,
      doorMap
    );

  if (
    !tooMany ||
    tooMany.success ||
    String(tooMany.error || "") !==
      "❌ รูดได้ไม่เกิน 6 ตัว"
  ) {
    throw new Error(
      "FAIL: รูด 8 ตัวต้องแจ้งตามข้อความที่กำหนด"
    );
  }

  return {
    success: true,
    maxAccepted: 7,
    overLimitMessage:
      tooMany.error
  };
}


// ============================================================================
// TEST 2026-09-23 — Bangkok Bank OCR target fix
// ============================================================================
function testBangkokBankReceiverFixV0923_() {
  const sample = [
    "Bangkok Bank",
    "รายการสำเร็จ",
    "23 ก.ย. 69, 19:22",
    "จำนวนเงิน",
    "80.00 THB",
    "จาก",
    "น.ส. ทัศนีย์",
    "624-0-xxx863",
    "ธนาคารกรุงเทพ",
    "ไปที่",
    "Bangkok",
    "นางสาว ชื่อบัญชี นามสกุลบัญชี",
    "505-4-xxx114",
    "ธนาคารไทยพาณิชย์",
    "ค่าธรรมเนียม",
    "0.00 THB"
  ].join("\\n");

  const parsed =
    parseBangkokBankSlipFromOcr_(
      sample
    );

  if (!parsed.success) {
    throw new Error(
      "FAIL: Bangkok Bank slip ต้องผ่าน"
    );
  }

  if (
    parsed.receiverName !==
    BANK_ACCOUNT_NAME
  ) {
    throw new Error(
      "FAIL receiverName: " +
      String(parsed.receiverName)
    );
  }

  if (
    parsed.senderName.indexOf(
      "ทัศนีย์"
    ) === -1
  ) {
    throw new Error(
      "FAIL senderName: " +
      String(parsed.senderName)
    );
  }

  if (
    Number(parsed.amount) !== 80
  ) {
    throw new Error(
      "FAIL amount: " +
      String(parsed.amount)
    );
  }

  const receiver =
    parsed.data.rawSlip.receiver;

  if (
    getSlipPartyName(receiver) !==
    BANK_ACCOUNT_NAME
  ) {
    throw new Error(
      "FAIL Flex receiver name"
    );
  }

  if (
    getSlipPartyBankText(receiver) !==
    "SCB"
  ) {
    throw new Error(
      "FAIL Flex receiver bank"
    );
  }

  return {
    success: true,
    sender: parsed.senderName,
    receiver: parsed.receiverName,
    amount: parsed.amount
  };
}


// ============================================================================
// TEST 2026-09-24
// 16,17,19,15,67,69,65,79,75,95=2×2บลก
// ============================================================================
function testExplicitTwoDigitEqualsBLGV0924_() {
  const input =
    "16,17,19,15,67,69,65,79,75,95=2×2บลก";

  const extracted =
    extractExplicitTwoDigitEqualsModeBets_(
      input
    );

  if (
    !extracted.success ||
    !extracted.groups ||
    extracted.groups.length !== 1
  ) {
    throw new Error(
      "FAIL: extractor ไม่ได้ 1 กลุ่ม"
    );
  }

  const g =
    extracted.groups[0];

  if (g.mode !== "บลก") {
    throw new Error(
      "FAIL mode: " + g.mode
    );
  }

  if (
    Number(g.top) !== 2 ||
    Number(g.bottom) !== 2
  ) {
    throw new Error(
      "FAIL price"
    );
  }

  // 10 เลขต้นทาง ไม่มีเบิ้ล
  // กลับเลขแล้ว = 20 รายการ
  if (
    (g.numbers || []).length !== 20
  ) {
    throw new Error(
      "FAIL numbers: " +
      (g.numbers || []).length
    );
  }

  // 20 รายการ × (บน2 + ล่าง2) = 80
  if (Number(g.subtotal) !== 80) {
    throw new Error(
      "FAIL subtotal: " +
      String(g.subtotal)
    );
  }

  // ทดสอบทั้งข้อความตามที่ลูกค้าส่ง
  const full = parseMessage(
    [
      "รูด 4=2×2บลก",
      "",
      input,
      "",
      "รูดเบิ้ล 2×2บลก"
    ].join("\n"),
    [],
    {}
  );

  if (
    !full ||
    !full.success
  ) {
    throw new Error(
      "FAIL full parse: " +
      String(
        full && full.error
          ? full.error
          : "unknown"
      )
    );
  }

  // รูด4 = 76
  // เลขชุด = 80
  // รูดเบิ้ล = 40
  // รวม = 196
  if (
    Number(full.totalAmount) !== 196
  ) {
    throw new Error(
      "FAIL total: expected 196, got " +
      String(full.totalAmount)
    );
  }

  return {
    success: true,
    listSubtotal: g.subtotal,
    totalAmount: full.totalAmount
  };
}


// ============================================================================
// TEST 2026-10-01 — Bangkok Bank auto accept
// ============================================================================
function testBangkokBankAutoAcceptNoVerifyV1001_() {
  const source =
    String(handleIncomingLineImage_);

  if (
    source.indexOf(
      "BBL_SLIP_AUTO_ACCEPT"
    ) === -1
  ) {
    throw new Error(
      "FAIL: ไม่มี Bangkok Bank auto-accept branch"
    );
  }

  if (
    source.indexOf(
      "buildSlipThankYouFlex()"
    ) === -1
  ) {
    throw new Error(
      "FAIL: Bangkok Bank ไม่ได้ส่ง Flex รับยอดทันที"
    );
  }

  return {
    success: true,
    mode: "NO_VERIFY_AUTO_ACCEPT"
  };
}


// ============================================================================
// TEST 2026-11-11 — ตรง 5×5 / กลับ 3×3 แยกราคา
// ============================================================================
function testDirectAndReturnSeparatePriceV1111_() {
  const input = [
    "80 81 83 84 85 86 87 89",
    "บนล่าง@5×5",
    "กลับ@3×3"
  ].join("\n");

  const result = parseMessage(input, [], {});

  if (!result || !result.success) {
    throw new Error(
      "FAIL parse: " +
      String(result && result.error ? result.error : "unknown")
    );
  }

  const groups = result.groups || [];
  if (groups.length !== 2) {
    throw new Error("FAIL groups: " + groups.length);
  }

  const direct = groups.find(function(g) {
    return g.mode === "บล" &&
      Number(g.top) === 5 &&
      Number(g.bottom) === 5;
  });

  const returned = groups.find(function(g) {
    return g.mode === "กลับ" &&
      Number(g.top) === 3 &&
      Number(g.bottom) === 3;
  });

  if (!direct || (direct.numbers || []).length !== 8) {
    throw new Error("FAIL direct group");
  }

  const expectedDirect =
    "80 81 83 84 85 86 87 89";

  if ((direct.numbers || []).join(" ") !== expectedDirect) {
    throw new Error(
      "FAIL direct nums: " +
      (direct.numbers || []).join(" ")
    );
  }

  const expectedReturn =
    "08 18 38 48 58 68 78 98";

  if (!returned || (returned.numbers || []).join(" ") !== expectedReturn) {
    throw new Error(
      "FAIL return nums: " +
      String(returned && returned.numbers ? returned.numbers.join(" ") : "")
    );
  }

  if (Number(direct.subtotal) !== 80) {
    throw new Error("FAIL direct subtotal: " + direct.subtotal);
  }

  if (Number(returned.subtotal) !== 48) {
    throw new Error("FAIL return subtotal: " + returned.subtotal);
  }

  if (Number(result.totalAmount) !== 128) {
    throw new Error("FAIL total: " + result.totalAmount);
  }

  if (Number(result.totalDoors) !== 16) {
    throw new Error("FAIL doors: " + result.totalDoors);
  }

  return {
    success: true,
    totalAmount: result.totalAmount,
    totalDoors: result.totalDoors,
    direct: direct.numbers,
    returned: returned.numbers
  };
}


// ============================================================================
// TEST 2026-11-11 — ราคาแบบต่อเนื่อง
// ============================================================================
function testStickyTwoDigitPriceGroupsV1111_() {
  const input = [
    "82=5×5",
    "84",
    "81",
    "85",
    "86",
    "87",
    "83",
    "38=3×3",
    "78",
    "68",
    "58",
    "18",
    "48",
    "28",
    "บนล่าง"
  ].join("\n");

  const extracted =
    extractStickyTwoDigitPriceGroups_(input);

  if (
    !extracted.success ||
    !extracted.groups ||
    extracted.groups.length !== 2
  ) {
    throw new Error(
      "FAIL sticky groups"
    );
  }

  const g1 = extracted.groups[0];
  const g2 = extracted.groups[1];

  if (
    g1.numbers.join(" ") !==
      "82 84 81 85 86 87 83" ||
    Number(g1.top) !== 5 ||
    Number(g1.bottom) !== 5 ||
    Number(g1.subtotal) !== 70
  ) {
    throw new Error(
      "FAIL sticky group1"
    );
  }

  if (
    g2.numbers.join(" ") !==
      "38 78 68 58 18 48 28" ||
    Number(g2.top) !== 3 ||
    Number(g2.bottom) !== 3 ||
    Number(g2.subtotal) !== 42
  ) {
    throw new Error(
      "FAIL sticky group2"
    );
  }

  const full = parseMessage(
    input,
    [],
    {}
  );

  if (
    !full ||
    !full.success ||
    Number(full.totalAmount) !== 112 ||
    Number(full.totalDoors) !== 14
  ) {
    throw new Error(
      "FAIL sticky full: " +
      JSON.stringify(full)
    );
  }

  return {
    success: true,
    totalAmount: full.totalAmount,
    totalDoors: full.totalDoors
  };
}


// ============================================================================
// TEST 2026-11-11
// ============================================================================
function testTrailingPriceModeTwoDigitBetsV1111_() {
  const input = [
    "02/19/36/45",
    "37/81/01",
    "8บลก",
    "11=40บล"
  ].join("\n");

  const r =
    extractTrailingPriceModeTwoDigitBets_(
      input
    );

  if (!r.success) {
    throw new Error("FAIL extractor");
  }

  if (
    !r.groups ||
    r.groups.length !== 2
  ) {
    throw new Error(
      "FAIL groups: " +
      String(
        r.groups &&
        r.groups.length
      )
    );
  }

  // 7 เลข บลก => กลับเป็น 14 รายการ
  // 14 * (8+8) = 224
  if (
    Number(r.groups[0].subtotal) !== 224
  ) {
    throw new Error(
      "FAIL group1 subtotal: " +
      String(r.groups[0].subtotal)
    );
  }

  // 11 บล 40*40 = 80
  if (
    Number(r.groups[1].subtotal) !== 80
  ) {
    throw new Error(
      "FAIL group2 subtotal: " +
      String(r.groups[1].subtotal)
    );
  }

  const total =
    r.groups.reduce(
      function(sum, g) {
        return sum +
          Number(g.subtotal || 0);
      },
      0
    );

  if (total !== 304) {
    throw new Error(
      "FAIL total: " +
      String(total)
    );
  }

  return {
    success: true,
    group1Subtotal:
      r.groups[0].subtotal,
    group2Subtotal:
      r.groups[1].subtotal,
    totalAmount: total
  };
}


// ============================================================================
// TEST 2026-11-15 — รุด6/9 + ราคาแยกบรรทัด
// ============================================================================
function testRoodTypoSeparatePriceV1115_() {
  const input = [
    "รุด6/9",
    "5*5",
    "นอยทีวี"
  ].join("\n");

  const normalized =
    normalizeRoodSeparatePriceLines_(
      input
    );

  const expected = [
    "รูด 6/9 บล 5*5",
    "นอยทีวี"
  ].join("\n");

  if (normalized !== expected) {
    throw new Error(
      "FAIL normalize: " +
      normalized
    );
  }

  // ทดสอบตัว parser รูดโดยตรง
  const result =
    parseRoodBlock(
      "รูด 6/9 บล",
      "รูด",
      5,
      5,
      {}
    );

  if (
    !result ||
    !result.success
  ) {
    throw new Error(
      "FAIL parseRoodBlock"
    );
  }

  if (
    Number(result.betCount) !== 38
  ) {
    throw new Error(
      "FAIL betCount: " +
      String(result.betCount)
    );
  }

  if (
    Number(result.subtotal) !== 380
  ) {
    throw new Error(
      "FAIL subtotal: " +
      String(result.subtotal)
    );
  }

  return {
    success: true,
    normalized: normalized,
    betCount: result.betCount,
    totalAmount: result.subtotal
  };
}


// ============================================================================
// REGRESSION TEST — เคสจาก LINE จริง 2026-11-15
// ============================================================================
function testLiveCasesV1115Fix_() {
  // เคส 1: ต้องรวม 304
  const case1 = [
    "02/19/36/45",
    "37/81/01",
    "8บลก",
    "11=40บล"
  ].join("\n");

  const r1 =
    extractTrailingPriceModeTwoDigitBets_(case1);

  if (
    !r1.success ||
    !r1.groups ||
    r1.groups.length !== 2
  ) {
    throw new Error("CASE1: ต้องได้ 2 กลุ่ม");
  }

  if (
    Number(r1.groups[0].subtotal) !== 224 ||
    Number(r1.groups[1].subtotal) !== 80
  ) {
    throw new Error(
      "CASE1 subtotal ผิด: " +
      JSON.stringify(
        r1.groups.map(function(g) {
          return {
            mode: g.mode,
            top: g.top,
            bottom: g.bottom,
            count: g.betCount,
            subtotal: g.subtotal
          };
        })
      )
    );
  }

  const case1Total =
    Number(r1.groups[0].subtotal) +
    Number(r1.groups[1].subtotal);

  if (case1Total !== 304) {
    throw new Error(
      "CASE1 total ต้องเป็น 304 แต่ได้ " +
      case1Total
    );
  }

  // เคส 2: รูด7/5 ราคาแยกบรรทัด
  const case2 = [
    "รูด7/5",
    "5*5",
    "ฮั่ง"
  ].join("\n");

  const normalized =
    normalizeRoodSeparatePriceLines_(case2);

  const expected =
    [
      "รูด 7/5 บล 5*5",
      "ฮั่ง"
    ].join("\n");

  if (normalized !== expected) {
    throw new Error(
      "CASE2 normalize ผิด: " +
      normalized
    );
  }

  const rood =
    parseRoodBlock(
      "รูด 7/5 บล",
      "รูด",
      5,
      5,
      {}
    );

  if (
    !rood ||
    !rood.success ||
    Number(rood.betCount) !== 38 ||
    Number(rood.subtotal2) !== 380
  ) {
    throw new Error(
      "CASE2 รูด7/5 ต้อง 38 รายการ 380 บาท: " +
      JSON.stringify(rood)
    );
  }

  return {
    success: true,
    version: LEGACY_BOT_PATCH_VERSION_,
    case1Total: case1Total,
    case2BetCount: rood.betCount,
    case2Total: rood.subtotal2
  };
}


// ============================================================================
// TEST — PRE-PARSE จริงก่อนเข้า parseMessage
// ============================================================================
function testMainPreParseCustomV1115_() {
  const a = [
    "🇻🇳 ฮานอย",
    "02/19/36/45",
    "37/81/01",
    "8บลก",
    "11=40บล",
    "ปกติ"
  ].join("\n");

  const na = normalizeMainPreParseCustomBet_(a, "ฮานอยปกติ");
  const expectedA = [
    "ฮานอยปกติ",
    "02 19 36 45 37 81 01 บลก 8*8",
    "11 บล 40*40"
  ].join("\n");

  if (na !== expectedA) {
    throw new Error("PREPARSE A FAIL: " + na);
  }

  const b = [
    "🇭🇰 ฮั่งเส็งบ่าย",
    "รุด7/5",
    "5*5",
    "ฮั่ง"
  ].join("\n");

  const nb = normalizeMainPreParseCustomBet_(b, "ฮั่งเส็งบ่าย");
  const expectedB = [
    "ฮั่งเส็งบ่าย",
    "รูด 7/5 บล 5*5"
  ].join("\n");

  if (nb !== expectedB) {
    throw new Error("PREPARSE B FAIL: " + nb);
  }

  return {success:true, a:na, b:nb};
}


// ============================================================================
// REGRESSION TEST — RAW HIGH PRIORITY
// ============================================================================
function testRawHighPriorityBetSyntaxV1115_() {
  const a = [
    "🇻🇳 ฮานอย",
    "02/19/36/45",
    "37/81/01",
    "8บลก",
    "11=40บล",
    "ปกติ"
  ].join("\n");

  const gotA = normalizeRawHighPriorityBetSyntax_(a);
  const expectedA = [
    "🇻🇳 ฮานอย",
    "02 19 36 45 37 81 01 บลก 8*8",
    "11 บล 40*40",
    "ปกติ"
  ].join("\n");

  if (gotA !== expectedA) {
    throw new Error("RAW A FAIL: " + gotA);
  }

  const b = [
    "🇭🇰 ฮั่งเส็งบ่าย",
    "รูด7/5",
    "5*5",
    "ฮั่ง"
  ].join("\n");

  const gotB = normalizeRawHighPriorityBetSyntax_(b);
  const expectedB = [
    "🇭🇰 ฮั่งเส็งบ่าย",
    "รูด 7/5 บล 5*5",
    "ฮั่ง"
  ].join("\n");

  if (gotB !== expectedB) {
    throw new Error("RAW B FAIL: " + gotB);
  }

  return {
    success: true,
    version: LEGACY_BOT_RAW_PATCH_VERSION_,
    case304: gotA,
    caseRood: gotB
  };
}


// ============================================================================
// TEST 2026-11-15 — วินรวมเบิ้ล
// ============================================================================
function testWinCombinedDoubleV1115_() {
  const expected = [
    "33","38","35","37","36","30","32",
    "88","85","87","86","80","82",
    "55","57","56","50","52",
    "77","76","70","72",
    "66","60","62",
    "00","02",
    "22"
  ];

  const generated =
    generateWinCombinedDouble2_("3857602");

  if (
    JSON.stringify(generated) !==
    JSON.stringify(expected)
  ) {
    throw new Error(
      "FAIL numbers: " +
      JSON.stringify(generated)
    );
  }

  const input = [
    "วิน3857602",
    "5*5",
    "บลก",
    "วินรวมเบิ้น"
  ].join("\\n");

  const extracted =
    extractWinCombinedDoubleBets_(input);

  if (
    !extracted.success ||
    extracted.groups.length !== 1
  ) {
    throw new Error("FAIL extractor");
  }

  const g = extracted.groups[0];

  if (g.numbers.length !== 28) {
    throw new Error(
      "FAIL count: " + g.numbers.length
    );
  }

  if (
    Number(g.top) !== 5 ||
    Number(g.bottom) !== 5
  ) {
    throw new Error("FAIL price");
  }

  if (Number(g.subtotal) !== 280) {
    throw new Error(
      "FAIL subtotal: " + g.subtotal
    );
  }

  return {
    success: true,
    count: g.numbers.length,
    totalAmount: g.subtotal,
    numbers: g.numbers
  };
}


// ============================================================================
// TEST 2026-11-15 — วินรวมเบิ้ล + บลก ต้องรวมตัวกลับ = 490 บาท
// ============================================================================
function testWinCombinedDoubleBLG490V1115_() {
  const input = [
    "วิน3857602",
    "5*5",
    "บลก",
    "วินรวมเบิ้น"
  ].join("\n");

  const r = extractWinCombinedDoubleBets_(input);
  if (!r.success || !r.groups || r.groups.length !== 1) {
    throw new Error("FAIL groups");
  }

  const g = r.groups[0];
  if (g.mode !== "บลก") throw new Error("FAIL mode: " + g.mode);
  if (Number(g.betCount) !== 49) throw new Error("FAIL count: " + g.betCount);
  if (Number(g.subtotal) !== 490) throw new Error("FAIL subtotal: " + g.subtotal);

  return {
    success: true,
    count: g.betCount,
    subtotal: g.subtotal,
    numbers: g.numbers
  };
}


// ============================================================================
// TEST 2026-11-15 — เคสจาก LINE จริง วิน3852601 + บลก = 490
// ============================================================================
function testWinCombinedDoubleScreenshotV1115_() {
  const input = [
    "วิน3852601",
    "5*5",
    "บลก",
    "วินรวมเบิ้น"
  ].join("\n");

  const r = extractWinCombinedDoubleBets_(input);
  if (!r.success || !r.groups || r.groups.length !== 1) {
    throw new Error("FAIL groups");
  }

  const g = r.groups[0];
  if (g.mode !== "บลก") throw new Error("FAIL mode: " + g.mode);
  if (Number(g.betCount) !== 49) throw new Error("FAIL count: " + g.betCount);
  if (Number(g.subtotal) !== 490) throw new Error("FAIL subtotal: " + g.subtotal);

  return {
    success: true,
    version: LEGACY_BOT_RAW_PATCH_VERSION_,
    count: g.betCount,
    subtotal: g.subtotal,
    numbers: g.numbers
  };
}


// ============================================================================
// TEST 2026-11-17 — รูปแบบภาษาพูด รูด/บลก
// ============================================================================
function testNaturalRoodModeV1117_() {
  const case1 = [
    "กาชาด บลก",
    "4/8=5×5"
  ].join("\n");

  const got1 =
    normalizeRawHighPriorityBetSyntax_(
      case1
    );

  const expected1 = [
    "กาชาด",
    "รูด 4/8 บลก 5*5"
  ].join("\n");

  if (got1 !== expected1) {
    throw new Error(
      "CASE1 normalize ผิด: " + got1
    );
  }

  const parsed1 =
    parseMessage(
      "รูด 4/8 บลก 5*5",
      [],
      {}
    );

  if (
    !parsed1.success ||
    Number(parsed1.totalAmount) !== 380 ||
    Number(parsed1.totalDoors) !== 38
  ) {
    throw new Error(
      "CASE1 parse ผิด: " +
      JSON.stringify(parsed1)
    );
  }

  const case2 = [
    "ลาวสตาร์",
    "รูด2/3ตัวล่ะ5บาท",
    "บน/ล่าง/กลับตัวล่ะ5บาท",
    "30/31/36/37/38/39"
  ].join("\n");

  const got2 =
    normalizeRawHighPriorityBetSyntax_(
      case2
    );

  const expected2 = [
    "ลาวสตาร์",
    "รูด 2/3 บลก 5*5",
    "30 31 36 37 38 39 บลก 5*5"
  ].join("\n");

  if (got2 !== expected2) {
    throw new Error(
      "CASE2 normalize ผิด: " + got2
    );
  }

  const parsed2 =
    parseMessage(
      [
        "รูด 2/3 บลก 5*5",
        "30 31 36 37 38 39 บลก 5*5"
      ].join("\n"),
      [],
      {}
    );

  if (
    !parsed2.success ||
    Number(parsed2.totalAmount) !== 500 ||
    Number(parsed2.totalDoors) !== 50
  ) {
    throw new Error(
      "CASE2 parse ผิด: " +
      JSON.stringify(parsed2)
    );
  }

  return {
    success: true,
    version:
      LEGACY_BOT_RAW_PATCH_VERSION_,
    case1: {
      normalized: got1,
      totalDoors: parsed1.totalDoors,
      totalAmount: parsed1.totalAmount
    },
    case2: {
      normalized: got2,
      totalDoors: parsed2.totalDoors,
      totalAmount: parsed2.totalAmount
    }
  };
}

;
    return { parseMessage: parseMessage, cleanIncomingLineText_: cleanIncomingLineText_, normalizeRawHighPriorityBetSyntax_: normalizeRawHighPriorityBetSyntax_, normalizeMainPreParseCustomBet_: normalizeMainPreParseCustomBet_, normalizePerLineAssignedBetFormat_: normalizePerLineAssignedBetFormat_, normalizeLegacyNaturalBetFormat_: normalizeLegacyNaturalBetFormat_, stripLottoNameFromText: stripLottoNameFromText, findLottoSafe_: findLottoSafe_, findMultipleLottosInText_: findMultipleLottosInText_, isClosed: isClosed };
  })(stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, stub, Utilities, quiet, sandbox);

  function cleanLines(t) {
    return String(t || "").split(/\n/).map(function (l) { return l.replace(/[ \t]+/g, " ").trim(); }).filter(Boolean).join("\n");
  }

  /* ตัดชื่อหวยแบบยืดหยุ่น: "ลาว extra" / "ลาวExtra" / "🇱🇦ลาวExtra" */
  function stripNameLoose(text, name) {
    var chars = Array.from(String(name).toLowerCase().replace(/[^ก-๙a-z0-9]/gi, ""));
    if (!chars.length) return text;
    var re = new RegExp(chars.map(function (c) { return c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).join("[^ก-๙a-z0-9\\n]*"), "gi");
    return String(text).replace(re, " ");
  }

  function parse(raw, lottos) {
    var data = { map: {}, aliasMap: {} };
    (lottos || []).forEach(function (l) {
      data.map[l.name] = l.closeTime;
      data.aliasMap[l.name] = (l.aliases || []).filter(Boolean);
    });

    var text = L.cleanIncomingLineText_(String(raw || ""));
    text = text.replace(/19\s*ประตู/g, "รูด");          // "19 ประตู" = รูด
    text = text.replace(/(\d)\s*และ\s*(?=\d)/g, "$1 ");  // "1 และ 2" (ตัว ล ใน "และ" ทำให้บอทเดิมหยุดอ่านเลข)

    var multi = L.findMultipleLottosInText_(text, data);
    if (multi && multi.length > 1) {
      return { ok: false, isSlip: true, error: "โพยมีหลายหวยในข้อความเดียว: " + multi.map(function (m) { return m.name; }).join(", ") };
    }

    var m = L.findLottoSafe_(text, data);
    if (!m) return { ok: false, isSlip: false, error: "ไม่พบชื่อหวย" };
    if (m.ambiguous || !m.name) {
      return { ok: false, isSlip: true, error: "ชื่อหวยกำกวม: " + (m.candidates || []).map(function (c) { return c.name; }).join(", ") };
    }

    // ตัดชื่อหวยออกก่อน แล้วค่อยแปลงข้อความ
    // (ตัวแปลงของบอทเดิมเปลี่ยน x เป็น * ทำให้ชื่อเช่น "ลาวExtra" เพี้ยน)
    text = L.stripLottoNameFromText(text, m.name, data, m.matchedToken);
    text = stripNameLoose(text, m.name);
    text = L.normalizeRawHighPriorityBetSyntax_(text);

    var pre = L.normalizeMainPreParseCustomBet_(text, m.name) ||
      L.normalizePerLineAssignedBetFormat_(text, m.name) ||
      L.normalizeLegacyNaturalBetFormat_(text, m.name);
    if (pre) text = pre;

    var cmd = cleanLines(text);
    if (!/\d/.test(cmd)) return { ok: false, isSlip: false, lotto: m.name, error: "ไม่มีเลขในข้อความ" };

    // ไม่ถามกลับเลข: ลงตามที่ลูกค้าพิมพ์
    var r = L.parseMessage(cmd, ["รูด"], {});
    if (!r || !r.success) return { ok: false, isSlip: true, lotto: m.name, error: String((r && r.error) || "อ่านโพยไม่ได้").replace(/^❌\s*/, "") };

    return { ok: true, isSlip: true, lotto: m.name, groups: r.groups, total: r.totalAmount, count: r.totalDoors, cmd: cmd };
  }

  window.LegacyParser = { parse: parse, _L: L, version: "V1117-RUDMODE-NATURAL" };
})(window);
