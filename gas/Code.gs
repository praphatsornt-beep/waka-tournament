// ──────────────────────────────────────────────────────────────────────────────
// Card Game Order System — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Who has access: "Anyone"
// ──────────────────────────────────────────────────────────────────────────────

// ── Secrets (เก็บใน Project Settings → Script Properties) ─────────────────────
const PROPS       = PropertiesService.getScriptProperties();
const CLAUDE_KEY  = PROPS.getProperty("CLAUDE_KEY");
const LINE_TOKEN  = PROPS.getProperty("LINE_TOKEN");
const SHEET_ID    = PROPS.getProperty("SHEET_ID");

// ── Sheet tab names ────────────────────────────────────────────────────────────
const TAB_ORDERS  = "orders";
const TAB_CATALOG = "_catalog";
const TAB_CONFIG  = "_config";

// ── CORS helper ────────────────────────────────────────────────────────────────
function _cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader("Access-Control-Allow-Origin", "*");
}

// ── GET: โหลด catalog + config สำหรับ LIFF ────────────────────────────────────
function doGet(e) {
  try {
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const catWs   = ss.getSheetByName(TAB_CATALOG);
    const cfgWs   = ss.getSheetByName(TAB_CONFIG);

    const catRows = catWs ? catWs.getDataRange().getValues() : [];
    const catalog = [];
    for (let i = 1; i < catRows.length; i++) {
      const [name, price, category, active, imageUrl] = catRows[i];
      if (name && active !== false && active !== "FALSE" && active !== 0)
        catalog.push({
          name:     String(name),
          price:    Number(price) || 0,
          category: String(category || ""),
          imageUrl: _driveUrl(String(imageUrl || "")),
        });
    }

    const cfgRows = cfgWs ? cfgWs.getDataRange().getValues() : [];
    const config  = {};
    for (let i = 1; i < cfgRows.length; i++) {
      const [key, value] = cfgRows[i];
      if (key) config[String(key)] = String(value || "");
    }

    return _cors(ContentService.createTextOutput(JSON.stringify({ catalog, config })));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── POST: รับ order จาก LIFF ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const orderId = _genOrderId();

    // 1. ตรวจสลิป
    let slipStatus = "รอตรวจ";
    let slipAmount = "";
    let slipTxnId  = "";
    let slipNote   = "";

    if (data.slipBase64) {
      const result = verifySlip(data.slipBase64);
      slipAmount   = result.amount || "";
      slipTxnId    = result.txn_id || "";
      const expected = Number(data.total) || 0;
      const actual   = Number(result.amount) || 0;

      if (!result.amount) {
        slipStatus = "อ่านไม่ได้";
        slipNote   = "ระบบอ่านสลิปไม่ได้ แอดมินจะตรวจสอบภายหลัง";
      } else if (Math.abs(actual - expected) <= 1) {
        slipStatus = "✅ ยืนยันแล้ว";
      } else {
        slipStatus = "❌ ยอดไม่ตรง";
        slipNote   = `ยอดในสลิป ${actual} บาท แต่ยอดสั่ง ${expected} บาท`;
      }
    }

    // 2. เขียน Sheets
    const ss = SpreadsheetApp.openById(SHEET_ID);
    writeOrder(ss, {
      orderId, timestamp: new Date().toISOString(),
      lineUserId:  data.lineUserId   || "",
      displayName: data.displayName  || "",
      itemsJson:   JSON.stringify(data.items || []),
      total:       data.total        || 0,
      realName:    data.realName     || "",
      nickname:    data.nickname     || "",
      pickup:      data.pickup       || "",
      slipStatus, slipAmount, slipTxnId,
      notes:       slipNote,
    });

    // 3. แจ้ง Line Group สาขา
    const cfgWs   = ss.getSheetByName(TAB_CONFIG);
    const groupId = _getBranchGroupId(cfgWs, data.pickup);
    if (groupId) notifyBranch(groupId, { orderId, ...data, slipStatus });

    // 4. ยืนยันกลับลูกค้า
    if (data.lineUserId) notifyCustomer(data.lineUserId, { orderId, ...data, slipStatus, slipNote });

    return _cors(ContentService.createTextOutput(JSON.stringify({ success: true, orderId, slipStatus })));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })));
  }
}

// ── verifySlip: เรียก Claude Haiku อ่านสลิป ──────────────────────────────────
function verifySlip(base64) {
  try {
    const resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        "x-api-key":         CLAUDE_KEY,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json",
      },
      payload: JSON.stringify({
        model:      "claude-haiku-4-5-20251001",
        max_tokens: 200,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 }},
            { type: "text",  text: 'อ่านสลิปโอนเงินนี้ ตอบเป็น JSON เท่านั้น ไม่ต้องอธิบาย: {"amount": 0, "date": "", "bank": "", "txn_id": ""}' },
          ],
        }],
      }),
    });

    const body = JSON.parse(resp.getContentText());
    const text = body.content?.[0]?.text || "{}";
    const json = text.match(/\{[\s\S]*\}/)?.[0] || "{}";
    return JSON.parse(json);
  } catch (_) {
    return {};
  }
}

// ── writeOrder: เขียน row ใหม่ใน orders tab ─────────────────────────────────
function writeOrder(ss, d) {
  let ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) {
    ws = ss.insertSheet(TAB_ORDERS);
    ws.appendRow([
      "order_id","timestamp","line_user_id","display_name",
      "items_json","total","real_name","nickname","pickup",
      "slip_status","slip_amount","slip_txn_id","notes",
    ]);
  }
  ws.appendRow([
    d.orderId, d.timestamp, d.lineUserId, d.displayName,
    d.itemsJson, d.total, d.realName, d.nickname, d.pickup,
    d.slipStatus, d.slipAmount, d.slipTxnId, d.notes,
  ]);
}

// ── notifyBranch: ส่งข้อความเข้า Line Group สาขา ────────────────────────────
function notifyBranch(groupId, order) {
  const items = (order.items || []).map(i => `  • ${i.name} ×${i.qty} = ${i.price * i.qty} บาท`).join("\n");
  const msg = [
    `🛒 ออเดอร์ใหม่ #${order.orderId}`,
    `👤 ${order.displayName} (${order.nickname || order.realName})`,
    `📦 ${order.pickup}`,
    ``,
    items,
    ``,
    `💰 ยอดรวม: ${order.total} บาท`,
    `🧾 สลิป: ${order.slipStatus}`,
  ].join("\n");
  _linePush(groupId, msg);
}

// ── notifyCustomer: ส่งยืนยันกลับลูกค้า ────────────────────────────────────
function notifyCustomer(userId, order) {
  const statusIcon = order.slipStatus?.includes("✅") ? "✅" : "⚠️";
  const items = (order.items || []).map(i => `  • ${i.name} ×${i.qty}`).join("\n");
  const msg = [
    `${statusIcon} ยืนยันการสั่งซื้อ #${order.orderId}`,
    ``,
    items,
    ``,
    `💰 ยอดรวม: ${order.total} บาท`,
    `📍 รับที่: ${order.pickup}`,
    `🧾 สถานะสลิป: ${order.slipStatus}`,
    order.notes ? `\n⚠️ ${order.notes}` : "",
    ``,
    `ขอบคุณที่สั่งซื้อครับ/ค่ะ 🎴`,
  ].filter(l => l !== undefined).join("\n");
  _linePush(userId, msg);
}

// ── helpers ───────────────────────────────────────────────────────────────────
function _linePush(to, text) {
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization:  "Bearer " + LINE_TOKEN,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({ to, messages: [{ type: "text", text }] }),
  });
}

function _getBranchGroupId(cfgWs, pickup) {
  if (!cfgWs || !pickup) return null;
  const rows = cfgWs.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [key, value] = rows[i];
    if (String(key).includes("group_") && pickup.includes(String(value || "").replace("group_","")))
      return value;
    // key pattern: "group_ต้นสัก" → value = "C123456789"
    if (String(key).startsWith("group_") && pickup.includes(String(key).replace("group_", "")))
      return String(value);
  }
  return null;
}

function _genOrderId() {
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `ORD${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${Math.floor(Math.random()*100)}`;
}

// แปลง Google Drive share URL → direct image URL
// รองรับทั้ง /file/d/ID/view และ URL ตรงอื่นๆ
function _driveUrl(url) {
  if (!url) return "";
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=view&id=${m[1]}`;
  return url; // ถ้าเป็น URL ตรงอื่นให้ใช้ได้เลย
}
