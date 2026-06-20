// ──────────────────────────────────────────────────────────────────────────────
// Card Game Order System — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Who has access: "Anyone"
// ──────────────────────────────────────────────────────────────────────────────

const PROPS      = PropertiesService.getScriptProperties();
const LINE_TOKEN = PROPS.getProperty("LINE_TOKEN");
const SHEET_ID   = PROPS.getProperty("SHEET_ID");

const TAB_ORDERS  = "orders";
const TAB_CATALOG = "_catalog";
const TAB_CONFIG  = "_config";

const BRANCH_TABS = {
  "ต้นสัก":      "stock_tonsak",
  "เมืองทอง":   "stock_muangthong",
  "ศรีนครินทร์": "stock_srinakarin",
};

function _cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader("Access-Control-Allow-Origin", "*");
}

// ── GET: โหลด catalog + config สำหรับ LIFF ────────────────────────────────────
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const catWs = ss.getSheetByName(TAB_CATALOG);
    const cfgWs = ss.getSheetByName(TAB_CONFIG);

    // catalog columns: name, category, price_box, price_pack, active, image_url
    const catRows = catWs ? catWs.getDataRange().getValues() : [];
    const catalog = [];
    for (let i = 1; i < catRows.length; i++) {
      const [name, category, price_box, price_pack, active, image_url] = catRows[i];
      if (!name) continue;
      if (active === false || active === "FALSE" || active === 0) continue;
      catalog.push({
        name:      String(name),
        category:  String(category || ""),
        price_box: Number(price_box) || 0,
        price_pack: Number(price_pack) || 0,
        imageUrl:  _driveUrl(String(image_url || "")),
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
    const data    = JSON.parse(e.postData.contents);
    const orderId = _genOrderId();
    const ss      = SpreadsheetApp.openById(SHEET_ID);

    // สลิป — รอตรวจโดย admin (ไม่ใช้ AI)
    const slipStatus = data.slipBase64 ? "รอตรวจ" : "ไม่มีสลิป";
    const slipNote   = data.slipBase64 ? "" : "ลูกค้าไม่ได้แนบสลิป";

    // เขียน order
    writeOrder(ss, {
      orderId,
      timestamp:   new Date().toISOString(),
      lineUserId:  data.lineUserId  || "",
      displayName: data.displayName || "",
      itemsJson:   JSON.stringify(data.items || []),
      total:       data.total       || 0,
      branch:      data.branch      || "",
      realName:    data.realName    || "",
      phone:       data.phone       || "",
      slipStatus,
      slipAmount:  "",
      slipTxnId:   "",
      notes:       slipNote,
    });

    // ตัดสต็อก
    if (data.items && data.branch) {
      deductStock(ss, data.branch, data.items);
    }

    // แจ้ง Line Group สาขา
    const cfgWs   = ss.getSheetByName(TAB_CONFIG);
    const groupId = _getBranchGroupId(cfgWs, data.branch);
    if (groupId) notifyBranch(groupId, { orderId, ...data, slipStatus });

    // ยืนยันกลับลูกค้า
    if (data.lineUserId) notifyCustomer(data.lineUserId, { orderId, ...data, slipStatus });

    return _cors(ContentService.createTextOutput(JSON.stringify({ success: true, orderId, slipStatus })));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })));
  }
}

// ── deductStock: ตัดสต็อกสาขา ────────────────────────────────────────────────
function deductStock(ss, branch, items) {
  const tabName = BRANCH_TABS[branch];
  if (!tabName) return;
  const ws = ss.getSheetByName(tabName);
  if (!ws) return;

  const rows = ws.getDataRange().getValues();
  // columns: name(0), category(1), qty_box(2), qty_pack(3)
  for (const item of items) {
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() !== String(item.name).trim()) continue;
      if (item.type === "box") {
        const cur = Number(rows[r][2]) || 0;
        ws.getRange(r + 1, 3).setValue(Math.max(0, cur - (item.qty || 1)));
      } else {
        const cur = Number(rows[r][3]) || 0;
        ws.getRange(r + 1, 4).setValue(Math.max(0, cur - (item.qty || 1)));
      }
      break;
    }
  }
}

// ── writeOrder ────────────────────────────────────────────────────────────────
function writeOrder(ss, d) {
  let ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) {
    ws = ss.insertSheet(TAB_ORDERS);
    ws.appendRow([
      "order_id","timestamp","line_user_id","display_name",
      "items_json","total","branch","real_name","phone",
      "slip_status","slip_amount","slip_txn_id","notes",
    ]);
  }
  ws.appendRow([
    d.orderId, d.timestamp, d.lineUserId, d.displayName,
    d.itemsJson, d.total, d.branch, d.realName, d.phone,
    d.slipStatus, d.slipAmount, d.slipTxnId, d.notes,
  ]);
}

// ── notifyBranch ──────────────────────────────────────────────────────────────
function notifyBranch(groupId, order) {
  const items = (order.items || []).map(i => {
    const unitLabel = i.type === "box" ? "กล่อง" : "ซอง";
    return `  • ${i.name} (${unitLabel}) ×${i.qty} = ${i.price * i.qty} บาท`;
  }).join("\n");
  const msg = [
    `🛒 ออเดอร์ใหม่ #${order.orderId}`,
    `👤 ${order.displayName}${order.realName ? " (" + order.realName + ")" : ""}`,
    `📦 รับที่: ${order.branch}`,
    ``,
    items,
    ``,
    `💰 ยอดรวม: ${order.total} บาท`,
    `🧾 สลิป: ${order.slipStatus}`,
  ].join("\n");
  _linePush(groupId, msg);
}

// ── notifyCustomer ────────────────────────────────────────────────────────────
function notifyCustomer(userId, order) {
  const items = (order.items || []).map(i => {
    const unitLabel = i.type === "box" ? "กล่อง" : "ซอง";
    return `  • ${i.name} (${unitLabel}) ×${i.qty}`;
  }).join("\n");
  const msg = [
    `⏳ รับออเดอร์แล้ว #${order.orderId}`,
    ``,
    items,
    ``,
    `💰 ยอดรวม: ${order.total} บาท`,
    `📍 รับที่: ${order.branch}`,
    `🧾 สถานะ: ${order.slipStatus}`,
    ``,
    `ทีมงานจะตรวจสอบสลิปและยืนยันอีกครั้งครับ/ค่ะ 🎴`,
  ].join("\n");
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

function _getBranchGroupId(cfgWs, branch) {
  if (!cfgWs || !branch) return null;
  const rows = cfgWs.getDataRange().getValues();
  const key  = "group_" + Object.keys(BRANCH_TABS).find(b => branch.includes(b) || b.includes(branch)) || "";
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key && rows[i][1]) return String(rows[i][1]);
  }
  return null;
}

function _genOrderId() {
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `ORD${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${Math.floor(Math.random()*100)}`;
}

function _driveUrl(url) {
  if (!url) return "";
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return `https://drive.google.com/uc?export=view&id=${m[1]}`;
  return url;
}
