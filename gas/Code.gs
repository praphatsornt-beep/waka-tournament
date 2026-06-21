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
const TAB_STOCK   = "stock";

const BRANCHES = ["ต้นสัก", "เมืองทอง", "ศรีนครินทร์"];

function _cors(output) {
  return output.setMimeType(ContentService.MimeType.JSON);
}

// GET: โหลด catalog หรือ ลูกค้ากดยืนยันรับของ
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "";

    if (action === "confirm") {
      return handleCustomerConfirm(e.parameter.order || "");
    }
    if (action === "staff") {
      return handleStaffPage(e.parameter.order || "", e.parameter.do || "");
    }

    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var catWs = ss.getSheetByName(TAB_CATALOG);
    var cfgWs = ss.getSheetByName(TAB_CONFIG);

    var catRows = catWs ? catWs.getDataRange().getValues() : [];
    var catalog = [];
    for (var i = 1; i < catRows.length; i++) {
      var name = catRows[i][0], category = catRows[i][1], price_box = catRows[i][2];
      var price_pack = catRows[i][3], active = catRows[i][4], image_url = catRows[i][5];
      if (!name) continue;
      if (active === false || active === "FALSE" || active === 0) continue;
      catalog.push({
        name:       String(name),
        category:   String(category || ""),
        price_box:  Number(price_box)  || 0,
        price_pack: Number(price_pack) || 0,
        imageUrl:   _driveUrl(String(image_url || "")),
      });
    }

    var cfgRows = cfgWs ? cfgWs.getDataRange().getValues() : [];
    var config  = {};
    for (var j = 1; j < cfgRows.length; j++) {
      if (cfgRows[j][0]) config[String(cfgRows[j][0])] = String(cfgRows[j][1] || "");
    }

    return _cors(ContentService.createTextOutput(JSON.stringify({ catalog: catalog, config: config })));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

function handleCustomerConfirm(orderId) {
  if (!orderId) {
    return HtmlService.createHtmlOutput("<h2>ไม่พบเลขออเดอร์</h2>");
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(TAB_ORDERS);
  var rows = ws.getDataRange().getValues();
  var hdr = rows[0];
  var col = function(name) { return hdr.indexOf(name); };

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][col("order_id")]) !== orderId) continue;

    var r = rows[i];
    var ff       = r[col("fulfillment")] || "รอเตรียม";
    var staffAt  = r[col("staff_confirmed_at")] || "";
    var custAt   = r[col("customer_confirmed_at")] || "";
    var slipSt   = r[col("slip_status")] || "";
    var branch   = r[col("branch")] || "";
    var ts       = r[col("timestamp")] || "";
    var isDelivery = branch === "จัดส่ง";

    var doConfirm = !custAt && staffAt;
    if (doConfirm) {
      var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");
      if (col("customer_confirmed_at") >= 0) ws.getRange(i + 1, col("customer_confirmed_at") + 1).setValue(now);
      if (col("fulfillment") >= 0) ws.getRange(i + 1, col("fulfillment") + 1).setValue("รับแล้ว");
      custAt = now;
      ff = "รับแล้ว";
    }

    var steps = [];
    steps.push({label: "สั่งซื้อสำเร็จ", done: true, time: ts ? ts.substring(0,16).replace("T"," ") : ""});
    steps.push({label: "ยืนยันชำระเงิน", done: slipSt === "ยืนยัน", time: slipSt === "ยืนยัน" ? "ผ่าน" : ""});

    if (isDelivery) {
      steps.push({label: "จัดส่งแล้ว", done: ff === "จัดส่งแล้ว" || ff === "รับแล้ว", time: ""});
    } else {
      steps.push({label: "กำลังจัดส่งไปสาขา", done: ["กำลังจัดส่งไปสาขา","พร้อมรับ","สาขายืนยัน","รับแล้ว"].indexOf(ff) >= 0, time: ""});
      steps.push({label: "พร้อมรับที่สาขา" + branch, done: ["พร้อมรับ","สาขายืนยัน","รับแล้ว"].indexOf(ff) >= 0, time: ""});
      steps.push({label: "สาขาส่งมอบ", done: !!staffAt, time: staffAt});
    }
    steps.push({label: "ลูกค้ายืนยันรับ", done: !!custAt, time: custAt});

    var html = '<div style="max-width:400px;margin:0 auto;padding:20px;font-family:sans-serif">';
    html += '<h2 style="color:#06c755;text-align:center">📦 สถานะออเดอร์</h2>';
    html += '<p style="text-align:center;color:#888">#' + orderId + '</p>';
    html += '<div style="padding:10px 0">';
    for (var s = 0; s < steps.length; s++) {
      var icon = steps[s].done ? "✅" : "⏳";
      var color = steps[s].done ? "#06c755" : "#ccc";
      html += '<div style="display:flex;gap:10px;margin:8px 0;align-items:center">';
      html += '<span style="font-size:20px">' + icon + '</span>';
      html += '<div style="flex:1"><div style="font-weight:bold;color:' + (steps[s].done ? '#333' : '#aaa') + '">' + steps[s].label + '</div>';
      if (steps[s].time) html += '<div style="font-size:12px;color:#999">' + steps[s].time + '</div>';
      html += '</div></div>';
      if (s < steps.length - 1) html += '<div style="margin-left:10px;border-left:2px solid ' + color + ';height:16px"></div>';
    }
    html += '</div>';

    if (doConfirm) {
      html += '<div style="text-align:center;margin-top:20px;padding:16px;background:#f0fbf4;border-radius:10px">';
      html += '<h3 style="color:#06c755">✅ ยืนยันรับของสำเร็จ!</h3>';
      html += '<p style="color:#888">บันทึกเมื่อ: ' + now + '</p></div>';
    } else if (staffAt && !custAt) {
      html += '<div style="text-align:center;margin-top:20px">';
      html += '<p>กดปุ่มด้านล่างเพื่อยืนยันรับของ</p>';
      html += '<a href="?action=confirm&order=' + orderId + '" style="display:inline-block;background:#06c755;color:#fff;padding:14px 32px;border-radius:24px;text-decoration:none;font-weight:bold;font-size:16px;margin-top:10px">✅ ยืนยันรับของ</a></div>';
    }

    html += '<p style="text-align:center;color:#aaa;margin-top:24px;font-size:12px">ขอบคุณที่ใช้บริการ 🎴</p></div>';
    return HtmlService.createHtmlOutput(html);
  }

  return HtmlService.createHtmlOutput("<h2>ไม่พบออเดอร์ #" + orderId + "</h2>");
}

// POST: รับ order จาก LIFF หรือ internal actions
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data._action === "sendConfirmLink") {
      return handleSendConfirmLink(data);
    }

    if (Array.isArray(data.events)) {
      for (var ev = 0; ev < data.events.length; ev++) {
        var evt = data.events[ev];
        var src = evt.source || {};
        var msgText = (evt.message && evt.message.text) || "";
        if (src.type === "group" && src.groupId && msgText.trim() === "!waka-setup") {
          var cfgWs = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_CONFIG);
          if (cfgWs) {
            cfgWs.appendRow(["group_staff", src.groupId]);
            _linePush(src.groupId, "ตั้งค่ากลุ่ม staff สำเร็จ!\nGroup ID: " + src.groupId);
          }
        }
      }
      return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
    }
    const orderId = _genOrderId();
    const ss      = SpreadsheetApp.openById(SHEET_ID);

    var slipStatus = "ไม่มีสลิป";
    var slipNote   = "ลูกค้าไม่ได้แนบสลิป";
    var slipUrl    = "";
    var slipAmount = "";
    var slipTxnId  = "";

    if (data.slipBase64) {
      slipUrl = saveSlipToDrive(data.slipBase64, orderId);
      var verify = verifySlipWithClaude(data.slipBase64);
      slipAmount = verify.amount || "";
      slipTxnId  = verify.ref || "";

      if (!verify.amount) {
        slipStatus = "รอตรวจ";
        slipNote   = verify.error || "อ่านสลิปไม่ได้";
      } else if (verify.suspicious) {
        slipStatus = "สงสัยปลอม";
        slipNote   = "Claude: " + (verify.suspicious_reason || "สลิปมีลักษณะผิดปกติ");
      } else if (slipTxnId && isDuplicateSlip(ss, slipTxnId)) {
        slipStatus = "สลิปซ้ำ";
        slipNote   = "เลขอ้างอิง " + slipTxnId + " เคยใช้แล้ว";
      } else if ((verify.to_account || verify.to_name) && !isCorrectAccount(ss, verify.to_account, verify.to_name)) {
        slipStatus = "บัญชีไม่ตรง";
        slipNote   = "โอนเข้า " + (verify.to_account || "") + " " + (verify.to_name || "") + " ไม่ตรงกับบัญชีร้าน";
      } else if (Number(verify.amount) < Number(data.total)) {
        slipStatus = "ยอดไม่ตรง";
        slipNote   = "Claude: สลิป " + verify.amount + " บาท แต่ออเดอร์ " + data.total + " บาท";
      } else {
        slipStatus = "ยืนยัน";
        slipNote   = "Claude: ยอดตรง " + verify.amount + " บาท, " + (verify.bank || "") + " " + (verify.date || "");
      }
    }

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
      address:     data.address     || "",
      slipStatus,
      slipUrl,
      slipAmount,
      slipTxnId,
      notes:       slipNote,
    });

    if (data.items) {
      deductStock(ss, data.items);
    }

    var cfgWs   = ss.getSheetByName(TAB_CONFIG);
    var groupId = _getConfigValue(cfgWs, "group_staff");
    if (groupId) notifyBranch(groupId, { orderId: orderId, items: data.items, displayName: data.displayName, realName: data.realName, phone: data.phone, branch: data.branch, address: data.address, total: data.total, slipStatus: slipStatus });

    if (data.lineUserId) notifyCustomer(data.lineUserId, { orderId: orderId, items: data.items, displayName: data.displayName, branch: data.branch, total: data.total, slipStatus: slipStatus });

    return _cors(ContentService.createTextOutput(JSON.stringify({ success: true, orderId, slipStatus })));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })));
  }
}

function deductStock(ss, items) {
  var ws = ss.getSheetByName(TAB_STOCK);
  if (!ws) return;

  var rows = ws.getDataRange().getValues();
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() !== String(item.name).trim()) continue;
      if (item.type === "box") {
        var curBox = Number(rows[r][2]) || 0;
        ws.getRange(r + 1, 3).setValue(Math.max(0, curBox - (item.qty || 1)));
      } else {
        var curPack = Number(rows[r][3]) || 0;
        ws.getRange(r + 1, 4).setValue(Math.max(0, curPack - (item.qty || 1)));
      }
      break;
    }
  }
}

function writeOrder(ss, d) {
  let ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) {
    ws = ss.insertSheet(TAB_ORDERS);
    ws.appendRow([
      "order_id","timestamp","line_user_id","display_name",
      "items_json","total","branch","real_name","phone","address",
      "slip_status","slip_url","slip_amount","slip_txn_id","notes",
      "fulfillment","fulfilled_at","staff_confirmed_at","customer_confirmed_at",
    ]);
  }
  ws.appendRow([
    d.orderId, d.timestamp, d.lineUserId, d.displayName,
    d.itemsJson, d.total, d.branch, d.realName, d.phone, d.address,
    d.slipStatus, d.slipUrl, d.slipAmount, d.slipTxnId, d.notes,
  ]);
}

function notifyBranch(groupId, order) {
  var items = (order.items || []).map(function(i) {
    var unitLabel = i.type === "box" ? "กล่อง" : "ซอง";
    return "  - " + i.name + " (" + unitLabel + ") x" + i.qty + " = " + (i.price * i.qty) + " บาท";
  }).join("\n");
  var isDelivery = order.branch === "จัดส่ง";
  var staffUrl = ScriptApp.getService().getUrl() + "?action=staff&order=" + order.orderId;
  var lines = [
    "ออเดอร์ใหม่ #" + order.orderId,
    "ลูกค้า: " + order.displayName + (order.realName ? " (" + order.realName + ")" : ""),
    "โทร: " + order.phone,
    isDelivery ? "จัดส่งพัสดุ" : "รับที่: " + order.branch,
  ];
  if (isDelivery && order.address) lines.push("ที่อยู่: " + order.address);
  lines.push("", items, "", "ยอดรวม: " + order.total + " บาท", "สลิป: " + order.slipStatus, "", "จัดการออเดอร์:", staffUrl);
  _linePush(groupId, lines.join("\n"));
}

function notifyCustomer(userId, order) {
  var items = (order.items || []).map(function(i) {
    var unitLabel = i.type === "box" ? "กล่อง" : "ซอง";
    return "  - " + i.name + " (" + unitLabel + ") x" + i.qty;
  }).join("\n");
  var isDelivery = order.branch === "จัดส่ง";
  var lines = [
    "รับออเดอร์แล้ว #" + order.orderId,
    "",
    items,
    "",
    "ยอดรวม: " + order.total + " บาท",
    isDelivery ? "จัดส่งพัสดุ" : "รับที่: " + order.branch,
    "",
    "ทีมงานจะตรวจสอบและแจ้งกลับทาง LINE ครับ/ค่ะ",
  ];
  _linePush(userId, lines.join("\n"));
}

function _linePush(to, text) {
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization:  "Bearer " + LINE_TOKEN,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({ to, messages: [{ type: "text", text: text }] }),
  });
}

function _getConfigValue(cfgWs, key) {
  if (!cfgWs) return null;
  const rows = cfgWs.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key && rows[i][1]) return String(rows[i][1]);
  }
  return null;
}


function _genOrderId() {
  var now = new Date();
  var pad = function(n) { return String(n).padStart(2, "0"); };
  var yy = String(now.getFullYear()).slice(-2);
  var rand = String(Math.floor(Math.random() * 100)).padStart(2, "0");
  return "WK" + yy + pad(now.getMonth()+1) + pad(now.getDate()) + pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds()) + rand;
}

function saveSlipToDrive(base64, orderId) {
  try {
    const folderId = PROPS.getProperty("SLIP_FOLDER_ID");
    const folder   = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const bytes    = Utilities.base64Decode(base64);
    const blob     = Utilities.newBlob(bytes, "image/jpeg", "slip_" + orderId + ".jpg");
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800";
  } catch (err) {
    return "";
  }
}

function handleStaffPage(orderId, action) {
  if (!orderId) return HtmlService.createHtmlOutput("<h2>ไม่พบออเดอร์</h2>");
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(TAB_ORDERS);
  var rows = ws.getDataRange().getValues();
  var hdr = rows[0];
  var col = function(name) { return hdr.indexOf(name); };
  var gasUrl = ScriptApp.getService().getUrl();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][col("order_id")]) !== orderId) continue;
    var r = rows[i];
    var ff = r[col("fulfillment")] || "รอเตรียม";
    var branch = r[col("branch")] || "";
    var isDelivery = branch === "จัดส่ง";
    var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");

    if (action === "shipping") {
      if (col("fulfillment") >= 0) ws.getRange(i+1, col("fulfillment")+1).setValue("กำลังจัดส่งไปสาขา");
      if (col("fulfilled_at") >= 0) ws.getRange(i+1, col("fulfilled_at")+1).setValue(now);
      ff = "กำลังจัดส่งไปสาขา";
      var uid = r[col("line_user_id")];
      if (uid) {
        var trackUrl = gasUrl + "?action=confirm&order=" + orderId;
        _linePush(uid, "สินค้ากำลังจัดส่งไปสาขา" + branch + "\n\nออเดอร์: #" + orderId + "\n\nดูสถานะ:\n" + trackUrl);
      }
    } else if (action === "ready") {
      if (col("fulfillment") >= 0) ws.getRange(i+1, col("fulfillment")+1).setValue("พร้อมรับ");
      if (col("fulfilled_at") >= 0) ws.getRange(i+1, col("fulfilled_at")+1).setValue(now);
      ff = "พร้อมรับ";
      var uid2 = r[col("line_user_id")];
      if (uid2) {
        var trackUrl2 = gasUrl + "?action=confirm&order=" + orderId;
        _linePush(uid2, "สินค้าพร้อมรับที่สาขา" + branch + " แล้ว!\n\nออเดอร์: #" + orderId + "\n\nดูสถานะ:\n" + trackUrl2);
      }
    } else if (action === "handover") {
      var ffValue = isDelivery ? "จัดส่งแล้ว" : "สาขายืนยัน";
      if (col("fulfillment") >= 0) ws.getRange(i+1, col("fulfillment")+1).setValue(ffValue);
      if (col("staff_confirmed_at") >= 0) ws.getRange(i+1, col("staff_confirmed_at")+1).setValue(now);
      ff = ffValue;
      var uid3 = r[col("line_user_id")];
      if (uid3) {
        var trackUrl3 = gasUrl + "?action=confirm&order=" + orderId;
        _linePush(uid3, "สาขาส่งมอบสินค้าแล้ว กรุณากดยืนยันรับของ\n\nออเดอร์: #" + orderId + "\n\nกดยืนยัน:\n" + trackUrl3);
      }
    }

    var items = [];
    try { items = JSON.parse(r[col("items_json")] || "[]"); } catch(e) {}
    var itemsHtml = "";
    for (var idx = 0; idx < items.length; idx++) {
      var unit = items[idx].type === "box" ? "กล่อง" : "ซอง";
      itemsHtml += "<div>" + items[idx].name + " (" + unit + ") x" + items[idx].qty + "</div>";
    }

    var baseUrl = gasUrl + "?action=staff&order=" + orderId + "&do=";
    var btnStyle = "display:block;width:100%;padding:14px;border:none;border-radius:10px;font-size:16px;font-weight:bold;color:#fff;cursor:pointer;margin:8px 0;text-decoration:none;text-align:center";

    var buttonsHtml = "";
    if (ff === "รอเตรียม" && !isDelivery) {
      buttonsHtml = '<a href="' + baseUrl + 'shipping" style="' + btnStyle + ';background:#2196F3">📤 จัดส่งไปสาขาแล้ว</a>';
    } else if (ff === "กำลังจัดส่งไปสาขา") {
      buttonsHtml = '<a href="' + baseUrl + 'ready" style="' + btnStyle + ';background:#FF9800">📍 ถึงสาขาแล้ว / พร้อมรับ</a>';
    } else if (ff === "พร้อมรับ") {
      buttonsHtml = '<a href="' + baseUrl + 'handover" style="' + btnStyle + ';background:#06c755">🤝 ส่งมอบสินค้าแล้ว</a>';
    } else if (ff === "รอเตรียม" && isDelivery) {
      buttonsHtml = '<a href="' + baseUrl + 'handover" style="' + btnStyle + ';background:#06c755">🚚 จัดส่งพัสดุแล้ว</a>';
    } else if (ff === "สาขายืนยัน" || ff === "รับแล้ว") {
      buttonsHtml = '<div style="text-align:center;padding:16px;background:#f0fbf4;border-radius:10px;color:#06c755;font-weight:bold">✅ ดำเนินการแล้ว</div>';
    }

    if (action) {
      buttonsHtml = '<div style="text-align:center;padding:16px;background:#f0fbf4;border-radius:10px;margin-bottom:12px"><b style="color:#06c755">✅ อัปเดตแล้ว!</b><br><span style="color:#888">' + now + '</span></div>' + buttonsHtml;
    }

    var html = '<div style="max-width:420px;margin:0 auto;padding:20px;font-family:sans-serif">'
      + '<h2 style="text-align:center;color:#333">📋 ออเดอร์ #' + orderId + '</h2>'
      + '<div style="background:#f9f9f9;border-radius:10px;padding:14px;margin:12px 0">'
      + '<div><b>ลูกค้า:</b> ' + (r[col("display_name")] || "") + ' (' + (r[col("real_name")] || "") + ')</div>'
      + '<div><b>โทร:</b> ' + (r[col("phone")] || "") + '</div>'
      + '<div><b>' + (isDelivery ? '🚚 จัดส่งพัสดุ' : '📦 รับที่สาขา: ' + branch) + '</b></div>'
      + (isDelivery && r[col("address")] ? '<div><b>ที่อยู่:</b> ' + r[col("address")] + '</div>' : '')
      + '</div>'
      + '<div style="background:#fff;border:1px solid #eee;border-radius:10px;padding:14px;margin:12px 0">'
      + '<div style="font-weight:bold;margin-bottom:8px">🎴 รายการ</div>' + itemsHtml
      + '<div style="margin-top:8px;font-weight:bold;color:#06c755">ยอดรวม: ' + r[col("total")] + ' บาท</div>'
      + '</div>'
      + '<div style="text-align:center;margin:12px 0;color:#888">สถานะ: <b>' + ff + '</b></div>'
      + buttonsHtml
      + '</div>';

    return HtmlService.createHtmlOutput(html);
  }
  return HtmlService.createHtmlOutput("<h2>ไม่พบออเดอร์ #" + orderId + "</h2>");
}

function handleSendConfirmLink(data) {
  var trackUrl = data.confirmUrl;
  var msg = "";
  if (data.msgType === "ready") {
    msg = "สินค้าพร้อมรับที่สาขาแล้ว!\n\nออเดอร์: #" + data.orderId + "\n\nกดดูสถานะ / ยืนยันรับของ:\n" + trackUrl;
  } else if (data.msgType === "shipped") {
    msg = "สินค้ากำลังจัดส่งไปสาขา\n\nออเดอร์: #" + data.orderId + "\n\nกดดูสถานะ:\n" + trackUrl;
  } else if (data.msgType === "handover") {
    msg = "สาขาส่งมอบสินค้าแล้ว กรุณากดยืนยันรับของ\n\nออเดอร์: #" + data.orderId + "\n\nกดยืนยัน:\n" + trackUrl;
  } else {
    msg = "อัปเดตออเดอร์ #" + data.orderId + "\n\nกดดูสถานะ:\n" + trackUrl;
  }
  _linePush(data.lineUserId, msg);
  return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
}

function isDuplicateSlip(ss, ref) {
  var ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) return false;
  var rows = ws.getDataRange().getValues();
  var hdr  = rows[0];
  var refCol = -1;
  for (var c = 0; c < hdr.length; c++) {
    if (hdr[c] === "slip_txn_id") { refCol = c; break; }
  }
  if (refCol < 0) return false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][refCol]).trim() === String(ref).trim()) return true;
  }
  return false;
}

function isCorrectAccount(ss, toAccount, toName) {
  var cfgWs = ss.getSheetByName(TAB_CONFIG);
  if (!cfgWs) return true;

  var acctOk = true;
  var shopAccount = _getConfigValue(cfgWs, "bank_account");
  if (shopAccount && toAccount) {
    var clean1 = String(toAccount).replace(/[-\s]/g, "");
    var clean2 = String(shopAccount).replace(/[-\s]/g, "");
    acctOk = clean1.indexOf(clean2) >= 0 || clean2.indexOf(clean1) >= 0;
  }

  var nameOk = true;
  if (toName) {
    var shopNameTh = _getConfigValue(cfgWs, "bank_account_name") || "";
    var shopNameEn = _getConfigValue(cfgWs, "bank_account_name_en") || "";
    var slipName = String(toName).toLowerCase().replace(/[.\s]+/g, " ").trim();
    if (shopNameTh || shopNameEn) {
      var matchTh = shopNameTh && nameMatch(slipName, shopNameTh.toLowerCase().trim());
      var matchEn = shopNameEn && nameMatch(slipName, shopNameEn.toLowerCase().trim());
      nameOk = matchTh || matchEn || (!shopNameTh && !shopNameEn);
    }
  }

  return acctOk && nameOk;
}

function nameMatch(slipName, shopName) {
  if (slipName.indexOf(shopName) >= 0) return true;
  if (shopName.indexOf(slipName) >= 0 && slipName.length >= 8) return true;
  var shorter = slipName.length < shopName.length ? slipName : shopName;
  var longer  = slipName.length < shopName.length ? shopName : slipName;
  return shorter.length >= 8 && longer.indexOf(shorter) === 0;
}

function verifySlipWithClaude(base64) {
  try {
    var claudeKey = PROPS.getProperty("CLAUDE_KEY");
    if (!claudeKey) return { error: "ไม่มี CLAUDE_KEY" };

    var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      muteHttpExceptions: true,
      headers: {
        "x-api-key": claudeKey,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 200,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
            { type: "text", text: 'อ่านสลิปโอนเงินนี้ ตอบเป็น JSON เท่านั้น ห้ามตอบอย่างอื่น: {"amount": 0, "date": "", "bank": "", "ref": "", "to_account": "", "to_name": "", "suspicious": false, "suspicious_reason": ""} โดย to_account=เลขบัญชีปลายทาง, to_name=ชื่อบัญชีปลายทาง, suspicious=true ถ้าสงสัยว่าสลิปปลอม เช่น ฟอนต์ผิดปกติ ภาพเบลอเฉพาะจุด มีรอยตัดต่อ ตัวเลขไม่ตรงแนว หรือรูปแบบไม่ตรงกับธนาคารที่ระบุ' }
          ]
        }]
      })
    });

    var body = JSON.parse(res.getContentText());
    if (body.error) return { error: body.error.message };
    var text = (body.content && body.content[0] && body.content[0].text) || "";
    var match = text.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    return { error: "อ่านสลิปไม่ได้" };
  } catch (err) {
    return { error: err.message };
  }
}

function _driveUrl(url) {
  if (!url) return "";
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w400";
  return url;
}

