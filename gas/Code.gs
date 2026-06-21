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
      return handleCustomerConfirm(e.parameter.order || "", e);
    }
    if (action === "staff") {
      return handleStaffPage(e.parameter.order || "", e.parameter.do || "");
    }
    if (action === "api") {
      return handleApi(e.parameter);
    }

    var cache = CacheService.getScriptCache();
    var cached = cache.get("catalog_config");
    if (cached) return _cors(ContentService.createTextOutput(cached));

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

    var jsonOut = JSON.stringify({ catalog: catalog, config: config });
    cache.put("catalog_config", jsonOut, 300);
    return _cors(ContentService.createTextOutput(jsonOut));
  } catch (err) {
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

function handleCustomerConfirm(orderId, e) {
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

    var canConfirm = !custAt && staffAt;
    var doConfirm = canConfirm && (e.parameter.do === "yes");
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
    } else if (canConfirm) {
      html += '<div style="text-align:center;margin-top:20px">';
      html += '<p style="font-family:sans-serif;font-size:14px;color:#555">ได้รับสินค้าเรียบร้อยแล้วใช่ไหมครับ?</p>';
      html += '<a href="?action=confirm&order=' + orderId + '&do=yes" style="display:inline-block;background:#06c755;color:#fff;padding:14px 32px;border-radius:24px;text-decoration:none;font-weight:bold;font-size:16px;margin-top:10px">✅ ยืนยันรับของแล้ว</a></div>';
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
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    var orderId = _genOrderId();
    var ss      = SpreadsheetApp.openById(SHEET_ID);

    var slipStatus = "ไม่มีสลิป";
    var slipNote   = "ลูกค้าไม่ได้แนบสลิป";
    var slipUrl    = "";
    var slipAmount = "";
    var slipTxnId  = "";

    if (data.slipBase64) {
      slipUrl = saveSlipToDrive(data.slipBase64, orderId);
      var verify = verifySlipWithSlipOK(data.slipBase64, data.total);
      var slipokError = verify.error || "";
      if (verify.error) verify = verifySlipWithClaude(data.slipBase64);
      slipAmount = verify.amount || "";
      slipTxnId  = verify.ref || "";

      var isSlipOK = verify.source === "slipok";

      var fallbackInfo = slipokError ? " [SlipOK: " + slipokError + "]" : "";

      if (!verify.amount) {
        slipStatus = "รอตรวจ";
        slipNote   = (verify.error || "อ่านสลิปไม่ได้") + fallbackInfo;
      } else if (!isSlipOK && verify.suspicious) {
        slipStatus = "สงสัยปลอม";
        slipNote   = "Claude: " + (verify.suspicious_reason || "สลิปมีลักษณะผิดปกติ");
      } else if (slipTxnId && isDuplicateSlip(ss, slipTxnId)) {
        slipStatus = "สลิปซ้ำ";
        slipNote   = "เลขอ้างอิง " + slipTxnId + " เคยใช้แล้ว";
      } else if (isSlipOK && (verify.to_account || verify.to_name) && !isCorrectAccount(ss, verify.to_account, verify.to_name)) {
        slipStatus = "บัญชีไม่ตรง";
        slipNote   = "SlipOK: โอนเข้า " + (verify.to_account || "") + " " + (verify.to_name || "") + " ไม่ตรงกับบัญชีร้าน";
      } else if (Number(verify.amount) < Number(data.total)) {
        slipStatus = "ยอดไม่ตรง";
        var src = isSlipOK ? "SlipOK" : "Claude";
        slipNote   = src + ": สลิป " + verify.amount + " บาท แต่ออเดอร์ " + data.total + " บาท" + fallbackInfo;
      } else if (isSlipOK) {
        slipStatus = "ยืนยัน";
        slipNote   = "SlipOK (QR verified): ยอดตรง " + verify.amount + " บาท, " + (verify.bank || "") + " " + (verify.date || "") + " " + (verify.to_name || "");
      } else {
        var cfgWs2 = ss.getSheetByName(TAB_CONFIG);
        var shopAcct = _getConfigValue(cfgWs2, "bank_account") || "";
        var shopNameTh = _getConfigValue(cfgWs2, "bank_account_name") || "";
        var shopNameEn = _getConfigValue(cfgWs2, "bank_account_name_en") || "";

        var amtOk = Number(verify.amount) >= Number(data.total);
        var acctOk = !shopAcct || !verify.to_account || isPartialMatch(verify.to_account, shopAcct);
        var nameOk = !verify.to_name || nameMatch(String(verify.to_name).toLowerCase(), shopNameTh.toLowerCase()) || nameMatch(String(verify.to_name).toLowerCase(), shopNameEn.toLowerCase());

        var details = [];
        details.push("ยอด: " + (amtOk ? "✅ ตรง" : "❌ สลิป " + verify.amount + " ≠ ออเดอร์ " + data.total));
        details.push("บัญชี: " + (acctOk ? "✅ ตรง" : "❌ อ่านได้ " + (verify.to_account || "-") + " ≠ " + shopAcct));
        details.push("ชื่อ: " + (nameOk ? "✅ ตรง" : "❌ อ่านได้ " + (verify.to_name || "-")));

        if (amtOk && acctOk && nameOk) {
          slipStatus = "ยืนยัน";
          slipNote   = "Claude: ตรงทุกรายการ — " + details.join(" | ") + fallbackInfo;
        } else {
          slipStatus = "รอตรวจเพิ่ม";
          slipNote   = "Claude: " + details.join(" | ") + " — admin กรุณาเช็คแอปธนาคาร" + fallbackInfo;
        }
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
    var problemSlip = ["สงสัยปลอม","สลิปซ้ำ","บัญชีไม่ตรง","ยอดไม่ตรง"].indexOf(slipStatus) >= 0;
    if (groupId && problemSlip) {
      _linePush(groupId, "⚠️ ออเดอร์มีปัญหา #" + orderId + "\nสถานะ: " + slipStatus + "\n" + (slipNote || "") + "\n\nตรวจสอบ:\nhttps://waka-liff.vercel.app/staff.html?order=" + orderId);
    }

    if (data.lineUserId) notifyCustomer(data.lineUserId, { orderId: orderId, items: data.items, displayName: data.displayName, branch: data.branch, address: data.address, total: data.total, slipStatus: slipStatus });

    lock.releaseLock();
    return _cors(ContentService.createTextOutput(JSON.stringify({ success: true, orderId: orderId, slipStatus: slipStatus })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
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
    d.orderId, d.timestamp, d.lineUserId, _sanitize(d.displayName),
    d.itemsJson, d.total, d.branch, _sanitize(d.realName), _sanitize(d.phone), _sanitize(d.address),
    d.slipStatus, d.slipUrl, d.slipAmount, d.slipTxnId, d.notes,
  ]);
}

function notifyBranch(groupId, order) {
  var items = (order.items || []).map(function(i) {
    var unitLabel = i.type === "box" ? "กล่อง" : "ซอง";
    return "  - " + i.name + " (" + unitLabel + ") x" + i.qty + " = " + (i.price * i.qty) + " บาท";
  }).join("\n");
  var isDelivery = order.branch === "จัดส่ง";
  var staffUrl = "https://waka-liff.vercel.app/staff.html?order=" + order.orderId;
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
    isDelivery ? "จัดส่งพัสดุ" : "รับที่สาขา: " + order.branch,
  ];
  if (isDelivery && order.address) {
    lines.push("ที่อยู่จัดส่ง: " + order.address);
    lines.push("");
    lines.push("หากที่อยู่ไม่ถูกต้อง กรุณาแจ้งพนักงานหรือแอดมินเพื่อดำเนินการแก้ไขด่วนครับ");
  }
  lines.push("");
  lines.push("ทีมงานจะตรวจสอบและแจ้งกลับทาง LINE ครับ");
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
  var prefix = yy + pad(now.getMonth()+1) + pad(now.getDate());

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(TAB_ORDERS);
  var count = 1;
  if (ws) {
    var rows = ws.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).indexOf(prefix) === 0) {
        var last = parseInt(String(rows[i][0]).slice(6), 10) || 0;
        count = last + 1;
        break;
      }
    }
  }
  return prefix + String(count).padStart(3, "0");
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

function handleApi(params) {
  var action = params.do || "";
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "no orders tab" })));
  var rows = ws.getDataRange().getValues();
  var hdr = rows[0];

  if (action === "search") {
    var q = String(params.q || "").toLowerCase().trim();
    if (!q) return _cors(ContentService.createTextOutput(JSON.stringify({ orders: [] })));
    var results = [];
    for (var i = 1; i < rows.length; i++) {
      var row = {};
      for (var c = 0; c < hdr.length; c++) {
        row[hdr[c]] = rows[i][c] != null ? String(rows[i][c]) : "";
      }
      var match = row.order_id.toLowerCase().indexOf(q) >= 0
        || row.real_name.toLowerCase().indexOf(q) >= 0
        || row.display_name.toLowerCase().indexOf(q) >= 0
        || row.phone.indexOf(q) >= 0;
      if (match) results.push(row);
    }
    results.reverse();
    return _cors(ContentService.createTextOutput(JSON.stringify({ orders: results.slice(0, 20) })));
  }

  if (action === "update") {
    var orderId = params.order || "";
    var newStatus = params.status || "";
    if (!orderId || !newStatus) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "missing params" })));

    var col = function(name) { return hdr.indexOf(name); };
    var gasUrl = ScriptApp.getService().getUrl();
    var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");

    for (var j = 1; j < rows.length; j++) {
      if (String(rows[j][col("order_id")]) !== orderId) continue;
      var branch = rows[j][col("branch")] || "";
      var isDelivery = branch === "จัดส่ง";
      var uid = rows[j][col("line_user_id")] || "";
      var trackUrl = gasUrl + "?action=confirm&order=" + orderId;

      if (newStatus === "shipping") {
        if (col("fulfillment") >= 0) ws.getRange(j+1, col("fulfillment")+1).setValue("กำลังจัดส่งไปสาขา");
        if (col("fulfilled_at") >= 0) ws.getRange(j+1, col("fulfilled_at")+1).setValue(now);
        if (uid) _linePush(uid, "สินค้ากำลังจัดส่งไปสาขา" + branch + "\n\nออเดอร์: #" + orderId + "\n\nดูสถานะ:\n" + trackUrl);
      } else if (newStatus === "ready") {
        if (col("fulfillment") >= 0) ws.getRange(j+1, col("fulfillment")+1).setValue("พร้อมรับ");
        if (col("fulfilled_at") >= 0) ws.getRange(j+1, col("fulfilled_at")+1).setValue(now);
        if (uid) _linePush(uid, "สินค้าพร้อมรับที่สาขา" + branch + " แล้ว!\n\nออเดอร์: #" + orderId + "\n\nดูสถานะ:\n" + trackUrl);
      } else if (newStatus === "handover") {
        var ffVal = isDelivery ? "จัดส่งแล้ว" : "สาขายืนยัน";
        if (col("fulfillment") >= 0) ws.getRange(j+1, col("fulfillment")+1).setValue(ffVal);
        if (col("staff_confirmed_at") >= 0) ws.getRange(j+1, col("staff_confirmed_at")+1).setValue(now);
        if (uid) _linePush(uid, "สาขาส่งมอบสินค้าแล้ว กรุณากดยืนยันรับของ\n\nออเดอร์: #" + orderId + "\n\nกดยืนยัน:\n" + trackUrl);
      }
      return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, status: newStatus, time: now })));
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: "order not found" })));
  }

  return _cors(ContentService.createTextOutput(JSON.stringify({ error: "unknown action" })));
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

function isPartialMatch(slipAcct, shopAcct) {
  var slip = String(slipAcct).replace(/[-\sx]/gi, "");
  var shop = String(shopAcct).replace(/[-\s]/g, "");
  if (!slip || !shop) return true;
  for (var i = 0; i < slip.length; i++) {
    var pos = shop.indexOf(slip[i]);
    if (pos >= 0) {
      var matchCount = 0;
      for (var j = 0; j < slip.length && (pos + j) < shop.length; j++) {
        if (slip[i + j] === shop[pos + j]) matchCount++;
      }
      if (matchCount >= 3) return true;
    }
  }
  return slip.length >= 3 && shop.indexOf(slip) >= 0;
}

function nameMatch(slipName, shopName) {
  if (slipName.indexOf(shopName) >= 0) return true;
  if (shopName.indexOf(slipName) >= 0 && slipName.length >= 8) return true;
  var shorter = slipName.length < shopName.length ? slipName : shopName;
  var longer  = slipName.length < shopName.length ? shopName : slipName;
  return shorter.length >= 8 && longer.indexOf(shorter) === 0;
}

function checkSlipOKQuota() {
  var slipokKey = PROPS.getProperty("SLIPOK_KEY");
  var slipokBranch = PROPS.getProperty("SLIPOK_BRANCH") || "1";
  var res = UrlFetchApp.fetch("https://api.slipok.com/api/line/apikey/" + slipokBranch + "/quota", {
    headers: { "x-authorization": slipokKey }
  });
  Logger.log(res.getContentText());
  return JSON.parse(res.getContentText());
}

function verifySlipWithSlipOK(base64, orderTotal) {
  try {
    var slipokKey = PROPS.getProperty("SLIPOK_KEY");
    var slipokBranch = PROPS.getProperty("SLIPOK_BRANCH") || "1";
    if (!slipokKey) return { error: "ไม่มี SLIPOK_KEY" };

    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, "image/jpeg", "slip.jpg");

    var payload = { files: blob, log: "true" };
    if (orderTotal) payload.amount = String(orderTotal);

    var res = UrlFetchApp.fetch("https://api.slipok.com/api/line/apikey/" + slipokBranch, {
      method: "post",
      muteHttpExceptions: true,
      headers: { "x-authorization": slipokKey },
      payload: payload
    });

    var rawText = res.getContentText();
    var body = JSON.parse(rawText);
    if (!body.success) return { error: "SlipOK: " + (body.message || rawText.substring(0, 200)) };
    if (!body.data) return { error: "SlipOK: no data" };

    var d = body.data;
    if (d.success === false) return { error: "SlipOK: QR ไม่ถูกต้อง - " + (d.message || "") };

    var rcvAcct = (d.receiver && d.receiver.account && d.receiver.account.value) || "";
    var rcvName = (d.receiver && (d.receiver.displayName || d.receiver.name)) || "";
    var sndName = (d.sender && (d.sender.displayName || d.sender.name)) || "";
    var bankCode = d.sendingBank || "";

    return {
      amount: Number(d.amount) || 0,
      date: (d.transDate || "") + " " + (d.transTime || ""),
      bank: bankCode,
      ref: d.transRef || "",
      to_account: rcvAcct,
      to_name: rcvName,
      sender_name: sndName,
      suspicious: false,
      suspicious_reason: "",
      source: "slipok"
    };
  } catch (err) {
    return { error: "SlipOK error: " + err.message };
  }
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
        max_tokens: 500,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: base64 } },
            { type: "text", text: 'อ่านสลิปโอนเงินไทยนี้ ตอบเป็น JSON เท่านั้น ห้ามครอบด้วย markdown: {"amount": 0, "date": "", "bank": "", "ref": "", "to_account": "", "to_name": "", "suspicious": false, "suspicious_reason": ""} โดย to_account=เลขบัญชีปลายทาง, to_name=ชื่อบัญชีปลายทาง หมายเหตุ: สลิปไทยปกติจะซ่อนเลขบัญชีบางส่วนเป็น xxx หรือ * อย่าถือว่าผิดปกติ และรูปสลิปจากมือถืออาจเบลอบ้างเป็นปกติ suspicious=true เฉพาะกรณีชัดเจนว่าตัดต่อเช่น ตัวเลขซ้อนกัน ฟอนต์คนละแบบในจุดเดียวกัน หรือ layout ไม่ตรงกับธนาคารที่ระบุเลย' }
          ]
        }]
      })
    });

    var rawText = res.getContentText();
    var body = JSON.parse(rawText);
    if (body.error) return { error: body.error.message || body.error.type || "API error" };
    var text = (body.content && body.content[0] && body.content[0].text) || "";
    if (!text) return { error: "Claude ไม่ตอบ: " + rawText.substring(0, 200) };
    text = text.replace(/```json\s*/gi, "").replace(/```\s*/g, "").trim();
    var match = text.match(/\{[\s\S]*\}/);
    if (match) {
      try { return JSON.parse(match[0]); }
      catch (pe) { return { error: "JSON parse error: " + match[0].substring(0, 200) }; }
    }
    return { error: "ไม่พบ JSON: " + text.substring(0, 200) };
  } catch (err) {
    return { error: err.message };
  }
}

function _sanitize(val) {
  var s = String(val || "");
  if (s.length > 0 && "=+-@\t\r".indexOf(s[0]) >= 0) s = "'" + s;
  return s;
}

function _driveUrl(url) {
  if (!url) return "";
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w400";
  return url;
}

