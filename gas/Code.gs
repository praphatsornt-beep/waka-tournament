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
const TAB_STOCK_BRANCH = "stock_branch";
const TAB_SHIPMENTS    = "shipments";

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
      var slug = catRows[i][8] || "";
      var limit_box  = catRows[i][9];
      var limit_pack = catRows[i][10];
      var barcode    = catRows[i][11] || "";
      catalog.push({
        name:       String(name),
        category:   String(category || ""),
        price_box:  Number(price_box)  || 0,
        price_pack: Number(price_pack) || 0,
        imageUrl:   _driveUrl(String(image_url || "")),
        slug:       String(slug),
        limit_box:  (limit_box === "" || limit_box === undefined || limit_box === null) ? -1 : Number(limit_box),
        limit_pack: (limit_pack === "" || limit_pack === undefined || limit_pack === null) ? -1 : Number(limit_pack),
        barcode:    String(barcode),
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
  var url = "https://waka-liff.vercel.app/confirm.html?order=" + encodeURIComponent(orderId || "");
  return HtmlService.createHtmlOutput('<script>window.top.location.href="' + url + '";</script>');
}

// POST: รับ order จาก LIFF หรือ internal actions
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data._action === "sendConfirmLink") {
      return handleSendConfirmLink(data);
    }

    if (data._action === "createShipment") {
      return handleCreateShipment(data);
    }

    if (data._action === "receiveShipment") {
      return handleReceiveShipment(data);
    }

    if (data._action === "handoverOrder") {
      return handleHandoverOrder(data);
    }

    if (data._action === "addStock") {
      return handleAddStock(data);
    }

    if (data._action === "addProduct") {
      return handleAddProduct(data);
    }

    if (data._action === "withdrawStock") {
      return handleWithdrawStock(data);
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

    // ── ตรวจ limit ใน _catalog ก่อนรับออเดอร์ ──
    if (data.items && data.items.length > 0) {
      var limitCheck = checkCatalogLimits(ss, data.items);
      if (limitCheck.error) {
        try { lock.releaseLock(); } catch(_) {}
        return _cors(ContentService.createTextOutput(JSON.stringify({ success: false, error: limitCheck.error })));
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

    if (data.items && data.items.length > 0) {
      deductCatalogLimits(ss, data.items);
      deductStock(ss, data.items);
    }

    // อัปโหลดสลิปหลัง write order — ไม่ block response
    if (data.slipBase64 && !slipUrl) {
      try {
        slipUrl = saveSlipToDrive(data.slipBase64, orderId);
        if (slipUrl) {
          var owsUpd = ss.getSheetByName(TAB_ORDERS);
          var lastRow = owsUpd.getLastRow();
          var hdrUpd = owsUpd.getRange(1, 1, 1, owsUpd.getLastColumn()).getValues()[0];
          var slipUrlCol = hdrUpd.indexOf("slip_url");
          if (slipUrlCol >= 0) owsUpd.getRange(lastRow, slipUrlCol + 1).setValue(slipUrl);
        }
      } catch(_) {}
    }

    lock.releaseLock();

    // LINE push หลัง release lock — ไม่ block order ถัดไป
    try {
      var cfgWs   = ss.getSheetByName(TAB_CONFIG);
      var problemSlip = ["สงสัยปลอม","สลิปซ้ำ","บัญชีไม่ตรง","ยอดไม่ตรง"].indexOf(slipStatus) >= 0;
      if (problemSlip) {
        var financeId = _getConfigValue(cfgWs, "finance_line_id");
        if (financeId) {
          _linePush(financeId, "⚠️ ออเดอร์มีปัญหา #" + orderId + "\nสถานะ: " + slipStatus + "\n" + (slipNote || "") + "\n\nตรวจสอบ:\nhttps://waka-tournament-e6wsqmhuhhexratyiub65f.streamlit.app");
        }
      }
      if (data.lineUserId) notifyCustomer(data.lineUserId, { orderId: orderId, items: data.items, displayName: data.displayName, branch: data.branch, address: data.address, total: data.total, slipStatus: slipStatus });
    } catch(_) {}

    return _cors(ContentService.createTextOutput(JSON.stringify({ success: true, orderId: orderId, slipStatus: slipStatus })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })));
  }
}

// ── Catalog limit: ตรวจจำนวนที่ปล่อยขาย (คอลัมน์ J=limit_box, K=limit_pack) ──
function checkCatalogLimits(ss, items) {
  var ws = ss.getSheetByName(TAB_CATALOG);
  if (!ws) return { ok: true };
  var rows = ws.getDataRange().getValues();
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() !== String(item.name).trim()) continue;
      var colIdx = item.type === "box" ? 9 : 10;
      var limit = rows[r][colIdx];
      if (limit === "" || limit === undefined || limit === null) break; // ไม่จำกัด
      limit = Number(limit);
      if (item.qty > limit) {
        var unitLabel = item.type === "box" ? "กล่อง" : "ซอง";
        if (limit <= 0) return { error: item.name + " (" + unitLabel + ") สินค้าหมดแล้ว" };
        return { error: item.name + " (" + unitLabel + ") เหลือเพียง " + limit + " " + unitLabel };
      }
      break;
    }
  }
  return { ok: true };
}

function deductCatalogLimits(ss, items) {
  var ws = ss.getSheetByName(TAB_CATALOG);
  if (!ws) return;
  var range = ws.getDataRange();
  var rows = range.getValues();
  var changed = false;
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() !== String(item.name).trim()) continue;
      var colIdx = item.type === "box" ? 9 : 10;
      var limit = rows[r][colIdx];
      if (limit === "" || limit === undefined || limit === null) break; // ไม่จำกัด
      rows[r][colIdx] = Math.max(0, Number(limit) - (item.qty || 1));
      changed = true;
      break;
    }
  }
  if (changed) {
    range.setValues(rows);
    CacheService.getScriptCache().remove("catalog_config");
  }
}

function deductStock(ss, items) {
  var ws = ss.getSheetByName(TAB_STOCK);
  if (!ws) return;

  var range = ws.getDataRange();
  var rows = range.getValues();
  var changed = false;
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() !== String(item.name).trim()) continue;
      if (item.type === "box") {
        rows[r][2] = Math.max(0, (Number(rows[r][2]) || 0) - (item.qty || 1));
      } else {
        rows[r][3] = Math.max(0, (Number(rows[r][3]) || 0) - (item.qty || 1));
      }
      changed = true;
      break;
    }
  }
  if (changed) range.setValues(rows);
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

  var propKey = "order_seq_" + prefix;
  var seq = parseInt(PROPS.getProperty(propKey) || "0", 10) + 1;
  PROPS.setProperty(propKey, String(seq));
  return prefix + String(seq).padStart(3, "0");
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
    } else if (action === "ready") {
      if (col("fulfillment") >= 0) ws.getRange(i+1, col("fulfillment")+1).setValue("พร้อมรับ");
      if (col("fulfilled_at") >= 0) ws.getRange(i+1, col("fulfilled_at")+1).setValue(now);
      ff = "พร้อมรับ";
      var uid2 = r[col("line_user_id")];
      if (uid2) {
        var trackUrl2 = "https://waka-liff.vercel.app/confirm.html?order=" + orderId;
        _linePush(uid2, "สินค้าพร้อมรับที่สาขา" + branch + " แล้ว!\n\nออเดอร์: #" + orderId + "\n\nดูสถานะ:\n" + trackUrl2);
      }
    } else if (action === "handover") {
      var ffValue = isDelivery ? "จัดส่งแล้ว" : "สาขายืนยัน";
      if (col("fulfillment") >= 0) ws.getRange(i+1, col("fulfillment")+1).setValue(ffValue);
      if (col("staff_confirmed_at") >= 0) ws.getRange(i+1, col("staff_confirmed_at")+1).setValue(now);
      ff = ffValue;
      var uid3 = r[col("line_user_id")];
      if (uid3) {
        var trackUrl3 = "https://waka-liff.vercel.app/confirm.html?order=" + orderId;
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
      var trackUrl = "https://waka-liff.vercel.app/confirm.html?order=" + orderId;

      if (newStatus === "shipping") {
        if (col("fulfillment") >= 0) ws.getRange(j+1, col("fulfillment")+1).setValue("กำลังจัดส่งไปสาขา");
        if (col("fulfilled_at") >= 0) ws.getRange(j+1, col("fulfilled_at")+1).setValue(now);
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

  if (action === "order_status") {
    var orderId = params.order || "";
    if (!orderId) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "missing order" })));
    var col = function(name) { return hdr.indexOf(name); };
    for (var k = 1; k < rows.length; k++) {
      if (String(rows[k][col("order_id")]) !== orderId) continue;
      var r = rows[k];
      return _cors(ContentService.createTextOutput(JSON.stringify({
        order_id: orderId,
        branch: r[col("branch")] || "",
        slip_status: r[col("slip_status")] || "",
        fulfillment: r[col("fulfillment")] || "",
        staff_confirmed_at: r[col("staff_confirmed_at")] || "",
        customer_confirmed_at: r[col("customer_confirmed_at")] || "",
        timestamp: r[col("timestamp")] || "",
        total: r[col("total")] || 0,
      })));
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: "order not found" })));
  }

  if (action === "customer_confirm") {
    var orderId = params.order || "";
    if (!orderId) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "missing order" })));
    var col = function(name) { return hdr.indexOf(name); };
    for (var m = 1; m < rows.length; m++) {
      if (String(rows[m][col("order_id")]) !== orderId) continue;
      var staffAt = rows[m][col("staff_confirmed_at")] || "";
      var custAt = rows[m][col("customer_confirmed_at")] || "";
      if (custAt) return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, already: true })));
      if (!staffAt) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "staff ยังไม่ส่งมอบ" })));
      var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");
      if (col("customer_confirmed_at") >= 0) ws.getRange(m + 1, col("customer_confirmed_at") + 1).setValue(now);
      if (col("fulfillment") >= 0) ws.getRange(m + 1, col("fulfillment") + 1).setValue("รับแล้ว");
      return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, time: now })));
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: "order not found" })));
  }

  // ── ออเดอร์ของสาขา ──
  if (action === "branch_orders") {
    var branchFilter = params.branch || "";
    if (!branchFilter) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "missing branch" })));
    var col = function(name) { return hdr.indexOf(name); };
    var orders = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][col("branch")] || "") !== branchFilter) continue;
      var slip = String(rows[i][col("slip_status")] || "");
      if (slip !== "ยืนยัน") continue;
      orders.push({
        order_id: String(rows[i][col("order_id")] || ""),
        real_name: String(rows[i][col("real_name")] || ""),
        display_name: String(rows[i][col("display_name")] || ""),
        phone: String(rows[i][col("phone")] || ""),
        items_json: String(rows[i][col("items_json")] || "[]"),
        total: String(rows[i][col("total")] || "0"),
        fulfillment: String(rows[i][col("fulfillment")] || ""),
        staff_confirmed_at: String(rows[i][col("staff_confirmed_at")] || ""),
        customer_confirmed_at: String(rows[i][col("customer_confirmed_at")] || ""),
        timestamp: String(rows[i][col("timestamp")] || ""),
      });
    }
    orders.reverse();
    return _cors(ContentService.createTextOutput(JSON.stringify({ orders: orders })));
  }

  // ── สรุปออเดอร์แต่ละสาขา (รวมเป็นรายสินค้า) ──
  if (action === "branch_summary") {
    var col = function(name) { return hdr.indexOf(name); };
    var summary = {};
    for (var i = 1; i < rows.length; i++) {
      var branch = rows[i][col("branch")] || "";
      var slip = rows[i][col("slip_status")] || "";
      var ff = rows[i][col("fulfillment")] || "";
      if (slip !== "ยืนยัน") continue;
      if (["กำลังจัดส่งไปสาขา","พร้อมรับ","สาขายืนยัน","รับแล้ว","จัดส่งแล้ว"].indexOf(ff) >= 0) continue;
      var items = [];
      try { items = JSON.parse(rows[i][col("items_json")] || "[]"); } catch(e) {}
      if (!summary[branch]) summary[branch] = {};
      for (var x = 0; x < items.length; x++) {
        var key = items[x].name;
        if (!summary[branch][key]) summary[branch][key] = { name: key, qty_box: 0, qty_pack: 0, order_count: 0 };
        if (items[x].type === "box") summary[branch][key].qty_box += (items[x].qty || 1);
        else summary[branch][key].qty_pack += (items[x].qty || 1);
        summary[branch][key].order_count++;
      }
    }
    var result = {};
    for (var b in summary) {
      result[b] = [];
      for (var k in summary[b]) result[b].push(summary[b][k]);
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ branches: result })));
  }

  // ── สต็อกกลาง ──
  if (action === "central_stock") {
    var stockWs = ss.getSheetByName(TAB_STOCK);
    if (!stockWs) return _cors(ContentService.createTextOutput(JSON.stringify({ stock: [] })));
    var sRows = stockWs.getDataRange().getValues();
    var stock = [];
    for (var i = 1; i < sRows.length; i++) {
      if (!sRows[i][0]) continue;
      stock.push({ name: String(sRows[i][0]), category: String(sRows[i][1] || ""), qty_box: Number(sRows[i][2]) || 0, qty_pack: Number(sRows[i][3]) || 0 });
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ stock: stock })));
  }

  // ── สต็อกสาขา ──
  if (action === "branch_stock") {
    var branchFilter = params.branch || "";
    var bsWs = ss.getSheetByName(TAB_STOCK_BRANCH);
    if (!bsWs) return _cors(ContentService.createTextOutput(JSON.stringify({ stock: [] })));
    var bsRows = bsWs.getDataRange().getValues();
    var bStock = [];
    for (var i = 1; i < bsRows.length; i++) {
      if (!bsRows[i][0]) continue;
      if (branchFilter && String(bsRows[i][2]) !== branchFilter) continue;
      bStock.push({ name: String(bsRows[i][0]), category: String(bsRows[i][1] || ""), branch: String(bsRows[i][2] || ""), qty_box: Number(bsRows[i][3]) || 0, qty_pack: Number(bsRows[i][4]) || 0 });
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ stock: bStock })));
  }

  // ── รายการ shipments ──
  if (action === "shipments") {
    var shWs = ss.getSheetByName(TAB_SHIPMENTS);
    if (!shWs) return _cors(ContentService.createTextOutput(JSON.stringify({ shipments: [] })));
    var shRows = shWs.getDataRange().getValues();
    var shList = [];
    for (var i = 1; i < shRows.length; i++) {
      shList.push({
        shipment_id: String(shRows[i][0] || ""),
        timestamp: String(shRows[i][1] || ""),
        to_branch: String(shRows[i][2] || ""),
        status: String(shRows[i][3] || ""),
        items_json: String(shRows[i][4] || "[]"),
        received_at: String(shRows[i][5] || ""),
        notes: String(shRows[i][6] || ""),
      });
    }
    shList.reverse();
    return _cors(ContentService.createTextOutput(JSON.stringify({ shipments: shList })));
  }

  // ── รายงานยอดขาย ──
  if (action === "report") {
    var col = function(name) { return hdr.indexOf(name); };

    // อ่าน cost จาก _catalog
    var catWs = ss.getSheetByName(TAB_CATALOG);
    var costMap = {};
    if (catWs) {
      var catRows = catWs.getDataRange().getValues();
      for (var ci = 1; ci < catRows.length; ci++) {
        if (!catRows[ci][0]) continue;
        costMap[String(catRows[ci][0])] = {
          cost_box: Number(catRows[ci][6]) || 0,
          cost_pack: Number(catRows[ci][7]) || 0,
          price_box: Number(catRows[ci][2]) || 0,
          price_pack: Number(catRows[ci][3]) || 0,
        };
      }
    }

    var byBranch = {};
    var byProduct = {};
    var byDate = {};
    var totalRevenue = 0, totalCost = 0;

    for (var i = 1; i < rows.length; i++) {
      var slip = rows[i][col("slip_status")] || "";
      if (slip !== "ยืนยัน") continue;

      var branch = rows[i][col("branch")] || "ไม่ระบุ";
      var ts = String(rows[i][col("timestamp")] || "");
      var dateKey = ts.substring(0, 10);
      var items = [];
      try { items = JSON.parse(rows[i][col("items_json")] || "[]"); } catch(e) {}

      var orderRev = 0, orderCost = 0;
      for (var x = 0; x < items.length; x++) {
        var it = items[x];
        var qty = it.qty || 1;
        var c = costMap[it.name] || {};
        var rev = (it.price || 0) * qty;
        var cost = (it.type === "box" ? (c.cost_box || 0) : (c.cost_pack || 0)) * qty;
        orderRev += rev;
        orderCost += cost;

        var pKey = it.name + "|" + it.type;
        if (!byProduct[pKey]) byProduct[pKey] = { name: it.name, type: it.type, qty: 0, revenue: 0, cost: 0 };
        byProduct[pKey].qty += qty;
        byProduct[pKey].revenue += rev;
        byProduct[pKey].cost += cost;
      }

      if (!byBranch[branch]) byBranch[branch] = { revenue: 0, cost: 0, orders: 0 };
      byBranch[branch].revenue += orderRev;
      byBranch[branch].cost += orderCost;
      byBranch[branch].orders++;

      if (dateKey) {
        if (!byDate[dateKey]) byDate[dateKey] = { revenue: 0, cost: 0, orders: 0 };
        byDate[dateKey].revenue += orderRev;
        byDate[dateKey].cost += orderCost;
        byDate[dateKey].orders++;
      }

      totalRevenue += orderRev;
      totalCost += orderCost;
    }

    return _cors(ContentService.createTextOutput(JSON.stringify({
      total: { revenue: totalRevenue, cost: totalCost, profit: totalRevenue - totalCost },
      by_branch: byBranch,
      by_product: Object.values(byProduct),
      by_date: byDate,
    })));
  }

  // ── ค้นหาสินค้าจาก barcode ──
  if (action === "lookup_barcode") {
    var barcode = String(params.barcode || "").trim();
    if (!barcode) return _cors(ContentService.createTextOutput(JSON.stringify({ error: "missing barcode" })));
    var catWs = ss.getSheetByName(TAB_CATALOG);
    if (!catWs) return _cors(ContentService.createTextOutput(JSON.stringify({ found: false })));
    var catRows = catWs.getDataRange().getValues();
    for (var i = 1; i < catRows.length; i++) {
      if (String(catRows[i][11] || "").trim() === barcode) {
        var stockWs = ss.getSheetByName(TAB_STOCK);
        var sBox = 0, sPack = 0;
        if (stockWs) {
          var sRows = stockWs.getDataRange().getValues();
          for (var s = 1; s < sRows.length; s++) {
            if (String(sRows[s][0]).trim() === String(catRows[i][0]).trim()) { sBox = Number(sRows[s][2]) || 0; sPack = Number(sRows[s][3]) || 0; break; }
          }
        }
        return _cors(ContentService.createTextOutput(JSON.stringify({
          found: true,
          product: {
            name: String(catRows[i][0]), category: String(catRows[i][1] || ""),
            price_box: Number(catRows[i][2]) || 0, price_pack: Number(catRows[i][3]) || 0,
            cost_box: Number(catRows[i][6]) || 0, cost_pack: Number(catRows[i][7]) || 0,
            barcode: barcode, stock_box: sBox, stock_pack: sPack,
          }
        })));
      }
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ found: false })));
  }

  // ── รายการสินค้าทั้งหมด (สำหรับหน้ารับสต็อก) ──
  if (action === "product_list") {
    var catWs = ss.getSheetByName(TAB_CATALOG);
    if (!catWs) return _cors(ContentService.createTextOutput(JSON.stringify({ products: [] })));
    var catRows = catWs.getDataRange().getValues();
    var stockWs = ss.getSheetByName(TAB_STOCK);
    var stockMap = {};
    if (stockWs) {
      var sRows = stockWs.getDataRange().getValues();
      for (var s = 1; s < sRows.length; s++) {
        if (sRows[s][0]) stockMap[String(sRows[s][0]).trim()] = { qty_box: Number(sRows[s][2]) || 0, qty_pack: Number(sRows[s][3]) || 0 };
      }
    }
    var products = [];
    for (var i = 1; i < catRows.length; i++) {
      if (!catRows[i][0]) continue;
      var n = String(catRows[i][0]).trim();
      var st = stockMap[n] || { qty_box: 0, qty_pack: 0 };
      products.push({
        name: n, category: String(catRows[i][1] || ""),
        price_box: Number(catRows[i][2]) || 0, price_pack: Number(catRows[i][3]) || 0,
        cost_box: Number(catRows[i][6]) || 0, cost_pack: Number(catRows[i][7]) || 0,
        barcode: String(catRows[i][11] || ""),
        limit_box: (catRows[i][9] === "" || catRows[i][9] === undefined || catRows[i][9] === null) ? -1 : Number(catRows[i][9]),
        limit_pack: (catRows[i][10] === "" || catRows[i][10] === undefined || catRows[i][10] === null) ? -1 : Number(catRows[i][10]),
        stock_box: st.qty_box, stock_pack: st.qty_pack,
      });
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({ products: products })));
  }

  // ── Dashboard KPI ──
  if (action === "dashboard") {
    var col = function(name) { return hdr.indexOf(name); };
    var today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd");
    var ordersToday = 0, revenueToday = 0, pendingCount = 0;
    var recentOrders = [];
    for (var i = rows.length - 1; i >= 1; i--) {
      var ts = String(rows[i][col("timestamp")] || "");
      var slip = String(rows[i][col("slip_status")] || "");
      if (ts.substring(0, 10) === today) {
        ordersToday++;
        if (slip === "ยืนยัน") revenueToday += Number(rows[i][col("total")]) || 0;
      }
      if (["รอตรวจ","รอตรวจเพิ่ม","ยอดไม่ตรง","สลิปซ้ำ","บัญชีไม่ตรง","สงสัยปลอม"].indexOf(slip) >= 0) pendingCount++;
      if (recentOrders.length < 10) {
        recentOrders.push({
          order_id: String(rows[i][col("order_id")] || ""),
          real_name: String(rows[i][col("real_name")] || ""),
          total: Number(rows[i][col("total")]) || 0,
          slip_status: slip,
          fulfillment: String(rows[i][col("fulfillment")] || ""),
          branch: String(rows[i][col("branch")] || ""),
          timestamp: ts,
        });
      }
    }
    return _cors(ContentService.createTextOutput(JSON.stringify({
      orders_today: ordersToday, revenue_today: revenueToday,
      pending_count: pendingCount, recent_orders: recentOrders,
    })));
  }

  // ── รายการเบิกสินค้า ──
  if (action === "withdrawals") {
    var wWs = ss.getSheetByName("withdrawals");
    if (!wWs) return _cors(ContentService.createTextOutput(JSON.stringify({ withdrawals: [] })));
    var wRows = wWs.getDataRange().getValues();
    var branchFilter = params.branch || "";
    var wList = [];
    for (var i = 1; i < wRows.length; i++) {
      if (branchFilter && String(wRows[i][1]) !== branchFilter) continue;
      wList.push({ timestamp: String(wRows[i][0] || ""), branch: String(wRows[i][1] || ""), name: String(wRows[i][2] || ""), type: String(wRows[i][3] || ""), qty: Number(wRows[i][4]) || 0, reason: String(wRows[i][5] || "") });
    }
    wList.reverse();
    return _cors(ContentService.createTextOutput(JSON.stringify({ withdrawals: wList.slice(0, 50) })));
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
  if (!ref) return false;
  var cache = CacheService.getScriptCache();
  var cacheKey = "slip_ref_" + String(ref).trim();
  if (cache.get(cacheKey)) return true;

  var ws = ss.getSheetByName(TAB_ORDERS);
  if (!ws) return false;
  var lastRow = ws.getLastRow();
  if (lastRow < 2) return false;
  var hdr = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
  var refCol = hdr.indexOf("slip_txn_id");
  if (refCol < 0) return false;
  var refs = ws.getRange(2, refCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < refs.length; i++) {
    if (String(refs[i][0]).trim() === String(ref).trim()) {
      cache.put(cacheKey, "1", 3600);
      return true;
    }
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

// ── Shipment: สร้างล็อตส่งสาขา ──────────────────────────────────────────────
// data: { to_branch, items: [{name, qty_box, qty_pack, qty_box_extra, qty_pack_extra}] }
function handleCreateShipment(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");
    var shipId = "SH" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMMddHHmm");

    var shWs = ss.getSheetByName(TAB_SHIPMENTS);
    if (!shWs) {
      shWs = ss.insertSheet(TAB_SHIPMENTS);
      shWs.appendRow(["shipment_id", "timestamp", "to_branch", "status", "items_json", "received_at", "notes"]);
    }

    var items = data.items || [];
    // ตัดสต็อกกลาง
    var stockWs = ss.getSheetByName(TAB_STOCK);
    if (stockWs) {
      var sRows = stockWs.getDataRange().getValues();
      for (var idx = 0; idx < items.length; idx++) {
        var it = items[idx];
        var totalBox = (it.qty_box || 0) + (it.qty_box_extra || 0);
        var totalPack = (it.qty_pack || 0) + (it.qty_pack_extra || 0);
        for (var r = 1; r < sRows.length; r++) {
          if (String(sRows[r][0]).trim() !== String(it.name).trim()) continue;
          if (totalBox > 0) {
            var curBox = Number(sRows[r][2]) || 0;
            stockWs.getRange(r + 1, 3).setValue(Math.max(0, curBox - totalBox));
          }
          if (totalPack > 0) {
            var curPack = Number(sRows[r][3]) || 0;
            stockWs.getRange(r + 1, 4).setValue(Math.max(0, curPack - totalPack));
          }
          break;
        }
      }
    }

    shWs.appendRow([shipId, now, data.to_branch || "", "จัดส่ง", JSON.stringify(items), "", ""]);
    lock.releaseLock();

    // LINE แจ้งกลุ่ม staff
    try {
      var cfgWs = ss.getSheetByName(TAB_CONFIG);
      var groupId = _getConfigValue(cfgWs, "group_staff");
      if (groupId) {
        var itemLines = items.map(function(it) {
          var parts = [];
          var tb = (it.qty_box || 0) + (it.qty_box_extra || 0);
          var tp = (it.qty_pack || 0) + (it.qty_pack_extra || 0);
          if (tb > 0) parts.push("Box " + tb + (it.qty_box_extra ? " (เผื่อ " + it.qty_box_extra + ")" : ""));
          if (tp > 0) parts.push("Pack " + tp + (it.qty_pack_extra ? " (เผื่อ " + it.qty_pack_extra + ")" : ""));
          return "  - " + it.name + ": " + parts.join(", ");
        }).join("\n");
        _linePush(groupId, "📦 สร้างล็อตส่งสาขา " + (data.to_branch || "") + "\n\n" + shipId + " — " + now + "\n\n" + itemLines);
      }
    } catch(_) {}

    return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, shipment_id: shipId })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── Shipment: สาขารับของ + เพิ่มสต็อกสาขา + แจ้งลูกค้า ───────────────────
// data: { shipment_id }
function handleReceiveShipment(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var shWs = ss.getSheetByName(TAB_SHIPMENTS);
    if (!shWs) { lock.releaseLock(); return _cors(ContentService.createTextOutput(JSON.stringify({ error: "no shipments tab" }))); }
    var shRows = shWs.getDataRange().getValues();
    var shHdr = shRows[0];
    var shipRow = -1, branch = "", items = [];
    for (var i = 1; i < shRows.length; i++) {
      if (String(shRows[i][0]) === data.shipment_id) {
        shipRow = i;
        branch = String(shRows[i][2]);
        try { items = JSON.parse(shRows[i][4] || "[]"); } catch(e) {}
        break;
      }
    }
    if (shipRow < 0) { lock.releaseLock(); return _cors(ContentService.createTextOutput(JSON.stringify({ error: "shipment not found" }))); }
    if (String(shRows[shipRow][3]) === "รับแล้ว") { lock.releaseLock(); return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, already: true }))); }

    var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");
    shWs.getRange(shipRow + 1, 4).setValue("รับแล้ว");
    shWs.getRange(shipRow + 1, 6).setValue(now);

    // เพิ่มสต็อกสาขา (ทั้งออเดอร์+เผื่อ)
    var bsWs = ss.getSheetByName(TAB_STOCK_BRANCH);
    if (!bsWs) {
      bsWs = ss.insertSheet(TAB_STOCK_BRANCH);
      bsWs.appendRow(["name", "category", "branch", "qty_box", "qty_pack"]);
    }
    var bsRows = bsWs.getDataRange().getValues();
    for (var idx = 0; idx < items.length; idx++) {
      var it = items[idx];
      var addBox = (it.qty_box || 0) + (it.qty_box_extra || 0);
      var addPack = (it.qty_pack || 0) + (it.qty_pack_extra || 0);
      var found = false;
      for (var r = 1; r < bsRows.length; r++) {
        if (String(bsRows[r][0]).trim() === String(it.name).trim() && String(bsRows[r][2]).trim() === branch) {
          if (addBox > 0) bsWs.getRange(r + 1, 4).setValue((Number(bsRows[r][3]) || 0) + addBox);
          if (addPack > 0) bsWs.getRange(r + 1, 5).setValue((Number(bsRows[r][4]) || 0) + addPack);
          found = true;
          break;
        }
      }
      if (!found) {
        bsWs.appendRow([it.name, it.category || "", branch, addBox, addPack]);
        bsRows = bsWs.getDataRange().getValues();
      }
    }

    // แจ้ง LINE ลูกค้าทุกคนที่มีออเดอร์ยืนยัน + สาขานี้ + ยังไม่ส่ง
    var ws = ss.getSheetByName(TAB_ORDERS);
    if (ws) {
      var oRows = ws.getDataRange().getValues();
      var oHdr = oRows[0];
      var oCol = function(name) { return oHdr.indexOf(name); };
      for (var j = 1; j < oRows.length; j++) {
        var oBranch = oRows[j][oCol("branch")] || "";
        var oSlip = oRows[j][oCol("slip_status")] || "";
        var oFf = oRows[j][oCol("fulfillment")] || "";
        if (oBranch !== branch || oSlip !== "ยืนยัน") continue;
        if (["พร้อมรับ","สาขายืนยัน","รับแล้ว"].indexOf(oFf) >= 0) continue;
        // อัปเดต fulfillment เป็น "พร้อมรับ"
        if (oCol("fulfillment") >= 0) ws.getRange(j + 1, oCol("fulfillment") + 1).setValue("พร้อมรับ");
        if (oCol("fulfilled_at") >= 0) ws.getRange(j + 1, oCol("fulfilled_at") + 1).setValue(now);
        var uid = oRows[j][oCol("line_user_id")] || "";
        var oid = String(oRows[j][oCol("order_id")] || "");
        if (uid) {
          var trackUrl = "https://waka-liff.vercel.app/confirm.html?order=" + oid;
          _linePush(uid, "สินค้าพร้อมรับที่สาขา" + branch + " แล้ว!\n\nออเดอร์: #" + oid + "\n\nดูสถานะ:\n" + trackUrl);
        }
      }
    }

    lock.releaseLock();
    return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, time: now })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── ส่งมอบลูกค้า: ตัดสต็อกสาขา ────────────────────────────────────────────
// data: { order_id }
function handleHandoverOrder(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(TAB_ORDERS);
    var oRows = ws.getDataRange().getValues();
    var oHdr = oRows[0];
    var oCol = function(name) { return oHdr.indexOf(name); };
    var now = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm");

    for (var i = 1; i < oRows.length; i++) {
      if (String(oRows[i][oCol("order_id")]) !== data.order_id) continue;
      var branch = oRows[i][oCol("branch")] || "";
      var items = [];
      try { items = JSON.parse(oRows[i][oCol("items_json")] || "[]"); } catch(e) {}

      // ตัดสต็อกสาขา
      var bsWs = ss.getSheetByName(TAB_STOCK_BRANCH);
      if (bsWs) {
        var bsRows = bsWs.getDataRange().getValues();
        for (var idx = 0; idx < items.length; idx++) {
          for (var r = 1; r < bsRows.length; r++) {
            if (String(bsRows[r][0]).trim() !== String(items[idx].name).trim()) continue;
            if (String(bsRows[r][2]).trim() !== branch) continue;
            if (items[idx].type === "box") {
              var curBox = Number(bsRows[r][3]) || 0;
              bsWs.getRange(r + 1, 4).setValue(Math.max(0, curBox - (items[idx].qty || 1)));
            } else {
              var curPack = Number(bsRows[r][4]) || 0;
              bsWs.getRange(r + 1, 5).setValue(Math.max(0, curPack - (items[idx].qty || 1)));
            }
            break;
          }
        }
      }

      // อัปเดต fulfillment
      if (oCol("fulfillment") >= 0) ws.getRange(i + 1, oCol("fulfillment") + 1).setValue("สาขายืนยัน");
      if (oCol("staff_confirmed_at") >= 0) ws.getRange(i + 1, oCol("staff_confirmed_at") + 1).setValue(now);

      // แจ้งลูกค้ากดยืนยันรับ
      var uid = oRows[i][oCol("line_user_id")] || "";
      if (uid) {
        var trackUrl = "https://waka-liff.vercel.app/confirm.html?order=" + data.order_id;
        _linePush(uid, "สาขาส่งมอบสินค้าแล้ว กรุณากดยืนยันรับของ\n\nออเดอร์: #" + data.order_id + "\n\nกดยืนยัน:\n" + trackUrl);
      }

      lock.releaseLock();
      return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true, time: now })));
    }
    lock.releaseLock();
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: "order not found" })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
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

// ── เพิ่มสต็อกสินค้าเดิม ──
// data: { name, add_box, add_pack }
function handleAddStock(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(TAB_STOCK);
    if (!ws) {
      ws = ss.insertSheet(TAB_STOCK);
      ws.appendRow(["name", "category", "qty_box", "qty_pack"]);
    }
    var rows = ws.getDataRange().getValues();
    var found = false;
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]).trim() === String(data.name).trim()) {
        if (data.add_box) ws.getRange(r + 1, 3).setValue((Number(rows[r][2]) || 0) + Number(data.add_box));
        if (data.add_pack) ws.getRange(r + 1, 4).setValue((Number(rows[r][3]) || 0) + Number(data.add_pack));
        found = true;
        break;
      }
    }
    if (!found) {
      ws.appendRow([data.name, data.category || "", Number(data.add_box) || 0, Number(data.add_pack) || 0]);
    }

    // อัปเดต limit ใน _catalog ถ้าส่งมา
    if (data.limit_box !== undefined && data.limit_box !== null || data.limit_pack !== undefined && data.limit_pack !== null) {
      var catWs = ss.getSheetByName(TAB_CATALOG);
      if (catWs) {
        var catRows = catWs.getDataRange().getValues();
        for (var c = 1; c < catRows.length; c++) {
          if (String(catRows[c][0]).trim() === String(data.name).trim()) {
            if (data.limit_box !== undefined && data.limit_box !== null) catWs.getRange(c + 1, 10).setValue(Number(data.limit_box));
            if (data.limit_pack !== undefined && data.limit_pack !== null) catWs.getRange(c + 1, 11).setValue(Number(data.limit_pack));
            CacheService.getScriptCache().remove("catalog_config");
            break;
          }
        }
      }
    }

    lock.releaseLock();
    return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

// ── เพิ่มสินค้าใหม่ใน _catalog + stock ──
// data: { name, category, price_box, price_pack, cost_box, cost_pack, barcode, initial_box, initial_pack }
function handleAddProduct(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var catWs = ss.getSheetByName(TAB_CATALOG);
    if (!catWs) { lock.releaseLock(); return _cors(ContentService.createTextOutput(JSON.stringify({ error: "no _catalog tab" }))); }

    // เช็คชื่อซ้ำ
    var catRows = catWs.getDataRange().getValues();
    for (var i = 1; i < catRows.length; i++) {
      if (String(catRows[i][0]).trim() === String(data.name).trim()) {
        lock.releaseLock();
        return _cors(ContentService.createTextOutput(JSON.stringify({ error: "สินค้าชื่อนี้มีอยู่แล้ว" })));
      }
    }

    // เพิ่มใน _catalog: A=name, B=category, C=price_box, D=price_pack, E=active, F=image_url, G=cost_box, H=cost_pack, I=slug, J=limit_box, K=limit_pack, L=barcode
    var limBox = (data.limit_box === "" || data.limit_box === undefined || data.limit_box === null) ? "" : Number(data.limit_box);
    var limPack = (data.limit_pack === "" || data.limit_pack === undefined || data.limit_pack === null) ? "" : Number(data.limit_pack);
    var newRow = [
      data.name, data.category || "", Number(data.price_box) || 0, Number(data.price_pack) || 0,
      "TRUE", "", Number(data.cost_box) || 0, Number(data.cost_pack) || 0,
      "", limBox, limPack, data.barcode || ""
    ];
    catWs.appendRow(newRow);

    // เพิ่มใน stock
    var stockWs = ss.getSheetByName(TAB_STOCK);
    if (!stockWs) {
      stockWs = ss.insertSheet(TAB_STOCK);
      stockWs.appendRow(["name", "category", "qty_box", "qty_pack"]);
    }
    stockWs.appendRow([data.name, data.category || "", Number(data.initial_box) || 0, Number(data.initial_pack) || 0]);

    CacheService.getScriptCache().remove("catalog_config");
    lock.releaseLock();
    return _cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
  } catch (err) {
    try { lock.releaseLock(); } catch(_) {}
    return _cors(ContentService.createTextOutput(JSON.stringify({ error: err.message })));
  }
}

function clearCache() {
  CacheService.getScriptCache().remove("catalog_config");
}

