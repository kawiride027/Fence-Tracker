// ══════════════════════════════════════════════════════════════════════
// Google Apps Script - Fence Inventory Backend v7
// ══════════════════════════════════════════════════════════════════════
//
// TABS CREATED:
//   1. "📊 Yard Inventory" — LIVE dashboard (what you have right now)
//   2. "Inventory Log"     — every job entry
//   3. "Extras Log"        — extras needing invoicing
//   4. "Sales Log"         — all material sales
//   5. "Damage Log"        — damaged material pulled from the yard   ← NEW in v7
//
// WHAT CHANGED FROM v6 (damage reporting):
//   • doPost routes a new "email_damaged" action
//   • updateYardInventory has a new "damaged" direction (subtracts from
//     in-yard New/Used; does NOT touch Out on Rent)
//   • handleInventory recognizes jobType "damaged", and now takes a script
//     lock so the pickup + damaged submissions can't race on the yard math
//   • handleDamagedEmail logs to Damage Log and emails the sales rep
//
// SETUP / REDEPLOY:
//   Paste over the old code → Save → Deploy → Manage deployments →
//   (pencil/Edit) → Version: New version → Deploy.
//   (No need to re-run initYardInventory unless you want to reset.)
//
// ══════════════════════════════════════════════════════════════════════

var SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

var ITEMS = [
  "6×10 Fence Panel",
  "8×10 Fence Panel",
  "6×12 Fence Panel",
  "8×12 Fence Panel",
  "6×10 Panel w/ Ped Gate",
  "8×10 Panel w/ Ped Gate",
  "Fence Posts",
  "Fence Bases / Stands",
  "Fence Clamps",
  "Fence Wheels",
  "Green 6ft Windscreen",
  "Green 8ft Windscreen",
  "Black 6ft Windscreen",
  "Black 8ft Windscreen",
  "Pedestrian Gates",
  "Barricades",
  "Sandbags"
];

// ══════════════════════════════════════════════════════════════════════
// RUN THIS ONCE: Creates the Yard Inventory dashboard
// Select initYardInventory from dropdown → click Run
// ══════════════════════════════════════════════════════════════════════
function initYardInventory() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var name = "📊 Yard Inventory";
  var sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(name, 0);

  var headers = [
    "Item",
    "New (In Yard)",
    "Used (In Yard)",
    "Out on Rent",
    "Total Sold",
    "Last Updated"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1a5632")
    .setFontColor("#ffffff")
    .setFontSize(11)
    .setHorizontalAlignment("center");

  // Item name column left-aligned
  sheet.getRange(1, 1).setHorizontalAlignment("left");

  for (var i = 0; i < ITEMS.length; i++) {
    var row = i + 2;
    sheet.getRange(row, 1).setValue(ITEMS[i]);
    sheet.getRange(row, 2).setValue(0); // New (In Yard)
    sheet.getRange(row, 3).setValue(0); // Used (In Yard)
    sheet.getRange(row, 4).setValue(0); // Out on Rent
    sheet.getRange(row, 5).setValue(0); // Total Sold
    sheet.getRange(row, 6).setValue(""); // Last Updated
  }

  // Column widths
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 160);

  // Color coding
  var numItems = ITEMS.length;
  sheet.getRange(2, 1, numItems, 1).setFontWeight("bold");
  sheet.getRange(2, 2, numItems, 1).setBackground("#dcfce7").setHorizontalAlignment("center"); // New = green
  sheet.getRange(2, 3, numItems, 1).setBackground("#dbeafe").setHorizontalAlignment("center"); // Used = blue
  sheet.getRange(2, 4, numItems, 1).setBackground("#fce7f3").setHorizontalAlignment("center"); // Out on Rent = pink
  sheet.getRange(2, 5, numItems, 1).setBackground("#fed7aa").setHorizontalAlignment("center"); // Sold = orange
  sheet.getRange(2, 6, numItems, 1).setFontColor("#888888").setFontSize(10);

  // Instructions
  var r = numItems + 3;
  sheet.getRange(r, 1).setValue("📋 HOW THIS WORKS:").setFontWeight("bold").setFontSize(12);
  sheet.getRange(r+1, 1).setValue("1. Enter your starting yard counts in columns B (New) and C (Used)");
  sheet.getRange(r+2, 1).setValue("2. When drivers submit jobs through the app, these numbers update automatically");
  sheet.getRange(r+3, 1).setValue("3. Deliveries subtract from yard → Pickups add back to yard");
  sheet.getRange(r+4, 1).setValue("4. Sales subtract from New permanently");
  sheet.getRange(r+5, 1).setValue("5. Out on Rent shows how many used items are at customer job sites right now");
  sheet.getRange(r+6, 1).setValue("6. Damaged items are pulled from yard stock and logged in the Damage Log tab");

  // Also init Inventory Log
  getOrCreateSheet("Inventory Log", [
    "Submitted", "Job Date", "Driver", "Job Type", "Direction",
    "Customer", "Job Site", "PO#",
    "Item", "Qty", "Unit", "Condition",
    "Payment Received", "Payment Type", "Payment Notes", "Notes"
  ]);

  SpreadsheetApp.flush();
  Logger.log("✅ Yard Inventory created! Enter your starting counts in columns B and C.");
}

// ══════════════════════════════════════════════════════════════════════
// WEB APP
// ══════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === "inventory") return handleInventory(data);
    if (data.action === "email_extras") return handleExtrasEmail(data);
    if (data.action === "email_sale") return handleSaleEmail(data);
    if (data.action === "email_damaged") return handleDamagedEmail(data); // ← NEW
    return jsonResponse({ status: "error", message: "Unknown action" });
  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "Fence Inventory API v7" });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f0f0f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function fmtDate(d) {
  if (!d) return "";
  var p = d.split("-");
  return p.length === 3 ? p[1]+"/"+p[2]+"/"+p[0] : d;
}

// ── Update Yard Inventory ────────────────────────────────────────────
// Columns: A=Item, B=New(InYard), C=Used(InYard), D=OutOnRent, E=TotalSold, F=LastUpdated
function updateYardInventory(itemName, qty, condition, direction) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("📊 Yard Inventory");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === itemName) { row = i + 1; break; }
  }
  if (row === -1) return;

  var isNew = (condition || "").toLowerCase() === "new";
  var now = new Date().toLocaleString();

  if (direction === "out") {
    // Item leaves yard (delivery, swap drop-off, extras)
    if (isNew) {
      var curNew = sheet.getRange(row, 2).getValue() || 0;
      sheet.getRange(row, 2).setValue(Math.max(0, curNew - qty));
    } else {
      var curUsed = sheet.getRange(row, 3).getValue() || 0;
      sheet.getRange(row, 3).setValue(Math.max(0, curUsed - qty));
    }
    // Track out on rent for ALL deliveries (new or used)
    var outRent = sheet.getRange(row, 4).getValue() || 0;
    sheet.getRange(row, 4).setValue(outRent + qty);
  } else if (direction === "in") {
    // Item returns to yard (pickup, swap pick-up)
    if (isNew) {
      var curNew2 = sheet.getRange(row, 2).getValue() || 0;
      sheet.getRange(row, 2).setValue(curNew2 + qty);
    } else {
      var curUsed2 = sheet.getRange(row, 3).getValue() || 0;
      sheet.getRange(row, 3).setValue(curUsed2 + qty);
    }
    // Reduce out on rent for ALL pickups
    var outRent2 = sheet.getRange(row, 4).getValue() || 0;
    sheet.getRange(row, 4).setValue(Math.max(0, outRent2 - qty));
  } else if (direction === "sold") {
    // Sale: new item gone permanently
    var curNew3 = sheet.getRange(row, 2).getValue() || 0;
    sheet.getRange(row, 2).setValue(Math.max(0, curNew3 - qty));
    var sold = sheet.getRange(row, 5).getValue() || 0;
    sheet.getRange(row, 5).setValue(sold + qty);
  } else if (direction === "damaged") {
    // ← NEW: Damaged material pulled from the yard. On a pickup the driver
    // enters everything returned (which the pickup adds back and reduces
    // Out on Rent), then flags the damaged subset here — so we ONLY subtract
    // it from in-yard stock and leave Out on Rent alone.
    if (isNew) {
      var curNewD = sheet.getRange(row, 2).getValue() || 0;
      sheet.getRange(row, 2).setValue(Math.max(0, curNewD - qty));
    } else {
      var curUsedD = sheet.getRange(row, 3).getValue() || 0;
      sheet.getRange(row, 3).setValue(Math.max(0, curUsedD - qty));
    }
  }

  sheet.getRange(row, 6).setValue(now);
}

// ── Handle Inventory Submission ──────────────────────────────────────
function handleInventory(data) {
  // Take a lock: a pickup-with-damage now sends the pickup and the damaged
  // rows as two separate requests that fire at the same time. Without this,
  // their read-modify-write on the same yard row can lose an update.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (lockErr) { /* proceed rather than drop the row */ }

  try {
    var HEADERS = [
      "Submitted", "Job Date", "Driver", "Job Type", "Direction",
      "Customer", "Job Site", "PO#",
      "Item", "Qty", "Unit", "Condition",
      "Payment Received", "Payment Type", "Payment Notes", "Notes"
    ];
    var sheet = getOrCreateSheet("Inventory Log", HEADERS);
    var jobDate = fmtDate(data.jobDate);
    var isSale = data.jobType === "sale";
    var isDamaged = data.jobType === "damaged";
    var payR = data.paymentReceived === true ? "YES" : data.paymentReceived === false ? "NO" : "";
    var payT = data.paymentType || "";
    var payN = data.paymentNotes || "";
    var rc = 0;

    // Standard items
    var items = data.items || [];
    items.forEach(function(item) {
      var dir = isSale ? "Sold"
        : isDamaged ? "🛠️ Damaged"
        : data.jobType === "pickup" ? "Picked Up"
        : "Delivered";
      sheet.appendRow([
        data.timestamp, jobDate, data.driver, data.jobType, dir,
        data.customer, data.jobSite, data.poNumber,
        item.item, item.quantity, item.unit,
        (item.condition || "used").toUpperCase(),
        isSale ? payR : "", isSale ? payT : "", isSale ? payN : "",
        data.notes || ""
      ]);
      if (isSale) {
        updateYardInventory(item.item, item.quantity, item.condition, "sold");
      } else if (data.jobType === "delivery") {
        updateYardInventory(item.item, item.quantity, item.condition, "out");
      } else if (data.jobType === "pickup") {
        updateYardInventory(item.item, item.quantity, item.condition, "in");
      } else if (isDamaged) {
        updateYardInventory(item.item, item.quantity, item.condition, "damaged");
      }
      rc++;
    });

    // Swap OUT (picked up from customer → returns to yard)
    var swapOut = data.swapOut || [];
    swapOut.forEach(function(item) {
      sheet.appendRow([
        data.timestamp, jobDate, data.driver, data.jobType, "Swap - Picked Up",
        data.customer, data.jobSite, data.poNumber,
        item.item, item.quantity, item.unit,
        (item.condition || "used").toUpperCase(),
        "", "", "", data.notes || ""
      ]);
      updateYardInventory(item.item, item.quantity, item.condition, "in");
      rc++;
    });

    // Swap IN (dropped off to customer → leaves yard)
    var swapIn = data.swapIn || [];
    swapIn.forEach(function(item) {
      sheet.appendRow([
        data.timestamp, jobDate, data.driver, data.jobType, "Swap - Dropped Off",
        data.customer, data.jobSite, data.poNumber,
        item.item, item.quantity, item.unit,
        (item.condition || "used").toUpperCase(),
        "", "", "", data.notes || ""
      ]);
      updateYardInventory(item.item, item.quantity, item.condition, "out");
      rc++;
    });

    // Extras (leave yard)
    var extras = data.extras || [];
    extras.forEach(function(item) {
      sheet.appendRow([
        data.timestamp, jobDate, data.driver, data.jobType, "⚠️ EXTRA",
        data.customer, data.jobSite, data.poNumber,
        item.item, item.quantity, item.unit,
        (item.condition || "used").toUpperCase(),
        "", "", "", data.extraNotes || ""
      ]);
      updateYardInventory(item.item, item.quantity, item.condition, "out");
      rc++;
    });

    return jsonResponse({ status: "success", rows: rc });
  } finally {
    try { lock.releaseLock(); } catch (relErr) {}
  }
}

// ── Handle Extras Email ──────────────────────────────────────────────
function handleExtrasEmail(data) {
  var salesEmail = data.salesEmail;
  if (!salesEmail) return jsonResponse({ status: "error", message: "No sales email" });

  var HEADERS = ["Submitted","Job Date","Driver","Customer","Job Site","PO#","Item","Qty","Unit","Condition","Extra Notes","Email Sent To"];
  var extrasSheet = getOrCreateSheet("Extras Log", HEADERS);
  var jobDate = fmtDate(data.jobDate);

  var extras = data.extras || [];
  extras.forEach(function(item) {
    extrasSheet.appendRow([data.timestamp, jobDate, data.driver, data.customer, data.jobSite, data.poNumber, item.item, item.quantity, item.unit, (item.condition||"used").toUpperCase(), data.extraNotes||"", salesEmail]);
  });

  var list = extras.map(function(i){return "  • "+i.quantity+" "+i.unit+" of "+i.item+" ("+(i.condition||"used").toUpperCase()+")";}).join("\n");
  var subject = "⚠️ EXTRA FENCING - " + data.customer + " - Charge Required";
  var body = ["EXTRA MATERIALS ALERT","═══════════════════","","Customer: "+data.customer,"Job Date: "+jobDate,"Job Site: "+(data.jobSite||"N/A"),"PO#: "+(data.poNumber||"N/A"),"Driver: "+data.driver,"","EXTRAS:","───────────────────",list,"","Notes: "+(data.extraNotes||"None"),"","ACTION: Invoice the customer.","","— TGD Fence Tracker"].join("\n");

  var eHtml = extras.map(function(i){var c=i.condition==="new"?"#16a34a":"#64748b";return '<tr><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;">'+i.item+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;font-weight:bold;">'+i.quantity+' '+i.unit+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;color:'+c+';font-weight:bold;">'+(i.condition||"used").toUpperCase()+'</td></tr>';}).join('');

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;"><div style="background:#f59e0b;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;"><h2 style="margin:0;">⚠️ Extra Materials</h2></div><div style="background:#fff;border:1px solid #e5e7eb;padding:24px;border-radius:0 0 8px 8px;"><table style="width:100%;border-collapse:collapse;margin-bottom:16px;"><tr><td style="padding:8px 0;color:#6b7280;width:100px;">Customer:</td><td style="font-weight:bold;">'+data.customer+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Date:</td><td style="font-weight:bold;">'+jobDate+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Site:</td><td>'+(data.jobSite||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">PO #:</td><td>'+(data.poNumber||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Driver:</td><td>'+data.driver+'</td></tr></table><div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:8px;padding:16px;margin-bottom:16px;"><h3 style="margin:0 0 12px;color:#92400e;">Extra Materials:</h3><table style="width:100%;border-collapse:collapse;"><tr style="background:#fef3c7;"><th style="padding:6px 12px;text-align:left;color:#92400e;">Item</th><th style="padding:6px 12px;text-align:left;color:#92400e;">Qty</th><th style="padding:6px 12px;text-align:left;color:#92400e;">Cond</th></tr>'+eHtml+'</table></div>'+(data.extraNotes?'<p style="color:#6b7280;"><strong>Notes:</strong> '+data.extraNotes+'</p>':'')+'<div style="background:#fee2e2;border:1px solid #ef4444;border-radius:8px;padding:12px;text-align:center;"><strong style="color:#dc2626;">ACTION REQUIRED</strong><p style="margin:4px 0 0;color:#991b1b;">Invoice customer for extras.</p></div></div></div>';

  GmailApp.sendEmail(salesEmail, subject, body, {htmlBody:html, name:"TGD Fence Tracker"});
  return jsonResponse({ status: "success", emailSent: true });
}

// ── Handle Sale Email ────────────────────────────────────────────────
function handleSaleEmail(data) {
  var salesEmail = data.salesEmail;
  if (!salesEmail) return jsonResponse({ status: "error", message: "No sales email" });

  var HEADERS = ["Submitted","Job Date","Driver","Customer","Job Site","PO#","Item","Qty","Unit","Payment Received","Payment Type","Payment Notes","Email Sent To"];
  var salesSheet = getOrCreateSheet("Sales Log", HEADERS);
  var jobDate = fmtDate(data.jobDate);

  var items = data.items || [];
  items.forEach(function(item) {
    salesSheet.appendRow([data.timestamp, jobDate, data.driver, data.customer, data.jobSite, data.poNumber, item.item, item.quantity, item.unit, data.paymentReceived?"YES":"NO", data.paymentType||"", data.paymentNotes||"", salesEmail]);
  });

  var types = {cc:"Credit Card",check:"Check",cash:"Cash",other:"Other"};
  var payLabel = data.paymentReceived ? (types[data.paymentType]||data.paymentType||"Unknown") : "";
  var itemsList = items.map(function(i){return "  • "+i.quantity+" "+i.unit+" of "+i.item;}).join("\n");
  var urg = data.paymentReceived ? "" : " - ⚠️ UNPAID";
  var subject = "💰 MATERIAL SALE - " + data.customer + urg;

  var body = ["MATERIAL SALE","═══════════════════","","Customer: "+data.customer,"Job Date: "+jobDate,"Job Site: "+(data.jobSite||"N/A"),"PO#: "+(data.poNumber||"N/A"),"Driver: "+data.driver,"","SOLD (NEW):","───────────────────",itemsList,"","PAYMENT:","───────────────────","Received: "+(data.paymentReceived?"YES":"NO"),data.paymentReceived?"Type: "+payLabel:"",data.paymentNotes?"Notes: "+data.paymentNotes:"","",data.paymentReceived?"✅ Paid.":"⚠️ NOT paid — follow up.","","— TGD Fence Tracker"].filter(function(l){return l!=="";}).join("\n");

  var iHtml = items.map(function(i){return '<tr><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;">'+i.item+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;font-weight:bold;">'+i.quantity+' '+i.unit+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;color:#16a34a;font-weight:bold;">NEW</td></tr>';}).join('');
  var pc = data.paymentReceived ? "#22c55e" : "#ef4444";
  var pb = data.paymentReceived ? "#f0fdf4" : "#fef2f2";

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;"><div style="background:#f59e0b;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;"><h2 style="margin:0;">💰 Material Sale'+(data.paymentReceived?'':' — Unpaid')+'</h2></div><div style="background:#fff;border:1px solid #e5e7eb;padding:24px;border-radius:0 0 8px 8px;"><table style="width:100%;border-collapse:collapse;margin-bottom:16px;"><tr><td style="padding:8px 0;color:#6b7280;width:100px;">Customer:</td><td style="font-weight:bold;">'+data.customer+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Date:</td><td style="font-weight:bold;">'+jobDate+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Site:</td><td>'+(data.jobSite||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">PO #:</td><td>'+(data.poNumber||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Driver:</td><td>'+data.driver+'</td></tr></table><div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:8px;padding:16px;margin-bottom:16px;"><h3 style="margin:0 0 12px;color:#92400e;">Materials Sold (New):</h3><table style="width:100%;border-collapse:collapse;"><tr style="background:#fef3c7;"><th style="padding:6px 12px;text-align:left;color:#92400e;">Item</th><th style="padding:6px 12px;text-align:left;color:#92400e;">Qty</th><th style="padding:6px 12px;text-align:left;color:#92400e;">Cond</th></tr>'+iHtml+'</table></div><div style="background:'+pb+';border:2px solid '+pc+';border-radius:8px;padding:16px;margin-bottom:16px;"><h3 style="margin:0 0 8px;color:'+pc+';">'+(data.paymentReceived?'✅':'❌')+' Payment</h3><table style="width:100%;border-collapse:collapse;"><tr><td style="padding:4px 0;color:#6b7280;width:100px;">Received:</td><td style="font-weight:bold;color:'+pc+';">'+(data.paymentReceived?"YES":"NO")+'</td></tr>'+(data.paymentReceived?'<tr><td style="padding:4px 0;color:#6b7280;">Type:</td><td style="font-weight:bold;">'+payLabel+'</td></tr>':'')+(data.paymentNotes?'<tr><td style="padding:4px 0;color:#6b7280;">Notes:</td><td>'+data.paymentNotes+'</td></tr>':'')+'</table></div>'+(data.paymentReceived?'':'<div style="background:#fee2e2;border:1px solid #ef4444;border-radius:8px;padding:12px;text-align:center;"><strong style="color:#dc2626;">ACTION REQUIRED</strong><p style="margin:4px 0 0;color:#991b1b;">Payment NOT collected.</p></div>')+'</div></div>';

  GmailApp.sendEmail(salesEmail, subject, body, {htmlBody:html, name:"TGD Fence Tracker"});
  return jsonResponse({ status: "success", emailSent: true });
}

// ── Handle Damaged Email ─────────────────────────────────────────────  ← NEW in v7
// Logs each damaged line to the "Damage Log" tab and emails the sales rep.
// The yard is adjusted separately, in handleInventory (jobType "damaged").
function handleDamagedEmail(data) {
  var salesEmail = data.salesEmail;
  if (!salesEmail) return jsonResponse({ status: "error", message: "No sales email" });

  var HEADERS = ["Submitted","Job Date","Driver","Customer","Job Site","PO#","Item","Qty","Unit","Condition","Reason","Email Sent To"];
  var dmgSheet = getOrCreateSheet("Damage Log", HEADERS);
  var jobDate = fmtDate(data.jobDate);
  var reason = data.damageNotes || "";

  var items = data.items || [];
  items.forEach(function(item) {
    dmgSheet.appendRow([data.timestamp, jobDate, data.driver, data.customer, data.jobSite, data.poNumber, item.item, item.quantity, item.unit, (item.condition||"used").toUpperCase(), reason, salesEmail]);
  });

  var list = items.map(function(i){return "  • "+i.quantity+" "+i.unit+" of "+i.item+" ("+(i.condition||"used").toUpperCase()+")";}).join("\n");
  var subject = "🛠️ DAMAGED FENCING - " + data.customer + " - Removed from yard";
  var body = ["DAMAGED MATERIAL ALERT","═══════════════════","","Customer: "+data.customer,"Job Date: "+jobDate,"Job Site: "+(data.jobSite||"N/A"),"PO#: "+(data.poNumber||"N/A"),"Driver: "+data.driver,"","DAMAGED (removed from yard):","───────────────────",list,"","What happened: "+(reason||"None given"),"","ACTION: This material was pulled from inventory. Repair, replace, or write off.","","— TGD Fence Tracker"].join("\n");

  var dHtml = items.map(function(i){var c=i.condition==="new"?"#16a34a":"#64748b";return '<tr><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;">'+i.item+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;font-weight:bold;">'+i.quantity+' '+i.unit+'</td><td style="padding:6px 12px;border-bottom:1px solid #f0f0f0;color:'+c+';font-weight:bold;">'+(i.condition||"used").toUpperCase()+'</td></tr>';}).join('');

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;"><div style="background:#ef4444;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;"><h2 style="margin:0;">🛠️ Damaged Material</h2></div><div style="background:#fff;border:1px solid #e5e7eb;padding:24px;border-radius:0 0 8px 8px;"><table style="width:100%;border-collapse:collapse;margin-bottom:16px;"><tr><td style="padding:8px 0;color:#6b7280;width:100px;">Customer:</td><td style="font-weight:bold;">'+data.customer+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Date:</td><td style="font-weight:bold;">'+jobDate+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Job Site:</td><td>'+(data.jobSite||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">PO #:</td><td>'+(data.poNumber||'N/A')+'</td></tr><tr><td style="padding:8px 0;color:#6b7280;">Driver:</td><td>'+data.driver+'</td></tr></table><div style="background:#fee2e2;border:1px solid #ef4444;border-radius:8px;padding:16px;margin-bottom:16px;"><h3 style="margin:0 0 12px;color:#991b1b;">Damaged — pulled from yard:</h3><table style="width:100%;border-collapse:collapse;"><tr style="background:#fee2e2;"><th style="padding:6px 12px;text-align:left;color:#991b1b;">Item</th><th style="padding:6px 12px;text-align:left;color:#991b1b;">Qty</th><th style="padding:6px 12px;text-align:left;color:#991b1b;">Cond</th></tr>'+dHtml+'</table></div>'+(reason?'<p style="color:#6b7280;"><strong>What happened:</strong> '+reason+'</p>':'')+'<div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:8px;padding:12px;text-align:center;"><strong style="color:#92400e;">HEADS UP</strong><p style="margin:4px 0 0;color:#92400e;">This material was removed from inventory. Repair, replace, or write off.</p></div></div></div>';

  GmailApp.sendEmail(salesEmail, subject, body, {htmlBody:html, name:"TGD Fence Tracker"});
  return jsonResponse({ status: "success", emailSent: true });
}
