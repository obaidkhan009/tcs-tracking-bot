/**
 * TCS Tracking Bot for Google Sheets.
 *
 * Adds a "TCS Tracking" menu to the spreadsheet:
 *   - Search box (sidebar): type a tracking number, see the status, the row is updated.
 *   - Check selected rows / Check all rows / Check pending rows.
 *   - Enable "Track" sheet: a sheet where typing a number in cell B2 shows the
 *     status right there, without opening anything.
 *
 * Install: Extensions > Apps Script, replace the content with this file, save,
 * reload the spreadsheet. The first use asks for permission (external requests
 * to tcsexpress.com and editing this sheet).
 */

var TRACKING_HEADERS = ["trackingnumber", "tracking", "trackingno", "tracking#", "cn", "cnnumber",
  "cnno", "consignment", "consignmentno", "consignmentnumber", "awb", "awbno"];
var COURIER_HEADERS = ["courier", "carrier", "courierservice", "shipper", "service"];
var OUTPUT_COLUMNS = ["Status", "Last Update", "Origin", "Destination", "Booked On",
  "Delivered On", "Received By", "Days In Transit", "ETA / Note", "Checked At"];
var COLORS = { "DELIVERED": "#c6efce", "IN TRANSIT": "#ffeb9c", "RETURNED": "#f4b084",
  "NOT FOUND": "#ffc7ce", "ERROR": "#ffc7ce" };
var TZ = null;  // resolved to the spreadsheet time zone on first use
var TRACK_SHEET = "Track";

// ---------------------------------------------------------------- menu

function onOpen() {
  SpreadsheetApp.getUi().createMenu("TCS Tracking")
    .addItem("Search box (sidebar)", "showSidebar")
    .addSeparator()
    .addItem("Check selected rows", "checkSelected")
    .addItem("Check pending rows (not delivered)", "checkPending")
    .addItem("Check all rows", "checkAllRows")
    .addSeparator()
    .addItem("Enable 'Track' sheet (search cell)", "enableTrackSheet")
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutput(SIDEBAR_HTML).setTitle("TCS Tracking").setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------------------------------------------------------------- TCS

function fetchTcs(number) {
  var payload = { body: { url: "trackapinew", type: "GET", headers: {},
    param: "consignee=" + String(number).replace(/\s+/g, ""), payload: {} } };
  var resp = UrlFetchApp.fetch("https://www.tcsexpress.com/apibridge", {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload),
    headers: { "Origin": "https://www.tcsexpress.com", "Referer": "https://www.tcsexpress.com/track/",
      "User-Agent": "Mozilla/5.0" },
    muteHttpExceptions: true, followRedirects: true
  });
  if (resp.getResponseCode() !== 200) throw new Error("TCS returned HTTP " + resp.getResponseCode());
  var data = JSON.parse(resp.getContentText());
  if (!data.isSuccess) throw new Error("TCS error: " + JSON.stringify(data.responseData));
  return data.responseData || {};
}

function parseDate(text) {
  if (!text) return null;
  var s = String(text).trim().replace(/\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/i, "").replace(/,/g, " ")
    .replace(/\s+/g, " ").trim();
  var d = new Date(s);
  if (isNaN(d.getTime())) { d = new Date(s.replace(/-/g, " ")); }
  if (isNaN(d.getTime())) {
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(\d{1,2}:\d{2})?/);   // dd/mm/yyyy
    if (m) d = new Date(m[3] + "-" + m[2] + "-" + m[1] + (m[4] ? "T" + m[4] : ""));
  }
  return isNaN(d.getTime()) ? null : d;
}

function tz() {
  if (!TZ) { try { TZ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); } catch (e) { TZ = Session.getScriptTimeZone(); } }
  return TZ;
}
function fmt(d) { return d ? Utilities.formatDate(d, tz(), "dd MMM yyyy hh:mm a") : ""; }

function pick(list, number) {
  if (!list) return {};
  if (!Array.isArray(list)) return list;
  for (var i = 0; i < list.length; i++) {
    if (list[i] && String(list[i].consignmentno || "").trim() === number) return list[i];
  }
  return list[0] || {};
}

function has(text, words) {
  var s = String(text || "").toLowerCase();
  for (var i = 0; i < words.length; i++) if (s.indexOf(words[i]) !== -1) return true;
  return false;
}
var NOT_DELIVERED = ["undeliver", "not deliver", "attempt", "refus", "return", "rts", "hold"];
var DELIVERED = ["delivered", "deliverd", "delivery done", "handed over"];
var RETURNED = ["returned to sender", "returned to shipper", "return to sender", "return to shipper", "rts"];
function isDelivered(s) { return !has(s, NOT_DELIVERED) && has(s, DELIVERED); }
function isReturned(s) { return has(s, RETURNED); }

/** Look a number up with TCS and return a plain result object. */
function trackTcs(number) {
  number = String(number).replace(/\s+/g, "");
  var data = fetchTcs(number);
  var info = pick(data.shipmentinfo, number);
  var delivery = pick(data.deliveryinfo, number);
  var cps = (data.checkpoints || []).filter(function (c) { return c; });
  if (cps.some(function (c) { return String(c.consignmentno || "") === number; })) {
    cps = cps.filter(function (c) { return String(c.consignmentno || "") === number; });
  }
  var summaryText = String(data.shipmentsummary || "").trim();

  if (!Object.keys(info).length && !cps.length) {
    return { number: number, found: false, state: "NOT FOUND",
      summary: summaryText || "No data found for this tracking number.", checkpoints: [] };
  }
  var checkpoints = cps.map(function (c) {
    return { time: parseDate(c.datetime), rawTime: String(c.datetime || ""),
      status: String(c.status || "").trim(), detail: String(c.recievedby || c.receivedby || "").trim() };
  });
  if (checkpoints.length && checkpoints.every(function (c) { return c.time; })) {
    checkpoints.sort(function (a, b) { return b.time - a.time; });
  }
  var booking = parseDate(info.bookingdate);
  var deliveredOn = Object.keys(delivery).length ? parseDate(delivery.datetime) : null;
  var receivedBy = String(delivery.recievedby || delivery.receivedby || "").trim();
  var latest = checkpoints[0] || null;
  var latestStatus = latest ? latest.status : "";
  var dStatus = String(delivery.status || "").trim();
  var dCode = String(delivery.code || "").trim().toUpperCase();
  var hasDelivery = Object.keys(delivery).length > 0;

  var returned = isReturned(dStatus) || dCode === "RS" || (!hasDelivery && isReturned(latestStatus));
  var delivered = !returned && (isDelivered(dStatus) || dCode === "OK" ||
    (!dStatus && hasDelivery && (deliveredOn || receivedBy)) ||
    (!hasDelivery && (isDelivered(latestStatus) || isDelivered(summaryText))));
  if ((delivered || returned) && !deliveredOn) {
    var hit = null;
    for (var i = 0; i < checkpoints.length; i++) {
      if (returned ? isReturned(checkpoints[i].status) : isDelivered(checkpoints[i].status)) { hit = checkpoints[i]; break; }
    }
    deliveredOn = hit ? hit.time : (latest ? latest.time : null);
    if (hit && delivered && !receivedBy) receivedBy = hit.detail;
  }
  var end = deliveredOn || new Date();
  var days = booking ? Math.floor((end - booking) / 86400000) : "";
  var state = delivered ? "DELIVERED" : returned ? "RETURNED" : "IN TRANSIT";
  var summary, note;
  if (state === "DELIVERED") {
    summary = "Delivered" + (receivedBy ? " to " + receivedBy : "") + (deliveredOn ? " on " + fmt(deliveredOn) : "");
    note = "Delivered on " + (deliveredOn ? fmt(deliveredOn) : "unknown date") + (receivedBy ? " (received by " + receivedBy + ")" : "") + ".";
  } else if (state === "RETURNED") {
    summary = "Returned to sender" + (deliveredOn ? " on " + fmt(deliveredOn) : "");
    note = "Returned to sender. Contact the local TCS office.";
  } else {
    summary = (latestStatus || summaryText || "In transit") +
      (latest && latest.time ? " (" + fmt(latest.time) + ")" : latest && latest.rawTime ? " (" + latest.rawTime + ")" : "");
    note = "In transit for " + (days === "" ? "unknown time" : days + " day(s)") + ". Latest scan: " +
      (latestStatus || "no scan yet") + ". TCS does not give an expected delivery date.";
  }
  return { number: number, found: true, state: state, summary: summary,
    origin: String(info.origin || ""), destination: String(info.destination || ""),
    bookedOn: fmt(booking), deliveredOn: fmt(deliveredOn), receivedBy: receivedBy,
    days: days, note: note,
    checkpoints: checkpoints.map(function (c) {
      return { time: c.time ? fmt(c.time) : c.rawTime, status: c.status, detail: c.detail };
    }) };
}

function safeTrack(number) {
  try { return trackTcs(number); }
  catch (e) { return { number: String(number), found: false, state: "ERROR", summary: String(e.message || e), checkpoints: [] }; }
}

// ---------------------------------------------------------------- sheet helpers

function norm(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9#]/g, ""); }
function cleanNumber(v) { return v === null || v === undefined ? "" : String(v).replace(/\s+/g, "").replace(/\.0$/, ""); }

/** The data sheet: the active sheet unless it is the Track sheet, then the first other sheet. */
function dataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  if (sh.getName() !== TRACK_SHEET) return sh;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (sheets[i].getName() !== TRACK_SHEET) return sheets[i];
  throw new Error("No data sheet found.");
}

/** Finds the header row and the tracking / courier columns; adds output columns. */
function layout(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var rowsToScan = Math.min(sheet.getLastRow(), 10) || 1;
  var values = sheet.getRange(1, 1, rowsToScan, lastCol).getValues();
  var headerRow = -1, trackCol = -1, courierCol = -1, headers = {};
  for (var r = 0; r < values.length && headerRow < 0; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (TRACKING_HEADERS.indexOf(norm(values[r][c])) !== -1) { headerRow = r + 1; break; }
    }
  }
  if (headerRow < 0) throw new Error("No 'Tracking Number' column found in sheet '" + sheet.getName() +
    "'. Name the column Tracking Number, CN, Consignment No or AWB.");
  var hdr = values[headerRow - 1];
  for (var c2 = 0; c2 < hdr.length; c2++) {
    var n = norm(hdr[c2]);
    if (n) headers[n] = c2 + 1;
    if (trackCol < 0 && TRACKING_HEADERS.indexOf(n) !== -1) trackCol = c2 + 1;
    if (courierCol < 0 && COURIER_HEADERS.indexOf(n) !== -1) courierCol = c2 + 1;
  }
  var outCols = {}, nextCol = lastCol + 1;
  OUTPUT_COLUMNS.forEach(function (name) {
    var col = headers[norm(name)];
    if (!col) {
      col = nextCol++;
      sheet.getRange(headerRow, col).setValue(name).setFontWeight("bold");
    }
    outCols[name] = col;
  });
  return { sheet: sheet, headerRow: headerRow, trackCol: trackCol, courierCol: courierCol, outCols: outCols };
}

function writeResult(L, row, res) {
  var values = [res.state, res.summary, res.origin || "", res.destination || "", res.bookedOn || "",
    res.deliveredOn || "", res.receivedBy || "", res.days === undefined ? "" : res.days, res.note || "",
    Utilities.formatDate(new Date(), tz(), "dd MMM yyyy hh:mm a")];
  OUTPUT_COLUMNS.forEach(function (name, i) { L.sheet.getRange(row, L.outCols[name]).setValue(values[i]); });
  L.sheet.getRange(row, L.outCols["Status"]).setBackground(COLORS[res.state] || null).setFontWeight("bold");
}

function findRows(L, number) {
  var target = cleanNumber(number);
  var last = L.sheet.getLastRow();
  if (last <= L.headerRow) return [];
  var nums = L.sheet.getRange(L.headerRow + 1, L.trackCol, last - L.headerRow, 1).getValues();
  var exact = [], partial = [];
  for (var i = 0; i < nums.length; i++) {
    var n = cleanNumber(nums[i][0]);
    if (!n) continue;
    if (n === target) exact.push({ row: L.headerRow + 1 + i, number: n });
    else if (target && n.indexOf(target) !== -1) partial.push({ row: L.headerRow + 1 + i, number: n });
  }
  return exact.length ? exact : partial;
}

// ---------------------------------------------------------------- actions (menu + sidebar)

/** Sidebar / Track-sheet entry: look up a number, update its row, return the result. */
function checkNumber(number) {
  var L = layout(dataSheet());
  var matches = findRows(L, number);
  var results = [];
  if (!matches.length) {
    var r0 = safeTrack(number); r0.row = null; r0.inSheet = false; results.push(r0);
    return results;
  }
  matches.forEach(function (m) {
    var r = safeTrack(m.number); r.row = m.row; r.inSheet = true;
    writeResult(L, m.row, r);
    results.push(r);
  });
  return results;
}

function checkRows(rowNumbers, L) {
  var counts = {};
  rowNumbers.forEach(function (row, i) {
    var num = cleanNumber(L.sheet.getRange(row, L.trackCol).getValue());
    if (!num) return;
    var r = safeTrack(num);
    writeResult(L, row, r);
    counts[r.state] = (counts[r.state] || 0) + 1;
    if (i < rowNumbers.length - 1) Utilities.sleep(500);
  });
  return counts;
}

function summarize(counts) {
  var parts = [];
  Object.keys(counts).sort().forEach(function (k) { parts.push(k + ": " + counts[k]); });
  return parts.length ? parts.join(", ") : "Nothing to check.";
}

function checkSelected() {
  var sh = dataSheet();
  var L = layout(sh);
  var ranges = sh.getActiveRangeList() ? sh.getActiveRangeList().getRanges() : [sh.getActiveRange()];
  var rows = {};
  ranges.forEach(function (rg) {
    for (var r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r > L.headerRow) rows[r] = true;
  });
  var list = Object.keys(rows).map(Number).sort(function (a, b) { return a - b; });
  SpreadsheetApp.getUi().alert("Done. " + summarize(checkRows(list, L)));
}

function allDataRows(L, pendingOnly) {
  var last = L.sheet.getLastRow();
  var rows = [];
  if (last <= L.headerRow) return rows;
  var statuses = L.sheet.getRange(L.headerRow + 1, L.outCols["Status"], last - L.headerRow, 1).getValues();
  for (var r = L.headerRow + 1; r <= last; r++) {
    var st = String(statuses[r - L.headerRow - 1][0] || "").toUpperCase();
    if (pendingOnly && (st === "DELIVERED" || st === "RETURNED")) continue;
    rows.push(r);
  }
  return rows;
}

function checkAllRows() {
  var L = layout(dataSheet());
  SpreadsheetApp.getUi().alert("Done. " + summarize(checkRows(allDataRows(L, false), L)));
}

function checkPending() {
  var L = layout(dataSheet());
  SpreadsheetApp.getUi().alert("Done. " + summarize(checkRows(allDataRows(L, true), L)));
}

// ---------------------------------------------------------------- "Track" sheet with a search cell

function enableTrackSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TRACK_SHEET) || ss.insertSheet(TRACK_SHEET);
  sh.clear();
  sh.getRange("A1").setValue("Type a tracking number in the yellow cell and press Enter:").setFontWeight("bold");
  sh.getRange("A2").setValue("Tracking number").setFontWeight("bold");
  sh.getRange("B2").setBackground("#fff2cc").setFontSize(12);
  sh.getRange("A4:B4").setValues([["Field", "Value"]]).setFontWeight("bold");
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 520);
  // one installable trigger so edits can call TCS
  var exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "onTrackEdit"; });
  if (!exists) ScriptApp.newTrigger("onTrackEdit").forSpreadsheet(ss).onEdit().create();
  ss.setActiveSheet(sh);
  sh.getRange("B2").activate();
  SpreadsheetApp.getUi().alert("Ready. Type a tracking number in cell B2 of the 'Track' sheet and press Enter.");
}

function onTrackEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== TRACK_SHEET || e.range.getA1Notation() !== "B2") return;
  var number = cleanNumber(e.range.getValue());
  if (!number) return;
  sh.getRange("A5:B40").clearContent().setBackground(null);
  sh.getRange("A5").setValue("Checking with TCS...");
  SpreadsheetApp.flush();
  var results = checkNumber(number);
  var r = results[0];
  var rows = [
    ["Status", r.state], ["Latest", r.summary], ["Route", (r.origin || "?") + " -> " + (r.destination || "?")],
    ["Booked on", r.bookedOn || ""], ["Delivered on", r.deliveredOn || ""], ["Received by", r.receivedBy || ""],
    ["Days in transit", r.days === undefined ? "" : r.days], ["Note", r.note || ""],
    ["In your sheet", r.inSheet ? "Yes, row " + r.row + " updated" : "No, not found in the data sheet"]
  ];
  if (!r.found) rows = [["Status", r.state], ["Message", r.summary]];
  sh.getRange(5, 1, rows.length, 2).setValues(rows);
  sh.getRange(5, 2).setBackground(COLORS[r.state] || null).setFontWeight("bold");
  var hist = r.checkpoints || [];
  if (hist.length) {
    var start = 5 + rows.length + 1;
    sh.getRange(start, 1, 1, 2).setValues([["History", ""]]).setFontWeight("bold");
    var hv = hist.slice(0, 25).map(function (c) { return [c.time, c.status + (c.detail ? " (" + c.detail + ")" : "")]; });
    sh.getRange(start + 1, 1, hv.length, 2).setValues(hv);
  }
}

// ---------------------------------------------------------------- sidebar

var SIDEBAR_HTML =
'<!DOCTYPE html><html><head><base target="_top">' +
'<style>body{font-family:Arial,sans-serif;font-size:13px;margin:10px}input{width:100%;padding:8px;font-size:14px;box-sizing:border-box}' +
'button{margin-top:8px;padding:8px 12px;font-size:13px;cursor:pointer}#out{margin-top:12px;white-space:pre-wrap}' +
'.badge{display:inline-block;padding:3px 8px;border-radius:4px;font-weight:bold;margin-bottom:6px}' +
'.DELIVERED{background:#c6efce}.IN{background:#ffeb9c}.RETURNED{background:#f4b084}.NOT{background:#ffc7ce}.ERROR{background:#ffc7ce}' +
'table{border-collapse:collapse;width:100%;margin-top:8px}td{border-bottom:1px solid #eee;padding:3px;vertical-align:top;font-size:12px}</style></head><body>' +
'<b>Tracking number</b><input id="n" placeholder="type and press Enter" autofocus>' +
'<button id="go">Check</button> <label><input type="checkbox" id="h" style="width:auto"> show history</label>' +
'<div id="out"></div>' +
'<script>' +
'var n=document.getElementById("n"),out=document.getElementById("out");' +
'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;");}' +
'function show(results){out.innerHTML="";results.forEach(function(r){var cls=r.state.split(" ")[0];' +
'var h="<div class=\\"badge "+cls+"\\">"+esc(r.state)+"</div><div><b>"+esc(r.number)+"</b></div>";' +
'if(r.found){h+="<div>"+esc(r.summary)+"</div><div>Route: "+esc(r.origin||"?")+" \\u2192 "+esc(r.destination||"?")+"</div>";' +
'if(r.bookedOn)h+="<div>Booked: "+esc(r.bookedOn)+"</div>";h+="<div style=\\"color:#555\\">"+esc(r.note)+"</div>";' +
'h+="<div style=\\"color:#555\\">"+(r.inSheet?"Row "+r.row+" updated in the sheet.":"Not in the sheet (nothing written).")+"</div>";' +
'if(document.getElementById("h").checked&&r.checkpoints.length){h+="<table>";r.checkpoints.forEach(function(c){h+="<tr><td>"+esc(c.time)+"</td><td>"+esc(c.status)+(c.detail?" ("+esc(c.detail)+")":"")+"</td></tr>";});h+="</table>";}}' +
'else{h+="<div>"+esc(r.summary)+"</div>";}' +
'out.innerHTML+=h+"<hr>";});}' +
'function go(){var v=n.value.trim();if(!v)return;out.textContent="Checking "+v+" with TCS...";' +
'google.script.run.withSuccessHandler(function(res){show(res);n.value="";n.focus();})' +
'.withFailureHandler(function(e){out.textContent="Error: "+e.message;}).checkNumber(v);}' +
'document.getElementById("go").onclick=go;n.addEventListener("keydown",function(e){if(e.key==="Enter")go();});' +
'</script></body></html>';
