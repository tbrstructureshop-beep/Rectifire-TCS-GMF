/**
 * ============================================================
 *  RECTIFIER MONITORING SYSTEM — Google Apps Script
 *  Electroplating Facility | Specialized Process Division
 * ============================================================
 *
 *  SETUP:
 *  1. Open your Google Spreadsheet (matching Rectifier_Monitoring.xlsx layout).
 *  2. Extensions → Apps Script → paste this file.
 *  3. Run  onOpen()  once to create the custom menu.
 *  4. Set up triggers via  setupTriggers()  (run once).
 *
 *  SHEET NAMES expected:
 *    "Rectifier Data"      — master rectifier table (rows 6–10, cols A–N)
 *    "Logbook"             — damage log (rows 4+, cols A–I)
 *    "Calibration Tracker" — auto-pulled, do not edit manually
 * ============================================================
 */

// ── CONSTANTS ──────────────────────────────────────────────────
const SHEET_RECT   = "Rectifier Data";
const SHEET_LOG    = "Logbook";
const SHEET_CAL    = "Calibration Tracker";

const RECT_DATA_START_ROW = 6;
const RECT_DATA_END_ROW   = 10;
const LOG_START_ROW       = 4;

// Column indices (1-based) in "Rectifier Data"
const COL_RECT_ID         = 2;   // B
const COL_RECT_NAME       = 3;   // C
const COL_RECT_STATUS     = 4;   // D
const COL_RECT_VMIN       = 5;   // E
const COL_RECT_VMAX       = 6;   // F
const COL_RECT_CMIN       = 7;   // G
const COL_RECT_CMAX       = 8;   // H
const COL_RECT_CALIBNO    = 9;   // I
const COL_RECT_CALIBDATE  = 10;  // J
const COL_RECT_DUEDATE    = 11;  // K
const COL_RECT_DAYS       = 12;  // L  (formula)
const COL_RECT_DAYSTATUS  = 13;  // M  (formula)

// Column indices in "Logbook"
const COL_LOG_NO          = 1;   // A
const COL_LOG_DATE        = 2;   // B
const COL_LOG_UNIT        = 3;   // C
const COL_LOG_UNITID      = 4;   // D
const COL_LOG_SEVERITY    = 5;   // E
const COL_LOG_DAMAGE      = 6;   // F
const COL_LOG_REPORTER    = 7;   // G
const COL_LOG_ACTION      = 8;   // H
const COL_LOG_STATUS      = 9;   // I

// Email for alerts
const ALERT_EMAIL = Session.getActiveUser().getEmail();  // change if needed


// ── MENU ───────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Rectifier System")
    .addItem("🔄 Sync Dashboard Summary",   "syncDashboard")
    .addItem("📋 Add Log Entry (Form)",      "showAddLogSidebar")
    .addItem("✏️  Edit Rectifier Unit",       "showEditRectSidebar")
    .addSeparator()
    .addItem("📧 Send Overdue Alert Email",  "sendOverdueAlertEmail")
    .addItem("📊 Refresh Calibration Tracker", "refreshCalibrationTracker")
    .addSeparator()
    .addItem("⚙️  Setup Auto Triggers",      "setupTriggers")
    .addItem("🗑️  Remove All Triggers",       "removeAllTriggers")
    .addToUi();
}


// ── SYNC DASHBOARD SUMMARY ─────────────────────────────────────
/**
 * Recalculates KPI cells in "Rectifier Data" rows 3–4.
 * Called automatically by time-driven trigger.
 */
function syncDashboard() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsR   = ss.getSheetByName(SHEET_RECT);
  if (!wsR) { logError_("Sheet not found: " + SHEET_RECT); return; }

  const data  = wsR.getRange(RECT_DATA_START_ROW, 1,
                  RECT_DATA_END_ROW - RECT_DATA_START_ROW + 1, 14).getValues();

  let total = 0, inService = 0, unservice = 0, overdue = 0;
  const today = new Date(); today.setHours(0,0,0,0);

  data.forEach(row => {
    if (!row[COL_RECT_ID - 1]) return;  // skip blank rows
    total++;
    const status = String(row[COL_RECT_STATUS - 1]).trim().toLowerCase();
    if (status === "in service")  inService++;
    else                          unservice++;

    const dueRaw = row[COL_RECT_DUEDATE - 1];
    if (dueRaw) {
      const dueDate = (dueRaw instanceof Date) ? dueRaw : new Date(dueRaw);
      if (dueDate < today) overdue++;
    }
  });

  // Write KPI values to the merged cells (row 4: A4, D4, G4, J4)
  wsR.getRange("A4").setValue(total);
  wsR.getRange("D4").setValue(inService);
  wsR.getRange("G4").setValue(unservice);
  wsR.getRange("J4").setValue(overdue);
  wsR.getRange("M3").setValue("Updated: " + Utilities.formatDate(new Date(),
      Session.getScriptTimeZone(), "dd MMM yyyy HH:mm"));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Dashboard synced — Total: ${total} | In Service: ${inService} | Unservice: ${unservice} | Overdue: ${overdue}`,
    "Sync Complete", 5);
}


// ── ADD LOG ENTRY SIDEBAR ──────────────────────────────────────
function showAddLogSidebar() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsR   = ss.getSheetByName(SHEET_RECT);
  const units = [];
  for (let r = RECT_DATA_START_ROW; r <= RECT_DATA_END_ROW; r++) {
    const id   = wsR.getRange(r, COL_RECT_ID).getValue();
    const name = wsR.getRange(r, COL_RECT_NAME).getValue();
    if (id) units.push({ id, name });
  }

  const unitOptions = units.map(u =>
    `<option value="${u.id}">${u.name} (${u.id})</option>`).join("");

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html>
<head>
<style>
  body { font-family: Arial, sans-serif; font-size: 13px; padding: 12px; background: #f8f9fa; }
  h3 { margin: 0 0 12px; color: #1F6FEB; font-size: 14px; }
  label { display: block; font-size: 11px; font-weight: bold; color: #555; margin: 10px 0 3px; }
  select, input, textarea {
    width: 100%; padding: 7px 9px; border: 1px solid #ccc; border-radius: 5px;
    font-size: 12px; box-sizing: border-box; font-family: Arial;
  }
  textarea { min-height: 60px; resize: vertical; }
  .btn { margin-top: 14px; width: 100%; padding: 9px; background: #1F6FEB;
    color: #fff; border: none; border-radius: 5px; cursor: pointer; font-size: 13px; }
  .btn:hover { background: #1558c0; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
</style>
</head><body>
<h3>📋 Add Log Entry</h3>
<label>Rectifier Unit</label>
<select id="unit">${unitOptions}</select>
<div class="row2">
  <div>
    <label>Date</label>
    <input type="date" id="date" value="${new Date().toISOString().split('T')[0]}"/>
  </div>
  <div>
    <label>Severity</label>
    <select id="severity">
      <option value="Critical">Critical</option>
      <option value="Minor" selected>Minor</option>
      <option value="Resolved">Resolved</option>
    </select>
  </div>
</div>
<label>Damage Detail</label>
<textarea id="damage" placeholder="Describe damage or issue..."></textarea>
<label>Reported By</label>
<input type="text" id="reporter" placeholder="Name / Employee ID"/>
<label>Action Taken</label>
<textarea id="action" placeholder="Action taken or pending..." style="min-height:48px;"></textarea>
<label>Status</label>
<select id="status">
  <option value="Minor">Open (Minor)</option>
  <option value="Critical">Critical – Immediate Action</option>
  <option value="Resolved">Resolved</option>
</select>
<button class="btn" onclick="submitLog()">💾 Save Log Entry</button>
<div id="msg" style="margin-top:8px;font-size:12px;"></div>
<script>
function submitLog() {
  const payload = {
    unit:     document.getElementById('unit').value,
    unitName: document.getElementById('unit').options[document.getElementById('unit').selectedIndex].text,
    date:     document.getElementById('date').value,
    severity: document.getElementById('severity').value,
    damage:   document.getElementById('damage').value,
    reporter: document.getElementById('reporter').value,
    action:   document.getElementById('action').value,
    status:   document.getElementById('status').value
  };
  if (!payload.damage.trim()) { document.getElementById('msg').innerHTML='<span style="color:red">Damage detail required.</span>'; return; }
  document.getElementById('msg').innerHTML = 'Saving...';
  google.script.run
    .withSuccessHandler(r => { document.getElementById('msg').innerHTML = '<span style="color:green">✅ ' + r + '</span>'; })
    .withFailureHandler(e => { document.getElementById('msg').innerHTML = '<span style="color:red">Error: ' + e.message + '</span>'; })
    .addLogEntryFromSidebar(payload);
}
</script>
</body></html>`)
    .setTitle("Add Log Entry")
    .setWidth(340);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called by sidebar. Appends a new row to the Logbook sheet.
 */
function addLogEntryFromSidebar(payload) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wsL = ss.getSheetByName(SHEET_LOG);
  if (!wsL) throw new Error("Logbook sheet not found.");

  // Find next blank row
  const lastRow = wsL.getLastRow();
  const nextRow = Math.max(LOG_START_ROW, lastRow + 1);

  // Entry number
  const entryNo = nextRow - LOG_START_ROW + 1;

  // Strip unit name from combined string (format: "Name (ID)")
  const match = payload.unitName.match(/^(.*?)\s*\(([^)]+)\)/);
  const unitName = match ? match[1].trim() : payload.unitName;
  const unitId   = payload.unit;

  wsL.getRange(nextRow, COL_LOG_NO).setValue(entryNo);
  wsL.getRange(nextRow, COL_LOG_DATE).setValue(new Date(payload.date));
  wsL.getRange(nextRow, COL_LOG_DATE).setNumberFormat("DD-MMM-YYYY");
  wsL.getRange(nextRow, COL_LOG_UNIT).setValue(unitName);
  wsL.getRange(nextRow, COL_LOG_UNITID).setValue(unitId);
  wsL.getRange(nextRow, COL_LOG_SEVERITY).setValue(payload.severity);
  wsL.getRange(nextRow, COL_LOG_DAMAGE).setValue(payload.damage);
  wsL.getRange(nextRow, COL_LOG_REPORTER).setValue(payload.reporter);
  wsL.getRange(nextRow, COL_LOG_ACTION).setValue(payload.action);
  wsL.getRange(nextRow, COL_LOG_STATUS).setValue(payload.status);

  // Apply row formatting
  formatLogRow_(wsL, nextRow, payload.severity);

  // Trigger alert if critical
  if (payload.severity === "Critical") {
    sendCriticalAlert_(unitName, unitId, payload.damage, payload.reporter);
  }

  return `Log #${entryNo} added for ${unitName}`;
}


// ── EDIT RECTIFIER SIDEBAR ─────────────────────────────────────
function showEditRectSidebar() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsR   = ss.getSheetByName(SHEET_RECT);
  const units = [];
  for (let r = RECT_DATA_START_ROW; r <= RECT_DATA_END_ROW; r++) {
    const row = wsR.getRange(r, 1, 1, 14).getValues()[0];
    units.push({
      row:       r,
      id:        row[COL_RECT_ID - 1],
      name:      row[COL_RECT_NAME - 1],
      status:    row[COL_RECT_STATUS - 1],
      voltMin:   row[COL_RECT_VMIN - 1],
      voltMax:   row[COL_RECT_VMAX - 1],
      currMin:   row[COL_RECT_CMIN - 1],
      currMax:   row[COL_RECT_CMAX - 1],
      calibNo:   row[COL_RECT_CALIBNO - 1],
      calibDate: row[COL_RECT_CALIBDATE - 1] instanceof Date
                   ? Utilities.formatDate(row[COL_RECT_CALIBDATE - 1], Session.getScriptTimeZone(), "yyyy-MM-dd")
                   : row[COL_RECT_CALIBDATE - 1],
      dueDate:   row[COL_RECT_DUEDATE - 1] instanceof Date
                   ? Utilities.formatDate(row[COL_RECT_DUEDATE - 1], Session.getScriptTimeZone(), "yyyy-MM-dd")
                   : row[COL_RECT_DUEDATE - 1],
    });
  }

  const unitOptionsJson = JSON.stringify(units);

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html>
<head>
<style>
  body { font-family: Arial, sans-serif; font-size: 13px; padding: 12px; }
  h3 { margin: 0 0 10px; color: #1F6FEB; font-size: 14px; }
  label { display: block; font-size: 11px; font-weight: bold; color: #555; margin: 8px 0 3px; }
  select, input { width: 100%; padding: 7px 9px; border: 1px solid #ccc; border-radius: 5px;
    font-size: 12px; box-sizing: border-box; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .btn { margin-top: 14px; width: 100%; padding: 9px; background: #1F6FEB;
    color: #fff; border: none; border-radius: 5px; cursor: pointer; font-size: 13px; }
  .btn:hover { background: #1558c0; }
</style>
</head><body>
<h3>✏️ Edit Rectifier Unit</h3>
<label>Select Unit</label>
<select id="unitSel" onchange="loadUnit(this.value)">
  <option value="">-- Choose unit --</option>
</select>
<div id="form" style="display:none">
  <label>Status</label>
  <select id="eStatus"><option>In Service</option><option>Unservice</option></select>
  <div class="row2">
    <div><label>Volt Min (VDC)</label><input type="number" id="eVMin"/></div>
    <div><label>Volt Max (VDC)</label><input type="number" id="eVMax"/></div>
  </div>
  <div class="row2">
    <div><label>Curr Min (A)</label><input type="number" id="eCMin"/></div>
    <div><label>Curr Max (A)</label><input type="number" id="eCMax"/></div>
  </div>
  <label>Calib. Reg. No.</label>
  <input type="text" id="eCalibNo"/>
  <div class="row2">
    <div><label>Calibration Date</label><input type="date" id="eCalibDate"/></div>
    <div><label>Due Date</label><input type="date" id="eDueDate"/></div>
  </div>
  <button class="btn" onclick="saveUnit()">💾 Save Changes</button>
</div>
<div id="msg" style="margin-top:8px;font-size:12px;"></div>
<script>
const units = ${unitOptionsJson};
const sel   = document.getElementById('unitSel');
units.forEach(u => {
  const o = document.createElement('option');
  o.value = u.row; o.textContent = u.name + ' (' + u.id + ')';
  sel.appendChild(o);
});
function loadUnit(row) {
  if (!row) { document.getElementById('form').style.display='none'; return; }
  const u = units.find(x => x.row == row);
  if (!u) return;
  document.getElementById('eStatus').value  = u.status;
  document.getElementById('eVMin').value    = u.voltMin;
  document.getElementById('eVMax').value    = u.voltMax;
  document.getElementById('eCMin').value    = u.currMin;
  document.getElementById('eCMax').value    = u.currMax;
  document.getElementById('eCalibNo').value = u.calibNo;
  document.getElementById('eCalibDate').value = u.calibDate ? u.calibDate.substring(0,10) : '';
  document.getElementById('eDueDate').value   = u.dueDate   ? u.dueDate.substring(0,10)   : '';
  document.getElementById('form').style.display='block';
}
function saveUnit() {
  const row = parseInt(document.getElementById('unitSel').value);
  const payload = {
    row,
    status:    document.getElementById('eStatus').value,
    voltMin:   parseFloat(document.getElementById('eVMin').value),
    voltMax:   parseFloat(document.getElementById('eVMax').value),
    currMin:   parseFloat(document.getElementById('eCMin').value),
    currMax:   parseFloat(document.getElementById('eCMax').value),
    calibNo:   document.getElementById('eCalibNo').value,
    calibDate: document.getElementById('eCalibDate').value,
    dueDate:   document.getElementById('eDueDate').value,
  };
  document.getElementById('msg').innerHTML = 'Saving...';
  google.script.run
    .withSuccessHandler(r => document.getElementById('msg').innerHTML = '<span style="color:green">✅ ' + r + '</span>')
    .withFailureHandler(e => document.getElementById('msg').innerHTML = '<span style="color:red">Error: ' + e.message + '</span>')
    .saveRectifierFromSidebar(payload);
}
</script>
</body></html>`)
    .setTitle("Edit Rectifier Unit")
    .setWidth(340);

  SpreadsheetApp.getUi().showSidebar(html);
}

function saveRectifierFromSidebar(payload) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wsR = ss.getSheetByName(SHEET_RECT);
  const r   = payload.row;

  wsR.getRange(r, COL_RECT_STATUS).setValue(payload.status);
  wsR.getRange(r, COL_RECT_VMIN).setValue(payload.voltMin);
  wsR.getRange(r, COL_RECT_VMAX).setValue(payload.voltMax);
  wsR.getRange(r, COL_RECT_CMIN).setValue(payload.currMin);
  wsR.getRange(r, COL_RECT_CMAX).setValue(payload.currMax);
  wsR.getRange(r, COL_RECT_CALIBNO).setValue(payload.calibNo);

  if (payload.calibDate) {
    wsR.getRange(r, COL_RECT_CALIBDATE).setValue(new Date(payload.calibDate));
    wsR.getRange(r, COL_RECT_CALIBDATE).setNumberFormat("DD-MMM-YYYY");
  }
  if (payload.dueDate) {
    wsR.getRange(r, COL_RECT_DUEDATE).setValue(new Date(payload.dueDate));
    wsR.getRange(r, COL_RECT_DUEDATE).setNumberFormat("DD-MMM-YYYY");
  }

  syncDashboard();
  return `Row ${r} updated successfully.`;
}


// ── REFRESH CALIBRATION TRACKER ────────────────────────────────
function refreshCalibrationTracker() {
  // The Calibration Tracker uses IMPORTRANGE-style cross-sheet formulas.
  // This function forces a recalculation by touching a helper cell.
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wsC = ss.getSheetByName(SHEET_CAL);
  if (!wsC) return;
  const helperCell = wsC.getRange("H1");
  helperCell.setValue(new Date());
  helperCell.setNumberFormat("DD-MMM-YYYY HH:mm");
  SpreadsheetApp.getActiveSpreadsheet().toast("Calibration Tracker refreshed.", "Done", 3);
}


// ── OVERDUE ALERT EMAIL ────────────────────────────────────────
function sendOverdueAlertEmail() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsR   = ss.getSheetByName(SHEET_RECT);
  const today = new Date(); today.setHours(0,0,0,0);

  const overdue = [];
  const dueSoon = [];

  for (let r = RECT_DATA_START_ROW; r <= RECT_DATA_END_ROW; r++) {
    const row     = wsR.getRange(r, 1, 1, 14).getValues()[0];
    const id      = row[COL_RECT_ID - 1];
    if (!id) continue;
    const name    = row[COL_RECT_NAME - 1];
    const status  = row[COL_RECT_STATUS - 1];
    const dueRaw  = row[COL_RECT_DUEDATE - 1];
    if (!dueRaw) continue;
    const dueDate = (dueRaw instanceof Date) ? dueRaw : new Date(dueRaw);
    const days    = Math.round((dueDate - today) / 86400000);

    if (days < 0)        overdue.push({ id, name, status, dueDate, days });
    else if (days <= 30) dueSoon.push({ id, name, status, dueDate, days });
  }

  if (overdue.length === 0 && dueSoon.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ No overdue or due-soon calibrations.", "All Clear", 4);
    return;
  }

  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM yyyy");

  let body = `<div style="font-family:Arial,sans-serif;max-width:640px">
  <div style="background:#0D1117;color:#E6EDF3;padding:20px 24px;border-radius:8px 8px 0 0;">
    <h2 style="margin:0;font-size:18px">⚡ Rectifier Calibration Alert</h2>
    <p style="margin:4px 0 0;color:#8B949E;font-size:12px;">Electroplating Facility — Specialized Process Division</p>
  </div>
  <div style="background:#f8f9fa;padding:20px 24px;border-radius:0 0 8px 8px;border:1px solid #dee2e6;">`;

  if (overdue.length > 0) {
    body += `<h3 style="color:#F85149;margin:0 0 10px;">🔴 OVERDUE Calibrations (${overdue.length})</h3>
    <table style="width:100%;border-collapse:collapse;margin-bottom:16px;">
      <tr style="background:#FDDCDC;"><th style="padding:8px;text-align:left;font-size:12px;">ID</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Unit</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Status</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Due Date</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Overdue By</th></tr>`;
    overdue.forEach(u => {
      body += `<tr style="border-bottom:1px solid #eee;">
        <td style="padding:8px;font-size:12px;font-weight:bold;">${u.id}</td>
        <td style="padding:8px;font-size:12px;">${u.name}</td>
        <td style="padding:8px;font-size:12px;">${u.status}</td>
        <td style="padding:8px;font-size:12px;">${fmt(u.dueDate)}</td>
        <td style="padding:8px;font-size:12px;color:#F85149;font-weight:bold;">${Math.abs(u.days)} days</td>
      </tr>`;
    });
    body += `</table>`;
  }

  if (dueSoon.length > 0) {
    body += `<h3 style="color:#D29922;margin:0 0 10px;">🟡 DUE SOON (within 30 days) — ${dueSoon.length}</h3>
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#FFF3CC;"><th style="padding:8px;text-align:left;font-size:12px;">ID</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Unit</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Due Date</th>
      <th style="padding:8px;text-align:left;font-size:12px;">Days Remaining</th></tr>`;
    dueSoon.forEach(u => {
      body += `<tr style="border-bottom:1px solid #eee;">
        <td style="padding:8px;font-size:12px;font-weight:bold;">${u.id}</td>
        <td style="padding:8px;font-size:12px;">${u.name}</td>
        <td style="padding:8px;font-size:12px;">${fmt(u.dueDate)}</td>
        <td style="padding:8px;font-size:12px;color:#D29922;font-weight:bold;">${u.days} days</td>
      </tr>`;
    });
    body += `</table>`;
  }

  body += `<p style="margin-top:16px;font-size:11px;color:#888;">
    Generated: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")} |
    Rectifier Monitoring System v1.0
  </p></div></div>`;

  MailApp.sendEmail({
    to:      ALERT_EMAIL,
    subject: `[Rectifier Alert] ${overdue.length} Overdue, ${dueSoon.length} Due Soon — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy")}`,
    htmlBody: body
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Email alert sent to ${ALERT_EMAIL}`, "Email Sent", 5);
}


// ── AUTO-TRIGGER SETUP ─────────────────────────────────────────
/**
 * Installs:
 *   • Daily 08:00 trigger → sendOverdueAlertEmail
 *   • Hourly trigger      → syncDashboard
 */
function setupTriggers() {
  removeAllTriggers();  // clean up first

  // Daily at 08:00
  ScriptApp.newTrigger("sendOverdueAlertEmail")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Every hour
  ScriptApp.newTrigger("syncDashboard")
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Triggers set: daily alert at 08:00 + hourly dashboard sync.", "Triggers Ready", 5);
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}


// ── ON EDIT TRIGGER ────────────────────────────────────────────
/**
 * Automatically syncs dashboard when rectifier status or due date changes.
 * Install via: Edit → Current project's triggers → onEdit → From spreadsheet → On edit
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_RECT) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Only react to status (D) or due date (K) changes in data rows
  if (row >= RECT_DATA_START_ROW && row <= RECT_DATA_END_ROW) {
    if (col === COL_RECT_STATUS || col === COL_RECT_DUEDATE) {
      syncDashboard();
    }
  }
}


// ── INTERNAL HELPERS ───────────────────────────────────────────
function formatLogRow_(sheet, rowNum, severity) {
  const colors = {
    "Critical": { bg: "FDDCDC", fg: "F85149" },
    "Minor":    { bg: "FFF3CC", fg: "D29922" },
    "Resolved": { bg: "D2F4DC", fg: "3FB950" },
  };
  const { bg, fg } = colors[severity] || { bg: "FFFFFF", fg: "000000" };

  // Severity cell (E) and Status cell (I)
  [COL_LOG_SEVERITY, COL_LOG_STATUS].forEach(col => {
    const cell = sheet.getRange(rowNum, col);
    cell.setBackground("#" + bg);
    cell.setFontColor("#" + fg);
    cell.setFontWeight("bold");
  });

  // Light styling for full row
  sheet.getRange(rowNum, 1, 1, 9).setBorder(
    true, true, true, true, true, true, "#D0D7DE",
    SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(rowNum, 42);
}

function sendCriticalAlert_(unitName, unitId, damage, reporter) {
  try {
    MailApp.sendEmail({
      to:      ALERT_EMAIL,
      subject: `🔴 CRITICAL DAMAGE ALERT — ${unitName} (${unitId})`,
      htmlBody: `<div style="font-family:Arial,sans-serif;">
        <div style="background:#F85149;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0;">
          <h2 style="margin:0;">🔴 Critical Damage Reported</h2>
        </div>
        <div style="background:#fff;padding:16px 20px;border:1px solid #FDDCDC;border-radius:0 0 8px 8px;">
          <p><strong>Unit:</strong> ${unitName} (${unitId})</p>
          <p><strong>Reported By:</strong> ${reporter}</p>
          <p><strong>Time:</strong> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")}</p>
          <div style="background:#FFF5F5;border:1px solid #FDDCDC;border-radius:6px;padding:12px;margin-top:10px;">
            <strong>Damage Detail:</strong><br>${damage}
          </div>
          <p style="margin-top:14px;font-size:11px;color:#888;">Rectifier Monitoring System v1.0 — Electroplating Facility</p>
        </div>
      </div>`
    });
  } catch(err) {
    logError_("sendCriticalAlert_: " + err.message);
  }
}

function logError_(msg) {
  console.error("[RectifierSystem] " + msg);
}
