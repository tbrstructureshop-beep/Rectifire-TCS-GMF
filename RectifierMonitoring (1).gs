/**
 * ============================================================
 *  RECTIFIER MONITORING SYSTEM — Google Apps Script
 *  Electroplating Facility | Specialized Process Division
 * ============================================================
 *
 *  SPREADSHEET ID sudah terhubung ke:
 *  https://docs.google.com/spreadsheets/d/1tF7ukyJjQ1vL1Ov9EtLI1d1EqWLzay9jp4Y9x262EHY
 *
 *  CARA PASANG:
 *  1. Buka Google Sheet di atas → Extensions → Apps Script
 *  2. Hapus isi editor, paste SELURUH file ini
 *  3. Klik Save (Ctrl+S), lalu jalankan onOpen() sekali
 *  4. Reload spreadsheet → menu "⚡ Rectifier System" akan muncul
 *  5. Jalankan setupTriggers() untuk aktifkan otomatisasi
 * ============================================================
 */

// ── SPREADSHEET TARGET ─────────────────────────────────────────
// Karena script dipasang langsung di dalam spreadsheet ini,
// SpreadsheetApp.getActiveSpreadsheet() sudah otomatis terhubung.
// SPREADSHEET_ID di bawah dipakai hanya untuk referensi / cross-script.
const SPREADSHEET_ID = "1tF7ukyJjQ1vL1Ov9EtLI1d1EqWLzay9jp4Y9x262EHY";

// Selalu gunakan helper ini — otomatis pakai active sheet saat dijalankan
// dari dalam spreadsheet, atau buka via ID jika dijalankan dari luar.
function getSpreadsheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ── NAMA SHEET ─────────────────────────────────────────────────
const SHEET_RECT   = "Rectifier Data";
const SHEET_LOG    = "Logbook";
const SHEET_CAL    = "Calibration Tracker";

// ── BARIS DATA ─────────────────────────────────────────────────
const RECT_DATA_START_ROW = 6;
const RECT_DATA_END_ROW   = 10;
const LOG_START_ROW       = 4;

// ── KOLOM "Rectifier Data" (1-based) ──────────────────────────
const COL_RECT_NO         = 1;   // A  No.
const COL_RECT_ID         = 2;   // B  Rectifier ID
const COL_RECT_NAME       = 3;   // C  Unit Name
const COL_RECT_STATUS     = 4;   // D  Status
const COL_RECT_VMIN       = 5;   // E  Volt Min
const COL_RECT_VMAX       = 6;   // F  Volt Max
const COL_RECT_CMIN       = 7;   // G  Curr Min
const COL_RECT_CMAX       = 8;   // H  Curr Max
const COL_RECT_CALIBNO    = 9;   // I  Calib Reg No
const COL_RECT_CALIBDATE  = 10;  // J  Calibration Date
const COL_RECT_DUEDATE    = 11;  // K  Due Date
const COL_RECT_DAYS       = 12;  // L  Days Remaining (formula)
const COL_RECT_DAYSTATUS  = 13;  // M  Days Status (formula)
const COL_RECT_OVERDUE    = 14;  // N  Overdue? (formula)

// ── KOLOM "Logbook" (1-based) ──────────────────────────────────
const COL_LOG_NO          = 1;   // A
const COL_LOG_DATE        = 2;   // B
const COL_LOG_UNIT        = 3;   // C
const COL_LOG_UNITID      = 4;   // D
const COL_LOG_SEVERITY    = 5;   // E
const COL_LOG_DAMAGE      = 6;   // F
const COL_LOG_REPORTER    = 7;   // G
const COL_LOG_ACTION      = 8;   // H
const COL_LOG_STATUS      = 9;   // I

// ── EMAIL ALERT TARGET ─────────────────────────────────────────
// Ganti jika ingin kirim ke email lain, atau tambah beberapa dipisah koma
const ALERT_EMAIL = Session.getActiveUser().getEmail();


// ══════════════════════════════════════════════════════════════
//  MENU
// ══════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Rectifier System")
    .addItem("🔄 Sync Dashboard Summary",      "syncDashboard")
    .addItem("📋 Add Log Entry",               "showAddLogSidebar")
    .addItem("✏️  Edit Rectifier Unit",          "showEditRectSidebar")
    .addSeparator()
    .addItem("📧 Send Overdue Alert Email",     "sendOverdueAlertEmail")
    .addItem("📊 Refresh Calibration Tracker",  "refreshCalibrationTracker")
    .addSeparator()
    .addItem("⚙️  Setup Auto Triggers (1x)",    "setupTriggers")
    .addItem("🗑️  Hapus Semua Triggers",         "removeAllTriggers")
    .addSeparator()
    .addItem("ℹ️  Info Spreadsheet ID",          "showSpreadsheetInfo")
    .addToUi();
}


// ══════════════════════════════════════════════════════════════
//  INFO SPREADSHEET
// ══════════════════════════════════════════════════════════════
function showSpreadsheetInfo() {
  const ss = getSpreadsheet_();
  SpreadsheetApp.getUi().alert(
    "ℹ️ Info Spreadsheet",
    `Nama   : ${ss.getName()}\nID      : ${ss.getId()}\nURL    : ${ss.getUrl()}\nSheets : ${ss.getSheets().map(s => s.getName()).join(", ")}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════
//  SYNC DASHBOARD SUMMARY
// ══════════════════════════════════════════════════════════════
/**
 * Menghitung ulang KPI dan menulis ke baris 4 di "Rectifier Data".
 * Dipanggil otomatis setiap jam (trigger) dan setiap onEdit status/due date.
 */
function syncDashboard() {
  const ss  = getSpreadsheet_();
  const wsR = ss.getSheetByName(SHEET_RECT);
  if (!wsR) { logError_("Sheet tidak ditemukan: " + SHEET_RECT); return; }

  const numRows = RECT_DATA_END_ROW - RECT_DATA_START_ROW + 1;
  const data    = wsR.getRange(RECT_DATA_START_ROW, 1, numRows, 14).getValues();

  let total = 0, inService = 0, unservice = 0, overdue = 0;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  data.forEach(row => {
    const id = row[COL_RECT_ID - 1];
    if (!id) return;
    total++;

    const status = String(row[COL_RECT_STATUS - 1]).trim().toLowerCase();
    if (status === "in service") inService++;
    else unservice++;

    const dueRaw = row[COL_RECT_DUEDATE - 1];
    if (dueRaw) {
      const dueDate = (dueRaw instanceof Date) ? dueRaw : new Date(dueRaw);
      dueDate.setHours(0, 0, 0, 0);
      if (dueDate < today) overdue++;
    }
  });

  // Tulis ke merged KPI cells (row 4)
  wsR.getRange("A4").setValue(total);
  wsR.getRange("D4").setValue(inService);
  wsR.getRange("G4").setValue(unservice);
  wsR.getRange("J4").setValue(overdue);
  wsR.getRange("M3").setValue(
    "Updated: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
  );

  ss.toast(
    `✅ Sync selesai — Total: ${total} | In Service: ${inService} | Unservice: ${unservice} | Overdue: ${overdue}`,
    "Dashboard Synced", 5
  );
}


// ══════════════════════════════════════════════════════════════
//  SIDEBAR: ADD LOG ENTRY
// ══════════════════════════════════════════════════════════════
function showAddLogSidebar() {
  const ss  = getSpreadsheet_();
  const wsR = ss.getSheetByName(SHEET_RECT);
  const units = [];

  for (let r = RECT_DATA_START_ROW; r <= RECT_DATA_END_ROW; r++) {
    const id   = wsR.getRange(r, COL_RECT_ID).getValue();
    const name = wsR.getRange(r, COL_RECT_NAME).getValue();
    if (id) units.push({ id: String(id), name: String(name) });
  }

  const unitOptions = units.map(u =>
    `<option value="${u.id}">${u.name} (${u.id})</option>`
  ).join("");

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html>
<head>
<meta charset="UTF-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Arial,sans-serif;font-size:13px;padding:14px;background:#f6f8fa;color:#24292f}
  h3{color:#1F6FEB;font-size:14px;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #1F6FEB}
  label{display:block;font-size:11px;font-weight:700;color:#57606a;margin:10px 0 3px;text-transform:uppercase;letter-spacing:.5px}
  select,input,textarea{width:100%;padding:7px 9px;border:1px solid #d0d7de;border-radius:6px;font-size:12px;font-family:Arial;background:#fff;color:#24292f}
  select:focus,input:focus,textarea:focus{outline:none;border-color:#1F6FEB;box-shadow:0 0 0 3px rgba(31,111,235,.15)}
  textarea{min-height:60px;resize:vertical}
  .row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  .btn{margin-top:14px;width:100%;padding:9px;background:#1F6FEB;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600}
  .btn:hover{background:#1558c0}
  .msg{margin-top:8px;font-size:12px;padding:6px 10px;border-radius:5px}
  .msg.ok{background:#d2f4dc;color:#1a7f37}
  .msg.err{background:#fddcdc;color:#cf222e}
  .badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700}
</style>
</head>
<body>
<h3>📋 Add Log Entry</h3>

<label>Rectifier Unit</label>
<select id="unit">${unitOptions}</select>

<div class="row2">
  <div>
    <label>Tanggal</label>
    <input type="date" id="date" value="${todayStr}"/>
  </div>
  <div>
    <label>Severity</label>
    <select id="severity">
      <option value="Critical">🔴 Critical</option>
      <option value="Minor" selected>🟡 Minor</option>
      <option value="Resolved">🟢 Resolved</option>
    </select>
  </div>
</div>

<label>Detail Kerusakan</label>
<textarea id="damage" placeholder="Jelaskan kerusakan atau masalah yang terjadi..."></textarea>

<label>Dilaporkan Oleh</label>
<input type="text" id="reporter" placeholder="Nama / Employee ID"/>

<label>Tindakan yang Diambil</label>
<textarea id="action" placeholder="Tindakan yang sudah atau akan diambil..." style="min-height:48px"></textarea>

<label>Status</label>
<select id="status">
  <option value="Minor">Terbuka (Minor)</option>
  <option value="Critical">Critical – Perlu Tindakan Segera</option>
  <option value="Resolved">Resolved – Sudah Diselesaikan</option>
</select>

<button class="btn" onclick="submitLog()">💾 Simpan Log Entry</button>
<div id="msg"></div>

<script>
function submitLog() {
  const damage = document.getElementById('damage').value.trim();
  if (!damage) {
    showMsg('Detail kerusakan wajib diisi.', 'err'); return;
  }
  const reporter = document.getElementById('reporter').value.trim();
  if (!reporter) {
    showMsg('Nama pelapor wajib diisi.', 'err'); return;
  }
  const unitSel = document.getElementById('unit');
  const payload = {
    unitId:   unitSel.value,
    unitName: unitSel.options[unitSel.selectedIndex].text,
    date:     document.getElementById('date').value,
    severity: document.getElementById('severity').value,
    damage,
    reporter,
    action:   document.getElementById('action').value.trim(),
    status:   document.getElementById('status').value
  };
  document.querySelector('.btn').disabled = true;
  showMsg('Menyimpan...', '');
  google.script.run
    .withSuccessHandler(r => {
      showMsg('✅ ' + r, 'ok');
      document.getElementById('damage').value = '';
      document.getElementById('reporter').value = '';
      document.getElementById('action').value = '';
      document.querySelector('.btn').disabled = false;
    })
    .withFailureHandler(e => {
      showMsg('Error: ' + e.message, 'err');
      document.querySelector('.btn').disabled = false;
    })
    .addLogEntryFromSidebar(payload);
}
function showMsg(text, type) {
  const el = document.getElementById('msg');
  el.className = 'msg ' + type;
  el.textContent = text;
}
</script>
</body></html>`)
    .setTitle("📋 Add Log Entry")
    .setWidth(350);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Dipanggil oleh sidebar — append baris baru ke sheet Logbook.
 */
function addLogEntryFromSidebar(payload) {
  const ss  = getSpreadsheet_();
  const wsL = ss.getSheetByName(SHEET_LOG);
  if (!wsL) throw new Error("Sheet Logbook tidak ditemukan.");

  const lastRow = wsL.getLastRow();
  const nextRow = Math.max(LOG_START_ROW, lastRow + 1);
  const entryNo = nextRow - LOG_START_ROW + 1;

  // Pisahkan unit name dari "Name (ID)" jika perlu
  const match    = payload.unitName.match(/^(.*?)\s*\(([^)]+)\)/);
  const unitName = match ? match[1].trim() : payload.unitName;

  wsL.getRange(nextRow, COL_LOG_NO).setValue(entryNo);

  const dateCell = wsL.getRange(nextRow, COL_LOG_DATE);
  dateCell.setValue(new Date(payload.date));
  dateCell.setNumberFormat("DD-MMM-YYYY");

  wsL.getRange(nextRow, COL_LOG_UNIT).setValue(unitName);
  wsL.getRange(nextRow, COL_LOG_UNITID).setValue(payload.unitId);
  wsL.getRange(nextRow, COL_LOG_SEVERITY).setValue(payload.severity);
  wsL.getRange(nextRow, COL_LOG_DAMAGE).setValue(payload.damage);
  wsL.getRange(nextRow, COL_LOG_REPORTER).setValue(payload.reporter);
  wsL.getRange(nextRow, COL_LOG_ACTION).setValue(payload.action);
  wsL.getRange(nextRow, COL_LOG_STATUS).setValue(payload.status);

  formatLogRow_(wsL, nextRow, payload.severity);

  if (payload.severity === "Critical") {
    sendCriticalAlert_(unitName, payload.unitId, payload.damage, payload.reporter);
  }

  return `Log #${entryNo} berhasil ditambahkan untuk ${unitName}`;
}


// ══════════════════════════════════════════════════════════════
//  SIDEBAR: EDIT RECTIFIER UNIT
// ══════════════════════════════════════════════════════════════
function showEditRectSidebar() {
  const ss  = getSpreadsheet_();
  const wsR = ss.getSheetByName(SHEET_RECT);
  const units = [];

  for (let r = RECT_DATA_START_ROW; r <= RECT_DATA_END_ROW; r++) {
    const row = wsR.getRange(r, 1, 1, 14).getValues()[0];
    const id  = row[COL_RECT_ID - 1];
    if (!id) continue;
    units.push({
      row,
      sheetRow:  r,
      id:        String(id),
      name:      String(row[COL_RECT_NAME - 1]),
      status:    String(row[COL_RECT_STATUS - 1]),
      voltMin:   row[COL_RECT_VMIN - 1],
      voltMax:   row[COL_RECT_VMAX - 1],
      currMin:   row[COL_RECT_CMIN - 1],
      currMax:   row[COL_RECT_CMAX - 1],
      calibNo:   String(row[COL_RECT_CALIBNO - 1]),
      calibDate: row[COL_RECT_CALIBDATE - 1] instanceof Date
                   ? Utilities.formatDate(row[COL_RECT_CALIBDATE - 1], Session.getScriptTimeZone(), "yyyy-MM-dd")
                   : String(row[COL_RECT_CALIBDATE - 1]),
      dueDate:   row[COL_RECT_DUEDATE - 1] instanceof Date
                   ? Utilities.formatDate(row[COL_RECT_DUEDATE - 1], Session.getScriptTimeZone(), "yyyy-MM-dd")
                   : String(row[COL_RECT_DUEDATE - 1]),
      damage:    String(row[13] || ""),  // kolom N setelah formula — ambil dari baris tambahan jika ada
    });
  }

  const unitsJson = JSON.stringify(units);

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html><html>
<head>
<meta charset="UTF-8">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Arial,sans-serif;font-size:13px;padding:14px;background:#f6f8fa;color:#24292f}
  h3{color:#1F6FEB;font-size:14px;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #1F6FEB}
  label{display:block;font-size:11px;font-weight:700;color:#57606a;margin:10px 0 3px;text-transform:uppercase;letter-spacing:.5px}
  select,input{width:100%;padding:7px 9px;border:1px solid #d0d7de;border-radius:6px;font-size:12px;font-family:Arial;background:#fff}
  select:focus,input:focus{outline:none;border-color:#1F6FEB;box-shadow:0 0 0 3px rgba(31,111,235,.15)}
  .row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  .btn{margin-top:14px;width:100%;padding:9px;background:#1F6FEB;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600}
  .btn:hover{background:#1558c0}
  .msg{margin-top:8px;font-size:12px;padding:6px 10px;border-radius:5px}
  .msg.ok{background:#d2f4dc;color:#1a7f37}
  .msg.err{background:#fddcdc;color:#cf222e}
  .divider{border:none;border-top:1px solid #d0d7de;margin:12px 0}
  .unit-card{background:#fff;border:1px solid #d0d7de;border-radius:8px;padding:10px;margin-bottom:4px}
  .unit-id{font-size:10px;color:#8b949e;font-family:monospace}
  .unit-name{font-size:13px;font-weight:700}
</style>
</head>
<body>
<h3>✏️ Edit Rectifier Unit</h3>

<label>Pilih Unit</label>
<select id="unitSel" onchange="loadUnit(this.value)">
  <option value="">— Pilih rectifier unit —</option>
</select>

<div id="unitInfo" style="display:none;background:#EEF2FF;border:1px solid #c8d8ff;border-radius:6px;padding:8px 10px;margin-top:8px;font-size:12px;"></div>

<hr class="divider">

<div id="form" style="display:none">
  <label>Status</label>
  <select id="eStatus">
    <option value="In Service">✅ In Service</option>
    <option value="Unservice">🔴 Unservice</option>
  </select>

  <div class="row2">
    <div><label>Volt Min (VDC)</label><input type="number" id="eVMin" step="0.1"/></div>
    <div><label>Volt Max (VDC)</label><input type="number" id="eVMax" step="0.1"/></div>
  </div>
  <div class="row2">
    <div><label>Curr Min (A)</label><input type="number" id="eCMin"/></div>
    <div><label>Curr Max (A)</label><input type="number" id="eCMax"/></div>
  </div>

  <label>Calib. Reg. No.</label>
  <input type="text" id="eCalibNo"/>

  <div class="row2">
    <div><label>Tanggal Kalibrasi</label><input type="date" id="eCalibDate"/></div>
    <div><label>Due Date Kalibrasi</label><input type="date" id="eDueDate"/></div>
  </div>

  <button class="btn" onclick="saveUnit()">💾 Simpan Perubahan</button>
</div>

<div id="msg"></div>

<script>
const units = ${unitsJson};
const sel   = document.getElementById('unitSel');
units.forEach(u => {
  const o = document.createElement('option');
  o.value       = u.sheetRow;
  o.textContent = u.name + ' · ' + u.id;
  sel.appendChild(o);
});

function loadUnit(row) {
  const form = document.getElementById('form');
  const info = document.getElementById('unitInfo');
  if (!row) { form.style.display='none'; info.style.display='none'; return; }
  const u = units.find(x => x.sheetRow == row);
  if (!u) return;

  info.style.display  = 'block';
  info.innerHTML      = '<span style="font-family:monospace;font-weight:700;color:#1F6FEB">' + u.id + '</span> &nbsp;|&nbsp; ' + u.name + ' &nbsp;|&nbsp; Status: <strong>' + u.status + '</strong>';

  document.getElementById('eStatus').value    = u.status;
  document.getElementById('eVMin').value       = u.voltMin;
  document.getElementById('eVMax').value       = u.voltMax;
  document.getElementById('eCMin').value       = u.currMin;
  document.getElementById('eCMax').value       = u.currMax;
  document.getElementById('eCalibNo').value    = u.calibNo;
  document.getElementById('eCalibDate').value  = (u.calibDate || '').substring(0, 10);
  document.getElementById('eDueDate').value    = (u.dueDate   || '').substring(0, 10);
  form.style.display = 'block';
}

function saveUnit() {
  const row = parseInt(document.getElementById('unitSel').value);
  if (!row) { showMsg('Pilih unit terlebih dahulu.', 'err'); return; }
  const payload = {
    sheetRow:  row,
    status:    document.getElementById('eStatus').value,
    voltMin:   parseFloat(document.getElementById('eVMin').value) || 0,
    voltMax:   parseFloat(document.getElementById('eVMax').value) || 0,
    currMin:   parseFloat(document.getElementById('eCMin').value) || 0,
    currMax:   parseFloat(document.getElementById('eCMax').value) || 0,
    calibNo:   document.getElementById('eCalibNo').value.trim(),
    calibDate: document.getElementById('eCalibDate').value,
    dueDate:   document.getElementById('eDueDate').value,
  };
  document.querySelector('.btn').disabled = true;
  showMsg('Menyimpan...', '');
  google.script.run
    .withSuccessHandler(r => {
      showMsg('✅ ' + r, 'ok');
      document.querySelector('.btn').disabled = false;
    })
    .withFailureHandler(e => {
      showMsg('Error: ' + e.message, 'err');
      document.querySelector('.btn').disabled = false;
    })
    .saveRectifierFromSidebar(payload);
}

function showMsg(text, type) {
  const el = document.getElementById('msg');
  el.className = 'msg ' + (type || '');
  el.textContent = text;
}
</script>
</body></html>`)
    .setTitle("✏️ Edit Rectifier Unit")
    .setWidth(360);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Dipanggil oleh sidebar — simpan perubahan ke baris rectifier.
 */
function saveRectifierFromSidebar(payload) {
  const ss  = getSpreadsheet_();
  const wsR = ss.getSheetByName(SHEET_RECT);
  if (!wsR) throw new Error("Sheet Rectifier Data tidak ditemukan.");

  const r = payload.sheetRow;

  wsR.getRange(r, COL_RECT_STATUS).setValue(payload.status);
  wsR.getRange(r, COL_RECT_VMIN).setValue(payload.voltMin);
  wsR.getRange(r, COL_RECT_VMAX).setValue(payload.voltMax);
  wsR.getRange(r, COL_RECT_CMIN).setValue(payload.currMin);
  wsR.getRange(r, COL_RECT_CMAX).setValue(payload.currMax);
  wsR.getRange(r, COL_RECT_CALIBNO).setValue(payload.calibNo);

  if (payload.calibDate) {
    const cd = wsR.getRange(r, COL_RECT_CALIBDATE);
    cd.setValue(new Date(payload.calibDate));
    cd.setNumberFormat("DD-MMM-YYYY");
  }
  if (payload.dueDate) {
    const dd = wsR.getRange(r, COL_RECT_DUEDATE);
    dd.setValue(new Date(payload.dueDate));
    dd.setNumberFormat("DD-MMM-YYYY");
  }

  syncDashboard();
  return `Baris ${r} (${wsR.getRange(r, COL_RECT_NAME).getValue()}) berhasil diupdate.`;
}


// ══════════════════════════════════════════════════════════════
//  REFRESH CALIBRATION TRACKER
// ══════════════════════════════════════════════════════════════
function refreshCalibrationTracker() {
  const ss  = getSpreadsheet_();
  const wsC = ss.getSheetByName(SHEET_CAL);
  if (!wsC) {
    ss.toast("Sheet Calibration Tracker tidak ditemukan.", "Error", 4);
    return;
  }
  // Touch helper cell untuk force recalculate cross-sheet formulas
  const h = wsC.getRange("H1");
  h.setValue(new Date());
  h.setNumberFormat("DD-MMM-YYYY HH:mm");
  ss.toast("Calibration Tracker direfresh.", "Selesai", 3);
}


// ══════════════════════════════════════════════════════════════
//  EMAIL ALERT — OVERDUE & DUE SOON
// ══════════════════════════════════════════════════════════════
function sendOverdueAlertEmail() {
  const ss    = getSpreadsheet_();
  const wsR   = ss.getSheetByName(SHEET_RECT);
  const today = new Date(); today.setHours(0, 0, 0, 0);

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

    const dueDate = (dueRaw instanceof Date) ? new Date(dueRaw) : new Date(dueRaw);
    dueDate.setHours(0, 0, 0, 0);
    const days = Math.round((dueDate - today) / 86400000);

    if (days < 0)        overdue.push({ id, name, status, dueDate, days });
    else if (days <= 30) dueSoon.push({ id, name, status, dueDate, days });
  }

  if (overdue.length === 0 && dueSoon.length === 0) {
    ss.toast("✅ Tidak ada kalibrasi overdue atau due soon.", "All Clear", 5);
    return;
  }

  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM yyyy");

  let body = `
<div style="font-family:Arial,sans-serif;max-width:660px;margin:auto">
  <div style="background:#0D1117;color:#E6EDF3;padding:20px 24px;border-radius:10px 10px 0 0">
    <h2 style="margin:0;font-size:18px">⚡ Rectifier Calibration Alert</h2>
    <p style="margin:5px 0 0;color:#8B949E;font-size:12px">
      Electroplating Facility — Specialized Process Division<br>
      Spreadsheet: <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}" style="color:#388BFD">Buka Google Sheet</a>
    </p>
  </div>
  <div style="background:#f8f9fa;padding:20px 24px;border-radius:0 0 10px 10px;border:1px solid #dee2e6">`;

  if (overdue.length > 0) {
    body += `
    <h3 style="color:#F85149;margin:0 0 10px">🔴 OVERDUE — ${overdue.length} unit</h3>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px">
      <thead>
        <tr style="background:#FDDCDC">
          <th style="padding:8px;text-align:left">ID</th>
          <th style="padding:8px;text-align:left">Unit</th>
          <th style="padding:8px;text-align:left">Status</th>
          <th style="padding:8px;text-align:left">Due Date</th>
          <th style="padding:8px;text-align:left">Overdue</th>
        </tr>
      </thead>
      <tbody>
        ${overdue.map(u => `
        <tr style="border-bottom:1px solid #eee">
          <td style="padding:8px;font-weight:700">${u.id}</td>
          <td style="padding:8px">${u.name}</td>
          <td style="padding:8px">${u.status}</td>
          <td style="padding:8px">${fmt(u.dueDate)}</td>
          <td style="padding:8px;color:#F85149;font-weight:700">${Math.abs(u.days)} hari</td>
        </tr>`).join("")}
      </tbody>
    </table>`;
  }

  if (dueSoon.length > 0) {
    body += `
    <h3 style="color:#D29922;margin:0 0 10px">🟡 DUE SOON (≤30 hari) — ${dueSoon.length} unit</h3>
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead>
        <tr style="background:#FFF3CC">
          <th style="padding:8px;text-align:left">ID</th>
          <th style="padding:8px;text-align:left">Unit</th>
          <th style="padding:8px;text-align:left">Due Date</th>
          <th style="padding:8px;text-align:left">Sisa Hari</th>
        </tr>
      </thead>
      <tbody>
        ${dueSoon.map(u => `
        <tr style="border-bottom:1px solid #eee">
          <td style="padding:8px;font-weight:700">${u.id}</td>
          <td style="padding:8px">${u.name}</td>
          <td style="padding:8px">${fmt(u.dueDate)}</td>
          <td style="padding:8px;color:#D29922;font-weight:700">${u.days} hari</td>
        </tr>`).join("")}
      </tbody>
    </table>`;
  }

  body += `
    <p style="margin-top:18px;font-size:11px;color:#888;border-top:1px solid #eee;padding-top:10px">
      Generated: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")} &nbsp;|&nbsp;
      Rectifier Monitoring System v1.0 &nbsp;|&nbsp; Electroplating Facility
    </p>
  </div>
</div>`;

  MailApp.sendEmail({
    to:       ALERT_EMAIL,
    subject:  `[Rectifier Alert] ${overdue.length} Overdue · ${dueSoon.length} Due Soon — ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy")}`,
    htmlBody: body
  });

  ss.toast(`📧 Email alert dikirim ke ${ALERT_EMAIL}`, "Email Sent", 5);
}


// ══════════════════════════════════════════════════════════════
//  TRIGGER SETUP
// ══════════════════════════════════════════════════════════════
/**
 * Pasang dua trigger:
 *   • Tiap jam   → syncDashboard
 *   • Setiap hari jam 08:00 → sendOverdueAlertEmail
 *
 * Jalankan SEKALI dari menu atau editor.
 */
function setupTriggers() {
  removeAllTriggers();

  ScriptApp.newTrigger("syncDashboard")
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger("sendOverdueAlertEmail")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  getSpreadsheet_().toast(
    "Triggers aktif: sync tiap jam + alert email setiap hari jam 08:00.",
    "Triggers Ready", 6
  );
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}


// ══════════════════════════════════════════════════════════════
//  ON EDIT TRIGGER (Simple Trigger — otomatis, tanpa instalasi)
// ══════════════════════════════════════════════════════════════
/**
 * Otomatis sync dashboard jika:
 *   - Sheet "Rectifier Data", kolom Status (D) atau Due Date (K) diubah
 *   - Sheet "Logbook", row baru ditambahkan
 */
function onEdit(e) {
  if (!e) return;
  const sheet  = e.range.getSheet();
  const sName  = sheet.getName();
  const row    = e.range.getRow();
  const col    = e.range.getColumn();

  if (sName === SHEET_RECT) {
    if (row >= RECT_DATA_START_ROW && row <= RECT_DATA_END_ROW) {
      if (col === COL_RECT_STATUS || col === COL_RECT_DUEDATE) {
        syncDashboard();
      }
    }
  }

  if (sName === SHEET_LOG) {
    // Auto-format severity dan status jika user input langsung ke sheet
    if (row >= LOG_START_ROW && col === COL_LOG_SEVERITY) {
      const val = String(e.value || "");
      if (val) formatLogRow_(sheet, row, val);
    }
  }
}


// ══════════════════════════════════════════════════════════════
//  HELPER FUNCTIONS
// ══════════════════════════════════════════════════════════════
function formatLogRow_(sheet, rowNum, severity) {
  const palette = {
    "Critical": { bg: "#FDDCDC", fg: "#CF222E" },
    "Minor":    { bg: "#FFF3CC", fg: "#9A6700" },
    "Resolved": { bg: "#D2F4DC", fg: "#1A7F37" },
  };
  const col = palette[severity] || { bg: "#FFFFFF", fg: "#24292F" };

  // Warnai sel Severity (E) dan Status (I)
  [COL_LOG_SEVERITY, COL_LOG_STATUS].forEach(c => {
    const cell = sheet.getRange(rowNum, c);
    cell.setBackground(col.bg);
    cell.setFontColor(col.fg);
    cell.setFontWeight("bold");
  });

  // Border seluruh baris
  sheet.getRange(rowNum, 1, 1, 9).setBorder(
    true, true, true, true, true, true,
    "#D0D7DE", SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.setRowHeight(rowNum, 42);
}

function sendCriticalAlert_(unitName, unitId, damage, reporter) {
  try {
    MailApp.sendEmail({
      to: ALERT_EMAIL,
      subject: `🔴 CRITICAL — ${unitName} (${unitId}) — Kerusakan Dilaporkan`,
      htmlBody: `
<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto">
  <div style="background:#CF222E;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0">
    <h2 style="margin:0;font-size:17px">🔴 Critical Damage Alert</h2>
    <p style="margin:4px 0 0;font-size:12px;opacity:.85">Rectifier Monitoring System — Electroplating Facility</p>
  </div>
  <div style="background:#fff;padding:18px 20px;border:1px solid #FDDCDC;border-radius:0 0 10px 10px">
    <table style="font-size:13px;border-collapse:collapse;width:100%">
      <tr><td style="padding:5px 0;color:#57606a;width:130px">Unit</td><td><strong>${unitName} (${unitId})</strong></td></tr>
      <tr><td style="padding:5px 0;color:#57606a">Dilaporkan Oleh</td><td><strong>${reporter}</strong></td></tr>
      <tr><td style="padding:5px 0;color:#57606a">Waktu</td><td>${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")}</td></tr>
    </table>
    <div style="background:#FFF5F5;border:1px solid #FDDCDC;border-radius:6px;padding:12px;margin-top:12px">
      <strong style="font-size:11px;text-transform:uppercase;color:#CF222E">Detail Kerusakan</strong><br>
      <span style="font-size:13px;line-height:1.6">${damage}</span>
    </div>
    <p style="margin-top:14px">
      <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}" style="background:#1F6FEB;color:#fff;padding:8px 14px;border-radius:6px;text-decoration:none;font-size:12px">
        📊 Buka Spreadsheet
      </a>
    </p>
    <p style="margin-top:14px;font-size:11px;color:#888">Rectifier Monitoring System v1.0</p>
  </div>
</div>`
    });
  } catch(err) {
    logError_("sendCriticalAlert_: " + err.message);
  }
}

function logError_(msg) {
  console.error("[RectifierGAS] " + msg);
}
