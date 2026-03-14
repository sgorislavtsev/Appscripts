// ═══════════════════════════════════════════════
//  HQ SCHEDULE HUB — Google Apps Script
//  Deploy as Web App: Execute as Me, Anyone access
// ═══════════════════════════════════════════════

const DATA_SHEET   = 'Schedules';
const CONFIG_SHEET = 'Config';

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  var action = p.action || '';

  try {
    if (action === 'ping') {
      return json({ ok: true, message: 'HQ Hub is live' });
    }
    if (action === 'pushSchedule' && p.payload) {
      var body2 = JSON.parse(p.payload);
      body2.action = 'pushSchedule';
      return json(pushSchedule(body2));
    }
    if (action === 'getAllStores') {
      return json(getAllStores());
    }
    if (action === 'getConfig') {
      return json(getConfig());
    }
    if (action === 'getStoreDirectory') {
      return json(getStoreDirectory());
    }
    if (action === 'getHandoff' && p.storeId) {
      return json(getHandoff(p.storeId));
    }
    if (action === 'getThresholds') {
      return json(getThresholds(p.email || ''));
    }
    if (action === 'getUserPrefs') {
      return json(getUserPrefs(p.email || ''));
    }
    return json({ error: 'Unknown action: ' + action });
  } catch(err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';

    if (action === 'pushSchedule') {
      return json(pushSchedule(body));
    }
    if (action === 'saveConfig') {
      return json(saveConfig(body.config));
    }
    if (action === 'pushHandoff') {
      return json(pushHandoff(body));
    }
    if (action === 'saveThresholds') {
      return json(saveThresholds(body.email || '', body.thresholds || {}));
    }
    if (action === 'saveUserPrefs') {
      return json(saveUserPrefs(body.email || '', body.prefs || {}));
    }
    return json({ error: 'Unknown action: ' + action });
  } catch(err) {
    return json({ error: err.message });
  }
}

// ── Store a schedule pushed from a store manager ──────────────
function pushSchedule(body) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, DATA_SHEET);

  // Headers on row 1 — department column added after storeId
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['storeId','department','storeName','region','district',
                     'weekOf','headcount','totalHours','weeklyBudget',
                     'lastPushed','payload']);
  } else {
    // ── Migrate old sheets that lack the 'department' column ──
    // Old header row: storeId(0) storeName(1) region(2) ...
    // New header row: storeId(0) department(1) storeName(2) region(3) ...
    // If col B header is not 'department', insert the column now.
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headerRow[1] !== 'department') {
      sheet.insertColumnBefore(2);
      sheet.getRange(1, 2).setValue('department');
      // Existing data rows already have '' in the new column (blank after insert)
    }
  }

  var storeId    = String(body.storeId    || '');
  var storeName  = String(body.storeName  || '');
  var payload    = body.payload || {};

  // Department: prefer top-level body field, fall back to payload field
  var department = String(body.department || payload.department || '').trim();

  // Compute summary stats from payload
  var weekOf = payload.dateFrom || '';

  // Check both 'weeklyBudget' and 'budget' keys
  var budget = Number(payload.weeklyBudget || payload.budget || 0);

  var members  = payload.teamMembers || [];
  var managers = payload.managers    || {};
  var schedule = payload.schedule    || {};

  // Count hours from slot-based schedule.
  // Format: schedule[day][member] = { "07:15": "s", "07:30": "s", ... }
  // Slot duration is auto-detected from the minute values in the time keys.
  // Supports both 15-min slots (0.25h) and 30-min slots (0.5h).
  var codes = payload.codes || [];
  var noCountCodes = {};
  codes.forEach(function(c) { if (c.countHours === false) noCountCodes[c.key] = true; });
  var hasCodes = codes.length > 0;

  // Detect slot size by finding the minimum gap between minute values across all schedule keys.
  var minuteSet = {};
  Object.keys(schedule).forEach(function(d) {
    Object.keys(schedule[d]).forEach(function(mem) {
      var slotMap = schedule[d][mem];
      if (!slotMap || typeof slotMap !== 'object') return;
      Object.keys(slotMap).forEach(function(t) {
        var parts = t.split(':');
        if (parts.length === 2) minuteSet[parseInt(parts[1], 10)] = true;
      });
    });
  });
  var uniqueMinutes = Object.keys(minuteSet).map(Number).sort(function(a,b){return a-b;});
  var slotHours = 0.5; // default: 30-min slots
  if (uniqueMinutes.length >= 2) {
    var minGap = 60;
    for (var mi = 1; mi < uniqueMinutes.length; mi++) {
      var gap = uniqueMinutes[mi] - uniqueMinutes[mi-1];
      if (gap > 0 && gap < minGap) minGap = gap;
    }
    slotHours = minGap / 60;
  } else if (uniqueMinutes.length === 1) {
    slotHours = 0.25; // only one minute value seen; assume 15-min
  }

  var totalHours = 0;
  var DAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  members.forEach(function(m) {
    DAYS.forEach(function(d) {
      var slots = (schedule[d] || {})[m];
      if (!slots || typeof slots !== 'object' || Array.isArray(slots)) return;
      Object.keys(slots).forEach(function(time) {
        var code = slots[time];
        if (!hasCodes || !noCountCodes[code]) {
          totalHours += slotHours;
        }
      });
    });
  });

  // Headcount = non-manager members
  var headcount = members.filter(function(m){ return !managers[m]; }).length;

  var now = new Date().toISOString();

  // ── Find existing row for this storeId + department combo and overwrite ──
  // This is the key change: composite key prevents departments from overwriting each other.
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    var rowStoreId = String(data[i][0]);
    var rowDept    = String(data[i][1]);
    if (rowStoreId === storeId && rowDept === department) {
      rowIdx = i + 1;
      break;
    }
  }

  // Column order: storeId | department | storeName | region | district |
  //               weekOf | headcount | totalHours | weeklyBudget | lastPushed | payload
  var rowData = [
    storeId,
    department,
    storeName,
    (payload.storeInfo && payload.storeInfo.region) || '',
    '',  // district — managed via dashboard Config
    weekOf,
    headcount,
    Math.round(totalHours * 10) / 10,
    budget,
    now,
    JSON.stringify(payload)
  ];

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  var msg = department
    ? 'Schedule saved for Store ' + storeId + ' — ' + department
    : 'Schedule saved for Store ' + storeId;
  return { ok: true, message: msg };
}

// ── Return all store records to the Dashboard ─────────────────
// Rows are keyed by storeId+department. This function merges all department
// rows for a given storeId into one store object, rolling up headcount,
// hours, and budget. The `departments` array carries per-dept detail.
function getAllStores() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, stores: [] };

  var data = sheet.getDataRange().getValues();
  var cfg  = loadConfigData();

  // Detect column layout: new layout has 'department' at index 1.
  // Old layout (no department col): storeId(0) storeName(1) region(2) district(3)
  //                                  weekOf(4) headcount(5) totalHours(6) weeklyBudget(7) lastPushed(8) payload(9)
  // New layout (with department col): storeId(0) department(1) storeName(2) region(3) district(4)
  //                                  weekOf(5) headcount(6) totalHours(7) weeklyBudget(8) lastPushed(9) payload(10)
  var headerRow = data[0];
  var hasDepCol = (headerRow[1] === 'department');

  // Column index helpers
  var C = hasDepCol
    ? { dep:1, name:2, reg:3, dist:4, week:5, hc:6, hrs:7, bgt:8, pushed:9, pay:10 }
    : { dep:-1, name:1, reg:2, dist:3, week:4, hc:5, hrs:6, bgt:7, pushed:8, pay:9 };

  // storeMap keyed by storeId for merge order
  var storeMap  = {};
  var storeOrder = [];

  for (var i = 1; i < data.length; i++) {
    var r       = data[i];
    var storeId = String(r[0] || '').trim();
    if (!storeId) continue;

    var department = hasDepCol ? String(r[C.dep] || '').trim() : '';
    var payObj     = {};
    try { payObj = JSON.parse(r[C.pay] || '{}'); } catch(e) {}

    var weeklyBudget = Number(r[C.bgt] || 0);
    var calcBudget   = Math.round(
      (Number(payObj.forecastSales || 0) * Number(payObj.payrollPct || 0)) / 100
    );
    var headcount  = Number(r[C.hc]     || 0);
    var totalHours = Number(r[C.hrs]    || 0);
    var weekOf     = r[C.week]  || '';
    var lastPushed = r[C.pushed] || '';
    var storeName  = String(r[C.name] || '');
    var region     = String(r[C.reg]  || '');
    var district   = String(r[C.dist] || '');

    var cfgInfo = (cfg.storeMap || {})[storeId] || {};

    var deptEntry = {
      department:   department,
      headcount:    headcount,
      totalHours:   totalHours,
      weeklyBudget: weeklyBudget,
      calcBudget:   calcBudget,
      weekOf:       weekOf,
      lastPushed:   lastPushed,
      payload:      payObj
    };

    if (!storeMap[storeId]) {
      storeOrder.push(storeId);
      storeMap[storeId] = {
        storeId:      storeId,
        storeName:    storeName,
        region:       cfgInfo.region   || region,
        district:     cfgInfo.district || district,
        weekOf:       weekOf,
        headcount:    0,
        totalHours:   0,
        weeklyBudget: 0,
        calcBudget:   0,
        lastPushed:   lastPushed,
        payload:      payObj,   // first/most-recent payload for backward compat
        departments:  []
      };
    }

    var s = storeMap[storeId];
    s.departments.push(deptEntry);

    // Roll up totals
    s.headcount    += headcount;
    s.totalHours    = Math.round((s.totalHours + totalHours) * 10) / 10;
    s.weeklyBudget += weeklyBudget;
    s.calcBudget   += calcBudget;

    // Keep most recent weekOf and lastPushed
    if (!s.weekOf || weekOf > s.weekOf) s.weekOf = weekOf;
    if (!s.lastPushed || lastPushed > s.lastPushed) {
      s.lastPushed = lastPushed;
      s.payload    = payObj;  // canonical payload = most recently pushed dept
    }
  }

  var stores = storeOrder.map(function(id) { return storeMap[id]; });
  return { ok: true, stores: stores };
}

// ── Config (region / district assignments) ────────────────────
function getConfig() {
  return { ok: true, config: loadConfigData() };
}

function saveConfig(config) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, CONFIG_SHEET);
  sheet.clearContents();
  sheet.appendRow(['config']);
  sheet.appendRow([JSON.stringify(config)]);
  return { ok: true };
}

function loadConfigData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { storeMap: {} };
  try {
    var val = sheet.getRange(2, 1).getValue();
    return JSON.parse(val) || { storeMap: {} };
  } catch(e) { return { storeMap: {} }; }
}


// ── Read Branch Manager Contact Sheet → return store directory ──
//  Sheet ID is the BRANCH_MANAGER_SHEET_ID constant below.
//  Set this to the ID from your Google Sheet URL.
var BRANCH_MANAGER_SHEET_ID = '1KTXtPoDFtvZ9Nx35RfhvMmFkzRICd0eabxiIKvOcQQE';
var BRANCH_MANAGER_GID      = '873143388';

// Column indices (0-based) in the Branch Manager Contact sheet
var COL_BRANCH  = 0;   // BRANCH #
var COL_NAME    = 1;   // STATE AND BRANCH NAME
var COL_ADDR    = 3;   // ADDRESS
var COL_CITY    = 4;   // CITY
var COL_STATE   = 5;   // STATE
var COL_ZIP     = 6;   // ZIP
var COL_PHONE   = 9;   // PHONE
var COL_MGR     = 10;  // MANAGER
var COL_MGRCELL = 11;  // MANAGER CELL#
var COL_EMAIL   = 12;  // STORE MANAGER EMAIL
var COL_REGION  = 18;  // REGION
var COL_SCDIREC = 19;  // SC DIRECTOR (used as "area" / DM group)
var COL_DM      = 20;  // DM
var COL_GM      = 21;  // GM
var COL_RVP     = 22;  // RVP

function getStoreDirectory() {
  try {
    var ss = SpreadsheetApp.openById(BRANCH_MANAGER_SHEET_ID);
    // Try to get the sheet by gid
    var sheets = ss.getSheets();
    var ws = null;
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getSheetId()) === BRANCH_MANAGER_GID) {
        ws = sheets[i]; break;
      }
    }
    if (!ws) ws = sheets[0]; // fallback to first sheet

    var data = ws.getDataRange().getValues();
    if (data.length < 2) return { status: 'ok', data: [] };

    var stores = [];
    var seen   = {};  // deduplicate by storeId

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rawId = row[COL_BRANCH];
      if (!rawId && rawId !== 0) continue;
      var storeId = String(rawId).replace(/\.0$/, '').trim();
      if (!storeId) continue;

      var name    = String(row[COL_NAME]    || '').trim();
      var city    = String(row[COL_CITY]    || '').trim();
      var state   = String(row[COL_STATE]   || '').trim();
      var addr    = String(row[COL_ADDR]    || '').trim();
      var phone   = _fmtPhone(row[COL_PHONE]);
      var mgr     = String(row[COL_MGR]     || '').trim();
      var mgrEmail= String(row[COL_EMAIL]   || '').trim();
      var region  = String(row[COL_REGION]  || '').trim();
      var dm      = String(row[COL_DM]      || '').trim();
      var gm      = String(row[COL_GM]      || '').trim();
      var rvp     = String(row[COL_RVP]     || '').trim();

      // Derive area from state or SC Director column
      var area = String(row[COL_SCDIREC] || '').trim() || state;

      // Build display name: "ST - City" style if raw name lacks it
      var displayName = name.length > 2 ? name : (state + ' - ' + city);

      stores.push({
        id:           storeId,
        name:         displayName,
        address:      addr,
        city:         city,
        state:        state,
        phone:        phone,
        manager:      mgr,
        managerEmail: mgrEmail,
        region:       region,
        area:         area,
        dm:           dm,
        dmEmail:      '',   // not in sheet; can be extended
        gm:           gm,
        gmEmail:      '',
        rvp:          rvp,
        rvpEmail:     ''
      });
    }

    return { status: 'ok', data: stores };
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

function _fmtPhone(val) {
  if (!val) return '';
  var s = String(val).replace(/\.0$/, '').trim();
  // If it looks like a raw 10-digit number, format it
  var digits = s.replace(/\D/g, '');
  if (digits.length === 10) {
    return '(' + digits.slice(0,3) + ') ' + digits.slice(3,6) + '-' + digits.slice(6);
  }
  return s;
}

// ── HQ → Scheduler handoff: DM pushes a schedule for a store manager to pick up ──
var HANDOFF_SHEET = 'Handoffs';

function pushHandoff(body) {
  var storeId  = String(body.storeId  || '');
  var pushedBy = String(body.pushedBy || 'HQ Dashboard');
  var payload  = body.payload;
  if (!storeId || !payload) return { ok: false, error: 'Missing storeId or payload' };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, HANDOFF_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['storeId', 'payload', 'pushedAt', 'pushedBy']);
  }

  var now  = new Date().toISOString();
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === storeId) { rowIdx = i + 1; break; }
  }

  var row = [storeId, JSON.stringify(payload), now, pushedBy];
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { ok: true, message: 'Handoff staged for Store ' + storeId };
}

function getHandoff(storeId) {
  storeId = String(storeId);
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HANDOFF_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, handoff: null };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === storeId) {
      var payload  = {};
      var pushedAt = data[i][2];
      var pushedBy = data[i][3];
      try { payload = JSON.parse(data[i][1] || '{}'); } catch(e) {}

      // Clear the row so it can only be picked up once
      sheet.getRange(i + 1, 1, 1, 4).setValues([['', '', '', '']]);

      return { ok: true, handoff: { payload: payload, pushedAt: pushedAt, pushedBy: pushedBy } };
    }
  }
  return { ok: true, handoff: null };
}

// ── Helpers ───────────────────────────────────────────────────
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function timeToHours(t) {
  if (!t) return 0;
  t = t.trim();
  var ampm  = /([ap]m)/i.exec(t);
  var parts = t.replace(/[ap]m/i, '').trim().split(':');
  var h = parseInt(parts[0] || 0), m = parseInt(parts[1] || 0);
  if (ampm) {
    var sfx = ampm[1].toLowerCase();
    if (sfx === 'pm' && h !== 12) h += 12;
    if (sfx === 'am' && h === 12) h = 0;
  }
  return h + m / 60;
}
// ═══════════════════════════════════════════════
//  USER THRESHOLDS  (tab: __THRESHOLDS__)
//  One row per email: email | JSON blob | lastUpdated
// ═══════════════════════════════════════════════

function getThresholds(email) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, '__THRESHOLDS__');
  if (!email) return { ok: false, error: 'Email required' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      try {
        var thr = JSON.parse(data[i][1] || '{}');
        return { ok: true, thresholds: thr };
      } catch(e) {
        return { ok: true, thresholds: {} };
      }
    }
  }
  return { ok: true, thresholds: null }; // null = use client defaults
}

function saveThresholds(email, thresholds) {
  if (!email) return { ok: false, error: 'Email required' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, '__THRESHOLDS__');
  var data  = sheet.getDataRange().getValues();
  var now   = new Date().toISOString();
  var blob  = JSON.stringify(thresholds);

  // Ensure header row
  if (data.length === 0 || data[0][0] !== 'email') {
    sheet.getRange(1, 1, 1, 3).setValues([['email', 'thresholds', 'lastUpdated']]);
    data = sheet.getDataRange().getValues();
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[email, blob, now]]);
      return { ok: true, saved: true };
    }
  }
  // New row
  sheet.appendRow([email, blob, now]);
  return { ok: true, saved: true };
}

// ═══════════════════════════════════════════════
//  USER PREFERENCES  (tab: __USERPREFS__)
//  One row per email: email | JSON blob | lastUpdated
//  Prefs include: defaultView, filterRegion, filterArea,
//                 filterWeek, filterDm, filterGm
// ═══════════════════════════════════════════════

function getUserPrefs(email) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, '__USERPREFS__');
  if (!email) return { ok: false, error: 'Email required' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      try {
        var prefs = JSON.parse(data[i][1] || '{}');
        return { ok: true, prefs: prefs };
      } catch(e) {
        return { ok: true, prefs: {} };
      }
    }
  }
  return { ok: true, prefs: null }; // null = no saved prefs
}

function saveUserPrefs(email, prefs) {
  if (!email) return { ok: false, error: 'Email required' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreate(ss, '__USERPREFS__');
  var data  = sheet.getDataRange().getValues();
  var now   = new Date().toISOString();
  var blob  = JSON.stringify(prefs);

  // Ensure header row
  if (data.length === 0 || data[0][0] !== 'email') {
    sheet.getRange(1, 1, 1, 3).setValues([['email', 'prefs', 'lastUpdated']]);
    data = sheet.getDataRange().getValues();
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[email, blob, now]]);
      return { ok: true, saved: true };
    }
  }
  // New row
  sheet.appendRow([email, blob, now]);
  return { ok: true, saved: true };
}
