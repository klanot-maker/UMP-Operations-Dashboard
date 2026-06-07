// ============================================================
// UMP OPERATIONS DASHBOARD — Code.gs  (v3)
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('UMP Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── SPREADSHEET IDs ──────────────────────────────────────────
const SS_FINANCIAL  = '1JvHBdSaeh6c2KIktDAOTub93jTOqs5znszzwKEdbe1o';
const SS_COMPLAINTS = '1j1wZ3z3YMvWhDpQbzYTNUAWlaj2rRt1RwP1ENSX3Lr0';
const SS_STAFF      = '1fA1ytS34bcUf-xua55d_RWCgmBwrZEYlvhQtSNwGPwE';

// ── MASTER LOADER (called once from client) ───────────────────
function getAllData() {
  return {
    financial:  getFinancialData(),
    complaints: getComplaintsData(),
    staff:      getStaffData()
  };
}

// ════════════════════════════════════════════════════════════
// FINANCIAL COST
// ════════════════════════════════════════════════════════════
function getFinancialData() {
  try {
    var ss    = SpreadsheetApp.openById(SS_FINANCIAL);
    var sheet = ss.getSheetByName('Financial cost') || ss.getSheets()[0];
    var data  = sheet.getDataRange().getValues();
    return { rows: data.map(function(row){ return row.map(function(c){ return String(c); }); }) };
  } catch(e) {
    return { error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
// UMP COMPLAINTS
// Columns: A=Date B=FO C=Quality D=Health E=Spilled F=Cold G=Dispatch H=Logistics I=Total
// Targets: W2=FO  X2=Quality  Z2=Spilled  AB2=Dispatch  AC2=Logistics  (0-indexed: 22,23,25,27,28)
// Pct row: N2:T2 = cols 13–19
// ════════════════════════════════════════════════════════════
function getComplaintsData() {
  try {
    var ss     = SpreadsheetApp.openById(SS_COMPLAINTS);
    var sheet  = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === 1266592099) { sheet = sheets[i]; break; }
    }
    if (!sheet) sheet = ss.getSheetByName('MOM Complaints') || sheets[0];

    var all  = sheet.getDataRange().getValues();
    var tRow = all[1] || [];

    var targets = {
      fo: safeNum(tRow[22]),   // W2
      ql: safeNum(tRow[23]),   // X2
      sl: safeNum(tRow[25]),   // Z2
      dp: safeNum(tRow[27]),   // AB2
      le: safeNum(tRow[28])    // AC2
    };

    var months = [];
    for (var r = 1; r <= 18; r++) {
      var row   = all[r] || [];
      var label = String(row[0] || '').trim();
      if (!label) continue;
      var total = safeNum(row[8]);
      if (total === 0 && r > 17) continue; // skip future empty months
      months.push({
        label: label,
        fo:    safeNum(row[1]),
        ql:    safeNum(row[2]),
        hr:    safeNum(row[3]),
        sl:    safeNum(row[4]),
        cs:    safeNum(row[5]),
        dp:    safeNum(row[6]),
        le:    safeNum(row[7]),
        total: total
      });
    }

    return { months: months, targets: targets };
  } catch(e) {
    return { error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
// PRODUCTION STAFF
// Sheet structure (each section):
//   Row N+0 : Department header (col A = "Kitchen Staff" etc, col B = blank)
//   Row N+1 : Column headers  (Production Date | Permanent Staff | ...)
//   Row N+2+: Daily data rows until next blank separator or next dept header
//
// Dept keywords detected: Kitchen, Steward, Dispatch, Office
// Office Support is typically after row 100 — server-side reads the full sheet
// ════════════════════════════════════════════════════════════
function getStaffData() {
  try {
    var ss     = SpreadsheetApp.openById(SS_STAFF);
    // Try the active June 2026 sheet first, then fallback to first sheet
    var sheet  = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === 773931869) { sheet = sheets[i]; break; }
    }
    if (!sheet) {
      for (var j = 0; j < sheets.length; j++) {
        var n = sheets[j].getName().toLowerCase();
        if (n.indexOf('june') > -1 || n.indexOf('jun') > -1) { sheet = sheets[j]; break; }
      }
    }
    if (!sheet) sheet = sheets[0];

    var all       = sheet.getDataRange().getValues();
    var totalRows = all.length;
    var KEYWORDS  = ['kitchen','steward','dispatch','office'];
    var sections  = [];
    var i         = 0;

    while (i < totalRows) {
      var cellA = String(all[i][0] || '').trim();
      var cellB = String(all[i][1] || '').trim();

      // Detect department header: col A contains a keyword, col B is empty (merged cell)
      var isDeptHdr = cellB === '' && (function(a) {
        for (var k = 0; k < KEYWORDS.length; k++) {
          if (a.toLowerCase().indexOf(KEYWORDS[k]) > -1) return true;
        }
        return false;
      })(cellA);

      if (isDeptHdr) {
        var deptName = cellA; // e.g. "Kitchen Staff", "Office Support"
        i += 2; // skip header row + column-label row

        var rows = [];
        while (i < totalRows) {
          var nA = String(all[i][0] || '').trim();
          var nB = String(all[i][1] || '').trim();

          // Next dept header → break
          var isNextHdr = nB === '' && (function(a2) {
            for (var k2 = 0; k2 < KEYWORDS.length; k2++) {
              if (a2.toLowerCase().indexOf(KEYWORDS[k2]) > -1) return true;
            }
            return false;
          })(nA);
          if (isNextHdr) break;

          // Skip fully blank rows
          if (!nA && !nB) { i++; continue; }

          var ps = safeNum(all[i][1]);
          if (ps > 0) {
            var dpsRaw = String(all[i][10] || '');
            var dps    = (dpsRaw.indexOf('#') > -1 || dpsRaw === '') ? 0 : safeNum(all[i][10]);
            var delRaw = safeNum(all[i][9]);
            rows.push({
              pd:  nA,
              ps:  ps,
              on:  safeNum(all[i][2]),
              off: safeNum(all[i][3]),
              al:  safeNum(all[i][4]),
              sl:  safeNum(all[i][5]),
              sup: safeNum(all[i][6]),
              tot: safeNum(all[i][7]),
              dd:  String(all[i][8] || ''),
              del: delRaw,
              dps: dps
            });
          }
          i++;
        }

        if (rows.length > 0) {
          var lastRow  = rows[rows.length - 1];
          var supVals  = rows.filter(function(r){ return r.sup > 0; }).map(function(r){ return r.sup; });
          var dpsVals  = rows.filter(function(r){ return r.dps > 0; }).map(function(r){ return r.dps; });
          var totalDel = rows.reduce(function(s,r){ return s + r.del; }, 0);
          var avgSup   = supVals.length ? Math.round(supVals.reduce(function(a,b){return a+b;},0)/supVals.length) : 0;
          var avgDps   = dpsVals.length ? Math.round(dpsVals.reduce(function(a,b){return a+b;},0)/dpsVals.length) : 0;

          sections.push({
            name:       deptName,
            permStaff:  lastRow.ps,
            avgSup:     avgSup,
            totalDel:   totalDel,
            avgDps:     avgDps,
            daysLogged: rows.length,
            rows:       rows
          });
        }
      } else {
        i++;
      }
    }

    return { sections: sections };
  } catch(e) {
    return { error: e.message };
  }
}

// ── HELPER ────────────────────────────────────────────────────
function safeNum(v) {
  var s = String(v || '').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}