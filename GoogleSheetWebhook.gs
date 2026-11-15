/**
 * Classroom Helper -> Google Sheets webhook
 *
 * SETUP (once per class / per Google Sheet):
 *  1. Create a new Google Sheet for the class.
 *  2. Extensions > Apps Script. Delete any code, paste this whole file, Save.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Authorize when asked, then copy the Web app URL (ends in /exec).
 *  5. In Classroom Helper, edit the class and paste the URL into
 *     "Google Sheet Webhook URL".
 *
 * The same script works for every class. Just repeat the steps with a new sheet.
 * If you change this code later: Deploy > Manage deployments > Edit > New version.
 *
 * LAYOUT (one tab per month for each type):
 *   "Attendance 2026-11"  1 = present, X = absent
 *   "Homework 2026-11"    1 = submitted, X = not submitted
 *   Row 1: class name, then 1, 2, 3 ... (class day number)
 *   Names down column A, one column per day
 *   Bottom: Date row (vertical text), and on Homework a Note row (assignment title)
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    if (data.action !== 'sync_days') return out_({ ok: false, error: 'Unknown action' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var students = data.students || [];
    var className = data.className || 'Class';

    (data.days || []).forEach(function (day) {
      if (!day.date) return;
      if (day.attendance) {
        writeDay_(ss, 'Attendance', className, day.date, students, day.attendance, 'present', 'absent', null);
      }
      if (day.homework) {
        writeDay_(ss, 'Homework', className, day.date, students, day.homework, 'done', 'missing', day.homeworkTopic || '');
      }
    });

    SpreadsheetApp.flush();
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  return out_({ ok: true, message: 'Classroom Helper webhook is running' });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getTab_(ss, kind, tabName) {
  var sh = ss.getSheetByName(tabName);
  if (sh) return sh;

  var sheets = ss.getSheets();
  var index;
  if (kind === 'Attendance') {
    // Attendance months stay in front of all Homework months
    index = 0;
    sheets.forEach(function (s) { if (s.getName().indexOf('Attendance ') === 0) index++; });
  } else {
    index = sheets.length;
  }
  sh = ss.insertSheet(tabName, index);

  // Remove the untouched default "Sheet1"
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  return sh;
}

function findDateRow_(sh) {
  var last = sh.getLastRow();
  if (last < 1) return -1;
  var colA = sh.getRange(1, 1, last, 1).getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]) === 'Date') return i + 1;
  }
  return -1;
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function writeDay_(ss, kind, className, dateStr, students, statuses, okWord, badWord, note) {
  var isHw = (kind === 'Homework');
  var tabName = kind + ' ' + String(dateStr).substring(0, 7);
  var sh = getTab_(ss, kind, tabName);

  // ---- First-time layout ----
  var dateRow = findDateRow_(sh);
  if (dateRow < 0) {
    var n = students.length;
    dateRow = n + 2;
    ensureSize_(sh, dateRow + (isHw ? 1 : 0), 2);
    if (n > 0) {
      sh.getRange(2, 1, n, 1).setValues(students.map(function (s) { return [s]; }));
    }
    sh.getRange(dateRow, 1).setValue('Date');
    if (isHw) sh.getRange(dateRow + 1, 1).setValue('Note');
  }

  // ---- Add students that are not on the sheet yet (new pupils) ----
  var existing = dateRow > 2
    ? sh.getRange(2, 1, dateRow - 2, 1).getValues().map(function (r) { return String(r[0]); })
    : [];
  students.forEach(function (name) {
    if (existing.indexOf(String(name)) === -1) {
      sh.insertRowBefore(dateRow);
      sh.getRange(dateRow, 1, 1, sh.getMaxColumns()).clearFormat();
      sh.getRange(dateRow, 1).setValue(name);
      existing.push(String(name));
      dateRow++;
    }
  });

  // ---- Find or create this day's column (kept in date order) ----
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var dateVals = lastCol > 1 ? sh.getRange(dateRow, 2, 1, lastCol - 1).getValues()[0].map(String) : [];
  while (dateVals.length && dateVals[dateVals.length - 1] === '') dateVals.pop();

  var idx = dateVals.indexOf(dateStr);
  var col;
  if (idx >= 0) {
    col = idx + 2;
  } else {
    var pos = dateVals.length;
    for (var i = 0; i < dateVals.length; i++) {
      if (dateVals[i] !== '' && dateVals[i] > dateStr) { pos = i; break; }
    }
    col = pos + 2;
    if (pos < dateVals.length) {
      sh.insertColumnBefore(col);
    } else {
      ensureSize_(sh, dateRow + (isHw ? 1 : 0), col);
    }
    sh.getRange(dateRow, col).setNumberFormat('@').setValue(dateStr);
  }

  // ---- Date / note cells (vertical text keeps columns thin) ----
  sh.getRange(dateRow, col).setNumberFormat('@').setValue(dateStr)
    .setTextRotation(90).setHorizontalAlignment('center').setVerticalAlignment('bottom').setFontWeight('bold');
  if (isHw) {
    sh.getRange(dateRow + 1, col).setNumberFormat('@').setValue(note || '')
      .setTextRotation(90).setHorizontalAlignment('center').setVerticalAlignment('top').setWrap(false);
  }

  // ---- Marks ----
  var nNames = dateRow - 2;
  if (nNames > 0) {
    var names = sh.getRange(2, 1, nNames, 1).getValues().map(function (r) { return String(r[0]); });
    var rng = sh.getRange(2, col, nNames, 1);
    var cur = rng.getValues();
    var vals = [], colors = [], weights = [];
    for (var r = 0; r < nNames; r++) {
      var st = statuses[names[r]];
      var v = cur[r][0];
      if (st === okWord) v = 1;
      else if (st === badWord) v = 'X';
      vals.push([v]);
      colors.push([v === 'X' ? '#d93025' : '#188038']);
      weights.push([v === 'X' ? 'bold' : 'normal']);
    }
    rng.setValues(vals).setFontColors(colors).setFontWeights(weights).setHorizontalAlignment('center');
  }

  // ---- Header row: class name + day numbers ----
  var lastCol2 = Math.max(sh.getLastColumn(), 1);
  var dv2 = lastCol2 > 1 ? sh.getRange(dateRow, 2, 1, lastCol2 - 1).getValues()[0].map(String) : [];
  while (dv2.length && dv2[dv2.length - 1] === '') dv2.pop();
  var nums = [];
  for (var k = 1; k <= dv2.length; k++) nums.push(k);
  sh.getRange(1, 1).setValue(className).setFontWeight('bold');
  if (nums.length) {
    sh.getRange(1, 2, 1, nums.length).setValues([nums]).setFontWeight('bold').setHorizontalAlignment('center');
    sh.setColumnWidths(2, nums.length, 30);
  }

  // ---- Sizing / freeze ----
  sh.setColumnWidth(1, 170);
  sh.setRowHeight(dateRow, 95);
  if (isHw) sh.setRowHeight(dateRow + 1, 170);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
}
