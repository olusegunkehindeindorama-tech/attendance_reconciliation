function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Attendance Reconciliation Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SHEET_NAME = "Sheet1";
const EMP_SHEET_NAME = "Emp_Details";
const ABSENT_CASES_SHEET = "Absent_Cases";

function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Normalize various date formats coming from Sheets into a stable key and display string.
 * Supports Date objects, "M/D/YYYY", "D/M/YYYY", "YYYY-MM-DD", "d-M-Y", etc.
 */
function normalizeDate(value) {
  if (value === null || value === undefined || value === "") return null;

  let d = null;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    d = value;
  } else {
    const s = value.toString().trim();
    d = new Date(s);
    if (isNaN(d.getTime())) {
      const m1 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
      if (m1) {
        const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
        const mon = months[m1[2].charAt(0).toUpperCase() + m1[2].slice(1).toLowerCase()];
        if (mon !== undefined) d = new Date(parseInt(m1[3],10), mon, parseInt(m1[1],10));
      }
    }
    if (isNaN(d.getTime())) {
      const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m2) {
        d = new Date(parseInt(m2[3],10), parseInt(m2[1],10) - 1, parseInt(m2[2],10));
      }
    }
  }

  if (!d || isNaN(d.getTime())) return null;

  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const display = d.getDate() + "-" + monthNames[d.getMonth()] + "-" + d.getFullYear();
  const key = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
  return { display: display, key: key, dateObj: d };
}

/**
 * Payroll cycle is 16th → 15th.
 * Day >= 16  → salary / arrears month = 1st of NEXT month
 * Day <= 15  → salary / arrears month = 1st of CURRENT month
 * Examples:
 *   17-May-2026 → 1-Jun-2026
 *   12-Jun-2026 → 1-Jun-2026
 *   19-Jul-2026 → 1-Aug-2026
 */
function getArrearsMonthForDate(displayDate) {
  const norm = normalizeDate(displayDate);
  if (!norm) return "";

  const d = norm.dateObj;
  const day = d.getDate();
  let year = d.getFullYear();
  let month = d.getMonth(); // 0-based

  if (day >= 16) {
    month += 1;
    if (month > 11) {
      month = 0;
      year += 1;
    }
  }

  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return "1-" + monthNames[month] + "-" + year;
}

/**
 * Load ALL absent cases once, grouped by Emp ID for fast client-side filtering.
 */
function loadAllAbsentCases() {
  const sheet = getSheet(ABSENT_CASES_SHEET);
  const byEmp = {};

  if (!sheet) return byEmp;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowEmp = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    if (!rowEmp) continue;

    const norm = normalizeDate(data[i][1]);
    if (!norm) continue;

    const status = data[i][2] ? data[i][2].toString().trim() : "";
    const selectable = status === "";

    if (!byEmp[rowEmp]) byEmp[rowEmp] = [];
    byEmp[rowEmp].push({
      display: norm.display,
      key: norm.key,
      status: status,
      selectable: selectable,
      sheetRow: i + 1
    });
  }

  // Sort each employee's cases by date
  for (const emp in byEmp) {
    byEmp[emp].sort(function(a, b) {
      return a.key < b.key ? -1 : (a.key > b.key ? 1 : 0);
    });
  }

  return byEmp;
}

function getInitialData() {
  const mainSheet = getSheet(SHEET_NAME);
  const data = mainSheet.getDataRange().getDisplayValues();

  // 1. Calculate last (highest) and next Rec No
  let maxRec = 403;
  for (let i = 1; i < data.length; i++) {
    const recStr = data[i][9] ? data[i][9].toString().toUpperCase() : "";
    if (recStr.startsWith("R")) {
      let num = parseInt(recStr.replace("R", ""), 10);
      if (!isNaN(num) && num > maxRec) {
        maxRec = num;
      }
    }
  }
  const lastRecNo = "R" + maxRec;          // highest existing (for Manual default)
  const nextRecNo = "R" + (maxRec + 1);    // next available (for Auto)

  // 2. Employee dictionary
  const empSheet = getSheet(EMP_SHEET_NAME);
  let empDict = {};
  if (empSheet) {
    const empData = empSheet.getDataRange().getValues();
    for (let i = 1; i < empData.length; i++) {
      let id = empData[i][2] ? empData[i][2].toString().toUpperCase().trim() : "";
      let name = empData[i][6] ? empData[i][6].toString().trim() : "";
      if (id && name) {
        empDict[id] = name;
      }
    }
  }

  // 3. History (last 50)
  let history = [];
  if (data.length > 1) {
    for (let i = data.length - 1; i >= Math.max(1, data.length - 50); i--) {
      history.push({ row: i + 1, record: data[i] });
    }
  }

  // 4. All Absent_Cases loaded once into memory (grouped by Emp ID)
  const absentCasesByEmp = loadAllAbsentCases();

  return {
    lastRecNo: lastRecNo,
    nextRecNo: nextRecNo,
    history: history,
    empDict: empDict,
    absentCasesByEmp: absentCasesByEmp
  };
}

function submitForm(formData) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();

  const targetEmpId = formData.empId.toUpperCase().trim();
  const datesArray = JSON.parse(formData.absentDates);

  // Build duplicate lookup: EMPID|DATE → Rec No
  const recordMap = {};
  for (let i = 1; i < data.length; i++) {
    const existingEmpId = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    const existingDate = data[i][3] ? data[i][3].toString().trim() : "";
    const existingRecNo = data[i][9] ? data[i][9].toString().trim() : "No Rec No";
    if (existingEmpId && existingDate) {
      recordMap[existingEmpId + "|" + existingDate] = existingRecNo;
    }
  }

  const duplicatesFound = [];
  datesArray.forEach(function(dateItem) {
    const checkKey = targetEmpId + "|" + dateItem.trim();
    if (recordMap[checkKey]) {
      duplicatesFound.push({ date: dateItem, recNo: recordMap[checkKey] });
    }
  });

  if (duplicatesFound.length > 0) {
    let duplicateSummary = "Submission Denied! The entry has already been processed:\n";
    duplicatesFound.forEach(function(dup) {
      duplicateSummary += `• Date [${dup.date}] is logged under Rec Code: ${dup.recNo}\n`;
    });
    return { success: false, message: duplicateSummary };
  }

  // Assign Rec No
  let assignedRecNo = "";
  if (formData.assignRec === "Yes") {
    // Recompute next to avoid race
    let maxRec = 403;
    for (let i = 1; i < data.length; i++) {
      const recStr = data[i][9] ? data[i][9].toString().toUpperCase() : "";
      if (recStr.startsWith("R")) {
        let num = parseInt(recStr.replace("R", ""), 10);
        if (!isNaN(num) && num > maxRec) maxRec = num;
      }
    }
    assignedRecNo = "R" + (maxRec + 1);
  } else {
    assignedRecNo = formData.manualRecNo || "";
  }

  // Build rows — arrears month calculated per date from payroll cycle
  const rowsToInsert = [];
  datesArray.forEach(function(dateItem) {
    const arrearsMonth = getArrearsMonthForDate(dateItem);
    rowsToInsert.push([
      targetEmpId,
      formData.entryMonth,
      arrearsMonth,          // auto-calculated per date
      dateItem,
      1,
      formData.reason,
      formData.proof,
      formData.action,
      formData.actionMonth,
      assignedRecNo
    ]);
  });

  if (rowsToInsert.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
  }

  // Mark Absent_Cases status (best-effort)
  try {
    updateAbsentCasesStatus(targetEmpId, datesArray, formData.reason || "Treated");
  } catch (e) {
    // non-fatal
  }

  return {
    success: true,
    recNo: assignedRecNo,
    message: `Successfully added ${datesArray.length} record(s). Assigned Rec No: ${assignedRecNo}`
  };
}

function updateAbsentCasesStatus(empId, displayDates, status) {
  const sheet = getSheet(ABSENT_CASES_SHEET);
  if (!sheet) return;

  const target = empId.toString().toUpperCase().trim();
  const wantedKeys = {};
  displayDates.forEach(function(d) {
    const norm = normalizeDate(d);
    if (norm) wantedKeys[norm.key] = true;
  });

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowEmp = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    if (rowEmp !== target) continue;
    const norm = normalizeDate(data[i][1]);
    if (!norm) continue;
    if (wantedKeys[norm.key]) {
      const current = data[i][2] ? data[i][2].toString().trim() : "";
      if (current === "") {
        sheet.getRange(i + 1, 3).setValue(status);
      }
    }
  }
}

function deleteRecord(rowNumber) {
  const sheet = getSheet(SHEET_NAME);
  sheet.deleteRow(rowNumber);
  return getInitialData();
}
