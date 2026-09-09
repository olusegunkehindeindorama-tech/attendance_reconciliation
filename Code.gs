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
    // Try native parse first
    d = new Date(s);
    if (isNaN(d.getTime())) {
      // Try d-M-Y (e.g. 20-Jun-2026)
      const m1 = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
      if (m1) {
        const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
        const mon = months[m1[2].charAt(0).toUpperCase() + m1[2].slice(1).toLowerCase()];
        if (mon !== undefined) d = new Date(parseInt(m1[3],10), mon, parseInt(m1[1],10));
      }
    }
    if (isNaN(d.getTime())) {
      // Try M/D/YYYY or D/M/YYYY — prefer US style from the sample sheet
      const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m2) {
        d = new Date(parseInt(m2[3],10), parseInt(m2[1],10) - 1, parseInt(m2[2],10));
      }
    }
  }

  if (!d || isNaN(d.getTime())) return null;

  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const display = d.getDate() + "-" + monthNames[d.getMonth()] + "-" + d.getFullYear();
  // ISO-like key for reliable matching
  const key = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
  return { display: display, key: key, dateObj: d };
}

function getInitialData() {
  const mainSheet = getSheet(SHEET_NAME);
  const data = mainSheet.getDataRange().getDisplayValues();
  
  // 1. Calculate next Rec No
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
  const nextRecNo = "R" + (maxRec + 1);

  // 2. Fetch Employee Dictionary for Name Lookup
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

  // 3. Return history (last 50 records)
  let history = [];
  if (data.length > 1) {
    for (let i = data.length - 1; i >= Math.max(1, data.length - 50); i--) {
      history.push({ row: i + 1, record: data[i] });
    }
  }

  return {
    nextRecNo: nextRecNo,
    history: history,
    empDict: empDict
  };
}

/**
 * Return absent cases for a given employee from the Absent_Cases sheet.
 * Each item: { display, key, status, selectable, sheetRow }
 * selectable = true only when Reconciliation_Status is empty.
 */
function getAbsentCasesForEmployee(empId) {
  const target = (empId || "").toString().toUpperCase().trim();
  if (!target) return { cases: [], message: "No employee ID provided." };

  const sheet = getSheet(ABSENT_CASES_SHEET);
  if (!sheet) {
    return { cases: [], message: "Absent_Cases sheet not found." };
  }

  const data = sheet.getDataRange().getValues(); // raw values so dates stay as Date objects when possible
  const cases = [];

  for (let i = 1; i < data.length; i++) {
    const rowEmp = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    if (rowEmp !== target) continue;

    const norm = normalizeDate(data[i][1]);
    if (!norm) continue;

    const status = data[i][2] ? data[i][2].toString().trim() : "";
    const selectable = status === "";

    cases.push({
      display: norm.display,
      key: norm.key,
      status: status,
      selectable: selectable,
      sheetRow: i + 1
    });
  }

  // Sort ascending by date key
  cases.sort(function(a, b) {
    return a.key < b.key ? -1 : (a.key > b.key ? 1 : 0);
  });

  return {
    cases: cases,
    message: cases.length === 0 ? "No absent cases found for this employee." : ""
  };
}

function submitForm(formData) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues(); 
  
  const targetEmpId = formData.empId.toUpperCase().trim();
  const datesArray = JSON.parse(formData.absentDates);
  
  // 1. Build a lookup map of all pre-existing entries: Key = "EMPID|DATE" -> Value = Rec No
  const recordMap = {};
  for (let i = 1; i < data.length; i++) {
    const existingEmpId = data[i][0] ? data[i][0].toString().toUpperCase().trim() : "";
    const existingDate = data[i][3] ? data[i][3].toString().trim() : "";
    const existingRecNo = data[i][9] ? data[i][9].toString().trim() : "No Rec No";
    
    if (existingEmpId && existingDate) {
      recordMap[existingEmpId + "|" + existingDate] = existingRecNo;
    }
  }
  
  // 2. Loop through each of the incoming absent days to look for an identity match
  const duplicatesFound = [];
  datesArray.forEach(function(dateItem) {
    const checkKey = targetEmpId + "|" + dateItem.trim();
    if (recordMap[checkKey]) {
      duplicatesFound.push({
        date: dateItem,
        recNo: recordMap[checkKey]
      });
    }
  });
  
  // 3. If any duplicate elements match, halt operation and report the match back to user
  if (duplicatesFound.length > 0) {
    let duplicateSummary = "Submission Denied! The entry has already been processed:\n";
    duplicatesFound.forEach(function(dup) {
      duplicateSummary += `• Date [${dup.date}] is logged under Rec Code: ${dup.recNo}\n`;
    });
    
    return {
      success: false,
      message: duplicateSummary
    };
  }
  
  // 4. Proceed with execution if clearance check passes cleanly
  let assignedRecNo = "";
  if (formData.assignRec === "Yes") {
    const initialData = getInitialData();
    assignedRecNo = initialData.nextRecNo;
  } else {
    assignedRecNo = formData.manualRecNo || "";
  }

  const rowsToInsert = [];
  datesArray.forEach(function(dateItem) {
    rowsToInsert.push([
      targetEmpId,                    
      formData.entryMonth,                
      formData.arrearsMonth,              
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

  // 5. Optionally mark the selected dates as reconciled in Absent_Cases
  //    using the reason as the status value (best-effort; does not fail the submit)
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

/**
 * Mark matching rows in Absent_Cases with the given status for the selected dates.
 */
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
      // Only write if currently empty (do not overwrite existing status)
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
