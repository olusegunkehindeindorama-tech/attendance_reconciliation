function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Attendance Reconciliation Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const SHEET_NAME = "Sheet1";
const EMP_SHEET_NAME = "Emp_Details";

function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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

  return {
    success: true,
    recNo: assignedRecNo,
    message: `Successfully added ${datesArray.length} record(s). Assigned Rec No: ${assignedRecNo}`
  };
}

function deleteRecord(rowNumber) {
  const sheet = getSheet(SHEET_NAME);
  sheet.deleteRow(rowNumber);
  return getInitialData(); 
}
