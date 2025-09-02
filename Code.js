function connectToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();         // connects to the file you're editing
  const sheet = ss.getSheetByName("DocumentDB");            // connects to the tab named "DocumentDB"
  const data = sheet.getDataRange().getValues();            // gets all data
  Logger.log(data);                                         // logs it for checking
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("drsindex")
    .setTitle("ISMD Document Registration Sheet")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Check for duplicate Reference Number
function checkDuplicate(refNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DocumentDB");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == refNo) {
      return true; // Duplicate found
    }
  }
  return false; // No duplicate
}

// Validate form data
function validateFormData(formData) {
  for (let key in formData) {
    if (!formData[key] || formData[key].toString().trim() === "") {
      throw new Error(
        "Field '" + key + "' cannot be empty. Please complete all fields."
      );
    }
  }
}

// Submit data to the Database sheet
function submitData(formData) {
  validateFormData(formData);

  if (checkDuplicate(formData.refNo)) {
    throw new Error(
      "Reference Number '" +
        formData.refNo +
        "' already exists. Please use a unique Reference Number."
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DocumentDB");

  const newRow = [
    formData.refNo,
    formData.officeOrigin,
    formData.officeDestination,
    formData.dateFiled,
    formData.subject,
    formData.signatory,
    formData.recordedBy,
    new Date(), // Timestamp
  ];

  sheet.appendRow(newRow);
  return "Record saved successfully with Reference No. " + formData.refNo + ".";
}

// Search record by Reference Number
function searchRecord(refNo) {
  refNo = refNo.trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DocumentDB");
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { found: false };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = 0; i < data.length; i++) {
    let rowRefNo = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowRefNo === refNo) {
      return {
        found: true,
        record: {
          refNo: data[i][0],
          officeOrigin: data[i][1],
          officeDestination: data[i][2],
          dateFiled:
            data[i][3] instanceof Date
              ? Utilities.formatDate(
                  data[i][3],
                  Session.getScriptTimeZone(),
                  "yyyy-MM-dd"
                )
              : data[i][3],
          subject: data[i][4],
          signatory: data[i][5],
          recordedBy: data[i][6],
        },
      };
    }
  }
  return { found: false };
}

// Delete record by Reference Number
function deleteRecord(refNo) {
  if (!refNo || refNo.trim() === "") {
    throw new Error("Reference Number is required for deletion.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DocumentDB");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == refNo) {
      sheet.deleteRow(i + 2); // Adjust for header
      return "Record with Reference No. " + refNo + " deleted successfully.";
    }
  }
  return "No record found for Reference No. " + refNo + ".";
}

// Modify record by Reference Number
function modifyRecord(formData) {
  validateFormData(formData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DocumentDB");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == formData.refNo) {
      sheet.getRange(i + 2, 1, 1, 8).setValues([
        [
          formData.refNo,
          formData.officeOrigin,
          formData.officeDestination,
          formData.dateFiled,
          formData.subject,
          formData.signatory,
          formData.recordedBy,
          new Date(), // Update timestamp
        ],
      ]);
      return (
        "Record with Reference No. " +
        formData.refNo +
        " modified successfully."
      );
    }
  }
  return "No record found for Reference No. " + formData.refNo + ".";
}
