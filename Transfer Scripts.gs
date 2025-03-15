const CELL_EDIT_LIMIT = 4;   // set number of cells that can be edited at once

// USED TO IMPORT NEW REGISTRATION FROM FILLOUT FORM
function onChange(e) {
  // Get details of edit event's sheet
  console.log({
    authMode:  e.authMode.toString(),
    changeType: e.changeType,
    user: e.user,
  });

  const thisSource = e.source;

  // Try-catch to prevent errors when sheetId cannot be found
  try {
    const thisSheetID = thisSource.getSheetId();
    const thisLastRow = thisSource.getLastRow();

    const thisChange = e.changeType;
    console.log(`Change Type: ${thisChange}`);

    if (thisSheetID === IMPORT_SHEET_ID && thisChange === 'INSERT_ROW') {
      console.log('Executing if-statement now...');
      const importSheet = thisSource.getSheetById(thisSheetID);
      const registrationObj = importSheet.getRange(thisLastRow, 1).getValue();

      const lastRow = copyToMain(registrationObj);
      onFormSubmit(lastRow);
      
      // Successful execution...
      console.log('Exiting if-statement successfully!');
    }
  }
  catch (error) {
    console.error(error);
  }
  finally {
    console.log(e);
  }

}

function transferLastImport() {
  const thisLastRow = IMPORT_SHEET.getLastRow();
  transferThisRow(thisLastRow);
}

function transferThisRow(row) {
  const registrationObj = IMPORT_SHEET.getRange(row, 1).getValue();
  const lastRow = copyToMain(registrationObj);
  onFormSubmit(lastRow);
}


function onEdit(e) {
  // Get details of edit event's sheet
  const thisRange = e.range;
  const thisSheet = thisRange.getSheet();
  const thisSheetName = thisSheet.getName();

  var debug_e = {
    //authMode:  e.authMode,
    range: e.range.getA1Notation(),
    sheetName: e.range.getSheet().getSheetName(),
    //source:  e.source,
    newValue: e.value,
    oldValue: e.oldValue
  }
  console.log(debug_e);

  if (thisRange.getNumRows() > 2) return;  // prevent sheet-wide changes
  else if (thisRange.getNumColumns() > CELL_EDIT_LIMIT) {
    // TODO: add function to individually process changes
    Logger.log(`More than ${CELL_EDIT_LIMIT} columns edited at once`);
  }

  console.log(`onEdit 1 -> thisSheetName: ${thisSheetName}`);

  // Check if legal sheet
  if (thisSheetName != SHEET_NAME && thisSheetName != MASTER_NAME) return;

  console.log("onEdit 1a -> Passed first check");

  //if(e.value == e.oldValue) return;   // Values have not changed. Edit was on sheet formatting.

  console.log("onEdit 1b -> Passed second check");

  // Check if legal edit
  if (!verifyLegalEditInRange(e, thisSheet)) return;

  console.log("onEdit 2 -> Passed \`verifyLegalEditInRange()\`");

  // Get the email column for the current sheet
  const thisEmailCol = GET_COL_MAP_(thisSheetName).emailCol;
  const thisRow = e.range.getRow();

  console.log(`onEdit 3 -> thisEmailCol: ${thisEmailCol} thisRow: ${thisRow}`);

  // Get email from `thisRow` and `thisEmailCol`
  const email = thisSheet.getRange(thisRow, thisEmailCol).getValue();

  const isMainSheet = (thisSheetName === SHEET_NAME);
  console.log(`onEdit 4 -> email: ${email} isMainSheet: ${isMainSheet}`);

  const sourceSheet = isMainSheet ? MAIN_SHEET : MASTER_SHEET;
  const targetSheet = isMainSheet ? MASTER_SHEET : MAIN_SHEET;
  const targetRow = findMemberByEmail(email, targetSheet);  // Find row of member in `targetSheet` using their email

  // Throw error message if member not in `targetSheet`
  if (targetRow == null) {
    const errorMessage = `
      --- onEdit() ---
      targetRow not found in ${targetSheet}. 
      Edit made in ${thisSheetName} at row ${thisRow}.
      Email of edited member: ${email}. Please review this error.
    `
    throw Error(errorMessage);
  }

  console.log(`onEdit 5 -> targetRow: ${targetRow} found by \`findMemberByEmail()\``);

  updateFeeInfo(e, thisSheetName, targetRow, targetSheet);
  console.log(`onEdit 6 -> successfully completed trigger check`);
}


/**
 * @param {Event} e  Event Object from `onEdit`.
 * @param {SpreadsheetApp.Sheet} sheet  Sheet where edit occurred.
 */

function verifyLegalEditInRange(e, sheet) {
  Logger.log("NOW ENTERING verifyLegalEditInRange()...");
  const sheetName = sheet.getName();
  var thisRow = e.range.getRow();
  var thisCol = e.range.getColumn();
  Logger.log(`verifyLegalEditInRange 1 -> sheetName: ${sheetName}`);

  // Function to get column mappings
  const feeStatus = GET_COL_MAP_(sheetName).feeStatus;
  const isInternalCollected = GET_COL_MAP_(sheetName).isInternalCollected;

  Logger.log(`verifyLegalEditInRange 2 -> feeStatusCol: ${feeStatus}, isInternalCollected: ${isInternalCollected}`);

  const feeEditRange = {
    top: 2,    // Skip header row
    bottom: sheet.getLastRow(),
    leftmost: feeStatus,
    rightmost: isInternalCollected,
  }

  // Helper function to log error message and exit function
  const logAndExitFalse = (cell) => { Logger.log(`${cell} is out of bounds`); return false; }

  // Exit if we're out of range
  if (thisRow < feeEditRange.top || thisRow > feeEditRange.bottom) logAndExitFalse("Row");
  if (thisCol < feeEditRange.leftmost || thisCol > feeEditRange.rightmost) logAndExitFalse("Column");

  return true;    // edit e is within legal edit range
}


/** 
 * Update fee status from `sourceSheet` to `targetSheet`.
 * 
 * @param {Event} e  Event Object from `onEdit`.
 * @param {string} sourceSheetName  Name of source sheet to extract fee info.
 * @param {number} targetRow  Target row to update.
 * @param {SpreadsheetApp.Sheet} targetSheet  Target sheet to update fee info.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 16, 2024
 * @update  Dec 18, 2024
 * 
 */

function updateFeeInfo(e, sourceSheetName, targetRow, targetSheet) {
  const thisRange = e.range;
  const thisCol = thisRange.getColumn();
  const targetSheetName = targetSheet.getSheetName();

  console.log(`NOW ENTERING updateFeeInfo()`);
  console.log(`Source: ${sourceSheetName}, thisCol: ${thisCol} && Target: ${targetSheetName}, targetRow: ${targetRow}`);

  const sourceCols = GET_COL_MAP_(sourceSheetName);   // Map of type of member data to its column index
  const targetCols = GET_COL_MAP_(targetSheetName);   // Get map of member data with respective column indices

  Logger.log("updateFeeInfo 1 -> Successfully got sourceCols and targetCols");

  // Find respective column where `targetCol` contains same data as `sourceCol`.
  const getTargetCol = (source) => {
    switch (source) {
      case (sourceCols.feeStatus): return targetCols.feeStatus;
      case (sourceCols.collectionDate): return targetCols.collectionDate;
      case (sourceCols.collector): return targetCols.collector;
      case (sourceCols.isInternalCollected): return targetCols.isInternalCollected;
    }
  };

  // Find which column was edited in `sourceSheet` and find respective col in `targetSheet`
  const targetCol = getTargetCol(thisCol);
  Logger.log(`updateFeeInfo 2 -> targetRow: ${targetRow} targetCol: ${targetCol}`);

  const targetRange = targetSheet.getRange(targetRow, targetCol);

  // Special case: MASTER stores payment history as semesterCode(s).
  // If isPaid, then add semesterCode to payment history, i.e. bool -> str
  // Otherwise, nothing to modify in MASTER for member's payment history
  if (targetSheetName == MASTER_NAME && targetCol == MASTER_PAYMENT_HIST) {
    console.log("updateFeeInfo 3 -> entering if statement");
    const value = thisRange.getValue() || "";
    const isPaid = parseBool(value);    // convert to bool
    console.log(`updateFeeInfo 3b -> Value: ${value} isPaid: ${isPaid}`);

    // Only modify payment history if isPaid == true.
    if (isPaid) {
      console.log("updateFeeInfo 3c -> entering isPaid");
      addPaidSemesterToHistory(targetRow, sourceSheetName);
    }
    else {
      console.log("updateFeeInfo 3c -> entering NOT(isPaid)");
    }

  }
  else if (sourceSheetName == MASTER_NAME && thisCol == MASTER_PAYMENT_HIST) {
    // CASE 2: Add history payment to sheet
    console.log("updateFeeInfo 3 ->  entering else if statement");
    const paymentHistory = thisRange.getValue() || "";
    updateIsFeePaid(paymentHistory, targetRow, targetCol, targetSheet);
  }
  else {
    console.log("updateFeeInfo 3 ->  entering else statement");
    thisRange.copyTo(targetRange, { contentsOnly: true });
  }

  console.log("updateFeeInfo 4 ->  finished updating payment history");
}


/** 
 * Transfer new member registration from `Import` to main sheet.
 * 
 * @trigger  New entry in `Import` sheet.
 * @param {Object} registration  Information on member registration.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 18, 2023
 * @update  Feb 23, 2025
 * 
 */

function copyToMain(registration, row = getLastSubmissionInMain()) {
  const mainSheet = MAIN_SHEET;
  const importMap = IMPORT_MAP;

  const registrationObj = JSON.parse(registration.replace(/[\n\r\t]/g, ' '));
  console.log(registrationObj);

  const startRow = row + 1;
  const colSize = mainSheet.getLastColumn();

  const valuesByIndex = Array(colSize);

  // Format timestamp correctly; otherwise GSheet will not understand datetime
  const timestamp = registrationObj['timestamp'];
  if (timestamp != '') {
    const formattedTimestamp = Utilities.formatDate(
      new Date(timestamp),
      TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );

    registrationObj['timestamp'] = formattedTimestamp;   // replace with formatted
  }

  const removeRegex = /^[,\s-]+|[,\s-]+$/g;

  for (let [key, value] of Object.entries(registrationObj)) {
    if (key in importMap) {
      value = (typeof (value) === 'object') ? Object.values(value).join(', ') : value;
      let indexInMain = importMap[key] - 1;   // Set 1-index to 0-index
      valuesByIndex[indexInMain] = value.replace(removeRegex, ''); // Remove all unwanted
    }
  }

  // Set values of registration
  const rangeToImport = mainSheet.getRange(startRow, 1, 1, colSize);
  rangeToImport.setValues([valuesByIndex]);

  return startRow;
}


function packageMemberInfoInRow_(row) {
  const sheet = MAIN_SHEET;
  const semesterName = SHEET_NAME;

  // Get member data to populate pass template
  const endCol = MEMBER_ID_COL; // From first name to id col
  const memberData = sheet.getSheetValues(row, 1, 1, endCol)[0];

  // Add entry to front to allow 1-indexed data access like for GSheet
  memberData.unshift('');

  // Stringify fee status
  const memberFeeStatus = parseBool(memberData[IS_FEE_PAID_COL]) ? 'Paid' : 'Unpaid';

  // Get membership expiration date via sheet name
  const semesterCode = getSemesterCode_(semesterName); // Get the semester code based on the sheet name
  const membershipExpiration = getExpirationDate(semesterCode);

  // Map member info to pass info
  const memberInfo = {
    email: memberData[EMAIL_COL],
    firstName: memberData[FIRST_NAME_COL],
    lastName: memberData[LAST_NAME_COL],
    memberId: memberData[MEMBER_ID_COL],
    memberStatus: 'Active',    // If email not found, then membership expired
    feeStatus: memberFeeStatus,
    expiry: membershipExpiration,
  }

  return memberInfo;
}


function testMigrate() {
  let ex = `{"timestamp":"2025-02-22T22:34:26.899Z",
  "email":"jiangforrest1@gmail.com",
  "firstName":"Forrest",
  "lastName":"Jiang",
  "preferredName":"",
  "year":"Non-McGillian",
  "program":"N/A",
  "memberDescription":"Fresh graduate from Hong Kong just landed in Montreal, enjoy running as part of my sporting mix.
  Running semi-regularly for two years, registered for 2025 half and full marathons in Montreal
  ",
  "paymentAmount":"10",
  "paymentMethod":"Interac e-Transfer", 
  "interacRef":"C1AqtpGKhzgD",
  "comments":"",
  "referral": 
  {
  "name":"",
  "sources":"Social Media",
  "platform":""
  },
  "discountFriendEmail": ""
  }`
    ;

  if (false) {
    Logger.log(ex);
    ex = ex.replace(/[\n\r\t]/g, ' ');
    console.log(ex);

    const testMe = JSON.parse(ex);
    console.log(testMe);
  }
  
  const newRowIndex = copyToMain(ex);
  Logger.log(newRowIndex);
}

