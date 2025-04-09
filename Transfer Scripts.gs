const CELL_EDIT_LIMIT = 4;   // set number of cells that can be edited at once

// USED TO IMPORT NEW REGISTRATION FROM FILLOUT FORM
function onChange(e) {
  // Get details of edit event's sheet
  console.log({
    authMode: e.authMode.toString(),
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
      console.log('Executing if block from onChange(e)...');
      const importSheet = thisSource.getSheetById(thisSheetID);
      const registrationObj = importSheet.getRange(thisLastRow, 1).getValue();

      const lastRow = copyFilloutRegToMain_(registrationObj);
      onFormSubmit(lastRow);

      // Successful execution...
      console.log('Exiting `onFormSubmit` from onChange(e) successfully!');
    }

    else if (thisSheetID === MASTER_SHEET.getSheetId() && thisChange === 'INSERT_ROW') {
      console.log('Executing else-if block from onChange(e)...');
      console.log('New row added to Master sheet');

      // Add formula, then copy submission to main sheet
      if(!isNewMemberViaApp_(thisLastRow)) {
        console.log('Exiting because member *not* added by app');
      }
      formatFeeCollection(thisLastRow);
      setMemberIdInRow_(MASTER_SHEET, thisLastRow);
      copyToMainFromMaster(thisLastRow);
      console.log('Added registration to main sheet from master!');

      notifyNewAppSubmission(thisLastRow);
      console.log('Sent an email notification!');

      sortMasterByEmail(); // Sort 'MASTER' by email once formatting completed
      
      // Successful execution...
      console.log('Exiting `formatFeeCollection` from onChange(e) successfully!');
    }
  }
  catch (error) {
    console.log('Whoops! Error raised in onChange(e)');
    Logger.log(error);
  }
}

function transferLastImport() {
  const thisLastRow = IMPORT_SHEET.getLastRow();
  transferThisRow_(thisLastRow);
}

function transferThisRow_(row) {
  const registrationObj = IMPORT_SHEET.getRange(row, 1).getValue();
  const lastRow = copyFilloutRegToMain_(registrationObj);
  onFormSubmit(lastRow);
}


function isNewMemberViaApp_(row) {
  const sheet = MASTER_SHEET;

  // STEP 1: Check if 3-char code for registration semester exists
  const rangeRegSem = sheet.getRange(row, MASTER_LAST_REG_SEM);
  const isBlank = rangeRegSem.trimWhitespace().isBlank();
  if (!isBlank) return false;   // Reg sem exists... not added via app

  // STEP 2: Get regSem using `SHEET_NAME`, then add to target row
  const regSem = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangeRegSem.setValue(regSem);
  
  // STEP 3: Finally return true once added
  return true;
}


function setMemberIdInRow_(sheet, row) {
  // STEP 1: Get email col in target sheet
  const sheetName = sheet.getSheetName();
  const colMap = GET_COL_MAP_(sheetName);
  const thisEmailCol = colMap.emailCol;

  // STEP 2: Encode email using `encodeFromInput`
  const email = sheet.getRange(row, thisEmailCol).getValue();
  const memberId = encodeFromInput(email);

  // STEP 3: Set member id in target row
  const thisIdCol = colMap.memberIdCol;
  sheet.getRange(row, thisIdCol).setValue(memberId);
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
  if (!verifyLegalEditInRange_(e, thisSheet)) return;

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
  if (targetRow === null) {
    const errorMessage = `
      --- onEdit() ---
      targetRow not found in ${targetSheet}. 
      Edit made in ${thisSheetName} at row ${thisRow}.
      Email of edited member: ${email}. Please review this error.
    `
    throw Error(errorMessage);
  }

  console.log(`onEdit 5 -> targetRow: ${targetRow} found by \`findMemberByEmail()\``);

  updateFeeInfo_(e, thisSheetName, targetRow, targetSheet);
  console.log(`onEdit 6 -> successfully completed trigger check`);
}


/**
 * @param {Event} e  Event Object from `onEdit`.
 * @param {SpreadsheetApp.Sheet} sheet  Sheet where edit occurred.
 */

function verifyLegalEditInRange_(e, sheet) {
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

  // Exit if we're out of range
  if (thisRow < feeEditRange.top || thisRow > feeEditRange.bottom) {
    Logger.log("Row is out of bounds");
    return false;
  }
  if (thisCol < feeEditRange.leftmost || thisCol > feeEditRange.rightmost) {
    Logger.log("Column is out of bounds");
    return false;
  }

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

function updateFeeInfo_(e, sourceSheetName, targetRow, targetSheet) {
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

    // Only modify payment history if isPaid=true.
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
 * @update  Apr 9, 2025
 * 
 */

function copyFilloutRegToMain_(registration, row = getLastSubmissionInMain()) {
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

  // Add additional information for payments occuring in other forms
  if (registrationObj['paymentAmount'] == 0) {
    const event = (registrationObj['referral'])['sources'] ?? "";
    const method = registrationObj['paymentMethod'];

    // Set new value
    registrationObj['paymentMethod'] = `${[event, method].join(': ')}`;
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


function copyToMainFromMaster(row) {
  //@todo: complete function!!!
}

function notifyNewAppSubmission(row) {
  //@todo: complete function!!!
  //STEP 1: Notify member to complete reg
}

