const CELL_EDIT_LIMIT = 4;   // set number of cells that can be edited at once

function onEdit(e) {
  // Get details of edit event's sheet
  const thisRange = e.range;
  const thisSheet = thisRange.getSheet();
  const thisSheetName = thisSheet.getName();

  var debug_e = {
    //authMode:  e.authMode,
    range:  e.range.getA1Notation(),
    sheetName : e.range.getSheet().getSheetName(),
    //source:  e.source,
    value:  e.value,
    oldValue: e.oldValue
  }
  console.log({test: 2, eventObject: debug_e});

  if(thisRange.getNumRows() > 2) return;  // prevent sheet-wide changes
  else if(thisRange.getNumColumns() > CELL_EDIT_LIMIT) {
    // TODO: add function to individually process changes
    Logger.log(`More than ${CELL_EDIT_LIMIT} columns edited at once`);
  }

  console.log(`onEdit 1 -> thisSheetName: ${thisSheetName}`);
  
  // Check if legal sheet
  if(thisSheetName != SHEET_NAME && thisSheetName != MASTER_NAME) return;

  console.log("onEdit 1a -> Passed first check");

  //if(e.value == e.oldValue) return;   // Values have not changed. Edit was on sheet formatting.

  console.log("onEdit 1b -> Passed second check");

  // Check if legal edit
  if(!verifyLegalEditInRange(e, thisSheet)) return;

  console.log("onEdit 2 -> Passed \`verifyLegalEditInRange()\`");

  // Get the email column for the current sheet
  const thisEmailCol = GET_COL_MAP_(thisSheetName).emailCol;
  const thisRow = e.range.getRow();

  console.log(`onEdit 3 -> thisEmailCol: ${thisEmailCol} thisRow: ${thisRow}`);

  // Get email from `thisRow` and `thisEmailCol`
  const email = thisSheet.getRange(thisRow, thisEmailCol).getValue();

  const isMainSheet = (thisSheetName == SHEET_NAME);
  console.log(`onEdit 4 -> email: ${email} isMainSheet: ${isMainSheet}`);

  const sourceSheet = isMainSheet ? MAIN_SHEET : MASTER_SHEET;
  const targetSheet = isMainSheet ? MASTER_SHEET : MAIN_SHEET;
  const targetRow = findMemberByEmail(email, targetSheet);  // Find row of member in `targetSheet` using their email

  // Throw error message if member not in `targetSheet`
  if(targetRow == null) {
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
    top : 2,    // Skip header row
    bottom : sheet.getLastRow(),
    leftmost : feeStatus,
    rightmost : isInternalCollected,
  }

  // Helper function to log error message and exit function
  const logAndExitFalse = (cell) => { Logger.log(`${cell} is out of bounds`); return false; }

  // Exit if we're out of range
  if (thisRow < feeEditRange.top || thisRow > feeEditRange.bottom) logAndExitFalse("Row");
  if (thisCol < feeEditRange.left || thisCol > feeEditRange.right) logAndExitFalse("Column");
  
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
    switch(source) {
      case(sourceCols.feeStatus) : return targetCols.feeStatus;
      case(sourceCols.collectionDate) : return targetCols.collectionDate;
      case(sourceCols.collector) : return targetCols.collector;
      case(sourceCols.isInternalCollected) : return targetCols.isInternalCollected;
    }
  };

  // Find which column was edited in `sourceSheet` and find respective col in `targetSheet`
  const targetCol = getTargetCol(thisCol);
  Logger.log(`updateFeeInfo 2 -> targetRow: ${targetRow} targetCol: ${targetCol}`);

  const targetRange = targetSheet.getRange(targetRow, targetCol);

  // Special case: MASTER stores payment history as semesterCode(s).
  // If isPaid, then add semesterCode to payment history, i.e. bool -> str
  // Otherwise, nothing to modify in MASTER for member's payment history
  if(targetSheetName == MASTER_NAME && targetCol == MASTER_PAYMENT_HIST) {
    console.log("updateFeeInfo 3 -> entering if statement");
    const value = thisRange.getValue() || "";
    const isPaid = parseBool(value);    // convert to bool
    console.log(`updateFeeInfo 3b -> Value: ${value} isPaid: ${isPaid}`);

    // Only modify payment history if isPaid == true.
    if(isPaid) {
      console.log("updateFeeInfo 3c -> entering isPaid");
      addPaidSemesterToHistory(targetRow, sourceSheetName);
    }
    else {
      console.log("updateFeeInfo 3c -> entering NOT(isPaid)");
    }
    
  }
  else if(sourceSheetName == MASTER_NAME && thisCol == MASTER_PAYMENT_HIST) {
    // CASE 2: Add history payment to sheet
    console.log("updateFeeInfo 3 ->  entering else if statement");
    const paymentHistory = thisRange.getValue() || "";
    updateIsFeePaid(paymentHistory, targetRow, targetCol, targetSheet);
  }
  else {
    console.log("updateFeeInfo 3 ->  entering else statement");
    thisRange.copyTo(targetRange, {contentsOnly: true});
  }

  console.log("updateFeeInfo 4 ->  finished updating payment history");
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 18, 2023
 * @update: Oct 18, 2023
 * 
 * Triggered by `checkURLFromIndex` when PassKit URL in `BACKUP` exists & !isCopied
 * 
 */

function copyToMain(url, targetEmail) {
  const mainSheet = MAIN_SHEET;
  const EMAIL_COL = 1;
  const URL_COL = 21;
  
  var mainData = mainSheet.getDataRange().getValues();
  var targetData = [];
  var mainRow, email, i;

  // Loop through rows in `main` until matching email entry is found
  for (i = 0; i < mainData.length; i++) {
    mainRow = mainData[i];
    email = mainRow[EMAIL_COL]; // get email using email column index
    
    if (email === targetEmail) {
      targetData.push(mainRow);
      break; // Exit the loop once the matching row is found
    }
  }

  // Copy URL when match is found
  if (targetData.length > 0) {
    targetRow = i + 1;
    mainSheet.getRange(targetRow, URL_COL).setValue(url);
    return true;
  }

  return false;
}

