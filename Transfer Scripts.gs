/* SHEET CONSTANTS */
const BACKUP_NAME = 'BACKUP';
const BACKUP_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP_NAME);

// Function to get column mappings
const GET_COL_MAP = (sheet) => { return sheetColumnMappings[sheet] || null };

const SHEET_COL_MAP = {
  [SHEET_NAME]: {
    emailCol: EMAIL_COL,
    memberIdCol: MEMBER_ID_COL,
    feeStatus: IS_FEE_PAID_COL,
    collectionDate: COLLECTION_DATE_COL,
    collector: COLLECTION_PERSON_COL,
    isInternalCollected: IS_INTERNAL_COLLECTED_COL,
  },
  [MASTER_NAME]: {
    emailCol: MASTER_EMAIL_COL,
    memberIdCol: MASTER_MEMBER_ID_COL,
    feeStatus: MASTER_FEE_STATUS,
    collectionDate: MASTER_COLLECTION_DATE,
    collector: MASTER_FEE_COLLECTOR,
    isInternalCollected: MASTER_IS_INTERNAL_COLLECTED,
  },
};


function onEdit(e) {
  // Get details of edit event's sheet
  const thisSheet = e.range.getSheet();
  const thisSheetName = thisSheet.getName();

  // Check if legal sheet
  if(thisSheetName != SHEET_NAME || thisSheetName != MASTER_NAME) return;

  // Check if legal edit
  if(!verifyLegalEditInRange(e, thisSheet)) return;

  // Get the email column for the current sheet
  const thisEmailCol = GET_COL_MAP(thisSheetName).emailCol;
  const thisRow = e.range.getRow();

  // Get email from `thisRow` and `thisEmailCol`
  const email = thisSheet.getRange(thisRow, thisEmailCol).getValue();

  const isMainSheet = (thisSheetName == SHEET_NAME);

  const sourceSheet = isMainSheet ? MAIN_SHEET : MASTER_SHEET;
  const targetSheet = isMainSheet ? MASTER_SHEET : MAIN_SHEET;
  const targetRow = isMainSheet ? null : findMemberByBinarySearch(email)
    
  updateFeeInfo(e, sourceSheet, targetRow, targetSheet);
}


/**
 * @param {Event} e  Event Object from `onEdit`.
 * @param {SpreadsheetApp.Sheet} sheet  Sheet where edit occurred.
 */

function verifyLegalEditInRange(e, sheet) {
  const sheetName = sheet.getName();
  var thisRow = e.range.getRow();
  var thisCol = e.range.getColumn();
  
  // Function to get column mappings
  const feeStatus = GET_COL_MAP(sheetName).feeStatus;
  const isInternalCollected = GET_COL_MAP(sheetName).isInternalCollected;

  Logger.log(`Now in verifyLegalEditInRange... feeStatusCol : ${feeStatus}, isInternalCollected : ${isInternalCollected}`);
  
  const feeEditRange = {
    top : 2,    // Skip header row
    bottom : sheet.getLastRow(),
    leftmost : feeStatus,
    rightmost : isInternalCollected,
  }

  // Helper function to log error message
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
 * @param {SpreadsheetApp.Sheet} sourceSheet  Source sheet to extract fee info.
 * @param {number} targetRow  Target row to update.
 * @param {SpreadsheetApp.Sheet} targetSheet  Target sheet to update fee info.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 16, 2024
 * @update  Dec 16, 2024
 * 
 */

function updateFeeInfo(e, sourceSheet, targetRow, targetSheet) {
  const thisRange = e.range;
  const thisCol = e.range.getColumn();

  const sourceCols = getColsFromSheet(sourceSheet);
  const targetCols = getColsFromSheet(targetSheet);

  const getTargetCol = (sourceCol) => {
    switch(sourceCol) {
      case(sourceCols.feeStatus) : return targetCols.feeStatus;
      case(sourceCols.collectionDate) : return targetCols.collectionDate;
      case(sourceCols.collector) : return targetCols.collector;
      case(sourceCols.isInternalCollected) : return target.isInternalCollected;
    }
  };

  const targetCol = getTargetCol(thisCol);
  const targetRange = targetSheet.getRange(targetRow, targetCol);
  thisRange.copyTo(targetRange, {contentsOnly: true});
}
  

function updateFeeInfo2_(sourceRow, sourceSheet, targetRow, targetSheet) {
  sourceCols = getColsFromSheet(sourceSheet);
  targetCols = getColsFromSheet(targetSheet);

  const transferValue = (fromCell, toCell) => {
    let fromValue = fromCell.getValue();
    Logger.log("Previous value: " + toValue.getValue());    // Add previous value to execution log
    toCell.setValue(fromValue);
  }

  oldFeeStatusRange = targetSheet.getRange(targetRow, targetCols.feeStatus);
  newFeeStatusRange = sourceSheet.getRange(sourceRow, sourceSheet.feeStatus);
  transferValue(oldFeeStatusRange, newFeeStatusRange);

  oldCollectionDateRange = targetSheet.getRange(targetRow, targetCols.collectionDate);
  newCollectionDateRange = source.getRange(row, sourceSheet.collectionDate);

  oldCollector = targetSheet.getRange(targetRow, targetCols.collector);
  newCollector = source.getRange(sourceRow, sourceSheet.collector);

  oldInternalCollectedRange = targetSheet.getRange(targetRow, targetCols.isInternalCollected);
  newInternalCollectedRange = source.getRange(sourceRow, sourceSheet.isInternalCollected);
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

