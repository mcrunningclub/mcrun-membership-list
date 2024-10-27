/* SHEET CONSTANTS */
const BACKUP_NAME = 'BACKUP';
const BACKUP_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP_NAME);

/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 18, 2023
 * 
 * Copy to 'BACKUP' sheet for Passkit + Zapier Automation.
 * Zapier automation only works when latest submission is on the last row.
 * 
 * Zapier automation: 1. Create New Member Pass ➜ 2. Return PassKit URL
 * 
 */
function copyToBackup() {
  const mainSheet = MAIN_SHEET;
  const backupSheet = BACKUP_SHEET;
  const IS_COPIED_COL = 22;

  // Get size of last row
  const numCol = mainSheet.getLastColumn();
  const lastRow = mainSheet.getLastRow();

  // Get `SpreadsheetApp.Range` object
  const sourceRow = mainSheet.getRange(lastRow, 1, 1, numCol);
  var rangeBackup = backupSheet.getRange(backupSheet.getLastRow() +1, 1, 1, numCol);
  
  // Copy and paste values
  const valuesToCopy = sourceRow.getValues();
  rangeBackup.setValues(valuesToCopy);

  backupSheet.getRange(backupSheet.getLastRow(), IS_COPIED_COL).setValue(false); // set flag to false
  
  return;
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 18, 2023
 * @update: Oct 18, 2023
 * 
 * Triggered when PassKit URL is added for latest submission in `BACKUP`.
 * Copies PassKit URL to respective member in `main`.
 * 
 */

function onEditPasskitURL() { 
  var BACKUP_SHEET;
  return;
  var sheet = BACKUP_SHEET;
  checkURLFromIndex(sheet.getLastRow());
  return;
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 18, 2023
 * @update: Oct 18, 2023
 * 
 * Triggered every 4 hours to ensure all URL are copied to `main`.
 */

function checkAllURL() {
  const sheet = BACKUP_SHEET;
  const IS_COPIED_COL = 22;

  for(var i = 2; i <= sheet.getLastRow(); i++) {
    var isCopied = sheet.getRange(i, IS_COPIED_COL).getValue();
    if (!isCopied) checkURLFromIndex(i);
  }

  return;
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 18, 2023
 * @update: Oct 18, 2023
 * 
 * Check row at `rowIndex` if PassKit URL already copied. Otherwise transfer URL to `main`.
 * 
 */

function checkURLFromIndex(rowIndex) {
  const sheet = BACKUP_SHEET;
  const EMAIL_COL = 2;
  const URL_COL = 21;
  const IS_COPIED_COL = 22;

  // Exit if URL data already copied to `main`
  const rangeIsCopied = sheet.getRange(rowIndex, IS_COPIED_COL);
  if ( rangeIsCopied.getValue() ) return;

  // Get member email and PassKit URL from row at `rowIndex`
  const memberEmail = sheet.getRange(rowIndex, EMAIL_COL).getValue();
  const url = sheet.getRange(rowIndex, URL_COL).getValue();
  
  // Only copy PassKit URL when nonempty
  if( String(url).length > 0) {
    const res = copyToMain(url, memberEmail);   // args: URL, email
    sheet.getRange(rowIndex, IS_COPIED_COL).setValue(res);  // Toggle flag to `TRUE` if URL copied successfully
  }
  return;
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


// DRAFT FUNCTIONS!
/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 18, 2023
 * 
 * Triggered when PassKit URL is added for latest submission in `BACKUP`
 * Zapier automation only works when latest submission is on the last row.
 * 
 * Zapier automation : 1. Create New Member Pass ➜ 2. Return PassKit URL
 * 
 */

// function onEditPasskitURL() {
//   const sheet = BACKUP_SHEET;
//   const EMAIL_COL = 2;
//   const URL_COL = 21;
//   const IS_COPIED_COL = 22;

//   const backupRowNum = sheet.getLastRow();  // index of latest submission
//   const memberEmail = sheet.getRange(backupRowNum, EMAIL_COL).getValue();

//   // Exit if URL data already copied to `main`
//   const rangeIsCopied = sheet.getRange(backupRowNum, IS_COPIED_COL);
//   if (rangeIsCopied.getValue()) return;

//   const url = sheet.getRange(backupRowNum, URL_COL).getValue();   // url for latest submission
  
//   // Only copy PassKit URL when nonempty
//   if( String(url).length > 0) {
//     const res = copyToMain(url, memberEmail);   // args: URL, email
//     sheet.getRange(backupRowNum, IS_COPIED_COL).setValue(res);  // Toggle flag to `TRUE` if URL copied successfully
//   }
//   return;
// }

