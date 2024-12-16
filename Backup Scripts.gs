/* SHEET CONSTANTS */
const BACKUP_NAME = 'BACKUP';
const BACKUP_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP_NAME);


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

