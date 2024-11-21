/**
 * Trim whitespace from specific columns in last row of `MAIN_SHEET`.
 * 
 * Range is FIRST_NAME_COL to REFERAL_COL (7 columns).
 * 
 * @trigger New form submission
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Oct 17, 2023
 */

function trimWhitespace_() {
  const sheet = MAIN_SHEET;
  
  const lastRow = sheet.getLastRow();
  const rangeNames = sheet.getRange(lastRow, FIRST_NAME_COL, 1, 7);
  rangeNames.trimWhitespace();

  return;
}


/**
 * Returns reg expression for target string.
 * 
 * @input {string}  targetSubstring  String used to create regex.
 * @return {RegExp}   Returns regular expression.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 8, 2023
 */

function getRegEx_(targetSubstring) {
  return RegExp(targetSubstring, 'g');
}


/**
 * Sorts `MAIN_SHEET` by first name, then last name.
 * 
 * @trigger New form submission.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Jun 1, 2024
 */

function sortMainByName() {
  const sheet = MAIN_SHEET;

  const numRows = sheet.getLastRow() - 1;     // Remove header row from count
  const numCols = sheet.getLastColumn();
  
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);
  
  // Sorts values by `First Name` then by `Last Name`
  range.sort([{column: 3, ascending: true}, {column: 4, ascending: true}]);
  return;
}


/**
 * Formats `MAIN_SHEET` for simple and uniform UX.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Oct 18, 2023
 */

function formatSpecificColumns() {
  var sheet = MAIN_SHEET;

  const rangeRegistration = sheet.getRange('A2:A');  // Range for Preferred Name/Pronouns
  const rangePreferredName = sheet.getRange('E2:E');  // Range for Preferred Name/Pronouns
  const rangeWaiver = sheet.getRange('J2:J');         // Range for Waiver
  const rangePaymentChoice = sheet.getRange('K2:K');  // Range for Payment Preferrence
  const rangeInteracRef = sheet.getRange('L2:L');     // Range for Interac e-Transfer Reference
  const rangeCollection = sheet.getRange('O2:P');     // Range for Collection Info
  const rangeMemberId = sheet.getRange('T2:T');       // Range for Member Id
  const rangeURL = sheet.getRange('U2:U');            // Range for PassKit URL

  // Set ranges to Bold
  rangeRegistration.setFontWeight('bold');
  rangePreferredName.setFontWeight('bold');
  rangePaymentChoice.setFontWeight('bold');
  rangeInteracRef.setFontWeight('bold');
  rangeCollection.setFontWeight('bold');
  rangeMemberId.setFontWeight('bold');
  rangeURL.setFontWeight('bold');

  // Set Text Wrapping to 'Clip'
  rangePaymentChoice.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  rangeWaiver.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Align ranges to Left
  rangePaymentChoice.setHorizontalAlignment('left');

  // Centre these ranges
  rangeInteracRef.setHorizontalAlignment('center');
  rangeCollection.setHorizontalAlignment('center');
  rangeMemberId.setHorizontalAlignment('center');
}


/**
 * Sorts `MASTER` by email instead of first name.
 * Required to ensure `findSubmissionByEmail` works properly.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 27, 2024
 */

function sortMasterByEmail() {
  const sheet = MASTER_SHEET;
  const numRows = sheet.getLastRow() - 1;   // Remove Header from count
  const numCols = sheet.getLastColumn();
    
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);
    
  // Sorts values by email
  range.sort([{column: 1, ascending: true}]);
  return;
}


/**
 * Create Member ID in last row of `MAIN_SHEET`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 20, 2024
 */

function encodeLastRow() {
  const sheet = MAIN_SHEET;
  const newSubmissionRow = getLastSubmission();
  
  const email = sheet.getRange(newSubmissionRow, EMAIL_COL).getValue();
  const member_id = MD5(email);
  sheet.getRange(newSubmissionRow, MEMBER_ID_COL).setValue(member_id);
}


/**
 * Create Member ID for every member in `sheet`.
 * 
 * @input {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Nov 13, 2024
 */

function encodeList(sheet) {
  let sheetCols = getColsToEncode(sheet);

  // Start at row 2 (1-indexed)
  for (var i = 2; i <= sheet.getMaxRows(); i++) {
    var email = sheet.getRange(i, sheetCols.emailCol).getValue();
    if (email === "") return;   // check for invalid row

    var member_id = MD5(email);
    sheet.getRange(i, sheetCols.memberIdCol).setValue(member_id);
  }
}


/**
 * Create single Member ID using row number from `sheet`.
 * 
 * @input {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 * @input {int} rowNumber  Row index in GSheet (1-indexed)
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Nov 13, 2024
 */

function encodeByRow(sheet, rowNumber) {
  let sheetCols = getColsToEncode(sheet);

  const email = sheet.getRange(rowNumber, sheetCols.emailCol).getValue();
  if (email === "") throw RangeError("Invalid index access");   // check for invalid index

  const member_id = MD5(email);
  sheet.getRange(i, sheetCols.memberIdCol).setValue(member_id);
}


/**
 * Retrieves column indices of email and member id in GSheet.
 * 
 * Helper for encoding functions (i.e. `encodeList`, `encodeByRow`)
 * 
 * @input {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 * @return {{emailCol, memberIdCol}}  Returns col indices for `sheet`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2024
 * @update  Nov 13, 2024
 */

function getColsToEncode(sheet) {
  let sheetCols = {emailCol, memberIdCol};

  switch (sheet) {
    case MAIN_SHEET:
      sheetCols.emailCol = EMAIL_COL;
      sheetCols.memberIdCol = MEMBER_ID_COL;
    break;
    
    case MASTER_SHEET:
      sheetCols.emailCol = MASTER_EMAIL_COL;
      sheetCols.memberIdCol = MASTER_MEMBER_ID_COL;
    break;
  }

  return sheetCols;

}

