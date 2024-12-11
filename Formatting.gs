/**
 * Trim whitespace from specific columns in last row of `MAIN_SHEET`.
 * 
 * Range is FIRST_NAME_COL to REFERAL_COL (7 columns).
 * 
 * @trigger New form submission
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Nov 22, 2024
 */

function trimWhitespace_() {
  const sheet = MAIN_SHEET;
  
  const lastRow = sheet.getLastRow();
  const rangeToFormat = sheet.getRange(lastRow, FIRST_NAME_COL, 1, 7);
  rangeToFormat.trimWhitespace();

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


///  👉 FUNCTIONS APPLIED TO MAIN_SHEET 👈  \\\

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
 * @update  Nov 22, 2024
 */

function formatMainView() {
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
 * Set letter case of specific columns in member entry as following:
 *  - Lower Case: [McGill Email Address] 
 *  - Capitalized: [First Name, Last Name, Preferred Name/Pronouns, Year, Program]
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()] 
 *                    Row number to target fix.
 *                    Defaults to last row (1-indexed).
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Dec 11, 2024
 * 
 */

function fixLetterCaseInRow_(row=MASTER_SHEET.getLastRow()) {
  const sheet = MAIN_SHEET;
  const lastRow = getLastSubmissionInMain();

  // Set to lower case
  const rangeToLowerCase = sheet.getRange(lastRow, EMAIL_COL);
  const rawValue = rangeToLowerCase.getValue().toString();
  rangeToLowerCase.setValue(rawValue.toLowerCase());

  // Set to Capitalized (first letter of word is UPPER)
  const rangeToCapitalize = sheet.getRange(lastRow, FIRST_NAME_COL, 1, 5);
  const valuesToCapitalize = rangeToCapitalize.getValues()[0]   // Flatten array
  
  // Capitalize each value in array
  const capitalizedValues = valuesToCapitalize.map(value => 
    value.replace(/\b\w/g, l => l.toUpperCase())
  );

  // Now replace raw values with capitalized ones
  rangeToCapitalize.setValues([ capitalizedValues ]); // setValues requires 2d array
}


///  👉 FUNCTIONS APPLIED TO MASTER_SHEET 👈  \\\

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
 * Formats `MASTER_SHEET` for simple and uniform UX.
 * 
 * Remove whitespace from `McGill Email Address` to  `Referral`
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 22, 2024
 * @update  Nov 22, 2024
 */

function formatMasterView() {
  var sheet = MASTER_SHEET;
  
}


/**
 * Clean latest member registration in `MASTER_SHEET`.
 * 
 * Data normalization includes:
 * 
 *  - Trim whitespace
 *  - Capitalize selected values e.g. name, year, program
 *  - Insert fee status formula in `Fee Paid` col
 *  - Format collection date correctly; append semester code if applicable
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 22, 2024
 * @update  Nov 22, 2024
 */

function cleanMasterRegistration() {
  var sheet = MASTER_SHEET;
  const lastRow = sheet.getLastRow();

  // STEP 1: Trim white space from `Email` col to `Referral` col
  const rangeToTrim = sheet.getRange(lastRow, MASTER_EMAIL_COL, 1, 9);
  rangeToTrim.trimWhitespace();

  // STEP 2: Capitalize selected value
  const rangeToCapitalize = sheet.getRange(lastRow, MASTER_FIRST_NAME_COL, 1, 5);
  var valuesToCapitalize = rangeToCapitalize.getValues()[0]; // Get all the values as 1D arr

  valuesToCapitalize.forEach((cell, colIndex) => {
    if (typeof cell === "string") {   // Ensure it's a string before capitalizing
      valuesToCapitalize[colIndex] = cell
        .toLowerCase()
        .replace(/\b\w/g, l => l.toUpperCase());
    }
  });

  // Replace values with formatted values
  rangeToCapitalize.setValues([valuesToCapitalize]);  // setValues() requires 2D array

}


/**
 * Recursive function to search for entry by email in `MASTER` sheet using binary search.
 * Returns email's row index in GSheet (1-indexed), or null if not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Nov 22, 2024
 * @update  Nov 22, 2024
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()]  
 *                      The starting row index for the search (1-indexed). 
 *                      Defaults to 1 (the first row).
 * 
 */

function formatFeeCollection(row=MASTER_SHEET.getLastRow()) {
  var sheet = MASTER_SHEET;

  // STEP 1: Check for current fee status to flag for later
  const rangeFeeStatus = sheet.getRange(row, MASTER_FEE_STATUS);
  const feeStatus = rangeFeeStatus.getValue().toString();

  const regex = new RegExp('unpaid', "i"); // Case insensitive
  const isUnpaid = feeStatus.includes("unpaid")           // FIX LINE!!
  
  //.search(regex);
  
  // STEP 2: Insert fee status formula in `Fee Paid` col
  rangeFeeStatus.setFormula(isFeePaidFormula);    // Formula found in `Semester Variables.gs`

  // If feeStatus is unpaid, formatting is completed.
  if(isUnpaid) return;

  // STEP 3: Format collection date correctly;
  const rangeCollectionDate = sheet.getRange(row, MASTER_COLLECTION_DATE);
  const collectionDate = rangeCollectionDate.getValue();   // Format is yyyy-mm-dd hh:mm

  const formattedDate = Utilities.formatDate(collectionDate, TIMEZONE, 'yyyy-MM-dd');
  rangeCollectionDate.setValue(formattedDate);

  // STEP 4: Append semester code if collection date non-empty
  if(!collectionDate) return;

  const rangePaymentHistory = sheet.getRange(row, MASTER_PAYMENT_HIST);
  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangePaymentHistory.setValue(semCode);
}


function insertRegistrationSem(row=MASTER_SHEET.getLastRow()) {
  var sheet = MASTER_SHEET;
  const rangeLatestRegSem = sheet.getRange(row, MASTER_LAST_REG_SEM);

  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangeLatestRegSem.setValue(semCode);
}


///  👉 FUNCTIONS FOR MEMBER ID ENCODING 👈  \\\

/**
 * Create Member ID in last row of `MAIN_SHEET`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update  Oct 20, 2024
 */

function encodeLastRow() {
  const sheet = MAIN_SHEET;
  const newSubmissionRow = getLastSubmissionInMain();
  
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
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to target
 * @param {integer} [row=sheet.getLastRow()]    The 1-indexed row in input `sheet`. 
 *                                              Defaults to the last row in the sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Dec 11, 2024
 */

function encodeByRow(sheet, row=sheet.getLastRow()) {
  let sheetCols = getColsToEncode(sheet);

  const email = sheet.getRange(row, sheetCols.emailCol).getValue();
  if (email === "") throw RangeError("Invalid index access");   // check for invalid index

  const member_id = MD5(email);
  sheet.getRange(row, sheetCols.memberIdCol).setValue(member_id);
}


/**
 * Retrieves column indices of email and member id in GSheet.
 * 
 * Helper for encoding functions (i.e. `encodeList`, `encodeByRow`)
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 * @return {{emailCol, memberIdCol}}  Returns col indices for `sheet`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2024
 * @update  Nov 13, 2024
 */

function getColsToEncode(sheet) {
  let sheetCols = {emailCol: -1, memberIdCol: -1};    // starter values

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

