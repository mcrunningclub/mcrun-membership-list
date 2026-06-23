/**
 * Trims whitespace from specific columns in the last row of the semester sheet.
 * 
 * This function targets the range from `SEMESTER_COLS.FIRST_NAME` to `REFERRAL_COL` (7 columns).
 * It ensures that unnecessary whitespace is removed from the latest member entry.
 * 
 * @trigger New form submission
 * 
 * @param {number} [row=SEMESTER_SHEET.getLastRow()] - The row number to target for trimming.
 *                                                     Defaults to the last row in semester sheet.
 * 
  * 
 * @see {@link fixRowCaseSemester_} for additional formatting applied to the same row.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Feb 5, 2025
 */

function trimWhitespaceSemester_(row = getLastSubmissionInSemester()) {
  const sheet = SEMESTER_SHEET;
  const rangeToFormat = sheet.getRange(row, SEMESTER_COLS.FIRST_NAME, 1, 7);
  rangeToFormat.trimWhitespace();
}


///  👉 FUNCTIONS APPLIED TO SEMESTER_SHEET 👈  \\\

/**
 * Sorts semester sheet by first name, then last name.
 * 
 * This function organizes the data in the semester sheet by sorting rows alphabetically
 * based on the `First Name` column and then the `Last Name` column.
 * 
 * @trigger New form submission or McRUN menu.
 * 
 * @see {@link tryAndSortSemester} for a safe way to sort with locking.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Jan 11, 2025
 */

function sortSemesterByName_() {
  const sheet = SEMESTER_SHEET;

  const numRows = sheet.getLastRow() - 1;   // Remove header row from count
  const numCols = sheet.getLastColumn();

  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by `First Name` then by `Last Name`
  range.sort([{ column: 3, ascending: true }, { column: 4, ascending: true }]);
}


/**
 * Sorts semester sheet only if the lock is free.
 * 
 * This function prevents concurrent processes from interfering with sorting
 * by acquiring a script lock before proceeding. If the lock is unavailable,
 * it logs a message and exits gracefully.
 *  
 * @see {@link sortSemesterByName_} for the actual sorting logic.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Mar 15, 2025
 * @update  Mar 15, 2025
 */

function tryAndSortSemester() {
  const lock = LockService.getScriptLock();

  // Try getting lock for up to 10 seconds
  if (lock.tryLock(10000)) {
    try {
      sortSemesterByName_();
      formatSemester();
    } finally {
      lock.releaseLock();
    }
  } else {
    console.log("Another script is running. Unable to sort now");
  }
}


/**
 * Formats semester sheet for a simple and uniform user experience.
 * 
 * - Freezing panes
 * - Adjusting font styles, sizes, and weights
 * - Setting column widths
 * - Applying number formats and text wrapping
 * - Aligning text horizontally and vertically
 * - Adding checkboxes to specific columns
 * - Ensuring proper letter casing for names and email addresses * 
 * - Adding hyperlinks to waivers
 * - Formatting collection dates
 * 
 * @see {@link sortSemesterByName_} for sorting logic applied to the same sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 5, 2025
 */

function formatSemester() {
  const sheet = SEMESTER_SHEET;
  const allCols = Object.entries(SEMESTER_COLS).length;
  const allRowsExceptHeader = sheet.getLastRow() - 1;

  //helper functions
  function getHeaderCell(column) {
    return sheet.getRange(1, column);
  }

  function getHeaderRow() {
    return sheet.getRange(1, 1, 1, allCols);
  }

  function getColumnExceptHeader(column) {
    return sheet.getRange(2, column, allRowsExceptHeader, 1);
  }

  function getColumnIncludingHeader(column) {
    return sheet.getRange(1, column, sheet.getLastRow(), 1);
  }

  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 2. Bold formatting
  getHeaderRow().setFontWeight('bold');
  var columns = [
    SEMESTER_COLS.REGISTRATION_DATE,
    SEMESTER_COLS.PREFERRED_NAME,
    SEMESTER_COLS.PAYMENT_METHOD,
    SEMESTER_COLS.INTERAC_REF,
    SEMESTER_COLS.COLLECTION_DATE,
    SEMESTER_COLS.COLLECTED_BY,
    SEMESTER_COLS.MEMBER_ID];
  for (const col of columns) {
    console.log(col);
    getColumnExceptHeader(col).setFontWeight('bold');
  }

  // 3. Font size adjustments
  getHeaderRow().setFontSize(11);
  columns = [
    SEMESTER_COLS.PREFERRED_NAME,
    SEMESTER_COLS.FEE_PAID,
    SEMESTER_COLS.ATTENDANCE_STATUS
  ];
  for (const col of columns) {
    getHeaderCell(col).setFontSize(10);
  }
  getHeaderCell(SEMESTER_COLS.IS_INTERNAL_COLLECTED).setFontSize(9);
  getColumnExceptHeader(SEMESTER_COLS.MEMBER_ID).setFontSize(9);
  getColumnIncludingHeader(SEMESTER_COLS.PAYMENT_METHOD).setFontSize(8);

  // 4. Font family adjustment for member ID
  getColumnExceptHeader(SEMESTER_COLS.MEMBER_ID).setFontFamily('Google Sans Mono');

  // 5. Format collection date
  getColumnExceptHeader(SEMESTER_COLS.REGISTRATION_DATE).setNumberFormat('yyyy-MM-dd hh:mm:ss');
  getColumnExceptHeader(SEMESTER_COLS.COLLECTION_DATE).setNumberFormat('mmm d, yyyy');

  // 6. Text wrapping set to 'Clip' (for Referral + Waiver + Payment Choice)
  columns = [
    SEMESTER_COLS.REFERRAL,
    SEMESTER_COLS.WAIVER,
    SEMESTER_COLS.PAYMENT_METHOD
  ]
  for (const col of columns) {
    getColumnExceptHeader(col).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // 7. Horizontal and vertical alignment
  columns = [
    SEMESTER_COLS.OPTED_INTO_NEWSLETTER,
    SEMESTER_COLS.OPTED_INTO_EMAILS,
    SEMESTER_COLS.INTERAC_REF,
    SEMESTER_COLS.FEE_PAID,
    SEMESTER_COLS.COLLECTION_DATE,
    SEMESTER_COLS.COLLECTED_BY,
    SEMESTER_COLS.IS_INTERNAL_COLLECTED,
    SEMESTER_COLS.ATTENDANCE_STATUS,
    SEMESTER_COLS.MEMBER_ID
  ]
  for (const col of columns) {
    getColumnExceptHeader(col).setHorizontalAlignment('center');
  }
  getColumnExceptHeader(SEMESTER_COLS.REGISTRATION_DATE).setHorizontalAlignment('right');

  // 8. Column width mapping
  const sizeMap = {
    [SEMESTER_COLS.REGISTRATION_DATE]: 140,
    [SEMESTER_COLS.EMAIL]: 245,
    [SEMESTER_COLS.FIRST_NAME]: 115,
    [SEMESTER_COLS.LAST_NAME]: 115,
    [SEMESTER_COLS.YEAR]: 120,
    [SEMESTER_COLS.YEAR]: 90,
    [SEMESTER_COLS.PROGRAM]: 240,
    [SEMESTER_COLS.DESCRIPTION]: 400,
    [SEMESTER_COLS.REFERRAL]: 145,
    [SEMESTER_COLS.WAIVER]: 185,
    [SEMESTER_COLS.PAYMENT_METHOD]: 155,
    [SEMESTER_COLS.INTERAC_REF]: 155,
    [SEMESTER_COLS.FEE_PAID]: 75,
    [SEMESTER_COLS.COLLECTION_DATE]: 150,
    [SEMESTER_COLS.COLLECTED_BY]: 160,
    [SEMESTER_COLS.IS_INTERNAL_COLLECTED]: 65,
    [SEMESTER_COLS.COMMENTS]: 255,
    [SEMESTER_COLS.ATTENDANCE_STATUS]: 125,
    [SEMESTER_COLS.MEMBER_ID]: 140,
  };

  // Resize columns based on `sizeMap`
  Object.entries(sizeMap).forEach(([col, width]) => {
    sheet.setColumnWidth(col, width);
  });
}


/**
 * Adds checkboxes to specific columns in the last row of semester sheet.
 * 
 * This function is used to ensure that the last row of semester sheet has checkboxes
 * in the `Fee Paid`, `Given to Internal`, and `Attendance Status` columns.
 * 
 * @param {number} row  Row number to target for formatting.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 5, 2025
 */
function addCheckboxSemester_(row) {
  const sheet = SEMESTER_SHEET;

  // Add checkboxes to target columns
  [SEMESTER_COLS.FEE_PAID,
    SEMESTER_COLS.IS_INTERNAL_COLLECTED,
    SEMESTER_COLS.ATTENDANCE_STATUS
  ].forEach(col => sheet.getRange(row, col).insertCheckboxes());

  // Copy the list item  in 'Collection Person' col from first entry
  //var collectorItem = sheet.getRange(5, SEMESTER_COLS.COLLECTED_BY).getDataValidation();
  //var targetCell = sheet.getRange(row, SEMESTER_COLS.COLLECTED_BY);

  // Set the collector item
  //targetCell.setDataValidation(collectorItem);
}


/**
 * Set letter case of specific columns in member entry as following:
 *  - Lower Case: [McGill Email Address] 
 *  - Capitalized: [First Name, Last Name, Preferred Name/Pronouns, Year, Program]
 * 
 * @param {number} [row=getLastSubmissionInMain()]  Row number to target fix.
 *                                                  Defaults to last row (1-indexed).
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 11, 2024
 * @update  Mar 15, 2025
 */

function fixRowCaseSemester_(row = getLastSubmissionInSemester()) {
  const sheet = SEMESTER_SHEET;

  // Set to lower case
  const rangeToLowerCase = sheet.getRange(row, SEMESTER_COLS.EMAIL);
  const rawValue = rangeToLowerCase.getValue().toString();
  rangeToLowerCase.setValue(rawValue.toLowerCase());

  // Set to Capitalized (first letter of word is UPPER)
  const rangeToCapitalize = sheet.getRange(row, SEMESTER_COLS.FIRST_NAME, 1, 5);

  // Capitalize each value in array
  const capitalizedValues = rangeToCapitalize.getValues()[0].map(
    value => formatName(value)
  );

  // Now replace raw values with capitalized ones
  rangeToCapitalize.setValues([capitalizedValues]);

  // Helper function
  function formatName(name) {
    // Step 1: Trim whitespace from the beginning and end
    let formattedName = name.trim();

    // Step 2: Normalize the string to NFC (Canonical Composition)
    formattedName = formattedName.normalize("NFC");

    // Step 3: Split by spaces and hyphens to handle individual words
    formattedName = formattedName.split(/[\s-]+/).map(word => {
      // Capitalize the first letter and lowercase the rest of each word
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }).join(" ");

    // Step 4: Handle hyphenated names by rejoining with the correct hyphen case
    formattedName = formattedName.replace(/(\w)-(\w)/g, (match, p1, p2) => {
      return p1.toUpperCase() + '-' + p2.toLowerCase();
    });

    return formattedName;
  }
}


///  👉 FUNCTIONS APPLIED TO MASTER_SHEET 👈  \\\

/**
 * Sorts master sheet by email instead of first name.
 * Required to ensure `findSubmissionByEmail` works properly.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 27, 2024
 * @update  Jan 11, 2025
 */

function sortMasterByEmail() {
  const sheet = MASTER_SHEET;
  const numRows = sheet.getLastRow() - 1;   // Remove Header from count
  const numCols = sheet.getLastColumn();

  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by email
  range.sort([{ column: 1, ascending: true }]);
}


/**
 * Formats master sheet for simple and uniform UX.
 * 
 * Remove whitespace from `McGill Email Address` to `Referral`
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 22, 2024
 * @update  Dec 15, 2024
 */

function formatMaster() {
  var sheet = MASTER_SHEET;

  // Set Text to Bold
  const rangeListToBold = sheet.getRangeList([
    'K2:K',   // Latest Reg Semester
    'N2:N',   // Fee Paid
  ]);
  rangeListToBold.setFontWeight('bold');

  // Reduce Font to 9
  const rangeListReduceFont = sheet.getRangeList([
    'H2:H',   // Member Description
    'I2:I',   // Referral
    'J2:J',   // Latest Registration Timestamp
    'L2:L',   // Registration History
    'N2:N',   // Fee Paid
    'Q2:Q',   // Collection Date
    'R2:R',   // Given to Internal
    'S2:S',   // Payment History
    'T2:T',   // Comments
    'U2:U',   // Attendance Status
    'V2:V',   // Member ID
  ]);
  rangeListReduceFont.setFontSize(9).set;

  // Change Font Family to 'Google Sans Mono'
  const rangeListToGoogleSansMono = sheet.getRangeList([
    'L2:L',   // Registration History
    'N2:N',   // Fee Paid
    'S2:S',   // Payment History
    'V2:V',   // Member ID
  ]);
  rangeListToGoogleSansMono.setFontFamily('Google Sans Mono');

  // Change Font Family to Helvetica
  const rangeListToHelvetica = sheet.getRangeList([
    'H2:H',   // Member Description
    'O2:O',   // Fee Expiration
    'P2:P',   // Collected By
    'T2:T',   // Comments
  ]);
  rangeListToHelvetica.setFontFamily('Helvetica');

  // Change to Clip Wrap
  const rangeListToClipWrap = sheet.getRangeList([
    'E2:E',   // Year
    'F2:F',   // Program
    'G2:G',   // Waiver
    'H2:H',   // Member Description
  ]);
  rangeListToClipWrap.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Modify to Middle-Centre Alignment
  const rangeListToCenter = sheet.getRangeList([
    'J2:J',   // Latest Registration Timestamp
    'K2:K',   // Latest Registration Code
    'L2:L',   // Registration History
    'N2:N',   // Fee Paid
    'O2:O',   // Fee Expiration
    'P2:P',   // Collected By
    'Q2:Q',   // Collection Date
    'R2:R',   // Given to Internal
    'S2:S',   // Payment History
    'U2:U',   // Attendance Status
    'V2:V',   // Member ID
  ]);
  rangeListToCenter.setHorizontalAlignment('center');
  rangeListToCenter.setVerticalAlignment('middle');

  // Show Hyperlink for Waivers
  const rangeListShowHyperlink = sheet.getRangeList(['G2:G']);
  rangeListShowHyperlink.setShowHyperlink(true);

  // Set formatting type of collection date
  const rangeCollectionDate = sheet.getRange('Q2:Q');
  rangeCollectionDate.setNumberFormat('yyyy-mm-dd');
}


/**
 * Clean latest member registration in master sheet.
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

function cleanLastRowMaster() {
  var sheet = MASTER_SHEET;
  const lastRow = sheet.getLastRow();

  // STEP 1: Trim white space from `Email` col to `Referral` col
  const rangeToTrim = sheet.getRange(lastRow, MASTER_COLS.EMAIL, 1, 9);
  rangeToTrim.trimWhitespace();

  // STEP 2: Capitalize selected value
  const rangeToCapitalize = sheet.getRange(lastRow, MASTER_COLS.FIRST_NAME, 1, 5);
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
 * Formats fee collection date and semester for the specified row of the master sheet.
 * 
 * Changes date to 'yyyy-MM-dd'. No formatting is done if fee is not paid.
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()]  The starting row index for the search (1-indexed). 
 *                                                  Defaults to last row.
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Nov 22, 2024
 * @update  Dec 11, 2024
 */

function formatFeeCollection_(row = MASTER_SHEET.getLastRow()) {
  const sheet = MASTER_SHEET;

  // STEP 1: Check for current fee status to flag for later
  const rangeFeeStatus = sheet.getRange(row, MASTER_COLS.FEE_PAID);
  const feeStatus = rangeFeeStatus.getValue().toString();

  const regex = new RegExp('unpaid', "i"); // Case insensitive
  const isUnpaid = regex.test(feeStatus);

  // STEP 2: Insert fee status formula in `Fee Paid` col
  rangeFeeStatus.setFormula(isFeePaidFormula);    // Formula found in `Semester Variables.gs`

  // If feeStatus is unpaid, formatting is completed.
  if (isUnpaid) return;

  // STEP 3: Format collection date correctly;
  const rangeCollectionDate = sheet.getRange(row, MASTER_COLS.COLLECTION_DATE);
  const collectionDate = rangeCollectionDate.getValue();   // Format is yyyy-mm-dd hh:mm

  const formattedDate = Utilities.formatDate(collectionDate, TIMEZONE, 'yyyy-MM-dd');
  rangeCollectionDate.setValue(formattedDate);

  // STEP 4: Append semester code if collection date non-empty
  if (!collectionDate) return;

  const rangePaymentHistory = sheet.getRange(row, MASTER_COLS.PAYMENT_HISTORY);
  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from semester sheet
  rangePaymentHistory.setValue(semCode);
}

/**
 * Inserts the 3-char semester code for the registration in the specified row of the master sheet.
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()] 
 *                    The row number to target for inserting the semester code.
 *                    Defaults to the last row.
 * 
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 15, 2024
 * @update  Dec 15, 2024
 */

function insertRegistrationSem_(row = MASTER_SHEET.getLastRow()) {
  var sheet = MASTER_SHEET;
  const rangeLatestRegSem = sheet.getRange(row, MASTER_COLS.LATEST_REG_SEMESTER);

  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangeLatestRegSem.setValue(semCode);
}


///  👉 FUNCTIONS FOR MEMBER ID ENCODING 👈  \\\

/**
 * Create Member ID from input.
 * 
 * @param {string} input Usually email
 * 
 * @return {string} Hash of input
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 15, 2024
 * @update  Dec 15, 2024
 */

function encodeFromInput_(input) {
  return MD5(input);
}

/**
 * Create Member ID in specified row of semester sheet.
 * 
 * @param {number} row  Row to encode. Defaults to last row.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update Feb 5, 2025
 */

function encodeRowSemester_(row = getLastSubmissionInSemester()) {
  const sheet = SEMESTER_SHEET;

  const email = sheet.getRange(row, SEMESTER_COLS.EMAIL).getValue();
  const memberID = MD5(email);
  sheet.getRange(row, SEMESTER_COLS.MEMBER_ID).setValue(memberID);
}


/**
 * Create Member ID for every member in given sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Dec 18, 2024
 */

function encodeList_(sheet) {
  const sheetName = sheet.getSheetName();
  let sheetCols = GET_COL_MAP_(sheetName);

  // Start at row 2 (1-indexed)
  for (var i = 2; i <= sheet.getMaxRows(); i++) {
    var email = sheet.getRange(i, sheetCols.emailCol).getValue();
    if (!email) return;   // check for invalid row

    var member_id = MD5(email);
    sheet.getRange(i, sheetCols.memberIdCol).setValue(member_id);
  }
}


/**
 * Create single Member ID using specified row number and sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to target
 * @param {integer} [row=sheet.getLastRow()]    The 1-indexed row in input `sheet`. 
 *                                              Defaults to the last row in the sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Dec 18, 2024
 */

function encodeByRow_(sheet, row = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  let sheetCols = GET_COL_MAP_(sheetName);

  const email = sheet.getRange(row, sheetCols.emailCol).getValue();
  if (!email) throw RangeError("Invalid index access");   // check for invalid index

  const member_id = MD5(email);
  sheet.getRange(row, sheetCols.memberIdCol).setValue(member_id);
}