/**
 * Trims whitespace from specific columns in the last row of `MAIN_SHEET`.
 * 
 * This function targets the range from `FIRST_NAME_COL` to `REFERRAL_COL` (7 columns).
 * It ensures that unnecessary whitespace is removed from the latest member entry.
 * 
 * @trigger New form submission
 * 
 * @param {number} [lastRow=MAIN_SHEET.getLastRow()] - The row number to target for trimming.
 *                                                     Defaults to the last row in `MAIN_SHEET`.
 * 
  * 
 * @see {@link fixLetterCaseInRow_} for additional formatting applied to the same row.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 17, 2023
 * @update  Feb 5, 2025
 */

function trimWhitespace_(lastRow = MAIN_SHEET.getLastRow()) {
  const sheet = MAIN_SHEET;
  const rangeToFormat = sheet.getRange(lastRow, FIRST_NAME_COL, 1, 7);
  rangeToFormat.trimWhitespace();
}


/**
 * Removes diacritics (accents) from a string.
 * 
 * This function normalizes the input string and removes any diacritical marks,
 * ensuring a clean, ASCII-compatible output.
 * 
 * @param {string} str  The string to normalize and strip of diacritics.
 * @return {string}  The normalized string without diacritics.
 * 
 * @example
 * const result = removeDiacritics("José");
 * console.log(result); // Outputs: "Jose"
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 5, 2025
 * @update  Mar 15, 2025
 */

function removeDiacritics(str) {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}


///  👉 FUNCTIONS APPLIED TO MAIN_SHEET 👈  \\\

/**
 * Sorts `MAIN_SHEET` by first name, then last name.
 * 
 * This function organizes the data in `MAIN_SHEET` by sorting rows alphabetically
 * based on the `First Name` column (column 3) and then the `Last Name` column (column 4).
 * 
 * @trigger New form submission or McRUN menu.
 * 
 * @see {@link tryAndSortMain} for a safe way to sort with locking.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Jan 11, 2025
 */

function sortMainByName() {
  const sheet = MAIN_SHEET;

  const numRows = sheet.getLastRow() - 1;   // Remove header row from count
  const numCols = sheet.getLastColumn();

  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by `First Name` then by `Last Name`
  range.sort([{ column: 3, ascending: true }, { column: 4, ascending: true }]);
}


/**
 * Sorts `MAIN_SHEET` only if the lock is free.
 * 
 * This function prevents concurrent processes from interfering with sorting
 * by acquiring a script lock before proceeding. If the lock is unavailable,
 * it logs a message and exits gracefully.
 *  
 * @see {@link sortMainByName} for the actual sorting logic.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Mar 15, 2025
 * @update  Mar 15, 2025
 */

function tryAndSortMain() {
  const lock = LockService.getScriptLock();

  // Try getting lock for up to 10 seconds
  if (lock.tryLock(10000)) {
    try {
      sortMainByName();
      formatMainView();
    } finally {
      lock.releaseLock();
    }
  } else {
    console.log("Another script is running. Unable to sort now");
  }
}


/**
 * Formats `MAIN_SHEET` for a simple and uniform user experience.
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
 * @see {@link sortMainByName} for sorting logic applied to the same sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 5, 2025
 */

function formatMainView() {
  const sheet = MAIN_SHEET;

  // Helper function to improve readability
  const getThisRange = (ranges) =>
    Array.isArray(ranges) ? sheet.getRangeList(ranges) : sheet.getRange(ranges);

  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 2. Bold formatting
  getThisRange([
    'A1:T1',  // Header Row
    'A2:A',   // Registration
    'E2:E',   // Preferred Name
    'K2:L',   // Payment Method + Interac Ref Number
    'O2:P',   // Collection Date + Collector
    'T2:T',   // Member ID
  ]).setFontWeight('bold');

  // 3. Font size adjustments
  getThisRange('A1:T1').setFontSize(11);  // Header row to size 11

  getThisRange([
    'E1',   // Preferred Name (Header Cell)
    'T2:T', // Member ID
    'N1',   // Fee Paid (Header Cell)
    'S1',   // Attendance Status (Header Cell)
  ]).setFontSize(10);

  getThisRange(['Q1', 'T2:T']).setFontSize(9);  // Given to Internal (Header Cell) + Member ID
  getThisRange('K1:L1').setFontSize(8);  // Payment Method headers

  // 4. Font family adjustment
  getThisRange('T2:T').setFontFamily('Google Sans Mono');

  // 5. Format collection date
  getThisRange('A2:A').setNumberFormat('yyyy-MM-dd hh:mm:ss');
  getThisRange('O2:O').setNumberFormat('mmm d, yyyy');

  // 6. Text wrapping set to 'Clip' (for Referral + Waiver + Payment Choice)
  getThisRange(['I2:I', 'J1:J', 'K2:K']).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // 7. Horizontal and vertical alignment
  getThisRange([
    'L2:L',   // Interac Ref
    'N2:Q',   // Fee Paid, ..., Given to Internal
    'S1:T',   // Attendance Status + Member ID
  ]).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  getThisRange(['A2:A', 'I1']).setHorizontalAlignment('right');   // Align right

  // 8. Column width mapping
  const sizeMap = {
    [REGISTRATION_DATE_COL]: 140,
    [EMAIL_COL]: 245,
    [FIRST_NAME_COL]: 115,
    [LAST_NAME_COL]: 115,
    [PREFERRED_NAME_COL]: 120,
    [YEAR_COL]: 90,
    [PROGRAM_COL]: 240,
    [DESCRIPTION_COL]: 400,
    [REFERRAL_COL]: 145,
    [WAIVER_COL]: 185,
    [PAYMENT_METHOD_COL]: 155,
    [INTERAC_REF_COL]: 155,
    [EMPTY_COL]: 40,
    [IS_FEE_PAID_COL]: 75,
    [COLLECTION_DATE_COL]: 150,
    [COLLECTION_PERSON_COL]: 160,
    [IS_INTERNAL_COLLECTED_COL]: 65,
    [COMMENTS_COL]: 255,
    [ATTENDANCE_STATUS_COL]: 125,
    [MEMBER_ID_COL]: 140,
  };

  // Resize columns based on `sizeMap`
  Object.entries(sizeMap).forEach(([col, width]) => {
    sheet.setColumnWidth(col, width);
  });
}


/**
 * Adds checkboxes to specific columns in the last row of `MAIN_SHEET`.
 * 
 * This function is used to ensure that the last row of `MAIN_SHEET` has checkboxes
 * in the `Fee Paid`, `Given to Internal`, and `Attendance Status` columns.
 * 
 * @param {number} row  Row number to target for formatting.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 5, 2025
 */
function addMissingItems_(row) {
  const sheet = MAIN_SHEET;

  // Add checkboxes to target columns
  [IS_FEE_PAID_COL,
    IS_INTERNAL_COLLECTED_COL,
    ATTENDANCE_STATUS_COL
  ].forEach(col => sheet.getRange(row, col).insertCheckboxes());

  // Copy the list item  in 'Collection Person' col from first entry
  //var collectorItem = sheet.getRange(5, COLLECTION_PERSON_COL).getDataValidation();
  //var targetCell = sheet.getRange(row, COLLECTION_PERSON_COL);

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

function fixLetterCaseInRow_(row = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;

  // Set to lower case
  const rangeToLowerCase = sheet.getRange(row, EMAIL_COL);
  const rawValue = rangeToLowerCase.getValue().toString();
  rangeToLowerCase.setValue(rawValue.toLowerCase());

  // Set to Capitalized (first letter of word is UPPER)
  const rangeToCapitalize = sheet.getRange(row, FIRST_NAME_COL, 1, 5);

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
 * Sorts `MASTER` by email instead of first name.
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
 * Formats `MASTER_SHEET` for simple and uniform UX.
 * 
 * Remove whitespace from `McGill Email Address` to  `Referral`
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 22, 2024
 * @update  Dec 15, 2024
 */

function formatMasterView() {
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
 * 
 * Returns email's row index in GSheet (1-indexed), or null if not found.
 * 
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()]  The starting row index for the search (1-indexed). 
 *                                                  Defaults to 1 (the first row).
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Nov 22, 2024
 * @update  Dec 11, 2024
 */

function formatFeeCollection(row = MASTER_SHEET.getLastRow()) {
  const sheet = MASTER_SHEET;

  // STEP 1: Check for current fee status to flag for later
  const rangeFeeStatus = sheet.getRange(row, MASTER_FEE_STATUS);
  const feeStatus = rangeFeeStatus.getValue().toString();

  const regex = new RegExp('unpaid', "i"); // Case insensitive
  const isUnpaid = regex.test(feeStatus);

  // STEP 2: Insert fee status formula in `Fee Paid` col
  rangeFeeStatus.setFormula(isFeePaidFormula);    // Formula found in `Semester Variables.gs`

  // If feeStatus is unpaid, formatting is completed.
  if (isUnpaid) return;

  // STEP 3: Format collection date correctly;
  const rangeCollectionDate = sheet.getRange(row, MASTER_COLLECTION_DATE);
  const collectionDate = rangeCollectionDate.getValue();   // Format is yyyy-mm-dd hh:mm

  const formattedDate = Utilities.formatDate(collectionDate, TIMEZONE, 'yyyy-MM-dd');
  rangeCollectionDate.setValue(formattedDate);

  // STEP 4: Append semester code if collection date non-empty
  if (!collectionDate) return;

  const rangePaymentHistory = sheet.getRange(row, MASTER_PAYMENT_HIST);
  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangePaymentHistory.setValue(semCode);
}

/**
 * Inserts the 3-char semester code for the registration.
 * 
 * @param {number} [row=MASTER_SHEET.getLastRow()] 
 *                    The row number to target for inserting the semester code.
 *                    Defaults to the last row in `MASTER_SHEET`.
 * 
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 15, 2024
 * @update  Dec 15, 2024
 */

function insertRegistrationSem(row = MASTER_SHEET.getLastRow()) {
  var sheet = MASTER_SHEET;
  const rangeLatestRegSem = sheet.getRange(row, MASTER_LAST_REG_SEM);

  const semCode = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangeLatestRegSem.setValue(semCode);
}


///  👉 FUNCTIONS FOR MEMBER ID ENCODING 👈  \\\

/**
 * Create Member ID from input.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 15, 2024
 * @update  Dec 15, 2024
 */

function encodeFromInput(input) {
  return MD5(input);
}

/**
 * Create Member ID in last row of `MAIN_SHEET`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 9, 2023
 * @update Feb 5, 2025
 */

function encodeLastRow_(newSubmissionRow = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;

  const email = sheet.getRange(newSubmissionRow, EMAIL_COL).getValue();
  const memberID = MD5(email);
  sheet.getRange(newSubmissionRow, MEMBER_ID_COL).setValue(memberID);
}


/**
 * Create Member ID for every member in `sheet`.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to encode
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Dec 18, 2024
 */

function encodeList(sheet) {
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
 * Create single Member ID using row number from `sheet`.
 * 
 * @param {SpreadsheetApp.Sheet} sheet  Sheet reference to target
 * @param {integer} [row=sheet.getLastRow()]    The 1-indexed row in input `sheet`. 
 *                                              Defaults to the last row in the sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 20, 2024
 * @update  Dec 18, 2024
 */

function encodeByRow(sheet, row = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  let sheetCols = GET_COL_MAP_(sheetName);

  const email = sheet.getRange(row, sheetCols.emailCol).getValue();
  if (!email) throw RangeError("Invalid index access");   // check for invalid index

  const member_id = MD5(email);
  sheet.getRange(row, sheetCols.memberIdCol).setValue(member_id);
}