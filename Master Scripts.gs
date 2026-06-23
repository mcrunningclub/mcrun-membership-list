/**
 * Creates the master sheet by consolidating member data from selected semester sheets.
 * 
 * Calls `consolidateMemberData` to fetch, process, and output data to the master sheet.
 *
 * @author [Andrey Gonzalez](andrey.gonzalez@mail.mcgill.ca)
 * @date  Oct 23, 2024
 *
 */

function createMaster() {
  consolidateMemberData_();
}

/**
 * Adds the last submission from semester sheet to the master sheet.
 *
 * This function processes the last row of the semester sheet, consolidates the data,
 * and ensures the master sheet is sorted by email after the new entry is added.
 *
 * @param {number} [lastRow=getLastSubmissionInSemester()] - The row number of the last submission in the semester.
 *                                                        Defaults to the last row.
 *
 *
 * @see {@link consolidateLastSubmission_} for the logic of consolidating the last submission.
 * @see {@link sortMasterByEmail} for sorting the `MASTER` sheet by email.
 *
 * @author Andrey Gonzalez
 * @date Oct 23, 2024
 */
function addLastSubmissionToMaster_(lastRow = getLastSubmissionInSemester()) {
  consolidateLastSubmission_(lastRow);
  sortMasterByEmail(); // Sort 'MASTER' by email once member entry added
}


/**
 * Updates Payment History in master sheet from the member's semester sheet where they registered.
 *
 * Appends the semester code to the payment history column if it is not already present.
 *
 * @param {number} memberRow  The 1-indexed row number of the member in master sheet.
 * @param {string} semesterSheetName  The name of the member's latest registration semester sheet.
 * 
 * @see {@link getSemesterCode_} for how the semester code is determined.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 17, 2024
 * @update  Dec 17, 2024
 */

function addPaidSemesterToMaster_(memberRow, semesterSheetName) {
  const sheet = MASTER_SHEET;
  const paymentHistoryCol = MASTER_COLS.PAYMENT_HISTORY;
  const semesterCode = getSemesterCode_(semesterSheetName); // Get the semester code based on the sheet name

  const rangePaymentHistory = sheet.getRange(memberRow, paymentHistoryCol);
  const paymentHistory = rangePaymentHistory.getValue();

  // If previous payment history, append with `semesterCode`
  // Otherwise only use `semesterCode`.
  const newHistory = paymentHistory ? `${paymentHistory}\n${semesterCode}` : semesterCode;

  // Only modify if paymentHistory **does not** contain semesterCode to prevent duplicates.
  if (!paymentHistory.includes(semesterCode)) rangePaymentHistory.setValue(newHistory);
}


/**
 * Updates the `isFeePaid` status in the member's semester sheet.
 *
 * This function checks if the member's semester code is present in their payment history.
 * If the code is found, it sets `isFeePaid` to `true`; otherwise, it sets it to `false`.
 *
 * @param {string} payHistory  The payment history of the member, stored as newline-separated semester codes.
 * @param {number} memberRow  The 1-indexed row number of the member in the semester sheet.
 * @param {number} isFeePaidCol  The 1-indexed column number of the `isFeePaid` field in the semester sheet.
 * @param {SpreadsheetApp.Sheet} semesterSheet  The member's latest registration sheet (e.g., "Fall 2024").
 *
 * @see {@link getSemesterCode_} for how the semester code is determined.
 *
 * @author Andrey Gonzalez
 * @date Dec 17, 2024
 */

function updateFeeStatusSemester_(payHistory, memberRow, isFeePaidCol, semesterSheet) {
  const paymentHistoryArray = payHistory.split('\n');
  const semesterCode = getSemesterCode_(semesterSheet.getSheetName()); // Get the semester code based on the sheet name

  // Returns false if no payment history or semester code not in payHistory
  const isFeePaid = paymentHistoryArray.includes(semesterCode);
  Logger.log(`updateIsFeePaid -> payHistory: ${payHistory} isFeePaid: ${isFeePaid}`);

  const rangeIsFeePaid = semesterSheet.getRange(memberRow, isFeePaidCol);
  rangeIsFeePaid.setValue(isFeePaid);
}


/**
 * Processes the last submitted row from the semester, adding semester codes
 * to relevant fields like `MEMBER_DESCR`, `REFERRAL`, `COMMENTS`, and payment history.
 * 
 * @param {number} [lastRow=getLastSubmissionInSemester()] - The row number of the last submission in the semester.
 *                                                        Defaults to the last row.
 *
 * @return {string[]} Array of processed values for the last submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 21, 2024
 * @update Mar 15, 2025
 */

function processLastSubmission_(lastRow = getLastSubmissionInSemester()) {
  const semesterCode = getSemesterCode_(SHEET_NAME); // Get the semester code based on the sheet name
  var lastSubmission = SEMESTER_SHEET.getSheetValues(lastRow, 1, 1, Object.entries(SEMESTER_COLS).length)[0];

  const indicesToProcess = [PROCESSED_ARR.DESCRIPTION, PROCESSED_ARR.REFERRAL, PROCESSED_ARR.COMMENTS];

  // Loop over the relevant indices and prepend semester code to non-empty fields
  indicesToProcess.forEach(index => {
    if (lastSubmission[index]) {
      lastSubmission[index] = `(${semesterCode}) ${lastSubmission[index]}`;
    }
  });

  // Append semester code to IS_FEE_PAID column in array
  if (lastSubmission[PROCESSED_ARR.FEE_PAID_SEM]) {
    lastSubmission[PROCESSED_ARR.FEE_PAID_SEM] = semesterCode;
  }

  // Add semester code for MASTER.LATEST_REG_SEMESTER and MASTER.REG_HISTORY
  lastSubmission.push(semesterCode);  // For MASTER.LATEST_REG_SEMESTER column
  lastSubmission.push("");            // For MASTER.REG_HISTORY column

  return lastSubmission;
}


/**
 * Consolidates the last submitted row from semester into master sheet.
 * 
 * Checks if an existing entry with the same email exists in the MASTER sheet:
 *   - If found, updates specific fields with concatenated data from both entries.
 *   - If not found, appends the new data as a fresh row.
 * 
 * @param {number} [lastRow=getLastSubmissionInSemester()] - The row number of the last submission in the semester.
 *                                                        Defaults to the last row.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 21, 2024
 * @update Mar 15, 2025
 */
// IMPROVE RUNTIME!
function consolidateLastSubmission_(lastRow = getLastSubmissionInSemester()) {
  var sheet = MASTER_SHEET;
  var lastSubmission = processLastSubmission_(lastRow);

  // Search for user in 'MASTER'
  const lastEmail = lastSubmission[PROCESSED_ARR.EMAIL];
  const memberRow = findMemberByEmail(lastEmail, sheet);   // Returns null if not found

  // Check if user already exists
  if (memberRow != null) {
    var existingEntry = sheet.getSheetValues(memberRow, 1, 1, Object.entries(MASTER_COLS).length)[0];

    // Data to append in latest registration: 
    const columnsToMerge = {
      [PROCESSED_ARR.DESCRIPTION]: MASTER_COLS.DESCRIPTION - 1,    // Describe Yourself 'H'
      [PROCESSED_ARR.REFERRAL]: MASTER_COLS.REFERRAL - 1,        // Referral 'I'
      [PROCESSED_ARR.FEE_PAID_SEM]: MASTER_COLS.PAYMENT_HISTORY - 1,  // Payment History 'S'
      [PROCESSED_ARR.COMMENTS]: MASTER_COLS.COMMENTS - 1        // Comments 'T'
    };

    // Only append if existingEntry[i] non-empty
    for (const column in columnsToMerge) {
      var existingValue = existingEntry[columnsToMerge[column]];
      var newValue = lastSubmission[column];

      if (existingValue && newValue) {
        lastSubmission[column] += "\n" + existingValue;
      }
    }

    // Data to keep if latest registration empty:
    const columnsToKeep = {
      [PROCESSED_ARR.COLLECTED_BY]: MASTER_COLS.COLLECTED_BY - 1,     // Collected By 'P'
      [PROCESSED_ARR.COLLECTION_DATE]: MASTER_COLS.COLLECTION_DATE - 1,   // Collection Date 'Q'
      [PROCESSED_ARR.IS_INTERNAL_COLLECTED]: MASTER_COLS.IS_INTERNAL_COLLECTED - 1,  // Given to Internal 'R'
    };

    for (const column in columnsToKeep) {
      var newValue = lastSubmission[column];
      if (newValue == "") {
        var existingValue = existingEntry[columnsToKeep[column]];
        lastSubmission[column] = existingValue;
      }
    }

    // If no previous registration -> move existing registration semester to history
    // If yes -> add existing code to existing history
    // And lastly add registration history
    var existingHistory = existingEntry[MASTER_COLS.REG_HISTORY - 1];
    var existingRegCode = existingEntry[MASTER_COLS.LATEST_REG_SEMESTER - 1];
    if (existingHistory == "") {
      lastSubmission[PROCESSED_ARR.REG_HISTORY] = existingRegCode;
    }
    else {
      lastSubmission[PROCESSED_ARR.REG_HISTORY] = existingRegCode + "\n" + existingHistory;
    }
    
  }

  // Select specific columns
  const indicesToSelect = [
    PROCESSED_ARR.EMAIL,
    PROCESSED_ARR.FIRST_NAME,
    PROCESSED_ARR.LAST_NAME,
    PROCESSED_ARR.PREFERRED_NAME,
    PROCESSED_ARR.YEAR,
    PROCESSED_ARR.PROGRAM,
    PROCESSED_ARR.WAIVER,
    PROCESSED_ARR.DESCRIPTION,
    PROCESSED_ARR.REFERRAL,
    PROCESSED_ARR.LATEST_REG_DATE,
    PROCESSED_ARR.LATEST_REG_SEM,
    PROCESSED_ARR.REG_HISTORY,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.COLLECTED_BY,
    PROCESSED_ARR.COLLECTION_DATE,
    PROCESSED_ARR.IS_INTERNAL_COLLECTED,
    PROCESSED_ARR.FEE_PAID_SEM,
    PROCESSED_ARR.COMMENTS,
    PROCESSED_ARR.ATTENDANCE_STATUS,
    PROCESSED_ARR.MEMBER_ID
  ];

  // Store selected data in new array
  var selectedData = indicesToSelect.map(index => lastSubmission[index] || "");
  var newEntryRow = sheet.getLastRow() + 1;

  // Output data to 'MASTER'
  // CASE 1 : User exists -> replace previous entry
  if (memberRow != null) {
    sheet.getRange(memberRow, 1, 1, selectedData.length).setValues([selectedData]);
  }
  // CASE 2: User does not exist in 'MASTER' -> add new entry to the bottom of sheet
  else {
    sheet.getRange(newEntryRow, 1, 1, selectedData.length).setValues([selectedData]);
  }

  // Get row of member if existing or non-existing
  const targetRow = memberRow ? memberRow : newEntryRow;

  // Add formula for `Fee Paid` col
  const isFeePaidCell = sheet.getRange(targetRow, MASTER_COLS.FEE_PAID);
  isFeePaidCell.setFormula(isFeePaidFormula);
}


/**
 * Processes data for a given semester sheet, adding semester codes to selected
 * fields and returning the formatted data.
 * 
 * Helper function for `consolidateMemberData()`.
 *
 * @param {string} sheetName  The name of the semester sheet to process (e.g., 'Fall 2024').
 * @return {string[][]}  Returns an array of processed row data for the given semester.
 *
 * @example `const processedData = processSemesterData('Fall 2024');`
 */

function processSemesterData_(sheetName) {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const semesterSheetRange = 'A2:T';

  var sheetData = SPREADSHEET.getSheetByName(sheetName).getSheetValues(semesterSheetRange);
  const semesterCode = getSemesterCode_(sheetName); // Get the semester code based on the sheet name

  const processedData = sheetData.map(function (row) {
    // Append semester code if entries are non-empty by looping over selected indices
    const indicesToProcess = [PROCESSED_ARR.DESCRIPTION, PROCESSED_ARR.REFERRAL, PROCESSED_ARR.COMMENTS];

    indicesToProcess.forEach(index => {
      if (lastSubmission[index]) {
        lastSubmission[index] = `(${semesterCode}) ${lastSubmission[index]}`;
      }
    });

    // Append semester code to payment history
    index = PROCESSED_ARR.FEE_PAID_SEM;
    row[index] = row[index] ? semesterCode : "";

    // Append row with semester code for MASTER.LATEST_REG_SEMESTER
    row.push(semesterCode);

    // Append row with semester code for MASTER.REG_HISTORY
    row.push("");

    return row;
  });

  return processedData;
}

/**
 * Combines data from 2024 semesters into new master sheet (overwrites existing)
 * 
 * Get and process semester data and concatenate, then create a map indexed by emails
 * to make sure entries are unique/combine entries with the same email. Sorts output by first name.
 */
function consolidateMemberData_() {
  // Verify if data overwrite wanted
  createMasterUI_();

  // Prevent data overwrite for now
  Logger.log("Currently preventing consolidateMemberData() from running. Remove return statement from code to continue.");
  return;

  // Get processed semester data
  var fall2024 = processSemesterData_('Fall 2024');
  var summer2024 = processSemesterData_('Summer 2024');
  var winter2024 = processSemesterData_('Winter 2024');

  // Combine all semester data
  var allMemberData = fall2024.concat(summer2024, winter2024);

  // Filter out empty rows
  allMemberData = allMemberData.filter(function (row) {
    return row[0] !== "" && row[1] !== "";
  });

  // Create an object to store the unique entries by email address
  const memberMap = {};

  allMemberData.forEach(function (row) {
    var email = row[1];

    if (!memberMap[email]) {
      // Initialize the entry for the email if it doesn't exist
      memberMap[email] = row;
    } else {
      // Concatenate the relevant columns if the email already added
      // Access memberMap is 0-indexed according to semester sheet
      var index, regHistory, semesterCode;

      index = PROCESSED_ARR.DESCRIPTION;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      index = PROCESSED_ARR.REFERRAL;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      index = PROCESSED_ARR.COMMENTS;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      // Append registration history using semester code
      index = PROCESSED_ARR.REG_HISTORY;
      semesterCode = row[PROCESSED_ARR.LATEST_REG_SEM];
      regHistory = memberMap[email][index] ? " " + semesterCode : semesterCode;
      memberMap[email][index] += regHistory;

      // Append payment history
      index = PROCESSED_ARR.FEE_PAID_SEM;
      if (row[index]) memberMap[email][index] += " " + row[index];
    }
  });

  // Convert the emailMap object back into an array
  var resultData = Object.values(memberMap);

  // Sort by 'First Name' column
  resultData.sort(function (a, b) {
    var firstNameA = a[2].toLowerCase(); // First Name in lowercase for consistent sorting
    var firstNameB = b[2].toLowerCase();
    return firstNameA.localeCompare(firstNameB);
  });


  // Select specific columns
  var selectedData = resultData.map(function (row) {
    return [
      row[PROCESSED_ARR.EMAIL],
      row[PROCESSED_ARR.FIRST_NAME],
      row[PROCESSED_ARR.LAST_NAME],
      row[PROCESSED_ARR.PREFERRED_NAME],
      row[PROCESSED_ARR.YEAR],
      row[PROCESSED_ARR.PROGRAM],
      row[PROCESSED_ARR.WAIVER],
      row[PROCESSED_ARR.DESCRIPTION],
      row[PROCESSED_ARR.REFERRAL],
      row[PROCESSED_ARR.LATEST_REG_DATE],
      row[PROCESSED_ARR.LATEST_REG_SEM],
      row[PROCESSED_ARR.REG_HISTORY],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.COLLECTED_BY],
      row[PROCESSED_ARR.COLLECTION_DATE],
      row[PROCESSED_ARR.IS_INTERNAL_COLLECTED],
      row[PROCESSED_ARR.FEE_PAID_SEM],
      row[PROCESSED_ARR.COMMENTS],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.MEMBER_ID]
    ]
  });

  // Output sorted unique data to another sheet or range
  MASTER_SHEET.getRange(2, 1, selectedData.length, selectedData[0].length).setValues(selectedData);
}

/**
 * Combine data from 2024 and sort it into a new master sheet (overwrites existing)
 * 
 * Combines sheet values with their sheet name, then removes duplicates and sorts by
 * first name.
 */
function sortUniqueData() {
  // Get the active sheet and data from each range
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var fall2024 = sheet.getSheetByName('Fall 2024').getSheetValues('A2:U');
  var summer2024 = sheet.getSheetByName('Summer 2024').getSheetValues('A2:U');
  var winter2024 = sheet.getSheetByName('Winter 2024').getSheetValues('A2:U');

  // Append the sheet name to each row
  fall2024 = fall2024.map(function (row) {
    return row.concat('Fall 2024'); // Append the sheet name 'Fall 2024'
  });

  summer2024 = summer2024.map(function (row) {
    return row.concat('Summer 2024'); // Append the sheet name 'Summer 2024'
  });

  winter2024 = winter2024.map(function (row) {
    return row.concat('Winter 2024'); // Append the sheet name 'Winter 2024'
  });


  // Combine data from all semesters
  var allData = fall2024.concat(summer2024, winter2024);

  // Filter out empty rows
  allData = allData.filter(function (row) {
    return row[0] !== "" && row[1] !== "";
  });

  // Remove duplicates
  var uniqueData = allData.filter((row, index, self) =>
    index === self.findIndex((r) => r[0] === row[0] && r[1] === row[1])
  );

  // Sort by the second column (ignoring accents)
  uniqueData.sort(function (a, b) {
    var nameA = removeDiacritics_(a[1]); // Remove accents from nameA
    var nameB = removeDiacritics_(b[1]); // Remove accents from nameB
    return nameA.localeCompare(nameB);
  });

  // Select specific columns: 4, 5, 6, 7, 10, 22, 1 (arrays are 0-indexed, so Col4 is index 3, Col1 is index 0, etc.)
  var selectedData = uniqueData.map(function (row) {
    return [row[1], row[2], row[3], row[4], row[5], row[6], row[9], row[0], row[21]];
  });

  // Output sorted unique data to another sheet or range
  MASTER_SHEET.getRange(2, 1, selectedData.length, selectedData[0].length).setValues(selectedData);

  //MASTER_SHEET.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
}
