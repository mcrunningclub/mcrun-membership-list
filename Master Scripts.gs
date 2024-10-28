/* SHEET CONSTANTS */
const MASTER_NAME = 'MASTER';
const MASTER_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_NAME);

const SEMESTER_CODE_MAP = new Map();
const ALL_SEMESTERS = ['Fall 2024', 'Summer 2024', 'Winter 2024'];
const COLUMN_SIZE = 20;   // Range 'A:T' in 'MASTER'

// Index of processed semester data arrays (0-indexed)
const PROCESSED_ARR = {
  LAST_REGISTRATION: 0,
  EMAIL: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  PREFERRED_NAME: 4,
  YEAR: 5,
  PROGRAM: 6,
  MEMBER_DESCR: 7,
  REFERRAL: 8,
  WAIVER: 9,
  EMPTY: 12,
  FEE_PAID_HIST: 13,
  COLLECTION_DATE: 14,
  COLLECTED_BY: 15,
  GIVEN_TO_INTERNAL: 16,
  COMMENTS: 17,
  ATTENDANCE_STATUS: 18,
  MEMBER_ID: 19,
  LAST_REG_CODE: 20,
  REGISTRATION_HIST: 21
};

/**
 * Creates the master sheet by consolidating member data from selected semester sheets.
 * Calls `consolidateMemberData` to fetch, process, and output data to the `MASTER` sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct , 2024
 *
 */

function createMaster() {
  consolidateMemberData();
}

function addLastSubmissionToMaster() {
  consolidateLastSubmission();
  sortMasterByEmail(); // Sort 'MASTER' by email once member entry added
}

/**
 * Sorts `MASTER` by email instead of first name.
 * Required to ensure `findSubmissionByEmail` works properly.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 27, 2024
 *
 */

function sortMasterByEmail() {
  const sheet = MASTER_SHEET;
  const numRows = sheet.getLastRow() - 1;
  const numCols = sheet.getLastColumn();
    
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);
    
  // Sorts values by email
  range.sort([{column: 1, ascending: true}]);
  return;
}


/**
 * Processes the last submitted row from the `MAIN_SHEET`, adding semester codes
 * to relevant fields like `MEMBER_DESCR`, `REFERRAL`, `COMMENTS`, and payment history.
 *
 * @return {string[]} Array of processed values for the last submission.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct , 2024
 * 
 * ```javascript
 * // Sample Script ➜ Storing processed submission.
 * const processedData = processLastSubmission();
 * ```
 */

function processLastSubmission() {
  const lastRowNum = getLastSubmission();     // Last row num from 'MAIN_SHEET'
  const semesterCode = getSemesterCode(SHEET_NAME); // Get the semester code based on the sheet name
  var lastSubmission = MAIN_SHEET.getRange(lastRowNum, 1, 1, COLUMN_SIZE).getValues()[0];
  
  const indicesToProcess = [PROCESSED_ARR.MEMBER_DESCR, PROCESSED_ARR.REFERRAL, PROCESSED_ARR.COMMENTS];

  // Loop over the relevant indices and append semester code to non-empty fields
  indicesToProcess.forEach(index => {
    if (lastSubmission[index]) {
      lastSubmission[index] = `(${semesterCode}) ${lastSubmission[index]}`;
    }
  });

  // Append semester code to IS_FEE_PAID column
  if (lastSubmission[PROCESSED_ARR.FEE_PAID_HIST]) {
    lastSubmission[PROCESSED_ARR.FEE_PAID_HIST] = semesterCode;
  }

  // Add semester code for MASTER.LAST_REG_CODE and MASTER.REGISTRATION_HIST
  lastSubmission.push(semesterCode);  // For MASTER.LAST_REG_CODE column
  lastSubmission.push("");            // For MASTER.REGISTRATION_HIST column


  return lastSubmission;
}

/**
 * Consolidates the last submitted row from `MAIN` into `MASTER`.
 * Checks if an existing entry with the same email exists in the MASTER sheet:
 *   - If found, updates specific fields with concatenated data from both entries.
 *   - If not found, appends the new data as a fresh row.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 21, 2024
 */

function consolidateLastSubmission() {
  var sheet = MASTER_SHEET;
  var processedLastSubmission = processLastSubmission();

  // Search for user in 'MASTER'
  const lastEmail = processedLastSubmission[PROCESSED_ARR.EMAIL];
  const indexSubmission = findSubmissionFromEmail(lastEmail);   // Returns null if not found
  
  // Check if user already exists
  if (indexSubmission != null) {
    var existingEntry = sheet.getRange(indexSubmission, 1, 1, COLUMN_SIZE).getValues()[0];

    // Data to append in latest registration: 
    const indicesToAppend = {
      [PROCESSED_ARR.MEMBER_DESCR]: existingEntry[7],   // Describe Yourself 'H'
      [PROCESSED_ARR.REFERRAL]: existingEntry[8],   // Referral 'I'
      [PROCESSED_ARR.FEE_PAID_HIST]: existingEntry[18],   // Payment History 'S'
      [PROCESSED_ARR.COMMENTS]: existingEntry[19]   // Comments 'T'
    };

    // Only append if existingEntry[i] non-empty
    for (const index in indicesToAppend) {
      var existingValue = indicesToAppend[index];
      if (existingValue) {
        var delimiter = processedLastSubmission[index] ? "\n" : "";
        processedLastSubmission[index] += delimiter + existingValue;
      }
    }

    // Data to keep if latest registration empty: 
    const indicesToKeep = {
      [PROCESSED_ARR.COLLECTED_BY]: existingEntry[16],     // Collected By 'P'
      [PROCESSED_ARR.COLLECTION_DATE]: existingEntry[17],   // Collection Date 'Q'
      [PROCESSED_ARR.GIVEN_TO_INTERNAL]: existingEntry[18],  // Given to Internal 'R'
    };

    for (const index in indicesToKeep) {
      if (processedLastSubmission[index] == "") {
        var existingValue = indicesToKeep[index];
        processedLastSubmission[index] = existingValue;
      }
    }

    // Case 1: no previous registration -> move old regCode to regHistory
    // Lastly add registration history
    var existingHistory = existingEntry[11];
    var existingRegCode = existingEntry[10];
    var latestHistory = processedLastSubmission[PROCESSED_ARR.REGISTRATION_HIST];

    // Registration History 'L'
    var delimiter = latestHistory ? "\n" : "";
    processedLastSubmission[PROCESSED_ARR.REGISTRATION_HIST] = existingRegCode + delimiter + existingHistory;    
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
    PROCESSED_ARR.MEMBER_DESCR,
    PROCESSED_ARR.REFERRAL,
    PROCESSED_ARR.LAST_REGISTRATION,
    PROCESSED_ARR.LAST_REG_CODE,
    PROCESSED_ARR.REGISTRATION_HIST,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.COLLECTED_BY,
    PROCESSED_ARR.COLLECTION_DATE,
    PROCESSED_ARR.GIVEN_TO_INTERNAL,
    PROCESSED_ARR.FEE_PAID_HIST,
    PROCESSED_ARR.COMMENTS,
    PROCESSED_ARR.EMPTY,
    PROCESSED_ARR.MEMBER_ID
  ];

  // Store selected data in new array
  var selectedData = indicesToSelect.map(index => processedLastSubmission[index] || "");

  // Output data to 'MASTER'
  // CASE 1 : User exists -> replace previous entry
  if (indexSubmission != null) {
    sheet.getRange(indexSubmission, 1, 1, selectedData.length).setValues([selectedData]);
  }
  // CASE 2: User does not exist in 'MASTER' -> add new entry to the bottom of sheet
  else {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, selectedData.length).setValues([selectedData]);
  }

  return;
}



/**
 * Processes data for a given semester sheet, adding semester codes to selected
 * fields and returning the formatted data.
 * 
 * Helper function for `consolidateMemberData()`.
 *
 * @param {string} sheetName  The name of the semester sheet to process (e.g., 'Fall 2024').
 * @return {Array<Array>}  Returns an array of processed row data for the given semester.
 *
 * @example `const processedData = processSemesterData('Fall 2024');`
 */

function processSemesterData(sheetName) {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const semesterSheetRange = 'A2:T';
  
  var sheetData = SPREADSHEET.getSheetByName(sheetName).getRange(semesterSheetRange).getValues();
  const semesterCode = getSemesterCode(sheetName); // Get the semester code based on the sheet name

  const processedData = sheetData.map(function (row) {
    // Append semester code if entries are non-empty by looping over selected indices
    const indicesToProcess = [PROCESSED_ARR.MEMBER_DESCR, PROCESSED_ARR.REFERRAL, PROCESSED_ARR.COMMENTS];
  
    indicesToProcess.forEach(index => {
      if (lastSubmission[index]) {
        lastSubmission[index] = `(${semesterCode}) ${lastSubmission[index]}`;
      }
    });

    // Append semester code to payment history
    index = PROCESSED_ARR.FEE_PAID_HIST;
    row[index] = row[index] ? semesterCode : "";

    // Append row with semester code for MASTER.LAST_REG_CODE
    row.push(semesterCode);

    // Append row with semester code for MASTER.REGISTRATION_HIST
    row.push("");

    return row;
  });

  return processedData;
}

/**
 * Recursive function to search for entry by email in `MASTER` sheet using binary search.
 * Returns email's row index in GSheet (1-indexed), or null if not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 21, 2024
 * @update  Oct 23, 2024
 * 
 * @param {string} emailToFind  The email address to search for in the sheet.
 * @param {number} [start=1]  The starting row index for the search (1-indexed). 
 *                            Defaults to 1 (the first row).
 * @param {number} [end=MASTER_SHEET.getLastRow()]  The ending row index for the search. 
 *                                                  Defaults to the last row in the sheet.
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @example `const submissionRowNumber = findSubmissionFromEmail('example@mail.com');`
 */

function findSubmissionFromEmail(emailToFind, start=1, end=MASTER_SHEET.getLastRow()) {
  const MASTER_EMAIL_COL = 1;
 
  // Base case: If start index exceeds the end index, the email is not found
  if (start > end) {
    return null;
  }

  // Find the middle point between the start and end indexes
  const mid = Math.floor((start + end) / 2);

  // Get the email value at the middle row
  const emailAtMid = MASTER_SHEET.getRange(mid, MASTER_EMAIL_COL).getValue();


  // Compare the target email with the middle email
  if (emailAtMid === emailToFind) {
    return mid;  // If the email matches, return the row index (1-indexed)
  
   // If the email at the middle row is alphabetically smaller, search the right half
   // Note: use localeString() to ensure string comparison matches GSheet
  } else if (emailAtMid.localeCompare(emailToFind) === -1) {
    return findSubmissionFromEmail(emailToFind, mid + 1, end);
  
  // If the email at the middle row is alphabetically larger, search the left half
  } else {
    return findSubmissionFromEmail(emailToFind, start, mid - 1);
  }

}


/**
 * Requires user validation to consolidate member registration.
 * Prevent unwanted data overwrite in `MASTER` sheet
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 28, 2024
 * 
 * @returns {boolean}  Returns user choice as a boolean.
 *
 * @example `const userChoice = confirmMasterOverwrite();`
 */

function confirmMasterOverwrite() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Do you want to consolidate member registrations?";
  const textMsg = "This action will overwrite present data in MASTER. Ensure that data has been copied beforehand.";

  var choice = ui.alert(
    headerMsg,
    textMsg,
    ui.ButtonSet.YES_NO,
  );

  const result = (choice == ui.Button.YES);

  // Process the user's response.
  if (result) {
    // User clicked "Yes".
    ui.alert("Confirmation received. Starting data consolidation...");
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Process cancelled.");
  }

  return result;
}


function consolidateMemberData() {
  // Verify if data overwrite wanted
  if(!confirmMasterOverwrite()) {
    return;
  }

  // Prevent data overwrite for now
  Logger.log("Currently preventing consolidateMemberData() from running. Remove return statement from code to continue.");
  return;

  // Get processed semester data
  var fall2024 = processSemesterData('Fall 2024');
  var summer2024 = processSemesterData('Summer 2024');
  var winter2024 = processSemesterData('Winter 2024');

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

      index = PROCESSED_ARR.MEMBER_DESCR;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      index = PROCESSED_ARR.REFERRAL;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      index = PROCESSED_ARR.COMMENTS;
      if (row[index]) memberMap[email][index] += "\n" + row[index];

      // Append registration history using semester code
      index = PROCESSED_ARR.REGISTRATION_HIST;
      semesterCode = row[PROCESSED_ARR.LAST_REG_CODE];
      regHistory = memberMap[email][index] ? " " + semesterCode : semesterCode;
      memberMap[email][index] += regHistory;

      // Append payment history
      index = PROCESSED_ARR.FEE_PAID_HIST;
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
      row[PROCESSED_ARR.MEMBER_DESCR],
      row[PROCESSED_ARR.REFERRAL],
      row[PROCESSED_ARR.LAST_REGISTRATION],
      row[PROCESSED_ARR.LAST_REG_CODE],
      row[PROCESSED_ARR.REGISTRATION_HIST],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.COLLECTED_BY],
      row[PROCESSED_ARR.COLLECTION_DATE],
      row[PROCESSED_ARR.GIVEN_TO_INTERNAL],
      row[PROCESSED_ARR.FEE_PAID_HIST],
      row[PROCESSED_ARR.COMMENTS],
      row[PROCESSED_ARR.EMPTY],
      row[PROCESSED_ARR.MEMBER_ID]
    ]
  });

  // Output sorted unique data to another sheet or range
  MASTER_SHEET.getRange(2, 1, selectedData.length, selectedData[0].length).setValues(selectedData);
}




// Helper Function
function getSemesterCode(semester) {
  // Return semester code if already in map
  if (semester in SEMESTER_CODE_MAP) {
    return SEMESTER_CODE_MAP.get(semester)
  }

  // Extract the first letter of the string and last two char representing the year
  const semesterType = semester.charAt(0);
  const semesterYear = semester.slice(-2);
  const code = semesterType + semesterYear;

  // Same key-value in map
  SEMESTER_CODE_MAP.set(semester, code);

  return code;
}



function sortUniqueData() {
  // Get the active sheet and data from each range
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var fall2024 = sheet.getSheetByName('Fall 2024').getRange('A2:U').getValues();
  var summer2024 = sheet.getSheetByName('Summer 2024').getRange('A2:U').getValues();
  var winter2024 = sheet.getSheetByName('Winter 2024').getRange('A2:U').getValues();

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
    var nameA = a[1].normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove accents from nameA
    var nameB = b[1].normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove accents from nameB
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

/**
 * @author: Andrey S Gonzalez
 * @date: Oct 23, 2024
 * @update: Oct 23, 2024
 * 
 * Recursive function to find submission in MASTER using email string.
 * Return null if not found.
 * @PARAM rowNumber (1-indexed according to GSheet)
 * 
 */

// Create single Member ID using row number from 'MASTER' sheet
function encodeByRow(rowNumber) {
  var sheet = MASTER_SHEET;
  const MASTER_EMAIL_COL = 1;
  const MASTER_MEMBER_ID_COL = 22;

  const email = sheet.getRange(rowNumber, MASTER_EMAIL_COL).getValue();
  if (email === "") throw RangeError("Invalid index access");   // check for invalid index

  const member_id = MD5(email);
  sheet.getRange(i, MASTER_MEMBER_ID_COL).setValue(member_id);
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 20, 2024
 * @update: Oct 23, 2024
 * 
 * Encode whole list in 'MASTER' sheet using MD5 algorithm
 * 
 */

// Create Member ID using email
function encodeMasterList() {
  var sheet = MASTER_SHEET;
  const MASTER_EMAIL_COL = 1;
  const MASTER_MEMBER_ID_COL = 22;
  var i, email;

  // Start at row 2 (1-indexed)
  for (i = 221; i <= sheet.getMaxRows(); i++) {
    email = sheet.getRange(i, MASTER_EMAIL_COL).getValue();
    if (email === "") return;   // check for invalid row

    var member_id = MD5(email);
    sheet.getRange(i, MASTER_MEMBER_ID_COL).setValue(member_id);
  }
}