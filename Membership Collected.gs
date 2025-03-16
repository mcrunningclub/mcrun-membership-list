/**
 * Run formatting functions after new member submits a registration form.
 * 
 * Prevents sorting in main if concurrent runs.
 * 
 * https://developers.google.com/apps-script/samples/automations/event-session-signup
 * 
 * https://stackoverflow.com/questions/62246016/how-to-check-if-current-form-submission-is-editing-response
 *
 */

function onFormSubmit(newRow = getLastSubmissionInMain()) {
  trimWhitespace_(newRow);
  fixLetterCaseInRow_(newRow);
  encodeLastRow_(newRow);   // create unique member ID
  addMissingItems_(newRow);

  // Get payment information from Interac or Zeffy email
  checkAndSetPaymentRef_(newRow);

  // Wrap around try-catch since GAS does not support async calls
  try {
    const memberInfo =  packageMemberInfoInRow_(newRow);
    console.log(memberInfo);
    NewMemberCommunications.createNewMemberCommunications(memberInfo);
  }
  catch (e) {
    console.log(`Could not transfer new registration to external sheet 'New Member Comms'`);
    throw e;
  }
  finally {
    // Must add and sort AFTER extracting Interac info from email
    setWaiverUrl(newRow);
    addLastSubmissionToMaster(newRow);
    tryAndSortMain();   // Can only sort and format view if lock not acquired (to prevent concurrent runs)
    SpreadsheetApp.flush();   // Applies all pending changes
  }
}


/**
 * Find row index of last submission, starting from bottom using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sept 1, 2024
 * @update  Dec 18, 2024
 */

function getLastSubmissionInMain() {
  const sheet = MAIN_SHEET;
  let lastRow = sheet.getLastRow();

  while (sheet.getRange(lastRow, REGISTRATION_DATE_COL).getValue() == "") {
    lastRow = lastRow - 1;
  }

  return lastRow;
}


/**
 * Searches for member entry by email in `sheet` by binary search.
 * If unsuccessful, searches again via top-to-bottom iteration.
 * 
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} emailToFind  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 16, 2024
 * @update  Dec 18, 2024
 * 
 * @example `const submissionRowNumber = findMemberByEmail('example@mail.com', MAIN_SHEET);`
 */

function findMemberByEmail(emailToFind, sheet) {
  // First try with binary search (faster)
  const resultBinarySearch = findMemberByBinarySearch(emailToFind, sheet);

  if (resultBinarySearch != null) return resultBinarySearch;   // success!

  // If binary search unsuccessful, try with iteration (2x slower)
  return findMemberByIteration(emailToFind, sheet);
}


/**
 * Searches for member entry by email in `sheet` using iteration.
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * See faster binary search function `findMemberByBinarySearch()`.
 * 
 * @param {string} emailToFind  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * @param {number} [start=2]  The starting row index for the search (1-indexed).
 *                            Defaults to 2 (the second row) to avoid the header row.
 * @param {number} [end=MASTER_SHEET.getLastRow()]  The ending row index for the search.
 *                                                  Defaults to the last row in the sheet.
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Dec 16, 2024
 * @update  Dec 18, 2024
 * 
 * @example `const submissionRowNumber = findMemberByIteration('example@mail.com', MAIN_SHEET);`
 */

function findMemberByIteration(emailToFind, sheet, start = 2, end = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  const thisEmailCol = GET_COL_MAP_(sheetName).emailCol;    // Get email col index of `sheet`

  for (var row = start; row <= end; row++) {
    let email = sheet.getRange(row, thisEmailCol).getValue();
    if (email === emailToFind) return row;    // Exit loop and return value;
  }

  return null;
}


/**
 * Recursive function to search for entry by email in `sheet` using binary search.
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * Previously `findSubmissionFromEmail` in `Master Scripts.gs`.
 * 
 * @param {string} emailToFind  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * @param {number} [start=2]  The starting row index for the search (1-indexed). 
 *                            Defaults to 2 (the second row) to avoid the header row.
 * @param {number} [end=MASTER_SHEET.getLastRow()]  The ending row index for the search. 
 *                                                  Defaults to the last row in the sheet.
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Oct 21, 2024
 * @update  Dec 17, 2024
 * 
 * @example `const submissionRowNumber = findMemberByBinarySearch('example@mail.com', MASTER_SHEET);`
 */

function findMemberByBinarySearch(emailToFind, sheet, start = 2, end = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  const emailCol = GET_COL_MAP_(sheetName).emailCol;  // Get email col from `sheet`

  // Base case: If start index exceeds the end index, the email is not found
  if (start > end) {
    return null;
  }

  // Find the middle point between the start and end indexes
  const mid = Math.floor((start + end) / 2);

  // Get the email value at the middle row
  const emailAtMid = sheet.getRange(mid, emailCol).getValue();

  // Compare the target email with the middle email
  if (emailAtMid === emailToFind) {
    return mid;  // If the email matches, return the row index (1-indexed)

    // If the email at the middle row is alphabetically smaller, search the right half.
    // Note: use localeString() to ensure string comparison matches GSheet.
  } else if (emailAtMid.localeCompare(emailToFind) === -1) {
    return findMemberByBinarySearch(emailToFind, sheet, mid + 1, end);

    // If the email at the middle row is alphabetically larger, search the left half.
  } else {
    return findMemberByBinarySearch(emailToFind, sheet, start, mid - 1);
  }

}


/**
 * Hash function using modified MD5 algorithm.
 * 
 * Used for members' External ID.
 * 
 * @param {string} input  The string to hash.
 * @return {string}  Returns MD5-hashed input.
 *  
 * @author https://stackoverflow.com/questions/7994410/hash-of-a-cell-text-in-google-spreadsheet
 * @date  Oct 8, 2023
 */

function MD5(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  var txtHash = '';
  for (i = 0; (i < rawHash.length); i++) {
    if (i % 2 == 0) continue;
    var hashVal = rawHash[i];

    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}


/**
 * Find and set waiver url to new member registration.
 * 
 * Waiver is automatically saved by Fillout to a specific folder.
 * 
 * @param {number} [row=getLastSubmissionInMain()]  Row index to find and set url.
 *                                                  Defaults to the last row in main sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 1, 2025
 * @update  Mar 15, 2025
 */

function setWaiverUrl(row = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;

  // Search for waiver link using member name
  const memberName = sheet.getSheetValues(row, FIRST_NAME_COL, 1, 2)[0].join(' ');
  const waiverUrl = findWaiverLink_(memberName);

  // Set value of waiver url
  sheet.getRange(row, WAIVER_COL).setValue(waiverUrl);
}


/**
 * Find waiver using member's name. Helper function for setWaiverUrl
 * 
 * Waiver is automatically saved by Fillout to folder with id `WAIVER_DRIVE_ID`.
 * 
 * @param {string} name  The name of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 1, 2025
 * @update  Mar 15, 2025
 */

function findWaiverLink_(name) {
  const folderId = WAIVER_DRIVE_ID;
  const waiverFolder = DriveApp.getFolderById(folderId);

  console.log(`Now searching for waiver with name ${name}`);
  const files = waiverFolder.searchFiles(`title contains \"${name}\"`);

  const results = [];

  while (files.hasNext()) {
    const file = files.next();
    console.log(file.getName());
    results.push(file.getUrl());
  }

  return results.join('\n');
}


/**
 * Get expiration date of member fee using `semCode` and `MEMBERSHIP_DURATION`.
 * 
 * @param {string} semCode  The 3-char code representing the semester and year.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 28, 2025
 * @update  Feb 28, 2025
 */


function getExpirationDate(semCode) {
  const validDuration = MEMBERSHIP_DURATION;

  const semester = semCode.charAt(0);
  const expirationYear = '20' + (parseInt(semCode.slice(-2)) + validDuration)

  switch (semester) {
    case ('F'): return `Sep ${expirationYear}`;
    case ('W'): return `Jan ${expirationYear}`;
    case ('S'): return `Jun ${expirationYear}`;
    default: return null;
  };

}

