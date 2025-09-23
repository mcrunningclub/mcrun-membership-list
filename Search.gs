// NAME OF STORE WITH INDEX OF LETTER
const INDEX_STORE_NAME = "letterIndexStore";

// Setter and getter for index store
function getIndexStore_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const store = scriptProperties.getProperty(INDEX_STORE_NAME);
  return JSON.parse(store) ?? setIndexStore();
}

function setIndexStore() {
  const store = buildLetterIndexStore_();
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(INDEX_STORE_NAME, JSON.stringify(store));
  return store;
}


/**
 * Builds an index store mapping the first letter of a key (i.e. email) to 
 * its first occurrence index in `MASTER_SHEET`. e.g { 'a': 2, 'b' : 21, ... }
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Sep 22, 2025
 * @update  Sep 22, 2025
 */

function buildLetterIndexStore_() {
  // Get all emails from `MASTER` sheet
  const startRow = 2;
  const numRows = MASTER_SHEET.getLastRow() - 1;
  const emails = MASTER_SHEET.getSheetValues(startRow, MASTER_COLS.email, numRows, 1);

  const store = {};

  for (let i = 0; i < emails.length; i++) {
    const email = (emails[i][0] || '').toString().trim();
    if (!email) continue;

    const letter = email[0].toLowerCase();
    if (store[letter] == undefined) {
      store[letter] = i + startRow;
    }
  }
  return store;
}


function testRuntime() {
  const email = 'example@mail.com';
  const startTime = new Date().getTime();

  /**
   * Runtime compared in ms: 
   * - findMemberByIteration: 1133, 1749, 1185, 1191
   * - findMemberByBinarySearch: 643, 1548, 926 
   * - findMemberWithStore: 275, 239, 223, 235
  */

  //findMemberByIteration(email, MASTER_SHEET);
  //findMemberByBinarySearch(email, MASTER_SHEET);
  findMemberWithStore_(email);
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
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
 * @update  Sep 22, 2025
 * 
 * @example `const submissionRowNumber = findMemberByEmail('example@mail.com', MAIN_SHEET);`
 */

function findMemberByEmail(emailToFind, sheet) {
  let result = null;

  // First try with letter index store and binary search (fastest)
  if (sheet === MASTER_SHEET) result = findMemberWithStore_(emailToFind);

  // Try with binary search with whole sheet if unsuccessful or sheet !== MASTER_SHEET
  result = result ?? findMemberByBinarySearch(emailToFind, sheet);

  // If binary search unsuccessful, try with iteration (slowest)
  result = result ?? findMemberByIteration(emailToFind, sheet);
  return result;
}


/**
 * Finds a member by setting a stricter bound using an index store of each letter.
 * Then searches `MASTER_SHEET` using binary search.
 * 
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} emailToFind  The email address to search for in `sheet`.
 * @param {Object} [store=getIndexStore()]  Object mapping first letter 
 *                                          to starting index, e.g. { 'a': 1, 'b': 21, ... }
 * 
 * @return {number|null}  Returns the 1-indexed row number where the email is found, 
 *                        or `null` if the email is not found.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Sep 22, 2025
 * @update  Sep 22, 2025
 * 
 * @example `const submissionRowNumber = findMemberWithStore('example@mail.com');`
 */
function findMemberWithStore_(emailToFind, store = getIndexStore_()) {
  const letter = emailToFind[0].toLowerCase();
  const BUFFER = 3;   // In case the store has not been updated

  // Set lower and upper bound for binary search
  const lowerBound = store[letter] - BUFFER;
  const upperBound = getUpperBound(letter) + BUFFER;

  // Use binary search with a smaller search range
  return findMemberByBinarySearch(emailToFind, MASTER_SHEET, lowerBound, upperBound);

  // Get upper bound using next letter found in store
  function getUpperBound(char) {
    let ascii = char.charCodeAt(0);
    while(store[String.fromCharCode(++ascii)] == undefined) {}
    return store[String.fromCharCode(ascii)];
  }
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

