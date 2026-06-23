/**
 * Retrieves index store from script properties and parses it.
 * 
 * If store is not found, call function to create it.
 * 
 * @returns {Object}  Mapping letters to the row with the first occurence 
 *                      of an email starting with that letter
 */
function getIndexStore() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const store = scriptProperties.getProperty(INDEX_STORE_NAME);
  return JSON.parse(store) ?? setIndexStore();
}

/**
 * Builds and sets index store as a script property.
 * 
 * @returns {Object}  Mapping letters to the row with the first occurence 
 *                      of an email starting with that letter
 */
function setIndexStore() {
  const store = buildIndexStore_();
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(INDEX_STORE_NAME, JSON.stringify(store));
  return store;
}


/**
 * Builds an index store mapping the first letter of a key (i.e. email) to 
 * its first occurrence index in the master sheet. e.g { 'a': 2, 'b' : 21, ... }
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  Sep 22, 2025
 * @update  Sep 22, 2025
 */

function buildIndexStore_() {
  // Get all emails from `MASTER` sheet
  const startRow = 2;
  const numRows = MASTER_SHEET.getLastRow() - 1;
  const emails = MASTER_SHEET.getSheetValues(startRow, MASTER_COLS.EMAIL, numRows, 1);

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


/**
 * Searches for member entry by email in `sheet` by binary search.
 * If unsuccessful, searches again via top-to-bottom iteration.
 * 
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} email  The email address to search for in `sheet`.
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

function findMemberByEmail(email, sheet) {
  let result = null;

  // First try with letter index store and binary search (fastest)
  if (sheet === MASTER_SHEET) result = findMemberWithStore(email);

  // Try with binary search with whole sheet if unsuccessful or sheet !== MASTER_SHEET
  result = result ?? findMemberByBinarySearch(email, sheet);

  // If binary search unsuccessful, try with iteration (slowest)
  result = result ?? findMemberByIteration(email, sheet);
  return result;
}


/**
 * Finds a member in the master sheet by setting a stricter bound using an
 * index store of each letter, searching with binary search.
 * 
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * @param {string} email  The email address to search for in `sheet`.
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
function findMemberWithStore(email, store = getIndexStore()) {
  const letter = email[0].toLowerCase();
  const BUFFER = 3;   // In case the store has not been updated

  // Set lower and upper bound for binary search
  const lowerBound = store[letter] - BUFFER;
  const upperBound = getUpperBound(letter) + BUFFER;

  // Use binary search with a smaller search range
  return findMemberByBinarySearch(email, MASTER_SHEET, lowerBound, upperBound);

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
 * @param {string} email  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * @param {number} [startRow=2]  The starting row index for the search (1-indexed).
 *                            Defaults to 2 (the second row) to avoid the header row.
 * @param {number} [endRow=MASTER_SHEET.getLastRow()]  The ending row index for the search.
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

function findMemberByIteration(email, sheet, startRow = 2, endRow = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  const thisEmailCol = GET_COL_MAP_(sheetName).emailCol;    // Get email col index of `sheet`

  for (var row = startRow; row <= endRow; row++) {
    let email = sheet.getRange(row, thisEmailCol).getValue();
    if (email === email) return row;    // Exit loop and return value;
  }

  return null;
}


/**
 * Recursive function to search for entry by email in `sheet` using binary search.
 * Returns row index of `email` in GSheet (1-indexed), or null if not found.
 * 
 * Previously `findSubmissionFromEmail` in `Master Scripts.gs`.
 * 
 * @param {string} email  The email address to search for in `sheet`.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * @param {number} [startRow=2]  The starting row index for the search (1-indexed). 
 *                            Defaults to 2 (the second row) to avoid the header row.
 * @param {number} [endRow=MASTER_SHEET.getLastRow()]  The ending row index for the search. 
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

function findMemberByBinarySearch(email, sheet, startRow = 2, endRow = sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  const emailCol = GET_COL_MAP_(sheetName).emailCol;  // Get email col from `sheet`

  // Base case: If start index exceeds the end index, the email is not found
  if (startRow > endRow) {
    return null;
  }

  if (startRow === endRow) {
    const emailAtRow = sheet.getRange(startRow, emailCol).getValue();
    if (emailAtRow === email) {
      return startRow;
    }
    else {
      return null;
    }
  }

  // Find the middle point between the start and end indexes
  const mid = Math.floor((startRow + endRow) / 2);

  // Get the email value at the middle row
  const emailAtMid = sheet.getRange(mid, emailCol).getValue();

  // Compare the target email with the middle email
  if (emailAtMid === email) {
    return mid;  // If the email matches, return the row index (1-indexed)

    // If the email at the middle row is alphabetically smaller, search the right half.
    // Note: use localeString() to ensure string comparison matches GSheet.
  } else if (emailAtMid.localeCompare(email) === -1) {
    return findMemberByBinarySearch(email, sheet, mid + 1, endRow);

    // If the email at the middle row is alphabetically larger, search the left half.
  } else {
    return findMemberByBinarySearch(email, sheet, startRow, mid);
  }
}


/**
 * Testing speed of different search functions.
 */
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
  findMemberWithStore(email);
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}
