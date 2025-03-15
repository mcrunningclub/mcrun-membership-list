/**
 * Run formatting functions after new member submits a registration form.
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
  formatMainView();
  getInteracRefNumberFromEmail_(newRow);

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
    addLastSubmissionToMaster();
    sortMainByName();
    SpreadsheetApp.flush();   // Applies all pending changes before executing again
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
 * Checks if new submission paid using Interac e-Transfer and completes collection info.
 * 
 * Must have the new member submission in the last row to work.
 * 
 * Helper function for getReferenceNumberFromEmail_()
 * 
 * @param {string} emailInteracRef  The Interac e-Transfer reference found in email.
 * @return {integer}  Returns status code.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 10, 2025
 */

function enterInteracRef_(emailInteracRef, row = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;

  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');
  const userInteracRef = sheet.getRange(row, INTERACT_REF_COL);

  if (userInteracRef.getValue() != emailInteracRef) {
    return false;
  }

  // Copy the '(isInterac)' list item in `Internal Fee Collection` to set in 'Collection Person' col
  const interacItem = getInteracItem();

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue(interacItem);
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();

  return true;   // Success!
}


/**
 * Look for new emails from Interac starting yesterday (cannot search for day of) and extract ref number.
 * 
 * Interac email address end in "interac.ca"
 * 
 * @trigger  New member registration.
 * @error  Send notification email to McRUN if no ref number found.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 11, 2025
 */

function getInteracRefNumberFromEmail_(row = MAIN_SHEET.getLastRow()) {
  const paymentForm = MAIN_SHEET.getRange(row, PAYMENT_METHOD_COL).getValue();

  if (!(String(paymentForm).includes('Interac'))) {
    return;
  }
  // else if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
  //   throw new Error('Please verify the club\'s inbox to search for the Interac email');
  // }

  Utilities.sleep(60 * 1000);   // If payment by Interac, allow *60 sec* for Interac email confirmation to arrive

  // Format start search date (yesterday) for GmailApp.search()
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd');

  const interacLabelName = "Interac Emails";
  const searchStr = `from:(interac.ca) in:inbox after:${formattedYesterday}`;
  const threads = GmailApp.search(searchStr, 0, 10);

  if (threads.length === 0) {
    throw new Error(`No Interac e-Transfer emails in inbox. Please verify again for latest member registration.`);
  }

  const checkTheseRef = [];
  const interacLabel = GmailApp.getUserLabelByName(interacLabelName);

  // Most Interac emails only have 1 message, so O(n) instead of O(n**2). Coded as safeguard.
  for (thread of threads) {
    for (message of thread.getMessages()) {
      const emailBody = message.getPlainBody();

      // Extract Interac e-transfer reference
      const interacRef = extractInteracRef_(emailBody);
      const isSuccess = enterInteracRef_(interacRef);

      // Success: Mark thread as read and archive it
      if (isSuccess) {
        thread.markRead();
        thread.moveToArchive();
        thread.addLabel(interacLabel);
      }
      else {
        checkTheseRef.push(interacRef);
      }
    }
  }

  if (checkTheseRef.length > 0) {
    const errorEmail = {
      to: 'mcrunningclub@ssmu.ca',
      subject: 'ATTENTION: Interac Reference(s) to CHECK!',
      body: 
  `
  Cannot identify new Interac e-Transfer Reference number(s): ${checkTheseRef.join(', ')}
      
  Please check the newest entry of the membership list.
      
  Automatic email created by 'Membership Collected (main)' script.
  `
    };

    // Send warning email for unlabeled interac emails in inbox
    GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
  }
}


/**
 * Extract Interac e-Transfer reference string.
 * 
 * Helper function for getReferenceNumberFromEmail_().
 * 
 * @param {string} emailBody  The body of the Interac e-Transfer email.
 * @return {string || null}  Returns extracted Interac Ref from `emailBody`, otherwise null.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2024
 * @update  Feb 10, 2025
 */

function extractInteracRef_(emailBody) {
  const searchPattern = /(Reference Number|Numero de reference)\s*:\s*(\w+)/;
  const match = emailBody.match(searchPattern);

  // If a reference is found, return it. Otherwise, return null
  // The interac reference is in the second capturing group i.e. match[2];
  if (match && match[2]) {
    return match[2].trim();
  }

  return null;
}


function setWaiverUrl(row = MAIN_SHEET.getLastRow()) {
  const sheet = MAIN_SHEET;

  // Search for waiver link using member name
  const memberName = sheet.getSheetValues(row, FIRST_NAME_COL, 1, 2)[0].join(' ');
  const waiverUrl = findWaiverLink_(memberName);

  // Set value of waiver url
  sheet.getRange(row, WAIVER_COL).setValue(waiverUrl);
}


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

