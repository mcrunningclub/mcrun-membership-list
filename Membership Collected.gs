/**
 * Runs formatting functions after new member submits a registration form.
 * 
 * https://developers.google.com/apps-script/samples/automations/event-session-signup
 * 
 * https://stackoverflow.com/questions/62246016/how-to-check-if-current-form-submission-is-editing-response
 *
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  May 28, 2024
 */

function onFormSubmit(newRow=getLastSubmissionInMain()) {
  trimWhitespace_(newRow);
  fixLetterCaseInRow_(newRow);
  encodeLastRow(newRow);   // create unique member ID

  addMissingItems_(newRow);
  formatMainView();
  getInteracRefNumberFromEmail_(newRow);
  
  // Must add and sort AFTER extracting Interac info from email
  addLastSubmissionToMaster();
  //sortMainByName();
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
  
  if(resultBinarySearch != null) return resultBinarySearch;   // success!

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

function findMemberByIteration(emailToFind, sheet, start=2, end=sheet.getLastRow()) {
  const sheetName = sheet.getSheetName();
  const thisEmailCol = GET_COL_MAP_(sheetName).emailCol;    // Get email col index of `sheet`
  
  for(var row=start; row <= end; row++) {
    let email = sheet.getRange(row, thisEmailCol).getValue();
    if(email === emailToFind) return row;    // Exit loop and return value;
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

function findMemberByBinarySearch(emailToFind, sheet, start=2, end=sheet.getLastRow()) {
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
    if (i%2 == 0) continue;
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
 * @update  Oct 8, 2023
 */

function enterInteracRef_(emailInteracRef) {
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');
  const sheet = MAIN_SHEET;

  const newSubmissionRow = sheet.getLastRow();
  const userInteracRef = sheet.getRange(newSubmissionRow, INTERACT_REF_COL);

  if (userInteracRef.getValue() != emailInteracRef) return false;
  
  // Copy the '(isInterac)' list item in `Internal Fee Collection` to set in 'Collection Person' col
  var interacItem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal Fee Collection").getRange(INTERAC_ITEM_COL).getValue();

  sheet.getRange(newSubmissionRow, IS_FEE_PAID_COL).check();
  sheet.getRange(newSubmissionRow, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(newSubmissionRow, COLLECTION_PERSON_COL).setValue(interacItem);
  sheet.getRange(newSubmissionRow, IS_INTERNAL_COLLECTED_COL).check();

  return true;   // Success!
}

function testIt() {
  getInteracRefNumberFromEmail_();
}

/**
 * Look for new emails from Interac starting today (form trigger date) and extract ref number.
 * 
 * Interac email address either "catch@payments.interac.ca" or "notify@payments.interac.ca"
 * 
 * @trigger  New member registration.
 * @error  Send notification email to McRUN if no ref number found.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 10, 2025
 */

function getInteracRefNumberFromEmail_(row=MAIN_SHEET.getLastRow()) {
  const paymentForm = MAIN_SHEET.getRange(row, PAYMENT_METHOD_COL).getValue();
  if ( !(String(paymentForm).includes('Interac')) ) return;

  // If payment by Interac, allow Interac email confirmation to arrive in inbox
  //Utilities.sleep(1 * 60 * 1000);   // 1 minute

  // Format start search date (yesterday) for GmailApp.search()
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd'); 

  const interacLabelName = "Interac Emails";
  //const threads = GmailApp.search(`from:(interac.ca) in:inbox NOT(label:"${interacLabel}") after:${formattedYesterday}`, 0, 10);
  const threads = GmailApp.search(`from:(interac.ca) in:inbox NOT(label:"${interacLabelName}")`, 0, 10);
  
  if (threads.length === 0) {
    throw new Error(`No Interac e-Transfer Ref matches the latest submission. Please verify inbox`);
  }

  let found = false;
  const refToCheck = [];

  const interacLabel = GmailApp.getUserLabelByName(interacLabelName);
  const firstThread = threads[0];

  for (message of firstThread.getMessages()) {
    const emailBody = message.getPlainBody();

    // Add label to thread
    firstThread.addLabel(interacLabel);

    // Extract Interac e-transfer reference
    const interacRef = extractInteracRef_(emailBody);
    const isSuccess = enterInteracRef_(interacRef);

    // Success: Mark thread as read and archive it
    if (isSuccess) {
      found = true;
      firstThread.markRead();
      firstThread.moveToArchive();
    }
    else {
      refToCheck.push(interacRef);
    }
  }

  if(true) {
    var errorEmail = {
      to: 'mcrunningclub@ssmu.ca',
      cc: '',
      subject: 'ATTENTION: Interac Reference to CHECK!',
      body: `
      Cannot identify new Interac e-Transfer Reference number: ${refToCheck.join('; ')}
      
      Please check the newest entry of the membership list.
      
      Automatic email created by 'Membership Collected (main)' script.`
    }
      
    // Send warning email if reference code cannot be found
    GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
  }

  //firstThread.markUnread();   // needed??

  return;

  




  /* const interacRef = messages.forEach((msg, i) => {
    threads[i].addLabel(GmailApp.getUserLabelByName("Interac Emails"));   // Label as `Interac Emails`

    const emailBody = msg.getPlainBody();
    const ref = extractInteracRef_(emailBody);

    if((enterInteracRef_(ref)) === 0) {
      firstThread.markRead();
      firstThread.moveToArchive();  // remove from inbox
      return ref;
    }
  });

  // Error handling: mark Interac email unread & send notification email to McRUN
  if (!interacRef) {
    firstThread.markUnread();

    var errorEmail = {
      to: 'mcrunningclub@ssmu.ca',
      cc: "",
      subject: 'ERROR: Interac Reference to CHECK!',
      body: `
      Cannot identify new Interac e-Transfer Reference number: ${referenceNumberString}
      
      Please check the newest entry of the membership list.
      
      Automatic email created by 'Membership Collected (main)' script.
      `
    }
        
    // Send warning email if reference code cannot be found
    GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
  }
  
  return;
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd'); 

  const threads = GmailApp.search(`from:(interac.ca) in:inbox after:${formattedYesterday}`, 0, 10);

  const firstThread = threads[0];
  const messages = firstThread.getMessages();  

  // Loop through messages
  for (let i=0; i<messages.length; i++) {
    const emailBody = messages[i].getPlainBody();
    threads[i].addLabel(GmailApp.getUserLabelByName("Interac Emails"));   // Label as `Interac Emails`

    referenceNumberString = extractInteracRef_(emailBody);  // search for Interac e-transfer ref in email
    var errorCode = enterInteracRef_(referenceNumberString);  // confirm number with newest entry in membership list

    // Email with ref is found
    if (errorCode === 0) {
      firstThread.markRead();
      firstThread.moveToArchive();  // remove from inbox
      break;    // email found, exit loop
    }
    
    // Error handling: mark Interac email unread & send notification email to McRUN
    firstThread.markUnread();

    var errorEmail = {
      to: "mcrunningclub@ssmu.ca",
      cc: "",
      subject: "ERROR: Interac Reference to CHECK!",
      body: "Cannot identify new Interac e-Transfer Reference number: " + referenceNumberString + "\n\nPlease check the newest entry of the membership list.\n\n\nAutomatic email created by \'Membership Collected (main)\' script."
    }
        
    // Send warning email if reference code cannot be found
    GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
  } */
}


/**
 * Extract Interac e-Transfer reference string.
 * 
 * Helper function for getReferenceNumberFromEmail_()
 * 
 * @param {string} emailBody  The body of the Interac e-Transfer email.
 * @return {string}  Returns extracted Interac Ref from `emailBody`.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2024
 * @update  Nov 13, 2024
 */

function extractInteracRef_(emailBody) {
  const searchString = "Reference Number:";
  const searchStringFR = "Numero de reference :";  // Accents not required

  // Try searching in English
  let startIndex = emailBody.indexOf(searchString) + searchString.length + 1;

  // Now in French
  if(startIndex < 0 ) {
    startIndex = emailBody.indexOf(searchStringFR) + searchStringFR.length + 2;
    searchString = searchStringFR;
  }

  // Extract substring of length 20, and split after '\n'
  var referenceNumberString = emailBody.substring(startIndex, startIndex + 20);
  var newlineIndex = referenceNumberString.indexOf('\n', 1);
    
  referenceNumberString = (referenceNumberString.substring(0, newlineIndex)).trim(); // trim everything after newline
  return referenceNumberString;
}

