// LAST UPDATED: MAR 15, 2025
// PLEASE UPDATE WHEN NEEDED
const ZEFFY_EMAIL = 'contact@zeffy.com';
const INTERAC_EMAIL = 'interac.ca';    // Interac email address end in "interac.ca"

const ZEFFY_LABEL = 'fee-payments-zeffy-emails';
const INTERAC_LABEL = 'fee-payments-interac-emails';

// Found in `Internal Fee Collection` sheet
const INTERAC_ITEM_COL = 'A3';
const ONLINE_PAYMENT_ITEM_COL = 'A4';

function getPaymentItem(colIndex) {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Internal Fee Collection")
    .getRange(colIndex)
    .getValue();
}

function getGmailLabel(labelName) {
  return GmailApp.getUserLabelByName(labelName);
}


function checkAndSetPaymentRef_(row = getLastSubmissionInMain()) {
  // Get values of member's registration
  const sheet = MAIN_SHEET;
  const values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn())[0].unshift('');

  const memberEmail = values[EMAIL_COL];
  const memberName = `${values[FIRST_NAME_COL]} ${values[LAST_NAME_COL]}`;
  const memberPaymentMethod = values[PAYMENT_METHOD_COL];
  const memberInteracRef = values[INTERAC_REF_COL];

  // Has the payment been found in inbox?
  const isFound = checkPayment(memberPaymentMethod);

  if (!isFound) {
    throw new Error(`Unable to find Zeffy email for member ${memberName}. Please verify again.`);
  }

  function checkPayment(paymentMethod) {
    if (paymentMethod.includes('CC')) {
      return checkAndSetZeffyPayment(row, { name: memberName, email: memberEmail });
    }
    else if (paymentMethod.includes('Interac')) {
      return checkAndSetInteracRef(row, memberInteracRef);
    }
    return false;
  }
}


/**
 * Updates member's fee information.
 * 
 * @param {number} Row index to enter information.
 * @param {string} listItem  The list item in `Internal Fee Collection` to set in 'Collection Person' col.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  March 16, 2025
 */

function setFeeDetails(row, listItem) {
  const sheet = MAIN_SHEET;
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue(listItem);
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();
}


function getMatchingPayments_(sender, maxMatches) {
  const searchStr = getGmailSearchString_(sender);
  let threads = [];
  let delay = 10000; // Start with 10 seconds

  // Search inbox until successful (max 3 tries)
  for (let tries = 0; tries < 3 && threads.length === 0; tries++) {
    if (tries > 0) Utilities.sleep(delay);  // Wait only on retries
    threads = GmailApp.search(searchStr, 0, maxMatches);
    delay *= 2; // Exponential backoff (10s → 20s → 40s)
  }

  return threads;
}


// Get threads from search (from:sender, starting:yesterday, in:inbox)
function getGmailSearchString_(sender) {
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd');
  return `from:(${sender}) in:inbox after:${formattedYesterday}`;
}


/**
 * Marks a fully processed thread as read, archives it, and moves it to the `label` folder.
 */

function cleanUpMatchedThread(thread, label) {
  thread.markRead();
  thread.moveToArchive();
  thread.addLabel(label);

  console.log('Thread cleaned up. Now removed from inbox');
}



///  👉 FUNCTIONS HANDLING ZEFFY TRANSACTIONS 👈  \\\

function setZeffyPaid_(row) {
  const onlinePaymentItem = getPaymentItem(ONLINE_PAYMENT_ITEM_COL);
  setFeeDetails(row, onlinePaymentItem);
}


/**
 * Verify Zeffy payment transaction for latest registration.
 * 
 * Must have the member submission in last row of main sheet to work.
 * 
 * @param {integer} row  Member's row index in GSheet.
 * 
 * @param {Object} member  Member information.
 * @param {string} member.name  Name of member.
 * @param {string} member.email  Email of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 15, 2025
 * @update  Mar 16, 2025
 */

function checkAndSetZeffyPayment(row, member) {
  const sender = ZEFFY_EMAIL;
  const threads = getMatchingPayments_(sender);

  let isFound = threads.some(thread => processZeffyThread(thread, member));
  if (isFound) {
    setZeffyPaid_(row);
  }

  return isFound;
}


/**
 * Process a single Gmail thread to find a matching member's payment.
 */
function processZeffyThread(thread, member) {
  const messages = thread.getMessages();
  let starredCount = 0;
  let isFound = false;

  for (const message of messages) {
    if (message.isStarred()) {
      starredCount++; // Already processed, skip
      continue;
    }

    const emailBody = message.getPlainBody();
    isFound = matchMemberInZeffyEmail(member, emailBody);

    if (isFound) {
      message.star();
      starredCount++;
    }
  }

  if (starredCount === messages.length) {
    const zeffyLabel = getGmailLabel(ZEFFY_LABEL);
    cleanUpMatchedThread(thread, zeffyLabel);
  }

  return isFound;


  /**
   * Checks if a member's name or email is present in the email body.
   */
  function matchMemberInZeffyEmail(member, emailBody) {
    const strippedName = removeDiacritics(member.name);
    const searchPattern = new RegExp(`\\b(${member.email}|${member.name}|${strippedName})\\b`, 'i');
    return searchPattern.test(emailBody);
  }
}



///  👉 FUNCTIONS HANDLING INTERAC TRANSACTIONS 👈  \\\


function setInteractPaid_(row) {
  const interacItem = getPaymentItem(INTERAC_ITEM_COL);
  setFeeDetails(row, interacItem);
}


/**
 * Look for new emails from Interac starting yesterday (cannot search from same day) and extract ref number.
 * 
 * @trigger  New member registration.
 * @error  Send notification email to McRUN if no ref number found.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Mar 16, 2025
 */

function checkAndSetInteracRef(row, memberInteracRef) {
  const sender = INTERAC_EMAIL;
  const maxMatches = 10;
  const threads = getMatchingPayments_(sender, maxMatches);

  // Save results of thread processing
  const store = { isFound : false, unidentified : [] };

  // Most Interac email threads only have 1 message, so O(n) instead of O(n**2). Coded as safeguard.
  for (const thread of threads) {
    const result = processInteracThreads_(thread, memberInteracRef);

    // Set store.isFound to true iff result.isFound=true
    if (result.isFound) store.isFound = true;
    unidentified.push(...result.unidentified);
  }

  // Update member's payment information
  if (isFound) {
    setInteractPaid_(row);
  }
  // Notify McRUN about references not identified
  else if (unidentified.length > 0){
    emailUnidentifiedRef(unidentified);
  }

  return store.isFound;
}


function processInteracThreads_(thread, memberInteracRef) {
  const messages = thread.getMessages();
  const result = {isFound : false, unidentified : []};

  for (message of messages) {
    const emailBody = message.getPlainBody();

    // Extract Interac e-Transfer reference from email
    const emailInteracRef = extractInteracRef_(emailBody);

    if (memberInteracRef === emailInteracRef) {
      result.isFound = true;
      cleanUpMatchedThread(thread, getGmailLabel(INTERAC_LABEL));
      continue;
    }

    // Store unidentified Interac references
    (result.unidentified).push(emailInteracRef);
  }

  return result;
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


function emailUnidentifiedRef(references) {
  const emailBody =
  `
  Cannot identify new Interac e-Transfer Reference number(s): ${references.join(', ')}
      
  Please check the newest entry of the membership list.
      
  Automatic email created by 'Membership Collected (main)' script.
  `
  const errorEmail = {
    to: 'mcrunningclub@ssmu.ca',
    subject: 'ATTENTION: Interac Reference(s) to CHECK!',
    body: emailBody
  };

  // Send warning email for unlabeled interac emails in inbox
  GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
}



