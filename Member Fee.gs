// LAST UPDATED: MAR 15, 2025
// PLEASE UPDATE WHEN NEEDED
const ZEFFY_EMAIL = 'contact@zeffy.com';
const INTERAC_EMAIL = 'interac.ca';    // Interac email address end in "interac.ca"

const ZEFFY_LABEL = 'fee-payments-zeffy-emails';
const INTERAC_LABEL = 'fee-payments-interac-emails';


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


// @TODO: combine with Interac version
function setZeffyPaid_(row) {
  const sheet = MAIN_SHEET;
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue('(Online Payment)');
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();
}


function getMatchingPayments_(sender) {
  const searchStr = getGmailSearchString(sender);
  let threads = [];
  let delay = 10000; // Start with 10 seconds

  // Search inbox until successful (max 3 tries)
  for (let tries = 0; tries < 3 && threads.length === 0; tries++) {
    if (tries > 0) Utilities.sleep(delay);  // Wait only on retries
    threads = GmailApp.search(searchStr, 0, 3);
    delay *= 2; // Exponential backoff (10s → 20s → 40s)
  }

  return threads;
}

// Get threads from search (from:sender, starting:yesterday, in:inbox)
function getGmailSearchString(sender) {
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd');
  return `from:(${sender}) in:inbox after:${formattedYesterday}`;
}


///  👉 FUNCTIONS HANDLING ZEFFY TRANSACTIONS 👈  \\\

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

  let isFound = threads.some(thread => processThread(thread, member));
  if (isFound) {
    setZeffyPaid_(row);
  }

  return isFound;
}

/**
 * Process a single Gmail thread to find a matching member's payment.
 */
function processThread(thread, member) {
  const messages = thread.getMessages();
  let starredCount = 0;
  let isFound = false;

  for (const message of messages) {
    if (message.isStarred()) {
      starredCount++; // Already processed, skip
      continue;
    }

    const emailBody = message.getPlainBody();
    isFound = matchMemberInEmail(member, emailBody);

    if (isFound) {
      message.star();
      starredCount++;
    }
  }

  if (starredCount === messages.length) {
    cleanUpMatchedThread(thread);
  }

  return isFound;


  /**
   * Checks if a member's name or email is present in the email body.
   */
  function matchMemberInEmail(member, emailBody) {
    const strippedName = removeDiacritics(member.name);
    const searchPattern = new RegExp(`\\b(${member.email}|${member.name}|${strippedName})\\b`, 'i');
    return searchPattern.test(emailBody);
  }

  /**
   * Marks a fully processed thread as read, archives it, and moves it to the Zeffy folder.
   */
  function cleanUpMatchedThread(thread) {
    thread.markRead();
    thread.moveToArchive();

    const zeffyLabel = GmailApp.getUserLabelByName(ZEFFY_LABEL);
    if (zeffyLabel) {
      thread.addLabel(zeffyLabel);
    }

    console.log('Thread fully matched. Now removed from inbox');
  }
}



///  👉 FUNCTIONS HANDLING INTERAC TRANSACTIONS 👈  \\\

/**
 * Checks if new submission paid using Interac e-Transfer and completes collection info.
 * 
 * Must have the new member submission in the last row to work.
 * 
 * Helper function for getReferenceNumberFromEmail_()
 * 
 * @param {string} emailInteracRef  The Interac e-Transfer reference found in email.
 * @param {number} [row=getLastSubmissionInMain()]  Row index to enter Interac ref.
 *                                                  Defaults to the last row in main sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  March 16, 2025
 */

function enterInteracRef_(row = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;

  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

  // Copy the '(e-Transfer)' list item in `Internal Fee Collection` to set in 'Collection Person' col
  const interacItem = getPaymentItem(INTERAC_ITEM_COL);

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue(interacItem);
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();
}


/**
 * Look for new emails from Interac starting yesterday (cannot search for day of) and extract ref number.
 * 
 * @trigger  New member registration.
 * @error  Send notification email to McRUN if no ref number found.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Feb 11, 2025
 */

function checkAndSetInteracRef(row = MAIN_SHEET.getLastRow()) {
  const sheet = MAIN_SHEET;
  const userInteracRef = sheet.getRange(row, INTERAC_REF_COL).getValue();

  Utilities.sleep(30 * 1000);   // If payment by Interac, allow *30 sec* for Interac email confirmation to arrive

  getGmailSearchString_(sender)

  // Format start search date (yesterday) for GmailApp.search()
  const interacLabelName = INTERAC_LABEL;

  const searchStr = getGmailSearchString_(INTERAC_EMAIL);
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

      if (userInteracRef.getValue() != emailInteracRef) {
        return false;
      }

      // Success: Mark thread as read and archive it
      if (userInteracRef === emailInteracRef) {
        enterInteracRef_(interacRef);

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
    const emailBody =
      `
    Cannot identify new Interac e-Transfer Reference number(s): ${checkTheseRef.join(', ')}
        
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
