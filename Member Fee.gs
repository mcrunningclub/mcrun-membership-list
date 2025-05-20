// LAST UPDATED: MAY 15, 2025
// PLEASE UPDATE WHEN NEEDED
const ZEFFY_EMAIL = 'contact@zeffy.com';
const INTERAC_EMAIL = 'interac.ca';    // Interac email addresses end in "interac.ca"
const STRIPE_EMAIL = 'stripe.com';

const ONLINE_LABEL = 'Fee Payments/Online Emails';
const INTERAC_LABEL = 'Fee Payments/Interac Emails';

// Found in `Internal Fee Collection` sheet
const INTERAC_ITEM_COL = 'A3';
const ONLINE_PAYMENT_ITEM_COL = 'A4';

function getPaymentItem_(colIndex) {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Internal Fee Collection")
    .getRange(colIndex)
    .getValue();
}

function getGmailLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName);
}


/**
 * Verify if member has paid fee using notification email sent by Interac, Stripe or Zeffy
 * 
 * Update member's information in MAIN_SHEET as required.
 * 
 * @param {number} Row index of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 16, 2025
 * @update  May 20, 2025
 */

function checkAndSetPaymentRef(row = getLastSubmissionInMain()) {
  const sheet = MAIN_SHEET;
  console.log('Entering `checkAndSetPaymentRef()` now...');

  // Get values of member's registration, and pack fee details for payment verifications
  const values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn())[0];
  const feeDetails = packFeeDetails(values);

  // Has the payment been found in inbox? 
  if (checkThisPayment(row, feeDetails)) {
    console.log(`Successfully found transaction email for '${feeDetails.memberName}'!`);  // Log success message
  }
  else {
    // 1) Create a scheduled trigger to recheck email inbox
    // 2) After max tries, send an email notification to RUN for missing payment
    console.error(`Unable to find payment confirmation email for '${feeDetails.memberName}'. Creating new scheduled trigger to check later.`);
    createNewFeeTrigger_(row, feeDetails);
  }

  /** Helper: pack fee details from GSheet values */
  function packFeeDetails(values) {
    // Allows to access 0-index array using GSheet indices (1-indexed)
    const getThis = (index) => values[index - 1];
    
    const obj = {
      'firstName': getThis(FIRST_NAME_COL), 
      'lastName': getThis(LAST_NAME_COL),
      'email': getThis(EMAIL_COL),
      'paymentMethod': getThis(PAYMENT_METHOD_COL),
      'interacRef' : getThis(INTERAC_REF_COL),
    };
    
    // Assemble full name for log and email
    obj.memberName = `${obj.firstName} ${obj.lastName}`;
    return obj;
  }
}

// Helper function for Stripe/Zeffy and Interac cases
function checkThisPayment(row, feeDetails) {
  const { firstName, lastName, email, paymentMethod, interacRef } = feeDetails;

  if (paymentMethod.includes('CC')) {
    return checkAndSetOnlinePayment(row, 
    { firstName: firstName, lastName: lastName, email: email });
  }
  else if (paymentMethod.includes('Interac')) {
    return checkAndSetInteracRef(row, 
    { firstName: firstName, lastName: lastName, interacRef: interacRef });
  }
  return false;
}


/**
 * Updates member's fee information.
 * 
 * @param {number} Row index to enter information.
 * @param {string} listItem  The list item in `Internal Fee Collection` to set in 'Collection Person' col.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Mar 16, 2025
 */

function setFeeDetails_(row, listItem) {
  const sheet = MAIN_SHEET;
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue(listItem);
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();
}


/**
 * Return latest emails of payment notification.
 * 
 * If not found, wait multiple times for email to arrive in McRUN inbox.
 * 
 * @param {string} sender  Email of sender (Interac, Stripe or Zeffy).
 * @param {number} maxMatches  Number of max tries.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 16, 2025
 * @update  Mar 17, 2025
 */

function getMatchingPayments_(sender, maxMatches, subject='') {
  // Ensure that correct mailbox is used
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw new Error('Wrong account! Please switch to McRUN\'s Gmail account');
  }

  const searchStr = getGmailSearchString_(sender, subject);
  let threads = [];
  let delay = 5 * 1000; // Start with 10 seconds

  // Search inbox until successful (max 2 tries)
  for (let tries = 0; tries < 2 && threads.length === 0; tries++) {
    if (tries > 0) Utilities.sleep(delay);  // Wait only on retries
    threads = GmailApp.search(searchStr, 0, maxMatches);
    delay *= 2; // Exponential backoff (5s → 10s → 20s)
  }

  return threads;
}


// Get threads from search (from:sender, starting:yesterday, in:inbox, [subject:partial-email-match])
function getGmailSearchString_(sender, subject = '') {
  const yesterday = new Date(Date.now() - 86400000); // Subtract 1 day in milliseconds
  const formattedYesterday = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy/MM/dd');
  return `from:(${sender}) in:inbox after:${formattedYesterday} subject:"${subject}"`;
}


/**
 * Marks a fully processed thread as read, archives it, and moves it to the `label` folder.
 */

function cleanUpMatchedThread_(thread, label) {
  thread.markRead();
  thread.moveToArchive();
  thread.addLabel(label);

  console.log('Thread cleaned up. Now removed from inbox');
}

/**
 * Checks if a member's information is present in the email body.
 * 
 * @param {string[]}  searchTerms. Search terms for match regex.
 * @param {string} emailBody  The body of the payment.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 15, 2025
 * @update Mar 21, 2025
 * 
 */

function matchMemberInPaymentEmail_(searchTerms, emailBody) {
  const formatedBody = emailBody.replace(/\*/g, '');    // Remove astericks around terms
  console.log(formatedBody);

  if (searchTerms.length === 0) return false; // Prevent empty regex errors

  const searchPattern = new RegExp(`\\b(${searchTerms.join('\\b|\\b')})\\b`, 'i');
  return searchPattern.test(formatedBody);
}


/**
 * Creates search terms for regex using member information.
 * 
 * Matches lastName whether hyphenated or not.
 * 
 * @param {Object}  Member information.
 * @param {string} member.firstName  Member's first name.
 * @param {string} member.lastName  Member's last name.
 * @param {string} [member.email]  Member's email address (if applicable).
 * @param {string} [member.interacRef]  Reference number of Interac e-Transfer (if applicable).
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 21, 2025
 * @update Mar 21, 2025
 * 
 */

function createSearchTerms(member) {
  const lastNameHyphenated = (member.lastName).replace(/[-\s]/, '[-\\s]?'); // handle hyphenated last names
  const fullName = `${member.firstName}\\s+${lastNameHyphenated}`;

  const searchTerms = [
    fullName,
    removeDiacritics(fullName),
    member.email,
    member.interacRef
  ].filter(Boolean); // Removes undefined, null, or empty strings

  return searchTerms;
}


///  👉 FUNCTIONS HANDLING STRIPE/ZEFFY TRANSACTIONS 👈  \\\

function setOnlinePaid_(row) {
  const onlinePaymentItem = getPaymentItem_(ONLINE_PAYMENT_ITEM_COL);
  setFeeDetails_(row, onlinePaymentItem);
}


/**
 * Verify Stripe/Zeffy payment transaction for latest registration.
 * 
 * Must have the member submission in last row of main sheet to work.
 * 
 * @param {integer} row  Member's row index in GSheet.
 * 
 * @param {Object} member  Member information.
 * @param {string} member.firstName  First name of member.
 * @param {string} member.lastName  Last name of member.
 * @param {string} member.email  Email of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 15, 2025
 * @update  May 17, 2025
 */

function checkAndSetOnlinePayment(row, member) {
  const sender = `${ZEFFY_EMAIL} OR ${STRIPE_EMAIL}`;
  const maxMatches = 5;
  const threads = getMatchingPayments_(sender, maxMatches, 'payment');

  const searchTerms = createSearchTerms(member);
  console.log('Search terms for email body', searchTerms);

  let isFound = threads.some(thread => processOnlineThread_(thread, searchTerms));
  if (isFound) {
    setOnlinePaid_(row);
  }

  return isFound;
}


/**
 * Process a single Gmail thread to find a matching member's payment.
 */

function processOnlineThread_(thread, searchTerms) {
  const messages = thread.getMessages();
  let starredCount = 0;
  let isFoundInMessage = false;

  for (const message of messages) {
    if (message.isStarred()) {
      starredCount++; // Already processed, skip
      continue;
    }

    const emailBody = message.getPlainBody();
    isFoundInMessage = matchMemberInPaymentEmail_(searchTerms, emailBody);

    if (isFoundInMessage) {
      message.star();
      starredCount++;
    }
  }

  if (starredCount === messages.length) {
    const onlineLabel = getGmailLabel_(ONLINE_LABEL);
    cleanUpMatchedThread_(thread, onlineLabel);
  }

  return isFoundInMessage;
}



///  👉 FUNCTIONS HANDLING INTERAC TRANSACTIONS 👈  \\\

function setInteractPaid_(row) {
  const interacItem = getPaymentItem_(INTERAC_ITEM_COL);
  setFeeDetails_(row, interacItem);
}

/**
 * Look for new emails from Interac starting yesterday (cannot search from same day) and extract ref number.
 * 
 * @trigger  New member registration.
 * @error  Send notification email to McRUN if no ref number found.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Apr 29, 2025
 */

function checkAndSetInteracRef(row, member) {
  const sender = INTERAC_EMAIL;
  const maxMatches = 10;
  const threads = getMatchingPayments_(sender, maxMatches);

  // Save results of thread processing
  let thisIsFound = false;
  const thisUnidentified = [];

  // Construct search terms once
  const searchTerms = createSearchTerms(member);
  console.log('Search terms for email body', searchTerms);

  // Most Interac email threads only have 1 message, so O(n) instead of O(n**2). Coded as safeguard.
  for (const thread of threads) {
    const result = processInteracThreads_(thread, searchTerms);

    // Set store.isFound to true iff result.isFound=true
    if (result.isFound) thisIsFound = true;
    thisUnidentified.push(...result.unidentified);
  }

  // Update member's payment information
  if (thisIsFound) {
    setInteractPaid_(row);
  }
  // Notify McRUN about references not identified
  else if (thisUnidentified.length > 0) {
    notifyUnidentifiedInteracRef_(thisUnidentified);
  }

  return thisIsFound;
}

// Interac e-Transfer emails can be matched by a reference number or full name
function processInteracThreads_(thread, searchTerms) {
  const messages = thread.getMessages();
  const result = { isFound: false, unidentified: [] };

  for (message of messages) {
    const emailBody = message.getPlainBody();

    // Try matching Interac e-Transfer email with member's reference number or name
    const isFoundInMessage = matchMemberInPaymentEmail_(searchTerms, emailBody);

    if (isFoundInMessage) {
      result.isFound = true;
      cleanUpMatchedThread_(thread, getGmailLabel_(INTERAC_LABEL));
      continue;
    }

    // If not found, store email's Interac references for later
    const emailInteracRef = extractInteracRef_(emailBody);
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
 * @return {string}  Returns extracted Interac Ref from `emailBody`, otherwise empty string.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 13, 2024
 * @update  Mar 16, 2025
 */

function extractInteracRef_(emailBody) {
  const searchPattern = /(Reference Number|Numero de reference)\s*:\s*(\w+)/;
  const match = emailBody.match(searchPattern);

  // If a reference is found, return it. Otherwise, return null
  // The Interac reference is in the second capturing group i.e. match[2];
  if (match && match[2]) {
    return match[2].trim();
  }

  return '';
}


function notifyUnidentifiedInteracRef_(references) {
  const emailBody =
  `
  Cannot identify new Interac e-Transfer Reference number(s): ${references.join(', ')}
      
  Please check the newest entry of the membership list.
      
  Automatic email created by 'Membership Collected (main)' script.
  `
  const errorEmail = {
    to: 'mcrunningclub@ssmu.ca',
    subject: 'ATTENTION: Interac Reference(s) to CHECK!',
    body: emailBody.replace(/[ \t]{2,}/g, '')
  };

  // Send warning email for unlabeled interac emails in inbox
  GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
}


function notifyUnidentifiedPayment_(name) {
  const emailBody =
  `
  Cannot find the payment confirmation email for member: ${name}
      
  Please manually check the inbox and update membership registry if required.

  If email not found, please notify member of outstanding member fee.
      
  Automatic email created by 'Membership Collected (main)' script.
  `
  const errorEmail = {
    to: 'mcrunningclub@ssmu.ca',
    subject: 'ATTENTION: Missing Member Payment!',
    body: emailBody.replace(/[ \t]{2,}/g, '')
  };

  GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
}
