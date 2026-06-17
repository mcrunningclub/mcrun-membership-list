/**
 * Retrieves the payment item from the "Internal Fee Collection" sheet.
 *
 * @param {string} cell  The cell reference (e.g., 'A3') to retrieve the payment item from.
 * @return {string}  The payment item value (e.g. "(Online Payment)") from the specified cell.
 *
 * @author Andrey Gonzalez
 * @date  May 24, 2025
 */
function getPaymentItem_(cell) {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Internal Fee Collection")
    .getRange(cell)
    .getValue();
}

/**
 * Retrieves a Gmail label by its name.
 *
 * This function fetches a Gmail label object based on the provided label name.
 *
 * @param {string} labelName  The name of the Gmail label to retrieve.
 * @return {GmailApp.GmailLabel} The Gmail label object corresponding to the provided name.
 *
 */
function getGmailLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName);
}



/**
 * Verify if member has paid fee using notification email sent by Interac, Stripe or Zeffy
 * 
 * Update member's information in the semester sheet as required.
 * 
 * @param {integer} [row=getLastSubmissionInMain()] index of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 16, 2025
 * @update  Jun 6, 2025
 */
function checkPaymentForSemester(row = getLastSubmissionInSemester()) {
  const sheet = SEMESTER_SHEET;
  console.log('Entering `checkAndSetPaymentRef()` now...');

  // Get values of member's registration, and pack fee details for payment verifications
  const values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn())[0];
  const feeDetails = packFeeDetails(values);

  // Has the payment been found in inbox or fee waived?
  if (isPaid_(row, feeDetails)) {
    console.log(`Successfully found transaction email for '${feeDetails.memberName}' (or fee waived)!`);  // Log success message
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

/**
 * Helper function for Stripe/Zeffy and Interac cases
 * 
 * Calls necessary function to check whether payment has been
 * made depending on type of payment
 * 
 * @param {number} row  Row of the member to check payment for
 * @param {struct} feeDetails  struct of fee information from sheet??
 * 
 * @return {bool}  Whether member's fee has been paid or not
 */
function isPaid_(row, feeDetails) {
  const { firstName, lastName, email, paymentMethod, interacRef } = feeDetails;

  if (paymentMethod.includes('CC')) {
    return checkAndSetOnlinePayment_(row, 
    { firstName, lastName, email});
  }
  else if (paymentMethod.includes('Interac')) {
    return checkAndSetInteracRef_(row, 
    { firstName, lastName, interacRef });
  }
  else if (paymentMethod.includes('Waived')) {
    setFeeWaived_(row);
    return true;  // Always returns true
  }
  return false;
}


/**
 * Updates member's fee information.
 * 
 * @param {integer} row  The index to enter information.
 * @param {string} collectedBy  The list item from `Internal Fee Collection` to put in 'Collection Person' col.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  Mar 16, 2025
 */
function setFeeDetailsInSemester_(row, collectedBy) {
  const sheet = SEMESTER_SHEET;
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

  sheet.getRange(row, IS_FEE_PAID_COL).check();
  sheet.getRange(row, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(row, COLLECTION_PERSON_COL).setValue(collectedBy);
  sheet.getRange(row, IS_INTERNAL_COLLECTED_COL).check();
}


/**
 * Updates member's fee information in the master sheet.
 * 
 * @param {integer} row  The index of the row with member's information.
 * @param {string} paymentMethod  How the fee was paid
 * @param {string} collectedBy  How the fee was collected
 * @param {string} date  Date of collection. Defaults to null (will be set to current date)
 */
function setFeeDetailsInMaster_(row, paymentMethod, collectedBy, date = null) {
  const sheet = MASTER_SHEET;
  const collectionDate = date ?? Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');
  sheet.getRange(row, MASTER_COLS.collectionDate).setValue(collectionDate);

  //sheet.getRange(row, MASTER_COLS.paymentHistory).setValue(...);

  // Get cleaner string of payment method using list item
  const collector = collectedBy ?? paymentMethodToItem_(paymentMethod).toString();
  sheet.getRange(row, MASTER_COLS.collectedBy).setValue(collector);
  sheet.getRange(row, MASTER_COLS.isInternalCollected).check();
}

/**
 * Updates fee payment information in master sheet given member's email.
 * 
 * @param {string} email  Email address of member to update info for.
 * @param {string} paymentMethod  How fee was paid
 * @param {number} row  Row of member (if known). Defaults to null.
 */
function updateMasterPayment_(email, paymentMethod, row = null) {
  // Find member row in MASTER_SHEET
  const rowInMaster = row ?? findMemberByEmail(email, MASTER_SHEET);
  
  // Set fee details in MASTER_SHEET
  setFeeDetailsInMaster_(rowInMaster, paymentMethod);
  addPaidSemesterToMaster_(rowInMaster, SHEET_NAME);
}


/**
 * For members whose fee is not paid in the master sheet, check whether
 * they have paid the fee in the semester sheet. Update master sheet if necessary.
 * 
 * Loop through every member in the master sheet. If their fee status is not paid,
 * check whether registration date is within the last semester and whether the semester
 * sheet contains their payment. If found, add to master sheet. If not found, add to
 * list of "unpaid" emails that is logged in console.
 */
function checkExistingPaymentInSemester() {
  const masterSheet = MASTER_SHEET;
  const startRow = 1;
  const numRowsMaster = masterSheet.getLastRow();

  // Get all emails and isPaid values in both MASTER and SEMESTER sheets
  const getColumnVals = (col) => masterSheet.getSheetValues(startRow, col, numRowsMaster, 1);
  const masterEmails = getColumnVals(MASTER_COLS.email);
  const masterRegDates =  getColumnVals(MASTER_COLS.latestRegistration);
  const masterFeeStatuses = getColumnVals(MASTER_COLS.feePaid);

  // Now in SEMESTER SHEET
  const semesterSheet = SEMESTER_SHEET;
  const semesterName = semesterSheet.getSheetName();
  const numRowsSemester = semesterSheet.getLastRow();

  const SEMESTER_MAP = GET_COL_MAP_(semesterName);
  const semesterFeeStatuses = semesterSheet.getSheetValues(1, SEMESTER_MAP.feeStatus, numRowsSemester, 1);

  const emailsNotFound = [];
  const today = new Date();

  // Iterate through rows and verify payment in SEMESTER_SHEET iff isPaid false
  for (let isPaid, lastReg, masterRow = startRow; masterRow <= numRowsMaster - 1; masterRow++) {
    isPaid = masterFeeStatuses[masterRow][0];
    lastReg = new Date(masterRegDates[masterRow][0]);

    const daysSinceReg = getNumberOfDays(lastReg, today);

    if (isPaid !== 'Paid' && daysSinceReg > 60 && daysSinceReg < 150) {
      let email = masterEmails[masterRow][0];
      const semesterRow = findMemberByEmail(email, semesterSheet);
      console.info(`Checking fee for '${email}' in semester sheet row #${semesterRow}`);

      if (!semesterRow) {
        emailsNotFound.push(email);   // Not found in 'SEMESTER'
      }
      else if (semesterFeeStatuses[semesterRow - 1][0]) {
        // Transfer from SEMESTER TO MASTER and add semester code
        const collectedBy = semesterSheet.getRange(semesterRow, SEMESTER_MAP.collector).getValue();
        const collectionDate = semesterSheet.getRange(semesterRow, SEMESTER_MAP.collectionDate).getValue();

        setFeeDetailsInMaster_(masterRow + 1, null, collectedBy, collectionDate)
        addPaidSemesterToMaster_(masterRow + 1, semesterName);
        logSuccessfulTransfer(semesterRow, masterRow + 1);
      }
    }
  }

  // Finally log all emails that could not be found in semester sheet
  if(emailsNotFound.length > 0) {
    console.error(`Unable to find in 'SEMESTER'...\n${emailsNotFound.join('\n')}`);
  }

  /** Helper function to calculate number of days */ 
  function getNumberOfDays(startDate, endDate) {
    const startTime = startDate.getTime();
    const endTime = endDate.getTime();

    const differenceInMilliseconds = endTime - startTime;
    const millisecondsPerDay = 1000 * 60 * 60 * 24;

    return Math.round(differenceInMilliseconds / millisecondsPerDay);
  }

  /** Helper function to add log messages for debugging */
  function logSuccessfulTransfer(sRow, mRow) {
    console.log(`Successfully transferred fee details from 'SEMESTER' row #${sRow} to 'MASTER' row #${mRow}`);
  }
}


/**
 * Return latest emails of payment notification.
 * 
 * If not found, wait multiple times for email to arrive in McRUN inbox.
 * Must use club email.
 * 
 * @param {string} sender  Email of sender (Interac, Stripe or Zeffy).
 * @param {integer} maxMatches  Number of max tries.
 * @param {string} subject  Email subject to filter by. Defaults to empty string
 * @return {GmailThread[]}  Gmail threads matching the search
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 16, 2025
 * @update  Mar 17, 2025
 */
function getMatchingEmails_(sender, maxMatches, subject='') {
  // Ensure that correct mailbox is used
  if (getCurrentUserEmail_() !== MCRUN_EMAIL) {
    throw new Error('Wrong account! Please switch to McRUN\'s Gmail account');
  }

  const searchStr = createGmailSearchString_(sender, subject);
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


/**
 * Create search string given sender and optional subject
 * 
 * In the form (from:sender, starting:yesterday, in:inbox, [subject:partial-email-match])
 */ 
function createGmailSearchString_(sender, subject = '') {
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
 * @param {string[]} searchTerms  Search terms for match regex.
 * @param {string} emailBody  The body of the payment.
 * @returns {boolean}  True if a match is found, false otherwise.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 15, 2025
 * @update  Jun 7, 2025
 */
function searchInEmail_(searchTerms, emailBody) {
  const formatedBody = emailBody.replace(/\*/g, '');   // Remove asterisks around terms
  console.log(formatedBody);

  if (!searchTerms.length) return false;  // Prevent empty regex errors
  const searchPattern = new RegExp(`${searchTerms.join('|')}`, 'i');
  return searchPattern.test(formatedBody);
}


/**
 * Creates search terms for regex matching using a member's information.
 * 
 * Handles optional hyphens/spaces in last names, and removes diacritics for better matching.
 * Improves matching accuracy in `searchInEmail_`.
 * 
 * @param {Object}  Member information. Contains attributes member.firstname, member.lastname,
 *                    member.email (if applicable), member.interacRef (if applicable).
 * @returns {string[]}  An array of search terms for regex matching.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 21, 2025
 * @update  Sep 22, 2025
 */
function createSearchTerms_(member) {
  const escapeRegex = str => str ? str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : '';

  const nameParts = [
    member.firstName,
    member.lastName.replace(/[-\s]/g, '[-\\s]?'),  // Allow optional hyphen/space
  ].filter(Boolean);

  // Construct regex for ordered matching: \bWord\b.*?\bWord\b...
  const orderedNamePattern = nameParts
    .map(part => `\\b${escapeRegex(part)}\\b`)
    .join('.*?');

  const diacriticPattern = removeDiacritics_(orderedNamePattern);

  const searchTerms = [
    orderedNamePattern,
    diacriticPattern,
    member.email,
    escapeRegex(member.interacRef)
  ].filter(Boolean);   // Removes undefined, null, or empty strings

  return searchTerms;
}


///  👉 FUNCTIONS FOR FEE WAIVED 👈  \\\

/**
 * Get payment item eg. "(Online Payment)" from payment method string
 * 
 * Gets standardized item from list in Internal Memberships Collected spreadsheet 
 * based on keywords in payment method. 
 */
function paymentMethodToItem_(paymentMethod) {
  const itemCol = (() => {
    if (paymentMethod.includes('CC')) {
      return ONLINE_PAYMENT_ITEM_COL;
    }
    else if (paymentMethod.includes('Interac')) {
      return INTERAC_ITEM_COL;
    }
    else if (paymentMethod.includes('Waived')) {
      return FEE_WAIVED_ITEM_COL;
    }
  })();

  return getPaymentItem_(itemCol);
}

/**
 * Sets fee status as waived in member registration.
 * 
 * @param {integer} row  The index to enter information.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Jun 6, 2025
 * @update  Jun 6, 2025
 */
function setFeeWaived_(row) {
  const feeWaivedItem = getPaymentItem_(FEE_WAIVED_ITEM_COL);
  setFeeDetailsInSemester_(row, feeWaivedItem);
}


///  👉 FUNCTIONS HANDLING STRIPE/ZEFFY TRANSACTIONS 👈  \\\

/**
 * Sets fee status as paid online in member registration.
 * 
 * @param {integer} row  The index to enter information.
 */
function setOnlinePaid_(row) {
  const onlinePaymentItem = getPaymentItem_(ONLINE_PAYMENT_ITEM_COL);
  setFeeDetailsInSemester_(row, onlinePaymentItem);
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
 * @return {boolean}  True if payment was found in emails.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 15, 2025
 * @update  May 17, 2025
 */
function checkAndSetOnlinePayment_(row, member) {
  const sender = `${ZEFFY_EMAIL} OR ${STRIPE_EMAIL}`;
  const maxMatches = 5;
  const threads = getMatchingEmails_(sender, maxMatches, 'payment');

  const searchTerms = createSearchTerms_(member);
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
    isFoundInMessage = searchInEmail_(searchTerms, emailBody);

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

/**
 * Sets fee status as paid through Interac in member registration.
 * 
 * @param {integer} row  The index to enter information.
 */
function setInteracPaid_(row) {
  const interacItem = getPaymentItem_(INTERAC_ITEM_COL);
  setFeeDetailsInSemester_(row, interacItem);
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
function checkAndSetInteracRef_(row, member) {
  const sender = INTERAC_EMAIL;
  const maxMatches = 10;
  const threads = getMatchingEmails_(sender, maxMatches);

  // Save results of thread processing
  let thisIsFound = false;
  const thisUnidentified = [];

  // Construct search terms once
  const searchTerms = createSearchTerms_(member);
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
    setInteracPaid_(row);
  }
  // Notify McRUN about references not identified
  // else if (thisUnidentified.length > 0) {
  //   notifyUnidentifiedInteracRef_(thisUnidentified);
  // }

  return thisIsFound;
}

// Interac e-Transfer emails can be matched by a reference number or full name
function processInteracThreads_(thread, searchTerms) {
  const messages = thread.getMessages();
  const result = { isFound: false, unidentified: [] };

  for (message of messages) {
    const emailBody = message.getPlainBody();

    // Try matching Interac e-Transfer email with member's reference number or name
    const isFoundInMessage = searchInEmail_(searchTerms, emailBody);

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

/**
 * Sends an email to the club with a list of Interac references that have not
 * been matched to a member registration.
 */
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

/**
 * Sends an email to the club with member whose payment emails has not been found.
 */
function notifyUnidentifiedPayment_(name) {
  const emailBody =
  `
  Cannot find the payment confirmation email for member: '${name}'
      
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
