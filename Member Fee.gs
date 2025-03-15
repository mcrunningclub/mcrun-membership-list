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
