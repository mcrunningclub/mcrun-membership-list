/**
 * Runs formatting functions after form submission by new member.
 * 
 * https://developers.google.com/apps-script/samples/automations/event-session-signup
 * 
 * https://stackoverflow.com/questions/62246016/how-to-check-if-current-form-submission-is-editing-response
 *
 * 
 * CURRENTLY IN REVIEW (may-28)
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 1, 2023
 * @update  May 28, 2024
 */

function onFormSubmit() {
  trimWhitespace_();
  fixLetterCaseInRow_();

  encodeLastRow();   // create unique member ID
  //copyNewMemberToPointsLedger();  // copy new member to `Points Ledger`

  formatMainView();
  //getReferenceNumberFromEmail_();
  
  // Must add and sort AFTER getting Interac info and copying
  addLastSubmissionToMaster();
  //sortMainByName();
}


/**
 * Find last submission using while loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sept 1, 2024
 * @update  Sept 1, 2024
 */

function getLastSubmissionInMain() {
  const sheet = MAIN_SHEET;
  const lastRow = sheet.getLastRow();

  while (sheet.getRange(lastRow, REGISTRATION_DATE_COL).getValue() == "") {
    lastRow = lastRow - 1;
  }

  return lastRow;
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

  if (userInteracRef.getValue() != emailInteracRef) return -1;
  
  // Copy the '(isInterac)' list item in `Internal Fee Collection` to set in 'Collection Person' col
  var interacItem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal Fee Collection").getRange(INTERAC_ITEM_COL).getValue();

  sheet.getRange(newSubmissionRow, IS_FEE_PAID_COL).check();
  sheet.getRange(newSubmissionRow, COLLECTION_DATE_COL).setValue(currentDate);
  sheet.getRange(newSubmissionRow, COLLECTION_PERSON_COL).setValue(interacItem);
  sheet.getRange(newSubmissionRow, IS_INTERNAL_COLLECTED_COL).check();

  return 0;   // Success!
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
 * @update  Nov 13, 2024
 */

function getReferenceNumberFromEmail_() {
  const sheet = MAIN_SHEET;
  const lastRow = sheet.getLastRow();
  
  const paymentForm = MAIN_SHEET.getRange(lastRow, PAYMENT_METHOD_COL).getValue();
  if ( !(String(paymentForm).includes('Interac')) ) return;   // Exit if Interac is not chosen

  // If payment by Interac, allow Interac email confirmation to arrive in inbox
  Utilities.sleep(1 * 60 * 1000);   // 1 minute
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
  const threads = GmailApp.search('from:payments.interac.ca "Reference Number" "label:inbox" after:' + currentDate, 0, 10);
  //const threads = GmailApp.search('from:payments.interac.ca "INTERAC" after:2024/10/20', 0, 5); // TEST
  
  // If none found, Interac email has not arrived or not found
  if (threads.length < 1) return;
  
  var firstThread = threads[0];
  var messages = firstThread.getMessages();
  
  // Loop through messages
  for (var i=0; i<messages.length; i++) {
    var emailBody = messages[i].getPlainBody();
    threads[i].addLabel(GmailApp.getUserLabelByName("Interac Emails"));   // Label as `Interac Emails`

    var referenceNumberString = extractInteracRef_(emailBody);  // search for Interac e-transfer ref in email
    var errorCode = enterInteracRef_(referenceNumberString);  // confirm number with newest entry in membership list

    // Email with ref is found
    if (errorCode == 0) {
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
  }
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


/**
 * Copies the new member to the `Points Ledger` spreadsheet.
 * 
 * @trigger  New form submission.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 28, 2023
 * @update  Nov 28, 2023
 */

function copyNewMemberToPointsLedger() {

  const LEDGER_SPREADSHEET = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1J-nSg2QLNYkVWc0PplfwQWM8fyujE1Dv_PURL6kBNXI/edit?usp=sharing");
  
  const copySheet = MAIN_SHEET;

  const pasteSheet = LEDGER_SPREADSHEET.getSheetByName('Member Points');
  const PASTE_MEMBER_ID_COL = 1;
  const PASTE_FEE_PAID_COL = 2;
  const PASTE_FULL_NAME_COL = 3;

  // Get last row index of copy sheet
  const copyRowNum = copySheet.getLastRow();
  const pasteRowNum = pasteSheet.getLastRow();

  // Set member id in `Member Points`
  const copyMemberIDRange = copySheet.getRange(copyRowNum, MEMBER_ID_COL);
  const pasteMemberIDRange = pasteSheet.getRange(pasteRowNum, PASTE_MEMBER_ID_COL);
  pasteMemberIDRange.setValue(copyMemberIDRange.getValue());

  // Set fee paid in `Member Points`
  const copyFeePaidRange = copySheet.getRange(copyRowNum, IS_FEE_PAID_COL);
  const pasteFeePaidRange = pasteSheet.getRange(pasteRowNum, PASTE_FEE_PAID_COL);
  pasteFeePaidRange.setValue(copyFeePaidRange.getValue());

  // Get full name from `MAIN_SHEET`
  const nameArray = copySheet.getRange(copyRowNum, FIRST_NAME_COL, 1, 2).getValues()[0];
  const firstName = nameArray[0];
  const lastName = nameArray[1];

  // Set full name in `Member Points` 
  const fullName = firstName + " " + lastName;
  pasteSheet.getRange(pasteRowNum, PASTE_FULL_NAME_COL).setValue(fullName);

  // Set formulas in `Member Points`
  const TPOINTS_FORMULA_COL = 4;
  const REGISTERED_FORMULA_COL = 5;
  const FEE_FORMULA_COL = 6;
  const PASS_SAVED_FORMULA_COL = 7;

  // e.g. `=SUM(E123:AA123)` & need a regEx to replace all instances of `{row}`
  const totalPointsFormula = "=SUM(E{row}:AA{row})".replace( getRegEx_('{row}'), pasteRowNum );
  pasteSheet.getRange(pasteRowNum, TPOINTS_FORMULA_COL).setValue(totalPointsFormula);

  const registrationFormula = "=pointsMemberRegistration";
  pasteSheet.getRange(pasteRowNum, REGISTERED_FORMULA_COL).setValue( registrationFormula );

  // e.g. `=IF(B123, pointsPaidFee, 0)`
  const feeFormula = "=IF(B{row}, pointsPaidFee, 0)".replace('{row}', pasteRowNum);
  pasteSheet.getRange(pasteRowNum, FEE_FORMULA_COL).setValue( feeFormula );

  // e.g. =IFNA(IF( INDEX('Import Pass Data'!AX$2:AX, ...
  const passSavedFormula = 
    "=IFNA(IF( INDEX('Import Pass Data'!AX$2:AX, MATCH(A{row}, 'Import Pass Data'!B$2:B, 0)) = \"PASS_INSTALLED\", pointsSavedDigitalPass, 0), \"UNKNOWN\")".replace('{row}', pasteRowNum);

  pasteSheet.getRange(pasteRowNum, PASS_SAVED_FORMULA_COL).setValue(passSavedFormula);
  return;
}
