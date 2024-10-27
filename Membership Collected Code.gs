/* SHEET CONSTANTS */
const SHEET_NAME = 'Fall 2024';       // MUST UPDATE EVERY SEMESTER!
const MAIN_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
const TIMEZONE = getUserTimeZone();


function getUserTimeZone() {
  return Session.getScriptTimeZone();
}

/**
 * @author: Andrey S Gonzalez
 * @date: Oct 1, 2023
 * @update: May 28, 2024
 * 
 * Runs these functions on new form submission.
 * 
 * CURRENTLY IN REVIEW (may-28)
 */
function onFormSubmit() {
  trimWhitespace();
  encodeLastRow();   // create unique member ID
  
  //copyToBackup();   // transfer info to `BACKUP` for Zapier Automation. PassKit URL copied back to `main`
  //copyNewMemberToPointsLedger();  // copy new member to `Points Ledger`

  formatSpecificColumns();
  getReferenceNumberFromEmail();
  
  // Must add and sort AFTER getting Interac info and copying
  addLastSubmissionToMaster();
  sortNameByAscending();
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 1, 2023
 * @update: Jun 1, 2024
 * 
 * Sorts `main' by first name.
 * Triggered by form submission.
 */
function sortNameByAscending() {
  const sheet = MAIN_SHEET;

  const numRows = sheet.getLastRow() - 1;     // Remove header row from count
  const numCols = sheet.getLastColumn();
  
  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);
  
  // Sorts values by `First Name` then by `Last Name`
  range.sort([{column: 3, ascending: true}, {column: 4, ascending: true}]);
  return;
}

/**
 * @author: Andrey S Gonzalez
 * @date: Nov 28, 2023
 * @update: Nov 28, 2023
 * 
 * Copies the new member to the `Points Ledger` spreadsheet
 * Triggered by new form submit.
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
  const totalPointsFormula = "=SUM(E{row}:AA{row})".replace( getRegEx('{row}'), pasteRowNum );
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


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 1, 2023
 * @update: Oct 8, 2023
 * 
 * Checks if new submission paid using Interac e-Transfer and completes collection info
 * Triggered by `getReferenceNumberFromEmail`
 * 
 * Must have the new submission in the last row to work.
 */

function checkInteracRef(emailInteracRef) {
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');
  const sheet = MAIN_SHEET;

  // Column numbers (double check if correct numbers)

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
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 17, 2023
 * 
 * Trim whitespace from specific columns in last row.
 * Trigger: new form submission
 */

function trimWhitespace() {
  const sheet = MAIN_SHEET;
  
  const lastRow = sheet.getLastRow();
  const rangeNames = sheet.getRange(lastRow, FIRST_NAME_COL, 1, 7);   // range is [First Name, Name of Referral]
  rangeNames.trimWhitespace();

  return;
}


/**
 * @author: Andrey S Gonzalez
 * @date: Sept 1, 2024
 * @update: Sept 1, 2024
 * 
 * Find last submission using while loop.
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @RETURN integer
 */

function getLastSubmission() {
  const sheet = MAIN_SHEET;
  const lastRow = sheet.getLastRow();

  while (sheet.getRange(lastRow, REGISTRATION_DATE_COL).getValue() == "") {
    lastRow = lastRow - 1;
  }

  return lastRow;

}

/**
 * @author: Andrey S Gonzalez
 * @date: Oct 9, 2023
 * @update: Oct 20, 2024
 * 
 * Modified MD5 hash function to define member_id from email *only*.
 * Changed `sheet.getLastRow()` to user-defined `getLastSubmission()`
 */
function encodeLastRow() {
  const sheet = MAIN_SHEET;
  //const sheet = BACKUP_SHEET;
  const newSubmissionRow = getLastSubmission();
  
  const email = sheet.getRange(newSubmissionRow, EMAIL_COL).getValue();
  const member_id = MD5(email);
  sheet.getRange(newSubmissionRow, MEMBER_ID_COL).setValue(member_id);
}


function encodeWholeList() {
  var sheet = MAIN_SHEET;
  var i, email;

  // Start at row 2 (1-indexed)
  for (i = 2; i <= sheet.getMaxRows(); i++) {
    email = sheet.getRange(i, EMAIL_COL).getValue();
    if (email === "") return;   // check for invalid row

    var member_id = MD5(email);
    sheet.getRange(i, MEMBER_ID_COL).setValue(member_id);
  }
}


/** 
 * @author: Andrey S Gonzalez
 * @date: Oct 1, 2023
 * @update: Oct 18, 2023
 * 
 * Formats spreadsheet for easy user experience
 */ 

function formatSpecificColumns() {
  var sheet = MAIN_SHEET;

  const rangeRegistration = sheet.getRange('A2:A');  // Range for Preferred Name/Pronouns
  const rangePreferredName = sheet.getRange('E2:E');  // Range for Preferred Name/Pronouns
  const rangeWaiver = sheet.getRange('J2:J');         // Range for Waiver
  const rangePaymentChoice = sheet.getRange('K2:K');  // Range for Payment Preferrence
  const rangeInteracRef = sheet.getRange('L2:L');     // Range for Interac e-Transfer Reference
  const rangeCollection = sheet.getRange('O2:P');     // Range for Collection Info
  const rangeMemberId = sheet.getRange('T2:T');       // Range for Member Id
  const rangeURL = sheet.getRange('U2:U');            // Range for PassKit URL

  // Set ranges to Bold
  rangeRegistration.setFontWeight('bold');
  rangePreferredName.setFontWeight('bold');
  rangePaymentChoice.setFontWeight('bold');
  rangeInteracRef.setFontWeight('bold');
  rangeCollection.setFontWeight('bold');
  rangeMemberId.setFontWeight('bold');
  rangeURL.setFontWeight('bold');

  // Set Text Wrapping to 'Clip'
  rangePaymentChoice.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  rangeWaiver.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Align ranges to Left
  rangePaymentChoice.setHorizontalAlignment('left');

  // Centre these ranges
  rangeInteracRef.setHorizontalAlignment('center');
  rangeCollection.setHorizontalAlignment('center');
  rangeMemberId.setHorizontalAlignment('center');
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 1, 2023
 * @update: Oct 24, 2024
 * 
 * Look for new emails from Interac starting today (form trigger date) and extract ref number
 * Triggered by new member registration. Creates error notification email if no ref number found.
 * 
 * Interac email address either "catch@payments.interac.ca" or "notify@payments.interac.ca"
 */

function getReferenceNumberFromEmail() {
  const sheet = MAIN_SHEET;
  const lastRow = sheet.getLastRow();
  
  const isInterac = MAIN_SHEET.getRange(lastRow, PAYMENT_METHOD_COL).getValue();
  if ( !(String(isInterac).includes('Interac')) ) return;   // Exit if Interac is not chosen

  // If payment by Interac, allow Interac email confirmation to arrive to inbox
  Utilities.sleep(1 * 60 * 1000);   // 1 minute
  const currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
  const threads = GmailApp.search('from:payments.interac.ca "Reference Number" "label:inbox" after:' + currentDate, 0, 10);
  //const threads = GmailApp.search('from:payments.interac.ca "INTERAC" after:2024/10/20', 0, 5); // TEST

  if (threads.length < 1) return;    // if no results, then interac email has not arrived
  var firstThread = threads[0];
  var messages = firstThread.getMessages();
  
  for (var i=0; i<messages.length; i++) {
    var emailBody = messages[i].getPlainBody();
    threads[i].addLabel(GmailApp.getUserLabelByName("Interac Emails"));   // Label as `Interac Emails`

    var searchString = "Reference Number:";
    var searchStringFR = "Numero de reference :";  // Accents not required

    // Try searching in English
    var startIndex = emailBody.indexOf(searchString) + searchString.length + 1;

    // Now in French
    if(startIndex < 0 ) {
      startIndex = emailBody.indexOf(searchStringFR) + searchStringFR.length + 2;
      searchString = searchStringFR;
    }

    // Extract substring of length 20, and split after '\n'
    var referenceNumberString = emailBody.substring(startIndex, startIndex + 20);
    var newlineIndex = referenceNumberString.indexOf('\n', 1);
    
    referenceNumberString = (referenceNumberString.substring(0, newlineIndex)).trim(); // trim everything after newline
    var errorCode = checkInteracRef(referenceNumberString); // confirm number with newest entry in membership list

    // Error Handling
    if (errorCode != 0) { 
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
    
    else { 
      firstThread.markRead();
      firstThread.moveToArchive(); // remove from inbox
      break;      //exit for-loop
    }
  }
}


/**
 * @author: Andrey S Gonzalez
 * @date: Oct 8, 2023
 * Returns regex expression for target string
 * 
 * @returns {string} 
 */

function getRegEx(targetSubstring) {
  return RegExp(targetSubstring, 'g');
}


/**
 * @author: https://stackoverflow.com/questions/7994410/hash-of-a-cell-text-in-google-spreadsheet
 * @date: Oct 8, 2023
 * Hash function using modified MD5 algorithm. Used for members' External ID
 * 
 * @returns {string} 
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
 * @author: Andrey S Gonzalez
 * @date: Oct 17, 2023
 * @update: Oct 17, 2023
 * 
 * Check if PassKit URL is entered.
 */

function isPasskitURL() {
  return;
  const sheet = MAIN_SHEET;
  const newSubmissionRow = sheet.getLastRow();
  const urlCol = sheet.getLastColumn();

  const rangeURL = sheet.getRange(newSubmissionRow, urlCol);
  return rangeURL.isBlank;
}


/* DEPRICATED OR JUNK FUNCTIONS */

function drafts() {
  return;

  /**
   * @author: Andrey S Gonzalez
   * @date: Oct 1, 2023
   * @update: Oct 8, 2023
   * 
   * Verifies if Collection Date and Collection Person were added before checking `Fee Paid` box
   * If not, displays up warning message
   */ 

  function _onFormSubmit(e) {
    //https://developers.google.com/apps-script/samples/automations/event-session-signup
    //https://stackoverflow.com/questions/62246016/how-to-check-if-current-form-submission-is-editing-response
  }

  function _checkInterac() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Fall 2023');
    
    var newSubmissionRow = sheet.getLastRow();
    var interacRef = sheet.getRange(newSubmissionRow, INTERACT_REF_COL);

    var trimmed = interacRef.trimWhitespace();
    var isSingleWord = (trimmed.getValue().split(' ').length) == 1; // Verifies if letters unseparated

    if (interacRef.getValue() != "" && isSingleWord) {
      var currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'MMM d, yyyy');

      // Used to copy the '(isInterac)' list item to set in 'Collection Person' col
      var interacItem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internal Fee Collection").getRange('V4').getValue();

      sheet.getRange(newSubmissionRow, IS_FEE_PAID_COL).check();
      sheet.getRange(newSubmissionRow, COLLECTION_DATE_COL).setValue(currentDate);
      sheet.getRange(newSubmissionRow, COLLECTION_PERSON_COL).setValue(interacItem);
      sheet.getRange(newSubmissionRow, IS_INTERNAL_COLLECTED_COL).check();
    }
  }


  function _getReferenceNumberFromEmail() {
    var currentDate = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    var threads = GmailApp.search('(from:notify@payments.interac.ca) "Reference Number" after:' + currentDate, 0, 1);
    //TEST: var threads = GmailApp.search('(from:notify@payments.interac.ca) "Reference Number" after:2023/10/01', 0, 1);
    
    for (var i=0; i<threads.length; i++) {
      var messages = threads[i].getMessages();

      for (var j=0; j<messages.length; j++) {
        var emailBody = messages[j].getPlainBody();
        Logger.log(emailBody);

        var searchString = "Reference Number:";
        var startIndex = emailBody.indexOf(searchString) + searchString.length + 1;

        // Extract substring of length 20, and split after '\n'
        var referenceNumberString = emailBody.substring(startIndex, startIndex + 20);
        var newlineIndex = referenceNumberString.indexOf('\n', 1);
        
        referenceNumberString = referenceNumberString.substring(0, newlineIndex); // trim everything after newline
        var errorCode = checkInterac(referenceNumberString.trim()); // confirm number with newest entry in membership list

        // Error Handling
        if (errorCode != 0) { 
          threads[i].markUnread();
          threads[i].addLabel(GmailApp.getUserLabelByName("Interac Emails"));

          var errorEmail = {
            to: "mcrunningclub@ssmu.ca",
            subject: "ERROR: Interac Reference to CHECK!",
            body: "Cannot identify new Interac e-Transfer Reference number: " + referenceNumberString + "\n\nPlease check the newest entry of the membership list."
          }
          
          // Send warning email if reference code cannot be found
          GmailApp.sendEmail(errorEmail.to, errorEmail.subject, errorEmail.body);
        }

      }
    }
  }


  function _onEdit(e) {
    var sheet = e.range.getSheet();
    if (sheet.getName() != 'Fall 2023') return;  // Exit if incorrect sheet
    if (e.range.getValue() != true) return;   // Exit if box not checked

    var editRange = { // L2:L
      top : 2,
      col : 12
    };

    // Find column and row of checked box
    var thisCol = e.range.getColumn();
    var thisRow = e.range.getRow();
    
    // Exit if we're out of range
    if (thisCol != editRange.col || thisRow < editRange.top) return;
    
    // Get value of neighbouring Date and Person cells
    var collectionDate = sheet.getRange("M" + thisRow).getValue();
    var collectionPerson = sheet.getRange("N" + thisRow).getValue();

    // If cells empty, issue warning and set note on the edited cell to indicate when it was changed.
    if(collectionDate == "" || collectionPerson == "") {
      var longMessage = 'Make sure that you enter your name and collection date.\nThank you!';
      SpreadsheetApp.getUi().alert('⚠️ Change Detected ⚠️', longMessage, SpreadsheetApp.getUi().ButtonSet.OK);

      e.range.setNote('Last modified ' + new Date() + '\n\n' + Session.getActiveUser().getEmail());
    }
  }


  /**
   * @author: Andrey S Gonzalez
   * @date: Oct 1, 2023
   * @update: Oct 1, 2023
   * 
   * Remove empty rows from `MASTER` sheet
   * @WARNING: no longer in use since `MASTER` deleted
   */   

  function _deleteEmptyRows() {
    Utilities.sleep(3000);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("MASTER");
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var row = sheet.getLastRow();

    while (row > 2) {
      var rec = data.pop();
      if (rec.join('').length === 0) {
        sheet.deleteRow(row);
      }
      row--;
    }

    var maxRows = sheet.getMaxRows(); 
    var lastRow = sheet.getLastRow();

    if (maxRows - lastRow != 0) {
      sheet.deleteRows(lastRow + 1, maxRows - lastRow);
    }
  }


  /** 
   * @author: Andrey S Gonzalez
   * @date: Oct 1, 2023
   * @update: Oct 8, 2023
   * 
   * Sorts MAIN_SHEET by first name
   * Triggered by form submission
   */
  function _sortNameByAscending() {
    const sheet = MAIN_SHEET;
    var lastColumnLetter = getLetterFromColumnIndex(sheet.getLastColumn());
    
    // Sort all the way to the last row
    var range = sheet.getRange("A2:" + lastColumnLetter + sheet.getLastRow());
    
    // Sorts values by the `First Name` column in ascending order
    range.sort(3);
  }


  /* HELPER FUNCTIONS */

  // Returns the letter representation of the column index
  // e.g. "1" returns "A"; "5" returns "E"
  function _getLetterFromColumnIndex(column) {
    var temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  // Returns the column number from letter representation of column
  // e.g. "A" returns "1"; "E" returns "5"
  function _getColumnFromLetter(letter) {
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }

  function _getAllSheetNames() {
    var out = new Array()
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i=0 ; i<sheets.length ; i++) 
      out.push( [ sheets[i].getName() ] )
    return out
  }

  /* END OF JUNK FUNCTIONS */
}