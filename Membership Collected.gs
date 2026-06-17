/**
 * Handles the submission of a new registration form.
 * 
 * This function processes the latest submission in the semester sheet by:
 * - Trimming whitespace
 * - Fixing letter case
 * - Generating a unique member ID
 * - Adding missing items (e.g., checkboxes)
 * - Verifying payment information
 * - Sending communications to the new member
 * 
 * It also ensures that the data is added to the master sheet and sorted appropriately.
 * 
 * @param {number} [newRow=getLastSubmissionInMain()] - The row number of the new submission.
 *                                                      Defaults to the last row in the semester sheet.
 *
 * @see {@link getLastSubmissionInSemester} for how the last row is determined.
 * @see {@link addLastSubmissionToMaster_} for how the data is added to the master sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 18, 2023
 */
function onFormSubmit(newRow = getLastSubmissionInSemester()) {
  trimWhitespaceSemester_(newRow);
  fixRowCaseSemester_(newRow);
  encodeRowSemester_(newRow);
  addCheckboxSemester_(newRow);
  
  // Wrap around try-catch since GAS does not support async calls
  try {
    checkPaymentForSemester(newRow);
    sendNewMemberCommunications(newRow);
  }
  catch (e) {
    console.log(`Could not find payment OR transfer new registration to 'New Member Comms'`);
    throw Error(e);
  }
  finally {
    // Must add and sort AFTER extracting payment info from email
    setWaiverUrl_(newRow);
    addLastSubmissionToMaster_(newRow);

    // Applies all pending changes before sorting
    SpreadsheetApp.flush();
    tryAndSortSemester();   // Can only sort and format view if lock not acquired (to prevent concurrent runs)
  }
}


/**
 * Sends communications to a new member.
 * 
 * This function packages the member's information and transfers it to the
 * New Member Communications sheet for further processing.
 * 
 * @param {integer} row - The row number of the new member in the semester.
 * 
 * @see {@link packageMemberInfoInRow_} for how member information is packaged.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 18, 2023
 */

function sendNewMemberCommunications(row) {
  const memberInfo = packageMemberInfoInRow_(row);
  console.log(`Member info to export to 'NewMemberComms'\n`, memberInfo);
  NewMemberCommunications.createNewMemberCommunications(memberInfo);   // Transfer new member's value to external sheet
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
 * Find and set waiver url to new member registration.
 * 
 * Waiver is automatically saved by Fillout to folder with id `WAIVER_DRIVE_ID`.
 * 
 * @param {number} [row=getLastSubmissionInMain()]  Row index to find and set url.
 *                                                  Defaults to the last row in main sheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 1, 2025
 * @update  Mar 15, 2025
 */

function setWaiverUrl_(row = getLastSubmissionInSemester()) {
  const sheet = SEMESTER_SHEET;

  // Search for waiver link using member name
  const memberName = sheet.getSheetValues(row, FIRST_NAME_COL, 1, 2)[0].join(' ');
  const waiverUrl = findWaiverLink_(memberName);

  // Set value of waiver url
  sheet.getRange(row, WAIVER_COL).setValue(waiverUrl);
}


/**
 * Find waiver using member's name. Helper function for setWaiverUrl_.
 * 
 * Waiver is automatically saved by Fillout to folder with id `WAIVER_DRIVE_ID`.
 * 
 * @param {string} name  The name of member.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 1, 2025
 * @update  Sep 22, 2025
 */

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


/**
 * Get expiration date of member fee from given semester code.
 * 
 * Depends on `MEMBERSHIP_DURATION`.
 * 
 * @param {string} semesterCode  The 3-char code representing the semester and year.
 * @return {string}  Month and year of expiration e.g. Sep 2025
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 28, 2025
 * @update  Feb 28, 2025
 */

function getExpirationDate_(semesterCode) {
  const validDuration = MEMBERSHIP_DURATION;

  const semester = semesterCode.charAt(0);
  const expirationYear = '20' + (parseInt(semesterCode.slice(-2)) + validDuration)

  switch (semester) {
    case ('F'): return `Sep ${expirationYear}`;
    case ('W'): return `Jan ${expirationYear}`;
    case ('S'): return `Jun ${expirationYear}`;
    default: return null;
  };
}
