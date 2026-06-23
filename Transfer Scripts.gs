/**
 * When a sheet is changed, check if a row was added and process it accordingly.
 * 
 * If new row added to Import sheet, try to add the data to the semester sheet, format
 * it, and add to master sheet. If new row added to master sheet by the app, try to format it and
 * add the data to the semester sheet.
 * 
 * @param {Event} e  Edit event
 */
function onChange(e) {
  // Get details of edit event's sheet
  console.log({
    authMode: e.authMode.toString(),
    changeType: e.changeType,
    user: e.user,
  });

  const source = e.source;

  // Try-catch to prevent errors when sheetId cannot be found
  try {
    const sheetChanged = source.getSheetId();

    const changeType = e.changeType;
    console.log(`Change Type: ${changeType}`);

    if (sheetChanged === IMPORT_SHEET_ID && changeType === 'INSERT_ROW') {
      console.log('Executing if block from onChange(e)...');
      const importSheet = source.getSheetById(sheetChanged);
      const newRow = source.getLastRow();
      const newRegistration = importSheet.getRange(newRow, 1).getValue();

      const lastRow = copyFilloutRegToSemester_(newRegistration);
      onFormSubmit(lastRow);

      // Successful execution...
      console.log('Exiting `onFormSubmit` from onChange(e) successfully!');
    }

    else if (sheetChanged === MASTER_SHEET.getSheetId() && changeType === 'INSERT_ROW') {
      console.log('Executing else-if block from onChange(e)...');
      console.log('New row added to Master sheet');

      const newRow = source.getLastRow();
      // Add formula, then copy submission to main sheet
      if(!isNewMemberViaApp_(newRow)) {
        console.log('Exiting because member *not* added by app');
      }
      formatFeeCollection_(newRow);
      setMemberId_(MASTER_SHEET, newRow);
      copyMasterRowToSemester(newRow);
      console.log('Added registration to main sheet from master!');

      notifyNewAppSubmission(newRow);
      console.log('Sent an email notification!');

      sortMasterByEmail(); // Sort 'MASTER' by email once formatting completed
      
      // Successful execution...
      console.log('Exiting `formatFeeCollection` from onChange(e) successfully!');
    }
  }
  catch (error) {
    console.log('Whoops! Error raised in onChange(e)');
    Logger.log(error);

    /* // Only alert if message is not "Please select an active sheet first"
    if (! /"active"/i.test(error.message)) {
      throw Error(error.message);
    } */
  }
}

/**
 * Checks whether a row (in the master sheet) was added from the app or not.
 * 
 * Registrations added from the app do not have the registration semester, so
 * it needs to be added.
 * 
 * @param {number} row  Row to check.
 * @return {boolean}  True if row was added from app, false if not.
 */
function isNewMemberViaApp_(row) {
  const sheet = MASTER_SHEET;

  // STEP 1: Check if 3-char code for registration semester exists
  const rangeRegSem = sheet.getRange(row, MASTER_COLS.LATEST_REG_SEMESTER);
  const isBlank = rangeRegSem.trimWhitespace().isBlank();
  if (!isBlank) return false;   // Reg sem exists... not added via app

  // STEP 2: Get regSem using semester sheet, then add to target row
  const regSem = getSemesterCode_(SHEET_NAME);   // Get semCode from `MAIN_SHEET`
  rangeRegSem.setValue(regSem);
  
  // STEP 3: Finally return true once added
  return true;
}

/**
 * Encode member's email from given sheet and row, and sets it in the Member ID column.
 * 
 * @param {SpreadSheetApp.Sheet} sheet  Sheet object that row is from.
 * @param {number} row  Row of member to make ID for.
 */
function setMemberId_(sheet, row) {
  // STEP 1: Get email col in target sheet
  const sheetName = sheet.getSheetName();
  const colMap = GET_COL_MAP_(sheetName);
  const thisEmailCol = colMap.emailCol;

  // STEP 2: Encode email using `encodeFromInput`
  const email = sheet.getRange(row, thisEmailCol).getValue();
  const memberId = encodeFromInput_(email);

  // STEP 3: Set member id in target row
  const thisIdCol = colMap.memberIdCol;
  sheet.getRange(row, thisIdCol).setValue(memberId);
}

/**
 * If master or semester sheet is edited, copies changes to the
 * other sheet as well.
 * 
 * When a sheet is edited, check whether it was the master or semester sheet, 
 * and if the edited range was within bounds of sheet contents. If so, get
 * member row in source (edited) and target (other) sheet, and call updateFeeInfo
 * 
 * @param {Event} e  Edit event
 */
function onEdit(e) {
  // Get details of edit event's sheet
  const rangeEdited = e.range;
  const sheetEdited = rangeEdited.getSheet();
  const sheetEditedName = sheetEdited.getName();

  var debug_e = {
    //authMode:  e.authMode,
    range: e.range.getA1Notation(),
    sheetName: e.range.getSheet().getSheetName(),
    //source:  e.source,
    newValue: e.value,
    oldValue: e.oldValue
  }
  console.log(debug_e);

  if (rangeEdited.getNumRows() > 2) return;  // prevent sheet-wide changes
  else if (rangeEdited.getNumColumns() > CELL_EDIT_LIMIT) {
    // TODO: add function to individually process changes
    Logger.log(`More than ${CELL_EDIT_LIMIT} columns edited at once`);
  }

  console.log(`onEdit 1 -> thisSheetName: ${sheetEditedName}`);

  // Check if legal sheet (neither SHEET_NAME OR MASTER_NAME)
  if (!(sheetEditedName == SHEET_NAME || sheetEditedName == MASTER_NAME)) return;

  console.log("onEdit 1a -> Passed first check");

  //if(e.value == e.oldValue) return;   // Values have not changed. Edit was on sheet formatting.
  //console.log("onEdit 1b -> Passed second check");

  // Check if legal edit
  if (!isLegalEdit_(rangeEdited, sheetEdited)) return;

  console.log("onEdit 2 -> Passed \`isLegalEdit()\`");

  // Get the email column for the current sheet
  const emailCol = GET_COL_MAP_(sheetEditedName).emailCol;
  const row = rangeEdited.getRow();

  console.log(`onEdit 3 -> thisEmailCol: ${emailCol} thisRow: ${row}`);

  // Get email from row and column
  const email = sheetEdited.getRange(row, emailCol).getValue();

  const isSemesterSheet = (sheetEditedName === SHEET_NAME);
  console.log(`onEdit 4 -> email: ${email}  |  isMainSheet: ${isSemesterSheet}`);

  const sourceSheet = isSemesterSheet ? SEMESTER_SHEET : MASTER_SHEET;
  const targetSheet = isSemesterSheet ? MASTER_SHEET : SEMESTER_SHEET;
  const targetRow = findMemberByEmail(email, targetSheet);  // Find row of member in `targetSheet` using their email

  // Throw error message if member not in `targetSheet`
  if (targetRow === null) {
    const errorMessage = `
      --- onEdit() ---
      targetRow not found in ${targetSheet}. 
      Edit made in ${sheetEditedName} at row ${row}.
      Email of edited member: ${email}. Please review this error.
    `
    throw Error(errorMessage);
  }

  console.log(`onEdit 5 -> targetRow: ${targetRow} found by \`findMemberByEmail()\``);

  updateFeeInfo_(rangeEdited, sheetEditedName, targetRow, targetSheet);
  console.log(`onEdit 6 -> successfully completed trigger check`);
}


/**
 * Check whether the given range is within valid range of the given sheet.
 * 
 * Valid range includes all rows except header rows, and all columns from leftmost
 * until the "Internal Collected" column
 * 
 * @param {Range} range  Range from Event Object from `onEdit`.
 * @param {SpreadsheetApp.Sheet} sheet  Sheet where edit occurred.
 * @return {boolean}  True if range and sheet are valid, False otherwise
 */

function isLegalEdit_(range, sheet) {
  Logger.log("NOW ENTERING isLegalEdit()...");
  const sheetName = sheet.getName();
  const thisRow = range.getRow();
  const thisCol = range.getColumn();
  Logger.log(`isLegalEdit 1 -> sheetName: ${sheetName}`);

  // Function to get column mappings
  const colMap = GET_COL_MAP_(sheetName);
  const getLeftRange = () => (sheetName === MASTER_NAME) ? colMap.collector : colMap.feeStatus;

  const feeEditRange = {
    top: 2,    // Skip header row
    bottom: sheet.getLastRow(),
    rightmost: colMap.isInternalCollected,   // For both sheet, this is rightmost col
    leftmost : getLeftRange()   // Get leftmost according to sheet
  }

  Logger.log(`isLegalEdit 2 -> feeEditRange: ${JSON.stringify(feeEditRange)}`);

  // Exit if we're out of range
  if (thisRow < feeEditRange.top || thisRow > feeEditRange.bottom) {
    Logger.log("Row is out of bounds");
    return false;
  }
  if (thisCol < feeEditRange.leftmost || thisCol > feeEditRange.rightmost) {
    Logger.log("Column is out of bounds");
    return false;
  }

  Logger.log("Edit is within legal edit range");
  return true;
}


/** 
 * Update fee status from `sourceSheet` to `targetSheet`.
 * 
 * Includes handling the different ways of storing fee payment (boolean in semester sheet,
 * list of semesters in master sheet)
 * 
 * @param {Range} range  Range from Event Object from `onEdit`.
 * @param {string} sourceSheetName  Name of source sheet to extract fee info.
 * @param {number} targetRow  Target row to update.
 * @param {SpreadsheetApp.Sheet} targetSheet  Target sheet to update fee info.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Dec 16, 2024
 * @update  Dec 18, 2024
 * 
 */

function updateFeeInfo_(range, sourceSheetName, targetRow, targetSheet) {
  const thisCol = range.getColumn();
  const targetSheetName = targetSheet.getSheetName();

  console.log(`NOW ENTERING ${updateFeeInfo_.name}`);
  console.log(`Source: ${sourceSheetName}, thisCol: ${thisCol} && Target: ${targetSheetName}, targetRow: ${targetRow}`);

  // Find which column was edited in `sourceSheet` and find respective col in `targetSheet`
  const targetCol = getRespectiveCol(thisCol, sourceSheetName, targetSheetName);
  console.log(`updateFeeInfo 1 -> Successfully got respective range (row=${targetRow},col=${targetCol})!`);

  const targetRange = targetSheet.getRange(targetRow, targetCol);
  console.log(`updateFeeInfo 2 -> Successfully got targetRange for ${targetSheetName}`);

  // Special case: MASTER stores payment history as semesterCode(s).
  // If isPaid, then add semesterCode to payment history, i.e. bool -> str
  // Otherwise, nothing to modify in MASTER for member's payment history
  if (targetSheetName == MASTER_NAME && targetCol == MASTER_COLS.PAYMENT_HISTORY) {
    console.log("updateFeeInfo 3 -> entering if statement");
    const value = range.getValue() || "";
    const isPaid = parseBool_(value);    // convert to bool
    console.log(`updateFeeInfo 3b -> Value: ${value} isPaid: ${isPaid}`);

    // Only modify payment history if isPaid
    if (isPaid) {
      console.log("updateFeeInfo 3c -> entering isPaid");
      addPaidSemesterToMaster_(targetRow, sourceSheetName);
    }
    else {
      console.log("updateFeeInfo 3c -> entering NOT(isPaid)");
    }

  }
  else if (sourceSheetName == MASTER_NAME && thisCol == MASTER_COLS.PAYMENT_HISTORY) {
    // CASE 2: Add payment history from MASTER to SEMESTER_SHEET
    console.log("updateFeeInfo 3a ->  MASTER to SEMESTER_SHEET");
    const paymentHistory = range.getValue() || "";
    updateFeeStatusSemester_(paymentHistory, targetRow, targetCol, targetSheet);
  }
  else {
    // CASE 3: Copy fee details from SEMESTER_SHEET to MASTER
    console.log("updateFeeInfo 3b ->  SEMESTER_SHEET to MASTER");
    range.copyTo(targetRange, { contentsOnly: true });
  }

  console.log(`Completed update succesfully\nNOW EXITING ${updateFeeInfo_.name}`);


  /** HELPERS */
  function getRespectiveCol(sourceCol, sourceSheetName, targetSheetName) {
    const sourceCols = GET_COL_MAP_(sourceSheetName);   // Map of type of member data to its column index
    const targetCols = GET_COL_MAP_(targetSheetName);   // Get map of member data with respective column indices

    // Find respective column where `targetCol` contains same data as `sourceCol`.
    switch (sourceCol) {
      case (sourceCols.feeStatus): return targetCols.feeStatus;
      case (sourceCols.collectionDate): return targetCols.collectionDate;
      case (sourceCols.collector): return targetCols.collector;
      case (sourceCols.isInternalCollected): return targetCols.isInternalCollected;
    }
  }
}

/** 
 * Transfer new member registration from `Import` to semester sheet.
 * 
 * @trigger  New entry in `Import` sheet.
 * @param {Object} registration  Information on member registration.
 * @param {integer} [row = getLastSubmissionInMain()]  Gsheet row number to target, defaults to last row
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct 18, 2023
 * @update  May 17, 2025
 */

function copyFilloutRegToSemester_(registration, row = getLastSubmissionInSemester()) {
  const semesterSheet = SEMESTER_SHEET;
  const importMap = IMPORT_MAP;

  const registrationObj = JSON.parse(registration.replace(/[\n\r\t]/g, ' '));
  console.log(registrationObj);

  const startRow = row + 1;
  const colSize = semesterSheet.getLastColumn();

  const valuesByIndex = Array(colSize);

  // Format timestamp correctly; otherwise GSheet will not understand datetime
  const timestamp = registrationObj['timestamp'];
  if (timestamp != '') {
    const formattedTimestamp = Utilities.formatDate(
      new Date(timestamp),
      TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );

    registrationObj['timestamp'] = formattedTimestamp;   // replace with formatted
  }

  // Add additional information for payments occuring in other forms
  if (registrationObj['paymentAmount'] == 0) {
    const event = (registrationObj['referral'])['sources'] || "";
    const method = registrationObj['paymentMethod'] || 'Fee Waived';

    // Set new value
    registrationObj['paymentMethod'] = `${[event, method].join(': ')}`;
  }

  const removeRegex = /^[,\s-]+|[,\s-]+$/g;

  for (let [key, value] of Object.entries(registrationObj)) {
    if (key in importMap) {
      value = (typeof (value) === 'object') ? Object.values(value).join(', ') : value;
      let indexInMain = importMap[key] - 1;   // Set 1-index to 0-index
      valuesByIndex[indexInMain] = value.replace(removeRegex, ''); // Remove all unwanted
    }
  }

  // Set values of registration
  const rangeToImport = semesterSheet.getRange(startRow, 1, 1, colSize);
  rangeToImport.setValues([valuesByIndex]);

  return startRow;
}

/**
 * Creates an object containing member information for creating membership pass
 * 
 * Gets member's email, first and last name, member ID, membership status, and
 * membership expiration.
 * 
 * @param {number} row  Row of member to package information for
 * @return {Object}  Member information
 */
function packageMemberInfo_(row) {
  const sheet = SEMESTER_SHEET;
  const semesterName = SHEET_NAME;

  // Get member data to populate pass template
  const endCol = SEMESTER_COLS.MEMBER_ID; // From first name to id col
  const memberData = sheet.getSheetValues(row, 1, 1, endCol)[0];

  // Create function to access elements in arr as 0-index (GSheet is 1-indexed)
  const extractValues = (index) => memberData[index - 1].toString().trim();

  // Stringify fee status
  //const memberFeeStatus = parseBool(extractValues(SEMESTER_COLS.FEE_PAID)) ? 'Paid' : 'Unpaid';
  const memberFeeStatus = 'Paid';   // Set as paid since updating pass not implemented yet

  // Get membership expiration date via sheet name
  const semesterCode = getSemesterCode_(semesterName);  // Get the semester code based on the sheet name
  const membershipExpiration = getExpirationDate_(semesterCode);

  // Map member info to pass info
  return {
    email: extractValues(SEMESTER_COLS.EMAIL),
    firstName: extractValues(SEMESTER_COLS.FIRST_NAME),
    lastName: extractValues(SEMESTER_COLS.LAST_NAME),
    memberId: extractValues(SEMESTER_COLS.MEMBER_ID),
    memberStatus: 'Active',    // If email not found, then membership expired
    feeStatus: memberFeeStatus,
    expiry: membershipExpiration,
  }
}


/** 
 * Copies a member's registration from the `MASTER` sheet to the `MAIN_SHEET`.
 *
 * This function is a placeholder for copying a member's registration from the `MASTER`
 * sheet to the `MAIN_SHEET`. It is currently not implemented.
 * 
 * @example
 * // Copy a member's registration to the main sheet
 * copyToMainFromMaster(5);
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date
 */
function copyMasterRowToSemester(row) {
  throw new Error('Function not implemented.');
}


/**
 * Sends a notification for a new app submission.
 *
 * This function is a placeholder for sending a notification when a new app submission
 * is added. It is currently not implemented.
 *
 * @example
 * // Notify about a new app submission
 * notifyNewAppSubmission(5);
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date 
 */
function notifyNewAppSubmission(row) {
  throw new Error('Function not implemented.');
}