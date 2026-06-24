/**
 * Logs user attempting to use custom McRUN menu.
 * 
 * If input empty, then extract email using `getCurrentUserEmail_()`.
 * 
 * @trigger User choice in custom menu.
 * 
 * @param {string} [email=""]  Email of active user. 
 *                             Defaults to empty string.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 22, 2024
 */

function logMenuAttempt_(email = "") {
  const userEmail = email.size > 0 ? email : getCurrentUserEmail_();
  Logger.log(`McRUN menu access attempt by: ${userEmail}`);
}

/**
 * Activate the sheet `sheetName` in Google Spreadsheet.
 * 
 * Changes view to `sheetName`.
 * 
 * @param {string} sheetName  Name of target sheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 21, 2024
 */

function changeSheetView_(sheetName) {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
}


/**
 * Creates custom menu to run frequently used scripts in Google App Script.
 * 
 * Extracting function name using `name` property to allow for refactoring.
 * 
 * Cannot check if user authorized, or custom menu will not be displayed.
 * This is due to Google App Script limitation.
 * 
 * @trigger Open Google Spreadsheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Mar 1, 2025
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏃‍♂️ McRUN Menu')
    .addItem('📢 Custom menu. Click for help.', helpUI_.name)
    .addSeparator()
    .addItem('Encode Text from Input', createMemberIDFromInputUI_.name)
    .addSeparator()

    .addSubMenu(ui.createMenu('Main Scripts')
      .addItem('Sort by Name', sortByNameUI_.name)
      .addItem('Submit Form', onFormSubmitUI_.name)
      .addItem('Prettify Main Sheet', prettifyMainUI_.name)
      .addItem('Find Waiver Url', findWaiverLinkUI_.name)
      .addItem('Create ID for Last Member', encodeLastRowUI_.name)

    )

    .addSubMenu(ui.createMenu('Master Scripts')
      .addItem('Create Master', createMasterUI_.name)
      .addItem('Sort by Email', sortMasterByEmailUI_.name)
      .addItem('Prettify Master Sheet', prettifyMasterUI_.name)
      .addItem('Add Last Submission from Main', addLastSubmissionToMasterUI_.name)
      .addItem('Add Specific Sheet Submission (draft)', addMemberFromSheetInRowUI_.name)
    )
    .addToUi()
    ;
}


/**
 * Displays a help message for the custom McRUN menu.
 * 
 * Accessible to all users.
 */

function helpUI_() {
  const ui = SpreadsheetApp.getUi();

  const helpMessage = `
    📋 McRUN Menu Help

    - This menu is only accessible to authorized members.

    - Scripts are applied to the sheet via the submenu.

    - Please contact the admin if you need access or assistance.
  `;

  // Display the help message
  ui.alert("McRUN Menu Help", helpMessage.trim(), ui.ButtonSet.OK);
}


/**
 * Boiler plate function to display custom UI to run scripts.
 * 
 * Verifies if user is authorized before executing script.
 * 
 * @trigger User choice in custom menu.
 * 
 * @param {string} functionName  Name of function to execute.
 * @param {string} sheetName  Name of sheet where `functionName` will run.
 * @param {string}  [additionalMsg=""]  Custom message for executing function.
 *                                      Defaults to empty string.
 * @param {string}  [funcArg=""]  Function argument to pass with `functionName`.
 *                                Defaults to empty string.
 * 
 * @return {string}  Return value of the executed function.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Mar 1, 2025
 */

function confirmAndRunUserChoice_(functionName, sheetName, additionalMsg = "", funcArg = "") {
  const ui = SpreadsheetApp.getUi();
  const userEmail = getCurrentUserEmail_();

  // Check if authorized user to prevent illegal executi on
  if (!ADMINS_.includes(userEmail)) {
    const warningMsgHeader = "🛑 You are not authorized 🛑"
    const warningMsgBody = "Please contact the exec team if you believe this is an error.";

    ui.alert(warningMsgHeader, warningMsgBody, ui.ButtonSet.OK);
    return;
  }

  // Continue execution if user is authorized
  let message = `
    ⚙️ Now executing ${functionName} in ${sheetName}.
  
    🚨 Press cancel to stop.
  `;

  // Append additional message if non-empty
  message += additionalMsg ? `\n🔔 ${additionalMsg}` : "";

  // Save user response
  const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);
  let retValue = "";

  if (response == ui.Button.OK) {
    // Execute function `functionName` (with arg if non-empty)
    retValue = funcArg ? this[functionName](funcArg) : this[functionName]();
  }
  else {
    ui.alert('Execution cancelled...');
  }

  // Change view to target sheet
  changeSheetView_(sheetName);

  // Log attempt in console using active user email
  logMenuAttempt_(userEmail);

  // Return value from executed function if required
  return retValue;
}


/**
 * Returns true if row is int and found in given sheet.
 *
 * Helper function for UI functions for McRUN menu.
 *
 * @param {integer}  The row number in `sheet` 1-indexed.
 * @param {SpreadsheetApp.Sheet} sheet  The sheet to search in.
 * 
 * @return {boolean}  The input is a number.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * @date  March 1, 2025
 * @update  March 1, 2025
 */

function isValidRow_(row, sheet = SEMESTER_SHEET) {
  const lastRow = sheet.getLastRow();
  const rowInt = parseInt(row);

  return (Number.isInteger(rowInt) && rowInt >= 2 && rowInt <= lastRow);
}


/** 
 * Scripts for semester sheet menu items: sort by name
 */
function sortByNameUI_() {
  const functionName = sortSemesterByName.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for semester sheet menu items: run submit form function
 */
function onFormSubmitUI_() {
  const functionName = onFormSubmit.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for semester sheet menu items: format sheet
 */
function prettifyMainUI_() {
  const functionName = formatSemester.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for semester sheet menu items: encode last row
 */
function encodeLastRowUI_() {
  const functionName = encodeRowSemester_.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for semester sheet menu items: ask for row and find waiver for
 * that member
 */
function findWaiverLinkUI_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Which row do you want to search for a waiver link?";
  const textMsg = "Enter the row number, or leave empty for the last row.";

  const response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK);
  const responseText = response.getResponseText().trim();

  const sheetName = SHEET_NAME;
  const functionName = findWaiverLink_.name;

  let customMsg = "";
  const rowNumber = Number.parseInt(responseText);

  if (responseText === "") {
    // User did not enter a row number; check last row only
    customMsg = "This function will only run in the last row.";
  }
  else if (isValidRow_(rowNumber)) {
    // Row is valid, can continue execution
    customMsg = `This function will run in row ${rowNumber}.`;
  }
  else {
    // Input value is invalid row. Stop execution.
    ui.alert("Incorrect row number, please try again with a valid row number.");
    return;
  }

  // Run respective function depending if-statement above
  confirmAndRunUserChoice_(functionName, sheetName, customMsg, rowNumber);
}

/** 
 * Scripts for semester sheet menu items: ask for text and encode it
 */
function createMemberIDFromInputUI_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Enter the text to encode";
  const textMsg = "";

  var response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK_CANCEL);
  const responseText = response.getResponseText().trim();
  const responseButton = response.getSelectedButton();

  // Process the user's response.
  if (responseText === "") {
    ui.alert("INVALID INPUT", "Please enter a non-empty string", ui.ButtonSet.OK);
  }
  else if (responseButton == ui.Button.OK) {
    // User clicked "OK" and response non-empty.
    const encoded = encodeFromInput_(responseText);
    ui.alert("Here is the encoded text:", encoded, ui.ButtonSet.OK);
  }
  else {
    // User clicked "Canceled" or X in the title bar.
    ui.alert('Execution cancelled...');
  }

  logMenuAttempt_();    // log attempt
}


/** 
 * Scripts for master sheet menu items: overwrite master sheet with a new one
 */
function createMasterUI_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Do you want to consolidate member registrations?";
  const textMsg = "This action will overwrite present data in MASTER. Ensure that data has been copied beforehand.";

  var choice = ui.alert(headerMsg, textMsg, ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (choice == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert("Confirmation received. Starting data consolidation...");
    //consolidateMemberData();
    ui.alert('Currently disabled to prevent overwrite');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Execution cancelled...');
  }

  logMenuAttempt_();    // log attempt
}

/** 
 * Scripts for master sheet menu items: format sheet
 */
function prettifyMasterUI_() {
  const functionName = formatMaster_.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for master sheet menu items: add last submission from semester sheet
 */
function addLastSubmissionToMasterUI_() {
  const functionName = addLastSubmissionToMaster_.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/** 
 * Scripts for master sheet menu items: sort by email
 */
function sortMasterByEmailUI_() {
  const functionName = sortMasterByEmail.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

/**
 * Adds a new member from a specific sheet submission to the `MASTER` sheet.
 *
 * This function is a placeholder for adding a member from a specific sheet submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date
 */
function addMemberFromSheetInRowUI_() { 
  throw new Error('Function not implemented.');
}
