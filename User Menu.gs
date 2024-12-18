/**
 * Users authorized to use the McRUN menu.
 * 
 * Prevents unwanted data overwrite in Gsheet.
 */
const PERM_USER_ = [ 
  'mcrunningclub@ssmu.ca',
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'gagnonjikael@gmail.com',
  'thecharlesvillegas@gmail.com',
];


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

function logMenuAttempt_(email="") {
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
 * Cannot check if user authorized, or custom menu will not be displayed
 * due to Google App Script limitation.
 * 
 * @trigger Open Google Spreadsheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 22, 2024
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏃‍♂️ McRUN Menu')
    .addItem('📢 Custom menu. Click for help.', helpUI_.name)
    .addSeparator()
    .addItem('Turn ON/OFF onEdit()', setOnEditFlagUI_.name)

    .addSubMenu(ui.createMenu('Main Scripts')
      .addItem('Sort by Name', sortByNameUI_.name)
      .addItem('Submit Form', onFormSubmitUI_.name)
      .addItem('Prettify Main Sheet', prettifyMainUI_.name)
      .addItem('Encode Text from Input', createMemberIDFromInputUI_.name)
      .addItem('Create ID for Last Member', encodeLastRowUI_.name)
      )

    .addSubMenu(ui.createMenu('Master Scripts')
      .addItem('Create Master', createMasterUI_.name)
      .addItem('Sort by Email', sortMasterByEmailUI_.name)
      .addItem('Prettify Master Sheet', prettifyMasterUI_.name )
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
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 21, 2024
 */

function confirmAndRunUserChoice_(functionName, sheetName) {
  const ui = SpreadsheetApp.getUi();
  const userEmail = getCurrentUserEmail_();

  // Check if authorized user to prevent illegal execution
  if (!PERM_USER_.includes(userEmail)) {
    const warningMsgHeader = "🛑 You are not authorized 🛑"
    const warningMsgBody = "Please contact the exec team if you believe this is an error.";
    
    ui.alert(warningMsgHeader, warningMsgBody, ui.ButtonSet.OK);
    return;
  }
  
  // Continue execution if user is authorized
  const message = `
    ⚙️ Now executing ${functionName} in ${sheetName}.
  
    🚨 Press cancel to stop.
  `;

  const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);

  if(response == ui.Button.OK) {
    this[functionName]();   // executing function with name `functionName`
  }
  else {
    ui.alert('Execution cancelled...');
  }

  // Change view to target sheet
  changeSheetView_(sheetName);

  // Log attempt in console using active user email
  logMenuAttempt_(userEmail);
}


/** 
 * Scripts for `MAIN_SHEET` menu items.
 * 
 * Extracting function name using `name` property to allow for refactoring.
 */

function sortByNameUI_() {
  const functionName = sortMainByName.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function onFormSubmitUI_() {
  const functionName = onFormSubmit.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function prettifyMainUI_() {
  const functionName = formatMainView.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function encodeLastRowUI_() {
  const functionName = encodeLastRow.name;
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function createMemberIDFromInputUI_() {
  const ui = SpreadsheetApp.getUi();
  const headerMsg = "Enter the text to encode";
  const textMsg = "";

  var response = ui.prompt(headerMsg, textMsg, ui.ButtonSet.OK_CANCEL);
  const responseText = response.getResponseText().trim();
  const responseButton = response.getSelectedButton();

  // Process the user's response.
  if(responseText === "") {
    ui.alert("INVALID INPUT", "Please enter a non-empty string", ui.ButtonSet.OK);
  }
  else if(responseButton == ui.Button.OK){
    // User clicked "OK" and response non-empty.
    const encoded = encodeFromInput(responseText);
    ui.alert("Here is the encoded text:", encoded, ui.ButtonSet.OK);
  }  
  else {
    // User clicked "Canceled" or X in the title bar.
    ui.alert('Execution cancelled...');
  }
  
  logMenuAttempt_();    // log attempt
}


/** 
 * Scripts for `MASTER_SHEET` menu items.
 * 
 * Extracting function name using `name` property to allow for refactoring.
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

function prettifyMasterUI_() {
  const functionName = formatMasterView.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function addLastSubmissionToMasterUI_() {
  const functionName = addLastSubmissionToMaster.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function sortMasterByEmailUI_() {
  const functionName = sortMasterByEmail.name;
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function addMemberFromSheetInRowUI_() {
  const functionName = addMemberFromSheetInRow.name;
}

