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
 * @trigger User choice in custom menu.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 21, 2024
 */

function logMenuAttempt_() {
  const userEmail = getCurrentUserEmail_().toString();
  Logger.log(`McRUN menu access attempt by: ${userEmail}`);
}

/**
 * Activate the sheet `sheetName` in Google Spreadsheet.
 * 
 * Changes view to `sheetName`.
 * 
 * @input {string}  sheetName  Name of target sheet.
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
 * @trigger Open Google Spreadsheet.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 21, 2024
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = getCurrentUserEmail_().toString();

  // Check if authorized user to prevent illegal execution
  if (!PERM_USER_.includes(userEmail)) return;

  ui.createMenu('🏃‍♂️ McRUN Menu')
    .addItem('📢 Custom menu. Click for help.', 'helpUI_')
    .addSeparator()

    .addSubMenu(ui.createMenu('Main Scripts')
      .addItem('Sort by Name', 'sortByNameUI_')
      .addItem('Submit Form', 'onFormSubmitUI_')
      .addItem('Format Sheet', 'formatSpecificColumnsUI_')
      .addItem('Create ID for Last Member', 'encodeLastRowUI_')
      )
    .addSeparator()

    .addSubMenu(ui.createMenu('Master Scripts')
      .addItem('Create Master', 'createMasterUI_')
      .addItem('Add Last Submission from Main', 'addLastSubmissionToMasterUI_')
      .addItem('Sort by Email', 'sortMasterByEmailUI_')
      .addItem('Process Last Submission', 'processLastSubmissionUI_')
    )
    .addToUi();
}


/**
 * Displays a help message for the custom McRUN menu.
 * 
 * Accessible only to authorized members.
 */

function helpUI_() {
  const ui = SpreadsheetApp.getUi();
  
  const helpMessage = `
    📋 McRUN Menu Help

    This menu is only accessible to selected members.

    - Scripts are applied to the sheet via the submenu.
    - To view or modify authorized users, open the Google Apps Script editor.

    Please contact the admin if you need access or assistance.
  `;

  // Display the help message
  ui.alert("McRUN Menu Help", helpMessage.trim(), ui.ButtonSet.OK);
}


/**
 * Boiler plate function to display custom UI to run scripts.
 * 
 * @trigger User choice in custom menu.
 * 
 * @input {string}  functionName  Name of function to execute.
 * @input {string}  sheetName  Name of sheet where `functionName` will run.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Nov 21, 2024
 * @update  Nov 21, 2024
 */

function confirmAndRunUserChoice_(functionName, sheetName) {
  const ui = SpreadsheetApp.getUi();
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
  logMenuAttempt_();  
}


/** 
 *  Scripts for `MAIN_SHEET` menu items.
 */

function sortByNameUI_() {
  const functionName = 'sortMainByName';
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function onFormSubmitUI_() {
  const functionName = 'onFormSubmit';
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function formatSpecificColumnsUI_() {
  const functionName = 'formatSpecificColumns';
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function encodeLastRowUI_() {
  const functionName = 'encodeLastRow';
  const sheetName = SHEET_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}


/** 
 *  Scripts for `MASTER_SHEET` menu items.
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
    ui.alert("Execution cancelled...");
  }
  
  logMenuAttempt_();    // log attempt
}

function addLastSubmissionToMasterUI_() {
  const functionName = 'addLastSubmissionToMaster';
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function sortMasterByEmailUI_() {
  const functionName = 'sortMasterByEmail';
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

function processLastSubmissionUI_() {
  const functionName = 'processLastSubmission';
  const sheetName = MASTER_NAME;
  confirmAndRunUserChoice_(functionName, sheetName);
}

