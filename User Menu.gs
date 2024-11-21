/**
 * Users authorized to use the McRUN menu.
 * 
 * Prevents unwanted data overwrite in Gsheet.
 * 
 */
const PERM_USER_ = [
  'mcrunningclub@ssmu.ca', 
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'gagnonjikael@gmail.com',
  'thecharlesvillegas@gmail.com',
];

/**
 * Creates custom menu to run frequently used scripts in Google App Script.
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

  ui.createMenu('McRUN Menu')

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
 * Boiler plate function to display custom UI to run scripts.
 * 
 * @trigger User choice in custom menu
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
  const message = `⚙️ Now executing ${functionName} in ${sheetName}. Press cancel to stop.`;

  const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);

  if(response == ui.Button.OK) {
    this[functionName]();   // executing function with name `functionName`;
  }
  else {
    ui.alert('Execution cancelled...');
  }
}

/** 
 *  Scripts for MAIN_SHEET submenu.
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
 *  Scripts for MASTER_SHEET submenu.
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

