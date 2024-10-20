/* SHEET CONSTANTS */
const MASTER_NAME = 'MASTER';
const MASTER_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_NAME);

const SEMESTER_CODE_MAP = new Map();
const ALL_SEMESTERS = ['Fall 2024', 'Summer 2024', 'Winter 2024'];


// Index of processed semester data arrays (0-indexed)
const PROCESSED_ARR = {
  LAST_REGISTRATION: 0,
  EMAIL: 1,
  FIRST_NAME: 2,
  LAST_NAME: 3,
  PREFERRED_NAME: 4,
  YEAR: 5,
  PROGRAM: 6,
  MEMBER_DESCR: 7,
  REFERRAL: 8,
  WAIVER: 9,
  EMPTY: 12,
  IS_FEE_PAID: 13,
  COLLECTION_DATE: 14,
  COLLECTED_BY: 15,
  GIVEN_TO_INTERNAL: 16,
  COMMENTS: 17,
  ATTENDANCE_STATUS: 18,
  MEMBER_ID: 19,
  LAST_REG_CODE: 20,
  REGISTATION_HIST: 21
};

function createMaster() {
  consolidateMemberData();
}

// Helper Function
function processSemesterData(sheetName) {
  const semesterSheetRange = 'A2:T';
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = SPREADSHEET.getSheetByName(sheetName).getRange(semesterSheetRange).getValues();
  var semesterCode = getSemesterCode(sheetName); // Get the semester code based on the sheet name
  
  const processedData = sheetData.map(function(row) {
    // Append semester code if entries are non-empty
    var index;

    index = PROCESSED_ARR.MEMBER_DESCR;
    row[index] = row[index] ? "(" + semesterCode + ") " + row[index] : "";

    index = PROCESSED_ARR.REFERRAL;
    row[index] = row[index] ? "(" + semesterCode + ") " + row[index] : "";

    index = PROCESSED_ARR.COMMENTS;
    row[index] = row[index] ? "(" + semesterCode + ") " + row[index] : "";

    // Append semester code to payment history
    index = PROCESSED_ARR.IS_FEE_PAID;
    row[index] = row[index] ? semesterCode : "";
    
    // Append row with semester code for MASTER.LAST_REG_CODE
    row.push(semesterCode);

    // Append row with semester code for MASTER.REGISTRATION_HIST
    row.push("");
    
    return row;
  });
  
  return processedData;
}


function consolidateMemberData() {
  // Get processed semester data
  var fall2024 = processSemesterData('Fall 2024');
  var summer2024 = processSemesterData('Summer 2024');
  var winter2024 = processSemesterData('Winter 2024');

  // Combine all semester data
  var allMemberData = fall2024.concat(summer2024, winter2024);

  // Filter out empty rows
  allMemberData = allMemberData.filter(function(row) {
    return row[0] !== "" && row[1] !== "";
  });
    
  // Create an object to store the unique entries by email address
  const memberMap = {};

  allMemberData.forEach(function(row) {
    var email = row[1];

    if (!memberMap[email]) {
      // Initialize the entry for the email if it doesn't exist
      memberMap[email] = row;
    } else {
      // Concatenate the relevant columns if the email already added
      // Access memberMap is 0-indexed according to semester sheet
      var index, regHistory, semesterCode;
      
      index = PROCESSED_ARR.MEMBER_DESCR;
      if(row[index]) memberMap[email][index] += "\n" + row[index];
      
      index = PROCESSED_ARR.REFERRAL;
      if(row[index]) memberMap[email][index] += "\n" + row[index];

      index = PROCESSED_ARR.COMMENTS;
      if(row[index]) memberMap[email][index] += "\n" + row[index];
      
      // Append registration history using semester code
      index = PROCESSED_ARR.REGISTATION_HIST;
      semesterCode = row[PROCESSED_ARR.LAST_REG_CODE];
      regHistory = memberMap[email][index] ? " " + semesterCode : semesterCode;
      memberMap[email][index] += regHistory;

      // Append payment history
      index = PROCESSED_ARR.IS_FEE_PAID;
      if(row[index]) memberMap[email][index] += " " + row[index];
    }
  });

  // Convert the emailMap object back into an array
  var resultData = Object.values(memberMap);

  // Sort by 'First Name' column
  resultData.sort(function(a, b) {
    var firstNameA = a[2].toLowerCase(); // First Name in lowercase for consistent sorting
    var firstNameB = b[2].toLowerCase();
    return firstNameA.localeCompare(firstNameB);
  });


  // Select specific columns
  var selectedData = resultData.map(function(row) {
    return [ 
    row[PROCESSED_ARR.EMAIL],
    row[PROCESSED_ARR.FIRST_NAME],
    row[PROCESSED_ARR.LAST_NAME],
    row[PROCESSED_ARR.PREFERRED_NAME],
    row[PROCESSED_ARR.YEAR],
    row[PROCESSED_ARR.PROGRAM], 
    row[PROCESSED_ARR.WAIVER],
    row[PROCESSED_ARR.MEMBER_DESCR],
    row[PROCESSED_ARR.REFERRAL],
    row[PROCESSED_ARR.LAST_REGISTRATION],
    row[PROCESSED_ARR.LAST_REG_CODE],
    row[PROCESSED_ARR.REGISTATION_HIST],
    row[PROCESSED_ARR.EMPTY],
    row[PROCESSED_ARR.EMPTY],
    row[PROCESSED_ARR.EMPTY],
    row[PROCESSED_ARR.COLLECTED_BY],
    row[PROCESSED_ARR.COLLECTION_DATE],
    row[PROCESSED_ARR.GIVEN_TO_INTERNAL], 
    row[PROCESSED_ARR.IS_FEE_PAID],
    row[PROCESSED_ARR.COMMENTS],
    row[PROCESSED_ARR.EMPTY],
    row[PROCESSED_ARR.MEMBER_ID]
    ]
  });
  
  // Output sorted unique data to another sheet or range
  MASTER_SHEET.getRange(2, 1, selectedData.length, selectedData[0].length).setValues(selectedData);
}




// Helper Function
function getSemesterCode(semester) {
  // Return semester code if already in map
  if(semester in SEMESTER_CODE_MAP) {
    return SEMESTER_CODE_MAP.get(semester)
  }

  // Extract the first letter of the string and last two char representing the year
  const semesterType = semester.charAt(0);
  const semesterYear = semester.slice(-2);
  const code = semesterType + semesterYear;

  // Same key-value in map
  SEMESTER_CODE_MAP.set(semester, code);

  return code;
}



function sortUniqueData() {
  // Get the active sheet and data from each range
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var fall2024 = sheet.getSheetByName('Fall 2024').getRange('A2:U').getValues();
  var summer2024 = sheet.getSheetByName('Summer 2024').getRange('A2:U').getValues();
  var winter2024 = sheet.getSheetByName('Winter 2024').getRange('A2:U').getValues();

  // Append the sheet name to each row
  fall2024 = fall2024.map(function(row) {
    return row.concat('Fall 2024'); // Append the sheet name 'Fall 2024'
  });

  summer2024 = summer2024.map(function(row) {
  return row.concat('Summer 2024'); // Append the sheet name 'Summer 2024'
  });

  winter2024 = winter2024.map(function(row) {
    return row.concat('Winter 2024'); // Append the sheet name 'Winter 2024'
  });

  
  // Combine data from all semesters
  var allData = fall2024.concat(summer2024, winter2024);
  
  // Filter out empty rows
  allData = allData.filter(function(row) {
    return row[0] !== "" && row[1] !== "";
  });
  
  // Remove duplicates
  var uniqueData = allData.filter((row, index, self) =>
    index === self.findIndex((r) => r[0] === row[0] && r[1] === row[1])
  );
  
  // Sort by the second column (ignoring accents)
  uniqueData.sort(function(a, b) {
    var nameA = a[1].normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove accents from nameA
    var nameB = b[1].normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove accents from nameB
    return nameA.localeCompare(nameB);
  });

  // Select specific columns: 4, 5, 6, 7, 10, 22, 1 (arrays are 0-indexed, so Col4 is index 3, Col1 is index 0, etc.)
  var selectedData = uniqueData.map(function(row) {
    return [ row[1], row[2], row[3], row[4], row[5], row[6], row[9], row[0], row[21] ];
  });
  
  // Output sorted unique data to another sheet or range
  MASTER_SHEET.getRange(2, 1, selectedData.length, selectedData[0].length).setValues(selectedData);

  //MASTER_SHEET.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);

}

// Create Member ID using email
function encodeMasterList() {
  var sheet = MASTER_SHEET;
  const MASTER_EMAIL_COL = 1;
  const MASTER_MEMBER_ID_COL = 19;
  var i, email;

  // Start at row 2 (1-indexed)
  for (i = 2; i < sheet.getMaxRows(); i++) {
    email = sheet.getRange(i, MASTER_EMAIL_COL).getValue();
    if (email === "") return;   // check for invalid row

    var member_id = MD5(email);
    sheet.getRange(i, MASTER_MEMBER_ID_COL).setValue(member_id);
  }
}

function queryData() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  var a2Value = MASTER_SHEET.getRange("A2").getValue(); // Get the value from A2
  if (!a2Value) {
    return; // Exit the function if A2 is empty
  }

  // Get data from the four sheets
  var fall2024 = sheet.getSheetByName('Fall 2024').getRange('A2:V').getValues();
  var summer2024 = sheet.getSheetByName('Summer 2024').getRange('A2:V').getValues();
  var winter2024 = sheet.getSheetByName('Winter 2024').getRange('A2:V').getValues();

  // Combine data from all semesters
  var allData = fall2024.concat(summer2024, winter2024);

  // Sort by the first column (index 0)
  allData.sort(function(a, b) {
    return a[0] > b[0] ? 1 : -1;
  });

  // Filter data where the second column (index 1) contains the value in A2
  var filteredData = allData.filter(function(row) {
    return row[1] && row[1].toString().indexOf(a2Value) !== -1;
  });

  // If there are no matches, exit
  if (filteredData.length === 0) {
    return;
  }

  // Sort the filtered data by the first column (index 0) in descending order
  filteredData.sort(function(a, b) {
    return a[0] < b[0] ? 1 : -1;
  });

  // Select specific columns: 4, 5, 6, 7, 10, 22, 1 (arrays are 0-indexed, so Col4 is index 3, Col1 is index 0, etc.)
  var selectedColumns = filteredData.map(function(row) {
    return [row[3], row[4], row[5], row[6], row[9], row[21], row[0]];
  });

  // Limit the results to 1 (since the formula does that)
  var result = selectedColumns.slice(0, 1);
  

  // Output the result to a specified location (adjust this as needed)
  var outputRange = MASTER_SHEET.getRange(2, 9, result.length, result[0].length); // Outputs to I2 (change as needed)
  outputRange.setValues(result);
}

