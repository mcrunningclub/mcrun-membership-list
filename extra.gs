/**
 * ## Sample text
 * sample *sample* **sample**
 *
 * ## Sample hyperlink
 * [Sample link](https://tanaikech.github.io/)
 *
 * ## Sample image
 * ![sample image](https://stackoverflow.design/assets/img/logos/se/se-icon.png)
 *
 * ### Sample script
 * ```javascript
 * const text = "sample";
 * const res = myFunction(text);
 * ```
 *
 * ## Sample json data
 * ```json
 * {
 *   "key1": "value1",
 *   "key2": "value3",
 * }
 * ```
 * @example
 * // returns 2
 * sample();
 *
 * @param {String} text Sample text.
 * @return {String} Output text.
 */
function sample() {
}


// CURRENT TEMPLATE
/**
 * DESCRIPTION
 *
 * @return {string[]} Returns
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Oct , 2024
 * 
 * ```javascript
 * // Sample Script : Storing processed submission.
 * const processedData = processLastSubmission();
 * ```
 */


function getFunctionNames_() {
  const allKeys = Object.keys(this); // Get all properties in the global scope
  const functionNames = allKeys.filter(key => typeof this[key] === "function"); // Filter out only functions
  return functionNames;
}

// Test
function printFunctions() {
  Logger.log(getFunctionNames_());
}

function extractFeeStatus() {
  const rangeFeeStatus = sheet.getRange(lastRow, MASTER_FEE_STATUS);
  const feeStatus = rangeFeeStatus.getValue();

  // Create a regular expression from `FEE_STATUS_ENUM` array
  const regex = new RegExp(FEE_STATUS_ENUM.join("|"), "i"); // Case insensitive
  
  // Find and replace match in `feeStatus`
  const match = feeStatus.match(regex)[0];
  rangeFeeStatus.setValue(match);
}


/* DEPRICATED OR JUNK FUNCTIONS */
function drafts_() {
  return;

  /**
   * @author: Andrey S Gonzalez
   * @date: Oct 17, 2023
   * @update: Oct 17, 2023
   * 
   * Check if PassKit URL is entered.
   */

  function isPasskitURL_() {
    return;
    const sheet = MAIN_SHEET;
    const newSubmissionRow = sheet.getLastRow();
    const urlCol = sheet.getLastColumn();

    const rangeURL = sheet.getRange(newSubmissionRow, urlCol);
    return rangeURL.isBlank;
  }


  /**
   * @author: Andrey S Gonzalez
   * @date: Oct 1, 2023
   * @update: Oct 8, 2023
   * 
   * Verifies if Collection Date and Collection Person were added before checking `Fee Paid` box
   * If not, displays up warning message
   */ 

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
   * @date: Oct 17, 2023
   * @update: Oct 17, 2023
   * 
   * Check if PassKit URL is entered.
   */

  function isPasskitURL_() {
    return;
    const sheet = MAIN_SHEET;
    const newSubmissionRow = sheet.getLastRow();
    const urlCol = sheet.getLastColumn();

    const rangeURL = sheet.getRange(newSubmissionRow, urlCol);
    return rangeURL.isBlank;
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

  
  // Returns the letter representation of the column index
  // e.g. "1" returns "A"; "5" returns "E"
  function getLetterFromColumnIndex_(column) {
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
  function getColumnFromLetter_(letter) {
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++) {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }

  function getAllSheetNames_() {
    var out = new Array()
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i=0 ; i<sheets.length ; i++) 
      out.push( [ sheets[i].getName() ] )
    return out
  }

  function getSheetRangeFromA1(a1Notation, sheet) {
    // Helper function to convert column letters to numeric index
    function columnToIndex(column) {
      let index = 0;
      for (let i = 0; i < column.length; i++) {
        index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
      }
      return index;
    }

    // Parse A1 notation
    const rangeMatch = a1Notation.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!rangeMatch) return null; // Default for invalid input

    const [_, col1, row1, col2, row2] = rangeMatch;

    // Convert start and end columns/rows
    const startColumn = columnToIndex(col1);
    const startRow = parseInt(row1, 10);

    const endColumn = col2 ? columnToIndex(col2) : startColumn; // If no range, end = start
    const endRow = row2 ? parseInt(row2, 10) : startRow;

    // Calculate number of rows and columns
    const numRows = endRow - startRow + 1;
    const numCols = endColumn - startColumn + 1;

    return sheet.getRange(startRow, startColumn, numRows, numCols);
  }
}