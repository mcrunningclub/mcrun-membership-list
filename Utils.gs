/**
 * Removes diacritics (accents) from a string.
 * 
 * This function normalizes the input string and removes any diacritical marks,
 * ensuring a clean, ASCII-compatible output.
 * 
 * @param {string} str  The string to normalize and strip of diacritics.
 * @return {string}  The normalized string without diacritics.
 * 
 * @example
 * const result = removeDiacritics("José");
 * console.log(result); // Outputs: "Jose"
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 5, 2025
 * @update  Mar 15, 2025
 */

function removeDiacritics_(str) {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Get semester code from semester sheet name in map, or creates if not found.
 * 
 * First letter of code is W/F/S corresponding to first letter of semester
 * and next two are digits YY corresponding to the year.
 * 
 * @param {string} semester  Semester name e.g. Fall 2024
 * @return {string}  Semester code e.g. F24
 */
function getSemesterCode_(semester) {
  // Return semester code if already in map
  if (semester in SEMESTER_CODE_MAP) {
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

/**
 * Find row index of last submission, starting from bottom using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Sep 1, 2024
 * @update  Dec 18, 2024
 */

function getLastSubmissionInSemester() {
  const sheet = SEMESTER_SHEET;
  let lastRow = sheet.getLastRow();

  while (sheet.getRange(lastRow, REGISTRATION_DATE_COL).getValue() == "") {
    lastRow = lastRow - 1;
  }

  return lastRow;
}


/**
 * Returns timezone for currently running script.
 * 
 * Prevents incorrect time formatting during time changes like Daylight Savings Time.
 *
 * @return {string}  Timezone as geographical location (e.g.`'America/Montreal'`).
 */

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}


/**
 * Returns email of current user executing Google Apps Script functions.
 * 
 * Prevents incorrect account executing Google automations (e.g. McRUN bot.)
 * 
 * @return {string}  Email of current user.
 */

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}


/**
 * Converts a string to a boolean value.
 * 
 * @param {string} val  A string that contains a boolean.
 * @return {Boolean}  Parsed value.
 */

function parseBool_(val) {
  return val === true || val === "true";
}