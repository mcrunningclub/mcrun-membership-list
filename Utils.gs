/**
 * Retrieves the column mapping for a given sheet.
 *
 * This function returns the column mapping object for the specified sheet name.
 * If the sheet name is not found in the mapping, it returns `null`.
 *
 * @param {string} sheet - The name of the sheet to retrieve the column mapping for.
 * @return {Object|null} The column mapping object for the sheet, or `null` if not found
 * 
 * @author  Andrey Gonzalez
 * @date  May 24, 2025
 * @update  Sep 26, 2025
 */

let SHEET_COL_MAP = null;

function GET_COL_MAP_(sheet) {
  /** If SHEET_COL_MAP not defined yet */
  if (!SHEET_COL_MAP) {
    SHEET_COL_MAP = {
      [SHEET_NAME]: {
        emailCol: EMAIL_COL,
        memberIdCol: MEMBER_ID_COL,
        feeStatus: IS_FEE_PAID_COL,   // Boolean value
        collectionDate: COLLECTION_DATE_COL,
        collector: COLLECTION_PERSON_COL,
        isInternalCollected: IS_INTERNAL_COLLECTED_COL,
      },
      [MASTER_NAME]: {
        emailCol: MASTER_EMAIL_COL,
        memberIdCol: MASTER_MEMBER_ID_COL,
        feeStatus: MASTER_PAYMENT_HIST,   // String with semester code(s)
        collectionDate: MASTER_COLLECTION_DATE,
        collector: MASTER_FEE_COLLECTOR,
        isInternalCollected: MASTER_IS_INTERNAL_COLLECTED,
      },
    };
  }

  return SHEET_COL_MAP[sheet] || null;
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