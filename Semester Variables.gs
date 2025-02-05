// SHEET CONSTANTS
const SHEET_NAME = 'Winter 2025';       // MUST UPDATE EVERY SEMESTER!
const MAIN_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

// IMPORT SHEET FOR FILLOUT REGISTRATIONS
const IMPORT_NAME = 'Import';
const IMPORT_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_NAME);
const IMPORT_SHEET_ID = 973209987;

const TIMEZONE = getUserTimeZone_();

// LIST OF COLUMNS IN SHEET_NAME
// Please update any changes made in active sheet
const REGISTRATION_DATE_COL = 1;      // Column 'A'
const EMAIL_COL = 2;                  // Column 'B'
const FIRST_NAME_COL = 3;             // Column 'C'
const LAST_NAME_COL = 4;              // Column 'D'
const PREFERRED_NAME_COL = 5;         // Column 'E'
const YEAR_COL = 6;                   // Column 'F'
const PROGRAM_COL = 7;                // Column 'G'
const DESCRIPTION_COL = 8;            // Column 'H'
const REFERRAL_COL = 9;               // Column 'I'
const WAIVER_COL = 10;                // Column 'J'
const PAYMENT_METHOD_COL = 11;        // Column 'K'
const INTERACT_REF_COL = 12;          // Column 'L'
const EMPTY_COL = 13;                 // Column 'M'
const IS_FEE_PAID_COL = 14;           // Column 'N'
const COLLECTION_DATE_COL = 15;       // Column 'O'
const COLLECTION_PERSON_COL = 16;       // Column 'P'
const IS_INTERNAL_COLLECTED_COL = 17;   // Column 'Q'
const COMMENTS_COL = 18;                // Column 'R'
const ATTENDANCE_STATUS_COL = 19;       // Column 'S'
const MEMBER_ID_COL = 20;               // Column 'T'


// MASTER SHEET CONSTANTS
const MASTER_NAME = 'MASTER';
const MASTER_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_NAME);
const MASTER_EMAIL_COL = 1;
const MASTER_FIRST_NAME_COL = 2;
const MASTER_LAST_NAME_COL = 3;
const MASTER_LAST_REG_SEM = 11;
const MASTER_FEE_STATUS = 14;   // Do not modify - Contains formula
const MASTER_FEE_EXPIRATION = 15;   // Do not modify - Contains formula
const MASTER_FEE_COLLECTOR = 16;  // Do not modify - Contains formula
const MASTER_COLLECTION_DATE = 17;
const MASTER_IS_INTERNAL_COLLECTED = 18;
const MASTER_PAYMENT_HIST = 19;
const MASTER_MEMBER_ID_COL = 22;


// MAPPING USED TO GET COL INDEX ACROSS SHEETS
const SHEET_COL_MAP = {
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


// MAPPING FROM FILLOUT REGISTRATION OBJ TO MAIN
const IMPORT_MAP = {
  'timestamp' : REGISTRATION_DATE_COL,
  'email' : EMAIL_COL,
  'firstName' : FIRST_NAME_COL,
  'lastName' : LAST_NAME_COL,
  'preferredName' : PREFERRED_NAME_COL,
  'year' : YEAR_COL,
  'program' : PROGRAM_COL,
  'memberDescription' : DESCRIPTION_COL,
  'referral' : REFERRAL_COL,
  'waiver' : WAIVER_COL,
  'paymentMethod' : PAYMENT_METHOD_COL,
  'interacRef' : INTERACT_REF_COL,
  'comments' : COMMENTS_COL,
}


// Function to get column mappings
function GET_COL_MAP_(sheet) { return SHEET_COL_MAP[sheet] || null };

const SEMESTER_CODE_MAP = new Map();
const ALL_SEMESTERS = ['Winter 2025', 'Fall 2024', 'Summer 2024', 'Winter 2024'];
const MASTER_COL_SIZE = 20;   // Range 'A:T' in 'MASTER'

const FEE_STATUS_ENUM = [
  "Paid",
  "Unpaid",
  "Expired"
];

// Found in `Internal Fee Collection` sheet
const INTERAC_ITEM_COL = 'A3'


// GSheet formula for IS_FEE_PAID_COL in `MASTER`
const isFeePaidFormula =`
  LET(row, ROW(),
      last_payment_sem, GET_FEE_EXPIRATION_DATE(INDIRECT("S" & row)), 
      expiration_date, SEMESTER_TO_DATE(last_payment_sem),
      IFS(INDIRECT("A" & row) = "", "", 
          INDIRECT("P" & row) = "(Fee Waived)", "Paid", 
          INDIRECT("S" & row) = "", "Unpaid", 
          expiration_date >= TODAY(), "Paid", 
          expiration_date < TODAY(), "Expired" )
)`;


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
 * 
 * @return {Boolean}  Parsed value.
 */

function parseBool(val) { 
  return val === true || val === "true";
}

