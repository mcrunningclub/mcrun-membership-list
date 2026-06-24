/**
 * Users authorized to use the McRUN menu.
 * 
 * Prevents unwanted data overwrite in Gsheet.
 * 
 * @const {string[]}
 */
const ADMINS_ = [
  'mcrunningclub@ssmu.ca',
  'ademetriou8@gmail.com',
  'andreysebastian10.g@gmail.com',
  'monaliu832@gmail.com',
];

/**
 * Sheet name corresponding to current semester
 * 
 * MUST UPDATE EVERY SEMESTER!
 * 
 * @constant {string}
 */
const SHEET_NAME = 'Summer 2026';

/**
 * Sheet object corresponding to current semester
 * 
 * @constant {SpreadsheetApp.Sheet}
 */
const SEMESTER_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

/**
 * Name of sheet with imports from registration form
 * 
 * @const {string}
 */
const IMPORT_NAME = 'Import';

/**
 * Sheet object corresponding to imports from registration form
 * 
 * @const {SpreadsheetApp.Sheet}
 */
const IMPORT_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_NAME);

/**
 * ID of sheet with imports from registration form
 * 
 * @const {number}
 */
const IMPORT_SHEET_ID = 973209987;

/**
 * Name of master sheet
 * @const {string}
 */
const MASTER_NAME = 'MASTER';

/**
 * Sheet object of master sheet
 * @const {SpreadsheetApp.Sheet}
 */
const MASTER_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_NAME);

/**
 * Current timezone
 * @const {string}
 */
const TIMEZONE = getUserTimeZone_();

/**
 * Club email
 * @const {string}
 */
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';

// LAST UPDATED: MAY 15, 2025
// PLEASE UPDATE WHEN NEEDED
/**
 * Email address of Zeffy emails
 * @const {string}
 */
const ZEFFY_EMAIL = 'contact@zeffy.com';

/**
 * Email address (ending) of Interac emails
 * @const {string}
 */
const INTERAC_EMAIL = 'interac.ca';    // Interac email addresses end in "interac.ca"

/**
 * Email address (ending) of Stripe emails
 * @const {string}
 */
const STRIPE_EMAIL = 'stripe.com';

/**
 * Gmail label for online payment emails
 * @const {string}
 */
const ONLINE_LABEL = 'Fee Payments/Online Emails';

/**
 * Gmail label for Interac payment emails
 * @const {string}
 */
const INTERAC_LABEL = 'Fee Payments/Interac Emails';

/**
 * Cells for each payment method. Found in `Internal Fee Collection` sheet.
 */
const INTERAC_ITEM_COL = 'A3';
const ONLINE_PAYMENT_ITEM_COL = 'A4';
const FEE_WAIVED_ITEM_COL = 'A5';


/**
 * Length of membership in years
 * @const {number}
 */
const MEMBERSHIP_DURATION = 1;

/**
 * DRIVE URL CONTAINING WAIVERS; NOT CONFIDENTIAL
 * 
 * @constant {string}
 */
const WAIVER_DRIVE_ID = '1lNAvGMsm-ixa7rAQHqTdd_gV-W4WNwpOdx4Zx7S_AZ8_6EQ8ammSEwy3A3xzbPsPp7eUnvnf';

/**
 * Maps column letters to numbers (1-indexed)
 * @const {Object}
 */
const COL = {
  A: 1,
  B: 2,
  C: 3,
  D: 4,
  E: 5,
  F: 6,
  G: 7,
  H: 8,
  I: 9,
  J: 10,
  K: 11,
  L: 12,
  M: 13,
  N: 14,
  O: 15,
  P: 16,
  Q: 17,
  R: 18,
  S: 19,
  T: 20,
  U: 21,
  V: 22,
  W: 23,
  X: 24,
  Y: 25,
  Z: 26
}

/** 
 * LATEST COLUMN MAPPING FOR SEMESTER SHEET (S26)
 * If REMOVING A CONSTANT, ENSURE IT IS NOT USED IN THE SCRIPTS!!
 * @const {Object}
 */
const SEMESTER_COLS = {
  REGISTRATION_DATE: COL.A,
  EMAIL: COL.B,
  FIRST_NAME: COL.C,
  LAST_NAME: COL.D,
  PREFERRED_NAME: COL.E,
  YEAR: COL.F,
  PROGRAM: COL.G,
  DESCRIPTION: COL.H,
  OPTED_INTO_NEWSLETTER: COL.I,
  OPTED_INTO_EMAILS: COL.J,
  REFERRAL: COL.K,
  WAIVER: COL.L,
  PAYMENT_METHOD: COL.M,
  PROMO: COL.N,
  INTERAC_REF: COL.O,
  FEE_AMOUNT: COL.P,
  FEE_PAID: COL.Q,
  COLLECTION_DATE: COL.R,
  COLLECTED_BY: COL.S,
  IS_INTERNAL_COLLECTED: COL.T,
  COMMENTS: COL.U,
  TOTAL_RUNS: COL.V,
  TOTAL_POINTS: COL.W,
  ATTENDANCE_STATUS: COL.X,
  MEMBER_ID: COL.Y
}


/** 
 * LATEST COLUMN MAPPING FOR MASTER SHEET (S26)
 * If REMOVING A CONSTANT, ENSURE IT IS NOT USED IN THE SCRIPTS!!
 * @const {Object}
 */            
const MASTER_COLS = {
  EMAIL: COL.A,
  FIRST_NAME: COL.B,
  LAST_NAME: COL.C,
  PREFERRED_NAME: COL.D,
  YEAR: COL.E,
  PROGRAM: COL.F,
  WAIVER: COL.G,
  DESCRIPTION: COL.H,
  REFERRAL: COL.I,
  LATEST_REG_DATE: COL.J,
  LATEST_REG_SEMESTER: COL.K,
  REG_HISTORY: COL.L,
  EMPTY: COL.M,
  FEE_PAID: COL.N,               // Do not modify - Contains formula
  FEE_EXPIRATION: COL.O,         // Do not modify - Contains formula
  COLLECTED_BY: COL.P,           // Do not modify - Contains formula
  COLLECTION_DATE: COL.Q,
  IS_INTERNAL_COLLECTED: COL.R,
  PAYMENT_HISTORY: COL.S,
  COMMENTS: COL.T,
  ATTENDANCE_STATUS: COL.U,
  MEMBER_ID: COL.V
}

/**
 * MAPPING FROM FILLOUT REGISTRATION OBJ TO SEMESTER SHEET
 * @const {Object}
 */ 
const IMPORT_MAP = {
  'timestamp': SEMESTER_COLS.REGISTRATION_DATE,
  'email': SEMESTER_COLS.EMAIL,
  'firstName': SEMESTER_COLS.FIRST_NAME,
  'lastName': SEMESTER_COLS.LAST_NAME,
  'preferredName': SEMESTER_COLS.YEAR,
  'year': SEMESTER_COLS.YEAR,
  'program': SEMESTER_COLS.PROGRAM,
  'memberDescription': SEMESTER_COLS.DESCRIPTION,
  'referral': SEMESTER_COLS.REFERRAL,
  'waiver': SEMESTER_COLS.WAIVER,
  'paymentAmount': SEMESTER_COLS.FEE_AMOUNT,
  'paymentMethod': SEMESTER_COLS.PAYMENT_METHOD,
  'interacRef': SEMESTER_COLS.INTERAC_REF,
  'comments': SEMESTER_COLS.COMMENTS,
  'automatedEmailConsent': SEMESTER_COLS.OPTED_INTO_EMAILS
}


/**
 * Fields in array from processing last row in semester sheet (semester columns but 0-indexed)
 * NOT ALL FIELDS ARE IN THIS MAPPING, only the ones needed to move to master sheet
 * @const {Object}
 */
const PROCESSED_ARR = {
  LATEST_REG_DATE: SEMESTER_COLS.REGISTRATION_DATE - 1,
  EMAIL: SEMESTER_COLS.EMAIL - 1,
  FIRST_NAME: SEMESTER_COLS.FIRST_NAME - 1,
  LAST_NAME: SEMESTER_COLS.LAST_NAME - 1,
  PREFERRED_NAME: SEMESTER_COLS.PREFERRED_NAME - 1,
  YEAR: SEMESTER_COLS.YEAR - 1,
  PROGRAM: SEMESTER_COLS.PROGRAM - 1,
  DESCRIPTION: SEMESTER_COLS.DESCRIPTION - 1,
  REFERRAL: SEMESTER_COLS.REFERRAL - 1,
  WAIVER: SEMESTER_COLS.WAIVER - 1,
  FEE_PAID_SEM: SEMESTER_COLS.FEE_PAID - 1,
  COLLECTION_DATE: SEMESTER_COLS.COLLECTION_DATE - 1,
  COLLECTED_BY: SEMESTER_COLS.COLLECTED_BY - 1,
  IS_INTERNAL_COLLECTED: SEMESTER_COLS.IS_INTERNAL_COLLECTED - 1,
  COMMENTS: SEMESTER_COLS.COMMENTS - 1,
  ATTENDANCE_STATUS: SEMESTER_COLS.ATTENDANCE_STATUS - 1,
  MEMBER_ID: SEMESTER_COLS.MEMBER_ID - 1,
  LATEST_REG_SEM: Object.entries(SEMESTER_COLS).length,
  REG_HISTORY: Object.entries(SEMESTER_COLS).length + 1,
  EMPTY: -1       // Currently no empty columns
};

/**
 * Number of cells that can be edited at once (for onEdit function)
 * @const {number}
 */
const CELL_EDIT_LIMIT = 4;

/**
 * Mapping from semesters names to semester codes e.g. Winter 2025 -> W25
 * @const {Map}
 */
const SEMESTER_CODE_MAP = new Map();

/**
 * List of all semesters (names) which have sheets
 * 
 * @const {string[]}
 */
const ALL_SEMESTERS = ['Summer 2026', 'Winter 2025', 'Fall 2024', 'Summer 2024', 'Winter 2024'];


/**
 * GSheet formula for IS_FEE_PAID_COL in master sheet
 * 
 * @const {string}
 */
const isFeePaidFormula = `
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
 * Name of property that index store is saved under
 * @const {string}
 */
const INDEX_STORE_NAME = "letterIndexStore";


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
        emailCol: SEMESTER_COLS.EMAIL,
        memberIdCol: SEMESTER_COLS.MEMBER_ID,
        feeStatus: SEMESTER_COLS.FEE_PAID,   // Boolean value
        collectionDate: SEMESTER_COLS.COLLECTION_DATE,
        collector: SEMESTER_COLS.COLLECTED_BY,
        isInternalCollected: SEMESTER_COLS.IS_INTERNAL_COLLECTED,
      },
      [MASTER_NAME]: {
        emailCol: MASTER_COLS.EMAIL,
        memberIdCol: MASTER_COLS.MEMBER_ID,
        feeStatus: MASTER_COLS.PAYMENT_HISTORY,   // String with semester code(s)
        collectionDate: MASTER_COLS.COLLECTION_DATE,
        collector: MASTER_COLS.COLLECTED_BY,
        isInternalCollected: MASTER_COLS.IS_INTERNAL_COLLECTED,
      },
    };
  }

  return SHEET_COL_MAP[sheet] || null;
}
