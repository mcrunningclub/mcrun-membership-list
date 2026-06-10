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
 * Number of (relevant???) columns in the master sheet
 * @const {number}
 */
const MASTER_COL_SIZE = 20;   // Range 'A:T' in 'MASTER'


/**
 * Current timezone
 * @const {string}
 */
const TIMEZONE = getUserTimeZone_();

/**
 * Club email
 * @const {str}
 */
const MCRUN_EMAIL = 'mcrunningclub@ssmu.ca';

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
 * LATEST COLUMN MAPPING FOR SEMESTER SHEET (S26) 
 */
const SEMESTER_COLS = {
  registrationDate: 1,
  email: 2,
  firstName: 3,
  lastName: 4,
  preferredName: 5,
  year: 6,
  program: 7,
  description: 8,
  optedIntoNewsletter: 9,
  optedIntoEmails: 10,
  referral: 11,
  waiver: 12,
  paymentMethod: 13,
  promo: 14,
  interacRef: 15,
  feeAmount: 16,
  feePaid: 17,
  collectionDate: 18,
  collectedBy: 19,
  isInternalCollected: 20,
  comments: 21,
  totalRuns: 22,
  totalPoints: 23,
  attendanceStatus: 24,
  memberId: 25
}


/** 
 * LATEST COLUMN MAPPING FOR MASTER SHEET (S26) 
 */            
const MASTER_COLS = {
  email: 1,            
  firstName: 2,            
  lastName: 3,            
  preferredName: 4,            
  year: 5,            
  program: 6,            
  waiver: 7,            
  description: 8,            
  referral: 9,            
  latestRegistration: 10,            
  latestSemester: 11,            
  regHistory: 12,            
  emptyCol: 13,            
  feePaid: 14,               // Do not modify - Contains formula
  feeExpiration: 15,               // Do not modify - Contains formula
  collectedBy: 16,               // Do not modify - Contains formula
  collectionDate: 17,            
  isInternalCollected: 18,            
  paymentHistory: 19,            
  comments: 20,            
  attendanceStatus: 21,
  memberId: 22            
}

/**
 * MAPPING FROM FILLOUT REGISTRATION OBJ TO MAIN
 */ 
const IMPORT_MAP = {
  'timestamp': REGISTRATION_DATE_COL,
  'email': EMAIL_COL,
  'firstName': FIRST_NAME_COL,
  'lastName': LAST_NAME_COL,
  'preferredName': PREFERRED_NAME_COL,
  'year': YEAR_COL,
  'program': PROGRAM_COL,
  'memberDescription': DESCRIPTION_COL,
  'referral': REFERRAL_COL,
  'waiver': WAIVER_COL,
  'paymentMethod': PAYMENT_METHOD_COL,
  'interacRef': INTERAC_REF_COL,
  'comments': COMMENTS_COL,
}

// Semester sheet constants
// Please update any changes made in active sheet
const REGISTRATION_DATE_COL = 1;      // Column 'A'
const EMAIL_COL = 2;                  // Column 'B'
const FIRST_NAME_COL = 3;             // Column 'C'
const LAST_NAME_COL = 4;              // Column 'D'
const PREFERRED_NAME_COL = 5;         // Column 'E'
const YEAR_COL = 6;                   // Column 'F'
const PROGRAM_COL = 7;                // Column 'G'
const DESCRIPTION_COL = 8;            // Column 'H'
const REFERRAL_COL = 11;               // Column K
const WAIVER_COL = 12;                // Column L
const PAYMENT_METHOD_COL = 13;        // Column M
const INTERAC_REF_COL = 15;          // Column O
const IS_FEE_PAID_COL = 17;           // Column Q
const COLLECTION_DATE_COL = 18;       // Column R
const COLLECTION_PERSON_COL = 19;       // Column S
const IS_INTERNAL_COLLECTED_COL = 20;   // Column T
const COMMENTS_COL = 21;                // Column U
const ATTENDANCE_STATUS_COL = 24;       // Column X
const MEMBER_ID_COL = 25;               // Column Y


// MASTER SHEET CONSTANTS
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


const SEMESTER_CODE_MAP = new Map();
const ALL_SEMESTERS = ['Winter 2025', 'Fall 2024', 'Summer 2024', 'Winter 2024'];


// GSheet formula for IS_FEE_PAID_COL in master sheet
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