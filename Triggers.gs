/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/**
 * Name of fee payment check trigger
 * @const {string}
 */
const TRIGGER_FUNC = runFeeChecker.name;

/**
 * ID of fee payment check trigger
 * @const {string}
 */
const TRIGGER_BASE_ID = 'feeCheckTrigger';

/**
 * Max number of times to check for fee payment
 * @const {number}
 */
const FEE_MAX_CHECKS = 3;

/**
 * Trigger frequency in minutes
 * @const {number}
 */
const TRIGGER_FREQUENCY = 5;


/**
 * Create time-based trigger to check fee payment.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 20, 2025
 * @update  May 20, 2025
 */

function createNewFeeTrigger_(row, feeDetails) {
  const scriptProperties = PropertiesService.getScriptProperties();

  const trigger = ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .everyMinutes(TRIGGER_FREQUENCY)
    .create();

  // Store trigger details using 'memberName' as key
  const triggerData = JSON.stringify({
    tries: 1,
    triggerId: trigger.getUniqueId(),
    feeDetails: feeDetails,
    memberRow: row,
  });

  // Label trigger key with member name, and log trigger data
  const key = TRIGGER_BASE_ID + " " + (feeDetails.memberName).replace(' ', '');
  
  scriptProperties.setProperty(key, triggerData);
  Logger.log(`Created new trigger '${key}', running every ${TRIGGER_FREQUENCY} min.\n\n${triggerData}`);
}


/**
 * Handler function for time-based trigger to check fee payment.
 * 
 * Includes helper functions to check for payment, increment the
 * number of times checked, and clean up/delete the trigger.
 * 
 * No arguments allowed since trigger does not accept any.
 * Workaround: store member details in script properties.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  May 20, 2025
 * @update  May 26, 2025
 */

function runFeeChecker() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const allProps = scriptProperties.getProperties();

  for (let key in allProps) {
    if (!key.startsWith(TRIGGER_BASE_ID)) continue;

    const triggerData = JSON.parse(allProps[key]);
    let { tries, triggerId, feeDetails, memberRow } = triggerData;
    console.log(`Trigger Data: ${allProps[key]}`);

    // First check memberRow and update 'triggerData' if needed
    memberRow = checkMemberRow_(feeDetails.email, memberRow);
    triggerData.memberRow = memberRow;
    
    if (isPaymentFound_(memberRow)) {
      // If found, clean up trigger and data in script properties
      cleanUpTrigger(key, triggerId);
      Logger.log(`✅ Payment found for member '${feeDetails.memberName}' after ${tries} tries`);
    }
    else if (tries < FEE_MAX_CHECKS) {
      // Limit not reach, check again and increment 'tries'
      incrementTries_(key, triggerData);
      const isPaid = isPaid_(memberRow, feeDetails);
      Logger.log(`Payment verification for member '${feeDetails.memberName}' returned: ${isPaid}`);
      
      // Only update if member fee has been paid
      if (isPaid) {
        updateMasterPayment_(feeDetails.email, feeDetails.paymentMethod);
        Logger.log(`Completed '${updateMasterPayment_.name}' in 'MASTER_SHEET'`);
      }
    }
    else {
      // Send email notification if limit is reached
      cleanUpTrigger(key, triggerId);
      notifyUnidentifiedPayment_(feeDetails.memberName);
      Logger.log(`❌ Max tries reached for member '${feeDetails.memberName}', sending email and stopping checks`);
    }
  }

  /** Helper: check if payment already found */
  function isPaymentFound_(memberRow) {
    const sheet = SEMESTER_SHEET;
    const currentFeeValue = sheet.getRange(memberRow, SEMESTER_COLS.FEE_PAID).getValue().toString();
    return parseBool_(currentFeeValue.trim());
  }

  /** Helper: validate memberRow, else return updated row */
  function checkMemberRow_(memberEmail, memberRow) {
    const sheet = SEMESTER_SHEET;
    const currentEmail = sheet.getRange(memberRow, SEMESTER_COLS.EMAIL).getValue();

    // If emails don't match, find updated memberRow
    if (currentEmail !== memberEmail) {
      memberRow = findMemberByEmail(memberEmail, sheet);
    }
    return memberRow;
  }

  /** Helper: increment tries and log data */
  function incrementTries_(key, triggerData) {
    Logger.log(`Fee payment check #${triggerData.tries} for member ${triggerData.feeDetails.memberName}`);
    triggerData.tries++;
    scriptProperties.setProperty(key, JSON.stringify(triggerData));
  }

  /** Helper: remove trigger and data in script properties */
  function cleanUpTrigger(key, triggerId) {
    deleteTriggerById_(triggerId);
    scriptProperties.deleteProperty(key);
  }

  /**
   * Deletes a trigger by its unique ID.
   *
   * This function iterates through all project triggers to find and delete the one
   * with the specified unique ID. If the trigger is not found, it throws an error.
   *
   * @param {string} triggerId - The unique ID of the trigger to delete.
   */
  function deleteTriggerById_(triggerId) {
    const triggers = ScriptApp.getProjectTriggers();

    for (let trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`Trigger with id ${triggerId} deleted!`);
        return;
      }
    }
    // If we reach here, the trigger was not found
    throw new Error(`⚠️ Trigger with id ${triggerId} not found`);
  }
}

