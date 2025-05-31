// HelperFunctions.gs

/**
 * Formats an API date string (e.g., "2025-05-21T00:00:00-04:00") to "MM/DD/YYYY".
 * @param {string} apiDateStr The date string from the API.
 * @return {string} The formatted date string.
 */
function formatApiDateToSheetDate(apiDateStr) {
  if (!apiDateStr) return '';
  try {
    const dateObj = new Date(apiDateStr);
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
  } catch (e) {
    Logger.log(`Error formatting date string: ${apiDateStr}. Error: ${e}`);
    return apiDateStr;
  }
}

/**
 * Formats an API localized time string (e.g., "2025-05-21T06:43:00-04:00") to "HH:MM:SS".
 * @param {string} apiLocalizedTimeStr The localized time string from API (e.g., StartTimeLocalized).
 * @return {string} The formatted time string.
 */
function formatApiTimeToSheetTime(apiLocalizedTimeStr) {
  if (!apiLocalizedTimeStr) return '';
  try {
    const timePart = apiLocalizedTimeStr.split('T')[1];
    return timePart.substring(0, 8);
  } catch (e) {
    Logger.log(`Error formatting time string: ${apiLocalizedTimeStr}. Error: ${e}`);
    return apiLocalizedTimeStr;
  }
}

/**
 * Calculates total meal break duration in seconds from the Slots array.
 * @param {Array<Object>} slotsArray The 'Slots' array from a timesheet entry.
 * @return {number} Total meal break duration in seconds.
 */
function calculateMealbreakDurationInSeconds(slotsArray) {
  if (!slotsArray || !Array.isArray(slotsArray)) return 0;
  let totalMealbreakSeconds = 0;
  slotsArray.forEach(slot => {
    if (slot.strType === 'B' && slot.strTypeName === 'Meal Break' &&
        typeof slot.intUnixEnd === 'number' && typeof slot.intUnixStart === 'number') {
      totalMealbreakSeconds += (slot.intUnixEnd - slot.intUnixStart);
    }
  });
  return totalMealbreakSeconds;
}

/**
 * Formats a duration in seconds to "H:MM:SS" string.
 * @param {number} totalSeconds The duration in seconds.
 * @return {string} The formatted duration string.
 */
function formatDurationFromSeconds(totalSeconds) {
  if (totalSeconds === null || totalSeconds === undefined || isNaN(totalSeconds)) return '0:00:00';
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = Math.round(totalSeconds % 60);

  return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
}

/**
 * Fetches all Leave Rules from Deputy and returns them as a map.
 * @param {string} accessToken The API access token.
 * @param {string} authType The API auth type ('Bearer' or 'DeputyKey').
 * @param {string} deputyApiBaseUrl The base URL for the Deputy API.
 * @return {Object} A map of { leaveRuleId: "Leave Rule Name" }. Returns empty object on failure.
 */
function fetchAllLeaveRules(accessToken, authType, deputyApiBaseUrl) {
  const leaveRuleMap = {};
  let currentStartRecord = 0;
  const MAX_RECORDS_PER_CALL = 500; // Deputy's default max
  let keepFetching = true;
  const leaveRuleQueryEndpoint = `${deputyApiBaseUrl}/resource/LeaveRule/QUERY`; // Or just /LeaveRule if GET based

  Logger.log("Fetching all leave rules from Deputy...");

  while (keepFetching) {
    // Assuming QUERY endpoint for consistency, even if search is empty for "all"
    const queryPayload = {
      "search": {}, // Empty search to get all rules
      "sort": { "Name": "asc" }, // Optional: sort by name
      "max": MAX_RECORDS_PER_CALL,
      "start": currentStartRecord
    };

    const options = {
      'method': 'POST', // If using /QUERY
      // 'method': 'GET', // If /LeaveRule supports GET with max/start query params
      'headers': { 'Authorization': `${authType} ${accessToken}` },
      'payload': JSON.stringify(queryPayload), // If POST
      'contentType': 'application/json',
      'muteHttpExceptions': true
    };

    // If /resource/LeaveRule is a GET endpoint that supports pagination via query parameters:
    // const leaveRuleListEndpoint = `${deputyApiBaseUrl}/resource/LeaveRule?max=${MAX_RECORDS_PER_CALL}&start=${currentStartRecord}`;
    // const options = {
    //   'method': 'GET',
    //   'headers': { 'Authorization': `${authType} ${accessToken}` },
    //   'contentType': 'application/json',
    //   'muteHttpExceptions': true
    // };
    // const response = UrlFetchApp.fetch(leaveRuleListEndpoint, options); // Use this if GET

    Logger.log(`Fetching leave rules. Start: ${currentStartRecord}, Max: ${MAX_RECORDS_PER_CALL}`);
    try {
      const response = UrlFetchApp.fetch(leaveRuleQueryEndpoint, options); // Use this if POST with QUERY
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const jsonData = JSON.parse(responseBody);
        if (Array.isArray(jsonData)) {
          Logger.log(`Fetched ${jsonData.length} leave rules in this call.`);
          jsonData.forEach(rule => {
            if (rule.Id && rule.Name) {
              leaveRuleMap[rule.Id] = rule.Name;
            }
          });
          if (jsonData.length < MAX_RECORDS_PER_CALL) {
            keepFetching = false;
          } else {
            currentStartRecord += jsonData.length;
          }
        } else {
          Logger.log(`LeaveRule API Response 200, but not an array. Body: ${responseBody}. Stopping pagination.`);
          keepFetching = false;
          break;
        }
      } else {
        Logger.log(`Deputy API Error fetching LeaveRules: ${responseCode}. Response: ${responseBody}.`);
        keepFetching = false; // Stop fetching if there's an error
        break;
      }
    } catch (e) {
      Logger.log(`Exception during LeaveRule API call: ${e.toString()}. Stack: ${e.stack}.`);
      keepFetching = false; // Stop fetching on exception
      break;
    }
  }
  Logger.log(`Finished fetching leave rules. Total rules mapped: ${Object.keys(leaveRuleMap).length}`);
  return leaveRuleMap;
}


/**
 * Looks up the name of a LeaveRule by its ID from a pre-fetched map.
 * @param {number} leaveRuleId The ID of the leave rule.
 * @param {Object} leaveRuleMap A map of { ruleId: "Rule Name" }.
 * @return {string} The name of the leave rule or a placeholder.
 */
function lookupLeaveRuleName(leaveRuleId, leaveRuleMap) {
  if (!leaveRuleId) {
    return ""; // No rule ID provided
  }
  if (leaveRuleMap && leaveRuleMap[leaveRuleId]) {
    return leaveRuleMap[leaveRuleId]; // Return cached/mapped name
  }
  Logger.log(`LeaveRule Name for ID ${leaveRuleId} not found in pre-fetched map.`);
  return `Unknown Rule ID (${leaveRuleId})`; // Fallback if not in map
}

/**
 * Calculates the start and end dates for the target pay period.
 * The period is a 14-day span (Monday to Sunday).
 * The Sunday end date is 3 days before the script's current run date.
 * @return {{startDate: string, endDate: string}} Object with ISO formatted start and end dates.
 */
function getPreviousPayPeriodDates() {
  const scriptRunDate = new Date(); // This is the date the script is actually running
  scriptRunDate.setHours(0, 0, 0, 0); // Normalize to start of the day for consistent date math

  // Pay Period End Date: Sunday, 3 days before the scriptRunDate (Wednesday)
  const payPeriodEndDate = new Date(scriptRunDate);
  payPeriodEndDate.setDate(scriptRunDate.getDate() - 3); // e.g., If Wed, this becomes Sunday
  // Ensure it's actually a Sunday. If script runs on a day other than Wed during manual test,
  // this might not be Sunday. For a trigger on Wed, it will be.
  // This is simple: the rule is "Sunday that ends 3 days before script run".
  // So, if scriptRunDate is Wed, then indeed payPeriodEndDate is the preceding Sunday.

  payPeriodEndDate.setHours(23, 59, 59, 999); // Ensure it's the end of that Sunday

  // Pay Period Start Date: Monday, 13 days before the payPeriodEndDate
  const payPeriodStartDate = new Date(payPeriodEndDate);
  payPeriodStartDate.setDate(payPeriodEndDate.getDate() - 13);
  payPeriodStartDate.setHours(0, 0, 0, 0); // Ensure it's the start of that Monday

  const isoFormat = (date) => {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
  };
  
  Logger.log(`Script Run Date (for calc): ${scriptRunDate.toDateString()}`);
  Logger.log(`Calculated Pay Period Start Date: ${payPeriodStartDate.toDateString()}`);
  Logger.log(`Calculated Pay Period End Date: ${payPeriodEndDate.toDateString()}`);

  return {
    startDate: isoFormat(payPeriodStartDate),
    endDate: isoFormat(payPeriodEndDate)
  };
}


/**
 * Retrieves existing Deputy Timesheet IDs from the specified column in the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number} idColumnIndex Zero-based index of the column containing Timesheet IDs.
 * @return {Set<string>} A Set of existing Timesheet IDs.
 */
function getExistingDeputyIDsFromSheet(sheet, idColumnIndex) {
  const ids = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { // Assuming row 1 is header
    return ids;
  }
  const range = sheet.getRange(2, idColumnIndex + 1, lastRow - 1, 1); // row, col, numRows, numCols
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null && values[i][0] !== undefined) {
      ids.add(values[i][0].toString());
    }
  }
  return ids;
}

/**
 * Checks if a header row exists and writes it if the sheet is empty.
 * Optionally verifies existing headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {Array<string>} expectedHeaders An array of expected header strings.
 */
function checkAndWriteHeaderRow(sheet, expectedHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    Logger.log("Appended header row to empty sheet.");
  } else {
    // Optional: More robust header check if needed. For now, just log if length differs.
    const currentHeaderRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    if (currentHeaderRange.getWidth() < expectedHeaders.length) {
        Logger.log(`WARNING: Existing headers in sheet are shorter than expected. Expected ${expectedHeaders.length} columns.`);
    }
    // Simple check based on first few headers
    const currentHeaders = sheet.getRange(1, 1, 1, Math.min(expectedHeaders.length, currentHeaderRange.getWidth())).getValues()[0];
    let partialMatch = true;
    for(let i = 0; i < Math.min(expectedHeaders.length, currentHeaders.length); i++){
        if(expectedHeaders[i] !== currentHeaders[i]){
            partialMatch = false;
            break;
        }
    }
    if(!partialMatch){
        Logger.log(`WARNING: Sheet headers do not seem to match expected headers.
                    Expected (start): ${expectedHeaders.slice(0,3).join(', ')}...
                    Actual (start):   ${currentHeaders.slice(0,3).join(', ')}...`);
    }
  }
}