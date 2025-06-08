// Code.gs - Main script file

// --- SCRIPT PROPERTIES (to be set in File > Project properties > Script properties) ---
// DEPUTY_INSTALL, DEPUTY_GEO, DEPUTY_ACCESS_TOKEN, DEPUTY_AUTH_TYPE
// SPREADSHEET_ID, SHEET_NAME, TIMESHEET_ID_COLUMN_INDEX
// ADMIN_EMAIL_FOR_NOTIFICATIONS

/**
 * Reusable helper function to fetch timesheets with a specific filter. Handles pagination.
 * @param {Object} searchFilter - The 'search' object for the Deputy API query.
 * @param {Object} config - An object containing API configuration.
 * @param {Array<string>} joinArray - The array of objects to join.
 * @return {Array} An array of timesheet objects.
 */
function fetchDeputyTimesheets(searchFilter, config, joinArray) {
  let allResults = [];
  let currentStartRecord = 0;
  let keepFetching = true;
  const MAX_RECORDS_PER_CALL = 500;

  while (keepFetching) {
    const queryPayload = {
      "search": searchFilter,
      "join": joinArray,
      "sort": { "Date": "asc", "StartTimeLocalized": "asc" },
      "max": MAX_RECORDS_PER_CALL,
      "start": currentStartRecord
    };

    const options = {
      'method': 'POST',
      'headers': { 'Authorization': `${config.authType} ${config.accessToken}` },
      'payload': JSON.stringify(queryPayload),
      'contentType': 'application/json',
      'muteHttpExceptions': true
    };

    Logger.log(`Fetching timesheets with filter: ${JSON.stringify(searchFilter)}. Start: ${currentStartRecord}`);
    try {
      const response = UrlFetchApp.fetch(config.endpoint, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const jsonData = JSON.parse(responseBody);
        if (Array.isArray(jsonData)) {
          allResults = allResults.concat(jsonData);
          if (jsonData.length < MAX_RECORDS_PER_CALL) {
            keepFetching = false;
          } else {
            currentStartRecord += jsonData.length;
          }
        } else {
          Logger.log(`API Response 200, but not an array. Stopping.`);
          keepFetching = false;
        }
      } else {
        const apiErrorMsg = `Deputy API Error: ${responseCode}. Filter: ${JSON.stringify(searchFilter)}. Response: ${responseBody}`;
        Logger.log(apiErrorMsg);
        if(config.adminEmail) MailApp.sendEmail(config.adminEmail, "Deputy Sync API Error", apiErrorMsg);
        return [];
      }
    } catch (e) {
      const scriptErrorMsg = `Exception during API call: ${e.toString()}`;
      Logger.log(scriptErrorMsg);
      if(config.adminEmail) MailApp.sendEmail(config.adminEmail, "Deputy Sync Script Error", scriptErrorMsg);
      return [];
    }
  }
  return allResults;
}

/**
 * Main function to fetch and process data.
 */
function mainDataFetchAndProcess() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {
    installName: scriptProperties.getProperty('DEPUTY_INSTALL'),
    geo: scriptProperties.getProperty('DEPUTY_GEO'),
    accessToken: scriptProperties.getProperty('DEPUTY_ACCESS_TOKEN'),
    authType: scriptProperties.getProperty('DEPUTY_AUTH_TYPE') || 'Bearer',
    spreadsheetId: scriptProperties.getProperty('SPREADSHEET_ID'),
    sheetName: scriptProperties.getProperty('SHEET_NAME'),
    timesheetIdColumnIndex: parseInt(scriptProperties.getProperty('TIMESHEET_ID_COLUMN_INDEX'), 10) || 0,
    adminEmail: scriptProperties.getProperty('ADMIN_EMAIL_FOR_NOTIFICATIONS')
  };

  if (!config.installName || !config.accessToken || !config.spreadsheetId || !config.adminEmail) {
    const errorMessage = "ERROR: One or more required Script Properties are not set.";
    Logger.log(errorMessage);
    if (config.adminEmail) MailApp.sendEmail(config.adminEmail, "Deputy Sync Configuration Error", errorMessage);
    return;
  }
  config.endpoint = `https://${config.installName}.${config.geo}.deputy.com/api/v1/resource/Timesheet/QUERY`;

  const { startDate, endDate } = getPreviousPayPeriodDates();
  Logger.log(`Querying Deputy Timesheets for dates: ${startDate} to ${endDate}`);

  const joinArray = ["EmployeeObject", "Leave", "LeaveRuleObject"];

  // --- Fetch Timesheets in Two Separate, Simple Batches ---
  
  // Filter 1: Manually approved shifts
  const manualApprovalFilter = {
    "s1": { "field": "Date", "type": "ge", "data": startDate },
    "s2": { "field": "Date", "type": "le", "data": endDate },
    "s3": { "field": "TimeApprover", "type": "gt", "data": 0 }
  };
  const manualApprovedTimesheets = fetchDeputyTimesheets(manualApprovalFilter, config, joinArray);
  
  // Filter 2: System-approved leave shifts
  const leaveApprovalFilter = {
    "s1": { "field": "Date", "type": "ge", "data": startDate },
    "s2": { "field": "Date", "type": "le", "data": endDate },
    "s3": { "field": "TimeApprover", "type": "eq", "data": -2 }
  };
  const leaveTimesheets = fetchDeputyTimesheets(leaveApprovalFilter, config, joinArray);

  // Combine and deduplicate results using a Map
  const allTimesheetsMap = new Map();
  manualApprovedTimesheets.forEach(ts => allTimesheetsMap.set(ts.Id, ts));
  leaveTimesheets.forEach(ts => allTimesheetsMap.set(ts.Id, ts));
  const allApiTimesheets = Array.from(allTimesheetsMap.values());
  
  Logger.log(`Total unique approved timesheets fetched: ${allApiTimesheets.length} (${manualApprovedTimesheets.length} manual, ${leaveTimesheets.length} leave)`);
  if (allApiTimesheets.length === 0) {
    Logger.log("No approved timesheets found for the specified period.");
    return;
  }

  // --- FINAL ADDITION: Sort the combined array by Timesheet ID in descending order ---
  Logger.log("Sorting all timesheets by ID in ascending order before processing.");
  allApiTimesheets.sort((a, b) => a.Id - b.Id);
  
  // --- Process and Write to Sheet ---
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    const sheetErrorMsg = `ERROR: Sheet "${config.sheetName}" not found.`;
    Logger.log(sheetErrorMsg);
    MailApp.sendEmail(config.adminEmail, "Deputy Sync Error - Sheet Not Found", sheetErrorMsg);
    return;
  }

  const existingTimesheetIDsInSheet = getExistingDeputyIDsFromSheet(sheet, config.timesheetIdColumnIndex);
  Logger.log(`Found ${existingTimesheetIDsInSheet.size} existing Timesheet IDs for deduplication.`);

  const rowsToAppend = [];
  const sheetHeaders = [ "TimeSheet ID", "Employee Id", "Employee Name", "Date", "Start", "End", "Mealbreak", "Total Hours", "Total Cost", "Employee Comment", "Area Name", "Location Name", "Leave", "Manager's Comment" ];
  checkAndWriteHeaderRow(sheet, sheetHeaders);

  for (const ts of allApiTimesheets) {
    if (existingTimesheetIDsInSheet.has(ts.Id.toString())) {
      Logger.log(`Skipping duplicate Timesheet ID: ${ts.Id}`);
      continue;
    }

    let rowData;
    const employeeObject = ts.EmployeeObject || {};
    const operationalUnitInfo = (ts._DPMetaData && ts._DPMetaData.OperationalUnitInfo) ? ts._DPMetaData.OperationalUnitInfo : {};

    if (ts.IsLeave) {
      Logger.log(`Processing Leave Timesheet ID: ${ts.Id}`);
      
      const leaveTypeName = (ts.LeaveRuleObject && ts.LeaveRuleObject.Name) 
          ? ts.LeaveRuleObject.Name 
          : "Leave (Rule Name Missing)";

      let startTime = '', endTime = '', totalHours = 0, leaveComment = '';
      
      if (ts.Leave && ts.Leave.LeavePayLineArray) {
        const payLine = ts.Leave.LeavePayLineArray.find(line => line.TimesheetId === ts.Id);
        if (payLine) {
          startTime = formatApiTimeToSheetTime(payLine.StartTimeLocalized);
          endTime = formatApiTimeToSheetTime(payLine.EndTimeLocalized);
          totalHours = parseFloat(payLine.Hours) || 0;
        }
      }
      
      if (!startTime && ts.Leave) { startTime = formatApiTimeToSheetTime(ts.Leave.StartTimeLocalized); }
      if (!endTime && ts.Leave) { endTime = formatApiTimeToSheetTime(ts.Leave.EndTimeLocalized); }
      if (totalHours === 0) { totalHours = (ts.Leave && ts.Leave.TotalHours) ? ts.Leave.TotalHours : ts.TotalTime || 0; }
      leaveComment = ts.Leave ? (ts.Leave.Comment || '') : ts.EmployeeComment || '';
      
      rowData = [ ts.Id.toString(), employeeObject.Id ? employeeObject.Id.toString() : '', employeeObject.DisplayName || '', formatApiDateToSheetDate(ts.Date), startTime, endTime, "0:00:00", totalHours, ts.Cost || 0, leaveComment, '', operationalUnitInfo.CompanyName || '', leaveTypeName, ts.SupervisorComment || ''];

    } else { // Regular work timesheet
      Logger.log(`Processing Regular Timesheet ID: ${ts.Id}`);
      rowData = [ ts.Id.toString(), employeeObject.Id ? employeeObject.Id.toString() : '', employeeObject.DisplayName || '', formatApiDateToSheetDate(ts.Date), formatApiTimeToSheetTime(ts.StartTimeLocalized), formatApiTimeToSheetTime(ts.EndTimeLocalized), formatDurationFromSeconds(calculateMealbreakDurationInSeconds(ts.Slots)), ts.TotalTime || 0, ts.Cost || 0, ts.EmployeeComment || '', operationalUnitInfo.OperationalUnitName || '', operationalUnitInfo.CompanyName || '', '', ts.SupervisorComment || ''];
    }
    rowsToAppend.push(rowData);
  }

  if (rowsToAppend.length > 0) {
    try {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, sheetHeaders.length).setValues(rowsToAppend);
      Logger.log(`Successfully appended ${rowsToAppend.length} new timesheet entries.`);
    } catch (e) {
      const sheetWriteErrorMsg = `Error writing to Google Sheet: ${e.toString()}`;
      Logger.log(sheetWriteErrorMsg);
      MailApp.sendEmail(config.adminEmail, "Deputy Sync Error - Sheet Write Failure", sheetWriteErrorMsg);
    }
  } else {
    Logger.log("No new timesheet entries to append.");
  }
}

/**
 * Wrapper function for the bi-weekly trigger. Checks if it's the correct week to run.
 */
function mainDataFetchAndProcess_biWeeklyCheck() {
  const firstRunYear = 2025, firstRunMonth = 5, firstRunDay = 4; // Wed, June 4th, 2025
  const firstScheduledRunDateUTC = new Date(Date.UTC(firstRunYear, firstRunMonth, firstRunDay));
  const now = new Date();
  const currentRunDateUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  const diffInDays = Math.floor((currentRunDateUTC - firstScheduledRunDateUTC) / (1000 * 60 * 60 * 24));

  Logger.log(`Bi-weekly check: First run date: ${firstScheduledRunDateUTC.toUTCString()}. Days since: ${diffInDays}.`);
  if (diffInDays >= 0 && diffInDays % 14 === 0) {
    Logger.log("This is an active Wednesday. Proceeding with data fetch.");
    mainDataFetchAndProcess();
  } else {
    Logger.log("Skipping run: Not a scheduled bi-weekly run date.");
  }
}

/**
 * Sets up the time-driven trigger. Run this function MANUALLY ONCE.
 */
function setupTrigger() {
  const functionName = "mainDataFetchAndProcess_biWeeklyCheck";
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted existing trigger for ${functionName}.`);
    }
  });
  ScriptApp.newTrigger(functionName)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .atHour(4)
      .inTimezone("America/New_York")
      .create();
  Logger.log(`New trigger created for ${functionName} to run every Wednesday at 4 AM EST/EDT.`);
}

// Code.gs

// ... (all your other working functions go here) ...


// --- TEMPORARY DEBUGGING FUNCTION TO FIND MANAGER'S COMMENT ---
/**
 * Fetches all timesheets for a specific date (May 20, 2025) to inspect all
 * potential comment fields returned by the API.
 */
function debugManagerComments() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {
    installName: scriptProperties.getProperty('DEPUTY_INSTALL'),
    geo: scriptProperties.getProperty('DEPUTY_GEO'),
    accessToken: scriptProperties.getProperty('DEPUTY_ACCESS_TOKEN'),
    authType: scriptProperties.getProperty('DEPUTY_AUTH_TYPE') || 'Bearer',
  };
  config.endpoint = `https://${config.installName}.${config.geo}.deputy.com/api/v1/resource/Timesheet/QUERY`;

  // --- We will look at a specific date you know has manager comments ---
  const targetDate = "2025-05-20"; 
  
  Logger.log(`DEBUG: Fetching ALL timesheets for date: ${targetDate} to find manager comments.`);

  // We are NOT filtering by approval status, to see everything.
  const queryPayload = {
    "search": {
      "s1": { "field": "Date", "type": "eq", "data": targetDate }
    },
    "join": ["EmployeeObject", "Leave", "LeaveRuleObject"]
  };

  const options = {
    'method': 'POST',
    'headers': { 'Authorization': `${config.authType} ${config.accessToken}` },
    'payload': JSON.stringify(queryPayload),
    'contentType': 'application/json',
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(config.endpoint, options);
  const responseBody = response.getContentText();
  Logger.log(`DEBUG: API Response Code: ${response.getResponseCode()}`);
  
  if (response.getResponseCode() === 200) {
    const jsonData = JSON.parse(responseBody);
    if (jsonData && jsonData.length > 0) {
      Logger.log(`---------- DEBUGGING COMMENTS FOR ${jsonData.length} TIMESHEETS ON ${targetDate} ----------`);
      
      jsonData.forEach((ts, index) => {
        const employeeName = ts.EmployeeObject ? ts.EmployeeObject.DisplayName : 'N/A';
        Logger.log(`--- Record #${index + 1} | Timesheet ID: ${ts.Id} | Employee: ${employeeName} ---`);
        Logger.log(`SupervisorComment (our current mapping): ${ts.SupervisorComment}`);
        Logger.log(`EmployeeComment (the other comment field): ${ts.EmployeeComment}`);
        // For leave, comments can be in other places too
        if (ts.IsLeave && ts.Leave) {
          Logger.log(`Leave Comment: ${ts.Leave.Comment}`);
          Logger.log(`Leave ApprovalComment: ${ts.Leave.ApprovalComment}`);
        }
        Logger.log(`Full Timesheet Object: ${JSON.stringify(ts)}`); // Log the whole object as a final check
        Logger.log(`-----------------------------------------------------`);
      });

    } else {
      Logger.log("DEBUG: No timesheets found for this date.");
    }
  } else {
    Logger.log(`DEBUG: API Error: ${responseBody}`);
  }
}