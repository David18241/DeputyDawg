// Code.gs - Main script file

// --- SCRIPT PROPERTIES (to be set in File > Project properties > Script properties) ---
// DEPUTY_INSTALL: your_install_name
// DEPUTY_GEO: your_geo_code (e.g., na, au, eu)
// DEPUTY_ACCESS_TOKEN: your_deputy_api_token
// DEPUTY_AUTH_TYPE: 'Bearer' or 'DeputyKey'
// SPREADSHEET_ID: id_of_your_google_sheet
// SHEET_NAME: name_of_the_tab (e.g., 'Time Card Data')
// TIMESHEET_ID_COLUMN_INDEX: 0 (if "DeputyTimesheetID" is in Col A for deduplication)
// ADMIN_EMAIL_FOR_NOTIFICATIONS: your_admin_email@example.com (for error alerts)

function mainDataFetchAndProcess() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const installName = scriptProperties.getProperty('DEPUTY_INSTALL');
  const geo = scriptProperties.getProperty('DEPUTY_GEO');
  const accessToken = scriptProperties.getProperty('DEPUTY_ACCESS_TOKEN');
  const authType = scriptProperties.getProperty('DEPUTY_AUTH_TYPE') || 'Bearer';

  const SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');
  const SHEET_NAME = scriptProperties.getProperty('SHEET_NAME');
  let TIMESHEET_ID_COLUMN_INDEX = parseInt(scriptProperties.getProperty('TIMESHEET_ID_COLUMN_INDEX'), 10);
  if (isNaN(TIMESHEET_ID_COLUMN_INDEX)) {
    Logger.log("TIMESHEET_ID_COLUMN_INDEX Script Property is not a valid number, defaulting to 0.");
    TIMESHEET_ID_COLUMN_INDEX = 0;
  }
  const ADMIN_EMAIL = scriptProperties.getProperty('ADMIN_EMAIL_FOR_NOTIFICATIONS');


  if (!installName || !geo || !accessToken || !SPREADSHEET_ID || !SHEET_NAME || !ADMIN_EMAIL) {
    const errorMessage = "ERROR: One or more required Script Properties are not set (DEPUTY_*, SPREADSHEET_*, ADMIN_EMAIL_*).";
    Logger.log(errorMessage);
    if (ADMIN_EMAIL) MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync Configuration Error", errorMessage);
    else Logger.log("Admin email not configured to send error notification.");
    return;
  }

  const deputyApiBaseUrl = `https://${installName}.${geo}.deputy.com/api/v1`;

  // --- 1. Fetch all Leave Rules first ---
  const leaveRuleMap = fetchAllLeaveRules(accessToken, authType, deputyApiBaseUrl);
  if (Object.keys(leaveRuleMap).length === 0) {
    Logger.log("Warning: No leave rules were fetched or mapped. Leave types might be incorrect.");
    // Decide if this is a critical error, for now, we'll proceed.
  }

  // --- 2. Fetch Timesheets ---
  const { startDate, endDate } = getPreviousPayPeriodDates();
  Logger.log(`Querying Deputy Timesheets for dates: ${startDate} to ${endDate}`);

  let allApiTimesheets = [];
  let currentStartRecord = 0;
  const MAX_RECORDS_PER_CALL = 500;
  let keepFetching = true;

  const timesheetQueryEndpoint = `${deputyApiBaseUrl}/resource/Timesheet/QUERY`;

  while (keepFetching) {
    const queryPayload = {
      "search": {
        "s1": { "field": "Date", "type": "ge", "data": startDate },
        "s2": { "field": "Date", "type": "le", "data": endDate },
        "s3": { "field": "TimeApprover", "type": "gt", "data": 0 }
      },
      "join": ["EmployeeObject"],
      "sort": { "Date": "asc", "StartTimeLocalized": "asc" },
      "max": MAX_RECORDS_PER_CALL,
      "start": currentStartRecord
    };

    const options = {
      'method': 'POST',
      'headers': { 'Authorization': `${authType} ${accessToken}` },
      'payload': JSON.stringify(queryPayload),
      'contentType': 'application/json',
      'muteHttpExceptions': true
    };

    Logger.log(`Fetching timesheets from Deputy. Start record: ${currentStartRecord}, Max: ${MAX_RECORDS_PER_CALL}`);
    try {
      const response = UrlFetchApp.fetch(timesheetQueryEndpoint, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const jsonData = JSON.parse(responseBody);
        if (Array.isArray(jsonData)) {
          Logger.log(`Fetched ${jsonData.length} timesheet records in this call.`);
          allApiTimesheets = allApiTimesheets.concat(jsonData);
          if (jsonData.length < MAX_RECORDS_PER_CALL) {
            keepFetching = false;
            Logger.log("Last page of timesheet data reached.");
          } else {
            currentStartRecord += jsonData.length;
          }
        } else {
          Logger.log(`Timesheet API Response 200, but not an array. Body: ${responseBody}. Stopping.`);
          keepFetching = false; break;
        }
      } else {
        Logger.log(`Deputy Timesheet API Error: ${responseCode}. Response: ${responseBody}. Stopping.`);
        MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync API Error (Timesheets)", `Failed to fetch timesheets. Code: ${responseCode}. Body: ${responseBody}`);
        keepFetching = false; break;
      }
    } catch (e) {
      Logger.log(`Exception during Deputy Timesheet API call: ${e.toString()}. Stack: ${e.stack}. Stopping.`);
      MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync Script Error (Timesheets)", `Exception: ${e.toString()}\nStack: ${e.stack}`);
      keepFetching = false; break;
    }
    // Utilities.sleep(200); // Optional delay
  }

  Logger.log(`Total approved timesheets fetched from Deputy API: ${allApiTimesheets.length}`);
  if (allApiTimesheets.length === 0 && !keepFetching) { // Check !keepFetching to ensure it wasn't an error stop
    Logger.log("No approved timesheets found for the specified period, or an error occurred before fetching any.");
    return;
  }


  // --- 3. Process and Write to Sheet ---
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const sheetErrorMsg = `ERROR: Sheet "${SHEET_NAME}" not found in Spreadsheet ID "${SPREADSHEET_ID}".`;
    Logger.log(sheetErrorMsg);
    MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync Error - Sheet Not Found", sheetErrorMsg);
    return;
  }

  const existingTimesheetIDsInSheet = getExistingDeputyIDsFromSheet(sheet, TIMESHEET_ID_COLUMN_INDEX);
  Logger.log(`Found ${existingTimesheetIDsInSheet.size} existing Timesheet IDs in the sheet for deduplication.`);

  const rowsToAppend = [];
  // Updated headers to match your simplified sheet
  const sheetHeaders = [
      "TimeSheet ID", "Employee Export Code", "Employee Name", "Date", "Start", "End", "Mealbreak",
      "Total Hours", "Total Cost", "Employee Comment", "Area Name", "Location Name", "Leave", "Manager's Comment"
  ];
  // Note: Your sample CSV has "TimeSheet ID" (with space) for the first column. My code uses "DeputyTimesheetID".
  // I'll use your CSV's "TimeSheet ID" here for consistency.
  checkAndWriteHeaderRow(sheet, sheetHeaders);

  for (const ts of allApiTimesheets) {
    const deputyTimesheetIdString = ts.Id.toString();
    if (existingTimesheetIDsInSheet.has(deputyTimesheetIdString)) {
      Logger.log(`Skipping duplicate Timesheet ID: ${deputyTimesheetIdString}`);
      continue;
    }

    let leaveTypeName = "";
    if (ts.IsLeave && ts.LeaveRule) {
      leaveTypeName = lookupLeaveRuleName(ts.LeaveRule, leaveRuleMap);
    } else if (ts.IsLeave) {
      leaveTypeName = "Leave (Rule ID Missing)";
    }
    
    const employeeObject = ts.EmployeeObject || {};
    const dpMetaData = ts._DPMetaData || {};
    const operationalUnitInfo = dpMetaData.OperationalUnitInfo || {};

    const rowData = [
      deputyTimesheetIdString,               // TimeSheet ID (from Deputy ts.Id)
      // --- CORRECTED MAPPING FOR EMPLOYEE EXPORT CODE ---
      employeeObject.Id ? employeeObject.Id.toString() : '', // Employee Export Code (from EmployeeObject.Id)
      employeeObject.DisplayName || '',      // Employee Name
      formatApiDateToSheetDate(ts.Date),     // Date
      formatApiTimeToSheetTime(ts.StartTimeLocalized), // Start
      formatApiTimeToSheetTime(ts.EndTimeLocalized),   // End
      formatDurationFromSeconds(calculateMealbreakDurationInSeconds(ts.Slots)), // Mealbreak
      ts.TotalTime || 0,                     // Total Hours
      ts.Cost || 0,                          // Total Cost
      ts.EmployeeComment || '',              // Employee Comment
      operationalUnitInfo.OperationalUnitName || '', // Area Name
      operationalUnitInfo.CompanyName || '', // Location Name
      leaveTypeName,                         // Leave
      ts.SupervisorComment || ''             // Manager's Comment
    ];
    rowsToAppend.push(rowData);
  }

  if (rowsToAppend.length > 0) {
    try {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, sheetHeaders.length)
           .setValues(rowsToAppend);
      Logger.log(`Successfully appended ${rowsToAppend.length} new timesheet entries to the sheet.`);
    } catch (e) {
        const sheetWriteErrorMsg = `Error writing to Google Sheet: ${e.toString()}. Stack: ${e.stack}`;
        Logger.log(sheetWriteErrorMsg);
        MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync Error - Sheet Write Failure", sheetWriteErrorMsg);
    }
  } else {
    Logger.log("No new timesheet entries to append (all were duplicates or no data fetched after filtering).");
  }
}

function mainDataFetchAndProcess_biWeeklyCheck() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const ADMIN_EMAIL = scriptProperties.getProperty('ADMIN_EMAIL_FOR_NOTIFICATIONS');

    // --- Define the very first "active" run date and time (EST) ---
    // We'll use UTC for calculations to avoid timezone issues in date math,
    // then convert the trigger's current time to a comparable format.
    const firstRunYear = 2025; // As per the example: June 4th, 2025
    const firstRunMonth = 5;   // 0-indexed for JavaScript Date (0=Jan, 5=June)
    const firstRunDay = 4;
    // No specific need for the hour here for the "is it the right week" check,
    // as the trigger handles the 4 AM run.

    // Create a UTC date for the first scheduled run day (ignoring time of day for this check)
    const firstScheduledRunDateUTC = new Date(Date.UTC(firstRunYear, firstRunMonth, firstRunDay));

    // Get current date in UTC (ignoring time of day for this check)
    const now = new Date();
    const currentRunDateUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));

    // Calculate the difference in days
    const millisecondsPerDay = 24 * 60 * 60 * 1000;
    const diffInTime = currentRunDateUTC.getTime() - firstScheduledRunDateUTC.getTime();
    const diffInDays = Math.floor(diffInTime / millisecondsPerDay);

    Logger.log(`Bi-weekly check: First scheduled run (UTC date): ${firstScheduledRunDateUTC.toUTCString()}`);
    Logger.log(`Bi-weekly check: Current run (UTC date): ${currentRunDateUTC.toUTCString()}`);
    Logger.log(`Bi-weekly check: Difference in days from first scheduled run: ${diffInDays}`);

    // Check if the current run date is on or after the first scheduled run
    // And if the number of days difference is a multiple of 14 (for bi-weekly)
    if (diffInDays >= 0 && diffInDays % 14 === 0) {
        Logger.log("This is an active Wednesday for the bi-weekly schedule. Proceeding with data fetch.");
        try {
            mainDataFetchAndProcess();
            // No need to store last run timestamp for *this specific* bi-weekly logic,
            // as it's based on a fixed schedule.
            Logger.log("mainDataFetchAndProcess completed successfully for scheduled bi-weekly run.");
        } catch (e) {
            const errorMessage = `Error during scheduled mainDataFetchAndProcess execution: ${e.toString()}. Stack: ${e.stack}`;
            Logger.log(errorMessage);
            if (ADMIN_EMAIL) {
                MailApp.sendEmail(ADMIN_EMAIL, "Deputy Sync CRITICAL ERROR", errorMessage);
            }
        }
    } else {
        Logger.log("Skipping mainDataFetchAndProcess: Not a scheduled bi-weekly run date.");
        if (diffInDays < 0) {
            Logger.log("Reason: Current date is before the first scheduled run date.");
        } else {
            Logger.log(`Reason: Day difference (${diffInDays}) is not a multiple of 14 from the first run date.`);
        }
    }
}

function setupTrigger() { // Renamed from setupBiWeeklyTrigger for clarity
  // Deletes all existing triggers for this script to avoid duplicates if re-run.
  // Be cautious if you have other triggers for other functions in this project.
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === "mainDataFetchAndProcess_biWeeklyCheck") {
        ScriptApp.deleteTrigger(existingTriggers[i]);
        Logger.log("Deleted existing trigger for mainDataFetchAndProcess_biWeeklyCheck.");
    }
  }

  // Create a new trigger
  ScriptApp.newTrigger('mainDataFetchAndProcess_biWeeklyCheck')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .atHour(4) // 4 AM
      .inTimezone("America/New_York") // Explicitly set EST/EDT
      .create();
  Logger.log("New trigger created for 'mainDataFetchAndProcess_biWeeklyCheck' to run every Wednesday at 4 AM EST/EDT.");
}