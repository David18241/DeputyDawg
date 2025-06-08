// HelperFunctions.gs

/**
 * Formats an API date string (e.g., "2025-05-21T00:00:00-04:00") to "MM/DD/YYYY".
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
 * Formats an API time string (e.g., "2025-05-21T06:43:00-04:00" OR "2025-02-19 09:00:00") to "HH:MM:SS".
 */
function formatApiTimeToSheetTime(apiTimeStr) {
  if (!apiTimeStr) return '';
  try {
    const separator = apiTimeStr.includes('T') ? 'T' : ' ';
    const parts = apiTimeStr.split(separator);
    if (parts.length > 1 && parts[1]) {
      return parts[1].substring(0, 8);
    } else {
      return '';
    }
  } catch (e) {
    Logger.log(`Error formatting time string: "${apiTimeStr}". Error: ${e}`);
    return apiTimeStr;
  }
}

/**
 * Calculates total meal break duration in seconds from the Slots array.
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
 */
function formatDurationFromSeconds(totalSeconds) {
  if (totalSeconds === null || totalSeconds === undefined || isNaN(totalSeconds)) return '0:00:00';
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = Math.round(totalSeconds % 60);
  return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
}

/**
 * Calculates the start and end dates for the target pay period.
 * The period is a 14-day span (Monday to Sunday). The end date is the most recent Sunday.
 */
function getPreviousPayPeriodDates() {
  const scriptRunDate = new Date();
  const dayOfWeek = scriptRunDate.getDay(); // Sunday = 0, Monday = 1, ..., Saturday = 6

  const payPeriodEndDate = new Date(scriptRunDate);
  payPeriodEndDate.setDate(scriptRunDate.getDate() - dayOfWeek);
  
  const payPeriodStartDate = new Date(payPeriodEndDate);
  payPeriodStartDate.setDate(payPeriodEndDate.getDate() - 13);
  
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
 */
function getExistingDeputyIDsFromSheet(sheet, idColumnIndex) {
  const ids = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;
  const range = sheet.getRange(2, idColumnIndex + 1, lastRow - 1, 1);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0]) {
      ids.add(values[i][0].toString());
    }
  }
  return ids;
}

/**
 * Checks if a header row exists and writes it if the sheet is empty.
 */
function checkAndWriteHeaderRow(sheet, expectedHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    Logger.log("Appended header row to empty sheet.");
  }
}