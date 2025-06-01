<p align="center">
  <img src="logo.png" alt="DeputyDawg Logo" width="200"/>
</p>

# Deputy Timesheet Sync to Google Sheets

This project automates the process of fetching approved timesheets from the Deputy API and writing them to a specified Google Sheet. It is designed to run as a Google Apps Script project, with optional local development and deployment using clasp.

## Features
- Fetches all approved timesheets for a specified pay period from Deputy.
- Deduplicates timesheet entries based on Timesheet ID.
- Writes timesheet data to a Google Sheet with a customizable header.
- Handles leave rules and maps leave types.
- Sends error notifications to an admin email if configuration or API errors occur.
- Supports bi-weekly automated scheduling via Apps Script triggers.

## File Structure
- `Code.js` — Main script logic for fetching, processing, and writing timesheet data.
- `Helper Functions.js` — Utility functions for date/time formatting, leave rule mapping, deduplication, and sheet header management.

## Setup & Configuration

### 1. Prerequisites
- Google account with access to Google Apps Script and Google Sheets.
- Deputy API access (API token, install name, geo code).
- [Node.js](https://nodejs.org/) and [clasp](https://github.com/google/clasp) for local development (optional).

### 2. Script Properties (Required)
Set these in **Apps Script > Project Settings > Script Properties**:
- `DEPUTY_INSTALL`: Your Deputy install name (e.g., `mycompany`)
- `DEPUTY_GEO`: Your Deputy geo code (e.g., `na`, `au`, `eu`)
- `DEPUTY_ACCESS_TOKEN`: Your Deputy API token
- `DEPUTY_AUTH_TYPE`: `Bearer` or `DeputyKey` (default: `Bearer`)
- `SPREADSHEET_ID`: The ID of your target Google Sheet
- `SHEET_NAME`: The name of the tab to write to (e.g., `Time Card Data`)
- `TIMESHEET_ID_COLUMN_INDEX`: The zero-based column index for Timesheet IDs (default: `0`)
- `ADMIN_EMAIL_FOR_NOTIFICATIONS`: Email for error notifications

### 3. Local Development (Optional)
- Clone this repo and `cd` into the project directory.
- Install clasp globally if not already: `npm install -g @google/clasp`
- Authenticate clasp: `clasp login`
- Link to your Apps Script project: `clasp clone <scriptId>`
- Push changes: `clasp push`

### 4. Deploying & Running
- Open the project in Google Apps Script.
- Set all required script properties.
- Run `setupTrigger()` to schedule the bi-weekly sync (every other Wednesday at 4 AM EST/EDT).
- You can also run `mainDataFetchAndProcess()` manually for immediate sync.

## Scheduling & Automation
- The script uses `setupTrigger()` to create a time-based trigger for bi-weekly execution.
- The function `mainDataFetchAndProcess_biWeeklyCheck()` ensures the sync only runs on the correct Wednesdays.

## Error Handling & Notifications
- If required script properties are missing or API errors occur, the script logs the error and sends an email to the configured admin address.
- Sheet header mismatches and duplicate timesheet IDs are logged for troubleshooting.

## Customization
- You can adjust the sheet headers and mapping logic in `Code.js` as needed for your organization.
- Helper functions in `Helper Functions.js` can be extended for additional data processing.

## Contributing
Pull requests and issues are welcome! Please open an issue to discuss major changes before submitting a PR.

## License
[MIT](LICENSE) (or specify your license here)

---

*This project is not affiliated with Deputy. Use at your own risk. For questions or support, contact the project maintainer.* 