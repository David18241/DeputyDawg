<p align="center">
  <img src="logo.png" alt="DeputyDawg Logo" width="200"/>
</p>

# Deputy Timesheet Sync to Google Sheets

A Google Apps Script solution that automates the process of fetching approved timesheets from the Deputy API and syncing them to Google Sheets. Built for bi-weekly payroll processing with comprehensive error handling and deduplication.

## ✨ Features

- **Automated Timesheet Sync**: Fetches approved timesheets for specified pay periods from Deputy API
- **Dual Approval Support**: Handles both manually approved shifts and system-approved leave
- **Smart Deduplication**: Prevents duplicate entries based on Timesheet ID
- **Leave Type Mapping**: Processes various leave types with detailed breakdowns
- **Bi-weekly Automation**: Runs automatically every other Wednesday at 4 AM EST/EDT
- **Error Notifications**: Sends email alerts for configuration or API errors
- **Comprehensive Logging**: Detailed logs for troubleshooting and monitoring

## 📁 Project Structure

```
DeputyDawg/
├── Code.js                  # Main script logic for data fetching and processing
├── Helper Functions.js      # Utility functions for formatting and calculations
├── README.md               # This documentation
├── logo.png                # Project logo
├── appsscript.json         # Apps Script manifest
├── .clasp.json             # Clasp configuration for local development
└── .gitignore              # Git ignore rules
```

## 🚀 Quick Start

### Prerequisites
- Google account with Google Apps Script and Google Sheets access
- Deputy API credentials (API token, install name, geo code)
- Target Google Sheet created and configured

### 1. Deploy to Google Apps Script
1. Open [Google Apps Script](https://script.google.com)
2. Create a new project
3. Copy the contents of `Code.js` and `Helper Functions.js` into your project
4. Save and name your project

### 2. Configure Script Properties
In **Apps Script > Settings > Script Properties**, add:

| Property | Description | Example |
|----------|-------------|---------|
| `DEPUTY_INSTALL` | Your Deputy install name | `mycompany` |
| `DEPUTY_GEO` | Your Deputy geo code | `na`, `au`, or `eu` |
| `DEPUTY_ACCESS_TOKEN` | Your Deputy API token | `your-api-token` |
| `DEPUTY_AUTH_TYPE` | Authentication type | `Bearer` (default) or `DeputyKey` |
| `SPREADSHEET_ID` | Target Google Sheet ID | `1ABC...XYZ` |
| `SHEET_NAME` | Sheet tab name | `Time Card Data` |
| `TIMESHEET_ID_COLUMN_INDEX` | Column index for Timesheet IDs | `0` (A column) |
| `ADMIN_EMAIL_FOR_NOTIFICATIONS` | Email for error alerts | `admin@company.com` |

### 3. Set Up Automation
1. Run the `setupTrigger()` function to enable bi-weekly automation
2. Test manually with `mainDataFetchAndProcess()`

## 💻 Local Development (Optional)

For developers who prefer local development:

```bash
# Install clasp globally
npm install -g @google/clasp

# Authenticate with Google
clasp login

# Clone your Apps Script project
clasp clone <your-script-id>

# Make changes locally and push
clasp push

# Pull changes from Apps Script
clasp pull
```

## 📊 Data Processing

### Timesheet Types Handled
- **Regular Shifts**: Standard work hours with start/end times, meal breaks, and costs
- **Leave Requests**: Vacation, sick leave, and other approved time off
- **Manager Comments**: Supervisor notes and approvals

### Sheet Output Format
| Column | Description |
|--------|-------------|
| TimeSheet ID | Unique Deputy timesheet identifier |
| Employee Id | Deputy employee ID |
| Employee Name | Employee display name |
| Date | Work/leave date |
| Start | Start time (HH:MM:SS) |
| End | End time (HH:MM:SS) |
| Mealbreak | Total meal break duration |
| Total Hours | Total hours worked/taken |
| Total Cost | Labor cost |
| Employee Comment | Employee notes |
| Area Name | Work area/department |
| Location Name | Work location |
| Leave | Leave type (if applicable) |
| Manager's Comment | Supervisor comments |

## ⚙️ Advanced Configuration

### Pay Period Logic
The system calculates pay periods as 14-day spans (Monday to Sunday), with the end date being the most recent Sunday before the script runs.

### Error Handling
- **API Errors**: Logged and emailed to admin
- **Configuration Issues**: Validates all required properties
- **Sheet Errors**: Handles missing sheets or permission issues
- **Duplicate Detection**: Skips existing timesheet IDs

### Customization Options
- Adjust pay period calculation in `getPreviousPayPeriodDates()`
- Modify sheet headers in the main processing function
- Extend leave type mapping as needed
- Customize error notification templates

## 🔧 Troubleshooting

### Common Issues
1. **"Required Script Properties not set"**: Verify all properties are configured
2. **"Sheet not found"**: Check sheet name and permissions
3. **API errors**: Validate Deputy credentials and permissions
4. **No data returned**: Verify date ranges and approval statuses

### Debugging Tools
- Use `debugManagerComments()` to test specific functionality
- Check execution logs in Apps Script console
- Monitor email notifications for runtime errors

## 🤝 Contributing

Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This project is not affiliated with Deputy. Use at your own risk and ensure compliance with your organization's data handling policies.

---

**Questions or Support?** Open an issue or contact the project maintainer. 