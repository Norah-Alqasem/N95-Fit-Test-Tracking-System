# N95 Fit Test Tracking System

An automated tracking system for hospital N95 mask fit test expirations, with dashboard analytics and email alerts. Built with Google Sheets + Apps Script (cloud) or Excel + VBA (offline).

## Why

OSHA and most hospital infection-control protocols require N95 fit tests to be repeated every 24 months. In a hospital with dozens of clinical units and hundreds of staff, tracking these renewals manually leads to missed expirations and compliance gaps. This system removes the manual work.

## Features

- **Auto-calculated expirations** — formula-based, never manually entered
- **Status flags** — Valid (>30 days), Expiring Soon (≤30 days), Expired (past date)
- **Color-coded dashboard** — KPIs and per-department breakdown across 33 clinical units
- **Conditional formatting** — green / yellow / red rows for instant visual scanning
- **Daily email alerts** — sent automatically to Infection Control and department leads
- **Cooldown system** — prevents duplicate alerts (default: 7-day window per employee)
- **Two implementations** — Google Sheets (cloud) or Excel (offline)
- **6,740 formulas, zero errors** — verified before delivery

## Screenshots

   ### Dashboard — Hospital-wide overview
   Real-time KPIs and per-department breakdown across all 33 clinical units.

   <img width="1026" height="956" alt="Dashboard" src="https://github.com/user-attachments/assets/d354f18a-cc11-4b55-bd19-06f527489edd" />


   ### Department sheet — Staff list
   Each department has its own sheet with auto-calculated expiration dates and color-coded status (green / yellow / red).
   
<img width="2010" height="617" alt="Staff-list" src="https://github.com/user-attachments/assets/9f73ef58-13ca-4930-8ab4-d84795cd518f" />

   ### Automated email alert
   Sent to Infection Control and the department lead 30 days before expiration.
 
   <img width="1250" height="1106" alt="Alert" src="https://github.com/user-attachments/assets/83ace975-64ec-44c6-bc25-49255e34a151" />

   ### Daily check execution
   Apps Script runs automatically every morning, scanning all sheets and sending alerts.

  <img width="2166" height="540" alt="Run" src="https://github.com/user-attachments/assets/a1c35b5a-ff5b-4afe-8104-6a209fc54245" />


## Tech Stack

| Component | Google Sheets version | Excel version |
|---|---|---|
| Spreadsheet | Google Sheets | Microsoft Excel (.xlsm) |
| Automation | Google Apps Script (JavaScript) | VBA |
| Email engine | Gmail (via `MailApp`) | Outlook (via VBA) |
| Scheduling | Apps Script Time-driven Trigger | Workbook_Open or Windows Task Scheduler |

## Repository Structure

```
.
├── N95_Fit_Test_Tracker.xlsx      # Workbook (33 departments + Dashboard + README)
├── N95_FitTest_AppsScript.gs      # Google Apps Script (cloud automation)
├── N95_FitTest_VBA.bas            # VBA module (Excel automation)
├── SETUP_STEPS.md                 # Google Sheets setup guide
├── SETUP_STEPS_EXCEL.md           # Excel setup guide
├── build_workbook.py              # Python script that generates the workbook
└── README.md                      # This file
```

## Departments Tracked

AICU, IMCU, CCU, CICU, SCBU, SICU, PICU, NICU, L&D, OR, Specialty Clinic, Dental, FMW, OB, NBN, Hemodialysis, Cardiac Center, PW, FSW, Endoscopy, HK, FMC, MMW, MSW, CW, Radiology, HHC, MS1, DAY CASE, LABORATORY, ER, RR, Physiotherapy

Each department has its own sheet with identical column structure. Add or remove departments as needed.

## Column Structure

| Col | Field | Type |
|---|---|---|
| A | Employee Name | Manual |
| B | Employee ID | Manual |
| C | Department Email | Manual (used for alerts) |
| D | N95 Mask Type | Manual |
| E | Fit Test Date | Manual (date) |
| F | Expiration Date | **Auto** — `=IF(E="","",EDATE(E,24))` |
| G | Validation Status | **Auto** — `Valid` / `Expiring Soon` / `Expired` |

## Quick Start (Google Sheets)

1. Upload `N95_Fit_Test_Tracker.xlsx` to Google Drive.
2. Right-click → Open with → Google Sheets → File → Save as Google Sheets.
3. Open **Extensions → Apps Script**, paste `N95_FitTest_AppsScript.gs`.
4. Edit the two CONFIG values at the top:
   ```javascript
   INFECTION_CONTROL_EMAIL: 'your-ic-team@hospital.com',
   HOSPITAL_NAME: 'Your Hospital',
   ```
5. Run `sendSelfTest` to authorize, then add a daily Time-driven Trigger on `checkN95FitTests`.

Full instructions: see [SETUP_STEPS.md](SETUP_STEPS.md).

## Quick Start (Excel)

1. Open `N95_Fit_Test_Tracker.xlsx` in Excel → Save As → Excel Macro-Enabled Workbook (`.xlsm`).
2. Press **Alt+F11** → right-click VBAProject → Import File → select `N95_FitTest_VBA.bas`.
3. Edit the two CONFIG constants in the module.
4. Reopen the file → click **Enable Content** on the security bar.
5. Press **Alt+F8** → run `SendSelfTest` to verify Outlook integration, then schedule via Workbook_Open or Windows Task Scheduler.

Full instructions: see [SETUP_STEPS_EXCEL.md](SETUP_STEPS_EXCEL.md).

## Configuration

Both versions share the same configuration concept:

| Setting | What it does | Default |
|---|---|---|
| `INFECTION_CONTROL_EMAIL` | Address that receives every alert | `infectioncontrol@hospital.com` |
| `HOSPITAL_NAME` | Appears in email header | `Hospital Name` |
| `COOLDOWN_DAYS` | Minimum days between alerts for the same employee | `7` |
| `EXCLUDED_SHEETS` | Sheets the script ignores | `Dashboard`, `README`, etc. |

## Status Logic

```
Expiration Date  =  Fit Test Date + 24 months           (EDATE)

Status           =  Expired         if Expiration Date < today
                    Expiring Soon   if Expiration Date - today <= 30 days
                    Valid           otherwise
```

## How the Alert Email Looks

Each alert is a styled HTML email containing:
- Employee name, ID, department, mask type
- Expiration date (highlighted)
- Days remaining
- Action prompt to coordinate with Infection Control

Recipients: department email (column C of the row) + Infection Control.

## Cooldown System

The script tracks the last sent timestamp per employee+expiration combo. By default, the same employee won't be alerted more than once every 7 days. Adjust via `COOLDOWN_DAYS`. Run `resetAlertCache` (Apps Script) or `ResetAlertCache` (VBA) to manually clear.

## Adding a New Department

1. Right-click any department tab → **Duplicate** → rename.
2. On the Dashboard, add a row to the breakdown table:
   ```
   =COUNTA('NewDept'!A3:A102)
   =COUNTIF('NewDept'!G3:G102,"Valid")
   =COUNTIF('NewDept'!G3:G102,"Expiring Soon")
   =COUNTIF('NewDept'!G3:G102,"Expired")
   ```
3. Update the TOTAL row's SUM ranges to include the new row.
4. The script automatically picks up the new sheet — no code changes needed.

## Limitations

- **Google Sheets version:** Free Gmail accounts have a 100-email/day limit; Google Workspace accounts get 1,500/day.
- **Excel version:** Requires Outlook installed and signed in on the same machine. Computer must be on for Task Scheduler runs.
- **Mac Excel:** Outlook for Mac VBA support is limited — use the Google Sheets version on Mac.
## Future Improvements

- [ ] Power BI integration for cross-hospital reporting
- [ ] Audit log sheet recording every alert sent
- [ ] Multi-language email templates (English/Arabic)
- [ ] Slack / Microsoft Teams webhook integration

## License

MIT — feel free to adapt this for your facility. Attribution appreciated but not required.

## Author

-  Norah Alqasem
-  Manar Alharthi
-  built as an internal tool for hospital infection control automation.

## Disclaimer

This system is a tracking aid, not a replacement for proper infection-control protocols. Always verify fit test compliance through your facility's official records before clinical use of N95 respirators.
