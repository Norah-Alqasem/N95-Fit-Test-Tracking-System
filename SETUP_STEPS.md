# N95 Fit Test Tracking System — Setup Guide

## What's Already Done (Ready in the file)

- A complete workbook with **33 departments** + Dashboard + README sheet.
- All columns and formulas (Expiration Date + Validation Status) pre-populated for 100 rows per department.
- Conditional formatting fully applied: green / yellow / red for status, plus subtle row tinting for Expired and Expiring Soon rows.
- Dashboard auto-aggregates from all departments: 4 KPI cards + detailed breakdown table + TOTAL row.
- All formulas validated: **6,740 formulas with zero errors**.
- A ready-to-paste Apps Script `.gs` file with a built-in cooldown system to prevent duplicate alerts.

---

## What You Need to Do (≈ 5 minutes)

Google requires the account owner to perform these steps personally for security reasons.

---

### Step 1 — Upload to Google Sheets

1. Open [drive.google.com](https://drive.google.com).
2. Drag the file `N95_Fit_Test_Tracker.xlsx` into Drive (or click **+ New → File upload**).
3. Once uploaded, right-click the file → **Open with → Google Sheets**.
4. From the top menu: **File → Save as Google Sheets**.
5. You now have a native Google Sheet with all formulas and formatting preserved. You can delete the original `.xlsx` copy.

> All the formulas (`EDATE`, `IF`, `TODAY`, `COUNTA`, `COUNTIF`, `SUM`) work in Google Sheets without any modification.

---

### Step 2 — Add the Apps Script Code

1. Inside your Google Sheet, go to **Extensions → Apps Script**.
2. A new tab opens. Delete any default code shown (e.g., `function myFunction() {}`).
3. Open `N95_FitTest_AppsScript.gs` on your computer, copy the entire contents, and paste it into the Apps Script editor.
4. **Edit the CONFIG section at the top:**

   ```javascript
   const CONFIG = {
     INFECTION_CONTROL_EMAIL: 'YOUR_INFECTION_CONTROL@hospital.com',  // ← change this
     HOSPITAL_NAME: 'Your Hospital Name',                              // ← change this
     ...
   };
   ```

5. Click the **Save** icon (💾) at the top. Give the project a name like `N95 Tracker`.

---

### Step 3 — Authorize and Test

1. At the top of the Apps Script editor, in the function dropdown next to the Run button, select `sendSelfTest`.
2. Click **Run** (▶).
3. An **Authorization required** dialog appears → click **Review permissions**.
4. Choose your account → you'll see a "Google hasn't verified this app" warning → click **Advanced → Go to N95 Tracker (unsafe) → Allow**.
   > This warning is normal — it appears because you wrote the code yourself and haven't published it publicly. The code is safe; you can review every line.
5. Check your inbox: you should receive a test email titled "N95 Tracker — self-test OK".
6. Now switch the function dropdown to `testRun` and click **Run**. This scans every department and sends real alerts for any employee with status "Expiring Soon".

---

### Step 4 — Schedule the Daily Check

1. In the Apps Script editor, click the **Triggers** icon (clock ⏰) in the left sidebar.
2. Click **+ Add Trigger** (bottom right).
3. Fill in the form:

   | Field | Value |
   |---|---|
   | Choose which function to run | `checkN95FitTests` |
   | Choose which deployment should run | `Head` |
   | Select event source | `Time-driven` |
   | Select type of time based trigger | `Day timer` |
   | Select time of day | `7am to 8am` (or whenever you prefer) |
   | Failure notification settings | `Notify me immediately` |

4. Click **Save** → it may ask for permissions again — approve.

Done! The system now runs automatically every day.

---

## Important Notes

### Department Emails
For each employee row, make sure column **Department Email (C)** has a valid email address. If it's blank or invalid, the script skips that recipient and logs a note in the execution log.

### Cooldown System
- By default, the script will not send an alert for the same employee more than once every 7 days.
- To change the cooldown period, edit `COOLDOWN_DAYS` in the CONFIG section.
- To clear the cooldown cache and re-trigger pending alerts, run the `resetAlertCache` function manually.

### Adding a New Department
1. In Google Sheets: right-click any department tab → **Duplicate** → rename it.
2. On the Dashboard, insert a new row in the "Breakdown by Department" table:
   - Column A: department name
   - B: `=COUNTA('NewDept'!A3:A102)`
   - C: `=COUNTIF('NewDept'!G3:G102,"Valid")`
   - D: `=COUNTIF('NewDept'!G3:G102,"Expiring Soon")`
   - E: `=COUNTIF('NewDept'!G3:G102,"Expired")`
3. Update the TOTAL row's SUM ranges to include the new row.
4. The script will pick up the new sheet automatically — no code changes needed.

### If You Exceed 100 Employees per Department
The formulas and conditional formatting are pre-populated for 100 rows per sheet. To extend:
1. Select the last formula row, copy it (Ctrl+C).
2. Select the new range, paste (Ctrl+V).
3. On Dashboard, update the range `A3:A102` (and `G3:G102`) to a larger range in each formula.

### Security
- The script never modifies or deletes your data — it only reads cells and sends emails.
- Recipients are limited to: Infection Control + the Department Email in the same row. Nowhere else.
- You can disable the trigger at any time from the Triggers tab in Apps Script.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| No emails received | Run `testRun` manually and check the execution log (View → Logs). Confirm at least one employee has status "Expiring Soon". |
| Email reached Infection Control but not the department | Check that the Department Email column is filled correctly for that employee. |
| "Service invoked too many times" error | You hit Google's daily email quota (100/day for free Gmail, 1500/day for Workspace). Wait 24 hours. |
| Trigger not firing | In Apps Script, open the Executions tab — if there are errors, click them to see the cause. |
| Conditional formatting missing after upload | Rare. Open Format → Conditional formatting and you'll find the rules saved — just click Done. |
| Sheet name with spaces breaks formula | Make sure the sheet name in formulas is wrapped in single quotes, e.g., `'Specialty Clinic'!A3:A102`. |

---

## File Structure Reference

| Sheet | Purpose |
|---|---|
| **Dashboard** | Auto-aggregated view of all departments with KPIs and breakdown |
| **AICU, IMCU, CCU, ...** (33 sheets) | One per department, each with identical column structure |
| **README** | Quick-reference instructions inside the workbook |

### Column Structure (every department sheet)

| Col | Field | Type |
|---|---|---|
| A | Employee Name | Manual entry |
| B | Employee ID | Manual entry |
| C | Department Email | Manual entry (used for alerts) |
| D | N95 Mask Type | Manual entry |
| E | Fit Test Date | Manual entry (date) |
| F | Expiration Date | **Auto-calculated** (Fit Test Date + 24 months) |
| G | Validation Status | **Auto-calculated** (Valid / Expiring Soon / Expired) |
