# N95 Fit Test Tracking — Excel Setup Guide

## Requirements

- **Microsoft Excel** (Windows version recommended — VBA + Outlook integration is most reliable on Windows).
- **Microsoft Outlook** installed and signed in on the same computer (the macro sends through Outlook's account).
- The file `N95_Fit_Test_Tracker.xlsx`.
- The VBA file `N95_FitTest_VBA.bas`.

> **Mac users:** Outlook for Mac has limited VBA support. The basic macro will work but may need small adjustments. If you're on Mac, tell me and I'll provide a Mac-compatible version.

---

## Step 1 — Open the Workbook in Excel

1. Locate `N95_Fit_Test_Tracker.xlsx` on your computer and double-click it.
2. Excel opens — you'll see the Dashboard plus 33 department tabs.
3. All formulas (Expiration Date, Validation Status) and conditional formatting work immediately. Try typing a Fit Test Date in any department sheet to confirm.

---

## Step 2 — Save as a Macro-Enabled Workbook

VBA macros only run in `.xlsm` (macro-enabled) workbooks, not in plain `.xlsx`.

1. **File → Save As**.
2. Choose a location.
3. From the **Save as type** dropdown, select **Excel Macro-Enabled Workbook (*.xlsm)**.
4. Click **Save**. Excel may ask whether to overwrite — confirm.
5. From now on, work in the `.xlsm` file. You can delete the original `.xlsx`.

---

## Step 3 — Enable the Developer Tab

If the **Developer** tab isn't already visible in the ribbon:

1. **File → Options → Customize Ribbon**.
2. In the right column, check the box next to **Developer**.
3. Click **OK**.

---

## Step 4 — Import the VBA Module

1. Press **Alt + F11** to open the VBA editor (or click **Developer → Visual Basic**).
2. In the left **Project Explorer** panel, right-click on `VBAProject (N95_Fit_Test_Tracker.xlsm)` → **Import File…**.
3. Browse to `N95_FitTest_VBA.bas` and click **Open**.
4. A new module called **N95Tracker** appears under the **Modules** folder. Double-click it to view the code.

---

## Step 5 — Edit the CONFIG

At the top of the module, edit these two lines:

```vba
Private Const INFECTION_CONTROL_EMAIL As String = "infectioncontrol@hospital.com"
Private Const HOSPITAL_NAME As String = "Hospital Name"
```

Replace with your actual values:

```vba
Private Const INFECTION_CONTROL_EMAIL As String = "ic.team@yourhospital.org"
Private Const HOSPITAL_NAME As String = "King Faisal Hospital"
```

Save the file with **Ctrl + S** (the file must already be saved as `.xlsm` — see Step 2).

---

## Step 6 — Trust the Workbook (Macro Security)

By default, Excel blocks macros from new files. To enable them for this file:

**Option A — Per-file trust (recommended):**
1. Close the workbook and reopen it.
2. A yellow **SECURITY WARNING — Macros have been disabled** bar appears at the top.
3. Click **Enable Content**.
4. From now on, Excel remembers this file as trusted.

**Option B — Trusted Location (for permanent trust):**
1. **File → Options → Trust Center → Trust Center Settings → Trusted Locations**.
2. Click **Add new location**, browse to the folder containing your `.xlsm` file, and check **Subfolders**.
3. Click **OK** twice.

---

## Step 7 — Test the System

1. In Excel, open any department sheet (e.g., **ER**) and add a test row in row 3:
   - Column A: `Test User`
   - Column B: `TEST-001`
   - Column C: your own email address
   - Column D: `3M 1860`
   - Column E: a date about 23 months ago, e.g., `2024-06-01`
2. Confirm column G turns yellow with text **Expiring Soon**.
3. Press **Alt + F8** to open the macro list.
4. Select **`SendSelfTest`** → click **Run**. Outlook should send a test email to yourself.
   > **Outlook security prompt:** the first time you run a macro that sends email, Outlook may pop up a dialog asking permission. Click **Allow** (and check "Allow access for 10 minutes" or longer).
5. Now run **`CheckN95FitTests`** (or **`TestRun`**). You should see a popup like:
   ```
   N95 daily check complete.
   Sent: 1
   Cooldown skipped: 0
   Errors: 0
   ```
6. Check Outlook → Sent Items, and the recipient inbox.
7. Delete the test row when done.

---

## Step 8 — Schedule Automatic Daily Runs

Excel doesn't have built-in time triggers. Choose one of these approaches:

### Option A — Run on file open (simplest)

The macro runs every time someone opens the workbook.

1. Press **Alt + F11**.
2. In the Project Explorer, double-click **ThisWorkbook** (under "Microsoft Excel Objects").
3. Paste this code:
   ```vba
   Private Sub Workbook_Open()
       ' Run the daily check whenever the file opens
       On Error Resume Next
       CheckN95FitTests
   End Sub
   ```
4. Save (Ctrl + S).

Now whenever the file is opened, it scans and sends alerts.

### Option B — Windows Task Scheduler (true automation)

Runs daily even if the file isn't open. Best for a shared computer that stays on.

1. Create a small VBS launcher file. Open Notepad, paste this:
   ```vbscript
   Set xl = CreateObject("Excel.Application")
   xl.Visible = False
   xl.DisplayAlerts = False
   Set wb = xl.Workbooks.Open("C:\Path\To\N95_Fit_Test_Tracker.xlsm")
   xl.Run "CheckN95FitTests"
   wb.Close False
   xl.Quit
   Set wb = Nothing
   Set xl = Nothing
   ```
2. Replace `C:\Path\To\...` with the actual full path of your `.xlsm` file.
3. Save the file as `RunN95Check.vbs` (in Notepad: **Save As** → set "Save as type" to **All Files** → filename `RunN95Check.vbs`).
4. Open **Task Scheduler** (Windows search: "Task Scheduler").
5. **Create Basic Task…** → Name it `N95 Daily Check`.
6. Trigger: **Daily**, start time `7:00 AM`.
7. Action: **Start a program** → Program/script: `wscript.exe` → Add arguments: `"C:\Path\To\RunN95Check.vbs"` (with quotes).
8. Finish. Right-click the task → **Run** to test it now.

> Task Scheduler runs even when the file is closed — make sure the computer is **on** at the scheduled time. For 24/7 reliability, run it on a workstation that stays on, or use a server.

### Option C — Outlook reminder (manual, no automation)

If automation is too complex, simply set yourself an Outlook recurring reminder to open the file and click **Run** on the macro every morning.

---

## Adding New Employees

Just type values into columns A-E of any department sheet. Columns F (Expiration Date) and G (Validation Status) calculate automatically.

## Adding a New Department

1. Right-click any department tab → **Move or Copy** → check **Create a copy** → click OK.
2. Rename the new tab.
3. On Dashboard, add a new row in the Breakdown table with COUNTIF formulas referencing the new sheet name. Update the TOTAL row's SUM ranges.
4. The macro automatically picks up the new sheet (it scans all sheets except the excluded ones).

## Changing Cooldown Period

In the VBA module, change:
```vba
Private Const COOLDOWN_DAYS As Integer = 7
```
to any number of days, or `0` to send every run.

## Clearing the Cooldown Cache

Press **Alt + F8** → run **`ResetAlertCache`**.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `Macros have been disabled` warning | Click **Enable Content** at the top of Excel. If missing, see Step 6 Option B. |
| `Run-time error '429': ActiveX component can't create object` | Outlook isn't installed, or COM access is blocked by IT. Use the SMTP version (ask me). |
| No email arrived, but macro says `Sent: 1` | Check Outlook → **Sent Items**, recipient's **Spam**, and Junk folders. Outlook may also be queuing in the Outbox if offline. |
| `Macro cannot be found` | Make sure you imported `N95_FitTest_VBA.bas` (Step 4) and the file is saved as `.xlsm` (Step 2). |
| Excel freezes during run | Outlook may be showing a security prompt off-screen. Click on Outlook in the taskbar and approve. |
| `Sent: 0` despite Dashboard showing Expiring Soon | Cooldown active — run **`ResetAlertCache`** then re-run. |
| Mac compatibility issues | Outlook for Mac VBA is limited. Tell me and I'll provide a Mac-compatible version using the macOS Mail app. |

---

## File Structure Reference

| File | Purpose |
|---|---|
| `N95_Fit_Test_Tracker.xlsx` | Original workbook — convert to `.xlsm` to use macros |
| `N95_FitTest_VBA.bas` | VBA module to import into Excel |

| Workbook Sheet | Purpose |
|---|---|
| Dashboard | Auto-aggregated KPIs and breakdown |
| AICU, IMCU, CCU, … (33 sheets) | One per department |
| README | Quick reference inside the workbook |
| `_AlertCache` | Hidden sheet the macro uses for cooldown tracking — don't edit |
