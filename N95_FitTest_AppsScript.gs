/**
 * N95 Fit Test Expiration Alert System
 * --------------------------------------
 * Scans every department sheet daily and emails an alert
 * for any employee whose Validation Status is "Expiring Soon".
 *
 * Setup:
 *   1. Update CONFIG below (Infection Control email + hospital name).
 *   2. Run testRun() once and grant permissions when prompted.
 *   3. Add a Time-driven Daily Trigger on checkN95FitTests().
 */

// ============================================================
// CONFIG — edit these values for your hospital
// ============================================================
const CONFIG = {
  INFECTION_CONTROL_EMAIL: 'infectioncontrol@hospital.com',
  HOSPITAL_NAME: 'Hospital Name',

  // Sheets to skip (don't change unless you renamed them)
  EXCLUDED_SHEETS: ['Dashboard', 'README', 'Settings', 'Logs'],

  // Column positions (0-based). Change only if you reorder columns.
  COLUMNS: {
    NAME: 0,           // A
    ID: 1,             // B
    DEPT_EMAIL: 2,     // C
    MASK_TYPE: 3,      // D
    FIT_TEST_DATE: 4,  // E
    EXP_DATE: 5,       // F
    STATUS: 6          // G
  },

  // Send each alert at most once per N days to avoid spamming.
  // Set to 0 to send every run.
  COOLDOWN_DAYS: 7
};
// ============================================================

/**
 * Main scheduled function. Triggered daily.
 */
function checkN95FitTests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const sheets = ss.getSheets();
  const cache = loadAlertCache_();
  const now = new Date();

  let alertsSent = 0;
  let skippedCooldown = 0;
  const errors = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (CONFIG.EXCLUDED_SHEETS.indexOf(sheetName) !== -1) return;

    try {
      const data = sheet.getDataRange().getValues();
      if (data.length < 3) return; // header rows + at least one data row

      // Data starts at row 3 (rows 1=title, 2=headers)
      for (let i = 2; i < data.length; i++) {
        const row = data[i];
        const status = String(row[CONFIG.COLUMNS.STATUS]).trim();
        const name = row[CONFIG.COLUMNS.NAME];

        if (!name || status !== 'Expiring Soon') continue;

        const employee = {
          name: String(name).trim(),
          id: String(row[CONFIG.COLUMNS.ID]).trim(),
          deptEmail: String(row[CONFIG.COLUMNS.DEPT_EMAIL]).trim(),
          maskType: String(row[CONFIG.COLUMNS.MASK_TYPE]).trim(),
          expDate: row[CONFIG.COLUMNS.EXP_DATE],
          department: sheetName
        };

        const key = employee.department + '|' + employee.id + '|' + (employee.expDate ? new Date(employee.expDate).getTime() : '');
        const lastSent = cache[key];

        if (CONFIG.COOLDOWN_DAYS > 0 && lastSent) {
          const daysSince = (now - new Date(lastSent)) / (1000 * 60 * 60 * 24);
          if (daysSince < CONFIG.COOLDOWN_DAYS) {
            skippedCooldown++;
            continue;
          }
        }

        const ok = sendAlertEmail_(employee, tz);
        if (ok) {
          cache[key] = now.toISOString();
          alertsSent++;
        }
      }
    } catch (err) {
      errors.push(sheetName + ': ' + err.message);
    }
  });

  saveAlertCache_(cache);

  Logger.log('N95 daily check complete | sent: ' + alertsSent +
             ' | cooldown skipped: ' + skippedCooldown +
             ' | errors: ' + errors.length);
  if (errors.length) Logger.log('Errors: ' + errors.join(' || '));
}

/**
 * Send a single alert email.
 */
function sendAlertEmail_(emp, tz) {
  try {
    if (!emp.deptEmail || emp.deptEmail.indexOf('@') === -1) {
      Logger.log('Skipped ' + emp.name + ' (' + emp.department + '): missing/invalid Department Email');
      return false;
    }

    const expDate = new Date(emp.expDate);
    const formattedDate = Utilities.formatDate(expDate, tz, 'dd MMM yyyy');
    const daysRemaining = Math.max(0, Math.ceil((expDate - new Date()) / (1000 * 60 * 60 * 24)));

    const subject = 'N95 Fit Test Expiring Soon — ' + emp.name + ' (' + emp.department + ')';

    const htmlBody =
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:620px;margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">' +
        '<div style="background-color:#f39c12;color:#ffffff;padding:20px;text-align:center;">' +
          '<h2 style="margin:0;font-size:20px;">N95 Fit Test Expiration Alert</h2>' +
          '<p style="margin:6px 0 0;font-size:13px;">' + escapeHtml_(CONFIG.HOSPITAL_NAME) + ' — Infection Control</p>' +
        '</div>' +
        '<div style="padding:24px;background-color:#ffffff;color:#333;">' +
          '<p style="font-size:15px;margin:0 0 12px;">Dear Department Lead and Infection Control Team,</p>' +
          '<p style="font-size:15px;margin:0 0 16px;">' +
            'This is an automated reminder that the N95 mask fit test for the following employee is ' +
            '<strong style="color:#e67e22;">expiring within ' + daysRemaining + ' day(s)</strong>. ' +
            'Please schedule a re-test to ensure continued compliance and staff safety.' +
          '</p>' +
          '<table style="width:100%;border-collapse:collapse;margin:18px 0;font-size:14px;">' +
            row_('Employee Name', emp.name, true) +
            row_('Employee ID', emp.id, false) +
            row_('Department', emp.department, true) +
            row_('N95 Mask Type', emp.maskType, false) +
            '<tr style="background-color:#fff3cd;">' +
              '<td style="padding:12px;border:1px solid #dee2e6;font-weight:bold;width:40%;">Expiration Date</td>' +
              '<td style="padding:12px;border:1px solid #dee2e6;color:#856404;font-weight:bold;">' + escapeHtml_(formattedDate) + '</td>' +
            '</tr>' +
          '</table>' +
          '<p style="font-size:14px;background-color:#e7f3fe;padding:12px;border-left:4px solid #2196F3;border-radius:4px;margin:0 0 12px;">' +
            '<strong>Action Required:</strong> Coordinate with Infection Control to re-schedule the fit test before the expiration date.' +
          '</p>' +
          '<p style="font-size:12px;color:#999;margin-top:24px;border-top:1px solid #eee;padding-top:12px;">' +
            'Automated notification from the N95 Fit Test Tracking System. Sent on ' +
            Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm') + '.' +
          '</p>' +
        '</div>' +
      '</div>';

    const recipients = [emp.deptEmail, CONFIG.INFECTION_CONTROL_EMAIL].filter(Boolean).join(',');

    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlBody,
      name: 'N95 Fit Test Tracking System'
    });

    Logger.log('Alert sent: ' + emp.name + ' (' + emp.department + ')');
    return true;
  } catch (err) {
    Logger.log('Failed for ' + emp.name + ': ' + err.message);
    return false;
  }
}

function row_(label, value, alt) {
  const bg = alt ? '#f8f9fa' : '#ffffff';
  return '<tr style="background-color:' + bg + ';">' +
           '<td style="padding:12px;border:1px solid #dee2e6;font-weight:bold;width:40%;">' + escapeHtml_(label) + '</td>' +
           '<td style="padding:12px;border:1px solid #dee2e6;">' + escapeHtml_(value || '—') + '</td>' +
         '</tr>';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Cooldown cache — uses Script Properties so it survives across runs.
 */
function loadAlertCache_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('N95_ALERT_CACHE');
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    return {};
  }
}

function saveAlertCache_(cache) {
  // Trim entries older than 90 days to keep the property small.
  const cutoff = Date.now() - (90 * 24 * 60 * 60 * 1000);
  const trimmed = {};
  Object.keys(cache).forEach(k => {
    if (new Date(cache[k]).getTime() >= cutoff) trimmed[k] = cache[k];
  });
  PropertiesService.getScriptProperties().setProperty('N95_ALERT_CACHE', JSON.stringify(trimmed));
}

/**
 * Manual test runner — execute this once after first install
 * to grant required permissions and verify it works.
 */
function testRun() {
  Logger.log('=== Manual test run starting ===');
  checkN95FitTests();
  Logger.log('=== Test run complete. Check execution log above. ===');
}

/**
 * Optional helper: send a test email to yourself only,
 * to confirm MailApp permissions without spamming departments.
 */
function sendSelfTest() {
  const me = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: me,
    subject: 'N95 Tracker — self-test OK',
    htmlBody: '<p>Apps Script email permissions are working correctly. You can now run testRun().</p>'
  });
  Logger.log('Self-test email sent to ' + me);
}

/**
 * Reset the cooldown cache (use if you want to re-trigger pending alerts).
 */
function resetAlertCache() {
  PropertiesService.getScriptProperties().deleteProperty('N95_ALERT_CACHE');
  Logger.log('Alert cache cleared.');
}
