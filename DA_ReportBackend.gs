/**
 * DA Church Monthly Report — Backend
 *
 * Google Apps Script — deploy as web app
 *
 * SETUP:
 * 1. Create a new Google Sheet called "DA Monthly Reports"
 * 2. Rename Sheet1 tab to "Submissions"
 * 3. Add headers in row 1 (the doPost will auto-create them on first submission if missing)
 * 4. Open Extensions > Apps Script, paste this code
 * 5. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the web app URL and paste it into the form's BACKEND_URL variable
 *
 * FILE: DA_ReportBackend.gs
 */

// ═══════════════════════════════════════
//  CONFIG
// ═══════════════════════════════════════
var SHEET_ID = '1lV2eqMUdTuwdvhNoHm-fqLiHGIl7OLVYawbezhXCrLM';
var COMPLIANCE_ALERT_EMAIL = 'mharper@northtexas.ag';
var ALERT_SUBJECT_PREFIX = 'COMPLIANCE ALERT';

// ═══════════════════════════════════════
//  doPost — receives form submissions
// ═══════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss = SHEET_ID
      ? SpreadsheetApp.openById(SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    var sheet = ss.getSheetByName('Submissions');
    if (!sheet) {
      sheet = ss.insertSheet('Submissions');
    }

    // Auto-create headers on first row if empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(getHeaders());
      sheet.getRange(1, 1, 1, getHeaders().length).setFontWeight('bold');
    }

    // Build the row
    var row = buildRow(data);
    sheet.appendRow(row);

    // Check compliance and send alert if needed
    checkCompliance(data);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════
//  doGet — serves CORS preflight / health check
// ═══════════════════════════════════════
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'DA Monthly Report Backend' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════
//  HEADERS & ROW BUILDER
// ═══════════════════════════════════════
function getHeaders() {
  return [
    'Timestamp',
    'Church Name',
    'Church Address',
    'Church Phone',
    'New Pastor',
    'Pastor First',
    'Pastor Last',
    'Pastor Email',
    'Pastor Phone',
    'Submitted By',
    'Submitter Email',
    'Sunday AM Adults',
    'Sunday AM Under 18',
    'Sunday AM Total',
    'Mid-Week Attendance',
    'Salvations',
    'Water Baptisms',
    'Spirit Baptisms',
    'Communion Times',
    'Total Income',
    'Total Expenses',
    'Checking Balance',
    'Savings Balance',
    'Special Events',
    'Missions Promotion',
    'Ministry Safe',
    'Aplos',
    'Insurance Issues',
    'Insurance Other',
    'Section Connection',
    'Presbyter Call',
    'Comments'
  ];
}

function buildRow(d) {
  var totalAttend = (parseInt(d.attendAdults) || 0) + (parseInt(d.attendMinors) || 0);
  return [
    new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' }),
    d.churchName || '',
    d.churchAddress || '',
    d.churchPhone || '',
    d.newPastor ? 'Yes' : 'No',
    d.pastorFirst || '',
    d.pastorLast || '',
    d.pastorEmail || '',
    d.pastorPhone || '',
    d.fillerName || '',
    d.fillerEmail || '',
    parseInt(d.attendAdults) || 0,
    parseInt(d.attendMinors) || 0,
    totalAttend,
    parseInt(d.attendMidweek) || 0,
    parseInt(d.salvations) || 0,
    parseInt(d.waterBaptisms) || 0,
    parseInt(d.spiritBaptisms) || 0,
    parseInt(d.communionCount) || 0,
    parseFloat(d.totalIncome) || 0,
    parseFloat(d.totalExpenses) || 0,
    parseFloat(d.checkingBalance) || 0,
    parseFloat(d.savingsBalance) || 0,
    d.specialEvents || '',
    d.missionsPromo || '',
    d.ministrySafe || '',
    d.aplos || '',
    Array.isArray(d.insurance) ? d.insurance.join(', ') : (d.insurance || ''),
    d.insuranceOther || '',
    d.sectionConnection || '',
    d.presbyCall || '',
    d.comments || ''
  ];
}

// ═══════════════════════════════════════
//  COMPLIANCE ALERT
// ═══════════════════════════════════════
function checkCompliance(data) {
  var alerts = [];

  if (data.ministrySafe === 'No') {
    alerts.push('• Ministry Safe: NOT in use');
  }
  if (data.aplos === 'No') {
    alerts.push('• Aplos: NOT in use');
  }

  if (alerts.length === 0) return; // All good, no email needed

  var churchName = data.churchName || 'Unknown Church';
  var pastorName = (data.pastorFirst || '') + ' ' + (data.pastorLast || '');
  var subject = ALERT_SUBJECT_PREFIX + ' — ' + churchName;

  var body = 'COMPLIANCE ALERT\n'
    + '===================================\n\n'
    + 'Church: ' + churchName + '\n'
    + 'Pastor: ' + pastorName.trim() + '\n'
    + 'Pastor Email: ' + (data.pastorEmail || 'Not provided') + '\n'
    + 'Pastor Phone: ' + (data.pastorPhone || 'Not provided') + '\n'
    + 'Report Date: ' + new Date().toLocaleDateString('en-US', { timeZone: 'America/Chicago' }) + '\n\n'
    + 'NON-COMPLIANT ITEMS:\n'
    + '-----------------------------------\n'
    + alerts.join('\n') + '\n\n'
    + 'This church requires follow-up to ensure compliance with DA minimum standards.\n\n'
    + '---\n'
    + 'Sent automatically by DA Church Monthly Report system';

  MailApp.sendEmail({
    to: COMPLIANCE_ALERT_EMAIL,
    subject: subject,
    body: body
  });
}
