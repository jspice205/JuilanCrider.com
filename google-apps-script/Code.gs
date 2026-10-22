/**
 * Lead receiver for juliancrider.com and the open house sign-in.
 * Paste into the Apps Script project behind SHEET_URL in index.html,
 * then Deploy > Manage deployments > Edit > Version: New version > Deploy
 * (editing the existing deployment keeps the same /exec URL).
 */
const SPREADSHEET_ID = '15iuFGQmmle8rrCMp1YLKhb2hUpGU23nA-dD71y-V6ws'; // JC Open House Leads
const NOTIFY_EMAIL = 'julianc@bcolv.com';

const TABS = {
  'Sign-Ins': {
    headers: ['Timestamp', 'Full Name', 'Phone', 'Email', 'Open House Address', 'Interest (1–10)',
              'Currently Renting or Owning?', 'Timeline to Buy', 'How They Heard', 'Agent Notes'],
    row: d => [d.name, d.phone, d.email, d.property, d.score, d.status, d.timeline, d.source, d.notes]
  },
  'Website Leads': {
    headers: ['Timestamp', 'Lead Type', 'Full Name', 'Phone', 'Email', 'Property', 'Timeline', 'Source', 'Notes'],
    row: d => [d.type, d.name, d.phone, d.email, d.property, d.timeline, d.source, d.notes]
  }
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const tabName = TABS[d.tab] ? d.tab : 'Website Leads';
    const tab = TABS[tabName];
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(tab.headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, tab.headers.length).setFontWeight('bold');
    }
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy, h:mm:ss a');
    sheet.appendRow([stamp].concat(tab.row(d).map(cell)));

    // Email for every website lead; open house sign-ins only when they came from the website.
    if (tabName === 'Website Leads' || /juliancrider\.com/.test(d.notes || '')) notify(tabName, d, stamp, ss);
    return json({ok: true});
  } catch (err) {
    console.error(err);
    return json({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function notify(tabName, d, stamp, ss) {
  const type = d.type || (tabName === 'Sign-Ins' ? 'Open house sign-in' : 'Website lead');
  const fields = [
    ['Lead type', type], ['Name', d.name], ['Phone', d.phone], ['Email', d.email],
    ['Property', d.property], ['Timeline', d.timeline], ['Interest (1–10)', d.score],
    ['Renting or owning', d.status], ['Source', d.source], ['Notes', d.notes], ['Received', stamp]
  ].filter(f => f[1]);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const tel = d.phone ? `<a href="tel:${esc(String(d.phone).replace(/[^0-9+]/g, ''))}">Call ${esc(d.phone)}</a>` : '';
  const html = `<h2 style="font-family:sans-serif;margin:0 0 12px">New lead: ${esc(type)}</h2>
    <table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">
    ${fields.map(f => `<tr><td style="padding:6px 16px 6px 0;color:#666;vertical-align:top">${esc(f[0])}</td><td style="padding:6px 0">${esc(f[1])}</td></tr>`).join('')}
    </table><p style="font-family:sans-serif;font-size:14px">${tel}${tel ? ' · ' : ''}<a href="${ss.getUrl()}">Open the lead sheet</a></p>`;
  const msg = {
    to: NOTIFY_EMAIL,
    subject: `New lead: ${type}${d.name ? ' – ' + d.name : ''}`,
    htmlBody: html,
    body: fields.map(f => `${f[0]}: ${f[1]}`).join('\n'),
    name: 'juliancrider.com'
  };
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email || '')) msg.replyTo = d.email;
  MailApp.sendEmail(msg);
}

// Store visitor input as plain text so values like "+1 702…" or "=…" are never run as formulas.
function cell(v) { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function doGet() { return json({ok: true, service: 'juliancrider.com leads'}); }

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to approve permissions and send yourself a test lead. */
function testLead() {
  doPost({postData: {contents: JSON.stringify({tab: 'Website Leads', type: 'Consultation: Test', name: 'Test Lead',
    phone: '7025550100', email: '', property: '', timeline: 'Just exploring', source: 'juliancrider.com', notes: 'Test from the script editor'})}});
}
