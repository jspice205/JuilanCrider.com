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

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'rates') return rates();
  return json({ok: true, service: 'juliancrider.com leads'});
}

/* ---------- Mortgage rates for the site's "Mortgage rates today" card ---------- */
// Mortgage News Daily's daily index, cached for 3 hours. Falls back to Freddie Mac's weekly survey (30 and 15 year only).
function rates() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('rates');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  let out = null;
  try { out = mndRates(); } catch (err) { console.warn('MND', err); }
  if (!out) try { out = freddieRates(); } catch (err) { console.warn('Freddie Mac', err); }
  if (!out) return json({ok: false});
  const text = JSON.stringify(out);
  cache.put('rates', text, 3 * 60 * 60);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function mndRates() {
  const res = UrlFetchApp.fetch('https://www.mortgagenewsdaily.com/mortgage-rates/mnd', {muteHttpExceptions: true, headers: {'User-Agent': 'Mozilla/5.0'}});
  if (res.getResponseCode() !== 200) return null;
  const text = res.getContentText()
    .replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
  const products = {f30: '30 ?Yr\\.? Fixed', f15: '15 ?Yr\\.? Fixed', fha30: '30 ?Yr\\.? FHA', va30: '30 ?Yr\\.? VA', jumbo30: '30 ?Yr\\.? Jumbo'};
  const out = {};
  Object.keys(products).forEach(k => {
    const m = text.match(new RegExp(products[k] + '(?:(?!Yr)[^%]){0,40}?(\\d{1,2}\\.\\d{2})%\\s*([+-]?\\d+\\.\\d{2})?', 'i'));
    if (m) out[k] = {rate: +m[1], change: m[2] ? +m[2] : null};
  });
  if (!out.f30) return null;
  return {ok: true, source: 'Mortgage News Daily', asOf: Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMM d, yyyy'), rates: out};
}

function freddieRates() {
  const csv = UrlFetchApp.fetch('https://www.freddiemac.com/pmms/docs/PMMS_history.csv', {muteHttpExceptions: true}).getContentText().trim().split(/\r?\n/);
  const head = csv[0].split(','), last = csv[csv.length - 1].split(',');
  const col = name => last[head.indexOf(name)];
  if (!col('pmms30')) return null;
  return {ok: true, source: 'Freddie Mac weekly survey', asOf: last[0],
    rates: {f30: {rate: +col('pmms30'), change: null}, f15: {rate: +col('pmms15'), change: null}}};
}

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to approve permissions and send yourself a test lead. */
function testLead() {
  doPost({postData: {contents: JSON.stringify({tab: 'Website Leads', type: 'Consultation: Test', name: 'Test Lead',
    phone: '7025550100', email: '', property: '', timeline: 'Just exploring', source: 'juliancrider.com', notes: 'Test from the script editor'})}});
}
