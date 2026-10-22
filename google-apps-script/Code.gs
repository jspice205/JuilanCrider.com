// ============================================================
// JC OPEN HOUSE LEADS — combined script
// Receives the open house sign-in form AND juliancrider.com leads,
// emails website leads, and serves mortgage rates to the site.
//
// Paste into the sheet's Extensions > Apps Script, run testLead once
// to approve permissions, then Deploy > Manage deployments > Edit >
// Version: New version > Deploy (keeps the same /exec URL).
// ============================================================

const SPREADSHEET_ID = '15iuFGQmmle8rrCMp1YLKhb2hUpGU23nA-dD71y-V6ws'; // JC Open House Leads
const NOTIFY_EMAIL = 'julianc@bcolv.com';

// ── TAB DEFINITIONS ───────────────────────────────────────
const TABS = [
  {
    name: 'Sign-Ins',
    color: '#1B3168',
    headers: ['Timestamp', 'Full Name', 'Phone', 'Email', 'Open House Address', 'Interest (1–10)',
              'Currently Renting or Owning?', 'Timeline to Buy', 'How They Heard', 'Agent Notes'],
    row: d => [d.name, d.phone, d.email, d.property, d.score, d.status, d.timeline, d.source, d.notes]
  },
  {
    name: 'Leads',
    color: '#C0392B',
    headers: ['Timestamp', 'Full Name', 'Phone', 'Email', 'Property Inquired About', 'Budget Range',
              'Timeline', 'Pre-Approved?', 'Bedrooms Needed', 'Neighborhoods Interested In', 'Notes'],
    row: d => [d.name, d.phone, d.email, d.property, d.budget, d.timeline, d.preapproved, d.bedrooms, d.neighborhoods, d.notes]
  },
  {
    name: 'Pipeline',
    color: '#2E7D32',
    headers: ['Timestamp', 'Full Name', 'Phone', 'Email', 'Source (Open House / IG / Referral)', 'Property Shown',
              'Interest Score', 'Follow-Up Date', 'Status', 'CRM Entered?', 'Notes'],
    row: d => [d.name, d.phone, d.email, d.source, d.property, d.score, d.followup, d.status, d.crm || 'No', d.notes]
  },
  {
    name: 'Website Leads',
    color: '#D13438',
    headers: ['Timestamp', 'Lead Type', 'Full Name', 'Phone', 'Email', 'Property', 'Timeline', 'Source', 'Notes'],
    row: d => [d.type, d.name, d.phone, d.email, d.property, d.timeline, d.source, d.notes]
  },
  {
    name: 'All Entries',
    color: '#6A1B9A',
    headers: ['Timestamp', 'Full Name', 'Phone', 'Email', 'Source Tab', 'Property', 'Score', 'Notes'],
    row: d => [d.name, d.phone, d.email, d.tab, d.property, d.score, d.notes]
  }
];
const tabDef = name => TABS.filter(t => t.name === name)[0];

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Creates a tab with styled headers if it doesn't exist yet.
function ensureTab(ss, tab, idx) {
  let sheet = ss.getSheetByName(tab.name);
  if (sheet) return sheet;
  sheet = idx == null ? ss.insertSheet(tab.name) : ss.insertSheet(tab.name, idx);
  styleHeader(sheet, tab);
  return sheet;
}

function styleHeader(sheet, tab) {
  sheet.getRange(1, 1, 1, tab.headers.length).setValues([tab.headers])
    .setBackground(tab.color).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, tab.headers.length);
  sheet.setTabColor(tab.color);
}

// ============================================================
// SETUP — run once from the editor (safe to re-run)
// ============================================================
function setupSheet() {
  const ss = getSS();
  const defaultSheet = ss.getSheetByName('Sheet1');
  TABS.forEach((tab, idx) => {
    const isNew = !ss.getSheetByName(tab.name);
    const sheet = ensureTab(ss, tab, idx);
    if (!isNew) styleHeader(sheet, tab);
    if (isNew) sheet.getRange(2, 1, 200, tab.headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  });
  if (defaultSheet && ss.getSheets().length > TABS.length) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }
  SpreadsheetApp.getUi().alert('✅ Setup Complete!\n\nTabs: ' + TABS.map(t => t.name).join(', '));
}

// ============================================================
// WEB APP — receives POST requests from landing pages and the website
// Deploy as a Web App (Execute as: Me, Anyone can access)
// ============================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSS();
    const timestamp = new Date().toLocaleString('en-US', {timeZone: 'America/Los_Angeles'});

    // ── Route to correct tab based on form type ──────────
    const tabName = data.tab || 'Sign-Ins';
    const tab = tabDef(tabName);
    if (!tab) throw new Error('Tab not found: ' + tabName);
    const sheet = ensureTab(ss, tab);
    sheet.appendRow([timestamp].concat(tab.row(data).map(cell)));

    // ── Also log to All Entries tab ───────────────────────
    if (tabName !== 'All Entries') {
      const allSheet = ss.getSheetByName('All Entries');
      const notes = tabName === 'Website Leads' ? [data.type, data.notes].filter(Boolean).join(' | ') : data.notes;
      if (allSheet) allSheet.appendRow([timestamp].concat([data.name, data.phone, data.email, tabName, data.property, data.score, notes].map(cell)));
    }

    // ── Email every website lead (open house sign-ins only when made on the website) ──
    if (tabName === 'Website Leads' || /juliancrider\.com/.test(data.notes || '')) {
      try { notify(tabName, data, timestamp, ss); } catch (err) { console.error('Email failed', err); }
      try { sendToFub(tabName, data); } catch (err) { console.error('Follow Up Boss failed', err); }
    }

    return json({status: 'success', tab: tabName});
  } catch (err) {
    return json({status: 'error', message: err.toString()});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'rates') return rates();
  return json({status: 'ok', message: 'JC Leads API is running'});
}

// Store visitor input as plain text so values like "+1 702…" or "=…" are never run as formulas.
function cell(v) { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ============================================================
// EMAIL NOTIFICATION
// ============================================================
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

// ============================================================
// FOLLOW UP BOSS — sends each website lead to FUB as a new lead event,
// so FUB assigns it and starts your action plans/automations.
// The API key lives in Project Settings > Script Properties as FUB_API_KEY
// (never in this code). If it's missing, this step is skipped.
// ============================================================
function fubEventType(tabName, d) {
  const t = (d.type || '').toLowerCase();
  if (tabName === 'Sign-Ins') return 'Visited Open House';
  if (/sell|home value/.test(t)) return 'Seller Inquiry';
  if (/new construction/.test(t)) return 'Property Inquiry';
  if (/account|guide/.test(t)) return 'Registration';
  return 'General Inquiry';
}

function sendToFub(tabName, d) {
  const key = PropertiesService.getScriptProperties().getProperty('FUB_API_KEY');
  if (!key) return;
  const parts = String(d.name || '').trim().split(/\s+/);
  const person = {firstName: parts.shift() || '', lastName: parts.join(' '), tags: ['juliancrider.com']};
  if (d.email) person.emails = [{value: d.email}];
  if (d.phone) person.phones = [{value: d.phone}];
  if (d.type) person.tags.push(d.type.replace(/^Consultation:\s*/, ''));
  const event = {
    source: 'juliancrider.com',
    system: 'juliancrider.com',
    type: fubEventType(tabName, d),
    message: [d.type, d.timeline && 'Timeline: ' + d.timeline, d.notes].filter(Boolean).join('\n'),
    person: person
  };
  if (d.property) event.property = {street: d.property};
  const res = UrlFetchApp.fetch('https://api.followupboss.com/v1/events', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(event), muteHttpExceptions: true,
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(key + ':')}
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('FUB ' + code + ': ' + res.getContentText().slice(0, 300));
  console.log('Sent to Follow Up Boss (' + code + ')');
}

/** Run once after adding FUB_API_KEY: creates a "FUB Test" lead in Follow Up Boss (delete it after). */
function testFub() {
  sendToFub('Website Leads', {type: 'Consultation: Test', name: 'FUB Test', email: 'fub-test@example.com',
    phone: '7025550100', timeline: 'Just exploring', notes: 'Test from the script editor — safe to delete'});
}

/** Run once from the editor to approve permissions and send yourself a test lead. */
function testLead() {
  const res = doPost({postData: {contents: JSON.stringify({tab: 'Website Leads', type: 'Consultation: Test', name: 'Test Lead',
    phone: '7025550100', email: '', property: '', timeline: 'Just exploring', source: 'juliancrider.com', notes: 'Test from the script editor'})}});
  console.log(res.getContent());
  console.log(rates().getContent());
}

// ============================================================
// MORTGAGE RATES for the site's "Mortgage rates today" card
// Mortgage News Daily's daily index, cached for 3 hours.
// Falls back to Freddie Mac's weekly survey (30 and 15 year only).
// ============================================================
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
