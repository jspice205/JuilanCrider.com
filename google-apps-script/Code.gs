// ============================================================
// JC OPEN HOUSE LEADS — combined script
// Receives the open house sign-in form AND juliancrider.com leads,
// emails website leads, and serves mortgage rates to the site.
//
// Paste into the sheet's Extensions > Apps Script, run testLead once
// to approve permissions, then Deploy > Manage deployments > Edit >
// Version: New version > Deploy (keeps the same /exec URL).
// After setupAutoUpdate (see SELF-UPDATE below), it updates itself from GitHub
// every hour, so pasting is only needed this one last time.
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
    if (e.parameter && e.parameter.bb) return bbWebhook(e);  // BlueBubbles (texting Mac) events
    if (e.parameter && e.parameter.vapi) return vapiWebhook(e);  // Shawna phone calls (Vapi end-of-call reports)
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'signAgreement') return signAgreement(data);
    if (data.action === 'book') return json(bookSlot(data));
    if (data.action === 'enrich') return json(enrichSignIn(data));
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
    }
    // ── Showing requests: send the buyer a confirmation right away ──
    if (tabName === 'Website Leads' && /showing/i.test(data.type || '') && validEmail(data.email)) {
      let signUrl = '';
      // The one-home agreement needs an address; without one, Julián sends it once they pick a home.
      if (data.property && agreementSettings().live && !/Working with agent: Yes, with a signed/.test(data.notes || '')) {
        try { signUrl = createAgreement(data); } catch (err) { console.error('Agreement not created', err); }
      }
      try { confirmShowing(data, signUrl); } catch (err) { console.error('Buyer confirmation failed', err); }
    }
    // ── Every website lead and every open house sign-in goes to Follow Up Boss ──
    if (tabName === 'Website Leads' || tabName === 'Sign-Ins') {
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
  if (e && e.parameter && e.parameter.action === 'agreement') return getAgreement(String(e.parameter.id || ''));
  if (e && e.parameter && e.parameter.action === 'slots') return json({status: 'ok', slots: bookingSlots().map(d => d.toISOString())});
  return json({status: 'ok', message: 'JC Leads API is running'});
}

// ============================================================
// BOOKING — "Book a consultation" on the site shows open times from your Google Calendar and books them.
// Anything timed on your main calendar (Tao shifts, showings, appointments) blocks that time. Change hours here.
// ============================================================
const BOOKING = {days: 10, minutes: 30, startHour: 9, endHour: 18, weekdays: [1, 2, 3, 4, 5, 6], noticeHours: 3, tz: 'America/Los_Angeles'};

function bookingSlots() {
  const now = new Date(), until = new Date(now.getTime() + (BOOKING.days + 1) * 864e5);
  const busy = CalendarApp.getDefaultCalendar().getEvents(now, until).filter(ev => !ev.isAllDayEvent())
    .map(ev => [ev.getStartTime().getTime(), ev.getEndTime().getTime()]);
  const earliest = now.getTime() + BOOKING.noticeHours * 36e5, len = BOOKING.minutes * 6e4, slots = [];
  for (let d = 0; d < BOOKING.days; d++) {
    const day = new Date(now.getTime() + d * 864e5);
    if (BOOKING.weekdays.indexOf(Number(Utilities.formatDate(day, BOOKING.tz, 'u')) % 7) < 0) continue;  // u: 1 = Monday … 7 = Sunday
    const ymd = Utilities.formatDate(day, BOOKING.tz, 'yyyy-MM-dd');
    for (let m = BOOKING.startHour * 60; m + BOOKING.minutes <= BOOKING.endHour * 60; m += BOOKING.minutes) {
      const start = Utilities.parseDate(`${ymd} ${Math.floor(m / 60)}:${m % 60}`, BOOKING.tz, 'yyyy-MM-dd H:m'), t = start.getTime();
      if (t >= earliest && !busy.some(b => b[0] < t + len && b[1] > t)) slots.push(start);
    }
  }
  return slots;
}

function bookSlot(d) {
  const slot = bookingSlots().filter(s => s.toISOString() === d.slot)[0];
  if (!slot) return {status: 'taken', message: 'That time was just taken. Please pick another.'};
  if (!String(d.name || '').trim() || !String(d.phone || '').trim()) return {status: 'error', message: 'Name and phone are required.'};
  const end = new Date(slot.getTime() + BOOKING.minutes * 6e4), email = validEmail(d.email) ? String(d.email).trim() : '';
  const label = Utilities.formatDate(slot, BOOKING.tz, "EEE, MMM d 'at' h:mm a");
  const showing = d.kind === 'showing' && d.property;
  CalendarApp.getDefaultCalendar().createEvent(showing ? `Showing: ${d.name}` : `Consultation: ${d.name} (${d.type || 'juliancrider.com'})`, slot, end, {
    description: [`Booked on juliancrider.com`, `Phone: ${d.phone}`, email && `Email: ${email}`, d.notes].filter(Boolean).join('\n'),
    location: showing ? String(d.property) : '', guests: email, sendInvites: !!email});
  try {  // log it on their Follow Up Boss contact (the lead itself arrives through the normal website form)
    const phone = String(d.phone).replace(/\D/g, '');
    const found = (email && fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people) ||
                  (phone && fubFetch('people?limit=1&phone=' + phone, null, 'get').people) || [];
    if (found.length) {
      fubFetch('notes', {personId: found[0].id, subject: showing ? 'Booked a showing' : 'Booked a consultation',
        body: `${label} (${BOOKING.minutes} min)${showing ? ' at ' + d.property : ''}, booked on juliancrider.com`});
      createCallTask(found[0].id, showing ? `showing ${label} at ${d.property}` : `booked consultation ${label}`);
    }
  } catch (err) { console.error('Booking note: ' + err); }
  try {  // confirmation text from Julián's number (the 15-minute reminder comes from remindUpcoming)
    const first = String(d.name).trim().split(/\s+/)[0];
    bbSend(d.phone, showing
      ? `Hey ${first}, it's Julián 👋🏾 You're all set to see ${String(d.property).split(',')[0]} ${label}. I'll meet you there. If anything changes, just text me here.`
      : `Hey ${first}, it's Julián 👋🏾 You're all set for ${label}. Looking forward to it! If anything changes, just text me here.`);
  } catch (err) { console.log('Booking confirmation text not sent: ' + err.message); }
  return {status: 'success', label: label};
}

// ============================================================
// SELF-UPDATE — after the one-time setup, this script keeps itself current with GitHub, so you never paste again.
//   One time: turn on "Google Apps Script API" at script.google.com/home/usersettings, then run setupAutoUpdate.
//   It checks GitHub every hour, installs the new Code.gs, and points the website's deployment at it.
//   If an update needs new Google permissions, it emails you instead: open Apps Script and run updateFromGitHub once.
// ============================================================
const GITHUB_RAW = 'https://raw.githubusercontent.com/jspice205/JuilanCrider.com/main/google-apps-script/';

function setupAutoUpdate() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoUpdateFromGitHub').forEach(t => ScriptApp.deleteTrigger(t));
  try { console.log(updateFromGitHub()); } catch (err) {
    // The script's built-in Google project can't use the Apps Script API; it has to be linked to a project you own.
    throw new Error('Auto-update is not on, and nothing was changed. Link this script to your own Google Cloud project ' +
      'with the Apps Script API enabled (Project Settings > Google Cloud Platform (GCP) Project), then run setupAutoUpdate again. ' +
      'Details: ' + String(err).slice(0, 200));
  }
  ScriptApp.newTrigger('autoUpdateFromGitHub').timeBased().everyHours(1).create();  // only once updates actually work
}

function autoUpdateFromGitHub() {
  console.log(updateFromGitHub(true));
  try { ensureTextQueueTrigger(); } catch (e) { console.log('Text triggers: ' + e.message); }
  try { trackEmailOpens(); } catch (e) { console.log('Email opens: ' + e.message); }
  try {  // one-time: put the current signature footer on the plain-text drip emails
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('EMAIL_FOOTER') !== EMAIL_FOOTER_VERSION) { applyEmailFooter(); props.setProperty('EMAIL_FOOTER', EMAIL_FOOTER_VERSION); }
  } catch (e) { console.log('Email footer: ' + e.message); }
}

// EMAIL OPENS — hourly (from autoUpdateFromGitHub). Reads FUB's email events (opens, clicks, unsubscribes from your
// action-plan emails) and tags your contacts "Opened Email" / "Clicked Email", so a Smart List on those tags shows
// everyone who has engaged. The "Email Opens" sheet tab keeps one row per person + email + event with first/last
// date and count, so you can see engagement grow over time. Tags are only added, never removed.
const EMAIL_OPENS_TAB = {name: 'Email Opens', color: '#0078d4',
  headers: ['Last', 'First', 'Name', 'Event', 'Email (campaign)', 'Count', 'FUB link', 'Person id']};
function trackEmailOpens() {
  const props = PropertiesService.getScriptProperties();
  const seen = props.getProperty('EM_EVENTS_SEEN') || '';
  const me = fubFetch('me', null, 'get').id;
  const events = []; let next = '';
  for (let i = 0; i < 100; i++) {
    const r = fubFetch('emEvents?limit=100' + (next ? '&next=' + encodeURIComponent(next) : ''), null, 'get');
    events.push(...(r.emEvents || []));
    next = r._metadata && r._metadata.next; if (!next) break;
  }
  const ss = getSS(); const sheet = ss.getSheetByName(EMAIL_OPENS_TAB.name) || ensureTab(ss, EMAIL_OPENS_TAB);
  const old = {}; sheet.getDataRange().getValues().slice(1).forEach(r => { if (r[7]) old[r[7] + '|' + r[4] + '|' + r[3]] = r; });
  const people = {}, fmt = iso => Utilities.formatDate(new Date(iso), 'America/Los_Angeles', 'M/d/yyyy h:mm a');
  const person = id => people[id] || (people[id] = (() => { try { return fubFetch('people/' + id + '?fields=id,name,tags,assignedUserId', null, 'get'); } catch (e) { return null; } })());
  let latest = seen, tagged = 0; const rows = [];
  events.sort((a, b) => a.updated < b.updated ? 1 : -1);
  for (const e of events) {
    const key = e.personId + '|' + e.campaignName + '|' + e.type, prev = old[key];
    let name = prev && prev[2];
    if (e.updated > seen && /open|click/.test(e.type)) {
      const p = person(e.personId);
      if (p && (!p.assignedUserId || p.assignedUserId === me)) {
        name = p.name; const want = ['Opened Email'].concat(e.type === 'click' ? ['Clicked Email'] : []);
        const tags = p.tags || [], add = want.filter(t => tags.indexOf(t) < 0);
        if (add.length) { try { fubFetch('people/' + p.id, {tags: tags.concat(add)}, 'put'); p.tags = tags.concat(add); tagged++; Utilities.sleep(120); } catch (err) { console.log('Tag ' + p.id + ': ' + err.message); } }
      }
    }
    if (!name) { const p = person(e.personId); name = p ? p.name : ''; }
    if (e.updated > latest) latest = e.updated;
    rows.push([fmt(e.updated), fmt(e.created), name, e.type, e.campaignName, e.count, 'https://app.followupboss.com/2/people/view/' + e.personId, e.personId].map(cell));
  }
  sheet.clearContents(); styleHeader(sheet, EMAIL_OPENS_TAB);
  if (rows.length) sheet.getRange(2, 1, rows.length, EMAIL_OPENS_TAB.headers.length).setValues(rows);
  if (latest) props.setProperty('EM_EVENTS_SEEN', latest);
  console.log(`Email events: ${rows.length} rows, tagged ${tagged} people.`);
}

// Signature footer for the plain-text drip emails (Hot/Warm/Legacy 141-152, Nurture 153-176, Farm 177-180).
// Small and simple on purpose: it keeps the personal look but has the brokerage, address, license and opt-out the law
// requires. Change the text and bump EMAIL_FOOTER_VERSION; the hourly update re-applies it to every template.
const EMAIL_FOOTER_VERSION = 'v2';
const EMAIL_FOOTER = '<!--jc-footer--><table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:22px;border-top:1px solid #e5e5e5;padding-top:12px"><tr>' +
  '<td valign="top" style="padding:12px 12px 0 0"><img src="https://juliancrider.com/emails/img/julian-sq.jpg" width="56" height="56" alt="Julián Crider" style="display:block;width:56px;height:56px;border-radius:50%;border:0"></td>' +
  '<td valign="top" style="padding-top:12px;font:13px/1.55 Helvetica,Arial,sans-serif;color:#111">' +
  '<b>Julián J. Crider</b><br>REALTOR® · Developer · Advisor<br>The DeBenedetto Team at Boeckle &amp; Co.<br>' +
  '<a href="tel:+17024982571" style="color:#111">702-498-2571</a> · <a href="https://juliancrider.com" style="color:#111">juliancrider.com</a> · ' +
  '<a href="https://www.instagram.com/juliancriderestates/" style="color:#111">Instagram</a></td></tr></table>' +
  '<p style="margin:8px 0 0 0;font:11px/1.5 Helvetica,Arial,sans-serif;color:#888">4775 W Teco Ave #220, Las Vegas, NV 89118 · NV S.0177760<br>' +
  'Prefer fewer emails? Just reply "stop."</p>';

function applyEmailFooter() {
  let n = 0;
  for (let id = 141; id <= 180; id++) {
    try {
      const t = fubFetch('templates/' + id, null, 'get');
      let b = String(t.body || '');
      const mark = b.indexOf('<!--jc-footer-->');
      if (mark >= 0) b = b.slice(0, mark);
      else { const i = b.lastIndexOf('<p>Julián'); if (i >= 0) b = b.slice(0, i); }
      b = b.replace(/(<p>Julián<\/p>)?\s*$/, '') + '<p>Julián</p>' + EMAIL_FOOTER;
      fubFetch('templates/' + id, {name: t.name, subject: t.subject, body: b}, 'put');
      n++; Utilities.sleep(250);
    } catch (err) { console.log('Footer not applied to template ' + id + ': ' + err.message); }
  }
  console.log('Footer applied to ' + n + ' templates.');
}

function updateFromGitHub(auto) {
  const raw = f => UrlFetchApp.fetch(GITHUB_RAW + f + '?nocache=' + Date.now()).getContentText();
  const code = raw('Code.gs'), manifest = raw('appsscript.json');
  const id = ScriptApp.getScriptId(), base = 'https://script.googleapis.com/v1/projects/' + id;
  const api = (method, path, body) => JSON.parse(UrlFetchApp.fetch(base + path, {method: method, contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, payload: body ? JSON.stringify(body) : undefined}).getContentText() || '{}');
  const files = api('get', '/content').files || [];
  const cur = name => (files.filter(f => f.name === name)[0] || {}).source;
  if (cur('Code') === undefined) return 'No file named "Code" in this project, so nothing was changed.';
  // "Up to date" only once the website (web app deployment) is also running this exact code.
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, code + manifest));
  const props = PropertiesService.getScriptProperties();
  if (cur('Code') === code && cur('appsscript') === manifest && props.getProperty('DEPLOYED_HASH') === hash) return 'Already up to date.';
  const scopes = s => JSON.stringify(((s && JSON.parse(s).oauthScopes) || []).slice().sort());
  if (auto && scopes(cur('appsscript')) !== scopes(manifest)) {
    if (props.getProperty('UPDATE_SCOPES_EMAILED') !== scopes(manifest)) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'Your website script has an update that needs one click',
        'An update to your JC Open House Leads script needs new Google permissions. Open the sheet, then Extensions > Apps Script, ' +
        'pick updateFromGitHub in the menu at the top and click Run, then approve. Everything else updates on its own.');
      props.setProperty('UPDATE_SCOPES_EMAILED', scopes(manifest));
    }
    return 'Update waiting for new permissions; emailed you.';
  }
  // Replace Code and the manifest; keep any other files in the project as they are.
  api('put', '/content', {files: files.filter(f => f.name !== 'Code' && f.name !== 'appsscript').concat(
    [{name: 'appsscript', type: 'JSON', source: manifest}, {name: 'Code', type: 'SERVER_JS', source: code}])});
  const version = api('post', '/versions', {description: 'Update from GitHub'}).versionNumber;
  // Repoint the website's deployment(s). Skip the read-only "Head" (test) deployment, which has no version number.
  (api('get', '/deployments').deployments || [])
    .filter(dep => (dep.entryPoints || []).some(e => e.entryPointType === 'WEB_APP') && (dep.deploymentConfig || {}).versionNumber)
    .forEach(dep => api('put', '/deployments/' + dep.deploymentId,
      {deploymentConfig: {scriptId: id, versionNumber: version, manifestFileName: 'appsscript', description: 'Update from GitHub'}}));
  props.setProperty('DEPLOYED_HASH', hash);
  return 'Updated to version ' + version + ' and pointed the website at it.';
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
// FOLLOW UP BOSS — sends each website lead and open house sign-in to FUB
// as a new lead event, so FUB assigns it and starts your action plans.
// Source and tags show where each lead came from:
//   Open house sign-ins → source "Open House", tag "Open House"
//   Website leads       → source "juliancrider.com", tags "Website Lead" + lead type
// The API key lives in Project Settings > Script Properties as FUB_API_KEY
// (never in this code). If it's missing, this step is skipped.
// ============================================================
function fubEventType(tabName, d) {
  const t = (d.type || '').toLowerCase();
  if (tabName === 'Sign-Ins') return 'Visited Open House';
  if (/sell|home value/.test(t)) return 'Seller Inquiry';
  if (/new construction|showing/.test(t)) return 'Property Inquiry';
  if (/account|guide/.test(t)) return 'Registration';
  return 'General Inquiry';
}

// Sign-ins through the website link (?signin funnel) count as website leads, even when they name an open house.
function fubSource(tabName, d) { return tabName === 'Sign-Ins' && !/Signed up: online|Form: funnel/.test((d && d.notes) || '') ? 'Open House' : 'juliancrider.com'; }

/** A "Call" task due today (Pacific) in Follow Up Boss, assigned to the account owner. */
function createCallTask(personId, reason, type) {
  type = type || 'Call';  // 'Email' for contacts whose phone isn't DNC-checked
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
  return fubFetch('tasks', {personId: personId, type: type, name: (type + ': ' + reason).slice(0, 250), dueDate: today});
}

function fubPerson(tabName, d) {
  const parts = String(d.name || '').trim().split(/\s+/);
  const tags = tabName === 'Sign-Ins' ? (/Signed up: online/.test(d.notes || '') ? ['Website Lead', 'Online Sign-In']
    : /Form: funnel/.test(d.notes || '') ? ['Website Lead', 'Open House'] : ['Open House']) : ['Website Lead', 'juliancrider.com'];
  if (tabName === 'Sign-Ins' && /juliancrider\.com/.test(d.notes || '')) tags.push('juliancrider.com');
  // "Do you live in this neighborhood? Yes" on the open-house form: neighbor welcome + seller-style follow-up.
  if (tabName === 'Sign-Ins' && /Neighbor: Yes/.test(d.notes || '')) tags.push('Open House Neighbor', 'Likely Seller');
  if (tabName === 'Sign-Ins' && /Text OK: Yes/.test(d.notes || '')) tags.push('Text OK');  // checked the text-consent box
  // Other agents previewing the home: tagged Partner (kept out of buyer emails, scoring and the funnel).
  const isAgent = tabName === 'Sign-Ins' && /Visitor: I.m a real estate agent/.test(d.notes || '');
  if (isAgent) tags.push('Agent', 'Partner');
  if (d.type) tags.push(d.type.replace(/^Consultation:\s*/, ''));
  if (tabName !== 'Sign-Ins') tags.push('Text OK');  // website forms carry the "agree to be contacted by text" line
  scoreLead(tabName, d).forEach(t => { if (tags.indexOf(t) < 0 && !(isAgent && /^(Hot|Warm|Likely Seller)$/.test(t))) tags.push(t); });
  const person = {firstName: parts.shift() || '', lastName: parts.join(' '), tags: tags};
  if (validEmail(d.email)) person.emails = [{value: String(d.email).trim()}];
  if (d.phone) person.phones = [{value: d.phone}];
  return person;
}

// Scores a new lead from what the form collected and returns tier tags
// ("Intake Tier A/B/C" = how strong the sign-up looked, plus today's "Tier A/B/C" and "Hot"/"Warm" so a strong new
// lead starts its action plan right away; the nightly scoreFubDatabase keeps Tier and Hot/Warm current after that).
function scoreLead(tabName, d) {
  const notes = String(d.notes || ''), type = String(d.type || '').toLowerCase(), tl = String(d.timeline || '');
  if (/Working with agent: Yes, with a signed/.test(notes)) return ['Intake Tier C', 'Tier C', 'Has Agent'];
  const tags = [];
  let pts = ({'ASAP / Already looking': 30, 'Already looking': 30, 'Ready now': 30, '1–3 months': 25, '3–6 months': 15, '6–12 months': 8}[tl] || 0);
  const interest = Number(d.score) || 0;
  pts += interest >= 8 ? 20 : interest >= 6 ? 10 : interest ? 3 : 0;
  if (/Pre-approved: Yes|Pre-approved: Paying cash|Paying cash/.test(notes)) pts += 10;
  if (/showing/.test(type)) { tags.push('Showing requested'); pts += 45; }
  else if (/sell|home value/.test(type)) { tags.push('Likely Seller'); pts += 20; }
  else if (/new construction|buying|relocat/.test(type)) pts += 10;
  if (/^Owning/.test(String(d.status || '')) && pts >= 15) tags.push('Likely Seller');
  // Julián's focus: out-of-state buyers at $1.3M+ moving within ~60 days (booking form: Price range, Lives in)
  const lux = /Price range: \$(1\.3M–\$3\.3M|3\.3M\+)/.test(notes);
  const away = /Lives in: (California|Another state|Outside)/.test(notes);
  if (lux) { tags.push('Luxury Buyer'); pts += 15; }
  if (away) tags.push('Relocation');
  if (lux && away && /^(Ready now|1–3 months)$/.test(tl)) { tags.push('Luxury Relocation 60 Days'); pts = Math.max(pts, 45); }
  const t = pts >= 45 ? 'A' : pts >= 25 ? 'B' : 'C';
  tags.unshift('Intake Tier ' + t, 'Tier ' + t);
  if (t !== 'C') tags.push(t === 'A' ? 'Hot' : 'Warm');
  return tags.filter((t, i) => tags.indexOf(t) === i);
}

// FUB rejects the whole lead over a mistyped email, so send only well-formed ones and keep the rest in the note.
function validEmail(e) { return /^[^@\s,;]+@[^@\s,;]+\.[a-z]{2,}$/i.test(String(e || '').trim()); }
function badEmailNote(e) { return e && !validEmail(e) ? 'Email as entered (invalid): ' + e : ''; }

function fubFetch(path, body, method) {
  const key = PropertiesService.getScriptProperties().getProperty('FUB_API_KEY');
  if (!key) throw new Error('FUB_API_KEY is not set in Script Properties');
  const opts = {method: method || 'post', muteHttpExceptions: true, headers: {Authorization: 'Basic ' + Utilities.base64Encode(key + ':')}};
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  const res = UrlFetchApp.fetch('https://api.followupboss.com/v1/' + path, opts);
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('FUB ' + code + ': ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText() || '{}');
}

function sendToFub(tabName, d) {
  if (!PropertiesService.getScriptProperties().getProperty('FUB_API_KEY')) return;
  const person = fubPerson(tabName, d);
  const event = {
    source: fubSource(tabName, d),
    system: 'juliancrider.com',
    type: fubEventType(tabName, d),
    message: [d.type, d.timeline && 'Timeline: ' + d.timeline, d.notes, badEmailNote(d.email)].filter(Boolean).join('\n'),
    person: person
  };
  if (d.property) event.property = {street: d.property};
  const res = fubFetch('events', event);
  console.log('Sent to Follow Up Boss');
  const personId = () => {
    let id = res.personId || (res.person && res.person.id);
    if (!id) {
      const email = (person.emails || [])[0], phone = String(d.phone || '').replace(/\D/g, '');
      const found = (email && fubFetch('people?limit=1&email=' + encodeURIComponent(email.value), null, 'get').people) ||
                    (phone && fubFetch('people?limit=1&phone=' + phone, null, 'get').people) || [];
      id = found.length && found[0].id;
    }
    if (!id && a.who) {  // name only: the event created them, so pick the newest person with that exact name
      const same = (fubFetch('people?limit=5&sort=-created&fields=id,name&q=' + encodeURIComponent(a.who), null, 'get').people || [])
        .filter(x => (x.name || '').toLowerCase() === String(a.who).toLowerCase());
      id = same.length && same[0].id;
    }
    return id;
  };
  // Welcome email (FUB Action Plan "JC Welcome", template 183) right after an open-house sign-in or a website
  // account/guide sign-up. Plans started through the API begin Paused, so switch it to Running.
  if ((tabName === 'Sign-Ins' || /account|guide/i.test(d.type || '')) && (person.emails || []).length && person.tags.indexOf('Agent') < 0) {
    try {
      const id = personId();
      if (id) {
        const plan = person.tags.indexOf('Open House Neighbor') >= 0 ? NEIGHBOR_WELCOME_PLAN : WELCOME_PLAN;
        const ap = fubFetch('actionPlansPeople', {personId: id, actionPlanId: plan});
        if (ap.status !== 'Running') fubFetch('actionPlansPeople/' + ap.id, {status: 'Running'}, 'put');
      }
    } catch (err) { console.error('Welcome email: ' + err); }
  }
  // Open-house follow-up texts (2 hours and 72 hours later), only for guests who checked the text-consent box.
  if (tabName === 'Sign-Ins' && person.tags.indexOf('Text OK') >= 0 && person.tags.indexOf('Has Agent') < 0 && d.phone) {
    const kind = person.tags.indexOf('Agent') >= 0 ? 'agent' : person.tags.indexOf('Open House Neighbor') >= 0 ? 'neighbor'
      : /Signed up: online/.test(d.notes || '') ? 'online' : 'guest';
    try { const id = personId(); if (id) queueOpenHouseTexts(id, d, kind); }
    catch (err) { console.error('Open house texts: ' + err); }
  }
  // A task only when they reached out from the website (showing request, consultation). Open house sign-ins get
  // the automated plans instead; Julián engages once they reply.
  if (person.tags.indexOf('Hot') >= 0 && tabName !== 'Sign-Ins') {
    try {
      const id = personId();
      if (id) createCallTask(id, `New ${d.type || 'lead'}: call within 5 minutes`);
    } catch (err) { console.error('Call task: ' + err); }
  }
}
const NEIGHBOR_WELCOME_PLAN = 32;  // "JC Neighbor Welcome (open house)": template 186, home value + off-market P.S.
const WELCOME_PLAN = 29;  // FUB Action Plan "JC Welcome (sign-in)": one email, template 183 "Thanks for stopping by"

/**
 * One-time import of the leads already in this sheet (Sign-Ins and Website Leads) into Follow Up Boss.
 * Uses FUB's people import (not lead events), so it does NOT start your new-lead automations,
 * and FUB matches existing people by email/phone instead of duplicating them.
 * Each person gets a note with everything from their row. Rows named "Test…" are skipped.
 * Safe to re-run: it remembers where it stopped and continues (Apps Script stops runs after 6 minutes).
 */
function importSheetToFub() {
  const ss = getSS(), props = PropertiesService.getScriptProperties(), started = Date.now();
  let sent = 0, skipped = 0, failed = 0;
  for (const tabName of ['Sign-Ins', 'Website Leads']) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;
    const values = sheet.getDataRange().getDisplayValues();
    const head = values[0].map(h => String(h).trim());
    const doneKey = 'FUB_IMPORT_' + tabName;
    for (let r = Number(props.getProperty(doneKey) || 1); r < values.length; r++) {
      if (Date.now() - started > 5 * 60 * 1000) {
        console.log(`Stopped at ${tabName} row ${r + 1} to stay under Google's time limit. Run importSheetToFub again to continue.`);
        console.log(`So far: ${sent} imported, ${skipped} skipped, ${failed} failed.`);
        return;
      }
      const row = values[r], get = name => { const i = head.indexOf(name); return i < 0 ? '' : String(row[i]).trim(); };
      const extra = row.slice(head.length).concat(head.map((h, i) => h ? '' : row[i])).filter(Boolean).join(' | ');
      const d = tabName === 'Sign-Ins'
        ? {name: get('Full Name'), phone: get('Phone'), email: get('Email'), property: get('Open House Address'), score: get('Interest (1–10)'),
           status: get('Currently Renting or Owning?'), timeline: get('Timeline to Buy'), source: get('How They Heard'), notes: [get('Agent Notes'), extra].filter(Boolean).join(' | ')}
        : {type: get('Lead Type'), name: get('Full Name'), phone: get('Phone'), email: get('Email'), property: get('Property'),
           timeline: get('Timeline'), source: get('Source'), notes: get('Notes')};
      if ((!d.name && !d.phone && !d.email) || /^test\b/i.test(d.name)) { skipped++; props.setProperty(doneKey, String(r + 1)); continue; }
      const emailNote = badEmailNote(d.email);
      try {
        const person = fubPerson(tabName, d);
        person.source = fubSource(tabName);
        const created = fubFetch('people?deduplicate=true', person);
        const lines = tabName === 'Sign-Ins'
          ? [['Signed in', get('Timestamp')], ['Open house', d.property], ['Interest (1–10)', d.score], ['Renting or owning', d.status],
             ['Timeline', d.timeline], ['How they heard', d.source], ['Notes', d.notes]]
          : [['Submitted', get('Timestamp')], ['Lead type', d.type], ['Property', d.property], ['Timeline', d.timeline], ['Notes', d.notes]];
        const body = lines.filter(l => l[1]).map(l => l[0] + ': ' + l[1]).concat(emailNote ? [emailNote] : []).join('\n');
        if (created.id && body) fubFetch('notes', {personId: created.id, subject: 'Imported from ' + tabName + ' sheet', body: body});
        sent++;
      } catch (err) {
        failed++;
        console.error(`${tabName} row ${r + 1} (${d.name}): ${err}`);
      }
      props.setProperty(doneKey, String(r + 1));
      Utilities.sleep(200);
    }
  }
  console.log(`Done: ${sent} imported, ${skipped} skipped, ${failed} failed.`);
}

/**
 * Adds the tags from the "Lead Scoring" sheet (Tier A/B/C, Partner, Likely Seller, Referrer,
 * SIG – verify, Has Agent) to the matching people in Follow Up Boss. Finds each person by email,
 * then phone, and keeps every tag they already have. Remembers its place, so it's safe to re-run.
 */
const SCORING_SHEET_ID = '1ohzgr9LgEfQZCF6Dvb_mrlcvzHL7lgqDsxIG-rQfp54';
function tagScoredLeadsInFub() {
  const props = PropertiesService.getScriptProperties(), started = Date.now();
  const values = SpreadsheetApp.openById(SCORING_SHEET_ID).getSheets()[0].getDataRange().getDisplayValues();
  const head = values[0];
  let tagged = 0, notFound = 0, failed = 0;
  for (let r = Number(props.getProperty('FUB_TAGGED_ROW_V2') || 1); r < values.length; r++) {
    if (Date.now() - started > 5 * 60 * 1000) { console.log(`Paused at row ${r + 1}. Run tagScoredLeadsInFub again to continue.`); break; }
    const get = name => String(values[r][head.indexOf(name)] || '').trim();
    const tier = get('Tier'), flags = get('Flags');
    const want = [tier === 'Partner' ? 'Partner' : 'Intake Tier ' + tier];
    if (get('Seller')) want.push('Likely Seller');
    if (get('Referrer')) want.push('Referrer');
    if (/SIG/.test(flags)) want.push('SIG – verify');
    if (/signed agent/i.test(flags)) want.push('Has Agent');
    try {
      const email = get('Email'), phone = get('Phone').replace(/\D/g, '');
      let found = validEmail(email) ? fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people : [];
      if (!(found && found.length) && phone) found = fubFetch('people?limit=1&phone=' + phone, null, 'get').people;
      if (!(found && found.length)) { notFound++; console.warn(`Not in FUB: ${get('Name')}`); }
      else {
        const person = found[0];
        const tags = (person.tags || []).filter(t => !/^Intake Tier [ABC]$/.test(t) && !(t === 'Hot' && tier !== 'A') && !(t === 'Warm' && tier !== 'B'));
        want.forEach(t => { if (tags.indexOf(t) < 0) tags.push(t); });
        fubFetch('people/' + person.id, {tags: tags}, 'put');
        tagged++;
      }
    } catch (err) { failed++; console.error(`${get('Name')}: ${err}`); }
    props.setProperty('FUB_TAGGED_ROW_V2', String(r + 1));
    Utilities.sleep(150);
  }
  console.log(`Tagged ${tagged}, not found ${notFound}, failed ${failed}.`);
}

// ============================================================
// FUB INBOX — drop a CSV or Google Sheet into the Drive folder "FUB Inbox" and it's
// imported into Follow Up Boss automatically (checked every 10 minutes), then moved
// to "FUB Imported" and you get an email summary.
// Columns it understands (any order, any capitalization): First Name, Last Name, Email,
// Phone, Phone 2, Address, City, State, Zip, Source, Tags, Notes. Without a Source or
// Tags column, the file name is used for both.
// Uses FUB's people import (matches existing people, starts no automations).
// ============================================================
const FUB_INBOX = 'FUB Inbox', FUB_DONE = 'FUB Imported';

function driveFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** Run once: creates the two Drive folders and the 10-minute check. */
function setupFubInbox() {
  const inbox = driveFolder(FUB_INBOX); driveFolder(FUB_DONE);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processFubInbox').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processFubInbox').timeBased().everyMinutes(10).create();
  console.log('Ready. Drop files into: ' + inbox.getUrl());
}

function inboxRows(file) {
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.open(file).getSheets()[0].getDataRange().getDisplayValues();
  return Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8').replace(/^﻿/, ''));
}

function processFubInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;  // a previous run is still going
  try {
    const props = PropertiesService.getScriptProperties(), started = Date.now();
    const done = driveFolder(FUB_DONE), files = driveFolder(FUB_INBOX).getFiles();
    while (files.hasNext()) {
      const file = files.next(), id = file.getId(), label = file.getName().replace(/\.(csv|xlsx?)$/i, '').replace(/^FUB Import - /i, '');
      const rows = inboxRows(file).filter(r => r.some(c => String(c).trim()));
      if (rows.length < 2) { file.moveTo(done); continue; }
      const head = rows[0].map(h => String(h).trim().toLowerCase());
      const col = n => head.indexOf(n);
      const stats = JSON.parse(props.getProperty('INBOX_' + id) || '{"row":1,"ok":0,"failed":0}');
      for (; stats.row < rows.length; stats.row++) {
        if (Date.now() - started > 5 * 60 * 1000) { props.setProperty('INBOX_' + id, JSON.stringify(stats)); return; }
        const r = rows[stats.row], get = n => col(n) < 0 ? '' : String(r[col(n)] || '').trim();
        const person = {firstName: get('first name'), lastName: get('last name'), source: get('source') || label,
          tags: (get('tags') || label).split(',').map(t => t.trim()).filter(Boolean)};
        if (!person.firstName && !person.lastName && !get('email') && !get('phone')) continue;
        if (validEmail(get('email'))) person.emails = [{value: get('email')}];
        const phones = [get('phone'), get('phone 2')].filter(p => p.replace(/\D/g, '').length >= 10);
        if (phones.length) person.phones = phones.map(p => ({value: p}));
        if (get('address')) person.addresses = [{type: 'home', street: get('address'), city: get('city'), state: get('state'), code: get('zip')}];
        try {
          const created = fubFetch('people?deduplicate=true', person);
          if (created.id && get('notes')) fubFetch('notes', {personId: created.id, subject: 'Imported: ' + label, body: get('notes')});
          stats.ok++;
        } catch (err) { stats.failed++; console.error(`${file.getName()} row ${stats.row + 1}: ${err}`); }
        Utilities.sleep(100);
      }
      props.deleteProperty('INBOX_' + id);
      file.moveTo(done);
      MailApp.sendEmail(NOTIFY_EMAIL, `Follow Up Boss import done: ${label}`,
        `Imported ${stats.ok} people from "${file.getName()}"${stats.failed ? ` (${stats.failed} failed; see Apps Script executions for details)` : ''}.\nThe file was moved to your "${FUB_DONE}" folder.`);
    }
  } finally { lock.releaseLock(); }
}

// ============================================================
// FUB SCORECARD — reads everyone in Follow Up Boss, ranks them, writes the
// "FUB Scorecard" tab and (weekly) emails you who to call.
//   scoreFubDatabase()   run anytime to refresh the scorecard tab
//   setupWeeklyDigest()  run once: emails the summary every Monday ~7am
//   weeklyFubDigest()    what the Monday trigger runs (also safe to run by hand)
// Tiers: Hot · Warm · Nurture · Cold · Past client / sphere · Partner · Do not contact
// ============================================================
function fubAllPeople() {
  const people = [];
  for (let offset = 0; offset < 10000; offset += 100) {
    const res = fubFetch(`people?limit=100&offset=${offset}&sort=created&fields=allFields`, null, 'get');
    const page = res.people || [];
    people.push(...page);
    if (page.length < 100) break;
    Utilities.sleep(150);
  }
  return people;
}

function daysAgo(iso) { const t = iso ? new Date(iso).getTime() : NaN; return isNaN(t) ? null : Math.floor((Date.now() - t) / 864e5); }

const LIVE_SINCE = new Date('2026-10-29T12:00:00Z');  // leads created before this were imported in bulk

// Activity = what the lead did, not what we did: replies, calls in, Instagram/Facebook messages, a new inquiry,
// site/IDX visits, email opens/clicks, RealScout views. Only lead-side fields are read, so importing, tagging or
// texting someone never raises their score. Recent activity counts most and fades out after 90 days.
function activityScore(p) {
  const why = []; let pts = 0;
  const fresh = (iso, d7, d30, d90) => { const d = daysAgo(iso); return d === null ? 0 : d <= 7 ? d7 : d <= 30 ? d30 : d <= 90 ? d90 : 0; };
  const replies = [['text', p.lastReceivedText], ['email', p.lastReceivedEmail], ['phone', p.lastIncomingCall], ['DM', p.lastReceivedInboxAppMessage]]
    .filter(r => r[1]).sort((a, b) => new Date(b[1]) - new Date(a[1]));
  if (replies.length) { const n = fresh(replies[0][1], 40, 25, 10); if (n) { pts += n; why.push(`replied by ${replies[0][0]} ${daysAgo(replies[0][1])}d ago`); } }
  // A new inquiry after they were added (showing request, form, IDX inquiry). An import sets this to the import time, so that's skipped.
  if (p.lastLeadActivity && new Date(p.lastLeadActivity) - new Date(p.created) > 36e5) {
    const n = fresh(p.lastLeadActivity, 25, 15, 5); if (n) { pts += n; why.push(`new inquiry ${daysAgo(p.lastLeadActivity)}d ago`); }
  }
  const site = fresh(p.lastIdxVisit, 15, 8, 3); if (site) { pts += site; why.push(`on your site ${daysAgo(p.lastIdxVisit)}d ago`); }
  if (Number(p.propertiesSaved) > 0) { pts += 5; why.push(`saved ${p.propertiesSaved} homes`); }
  const em = fresh(p.lastEmEventActivity, 10, 5, 0); if (em) { pts += em; why.push('opened or clicked an email'); }
  const rsViews = Number(p.customRealScoutPropertyViewsCount) || 0;
  if (rsViews >= 3) { pts += 5; why.push(`viewed ${rsViews} homes in RealScout`); }
  if (Number(p.customRealScoutHomeValueAlertViewsCount) > 0) { pts += 10; why.push('checking their home value'); }
  // Going quiet: 3+ touches from us and never an answer in any channel.
  const n = k => Number(p[k]) || 0;
  const out = n('textsSent') + n('emailsSent') + n('callsOutgoing');
  const heard = n('textsReceived') + n('emailsReceived') + n('callsIncoming') + n('inboxAppMessagesReceived');
  if (out >= 3 && !heard && !pts) { pts -= 10; why.push(`quiet: no reply after ${out} touches`); }
  return {pts, why};
}

function scorePerson(p) {
  const tags = (p.tags || []).map(String), has = t => tags.some(x => x.toLowerCase() === t.toLowerCase());
  const stage = String(p.stage || ''), source = String(p.source || '');
  const sinceCreated = daysAgo(p.created);
  const contacted = p.contacted === true || p.contacted === 1 || Number(p.contacted) > 0;
  // Last activity isn't used: Follow Up Boss updates it whenever anything edits a contact (tags, notes, this script),
  // so it can't tell a lead who's browsing from one we just tagged. "New lead" only counts for leads that arrived
  // live (after LIVE_SINCE); everything created earlier was a bulk import of older sign-ins and lists.
  const bulk = /farm|legacy|import|purchased|list/i.test(source) || has('Farm') || has('Legacy – Seed') || has('Legacy – High');
  const live = !bulk && new Date(p.created) >= LIVE_SINCE;
  const why = []; let score = 0;
  if (/trash|unresponsive|do not/i.test(stage) || has('Has Agent')) return {tier: 'Do not contact', score: 0, why: stage || 'Working with another agent', next: 'Leave out of campaigns'};
  if (/past client|sphere/i.test(stage) || has('Past Client')) return {tier: 'Past client / sphere', score: 0, why: stage || 'Past client', next: 'Referral and anniversary touches'};
  if (has('Partner') || /agent|lender|vendor/i.test(stage)) return {tier: 'Partner', score: 0, why: 'Agent / industry / vendor', next: 'Agent Network, not client campaigns'};
  if (/active client|pending|under contract/i.test(stage)) return {tier: 'Active client', score: 100, why: stage, next: 'Keep the deal moving'};
  if (/hot/i.test(stage)) { score += 40; why.push('stage: ' + stage); }
  if (has('Showing requested')) { score += 40; why.push('requested a showing'); }
  if (has('Agreement signed')) { score += 30; why.push('signed a buyer agreement'); }
  // Intake Tier = how strong their sign-up was. (Plain "Tier A/B/C" is this score's own output, so it isn't an input.)
  if (has('Intake Tier A')) { score += 45; why.push('strong sign-up (intake A)'); } else if (has('Intake Tier B')) { score += 25; why.push('good sign-up (intake B)'); }
  if (has('Legacy – High')) { score += 20; why.push('high-value legacy seed'); } else if (has('Legacy – Seed')) { score += 10; why.push('legacy seed'); }
  if (has('Farm Rank A')) { score += 10; why.push('farm rank A (possible seller)'); }
  if (has('Likely Seller')) { score += 10; why.push('likely seller'); }
  if (live && sinceCreated <= 14) { score += 15; why.push(`new lead (${sinceCreated}d)`); }
  if (/juliancrider|website|zillow|realtor|open house/i.test(source)) { score += 5; }
  if (Number(p.price) >= 1000000) { score += 5; why.push('$' + Math.round(Number(p.price) / 1e5) / 10 + 'M price range'); }
  const act = activityScore(p);
  score = Math.max(0, score + act.pts); why.push(...act.why);
  const tier = score >= 50 ? 'Hot' : score >= 25 ? 'Warm' : score >= 10 ? 'Nurture' : 'Cold';
  // Plans do the reaching out; Julián steps in when the lead does something.
  let next = {Hot: act.pts >= 25 ? 'They reached out: answer today' : 'Hot plan running: answer when they reply',
    Warm: 'Warm plan running: answer when they reply', Nurture: 'Monthly email', Cold: 'Newsletter only'}[tier];
  if (has('Not DNC checked')) { next = 'Mail or email only until the phone is DNC-checked'; why.push('DNC not checked'); }
  else if (live && !contacted && sinceCreated <= 30) { next = 'NOT CONTACTED YET: call today'; why.push('never contacted'); }
  else if (act.pts >= 25 && tier !== 'Hot') next = 'Engaged recently: answer personally this week';
  // Engaged = they did something themselves (replied, called, DM'd, new inquiry, site visits). Only engaged people get tasks.
  return {tier, score, why: why.join('; ') || 'no recent signals', next, engaged: act.pts >= 25};
}

const SCORECARD_TAB = {name: 'FUB Scorecard', color: '#d13438',
  headers: ['Tier', 'Score', 'Name', 'Phone', 'Email', 'Stage', 'Source', 'Why', 'Next step', 'Last activity', 'Created', 'Tags', 'FUB link']};
const TIER_ORDER = ['Active client', 'Hot', 'Warm', 'Nurture', 'Cold', 'Past client / sphere', 'Partner', 'Do not contact'];

/** Continues a scoring run that hit Apps Script's time limit (scheduled automatically). */
function scoreFubContinue() { scoreFubDatabase(); }

function scoreFubDatabase() {
  const started = Date.now();
  try { ensureTextQueueTrigger(); } catch (e) { console.log('Text queue trigger: ' + e.message); }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'scoreFubContinue').forEach(t => ScriptApp.deleteTrigger(t));
  // Shared brokerage account: score and tag only contacts assigned to Julián (the owner of the API key).
  const me = fubFetch('me', null, 'get').id;
  const people = fubAllPeople().filter(p => !p.assignedUserId || p.assignedUserId === me);
  const rows = people.map(p => ({p, s: scorePerson(p)}))
    .sort((a, b) => TIER_ORDER.indexOf(a.s.tier) - TIER_ORDER.indexOf(b.s.tier) || b.s.score - a.s.score);
  const ss = getSS(); let sheet = ss.getSheetByName(SCORECARD_TAB.name);
  if (sheet) sheet.clearContents(); else sheet = ensureTab(ss, SCORECARD_TAB);
  styleHeader(sheet, SCORECARD_TAB);
  const d = iso => iso ? Utilities.formatDate(new Date(iso), 'America/Los_Angeles', 'M/d/yyyy') : '';
  const values = rows.map(({p, s}) => [s.tier, s.score, p.name || [p.firstName, p.lastName].join(' '), ((p.phones || [])[0] || {}).value || '',
    ((p.emails || [])[0] || {}).value || '', p.stage || '', p.source || '', s.why, s.next, d(p.lastActivity), d(p.created), (p.tags || []).join(', '),
    'https://app.followupboss.com/2/people/view/' + p.id].map(cell));
  if (values.length) sheet.getRange(2, 1, values.length, SCORECARD_TAB.headers.length).setValues(values);
  const counts = {}; rows.forEach(r => counts[r.s.tier] = (counts[r.s.tier] || 0) + 1);
  // Keep tags current in FUB. Everyone scored gets exactly one of Tier A (Hot) / Tier B (Warm) / Tier C (Nurture or
  // Cold); "Do not contact" loses its Tier tag. Hot/Warm tags start your action-plan automations, so they're only
  // added, never removed (Hot replaces Warm; the Hot plan hands off to Warm itself), and never for Farm,
  // Legacy or not-DNC-checked contacts. Set the Script Property SCORECARD_TAGS to "no" to turn this off.
  const TIERS = {Hot: 'Tier A', Warm: 'Tier B', Nurture: 'Tier C', Cold: 'Tier C', 'Do not contact': ''};
  let tagged = 0, tasks = 0, unfinished = false;
  if (PropertiesService.getScriptProperties().getProperty('SCORECARD_TAGS') !== 'no') {
    for (const {p, s} of rows) {
      if (!(s.tier in TIERS)) continue;  // past clients, partners and active clients keep their own tags
      const old = (p.tags || []).slice(), hasT = t => old.indexOf(t) >= 0;
      let next = old.filter(t => !/^Tier [ABC]$/.test(t));
      if (TIERS[s.tier]) next.push(TIERS[s.tier]);
      // Farm, legacy and not-DNC-checked contacts get a Tier but never an automatic plan (legacy goes out in batches).
      if (!hasT('Farm') && !hasT('Not DNC checked') && !old.some(t => /^Legacy/.test(t))) {
        if (s.tier === 'Hot' && !hasT('Hot')) { next = next.filter(t => t !== 'Warm').concat('Hot'); s.newlyHot = p.newHot = true; }
        if (s.tier === 'Warm' && !hasT('Warm') && !hasT('Hot')) { next.push('Warm'); p.newWarm = true; }
      } else if (s.tier === 'Hot' && !hasT('Tier A')) s.newlyHot = true;  // a task for you, never a plan
      if (next.length === old.length && next.every(t => hasT(t))) continue;
      if (Date.now() - started > 5 * 60 * 1000) { unfinished = true; break; }
      try {
        fubFetch('people/' + p.id, {tags: next}, 'put'); p.tags = next; tagged++;
        if (s.newlyHot && s.engaged) { createCallTask(p.id, 'they reached out: ' + s.why, hasT('Not DNC checked') ? 'Email' : 'Call'); tasks++; }
        Utilities.sleep(120);
      } catch (err) { console.error(`Tagging ${p.name}: ${err}`); }
    }
  }
  console.log(`Updated tags on ${tagged} people in Follow Up Boss; created ${tasks} call tasks for newly Hot people.`);
  if (unfinished) {  // big first run: pick up where we left off in a minute
    ScriptApp.newTrigger('scoreFubContinue').timeBased().after(60 * 1000).create();
    console.log('Ran out of time; continuing automatically in about a minute.');
  } else {
    try { runFunnel(people); } catch (err) { console.error('Funnel: ' + err); }
  }
  console.log(`Scored ${rows.length} people: ` + TIER_ORDER.filter(t => counts[t]).map(t => `${t} ${counts[t]}`).join(', '));
  return {rows, counts};
}

// ============================================================
// FUNNEL — moves everyone through your Follow Up Boss automations for you, a few dozen people a night (so it reads
// as personal and protects your email reputation). Runs after the nightly scoring; you never start anything by hand.
//   Hot (Tier A) or Warm (Tier B) or Legacy re-intro or Farm – Best  →  then the 3-year Long-Term Nurture.
//   Tier C (not farm/legacy) goes straight to Long-Term Nurture. Hot upgrades anyone in Warm or Nurture.
// How it starts them (Oct 31): the API can't touch Automations 2.0, but it can still build and start Action Plans.
// Claude built JC Hot/Warm/Legacy Re-Intro/Farm Best/Long-Term Nurture as Action Plans (ids in FUNNEL.plans, emails
// from templates 141-180, stop-on-contact OFF). The script starts the plan (the API starts it Paused, so it then sets
// it Running), stops the person's previous JC plan, and adds the matching JC tag so Smart Lists can show the stage.
// People tagged "Text OK" get the "+ Text" tag (e.g. "JC Hot + Text") once that stage is in textLive.
// Replies don't stop anything: the plan keeps going until they book or you pause it.
// Pause: someone in a deal (stage Active Client, Pending, Under Contract or Closed) or tagged "Pause" gets their JC
// plan paused and tag removed. When that changes (deal fell through, you removed "Pause"), they restart in Long-Term
// Nurture. Never started: no good email, past clients/sphere/trash, Has Agent. Only stages in LIVE are started.
// The "Funnel" tab shows who is in which stage.
// ============================================================
const FUNNEL = {
  perNight: 50,
  // One-time catch-up (Nov 1 bug: these people got their plain first email before the designed re-intro).
  // On this date (Tue) they get the re-intro; their paused plan then resumes at email 2 four days later.
  // Emilee (304) replied and 267/422/427 bounced, so they're left out. Counts toward the day's 50.
  rerun: {date: '2026-11-03', ids: [242, 246, 247, 251, 253, 256, 258, 260, 263, 264, 272, 275, 276, 278, 1906,
    244, 245, 250, 252, 254, 257, 259, 262, 268, 269, 270, 281, 286, 288, 293, 301, 306, 310,
    419, 420, 421, 423, 424, 425, 426, 428, 429, 430, 431, 432, 433]},
  tags: {hot: 'JC Hot', warm: 'JC Warm', legacy: 'JC Legacy Re-Intro', farm: 'JC Farm Best', nurture: 'JC Long-Term Nurture'},
  plans: {hot: 22, warm: 23, legacy: 24, farm: 25, nurture: 26},  // Follow Up Boss Action Plan ids
  // Older contacts (in FUB 30+ days, so they may not have heard from Julián in months or years) first get the designed
  // "It's been a while" re-introduction (plan 30, template 184); their personal plan starts reintroDays later.
  // Farm contacts skip it (they've never talked to Julián; Farm Best 1 is their intro).
  reintroPlan: 30, neighborPlan: 31, legacyAfterIntroPlan: 33, reintroDays: 4, reintroAfterDays: 30,
  live: ['hot', 'warm', 'legacy', 'farm', 'nurture'],  // stages that start people (all five built Oct 31)
  textLive: [],  // stages whose "+ Text" automation exists, e.g. ['hot', 'warm', 'nurture']
  handoffDays: {hot: 25, warm: 65, legacy: 25, farm: 60}  // days after starting, move on to Long-Term Nurture
};
const FUNNEL_TAB = {name: 'Funnel', color: '#0B8043', headers: ['Person ID', 'Name', 'Stage', 'Started', 'Trigger tag', 'Why']};

function runFunnel(people) {
  if (!FUNNEL.live.length) { console.log('Funnel: no automations are marked live yet, so nobody was started.'); return 0; }
  const allTags = Object.keys(FUNNEL.tags).reduce((list, k) => list.concat([FUNNEL.tags[k], FUNNEL.tags[k] + ' + Text']), []);
  const sheet = ensureTab(getSS(), FUNNEL_TAB), last = sheet.getLastRow();
  const rows = {};
  (last > 1 ? sheet.getRange(2, 1, last - 1, FUNNEL_TAB.headers.length).getValues() : [])
    .forEach((r, i) => rows[r[0]] = {row: i + 2, stage: r[2], started: r[3]});
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
  let used = props.getProperty('FUNNEL_DAY') === today ? Number(props.getProperty('FUNNEL_COUNT')) || 0 : 0, started = 0;

  const paused = p => /active client|pending|under contract|closed/i.test(String(p.stage || '')) || has(p, 'Pause');
  const ok = p => {
    const tags = p.tags || [], stage = String(p.stage || '');
    if (!(p.emails || []).some(e => validEmail(e.value) && !/invalid|bounce|bad/i.test(e.status || ''))) return false;
    if (/trash|unresponsive|do not|past client|sphere/i.test(stage) || paused(p)) return false;
    // Brand-new people just got the welcome email; start their drip after 2 days so they don't get two at once.
    if (p.created && (Date.now() - new Date(p.created)) / 864e5 < 2) return false;
    return !tags.some(t => /^(Has Agent|Partner|Do Not Contact)$|^Test\b/i.test(t));
  };
  const record = (p, stage, tag, why) => {
    const cur = rows[p.id], vals = [p.id, p.name || [p.firstName, p.lastName].join(' '), stage, today, tag, why];
    if (cur) sheet.getRange(cur.row, 1, 1, vals.length).setValues([vals]); else sheet.appendRow(vals);
    rows[p.id] = {row: cur ? cur.row : sheet.getLastRow(), stage: stage, started: today};
  };
  const planIds = Object.keys(FUNNEL.plans).map(k => FUNNEL.plans[k]).concat(FUNNEL.legacyAfterIntroPlan);
  // Pause any JC plan this person is running (moving stages, or a deal/Pause tag).
  const stopPlans = p => (fubFetch('actionPlansPeople?personId=' + p.id, null, 'get').actionPlansPeople || [])
    .filter(a => planIds.indexOf(a.actionPlanId) >= 0 && a.status === 'Running')
    .forEach(a => fubFetch('actionPlansPeople/' + a.id, {status: 'Paused'}, 'put'));
  const startPlan = (p, planId) => {
    const ap = fubFetch('actionPlansPeople', {personId: p.id, actionPlanId: planId});
    if (ap.status !== 'Running') fubFetch('actionPlansPeople/' + ap.id, {status: 'Running'}, 'put');
  };
  const start = (p, stage, why) => {
    if (used >= FUNNEL.perNight || FUNNEL.live.indexOf(stage) < 0) return;
    // First email is a designed intro, then the plain-text drip a few days later:
    // Farm – Best homeowners get the "Neighbor Note" (31); anyone else in FUB 30+ days gets "It's been a while" (30).
    // Open-house neighbors already got their neighbor welcome at sign-in, so they skip it.
    // "Old" = imported in the Oct 28-29 bulk load (FUB's created date is the import date, not when Julián met them)
    // or in Follow Up Boss 30+ days.
    const old = p.created && (new Date(p.created) < LIVE_SINCE || (Date.now() - new Date(p.created)) / 864e5 >= FUNNEL.reintroAfterDays);
    const intro = stageOf(p) ? 0 : stage === 'farm' ? (has(p, 'Open House Neighbor') ? 0 : FUNNEL.neighborPlan) : old ? FUNNEL.reintroPlan : 0;
    if (intro) {
      try { startPlan(p, intro); }
      catch (e) { console.log('Funnel: could not send re-intro to ' + p.id + ': ' + e.message); return; }
      record(p, 'reintro-' + stage, '', 're-intro first, then ' + stage + ' (' + why + ')');
      used++; started++; Utilities.sleep(120);
      return;
    }
    const tag = FUNNEL.tags[stage] + (has(p, 'Text OK') && FUNNEL.textLive.indexOf(stage) >= 0 ? ' + Text' : '');
    try {
      stopPlans(p);
      // Legacy 1 is a plain "It's been a while" email, so people who just got the designed re-intro skip it.
      // A paused plan for this stage (e.g. the Nov 2 catch-up) resumes where it left off instead of starting over.
      const alts = stage === 'legacy' ? [FUNNEL.plans.legacy, FUNNEL.legacyAfterIntroPlan] : [FUNNEL.plans[stage]];
      const paused = (fubFetch('actionPlansPeople?personId=' + p.id, null, 'get').actionPlansPeople || [])
        .filter(a => alts.indexOf(a.actionPlanId) >= 0 && a.status === 'Paused')[0];
      if (paused) fubFetch('actionPlansPeople/' + paused.id, {status: 'Running'}, 'put');
      else startPlan(p, stage === 'legacy' && /^reintro-/.test(stageOf(p) || '') ? FUNNEL.legacyAfterIntroPlan : FUNNEL.plans[stage]);
    } catch (e) { console.log('Funnel: could not start ' + stage + ' for ' + p.id + ': ' + e.message); return; }
    const next = (p.tags || []).filter(t => allTags.indexOf(t) < 0).concat(tag);
    fubFetch('people/' + p.id, {tags: next}, 'put'); p.tags = next;
    record(p, stage, tag, why);
    used++; started++; Utilities.sleep(120);
  };
  const has = (p, t) => (p.tags || []).indexOf(t) >= 0, stageOf = p => (rows[p.id] || {}).stage;
  // Pause anyone now in a deal or tagged "Pause" (doesn't count toward tonight's limit; nothing is sent).
  people.filter(p => stageOf(p) && stageOf(p) !== 'paused' && paused(p)).forEach(p => {
    const next = (p.tags || []).filter(t => allTags.indexOf(t) < 0);
    try { stopPlans(p); } catch (e) { console.log('Funnel: could not pause plan for ' + p.id + ': ' + e.message); }
    if (next.length !== (p.tags || []).length) { fubFetch('people/' + p.id, {tags: next}, 'put'); p.tags = next; }
    record(p, 'paused', '', 'paused from ' + stageOf(p) + (has(p, 'Pause') ? ' (Pause tag)' : ' (' + p.stage + ')'));
  });
  const eligible = people.filter(ok);
  // One-time catch-up: designed re-intro for people whose plain first email went out before it (see FUNNEL.rerun).
  if (FUNNEL.rerun && FUNNEL.rerun.date === today) {
    eligible.filter(p => FUNNEL.rerun.ids.indexOf(p.id) >= 0 && !/^reintro-|paused/.test(stageOf(p) || '')).forEach(p => {
      if (used >= FUNNEL.perNight) return;
      const stage = stageOf(p) || 'nurture';
      try { stopPlans(p); startPlan(p, FUNNEL.reintroPlan); }
      catch (e) { console.log('Funnel: catch-up re-intro failed for ' + p.id + ': ' + e.message); return; }
      record(p, 'reintro-' + stage, '', 'catch-up re-intro, then resume ' + stage);
      used++; started++; Utilities.sleep(120);
    });
  }
  // People who got the re-introduction: start their real plan once it's had time to land.
  eligible.filter(p => /^reintro-/.test(stageOf(p) || '') && daysAgo(rows[p.id].started) >= FUNNEL.reintroDays)
    .forEach(p => start(p, stageOf(p).slice(8), 'after re-intro'));
  // In priority order, until tonight's limit is used up.
  // Hot/Warm tags stay on for good, so only a tag added tonight moves someone up from Warm or Nurture.
  eligible.filter(p => has(p, 'Hot') && (!stageOf(p) || (p.newHot && /warm|nurture/.test(stageOf(p))))).forEach(p => start(p, 'hot', 'Tier A'));
  eligible.filter(p => has(p, 'Warm') && !has(p, 'Hot') && (!stageOf(p) || (p.newWarm && stageOf(p) === 'nurture'))).forEach(p => start(p, 'warm', 'Tier B'));
  eligible.filter(p => FUNNEL.handoffDays[stageOf(p)] && daysAgo(rows[p.id].started) >= FUNNEL.handoffDays[stageOf(p)])
    .forEach(p => start(p, 'nurture', 'finished ' + stageOf(p)));
  eligible.filter(p => stageOf(p) === 'paused').forEach(p => start(p, 'nurture', 'resumed after pause'));
  eligible.filter(p => !stageOf(p) && (p.tags || []).some(t => /^Legacy/.test(t))).forEach(p => start(p, 'legacy', 'Legacy contact'));
  eligible.filter(p => !stageOf(p) && has(p, 'Farm – Best')).forEach(p => start(p, 'farm', 'Farm – Best'));
  eligible.filter(p => !stageOf(p) && has(p, 'Open House Neighbor')).forEach(p => start(p, 'farm', 'open-house neighbor'));
  eligible.filter(p => !stageOf(p) && has(p, 'Tier C') && !has(p, 'Farm') && !(p.tags || []).some(t => /^Legacy/.test(t)))
    .forEach(p => start(p, 'nurture', 'Tier C'));
  props.setProperty('FUNNEL_DAY', today); props.setProperty('FUNNEL_COUNT', String(used));
  console.log(`Funnel: started ${started} people tonight (${used}/${FUNNEL.perNight} for today).`);
  return started;
}

function weeklyFubDigest() {
  const {rows, counts} = scoreFubDatabase();
  const e = x => String(x || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const line = ({p, s}) => `<tr><td style="padding:6px 10px 6px 0"><a href="https://app.followupboss.com/2/people/view/${p.id}" style="color:#111;font-weight:600">${e(p.name || p.firstName)}</a></td><td style="padding:6px 10px 6px 0;color:#444">${e(((p.phones || [])[0] || {}).value || '')}</td><td style="padding:6px 0;color:#666">${e(s.why)}</td></tr>`;
  const section = (title, list) => list.length ? `<h3 style="margin:18px 0 6px">${title} (${list.length})</h3><table style="border-collapse:collapse;font-size:14px">${list.slice(0, 15).map(line).join('')}</table>` : '';
  const notContacted = rows.filter(r => /NOT CONTACTED/.test(r.s.next));
  const hot = rows.filter(r => r.s.tier === 'Hot' && notContacted.indexOf(r) < 0);
  const quiet = rows.filter(r => r.s.tier === 'Warm' && /quiet/.test(r.s.why));
  const signed = rows.filter(r => /signed a buyer agreement/.test(r.s.why));
  const html = `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#111;max-width:720px">
    <h2 style="margin:0 0 6px">Your week in Follow Up Boss</h2>
    <p style="color:#555;margin:0 0 10px">${TIER_ORDER.filter(t => counts[t]).map(t => `<b>${counts[t]}</b> ${t}`).join(' · ')}</p>
    ${section('Call today: new leads not contacted yet', notContacted)}${section('Hot', hot)}${section('Signed buyer agreements', signed)}${section('Warm but going quiet', quiet)}
    <p style="margin-top:18px"><a href="${getSS().getUrl()}">Open the full FUB Scorecard</a></p></div>`;
  MailApp.sendEmail({to: NOTIFY_EMAIL, subject: `Who to call this week: ${(counts.Hot || 0) + notContacted.length} priority leads`, htmlBody: html, name: 'juliancrider.com'});
}

/** Run once: sends the digest every Monday morning. */
function setupWeeklyDigest() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklyFubDigest').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyFubDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone('America/Los_Angeles').create();
  console.log('Weekly digest scheduled for Mondays around 7am Pacific.');
}

/** Run once: re-scores everyone in Follow Up Boss every morning around 7am Pacific (so the funnel emails land as people wake up) and adds Hot/Warm tags,
 *  so your FUB automations start the right action plan while you sleep. Check the "FUB Scorecard" tab each morning. */
function setupNightlyScoring() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'scoreFubDatabase').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scoreFubDatabase').timeBased().everyDays(1).atHour(7).inTimezone('America/Los_Angeles').create();
  console.log('Daily scoring + emails scheduled for around 7am Pacific.');
}

/** Start the sheet import over from the top (only needed if you want to re-send everything). */
function resetSheetImport() {
  ['Sign-Ins', 'Website Leads'].forEach(t => PropertiesService.getScriptProperties().deleteProperty('FUB_IMPORT_' + t));
}

/** Run once after adding FUB_API_KEY: creates a "FUB Test" lead in Follow Up Boss (delete it after). */
function testFub() {
  sendToFub('Website Leads', {type: 'Consultation: Test', name: 'FUB Test', email: 'fub-test@example.com',
    phone: '7025550100', timeline: 'Just exploring', notes: 'Test from the script editor — safe to delete'});
}

// ============================================================
// EMAIL TEMPLATES — run loadFubTemplates once to add the Hot/Warm/Legacy plan emails to Follow Up Boss.
// Safe to run again: templates that already exist (same name) are skipped.
// ============================================================
const SIG_FULL = 'Julián J. Crider<br>702-498-2571 · juliancrider.com', SIG = 'Julián';
const FUB_TEMPLATES = [
  ['Hot 1 - Quick question', 'Quick question, %contact_first_name%', [
    'Hi %contact_first_name%,',
    'Thanks for reaching out. You probably have a dozen agents and websites in your inbox right now, so I\'ll keep this short.',
    'What would have to be true for the next home to feel like the right one?',
    'Reply with one line. I read every one myself.', SIG_FULL]],
  ['Hot 2 - Homes I\'d look at', 'The homes I\'d look at if I were you', [
    'Hi %contact_first_name%,',
    'I pulled a few homes that fit what you were looking at, including one or two that aren\'t getting much attention online yet.',
    'Would you be opposed to me sending them over? If the timing\'s off, just say "not now". No hard feelings.', SIG]],
  ['Hot 3 - Is it the market', 'Is it the market?', [
    'Hi %contact_first_name%,',
    'A lot of people I talk to right now feel like they might buy at the wrong time: rates, prices, all of it. It\'s a fair concern.',
    'How would you know it\'s the right time for you?',
    'If it helps, I can walk you through what\'s actually happening in your price range. Ten minutes, no pressure.', SIG]],
  ['Hot 4 - Close your file', 'Should I close your file?', [
    'Hi %contact_first_name%,',
    'I haven\'t heard back, which usually means one of three things: you found a home, the timing changed, or my emails are hitting the wrong inbox.',
    'Is it a bad idea for me to keep sending homes your way? Just reply "keep" or "stop". Either is fine.', SIG]],
  ['Warm 1 - No rush', 'No rush, %contact_first_name%', [
    'Hi %contact_first_name%,',
    'It sounds like the timing isn\'t quite right yet, and that\'s completely okay. Most of my clients took months between "just looking" and "let\'s go."',
    'What would you want to know now so you\'re ready when the time comes?', SIG_FULL]],
  ['Warm 2 - Market check', 'What homes are doing in Las Vegas right now', [
    'Hi %contact_first_name%,',
    'The Las Vegas market changes month to month, and what\'s true for one price range often isn\'t true for another.',
    'If you tell me the area and price range you have in mind, I\'ll send you a quick, honest snapshot of what\'s selling, how fast, and for how much.',
    'How would that change your plans, if at all?', SIG]],
  ['Warm 3 - Home value', 'Curious what your home is worth?', [
    'Hi %contact_first_name%,',
    'If you own now, you\'re probably wondering what your place would sell for. Would it be a bad idea for me to put together a real number for you, not a Zestimate?',
    'It takes me a day, and there\'s no obligation.', SIG]],
  ['Warm 4 - What makes the difference', 'What finally made the difference', [
    'Hi %contact_first_name%,',
    'Most of the families I help were "just looking" for months. What finally made the difference was usually one thing: the right home that never hit Zillow, a payment that finally worked, or a change in life that made the timing obvious.',
    'How close are you to your version of that?', SIG]],
  // Legacy re-introduction (old sphere contacts): reconnect, website, home value, socials, referral.
  ['Legacy 1 - It\'s been a while', 'It\'s been a while, %contact_first_name%', [
    'Hi %contact_first_name%,',
    'It\'s been a while, and I\'ll admit this email is overdue.',
    'I\'m still helping people buy and sell here in Las Vegas, now with Boeckle & Co., and I just finished building a new home base: <a href="https://juliancrider.com">juliancrider.com</a>. Homes, neighborhood guides, and a quick home-value check, all in one place.',
    'No ask here. I just didn\'t want to be the person who only reaches out when they need something.',
    'What\'s new with you?', SIG_FULL]],
  ['Legacy 2 - Equity check', 'You might be sitting on more than you think', [
    'Hi %contact_first_name%,',
    'A lot of people I\'ve worked with over the years are sitting on more equity than they realize, and most have never had anyone run a real number for them.',
    'Would it be a bad idea for me to put one together for your place? Not a Zestimate. It takes me a day, and there\'s no obligation.',
    'Or if you\'re curious first, there\'s a quick home-value check on <a href="https://juliancrider.com">juliancrider.com</a>.', SIG]],
  ['Legacy 3 - Where I share the good stuff', 'Where I share the good stuff', [
    'Hi %contact_first_name%,',
    'I\'ve started sharing what I see every day: new homes before they get crowded, honest market updates, and the Las Vegas spots locals actually go to.',
    'It\'s all on <a href="https://juliancrider.com">juliancrider.com</a>, with links to my Instagram and YouTube.',
    'What would you want to see more of?', SIG]],
  ['Legacy 4 - A small favor', 'Would you be against a small favor?', [
    'Hi %contact_first_name%,',
    'Most of my business comes from people like you, so I\'ll just ask.',
    'If someone you know mentions moving to, from, or around Las Vegas, would you be opposed to passing along my name? I\'ll take great care of them, and I\'ll keep you in the loop.',
    'Either way, it\'s good to be back in touch.', SIG_FULL]]
];

function loadFubTemplates() {
  const existing = (fubFetch('templates?limit=100', null, 'get').templates || []).map(t => t.name);
  const made = [], skipped = [];
  FUB_TEMPLATES.forEach(([name, subject, paras]) => {
    if (existing.indexOf(name) >= 0) return skipped.push(name);
    const t = fubFetch('templates', {name, subject, body: paras.map(x => `<p>${x}</p>`).join('')});
    made.push(`${name} (id ${t.id})`);
  });
  console.log(`Created ${made.length}: ${made.join(', ') || 'none'}\nAlready there: ${skipped.join(', ') || 'none'}`);
}

// ============================================================
// BUYER CONFIRMATION — branded email to the buyer after a showing request
// ============================================================
function confirmShowing(d, signUrl) {
  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const first = esc(String(d.name || '').trim().split(/\s+/)[0] || 'there');
  const when = [d.timeline, (String(d.notes || '').match(/Preferred time: ([^|]+)/) || [])[1]].filter(Boolean).map(s => esc(s.trim())).join(', ');
  const hasAgent = /Working with agent: Yes, with a signed/.test(d.notes || '');
  const body = hasAgent
    ? `<p>Thanks for your interest in ${d.property ? `<b>${esc(d.property)}</b>` : 'seeing homes'}. Since you have a signed agreement with another agent, they should schedule this tour for you, and they'll be able to get you in quickly.</p>
       <p>If anything changes, or you have questions about the home or the neighborhood, I'm always happy to help.</p>`
    : !d.property ? `<p>Thanks for reaching out about seeing homes${when ? ` (${when})` : ''}. I'll be in touch shortly to hear what you're looking for, then line up the homes worth seeing.</p>
       <p>Before any tour, Nevada requires a short written buyer agreement. Once we pick a home, I'll send a non-exclusive one that covers only that home. It takes about a minute to sign.</p>`
    : `<p>Thanks for requesting a showing of <b>${esc(d.property)}</b>${when ? ` (${when})` : ''}. I'll confirm a time with you shortly.</p>
       <p style="margin:18px 0 6px"><b>What happens next</b></p>
       <ol style="margin:0 0 14px 18px;padding:0">
         <li style="margin-bottom:6px">I confirm the day and time with you.</li>
         <li style="margin-bottom:6px">${signUrl ? 'Sign' : 'You get'} a short buyer agreement${signUrl ? '' : ' to e-sign'}. It's <b>non-exclusive and covers only this one home</b>, so you're free to work with anyone else. Nevada requires a written agreement before any tour, and it takes about a minute.${signUrl ? `<br><a href="${signUrl}" style="display:inline-block;margin-top:10px;background:#000;color:#fff;text-decoration:none;font-weight:700;padding:11px 18px;border-radius:999px">Review &amp; sign</a>` : ''}</li>
         <li>We meet at the home. If you're financing, bring your pre-approval letter if you have one.</li>
       </ol>`;
  const html = `<div style="background:#f4f4f5;padding:24px 12px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
    <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden">
      <div style="background:#000;color:#fff;padding:20px 24px;font-size:18px;font-weight:700">Julián J. Crider <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#d13438;margin-left:4px"></span></div>
      <div style="padding:24px;color:#1c1c1e;font-size:15px;line-height:1.55">
        <p style="font-size:20px;font-weight:700;margin:0 0 12px">Hi ${first}, your showing request is in.</p>
        ${body}
        <p style="margin:18px 0 0">Talk soon,<br><b>Julián J. Crider</b><br><span style="color:#6e6e73">Las Vegas Realtor® · <a href="tel:+17024982571" style="color:#6e6e73">702-498-2571</a> · <a href="https://juliancrider.com" style="color:#6e6e73">juliancrider.com</a></span></p>
      </div>
    </div>
    <p style="max-width:560px;margin:12px auto 0;text-align:center;color:#8e8e93;font-size:11px">You're receiving this because you requested a showing on juliancrider.com. Reply to this email with any questions.</p>
  </div>`;
  MailApp.sendEmail({
    to: String(d.email).trim(), replyTo: NOTIFY_EMAIL, name: 'Julián J. Crider',
    subject: hasAgent ? 'About your showing request' : 'Your showing request: ' + String(d.property || '').slice(0, 60),
    htmlBody: html, body: html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
  });
}

// ============================================================
// BUYER AGREEMENT — single-property, non-exclusive, signed on juliancrider.com/sign.html
// Broker-approved (3% fee). Live by default; set the Script Property AGREEMENT_LIVE to "no" to pause.
// The buyer also acknowledges Nevada's Duties Owed by a Nevada Real Estate Licensee (form 525)
// with the same signature. previewAgreement() creates a test link.
// Optional Script Properties (Project Settings > Script Properties) override the defaults:
//   AGREEMENT_LIVE "no" to pause · BUYER_FEE · AGREEMENT_TERM_DAYS (30) · PROTECTION_DAYS (30)
//   LICENSEE_NAME · LICENSE_NUMBER · BROKERAGE · BROKER_NAME (Landon Boeckle) · BROKER_LICENSE (B.1001882)
// ============================================================
const SIGN_PAGE = 'https://juliancrider.com/sign.html';
const AGREEMENT_TAB = {name: 'Buyer Agreements', color: '#000000',
  headers: ['Created', 'Agreement ID', 'Buyer', 'Email', 'Phone', 'Property', 'Fee', 'Term (days)', 'Status', 'Signed at', 'Signed name', 'Signed PDF']};

function agreementSettings() {
  const p = PropertiesService.getScriptProperties(), g = (k, d) => p.getProperty(k) || d;
  return {live: g('AGREEMENT_LIVE', 'yes') !== 'no', fee: g('BUYER_FEE', '3% of the purchase price'), termDays: Number(g('AGREEMENT_TERM_DAYS', 30)),
    protectDays: Number(g('PROTECTION_DAYS', 30)), licensee: g('LICENSEE_NAME', 'Julián J. Crider'), licenseNo: g('LICENSE_NUMBER', 'S.0177760'),
    brokerage: g('BROKERAGE', 'Boeckle & Co.'), brokerName: g('BROKER_NAME', 'Landon Boeckle'), brokerLicense: g('BROKER_LICENSE', 'B.1001882')};
}

function createAgreement(d) {
  const s = agreementSettings(), id = Utilities.getUuid();
  ensureTab(getSS(), AGREEMENT_TAB).appendRow([new Date().toLocaleString('en-US', {timeZone: 'America/Los_Angeles'}), id, d.name, d.email, d.phone, d.property,
    s.fee, s.termDays, 'Sent', '', '', ''].map(cell));
  return SIGN_PAGE + '?id=' + id;
}

function findAgreement(id) {
  if (!/^[0-9a-f-]{36}$/i.test(id)) return null;
  const sheet = getSS().getSheetByName(AGREEMENT_TAB.name);
  if (!sheet) return null;
  const v = sheet.getDataRange().getDisplayValues();
  for (let r = 1; r < v.length; r++) if (v[r][1] === id) {
    const [created, , buyer, email, phone, property, fee, termDays, status, signedAt, signedName, pdf] = v[r];
    return {sheet, row: r + 1, data: {created, id, buyer, email, phone, property, fee, termDays: Number(termDays) || 30, status, signedAt, signedName, pdf}};
  }
  return null;
}

function getAgreement(id) {
  const a = findAgreement(id);
  if (!a) return json({ok: false, error: 'This signing link is not valid. Please contact Julián.'});
  const s = agreementSettings();
  return json({ok: true, live: s.live, status: a.data.status, buyer: a.data.buyer, property: a.data.property, html: agreementHtml(a.data, s), duties: dutiesHtml(a.data, s, '')});
}

// The agreement wording (shared by the signing page and the signed PDF).
function agreementHtml(a, s) {
  const e = x => String(x || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<h1>Single-Property Non-Exclusive Buyer Representation Agreement</h1>
<p><b>Buyer:</b> ${e(a.buyer)}<br><b>Broker:</b> ${e(s.brokerage)} (${e(s.brokerName)}, broker, license ${e(s.brokerLicense)}), through licensee ${e(s.licensee)}${s.licenseNo ? ` (Nevada license #${e(s.licenseNo)})` : ''}<br><b>Property:</b> ${e(a.property)}</p>
<h2>1. What this covers</h2><p>Buyer engages Broker to show and provide services for <b>only the Property listed above</b>. This agreement is <b>non-exclusive</b>: Buyer is free to work with any other agent on any other property, and nothing here limits that.</p>
<h2>2. How long it lasts</h2><p>This agreement starts when Buyer signs it and ends ${a.termDays} days later at 11:59 p.m., or at closing if Buyer is under contract to buy the Property, whichever is later. Either party may cancel at any time with written notice (email is fine), unless Buyer is already under contract to buy the Property.</p>
<h2>3. Broker's services</h2><p>Broker will make good-faith efforts to show the Property, share information about it, and, if Buyer chooses, help Buyer prepare, negotiate and complete an offer and purchase.</p>
<h2>4. Broker's compensation</h2><p>If Buyer buys the Property during this agreement, or within ${s.protectDays} days after it ends, Buyer agrees Broker is owed <b>${e(s.fee)}</b>, due at closing. <b>This fee is not set by law and is fully negotiable.</b> Buyer may ask that the seller pay all or part of it as part of the offer, and Broker will help negotiate that. Broker will not accept more than this amount from all sources combined.</p>
<h2>5. Other buyers and dual agency</h2><p>Broker may work with other buyers interested in the same or similar properties. If Broker would also represent the seller of the Property, Broker must first obtain Buyer's written consent on Nevada's Consent to Act form.</p>
<h2>6. Nevada disclosures</h2><p>Buyer acknowledges receiving the Nevada <i>Duties Owed by a Nevada Real Estate Licensee</i> disclosure (form 525), included with this agreement.</p>
<h2>7. Wire fraud warning</h2><p>Criminals target real estate payments. Broker will never send wiring instructions by email. Always confirm wiring instructions by phone with the title or escrow company, using a number you find independently, before sending money.</p>
<h2>8. Fair housing</h2><p>Broker follows all federal, state and local fair housing laws and does not discriminate against anyone.</p>
<h2>9. Nevada law</h2><p>Nevada law governs this agreement.</p>
<h2>10. Electronic signature</h2><p>Buyer agrees to sign electronically and agrees that an electronic signature has the same effect as a handwritten one. Buyer will receive a signed copy by email.</p>`;
}

// Nevada Real Estate Division form 525, filled in for the buyer side.
function dutiesHtml(a, s, dual, initials) {
  const e = x => String(x || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const box = on => on ? '&#9746;' : '&#9744;';
  const ini = initials ? e(initials) : '______';
  return `<h1>Duties Owed by a Nevada Real Estate Licensee</h1>
<p><i>This form does not constitute a contract for services nor an agreement to pay compensation.</i></p>
<p>In Nevada, a real estate licensee is required to provide a form setting forth the duties owed by the licensee to: a) each party for whom the licensee is acting as an agent in the real estate transaction, and b) each unrepresented party to the real estate transaction, if any.</p>
<p><b>Licensee:</b> The licensee in the real estate transaction is <b>${e(s.licensee)}</b> whose license/permit number is <b>${e(s.licenseNo)}</b>.<br>
The licensee is acting for <b>${e(a.buyer)}</b>, who is/are the ${box(false)} Seller/Landlord ${box(true)} Buyer/Tenant.<br>
<b>Broker:</b> The Broker is <b>${s.brokerName ? e(s.brokerName) : 'the designated broker of ' + e(s.brokerage)}</b>, whose company is <b>${e(s.brokerage)}</b>.<br>
Are there additional licensees involved in this transaction? ${box(false)} Yes ${box(true)} No</p>
<h2>Licensee's duties owed to all parties</h2><p>A Nevada real estate licensee shall: 1. Not deal with any party to a real estate transaction in a manner which is deceitful, fraudulent or dishonest. 2. Exercise reasonable skill and care with respect to all parties to the real estate transaction. 3. Disclose to each party to the real estate transaction as soon as practicable: a. any material and relevant facts, data or information which licensee knows, or with reasonable care and diligence the licensee should know, about the property; b. each source from which licensee will receive compensation. 4. Abide by all other duties, responsibilities and obligations required of the licensee in law or regulations.</p>
<h2>Licensee's duties owed to the client</h2><p>A Nevada real estate licensee shall: 1. Exercise reasonable skill and care to carry out the terms of the brokerage agreement and the licensee's duties in the brokerage agreement or property management agreement; 2. Not disclose, except to the licensee's broker, confidential information relating to a client for 1 year after the revocation or termination of the brokerage agreement, unless licensee is required to do so by court order or the client gives written permission; 3. Seek a sale, purchase, option, rental or lease of real property at the price and terms stated in the brokerage agreement or at a price acceptable to the client; 4. Present all offers made to, or by the client as soon as practicable, unless the client chooses to waive the duty of the licensee to present all offers and signs a waiver of the duty on a form prescribed by the Division; 5. Disclose to the client material facts of which the licensee has knowledge concerning the real estate transaction; 6. Advise the client to obtain advice from an expert relating to matters which are beyond the expertise of the licensee; and 7. Account to the client for all money and property the licensee receives in which the client may have an interest.</p>
<h2>Duties owed by a broker who assigns different licensees affiliated with the brokerage to separate parties</h2><p>Each licensee shall not disclose, except to the real estate broker, confidential information relating to client.</p>
<h2>Licensee acting for both parties</h2><p>The Licensee <b>MAY [${dual === 'MAY' ? ini : '______'}]</b> OR <b>MAY NOT [${dual === 'MAY NOT' ? ini : '______'}]</b> in the future act for two or more parties who have interests adverse to each other. In acting for these parties, the licensee has a conflict of interest. Before a licensee may act for two or more parties, the licensee must give you a "Consent to Act" form to sign.</p>
<p><b>I/We acknowledge receipt of a copy of this list of licensee duties, and have read and understand this disclosure.</b></p>
<p style="font-size:12px;color:#666">Nevada Real Estate Division form 525, revised 10/03/2023.</p>`;
}

function signAgreement(d) {
  const a = findAgreement(String(d.id || ''));
  if (!a) return json({ok: false, error: 'This signing link is not valid. Please contact Julián.'});
  if (a.data.status === 'Signed') return json({ok: true, already: true});
  const name = String(d.typedName || '').trim(), sig = String(d.signature || '');
  if (name.length < 2 || d.consent !== true) return json({ok: false, error: 'Please type your full name and check all the boxes.'});
  if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(sig) || sig.length > 400000) return json({ok: false, error: 'Please draw your signature in the box.'});
  const dual = String(d.dualAgency || '');
  if (d.dutiesAck !== true || ['MAY', 'MAY NOT'].indexOf(dual) < 0) return json({ok: false, error: 'Please choose MAY or MAY NOT and acknowledge the Duties Owed form.'});
  const initials = name.split(/\s+/).map(w => w[0]).join('').toUpperCase().slice(0, 4);
  const s = agreementSettings(), when = new Date().toLocaleString('en-US', {timeZone: 'America/Los_Angeles'});
  const e = x => String(x || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const doc = `<html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:11pt;line-height:1.45;color:#111}h1{font-size:16pt}h2{font-size:11.5pt;margin:14px 0 4px}.sig{margin-top:22px;border-top:1px solid #000;padding-top:10px}.audit{margin-top:18px;font-size:8.5pt;color:#555}</style></head><body>
${s.live ? '' : '<p style="color:#b00;font-weight:bold">DRAFT FOR BROKER REVIEW: not a binding agreement.</p>'}
${agreementHtml(a.data, s)}
<div class="sig"><p><b>Buyer signature</b></p><img src="${sig}" style="height:70px"><p>${e(name)}<br>Signed ${e(when)} (Pacific)</p>
<p><b>Broker:</b> ${e(s.brokerage)} (broker ${e(s.brokerName)}, license ${e(s.brokerLicense)}), by ${e(s.licensee)}${s.licenseNo ? `, Nevada license #${e(s.licenseNo)}` : ''}</p></div>
<div style="page-break-before:always"></div>
${dutiesHtml(a.data, s, dual, initials)}
<div class="sig"><p><b>Buyer/Tenant</b></p><img src="${sig}" style="height:70px"><p>${e(name)}<br>Date/time: ${e(when)} (Pacific)</p></div>
<div class="audit">Signed electronically on juliancrider.com. Buyer chose: licensee ${e(dual)} act for two or more parties (initials ${e(initials)}). Agreement ID ${e(a.data.id)}. Buyer email ${e(a.data.email)}. Device: ${e(String(d.ua || '').slice(0, 200))}.</div></body></html>`;
  const pdf = Utilities.newBlob(doc, 'text/html', 'agreement.html').getAs('application/pdf')
    .setName(`Buyer Agreement - ${name} - ${String(a.data.property).slice(0, 60)}.pdf`.replace(/[\\/:*?"<>|]/g, ''));
  const file = driveFolder('Signed Buyer Agreements').createFile(pdf);
  a.sheet.getRange(a.row, 9, 1, 4).setValues([['Signed', when, cell(name), file.getUrl()]]);
  MailApp.sendEmail({to: a.data.email, cc: NOTIFY_EMAIL, replyTo: NOTIFY_EMAIL, name: 'Julián J. Crider',
    subject: 'Your signed buyer agreement: ' + String(a.data.property).slice(0, 60), attachments: [pdf],
    htmlBody: `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;font-size:15px;line-height:1.55;color:#1c1c1e"><p>Hi ${e(String(a.data.buyer).split(/\s+/)[0])},</p><p>Thanks for signing. Your copy of the one-home buyer agreement for <b>${e(a.data.property)}</b> is attached. I'll see you at the showing.</p><p>Julián J. Crider<br><span style="color:#6e6e73">702-498-2571 · juliancrider.com</span></p></div>`});
  try { tagInFub(a.data.email, ['Agreement signed'], `Signed single-property buyer agreement for ${a.data.property} on ${when}. PDF: ${file.getUrl()}`); }
  catch (err) { console.error('FUB tag failed', err); }
  return json({ok: true});
}

function tagInFub(email, tags, note) {
  if (!PropertiesService.getScriptProperties().getProperty('FUB_API_KEY') || !validEmail(email)) return;
  const found = fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people;
  if (!(found && found.length)) return;
  const merged = (found[0].tags || []).slice();
  tags.forEach(t => { if (merged.indexOf(t) < 0) merged.push(t); });
  fubFetch('people/' + found[0].id, {tags: merged}, 'put');
  if (note) fubFetch('notes', {personId: found[0].id, subject: 'Buyer agreement', body: note});
}

/** Creates a test agreement (works before AGREEMENT_LIVE) and logs the link for you and your broker to review. */
function previewAgreement() {
  const url = createAgreement({name: 'Test Buyer', email: NOTIFY_EMAIL, phone: '702-555-0100', property: '11616 Zagarolo Ln, Las Vegas, NV'});
  console.log('Open this to review and test-sign: ' + url);
}

/** Sends the buyer confirmation to you so you can see it (uses your notification email). */
function testShowingConfirmation() {
  confirmShowing({name: 'Julián Test', email: NOTIFY_EMAIL, property: '11616 Zagarolo Ln', timeline: 'This weekend',
    notes: 'Preferred time: Saturday after 2 pm | Pre-approved: Yes | Working with agent: No'});
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


// ============================================================
// TEXTING: blue iMessages from Julián's Mac (BlueBubbles server app, Messages signed in to his Apple ID).
// Script Properties: BB_URL (server address; BlueBubbles reports new ones by itself), BB_PASSWORD (server password),
// BB_KEY (made by setupTexting; secret part of the webhook URL). Run setupTexting once, paste the webhook URL it
// prints into BlueBubbles > API & Webhooks (events: New Messages, New Server URL), then run testText.
// Only contacts tagged "Text OK" are ever texted. "STOP" replies remove Text OK and add "Text Opt-Out".
// Texts from people who aren't in Follow Up Boss (friends, family) are ignored and never stored.
// ============================================================
const TEXTING = {webApp: 'https://script.google.com/macros/s/AKfycbzqOjOSjwfZynEC6uNUlhoewr9k34HKZhIfzKaUBScpcoZKHEQykJQJyEW7d9tETkK7/exec',
  myNumber: '+17024982571'};

function bbE164(phone) {
  const d = String(phone || '').replace(/\D/g, '');
  return d.length === 10 ? '+1' + d : d.length === 11 && d[0] === '1' ? '+' + d : '';
}

function bbApi(path, body) {
  const p = PropertiesService.getScriptProperties(), url = p.getProperty('BB_URL'), pw = p.getProperty('BB_PASSWORD');
  if (!url || !pw) throw new Error('Texting is not set up yet: add BB_URL and BB_PASSWORD in Script Properties.');
  const res = UrlFetchApp.fetch(url.replace(/\/+$/, '') + '/api/v1/' + path + (path.indexOf('?') < 0 ? '?' : '&') +
    'password=' + encodeURIComponent(pw), body ? {method: 'post', contentType: 'application/json', payload: JSON.stringify(body),
    muteHttpExceptions: true} : {muteHttpExceptions: true});
  const code = res.getResponseCode();
  if (code >= 300) {
    let why = res.getContentText();
    try { const j = JSON.parse(why); why = [j.message, j.error && (j.error.message || j.error)].filter(Boolean).join(': '); } catch (e) {}
    throw new Error('Texting Mac answered ' + code + ': ' + String(why).replace(/\\"/g, '"').slice(0, 900) +
      (code === 404 || code === 502 || code === 530 ? ' (is the Mac on and BlueBubbles open?)' : ''));
  }
  return JSON.parse(res.getContentText() || '{}');
}

/** Sends one iMessage (falls back to a green SMS for non-iPhones, if Text Message Forwarding is on). */
function bbSend(phone, text) {
  const to = bbE164(phone);
  if (!to) throw new Error('Not a US mobile number: ' + phone);
  // Existing conversation: send into it. New number: on this Mac's macOS (10.14) Messages can't address an unknown
  // "buddy", so start a new chat with the message instead. iMessage first, then green SMS (needs Text Message Forwarding).
  const tries = [
    ['iMessage, existing chat', () => bbApi('message/text', {chatGuid: 'iMessage;-;' + to, tempGuid: 'jc-' + Utilities.getUuid(), message: text, method: 'apple-script'})],
    ['iMessage, new chat', () => bbApi('chat/new', {addresses: [to], message: text, service: 'iMessage', method: 'apple-script', tempGuid: 'jc-' + Utilities.getUuid()})],
    ['SMS, existing chat', () => bbApi('message/text', {chatGuid: 'SMS;-;' + to, tempGuid: 'jc-' + Utilities.getUuid(), message: text, method: 'apple-script'})],
    ['SMS, new chat', () => bbApi('chat/new', {addresses: [to], message: text, service: 'SMS', method: 'apple-script', tempGuid: 'jc-' + Utilities.getUuid()})]];
  let last;
  for (const [label, fn] of tries) {
    try { const r = fn(); console.log('Texted ' + to + ' (' + label + ').'); return {service: label.split(',')[0], result: r}; }
    catch (err) { last = err; console.log(label + ' failed: ' + String(err.message).slice(0, 300)); }
  }
  throw last;
}

/** Texts a Follow Up Boss contact (only if tagged "Text OK") and logs it on their timeline. */
function bbTextPerson(personId, text, opts) {
  const p = fubFetch('people/' + personId + '?fields=id,firstName,tags,phones', null, 'get');
  const tags = p.tags || [];
  if (tags.indexOf('Text OK') < 0 || tags.indexOf('Text Opt-Out') >= 0) return 'skipped: no text consent';
  const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
  if (!phone) return 'skipped: no mobile number';
  const msg = String(text).replace(/%contact_first_name%/g, p.firstName || 'there');
  const sent = bbSendParts(phone, msg);
  // Optional voice note right after the text (same conversation), e.g. "Julians Open House Audio" in Google Drive.
  (opts && opts.files || []).forEach(name => {
    try { bbSendAudio(phone, sent.service, name); fubFetch('notes', {personId: p.id, subject: 'Sent by text: ' + name, body: name}); }
    catch (err) { console.log('Text sent, "' + name + '" not sent: ' + err.message); }
  });
  try { fubFetch('textMessages', {personId: p.id, message: msg, toNumber: phone, fromNumber: TEXTING.myNumber, isIncoming: false}); }
  catch (err) {  // FUB only takes textMessages when its own texting is on; save it as a note instead
    try { fubFetch('notes', {personId: p.id, subject: 'Text sent to ' + (p.firstName || 'contact'), body: msg}); }
    catch (e2) { console.log('Sent, but could not log the text in Follow Up Boss: ' + e2.message); }
  }
  return 'sent';
}

/** Sends "bubble || [voice] || bubble || [contact card]" as separate messages, in order. Returns the first send. */
function bbSendParts(phone, msg) {
  // [welcome voice] = Drive file "Welcome hello" (a general hello for online sign-ups); skipped quietly until it exists.
  const files = {'[voice]': OPEN_HOUSE_VOICE, '[welcome voice]': 'Welcome hello', '[contact card]': CONTACT_CARD};
  let first = null;
  String(msg).split(/\s*\|\|\s*/).filter(Boolean).forEach(part => {
    if (files[part.toLowerCase()]) {
      if (!first) throw new Error('The first bubble has to be text');
      try { bbSendAudio(phone, first.service, files[part.toLowerCase()]); Utilities.sleep(1500); }
      catch (err) { console.log(part + ' not sent: ' + err.message); }
    } else {
      const r = bbSend(phone, part); first = first || r; Utilities.sleep(1500);
    }
  });
  return first;
}

/** Sends an audio file from Google Drive (found by name) into the conversation with this number. */
function bbSendAudio(phone, service, driveName) {
  if (/^https?:\/\//.test(driveName)) {  // a file hosted on the website (e.g. the contact card)
    const blob = UrlFetchApp.fetch(driveName).getBlob().setName(/\.vcf$/i.test(driveName) ? 'Julián Crider.vcf' : driveName.split('/').pop());
    return bbPostAttachment(phone, service, blob);
  }
  const props = PropertiesService.getScriptProperties(), cacheKey = 'AUDIO_' + driveName;
  let file;
  try { file = props.getProperty(cacheKey) && DriveApp.getFileById(props.getProperty(cacheKey)); if (file && file.isTrashed()) file = null; } catch (e) {}
  if (!file) {
    const it = DriveApp.searchFiles('title contains "' + driveName.replace(/"/g, '') + '" and trashed = false');
    if (!it.hasNext()) throw new Error('No Google Drive file named "' + driveName + '"');
    file = it.next(); props.setProperty(cacheKey, file.getId());
  }
  const blob = file.getBlob(), name = /\.[a-z0-9]{2,4}$/i.test(file.getName()) ? file.getName()
    : file.getName() + (/vcard|vcf/i.test(blob.getContentType()) ? '.vcf' : '.m4a');
  return bbPostAttachment(phone, service, blob.setName(name));
}

function bbPostAttachment(phone, service, blob) {
  const name = blob.getName();
  const p = PropertiesService.getScriptProperties(), url = p.getProperty('BB_URL'), pw = p.getProperty('BB_PASSWORD');
  const res = UrlFetchApp.fetch(url.replace(/\/+$/, '') + '/api/v1/message/attachment?password=' + encodeURIComponent(pw), {
    method: 'post', muteHttpExceptions: true, payload: {chatGuid: (service || 'iMessage') + ';-;' + bbE164(phone),
      tempGuid: 'jc-' + Utilities.getUuid(), name: name, method: 'apple-script', attachment: blob}});
  if (res.getResponseCode() >= 300) throw new Error('Voice note: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
}

function bbWebhook(e) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BB_KEY') || e.parameter.bb !== props.getProperty('BB_KEY')) return json({status: 'ignored'});
  const ev = JSON.parse((e.postData && e.postData.contents) || '{}'), data = ev.data || {};
  if (ev.type === 'new-server') {
    const url = typeof data === 'string' ? data : data.url || data.serverAddress || '';
    if (/^https?:\/\//.test(url)) props.setProperty('BB_URL', url);
    return json({status: 'ok'});
  }
  if (ev.type === 'new-message' && data.isFromMe) {  // Julián (or Shawna) texted: cancel any reply Shawna was holding
    const to = bbE164((data.handle && data.handle.address) || ((data.chats || [])[0] || {}).chatIdentifier || '');
    if (to) {
      assistantCancel(to);
      const sentByShawna = Number(props.getProperty('AI_SENDING_' + to) || 0) > Date.now() - 90 * 1000;
      if (!sentByShawna) props.setProperty('AI_MUTE_' + to, String(Date.now() + Number(props.getProperty('ASSISTANT_MUTE_MIN') || 30) * 6e4));
    }
    return json({status: 'ok'});
  }
  if (ev.type !== 'new-message' || !data.handle || !data.text) return json({status: 'ok'});
  const from = bbE164(data.handle.address);
  if (!from) return json({status: 'ok'});  // email handles etc.
  // Julián texting his own business number from his personal phone = talking to Shawna (his assistant).
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  if (boss && from === boss) { runBossCommand(String(data.text).slice(0, 1500), 0); return json({status: 'ok'}); }
  const found = fubFetch('people?limit=1&fields=id,firstName,lastName,tags&phone=' + encodeURIComponent(from.slice(2)), null, 'get').people || [];
  if (!found.length) return json({status: 'ok'});  // not a Follow Up Boss contact: personal text, leave it alone
  const person = found[0], text = String(data.text).slice(0, 1000);
  try { fubFetch('textMessages', {personId: person.id, message: text, fromNumber: from, toNumber: TEXTING.myNumber, isIncoming: true}); }
  catch (err) { fubFetch('notes', {personId: person.id, subject: 'Text from ' + (person.firstName || 'contact'), body: text}); }
  if (/^\s*(stop|stopall|unsubscribe|end|quit|cancel|opt ?out)\b/i.test(text)) {
    const tags = (person.tags || []).filter(t => t !== 'Text OK').concat('Text Opt-Out');
    fubFetch('people/' + person.id, {tags: tags}, 'put');
    return json({status: 'ok'});
  }
  // "JULIAN": they want him, not Shawna. Tell Julián on his personal phone, answer them, skip Shawna.
  if (/^\s*juli[aá]n\s*[.!?]*\s*$/i.test(text)) {
    assistantCancel(from);
    try { bbSend(from, "Got it. I'm letting Julián know right now. He'll text you back or give you a call shortly."); } catch (e) {}
    notifyJulian((person.firstName || 'A contact') + ' asked for you', 'They texted JULIAN. Check your messages and reply or call.', 'https://app.followupboss.com/2/people/view/' + person.id);
    createCallTask(person.id, 'They asked for you directly (texted JULIAN). Reply or call now.', 'Text');
    return json({status: 'ok'});
  }
  try { notifyJulian([person.firstName, person.lastName].join(' ').trim() + ' texted you', '"' + text.slice(0, 140) + '" Check your messages.', 'https://app.followupboss.com/2/people/view/' + person.id); } catch (e) {}
  // Shawna (AI assistant) drafts or sends a reply. Never blocks the webhook if Groq is down.
  let handled = false;
  try { handled = assistantReply(person, from, text); } catch (err) { console.log('Assistant: ' + err.message); }
  // One "answer today" task per person per day (inbound-only rule: tasks only when they reach out).
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'), key = 'BB_TASK_' + person.id;
  if (props.getProperty(key) !== today) {
    if (!handled) createCallTask(person.id, 'They texted back: "' + text.slice(0, 120) + '" Answer today.', 'Text');
    props.setProperty(key, today);  // also tells the text queue they replied (skip the next scheduled text)
  }
  return json({status: 'ok'});
}

// ============================================================
// AI ASSISTANT ("Shawna" by default; ASSISTANT_NAME in Script Properties, so any agent can name their own).
// When a Follow Up Boss contact texts in, Groq writes a short reply in Julián's voice (Chris Voss style).
//   ASSISTANT_MODE = off   → nothing (default if no GROQ_API_KEY)
//                  = draft → (default) reply saved as a FUB note "Shawna draft" + in the "answer today" task;
//                            Julián sends it himself (or copies it into a "Send text" note).
//                  = auto  → Shawna texts the reply herself, 8am-9pm PT, max 3 per person per day, first message
//                            says she's Julián's AI assistant. Anything sensitive is escalated to Julián instead.
// Script Properties: GROQ_API_KEY (console.groq.com → API Keys), optional GROQ_MODEL, ASSISTANT_NAME, ASSISTANT_MODE.
// Returns true when the reply was sent automatically (so no task is needed).
// ============================================================
const ASSISTANT = {
  model: 'llama-3.3-70b-versatile',
  escalate: /\b(offer|counter|contract|escrow|earnest|commission|lawyer|attorney|sue|complain|discriminat|inspection report|appraisal)\b/i,
};
// What Shawna knows about Julián's business. Julián adds more by typing one fact per row in the "Shawna Knowledge" tab.
const SHAWNA_FACTS = `About Julián: Julián J. Crider, Realtor® (NV license S.0177760), The DeBenedetto Team at Boeckle & Co. (broker Landon Boeckle, B.1001882), 4775 W Teco Ave #220, Las Vegas, NV 89118, office 725-465-7646. Cell/text 702-498-2571. Email julianc@bcolv.com. Background: started as a listing agent at one of the largest brokerages in the country, then moved into development (land, construction, getting housing built in the valley), so he looks at a home's value five years out, build quality, and whether it's the right move for the money. Promise: straight talk even when it costs him the deal, does what he says, answers the phone.
Focus: luxury homes and relocation in Las Vegas and Henderson, especially out-of-state buyers (California, the West Coast, Chicago, Seattle and beyond). Areas: Summerlin, Henderson, MacDonald Highlands, The Ridges, Red Rock Country Club, The Summit, Southern Highlands, Lake Las Vegas, Desert Shores, Skye Canyon and the Northwest, new-construction communities. Golf communities are a specialty (he plays).
How he works: free 30-minute game-plan call (Zoom/Google Meet for out-of-town clients, or phone); previews homes before clients tour; live FaceTime showings for out-of-state buyers; remote offers and closings; comp reports like an appraiser; seller net sheets; buyer agreements signed online. Buyers: get pre-approved with a local lender (he can recommend one), then a home search, showings, offer, escrow, close.
Links: website juliancrider.com · book a call juliancrider.com/?book · request a showing juliancrider.com/?showing · home value juliancrider.com/?value · relocation guide juliancrider.com/?relocate · open-house sign-in juliancrider.com/?signin · Instagram @juliancriderestates (instagram.com/juliancriderestates).
The team (thedebenedettoteamlv.com, "The valley's quiet standard"): luxury representation for buyers and sellers in Las Vegas, Henderson and Summerlin. "A short bench, on purpose": people who answer their own phones, and every listing is handled by the person you met, from first walkthrough to closing. Joey DeBenedetto, Team Leader (runs the listing side and takes the call when a negotiation gets difficult; started selling luxury real estate in New York City and built a team with more than $200 million in luxury sales; CEO of Keller Williams Cupertino and Palo Alto; Redfin's first outside senior executive, shaping its Premier luxury division; Top 100 Most Influential People in Real Estate; Silicon Valley "30 Leaders to Watch"; Certified Negotiation Expert; Las Vegas Real Producers cover, September 2026; 2026 Top 25 Realtor in Las Vegas by closed production). Julián Crider, Senior Buyers Agent (business built almost entirely on referral; buyers who work with him see fewer houses and write fewer offers because the guidance comes before the showings). Tana Cogan, Senior Buyers Specialist (investor and renovator background; Summerlin, The Ridges and MacDonald Highlands; short-term rentals, long-term holds, primary homes and renovation projects). Team line 702-300-4349.
Team results as published on the team site: 64 closings in the twelve months from September 2025 to September 2026 (49 listings, 15 buyers); January 1 to August 21, 2026: 36 closings, $38.2M sold at 98.01% of list, median 19 days on market. A seller at Anthem Country Club: "They didn't just sell my house, they sold it in 3 days for full ask." A good share of what the team sells never reaches the open market.
Neighborhoods the team knows best: Summerlin (master-planned, walkable, deepest inventory, from the Vistas to Red Rock Country Club), The Ridges (guard-gated estates, custom lots, Strip views), Anthem Country Club (Henderson's golf-course guard gate, the team's most-sold neighborhood), MacDonald Highlands (hillside architecture, skyline views), Green Valley (established Henderson, mature trees), Centennial and the Northwest (newer construction, bigger lots, mountain access).
Home search process: set your criteria (price, beds, area, lot, view, single story, guard gate); the team sees what you saved; new matches reach you the moment they hit the MLS, and the ones worth seeing come with a personal note (not an automated drip).
Home valuation process: send the address (no listing agreement, nothing made public); a 45-minute in-person walkthrough; a written opinion of value with comparable evidence, a recommended range and an honest read on timing, yours to keep either way; the team replies within one business day.
Never quote list prices, rates or market stats beyond these; say "I'll confirm that for you" and pass it to Julián.`;

function assistantKnowledge() {
  let extra = '';
  try {
    const ss = getSS();
    const sh = ss.getSheetByName('Shawna Knowledge') || ensureTab(ss, {name: 'Shawna Knowledge', color: '#6B4FBB', headers: ['One fact per row (Shawna reads these)']});
    if (sh.getLastRow() > 1) extra = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean).join('\n');
  } catch (e) {}
  return SHAWNA_FACTS + (extra ? '\nMore facts from Julián:\n' + extra : '');
}

/** Removes the dashes that make texts read like AI (keeps phone numbers and hyphenated words). */
function humanize(t) {
  return String(t || '').replace(/\s*[—–]\s*/g, ', ').replace(/\s+-\s+/g, ', ').replace(/,\s*,/g, ',').trim();
}

/** First text Shawna sends anyone (ASSISTANT_INTRO overrides it; {name} = her name). */
function assistantIntro() {
  const t = PropertiesService.getScriptProperties().getProperty('ASSISTANT_INTRO') ||
    "Hi, it's {name}, Julián's AI assistant. Julián's away right now, but I can answer your questions. If you'd like to talk with Julián directly, just text JULIAN and he'll text you back or give you a call.";
  return t.replace(/\{name\}/g, assistantName());
}
function assistantName() { return PropertiesService.getScriptProperties().getProperty('ASSISTANT_NAME') || 'Shawna'; }

function assistantReply(person, from, text) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GROQ_API_KEY'), mode = (props.getProperty('ASSISTANT_MODE') || 'draft').toLowerCase();
  if (!key || mode === 'off') return false;
  const name = assistantName();
  const p = fubFetch('people/' + person.id + '?fields=id,firstName,lastName,tags,stage,price', null, 'get');
  const tags = p.tags || [];
  const notes = (fubFetch('notes?personId=' + person.id + '&limit=8', null, 'get').notes || []).reverse()
    .map(n => (n.subject ? n.subject + ': ' : '') + String(n.body || '').replace(/<[^>]+>/g, ' ').slice(0, 400)).join('\n');
  const r = groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: assistantPrompt(name)},
    {role: 'user', content: 'Contact: ' + [p.firstName, p.lastName].join(' ') + '\nStage: ' + (p.stage || '') + '\nTags: ' + tags.join(', ') +
      '\nRecent history (oldest first):\n' + notes + '\n\nThey just texted: "' + text + '"\n\nReply as JSON only.'}]);
  let out; try { out = JSON.parse(r); } catch (e) { out = {reply: '', escalate: true, reason: 'unreadable AI answer'}; }
  const reply = humanize(out.reply);
  const sensitive = ASSISTANT.escalate.test(text) || tags.indexOf('Has Agent') >= 0 || tags.indexOf('Text Opt-Out') >= 0;
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'), cnt = 'AI_CNT_' + person.id;
  const used = (props.getProperty(cnt) || '').split('|'), n = used[0] === today ? Number(used[1]) : 0;
  const muted = Number(props.getProperty('AI_MUTE_' + from) || 0) > Date.now();  // Julián is talking with them himself
  const ok = reply && !out.escalate && !sensitive && n < 3 && tags.indexOf('Text OK') >= 0 && !muted;
  if (mode === 'backup' && ok && hour >= 8 && hour < 21) {  // hold it; send only if Julián hasn't replied in time
    const delay = Number(props.getProperty('ASSISTANT_DELAY_MIN') || 2);
    const q = JSON.parse(props.getProperty('AI_PENDING') || '{}');
    q[from] = {personId: person.id, reply: reply, next: out.next || '', due: Date.now() + delay * 6e4};
    props.setProperty('AI_PENDING', JSON.stringify(q));
    return true;
  }
  const auto = mode === 'auto' && ok && hour >= 8 && hour < 21;
  if (!auto) {
    fubFetch('notes', {personId: person.id, subject: name + ' draft', body: (reply ? 'Suggested reply:\n' + reply : 'No reply drafted.') +
      (out.escalate || sensitive ? '\n\nNeeds you: ' + (out.reason || 'sensitive topic') : '') + (out.next ? '\nNext step: ' + out.next : '')});
    return false;
  }
  const intro = props.getProperty('AI_INTRO_' + person.id) ? '' : assistantIntro() + ' ';
  props.setProperty('AI_SENDING_' + from, String(Date.now()));
  bbSendParts(from, intro + reply);
  props.setProperty('AI_INTRO_' + person.id, '1'); props.setProperty(cnt, today + '|' + (n + 1));
  fubFetch('notes', {personId: person.id, subject: 'Text sent', body: intro + reply + '\n\n(sent by ' + name + ')' + (out.next ? '\nNext step: ' + out.next : '')});
  return true;
}

function assistantPrompt(name) {
  return `You are ${name}, the AI assistant for Julián J. Crider, a Realtor® with Boeckle & Co. in Las Vegas and Henderson, Nevada (NV S.0177760, phone 702-498-2571).
You are also a top luxury real estate advisor in your own right: confident, polished, warm, and genuinely helpful.
${assistantKnowledge()}
You answer text messages from his contacts. Never use dashes of any kind (no "—", "–" or " - "); use commas or periods. Write like Julián: short (1-3 sentences), warm, plain, no corporate words, at most one emoji (👋🏾 or 🏡). Use Chris Voss style: a label ("Sounds like…", "It seems…"), a no-oriented question ("Would it be a bad idea if…?") or a calibrated what/how question. Ask at most one question.
Sound human and trustworthy, like someone who will tell them the truth: it's fine to say no or "this one isn't for everyone" when a home or idea doesn't fit them.
Word choices (always): instead of "sorry, I can't" say "here's what I can do for you"; instead of "if you decide to" say "when we work together"; instead of "I'm not sure" say "I'll confirm that for you"; never say "sorry to bother you", say "I was thinking of you" or "I thought of you"; instead of "what happens next" say "the next steps are". Don't apologize unless something actually went wrong.
Goals: understand what they need, and when they're ready, offer a time to talk (juliancrider.com/?book) or a showing (juliancrider.com/?showing). Home value questions: juliancrider.com/?value.
Rules: never invent facts, prices, availability, rates, or promises. Never give legal, tax or lending advice; suggest talking with Julián or their CPA/lender. Fair Housing: never comment on who lives in an area, schools' quality, safety/crime, religion, race, family status, disability or any protected class; point to public sources instead. If they mention an offer, contract, negotiation, commission, a complaint, anything legal or upset, or anything you're unsure about, set escalate true and write a short holding reply ("Got it, I'll have Julián reach out today.").
Return JSON only: {"reply": "...", "escalate": true|false, "reason": "why Julián is needed (if escalate)", "next": "the next step for Julián, short"}`;
}

function groqChat(key, model, messages) {
  const send = m => UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {method: 'post', contentType: 'application/json',
    muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + key},
    payload: JSON.stringify({model: m, messages: messages, temperature: 0.4, max_tokens: 400, response_format: {type: 'json_object'}})});
  let res = send(model);
  if (res.getResponseCode() === 404 || /model_not_found|decommissioned/.test(res.getContentText())) {
    model = groqPickModel(key);  // Groq retires models now and then: pick the best one this account has and remember it
    PropertiesService.getScriptProperties().setProperty('GROQ_MODEL', model);
    res = send(model);
  }
  if (res.getResponseCode() >= 300) throw new Error('Groq ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  return JSON.parse(res.getContentText()).choices[0].message.content;
}

/** Lists the chat models on this Groq account and returns the best fit (big general models first). */
function groqPickModel(key) {
  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/models', {muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + key}});
  if (res.getResponseCode() >= 300) throw new Error('Groq models ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  const ids = (JSON.parse(res.getContentText()).data || []).filter(m => m.active !== false).map(m => m.id)
    .filter(id => !/whisper|tts|guard|embed|playai|orpheus|compound|prompt/i.test(id));
  console.log('Groq models on this account: ' + ids.join(', '));
  const prefs = [/gpt-oss-120b/, /llama.*70b/, /llama-4.*maverick/, /kimi/, /qwen.*3.*32b/, /llama-4.*scout/, /gpt-oss/, /llama/];
  for (const re of prefs) { const hit = ids.find(id => re.test(id)); if (hit) return hit; }
  if (!ids.length) throw new Error('No Groq chat models available on this account');
  return ids[0];
}

function assistantCancel(phone) {
  const props = PropertiesService.getScriptProperties(), q = JSON.parse(props.getProperty('AI_PENDING') || '{}');
  if (q[phone]) { delete q[phone]; props.setProperty('AI_PENDING', JSON.stringify(q)); }
}

/** Every minute: sends the replies Shawna was holding (backup mode) once Julián hasn't answered within the delay. */
function processAssistantQueue() {
  const props = PropertiesService.getScriptProperties();
  const retry = JSON.parse(props.getProperty('BOSS_RETRY') || '[]');
  if (retry.length) {
    const due = retry.filter(r => r.at <= Date.now()), later = retry.filter(r => r.at > Date.now());
    props.setProperty('BOSS_RETRY', JSON.stringify(later));
    due.forEach(r => runBossCommand(r.text, r.tries));
  }
  const raw = props.getProperty('AI_PENDING');
  if (!raw || raw === '{}') return;
  const q = JSON.parse(raw), name = assistantName(), today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
  Object.keys(q).forEach(phone => {
    const it = q[phone]; if (it.due > Date.now()) return;
    delete q[phone]; props.setProperty('AI_PENDING', JSON.stringify(q));  // remove first: never double-send
    try {
      const intro = props.getProperty('AI_INTRO_' + it.personId) ? '' : assistantIntro() + ' ';
      if (Number(props.getProperty('AI_MUTE_' + phone) || 0) > Date.now()) return;  // he jumped in meanwhile
      props.setProperty('AI_SENDING_' + phone, String(Date.now()));
      bbSendParts(phone, intro + it.reply);
      const cnt = 'AI_CNT_' + it.personId, used = (props.getProperty(cnt) || '').split('|');
      props.setProperty('AI_INTRO_' + it.personId, '1'); props.setProperty(cnt, today + '|' + ((used[0] === today ? Number(used[1]) : 0) + 1));
      fubFetch('notes', {personId: it.personId, subject: 'Text sent', body: intro + it.reply + '\n\n(sent by ' + name + ', you hadn\'t replied yet)' + (it.next ? '\nNext step: ' + it.next : '')});
      try {
        const who = (fubFetch('people/' + it.personId + '?fields=firstName', null, 'get').firstName) || 'them';
        notifyJulian('You received a text from ' + who, "I'm speaking with " + who + ' now: "' + it.reply.slice(0, 140) + '" If you want to jump in, just reply to ' + who + " yourself and I'll stay out of that conversation for 30 minutes.");
      } catch (e) {}
    } catch (err) {
      fubFetch('notes', {personId: it.personId, subject: name + ' draft', body: 'Could not send (' + err.message + '). Suggested reply:\n' + it.reply});
      createCallTask(it.personId, 'They texted and ' + name + ' couldn\'t reply. Answer today.', 'Text');
    }
  });
}

// ============================================================
// EMAIL ALERTS: Follow Up Boss logs emails from contacts (status "Received"; it can't show us the text). Every 5
// minutes we look at contacts with fresh activity; a new email from one of them sends Julián a phone push
// notification (free ntfy app, topic in Script Property NTFY_TOPIC) and, if he still hasn't emailed back after
// EMAIL_NUDGE_MIN (default 60) minutes, one reminder. Push only between 7am and 10pm PT; later ones wait for morning.
// ============================================================
function notifyJulian(title, msg, url) {
  // Two ways to reach Julián: a text from his business number to his personal phone (ALERT_TEXT_TO, falls back to
  // TEST_TEXT_TO; a different phone, so it actually alerts him) and/or a push through the free ntfy app (NTFY_TOPIC).
  const props = PropertiesService.getScriptProperties(), topic = props.getProperty('NTFY_TOPIC');
  const to = props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO');
  let sent = false;
  if (to) { try { bbSend(to, title + '. ' + msg + (url ? ' ' + url : '')); sent = true; } catch (e) { console.log('Alert text: ' + e.message); } }
  if (topic) {
    UrlFetchApp.fetch('https://ntfy.sh/' + topic, {method: 'post', payload: msg, muteHttpExceptions: true,
      headers: Object.assign({Title: title.replace(/[^\x20-\x7E]/g, ''), Priority: 'high', Tags: 'email'}, url ? {Click: url} : {})});
    sent = true;
  }
  return sent;
}

function checkClientEmails() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('NTFY_TOPIC') && !props.getProperty('ALERT_TEXT_TO') && !props.getProperty('TEST_TEXT_TO')) return;
  const me = Number(props.getProperty('FUB_ME') || 0) || (() => { const id = fubFetch('me', null, 'get').id; props.setProperty('FUB_ME', id); return id; })();
  const since = Date.now() - 12 * 6e4, hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  const watch = JSON.parse(props.getProperty('EMAIL_WATCH') || '{}');  // emailId -> {p, name, at, nudged}
  const people = (fubFetch('people?limit=25&sort=-lastActivity&fields=id,name,lastActivity,assignedUserId', null, 'get').people || [])
    .filter(p => new Date(p.lastActivity).getTime() > since && (!p.assignedUserId || p.assignedUserId === me));
  people.forEach(p => {
    const em = (fubFetch('emails?personId=' + p.id + '&limit=3&fields=id,status,date', null, 'get').emails || []);
    const last = em[0];
    if (last && last.status === 'Received' && !watch[last.id] && new Date(last.date).getTime() > Date.now() - 6 * 36e5)
      watch[last.id] = {p: p.id, name: p.name, at: new Date(last.date).getTime(), told: 0, nudged: 0};
  });
  const nudgeMin = Number(props.getProperty('EMAIL_NUDGE_MIN') || 60);
  Object.keys(watch).forEach(id => {
    const w = watch[id], link = 'https://app.followupboss.com/2/people/view/' + w.p;
    if (Date.now() - w.at > 2 * 864e5) { delete watch[id]; return; }
    if (hour < 7 || hour >= 22) return;
    if (!w.told) { notifyJulian(w.name + ' emailed you', 'New email from ' + w.name + '. Tap to open them in Follow Up Boss.', link); w.told = 1; return; }
    if (!w.nudged && Date.now() - w.at > nudgeMin * 6e4) {
      const replied = (fubFetch('emails?personId=' + w.p + '&limit=1&fields=id,status,date', null, 'get').emails || [])
        .some(e => e.status === 'Sent' && new Date(e.date).getTime() > w.at);
      if (!replied) notifyJulian('Still waiting: ' + w.name, w.name + ' emailed you ' + Math.round((Date.now() - w.at) / 6e4) + ' min ago and hasn\'t heard back yet.', link);
      w.nudged = 1;
    }
  });
  props.setProperty('EMAIL_WATCH', JSON.stringify(watch));
}

/** Sends a test push to Julián's phone (after adding NTFY_TOPIC and subscribing to it in the ntfy app). */
function testNotify() {
  console.log(notifyJulian('I am connected', 'You will get a message like this when a client emails you or I answer a text for you.') ? 'Sent. Check your phone.' : 'Add ALERT_TEXT_TO or NTFY_TOPIC first.');
}

// ============================================================
// TALK TO SHAWNA: Julián texts his business number from his personal phone (ALERT_TEXT_TO / TEST_TEXT_TO) and Shawna
// acts as his assistant. She understands plain English (Groq) and can: text a contact (always asks "Reply YES" first),
// add a note, create a task, list who emailed him, read today's calendar, or save a bigger request for Claude
// ("Requests for Claude" sheet tab; Julián opens the Claude app to have it built). Anything else: a short answer.
// ============================================================
const REQUESTS_TAB = {name: 'Requests for Claude', color: '#6B4FBB', headers: ['When', 'Request', 'Status']};

function bossCommand(text) {
  const props = PropertiesService.getScriptProperties(), name = assistantName();
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO'));
  const say = msg => bbSend(boss, humanize(String(msg).replace(new RegExp('^' + name + ':\\s*'), '')));
  // A pending "send this text?" waiting for YES / NO
  const pend = JSON.parse(props.getProperty('BOSS_PENDING') || 'null');
  if (pend && pend.kind !== 'schedule' && /^\s*(yes|y|send|send it|go|ok|okay)\s*[.!]*\s*$/i.test(text)) {
    props.deleteProperty('BOSS_PENDING');
    if (Date.now() - pend.at > 30 * 6e4) return say(name + ': that one expired. Ask me again.');
    if (pend.kind === 'voice') {
      const audio = shawnaVoice(pend.message);
      props.setProperty('AI_SENDING_' + pend.phone, String(Date.now()));
      const first = String(pend.who || '').split(' ')[0];
      const r = bbSend(pend.phone, 'Hi' + (first ? ' ' + first : '') + ', it\'s Shawna, Julián\'s assistant. Here\'s a quick voice note 👇🏾');
      Utilities.sleep(1500);
      bbPostAttachment(pend.phone, r && r.service, audio);
      fubFetch('notes', {personId: pend.personId, subject: 'Voice memo sent', body: pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
      return say('voice memo sent to ' + pend.who + '. ✅');
    }
    if (pend.kind === 'email') {
      const r = sendFromBcolv(pend.email, pend.subject, pend.message);
      if (!r.ok) return say('the email did NOT go out: ' + (r.error || 'unknown error') + '.');
      fubFetch('notes', {personId: pend.personId, subject: 'Email sent', body: 'To ' + pend.email + ' from ' + (r.from || 'bcolv') + ': "' + pend.subject + '"\n\n' + pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
      return say('emailed ' + pend.who + ' from ' + (r.from || 'your bcolv email') + '. ✅');
    }
    props.setProperty('AI_SENDING_' + pend.phone, String(Date.now()));
    bbSendParts(pend.phone, pend.message);
    fubFetch('notes', {personId: pend.personId, subject: 'Text sent', body: pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
    return say(name + ': sent to ' + pend.who + '. ✅');
  }
  if (pend && /^\s*(no|n|cancel|stop|don.?t)\b/i.test(text)) { props.deleteProperty('BOSS_PENDING'); return say(name + ': canceled. Nothing was sent.'); }

  const key = props.getProperty('GROQ_API_KEY');
  if (!key) return say(name + ': add GROQ_API_KEY first.');
  const sys = `You are ${name}, the personal AI assistant of Julián Crider, a Las Vegas Realtor. Julián is texting you. Decide what he wants and return JSON only:
{"action":"text|voicememo|email|note|task|contact|schedule|lookup|hot|emails|calendar|callme|claude|chat", "who":"contact's full name if any", "phone":"their phone if he gave one", "email":"their email if he gave one", "subject":"email subject (for email)", "message":"the text or email body to send / note / task / event title", "when":"for schedule: start as YYYY-MM-DDTHH:mm, Las Vegas time", "minutes":30, "calendar":"personal|business|unsaid (personal = julian.crider gmail, business = bcolv work calendar; unsaid if he didn't say)", "reply":"your short answer to Julián (for chat)"}
Today is ${Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE yyyy-MM-dd HH:mm')} in Las Vegas.
${assistantKnowledge()}
- text: he wants you to text a client (e.g. "text Lisa and ask if 3pm works"). Write the message in his voice: short, warm, Chris Voss style, no apology, "I was thinking of you" not "sorry to bother you".
- voicememo: he wants you to send someone a voice memo / voice note / audio message in your voice ("message" = what you'll say, written to be spoken, 15 to 40 seconds).
- email: he wants you to email someone (write a short, warm email in his voice, plain paragraphs, no signature; it's signed automatically).
- note: save a note on a contact. task: remind him to do something about a contact.
- emails: who emailed him / unanswered emails. calendar: what's on his schedule today.
- contact: add a new person to Follow Up Boss (name, phone, email).
- schedule: put an appointment on his Google Calendar (who, when, minutes, message = what it is, phone if given). New people are added to Follow Up Boss too.
- lookup: tell him about a contact from Follow Up Boss (stage, tags, recent notes).
- hot: list his hottest leads right now.
- callme: he wants you to phone him (e.g. "call me", "give me a call"). Put anything he wants to talk about in "message".
- Never use dashes in anything you write.
- claude: anything you can't do with the actions above, and anything bigger: website changes, designs, emails campaigns, social posts, videos, reports, code, Follow Up Boss setup. Put his request in "message".
- chat: anything else; answer in 1-3 short sentences. Never invent facts about clients or listings.`;
  let a;
  const calPick = pend && pend.kind === 'schedule' && pickCalendar(text);
  if (calPick) { props.deleteProperty('BOSS_PENDING'); a = Object.assign(pend.a, {calendar: calPick}); }
  else {
    try { a = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [{role: 'system', content: sys}, {role: 'user', content: text}])); }
    catch (e) { return say(name + ': I didn\'t catch that. Try again?'); }
  }
  const findOne = who => {
    const q = String(who || '').trim(), get = term => fubFetch('people?limit=8&fields=id,name,firstName,lastName,tags,phones,created&q=' + encodeURIComponent(term), null, 'get').people || [];
    let ppl = q ? get(q) : [];
    if (!ppl.length && q.split(/\s+/).length > 1) {  // "Lily Montenegro" → search the last name, keep first names that start the same
      const parts = q.toLowerCase().split(/\s+/), first = parts[0].slice(0, 3);
      ppl = get(parts[parts.length - 1]).filter(p => String(p.firstName || p.name || '').toLowerCase().indexOf(first) === 0);
    }
    const exact = ppl.filter(p => (p.name || '').toLowerCase() === q.toLowerCase());
    if (exact.length) return exact.sort((x, y) => String(x.created).localeCompare(String(y.created)))[0];  // oldest = the real one
    return ppl.length === 1 ? ppl[0] : (ppl.length ? {many: ppl} : null);
  };
  const addContact = () => {
    const parts = String(a.who || '').trim().split(/\s+/);
    const r = fubFetch('events', {source: 'Julián (via ' + name + ')', system: 'juliancrider.com', type: 'General Inquiry',
      message: 'Added by text from Julián' + (a.message ? ': ' + a.message : ''), person: {firstName: parts[0] || 'Unknown', lastName: parts.slice(1).join(' '),
      phones: a.phone ? [{value: a.phone}] : [], emails: a.email ? [{value: a.email}] : [], tags: ['Added by ' + name]}});
    let id = r.personId || (r.person && r.person.id) || r.id;
    if (!id) {  // some event responses don't include the person: find them by the phone or email we just sent
      const digits = String(a.phone || '').replace(/\D/g, '').slice(-10);
      const found = (digits.length === 10 && fubFetch('people?limit=1&fields=id&phone=' + digits, null, 'get').people) ||
        (a.email && fubFetch('people?limit=1&fields=id&email=' + encodeURIComponent(a.email), null, 'get').people) || [];
      id = found.length && found[0].id;
    }
    return id ? Object.assign({id: id, name: a.who, tags: ['Added by ' + name]}, a.phone ? {phones: [{value: a.phone}]} : {}) : null;
  };
  // Who is this about? A phone number Julián gives wins (search by it); then the name; if nobody matches, add them.
  // Returns {p, isNew}. A person found by name who has no phone gets the number he gave.
  const resolvePerson = () => {
    const digits = String(a.phone || '').replace(/\D/g, '').slice(-10);
    if (digits.length === 10) {
      const byPhone = (fubFetch('people?limit=1&fields=id,name,firstName,tags,phones&phone=' + digits, null, 'get').people || [])[0];
      if (byPhone) return {p: byPhone, isNew: false};
      const named = a.who ? findOne(a.who) : null;
      if (named && !named.many && !(named.phones || []).length) {
        fubFetch('people/' + named.id, {phones: [{value: digits, type: 'mobile'}]}, 'put');
        named.phones = [{value: digits}];
        return {p: named, isNew: false};
      }
      return {p: addContact(), isNew: true};  // new lead, or a different person with the same name
    }
    const named = a.who ? findOne(a.who) : null;
    if (named) return {p: named, isNew: false};
    if (a.who && !a.phone && !a.email) {  // not in FUB: try the contacts on his phone (synced to the texting Mac)
      const c = macContact(a.who);
      if (c) { a.phone = c.phone; a.email = a.email || c.email; a.who = c.name || a.who; }
    }
    return {p: a.who ? addContact() : null, isNew: true};
  };
  if (a.action === 'contact') {
    const p = addContact();
    if (!p) fubFetch('notes', {personId: Number(props.getProperty('REQUESTS_PERSON_ID') || 2221), subject: 'Claude request', body: 'Add contact: ' + text});
    return say(p ? 'added ' + a.who + ' to Follow Up Boss. ✅' : 'I couldn\'t add ' + (a.who || 'them') + ', so I passed it to Claude.');
  }
  if (a.action === 'email') {
    if (!PropertiesService.getScriptProperties().getProperty('BCOLV_MAILER_URL')) return say('email isn\'t connected yet. Set up the bcolv mailer first (Claude has the steps).');
    const {p} = resolvePerson();
    if (!p || p.many) return say(p && p.many ? 'which one? ' + p.many.map(x => x.name).join(', ') : 'who should I email? Send me their name and email.');
    let email = a.email;
    if (!email) { const full = fubFetch('people/' + p.id + '?fields=emails', null, 'get'); email = ((full.emails || []).filter(e => !/invalid|bounced/i.test(e.status || ''))[0] || {}).value; }
    if (!email) return say((p.name || a.who) + ' has no email yet. Send me their email and I\'ll write it.');
    const subject = humanize(a.subject || 'Quick note from Julián'), body = humanize(a.message || '');
    props.setProperty('BOSS_PENDING', JSON.stringify({kind: 'email', personId: p.id, who: p.name || a.who, email: email, subject: subject, message: body, at: Date.now()}));
    return say('send this email to ' + (p.name || a.who) + ' (' + email + ')?\n\nSubject: ' + subject + '\n\n' + body + '\n\nReply YES to send or NO to cancel.');
  }
  if (a.action === 'voicememo') {
    if (!props.getProperty('ELEVENLABS_API_KEY') || !props.getProperty('ELEVENLABS_VOICE_ID')) return say('voice memos aren\'t connected yet. Add ELEVENLABS_API_KEY and ELEVENLABS_VOICE_ID (Claude has the steps).');
    const {p} = resolvePerson();
    if (!p || p.many) return say(p && p.many ? 'which one? ' + p.many.map(x => x.name).join(', ') : 'who should get the voice memo? Send me their name and number.');
    const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
    if (!phone) return say((p.name || a.who) + ' has no mobile number yet. Send me their number.');
    props.setProperty('BOSS_PENDING', JSON.stringify({kind: 'voice', personId: p.id, who: p.name || a.who, phone: phone, message: humanize(a.message), at: Date.now()}));
    return say('send ' + (p.name || a.who) + ' a voice memo saying:\n\n"' + humanize(a.message) + '"\n\nReply YES to send or NO to cancel.');
  }
  if (a.action === 'text' || a.action === 'note' || a.action === 'task') {
    let {p, isNew} = resolvePerson();
    if (!p) {
      fubFetch('notes', {personId: Number(props.getProperty('REQUESTS_PERSON_ID') || 2221), subject: 'Claude request', body: text});
      return say('I couldn\'t find or add "' + (a.who || '?') + '", so I passed it to Claude.');
    }
    if (isNew && a.action !== 'text') say('(' + a.who + ' wasn\'t in Follow Up Boss, so I added them.)');
    if (p.many) return say(name + ': which one? ' + p.many.map(x => x.name).join(', '));
    if (a.action === 'note') { fubFetch('notes', {personId: p.id, subject: 'Note from Julián', body: a.message}); return say(name + ': note saved on ' + p.name + '.'); }
    if (a.action === 'task') { createCallTask(p.id, a.message, 'Follow Up'); return say(name + ': task added for ' + p.name + ' (due today).'); }
    const tags = p.tags || [];
    if (tags.indexOf('Text Opt-Out') >= 0) return say(name + ': ' + p.name + ' opted out of texts, so I can\'t text them. Here\'s what I can do for you: add a task to call instead.');
    const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
    if (!phone) return say(p.name + ' has no mobile number yet. Send me their number and I\'ll text them.');
    props.setProperty('BOSS_PENDING', JSON.stringify({personId: p.id, who: p.name, phone: phone, message: a.message, at: Date.now()}));
    return say(name + ': send this to ' + p.name + (tags.indexOf('Text OK') < 0 ? ' (heads up: no Text OK tag)' : '') + '?\n\n"' + a.message + '"\n\nReply YES to send or NO to cancel.');
  }
  if (a.action === 'schedule') {
    const m = String(a.when || '').match(/(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
    if (!m) return say('what day and time should I put it on your calendar?');
    const start = Utilities.parseDate(m[1] + ' ' + m[2], 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
    const end = new Date(start.getTime() + Number(a.minutes || 30) * 6e4);
    const which = pickCalendar(a.calendar || '');
    if (!which) {
      props.setProperty('BOSS_PENDING', JSON.stringify({kind: 'schedule', a: a, at: Date.now()}));
      return say('personal or business calendar for this one?');
    }
    const cal = whichCalendar(which);
    if (!cal) return say('I can\'t reach your business calendar yet. Share the bcolv calendar with julian.crider@gmail.com (Make changes to events), then ask me again.');
    let p = (a.who || a.phone) ? resolvePerson().p : null;
    if (p && p.many) p = null;
    const title = (a.message || 'Appointment') + (a.who ? ': ' + a.who : '');
    const phone = a.phone || (p && p.phones && p.phones[0] && p.phones[0].value) || '';
    cal.createEvent(title, start, end, {description: (phone ? 'Phone: ' + phone + '\n' : '') + 'Added by ' + name + ' from Julián\'s text.'});
    if (p && p.id) fubFetch('notes', {personId: p.id, subject: 'Appointment', body: title + ' on ' + Utilities.formatDate(start, 'America/Los_Angeles', 'EEE MMM d, h:mm a')});
    return say('on your ' + which + ' calendar: ' + title + ', ' + Utilities.formatDate(start, 'America/Los_Angeles', 'EEE MMM d, h:mm a') + '. ✅');
  }
  if (a.action === 'lookup') {
    const p = findOne(a.who);
    if (!p) return say('I couldn\'t find "' + (a.who || '?') + '" in Follow Up Boss.');
    if (p.many) return say('which one? ' + p.many.map(x => x.name).join(', '));
    const full = fubFetch('people/' + p.id + '?fields=id,name,stage,tags,source,phones,emails,lastActivity', null, 'get');
    const notes = (fubFetch('notes?personId=' + p.id + '&limit=3', null, 'get').notes || []).map(n => '• ' + (n.subject ? n.subject + ': ' : '') + String(n.body || '').replace(/<[^>]+>/g, ' ').slice(0, 120)).join('\n');
    return say(full.name + ' (' + (full.stage || 'Lead') + (full.source ? ', ' + full.source : '') + ')\n' + [((full.phones || [])[0] || {}).value, ((full.emails || [])[0] || {}).value].filter(Boolean).join(' · ') +
      '\nTags: ' + (full.tags || []).slice(0, 8).join(', ') + (notes ? '\nLatest:\n' + notes : ''));
  }
  if (a.action === 'hot') {
    const ppl = fubFetch('people?limit=8&sort=-lastActivity&tags=' + encodeURIComponent('Tier A') + '&fields=id,name,lastActivity', null, 'get').people || [];
    return say(ppl.length ? 'your hottest leads:\n' + ppl.map(x => '• ' + x.name).join('\n') : 'no Tier A leads right now.');
  }
  if (a.action === 'callme') {
    const r = vapiCall(boss, a.message || '');
    return say(r.ok ? 'calling you now 📞' : 'I couldn\'t start the call: ' + r.error);
  }
  if (a.action === 'emails') {
    const w = JSON.parse(props.getProperty('EMAIL_WATCH') || '{}'), list = Object.keys(w).map(k => w[k]).sort((x, y) => y.at - x.at).slice(0, 6);
    return say(name + ': ' + (list.length ? 'recent emails from clients:\n' + list.map(x => '• ' + x.name + ' (' + Utilities.formatDate(new Date(x.at), 'America/Los_Angeles', 'EEE h:mm a') + ')').join('\n') : 'no new client emails in the last 2 days.'));
  }
  if (a.action === 'calendar') {
    const now = new Date(), end = new Date(now); end.setHours(23, 59, 59);
    const evs = [];
    ['personal', 'business'].forEach(w => { const c = whichCalendar(w); if (c) c.getEvents(now, end).forEach(e => evs.push({e: e, w: w})); });
    evs.sort((x, y) => x.e.getStartTime() - y.e.getStartTime());
    return say(evs.length ? 'the rest of today:\n' + evs.map(x => '• ' + Utilities.formatDate(x.e.getStartTime(), 'America/Los_Angeles', 'h:mm a') + ' ' + x.e.getTitle() + (x.w === 'business' ? ' (business)' : '')).join('\n') : 'nothing else on your calendars today.');
  }
  if (a.action === 'claude') {
    const sh = getSS().getSheetByName(REQUESTS_TAB.name) || ensureTab(getSS(), REQUESTS_TAB);
    sh.appendRow([Utilities.formatDate(new Date(), 'America/Los_Angeles', 'M/d h:mm a'), cell(a.message || text), 'new']);
    // Also a FUB note on Julián's own test contact, where Claude's hourly check picks it up (and texts him back there).
    fubFetch('notes', {personId: Number(props.getProperty('REQUESTS_PERSON_ID') || 2221), subject: 'Claude request', body: a.message || text});
    return say(name + ': got it ✅ "' + (a.message || text).slice(0, 120) + '". Claude checks every hour and will text you when it\'s done.');
  }
  return say(name + ': ' + (a.reply || 'Got it.'));
}

/** Runs one of Julián's texts to Shawna with a log and a safety net: if it fails (Mac busy, Groq hiccup), it's queued
 *  and retried by processAssistantQueue a minute later (up to 3 tries), and he's told if it still fails. */
function runBossCommand(text, tries) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {}  // one of his messages at a time, so they don't trip over each other
  try {
    shawnaLog('in', text, tries ? 'retry ' + tries : '');
    bossCommand(text);
    shawnaLog('ok', text, '');
  } catch (err) {
    shawnaLog('error', text, err.message);
    const props = PropertiesService.getScriptProperties(), q = JSON.parse(props.getProperty('BOSS_RETRY') || '[]');
    if (tries < 3) { q.push({text: text, tries: tries + 1, at: Date.now() + 60000}); props.setProperty('BOSS_RETRY', JSON.stringify(q)); }
    else {
      try { bbSend(bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO')), 'I couldn\'t finish "' + text.slice(0, 60) + '" (' + err.message.slice(0, 120) + '). Try again in a minute?'); }
      catch (e) { fubFetch('notes', {personId: Number(props.getProperty('REQUESTS_PERSON_ID') || 2221), subject: 'Claude request', body: 'Shawna failed 3 times on: ' + text + '\nError: ' + err.message}); }
    }
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/** "Shawna Log" tab: every message Julián sends her, and what happened (so problems can be traced). Keeps 500 rows. */
function shawnaLog(kind, text, detail) {
  try {
    const ss = getSS(), sh = ss.getSheetByName('Shawna Log') || ensureTab(ss, {name: 'Shawna Log', color: '#6B4FBB', headers: ['When', 'What', 'Message', 'Detail']});
    sh.appendRow([Utilities.formatDate(new Date(), 'America/Los_Angeles', 'M/d h:mm:ss a'), kind, cell(text.slice(0, 300)), cell(String(detail || '').slice(0, 300))]);
    if (sh.getLastRow() > 600) sh.deleteRows(2, 100);
  } catch (e) {}
}

/** "personal" / "business" from Julián's words (or null if he didn't say). */
function pickCalendar(t) {
  t = String(t || '').toLowerCase();
  if (/business|work|bcolv|office|team/.test(t)) return 'business';
  if (/personal|julian\.crider|gmail|my own|private/.test(t)) return 'personal';
  return null;
}

/** Personal = this account's main calendar (julian.crider@gmail.com). Business = BUSINESS_CALENDAR_ID (default the
 *  bcolv calendar, which must be shared with julian.crider@gmail.com with "Make changes to events"). */
function whichCalendar(which) {
  if (which !== 'business') return CalendarApp.getDefaultCalendar();
  const id = PropertiesService.getScriptProperties().getProperty('BUSINESS_CALENDAR_ID') || 'julianc@bcolv.com';
  try { return CalendarApp.getCalendarById(id) || (CalendarApp.subscribeToCalendar(id), CalendarApp.getCalendarById(id)); } catch (e) { return null; }
}

/** Shawna phones a number through Vapi (VAPI_API_KEY = Vapi private key, VAPI_ASSISTANT_ID, VAPI_PHONE_NUMBER_ID).
 *  Only used for Julián himself for now: AI calls to clients need their prior consent (FCC). */
function vapiCall(number, topic) {
  const p = PropertiesService.getScriptProperties(), key = p.getProperty('VAPI_API_KEY');
  if (!key || !p.getProperty('VAPI_ASSISTANT_ID') || !p.getProperty('VAPI_PHONE_NUMBER_ID')) return {ok: false, error: 'add VAPI_API_KEY, VAPI_ASSISTANT_ID and VAPI_PHONE_NUMBER_ID in Script Properties'};
  const body = {assistantId: p.getProperty('VAPI_ASSISTANT_ID'), phoneNumberId: p.getProperty('VAPI_PHONE_NUMBER_ID'), customer: {number: number, name: 'Julian'},
    assistantOverrides: {firstMessage: 'Hey Julian, it\'s Shawna. ' + (topic ? 'You wanted to talk about ' + topic + '. ' : '') + 'What can I do for you?'}};
  const res = UrlFetchApp.fetch('https://api.vapi.ai/call', {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Bearer ' + key}, payload: JSON.stringify(body)});
  if (res.getResponseCode() >= 300) return {ok: false, error: 'Vapi ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 160)};
  return {ok: true};
}

/** Turns text into an MP3 in Shawna's ElevenLabs voice (ELEVENLABS_API_KEY, ELEVENLABS_VOICE_ID = the same voice as on
 *  her Vapi calls). "Julián" is written "Julian" so she says it the way he does. */
function shawnaVoice(text) {
  const p = PropertiesService.getScriptProperties();
  const res = UrlFetchApp.fetch('https://api.elevenlabs.io/v1/text-to-speech/' + p.getProperty('ELEVENLABS_VOICE_ID') + '?output_format=mp3_44100_128', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: {'xi-api-key': p.getProperty('ELEVENLABS_API_KEY')},
    payload: JSON.stringify({text: String(text).replace(/Julián/g, 'Julian'), model_id: p.getProperty('ELEVENLABS_MODEL') || 'eleven_multilingual_v2'})});
  if (res.getResponseCode() >= 300) throw new Error('ElevenLabs ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 160));
  return res.getBlob().setName('Shawna.mp3');
}

/** Run once after adding the ElevenLabs keys: texts Julián's personal phone a sample voice memo. */
function testVoiceMemo() {
  const p = PropertiesService.getScriptProperties(), to = bbE164(p.getProperty('ALERT_TEXT_TO') || p.getProperty('TEST_TEXT_TO'));
  const r = bbSend(to, 'Test voice memo from Shawna 👇🏾'); Utilities.sleep(1500);
  bbPostAttachment(to, r && r.service, shawnaVoice("Hi Julian, it's Shawna. If you can hear this, voice memos are working. Talk soon."));
  console.log('Sent. Check your personal phone.');
}

/** Finds a person in the contacts on Julián's phone (iCloud contacts synced to the texting Mac, via BlueBubbles). */
function macContact(who) {
  try {
    const want = String(who || '').toLowerCase().trim();
    if (!want) return null;
    const list = (bbApi('contact').data || []);
    const nm = c => (c.displayName || [c.firstName, c.lastName].filter(Boolean).join(' ')).trim();
    const hits = list.filter(c => nm(c).toLowerCase() === want).concat(list.filter(c => nm(c).toLowerCase().indexOf(want) >= 0));
    const c = hits[0];
    if (!c) return null;
    const phone = ((c.phoneNumbers || [])[0] || {}).address || '', email = ((c.emails || [])[0] || {}).address || '';
    return phone || email ? {name: nm(c), phone: phone, email: email} : null;
  } catch (e) { console.log('Mac contacts: ' + e.message); return null; }
}

/** Sends an email FROM julianc@bcolv.com through the small "Shawna Mailer" script in that account. */
function sendFromBcolv(to, subject, text) {
  const props = PropertiesService.getScriptProperties(), url = props.getProperty('BCOLV_MAILER_URL');
  if (!url) return {ok: false, error: 'bcolv mailer not set up'};
  const html = String(text).split(/\n{2,}/).map(par => '<p style="margin:0 0 12px">' + par.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>') + '</p>').join('') +
    '<p style="margin:0">Julián</p>' + EMAIL_FOOTER;
  const res = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({key: props.getProperty('BCOLV_MAILER_KEY'), to: to, subject: subject, text: text + '\n\nJulián', html: '<div style="font:15px/1.6 Helvetica,Arial,sans-serif;color:#111">' + html + '</div>'})});
  try { return JSON.parse(res.getContentText()); } catch (e) { return {ok: false, error: 'mailer answered ' + res.getResponseCode()}; }
}

/** Run once after setting BCOLV_MAILER_URL/KEY: emails Julián's personal Gmail from bcolv as a test. */
function testBcolvMailer() {
  console.log(JSON.stringify(sendFromBcolv(NOTIFY_EMAIL, 'Shawna mailer test', 'If you can read this, Shawna can send email from your bcolv address.')));
}

/** Run once after adding GROQ_API_KEY: asks Shawna a sample question and prints her answer (sends nothing). */
function testAssistant() {
  const props = PropertiesService.getScriptProperties();
  const out = groqChat(props.getProperty('GROQ_API_KEY'), props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: assistantPrompt(assistantName())},
    {role: 'user', content: 'Contact: Test Person\nStage: Lead\nTags: Text OK, open house\nRecent history: Came to an open house Sunday.\n\nThey just texted: "Hi! Is the house on Sunday still available? We might want to see a couple more this weekend."\n\nReply as JSON only.'}]);
  console.log('Model: ' + (props.getProperty('GROQ_MODEL') || ASSISTANT.model));
  console.log(assistantName() + ' says: ' + out);
}

/** Run once after adding BB_URL and BB_PASSWORD: checks the Mac and prints the webhook URL for BlueBubbles. */
function setupTexting() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BB_KEY')) props.setProperty('BB_KEY', Utilities.getUuid().replace(/-/g, ''));
  const info = bbApi('server/info');
  console.log('Connected to the texting Mac' + (info.data && info.data.detected_icloud ? ' (' + info.data.detected_icloud + ')' : '') + '.');
  console.log('In BlueBubbles > API & Webhooks, add a webhook with this URL (events: New Messages and New Server URL):');
  console.log(TEXTING.webApp + '?bb=' + props.getProperty('BB_KEY'));
}

/** Sends a test text to Julián's own number. */
function testText() {
  // Texts TEST_TEXT_TO (Script Property, e.g. another phone) if set; otherwise Julián's own number.
  const to = PropertiesService.getScriptProperties().getProperty('TEST_TEXT_TO') || TEXTING.myNumber;
  bbSend(to, 'Test from your website script: texting is connected. ✅');
  console.log('Test text sent to ' + to + '. Check that phone.');
}


// ============================================================
// TEXT QUEUE: scheduled texts (sent from Julián's number through the texting Mac). Each row in the "Text Queue" tab is
// one text; you can edit the wording in the sheet before it goes out. processTextQueue runs every 15 minutes, only
// sends 9am-9pm Pacific (federal TCPA limit for marketing texts is 8am-9pm), retries for 6 hours if the Mac is off, and skips a text if the person already texted back.
// ============================================================
const TEXT_QUEUE_TAB = {name: 'Text Queue', color: '#1C1C1E',
  headers: ['Due', 'Person ID', 'Name', 'Step', 'Status', 'Tries', 'Sent', 'Text']};

/** Open-house follow-up texts, Chris Voss style. {first} and {street} are filled in. */
const OPEN_HOUSE_VOICE = 'Quick hello';  // Google Drive voice note (the file name is what people see, so keep it casual)
const CONTACT_CARD = 'https://juliancrider.com/emails/julian-crider.vcf';  // HD contact card hosted on the site
// Each text can be several bubbles separated by " || ". [voice] = the voice note, [contact card] = the contact card.
// The first bubble has to be a text (Messages can't start a new conversation with a file on this Mac).
const OPEN_HOUSE_TEXTS = {
  // Buyers and lookers: the goal is a second showing. Later texts are skipped once they reply (Julián takes over).
  guest: [
    {step: 'oh-2h', hours: 2, text: "Hey {first}, it's Julián from the open house today 👋🏾 Thanks again for coming by. Just curious, was it what you expected or not quite? || [contact card] || juliancrider.com/?relocate || [voice]"},
    {step: 'oh-24h', hours: 24, text: "Hey {first}, Julián again. If that one wasn't quite the one, what would the next home need to have that it didn't? 🏡"},
    {step: 'oh-72h', hours: 72, text: "Would it be a bad idea if I lined up 2 or 3 homes for you to see this weekend? You can pick what you'd like to see here: juliancrider.com/?showing"}],
  neighbor: [
    {step: 'oh-2h', hours: 2, text: "Hey {first}, it's Julián from the open house today 👋🏾 Thanks for stopping by, neighbor. It's always good to meet the people who actually live here. || [contact card] || juliancrider.com/?relocate || [voice]"},
    {step: 'oh-72h', hours: 72, text: "Hey {first}, Julián again. A few neighbors asked what their own homes would sell for after the open house. Would it be crazy if I ran a quick, private number for yours? No visit needed. 🏡"}],
  // Signed up online through the funnel (no open house given): quick hello within minutes, no open-house voice note.
  online: [
    {step: 'oh-2h', hours: 0.05, text: "Hey {first}, it's Julián Crider with The DeBenedetto Team 👋🏾 Welcome, and thanks for signing up. As a developer and luxury advisor, I help people buy and sell homes across Las Vegas and Henderson, and you can trust me to make this an experience you actually enjoy. Just curious, what has you looking right now? || [contact card] || [welcome voice] || juliancrider.com/?relocate"},
    {step: 'oh-24h', hours: 24, text: "Hey {first}, Julián again. Whether it's an open house this weekend or a private tour, what would your next home need to have that your current one doesn't? 🏡"},
    {step: 'oh-72h', hours: 72, text: "Would it be a bad idea if I lined up 2 or 3 homes for you to see? You can pick what you'd like to see here: juliancrider.com/?showing"}],
  agent: [
    {step: 'oh-2h', hours: 2, text: "Hey {first}, Julián here 👋🏾 Thanks for previewing the home today. Any feedback from you or your buyers? It really helps the sellers."}]
};

function queueOpenHouseTexts(personId, d, kind) {
  const first = String(d.name || '').trim().split(/\s+/)[0] || 'there';
  const street = String(d.property || '').split(',')[0].trim() || 'the open house';
  const sheet = ensureTab(getSS(), TEXT_QUEUE_TAB);
  (OPEN_HOUSE_TEXTS[kind] || OPEN_HOUSE_TEXTS.guest).forEach(t => sheet.appendRow([
    new Date(Date.now() + t.hours * 3600e3), personId, String(d.name || ''), t.step, 'pending', 0, '',
    t.text.replace(/\{first\}/g, first).replace(/\{on_street\}/g, street === 'the open house' ? '' : ' on ' + street)
      .replace(/\{street\}/g, street)]));
  ensureTextQueueTrigger();
}

function ensureTextQueueTrigger() {
  const has = name => ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === name);
  if (!has('processTextQueue')) ScriptApp.newTrigger('processTextQueue').timeBased().everyMinutes(15).create();
  if (!has('remindUpcoming')) ScriptApp.newTrigger('remindUpcoming').timeBased().everyMinutes(5).create();
  if (!has('processAssistantQueue')) ScriptApp.newTrigger('processAssistantQueue').timeBased().everyMinutes(1).create();
}

// ============================================================
// APPOINTMENT REMINDERS: a text about 15 minutes before any event on Julián's Google Calendar that has a phone number
// in its notes (website bookings add "Phone: …" automatically; for a showing, add the client's number to the event).
// Title with "Showing" → showing wording (uses the event's location). Skips anyone tagged Text Opt-Out.
// ============================================================
function remindUpcoming() {
  try { processFubOutbox(); } catch (e) { console.log('Outbox: ' + e.message); }
  try { checkClientEmails(); } catch (e) { console.log('Email alerts: ' + e.message); }
  const now = Date.now(), props = PropertiesService.getScriptProperties();
  const events = CalendarApp.getDefaultCalendar().getEvents(new Date(now + 9 * 6e4), new Date(now + 21 * 6e4));
  events.forEach(ev => {
    const key = 'REM_' + ev.getId().slice(0, 40) + '_' + ev.getStartTime().getTime();
    if (props.getProperty(key) || ev.isAllDayEvent()) return;
    const blob = [ev.getDescription(), ev.getLocation(), ev.getTitle()].join('\n');
    const m = blob.match(/(?:\+?1[\s.-]?)?\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/);
    if (!m) return;
    const phone = m[1] + m[2] + m[3];
    if (bbE164(phone) === TEXTING.myNumber) return;
    props.setProperty(key, '1');  // one reminder per event, even if sending fails
    let person = null;
    try { person = (fubFetch('people?limit=1&fields=id,firstName,tags&phone=' + phone, null, 'get').people || [])[0] || null; } catch (e) {}
    if (person && (person.tags || []).indexOf('Text Opt-Out') >= 0) return;
    const titleName = (ev.getTitle().match(/:\s*([A-Za-z\u00C0-\u017F'-]+)/) || [])[1];
    const first = (person && person.firstName) || titleName || 'there';
    const time = Utilities.formatDate(ev.getStartTime(), 'America/Los_Angeles', 'h:mm a');
    const where = String(ev.getLocation() || '').split(',')[0].trim();
    const text = /showing|tour/i.test(ev.getTitle())
      ? `Hey ${first}, Julián here 👋🏾 Quick reminder, I'll see you ${where ? 'at ' + where + ' ' : ''}at ${time}. If anything changed, just text me. See you soon!`
      : `Hey ${first}, Julián here 👋🏾 Quick reminder about our ${time} today. See you soon!`;
    try {
      bbSend(phone, text);
      if (person) fubFetch('notes', {personId: person.id, subject: 'Reminder text sent', body: text});
    } catch (err) { console.log('Reminder not sent for "' + ev.getTitle() + '": ' + err.message); }
  });
}

function processTextQueue() {
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  if (hour < 9 || hour >= 21) return;  // quiet hours: wait for the morning
  const sheet = getSS().getSheetByName(TEXT_QUEUE_TAB.name);
  if (!sheet || sheet.getLastRow() < 2) return;
  const props = PropertiesService.getScriptProperties();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TEXT_QUEUE_TAB.headers.length).getValues();
  let sentNow = 0;
  rows.forEach((r, i) => {
    const [due, id, , step, status, tries, , text] = r;
    if (status !== 'pending' || !(due instanceof Date) || due > new Date() || sentNow >= 20) return;
    const cell = (col, v) => sheet.getRange(i + 2, col).setValue(v);
    // Later steps are skipped once they've texted back: Julián is in the conversation now.
    if (!/^oh-2h/.test(step) && props.getProperty('BB_TASK_' + id)) return cell(5, 'skipped: they replied');
    try {
      const res = bbTextPerson(id, text);
      cell(5, res); if (res === 'sent') { cell(7, new Date()); sentNow++; Utilities.sleep(3000); }
    } catch (err) {
      const n = (Number(tries) || 0) + 1;
      cell(6, n);
      if (n >= 24) cell(5, 'failed: ' + String(err.message).slice(0, 150));  // ~6 hours of retries
    }
  });
}

/** Test: sends the first open-house message (greeting, voice note, question, contact card) now to TEST_TEXT_TO. */
function testOpenHouseText() {
  const to = PropertiesService.getScriptProperties().getProperty('TEST_TEXT_TO') || TEXTING.myNumber;
  bbSendParts(to, OPEN_HOUSE_TEXTS.guest[0].text.replace(/\{first\}/g, 'there'));
  console.log('Sent to ' + to + '. Lines above say if the voice note or contact card failed, and why.');
}


// ============================================================
// OUTBOX: a Follow Up Boss note with the subject "Send text", written by Julián's account (user 2, e.g. Claude on his
// behalf), is texted to that contact through BlueBubbles within ~5 minutes, then logged as "Text sent".
// Bubbles can be split with " || ". Skips Text Opt-Out and notes older than 24 hours. Each note is sent once.
// ============================================================
// EMAIL OUTBOX: a note with the subject "Send email" by user 2 emails a FUB template to that contact from Julián's
// Gmail (with the signature footer), because FUB's API can't send email and action plans skip people who replied.
// Note body lines:  template: 187   (required)   ·   text: <a text to send right after, through the text outbox>
function processEmailOutbox() {
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('EMAIL_OUTBOX_DONE') || '[]');
  const notes = (fubFetch('notes?limit=25&sort=-created', null, 'get').notes || [])
    .filter(n => /^send email$/i.test(String(n.subject || '').trim()) && n.createdById === 2 && done.indexOf(n.id) < 0 &&
      Date.now() - new Date(n.created) < 36 * 36e5);
  notes.reverse().forEach(n => {
    done.push(n.id); props.setProperty('EMAIL_OUTBOX_DONE', JSON.stringify(done.slice(-200)));
    const body = String(n.body || ''), field = k => ((body.match(new RegExp('^\\s*' + k + '\\s*:\\s*(.+)$', 'im')) || [])[1] || '').trim();
    try {
      const t = fubFetch('templates/' + Number(field('template')), null, 'get');
      if (!t.body) throw new Error('template not found');
      const p = fubFetch('people/' + n.personId + '?fields=id,firstName,lastName,emails,tags', null, 'get');
      const to = (p.emails || []).filter(e => !/invalid|bounced/i.test(e.status || ''))[0];
      if (!to) throw new Error('no good email address');
      const fill = x => String(x).replace(/%contact_first_name%/g, p.firstName || '').replace(/%contact_name%/g, [p.firstName, p.lastName].join(' ').trim());
      const html = fill(t.body) + (t.body.indexOf('jc-footer') < 0 ? EMAIL_FOOTER : '');
      MailApp.sendEmail({to: to.value, subject: fill(t.subject), htmlBody: html, name: 'Julián Crider'});
      fubFetch('notes', {personId: p.id, subject: 'Email sent', body: 'To ' + to.value + ': "' + fill(t.subject) + '" (template ' + t.id + ', from Gmail)'});
      if (field('text')) fubFetch('notes', {personId: p.id, subject: 'Send text', body: field('text')});
    } catch (err) {
      fubFetch('notes', {personId: n.personId, subject: 'Email NOT sent', body: err.message + '\n\n' + body});
    }
  });
}

function processFubOutbox() {
  // Off unless Script Property EMAIL_OUTBOX = on: client emails must come from julianc@bcolv.com, not the script
  // owner's personal Gmail (Julián, Nov 8). Turn on only once it sends from bcolv.
  if (PropertiesService.getScriptProperties().getProperty('EMAIL_OUTBOX') === 'on') {
    try { processEmailOutbox(); } catch (e) { console.log('Email outbox: ' + e.message); }
  }
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  if (hour < 9 || hour >= 21) return;  // quiet hours: texts wait until 9am
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('OUTBOX_DONE') || '[]');
  const notes = (fubFetch('notes?limit=25&sort=-created', null, 'get').notes || [])
    .filter(n => /^send text$/i.test(String(n.subject || '').trim()) && n.createdById === 2 && done.indexOf(n.id) < 0 &&
      Date.now() - new Date(n.created) < 36 * 36e5);
  notes.reverse().forEach(n => {
    done.push(n.id); props.setProperty('OUTBOX_DONE', JSON.stringify(done.slice(-200)));  // mark first: never double-send
    try {
      const p = fubFetch('people/' + n.personId + '?fields=id,firstName,tags,phones', null, 'get');
      if ((p.tags || []).indexOf('Text Opt-Out') >= 0) throw new Error('they opted out of texts');
      const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
      if (!phone) throw new Error('no mobile number');
      bbSendParts(phone, String(n.body || '').trim());
      fubFetch('notes', {personId: p.id, subject: 'Text sent', body: n.body});
    } catch (err) {
      fubFetch('notes', {personId: n.personId, subject: 'Text NOT sent', body: err.message + '\n\n' + n.body});
    }
  });
}


// ============================================================
// OPEN HOUSE FUNNEL: answers given after the step-by-step sign-in (juliancrider.com/?openhouse) are added to the
// contact already created at step 1: FUB note + refreshed tier tags, the Sign-Ins sheet row, and if they turn out to
// have a signed agent, their scheduled texts are cancelled.
// ============================================================
function enrichSignIn(d) {
  const phone = String(d.phone || '').replace(/\D/g, ''), email = validEmail(d.email) ? String(d.email).trim() : '';
  const lines = [['Interest (1–10)', d.score], ['Currently', d.status], ['Timeline', d.timeline], ['How they heard', d.source], ['Details', d.notes]]
    .filter(x => x[1]).map(x => x[0] + ': ' + x[1]);
  // Sign-Ins sheet: fill in the row created at step 1 (latest row with this phone).
  try {
    const sh = getSS().getSheetByName('Sign-Ins');
    if (sh && sh.getLastRow() > 1 && phone) {
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][2]).replace(/\D/g, '').slice(-10) === phone.slice(-10)) {
          sh.getRange(i + 2, 6, 1, 5).setValues([[cell(d.score), cell(d.status), cell(d.timeline), cell(d.source),
            cell([vals[i][9], d.notes].filter(Boolean).join(' | '))]]);
          break;
        }
      }
    }
  } catch (err) { console.error('Enrich sheet: ' + err); }
  if (!PropertiesService.getScriptProperties().getProperty('FUB_API_KEY')) return {status: 'ok'};
  try {
    const found = (email && fubFetch('people?limit=1&fields=id,tags&email=' + encodeURIComponent(email), null, 'get').people) ||
                  (phone && fubFetch('people?limit=1&fields=id,tags&phone=' + phone, null, 'get').people) || [];
    if (!found.length) return {status: 'ok'};
    const person = found[0];
    if (lines.length) fubFetch('notes', {personId: person.id, subject: 'Open house details (funnel)', body: lines.join('\n')});
    // Re-score with the full answers; keep every other tag.
    const fresh = scoreLead('Sign-Ins', d);
    const keep = (person.tags || []).filter(t => !/^(Intake )?Tier [ABC]$|^(Hot|Warm)$/.test(t) || fresh.indexOf(t) >= 0);
    const tags = keep.concat(fresh.filter(t => keep.indexOf(t) < 0));
    if (tags.indexOf('Agent') >= 0) tags.splice(0, tags.length, ...tags.filter(t => !/^(Hot|Warm|Likely Seller)$/.test(t)));
    fubFetch('people/' + person.id, {tags: tags}, 'put');
    if (tags.indexOf('Has Agent') >= 0) {  // signed with another agent: no follow-up texts
      const sh = getSS().getSheetByName(TEXT_QUEUE_TAB.name);
      if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach((r, i) => {
        if (String(r[1]) === String(person.id) && r[4] === 'pending') sh.getRange(i + 2, 5).setValue('skipped: has agent');
      });
    }
  } catch (err) { console.error('Enrich FUB: ' + err); }
  return {status: 'ok'};
}
