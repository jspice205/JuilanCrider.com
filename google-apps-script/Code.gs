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
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'signAgreement') return signAgreement(data);
    if (data.action === 'book') return json(bookSlot(data));
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
      if (agreementSettings().live && !/Working with agent: Yes, with a signed/.test(data.notes || '')) {
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
  CalendarApp.getDefaultCalendar().createEvent(`Consultation: ${d.name} (${d.type || 'juliancrider.com'})`, slot, end, {
    description: [`Booked on juliancrider.com`, `Phone: ${d.phone}`, email && `Email: ${email}`, d.notes].filter(Boolean).join('\n'),
    guests: email, sendInvites: !!email});
  try {  // log it on their Follow Up Boss contact (the lead itself arrives through the normal website form)
    const phone = String(d.phone).replace(/\D/g, '');
    const found = (email && fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people) ||
                  (phone && fubFetch('people?limit=1&phone=' + phone, null, 'get').people) || [];
    if (found.length) {
      fubFetch('notes', {personId: found[0].id, subject: 'Booked a consultation', body: `${label} (${BOOKING.minutes} min), booked on juliancrider.com`});
      createCallTask(found[0].id, `booked consultation ${label}`);
    }
  } catch (err) { console.error('Booking note: ' + err); }
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

function autoUpdateFromGitHub() { console.log(updateFromGitHub(true)); }

function updateFromGitHub(auto) {
  const raw = f => UrlFetchApp.fetch(GITHUB_RAW + f + '?nocache=' + Date.now()).getContentText();
  const code = raw('Code.gs'), manifest = raw('appsscript.json');
  const id = ScriptApp.getScriptId(), base = 'https://script.googleapis.com/v1/projects/' + id;
  const api = (method, path, body) => JSON.parse(UrlFetchApp.fetch(base + path, {method: method, contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, payload: body ? JSON.stringify(body) : undefined}).getContentText() || '{}');
  const files = api('get', '/content').files || [];
  const cur = name => (files.filter(f => f.name === name)[0] || {}).source;
  if (cur('Code') === undefined) return 'No file named "Code" in this project, so nothing was changed.';
  if (cur('Code') === code && cur('appsscript') === manifest) return 'Already up to date.';
  const scopes = s => JSON.stringify(((s && JSON.parse(s).oauthScopes) || []).slice().sort());
  if (auto && scopes(cur('appsscript')) !== scopes(manifest)) {
    const props = PropertiesService.getScriptProperties();
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
  (api('get', '/deployments').deployments || []).filter(dep => (dep.entryPoints || []).some(e => e.entryPointType === 'WEB_APP'))
    .forEach(dep => api('put', '/deployments/' + dep.deploymentId,
      {deploymentConfig: {scriptId: id, versionNumber: version, manifestFileName: 'appsscript', description: 'Update from GitHub'}}));
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

function fubSource(tabName) { return tabName === 'Sign-Ins' ? 'Open House' : 'juliancrider.com'; }

/** A "Call" task due today (Pacific) in Follow Up Boss, assigned to the account owner. */
function createCallTask(personId, reason, type) {
  type = type || 'Call';  // 'Email' for contacts whose phone isn't DNC-checked
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
  return fubFetch('tasks', {personId: personId, type: type, name: (type + ': ' + reason).slice(0, 250), dueDate: today});
}

function fubPerson(tabName, d) {
  const parts = String(d.name || '').trim().split(/\s+/);
  const tags = tabName === 'Sign-Ins' ? ['Open House'] : ['Website Lead', 'juliancrider.com'];
  if (tabName === 'Sign-Ins' && /juliancrider\.com/.test(d.notes || '')) tags.push('juliancrider.com');
  if (d.type) tags.push(d.type.replace(/^Consultation:\s*/, ''));
  if (tabName !== 'Sign-Ins') tags.push('Text OK');  // website forms carry the "agree to be contacted by text" line
  scoreLead(tabName, d).forEach(t => tags.push(t));
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
  let pts = ({'ASAP / Already looking': 30, 'Ready now': 30, '1–3 months': 25, '3–6 months': 15, '6–12 months': 8}[tl] || 0);
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
    source: fubSource(tabName),
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
    return id;
  };
  // Welcome email (FUB Action Plan "JC Welcome", template 183) right after an open-house sign-in or a website
  // account/guide sign-up. Plans started through the API begin Paused, so switch it to Running.
  if ((tabName === 'Sign-Ins' || /account|guide/i.test(d.type || '')) && (person.emails || []).length) {
    try {
      const id = personId();
      if (id) {
        const ap = fubFetch('actionPlansPeople', {personId: id, actionPlanId: WELCOME_PLAN});
        if (ap.status !== 'Running') fubFetch('actionPlansPeople/' + ap.id, {status: 'Running'}, 'put');
      }
    } catch (err) { console.error('Welcome email: ' + err); }
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
  tags: {hot: 'JC Hot', warm: 'JC Warm', legacy: 'JC Legacy Re-Intro', farm: 'JC Farm Best', nurture: 'JC Long-Term Nurture'},
  plans: {hot: 22, warm: 23, legacy: 24, farm: 25, nurture: 26},  // Follow Up Boss Action Plan ids
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
  const planIds = Object.keys(FUNNEL.plans).map(k => FUNNEL.plans[k]);
  // Pause any JC plan this person is running (moving stages, or a deal/Pause tag).
  const stopPlans = p => (fubFetch('actionPlansPeople?personId=' + p.id, null, 'get').actionPlansPeople || [])
    .filter(a => planIds.indexOf(a.actionPlanId) >= 0 && a.status === 'Running')
    .forEach(a => fubFetch('actionPlansPeople/' + a.id, {status: 'Paused'}, 'put'));
  const start = (p, stage, why) => {
    if (used >= FUNNEL.perNight || FUNNEL.live.indexOf(stage) < 0) return;
    const tag = FUNNEL.tags[stage] + (has(p, 'Text OK') && FUNNEL.textLive.indexOf(stage) >= 0 ? ' + Text' : '');
    try {
      stopPlans(p);
      const ap = fubFetch('actionPlansPeople', {personId: p.id, actionPlanId: FUNNEL.plans[stage]});
      if (ap.status !== 'Running') fubFetch('actionPlansPeople/' + ap.id, {status: 'Running'}, 'put');
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
  // In priority order, until tonight's limit is used up.
  // Hot/Warm tags stay on for good, so only a tag added tonight moves someone up from Warm or Nurture.
  eligible.filter(p => has(p, 'Hot') && (!stageOf(p) || (p.newHot && /warm|nurture/.test(stageOf(p))))).forEach(p => start(p, 'hot', 'Tier A'));
  eligible.filter(p => has(p, 'Warm') && !has(p, 'Hot') && (!stageOf(p) || (p.newWarm && stageOf(p) === 'nurture'))).forEach(p => start(p, 'warm', 'Tier B'));
  eligible.filter(p => FUNNEL.handoffDays[stageOf(p)] && daysAgo(rows[p.id].started) >= FUNNEL.handoffDays[stageOf(p)])
    .forEach(p => start(p, 'nurture', 'finished ' + stageOf(p)));
  eligible.filter(p => stageOf(p) === 'paused').forEach(p => start(p, 'nurture', 'resumed after pause'));
  eligible.filter(p => !stageOf(p) && (p.tags || []).some(t => /^Legacy/.test(t))).forEach(p => start(p, 'legacy', 'Legacy contact'));
  eligible.filter(p => !stageOf(p) && has(p, 'Farm – Best')).forEach(p => start(p, 'farm', 'Farm – Best'));
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

/** Run once: re-scores everyone in Follow Up Boss every night around 2am Pacific and adds Hot/Warm tags,
 *  so your FUB automations start the right action plan while you sleep. Check the "FUB Scorecard" tab each morning. */
function setupNightlyScoring() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'scoreFubDatabase').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scoreFubDatabase').timeBased().everyDays(1).atHour(2).inTimezone('America/Los_Angeles').create();
  console.log('Nightly scoring scheduled for around 2am Pacific.');
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
    ? `<p>Thanks for your interest in <b>${esc(d.property)}</b>. Since you have a signed agreement with another agent, they should schedule this tour for you, and they'll be able to get you in quickly.</p>
       <p>If anything changes, or you have questions about the home or the neighborhood, I'm always happy to help.</p>`
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
