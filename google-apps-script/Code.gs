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
    }
    // ── Showing requests: send the buyer a confirmation right away ──
    if (tabName === 'Website Leads' && /showing/i.test(data.type || '') && validEmail(data.email)) {
      try { confirmShowing(data); } catch (err) { console.error('Buyer confirmation failed', err); }
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

function fubPerson(tabName, d) {
  const parts = String(d.name || '').trim().split(/\s+/);
  const tags = tabName === 'Sign-Ins' ? ['Open House'] : ['Website Lead', 'juliancrider.com'];
  if (tabName === 'Sign-Ins' && /juliancrider\.com/.test(d.notes || '')) tags.push('juliancrider.com');
  if (d.type) tags.push(d.type.replace(/^Consultation:\s*/, ''));
  scoreLead(tabName, d).forEach(t => tags.push(t));
  const person = {firstName: parts.shift() || '', lastName: parts.join(' '), tags: tags};
  if (validEmail(d.email)) person.emails = [{value: String(d.email).trim()}];
  if (d.phone) person.phones = [{value: d.phone}];
  return person;
}

// Scores a new lead from what the form collected and returns tier tags
// ("Tier A/B/C", "Likely Seller", "Has Agent"), so FUB Smart Lists and automations can act on them.
function scoreLead(tabName, d) {
  const notes = String(d.notes || ''), type = String(d.type || '').toLowerCase(), tl = String(d.timeline || '');
  if (/Working with agent: Yes, with a signed/.test(notes)) return ['Tier C', 'Has Agent'];
  const tags = [];
  let pts = ({'ASAP / Already looking': 30, 'Ready now': 30, '1–3 months': 25, '3–6 months': 15, '6–12 months': 8}[tl] || 0);
  const interest = Number(d.score) || 0;
  pts += interest >= 8 ? 20 : interest >= 6 ? 10 : interest ? 3 : 0;
  if (/Pre-approved: Yes|Pre-approved: Paying cash|Paying cash/.test(notes)) pts += 10;
  if (/showing/.test(type)) { tags.push('Showing requested'); pts += 45; }
  else if (/sell|home value/.test(type)) { tags.push('Likely Seller'); pts += 20; }
  else if (/new construction|buying|relocat/.test(type)) pts += 10;
  if (/^Owning/.test(String(d.status || '')) && pts >= 15) tags.push('Likely Seller');
  tags.unshift(pts >= 45 ? 'Tier A' : pts >= 25 ? 'Tier B' : 'Tier C');
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
  fubFetch('events', event);
  console.log('Sent to Follow Up Boss');
}

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
  for (let r = Number(props.getProperty('FUB_TAGGED_ROW') || 1); r < values.length; r++) {
    if (Date.now() - started > 5 * 60 * 1000) { console.log(`Paused at row ${r + 1}. Run tagScoredLeadsInFub again to continue.`); break; }
    const get = name => String(values[r][head.indexOf(name)] || '').trim();
    const tier = get('Tier'), flags = get('Flags');
    const want = [tier === 'Partner' ? 'Partner' : 'Tier ' + tier];
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
        const person = found[0], tags = (person.tags || []).slice();
        want.forEach(t => { if (tags.indexOf(t) < 0) tags.push(t); });
        fubFetch('people/' + person.id, {tags: tags}, 'put');
        tagged++;
      }
    } catch (err) { failed++; console.error(`${get('Name')}: ${err}`); }
    props.setProperty('FUB_TAGGED_ROW', String(r + 1));
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
// BUYER CONFIRMATION — branded email to the buyer after a showing request
// ============================================================
function confirmShowing(d) {
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
         <li style="margin-bottom:6px">You get a short buyer agreement to e-sign. It's <b>non-exclusive and covers only this one home</b>, so you're free to work with anyone else. Nevada requires a written agreement before any tour, and it takes about a minute.</li>
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
