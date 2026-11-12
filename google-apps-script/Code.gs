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
    if (e.parameter && e.parameter.ig) return instagramWebhook(e);  // ManyChat (Instagram DM automation) leads
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
      // Julián, Nov 12: send the one-home agreement + Duties Owed right away, even without an address. Without one, it's a
      // review copy (the home is filled in, and it can be signed, once they pick one with Julián).
      if (agreementSettings().live && !/Working with agent: Yes, with a signed/.test(data.notes || '')) {
        try { signUrl = createAgreement(data.property ? data : Object.assign({}, data, {property: AGREEMENT_TBD})); } catch (err) { console.error('Agreement not created', err); }
      }
      try { confirmShowing(data, signUrl); } catch (err) { console.error('Buyer confirmation failed', err); }
    }
    // ── Shawna texts Julián about every website lead right away (Nov 12) ──
    if (tabName === 'Website Leads' || /juliancrider\.com/.test(data.notes || '')) {
      try {
        const bits = [data.phone, data.email, data.property && 'Home: ' + data.property, data.timeline && 'When: ' + data.timeline,
          String(data.notes || '').split(' | ').filter(x => /Requested time|Preferred time|Pre-approved|Working with agent|Price range|Lives in/.test(x) && !/: *$/.test(x)).join(', ')].filter(Boolean);
        notifyJulian('🌐 ' + (data.type || tabName) + ': ' + (data.name || 'someone'), 'from juliancrider.com. ' + bits.join(' · '), '', data.name || '');
      } catch (err) { console.log('Lead alert: ' + err.message); }
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

const CODE_VERSION = '2026-11-12-teamrecap';  // bump when a change must be confirmed live (GET ?action=version)
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'version') return json({status: 'ok', version: CODE_VERSION});
  if (e && e.parameter && e.parameter.action === 'rates') return rates();
  if (e && e.parameter && e.parameter.action === 'agreement') return getAgreement(String(e.parameter.id || ''));
  if (e && e.parameter && e.parameter.action === 'slots') return json({status: 'ok', slots: bookingSlots(meetingFor(e.parameter)).map(d => d.toISOString())});
  return json({status: 'ok', message: 'JC Leads API is running'});
}

// ============================================================
// BOOKING — "Book a consultation" on the site shows open times from your Google Calendar and books them.
// Anything timed on your main calendar (Tao shifts, showings, appointments) blocks that time. Change hours here.
// ============================================================
const BOOKING = {days: 10, minutes: 30, startHour: 9, endHour: 18, weekdays: [1, 2, 3, 4, 5, 6], noticeHours: 3, tz: 'America/Los_Angeles'};

// Meeting types (Julián's choices, Nov 2): phone and video Mon–Sat 9–6, 30 min (video can be 45); golf by private link only,
// Mon/Tue/Fri tee times 8am–3pm and Thu 11am–4pm, 9 holes ≈ 2.5 h or 18 holes ≈ 5 h. No coffee, lunch or drinks.
const MEETINGS = {
  consult: {label: 'Consultation', minutes: 30, hours: {1: [9, 18], 2: [9, 18], 3: [9, 18], 4: [9, 18], 5: [9, 18], 6: [9, 18]}},
  phone: {label: 'Phone call', minutes: 30, hours: {1: [9, 18], 2: [9, 18], 3: [9, 18], 4: [9, 18], 5: [9, 18], 6: [9, 18]}},
  video: {label: 'Video call', minutes: 30, max: 45, hours: {1: [9, 18], 2: [9, 18], 3: [9, 18], 4: [9, 18], 5: [9, 18], 6: [9, 18]}},
  showing: {label: 'Showing', minutes: 30, hours: {1: [9, 18], 2: [9, 18], 3: [9, 18], 4: [9, 18], 5: [9, 18], 6: [9, 18]}},
  golf: {label: 'Golf', minutes: 300, startsOnly: true, step: 60, days: 21, hours: {1: [8, 15], 2: [8, 15], 4: [11, 16], 5: [8, 15]}},
};
function meetingFor(q) {
  const kind = MEETINGS[q.kind] ? q.kind : 'consult', m = Object.assign({kind: kind}, MEETINGS[kind]);
  if (kind === 'video' && Number(q.minutes) === 45) m.minutes = 45;
  if (kind === 'golf') m.minutes = String(q.holes) === '9' ? 150 : 300;
  return m;
}

function bookingSlots(m) {
  m = m || meetingFor({});
  const now = new Date(), days = m.days || BOOKING.days, until = new Date(now.getTime() + (days + 1) * 864e5);
  const busy = CalendarApp.getDefaultCalendar().getEvents(now, until).filter(ev => !ev.isAllDayEvent())
    .map(ev => [ev.getStartTime().getTime(), ev.getEndTime().getTime()]);
  try {  // the bcolv (business) calendar counts as busy too
    (bcolvCall({action: 'events', from: now.toISOString(), to: until.toISOString()}).events || [])
      .forEach(e => busy.push([e.start, e.end || e.start + 30 * 6e4]));
  } catch (e) {}
  const earliest = now.getTime() + BOOKING.noticeHours * 36e5, len = m.minutes * 6e4, step = (m.step || 30), slots = [];
  for (let d = 0; d < days; d++) {
    const day = new Date(now.getTime() + d * 864e5);
    const hrs = m.hours[Number(Utilities.formatDate(day, BOOKING.tz, 'u')) % 7];  // u: 1 = Monday … 7 = Sunday
    if (!hrs) continue;
    const ymd = Utilities.formatDate(day, BOOKING.tz, 'yyyy-MM-dd');
    // Golf windows are tee-time windows (the round can run past them); everything else must fit inside the hours.
    const lastStart = m.startsOnly ? hrs[1] * 60 : hrs[1] * 60 - m.minutes;
    for (let mm = hrs[0] * 60; mm <= lastStart; mm += step) {
      const start = Utilities.parseDate(`${ymd} ${Math.floor(mm / 60)}:${mm % 60}`, BOOKING.tz, 'yyyy-MM-dd H:m'), t = start.getTime();
      if (t >= earliest && !busy.some(b => b[0] < t + len && b[1] > t)) slots.push(start);
    }
  }
  return slots;
}

/** Creates the calendar event; video calls get a Google Meet link (Calendar API), with VIDEO_LINK as a fallback. */
function createBookingEvent(title, start, end, description, location, email, video) {
  if (video) {
    try {
      const res = UrlFetchApp.fetch('https://www.googleapis.com/calendar/v3/calendars/primary/events?conferenceDataVersion=1&sendUpdates=' + (email ? 'all' : 'none'), {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
        payload: JSON.stringify({summary: title, description: description, start: {dateTime: start.toISOString()}, end: {dateTime: end.toISOString()},
          attendees: email ? [{email: email}] : [], conferenceData: {createRequest: {requestId: Utilities.getUuid(), conferenceSolutionKey: {type: 'hangoutsMeet'}}}})});
      const ev = JSON.parse(res.getContentText());
      if (res.getResponseCode() < 300) return ev.hangoutLink || ((ev.conferenceData || {}).entryPoints || [{}])[0].uri || '';
      console.log('Meet link: ' + res.getContentText().slice(0, 200));
    } catch (e) { console.log('Meet link: ' + e.message); }
  }
  const link = video ? (PropertiesService.getScriptProperties().getProperty('VIDEO_LINK') || '') : '';
  CalendarApp.getDefaultCalendar().createEvent(title, start, end, {description: description + (link ? '\nVideo: ' + link : ''),
    location: location || link || '', guests: email, sendInvites: !!email});
  return link;
}

function bookSlot(d) {
  const m = meetingFor(d);
  const slot = bookingSlots(m).filter(s => s.toISOString() === d.slot)[0];
  if (!slot) return {status: 'taken', message: 'That time was just taken. Please pick another.'};
  if (!String(d.name || '').trim() || !String(d.phone || '').trim()) return {status: 'error', message: 'Name and phone are required.'};
  const end = new Date(slot.getTime() + m.minutes * 6e4), email = validEmail(d.email) ? String(d.email).trim() : '';
  const label = Utilities.formatDate(slot, BOOKING.tz, "EEE, MMM d 'at' h:mm a");
  const showing = m.kind === 'showing' && d.property, golf = m.kind === 'golf', video = m.kind === 'video';
  const title = showing ? `Showing: ${d.name}` : golf ? `Golf: ${d.name} (${d.holes || 18} holes)` :
    `${m.label}: ${d.name}${d.type ? ' (' + d.type + ')' : ''}`;
  const desc = ['Booked on juliancrider.com', `Phone: ${d.phone}`, email && `Email: ${email}`,
    golf && `Holes: ${d.holes || 18}`, golf && `Tee time booked: ${d.teeBooked || 'not yet'}`, golf && d.course && `Course: ${d.course}`, d.notes].filter(Boolean).join('\n');
  const link = createBookingEvent(title, slot, end, desc, showing ? String(d.property) : golf ? (d.course || '') : '', email, video);
  try {  // Follow Up Boss: find or add the person, then appointment + note + task
    const phone = String(d.phone).replace(/\D/g, '');
    let found = (email && fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people) ||
                (phone && fubFetch('people?limit=1&phone=' + phone, null, 'get').people) || [];
    if (!found.length) {
      const parts = String(d.name).trim().split(/\s+/);
      fubFetch('events', {source: 'juliancrider.com', system: 'juliancrider.com', type: 'General Inquiry', message: title + ' ' + label,
        person: {firstName: parts[0], lastName: parts.slice(1).join(' '), phones: [{value: phone}], emails: email ? [{value: email}] : [], tags: ['Website Lead', golf ? 'Golf' : 'Booked']}});
      found = (phone && fubFetch('people?limit=1&phone=' + phone, null, 'get').people) || [];
    }
    if (found.length) {
      const pid = found[0].id;
      try { fubFetch('appointments', {title: title, description: desc, start: slot.toISOString(), end: end.toISOString(),
        location: showing ? String(d.property) : link || (golf ? d.course || '' : ''), invitees: [{personId: pid}]}); } catch (e) { console.log('FUB appointment: ' + e.message); }
      fubFetch('notes', {personId: pid, subject: 'Booked: ' + m.label, body: `${label} (${m.minutes} min)${showing ? ' at ' + d.property : ''}${link ? '\nVideo: ' + link : ''}\n${desc}`});
      if (!golf) createCallTask(pid, showing ? `showing ${label} at ${d.property}` : `${m.label.toLowerCase()} ${label}`);
    }
  } catch (err) { console.error('Booking FUB: ' + err); }
  try {  // confirmation text (the 15-minute reminder comes from remindUpcoming)
    const first = String(d.name).trim().split(/\s+/)[0];
    const sent = bbSend(d.phone, showing ? `Hey ${first}, it's Julián 👋🏾 You're all set to see ${String(d.property).split(',')[0]} ${label}. I'll meet you there. If anything changes, just text me here.`
      : golf ? `Hey ${first}, it's Julián ⛳ We're on for ${d.holes || 18} holes ${label}${d.course ? ' at ' + d.course : ''}. Looking forward to it!`
      : video ? `Hey ${first}, it's Julián 👋🏾 You're all set for our video call ${label}.${link ? " Here's the link: " + link : " I'll send the link shortly."}`
      : `Hey ${first}, it's Julián 👋🏾 You're all set for ${label}. I'll call you then. If anything changes, just text me here.`);
    try {  // tap-to-add calendar invite right under the confirmation
      Utilities.sleep(1500);
      const invTitle = showing ? 'Showing with Julián: ' + String(d.property).split(',')[0] : golf ? 'Golf with Julián' : video ? 'Video call with Julián Crider' : 'Call with Julián Crider';
      bbPostAttachment(d.phone, sent && sent.service, icsInvite(invTitle, slot, end, showing ? String(d.property) : golf ? (d.course || '') : (link || 'Julián will call you'),
        (link ? 'Join: ' + link + '\n' : '') + 'Questions or changes: text Julián at 702-498-2571.'));
    } catch (e) { console.log('Calendar invite not sent: ' + e.message); }
  } catch (err) { console.log('Booking confirmation text not sent: ' + err.message); }
  try { notifyJulian('📅 Booked on your site: ' + String(d.name || '').trim(), m.label + ', ' + label + (d.property ? ' at ' + d.property : '') + '. ' + (d.phone || '') + '. It\'s on your calendar.', '', String(d.name || '').trim()); } catch (e) {}
  return {status: 'success', label: label, link: link};
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
  try { cleanupProperties(); } catch (e) { console.log('Property cleanup: ' + e.message); }
  try { processAllowNotes(); } catch (e) { console.log('Shawna allow: ' + e.message); }
  try { cleanupOldTasks(); } catch (e) { console.log('Old tasks: ' + e.message); }
  try { syncVapiAssistant(); } catch (e) { console.log('Vapi sync: ' + e.message); }
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
    if (e.updated > seen && e.type === 'bounce') {  // tag bounced addresses so the funnel stops emailing them
      const p = person(e.personId);
      if (p && (p.tags || []).indexOf('Bounced') < 0) { try { fubFetch('people/' + p.id, {tags: (p.tags || []).concat('Bounced')}, 'put'); p.tags = (p.tags || []).concat('Bounced'); } catch (err) {} }
    }
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
  const me = Number(PropertiesService.getScriptProperties().getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
  return fubFetch('tasks', {personId: personId, assignedUserId: me, type: type, name: (type + ': ' + reason).slice(0, 250), dueDate: today});
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
    if (!id && d.name) {  // name only: the event created them, so pick the newest person with that exact name
      const same = (fubFetch('people?limit=5&sort=-created&fields=id,name&q=' + encodeURIComponent(d.name), null, 'get').people || [])
        .filter(x => (x.name || '').toLowerCase() === String(d.name).toLowerCase());
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
  // Website forms: a welcome text from Julián's number (the forms carry the text-consent line). Bookings and showing
  // requests skip it because they get a booking confirmation text instead.
  if (tabName !== 'Sign-Ins' && d.phone && person.tags.indexOf('Has Agent') < 0 && !/consultation|showing|book/i.test(String(d.type || ''))) {
    try { const id = personId(); if (id) queueOpenHouseTexts(id, d, 'website'); }
    catch (err) { console.error('Website welcome text: ' + err); }
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
    try { findDuplicates(people); } catch (err) { console.error('Duplicates: ' + err); }
    try { runPastClientLoop(people); } catch (err) { console.error('Closed client loop: ' + err); }
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

  const paused = p => /active client|pending|under contract|closed/i.test(String(p.stage || '')) || has(p, 'Pause') || has(p, 'Not Interested');
  const isBounced = p => has(p, 'Bounced') && !has(p, 'Opened Email');
  const ok = p => {
    const tags = p.tags || [], stage = String(p.stage || '');
    if (!(p.emails || []).some(e => validEmail(e.value) && !/invalid|bounce|bad/i.test(e.status || ''))) return false;
    if (isBounced(p)) return false;
    if (/trash|unresponsive|do not|past client|sphere/i.test(stage) || paused(p)) return false;
    if (tags.some(t => /^past client$/i.test(t))) return false;  // the closed-client loop handles them
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
  // Bounced emails hurt deliverability for everyone: stop their email drips once (they keep their phone, notes and tags).
  const cleaned = JSON.parse(props.getProperty('BOUNCE_CLEANED') || '[]');
  people.filter(isBounced).filter(p => cleaned.indexOf(p.id) < 0).slice(0, 60).forEach(p => {
    try { stopPlans(p); cleaned.push(p.id); Utilities.sleep(100); } catch (e) { console.log('Bounce cleanup ' + p.id + ': ' + e.message); }
  });
  props.setProperty('BOUNCE_CLEANED', JSON.stringify(cleaned.slice(-1500)));
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
  // They said "not right now": off the Hot/Warm/Legacy/Farm drip and onto the slower Long-Term Nurture (Nov 11).
  eligible.filter(p => has(p, 'Not Now') && stageOf(p) && !/^nurture$|^paused$/.test(stageOf(p))).forEach(p => start(p, 'nurture', 'said not right now'));
  eligible.filter(p => has(p, 'Hot') && !has(p, 'Not Now') && (!stageOf(p) || (p.newHot && /warm|nurture/.test(stageOf(p))))).forEach(p => start(p, 'hot', 'Tier A'));
  eligible.filter(p => has(p, 'Warm') && !has(p, 'Hot') && !has(p, 'Not Now') && (!stageOf(p) || (p.newWarm && stageOf(p) === 'nurture'))).forEach(p => start(p, 'warm', 'Tier B'));
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

// ============================================================
// CLOSED CLIENT LOOP (Nov 10). Where referrals and repeat business come from. Runs every morning right after the
// funnel. You only move the deal card in Follow Up Boss; everything else follows on its own:
//   Deal stage → contact stage: Hot Buyers / Listing Prep / Active Listings → Active Client; Offer, Due Diligence,
//   Loan/Appraisal → Pending (both pause the email drips); CLOSED → Closed + this loop. Never moves anyone backwards.
//   Signed buyer agreement on the site → a deal card in Buyers › HOT BUYERS for you.
// The loop (from the closing date; texts only to Text OK, emails from julianc@bcolv.com via the Shawna Mailer):
//   day 3 text (or email): how are the first nights · day 7 email: first week + vendor list · day 30: review ask
//   (waits until the Script Property REVIEW_LINK is set) · day 60: referral ask · every 90 days: home value check-in
//   (4 rotating notes) · every closing anniversary: happy home anniversary + a soft referral P.S.
// After day 30 the contact moves from Closed to Past Client. Contacts already tagged "past client" (your older
// clients) join at the quarterly check-ins, the first one 30-60 days out, and their old drip is paused.
// Skipped: Trash, Deceased, Has Agent, tags "Pause" / "No Past Client Loop". The "Past Clients" tab shows everyone,
// their closing date and every touch sent. Max PAST_CLIENTS.perDay touches a day.
// ============================================================
const PAST_CLIENTS = {
  live: true, perDay: 15,
  dealStages: {active: [19, 20, 22], pending: [21, 28, 23, 36, 24, 29], closed: [25, 26]},
  agreementStage: 19,  // Buyers › HOT BUYERS
  steps: [
    {step: 'pc-3d', day: 3, text: "Hey {first}, Julián here 🔑 How are the first few nights in the new place? Anything not working the way it should?",
      subject: 'How are the first few nights?', email: "Hi {first},\n\nHow are the first few nights in the new place? If anything isn't working the way it should, tell me and I'll get the right person out there."},
    {step: 'pc-1w', day: 7, subject: 'Your first week',
      email: "Hi {first},\n\nOne week in. How's it feeling?\n\nA few things people usually need in the first month: a handyman, HVAC service, a pool or landscape company, someone for blinds. I keep a short list of people I'd send my own family to. Just reply with what you need and I'll send names.\n\nAnd if anything about the house surprised you, I want to hear it."},
    {step: 'pc-30d', day: 30, needsReview: true, subject: 'Would it be a bad idea to ask a favor?',
      text: "Hey {first}, Julián here. Would it be a bad idea to ask you a small favor? If I did a good job for you, a quick review would mean a lot: {review} 🙏🏾",
      email: "Hi {first},\n\nWould it be a bad idea to ask you a small favor?\n\nMost of my business comes from people like you telling someone else. If I did a good job for you, a short review would mean a lot, even two sentences:\n\n{review}\n\nAnd if I didn't, I'd rather hear that from you directly."},
    {step: 'pc-60d', day: 60, subject: "Who's the next person?",
      text: "Hey {first}, Julián here 👋🏾 Quick one. Who's the one person you know who keeps talking about moving? Would you be against passing my number along? I'll take good care of them.",
      email: "Hi {first},\n\nQuick one. Who's the one person you know who keeps talking about moving, either here to Vegas or within the valley?\n\nWould you be against passing my name along? I'll take care of them the way I took care of you, and I'll keep you posted."}],
  quarterly: [
    {subject: 'What your home is worth right now', text: "Hey {first}, Julián here 🏡 Would it be crazy if I ran a fresh number on what your home is worth today? Takes me 10 minutes, no visit needed.",
      email: "Hi {first},\n\nIt's been a few months, so I wanted to check in. Would it be crazy if I ran a fresh number on what your home is worth today? It takes me about 10 minutes, no visit needed, and it's good to know for insurance, refinancing or just peace of mind.\n\nJust reply \"yes\" or use juliancrider.com/?value"},
    {subject: 'Season check for the house', text: "Hey {first}, Julián here. No real estate talk, just checking in. How's the house treating you? ☀️",
      email: "Hi {first},\n\nNo real estate talk today, just checking in. How's the house treating you?\n\nIf you need a good HVAC, pool or roof person before the season turns, reply and I'll send who I trust."},
    {subject: "What's selling near you", text: "Hey {first}, Julián here. Homes near you have been moving. Would you be against me sending a quick snapshot of what sold nearby?",
      email: "Hi {first},\n\nHomes near you have been moving. Would you be against me sending you a quick snapshot of what sold nearby and for how much? Just reply and I'll put it together."},
    {subject: 'Quick favor', text: "Hey {first}, hope all is good 👋🏾 If anyone you know is thinking about a move, would it be a bad idea to pass along my number? I'll take great care of them.",
      email: "Hi {first},\n\nHope all is good with you. If anyone you know is thinking about buying or selling, would it be a bad idea to pass along my name? I'll take great care of them, the same way I tried to take care of you."}],
  anniversary: {subject: 'Happy home anniversary 🏡',
    text: "Happy {years} home anniversary, {first}! 🎉🏡 Hard to believe it's been {yearsWord} already. Hope the house has been good to you. Julián",
    email: "Hi {first},\n\nHappy {years} home anniversary! Hard to believe it's been {yearsWord} since you got the keys.\n\nI hope the house has been good to you. If you'd ever like to know what it's worth today, just reply and I'll run the number.\n\nP.S. If someone you know is thinking about a move, I'd be grateful for the introduction."}
};
const PAST_CLIENTS_TAB = {name: 'Past Clients', color: '#C9A227', headers: ['Person ID', 'Name', 'Closed on', 'Joined as', 'Touches sent', 'Last touch', 'Status']};

function pcFill(t, p, extra) {
  const first = p.firstName || String(p.name || '').split(/\s+/)[0] || 'there';
  let s = String(t || '').replace(/\{first\}/g, first);
  Object.keys(extra || {}).forEach(k => s = s.split('{' + k + '}').join(extra[k]));
  return s;
}

/** Moves contact stages to follow their deal cards, and returns {personId: closeDate} for deals in CLOSED. */
function syncDealStages(peopleById) {
  const order = ['active client', 'pending', 'closed', 'past client'], closed = {};
  const deals = [];
  for (let offset = 0; offset < 2000; offset += 100) {
    const page = fubFetch('deals?limit=100&offset=' + offset, null, 'get').deals || [];
    deals.push(...page); if (page.length < 100) break;
  }
  deals.forEach(d => {
    const sid = Number(d.stageId), S = PAST_CLIENTS.dealStages;
    const want = S.closed.indexOf(sid) >= 0 ? 'Closed' : S.pending.indexOf(sid) >= 0 ? 'Pending' : S.active.indexOf(sid) >= 0 ? 'Active Client' : '';
    if (!want) return;
    (d.people || []).forEach(x => {
      const p = peopleById[x.id]; if (!p) return;
      if (want === 'Closed') closed[p.id] = d.projectedCloseDate || d.closeDate || d.enteredStageAt || d.updatedAt || new Date().toISOString();
      const cur = order.indexOf(String(p.stage || '').toLowerCase());
      if (cur >= order.indexOf(want.toLowerCase())) return;  // never move anyone backwards
      try {
        fubFetch('people/' + p.id, {stage: want}, 'put');
        fubFetch('notes', {personId: p.id, subject: 'Stage: ' + want, body: `Moved to ${want} because the deal "${d.name}" is in ${d.stageName || 'stage ' + sid}.`});
        p.stage = want; Utilities.sleep(120);
      } catch (e) { console.log('Deal stage sync ' + p.id + ': ' + e.message); }
    });
  });
  return closed;
}

function runPastClientLoop(people) {
  if (!PAST_CLIENTS.live) return 0;
  const props = PropertiesService.getScriptProperties();
  const byId = {}; people.forEach(p => byId[p.id] = p);
  let closedDeals = {};
  try { closedDeals = syncDealStages(byId); } catch (e) { console.log('Deal sync: ' + e.message); }
  const sheet = ensureTab(getSS(), PAST_CLIENTS_TAB), last = sheet.getLastRow(), rows = {};
  (last > 1 ? sheet.getRange(2, 1, last - 1, PAST_CLIENTS_TAB.headers.length).getValues() : [])
    .forEach((r, i) => rows[r[0]] = {row: i + 2, closed: r[2], as: r[3], done: String(r[4] || '').split(',').filter(Boolean)});
  const tagHas = (p, t) => (p.tags || []).some(x => String(x).toLowerCase() === t.toLowerCase());
  const skip = p => /trash/i.test(String(p.stage || '')) || ['Deceased', 'Has Agent', 'Pause', 'No Past Client Loop'].some(t => tagHas(p, t)) || (p.tags || []).some(t => /^Test\b/i.test(t));
  const planIds = Object.keys(FUNNEL.plans).map(k => FUNNEL.plans[k]).concat(FUNNEL.legacyAfterIntroPlan, FUNNEL.reintroPlan, FUNNEL.neighborPlan);
  const today = new Date(), day = Utilities.formatDate(today, 'America/Los_Angeles', 'yyyy-MM-dd');
  // Join: deal closed, stage Closed / Past Client, or tagged "past client".
  people.forEach(p => {
    if (rows[p.id] || skip(p)) return;
    const viaDeal = closedDeals[p.id], viaStage = /^closed$/i.test(String(p.stage || ''));
    const older = !viaDeal && !viaStage && (/past client/i.test(String(p.stage || '')) || tagHas(p, 'past client'));
    if (!viaDeal && !viaStage && !older) return;
    const closedOn = viaDeal ? new Date(viaDeal) : older ? '' : today;
    const as = older ? 'past client' : 'closed';
    try {  // their regular drip stops; this loop takes over
      (fubFetch('actionPlansPeople?personId=' + p.id, null, 'get').actionPlansPeople || [])
        .filter(a => planIds.indexOf(a.actionPlanId) >= 0 && a.status === 'Running')
        .forEach(a => fubFetch('actionPlansPeople/' + a.id, {status: 'Paused'}, 'put'));
    } catch (e) { console.log('Past client ' + p.id + ' plan pause: ' + e.message); }
    // Older clients: the first quarterly check-in lands 30-60 days out, spread so they don't all go the same day.
    const start = older ? new Date(today.getTime() + (30 + (p.id % 30)) * 864e5) : '';
    if (closedOn && !p.customClosingAnniversary) {  // fills FUB's Closing Anniversary field so it shows on their profile
      try { fubFetch('people/' + p.id, {customClosingAnniversary: Utilities.formatDate(closedOn, 'America/Los_Angeles', 'yyyy-MM-dd')}, 'put'); } catch (e) {}
    }
    sheet.appendRow([p.id, p.name || [p.firstName, p.lastName].join(' '), closedOn || start, as, '', '', older ? 'first check-in ' + Utilities.formatDate(start, 'America/Los_Angeles', 'M/d') : 'active']);
    rows[p.id] = {row: sheet.getLastRow(), closed: closedOn || start, as: as, done: []};
    try { fubFetch('notes', {personId: p.id, subject: 'Closed client loop', body: older ? 'Joined the past-client loop (quarterly check-ins, referral asks). Their regular email drip is paused.' : 'Joined the closed-client loop: day 3, week 1, review ask, referral ask, quarterly check-ins, home anniversary.'}); } catch (e) {}
  });
  // Send what's due.
  let sent = Number(props.getProperty('PC_DAY') === day ? props.getProperty('PC_COUNT') : 0) || 0;
  const review = props.getProperty('REVIEW_LINK') || '';
  Object.keys(rows).forEach(id => {
    const r = rows[id], p = byId[id];
    if (!p || skip(p) || sent >= PAST_CLIENTS.perDay || !(r.closed instanceof Date || r.closed)) return;
    const closedOn = new Date(r.closed), age = Math.floor((today - closedOn) / 864e5);
    if (isNaN(age) || age < 0) return;
    let due = null;
    if (r.as === 'closed') due = PAST_CLIENTS.steps.filter(s => age >= s.day && age < s.day + 30 && r.done.indexOf(s.step) < 0 && (!s.needsReview || review))[0] || null;
    // Anniversary: FUB's "Closing Anniversary" field if filled in (works for older clients too), else the closing date.
    const annDate = p.customClosingAnniversary ? new Date(p.customClosingAnniversary) : r.as === 'closed' ? closedOn : null;
    const annAge = annDate && !isNaN(annDate) ? Math.floor((today - annDate) / 864e5) : -1;
    const years = annAge > 0 ? Math.floor(annAge / 365.25) : 0;
    if (!due && years >= 1 && r.done.indexOf('pc-a' + years) < 0 && annAge - Math.round(years * 365.25) < 14) {
      const A = PAST_CLIENTS.anniversary, ord = ['', '1-year', '2-year', '3-year', '4-year', '5-year'][years] || years + '-year';
      due = {step: 'pc-a' + years, subject: A.subject, text: A.text, email: A.email, extra: {years: ord, yearsWord: years === 1 ? 'a year' : years + ' years'}};
    }
    const q = r.as === 'closed' ? Math.floor((age - 90) / 90) : Math.floor(age / 90);  // older clients: closed = first check-in date
    if (!due && (r.as === 'closed' ? age >= 90 : age >= 0) && r.done.indexOf('pc-q' + q) < 0 && !(annAge > 0 && annAge % 365 > 350)) {
      const Q = PAST_CLIENTS.quarterly[q % PAST_CLIENTS.quarterly.length];
      due = {step: 'pc-q' + q, subject: Q.subject, text: Q.text, email: Q.email};
    }
    if (!due) return;
    const extra = Object.assign({review: review}, due.extra || {});
    const results = [];
    const email = (p.emails || []).map(e => e.value).filter(validEmail)[0];
    const textOk = tagHas(p, 'Text OK') && !tagHas(p, 'Text Opt-Out') && due.text;
    if (textOk) {
      ensureTab(getSS(), TEXT_QUEUE_TAB).appendRow([new Date(), p.id, p.name || p.firstName || '', due.step, 'pending', 0, '', pcFill(due.text, p, extra)]);
      results.push('text queued');
    }
    if (due.email && email && !(tagHas(p, 'Bounced') && !tagHas(p, 'Opened Email')) && (!textOk || due.step !== 'pc-3d')) {
      const res = sendFromBcolv(email, pcFill(due.subject, p, extra), pcFill(due.email, p, extra));
      if (res && res.ok) {
        results.push('email sent');
        try { fubFetch('notes', {personId: p.id, subject: 'Email sent: ' + pcFill(due.subject, p, extra), body: 'Closed client loop (' + due.step + '), from julianc@bcolv.com:\n\n' + pcFill(due.email, p, extra)}); } catch (e) {}
      } else results.push('email NOT sent: ' + ((res && res.error) || 'no answer'));
    }
    if (!results.length) results.push('nothing to send (no Text OK, no good email)');
    r.done.push(due.step); sent++;
    sheet.getRange(r.row, 5, 1, 3).setValues([[r.done.join(','), today, due.step + ': ' + results.join(', ')]]);
    // Closed → Past Client once the first month of touches is done.
    if (r.as === 'closed' && age >= 30 && /^closed$/i.test(String(p.stage || ''))) {
      try { fubFetch('people/' + p.id, {stage: 'Past Client'}, 'put'); p.stage = 'Past Client'; } catch (e) {}
    }
    Utilities.sleep(200);
  });
  props.setProperty('PC_DAY', day); props.setProperty('PC_COUNT', String(sent));
  try { ensureTextQueueTrigger(); } catch (e) {}
  console.log('Closed client loop: ' + sent + ' touches today.');
  return sent;
}

/** Signed buyer agreement → a deal card in Buyers › HOT BUYERS (once per person and property). */
function dealForAgreement(email, buyer, property) {
  if (!validEmail(email)) return;
  const p = (fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people || [])[0];
  if (!p) return;
  const name = (buyer || p.name || 'Buyer') + ' · ' + String(property || 'home search').split(',')[0];
  const existing = (fubFetch('deals?personId=' + p.id + '&limit=50', null, 'get').deals || []).filter(d => d.name === name);
  if (existing.length) return;
  fubFetch('deals', {name: name, stageId: PAST_CLIENTS.agreementStage, peopleIds: [p.id], userIds: [fubFetch('me', null, 'get').id]});
  fubFetch('notes', {personId: p.id, subject: 'Deal created', body: 'Buyer agreement signed, so a deal card was added in Buyers › HOT BUYERS: ' + name + '. Move the card as it progresses; the contact stage and follow-up follow it.'});
}

/** Preview: logs who would join the loop and what each would get next, without sending anything. */
function previewPastClientLoop() {
  const me = fubFetch('me', null, 'get').id;
  const people = fubAllPeople().filter(p => !p.assignedUserId || p.assignedUserId === me);
  const list = people.filter(p => /closed|past client/i.test(String(p.stage || '')) || (p.tags || []).some(t => /^past client$/i.test(t)));
  console.log(list.length + ' would be in the loop:\n' + list.map(p => p.id + ' ' + p.name + ' (' + p.stage + ')').join('\n'));
}


// ============================================================
// INSTAGRAM DMs (ManyChat). Your ManyChat flow asks for their email (and optionally phone), then an "External Request"
// action (ManyChat Pro) POSTs JSON to this web app at ?ig=IG_KEY. Run setupInstagram once: it makes the key and prints
// the URL and the body to paste into ManyChat. Each DM lead → Follow Up Boss (source "Instagram", tag "Instagram DM",
// a note with what they said) → a short welcome email from julianc@bcolv.com with your booking link → the nightly
// funnel picks them up after 2 days like any website lead. Texts only if the flow sends "text_ok": "yes" (they agreed).
// A DM is someone reaching out, so they also get one "answer their DM" task (inbound-only rule).
// ============================================================
function instagramWebhook(e) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('IG_KEY') || e.parameter.ig !== props.getProperty('IG_KEY')) return json({status: 'ignored'});
  const b = JSON.parse((e.postData && e.postData.contents) || '{}');
  const clean = v => { v = String(v == null ? '' : v).trim(); return /^\{\{.*\}\}$/.test(v) || /^(null|undefined)$/i.test(v) ? '' : v; };
  const name = clean(b.name) || [clean(b.first_name), clean(b.last_name)].filter(Boolean).join(' ');
  const email = clean(b.email), phone = clean(b.phone).replace(/[^\d+]/g, ''), ig = clean(b.ig_username).replace(/^@/, '');
  const said = clean(b.message || b.last_input), interest = clean(b.interest);
  // Shawna answers the DM (or comment) right away; ManyChat sends her "reply" back to them (map $.reply in the
  // External Request's response). No email/phone yet: she still answers and works toward getting one.
  const reply = said ? igReply(ig || name || 'instagram', said, name, clean(b.source) || 'dm') : '';
  if (!validEmail(email) && phone.replace(/\D/g, '').length < 10) return json({status: 'ok', reply: reply});
  const textOk = /^(yes|true|1)$/i.test(clean(b.text_ok));
  const parts = name.split(/\s+/), first = parts[0] || 'there';
  const tags = ['Instagram DM'].concat(textOk ? ['Text OK'] : []);
  const person = {firstName: parts[0] || ig || 'Instagram', lastName: parts.slice(1).join(' '), tags: tags};
  if (validEmail(email)) person.emails = [{value: email}];
  if (phone) person.phones = [{value: phone}];
  const msg = ['Instagram DM' + (ig ? ' from @' + ig : ''), interest && 'Interested in: ' + interest, said && 'They said: ' + said,
    textOk ? 'Agreed to texts in the DM flow.' : ''].filter(Boolean).join('\n');
  const res = fubFetch('events', {source: 'Instagram', system: 'juliancrider.com', type: 'General Inquiry', message: msg, person: person});
  let id = res.personId || (res.person && res.person.id);
  if (!id && validEmail(email)) id = ((fubFetch('people?limit=1&email=' + encodeURIComponent(email), null, 'get').people || [])[0] || {}).id;
  if (!id && phone) id = ((fubFetch('people?limit=1&phone=' + phone.replace(/\D/g, '').slice(-10), null, 'get').people || [])[0] || {}).id;
  if (id) {
    try { fubFetch('notes', {personId: id, subject: 'Instagram DM', body: msg}); } catch (err) {}
    try { createCallTask(id, 'Instagram DM: answer them in Instagram' + (ig ? ' (@' + ig + ')' : ''), 'Follow Up'); } catch (err) {}
    if (validEmail(email) && !props.getProperty('IG_WELCOME_' + id)) {
      const r = sendFromBcolv(email, 'Thanks for the DM, ' + first, `Hi ${first},\n\nThanks for reaching out on Instagram. I'm Julián, a Realtor with The DeBenedetto Team here in Las Vegas and Henderson.\n\nIf it's easier to talk it through, you can grab a time on my calendar here: juliancrider.com/?book\n\nJust curious, what has you looking at Vegas right now?`);
      if (r && r.ok) { props.setProperty('IG_WELCOME_' + id, '1'); try { fubFetch('notes', {personId: id, subject: 'Email sent: Thanks for the DM', body: 'Instagram welcome email from julianc@bcolv.com'}); } catch (err) {} }
    }
    if (textOk && phone) { try { queueOpenHouseTexts(id, {name: name}, 'website'); } catch (err) {} }
  }
  return json({status: 'ok', personId: id || null, reply: reply});
}

/** Shawna's Instagram reply (DM or comment), as Julián, remembering the last few messages with that account. */
function igReply(handle, text, name, source) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('GROQ_API_KEY');
  if (!key) return '';
  const k = 'IG_' + String(handle).toLowerCase().replace(/[^a-z0-9_.]/g, '').slice(0, 40);
  const hist = JSON.parse(props.getProperty(k) || '[]');
  let out = {};
  try {
    out = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
      {role: 'system', content: assistantPrompt(assistantName(), 'julian') + '\nTHIS IS INSTAGRAM' + (source === 'comment' ? ' (a reply to their comment, sent to them by DM)' : ' DM') +
        '. Keep it short and natural for Instagram. Answer what they said, then when it fits, offer the next step: a quick call (juliancrider.com/?book), homes or a home value, and ask for the best email or number so you can send it. Never pushy.'},
      {role: 'user', content: (name ? 'Their name: ' + name + '\n' : '') + (hist.length ? 'Conversation so far:\n' + hist.map(h => h.r + ': ' + h.t).join('\n') + '\n' : '') +
        'They just ' + (source === 'comment' ? 'commented' : 'messaged') + ': "' + text + '"\n\nReply as JSON only.'}]));
  } catch (e) { console.log('Instagram reply: ' + e.message); }
  const reply = humanize(out.reply || '');
  hist.push({r: 'Them', t: String(text).slice(0, 200)}); if (reply) hist.push({r: 'Julián', t: reply.slice(0, 200)});
  props.setProperty(k, JSON.stringify(hist.slice(-8)));
  if (reply) stat('igReplies');
  return reply;
}

/** Run once: makes the Instagram key and prints what to paste into ManyChat's External Request. */
function setupInstagram() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('IG_KEY')) props.setProperty('IG_KEY', 'ig-' + Utilities.getUuid().slice(0, 8));
  const url = ScriptApp.getService().getUrl() + '?ig=' + props.getProperty('IG_KEY');
  console.log('ManyChat → External Request → POST to:\n' + url + '\n\nBody (JSON), using ManyChat\'s "Add Full Contact Data"/fields:\n' +
    '{"name": "{{full_name}}", "email": "{{email}}", "phone": "{{phone}}", "ig_username": "{{ig_username}}", "message": "{{last_input_text}}", "interest": "{{interest}}", "source": "dm"}\n\n' +
    'In Response mapping, map $.reply to a custom field (e.g. shawna_reply) and send it as the next message. For comment automations use "source": "comment" and {{last_comment_text}} or the comment field.\n\n' +
    'Only add "text_ok": "yes" if your flow asked "Is it OK if I text you?" and they said yes.');
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
       <p>Before any tour, Nevada requires a short written buyer agreement. Mine is non-exclusive and covers only the one home you tour, so you're free to work with anyone else. ${signUrl ? `<a href="${signUrl}">Here it is to look over</a>, along with the Nevada Duties Owed form. Once we pick the home, I'll fill it in and you can sign in about a minute.` : 'Once we pick a home, I\'ll send it to sign. It takes about a minute.'}</p>`
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

const AGREEMENT_TBD = 'To be confirmed with Julián (the home you choose to tour)';
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
  return json({ok: true, live: s.live, status: a.data.status, buyer: a.data.buyer, property: a.data.property, reviewOnly: a.data.property === AGREEMENT_TBD, html: agreementHtml(a.data, s), duties: dutiesHtml(a.data, s, '')});
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
  if (a.data.property === AGREEMENT_TBD) return json({ok: false, error: 'Julián will add the home address first, then send you the link to sign.'});
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
  try { dealForAgreement(a.data.email, a.data.buyer, a.data.property); } catch (err) { console.error('Deal card failed', err); }
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
/** Downloads an audio attachment from the texting Mac and turns it into text (Groq Whisper, free plan). */
function transcribeAudio(att) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('GROQ_API_KEY');
  if (!key || !att || !att.guid) return '';
  const url = props.getProperty('BB_URL').replace(/\/+$/, '') + '/api/v1/attachment/' + encodeURIComponent(att.guid) + '/download?password=' + encodeURIComponent(props.getProperty('BB_PASSWORD'));
  let res;
  for (let i = 0; i < 3; i++) {  // a brand-new attachment can take a few seconds to finish arriving on the Mac
    res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (res.getResponseCode() < 300 && res.getBlob().getBytes().length > 1000) break;
    Utilities.sleep(4000);
  }
  if (res.getResponseCode() >= 300) throw new Error('download ' + res.getResponseCode());
  const name = String(att.transferName || 'audio.m4a').replace(/\.caf$/i, '.m4a');
  const blob = res.getBlob().setName(name);
  const r = UrlFetchApp.fetch('https://api.groq.com/openai/v1/audio/transcriptions', {method: 'post', muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + key},
    payload: {file: blob, model: props.getProperty('GROQ_WHISPER_MODEL') || 'whisper-large-v3-turbo', response_format: 'json'}});
  if (r.getResponseCode() >= 300) throw new Error('Groq audio ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  return String(JSON.parse(r.getContentText()).text || '').trim();
}

function bbSend(phone, text) {
  // One send at a time per person, so two of Shawna's messages never cross and arrive out of order (Nov 12).
  const busyKey = 'BUSY_' + String(phone).replace(/\D/g, '').slice(-10), cache = CacheService.getScriptCache();
  for (let i = 0; i < 4 && cache.get(busyKey); i++) Utilities.sleep(2500);
  cache.put(busyKey, '1', 12);
  try { return bbSendNow(phone, text); } finally { try { cache.remove(busyKey); } catch (e) {} }
}
function bbSendNow(phone, text) {
  // iPhone users sometimes text from their Apple ID email instead of their number: reply to that address too.
  const to = bbE164(phone) || (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(phone || '').trim()) ? String(phone).trim().toLowerCase() : '');
  if (!to) throw new Error('Not a US mobile number: ' + phone);
  // Existing conversation: send into it. New number: on this Mac's macOS (10.14) Messages can't address an unknown
  // "buddy", so start a new chat with the message instead. iMessage first, then green SMS (needs Text Message Forwarding).
  const known = PropertiesService.getScriptProperties().getProperty('BB_CHAT_' + to);
  const tries = (known ? [[known.split(';')[0] + ', the chat they texted from', () => bbApi('message/text', {chatGuid: known, tempGuid: 'jc-' + Utilities.getUuid(), message: text, method: 'apple-script'})]] : []).concat([
    ['iMessage, existing chat', () => bbApi('message/text', {chatGuid: 'iMessage;-;' + to, tempGuid: 'jc-' + Utilities.getUuid(), message: text, method: 'apple-script'})],
    ['iMessage, new chat', () => bbApi('chat/new', {addresses: [to], message: text, service: 'iMessage', method: 'apple-script', tempGuid: 'jc-' + Utilities.getUuid()})],
    ['SMS, existing chat', () => bbApi('message/text', {chatGuid: 'SMS;-;' + to, tempGuid: 'jc-' + Utilities.getUuid(), message: text, method: 'apple-script'})],
    ['SMS, new chat', () => bbApi('chat/new', {addresses: [to], message: text, service: 'SMS', method: 'apple-script', tempGuid: 'jc-' + Utilities.getUuid()})]]);
  let last;
  for (const [label, fn] of tries) {
    try {
      const r = fn();
      // First text to a number we've never had a working chat with: make sure the iMessage really went. Android and
      // Google Voice numbers can't get iMessage; Messages marks it "Not Delivered" and we resend it as a green SMS.
      if (!known && /^iMessage/.test(label) && /^\+1/.test(to) && !bbDelivered(to)) { console.log(label + ' not delivered: trying SMS'); last = new Error('iMessage not delivered'); continue; }
      if (!known && /^\+1/.test(to)) PropertiesService.getScriptProperties().setProperty('BB_CHAT_' + to, (label.split(',')[0] === 'SMS' ? 'SMS' : 'iMessage') + ';-;' + to);
      console.log('Texted ' + to + ' (' + label + ').');
      return {service: label.split(',')[0], result: r};
    }
    catch (err) { last = err; console.log(label + ' failed: ' + String(err.message).slice(0, 300)); }
  }
  throw last;
}

/** After an iMessage to a new number: false if Messages marked it "Not Delivered" (the number can't get iMessage). */
function bbDelivered(to) {
  Utilities.sleep(7000);
  try {
    const m = (bbApi('chat/' + encodeURIComponent('iMessage;-;' + to) + '/message?limit=1&sort=DESC').data || [])[0];
    return !(m && m.isFromMe && Number(m.error) > 0);
  } catch (e) { return true; }  // can't tell: assume it went
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

/** A tap-to-add calendar invite (.ics) for a client's phone, with a 30-minute reminder. */
function icsInvite(title, start, end, location, description) {
  const esc = v => String(v || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
  const utc = d => Utilities.formatDate(new Date(d), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//juliancrider.com//Shawna//EN', 'METHOD:PUBLISH', 'BEGIN:VEVENT',
    'UID:' + Utilities.getUuid() + '@juliancrider.com', 'DTSTAMP:' + utc(new Date()), 'DTSTART:' + utc(start), 'DTEND:' + utc(end),
    'SUMMARY:' + esc(title), 'LOCATION:' + esc(location), 'DESCRIPTION:' + esc(description),
    'BEGIN:VALARM', 'TRIGGER:-PT30M', 'ACTION:DISPLAY', 'DESCRIPTION:' + esc(title), 'END:VALARM', 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
  return Utilities.newBlob(ics, 'text/calendar', title.replace(/[\\/:*?"<>|]/g, '').slice(0, 60) + '.ics');
}

/** The main photo of a listing page (its og:image), or null. Used to send a photo bubble under a listing link.
 *  Some sites (e.g. Zillow) block this; then the link simply goes alone. */
function listingPhoto(url) {
  if (!/^https?:\/\/[^\s]+/.test(url) || /juliancrider\.com\/?(\?|$)/.test(url)) return null;
  const cache = CacheService.getScriptCache(), key = 'OG_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url));
  let img = cache.get(key);
  if (img === 'none') return null;
  try {
    if (!img) {
      const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true});
      const m = res.getResponseCode() < 300 && res.getContentText().slice(0, 200000).match(/<meta[^>]+property=["']og:image(?::secure_url)?["'][^>]+content=["']([^"']+)["']/i);
      img = m ? m[1].replace(/&amp;/g, '&') : 'none';
      cache.put(key, img, 21600);
      if (img === 'none') return null;
    }
    const r = UrlFetchApp.fetch(img, {muteHttpExceptions: true});
    if (r.getResponseCode() >= 300 || r.getBlob().getBytes().length > 4e6 || !/^image\//.test(r.getBlob().getContentType())) return null;
    return r.getBlob().setName('home.' + (/png/.test(r.getBlob().getContentType()) ? 'png' : 'jpg'));
  } catch (e) { return null; }
}

/** Sends "bubble || [voice] || bubble || [contact card]" as separate messages, in order. Returns the first send. */
function bbSendParts(phone, msg) {
  // [welcome voice] = Drive file "Welcome hello" (a general hello for online sign-ups); skipped quietly until it exists.
  const files = {'[voice]': OPEN_HOUSE_VOICE, '[welcome voice]': 'Welcome hello', '[contact card]': CONTACT_CARD};
  let first = null;
  String(msg).split(/\s*\|\|\s*/).filter(Boolean).forEach(part => {
    const fileTok = part.match(/^\[file:\s*(.+?)\s*\]$/i);  // [file: Drive file name or id], e.g. a CMA PDF
    if (files[part.toLowerCase()] || fileTok) {
      if (!first) throw new Error('The first bubble has to be text');
      try { bbSendAudio(phone, first.service, fileTok ? fileTok[1] : files[part.toLowerCase()]); Utilities.sleep(1500); }
      catch (err) { if (fileTok) throw err; console.log(part + ' not sent: ' + err.message); }
    } else {
      const r = bbSend(phone, part); first = first || r; Utilities.sleep(1500);
      // A home link in the message: send its photo right under it, so it reads like a listing card.
      const url = (part.match(/https?:\/\/[^\s)]+/) || [])[0];
      if (url && /zillow|redfin|realtor\.com|homes\.com|realscout|compass|trulia|lvrealtors|gvar|idx|listing|property|homedetails/i.test(url)) {
        const photo = listingPhoto(url);
        if (photo) { try { bbPostAttachment(phone, first && first.service, photo); Utilities.sleep(1500); } catch (e) { console.log('Listing photo: ' + e.message); } }
      }
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
    if (/^[\w-]{25,}$/.test(driveName)) { try { file = DriveApp.getFileById(driveName); } catch (e) {} }  // a Drive file id
    if (!file) {
      const it = DriveApp.searchFiles('title contains "' + driveName.replace(/"/g, '') + '" and trashed = false');
      if (!it.hasNext()) throw new Error('No Google Drive file named "' + driveName + '"');
      file = it.next(); props.setProperty(cacheKey, file.getId());
    }
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
  // AUDIO (Julián, Nov 12): Shawna listens. Voice messages and audio files (a call recording shared from Notes) are turned
  // into text with Groq's free speech-to-text and handled like a text. From Julián (or his own business line): call
  // notes, or a spoken request if it's short. From anyone else: their message, answered as usual.
  if (ev.type === 'new-message') {
    const att = (data.attachments || []).filter(a => /^audio\//i.test(a.mimeType || '') || /\.(m4a|mp3|wav|caf|aac|amr|ogg|opus|mp4)$/i.test(a.transferName || ''))[0];
    const addr = bbE164(String((data.handle && data.handle.address) || ((data.chats || [])[0] || {}).chatIdentifier || '').trim());
    const bossA = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || ''), selfA = addr === TEXTING.myNumber;
    if (att && (!data.isFromMe || selfA)) {
      const c = CacheService.getScriptCache(), dk = 'AUD_' + String(att.guid || att.transferName).slice(0, 80);
      if (c.get(dk)) return json({status: 'ok'});  // the same audio can arrive twice (sent + received copies)
      c.put(dk, '1', 6 * 3600);
      let heard = '';
      try { heard = transcribeAudio(att); } catch (err) { console.log('Audio: ' + err.message); }
      if (addr === bossA || selfA) {
        if (!heard) { try { notifyJulian('🎤 Shawna', 'I got your audio but couldn\'t listen to it (' + (att.transferName || 'audio') + '). Try sharing it as an m4a, or copy the transcript instead.'); } catch (err) {} return json({status: 'ok'}); }
        const who = (String(att.transferName || '').match(/^call (?:with|from) (.+?)\.\w+$/i) || [])[1] || '';
        const isCall = !!who || /^call/i.test(att.transferName || '') || heard.length > 600;
        if (!isCall) c.put('BOSS_VOICE', '1', 120);  // a spoken request: Shawna answers with a voice memo too
        runBossCommand(stashCallNotes(isCall ? 'call notes ' + who + '\n' + heard : heard), 0);
        c.remove('BOSS_VOICE');
        return json({status: 'ok'});
      }
      data.text = heard ? '(voice message) ' + heard : '(they sent a voice message I could not play)';
    }
  }
  // Julián texting his own business number (the "Send call to Shawna" shortcut on the business iPhone sends 2571 → 2571):
  // call notes go to Shawna; anything else to himself is ignored. A note-to-self can arrive twice (sent + received).
  if (ev.type === 'new-message' && data.text) {
    const self = bbE164(String((data.handle && data.handle.address) || ((data.chats || [])[0] || {}).chatIdentifier || '').trim()) === TEXTING.myNumber;
    if (self) {
      if (CALL_NOTES_RE.test(String(data.text))) {
        const c = CacheService.getScriptCache(), k = 'SELF_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(data.text).slice(0, 2000))).slice(0, 24);
        if (!c.get(k)) { c.put(k, '1', 600); runBossCommand(stashCallNotes(String(data.text)), 0); }
      }
      return json({status: 'ok'});
    }
  }
  if (ev.type === 'new-message' && data.isFromMe) {  // Julián (or Shawna) texted: cancel any reply Shawna was holding
    const raw = String((data.handle && data.handle.address) || ((data.chats || [])[0] || {}).chatIdentifier || '').trim();
    const to = bbE164(raw) || (/@/.test(raw) ? raw.toLowerCase() : '');
    if (to) {
      assistantCancel(to);
      const sentByShawna = Number(props.getProperty('AI_SENDING_' + to) || 0) > Date.now() - 90 * 1000;
      const group = /;\+;/.test(String(((data.chats || [])[0] || {}).guid || ''));
      // "Hey Shawna" in his own text to someone = bring Shawna into that conversation (friends, saved contacts, anyone).
      if (!sentByShawna && !group && /\b(?:hey|hi|yo|ok|okay)[,!]?\s+sh(?:aw|au|w)n+a\b/i.test(String(data.text || ''))) {
        props.setProperty('SHAWNA_ON_' + to, String(Date.now()));
        props.deleteProperty('AI_MUTE_' + to);
        // He's watching the thread: she answers right away, and if she can't, he hears why (never silence).
        let done = false, why = '';
        try { done = shawnaDirectReply(to, '', String(data.text)); } catch (e) { why = String(e.message).slice(0, 100); console.log('Hey Shawna: ' + e.message); }
        if (!done) { try { notifyJulian('Shawna here', 'I couldn\'t jump into that chat just now' + (why ? ' (' + why + ')' : '') + '. Text me what you need and I\'ll handle it.'); } catch (e) {} }
        return json({status: 'ok'});
      }
      if (!sentByShawna) props.setProperty('AI_MUTE_' + to, String(Date.now() + Number(props.getProperty('ASSISTANT_MUTE_MIN') || 30) * 6e4));
      if (!sentByShawna) { props.setProperty('AI_LASTME_' + to, String(Date.now())); textLogAnswered(to); }
      const boss1 = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
      if (!sentByShawna && !group && to !== boss1 && /^\+1\d{10}$/.test(to)) {
        try {
          const fp = (fubFetch('people?limit=1&fields=id&phone=' + to.slice(2), null, 'get').people || [])[0];
          if (fp) {
            // Julián, Nov 12: his side of the conversation goes into Follow Up Boss too, so the history is complete.
            try { fubFetch('notes', {personId: fp.id, subject: 'Text from Julián', body: String(data.text).slice(0, 2000)}); } catch (e) {}
            rememberJulianText(to, data.text); detectMeeting(to, fp.id, data.text);
          }
        } catch (e) { console.log('His text: ' + e.message); }
      }
    }
    return json({status: 'ok'});
  }
  if (ev.type !== 'new-message' || !data.handle || !data.text) return json({status: 'ok'});
  const handle = String(data.handle.address || '').trim();
  const isEmailHandle = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(handle);
  const from = bbE164(handle) || (isEmailHandle ? handle.toLowerCase() : '');
  if (!from) return json({status: 'ok'});
  // Remember the exact chat (iMessage or green SMS) they texted from, so replies go back the same way. Google Voice and
  // Android numbers only get SMS; an "iMessage" reply to them looks sent but never arrives.
  const chatGuid = String(((data.chats || [])[0] || {}).guid || '');
  if (chatGuid) props.setProperty('BB_CHAT_' + from, chatGuid);
  // Julián texting his own business number from his personal phone = talking to Shawna (his assistant).
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  if (boss && from === boss) { runBossCommand(stashCallNotes(String(data.text)), 0); return json({status: 'ok'}); }
  let found = (isEmailHandle
    ? fubFetch('people?limit=1&fields=id,firstName,lastName,tags&email=' + encodeURIComponent(from), null, 'get').people
    : fubFetch('people?limit=1&fields=id,firstName,lastName,tags&phone=' + encodeURIComponent(from.slice(2)), null, 'get').people) || [];
  if (!found.length && isEmailHandle) return json({status: 'ok'});  // unknown Apple ID email: leave it alone
  // Toll-free numbers (800, 833, 844, 855, 866, 877, 888) are businesses and verification codes, never leads.
  if (!found.length && /^\+18(00|33|44|55|66|77|88)/.test(from)) return json({status: 'ok'});
  if (!found.length) {
    // Someone new texted the business number. People saved in his iPhone contacts are personal (friends, family):
    // leave them alone. Anyone else is a new lead: add them to FUB so Shawna can answer (they texted first).
    const mode = (props.getProperty('ASSISTANT_MODE') || 'draft').toLowerCase();
    // Saved in his phone = personal, left alone. But if Shawna was already talking with them before he saved the number
    // (a new lead he met and saved), she keeps going.
    // Contacts he allowed (texted "you can talk to Kyle") or brought in with "Hey Shawna": she answers, no CRM entry needed.
    if (Number(props.getProperty('SHAWNA_ON_' + from) || 0) > Date.now() - 3 * 36e5 || JSON.parse(props.getProperty('SHAWNA_ALLOW') || '[]').indexOf(from) >= 0) {
      try { shawnaDirectReply(from, String(data.text).slice(0, 1000), ''); } catch (e) { console.log('Allowed contact: ' + e.message); }
      return json({status: 'ok'});
    }
    const pending = !!props.getProperty('NEWNUM_' + from);
    if (mode === 'off' || props.getProperty('SHAWNA_NEW_NUMBERS') === 'off' || (!pending && macContactByPhone(from))) return json({status: 'ok'});
    // Julián, Nov 11: no more "No Name" contacts. Shawna answers (they texted first) and asks who it is; they go into
    // Follow Up Boss once they've said their name. Until then they're listed in the "New Numbers" sheet tab.
    const added = newNumber(from, String(data.text).slice(0, 1000), mode);
    if (!added) return json({status: 'ok'});
    found = [added];
  }
  const person = found[0], text = String(data.text).slice(0, 1000);
  try { fubFetch('textMessages', {personId: person.id, message: text, fromNumber: from, toNumber: TEXTING.myNumber, isIncoming: true}); }
  catch (err) { fubFetch('notes', {personId: person.id, subject: 'Text from ' + (person.firstName || 'contact'), body: text}); }
  try { detectMeeting(from, person.id, text); } catch (e) { console.log('Meeting: ' + e.message); }
  if (/^\s*(stop|stopall|unsubscribe|end|quit|cancel|opt ?out|stop texting( me)?|remove me)\s*[.!]*\s*$/i.test(text)) {  // the whole text only
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
  textLog({who: [realFirst(person), person.lastName].join(' ').trim() || from, phone: from, text: text.slice(0, 200)});
  if (props.getProperty('TEXT_ALERTS') === 'on') {
    try { notifyJulian([person.firstName, person.lastName].join(' ').trim() + ' texted you', '"' + text.slice(0, 140) + '" Check your messages.'); } catch (e) {}
  }
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

// LIVE INFO (Julián, Nov 11): Shawna should know today's mortgage rates, the market and what's going on in the world,
// like a well read friend. Rates come from the same feed as the website card (Mortgage News Daily, Freddie Mac backup).
// Anything current (news, sports, pop culture, the Vegas market, events, weather) comes from Groq's web search model
// (GROQ_WEB_MODEL, default groq/compound-mini, on the free plan). Answers are cached for an hour.
const RATE_Q = /\b(rates?|interest|apr|mortgage|refi|refinanc\w*|points|fha|va loan|jumbo|conventional loan|arm loan|lender)\b/i;
const WEB_Q = /\b(market|home prices|prices in|inventory|news|who won|score|weather|concert|festival|election|economy|recession|stock market|the fed|trending|new album|movie|nba|nfl|ufc|super bowl|raiders|golden knights|aces|f1|what'?s happening|what'?s going on in|events?|things to do|free in vegas)\b/i;
function ratesText() {
  try {
    const r = JSON.parse(rates().getContent());
    if (!r.ok || !r.rates) return '';
    const n = {f30: '30 yr fixed', f15: '15 yr fixed', fha30: '30 yr FHA', va30: '30 yr VA', jumbo30: '30 yr jumbo'};
    return 'Average mortgage rates (' + r.source + ', ' + r.asOf + '): ' + Object.keys(r.rates).map(k => (n[k] || k) + ' ' + (r.rates[k].rate || r.rates[k]) + '%').join(', ') +
      '. These are averages; a person\'s own rate depends on credit, down payment, loan type and points.';
  } catch (e) { return ''; }
}
function webAnswer(q) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('GROQ_API_KEY');
  q = String(q || '').trim().slice(0, 400);
  if (!key || !q) return '';
  const cache = CacheService.getScriptCache();
  const ck = 'WEB_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q.toLowerCase())).slice(0, 30);
  const hit = cache.get(ck); if (hit !== null) return hit;
  try {
    const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: {Authorization: 'Bearer ' + key}, payload: JSON.stringify({model: props.getProperty('GROQ_WEB_MODEL') || 'groq/compound-mini', max_tokens: 500, temperature: 0.3,
        messages: [{role: 'system', content: 'Search the web and answer with current facts in 2 to 4 short sentences. If it is about real estate, homes or the local scene, use Las Vegas and Henderson, Nevada. Say where the numbers come from and how recent they are. Today is ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE MMMM d, yyyy') + '.'},
          {role: 'user', content: q}]})});
    if (res.getResponseCode() !== 200) { console.log('Web answer ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200)); return ''; }
    const t = String(((JSON.parse(res.getContentText()).choices || [])[0] || {}).message.content || '').replace(/<think>[\s\S]*?<\/think>/g, '').trim().slice(0, 1500);
    cache.put(ck, t, 3600);
    return t;
  } catch (e) { console.log('Web answer: ' + e.message); return ''; }
}
/** Live facts for a text, only when it asks about rates or something current (keeps Groq's free limits for the rest). */
function liveInfo(text, always) {
  const parts = [];
  (String(text || '').match(/https?:\/\/[^\s<>"]+/g) || []).slice(0, 2).forEach(u => {
    const t = pageText(u);
    parts.push(t ? 'The page at ' + u + ' says: ' + t : 'The link ' + u + ' could not be opened (the site blocks it); ask them what they like about it.');
  });
  if (RATE_Q.test(text)) parts.push(ratesText());
  if (always || WEB_Q.test(text)) parts.push(webAnswer(text));
  const out = parts.filter(Boolean).join('\n');
  return out ? '\nLIVE INFO (current, use it if it helps and say where numbers come from):\n' + out : '';
}

// HIS VOICE (Julián, Nov 11): Shawna learns how he texts from his own replies to clients (not hers), last 40 kept.
function rememberJulianText(to, text) {
  text = String(text || '').trim();
  if (text.length < 2 || text.length > 400 || /\bsh(?:aw|au|w)n+a\b/i.test(text) || /^\[|^call notes/i.test(text)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    let v = JSON.parse(props.getProperty('JULIAN_VOICE') || '[]');
    v.push(text.replace(/\s+/g, ' ')); v = v.slice(-40);
    while (JSON.stringify(v).length > 7000) v.shift();
    props.setProperty('JULIAN_VOICE', JSON.stringify(v));
  } catch (e) {}
}
function julianVoice() {
  const v = JSON.parse(PropertiesService.getScriptProperties().getProperty('JULIAN_VOICE') || '[]').slice(-18);
  return v.length < 3 ? '' : '\nHOW JULIÁN REALLY TEXTS (his own recent texts to clients; match his tone, length, words and emoji habits, never copy them word for word):\n' + v.map(t => '• ' + t).join('\n');
}

// MEETINGS FROM TEXTS (Julián, Nov 11): when he and a client agree on a time in their texts (whoever wrote it), Shawna
// puts it on his calendar + Follow Up Boss, and offers to send the client a tap-to-add invite (Reply YES).
const MEET_HINT = /\b(\d{1,2}(:\d\d)?\s*(am|pm)|noon|morning|afternoon|evening|tonight|tomorrow|today|monday|tuesday|wednesday|thursday|friday|saturday|sunday|mon|tue|wed|thu|fri|sat|sun)\b|\b(works|that works|see you|confirmed|sounds good|perfect|deal|i'?ll be there)\b/i;
function detectMeeting(phone, personId, lastText) {
  if (!personId || !MEET_HINT.test(String(lastText || ''))) return;
  const props = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache(), key = props.getProperty('GROQ_API_KEY');
  if (!key || cache.get('MEETCHK_' + personId)) return;
  cache.put('MEETCHK_' + personId, '1', 60);
  const thread = bbThread(phone);
  if (!thread) return;
  let r = {};
  try {
    r = JSON.parse(groqChat(key, props.getProperty('GROQ_ROUTER_MODEL') || 'openai/gpt-oss-20b', [
      {role: 'system', content: 'Read this text conversation between Julián (a Las Vegas Realtor) and a client. Today is ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE yyyy-MM-dd HH:mm') +
        ' in Las Vegas. Did they just AGREE on a specific meeting (a showing, walkthrough, call or meeting) with a day AND a time, confirmed by both sides? Return JSON only: {"agreed": true|false, "start": "YYYY-MM-DDTHH:mm Las Vegas time", "minutes": 60, "title": "short title like Onyx Crest walkthrough", "location": "address or place if said, else empty"}. If the day or time is missing or not yet confirmed, agreed is false.'},
      {role: 'user', content: thread.split('\n').slice(-14).join('\n')}], {maxTokens: 500, noPick: true}));
  } catch (e) { console.log('Meeting check: ' + e.message); return; }
  const m = String(r.start || '').match(/(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
  if (!r.agreed || !m) return;
  const start = Utilities.parseDate(m[1] + ' ' + m[2], 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
  if (start.getTime() < Date.now()) return;
  const dk = 'MEET_' + personId + '_' + m[1] + m[2].replace(':', '');
  if (props.getProperty(dk)) return;
  props.setProperty(dk, '1');
  const p = fubFetch('people/' + personId + '?fields=id,name,firstName,lastName', null, 'get'), who = p.name || [p.firstName, p.lastName].join(' ');
  const end = new Date(start.getTime() + (Number(r.minutes) || 60) * 6e4), title = who + ': ' + humanize(r.title || 'meeting');
  const cal = CalendarApp.getDefaultCalendar();
  const dup = cal.getEvents(new Date(start.getTime() - 36e5), new Date(start.getTime() + 36e5)).some(e => e.getTitle().toLowerCase().indexOf(String(p.firstName || who).toLowerCase()) >= 0);
  if (!dup) {
    cal.createEvent(title, start, end, {location: r.location || '', description: 'Phone: ' + phone + '\nSet in your texts with ' + who + ' (added by ' + assistantName() + ').'});
    try { fubFetch('appointments', {title: title, start: start.toISOString(), end: end.toISOString(), location: r.location || '', invitees: [{personId: personId}]}); } catch (e) {}
  }
  const n = addPending({kind: 'invite', personId: personId, who: who, phone: phone, title: humanize(r.title || 'Meeting with Julián'), start: start.getTime(), end: end.getTime(), location: r.location || ''});
  notifyJulian('📅 ' + who, (dup ? 'You already have ' : 'I put ') + humanize(r.title || 'your meeting') + ', ' + Utilities.formatDate(start, 'America/Los_Angeles', "EEE MMM d 'at' h:mm a") +
    (dup ? ' on your calendar.' : ' on your calendar from your texts.') + ' Want me to send ' + (p.firstName || who) + ' a tap-to-add calendar invite? ' + yesHint(n) + '.', '', who);
}

/** Shawna writes Julián's reply to a client's text thread (he asked her to: "reply to Lisa", "handle it"). */
function draftThreadReply(p, phone, instruction) {
  const props = PropertiesService.getScriptProperties();
  const thread = bbThread(phone), notes = personHistory(p.id), camp = personCampaigns(p.id);
  const out = JSON.parse(groqChat(props.getProperty('GROQ_API_KEY'), props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: assistantPrompt(assistantName(), 'julian') + '\nJULIÁN ASKED YOU TO WRITE HIS NEXT TEXT IN THIS CONVERSATION. Read the whole thread and notes, then answer their latest message(s) in context, as Julián. ' +
      'For price or negotiation questions: show you heard them and that you will work it with them (a call or meeting), but never promise a price, a discount or what a seller will accept. Set escalate false.'},
    {role: 'user', content: 'Contact: ' + (p.name || '') + (thread ? '\nThe text conversation (oldest first):\n' + thread : '') + '\nFollow Up Boss notes:\n' + notes + (camp ? '\nCAMPAIGNS:\n' + camp : '') +
      (instruction ? '\n\nJulián\'s instructions for this reply: ' + instruction : '') + '\n\nReply as JSON only.'}]));
  return humanize(out.reply || '');
}

/** Opens a link someone texted (Julián, Nov 11: Shawna can open links and read web pages), as plain text. */
function pageText(url) {
  const cache = CacheService.getScriptCache(), k = 'PG_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url)).slice(0, 30);
  const hit = cache.get(k); if (hit !== null) return hit;
  let t = '';
  try {
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true, headers: {'User-Agent': 'Mozilla/5.0 (Macintosh) Safari/605'}});
    if (res.getResponseCode() < 300) t = res.getContentText().slice(0, 400000).replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>|<noscript[\s\S]*?<\/noscript>/gi, ' ')
      .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;|&rsquo;/g, "'").replace(/\s+/g, ' ').trim().slice(0, 3500);
  } catch (e) {}
  cache.put(k, t, 3600);
  return t;
}

function assistantReply(person, from, text) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GROQ_API_KEY'), mode = (props.getProperty('ASSISTANT_MODE') || 'draft').toLowerCase();
  if (!key || mode === 'off') return false;
  const name = assistantName();
  const p = fubFetch('people/' + person.id + '?fields=id,firstName,lastName,tags,stage,price', null, 'get');
  const tags = p.tags || [];
  const notes = personHistory(person.id);
  const thread = bbThread(from), first = realFirst(p);
  const r = groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: assistantPrompt(name, assistantPersona(from))},
    {role: 'user', content: 'Contact: ' + (first ? [first, p.lastName].join(' ') : '(name unknown, never guess one)') + '\nStage: ' + (p.stage || '') + '\nTags: ' + tags.join(', ') +
      (thread ? '\nThe text conversation so far (oldest first, from the phone):\n' + thread : '') +
      '\nFollow Up Boss notes, emails and calls (oldest first):\n' + notes + (function () { const c = personCampaigns(p.id); return c ? '\nCAMPAIGNS (set to go out to them):\n' + c : ''; })() + (function () { try { const f = fubFetch('people/' + p.id + '?fields=emails', null, 'get'); const m = emailHistory((f.emails || []).map(e => e.value), 14); return m ? '\nTheir recent emails (oldest first):\n' + m : ''; } catch (e) { return ''; } })() +
      liveInfo(text) + '\n\nThey just texted: "' + text + '"\n\nReply as JSON only.'}]);
  let out; try { out = JSON.parse(r); } catch (e) { out = {reply: '', escalate: true, reason: 'unreadable AI answer'}; }
  learnFromReply(p, out, 'a text');
  if (out.important || out.escalate) stat('neededYou');
  const sens0 = ASSISTANT.escalate.test(text) || tags.indexOf('Has Agent') >= 0 || tags.indexOf('No Shawna') >= 0;
  if (out.important || out.escalate) {  // only what matters reaches Julián's phone (once per person every 2 hours)
    props.setProperty('LAST_ALERT_ID', JSON.stringify({id: person.id, at: Date.now()}));
    const k = 'IMP_' + person.id;
    if (Date.now() - Number(props.getProperty(k) || 0) > 2 * 36e5) {
      props.setProperty(k, String(Date.now()));
      try { notifyJulian((first || 'A client') + ' needs you', (first || 'They') + ' texted: "' + String(text).slice(0, 140) + '"' + (out.why || out.reason ? ' (' + (out.why || out.reason) + ')' : '') + (out.escalate || sens0 ? '\n\nThis one\'s for you. Want me to write the reply? Just say "reply to ' + (first || 'them') + '" and tell me anything to include.' : ' I\'m handling the reply.'), '', [first, p.lastName].join(' ').trim()); } catch (e) {}
    }
  }
  const reply = humanize(out.reply);
  const sensitive = ASSISTANT.escalate.test(text) || tags.indexOf('Has Agent') >= 0 || tags.indexOf('Text Opt-Out') >= 0 || tags.indexOf('No Shawna') >= 0;
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'), cnt = 'AI_CNT_' + person.id;
  const used = (props.getProperty(cnt) || '').split('|'), n = used[0] === today ? Number(used[1]) : 0;
  const muted = Number(props.getProperty('AI_MUTE_' + from) || 0) > Date.now();  // Julián is talking with them himself
  // He said "Hey Shawna" in this thread in the last 3 hours: answer right away, no human pacing, any mode.
  const invited = Number(props.getProperty('SHAWNA_ON_' + from) || 0) > Date.now() - 3 * 36e5;
  const ok = reply && !out.escalate && !sensitive && n < Number(props.getProperty('ASSISTANT_MAX_PER_DAY') || 12) && !muted;
  // Replies to someone who texted first may go until 10pm (ASSISTANT_END_HOUR); texts Julián starts stay 9am-9pm (TCPA).
  const endHour = Number(props.getProperty('ASSISTANT_END_HOUR') || 22);
  if (mode === 'backup' && ok && hour >= 8 && hour < endHour && !invited && props.getProperty('ASSISTANT_FAST') === 'off') {  // hold it; send only if Julián hasn't replied in time
    const delay = Number(props.getProperty('ASSISTANT_DELAY_MIN') || 2);
    const q = JSON.parse(props.getProperty('AI_PENDING') || '{}');
    q[from] = {personId: person.id, reply: reply, next: out.next || '', due: Date.now() + delay * 6e4};
    props.setProperty('AI_PENDING', JSON.stringify(q));
    return true;
  }
  if (mode === 'auto' && ok && hour >= 8 && hour < endHour && !invited && props.getProperty('ASSISTANT_INSTANT') !== 'on' && props.getProperty('ASSISTANT_FAST') === 'off') {
    // Like a person (Julián's own texts to Shawna are still answered instantly). If Julián texts them himself in the
    // meantime, this reply is canceled.
    const q = JSON.parse(props.getProperty('AI_PENDING') || '{}');
    q[from] = {personId: person.id, reply: reply, next: out.next || '', due: Date.now() + humanDelay(text, reply, from) * 1000, paced: true};
    props.setProperty('AI_PENDING', JSON.stringify(q));
    return true;
  }
  const auto = (mode === 'auto' || (mode === 'backup' && props.getProperty('ASSISTANT_FAST') !== 'off') || invited) && ok && hour >= 8 && hour < endHour;
  // Julián, Nov 11: answer within about 10 seconds. A short human pause (3 to 8 s, shorter for short replies), then send,
  // unless he jumped into the thread himself in the meantime.
  if (auto && !invited && props.getProperty('ASSISTANT_INSTANT') !== 'on') {
    Utilities.sleep(Math.round(Math.min(8, 3 + String(reply).length / 60 + Math.random() * 2) * 1000));
    if (Number(props.getProperty('AI_MUTE_' + from) || 0) > Date.now()) return false;
  }
  if (!auto) {
    fubFetch('notes', {personId: person.id, subject: name + ' draft', body: (reply ? 'Suggested reply:\n' + reply : 'No reply drafted.') +
      (out.escalate || sensitive ? '\n\nNeeds you: ' + (out.reason || 'sensitive topic') : '') + (out.next ? '\nNext step: ' + out.next : '')});
    return false;
  }
  const intro = (props.getProperty('AI_INTRO_' + person.id) || assistantPersona(from) === 'julian') ? '' : assistantIntro() + ' ';
  props.setProperty('AI_SENDING_' + from, String(Date.now()));
  bbSendParts(from, intro + reply);
  props.setProperty('AI_INTRO_' + person.id, '1'); props.setProperty(cnt, today + '|' + (n + 1));
  fubFetch('notes', {personId: person.id, subject: 'Text sent', body: intro + reply + '\n\n(sent by ' + name + ')' + (out.next ? '\nNext step: ' + out.next : '')});
  return true;
}

/** How long a person would take to answer (seconds, before the queue's own 0-60 s tick): time to read their text,
 *  time to type the reply, a moment to think, all with natural randomness. Faster when they're texting back and forth
 *  right now. Research (Nov 11): dynamic, length-based delays feel most human; a short "ok" shouldn't take as long as a
 *  paragraph; and keep a live conversation inside a couple of minutes. Actual send: about 10 s to 3 min. */
function humanDelay(theirText, reply, phone) {
  const read = Math.min(String(theirText || '').length / 30, 10);   // reading their text, ~30 characters a second
  const type = Math.min(String(reply || '').length / 5, 60);        // thumb typing, ~5 characters a second
  const think = 2 + Math.random() * 10;
  const lastOut = Number(PropertiesService.getScriptProperties().getProperty('AI_LAST_OUT_' + phone) || 0);
  const pace = Date.now() - lastOut < 3 * 6e4 ? 0.5 : 1;             // they answered within 3 min: keep the rhythm going
  const jitter = 0.7 + Math.random() * 0.6;
  return Math.round(Math.min((read + type + think) * pace * jitter, 10));  // Julián, Nov 11: never longer than 10 s
}

/** Today's (or this week's) calendar, numbered to-dos and what's waiting on Julián. Numbered tasks are remembered so he
 *  can text "do 2" / "done 2". Used by the agenda command and the morning briefing. */
/** "today" | "tomorrow" | "week" from what he said. */
function agendaRange(t) { t = String(t || ''); return /week|7 days|next few days/i.test(t) ? 'week' : /tomorrow|next day|tmrw|tmr/i.test(t) ? 'tomorrow' : 'today'; }
function rangeBounds(range) {
  const now = new Date(), day = d => Utilities.parseDate(Utilities.formatDate(d, 'America/Los_Angeles', 'yyyy-MM-dd') + ' 00:00', 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
  const start = range === 'tomorrow' ? new Date(day(now).getTime() + 864e5) : now;
  const end = new Date(day(now).getTime() + (range === 'week' ? 7 : range === 'tomorrow' ? 2 : 1) * 864e5 - 1);
  return {start: start, end: end};
}

/** His calendars only (Julián, Nov 12: calendar and to-dos are different things). Personal + bcolv business, with
 *  the copies Shawna syncs onto the personal calendar left out so nothing shows twice. */
function calendarText(range) {
  const {start, end} = rangeBounds(range), fmt = (d, f) => Utilities.formatDate(new Date(d), 'America/Los_Angeles', f), evs = [], seen = {};
  const add = (t, s, biz) => { const k = String(t).toLowerCase().replace(/\W+/g, '') + '|' + Math.round(s / 6e4); if (seen[k]) return; seen[k] = 1; evs.push({t: t, s: s, biz: biz}); };
  try { CalendarApp.getDefaultCalendar().getEvents(start, end).forEach(e => { if (!/\[synced by Shawna\]/.test(e.getDescription() || '')) add(e.getTitle(), e.getStartTime().getTime(), false); }); } catch (e) {}
  try { ((bcolvCall({action: 'events', from: start.toISOString(), to: end.toISOString()}) || {}).events || []).forEach(e => add(e.title, e.start, true)); } catch (e) {}
  evs.sort((x, y) => x.s - y.s);
  const label = range === 'week' ? 'this week' : range === 'tomorrow' ? 'tomorrow' : 'today';
  if (!evs.length) return '📅 nothing on your calendars ' + label + '.';
  return '📅 ' + label + ':\n' + evs.slice(0, 15).map(x => '• ' + fmt(x.s, range === 'week' ? 'EEE h:mm a' : 'h:mm a') + ' ' + x.t + (x.biz ? ' (business)' : '')).join('\n');
}

/** His to-do list only: Follow Up Boss tasks due today (overdue included), tomorrow, or this week, plus texts waiting
 *  on him and approvals. Numbered so he can text "do 2" / "done 2". */
function todoText(range) {
  const props = PropertiesService.getScriptProperties(), lines = [], {end} = rangeBounds(range);
  const fmt = (d, f) => Utilities.formatDate(new Date(d), 'America/Los_Angeles', f), today = fmt(new Date(), 'yyyy-MM-dd');
  const tomorrow = fmt(new Date(Date.now() + 864e5), 'yyyy-MM-dd'), last = fmt(end, 'yyyy-MM-dd');
  try {
    const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
    const all = fubFetch('tasks?assignedUserId=' + me + '&isCompleted=false&limit=60&sort=dueDate&fields=id,name,dueDate,personId', null, 'get').tasks || [];
    const tasks = all.filter(t => { const d = String(t.dueDate || '').slice(0, 10); return range === 'tomorrow' ? d === tomorrow : d <= last; });
    const old = tasks.filter(t => String(t.dueDate || '').slice(0, 10) < today).length;
    const list = tasks.slice(0, 10), names = {};
    list.forEach(t => { if (t.personId && !(t.personId in names)) { try { const pp = fubFetch('people/' + t.personId + '?fields=id,name,firstName,lastName', null, 'get'); names[t.personId] = [realFirst(pp), pp.lastName].join(' ').trim() || pp.name; } catch (e) { names[t.personId] = ''; } } });
    if (list.length) props.setProperty('AGENDA_TASKS', JSON.stringify(list.map(t => ({id: t.id, personId: t.personId, due: String(t.dueDate || '').slice(0, 10), name: String(t.name || '').slice(0, 200), who: names[t.personId] || ''}))));
    const label = range === 'week' ? 'this week' : range === 'tomorrow' ? 'tomorrow' : 'today';
    if (list.length) lines.push('✅ to do ' + label + ' (' + tasks.length + (old ? ', ' + old + ' past due' : '') + '):\n' +
      list.map((t, i) => (i + 1) + '. ' + (names[t.personId] ? names[t.personId] + ': ' : '') + String(t.name || '').replace(/^\w+: /, '').replace(/ Answer today\.| Call them back\./, '').slice(0, 80) +
        (String(t.dueDate || '').slice(0, 10) > today ? ' (' + fmt(Utilities.parseDate(String(t.dueDate).slice(0, 10), 'America/Los_Angeles', 'yyyy-MM-dd'), 'EEE') + ')' : '')).join('\n') +
      '\n\n"do 2" to start one, "done 2" to check it off' + (old ? ', "clear all" for everything past due' : '') + '.');
    else lines.push('✅ nothing on your to-do list ' + label + '.');
  } catch (e) { lines.push('✅ I couldn\'t open your tasks just now.'); }
  try {
    const waiting = JSON.parse(props.getProperty('TEXT_LOG') || '[]').filter(x => !x.reply && !x.read && !isOwnNumber(x.phone) && Date.now() - x.at < 2 * 864e5).length;
    if (waiting) lines.push('💬 ' + waiting + ' text' + (waiting > 1 ? 's' : '') + ' may be waiting on you ("who haven\'t I answered").');
    try {
      const sh = getSS().getSheetByName(TEXT_QUEUE_TAB.name);
      if (sh && sh.getLastRow() > 1) {
        const sc = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => r[3] === 'scheduled' && r[4] === 'pending' && r[0] instanceof Date && r[0] <= end);
        if (sc.length) lines.push('🕒 scheduled texts:\n' + sc.slice(0, 6).map(r => '• ' + fmt(r[0], 'EEE h:mm a') + ' to ' + r[2]).join('\n'));
      }
    } catch (e) {}
    const q = pendingQueue();
    if (q.length) lines.push('⏳ waiting for your OK:\n' + q.map(x => '#' + x.n + ' ' + pendingLabel(x)).join('\n') + '\nReply "YES 2" or "NO 2" (or "no all").');
  } catch (e) {}
  return lines.join('\n\n');
}

/** Morning briefing: today's calendar, then today's to-dos. */
function agendaText(week) {
  const range = week ? 'week' : 'today';
  return calendarText(range) + '\n\n' + todoText(range);
}

// ---------- Approvals: several can wait at once, numbered (#1, #2…); plain YES = the newest ----------
function pendingQueue() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('BOSS_QUEUE') || '[]').filter(x => Date.now() - x.at < 6 * 36e5 && !(x.phone && bbE164(x.phone) === TEXTING.myNumber));
}
function addPending(item) {
  const props = PropertiesService.getScriptProperties();
  let q = pendingQueue();
  item.at = item.at || Date.now();
  item.n = q.length ? Math.max.apply(null, q.map(x => x.n || 0)) + 1 : 1;
  q.push(item); q = q.slice(-6);
  props.setProperty('BOSS_QUEUE', JSON.stringify(q));
  props.setProperty('BOSS_PENDING', JSON.stringify(item));
  return item.n;
}
function pendingLabel(x) {
  if (!x.who || /^(undefined|unknown|null)$/i.test(String(x.who))) x = Object.assign({}, x, {who: x.phone ? String(x.phone).replace(/^\+1/, '') : 'someone'});
  return x.kind === 'email' ? 'email reply to ' + x.who : x.kind === 'voice' ? 'voice memo to ' + x.who : x.kind === 'deal' ? 'move ' + x.who + '\'s deal to ' + x.stageName
    : x.kind === 'event' ? 'add ' + x.title + ' to your calendar' : x.kind === 'invite' ? 'calendar invite to ' + x.who : 'text to ' + x.who;
}
function yesHint(n) { return pendingQueue().length > 1 ? 'Reply YES ' + n : 'Reply YES'; }

// ---------- Weekly numbers (for the Friday scorecard) ----------
function weekKey() { return 'STAT_' + Utilities.formatDate(new Date(), 'America/Los_Angeles', "YYYY-'W'ww"); }
function stat(key, n) {
  try {
    const props = PropertiesService.getScriptProperties(), k = weekKey(), o = JSON.parse(props.getProperty(k) || '{}');
    o[key] = (o[key] || 0) + (n || 1); props.setProperty(k, JSON.stringify(o));
  } catch (e) {}
}

// ---------- Deals: Shawna notices offers, contracts and closings and offers to move the deal card ----------
const DEAL_SIGNALS = {
  offer: {buyer: 21, seller: 36, label: 'OFFER', contact: 'Pending'}, under_contract: {buyer: 28, seller: 24, label: 'UNDER CONTRACT (Due Diligence)', contact: 'Pending'},
  closed: {buyer: 25, seller: 26, label: 'CLOSED', contact: 'Closed'}, listing: {buyer: 20, seller: 20, label: 'LISTING PREP', contact: 'Active Client'}};
function suggestDeal(personId, signal, why) {
  const sg = DEAL_SIGNALS[signal]; if (!sg || !personId) return;
  const props = PropertiesService.getScriptProperties(), k = 'DEAL_SUG_' + personId + '_' + signal;
  if (Date.now() - Number(props.getProperty(k) || 0) < 3 * 864e5) return;  // once every 3 days per person and signal
  props.setProperty(k, String(Date.now()));
  const pp = fubFetch('people/' + personId + '?fields=id,name,firstName,lastName', null, 'get'), who = [realFirst(pp), pp.lastName].join(' ').trim() || pp.name;
  const n = addPending({kind: 'deal', personId: personId, who: who, signal: signal, stageName: sg.label});
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  const msg = '🏠 Sounds like ' + who + ' ' + {offer: 'is making an offer', under_contract: 'is under contract', closed: 'closed', listing: 'wants to list'}[signal] +
    (why ? ' (' + String(why).slice(0, 120) + ')' : '') + '. Want me to move their deal to ' + sg.label + ' in Follow Up Boss? ' + yesHint(n) + '.';
  if (boss) { try { bbSend(boss, msg); bossMemory('assistant', msg); } catch (e) {} }
}
function moveDeal(pend) {
  const sg = DEAL_SIGNALS[pend.signal], me = fubFetch('me', null, 'get').id;
  const deals = fubFetch('deals?personId=' + pend.personId + '&limit=10', null, 'get').deals || [];
  const open = deals.filter(d => !/closed|lost/i.test(String(d.stageName || '')))[0] || deals[0];
  const seller = open ? Number(open.pipelineId) === 2 : pend.signal === 'listing';
  const stageId = seller ? sg.seller : sg.buyer;
  if (open) fubFetch('deals/' + open.id, {stageId: stageId}, 'put');
  else fubFetch('deals', {name: pend.who + ' · ' + sg.label.toLowerCase(), stageId: stageId, peopleIds: [pend.personId], userIds: [me]});
  try { fubFetch('people/' + pend.personId, {stage: sg.contact}, 'put'); } catch (e) {}
  fubFetch('notes', {personId: pend.personId, subject: 'Deal moved', body: 'Deal moved to ' + sg.label + ' (Julián approved by text).'});
  return sg.label;
}

/** 8am every day: Shawna texts Julián his day (calendar, numbered to-dos, what's waiting) plus hot leads who opened or
 *  clicked his emails overnight. Turn off with Script Property MORNING_BRIEFING = off. */
function morningBriefing() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MORNING_BRIEFING') === 'off') return;
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  if (!boss) return;
  const parts = ['Good morning ☀️ here\'s your ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE') + ':', agendaText(false)];
  try {
    const sh = getSS().getSheetByName(EMAIL_OPENS_TAB.name);
    if (sh && sh.getLastRow() > 1) {
      const rows = sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 300), 5).getValues().filter(r => r[0] instanceof Date && Date.now() - r[0].getTime() < 16 * 36e5);
      const seen = {}, hot = [];
      rows.forEach(r => { if (!seen[r[2]] && hot.length < 5) { seen[r[2]] = 1; hot.push('• ' + r[2] + (r[3] === 'click' ? ' clicked ' : ' opened ') + '"' + String(r[4]).slice(0, 40) + '"'); } });
      if (hot.length) parts.push('👀 opened your emails since last night:\n' + hot.join('\n'));
    }
  } catch (e) {}
  try { const d = Number(props.getProperty('DUP_COUNT') || 0); if (d) parts.push('🧹 ' + d + ' possible duplicate contacts are listed in the "Duplicates" tab.'); } catch (e) {}
  const msg = parts.join('\n\n');
  try { bbSend(boss, humanize(msg)); bossMemory('assistant', msg); } catch (e) { console.log('Morning briefing: ' + e.message); }
}

/** Fridays 9am: how the week went, by text (and a copy by email to julianc@bcolv.com). */
function weeklyScorecard() {
  const props = PropertiesService.getScriptProperties(), boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  const st = JSON.parse(props.getProperty(weekKey()) || '{}'), lines = [];
  try {
    const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
    const fresh = (fubFetch('people?limit=100&sort=-created&fields=id,source,created,assignedUserId,tags', null, 'get').people || [])
      .filter(p => Date.now() - new Date(p.created) < 7 * 864e5 && (!p.assignedUserId || p.assignedUserId === me) && !(p.tags || []).some(t => /^Test\b/i.test(t)));
    const by = {}; fresh.forEach(p => by[p.source || 'other'] = (by[p.source || 'other'] || 0) + 1);
    lines.push('🆕 new leads: ' + fresh.length + (fresh.length ? ' (' + Object.keys(by).map(k => k + ' ' + by[k]).join(', ') + ')' : ''));
  } catch (e) {}
  try {
    const sh = getSS().getSheetByName(EMAIL_OPENS_TAB.name);
    if (sh && sh.getLastRow() > 1) {
      const rows = sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 1000), 4).getValues().filter(r => r[0] instanceof Date && Date.now() - r[0].getTime() < 7 * 864e5);
      lines.push('📧 people who opened or clicked your emails: ' + Object.keys(rows.reduce((o, r) => (o[r[2]] = 1, o), {})).length + ' (' + rows.filter(r => r[3] === 'click').length + ' clicked)');
    }
  } catch (e) {}
  lines.push('📞 calls Shawna answered: ' + (st.calls || 0) + ' · 📅 appointments confirmed: ' + (st.booked || 0));
  lines.push('💬 texts Shawna handled: ' + (st.shawnaTexts || 0) + ' · needed you: ' + (st.neededYou || 0) + ' · client emails briefed: ' + (st.emailsIn || 0) + ' · your approvals: ' + (st.approved || 0));
  const msg = 'Your week 📊\n\n' + lines.join('\n') + '\n\nHave a great weekend. Reply "how\'s my week look" anytime for what\'s next.';
  if (boss) { try { bbSend(boss, msg); bossMemory('assistant', msg); } catch (e) {} }
  try { MailApp.sendEmail({to: NOTIFY_EMAIL, subject: 'Your week with Shawna', body: msg, name: 'Shawna'}); } catch (e) {}
}

/** Nightly (after scoring): possible duplicate contacts (same phone, email or full name) → "Duplicates" tab. */
const DUPLICATES_TAB = {name: 'Duplicates', color: '#9E9E9E', headers: ['Why', 'Contacts (Follow Up Boss IDs)', 'Suggestion']};
function findDuplicates(people) {
  const groups = {}, add = (k, p) => { if (!k) return; (groups[k] = groups[k] || []).push(p); };
  people.filter(p => !/trash/i.test(String(p.stage || ''))).forEach(p => {
    ((p.phones || []).map(x => String(x.value || '').replace(/\D/g, '').slice(-10)).filter(x => x.length === 10)).forEach(d => add('phone ' + d, p));
    ((p.emails || []).map(x => String(x.value || '').toLowerCase().trim()).filter(validEmail)).forEach(e => add('email ' + e, p));
    const nm = String(p.name || '').toLowerCase().replace(/\s+/g, ' ').trim(); if (nm.split(' ').length > 1) add('name ' + nm, p);
  });
  const seen = {}, rows = [];
  Object.keys(groups).forEach(k => {
    const g = groups[k].filter((p, i, arr) => arr.findIndex(x => x.id === p.id) === i);
    if (g.length < 2) return;
    const ids = g.map(p => p.id).sort((a, b) => a - b).join(','); if (seen[ids]) return; seen[ids] = 1;
    const test = g.filter(p => (p.tags || []).some(t => /^Test\b/i.test(t)));
    rows.push(['same ' + k.split(' ')[0], g.map(p => p.name + ' (' + p.id + ')').join(' · '),
      test.length ? 'Test contact(s): ' + test.map(p => p.name).join(', ') + ' can go to Trash' : 'Merge into the oldest (' + Math.min.apply(null, g.map(p => p.id)) + ')']);
  });
  const sh = ensureTab(getSS(), DUPLICATES_TAB);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows.slice(0, 500).map(r => r.map(cell)));
  PropertiesService.getScriptProperties().setProperty('DUP_COUNT', String(rows.length));
}

/** A usable first name, or ''. Follow Up Boss shows "No Name" when it has none: never greet anyone as "No". */
function realFirst(p) {
  const f = String((p && p.firstName) || '').trim();
  return !f || /^(no|name|no name|unknown|caller|texted in|\+?\d[\d\s().-]*)$/i.test(f) ? '' : f;
}

/** The last ~20 messages of the real text thread with this number (read from the texting Mac), oldest first. */
function bbThread(phone) {
  if (!/^\+1\d{10}$|@/.test(String(phone || ''))) return '';
  const known = PropertiesService.getScriptProperties().getProperty('BB_CHAT_' + phone);
  const guids = [known, 'iMessage;-;' + phone, 'SMS;-;' + phone].filter((g, i, a) => g && a.indexOf(g) === i);
  const fmt = msgs => msgs.map(m => (m.dateCreated ? Utilities.formatDate(new Date(m.dateCreated), 'America/Los_Angeles', 'M/d h:mma') + ' ' : '') +
    (m.isFromMe ? 'Julián: ' : 'Them: ') + String(m.text).replace(/\s+/g, ' ').slice(0, 300)).join('\n');
  for (const g of guids) {
    try {
      const msgs = (bbApi('chat/' + encodeURIComponent(g) + '/message?limit=25&sort=DESC').data || []).filter(m => m.text).reverse();
      if (msgs.length) return fmt(msgs);
    } catch (e) {}
    try {  // older BlueBubbles: the message query endpoint
      const msgs = (bbApi('message/query', {chatGuid: g, limit: 25, offset: 0, sort: 'DESC'}).data || []).filter(m => m.text).reverse();
      if (msgs.length) return fmt(msgs);
    } catch (e) {}
  }
  return '';
}

/** Which Shawna Mailer version answers at BCOLV_MAILER_URL (0 = not reachable). Cached 10 minutes. */
function mailerVersion() {
  const c = CacheService.getScriptCache(), hit = c.get('MAILER_V');
  if (hit) return Number(hit);
  let v = 0;
  try { const r = bcolvCall({action: 'version'}); v = r && r.version ? Number(r.version) : (r && r.error === 'denied' ? -1 : 1); } catch (e) {}
  c.put('MAILER_V', String(v), 600);
  return v;
}
function mailerProblem(v) {
  return v === -1 ? 'The Shawna Mailer says the key doesn\'t match (BCOLV_MAILER_KEY).'
    : v === 0 ? 'The Shawna Mailer isn\'t answering (check BCOLV_MAILER_URL).'
    : 'The Shawna Mailer is still on an older version (' + v + '). In the Shawna Mailer: Deploy → Manage deployments → pencil → Version: New version → Deploy.';
}

/** Run by hand: checks that Shawna can read the bcolv inbox, and prints what she sees. */
function testShawnaEmail() {
  CacheService.getScriptCache().remove('MAILER_V');
  const v = mailerVersion();
  console.log('Shawna Mailer version: ' + v + (v >= 4 ? ' (good)' : ' — ' + mailerProblem(v)));
  if (v >= 4) {
    const r = bcolvCall({action: 'search', query: 'newer_than:3d -category:promotions', max: 5});
    console.log((r.messages || []).map(m => (m.fromMe ? 'You → ' + m.to : m.from) + ' | ' + m.subject).join('\n') || 'No emails found in 3 days.');
  }
}

/** Their emails with Julián over the last 2 weeks (read or unread, both directions), from julianc@bcolv.com via the
 *  Shawna Mailer v4. Oldest first. '' if the mailer is older or they have no email. */
function emailHistory(emails, days) {
  const list = (emails || []).map(e => String(e).trim().toLowerCase()).filter(validEmail).slice(0, 2);
  if (!list.length || !PropertiesService.getScriptProperties().getProperty('BCOLV_MAILER_URL')) return '';
  const q = '{' + list.map(e => 'from:' + e + ' to:' + e).join(' ') + '} newer_than:' + (days || 14) + 'd';
  const r = bcolvCall({action: 'search', query: q, max: 8});
  if (!r || !r.ok) return '';
  return (r.messages || []).sort((x, y) => x.date - y.date).slice(-8).map(m => Utilities.formatDate(new Date(m.date), 'America/Los_Angeles', 'M/d h:mma') + ' ' +
    (m.fromMe ? 'Julián' : 'Them') + ' (email "' + m.subject + '"): ' + String(m.body || '').replace(/\s+/g, ' ').slice(0, 500)).join('\n');
}

/** Everything about a person from the last ~2 weeks, in one place: texts (from the Mac), emails (Gmail), and calls,
 *  voice memos and notes (Follow Up Boss). Used for read-backs, replies and lookups. */
function personTimeline(p, phone) {
  const full = p && p.id ? fubFetch('people/' + p.id + '?fields=id,name,phones,emails', null, 'get') : null;
  phone = phone || (full ? bbE164(((full.phones || [])[0] || {}).value || '') : '');
  const parts = [];
  const texts = phone ? bbThread(phone) : '';
  if (texts) parts.push('TEXTS:\n' + texts.split('\n').slice(-12).join('\n'));
  const mails = full ? emailHistory((full.emails || []).map(e => e.value), 14) : '';
  if (mails) parts.push('EMAILS:\n' + mails);
  if (full) {
    const notes = (fubFetch('notes?personId=' + full.id + '&limit=25', null, 'get').notes || [])
      .filter(n => Date.now() - new Date(n.created) < 14 * 864e5 && !/draft$/i.test(String(n.subject || '')))
      .slice(0, 10).reverse()
      .map(n => String(n.created || '').slice(5, 10) + ' ' + (n.subject ? n.subject + ': ' : '') + String(n.body || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 400));
    if (notes.length) parts.push('CALLS, VOICE MEMOS AND NOTES (Follow Up Boss):\n' + notes.join('\n'));
    const camp = personCampaigns(full.id);
    if (camp) parts.push('CAMPAIGNS (email drips and scheduled texts):\n' + camp);
  }
  return parts.join('\n\n');
}

/** The person's recent Follow Up Boss history for Shawna: notes (texts, calls, emails sent), oldest first. */
function personHistory(personId) {
  return (fubFetch('notes?personId=' + personId + '&limit=15', null, 'get').notes || []).reverse()
    .filter(n => !/draft$/i.test(String(n.subject || '')))
    .map(n => String(n.created || '').slice(0, 10) + ' ' + (n.subject ? n.subject + ': ' : '') + String(n.body || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 350)).join('\n');
}

/** What's set to go out to this person (Julián, Nov 11: Shawna should see the campaigns): their Follow Up Boss action
 *  plans (email drips, running / paused / finished) and texts waiting in the Text Queue tab. FUB's own mass email
 *  (Campaigns) isn't in the API, so it can't be listed. */
function personCampaigns(personId) {
  const lines = [];
  try {
    const cache = CacheService.getScriptCache();
    let names = JSON.parse(cache.get('AP_NAMES') || 'null');
    if (!names) {
      names = {};
      (fubFetch('actionPlans?limit=100', null, 'get').actionPlans || []).forEach(a => names[a.id] = a.name);
      cache.put('AP_NAMES', JSON.stringify(names), 6 * 3600);
    }
    (fubFetch('actionPlansPeople?personId=' + personId, null, 'get').actionPlansPeople || []).forEach(a =>
      lines.push((names[a.actionPlanId] || 'Plan ' + a.actionPlanId) + ': ' + ({Running: 'running', Paused: 'paused', Completed: 'finished'}[a.status] || a.status) + ' (started ' + String(a.created || '').slice(0, 10) + ')'));
  } catch (e) {}
  try {
    const sh = getSS().getSheetByName(TEXT_QUEUE_TAB.name);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
      .filter(r => String(r[1]) === String(personId) && r[4] === 'pending')
      .forEach(r => lines.push('Text "' + r[3] + '" on ' + Utilities.formatDate(new Date(r[0]), 'America/Los_Angeles', 'EEE MMM d h:mm a') + ': ' + String(r[7]).slice(0, 80)));
  } catch (e) {}
  return lines.join('\n');
}

// Tags Shawna adds on her own when a text, email or DM tells her something (Julián, Nov 11). Partner keeps them out of
// the buyer/seller drips (the funnel skips Partner and Has Agent).
const SHAWNA_TAGS = {
  moved_away: ['Moved Away', 'Out of area'], agent: ['Agent', 'Partner'], referral_partner: ['Referral Partner', 'Partner'],
  relocation: ['Relocation'], seller: ['Likely Seller'], investor: ['Investor'], has_agent: ['Has Agent'],
  bought_elsewhere: ['Bought Elsewhere'], renter: ['Renter'], past_client: ['past client']
};
function applyShawnaTags(personId, keys, where) {
  const add = [];
  (Array.isArray(keys) ? keys : String(keys || '').split(/\s*,\s*/)).forEach(k => (SHAWNA_TAGS[String(k).trim()] || []).forEach(t => { if (add.indexOf(t) < 0) add.push(t); }));
  if (!add.length) return [];
  const p = fubFetch('people/' + personId + '?fields=id,name,tags', null, 'get'), have = p.tags || [];
  const fresh = add.filter(t => have.indexOf(t) < 0);
  if (!fresh.length) return [];
  fubFetch('people/' + personId, {tags: have.concat(fresh)}, 'put');
  fubFetch('notes', {personId: personId, subject: 'Shawna tagged: ' + fresh.join(', '), body: 'From ' + where + '. Remove the tag to undo.'});
  return fresh;
}

/** After Shawna reads a reply: save their name if we didn't have it, and act on "not right now" / "not interested". */
function learnFromReply(p, out, where) {
  const nm = String((out && out.name) || '').trim();
  if (nm && !realFirst(p) && !/^(no|unknown|shawna|juli[aá]n)\b/i.test(nm)) {
    const parts = nm.split(/\s+/);
    try { fubFetch('people/' + p.id, {firstName: parts[0], lastName: parts.slice(1).join(' ')}, 'put'); p.firstName = parts[0]; } catch (e) {}
  }
  try { applyIntent(p.id, String((out && out.intent) || ''), where); } catch (e) { console.log('Intent: ' + e.message); }
  try { if (out && DEAL_SIGNALS[out.deal]) suggestDeal(p.id, out.deal, 'from ' + where); } catch (e) { console.log('Deal: ' + e.message); }
  try { if (out && out.tags) applyShawnaTags(p.id, out.tags, where); } catch (e) { console.log('Tags: ' + e.message); }
}

/** "Not right now" → tag Not Now: the morning funnel moves them to the slower Long-Term Nurture (about monthly) and stops
 *  the Hot/Warm drip. "Not interested" → tag Not Interested: their JC drip pauses. Julián removes the tag to undo. */
function applyIntent(personId, intent, where) {
  const tag = intent === 'not_now' ? 'Not Now' : intent === 'not_interested' ? 'Not Interested' : '';
  if (!tag) return;
  const p = fubFetch('people/' + personId + '?fields=id,tags', null, 'get');
  if ((p.tags || []).indexOf(tag) >= 0) return;
  fubFetch('people/' + personId, {tags: (p.tags || []).concat(tag)}, 'put');
  fubFetch('notes', {personId: personId, subject: tag, body: tag === 'Not Now'
    ? 'They said "not right now" in ' + where + '. Next morning they move to the slower Long-Term Nurture emails (about once a month) and the Hot/Warm emails stop. Remove the "Not Now" tag to undo.'
    : 'They said they\'re not interested in ' + where + '. Their email drip pauses next morning. Remove the "Not Interested" tag to undo.'});
}

/** Lets Shawna text and answer people saved in his phone (by name: Follow Up Boss first, then his phone's contacts). */
function allowShawna(names) {
  const props = PropertiesService.getScriptProperties(), allow = JSON.parse(props.getProperty('SHAWNA_ALLOW') || '[]'), ok = [], missing = [];
  names.forEach(nm => {
    let phone = '';
    const hits = (fubFetch('people?limit=5&fields=id,name,phones&q=' + encodeURIComponent(nm), null, 'get').people || []).filter(p => String(p.name || '').toLowerCase() === nm.toLowerCase());
    if (hits.length === 1) phone = bbE164(((hits[0].phones || [])[0] || {}).value || '');
    if (!phone) { const c = macContact(nm); if (c) phone = bbE164(c.phone); }
    if (phone) { if (allow.indexOf(phone) < 0) allow.push(phone); ok.push(nm); } else missing.push(nm);
  });
  props.setProperty('SHAWNA_ALLOW', JSON.stringify(allow.slice(-300)));
  return {ok: ok, missing: missing};
}

/** Hourly: a FUB note "Shawna allow" on contact 2221 (written by Claude) lists names Shawna may text; processed once. */
function processAllowNotes() {
  const pid = Number(PropertiesService.getScriptProperties().getProperty('REQUESTS_PERSON_ID') || 2221);
  (fubFetch('notes?personId=' + pid + '&limit=20', null, 'get').notes || []).filter(n => /^shawna allow$/i.test(String(n.subject || '').trim())).forEach(n => {
    const names = String(n.body || '').split(/\n|,/).map(x => x.trim()).filter(Boolean);
    const r = allowShawna(names);
    fubFetch('notes/' + n.id, {subject: 'Shawna allow - done', body: String(n.body || '') + '\n\nAllowed: ' + (r.ok.join(', ') || 'none') + (r.missing.length ? '\nNo number found: ' + r.missing.join(', ') : '')}, 'put');
  });
}

/** Shawna texting someone who isn't a lead in Follow Up Boss (a saved contact he allowed, or "Hey Shawna" in his own
 *  text). invite = Julián's "Hey Shawna..." message (she joins and introduces herself); else they just texted her. */
function shawnaDirectReply(phone, text, invite) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('GROQ_API_KEY');
  const mode = (props.getProperty('ASSISTANT_MODE') || 'draft').toLowerCase();
  if (!key || mode === 'off') return false;
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H')), endHour = Number(props.getProperty('ASSISTANT_END_HOUR') || 22);
  if (!invite && (hour < 8 || hour >= endHour)) return false;  // he's in the thread when he invites her: any hour
  if (!invite && Number(props.getProperty('AI_MUTE_' + phone) || 0) > Date.now()) return false;  // he's talking with them himself
  // Friends and family: if Julián texted them in the last 6 hours it's his conversation, so Shawna stays out (Nov 12).
  if (!invite && Number(props.getProperty('AI_LASTME_' + phone) || 0) > Date.now() - 6 * 36e5) return false;
  if (!invite && mode === 'draft') return false;
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'), cnt = 'AI_CNT_' + phone;
  const used = (props.getProperty(cnt) || '').split('|'), n = used[0] === today ? Number(used[1]) : 0;
  if (n >= Number(props.getProperty('ASSISTANT_MAX_PER_DAY') || 12)) return false;
  const name = assistantName(), thread = bbThread(phone);
  const out = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: assistantPrompt(name, invite ? 'shawna' : assistantPersona(phone)) + '\nThis person is someone Julián knows personally (a friend, family or colleague, not a lead). Be warm and natural, like a helpful friend. Never pitch a call, a meeting, real estate or Julián\'s services, and never ask to "hop on a call"; real estate only if they bring it up or Julián asked. If their message was meant for Julián, keep it short and say you\'ll pass it along. Never greet anyone as "Unknown".'},
    {role: 'user', content: (thread ? 'The text conversation so far (oldest first):\n' + thread + '\n\n' : '') +
      (invite ? 'Julián just brought you into this conversation by texting: "' + invite + '". Say a quick hi as ' + name + ', his AI assistant, and do what he asked or help with what they need.'
        : 'They just texted: "' + text + '"') + liveInfo(invite || text) + '\n\nReply as JSON only.'}]));
  const reply = humanize(out.reply || '');
  if (!reply || (out.escalate && !invite)) return false;
  props.setProperty('AI_SENDING_' + phone, String(Date.now()));
  bbSend(phone, reply);
  props.setProperty(cnt, today + '|' + (n + 1));
  textLog({who: phone, phone: phone, text: text.slice(0, 200), reply: reply.slice(0, 200)});
  return true;
}

// NEW NUMBERS (Julián, Nov 11): a number that isn't in Follow Up Boss or his phone's contacts texts the business line.
// Shawna answers (they texted first), asks who it is (once), and adds them to FUB only when they've said their name.
// Max 3 replies while the name is unknown. Everything is listed in the "New Numbers" sheet tab.
const NEW_NUMBERS_TAB = {name: 'New Numbers', color: '#F4B400', headers: ['When', 'Number', 'What they texted', 'Status']};
function newNumber(from, text, mode) {
  const props = PropertiesService.getScriptProperties(), k = 'NEWNUM_' + from;
  const st = JSON.parse(props.getProperty(k) || '{"texts":[],"asks":0}');
  st.texts.push(text.slice(0, 300)); st.at = st.at || Date.now();
  const sheet = ensureTab(getSS(), NEW_NUMBERS_TAB), key = props.getProperty('GROQ_API_KEY');
  let out = {};
  if (key) {
    try {
      out = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
        {role: 'system', content: assistantPrompt(assistantName(), assistantPersona(from)) + '\nTHIS PERSON IS NEW: they are not in the contacts and you do not know their name. Answer what they said briefly and warmly, ' +
          (st.asks ? 'and do not ask their name again.' : 'and naturally ask who you\'re texting with (for example "Who do I have the pleasure of texting with?").') + ' Put their name in "name" only if they actually said it.'},
        {role: 'user', content: 'Text conversation so far:\n' + (bbThread(from) || st.texts.map(t => 'Them: ' + t).join('\n')) + '\n\nThey just texted: "' + text + '"\n\nReply as JSON only.'}]));
    } catch (e) { console.log('New number: ' + e.message); }
  }
  const name = String(out.name || '').trim();
  if (name && !/^(no|unknown|shawna|juli[aá]n)\b/i.test(name)) {
    const parts = name.split(/\s+/);
    fubFetch('events', {source: 'Text message', system: 'juliancrider.com', type: 'General Inquiry',
      message: 'Texted the business number:\n' + st.texts.join('\n'), person: {firstName: parts[0], lastName: parts.slice(1).join(' '), phones: [{value: from}], tags: ['Texted In']}});
    props.deleteProperty(k);
    sheet.appendRow([new Date(), from, cell(st.texts.join(' | ').slice(0, 1000)), 'Added to Follow Up Boss as ' + cell(name)]);
    return (fubFetch('people?limit=1&fields=id,firstName,lastName,tags&phone=' + encodeURIComponent(from.slice(2)), null, 'get').people || [])[0] || null;
  }
  if (st.texts.length === 1) sheet.appendRow([new Date(), from, cell(text.slice(0, 300)), 'Waiting for their name (not in Follow Up Boss yet)']);
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H')), endHour = Number(props.getProperty('ASSISTANT_END_HOUR') || 22);
  const reply = humanize(out.reply || ''), muted = Number(props.getProperty('AI_MUTE_' + from) || 0) > Date.now();
  if (reply && !out.escalate && !muted && st.asks < 3 && (mode === 'auto' || mode === 'backup') && hour >= 8 && hour < endHour) {
    props.setProperty('AI_SENDING_' + from, String(Date.now()));
    try { bbSend(from, reply); st.asks++; } catch (e) { console.log('New number reply not sent: ' + e.message); }
  }
  props.setProperty(k, JSON.stringify(st));
  return null;
}

// ASSISTANT_PERSONA = julian (default, Julián's choice Nov 9: replies are written as Julián himself, no intro) or shawna.
function assistantPersona(phone) {
  // California law (B&P 17941) bars using a bot to hide that it's AI in sales conversations: California numbers always
  // get the open "Shawna, Julián's AI assistant" version, with the intro on the first text.
  if (phone && isCaliforniaNumber(phone)) return 'shawna';
  return (PropertiesService.getScriptProperties().getProperty('ASSISTANT_PERSONA') || 'julian').toLowerCase();
}

const CA_AREA_CODES = '209 213 279 310 323 341 350 369 408 415 424 442 510 530 559 562 619 626 628 650 657 661 669 707 714 747 760 805 818 820 831 840 858 909 916 925 949 951'.split(' ');
function isCaliforniaNumber(phone) {
  const d = String(phone).replace(/\D/g, '').slice(-10);
  return d.length === 10 && CA_AREA_CODES.indexOf(d.slice(0, 3)) >= 0;
}

function assistantPrompt(name, persona) {
  const asJulian = (persona || assistantPersona()) === 'julian';
  const who = asJulian
    ? `You ARE Julián J. Crider texting his own clients from his phone. Write in first person as Julián ("I", "me"). Never mention an assistant or AI on your own. If someone sincerely asks whether they're talking to a person, an AI or a bot, be honest: "Ha, good question. You're texting with ${name}, my AI assistant. I read every conversation and I'm happy to jump on a call anytime." Never deny being AI.`
    : `You are ${name}, Julián J. Crider's AI assistant, texting his clients. Be open that you're his assistant.`;
  const voice = asJulian ? julianVoice() : '';
  return `${who}${voice}
Julián is a Realtor® and developer with The DeBenedetto Team at Boeckle & Co. in Las Vegas and Henderson (NV S.0177760, 702-498-2571). Think like a top luxury advisor and a good friend: warm, sharp, genuinely interested in the person, never salesy.
${assistantKnowledge()}

HOW TO TEXT (this matters most):
- Sound like a real person texting a friend he respects. Relaxed, warm, specific to what THEY said. Contractions, natural rhythm, a little personality.
- Catch up like a human first. If they share something personal or casual, respond to that first (a quick reaction, a light question about them) before anything about real estate. Don't rush to business.
- 1 to 3 short sentences. One question at most, and not every text needs a question.
- Emojis: most texts have none. Use one only now and then when it truly fits (😊 🙌🏾 🏡 👀 ☀️ 🔑 👋🏾 😂), never by habit at the end of every text and never in back-to-back texts.
- Longer replies: break them into 2 or 3 short paragraphs with a blank line between, the way a person texts. Short replies stay one line.
- Read the WHOLE conversation first and answer what they actually said in context (who they are, what you already told them, what was planned). Never answer only the last line.
- Chris Voss, used lightly and naturally: mirror a few of their words, label a feeling ("Sounds like…", "Seems like…"), or ask a calm "what" or "how" question. Use a no-oriented question ("Would it be crazy if…?", "Is it a bad idea to…?") only once in a while. Never stack techniques, never sound scripted.
- No corporate phrases ("I'd be happy to assist", "Great question", "valued client", "touch base"). No dashes of any kind (no "—", "–" or " - "); use commas or periods.
- Word choices: "here's what I can do for you" (never "sorry, I can't"), "when we work together" (not "if you decide to"), "I'll confirm that for you" (not "I'm not sure"), "I was thinking of you" (never "sorry to bother you"), "the next steps are". It's fine to say "honestly, this one isn't for everyone" when something isn't a fit.
- Read the conversation history and never repeat yourself or re-introduce yourself.
- LANGUAGE: always reply in the language of their latest text (Spanish, Tagalog, Chinese, anything), even if earlier texts were in English; switch back if they do.
- Never offer to "check back in a few weeks" (or any later check-in) unless they told you they're traveling or away and when they'll be back. Keep the conversation going now instead.

EXAMPLES of the right feel:
Them: "Hey! Crazy week, just got back from Cabo" → "Cabo! 😎 Okay that sounds amazing. How was it? Bet coming back to real life hit different 😂"
Them: "We're still thinking about moving next year" → "That makes sense, no rush at all. What's the one thing that would make next year feel like the right time?"
Them: "Is the Summerlin one still available?" → "Let me confirm that for you right now 👀 Are you hoping to see it this week, or just keeping an eye on it?"
Them: "Not ready yet" → "Totally fair 🙌🏾 I'm here whenever you are. Mind if I send you a couple I think you'd love in the meantime?"

GOALS (gently, over time): understand what they want and when, help them feel taken care of, and when they're ready, offer a quick call (juliancrider.com/?book), a showing (juliancrider.com/?showing) or a home value (juliancrider.com/?value).
WELL ROUNDED: you know real estate and lending inside out and explain it in plain words: loan types (conventional, FHA, VA, jumbo, ARM), pre approval, down payment, PMI, points, DTI, closing costs, earnest money, inspection, appraisal, escrow, title, HOAs, Nevada and Las Vegas specifics, the steps of buying and selling. You also chat easily about Las Vegas, sports, food, pop culture and what's going on in the world, like a well read friend. For anything current (today's rates, the market, news, scores, events) use only the LIVE INFO given and say where it's from; if there's none, say you'll confirm. Rates: share today's averages, say their own rate depends on credit, down payment and loan type, and offer to connect them with a great lender.
RULES: never invent facts, listings, prices, availability, rates or promises. General education about loans and the process is great; no personal lending, legal or tax advice (point to their lender or CPA, or offer a call). Fair Housing: never comment on who lives in an area, schools' quality, safety or crime, religion, race, family status, disability or any protected class. If they bring up an offer, contract, negotiation, commission, a complaint, anything legal, they're upset, or you're unsure, set escalate true and reply with a short, warm holding line (as Julián: "Got it, let me look at that and call you today.").
NAMES: if the contact's name is unknown, never guess or use a placeholder (never "No", "Name" or "there"); when it fits naturally, ask who you're talking with, once.
Return JSON only: {"important": true if Julián should know about this soon (they want to see homes or talk now, a price/offer/contract question, ready to list or buy, upset, urgent or time sensitive), else false, "why": "a few words on why it's important", "reply": "...", "escalate": true|false, "reason": "why Julián needs to step in (if escalate)", "next": "short next step for Julián", "name": "their full name ONLY if they told you in this conversation, else empty", "intent": "not_now if they say not right now, maybe later, not ready yet or check back later; not_interested if they clearly say no, not interested, already bought or stop contacting me; otherwise empty", "deal": "offer if they say they are making or made an offer; under_contract if they're under contract or in escrow; closed if they closed or got the keys; listing if they want to list or sell their home now; otherwise empty", "tags": "comma separated, only what they clearly told you, else empty: moved_away (they moved out of Las Vegas), agent (they're a real estate agent), referral_partner (they offer to send or swap referrals), relocation (moving to Las Vegas from out of state), seller (thinking about selling their home), investor (buying rentals or investments), has_agent (working with another agent), bought_elsewhere (already bought), renter (renting, not buying), past_client (Julián helped them buy or sell before)"}
The CAMPAIGNS section (if given) lists the emails and texts already set to go out to them: don't repeat what a campaign is about to say.`;
}

function groqChat(key, model, messages, opts) {
  opts = opts || {};
  const send = m => UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {method: 'post', contentType: 'application/json',
    muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + key},
    // gpt-oss models think before answering and that thinking counts toward max_tokens: keep it short and leave room,
    // or the JSON gets cut off (Groq 400 "Failed to validate JSON").
    // Groq's free plan counts the max_tokens we ask for against the per-minute budget, so ask only for what's needed.
    payload: JSON.stringify(Object.assign({model: m, messages: messages, temperature: 0.7, max_tokens: opts.maxTokens || 900, response_format: {type: 'json_object'}},
      /gpt-oss/.test(m) ? {reasoning_effort: 'low'} : {}))});
  let res = send(model);
  if (res.getResponseCode() === 400 && /json_validate_failed|Failed to validate JSON/.test(res.getContentText())) {
    try {  // the model's answer is usually still in "failed_generation": use it if it holds a JSON object
      const fg = String((JSON.parse(res.getContentText()).error || {}).failed_generation || ''), j = fg.match(/\{[\s\S]*\}/);
      if (j) { JSON.parse(j[0]); return j[0]; }
    } catch (e) {}
    res = send(model);  // otherwise just ask again once
  }
  if (opts.noPick && (res.getResponseCode() === 404 || /model_not_found|decommissioned/.test(res.getContentText()))) throw new Error('Groq 404 model ' + model);
  if (res.getResponseCode() === 404 || /model_not_found|decommissioned/.test(res.getContentText())) {
    const wasMain = model === (PropertiesService.getScriptProperties().getProperty('GROQ_MODEL') || ASSISTANT.model);
    model = groqPickModel(key);  // Groq retires models now and then: pick the best one this account has and remember it
    if (wasMain) PropertiesService.getScriptProperties().setProperty('GROQ_MODEL', model);
    res = send(model);
  }
  if (res.getResponseCode() === 429 || res.getResponseCode() >= 500) {  // busy or rate-limited: wait a few seconds, try once more
    const hint = Number((res.getContentText().match(/try again in ([\d.]+)s/) || [])[1] || (res.getHeaders() || {})['retry-after'] || 4);
    const wait = Math.min(Math.ceil(hint) + 1, 20);
    Utilities.sleep(wait * 1000);
    res = send(model);
  }
  if (res.getResponseCode() >= 300) {  // still failing: a smaller, faster model on the same account usually has room
    const fb = PropertiesService.getScriptProperties().getProperty('GROQ_FALLBACK') || 'llama-3.1-8b-instant';
    if (fb !== model) { const r2 = send(fb); if (r2.getResponseCode() < 300) { console.log('Groq: used fallback ' + fb); res = r2; } }
  }
  if (res.getResponseCode() >= 300) {  // Groq is down, out of free allowance, or the key stopped working: use the backup brain
    const backup = backupBrain(messages, opts);
    if (backup) return backup;
    throw new Error('Groq ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  return JSON.parse(res.getContentText()).choices[0].message.content;
}

/** Backup brain for when Groq fails (Nov 12: "my Groq died"). Google Gemini's free API key (aistudio.google.com),
 *  saved as GEMINI_API_KEY in the Settings tab. Same messages, same JSON answers. Returns '' when there's no key or it fails too. */
function backupBrain(messages, opts) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('GEMINI_API_KEY');
  if (!key) return '';
  try {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/openai/chat/completions', {method: 'post',
      contentType: 'application/json', muteHttpExceptions: true, headers: {Authorization: 'Bearer ' + key},
      payload: JSON.stringify({model: props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash', messages: messages, temperature: 0.7,
        max_tokens: Math.max((opts && opts.maxTokens) || 900, 2000), response_format: {type: 'json_object'}})});
    if (res.getResponseCode() >= 300) { console.log('Backup brain ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200)); return ''; }
    const out = JSON.parse(res.getContentText()).choices[0].message.content;
    if (out) console.log('Groq failed: used the Gemini backup');
    return out || '';
  } catch (e) { console.log('Backup brain error: ' + e.message); return ''; }
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
/** A text to Julián that must not get lost: if the texting Mac doesn't take it, it's retried every minute (6 h). */
function bossTextSafe(msg) {
  const props = PropertiesService.getScriptProperties(), boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  if (!boss) return false;
  try { bbSend(boss, msg); return true; }
  catch (e) {
    console.log('Text to Julián failed, will retry: ' + e.message);
    const q = JSON.parse(props.getProperty('BOSS_OUTBOX') || '[]'); q.push({m: String(msg).slice(0, 1500), at: Date.now()});
    while (JSON.stringify(q).length > 8000) q.shift();
    props.setProperty('BOSS_OUTBOX', JSON.stringify(q));
    return false;
  }
}

function processAssistantQueue() {
  const props = PropertiesService.getScriptProperties();
  try {
    const out = JSON.parse(props.getProperty('BOSS_OUTBOX') || '[]').filter(x => Date.now() - x.at < 6 * 36e5);
    if (out.length) {
      const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || ''), left = [];
      out.forEach((x, i) => { if (left.length) { left.push(x); return; } try { bbSend(boss, x.m); } catch (e) { left.push(x); } });
      props.setProperty('BOSS_OUTBOX', JSON.stringify(left));
    } else props.deleteProperty('BOSS_OUTBOX');
  } catch (e) {}
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
      const intro = (props.getProperty('AI_INTRO_' + it.personId) || assistantPersona(phone) === 'julian') ? '' : assistantIntro() + ' ';
      if (Number(props.getProperty('AI_MUTE_' + phone) || 0) > Date.now()) return;  // he jumped in meanwhile
      props.setProperty('AI_SENDING_' + phone, String(Date.now()));
      bbSendParts(phone, intro + it.reply);
      props.setProperty('AI_LAST_OUT_' + phone, String(Date.now()));
      stat('shawnaTexts');
      const cnt = 'AI_CNT_' + it.personId, used = (props.getProperty(cnt) || '').split('|');
      props.setProperty('AI_INTRO_' + it.personId, '1'); props.setProperty(cnt, today + '|' + ((used[0] === today ? Number(used[1]) : 0) + 1));
      fubFetch('notes', {personId: it.personId, subject: 'Text sent', body: intro + it.reply + '\n\n(sent by ' + name + (it.paced ? ')' : ', you hadn\'t replied yet)') + (it.next ? '\nNext step: ' + it.next : '')});
      try {
        const who = (fubFetch('people/' + it.personId + '?fields=firstName', null, 'get').firstName) || 'them';
        textLog({who: who, text: '', reply: it.reply.slice(0, 200)});
        if (props.getProperty('TEXT_ALERTS') === 'on') notifyJulian('You received a text from ' + who, "I'm speaking with " + who + ' now: "' + it.reply.slice(0, 140) + '" If you want to jump in, just reply to ' + who + " yourself and I'll stay out of that conversation for 30 minutes.");
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
function notifyJulian(title, msg, url, who) {
  // Two ways to reach Julián: a text from his business number to his personal phone (ALERT_TEXT_TO, falls back to
  // TEST_TEXT_TO; a different phone, so it actually alerts him) and/or a push through the free ntfy app (NTFY_TOPIC).
  const props = PropertiesService.getScriptProperties(), topic = props.getProperty('NTFY_TOPIC');
  const to = props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO');
  let sent = false;
  if (to) sent = bossTextSafe(title + '. ' + msg) || true;  // retried automatically if the Mac is asleep
  // Shawna remembers what she told him, so "what does the email say?" right after an alert knows who he means.
  if (sent) { bossMemory('assistant', title + '. ' + msg); if (who) props.setProperty('LAST_ALERT_WHO', String(who)); }
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
    const em = (fubFetch('emails?personId=' + p.id + '&limit=3&fields=id,status,date,relatedPeople', null, 'get').emails || []);
    const last = em[0];
    // FUB logs Julián's own bcolv emails as "Received" too; only count emails the contact actually sent.
    const fromThem = e => (e.relatedPeople || []).some(r => r.personId === p.id && r.sentByPerson) || !(e.relatedPeople || []).length;
    if (last && last.status === 'Received' && fromThem(last) && !watch[last.id] && new Date(last.date).getTime() > Date.now() - 6 * 36e5)
      watch[last.id] = {p: p.id, name: p.name, at: new Date(last.date).getTime(), told: 0, nudged: 0, fresh: 1};
  });
  // Replies to the drip campaigns land in Follow Up Boss and Gmail. Even if he already opened it, Shawna reads it from
  // Gmail (Shawna Mailer v4), learns from it (not now / not interested / name) and drafts the answer.
  Object.keys(watch).filter(id => watch[id].fresh).forEach(id => {
    const w = watch[id]; delete w.fresh;
    try {
      const em = ((fubFetch('people/' + w.p + '?fields=emails', null, 'get').emails) || []).map(e => String(e.value).toLowerCase()).filter(validEmail)[0];
      if (!em) return;
      const r = bcolvCall({action: 'search', query: 'from:' + em + ' newer_than:2d', max: 2});
      const msgs = ((r && r.messages) || []).filter(m => !m.fromMe).sort((x, y) => x.date - y.date).slice(-1);
      if (msgs.length && processEmailAssistant(msgs) > 0) w.told = 1;  // Shawna already texted him about it
    } catch (e) { console.log('Campaign reply: ' + e.message); }
  });
  const nudgeMin = Number(props.getProperty('EMAIL_NUDGE_MIN') || 60);
  Object.keys(watch).forEach(id => {
    const w = watch[id], link = 'https://app.followupboss.com/2/people/view/' + w.p;
    if (Date.now() - w.at > 2 * 864e5) { delete watch[id]; return; }
    if (hour < 7 || hour >= 22) return;
    if (!w.told) { w.told = 1; return; }  // no bare "you got an email" alert: Shawna's email brief tells him what it says
    if (!w.nudged && Date.now() - w.at > nudgeMin * 6e4) {
      const replied = (fubFetch('emails?personId=' + w.p + '&limit=3&fields=id,status,date,relatedPeople', null, 'get').emails || [])
        .some(e => new Date(e.date).getTime() > w.at && (e.status === 'Sent' || (e.relatedPeople || []).some(r => r.personId === w.p && r.sentByPerson === false)));
      if (!replied) notifyJulian('Still waiting: ' + w.name, w.name + ' emailed you ' + Math.round((Date.now() - w.at) / 6e4) + ' min ago and hasn\'t heard back yet. Want me to read it to you?', link, w.name);
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

function bossCommand(text, tries) {
  const props = PropertiesService.getScriptProperties(), name = assistantName();
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO'));
  const say = msg => {
    const t = humanize(String(msg).replace(new RegExp('^' + name + ':\\s*'), '')); bossMemory('assistant', t);
    // He talked to her with a voice message: she answers with a voice memo only (Julián, Nov 12). The written version is
    // kept, so "transcript" / "text me that" sends it. If the voice can't be made, the text goes instead.
    if (CacheService.getScriptCache().get('BOSS_VOICE') && props.getProperty('ELEVENLABS_API_KEY') && t.length < 900) {
      try {
        bbPostAttachment(boss, '', shawnaVoice(t.replace(/[\u{1F300}-\u{1FAFF}\u2600-\u27BF]/gu, '').replace(/\n+/g, '. ')));
        props.setProperty('LAST_VOICE_TEXT', JSON.stringify({t: t, at: Date.now()}));
        return {service: 'iMessage'};
      } catch (e) { console.log('Voice reply: ' + e.message); }
    }
    return bbSend(boss, t);
  };
  bossMemory('user', text);
  // "transcript" / "text me that" after a voice-memo answer: the written version.
  if (/^\s*(transcript|text( me)?( it| that| back)?( your)?( voice ?memo)?( back)?|type it( out)?|send (me )?the text|in text)\s*[.!?]*\s*$/i.test(text)) {
    const lv = JSON.parse(props.getProperty('LAST_VOICE_TEXT') || 'null');
    if (lv && Date.now() - lv.at < 12 * 36e5) return bbSend(boss, lv.t);
  }
  // A pending "send this text?" waiting for YES / NO
  let pend = JSON.parse(props.getProperty('BOSS_PENDING') || 'null');
  const queue = pendingQueue();
  const ym = text.match(/^\s*(?:yes|y|yeah|yep|yup|send|send it|go|go ahead|ok|okay|bet|run it|do it|approved?)\s*(?:#|number|no\.)?\s*(\d)?\s*[.!👍🏾👍]*\s*$/i);
  // "no all" / "cancel all": clear every approval that's waiting.
  if (/^\s*(no|cancel|clear|nah)\s+(them\s+)?all\s*(approvals?)?\s*[.!]*\s*$/i.test(text) && /^\s*(no|nah|cancel)|approval/i.test(text)) {
    const qn = pendingQueue().length; props.deleteProperty('BOSS_QUEUE'); props.deleteProperty('BOSS_PENDING');
    return say('cleared ' + qn + ' waiting approval' + (qn === 1 ? '' : 's') + '. Nothing was sent.');
  }
  // "clear all" / "read all" / "delete all" / "clear everything": past-due and today's tasks checked off, waiting texts
  // marked read, expired approvals removed. Future tasks stay (those go one at a time: "done 3").
  if (/^\s*(clear|delete|read|done|mark)\s+(them\s+|it\s+)?(all|everything)(\s+(done|read|tasks?|messages?|texts?))?\s*[.!]*\s*$/i.test(text)) {
    let tn = 0, mn = 0;
    try {
      const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id, t0 = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
      (fubFetch('tasks?assignedUserId=' + me + '&isCompleted=false&limit=100&fields=id,dueDate', null, 'get').tasks || [])
        .filter(t => String(t.dueDate || '').slice(0, 10) <= t0).forEach(t => { try { fubFetch('tasks/' + t.id, {isCompleted: true}, 'put'); tn++; } catch (e) {} });
    } catch (e) {}
    try { const log = JSON.parse(props.getProperty('TEXT_LOG') || '[]'); log.forEach(x => { if (!x.reply && !x.read) { x.read = 1; mn++; } }); props.setProperty('TEXT_LOG', JSON.stringify(log)); } catch (e) {}
    return say('done ✅ ' + tn + ' past-due task' + (tn === 1 ? '' : 's') + ' checked off and ' + mn + ' waiting text' + (mn === 1 ? '' : 's') + ' marked read. Future tasks are still there; clear those one at a time ("done 3").');
  }
  const nm = text.match(/^\s*(?:no|n|nah|cancel|stop|don.?t)\s*(?:#|number)?\s*(\d)?(?:\s*[.!]*\s*$|\b)/i);
  const picked = (ym && ym[1]) || (nm && nm[1]);
  if (picked) pend = queue.filter(x => x.n === Number(picked))[0] || null;
  const dropPend = x => { props.setProperty('BOSS_QUEUE', JSON.stringify(pendingQueue().filter(q => q.at !== x.at)));
    const cur = JSON.parse(props.getProperty('BOSS_PENDING') || 'null'); if (cur && cur.at === x.at) props.deleteProperty('BOSS_PENDING'); };
  if (picked && !pend) return say('I don\'t have a #' + picked + ' waiting. ' + (queue.length ? 'Waiting now: ' + queue.map(x => '#' + x.n + ' ' + pendingLabel(x)).join(', ') : 'Nothing is waiting for your OK.'));
  if (pend && pend.kind !== 'schedule' && ym) {
    dropPend(pend);
    if (Date.now() - pend.at > (/^(email|deal|event|invite)$/.test(pend.kind || '') || pend.long ? 6 * 36e5 : 30 * 6e4)) return say(name + ': that one expired. Ask me again.');
    stat('approved');
    if (pend.kind === 'deal') {
      try { return say('done. ' + pend.who + '\'s deal is now in ' + moveDeal(pend) + '. ✅'); } catch (e) { return say('I couldn\'t move that deal: ' + String(e.message).slice(0, 120)); }
    }
    if (pend.kind === 'invite') {  // tap-to-add calendar invite for a meeting set in their texts
      props.setProperty('AI_SENDING_' + pend.phone, String(Date.now()));
      const first = String(pend.who || '').split(' ')[0];
      const r = bbSend(pend.phone, 'Here\'s the invite for ' + Utilities.formatDate(new Date(pend.start), 'America/Los_Angeles', "EEEE 'at' h:mm a") + (first ? ', ' + first : '') + '. Tap it to add it to your calendar.');
      Utilities.sleep(1200);
      try { bbPostAttachment(pend.phone, r && r.service, icsInvite(pend.title + ' with Julián', pend.start, pend.end, pend.location, 'Julián Crider, 702-498-2571')); } catch (e) { return say('the text went, but the invite file didn\'t: ' + String(e.message).slice(0, 80)); }
      fubFetch('notes', {personId: pend.personId, subject: 'Calendar invite sent', body: pend.title + ', ' + Utilities.formatDate(new Date(pend.start), 'America/Los_Angeles', 'EEE MMM d h:mm a') + '\n(sent by ' + name + ' at Julián\'s request)'});
      return say('invite sent to ' + pend.who + '. ✅');
    }
    if (pend.kind === 'event') {
      const start = new Date(pend.start), end = new Date(start.getTime() + (pend.minutes || 30) * 6e4);
      CalendarApp.getDefaultCalendar().createEvent(pend.title, start, end, {description: (pend.phone ? 'Phone: ' + pend.phone + '\n' : '') + 'Booked on a call with ' + name + '.'});
      try { fubFetch('appointments', {title: pend.title, start: start.toISOString(), end: end.toISOString(), invitees: [{personId: pend.personId}]}); } catch (e) {}
      return say('added to your calendar: ' + pend.title + ', ' + Utilities.formatDate(start, 'America/Los_Angeles', 'EEE MMM d, h:mm a') + '. ✅');
    }
    if (pend.kind === 'voice') {
      const audio = shawnaVoice(pend.message);
      props.setProperty('AI_SENDING_' + pend.phone, String(Date.now()));
      let first = String(pend.who || '').split(' ')[0];
      if (/^(unknown|no|name|caller|contact|there|\+?\d)/i.test(first)) first = '';  // never "Hi Unknown"
      const r = bbSend(pend.phone, 'Hi' + (first ? ' ' + first : '') + ', it\'s Shawna, Julián\'s assistant. Here\'s a quick voice note.');
      Utilities.sleep(1500);
      bbPostAttachment(pend.phone, r && r.service, audio);
      fubFetch('notes', {personId: pend.personId, subject: 'Voice memo sent', body: pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
      return say('voice memo sent to ' + pend.who + '. ✅');
    }
    if (pend.kind === 'email') {
      let r;
      if (pend.threadId) {  // reply inside their email thread, from julianc@bcolv.com
        const html = '<div style="font:15px/1.6 Helvetica,Arial,sans-serif;color:#111">' + String(pend.message).split(/\n{2,}|\n/).map(par => '<p style="margin:0 0 12px">' +
          par.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</p>').join('') + '<p style="margin:0">Julián</p>' + EMAIL_FOOTER + '</div>';
        r = bcolvCall({action: 'reply', threadId: pend.threadId, text: pend.message + '\n\nJulián', html: html, draft: false});
      } else r = sendFromBcolv(pend.email, pend.subject, pend.message);
      if (!r.ok) return say('the email did NOT go out: ' + (r.error || 'unknown error') + '.');
      fubFetch('notes', {personId: pend.personId, subject: 'Email sent', body: 'To ' + pend.email + ' from ' + (r.from || 'bcolv') + ': "' + pend.subject + '"\n\n' + pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
      return say('emailed ' + pend.who + ' from ' + (r.from || 'your bcolv email') + '. ✅');
    }
    if (pend.sendAt && pend.sendAt > Date.now()) {  // scheduled for later: it waits in the Text Queue tab (editable there)
      ensureTab(getSS(), TEXT_QUEUE_TAB).appendRow([new Date(pend.sendAt), pend.personId, cell(pend.who), 'scheduled', 'pending', 0, '', cell(pend.message)]);
      ensureTextQueueTrigger();
      fubFetch('notes', {personId: pend.personId, subject: 'Text scheduled', body: 'For ' + Utilities.formatDate(new Date(pend.sendAt), 'America/Los_Angeles', "EEE MMM d 'at' h:mm a") + ':\n' + pend.message});
      return say('scheduled ✅ ' + pend.who + ' gets it ' + Utilities.formatDate(new Date(pend.sendAt), 'America/Los_Angeles', "EEE MMM d 'at' h:mm a") + '. Text "cancel the text to ' + String(pend.who).split(' ')[0] + '" to stop it.');
    }
    props.setProperty('AI_SENDING_' + pend.phone, String(Date.now()));
    bbSendParts(pend.phone, pend.message);
    fubFetch('notes', {personId: pend.personId, subject: 'Text sent', body: pend.message + '\n\n(sent by ' + name + ' at Julián\'s request)'});
    if (pend.taskId) { try { fubFetch('tasks/' + pend.taskId, {isCompleted: true}, 'put'); } catch (e) {} }
    return say(name + ': sent to ' + pend.who + '. ✅');
  }
  if (pend && pend.kind !== 'schedule' && nm) { dropPend(pend); return say(name + ': canceled ' + pendingLabel(pend) + '. Nothing was sent.'); }

  const key = props.getProperty('GROQ_API_KEY');
  if (!key) return say(name + ': add GROQ_API_KEY first.');
  const sys = `You are ${name}, the personal AI assistant of Julián Crider, a Las Vegas Realtor. Julián is texting you. Decide what he wants and return JSON only:
{"action":"jot|notes|tag|research|agenda|taskdone|taskstart|unschedule|text|voicememo|email|note|task|contact|schedule|lookup|hot|emails|calendar|callme|remember|texts|read|unanswered|allow|ask|alerts|claude|chat", "query":"for ask: a Gmail search if email could help (e.g. from:joey newer_than:14d, or summerlin newer_than:7d)", "who":"contact's full name if any", "phone":"their phone if he gave one", "email":"their email if he gave one", "subject":"email subject (for email)", "message":"the text or email body to send / note / task / event title", "when":"for schedule, for a task/reminder due on a later day, and for a text he wants sent LATER (Friday, next week, in 2 weeks, tomorrow morning): YYYY-MM-DDTHH:mm Las Vegas time (no time given: 10:00); empty = send now", "minutes":30, "calendar":"personal|business|unsaid (personal = julian.crider gmail, business = bcolv work calendar; unsaid if he didn't say)", "reply":"your short answer to Julián (for chat)"}
Today is ${Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE yyyy-MM-dd HH:mm')} in Las Vegas.
About Julián: Realtor® and developer with The DeBenedetto Team at Boeckle & Co. in Las Vegas and Henderson; business line 702-498-2571, julianc@bcolv.com, juliancrider.com.
- agenda: his TO-DO list only (tasks, follow-ups, who's waiting on him): "what do I need to do", "anything else I need to do today?", "what's on deck", "what do I got going on", "what's on my to-do list tomorrow". Put "today", "tomorrow" or "week" in "message" (today unless he says otherwise).
- unschedule: cancel a text he scheduled for later ("cancel the text to Lisa", "don't send Dan that message Friday"). Put the person in "who".
- taskdone: he says something is already done or wants tasks cleared ("done 2", "mark 1 and 3 done", "I called Sonia", "I already texted Lisa back", "Dan's handled", "already took care of that", "that's done", "clear the old tasks", "clear them all"). Whenever he says he already did something, it's taskdone (with "who" if he names a person), never a new task or text. "message": the numbers from your last to-do list (comma separated), or "old" (older than a week), or "all"; "who" if he names a person.
- taskstart: he wants to work a task from your to-do list ("do 2", "let's knock out number 3", "start the Sonia one", "handle 1"). "message": the number; "who": the person if he names one.
- text: he wants you to text a client (e.g. "text Lisa and ask if 3pm works"). Write the message in his voice: short, warm, Chris Voss style, no apology, "I was thinking of you" not "sorry to bother you".
- If he wants YOU to answer someone ("reply to Lisa", "you handle it", "answer her", "what should I say back", "write her back"), return action text with "draft": true and put any instructions he gave in "message" (or leave it empty); you will write the reply from the real conversation. If you just told him a client TEXTED, replying means a TEXT, never an email. Use email only when he says email or they emailed.
- Never invent an email address or phone number: only fill "email"/"phone" with ones he actually typed.
- voicememo: he wants you to send someone a voice memo / voice note / audio message in your voice ("message" = what you'll say, written to be spoken, 15 to 40 seconds).
- If you just suggested an email reply and he says how to change it ("make it shorter", "add that I'm free Thursday", "say it like this: ..."), return action email for that same person with the updated "message".
- email: he wants you to email someone or answer their email ("reply to Ruth", "email Ruth back and say..."). Write a short, warm email in his voice, plain paragraphs, no signature (it's signed automatically). If they emailed recently it goes as a reply in that same thread. If he asks you to answer their email, read it in your recent conversation or the data first so the reply fits what they said.
- note: save a note on a contact. task: remind him to do something about a contact.
- emails: who emailed him / did he get any emails ("did I receive any emails", "any emails", "who emailed me", "check my inbox"). calendar: his CALENDAR only (appointments, meetings): "what's on my calendar", "what's my schedule tomorrow", "how's my week look", "am I free Thursday". Put "today", "tomorrow" or "week" in "message". To-dos and calendar are different: never mix them.
- contact: add a new person to Follow Up Boss (name, phone, email).
- schedule: put an appointment on his Google Calendar (who, when, minutes, message = what it is, phone if given). New people are added to Follow Up Boss too.
- lookup: tell him about a contact from Follow Up Boss (stage, tags, recent notes).
- hot: list his hottest leads right now.
- texts: he asks who texted him / what came in / what you answered for him.
- read: he wants to hear what someone actually said: their latest texts and emails ("what did Kyle say?", "read me Sandy's email", "what did these people say?"). Put the person or people in "who", comma separated (use names from your recent conversation when he says "them" or "these people").
- unanswered: who he hasn't replied to yet (texts and emails), with what they said.
- ask: ANY question about his business data: clients, leads, texts, emails, calls you took, website sign-ups and forms, Instagram leads, his calendar, Follow Up Boss ("who called today?", "did anyone ask about Summerlin?", "what did Joey email me?", "any new leads from the website?", "how many open house sign-ins this week?", "what's going on with the Juniper Springs guy?"). Put his question in "message", names in "who", and a Gmail search in "query" when email might help. Prefer ask over chat whenever real data could answer it.
- tag: he wants tags added to or removed from a contact ("tag Margaret referral partner", "put Dan in the investor bucket", "take the Hot tag off Lisa", "she moved to Nashville, tag her right"). "who" = the person; "message" = tags to add, comma separated, written as Follow Up Boss tags (e.g. Referral Partner, Investor, Moved Away, Out of area, Likely Seller, Relocation, Has Agent, Agent, Partner, Not Now, Not Interested, Pause); put tags to remove after "remove:" (e.g. "Investor, remove: Hot").
- lookup also covers "what campaigns is Dan on", "what's set to go out to Lisa", "is she getting emails".
- research: general knowledge or anything current that isn't his own data: mortgage rates today, the Las Vegas market, how loans and the buying or selling process work, news, sports, pop culture, weather, events ("what are rates today", "how's the Vegas market", "what's an FHA loan", "who won the Raiders game", "what's free in Vegas this weekend"). Put his question in "message".
- If you just told him someone emailed or texted and he asks "what does it say", "what's the email say", "read it" or "yes", that's read with that person in "who".
- allow: he wants you to be able to text or answer someone who's saved in his phone contacts (friends or clients he saved), e.g. "you can talk to Kyle Henry". Names in "who", comma separated.
- If he answers "yes", "sure", "please do" or similar to something you just offered, return THAT action (e.g. read) with the same people. Never answer "yes" with chat.
SEVERAL REQUESTS IN ONE TEXT: if he asks for more than one thing ("text Lisa back, remind me to call Dan tomorrow and what's my week look like"), return {"actions": [ {one object per request, same fields as above, in the order he said them} ]} with up to 5. One request = one object as usual.
BE A REAL ASSISTANT: understand every message the way a smart human assistant would, whatever the wording, slang, typos or voice-to-text. Never get stuck and never say you can't. If no action matches exactly, pick the one that gets him closest: ask (answer anything from his calendar, texts, emails, calls, Follow Up Boss, website), text or email (do it for him), agenda (his day or week), or claude (anything that needs something new built). If you need one missing detail, ask one short question in "reply" with action chat. If he asks for several things at once, do the most important one and say in "reply" what you'll do next.
UNDERSTAND HIM LIKE A HUMAN WOULD: Julián texts fast, casual and sometimes urban, often by voice-to-text, with typos, slang and missing words. Figure out what he means, not the exact words, and pick the action. Never say you didn't understand if a reasonable person would get it; if it's truly unclear, ask one short question.
Examples (all mean the same as the action shown):
"who haven't I replied to" / "who I ain't hit back" / "who's waiting on me" / "who I left on read" / "anybody I need to get back to" / "who still needs a reply" → unanswered
"what Kyle say" / "what's Kyle talkin bout" / "what he say" / "read me what Sandy sent" / "what's the word from them" / "lemme see what they said" → read
"hit Kyle up" / "shoot Lisa a text" / "tell him I'm on the way" / "let her know 3 works" / "text bro back" → text
"run it" / "send it" / "bet" / "yup" / "go ahead" / "do that" / "yes please" → yes to what you just offered
"who been hittin my line" / "anybody text me" / "what came in" → texts; "who emailed" / "any emails" / "inbox" → emails
"anything else I need to do" / "what I got today" / "what's on deck" / "what's on my plate" → agenda (to-dos); "what's on my calendar" / "how's my week look" / "what's my schedule tomorrow" / "am I free at 3" → calendar; "who's hot" / "who's ready to move" → hot
"pull up Dan" / "what's up with Dan" / "who is Dan" → lookup; "remind me to call Dan" → task; "put Dan down for Thursday at 2" → schedule
"you can talk to Kyle" / "Kyle's cool, you can text him" / "add Kyle to your list" → allow
- alerts: he wants text alerts turned "on" or "off" (put on or off in "message").
- jot: he wants something saved in one of HIS folders (his personal notebook, not a client's record): meetings (notes from his weekly meetings with Joey or the team), listings (properties: coming available / coming soon, being renovated, active, pending, sold; one property per jot, e.g. "123 Main St coming soon in December"), golf (swing thoughts, his game, practice, rounds), links (web links he sends you), socialpod (anything about SocialPod), personal (anything he just wants written down and remembered: "write this down", "note to self", "save this", "remember that I parked on level 3"). Put the folder in "folder" (golf, links, socialpod, personal, or a new folder name he gives like "recipes") and the note in "message", in his words.
- notes: he wants to see or search a folder ("what listings are coming available", "which ones are we renovating", "what did Joey and I cover last meeting", "what are my swing thoughts", "show my golf notes", "what links did I save about mortgages", "what did I write about SocialPod"). "folder" = which one; "query" = his question if he asked one.
- remember: a business fact or preference for YOU to use when you work (how he likes things done, a fact about his business or a client). Put the fact in "message", written as a full sentence. Personal reminders and notes are jot, not remember.
- callme: he wants you to phone him (e.g. "call me", "give me a call"). Put anything he wants to talk about in "message".
- Never use dashes in anything you write.
- claude: anything you can't do with the actions above, and anything bigger: website changes, designs, emails campaigns, social posts, videos, reports, code, Follow Up Boss setup. Put his request in "message".
- chat: small talk or general questions that don't need his data; answer in 1-3 short sentences. Never invent facts about clients, leads or listings: use ask for those.`;
  let a;
  const calPick = pend && pend.kind === 'schedule' && pickCalendar(text);
  // Call notes: straight to the call-notes action (no routing). If she asked "who was that call with?", a short answer
  // within 30 minutes is the name for those notes.
  let cnKey = (text.match(/\[callnotes:(\w+)\]/) || [])[1] || '', cnWho = '';
  const cnWait = JSON.parse(props.getProperty('CALLNOTES_WAIT') || 'null');
  if (!cnKey && cnWait && Date.now() - cnWait.at < 30 * 6e4 && text.length < 80 && !/\?/.test(text)) { cnKey = cnWait.key; cnWho = text.trim(); }
  if (cnKey) props.deleteProperty('CALLNOTES_WAIT');
  if (calPick) { props.deleteProperty('BOSS_PENDING'); a = Object.assign(pend.a, {calendar: calPick}); }
  else if (/^\s*(save\s+(this|that)\s*(link)?\s*[:\-]?\s*)?https?:\/\/\S+\s*$/i.test(text)) a = {action: 'jot', folder: 'links', message: text};
  else if (/^\s*(golf|swing( thoughts?)?|golf notes?)\s*[:\-]/i.test(text)) a = {action: 'jot', folder: 'golf', message: text.replace(/^\s*[^:\-]+[:\-]\s*/, '')};
  else if (/^\s*(meeting( notes?)?|joey meeting|team meeting)\s*[:\-]/i.test(text)) a = {action: 'jot', folder: 'meetings', message: text.replace(/^\s*[^:\-]+[:\-]\s*/, '')};
  else if (/^\s*(listings?|new listing|coming soon|renovating|renovation)\s*[:\-]/i.test(text)) a = {action: 'jot', folder: 'listings', message: text};
  else if (/^\s*social\s*pod( notes?)?\s*[:\-]/i.test(text)) a = {action: 'jot', folder: 'socialpod', message: text.replace(/^\s*[^:\-]+[:\-]\s*/, '')};
  else if (/^\s*(note to self|personal( note)?|jot( this| that)?( down)?|write (this|that) down)\s*[:\-]/i.test(text)) a = {action: 'jot', folder: 'personal', message: text.replace(/^\s*[^:\-]+[:\-]\s*/, '')};
  else if (/^\s*(done|finished|complete|check off)\s*#?\d+([\s,&]+(and\s+)?#?\d+)*\s*[.!]*\s*$/i.test(text)) a = {action: 'taskdone', message: text.replace(/^\D+/, '')};  // "done 5": no AI needed
  else if (/^\s*(do|start)\s*#?\d+\s*[.!]*\s*$/i.test(text)) a = {action: 'taskstart', message: text.replace(/^\D+/, '')};
  else if (cnKey) {
    const full = CacheService.getScriptCache().get(cnKey) || text;
    const d = callNotesParse(full);
    const head = (String(full).split('\n')[0].match(CALL_NOTES_RE) ? String(full).split('\n')[0].replace(CALL_NOTES_RE, '').replace(/^[\s:,-]*(with|from)?\s*/i, '').replace(/^call\s+(with|from)\s+/i, '').split(/[:.!?–—]|\s-\s/)[0].trim() : '');
    const headPhone = /^[+\d\s().-]{10,}$/.test(head) ? head : '';
    const dk = 'CNDONE_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(full).slice(0, 3000))).slice(0, 24);
    if (!cnWho && CacheService.getScriptCache().get(dk)) return;  // the same call sent twice (automation + by hand): once is enough
    CacheService.getScriptCache().put(dk, '1', 6 * 3600);
    a = {action: 'callnotes', who: cnWho || (head && !headPhone && head.length < 40 ? head : '') || d.who || '', phone: headPhone || d.phone || '', email: d.email || '', data: d, key: cnKey, full: String(full).slice(0, 8000)};
  }
  else {
    const history = JSON.parse(props.getProperty('BOSS_HISTORY') || '[]').slice(0, -1).slice(-12)  // earlier messages, not this one
      .map(h => ({r: h.r, t: String(h.t).slice(0, 250)}));
    const msgs = [{role: 'system', content: sys + '\nYour recent conversation with Julián is below; use it so he never has to repeat himself (e.g. "her", "that time", "same as before").'}]
      .concat(history.map(h => ({role: h.r, content: h.r === 'assistant' ? JSON.stringify({action: 'chat', reply: h.t}) : h.t})), [{role: 'user', content: text}]);
    // Groq sometimes stumbles on a long history: retry with just the last few messages before giving up.
    // Working out what he wants runs on a fast, smaller model with its own Groq allowance (GROQ_ROUTER_MODEL), so his
    // texts rarely wait behind client replies and email briefs; the main model is the backup.
    const ask = (m, model) => {
      let raw;
      try { raw = groqChat(key, model, m, {maxTokens: 1000, noPick: model !== (props.getProperty('GROQ_MODEL') || ASSISTANT.model)}); }
      catch (e) {
        if (!/404/.test(e.message)) throw e;
        props.setProperty('GROQ_ROUTER_MODEL', 'llama-3.1-8b-instant');  // that model isn't on this account: use the small one
        raw = groqChat(key, 'llama-3.1-8b-instant', m, {maxTokens: 1000, noPick: true});
      }
      try { return JSON.parse(raw); } catch (e) { const j = String(raw).match(/\{[\s\S]*\}/); if (j) return JSON.parse(j[0]); throw e; } };
    try { a = ask(msgs, props.getProperty('GROQ_ROUTER_MODEL') || 'openai/gpt-oss-20b'); }
    catch (e) {
      shawnaLog('error', text, 'first try: ' + String(e.message).slice(0, 200));
      try { a = ask([msgs[0]].concat(msgs.slice(-5)), props.getProperty('GROQ_MODEL') || ASSISTANT.model); }
      catch (e2) {
        shawnaLog('error', text, 'retry: ' + String(e2.message).slice(0, 200));
        const why = /per day|TPD|RPD/i.test(e2.message) ? 'Groq\'s free plan used up today\'s allowance' : /429|rate/i.test(e2.message) ? 'Groq\'s free plan hit its per-minute limit' : /413|too large/i.test(e2.message) ? 'that request was too big for Groq\'s free plan'
          : /401|invalid_api_key|Invalid API Key/i.test(e2.message) ? 'my Groq key was rejected: make a new one at console.groq.com and put it in the Settings tab as GROQ_API_KEY'
          : /5\d\d/.test(e2.message) ? 'Groq is having a hiccup' : 'Groq error ' + String(e2.message).slice(0, 60);
        if (!tries) say('give me a minute (' + why + '). I\'ll answer you as soon as it\'s back.');
        throw new Error('Groq: ' + String(e2.message).slice(0, 120));  // runBossCommand retries in a minute (up to 3 times)
      }
    }
  }
  // The router sometimes invents an email (lisa.dragos@example.com) or phone: keep only ones he actually typed.
  if (!cnKey && !calPick && a) {
    const digits = text.replace(/\D/g, ''), low = text.toLowerCase();
    (Array.isArray(a.actions) ? a.actions : [a]).forEach(x => {
      if (!x) return;
      if (x.email && low.indexOf(String(x.email).toLowerCase()) < 0) delete x.email;
      if (x.phone) { const d = String(x.phone).replace(/\D/g, '').slice(-7); if (d.length < 7 || digits.indexOf(d) < 0) delete x.phone; }
    });
  }
  // Several requests in one text ("text Lisa, remind me to call Dan and what's my week look like"): do each in order.
  const doAction = a => {
    const findOne = who => {
      const q = String(who || '').trim(), get = term => fubFetch('people?limit=8&fields=id,name,firstName,lastName,tags,phones,created&q=' + encodeURIComponent(term), null, 'get').people || [];
      let ppl = q ? get(q) : [];
      if (!ppl.length && q.split(/\s+/).length > 1) {  // "Lily Montenegro" → search the last name, keep first names that start the same
        const parts = q.toLowerCase().split(/\s+/), first = parts[0].slice(0, 3);
        ppl = get(parts[parts.length - 1]).filter(p => String(p.firstName || p.name || '').toLowerCase().indexOf(first) === 0);
      }
      const exact = ppl.filter(p => (p.name || '').toLowerCase() === q.toLowerCase());
      if (exact.length) return exact.sort((x, y) => String(x.created).localeCompare(String(y.created)))[0];  // oldest = the real one
      if (ppl.length > 1) {  // "Lisa" right after "Lisa needs you": the Lisa she just told him about
        const la = JSON.parse(props.getProperty('LAST_ALERT_ID') || 'null');
        const hit = la && Date.now() - la.at < 12 * 36e5 && ppl.filter(x => x.id === la.id)[0];
        if (hit) return hit;
      }
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
      let subject = humanize(a.subject || 'Quick note from Julián'), body = humanize(a.message || ''), threadId = '';
      // Keep the conversation in one thread: answer in the latest email thread with them from the last 60 days (opened or
      // not, theirs or a campaign email we sent). With Shawna Mailer v5 the reply always goes to them; with an older
      // mailer only a thread whose latest message is theirs is safe to reply in.
      try {
        const v5 = mailerVersion() >= 5;
        const r = bcolvCall({action: 'search', query: '{from:' + email + ' to:' + email + '} newer_than:60d', max: 3});
        const msgs = ((r && r.messages) || []).sort((x, y) => y.date - x.date);
        const pick = msgs.find(m => v5 || (!m.fromMe && !msgs.some(o => o.threadId === m.threadId && o.fromMe && o.date > m.date)));
        if (pick) { threadId = pick.threadId; subject = /^re:/i.test(pick.subject) ? pick.subject : 'Re: ' + pick.subject; }
      } catch (e) {}
      addPending({kind: 'email', personId: p.id, who: p.name || a.who, email: email, subject: subject, message: body, threadId: threadId, at: Date.now()});
      return say((threadId ? 'reply to ' + (p.name || a.who) + '\'s email' : 'send this email to ' + (p.name || a.who)) + ' (' + email + ')?\n\nSubject: ' + subject + '\n\n' + body + '\n\nReply YES to send or NO to cancel.');
    }
    if (a.action === 'voicememo') {
      if (!props.getProperty('ELEVENLABS_API_KEY') || !props.getProperty('ELEVENLABS_VOICE_ID')) return say('voice memos aren\'t connected yet. Add ELEVENLABS_API_KEY and ELEVENLABS_VOICE_ID (Claude has the steps).');
      const {p} = resolvePerson();
      if (!p || p.many) return say(p && p.many ? 'which one? ' + p.many.map(x => x.name).join(', ') : 'who should get the voice memo? Send me their name and number.');
      const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
      if (!phone) return say((p.name || a.who) + ' has no mobile number yet. Send me their number.');
      addPending({kind: 'voice', personId: p.id, who: p.name || a.who, phone: phone, message: humanize(a.message), at: Date.now()});
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
      if (a.action === 'task') {  // a reminder on the day he said ("Friday", "next week"), else today
        const wm = String(a.when || '').match(/(\d{4}-\d{2}-\d{2})/), due = wm ? wm[1] : Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
        fubFetch('tasks', {personId: p.id, type: /\btext\b/i.test(a.message || '') ? 'Text' : /\bcall\b/i.test(a.message || '') ? 'Call' : /\bemail\b/i.test(a.message || '') ? 'Email' : 'Follow Up', name: String(a.message || 'Follow up').slice(0, 250), dueDate: due});
        return say(name + ': reminder added for ' + p.name + ' (' + (wm ? Utilities.formatDate(Utilities.parseDate(due, 'America/Los_Angeles', 'yyyy-MM-dd'), 'America/Los_Angeles', 'EEE MMM d') : 'today') + '). ✅');
      }
      const tags = p.tags || [];
      if (tags.indexOf('Text Opt-Out') >= 0) return say(name + ': ' + p.name + ' opted out of texts, so I can\'t text them. Here\'s what I can do for you: add a task to call instead.');
      const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
      if (!phone) return say(p.name + ' has no mobile number yet. Send me their number and I\'ll text them.');
      const wm = String(a.when || '').match(/(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
      let sendAt = wm ? Utilities.parseDate(wm[1] + ' ' + wm[2], 'America/Los_Angeles', 'yyyy-MM-dd HH:mm').getTime() : 0;
      if (sendAt && sendAt < Date.now() + 10 * 6e4) sendAt = 0;  // "now" or the past: just send it
      if (sendAt) {  // texts only go out 9am-9pm: nudge a late or early time into that window
        const h = Number(Utilities.formatDate(new Date(sendAt), 'America/Los_Angeles', 'H'));
        if (h < 9) sendAt += (9 - h) * 36e5; else if (h >= 21) sendAt += (24 - h + 9) * 36e5;
      }
      const whenTxt = sendAt ? Utilities.formatDate(new Date(sendAt), 'America/Los_Angeles', "EEE MMM d 'at' h:mm a") : '';
      if (a.draft || !String(a.message || '').trim()) {  // "reply to Lisa" / "handle it": written from the real conversation
        try { a.message = draftThreadReply(p, phone, a.draft ? String(a.message || '') : ''); } catch (e) {}
        if (!a.message) return say('I couldn\'t write that one just now. Tell me what to say and I\'ll send it.');
      }
      const n = addPending({personId: p.id, who: p.name, phone: phone, message: a.message, at: Date.now(), sendAt: sendAt});
      return say(name + ': ' + (sendAt ? 'schedule this for ' + p.name + ' on ' + whenTxt : 'send this to ' + p.name) + (tags.indexOf('Text OK') < 0 ? ' (heads up: no Text OK tag)' : '') + '?\n\n"' + a.message + '"\n\n' + yesHint(n) + ' or NO to cancel.');
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
      const cal = which === 'business' ? null : whichCalendar('personal');
      let p = (a.who || a.phone) ? resolvePerson().p : null;
      if (p && p.many) p = null;
      const title = (a.message || 'Appointment') + (a.who ? ': ' + a.who : '');
      const phone = a.phone || (p && p.phones && p.phones[0] && p.phones[0].value) || '';
      const desc = (phone ? 'Phone: ' + phone + '\n' : '') + 'Added by ' + name + ' from Julián\'s text.';
      if (cal) cal.createEvent(title, start, end, {description: desc});
      else {
        const r = bcolvCall({action: 'event', title: title, start: start.toISOString(), end: end.toISOString(), description: desc});
        if (!r.ok) return say('I couldn\'t add it to your business calendar: ' + (r.error || 'unknown error') + '. Run setupMailer again in the Shawna Mailer and redeploy it.');
      }
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
        '\nTags: ' + (full.tags || []).slice(0, 8).join(', ') + (function () { const c = personCampaigns(p.id); return c ? '\nCampaigns:\n' + c.split('\n').map(x => '• ' + x).join('\n') : '\nCampaigns: none'; })() + (notes ? '\nLatest:\n' + notes : ''));
    }
    if (a.action === 'hot') {
      const ppl = fubFetch('people?limit=8&sort=-lastActivity&tags=' + encodeURIComponent('Tier A') + '&fields=id,name,lastActivity', null, 'get').people || [];
      return say(ppl.length ? 'your hottest leads:\n' + ppl.map(x => '• ' + x.name).join('\n') : 'no Tier A leads right now.');
    }
    if (a.action === 'texts') {
      const log = JSON.parse(props.getProperty('TEXT_LOG') || '[]').slice(-8).reverse();
      return say(log.length ? 'latest texts:\n' + log.map(x => '• ' + x.who + ' (' + Utilities.formatDate(new Date(x.at), 'America/Los_Angeles', 'EEE h:mm a') + '): ' +
        (x.reply ? 'I answered "' + x.reply.slice(0, 80) + '"' : '"' + x.text.slice(0, 80) + '"')).join('\n') : 'no client texts in the last few days.');
    }
    if (a.action === 'read') {
      let names = String(a.who || '').split(/\s*(?:,|&|\band\b)\s*/i).map(x => x.trim()).filter(Boolean).slice(0, 4);
      if (!names.length && props.getProperty('LAST_ALERT_WHO')) names = [props.getProperty('LAST_ALERT_WHO')];  // "what does the email say?"
      if (!names.length) return say('whose messages should I read you?');
      const blocks = [];
      names.forEach(nm => {
        const p = findOne(nm);
        // Two people with the same first name (e.g. two Ruths): read each and keep the ones with recent activity.
        const cands = p ? (p.many ? p.many.slice(0, 4) : [p]) : [];
        const found = cands.map(c => ({who: c.name, tl: personTimeline(c, '')})).filter(b => b.tl);
        if (found.length) { found.forEach(b => blocks.push(b)); return; }
        const c = macContact(nm), phone = c ? bbE164(c.phone) : '';
        blocks.push({who: (c && c.name) || nm, tl: phone ? personTimeline(null, phone) : ''});
      });
      const mv = mailerVersion(), mailNote = mv < 4 ? '\n\n(Heads up: I can\'t read your bcolv emails yet. ' + mailerProblem(mv) + ')' : '';
      if (blocks.every(b => !b.tl)) return say('I don\'t see texts, emails or calls with ' + names.join(', ') + ' in the last 2 weeks.' + mailNote);
      let answer = '';
      try {
        const r = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
          {role: 'system', content: 'You are ' + name + ', Julián\'s assistant. For each person below, tell Julián in 1 to 3 short lines what they most recently said or asked (quote short bits), across texts, emails and calls, with the day, and whether they\'re still waiting on a reply. Plain text, no dashes. Return JSON {"summary": "..."} with one bullet (•) per person.'},
          {role: 'user', content: blocks.map(b => '### ' + b.who + '\n' + (b.tl || 'nothing in the last 2 weeks')).join('\n\n').slice(0, 8000)}]));
        answer = humanize(r.summary || '');
      } catch (e) {}
      if (!answer) answer = blocks.map(b => '• ' + b.who + '\n' + (b.tl ? b.tl.split('\n').filter(l => /Them|emailed|Call with/.test(l)).slice(-3).join('\n') : 'nothing in the last 2 weeks')).join('\n\n');
      return say(answer + mailNote + '\n\nWant me to answer any of them? Just tell me what to say.');
    }
    if (a.action === 'unanswered') {
      const seen = {}, lines = [];
      JSON.parse(props.getProperty('TEXT_LOG') || '[]').filter(x => Date.now() - x.at < 7 * 864e5 && !x.read && !isOwnNumber(x.phone)).reverse().forEach(x => {
        const ph = x.phone || '', k = ph || x.who;
        if (seen[k] || lines.length >= 8) return; seen[k] = 1;
        let who = x.who;
        if (ph) {
          const th = bbThread(ph).split('\n').filter(Boolean), last = th[th.length - 1] || '';
          if (/Julián: /.test(last)) return;  // someone (Julián or Shawna) already answered
          const p = (fubFetch('people?limit=1&fields=id,firstName,lastName,tags&phone=' + ph.slice(2), null, 'get').people || [])[0];
          if (p && (p.tags || []).some(t => /^Test\b/i.test(t))) return;
          who = p ? ([realFirst(p), p.lastName].join(' ').trim() || ph) : ph;
        } else if (/^no name$/i.test(String(who))) return;
        lines.push('• ' + who + ' (text): "' + String(x.text || '').slice(0, 100) + '"');
      });
      const r = bcolvCall({action: 'inbox', days: 7, max: 10});
      ((r && r.emails) || []).slice(0, 6).forEach(m => lines.push('• ' + String(m.from).replace(/<.*>/, '').replace(/"/g, '').trim() + ' (email "' + m.subject + '"): ' + String(m.body || '').replace(/\s+/g, ' ').slice(0, 100)));
      return say(lines.length ? 'still waiting on you:\n' + lines.join('\n') + '\n\nTell me what to say and I\'ll reply for you.' : 'you\'re all caught up. No unanswered texts or unread client emails this week.');
    }
    if (a.action === 'agenda') return say(todoText(agendaRange(String(a.message || '') + ' ' + text)));
    if (a.action === 'unschedule') {
    const sh = getSS().getSheetByName(TEXT_QUEUE_TAB.name);
    if (!sh || sh.getLastRow() < 2) return say('nothing is scheduled.');
    const want = String(a.who || '').toLowerCase().split(' ')[0];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    let n = 0;
    rows.forEach((r, i) => { if (r[3] === 'scheduled' && r[4] === 'pending' && (!want || String(r[2]).toLowerCase().indexOf(want) >= 0)) { sh.getRange(i + 2, 5).setValue('canceled by Julián'); n++; } });
    return say(n ? 'canceled ' + n + ' scheduled text' + (n > 1 ? 's' : '') + (a.who ? ' to ' + a.who : '') + '. ✅' : 'I don\'t see a scheduled text' + (a.who ? ' to ' + a.who : '') + '.');
  }
  if (a.action === 'taskdone' || a.action === 'taskstart') {
      const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
      let saved = JSON.parse(props.getProperty('AGENDA_TASKS') || '[]');
      const arg = String(a.message || '').toLowerCase();
      const asked = ((arg.match(/\d+/g) || (/^\s*(done|do|finish|start|complete)\b/i.test(text) ? text.match(/\d+/g) : null)) || []).map(Number).filter(n => n >= 1 && n < 50);
      // Still check each task is open; a number that's not on the list he's looking at → show him today's real list.
      const stillOpen = (fubFetch('tasks?assignedUserId=' + me + '&isCompleted=false&limit=100&sort=dueDate&fields=id,name,dueDate,personId', null, 'get').tasks || []);
      const openIds = stillOpen.map(t => t.id);
      saved = saved.filter(t => openIds.indexOf(t.id) >= 0 || asked.length);
      if (asked.length && asked.some(n => n > saved.length || openIds.indexOf((saved[n - 1] || {}).id) < 0)) {
        if (!stillOpen.length) return say('you\'re all clear: no open tasks in Follow Up Boss. ✅');
        const fresh = stillOpen.slice(0, 10).map(t => ({id: t.id, personId: t.personId, due: String(t.dueDate || '').slice(0, 10), name: String(t.name || '').slice(0, 200), who: ''}));
        fresh.forEach(t => { if (t.personId) { try { const pp = fubFetch('people/' + t.personId + '?fields=id,name,firstName,lastName', null, 'get'); t.who = [realFirst(pp), pp.lastName].join(' ').trim() || pp.name; } catch (e) {} } });
        props.setProperty('AGENDA_TASKS', JSON.stringify(fresh));
        return say('that number isn\'t on your list anymore (it may already be done). Here\'s what\'s open now:\n' +
          fresh.map((t, i) => (i + 1) + '. ' + (t.who ? t.who + ': ' : '') + t.name.replace(/^\w+: /, '').slice(0, 70) + (t.due ? ' (' + t.due.slice(5) + ')' : '')).join('\n') + '\n\nText "done 1" (or the right number).');
      }
      const nums = asked.filter(n => n <= saved.length);
      let picks = nums.map(n => saved[n - 1]);
      if (!picks.length && a.who) picks = saved.filter(t => String(t.who + ' ' + t.name).toLowerCase().indexOf(String(a.who).toLowerCase().split(' ')[0]) >= 0).slice(0, 1);
      if (a.action === 'taskdone') {
        const today0 = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
        if (!picks.length && /old|all|past|everything/.test(arg)) {  // "clear all": everything due today or earlier, never future ones
          const open = fubFetch('tasks?assignedUserId=' + me + '&isCompleted=false&limit=100&fields=id,dueDate', null, 'get').tasks || [];
          picks = open.filter(t => String(t.dueDate || '').slice(0, 10) <= today0);
        }
        const future = picks.filter(t => t.due && t.due > today0);
        if (future.length > 1) picks = picks.filter(t => !(t.due && t.due > today0)).concat(future.slice(0, 1));  // future ones: one at a time
        if (!picks.length && a.who) {  // "I already called Dan" / "Lisa's handled": that person's open tasks due today or earlier (else their next one)
          const pw = findOne(a.who);
          if (pw && pw.many) return say('which ' + a.who + '? ' + pw.many.map(x => x.name).join(', '));
          if (pw) {
            const open = (fubFetch('tasks?personId=' + pw.id + '&isCompleted=false&limit=20&sort=dueDate&fields=id,name,dueDate', null, 'get').tasks || []).map(t => ({id: t.id, name: t.name, due: String(t.dueDate || '').slice(0, 10)}));
            picks = open.filter(t => t.due <= today0);
            if (!picks.length && open.length) picks = open.slice(0, 1);
            if (!picks.length) return say('I don\'t see an open task for ' + pw.name + '. Nothing to check off.');
          }
        }
        if (!picks.length) {  // "done" / "that's done" right after she started one
          const lt = JSON.parse(props.getProperty('LAST_TASK') || 'null');
          if (lt && Date.now() - lt.at < 12 * 36e5) picks = [{id: lt.id}];
        }
        if (!picks.length) return say('which ones? Text "done" with the numbers from my list, like "done 2" (or "clear all" for everything past due).');
        let n = 0; picks.forEach(t => { try { fubFetch('tasks/' + t.id, {isCompleted: true}, 'put'); n++; } catch (e) {} });
        return say('checked off ' + n + ' task' + (n === 1 ? '' : 's') + '. ✅');
      }
      const t = picks[0];
      if (t) props.setProperty('LAST_TASK', JSON.stringify({id: t.id, at: Date.now()}));
      if (!t) return say('which one? Text "do" with the number from my list, like "do 2". Ask "what do I need to do today" to see the list.');
      if (!t.personId) return say('that one isn\'t tied to a person: "' + t.name + '". Text "done ' + (saved.indexOf(t) + 1) + '" when it\'s handled.');
      const pp = fubFetch('people/' + t.personId + '?fields=id,name,firstName,lastName,tags,phones,emails', null, 'get'), who = [realFirst(pp), pp.lastName].join(' ').trim() || pp.name;
      const phone = (pp.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0], tl = personTimeline(pp, phone);
      const isCall = /^call/i.test(t.name);
      const r = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
        {role: 'system', content: assistantPrompt(name, 'julian') + '\nJulián wants to handle this to-do now: "' + t.name + '". Using everything below, write ' + (isCall ? 'a 2-line briefing for his call (who, what they want, what to say first) in "reply", plus a short text he could send instead in "text".' : 'the text he should send them now in "reply" (his voice, picks up exactly where things left off).') + ' Return JSON {"reply": "...", "text": "..."}.'},
        {role: 'user', content: 'Contact: ' + who + '\n' + (tl || 'no history') }]));
      if (isCall) {
        const alt = humanize(r.text || '');
        if (alt && phone && (pp.tags || []).indexOf('Text Opt-Out') < 0) addPending({personId: pp.id, who: who, phone: phone, message: alt, at: Date.now(), taskId: t.id});
        return say('📞 ' + who + (phone ? ' ' + phone.slice(2).replace(/(\d{3})(\d{3})(\d{4})/, '($1) $2-$3') : '') + '\n' + humanize(r.reply || '') +
          (alt && phone ? '\n\nOr I can text them:\n"' + alt + '"\nReply YES to send it.' : '') + '\n\nText "done ' + (saved.indexOf(t) + 1) + '" after you talk.');
      }
      const msg = humanize(r.reply || '');
      if (!phone || !msg) return say('I couldn\'t draft that one (no number for ' + who + '). Text "done ' + (saved.indexOf(t) + 1) + '" when it\'s handled.');
      addPending({personId: pp.id, who: who, phone: phone, message: msg, at: Date.now(), taskId: t.id});
      return say('here\'s what I\'d text ' + who + ':\n\n"' + msg + '"\n\nReply YES to send it (I\'ll check the task off), or tell me what to change.');
    }
    if (a.action === 'ask') {
      const question = a.message || text, ctx = [], clip = (x, n) => String(x || '').replace(/\s+/g, ' ').slice(0, n);
      String(a.who || '').split(/\s*(?:,|&|\band\b)\s*/i).map(x => x.trim()).filter(Boolean).slice(0, 2).forEach(nm => {
        const p = findOne(nm); let phone = '';
        if (!p || p.many) { const c = macContact(nm); if (c) phone = bbE164(c.phone); }
        const tl = personTimeline(p && !p.many ? p : null, phone);
        if (tl) ctx.push('### EVERYTHING WITH ' + ((p && !p.many && p.name) || nm).toUpperCase() + '\n' + tl);
      });
      try {  // Follow Up Boss: calls Shawna took, texts, emails sent, notes (all contacts), newest first
        const notes = (fubFetch('notes?limit=40&sort=-created', null, 'get').notes || []).filter(n => !/^(send text|send bcolv email|claude request|shawna allow)/i.test(String(n.subject || '')));
        const names = {};
        notes.forEach(n => { if (!(n.personId in names) && Object.keys(names).length < 20) { try { const pp = fubFetch('people/' + n.personId + '?fields=id,name,firstName,lastName', null, 'get'); names[n.personId] = [realFirst(pp), pp.lastName].join(' ').trim() || pp.name; } catch (e) { names[n.personId] = '?'; } } });
        ctx.push('FOLLOW UP BOSS ACTIVITY (calls, texts, emails, notes; newest first):\n' + notes.map(n => String(n.created).slice(5, 16).replace('T', ' ') + ' ' + (names[n.personId] || 'contact') + ' | ' + clip(n.subject, 60) + ': ' + clip(n.body, 220)).join('\n'));
        const fresh = fubFetch('people?limit=12&sort=-created&fields=id,name,source,tags,created,stage', null, 'get').people || [];
        ctx.push('NEWEST CONTACTS (website, open house, Instagram, calls, texts):\n' + fresh.map(x => String(x.created).slice(5, 10) + ' ' + x.name + ' | source ' + (x.source || '') + ' | ' + (x.tags || []).slice(0, 6).join(', ')).join('\n'));
      } catch (e) { ctx.push('(Follow Up Boss unavailable: ' + e.message + ')'); }
      try {  // texts on the Mac: latest conversations
        const chats = bbApi('chat/query', {limit: 12, offset: 0, with: ['lastMessage'], sort: 'lastmessage'}).data || [];
        ctx.push('LATEST TEXT CONVERSATIONS (his phone):\n' + chats.filter(c => c.lastMessage && c.lastMessage.text).map(c => (c.displayName || c.chatIdentifier || '') + ' | ' +
          (c.lastMessage.isFromMe ? 'Julián/Shawna: ' : 'Them: ') + clip(c.lastMessage.text, 160)).join('\n'));
      } catch (e) {}
      try {  // email (julianc@bcolv.com)
        const r = bcolvCall({action: 'search', query: a.query || 'newer_than:7d -category:promotions -category:social -category:updates', max: 12});
        if (r && r.ok) ctx.push('EMAIL (julianc@bcolv.com, newest first):\n' + (r.messages || []).sort((x, y) => y.date - x.date).slice(0, 20).map(m =>
          Utilities.formatDate(new Date(m.date), 'America/Los_Angeles', 'M/d h:mma') + ' ' + (m.fromMe ? 'Julián to ' + clip(m.to, 40) : 'from ' + clip(m.from, 50)) + ' | "' + clip(m.subject, 70) + '": ' + clip(m.body, 260)).join('\n'));
      } catch (e) {}
      try {  // website and open house forms
        ['Website Leads', 'Sign-Ins'].forEach(t => {
          const sh = getSS().getSheetByName(t); if (!sh || sh.getLastRow() < 2) return;
          const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], n = Math.min(8, sh.getLastRow() - 1);
          const rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, sh.getLastColumn()).getValues().reverse();
          ctx.push(t.toUpperCase() + ' (website forms, newest first):\n' + rows.map(r => r.map((v, i) => v === '' ? '' : head[i] + ': ' + clip(v instanceof Date ? Utilities.formatDate(v, 'America/Los_Angeles', 'M/d h:mma') : v, 80)).filter(Boolean).join(' | ')).join('\n'));
        });
      } catch (e) {}
      try {  // calendar, next 7 days
        const now = new Date(), end = new Date(now.getTime() + 7 * 864e5), evs = [];
        CalendarApp.getDefaultCalendar().getEvents(now, end).forEach(e => evs.push({t: e.getTitle(), s: e.getStartTime().getTime(), w: 'personal'}));
        ((bcolvCall({action: 'events', from: now.toISOString(), to: end.toISOString()}) || {}).events || []).forEach(e => evs.push({t: e.title, s: e.start, w: 'business'}));
        ctx.push('CALENDAR (next 7 days):\n' + evs.sort((x, y) => x.s - y.s).map(x => Utilities.formatDate(new Date(x.s), 'America/Los_Angeles', 'EEE M/d h:mma') + ' ' + x.t + ' (' + x.w + ')').join('\n'));
      } catch (e) {}
      try {
        const r = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
          {role: 'system', content: 'You are ' + name + ', Julián\'s assistant. Today is ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE M/d h:mma') + ' in Las Vegas. Answer his question using ONLY the data below (Follow Up Boss including calls you took on the phone, his texts, his bcolv email, website and open house forms, his calendar). Be specific: names, days, short quotes, numbers. If the data doesn\'t answer it, say so and say what you checked. Short, plain, friendly, no dashes. Return JSON {"answer": "..."}.'},
          {role: 'user', content: 'His question: "' + question + '"\n\n' + ctx.join('\n\n').slice(0, 12000)}]));
        return say(humanize(r.answer || 'I checked everything and came up empty on that one.'));
      } catch (e) { return say('I couldn\'t pull that together just now (' + String(e.message).slice(0, 80) + '). Try again in a minute?'); }
    }
    if (a.action === 'allow') {
      const names = String(a.who || '').split(/\s*(?:,|&|\band\b)\s*/i).map(x => x.trim()).filter(Boolean);
      const r = allowShawna(names);
      return say((r.ok.length ? 'done. I can now text and answer ' + r.ok.join(', ') + '. ✅' : '') + (r.missing.length ? (r.ok.length ? '\n' : '') + 'I couldn\'t find a number for ' + r.missing.join(', ') + '. Send me their numbers.' : ''));
    }
    if (a.action === 'alerts') {
      const on = /on|yes|start/i.test(String(a.message || text)) && !/off|stop|no\b/i.test(String(a.message || ''));
      props.setProperty('TEXT_ALERTS', on ? 'on' : 'off');
      return say(on ? 'text alerts are on. I\'ll tell you when clients text.' : 'text alerts are off. I\'ll handle texts quietly. Ask me "who texted me?" anytime.');
    }
    if (a.action === 'jot') return say(jotNote(a.folder, a.message || text));
    if (a.action === 'notes') return say(readNotes(a.folder, a.query || (/\?$|what|which|any|when|how/i.test(text) ? text : '')));
    if (a.action === 'remember') {
      const ss = getSS(), sh = ss.getSheetByName('Shawna Knowledge') || ensureTab(ss, {name: 'Shawna Knowledge', color: '#6B4FBB', headers: ['One fact per row (Shawna reads these)']});
      sh.appendRow([cell(a.message || text)]);
      return say('got it, I\'ll remember that. ✅');
    }
    if (a.action === 'callme') {
      const r = vapiCall(boss, a.message || '');
      return say(r.ok ? 'calling you now 📞' : 'I couldn\'t start the call: ' + r.error);
    }
    if (a.action === 'emails') {  // who wrote to him and what they said: one line each, no system mail, nothing he sent
      let msgs = [];
      try {
        const r = bcolvCall({action: 'search', query: 'in:inbox newer_than:' + (/week|7/i.test(text) ? '7d' : '2d') + ' -from:me', max: 15});
        msgs = ((r && r.messages) || []).filter(x => !x.fromMe && !/mailer-daemon|postmaster|no-?reply|notification|followupboss|google\.com|calendar/i.test(x.from) &&
          !/delivery status|undeliverable|auto.?reply|out of office/i.test(x.subject || ''));
      } catch (e) {}
      const seen = {}; msgs = msgs.sort((x, y) => y.date - x.date).filter(x => { const k = String(x.from).toLowerCase(); if (seen[k]) return false; seen[k] = 1; return true; }).slice(0, 8);
      if (!msgs.length) return say('no new emails from people in the last ' + (/week|7/i.test(text) ? 'week' : '2 days') + '.');
      let lines = msgs.map(x => ({who: String(x.from).replace(/<.*>/, '').replace(/"/g, '').trim() || x.from, at: x.date, body: String(x.body || '').replace(/\s+/g, ' ').slice(0, 500)}));
      try {
        const sum = JSON.parse(groqChat(key, props.getProperty('GROQ_ROUTER_MODEL') || 'openai/gpt-oss-20b', [
          {role: 'system', content: 'For each email, write what the person said in at most 12 plain words (no dashes). Return JSON {"s": ["...", "..."]} in the same order.'},
          {role: 'user', content: lines.map((l, i) => (i + 1) + '. ' + l.who + ': ' + l.body).join('\n')}], {maxTokens: 600, noPick: true})).s || [];
        lines = lines.map((l, i) => Object.assign(l, {said: sum[i] || ''}));
      } catch (e) {}
      return say('📧 emails:\n' + lines.map(l => '• ' + l.who + ' (' + Utilities.formatDate(new Date(l.at), 'America/Los_Angeles', 'EEE h:mm a') + '): ' + (l.said || l.body.slice(0, 90))).join('\n'));
    }
    if (a.action === 'calendar') return say(calendarText(agendaRange(String(a.message || '') + ' ' + text)));
    if (a.action === 'callnotes') {
      const d = a.data || {};
      if (!a.who && !a.phone) {
        props.setProperty('CALLNOTES_WAIT', JSON.stringify({key: a.key, at: Date.now()}));
        return say('📞 got the call notes. Who was the call with? (just the name)');
      }
      const {p, isNew} = resolvePerson();
      if (!p || p.many) {
        props.setProperty('CALLNOTES_WAIT', JSON.stringify({key: a.key, at: Date.now()}));
        return say(p && p.many ? 'which ' + a.who + '? ' + p.many.map(x => x.name).join(', ') + ' (reply with the full name)' : 'who was the call with? (just the name)');
      }
      const who = p.name || a.who;
      const body = (d.summary ? d.summary + '\n\n' : '') + (d.details ? 'Details:\n' + d.details + '\n\n' : '') + (d.next ? 'Next step: ' + d.next + '\n\n' : '') +
        'Notes / transcript (from Julián\'s iPhone call recording):\n' + String(a.full || '').slice(0, 6000);
      const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0] || bbE164(a.phone || '');
      let logged = false;
      try { fubFetch('calls', {personId: p.id, phone: phone ? phone.slice(2) : '', isIncoming: /\b(called me|inbound|incoming|they called)\b/i.test(a.full || ''), outcome: 'Interested', note: ((d.summary || '') + (d.next ? '\nNext step: ' + d.next : '') || 'Call notes saved as a note.').slice(0, 2000)}); logged = true; } catch (e) {}
      fubFetch('notes', {personId: p.id, subject: 'Call notes' + (d.next ? ': ' + String(d.next).slice(0, 80) : ''), body: body.slice(0, 9000)});
      const out = ['📞 logged your call with ' + who + (isNew ? ' (new contact)' : '') + '.'];
      if (d.summary) out.push(humanize(d.summary));
      const fu = /^\d{4}-\d{2}-\d{2}$/.test(d.followup || '') ? d.followup : Utilities.formatDate(new Date(Date.now() + 2 * 864e5), 'America/Los_Angeles', 'yyyy-MM-dd');
      try {
        fubFetch('tasks', {personId: p.id, type: /^(Call|Text|Email|Showing)$/.test(d.followupType || '') ? d.followupType : 'Follow Up', name: String(d.next || 'Follow up from the call').slice(0, 250), dueDate: fu});
        out.push('✅ task: ' + (d.next || 'follow up') + ' (' + Utilities.formatDate(Utilities.parseDate(fu, 'America/Los_Angeles', 'yyyy-MM-dd'), 'America/Los_Angeles', 'EEE MMM d') + ')');
      } catch (e) {}
      try { const t = applyShawnaTags(p.id, d.tags, 'a phone call'); if (t.length) out.push('🏷️ tagged ' + t.join(', ')); } catch (e) {}
      try { applyIntent(p.id, String(d.intent || ''), 'a phone call'); } catch (e) {}
      try { if (DEAL_SIGNALS[d.deal]) suggestDeal(p.id, d.deal, 'from your call'); } catch (e) {}
      const optOut = (p.tags || []).indexOf('Text Opt-Out') >= 0;
      if (d.text && phone && !optOut) {
        const n = addPending({personId: p.id, who: who, phone: phone, message: humanize(d.text), at: Date.now(), long: true});
        out.push('\n💬 follow-up text (' + yesHint(n) + '):\n"' + humanize(d.text) + '"');
      }
      let email = a.email;
      if (!email) { try { const full = fubFetch('people/' + p.id + '?fields=emails', null, 'get'); email = ((full.emails || []).filter(e => !/invalid|bounced/i.test(e.status || ''))[0] || {}).value; } catch (e) {} }
      if (d.emailBody && email && props.getProperty('BCOLV_MAILER_URL')) {
        const n = addPending({kind: 'email', personId: p.id, who: who, email: email, subject: humanize(d.emailSubject || 'Great talking with you'), message: humanize(d.emailBody), at: Date.now()});
        out.push('\n📧 recap email (' + yesHint(n) + '):\nSubject: ' + humanize(d.emailSubject || 'Great talking with you') + '\n' + humanize(d.emailBody));
      }
      if (pendingQueue().length > 1) out.push('\nReply YES 1, YES 2… for each one, or tell me what to change.');
      return say(out.join('\n'));
    }
    if (a.action === 'tag') {
      const p = findOne(a.who);
      if (!p) return say('I couldn\'t find "' + (a.who || '?') + '" in Follow Up Boss.');
      if (p.many) return say('which one? ' + p.many.map(x => x.name).join(', '));
      const m = String(a.message || ''), cut = m.split(/remove:/i);
      const clean = x => String(x || '').split(/\s*,\s*/).map(t => t.trim()).filter(Boolean);
      const add = clean(cut[0]), remove = clean(cut[1]).map(t => t.toLowerCase());
      const full = fubFetch('people/' + p.id + '?fields=id,name,tags', null, 'get');
      const tags = (full.tags || []).filter(t => remove.indexOf(t.toLowerCase()) < 0);
      add.forEach(t => { if (!tags.some(x => x.toLowerCase() === t.toLowerCase())) tags.push(t); });
      fubFetch('people/' + p.id, {tags: tags}, 'put');
      fubFetch('notes', {personId: p.id, subject: 'Tags updated', body: (add.length ? 'Added: ' + add.join(', ') : '') + (remove.length ? (add.length ? '\n' : '') + 'Removed: ' + clean(cut[1]).join(', ') : '') + '\n(by ' + name + ' at Julián\'s request)'});
      return say('done ✅ ' + full.name + (add.length ? ' tagged ' + add.join(', ') : '') + (remove.length ? (add.length ? ';' : '') + ' removed ' + clean(cut[1]).join(', ') : '') + '.');
    }
    if (a.action === 'research') {
      const q = String(a.message || text);
      const info = liveInfo(q, true);
      let ans = '';
      try {
        ans = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
          {role: 'system', content: 'You are ' + name + ', Julián\'s AI assistant. He is a Las Vegas Realtor. Answer his question like a sharp, well read friend who knows real estate and lending: 1 to 4 short sentences, plain text, no dashes. Use the LIVE INFO for anything current and say where numbers come from; never invent current numbers or news. Return JSON {"reply": "..."}.'},
          {role: 'user', content: q + info}], {maxTokens: 700})).reply || '';
      } catch (e) { ans = info.replace(/^\s*LIVE INFO[^\n]*\n/, ''); }
      return say(ans || 'I couldn\'t look that up right now. Ask me again in a minute.');
    }
    if (a.action === 'claude') {
      const sh = getSS().getSheetByName(REQUESTS_TAB.name) || ensureTab(getSS(), REQUESTS_TAB);
      sh.appendRow([Utilities.formatDate(new Date(), 'America/Los_Angeles', 'M/d h:mm a'), cell(a.message || text), 'new']);
      // Also a FUB note on Julián's own test contact, where Claude's hourly check picks it up (and texts him back there).
      fubFetch('notes', {personId: Number(props.getProperty('REQUESTS_PERSON_ID') || 2221), subject: 'Claude request', body: a.message || text});
      return say(name + ': got it ✅ "' + (a.message || text).slice(0, 120) + '". Claude checks every hour and will text you when it\'s done.');
    }
    return say(name + ': ' + (a.reply || 'Got it.'));
  };
  const list = Array.isArray(a.actions) && a.actions.length ? a.actions.filter(x => x && x.action).slice(0, 5) : [a];
  if (list.length <= 1) return doAction(list[0] || a);
  list.forEach((x, i) => {
    try { doAction(x); } catch (e) { say('I couldn\'t finish #' + (i + 1) + ' (' + String(e.message).slice(0, 80) + '). Ask me that one again?'); }
    Utilities.sleep(800);
  });
}

// CALL NOTES (Julián, Nov 11): he records a call on his iPhone (business line 702-498-2571), copies the transcript or
// summary from the Notes app and texts it to Shawna ("call notes Lisa" + paste, or the "Send call to Shawna" shortcut).
// Shawna logs the call in Follow Up Boss, tags, sets the follow-up task, suggests a deal move, and drafts the follow-up
// text and email for his YES. Long transcripts are kept whole in the cache (6 h); his message carries a [callnotes:key].
const CALL_NOTES_RE = /^\s*(?:📞\s*)?(?:call\s*notes?|call\s+(?:summary|recap|transcript)|notes?\s+(?:from|on)\s+(?:my\s+)?call|transcript)\b/i;
function stashCallNotes(text) {
  if (!CALL_NOTES_RE.test(text) && !(text.length > 1400 && /transcript|summary|call/i.test(text.slice(0, 400)))) return text.slice(0, 1500);
  const key = 'CN' + Date.now();
  CacheService.getScriptCache().put(key, text.slice(0, 90000), 6 * 3600);
  return text.slice(0, 160).replace(/\s+/g, ' ') + ' [callnotes:' + key + ']';
}
function callNotesParse(full) {
  const props = PropertiesService.getScriptProperties();
  const raw = groqChat(props.getProperty('GROQ_API_KEY'), props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
    {role: 'system', content: 'Julián Crider is a Las Vegas Realtor. He just had a phone call and pasted his notes, the call summary or the transcript (the first line may say who it was with). Today is ' +
      Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE yyyy-MM-dd') + '. Return JSON only:\n' +
      '{"who": "the other person\'s full name (as Julián wrote it, else from the call; empty if unknown)", "phone": "their phone if said", "email": "their email if said", ' +
      '"summary": "2 to 4 short lines: who they are and what was discussed", "details": "key facts, one per line: budget, areas, beds, timeline, pre-approval, must haves, their situation, anything personal worth remembering", ' +
      '"next": "the next step that was agreed", "followup": "YYYY-MM-DD when Julián should follow up (what they agreed, else 2 days from today)", "followupType": "Call, Text, Email, Showing or Follow Up", ' +
      '"tags": "comma separated, only if clear: moved_away, agent, referral_partner, relocation, seller, investor, has_agent, bought_elsewhere, renter, past_client", ' +
      '"intent": "not_now, not_interested or empty", "deal": "offer, under_contract, closed, listing or empty", ' +
      '"text": "a short follow-up text from Julián in his voice: thank them, recap one specific thing they said, the next step. Warm, human, light Chris Voss, 1 emoji at most, no dashes", ' +
      '"emailSubject": "short subject", "emailBody": "a short recap email (3 to 6 sentences, plain paragraphs, no signature) if the call had details worth putting in writing, else empty"}'},
    {role: 'user', content: String(full).slice(0, 14000)}], {maxTokens: 1400});
  try { return JSON.parse(raw); } catch (e) { const j = String(raw).match(/\{[\s\S]*\}/); return j ? JSON.parse(j[0]) : {}; }
}

// SHAWNA'S FOLDERS (Julián, Nov 12): his personal notebooks, kept as tabs in the JC Open House Leads sheet so he can
// open them in the Google Sheets app. Golf (swing thoughts, his game), Links (web links he sends), SocialPod, Personal.
// Anything else he names ("put this in my recipes folder") gets its own new folder tab.
const NOTE_FOLDERS = {
  golf: {name: 'Golf Notes', color: '#188038', headers: ['When', 'Note']},
  links: {name: 'Links', color: '#1A73E8', headers: ['When', 'Link', 'Title', 'Note']},
  socialpod: {name: 'SocialPod Notes', color: '#9334E6', headers: ['When', 'Note']},
  personal: {name: 'Personal Notes', color: '#5F6368', headers: ['When', 'Note']},
  meetings: {name: 'Meeting Notes', color: '#0B8043', headers: ['When', 'Note']},
  listings: {name: 'Listings', color: '#B3261E', headers: ['When', 'Address', 'Status', 'Notes']}
};
function noteFolder(folder) {
  const k = String(folder || 'personal').toLowerCase().replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim()
    .replace(/^social ?pod.*/, 'socialpod').replace(/^(golf|swing).*/, 'golf').replace(/^(links?|websites?|urls?)$/, 'links').replace(/^(personal|me|self|notes?)$/, 'personal')
    .replace(/^(meetings?|meeting notes?|team meetings?|joey).*/, 'meetings').replace(/^(listings?|properties|inventory|coming soon|renovat\w*).*/, 'listings');
  if (NOTE_FOLDERS[k]) return NOTE_FOLDERS[k];
  const nm = k.replace(/\b\w/g, c => c.toUpperCase()).slice(0, 30) + ' Notes';
  return {name: nm, color: '#E37400', headers: ['When', 'Note']};
}
/** Saves a note (or a link) into one of his folders. Returns a short confirmation. */
function jotNote(folder, text) {
  const f = noteFolder(folder), sh = ensureTab(getSS(), f), when = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'M/d/yyyy h:mm a');
  if (f.name === 'Listings') {  // "123 Main St coming soon", "456 Oak renovating", "123 Main is active now"
    const st = /renovat|rehab|fixing|construction|remodel/i.test(text) ? 'Renovating' : /coming|soon|upcoming|available soon|pocket|off.?market/i.test(text) ? 'Coming soon'
      : /\bactive\b|listed|went live|on the market/i.test(text) ? 'Active' : /pending|under contract|escrow/i.test(text) ? 'Pending' : /sold|closed/i.test(text) ? 'Sold' : 'Coming soon';
    const addr = (String(text).match(/\d{2,6}\s+[A-Za-z0-9 .'#-]{3,60}?(?=,|\s+(?:is|are|was|coming|renovat\w*|active|pending|sold|listed|status|in|by|for|done|will|should|needs?|going)\b|$)/i) || [''])[0].trim();
    if (addr && sh.getLastRow() > 1) {  // already on the list: update its status
      const vals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < vals.length; i++) if (vals[i][0] && vals[i][0].toLowerCase().indexOf(addr.toLowerCase().split(' ').slice(0, 2).join(' ')) === 0) {
        sh.getRange(i + 2, 3).setValue(st); sh.getRange(i + 2, 4).setValue(cell(String(text).trim())); sh.getRange(i + 2, 1).setValue(when);
        return 'updated ' + vals[i][0] + ' → ' + st + ' 🏠';
      }
    }
    sh.appendRow([when, cell(addr || String(text).trim().slice(0, 80)), st, cell(String(text).trim())]);
    return 'added to Listings (' + st + ') 🏠';
  }
  if (f.name === 'Links') {
    const urls = String(text).match(/https?:\/\/[^\s<>"]+/g) || [];
    const note = String(text).replace(/https?:\/\/[^\s<>"]+/g, '').replace(/^\s*(save|saved|link|links?)\s*(this|that)?\s*(link)?\s*[:\-]?\s*/i, '').trim();
    urls.forEach(u => {
      let title = '';
      try { const h = UrlFetchApp.fetch(u, {muteHttpExceptions: true, followRedirects: true, headers: {'User-Agent': 'Mozilla/5.0 (Macintosh) Safari/605'}}).getContentText().slice(0, 100000);
        title = ((h.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)/i) || h.match(/<title[^>]*>([^<]+)/i) || [])[1] || '').replace(/&amp;/g, '&').replace(/&#39;|&rsquo;/g, "'").trim().slice(0, 150); } catch (e) {}
      sh.appendRow([when, u, cell(title), cell(note)]);
    });
    return urls.length ? 'saved to Links 🔗' + (urls.length > 1 ? ' (' + urls.length + ' links)' : '') : 'I didn\'t see a link in that one.';
  }
  sh.appendRow([when, cell(String(text).trim())]);
  return 'saved to ' + f.name.replace(/ Notes$/, '') + ' 📝';
}
/** Reads a folder back: the latest notes, or an answer to his question from what's in it. */
function readNotes(folder, question) {
  const f = noteFolder(folder), sh = getSS().getSheetByName(f.name);
  if (!sh || sh.getLastRow() < 2) return 'your ' + f.name.replace(/ Notes$/, '') + ' folder is empty so far.';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, f.headers.length).getDisplayValues();
  const lines = rows.map(r => f.name === 'Links' ? r[0] + ' ' + r[1] + (r[2] ? ' (' + r[2] + ')' : '') + (r[3] ? ': ' + r[3] : '') : r[0] + ': ' + r[1]);
  const key = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (question && key) {
    try {
      return JSON.parse(groqChat(key, PropertiesService.getScriptProperties().getProperty('GROQ_MODEL') || ASSISTANT.model, [
        {role: 'system', content: 'You are Shawna, Julián\'s personal assistant. Answer his question using only his saved notes below (newest last). Short and plain, no dashes; list the matching notes or links if that helps. Return JSON {"reply": "..."}.'},
        {role: 'user', content: 'His ' + f.name + ':\n' + lines.join('\n').slice(-9000) + '\n\nHis question: ' + question}], {maxTokens: 700})).reply || '';
    } catch (e) {}
  }
  return f.name + ' (latest):\n' + lines.slice(-8).map(l => '• ' + l.slice(0, 160)).join('\n');
}

/** Runs one of Julián's texts to Shawna with a log and a safety net: if it fails (Mac busy, Groq hiccup), it's queued
 *  and retried by processAssistantQueue a minute later (up to 3 tries), and he's told if it still fails. */
function runBossCommand(text, tries) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {}  // one of his messages at a time, so they don't trip over each other
  try {
    var logRow = shawnaLog('in', text, tries ? 'retry ' + tries : '');
    bossCommand(text, tries);
    shawnaLogDone(logRow, 'ok', text, tries ? 'retry ' + tries : '');
  } catch (err) {
    shawnaLogDone(logRow, 'error', text, err.message);
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
    if (sh.getLastRow() > 600) sh.deleteRows(2, 100);
    sh.appendRow([Utilities.formatDate(new Date(), 'America/Los_Angeles', 'M/d h:mm:ss a'), kind, cell(text.slice(0, 300)), cell(String(detail || '').slice(0, 300))]);
    return sh.getLastRow();
  } catch (e) { return 0; }
}
/** Marks the message's own "in" row done (one row per message instead of an "in" row plus an "ok" row). */
function shawnaLogDone(row, kind, text, detail) {
  try {
    const sh = getSS().getSheetByName('Shawna Log');
    if (row && sh && String(sh.getRange(row, 3).getValue()) === cell(text.slice(0, 300))) {
      sh.getRange(row, 2).setValue(kind);
      if (detail) sh.getRange(row, 4).setValue(cell(String(detail).slice(0, 300)));
      return;
    }
  } catch (e) {}
  shawnaLog(kind, text, detail);
}

/** Shawna's short-term memory of her texts with Julián (last 16 messages, about the last day of back and forth). */
function bossMemory(role, text) {
  try {
    const props = PropertiesService.getScriptProperties();
    let h = JSON.parse(props.getProperty('BOSS_HISTORY') || '[]').filter(x => Date.now() - x.at < 14 * 864e5);  // 2 weeks
    h.push({r: role, t: String(text).slice(0, 300), at: Date.now()});
    h = h.slice(-30);
    while (JSON.stringify(h).length > 8500) h.shift();  // stay under the 9 KB limit of one Script Property
    props.setProperty('BOSS_HISTORY', JSON.stringify(h));
  } catch (e) {}
}

/** Recent client texts and Shawna's replies (last 30, 3 days), so Julián can ask "who texted me?". */
/** Julián's own numbers (business line, personal phone, test phone, plus OWN_NUMBERS): never "waiting on you". */
function isOwnNumber(ph) {
  const n = bbE164(ph || ''); if (!n) return false;
  const p = PropertiesService.getScriptProperties();
  return [TEXTING.myNumber, p.getProperty('ALERT_TEXT_TO'), p.getProperty('TEST_TEXT_TO'), '+17028751382', '+17024982571', '+17027631554']
    .concat(String(p.getProperty('OWN_NUMBERS') || '').split(',')).map(x => bbE164(x || '')).indexOf(n) >= 0;
}
/** He answered someone himself: their texts stop counting as waiting on him. */
function textLogAnswered(ph) {
  try {
    const props = PropertiesService.getScriptProperties(), n = bbE164(ph || ''), log = JSON.parse(props.getProperty('TEXT_LOG') || '[]');
    let ch = false; log.forEach(x => { if (bbE164(x.phone || '') === n && !x.read && !x.reply) { x.read = 1; ch = true; } });
    if (ch) props.setProperty('TEXT_LOG', JSON.stringify(log));
  } catch (e) {}
}

function textLog(entry) {
  if (entry && isOwnNumber(entry.phone)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const log = JSON.parse(props.getProperty('TEXT_LOG') || '[]').filter(x => Date.now() - x.at < 3 * 864e5);
    log.push(Object.assign({at: Date.now()}, entry));
    props.setProperty('TEXT_LOG', JSON.stringify(log.slice(-30)));
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

/** True if this number is saved in Julián's iPhone contacts (synced to the texting Mac), i.e. someone he knows personally. */
function macContactByPhone(e164) {
  try {
    const want = String(e164).replace(/\D/g, '').slice(-10);
    const props = PropertiesService.getScriptProperties();
    let digits = null;
    const cached = CacheService.getScriptCache().get('MAC_PHONES');
    if (cached) digits = JSON.parse(cached);
    else {
      digits = [];
      (bbApi('contact').data || []).forEach(c => (c.phoneNumbers || []).forEach(n => digits.push(String(n.address || '').replace(/\D/g, '').slice(-10))));
      CacheService.getScriptCache().put('MAC_PHONES', JSON.stringify(digits).slice(0, 99000), 21600);
    }
    return digits.indexOf(want) >= 0;
  } catch (e) { return true; }  // can't check: play it safe and treat as personal
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

/** Calls the "Shawna Mailer" web app running inside julianc@bcolv.com (email + business calendar). */
function bcolvCall(payload) {
  const props = PropertiesService.getScriptProperties(), url = props.getProperty('BCOLV_MAILER_URL');
  if (!url) return {ok: false, error: 'bcolv mailer not set up'};
  const res = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({key: props.getProperty('BCOLV_MAILER_KEY')}, payload))});
  try { return JSON.parse(res.getContentText()); } catch (e) { return {ok: false, error: 'mailer answered ' + res.getResponseCode()}; }
}

/** Run once after updating the Shawna Mailer: adds a 5-minute test event to the business calendar tomorrow at 8am. */
function testBusinessCalendar() {
  const t = new Date(); t.setDate(t.getDate() + 1); t.setHours(8, 0, 0, 0);
  console.log(JSON.stringify(bcolvCall({action: 'event', title: 'Test from Shawna (delete me)', start: t.toISOString(), end: new Date(t.getTime() + 5 * 6e4).toISOString(), description: 'Business calendar test'})));
}

/** Sends an email FROM julianc@bcolv.com through the small "Shawna Mailer" script in that account. */
function sendFromBcolv(to, subject, text, attachments, plain) {
  const props = PropertiesService.getScriptProperties(), url = props.getProperty('BCOLV_MAILER_URL');
  if (!url) return {ok: false, error: 'bcolv mailer not set up'};
  const html = String(text).split(/\n{2,}/).map(par => '<p style="margin:0 0 12px">' + par.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>') + '</p>').join('') +
    (plain ? '' : '<p style="margin:0">Julián</p>' + EMAIL_FOOTER);
  const res = UrlFetchApp.fetch(url, {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({key: props.getProperty('BCOLV_MAILER_KEY'), to: to, subject: subject, text: plain ? text : text + '\n\nJulián', html: '<div style="font:15px/1.6 Helvetica,Arial,sans-serif;color:#111">' + html + '</div>', attachments: attachments || undefined})});
  try { return JSON.parse(res.getContentText()); } catch (e) { return {ok: false, error: 'mailer answered ' + res.getResponseCode()}; }
}

/** Run once after setting BCOLV_MAILER_URL/KEY: emails Julián's personal Gmail from bcolv as a test. */
function testBcolvMailer() {
  console.log(JSON.stringify(sendFromBcolv(NOTIFY_EMAIL, 'Shawna mailer test', 'If you can read this, Shawna can send email from your bcolv address.')));
}

// ============================================================
// SHAWNA ON THE PHONE (Vapi). Missed calls forward to a Vapi number where Shawna answers (Groq model). When a call
// ends, Vapi posts an "end-of-call-report" to this web app (?vapi=VAPI_KEY). We find the caller in Follow Up Boss by
// phone (or create them, source "Phone call"), add a note with the summary, transcript and recording link, and create
// one "call them back" task (they reached out, so this fits the inbound-only rule). No texts: a call isn't text consent.
// Setup: run setupVoice once; it prints the Server URL to paste into the Vapi assistant.
// ============================================================
function vapiWebhook(e) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('VAPI_KEY') || e.parameter.vapi !== props.getProperty('VAPI_KEY')) return json({status: 'ignored'});
  const m = (JSON.parse((e.postData && e.postData.contents) || '{}').message) || {};
  if (m.type === 'tool-calls') {  // Shawna asks, at the start of a call: who is this caller?
    const caller = bbE164((((m.call || {}).customer) || (m.customer || {})).number || '');
    return json({results: (m.toolCallList || []).map(c => ({toolCallId: c.id, result: vapiLookup(caller)}))});
  }
  if (m.type !== 'end-of-call-report') return json({status: 'ok'});
  try { return vapiEndOfCall(m, props); }
  catch (err) {
    console.log('Vapi call: ' + err.message);
    const cnum = bbE164((((m.call || {}).customer) || {}).number || '');
    try { notifyJulian('📞 A call came in', 'Shawna took a call' + (cnum ? ' from ' + cnum.slice(2).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3') : '') + ' but I couldn\'t log it (' + String(err.message).slice(0, 90) + '). ' + String((m.analysis && m.analysis.summary) || m.summary || '').slice(0, 200)); } catch (e) {}
    return json({status: 'error'});
  }
}
/** Phone numbers of everyone on the Follow Up Boss account (Joey, Tana…), cached 6 h. Calls from them are team calls. */
// Team members who aren't Follow Up Boss users (Julián, Nov 12): Kayleigh on the team, Lily = Joey's assistant.
const TEAM_EXTRA = '7029854247 Kayleigh, 9253957855 Lily';
function teamNumbers() { return Object.keys(teamDirectory()); }
/** {phone: name} for the team: Follow Up Boss users (not Julián) + TEAM_NUMBERS ("7025551234 Joey, 7025559876 Tana" or bare numbers). */
function teamDirectory() {
  const c = CacheService.getScriptCache(), hit = c.get('TEAM_DIR2');
  if (hit) return JSON.parse(hit);
  const dir = {};
  try { (fubFetch('users?limit=50', null, 'get').users || []).forEach(u => { if (u.id === 2) return; [u.phone, u.mobilePhone, u.cellPhone].forEach(x => { const n = bbE164(x || ''); if (n) dir[n] = u.name; }); }); } catch (e) {}
  (TEAM_EXTRA + ',' + (PropertiesService.getScriptProperties().getProperty('TEAM_NUMBERS') || '')).split(',').forEach(x => {
    const n = bbE164((String(x).match(/[\d()+.\s-]{10,}/) || [''])[0]); if (!n) return;
    dir[n] = String(x).replace(/[\d()+.-]/g, '').trim() || dir[n] || 'a team member';
  });
  c.put('TEAM_DIR2', JSON.stringify(dir), 6 * 3600);
  return dir;
}
/** True only if this Follow Up Boss contact belongs to Julián (never write calls into Joey's or Tana's contacts). */
function isMine(personId) {
  try {
    const me = Number(PropertiesService.getScriptProperties().getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
    const p = fubFetch('people/' + personId + '?fields=id,assignedUserId', null, 'get');
    return !p.assignedUserId || p.assignedUserId === me;
  } catch (e) { return false; }  // can't even read it: it's someone else's
}

function vapiEndOfCall(m, props) {
  const call = m.call || {}, num = bbE164(((call.customer || {}).number) || '');
  let summary = (m.analysis && m.analysis.summary) || m.summary || '';
  const transcript = String(m.transcript || (m.artifact && m.artifact.transcript) || '').slice(0, 6000);
  const rec = m.recordingUrl || (m.artifact && m.artifact.recordingUrl) || '';
  const mins = m.durationSeconds ? Math.round(m.durationSeconds / 6) / 10 + ' min' : '';
  let callerName = '', recapInfo = null;
  if (transcript && props.getProperty('GROQ_API_KEY')) {
    try {
      const r = JSON.parse(groqChat(props.getProperty('GROQ_API_KEY'), props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
        {role: 'system', content: 'This is a phone call answered by ' + assistantName() + ', the AI assistant of Julián, a real estate agent. Summarize it for Julián in 2 or 3 short sentences: who called, what they want (timeline, price range, area, pre-approval, agent status if mentioned), and the next step. No dashes. Also give the CALLER\'s name only if the caller said it (never the assistant\'s or Julián\'s name). textOk: true ONLY if the caller clearly agreed to get a text: they said yes when asked, OR they asked to be texted themselves ("just text me", "text me the details"). appointment: the date, time, type and place of any appointment that was booked or that the caller asked Julián to put on his calendar (a showing, open house, team meeting, dinner, call), as a short title like "Team meeting with Joey at the office", else empty. appointmentISO: the appointment start as YYYY-MM-DDTHH:mm in Las Vegas time (today is ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE yyyy-MM-dd') + '), else empty. minutes: its length (default 30). deal: offer, under_contract, closed or listing if the caller said so, else empty. recap: 1 or 2 warm sentences to the caller recapping what they talked about and the next step, written as ' + assistantName() + ' (no dashes). Return JSON {"summary": "...", "name": "caller full name or empty", "textOk": true|false, "appointment": "...", "appointmentISO": "...", "minutes": 30, "deal": "...", "recap": "..."}.'},
        {role: 'user', content: transcript.slice(0, 5000)}]));
      if (!summary) summary = humanize(r.summary || '');
      callerName = String(r.name || '').trim();
      recapInfo = {ok: r.textOk === true, appt: String(r.appointment || '').trim(), iso: String(r.appointmentISO || ''), minutes: Number(r.minutes) || 30, deal: String(r.deal || ''), recap: humanize(r.recap || '')};
      if (/^(no|unknown|shawna|juli[aá]n)\b/i.test(callerName)) callerName = '';
    } catch (e) {}
  }
  const boss0 = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  const fmtNum = n => n ? n.slice(2).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3') : 'an unknown number';
  const mate = num && num !== boss0 ? teamDirectory()[num] : '';
  if (mate) {  // Joey / Tana: Shawna works for Julián here too. Their request goes to him only, never into anyone's CRM.
    let added = '';
    const tm = recapInfo && recapInfo.iso.match(/(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
    if (tm) {
      try {
        const start = Utilities.parseDate(tm[1] + ' ' + tm[2], 'America/Los_Angeles', 'yyyy-MM-dd HH:mm'), end = new Date(start.getTime() + (recapInfo.minutes || 60) * 6e4);
        const title = recapInfo.appt || ('Meeting with ' + mate);
        const cal = CalendarApp.getDefaultCalendar();
        if (!cal.getEvents(new Date(start.getTime() - 6e4), new Date(start.getTime() + 6e4)).some(e => e.getTitle() === title))
          cal.createEvent(title, start, end, {description: 'Set by ' + mate + ' on a call with ' + assistantName() + '.\n' + String(summary || '')});
        added = '\n📅 Added to your calendar: ' + title + ', ' + Utilities.formatDate(start, 'America/Los_Angeles', "EEE MMM d 'at' h:mm a") + '.';
      } catch (e) { added = '\n(I couldn\'t add it to your calendar: ' + String(e.message).slice(0, 80) + ')'; }
    }
    // Teammates always get the recap text (colleagues who called; Julián, Nov 12). Never logged in anyone's CRM.
    let sentRecap = '';
    if (recapInfo && (recapInfo.recap || recapInfo.appt)) {
      try {
        const first = String(mate).split(/[\s(]/)[0];
        bbSend(num, 'Hi ' + first + ', it\'s ' + assistantName() + ', Julián\'s assistant 👋🏾 Thanks for calling! ' + (recapInfo.recap || '') +
          (recapInfo.appt ? ' You\'re set for ' + recapInfo.appt + '.' : '') + ' I let Julián know.');
        sentRecap = '\n💬 Recap texted to ' + first + '.';
      } catch (e) { sentRecap = '\n(The recap text to ' + mate + ' didn\'t go: ' + String(e.message).slice(0, 60) + ')'; }
    }
    notifyJulian('📞 ' + mate + ' called', String(summary || 'no summary').slice(0, 400) + added + sentRecap + (rec ? '\nRecording: ' + rec : ''), '', mate);
    return json({status: 'team'});
  }
  let person = null;
  if (num) person = (fubFetch('people?limit=1&fields=id,firstName,lastName&phone=' + encodeURIComponent(num.slice(2)), null, 'get').people || [])[0];
  const wasKnown = !!person;
  if (person && callerName && !realFirst(person) && num !== boss0) {  // they told Shawna their name: fix "No Name"
    const parts = callerName.split(/\s+/);
    try { fubFetch('people/' + person.id, {firstName: parts[0], lastName: parts.slice(1).join(' ')}, 'put'); } catch (e) {}
  }
  if (!person) {
    // Never a "No Name" contact: their name if they gave it, else "Caller" + their number so the call-back task makes sense.
    const parts = callerName ? callerName.split(/\s+/) : ['Caller', num ? num.slice(2).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3') : ''];
    const r = fubFetch('events', {source: 'Phone call', system: 'juliancrider.com', type: 'Inquiry',
      message: 'Called and talked with ' + assistantName() + ': ' + summary, person: {firstName: parts[0], lastName: parts.slice(1).join(' '), phones: num ? [{value: num}] : [], tags: ['Phone Lead'],
        assignedUserId: Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id}});
    person = (r.personId && {id: r.personId}) || r.person || (num && (fubFetch('people?limit=1&fields=id,firstName&phone=' + encodeURIComponent(num.slice(2)), null, 'get').people || [])[0]);
  }
  if (!person) { notifyJulian('📞 ' + (callerName || fmtNum(num)) + ' called', String(summary || 'no summary').slice(0, 400)); return json({status: 'no person'}); }
  if (!isMine(person.id)) {  // their contact belongs to Joey or Tana: tell Julián, don't write in their record
    notifyJulian('📞 ' + (callerName || fmtNum(num)) + ' called', String(summary || 'no summary').slice(0, 400) + (rec ? '\nRecording: ' + rec : '') +
      '\n(They\'re in another agent\'s Follow Up Boss, so I didn\'t log it there.)');
    return json({status: 'not mine'});
  }
  fubFetch('notes', {personId: person.id, subject: 'Call with ' + assistantName() + (mins ? ' (' + mins + ')' : ''),
    body: (summary || 'No summary.') + (rec ? '\n\nRecording: ' + rec : '') + (transcript ? '\n\nTranscript:\n' + transcript : '')});
  const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
  if (num && num === boss) {  // Julián testing from his own phone: no call-back task, but show him what a caller would get
    try {
      notifyJulian('🧪 Test call logged', String(summary || 'no summary').slice(0, 300) +
        (recapInfo && recapInfo.ok ? '\n\nThe caller said yes to a recap, so they\'d get this text:\n"' + (recapInfo.recap || '') + (recapInfo.appt ? ' You\'re set for ' + recapInfo.appt + '.' : '') + '"' : '\n\nNo recap text (they didn\'t say yes to one on the call).'));
    } catch (e) {}
    return json({status: 'ok'});
  }
  try { notifyJulian((callerName || 'Someone') + ' called', assistantName() + ' took a call' + (num ? ' from ' + num.slice(2).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3') : '') + ': ' + String(summary || 'no summary').slice(0, 220) + ' There\'s a call-back task in Follow Up Boss.'); } catch (e) {}
  createCallTask(person.id, assistantName() + ' took their call: ' + summary.slice(0, 140) + ' Call them back.', 'Call');
  stat('calls');
  // Did the appointment Shawna booked actually land on his calendar? If yes, mirror it in Follow Up Boss; if not, ask him.
  try {
    const m2 = recapInfo && recapInfo.iso.match(/(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
    if (m2) {
      const start = Utilities.parseDate(m2[1] + ' ' + m2[2], 'America/Los_Angeles', 'yyyy-MM-dd HH:mm'), mins = recapInfo.minutes || 30;
      const from = new Date(start.getTime() - 45 * 6e4), to = new Date(start.getTime() + 45 * 6e4);
      const onCal = CalendarApp.getDefaultCalendar().getEvents(from, to).length > 0 ||
        ((bcolvCall({action: 'events', from: from.toISOString(), to: to.toISOString()}) || {}).events || []).length > 0;
      const who2 = callerName || 'caller', title = 'Call with ' + who2;
      if (onCal) {
        fubFetch('appointments', {title: title, start: start.toISOString(), end: new Date(start.getTime() + mins * 6e4).toISOString(), invitees: [{personId: person.id}]});
        fubFetch('notes', {personId: person.id, subject: 'Appointment confirmed', body: 'Booked on the call with ' + assistantName() + ' and found on your calendar: ' + recapInfo.appt});
        stat('booked');
      } else {
        const n = addPending({kind: 'event', personId: person.id, who: who2, title: title, start: start.getTime(), minutes: mins, phone: num});
        notifyJulian('📅 Check this booking', assistantName() + ' told ' + who2 + ' they\'re set for ' + recapInfo.appt + ', but I don\'t see it on your calendar. Want me to add it? ' + yesHint(n) + '.');
      }
    }
    if (recapInfo && DEAL_SIGNALS[recapInfo.deal]) suggestDeal(person.id, recapInfo.deal, 'said on a call');
  } catch (e) { console.log('Call booking check: ' + e.message); }
  // Recap text after the call: only when the caller said yes to a text on the call (a call alone isn't text consent).
  if (recapInfo && recapInfo.ok && num && (recapInfo.recap || recapInfo.appt)) {
    try {
      const pp = fubFetch('people/' + person.id + '?fields=id,firstName,tags', null, 'get'), tags = pp.tags || [];
      if (tags.indexOf('Text Opt-Out') < 0) {
        if (tags.indexOf('Text OK') < 0) fubFetch('people/' + person.id, {tags: tags.concat('Text OK')}, 'put');
        fubFetch('notes', {personId: person.id, subject: 'Text consent', body: 'Said yes on the call with ' + assistantName() + ' to getting a text recap.'});
        const first = realFirst(pp) || callerName.split(' ')[0] || '';
        const msg = 'Hi' + (first ? ' ' + first : '') + ', it\'s ' + assistantName() + ', Julián\'s assistant 👋🏾 Thanks for calling! ' + (recapInfo.recap || '') +
          (recapInfo.appt ? ' You\'re set for ' + recapInfo.appt + '.' : '') + ' If anything changes, just text me here.';
        ensureTab(getSS(), TEXT_QUEUE_TAB).appendRow([new Date(Date.now() + 60e3), person.id, cell(first || callerName), 'call-recap', 'pending', 0, '', cell(msg)]);
        ensureTextQueueTrigger();
      }
    } catch (e) { console.log('Call recap: ' + e.message); }
  } else if (num) {
    // Julián, Nov 12: everyone who calls gets ONE text after the call (they called us; informational, with STOP), from his
    // number, ~1 minute after hanging up (9am-9pm; later calls get it in the morning). Never for opted-out numbers.
    try {
      const pp = fubFetch('people/' + person.id + '?fields=id,firstName,lastName,tags', null, 'get'), tags = pp.tags || [];
      if (tags.indexOf('Text Opt-Out') < 0 && !tags.some(t => /^Test\b/i.test(t))) {
        const first = realFirst(pp) || (callerName && !/^caller\b/i.test(callerName) ? callerName.split(' ')[0] : '');
        const gist = recapInfo && recapInfo.recap ? ' ' + recapInfo.recap : ' I passed your message to Julián and he\'ll get back to you soon.';
        const msg = 'Hi' + (first ? ' ' + first : '') + ', it\'s ' + assistantName() + ', Julián\'s assistant. Thanks for calling!' + gist +
          (recapInfo && recapInfo.appt ? ' You\'re set for ' + recapInfo.appt + '.' : '') + ' Text me here anytime. (Reply STOP to opt out.)';
        ensureTab(getSS(), TEXT_QUEUE_TAB).appendRow([new Date(Date.now() + 60e3), person.id, cell(first || pp.firstName || ''), 'call-thanks', 'pending', 0, '', cell(msg)]);
        ensureTextQueueTrigger();
      }
    } catch (e) { console.log('Call thank-you: ' + e.message); }
  }
  return json({status: 'ok'});
}

// ============================================================
// SETTINGS TAB (Nov 11). Apps Script's Script Properties page goes read-only after 50 entries, so Julián adds keys and
// settings in the "Settings" sheet tab instead: setting name in column A, value in column B. Every 5 minutes the script
// saves it into Script Properties and erases the value from the sheet (keys never sit in the sheet).
// ============================================================
const SETTINGS_TAB = {name: 'Settings', color: '#5F6368', headers: ['Setting name (for example VAPI_API_KEY)', 'Value (saved into the script within 5 minutes, then erased here)', 'Status']};
function applySettings() {
  const sheet = ensureTab(getSS(), SETTINGS_TAB), last = sheet.getLastRow();
  if (last < 2) return;
  const props = PropertiesService.getScriptProperties();
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach((r, i) => {
    const name = String(r[0] || '').trim().toUpperCase().replace(/\s+/g, '_'), value = String(r[1] || '').trim();
    if (!name || !value) return;
    const row = i + 2;
    if (!/^[A-Z][A-Z0-9_]{2,60}$/.test(name)) { sheet.getRange(row, 3).setValue('Not saved: the name can only have letters, numbers and _'); return; }
    props.setProperty(name, value);
    sheet.getRange(row, 2).setValue('');
    sheet.getRange(row, 3).setValue('Saved ' + Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMM d, h:mm a') + ' (ends in …' + value.slice(-4) + ')');
  });
}

/** Hourly: removes housekeeping entries that are no longer needed (old reminder marks, expired mutes), so Script
 *  Properties stays small. Never touches keys or settings. */
/** Daily: the automatic reminder tasks ("They texted back…", "they reached out…", "New Showing request…", "Shawna took
 *  their call…") are checked off after 7 days, so his task list only shows what's still real. */
function cleanupOldTasks() {
  const props = PropertiesService.getScriptProperties(), day = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
  if (props.getProperty('TASK_CLEAN_DAY') === day) return;
  props.setProperty('TASK_CLEAN_DAY', day);
  const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
  (fubFetch('tasks?assignedUserId=' + me + '&isCompleted=false&limit=100&fields=id,name,dueDate', null, 'get').tasks || [])
    .filter(t => (Date.now() - new Date(t.dueDate)) / 864e5 > 7 && /They texted back|they reached out|New Showing request|took their call|They asked for you/i.test(String(t.name || '')))
    .forEach(t => { try { fubFetch('tasks/' + t.id, {isCompleted: true}, 'put'); Utilities.sleep(100); } catch (e) {} });
}

function cleanupProperties() {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => {
    const v = all[k];
    if (/^REM_/.test(k) && Number(k.split('_').pop()) < now - 864e5) props.deleteProperty(k);  // reminder already sent
    else if (/^(AI_MUTE_|AI_SENDING_|AI_LAST_OUT_)/.test(k) && Number(v) < now - 36e5) props.deleteProperty(k);
    else if (/^NEWNUM_/.test(k)) { try { if (JSON.parse(v).at < now - 7 * 864e5) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
  });
}

/** What Shawna should know about a caller, from Follow Up Boss (read to herself, never out loud). */
function vapiLookup(num) {
  try {
    if (!num) return 'No caller ID. Ask their name.';
    const props = PropertiesService.getScriptProperties();
    if (num === bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '')) return 'This is Julián himself calling to test you. Be relaxed and natural.';
    const mate = teamDirectory()[num];
    if (mate) return 'TEAM MEMBER: ' + mate + ', on Julián\'s team (The DeBenedetto Team). Greet them by first name like a colleague. You are Julián\'s assistant: take whatever they need for Julián, like setting a team meeting, an open house, a showing, a dinner or a call. Get what it is, the day, the time and the place, read it back once, and tell them you\'ll put it on Julián\'s calendar and let him know right away. Don\'t qualify them as a buyer or seller, don\'t ask for email. Before goodbye, ask the RECAP TEXT question like with anyone else (they get the recap by text).';
    const p = (fubFetch('people?limit=1&fields=id,firstName,lastName,stage,tags,price&phone=' + encodeURIComponent(num.slice(2)), null, 'get').people || [])[0];
    if (!p) return 'New caller, not in Julián\'s contacts. Ask their name early.';
    const first = realFirst(p), tags = (p.tags || []).filter(t => /buyer|seller|relocation|luxury|open house|past client|hot|warm|agent|not now|pre-?approved/i.test(t));
    return ['Known contact: ' + (first ? [first, p.lastName].join(' ') : 'name unknown, ask it'), 'Stage: ' + (p.stage || 'Lead'),
      tags.length ? 'Tags: ' + tags.join(', ') : '', p.price ? 'Price range: $' + Number(p.price).toLocaleString('en-US') : '',
      'Recent history (oldest first, for you only, never read it out):\n' + personHistory(p.id).split('\n').slice(-8).join('\n')].filter(Boolean).join('\n').slice(0, 2500);
  } catch (e) { return 'Lookup unavailable right now. Ask their name.'; }
}

/** Shawna's phone instructions (Vapi). Built from the same facts she uses for texts, so both stay in sync. */
function vapiPrompt() {
  const name = assistantName();
  return `You are ${name}, the inbound AI receptionist for Julián J. Crider, a luxury real estate agent and developer (Realtor®) with The DeBenedetto Team at Boeckle & Co., serving Las Vegas and Henderson, Nevada. Your greeting already said you're Julián's assistant and that the call may be recorded. Never claim to be human; if asked, say warmly that you're Julián's AI assistant.
${assistantKnowledge()}

FIRST: call the lookup_caller tool right away, quietly, before your second sentence. If they're a known contact, use their first name and pick up where things left off (for example "Hi Dan, good to hear from you. Did the report make sense?"). Never read the notes out loud or mention a CRM, and never ask for something you already know.

IDENTITY AND STYLE
- Warm, confident, polished and composed, like a top luxury real estate advisor and a genuinely kind person.
- Short, natural sentences, calm pace. Ask ONE question, then stop and wait for the answer. Never ask two things at once, never answer for them, never assume their answer, and never list what you "have so far" (no "phone awaiting").
- If they pause, give them time. A short "take your time" is fine.
- Chris Voss, used naturally and not overused: mirror their last few words, labels ("It sounds like..."), calm "what" and "how" questions, and once in a while a no-oriented question ("Would it be a bad idea to...?"). When they answer "no" to a no-oriented question, that means yes, go ahead, so keep helping them.
- Never say "sorry, I can't." Say "Here's what I can do for you." Never say "I'm not sure." Say "I'll confirm that for you."

PRIMARY GOAL: understand why they're calling, help them, qualify genuine buyers and sellers, collect contact details, and offer a call or showing with Julián.

CONTACT DETAILS (one item at a time): their full name first (if you don't already have it), then what they're looking for. Email is OPTIONAL: ask for it once, and only if they want things by email. If they say "just text me" or prefer text, never ask for email again: say "Perfect, I'll text you at this number" and move on. You already have their phone number from caller ID: never ask for it; confirm it once near the end.
NEVER GET STUCK: ask for any detail (email, address, spelling) at most twice. If you still don't have it clearly, say "No problem, Julián will confirm that with you by text" and move on. Never repeat the same question three times.
ADDRESSES: repeat it back once in plain words ("4775 West Teco Avenue, is that right?"). If they say it's wrong, ask them to say just the street name slowly, or to spell it. If it's still unclear, note it as "address to confirm" and keep going. Numbers: say each digit clearly.

BUYERS (one at a time, conversationally, skip anything they already told you): what's prompting the move, when they'd like to move, areas or neighborhoods, price range, whether they're pre-approved or paying cash, whether they're working with a real estate agent, and where they live now. Don't interrogate.
SELLERS (one at a time): the property address and their selling timeline, then their name and email if you don't have them. Julián can do a private home valuation (no listing agreement, nothing made public).

SCHEDULING
- Offer to book a call (phone or video) or a showing with Julián using the connected calendar tools, if they're available.
- Hours: Monday through Saturday, 9:00 AM to 6:00 PM Pacific Time. Never offer Sundays or times outside these hours.
- First check calendar availability and offer a couple of options. Confirm the exact date, time, time zone, type of appointment, their name, email and the property or topic before creating the event.
- Never say an appointment is booked unless the calendar tool succeeded. If scheduling can't be completed, get their preferred times and say Julián will follow up to confirm. They can also book themselves at juliancrider.com/?book.

TEAM MEMBERS (the lookup says TEAM MEMBER, e.g. Joey or Tana): you are Julián's assistant for them too. Take the request (team meeting, open house, showing, dinner, call), get what, day, time and place, read it back once, and say "I'll put it on Julián's calendar and let him know right now." No buyer questions, no email. Then ask the RECAP TEXT question.
IF THEY ASK FOR JULIÁN: say he'll call them back shortly; confirm their name, email and the reason for the call. Don't transfer calls. Then, in the meantime, offer to help yourself: "While you wait for Julián, I'm happy to answer any questions." Answer anything you can (how buying or selling works, loans in general, neighborhoods, today's average rates, how Julián works), so they leave the call with what they needed.

INTEREST RATES AND LENDING: you can share today's averages: ${ratesText() || 'not available right now'} Never quote their own rate or payment: that depends on credit, down payment and loan type, so the right number comes from a lender. You can explain how loans and the buying or selling process work in plain words. Offer to have Julián connect them with a local lender he trusts, or set up a quick call with Julián. There's also a live rates page and payment calculator on juliancrider.com.

LANGUAGE: always answer in the language the caller is speaking right now. If they switch languages mid-call (even after speaking English), switch with them immediately, and switch back if they do. Keep the same warmth and rules in every language. If Julián will need to follow up in that language, note it in the next step.

SAFETY, COMPLIANCE AND BOUNDARIES
- Never quote or invent listing prices, availability, interest rates or market statistics (you may ask for their own target price range).
- No legal, tax, lending or loan advice; offer a follow-up from Julián or suggest the right licensed professional.
- Fair Housing: never comment on schools' quality, safety or crime, demographics, religion, race, family status, disability or any protected class, or who lives in a neighborhood, and never steer anyone toward or away from an area for those reasons.
- No guarantees about properties, returns, appreciation, availability or outcomes. Don't invent facts: "I'll confirm that for you."
- If it's about an offer, a contract, a negotiation, commission, a complaint or anything urgent, say Julián will call them back today.
- Collect only what's needed. No dashes of any kind in what you say.

RECAP TEXT (REQUIRED ON EVERY CALL, team members included, never skip it): before you say goodbye, always ask: "Before you go, would it be okay if I text you a quick recap of our call?" (add "with your appointment details" if one was booked). Ask it even on short calls, even if they only left a message for Julián or asked for someone else, and even if the call wandered. If they already asked to be texted, that's a yes: just say "I'll text you a quick recap." The recap goes by text to the number they're calling from: never ask for their email or number for the recap. Only a clear yes counts; if they hesitate or say no, that's completely fine.
ENDING THE CALL: first ask the RECAP TEXT question if you haven't yet (it is required). Then briefly repeat the next step (for example "Julián will call you this afternoon about the Summerlin homes"), confirm the best callback details, and ask if there's anything else you can help with. Only end the call after they say goodbye or clearly want to hang up. Always close warmly with "Have a wonderful rest of your day, and we'll speak with you soon." Never end with just "goodbye."`;
}

/** Hourly (from autoUpdateFromGitHub): pushes Shawna's phone prompt, the caller lookup tool and a calmer turn-taking
 *  setting to the Vapi assistant, only when something changed. Needs VAPI_API_KEY + VAPI_ASSISTANT_ID. */
function syncVapiAssistant(force) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('VAPI_API_KEY'), id = props.getProperty('VAPI_ASSISTANT_ID');
  if (!key || !id || !props.getProperty('VAPI_KEY')) return 'Add VAPI_API_KEY and VAPI_ASSISTANT_ID in Script Properties first.';
  const prompt = vapiPrompt(), url = TEXTING.webApp + '?vapi=' + props.getProperty('VAPI_KEY');
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, prompt + url + 'v4'));
  const api = (method, body) => UrlFetchApp.fetch('https://api.vapi.ai/assistant/' + id, {method: method, muteHttpExceptions: true,
    contentType: 'application/json', headers: {Authorization: 'Bearer ' + key}, payload: body ? JSON.stringify(body) : undefined});
  const cur = api('get');
  if (cur.getResponseCode() >= 300) throw new Error('Vapi ' + cur.getResponseCode() + ': ' + cur.getContentText().slice(0, 200));
  const a = JSON.parse(cur.getContentText()), model = Object.assign({}, a.model || {});
  // Up to date only if Vapi really has today's prompt and the caller lookup: someone publishing an old dashboard draft
  // (or changing the model) is put right within the hour. The model/provider he picks in Vapi is always kept.
  const live = ((model.messages || [])[0] || {}).content === prompt && (model.tools || []).some(t => t.function && t.function.name === 'lookup_caller');
  if (!force && live && props.getProperty('VAPI_SYNC') === hash) return 'Vapi already up to date.';
  model.messages = [{role: 'system', content: prompt}];
  model.tools = (model.tools || []).filter(t => !(t.function && t.function.name === 'lookup_caller')).concat([{
    type: 'function', async: false, server: {url: url}, messages: [],
    function: {name: 'lookup_caller', description: 'Looks up the caller in Julián\'s contacts by their phone number: name, what they want and recent history. Call it once at the very start of every call.',
      parameters: {type: 'object', properties: {}}}}]);
  // Hear and speak any language: Deepgram Nova 3 "multi" transcription, and the multilingual version of her ElevenLabs voice.
  const transcriber = {provider: 'deepgram', model: 'nova-3', language: 'multi'};
  const voice = a.voice && a.voice.provider === '11labs' && /^eleven_(turbo_v2|flash_v2|monolingual_v1)$/.test(String(a.voice.model || ''))
    ? Object.assign({}, a.voice, {model: String(a.voice.model).replace(/^eleven_monolingual_v1$/, 'eleven_multilingual_v2').replace(/_v2$/, '_v2_5')}) : null;
  const body = {model: model, transcriber: transcriber, endCallMessage: 'Have a wonderful rest of your day, and we\'ll speak with you soon.',
    startSpeakingPlan: Object.assign({}, a.startSpeakingPlan || {}, {waitSeconds: 0.8})};
  if (voice) body.voice = voice;
  const res = api('patch', body);
  if (res.getResponseCode() >= 300) throw new Error('Vapi update ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  props.setProperty('VAPI_SYNC', hash);
  return 'Vapi updated: new phone instructions, caller lookup and calmer pacing.';
}

/** Run once by hand to push Shawna's phone setup to Vapi now (it also happens on its own every hour). */
function updateVapiNow() { console.log(syncVapiAssistant(true)); }

/** Run once: makes the secret key and prints the Server URL for the Vapi assistant. */
function setupVoice() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('VAPI_KEY')) props.setProperty('VAPI_KEY', Utilities.getUuid().replace(/-/g, ''));
  console.log('In Vapi > Assistants > ' + assistantName() + ' > Advanced > Server URL, paste:');
  console.log(TEXTING.webApp + '?vapi=' + props.getProperty('VAPI_KEY'));
  console.log('Under Server Messages, keep only "end-of-call-report".');
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
    {step: 'oh-72h', hours: 72, text: "Would it be a bad idea if I lined up 2 or 3 homes for you to see this weekend? You can pick what you'd like to see here: juliancrider.com/?showing"},
    // 90-day follow-up (Julián approved Nov 10; modeled on Joey's 36-touch arc, rewritten in his voice). Stops once they reply.
    {step: 'oh-d6', hours: 6 * 24, text: "Hey {first}, Julián here. Sounds like you're still figuring out what the right fit looks like. Would it be a bad idea if I sent you a couple that feel like the one you toured? 🏡"},
    {step: 'oh-d10', hours: 10 * 24, text: "Quick one. Seems like the market noise can be a lot right now. Would you be against me sending you a short snapshot of what's actually selling near that house?"},
    {step: 'oh-d17', hours: 17 * 24, text: "Is it ridiculous to ask what would have made that house a yes for you? It helps me skip the ones that would waste your time."},
    {step: 'oh-d35', hours: 35 * 24, text: "Seems like timing is the big question for most people I talk to. Would it be off base to say you're more in the someday camp than the next few months camp?"},
    {step: 'oh-d45', hours: 45 * 24, text: "A lot of people I meet at open houses ask where to even start with neighborhoods here. Would you be against me sending you my guide? It's the stuff I wish someone told me 📩"},
    {step: 'oh-d60', hours: 60 * 24, text: "Have you given up on finding the right place, or just hit pause?"},
    {step: 'oh-d75', hours: 75 * 24, text: "No house talk today. Just hoping things are good on your end, {first} ☀️"},
    {step: 'oh-d90', hours: 90 * 24, text: "Have you decided to stay put for now? Totally fine if so and I'll stop checking in. If not, here's my calendar: juliancrider.com/?book"}],
  neighbor: [
    {step: 'oh-2h', hours: 2, text: "Hey {first}, it's Julián from the open house today 👋🏾 Thanks for stopping by, neighbor. It's always good to meet the people who actually live here. || [contact card] || juliancrider.com/?relocate || [voice]"},
    {step: 'oh-72h', hours: 72, text: "Hey {first}, Julián again. A few neighbors asked what their own homes would sell for after the open house. Would it be crazy if I ran a quick, private number for yours? No visit needed. 🏡"}],
  // Website forms (home value, guides, account, questions) that didn't already book a time: a welcome text.
  website: [
    {step: 'web-welcome', hours: 0.05, text: "Hey {first}, it's Julián Crider with The DeBenedetto Team 👋🏾 Thanks for reaching out through my site. As a developer and luxury advisor, I help people buy and sell across Las Vegas and Henderson, and I'm glad you found me. Just curious, what has you looking right now? || [contact card]"},
    {step: 'web-2d', hours: 48, text: "Hey {first}, Julián again. What would make the next few months feel like a win for you on the real estate side? 🏡"}],
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
  if (!has('morningBriefing')) ScriptApp.newTrigger('morningBriefing').timeBased().atHour(8).everyDays(1).inTimezone('America/Los_Angeles').create();
  if (!has('weeklyScorecard')) ScriptApp.newTrigger('weeklyScorecard').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).inTimezone('America/Los_Angeles').create();
  if (!has('processTextQueue')) ScriptApp.newTrigger('processTextQueue').timeBased().everyMinutes(15).create();
  if (!has('remindUpcoming')) ScriptApp.newTrigger('remindUpcoming').timeBased().everyMinutes(5).create();
  if (!has('processAssistantQueue')) ScriptApp.newTrigger('processAssistantQueue').timeBased().everyMinutes(1).create();
}

// ============================================================
// APPOINTMENT REMINDERS: a text about 15 minutes before any event on Julián's Google Calendar that has a phone number
// in its notes (website bookings add "Phone: …" automatically; for a showing, add the client's number to the event).
// Title with "Showing" → showing wording (uses the event's location). Skips anyone tagged Text Opt-Out.
// ============================================================
// CALENDAR SYNC (Julián, Nov 12: "everything on my bcolv calendar, add it to my personal calendar"). Every 15 minutes
// the next 30 days of his bcolv (business) calendar and his Follow Up Boss appointments are copied onto his personal
// Google Calendar (the one this script runs in). Moved or deleted ones follow; anything already there isn't doubled.
function syncCalendars() {
  const cache = CacheService.getScriptCache();
  if (cache.get('CALSYNC_RUN')) return;
  cache.put('CALSYNC_RUN', '1', 14 * 60);
  const props = PropertiesService.getScriptProperties(), cal = CalendarApp.getDefaultCalendar();
  const now = new Date(), until = new Date(Date.now() + 30 * 864e5), want = {};
  let bcolvOk = false, fubOk = false;
  if (props.getProperty('BCOLV_MAILER_URL')) {
    try {
      const r = bcolvCall({action: 'events', from: now.toISOString(), to: until.toISOString()}) || {};
      if (r.ok !== false && r.events) bcolvOk = true;
      (r.events || []).forEach(e => { want['b|' + e.title + '|' + e.start] = {title: e.title, start: e.start, end: e.end, note: 'From your bcolv (business) calendar.' + (e.description ? '\n\n' + e.description : '')}; });
    } catch (e) { console.log('bcolv events: ' + e.message); }
  }
  try {
    const me = Number(props.getProperty('FUB_ME') || 0) || fubFetch('me', null, 'get').id;
    (fubFetch('appointments?limit=100&sort=-start', null, 'get').appointments || []).forEach(a => {
      const st = new Date(a.start).getTime();
      if (!(st > now.getTime() && st < until.getTime())) return;
      if (a.createdById !== me && !(a.invitees || []).some(i => i.userId === me)) return;
      const people = (a.invitees || []).filter(i => i.personId).map(i => i.name).join(', ');
      want['f|' + a.id + '|' + st] = {title: a.title, start: st, end: new Date(a.end || st + 36e5).getTime(), location: a.location || '',
        note: 'From Follow Up Boss' + (people ? ' with ' + people : '') + '.' + (a.description ? '\n\n' + a.description : '')};
    });
    fubOk = true;
  } catch (e) { console.log('FUB appointments: ' + e.message); }
  const map = JSON.parse(props.getProperty('CALSYNC') || '{}');
  Object.keys(want).forEach(k => {
    if (map[k]) return;
    const w = want[k], start = new Date(w.start), end = new Date(Math.max(w.end || 0, w.start + 15 * 6e4));
    // Already on his calendar (booked on the site, by Shawna, or by hand)? Then don't add a second one.
    const same = cal.getEvents(new Date(w.start - 6e4), new Date(w.start + 6e4)).filter(e => Math.abs(e.getStartTime().getTime() - w.start) < 6e4 &&
      (e.getTitle().toLowerCase() === String(w.title).toLowerCase() || String(w.title).toLowerCase().indexOf(e.getTitle().toLowerCase().split(':')[0]) >= 0));
    if (same.length) { map[k] = {id: '', start: w.start}; return; }
    const ev = cal.createEvent(w.title, start, end, {location: w.location || '', description: w.note + '\n[synced by Shawna]'});
    map[k] = {id: ev.getId(), start: w.start};
  });
  Object.keys(map).forEach(k => {
    const m = map[k];
    if (m.start < now.getTime()) { delete map[k]; return; }
    const src = k.charAt(0);
    if (!want[k] && ((src === 'b' && bcolvOk) || (src === 'f' && fubOk))) {  // moved or canceled at the source: remove our copy
      if (m.id) { try { const ev = cal.getEventById(m.id); if (ev) ev.deleteEvent(); } catch (e) {} }
      delete map[k];
    }
  });
  let json = JSON.stringify(map);
  while (json.length > 8500) { const ks = Object.keys(map).sort((x, y) => map[y].start - map[x].start); delete map[ks[0]]; json = JSON.stringify(map); }
  props.setProperty('CALSYNC', json);
}

function remindUpcoming() {
  try { processFubOutbox(); } catch (e) { console.log('Outbox: ' + e.message); }
  try { syncCalendars(); } catch (e) { console.log('Calendar sync: ' + e.message); }
  try { processPersonalEmailOutbox(); } catch (e) { console.log('Personal email: ' + e.message); }
  try { checkClientEmails(); } catch (e) { console.log('Email alerts: ' + e.message); }
  try { processEmailAssistant(); } catch (e) { console.log('Email assistant: ' + e.message); }
  try { applySettings(); } catch (e) { console.log('Settings: ' + e.message); }
  const now = Date.now(), props = PropertiesService.getScriptProperties();
  const events = CalendarApp.getDefaultCalendar().getEvents(new Date(now + 9 * 6e4), new Date(now + 21 * 6e4));
  events.forEach(ev => {
    const key = 'REM_' + ev.getId().slice(0, 40) + '_' + ev.getStartTime().getTime();
    if (props.getProperty(key) || ev.isAllDayEvent()) return;
    if (/\[synced by Shawna\]/.test(ev.getDescription() || '')) return;  // copies from bcolv/FUB: no automatic client reminders
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

/** A text Julián approved (scheduled for later): goes to their number unless they opted out, logged in Follow Up Boss. */
function sendApprovedText(personId, text) {
  const p = fubFetch('people/' + personId + '?fields=id,firstName,tags,phones', null, 'get');
  if ((p.tags || []).indexOf('Text Opt-Out') >= 0) return 'skipped: they opted out';
  const phone = (p.phones || []).map(x => bbE164(x.value)).filter(Boolean)[0];
  if (!phone) return 'skipped: no mobile number';
  PropertiesService.getScriptProperties().setProperty('AI_SENDING_' + phone, String(Date.now()));
  bbSendParts(phone, text);
  try { fubFetch('notes', {personId: p.id, subject: 'Text sent', body: text + '\n\n(scheduled by Julián)'}); } catch (e) {}
  return 'sent';
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
    if (!/^(oh-2h|pc-|call-recap|call-followup|call-thanks|scheduled)/.test(step) && props.getProperty('BB_TASK_' + id)) return cell(5, 'skipped: they replied');
    try {
      const res = step === 'scheduled' || step === 'call-thanks' ? sendApprovedText(id, text) : bbTextPerson(id, text);
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

// ============================================================
// EMAIL ASSISTANT (Nov 11). Every 5 minutes, 7am-10pm, Shawna reads new unread emails in julianc@bcolv.com (through the
// Shawna Mailer v3) from people who are in Follow Up Boss and assigned to Julián, and writes a reply as Julián, using
// their texts, notes and past emails. EMAIL_ASSISTANT = draft (default): the reply is saved as a Gmail draft in that
// thread and Julián gets a text to review and send it. auto: she sends it herself (but anything that needs him, like
// offers, contracts, complaints or Has Agent, is still only drafted). off: nothing. Emails from people who aren't in
// Follow Up Boss (vendors, newsletters, other agents) are left alone. "Not right now" / "not interested" emails change
// their campaign (applyIntent). Emails stay unread in his inbox.
// ============================================================
function processEmailAssistant(list) {
  const props = PropertiesService.getScriptProperties(), mode = (props.getProperty('EMAIL_ASSISTANT') || 'draft').toLowerCase();
  const key = props.getProperty('GROQ_API_KEY');
  if (mode === 'off' || !key || !props.getProperty('BCOLV_MAILER_URL')) return;
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  if (hour < 7 || hour >= 22) return;
  const v = bcolvCall({action: 'version'});
  if (!v || !(v.version >= 3)) return;  // the Shawna Mailer needs the v3 update to read email
  let emails = Array.isArray(list) ? list : null;
  if (!emails) { const res = bcolvCall({action: 'inbox', days: 2, max: 20}); if (!res || !res.ok) return 0; emails = res.emails || []; }
  let handled = 0;
  const me = Number(props.getProperty('FUB_ME') || 0) || (() => { const id = fubFetch('me', null, 'get').id; props.setProperty('FUB_ME', id); return id; })();
  const done = JSON.parse(props.getProperty('EMAIL_AI_DONE') || '[]'), name = assistantName();
  emails.filter(m => done.indexOf(m.id) < 0 && !m.fromMe).slice(0, 5).forEach(m => {
    done.push(m.id); props.setProperty('EMAIL_AI_DONE', JSON.stringify(done.slice(-400)));  // mark first: never answer twice
    const addr = String((String(m.from).match(/<([^>]+)>/) || [null, m.from])[1] || '').trim().toLowerCase();
    // Nothing new from them (a copy of Julián's own email, or just a signature): skip it.
    const newWords = String(m.body || '').replace(/juli[aá]n\s*crider|\(?702\)?[\s.-]*498[\s.-]*2571|sent from my \w+/gi, '').replace(/\s+/g, ' ').trim();
    if (newWords.length < 3) return;
    if (!validEmail(addr) || /no-?reply|notification|mailer-daemon|postmaster|bcolv\.com|followupboss|zillow|realscout|^info@|^support@|^hello@/i.test(addr)) return;
    // Automatic replies are not people writing back (Nov 12: a cigar shop's "we will contact you in the order received").
    if (/auto.?reply|automatic reply|out of (the )?office|thank you for (contacting|your (email|message|inquiry))|we (will|'ll) (contact|get back to) you|in the order (it|your email) was received|do not reply|this mailbox is not monitored/i.test(m.subject + ' ' + m.body)) return;
    const p = (fubFetch('people?limit=1&fields=id,firstName,lastName,tags,stage,assignedUserId,phones&email=' + encodeURIComponent(addr), null, 'get').people || [])[0];
    if (!p || (p.assignedUserId && p.assignedUserId !== me)) return;
    const tags = p.tags || [];
    if (tags.indexOf('No Shawna') >= 0 || tags.indexOf('Partner') >= 0) return;
    const phone = bbE164(((p.phones || [])[0] || {}).value || ''), first = realFirst(p);
    let out;
    try {
      out = JSON.parse(groqChat(key, props.getProperty('GROQ_MODEL') || ASSISTANT.model, [
        {role: 'system', content: assistantPrompt(name, assistantPersona(phone || '')) + '\nTHIS IS AN EMAIL REPLY, not a text: 2 to 5 short sentences, warm and plain. Start with "Hi ' + (first || 'there') + '," on its own line. No subject line and no sign-off or signature (they are added automatically). Fewer emojis than in texts (0 or 1). Also add "summary": what they wrote, in 1 or 2 plain sentences for Julián (keep their key words, numbers, dates, addresses).'},
        {role: 'user', content: 'Contact: ' + (first ? [first, p.lastName].join(' ') : '(name unknown)') + '\nStage: ' + (p.stage || '') + '\nTags: ' + tags.join(', ') +
          (phone ? '\nTheir text conversation (oldest first):\n' + bbThread(phone) : '') + '\nFollow Up Boss history (oldest first):\n' + personHistory(p.id) +
          '\n\nThey just emailed. Subject: "' + m.subject + '"\n"' + String(m.body || '').slice(0, 2500) + '"\n\nReply as JSON only.'}]));
    } catch (e) { console.log('Email assistant (' + addr + '): ' + e.message); return; }
    learnFromReply(p, out, 'an email');
    const reply = humanize(out.reply || '');
    if (!reply) return;
    const sensitive = !!out.escalate || ASSISTANT.escalate.test(m.body + ' ' + m.subject) || tags.indexOf('Has Agent') >= 0;
    const who = [first, p.lastName].join(' ').trim() || addr, said = humanize(out.summary || String(m.body || '').replace(/\s+/g, ' ').slice(0, 300));
    const subject = /^re:/i.test(m.subject) ? m.subject : 'Re: ' + m.subject;
    const boss = bbE164(props.getProperty('ALERT_TEXT_TO') || props.getProperty('TEST_TEXT_TO') || '');
    if (mode === 'auto' && !sensitive) {  // EMAIL_ASSISTANT = auto: she answers, then tells him what she said
      const html = '<div style="font:15px/1.6 Helvetica,Arial,sans-serif;color:#111">' + reply.split(/\n{2,}|\n/).map(par => '<p style="margin:0 0 12px">' +
        par.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</p>').join('') + '<p style="margin:0">Julián</p>' + EMAIL_FOOTER + '</div>';
      const r = bcolvCall({action: 'reply', threadId: m.threadId, text: reply + '\n\nJulián', html: html, draft: false});
      if (!r || !r.ok) { console.log('Email assistant: reply not sent for ' + addr + ': ' + ((r && r.error) || 'no answer')); return; }
      fubFetch('notes', {personId: p.id, subject: 'Email sent: ' + subject, body: 'They emailed: "' + String(m.body || '').slice(0, 600) + '"\n\nReplied from julianc@bcolv.com:\n' + reply});
      const msg = '📧 ' + who + ' emailed you: ' + said + '\n\nI answered for you:\n"' + reply + '"';
      if (boss) { bossTextSafe(msg); bossMemory('assistant', msg); }
    } else {  // default: tell him what it says and what she'd send; YES sends it in the same thread, or he says what to change
      const pn = addPending({kind: 'email', personId: p.id, who: who, email: addr, subject: subject, message: reply, threadId: m.threadId, at: Date.now()});
      stat('emailsIn');
      fubFetch('notes', {personId: p.id, subject: name + ' suggested reply', body: 'They emailed: "' + String(m.body || '').slice(0, 600) + '"\n\nSuggested reply (waiting for Julián):\n' + reply +
        (sensitive ? '\n\nNeeds you: ' + (out.reason || 'sensitive topic') : '')});
      const msg = '📧 ' + who + ' emailed you: ' + said + (sensitive ? '\n\n(Heads up: ' + (out.reason || 'this one needs you') + '.)' : '') +
        '\n\nHere\'s what I\'d send back:\n"' + reply + '"\n\n' + yesHint(pn) + ' to send it, or tell me what to change.';
      if (boss) { bossTextSafe(msg); bossMemory('assistant', msg); }
    }
    handled++;
  });
  return handled;
}

/** EMAIL OUTBOX (bcolv): a FUB note with subject "Send bcolv email" by Julián's account (user 2, e.g. Claude) is
 *  emailed to that contact FROM julianc@bcolv.com through the Shawna Mailer. Body: first line "Subject: …", optional
 *  line "File: <Google Drive file id>" (attached, e.g. a CMA PDF), then the email text. Attachments need Shawna Mailer
 *  v2; until Julián re-pastes it, the note waits (up to 36h) instead of sending without the file. */
function processBcolvEmailOutbox() {
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('BCOLV_OUTBOX_DONE') || '[]');
  const notes = (fubFetch('notes?limit=25&sort=-created', null, 'get').notes || [])
    .filter(n => /^send bcolv email$/i.test(String(n.subject || '').trim()) && n.createdById === 2 && done.indexOf(n.id) < 0 &&
      Date.now() - new Date(n.created) < 36 * 36e5);
  notes.reverse().forEach(n => {
    const fresh = fubFetch('notes/' + n.id, null, 'get');
    if (!/^send bcolv email$/i.test(String(fresh.subject || '').trim())) return;
    const lines = String(fresh.body || '').replace(/\r/g, '').split('\n');
    let subject = 'Following up', fileId = '', plain = false;
    while (lines.length && /^(subject|file|signature):/i.test(lines[0])) {
      const l = lines.shift();
      if (/^subject:/i.test(l)) subject = l.replace(/^subject:\s*/i, '');
      else if (/^signature:/i.test(l)) plain = /none|no|off|plain/i.test(l.replace(/^signature:\s*/i, ''));  // personal letters: no "Julián" + Realtor footer
      else fileId = l.replace(/^file:\s*/i, '').trim();
    }
    const text = lines.join('\n').trim();
    let attachments;
    if (fileId) {
      const v = bcolvCall({action: 'version'});
      if (!v || !(v.version >= 2)) { console.log('bcolv email waiting: Shawna Mailer needs the v2 update for attachments'); return; }
      const f = DriveApp.getFileById(fileId), b = f.getBlob();
      attachments = [{name: f.getName(), mimeType: b.getContentType(), base64: Utilities.base64Encode(b.getBytes())}];
    }
    done.push(n.id); props.setProperty('BCOLV_OUTBOX_DONE', JSON.stringify(done.slice(-200)));  // mark first: never double-send
    try {
      const p = fubFetch('people/' + n.personId + '?fields=id,firstName,emails', null, 'get');
      const to = (p.emails || []).map(e => e.value).filter(validEmail)[0];
      if (!to) throw new Error('no valid email');
      const r = sendFromBcolv(to, subject, text, attachments, plain);
      if (!r || !r.ok) throw new Error((r && r.error) || 'mailer did not answer');
      if (attachments && r.attachments !== attachments.length) throw new Error('sent, but the attachment did not go');
      fubFetch('notes/' + n.id, {subject: 'Send bcolv email - sent'}, 'put');
      fubFetch('notes', {personId: p.id, subject: 'Email sent: ' + subject, body: 'From julianc@bcolv.com to ' + to + (attachments ? ' with ' + attachments[0].name : '') + ':\n\n' + text});
    } catch (err) {
      fubFetch('notes', {personId: n.personId, subject: 'Email NOT sent', body: err.message + '\n\n' + text});
    }
  });
}

/** Personal letters from Julián's own Gmail (Nov 12, his request for a family/HOA letter): a FUB note "Send personal
 *  email" by user 2, body "Subject: …" then the text. Sent as written (no sign-off, no Realtor footer), then the note is
 *  renamed "Send personal email - sent". Never for client business (those go from bcolv). */
function processPersonalEmailOutbox() {
  const props = PropertiesService.getScriptProperties(), done = JSON.parse(props.getProperty('PERSONAL_OUTBOX_DONE') || '[]');
  (fubFetch('notes?limit=25&sort=-created', null, 'get').notes || [])
    .filter(n => /^send personal email$/i.test(String(n.subject || '').trim()) && n.createdById === 2 && done.indexOf(n.id) < 0 && Date.now() - new Date(n.created) < 36 * 36e5)
    .forEach(n => {
      done.push(n.id); props.setProperty('PERSONAL_OUTBOX_DONE', JSON.stringify(done.slice(-100)));
      const lines = String(n.body || '').replace(/\r/g, '').split('\n');
      let subject = 'Hello';
      if (/^subject:/i.test(lines[0] || '')) subject = lines.shift().replace(/^subject:\s*/i, '');
      const text = lines.join('\n').trim();
      try {
        const p = fubFetch('people/' + n.personId + '?fields=id,emails', null, 'get');
        const to = (p.emails || []).map(e => e.value).filter(validEmail)[0];
        if (!to) throw new Error('no valid email');
        const html = '<div style="font:15px/1.6 Helvetica,Arial,sans-serif;color:#111">' + text.split(/\n{2,}/).map(par => '<p style="margin:0 0 12px">' +
          par.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>') + '</p>').join('') + '</div>';
        MailApp.sendEmail(to, subject, text, {htmlBody: html});
        fubFetch('notes/' + n.id, {subject: 'Send personal email - sent'}, 'put');
        fubFetch('notes', {personId: n.personId, subject: 'Email sent (personal Gmail): ' + subject, body: 'To ' + to + ' from Julián\'s personal Gmail.'});
      } catch (err) { fubFetch('notes', {personId: n.personId, subject: 'Email NOT sent', body: err.message}); }
    });
}

function processFubOutbox() {
  // Off unless Script Property EMAIL_OUTBOX = on: client emails must come from julianc@bcolv.com, not the script
  // owner's personal Gmail (Julián, Nov 8). Turn on only once it sends from bcolv.
  if (PropertiesService.getScriptProperties().getProperty('EMAIL_OUTBOX') === 'on') {
    try { processEmailOutbox(); } catch (e) { console.log('Email outbox: ' + e.message); }
  }
  const hour = Number(Utilities.formatDate(new Date(), 'America/Los_Angeles', 'H'));
  if (hour < 9 || hour >= 21) return;  // quiet hours: texts and emails wait until 9am
  try { processBcolvEmailOutbox(); } catch (e) { console.log('bcolv email outbox: ' + e.message); }
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('OUTBOX_DONE') || '[]');
  const notes = (fubFetch('notes?limit=25&sort=-created', null, 'get').notes || [])
    .filter(n => /^send text$/i.test(String(n.subject || '').trim()) && n.createdById === 2 && done.indexOf(n.id) < 0 &&
      Date.now() - new Date(n.created) < 36 * 36e5);
  notes.reverse().forEach(n => {
    done.push(n.id); props.setProperty('OUTBOX_DONE', JSON.stringify(done.slice(-200)));  // mark first: never double-send
    try {
      // Re-read the note right before sending: if it was canceled or edited a moment ago, respect that.
      const fresh = fubFetch('notes/' + n.id, null, 'get');
      if (!/^send text$/i.test(String(fresh.subject || '').trim())) return;
      n.body = fresh.body;
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
