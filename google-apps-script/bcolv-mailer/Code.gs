// Shawna's mailer: lives in Julián's julianc@bcolv.com Google account (script.google.com while signed in as bcolv),
// so emails Shawna sends come FROM julianc@bcolv.com and Follow Up Boss (connected to that inbox) tracks them.
// It also adds/reads appointments on the bcolv (business) calendar for Shawna.
// The main website script calls this web app with a secret key; nobody else can use it.
// Setup (one time):
//   1. script.google.com (signed in as julianc@bcolv.com) → New project → name it "Shawna Mailer" → paste this file.
//   2. Run setupMailer once (approve the Gmail permission). Copy the key it prints.
//   3. Deploy → New deployment → type "Web app" → Execute as: Me (julianc@bcolv.com), Who has access: Anyone → Deploy.
//      Copy the Web app URL.
//   4. In the MAIN script (JC Open House Leads) → Project Settings → Script Properties add:
//      BCOLV_MAILER_URL = the Web app URL, BCOLV_MAILER_KEY = the key from step 2.

function setupMailer() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('MAILER_KEY')) p.setProperty('MAILER_KEY', Utilities.getUuid().replace(/-/g, ''));
  GmailApp.getAliases();  // asks for the Gmail permission now
  CalendarApp.getDefaultCalendar();  // and the Calendar permission (business appointments)
  console.log('BCOLV_MAILER_KEY = ' + p.getProperty('MAILER_KEY'));
  console.log('Sending as: ' + Session.getActiveUser().getEmail());
}

function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let d; try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ok: false, error: 'bad request'}); }
  if (!d.key || d.key !== PropertiesService.getScriptProperties().getProperty('MAILER_KEY')) return out({ok: false, error: 'denied'});
  // Business calendar (Boeckle's Google won't let outside accounts edit it, so the main script asks this one).
  if (d.action === 'event') {
    const ev = CalendarApp.getDefaultCalendar().createEvent(String(d.title || 'Appointment'), new Date(d.start), new Date(d.end), {description: String(d.description || '')});
    return out({ok: true, id: ev.getId()});
  }
  if (d.action === 'events') {
    const evs = CalendarApp.getDefaultCalendar().getEvents(new Date(d.from), new Date(d.to))
      .map(e => ({title: e.getTitle(), start: e.getStartTime().getTime(), description: e.getDescription()}));
    return out({ok: true, events: evs});
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.to || ''))) return out({ok: false, error: 'bad address'});
  GmailApp.sendEmail(d.to, String(d.subject || '').slice(0, 200), String(d.text || ''),
    {htmlBody: d.html || undefined, name: d.name || 'Julián Crider'});
  return out({ok: true, from: Session.getEffectiveUser().getEmail()});
}
