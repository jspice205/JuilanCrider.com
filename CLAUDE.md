# juliancrider.com — project guide for Claude

Owner: Julián J. Crider, Las Vegas / Henderson Realtor, license S.0177760, Boeckle & Co.
(broker Landon Boeckle, B.1001882; office 4775 W Teco Ave #220, Las Vegas NV 89118, 725-465-7646).
Julián's number for clients: 702-498-2571. Lead emails go to julianc@bcolv.com.
Julián is not a developer: explain steps click by click, in plain language, and keep all site and
Follow Up Boss work in one thread unless he starts a thread for a specific task.

## How changes go live
The site is served from `main`. Work on the session branch, then open a PR and squash-merge it
yourself (Julián approved this workflow). Never commit private contact data or API keys.

## What's here
- `index.html` — the whole site (single-page app). 4 tabs only: Feed, Homes, Relocate, Julián.
  **Don't add tabs**; use a hamburger/menu if more space is needed. Key pieces: `showTab()`,
  `openRaw()` sheets, `bindOpeners()`, Request a showing form, Book a consultation (`bookFlow`: screens pre-approval + agent status; signed-with-agent gets a polite no;
  then `pickTime` shows open slots from his calendar via the Apps Script), first-visit "What brings you here?"
  question (`askIntent`, shows until answered up to 3 visits; `?start` forces it).
  Mortgage rates + calculator live under Homes > Home value (`showValueRates`), not in the Feed. Book flow also asks Price range and
  "Where do you live now?"; `scoreLead` tags `Luxury Buyer` ($1.3M+), `Relocation` (lives outside NV) and
  `Luxury Relocation 60 Days` (both + Ready now/1–3 months → Hot). Julián's target: out-of-state $1.3M–$3.3M buyers;
  never target/exclude by sex, family status or age (Fair Housing), community map
  (Leaflet, loads on tap/consent), JSON-LD for SEO. Images live in `img/`.
- `sign.html` — on-site e-sign page for the single-property, non-exclusive buyer agreement +
  Nevada Duties Owed (form 525). Terms: 3% of purchase price, 30 days, live.
- `robots.txt`, `sitemap.xml`, `llms.txt` — SEO and AI-search files.
- `google-apps-script/Code.gs` + `appsscript.json` — pasted by Julián into the Apps Script of his
  "JC Open House Leads" Google Sheet, deployed as the web app the site posts to (`SHEET_URL`).
  Handles: website leads → sheet tabs + email + Follow Up Boss; mortgage rates; buyer agreements
  (PDF to Drive, emails); FUB Inbox auto-import from Google Drive; lead scoring; FUB Scorecard +
  Hot/Warm tagging (`trackEmailOpens` hourly: FUB emEvents → tags `Opened Email`/`Clicked Email` + "Email Opens" sheet tab;
  Smart List "Opened my emails" = tag Opened Email; `scoreFubDatabase`, nightly via `setupNightlyScoring`; `activityScore` adds lead-side activity:
  replies, calls in, DMs, new inquiries, site/IDX visits, email clicks, RealScout views, fading after 90 days, and
  -10 for "quiet: no reply after 3+ touches"); Monday digest; `loadFubTemplates` (Hot/Warm emails).
  Keys live in Script Properties (`FUB_API_KEY`, etc.), never in code.
  Self-updates hourly from GitHub `main` (`updateFromGitHub`/`setupAutoUpdate`: Apps Script API puts Code + manifest,
  makes a version, repoints the web-app deployment). If an update adds OAuth scopes, the hourly run emails Julián to
  run `updateFromGitHub` once by hand. Until he has run setupAutoUpdate, he pastes and redeploys manually.
  Oct 29: the installer hit 403 "Apps Script API has not been used in project 580055367952" (the script's hidden
  default GCP project). Self-update only works if that API can be enabled there or the script is moved to a standard
  GCP project; otherwise give Julián the raw-link paste steps (raw.githubusercontent.com/.../Code.gs, Ctrl+A/C,
  paste, Deploy > Manage deployments > New version).
  Booking: `bookingSlots(meetingFor(q))`/`bookSlot` (GET ?action=slots&kind=&minutes=&holes=, POST action=book). MEETINGS:
  phone + video Mon–Sat 9–6 (video 30 or 45 min, Google Meet link via Calendar REST API, VIDEO_LINK fallback), showing,
  golf by private link juliancrider.com/?golf only (Mon/Tue/Fri tee times 8–3, Thu 11–4; 9 holes 2.5 h / 18 holes 5 h;
  asks tee time booked + course). Busy = personal calendar + bcolv calendar (via mailer). Every booking: calendar event,
  FUB person (created if new), FUB appointment + note (+ task, not for golf), confirmation text; 15-min reminder from
  remindUpcoming. No coffee/lunch/in-person consults (Julián, Nov 2).
- Texting (bottom of Code.gs): BlueBubbles 1.9.7 on Mariel's MacBook Air (macOS 10.14, so newer versions won't run;
  auto-update off), Messages signed in as julian.crider@gmail.com / 702-498-2571, Cloudflare proxy (URL changes on
  restart; the `new-server` webhook updates BB_URL). Script Properties BB_URL, BB_PASSWORD, BB_KEY. `bbTextPerson` texts
  only `Text OK` (never `Text Opt-Out`) and logs to FUB textMessages; webhook logs replies from FUB contacts + one
  "They texted back" Text task per day; STOP → `Text Opt-Out`; non-FUB senders ignored. `setupTexting`, `testText`.
  Open-house texts: sign-in form shows a consent card under Phone ("By entering your phone number, you agree… Reply STOP")
  with a "Do not text me. Email only." option; phone given and not opted out → Text OK: Yes → tag `Text OK`; consenting guests
  (not Has Agent) get 2 texts via the "Text Queue" sheet tab (editable before send): +2h thank-you and +72h no-oriented
  question (OPEN_HOUSE_TEXTS guest: 2h "Hey… question || [contact card] || juliancrider.com/?relocate || [voice]"/24h/72h aiming for a 2nd showing, then a 90-day follow-up oh-d6/d10/d17/d35/d45/d60/d75/d90 (Joey's 36-touch arc rewritten in
  Julián's voice, approved Nov 10; stops when they reply); website forms (not bookings/showings) get OPEN_HOUSE_TEXTS.website
  (welcome + contact card, then +48h); no street address; neighbor: 2h+voice/72h value;
  agent: 2h feedback). Form asks "What brings you in today?" (Looking to buy / Just looking / I live nearby / I'm a real
  estate agent → tags Agent+Partner, no welcome email or buyer drip). Voice note = Drive file "Quick hello.m4a" (name shows to recipients); contact card = emails/julian-crider.vcf (HD photo from img/julian-headshot-hd.jpg)
  (bbSendAudio, /message/attachment); if missing, the text still goes. processTextQueue every 15 min, 9am-9pm PT (TCPA: never after 9pm), retries ~6h, skips later
  steps once they've texted back. FUB rejects textMessages (its texting is off), so texts are logged as notes.
  Sign-in funnel: juliancrider.com/?signin (or ?openhouse; also a button on Homes > Open house) = same questions one per
  screen, not tied to a listing (open house optional). Every ?signin lead → source juliancrider.com + tag Website Lead (plus Open House if named). No open house → "Signed up: online" → tags Website Lead + Online
  Sign-In, source juliancrider.com, OPEN_HOUSE_TEXTS.online (first text ~3 min: welcome + introduces Julián and the team, [welcome voice] = Drive file "Welcome hello"
  if he records one, no open-house voice note). Step 1 (name,
  phone, email, visitor type) signs them in, later answers post action 'enrich' (enrichSignIn: FUB note, re-tier tags, fills the
  Sign-Ins row, cancels queued texts if Has Agent). Notes carry "Form: funnel" / "Form: page" for the A/B answer-rate test.
  Appointments: bookSlot texts a confirmation; remindUpcoming (every 5 min) texts ~15 min before any Google Calendar
  event with a phone number in its notes/location/title ("Showing"/"Tour" in title → showing wording with location);
  one per event, skips Text Opt-Out, logs a FUB note. Triggers are ensured hourly by autoUpdateFromGitHub.
  Request a showing (button "Next: pick a time") with an address (not signed with an agent) then opens pickTime; bookSlot with kind 'showing'
  creates "Showing: Name" (location = address, phone in notes) → confirmation text + 15-min reminder.
  Outbox: a FUB note with subject "Send text" by user 2 (Claude via fub_api) is texted to that contact within ~5 min
  (processFubOutbox, from remindUpcoming; 9am-9pm PT (TCPA: never after 9pm) only, notes up to 36h old), then logged "Text sent"/"Text NOT sent". Use this to text clients for Julián.
  Email outbox: note subject "Send email" by user 2, body "template: <id>" (+ optional "text: <follow-up text>") →
  processEmailOutbox emails that FUB template from his Gmail (MailApp, adds EMAIL_FOOTER), logs "Email sent", then queues
  the text. OFF since Nov 8 (needs Script Property EMAIL_OUTBOX=on): it sent from his personal Gmail.
  Files by text: a bubble "[file: <Drive file id or name>]" sends that Drive file (e.g. a CMA PDF) as an attachment.
  bcolv email outbox: note subject "Send bcolv email" by user 2, body "Subject: …" / optional "File: <Drive id>" / text →
  processBcolvEmailOutbox emails FROM julianc@bcolv.com (attachments need Shawna Mailer v2; the note waits until then),
  logs "Email sent: …" and renames the note "Send bcolv email - sent". A "Signature: none" header line sends it without the "Julián" + Realtor footer (personal letters, Nov 12).
  Personal letters from his own Gmail (his explicit ask, Nov 12): FUB note "Send personal email" by user 2 ("Subject: …" +
  text) → processPersonalEmailOutbox (remindUpcoming tick, MailApp, no sign-off/footer) → "Send personal email - sent". Never for client business. Web app GET ?action=version shows CODE_VERSION
  (check it before relying on new code). CMAs: reports/cma-template.html + render-cma.js; client data stays local.
  RULES (Julián, Nov 8): client emails always go from julianc@bcolv.com, never julian.crider@gmail.com. Never tell
  Julián something was sent until it's verified (FUB GET emails?personId=… shows status Sent, or a "Text sent" note);
  an action plan reading "Completed" does NOT mean the email went out.
- AI assistant "Shawna" (`assistantReply` in Code.gs, called from bbWebhook): Groq (OpenAI-compatible API, model in
  GROQ_MODEL; on 404 groqPickModel lists his account's models and saves the best, since llama-3.3-70b-versatile was retired Nov 2026) drafts replies to inbound texts from FUB contacts. ASSISTANT_NAME
  (default Shawna; future users name their own), ASSISTANT_MODE draft (default: FUB note "Shawna draft" + task) / auto
  (replies 8am-10pm (ASSISTANT_END_HOUR 22; they texted first), scheduled/outbox texts still 9am-9pm; first text = assistantIntro(): "Hi, it's Shawna, Julián's AI
  assistant. Julián's away… text JULIAN to reach him" (ASSISTANT_INTRO overrides); a client texting JULIAN → Julián alerted +
  task, Shawna stays out; client-text alerts only when TEXT_ALERTS=on (off by default since Nov 9; "who texted me?" → TEXT_LOG; texts to him carry no
  FUB links, which open a Zillow Pro sign-in); escalates offers,
  contracts, commission, legal, complaints, Has Agent) / backup (Julián's choice Nov 9: reply held in AI_PENDING,
  sent by processAssistantQueue (1-min trigger) only if he hasn't texted that person within ASSISTANT_DELAY_MIN,
  default 2; his outgoing texts (BlueBubbles isFromMe) cancel it and mute Shawna for that person for ASSISTANT_MUTE_MIN,
  default 30 (AI_MUTE_<phone>; her own sends are marked AI_SENDING_<phone> so they don't count)) / off. Needs GROQ_API_KEY; `testAssistant`.
  Voice rules (Julián): human + Chris Voss, OK to say no / "this isn't for everyone"; "here's what I can do for you",
  "when we work together", "I'll confirm that for you", "I was thinking of you", "the next steps are".
  Email alerts: `checkClientEmails` (5-min tick) pushes "X emailed you" via ntfy (NTFY_TOPIC) for FUB emails with
  status Received, plus one nudge after EMAIL_NUDGE_MIN (60) with no Sent reply; 7am-10pm. notifyJulian also texts his
  personal phone (ALERT_TEXT_TO, else TEST_TEXT_TO) from the business number, and tells him when Shawna answered for him. FUB hides email bodies,
  so answering emails would need a script inside the bcolv Google account.
- Shawna knows the business from SHAWNA_FACTS in Code.gs + one fact per row in the "Shawna Knowledge" sheet tab
  (assistantKnowledge); humanize() strips dashes from everything she writes. Boss actions also: schedule (personal = his gmail
  calendar or business = bcolv calendar via the Shawna Mailer (bcolvCall action event/events; Boeckle's Workspace
  blocks outside edit sharing, Nov 9); asks "personal or
  business?" if he didn't say; + FUB note, adds new people); calendar shows both, lookup (FUB summary), hot (Tier A list), voicememo (shawnaVoice:
  ElevenLabs TTS, ELEVENLABS_API_KEY + ELEVENLABS_VOICE_ID = her Vapi voice; Reply YES first; text bubble then MP3; testVoiceMemo), callme
  (vapiCall: Vapi POST /call to Julián only; VAPI_API_KEY, VAPI_ASSISTANT_ID, VAPI_PHONE_NUMBER_ID; no AI calls to clients without consent).
- Nov 11 fixes: no "No Name" contacts. Unknown numbers (not in FUB, not in his iPhone contacts) → `newNumber`: Shawna
  replies (they texted first), asks their name once, max 3 replies; added to FUB (tag Texted In) only once Groq hears
  their name; "New Numbers" sheet tab. `realFirst` never greets anyone as "No". Shawna reads the real text thread from the
  Mac (`bbThread`, BlueBubbles GET chat/<guid>/message) + 15 FUB notes (`personHistory`). Replies go back on the chat
  they texted from (BB_CHAT_<number>, so Google Voice/Android get SMS, not iMessage). Groq JSON also returns name +
  intent: `learnFromReply` saves a missing name; `applyIntent` tags `Not Now` (funnel moves them to Long-Term Nurture, no
  Hot/Warm) or `Not Interested` (drip paused). Vapi calls: Groq pulls the caller's name; new callers without one are
  "Caller 702-…", never No Name. Email assistant `processEmailAssistant` (5-min tick, 7am-10pm, Shawna Mailer v3
  actions inbox/reply): unread emails from FUB contacts → reply as Julián; EMAIL_ASSISTANT draft (default, Gmail draft +
  text to Julián) / auto / off; sensitive ones only drafted. Numbers saved in his iPhone contacts stay ignored (personal).
- Vapi sync (Nov 11): `syncVapiAssistant` (hourly from autoUpdateFromGitHub, only when changed; `updateVapiNow` by hand)
  PATCHes the Vapi assistant: system prompt `vapiPrompt()` (built from SHAWNA_FACTS + Shawna Knowledge: one question at a
  time, never ask for the phone number, rates → lender/Julián, Chris Voss), tool `lookup_caller` (server = web app ?vapi=,
  vapiWebhook answers `tool-calls` with `vapiLookup`: FUB name/stage/tags/last notes), endCallMessage, startSpeakingPlan
  waitSeconds 0.8, multilingual (transcriber deepgram nova-3 'multi', ElevenLabs *_v2_5 voice; prompt: answer in the caller's
  current language; texts too). Needs VAPI_API_KEY + VAPI_ASSISTANT_ID. Saved-in-contacts numbers stay ignored unless
  Shawna was already talking with them (NEWNUM_ pending); FUB contacts always get Shawna whether saved or not. First texts to new numbers: if the iMessage is "Not
  Delivered" (Android/Google Voice), `bbDelivered` → resend as SMS; the working chat is saved in BB_CHAT_<number>.
  Test numbers: 2493 = Julián's Google Voice 702-763-1554, 2492 = friend; both tagged "Test - do not market".
- Settings tab (Nov 11): Script Properties UI is read-only past 50 entries, so Julián adds keys in the sheet's "Settings"
  tab (name in A, value in B); `applySettings` (5-min tick) saves them to Script Properties and erases the value.
  `cleanupProperties` (hourly) deletes old REM_/AI_MUTE_/AI_SENDING_/NEWNUM_ entries.
- Memory (Nov 11): BOSS_HISTORY 14 days / 30 msgs (<8.5 KB). `personTimeline` = texts (bbThread, falls back to BlueBubbles
  message/query) + 2 weeks of Gmail both ways (`emailHistory`, Shawna Mailer v4 action search) + FUB calls/voice memos/
  notes; boss `read` summarizes it with Groq; client text replies also get their emails. Client auto replies are paced
  like a person (AI_PENDING, `humanDelay`: read + type + think time with jitter, halved in a live back-and-forth, max 120 s,
  plus the 1-min queue tick → ~10 s to 3 min; ASSISTANT_INSTANT=on to send at once); Julián's own commands stay instant.
  `unanswered` skips threads whose last message is from Julián/Shawna, test contacts and "No name".
- Email briefs (Nov 11): processEmailAssistant (unread inbox + FUB "Received" via checkClientEmails) no longer makes
  Gmail drafts or "X emailed you" alerts: it texts Julián "📧 X emailed you: <summary>" + the reply Shawna would send;
  BOSS_PENDING kind email (threadId, valid 6 h) → YES sends in the thread; "change it…" rewrites. EMAIL_ASSISTANT=auto
  sends non-sensitive replies and tells him after. FUB note "Shawna suggested reply".
- Email threads (Nov 11): Shawna Mailer v5 `reply` answers the client's latest message in the thread (reply), or reply-all
  to Julián's latest when the thread is only his (campaign) — never to himself. Boss `email` replies in the latest thread
  with them from 60 days ({from: to:} search); with mailer <5 only threads whose latest message is theirs.
- Campaign replies (Nov 11): checkClientEmails sees FUB "Received" emails (FUB hides bodies) → mailer search from:<their
  email> → processEmailAssistant(list) drafts/intent even if Julián already opened it. ask/read/unanswered send an
  instant "On it 👀 one sec." before the slower answer. FUB sends drips through a connected Gmail (email account 683781,
  gmail threads since Nov 10); replies arrive there.
- Boss action `ask` (Nov 11): answers any question from all sources at once: personTimeline for named people, FUB
  notes feed (calls/texts/emails/notes, 40 newest) + newest contacts, BlueBubbles chat/query latest threads, Gmail
  search (Groq writes the query; mailer v4), Website Leads + Sign-Ins tabs, both calendars (7 days) → Groq answer.
  Alerts that matter: client texts Groq marks important/escalate → notifyJulian (once per person per 2 h, even with
  TEXT_ALERTS off); every real Vapi caller → notifyJulian with the summary.
- Scheduled texts (Nov 11): boss text with "when" (Friday / next week / in 2 weeks; default 10:00, nudged into 9am-9pm) →
  approval shows the time; YES → Text Queue row step "scheduled" (editable in the tab), sent by processTextQueue via
  `sendApprovedText` (Julián-approved: respects Text Opt-Out, not the Text OK tag). "cancel the text to X" → unschedule.
  agenda lists scheduled texts.
- Multi-requests (Nov 11): routing may return {"actions": [...]} (≤5); bossCommand runs `doAction` for each in order (each
  answers separately; approvals get numbered). Routing maxTokens 1000.
- groqChat(key, model, messages, {maxTokens, noPick}) (Nov 11): default max_tokens 900 (routing 700; Groq free tier counts
  requested max_tokens against tokens-per-minute), waits Groq's "try again in Xs" up to 20 s; routing model 404 →
  llama-3.1-8b-instant saved as GROQ_ROUTER_MODEL; "give me a minute (<reason>)" names the cause; ask/read context capped
  at 12k/8k chars. Earlier: reasoning_effort low for gpt-oss (thinking counts toward tokens; 400 cut off JSON);
  on 400 json_validate_failed it uses error.failed_generation if it parses, else retries once.
- Boss routing model (Nov 11): GROQ_ROUTER_MODEL (default openai/gpt-oss-20b, separate Groq rate bucket); main model is the
  retry. "give me a minute" = Groq free-tier rate limits.
- Boss reliability (Nov 11): routing prompt is slim (no SHAWNA_FACTS; last 12 history msgs ≤250 chars); groqChat waits and
  retries once on 429/5xx; if both routing tries fail she says "give me a minute…" once and throws so runBossCommand retries
  (3x). Boss action `agenda` (today/week): both calendars + FUB open tasks (tasks?assignedUserId&isCompleted=false) +
  waiting texts + pending email approval.
- Assistant upgrades (Nov 11): approvals queue BOSS_QUEUE (numbered #n, 6 h; plain YES = newest; "YES 2"/"NO 2");
  `addPending` everywhere (text, email, voice, deal, event). `agendaText(week)` shared by agenda + `morningBriefing`
  (8am daily trigger; MORNING_BRIEFING=off; adds overnight email opens + duplicate count). `weeklyScorecard` Fridays 9am
  (text + email; counters via `stat()` in STAT_<week>: calls, booked, shawnaTexts, neededYou, emailsIn, approved).
  Deal signals: Groq "deal" (offer/under_contract/closed/listing) from texts, emails, calls → `suggestDeal` (YES → `moveDeal`
  moves/creates the FUB deal + contact stage). Vapi bookings: appointmentISO checked against both calendars → FUB
  appointment if found, else pending kind event (YES adds it). `findDuplicates` nightly → "Duplicates" tab. Boss prompt
  "BE A REAL ASSISTANT": never stuck, closest action, one short question if needed.
- Rich texts (Nov 11): website bookings send a tap-to-add calendar invite (`icsInvite`, .ics with 30-min alarm) under the
  confirmation; any text bubble with a listing link (zillow/redfin/realtor/realscout/…) gets its og:image sent as a photo
  bubble (`listingPhoto`, cached 6 h; blocked sites just send the link). Julián rejected numbered quick replies (too old
  school). Tapbacks/typing/effects need a newer Mac + BlueBubbles Private API.
- Tasks from the text thread (Nov 11): agenda lists up to 10 open FUB tasks numbered (AGENDA_TASKS); "do 2" → taskstart
  (Groq drafts the text from personTimeline → BOSS_PENDING with taskId; YES sends and checks the task off; call tasks get
  a briefing + optional text); "done 2" / "clear the old tasks" → taskdone. `cleanupOldTasks` (daily, hourly hook)
  completes automatic reminder tasks older than 7 days. Groq falls back to GROQ_FALLBACK (default llama-3.1-8b-instant)
  when the main model keeps failing.
- Boss prompt has an "understand him like a human" section: casual/urban/voice-to-text phrasings mapped to actions
  (e.g. "who I ain't hit back" → unanswered, "bet"/"run it" → yes). Add new phrasings there when he reports a miss.
- Nov 11 (later): boss commands parse more robustly (retry with short history; "yes please" repeats the offered action).
  New boss actions: read (latest texts via bbThread + latest unread email via mailer inbox or FUB email notes),
  unanswered (TEXT_LOG + unread inbox), allow (`allowShawna`: names → numbers from FUB or his phone contacts →
  SHAWNA_ALLOW). Allowed numbers / "Hey Shawna" threads get `shawnaDirectReply` (no FUB entry). "Hey Shawna" in a text
  Julián sends anyone → SHAWNA_ON_<number>, Shawna joins that thread as his AI assistant. Claude can allow people by a
  FUB note on 2221 with subject "Shawna allow" (one name per line; `processAllowNotes` hourly → "Shawna allow - done").
- Live info (Nov 11): `liveInfo(text)` adds today's rates (`ratesText`, same feed as the site card) when a text mentions
  rates/loans (RATE_Q), and a web answer (`webAnswer`: Groq GROQ_WEB_MODEL, default groq/compound-mini, cached 1 h) for
  current things (WEB_Q: market, news, sports, events…) to client replies, "Hey Shawna" chats and boss action `research`.
  Prompt "WELL ROUNDED": explains loans/process, chats about the world; current numbers only from LIVE INFO. Vapi prompt
  includes today's averages (never their own rate). "Hey Shawna" (also hi/yo, shauna/shwna; not "Shana") answers
  instantly at any hour, and client replies in that thread skip pacing for 3 h (SHAWNA_ON_); if she can't, Julián is told.
  notifyJulian alerts go into BOSS_HISTORY + LAST_ALERT_WHO, so "what does the email say?" reads that person.
- Campaigns + tags (Nov 11): `personCampaigns(id)` = FUB action plans (running/paused/finished; names cached 6 h) + pending
  Text Queue rows; shown in personTimeline (read/ask), lookup ("what campaigns is Dan on") and client-reply context
  (CAMPAIGNS: don't repeat what a drip is about to say). FUB mass-email Campaigns aren't in the API. Groq JSON "tags"
  (SHAWNA_TAGS: moved_away, agent, referral_partner, relocation, seller, investor, has_agent, bought_elsewhere, renter,
  past_client) → `applyShawnaTags` from texts/emails/DMs + note "Shawna tagged: …". Boss action `tag` (add, "remove:").
  checkClientEmails ignores FUB "Received" copies of his own emails (relatedPeople sentByPerson false) and counts them
  as his reply. Margaret Anjewel 1970 (moved to Nashville): Referral Partner, Partner, Moved Away.
- Call notes (Nov 11, option 1): Julián records calls on his iPhone (native recording, both sides hear the notice),
  copies the transcript/summary from Notes and texts Shawna "call notes <name>" + paste (or the "Send call to Shawna"
  shortcut). Webhook `stashCallNotes` keeps long text whole in cache ([callnotes:key], 6 h; CALL_NOTES_RE); bossCommand skips
  routing → `callNotesParse` (Groq: who, summary, details, next, followup date/type, tags, intent, deal, text, email) →
  action `callnotes`: resolvePerson (creates new), FUB call log + "Call notes" note, follow-up task on the agreed day,
  applyShawnaTags/applyIntent/suggestDeal, pending follow-up text (long: 6 h to approve) + recap email from bcolv.
  Unknown/ambiguous name → asks; a short reply within 30 min (CALLNOTES_WAIT) names it.
- Saw-you-called (Nov 11): after a Vapi call from someone already in FUB with Text OK (no recap consent path), Text
  Queue step `call-followup` (~1 min, 9am-9pm): "Hey X, it's Julián. Saw you just called… good time to call you back?"
  Fast replies (Nov 11, Julián: never longer than 10 s): auto and backup modes sleep 3-8 s in the webhook and send (no
  1-min queue; skipped if he texted them meanwhile); ASSISTANT_FAST=off restores the old pacing/backup hold; humanDelay
  max 10 s. Self-texts on the business line (shortcut run on the business iPhone, 2571 → 2571) with "call notes" go to
  Shawna (deduped SELF_ cache); other self-texts are ignored. Vapi: when they ask for Julián, Shawna offers to answer
  their questions in the meantime.
  Vapi versions: syncVapiAssistant PATCH publishes (v35-37 on Nov 11); the dashboard "Draft" is an old local draft, never
  publish it (it has the old receptionist prompt).
- Nov 12 fixes (Lisa thread): STOP opt-out only when the WHOLE text is an opt-out word ("End of Dec you said?" had
  tagged Lisa Dragos 2223 Text Opt-Out; restored Text OK + note). Prompt: few emojis (not every text), paragraphs for
  longer replies, read the whole thread. `rememberJulianText`/`julianVoice` (JULIAN_VOICE, his last 40 texts to FUB
  contacts) → "HOW JULIÁN REALLY TEXTS" in assistantPrompt (persona julian). `detectMeeting` (his outgoing texts and
  client texts matching MEET_HINT; router model reads last 14 thread lines; agreed day+time → calendar event + FUB
  appointment unless one exists, MEET_<id>_<time> dedupe) → notifyJulian + pending kind `invite` (YES texts an .ics).
  Alerts: escalated/sensitive texts say "This one's for you. Want me to write the reply?" (not "I'm handling it");
  LAST_ALERT_ID makes "Lisa" resolve to the Lisa she just alerted about. Router text action "draft": true →
  `draftThreadReply` (whole bbThread + notes + campaigns; no price promises); replying to a text is a text, not email;
  router-invented emails/phones are scrubbed unless he typed them. `liveInfo` opens links (`pageText`, cached 1 h).
  His own texts to FUB contacts are logged as notes "Text from Julián" (Nov 12, his ask). Umar Tomlin 703 = his uncle:
  Trash + Deceased (Nov 12). Lisa Dragos: Onyx Crest walkthrough with Tom, Sat Nov 14 9:30am (FUB appointment 3102).
  Calendar sync (Nov 12): `syncCalendars` (remindUpcoming tick, every 15 min) copies the next 30 days of bcolv calendar
  events (mailer action events) and his FUB appointments onto his personal Google Calendar ("[synced by Shawna]" in the
  description; CALSYNC map; skips ones already there; removes copies when the source moves/cancels). Synced copies get no
  15-min client reminder texts.
  Showing flow (Nov 12, `showingFlow` in index.html): step 1 pick a day/time from his calendar (slots kind showing) or
  "Not sure yet, after talking with Julián"; step 2 home (optional), pre-approval, agent, contact → sendLead + bookSlot
  (property "Home to be chosen with Julián" if blank). The one-home agreement + Duties Owed is emailed even without an
  address: property AGREEMENT_TBD = review copy (sign.html shows it read-only; signAgreement refuses it). Every website
  lead → notifyJulian "🌐 <type>: <name>"; every site booking → "📅 Booked on your site". Shawna: no "Hi Unknown" in voice
  memos; "Hey Shawna" threads expire after 3 h; friends/family threads: she stays out 6 h after Julián texts them
  (AI_LASTME_) and never pitches calls/real estate; bbSend sends one message at a time per person (BUSY_ cache) so
  bubbles don't cross. Call notes: the shortcut can send "call notes <Notes title>\n<body>" after "Add Transcript to Note";
  "Call with X" titles become the name, phone-number titles the phone; the same notes twice are processed once (CNDONE_).
  Audio (Nov 12): bbWebhook transcribes audio attachments (`transcribeAudio`: BlueBubbles attachment/<guid>/download →
  Groq GROQ_WHISPER_MODEL, default whisper-large-v3-turbo; AUD_<guid> dedupe). From Julián/his business line: "Call
  with X.m4a" or >600 chars → call notes (name from the file name), else a spoken boss command. From clients: text
  becomes "(voice message) <words>" and flows as usual. iPhone voice bubbles are .caf; if Whisper rejects one, the client
  text says it couldn't be played. A spoken request from Julián (BOSS_VOICE cache, 2 min) makes `say` also send a
  voice memo of her answer (shawnaVoice, ElevenLabs) under the text. Notes can't add transcripts or share audio automatically (no Shortcuts action).
  To-dos vs calendar (Nov 12): boss `agenda` = `todoText(range)` (FUB tasks due today incl. past due / tomorrow / week,
  waiting texts, scheduled texts, approvals; AGENDA_TASKS keeps due dates); `calendar` = `calendarText(range)` (personal +
  bcolv, Shawna's synced copies skipped, de-duplicated). `agendaRange`: today unless "tomorrow"/"week". morningBriefing
  uses both (agendaText). "clear all"/"read all"/"delete all"/"clear everything" → past-due + today's tasks completed and
  waiting texts marked read (TEXT_LOG read flag); future tasks only one at a time; "no all"/"cancel all" empties the
  approvals queue; pendingLabel never says "undefined".
  Voice in → voice only out (Nov 12): with BOSS_VOICE, `say` sends just the voice memo (text if ElevenLabs fails) and
  keeps LAST_VOICE_TEXT; "transcript" / "text me that" / "text me back your voice memo" sends the written version.
  Already done: "I already called Dan" / "Lisa's handled" → taskdone for that person (their tasks due today or earlier,
  else their next one, even if not on the numbered list); "that's done" → LAST_TASK (the one "do N" started).
  Own numbers (Nov 12): `isOwnNumber` (business 2571, personal 1382, Google Voice 1554, ALERT/TEST_TEXT_TO, OWN_NUMBERS)
  never count as "waiting on you" (textLog skips them; agenda/unanswered filter them); his own reply to someone marks
  their waiting texts read (`textLogAnswered`); approvals addressed to the business number itself are dropped.
  "done N"/"do N" (Nov 12): handled without the router (regex fast path). Numbers come from his text; each is checked
  against open FUB tasks; a number not on the current list (already done, or a stale list) → Shawna shows the real
  open list, renumbered, instead of "which ones?". A view with no tasks never wipes AGENDA_TASKS.
  Folders (Nov 12): Julián's personal notebooks as sheet tabs (NOTE_FOLDERS: Golf Notes, Links [When, Link, Title,
  Note; title from og:title/<title>], SocialPod Notes, Personal Notes; any other name → "<Name> Notes"). Boss actions
  `jot` (save) and `notes` (read latest 8 or answer a question over the folder with Groq). Fast paths: a bare link →
  Links; "golf:/swing thought:", "socialpod:", "note to self:/personal:/jot:/write this down:" prefixes. `remember` stays
  for business facts Shawna uses (Shawna Knowledge tab).
  Nov 12 (Joey test call): vapiWebhook → `vapiEndOfCall`; any error still texts Julián "📞 A call came in" + Vapi's summary.
  Vapi prompt: email optional (asked once; "just text me" = never ask again, that's text consent), NEVER GET STUCK (any
  detail at most twice, then "Julián will confirm by text"), addresses read back once then spelled/"address to confirm".
  Summary textOk also true when the caller asked to be texted. `bossTextSafe`: texts to Julián that fail (Mac asleep)
  go to BOSS_OUTBOX and processAssistantQueue retries every minute for 6 h (notifyJulian + email briefs use it). Email
  assistant skips automatic replies (out of office, "thank you for contacting us", info@/support@); Victor Cohen 731's
  wrong Partner tags removed. Boss `emails`: one line per person (inbox, last 2 days or week, no system mail, nothing he
  sent) with a ≤12-word summary. Folders: Meeting Notes (weekly meetings with Joey) and Listings [When, Address, Status
  (Coming soon / Renovating / Active / Pending / Sold), Notes]; jotting an address already listed updates its status;
  prefixes "meeting:" and "listing:/coming soon:/renovating:".
  Team calls (Nov 12, Joey's test call landed in Joey's FUB and no alert came): vapiEndOfCall → callers whose number is a
  FUB user's (`teamNumbers`: users API phones + TEAM_NUMBERS setting, cached 6 h) get only a "📞 Team call" text, nothing
  in FUB; a matched/created contact that isn't Julián's (`isMine`: assignedUserId ≠ me or unreadable) gets a text only, no
  note/task; new phone leads are created with assignedUserId = Julián; createCallTask always assigns Julián. FUB users:
  Joey 1 (702-300-4349), Julián 2, Tana 3 (702-622-5232).
  Team assistant (Nov 12, Julián: Shawna is his assistant for the team too): `teamDirectory` {phone: name} (FUB users except
  Julián + TEAM_NUMBERS "7025551234 Joey, …"); vapiLookup tells Shawna "TEAM MEMBER: <name>" (take meetings/open houses/
  showings/dinners for Julián: what, day, time, place; no buyer questions, no email, no recap text). After a team call the
  requested appointment (summary appointment/appointmentISO now include requests) goes straight on his personal calendar
  and he gets "📞 <name> called" + summary + "📅 Added to your calendar". Nothing in FUB.
  Every caller gets one text (Nov 12, Julián's ask): after any non-team call, Text Queue step `call-thanks` (~1 min,
  9am-9pm, sent via sendApprovedText: respects Text Opt-Out, no Text OK needed, one-time informational reply to someone who
  called, ends "Reply STOP to opt out"); callers who said yes to a recap still get the recap (and Text OK). Joey's Nov 12
  call: Shawna went silent after he spelled his email (audio: he said "Hello" twice, ~40 s of nothing) = the AI model
  stopped answering mid-call, most likely Groq free-tier rate limits on a long prompt; recommended moving Vapi's model to a
  paid provider (e.g. OpenAI gpt-4.1-mini in Vapi) or a paid Groq tier.
  syncVapiAssistant now GETs the assistant hourly and re-pushes the prompt + lookup_caller tool whenever Vapi's live prompt
  differs (e.g. an old dashboard draft got published); the model/provider Julián picks in Vapi is kept (Nov 12: he moved
  the phone brain off Groq to a Vapi-billed model).
  Vapi calls from his own phone (boss) now text him "🧪 Test call logged" + the recap a caller would get.
- iMessage from an Apple ID email (no phone): webhook matches FUB by email and Shawna replies to that address (bbSend
  accepts emails). Toll-free senders (8xx) are never added as leads.
- Open to everyone (Nov 9): any number texting the business line gets Shawna (no Text OK needed, they texted first).
  Not in FUB → added (source "Text message", tag "Texted In") unless saved in his iPhone contacts (macContactByPhone,
  cached 6h; personal people are left alone) or SHAWNA_NEW_NUMBERS=off. Tag `No Shawna` opts a contact out. Up to
  ASSISTANT_MAX_PER_DAY (12) replies per person per day.
- ASSISTANT_PERSONA = julian (default, Nov 9): Shawna's client replies are written AS Julián, first person, no intro; if a
  client sincerely asks if it's AI she says yes (never denies). shawna = replies as his assistant with the intro. California
  numbers (CA_AREA_CODES) always get the shawna version + intro (CA bot disclosure law, Julián's choice Nov 9). Prompt has
  human/Chris Voss texting examples, catch up first, 1-2 emojis; temperature 0.7.
- Shawna remembers: BOSS_HISTORY (last 16 messages with Julián, 3 days) is sent to Groq with each command; "remember …" adds
  a row to the Shawna Knowledge tab. Calls: no summary from Vapi → Groq writes one; Julián's own test calls get no task.
- Julián's texts to Shawna go through runBossCommand: script lock, "Shawna Log" sheet tab (in/ok/error), failures retried
  by processAssistantQueue up to 3 times (BOSS_RETRY), then he's told. Check that tab first when something "didn't respond".
- Talk to Shawna (`bossCommand`): texts from Julián's personal phone (ALERT_TEXT_TO / TEST_TEXT_TO) to the business
  number are commands: text a contact (always "Reply YES" first, BOSS_PENDING, 30 min), note, task, who emailed,
  today's calendar, add a contact (resolvePerson: a phone he gives is looked up first, then the name; nobody found → created right away
  via FUB events, tag "Added by Shawna", and the text/note/task/appointment still happens), or bigger work /
  anything she can't do → "Requests for Claude" sheet tab + FUB note subject "Claude request" on contact
  2221 (REQUESTS_PERSON_ID; Julián's test contact, phone = his personal phone). An hourly scheduled task (9am-8pm PT)
  wakes this Claude session: GET notes?personId=2221, do each "Claude request", PUT the note subject to
  "Claude request - done", and text him the result with a "Send text" note on 2221. Chat answers otherwise.
- Shawna email: `google-apps-script/bcolv-mailer/Code.gs` runs in Julián's julianc@bcolv.com Google account as a web app
  (BCOLV_MAILER_URL + BCOLV_MAILER_KEY in the main script); `sendFromBcolv` emails FROM bcolv (FUB tracks it via the
  connected inbox). bossCommand action "email" (Reply YES first). `testBcolvMailer`. Mac contacts: `macContact` searches
  his iPhone contacts synced to the texting Mac (BlueBubbles GET contact) when a name isn't in FUB and no number given.
- Shawna on the phone (plan, Nov 9): Vapi assistant (Groq model, Google Calendar tools, free Vapi number; ~$10-15/mo
  for ~90 min) answers forwarded missed calls (carrier no-answer forwarding). End-of-call report → `vapiWebhook`
  (?vapi=VAPI_KEY; `setupVoice` prints the URL) → FUB person by phone (or new, source "Phone call", tag Phone Lead) +
  note (summary, transcript, recording) + call-back task. Recap text (Nov 11): Shawna asks on the call "okay if I text you a
  quick recap?"; only a clear yes (Groq textOk) → tag Text OK + consent note + Text Queue step call-recap (recap + booked
  appointment), 9am-9pm. Otherwise no texts to callers (a call alone isn't text consent). Email briefs skip messages with
  no new words from the client (bounced copies of his own email). Greeting
  must say the call may be recorded (Nevada all-party consent, NRS 200.620). His own calls: iPhone native recording
  (both sides hear a notice; also ask verbally) → paste summary to Claude to log; FUB Calling add-on (~$39/mo, admin) later.
- `fub-connector/worker.js` — Follow Up Boss connector for Claude (Cloudflare Worker "fub-claude",
  MCP server). Tools: hot_list, search_people, get_person, update_person, add_note, log_call,
  create_task, list/start action plans, pipelines and deals, email templates, and `fub_api`
  (any FUB endpoint; DELETE needs confirm). Prefer `fub_api` over adding new tools, because a
  running thread doesn't see tools added after it starts. Julián updates the Worker by pasting code
  in Cloudflare → Edit code → Deploy.

## Follow Up Boss facts
- Deceased contacts: stage Trash + tag `Deceased` (Jared Weiss 437, Don Alves 471, Bishop Cooper 460; Nov 2).
- Shared brokerage account: Joey DeBenedetto (Broker, admin, user 1), Julián (Agent, user 2), Tana Cogan (Agent, 3).
  Julián can't create Automations or enable texting (admin must submit texting registration); his email isn't
  connected to FUB yet. Scoring and funnel only touch contacts assigned to user 2 (checked via GET me).
- Tags: every scored contact has one of `Tier A` (Hot) / `Tier B` (Warm) / `Tier C` (Nurture or Cold), refreshed
  nightly. `Intake Tier A/B/C` = how strong their sign-up was (an input, never overwritten). `Hot`/`Warm` tags are
  never added automatically for Farm, Legacy or "Not DNC checked" contacts.
- Funnel (`runFunnel`, after daily scoring at ~7am PT so emails land in the morning), max 50/day. The API can't touch Automations 2.0 (403), but on Oct 31
  Claude found it can still create/start **Action Plans**: JC Hot 22, JC Warm 23, JC Legacy Re-Intro 24, JC Farm Best 25,
  JC Long-Term Nurture 26 (FUNNEL.plans; stopOnContacted false). API-started plans begin Paused, so the script PUTs
  status Running; it pauses the old JC plan on stage moves/pauses. Tested: Hot 1 delivered to test contact 2221 (Trash).
  Welcome email: open-house sign-ins and website account/guide sign-ups get FUB plan "JC Welcome (sign-in)" (29,
  HTML template 183, emails/welcome.html) right away; the funnel waits 2 days after `created` before starting the drip.
  Re-intro: contacts in FUB 30+ days with no funnel stage (not Farm) first get plan "JC Re-Intro" (30, HTML template 184,
  emails/reintro.html, "It's been a while"); Funnel tab records `reintro-<stage>` and their real drip starts 4+ days later.
  "Old" also = created before LIVE_SINCE (Oct 28-29 bulk import; FUB created = import date). Bug on Nov 1: first run
  skipped re-intro for 50 people (Hot 1/Warm 1/Legacy 1 went first). FUNNEL.rerun sends those 46 (not Emilee 304,
  bounced 267/422/427) the re-intro on Nov 3 and their paused plans resume at email 2 after 4 days (start() resumes a
  paused plan for the stage instead of restarting). Legacy after re-intro uses plan 33 (Legacy 2-4).
  `juliancrider.com/?showing` opens Request a showing.
  Farm – Best get "JC Neighbor Note" (31, template 185, emails/neighbor-note.html) first instead; plan 25 now starts at
  Farm Best 2 (template 177 retired). Open-house form asks "Do you live in this neighborhood?"; Yes → tags `Open House
  Neighbor` + `Likely Seller`, plan "JC Neighbor Welcome" (32, template 186, emails/welcome-neighbor.html) instead of 29,
  then the Farm drip (no Neighbor Note). `juliancrider.com/?value` opens the home-value form.
  Plain-text drip templates 141-180 end with "<p>Julián</p>" + EMAIL_FOOTER (name, title, brokerage, phone, site,
  Instagram, address, license, reply-stop); applyEmailFooter runs once per EMAIL_FOOTER_VERSION from the hourly update.
  When rewriting a template, keep the <!--jc-footer--> part (or bump the version to re-apply).
  Newsletter/content-drop sends: plans 27/28 (templates 181/182, emails/*.html). Contacts tagged "Test…" are skipped.
  The JC tags are still added so Smart Lists show the stage: Hot → `JC Hot`, Warm → `JC Warm`, Legacy →
  `JC Legacy Re-Intro`, Farm – Best → `JC Farm Best`, Tier C (not farm/legacy) → `JC Long-Term Nurture`; after
  25/65/25/60 days they move to Long-Term Nurture (old JC tag removed). `Text OK` people get "<tag> + Text" for stages
  in `FUNNEL.textLive`. Nothing is tagged for a stage until it's in `FUNNEL.live` — flip those in Code.gs when Julián
  says an automation is built (auto-update deploys it within the hour). Replies never stop it. Pause = stage Active Client/Pending/Under
  Contract/Closed or tag `Pause` → JC tag removed; when that clears they restart in Long-Term Nurture. Never starts
  no/bad email, `Bounced` (unless they also have `Opened Email`; bounced people's JC plans are paused once, BOUNCE_CLEANED;
  trackEmailOpens tags bounces), past clients/sphere/trash, Has Agent. In each FUB automation, turn OFF stop-on-reply. "Funnel" sheet tab tracks stage. Text templates: JC Hot Text 1–3
  (31–33), JC Warm Text 1–2 (34–35), JC Nurture Text 1–8 (36–43). `juliancrider.com/?book` opens Book a consultation
  directly (link for Instagram/text auto-replies); site has no sms: links (they open the booking flow).
- Closed client loop (Nov 10, `runPastClientLoop`, after runFunnel each morning; PAST_CLIENTS in Code.gs): Julián only
  moves deal cards. `syncDealStages`: Hot Buyers/Listing Prep/Active Listings → contact stage Active Client; Offer/Due
  Diligence/Loan/Appraisal → Pending; CLOSED (25/26) → Closed (never backwards). Signed buyer agreement → deal card in
  Buyers › HOT BUYERS (`dealForAgreement`). Loop from closing date: day 3 text (email if no Text OK), day 7 email, day 30
  review ask (waits for Script Property REVIEW_LINK), day 60 referral ask, quarterly check-ins (4 rotating), yearly
  anniversary (FUB field customClosingAnniversary if set, also for older clients; filled in at closing); Closed → Past Client after day 30. Tag "past client" (older clients) → quarterly only, first 30-60 days out,
  JC drips paused, funnel skips them. Emails via sendFromBcolv (bcolv), texts via Text Queue (Text OK only). "Past Clients"
  sheet tab; max 15/day; skip tags Pause / No Past Client Loop / Deceased / Has Agent. `previewPastClientLoop`.
- Instagram DMs (Nov 10): ManyChat flow collects email (+ optional phone/text consent) → External Request (ManyChat Pro)
  POST to web app ?ig=IG_KEY → `instagramWebhook`: FUB event source "Instagram", tag "Instagram DM", note, one Follow Up
  task, welcome email from bcolv with juliancrider.com/?book; texts only if body has "text_ok": "yes". `setupInstagram`
  prints URL + JSON body. Free-plan fallback: ManyChat replies with juliancrider.com/?signin (Julián set this up Nov 11).
  `igReply`: the webhook also returns {reply} = Shawna's answer (as Julián, IG_<handle> keeps 8 msgs) for ManyChat Pro to
  send back (map $.reply); works without email/phone; "source": "comment" for comment-to-DM automations.
- His API key cannot create pipelines (403). FUB's API can't create action plans, automations,
  smart lists, or send texts/emails, so give Julián click-by-click steps for those.
- Pipelines: Buyers (Nurture, Hot Buyers, Offer, Due Diligence, Loan/Appraisal, Closed) and
  Sellers (9 stages). Plan: rename Buyers stages to the deal flow from his Asana board
  (Consult Set → Buyer Package + Duties Owed → Pre-Approved → Actively Looking → Offer Submitted →
  In Escrow → Closed).
- Inbound-only rule (Julián's choice, Oct 29): plans do the reaching out; tasks are created only when a lead
  reaches out (website showing/consultation request, or activity score 25+: reply, call, DM, new inquiry, site visits).
  Intake score alone never creates a task. Don't add call lists or outreach tasks.
- Messaging style: Chris Voss (labels, no-oriented questions, calibrated how/what questions).
- Texting only to people who gave consent. Sahara Investment Group (SIG) outreach needs
  compliance approval first.

## Open to-dos (as of Oct 2026)
1. Templates loaded: Hot/Warm/Legacy 1–4 (141–152), JC Nurture 01–24, JC Farm Best 1–4. Julián builds the JC
   Automations in FUB (trigger: tag added; build sheet in his Drive); then set FUNNEL.live and the nightly funnel runs.
2. Texts only for `Text OK` (open house sign-in form has no text-consent line yet); SIG content only after compliance approval.
3. Smart list cleanup (7 lists: Call Today, New Not Contacted, Active on Site, Warm Check-in,
   Likely Sellers, Under Contract, Past Clients & Sphere) and pause unused starter plans.
4. Apps Script: make sure the latest Code.gs + manifest are pasted and redeployed; run
   setupFubInbox, setupWeeklyDigest, previewAgreement.
5. Later: agent network list, SIG/Apollo (after compliance), social posting, Meta audiences.
