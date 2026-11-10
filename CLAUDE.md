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
  Booking: `bookingSlots`/`bookSlot` (GET ?action=slots, POST action=book) read free time from his default Google
  Calendar (Mon–Sat 9–6 PT, 30 min, 3 h notice) and create the event + FUB note/task.
- Texting (bottom of Code.gs): BlueBubbles 1.9.7 on Mariel's MacBook Air (macOS 10.14, so newer versions won't run;
  auto-update off), Messages signed in as julian.crider@gmail.com / 702-498-2571, Cloudflare proxy (URL changes on
  restart; the `new-server` webhook updates BB_URL). Script Properties BB_URL, BB_PASSWORD, BB_KEY. `bbTextPerson` texts
  only `Text OK` (never `Text Opt-Out`) and logs to FUB textMessages; webhook logs replies from FUB contacts + one
  "They texted back" Text task per day; STOP → `Text Opt-Out`; non-FUB senders ignored. `setupTexting`, `testText`.
  Open-house texts: sign-in form shows a consent card under Phone ("By entering your phone number, you agree… Reply STOP")
  with a "Do not text me. Email only." option; phone given and not opted out → Text OK: Yes → tag `Text OK`; consenting guests
  (not Has Agent) get 2 texts via the "Text Queue" sheet tab (editable before send): +2h thank-you and +72h no-oriented
  question (OPEN_HOUSE_TEXTS guest: 2h "Hey… question || [contact card] || juliancrider.com/?relocate || [voice]"/24h/72h aiming for a 2nd showing; no street address; neighbor: 2h+voice/72h value;
  agent: 2h feedback). Form asks "What brings you in today?" (Looking to buy / Just looking / I live nearby / I'm a real
  estate agent → tags Agent+Partner, no welcome email or buyer drip). Voice note = Drive file "Quick hello.m4a" (name shows to recipients); contact card = emails/julian-crider.vcf (HD photo from img/julian-headshot-hd.jpg)
  (bbSendAudio, /message/attachment); if missing, the text still goes. processTextQueue every 15 min, 9am-9pm PT (TCPA: never after 9pm), retries ~6h, skips later
  steps once they've texted back. FUB rejects textMessages (its texting is off), so texts are logged as notes.
  Sign-in funnel: juliancrider.com/?signin (or ?openhouse; also a button on Homes > Open house) = same questions one per
  screen, not tied to a listing (open house optional). Every ?signin lead → source juliancrider.com + tag Website Lead (plus Open House if named). No open house → "Signed up: online" → tags Website Lead + Online
  Sign-In, source juliancrider.com, OPEN_HOUSE_TEXTS.online (first text ~3 min, no open-house voice note). Step 1 (name,
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
  RULES (Julián, Nov 8): client emails always go from julianc@bcolv.com, never julian.crider@gmail.com. Never tell
  Julián something was sent until it's verified (FUB GET emails?personId=… shows status Sent, or a "Text sent" note);
  an action plan reading "Completed" does NOT mean the email went out.
- AI assistant "Shawna" (`assistantReply` in Code.gs, called from bbWebhook): Groq (OpenAI-compatible API, model in
  GROQ_MODEL; on 404 groqPickModel lists his account's models and saves the best, since llama-3.3-70b-versatile was retired Nov 2026) drafts replies to inbound texts from FUB contacts. ASSISTANT_NAME
  (default Shawna; future users name their own), ASSISTANT_MODE draft (default: FUB note "Shawna draft" + task) / auto
  (sends 8am-9pm, max 3/person/day, only Text OK; first text = assistantIntro(): "Hi, it's Shawna, Julián's AI
  assistant. Julián's away… text JULIAN to reach him" (ASSISTANT_INTRO overrides); a client texting JULIAN → Julián alerted +
  task, Shawna stays out; every inbound client text alerts his personal phone (no "Shawna:" prefix); escalates offers,
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
  calendar or business = BUSINESS_CALENDAR_ID, default julianc@bcolv.com shared to his gmail; asks "personal or
  business?" if he didn't say; + FUB note, adds new people); calendar shows both, lookup (FUB summary), hot (Tier A list).
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
  note (summary, transcript, recording) + call-back task. No texts to callers (a call isn't text consent). Greeting
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
  no/bad email, past clients/sphere/trash, Has Agent. In each FUB automation, turn OFF stop-on-reply. "Funnel" sheet tab tracks stage. Text templates: JC Hot Text 1–3
  (31–33), JC Warm Text 1–2 (34–35), JC Nurture Text 1–8 (36–43). `juliancrider.com/?book` opens Book a consultation
  directly (link for Instagram/text auto-replies); site has no sms: links (they open the booking flow).
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
