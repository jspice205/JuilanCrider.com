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
  question (`askIntent`, shows until answered up to 3 visits; `?start` forces it), community map
  (Leaflet, loads on tap/consent), JSON-LD for SEO. Images live in `img/`.
- `sign.html` — on-site e-sign page for the single-property, non-exclusive buyer agreement +
  Nevada Duties Owed (form 525). Terms: 3% of purchase price, 30 days, live.
- `robots.txt`, `sitemap.xml`, `llms.txt` — SEO and AI-search files.
- `google-apps-script/Code.gs` + `appsscript.json` — pasted by Julián into the Apps Script of his
  "JC Open House Leads" Google Sheet, deployed as the web app the site posts to (`SHEET_URL`).
  Handles: website leads → sheet tabs + email + Follow Up Boss; mortgage rates; buyer agreements
  (PDF to Drive, emails); FUB Inbox auto-import from Google Drive; lead scoring; FUB Scorecard +
  Hot/Warm tagging (`scoreFubDatabase`, nightly via `setupNightlyScoring`; `activityScore` adds lead-side activity:
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
- `fub-connector/worker.js` — Follow Up Boss connector for Claude (Cloudflare Worker "fub-claude",
  MCP server). Tools: hot_list, search_people, get_person, update_person, add_note, log_call,
  create_task, list/start action plans, pipelines and deals, email templates, and `fub_api`
  (any FUB endpoint; DELETE needs confirm). Prefer `fub_api` over adding new tools, because a
  running thread doesn't see tools added after it starts. Julián updates the Worker by pasting code
  in Cloudflare → Edit code → Deploy.

## Follow Up Boss facts
- Tags: every scored contact has one of `Tier A` (Hot) / `Tier B` (Warm) / `Tier C` (Nurture or Cold), refreshed
  nightly. `Intake Tier A/B/C` = how strong their sign-up was (an input, never overwritten). `Hot`/`Warm` tags are
  never added automatically for Farm, Legacy or "Not DNC checked" contacts.
- Funnel (`runFunnel`, after nightly scoring), max 50/night. FUB moved Action Plans to Automations 2.0 (Oct 29)
  and the API can't start automations (403 on automations/automationsPeople), so the script adds a **trigger tag**
  and Julián's FUB Automation of the same name starts on "Tag added": Hot → `JC Hot`, Warm → `JC Warm`, Legacy →
  `JC Legacy Re-Intro`, Farm – Best → `JC Farm Best`, Tier C (not farm/legacy) → `JC Long-Term Nurture`; after
  25/65/25/60 days they move to Long-Term Nurture (old JC tag removed). `Text OK` people get "<tag> + Text" for stages
  in `FUNNEL.textLive`. Nothing is tagged for a stage until it's in `FUNNEL.live` — flip those in Code.gs when Julián
  says an automation is built (auto-update deploys it within the hour). Skips no/bad email, clients, Has Agent, and
  anyone who replied or was called in 30 days. "Funnel" sheet tab tracks stage. Text templates: JC Hot Text 1–3
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
