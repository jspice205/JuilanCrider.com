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
  `openRaw()` sheets, `bindOpeners()`, Request a showing form, first-visit "What brings you here?"
  question (`askIntent`, shows until answered up to 3 visits; `?start` forces it), community map
  (Leaflet, loads on tap/consent), JSON-LD for SEO. Images live in `img/`.
- `sign.html` — on-site e-sign page for the single-property, non-exclusive buyer agreement +
  Nevada Duties Owed (form 525). Terms: 3% of purchase price, 30 days, live.
- `robots.txt`, `sitemap.xml`, `llms.txt` — SEO and AI-search files.
- `google-apps-script/Code.gs` + `appsscript.json` — pasted by Julián into the Apps Script of his
  "JC Open House Leads" Google Sheet, deployed as the web app the site posts to (`SHEET_URL`).
  Handles: website leads → sheet tabs + email + Follow Up Boss; mortgage rates; buyer agreements
  (PDF to Drive, emails); FUB Inbox auto-import from Google Drive; lead scoring; FUB Scorecard +
  Hot/Warm tagging (`scoreFubDatabase`, nightly via `setupNightlyScoring`); Monday digest; `loadFubTemplates` (Hot/Warm emails).
  Keys live in Script Properties (`FUB_API_KEY`, etc.), never in code.
  After changing it, Julián must paste it in and redeploy (Deploy → Manage → Edit → New version).
- `fub-connector/worker.js` — Follow Up Boss connector for Claude (Cloudflare Worker "fub-claude",
  MCP server). Tools: hot_list, search_people, get_person, update_person, add_note, log_call,
  create_task, list/start action plans, pipelines and deals, email templates, and `fub_api`
  (any FUB endpoint; DELETE needs confirm). Prefer `fub_api` over adding new tools, because a
  running thread doesn't see tools added after it starts. Julián updates the Worker by pasting code
  in Cloudflare → Edit code → Deploy.

## Follow Up Boss facts
- Tags: every scored contact has one of `Tier A` (Hot) / `Tier B` (Warm) / `Tier C` (Nurture or Cold), refreshed
  nightly. `Intake Tier A/B/C` = how strong their sign-up was (an input, never overwritten). `Hot`/`Warm` tags are
  the only automation triggers, and are never added automatically for Farm, Legacy or "Not DNC checked" contacts.
- His API key cannot create pipelines (403). FUB's API can't create action plans, automations,
  smart lists, or send texts/emails, so give Julián click-by-click steps for those.
- Pipelines: Buyers (Nurture, Hot Buyers, Offer, Due Diligence, Loan/Appraisal, Closed) and
  Sellers (9 stages). Plan: rename Buyers stages to the deal flow from his Asana board
  (Consult Set → Buyer Package + Duties Owed → Pre-Approved → Actively Looking → Offer Submitted →
  In Escrow → Closed).
- Messaging style: Chris Voss (labels, no-oriented questions, calibrated how/what questions).
- Texting only to people who gave consent. Sahara Investment Group (SIG) outreach needs
  compliance approval first.

## Open to-dos (as of Oct 2026)
1. Load the 8 Hot/Warm email templates (connector `fub_api` POST `templates`, or Apps Script
   `loadFubTemplates`), then help Julián build the Hot and Warm action plans + automations
   (tag `Hot`/`Warm` → start plan) in FUB.
2. After the automations exist, run `scoreFubDatabase` so the Hot/Warm tags start the plans.
3. Smart list cleanup (7 lists: Call Today, New Not Contacted, Active on Site, Warm Check-in,
   Likely Sellers, Under Contract, Past Clients & Sphere) and pause unused starter plans.
4. Apps Script: make sure the latest Code.gs + manifest are pasted and redeployed; run
   setupFubInbox, setupWeeklyDigest, previewAgreement.
5. Later: agent network list, SIG/Apollo (after compliance), social posting, Meta audiences.
