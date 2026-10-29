# Follow Up Boss connector for Claude

Lets Claude (claude.ai, the Claude app, and Claude Code) work in your Follow Up Boss account directly:
rank Hot/Warm leads, look people up, change stages and tags, add notes, log calls, create tasks and deals,
and start action plans you already built.

It's one small file (`worker.js`) that runs free on Cloudflare Workers. Your Follow Up Boss API key is stored
as a Cloudflare secret, never in this repo or in chat.

## Set up (about 10 minutes, one time)

1. **Make a free Cloudflare account** at https://dash.cloudflare.com/sign-up.
2. **Create the Worker:** in the left menu click **Workers & Pages** (not anything under AI) → **Create** → **Start with Hello World** → name it `fub-claude` → **Deploy**.
3. **Paste the code:** click **Edit code**, delete everything, paste all of `worker.js`, then click **Deploy**.
4. **Add two secrets:** go to the Worker's **Settings → Variables and Secrets → Add**, type **Secret**:
   - `FUB_API_KEY`: your Follow Up Boss API key (the same one in your Google Apps Script).
   - `CONNECTOR_SECRET`: a long random password, letters and numbers only, 30 or more characters.
     (A password manager's generator works well.)
5. **Your connector URL** is:
   `https://fub-claude.<your-subdomain>.workers.dev/<CONNECTOR_SECRET>/mcp`
   (the first part is shown at the top of the Worker page).
6. **Add it to Claude:** go to https://claude.ai/customize/connectors → **Add custom connector** →
   name it `Follow Up Boss` → paste the URL → **Add**.
7. **Start a new conversation** (or a new Claude Code session) so the tools load. Try:
   *"Show me my Hot leads in Follow Up Boss."*

**Keep the URL private.** Anyone who has it can read and change your contacts. If it ever leaks, change
`CONNECTOR_SECRET` in Cloudflare and update the URL in Claude.

## What Claude can do with it

| Tool | What it does |
|---|---|
| `hot_list` | Ranks everyone Hot / Warm / Nurture / Cold, with the reason (same rules as the FUB Scorecard sheet) |
| `search_people` | Find people by name, email, phone, tag or stage |
| `get_person` | Full contact plus their recent notes |
| `update_person` | Change stage, add or remove tags |
| `add_note` | Add a note to a contact |
| `log_call` | Log a call with its outcome |
| `create_task` | Create a follow-up task with a due date |
| `list_action_plans` / `start_action_plan` | Start an existing action plan for one or more people |
| `list_pipelines` / `create_pipeline` | See your deal flows, or build a new one with its stages |
| `list_email_templates` / `create_email_template` / `update_email_template` | Load ready-to-use email templates (with merge fields) for action plans |
| `create_deal` / `update_deal` / `list_deals` | Create deals, move them through stages, update price and close date |

## What it can't do

Follow Up Boss doesn't allow these through its API, so they stay manual in FUB:
creating or editing **action plans**, **automations**, **smart lists**, and **text templates**, and **sending texts or emails**.
Claude can still start plans you've already built, which then send the messages.

## Updating the connector

When `worker.js` changes: open the Worker in Cloudflare → **Edit code** → select all, paste the new file → **Deploy**.
Your secrets and link stay the same. Start a new Claude conversation to see new tools.
