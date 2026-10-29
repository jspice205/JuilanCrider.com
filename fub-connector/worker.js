// Follow Up Boss connector for Claude (remote MCP server on Cloudflare Workers).
// Paste this file into a Cloudflare Worker and add two secrets (Settings > Variables and Secrets):
//   FUB_API_KEY       your Follow Up Boss API key
//   CONNECTOR_SECRET  a long random string; it becomes part of the private connector URL
// Connector URL to add in Claude: https://<your-worker>.workers.dev/<CONNECTOR_SECRET>/mcp
// Treat that URL like a password: anyone who has it can read and edit your Follow Up Boss contacts.

const FUB = 'https://api.followupboss.com/v1/';
const PROTOCOL = '2025-06-18';

async function fub(env, path, {method = 'GET', body} = {}) {
  const res = await fetch(FUB + path, {
    method,
    headers: {Authorization: 'Basic ' + btoa(env.FUB_API_KEY + ':'), 'Content-Type': 'application/json', 'X-System': 'juliancrider-claude'},
    body: body ? JSON.stringify(body) : undefined
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Follow Up Boss ${res.status}: ${text.slice(0, 300)}`);
  return text ? JSON.parse(text) : {};
}

async function allPeople(env, max = 2000) {
  const out = [];
  for (let offset = 0; offset < max; offset += 100) {
    const page = (await fub(env, `people?limit=100&offset=${offset}&sort=created&fields=allFields`)).people || [];
    out.push(...page);
    if (page.length < 100) break;
  }
  return out;
}

const brief = p => ({id: p.id, name: p.name || [p.firstName, p.lastName].filter(Boolean).join(' '), stage: p.stage, source: p.source,
  phone: ((p.phones || [])[0] || {}).value, email: ((p.emails || [])[0] || {}).value, tags: p.tags || [],
  lastActivity: p.lastActivity, created: p.created, price: p.price, link: `https://app.followupboss.com/2/people/view/${p.id}`});

// Same rules as the Google Sheet "FUB Scorecard" (google-apps-script/Code.gs scorePerson).
function score(p) {
  const tags = (p.tags || []).map(String), has = t => tags.some(x => x.toLowerCase() === t.toLowerCase());
  const stage = String(p.stage || ''), days = iso => iso ? Math.floor((Date.now() - new Date(iso)) / 864e5) : null;
  const active = days(p.lastActivity), created = days(p.created), contacted = Number(p.contacted) > 0 || p.contacted === true;
  if (/trash|unresponsive|do not/i.test(stage) || has('Has Agent')) return {tier: 'Do not contact', score: 0, why: stage || 'has another agent'};
  if (/past client|sphere/i.test(stage) || has('Past Client')) return {tier: 'Past client / sphere', score: 0, why: stage || 'past client'};
  if (has('Partner')) return {tier: 'Partner', score: 0, why: 'agent / vendor'};
  if (/active client|pending|under contract/i.test(stage)) return {tier: 'Active client', score: 100, why: stage};
  let s = 0; const why = [];
  const add = (cond, pts, label) => { if (cond) { s += pts; why.push(label); } };
  add(/hot/i.test(stage), 40, 'stage ' + stage); add(has('Showing requested'), 40, 'requested a showing'); add(has('Agreement signed'), 30, 'signed buyer agreement');
  add(has('Tier A'), 30, 'Tier A'); add(!has('Tier A') && has('Tier B'), 15, 'Tier B'); add(has('Legacy – High'), 20, 'high-value legacy seed');
  add(has('Farm Rank A'), 10, 'farm rank A'); add(has('Likely Seller'), 10, 'likely seller');
  add(active !== null && active <= 7, 30, `active ${active}d ago`); add(active !== null && active > 7 && active <= 30, 15, `active ${active}d ago`);
  add(created !== null && created <= 14, 15, `new lead (${created}d)`); add(Number(p.price) >= 1e6, 5, 'luxury price range');
  if (!contacted && created !== null && created <= 30) why.push('NEVER CONTACTED');
  return {tier: s >= 50 ? 'Hot' : s >= 25 ? 'Warm' : s >= 10 ? 'Nurture' : 'Cold', score: s, why: why.join('; ') || 'no recent signals'};
}

const TOOLS = [
  {name: 'hot_list', description: 'Rank everyone in Follow Up Boss (Hot, Warm, Nurture, Cold, active clients, past clients, partners) and return the top people to call, with reasons.',
    inputSchema: {type: 'object', properties: {tier: {type: 'string', description: 'Optional: only this tier, e.g. Hot or Warm'}, limit: {type: 'number', description: 'How many to return (default 25)'}}},
    run: async (env, a) => {
      const ranked = (await allPeople(env)).map(p => ({...brief(p), ...score(p)}))
        .filter(x => !a.tier || x.tier.toLowerCase() === String(a.tier).toLowerCase())
        .sort((x, y) => y.score - x.score);
      const counts = {}; ranked.forEach(x => counts[x.tier] = (counts[x.tier] || 0) + 1);
      return {counts, people: ranked.slice(0, a.limit || 25)};
    }},
  {name: 'search_people', description: 'Find people in Follow Up Boss by name, email, phone, tag, stage or source (partial matches).',
    inputSchema: {type: 'object', properties: {query: {type: 'string'}, tag: {type: 'string'}, stage: {type: 'string'}, source: {type: 'string'}, limit: {type: 'number'}}},
    run: async (env, a) => {
      const q = String(a.query || '').toLowerCase(), digits = q.replace(/\D/g, '');
      return (await allPeople(env)).filter(p => {
        const b = brief(p);
        if (q && !([b.name, b.email].join(' ').toLowerCase().includes(q) || (digits.length >= 4 && String(b.phone || '').replace(/\D/g, '').includes(digits)))) return false;
        if (a.tag && !b.tags.some(t => t.toLowerCase().includes(String(a.tag).toLowerCase()))) return false;
        if (a.stage && !String(b.stage || '').toLowerCase().includes(String(a.stage).toLowerCase())) return false;
        if (a.source && !String(b.source || '').toLowerCase().includes(String(a.source).toLowerCase())) return false;
        return true;
      }).slice(0, a.limit || 50).map(brief);
    }},
  {name: 'get_person', description: 'Full details for one person, including their 10 most recent notes.',
    inputSchema: {type: 'object', properties: {person_id: {type: 'number'}}, required: ['person_id']},
    run: async (env, a) => {
      const p = await fub(env, `people/${a.person_id}?fields=allFields`);
      const notes = (await fub(env, `notes?personId=${a.person_id}&limit=10&sort=-created`)).notes || [];
      return {...brief(p), score: score(p), addresses: p.addresses, notes: notes.map(n => ({created: n.created, subject: n.subject, body: n.body}))};
    }},
  {name: 'update_person', description: 'Change a person\'s stage, price range, or tags (add and/or remove). Existing tags are kept unless removed.',
    inputSchema: {type: 'object', properties: {person_id: {type: 'number'}, stage: {type: 'string'}, price: {type: 'number'},
      add_tags: {type: 'array', items: {type: 'string'}}, remove_tags: {type: 'array', items: {type: 'string'}}}, required: ['person_id']},
    run: async (env, a) => {
      const body = {};
      if (a.stage) body.stage = a.stage;
      if (a.price) body.price = a.price;
      if ((a.add_tags || []).length || (a.remove_tags || []).length) {
        const cur = (await fub(env, `people/${a.person_id}`)).tags || [];
        const rm = (a.remove_tags || []).map(t => t.toLowerCase());
        body.tags = [...new Set(cur.filter(t => !rm.includes(t.toLowerCase())).concat(a.add_tags || []))];
      }
      return brief(await fub(env, `people/${a.person_id}`, {method: 'PUT', body}));
    }},
  {name: 'add_note', description: 'Add a note to a person.',
    inputSchema: {type: 'object', properties: {person_id: {type: 'number'}, body: {type: 'string'}, subject: {type: 'string'}}, required: ['person_id', 'body']},
    run: async (env, a) => fub(env, 'notes', {method: 'POST', body: {personId: a.person_id, subject: a.subject || 'Note', body: a.body}})},
  {name: 'log_call', description: 'Log a phone call with a person (outcome e.g. "Talked", "Left message", "No answer").',
    inputSchema: {type: 'object', properties: {person_id: {type: 'number'}, outcome: {type: 'string'}, note: {type: 'string'}, phone: {type: 'string'}}, required: ['person_id', 'outcome']},
    run: async (env, a) => fub(env, 'calls', {method: 'POST', body: {personId: a.person_id, outcome: a.outcome, note: a.note || '', phone: a.phone || '', isIncoming: false}})},
  {name: 'create_task', description: 'Create a follow-up task (Call, Text, Email, Follow Up, Showing, Appointment, Other) with a due date (YYYY-MM-DD).',
    inputSchema: {type: 'object', properties: {person_id: {type: 'number'}, name: {type: 'string'}, type: {type: 'string'}, due_date: {type: 'string'}}, required: ['person_id', 'name', 'due_date']},
    run: async (env, a) => fub(env, 'tasks', {method: 'POST', body: {personId: a.person_id, name: a.name, type: a.type || 'Follow Up', dueDate: a.due_date}})},
  {name: 'list_action_plans', description: 'List your Follow Up Boss action plans (id and name).',
    inputSchema: {type: 'object', properties: {}},
    run: async env => ((await fub(env, 'actionPlans?limit=100')).actionPlans || []).map(x => ({id: x.id, name: x.name, status: x.status}))},
  {name: 'start_action_plan', description: 'Put one or more people on an existing action plan.',
    inputSchema: {type: 'object', properties: {action_plan_id: {type: 'number'}, person_ids: {type: 'array', items: {type: 'number'}}}, required: ['action_plan_id', 'person_ids']},
    run: async (env, a) => {
      const done = [];
      for (const id of a.person_ids) done.push(await fub(env, 'actionPlansPeople', {method: 'POST', body: {personId: id, actionPlanId: a.action_plan_id}}));
      return {started: done.length};
    }},
  {name: 'list_email_templates', description: 'List your Follow Up Boss email templates (id, name, subject).',
    inputSchema: {type: 'object', properties: {}},
    run: async env => ((await fub(env, 'templates?limit=100')).templates || []).map(t => ({id: t.id, name: t.name, subject: t.subject}))},
  {name: 'create_email_template', description: 'Create an email template in Follow Up Boss, ready to use in action plans. ' +
      'Body is HTML; merge fields like %contact_first_name% work. Write short, plain emails and sign as Julián.',
    inputSchema: {type: 'object', properties: {name: {type: 'string'}, subject: {type: 'string'}, body: {type: 'string', description: 'HTML, e.g. <p>Hi %contact_first_name%,</p>'}},
      required: ['name', 'subject', 'body']},
    run: async (env, a) => fub(env, 'templates', {method: 'POST', body: {name: a.name, subject: a.subject, body: a.body}})},
  {name: 'update_email_template', description: 'Change the name, subject or body of an existing email template.',
    inputSchema: {type: 'object', properties: {template_id: {type: 'number'}, name: {type: 'string'}, subject: {type: 'string'}, body: {type: 'string'}}, required: ['template_id']},
    run: async (env, a) => {
      const body = {name: a.name, subject: a.subject, body: a.body};
      Object.keys(body).forEach(k => body[k] === undefined && delete body[k]);
      return fub(env, 'templates/' + a.template_id, {method: 'PUT', body});
    }},
  {name: 'list_pipelines', description: 'List deal pipelines and their stages (ids needed for create_deal).',
    inputSchema: {type: 'object', properties: {}},
    run: async env => ((await fub(env, 'pipelines')).pipelines || []).map(p => ({id: p.id, name: p.name, stages: (p.stages || []).map(s => ({id: s.id, name: s.name}))}))},
  {name: 'create_deal', description: 'Create a deal in a pipeline stage, linked to people.',
    inputSchema: {type: 'object', properties: {name: {type: 'string'}, stage_id: {type: 'number'}, person_ids: {type: 'array', items: {type: 'number'}},
      price: {type: 'number'}, projected_close_date: {type: 'string', description: 'YYYY-MM-DD'}}, required: ['name', 'stage_id']},
    run: async (env, a) => fub(env, 'deals', {method: 'POST', body: {name: a.name, stageId: a.stage_id, peopleIds: a.person_ids || [],
      price: a.price, projectedCloseDate: a.projected_close_date}})},
  {name: 'update_deal', description: 'Move a deal to another stage, or change its name, price, close date or people.',
    inputSchema: {type: 'object', properties: {deal_id: {type: 'number'}, stage_id: {type: 'number'}, name: {type: 'string'}, price: {type: 'number'},
      projected_close_date: {type: 'string', description: 'YYYY-MM-DD'}, person_ids: {type: 'array', items: {type: 'number'}}}, required: ['deal_id']},
    run: async (env, a) => {
      const body = {stageId: a.stage_id, name: a.name, price: a.price, projectedCloseDate: a.projected_close_date, peopleIds: a.person_ids};
      Object.keys(body).forEach(k => body[k] === undefined && delete body[k]);
      return fub(env, 'deals/' + a.deal_id, {method: 'PUT', body});
    }},
  {name: 'create_pipeline', description: 'Create a new deal pipeline (deal flow) with its stages in order, e.g. Buyers: Showing, Offer, Under Contract, Closed.',
    inputSchema: {type: 'object', properties: {name: {type: 'string'}, stages: {type: 'array', items: {type: 'string'}, description: 'Stage names, first to last'}},
      required: ['name', 'stages']},
    run: async (env, a) => fub(env, 'pipelines', {method: 'POST', body: {name: a.name, stages: a.stages.map(name => ({name}))}})},
  {name: 'list_deals', description: 'List open deals (optionally for one pipeline).',
    inputSchema: {type: 'object', properties: {pipeline_id: {type: 'number'}}},
    run: async (env, a) => ((await fub(env, `deals?limit=100${a.pipeline_id ? '&pipelineId=' + a.pipeline_id : ''}`)).deals || [])
      .map(d => ({id: d.id, name: d.name, stage: d.stageName || d.stageId, price: d.price, projectedCloseDate: d.projectedCloseDate}))}
];

function rpc(id, result) { return {jsonrpc: '2.0', id, result}; }
function rpcError(id, code, message) { return {jsonrpc: '2.0', id, error: {code, message}}; }

async function handle(msg, env) {
  const {id, method, params = {}} = msg;
  if (method === 'initialize') return rpc(id, {protocolVersion: params.protocolVersion || PROTOCOL, capabilities: {tools: {}},
    serverInfo: {name: 'follow-up-boss', version: '1.0.0'},
    instructions: 'Tools for Julián Crider\'s Follow Up Boss: rank leads, search people, update stages and tags, add notes, log calls, create tasks and deals, create email templates, and start existing action plans. Automations and action plans themselves cannot be created through the API.'});
  if (method === 'ping') return rpc(id, {});
  if (method === 'tools/list') return rpc(id, {tools: TOOLS.map(({name, description, inputSchema}) => ({name, description, inputSchema}))});
  if (method === 'tools/call') {
    const tool = TOOLS.find(t => t.name === params.name);
    if (!tool) return rpcError(id, -32602, 'Unknown tool: ' + params.name);
    try {
      const out = await tool.run(env, params.arguments || {});
      return rpc(id, {content: [{type: 'text', text: JSON.stringify(out, null, 1)}]});
    } catch (err) {
      return rpc(id, {content: [{type: 'text', text: String(err.message || err)}], isError: true});
    }
  }
  if (id === undefined) return null;  // notifications (e.g. notifications/initialized) need no reply
  return rpcError(id, -32601, 'Method not found: ' + method);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!env.CONNECTOR_SECRET || !env.FUB_API_KEY) return new Response('Connector not configured: add FUB_API_KEY and CONNECTOR_SECRET.', {status: 500});
    if (url.pathname !== `/${env.CONNECTOR_SECRET}/mcp`) return new Response('Not found', {status: 404});
    if (request.method === 'GET') return new Response('Method not allowed', {status: 405, headers: {Allow: 'POST'}});
    if (request.method !== 'POST') return new Response(null, {status: 405});
    let body;
    try { body = await request.json(); } catch { return Response.json(rpcError(null, -32700, 'Parse error'), {status: 400}); }
    const replies = (await Promise.all((Array.isArray(body) ? body : [body]).map(m => handle(m, env)))).filter(Boolean);
    if (!replies.length) return new Response(null, {status: 202});
    return Response.json(Array.isArray(body) ? replies : replies[0], {headers: {'MCP-Protocol-Version': PROTOCOL}});
  }
};
