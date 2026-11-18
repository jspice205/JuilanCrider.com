// Static article pages for search engines and AI search (Nov 18, 2026).
// The site is one page that draws articles with JavaScript, which most crawlers and AI tools don't run.
// This writes a plain HTML copy of every article in BLOGS (index.html) to blog/<slug>/index.html, with its own
// title, description, canonical link and Article schema, then refreshes sitemap.xml and the article list in llms.txt.
// Run after adding or editing an article:  node blog/build.js
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..'), SITE = 'https://juliancrider.com';
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const i = html.indexOf('const BLOGS = ['), j = html.indexOf('\n];', i);
const BLOGS = eval(html.slice(i + 'const BLOGS = '.length, j + 3));
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const text = s => String(s).replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
const today = new Date().toISOString().slice(0, 10);
const CTA = {book: ['Book a free 30-minute call', '/?book'], value: ['Get a free home value report', '/?value'], net: ['Estimate my net proceeds', '/?value'],
  search: ['Search homes', '/?start'], register: ['Register my builder visit', '/?book'], relocate: ['Open the relocation guide', '/?relocate']};

const page = b => {
  const url = `${SITE}/blog/${b.slug}/`, img = /^https?:/.test(b.img) ? b.img : `${SITE}/${b.img}`;
  const body = b.body.filter(p => p !== '[[ig]]').map(p => p.startsWith('## ') ? `<h2>${p.slice(3)}</h2>` : `<p>${p}</p>`).join('\n');
  const cta = CTA[b.cta] || CTA.book;
  const ld = {'@context': 'https://schema.org', '@type': 'Article', headline: text(b.title), description: text(b.ex), image: img,
    mainEntityOfPage: url, inLanguage: 'en-US', keywords: b.kw || '',
    ...(b.date ? {datePublished: b.date, dateModified: b.date} : {}),
    author: {'@type': 'Person', name: 'Julián J. Crider', jobTitle: 'Realtor', url: SITE + '/',
      worksFor: {'@type': 'RealEstateAgent', name: 'Boeckle & Co.', address: {'@type': 'PostalAddress', streetAddress: '4775 W Teco Ave #220', addressLocality: 'Las Vegas', addressRegion: 'NV', postalCode: '89118'}}},
    publisher: {'@type': 'Organization', name: 'Julián J. Crider, Realtor', url: SITE + '/'},
    ...(b.refs ? {citation: b.refs.map(r => ({'@type': 'CreativeWork', name: text(r[0]), url: r[1]}))} : {})};
  return `<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(text(b.title))} | Julián J. Crider, Las Vegas Realtor</title>
<meta name="description" content="${esc(text(b.ex))}">
<link rel="canonical" href="${url}">
<meta property="og:type" content="article"><meta property="og:title" content="${esc(text(b.title))}">
<meta property="og:description" content="${esc(text(b.ex))}"><meta property="og:url" content="${url}"><meta property="og:image" content="${img}">
<meta name="twitter:card" content="summary_large_image">
<link rel="icon" href="/favicon.png">
<script type="application/ld+json">${JSON.stringify(ld)}</script>
<style>
:root{--bg:#fff;--ink:#111;--muted:#6e6e73;--line:#e5e5ea;--red:#d13438}
@media (prefers-color-scheme:dark){:root{--bg:#000;--ink:#f5f5f7;--muted:#98989d;--line:#2c2c2e}}
body{margin:0;background:var(--bg);color:var(--ink);font:17px/1.6 -apple-system,BlinkMacSystemFont,"Inter","Segoe UI",sans-serif}
main{max-width:680px;margin:0 auto;padding:20px 16px 60px}
header a{color:var(--ink);text-decoration:none;font-weight:700}
img.hero{width:100%;border-radius:16px;margin:16px 0;aspect-ratio:4/3;object-fit:cover}
.kick{color:var(--muted);font-size:14px}
h1{font-size:30px;line-height:1.2;margin:6px 0 14px}h2{font-size:21px;margin:28px 0 6px}
a{color:var(--red)}.cta{display:block;text-align:center;background:var(--ink);color:var(--bg);padding:14px;border-radius:999px;text-decoration:none;font-weight:700;margin:28px 0 10px}
.refs{font-size:15px}.note,.foot{color:var(--muted);font-size:14px}.foot{border-top:1px solid var(--line);margin-top:36px;padding-top:16px}
</style></head>
<body><main>
<header><a href="/">Julián J. Crider · Las Vegas &amp; Henderson Real Estate</a></header>
<article>
<img class="hero" src="${/^https?:/.test(b.img) ? b.img : '/' + b.img}" alt="">
<div class="kick">${esc(b.tag)} · ${b.mins} min read${b.date ? ' · ' + new Date(b.date + 'T12:00:00Z').toLocaleDateString('en-US', {month: 'long', day: 'numeric', year: 'numeric', timeZone: 'UTC'}) : ''} · By Julián J. Crider</div>
<h1>${b.title}</h1>
${body}
${b.refs ? `<h2>Sources</h2><ol class="refs">${b.refs.map(r => `<li><a href="${r[1]}" rel="noopener">${r[0]}</a></li>`).join('')}</ol>${b.refnote ? `<p class="note">${b.refnote}</p>` : ''}` : ''}
<a class="cta" href="${cta[1]}">${cta[0]}</a>
</article>
<p class="foot">Julián J. Crider, Realtor® · The DeBenedetto Team at Boeckle &amp; Co. · NV S.0177760 · 702-498-2571 · <a href="/">juliancrider.com</a></p>
</main></body></html>
`;
};

const dir = path.join(ROOT, 'blog');
BLOGS.forEach(b => {
  fs.mkdirSync(path.join(dir, b.slug), {recursive: true});
  fs.writeFileSync(path.join(dir, b.slug, 'index.html'), page(b));
});

// sitemap.xml: one /blog/<slug>/ line per article (the ?blog= lines go; the static page is the canonical one).
const smPath = path.join(ROOT, 'sitemap.xml');
let sm = fs.readFileSync(smPath, 'utf8').split('\n').filter(l => !/\?blog=|\/blog\//.test(l)).join('\n');
const lines = BLOGS.map(b => `  <url><loc>${SITE}/blog/${b.slug}/</loc><lastmod>${b.date || today}</lastmod><changefreq>monthly</changefreq><priority>0.6</priority></url>`).join('\n');
sm = sm.replace('</urlset>', lines + '\n</urlset>');
fs.writeFileSync(smPath, sm);

// llms.txt: an "Articles" section AI tools can follow.
const llPath = path.join(ROOT, 'llms.txt');
let ll = fs.readFileSync(llPath, 'utf8').replace(/\n## Articles\n[\s\S]*?(?=\n## |$)/, '');
ll = ll.trimEnd() + '\n\n## Articles\n' + BLOGS.map(b => `- [${text(b.title)}](${SITE}/blog/${b.slug}/): ${text(b.ex)}`).join('\n') + '\n';
fs.writeFileSync(llPath, ll);
console.log('Built ' + BLOGS.length + ' article pages');
