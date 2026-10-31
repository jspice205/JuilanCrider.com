// "Desert Sport" carousels + Reel versions for @juliancriderestates (style D, chosen Oct 30 2026).
// Usage: node social/sport.js [setKey]
//   -> social/carousels/<key>-N.jpg (1080x1350 slides) and social/reels/<key>.mp4 (1080x1920 slideshow, no audio;
//      music is added from Instagram's library when scheduling the Reel in Metricool).
// Add new topics to SETS. bg = a photo in img/. Keep copy evergreen: no prices, rates or stats.
const {chromium} = require('playwright');
const fs = require('fs'), path = require('path'), {execFileSync} = require('child_process');
const ROOT = path.join(__dirname, '..'), IMG = path.join(ROOT, 'img'), FONTS = path.join(__dirname, 'fonts');
const OUT = path.join(__dirname, 'carousels'), REELS = path.join(__dirname, 'reels');
const SIGN = 'Julián Crider, Realtor® · Boeckle &amp; Co. · NV S.0177760';

const SETS = {
  hoa: {bg: 'henderson-southeast.jpg', tag: 'Buyer Playbook', cover: ['5 Questions', 'Before the HOA'], sub: 'Las Vegas · Henderson · Save this',
    slides: [['What does the fee actually cover?', 'Landscaping, gates, pools, water, insurance. Two HOAs with the same fee can cover very different things.'],
             ['How healthy is the reserve fund?', 'A thin reserve can mean a surprise special assessment later. Ask for the budget and the reserve study.'],
             ['Any special assessments coming?', 'Ask what’s planned or being discussed: roofs, roads, a clubhouse remodel.'],
             ['Which rules will you actually notice?', 'Rentals, parking, RVs, pets, paint, solar, backyard changes. Read the CC&Rs before you fall in love.'],
             ['Is there a master HOA and a sub-HOA?', 'Many Las Vegas communities have two. Budget for both fees.']]},
  summerlin: {bg: 'northwest-summerlin.jpg', tag: 'Neighborhood Scout', cover: ['Pick Your', 'Summerlin Village'], sub: '4 things that matter more than the zip',
    slides: [['Your daily drive', 'Work, the Strip, the airport, Downtown Summerlin. The right village can save you 20 minutes a day.'],
             ['Home age and style', 'Older villages: bigger lots and mature trees. Newer ones: modern layouts and newer systems.'],
             ['What the HOA includes', 'Pools, parks, gates, trails and community centers vary village to village, and so do the fees.'],
             ['Your lifestyle', 'Golf, trails, Red Rock access, quiet streets or walkable shopping. Pick the village that fits how you live.']]},
  california: {bg: 'relocate-park.jpg', tag: 'Relocation Game Plan', cover: ['Leaving', 'California?'], sub: 'Your Las Vegas move checklist',
    slides: [['No state income tax', 'Nevada doesn’t have one. Ask your CPA what that means for you and when residency starts.'],
             ['License and plates', 'New residents need a Nevada license and registration soon after moving. Check the NV DMV for deadlines.'],
             ['Budget for the heat', 'Check the AC, shade, pool and sun direction. Summer power bills are part of the real cost.'],
             ['Don’t pick from a map', 'Neighborhoods change block to block. Tour a few areas at different times of day.'],
             ['Get a local lender early', 'A Nevada pre-approval makes your offer stronger and shows what you can comfortably afford.']]},
  buydowns: {bg: 'escrow-strip-night.jpg', tag: 'Money Moves', cover: ['Same House,', 'Lower Payment'], sub: 'Buydowns and seller credits, simply',
    slides: [['What’s a seller credit?', 'Money the seller puts toward your costs at closing. It’s negotiated as part of your offer.'],
             ['Temporary buydown', 'The credit lowers your rate for the first year or two (like a 2-1). Your payment steps up later.'],
             ['Permanent buydown', 'Paying “points” up front lowers your rate for the life of the loan. Best if you’ll keep the loan for years.'],
             ['Which one fits you?', 'It depends on your budget and timeline. Have your lender run both side by side before you write the offer.']]},
};

// Color palettes. Pick with PALETTE=<name>. Default "quiet" (black/white, red only on the tag and button; Julián's pick Oct 31).
// Julián doesn't want gold ("everyone uses it"); "site" is the bolder red version, "desert" the original terracotta.
const PALETTES = {
  desert: {paper:'#e9dcc6', ink:'#1e1b18', ink2:'#3b342c', muted:'#7a6b58', accent:'#c2603d', 'on-accent':'#f6efe3', hi:'#e8a27d',
    dark:'#1e1b18', light:'#f6efe3', soft:'#d9cbb4', 'muted-dark':'#8d7f6c', shade:'rgba(30,27,24,.88)', photo:'saturate(.85) contrast(1.05) sepia(.18)'},
  site: {paper:'#f4f4f2', ink:'#111111', ink2:'#3a3a3c', muted:'#8a8a8e', accent:'#d13438', 'on-accent':'#ffffff', hi:'#ef4044',
    dark:'#0b0b0c', light:'#ffffff', soft:'#c7c7cc', 'muted-dark':'#6e6e73', shade:'rgba(0,0,0,.88)', photo:'saturate(.95) contrast(1.08)'},
  // Calmer version of "site": red only on the small tag and button; headlines and numbers stay black/white.
  quiet: {paper:'#f4f4f2', ink:'#111111', ink2:'#3a3a3c', muted:'#8a8a8e', accent:'#b3262b', 'on-accent':'#ffffff', hi:'#ffffff', big:'#111111',
    dark:'#0b0b0c', light:'#ffffff', soft:'#c7c7cc', 'muted-dark':'#6e6e73', shade:'rgba(0,0,0,.88)', photo:'saturate(.95) contrast(1.08)'},
  gold: {paper:'#f3f1ec', ink:'#0e0e0e', ink2:'#3a3835', muted:'#8c857a', accent:'#b8963e', 'on-accent':'#0e0e0e', hi:'#d9b860',
    dark:'#0e0e0e', light:'#f7f5f0', soft:'#cfc8bb', 'muted-dark':'#7d776d', shade:'rgba(10,10,10,.9)', photo:'saturate(.8) contrast(1.1)'},
  navy: {paper:'#eef1f4', ink:'#0f1b2d', ink2:'#33405a', muted:'#7d889b', accent:'#1f6feb', 'on-accent':'#ffffff', hi:'#8fd3ff',
    dark:'#0f1b2d', light:'#ffffff', soft:'#bcc6d6', 'muted-dark':'#6d7a91', shade:'rgba(10,20,36,.9)', photo:'saturate(.9) contrast(1.08)'},
};
const PAL = PALETTES[process.env.PALETTE || 'quiet'];
const VARS = ':root{' + Object.entries(PAL).map(([k, v]) => `--${k}:${v}`).join(';') + '}';
const data = (f, t) => `data:${t};base64,` + fs.readFileSync(f).toString('base64');
const font = (n, f) => `@font-face{font-family:${n};src:url(${data(path.join(FONTS, f), 'font/ttf')});font-weight:100 900}`;
const FONT_CSS = font('Anton', 'Anton-Regular.ttf') + font('Osw', 'Oswald.ttf') + font('Inter', 'Inter.ttf');
const GRAIN = `url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='300' height='300'><filter id='n'><feTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='3' stitchTiles='stitch'/></filter><rect width='100%' height='100%' filter='url(%23n)' opacity='.55'/></svg>")`;
const CSS = FONT_CSS + VARS + `*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1350px}
.s{position:relative;width:1080px;height:1350px;overflow:hidden;background:var(--paper)}
.grain{position:absolute;inset:0;background:${GRAIN};mix-blend-mode:multiply;opacity:.35}
.bg{position:absolute;inset:0;background-size:cover;background-position:center;filter:var(--photo)}
.tint{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.05),rgba(0,0,0,.15) 40%,var(--shade))}
.tag{position:absolute;top:66px;left:70px;background:var(--accent);color:var(--on-accent);font:600 24px Osw;letter-spacing:.18em;padding:10px 20px;text-transform:uppercase}
.no{position:absolute;top:60px;right:70px;font:400 40px Anton;color:var(--light)}
.cover .t{position:absolute;left:70px;right:70px;bottom:120px;color:var(--light)}
.cover h1{font:400 150px/.92 Anton;text-transform:uppercase}.cover h1 span{color:var(--hi)}
.cover .sub{margin-top:26px;font:500 32px Osw;letter-spacing:.14em;text-transform:uppercase;color:var(--soft)}
.yard{position:absolute;left:70px;right:70px;bottom:70px;display:flex;gap:16px;align-items:center;font:500 20px Osw;letter-spacing:.24em;color:var(--soft)}
.yard i{flex:1;height:2px;background:repeating-linear-gradient(90deg,var(--soft) 0 18px,transparent 18px 30px)}
.card .col{position:absolute;top:110px;left:70px;right:90px}
.card .big{font:400 300px/1 Anton;color:var(--big,var(--accent))}
.card h2{margin-top:40px;font:400 92px/1 Anton;text-transform:uppercase;color:var(--ink)}
.card .bar{margin:44px 0 36px;width:120px;height:8px;background:var(--ink)}
.card p{font:400 38px/1.45 Inter;color:var(--ink2)}
.card .ft{position:absolute;bottom:64px;left:70px;right:70px;display:flex;justify-content:space-between;font:500 22px Osw;letter-spacing:.22em;color:var(--muted);text-transform:uppercase}
.cta{background:var(--dark);color:var(--light)}
.cta .ph{position:absolute;top:150px;left:70px;width:400px;height:500px;background-size:cover;background-position:center 15%}
.cta .ph:after{content:'';position:absolute;left:-14px;top:-14px;right:14px;bottom:14px;border:4px solid var(--accent)}
.cta .side{position:absolute;top:180px;left:520px;right:70px;font:500 28px/1.6 Osw;letter-spacing:.14em;text-transform:uppercase;color:var(--soft)}
.cta .side b{display:block;font:400 64px/1 Anton;color:var(--light);margin-bottom:14px}
.cta h2{position:absolute;top:720px;left:70px;right:70px;font:400 120px/.95 Anton;text-transform:uppercase}.cta h2 span{color:var(--hi)}
.cta .btn{position:absolute;top:1010px;left:70px;background:var(--accent);color:var(--on-accent);font:600 34px Osw;letter-spacing:.12em;padding:22px 34px;text-transform:uppercase}
.cta .alt{position:absolute;top:1120px;left:70px;font:500 28px Osw;letter-spacing:.14em;color:var(--soft);text-transform:uppercase}
.cta .lic{position:absolute;bottom:64px;left:70px;font:500 20px Osw;letter-spacing:.22em;color:var(--muted-dark);text-transform:uppercase}
.reel{width:1080px;height:1920px;background:var(--dark);position:relative}
.reel .frame{position:absolute;top:285px;left:0;width:1080px;height:1350px}
.reel .top,.reel .bot{position:absolute;left:0;right:0;text-align:center;font:500 30px Osw;letter-spacing:.3em;color:var(--soft);text-transform:uppercase}
.reel .top{top:150px}.reel .bot{bottom:140px}`;

function slidesFor(c) {
  const n = c.slides.length + 2, pad = i => String(i).padStart(2, '0');
  const out = [`<div class="s cover"><div class="bg" style="background-image:url(${data(path.join(IMG, c.bg), 'image/jpeg')})"></div><div class="tint"></div><div class="grain"></div><div class="tag">${c.tag}</div><div class="no">01/${pad(n)}</div><div class="t"><h1>${c.cover[0]}<br><span>${c.cover[1]}</span></h1><div class="sub">${c.sub}</div></div><div class="yard"><span>TEE</span><i></i><span>CLOSE</span></div></div>`];
  c.slides.forEach((s, i) => out.push(`<div class="s card"><div class="grain"></div><div class="col"><div class="big">${pad(i + 1)}</div><h2>${s[0]}</h2><div class="bar"></div><p>${s[1]}</p></div><div class="ft"><span>Julián J. Crider</span><span>${pad(i + 2)} / ${pad(n)}</span></div></div>`));
  out.push(`<div class="s cta"><div class="ph" style="background-image:url(${data(path.join(IMG, 'julian-headshot.jpg'), 'image/jpeg')})"></div><div class="side"><b>Your Vegas move, coached.</b>Golf · Pickleball<br>Relocation · New builds</div><div class="grain"></div><h2>Book your<br><span>game plan</span></h2><div class="btn">30-min call · link in bio</div><div class="alt">or DM me HOMES</div><div class="lic">${SIGN}</div></div>`);
  return out;
}

(async () => {
  const only = process.argv[2], b = await chromium.launch();
  const p = await b.newPage({viewport: {width: 1080, height: 1350}});
  const r = await b.newPage({viewport: {width: 1080, height: 1920}});
  for (const [key, c] of Object.entries(SETS)) {
    if (only && key !== only) continue;
    const html = slidesFor(c), frames = [];
    for (let i = 0; i < html.length; i++) {
      await p.setContent(`<html><head><style>${CSS}</style></head><body>${html[i]}</body></html>`);
      await p.waitForTimeout(300);
      const png = path.join(OUT, `${key}-${i + 1}.png`); await p.screenshot({path: png});
      execFileSync('python3', ['-c', 'import sys;from PIL import Image;Image.open(sys.argv[1]).convert("RGB").save(sys.argv[2],"JPEG",quality=88)', png, png.replace(/\.png$/, '.jpg')]);
      // Reel frame: the slide centered on a 9:16 canvas
      await r.setContent(`<html><head><style>${CSS}body{width:1080px;height:1920px}</style></head><body><div class="reel"><div class="top">Save this ↓</div><div class="frame">${html[i]}</div><div class="bot">@juliancriderestates</div></div></body></html>`);
      await r.waitForTimeout(300);
      const fr = path.join(OUT, `${key}-reel-${i + 1}.png`); await r.screenshot({path: fr}); frames.push(fr);
      fs.unlinkSync(png);
    }
    // Slideshow: 3.2s per slide (cover 2.6s), 0.4s crossfades, H.264 for Instagram
    const durs = frames.map((_, i) => i === 0 ? 2.6 : 3.2), args = [];
    frames.forEach((f, i) => args.push('-loop', '1', '-t', String(durs[i]), '-i', f));
    let filter = '', last = '[0:v]', t = durs[0];
    for (let i = 1; i < frames.length; i++) { const lbl = `[x${i}]`; filter += `${last}[${i}:v]xfade=transition=fade:duration=0.4:offset=${(t - 0.4).toFixed(2)}${lbl};`; last = lbl; t += durs[i] - 0.4; }
    filter += `${last}format=yuv420p,fps=30[v]`;
    execFileSync('ffmpeg', ['-y', '-loglevel', 'error', ...args, '-filter_complex', filter, '-map', '[v]', '-c:v', 'libx264', '-preset', 'medium', '-crf', '21', '-movflags', '+faststart', path.join(REELS, `${key}.mp4`)]);
    frames.forEach(f => fs.unlinkSync(f));
    console.log(`${key}: ${html.length} slides + reel (${t.toFixed(1)}s)`);
  }
  await b.close();
})();
