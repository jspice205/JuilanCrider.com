// Instagram Story graphics (1080x1920) in the calm black-and-white look: reviews, the Book story and Start Here.
// Usage: node social/stories.js  -> social/stories/<name>.jpg
// Reviews must be real and quoted word for word (first name + last initial, where it came from). Never make one up.
const {chromium} = require('playwright');
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..'), IMG = path.join(ROOT, 'img'), FONTS = path.join(__dirname, 'fonts');
const OUT = path.join(__dirname, 'stories');
const data = (f, t) => `data:${t};base64,` + fs.readFileSync(f).toString('base64');
const font = (n, f, st = 'normal') => `@font-face{font-family:${n};font-style:${st};src:url(${data(path.join(FONTS, f), 'font/ttf')});font-weight:100 900}`;
const CSS = font('Anton', 'Anton-Regular.ttf') + font('Osw', 'Oswald.ttf') + font('Inter', 'Inter.ttf') + font('Corm', 'Cormorant.ttf') +
  font('Corm', 'Cormorant-Italic.ttf', 'italic') + `*{margin:0;padding:0;box-sizing:border-box}
body{width:1080px;height:1920px;background:#0b0b0c;color:#fff;position:relative;overflow:hidden}
.bg{position:absolute;inset:0;background-size:cover;background-position:center}
.dim{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.35),rgba(0,0,0,.55) 45%,rgba(0,0,0,.85))}
.tag{position:absolute;top:250px;left:90px;border:1px solid #fff;font:500 26px Osw;letter-spacing:.32em;padding:12px 22px;text-transform:uppercase}
.rule{width:90px;height:4px;background:#d13438;margin:0 0 40px}
.foot{position:absolute;bottom:250px;left:90px;right:90px;font:500 24px Osw;letter-spacing:.26em;color:#c7c7cc;text-transform:uppercase}
.q{position:absolute;left:90px;right:90px;top:520px}
.q .mark{font:italic 400 260px/0.6 Corm;color:#d13438;height:120px}
.q p{font:italic 400 64px/1.25 Corm;margin-top:30px}
.q .stars{margin-top:60px;font:400 44px Inter;letter-spacing:.2em;color:#fff}
.q .who{margin-top:24px;font:500 30px Osw;letter-spacing:.2em;text-transform:uppercase;color:#c7c7cc}
.big{position:absolute;left:90px;right:90px;top:600px}
.big h1{font:500 128px/.98 Corm}.big h1 i{font-weight:400}
.big p{margin-top:44px;font:400 40px/1.5 Inter;color:#e5e5ea}
.big ul{margin-top:44px;list-style:none;font:500 40px/1.9 Osw;letter-spacing:.06em;text-transform:uppercase}
.big li:before{content:'— ';color:#d13438}
.sticker{display:none}
.ph{position:absolute;top:380px;left:90px;width:420px;height:520px;background-size:cover;background-position:center 15%}
.ph:after{content:'';position:absolute;left:-16px;top:-16px;right:16px;bottom:16px;border:2px solid #d13438}`;
const FOOT = 'Julián Crider · Realtor® · Boeckle &amp; Co. · S.0177760';

// Real reviews from Zillow, as shown on juliancrider.com. Add more from Zillow/Google word for word.
const REVIEWS = [
  ['Julian has been amazing. He helped us sell our home and buy new in Las Vegas. He is super responsive, got us all we asked for, has great taste, and is a true expert in his field.', 'Jessica S.', 'Bought and sold · Zillow'],
  ['He found me the perfect condo even after I set up almost impossible must-haves… and more I didn’t know I wanted.', 'Katie H.', 'Spring Valley · Zillow'],
  ['He didn’t just send us a list of homes and email us “What do you think.” He knew the homes before we saw them.', 'Buyer', 'Summerlin North · Zillow'],
  ['He is professional and personal. He is knowledgeable about the market… He gave great advice when it came to making offers. I highly recommend Julian.', 'Cheryl H.', 'Providence · Zillow'],
  ['Our family is so pleased how he has a wealth of knowledge and true compassion for first-time home buyers.', 'Suzette P.', 'Spring Valley · Zillow'],
];

const pages = {};
REVIEWS.forEach(([q, who, where], i) => pages[`review-${String(i + 1).padStart(2, '0')}`] =
  `<div class="tag">Client Reviews</div><div class="q"><div class="mark">“</div><p>${q}</p><div class="stars">★★★★★</div><div class="who">${who} · ${where}</div></div><div class="foot">${FOOT}</div>`);
pages['book'] = `<div class="bg" style="background-image:url(${data(path.join(IMG, 'golf.jpg'), 'image/jpeg')})"></div><div class="dim"></div>
  <div class="tag">Book a Call</div><div class="big" style="top:520px"><div class="rule"></div><h1>Your Vegas<br><i>game plan.</i></h1>
  <p>Free 30-minute call. No pressure, no obligation.</p><ul><li>Neighborhoods for your budget</li><li>What it really costs</li><li>Your next 3 steps</li></ul></div>
  <div class="sticker">Link sticker here ↓</div><div class="foot">${FOOT}</div>`;
pages['start-01-hi'] = `<div class="ph" style="background-image:url(${data(path.join(IMG, 'julian-headshot.jpg'), 'image/jpeg')})"></div>
  <div class="tag" style="top:180px">Start Here · 1/4</div><div class="big" style="top:1010px"><div class="rule"></div><h1>Hi, I’m <i>Julián.</i></h1>
  <p>Las Vegas &amp; Henderson Realtor®. I help people buy, sell and move here, without the stress.</p></div><div class="foot">${FOOT}</div>`;
pages['start-02-where'] = `<div class="bg" style="background-image:url(${data(path.join(IMG, 'northwest-summerlin.jpg'), 'image/jpeg')})"></div><div class="dim"></div>
  <div class="tag">Start Here · 2/4</div><div class="big"><div class="rule"></div><h1>Where I <i>help.</i></h1>
  <ul><li>Summerlin</li><li>Henderson</li><li>Southwest &amp; Spring Valley</li><li>Skye Canyon &amp; the Northwest</li><li>New construction valley-wide</li></ul></div><div class="foot">${FOOT}</div>`;
pages['start-03-how'] = `<div class="bg" style="background-image:url(${data(path.join(IMG, 'henderson-southeast.jpg'), 'image/jpeg')})"></div><div class="dim"></div>
  <div class="tag">Start Here · 3/4</div><div class="big"><div class="rule"></div><h1>How I <i>help.</i></h1>
  <ul><li>Buying &amp; first homes</li><li>Moving from out of state</li><li>Selling for top dollar</li><li>Golf &amp; pickleball communities</li></ul>
  <p>I preview homes before you see them, so you tour fewer houses and the right ones.</p></div><div class="foot">${FOOT}</div>`;
pages['start-04-next'] = `<div class="bg" style="background-image:url(${data(path.join(IMG, 'explore-high-roller.jpg'), 'image/jpeg')})"></div><div class="dim" style="background:rgba(0,0,0,.62)"></div>
  <div class="tag">Start Here · 4/4</div><div class="big" style="top:520px"><div class="rule"></div><h1>Your next <i>step.</i></h1>
  <ul><li>Book a free 30-min call</li><li>DM me HOMES</li><li>Get my relocation guide</li></ul></div>
  <div class="sticker">Link sticker here ↓</div><div class="foot">${FOOT}</div>`;

(async () => {
  fs.mkdirSync(OUT, {recursive: true});
  const b = await chromium.launch(), p = await b.newPage({viewport: {width: 1080, height: 1920}});
  for (const [name, html] of Object.entries(pages)) {
    await p.setContent(`<html><head><style>${CSS}</style></head><body>${html}</body></html>`);
    await p.waitForTimeout(200);
    await p.screenshot({path: path.join(OUT, `${name}.jpg`), type: 'jpeg', quality: 88});
  }
  await b.close();
  console.log('done:', Object.keys(pages).join(', '));
})();
