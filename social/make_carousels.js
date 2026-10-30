// Instagram carousel maker for @juliancriderestates: 1080x1350 slides in the site's style.
// Add a new set to `sets` (bg = a photo in img/), run it, convert PNGs to JPG, commit, then schedule in Metricool
// using https://juliancrider.com/social/carousels/<key>-N.jpg.
const {chromium}=require('playwright'); const fs=require('fs'), path=require('path');
const IMGDIR=path.join(__dirname,'..','img')+'/'; const IMG=''; const data=f=>'data:image/jpeg;base64,'+fs.readFileSync(IMGDIR+f).toString('base64');
const OUT=path.join(__dirname,'carousels');
// Usage: node social/make_carousels.js [setKey]  -> writes social/carousels/<key>-N.png (convert to .jpg before committing)
const ONLY=process.argv[2];
const SIGN='Julián Crider, Realtor® · Boeckle & Co. · NV S.0177760';
const sets={
 'hoa':{bg:'henderson-southeast.jpg', kicker:'Buying in Las Vegas', title:'5 questions to ask before you buy in an HOA', sub:'Save this for your next showing',
   slides:[['What does the monthly fee actually cover?','Landscaping, gates, pools, water, insurance? Two HOAs with the same fee can cover very different things.'],
           ['How healthy is the reserve fund?','A thin reserve can mean a surprise special assessment later. Ask for the budget and reserve study.'],
           ['Any special assessments coming up?','Ask if anything is planned or being discussed, like roofs, roads or a clubhouse remodel.'],
           ['What are the rules I’d actually notice?','Rentals, parking, RVs, pets, paint colors, solar, backyard changes. Read the CC&Rs before you fall in love.'],
           ['Is there a master HOA and a sub-HOA?','Many Las Vegas communities have two. Make sure you’re budgeting for both fees.']]},
 'summerlin':{bg:'northwest-summerlin.jpg', kicker:'Summerlin, Las Vegas', title:'How to pick the right Summerlin village', sub:'4 things that matter more than the zip code',
   slides:[['Your daily drive','Think about where you go most: work, the Strip, the airport, Downtown Summerlin. Some villages save you 20 minutes a day.'],
           ['Home age and style','Older villages have bigger lots and established trees. Newer ones have modern layouts and newer systems.'],
           ['What the HOA includes','Pools, parks, gates, trails and community centers vary village to village, and so do the fees.'],
           ['Lifestyle fit','Golf, trails, Red Rock Canyon access, quiet streets or walkable shopping. Pick the village that fits how you live.']]},
 'california':{bg:'relocate-park.jpg', kicker:'Moving to Las Vegas', title:'Moving from California? Your checklist', sub:'Save this before you pack',
   slides:[['Nevada has no state income tax','Talk to your CPA about what that means for you and when your residency starts.'],
           ['Update your license and plates','New residents need a Nevada license and registration soon after moving. Check the NV DMV for current deadlines.'],
           ['Budget for the heat','Look closely at the AC, shade, pool and sun direction. Summer power bills are part of the real cost.'],
           ['Don’t pick from a map','Neighborhoods change a lot block to block. Tour a few areas at different times of day before you choose.'],
           ['Line up a local lender early','A Nevada pre-approval makes your offer stronger and shows you what you can comfortably afford.']]},
 'buydowns':{bg:'escrow-strip-night.jpg', kicker:'Buying smarter', title:'Buydowns and seller credits, explained simply', sub:'Lower payment, same house',
   slides:[['What’s a seller credit?','Money the seller agrees to put toward your costs at closing. It’s negotiated as part of your offer.'],
           ['Temporary buydown','The credit pays to lower your rate for the first year or two (like a 2-1 buydown). Your payment steps up later.'],
           ['Permanent buydown','Paying “points” up front lowers your rate for the life of the loan. It works best if you’ll keep the loan for years.'],
           ['Which one is right for you?','It depends on your budget, timeline and lender. Ask your lender to run both side by side before you write the offer.']]}
};
const css=`*{margin:0;padding:0;box-sizing:border-box}body{width:1080px;height:1350px;font-family:-apple-system,"SF Pro Display","Inter","Helvetica Neue",Arial,sans-serif;-webkit-font-smoothing:antialiased}
.s{position:relative;width:1080px;height:1350px;overflow:hidden;background:#0b0b0c;color:#fff}
.bg{position:absolute;inset:0;background-size:cover;background-position:center}
.shade{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.15) 0%,rgba(0,0,0,.35) 40%,rgba(0,0,0,.88) 100%)}
.brand{position:absolute;top:64px;left:72px;font-weight:700;font-size:30px;letter-spacing:.01em}
.brand i{display:inline-block;width:14px;height:14px;border-radius:50%;background:#e0353a;margin-left:8px;vertical-align:middle}
.cover .txt{position:absolute;left:72px;right:72px;bottom:120px}
.kick{font-size:30px;font-weight:600;text-transform:uppercase;letter-spacing:.14em;color:#ff5a5f;margin-bottom:22px}
h1{font-size:92px;line-height:1.02;font-weight:800;letter-spacing:-.025em}
.sub{margin-top:30px;font-size:36px;color:#e6e6e6}
.swipe{position:absolute;right:72px;bottom:56px;font-size:28px;color:#cfcfcf}
.card{background:#f4f4f5;color:#111}
.card .num{position:absolute;top:150px;left:72px;font-size:200px;font-weight:800;color:#e0353a;line-height:1;letter-spacing:-.04em}
.card h2{position:absolute;top:420px;left:72px;right:72px;font-size:74px;line-height:1.06;font-weight:800;letter-spacing:-.02em}
.card p{position:absolute;top:760px;left:72px;right:90px;font-size:42px;line-height:1.4;color:#3a3a3c}
.card .brand{color:#111}
.foot{position:absolute;left:72px;right:72px;bottom:56px;font-size:24px;color:#8e8e93;display:flex;justify-content:space-between}
.cta{background:#0b0b0c}
.cta .ph{position:absolute;top:150px;left:50%;transform:translateX(-50%);width:330px;height:330px;border-radius:50%;background-size:cover;background-position:center 20%;border:6px solid #fff}
.cta h2{position:absolute;top:540px;left:72px;right:72px;text-align:center;font-size:80px;line-height:1.05;font-weight:800;letter-spacing:-.02em}
.cta .btn{position:absolute;top:800px;left:50%;transform:translateX(-50%);background:#fff;color:#000;font-size:40px;font-weight:700;padding:28px 52px;border-radius:999px;white-space:nowrap}
.cta .alt{position:absolute;top:960px;left:72px;right:72px;text-align:center;font-size:36px;color:#d1d1d6}
.cta .sign{position:absolute;bottom:70px;left:72px;right:72px;text-align:center;font-size:26px;color:#8e8e93}`;
const brand=`<div class="brand">Julián J. Crider<i></i></div>`;
(async()=>{
  const b=await chromium.launch(); const p=await b.newPage({viewport:{width:1080,height:1350}});
  const made=[];
  for (const [key,c] of Object.entries(sets)) {
    if (ONLY && key!==ONLY) continue;
    const total=c.slides.length+2; const html=[];
    html.push(`<div class="s cover"><div class="bg" style="background-image:url('${data(c.bg)}')"></div><div class="shade"></div>${brand}<div class="txt"><div class="kick">${c.kicker}</div><h1>${c.title}</h1><div class="sub">${c.sub}</div></div><div class="swipe">Swipe →</div></div>`);
    c.slides.forEach((s,i)=>html.push(`<div class="s card">${brand}<div class="num">${String(i+1).padStart(2,'0')}</div><h2>${s[0]}</h2><p>${s[1]}</p><div class="foot"><span>${c.kicker}</span><span>${i+2} / ${total}</span></div></div>`));
    html.push(`<div class="s cta"><div class="ph" style="background-image:url('${data('julian-headshot.jpg')}')"></div><h2>Want help with your move?</h2><div class="btn">Book a 30-min call · link in bio</div><div class="alt">or DM me <b>HOMES</b></div><div class="sign">${SIGN}</div></div>`);
    for (let i=0;i<html.length;i++){
      await p.setContent(`<html><head><style>${css}</style></head><body>${html[i]}</body></html>`,{waitUntil:'load'});
      await p.waitForTimeout(150);
      const f=path.join(OUT,`${key}-${i+1}.png`); await p.screenshot({path:f}); made.push(f);
    }
  }
  console.log(made.length+' slides'); await b.close();
})();
