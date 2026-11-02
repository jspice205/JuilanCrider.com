const {chromium}=require('playwright');const fs=require('fs');const path=require('path');
(async()=>{const [,,jsonPath,out]=process.argv;const repo=process.cwd();
let data=fs.readFileSync(jsonPath,'utf8').replace(/IMG\//g,'file://'+repo+'/img/').replace(/SCR\//g,'file://'+path.resolve(path.dirname(jsonPath),'..')+'/');
const b=await chromium.launch();const p=await b.newPage();
await p.addInitScript('window.CMA='+data);
await p.goto('file://'+repo+'/reports/cma-template.html');await p.waitForTimeout(1500);
await p.pdf({path:out,format:'Letter',printBackground:true});
await p.setViewportSize({width:816,height:1056});
const secs=await p.$$('section.pg');for(let i=0;i<secs.length;i++){const el=secs[i];await el.screenshot({path:out.replace('.pdf','-'+(i+1)+'.png')});}
await b.close();})();
