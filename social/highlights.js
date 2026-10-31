// Instagram Highlight covers for @juliancriderestates in the calm black-and-white look (red as a small accent).
// Usage: node social/highlights.js  -> social/highlights/<name>.jpg (1080x1920; Instagram shows the center circle)
const {chromium} = require('playwright');
const path = require('path');
const OUT = path.join(__dirname, 'highlights');
// Simple line icons on a 24x24 grid
const ICONS = {
  'start-here': '<path d="M12 3l2.2 5.6L20 9.3l-4.5 3.8L17 19l-5-3.1L7 19l1.5-5.9L4 9.3l5.8-.7z"/>',
  'tours': '<rect x="3" y="7" width="12" height="10" rx="2"/><path d="M15 11l6-3.5v9L15 13"/>',
  'homes': '<path d="M3 11l9-7 9 7"/><path d="M5 9.5V20h14V9.5"/><path d="M10 20v-6h4v6"/>',
  'moving-to-vegas': '<path d="M12 21s-7-6.2-7-11.5a7 7 0 0114 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>',
  'reviews': '<path d="M4 5h16v11H9l-5 4z"/><path d="M8 9.5h8M8 12.5h5"/>',
  'book': '<rect x="3.5" y="5" width="17" height="15" rx="2"/><path d="M3.5 9.5h17M8 3v4M16 3v4"/><path d="M9 14.5l2 2 4-4"/>',
};
(async () => {
  const b = await chromium.launch(), p = await b.newPage({viewport: {width: 1080, height: 1920}});
  for (const [name, svg] of Object.entries(ICONS)) {
    await p.setContent(`<html><body style="margin:0;width:1080px;height:1920px;background:#0b0b0c;display:flex;align-items:center;justify-content:center">
      <div style="position:relative;width:860px;height:860px;border-radius:50%;border:4px solid #2a2a2c;display:flex;align-items:center;justify-content:center">
        <svg width="520" height="520" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="1" stroke-linecap="round" stroke-linejoin="round">${svg}</svg>
        <div style="position:absolute;bottom:120px;width:80px;height:6px;background:#b3262b"></div>
      </div></body></html>`);
    await p.screenshot({path: path.join(OUT, `${name}.jpg`), type: "jpeg", quality: 85});
  }
  await b.close();
  console.log('done:', Object.keys(ICONS).join(', '));
})();
