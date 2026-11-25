// Wandelt eine HTML-Vorschau (z. B. macOS Quick Look einer .pptx) in ein PDF mit genau einer Folie pro Seite um.
// Die Folien werden im Dokument gesucht (gleich große, querformatige Blöcke untereinander), einzeln
// abfotografiert und dann als Seiten im Folienformat zusammengesetzt.
'use strict';
const { BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');
const os = require('os');

const FIND = `(() => {
  const all = [...document.body.querySelectorAll('*')];
  const sx = window.scrollX, sy = window.scrollY;
  let c = all.map(el => { const r = el.getBoundingClientRect(); return { el, x: r.left + sx, y: r.top + sy, w: r.width, h: r.height }; })
    .filter(o => o.w >= 200 && o.h >= 100 && o.w / o.h > 1.15 && o.w / o.h < 2.4);
  // nur äußerste Kandidaten behalten
  c = c.filter(o => !c.some(p => p !== o && p.el.contains(o.el)));
  // die häufigste Größe ist die Foliengröße
  const key = o => Math.round(o.w / 8) + 'x' + Math.round(o.h / 8);
  const cnt = {}; for (const o of c) cnt[key(o)] = (cnt[key(o)] || 0) + 1;
  let best = null; for (const k in cnt) if (!best || cnt[k] > cnt[best]) best = k;
  const s = c.filter(o => key(o) === best).sort((a, b) => a.y - b.y || a.x - b.x);
  return s.map(o => ({ x: o.x, y: o.y, w: o.w, h: o.h }));
})()`;

async function htmlSlidesToPdf(htmlFile, outF, { target = 1920 } = {}) {
  const w = new BrowserWindow({ show: false, width: 1400, height: 1000, webPreferences: { offscreen: true, sandbox: true, contextIsolation: true, nodeIntegration: false } });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-slides-'));
  try {
    await w.loadFile(htmlFile);
    await new Promise(r => setTimeout(r, 1200));
    let slides = await w.webContents.executeJavaScript(FIND);
    if (!slides.length) return 0;
    const z = Math.max(0.5, Math.min(3, target / slides[0].w));
    w.webContents.setZoomFactor(z);
    const W = Math.ceil(slides[0].w * z) + 4, H = Math.ceil(slides[0].h * z) + 4;
    w.setContentSize(W + 40, H + 40);
    await new Promise(r => setTimeout(r, 600));
    slides = await w.webContents.executeJavaScript(FIND);
    const imgs = [];
    for (let i = 0; i < slides.length; i++) {
      const s = slides[i];
      await w.webContents.executeJavaScript(`window.scrollTo(${Math.max(0, s.x - 2)}, ${Math.max(0, s.y - 2)}); [window.scrollX, window.scrollY]`).then(async ([vx, vy]) => {
        await new Promise(r => setTimeout(r, 250));
        const img = await w.webContents.capturePage({ x: Math.round(s.x - vx), y: Math.round(s.y - vy), width: Math.round(s.w), height: Math.round(s.h) });
        const f = path.join(tmp, 's' + i + '.png'); fs.writeFileSync(f, img.toPNG()); imgs.push(f);
      });
    }
    const ar = slides[0].w / slides[0].h, pw = 13.333, ph = pw / ar;
    const page = `<!doctype html><html><head><style>@page{size:${pw}in ${ph}in;margin:0}html,body{margin:0;padding:0;background:#fff}
      div{width:${pw}in;height:${ph}in;page-break-after:always;break-after:page;overflow:hidden}div:last-child{page-break-after:auto;break-after:auto}
      img{width:100%;height:100%;object-fit:contain;display:block}</style></head><body>${imgs.map(f => `<div><img src="${path.basename(f)}"></div>`).join('')}</body></html>`;
    const pf = path.join(tmp, 'pages.html'); fs.writeFileSync(pf, page);
    await w.loadFile(pf);
    await new Promise(r => setTimeout(r, 500));
    const data = await w.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true, margins: { marginType: 'none' } });
    fs.writeFileSync(outF, data);
    return imgs.length;
  } finally { w.destroy(); fs.rm(tmp, { recursive: true, force: true }, () => {}); }
}
module.exports = { htmlSlidesToPdf };
