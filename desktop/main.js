// CueStack Desktop – Electron-Hülle um web/index.html
// Zusätzlich: NDI-Ausgabe, HTTP-API für Bitfocus Companion, Office→PDF, Bildschirmauswahl, Vollbild-Ausgänge.
'use strict';
const { app, BrowserWindow, ipcMain, screen, desktopCapturer, session, shell } = require('electron');
const path = require('path');
const http = require('http');
const os = require('os');
const fs = require('fs');
const { execFile } = require('child_process');

let grandiose = null, ndiError = null;
try { grandiose = require('@stagetimerio/grandiose'); } catch (e) { ndiError = e.message; }

const PORT = Number(process.env.CUESTACK_PORT) || 9876;
const HOST = process.env.CUESTACK_HOST || '0.0.0.0';
let win = null;
let lastState = {};
let rid = 0;
const pending = new Map();

if (!app.requestSingleInstanceLock()) { app.quit(); }
app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });

function lanIp() {
  for (const list of Object.values(os.networkInterfaces())) {
    for (const a of list || []) if (a.family === 'IPv4' && !a.internal) return a.address;
  }
  return '127.0.0.1';
}

function createWindow() {
  win = new BrowserWindow({
    width: 1500, height: 950, minWidth: 900, minHeight: 600,
    backgroundColor: '#0e1014', title: 'CueStack', autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true, nodeIntegration: false, backgroundThrottling: false,
    },
  });
  win.loadFile(path.join(__dirname, '..', 'web', 'index.html'));

  win.webContents.setWindowOpenHandler(({ url, frameName }) => {
    if (frameName && frameName.startsWith('cuestack_out_')) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          backgroundColor: '#000000', autoHideMenuBar: true, title: 'CueStack Ausgang',
          webPreferences: { backgroundThrottling: false },
        },
      };
    }
    if (/^https?:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  win.webContents.on('did-create-window', (child, { frameName }) => {
    child.setMenuBarVisibility(false);
    const m = /__fs_(-?\d+)/.exec(frameName || '');
    if (m) {
      const d = screen.getAllDisplays().find(x => String(x.id) === m[1]);
      if (d) { child.setBounds(d.bounds); child.once('ready-to-show', () => child.setFullScreen(true)); setTimeout(() => !child.isDestroyed() && child.setFullScreen(true), 400); }
    }
  });

  // Absturzschutz: Renderer neu laden, Zustand kommt aus der lokalen Sicherung
  win.webContents.on('render-process-gone', () => {
    setTimeout(() => { if (win && !win.isDestroyed()) win.reload(); }, 800);
  });
  win.on('closed', () => { win = null; });
}

/* ---------- Bildschirm-Capture: Auswahl im CueStack-Fenster ---------- */
function askRenderer(channel, payload, timeout = 120000) {
  return new Promise((resolve) => {
    if (!win) return resolve(null);
    const id = ++rid;
    const t = setTimeout(() => { pending.delete(id); resolve(null); }, timeout);
    pending.set(id, (v) => { clearTimeout(t); pending.delete(id); resolve(v); });
    win.webContents.send(channel, { rid: id, ...payload });
  });
}
ipcMain.on('reply', (_e, { rid: id, value }) => { const f = pending.get(id); if (f) f(value); });

function setupCapture() {
  session.defaultSession.setDisplayMediaRequestHandler(async (_req, callback) => {
    try {
      const sources = await desktopCapturer.getSources({ types: ['screen', 'window'], thumbnailSize: { width: 320, height: 180 } });
      const list = sources.map(s => ({ id: s.id, name: s.name, thumb: s.thumbnail.toDataURL() }));
      const id = await askRenderer('pick-source', { list });
      const src = sources.find(s => s.id === id);
      if (src) callback({ video: src }); else callback({});
    } catch (e) { callback({}); }
  });
}

/* ---------- NDI ---------- */
const senders = new Map();
async function getSender(id, name) {
  const s = senders.get(id);
  if (s && s.name === name) return s.p;
  if (s) { senders.delete(id); try { await (await s.p).destroy(); } catch (e) {} }
  const p = grandiose.send({ name, clockVideo: false, clockAudio: false });
  senders.set(id, { name, p });
  p.catch(() => senders.delete(id));
  return p;
}
ipcMain.handle('ndi-frame', async (_e, { id, name, w, h, fps, buf }) => {
  if (!grandiose) return { error: ndiError || 'NDI-Bibliothek nicht geladen' };
  try {
    const snd = await getSender(id, name);
    await snd.video({
      xres: w, yres: h, frameRateN: fps * 1000, frameRateD: 1000,
      fourCC: grandiose.FOURCC_RGBA, pictureAspectRatio: w / h,
      frameFormatType: grandiose.FORMAT_TYPE_PROGRESSIVE, lineStrideBytes: w * 4,
      data: Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength),
    });
    return { ok: true };
  } catch (err) { return { error: err.message }; }
});
ipcMain.handle('ndi-stop', async (_e, id) => {
  const s = senders.get(id); senders.delete(id);
  if (s) { try { await (await s.p).destroy(); } catch (e) {} }
  return true;
});

/* ---------- Infos, Bildschirme, Status ---------- */
ipcMain.handle('info', () => ({ ndi: !!grandiose, ndiError, port: PORT, ip: lanIp(), version: app.getVersion(), platform: process.platform }));
ipcMain.handle('displays', () => {
  const prim = screen.getPrimaryDisplay().id;
  return screen.getAllDisplays().map((d, i) => ({ id: d.id, label: d.label || `Bildschirm ${i + 1}`, bounds: d.bounds, primary: d.id === prim }));
});
ipcMain.on('state', (_e, s) => { lastState = s || {}; });

/* ---------- Office → PDF ---------- */
const PS_SCRIPT = `
$ErrorActionPreference = 'Stop'
$in = $args[0]; $out = $args[1]
$ext = [IO.Path]::GetExtension($in).ToLower()
if ($ext -match 'ppt|pps|odp') {
  $app = New-Object -ComObject PowerPoint.Application
  $p = $app.Presentations.Open($in, -1, 0, 0)
  $p.SaveAs($out, 32); $p.Close(); $app.Quit()
} elseif ($ext -match 'doc|odt|rtf') {
  $app = New-Object -ComObject Word.Application; $app.Visible = $false
  $d = $app.Documents.Open($in, $false, $true)
  $d.ExportAsFixedFormat($out, 17); $d.Close($false); $app.Quit()
} elseif ($ext -match 'xls|ods') {
  $app = New-Object -ComObject Excel.Application; $app.Visible = $false; $app.DisplayAlerts = $false
  $w = $app.Workbooks.Open($in, 0, $true)
  $w.ExportAsFixedFormat(0, $out); $w.Close($false); $app.Quit()
} else { throw "Dateityp nicht unterstützt" }
`;
const AS_SCRIPT = `
on run argv
  set inF to POSIX file (item 1 of argv)
  set outF to POSIX file (item 2 of argv)
  set appName to item 3 of argv
  tell application appName
    set d to open inF
    delay 1
    export d to outF as PDF
    close d saving no
  end tell
end run
`;
function run(cmd, args, timeout = 180000) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, { timeout, windowsHide: true }, (err, stdout, stderr) => err ? reject(new Error((stderr || err.message).toString().trim().slice(0, 300))) : resolve(stdout));
  });
}
function findSoffice() {
  const c = process.platform === 'win32'
    ? ['C:\\Program Files\\LibreOffice\\program\\soffice.exe', 'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe']
    : process.platform === 'darwin' ? ['/Applications/LibreOffice.app/Contents/MacOS/soffice'] : ['/usr/bin/soffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice'];
  return c.find(p => fs.existsSync(p)) || null;
}
ipcMain.handle('convert-office', async (_e, { name, buf }) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-'));
  const safe = String(name).replace(/[^\w.\- ]+/g, '_');
  const inF = path.join(dir, safe), outF = path.join(dir, safe.replace(/\.[^.]+$/, '') + '.pdf');
  fs.writeFileSync(inF, Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength));
  const errors = [];
  const tryIt = async (label, fn) => { try { await fn(); if (fs.existsSync(outF)) return label; } catch (e) { errors.push(label + ': ' + e.message); } return null; };
  let via = null;
  const ext = path.extname(safe).toLowerCase();
  if (process.platform === 'win32') {
    via = await tryIt('Microsoft Office', async () => {
      const ps = path.join(dir, 'convert.ps1'); fs.writeFileSync(ps, '\ufeff' + PS_SCRIPT, 'utf8');
      await run('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps, inF, outF]);
    });
  } else if (process.platform === 'darwin') {
    const appName = /\.(pptx?|ppsx?|key|odp)$/.test(ext) ? 'Keynote' : /\.(xlsx?|ods)$/.test(ext) ? 'Numbers' : 'Pages';
    if (fs.existsSync(`/Applications/${appName}.app`)) {
      via = await tryIt(appName, async () => {
        const as = path.join(dir, 'convert.applescript'); fs.writeFileSync(as, AS_SCRIPT);
        await run('osascript', [as, inF, outF, appName]);
      });
    }
  }
  if (!via) {
    const so = findSoffice();
    if (so) via = await tryIt('LibreOffice', () => run(so, ['--headless', '--convert-to', 'pdf', '--outdir', dir, inF]));
  }
  try {
    if (!via) return { error: errors.length ? errors.join(' | ') : 'Kein Programm zum Umwandeln gefunden. Bitte Microsoft PowerPoint, Keynote oder LibreOffice installieren (LibreOffice ist kostenlos).' };
    return { pdf: fs.readFileSync(outF), via };
  } finally { fs.rm(dir, { recursive: true, force: true }, () => {}); }
});

/* ---------- HTTP-API für Companion & Handy-Fernbedienung ---------- */
const REMOTE_HTML = `<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CueStack Fernbedienung</title><style>
body{margin:0;background:#0e1014;color:#e9ebf1;font:16px system-ui,sans-serif;padding:14px}
h1{font-size:18px;margin:0 0 10px}#st{color:#8d95a8;margin-bottom:12px;min-height:22px}#t{font:700 48px Arial;text-align:center;margin:8px 0}
.g{display:grid;grid-template-columns:1fr 1fr;gap:10px}button{min-height:64px;font-size:18px;font-weight:700;border-radius:12px;border:1px solid #2a2f3a;background:#1d212a;color:#e9ebf1}
.p{background:#3ddc97;color:#06130d;border-color:#3ddc97}.r{background:#ff3b47;color:#fff;border-color:#ff3b47}.w{grid-column:1/-1}
input{width:100%;min-height:48px;font-size:16px;border-radius:10px;border:1px solid #2a2f3a;background:#0b0d11;color:#e9ebf1;padding:0 10px;box-sizing:border-box}
</style></head><body><h1>CueStack</h1><div id="t">–</div><div id="st"></div>
<div class="g"><button data-c="prev">◀ Zurück</button><button class="p" data-c="next">Weiter ▶</button>
<button data-c="timer/toggle">Timer Start/Pause</button><button data-c="timer/reset">Timer Reset</button>
<button data-c="timer/add?sec=60">+1 min</button><button data-c="timer/add?sec=-60">−1 min</button>
<button class="r w" data-c="blackout/toggle">BLACKOUT</button>
<input id="m" placeholder="Nachricht an die Bühne"><button data-c="stage/clear">Nachricht löschen</button>
<button class="w" id="send">Nachricht senden</button></div>
<script>
const st=document.getElementById('st'),t=document.getElementById('t');
async function call(c){try{const r=await fetch('/api/'+c,{method:'POST'});show(await r.json());}catch(e){st.textContent='Keine Verbindung';}}
function show(j){const s=j.state||j;if(!s.timer)return;t.textContent=s.timer.text;t.style.color=s.blackout?'#ff3b47':'';
 st.textContent=(s.blackout?'BLACKOUT · ':'')+'Folie '+s.slides.n+'/'+s.slides.total+(s.stage.message?' · Bühne: '+s.stage.message:'');}
document.querySelectorAll('[data-c]').forEach(b=>b.onclick=()=>call(b.dataset.c));
document.getElementById('send').onclick=()=>call('stage/message?text='+encodeURIComponent(document.getElementById('m').value));
setInterval(()=>fetch('/api/state').then(r=>r.json()).then(show).catch(()=>{st.textContent='Keine Verbindung';}),1000);
</script></body></html>`;

function startHttp() {
  const srv = http.createServer(async (req, res) => {
    const u = new URL(req.url, 'http://localhost');
    const jh = { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'no-store' };
    if (req.method === 'OPTIONS') { res.writeHead(204, { ...jh, 'Access-Control-Allow-Methods': 'GET,POST', 'Access-Control-Allow-Headers': 'Content-Type' }); return res.end(); }
    if (!u.pathname.startsWith('/api')) { res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); return res.end(REMOTE_HTML); }
    const cmd = decodeURIComponent(u.pathname.replace(/^\/api\/?/, ''));
    const p = Object.fromEntries(u.searchParams);
    if (req.method === 'POST') {
      const body = await new Promise(r => { let b = ''; req.on('data', d => { b += d; if (b.length > 1e5) req.destroy(); }); req.on('end', () => r(b)); });
      if (body) { try { Object.assign(p, JSON.parse(body)); } catch (e) { Object.assign(p, Object.fromEntries(new URLSearchParams(body))); } }
    }
    if (!cmd || cmd === 'state') { res.writeHead(200, jh); return res.end(JSON.stringify(lastState)); }
    const r = await askRenderer('remote', { cmd, p }, 5000);
    if (!r) { res.writeHead(503, jh); return res.end(JSON.stringify({ ok: false, error: 'CueStack antwortet nicht' })); }
    res.writeHead(r.ok ? 200 : 400, jh); res.end(JSON.stringify(r));
  });
  srv.on('error', (e) => console.error('HTTP-API:', e.message));
  srv.listen(PORT, HOST);
}

app.whenReady().then(() => {
  setupCapture();
  createWindow();
  startHttp();
  const notify = () => win && win.webContents.send('displays-changed');
  screen.on('display-added', notify); screen.on('display-removed', notify); screen.on('display-metrics-changed', notify);
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', async () => {
  for (const s of senders.values()) { try { await (await s.p).destroy(); } catch (e) {} }
  app.quit();
});
