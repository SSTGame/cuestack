// CueStack Desktop – Electron-Hülle um web/index.html
// Zusätzlich: NDI-Ausgabe, HTTP-API für Bitfocus Companion, Office→PDF, Bildschirmauswahl, Vollbild-Ausgänge.
'use strict';
const { app, BrowserWindow, ipcMain, screen, desktopCapturer, session, shell } = require('electron');
const path = require('path');
const http = require('http');
const os = require('os');
const fs = require('fs');
const { execFile, spawn } = require('child_process');
const { htmlSlidesToPdf } = require('./html-slides');

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
  $p.SaveAs($out, 32); $p.Close(); if ($app.Presentations.Count -eq 0) { $app.Quit() }
} elseif ($ext -match 'doc|odt|rtf') {
  $app = New-Object -ComObject Word.Application; $app.Visible = $false
  $d = $app.Documents.Open($in, $false, $true)
  $d.ExportAsFixedFormat($out, 17); $d.Close($false); if ($app.Documents.Count -eq 0) { $app.Quit() }
} elseif ($ext -match 'xls|ods') {
  $app = New-Object -ComObject Excel.Application; $app.Visible = $false; $app.DisplayAlerts = $false
  $w = $app.Workbooks.Open($in, 0, $true)
  $w.ExportAsFixedFormat(0, $out); $w.Close($false); if ($app.Workbooks.Count -eq 0) { $app.Quit() }
} else { throw "Dateityp nicht unterstützt" }
`;
// AppleScripts: der App-Name muss fest im Skript stehen, sonst kennt osascript die Befehle der App nicht
const AS = {
  'Microsoft PowerPoint': `on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  tell application "Microsoft PowerPoint"
    open (POSIX file inPath)
    delay 1
    set p to active presentation
    try
      save p in outPath as save as PDF
    on error
      save p in ((POSIX file outPath) as text) as save as PDF
    end try
    close p saving no
  end tell
end run`,
  'Microsoft Word': `on run argv
  set inPath to item 1 of argv
  set outPath to item 2 of argv
  tell application "Microsoft Word"
    open (POSIX file inPath)
    delay 1
    set d to active document
    save as d file name outPath file format format PDF
    close d saving no
  end tell
end run`,
  Keynote: `on run argv
  set inF to POSIX file (item 1 of argv)
  set outF to POSIX file (item 2 of argv)
  tell application "Keynote"
    set d to open inF
    delay 1
    export d to outF as PDF
    close d saving no
  end tell
end run`,
  Pages: `on run argv
  set inF to POSIX file (item 1 of argv)
  set outF to POSIX file (item 2 of argv)
  tell application "Pages"
    set d to open inF
    delay 1
    export d to outF as PDF
    close d saving no
  end tell
end run`,
  Numbers: `on run argv
  set inF to POSIX file (item 1 of argv)
  set outF to POSIX file (item 2 of argv)
  tell application "Numbers"
    set d to open inF
    delay 1
    export d to outF as PDF
    close d saving no
  end tell
end run`,
};
const MS_CONTAINER = { 'Microsoft PowerPoint': 'com.microsoft.Powerpoint', 'Microsoft Word': 'com.microsoft.Word' };
// Microsoft-Apps auf dem Mac sind „sandboxed“: Dateien in ihrem eigenen Container dürfen sie ohne Rückfrage lesen/schreiben
function macWorkDir(appName) {
  const c = MS_CONTAINER[appName];
  const cands = [];
  if (c) cands.push(path.join(os.homedir(), 'Library', 'Containers', c, 'Data', 'Documents', 'CueStack'));
  cands.push(path.join(os.homedir(), 'Documents', 'CueStack-Umwandlung'));
  for (const d of cands) { try { fs.mkdirSync(d, { recursive: true }); fs.accessSync(d, fs.constants.W_OK); return d; } catch (e) {} }
  return null;
}
/* Mac: Office- und iWork-Apps sind „sandboxed“. Darum wird die Datei wie per Doppelklick über macOS
   übergeben (open -a → die App darf genau diese Datei öffnen) und das PDF im eigenen Ordner der App
   gespeichert (dort darf sie immer schreiben). CueStack holt das PDF dort ab. */
const MAC_CONTAINER = { 'Microsoft PowerPoint': 'com.microsoft.Powerpoint', 'Microsoft Word': 'com.microsoft.Word', Keynote: 'com.apple.iWork.Keynote', Pages: 'com.apple.iWork.Pages', Numbers: 'com.apple.iWork.Numbers' };
const MAC_EXPORT = {
  'Microsoft PowerPoint': `on run argv
  set docName to item 1 of argv
  set docBase to item 2 of argv
  set outPath to item 3 of argv
  tell application "Microsoft PowerPoint"
    set thePres to missing value
    repeat 60 times
      try
        set nameList to name of every presentation
      on error
        set nameList to {}
      end try
      repeat with idx from 1 to count of nameList
        set oneName to (item idx of nameList) as text
        if oneName is docName or oneName is docBase then
          set thePres to presentation idx
          exit repeat
        end if
      end repeat
      if thePres is not missing value then exit repeat
      delay 0.5
    end repeat
    if thePres is missing value then
      try
        if (count of presentations) > 0 then set thePres to active presentation
      end try
    end if
    if thePres is missing value then error "PowerPoint hat die Datei nicht geöffnet (Zeitüberschreitung)."
    activate
    set report to ""
    repeat with k from 3 to count of argv
      set outPath to item k of argv
      set okFlag to false
      try
        save thePres in ((POSIX file outPath) as text) as save as PDF
        set okFlag to true
        set report to report & "HFS ok: " & outPath & linefeed
      on error errA number numA
        set report to report & "HFS " & numA & " " & errA & ": " & outPath & linefeed
      end try
      if not okFlag then
        try
          save thePres in outPath as save as PDF
          set report to report & "POSIX ok: " & outPath & linefeed
        on error errB number numB
          set report to report & "POSIX " & numB & " " & errB & ": " & outPath & linefeed
        end try
      end if
    end repeat
    close thePres saving no
  end tell
  return report
end run`,
  'Microsoft Word': `on run argv
  set docName to item 1 of argv
  set docBase to item 2 of argv
  set outPath to item 3 of argv
  tell application "Microsoft Word"
    set theDoc to missing value
    repeat 60 times
      try
        set nameList to name of every document
      on error
        set nameList to {}
      end try
      repeat with idx from 1 to count of nameList
        set oneName to (item idx of nameList) as text
        if oneName is docName or oneName is docBase then
          set theDoc to document idx
          exit repeat
        end if
      end repeat
      if theDoc is not missing value then exit repeat
      delay 0.5
    end repeat
    if theDoc is missing value then
      try
        if (count of documents) > 0 then set theDoc to active document
      end try
    end if
    if theDoc is missing value then error "Word hat die Datei nicht geöffnet (Zeitüberschreitung)."
    activate
    save as theDoc file name outPath file format format PDF
    close theDoc saving no
  end tell
end run`,
};
for (const a of ['Keynote', 'Pages', 'Numbers']) MAC_EXPORT[a] = `on run argv
  set docName to item 1 of argv
  set docBase to item 2 of argv
  set outPath to item 3 of argv
  tell application "${a}"
    set theDoc to missing value
    repeat 60 times
      try
        set nameList to name of every document
      on error
        set nameList to {}
      end try
      repeat with idx from 1 to count of nameList
        set oneName to (item idx of nameList) as text
        if oneName is docName or oneName is docBase then
          set theDoc to document idx
          exit repeat
        end if
      end repeat
      if theDoc is not missing value then exit repeat
      delay 0.5
    end repeat
    if theDoc is missing value then
      try
        if (count of documents) > 0 then set theDoc to front document
      end try
    end if
    if theDoc is missing value then error "${a} hat die Datei nicht geöffnet (Zeitüberschreitung)."
    export theDoc to (POSIX file outPath) as PDF
    close theDoc saving no
  end tell
end run`;
/* Austauschordner gehört CueStack selbst – dorthin darf CueStack immer schreiben. PowerPoint fragt beim
   ersten Mal einmal „Zugriff gewähren“ für diesen Ordner und merkt sich das danach. */
function macExchangeDir() {
  const d = path.join(app.getPath('userData'), 'Office-Austausch');
  fs.mkdirSync(d, { recursive: true });
  return d;
}
/* Mögliche Ablageorte, in die das sandboxed PowerPoint schreiben darf. Der erste lesbare gewinnt. */
function macTargets(appName, base) {
  const h = os.homedir(), t = [];
  if (/^Microsoft/.test(appName)) {
    t.push(path.join(h, 'Library/Group Containers/UBF8T346G9.Office', base + '.pdf'));
    const c = MAC_CONTAINER[appName];
    if (c) t.push(path.join(h, 'Library/Containers', c, 'Data/Library/Caches', base + '.pdf'));
    t.push(path.join(h, 'Downloads', base + '.pdf'));
  }
  t.push(path.join(macExchangeDir(), base + '.pdf'));
  return t;
}
async function macConvert(appName, srcFile, outF) {
  const dir = macExchangeDir();
  const docName = 'CueStack_' + Date.now() + path.extname(srcFile).toLowerCase();
  const docBase = docName.replace(/\.[^.]+$/, '');
  const inF = path.join(dir, docName);
  fs.copyFileSync(srcFile, inF);
  const targets = /^Microsoft PowerPoint/.test(appName) ? macTargets(appName, docBase) : [path.join(dir, docBase + '.pdf')];
  const as = path.join(dir, 'export-' + appName.replace(/\W/g, '') + '.applescript');
  fs.writeFileSync(as, MAC_EXPORT[appName]);
  const short = appName.replace('Microsoft ', '');
  if (win && !win.isDestroyed()) win.webContents.send('hint', `${short} wandelt um … Falls ${short} nach Zugriff fragt: „Zugriff gewähren“ klicken (nur beim ersten Mal).`);
  try {
    await run('open', ['-g', '-a', appName, inF], 60000);           // macOS übergibt die Datei an die App
    const out = await run('osascript', [as, docName, docBase, ...targets], 300000);
    if (out && String(out).trim()) logLine(short + '-Export: ' + String(out).trim().replace(/\n/g, ' | '));
    let hit = null;
    for (let i = 0; i < 40 && !hit; i++) {
      for (const f of targets) {
        try { const b = fs.readFileSync(f); if (b.length > 100 && b.slice(0, 4).toString() === '%PDF') { hit = f; fs.writeFileSync(outF, b); break; } } catch (e) {}
      }
      if (!hit) await new Promise(r => setTimeout(r, 250));
    }
    if (!hit) {
      // PDF liegt evtl. in PowerPoints geschütztem Bereich: der Finder darf es herauskopieren
      for (const f of targets) {
        if (!fs.existsSync(f) || f.startsWith(dir)) continue;
        try {
          await run('osascript', ['-e', 'on run argv', '-e', 'tell application "Finder" to duplicate (POSIX file (item 1 of argv) as alias) to (POSIX file (item 2 of argv) as alias) with replacing', '-e', 'end run', f, dir], 30000);
          const c = path.join(dir, path.basename(f));
          const b = fs.readFileSync(c);
          if (b.slice(0, 4).toString() === '%PDF') { fs.writeFileSync(outF, b); hit = f + ' (über Finder)'; fs.rm(c, { force: true }, () => {}); break; }
        } catch (e) { logLine('  Finder-Kopie ' + f + ': ' + e.message); }
      }
    }
    for (const f of targets) {
      let st = 'fehlt';
      try { fs.accessSync(f, fs.constants.R_OK); st = 'lesbar'; } catch (e) { if (e.code !== 'ENOENT') st = e.code; }
      logLine('  Ziel ' + f + ': ' + st + (hit && hit.startsWith(f) ? ' ← benutzt' : ''));
    }
    if (!hit) throw new Error('PDF wurde nicht erzeugt (alle Ablageorte geprüft, siehe Protokoll)');
  } finally {
    for (const f of [inF, ...targets]) fs.rm(f, { force: true }, () => {});
    setTimeout(() => { if (win && !win.isDestroyed()) win.focus(); }, 300);
  }
}
/* Ersatz ohne PowerPoint: macOS-Vorschau (Quick Look) erzeugt eine Vorschau der Präsentation */
async function macQuickLook(srcFile, outF) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-ql-'));
  try {
    await run('qlmanage', ['-p', '-o', dir, srcFile], 120000);
    const ql = fs.readdirSync(dir).find(n => n.endsWith('.qlpreview'));
    if (!ql) throw new Error('Quick Look hat keine Vorschau erzeugt');
    const qd = path.join(dir, ql);
    try { logLine('  macOS-Vorschau Inhalt: ' + fs.readdirSync(qd).join(', ')); } catch (e) {}
    const pdf = path.join(qd, 'Preview.pdf');
    if (fs.existsSync(pdf)) { fs.copyFileSync(pdf, outF); return; }
    const html = fs.readdirSync(qd).find(n => /^Preview\.html?$/i.test(n));
    if (!html) throw new Error('Quick Look: weder PDF noch HTML in der Vorschau');
    const n = await htmlSlidesToPdf(path.join(qd, html), outF);
    logLine('  macOS-Vorschau: ' + n + ' Folien erkannt');
    if (!n) throw new Error('macOS-Vorschau: keine einzelnen Folien erkannt');
  } finally { fs.rm(dir, { recursive: true, force: true }, () => {}); }
}
/* Protokoll für Fehler, damit nichts unbemerkt verschwindet */
function logLine(msg) {
  try { const d = app.getPath('logs'); fs.mkdirSync(d, { recursive: true }); fs.appendFileSync(path.join(d, 'cuestack.log'), new Date().toISOString() + '  ' + msg + '\n'); } catch (e) {}
}
ipcMain.handle('log-path', () => { try { return path.join(app.getPath('logs'), 'cuestack.log'); } catch (e) { return ''; } });
ipcMain.handle('log-read', () => { try { const t = fs.readFileSync(path.join(app.getPath('logs'), 'cuestack.log'), 'utf8'); return t.slice(-12000); } catch (e) { return ''; } });
ipcMain.handle('show-log', () => { try { const f = path.join(app.getPath('logs'), 'cuestack.log'); if (fs.existsSync(f)) shell.showItemInFolder(f); else shell.openPath(app.getPath('logs')); } catch (e) {} return true; });
const macHas = (appName) => fs.existsSync(`/Applications/${appName}.app`);
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

/* ---------- PowerPoint mit Effekten → Effekt-Video (Windows, unsichtbar im Hintergrund) ----------
   In einer Kopie wird jede Klick-Animation zu „Nach vorheriger“ mit einer festen Pause (HOLD) und jede
   Folie wechselt nach HOLD automatisch. PowerPoint rendert daraus ein Video mit allen Effekten; CueStack
   findet die Pausen wieder und hält dort an – jeder Klick spielt bis zur nächsten Pause weiter. */
const FX_HOLD = 1.5;
const PS_FX = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$in = $args[0]; $out = $args[1]; $planFile = $args[2]; $hold = [double]$args[3]; $height = [int]$args[4]
$app = New-Object -ComObject PowerPoint.Application
$p = $app.Presentations.Open($in, 0, 0, 0)
$plan = @()
foreach ($s in $p.Slides) {
  $t = $s.SlideShowTransition
  if ($t.Hidden) { continue }
  $steps = @()
  $seq = $s.TimeLine.MainSequence
  for ($i = 1; $i -le $seq.Count; $i++) {
    $e = $seq.Item($i)
    if ($e.Timing.TriggerType -eq 1) {
      $d = [double]$e.Timing.TriggerDelayTime
      $e.Timing.TriggerType = 3
      $e.Timing.TriggerDelayTime = $hold + $d
      $steps += $d
    }
  }
  $endStop = $true
  if ($t.AdvanceOnTime -and -not $t.AdvanceOnClick) { $endStop = $false }
  else { $t.AdvanceOnClick = 0; $t.AdvanceOnTime = -1; $t.AdvanceTime = $hold }
  $plan += @{ slide = $s.SlideIndex; steps = @($steps); endStop = $endStop }
}
$w = $p.PageSetup.SlideWidth; $h = $p.PageSetup.SlideHeight
($(@{ hold = $hold; slides = $plan; w = $w; h = $h }) | ConvertTo-Json -Depth 5 -Compress) | Set-Content -Path $planFile -Encoding UTF8
$p.CreateVideo($out, $true, 5, $height, 30, 90)
$start = Get-Date
while ($true) {
  Start-Sleep -Milliseconds 700
  $st = $p.CreateVideoStatus
  if ($st -eq 3) { break }
  if ($st -eq 4) { throw 'PowerPoint konnte das Video nicht erstellen.' }
  if (((Get-Date) - $start).TotalMinutes -gt 60) { throw 'Zeitüberschreitung beim Videoexport.' }
}
$p.Saved = -1
$p.Close()
if ($app.Presentations.Count -eq 0) { $app.Quit() }
`;
ipcMain.handle('fx-prepare', async (_e, { name, buf, height }) => {
  if (process.platform !== 'win32') return { error: 'mac' };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-fx-'));
  const safe = String(name).replace(/[^\w.\- ]+/g, '_');
  const inF = path.join(dir, safe), outF = path.join(dir, 'effekte.mp4'), planF = path.join(dir, 'plan.json');
  fs.writeFileSync(inF, Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength));
  const ps = path.join(dir, 'fx.ps1'); fs.writeFileSync(ps, '﻿' + PS_FX, 'utf8');
  try {
    await run('powershell.exe', ['-NoProfile', '-Sta', '-ExecutionPolicy', 'Bypass', '-File', ps, inF, outF, planF, String(FX_HOLD), String(height || 1080)], 65 * 60 * 1000);
    if (!fs.existsSync(outF)) throw new Error('Kein Video erzeugt.');
    const plan = JSON.parse(fs.readFileSync(planF, 'utf8').replace(/^﻿/, ''));
    if (!Array.isArray(plan.slides)) plan.slides = plan.slides ? [plan.slides] : [];
    plan.slides.forEach(s => { if (!Array.isArray(s.steps)) s.steps = s.steps == null ? [] : [s.steps]; });
    return { video: fs.readFileSync(outF), plan };
  } catch (e) {
    return { error: /PowerPoint\.Application|80040154|ComObject/i.test(e.message) ? 'Microsoft PowerPoint wurde nicht gefunden.' : e.message };
  } finally { fs.rm(dir, { recursive: true, force: true }, () => {}); }
});

ipcMain.handle('convert-office', async (_e, { name, buf }) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-'));
  const safe = String(name).replace(/[^\w.\- ]+/g, '_');
  const inF = path.join(dir, safe), outF = path.join(dir, safe.replace(/\.[^.]+$/, '') + '.pdf');
  fs.writeFileSync(inF, Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength));
  const errors = [];
  const tryIt = async (label, fn) => { try { await fn(); if (fs.existsSync(outF)) return label; errors.push(label + ': kein PDF erzeugt'); } catch (e) { errors.push(label + ': ' + e.message); } return null; };
  let via = null;
  const ext = path.extname(safe).toLowerCase();
  if (process.platform === 'win32') {
    via = await tryIt('Microsoft Office', async () => {
      const ps = path.join(dir, 'convert.ps1'); fs.writeFileSync(ps, '\ufeff' + PS_SCRIPT, 'utf8');
      await run('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps, inF, outF]);
    });
  } else if (process.platform === 'darwin') {
    const order = /\.(pptx?|ppsx?|odp)$/.test(ext) ? (macHas('Microsoft PowerPoint') ? ['Microsoft PowerPoint'] : ['Keynote'])
      : ext === '.key' ? ['Keynote']
      : /\.(xlsx?|ods)$/.test(ext) ? ['Numbers']
      : (macHas('Microsoft Word') ? ['Microsoft Word'] : ['Pages']);
    for (const appName of order) {
      if (via) break;
      if (macHas(appName)) via = await tryIt(appName, () => macConvert(appName, inF, outF));
    }
  }
  if (!via) {
    const so = findSoffice();
    if (so) via = await tryIt('LibreOffice', () => run(so, ['--headless', '--convert-to', 'pdf', '--outdir', dir, inF]));
  }
  if (!via && process.platform === 'darwin' && /\.(pptx?|ppsx?|key|docx?|xlsx?)$/.test(ext)) via = await tryIt('macOS-Vorschau', () => macQuickLook(inF, outF));
  try {
    if (!via) logLine('Umwandlung ' + name + ' fehlgeschlagen: ' + errors.join(' | '));
    else logLine('Umwandlung ' + name + ' ok über ' + via + (errors.length ? ' (vorher: ' + errors.join(' | ') + ')' : ''));
    if (!via) return { error: errors.length ? errors.join(' | ') : 'Kein Programm zum Umwandeln gefunden. Bitte Microsoft PowerPoint, Keynote oder LibreOffice installieren (LibreOffice ist kostenlos). Auf dem Mac beim ersten Mal die Rückfragen „CueStack möchte … steuern“ bzw. „auf Daten anderer Apps zugreifen“ erlauben.' };
    return { pdf: fs.readFileSync(outF), via };
  } finally { fs.rm(dir, { recursive: true, force: true }, () => {}); }
});


/* ---------- PowerPoint live: echte Bildschirmpräsentation mit allen Effekten, ferngesteuert ---------- */
const PS_LIVE = `
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$app = $null; $pres = $null; $win = $null
function Send($o) { [Console]::Out.WriteLine(($o | ConvertTo-Json -Compress)); [Console]::Out.Flush() }
function State() {
  try { $v = $win.View; return @{ ok = $true; slide = $v.Slide.SlideIndex; total = $pres.Slides.Count; black = ($v.State -eq 3); app = 'PowerPoint' } }
  catch { return @{ ok = $true; ended = $true } }
}
while ($true) {
  $line = [Console]::In.ReadLine()
  if ($line -eq $null) { break }
  try {
    $c = $line | ConvertFrom-Json
    switch ($c.cmd) {
      'open' {
        $app = New-Object -ComObject PowerPoint.Application
        $pres = $app.Presentations.Open($c.path, -1, 0, 0)
        $s = $pres.SlideShowSettings
        $s.ShowType = 1
        try { $s.ShowPresenterView = 0 } catch {}
        $win = $s.Run()
        if ($c.bounds) {
          $win.Left = $c.bounds.x * 0.75; $win.Top = $c.bounds.y * 0.75
          $win.Width = $c.bounds.width * 0.75; $win.Height = $c.bounds.height * 0.75
        }
        Send (State)
      }
      'next'  { $win.View.Next(); Send (State) }
      'prev'  { $win.View.Previous(); Send (State) }
      'goto'  { $win.View.GotoSlide([int]$c.n); Send (State) }
      'black' { if ($c.on) { $win.View.State = 3 } else { $win.View.State = 1 }; Send (State) }
      'state' { Send (State) }
      'end'   { try { $win.View.Exit() } catch {}; try { $pres.Close() } catch {}; Send @{ ok = $true; ended = $true } }
      default { Send @{ ok = $false; error = 'unbekannter Befehl' } }
    }
  } catch { Send @{ ok = $false; error = $_.Exception.Message } }
}
`;
function winBridge() {
  const ps = path.join(os.tmpdir(), 'cuestack-live.ps1');
  fs.writeFileSync(ps, '﻿' + PS_LIVE, 'utf8');
  const proc = spawn('powershell.exe', ['-NoProfile', '-Sta', '-ExecutionPolicy', 'Bypass', '-File', ps], { windowsHide: true });
  let buf = ''; const waiters = [];
  proc.stdout.on('data', (d) => {
    buf += d.toString('utf8'); let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
      if (!line || line[0] !== '{') continue;
      const w = waiters.shift(); if (!w) continue;
      try { w(JSON.parse(line)); } catch (e) { w({ ok: false, error: line.slice(0, 200) }); }
    }
  });
  proc.on('exit', () => { while (waiters.length) waiters.shift()({ ok: false, ended: true, error: 'PowerPoint-Verbindung beendet' }); });
  let chain = Promise.resolve();
  const send = (obj, timeout = 20000) => (chain = chain.then(() => new Promise((res) => {
    const w = (v) => { clearTimeout(t); res(v); };
    const t = setTimeout(() => { const k = waiters.indexOf(w); if (k >= 0) waiters.splice(k, 1); res({ ok: false, error: 'PowerPoint antwortet nicht' }); }, timeout);
    waiters.push(w); proc.stdin.write(JSON.stringify(obj) + '\n');
  })));
  return { kind: 'win', proc, send, close: () => { try { proc.stdin.end(); } catch (e) {} setTimeout(() => { try { proc.kill(); } catch (e) {} }, 3000); } };
}
const MAC_LIVE = {
  'Microsoft PowerPoint': {
    open: `on run argv
  set docName to item 1 of argv
  set docBase to item 2 of argv
  tell application "Microsoft PowerPoint"
    set thePres to missing value
    repeat 60 times
      try
        set nameList to name of every presentation
      on error
        set nameList to {}
      end try
      repeat with idx from 1 to count of nameList
        set oneName to (item idx of nameList) as text
        if oneName is docName or oneName is docBase then
          set thePres to presentation idx
          exit repeat
        end if
      end repeat
      if thePres is not missing value then exit repeat
      delay 0.5
    end repeat
    if thePres is missing value then
      try
        if (count of presentations) > 0 then set thePres to active presentation
      end try
    end if
    if thePres is missing value then error "PowerPoint hat die Datei nicht geöffnet (Zeitüberschreitung)."
    activate
    set showSettings to slide show settings of thePres
    set show type of showSettings to slide show type speaker
    run slide show showSettings
  end tell
end run`,
    next: 'tell application "Microsoft PowerPoint" to go to next slide slide show view of slide show window 1',
    prev: 'tell application "Microsoft PowerPoint" to go to previous slide slide show view of slide show window 1',
    goto: `on run argv
  tell application "Microsoft PowerPoint" to go to slide (slide show view of slide show window 1) number ((item 1 of argv) as integer)
end run`,
    black: `on run argv
  tell application "Microsoft PowerPoint"
    if (item 1 of argv) is "1" then
      set slide state of slide show view of slide show window 1 to slide show state black screen
    else
      set slide state of slide show view of slide show window 1 to slide show state running
    end if
  end tell
end run`,
    state: `tell application "Microsoft PowerPoint"
  set theView to slide show view of slide show window 1
  return ((slide index of slide of theView) as text) & "/" & ((count of slides of active presentation) as text)
end tell`,
    end: 'tell application "Microsoft PowerPoint" to exit slide show slide show view of slide show window 1',
  },
  Keynote: {
    open: `on run argv
  set docName to item 1 of argv
  set docBase to item 2 of argv
  tell application "Keynote"
    set theDoc to missing value
    repeat 60 times
      try
        set nameList to name of every document
      on error
        set nameList to {}
      end try
      repeat with idx from 1 to count of nameList
        set oneName to (item idx of nameList) as text
        if oneName is docName or oneName is docBase then
          set theDoc to document idx
          exit repeat
        end if
      end repeat
      if theDoc is not missing value then exit repeat
      delay 0.5
    end repeat
    if theDoc is missing value then
      try
        if (count of documents) > 0 then set theDoc to front document
      end try
    end if
    if theDoc is missing value then error "Keynote hat die Datei nicht geöffnet (Zeitüberschreitung)."
    activate
    start theDoc from slide 1 of theDoc
  end tell
end run`,
    next: 'tell application "Keynote" to show next',
    prev: 'tell application "Keynote" to show previous',
    goto: `on run argv
  tell application "Keynote" to start front document from slide ((item 1 of argv) as integer) of front document
end run`,
    black: 'return ""',
    state: `tell application "Keynote"
  if not playing then error "beendet"
  return ((slide number of current slide of front document) as text) & "/" & ((count of slides of front document) as text)
end tell`,
    end: 'tell application "Keynote" to stop front document',
  },
};
function macBridge(appName) {
  const scripts = MAC_LIVE[appName];
  const runAs = async (op, args = []) => {
    const f = path.join(os.tmpdir(), `cuestack-live-${appName.replace(/\W/g, '')}-${op}.applescript`);
    fs.writeFileSync(f, scripts[op]); return run('osascript', [f, ...args], op === 'open' ? 120000 : 15000);
  };
  const state = async () => {
    try { const [a, b] = String(await runAs('state')).trim().split('/').map(Number); return { ok: true, slide: a, total: b, app: appName.replace('Microsoft ', '') }; }
    catch (e) { return { ok: true, ended: true }; }
  };
  let chain = Promise.resolve();
  const send = (c) => (chain = chain.then(async () => {
    try {
      if (c.cmd === 'open') {
        const docName = path.basename(c.path), docBase = docName.replace(/\.[^.]+$/, '');
        await run('open', ['-g', '-a', appName, c.path], 60000);   // macOS übergibt die Datei an die App
        await runAs('open', [docName, docBase]); await new Promise(r => setTimeout(r, 1500));
      }
      else if (c.cmd === 'goto') await runAs('goto', [String(c.n)]);
      else if (c.cmd === 'black') await runAs('black', [c.on ? '1' : '0']);
      else if (c.cmd === 'end') { await runAs('end'); return { ok: true, ended: true }; }
      else if (c.cmd !== 'state') await runAs(c.cmd);
      const st = await state(); if (c.cmd === 'black') st.black = !!c.on; return st;
    } catch (e) { logLine('PowerPoint live (' + appName + ') ' + c.cmd + ': ' + e.message); return { ok: false, error: e.message }; }
  }));
  return { kind: 'mac', send, close: () => {} };
}
let live = null;
ipcMain.handle('ppt-live', async (_e, c) => {
  try {
    if (c.cmd === 'open') {
      if (live) { try { await live.send({ cmd: 'end' }); } catch (e) {} live.close(); live = null; }
      const safe = String(c.name || 'praesentation.pptx').replace(/[^\w.\- ]+/g, '_');
      let appName = null;
      if (process.platform === 'darwin') {
        const isKey = /\.key$/i.test(safe);
        appName = !isKey && macHas('Microsoft PowerPoint') ? 'Microsoft PowerPoint' : macHas('Keynote') ? 'Keynote' : null;
        if (!appName) return { ok: false, error: 'Weder Microsoft PowerPoint noch Keynote gefunden.' };
      } else if (process.platform !== 'win32') return { ok: false, error: 'PowerPoint live gibt es nur unter Windows und macOS.' };
      const dir = process.platform === 'darwin' ? macExchangeDir() : path.join(os.tmpdir(), 'cuestack-live');
      fs.mkdirSync(dir, { recursive: true });
      const file = path.join(dir, 'live_' + Date.now() + '_' + safe);
      fs.writeFileSync(file, Buffer.from(c.buf.buffer, c.buf.byteOffset, c.buf.byteLength));
      live = process.platform === 'win32' ? winBridge() : macBridge(appName);
      const r = await live.send({ cmd: 'open', path: file, bounds: c.bounds || null }, 120000);
      setTimeout(() => { if (win && !win.isDestroyed()) win.focus(); }, 1200);
      return r;
    }
    if (!live) return { ok: true, ended: true };
    const r = await live.send(c);
    if (c.cmd === 'end' || r.ended) { const l = live; live = null; l.close(); }
    return r;
  } catch (e) { return { ok: false, error: e.message }; }
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

/* ---------- Updates direkt aus GitHub-Releases ---------- */
const https = require('https');
const REPO = 'SSTGame/cuestack';
function httpsGet(url, { json = false, file = null, onProgress = null } = {}, hops = 0) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers: { 'User-Agent': 'CueStack/' + app.getVersion(), Accept: json ? 'application/vnd.github+json' : 'application/octet-stream' } }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location && hops < 8) {
        res.resume(); return resolve(httpsGet(new URL(res.headers.location, url).toString(), { json, file, onProgress }, hops + 1));
      }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error('HTTP ' + res.statusCode)); }
      const total = Number(res.headers['content-length']) || 0;
      let got = 0, last = 0;
      if (file) {
        const out = fs.createWriteStream(file);
        res.on('data', (c) => { got += c.length; const n = Date.now(); if (onProgress && n - last > 200) { last = n; onProgress(got, total); } });
        res.pipe(out);
        out.on('finish', () => out.close(() => resolve(file)));
        out.on('error', reject); res.on('error', reject);
      } else {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => { const t = Buffer.concat(chunks).toString('utf8'); try { resolve(json ? JSON.parse(t) : t); } catch (e) { reject(e); } });
        res.on('error', reject);
      }
    });
    req.setTimeout(30000, () => req.destroy(new Error('Zeitüberschreitung')));
    req.on('error', reject);
  });
}
function verCmp(a, b) {
  const pa = String(a).replace(/^v/, '').split('.').map(Number), pb = String(b).replace(/^v/, '').split('.').map(Number);
  for (let i = 0; i < 3; i++) { const d = (pa[i] || 0) - (pb[i] || 0); if (d) return d; }
  return 0;
}
let updInfo = null;
ipcMain.handle('update-check', async () => {
  try {
    const r = await httpsGet(`https://api.github.com/repos/${REPO}/releases/latest`, { json: true });
    const re = process.platform === 'darwin' ? /Mac.*\.zip$/i : process.platform === 'win32' ? /Setup.*\.exe$/i : null;
    const asset = re && (r.assets || []).find(a => re.test(a.name));
    const latest = String(r.tag_name || '').replace(/^v/, '');
    updInfo = { current: app.getVersion(), latest, newer: !!latest && verCmp(latest, app.getVersion()) > 0, url: asset && asset.browser_download_url, size: asset && asset.size, asset: asset && asset.name, notes: String(r.body || '').slice(0, 2000), page: r.html_url };
    if (updInfo.newer && !asset) updInfo.error = process.platform === 'darwin' ? 'Für diese Version gibt es noch kein Mac-Update-Paket – bitte einmal die DMG von der Release-Seite installieren.' : 'Kein Update-Paket für dieses System gefunden.';
    return updInfo;
  } catch (e) { return { current: app.getVersion(), error: 'Update-Prüfung fehlgeschlagen: ' + e.message }; }
});
ipcMain.handle('update-install', async () => {
  const send = (m) => { if (win && !win.isDestroyed()) win.webContents.send('update-progress', m); };
  try {
    if (!updInfo || !updInfo.url) throw new Error('Kein Update gefunden');
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cuestack-update-'));
    const file = path.join(dir, updInfo.asset);
    await httpsGet(updInfo.url, { file, onProgress: (g, t) => send({ phase: 'download', got: g, total: t || updInfo.size || 0 }) });
    send({ phase: 'install' });
    logLine('Update ' + updInfo.current + ' → ' + updInfo.latest + ' geladen: ' + file);
    if (process.platform === 'darwin') {
      const appPath = path.resolve(process.execPath, '../../..');           // …/CueStack.app
      if (!/\.app$/.test(appPath)) throw new Error('App-Pfad nicht erkannt: ' + appPath);
      const ex = path.join(dir, 'x'); fs.mkdirSync(ex);
      await run('ditto', ['-x', '-k', file, ex], 300000);
      const newApp = fs.readdirSync(ex).find(n => n.endsWith('.app'));
      if (!newApp) throw new Error('Im Update-Paket ist keine App');
      try { fs.accessSync(path.dirname(appPath), fs.constants.W_OK); } catch (e) { throw new Error('Keine Schreibrechte für ' + path.dirname(appPath) + ' – bitte die DMG manuell installieren.'); }
      const sh = path.join(dir, 'update.sh');
      fs.writeFileSync(sh, `#!/bin/sh
while kill -0 ${process.pid} 2>/dev/null; do sleep 0.3; done
sleep 0.5
rm -rf "${appPath}.old"
mv "${appPath}" "${appPath}.old" || exit 1
if ditto "${path.join(ex, newApp)}" "${appPath}"; then rm -rf "${appPath}.old"; else rm -rf "${appPath}"; mv "${appPath}.old" "${appPath}"; fi
xattr -cr "${appPath}" 2>/dev/null
open "${appPath}"
rm -rf "${dir}"
`, { mode: 0o755 });
      spawn('/bin/sh', [sh], { detached: true, stdio: 'ignore' }).unref();
    } else if (process.platform === 'win32') {
      const cmd = path.join(dir, 'update.cmd');
      fs.writeFileSync(cmd, `@echo off\r\n:w\r\ntasklist /FI "PID eq ${process.pid}" | find "${process.pid}" >nul && (timeout /t 1 /nobreak >nul & goto w)\r\n"${file}" /S\r\nstart "" "${process.execPath}"\r\n`);
      spawn('cmd.exe', ['/c', cmd], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
    } else throw new Error('Auf diesem System nicht unterstützt');
    setTimeout(() => app.quit(), 400);
    return { ok: true };
  } catch (e) { logLine('Update fehlgeschlagen: ' + e.message); send({ phase: 'error', error: e.message }); return { error: e.message }; }
});

app.whenReady().then(() => {
  setupCapture();
  createWindow();
  startHttp();
  const notify = () => win && win.webContents.send('displays-changed');
  screen.on('display-added', notify); screen.on('display-removed', notify); screen.on('display-metrics-changed', notify);
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', async () => {
  if (live) { try { await live.send({ cmd: 'end' }); } catch (e) {} live.close(); live = null; }
  for (const s of senders.values()) { try { await (await s.p).destroy(); } catch (e) {} }
  app.quit();
});
