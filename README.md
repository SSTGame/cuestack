# CueStack – kostenloser Show-Cue-Helfer

Alle Cues einer Show in einer App: **Timer, PowerPoint/PDF-Folien mit Videoclips, Dokumente, Lyrics, Referentenansicht, Stage-Monitor, 4× Desktop-Capture und 4 Mixe**. Jedes Modul hat seinen eigenen Ausgang – als Fenster, Vollbild auf Projektor/LED oder als **NDI** (mit Transparenz) für Resolume, vMix, OBS und TouchDesigner.

- **Web-Version (ohne Installation):** https://sstgame.github.io/cuestack/
- **Desktop-App für Windows und Mac:** [Releases](https://github.com/SSTGame/cuestack/releases/latest)

## Was kann was?

| Funktion | Web | Desktop |
|---|:-:|:-:|
| Timer (Countdown, Hochzählen, Uhr, Warnungen, „TIME IS UP“) | ✓ | ✓ |
| **4 Folien-Stapel** (PDF/Bilder/PowerPoint), 9 Übergänge, Autoplay, Clicker | ✓ | ✓ |
| **iPhone-Fotos (HEIC/HEIF)** werden automatisch umgewandelt | ✓ | ✓ |
| Videoclips (MP4/WebM) zwischen den Folien | ✓ | ✓ |
| **PowerPoint (.pptx), Word, Excel direkt öffnen** | – (als PDF exportieren) | ✓ |
| **PowerPoint live**: echte Bildschirmpräsentation mit allen Animationen/Effekten, ferngesteuert per Clicker/Companion | – | ✓ |
| Referentenansicht mit Sprechernotizen aus der .pptx | ✓ | ✓ (automatisch) |
| Stage-Monitor mit Nachrichten an die Bühne | ✓ | ✓ |
| **4 Dokumente** mit weichem Scrollen, Zoom, Teleprompter | ✓ | ✓ |
| Lyrics (.lrc/.srt synchron, Online-Suche, 30 Stile) | ✓ | ✓ |
| 4 Medien-Slots für Bilder/Videos (Logo, Pausenbild, Einspieler) | ✓ | ✓ |
| 4× Desktop-Capture, 4 Mixe à 8 Ebenen | ✓ | ✓ |
| **Mix-Szenen**: Abläufe wie Bild → PowerPoint 1 → Bild → PowerPoint 2 → PDF, per GO, Clicker oder Companion | ✓ | ✓ |
| Ausgänge als Fenster / Vollbild, Testbild, Chroma-Hintergrund | ✓ | ✓ (Vollbild automatisch) |
| **NDI-Ausgabe** inkl. Alpha-Transparenz | – | ✓ |
| **Bitfocus Companion / HTTP-Fernsteuerung**, Handy-Fernbedienung | – | ✓ |
| BLACKOUT, Absturz-Wiederherstellung, Show-Dateien | ✓ | ✓ |

## Desktop-App installieren

**Windows:** `CueStack-Setup-x.y.z.exe` herunterladen und starten. Da die App nicht kostenpflichtig signiert ist, zeigt Windows „Der Computer wurde geschützt“ → **Weitere Informationen → Trotzdem ausführen**.

**Mac (Apple Silicon M1–M4):** `.dmg` öffnen, CueStack in „Programme“ ziehen. Beim ersten Start: Rechtsklick auf die App → **Öffnen**. Falls macOS „beschädigt“ meldet: Terminal öffnen und `xattr -cr /Applications/CueStack.app` eingeben. Für Intel-Macs bitte die Web-Version nutzen.

**PowerPoint als Folien:** Windows nutzt das installierte Microsoft PowerPoint, der Mac Microsoft PowerPoint oder Keynote. Fehlt das, springt das kostenlose [LibreOffice](https://de.libreoffice.org/) ein. Folien werden dabei Standbilder (ideal für Mix-Szenen, NDI, Stage-Monitor). Sprechernotizen landen automatisch in der Referentenansicht. Mac: beim ersten Mal die Rückfragen „CueStack möchte PowerPoint steuern“ und ggf. „auf Daten anderer Apps zugreifen“ erlauben.

**PowerPoint live (mit Effekten):** Modul „PowerPoint live“ startet die Präsentation im echten PowerPoint (Windows) bzw. PowerPoint/Keynote (Mac) – alle Animationen, Übergänge und Videos laufen original. CueStack steuert sie per Clicker, Companion (`ppt/next`, `ppt/prev`, `ppt/goto?n=3`, `ppt/black`, `ppt/end`) und BLACKOUT und zeigt die Notizen in der Referentenansicht. Für NDI oder Mix-Szenen den PowerPoint-Bildschirm in einen Capture-Slot holen.

## Bitfocus Companion

Modul **Generic HTTP** anlegen, Basis-URL `http://<IP-des-CueStack-Rechners>:9876/api/`. Die IP steht in der App unter „Bitfocus Companion / Fernsteuerung“. Unter `http://<IP>:9876/` gibt es außerdem eine Handy-Fernbedienung.

| Befehl | Wirkung |
|---|---|
| `next` / `prev` | weiter/zurück (je nach „Clicker steuert“) |
| `slides/next`, `slides/prev`, `slides/goto?n=3`, `slides/video` | Folien, Video Play/Pause – mit `&d=2` für Folien 2 (ohne `d`: zuletzt benutzter Stapel) |
| `doc/next`, `doc/prev`, `doc/goto?n=2`, `doc/scroll?px=300` | Dokument – mit `&d=2` für Dokument 2 |
| `lyrics/next`, `lyrics/prev`, `lyrics/play`, `lyrics/pause`, `lyrics/stop`, `lyrics/blank` | Lyrics |
| `timer/start`, `timer/pause`, `timer/toggle`, `timer/reset` | Timer |
| `timer/set?t=10:00` oder `timer/set?sec=600`, `timer/add?sec=60`, `timer/mode?m=down\|up\|clock` | Timer stellen |
| `blackout/toggle`, `blackout/on`, `blackout/off` | Blackout |
| `stage/message?text=Noch%205%20Minuten`, `stage/clear` | Nachricht auf Stage-Monitor |
| `mix/next?m=1`, `mix/prev?m=1`, `mix/scene?m=1&s=3`, `mix/step?m=1` | Mix-Szenen schalten (`step` blättert erst durch Folien/Dokument) |
| `clicker?target=auto\|slides\|doc\|lyrics\|mix1` | Ziel des Clickers |
| `state` | Status als JSON (Timer-Text, Folie, Blackout …) – für Companion-Feedbacks/Variablen |

GET und POST funktionieren beide. Port ändern: Umgebungsvariable `CUESTACK_PORT`. Nur lokal erreichbar machen: `CUESTACK_HOST=127.0.0.1`.

## Signalwege

- **Projektor/LED:** Ausgang mit gewähltem Bildschirm öffnen (Desktop: automatisch Vollbild, Web: Doppelklick).
- **Resolume / vMix / TouchDesigner / OBS:** Desktop-App → „+ NDI-Ausgang“, Hintergrund „Transparent“ für Overlays mit Alpha.
- **Web-Version → OBS:** Fensteraufnahme des Ausgangsfensters, für Freistellung Hintergrund Grün + Chroma Key.

## Selbst bauen

```bash
npm ci          # lädt Electron + NDI-SDK, baut das NDI-Modul
npm start       # Desktop-App starten
npm run dist    # Installer bauen (Ausgabe in dist/)
```

Jeder Push auf `main` baut auf GitHub automatisch Windows- und Mac-Installer und hängt sie an das Release zur Version aus `package.json` (Version erhöhen = neues Release). Änderungen in `web/` gehen automatisch auf GitHub Pages.

## Lizenzen

CueStack: MIT. Enthält [PDF.js](https://mozilla.github.io/pdf.js/) (Apache 2.0) [grandiose](https://github.com/stagetimerio/grandiose) (Apache 2.0) und [heic-to](https://github.com/hoppergee/heic-to) (LGPL-3.0, unverändert als eigene Datei unter `web/vendor/heic/`). NDI® ist eine eingetragene Marke von Vizrt NDI AB – https://ndi.video. Lyrics-Suche über [LRCLIB](https://lrclib.net).
