'use strict';
const { contextBridge, ipcRenderer } = require('electron');

const reply = (rid, value) => ipcRenderer.send('reply', { rid, value });

contextBridge.exposeInMainWorld('CueNative', {
  info: () => ipcRenderer.invoke('info'),
  displays: () => ipcRenderer.invoke('displays'),
  onDisplaysChanged: (cb) => ipcRenderer.on('displays-changed', () => cb()),
  ndiFrame: (id, name, w, h, fps, buf) => ipcRenderer.invoke('ndi-frame', { id, name, w, h, fps, buf: new Uint8Array(buf) }),
  ndiStop: (id) => ipcRenderer.invoke('ndi-stop', id),
  pushState: (s) => ipcRenderer.send('state', s),
  onRemote: (cb) => ipcRenderer.on('remote', (_e, { rid, cmd, p }) => {
    let r; try { r = cb(cmd, p); } catch (err) { r = { ok: false, error: String(err && err.message || err) }; }
    reply(rid, r);
  }),
  onPickSource: (cb) => ipcRenderer.on('pick-source', async (_e, { rid, list }) => {
    let id = null; try { id = await cb(list); } catch (err) {}
    reply(rid, id);
  }),
  convertOffice: (name, buf) => ipcRenderer.invoke('convert-office', { name, buf: new Uint8Array(buf) }),
});
