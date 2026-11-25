/**
 * Mállalo X 2.0 · Servidor de datos del equipo (Google Apps Script)
 *
 * Guarda todo lo que el equipo apunta en la app (checklist, stock, compras,
 * tareas añadidas, fichas editadas) en la hoja «Mállalo X 2.0 · Datos» de tu
 * Google Drive, y las fotos nuevas en la carpeta «Mállalo X 2.0 · Fotos».
 *
 * Implementar > Gestionar implementaciones > Editar > Versión: Nueva > Implementar
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 */

const TZ = 'Europe/Madrid';
const COLS = ['cl', 'st', 'inv', 'compra', 'kx', 'fx'];

function doGet(e) {
  return json_(all_());
}

function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.op === 'set') out = set_(b.col, b.id, b.data, b.who);
    else if (b.op === 'del') out = del_(b.col, b.id, b.who);
    else if (b.op === 'seed') out = seed_(b.data);
    else if (b.op === 'upload') out = upload_(b.name, b.type, b.b64);
    else if (b.op === 'unupload') out = unupload_(b.id);
    else out = { error: 'op' };
  } catch (err) { out = { error: String(err) }; }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DATOS2');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('Mállalo X 2.0 · Datos');
  props.setProperty('DATOS2', ss.getId());
  const sh = ss.getSheets()[0];
  sh.setName('Datos');
  sh.getRange(1, 1, 1, 4).setValues([['coleccion', 'id', 'datos', 'actualizado']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:C').setNumberFormat('@');
  const his = ss.insertSheet('Historial');
  his.appendRow(['fecha', 'quien', 'coleccion', 'id', 'datos']);
  return ss;
}

function rows_(sh) {
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, 3).getValues();
}

function all_() {
  const sh = ss_().getSheetByName('Datos');
  const out = {};
  COLS.forEach(function (c) { out[c] = {}; });
  rows_(sh).forEach(function (r) {
    if (!r[0] || !r[1] || !out[r[0]]) return;
    try { out[r[0]][r[1]] = JSON.parse(r[2]); } catch (e) {}
  });
  return { t: Date.now(), d: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), data: out, empty: sh.getLastRow() < 2 };
}

function check_(col, id) {
  if (COLS.indexOf(col) < 0) throw new Error('coleccion');
  id = String(id || '');
  if (!id || id.length > 120) throw new Error('id');
  return id;
}

function set_(col, id, data, who) {
  id = check_(col, id);
  const s = JSON.stringify(data || {});
  if (s.length > 45000) throw new Error('demasiado grande');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_(), sh = ss.getSheetByName('Datos');
    const r = rows_(sh);
    let row = 0;
    for (let i = 0; i < r.length; i++) if (r[i][0] === col && r[i][1] === id) { row = i + 2; break; }
    const v = [[col, id, s, new Date()]];
    if (row) sh.getRange(row, 1, 1, 4).setValues(v); else sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues(v);
    if (col !== 'cl') ss.getSheetByName('Historial').appendRow([new Date(), String(who || '').slice(0, 40), col, id, s.slice(0, 500)]);
  } finally { lock.releaseLock(); }
  return { ok: 1 };
}

function del_(col, id, who) {
  id = check_(col, id);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_(), sh = ss.getSheetByName('Datos');
    const r = rows_(sh);
    for (let i = 0; i < r.length; i++) if (r[i][0] === col && r[i][1] === id) { sh.deleteRow(i + 2); break; }
    ss.getSheetByName('Historial').appendRow([new Date(), String(who || '').slice(0, 40), col, id, '(borrado)']);
  } finally { lock.releaseLock(); }
  return { ok: 1 };
}

function seed_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ss_().getSheetByName('Datos');
    if (sh.getLastRow() > 1) return { ok: 0, skipped: 1 };
    const rows = [];
    Object.keys(data || {}).forEach(function (col) {
      if (COLS.indexOf(col) < 0) return;
      Object.keys(data[col]).forEach(function (id) { rows.push([col, String(id), JSON.stringify(data[col][id]), new Date()]); });
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    return { ok: 1, n: rows.length };
  } finally { lock.releaseLock(); }
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOTOS2');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Mállalo X 2.0 · Fotos');
  props.setProperty('FOTOS2', f.getId());
  return f;
}

function upload_(name, type, b64) {
  if (!/^image\//.test(type || '')) throw new Error('tipo');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), type, String(name || 'foto.jpg').slice(0, 80));
  const file = folder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return { id: id, url: 'https://lh3.googleusercontent.com/d/' + id + '=w1200' };
}

function unupload_(id) {
  try {
    const f = DriveApp.getFileById(String(id));
    if (f.getParents().hasNext() && f.getParents().next().getId() === folder_().getId()) f.setTrashed(true);
  } catch (e) {}
  return { ok: 1 };
}
