const SHEET_NAMES = {
  SESIONES: 'Sesiones',
  PROFESIONALES: 'Profesionales',
  TIPOS: 'TiposSesion',
  HISTORIAL: 'HistorialCambios'
};

const HEADERS = {
  SESIONES: ['idSesion', 'fecha', 'hora', 'duracionMin', 'tipoSesion', 'tema', 'ponente', 'area', 'estado', 'sustituto', 'enlaceMaterial', 'observaciones', 'fechaCreacion', 'ultimaModificacion', 'modificadoPor'],
  PROFESIONALES: ['nombre', 'area', 'email', 'activo', 'observaciones'],
  TIPOS: ['tipoSesion', 'activo'],
  HISTORIAL: ['fechaCambio', 'idSesion', 'campoModificado', 'valorAnterior', 'valorNuevo', 'usuario', 'motivoCambio']
};

const TZ = 'Europe/Madrid';
const ESTADOS_SESION = ['Programada', 'Realizada', 'Cancelada'];

function doGet() { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Sesiones Clínicas - Servicio de Farmacia').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getInitialData() {
  const diagnostics = { timestamp: new Date().toISOString(), sheetNamesFound: [], headers: {}, errors: [], duplicatedSessionIds: [], duplicatedSessions: [] };
  try {
    const ss = getSpreadsheet_();
    diagnostics.sheetNamesFound = ss.getSheets().map((s) => s.getName());
    const profesionales = getProfesionalesActivosFromSheet_(getSheetOrThrow_(ss, SHEET_NAMES.PROFESIONALES));
    const tiposSesion = getTiposActivosFromSheet_(getSheetOrThrow_(ss, SHEET_NAMES.TIPOS));
    const sesionesResult = getSesionesFromSheet_(getSheetOrThrow_(ss, SHEET_NAMES.SESIONES));
    diagnostics.duplicatedSessionIds = sesionesResult.duplicatedSessionIds;
    diagnostics.duplicatedSessions = sesionesResult.duplicates;
    return { ok: true, profesionales, tiposSesion, estados: ESTADOS_SESION.slice(), sesiones: sesionesResult.sesiones, incidenciasSesiones: sesionesResult.incidencias, diagnostics };
  } catch (err) {
    diagnostics.errors.push(err && err.message ? err.message : String(err));
    console.error('[getInitialData] error', err);
    return { ok: false, profesionales: [], tiposSesion: [], estados: ESTADOS_SESION.slice(), sesiones: [], incidenciasSesiones: [], message: "No se pudieron cargar los datos maestros. Revise las pestañas 'Profesionales' y 'TiposSesion', sus encabezados y filas activas ('Sí').", diagnostics };
  }
}

function getSpreadsheet_() { return (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function getSheetOrThrow_(ss, name) { const sh = ss.getSheetByName(name); if (!sh) throw new Error(`No existe la pestaña ${name}.`); return sh; }
function getSheetHeaders_(sheet) { if (sheet.getLastRow() < 1) return []; return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((v) => String(v || '')); }
function normalizeText_(v) { return v === null || v === undefined ? '' : String(v).replace(/\uFEFF/g, '').replace(/[\u200B-\u200D\u2060]/g, '').trim(); }
function normalizeHeader_(v) { return normalizeText_(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '').toLowerCase(); }
function isActive_(v) { if (v === true) return true; return ['sí', 'si', 'true', '1', 'activo'].includes(normalizeText_(v).toLowerCase()); }
function formatDateTime_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); }
function getUserEmail_() { return Session.getActiveUser().getEmail() || 'usuario_no_identificado'; }

function getHeaderMap_(headers) {
  const normalizedToIndex = {};
  headers.forEach((h, i) => { const n = normalizeHeader_(h); if (n && normalizedToIndex[n] === undefined) normalizedToIndex[n] = i; });
  const map = {};
  [].concat(HEADERS.SESIONES, HEADERS.PROFESIONALES, HEADERS.TIPOS, HEADERS.HISTORIAL).forEach((c) => {
    const n = normalizeHeader_(c);
    if (normalizedToIndex[n] !== undefined) map[c] = normalizedToIndex[n];
  });
  return map;
}

function getRowsAsObjects_(sheet, expectedHeaders) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headerMap = getHeaderMap_(values[0]);
  const missing = expectedHeaders.filter((h) => headerMap[h] === undefined);
  if (missing.length) throw new Error(`Encabezados inválidos en ${sheet.getName()}. Faltan: ${missing.join(', ')}`);
  return values.slice(1).filter((r) => r.some((c) => normalizeText_(c) !== '')).map((row) => {
    const obj = {};
    expectedHeaders.forEach((h) => { obj[h] = row[headerMap[h]]; });
    return obj;
  });
}

function getProfesionalesActivosFromSheet_(sheet) { return getRowsAsObjects_(sheet, HEADERS.PROFESIONALES).filter((r) => isActive_(r.activo) && normalizeText_(r.nombre)).map((r) => ({ nombre: normalizeText_(r.nombre), area: normalizeText_(r.area), email: normalizeText_(r.email) })).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); }
function getTiposActivosFromSheet_(sheet) { return getRowsAsObjects_(sheet, HEADERS.TIPOS).filter((r) => isActive_(r.activo) && normalizeText_(r.tipoSesion)).map((r) => normalizeText_(r.tipoSesion)).sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' })); }

function normalizeEstado_(e) { const t = normalizeText_(e).toLowerCase(); if (t === 'programada') return 'Programada'; if (t === 'realizada') return 'Realizada'; if (t === 'cancelada') return 'Cancelada'; throw new Error(`Estado inválido: ${e}.`); }
function parseDateTimeSafe_(v) { const txt = normalizeText_(v); if (!txt) return null; const d = new Date(txt.replace(' ', 'T')); return isNaN(d.getTime()) ? null : d.getTime(); }
function parseDateForSheet_(value) { const d = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(normalizeText_(value)); if (isNaN(d.getTime())) throw new Error('Fecha inválida. Debe tener formato yyyy-MM-dd.'); return d; }
function normalizeDate_(v) { return Utilities.formatDate(parseDateForSheet_(v), TZ, 'yyyy-MM-dd'); }
function formatDateForInputClient(value) { return normalizeDate_(value); }
function normalizeTime_(value) {
  if (value === null || value === undefined || normalizeText_(value) === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return Utilities.formatDate(value, TZ, 'HH:mm');
  if (typeof value === 'number' && Number.isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60); const m = ((totalMinutes % 1440) + 1440) % 1440;
    return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
  }
  const m = normalizeText_(value).match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if (!m) throw new Error('Hora inválida. Use HH:mm.');
  const hh = Number(m[1]); const mm = m[2] === undefined ? 0 : Number(m[2]); const ss = m[3] === undefined ? 0 : Number(m[3]);
  if (![hh, mm, ss].every(Number.isInteger) || hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) throw new Error('Hora inválida. Use HH:mm con valores entre 00:00 y 23:59.');
  return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}
function parseTimeForDisplay_(value) { return normalizeTime_(value); }
function formatTimeForInputClient(value) { return normalizeTime_(value); }

function pickCanonicalSessionRecord_(records) {
  let canonical = records[records.length - 1];
  let reason = 'last physical row';
  const dated = records.filter((r) => parseDateTimeSafe_(r.ultimaModificacion) !== null);
  if (dated.length) {
    canonical = dated.sort((a, b) => parseDateTimeSafe_(b.ultimaModificacion) - parseDateTimeSafe_(a.ultimaModificacion))[0];
    reason = 'latest ultimaModificacion';
  }
  return { canonical, reason };
}

function getSesionesFromSheet_(sheet) {
  const rows = getRowsAsObjects_(sheet, HEADERS.SESIONES);
  const byId = {};
  const incidencias = [];
  rows.forEach((row, idx) => {
    try {
      const idSesion = normalizeText_(row.idSesion); if (!idSesion) throw new Error('idSesion vacío.');
      const rec = { _rowNumber: idx + 2, idSesion, fecha: formatDateForInputClient(row.fecha), hora: formatTimeForInputClient(row.hora), duracionMin: Number(row.duracionMin) || 0, tipoSesion: normalizeText_(row.tipoSesion), tema: normalizeText_(row.tema), ponente: normalizeText_(row.ponente), area: normalizeText_(row.area), estado: normalizeEstado_(row.estado || 'Programada'), sustituto: normalizeText_(row.sustituto), enlaceMaterial: normalizeText_(row.enlaceMaterial), observaciones: normalizeText_(row.observaciones), fechaCreacion: normalizeText_(row.fechaCreacion), ultimaModificacion: normalizeText_(row.ultimaModificacion), modificadoPor: normalizeText_(row.modificadoPor) };
      byId[idSesion] = byId[idSesion] || [];
      byId[idSesion].push(rec);
    } catch (e) { incidencias.push(`Fila ${idx + 2}: ${e.message || e}`); }
  });
  const duplicates = [];
  const sesiones = Object.keys(byId).map((id) => {
    const { canonical, reason } = pickCanonicalSessionRecord_(byId[id]);
    if (byId[id].length > 1) {
      duplicates.push({ idSesion: id, rows: byId[id].map((r) => r._rowNumber), estados: Array.from(new Set(byId[id].map((r) => r.estado))), fechas: byId[id].map((r) => r.fecha), horas: byId[id].map((r) => r.hora), ultimaModificacion: byId[id].map((r) => r.ultimaModificacion), canonicalRow: canonical._rowNumber, reasonCanonical: reason });
    }
    const out = { ...canonical }; delete out._rowNumber; return out;
  }).sort((a, b) => `${a.fecha} ${a.hora}`.localeCompare(`${b.fecha} ${b.hora}`));
  return { sesiones, incidencias, duplicatedSessionIds: duplicates.map((d) => d.idSesion), duplicates };
}

function validateSesion_(payload, isEdit) {
  ['fecha', 'hora', 'tipoSesion', 'tema', 'ponente', 'estado'].forEach((f) => { if (!normalizeText_(payload[f])) throw new Error(`El campo ${f} es obligatorio.`); });
  const dur = Number(payload.duracionMin); if (!Number.isFinite(dur) || dur <= 0) throw new Error('duracionMin debe ser numérico y mayor que 0.');
  normalizeDate_(payload.fecha); normalizeTime_(payload.hora); normalizeEstado_(payload.estado); if (isEdit && !normalizeText_(payload.idSesion)) throw new Error('idSesion es obligatorio.');
}

function getExistingSessionIds_(sheet) {
  const v = sheet.getDataRange().getValues(); if (v.length < 2) return {};
  const idx = getHeaderMap_(v[0]); const set = {};
  for (let i = 1; i < v.length; i++) { const id = normalizeText_(v[i][idx.idSesion]); if (id) set[id] = true; }
  return set;
}

function generateUniqueIdSesion_(sheet, fecha) {
  const year = String(fecha).slice(0, 4); const ids = getExistingSessionIds_(sheet);
  let seq = 1;
  Object.keys(ids).forEach((id) => { const m = id.match(new RegExp(`^SES-${year}-(\\d+)$`)); if (m) seq = Math.max(seq, Number(m[1]) + 1); });
  let candidate = `SES-${year}-${String(seq).padStart(3, '0')}`;
  while (ids[candidate]) { seq += 1; candidate = `SES-${year}-${String(seq).padStart(3, '0')}`; }
  return candidate;
}

function rowArrayToObj_(headers, row) { const o = {}; headers.forEach((h, i) => { o[h] = row[i]; }); return o; }

function findSessionRowById_(sheet, idSesion) {
  const idBuscado = normalizeText_(idSesion); if (!idBuscado) throw new Error('idSesion es obligatorio.');
  const values = sheet.getDataRange().getValues(); if (values.length < 2) throw new Error('No hay sesiones registradas.');
  const idx = getHeaderMap_(values[0]); if (idx.idSesion === undefined) throw new Error("No existe la columna 'idSesion' en Sesiones.");
  const records = [];
  for (let i = 1; i < values.length; i++) if (normalizeText_(values[i][idx.idSesion]) === idBuscado) records.push({ row: i + 1, ultimaModificacion: normalizeText_(values[i][idx.ultimaModificacion]) });
  if (!records.length) throw new Error(`No se encontró la sesión con idSesion ${idBuscado}.`);
  const c = pickCanonicalSessionRecord_(records.map((r) => ({ _rowNumber: r.row, ultimaModificacion: r.ultimaModificacion })));
  const duplicates = records.length > 1 ? [{ idSesion: idBuscado, rows: records.map((r) => r.row), canonicalRow: c.canonical._rowNumber, reasonCanonical: c.reason }] : [];
  return { rowIndex: c.canonical._rowNumber, diagnostics: { duplicatedSessionIds: duplicates.map((d) => d.idSesion), duplicatedSessions: duplicates } };
}

function resolveArea_(ponente, areaInput) { const area = normalizeText_(areaInput); if (area) return area; const p = getProfesionalesActivosFromSheet_(getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.PROFESIONALES)).find((x) => normalizeText_(x.nombre).toLowerCase() === normalizeText_(ponente).toLowerCase()); return p ? normalizeText_(p.area) : ''; }

function logCambio_(idSesion, campo, valorAnterior, valorNuevo, usuario, motivoCambio) { getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.HISTORIAL).appendRow([formatDateTime_(new Date()), idSesion, campo, normalizeText_(valorAnterior), normalizeText_(valorNuevo), usuario, motivoCambio]); }
function logCambiosBatch_(idSesion, oldObj, newObj, motivo, user) { ['fecha', 'hora', 'duracionMin', 'tipoSesion', 'tema', 'ponente', 'area', 'estado', 'sustituto', 'enlaceMaterial', 'observaciones'].forEach((campo) => { const a = oldObj ? normalizeText_(oldObj[campo]) : ''; const n = normalizeText_(newObj[campo]); if (!oldObj || a !== n) logCambio_(idSesion, campo, a, n, user, motivo); }); }

function createSesion(payload) {
  validateSesion_(payload); const ss = getSpreadsheet_(); const sheet = getSheetOrThrow_(ss, SHEET_NAMES.SESIONES);
  const now = new Date(); const user = getUserEmail_(); const fecha = normalizeDate_(payload.fecha); const idSesion = generateUniqueIdSesion_(sheet, fecha);
  const row = { idSesion, fecha, hora: normalizeTime_(payload.hora), duracionMin: Number(payload.duracionMin), tipoSesion: normalizeText_(payload.tipoSesion), tema: normalizeText_(payload.tema), ponente: normalizeText_(payload.ponente), area: resolveArea_(payload.ponente, payload.area), estado: normalizeEstado_(payload.estado), sustituto: normalizeText_(payload.sustituto), enlaceMaterial: normalizeText_(payload.enlaceMaterial), observaciones: normalizeText_(payload.observaciones), fechaCreacion: formatDateTime_(now), ultimaModificacion: formatDateTime_(now), modificadoPor: user };
  sheet.appendRow(HEADERS.SESIONES.map((h) => row[h]));
  logCambiosBatch_(idSesion, null, row, 'Creación de sesión', user);
  return { ok: true, idSesion, diagnostics: { duplicatedSessionIds: [] } };
}

function updateSesion(payload) {
  validateSesion_(payload, true);
  const sheet = getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.SESIONES);
  const located = findSessionRowById_(sheet, payload.idSesion);
  const headers = getSheetHeaders_(sheet);
  const oldRow = rowArrayToObj_(headers, sheet.getRange(located.rowIndex, 1, 1, headers.length).getValues()[0]);
  const user = getUserEmail_();
  const updated = { ...oldRow, idSesion: normalizeText_(payload.idSesion), fecha: normalizeDate_(payload.fecha), hora: normalizeTime_(payload.hora), duracionMin: Number(payload.duracionMin), tipoSesion: normalizeText_(payload.tipoSesion), tema: normalizeText_(payload.tema), ponente: normalizeText_(payload.ponente), area: resolveArea_(payload.ponente, payload.area), estado: normalizeEstado_(payload.estado), sustituto: normalizeText_(payload.sustituto), enlaceMaterial: normalizeText_(payload.enlaceMaterial), observaciones: normalizeText_(payload.observaciones), ultimaModificacion: formatDateTime_(new Date()), modificadoPor: user };
  sheet.getRange(located.rowIndex, 1, 1, headers.length).setValues([headers.map((h) => updated[h])]);
  logCambiosBatch_(payload.idSesion, oldRow, updated, 'Edición de sesión', user);
  return { ok: true, diagnostics: located.diagnostics };
}

function updateSessionById_(idSesion, patch, motivoCambio) {
  const sheet = getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.SESIONES);
  const located = findSessionRowById_(sheet, idSesion);
  const headers = getSheetHeaders_(sheet);
  const oldRow = rowArrayToObj_(headers, sheet.getRange(located.rowIndex, 1, 1, headers.length).getValues()[0]);
  const updated = { ...oldRow };
  Object.keys(patch || {}).forEach((k) => { updated[k] = k === 'estado' ? normalizeEstado_(patch[k]) : patch[k]; });
  updated.ultimaModificacion = formatDateTime_(new Date()); updated.modificadoPor = getUserEmail_();
  sheet.getRange(located.rowIndex, 1, 1, headers.length).setValues([headers.map((h) => updated[h])]);
  logCambiosBatch_(normalizeText_(idSesion), oldRow, updated, motivoCambio || 'Actualización de sesión', updated.modificadoPor);
  return { ok: true, diagnostics: located.diagnostics };
}

function cambiarEstado_(idSesion, nuevoEstado, motivo) {
  const sheet = getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.SESIONES);
  const located = findSessionRowById_(sheet, idSesion);
  const headers = getSheetHeaders_(sheet);
  const current = rowArrayToObj_(headers, sheet.getRange(located.rowIndex, 1, 1, headers.length).getValues()[0]);
  const normalizedTarget = normalizeEstado_(nuevoEstado);
  if (normalizeEstado_(current.estado || 'Programada') === normalizedTarget) return { ok: true, unchanged: true, diagnostics: located.diagnostics };
  return updateSessionById_(idSesion, { estado: normalizedTarget }, motivo);
}

function marcarRealizada(idSesion) { return cambiarEstado_(idSesion, 'Realizada', 'Marcada como realizada'); }
function cancelarSesion(idSesion) { return cambiarEstado_(idSesion, 'Cancelada', 'Cancelación de sesión'); }

function debugDuplicateSessions() {
  const data = getSesionesFromSheet_(getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.SESIONES));
  const result = { ok: true, totalRows: data.sesiones.length + data.incidencias.length, totalUniqueSessions: data.sesiones.length, duplicatedIds: data.duplicates };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function reconcileDuplicateSessionsDryRun() {
  const debug = debugDuplicateSessions();
  return { ok: true, totalDuplicatedIds: debug.duplicatedIds.length, proposals: debug.duplicatedIds.map((d) => ({ idSesion: d.idSesion, canonicalRow: d.canonicalRow, nonCanonicalRows: d.rows.filter((r) => r !== d.canonicalRow), reasonCanonical: d.reasonCanonical, proposedAction: 'No aplicar cambios automáticos. Revisar manualmente filas no canónicas.' })) };
}
