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

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sesiones Clínicas - Servicio de Farmacia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialData() {
  const diagnostics = { timestamp: new Date().toISOString(), sheetNamesFound: [], headers: {}, errors: [] };
  try {
    const ss = getSpreadsheet_();
    diagnostics.sheetNamesFound = ss.getSheets().map((s) => s.getName());

    const profesionalesSheet = getSheetOrThrow_(ss, SHEET_NAMES.PROFESIONALES);
    const tiposSheet = getSheetOrThrow_(ss, SHEET_NAMES.TIPOS);
    const sesionesSheet = getSheetOrThrow_(ss, SHEET_NAMES.SESIONES);

    diagnostics.headers[SHEET_NAMES.PROFESIONALES] = getSheetHeaders_(profesionalesSheet);
    diagnostics.headers[SHEET_NAMES.TIPOS] = getSheetHeaders_(tiposSheet);
    diagnostics.headers[SHEET_NAMES.SESIONES] = getSheetHeaders_(sesionesSheet);

    const profesionales = getProfesionalesActivosFromSheet_(profesionalesSheet);
    const tiposSesion = getTiposActivosFromSheet_(tiposSheet);
    const sesionesResult = getSesionesFromSheet_(sesionesSheet);

    diagnostics.activeProfesionales = profesionales.length;
    diagnostics.activeTiposSesion = tiposSesion.length;
    diagnostics.estados = ESTADOS_SESION.slice();
    diagnostics.incidenciasSesiones = sesionesResult.incidencias.length;

    return {
      ok: true,
      profesionales,
      tiposSesion,
      estados: ESTADOS_SESION.slice(),
      sesiones: sesionesResult.sesiones,
      incidenciasSesiones: sesionesResult.incidencias,
      message: 'Datos iniciales cargados correctamente.',
      diagnostics
    };
  } catch (err) {
    const errMsg = err && err.message ? err.message : String(err);
    diagnostics.errors.push(errMsg);
    console.error('[getInitialData] error', err);
    Logger.log(`Error en getInitialData: ${errMsg}`);
    return {
      ok: false,
      profesionales: [],
      tiposSesion: [],
      estados: ESTADOS_SESION.slice(),
      sesiones: [],
      incidenciasSesiones: [],
      message: "No se pudieron cargar los datos maestros. Revise que existan las pestañas 'Profesionales' y 'TiposSesion', que sus encabezados sean correctos y que haya filas activas marcadas como 'Sí'.",
      diagnostics
    };
  }
}

function debugInitialData() {
  const data = getInitialData();
  Logger.log(JSON.stringify({
    sheetNamesFound: (data.diagnostics && data.diagnostics.sheetNamesFound) || [],
    headers: (data.diagnostics && data.diagnostics.headers) || {},
    activeProfesionales: (data.diagnostics && data.diagnostics.activeProfesionales) || 0,
    activeTiposSesion: (data.diagnostics && data.diagnostics.activeTiposSesion) || 0,
    estados: data.estados || [],
    payload: data
  }, null, 2));
  return data;
}

function getSpreadsheet_() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No se pudo abrir el Spreadsheet activo. Vincule el proyecto a una hoja o configure SPREADSHEET_ID.');
  return ss;
}

function getSheetOrThrow_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`No existe la pestaña ${sheetName}.`);
  return sheet;
}

function getSheetHeaders_(sheet) {
  if (sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((v) => String(v || ''));
}

function getRowsAsObjects_(sheet, expectedHeaders) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const headerMap = getHeaderMap_(headers);
  const missing = expectedHeaders.filter((h) => headerMap[h] === undefined);
  if (missing.length) throw new Error(`Encabezados inválidos en ${sheet.getName()}. Faltan: ${missing.join(', ')}`);

  return values.slice(1)
    .filter((row) => row.some((c) => normalizeText_(c) !== ''))
    .map((row) => {
      const obj = {};
      expectedHeaders.forEach((h) => {
        obj[h] = row[headerMap[h]];
      });
      return obj;
    });
}

function getHeaderMap_(headers) {
  const normalizedToIndex = {};
  headers.forEach((header, index) => {
    const norm = normalizeHeader_(header);
    if (norm && normalizedToIndex[norm] === undefined) normalizedToIndex[norm] = index;
  });

  const map = {};
  const allCanonical = [].concat(HEADERS.SESIONES, HEADERS.PROFESIONALES, HEADERS.TIPOS, HEADERS.HISTORIAL);
  allCanonical.forEach((canonical) => {
    const normCanonical = normalizeHeader_(canonical);
    if (normalizedToIndex[normCanonical] !== undefined) map[canonical] = normalizedToIndex[normCanonical];
  });
  return map;
}

function normalizeHeader_(value) {
  return normalizeText_(value)
    .replace(/\uFEFF/g, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\uFEFF/g, '').replace(/[\u200B-\u200D\u2060]/g, '').trim();
}

function isActive_(value) {
  if (value === true) return true;
  const txt = normalizeText_(value).toLowerCase();
  return ['sí', 'si', 'true', '1', 'activo'].includes(txt);
}

function getProfesionalesActivosFromSheet_(sheet) {
  return getRowsAsObjects_(sheet, HEADERS.PROFESIONALES)
    .filter((row) => isActive_(row.activo) && normalizeText_(row.nombre))
    .map((row) => ({ nombre: normalizeText_(row.nombre), area: normalizeText_(row.area), email: normalizeText_(row.email) }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));
}

function getTiposActivosFromSheet_(sheet) {
  return getRowsAsObjects_(sheet, HEADERS.TIPOS)
    .filter((row) => isActive_(row.activo) && normalizeText_(row.tipoSesion))
    .map((row) => normalizeText_(row.tipoSesion))
    .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
}

function getSesionesFromSheet_(sheet) {
  const rows = getRowsAsObjects_(sheet, HEADERS.SESIONES);
  const sesiones = [];
  const incidencias = [];
  rows.forEach((row, idx) => {
    try {
      const fecha = normalizeDate_(row.fecha);
      const hora = normalizeTime_(row.hora);
      sesiones.push({
        idSesion: normalizeText_(row.idSesion),
        fecha,
        hora,
        duracionMin: Number(row.duracionMin) || 0,
        tipoSesion: normalizeText_(row.tipoSesion),
        tema: normalizeText_(row.tema),
        ponente: normalizeText_(row.ponente),
        area: normalizeText_(row.area),
        estado: normalizeText_(row.estado),
        sustituto: normalizeText_(row.sustituto),
        enlaceMaterial: normalizeText_(row.enlaceMaterial),
        observaciones: normalizeText_(row.observaciones),
        fechaCreacion: normalizeText_(row.fechaCreacion),
        ultimaModificacion: normalizeText_(row.ultimaModificacion),
        modificadoPor: normalizeText_(row.modificadoPor)
      });
    } catch (err) {
      incidencias.push(`Fila ${idx + 2}: ${err.message || err}`);
    }
  });
  sesiones.sort((a, b) => `${a.fecha} ${a.hora}`.localeCompare(`${b.fecha} ${b.hora}`));
  return { sesiones, incidencias };
}

function normalizeTime_(value) {
  if (value === null || value === undefined || normalizeText_(value) === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60);
    const dayMinutes = 24 * 60;
    const m = ((totalMinutes % dayMinutes) + dayMinutes) % dayMinutes;
    return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
  }
  const match = normalizeText_(value).match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if (!match) throw new Error('Hora inválida. Use HH:mm.');
  const hh = Number(match[1]);
  const mm = match[2] === undefined ? 0 : Number(match[2]);
  const ss = match[3] === undefined ? 0 : Number(match[3]);
  if (![hh, mm, ss].every(Number.isInteger) || hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) {
    throw new Error('Hora inválida. Use HH:mm con valores entre 00:00 y 23:59.');
  }
  return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

// Alias conservado para compatibilidad interna
function normalizeHora_(value) { return normalizeTime_(value); }

function normalizeDate_(value) {
  if (!value) throw new Error('Fecha inválida.');
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  const d = new Date(normalizeText_(value));
  if (isNaN(d.getTime())) throw new Error('Fecha inválida. Debe tener formato yyyy-MM-dd.');
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function createSesion(payload) {
  validateSesion_(payload);
  const ss = getSpreadsheet_();
  const sheet = getSheetOrThrow_(ss, SHEET_NAMES.SESIONES);
  const now = new Date();
  const user = getUserEmail_();
  const fecha = normalizeDate_(payload.fecha);
  const idSesion = generateIdSesion_(fecha);
  const area = resolveArea_(payload.ponente, payload.area);

  const row = {
    idSesion,
    fecha,
    hora: normalizeTime_(payload.hora),
    duracionMin: Number(payload.duracionMin),
    tipoSesion: normalizeText_(payload.tipoSesion),
    tema: normalizeText_(payload.tema),
    ponente: normalizeText_(payload.ponente),
    area,
    estado: normalizeText_(payload.estado),
    sustituto: normalizeText_(payload.sustituto),
    enlaceMaterial: normalizeText_(payload.enlaceMaterial),
    observaciones: normalizeText_(payload.observaciones),
    fechaCreacion: formatDateTime_(now),
    ultimaModificacion: formatDateTime_(now),
    modificadoPor: user
  };

  sheet.appendRow(HEADERS.SESIONES.map((h) => row[h]));
  logCambiosBatch_(idSesion, null, row, 'Creación de sesión', user);
  return { ok: true, idSesion };
}

function updateSesion(payload) {
  validateSesion_(payload, true);
  if (!payload.idSesion) throw new Error('idSesion es obligatorio para editar.');
  const ss = getSpreadsheet_();
  const sheet = getSheetOrThrow_(ss, SHEET_NAMES.SESIONES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = findHeaderIndexes_(headers);
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.idSesion]) === String(payload.idSesion)) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) throw new Error('No se encontró la sesión a editar.');
  const oldRowArray = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const oldRow = rowArrayToObj_(headers, oldRowArray);
  const user = getUserEmail_();
  const updated = {
    ...oldRow,
    fecha: normalizeDate_(payload.fecha),
    hora: normalizeTime_(payload.hora),
    duracionMin: Number(payload.duracionMin),
    tipoSesion: normalizeText_(payload.tipoSesion),
    tema: normalizeText_(payload.tema),
    ponente: normalizeText_(payload.ponente),
    area: resolveArea_(payload.ponente, payload.area),
    estado: normalizeText_(payload.estado),
    sustituto: normalizeText_(payload.sustituto),
    enlaceMaterial: normalizeText_(payload.enlaceMaterial),
    observaciones: normalizeText_(payload.observaciones),
    ultimaModificacion: formatDateTime_(new Date()),
    modificadoPor: user
  };
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map((h) => updated[h])]);
  logCambiosBatch_(payload.idSesion, oldRow, updated, 'Edición de sesión', user);
  return { ok: true };
}
function marcarRealizada(idSesion) { return cambiarEstado_(idSesion, 'Realizada', 'Marcada como realizada'); }
function cancelarSesion(idSesion) { return cambiarEstado_(idSesion, 'Cancelada', 'Cancelación de sesión'); }

function cambiarEstado_(idSesion, nuevoEstado, motivo) {
  if (!idSesion) throw new Error('idSesion es obligatorio.');
  const ss = getSpreadsheet_();
  const sheet = getSheetOrThrow_(ss, SHEET_NAMES.SESIONES);
  const values = sheet.getDataRange().getValues();
  const idx = findHeaderIndexes_(values[0]);
  const user = getUserEmail_();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.idSesion]) === String(idSesion)) {
      const rowIndex = i + 1;
      const oldEstado = values[i][idx.estado];
      sheet.getRange(rowIndex, idx.estado + 1).setValue(nuevoEstado);
      sheet.getRange(rowIndex, idx.ultimaModificacion + 1).setValue(formatDateTime_(new Date()));
      sheet.getRange(rowIndex, idx.modificadoPor + 1).setValue(user);
      logCambio_(idSesion, 'estado', oldEstado, nuevoEstado, user, motivo);
      return { ok: true };
    }
  }
  throw new Error('No se encontró la sesión.');
}

function validateSesion_(payload, isEdit) {
  const required = ['fecha', 'hora', 'tipoSesion', 'tema', 'ponente', 'estado'];
  required.forEach((f) => { if (!normalizeText_(payload[f])) throw new Error(`El campo ${f} es obligatorio.`); });
  const dur = Number(payload.duracionMin);
  if (!Number.isFinite(dur) || dur <= 0) throw new Error('duracionMin debe ser numérico y mayor que 0.');
  normalizeDate_(payload.fecha);
  normalizeTime_(payload.hora);
  if (isEdit && !payload.idSesion) throw new Error('idSesion es obligatorio.');
}

function getUserEmail_() { return Session.getActiveUser().getEmail() || 'usuario_no_identificado'; }
function cleanText_(v) { return normalizeText_(v); }
function isActivo_(v) { return isActive_(v); }

function rowArrayToObj_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function findHeaderIndexes_(headers) {
  const map = getHeaderMap_(headers);
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  Object.keys(map).forEach((k) => { idx[k] = map[k]; });
  return idx;
}

function resolveArea_(ponente, areaInput) {
  const areaManual = normalizeText_(areaInput);
  if (areaManual) return areaManual;
  const ss = getSpreadsheet_();
  const profesionales = getProfesionalesActivosFromSheet_(getSheetOrThrow_(ss, SHEET_NAMES.PROFESIONALES));
  const p = profesionales.find((pro) => normalizeText_(pro.nombre).toLowerCase() === normalizeText_(ponente).toLowerCase());
  return p ? normalizeText_(p.area) : '';
}

function formatDateTime_(dateObj) { return Utilities.formatDate(dateObj, TZ, 'yyyy-MM-dd HH:mm:ss'); }

function logCambiosBatch_(idSesion, oldObj, newObj, motivo, user) {
  const camposAuditables = ['fecha', 'hora', 'duracionMin', 'tipoSesion', 'tema', 'ponente', 'area', 'estado', 'sustituto', 'enlaceMaterial', 'observaciones'];
  camposAuditables.forEach((campo) => {
    const anterior = oldObj ? normalizeText_(oldObj[campo]) : '';
    const nuevo = normalizeText_(newObj[campo]);
    if (!oldObj || anterior !== nuevo) logCambio_(idSesion, campo, anterior, nuevo, user, motivo);
  });
}

function logCambio_(idSesion, campo, valorAnterior, valorNuevo, usuario, motivoCambio) {
  const ss = getSpreadsheet_();
  const sheet = getSheetOrThrow_(ss, SHEET_NAMES.HISTORIAL);
  sheet.appendRow([formatDateTime_(new Date()), idSesion, campo, normalizeText_(valorAnterior), normalizeText_(valorNuevo), usuario, motivoCambio]);
}

function generateIdSesion_(fechaStr) {
  const year = String(fechaStr).slice(0, 4);
  const ss = getSpreadsheet_();
  const sesiones = getSesionesFromSheet_(getSheetOrThrow_(ss, SHEET_NAMES.SESIONES)).sesiones;
  const sameYear = sesiones.filter((s) => String(s.idSesion).startsWith(`SES-${year}-`));
  const maxSeq = sameYear.reduce((max, s) => Math.max(max, Number(String(s.idSesion).split('-')[2]) || 0), 0);
  return `SES-${year}-${String(maxSeq + 1).padStart(3, '0')}`;
}
