const SHEET_NAMES = {
  SESIONES: 'Sesiones',
  PROFESIONALES: 'Profesionales',
  TIPOS: 'TiposSesion',
  HISTORIAL: 'HistorialCambios'
};

const HEADERS = {
  SESIONES: ['idSesion','fecha','hora','duracionMin','tipoSesion','tema','ponente','area','estado','sustituto','enlaceMaterial','observaciones','fechaCreacion','ultimaModificacion','modificadoPor'],
  PROFESIONALES: ['nombre','area','email','activo','observaciones'],
  TIPOS: ['tipoSesion','activo'],
  HISTORIAL: ['fechaCambio','idSesion','campoModificado','valorAnterior','valorNuevo','usuario','motivoCambio']
};

const TZ = 'Europe/Madrid';

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
  const sesionesResult = getSesiones_();
  const profesionales = getProfesionalesActivos_();
  const tiposSesion = getTiposActivos_();

  return {
    sesiones: sesionesResult.sesiones,
    incidenciasSesiones: sesionesResult.incidencias,
    profesionales,
    tiposSesion,
    estados: ['Programada', 'Confirmada', 'Pendiente', 'Cambiada', 'Cancelada', 'Realizada']
  };
}

function createSesion(payload) {
  validateSesion_(payload);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SESIONES);
  const now = new Date();
  const user = getUserEmail_();

  const fecha = normalizeDate_(payload.fecha);
  const idSesion = generateIdSesion_(fecha);
  const area = resolveArea_(payload.ponente, payload.area);

  const row = {
    idSesion,
    fecha,
    hora: normalizeHora_(payload.hora),
    duracionMin: Number(payload.duracionMin),
    tipoSesion: cleanText_(payload.tipoSesion),
    tema: cleanText_(payload.tema),
    ponente: cleanText_(payload.ponente),
    area,
    estado: cleanText_(payload.estado),
    sustituto: cleanText_(payload.sustituto),
    enlaceMaterial: cleanText_(payload.enlaceMaterial),
    observaciones: cleanText_(payload.observaciones),
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SESIONES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = findHeaderIndexes_(headers);

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.idSesion]) === String(payload.idSesion)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('No se encontró la sesión a editar.');

  const oldRowArray = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const oldRow = rowArrayToObj_(headers, oldRowArray);
  const user = getUserEmail_();
  const now = new Date();

  const updated = {
    ...oldRow,
    fecha: normalizeDate_(payload.fecha),
    hora: normalizeHora_(payload.hora),
    duracionMin: Number(payload.duracionMin),
    tipoSesion: cleanText_(payload.tipoSesion),
    tema: cleanText_(payload.tema),
    ponente: cleanText_(payload.ponente),
    area: resolveArea_(payload.ponente, payload.area),
    estado: cleanText_(payload.estado),
    sustituto: cleanText_(payload.sustituto),
    enlaceMaterial: cleanText_(payload.enlaceMaterial),
    observaciones: cleanText_(payload.observaciones),
    ultimaModificacion: formatDateTime_(now),
    modificadoPor: user
  };

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map((h) => updated[h])]);

  logCambiosBatch_(payload.idSesion, oldRow, updated, 'Edición de sesión', user);

  return { ok: true };
}

function marcarRealizada(idSesion) {
  return cambiarEstado_(idSesion, 'Realizada', 'Marcada como realizada');
}

function cancelarSesion(idSesion) {
  return cambiarEstado_(idSesion, 'Cancelada', 'Cancelación de sesión');
}

function cambiarEstado_(idSesion, nuevoEstado, motivo) {
  if (!idSesion) throw new Error('idSesion es obligatorio.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SESIONES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = findHeaderIndexes_(headers);
  const user = getUserEmail_();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.idSesion]) === String(idSesion)) {
      const rowIndex = i + 1;
      const oldEstado = data[i][idx.estado];
      if (String(oldEstado) === String(nuevoEstado)) return { ok: true };

      sheet.getRange(rowIndex, idx.estado + 1).setValue(nuevoEstado);
      sheet.getRange(rowIndex, idx.ultimaModificacion + 1).setValue(formatDateTime_(now));
      sheet.getRange(rowIndex, idx.modificadoPor + 1).setValue(user);

      logCambio_(idSesion, 'estado', oldEstado, nuevoEstado, user, motivo === 'Cancelación de sesión' ? 'Cancelación de sesión' : 'Cambio de estado');
      if (motivo === 'Marcada como realizada') {
        logCambio_(idSesion, 'estado', oldEstado, nuevoEstado, user, 'Marcada como realizada');
      }
      return { ok: true };
    }
  }

  throw new Error('No se encontró la sesión.');
}

function getSesiones_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SESIONES);
  if (!sheet) throw new Error('No existe la pestaña Sesiones.');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { sesiones: [], incidencias: [] };
  const headers = values[0];
  const sesiones = [];
  const incidencias = [];

  values.slice(1).forEach((row, idx) => {
    if (!row.some((c) => String(c).trim() !== '')) return;
    try {
      const s = rowArrayToObj_(headers, row);
      sesiones.push({
        ...s,
        fecha: normalizeDate_(s.fecha),
        hora: normalizeHora_(s.hora),
        duracionMin: Number(s.duracionMin) || 0
      });
    } catch (err) {
      const rowNumber = idx + 2;
      const msg = `Fila ${rowNumber}: ${err && err.message ? err.message : 'Dato inválido'}`;
      incidencias.push(msg);
      Logger.log(`Sesión omitida por validación: ${msg}`);
    }
  });

  sesiones.sort((a, b) => (a.fecha + ' ' + a.hora).localeCompare(b.fecha + ' ' + b.hora));
  return { sesiones, incidencias };
}

function getProfesionalesActivos_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PROFESIONALES);
  if (!sheet) throw new Error('No existe la pestaña Profesionales.');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1)
    .filter((row) => row.some((c) => String(c).trim() !== ''))
    .map((row) => rowArrayToObj_(headers, row))
    .filter((p) => isActivo_(p.activo));
}

function getTiposActivos_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TIPOS);
  if (!sheet) throw new Error('No existe la pestaña TiposSesion.');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1)
    .filter((row) => row.some((c) => String(c).trim() !== ''))
    .map((row) => rowArrayToObj_(headers, row))
    .filter((t) => isActivo_(t.activo))
    .map((t) => t.tipoSesion);
}

function generateIdSesion_(fechaStr) {
  const year = String(fechaStr).slice(0, 4);
  const sesiones = getSesiones_().sesiones;
  const sameYear = sesiones.filter((s) => String(s.idSesion).startsWith(`SES-${year}-`));
  const maxSeq = sameYear.reduce((max, s) => {
    const parts = String(s.idSesion).split('-');
    const n = Number(parts[2]);
    return Number.isFinite(n) ? Math.max(max, n) : max;
  }, 0);
  const next = String(maxSeq + 1).padStart(3, '0');
  return `SES-${year}-${next}`;
}

function validateSesion_(payload, isEdit) {
  const required = ['fecha', 'hora', 'tipoSesion', 'tema', 'ponente', 'estado'];
  required.forEach((f) => {
    if (!cleanText_(payload[f])) throw new Error(`El campo ${f} es obligatorio.`);
  });

  const dur = Number(payload.duracionMin);
  if (!Number.isFinite(dur) || dur <= 0) {
    throw new Error('duracionMin debe ser numérico y mayor que 0.');
  }

  normalizeDate_(payload.fecha);
  normalizeHora_(payload.hora);

  if (isEdit && !payload.idSesion) throw new Error('idSesion es obligatorio.');
}

function normalizeDate_(value) {
  if (!value) throw new Error('Fecha inválida.');
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const txt = String(value).trim();
  const d = new Date(txt);
  if (isNaN(d.getTime())) throw new Error('Fecha inválida. Debe tener formato yyyy-MM-dd.');
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function normalizeHora_(value) {
  if (value === null || value === undefined || cleanText_(value) === '') {
    throw new Error('Hora inválida. Debe tener formato HH:mm.');
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60);
    const dayMinutes = 24 * 60;
    const normalizedMinutes = ((totalMinutes % dayMinutes) + dayMinutes) % dayMinutes;
    const hh = String(Math.floor(normalizedMinutes / 60)).padStart(2, '0');
    const mm = String(normalizedMinutes % 60).padStart(2, '0');
    return `${hh}:${mm}`;
  }

  const txt = String(value).trim();
  const m = txt.match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if (!m) {
    throw new Error('Hora inválida. Use HH:mm (ej: 08:15).');
  }

  const hh = Number(m[1]);
  const mm = m[2] === undefined ? 0 : Number(m[2]);
  const ss = m[3] === undefined ? 0 : Number(m[3]);

  if (![hh, mm, ss].every(Number.isInteger) || hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) {
    throw new Error('Hora inválida. Use HH:mm con valores entre 00:00 y 23:59.');
  }

  return `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

function getUserEmail_() {
  const email = Session.getActiveUser().getEmail();
  return email ? email : 'usuario_no_identificado';
}

function cleanText_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function isActivo_(v) {
  const normalized = String(v === true ? 'true' : v || '').trim().toLowerCase();
  return ['sí', 'si', 'true'].includes(normalized);
}

function rowArrayToObj_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function findHeaderIndexes_(headers) {
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  return idx;
}

function resolveArea_(ponente, areaInput) {
  const areaManual = cleanText_(areaInput);
  if (areaManual) return areaManual;

  const profesionales = getProfesionalesActivos_();
  const p = profesionales.find((pro) => cleanText_(pro.nombre).toLowerCase() === cleanText_(ponente).toLowerCase());
  return p ? cleanText_(p.area) : '';
}

function formatDateTime_(dateObj) {
  return Utilities.formatDate(dateObj, TZ, 'yyyy-MM-dd HH:mm:ss');
}

function logCambiosBatch_(idSesion, oldObj, newObj, motivo, user) {
  const camposAuditables = ['fecha','hora','duracionMin','tipoSesion','tema','ponente','area','estado','sustituto','enlaceMaterial','observaciones'];

  camposAuditables.forEach((campo) => {
    const anterior = oldObj ? cleanText_(oldObj[campo]) : '';
    const nuevo = cleanText_(newObj[campo]);
    if (!oldObj || anterior !== nuevo) {
      logCambio_(idSesion, campo, anterior, nuevo, user, motivo);
    }
  });
}

function logCambio_(idSesion, campo, valorAnterior, valorNuevo, usuario, motivoCambio) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HISTORIAL);
  if (!sheet) throw new Error('No existe la pestaña HistorialCambios.');

  sheet.appendRow([
    formatDateTime_(new Date()),
    idSesion,
    campo,
    cleanText_(valorAnterior),
    cleanText_(valorNuevo),
    usuario,
    motivoCambio
  ]);
}
