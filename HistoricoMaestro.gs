/**
 * ============================================================================
 * HISTORICO MAESTRO PARA LOOKER STUDIO  -  Apps Script (V8)
 * ============================================================================
 *
 * Que hace
 *   Lee los ficheros de periodo de una carpeta de Drive (una hoja por periodo),
 *   y archiva las pestanas de datos que le indiques en un fichero maestro con
 *   ID fijo, una pestana destino por cada entidad (nunca una pestana por periodo).
 *
 * Modelo del maestro (por cada entidad):
 *   fila 1:  periodo | anio | semana | fecha_reporte | <columnas del origen> | fecha_carga
 *   fila 2+: una fila por registro, con el periodo del fichero del que viene.
 *
 * Garantias
 *   - Idempotente: reejecutar no duplica. Si el periodo ya existe, se borra y se
 *     reescribe (upsert).
 *   - Cabeceras mapeadas POR NOMBRE, nunca por posicion.
 *   - Columnas nuevas en el origen: se anaden solas al maestro.
 *   - Columnas que faltan: si faltan en el fichero MAS RECIENTE se considera
 *     renombrado/borrado y aborta. Si faltan en un fichero antiguo (la columna
 *     se creo despues) se rellenan vacias y se avisa en el resumen.
 *   - Retencion por pestana: conserva los ultimos CONFIG.windowPeriods periodos.
 *   - La fila 1 no se toca jamas: el data source de Looker no se rompe.
 *   - Solo hojas de Google: si en la carpeta aparece otro tipo de fichero con
 *     nombre de periodo, se rechaza con error explicito.
 *
 * Puesta en marcha
 *   1. Rellena CONFIG (IDs, correo y entidades).
 *   2. Ejecuta dryRun() y revisa el log: ficheros detectados y estado del maestro.
 *   3. Ejecuta ingestManual() con un periodo concreto para la primera carga.
 *   4. Ejecuta installTrigger() para dejarlo automatico.
 *   5. Comparte el maestro SOLO EN LECTURA. Nadie debe editar ni formatear ese
 *      fichero: aplicar formato a columnas enteras lo devuelve al problema de
 *      memoria del navegador.
 *
 * Notas de mantenimiento
 *   - El maestro debe vivir en Mi unidad, no en una unidad compartida: el
 *     conector de Google Sheets de Looker Studio no accede a unidades compartidas.
 *   - El scope de Drive es de solo lectura: el script nunca modifica ni borra
 *     nada fuera del fichero maestro.
 *   - Reconstruir muchos periodos de golpe puede superar los 6 minutos de
 *     ejecucion: hazlo por tandas con ingestPeriod('AAAA-Wnn').
 *   - Los numeros con formato local ("1.234,56") viajan como texto. Si Looker
 *     los necesita como numero, arreglalo en el origen.
 * ============================================================================
 */

// ============================================================================
// CONFIGURACION  -  lo unico que tienes que editar
// ============================================================================

var CONFIG = {

  // Fichero maestro (el ID aparece en la URL: /spreadsheets/d/<ID>/edit)
  masterFileId: 'PON_AQUI_EL_ID_DEL_MAESTRO',

  // Carpeta de Drive con los ficheros de periodo. No es recursiva.
  sourceFolderId: 'PON_AQUI_EL_ID_DE_LA_CARPETA',

  // Avisos de error
  notifyEmail: 'tu.correo@empresa.com',

  timezone: 'Europe/Madrid',

  // Periodos que se conservan por pestana destino. 6 a 12 segun tu caso.
  windowPeriods: 12,

  // 'ISO_WEEK' -> nombres tipo 2026-W47   |   'MONTH' -> nombres tipo 2026-11
  periodFormat: 'ISO_WEEK',

  // Opcional. Si lo dejas vacio se deriva de periodFormat.
  // Ejemplos: '^(\\d{4})-W(\\d{2})$'  o  '^(\\d{4})-(\\d{2})$'
  filenamePattern: '',

  // Fila donde estan las cabeceras en las pestanas de origen
  headerRow: 1,

  // Aborta si el origen trae celdas con error (#REF!, #N/A, ...)
  failOnErrorCells: true,

  // Una entrada por cada pestana de datos que quieras archivar.
  // destino : nombre de la pestana en el maestro (se crea si no existe)
  // origen  : nombre de la pestana en el fichero de periodo
  // headerRow: opcional, sobrescribe CONFIG.headerRow para esta entidad
  entities: [
    { destino: 'entidad_a', origen: 'Hoja 1' },
    { destino: 'entidad_b', origen: 'Hoja 2' },
    { destino: 'entidad_c', origen: 'Hoja 3' }
  ]
};

// Periodo que usa ingestManual(). Cambialo y pulsa Ejecutar.
var MANUAL_PERIOD = '2026-W47';

// ============================================================================
// CONSTANTES INTERNAS
// ============================================================================

var META_HEADERS = ['periodo', 'anio', 'semana', 'fecha_reporte'];
var CACHE_HEADER = 'fecha_carga';
var LOCK_WAIT_MS = 30000;
var SHEETS_MIME = 'application/vnd.google-apps.spreadsheet';

// ============================================================================
// PUNTOS DE ENTRADA
// ============================================================================

/** Trigger automatico: procesa solo los periodos que falten. */
function ingestScheduled() {
  execute_(function () {
    return ingest_({ onlyMissing: true });
  });
}

/** Carga inicial o reconstruccion: reprocesa todos los ficheros de la carpeta. */
function rebuildAll() {
  execute_(function () {
    return ingest_({ onlyMissing: false });
  });
}

/** Rehace un unico periodo (util tras corregir un fichero ya cargado). */
function ingestManual() {
  ingestPeriod(MANUAL_PERIOD);
}

/** Rehace el periodo indicado, aunque ya estuviera cargado. */
function ingestPeriod(periodo) {
  execute_(function () {
    return ingest_({ onlyMissing: false, onlyPeriod: String(periodo) });
  });
}

/**
 * Diagnostico sin escribir nada: ficheros detectados, periodos y estado del
 * maestro. Es lo primero que debes ejecutar.
 */
function dryRun() {
  var master = SpreadsheetApp.openById(CONFIG.masterFileId);
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var files = listSourceFiles_(folder);
  var lines = [];

  lines.push('Formato de periodo: ' + CONFIG.periodFormat + ' | ventana: ' + CONFIG.windowPeriods);
  lines.push('Ficheros detectados en la carpeta: ' + files.length);
  files.forEach(function (f) {
    lines.push('  - ' + f.name + '  ->  ' + f.periodo + '  (fecha_reporte ' + formatDate_(fechaReporte_(f.periodo)) + ')');
  });
  lines.push('');
  lines.push('Estado del maestro:');
  CONFIG.entities.forEach(function (e) {
    var sheet = master.getSheetByName(e.destino);
    if (!sheet) {
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): no existe, se creara');
      return;
    }
    try {
      var periods = distinct_(readPeriodColumn_(sheet));
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): ' + Math.max(sheet.getLastRow() - 1, 0) +
        ' filas, ' + periods.length + ' periodos' + (periods.length ? ' [' + periods.join(', ') + ']' : ''));
    } catch (err) {
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): PROBLEMA -> ' + err.message);
    }
  });

  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/** Deja el trigger automatico (lunes 06:00 o dia 1 06:00 segun periodFormat). */
function installTrigger() {
  removeTriggers();
  var builder = ScriptApp.newTrigger('ingestScheduled').timeBased().inTimezone(CONFIG.timezone);
  if (CONFIG.periodFormat === 'MONTH') {
    builder.onMonthDay(1).atHour(6).create();
  } else {
    builder.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  }
  Logger.log('Trigger instalado.');
}

/** Quita el trigger automatico. */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'ingestScheduled') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ============================================================================
// MOTOR
// ============================================================================

/**
 * Envoltorio comun: lock de ejecucion, aviso por correo ante error y relanzado
 * para que la ejecucion quede marcada como fallida.
 */
function execute_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    throw new Error('Ya hay una ejecucion en curso del historico. Abortado para no escribir dos veces.');
  }
  try {
    var summary = fn();
    Logger.log(JSON.stringify(summary, null, 2));
    return summary;
  } catch (err) {
    notify_('ERROR', (err && err.stack) ? err.stack : String(err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function ingest_(opts) {
  opts = opts || {};
  var master = SpreadsheetApp.openById(CONFIG.masterFileId);
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var allFiles = listSourceFiles_(folder);

  if (!allFiles.length) {
    throw new Error('No hay ficheros en la carpeta que cumplan el patron de nombre. Patron usado: ' + pattern_());
  }

  // El periodo mas reciente de la carpeta define el esquema "vigente": solo ese
  // fichero puede provocar un aborto por columna desaparecida.
  var latestPeriod = allFiles[allFiles.length - 1].periodo;

  var files = allFiles;
  if (opts.onlyPeriod) {
    files = allFiles.filter(function (f) { return f.periodo === opts.onlyPeriod; });
    if (!files.length) {
      throw new Error('No hay ningun fichero para el periodo ' + opts.onlyPeriod + '. Revisa el nombre del fichero y CONFIG.filenamePattern.');
    }
  }

  // Cache de periodos ya presentes por entidad, para no leer la columna N veces.
  var periodCache = {};
  CONFIG.entities.forEach(function (e) {
    periodCache[e.destino] = {};
    readPeriodColumn_(master.getSheetByName(e.destino)).forEach(function (p) {
      periodCache[e.destino][p] = true;
    });
  });

  var summary = { ok: [], skipped: [], failed: [], warned: [] };

  files.forEach(function (entry) {
    var pending = CONFIG.entities.filter(function (e) {
      return !(opts.onlyMissing && periodCache[e.destino][entry.periodo]);
    });
    if (!pending.length) {
      summary.skipped.push(entry.periodo);
      return;
    }

    var srcSs;
    try {
      srcSs = openSource_(entry.id);
    } catch (err) {
      summary.failed.push(entry.periodo + ' / abrir fichero: ' + err.message);
      return;
    }

    pending.forEach(function (e) {
      try {
        ingestEntity_(master, srcSs, e, entry, entry.periodo === latestPeriod, summary.warned);
        periodCache[e.destino][entry.periodo] = true;
        summary.ok.push(e.destino + ' @ ' + entry.periodo);
      } catch (err) {
        summary.failed.push(entry.periodo + ' / ' + e.destino + ': ' + err.message);
      }
    });
  });

  // Retencion al final, una vez por entidad.
  CONFIG.entities.forEach(function (e) {
    try {
      applyRetention_(master, e);
    } catch (err) {
      summary.failed.push('retencion / ' + e.destino + ': ' + err.message);
    }
  });

  SpreadsheetApp.flush();

  if (summary.warned.length) {
    Logger.log('Avisos (' + summary.warned.length + '):\n' + summary.warned.join('\n'));
  }

  if (summary.failed.length) {
    throw new Error('Fallos (' + summary.failed.length + '):\n' + summary.failed.join('\n') +
      '\n\nCargados: ' + summary.ok.length + ' | Omitidos: ' + summary.skipped.length +
      ' | Avisos: ' + summary.warned.length);
  }
  return summary;
}

/**
 * Archiva una entidad de un fichero de periodo en su pestana del maestro.
 * isLatest: true si este fichero es el mas reciente de la carpeta; solo entonces
 * una columna ausente se trata como renombrado/borrado y aborta.
 */
function ingestEntity_(master, srcSs, entity, entry, isLatest, warnings) {
  var masterSheet = ensureMasterTab_(master, entity);
  var srcSheet = srcSs.getSheetByName(resolveSheetName_(srcSs, entity.origen));

  var headerRow = entity.headerRow || CONFIG.headerRow;
  var lastRow = srcSheet.getLastRow();
  var lastCol = srcSheet.getLastColumn();
  if (lastRow <= headerRow || lastCol === 0) {
    throw new Error('la pestana "' + entity.origen + '" no tiene datos bajo la fila ' + headerRow);
  }

  var block = srcSheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getValues();
  var srcHeader = block.shift().map(function (h) { return String(h).trim(); });
  assertHeader_(srcHeader, entity.origen);

  if (CONFIG.failOnErrorCells) {
    assertNoErrors_(block, entity.origen, headerRow);
  }

  var rows = block.filter(function (r) {
    return r.some(function (c) { return c !== '' && c !== null; });
  });
  if (!rows.length) {
    throw new Error('la pestana "' + entity.origen + '" no tiene filas con datos');
  }

  var masterHeader = ensureColumns_(masterSheet, srcHeader, isLatest, warnings, entity.destino, entry.periodo);
  var now = new Date();

  var payload = rows.map(function (r) {
    return masterHeader.map(function (col) {
      if (col === 'periodo') return entry.periodo;
      if (col === 'anio') return entry.anio;
      if (col === 'semana') return entry.semana;
      if (col === 'fecha_reporte') return entry.fechaReporte;
      if (col === CACHE_HEADER) return now;
      var i = srcHeader.indexOf(col);
      return i === -1 ? '' : r[i];
    });
  });

  // Upsert: fuera las filas previas de este periodo, luego append de una sola escritura.
  deletePeriodRows_(masterSheet, entry.periodo);
  var startRow = masterSheet.getLastRow() + 1;
  masterSheet.getRange(startRow, 1, payload.length, masterHeader.length).setValues(payload);
}

// ============================================================================
// ESTRUCTURA DEL MAESTRO
// ============================================================================

function ensureMasterTab_(master, entity) {
  var sheet = master.getSheetByName(entity.destino);
  if (!sheet) {
    sheet = master.insertSheet(entity.destino);
    Logger.log('Creada la pestana destino "' + entity.destino + '".');
  }

  if (sheet.getLastRow() === 0) {
    var header = META_HEADERS.concat([CACHE_HEADER]);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var current = readHeader_(sheet);
  for (var i = 0; i < META_HEADERS.length; i++) {
    if (current[i] !== META_HEADERS[i]) {
      throw new Error('la pestana destino "' + entity.destino + '" no empieza con la cabecera esperada: ' +
        'columna ' + (i + 1) + ' deberia ser "' + META_HEADERS[i] + '" y es "' + (current[i] || '') + '". ' +
        'Renombra el fichero maestro o corrige la cabecera a mano.');
    }
  }
  if (current[current.length - 1] !== CACHE_HEADER) {
    throw new Error('en la pestana destino "' + entity.destino + '" la ultima columna debe ser "' + CACHE_HEADER +
      '". Revisa la cabecera del maestro.');
  }
  return sheet;
}

function readHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
}

/**
 * Alinea las columnas del origen con las del maestro.
 *   - Aditivas: se anaden al maestro.
 *   - Ausentes en el fichero mas reciente: aborta (renombrado o borrado real).
 *   - Ausentes en un fichero antiguo: se rellenan vacias y se avisa.
 */
function ensureColumns_(sheet, srcHeader, isLatest, warnings, destino, periodo) {
  var header = readHeader_(sheet);
  var lastDataPos = header.length - 1; // 1-based: ultima columna de datos (justo antes de fecha_carga)
  var masterData = header.slice(META_HEADERS.length, header.length - 1);

  var missing = masterData.filter(function (c) { return srcHeader.indexOf(c) === -1; });
  if (missing.length) {
    if (isLatest) {
      throw new Error('columnas que existen en el maestro y no en el origen: ' + missing.join(', ') +
        '. Si es un renombrado, actualiza la cabecera del maestro a mano; si no, revisa el fichero de origen.');
    }
    if (warnings) {
      warnings.push(destino + ' @ ' + periodo + ': columnas ausentes en el origen rellenadas vacias -> ' + missing.join(', '));
    }
  }

  var newCols = srcHeader.filter(function (c) { return c !== '' && header.indexOf(c) === -1; });
  if (newCols.length) {
    sheet.insertColumnsAfter(lastDataPos, newCols.length);
    sheet.getRange(1, lastDataPos + 1, 1, newCols.length).setValues([newCols]);
    Logger.log('Columnas nuevas en "' + sheet.getName() + '": ' + newCols.join(', '));
  }
  return readHeader_(sheet);
}

function periodoColumn_(sheet) {
  var idx = readHeader_(sheet).indexOf('periodo');
  if (idx === -1) {
    throw new Error('la pestana "' + sheet.getName() + '" no tiene columna "periodo"');
  }
  return idx + 1;
}

function readPeriodColumn_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var col = periodoColumn_(sheet);
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(function (r) {
    return String(r[0]);
  });
}

function deletePeriodRows_(sheet, periodo) {
  var periods = readPeriodColumn_(sheet);
  var rows = [];
  periods.forEach(function (p, i) {
    if (p === periodo) rows.push(i + 2);
  });
  deleteRowsGrouped_(sheet, rows);
  return rows.length;
}

/** Retencion: conserva los ultimos CONFIG.windowPeriods periodos de la pestana. */
function applyRetention_(master, entity) {
  var sheet = master.getSheetByName(entity.destino);
  if (!sheet || sheet.getLastRow() < 2) return;

  var periods = readPeriodColumn_(sheet);
  var ordered = distinct_(periods).sort();
  if (ordered.length <= CONFIG.windowPeriods) return;

  var keep = ordered.slice(ordered.length - CONFIG.windowPeriods);
  var rows = [];
  periods.forEach(function (p, i) {
    if (keep.indexOf(p) === -1) rows.push(i + 2);
  });
  deleteRowsGrouped_(sheet, rows);
  Logger.log('Retencion "' + sheet.getName() + '": ' + rows.length + ' filas eliminadas, quedan ' + keep.length +
    ' periodos [' + keep.join(', ') + ']');
}

/** Borra filas por indice agrupando bloques consecutivos y de abajo arriba. */
function deleteRowsGrouped_(sheet, rows) {
  if (!rows || !rows.length) return;
  var sorted = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < sorted.length) {
    var end = sorted[i];
    var start = end;
    var j = i + 1;
    while (j < sorted.length && sorted[j] === start - 1) {
      start = sorted[j];
      j++;
    }
    sheet.deleteRows(start, end - start + 1);
    i = j;
  }
}

// ============================================================================
// ORIGENES
// ============================================================================

function listSourceFiles_(folder) {
  var re = new RegExp(pattern_());
  var it = folder.getFiles();
  var out = [];
  while (it.hasNext()) {
    var file = it.next();
    var base = file.getName().replace(/\.[A-Za-z0-9]+$/, '');
    var m = base.match(re);
    if (!m) continue;

    var year = Number(m[1]);
    var num = Number(m[2]);
    var periodo = (CONFIG.periodFormat === 'MONTH') ? (m[1] + '-' + pad_(num)) : (m[1] + '-W' + pad_(num));

    out.push({
      id: file.getId(),
      name: file.getName(),
      periodo: periodo,
      anio: year,
      semana: (CONFIG.periodFormat === 'MONTH') ? '' : num,
      fechaReporte: fechaReporte_(periodo)
    });
  }
  out.sort(function (a, b) { return a.periodo < b.periodo ? -1 : (a.periodo > b.periodo ? 1 : 0); });
  return out;
}

function pattern_() {
  if (CONFIG.filenamePattern) return CONFIG.filenamePattern;
  return (CONFIG.periodFormat === 'MONTH') ? '^(\\d{4})-(\\d{2})$' : '^(\\d{4})-W(\\d{2})$';
}

/** Abre una hoja de Google y rechaza cualquier otro tipo de fichero. */
function openSource_(fileId) {
  var file = DriveApp.getFileById(fileId);
  if (file.getMimeType() !== SHEETS_MIME) {
    throw new Error('"' + file.getName() + '" no es una hoja de Google (tipo: ' +
      file.getMimeType() + '). Solo se admiten hojas de Google.');
  }
  return SpreadsheetApp.openById(fileId);
}

function resolveSheetName_(ss, wanted) {
  var target = normalize_(wanted);
  var matches = ss.getSheets().filter(function (s) { return normalize_(s.getName()) === target; });
  if (matches.length === 1) return matches[0].getName();
  var available = ss.getSheets().map(function (s) { return '"' + s.getName() + '"'; }).join(', ');
  throw new Error('no encuentro una unica pestana llamada "' + wanted + '". Disponibles: ' + available);
}

// ============================================================================
// FECHAS Y PERIODOS
// ============================================================================

/** Primer dia (lunes) de una semana ISO, a mediodia UTC para evitar saltos de zona. */
function isoWeekToMonday_(year, week) {
  var jan4 = new Date(Date.UTC(year, 0, 4, 12, 0, 0));
  var dow = jan4.getUTCDay() || 7;
  var mondayW1 = new Date(jan4.getTime() - (dow - 1) * 86400000);
  return new Date(mondayW1.getTime() + (week - 1) * 7 * 86400000);
}

function fechaReporte_(periodo) {
  if (CONFIG.periodFormat === 'MONTH') {
    var pm = periodo.split('-');
    return new Date(Date.UTC(Number(pm[0]), Number(pm[1]) - 1, 1, 12, 0, 0));
  }
  var pw = periodo.split('-W');
  var week = Number(pw[1]);
  if (!(week >= 1 && week <= 53)) {
    throw new Error('semana ISO fuera de rango: ' + periodo);
  }
  return isoWeekToMonday_(Number(pw[0]), week);
}

function formatDate_(d) {
  return Utilities.formatDate(d, CONFIG.timezone, 'yyyy-MM-dd');
}

// ============================================================================
// UTILIDADES
// ============================================================================

function normalize_(s) {
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function pad_(n) {
  return (n < 10 ? '0' : '') + n;
}

function distinct_(arr) {
  var seen = {};
  var out = [];
  arr.forEach(function (v) {
    if (v && !seen[v]) {
      seen[v] = true;
      out.push(v);
    }
  });
  return out;
}

function assertHeader_(header, sheetName) {
  var seen = {};
  header.forEach(function (c) {
    if (!c) {
      throw new Error('cabecera vacia en la pestana "' + sheetName + '"');
    }
    var key = normalize_(c);
    if (seen[key]) {
      throw new Error('cabecera duplicada en la pestana "' + sheetName + '": "' + c + '"');
    }
    seen[key] = true;
  });
  META_HEADERS.concat([CACHE_HEADER]).forEach(function (reserved) {
    header.forEach(function (c) {
      if (normalize_(c) === reserved) {
        throw new Error('la pestana "' + sheetName + '" usa la columna "' + c + '", reservada por el maestro');
      }
    });
  });
}

function assertNoErrors_(rows, sheetName, headerRow) {
  var pattern = /^#(REF|N\/A|DIV\/0|VALUE|NAME|NULL|NUM|ERROR)/;
  var bad = [];
  rows.forEach(function (row, ri) {
    row.forEach(function (cell) {
      if (typeof cell === 'string' && pattern.test(cell.trim())) {
        if (bad.length < 5) bad.push('fila ' + (ri + headerRow + 1) + ': ' + cell.trim());
      }
    });
  });
  if (bad.length) {
    throw new Error('celdas con error en la pestana "' + sheetName + '": ' + bad.join(' | '));
  }
}

function notify_(subject, body) {
  try {
    MailApp.sendEmail(CONFIG.notifyEmail, '[Historico] ' + subject, body);
  } catch (e) {
    Logger.log('No se pudo enviar el correo de aviso: ' + e.message);
  }
}
