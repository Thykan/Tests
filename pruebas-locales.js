/**
 * Pruebas locales de HistoricoMaestro.gs, sin Google Apps Script real.
 * Simula SpreadsheetApp, DriveApp, MailApp y LockService en memoria y ejecuta
 * el motor completo: fechas ISO, upsert, retencion, deriva de columnas, avisos.
 *
 * Ejecutar:  bun pruebas-locales.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = fs.readFileSync(path.join(__dirname, 'HistoricoMaestro.gs'), 'utf8');

let fails = 0;
function check(name, cond, extra) {
  if (cond) { console.log('PASS  ' + name); }
  else { fails++; console.log('FAIL  ' + name + (extra !== undefined ? '  -> ' + JSON.stringify(extra) : '')); }
}

function cellAt(rows, r, c) {
  const row = rows[r - 1];
  if (!row) return '';
  const v = row[c - 1];
  return v === undefined || v === null ? '' : v;
}
function lastRow(rows) {
  for (let r = rows.length; r >= 1; r--) {
    if ((rows[r - 1] || []).some(v => v !== '' && v != null)) return r;
  }
  return 0;
}
function lastCol(rows) {
  let max = 0;
  rows.forEach(row => {
    if (!row) return;
    for (let c = row.length; c >= 1; c--) {
      const v = row[c - 1];
      if (v !== '' && v != null) { if (c > max) max = c; break; }
    }
  });
  return max;
}
function makeSheet(name, rows) {
  const data = rows ? rows.map(r => r.slice()) : [];
  const ensure = (r, c) => {
    while (data.length < r) data.push([]);
    data.forEach(row => { while (row.length < c) row.push(''); });
  };
  return {
    getName: () => name,
    getLastRow: () => lastRow(data),
    getLastColumn: () => lastCol(data),
    setFrozenRows: () => {},
    insertColumnsAfter(pos, k) {
      data.forEach(row => { row.splice(pos, 0, ...Array(k).fill('')); });
    },
    deleteRows(start, k) { data.splice(start - 1, k); },
    getRange(r, c, nr, nc) {
      return {
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = [];
            for (let j = 0; j < nc; j++) row.push(cellAt(data, r + i, c + j));
            out.push(row);
          }
          return out;
        },
        setValues(vals) {
          ensure(r + nr - 1, c + nc - 1);
          for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = vals[i][j];
          return this;
        },
        setNumberFormat() { return this; }
      };
    },
    _dump: () => data.map(r => r.slice())
  };
}
function makeSpreadsheet(sheets) {
  const list = sheets || [];
  return {
    getSheetByName: n => list.find(s => s.getName() === n) || null,
    getSheets: () => list,
    insertSheet: n => { const s = makeSheet(n); list.push(s); return s; }
  };
}
function makeFolder(items) {
  return {
    getFiles() {
      let i = 0;
      return { hasNext: () => i < items.length, next: () => items[i++] };
    }
  };
}
function makeFile(id, name, mime) {
  return { getId: () => id, getName: () => name, getMimeType: () => mime };
}

const sheets = {};
const files = {};
const logs = [];
const mails = [];
let folderFiles = [];

const ctx = {
  console,
  Date, RegExp, String, Number, Array, Object, JSON, Math, Error, Set,
  SpreadsheetApp: { openById: id => sheets[id], flush: () => {} },
  DriveApp: {
    getFileById: id => {
      const f = files[id];
      if (!f) throw new Error('fichero no encontrado: ' + id);
      return f;
    },
    getFolderById: () => makeFolder(folderFiles)
  },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  ScriptApp: {
    getOAuthToken: () => 'tok',
    getProjectTriggers: () => [],
    deleteTrigger: () => {},
    newTrigger: () => { throw new Error('no en test'); }
  },
  MailApp: { sendEmail: (to, subj, body) => mails.push({ to, subj, body }) },
  Utilities: { formatDate: d => d.toISOString().slice(0, 10) },
  Logger: { log: m => logs.push(m) }
};
vm.createContext(ctx);
vm.runInContext(SRC, ctx);

// 1. fechas ISO de referencia
const d = p => ctx.fechaReporte_(p).toISOString().slice(0, 10);
check('ISO 2026-W01 = 2025-12-29', d('2026-W01') === '2025-12-29', d('2026-W01'));
check('ISO 2026-W47 = 2026-11-16', d('2026-W47') === '2026-11-16', d('2026-W47'));
check('ISO 2024-W01 = 2024-01-01', d('2024-W01') === '2024-01-01', d('2024-W01'));
check('ISO 2020-W53 = 2020-12-28', d('2020-W53') === '2020-12-28', d('2020-W53'));
check('ISO 2026-W52 = 2026-12-21', d('2026-W52') === '2026-12-21', d('2026-W52'));

// 2. borrado agrupado: 1=h 2=a 3=b 4=c 5=d 6=e 7=f, fuera 2,3,5,7 -> h,c,e
(function () {
  const s = makeSheet('t', [['h'], ['a'], ['b'], ['c'], ['d'], ['e'], ['f']]);
  ctx.deleteRowsGrouped_(s, [2, 3, 5, 7]);
  check('deleteRowsGrouped_ deja [h,c,e]', JSON.stringify(s._dump()) === JSON.stringify([['h'], ['c'], ['e']]), s._dump());
})();

// 3. retencion por ventana de periodos
(function () {
  const rows = [['periodo', 'anio', 'semana', 'fecha_reporte', 'x', 'fecha_carga']];
  ['2026-W01', '2026-W02', '2026-W03', '2026-W10'].forEach(p => {
    rows.push([p, 2026, 1, '', 'v', '']);
    rows.push([p, 2026, 1, '', 'v2', '']);
  });
  const master = makeSpreadsheet([makeSheet('ent', rows)]);
  ctx.CONFIG.masterFileId = 'm1';
  ctx.CONFIG.windowPeriods = 2;
  ctx.applyRetention_(master, { destino: 'ent' });
  const periods = [...new Set(master.getSheetByName('ent')._dump().slice(1).map(r => r[0]))];
  check('retencion conserva los 2 ultimos periodos', JSON.stringify(periods) === JSON.stringify(['2026-W03', '2026-W10']), periods);
  check('retencion mantiene cabecera', master.getSheetByName('ent')._dump()[0][0] === 'periodo');
})();

// 4. ingesta end-to-end semanal
(function () {
  const SPM = 'application/vnd.google-apps.spreadsheet';
  const master = makeSpreadsheet([]);
  sheets['MASTER'] = master;

  const w46 = makeSpreadsheet([
    makeSheet('Hoja 1', [['cliente', 'importe'], ['ACME', 100], ['Globex', 250]]),
    makeSheet('Hoja 2', [['kpi', 'valor'], ['visitas', 900]]),
    makeSheet('Resumen', [['esto', 'se ignora'], ['x', 1]])
  ]);
  const w47 = makeSpreadsheet([
    makeSheet('Hoja 1', [['cliente', 'importe', 'canal'], ['Initech', 400, 'web'], ['ACME', 120, 'email']]),
    makeSheet('Hoja 2', [['kpi', 'valor'], ['visitas', 1100]])
  ]);
  sheets['W46'] = w46;
  sheets['W47'] = w47;
  files['W46'] = makeFile('W46', '2026-W46.xlsx', SPM);
  files['W47'] = makeFile('W47', '2026-W47.xlsx', SPM);
  files['OTHER'] = makeFile('OTHER', 'informe_borrador.xlsx', SPM);
  folderFiles = [files['W46'], files['W47'], files['OTHER']];

  ctx.CONFIG.masterFileId = 'MASTER';
  ctx.CONFIG.periodFormat = 'ISO_WEEK';
  ctx.CONFIG.windowPeriods = 12;
  ctx.CONFIG.failOnErrorCells = true;
  ctx.CONFIG.entities = [
    { destino: 'entidad_a', origen: 'Hoja 1' },
    { destino: 'entidad_b', origen: 'Hoja 2' }
  ];

  const s1 = ctx.ingest_({ onlyMissing: true });
  const a = master.getSheetByName('entidad_a');
  const b = master.getSheetByName('entidad_b');

  check('ingesta: 4 cargas (2 entidades x 2 periodos)', s1.ok.length === 4, s1);
  check('ingesta: no crea pestana Resumen', master.getSheetByName('Resumen') === null);
  check('ingesta: entidad_a 4 filas', a.getLastRow() === 5, a._dump());
  check('ingesta: cabecera con columna nueva al final', JSON.stringify(a._dump()[0]) ===
    JSON.stringify(['periodo', 'anio', 'semana', 'fecha_reporte', 'cliente', 'importe', 'canal', 'fecha_carga']), a._dump()[0]);
  check('ingesta: periodo correcto en meta', a._dump()[1][0] === '2026-W46' && a._dump()[2][0] === '2026-W46');
  check('ingesta: fila W46 sin canal queda vacia', a._dump()[1][6] === '', a._dump()[1]);
  check('ingesta: anio y semana', a._dump()[1][1] === 2026 && a._dump()[1][2] === 46);
  check('ingesta: fecha_reporte es Date', a._dump()[1][3] instanceof Date);
  check('ingesta: entidad_b 3 filas', b.getLastRow() === 3);

  const s2 = ctx.ingest_({ onlyMissing: true });
  check('2a ejecucion omite todo', s2.ok.length === 0 && s2.skipped.length === 2, s2);
  check('2a ejecucion no duplica', a.getLastRow() === 5);

  ctx.ingest_({ onlyMissing: false });
  check('rebuild no duplica (upsert)', a.getLastRow() === 5 && b.getLastRow() === 3, { a: a.getLastRow(), b: b.getLastRow() });

  ctx.CONFIG.windowPeriods = 1;
  ctx.applyRetention_(master, { destino: 'entidad_a' });
  const left = a._dump().slice(1).map(r => r[0]);
  check('retencion ventana=1 deja solo 2026-W47', JSON.stringify([...new Set(left)]) === JSON.stringify(['2026-W47']), left);

  ctx.CONFIG.windowPeriods = 12;
  const w48 = makeSpreadsheet([
    makeSheet('Hoja 1', [['cliente', 'importe_total'], ['Umbrella', 999]]),
    makeSheet('Hoja 2', [['kpi', 'valor'], ['visitas', 1]])
  ]);
  sheets['W48'] = w48;
  files['W48'] = makeFile('W48', '2026-W48.xlsx', SPM);
  folderFiles = [files['W48']];
  let aborted = null;
  try { ctx.ingest_({ onlyMissing: true }); } catch (e) { aborted = e.message; }
  check('renombrado de columna aborta', !!aborted && /importe/.test(aborted), aborted);

  const w49 = makeSpreadsheet([
    makeSheet('Hoja 1', [['cliente', 'importe', 'canal', 'pais'], ['Stark', 50, 'web', 'ES']]),
    makeSheet('Hoja 2', [['kpi', 'valor'], ['visitas', 5]])
  ]);
  sheets['W49'] = w49;
  files['W49'] = makeFile('W49', '2026-W49.xlsx', SPM);
  folderFiles = [files['W49']];
  ctx.ingest_({ onlyMissing: true });
  check('columna nueva pais anadida', a._dump()[0].indexOf('pais') !== -1, a._dump()[0]);
  check('pais queda antes de fecha_carga', a._dump()[0][a._dump()[0].length - 1] === 'fecha_carga');

  let noSheet = null;
  try {
    const bad = makeSpreadsheet([makeSheet('Otra', [['a'], [1]])]);
    sheets['BAD'] = bad;
    files['BAD'] = makeFile('BAD', '2026-W50.xlsx', SPM);
    folderFiles = [files['BAD']];
    ctx.ingest_({ onlyMissing: true });
  } catch (e) { noSheet = e.message; }
  check('pestana origen inexistente da error claro', !!noSheet && /Hoja 1/.test(noSheet), noSheet);
})();

// 5. formato mensual
(function () {
  const SPM = 'application/vnd.google-apps.spreadsheet';
  const master = makeSpreadsheet([]);
  sheets['M2'] = master;
  sheets['M2SRC'] = makeSpreadsheet([makeSheet('datos', [['v'], [10]])]);
  files['M2SRC'] = makeFile('M2SRC', '2026-11.xlsx', SPM);
  folderFiles = [files['M2SRC']];

  ctx.CONFIG.masterFileId = 'M2';
  ctx.CONFIG.periodFormat = 'MONTH';
  ctx.CONFIG.windowPeriods = 12;
  ctx.CONFIG.entities = [{ destino: 'mes', origen: 'datos' }];

  ctx.ingest_({ onlyMissing: true });
  const s = master.getSheetByName('mes');
  check('mensual: periodo 2026-11', s._dump()[1][0] === '2026-11', s._dump());
  check('mensual: semana vacia', s._dump()[1][2] === '', s._dump()[1]);
  check('mensual: fecha_reporte dia 1', s._dump()[1][3].toISOString().slice(0, 10) === '2026-11-01', s._dump()[1][3]);
})();

// 6. resolucion de nombres de pestana
(function () {
  const sp = makeSpreadsheet([makeSheet('Datos ', []), makeSheet('Otro', [])]);
  check('resolveSheetName_ tolera espacios', ctx.resolveSheetName_(sp, 'datos') === 'Datos ');
  let dup = null;
  try {
    const sp2 = makeSpreadsheet([makeSheet('Hoja 1', []), makeSheet('Hoja 1 ', [])]);
    ctx.resolveSheetName_(sp2, 'Hoja 1');
  } catch (e) { dup = e.message; }
  check('resolveSheetName_ falla con duplicados', !!dup && /Disponibles/.test(dup), dup);
})();

// 7. columna ausente en periodo antiguo: aviso, no aborto
(function () {
  const SPM = 'application/vnd.google-apps.spreadsheet';
  const master = makeSpreadsheet([]);
  sheets['M3'] = master;
  sheets['M3N'] = makeSpreadsheet([makeSheet('datos', [['a', 'b2'], [1, 2]])]);
  sheets['M3V'] = makeSpreadsheet([makeSheet('datos', [['a'], [9]])]);
  files['M3N'] = makeFile('M3N', '2026-W48.xlsx', SPM);
  files['M3V'] = makeFile('M3V', '2026-W47.xlsx', SPM);
  folderFiles = [files['M3V'], files['M3N']];

  ctx.CONFIG.masterFileId = 'M3';
  ctx.CONFIG.periodFormat = 'ISO_WEEK';
  ctx.CONFIG.windowPeriods = 12;
  ctx.CONFIG.entities = [{ destino: 'd', origen: 'datos' }];

  ctx.ingest_({ onlyMissing: true });
  const r = ctx.ingest_({ onlyMissing: false });

  check('aviso al reingestar periodo antiguo sin columna', r.warned.length === 1 && /b2/.test(r.warned[0]), r.warned);
  const dump = master.getSheetByName('d')._dump();
  check('periodo antiguo conserva fila con b2 vacio', dump.length === 3 && dump[1][0] === '2026-W47' && dump[1][5] === '', dump);

  let text = null;
  try { text = ctx.dryRun(); } catch (e) { text = 'ERR ' + e.message; }
  check('dryRun detecta ficheros y estado', /Ficheros detectados en la carpeta: 2/.test(text) && /"d"/.test(text), text);
})();

console.log('\n' + (fails === 0 ? 'TODO OK' : fails + ' FALLOS'));
process.exit(fails === 0 ? 0 : 1);
