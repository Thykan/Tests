/**
 * ============================================================================
 * HISTORICO MAESTRO PARA LOOKER STUDIO  -  Apps Script (V8)
 * ============================================================================
 *
 * QUE HACE ESTE SCRIPT
 *   Coge los ficheros de periodo que ya se guardan en una carpeta de Drive (una
 *   hoja de Google por semana o por mes), les lee las pestanas de datos y va
 *   copiando las filas a un unico fichero "maestro" con ID fijo. Ese maestro es
 *   el que lee Looker Studio, y al tener siempre el mismo ID nunca se rompe la
 *   conexion. Los periodos viejos se van borrando para que el fichero no crezca.
 *
 *   Nunca se crea una pestana por periodo. Se crea UNA PESTANA POR ENTIDAD
 *   (una cosa que se mide), y todas las filas de todos los periodos se apilan
 *   ahi dentro. La columna "periodo" dice de que semana o mes viene cada fila.
 *
 * COMO QUEDA EL MAESTRO (por cada entidad)
 *   fila 1:  periodo | anio | semana | fecha_reporte | <columnas del origen> | fecha_carga
 *   fila 2+: una fila por registro, con el periodo del fichero del que viene.
 *
 * ----------------------------------------------------------------------------
 * COMO LEER ESTE FICHERO SI NO PROGRAMAS JAVASCRIPT
 * ----------------------------------------------------------------------------
 *
 *   function nombre(...) { ... }
 *       Es una "funcion": un bloque de instrucciones con nombre. Cuando tu
 *       pulsas Ejecutar en el editor de Apps Script, se ejecuta UNA funcion
 *       concreta: la que tengas seleccionada en el desplegable de arriba.
 *
 *   var algo = ...
 *       Crea una caja con nombre donde se guarda un valor. "var" es la forma
 *       antigua de declararla; aqui se usa para que el codigo sea compatible
 *       con cualquier version del motor de Apps Script.
 *
 *   return algo;
 *       Devuelve un valor como resultado de la funcion y termina ahi mismo.
 *       Lo que hay debajo de un "return" ya no se ejecuta.
 *
 *   throw new Error('mensaje');
 *       Lanza un error a proposito: la ejecucion se corta y el mensaje se ve en
 *       el registro (Ver > Registro de ejecuciones). Se usa para avisar de algo
 *       que no debe pasar en silencio, como una columna que ha desaparecido.
 *
 *   try { ... } catch (err) { ... }
 *       "Intenta esto; si falla, haz esto otro en vez de romperse". El error
 *       queda en la variable "err" y ahi se puede leer err.message.
 *
 *   lista.forEach(function (x) { ... })
 *       "Para cada elemento de la lista, haz esto". Es la forma de recorrer una
 *       lista sin escribir un bucle a mano.
 *
 *   lista.map(function (x) { ... return y; })
 *       "Convierte cada elemento de la lista en otra cosa". Devuelve una lista
 *       nueva del mismo tamano, transformada.
 *
 *   lista.filter(function (x) { ... })
 *       "Quedate solo con los elementos que cumplan la condicion". Devuelve una
 *       lista nueva mas corta.
 *
 *   funcion_(...)
 *       El guion bajo al final es solo una convencion de nombre: significa
 *       "funcion interna, no la pulses tu en el editor". No cambia como
 *       funciona; sirve para distinguir de un vistazo los botones de verdad.
 *
 *   OBJETOS DE GOOGLE QUE APARECEN (todos empiezan por mayuscula)
 *       SpreadsheetApp : abre hojas de calculo por su ID.
 *       DriveApp       : accede a ficheros y carpetas de Drive.
 *       Sheet          : una pestana concreta dentro de una hoja.
 *       Range          : un trozo rectangular de celdas.
 *       LockService    : impide que el script se ejecute dos veces a la vez.
 *       MailApp        : manda correos.
 *       Logger         : escribe en el registro de ejecuciones.
 *       ScriptApp      : gestiona los disparadores automaticos (triggers).
 *       Utilities      : utilidades varias, como formatear fechas.
 *
 *   getRange(fila, columna, numeroDeFilas, numeroDeColumnas)
 *       Selecciona un rectangulo de celdas. Ojo: las filas y columnas empiezan
 *       a contar en 1 (no en 0). getRange(2, 1, 10, 3) = desde la fila 2,
 *       columna 1 (A), 10 filas hacia abajo y 3 columnas a la derecha (A:C).
 *
 * ----------------------------------------------------------------------------
 * GARANTIAS DE DISENO
 * ----------------------------------------------------------------------------
 *   - Idempotente: reejecutar no duplica. Si el periodo ya existe, se borra y
 *     se reescribe (upsert).
 *   - Cabeceras mapeadas POR NOMBRE, nunca por posicion. Si el origen reordena
 *     sus columnas, los datos siguen cayendo en el sitio correcto.
 *   - Columnas nuevas en el origen: se anaden solas al maestro.
 *   - Columnas que faltan: si faltan en el fichero MAS RECIENTE se considera
 *     renombrado/borrado y aborta. Si faltan en un fichero antiguo (la columna
 *     se creo despues) se rellenan vacias y se avisa en el resumen.
 *   - Retencion por pestana: conserva los ultimos CONFIG.windowPeriods periodos.
 *   - La fila 1 no se toca jamas: el data source de Looker no se rompe.
 *   - Solo hojas de Google: si en la carpeta aparece otro tipo de fichero con
 *     nombre de periodo, se rechaza con error explicito.
 *
 * ----------------------------------------------------------------------------
 * PUESTA EN MARCHA (en este orden)
 * ----------------------------------------------------------------------------
 *   1. Rellena CONFIG (IDs, correo y entidades).
 *   2. Ejecuta dryRun() y revisa el registro: ficheros detectados y estado.
 *   3. Pon un periodo en MANUAL_PERIOD y ejecuta ingestManual().
 *   4. Ejecuta installTrigger() para dejarlo automatico.
 *   5. Comparte el maestro SOLO EN LECTURA. Nadie debe editar ni formatear ese
 *      fichero: aplicar formato a columnas enteras lo devuelve al problema de
 *      memoria del navegador.
 *
 * ----------------------------------------------------------------------------
 * NOTAS DE MANTENIMIENTO
 * ----------------------------------------------------------------------------
 *   - El maestro debe vivir en Mi unidad, no en una unidad compartida: el
 *     conector de Google Sheets de Looker Studio no accede a unidades compartidas.
 *   - El scope de Drive es de solo lectura: el script nunca modifica ni borra
 *     nada fuera del fichero maestro.
 *   - Reconstruir muchos periodos de golpe puede superar los 6 minutos de
 *     ejecucion: hazlo por tandas con ingestPeriod('Y26W48').
 *   - Los numeros con formato local ("1.234,56") viajan como texto. Si Looker
 *     los necesita como numero, arreglalo en el origen.
 *   - Este fichero esta escrito sin tildes ni enes a proposito, para que no haya
 *     problemas de codificacion al copiarlo y pegarlo.
 * ============================================================================
 */

// ============================================================================
// CONFIGURACION  -  lo unico que tienes que editar
// ============================================================================

// CONFIG es el objeto (la caja grande) con todos los ajustes. Solo cambia los
// valores que hay a la derecha de cada dos puntos. No borres las comas.
var CONFIG = {

  // Fichero maestro. El ID es la parte larga que aparece en su URL, entre
  // /spreadsheets/d/ y /edit. Debe estar en Mi unidad, no en unidad compartida.
  masterFileId: 'PON_AQUI_EL_ID_DEL_MAESTRO',

  // Carpeta de Drive donde estan los ficheros de periodo. Solo mira los ficheros
  // que cuelgan directamente de ella; no entra en subcarpetas.
  sourceFolderId: 'PON_AQUI_EL_ID_DE_LA_CARPETA',

  // Correo al que llega el aviso cuando una ejecucion falla.
  notifyEmail: 'tu.correo@empresa.com',

  // Zona horaria usada para interpretar fechas y para la hora del trigger.
  timezone: 'Europe/Madrid',

  // Cuantos periodos se conservan en cada pestana del maestro. Los mas viejos
  // se borran automaticamente. 12 semanas son unos 3 meses de historico.
  windowPeriods: 12,

  // Que tipo de periodo maneja este script:
  //   'ISO_WEEK' -> semanal, nombres tipo Y26W48 (semana ISO, empieza en lunes)
  //   'MONTH'    -> mensual, nombres tipo Y26M11
  // La letra W o M va en el propio nombre del fichero; este ajuste sirve para
  // saber que letra esperar y para poner el trigger semanal o mensual.
  periodFormat: 'ISO_WEEK',

  // Patron del nombre de fichero. Si lo dejas vacio se deduce solo de
  // periodFormat. Solo lo cambias si tus ficheros se llaman de otra forma.
  // Ejemplos para el formato Y26W48 / Y26M11:
  //   '^Y(\\d{2})W(\\d{1,2})$'   (semanal)
  //   '^Y(\\d{2})M(\\d{1,2})$'   (mensual)
  // Los dobles backslash son obligatorios: es la forma de escribir \d en texto.
  // \\d{1,2} acepta tanto Y26W5 como Y26W05.
  filenamePattern: '',

  // Fila donde empiezan las cabeceras en las pestanas de origen. Casi siempre 1.
  headerRow: 1,

  // Si es true, la ejecucion se corta cuando encuentra celdas con error
  // (#REF!, #N/A, #DIV/0!...). Mejor dejarlo en true: no quieres archivar basura.
  failOnErrorCells: true,

  // Una entrada por cada pestana de datos que quieras archivar.
  // Cada entrada es un objeto con tres campos posibles:
  //   destino  : nombre de la pestana en el maestro. Si no existe, se crea.
  //              Es el nombre que veras luego en Looker Studio.
  //   origen   : nombre de la pestana dentro del fichero de periodo.
  //   headerRow: opcional. Solo si esa pestana tiene las cabeceras en otra fila
  //              distinta de CONFIG.headerRow.
  // Puedes tener tantas entradas como necesites; con dejar las que uses basta.
  entities: [
    { destino: 'entidad_a', origen: 'Hoja 1' },
    { destino: 'entidad_b', origen: 'Hoja 2' },
    { destino: 'entidad_c', origen: 'Hoja 3' }
  ]
};

// Periodo que usa ingestManual() cuando lo pulsas a mano. Cambialo aqui y pulsa
// Ejecutar. El formato tiene que coincidir con periodFormat: Y26W48 o Y26M11.
var MANUAL_PERIOD = 'Y26W48';

// ============================================================================
// CONSTANTES INTERNAS  -  no hace falta tocarlas
// ============================================================================

// Columnas que el script pone siempre al principio de cada pestana del maestro.
// El orden importa: es el contrato con Looker Studio y no debe cambiar.
var META_HEADERS = ['periodo', 'anio', 'semana', 'fecha_reporte'];

// Columna que se pone siempre al final: cuando se copio esa fila al maestro.
var CACHE_HEADER = 'fecha_carga';

// Cuanto espera (en milisegundos) si ya hay otra ejecucion en marcha antes de
// rendirse. 30000 = 30 segundos.
var LOCK_WAIT_MS = 30000;

// Identificador interno que usa Google para decir "esto es una hoja de calculo".
var SHEETS_MIME = 'application/vnd.google-apps.spreadsheet';

// ============================================================================
// PUNTOS DE ENTRADA  -  las funciones que tu pulsas o que lanza el trigger
// ============================================================================
//
// Todo lo que empieza por "function" y NO acaba en guion bajo es un boton que
// puedes ejecutar desde el editor de Apps Script (desplegable de funciones +
// boton Ejecutar) o que puede lanzar un trigger automatico.

/**
 * QUE HACE: ejecucion automatica del dia a dia. Recorre la carpeta y carga en el
 *           maestro solo los periodos que todavia no estan dentro. Si un periodo
 *           ya esta, lo salta. Es lo que lanza el trigger semanal o mensual.
 * CUANDO SE EJECUTA: solo por el trigger automatico (o a mano, si quieres
 *           simular lo que hara el trigger).
 * NO HACE: no rehace periodos ya cargados. Para eso esta rebuildAll().
 */
function ingestScheduled() {
  // execute_ es el envoltorio que pone el candado y manda correo si algo falla.
  // Le pasamos una funcion anonima: el trabajo real que debe ejecutar.
  execute_(function () {
    return ingest_({ onlyMissing: true });
  });
}

/**
 * QUE HACE: carga inicial o reconstruccion completa. Vuelve a procesar todos los
 *           ficheros de la carpeta, reescribiendo cada periodo que encuentre.
 * CUANDO SE EJECUTA: la primera vez que montas el maestro, o si sospechas que
 *           hay periodos mal cargados.
 * OJO: con muchos periodos puede superar el limite de 6 minutos por ejecucion.
 *      En ese caso, hazlo por tandas con ingestPeriod('Y26W40') y siguientes.
 */
function rebuildAll() {
  execute_(function () {
    return ingest_({ onlyMissing: false });
  });
}

/**
 * QUE HACE: rehace UN periodo concreto, el que tengas en MANUAL_PERIOD.
 * CUANDO SE EJECUTA: cuando alguien corrige un fichero ya cargado y quieres
 *           reflejar la correccion sin tocar el resto.
 */
function ingestManual() {
  ingestPeriod(MANUAL_PERIOD);
}

/**
 * QUE HACE: rehace el periodo que le pases entre comillas como argumento.
 * CUANDO SE EJECUTA: desde el editor, cambiando MANUAL_PERIOD, o desde otro
 *           script llamandolo como ingestPeriod('Y26W48').
 */
function ingestPeriod(periodo) {
  execute_(function () {
    return ingest_({ onlyMissing: false, onlyPeriod: String(periodo) });
  });
}

/**
 * QUE HACE: diagnostico de solo lectura. No escribe nada en ningun sitio: solo
 *           mira la carpeta y el maestro, y escribe un resumen en el registro.
 * CUANDO SE EJECUTA: siempre lo PRIMERO, antes de instalar el trigger. Sirve
 *           para comprobar que los IDs son correctos y que el script ve los
 *           ficheros y las pestanas que esperas.
 * COMO VER EL RESULTADO: Ver > Registro de ejecuciones (o Ctrl+Enter).
 */
function dryRun() {
  var master = SpreadsheetApp.openById(CONFIG.masterFileId);
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var files = listSourceFiles_(folder);

  // "lines" es una lista de textos; al final se unen todos con saltos de linea.
  var lines = [];

  lines.push('Formato de periodo: ' + CONFIG.periodFormat + ' | ventana: ' + CONFIG.windowPeriods);
  lines.push('Ficheros detectados en la carpeta: ' + files.length);

  // Una linea por fichero: nombre, periodo deducido y fecha que se archivara.
  files.forEach(function (f) {
    lines.push('  - ' + f.name + '  ->  ' + f.periodo + '  (fecha_reporte ' + formatDate_(fechaReporte_(f.periodo)) + ')');
  });

  // Aviso util: ficheros con pinta de periodo pero del otro tipo (por ejemplo
  // un mensual Y26M11 en un script semanal). Al cargar se ignoran en silencio,
  // asi que aqui se delatan para que no parezca que el script no los ve.
  var letraEsperada = (CONFIG.periodFormat === 'MONTH') ? 'M' : 'W';
  var parecePeriodo = /^Y\d{2}([WM])\d{1,2}$/;
  var descartados = [];
  var it2 = folder.getFiles();
  while (it2.hasNext()) {
    var otro = it2.next();
    var otroBase = otro.getName().replace(/\.[A-Za-z0-9]+$/, '');
    var coincidencia = otroBase.match(parecePeriodo);
    if (coincidencia && coincidencia[1] !== letraEsperada) {
      descartados.push(otro.getName() + ' (tipo ' + coincidencia[1] + ', se esperaba ' + letraEsperada + ')');
    }
  }
  if (descartados.length) {
    lines.push('OJO, ignorados con pinta de periodo: ' + descartados.join(' | '));
  }

  lines.push('');
  lines.push('Estado del maestro:');

  // Una linea por entidad: si la pestana existe, cuantas filas y periodos tiene.
  CONFIG.entities.forEach(function (e) {
    var sheet = master.getSheetByName(e.destino);

    // Si la pestana no existe todavia, se creara en la primera carga real.
    if (!sheet) {
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): no existe, se creara');
      return;
    }

    // Se protege con try/catch para que una pestana mal montada no impida
    // diagnosticar las demas: se informa del problema y se sigue.
    try {
      var periods = distinct_(readPeriodColumn_(sheet));
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): ' + Math.max(sheet.getLastRow() - 1, 0) +
        ' filas, ' + periods.length + ' periodos' + (periods.length ? ' [' + periods.join(', ') + ']' : ''));
    } catch (err) {
      lines.push('  "' + e.destino + '" (origen "' + e.origen + '"): PROBLEMA -> ' + err.message);
    }
  });

  // join('\n') pega todas las lineas separandolas con un salto de linea.
  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * QUE HACE: deja programada la ejecucion automatica (el "trigger").
 * CUANDO SE EJECUTA: una sola vez, cuando ya has probado dryRun e ingestManual.
 * DETALLE: lunes a las 06:00 si es semanal; dia 1 a las 06:00 si es mensual.
 *          Antes de crear uno nuevo borra el anterior, para no duplicar.
 */
function installTrigger() {
  removeTriggers();

  // Se va montando el trigger por partes: primero "que funcion llamar" y con
  // que zona horaria; luego la frecuencia.
  var builder = ScriptApp.newTrigger('ingestScheduled').timeBased().inTimezone(CONFIG.timezone);
  if (CONFIG.periodFormat === 'MONTH') {
    builder.onMonthDay(1).atHour(6).create();
  } else {
    builder.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  }

  Logger.log('Trigger instalado.');
}

/**
 * QUE HACE: borra el trigger automatico de este script, si existe.
 * CUANDO SE EJECUTA: si quieres parar la automatizacion sin borrar el proyecto.
 */
function removeTriggers() {
  // getProjectTriggers() devuelve todos los triggers del proyecto. Solo se
  // borran los que llaman a ingestScheduled: el resto se deja intacto.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'ingestScheduled') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ============================================================================
// MOTOR
// ============================================================================
//
// A partir de aqui son funciones internas (acaban en guion bajo). No las pulses
// a mano: las llaman las de arriba.

/**
 * QUE HACE: envuelve cualquier trabajo en tres protecciones comunes.
 *   1. Un candado (lock) para que dos ejecuciones no escriban a la vez.
 *   2. Si el trabajo termina bien, guarda el resumen en el registro.
 *   3. Si falla, manda un correo con el error y vuelve a lanzarlo para que la
 *      ejecucion quede marcada como fallida en el panel de Apps Script.
 * RECIBE: fn, la funcion con el trabajo real.
 */
function execute_(fn) {
  var lock = LockService.getScriptLock();

  // tryLock devuelve false si otro proceso tiene el candado. En ese caso no se
  // espera indefinidamente: se aborta, porque escribir dos veces seria peor.
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    throw new Error('Ya hay una ejecucion en curso del historico. Abortado para no escribir dos veces.');
  }

  try {
    // Aqui se ejecuta el trabajo de verdad.
    var summary = fn();
    Logger.log(JSON.stringify(summary, null, 2));
    return summary;
  } catch (err) {
    // "err && err.stack" significa: si hay error y tiene traza, usa la traza
    // (que dice la linea exacta); si no, usa el texto del error.
    notify_('ERROR', (err && err.stack) ? err.stack : String(err));
    throw err;
  } finally {
    // "finally" se ejecuta siempre, haya ido bien o mal: asi el candado nunca
    // se queda echado por un fallo.
    lock.releaseLock();
  }
}

/**
 * QUE HACE: el recorrido principal. Lee la carpeta, decide que periodos hay que
 *           cargar y los va archivando entidad por entidad.
 * RECIBE: opts.onlyMissing (true = saltar periodos ya cargados) y
 *         opts.onlyPeriod (cargar solo ese periodo).
 * DEVUELVE: un resumen con lo cargado, lo omitido, los fallos y los avisos.
 */
function ingest_(opts) {
  // Si no llega nada, se usa un objeto vacio, para poder leer opts.onlyMissing
  // sin que el programa falle.
  opts = opts || {};

  var master = SpreadsheetApp.openById(CONFIG.masterFileId);
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);

  // Lista de ficheros de la carpeta que cumplen el patron, ordenada del periodo
  // mas antiguo al mas reciente.
  var allFiles = listSourceFiles_(folder);

  if (!allFiles.length) {
    throw new Error('No hay ficheros en la carpeta que cumplan el patron de nombre. Patron usado: ' + pattern_());
  }

  // El periodo mas reciente de la carpeta define el esquema "vigente": solo ese
  // fichero puede provocar un aborto por columna desaparecida.
  var latestPeriod = allFiles[allFiles.length - 1].periodo;

  // Si se pide un periodo concreto, se queda solo con ese fichero.
  var files = allFiles;
  if (opts.onlyPeriod) {
    files = allFiles.filter(function (f) { return f.periodo === opts.onlyPeriod; });
    if (!files.length) {
      throw new Error('No hay ningun fichero para el periodo ' + opts.onlyPeriod + '. Revisa el nombre del fichero y CONFIG.filenamePattern.');
    }
  }

  // Cache de periodos ya presentes por entidad, para no leer la columna N veces.
  // periodCache['entidad_a']['Y26W47'] = true significa "ese periodo ya esta".
  var periodCache = {};
  CONFIG.entities.forEach(function (e) {
    periodCache[e.destino] = {};
    readPeriodColumn_(master.getSheetByName(e.destino)).forEach(function (p) {
      periodCache[e.destino][p] = true;
    });
  });

  // Aqui se va apuntando lo que pasa, para el resumen final y el correo.
  var summary = { ok: [], skipped: [], failed: [], warned: [] };

  // ---- Bucle principal: un fichero de periodo cada vez ----
  files.forEach(function (entry) {
    // De las entidades configuradas, cuales hay que cargar para este fichero.
    // Con onlyMissing, se descartan las que ya tienen ese periodo.
    var pending = CONFIG.entities.filter(function (e) {
      return !(opts.onlyMissing && periodCache[e.destino][entry.periodo]);
    });

    if (!pending.length) {
      summary.skipped.push(entry.periodo);
      return;
    }

    // Se abre el fichero de origen una sola vez y se reutiliza para todas sus
    // pestanas. Si no se puede abrir, se apunta el fallo y se sigue con el
    // siguiente fichero: un fichero roto no debe parar los demas.
    var srcSs;
    try {
      srcSs = openSource_(entry.id);
    } catch (err) {
      summary.failed.push(entry.periodo + ' / abrir fichero: ' + err.message);
      return;
    }

    // Un fallo en una entidad tampoco debe impedir cargar las otras.
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

  // Retencion al final, una vez por entidad: aqui se borran los periodos viejos.
  CONFIG.entities.forEach(function (e) {
    try {
      applyRetention_(master, e);
    } catch (err) {
      summary.failed.push('retencion / ' + e.destino + ': ' + err.message);
    }
  });

  // Fuerza a que todos los cambios pendientes se escriban antes de terminar.
  SpreadsheetApp.flush();

  // Los avisos no son errores: se registran pero la ejecucion se considera bien.
  if (summary.warned.length) {
    Logger.log('Avisos (' + summary.warned.length + '):\n' + summary.warned.join('\n'));
  }

  // Si hubo fallos, se lanza un error con todos juntos. execute_ lo mandara por
  // correo y el panel de Apps Script marcara la ejecucion como fallida.
  if (summary.failed.length) {
    throw new Error('Fallos (' + summary.failed.length + '):\n' + summary.failed.join('\n') +
      '\n\nCargados: ' + summary.ok.length + ' | Omitidos: ' + summary.skipped.length +
      ' | Avisos: ' + summary.warned.length);
  }
  return summary;
}

/**
 * QUE HACE: archiva una entidad de un fichero de periodo en su pestana del
 *           maestro. Es el corazon del script.
 * RECIBE:
 *   master    : la hoja de calculo maestra ya abierta.
 *   srcSs     : la hoja de calculo de origen ya abierta.
 *   entity    : la entrada de CONFIG.entities que toca ({destino, origen}).
 *   entry     : datos del fichero ({periodo, anio, semana, fechaReporte}).
 *   isLatest  : true si este fichero es el mas reciente de la carpeta; solo
 *               entonces una columna ausente se trata como renombrado y aborta.
 *   warnings  : lista compartida donde se apuntan los avisos no criticos.
 * PASOS: preparar pestana destino, leer origen, validar, alinear columnas,
 *        borrar el periodo si ya estaba y escribir de una sola vez.
 */
function ingestEntity_(master, srcSs, entity, entry, isLatest, warnings) {
  // 1. Pestana destino del maestro (se crea si no existe).
  var masterSheet = ensureMasterTab_(master, entity);

  // 2. Pestana de origen. resolveSheetName_ tolera mayusculas, espacios y
  //    tildes al comparar nombres, y falla con un mensaje util si hay dudas.
  var srcSheet = srcSs.getSheetByName(resolveSheetName_(srcSs, entity.origen));

  var headerRow = entity.headerRow || CONFIG.headerRow;
  var lastRow = srcSheet.getLastRow();
  var lastCol = srcSheet.getLastColumn();

  // Si la pestana esta vacia o solo tiene cabeceras, no hay nada que archivar.
  if (lastRow <= headerRow || lastCol === 0) {
    throw new Error('la pestana "' + entity.origen + '" no tiene datos bajo la fila ' + headerRow);
  }

  // 3. Se lee el bloque entero de una vez (mucho mas rapido que celda a celda).
  var block = srcSheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastCol).getValues();

  // shift() saca la primera fila del bloque (las cabeceras) y deja el resto.
  // A cada nombre se le quitan los espacios sobrantes.
  var srcHeader = block.shift().map(function (h) { return String(h).trim(); });
  assertHeader_(srcHeader, entity.origen);

  // 4. Si alguna celda trae un error de formula, se aborta antes de archivar.
  if (CONFIG.failOnErrorCells) {
    assertNoErrors_(block, entity.origen, headerRow);
  }

  // 5. Se descartan las filas totalmente vacias que suelen quedar al final.
  //    some() devuelve true en cuanto encuentra una celda con algo dentro.
  var rows = block.filter(function (r) {
    return r.some(function (c) { return c !== '' && c !== null; });
  });
  if (!rows.length) {
    throw new Error('la pestana "' + entity.origen + '" no tiene filas con datos');
  }

  // 6. Se alinean las columnas con el maestro (anade las nuevas, valida las
  //    que faltan) y se obtiene la cabecera definitiva del destino.
  var masterHeader = ensureColumns_(masterSheet, srcHeader, isLatest, warnings, entity.destino, entry.periodo);

  // La misma marca de tiempo para todas las filas de esta carga.
  var now = new Date();

  // 7. Se construye la tabla final: por cada fila de origen y por cada columna
  //    del maestro, se decide que valor va ahi. Esto es lo que hace que el
  //    orden de las columnas del origen deje de importar.
  var payload = rows.map(function (r) {
    return masterHeader.map(function (col) {
      if (col === 'periodo') return entry.periodo;
      if (col === 'anio') return entry.anio;
      if (col === 'semana') return entry.semana;
      if (col === 'fecha_reporte') return entry.fechaReporte;
      if (col === CACHE_HEADER) return now;

      // Columna normal: se busca por su nombre en la cabecera de origen.
      // Si no estuviera, se deja vacia en vez de desplazar los datos.
      var i = srcHeader.indexOf(col);
      return i === -1 ? '' : r[i];
    });
  });

  // Upsert: fuera las filas previas de este periodo, luego append de una sola
  // escritura. getLastRow() se consulta despues de borrar, para escribir justo
  // debajo de la ultima fila que queda.
  deletePeriodRows_(masterSheet, entry.periodo);
  var startRow = masterSheet.getLastRow() + 1;
  masterSheet.getRange(startRow, 1, payload.length, masterHeader.length).setValues(payload);
}

// ============================================================================
// ESTRUCTURA DEL MAESTRO
// ============================================================================

/**
 * QUE HACE: devuelve la pestana destino, creandola si no existe, y comprueba
 *           que su cabecera es la esperada.
 * POR QUE: si alguien renombra o reordena las columnas del maestro a mano, el
 *          script debe pararse en vez de escribir en el sitio equivocado.
 */
function ensureMasterTab_(master, entity) {
  var sheet = master.getSheetByName(entity.destino);

  // No existe: se crea vacia y se le pone la cabecera minima.
  if (!sheet) {
    sheet = master.insertSheet(entity.destino);
    Logger.log('Creada la pestana destino "' + entity.destino + '".');
  }

  // Hoja recien creada o completamente vacia: se escribe la cabecera base
  // (las columnas de datos se anaden despues, al leer el primer origen).
  if (sheet.getLastRow() === 0) {
    var header = META_HEADERS.concat([CACHE_HEADER]);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  // La pestana ya tenia contenido: se valida que empiece como debe.
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

/**
 * QUE HACE: lee la fila 1 (las cabeceras) y la devuelve como lista de textos.
 * DETALLE: getValues() devuelve una tabla de filas x columnas; [0] toma la
 *          primera fila. trim() quita espacios sobrantes a cada nombre.
 */
function readHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
}

/**
 * QUE HACE: alinea las columnas del origen con las del maestro.
 *   - Aditivas: se anaden al maestro.
 *   - Ausentes en el fichero mas reciente: aborta (renombrado o borrado real).
 *   - Ausentes en un fichero antiguo: se rellenan vacias y se avisa.
 * DEVUELVE: la cabecera definitiva del maestro, ya con las columnas nuevas.
 */
function ensureColumns_(sheet, srcHeader, isLatest, warnings, destino, periodo) {
  var header = readHeader_(sheet);

  // Posicion (empezando en 1) de la ultima columna de datos: justo antes de
  // fecha_carga, que siempre va al final.
  var lastDataPos = header.length - 1;

  // slice(desde, hasta) coge un trozo de la lista: aqui, solo las columnas de
  // datos del maestro, sin las cuatro de control ni fecha_carga.
  var masterData = header.slice(META_HEADERS.length, header.length - 1);

  // Columnas que el maestro ya tiene y el origen no trae.
  // indexOf devuelve -1 cuando no encuentra el nombre.
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

  // Columnas nuevas: estan en el origen y no en el maestro. Se insertan antes
  // de fecha_carga, para que esa siga siendo la ultima columna.
  var newCols = srcHeader.filter(function (c) { return c !== '' && header.indexOf(c) === -1; });
  if (newCols.length) {
    sheet.insertColumnsAfter(lastDataPos, newCols.length);
    sheet.getRange(1, lastDataPos + 1, 1, newCols.length).setValues([newCols]);
    Logger.log('Columnas nuevas en "' + sheet.getName() + '": ' + newCols.join(', '));
  }
  return readHeader_(sheet);
}

/**
 * QUE HACE: devuelve el numero de columna (empezando en 1) donde esta "periodo".
 * POR QUE: en vez de suponer que es la 1, se busca por nombre; si algun dia
 *          cambia el orden, el script sigue acertando.
 */
function periodoColumn_(sheet) {
  var idx = readHeader_(sheet).indexOf('periodo');
  if (idx === -1) {
    throw new Error('la pestana "' + sheet.getName() + '" no tiene columna "periodo"');
  }
  return idx + 1;
}

/**
 * QUE HACE: devuelve la lista de periodos de la pestana, en el orden en que
 *          estan escritos (una entrada por fila, repetidos incluidos).
 * DETALLE: empieza en la fila 2 porque la 1 son las cabeceras.
 */
function readPeriodColumn_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var col = periodoColumn_(sheet);
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(function (r) {
    return String(r[0]);
  });
}

/**
 * QUE HACE: borra del maestro todas las filas de un periodo concreto.
 * PARA QUE: es la primera mitad del upsert; la segunda es volver a escribirlas.
 * DEVUELVE: cuantas filas se han borrado.
 * DETALLE: se suman 2 al indice porque la lista empieza en la fila 2 (la 1 es
 *          la cabecera y no se toca nunca).
 */
function deletePeriodRows_(sheet, periodo) {
  var periods = readPeriodColumn_(sheet);
  var rows = [];
  periods.forEach(function (p, i) {
    if (p === periodo) rows.push(i + 2);
  });
  deleteRowsGrouped_(sheet, rows);
  return rows.length;
}

/**
 * QUE HACE: retencion. Se queda con los CONFIG.windowPeriods periodos mas
 *          recientes y borra las filas de todos los demas.
 * DETALLE: el orden se calcula con ordenDe_ (numero), nunca comparando el texto
 *          del periodo: "Y26W5" ordenaria detras de "Y26W48" como texto.
 */
function applyRetention_(master, entity) {
  var sheet = master.getSheetByName(entity.destino);
  if (!sheet || sheet.getLastRow() < 2) return;

  var periods = readPeriodColumn_(sheet);
  if (!periods.length) return;

  // Una pestana no puede mezclar semanas (W) y meses (M): sus claves numericas
  // se solaparian (Y26W11 y Y26M11 darian el mismo numero) y se borraria mal.
  var letras = distinct_(periods.map(function (p) { return String(p).charAt(3); }));
  if (letras.length > 1) {
    throw new Error('la pestana "' + sheet.getName() + '" mezcla periodos semanales (W) y mensuales (M). ' +
      'Cada pestana debe contener un solo tipo de periodo.');
  }

  var ordered = distinct_(periods);
  ordered.sort(function (a, b) { return ordenDe_(a) - ordenDe_(b); });

  // Si aun no hay mas periodos que la ventana, no hay nada que borrar.
  if (ordered.length <= CONFIG.windowPeriods) return;

  // Los que se conservan son los ultimos de la lista ordenada.
  var keep = ordered.slice(ordered.length - CONFIG.windowPeriods);

  // Se marcan las filas de los periodos que NO estan en "keep".
  var rows = [];
  periods.forEach(function (p, i) {
    if (keep.indexOf(p) === -1) rows.push(i + 2);
  });
  deleteRowsGrouped_(sheet, rows);
  Logger.log('Retencion "' + sheet.getName() + '": ' + rows.length + ' filas eliminadas, quedan ' + keep.length +
    ' periodos [' + keep.join(', ') + ']');
}

/**
 * QUE HACE: borra filas por numero, agrupando las consecutivas y de abajo hacia
 *          arriba.
 * POR QUE ASI: borrar fila a fila es lentisimo, y borrar de arriba abajo
 *          desplaza los numeros de las filas siguientes y se borraria lo que no
 *          toca. De abajo arriba, los numeros de arriba no cambian.
 * EJEMPLO: filas [2,3,5,7] -> borra el bloque 2-3, luego la 5, luego la 7.
 */
function deleteRowsGrouped_(sheet, rows) {
  if (!rows || !rows.length) return;

  // slice() copia la lista para no modificar la original; sort compara de mayor
  // a menor (b - a) para ir de abajo arriba.
  var sorted = rows.slice().sort(function (a, b) { return b - a; });

  var i = 0;
  while (i < sorted.length) {
    var end = sorted[i];
    var start = end;

    // Se amplia el bloque mientras la fila siguiente sea justo la anterior.
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
// ORIGENES  -  leer la carpeta y deducir el periodo de cada fichero
// ============================================================================

/**
 * QUE HACE: recorre la carpeta y devuelve solo los ficheros cuyo nombre encaja
 *          con el patron, con su periodo ya deducido.
 * DEVUELVE: lista de objetos {id, name, periodo, anio, semana, orden,
 *          fechaReporte}, ordenada del periodo mas antiguo al mas reciente.
 * DETALLE: la extension se quita antes de comparar, asi que funciona igual con
 *          "Y26W48" que con "Y26W48.xlsx" o "Y26W48 copia".
 */
function listSourceFiles_(folder) {
  var re = new RegExp(pattern_());
  var it = folder.getFiles();
  var out = [];

  // Sirve para detectar el mismo periodo escrito de dos formas (Y26W5 / Y26W05).
  var porOrden = {};

  // getFiles() devuelve un "iterador": se recorre con hasNext() y next().
  while (it.hasNext()) {
    var file = it.next();

    // Quita la extension final, si la hay (punto + letras o numeros).
    var base = file.getName().replace(/\.[A-Za-z0-9]+$/, '');
    var m = base.match(re);

    // Si el nombre no encaja con el patron, se ignora sin avisar.
    if (!m) continue;

    // m[1] son las dos cifras del ano (26) y m[2] el numero de semana o de mes.
    var anoCorto = m[1];
    var num = Number(m[2]);

    // El periodo se guarda TAL CUAL se escribio en el nombre: Y26W5 se queda
    // como Y26W5, no se rellena a Y26W05. Asi el maestro usa tu nomenclatura.
    var periodo = 'Y' + anoCorto + (CONFIG.periodFormat === 'MONTH' ? 'M' : 'W') + m[2];

    // Ano completo: Y26 se interpreta como 2026.
    var year = 2000 + Number(anoCorto);

    // Clave numerica para ordenar y para la retencion. No se puede ordenar el
    // texto tal cual: por orden alfabetico "Y26W5" iria DESPUES de "Y26W48".
    var orden = year * 100 + num;

    // Dos ficheros con la misma clave son el mismo periodo escrito de dos
    // formas distintas: eso descuadraria la retencion, asi que se para aqui.
    if (porOrden[orden] && porOrden[orden] !== periodo) {
      throw new Error('dos ficheros representan el mismo periodo: "' + porOrden[orden] + '" y "' +
        periodo + '". Deja solo uno y unifica el formato.');
    }
    porOrden[orden] = periodo;

    out.push({
      id: file.getId(),
      name: file.getName(),
      periodo: periodo,
      anio: year,
      semana: (CONFIG.periodFormat === 'MONTH') ? '' : num,
      orden: orden,
      fechaReporte: fechaReporte_(periodo)
    });
  }

  // Orden por la clave numerica, que si es cronologico de verdad.
  out.sort(function (a, b) { return a.orden - b.orden; });
  return out;
}

/**
 * QUE HACE: devuelve el patron (expresion regular) que deben cumplir los
 *          nombres de fichero. Si CONFIG.filenamePattern esta relleno, se usa
 *          tal cual; si no, se deduce de periodFormat.
 * EJEMPLO SEMANAL: Y26W48 -> Y = ano, 26 = 2026, W = semana, 48 = semana 48.
 * DETALLE: \\d{1,2} acepta una o dos cifras, para que Y26W5 y Y26W05 valgan.
 */
function pattern_() {
  if (CONFIG.filenamePattern) return CONFIG.filenamePattern;
  return (CONFIG.periodFormat === 'MONTH') ? '^Y(\\d{2})M(\\d{1,2})$' : '^Y(\\d{2})W(\\d{1,2})$';
}

/**
 * QUE HACE: abre una hoja de Google y rechaza cualquier otro tipo de fichero.
 * POR QUE: si en la carpeta se cuela un PDF o una hoja de Excel con nombre de
 *          periodo, el mensaje de error dice exactamente que paso.
 */
function openSource_(fileId) {
  var file = DriveApp.getFileById(fileId);
  if (file.getMimeType() !== SHEETS_MIME) {
    throw new Error('"' + file.getName() + '" no es una hoja de Google (tipo: ' +
      file.getMimeType() + '). Solo se admiten hojas de Google.');
  }
  return SpreadsheetApp.openById(fileId);
}

/**
 * QUE HACE: busca una pestana por nombre dentro de una hoja y devuelve su nombre
 *          real. Tolera diferencias de mayusculas, espacios y tildes.
 * POR QUE: getSheetByName es estricto; asi un "Hoja 1" con un espacio de mas no
 *          rompe la ejecucion.
 * SI HAY DUDAS: falla y lista las pestanas disponibles, para que se vea al instante
 *          cual era el nombre correcto.
 */
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

/**
 * QUE HACE: convierte "ano + numero de semana ISO" en la fecha del lunes de esa
 *          semana.
 * POR QUE NO SE USA EL CALCULO FACIL: las semanas ISO no empiezan el 1 de enero;
 *          la semana 1 es la que contiene el 4 de enero. Por eso el calculo
 *          arranca en el 4 de enero, retrocede hasta su lunes y avanza semanas.
 * DETALLE: se usa mediodia UTC para que ningun cambio de hora mueva la fecha un
 *          dia arriba o abajo al guardarla en la hoja.
 */
function isoWeekToMonday_(year, week) {
  var jan4 = new Date(Date.UTC(year, 0, 4, 12, 0, 0));
  var dow = jan4.getUTCDay() || 7;
  var mondayW1 = new Date(jan4.getTime() - (dow - 1) * 86400000);
  return new Date(mondayW1.getTime() + (week - 1) * 7 * 86400000);
}

/**
 * QUE HACE: convierte un periodo ("Y26W47" o "Y26M11") en su fecha de
 *          referencia, que es la que se archiva en la columna fecha_reporte.
 *          Semanal -> el lunes de esa semana. Mensual -> el dia 1 del mes.
 * PARA QUE: Looker Studio necesita una fecha de verdad para ordenar y para los
 *          filtros de "ultimas N semanas/meses". La columna periodo es texto y
 *          no sirve para eso.
 */
function fechaReporte_(periodo) {
  if (CONFIG.periodFormat === 'MONTH') {
    var pm = String(periodo).match(/^Y(\d{2})M(\d{1,2})$/);
    if (!pm) {
      throw new Error('periodo mensual con formato inesperado: ' + periodo + ' (se esperaba algo como Y26M11)');
    }
    var mes = Number(pm[2]);
    if (!(mes >= 1 && mes <= 12)) {
      throw new Error('mes fuera de rango en el periodo: ' + periodo);
    }
    return new Date(Date.UTC(2000 + Number(pm[1]), mes - 1, 1, 12, 0, 0));
  }

  var pw = String(periodo).match(/^Y(\d{2})W(\d{1,2})$/);
  if (!pw) {
    throw new Error('periodo semanal con formato inesperado: ' + periodo + ' (se esperaba algo como Y26W48)');
  }
  var week = Number(pw[2]);

  // Una semana ISO solo puede ir de 1 a 53; fuera de eso, el fichero esta mal.
  if (!(week >= 1 && week <= 53)) {
    throw new Error('semana ISO fuera de rango: ' + periodo);
  }
  return isoWeekToMonday_(2000 + Number(pw[1]), week);
}

/**
 * QUE HACE: convierte un periodo en un numero comparable (Y26W48 -> 202648) y
 *          devuelve -1 si el texto no tiene el formato esperado.
 * PARA QUE: hace falta para ordenar. Comparando texto, "Y26W5" quedaria por
 *          detras de "Y26W48" porque el caracter "5" va despues del "4".
 */
function ordenDe_(periodo) {
  var m = String(periodo).match(/^Y(\d{2})[WM](\d{1,2})$/);
  if (!m) return -1;
  return (2000 + Number(m[1])) * 100 + Number(m[2]);
}

/**
 * QUE HACE: pasa una fecha a texto "ano-mes-dia" para poder mostrarla en el
 *          registro de ejecuciones.
 */
function formatDate_(d) {
  return Utilities.formatDate(d, CONFIG.timezone, 'yyyy-MM-dd');
}

// ============================================================================
// UTILIDADES
// ============================================================================

/**
 * QUE HACE: deja un texto en su forma minima para poder comparar nombres:
 *          minusculas, sin tildes, sin espacios de mas.
 * SE USA PARA: que "Hoja 1", "hoja 1 " y "HOJA 1" se consideren lo mismo.
 */
function normalize_(s) {
  return String(s)
    // NFD separa las tildes de las letras y esta linea las elimina.
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    // Varios espacios seguidos se convierten en uno solo.
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * QUE HACE: quita los elementos repetidos de una lista, conservando el orden.
 * TECNICA: un objeto hace de lista de "ya vistos"; seen[v] es true si ya salio.
 */
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

/**
 * QUE HACE: valida la cabecera de una pestana de origen antes de usarla.
 * COMPRUEBA: que ningun nombre este vacio, que no haya dos columnas llamadas
 *            igual, y que no se use un nombre reservado por el maestro
 *            (periodo, anio, semana, fecha_reporte, fecha_carga).
 * SI ALGO FALLA: lanza un error con el nombre de la pestana, para localizarlo.
 */
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

  // Recorre los nombres reservados y comprueba que el origen no use ninguno.
  META_HEADERS.concat([CACHE_HEADER]).forEach(function (reserved) {
    header.forEach(function (c) {
      if (normalize_(c) === reserved) {
        throw new Error('la pestana "' + sheetName + '" usa la columna "' + c + '", reservada por el maestro');
      }
    });
  });
}

/**
 * QUE HACE: busca celdas que contengan errores de formula (#REF!, #N/A...).
 * POR QUE: si una formula del origen esta rota, es mejor parar que archivar el
 *          error y arrastrarlo al maestro y a Looker.
 * DETALLE: solo se listan las 5 primeras para que el mensaje sea legible.
 */
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

/**
 * QUE HACE: manda un correo de aviso a CONFIG.notifyEmail.
 * DETALLE: si el correo fallara (por ejemplo, sin permiso), el error se apunta
 *          en el registro pero no se relanza, para no tapar el error original.
 */
function notify_(subject, body) {
  try {
    MailApp.sendEmail(CONFIG.notifyEmail, '[Historico] ' + subject, body);
  } catch (e) {
    Logger.log('No se pudo enviar el correo de aviso: ' + e.message);
  }
}
