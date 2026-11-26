/**
 * ============================================================================
 * HISTORICO MAESTRO PARA LOOKER STUDIO  -  Apps Script (V8)
 * ============================================================================
 *
 * QUE HACE ESTE SCRIPT
 *   Recorre una carpeta de Drive y TODAS sus subcarpetas (puedes tenerlas
 *   ordenadas por ano: 2025/, 2026/...). Dentro encuentra una hoja de Google por
 *   periodo, lee las pestanas de datos que le indiques y va copiando las filas a
 *   un unico fichero "maestro" con ID fijo. Ese maestro es el que lee Looker
 *   Studio, y al tener siempre el mismo ID nunca se rompe la conexion. Los
 *   periodos viejos se van borrando para que el fichero no crezca.
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
 *   - Fechas siempre en formato europeo dd/mm/aaaa, nunca anglosajon. El valor
 *     sigue siendo una fecha de verdad (para que Looker ordene y filtre); lo que
 *     se fija es como se muestra, aunque la hoja este en otro idioma.
 *   - #N/A se archiva tal cual: una formula sin resultados es un dato, no un
 *     fallo de ejecucion. Solo se informa en el resumen. Los errores de verdad
 *     (#REF!, #VALUE!, #DIV/0!...) si abortan.
 *
 * ----------------------------------------------------------------------------
 * PUESTA EN MARCHA (en este orden)
 * ----------------------------------------------------------------------------
 *   1. Rellena los IDs y el correo en CONFIG.
 *   2. Ejecuta listarPestanas(): te dice como se llaman las pestanas de tus
 *      ficheros y te escribe el bloque listo para pegar en CONFIG.entities.
 *   3. Ejecuta dryRun() y revisa el registro.
 *   4. Primera carga. Elige segun cuantos ficheros tengas:
 *        - ANOS de ficheros y solo quieres la ventana -> cargarUltimos()
 *          (carga los ultimos CONFIG.windowPeriods y nada mas)
 *        - pocos ficheros, quieres todos -> cargarNuevos()
 *        - un periodo concreto -> MANUAL_PERIOD + ingestManual()
 *      NO uses rebuildAll() con anos de ficheros: supera los 6 minutos y ademas
 *      la retencion borraria casi todo lo cargado.
 *   5. Uso normal, SIN trigger: guarda el fichero de la semana o del mes y
 *      ejecuta cargarNuevos(). Eso es todo, cada vez.
 *   6. Comparte el maestro SOLO EN LECTURA. Nadie debe editar ni formatear ese
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

  // Carpeta de Drive donde estan los ficheros de periodo. Se leen tambien
  // TODAS sus subcarpetas, a cualquier profundidad: puedes tenerlos ordenados
  // por ano (2025/, 2026/...) y el script los encuentra igual.
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

  // Patron que deben cumplir los nombres de fichero. Por defecto vale cualquier
  // texto delante y exige que el nombre TERMINE con el periodo, porque es lo
  // habitual en tus ficheros: "fichero pruebas - Y26W48".
  // Ejemplos para Y26W48 / Y26M11:
  //   'Y(\\d{2})W(\\d{1,2})$'   (semanal: el periodo al final)
  //   'Y(\\d{2})M(\\d{1,2})$'   (mensual: el periodo al final)
  // Si el nombre fuera SOLO el periodo, se pondria con ^ delante:
  //   '^Y(\\d{2})W(\\d{1,2})$'
  // Los dobles backslash son obligatorios: es la forma de escribir \d en texto.
  // \\d{1,2} acepta tanto Y26W5 como Y26W05.
  filenamePattern: '',

  // Fila donde empiezan las cabeceras en las pestanas de origen. Casi siempre 1.
  headerRow: 1,

  // Si es true, la ejecucion se corta cuando encuentra errores de formula de
  // verdad (#REF!, #VALUE!, #DIV/0!...). #N/A NO cuenta como error: significa
  // "la formula no encontro resultados" y se archiva tal cual, avisando en el
  // resumen. Dejalo en true para no archivar formulas rotas.
  failOnErrorCells: true,

  // ESTA ES LA LISTA DE PESTANAS QUE SE HISTORIZAN. Una linea por pestana.
  // Si aqui no esta, no se copia. Las tablas, resumenes y portadas del fichero
  // se quedan fuera simplemente no poniendolas.
  //
  //   origen   : nombre de la pestana DENTRO del fichero de periodo. Tiene que
  //              ser el nombre real; si no lo sabes, ejecuta listarPestanas() y
  //              copia lo que te escriba en el registro.
  //   destino  : nombre de la pestana en el maestro (el que veras en Looker).
  //              Puede ser distinto del origen; si no existe, se crea.
  //   headerRow: opcional. Solo si esa pestana tiene las cabeceras en otra fila
  //              distinta de CONFIG.headerRow.
  //
  // El valor de abajo es un ejemplo que HAY QUE SUSTITUIR por las pestanas
  // reales. Si lo dejas tal cual, la ejecucion fallara diciendo que no encuentra
  // esa pestana, que es justo lo que queremos: mejor fallar claro que copiar
  // datos de un sitio que no toca.
  entities: [
    { destino: 'ventas', origen: 'PON_AQUI_EL_NOMBRE_DE_LA_PESTANA' }
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

// Formato de fecha OBLIGATORIO en el maestro: dia/mes/ano, nunca mes/dia/ano.
// Se aplica como formato de celda, asi que se respeta aunque la hoja este
// configurada en ingles. El valor guardado sigue siendo una fecha real.
var FORMATO_FECHA = 'dd/MM/yyyy';
var FORMATO_FECHA_HORA = 'dd/MM/yyyy HH:mm:ss';

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
 * QUE HACE: EL BOTON DEL DIA A DIA. Recorre la carpeta y sus subcarpetas y carga
 *           en el maestro solo los periodos que todavia no estan dentro. Si un
 *           periodo ya esta, lo salta sin abrir siquiera el fichero.
 * CUANDO SE EJECUTA: a mano, justo despues de guardar el fichero de la semana o
 *           del mes. Tambien es lo que lanzaria el trigger automatico si algun
 *           dia lo instalas, pero no hace falta: pensado para lanzarlo tu.
 * NO HACE: no rehace periodos ya cargados. Para corregir uno ya cargado esta
 *           MANUAL_PERIOD + ingestManual().
 * OJO: la primera vez, si la carpeta tiene anos de ficheros, carga TODOS los que
 *           falten y puede pasarse de los 6 minutos. Solo para esa primera vez
 *           usa cargarUltimos(), que trae unicamente la ventana.
 */
function cargarNuevos() {
  // execute_ es el envoltorio que pone el candado y manda correo si algo falla.
  // Le pasamos una funcion anonima: el trabajo real que debe ejecutar.
  execute_(function () {
    return ingest_({ onlyMissing: true });
  });
}

/**
 * QUE HACE: reconstruccion completa. Vuelve a procesar TODOS los ficheros de la
 *           carpeta, reescribiendo cada periodo que encuentre.
 * CUANDO SE EJECUTA: solo si sospechas que hay periodos mal cargados y quieres
 *           rehacerlos desde los ficheros. NO es la primera carga.
 * OJO: con muchos periodos supera el limite de 6 minutos por ejecucion. Con
 *      anos de ficheros usa cargarUltimos() para la primera carga, o hazlo por
 *      tandas con ingestPeriod('Y26W40') y siguientes.
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
  // Si se pulsa desde el editor sin argumento, el valor llega vacio. Mejor un
  // mensaje util que un error raro buscando un periodo llamado "undefined".
  if (!periodo || typeof periodo !== 'string') {
    throw new Error('ingestPeriod necesita un periodo, por ejemplo ingestPeriod("Y26W48"). ' +
      'Para lanzarlo a mano desde el editor usa ingestManual() y edita MANUAL_PERIOD.');
  }
  execute_(function () {
    return ingest_({ onlyMissing: false, onlyPeriod: String(periodo) });
  });
}

/**
 * QUE HACE: te dice como se llaman las pestanas que tiene el fichero de periodo
 *          MAS RECIENTE de la carpeta, cuantas filas y columnas trae cada una,
 *          y si ya esta configurada en CONFIG.entities. Ademas escribe el
 *          bloque listo para pegar en CONFIG.entities.
 * CUANDO SE EJECUTA: antes de rellenar CONFIG.entities, para no adivinar
 *          nombres. No escribe nada en ningun sitio.
 * RESULTADO: Ver > Registro de ejecuciones (o Ctrl+Enter).
 */
function listarPestanas() {
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var files = listSourceFiles_(folder);
  if (!files.length) {
    throw new Error('No hay ficheros de periodo en la carpeta. Patron usado: ' + pattern_());
  }

  // El ultimo de la lista es el periodo mas reciente: el que mejor representa
  // como son los ficheros hoy.
  var ultimo = files[files.length - 1];
  var ss = openSource_(ultimo.id);
  var lines = [];

  lines.push('Fichero mas reciente: ' + rutaDe_(ultimo) + '  (periodo ' + ultimo.periodo + ')');
  lines.push('Pestanas que tiene:');

  // Mapa de pestanas ya configuradas, en forma normalizada para comparar igual
  // que el resto del script (sin tildes, sin espacios de mas, en minusculas).
  var configuradas = {};
  CONFIG.entities.forEach(function (e) {
    configuradas[normalize_(e.origen)] = e.destino;
  });

  var candidatas = [];
  ss.getSheets().forEach(function (s) {
    var filas = Math.max(s.getLastRow() - 1, 0);
    var cols = s.getLastColumn();
    var destino = configuradas[normalize_(s.getName())];
    var marca = destino ? ('  -> se guarda en "' + destino + '"') : '  -> NO configurada';
    lines.push('  - "' + s.getName() + '": ' + filas + ' filas de datos, ' + cols + ' columnas' + marca);

    // Solo se proponen las que tienen datos: las vacias no aportan nada.
    if (filas > 0) {
      candidatas.push('    { destino: "' + s.getName() + '", origen: "' + s.getName() + '" }');
    }
  });

  lines.push('');
  lines.push('Para historizar alguna, pega este bloque dentro de CONFIG.entities');
  lines.push('y borra las lineas de las pestanas que no quieras:');
  lines.push('  entities: [');
  lines.push(candidatas.join(',\n'));
  lines.push('  ]');

  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * QUE HACE: foto detallada de como ve el script cada pestana de origen en el
 *          fichero mas reciente, y del estado de la pestana equivalente en el
 *          maestro. No escribe nada en ningun sitio.
 * CUANDO SE EJECUTA: cuando algo no cuadra. Contesta preguntas como "por que
 *          solo entra una fila", "cuantas celdas tienen valor en cada columna"
 *          o "por que el maestro tiene miles de filas vacias".
 * RESULTADO: Ver > Registro de ejecuciones. Es el texto que hay que revisar
 *          (o enviar) cuando una pestana se comporta raro.
 */
function diagnosticoPestana() {
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var files = listSourceFiles_(folder);
  if (!files.length) {
    throw new Error('No hay ficheros de periodo en la carpeta. Patron usado: ' + pattern_());
  }

  var ultimo = files[files.length - 1];
  var ss = openSource_(ultimo.id);
  var master = SpreadsheetApp.openById(CONFIG.masterFileId);

  var lines = [];
  lines.push('Fichero analizado: ' + rutaDe_(ultimo) + '  (periodo ' + ultimo.periodo + ')');

  CONFIG.entities.forEach(function (e) {
    lines.push('');
    lines.push('=== destino "' + e.destino + '"   <-   pestana "' + e.origen + '" ===');

    // --- Como esta la pestana en el maestro ---
    var m = master.getSheetByName(e.destino);
    if (!m) {
      lines.push('  maestro: la pestana no existe todavia');
    } else {
      var conPeriodo = 0;
      try {
        conPeriodo = readPeriodColumn_(m).filter(function (p) { return p !== ''; }).length;
      } catch (err) {
        lines.push('  maestro: PROBLEMA -> ' + err.message);
      }
      lines.push('  maestro: getLastRow=' + m.getLastRow() + ' | filas con periodo=' + conPeriodo);
      if (m.getLastRow() - 1 > conPeriodo) {
        lines.push('  OJO: ' + (m.getLastRow() - 1 - conPeriodo) +
          ' filas las cuenta getLastRow pero no son datos (formato aplicado en vacio).');
      }
    }

    // --- Como esta la pestana en el origen ---
    var nombre;
    try {
      nombre = resolveSheetName_(ss, e.origen);
    } catch (err) {
      lines.push('  origen: ' + err.message);
      return;
    }

    var sh = ss.getSheetByName(nombre);
    var headerRow = e.headerRow || CONFIG.headerRow;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    lines.push('  origen: getLastRow=' + lastRow + ' | getLastColumn=' + lastCol +
      ' | cabecera en fila=' + headerRow);

    if (lastRow <= headerRow || lastCol === 0) {
      lines.push('  origen: no hay datos debajo de la cabecera');
      return;
    }

    var header = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0]
      .map(function (v) { return String(v).trim(); });

    lines.push('  cabecera: ' + header.map(function (h, i) {
      return (i + 1) + '="' + h + '"';
    }).join(' | '));

    var vacias = header.filter(function (h) { return h === ''; }).length;
    if (vacias) {
      lines.push('  OJO: ' + vacias + ' cabeceras vacias. Casi siempre son celdas combinadas:');
      lines.push('       el script solo puede leer el valor de la primera celda del bloque.');
    }

    var nFilas = lastRow - headerRow;
    var datos = sh.getRange(headerRow + 1, 1, nFilas, lastCol).getValues();
    var conAlgo = datos.filter(function (r) {
      return r.some(function (c) { return c !== '' && c !== null; });
    }).length;
    lines.push('  filas de datos: ' + nFilas + ' | con algun valor: ' + conAlgo);

    // Cuantas celdas con valor hay en cada columna. Si una columna tiene 1 y
    // deberia tener 500, ahi esta el problema (celdas combinadas o datos
    // colocados en otro sitio).
    var conteo = [];
    for (var c = 0; c < lastCol; c++) {
      var n = 0;
      for (var f = 0; f < datos.length; f++) {
        if (datos[f][c] !== '' && datos[f][c] !== null) n++;
      }
      conteo.push((c + 1) + '=' + n);
    }
    lines.push('  celdas con valor por columna: ' + conteo.join(' | '));

    lines.push('  primeras 3 filas de datos:');
    datos.slice(0, 3).forEach(function (r) {
      lines.push('    ' + r.map(function (v) {
        var t = (v === '' || v === null) ? '(vacio)' : String(v);
        return t.length > 22 ? t.slice(0, 22) + '...' : t;
      }).join(' | '));
    });
  });

  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * QUE HACE: carga en el maestro solo los ultimos N periodos que hay en la
 *          carpeta, del mas antiguo al mas reciente.
 * CUANDO SE EJECUTA: la primera vez, cuando la carpeta tiene anos de ficheros y
 *          solo quieres la ventana. Con 150 ficheros semanales, cargarlos todos
 *          para que la retencion borre 138 es tiempo tirado, y ademas se supera
 *          el limite de 6 minutos por ejecucion.
 * RECIBE: n. Si lo llamas sin argumento, usa CONFIG.windowPeriods (12).
 * EJEMPLO: cargarUltimos() carga 12 semanas; cargarUltimos(26), medio ano.
 */
function cargarUltimos(n) {
  // Sin argumento se usa la ventana configurada. Con argumento, se respeta:
  // cargarUltimos(0) es un error, no "usa la ventana".
  var cuantos = (n === undefined || n === null || n === '') ? CONFIG.windowPeriods : Number(n);
  if (!(cuantos >= 1)) {
    throw new Error('cargarUltimos necesita un numero de periodos mayor que 0.');
  }

  return execute_(function () {
    var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
    var files = listSourceFiles_(folder);
    if (!files.length) {
      throw new Error('No hay ficheros de periodo en la carpeta. Patron usado: ' + pattern_());
    }

    // files ya viene ordenado de mas antiguo a mas reciente.
    var ultimos = files.slice(Math.max(files.length - cuantos, 0)).map(function (f) {
      return f.periodo;
    });
    Logger.log('Se van a cargar ' + ultimos.length + ' periodos: ' + ultimos.join(', '));
    return ingest_({ onlyMissing: true, onlyPeriods: ultimos });
  });
}

/**
 * QUE HACE: radiografia de las columnas del maestro. Para cada pestana dice
 *          cuantas celdas con valor tiene cada columna, y avisa cuando dos
 *          columnas parecen la misma cosa partida en dos.
 * CUANDO SE EJECUTA: cuando sospechas que un renombrado en el origen dejo datos
 *          repartidos entre la columna vieja y la nueva. Es la forma de verlo.
 * COMO SE LEE: si "importe" tiene 26.000 valores y "importe_total" otros 26.000,
 *          y el total de filas es 52.000, es un renombrado: la mitad de los
 *          periodos estan en una columna y la otra mitad en la otra.
 * RESULTADO: Ver > Registro de ejecuciones.
 */
function diagnosticoColumnas() {
  var master = SpreadsheetApp.openById(CONFIG.masterFileId);
  var lines = ['Columnas del maestro y celdas con valor en cada una:'];

  CONFIG.entities.forEach(function (e) {
    lines.push('');
    var sh = master.getSheetByName(e.destino);
    if (!sh) {
      lines.push('=== "' + e.destino + '": la pestana no existe');
      return;
    }

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    lines.push('=== "' + e.destino + '": ' + Math.max(lastRow - 1, 0) + ' filas x ' + lastCol + ' columnas');
    if (lastRow < 2 || lastCol === 0) return;

    var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var header = vals.shift().map(function (v) { return String(v).trim(); });

    var conteo = [];
    for (var c = 0; c < lastCol; c++) conteo.push(0);
    vals.forEach(function (row) {
      for (var j = 0; j < lastCol; j++) {
        if (row[j] !== '' && row[j] !== null) conteo[j]++;
      }
    });

    header.forEach(function (h, i) {
      lines.push('  ' + (i + 1) + '. ' + (h || '(sin nombre)') + ': ' + conteo[i] + ' celdas con valor');
    });

    // Heuristica de renombrado: dos columnas de datos cuyos recuentos suman
    // aproximadamente el total de filas. Si ambas estuvieran completas sumarian
    // el doble, asi que esa suma solo cuadra cuando cada una tiene la mitad.
    var total = vals.length;
    if (total < 2) return;

    for (var a = META_HEADERS.length; a < lastCol - 1; a++) {
      for (var b = a + 1; b < lastCol - 1; b++) {
        var suma = conteo[a] + conteo[b];
        if (conteo[a] > 0 && conteo[b] > 0 && Math.abs(suma - total) <= Math.max(2, total * 0.02)) {
          lines.push('  OJO: "' + header[a] + '" y "' + header[b] + '" suman ' + suma + ' de ' + total +
            ' filas. Huele a renombrado: los datos estan partidos entre las dos.');
        }
      }
    }
  });

  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * QUE HACE: recorre TODOS los ficheros de periodo y compara las columnas de cada
 *          pestana con el esquema habitual (el que tiene la mayoria). Devuelve
 *          la lista de ficheros que se desvian y en que columna exactamente.
 * PARA QUE: si el maestro tiene datos partidos entre dos columnas, esto dice
 *          QUE FICHEROS hay que corregir en el origen, no solo que hay un lio.
 * COMO SE LEE:
 *          "esquema habitual (147 ficheros): cliente, importe, canal"
 *          "desviaciones: 3"
 *          "  - 2024/f - Y24W31: falta: canal | sobra: canal_venta"
 *          Ese fichero tiene la columna renombrada: se corrige en el origen.
 * TARDA: abre cada fichero, asi que con 150 ficheros puede irse a un par de
 *          minutos. Es de solo lectura: no escribe nada.
 * RESULTADO: Ver > Registro de ejecuciones.
 */
function diagnosticoFicheros() {
  var folder = DriveApp.getFolderById(CONFIG.sourceFolderId);
  var files = listSourceFiles_(folder);
  if (!files.length) {
    throw new Error('No hay ficheros de periodo en la carpeta. Patron usado: ' + pattern_());
  }

  var master = SpreadsheetApp.openById(CONFIG.masterFileId);

  // 1. Se abre cada fichero UNA vez y se guardan las cabeceras de cada entidad.
  var recogidas = {};
  var erroresApertura = {};
  CONFIG.entities.forEach(function (e) { recogidas[e.destino] = []; });

  files.forEach(function (f) {
    var ss = null;
    var fallo = null;
    try {
      ss = openSource_(f.id);
    } catch (err) {
      fallo = err.message;
    }

    CONFIG.entities.forEach(function (e) {
      if (!ss) {
        erroresApertura[e.destino] = erroresApertura[e.destino] || [];
        erroresApertura[e.destino].push(rutaDe_(f) + ': ' + fallo);
        return;
      }
      try {
        var sh = ss.getSheetByName(resolveSheetName_(ss, e.origen));
        var hr = e.headerRow || CONFIG.headerRow;
        if (sh.getLastRow() < hr || sh.getLastColumn() === 0) {
          recogidas[e.destino].push({ ruta: rutaDe_(f), error: 'la pestana "' + e.origen + '" no tiene datos' });
          return;
        }
        var h = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0]
          .map(function (v) { return String(v).trim(); })
          .filter(function (v) { return v !== ''; });
        recogidas[e.destino].push({ ruta: rutaDe_(f), header: h });
      } catch (err) {
        recogidas[e.destino].push({ ruta: rutaDe_(f), error: err.message });
      }
    });
  });

  // 2. Por entidad: esquema mas frecuente y ficheros que se desvian de el.
  var lines = ['Comparacion de las columnas de cada fichero con el esquema habitual.'];
  lines.push('Ficheros de periodo encontrados: ' + files.length);

  CONFIG.entities.forEach(function (e) {
    lines.push('');
    lines.push('=== destino "' + e.destino + '"  <-  pestana "' + e.origen + '" ===');

    var lista = recogidas[e.destino];
    var cuenta = {};
    lista.forEach(function (x) {
      if (!x.header) return;
      var clave = x.header.join(' | ');
      cuenta[clave] = (cuenta[clave] || 0) + 1;
    });

    // El esquema habitual es el que mas se repite. Se compara con el, no con el
    // maestro: si el maestro ya tiene una columna fantasma, compararlo con el
    // marcaria como raros los 147 ficheros buenos.
    var habitual = null;
    var veces = 0;
    Object.keys(cuenta).forEach(function (k) {
      if (cuenta[k] > veces) { veces = cuenta[k]; habitual = k; }
    });

    if (!habitual) {
      lines.push('  no se ha podido leer ninguna cabecera');
      return;
    }

    var columnasHabituales = habitual === '' ? [] : habitual.split(' | ');
    lines.push('  esquema habitual (' + veces + ' ficheros): ' + habitual);

    // Referencia: lo que tiene el maestro ahora mismo.
    var shMaster = master.getSheetByName(e.destino);
    if (shMaster) {
      var hMaster = readHeader_(shMaster);
      var hm = hMaster.slice(META_HEADERS.length, hMaster.length - 1);
      lines.push('  el maestro tiene: ' + (hm.join(' | ') || '(ninguna)'));
      var sobranEnMaestro = hm.filter(function (c) { return columnasHabituales.indexOf(c) === -1; });
      if (sobranEnMaestro.length) {
        lines.push('  OJO: el maestro conserva columnas que el origen ya no trae: ' + sobranEnMaestro.join(', ') +
          '. Esas son las que hay que borrar del maestro.');
      }
    } else {
      lines.push('  el maestro todavia no tiene esa pestana');
    }

    // Desviaciones: ficheros cuyo juego de columnas no es el habitual.
    var desvios = lista.filter(function (x) {
      if (x.error) return true;
      return x.header.join(' | ') !== habitual;
    });

    lines.push('  desviaciones: ' + desvios.length + ' de ' + lista.length);
    desvios.sort(function (a, b) { return a.ruta < b.ruta ? -1 : 1; });
    desvios.forEach(function (x) {
      if (x.error) {
        lines.push('  - ' + x.ruta + ': ' + x.error);
        return;
      }
      var faltan = columnasHabituales.filter(function (c) { return x.header.indexOf(c) === -1; });
      var sobran = x.header.filter(function (c) { return columnasHabituales.indexOf(c) === -1; });
      var detalle = [];
      if (faltan.length) detalle.push('falta: ' + faltan.join(', '));
      if (sobran.length) detalle.push('sobra: ' + sobran.join(', '));
      lines.push('  - ' + x.ruta + ': ' + detalle.join(' | '));
    });

    var fallos = erroresApertura[e.destino] || [];
    if (fallos.length) {
      lines.push('  no se pudieron abrir: ' + fallos.length);
      fallos.slice(0, 10).forEach(function (m) { lines.push('  - ' + m); });
    }
  });

  lines.push('');
  lines.push('Cada linea "- ..." es un fichero a corregir en el origen.');

  var out = lines.join('\n');
  Logger.log(out);
  return out;
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
    lines.push('  - ' + rutaDe_(f) + '  ->  ' + f.periodo + '  (fecha_reporte ' + formatDate_(fechaReporte_(f.periodo)) + ')');
  });

  // Aviso util: se repasa la carpeta y se separan los ficheros que NO encajan
  // con el patron. Los que parecen de periodo pero son del otro tipo (un
  // mensual Y26M11 en un script semanal) se nombran aparte, porque al cargar se
  // ignoran en silencio y pareceria que el script no los ve.
  var letraEsperada = (CONFIG.periodFormat === 'MONTH') ? 'M' : 'W';
  var parecePeriodo = /Y\d{2}([WM])\d{1,2}$/i;
  var rePatron = new RegExp(pattern_(), 'i');
  var otroTipo = [];
  var noEncajan = [];
  todosLosFicheros_(folder).forEach(function (item) {
    var nombreOtro = item.name;
    var baseOtro = nombreOtro.replace(/\.[A-Za-z0-9]+$/, '');
    if (baseOtro.match(rePatron) || nombreOtro.match(rePatron)) return;

    var coincidencia = baseOtro.match(parecePeriodo) || nombreOtro.match(parecePeriodo);
    if (coincidencia) {
      otroTipo.push(rutaDe_(item) + ' (tipo ' + coincidencia[1].toUpperCase() + ', se esperaba ' + letraEsperada + ')');
    } else if (noEncajan.length < 5) {
      noEncajan.push(rutaDe_(item));
    }
  });
  if (otroTipo.length) {
    lines.push('OJO, ignorados por ser del otro tipo: ' + otroTipo.join(' | '));
  }
  if (noEncajan.length) {
    lines.push('No encajan con el patron (se ignoran): ' + noEncajan.join(' | '));
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

  // Comprobacion de que las pestanas configuradas existen de verdad en el
  // fichero mas reciente: asi se descubre aqui y no a mitad de la carga.
  if (files.length) {
    var ultimoFichero = files[files.length - 1];
    var fuente = openSource_(ultimoFichero.id);
    lines.push('');
    lines.push('Pestanas configuradas, comprobadas en "' + rutaDe_(ultimoFichero) + '":');
    CONFIG.entities.forEach(function (e) {
      try {
        var nombreReal = resolveSheetName_(fuente, e.origen);
        lines.push('  OK  "' + nombreReal + '"  ->  se guarda en "' + e.destino + '"');
      } catch (err) {
        lines.push('  FALTA  "' + e.origen + '"  ->  ' + err.message);
      }
    });
  }

  // join('\n') pega todas las lineas separandolas con un salto de linea.
  var out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * OPCIONAL. QUE HACE: deja programada la ejecucion automatica (el "trigger").
 * SI lanzas el script a mano, NO ejecutes esto: no hace falta y ademas puede
 *          adelantarse a que los datos esten cerrados.
 * CUANDO SE EJECUTA: una sola vez, si algun dia quieres que se cargue solo.
 * DETALLE: lunes a las 06:00 si es semanal; dia 1 a las 06:00 si es mensual.
 *          Antes de crear uno nuevo borra el anterior, para no duplicar.
 */
function installTrigger() {
  removeTriggers();

  // Se va montando el trigger por partes: primero "que funcion llamar" y con
  // que zona horaria; luego la frecuencia.
  var builder = ScriptApp.newTrigger('cargarNuevos').timeBased().inTimezone(CONFIG.timezone);
  if (CONFIG.periodFormat === 'MONTH') {
    builder.onMonthDay(1).atHour(6).create();
  } else {
    builder.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  }

  Logger.log('Trigger instalado.');
}

/**
 * QUE HACE: borra el trigger automatico de este script, si existe.
 * CUANDO SE EJECUTA: si quieres parar la automatizacion sin borrar el proyecto,
 *          o para dejar limpio un trigger antiguo que ya no quieras.
 */
function removeTriggers() {
  // getProjectTriggers() devuelve todos los triggers del proyecto. Solo se
  // borran los de este script; el resto se deja intacto. Se mira tambien el
  // nombre antiguo ("ingestScheduled") por si instalaste un trigger con una
  // version anterior: si no, se quedaria ahi fallando cada semana.
  var nombres = ['cargarNuevos', 'ingestScheduled'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (nombres.indexOf(t.getHandlerFunction()) !== -1) {
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
    // Cuando no encaja ningun fichero, el mensaje muestra lo que si hay en la
    // carpeta (y sus subcarpetas): casi siempre es un nombre mal escrito o un
    // ajuste del patron.
    var muestra = todosLosFicheros_(folder).slice(0, 5).map(rutaDe_);
    throw new Error('No hay ficheros que cumplan el patron de nombre. Patron usado: ' + pattern_() +
      (muestra.length ? '. En la carpeta hay: ' + muestra.join(' | ') : '. La carpeta esta vacia o no es accesible.') +
      '. Revisa CONFIG.filenamePattern.');
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

  // Si se pide una lista de periodos (lo usa cargarUltimos), solo esos.
  if (opts.onlyPeriods) {
    files = allFiles.filter(function (f) { return opts.onlyPeriods.indexOf(f.periodo) !== -1; });
    if (!files.length) {
      throw new Error('Ningun fichero encaja con los periodos pedidos: ' + opts.onlyPeriods.join(', '));
    }
  }

  // Primero se valida (o se crea) cada pestana destino del maestro. Se hace
  // antes de leer nada para que cualquier problema de estructura salga aqui,
  // con un mensaje que diga que pasa, y no como un error suelto mas adelante.
  CONFIG.entities.forEach(function (e) {
    ensureMasterTab_(master, e);
  });

  // Estado de cada pestana destino: que periodos tiene, en que filas, y cual es
  // la ultima fila de datos. Se lee UNA vez al principio y se va actualizando en
  // memoria. Antes se releia la columna de periodo en cada fichero, y con 52
  // semanas eso crece hasta hacerse cuadratico: por eso se agotaba el tiempo.
  var estados = {};
  CONFIG.entities.forEach(function (e) {
    estados[e.destino] = crearEstadoEntidad_(master.getSheetByName(e.destino));
  });

  // Aqui se va apuntando lo que pasa, para el resumen final y el correo.
  var summary = { ok: [], skipped: [], failed: [], warned: [] };

  // ---- Bucle principal: un fichero de periodo cada vez ----
  files.forEach(function (entry) {
    // De las entidades configuradas, cuales hay que cargar para este fichero.
    // Con onlyMissing, se descartan las que ya tienen ese periodo.
    var pending = CONFIG.entities.filter(function (e) {
      return !(opts.onlyMissing && estados[e.destino].filas[entry.periodo]);
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
        estados[e.destino] = ingestEntity_(master, srcSs, e, entry, entry.periodo === latestPeriod,
          summary.warned, estados[e.destino]);
        summary.ok.push(e.destino + ' @ ' + entry.periodo);
      } catch (err) {
        summary.failed.push(entry.periodo + ' / ' + e.destino + ': ' + err.message);
      }
    });
  });

  // Retencion al final, una vez por entidad: aqui se borran los periodos viejos.
  // Despues se fija el formato europeo de las columnas de fecha.
  CONFIG.entities.forEach(function (e) {
    try {
      applyRetention_(master, e);
      formatearFechas_(master.getSheetByName(e.destino));
    } catch (err) {
      summary.failed.push('retencion/formatos / ' + e.destino + ': ' + err.message);
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
 * QUE HACE: lee una vez la columna de periodo de una pestana del maestro y
 *          devuelve un resumen en memoria: que periodos hay, en que filas y
 *          cual es la ultima fila de datos.
 * PARA QUE: asi el script NO tiene que releer la columna por cada fichero que
 *          carga. Con 52 semanas esa relectura crecia hasta agotar el tiempo.
 * DEVUELVE: { filas: { "Y26W48": [1201, 1202, ...] }, ultima: 1202 }
 */
function crearEstadoEntidad_(sheet) {
  var estado = { filas: {}, ultima: 1 };
  if (!sheet || sheet.getLastRow() < 2) return estado;

  var col = periodoColumn_(sheet);
  var vals = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();

  for (var i = 0; i < vals.length; i++) {
    var p = String(vals[i][0]).trim();
    if (!p) continue;
    if (!estado.filas[p]) estado.filas[p] = [];
    estado.filas[p].push(i + 2);
    estado.ultima = i + 2;
  }
  return estado;
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
 *   estado    : resumen en memoria de la pestana destino (ver crearEstadoEntidad_).
 * DEVUELVE: el estado actualizado, para no volver a leer la columna.
 * PASOS: preparar pestana destino, leer origen, validar, alinear columnas,
 *        borrar el periodo si ya estaba y escribir de una sola vez.
 */
function ingestEntity_(master, srcSs, entity, entry, isLatest, warnings, estado) {
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

  // 4. Celdas de error: #N/A se archiva (solo avisa), el resto de errores para.
  if (CONFIG.failOnErrorCells) {
    assertNoErrors_(block, entity.origen, headerRow, warnings, entity.destino, entry.periodo);
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

  // Upsert sin releer la hoja:
  //  - si el periodo ya estaba, se borran sus filas y se recalcula el estado,
  //    porque los indices de las filas de debajo cambian;
  //  - si no estaba, se escribe al final de los datos, sin ninguna lectura.
  // La fila de destino se calcula por la ultima fila con periodo, NO por
  // getLastRow(): si la pestana trae formato aplicado muy abajo, getLastRow()
  // devuelve miles de filas y el script escribiria al final del vacio.
  if (estado.filas[entry.periodo]) {
    deleteRowsGrouped_(masterSheet, estado.filas[entry.periodo]);
    estado = crearEstadoEntidad_(masterSheet);
  }

  var startRow = estado.ultima + 1;
  escribirFilas_(masterSheet, startRow, payload, warnings, entity.destino, entry.periodo);

  // El estado queda al dia para el siguiente fichero: cero lecturas extra.
  estado.ultima = startRow + payload.length - 1;
  estado.filas[entry.periodo] = [];
  for (var r = startRow; r <= estado.ultima; r++) estado.filas[entry.periodo].push(r);
  return estado;
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
      throw new Error('la pestana "' + entity.destino + '" ya existe en el maestro pero no tiene el formato del maestro: ' +
        'su columna ' + (i + 1) + ' es "' + (current[i] || 'vacia') + '" y deberia ser "' + META_HEADERS[i] + '". ' +
        'Esas columnas las genera el script a partir de la fecha del fichero, no tienen que venir en el origen. ' +
        'Revisa que masterFileId apunte al fichero maestro y no a uno de los ficheros de periodo, ' +
        'y que el "destino" de esta entidad no choque con una pestana que ya existia. ' +
        'Si esa pestana sobra, renombrala o borrala.');
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

    // Se avisa en el resumen, no solo en el registro: una columna nueva en un
    // fichero antiguo casi siempre es un renombrado hecho sin querer, y deja los
    // datos partidos entre la columna vieja y la nueva.
    if (warnings) {
      warnings.push(destino + ' @ ' + periodo + ': columnas nuevas anadidas -> ' + newCols.join(', ') +
        (isLatest ? '' : ' (ojo: viene de un fichero que NO es el mas reciente, puede ser un renombrado)'));
    }
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
    throw new Error('la pestana "' + sheet.getName() + '" del maestro no tiene columna "periodo". ' +
      'Esa columna la genera el script a partir de la fecha del fichero; si esa pestana no la creo el script, ' +
      'renombrala o cambia el "destino" de la entidad.');
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
 * QUE HACE: devuelve el numero de la ultima fila que tiene periodo escrito, es
 *          decir, la ultima fila que el script reconocera como dato.
 * POR QUE NO getLastRow(): getLastRow() cuenta tambien las filas que solo tienen
 *          FORMATO (color, bordes, ancho fijado). Si alguien dejo formato
 *          aplicado hasta la fila 5000, getLastRow() devuelve 5000 y el script
 *          escribiria en la 5001, dejando miles de filas vacias en medio.
 */
function ultimaFilaDeDatos_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 1;

  var col = periodoColumn_(sheet);
  var vals = sheet.getRange(2, col, last - 1, 1).getValues();

  // Se recorre de abajo arriba: la primera fila con algo es la ultima de datos.
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

/**
 * QUE HACE: escribe las filas en el maestro.
 * SI LA HOJA SE QUEJA DE "data validation": la pestana destino trae una regla
 *   de validacion con "rechazar entrada" (listas, casillas, numeros con rango)
 *   que bloquea la escritura. En una pestana que escribe el script esa regla no
 *   aporta nada y rompe la carga, asi que se quita solo del rango que se va a
 *   escribir y se reintenta. Queda anotado como aviso en el resumen.
 */
function escribirFilas_(sheet, startRow, payload, warnings, destino, periodo) {
  var rango = sheet.getRange(startRow, 1, payload.length, payload[0].length);

  try {
    rango.setValues(payload);
    return;
  } catch (err) {
    Logger.log('La escritura fallo (posible validacion de datos). Se quita la validacion del rango y se reintenta. ' +
      'Detalle: ' + err.message);
  }

  try {
    rango.setDataValidation(null);
    rango.setValues(payload);
  } catch (err2) {
    throw new Error('no se pudo escribir en la pestana "' + sheet.getName() + '": ' + err2.message);
  }

  if (warnings) {
    warnings.push(destino + ' @ ' + periodo +
      ': habia validacion de datos en el destino y se ha quitado para poder escribir');
  }
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
 * QUE HACE: fija el formato europeo (dia/mes/ano) de las columnas de fecha del
 *          maestro.
 * POR QUE: el valor guardado es una fecha real y no cambia; lo que se fija es
 *          como se pinta. Sin esto, una hoja configurada en ingles mostraria
 *          mes/dia/ano, y el requisito es que NUNCA salga el formato anglosajon.
 * CUANDO: al final de cada ejecucion, sobre el rango que tiene datos (nunca
 *          sobre la columna entera, para no engordar la hoja).
 */
function formatearFechas_(sheet) {
  if (!sheet) return;

  // Se usa la ultima fila con periodo, no getLastRow(): si la hoja trae formato
  // hasta abajo, no queremos repintar miles de filas vacias (eso perpetuaria el
  // problema de tamano).
  var last = ultimaFilaDeDatos_(sheet);
  if (last < 2) return;

  var header = readHeader_(sheet);
  var columnas = [
    ['fecha_reporte', FORMATO_FECHA],
    [CACHE_HEADER, FORMATO_FECHA_HORA]
  ];

  columnas.forEach(function (par) {
    var idx = header.indexOf(par[0]);
    if (idx === -1) return;
    sheet.getRange(2, idx + 1, last - 1, 1).setNumberFormat(par[1]);
  });
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
 * QUE HACE: recorre la carpeta y TODAS sus subcarpetas, y devuelve todos los
 *          ficheros que encuentra, cada uno con la ruta de carpetas donde esta.
 * DEVUELVE: lista de {file, name, ruta}. "ruta" es algo como "2026/Semanas/".
 * PARA QUE SIRVA LA RUTA: con carpetas por ano, dos ficheros pueden llamarse
 *          igual; la ruta es lo que permite decir cual es cual en los mensajes.
 * DETALLE: la recursion entra a cualquier profundidad. Drive no permite bucles
 *          de carpetas, asi que no hay riesgo de quedarse dando vueltas.
 */
function todosLosFicheros_(folder, prefijo, out) {
  out = out || [];
  prefijo = prefijo || '';

  // Primero los ficheros de este nivel.
  var itF = folder.getFiles();
  while (itF.hasNext()) {
    var f = itF.next();
    out.push({ file: f, name: f.getName(), ruta: prefijo });
  }

  // Despues se baja a cada subcarpeta, anadiendo su nombre a la ruta.
  var itD = folder.getFolders();
  while (itD.hasNext()) {
    var sub = itD.next();
    todosLosFicheros_(sub, prefijo + sub.getName() + '/', out);
  }

  return out;
}

/** Devuelve la ruta completa de un fichero, para los mensajes y el registro. */
function rutaDe_(entrada) {
  return (entrada.ruta || '') + entrada.name;
}

/**
 * QUE HACE: recorre la carpeta y sus subcarpetas, y devuelve solo los ficheros
 *          cuyo nombre encaja con el patron, con su periodo ya deducido.
 * DEVUELVE: lista de objetos {id, name, ruta, periodo, anio, semana, orden,
 *          fechaReporte}, ordenada del periodo mas antiguo al mas reciente.
 * DETALLE: vale cualquier texto delante del periodo ("fichero pruebas - Y26W48")
 *          y tambien con extension ("... - Y26W48.xlsx"), porque la extension
 *          se quita antes de comparar. Mayusculas y minusculas dan igual.
 */
function listSourceFiles_(folder) {
  var re = new RegExp(pattern_(), 'i');
  var todos = todosLosFicheros_(folder);
  var out = [];

  // Sirve para detectar dos ficheros del mismo periodo (mismo fichero duplicado
  // en dos carpetas, o Y26W5 / Y26W05). Guarda la ruta para poder senalarlos.
  var porOrden = {};

  todos.forEach(function (item) {
    var file = item.file;
    var ruta = rutaDe_(item);

    // Quita la extension final, si la hay (punto + letras o numeros).
    var nombre = item.name;
    var base = nombre.replace(/\.[A-Za-z0-9]+$/, '');

    // Se prueba con y sin extension, para que valgan tanto "x - Y26W48" como
    // "x - Y26W48.xlsx". El patron exige que el periodo quede al final.
    var m = base.match(re) || nombre.match(re);

    // Si el nombre no encaja con el patron, se ignora sin avisar.
    if (!m) return;

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

    // Dos ficheros con la misma clave son el mismo periodo: al cargar uno
    // pisaria al otro y no sabrias cual quedo. Se para y se dicen las rutas.
    if (porOrden[orden]) {
      throw new Error('hay dos ficheros para el mismo periodo (' + periodo + '): "' + porOrden[orden] +
        '" y "' + ruta + '". Deja solo uno.');
    }
    porOrden[orden] = ruta;

    out.push({
      id: file.getId(),
      name: nombre,
      ruta: item.ruta,
      periodo: periodo,
      anio: year,
      semana: (CONFIG.periodFormat === 'MONTH') ? '' : num,
      orden: orden,
      fechaReporte: fechaReporte_(periodo)
    });
  });

  // Orden por la clave numerica, que si es cronologico de verdad.
  out.sort(function (a, b) { return a.orden - b.orden; });
  return out;
}

/**
 * QUE HACE: devuelve el patron (expresion regular) que deben cumplir los
 *          nombres de fichero. Si CONFIG.filenamePattern esta relleno, se usa
 *          tal cual; si no, se deduce de periodFormat.
 * EJEMPLO SEMANAL: "fichero pruebas - Y26W48" -> vale, porque el periodo esta
 *          al final. Y = ano, 26 = 2026, W = semana, 48 = semana 48.
 * DETALLE: \\d{1,2} acepta una o dos cifras, para que Y26W5 y Y26W05 valgan.
 */
function pattern_() {
  if (CONFIG.filenamePattern) return CONFIG.filenamePattern;
  return (CONFIG.periodFormat === 'MONTH') ? 'Y(\\d{2})M(\\d{1,2})$' : 'Y(\\d{2})W(\\d{1,2})$';
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
 * QUE HACE: pasa una fecha a texto dia/mes/ano para mostrarla en el registro de
 *          ejecuciones. Formato europeo, nunca anglosajon.
 */
function formatDate_(d) {
  return Utilities.formatDate(d, CONFIG.timezone, FORMATO_FECHA);
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
 * QUE HACE: revisa las celdas de error de formula y trata cada tipo como toca.
 *   - #N/A: es "la formula no encontro resultados". No es un fallo: es un dato
 *     que hay que archivar. Se cuenta y se avisa, pero no se para.
 *   - #REF!, #VALUE!, #DIV/0!, #NAME?, #NULL!, #NUM!, #ERROR!: eso si es una
 *     formula rota. Se para la ejecucion para no arrastrarla al maestro.
 * DETALLE: solo se listan las 5 primeras rotas, para que el mensaje se lea.
 */
function assertNoErrors_(rows, sheetName, headerRow, warnings, destino, periodo) {
  var rotas = /^#(REF|VALUE|DIV\/0|NAME|NULL|NUM|ERROR)/;
  var sinResultado = /^#N\/A/;
  var bad = [];
  var na = 0;

  rows.forEach(function (row, ri) {
    row.forEach(function (cell) {
      if (typeof cell !== 'string') return;
      var v = cell.trim();
      if (rotas.test(v)) {
        if (bad.length < 5) bad.push('fila ' + (ri + headerRow + 1) + ': ' + v);
      } else if (sinResultado.test(v)) {
        na++;
      }
    });
  });

  if (bad.length) {
    throw new Error('celdas con error en la pestana "' + sheetName + '": ' + bad.join(' | '));
  }
  if (na && warnings) {
    warnings.push(destino + ' @ ' + periodo + ': ' + na + ' celdas con #N/A archivadas tal cual');
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
