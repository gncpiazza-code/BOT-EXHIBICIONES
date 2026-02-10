/***** main.gs - VERSIÓN COMPLETA Y CORREGIDA *****/

// ==========================================
// 🔧 Este archivo ahora usa CONFIG.gs
// ==========================================

/**
 * MENÚ INICIAL
 * - Ejecutar distribución AHORA = modo viejo (una sola corrida monolítica)
 * - Iniciar/Reanudar cola      = modo nuevo con cola/checkpoints + trigger
 * - Pausar cola                = apaga el trigger recurrente (mantiene la cola)
 * - Borrar y Detener Cola      = apaga el trigger Y borra la cola (Reset)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🤖 Robot de Reportes')
    .addItem('Ejecutar distribución AHORA', 'ejecutarManualmente')
    .addSeparator()
    .addItem('▶️ Iniciar/Reanudar cola automática', 'iniciarColaAuto')
    .addItem('⏸️ Pausar cola (Solo detener trigger)', 'detenerColaAuto')
    .addItem('⏹️ Borrar y Detener Cola (Reset)', 'borrarColaYDetener') 
    .addToUi();
  // Inicializa consola visual (sin borrar)
  try { consolaInit_Pro(false); } catch(e) {}
}


/**
 * BOTÓN MANUAL (modo actual, todo de una vez)
 */
function ejecutarManualmente() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmar',
    '¿Iniciar el proceso de distribución de reportes ahora?',
    ui.ButtonSet.YES_NO
  );
  if (respuesta == ui.Button.YES) {
    
    // --- BORRAR CONSOLA ANTES DE INICIAR ---
    try {
      consolaInit_Pro(true); // true = Borra todo y resetea encabezados
    } catch(e) {
      console.error("No se pudo resetear la consola: " + e);
    }

    logAConsola("🚀 Inicio de ejecución MANUAL...", "info");
    ejecutarDistribucionDeReportes();
  }
}


/**
 * FLUJO NUEVO (copia nativa de hojas completas):
 * - Escanea carpeta
 * - Procesa TODOS los archivos .xlsx
 * - Convierte cada .xlsx a Google Sheets temporal
 * - Copia hojas completas (con estilos) a cada vendedor
 * - Borra Sheets temporal
 * - Mueve archivo a "Archivados"
 */
function ejecutarDistribucionDeReportes() {
  try {
    logAConsola("Iniciando robot de reportes (Modo: Copia Nativa / manual)...", "info");
    
    // 1. Obtener mapeo de vendedores (AHORA INCLUYE TELEGRAM)
    const mapeoFiltrado = getMapeoFiltrado_();
    logAConsola(`Mapeo de ${mapeoFiltrado.length} vendedores cargado.`, "info");

    // 2. Obtener carpetas DESDE CONFIG
    const carpetaProcesar = DriveApp.getFolderById(getConfig("google.carpeta_procesar"));
    const carpetaArchivar = DriveApp.getFolderById(getConfig("google.carpeta_archivados"));

    // 3. Buscar archivos .xlsx
    const archivos = carpetaProcesar.getFiles();
    let archivosProcesados  = 0;
    let archivosEncontrados = 0;
    
    while (archivos.hasNext()) {
      const archivoXlsx   = archivos.next();
      const nombreArchivo = archivoXlsx.getName();

      if (!esArchivoValidoExcel_(nombreArchivo)) {
        continue;
      }

      archivosEncontrados++;
      logAConsola(`--- Procesando archivo: ${nombreArchivo} ---`, "info");
      
      // Determinar nombre de pestaña destino (Ventas/Mes o CC o genérico)
      const pestañaDestinoNombre = getPestañaDestinoParaArchivo_(nombreArchivo);
      
      // 🆕 NUEVO: Convertir .xlsx a Google Sheets temporal
      let spreadsheetTemporal;
      let idTemporal;
      
      try {
        logAConsola(`-> Convirtiendo ${nombreArchivo} a Google Sheets temporal...`, "info");
        idTemporal = convertirXlsxASheets_(archivoXlsx);
        spreadsheetTemporal = SpreadsheetApp.openById(idTemporal);
        logAConsola(`-> Conversión exitosa. ID temporal: ${idTemporal}`, "info");
      } catch (e) {
        logAConsola(`❌ ERROR: No se pudo convertir el archivo .xlsx (${nombreArchivo}). Detalle: ${e}`, "error");
        continue;
      }

      // 🆕 NUEVO: Distribuir usando copia nativa (en lugar de SheetJS)
      try {
        distribuirDatos_NativoCopy(spreadsheetTemporal, mapeoFiltrado, pestañaDestinoNombre, nombreArchivo);
      } catch (e) {
        logAConsola(`❌ ERROR distribuyendo datos de ${nombreArchivo}: ${e} (Stack: ${e.stack})`, "error");
      }

      // 🆕 NUEVO: Borrar Sheets temporal
      try {
        borrarSheetsTemporal_(idTemporal);
        logAConsola(`-> Sheets temporal borrado (${idTemporal}).`, "info");
      } catch (e) {
        logAConsola(`⚠️ AVISO: No se pudo borrar Sheets temporal: ${e}`, "warn");
      }

      // Mover archivo a Archivados
      try {
        archivoXlsx.moveTo(carpetaArchivar);
        logAConsola(`-> Archivo ${nombreArchivo} procesado y archivado.`, "ok");
        archivosProcesados++; 
      } catch (e) {
        logAConsola(`ERROR al archivar ${nombreArchivo}: ${e}`, "error");
      }
    }

    if (archivosEncontrados === 0) {
      logAConsola("No se encontraron archivos .xlsx para procesar.", "warn");
    }
    logAConsola(`✔ Proceso manual completado. ${archivosProcesados} archivo(s) procesado(s).`, "ok");
  } catch (e) {
    logAConsola(`ERROR FATAL en ejecutarDistribucionDeReportes: ${e}`, "error");
  }
}


// =======================================================
// ==========  MODO AUTOMÁTICO CON COLA / TRIGGER  =======
// =======================================================

// Construye la cola inicial y habilita el trigger recurrente
function iniciarColaAuto() {
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);

  try {
    // --- BORRAR CONSOLA ANTES DE INICIAR ---
    consolaInit_Pro(true); // true = Borra todo y resetea encabezados

    logAConsola("🚀 Inicio cola automática solicitado.", "info");
    logAConsola("... Re-escaneando carpeta 'archivos a procesar'...", "info");
    
    // 1. SIEMPRE volvemos a escanear la carpeta y creamos la cola
    const nuevaCola = buildQueueFromFolder_();
    
    // Obtener constantes desde CONFIG
    const PROP_QUEUE_JSON = "QUEUE_JSON";
    const PROP_QUEUE_INDEX = "QUEUE_INDEX";
    
    saveQueueState_(nuevaCola, 0); // Guardamos la *nueva* cola y reseteamos el índice

    logAConsola(`📦 Nueva cola creada: ${nuevaCola.length} archivo(s) pendiente(s).`, "info");
    if (nuevaCola.length > 0) {
      logAConsola(`📌 Checkpoint inicial: índice 0/${nuevaCola.length} — Próximo "${nuevaCola[0].fileName}"`, "info");
    }

    // 2. Asegurar trigger cada 1 minuto
    ensureProcessingTrigger_();
    logAConsola("▶️ Ejecución automática habilitada (cada 1 min).", "ok");

  } catch (e) {
    logAConsola(`❌ ERROR en iniciarColaAuto(): ${e}`, "error");
  } finally {
    lock.releaseLock();
  }
}


// Se llama automatically por el trigger cada 1 min hasta que la cola termine
function continuarEjecucionAuto() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    // Throttle del WARN de lock a 3 minutos
    logWarnThrottled('lock_conflict',
      "⏳ continuarEjecucionAuto(): no se obtuvo lock, puede haber otra corrida activa."
    );
    return;
  }

  // Se obtuvo lock: limpiar el throttle de este WARN
  clearWarnThrottle_('lock_conflict');

  const startMs = Date.now();
  const MAX_MS  = 5 * 60 * 1000; // ~5 minutos de trabajo antes de cortar limpio

  try {
    // 1. Cargar estado
    let state = loadQueueState_();
    if (!state || !state.queue || state.queue.length === 0) {
      logAConsola("ℹ️ No hay cola pendiente. Deteniendo trigger.", "info");
      stopProcessingTrigger_();
      clearQueueState_();
      return;
    }

    const queue = state.queue;
    let index   = state.index;
    
    // Si estamos fuera de rango => terminar
    if (index >= queue.length) {
      logAConsola("✅ Cola completada. Deteniendo trigger.", "ok");
      stopProcessingTrigger_();
      clearQueueState_();
      return;
    }

    // Obtenemos mapeoFiltrado (vendedores) para esta corrida
    const mapeoFiltrado = getMapeoFiltrado_();
    logAConsola(`Mapeo (${mapeoFiltrado.length} vendedores) cargado para corrida automática.`, "info");

    // 2. Procesar archivos mientras haya tiempo
    while (index < queue.length) {
      const item = queue[index];
      
      // Saltar ítems ya 'done'
      if (item.status === "done") {
        index++;
        continue;
      }

      // Mensaje de checkpoint
      logAConsola(`🔁 Retomando desde checkpoint: índice ${index}/${queue.length} — Procesando "${item.fileName}"`, "info");
      
      // Marcar processing
      item.status = "processing";
      
      // Intentar procesar el archivo concreto
      try {
        procesarUnArchivo_(item.fileId, item.fileName, mapeoFiltrado);
        item.status = "done";
        logAConsola(`✅ Terminado "${item.fileName}" (índice ${index}/${queue.length}).`, "ok");
      } catch (errProc) {
        item.status = "error";
        logAConsola(`❌ Error en "${item.fileName}": ${errProc}`, "error");
      }

      index++;
      saveQueueState_(queue, index);

      // 3. Chequeamos tiempo
      const elapsed = Date.now() - startMs;
      if (elapsed > MAX_MS) {
        logAConsola(`⏱️ Se agotó el tiempo de ejecución (~5min). Guardando checkpoint en índice ${index}.`, "warn");
        saveQueueState_(queue, index);
        return; // El trigger volverá a llamar en 1 min
      }
    }

    // Si llegamos acá, terminamos la cola
    logAConsola("✅ Cola completada sin timeout. Deteniendo trigger.", "ok");
    stopProcessingTrigger_();
    clearQueueState_();

  } catch (e) {
    logAConsola(`❌ ERROR en continuarEjecucionAuto: ${e}`, "error");
  } finally {
    lock.releaseLock();
  }
}


// Detiene el trigger automático pero deja la cola intacta (se puede reanudar)
function detenerColaAuto() {
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);

  try {
    stopProcessingTrigger_();
    logAConsola("⏸️ Cola pausada. El trigger automático ha sido detenido.", "ok");
    
    SpreadsheetApp.getUi().alert(
      "Proceso Pausado", 
      "El trigger automático ha sido detenido.\n\n" +
      "Los archivos parcialmente procesados permanecen en la cola.\n\n" +
      "Para reanudar, usá 'Iniciar/Reanudar cola automática'.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    logAConsola(`❌ ERROR en detenerColaAuto(): ${e}`, "error");
  } finally {
    lock.releaseLock();
  }
}


// Borra la cola Y detiene el trigger (reset total)
function borrarColaYDetener() {
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);

  try {
    // 1. Detener triggers
    stopProcessingTrigger_();

    // 2. Borrar el estado guardado
    clearQueueState_();

    logAConsola("⏹️ Cola borrada y trigger detenido (Reset).", "ok");
    SpreadsheetApp.getUi().alert(
      "Proceso Detenido y Borrado", 
      "La cola de archivos pendientes ha sido borrada y el trigger automático ha sido detenido.", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    logAConsola(`❌ ERROR en borrarColaYDetener(): ${e}`, "error");
  } finally {
    lock.releaseLock();
  }
}


// -------------------------------------------------------
// 🆕 FUNCIONES: Conversión .xlsx ↔ Sheets
// -------------------------------------------------------

/**
 * Convierte un archivo .xlsx a Google Sheets temporal
 * @param {File} archivoXlsx - Objeto File de Drive del .xlsx
 * @returns {string} - ID del Spreadsheet temporal creado
 */
function convertirXlsxASheets_(archivoXlsx) {
  const carpetaTemporal = DriveApp.getFolderById(getConfig("google.carpeta_temporal"));
  
  // Nombre para el archivo temporal
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const nombreTemporal = `TEMP_${timestamp}_${archivoXlsx.getName()}`;
  
  try {
    // Copiar el archivo y convertirlo a Google Sheets
    const resource = {
      title: nombreTemporal,
      mimeType: 'application/vnd.google-apps.spreadsheet',
      parents: [{id: carpetaTemporal.getId()}]
    };
    
    // Usar Drive API avanzada para la conversión
    const file = Drive.Files.copy(resource, archivoXlsx.getId(), {convert: true});
    
    logAConsola(`   ✅ Conversión exitosa: ${nombreTemporal} (ID: ${file.id})`, "info");
    return file.id;
    
  } catch (e) {
    logAConsola(`   ❌ Error en conversión: ${e}`, "error");
    throw new Error(`No se pudo convertir ${archivoXlsx.getName()}: ${e}`);
  }
}


/**
 * Borra un Spreadsheet temporal de Drive
 * @param {string} spreadsheetId - ID del Spreadsheet a borrar
 */
function borrarSheetsTemporal_(spreadsheetId) {
  try {
    const file = DriveApp.getFileById(spreadsheetId);
    file.setTrashed(true);
    logAConsola(`   🗑️ Sheets temporal borrado: ${spreadsheetId}`, "info");
  } catch (e) {
    logAConsola(`   ⚠️ No se pudo borrar temporal ${spreadsheetId}: ${e}`, "warn");
    // No es crítico, la carpeta TEMP se puede limpiar manualmente después
  }
}


// -------------------------------------------------------
// Helpers internos para la cola / trigger / parsing
// -------------------------------------------------------

// Devuelve [{fileId, fileName, status:'pending'}, ...] desde carpeta de entrada
function buildQueueFromFolder_() {
  const carpetaProcesar = DriveApp.getFolderById(getConfig("google.carpeta_procesar"));
  const archivos        = carpetaProcesar.getFiles();
  const queue           = [];
  
  while (archivos.hasNext()) {
    const f   = archivos.next();
    const nom = f.getName();
    if (!esArchivoValidoExcel_(nom)) continue;
    queue.push({
      fileId:   f.getId(),
      fileName: nom,
      status:   "pending"
    });
  }

  return queue;
}


// Procesa UN archivo puntual de la cola:
// - detecta tipo (Ventas/Cuentas/Genérico)
// - convierte a Sheets temporal
// - distribuirDatos_NativoCopy(...) -> copia hojas completas
// - borra temporal
// - mueve el archivo a Archivados
function procesarUnArchivo_(fileId, fileName, mapeoFiltrado) {
  const carpetaArchivar = DriveApp.getFolderById(getConfig("google.carpeta_archivados"));
  const fileObj         = DriveApp.getFileById(fileId);
  
  // Determinar pestaña destino
  const pestañaDestinoNombre = getPestañaDestinoParaArchivo_(fileName);
  
  // 🆕 NUEVO: Convertir a Sheets temporal
  logAConsola(`-> Convirtiendo ${fileName} a Sheets temporal...`, "info");
  let spreadsheetTemporal;
  let idTemporal;
  
  try {
    idTemporal = convertirXlsxASheets_(fileObj);
    spreadsheetTemporal = SpreadsheetApp.openById(idTemporal);
    logAConsola(`-> Conversión exitosa.`, "info");
  } catch (e) {
    logAConsola(`❌ ERROR al convertir ${fileName}: ${e}`, "error");
    throw e;
  }

  // 🆕 NUEVO: Distribuir usando copia nativa
  try {
    distribuirDatos_NativoCopy(spreadsheetTemporal, mapeoFiltrado, pestañaDestinoNombre, fileName);
  } catch (e) {
    logAConsola(`❌ ERROR al distribuir datos de ${fileName}: ${e} (Stack: ${e.stack})`, "error");
    throw e;
  }

  // 🆕 NUEVO: Borrar temporal
  try {
    borrarSheetsTemporal_(idTemporal);
    logAConsola(`-> Sheets temporal borrado.`, "info");
  } catch (e) {
    logAConsola(`⚠️ No se pudo borrar temporal: ${e}`, "warn");
  }

  // Mover archivo a ARCHIVADOS
  try {
    fileObj.moveTo(carpetaArchivar);
    logAConsola(`-> Archivo ${fileName} procesado y archivado.`, "ok");
  } catch (e) {
    logAConsola(`ERROR moviendo ${fileName} a Archivados: ${e}`, "error");
    throw e;
  }
}


// ==========================================
// 🆕 FUNCIÓN MODIFICADA - DETECCIÓN DE MES CON AÑO
// ==========================================
// Devuelve el nombre de pestaña de destino según el archivo (Ventas mes año / CC / nombre archivo)
function getPestañaDestinoParaArchivo_(nombreArchivo) {
  // Obtener meses abreviados desde configyconsola.gs
  const MESES_ABREVIADOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  
  const lowerName = nombreArchivo.toLowerCase();
  
  // Detectar si es de Ventas
  if (lowerName.includes("ventas") || lowerName.includes("venta")) {
    let mesAbreviado = null;
    let añoCorto = null;
    
    // 🆕 NUEVO: Detectar formatos de fecha
    // Formato 1: "DD-MM al DD-MM" (sin año) - Ejemplo: "02-01 al 31-01"
    const regexSinAño = /(\d{2})-(\d{2})\s+al\s+(\d{2})-(\d{2})/i;
    const matchSinAño = nombreArchivo.match(regexSinAño);
    
    if (matchSinAño) {
      // Extraer el mes del primer grupo de fecha
      const mesIndex = parseInt(matchSinAño[2], 10) - 1; // matchSinAño[2] es el mes
      
      if (mesIndex >= 0 && mesIndex < 12) {
        mesAbreviado = MESES_ABREVIADOS[mesIndex];
        
        // Obtener año actual y extraer últimos 2 dígitos
        const añoActual = new Date().getFullYear();
        añoCorto = String(añoActual).slice(-2); // "2026" -> "26"
        
        logAConsola(`-> ✅ Detectado: Ventas de ${mesAbreviado} ${añoCorto} (formato sin año en archivo)`, "info");
        
        return `Ventas (${mesAbreviado} ${añoCorto})`;
      }
    }
    
    // Formato 2 (RETROCOMPATIBILIDAD): "DD-MM-YYYY" - Ejemplo: "02-01-2026"
    const regexConAño = /(\d{2})-(\d{2})-(\d{4})/;
    const matchConAño = nombreArchivo.match(regexConAño);
    
    if (matchConAño) {
      const mesIndex = parseInt(matchConAño[2], 10) - 1;
      const añoCompleto = matchConAño[3];
      
      if (mesIndex >= 0 && mesIndex < 12) {
        mesAbreviado = MESES_ABREVIADOS[mesIndex];
        añoCorto = añoCompleto.slice(-2); // Tomar últimos 2 dígitos del año
        
        logAConsola(`-> ✅ Detectado: Ventas de ${mesAbreviado} ${añoCorto} (formato con año en archivo)`, "info");
        
        return `Ventas (${mesAbreviado} ${añoCorto})`;
      }
    }
    
    // Si no se detectó ningún formato válido
    logAConsola(`-> ⚠️ No se pudo detectar mes/año en "${nombreArchivo}". Usando "Ventas (General)"`, "warn");
    return "Ventas (General)";
  } 
  
  // Detectar si es de Cuentas Corrientes
  else if (lowerName.includes("cuentas") || lowerName.includes("ctacte") || lowerName.includes("cta cte") || lowerName.includes("cta") || lowerName.includes("cte")) {
    return "Cuentas Corrientes";
  }

  // 🆕 NUEVO: Si no es ni Ventas ni Cuentas, retornar el nombre del archivo
  // (se usará para crear pestaña personalizada tipo "REPORTE SIGO - 00009-juan perez")
  return nombreArchivo.replace('.xlsx', '').replace('.XLSX', '');
}


// Devuelve mapeoFiltrado = [[nombreVendedor, idSpreadsheetDestino, idTelegram], ...] sin vacíos
function getMapeoFiltrado_() {
  const hojaMapeo = SpreadsheetApp
    .openById(getConfig("google.hoja_mapeo"))
    .getSheetByName("Mapeo");
    
  // 🆕 MEJORADO: Leer un rango amplio para detectar vendedores agregados en el medio
  const maxFilas = hojaMapeo.getMaxRows();
  const filaInicio = 2; // Empezar después del encabezado
  
  // Leer hasta la última fila con datos (no solo con formato)
  let ultimaFilaConDatos = hojaMapeo.getLastRow();
  
  // Si no hay datos, retornar vacío
  if (ultimaFilaConDatos < filaInicio) return [];
  
  // Calcular cuántas filas leer
  const cantidadFilas = ultimaFilaConDatos - filaInicio + 1;
  
  // Leer 3 columnas: Nombre (A), ID Sheet (B), Chat ID Telegram (C)
  const mapeoVendedores = hojaMapeo.getRange(
    filaInicio, 
    1, 
    cantidadFilas, 
    3
  ).getValues();
  
  // 🆕 MEJORADO: Filtrar filas que tengan al menos Nombre (columna A) Y ID Sheet (columna B)
  // Esto ignora automáticamente filas vacías en el medio
  const mapeoFiltrado = mapeoVendedores.filter(fila => {
    const nombre = String(fila[0]).trim();
    const idSheet = String(fila[1]).trim();
    
    // Debe tener nombre Y debe tener ID de Sheet
    return nombre.length > 0 && idSheet.length > 0;
  });
  
  logAConsola(`📋 Mapeo cargado: ${mapeoFiltrado.length} vendedores válidos detectados`, "info");
  
  return mapeoFiltrado;
}


// Determina si un nombre de archivo es Excel válido de entrada
function esArchivoValidoExcel_(nombreArchivo) {
  if (!nombreArchivo.toLowerCase().endsWith(".xlsx")) return false;
  if (nombreArchivo.startsWith("~")) return false;
  return true;
}


// Lee y guarda el estado de cola en ScriptProperties
function loadQueueState_() {
  const PROP_QUEUE_JSON = "QUEUE_JSON";
  const PROP_QUEUE_INDEX = "QUEUE_INDEX";
  
  const props = PropertiesService.getScriptProperties();
  const rawQ  = props.getProperty(PROP_QUEUE_JSON);
  const rawI  = props.getProperty(PROP_QUEUE_INDEX);
  
  if (!rawQ) {
    return { queue: [], index: 0 };
  }
  
  let queue;
  try {
    queue = JSON.parse(rawQ);
  } catch(e) {
    queue = [];
  }

  let idx = 0;
  if (rawI && !isNaN(parseInt(rawI,10))) {
    idx = parseInt(rawI,10);
  }

  // Mini-autocuración: si algún item estaba "processing" y cortó,
  // lo dejamos pendiente otra vez para reintentar
  for (let k=0; k<queue.length; k++) {
    if (queue[k].status === "processing") {
      queue[k].status = "pending";
    }
  }

  return { queue: queue, index: idx };
}


// Guarda cola e índice
function saveQueueState_(queue, index) {
  const PROP_QUEUE_JSON = "QUEUE_JSON";
  const PROP_QUEUE_INDEX = "QUEUE_INDEX";
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_QUEUE_JSON, JSON.stringify(queue));
  props.setProperty(PROP_QUEUE_INDEX, String(index));
}


// Limpia cola al terminar
function clearQueueState_() {
  const PROP_QUEUE_JSON = "QUEUE_JSON";
  const PROP_QUEUE_INDEX = "QUEUE_INDEX";
  
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_QUEUE_JSON);
  props.deleteProperty(PROP_QUEUE_INDEX);
}


// Crea un trigger time-driven cada 1 minuto para continuarEjecucionAuto()
// si aún no existe
function ensureProcessingTrigger_() {
  const TRIGGER_FUNC_NAME = "continuarEjecucionAuto";
  
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === TRIGGER_FUNC_NAME);

  if (!exists) {
    ScriptApp.newTrigger(TRIGGER_FUNC_NAME)
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}


// Elimina TODOS los triggers que llamen continuarEjecucionAuto()
function stopProcessingTrigger_() {
  const TRIGGER_FUNC_NAME = "continuarEjecucionAuto";
  
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction && t.getHandlerFunction() === TRIGGER_FUNC_NAME) {
      ScriptApp.deleteTrigger(t);
    }
  }
}


// ==========================================
// 🆕 FUNCIÓN DE TEST PARA VERIFICAR DETECCIÓN
// ==========================================
function TEST_deteccionMesYAño() {
  logAConsola("=".repeat(60), "info");
  logAConsola("🧪 TEST: Detección de Mes y Año en nombres de archivo", "info");
  logAConsola("=".repeat(60), "info");
  
  const archivosTest = [
    "Reporte Ventas - Cordoba - 02-01 al 31-01.xlsx",
    "Ventas Sucursal Norte 15-12 al 31-12.xlsx",
    "VENTAS - 01-03 al 30-03.xlsx",
    "Reporte Ventas - 05-07-2026 al 31-07-2026.xlsx", // Formato con año
    "Cuentas Corrientes - Enero.xlsx",
    "Reporte SIGO - 15/01/26.xlsx"
  ];
  
  archivosTest.forEach(nombre => {
    const resultado = getPestañaDestinoParaArchivo_(nombre);
    logAConsola(`📄 "${nombre}" → "${resultado}"`, "info");
  });
  
  logAConsola("=".repeat(60), "info");
  logAConsola("✅ Test completado - Revisá los resultados en la Consola", "ok");
}