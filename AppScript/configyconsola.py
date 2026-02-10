/***** configyconsola.gs *****/

// ==========================================
// 🔧 Este archivo ahora usa CONFIG.gs
// ==========================================

// Cola / Trigger (constantes que NO están en CONFIG)
const PROP_QUEUE_JSON   = "QUEUE_JSON";
const PROP_QUEUE_INDEX  = "QUEUE_INDEX";
const TRIGGER_FUNC_NAME = "continuarEjecucionAuto";

const MESES_ABREVIADOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];


// ==========================================
// 📋 CONSOLA PRO
// ==========================================

/**
 * Inicializa o resetea la hoja de consola con formato profesional
 */
function consolaInit_Pro(reset) {
  const hojaId = getConfig("google.hoja_mapeo");
  const nombreConsola = getConfig("google.pestana_consola") || "Consola";
  
  if (!hojaId) {
    Logger.log("⚠️ ID de hoja no configurado");
    return null;
  }
  
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName(nombreConsola);
  
  if (!hoja) {
    hoja = ss.insertSheet(nombreConsola);
  }
  
  if (reset) {
    hoja.clear();
  }

  // Crear encabezados si no existen
  if (hoja.getLastRow() === 0) {
    hoja.insertRows(1);
  }
  
  hoja.getRange(1, 1, 1, 3).setValues([["⏰ Timestamp", "📊 Nivel", "💬 Mensaje"]]);

  hoja.getRange(1, 1, 1, 3)
    .setFontWeight("bold")
    .setFontFamily("Inter, -apple-system, Arial")
    .setFontSize(10)
    .setBackground("#111827")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(false)
    .setBorder(
      true, true, true, true, true, true,
      "#374151",
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );

  hoja.setFrozenRows(1);
  hoja.setColumnWidth(1, 180);   // Timestamp
  hoja.setColumnWidth(2, 90);    // Nivel
  hoja.setColumnWidth(3, 1000);  // Mensaje

  // Formato de fecha/hora legible
  hoja.getRange("A2:A").setNumberFormat("dd/MM/yyyy HH:mm:ss");

  return hoja;
}


/**
 * Agrega una línea a la consola con colores y formato
 * @param {string} mensaje - Mensaje a loguear
 * @param {string} nivel - ok|success|info|warn|error
 */
function logAConsola(mensaje, nivel) {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const nombreConsola = getConfig("google.pestana_consola") || "Consola";
    
    if (!hojaId) {
      // Fallback a Logger si no hay config
      Logger.log(`[${nivel || 'INFO'}] ${mensaje}`);
      return;
    }
    
    const ss = SpreadsheetApp.openById(hojaId);
    let hoja = ss.getSheetByName(nombreConsola);
    
    if (!hoja) {
      hoja = consolaInit_Pro(false);
    }

    const lvl = (nivel || "info").toLowerCase();
    
    // Emojis según nivel
    let icon = "";
    let colorTexto = "#1e293b";
    
    switch(lvl) {
      case "ok":
      case "success":
        icon = "✅";
        colorTexto = "#059669"; // Verde
        break;
      case "warn":
      case "warning":
        icon = "⚠️";
        colorTexto = "#d97706"; // Naranja
        break;
      case "error":
      case "err":
        icon = "❌";
        colorTexto = "#dc2626"; // Rojo
        break;
      case "debug":
        icon = "🔍";
        colorTexto = "#6366f1"; // Índigo
        break;
      default:
        icon = "ℹ️";
        colorTexto = "#64748b"; // Gris
    }

    // Agregar fila
    hoja.appendRow([new Date(), icon + " " + lvl.toUpperCase(), String(mensaje)]);

    const last = hoja.getLastRow();
    
    if (last >= 2) {
      const rowRange = hoja.getRange(last, 1, 1, 3);

      // Zebra rows (filas alternadas)
      const zebraBg = (last % 2 === 0) ? "#ffffff" : "#f9fafb";

      rowRange
        .setFontFamily("Inter, -apple-system, Arial")
        .setFontSize(10)
        .setVerticalAlignment("middle")
        .setWrap(true)
        .setBackground(zebraBg)
        .setBorder(
          false, false, false, false, false, false,
          "#e5e7eb",
          SpreadsheetApp.BorderStyle.SOLID_DOTTED
        );
      
      // Color del texto según nivel
      rowRange.getCell(1, 2).setFontColor(colorTexto).setFontWeight("bold");
      
      // Ajustar anchos (por si se desajustaron)
      hoja.setColumnWidth(1, 180);
      hoja.setColumnWidth(2, 90);
      hoja.setColumnWidth(3, 1000);
    }

    // Auto-limpieza: mantener solo las últimas X filas
    const maxRows = getConfig("logs.max_lineas_consola") || 2000;
    const total = hoja.getLastRow();
    
    if (total > maxRows) {
      const filasAEliminar = total - maxRows;
      hoja.deleteRows(2, filasAEliminar);
    }

  } catch (e) {
    // Si falla la consola, al menos loguear en Logger
    Logger.log(`CONSOLA ERROR: ${e} :: ${mensaje}`);
  }
}


// ==========================================
// 🛡️ THROTTLE DE WARNINGS (Anti-spam)
// ==========================================

const WARN_TS_PREFIX = 'WARN_TS__';

/**
 * Verifica si ya pasó el tiempo mínimo para repetir un warning
 * @private
 */
function shouldLogWarn_(key, minMinutes) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const minMs = Math.max(1, Number(minMinutes || 3)) * 60 * 1000;
  const lastStr = props.getProperty(WARN_TS_PREFIX + key) || '0';
  const last = Number(lastStr);
  
  if (now - last >= minMs) {
    props.setProperty(WARN_TS_PREFIX + key, String(now));
    return true;
  }
  
  return false;
}


/**
 * Loguea un warning con throttle (evita spam)
 * @param {string} key - Clave única para identificar el tipo de warning
 * @param {string} message - Mensaje a loguear
 * @param {number} minMinutes - Minutos mínimos entre repeticiones (default: 3)
 */
function logWarnThrottled(key, message, minMinutes) {
  if (shouldLogWarn_(key, minMinutes || 3)) {
    logAConsola(message, 'warn');
  }
}


/**
 * Limpia el throttle de una clave (útil cuando se resuelve el problema)
 * @param {string} key - Clave del warning a limpiar
 */
function clearWarnThrottle_(key) {
  PropertiesService.getScriptProperties().deleteProperty(WARN_TS_PREFIX + key);
}


// ==========================================
// 📊 UTILIDADES DE CONSOLA
// ==========================================

/**
 * Limpia la consola completamente
 */
function limpiarConsola() {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const nombreConsola = getConfig("google.pestana_consola") || "Consola";
    
    if (!hojaId) return;
    
    const ss = SpreadsheetApp.openById(hojaId);
    const hoja = ss.getSheetByName(nombreConsola);
    
    if (hoja) {
      consolaInit_Pro(true); // Reset completo
      logAConsola("🧹 Consola limpiada manualmente", "info");
    }
  } catch (e) {
    Logger.log(`Error limpiando consola: ${e}`);
  }
}


/**
 * Registra inicio de una operación (útil para medir duración)
 * @param {string} operacion - Nombre de la operación
 * @returns {number} - Timestamp de inicio
 */
function logInicioOperacion(operacion) {
  const inicio = Date.now();
  logAConsola(`🚀 Iniciando: ${operacion}`, "info");
  return inicio;
}


/**
 * Registra fin de una operación con duración
 * @param {string} operacion - Nombre de la operación
 * @param {number} timestampInicio - Timestamp retornado por logInicioOperacion
 */
function logFinOperacion(operacion, timestampInicio) {
  const fin = Date.now();
  const duracionMs = fin - timestampInicio;
  const duracionSeg = (duracionMs / 1000).toFixed(2);
  
  logAConsola(`✅ Completado: ${operacion} (${duracionSeg}s)`, "success");
}


/**
 * Loguea un separador visual (útil para separar secciones de logs)
 * @param {string} titulo - Título del separador (opcional)
 */
function logSeparador(titulo) {
  const linea = "─".repeat(50);
  
  if (titulo) {
    logAConsola(`${linea} ${titulo} ${linea}`, "info");
  } else {
    logAConsola(linea, "info");
  }
}


/**
 * Loguea información estructurada (objeto convertido a JSON legible)
 * @param {string} etiqueta - Etiqueta descriptiva
 * @param {Object} objeto - Objeto a loguear
 */
function logObjeto(etiqueta, objeto) {
  try {
    const json = JSON.stringify(objeto, null, 2);
    logAConsola(`${etiqueta}:\n${json}`, "debug");
  } catch (e) {
    logAConsola(`${etiqueta}: [Error al serializar objeto]`, "error");
  }
}


// ==========================================
// 🧪 FUNCIONES DE TEST
// ==========================================

/**
 * Test: Genera logs de todos los niveles
 */
function TEST_consolaNiveles() {
  logSeparador("TEST DE CONSOLA");
  
  logAConsola("Este es un mensaje de DEBUG", "debug");
  logAConsola("Este es un mensaje de INFO", "info");
  logAConsola("Este es un mensaje de WARNING", "warn");
  logAConsola("Este es un mensaje de ERROR", "error");
  logAConsola("Este es un mensaje de SUCCESS", "success");
  
  logSeparador();
  
  Logger.log("✅ Test completado - Revisá la pestaña Consola");
}


/**
 * Test: Verifica funcionamiento del throttle
 */
function TEST_consolaThrottle() {
  logSeparador("TEST DE THROTTLE");
  
  logWarnThrottled("test_key", "⚠️ Este warning debería aparecer", 1);
  Utilities.sleep(500);
  logWarnThrottled("test_key", "⚠️ Este warning NO debería aparecer (throttled)", 1);
  
  logAConsola("Esperá 1 minuto y ejecutá de nuevo para ver el siguiente warning", "info");
  
  logSeparador();
  
  Logger.log("✅ Test completado");
}


/**
 * Test: Verifica medición de duración de operaciones
 */
function TEST_consolaDuracion() {
  logSeparador("TEST DE DURACIÓN");
  
  const inicio = logInicioOperacion("Operación de prueba");
  
  // Simular trabajo
  Utilities.sleep(2000); // 2 segundos
  
  logFinOperacion("Operación de prueba", inicio);
  
  logSeparador();
  
  Logger.log("✅ Test completado");
}