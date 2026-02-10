/***** InterfazBotones.gs *****/

// ==========================================
// 🎛️ SISTEMA DE INTERFAZ CON BOTONES
// ==========================================

/**
 * Inicializa todos los botones en las pestañas
 * Se llama automáticamente en onOpen()
 */
function inicializarBotonesPestañas() {
  try {
    logAConsola("🎨 Inicializando botones en pestañas...", "info");
    
    crearBotonesAnalytics();
    crearBotonesBroadcast();
    crearBotonesMapeo();
    
    logAConsola("✅ Botones inicializados correctamente", "ok");
  } catch (e) {
    logAConsola(`⚠️ Error inicializando botones: ${e}`, "warn");
  }
}


/**
 * Crea botones en ANALYTICS_DASHBOARD
 */
function crearBotonesAnalytics() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName("ANALYTICS_DASHBOARD");
  
  if (!hoja) {
    hoja = crearPestañaAnalytics();
  }
  
  // Limpiar botones existentes en el área
  limpiarBotonesEnArea(hoja, 36, 1, 5, 6); // Filas 36-40, columnas A-F
  
  // TÍTULO DE SECCIÓN
  hoja.getRange("A36:F36").merge();
  hoja.getRange("A36")
    .setValue("🎛️ PANEL DE CONTROL")
    .setBackground("#0f172a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center");
  hoja.setRowHeight(36, 35);
  
  // BOTONES DE DISTRIBUCIÓN
  hoja.getRange("A38").setValue("📦 DISTRIBUCIÓN:").setFontWeight("bold");
  
  // Botón 1: EJECUTAR AHORA
  crearBoton(hoja, "B38", "▶️ EJECUTAR AHORA", "ejecutarManualmente", "#10b981", "#ffffff");
  
  // Botón 2: INICIAR COLA AUTOMÁTICA
  crearBoton(hoja, "C38", "📦 COLA AUTOMÁTICA", "iniciarColaAuto", "#3b82f6", "#ffffff");
  
  // Botón 3: PAUSAR COLA
  crearBoton(hoja, "D38", "⏸️ PAUSAR", "detenerColaAuto", "#f59e0b", "#ffffff");
  
  // Botón 4: DETENER Y BORRAR
  crearBoton(hoja, "E38", "⏹️ RESET", "borrarColaYDetener", "#ef4444", "#ffffff");
  
  // Botón 5: VER ESTADO
  crearBoton(hoja, "F38", "📊 ESTADO", "verEstadoCola", "#6366f1", "#ffffff");
  
  // BOTONES DE ANALYTICS
  hoja.getRange("A40").setValue("📊 ANALYTICS:").setFontWeight("bold");
  
  // Botón: ACTUALIZAR DASHBOARD
  crearBoton(hoja, "B40", "🔄 ACTUALIZAR", "ANALYTICS_actualizarDashboard", "#8b5cf6", "#ffffff");
  
  logAConsola("  -> Botones de ANALYTICS_DASHBOARD creados", "info");
}


/**
 * Crea botones en BROADCAST
 */
function crearBotonesBroadcast() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName("BROADCAST");
  
  if (!hoja) {
    hoja = crearPestañaBroadcast();
  }
  
  // Limpiar área de botones
  limpiarBotonesEnArea(hoja, 19, 1, 1, 6);
  hoja.getRange("A19:F19").merge();
  
  // Crear contenedor para botones
  hoja.getRange("A19:F19")
    .setBackground("#f8fafc")
    .setBorder(true, true, true, true, null, null, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  hoja.setRowHeight(19, 70);
  
  // BOTONES (usando celdas sin merge para que se vean en fila)
  const fila = 19;
  
  // Botón 1: ENVIAR A TODOS
  crearBotonEnCelda(hoja, `A${fila}`, "📤 TODOS", "BROADCAST_enviarATodos", "#10b981", "#ffffff", 140, 50);
  
  // Botón 2: ENVIAR A SELECCIONADOS
  crearBotonEnCelda(hoja, `B${fila}`, "📤 SELECCIONADOS", "BROADCAST_enviarASeleccionados", "#3b82f6", "#ffffff", 180, 50);
  
  // Botón 3: ENVIAR A ACTIVOS
  crearBotonEnCelda(hoja, `D${fila}`, "📤 ACTIVOS", "BROADCAST_enviarAActivos", "#f59e0b", "#ffffff", 140, 50);
  
  // Botón 4: VISTA PREVIA
  crearBotonEnCelda(hoja, `E${fila}`, "👁️ VISTA PREVIA", "BROADCAST_vistaPrevia", "#6366f1", "#ffffff", 160, 50);
  
  logAConsola("  -> Botones de BROADCAST creados", "info");
}


/**
 * Crea botones en Mapeo
 */
function crearBotonesMapeo() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName("Mapeo");
  
  if (!hoja) return;
  
  // Agregar columna de checkboxes si no existe
  if (hoja.getRange("D1").getValue() !== "☑️ Seleccionar") {
    hoja.getRange("D1")
      .setValue("☑️ Seleccionar")
      .setBackground("#0f172a")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    
    // Agregar checkboxes en todas las filas con datos
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila >= 2) {
      const rangoCheckboxes = hoja.getRange(2, 4, ultimaFila - 1, 1);
      rangoCheckboxes.insertCheckboxes();
    }
    
    hoja.setColumnWidth(4, 120);
  }
  
  // Crear área de botones (debajo de la tabla)
  const filaBoton = hoja.getLastRow() + 3;
  
  // Botón: SELECCIONAR TODOS
  crearBoton(hoja, `A${filaBoton}`, "✅ SELECCIONAR TODOS", "MAPEO_seleccionarTodos", "#10b981", "#ffffff");
  
  // Botón: DESELECCIONAR TODOS
  crearBoton(hoja, `B${filaBoton}`, "❌ DESELECCIONAR TODOS", "MAPEO_deseleccionarTodos", "#ef4444", "#ffffff");
  
  logAConsola("  -> Botones de MAPEO creados", "info");
}


/**
 * Crea un botón estándar en una celda
 */
function crearBoton(hoja, celda, texto, funcionAsignada, colorFondo, colorTexto) {
  const range = hoja.getRange(celda);
  
  // Configurar celda
  range
    .setValue(texto)
    .setBackground(colorFondo)
    .setFontColor(colorTexto)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  
  hoja.setRowHeight(range.getRow(), 40);
  
  // Asignar macro (solo si existe la función)
  try {
    range.setFontFamily("Arial");
    // Note: Apps Script no permite asignar macros via código directamente
    // Los usuarios deberán hacer click derecho > Asignar secuencia de comandos
    // y escribir el nombre de la función manualmente la primera vez
  } catch (e) {
    // Ignorar si falla
  }
}


/**
 * Crea un botón con tamaño personalizado
 */
function crearBotonEnCelda(hoja, celda, texto, funcionAsignada, colorFondo, colorTexto, ancho, alto) {
  const range = hoja.getRange(celda);
  
  range
    .setValue(texto)
    .setBackground(colorFondo)
    .setFontColor(colorTexto)
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Ajustar tamaños
  if (alto) hoja.setRowHeight(range.getRow(), alto);
  if (ancho) hoja.setColumnWidth(range.getColumn(), ancho);
}


/**
 * Limpia botones en un área específica
 */
function limpiarBotonesEnArea(hoja, filaInicio, colInicio, cantFilas, cantCols) {
  try {
    const rango = hoja.getRange(filaInicio, colInicio, cantFilas, cantCols);
    
    // Buscar y eliminar imágenes/dibujos en el área
    const drawings = hoja.getDrawings();
    drawings.forEach(drawing => {
      try {
        const container = drawing.getContainerInfo();
        if (container) {
          const anchorRow = container.getAnchorRow();
          const anchorCol = container.getAnchorColumn();
          
          if (anchorRow >= filaInicio && anchorRow < filaInicio + cantFilas &&
              anchorCol >= colInicio && anchorCol < colInicio + cantCols) {
            drawing.remove();
          }
        }
      } catch (e) {
        // Ignorar errores al eliminar
      }
    });
  } catch (e) {
    // No crítico
  }
}


// ==========================================
// FUNCIONES PARA BOTONES DE MAPEO
// ==========================================

/**
 * Selecciona todos los checkboxes en Mapeo
 */
function MAPEO_seleccionarTodos() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName("Mapeo");
  
  if (!hoja) return;
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;
  
  const rangoCheckboxes = hoja.getRange(2, 4, ultimaFila - 1, 1);
  rangoCheckboxes.check();
  
  SpreadsheetApp.getUi().alert(
    'Selección Completa',
    `Se seleccionaron ${ultimaFila - 1} vendedor(es).`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Deselecciona todos los checkboxes en Mapeo
 */
function MAPEO_deseleccionarTodos() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName("Mapeo");
  
  if (!hoja) return;
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;
  
  const rangoCheckboxes = hoja.getRange(2, 4, ultimaFila - 1, 1);
  rangoCheckboxes.uncheck();
  
  SpreadsheetApp.getUi().alert(
    'Deselección Completa',
    'Se deseleccionaron todos los vendedores.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ==========================================
// FUNCIÓN PARA BOTÓN: VER ESTADO DE COLA
// ==========================================

/**
 * Muestra el estado actual de la cola de procesamiento
 */
function verEstadoCola() {
  const state = loadQueueState_();
  const ui = SpreadsheetApp.getUi();
  
  let mensaje = "";
  
  if (!state || !state.queue || state.queue.length === 0) {
    mensaje = "📭 No hay cola activa.\n\n";
    mensaje += "La cola se crea automáticamente al iniciar el procesamiento.";
  } else {
    const queue = state.queue;
    const index = state.index;
    
    const total = queue.length;
    const completados = queue.filter(item => item.status === "done").length;
    const pendientes = queue.filter(item => item.status === "pending").length;
    const enProceso = queue.filter(item => item.status === "processing").length;
    const errores = queue.filter(item => item.status === "error").length;
    
    mensaje = "📊 ESTADO DE LA COLA DE PROCESAMIENTO\n";
    mensaje += "═".repeat(40) + "\n\n";
    mensaje += `📦 Total de archivos: ${total}\n`;
    mensaje += `✅ Completados: ${completados}\n`;
    mensaje += `⏳ Pendientes: ${pendientes}\n`;
    mensaje += `🔄 En proceso: ${enProceso}\n`;
    mensaje += `❌ Errores: ${errores}\n\n`;
    mensaje += `📍 Posición actual: ${index}/${total}\n\n`;
    
    if (index < total) {
      mensaje += `⏭️ Próximo archivo:\n`;
      mensaje += `   "${queue[index].fileName}"`;
    } else {
      mensaje += `✅ Cola completada`;
    }
  }
  
  ui.alert('Estado de la Cola', mensaje, ui.ButtonSet.OK);
}


// ==========================================
// FUNCIÓN HELPER: Crear botones "reales" con Apps Script
// ==========================================

/**
 * Crea un botón real usando imágenes y scripts asignados
 * (Alternativa a celdas formateadas)
 */
function crearBotonReal(hoja, fila, columna, texto, funcionAsignada, color) {
  // Esta función crea botones "reales" usando Drawing API
  // Es más complejo pero permite asignar macros automáticamente
  
  // Por simplicidad, en esta versión usamos celdas formateadas
  // Los usuarios deben asignar las macros manualmente la primera vez:
  // Click derecho en la celda > "Asignar secuencia de comandos" > Nombre de función
}
