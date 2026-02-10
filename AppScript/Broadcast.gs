/***** Broadcast.gs *****/

// ==========================================
// 📢 SISTEMA DE COMUNICACIÓN MASIVA
// ==========================================

/**
 * Crea la pestaña BROADCAST si no existe
 */
function crearPestañaBroadcast() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName("BROADCAST");
  
  if (hoja) return hoja; // Ya existe
  
  // Crear pestaña
  hoja = ss.insertSheet("BROADCAST");
  
  // --- DISEÑO DE LA PESTAÑA ---
  
  // TÍTULO
  hoja.getRange("A1:F1").merge();
  hoja.getRange("A1")
    .setValue("📢 SISTEMA DE BROADCAST - COMUNICACIÓN MASIVA")
    .setBackground("#0f172a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  hoja.setRowHeight(1, 45);
  
  // SECCIÓN: MENSAJE
  hoja.getRange("A3").setValue("📝 MENSAJE A ENVIAR:").setFontWeight("bold").setFontSize(11);
  
  // Celda grande para escribir mensaje
  hoja.getRange("A4:F12").merge();
  hoja.getRange("A4")
    .setValue("Escribí acá tu mensaje...\n\nPodés usar HTML:\n<b>negrita</b>\n<i>cursiva</i>\n\nEl mensaje se enviará tal cual lo escribas.")
    .setBackground("#f8fafc")
    .setBorder(true, true, true, true, null, null, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("top")
    .setWrap(true)
    .setFontColor("#64748b");
  
  // SECCIÓN: OPCIONES
  hoja.getRange("A14").setValue("🎯 DESTINATARIOS:").setFontWeight("bold").setFontSize(11);
  
  hoja.getRange("A15").setValue("⚙️ Modo:");
  hoja.getRange("B15")
    .setValue("TODOS")
    .setBackground("#fef3c7")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  hoja.getRange("A16").setValue("📊 Se enviará a:");
  hoja.getRange("B16")
    .setValue("Calculando...")
    .setBackground("#dbeafe")
    .setHorizontalAlignment("center");
  
  // SECCIÓN: BOTONES (Espacio reservado - los botones se crean con InterfazBotones.gs)
  hoja.getRange("A18").setValue("📤 ACCIONES:").setFontWeight("bold").setFontSize(11);
  hoja.getRange("A19:F19").merge();
  hoja.getRange("A19")
    .setValue("[Los botones se crearán automáticamente al cargar el proyecto]")
    .setBackground("#f1f5f9")
    .setFontColor("#94a3b8")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  hoja.setRowHeight(19, 60);
  
  // SECCIÓN: HISTORIAL
  hoja.getRange("A22").setValue("📊 HISTORIAL DE BROADCASTS").setFontWeight("bold").setFontSize(12);
  
  hoja.getRange("A23:F23").setValues([[
    "Fecha", "Hora", "Destinatarios", "Total", "Exitosos", "Vista Previa"
  ]]);
  
  hoja.getRange("A23:F23")
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  hoja.setFrozenRows(23);
  
  // Anchos de columna
  hoja.setColumnWidth(1, 100);  // Fecha
  hoja.setColumnWidth(2, 80);   // Hora
  hoja.setColumnWidth(3, 150);  // Destinatarios
  hoja.setColumnWidth(4, 70);   // Total
  hoja.setColumnWidth(5, 80);   // Exitosos
  hoja.setColumnWidth(6, 300);  // Vista Previa
  
  logAConsola("✅ Pestaña BROADCAST creada", "ok");
  return hoja;
}


/**
 * Actualiza el contador de destinatarios según el modo seleccionado
 */
function actualizarContadorDestinatarios() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaBroadcast = ss.getSheetByName("BROADCAST");
  
  if (!hojaBroadcast) return;
  
  const modo = hojaBroadcast.getRange("B15").getValue().toString().toUpperCase();
  let total = 0;
  
  const mapeo = getMapeoFiltrado_();
  
  if (modo === "TODOS") {
    total = mapeo.length;
  } else if (modo === "SELECCIONADOS") {
    total = contarVendedoresSeleccionados();
  } else if (modo === "ACTIVOS") {
    total = contarVendedoresActivos();
  }
  
  hojaBroadcast.getRange("B16").setValue(total + " vendedores");
}


/**
 * Cuenta vendedores con checkbox activado
 */
function contarVendedoresSeleccionados() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaMapeo = ss.getSheetByName("Mapeo");
  
  if (!hojaMapeo) return 0;
  
  // Leer columna D (checkboxes) desde fila 2 hasta última con datos
  const ultimaFila = hojaMapeo.getLastRow();
  if (ultimaFila < 2) return 0;
  
  const checks = hojaMapeo.getRange(2, 4, ultimaFila - 1, 1).getValues();
  let contador = 0;
  
  checks.forEach(fila => {
    if (fila[0] === true) contador++;
  });
  
  return contador;
}


/**
 * Cuenta vendedores que abrieron reportes en los últimos 7 días
 */
function contarVendedoresActivos() {
  const hojaId = getConfig("google.hoja_tracking");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaTracking = ss.getSheetByName("TRACKING");
  
  if (!hojaTracking) return 0;
  
  // Lógica simple: vendedores únicos en TRACKING de últimos 7 días
  const datos = hojaTracking.getDataRange().getValues();
  const ahora = new Date();
  const hace7dias = new Date(ahora.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  const vendedoresActivos = new Set();
  
  for (let i = 1; i < datos.length; i++) {
    try {
      const fechaStr = datos[i][0]; // Columna A: Fecha
      if (!fechaStr) continue;
      
      const partes = fechaStr.split('/');
      if (partes.length !== 3) continue;
      
      const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      
      if (fecha >= hace7dias) {
        const vendedor = datos[i][2]; // Columna C: Vendedor
        if (vendedor) vendedoresActivos.add(vendedor);
      }
    } catch (e) {
      // Ignorar errores de parsing
    }
  }
  
  return vendedoresActivos.size;
}


/**
 * Obtiene lista de destinatarios según modo
 */
function obtenerDestinatarios(modo) {
  const mapeo = getMapeoFiltrado_();
  const destinatarios = [];
  
  if (modo === "TODOS") {
    mapeo.forEach(([nombre, idSheet, idTelegram]) => {
      if (idTelegram) {
        destinatarios.push({ nombre, id: idTelegram });
      }
    });
  } 
  else if (modo === "SELECCIONADOS") {
    const hojaId = getConfig("google.hoja_mapeo");
    const ss = SpreadsheetApp.openById(hojaId);
    const hojaMapeo = ss.getSheetByName("Mapeo");
    const ultimaFila = hojaMapeo.getLastRow();
    
    if (ultimaFila >= 2) {
      const datos = hojaMapeo.getRange(2, 1, ultimaFila - 1, 4).getValues();
      
      datos.forEach(fila => {
        const nombre = fila[0];
        const idTelegram = fila[2];
        const seleccionado = fila[3]; // Columna D: checkbox
        
        if (seleccionado === true && idTelegram) {
          destinatarios.push({ nombre, id: idTelegram });
        }
      });
    }
  }
  else if (modo === "ACTIVOS") {
    // Obtener vendedores activos de últimos 7 días
    const hojaId = getConfig("google.hoja_tracking");
    const ss = SpreadsheetApp.openById(hojaId);
    const hojaTracking = ss.getSheetByName("TRACKING");
    
    if (hojaTracking) {
      const datos = hojaTracking.getDataRange().getValues();
      const ahora = new Date();
      const hace7dias = new Date(ahora.getTime() - 7 * 24 * 60 * 60 * 1000);
      
      const nombresActivos = new Set();
      
      for (let i = 1; i < datos.length; i++) {
        try {
          const fechaStr = datos[i][0];
          if (!fechaStr) continue;
          
          const partes = fechaStr.split('/');
          if (partes.length !== 3) continue;
          
          const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
          
          if (fecha >= hace7dias) {
            nombresActivos.add(datos[i][2]);
          }
        } catch (e) {}
      }
      
      // Mapear a IDs de Telegram
      mapeo.forEach(([nombre, idSheet, idTelegram]) => {
        if (nombresActivos.has(nombre) && idTelegram) {
          destinatarios.push({ nombre, id: idTelegram });
        }
      });
    }
  }
  
  return destinatarios;
}


/**
 * FUNCIÓN PRINCIPAL: Enviar broadcast
 */
function enviarBroadcast(modo) {
  const ui = SpreadsheetApp.getUi();
  
  // Leer mensaje
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaBroadcast = ss.getSheetByName("BROADCAST");
  
  if (!hojaBroadcast) {
    ui.alert("Error", "No se encontró la pestaña BROADCAST", ui.ButtonSet.OK);
    return;
  }
  
  const mensaje = hojaBroadcast.getRange("A4").getValue();
  
  if (!mensaje || mensaje.trim() === "" || mensaje.includes("Escribí acá tu mensaje")) {
    ui.alert("Error", "Debes escribir un mensaje antes de enviar", ui.ButtonSet.OK);
    return;
  }
  
  // Obtener destinatarios
  const destinatarios = obtenerDestinatarios(modo);
  
  if (destinatarios.length === 0) {
    ui.alert("Error", "No hay destinatarios para enviar", ui.ButtonSet.OK);
    return;
  }
  
  // Confirmar envío
  const respuesta = ui.alert(
    'Confirmar Broadcast',
    `¿Enviar mensaje a ${destinatarios.length} vendedor(es)?`,
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta !== ui.Button.YES) {
    return;
  }
  
  // Enviar mensajes
  logAConsola(`📢 Iniciando broadcast a ${destinatarios.length} destinatarios (${modo})...`, "info");
  
  let exitosos = 0;
  let fallidos = 0;
  
  destinatarios.forEach((dest, index) => {
    if (index > 0) {
      Utilities.sleep(getConfig("telegram.delay_entre_mensajes_ms") || 100);
    }
    
    const enviado = enviarMensajeTelegram(dest.id, mensaje);
    
    if (enviado) {
      exitosos++;
    } else {
      fallidos++;
    }
  });
  
  // Registrar en historial
  registrarBroadcastEnHistorial(modo, destinatarios.length, exitosos, mensaje);
  
  logAConsola(`📊 Broadcast completado: ${exitosos} ✅ | ${fallidos} ❌`, exitosos > 0 ? "ok" : "error");
  
  ui.alert(
    'Broadcast Enviado',
    `Mensajes enviados: ${exitosos}\nFallidos: ${fallidos}`,
    ui.ButtonSet.OK
  );
}


/**
 * Registra el broadcast en el historial
 */
function registrarBroadcastEnHistorial(modo, total, exitosos, mensaje) {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const ss = SpreadsheetApp.openById(hojaId);
    const hojaBroadcast = ss.getSheetByName("BROADCAST");
    
    if (!hojaBroadcast) return;
    
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm:ss");
    
    // Truncar mensaje para vista previa (primeros 50 caracteres)
    const vistaPrevia = mensaje.length > 50 ? mensaje.substring(0, 50) + "..." : mensaje;
    
    hojaBroadcast.appendRow([
      fecha,
      hora,
      modo,
      total,
      exitosos,
      vistaPrevia
    ]);
    
  } catch (e) {
    logAConsola(`⚠️ Error registrando historial: ${e}`, "warn");
  }
}


/**
 * Función para los botones: Enviar a TODOS
 */
function BROADCAST_enviarATodos() {
  enviarBroadcast("TODOS");
}


/**
 * Función para los botones: Enviar a SELECCIONADOS
 */
function BROADCAST_enviarASeleccionados() {
  enviarBroadcast("SELECCIONADOS");
}


/**
 * Función para los botones: Enviar a ACTIVOS
 */
function BROADCAST_enviarAActivos() {
  enviarBroadcast("ACTIVOS");
}


/**
 * Función para los botones: Vista previa
 */
function BROADCAST_vistaPrevia() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaBroadcast = ss.getSheetByName("BROADCAST");
  
  if (!hojaBroadcast) return;
  
  const mensaje = hojaBroadcast.getRange("A4").getValue();
  const modo = hojaBroadcast.getRange("B15").getValue();
  const destinatarios = obtenerDestinatarios(modo);
  
  const ui = SpreadsheetApp.getUi();
  
  let preview = `═══════════════════════════\n`;
  preview += `📢 VISTA PREVIA DEL BROADCAST\n`;
  preview += `═══════════════════════════\n\n`;
  preview += `🎯 Modo: ${modo}\n`;
  preview += `📊 Destinatarios: ${destinatarios.length}\n\n`;
  preview += `📝 MENSAJE:\n`;
  preview += `${mensaje}\n\n`;
  preview += `═══════════════════════════\n`;
  preview += `Primeros 5 destinatarios:\n`;
  
  destinatarios.slice(0, 5).forEach((d, i) => {
    preview += `${i + 1}. ${d.nombre}\n`;
  });
  
  if (destinatarios.length > 5) {
    preview += `... y ${destinatarios.length - 5} más`;
  }
  
  ui.alert("Vista Previa", preview, ui.ButtonSet.OK);
}
