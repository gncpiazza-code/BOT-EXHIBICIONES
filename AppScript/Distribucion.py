/***** distribucion.gs *****/

// ==========================================
// 🔧 Este archivo ahora usa CONFIG.gs
// ==========================================

// ---
// FUNCIÓN PRINCIPAL DE DISTRIBUCIÓN (COPIA NATIVA)
// MENSAJE SOLO SE ENVÍA CUANDO SE ACTUALIZA LA HOJA DEL VENDEDOR
// ---
function distribuirDatos_NativoCopy(spreadsheetTemporal, mapeoCompleto, tipoReporte, nombreArchivo) {
  
  // 🚦 ACTIVAR SEMÁFORO: Avisar al bot Python que estamos distribuyendo
  if (getConfig("semaforo.habilitar")) {
    activarSemaforo("DISTRIBUYENDO", mapeoCompleto.length);
  }
  
  // 📊 Estadísticas de la distribución
  const stats = {
    total: 0,
    exitosos: 0,
    fallidos: 0,
    omitidos: 0,
    vendedores: []
  };
  
  logAConsola(`🚀 Iniciando distribución de "${nombreArchivo}"...`, "info");
  
  // Verificar timeout de ejecución
  const tiempoInicio = Date.now();
  const tiempoMaxMs = (getConfig("distribucion.tiempo_maximo_ejecucion_min") || 5.5) * 60 * 1000;
  
  // Obtener todas las hojas del Spreadsheet temporal
  const hojasTemporal = spreadsheetTemporal.getSheets();
  
  logAConsola(`📦 Hojas detectadas: ${hojasTemporal.map(h => h.getName()).join(", ")}`, "info");
  
  // --- PROCESAMIENTO DE CADA HOJA ---
  for (const hojaTemporal of hojasTemporal) {
    
    // ⏱️ Control de timeout
    if (Date.now() - tiempoInicio > tiempoMaxMs) {
      logAConsola("⚠️ TIMEOUT: Alcanzado tiempo máximo de ejecución. Abortando...", "warn");
      break;
    }
    
    stats.total++;
    
    const nombreHojaTemporal = hojaTemporal.getName();
    
    let idDestino = null;
    let nombreEncontrado = null; 
    let idTelegram = null;
    let codigoVendedor = null;

    // 1. Buscar vendedor en el mapeo
    for (const [nombreMapeo, id, idChat] of mapeoCompleto) {
      const nombreHojaLimpio = nombreHojaTemporal.toUpperCase().trim();
      const nombreMapeoLimpio = nombreMapeo.toUpperCase().trim();
      
      if (nombreMapeoLimpio.includes(nombreHojaLimpio) || nombreHojaLimpio.includes(nombreMapeoLimpio)) {
        idDestino = id; 
        nombreEncontrado = nombreMapeo; 
        idTelegram = idChat;
        
        // Extraer código de vendedor (ej: "0009 - Juan Perez" → "0009")
        const matchCodigo = nombreMapeo.match(/^(\d{4,})/);
        codigoVendedor = matchCodigo ? matchCodigo[1] : nombreMapeo.substring(0, 5);
        
        break;
      }
    }

    if (!idDestino) {
      logAConsola(`  -> ⚠️ Hoja '${nombreHojaTemporal}' ignorada (sin mapeo).`, "warn");
      stats.omitidos++;
      continue;
    }

    logAConsola(`  -> 📋 Procesando '${nombreHojaTemporal}' para ${nombreEncontrado}...`, "info");
    
    try {
      // 2. DETERMINAR NOMBRE DE PESTAÑA DESTINO
      let nombrePestañaDestino;
      let esArchivoGenerico = false;
      
      if (tipoReporte.toUpperCase().includes("VENTAS") || tipoReporte.toUpperCase().includes("CUENTAS")) {
        // Es archivo de Ventas o Cuentas Corrientes
        nombrePestañaDestino = tipoReporte;
        esArchivoGenerico = false;
      } else {
        // Es archivo genérico (ej: "REPORTE SIGO - 15/01/26")
        nombrePestañaDestino = `${tipoReporte} - ${codigoVendedor}-${nombreEncontrado.replace(/^\d{4,}\s*-\s*/, '')}`;
        esArchivoGenerico = true;
      }
      
      logAConsola(`    -> Nombre de pestaña destino: "${nombrePestañaDestino}"`, "info");

      // 3. COPIAR HOJA COMPLETA CON TODOS LOS ESTILOS
      const resultadoCopia = copiarHojaCompletaConEstilos_(
        hojaTemporal, 
        idDestino, 
        nombrePestañaDestino,
        esArchivoGenerico
      );
      
      if (!resultadoCopia.exito) {
        logAConsola(`    -> ❌ Error copiando hoja: ${resultadoCopia.error}`, "error");
        stats.fallidos++;
        continue;
      }
      
      const urlDirecta = resultadoCopia.url;
      
      logAConsola(`    -> ✅ COPIA COMPLETA OK para ${nombreEncontrado}.`, "ok");

      // 4. ✅ Enviar notificación por Telegram SOLO A ESTE VENDEDOR
      // Este es el ÚNICO lugar donde se envían mensajes
      if (idTelegram) {
        const resultadoEnvio = enviarNotificacionReporte(
          idTelegram, 
          nombreEncontrado, 
          tipoReporte,
          nombreArchivo,
          urlDirecta,
          esArchivoGenerico
        );
        
        if (resultadoEnvio.enviado) {
          stats.exitosos++;
          stats.vendedores.push(nombreEncontrado);
        } else if (resultadoEnvio.omitido) {
          logAConsola(`    -> 🛑 Omitido: ${resultadoEnvio.motivo}`, "info");
          stats.omitidos++;
        } else {
          stats.fallidos++;
        }
      } else {
        logAConsola(`    -> ℹ️ Sin ID de Telegram configurado.`, "info");
        stats.exitosos++;
      }

    } catch (e) {
      logAConsola(`    -> ❌ ERROR procesando ${nombreEncontrado}: ${e}`, "error");
      stats.fallidos++;
    }
    
    // Pequeña pausa entre vendedores (evita saturación)
    Utilities.sleep(getConfig("distribucion.delay_envios_ms") || 500);
  }
  
  // --- FINALIZACIÓN ---
  
  // 📊 Log de estadísticas finales
  logAConsola("=".repeat(60), "info");
  logAConsola(`📊 DISTRIBUCIÓN COMPLETADA`, "ok");
  logAConsola(`   Total procesados: ${stats.total}`, "info");
  logAConsola(`   ✅ Exitosos: ${stats.exitosos}`, "ok");
  logAConsola(`   ❌ Fallidos: ${stats.fallidos}`, stats.fallidos > 0 ? "error" : "info");
  logAConsola(`   ⚠️ Omitidos: ${stats.omitidos}`, "info");
  logAConsola("=".repeat(60), "info");
  
  // 🚦 DESACTIVAR SEMÁFORO: El bot Python puede continuar
  if (getConfig("semaforo.habilitar")) {
    desactivarSemaforo();
  }
  
  // Guardar dashboard de la distribución
  registrarDistribucionEnDashboard(nombreArchivo, stats);
}


// ==========================================
// COPIA COMPLETA DE HOJA
// ==========================================
function copiarHojaCompletaConEstilos_(hojaTemporal, idDestino, nombrePestañaDestino, esArchivoGenerico) {
  try {
    // 1. Abrir spreadsheet destino
    const spreadsheetDestino = SpreadsheetApp.openById(idDestino);
    const urlBase = spreadsheetDestino.getUrl();
    
    // 2. Verificar si la hoja destino ya existe
    let hojaExistente = spreadsheetDestino.getSheetByName(nombrePestañaDestino);
    
    if (hojaExistente) {
      // Borrar hoja existente para reemplazarla
      spreadsheetDestino.deleteSheet(hojaExistente);
      logAConsola(`      -> Hoja existente "${nombrePestañaDestino}" eliminada para reemplazar`, "info");
    }
    
    // 3. COPIAR HOJA COMPLETA usando copyTo()
    const hojaCopiada = hojaTemporal.copyTo(spreadsheetDestino);
    
    // 4. Renombrar la hoja copiada
    hojaCopiada.setName(nombrePestañaDestino);
    
    logAConsola(`      -> Hoja copiada y renombrada a "${nombrePestañaDestino}"`, "info");
    
    // 5. AGREGAR TIMESTAMP VISIBLE
    agregarTimestampDestacado_(hojaCopiada, esArchivoGenerico);
    
    // 6. Obtener URL directa con #gid=
    const sheetId = hojaCopiada.getSheetId();
    const urlDirecta = `${urlBase}#gid=${sheetId}`;
    
    return { exito: true, url: urlDirecta };
    
  } catch (e) {
    return { exito: false, error: String(e) };
  }
}


// ==========================================
// AGREGAR TIMESTAMP VISIBLE
// ==========================================
function agregarTimestampDestacado_(hoja, esArchivoGenerico) {
  try {
    // Generar timestamp formateado
    const ahora = new Date();
    const dia = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd-MM-yyyy");
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm");
    const textoTimestamp = `REPORTE ACTUALIZADO EL ${dia} a las ${hora}`;
    
    if (esArchivoGenerico) {
      // ARCHIVOS GENÉRICOS: Insertar fila en A1 y desplazar todo
      
      // Insertar nueva fila al inicio
      hoja.insertRowBefore(1);
      
      // Determinar ancho de fusión (máximo 10 columnas o lo que haya)
      const maxCols = Math.min(hoja.getMaxColumns(), 10);
      
      // Fusionar celdas A1:J1 (o hasta donde haya)
      if (maxCols > 1) {
        hoja.getRange(1, 1, 1, maxCols).merge();
      }
      
      // Celda A1
      const celda = hoja.getRange("A1");
      celda.setValue(textoTimestamp);
      
      // Aplicar formato destacado
      celda.setBackground("#fff2cc")        // Amarillo claro
           .setFontWeight("bold")
           .setFontSize(11)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle")
           .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
      
      // Ajustar altura de fila
      hoja.setRowHeight(1, 30);
      
      logAConsola(`      -> Timestamp agregado en A1 (archivo genérico)`, "info");
      
    } else {
      // VENTAS/CUENTAS CORRIENTES: Colocar en J2 (esquina derecha)
      
      const celda = hoja.getRange("J2");
      celda.setValue(textoTimestamp);
      
      // Aplicar formato destacado
      celda.setBackground("#fff2cc")        // Amarillo claro
           .setFontWeight("bold")
           .setFontSize(10)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle")
           .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Asegurar que la columna J sea visible y tenga ancho suficiente
      hoja.setColumnWidth(10, 280); // J = columna 10
      
      logAConsola(`      -> Timestamp agregado en J2 (Ventas/Cuentas)`, "info");
    }
    
  } catch (e) {
    logAConsola(`      -> ⚠️ Error agregando timestamp: ${e}`, "warn");
    // No es crítico, no detenemos el proceso
  }
}


// ==========================================
// ENVÍO DE NOTIFICACIONES
// ==========================================
function enviarNotificacionReporte(chatId, nombreVendedor, tipoReporte, nombreArchivo, urlDirecta, esArchivoGenerico) {
  
  // 1. Verificar anti-duplicados
  const yaEnviado = verificarDuplicado(chatId, nombreArchivo);
  
  if (yaEnviado) {
    return { 
      enviado: false, 
      omitido: true, 
      motivo: "Ya enviado recientemente (cache anti-spam)" 
    };
  }
  
  // 2. Generar URL con tracking (si está habilitado)
  let urlFinal = urlDirecta;
  
  if (getConfig("tracker.habilitar_tracking")) {
    urlFinal = generarUrlConTracking(urlDirecta, nombreVendedor, nombreArchivo);
  }
  
  // 3. Detectar tipo CORRECTO del reporte
  let tipoDetectado = "GENÉRICO";
  
  if (tipoReporte.toUpperCase().includes("VENTAS") || tipoReporte.toUpperCase().includes("VENTA")) {
    tipoDetectado = "VENTAS";
  } else if (tipoReporte.toUpperCase().includes("CUENTAS") || tipoReporte.toUpperCase().includes("CTA")) {
    tipoDetectado = "CUENTAS";
  }
  
  // Si es archivo genérico, usar tipo GENÉRICO
  if (esArchivoGenerico) {
    tipoDetectado = "GENÉRICO";
  }
  
  // 4. Generar mensaje personalizado
  const mensaje = generarMensajeReporte(nombreVendedor, tipoDetectado, nombreArchivo, urlFinal);
  
  // 5. Enviar mensaje
  const enviado = enviarMensajeTelegram(chatId, mensaje);
  
  if (enviado) {
    // Marcar como enviado en cache (anti-spam)
    marcarComoEnviado(chatId, nombreArchivo);
    
    return { enviado: true };
  } else {
    return { enviado: false, omitido: false };
  }
}


// ==========================================
// GENERAR MENSAJE PERSONALIZADO
// ==========================================
function generarMensajeReporte(nombre, tipoReporte, nombreArchivo, url) {
  
  // Emojis desde config
  const emojiVentas = getConfig("mensajes.emoji_ventas") || "💰";
  const emojiCuentas = getConfig("mensajes.emoji_cuentas") || "📉";
  
  let cuerpo = "";
  
  const tipo = tipoReporte.toUpperCase();
  
  if (tipo.includes("VENTAS") || tipo.includes("VENTA")) {
    // Mensaje para VENTAS
    cuerpo = `${emojiVentas} <b>Reporte de Ventas</b>\n\n` +
             `📊 <b>¿Qué incluye?</b>\n` +
             `Desglose detallado de todas las operaciones facturadas en VENDO. ` +
             `Podrás ver fechas, clientes y montos totales recaudados por día.\n\n`;
             
  } else if (tipo.includes("CUENTAS") || tipo.includes("CTA")) {
    // Mensaje para CUENTAS CORRIENTES
    cuerpo = `${emojiCuentas} <b>Estado de Cuentas Corrientes</b>\n\n` +
             `📊 <b>¿Qué incluye?</b>\n` +
             `Listado de clientes con saldo pendiente, ordenado por ` +
             `<b>antigüedad de la deuda</b>. Incluye cantidad de comprobantes ` +
             `adeudados y monto exacto a cobrar.\n\n`;
             
  } else {
    // Mensaje GENÉRICO para otros tipos de reportes
    cuerpo = `📄 <b>Nuevo Reporte Disponible</b>\n\n` +
             `Se ha generado un nuevo reporte: <b>${nombreArchivo}</b>\n\n` +
             `Ya está disponible en tu hoja de cálculo.\n\n`;
  }
  
  // Agregar saludo personalizado si está habilitado
  let mensaje = "";
  if (getConfig("mensajes.usar_saludo_por_hora")) {
    const saludo = obtenerSaludoSegunHora();
    mensaje = `${saludo} <b>${nombre}</b>,\n\n${cuerpo}`;
  } else {
    mensaje = `👋 Hola <b>${nombre}</b>,\n\n${cuerpo}`;
  }
  
  // Link de acceso
  mensaje += `🔗 <b>ACCESO:</b> <a href="${url}">👉 Abrir Planilla</a>`;
  
  return mensaje;
}


function obtenerSaludoSegunHora() {
  const hora = new Date().getHours();
  
  if (hora >= 6 && hora < 12) return "☀️ Buenos días";
  if (hora >= 12 && hora < 19) return "🌤️ Buenas tardes";
  return "🌙 Buenas noches";
}


// ==========================================
// GENERACIÓN DE URL CON TRACKING
// ==========================================
function generarUrlConTracking(urlDirecta, nombreVendedor, nombreArchivo) {
  const urlWebApp = getConfig("tracker.url_webapp");
  
  if (!urlWebApp) {
    return urlDirecta;
  }
  
  try {
    const ahora = new Date();
    const timestampEnvio = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    const params = {
      url: encodeURIComponent(urlDirecta),
      vendedor: encodeURIComponent(nombreVendedor),
      archivo: encodeURIComponent(nombreArchivo),
      envio: encodeURIComponent(timestampEnvio)
    };
    
    const queryString = Object.keys(params)
      .map(key => `${key}=${params[key]}`)
      .join('&');
    
    return `${urlWebApp}?${queryString}`;
    
  } catch (e) {
    logAConsola(`⚠️ Error generando URL tracking: ${e}`, "warn");
    return urlDirecta;
  }
}


// ==========================================
// ANTI-DUPLICADOS
// ==========================================
function verificarDuplicado(chatId, nombreArchivo) {
  const cache = CacheService.getScriptCache();
  const horasCache = getConfig("distribucion.cache_duplicados_horas") || 6;
  const safeFileName = nombreArchivo.replace(/[^a-zA-Z0-9]/g, '');
  const cacheKey = `MSG_SENT_${chatId}_${safeFileName}`;
  
  return cache.get(cacheKey) !== null;
}


function marcarComoEnviado(chatId, nombreArchivo) {
  const cache = CacheService.getScriptCache();
  const horasCache = getConfig("distribucion.cache_duplicados_horas") || 6;
  const safeFileName = nombreArchivo.replace(/[^a-zA-Z0-9]/g, '');
  const cacheKey = `MSG_SENT_${chatId}_${safeFileName}`;
  
  cache.put(cacheKey, "true", horasCache * 3600);
}


// ==========================================
// SISTEMA DE SEMÁFORO
// ==========================================
function activarSemaforo(estado, totalArchivos) {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const pestañaControl = getConfig("google.pestana_control_bot") || "BOT_CONTROL";
    
    const ss = SpreadsheetApp.openById(hojaId);
    let hoja = ss.getSheetByName(pestañaControl);
    
    if (!hoja) {
      hoja = ss.insertSheet(pestañaControl);
      hoja.getRange("A1:D1").setValues([["ESTADO", "INICIO", "ARCHIVOS_TOTAL", "PROGRESO"]]);
      hoja.getRange("A1:D1").setFontWeight("bold").setBackground("#1a1a1a").setFontColor("#ffffff");
    }
    
    const ahora = new Date();
    const timestamp = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    hoja.getRange("A2").setValue(estado);
    hoja.getRange("B2").setValue(timestamp);
    hoja.getRange("C2").setValue(totalArchivos || 0);
    hoja.getRange("D2").setValue("0/" + (totalArchivos || 0));
    
    logAConsola(`🚦 Semáforo activado: ${estado}`, "info");
    
  } catch (e) {
    logAConsola(`⚠️ Error activando semáforo: ${e}`, "warn");
  }
}


function desactivarSemaforo() {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const pestañaControl = getConfig("google.pestana_control_bot") || "BOT_CONTROL";
    
    const ss = SpreadsheetApp.openById(hojaId);
    const hoja = ss.getSheetByName(pestañaControl);
    
    if (hoja) {
      hoja.getRange("A2").setValue("LIBRE");
      hoja.getRange("B2").setValue("");
      hoja.getRange("C2").setValue(0);
      hoja.getRange("D2").setValue("");
    }
    
    logAConsola(`🚦 Semáforo desactivado: LIBRE`, "ok");
    
  } catch (e) {
    logAConsola(`⚠️ Error desactivando semáforo: ${e}`, "warn");
  }
}


function actualizarProgresoSemaforo(actual, total) {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const pestañaControl = getConfig("google.pestana_control_bot") || "BOT_CONTROL";
    
    const ss = SpreadsheetApp.openById(hojaId);
    const hoja = ss.getSheetByName(pestañaControl);
    
    if (hoja) {
      hoja.getRange("D2").setValue(`${actual}/${total}`);
    }
  } catch (e) {
    // No crítico
  }
}


// ==========================================
// DASHBOARD DE DISTRIBUCIONES
// ==========================================
function registrarDistribucionEnDashboard(nombreArchivo, stats) {
  try {
    const hojaId = getConfig("google.hoja_mapeo");
    const ss = SpreadsheetApp.openById(hojaId);
    let dashboard = ss.getSheetByName("DASHBOARD_DISTRIBUCIONES");
    
    if (!dashboard) {
      dashboard = ss.insertSheet("DASHBOARD_DISTRIBUCIONES");
      dashboard.getRange("A1:G1").setValues([[
        "Fecha", "Hora", "Archivo", "Exitosos", "Fallidos", "Omitidos", "Duración"
      ]]);
      dashboard.getRange("A1:G1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
    
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm:ss");
    
    dashboard.appendRow([
      fecha,
      hora,
      nombreArchivo,
      stats.exitosos,
      stats.fallidos,
      stats.omitidos,
      "-"
    ]);
    
  } catch (e) {
    logAConsola(`⚠️ Error registrando en dashboard: ${e}`, "warn");
  }
}