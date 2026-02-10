/***** telegram.gs *****/

// ==========================================
// 🔧 Este archivo ahora usa CONFIG.gs
// ==========================================

/**
 * Envía mensaje a un chat de Telegram con reintentos automáticos y rate limiting
 * @param {string|number} chatId - ID del chat destino
 * @param {string} mensaje - Mensaje en HTML
 * @param {Object} opciones - Opciones adicionales (opcional)
 * @returns {boolean} - true si se envió correctamente, false si falló
 */
function enviarMensajeTelegram(chatId, mensaje, opciones = {}) {
  // Obtener token desde CONFIG
  const token = getConfig("telegram.token");
  
  if (!token || !chatId) {
    logAConsola("⚠️ Faltan datos: Token o ChatID no configurados.", "warn");
    return false;
  }

  // Aplicar rate limiting (evitar saturar API de Telegram)
  aplicarRateLimiting_(chatId);

  const url = `https://api.telegram.org/bot${token}/sendMessage`;

  const payload = {
    'chat_id': String(chatId),
    'text': mensaje,
    'parse_mode': opciones.parseMode || 'HTML',
    'disable_web_page_preview': opciones.disablePreview !== false // true por defecto
  };

  const fetchOptions = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  // Obtener configuración de reintentos
  const maxReintentos = getConfig("distribucion.max_reintentos") || 3;
  const delayReintentos = getConfig("distribucion.delay_entre_reintentos_ms") || 2000;

  // Intentar envío con reintentos
  for (let intento = 1; intento <= maxReintentos; intento++) {
    try {
      const response = UrlFetchApp.fetch(url, fetchOptions);
      const resultado = JSON.parse(response.getContentText());
      
      if (resultado.ok) {
        logAConsola(`✅ Telegram enviado al ID ${chatId}`, "ok");
        return true;
      } else {
        // API respondió pero con error
        logAConsola(`⚠️ Telegram rechazó mensaje al ID ${chatId}: ${resultado.description}`, "warn");
        
        // Si es error de rate limit (429), esperar más tiempo
        if (resultado.error_code === 429) {
          const retryAfter = resultado.parameters?.retry_after || 5;
          logAConsola(`⏳ Rate limit alcanzado. Esperando ${retryAfter}s...`, "warn");
          Utilities.sleep(retryAfter * 1000);
        }
      }
      
    } catch (e) {
      logAConsola(`❌ Error enviando a Telegram (${chatId}), intento ${intento}/${maxReintentos}: ${e}`, "error");
      
      // Si no es el último intento, esperar antes de reintentar
      if (intento < maxReintentos) {
        Utilities.sleep(delayReintentos);
      }
    }
  }
  
  // Si llegamos aquí, todos los intentos fallaron
  logAConsola(`❌ FALLO CRÍTICO: No se pudo enviar mensaje al ID ${chatId} después de ${maxReintentos} intentos`, "error");
  return false;
}


/**
 * Envía mensaje con saludo personalizado según la hora
 * @param {string|number} chatId - ID del chat
 * @param {string} nombre - Nombre del destinatario
 * @param {string} cuerpoMensaje - Contenido del mensaje
 */
function enviarMensajeConSaludo(chatId, nombre, cuerpoMensaje) {
  const usarSaludo = getConfig("mensajes.usar_saludo_por_hora");
  let mensaje = "";
  
  if (usarSaludo) {
    const saludo = obtenerSaludoSegunHora_();
    mensaje = `${saludo} <b>${nombre}</b>,\n\n${cuerpoMensaje}`;
  } else {
    mensaje = `👋 Hola <b>${nombre}</b>,\n\n${cuerpoMensaje}`;
  }
  
  return enviarMensajeTelegram(chatId, mensaje);
}


/**
 * Notifica a múltiples usuarios (batch)
 * Útil para "Reportes en proceso..." al inicio de distribución
 * @param {Array} destinatarios - [{id: chatId, nombre: "Juan"}, ...]
 * @param {string} mensaje - Mensaje a enviar
 */
function notificarBatch(destinatarios, mensaje) {
  if (!Array.isArray(destinatarios) || destinatarios.length === 0) {
    return;
  }
  
  logAConsola(`📢 Enviando notificación a ${destinatarios.length} destinatarios...`, "info");
  
  let exitosos = 0;
  let fallidos = 0;
  
  destinatarios.forEach((dest, index) => {
    // Pequeña pausa entre envíos para no saturar
    if (index > 0) {
      Utilities.sleep(200);
    }
    
    const enviado = enviarMensajeTelegram(dest.id, mensaje);
    if (enviado) {
      exitosos++;
    } else {
      fallidos++;
    }
  });
  
  logAConsola(`📊 Notificación batch completada: ${exitosos} ✅ | ${fallidos} ❌`, "info");
}


// ==========================================
// ⚙️ FUNCIONES AUXILIARES PRIVADAS
// ==========================================

/**
 * Obtiene saludo según la hora del día
 * @private
 */
function obtenerSaludoSegunHora_() {
  const hora = new Date().getHours();
  
  if (hora >= 6 && hora < 12) {
    return "☀️ Buenos días";
  } else if (hora >= 12 && hora < 19) {
    return "🌤️ Buenas tardes";
  } else {
    return "🌙 Buenas noches";
  }
}


/**
 * Rate Limiting: Evita saturar API de Telegram
 * Guarda timestamp del último envío por chatId
 * @private
 */
function aplicarRateLimiting_(chatId) {
  const delayMinimo = getConfig("telegram.delay_entre_mensajes_ms") || 100;
  const cache = CacheService.getScriptCache();
  const key = `TELEGRAM_LAST_SEND_${chatId}`;
  
  const ultimoEnvio = cache.get(key);
  
  if (ultimoEnvio) {
    const tiempoTranscurrido = Date.now() - Number(ultimoEnvio);
    
    if (tiempoTranscurrido < delayMinimo) {
      const esperaRestante = delayMinimo - tiempoTranscurrido;
      Utilities.sleep(esperaRestante);
    }
  }
  
  // Actualizar timestamp
  cache.put(key, String(Date.now()), 10); // Cache por 10 segundos
}


/**
 * Verifica si el bot está operativo (test de conexión)
 * @returns {boolean}
 */
function verificarConexionTelegram() {
  const token = getConfig("telegram.token");
  
  if (!token) {
    logAConsola("❌ Token de Telegram no configurado", "error");
    return false;
  }
  
  const url = `https://api.telegram.org/bot${token}/getMe`;
  
  try {
    const response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    const resultado = JSON.parse(response.getContentText());
    
    if (resultado.ok) {
      logAConsola(`✅ Bot conectado: @${resultado.result.username}`, "ok");
      return true;
    } else {
      logAConsola(`❌ Error de autenticación: ${resultado.description}`, "error");
      return false;
    }
  } catch (e) {
    logAConsola(`❌ Error verificando conexión: ${e}`, "error");
    return false;
  }
}


// ==========================================
// 🧪 FUNCIONES DE TEST (Opcional)
// ==========================================

/**
 * Test: Enviar mensaje de prueba
 * Ejecutá esta función desde el editor para probar la configuración
 */
function TEST_enviarMensajePrueba() {
  // Cambiá este ID por el tuyo para probar
  const miChatId = "2037005531";
  
  const mensaje = "🧪 <b>Test del sistema de mensajería</b>\n\n" +
                  "Si recibiste esto, la configuración funciona correctamente.\n\n" +
                  `Hora: ${new Date().toLocaleString()}`;
  
  const exito = enviarMensajeTelegram(miChatId, mensaje);
  
  if (exito) {
    Logger.log("✅ Test exitoso - Revisá tu Telegram");
  } else {
    Logger.log("❌ Test fallido - Revisá los logs en la pestaña Consola");
  }
}


/**
 * Test: Verificar conexión del bot
 */
function TEST_verificarBot() {
  const conectado = verificarConexionTelegram();
  
  if (conectado) {
    Logger.log("✅ Bot configurado correctamente");
  } else {
    Logger.log("❌ Problema con el bot - Revisá el token en CONFIG.gs");
  }
}