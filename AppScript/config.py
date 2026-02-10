/***** CONFIG.gs *****/

// ==========================================
// 🔧 CONFIGURACIÓN GLOBAL DEL SISTEMA
// ==========================================
// Editá solo los valores dentro de las comillas.
// ⚠️ NO modifiques la estructura (llaves, comas, etc)

const CONFIG = {
  
  // --- TELEGRAM ---
  telegram: {
    token: "TOKEN_ACA_EN_CONFIG_GLOBAL",
    
    // Notificaciones automáticas
    notificar_inicio_distribucion: true,   // Avisar "Procesando reportes..."
    notificar_fin_distribucion: true,      // Avisar "Reportes listos"
    
    // Rate Limiting (para no saturar API de Telegram)
    delay_entre_mensajes_ms: 100,          // Milisegundos entre mensajes
    max_mensajes_por_segundo: 20           // Límite de seguridad
  },
  
  // --- GOOGLE DRIVE / SHEETS ---
  google: {
    // IDs de carpetas
    carpeta_procesar: "1jifljJwB1uk-fSkUxafG4TrNXknp_UTi",
    carpeta_archivados: "1beUEBxbZh3dG2sxnXPUrVuCE_FeEUB39",
    carpeta_temporal: "1ANyns3M3Yt2G6kPYTGiCHmu5ZBF1wLNe",  // 🆕 NUEVA: Para conversiones temporales
    
    // IDs de hojas de cálculo
    hoja_mapeo: "1n2jtp4ZdBh0PlqDurLjkellNK8Vdn5Zxb3qwXaLVftk",
    hoja_tracking: "1n2jtp4ZdBh0PlqDurLjkellNK8Vdn5Zxb3qwXaLVftk",  // Puede ser la misma
    
    // Nombres de pestañas
    pestana_consola: "Consola",
    pestana_tracking: "TRACKING",
    pestana_control_bot: "BOT_CONTROL",        // Para semáforo
    pestana_cola_imagenes: "COLA_IMAGENES"     // Para imágenes pendientes
  },
  
  // --- TRACKER (Sistema de seguimiento de clics) ---
  tracker: {
    // URL de tu Web App (la obtenés después de deployar tracker.gs)
    url_webapp: "https://script.google.com/macros/s/AKfycbxyfWdqA1DqxUMQArEyt9WX26yTDuyLO4ogM3y-5xrne3MjZQWMoOpHNUP4HGt6Wb7M/exec",
    
    habilitar_tracking: true,                  // false = envía links directos sin tracking
    tiempo_redirect_ms: 500                    // Tiempo en pantalla de "Abriendo..."
  },
  
  // --- DISTRIBUCIÓN (Proceso de reparto de reportes) ---
  distribucion: {
    max_reintentos: 3,                         // Intentos si falla envío a Telegram
    delay_entre_reintentos_ms: 2000,           // Espera entre reintentos (2 segundos)
    
    batch_size: 5,                             // Procesar de a X vendedores (evita timeout)
    
    // Anti-spam
    cache_duplicados_horas: 6,                 // No re-enviar mismo archivo en X horas
    
    // Timeout de seguridad
    tiempo_maximo_ejecucion_min: 5.5,          // Apps Script tiene límite de 6 min
    
    // 🆕 NUEVO: Delay entre vendedores (para no saturar)
    delay_envios_ms: 500                       // Milisegundos entre cada vendedor
  },
  
  // --- MENSAJES PERSONALIZADOS ---
  mensajes: {
    usar_saludo_por_hora: true,                // "Buenos días", "Buenas tardes", etc.
    
    // Emojis en mensajes (cambialos si querés)
    emoji_ventas: "💰",
    emoji_cuentas: "📉",
    emoji_proceso: "⏳",
    emoji_listo: "✅",
    emoji_error: "❌"
  },
  
  // --- LOGS Y DEBUG ---
  logs: {
    nivel: "INFO",                             // DEBUG | INFO | WARN | ERROR
    guardar_en_sheet: true,                    // Escribir logs en pestaña Consola
    max_lineas_consola: 2000,                  // Limpiar si excede este número
    formato_fecha: "dd/MM/yyyy HH:mm:ss"
  },
  
  // --- SISTEMA DE SEMÁFORO (Coordinación Bot Python ↔ Apps Script) ---
  semaforo: {
    habilitar: true,                           // false = deshabilita coordinación
    intervalo_verificacion_seg: 5,             // Cada cuánto verificar estado
    timeout_distribucion_min: 20               // Si pasa este tiempo, forzar LIBRE
  }
};

// ==========================================
// ⛔ NO TOCAR DE AQUÍ PARA ABAJO
// ==========================================

/**
 * Función helper para obtener config con validación
 * Uso: const token = getConfig("telegram.token");
 */
function getConfig(path) {
  const parts = path.split('.');
  let value = CONFIG;
  
  for (const part of parts) {
    if (value && typeof value === 'object' && part in value) {
      value = value[part];
    } else {
      Logger.log(`⚠️ Config no encontrada: ${path}`);
      return null;
    }
  }
  
  return value;
}

/**
 * Validación de config al iniciar (llamar desde main.gs si querés)
 */
function validarConfig() {
  const errores = [];
  
  // Validar token
  if (!CONFIG.telegram.token || CONFIG.telegram.token.length < 40) {
    errores.push("❌ Token de Telegram inválido o vacío");
  }
  
  // Validar IDs de carpetas (deben tener 33 caracteres aprox)
  if (!CONFIG.google.carpeta_procesar || CONFIG.google.carpeta_procesar.length < 20) {
    errores.push("❌ ID de carpeta 'procesar' inválido");
  }
  
  if (!CONFIG.google.carpeta_temporal || CONFIG.google.carpeta_temporal.length < 20) {
    errores.push("❌ ID de carpeta 'temporal' inválido");
  }
  
  if (!CONFIG.google.hoja_mapeo || CONFIG.google.hoja_mapeo.length < 20) {
    errores.push("❌ ID de hoja de mapeo inválido");
  }
  
  // Validar URL del tracker
  if (CONFIG.tracker.habilitar_tracking && !CONFIG.tracker.url_webapp.includes("script.google.com")) {
    errores.push("⚠️ URL del tracker parece incorrecta");
  }
  
  // Mostrar resultados
  if (errores.length > 0) {
    Logger.log("=".repeat(50));
    Logger.log("⚠️ ERRORES EN CONFIGURACIÓN:");
    errores.forEach(e => Logger.log(e));
    Logger.log("=".repeat(50));
    return false;
  } else {
    Logger.log("✅ Configuración validada correctamente");
    return true;
  }
}

/**
 * Función para mostrar la config actual (útil para debugging)
 */
function mostrarConfigActual() {
  Logger.log("=".repeat(60));
  Logger.log("📋 CONFIGURACIÓN ACTUAL DEL SISTEMA");
  Logger.log("=".repeat(60));
  Logger.log(JSON.stringify(CONFIG, null, 2));
  Logger.log("=".repeat(60));
}
