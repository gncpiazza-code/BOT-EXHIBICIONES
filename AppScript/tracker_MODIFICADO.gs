/***** tracker.gs - VERSIÓN CON TRACKING MEJORADO *****/

// ==========================================
// 🎯 TRACKER PRO - Sistema de Seguimiento
// VERSIÓN 4.3 - Con columna "Abrió" y formato condicional
// ==========================================

function doGet(e) {
  const params = e.parameter || {};
  
  // ═══════════════════════════════════════
  // FASE 2: Actualización (Click en "Entrar")
  // ═══════════════════════════════════════
  if (params.action === "update" && params.vendedor) {
    try {
      actualizarMetadataUltimoRegistro(
        params.vendedor, 
        params.ua, 
        params.geo 
      );
      return ContentService.createTextOutput("OK");
    } catch (e) {
      return ContentService.createTextOutput("ERROR: " + e);
    }
  }
  
  // ═══════════════════════════════════════
  // FASE 1: Registro inicial
  // ═══════════════════════════════════════
  
  const validacion = validarParametros(e);
  if (!validacion.valido) {
    return mostrarPantallaError(validacion.error);
  }
  
  const datos = validacion.datos;
  
  try {
    registrarClickInicial(datos);
  } catch (err) {
    console.error("⚠️ Error en tracking: " + err);
  }
  
  return generarPantallaRedirect(datos);
}


// ==========================================
// 📝 REGISTRO EN GOOGLE SHEETS
// ==========================================
function registrarClickInicial(datos) {
  const hojaId = getConfig("google.hoja_tracking");
  const pestañaTracking = getConfig("google.pestana_tracking") || "TRACKING";
  
  if (!hojaId) return;
  
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName(pestañaTracking);
  
  if (!hoja) {
    hoja = crearHojaTrackingPro(ss, pestañaTracking);
  } else {
    // Aseguramos que se vean las columnas (ahora son J y K)
    try {
      hoja.showColumns(10); // J
      hoja.showColumns(11); // K
      verificarEncabezados(hoja);
    } catch(e) {}
  }
  
  const ahora = new Date();
  const fechaVista = Utilities.formatDate(ahora, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  const horaVista = Utilities.formatDate(ahora, ss.getSpreadsheetTimeZone(), "HH:mm:ss");
  
  const tipoReporte = detectarTipoReporte(datos.archivo);
  const tiempoReaccion = calcularTiempoReaccion(datos.fechaEnvio, ahora);
  
  // Fila inicial (Estructura con columna "Abrió" K)
  const fila = [
    fechaVista,                    // A: Fecha
    horaVista,                     // B: Hora
    datos.vendedor,                // C: Vendedor
    tipoReporte,                   // D: Tipo
    datos.fechaEnvio,              // E: Enviado
    tiempoReaccion.texto,          // F: Reacción
    "Detectando...",               // G: Dispositivo
    "Detectando...",               // H: Navegador
    "Detectando...",               // I: Sistema
    datos.url,                     // J: LINK
    "No"                           // K: ✅ ABRIÓ (inicialmente "No", se actualiza al hacer click)
  ];
  
  hoja.appendRow(fila);
  
  // 🆕 NUEVO: Aplicar formato condicional a la fila
  const ultimaFila = hoja.getLastRow();
  aplicarFormatoCondicional(hoja, ultimaFila, tiempoReaccion.minutos);
}

// Verifica y corrige encabezados si faltan
function verificarEncabezados(hoja) {
  const rango = hoja.getRange("J1:K1");
  const valores = rango.getValues()[0];
  
  // Link en columna 10 (J)
  if (valores[0] !== "🔗 LINK") {
    hoja.getRange("J1").setValue("🔗 LINK");
    hoja.setColumnWidth(10, 250); 
  }
  
  // Abrió en columna 11 (K)
  if (valores[1] !== "✅ ABRIÓ") {
    hoja.getRange("K1").setValue("✅ ABRIÓ");
    hoja.setColumnWidth(11, 100);
  }
  
  // Estilo
  hoja.getRange("J1:K1")
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#0f172a")
    .setHorizontalAlignment("center");
}


// ==========================================
// 🔄 ACTUALIZAR METADATA
// ==========================================
function actualizarMetadataUltimoRegistro(vendedor, userAgent, ubicacion) {
  const hojaId = getConfig("google.hoja_tracking");
  const pestañaTracking = getConfig("google.pestana_tracking") || "TRACKING";
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName(pestañaTracking);
  
  const datos = hoja.getDataRange().getValues();
  let filaActualizar = null;
  
  // Busca de abajo hacia arriba
  for (let i = datos.length - 1; i >= 1; i--) { 
    if (datos[i][2] === vendedor) { 
      filaActualizar = i + 1;
      break;
    }
  }
  
  if (!filaActualizar) return;
  
  const metadata = parseUserAgent(userAgent);
  
  // Actualiza G, H, I
  hoja.getRange(filaActualizar, 7).setValue(metadata.dispositivo); // G
  hoja.getRange(filaActualizar, 8).setValue(metadata.navegador);   // H
  hoja.getRange(filaActualizar, 9).setValue(metadata.so);          // I
  
  // 🆕 NUEVO: Marcar como "Sí" en la columna K (Abrió)
  hoja.getRange(filaActualizar, 11).setValue("Sí");
  
  // Actualizar ubicación si existe
  if (ubicacion) {
    // Si hay una columna L para ubicación detallada, usar esa
    // Por ahora, la ubicación está en I (Sistema), así que se puede agregar
    // Opcional: agregar columna nueva o dejar en I
  }
  
  // 🆕 NUEVO: Actualizar formato condicional (ahora que sabemos que abrió)
  const tiempoReaccionStr = datos[filaActualizar - 1][5]; // Columna F
  const minutos = parsearTiempoReaccion(tiempoReaccionStr);
  aplicarFormatoCondicional(hoja, filaActualizar, minutos);
}


// ==========================================
// 🎨 FORMATO CONDICIONAL
// ==========================================
/**
 * Aplica color de fondo según tiempo de reacción
 * Verde: < 30 min
 * Amarillo: 30 min - 2 hs
 * Naranja: 2 hs - 24 hs
 * Rojo: > 24 hs o no abrió
 */
function aplicarFormatoCondicional(hoja, fila, minutos) {
  const rangoFila = hoja.getRange(fila, 1, 1, 11); // Columnas A-K
  
  let color = "#ffffff"; // Blanco por defecto
  
  if (minutos === 0) {
    // No se calculó tiempo (todavía no abrió)
    color = "#fecaca"; // Rojo claro
  } else if (minutos < 30) {
    color = "#d1fae5"; // Verde claro
  } else if (minutos < 120) { // 2 horas
    color = "#fef3c7"; // Amarillo claro
  } else if (minutos < 1440) { // 24 horas
    color = "#fed7aa"; // Naranja claro
  } else {
    color = "#fecaca"; // Rojo claro
  }
  
  rangoFila.setBackground(color);
}


// ==========================================
// 🎬 PANTALLA DE REDIRECCIÓN (LÓGICA V4.2)
// ==========================================
function generarPantallaRedirect(datos) {
  const urlWebApp = ScriptApp.getService().getUrl();
  
  // Limpiar nombre
  let nombreMostrar = datos.vendedor;
  if (nombreMostrar.includes("-")) {
    nombreMostrar = nombreMostrar.split("-")[1].trim();
  }

  const html = `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Acceso Reporte</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
      min-height: 100vh;
      display: flex; align-items: center; justify-content: center;
      margin: 0; padding: 20px; color: white;
    }
    .container {
      background: rgba(255, 255, 255, 0.1);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(255,255,255,0.2);
      border-radius: 24px;
      padding: 40px 30px;
      max-width: 400px;
      width: 100%;
      text-align: center;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }
    h1 { font-size: 26px; margin-bottom: 5px; font-weight: 700; }
    .report-name { font-size: 16px; color: #94a3b8; margin-bottom: 25px; }
    .report-name strong { color: #fff; display: block; margin-top: 5px; font-size: 18px;}
    
    .update-box {
      background: rgba(16, 185, 129, 0.2);
      border: 1px solid rgba(16, 185, 129, 0.4);
      color: #34d399;
      padding: 12px;
      border-radius: 12px;
      margin-bottom: 20px;
      font-size: 14px;
      display: flex; align-items: center; justify-content: center; gap: 8px;
    }
    
    .date-display {
      font-size: 14px; color: #cbd5e1; margin-bottom: 30px;
      background: rgba(0,0,0,0.3); padding: 8px 16px;
      border-radius: 50px; display: inline-block;
    }

    .btn-action {
      background: #3b82f6; color: white; border: none; width: 100%;
      padding: 16px; font-size: 16px; font-weight: 700;
      border-radius: 12px; cursor: pointer; transition: all 0.2s;
      box-shadow: 0 4px 12px rgba(59, 130, 246, 0.5);
      display: flex; justify-content: center; align-items: center;
    }
    .btn-action:hover { background: #2563eb; transform: translateY(-2px); }
    .btn-action:disabled { background: #475569; cursor: wait; transform: none; box-shadow: none;}
    
    .spinner { display: none; width: 18px; height: 18px; border: 3px solid #fff; border-top-color: transparent; border-radius: 50%; animation: spin 1s linear infinite; margin-right: 10px; }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div class="container">
    <h1>Hola, ${escapeHtml(nombreMostrar)}</h1>
    <p class="report-name">
      Reporte disponible:
      <strong>${escapeHtml(datos.archivo)}</strong>
    </p>

    <div class="update-box">
      <span>✅</span> Reporte Actualizado
    </div>

    <div class="date-display" id="reloj">
      Cargando hora...
    </div>
    
    <button id="btnEntrar" class="btn-action" onclick="iniciarProceso()">
      <span class="spinner" id="spinner"></span>
      <span id="btnText">VER REPORTE</span>
    </button>
  </div>
  
  <script>
    // 1. Mostrar Hora Local
    const ahora = new Date();
    const opciones = { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' };
    document.getElementById('reloj').innerText = "Acceso: " + ahora.toLocaleDateString('es-AR', opciones);

    // 2. PROCESO DE ANÁLISIS Y ENVÍO
    function iniciarProceso() {
      const btn = document.getElementById('btnEntrar');
      const spinner = document.getElementById('spinner');
      const btnText = document.getElementById('btnText');
      
      btn.disabled = true;
      spinner.style.display = 'block';
      btnText.innerText = "ABRIENDO...";

      const urlWebApp = "${urlWebApp}";
      const vendedor = "${escapeHtml(datos.vendedor)}";
      const ua = navigator.userAgent;
      
      // Geo simple (no necesitamos triangulación compleja para esta versión)
      const geo = "Detectado";
      
      const urlUpdate = urlWebApp + 
        "?action=update" +
        "&vendedor=" + encodeURIComponent(vendedor) +
        "&ua=" + encodeURIComponent(ua) +
        "&geo=" + encodeURIComponent(geo);

      fetch(urlUpdate, { mode: 'no-cors' })
        .then(() => redirigir())
        .catch(() => redirigir());

      setTimeout(redirigir, 2000);
    }

    let redirigiendo = false;
    function redirigir() {
      if (redirigiendo) return;
      redirigiendo = true;
      window.location.href = "${datos.url}";
    }
  </script>
</body>
</html>
  `;
  return HtmlService.createHtmlOutput(html)
    .setTitle("Acceso Reporte")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// UTILS
// ==========================================
function validarParametros(e) {
  const params = e.parameter || {};
  if (!params.url) return { valido: false, error: "URL no especificada" };
  let url = decodeURIComponent(params.url);
  if (!url.startsWith('https://')) return { valido: false, error: "URL inválida" };
  return {
    valido: true,
    datos: {
      url: url,
      vendedor: params.vendedor || "Desconocido",
      archivo: params.archivo || "Reporte",
      fechaEnvio: params.envio || ""
    }
  };
}

function crearHojaTrackingPro(ss, nombre) {
  const hoja = ss.insertSheet(nombre);
  hoja.getRange("A1:K1").setValues([[
    "📅 FECHA", "⏰ HORA", "👤 VENDEDOR", "📂 TIPO", 
    "📤 ENVIADO", "⏱️ REACCIÓN", 
    "📱 DISPOSITIVO", "🌐 NAVEGADOR", "💻 SISTEMA", 
    "🔗 LINK", "✅ ABRIÓ"
  ]]).setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
  hoja.setFrozenRows(1);
  
  // Anchos de columna
  hoja.setColumnWidth(11, 100); // K: Abrió
  
  return hoja;
}

function detectarTipoReporte(nombre) {
  const n = nombre.toUpperCase();
  if (n.includes("VENTA")) return "💰 VENTAS";
  if (n.includes("STOCK")) return "📦 STOCK";
  if (n.includes("CTA")) return "📉 CUENTAS";
  return "📄 OTROS";
}

function calcularTiempoReaccion(envio, ahora) {
  if(!envio) return { texto: "-", minutos: 0 };
  try {
     const partes = envio.split(' ');
     const f = partes[0].split('/');
     const h = partes[1].split(':');
     const fechaEnvio = new Date(f[2], f[1]-1, f[0], h[0], h[1]);
     const min = Math.floor((ahora - fechaEnvio)/60000);
     if (min < 60) return { minutos: min, texto: min + " min" };
     return { minutos: min, texto: Math.floor(min/60) + " hs" };
  } catch(e) { return { texto: "-", minutos: 0 }; }
}

function parsearTiempoReaccion(str) {
  if (!str || str === "-") return 0;
  
  const minMatch = str.match(/(\d+)\s*min/i);
  if (minMatch) return parseInt(minMatch[1], 10);
  
  const hsMatch = str.match(/(\d+)\s*hs/i);
  if (hsMatch) return parseInt(hsMatch[1], 10) * 60;
  
  return 0;
}

function parseUserAgent(ua) {
  let d = "PC", n = "Chrome", s = "Windows";
  const u = ua.toLowerCase();
  if(u.includes("iphone")) { d="Móvil"; s="iOS"; n="Safari"; }
  else if(u.includes("android")) { d="Móvil"; s="Android"; }
  return { dispositivo: d, navegador: n, so: s };
}

function escapeHtml(text) { return String(text).replace(/[&<>"']/g, ""); }
function mostrarPantallaError(e) { return HtmlService.createHtmlOutput(e); }
function getConfig(key) { return ""; }
