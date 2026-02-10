/***** Analytics.gs *****/

// ==========================================
// 📊 SISTEMA DE ANALYTICS AUTOMÁTICO
// ==========================================

/**
 * Actualiza todo el dashboard automáticamente
 * Se llama al finalizar distribución o manualmente
 */
function actualizarDashboard() {
  try {
    logAConsola("📊 Actualizando ANALYTICS_DASHBOARD...", "info");
    
    const inicio = Date.now();
    
    // Actualizar KPIs
    actualizarKPIs();
    
    // Actualizar gráficos
    actualizarGraficos();
    
    // Actualizar alertas
    actualizarAlertas();
    
    const duracion = ((Date.now() - inicio) / 1000).toFixed(2);
    logAConsola(`✅ Dashboard actualizado en ${duracion}s`, "ok");
    
  } catch (e) {
    logAConsola(`❌ Error actualizando dashboard: ${e}`, "error");
  }
}


/**
 * Actualiza los KPIs en la pestaña ANALYTICS_DASHBOARD
 */
function actualizarKPIs() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  let hoja = ss.getSheetByName("ANALYTICS_DASHBOARD");
  
  if (!hoja) {
    hoja = crearPestañaAnalytics();
  }
  
  // Calcular métricas
  const metricas = calcularMetricas();
  
  // Actualizar celdas de KPIs
  // Asumiendo que ya tienes una estructura, vamos a agregar en columnas B-C
  
  // Total Clics Hoy (ya existe en tu imagen, fila 3)
  hoja.getRange("C3").setValue(metricas.clicsHoy);
  
  // Tiempo Reacción Promedio (fila 4)
  hoja.getRange("C4").setValue(metricas.tiempoReaccionPromedio);
  
  // Vendedor Más Rápido (fila 5)
  hoja.getRange("C5").setValue(metricas.vendedorMasRapido);
  
  // % Móviles (fila 6)
  hoja.getRange("C6").setValue(metricas.porcentajeMoviles);
  
  // NUEVOS KPIs
  // Reportes del mes (fila 8)
  if (!hoja.getRange("B8").getValue()) {
    hoja.getRange("B8").setValue("📊 Reportes Distribuidos (Mes)");
  }
  hoja.getRange("C8").setValue(metricas.reportesMes);
  
  // Tasa de apertura (fila 9)
  if (!hoja.getRange("B9").getValue()) {
    hoja.getRange("B9").setValue("✅ Tasa de Apertura");
  }
  hoja.getRange("C9").setValue(metricas.tasaApertura);
  
  // Top 3 más rápidos (fila 11-13)
  if (!hoja.getRange("B11").getValue()) {
    hoja.getRange("B11").setValue("🏆 TOP 3 MÁS RÁPIDOS");
    hoja.getRange("B11").setFontWeight("bold").setBackground("#fef3c7");
  }
  
  metricas.top3Rapidos.forEach((vendedor, index) => {
    const fila = 12 + index;
    hoja.getRange(`B${fila}`).setValue(`${index + 1}. ${vendedor.nombre}`);
    hoja.getRange(`C${fila}`).setValue(vendedor.tiempo);
  });
  
  logAConsola("  -> KPIs actualizados", "info");
}


/**
 * Calcula todas las métricas necesarias
 */
function calcularMetricas() {
  const hojaId = getConfig("google.hoja_tracking");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaTracking = ss.getSheetByName("TRACKING");
  
  const metricas = {
    clicsHoy: 0,
    tiempoReaccionPromedio: "-",
    vendedorMasRapido: "-",
    porcentajeMoviles: "0%",
    reportesMes: 0,
    tasaApertura: "0%",
    top3Rapidos: []
  };
  
  if (!hojaTracking) return metricas;
  
  const datos = hojaTracking.getDataRange().getValues();
  const hoy = new Date();
  const inicioMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  
  let clicsHoy = 0;
  let clicsMes = 0;
  let abrieronMes = 0;
  let tiemposReaccion = [];
  let moviles = 0;
  const vendedoresTiempos = {};
  
  for (let i = 1; i < datos.length; i++) {
    try {
      const fechaStr = datos[i][0];
      if (!fechaStr) continue;
      
      const partes = fechaStr.split('/');
      if (partes.length !== 3) continue;
      
      const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      const vendedor = datos[i][2];
      const tiempoReaccionStr = datos[i][5];
      const dispositivo = datos[i][6];
      const abrio = datos[i][10]; // Columna K: ✅ ABRIÓ
      
      // Clics hoy
      if (fecha.toDateString() === hoy.toDateString()) {
        clicsHoy++;
      }
      
      // Clics del mes
      if (fecha >= inicioMes) {
        clicsMes++;
        
        // Tasa de apertura
        if (abrio === "Sí" || abrio === true) {
          abrieronMes++;
        }
      }
      
      // Tiempo de reacción
      if (tiempoReaccionStr && tiempoReaccionStr !== "-") {
        const minutos = parsearTiempoReaccion(tiempoReaccionStr);
        if (minutos > 0) {
          tiemposReaccion.push(minutos);
          
          // Por vendedor
          if (!vendedoresTiempos[vendedor]) {
            vendedoresTiempos[vendedor] = [];
          }
          vendedoresTiempos[vendedor].push(minutos);
        }
      }
      
      // % Móviles
      if (dispositivo && dispositivo.toLowerCase().includes("móvil")) {
        moviles++;
      }
      
    } catch (e) {
      // Ignorar errores de parsing
    }
  }
  
  // Calcular promedios
  metricas.clicsHoy = clicsHoy;
  metricas.reportesMes = clicsMes;
  
  if (tiemposReaccion.length > 0) {
    const promedio = tiemposReaccion.reduce((a, b) => a + b, 0) / tiemposReaccion.length;
    metricas.tiempoReaccionPromedio = formatearTiempo(promedio);
  }
  
  if (datos.length > 1) {
    metricas.porcentajeMoviles = ((moviles / (datos.length - 1)) * 100).toFixed(0) + "%";
  }
  
  if (clicsMes > 0) {
    metricas.tasaApertura = ((abrieronMes / clicsMes) * 100).toFixed(0) + "%";
  }
  
  // Top 3 más rápidos
  const vendedoresPromedio = [];
  for (const [nombre, tiempos] of Object.entries(vendedoresTiempos)) {
    const promedio = tiempos.reduce((a, b) => a + b, 0) / tiempos.length;
    vendedoresPromedio.push({ nombre, promedio });
  }
  
  vendedoresPromedio.sort((a, b) => a.promedio - b.promedio);
  
  metricas.top3Rapidos = vendedoresPromedio.slice(0, 3).map(v => ({
    nombre: v.nombre,
    tiempo: formatearTiempo(v.promedio)
  }));
  
  // Vendedor más rápido
  if (metricas.top3Rapidos.length > 0) {
    metricas.vendedorMasRapido = metricas.top3Rapidos[0].nombre;
  }
  
  return metricas;
}


/**
 * Parsea string de tiempo a minutos
 */
function parsearTiempoReaccion(str) {
  if (!str || str === "-") return 0;
  
  const minMatch = str.match(/(\d+)\s*min/i);
  if (minMatch) return parseInt(minMatch[1], 10);
  
  const hsMatch = str.match(/(\d+)\s*hs/i);
  if (hsMatch) return parseInt(hsMatch[1], 10) * 60;
  
  return 0;
}


/**
 * Formatea minutos a string legible
 */
function formatearTiempo(minutos) {
  if (minutos < 60) {
    return Math.round(minutos) + " min";
  }
  return (minutos / 60).toFixed(1) + " hs";
}


/**
 * Actualiza los gráficos en el dashboard
 */
function actualizarGraficos() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName("ANALYTICS_DASHBOARD");
  
  if (!hoja) return;
  
  // Limpiar gráficos existentes
  const charts = hoja.getCharts();
  charts.forEach(chart => hoja.removeChart(chart));
  
  // Crear área de datos para gráficos (columnas H-M)
  prepararDatosGraficos(hoja);
  
  // GRÁFICO 1: Reportes por semana (líneas)
  crearGraficoReportesPorSemana(hoja);
  
  // GRÁFICO 2: Tipos de reportes (torta)
  crearGraficoTiposReportes(hoja);
  
  logAConsola("  -> Gráficos actualizados", "info");
}


/**
 * Prepara los datos para los gráficos
 */
function prepararDatosGraficos(hoja) {
  const hojaId = getConfig("google.hoja_tracking");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaTracking = ss.getSheetByName("TRACKING");
  
  if (!hojaTracking) return;
  
  // DATOS PARA GRÁFICO DE LÍNEAS (Reportes por semana)
  hoja.getRange("H1").setValue("📊 DATOS - REPORTES POR SEMANA");
  hoja.getRange("H2:I2").setValues([["Semana", "Cantidad"]]);
  
  const reportesPorSemana = calcularReportesPorSemana(hojaTracking);
  
  let fila = 3;
  reportesPorSemana.forEach(item => {
    hoja.getRange(`H${fila}`).setValue(item.semana);
    hoja.getRange(`I${fila}`).setValue(item.cantidad);
    fila++;
  });
  
  // DATOS PARA GRÁFICO DE TORTA (Tipos de reportes)
  hoja.getRange("K1").setValue("📊 DATOS - TIPOS DE REPORTES");
  hoja.getRange("K2:L2").setValues([["Tipo", "Cantidad"]]);
  
  const tiposReportes = calcularTiposReportes(hojaTracking);
  
  fila = 3;
  for (const [tipo, cantidad] of Object.entries(tiposReportes)) {
    hoja.getRange(`K${fila}`).setValue(tipo);
    hoja.getRange(`L${fila}`).setValue(cantidad);
    fila++;
  }
}


/**
 * Calcula reportes agrupados por semana (últimas 4 semanas)
 */
function calcularReportesPorSemana(hojaTracking) {
  const datos = hojaTracking.getDataRange().getValues();
  const hoy = new Date();
  const semanas = [];
  
  // Últimas 4 semanas
  for (let i = 3; i >= 0; i--) {
    const inicioSemana = new Date(hoy.getTime() - i * 7 * 24 * 60 * 60 * 1000);
    semanas.push({
      semana: `Sem ${4 - i}`,
      inicio: inicioSemana,
      fin: new Date(inicioSemana.getTime() + 7 * 24 * 60 * 60 * 1000),
      cantidad: 0
    });
  }
  
  // Contar reportes por semana
  for (let i = 1; i < datos.length; i++) {
    try {
      const fechaStr = datos[i][0];
      if (!fechaStr) continue;
      
      const partes = fechaStr.split('/');
      if (partes.length !== 3) continue;
      
      const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      
      semanas.forEach(sem => {
        if (fecha >= sem.inicio && fecha < sem.fin) {
          sem.cantidad++;
        }
      });
    } catch (e) {}
  }
  
  return semanas;
}


/**
 * Calcula cantidad de reportes por tipo
 */
function calcularTiposReportes(hojaTracking) {
  const datos = hojaTracking.getDataRange().getValues();
  const tipos = {
    "Ventas": 0,
    "Cuentas": 0,
    "Otros": 0
  };
  
  for (let i = 1; i < datos.length; i++) {
    const tipo = datos[i][3]; // Columna D: Tipo
    
    if (!tipo) continue;
    
    if (tipo.includes("VENTAS")) {
      tipos["Ventas"]++;
    } else if (tipo.includes("CUENTAS")) {
      tipos["Cuentas"]++;
    } else {
      tipos["Otros"]++;
    }
  }
  
  return tipos;
}


/**
 * Crea gráfico de líneas para reportes por semana
 */
function crearGraficoReportesPorSemana(hoja) {
  const rango = hoja.getRange("H2:I6"); // 4 semanas + header
  
  const chart = hoja.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rango)
    .setPosition(16, 1, 0, 0) // Fila 16, columna 1
    .setOption('title', '📈 Reportes Distribuidos (Últimas 4 Semanas)')
    .setOption('width', 600)
    .setOption('height', 300)
    .setOption('legend', { position: 'bottom' })
    .setOption('curveType', 'function')
    .setOption('colors', ['#3b82f6'])
    .build();
  
  hoja.insertChart(chart);
}


/**
 * Crea gráfico de torta para tipos de reportes
 */
function crearGraficoTiposReportes(hoja) {
  const rango = hoja.getRange("K2:L5"); // 3 tipos + header
  
  const chart = hoja.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(rango)
    .setPosition(16, 7, 0, 0) // Fila 16, columna 7
    .setOption('title', '📊 Distribución por Tipo de Reporte')
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('pieSliceText', 'value')
    .setOption('colors', ['#10b981', '#f59e0b', '#6366f1'])
    .build();
  
  hoja.insertChart(chart);
}


/**
 * Actualiza la sección de alertas
 */
function actualizarAlertas() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hoja = ss.getSheetByName("ANALYTICS_DASHBOARD");
  
  if (!hoja) return;
  
  // Detectar alertas
  const alertas = detectarAlertas();
  
  // Escribir en fila 28 (debajo de los gráficos)
  if (!hoja.getRange("B28").getValue()) {
    hoja.getRange("B28").setValue("⚠️ ALERTAS ACTIVAS");
    hoja.getRange("B28").setFontWeight("bold").setFontSize(12).setBackground("#fef3c7");
  }
  
  // Limpiar alertas anteriores (filas 29-35)
  hoja.getRange("B29:F35").clearContent();
  
  let fila = 29;
  if (alertas.length === 0) {
    hoja.getRange(`B${fila}`).setValue("🟢 Sistema operando normalmente");
    hoja.getRange(`B${fila}`).setBackground("#d1fae5").setFontColor("#065f46");
  } else {
    alertas.forEach(alerta => {
      hoja.getRange(`B${fila}`).setValue(alerta.icono + " " + alerta.mensaje);
      hoja.getRange(`B${fila}`).setBackground(alerta.color).setFontColor("#1e293b");
      fila++;
    });
  }
  
  logAConsola("  -> Alertas actualizadas", "info");
}


/**
 * Detecta alertas del sistema
 */
function detectarAlertas() {
  const alertas = [];
  
  // Alerta 1: Vendedores sin abrir reportes en 24hs
  const sinAbrir = detectarVendedoresSinAbrir();
  if (sinAbrir > 0) {
    alertas.push({
      icono: "🔴",
      mensaje: `${sinAbrir} vendedor(es) no abrieron reportes en las últimas 24hs`,
      color: "#fecaca"
    });
  }
  
  // Alerta 2: Errores de distribución hoy
  const erroresHoy = contarErroresHoy();
  if (erroresHoy > 2) {
    alertas.push({
      icono: "🟡",
      mensaje: `${erroresHoy} error(es) de distribución detectados hoy`,
      color: "#fef3c7"
    });
  }
  
  // Alerta 3: Vendedores inactivos (sin clicks en 7 días)
  const inactivos = detectarVendedoresInactivos();
  if (inactivos > 0) {
    alertas.push({
      icono: "🟠",
      mensaje: `${inactivos} vendedor(es) sin actividad en 7 días`,
      color: "#fed7aa"
    });
  }
  
  return alertas;
}


/**
 * Detecta vendedores que no abrieron reportes en 24hs
 */
function detectarVendedoresSinAbrir() {
  const hojaId = getConfig("google.hoja_tracking");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaTracking = ss.getSheetByName("TRACKING");
  
  if (!hojaTracking) return 0;
  
  const datos = hojaTracking.getDataRange().getValues();
  const hace24hs = new Date(Date.now() - 24 * 60 * 60 * 1000);
  
  let sinAbrir = 0;
  
  for (let i = 1; i < datos.length; i++) {
    try {
      const fechaStr = datos[i][0];
      if (!fechaStr) continue;
      
      const partes = fechaStr.split('/');
      if (partes.length !== 3) continue;
      
      const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      const abrio = datos[i][10]; // Columna K
      
      if (fecha >= hace24hs && (abrio === "No" || abrio === false || !abrio)) {
        sinAbrir++;
      }
    } catch (e) {}
  }
  
  return sinAbrir;
}


/**
 * Cuenta errores en la consola de hoy
 */
function contarErroresHoy() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaConsola = ss.getSheetByName("Consola");
  
  if (!hojaConsola) return 0;
  
  const datos = hojaConsola.getDataRange().getValues();
  const hoy = new Date().toDateString();
  
  let errores = 0;
  
  for (let i = 1; i < datos.length; i++) {
    try {
      const fecha = datos[i][0];
      if (fecha instanceof Date && fecha.toDateString() === hoy) {
        const nivel = datos[i][1];
        if (nivel && nivel.toString().includes("ERROR")) {
          errores++;
        }
      }
    } catch (e) {}
  }
  
  return errores;
}


/**
 * Detecta vendedores sin actividad en 7 días
 */
function detectarVendedoresInactivos() {
  const hojaId = getConfig("google.hoja_tracking");
  const ss = SpreadsheetApp.openById(hojaId);
  const hojaTracking = ss.getSheetByName("TRACKING");
  
  if (!hojaTracking) return 0;
  
  const mapeo = getMapeoFiltrado_();
  const datos = hojaTracking.getDataRange().getValues();
  const hace7dias = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  
  const vendedoresActivos = new Set();
  
  for (let i = 1; i < datos.length; i++) {
    try {
      const fechaStr = datos[i][0];
      if (!fechaStr) continue;
      
      const partes = fechaStr.split('/');
      if (partes.length !== 3) continue;
      
      const fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      
      if (fecha >= hace7dias) {
        vendedoresActivos.add(datos[i][2]);
      }
    } catch (e) {}
  }
  
  return mapeo.length - vendedoresActivos.size;
}


/**
 * Crea la pestaña ANALYTICS_DASHBOARD si no existe
 */
function crearPestañaAnalytics() {
  const hojaId = getConfig("google.hoja_mapeo");
  const ss = SpreadsheetApp.openById(hojaId);
  
  let hoja = ss.getSheetByName("ANALYTICS_DASHBOARD");
  if (hoja) return hoja;
  
  hoja = ss.insertSheet("ANALYTICS_DASHBOARD");
  
  // Configuración básica
  hoja.getRange("A1:F1").merge();
  hoja.getRange("A1")
    .setValue("📊 ANALYTICS - Sistema de Tracking")
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center");
  
  logAConsola("✅ Pestaña ANALYTICS_DASHBOARD creada", "ok");
  return hoja;
}


/**
 * Función para botón: Actualizar Dashboard manualmente
 */
function ANALYTICS_actualizarDashboard() {
  actualizarDashboard();
  
  SpreadsheetApp.getUi().alert(
    'Dashboard Actualizado',
    'Los KPIs, gráficos y alertas han sido actualizados con los datos más recientes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
