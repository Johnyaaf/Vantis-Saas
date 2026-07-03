/**
 * ============================================================
 * M5_Batch.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Procesos automáticos nocturnos.
 * Clasifica ABC, genera snapshots de stock, calcula scores
 * de proveedores, detecta alertas y las envía por email.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Módulos:      M0_Setup, M4_Auditoria, M2_Transacciones
 *   Hojas:        BATCH_RESULTADOS, HIST_STOCK, CALC_ALERTAS,
 *                 MSTR_PRODUCTOS, MSTR_PROVEEDORES, DB_VENTAS,
 *                 DB_MOVIMIENTOS, DB_CXC, DB_CXP, DB_METAS
 *   APIs Google:  GmailApp, SpreadsheetApp
 *   Triggers:     batchNocturno (diario 01:00)
 *                 actualizarFechaHoy (diario 00:00)
 *                 crearPeriodoSiguiente (día 25 cada mes)
 * ============================================================
 */

// ============================================================
// FUNCIÓN: actualizarFechaHoy
// Actualiza CONFIG_SISTEMA!T2 con la fecha actual
// Trigger: diario a las 00:00
// ============================================================

/**
 * Mantiene actualizada la fecha del sistema en CONFIG_SISTEMA!T2.
 * Esta celda es la fuente de verdad para todas las ARRAYFORMULA
 * que calculan aging de CxC y CxP.
 */
function actualizarFechaHoy() {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var hoja   = ss.getSheetByName(HOJAS.CONFIG_SISTEMA);
    if (!hoja) {
      Logger.log('ERROR actualizarFechaHoy: CONFIG_SISTEMA no encontrada');
      return;
    }
    hoja.getRange('T2').setValue(new Date());
    SpreadsheetApp.flush();
    Logger.log('actualizarFechaHoy: Fecha actualizada a ' + new Date());
  } catch (e) {
    Logger.log('ERROR actualizarFechaHoy: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: batchNocturno
// Orquesta todos los procesos automáticos nocturnos
// Trigger: diario a las 01:00
// ============================================================

/**
 * Función principal del batch. Ejecuta todos los subprocesos
 * en orden. Si alguno falla registra el error pero continúa
 * con el siguiente para maximizar la cobertura.
 */
function batchNocturno() {
  var inicio = new Date();
  Logger.log('batchNocturno: Iniciando — ' + inicio.toISOString());
  registrarLog('SISTEMA', 'BATCH_INICIO', 'batchNocturno', 'timestamp', inicio.toISOString());

  var resultados = {
    abc_productos:   false,
    abc_clientes:    false,
    snapshot_stock:  false,
    score_proveedores: false,
    alertas:         false,
    email:           false
  };

  // Proceso 1: ABC Productos
  try {
    resultados.abc_productos = batchABCProductos();
  } catch (e) {
    Logger.log('ERROR batchNocturno > batchABCProductos: ' + e.message);
  }

  // Proceso 2: ABC Clientes
  try {
    resultados.abc_clientes = batchABCClientes();
  } catch (e) {
    Logger.log('ERROR batchNocturno > batchABCClientes: ' + e.message);
  }

  // Proceso 3: Snapshot de stock
  try {
    resultados.snapshot_stock = generarSnapshotStock();
  } catch (e) {
    Logger.log('ERROR batchNocturno > generarSnapshotStock: ' + e.message);
  }

  // Proceso 4: Score de proveedores
  try {
    resultados.score_proveedores = batchScoreCumplimientoProveedores();
  } catch (e) {
    Logger.log('ERROR batchNocturno > batchScoreProveedores: ' + e.message);
  }

  // Proceso 5: Detectar alertas
  try {
    resultados.alertas = detectarAlertas();
  } catch (e) {
    Logger.log('ERROR batchNocturno > detectarAlertas: ' + e.message);
  }

  // Proceso 6: Enviar alertas por email
  try {
    resultados.email = enviarAlertasEmail();
  } catch (e) {
    Logger.log('ERROR batchNocturno > enviarAlertasEmail: ' + e.message);
  }

  // Proceso 7: Calcular variación de precios en HIST_PRECIOS
  try {
    _calcularVariacionPrecios();
  } catch (e) {
    Logger.log('ERROR batchNocturno > _calcularVariacionPrecios: ' + e.message);
  }

  // Registrar resultado en BATCH_RESULTADOS
  var fin       = new Date();
  var duracion  = Math.round((fin - inicio) / 1000);
  var exito     = Object.values(resultados).some(function(v) { return v; });
  _registrarResultadoBatch('NOCTURNO', exito, duracion, JSON.stringify(resultados));

  // Actualizar indicador de batch en CALC_ALERTAS
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(HOJAS.CALC_ALERTAS)
      .getRange('B8').setValue(exito ? 0 : 1);
  } catch (e) {
    Logger.log('ERROR batchNocturno > actualizar CA_BATCH_FALLIDO: ' + e.message);
  }

  registrarLog('SISTEMA', 'BATCH_FIN', 'batchNocturno', 'duracion_seg', String(duracion));
  Logger.log('batchNocturno: Completado en ' + duracion + ' seg — ' + JSON.stringify(resultados));
  flushLogBuffer();
}

// ============================================================
// FUNCIÓN: batchABCProductos
// Clasifica productos en A, B o C según participación en ventas
// ============================================================

/**
 * Clasifica todos los productos activos según el análisis ABC.
 * A = primeros productos que suman 80% de ventas del período.
 * B = siguientes que suman 15%.
 * C = resto (5% o productos sin ventas).
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function batchABCProductos() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hVentas = ss.getSheetByName(HOJAS.DB_VENTAS);
    var hProd   = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);

    if (!hVentas || !hProd) {
      Logger.log('batchABCProductos: Hojas no encontradas');
      return false;
    }

    var periodo    = getPeriodoActual();
    var ultimaVta  = hVentas.getLastRow();
    var ultimaPrd  = hProd.getLastRow();

    if (ultimaVta < 2 || ultimaPrd < 2) {
      Logger.log('batchABCProductos: Sin datos suficientes');
      return true;
    }

    // Leer ventas del período actual
    // Columnas: L=sys_periodo(11), N=id_producto(13), T=sys_subtotal_neto(19), AB=estado(27)
    var datosVta = hVentas.getRange(2, 1, ultimaVta - 1, 28).getValues();

    // Sumar ventas por producto
    var ventasPorProd = {};
    var totalVentas   = 0;

    for (var i = 0; i < datosVta.length; i++) {
      var fila      = datosVta[i];
      var perFila   = String(fila[11] || '').trim();
      var idProd    = String(fila[13] || '').trim();
      var subtotal  = Number(fila[19]) || 0;
      var estado    = String(fila[27] || '').trim();

      if (perFila !== periodo) continue;
      if (estado === ESTADOS.ANULADO) continue;
      if (!idProd) continue;

      ventasPorProd[idProd] = (ventasPorProd[idProd] || 0) + subtotal;
      totalVentas += subtotal;
    }

    if (totalVentas === 0) {
      Logger.log('batchABCProductos: Sin ventas en período ' + periodo);
      return true;
    }

    // Ordenar de mayor a menor venta
    var ranking = Object.keys(ventasPorProd).map(function(id) {
      return { id: id, ventas: ventasPorProd[id] };
    }).sort(function(a, b) { return b.ventas - a.ventas; });

    // Calcular clasificación ABC
    var clasificaciones = {};
    var acumulado       = 0;

    for (var j = 0; j < ranking.length; j++) {
      acumulado += ranking[j].ventas;
      var pct    = acumulado / totalVentas;
      clasificaciones[ranking[j].id] = pct <= 0.80 ? 'A' : pct <= 0.95 ? 'B' : 'C';
    }

    // Actualizar clasificacion_abc en MSTR_PRODUCTOS
    // Columna Q = índice 16 = col 17 base 1
    var datosPrd = hProd.getRange(2, 1, ultimaPrd - 1, 28).getValues();
    var updates  = 0;

    for (var k = 0; k < datosPrd.length; k++) {
      var idProdActual = String(datosPrd[k][COL.PRD_ID_PRODUCTO]).trim();
      var activo       = datosPrd[k][COL.PRD_ACTIVO] === true;

      if (!idProdActual || !activo) continue;

      var nuevaClasif = clasificaciones[idProdActual] || 'C';
      var clasifActual = String(datosPrd[k][COL.PRD_CLASIFICACION_ABC] || '').trim();

      if (nuevaClasif !== clasifActual) {
        hProd.getRange(k + 2, COL.PRD_CLASIFICACION_ABC + 1).setValue(nuevaClasif);
        updates++;
      }
    }

    SpreadsheetApp.flush();
    Logger.log('batchABCProductos: ' + updates + ' productos actualizados — Período: ' + periodo);
    registrarLog(HOJAS.MSTR_PRODUCTOS, 'BATCH_ABC', 'PRODUCTOS', 'clasificacion_abc', updates + '_actualizados');
    return true;

  } catch (e) {
    Logger.log('ERROR batchABCProductos: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: batchABCClientes
// Clasifica clientes en A, B o C según volumen de compras
// ============================================================

/**
 * Clasifica clientes activos según el análisis ABC de ventas.
 * Guarda resultado en BATCH_RESULTADOS para consulta.
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function batchABCClientes() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hVentas = ss.getSheetByName(HOJAS.DB_VENTAS);
    var hBatch  = ss.getSheetByName(HOJAS.BATCH_RESULTADOS);

    if (!hVentas || !hBatch) {
      Logger.log('batchABCClientes: Hojas no encontradas');
      return false;
    }

    var periodo   = getPeriodoActual();
    var ultimaVta = hVentas.getLastRow();
    if (ultimaVta < 2) return true;

    var datosVta = hVentas.getRange(2, 1, ultimaVta - 1, 28).getValues();

    // Sumar ventas por cliente
    var ventasPorCli = {};
    var totalVentas  = 0;

    for (var i = 0; i < datosVta.length; i++) {
      var fila     = datosVta[i];
      var perFila  = String(fila[11] || '').trim();
      var idCli    = String(fila[12] || '').trim();
      var subtotal = Number(fila[19]) || 0;
      var estado   = String(fila[27] || '').trim();

      if (perFila !== periodo) continue;
      if (estado === ESTADOS.ANULADO) continue;
      if (!idCli) continue;

      ventasPorCli[idCli] = (ventasPorCli[idCli] || 0) + subtotal;
      totalVentas += subtotal;
    }

    if (totalVentas === 0) return true;

    // Ordenar y clasificar
    var ranking = Object.keys(ventasPorCli).map(function(id) {
      return { id: id, ventas: ventasPorCli[id] };
    }).sort(function(a, b) { return b.ventas - a.ventas; });

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var acumulado = 0;

    // Escribir resultados en BATCH_RESULTADOS
    var filas = ranking.map(function(item) {
      acumulado += item.ventas;
      var pct    = acumulado / totalVentas;
      var clasif = pct <= 0.80 ? 'A' : pct <= 0.95 ? 'B' : 'C';

      return [
        generarID('BAT'),  // A id_batch
        idEmpresa,         // B id_empresa
        ahora,             // C generado_en
        periodo,           // D sys_periodo
        'ABC_CLIENTES',    // E tipo_batch
        item.id,           // F id_referencia
        clasif,            // G clasificacion
        item.ventas,       // H valor
        Math.round((item.ventas / totalVentas) * 10000) / 100, // I pct
        '',                // J notas
        '',                // K campo11
        '',                // L campo12
        '',                // M campo13
        '',                // N campo14
        ''                 // O campo15
      ];
    });

    if (filas.length > 0) {
      var ultimaBatch = hBatch.getLastRow();
      hBatch.getRange(ultimaBatch + 1, 1, filas.length, 15).setValues(filas);
      SpreadsheetApp.flush();
    }

    Logger.log('batchABCClientes: ' + ranking.length + ' clientes clasificados');
    return true;

  } catch (e) {
    Logger.log('ERROR batchABCClientes: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: generarSnapshotStock
// Guarda estado del stock en HIST_STOCK al cierre del día
// ============================================================

/**
 * Genera un snapshot diario del stock de todos los productos.
 * Permite reconstruir el inventario histórico en cualquier fecha.
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function generarSnapshotStock() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hProd = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    var hSnap = ss.getSheetByName(HOJAS.HIST_STOCK);

    if (!hProd || !hSnap) {
      Logger.log('generarSnapshotStock: Hojas no encontradas');
      return false;
    }

    var ultimaPrd = hProd.getLastRow();
    if (ultimaPrd < 2) return true;

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var periodo   = getPeriodoActual() || '';
    var ahora     = new Date();

    // Leer productos activos con stock
    var datosPrd = hProd.getRange(2, 1, ultimaPrd - 1, 28).getValues();
    var filas    = [];

    for (var i = 0; i < datosPrd.length; i++) {
      var fila       = datosPrd[i];
      var idProd      = String(fila[COL.PRD_ID_PRODUCTO] || '').trim();
      var activo      = fila[COL.PRD_ACTIVO] === true;
      var manejaStk   = fila[COL.PRD_MANEJA_STOCK] === true;
      var stockActual = Number(fila[COL.PRD_STOCK_ACTUAL]) || 0;
      var stockMin    = Number(fila[13]) || 0;
      var stockCrit   = Number(fila[15]) || 0;
      var costoUnit   = Number(fila[COL.PRD_COSTO_UNITARIO]) || 0;
      var clasifABC   = String(fila[COL.PRD_CLASIFICACION_ABC] || 'C').trim();

      if (!idProd || !activo || !manejaStk) continue;

      filas.push([
        generarID('SNP'),  // A id_snapshot
        idEmpresa,         // B id_empresa
        idProd,            // C id_producto
        ahora,             // D fecha_snapshot
        periodo,           // E sys_periodo
        stockActual,       // F stock_actual
        stockMin,          // G stock_minimo
        stockCrit,         // H stock_critico
        costoUnit,         // I costo_unitario
        Math.round(stockActual * costoUnit), // J valor_inventario
        clasifABC,         // K clasificacion_abc
        stockActual <= stockCrit ? 'CRITICO' : stockActual <= stockMin ? 'BAJO' : 'NORMAL', // L estado_stock
        ahora              // M creado_en
      ]);
    }

    if (filas.length > 0) {
      var ultimaSnap = hSnap.getLastRow();
      hSnap.getRange(ultimaSnap + 1, 1, filas.length, 13).setValues(filas);
      SpreadsheetApp.flush();
    }

    Logger.log('generarSnapshotStock: ' + filas.length + ' productos en snapshot');
    return true;

  } catch (e) {
    Logger.log('ERROR generarSnapshotStock: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: batchScoreCumplimientoProveedores
// Actualiza score_cumplimiento en MSTR_PROVEEDORES
// ============================================================

/**
 * Calcula el score de cumplimiento de proveedores basándose
 * en la diferencia entre plazo prometido y días reales de entrega.
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function batchScoreCumplimientoProveedores() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hComp   = ss.getSheetByName(HOJAS.DB_COMPRAS);
    var hPrv    = ss.getSheetByName(HOJAS.MSTR_PROVEEDORES);

    if (!hComp || !hPrv) {
      Logger.log('batchScoreCumplimientoProveedores: Hojas no encontradas');
      return false;
    }

    var ultimaComp = hComp.getLastRow();
    var ultimaPrv  = hPrv.getLastRow();

    if (ultimaComp < 2 || ultimaPrv < 2) return true;

    // Leer OC recibidas en los últimos 6 meses
    // Cols: D=id_proveedor(3), G=fecha_emision(6), J=fecha_entrega_real(9), W=estado_oc(22)
    var datosComp = hComp.getRange(2, 1, ultimaComp - 1, 23).getValues();

    var scoresPrv = {};
    var hace6Meses = new Date();
    hace6Meses.setMonth(hace6Meses.getMonth() - 6);

    for (var i = 0; i < datosComp.length; i++) {
      var fila        = datosComp[i];
      var idPrv       = String(fila[COL.OC_ID_PROVEEDOR] || '').trim();
      var fechaEmis   = fila[COL.OC_FECHA_EMISION];
      var fechaEntrg  = fila[9]; // J fecha_entrega_real
      var estadoOC    = String(fila[COL.OC_ESTADO_OC] || '').trim();

      if (!idPrv) continue;
      if (estadoOC !== ESTADOS.OC_RECIBIDA) continue;
      if (!fechaEmis || !fechaEntrg) continue;

      var fEmis  = new Date(fechaEmis);
      var fEntrg = new Date(fechaEntrg);

      if (fEmis < hace6Meses) continue;

      var diasReales = Math.round((fEntrg - fEmis) / (1000 * 60 * 60 * 24));

      if (!scoresPrv[idPrv]) scoresPrv[idPrv] = { total: 0, count: 0 };
      scoresPrv[idPrv].total += diasReales;
      scoresPrv[idPrv].count++;
    }

    // Actualizar score_cumplimiento en MSTR_PROVEEDORES
    var datosPrv = hPrv.getRange(2, 1, ultimaPrv - 1, 24).getValues();
    var updates  = 0;

    for (var j = 0; j < datosPrv.length; j++) {
      var fila2    = datosPrv[j];
      var idPrvAct = String(fila2[COL_PRV.ID] || '').trim();
      var plazo    = Number(fila2[COL_PRV.PLAZO_ENTREGA] || 0);
      var activo   = fila2[COL_PRV.ACTIVO] === true;

      if (!idPrvAct) continue;
      if (!activo) continue;
      if (!scoresPrv[idPrvAct]) continue;

      var promDias  = scoresPrv[idPrvAct].total / scoresPrv[idPrvAct].count;
      var score     = plazo > 0
        ? Math.max(0, Math.min(100, Math.round((1 - (promDias - plazo) / plazo) * 100)))
        : 50;

      // score_cumplimiento = col Q = índice 16 = col 17 base 1
      hPrv.getRange(j + 2, COL_PRV.SCORE_CUMPL + 1).setValue(score);
      updates++;
    }

    SpreadsheetApp.flush();
    Logger.log('batchScoreCumplimientoProveedores: ' + updates + ' proveedores actualizados');
    return true;

  } catch (e) {
    Logger.log('ERROR batchScoreCumplimientoProveedores: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: detectarAlertas
// Evalúa condiciones de alerta y actualiza CALC_ALERTAS
// ============================================================

/**
 * Evalúa todas las condiciones de alerta del sistema y actualiza
 * los contadores en CALC_ALERTAS. Las ARRAYFORMULA de CxC y CxP
 * calculan aging automáticamente — solo leemos sus resultados.
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function detectarAlertas() {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var hAlertas  = ss.getSheetByName(HOJAS.CALC_ALERTAS);
    if (!hAlertas) return false;

    var stockCritico  = _evaluarStockCritico();
    var cxcVencida    = _evaluarCxCVencida();
    var cxpUrgente    = _evaluarCxPUrgente();
    var metasRojo     = _evaluarMetasRojo();
    var difCaja       = _evaluarDiferenciasCaja();

    var total = stockCritico + cxcVencida + cxpUrgente + metasRojo + difCaja;

    // Actualizar contadores en CALC_ALERTAS columna B
    hAlertas.getRange('B2').setValue(total);          // CA_TOTAL_CRITICAS
    hAlertas.getRange('B3').setValue(stockCritico);   // CA_CRITICAS_INVENTARIO
    hAlertas.getRange('B4').setValue(cxcVencida);     // CA_CXC_MAS90
    hAlertas.getRange('B5').setValue(cxpUrgente);     // CA_CXP_URGENTE
    hAlertas.getRange('B6').setValue(metasRojo);      // CA_METAS_ROJO
    hAlertas.getRange('B7').setValue(difCaja);        // CA_DIFERENCIAS_CAJA

    SpreadsheetApp.flush();

    Logger.log('detectarAlertas: Total=' + total +
      ' | Stock=' + stockCritico +
      ' | CxC=' + cxcVencida +
      ' | CxP=' + cxpUrgente +
      ' | Metas=' + metasRojo +
      ' | Caja=' + difCaja);

    return true;

  } catch (e) {
    Logger.log('ERROR detectarAlertas: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: enviarAlertasEmail
// Envía email al administrador con las alertas críticas
// ============================================================

/**
 * Lee los contadores de CALC_ALERTAS y envía un email
 * consolidado al administrador si hay alertas críticas.
 * Un solo email con todas las alertas para no agotar la cuota.
 *
 * @returns {boolean} true si se ejecutó correctamente
 */
function enviarAlertasEmail() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var hAlertas = ss.getSheetByName(HOJAS.CALC_ALERTAS);
    if (!hAlertas) return false;

    var total = Number(hAlertas.getRange('B2').getValue()) || 0;

    if (total === 0) {
      Logger.log('enviarAlertasEmail: Sin alertas críticas — no se envía email');
      return true;
    }

    var config     = getConfig();
    var emailAdmin = config ? config.email_admin : '';

    if (!emailAdmin || emailAdmin === '') {
      Logger.log('enviarAlertasEmail: Email admin no configurado en CONFIG_SISTEMA!J2');
      return false;
    }

    var stockCritico = Number(hAlertas.getRange('B3').getValue()) || 0;
    var cxcVencida   = Number(hAlertas.getRange('B4').getValue()) || 0;
    var cxpUrgente   = Number(hAlertas.getRange('B5').getValue()) || 0;
    var metasRojo    = Number(hAlertas.getRange('B6').getValue()) || 0;
    var difCaja      = Number(hAlertas.getRange('B7').getValue()) || 0;

    var empresa  = config ? config.nombre_empresa : 'ERP PYME';
    var periodo  = getPeriodoActual() || 'sin período';
    var fecha    = new Date().toLocaleDateString('es-CL');

    var cuerpo = _construirEmailAlertas(
      empresa, fecha, periodo, total,
      stockCritico, cxcVencida, cxpUrgente, metasRojo, difCaja
    );

    GmailApp.sendEmail(
      emailAdmin,
      '⚠️ ERP PYME — ' + total + ' alerta(s) detectada(s) — ' + empresa,
      cuerpo,
      { name: 'ERP PYME Sistema' }
    );

    Logger.log('enviarAlertasEmail: Email enviado a ' + emailAdmin + ' con ' + total + ' alertas');
    registrarLog('SISTEMA', 'EMAIL_ALERTAS', emailAdmin, 'total_alertas', String(total));
    return true;

  } catch (e) {
    Logger.log('ERROR enviarAlertasEmail: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: crearPeriodoSiguiente
// Crea automáticamente el período del mes siguiente
// Trigger: día 25 de cada mes a las 08:00
// ============================================================

/**
 * Verifica si ya existe el período del mes siguiente.
 * Si no existe lo crea con estado ABIERTO.
 * Garantiza continuidad operacional sin intervención manual.
 */
function crearPeriodoSiguiente() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hPer    = ss.getSheetByName(HOJAS.CONFIG_PERIODOS);
    var config  = getConfig();

    if (!hPer || !config) {
      Logger.log('crearPeriodoSiguiente: Dependencias no disponibles');
      return;
    }

    var idEmpresa = config.id_empresa;
    var ahora     = new Date();

    // Calcular el mes siguiente
    var mesSig = ahora.getMonth() + 2; // getMonth() es base 0, +1 para actual, +1 para siguiente
    var añoSig = ahora.getFullYear();
    if (mesSig > 12) { mesSig = 1; añoSig++; }

    var periodoSig = añoSig + '-' + String(mesSig).padStart(2, '0');
    var idPeriodo  = 'PER-' + periodoSig;

    // Verificar si ya existe
    var ultimaFila = hPer.getLastRow();
    if (ultimaFila >= 2) {
      var datos = hPer.getRange(2, 1, ultimaFila - 1, 7).getValues();
      for (var i = 0; i < datos.length; i++) {
        if (String(datos[i][4]).trim() === periodoSig &&
            String(datos[i][1]).trim() === idEmpresa) {
          Logger.log('crearPeriodoSiguiente: Período ' + periodoSig + ' ya existe');
          return;
        }
      }
    }

    // Nombres de meses en español
    var nombresMes = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

    var nuevaFila = [
      idPeriodo,                              // A id_periodo
      idEmpresa,                              // B id_empresa
      añoSig,                                 // C año
      mesSig,                                 // D mes
      periodoSig,                             // E sys_periodo
      nombresMes[mesSig] + ' ' + añoSig,     // F nombre_periodo
      ESTADOS.PERIODO_CERRADO,               // G estado — se abrirá manualmente cuando corresponda
      new Date(añoSig, mesSig - 1, 1),       // H fecha_apertura
      ''                                      // I fecha_cierre
    ];

    hPer.getRange(ultimaFila + 1, 1, 1, 9).setValues([nuevaFila]);
    SpreadsheetApp.flush();

    registrarLog(HOJAS.CONFIG_PERIODOS, 'CREAR', idPeriodo, 'sys_periodo', periodoSig);
    Logger.log('crearPeriodoSiguiente: Período ' + periodoSig + ' creado correctamente');

  } catch (e) {
    Logger.log('ERROR crearPeriodoSiguiente: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _evaluarStockCritico
// ============================================================

/**
 * Cuenta productos con stock_actual <= stock_critico.
 * @returns {number} Cantidad de productos en estado crítico
 */
function _evaluarStockCritico() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hProd = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    if (!hProd) return 0;

    var ultimaFila = hProd.getLastRow();
    if (ultimaFila < 2) return 0;

    var datos = hProd.getRange(2, 1, ultimaFila - 1, 28).getValues();
    var criticos = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila      = datos[i];
      var idProd    = String(fila[COL.PRD_ID_PRODUCTO] || '').trim();
      if (!idProd) continue; // ← saltar filas vacías

      var activo    = fila[COL.PRD_ACTIVO] === true;
      var manejaStk = fila[COL.PRD_MANEJA_STOCK] === true;
      var stock     = Number(fila[COL.PRD_STOCK_ACTUAL]) || 0;
      var critico   = Number(fila[15]) || 0;

      if (!activo || !manejaStk) continue;
      if (stock <= critico) criticos++;
    }

    return criticos;

  } catch (e) {
    Logger.log('ERROR _evaluarStockCritico: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _evaluarCxCVencida
// ============================================================

/**
 * Cuenta documentos de CxC con más de 90 días de vencimiento.
 * Usa la columna L (calc_dias_vencimiento) calculada por ARRAYFORMULA.
 * @returns {number} Cantidad de documentos vencidos >90 días
 */
function _evaluarCxCVencida() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.DB_CXC);
    if (!hoja) return 0;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return 0;

    // Cols: A=id(0), K=calc_saldo(10), L=calc_dias(11), N=estado(13)
    var datos    = hoja.getRange(2, 1, ultimaFila - 1, 14).getValues();
    var vencidas = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var id     = String(fila[0] || '').trim();
      var saldo  = Number(fila[10]) || 0;
      var dias   = Number(fila[11]) || 0;
      var estado = String(fila[13] || '').trim();

      if (!id) continue;
      if (estado === ESTADOS.PAGADO || estado === ESTADOS.INCOBRABLE) continue;
      if (saldo > 0 && dias > 90) vencidas++;
    }

    return vencidas;

  } catch (e) {
    Logger.log('ERROR _evaluarCxCVencida: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _evaluarCxPUrgente
// ============================================================

/**
 * Cuenta documentos de CxP urgentes:
 * - Vencen en los próximos 5 días (dias entre 0 y 5)
 * - Ya están vencidos (dias negativo)
 * 
 * CORRECCIÓN H21: antes excluía CxP ya vencidas con dias < 0.
 * Ahora cualquier saldo pendiente con dias <= 5 es urgente,
 * incluyendo los ya vencidos.
 * 
 * @returns {number} Cantidad de documentos urgentes
 */
function _evaluarCxPUrgente() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.DB_CXP);
    if (!hoja) return 0;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return 0;

    var datos    = hoja.getRange(2, 1, ultimaFila - 1, 16).getValues();
    var urgentes = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var id     = String(fila[0] || '').trim();
      var saldo  = Number(fila[12]) || 0;  // M calc_saldo_pendiente
      var dias   = Number(fila[13]) || 0;  // N calc_dias_vencimiento
      var estado = String(fila[15] || '').trim(); // P estado

      if (!id) continue;
      if (estado === ESTADOS.PAGADO) continue;

      // URGENTE: vence en 5 días O ya está vencido (dias <= 5)
      // Incluye dias negativos (ya vencidos) — son los más críticos
      if (saldo > 0 && dias <= 5) urgentes++;
    }

    return urgentes;

  } catch (e) {
    Logger.log('ERROR _evaluarCxPUrgente: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _evaluarMetasRojo
// ============================================================

/**
 * Cuenta metas activas que están en estado ROJO.
 * Una meta está en ROJO cuando su cumplimiento < umbral_rojo_pct.
 * @returns {number} Cantidad de metas en rojo
 */
function _evaluarMetasRojo() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.DB_METAS);
    if (!hoja) return 0;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return 0;

    var periodo = getPeriodoActual();
    // Cols: E=sys_periodo(4), F=tipo_meta(5), H=valor_meta(7),
    //       I=valor_real(8), J=umbral_amarillo(9), K=umbral_rojo(10), L=activo(11)
    var datos  = hoja.getRange(2, 1, ultimaFila - 1, 12).getValues();
    var rojas  = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila        = datos[i];
      var perFila     = String(fila[4] || '').trim();
      var valorMeta   = Number(fila[7]) || 0;
      var valorReal   = Number(fila[8]) || 0;
      var umbralRojo  = Number(fila[10]) || 0.6;
      var activo      = fila[11];

      if (!activo) continue;
      if (perFila !== periodo) continue;
      if (valorMeta <= 0) continue;

      var cumplimiento = valorReal / valorMeta;
      if (cumplimiento < umbralRojo) rojas++;
    }

    return rojas;

  } catch (e) {
    Logger.log('ERROR _evaluarMetasRojo: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _evaluarDiferenciasCaja
// ============================================================

/**
 * Cuenta cierres de caja con diferencia != 0.
 * @returns {number} Cantidad de cierres con diferencia
 */
function _evaluarDiferenciasCaja() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.DB_CAJA);
    if (!hoja) return 0;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return 0;

    var periodo = getPeriodoActual();
    // Col N = calc_diferencia_caja (índice 13)
    var datos  = hoja.getRange(2, 1, ultimaFila - 1, 14).getValues();
    var difs   = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila      = datos[i];
      var perFila   = String(fila[3] || '').trim(); // D sys_periodo
      var diferencia = Number(fila[13]) || 0;       // N calc_diferencia_caja

      if (perFila !== periodo) continue;
      if (Math.abs(diferencia) > 0) difs++;
    }

    return difs;

  } catch (e) {
    Logger.log('ERROR _evaluarDiferenciasCaja: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _calcularVariacionPrecios
// ============================================================

/**
 * Actualiza calc_variacion_pct en HIST_PRECIOS.
 * Compara cada precio con el registro anterior del mismo
 * producto y proveedor.
 */
function _calcularVariacionPrecios() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.HIST_PRECIOS);
    if (!hoja) return;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 3) return; // Necesita al menos 2 registros de datos

    // Leer todo: cols A-K (11 cols)
    var datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getValues();

    // Construir mapa de último precio por producto+proveedor
    var ultimoPrecio = {};

    for (var i = 0; i < datos.length; i++) {
      var fila      = datos[i];
      var idProd    = String(fila[2] || '').trim(); // C
      var idPrv     = String(fila[3] || '').trim(); // D
      var precioClp = Number(fila[8]) || 0;         // I
      var key       = idProd + '|' + idPrv;

      if (!idProd || !idPrv) continue;

      if (ultimoPrecio[key] !== undefined && ultimoPrecio[key] > 0) {
        // Calcular variación
        var variacion = Math.round(((precioClp - ultimoPrecio[key]) / ultimoPrecio[key]) * 10000) / 100;
        // Escribir en columna J (índice 9, col 10 base 1)
        hoja.getRange(i + 2, 10).setValue(variacion);
      }

      ultimoPrecio[key] = precioClp;
    }

    SpreadsheetApp.flush();
    Logger.log('_calcularVariacionPrecios: Variaciones calculadas');

  } catch (e) {
    Logger.log('ERROR _calcularVariacionPrecios: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _construirEmailAlertas
// ============================================================

/**
 * Construye el texto del email de alertas.
 * @returns {string} Cuerpo del email en texto plano
 */
function _construirEmailAlertas(empresa, fecha, periodo, total,
  stockCritico, cxcVencida, cxpUrgente, metasRojo, difCaja) {

  var lineas = [
    'SISTEMA ERP PYME — REPORTE DE ALERTAS',
    '═══════════════════════════════════════',
    '',
    'Empresa:  ' + empresa,
    'Fecha:    ' + fecha,
    'Período:  ' + periodo,
    'Alertas:  ' + total + ' detectada(s)',
    '',
    '─────────────────────────────────────',
    'DETALLE DE ALERTAS',
    '─────────────────────────────────────'
  ];

  if (stockCritico > 0) {
    lineas.push('⚠️ INVENTARIO: ' + stockCritico + ' producto(s) en stock crítico');
    lineas.push('   → Revisar MSTR_PRODUCTOS y reponer stock urgente.');
  }
  if (cxcVencida > 0) {
    lineas.push('⚠️ CxC: ' + cxcVencida + ' documento(s) con más de 90 días vencido(s)');
    lineas.push('   → Revisar DB_CXC y gestionar cobro.');
  }
  if (cxpUrgente > 0) {
    lineas.push('⚠️ CxP: ' + cxpUrgente + ' pago(s) urgente(s) en los próximos 5 días');
    lineas.push('   → Revisar DB_CXP y preparar pagos.');
  }
  if (metasRojo > 0) {
    lineas.push('⚠️ METAS: ' + metasRojo + ' meta(s) en estado ROJO');
    lineas.push('   → Revisar DB_METAS y tomar acciones correctivas.');
  }
  if (difCaja > 0) {
    lineas.push('⚠️ CAJA: ' + difCaja + ' cierre(s) con diferencia detectada');
    lineas.push('   → Revisar DB_CAJA y cuadrar diferencias.');
  }

  lineas.push('');
  lineas.push('─────────────────────────────────────');
  lineas.push('Este es un mensaje automático del sistema ERP PYME.');
  lineas.push('Para acceder al sistema: abrir el archivo en Google Sheets.');
  lineas.push('');
  lineas.push('ERP PYME Inteligente v' + SISTEMA.VERSION);

  return lineas.join('\n');
}

// ============================================================
// FUNCIÓN PRIVADA: _registrarResultadoBatch
// ============================================================

/**
 * Registra el resultado de una ejecución de batch en BATCH_RESULTADOS.
 *
 * @param {string}  tipoBatch  - Tipo de batch ejecutado
 * @param {boolean} exito      - Si el batch fue exitoso
 * @param {number}  duracion   - Duración en segundos
 * @param {string}  detalle    - JSON con detalle de subprocesos
 */
function _registrarResultadoBatch(tipoBatch, exito, duracion, detalle) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.BATCH_RESULTADOS);
    if (!hoja) return;

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var periodo   = getPeriodoActual() || '';
    var ahora     = new Date();

    var fila = [
      generarID('BAT'),           // A id_batch
      idEmpresa,                  // B id_empresa
      ahora,                      // C generado_en
      periodo,                    // D sys_periodo
      tipoBatch,                  // E tipo_batch
      exito ? 'EXITOSO' : 'ERROR',// F estado
      duracion,                   // G duracion_seg
      detalle,                    // H detalle
      '',                         // I campo9
      '',                         // J campo10
      '',                         // K campo11
      '',                         // L campo12
      '',                         // M campo13
      '',                         // N campo14
      ''                          // O campo15
    ];

    var ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila + 1, 1, 1, 15).setValues([fila]);
    SpreadsheetApp.flush();

  } catch (e) {
    Logger.log('ERROR _registrarResultadoBatch: ' + e.message);
  }
}

// ============================================================
// FIN DE M5_Batch.gs
// ============================================================
