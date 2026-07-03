function limpiarFilasFantasmaSeguro(nombreHoja) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var hoja     = ss.getSheetByName(nombreHoja);
  var filaReal = _ultimaFilaReal(hoja);
  var lastRow  = hoja.getLastRow();

  if (lastRow <= filaReal + 5) {
    Logger.log(nombreHoja + ': sin filas fantasma significativas');
    return;
  }

  // ÚNICO método permitido — preserva validaciones y formato
  var filasEliminar = lastRow - filaReal - 1;
  if (filasEliminar > 0) {
    hoja.deleteRows(filaReal + 2, filasEliminar);
    SpreadsheetApp.flush();
  }

  Logger.log(nombreHoja + ': limpieza completada | nuevo lastRow=' + hoja.getLastRow());
}
function verificarTodasLasCorrecciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== VERIFICACIÓN CORRECCIONES C1-C4 ===');
  Logger.log('');

  var hFin    = ss.getSheetByName('CALC_FINANCIERO');
  var hCom    = ss.getSheetByName('CALC_COMERCIAL');
  var hSupply = ss.getSheetByName('CALC_SUPPLY');

  // C1 — CALC_SUPPLY entradas sin APERTURA
  Logger.log('── C1 CALC_SUPPLY ──');
  var f17   = hSupply.getRange(17,2).getFormula();
  var c1ok  = f17.indexOf('"APERTURA"') !== -1;
  Logger.log((c1ok?'✅':'❌') + ' Excluye APERTURA');
  Logger.log('  entradas_periodo: ' + hSupply.getRange(17,2).getValue());
  Logger.log('  stock_total:      ' + hSupply.getRange(25,2).getValue());
  Logger.log('  valor_inventario: $' + hSupply.getRange(12,2).getValue().toLocaleString('es-CL'));

  // C2 — CALC_COMERCIAL usa YEAR+MONTH
  Logger.log('');
  Logger.log('── C2 CALC_COMERCIAL ──');
  var f7   = hCom.getRange(7,2).getFormula();
  var c2ok = f7.indexOf('YEAR(') !== -1 && f7.indexOf('MONTH(') !== -1;
  Logger.log((c2ok?'✅':'❌') + ' Usa YEAR+MONTH');
  Logger.log('  ventas_neto:      $' + hCom.getRange(7,2).getValue().toLocaleString('es-CL'));
  Logger.log('  documentos_mes:   '  + hCom.getRange(13,2).getValue());
  Logger.log('  ticket_promedio:  $' + hCom.getRange(14,2).getValue().toLocaleString('es-CL'));

  // C3 — CALC_FINANCIERO valores correctos
  Logger.log('');
  Logger.log('── C3 CALC_FINANCIERO ──');
  var ingNeto = hFin.getRange(8,2).getValue();
  var flujo   = hFin.getRange(12,2).getValue();
  var cxcTot  = hFin.getRange(17,2).getValue();
  var cxpTot  = hFin.getRange(30,2).getValue();
  var liq     = hFin.getRange(45,2).getValue();
  Logger.log('  ingresos_neto:    $' + ingNeto.toLocaleString('es-CL'));
  Logger.log('  flujo_neto:       $' + flujo.toLocaleString('es-CL'));
  Logger.log('  cxc_total:        $' + cxcTot.toLocaleString('es-CL'));
  Logger.log('  cxp_total:        $' + cxpTot.toLocaleString('es-CL'));
  Logger.log('  liquidez_proy:    $' + liq.toLocaleString('es-CL'));
  Logger.log(ingNeto === 1395000 ? '✅ ingresos OK' : '❌ ingresos: ' + ingNeto);
  Logger.log(cxcTot  === 1490000 ? '✅ cxc OK'     : '❌ cxc: '     + cxcTot);
  Logger.log(cxpTot  === 595000  ? '✅ cxp OK'     : '❌ cxp: '     + cxpTot);

  // C4 — limpiarFilasFantasmaSeguro existe
  Logger.log('');
  Logger.log('── C4 LIMPIEZA SEGURA ──');
  var c4ok = typeof limpiarFilasFantasmaSeguro === 'function';
  Logger.log((c4ok?'✅':'❌') + ' limpiarFilasFantasmaSeguro existe');

  var todoOk = c1ok && c2ok &&
               ingNeto === 1395000 &&
               hCom.getRange(7,2).getValue() === 46050 && c4ok;

  Logger.log('');
  Logger.log(todoOk ?
    '✅ TODAS LAS CORRECCIONES VERIFICADAS' :
    '❌ REVISAR PUNTOS MARCADOS');
}
function auditarFuentesSupply() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES CALC_SUPPLY ===');

  // MSTR_PRODUCTOS
  var hPrd  = ss.getSheetByName('MSTR_PRODUCTOS');
  var fPrd  = _ultimaFilaReal(hPrd);
  Logger.log('');
  Logger.log('── MSTR_PRODUCTOS ──');
  Logger.log('Registros reales: ' + (fPrd-1));
  var hdrs = hPrd.getRange(1,1,1,hPrd.getLastColumn()).getValues()[0];
  hdrs.forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });
  if(fPrd>=2){
    var datos = hPrd.getRange(2,1,fPrd-1,hPrd.getLastColumn()).getValues();
    datos.forEach(function(f,i){
      if(f[0]==='') return;
      Logger.log('  Fila '+(i+2)+': id=['+f[0]+'] nombre=['+f[3]+
                 '] stock=['+f[11]+'] min=['+f[13]+'] critico=['+f[15]+
                 '] costo=['+f[7]+'] activo=['+f[27]+'] manejaStock=['+f[19]+']');
    });
  }

  // DB_MOVIMIENTOS
  var hMov = ss.getSheetByName('DB_MOVIMIENTOS');
  var fMov = _ultimaFilaReal(hMov);
  Logger.log('');
  Logger.log('── DB_MOVIMIENTOS ──');
  Logger.log('Registros reales: ' + (fMov-1));
  if(fMov>=2){
    var movs = hMov.getRange(2,1,fMov-1,16).getValues();
    movs.forEach(function(f,i){
      if(f[0]==='') return;
      Logger.log('  Fila '+(i+2)+': tipo=['+f[7]+'] cant=['+f[8]+
                 '] impacto=['+f[9]+'] producto=['+f[4]+'] fecha=['+f[2]+']');
    });
  }

  // HIST_PRECIOS
  var hHP  = ss.getSheetByName('HIST_PRECIOS');
  var fHP  = _ultimaFilaReal(hHP);
  Logger.log('');
  Logger.log('── HIST_PRECIOS ──');
  Logger.log('Registros reales: ' + (fHP-1));
  if(fHP>=2){
    var hpDatos = hHP.getRange(2,1,fHP-1,11).getValues();
    hpDatos.forEach(function(f,i){
      if(f[0]==='') return;
      Logger.log('  Fila '+(i+2)+': producto=['+f[2]+'] precio_clp=['+f[8]+
                 '] variacion=['+f[9]+'] periodo=['+f[5]+']');
    });
  }
}
function auditarCALC_ALERTAS() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_ALERTAS');
  Logger.log('=== AUDITORÍA CALC_ALERTAS ===');
  Logger.log('lastRow: ' + hoja.getLastRow());
  Logger.log('lastCol: ' + hoja.getLastColumn());
  Logger.log('');
  var datos = hoja.getRange(1,1,hoja.getLastRow(),
              hoja.getLastColumn()).getValues();
  datos.forEach(function(f,i){
    if(f[0]!==''||f[1]!==''){
      Logger.log('Fila '+(i+1)+': ['+f[0]+'] = ['+f[1]+']');
    }
  });
}
function diagnosticarHojasProyecto() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var todasLasHojas = ss.getSheets();
  
  Logger.log('=== DIAGNÓSTICO HOJAS DEL PROYECTO ===');
  Logger.log('Archivo: ' + ss.getName());
  Logger.log('ID:      ' + ss.getId());
  Logger.log('Total hojas: ' + todasLasHojas.length);
  Logger.log('');
  Logger.log('── LISTADO COMPLETO ──');
  
  var nombres = [];
  todasLasHojas.forEach(function(hoja, i) {
    var nombre  = hoja.getName();
    var oculta  = hoja.isSheetHidden() ? ' [OCULTA]' : '';
    nombres.push(nombre);
    Logger.log((i+1) + '. ' + nombre + oculta);
  });

  Logger.log('');
  Logger.log('── VERIFICACIÓN HOJAS CALC_ ──');
  var hojasClave = [
    'CALC_FINANCIERO',
    'CALC_COMERCIAL', 
    'CALC_SUPPLY',
    'CALC_GERENCIAL',
    'CALC_ALERTAS',
    'DASH_EJECUTIVO'
  ];

  hojasClave.forEach(function(nombre) {
    var existe = nombres.indexOf(nombre) !== -1;
    var hoja   = ss.getSheetByName(nombre);
    var filas  = hoja ? hoja.getLastRow() : 0;
    Logger.log((existe ? '✅' : '❌') + ' ' + nombre + 
               (existe ? ' — filas: ' + filas : ' — NO EXISTE'));
  });

  Logger.log('');
  Logger.log('── ANÁLISIS CAUSA RAÍZ ──');
  Logger.log('crearCALC_GERENCIAL busca: ss.getSheetByName("CALC_GERENCIAL")');
  Logger.log('Si retorna null → muestra alert y termina sin crear nada');
  Logger.log('Solución: agregar insertSheet() si la hoja no existe');
  Logger.log('');
  Logger.log('Spreadsheet activo en script: ' + ss.getName());
  Logger.log('¿Es el archivo DEV?: ' + (ss.getName().indexOf('DEV') !== -1 ? 'SÍ' : 'VERIFICAR'));
}
function auditarSemaforosCALC_GERENCIAL() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_GERENCIAL');
  
  Logger.log('=== AUDITORÍA SEMÁFOROS CALC_GERENCIAL ===');
  
  var celdas = [
    {fila: 15, label: 'semaforo_liquidez'},
    {fila: 23, label: 'semaforo_resultado'},
    {fila: 32, label: 'semaforo_ventas'},
    {fila: 42, label: 'semaforo_cobranza'},
    {fila: 50, label: 'semaforo_inventario'},
    {fila: 60, label: 'score_liquidez'},
    {fila: 61, label: 'score_resultado'},
    {fila: 62, label: 'score_ventas'},
    {fila: 63, label: 'score_cobranza'},
    {fila: 64, label: 'score_total'},
    {fila: 65, label: 'estado_general'}
  ];

  celdas.forEach(function(c) {
    var formula  = hoja.getRange(c.fila, 2).getFormula();
    var valor    = hoja.getRange(c.fila, 2).getValue();
    var display  = hoja.getRange(c.fila, 2).getDisplayValue();
    Logger.log('');
    Logger.log('Fila ' + c.fila + ' — ' + c.label);
    Logger.log('  Fórmula: [' + formula + ']');
    Logger.log('  getValue: [' + valor + ']');
    Logger.log('  display:  [' + display + ']');
    Logger.log('  isNaN:    ' + isNaN(Number(valor)));
  });

  // También verificar los valores fuente
  Logger.log('');
  Logger.log('── VALORES FUENTE ──');
  var hFin = ss.getSheetByName('CALC_FINANCIERO');
  var hCom = ss.getSheetByName('CALC_COMERCIAL');
  var hSup = ss.getSheetByName('CALC_SUPPLY');

  Logger.log('CALC_FINANCIERO!B46 (ratio_liquidez):  [' + hFin.getRange(46,2).getValue() + '] tipo: ' + typeof hFin.getRange(46,2).getValue());
  Logger.log('CALC_FINANCIERO!B14 (margen_pct):      [' + hFin.getRange(14,2).getValue() + '] tipo: ' + typeof hFin.getRange(14,2).getValue());
  Logger.log('CALC_FINANCIERO!B18 (cxc_vencida):     [' + hFin.getRange(18,2).getValue() + '] tipo: ' + typeof hFin.getRange(18,2).getValue());
  Logger.log('CALC_FINANCIERO!B26 (cxc_riesgo_pct):  [' + hFin.getRange(26,2).getValue() + '] tipo: ' + typeof hFin.getRange(26,2).getValue());
  Logger.log('CALC_COMERCIAL!B7  (ventas_neto):      [' + hCom.getRange(7,2).getValue()  + '] tipo: ' + typeof hCom.getRange(7,2).getValue());
  Logger.log('CALC_COMERCIAL!B11 (margen_bruto_pct): [' + hCom.getRange(11,2).getValue() + '] tipo: ' + typeof hCom.getRange(11,2).getValue());
  Logger.log('CALC_SUPPLY!B10   (prod_criticos):     [' + hSup.getRange(10,2).getValue() + '] tipo: ' + typeof hSup.getRange(10,2).getValue());
}
function testLineaError() {
  // Test de las 4 fórmulas corregidas individualmente
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_GERENCIAL');
  
  Logger.log('=== TEST FÓRMULAS CORREGIDAS ===');
  
  try {
    hoja.getRange(15,2).setFormula('=SUMPRODUCT((CALC_FINANCIERO!B46>=1,5)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B46>=1)*(CALC_FINANCIERO!B46<1,5)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B46<1)*1)*1');
    Logger.log('✅ Fila 15 OK: ' + hoja.getRange(15,2).getValue());
  } catch(e) { Logger.log('❌ Fila 15: ' + e.message); }

  try {
    hoja.getRange(23,2).setFormula('=SUMPRODUCT((CALC_FINANCIERO!B14>=0,2)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B14>=0)*(CALC_FINANCIERO!B14<0,2)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B14<0)*1)*1');
    Logger.log('✅ Fila 23 OK: ' + hoja.getRange(23,2).getValue());
  } catch(e) { Logger.log('❌ Fila 23: ' + e.message); }

  try {
    hoja.getRange(32,2).setFormula('=SUMPRODUCT((CALC_COMERCIAL!B7>0)*1)*2+SUMPRODUCT((CALC_COMERCIAL!B11>=0,2)*1)');
    Logger.log('✅ Fila 32 OK: ' + hoja.getRange(32,2).getValue());
  } catch(e) { Logger.log('❌ Fila 32: ' + e.message); }

  try {
    hoja.getRange(42,2).setFormula('=SUMPRODUCT((CALC_FINANCIERO!B18=0)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B18>0)*(CALC_FINANCIERO!B26<0,1)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B26>=0,1)*1)*1');
    Logger.log('✅ Fila 42 OK: ' + hoja.getRange(42,2).getValue());
  } catch(e) { Logger.log('❌ Fila 42: ' + e.message); }

  SpreadsheetApp.flush();
  Utilities.sleep(500);

  // Verificar scores que dependen de semáforos
  Logger.log('');
  Logger.log('── SCORES ──');
  Logger.log('B15 semaforo_liquidez:  ' + hoja.getRange(15,2).getValue());
  Logger.log('B23 semaforo_resultado: ' + hoja.getRange(23,2).getValue());
  Logger.log('B32 semaforo_ventas:    ' + hoja.getRange(32,2).getValue());
  Logger.log('B42 semaforo_cobranza:  ' + hoja.getRange(42,2).getValue());
  Logger.log('B60 score_liquidez:     ' + hoja.getRange(60,2).getValue());
  Logger.log('B64 score_total:        ' + hoja.getRange(64,2).getValue());
}
function verificacionFinalMotores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== VERIFICACIÓN FINAL TODOS LOS MOTORES ===');
  Logger.log('');

  var modulos = [
    { nombre: 'CALC_FINANCIERO', kpis: [
      {fila:8,  label:'ingresos_neto',     esperado:1395000},
      {fila:12, label:'flujo_neto',        esperado:1047400},
      {fila:17, label:'cxc_total',         esperado:1490000},
      {fila:30, label:'cxp_total',         esperado:595000},
      {fila:45, label:'liquidez_proy',     esperado:1942400}
    ]},
    { nombre: 'CALC_COMERCIAL', kpis: [
      {fila:7,  label:'ventas_neto',       esperado:46050},
      {fila:13, label:'documentos_mes',    esperado:5},
      {fila:14, label:'ticket_promedio',   esperado:9210},
      {fila:25, label:'clientes_activos',  esperado:1}
    ]},
    { nombre: 'CALC_SUPPLY', kpis: [
      {fila:7,  label:'productos_activos', esperado:2},
      {fila:12, label:'valor_inventario',  esperado:45000},
      {fila:25, label:'stock_total',       esperado:45}
    ]},
    { nombre: 'CALC_GERENCIAL', kpis: [
      {fila:11, label:'liquidez_proy',     esperado:1942400},
      {fila:26, label:'ventas_mes',        esperado:46050},
      {fila:64, label:'score_total',       esperado:100},
      {fila:65, label:'estado_general',    esperado:3}
    ]}
  ];

  var totalOk    = 0;
  var totalFallo = 0;

  modulos.forEach(function(mod) {
    var hoja = ss.getSheetByName(mod.nombre);
    Logger.log('── ' + mod.nombre + ' ──');

    mod.kpis.forEach(function(kpi) {
      var valor = Number(hoja.getRange(kpi.fila, 2).getValue());
      var ok    = Math.abs(valor - kpi.esperado) < 1;
      Logger.log((ok?'✅':'❌') + ' ' + kpi.label + ': ' + valor +
                 (ok ? '' : ' (esperado: ' + kpi.esperado + ')'));
      if (ok) totalOk++; else totalFallo++;
    });
    Logger.log('');
  });

  // Verificar funciones críticas del sistema
  Logger.log('── FUNCIONES CRÍTICAS ──');
  var funciones = [
    'crearCALC_FINANCIERO','crearCALC_COMERCIAL',
    'crearCALC_SUPPLY','crearCALC_GERENCIAL',
    'validarCALC_FINANCIERO','validarCALC_SUPPLY',
    'limpiarFilasFantasmaSeguro','getPeriodoActual'
  ];
  funciones.forEach(function(f) {
    var existe = typeof eval(f) === 'function';
    Logger.log((existe?'✅':'❌') + ' ' + f);
    if (existe) totalOk++; else totalFallo++;
  });

  Logger.log('');
  Logger.log('=== RESULTADO FINAL ===');
  Logger.log('✅ OK:     ' + totalOk);
  Logger.log('❌ Fallos: ' + totalFallo);
  Logger.log('');
  Logger.log(totalFallo === 0 ?
    '✅ SISTEMA LISTO PARA DASH_EJECUTIVO' :
    '❌ CORREGIR ANTES DE AVANZAR');
}
function testNuevoCampos() {
  var d = getDashboardData();
  Logger.log('=== TEST CAMPOS NUEVOS ===');
  if (!d) { Logger.log('ERROR: retornó null'); return; }
  Logger.log('semExecFinanzas:    ' + d.semExecFinanzas);
  Logger.log('semExecComercial:   ' + d.semExecComercial);
  Logger.log('semExecOperaciones: ' + d.semExecOperaciones);
  Logger.log('diasCaja:           ' + d.diasCaja);
  Logger.log('metaVentas:         ' + d.metaVentas);
  Logger.log('skusQuiebre:        ' + JSON.stringify(d.skusQuiebre));
}
function auditarFuentesComercial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO COMERCIAL ===');

  // CALC_COMERCIAL
  var cc = ss.getSheetByName('CALC_COMERCIAL');
  Logger.log('── CALC_COMERCIAL ──');
  [[7,'ventas_neto'],[8,'ventas_bruto'],[10,'margen_bruto'],
   [11,'margen_bruto_pct'],[13,'documentos_mes'],[14,'ticket_promedio'],
   [19,'ventas_dia_neto'],[21,'docs_dia'],[25,'clientes_activos'],
   [27,'productos_activos']
  ].forEach(function(k) {
    Logger.log('  B'+k[0]+' '+k[1]+': '+cc.getRange(k[0],2).getValue());
  });

  // DB_VENTAS — columnas clave
  var hV  = ss.getSheetByName('DB_VENTAS');
  var filV = _ultimaFilaReal(hV);
  Logger.log('── DB_VENTAS ('+( filV-1)+' registros) ──');
  if (filV >= 2) {
    hV.getRange(2,1,filV-1,25).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+
        ': tipo=['+f[3]+'] num=['+f[4]+
        '] fecha=['+f[10]+'] cliente=['+f[12]+
        '] producto=['+f[13]+'] cant=['+f[16]+
        '] neto=['+f[19]+'] margen=['+f[23]+
        '] margenPct=['+f[24]+'] estado=['+f[27]+']');
    });
  }

  // MSTR_CLIENTES — columnas nombre
  var hC  = ss.getSheetByName('MSTR_CLIENTES');
  var filC = _ultimaFilaReal(hC);
  Logger.log('── MSTR_CLIENTES ('+( filC-1)+' registros) ──');
  if (filC >= 2) {
    hC.getRange(2,1,filC-1,10).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+': id=['+f[0]+'] nombre=['+f[3]+'] rut=['+f[2]+']');
    });
  }

  // MSTR_PRODUCTOS — columnas nombre
  var hP  = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  Logger.log('── MSTR_PRODUCTOS ('+( filP-1)+' registros) ──');
  if (filP >= 2) {
    hP.getRange(2,1,filP-1,10).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+': id=['+f[0]+'] sku=['+f[2]+'] nombre=['+f[3]+']');
    });
  }
}
function auditarColumnasTransacciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA COLUMNAS TRANSACCIONES ===');

  var hojas = ['DB_VENTAS','DB_COMPRAS','DB_INGRESOS','DB_EGRESOS','DB_CXC','DB_CXP'];

  hojas.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) { Logger.log(nombre + ': NO EXISTE'); return; }
    var hdrs = hoja.getRange(1,1,1,hoja.getLastColumn()).getValues()[0];
    Logger.log('');
    Logger.log('── ' + nombre + ' ──');
    hdrs.forEach(function(h,i) {
      if (h!=='') Logger.log('  col '+(i+1)+': '+h);
    });
  });
}
function auditarFuentesInventario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO INVENTARIO ===');

  var cs = ss.getSheetByName('CALC_SUPPLY');
  Logger.log('── CALC_SUPPLY ──');
  [[7,'productos_activos'],[8,'con_stock'],[9,'sin_stock'],
   [10,'criticos'],[12,'valor_costo'],[13,'valor_venta'],
   [17,'entradas_periodo'],[18,'salidas_periodo'],[20,'unid_vendidas'],
   [21,'unid_compradas'],[22,'ratio_rotacion'],[25,'stock_total']
  ].forEach(function(k) {
    Logger.log('  B'+k[0]+' '+k[1]+': '+cs.getRange(k[0],2).getValue());
  });

  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var fP = _ultimaFilaReal(hP);
  Logger.log('── MSTR_PRODUCTOS ('+( fP-1)+' registros) ──');
  if (fP >= 2) {
    hP.getRange(2,1,fP-1,28).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+
        ': sku=['+f[2]+'] nombre=['+f[3]+
        '] stock=['+f[11]+'] min=['+f[13]+
        '] critico=['+f[15]+'] costo=['+f[7]+
        '] precio=['+f[8]+'] activo=['+f[27]+']');
    });
  }
}
function diagnosticarDiasSinVenta() {
  var d = getInventarioData();
  Logger.log('=== DIAGNÓSTICO DÍAS SIN VENTA ===');
  d.productos.forEach(function(p) {
    Logger.log(p.nombre + ': diasSinVenta=[' + p.diasSinVenta + '] tipo=' + typeof p.diasSinVenta);
  });
}
function auditarFuentesCobranza() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO COBRANZA ===');

  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('── CALC_FINANCIERO ──');
  [[17,'cxc_total'],[18,'cxc_vencida'],[19,'cxc_030'],
   [20,'cxc_3160'],[21,'cxc_mas60'],[22,'docs_pendientes'],
   [23,'docs_vencidos']
  ].forEach(function(k) {
    Logger.log('  B'+k[0]+' '+k[1]+': '+cf.getRange(k[0],2).getValue());
  });

  var hCxC = ss.getSheetByName('DB_CXC');
  var filCxC = _ultimaFilaReal(hCxC);
  Logger.log('── DB_CXC ('+( filCxC-1)+' registros) ──');
  if (filCxC >= 2) {
    hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+
        ': cliente=['+f[2]+'] doc=['+f[4]+
        '] emision=['+f[6]+'] vence=['+f[7]+
        '] saldo=['+f[10]+'] dias=['+f[11]+
        '] tramo=['+f[12]+'] estado=['+f[13]+']');
    });
  }

  var hC = ss.getSheetByName('MSTR_CLIENTES');
  var filC = _ultimaFilaReal(hC);
  Logger.log('── MSTR_CLIENTES condicion_pago ──');
  if (filC >= 2) {
    hC.getRange(2,1,filC-1,15).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+': id=['+f[0]+'] condicion=['+f[14]+']');
    });
  }
}
function testCobranzaData() {
  var d = getCobranzaData();
  Logger.log('resultado: ' + (d === null ? 'NULL' : 'OK'));
  if (d) Logger.log(JSON.stringify(d).substring(0,300));
}
function auditarFuentesProveedores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO PROVEEDORES ===');

  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('── CALC_FINANCIERO ──');
  [[30,'cxp_total'],[31,'cxp_vencida'],[32,'cxp_urgente_0_5'],
   [33,'cxp_6_30'],[34,'cxp_mas30'],[35,'docs_pendientes'],
   [36,'docs_vencidos'],[42,'caja_periodo']
  ].forEach(function(k) {
    Logger.log('  B'+k[0]+' '+k[1]+': '+cf.getRange(k[0],2).getValue());
  });

  var hCxP = ss.getSheetByName('DB_CXP');
  var filCxP = _ultimaFilaReal(hCxP);
  Logger.log('── DB_CXP ('+( filCxP-1)+' registros) ──');
  if (filCxP >= 2) {
    hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+
        ': proveedor=['+f[2]+'] doc=['+f[4]+
        '] emision=['+f[6]+'] vence=['+f[7]+
        '] saldo=['+f[12]+'] dias=['+f[13]+
        '] tramo=['+f[14]+'] estado=['+f[15]+']');
    });
  }

  var hP = ss.getSheetByName('MSTR_PROVEEDORES');
  var filP = _ultimaFilaReal(hP);
  Logger.log('── MSTR_PROVEEDORES ('+( filP-1)+' registros) ──');
  if (filP >= 2) {
    hP.getRange(2,1,filP-1,15).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  F'+(i+2)+': id=['+f[0]+'] nombre=['+f[3]+'] rut=['+f[2]+'] condicion=['+f[10]+']');
    });
  }
}
function auditarFuentesAlertas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO ALERTAS ===');

  var ca = ss.getSheetByName('CALC_ALERTAS');
  Logger.log('── CALC_ALERTAS ──');
  var datos = ca.getRange(1,1,ca.getLastRow(),ca.getLastColumn()).getValues();
  datos.forEach(function(f,i) {
    if (f[0]!=='' || f[1]!=='') Logger.log('  Fila '+(i+1)+': ['+f[0]+'] = ['+f[1]+']');
  });

  var hBatch = ss.getSheetByName('BATCH_RESULTADOS');
  if (hBatch) {
    var filB = _ultimaFilaReal(hBatch);
    Logger.log('── BATCH_RESULTADOS ('+( filB-1)+' registros) ──');
    if (filB >= 2) {
      hBatch.getRange(Math.max(2,filB-3),1,Math.min(4,filB-1),10).getValues().forEach(function(f,i) {
        if (f[0]==='') return;
        Logger.log('  F: '+JSON.stringify(f).substring(0,200));
      });
    }
  } else {
    Logger.log('BATCH_RESULTADOS: no existe');
  }
}
function auditarFuentesAlertasV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA ALERTAS V2 ===');

  // Verificar columnas DB_VENTAS para margen negativo
  var hV = ss.getSheetByName('DB_VENTAS');
  var filV = _ultimaFilaReal(hV);
  Logger.log('DB_VENTAS: '+(filV-1)+' registros');
  if (filV >= 2) {
    hV.getRange(2,1,filV-1,28).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  margen=['+f[23]+'] margenPct=['+f[24]+'] doc=['+f[4]+'] producto=['+f[13]+']');
    });
  }

  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('cxc_mas60: '+cf.getRange('B21').getValue());
  Logger.log('cxp_total: '+cf.getRange('B30').getValue());
}
function auditarPlanCuentas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA MSTR_PLAN_CUENTAS ===');

  var hPC = ss.getSheetByName('MSTR_PLAN_CUENTAS');
  if (!hPC) { Logger.log('NO EXISTE MSTR_PLAN_CUENTAS'); return; }

  var hdrs = hPC.getRange(1,1,1,hPC.getLastColumn()).getValues()[0];
  Logger.log('── Columnas ──');
  hdrs.forEach(function(h,i) { if(h!=='') Logger.log('  col '+(i+1)+': '+h); });

  var filPC = hPC.getLastRow();
  Logger.log('');
  Logger.log('── Cuentas existentes ('+(filPC-1)+') ──');
  if (filPC >= 2) {
    hPC.getRange(2,1,filPC-1,8).getValues().forEach(function(f,i) {
      if (f[0]==='') return;
      Logger.log('  '+f[0]+' | '+f[1]+' | tipo=['+f[2]+'] | nivel=['+f[3]+']');
    });
  }
}
function agregarParametroRetencionBH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('CONFIG_SISTEMA');

  Logger.log('=== AGREGANDO PARÁMETRO RETENCIÓN BH ===');

  // Buscar primera fila vacía en columnas de parámetros (asumiendo estructura clave-valor adicional)
  // Verificamos si ya existe
  var datos = cfg.getDataRange().getValues();
  var yaExiste = false;
  datos.forEach(function(fila) {
    fila.forEach(function(celda) {
      if (String(celda).indexOf('retencion_bh') !== -1) yaExiste = true;
    });
  });

  if (yaExiste) {
    Logger.log('El parámetro ya existe, no se duplica');
    return;
  }

  // Agregamos en una zona segura: columna V y W (ajustar si están ocupadas)
  cfg.getRange('V1').setValue('retencion_bh_pct');
  cfg.getRange('V2').setValue(0.1525);
  cfg.getRange('W1').setValue('retencion_bh_actualizado');
  cfg.getRange('W2').setValue(new Date());

  SpreadsheetApp.flush();
  Logger.log('✅ Parámetro creado: retencion_bh_pct = 15.25% en CONFIG_SISTEMA!V2');
  Logger.log('Verificar visualmente que V1/W1 no chocan con otra columna existente');
}
function agregarParametroRetencionBH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('CONFIG_SISTEMA');

  Logger.log('=== AGREGANDO PARÁMETRO RETENCIÓN BH ===');

  var datos = cfg.getDataRange().getValues();
  var yaExiste = false;
  datos.forEach(function(fila) {
    fila.forEach(function(celda) {
      if (String(celda).indexOf('retencion_bh') !== -1) yaExiste = true;
    });
  });

  if (yaExiste) {
    Logger.log('El parámetro ya existe, no se duplica');
    return;
  }

  cfg.getRange('Y1').setValue('retencion_bh_pct');
  cfg.getRange('Y2').setValue(0.1525);
  cfg.getRange('Z1').setValue('retencion_bh_actualizado');
  cfg.getRange('Z2').setValue(new Date());

  SpreadsheetApp.flush();
  Logger.log('✅ Parámetro creado: retencion_bh_pct = 15.25% en CONFIG_SISTEMA!Z2');
  Logger.log('Y1=retencion_bh_pct | Y2=0.1525 | Z1=retencion_bh_actualizado | Z2=fecha');
}
function auditarFichaClienteProveedor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA FICHA CLIENTE/PROVEEDOR ===');

  var hC = ss.getSheetByName('MSTR_CLIENTES');
  Logger.log('── MSTR_CLIENTES columnas ──');
  hC.getRange(1,1,1,hC.getLastColumn()).getValues()[0].forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });

  var hP = ss.getSheetByName('MSTR_PROVEEDORES');
  Logger.log('── MSTR_PROVEEDORES columnas ──');
  hP.getRange(1,1,1,hP.getLastColumn()).getValues()[0].forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });

  Logger.log('── DB_CXC columnas ──');
  var hCxC = ss.getSheetByName('DB_CXC');
  hCxC.getRange(1,1,1,hCxC.getLastColumn()).getValues()[0].forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });

  Logger.log('── DB_CXP columnas ──');
  var hCxP = ss.getSheetByName('DB_CXP');
  hCxP.getRange(1,1,1,hCxP.getLastColumn()).getValues()[0].forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });

  // Verificar si DB_CXC/CXP tienen columna para relacionar NC con su factura origen
  Logger.log('');
  Logger.log('¿Existe columna id_referencia en DB_CXC? Buscar manualmente arriba');
}
function testFichas() {
  var c = getFichaCliente('CLI-20260615-223109-222');
  Logger.log('Cliente: ' + JSON.stringify(c).substring(0,300));
  
  var p = getFichaProveedor('PRV-20260617-084029-660');
  Logger.log('Proveedor: ' + JSON.stringify(p).substring(0,300));
}
function verificarIdEnTopProveedores() {
  var d = getProveedoresData();
  Logger.log(JSON.stringify(d.topProveedores));
}
function verificarIdEnTopDeudores() {
  var d = getComercialData();
  Logger.log(JSON.stringify(d.topClientes));
}
function testFuncionFicha() {
  Logger.log(typeof pintarFichaCliente);
}
function auditarDatosParaVenta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== AUDITORÍA DATOS PARA NUEVA VENTA ===');

  var hC = ss.getSheetByName('MSTR_CLIENTES');
  var filC = _ultimaFilaReal(hC);
  Logger.log('MSTR_CLIENTES: '+(filC-1)+' registros, '+hC.getLastColumn()+' columnas');

  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  Logger.log('MSTR_PRODUCTOS: '+(filP-1)+' registros, '+hP.getLastColumn()+' columnas');

  // Confirmar índices clave de producto
  if (filP >= 2) {
    var muestra = hP.getRange(2,1,1,28).getValues()[0];
    Logger.log('Producto ejemplo — sku=['+muestra[2]+'] nombre=['+muestra[3]+'] unidad=['+muestra[6]+'] precio=['+muestra[8]+'] stock=['+muestra[11]+'] activo=['+muestra[27]+']');
  }

  var cfg = ss.getSheetByName('CONFIG_SISTEMA');
  Logger.log('tasa_iva: '+cfg.getRange('L2').getValue());

  // Verificar campos de CONFIG_PERIODOS abierto
  var hPer = ss.getSheetByName('CONFIG_PERIODOS');
  if (hPer) {
    var filPer = hPer.getLastRow();
    Logger.log('── CONFIG_PERIODOS ──');
    hPer.getRange(2,1,filPer-1,5).getValues().forEach(function(f){
      if (f[0]!=='') Logger.log('  '+f[0]+' = '+f[1]);
    });
  }
}
function testBusquedas() {
  Logger.log('Clientes "negocio": ' + JSON.stringify(buscarClientes('negocio')));
  Logger.log('Clientes "60803": ' + JSON.stringify(buscarClientes('60803')));
  Logger.log('Productos "test": ' + JSON.stringify(buscarProductos('test')));
  Logger.log('Productos "TEST-M3": ' + JSON.stringify(buscarProductos('TEST-M3')));
}
function crearClienteGenerico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('MSTR_CLIENTES');

  Logger.log('=== CREANDO CLIENTE GENÉRICO ===');

  var ctx = getDatosContexto();
  var empresa = ctx.idEmpresa;

  // Verificar si ya existe
  var filC = _ultimaFilaReal(hoja);
  var yaExiste = false;
  if (filC >= 2) {
    hoja.getRange(2,1,filC-1,1).getValues().forEach(function(f) {
      if (String(f[0]) === 'CLI-GENERICO') yaExiste = true;
    });
  }

  if (yaExiste) {
    Logger.log('El Cliente Genérico ya existe, no se duplica');
    return;
  }

  var usuario = Session.getActiveUser().getEmail() || 'sistema';
  var fecha = new Date();

  // Estructura: id, id_empresa, tipo_cliente, rut_normalizado, rut_cliente,
  // razon_social, nombre_fantasia, giro, email, telefono, direccion, comuna,
  // region, id_segmento, condicion_pago, limite_credito, id_vendedor, activo,
  // fecha_inactivacion, creado_en, creado_por, modificado_en, modificado_por
  var fila = [
    'CLI-GENERICO',           // id_cliente
    empresa,                  // id_empresa
    'GENERICO',                // tipo_cliente
    '',                        // rut_normalizado
    '',                        // rut_cliente
    'Cliente Genérico — Venta Mostrador', // razon_social
    'Venta Mostrador',         // nombre_fantasia
    '',                        // giro
    '',                        // email
    '',                        // telefono
    '',                        // direccion
    '',                        // comuna
    '',                        // region
    '',                        // id_segmento
    'CONTADO',                 // condicion_pago — forzado, nunca crédito
    0,                         // limite_credito
    '',                        // id_vendedor
    true,                      // activo
    '',                        // fecha_inactivacion
    fecha,                     // creado_en
    usuario,                   // creado_por
    '',                        // modificado_en
    ''                         // modificado_por
  ];

  var primeraFilaVacia = _ultimaFilaReal(hoja) + 1;
  hoja.getRange(primeraFilaVacia, 1, 1, 23).setValues([fila]);

  SpreadsheetApp.flush();
  Logger.log('✅ Cliente Genérico creado en fila ' + primeraFilaVacia);
  Logger.log('ID: CLI-GENERICO | Condición: CONTADO (forzada)');
}
function testTodoOK2() {
  Logger.log('getDatosContexto: ' + (getDatosContexto() === null ? 'NULL ❌' : 'OK ✅'));
  Logger.log('getDashboardData: ' + (getDashboardData() === null ? 'NULL ❌' : 'OK ✅'));
  Logger.log('getFinanzasData: ' + (getFinanzasData() === null ? 'NULL ❌' : 'OK ✅'));
}
function testFinanzasDirecto() {
  var r = getFinanzasData();
  Logger.log('getFinanzasData resultado: ' + (r === null ? 'NULL' : 'OK'));
}
function verificarTodasLasFunciones() {
  var funciones = ['getDashboardData','getFinanzasData','getComercialData','getInventarioData','getCobranzaData','getProveedoresData','getAlertasData'];
  funciones.forEach(function(nombre) {
    try {
      var fn = this[nombre];
      Logger.log(nombre + ': ' + (typeof fn === 'function' ? 'EXISTE' : 'NO EXISTE'));
    } catch(e) {
      Logger.log(nombre + ': ERROR AL ACCEDER - ' + e.message);
    }
  });

  // Detectar si hay 2 archivos con la misma función (conflicto de nombres)
  Logger.log('');
  Logger.log('=== PROBANDO EJECUCIÓN REAL UNA POR UNA ===');
  funciones.forEach(function(nombre) {
    try {
      var resultado = eval(nombre + '()');
      Logger.log(nombre + '() => ' + (resultado === null ? 'NULL ❌' : 'OK ✅'));
    } catch(e) {
      Logger.log(nombre + '() => EXCEPCIÓN: ' + e.message);
    }
  });
}
function doGet_test() {
  var resultado = doGet(null);
  Logger.log('doGet ejecutó sin morir: OK');
}
function verificarCambioOpenById() {
  var codigoFinanzas = getFinanzasData.toString();
  var tieneOpenById = codigoFinanzas.indexOf('openById') !== -1;
  var tieneGetActive = codigoFinanzas.indexOf('getActiveSpreadsheet') !== -1;
  Logger.log('getFinanzasData usa openById: ' + tieneOpenById);
  Logger.log('getFinanzasData usa getActiveSpreadsheet: ' + tieneGetActive);
  Logger.log('');
  Logger.log('Primeras 150 chars de la función:');
  Logger.log(codigoFinanzas.substring(0,150));
}
function verificarCatchActivo() {
  var codigo = getFinanzasData.toString();
  Logger.log('Tiene errorReal: ' + (codigo.indexOf('errorReal') !== -1));
}
function buscarDuplicados() {
  var files = DriveApp.getFileById(ScriptApp.getScriptId());
}
function simularLlamadaWeb() {
  try {
    var r = getFinanzasData();
    Logger.log('Resultado tipo: ' + typeof r);
    Logger.log('Resultado: ' + JSON.stringify(r).substring(0,300));
  } catch(e) {
    Logger.log('EXCEPCION FUERA DEL TRY DE getFinanzasData: ' + e.message);
  }
}
function verificarCeldasCALC_FINANCIERO() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  ['B8','B10','B12','B14','B17','B42','B45','B46','B48'].forEach(function(celda) {
    var valor = cf.getRange(celda).getValue();
    Logger.log(celda + ' = [' + valor + '] tipo: ' + typeof valor);
  });
}
function verFormulasRotas() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  ['B8','B12','B14','B17','B42','B45','B46','B48'].forEach(function(celda) {
    var formula = cf.getRange(celda).getFormula();
    Logger.log(celda + ' => ' + formula);
  });
}
function verificarDB_INGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_INGRESOS');
  Logger.log('Hoja existe: ' + (hoja !== null));
  if (hoja) {
    Logger.log('Última fila: ' + hoja.getLastRow());
    Logger.log('Última columna: ' + hoja.getLastColumn());
    Logger.log('Valor B2: ' + hoja.getRange('B2').getValue());
    Logger.log('Valor D2: ' + hoja.getRange('D2').getValue());
    Logger.log('Valor H2: ' + hoja.getRange('H2').getValue());
  }

  // Probar la fórmula directamente en una celda de prueba
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  cf.getRange('Z1').setFormula('=SUMPRODUCT((DB_INGRESOS!D2:D5000="2026-06")*(DB_INGRESOS!B2:B5000="EMP-01")*(DB_INGRESOS!H2:H5000))');
  Utilities.sleep(1000);
  Logger.log('Resultado prueba Z1: ' + cf.getRange('Z1').getValue());
  cf.getRange('Z1').clearContent();
}
function verificarColumnaD_INGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_INGRESOS');
  var lastRow = hoja.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var formula = hoja.getRange('D'+i).getFormula();
    var valor = hoja.getRange('D'+i).getValue();
    Logger.log('D'+i+': formula=[' + formula + '] valor=[' + valor + ']');
  }
}
function corregirFormulaD_INGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_INGRESOS');
  var lastRow = hoja.getLastRow();

  // Eliminar el ARRAYFORMULA roto
  hoja.getRange('D2:D' + lastRow).clearContent();
  SpreadsheetApp.flush();

  // Escribir el período como VALOR fijo (string), no como fórmula,
  // calculado directamente en JavaScript desde la columna C (fecha)
  var datos = hoja.getRange('C2:C' + lastRow).getValues();
  var periodos = datos.map(function(f) {
    if (!(f[0] instanceof Date)) return [''];
    var anio = f[0].getFullYear();
    var mes  = String(f[0].getMonth()+1).padStart(2,'0');
    return [anio + '-' + mes];
  });
  hoja.getRange('D2:D' + lastRow).setValues(periodos);

  SpreadsheetApp.flush();
  Logger.log('✅ Columna D corregida con valores fijos');
  for (var i = 2; i <= lastRow; i++) {
    Logger.log('D'+i+': ' + hoja.getRange('D'+i).getValue());
  }
}
function corregirFormulaD_INGRESOS_v2() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_INGRESOS');
  var lastRow = hoja.getLastRow();

  // Forzar formato de texto en la columna D antes de escribir
  hoja.getRange('D2:D' + lastRow).setNumberFormat('@');

  var datos = hoja.getRange('C2:C' + lastRow).getValues();
  var periodos = datos.map(function(f) {
    if (!(f[0] instanceof Date)) return [''];
    var anio = f[0].getFullYear();
    var mes  = String(f[0].getMonth()+1).padStart(2,'0');
    return [String(anio + '-' + mes)];
  });
  hoja.getRange('D2:D' + lastRow).setValues(periodos);

  SpreadsheetApp.flush();
  Logger.log('✅ Columna D corregida con formato texto forzado');
  for (var i = 2; i <= lastRow; i++) {
    var v = hoja.getRange('D'+i).getValue();
    Logger.log('D'+i+': [' + v + '] tipo: ' + typeof v);
  }

  // Verificar que CALC_FINANCIERO ya no tiene #REF!
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  SpreadsheetApp.flush();
  Utilities.sleep(1500);
  Logger.log('');
  Logger.log('=== VERIFICACIÓN CALC_FINANCIERO ===');
  Logger.log('B8: ' + cf.getRange('B8').getValue());
  Logger.log('B12: ' + cf.getRange('B12').getValue());
  Logger.log('B42: ' + cf.getRange('B42').getValue());
  Logger.log('B45: ' + cf.getRange('B45').getValue());
}
function verificarB43_B44() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');

  ['B43','B44'].forEach(function(celda) {
    var formula = cf.getRange(celda).getFormula();
    var valor = cf.getRange(celda).getValue();
    Logger.log(celda + ' formula=[' + formula + '] valor=[' + valor + ']');
  });
}
function verificarB19() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  var formula = cf.getRange('B19').getFormula();
  var valor = cf.getRange('B19').getValue();
  Logger.log('B19 formula=[' + formula + '] valor=[' + valor + ']');
}
function verificarColumnasDB_CXC() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_CXC');
  var lastRow = hoja.getLastRow();

  Logger.log('Última fila: ' + lastRow);
  for (var i = 2; i <= lastRow; i++) {
    var k = hoja.getRange('K'+i).getValue();
    var l = hoja.getRange('L'+i).getValue();
    var formK = hoja.getRange('K'+i).getFormula();
    var formL = hoja.getRange('L'+i).getFormula();
    Logger.log('Fila '+i+': K=['+k+'] formK=['+formK+'] | L=['+l+'] formL=['+formL+']');
  }
}
function corregirDB_CXC_K_L() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_CXC');
  var lastRow = hoja.getLastRow();

  // Leer columnas necesarias: A(id), H(fecha_vencimiento), I(monto_original), J(monto_pagado)
  var datos = hoja.getRange('A2:J' + lastRow).getValues();
  var hoy = new Date();

  var saldos = [];
  var dias = [];

  datos.forEach(function(f) {
    var id = f[0];
    if (id === '') {
      saldos.push(['']);
      dias.push(['']);
      return;
    }
    var fechaVenc = f[7]; // columna H
    var montoOriginal = Number(f[8]) || 0; // columna I
    var montoPagado = Number(f[9]) || 0;   // columna J

    var saldo = montoOriginal - montoPagado;
    saldos.push([saldo]);

    if (fechaVenc instanceof Date) {
      var diasCalc = Math.floor((fechaVenc - hoy) / (1000*60*60*24));
      dias.push([diasCalc]);
    } else {
      dias.push(['']);
    }
  });

  // Limpiar fórmulas rotas y escribir valores fijos
  hoja.getRange('K2:K' + lastRow).clearContent();
  hoja.getRange('L2:L' + lastRow).clearContent();
  SpreadsheetApp.flush();

  hoja.getRange('K2:K' + lastRow).setValues(saldos);
  hoja.getRange('L2:L' + lastRow).setValues(dias);

  SpreadsheetApp.flush();
  Logger.log('✅ Columnas K y L de DB_CXC corregidas con valores fijos');

  for (var i = 2; i <= lastRow; i++) {
    Logger.log('Fila '+i+': K=['+hoja.getRange('K'+i).getValue()+'] L=['+hoja.getRange('L'+i).getValue()+']');
  }

  // Verificar CALC_FINANCIERO completo
  Utilities.sleep(1500);
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('');
  Logger.log('=== VERIFICACIÓN FINAL CALC_FINANCIERO ===');
  ['B8','B12','B17','B19','B42','B43','B44','B45','B46','B48'].forEach(function(celda) {
    Logger.log(celda + ': ' + cf.getRange(celda).getValue());
  });
}
function verificarFinanzasCompleto() {
  var r = getFinanzasData();
  Logger.log(JSON.stringify(r));
}
function verificarInventarioCompleto() {
  var r = getInventarioData();
  Logger.log('Tipo: ' + typeof r);
  Logger.log(JSON.stringify(r));
}
function verificarCALC_SUPPLY() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cs = ss.getSheetByName('CALC_SUPPLY');
  ['B7','B9','B10','B12','B13','B17','B18','B25'].forEach(function(celda) {
    var formula = cs.getRange(celda).getFormula();
    var valor = cs.getRange(celda).getValue();
    Logger.log(celda + ' formula=[' + formula + '] valor=[' + valor + ']');
  });
}
function verificarDB_MOVIMIENTOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_MOVIMIENTOS');
  var lastRow = hoja.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var c = hoja.getRange('C'+i).getValue();
    var j = hoja.getRange('J'+i).getValue();
    var formC = hoja.getRange('C'+i).getFormula();
    var formJ = hoja.getRange('J'+i).getFormula();
    Logger.log('Fila '+i+': C=['+c+'] formC=['+formC+'] | J=['+j+'] formJ=['+formJ+']');
  }
}
function corregirDB_MOVIMIENTOS_J() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_MOVIMIENTOS');
  var lastRow = hoja.getLastRow();

  // Leer columnas: H(tipo_movimiento), I(cantidad)
  var datos = hoja.getRange('A2:I' + lastRow).getValues();

  var impactos = datos.map(function(f) {
    var id = f[0];
    if (id === '') return [''];
    var tipo = String(f[7] || '');
    var cant = Number(f[8]) || 0;

    var impacto;
    if (tipo === 'SALIDA_VENTA' || tipo === 'AJUSTE_NEGATIVO' || tipo === 'DEVOLUCION_PROVEEDOR') {
      impacto = -cant;
    } else if (tipo === 'TRASLADO') {
      impacto = 0;
    } else {
      impacto = cant;
    }
    return [impacto];
  });

  hoja.getRange('J2:J' + lastRow).clearContent();
  SpreadsheetApp.flush();
  hoja.getRange('J2:J' + lastRow).setValues(impactos);

  SpreadsheetApp.flush();
  Logger.log('✅ Columna J de DB_MOVIMIENTOS corregida');

  for (var i = 2; i <= lastRow; i++) {
    Logger.log('Fila '+i+': J=['+hoja.getRange('J'+i).getValue()+']');
  }

  // Verificar CALC_SUPPLY
  Utilities.sleep(1500);
  var cs = ss.getSheetByName('CALC_SUPPLY');
  Logger.log('');
  Logger.log('=== VERIFICACIÓN CALC_SUPPLY ===');
  Logger.log('B17: ' + cs.getRange('B17').getValue());
  Logger.log('B18: ' + cs.getRange('B18').getValue());
}
function verificarInventarioFinal() {
  var r = getInventarioData();
  Logger.log(JSON.stringify(r));
}
function auditarDatosParaCompra() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  Logger.log('=== AUDITORÍA DATOS PARA NUEVA COMPRA/GASTO ===');

  var hP = ss.getSheetByName('MSTR_PROVEEDORES');
  Logger.log('MSTR_PROVEEDORES: ' + hP.getLastRow() + ' filas, ' + hP.getLastColumn() + ' cols');

  var hPC = ss.getSheetByName('MSTR_PLAN_CUENTAS');
  Logger.log('MSTR_PLAN_CUENTAS: ' + hPC.getLastRow() + ' filas');
  var cuentas = hPC.getRange(2,1,hPC.getLastRow()-1,4).getValues();
  cuentas.forEach(function(c) {
    if (c[0]!=='') Logger.log('  ' + c[0] + ' - ' + c[2] + ' (' + c[3] + ')');
  });

  var hC = ss.getSheetByName('DB_COMPRAS');
  Logger.log('DB_COMPRAS columnas:');
  hC.getRange(1,1,1,hC.getLastColumn()).getValues()[0].forEach(function(h,i){
    if(h!=='') Logger.log('  col '+(i+1)+': '+h);
  });
}
function testValidarRUT() {
  Logger.log('60.803.000-K válido: ' + validarRUT('60.803.000-K'));
  Logger.log('76.354.771-K válido: ' + validarRUT('76.354.771-K'));
  Logger.log('Búsqueda proveedores "test": ' + JSON.stringify(buscarProveedores('test')));
  Logger.log('Búsqueda cuentas "suministros": ' + JSON.stringify(buscarCuentasContables('suministros')));
}
function normalizarCuentasOriginales() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('MSTR_PLAN_CUENTAS');
  var filas = hoja.getLastRow();
  var datos = hoja.getRange(2,1,filas-1,8).getValues();

  for (var i = 0; i < datos.length; i++) {
    var id = String(datos[i][0]);
    var nivelActual = datos[i][7];
    if (nivelActual === 1) {
      hoja.getRange(i+2, 8).setValue('GRUPO');
      hoja.getRange(i+2, 9).setValue(false); // imputa_directo
    } else if (nivelActual === 2) {
      hoja.getRange(i+2, 8).setValue('CLASE');
      hoja.getRange(i+2, 9).setValue(false);
    } else if (nivelActual === 3) {
      hoja.getRange(i+2, 8).setValue('CUENTA');
      hoja.getRange(i+2, 9).setValue(true);
    }
  }

  // Corregir id_cuenta_padre faltante en las originales
  hoja.getRange('E4').setValue('4.1');   // 4.1.1 -> padre 4.1
  hoja.getRange('E7').setValue('5.2');   // 5.2.1 -> padre 5.2

  SpreadsheetApp.flush();
  Logger.log('✅ Cuentas originales normalizadas');

  var verif = hoja.getRange(2,1,filas-1,8).getValues();
  verif.forEach(function(f) {
    if (f[0]!=='') Logger.log(f[0] + ' | nivel=[' + f[7] + ']');
  });
}
function testProcesarCompra() {
  var datos = {
    esMercaderia: false,
    tipoDoc: 'BOLETA_HONORARIOS',
    proveedor: { id: 'PRV-20260617-084029-660', nombre: 'Proveedor Test H18 SpA' },
    fecha: '2026-06-20',
    numeroDoc: 'BH-TEST-001',
    condicionPago: 'CONTADO',
    cuentaContable: { id: '5.2.5', nombre: 'Honorarios' },
    glosa: 'Servicio de prueba honorarios',
    montoNeto: 100000,
    centroCosto: ''
  };
  var r = procesarNuevaCompra(datos);
  Logger.log(JSON.stringify(r));
}
function verificarValidacionDB_EGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var validacion = hoja.getRange('E6').getDataValidation();
  if (validacion) {
    Logger.log('Tipo: ' + validacion.getCriteriaType());
    Logger.log('Valores: ' + JSON.stringify(validacion.getCriteriaValues()));
  } else {
    Logger.log('Sin validación en E6 directamente — revisar si es de toda la columna');
  }

  // Verificar también E2 a E10 por si la validación varía
  for (var i = 2; i <= 10; i++) {
    var v = hoja.getRange('E'+i).getDataValidation();
    Logger.log('E'+i+': ' + (v ? v.getCriteriaType() : 'sin validación'));
  }
}
function verVerdaderoRangoValidacion() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var validacion = hoja.getRange('E6').getDataValidation();
  var criterios = validacion.getCriteriaValues();
  var rango = criterios[0]; // el rango de celdas permitido

  if (rango && rango.getValues) {
    var valores = rango.getValues();
    Logger.log('Rango: ' + rango.getA1Notation());
    Logger.log('Hoja del rango: ' + rango.getSheet().getName());
    valores.forEach(function(v) {
      if (v[0] !== '') Logger.log('Valor permitido: [' + v[0] + ']');
    });
  } else {
    Logger.log('No es un rango: ' + JSON.stringify(criterios));
  }
}
function actualizarListaValidacionCuentas() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hojaListas = ss.getSheetByName('LISTAS_VALIDACION');
  var hojaPlan = ss.getSheetByName('MSTR_PLAN_CUENTAS');

  // Obtener todas las cuentas tipo CUENTA (imputables)
  var filasPlan = hojaPlan.getLastRow();
  var datosPlan = hojaPlan.getRange(2,1,filasPlan-1,8).getValues();
  var cuentasImputables = [];
  datosPlan.forEach(function(f) {
    if (f[0] !== '' && String(f[7]) === 'CUENTA') {
      cuentasImputables.push([f[0]]);
    }
  });

  Logger.log('Cuentas imputables encontradas: ' + cuentasImputables.length);

  // Limpiar y reescribir columna E de LISTAS_VALIDACION
  hojaListas.getRange('E2:E1000').clearContent();
  hojaListas.getRange(2, 5, cuentasImputables.length, 1).setValues(cuentasImputables);

  SpreadsheetApp.flush();
  Logger.log('✅ Lista de validación actualizada');

  cuentasImputables.forEach(function(c) {
    Logger.log('  ' + c[0]);
  });
}
function testProcesarCompra() {
  var datos = {
    esMercaderia: false,
    tipoDoc: 'BOLETA_HONORARIOS',
    proveedor: { id: 'PRV-20260617-084029-660', nombre: 'Proveedor Test H18 SpA' },
    fecha: '2026-06-20',
    numeroDoc: 'BH-TEST-001',
    condicionPago: 'CONTADO',
    cuentaContable: { id: '5.2.5', nombre: 'Honorarios' },
    glosa: 'Servicio de prueba honorarios',
    montoNeto: 100000,
    centroCosto: ''
  };
  var r = procesarNuevaCompra(datos);
  Logger.log(JSON.stringify(r));
}
function verificarEstadoDB_EGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var lastRow = hoja.getLastRow();
  Logger.log('Última fila con getLastRow: ' + lastRow);
  Logger.log('Última fila real (_ultimaFilaReal): ' + _ultimaFilaReal(hoja));

  for (var i = 1; i <= lastRow + 2; i++) {
    var a = hoja.getRange('A'+i).getValue();
    var e = hoja.getRange('E'+i).getValue();
    Logger.log('Fila '+i+': A=['+a+'] E=['+e+']');
  }
}
function limpiarFilasFallidasDB_EGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');

  // Eliminar filas 6 y 7 que quedaron incompletas
  hoja.deleteRows(6, 2);

  SpreadsheetApp.flush();
  Logger.log('✅ Filas incompletas eliminadas');
  Logger.log('Última fila ahora: ' + hoja.getLastRow());
}
function diagnosticarUltimaFilaEgresos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  Logger.log('getLastRow: ' + hoja.getLastRow());
  Logger.log('_ultimaFilaReal: ' + _ultimaFilaReal(hoja));

  for (var i = 1; i <= 10; i++) {
    Logger.log('Fila '+i+': A=['+hoja.getRange('A'+i).getValue()+']');
  }
}
function testEscrituraDirectaE() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');

  try {
    hoja.getRange('E8').setValue('5.2.5');
    Logger.log('✅ Escritura exitosa: ' + hoja.getRange('E8').getValue());
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
  }

  hoja.getRange('E8').clearContent();
}
function verificarValidacionE8() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var validacion = hoja.getRange('E8').getDataValidation();
  var criterios = validacion.getCriteriaValues();
  var rango = criterios[0];

  Logger.log('Rango: ' + rango.getA1Notation());
  Logger.log('Hoja: ' + rango.getSheet().getName());
  var valores = rango.getValues();
  valores.forEach(function(v) {
    if (v[0] !== '') Logger.log('Permitido: [' + v[0] + ']');
  });

  // Verificar también la hoja LISTAS_VALIDACION directamente
  var hojaListas = ss.getSheetByName('LISTAS_VALIDACION');
  Logger.log('');
  Logger.log('=== Contenido real de LISTAS_VALIDACION columna E ===');
  for (var i = 1; i <= 20; i++) {
    var v = hojaListas.getRange('E'+i).getValue();
    if (v !== '') Logger.log('E'+i+': [' + v + ']');
  }
}
function eliminarValidacionCuentaEgresos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');

  hoja.getRange('E2:E5000').clearDataValidations();
  hoja.getRange('E2:E5000').setNumberFormat('@');

  SpreadsheetApp.flush();
  Logger.log('✅ Validación eliminada de DB_EGRESOS columna E');

  hoja.getRange('E8').setValue('5.2.5');
  Logger.log('Prueba E8: ' + hoja.getRange('E8').getValue());
  hoja.getRange('E8').clearContent();
}
function limpiarYProbarDeNuevo() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var lastRow = hoja.getLastRow();

  // Eliminar filas residuales desde la 6 en adelante (las que tenían A pero no E)
  if (lastRow > 5) {
    hoja.deleteRows(6, lastRow - 5);
  }
  SpreadsheetApp.flush();
  Logger.log('Última fila ahora: ' + hoja.getLastRow());

  var datos = {
    esMercaderia: false,
    tipoDoc: 'BOLETA_HONORARIOS',
    proveedor: { id: 'PRV-20260617-084029-660', nombre: 'Proveedor Test H18 SpA' },
    fecha: '2026-06-20',
    numeroDoc: 'BH-TEST-001',
    condicionPago: 'CONTADO',
    cuentaContable: { id: '5.2.5', nombre: 'Honorarios' },
    glosa: 'Servicio de prueba honorarios',
    montoNeto: 100000,
    centroCosto: ''
  };
  var r = procesarNuevaCompra(datos);
  Logger.log(JSON.stringify(r));
}
function diagnosticoRapidoFinanzasCobranza() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('=== CALC_FINANCIERO ===');
  ['B8','B12','B17','B19','B42','B45'].forEach(function(c) {
    Logger.log(c + ': ' + cf.getRange(c).getValue());
  });

  Logger.log('');
  Logger.log('=== getFinanzasData() directo ===');
  var r = getFinanzasData();
  Logger.log('Es null: ' + (r === null));
  Logger.log('ok===false: ' + (r && r.ok === false));
  if (r && r.ok === false) Logger.log('Error: ' + r.errorReal);

  Logger.log('');
  Logger.log('=== getCobranzaData() directo ===');
  var r2 = getCobranzaData();
  Logger.log('Es null: ' + (r2 === null));
}
function diagnosticoB10yB12() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('B10 valor: ' + cf.getRange('B10').getValue());
  Logger.log('B10 formula: ' + cf.getRange('B10').getFormula());
  Logger.log('B12 formula: ' + cf.getRange('B12').getFormula());
}
function diagnosticarColumnaD_EGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var lastRow = hoja.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var d = hoja.getRange('D'+i).getValue();
    var formD = hoja.getRange('D'+i).getFormula();
    Logger.log('Fila '+i+': D=['+d+'] formula=['+formD+']');
  }
}
function corregirFormulaD_EGRESOS() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hoja = ss.getSheetByName('DB_EGRESOS');
  var lastRow = hoja.getLastRow();

  // Forzar formato texto en la columna D antes de escribir
  hoja.getRange('D2:D' + lastRow).setNumberFormat('@');

  var datos = hoja.getRange('C2:C' + lastRow).getValues();
  var periodos = datos.map(function(f) {
    if (!(f[0] instanceof Date)) return [''];
    var anio = f[0].getFullYear();
    var mes  = String(f[0].getMonth()+1).padStart(2,'0');
    return [String(anio + '-' + mes)];
  });
  hoja.getRange('D2:D' + lastRow).setValues(periodos);

  SpreadsheetApp.flush();
  Logger.log('✅ Columna D de DB_EGRESOS corregida');
  for (var i = 2; i <= lastRow; i++) {
    var v = hoja.getRange('D'+i).getValue();
    Logger.log('D'+i+': [' + v + '] tipo: ' + typeof v);
  }

  Utilities.sleep(1500);
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('');
  Logger.log('=== VERIFICACIÓN FINAL ===');
  ['B8','B10','B12','B42','B45'].forEach(function(c) {
    Logger.log(c + ': ' + cf.getRange(c).getValue());
  });
}
function buscarValidacionesDB_COMPRAS() {
  var ss = _getSS();
  var hC = ss.getSheetByName('DB_COMPRAS');
  for (var col = 1; col <= 30; col++) {
    var v = hC.getRange(2, col).getDataValidation();
    if (v) {
      var letra = String.fromCharCode(64 + col);
      Logger.log('Columna ' + letra + ' (col ' + col + '): TIENE VALIDACIÓN - ' + v.getCriteriaType());
    }
  }
  Logger.log('Búsqueda completada');
}
function eliminarValidacionesDB_COMPRAS() {
  var ss = _getSS();
  var hC = ss.getSheetByName('DB_COMPRAS');

  // Eliminar validaciones de las columnas D, E, L, S, W (y por seguridad todo el rango usado)
  hC.getRange('A2:AD5000').clearDataValidations();

  // Forzar formato texto en columnas de ID para evitar conversión a fecha
  hC.getRange('D2:E5000').setNumberFormat('@');
  hC.getRange('S2:S5000').setNumberFormat('@');

  SpreadsheetApp.flush();
  Logger.log('✅ Validaciones eliminadas de DB_COMPRAS');

  // Verificar
  var quedan = 0;
  for (var col = 1; col <= 30; col++) {
    var v = hC.getRange(2, col).getDataValidation();
    if (v) quedan++;
  }
  Logger.log('Validaciones restantes: ' + quedan);
}
function limpiarFilaFallidaCompra() {
  var ss = _getSS();
  var hC = ss.getSheetByName('DB_COMPRAS');
  var lastRow = hC.getLastRow();
  Logger.log('Última fila antes: ' + lastRow);

  // Eliminar fila 2 (la que quedó a medio escribir)
  if (lastRow >= 2) {
    hC.deleteRow(2);
  }

  SpreadsheetApp.flush();
  Logger.log('Última fila después: ' + hC.getLastRow());

  // Revertir el stock que se sumó incorrectamente con esa compra fallida
  // (el movimiento SÍ se registró con cantidad 20, pero la línea de compra falló)
  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
    if (f[0]==='') return;
    if (String(f[0]) === 'PRD-20260616-222404-458') {
      Logger.log('Stock actual H07: ' + f[11] + ' — revertir a 0 para repetir prueba limpia');
      hP.getRange(idx+2, 12).setValue(0);
    }
  });

  // Eliminar el movimiento de entrada fallido también
  var hM = ss.getSheetByName('DB_MOVIMIENTOS');
  var filM = hM.getLastRow();
  var datosM = hM.getRange(2,1,filM-1,12).getValues();
  for (var i = datosM.length-1; i >= 0; i--) {
    if (String(datosM[i][7]) === 'ENTRADA_COMPRA' && Number(datosM[i][8]) === 20) {
      Logger.log('Eliminando movimiento fila ' + (i+2));
      hM.deleteRow(i+2);
      break;
    }
  }

  SpreadsheetApp.flush();
  Logger.log('✅ Limpieza completa — listo para repetir la prueba');
}
function verificarCompraNueva() {
  var ss = _getSS();
  var hC = ss.getSheetByName('DB_COMPRAS');
  var filC = _ultimaFilaReal(hC);
  Logger.log('Última fila DB_COMPRAS: ' + filC);
  if (filC >= 2) {
    var fila = hC.getRange(filC,1,1,30).getValues()[0];
    var headers = hC.getRange(1,1,1,30).getValues()[0];
    for (var i=0; i<30; i++) {
      if (fila[i] !== '') Logger.log(headers[i] + ': [' + fila[i] + ']');
    }
  }

  Logger.log('');
  Logger.log('=== Costo actual del producto ===');
  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
    if (String(f[0]) === 'PRD-20260616-222404-458') {
      Logger.log('Costo=['+f[7]+'] Stock=['+f[11]+']');
    }
  });
}
function corregirCostoActualH07() {
  var ss = _getSS();
  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
    if (String(f[0]) === 'PRD-20260616-222404-458') {
      // Stock=20, comprado a 1000, stock previo era 0, entonces costo = 1000
      hP.getRange(idx+2, 8).setValue(1000);
      Logger.log('✅ Costo corregido a 1000');
    }
  });
}
function auditarColumnasParaCRUD() {
  var ss = _getSS();

  Logger.log('=== MSTR_CLIENTES headers ===');
  var hC = ss.getSheetByName('MSTR_CLIENTES');
  var headersC = hC.getRange(1,1,1,23).getValues()[0];
  headersC.forEach(function(h,i) { if(h!=='') Logger.log('col '+(i+1)+' (idx '+i+'): '+h); });

  Logger.log('');
  Logger.log('=== MSTR_PROVEEDORES headers ===');
  var hP = ss.getSheetByName('MSTR_PROVEEDORES');
  var headersP = hP.getRange(1,1,1,24).getValues()[0];
  headersP.forEach(function(h,i) { if(h!=='') Logger.log('col '+(i+1)+' (idx '+i+'): '+h); });

  Logger.log('');
  Logger.log('=== Validaciones de datos en ambas hojas ===');
  for (var col=1; col<=23; col++) {
    var v = hC.getRange(2,col).getDataValidation();
    if (v) Logger.log('MSTR_CLIENTES col '+col+': '+v.getCriteriaType());
  }
  for (var col2=1; col2<=24; col2++) {
    var v2 = hP.getRange(2,col2).getDataValidation();
    if (v2) Logger.log('MSTR_PROVEEDORES col '+col2+': '+v2.getCriteriaType());
  }
}
function eliminarValidacionesClientesProveedores() {
  var ss = _getSS();

  var hC = ss.getSheetByName('MSTR_CLIENTES');
  hC.getRange('O2:O5000').clearDataValidations(); // condicion_pago col 15

  var hP = ss.getSheetByName('MSTR_PROVEEDORES');
  hP.getRange('K2:K5000').clearDataValidations(); // condicion_pago col 11
  hP.getRange('M2:M5000').clearDataValidations(); // moneda_habitual col 13

  SpreadsheetApp.flush();
  Logger.log('✅ Validaciones eliminadas de MSTR_CLIENTES y MSTR_PROVEEDORES');
}
function testCrearCliente() {
  var datos = {
    razonSocial: 'Cliente Prueba CRUD SpA',
    rut: '12345678-5',
    giro: 'Comercio',
    email: 'test@crud.cl',
    condicionPago: '30_DIAS',
    limiteCredito: 1000000
  };
  var r = procesarCliente(datos);
  Logger.log(JSON.stringify(r));
}
function testEditarYProveedor() {
  // Editar el cliente recién creado
  var datosEdit = {
    id: 'CLI-20260622-212045-385',
    razonSocial: 'Cliente Prueba CRUD SpA EDITADO',
    rut: '12345678-5',
    giro: 'Comercio Mayorista',
    email: 'editado@crud.cl',
    condicionPago: '60_DIAS',
    limiteCredito: 2000000
  };
  var r1 = procesarCliente(datosEdit);
  Logger.log('Edición cliente: ' + JSON.stringify(r1));

  // Crear proveedor
  var datosProv = {
    razonSocial: 'Proveedor Prueba CRUD SpA',
    rut: '76123456-3',
    giro: 'Distribución',
    condicionPago: '30_DIAS',
    plazoEntrega: 10
  };
  var r2 = procesarProveedor(datosProv);
  Logger.log('Creación proveedor: ' + JSON.stringify(r2));

  // Probar RUT inválido
  var datosInvalido = {
    razonSocial: 'Test RUT Malo',
    rut: '12345678-9'
  };
  var r3 = procesarCliente(datosInvalido);
  Logger.log('RUT inválido (debe fallar): ' + JSON.stringify(r3));
}
function testRUTsConocidos() {
  Logger.log('78.149.043-1 (empresa real): ' + validarRUT('78.149.043-1'));
  Logger.log('60.803.000-K (cliente real): ' + validarRUT('60.803.000-K'));
  Logger.log('76.354.771-K (proveedor real): ' + validarRUT('76.354.771-K'));
  Logger.log('76123456-3 (el que probé): ' + validarRUT('76123456-3'));
}
function diagnosticoUrgente() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log('LARGO DEL ARCHIVO: ' + html.length);
  Logger.log('CONTIENE pintarFormCliente: ' + (html.indexOf('pintarFormCliente') !== -1));
  Logger.log('CONTIENE nuevo-cliente: ' + (html.indexOf("'nuevo-cliente'") !== -1));
  Logger.log('ULTIMOS 200 CARACTERES: ' + html.substring(html.length-200));
}
function testUrgente() {
  Logger.log('=== FINANZAS ===');
  var f = getFinanzasData();
  Logger.log(JSON.stringify(f).substring(0,500));

  Logger.log('=== PROVEEDORES ===');
  var p = getProveedoresData();
  Logger.log(JSON.stringify(p));
}
function testDash() {
  Logger.log(JSON.stringify(getDashboardData()));
}
function verCeldasFinanciero() {
  var cf = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4').getSheetByName('CALC_FINANCIERO');
  ['B30','B38','B45','B46'].forEach(function(c) {
    var v = cf.getRange(c).getValue();
    Logger.log(c + ' = [' + v + '] tipo=' + typeof v);
  });
}
function verColumnasCXP() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('DB_CXP');
  Logger.log('Encabezados fila 1: ' + JSON.stringify(h.getRange(1,1,1,h.getLastColumn()).getValues()[0]));
  Logger.log('Fila 2 ejemplo: ' + JSON.stringify(h.getRange(2,1,1,h.getLastColumn()).getValues()[0]));
}
function checkFunciones() {
  Logger.log('listarClientes existe: ' + (typeof listarClientes === 'function'));
  Logger.log('listarProveedores existe: ' + (typeof listarProveedores === 'function'));
}
function verPlanCuentas() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('MSTR_PLAN_CUENTAS');
  Logger.log('Encabezados: ' + JSON.stringify(h.getRange(1,1,1,h.getLastColumn()).getValues()[0]));
  Logger.log('Fila 2: ' + JSON.stringify(h.getRange(2,1,1,h.getLastColumn()).getValues()[0]));
  Logger.log('Fila 3: ' + JSON.stringify(h.getRange(3,1,1,h.getLastColumn()).getValues()[0]));
}
function verificarTriggersActivos() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers instalados: ' + triggers.length);
  triggers.forEach(function(t) {
    Logger.log('Función: ' + t.getHandlerFunction() + ' | Tipo: ' + t.getEventType());
  });
}function crearHojasFase1() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');

  if (!ss.getSheetByName('DB_COBROS_CXC')) {
    var h1 = ss.insertSheet('DB_COBROS_CXC');
    h1.getRange(1,1,1,18).setValues([[
      'id_cobro','uuid_transaccion','id_empresa','id_cxc','id_cliente',
      'fecha_cobro','fecha_registro','monto','id_medio_pago','referencia_pago',
      'conciliado','fecha_conciliacion','anulado','motivo_anulacion',
      'anulado_por','fecha_anulacion','creado_en','creado_por'
    ]]);
    h1.setFrozenRows(1);
  }

  if (!ss.getSheetByName('MSTR_MEDIOS_PAGO')) {
    var h2 = ss.insertSheet('MSTR_MEDIOS_PAGO');
    h2.getRange(1,1,1,5).setValues([['id_medio_pago','id_empresa','nombre','activo','orden_visualizacion']]);
    h2.getRange(2,1,5,5).setValues([
      ['EFECTIVO','EMP-01','Efectivo',true,1],
      ['TRANSFERENCIA','EMP-01','Transferencia',true,2],
      ['CHEQUE','EMP-01','Cheque',true,3],
      ['TARJETA_DEBITO','EMP-01','Tarjeta Débito',true,4],
      ['TARJETA_CREDITO','EMP-01','Tarjeta Crédito',true,5]
    ]);
    h2.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ Hojas Fase 1 creadas');
}
function crearHojaPagosCXP() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('DB_PAGOS_CXP')) {
    var h = ss.insertSheet('DB_PAGOS_CXP');
    h.getRange(1,1,1,18).setValues([[
      'id_pago','uuid_transaccion','id_empresa','id_cxp','id_proveedor',
      'fecha_pago','fecha_registro','monto','id_medio_pago','referencia_pago',
      'conciliado','fecha_conciliacion','anulado','motivo_anulacion',
      'anulado_por','fecha_anulacion','creado_en','creado_por'
    ]]);
    h.setFrozenRows(1);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ DB_PAGOS_CXP creada');
}
function verB8B10() {
  var cf = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4').getSheetByName('CALC_FINANCIERO');
  Logger.log('B8 ingresos_neto: ' + cf.getRange('B8').getValue());
  Logger.log('B10 egresos_neto: ' + cf.getRange('B10').getValue());
}
function agregarColumnasAnulacion() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hV = ss.getSheetByName('DB_VENTAS');
  var hC = ss.getSheetByName('DB_COMPRAS');

  if (hV.getLastColumn() < 35) {
    hV.getRange(1,32,1,3).setValues([['motivo_anulacion','anulado_por','fecha_anulacion']]);
  }
  if (hC.getLastColumn() < 27) {
    hC.getRange(1,24,1,3).setValues([['motivo_anulacion','anulado_por','fecha_anulacion']]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ Columnas de anulación agregadas');
}
function testAnularVenta2() {
  var r = anularVenta('DOC-20260617-171310-202', 'Prueba completa de anulación');
  Logger.log(JSON.stringify(r));
}
function quitarValidacionTipoMovimiento() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('DB_MOVIMIENTOS');
  h.getRange('H2:H5000').clearDataValidations();
  SpreadsheetApp.flush();
  Logger.log('✅ Validación eliminada de DB_MOVIMIENTOS columna H');
}
function diagnosticarDocAnulado() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hV = ss.getSheetByName('DB_VENTAS');
  var filV = _ultimaFilaReal(hV);
  hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
    if (String(f[1]) === 'DOC-20260617-171310-720') {
      Logger.log('Producto: ' + f[13] + ' | Cantidad: ' + f[16] + ' | Estado: ' + f[27]);
    }
  });
}
function corregirStockManual() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hP = ss.getSheetByName('MSTR_PRODUCTOS');
  var filP = _ultimaFilaReal(hP);
  hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
    if (String(f[0]) === 'PRD-20260616-203147-190') {
      var stockActual = Number(f[11])||0;
      hP.getRange(idx+2, 12).setValue(stockActual + 5);
      Logger.log('Stock corregido: ' + stockActual + ' + 5 = ' + (stockActual+5));
    }
  });
  SpreadsheetApp.flush();
}
function verIngresoDoc202() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hI = ss.getSheetByName('DB_INGRESOS');
  var filI = _ultimaFilaReal(hI);
  hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
    if (String(f[12]).indexOf('171310-202') !== -1) {
      Logger.log(JSON.stringify(f));
    }
  });
}
function verTodosLosIngresos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var hI = ss.getSheetByName('DB_INGRESOS');
  var filI = _ultimaFilaReal(hI);
  hI.getRange(2,1,filI-1,15).getValues().forEach(function(f, idx) {
    if (f[0]==='') return;
    Logger.log('Fila '+(idx+2)+': monto='+f[7]+' | id_origen_externo=['+f[12]+']');
  });
}
function verificarCambiosNC() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log('Contiene irANCDesdeDoc: ' + (html.indexOf('irANCDesdeDoc') !== -1));
  Logger.log('Contiene Generar Nota de Crédito: ' + (html.indexOf('Generar Nota de Crédito') !== -1));
  Logger.log('Contiene d.idCliente en getDetalleDocumentoVenta: ' + (html.indexOf('idCliente: String(f[12]') !== -1));
}
function testTipoDocReal() {
  var d = getDetalleDocumentoVenta('FAC-20260623143228');
  Logger.log('tipoDoc recibido: [' + d.tipoDoc + ']');
  Logger.log('estado: [' + d.estado + ']');
}
function crearProductoDescuentoOmitido() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('MSTR_PRODUCTOS');
  var fil = _ultimaFilaReal(h);
  var yaExiste = false;
  if (fil >= 2) {
    h.getRange(2,1,fil-1,1).getValues().forEach(function(f) {
      if (String(f[0]) === 'PRD-DESCUENTO-OMITIDO') yaExiste = true;
    });
  }
  if (yaExiste) { Logger.log('Ya existe'); return; }

  var fila = new Array(28).fill('');
  fila[0] = 'PRD-DESCUENTO-OMITIDO';
  fila[1] = 'EMP-01';
  fila[2] = 'DESC-OMIT';
  fila[3] = 'Descuento Omitido';
  fila[6] = 'UN';
  fila[7] = 0;
  fila[8] = 0;
  fila[11] = 999999;
  fila[18] = true;
  fila[19] = false;
  fila[27] = true;

  var primeraFilaVacia = _ultimaFilaReal(h) + 1;
  h.getRange(primeraFilaVacia, 1, 1, 28).setValues([fila]);
  SpreadsheetApp.flush();
  Logger.log('✅ Producto Descuento Omitido creado');
}
function crearTablaAnticipos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('DB_ANTICIPOS')) {
    var h = ss.insertSheet('DB_ANTICIPOS');
    h.getRange(1,1,1,17).setValues([[
      'id_anticipo','id_empresa','id_entidad','tipo_entidad','origen_id',
      'monto_original','monto_usado','fecha_origen','motivo','estado',
      'fecha_devolucion','id_medio_pago_devolucion','anulado','motivo_anulacion',
      'anulado_por','creado_en','creado_por'
    ]]);
    h.setFrozenRows(1);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ DB_ANTICIPOS creada');
}
function agregarMedioAnticipo() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
  var fil = h.getLastRow();
  var yaExiste = false;
  h.getRange(2,1,fil-1,1).getValues().forEach(function(f){ if(String(f[0])==='ANTICIPO') yaExiste=true; });
  if (!yaExiste) {
    h.getRange(fil+1,1,1,5).setValues([['ANTICIPO','EMP-01','Anticipo / Saldo a favor',true,6]]);
    SpreadsheetApp.flush();
  }
  Logger.log('OK');
}
function crearSistemaCierres() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');

  if (!ss.getSheetByName('DB_CIERRES')) {
    var h = ss.insertSheet('DB_CIERRES');
    h.getRange(1,1,1,13).setValues([[
      'id_cierre','id_empresa','periodo','fecha_cierre','usuario',
      'ventas_netas','costo_venta','gastos_operacionales','utilidad_neta',
      'activo_total','pasivo_total','patrimonio_total','estado'
    ]]);
    h.setFrozenRows(1);
  }

  var cfg = ss.getSheetByName('CONFIG_SISTEMA');
  if (!cfg.getRange('AA1').getValue()) {
    cfg.getRange('AA1').setValue('clave_admin_cierre');
    cfg.getRange('AA2').setValue('vantis2026');
  }

  SpreadsheetApp.flush();
  Logger.log('✅ Sistema de cierres creado. Clave admin: ' + cfg.getRange('AA2').getValue());
}
function testBloqueoPeriodo() {
  var r = _validarPeriodoNoBloqueado('2026-06-15');
  Logger.log(JSON.stringify(r));
}
function verDBCierres() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('DB_CIERRES');
  var fil = h.getLastRow();
  Logger.log('Última fila: ' + fil);
  if (fil >= 2) {
    h.getRange(2,1,fil-1,13).getValues().forEach(function(f, idx) {
      Logger.log('Fila '+(idx+2)+': periodo=['+f[2]+'] estado=['+f[12]+']');
    });
  }
}
function corregirDBCierres() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('DB_CIERRES');

  // Eliminar filas duplicadas/corruptas
  h.getRange(2, 1, h.getLastRow()-1, 13).clearContent();
  SpreadsheetApp.flush();

  // Forzar columna C (periodo) a texto
  h.getRange('C2:C5000').setNumberFormat('@');
  SpreadsheetApp.flush();

  Logger.log('✅ DB_CIERRES limpiada. Vuelve a cerrar el período manualmente desde el ERP.');
}
function verificarSintaxisHTML() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  // Extraer solo el contenido del script
  var inicio = html.indexOf('<script>') + 8;
  var fin = html.lastIndexOf('</script>');
  var js = html.substring(inicio, fin);
  try {
    new Function(js);
    Logger.log('✅ Sintaxis JS válida');
  } catch(e) {
    Logger.log('❌ ERROR DE SINTAXIS: ' + e.message);
  }
}
function auditarUsoCuentasGasto() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('DB_EGRESOS');
  var fil = _ultimaFilaReal(h);
  var porCuenta = {};
  h.getRange(2,1,fil-1,18).getValues().forEach(function(f) {
    if (f[0]==='') return;
    var cuenta = String(f[4]);
    var desc = String(f[6]);
    var monto = Number(f[7])||0;
    if (!porCuenta[cuenta]) porCuenta[cuenta] = { count:0, total:0, ejemplos:[] };
    porCuenta[cuenta].count++;
    porCuenta[cuenta].total += monto;
    if (porCuenta[cuenta].ejemplos.length < 3) porCuenta[cuenta].ejemplos.push(desc);
  });
  Object.keys(porCuenta).sort().forEach(function(c) {
    Logger.log('Cuenta ' + c + ': ' + porCuenta[c].count + ' registros, total $' + porCuenta[c].total + ' | Ejemplos: ' + porCuenta[c].ejemplos.join(' | '));
  });
}
function ampliarPlanCuentas() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('MSTR_PLAN_CUENTAS');

  if (h.getLastColumn() < 13) {
    h.getRange(1,13,1,5).setValues([[
      'subtipo_eerr','id_centro_costo_default','naturaleza_flujo','es_depreciable','orden_presentacion'
    ]]);
  }

  var fil = h.getLastRow();
  var datos = h.getRange(2,1,fil-1,8).getValues();
  var clasificacion = {
    '4.1.1': ['INGRESO_OPERACIONAL','OPERACIONAL'],
    '1.1.4': ['NO_APLICA','OPERACIONAL'],
    '5.2.1': ['GASTO_OPERACIONAL','OPERACIONAL'],
    '5.2.2': ['GASTO_OPERACIONAL','OPERACIONAL'],
    '5.2.3': ['GASTO_OPERACIONAL','OPERACIONAL'],
    '5.2.4': ['GASTO_OPERACIONAL','OPERACIONAL'],
    '5.2.5': ['GASTO_OPERACIONAL','OPERACIONAL'],
    '5.2.6': ['GASTO_OPERACIONAL','OPERACIONAL']
  };

  for (var i=0; i<datos.length; i++) {
    var id = String(datos[i][0]);
    var subtipo = clasificacion[id] ? clasificacion[id][0] : 'NO_APLICA';
    var flujo = clasificacion[id] ? clasificacion[id][1] : 'NO_APLICA';
    h.getRange(i+2, 13).setValue(subtipo);
    h.getRange(i+2, 15).setValue(flujo);
    h.getRange(i+2, 16).setValue(false);
    h.getRange(i+2, 17).setValue(i+1);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ Plan de Cuentas ampliado y clasificado');
}
function medirFlujoCajaMultiMes() {
  var t0 = new Date().getTime();
  var r = getFlujoCajaMultiMes(6);
  var t1 = new Date().getTime();
  Logger.log('Tiempo total: ' + (t1-t0) + 'ms para ' + r.length + ' meses');
}
function medirCobranza() {
  var t0 = new Date().getTime();
  var r = getCobranzaData();
  var t1 = new Date().getTime();
  Logger.log('getCobranzaData: ' + (t1-t0) + 'ms');
}

function medirFinanzas() {
  var t0 = new Date().getTime();
  var r = getFinanzasData();
  var t1 = new Date().getTime();
  Logger.log('getFinanzasData: ' + (t1-t0) + 'ms');
}

function medirDashboard() {
  var t0 = new Date().getTime();
  var r = getDashboardData();
  var t1 = new Date().getTime();
  Logger.log('getDashboardData: ' + (t1-t0) + 'ms');
}

function medirPartesFinanzas() {
  var ss = _getSS();
  var ctx = getDatosContexto();
  var empresa = ctx.idEmpresa;

  var t0 = new Date().getTime();
  var anulados = _getDocumentosAnulados();
  var t1 = new Date().getTime();
  Logger.log('_getDocumentosAnulados: ' + (t1-t0) + 'ms');

  var t2 = new Date().getTime();
  var caja = _calcularCajaReal(empresa, ctx, anulados);
  var t3 = new Date().getTime();
  Logger.log('_calcularCajaReal: ' + (t3-t2) + 'ms');

  var hI = ss.getSheetByName('DB_INGRESOS');
  var filI = _ultimaFilaReal(hI);
  var t4 = new Date().getTime();
  var datosI = hI.getRange(2,1,filI-1,15).getValues();
  var t5 = new Date().getTime();
  Logger.log('Lectura DB_INGRESOS (' + filI + ' filas): ' + (t5-t4) + 'ms');

  var hE = ss.getSheetByName('DB_EGRESOS');
  var filE = _ultimaFilaReal(hE);
  var t6 = new Date().getTime();
  var datosE = hE.getRange(2,1,filE-1,18).getValues();
  var t7 = new Date().getTime();
  Logger.log('Lectura DB_EGRESOS (' + filE + ' filas): ' + (t7-t6) + 'ms');

  var cg = ss.getSheetByName('CALC_GERENCIAL');
  var t8 = new Date().getTime();
  var v = cg.getRange('B15').getValue();
  var t9 = new Date().getTime();
  Logger.log('Lectura celda CALC_GERENCIAL: ' + (t9-t8) + 'ms');
}
function testConsolidado() {
  var r = getConsolidadoFinanciero('2026-06');
  Logger.log(JSON.stringify(r));
}
function verificarFuncionFlujoCaja() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log('Contiene renderFlujoCajaReporte: ' + (html.indexOf('function renderFlujoCajaReporte') !== -1));
  Logger.log('Contiene getFlujoCajaMultiMes en llamada: ' + (html.indexOf('.getFlujoCajaMultiMes(') !== -1));
  var idx = html.indexOf('.getFlujoCajaMultiMes(');
  if (idx !== -1) {
    Logger.log('Contexto: ' + html.substring(idx-300, idx+100));
  }
}
function verificarFlujoCajaReporte2() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log('Contiene function renderFlujoCajaReporte: ' + (html.indexOf('function renderFlujoCajaReporte') !== -1));
}
function verificarContextoFuncion() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  var idx = html.indexOf('function renderFlujoCajaReporte');
  Logger.log('200 caracteres ANTES:');
  Logger.log(html.substring(idx-200, idx));
  Logger.log('---');
  Logger.log('100 caracteres DESPUÉS del inicio:');
  Logger.log(html.substring(idx, idx+100));
}
function confirmarFuncionPegada() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log(html.indexOf('function renderFlujoCajaReporte') !== -1 ? 'PEGADA CORRECTAMENTE' : 'NO ESTÁ — revisa de nuevo');
}
function crearPlantillaCargaClientes() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_CLIENTES')) {
    var h = ss.insertSheet('CARGA_CLIENTES');
    h.getRange(1,1,1,9).setValues([[
      'razon_social','rut','giro','email','telefono','direccion','comuna','condicion_pago','limite_credito'
    ]]);
    h.setFrozenRows(1);
    h.getRange(2,1,2,9).setValues([
      ['Ejemplo Cliente Uno SpA','76123456-3','Comercio','contacto@ejemplo.cl','+56912345678','Av. Siempre Viva 123','Santiago','CONTADO',0],
      ['Ejemplo Cliente Dos Ltda','77654321-8','Servicios','ventas@dos.cl','+56987654321','Calle Falsa 456','Providencia','30_DIAS',500000]
    ]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_CLIENTES creada');
}
function crearPlantillaCargaProveedores() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_PROVEEDORES')) {
    var h = ss.insertSheet('CARGA_PROVEEDORES');
    h.getRange(1,1,1,7).setValues([[
      'razon_social','rut','giro','email','telefono','condicion_pago','plazo_entrega_dias'
    ]]);
    h.setFrozenRows(1);
    h.getRange(2,1,2,7).setValues([
      ['Ejemplo Proveedor Uno SpA','76234567-1','Distribución','contacto@prov1.cl','+56911112222','30_DIAS',7],
      ['Ejemplo Proveedor Dos Ltda','77345678-9','Manufactura','ventas@prov2.cl','+56933334444','CONTADO',3]
    ]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_PROVEEDORES creada');
}
function crearPlantillaCargaProductos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_PRODUCTOS')) {
    var h = ss.insertSheet('CARGA_PRODUCTOS');
    h.getRange(1,1,1,7).setValues([[
      'sku','nombre','unidad','costo','precio_venta','stock_minimo','stock_critico'
    ]]);
    h.setFrozenRows(1);
    h.getRange(2,1,2,7).setValues([
      ['PRD-001','Producto Ejemplo Uno','UN',1000,1500,10,5],
      ['PRD-002','Producto Ejemplo Dos','KG',500,800,20,8]
    ]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_PRODUCTOS creada');
}
function crearPlantillaSaldosInventario() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_SALDOS_INVENTARIO')) {
    var h = ss.insertSheet('CARGA_SALDOS_INVENTARIO');
    h.getRange(1,1,1,3).setValues([['sku','cantidad_inicial','costo_unitario']]);
    h.setFrozenRows(1);
    h.getRange(2,1,2,3).setValues([
      ['PRD-001',50,1000],
      ['PRD-002',30,500]
    ]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_SALDOS_INVENTARIO creada');
}
function crearPlantillaSaldosCxC() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_SALDOS_CXC')) {
    var h = ss.insertSheet('CARGA_SALDOS_CXC');
    h.getRange(1,1,1,5).setValues([['rut_cliente','numero_documento','fecha_emision','fecha_vencimiento','monto']]);
    h.setFrozenRows(1);
    h.getRange(2,1,1,5).setValues([['76123456-3','FAC-INICIAL-001','2026-06-01','2026-07-01',150000]]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_SALDOS_CXC creada');
}
function crearPlantillaSaldosCxP() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  if (!ss.getSheetByName('CARGA_SALDOS_CXP')) {
    var h = ss.insertSheet('CARGA_SALDOS_CXP');
    h.getRange(1,1,1,5).setValues([['rut_proveedor','numero_documento','fecha_emision','fecha_vencimiento','monto']]);
    h.setFrozenRows(1);
    h.getRange(2,1,1,5).setValues([['76234567-1','FAC-PROV-INICIAL-001','2026-06-01','2026-07-01',80000]]);
  }
  SpreadsheetApp.flush();
  Logger.log('✅ CARGA_SALDOS_CXP creada');
}
function verificarSidebarCarga() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  Logger.log('Contiene data-mod="carga-masiva": ' + (html.indexOf('data-mod="carga-masiva"') !== -1));
  Logger.log('Contiene data-mod="cierres": ' + (html.indexOf('data-mod="cierres"') !== -1));
}
function verSidebarCompleto() {
  var html = HtmlService.createHtmlOutputFromFile('vantis_app').getContent();
  var idx = html.indexOf('data-mod="reportes"');
  Logger.log(html.substring(idx-50, idx+600));
}
function verHeadersProductos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('MSTR_PRODUCTOS');
  var headers = h.getRange(1,1,1,28).getValues()[0];
  headers.forEach(function(val, idx) {
    Logger.log('Columna ' + (idx+1) + ' (índice ' + idx + '): ' + val);
  });
}
function actualizarPlantillaCargaProductos() {
  var ss = SpreadsheetApp.openById('11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4');
  var h = ss.getSheetByName('CARGA_PRODUCTOS');
  h.getRange(1,1,h.getLastRow(),h.getLastColumn()).clearContent();

  h.getRange(1,1,1,8).setValues([[
    'sku','nombre','unidad','costo','precio_venta_neto','afecto_iva_SI_NO','maneja_stock_SI_NO','categoria'
  ]]);
  h.setFrozenRows(1);
  h.getRange(2,1,2,8).setValues([
    ['PRD-001','Producto Ejemplo Uno','UN',1000,1500,'SI','SI','General'],
    ['PRD-002','Servicio Ejemplo Dos','UN',0,5000,'SI','NO','Servicios']
  ]);

  // Agregar columnas extra: stock_minimo, stock_critico al final
  h.getRange(1,9,1,2).setValues([['stock_minimo','stock_critico']]);
  h.getRange(2,9,2,2).setValues([[10,5],[0,0]]);

  SpreadsheetApp.flush();
  Logger.log('✅ Plantilla CARGA_PRODUCTOS actualizada');
}
function verificarBackendActualizado() {
  Logger.log(procesarCargaMasivaProductos.toString().indexOf('afectoIva') !== -1 ? 'BACKEND ACTUALIZADO' : 'BACKEND VIEJO — falta pegar');
}
function diagnosticarSubtipoEERR() {
  var ss = _getSS();
  var h = ss.getSheetByName('MSTR_PLAN_CUENTAS');
  var fil = h.getLastRow();
  var datos = h.getRange(2,1,fil-1,17).getValues();
  Logger.log('=== PLAN DE CUENTAS — subtipo_eerr (columna M) ===');
  datos.forEach(function(f) {
    if (f[0]==='') return;
    Logger.log(f[0] + ' | ' + f[2] + ' | nivel=' + f[7] + ' | subtipo_eerr=' + (f[12] || '(VACÍO)'));
  });
}
function diagnosticarConfigSistema() {
  var ss = _getSS();
  var h = ss.getSheetByName('CONFIG_SISTEMA');
  var fil = h.getLastRow();
  var col = h.getLastColumn();
  Logger.log('=== CONFIG_SISTEMA: ' + fil + ' filas x ' + col + ' columnas ===');
  var headers = h.getRange(1,1,1,col).getValues()[0];
  var valores = h.getRange(2,1,1,col).getValues()[0];
  for (var i=0; i<col; i++) {
    var letra = String.fromCharCode(65 + (i % 26));
    if (i >= 26) letra = String.fromCharCode(64 + Math.floor(i/26)) + letra;
    Logger.log(letra + i%0 + (i+1) + ' | header="' + headers[i] + '" | valor="' + valores[i] + '"');
  }

  Logger.log('=== Hojas existentes con MSTR o ALERTA ===');
  ss.getSheets().forEach(function(s) {
    var n = s.getName();
    if (n.indexOf('MSTR') !== -1 || n.indexOf('ALERTA') !== -1 || n.indexOf('UMBRAL') !== -1) {
      Logger.log(n + ' (' + s.getLastRow() + ' filas, ' + s.getLastColumn() + ' columnas)');
    }
  });
}
function diagnosticarMediosPago() {
  var ss = _getSS();
  var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
  var fil = h.getLastRow();
  var col = h.getLastColumn();
  var headers = h.getRange(1,1,1,col).getValues()[0];
  Logger.log('=== MSTR_MEDIOS_PAGO headers ===');
  Logger.log(JSON.stringify(headers));
  var datos = h.getRange(2,1,fil-1,col).getValues();
  datos.forEach(function(f) { Logger.log(JSON.stringify(f)); });
}
function setupEstructuraConfiguracion() {
  var ss = _getSS();

  // 1. Columnas nuevas en CONFIG_SISTEMA
  var cfg = ss.getSheetByName('CONFIG_SISTEMA');
  var headersActuales = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
  var nuevasCols = ['nombre_fantasia','decimales','formato_fecha','requiere_confirmacion_anulacion'];
  nuevasCols.forEach(function(col) {
    if (headersActuales.indexOf(col) === -1) {
      var nuevaColIdx = cfg.getLastColumn() + 1;
      cfg.getRange(1, nuevaColIdx).setValue(col);
    }
  });
  // Valores default fila 2
  var headersNuevo = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
  function setDefault(nombreCol, valor) {
    var idx = headersNuevo.indexOf(nombreCol) + 1;
    if (idx > 0 && cfg.getRange(2, idx).getValue() === '') {
      cfg.getRange(2, idx).setValue(valor);
    }
  }
  setDefault('nombre_fantasia', '');
  setDefault('decimales', 0);
  setDefault('formato_fecha', 'dd-MM-yyyy');
  setDefault('requiere_confirmacion_anulacion', true);

  // 2. Columna es_default en MSTR_MEDIOS_PAGO
  var hMP = ss.getSheetByName('MSTR_MEDIOS_PAGO');
  var headersMP = hMP.getRange(1,1,1,hMP.getLastColumn()).getValues()[0];
  if (headersMP.indexOf('es_default') === -1) {
    var colIdx = hMP.getLastColumn() + 1;
    hMP.getRange(1, colIdx).setValue('es_default');
    var filMP = hMP.getLastRow();
    for (var i=2; i<=filMP; i++) {
      hMP.getRange(i, colIdx).setValue(i === 2); // EFECTIVO como default inicial
    }
  }

  // 3. Hoja CONFIG_COMERCIAL
  if (!ss.getSheetByName('CONFIG_COMERCIAL')) {
    var hCom = ss.insertSheet('CONFIG_COMERCIAL');
    hCom.getRange(1,1,1,7).setValues([[
      'id_empresa','meta_mensual_default','meta_anual','ticket_objetivo',
      'dias_credito_default','descuento_max_pct','lista_precio_default'
    ]]);
    hCom.getRange(2,1,1,7).setValues([['EMP-01',0,0,0,30,0,'GENERAL']]);
  }

  // 4. Hoja CONFIG_INVENTARIO
  if (!ss.getSheetByName('CONFIG_INVENTARIO')) {
    var hInv = ss.insertSheet('CONFIG_INVENTARIO');
    hInv.getRange(1,1,1,5).setValues([[
      'id_empresa','stock_critico_default','politica_costeo','permite_stock_negativo','dias_reposicion_default'
    ]]);
    hInv.getRange(2,1,1,5).setValues([['EMP-01',5,'CPP',false,7]]);
  }

  // 5. Hoja CONFIG_ALERTAS
  if (!ss.getSheetByName('CONFIG_ALERTAS')) {
    var hAl = ss.insertSheet('CONFIG_ALERTAS');
    hAl.getRange(1,1,1,7).setValues([[
      'id_empresa','liquidez_minima','margen_minimo','stock_critico_umbral',
      'cxc_vencida_max','cxp_urgente_max','score_minimo'
    ]]);
    hAl.getRange(2,1,1,7).setValues([['EMP-01',1.5,0.15,5,500000,500000,60]]);
  }

  Logger.log('✅ Estructura de Configuración creada/verificada correctamente');
}
function ampliarConfigComercial() {
  var ss = _getSS();
  var h = ss.getSheetByName('CONFIG_COMERCIAL');
  var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];

  var nuevasCols = [
    'margen_bruto_objetivo', 'crecimiento_mensual_esperado', 'venta_minima_diaria',
    'objetivo_clientes_nuevos', 'limite_credito_default',
    'bloquear_venta_sin_stock', 'autorizar_descuento_alto', 'permitir_modificar_precio',
    'mostrar_costo_productos'
  ];
  nuevasCols.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      h.getRange(1, h.getLastColumn()+1).setValue(col);
    }
  });

  var headersNuevo = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
  function setDefault(nombreCol, valor) {
    var idx = headersNuevo.indexOf(nombreCol) + 1;
    if (idx > 0) {
      var actual = h.getRange(2, idx).getValue();
      if (actual === '') h.getRange(2, idx).setValue(valor);
    }
  }
  setDefault('margen_bruto_objetivo', 0.25);
  setDefault('crecimiento_mensual_esperado', 0.05);
  setDefault('venta_minima_diaria', 0);
  setDefault('objetivo_clientes_nuevos', 0);
  setDefault('limite_credito_default', 0);
  setDefault('bloquear_venta_sin_stock', true);
  setDefault('autorizar_descuento_alto', false);
  setDefault('permitir_modificar_precio', true);
  setDefault('mostrar_costo_productos', true);

  Logger.log('✅ CONFIG_COMERCIAL ampliada correctamente');
}
function diagnosticarBotonCobro(idCliente) {
  var ss = _getSS();
  var hCxC = ss.getSheetByName('DB_CXC');
  var fil = _ultimaFilaReal(hCxC);
  Logger.log('=== DB_CXC para cliente ' + idCliente + ' ===');
  if (fil >= 2) {
    hCxC.getRange(2,1,fil-1,21).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[2]) !== idCliente) return;
      Logger.log('idCxc=' + f[0] + ' | doc=' + f[4] + ' | tipoDoc=' + f[5] + ' | monto=' + f[8] + ' | estado(col14)=' + f[13]);
    });
  }
  Logger.log('=== Resultado de getFichaCliente ===');
  var ficha = getFichaCliente(idCliente);
  if (ficha) {
    ficha.documentos.forEach(function(d) {
      Logger.log('doc=' + d.doc + ' | idCxc=' + d.idCxc + ' | saldoFinal=' + d.saldoFinal + ' | estado=' + d.estado);
    });
  }
}
function buscarIdClientePrueba() {
  var lista = listarClientes();
  lista.forEach(function(c) {
    if (c.nombre.indexOf('Prueba Negocio') !== -1) {
      Logger.log('ID encontrado: ' + c.id + ' | nombre: ' + c.nombre);
    }
  });
}
function diagnosticarBotonCobroTEST() {

  diagnosticarBotonCobro('Cliente Prueba Negocio SpA');

}
function diagnosticarFacturaFAC() {
  var detalle = getDetalleDocumentoVenta('FAC-20260624083735');
  Logger.log('Detalle encontrado: ' + (detalle !== null));
  if (detalle) Logger.log(JSON.stringify(detalle));

  var ss = _getSS();
  var hV = ss.getSheetByName('DB_VENTAS');
  var fil = _ultimaFilaReal(hV);
  Logger.log('=== Buscando FAC-20260624083735 en DB_VENTAS ===');
  var encontrado = false;
  hV.getRange(2,1,fil-1,5).getValues().forEach(function(f, idx) {
    if (String(f[4]).indexOf('20260624083735') !== -1) {
      Logger.log('Fila ' + (idx+2) + ': id=' + f[0] + ' | idDoc=' + f[1] + ' | numeroDoc=' + f[4]);
      encontrado = true;
    }
  });
  if (!encontrado) Logger.log('NINGUNA fila coincide con ese número de documento');
}
function diagnosticarIdCxcFrontend() {
  var ficha = getFichaCliente('CLI-20260615-223109-222');
  ficha.documentos.forEach(function(d) {
    Logger.log('doc=' + d.doc + ' | idCxc="' + d.idCxc + '" | saldoFinal=' + d.saldoFinal);
  });
}