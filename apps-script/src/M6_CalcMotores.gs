function crearCALC_SUPPLY() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_SUPPLY');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('Hoja CALC_SUPPLY no encontrada');
    return;
  }

  Logger.log('=== CONSTRUYENDO CALC_SUPPLY v1.0 ===');
  hoja.clearContents();
  hoja.clearFormats();

  var cfg     = ss.getSheetByName('CONFIG_SISTEMA');
  var fechaT2 = cfg.getRange('T2').getValue();
  var empresa = String(cfg.getRange('A2').getValue());

  var anio, mes;
  if (fechaT2 instanceof Date) {
    anio = fechaT2.getFullYear();
    mes  = fechaT2.getMonth() + 1;
  } else {
    var parts = String(fechaT2).substring(0,7).split('-');
    anio = Number(parts[0]);
    mes  = Number(parts[1]);
  }

  var periodo = anio + '-' + String(mes).padStart(2,'0');
  Logger.log('Período: ' + periodo + ' | Empresa: ' + empresa);

  var e = '"' + empresa + '"';

  var estructura = [
    [1,  '—CONTEXTO—',              'CALC_SUPPLY v1.0',   false],
    [2,  'periodo_ref',             periodo,              false],
    [3,  'id_empresa',              empresa,              false],
    [4,  'fecha_calculo',           new Date(),           false],
    [5,  '',                        '',                   false],
    [6,  '—INVENTARIO_GENERAL—',    '',                   false],
    [7,  'productos_activos',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*1)',
      true],
    [8,  'productos_con_stock',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000>0)*1)',
      true],
    [9,  'productos_sin_stock',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000=0)*1)',
      true],
    [10, 'productos_criticos',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000<=MSTR_PRODUCTOS!P2:P5000)*(MSTR_PRODUCTOS!P2:P5000>0)*1)',
      true],
    [11, 'productos_bajo_minimo',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000<MSTR_PRODUCTOS!N2:N5000)*(MSTR_PRODUCTOS!N2:N5000>0)*1)',
      true],
    [12, 'valor_inventario_costo',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000)*(MSTR_PRODUCTOS!H2:H5000))',
      true],
    [13, 'valor_inventario_venta',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000)*(MSTR_PRODUCTOS!I2:I5000))',
      true],
    [14, 'margen_inventario_potencial',
      '=B13-B12',
      true],
    [15, '',                        '',                   false],
    [16, '—MOVIMIENTOS_PERIODO—',   '',                   false],
    // CORRECCIÓN C1: excluir APERTURA de entradas
    [17, 'entradas_periodo',
      '=SUMPRODUCT((DB_MOVIMIENTOS!C2:C5000<>"")*(YEAR(DB_MOVIMIENTOS!C2:C5000)=' + anio + ')*(MONTH(DB_MOVIMIENTOS!C2:C5000)=' + mes + ')*(DB_MOVIMIENTOS!H2:H5000<>"APERTURA")*(DB_MOVIMIENTOS!J2:J5000>0)*(DB_MOVIMIENTOS!J2:J5000))',
      true],
    [18, 'salidas_periodo',
      '=ABS(SUMPRODUCT((DB_MOVIMIENTOS!C2:C5000<>"")*(YEAR(DB_MOVIMIENTOS!C2:C5000)=' + anio + ')*(MONTH(DB_MOVIMIENTOS!C2:C5000)=' + mes + ')*(DB_MOVIMIENTOS!J2:J5000<0)*(DB_MOVIMIENTOS!J2:J5000)))',
      true],
    [19, 'total_movimientos_periodo',
      '=SUMPRODUCT((DB_MOVIMIENTOS!C2:C5000<>"")*(YEAR(DB_MOVIMIENTOS!C2:C5000)=' + anio + ')*(MONTH(DB_MOVIMIENTOS!C2:C5000)=' + mes + ')*(DB_MOVIMIENTOS!E2:E5000<>"")*1)',
      true],
    [20, 'unidades_vendidas_periodo',
      '=SUMPRODUCT((DB_MOVIMIENTOS!C2:C5000<>"")*(YEAR(DB_MOVIMIENTOS!C2:C5000)=' + anio + ')*(MONTH(DB_MOVIMIENTOS!C2:C5000)=' + mes + ')*(DB_MOVIMIENTOS!H2:H5000="SALIDA_VENTA")*ABS(DB_MOVIMIENTOS!J2:J5000))',
      true],
    // CORRECCIÓN C1: excluir APERTURA de compradas
    [21, 'unidades_compradas_periodo',
      '=SUMPRODUCT((DB_MOVIMIENTOS!C2:C5000<>"")*(YEAR(DB_MOVIMIENTOS!C2:C5000)=' + anio + ')*(MONTH(DB_MOVIMIENTOS!C2:C5000)=' + mes + ')*(DB_MOVIMIENTOS!H2:H5000="ENTRADA_COMPRA")*(DB_MOVIMIENTOS!J2:J5000))',
      true],
    [22, 'ratio_rotacion',
      '=SUMPRODUCT((B12>0)*B18/MAX(B12;1))',
      true],
    [23, '',                        '',                   false],
    [24, '—ALERTAS_STOCK—',         '',                   false],
    [25, 'stock_total_unidades',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*(MSTR_PRODUCTOS!T2:T5000=TRUE)*(MSTR_PRODUCTOS!L2:L5000))',
      true],
    [26, 'pct_productos_criticos',
      '=SUMPRODUCT((B7>0)*B10/MAX(B7;1))',
      true],
    [27, 'pct_productos_sin_stock',
      '=SUMPRODUCT((B7>0)*B9/MAX(B7;1))',
      true],
    [28, 'dias_inventario_disponible',
      '=SUMPRODUCT((B18>0)*B25/MAX(B18/30;1))',
      true],
    [29, 'cobertura_dias_promedio',
      '=SUMPRODUCT((B20>0)*B25/MAX(B20/30;1))',
      true],
    [30, '',                        '',                   false],
    [31, '—HISTORIAL_PRECIOS—',     '',                   false],
    [32, 'registros_hist_precios',
      '=SUMPRODUCT((HIST_PRECIOS!B2:B5000=' + e + ')*(HIST_PRECIOS!A2:A5000<>"")*1)',
      true],
    [33, 'precio_promedio_compra',
      '=SUMPRODUCT((HIST_PRECIOS!B2:B5000=' + e + ')*(HIST_PRECIOS!I2:I5000))/MAX(B32;1)',
      true],
    [34, 'ultimo_precio_compra',
      '=SUMPRODUCT((HIST_PRECIOS!B2:B5000=' + e + ')*(HIST_PRECIOS!A2:A5000<>"")*(ROW(HIST_PRECIOS!A2:A5000)=SUMPRODUCT(MAX((HIST_PRECIOS!B2:B5000=' + e + ')*(HIST_PRECIOS!A2:A5000<>"")*ROW(HIST_PRECIOS!A2:A5000))))*(HIST_PRECIOS!I2:I5000))',
      true],
    [35, '',                        '',                   false],
    [36, '—VALIDACIONES—',          '',                   false],
    [37, 'val_productos_activos',   '=SUMPRODUCT((B7>0)*1)',   true],
    [38, 'val_valor_inventario',    '=SUMPRODUCT((B12>=0)*1)', true],
    [39, 'val_stock_no_negativo',   '=SUMPRODUCT((B25>=0)*1)', true],
    [40, 'val_criticos_ok',         '=SUMPRODUCT((B10>=0)*1)', true],
    [41, 'val_rotacion_ok',         '=SUMPRODUCT((B22>=0)*1)', true]
  ];

  estructura.forEach(function(item) {
    var fila = item[0]; var etiq = item[1];
    var valor = item[2]; var esForm = item[3];
    if (etiq !== '') hoja.getRange(fila, 1).setValue(etiq);
    if (valor !== '') {
      if (esForm) hoja.getRange(fila, 2).setFormula(valor);
      else hoja.getRange(fila, 2).setValue(valor);
    }
  });

  SpreadsheetApp.flush();

  [1,6,16,24,31,36].forEach(function(f) {
    var r = hoja.getRange(f,1,1,2);
    r.setBackground('#2C3E50');
    r.setFontColor('#FFFFFF');
    r.setFontWeight('bold');
  });
  [12,13,14,17,18,33,34].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('$#,##0'); });
  [26,27].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('0.00%'); });
  [7,8,9,10,11,19,20,21,25,32].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('#,##0'); });
  [22,28,29].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('#,##0.00'); });
  hoja.getRange(4,2).setNumberFormat('DD/MM/YYYY');
  hoja.setColumnWidth(1,240);
  hoja.setColumnWidth(2,160);

  Logger.log('✅ CALC_SUPPLY v1.0 construida y corregida');
}
function crearCALC_COMERCIAL() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_COMERCIAL');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('Hoja CALC_COMERCIAL no encontrada');
    return;
  }

  Logger.log('=== CONSTRUYENDO CALC_COMERCIAL v1.0 ===');
  hoja.clearContents();
  hoja.clearFormats();

  var cfg     = ss.getSheetByName('CONFIG_SISTEMA');
  var fechaT2 = cfg.getRange('T2').getValue();
  var empresa = String(cfg.getRange('A2').getValue());

  var anio, mes;
  if (fechaT2 instanceof Date) {
    anio = fechaT2.getFullYear();
    mes  = fechaT2.getMonth() + 1;
  } else {
    var parts = String(fechaT2).substring(0,7).split('-');
    anio = Number(parts[0]);
    mes  = Number(parts[1]);
  }

  var periodo = anio + '-' + String(mes).padStart(2,'0');
  Logger.log('Período: ' + periodo + ' | Empresa: ' + empresa);

  var e = '"' + empresa + '"';

  // CORRECCIÓN C2: filtro período usa YEAR+MONTH desde fecha_venta (col K)
  // sys_periodo (col L) almacena como Date — no usar para comparación
  var filtPer  = '(YEAR(DB_VENTAS!K2:K5000)=' + anio + ')*(MONTH(DB_VENTAS!K2:K5000)=' + mes + ')';
  var filtEmp  = '(DB_VENTAS!C2:C5000=' + e + ')';
  var filtVig  = '(DB_VENTAS!AB2:AB5000<>"ANULADO")';
  var filtBase = filtPer + '*' + filtEmp + '*' + filtVig;
  var filtDia  = '(INT(DB_VENTAS!K2:K5000)=INT(TODAY()))*(DB_VENTAS!C2:C5000=' + e + ')*(DB_VENTAS!AB2:AB5000<>"ANULADO")';

  var estructura = [
    [1,  '—CONTEXTO—',              'CALC_COMERCIAL v1.0', false],
    [2,  'periodo_ref',             periodo,               false],
    [3,  'id_empresa',              empresa,               false],
    [4,  'fecha_calculo',           new Date(),            false],
    [5,  '',                        '',                    false],
    [6,  '—VENTAS_PERIODO—',        '',                    false],
    [7,  'ventas_neto',             '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!T2:T5000))',  true],
    [8,  'ventas_bruto',            '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!V2:V5000))',  true],
    [9,  'costo_ventas',            '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!Q2:Q5000)*(DB_VENTAS!W2:W5000))', true],
    [10, 'margen_bruto',            '=B7-B9',              true],
    [11, 'margen_bruto_pct',        '=SUMPRODUCT((B7>0)*B10/MAX(B7;1))', true],
    [12, 'unidades_vendidas',       '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!Q2:Q5000))', true],
    [13, 'documentos_mes',          '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!A2:A5000<>"")*1)', true],
    [14, 'ticket_promedio_neto',    '=SUMPRODUCT((B13>0)*B7/MAX(B13;1))', true],
    [15, 'ticket_promedio_bruto',   '=SUMPRODUCT((B13>0)*B8/MAX(B13;1))', true],
    [16, 'unidades_por_doc',        '=SUMPRODUCT((B13>0)*B12/MAX(B13;1))', true],
    [17, '',                        '',                    false],
    [18, '—VENTAS_DIA—',            '',                    false],
    [19, 'ventas_dia_neto',         '=SUMPRODUCT((' + filtDia + ')*(DB_VENTAS!T2:T5000))', true],
    [20, 'ventas_dia_bruto',        '=SUMPRODUCT((' + filtDia + ')*(DB_VENTAS!V2:V5000))', true],
    [21, 'docs_dia',                '=SUMPRODUCT((' + filtDia + ')*(DB_VENTAS!A2:A5000<>"")*1)', true],
    [22, 'unidades_dia',            '=SUMPRODUCT((' + filtDia + ')*(DB_VENTAS!Q2:Q5000))', true],
    [23, '',                        '',                    false],
    [24, '—CLIENTES_PRODUCTOS—',    '',                    false],
    [25, 'clientes_activos',
      '=SUMPRODUCT((MSTR_CLIENTES!B2:B5000=' + e + ')*(MSTR_CLIENTES!R2:R5000=TRUE)*1)',
      true],
    [26, 'clientes_con_venta_mes',
      '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!M2:M5000<>"")*1)',
      true],
    [27, 'productos_activos',
      '=SUMPRODUCT((MSTR_PRODUCTOS!B2:B5000=' + e + ')*(MSTR_PRODUCTOS!AB2:AB5000=TRUE)*1)',
      true],
    [28, 'productos_con_venta_mes',
      '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!N2:N5000<>"")*1)',
      true],
    [29, 'pct_productos_activos_vendidos',
      '=SUMPRODUCT((B27>0)*B28/MAX(B27;1))',
      true],
    [30, '',                        '',                    false],
    [31, '—MARGEN_RENTABILIDAD—',   '',                    false],
    [32, 'margen_total_periodo',    '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!X2:X5000))', true],
    [33, 'margen_promedio_pct',     '=SUMPRODUCT(' + filtBase + '*(DB_VENTAS!Y2:Y5000))/MAX(B13;1)', true],
    [34, 'ventas_anuladas_mes',
      '=SUMPRODUCT(' + filtPer + '*' + filtEmp + '*(DB_VENTAS!AB2:AB5000="ANULADO")*(DB_VENTAS!T2:T5000))',
      true],
    [35, 'docs_anulados_mes',
      '=SUMPRODUCT(' + filtPer + '*' + filtEmp + '*(DB_VENTAS!AB2:AB5000="ANULADO")*(DB_VENTAS!A2:A5000<>"")*1)',
      true],
    [36, '',                        '',                    false],
    [37, '—VALIDACIONES—',          '',                    false],
    [38, 'val_ventas_positivas',    '=SUMPRODUCT((B7>=0)*1)',  true],
    [39, 'val_margen_positivo',     '=SUMPRODUCT((B10>=0)*1)', true],
    [40, 'val_docs_consistentes',   '=SUMPRODUCT((B13>=0)*1)', true],
    [41, 'val_ticket_consistente',  '=SUMPRODUCT((B14>=0)*1)', true],
    [42, 'val_clientes_activos',    '=SUMPRODUCT((B25>0)*1)',  true]
  ];

  estructura.forEach(function(item) {
    var fila = item[0]; var etiq = item[1];
    var valor = item[2]; var esForm = item[3];
    if (etiq !== '') hoja.getRange(fila, 1).setValue(etiq);
    if (valor !== '') {
      if (esForm) hoja.getRange(fila, 2).setFormula(valor);
      else hoja.getRange(fila, 2).setValue(valor);
    }
  });

  SpreadsheetApp.flush();

  [1,6,18,24,31,37].forEach(function(f) {
    var r = hoja.getRange(f,1,1,2);
    r.setBackground('#2C3E50');
    r.setFontColor('#FFFFFF');
    r.setFontWeight('bold');
  });
  [7,8,9,10,14,15,19,20,32,34].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('$#,##0'); });
  [11,29,33].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('0.00%'); });
  [12,13,16,21,22,25,26,27,28,35].forEach(function(f) { hoja.getRange(f,2).setNumberFormat('#,##0'); });
  hoja.getRange(4,2).setNumberFormat('DD/MM/YYYY');
  hoja.setColumnWidth(1,240);
  hoja.setColumnWidth(2,160);

  Logger.log('✅ CALC_COMERCIAL v1.0 construida y corregida');
}
function validarCALC_SUPPLY() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_SUPPLY');
  Logger.log('=== VALIDACIÓN CALC_SUPPLY v1.0 ===');
  Logger.log('');

  function leer(f) { return hoja.getRange(f, 2).getValue(); }

  // BLOQUE 2 — Inventario general
  Logger.log('── BLOQUE 2 INVENTARIO GENERAL ──');
  var prdAct  = Number(leer(7));
  var prdStk  = Number(leer(8));
  var prdSin  = Number(leer(9));
  var prdCrit = Number(leer(10));
  var prdMin  = Number(leer(11));
  var valCost = Number(leer(12));
  var valVta  = Number(leer(13));
  var margen  = Number(leer(14));

  Logger.log('productos_activos:        ' + prdAct);
  Logger.log('productos_con_stock:      ' + prdStk);
  Logger.log('productos_sin_stock:      ' + prdSin);
  Logger.log('productos_criticos:       ' + prdCrit);
  Logger.log('productos_bajo_minimo:    ' + prdMin);
  Logger.log('valor_inventario_costo:   $' + valCost.toLocaleString('es-CL'));
  Logger.log('valor_inventario_venta:   $' + valVta.toLocaleString('es-CL'));
  Logger.log('margen_inventario_pot:    $' + margen.toLocaleString('es-CL'));

  Logger.log(prdAct === 2  ? '✅ productos_activos OK'  : '❌ esperado 2 | obtenido ' + prdAct);
  Logger.log(prdStk === 1  ? '✅ con_stock OK'           : '❌ esperado 1 | obtenido ' + prdStk);
  Logger.log(prdSin === 1  ? '✅ sin_stock OK'           : '❌ esperado 1 | obtenido ' + prdSin);
  Logger.log(valCost === 45000 ? '✅ valor_costo OK'     : '❌ esperado $45.000 | obtenido $' + valCost);

  // BLOQUE 3 — Movimientos
  Logger.log('');
  Logger.log('── BLOQUE 3 MOVIMIENTOS PERÍODO ──');
  var entradas = Number(leer(17));
  var salidas  = Number(leer(18));
  var totalMov = Number(leer(19));
  var uVend    = Number(leer(20));
  var uCompra  = Number(leer(21));
  var rotacion = Number(leer(22));

  Logger.log('entradas_periodo:         ' + entradas + ' unidades');
  Logger.log('salidas_periodo:          ' + salidas  + ' unidades');
  Logger.log('total_movimientos:        ' + totalMov);
  Logger.log('unidades_vendidas:        ' + uVend);
  Logger.log('unidades_compradas:       ' + uCompra);
  Logger.log('ratio_rotacion:           ' + rotacion.toFixed(2));

  Logger.log(uVend === 15   ? '✅ unidades_vendidas OK'  : '❌ esperado 15 | obtenido ' + uVend);
  Logger.log(uCompra === 60 ? '✅ unidades_compradas OK' : '❌ esperado 60 | obtenido ' + uCompra);

  // BLOQUE 4 — Alertas
  Logger.log('');
  Logger.log('── BLOQUE 4 ALERTAS STOCK ──');
  var stockTotal = Number(leer(25));
  var pctCrit    = Number(leer(26));
  var pctSin     = Number(leer(27));
  var diasInv    = Number(leer(28));

  Logger.log('stock_total_unidades:     ' + stockTotal);
  Logger.log('pct_productos_criticos:   ' + (pctCrit*100).toFixed(1) + '%');
  Logger.log('pct_sin_stock:            ' + (pctSin*100).toFixed(1) + '%');
  Logger.log('dias_inventario:          ' + diasInv.toFixed(0));

  Logger.log(stockTotal === 45 ? '✅ stock_total OK' : '❌ esperado 45 | obtenido ' + stockTotal);

  // BLOQUE 5 — Hist precios
  Logger.log('');
  Logger.log('── BLOQUE 5 HISTORIAL PRECIOS ──');
  var regHP  = Number(leer(32));
  var prcProm = Number(leer(33));
  Logger.log('registros_hist_precios:   ' + regHP);
  Logger.log('precio_promedio_compra:   $' + prcProm.toLocaleString('es-CL'));
  Logger.log(regHP === 3 ? '✅ registros_hist OK' : '❌ esperado 3 | obtenido ' + regHP);

  // BLOQUE 6 — Validaciones
  Logger.log('');
  Logger.log('── BLOQUE 6 VALIDACIONES ──');
  var valsOk = true;
  [{f:37,l:'val_productos_activos'},
   {f:38,l:'val_valor_inventario'},
   {f:39,l:'val_stock_no_negativo'},
   {f:40,l:'val_criticos_ok'},
   {f:41,l:'val_rotacion_ok'}
  ].forEach(function(v) {
    var val = Number(leer(v.f));
    var ok  = val === 1;
    Logger.log((ok?'✅':'❌') + ' ' + v.l);
    if (!ok) valsOk = false;
  });

  var todoOk = prdAct===2 && prdStk===1 && valCost===45000 &&
               uVend===15 && stockTotal===45 && valsOk;

  Logger.log('');
  Logger.log(todoOk ?
    '✅ CALC_SUPPLY v1.0 — VALIDACIÓN APROBADA' :
    '❌ CALC_SUPPLY — REVISAR VALORES MARCADOS');
}
function validarCALC_GERENCIAL() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_GERENCIAL');
  Logger.log('=== VALIDACIÓN CALC_GERENCIAL v1.0 ===');
  Logger.log('');

  function leer(f) { return hoja.getRange(f,2).getValue(); }

  // BLOQUE 2 — Liquidez
  Logger.log('── BLOQUE 2 LIQUIDEZ ──');
  var caja    = Number(leer(8));
  var cobros  = Number(leer(9));
  var pagos   = Number(leer(10));
  var liq30   = Number(leer(11));
  var ratio   = Number(leer(12));
  var dias    = Number(leer(14));
  var semLiq  = Number(leer(15));
  Logger.log('caja_periodo:          $' + caja.toLocaleString('es-CL'));
  Logger.log('cobros_proyectados:    $' + cobros.toLocaleString('es-CL'));
  Logger.log('pagos_proyectados:     $' + pagos.toLocaleString('es-CL'));
  Logger.log('liquidez_proyectada:   $' + liq30.toLocaleString('es-CL'));
  Logger.log('ratio_liquidez:        '  + ratio.toFixed(2));
  Logger.log('dias_caja:             '  + dias.toFixed(0));
  Logger.log('semaforo_liquidez:     '  + semLiq + ' (' + (semLiq===3?'VERDE':semLiq===2?'AMARILLO':'ROJO') + ')');
  Logger.log(caja  === 1047400  ? '✅ caja OK'   : '❌ caja: '  + caja);
  Logger.log(liq30 === 1942400  ? '✅ liq30 OK'  : '❌ liq30: ' + liq30);
  Logger.log(ratio  > 1         ? '✅ ratio > 1' : '❌ ratio: ' + ratio);

  // BLOQUE 3 — Resultado
  Logger.log('');
  Logger.log('── BLOQUE 3 RESULTADO ──');
  var ingNeto = Number(leer(18));
  var flujo   = Number(leer(20));
  var margen  = Number(leer(22));
  var semRes  = Number(leer(23));
  Logger.log('ingresos_neto:         $' + ingNeto.toLocaleString('es-CL'));
  Logger.log('flujo_neto:            $' + flujo.toLocaleString('es-CL'));
  Logger.log('margen_operacional:    '  + (margen*100).toFixed(1) + '%');
  Logger.log('semaforo_resultado:    '  + semRes + ' (' + (semRes===3?'VERDE':semRes===2?'AMARILLO':'ROJO') + ')');
  Logger.log(ingNeto === 1395000 ? '✅ ingresos OK' : '❌ ingresos: ' + ingNeto);
  Logger.log(flujo   === 1047400 ? '✅ flujo OK'    : '❌ flujo: '    + flujo);

  // BLOQUE 4 — Ventas
  Logger.log('');
  Logger.log('── BLOQUE 4 VENTAS ──');
  var vMes   = Number(leer(26));
  var ticket = Number(leer(29));
  var docs   = Number(leer(30));
  Logger.log('ventas_mes_neto:       $' + vMes.toLocaleString('es-CL'));
  Logger.log('ticket_promedio:       $' + ticket.toLocaleString('es-CL'));
  Logger.log('documentos_mes:        '  + docs);
  Logger.log(vMes   === 46050 ? '✅ ventas OK'  : '❌ ventas: '  + vMes);
  Logger.log(ticket === 9210  ? '✅ ticket OK'  : '❌ ticket: '  + ticket);
  Logger.log(docs   === 5     ? '✅ docs OK'    : '❌ docs: '    + docs);

  // BLOQUE 5 — Cobranza
  Logger.log('');
  Logger.log('── BLOQUE 5 COBRANZA ──');
  var cxcTot = Number(leer(35));
  var cxcVen = Number(leer(36));
  var cxpTot = Number(leer(40));
  var semCob = Number(leer(42));
  Logger.log('cxc_total_pendiente:   $' + cxcTot.toLocaleString('es-CL'));
  Logger.log('cxc_vencida:           $' + cxcVen.toLocaleString('es-CL'));
  Logger.log('cxp_total_pendiente:   $' + cxpTot.toLocaleString('es-CL'));
  Logger.log('semaforo_cobranza:     '  + semCob + ' (' + (semCob===3?'VERDE':semCob===2?'AMARILLO':'ROJO') + ')');
  Logger.log(cxcTot === 1490000 ? '✅ cxc OK' : '❌ cxc: ' + cxcTot);
  Logger.log(cxcVen === 0       ? '✅ sin vencidas' : '⚠️  vencidas: $' + cxcVen);

  // BLOQUE 6 — Inventario
  Logger.log('');
  Logger.log('── BLOQUE 6 INVENTARIO ──');
  var prdAct  = Number(leer(45));
  var prdCrit = Number(leer(46));
  var valInv  = Number(leer(48));
  var semInv  = Number(leer(50));
  Logger.log('productos_activos:     '  + prdAct);
  Logger.log('productos_criticos:    '  + prdCrit);
  Logger.log('valor_inventario:      $' + valInv.toLocaleString('es-CL'));
  Logger.log('semaforo_inventario:   '  + semInv + ' (' + (semInv===3?'VERDE':semInv===2?'AMARILLO':'ROJO') + ')');
  Logger.log(prdAct === 2     ? '✅ productos OK'  : '❌ productos: '  + prdAct);
  Logger.log(valInv === 45000 ? '✅ inventario OK' : '❌ inventario: ' + valInv);

  // BLOQUE 7 — Alertas
  Logger.log('');
  Logger.log('── BLOQUE 7 ALERTAS ──');
  var alertas = Number(leer(53));
  Logger.log('alertas_criticas_total: ' + alertas);
  Logger.log(alertas >= 0 ? '✅ alertas OK' : '❌ alertas: ' + alertas);

  // BLOQUE 8 — Score
  Logger.log('');
  Logger.log('── BLOQUE 8 SCORE SALUD ──');
  var sLiq  = Number(leer(60));
  var sRes  = Number(leer(61));
  var sVta  = Number(leer(62));
  var sCob  = Number(leer(63));
  var sTotal= Number(leer(64));
  var estado= Number(leer(65));
  Logger.log('score_liquidez:        ' + sLiq  + '/40');
  Logger.log('score_resultado:       ' + sRes  + '/30');
  Logger.log('score_ventas:          ' + sVta  + '/20');
  Logger.log('score_cobranza:        ' + sCob  + '/10');
  Logger.log('score_total:           ' + sTotal + '/100');
  Logger.log('estado_general:        ' + estado + ' (' + (estado===3?'SALUDABLE':estado===2?'ATENCIÓN':'CRÍTICO') + ')');
  Logger.log(sTotal > 0 ? '✅ score calculado' : '❌ score en cero');

  // BLOQUE 9 — Acciones
  Logger.log('');
  Logger.log('── BLOQUE 9 ACCIONES PRIORITARIAS ──');
  var accTotal = Number(leer(74));
  Logger.log('total_acciones_urgentes: ' + accTotal);

  var todoOk = caja===1047400 && liq30===1942400 &&
               ingNeto===1395000 && vMes===46050 &&
               cxcTot===1490000 && prdAct===2 && sTotal>0;

  Logger.log('');
  Logger.log(todoOk ?
    '✅ CALC_GERENCIAL v1.0 — VALIDACIÓN APROBADA' :
    '❌ CALC_GERENCIAL — REVISAR VALORES MARCADOS');
}
function crearCALC_FINANCIERO() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_FINANCIERO');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('Hoja CALC_FINANCIERO no encontrada');
    return;
  }

  Logger.log('=== CONSTRUYENDO CALC_FINANCIERO v1.0 ===');
  hoja.clearContents();
  hoja.clearFormats();

  // Obtener período correctamente
  var cfg      = ss.getSheetByName('CONFIG_SISTEMA');
  var fechaT2  = cfg.getRange('T2').getValue();
  var empresa  = String(cfg.getRange('A2').getValue());
  var moneda   = String(cfg.getRange('K2').getValue());

  // Construir período YYYY-MM desde el objeto Date
  var periodo;
  if (fechaT2 instanceof Date) {
    var yyyy = fechaT2.getFullYear();
    var mm   = String(fechaT2.getMonth() + 1).padStart(2, '0');
    periodo  = yyyy + '-' + mm;
  } else {
    periodo = String(fechaT2).substring(0, 7);
  }

  Logger.log('Período detectado: [' + periodo + ']');
  Logger.log('Empresa detectada: [' + empresa + ']');

  // ── ESCRIBIR ESTRUCTURA COMPLETA FILA POR FILA ──
  // Esto garantiza que cada fórmula va exactamente donde corresponde

  var estructura = [
    // fila, col A etiqueta, col B valor/fórmula, es_formula
    [1,  '—CONTEXTO—',             'CALC_FINANCIERO v1.0',  false],
    [2,  'periodo_ref',            periodo,                  false],
    [3,  'id_empresa',             empresa,                  false],
    [4,  'moneda_base',            moneda,                   false],
    [5,  'fecha_calculo',          new Date(),               false],
    [6,  '',                       '',                       false],
    [7,  '—FLUJO_PERIODO—',        '',                       false],
    [8,  'ingresos_neto',          0, false],
    [9,  'ingresos_bruto',         '=SUMPRODUCT((DB_INGRESOS!D2:D5000="' + periodo + '")*(DB_INGRESOS!B2:B5000="' + empresa + '")*(DB_INGRESOS!J2:J5000))', true],
    [10, 'egresos_neto',           0, false],
    [11, 'egresos_bruto',          '=SUMPRODUCT((DB_EGRESOS!D2:D5000="' + periodo + '")*(DB_EGRESOS!B2:B5000="' + empresa + '")*(DB_EGRESOS!J2:J5000))',  true],
    [12, 'flujo_neto',             '=B8-B10',                true],
    [13, 'resultado_neto',         '=B9-B11',                true],
    [14, 'margen_operacional_pct', '=SUMPRODUCT((B8>0)*B12/MAX(B8;1))', true],
    [15, '',                       '',                       false],
    [16, '—CXC—',                  '',                       false],
    [17, 'cxc_total_pendiente',    '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!K2:K5000))', true],
    [18, 'cxc_vencida',            '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000<0)*(DB_CXC!K2:K5000))', true],
    [19, 'cxc_0_30_dias',          '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000>=0)*(DB_CXC!L2:L5000<=30)*(DB_CXC!K2:K5000))', true],
    [20, 'cxc_31_60_dias',         '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000>30)*(DB_CXC!L2:L5000<=60)*(DB_CXC!K2:K5000))', true],
    [21, 'cxc_mas_60_dias',        '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000>60)*(DB_CXC!K2:K5000))', true],
    [22, 'cxc_docs_pendientes',    '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!A2:A5000<>"")*1)', true],
    [23, 'cxc_docs_vencidos',      '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000<0)*(DB_CXC!A2:A5000<>"")*1)', true],
    [24, 'cxc_pct_vencida',        '=SUMPRODUCT((B17>0)*B18/MAX(B17;1))', true],
    [25, 'cxc_riesgo_pct',         '=SUMPRODUCT((B17>0)*(B18+B21)/MAX(B17;1))', true],
    [26, 'cxc_cobertura_egresos',  '=SUMPRODUCT((B10>0)*B17/MAX(B10;1))', true],
    [27, 'cxc_dias_promedio',      '=SUMPRODUCT((DB_CXC!B2:B5000="' + empresa + '")*(DB_CXC!N2:N5000<>"PAGADO")*(DB_CXC!L2:L5000)*(DB_CXC!K2:K5000))/MAX(B17;1)', true],
    [28, '',                       '',                       false],
    [29, '—CXP—',                  '',                       false],
    [30, 'cxp_total_pendiente',    '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!M2:M5000))', true],
    [31, 'cxp_vencida',            '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!N2:N5000<0)*(DB_CXP!M2:M5000))', true],
    [32, 'cxp_urgente_0_5dias',    '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!N2:N5000>=0)*(DB_CXP!N2:N5000<=5)*(DB_CXP!M2:M5000))', true],
    [33, 'cxp_6_30_dias',          '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!N2:N5000>5)*(DB_CXP!N2:N5000<=30)*(DB_CXP!M2:M5000))', true],
    [34, 'cxp_mas_30_dias',        '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!N2:N5000>30)*(DB_CXP!M2:M5000))', true],
    [35, 'cxp_docs_pendientes',    '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!A2:A5000<>"")*1)', true],
    [36, 'cxp_docs_vencidos',      '=SUMPRODUCT((DB_CXP!B2:B5000="' + empresa + '")*(DB_CXP!P2:P5000<>"PAGADO")*(DB_CXP!N2:N5000<0)*(DB_CXP!A2:A5000<>"")*1)', true],
    [37, 'cxp_pct_vencida',        '=SUMPRODUCT((B30>0)*B31/MAX(B30;1))', true],
    [38, 'cxp_presion_liquidez',   '=B31+B32+B33', true],
    [39, 'cxp_cobertura_flujo',    '=SUMPRODUCT((B30>0)*B12/MAX(B30;1))', true],
    [40, '',                       '',                       false],
    [41, '—LIQUIDEZ—',             '',                       false],
    [42, 'caja_periodo',           '=B12',                   true],
    [43, 'cobros_proyectados_30d', '=B19',                   true],
    [44, 'pagos_proyectados_30d',  '=B38',                   true],
    [45, 'liquidez_proyectada_30d','=B42+B43-B44',           true],
    [46, 'ratio_liquidez',         '=SUMPRODUCT((B44>0)*(B42+B43)/MAX(B44;1))+(B44=0)*1', true],
    [47, 'brecha_liquidez',        '=SUMPRODUCT((B45<0)*B45)', true],
    [48, 'dias_caja_disponible',   '=SUMPRODUCT((B10>0)*(B42/MAX(B10;1))*30)+(B10=0)*999', true],
    [49, '',                       '',                       false],
    [50, '—VALIDACIONES—',         '',                       false],
    [51, 'val_periodo_activo',     '=SUMPRODUCT((B2<>"")*1)', true],
    [52, 'val_ingresos_positivos', '=SUMPRODUCT((B8>=0)*1)',  true],
    [53, 'val_cxc_no_negativa',    '=SUMPRODUCT((B17>=0)*1)', true],
    [54, 'val_cxp_no_negativa',    '=SUMPRODUCT((B30>=0)*1)', true],
    [55, 'val_suma_cxc',           '=SUMPRODUCT((ABS(B17-(B18+B19+B20+B21))<=1)*1)', true],
    [56, 'val_suma_cxp',           '=SUMPRODUCT((ABS(B30-(B31+B32+B33+B34))<=1)*1)', true],
    [57, 'val_liquidez',           '=SUMPRODUCT((B45>=0)*1)', true]
  ];

  // Escribir fila por fila — garantiza posición exacta
  estructura.forEach(function(item) {
    var fila    = item[0];
    var etiq    = item[1];
    var valor   = item[2];
    var esForm  = item[3];

    if (etiq !== '') hoja.getRange(fila, 1).setValue(etiq);
    if (valor !== '') {
      if (esForm) {
        hoja.getRange(fila, 2).setFormula(valor);
      } else {
        hoja.getRange(fila, 2).setValue(valor);
      }
    }
  });

  SpreadsheetApp.flush();

  // ── FORMATOS ──
  [1, 7, 16, 29, 41, 50].forEach(function(f) {
    var r = hoja.getRange(f, 1, 1, 2);
    r.setBackground('#2C3E50');
    r.setFontColor('#FFFFFF');
    r.setFontWeight('bold');
  });

  // Porcentajes
  [14, 24, 25, 37].forEach(function(f) {
    hoja.getRange(f, 2).setNumberFormat('0.00%');
  });

  // Moneda
  [8,9,10,11,12,13,17,18,19,20,21,26,27,30,31,32,33,34,38,39,42,43,44,45,47].forEach(function(f) {
    hoja.getRange(f, 2).setNumberFormat('$#,##0');
  });

  // Fecha
  hoja.getRange(5, 2).setNumberFormat('DD/MM/YYYY');

  // Ancho columnas
  hoja.setColumnWidth(1, 220);
  hoja.setColumnWidth(2, 160);

  Logger.log('✅ CALC_FINANCIERO v1.0 construida — ' + estructura.length + ' filas');
  Logger.log('Período: ' + periodo + ' | Empresa: ' + empresa);
  Logger.log('Ejecuta validarCALC_FINANCIERO() para verificar');
}
function validarCALC_FINANCIERO() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_FINANCIERO');
  Logger.log('=== VALIDACIÓN CALC_FINANCIERO v1.0 ===');
  Logger.log('');

  function leer(f) { return hoja.getRange(f, 2).getValue(); }

  Logger.log('── BLOQUE 1 CONTEXTO ──');
  Logger.log('periodo_ref: [' + leer(2) + ']');
  Logger.log('id_empresa:  [' + leer(3) + ']');
  Logger.log(leer(3) !== '' ? '✅ Contexto OK' : '❌ Empresa vacía');

  Logger.log('');
  Logger.log('── BLOQUE 2 FLUJO ──');
  var ingNeto = Number(leer(8));
  var egrNeto = Number(leer(10));
  var flujo   = Number(leer(12));
  var margen  = Number(leer(14));
  Logger.log('ingresos_neto:        $' + ingNeto.toLocaleString('es-CL'));
  Logger.log('egresos_neto:         $' + egrNeto.toLocaleString('es-CL'));
  Logger.log('flujo_neto:           $' + flujo.toLocaleString('es-CL'));
  Logger.log('margen_operacional:   ' + (margen*100).toFixed(1) + '%');
  Logger.log(Math.abs(ingNeto-1395000)<1 ? '✅ ingresos OK' : '❌ ingresos: ' + ingNeto);
  Logger.log(Math.abs(egrNeto-347600)<1  ? '✅ egresos OK'  : '❌ egresos: '  + egrNeto);
  Logger.log(Math.abs(flujo-1047400)<1   ? '✅ flujo OK'    : '❌ flujo: '    + flujo);

  Logger.log('');
  Logger.log('── BLOQUE 3 CXC ──');
  var cxcTotal = Number(leer(17));
  var cxcVenc  = Number(leer(18));
  var cxc030   = Number(leer(19));
  Logger.log('cxc_total_pendiente:  $' + cxcTotal.toLocaleString('es-CL'));
  Logger.log('cxc_vencida:          $' + cxcVenc.toLocaleString('es-CL'));
  Logger.log('cxc_0_30_dias:        $' + cxc030.toLocaleString('es-CL'));
  Logger.log(Math.abs(cxcTotal-1490000)<1 ? '✅ cxc_total OK' : '❌ cxc_total: ' + cxcTotal);
  Logger.log(cxcVenc === 0               ? '✅ sin vencidas' : '⚠️  vencidas: ' + cxcVenc);

  Logger.log('');
  Logger.log('── BLOQUE 4 CXP ──');
  var cxpTotal = Number(leer(30));
  var cxpVenc  = Number(leer(31));
  var presion  = Number(leer(38));
  Logger.log('cxp_total_pendiente:  $' + cxpTotal.toLocaleString('es-CL'));
  Logger.log('cxp_vencida:          $' + cxpVenc.toLocaleString('es-CL'));
  Logger.log('cxp_presion_30d:      $' + presion.toLocaleString('es-CL'));
  Logger.log(Math.abs(cxpTotal-595000)<1 ? '✅ cxp_total OK' : '❌ cxp_total: ' + cxpTotal);
  Logger.log(cxpVenc === 0              ? '✅ sin vencidas' : '⚠️  vencidas: ' + cxpVenc);

  Logger.log('');
  Logger.log('── BLOQUE 5 LIQUIDEZ ──');
  var caja  = Number(leer(42));
  var liq30 = Number(leer(45));
  var ratio = Number(leer(46));
  var dias  = Number(leer(48));
  Logger.log('caja_periodo:          $' + caja.toLocaleString('es-CL'));
  Logger.log('liquidez_proyectada:   $' + liq30.toLocaleString('es-CL'));
  Logger.log('ratio_liquidez:        '  + ratio.toFixed(2));
  Logger.log('dias_caja:             '  + dias.toFixed(0));
  Logger.log(liq30 > 0  ? '✅ liquidez positiva' : '❌ liquidez negativa');
  Logger.log(ratio >= 1 ? '✅ ratio saludable'   : '⚠️  ratio < 1');

  Logger.log('');
  Logger.log('── BLOQUE 6 VALIDACIONES ──');
  var valsOk = true;
  [{f:51,l:'val_periodo_activo'},
   {f:52,l:'val_ingresos_positivos'},
   {f:53,l:'val_cxc_no_negativa'},
   {f:54,l:'val_cxp_no_negativa'},
   {f:55,l:'val_suma_cxc'},
   {f:56,l:'val_suma_cxp'},
   {f:57,l:'val_liquidez'}
  ].forEach(function(v) {
    var val = Number(leer(v.f));
    var ok  = val === 1;
    Logger.log((ok?'✅':'❌') + ' ' + v.l);
    if (!ok) valsOk = false;
  });

  var todoOk = Math.abs(ingNeto-1395000)<1 &&
               Math.abs(cxcTotal-1490000)<1 &&
               Math.abs(cxpTotal-595000)<1  &&
               liq30 > 0 && valsOk;

  Logger.log('');
  Logger.log(todoOk ?
    '✅ CALC_FINANCIERO v1.0 — VALIDACIÓN APROBADA' :
    '❌ CALC_FINANCIERO — REVISAR VALORES MARCADOS');
}
function verificarEstructuraArchivos() {
  Logger.log('=== VERIFICACIÓN ESTRUCTURA ARCHIVOS ===');
  Logger.log('');

  var funciones = [
    'generarID',
    'getConfig',
    'getPeriodoActual',
    'crearCALC_FINANCIERO',
    'validarCALC_FINANCIERO',
    'crearCALC_COMERCIAL',
    'validarCALC_SUPPLY',
    'crearCALC_SUPPLY',
    'crearCALC_GERENCIAL',
    'validarCALC_GERENCIAL',
    'limpiarFilasFantasmaSeguro',
    'verificarTodasLasCorrecciones',
    'auditarFuentesSupply',
    'auditarCALC_ALERTAS'
  ];

  var todoOk = true;
  funciones.forEach(function(nombre) {
    try {
      var existe = typeof eval(nombre) === 'function';
      Logger.log((existe ? '✅' : '❌') + ' ' + nombre);
      if (!existe) todoOk = false;
    } catch(e) {
      Logger.log('❌ ' + nombre + ' — no encontrada');
      todoOk = false;
    }
  });

  Logger.log('');
  Logger.log(todoOk ?
    '✅ Estructura correcta — listo para continuar' :
    '❌ Faltan funciones — revisar marcadas');
}
function crearCALC_GERENCIAL() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('CALC_GERENCIAL');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('Hoja CALC_GERENCIAL no encontrada');
    return;
  }

  Logger.log('=== CONSTRUYENDO CALC_GERENCIAL v1.0 ===');
  hoja.clearContents();
  hoja.clearFormats();

  var cfg       = ss.getSheetByName('CONFIG_SISTEMA');
  var fechaT2   = cfg.getRange('T2').getValue();
  var empresa   = String(cfg.getRange('A2').getValue());
  var nombreEmp = String(cfg.getRange('B2').getValue());

  var anio, mes;
  if (fechaT2 instanceof Date) {
    anio = fechaT2.getFullYear();
    mes  = fechaT2.getMonth() + 1;
  } else {
    var parts = String(fechaT2).substring(0,7).split('-');
    anio = Number(parts[0]);
    mes  = Number(parts[1]);
  }
  var periodo = anio + '-' + String(mes).padStart(2,'0');
  Logger.log('Período: ' + periodo + ' | Empresa: ' + empresa);

  var estructura = [
    [1,  '—CONTEXTO—',               'CALC_GERENCIAL v1.0', false],
    [2,  'periodo_ref',               periodo,               false],
    [3,  'id_empresa',                empresa,               false],
    [4,  'nombre_empresa',            nombreEmp,             false],
    [5,  'fecha_calculo',             new Date(),            false],
    [6,  '',                          '',                    false],
    [7,  '—LIQUIDEZ—',                '',                    false],
    [8,  'caja_periodo',              '=CALC_FINANCIERO!B42', true],
    [9,  'cobros_proyectados_30d',    '=CALC_FINANCIERO!B43', true],
    [10, 'pagos_proyectados_30d',     '=CALC_FINANCIERO!B44', true],
    [11, 'liquidez_proyectada_30d',   '=CALC_FINANCIERO!B45', true],
    [12, 'ratio_liquidez',            '=CALC_FINANCIERO!B46', true],
    [13, 'brecha_liquidez',           '=CALC_FINANCIERO!B47', true],
    [14, 'dias_caja_disponible',      '=CALC_FINANCIERO!B48', true],
    [15, 'semaforo_liquidez',
  '=SUMPRODUCT((CALC_FINANCIERO!B46>=1,5)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B46>=1)*(CALC_FINANCIERO!B46<1,5)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B46<1)*1)*1',
  true],
    [16, '',                          '',                    false],
    [17, '—RESULTADO_PERIODO—',       '',                    false],
    [18, 'ingresos_neto',             '=CALC_FINANCIERO!B8',  true],
    [19, 'egresos_neto',              '=CALC_FINANCIERO!B10', true],
    [20, 'flujo_neto',                '=CALC_FINANCIERO!B12', true],
    [21, 'resultado_neto',            '=CALC_FINANCIERO!B13', true],
    [22, 'margen_operacional_pct',    '=CALC_FINANCIERO!B14', true],
    [23, 'semaforo_resultado',
  '=SUMPRODUCT((CALC_FINANCIERO!B14>=0,2)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B14>=0)*(CALC_FINANCIERO!B14<0,2)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B14<0)*1)*1',
  true],
    [24, '',                          '',                    false],
    [25, '—VENTAS—',                  '',                    false],
    [26, 'ventas_mes_neto',           '=CALC_COMERCIAL!B7',  true],
    [27, 'ventas_mes_bruto',          '=CALC_COMERCIAL!B8',  true],
    [28, 'ventas_dia_neto',           '=CALC_COMERCIAL!B19', true],
    [29, 'ticket_promedio',           '=CALC_COMERCIAL!B14', true],
    [30, 'documentos_mes',            '=CALC_COMERCIAL!B13', true],
    [31, 'margen_bruto_comercial',    '=CALC_COMERCIAL!B10', true],
    [32, 'semaforo_ventas',
  '=SUMPRODUCT((CALC_COMERCIAL!B7>0)*1)*2+SUMPRODUCT((CALC_COMERCIAL!B11>=0,2)*1)',
  true],
    [33, '',                          '',                    false],
    [34, '—COBRANZA—',                '',                    false],
    [35, 'cxc_total_pendiente',       '=CALC_FINANCIERO!B17', true],
    [36, 'cxc_vencida',               '=CALC_FINANCIERO!B18', true],
    [37, 'cxc_0_30_dias',             '=CALC_FINANCIERO!B19', true],
    [38, 'cxc_docs_pendientes',       '=CALC_FINANCIERO!B22', true],
    [39, 'cxc_riesgo_pct',            '=CALC_FINANCIERO!B26', true],
    [40, 'cxp_total_pendiente',       '=CALC_FINANCIERO!B30', true],
    [41, 'cxp_presion_liquidez',      '=CALC_FINANCIERO!B38', true],
    [42, 'semaforo_cobranza',
  '=MIN(3;SUMPRODUCT((CALC_FINANCIERO!B18=0)*1)*3+SUMPRODUCT((CALC_FINANCIERO!B18>0)*(CALC_FINANCIERO!B26<0,1)*1)*2+SUMPRODUCT((CALC_FINANCIERO!B26>=0,1)*1)*1)',
  true],
    [43, '',                          '',                    false],
    [44, '—INVENTARIO—',              '',                    false],
    [45, 'productos_activos',         '=CALC_SUPPLY!B7',     true],
    [46, 'productos_criticos',        '=CALC_SUPPLY!B10',    true],
    [47, 'productos_sin_stock',       '=CALC_SUPPLY!B9',     true],
    [48, 'valor_inventario_costo',    '=CALC_SUPPLY!B12',    true],
    [49, 'dias_inventario_disponible','=CALC_SUPPLY!B28',    true],
    [50, 'semaforo_inventario',
      '=SUMPRODUCT((CALC_SUPPLY!B10=0)*1)*3+SUMPRODUCT((CALC_SUPPLY!B10>0)*(CALC_SUPPLY!B10<=2)*1)*2+SUMPRODUCT((CALC_SUPPLY!B10>2)*1)*1',
      true],
    [51, '',                          '',                    false],
    [52, '—ALERTAS—',                 '',                    false],
    [53, 'alertas_criticas_total',    '=CALC_ALERTAS!B2',    true],
    [54, 'alertas_stock',             '=CALC_ALERTAS!B3',    true],
    [55, 'alertas_cxc',               '=CALC_ALERTAS!B4',    true],
    [56, 'alertas_cxp',               '=CALC_ALERTAS!B5',    true],
    [57, 'alertas_batch',             '=CALC_ALERTAS!B8',    true],
    [58, '',                          '',                    false],
    [59, '—SCORE_SALUD—',             '',                    false],
    [60, 'score_liquidez',
      '=SUMPRODUCT((B15=3)*40)+SUMPRODUCT((B15=2)*24)+SUMPRODUCT((B15=1)*8)',
      true],
    [61, 'score_resultado',
      '=SUMPRODUCT((B23=3)*30)+SUMPRODUCT((B23=2)*18)+SUMPRODUCT((B23=1)*6)',
      true],
    [62, 'score_ventas',
      '=SUMPRODUCT((B32>=3)*20)+SUMPRODUCT((B32=2)*12)+SUMPRODUCT((B32<=1)*4)',
      true],
    [63, 'score_cobranza',
      '=SUMPRODUCT((B42=3)*10)+SUMPRODUCT((B42=2)*6)+SUMPRODUCT((B42=1)*2)',
      true],
    [64, 'score_total',               '=B60+B61+B62+B63',   true],
    [65, 'estado_general',
      '=SUMPRODUCT((B64>=80)*3)+SUMPRODUCT((B64>=60)*(B64<80)*2)+SUMPRODUCT((B64<60)*1)',
      true],
    [66, 'label_estado',              '=B64',                true],
    [67, 'alertas_activas',           '=B53',                true],
    [68, '',                          '',                    false],
    [69, '—ACCIONES_PRIORITARIAS—',   '',                    false],
    [70, 'accion_liquidez',           '=SUMPRODUCT((B15=1)*1)', true],
    [71, 'accion_cobranza',           '=SUMPRODUCT((B42=1)*1)', true],
    [72, 'accion_inventario',         '=SUMPRODUCT((B50=1)*1)', true],
    [73, 'accion_resultado',          '=SUMPRODUCT((B23=1)*1)', true],
    [74, 'total_acciones_urgentes',   '=B70+B71+B72+B73',   true],
    [75, '',                          '',                    false],
    [76, '—SEMAFOROS_EJECUTIVOS—',    '',                    false],
    [77, 'sem_exec_finanzas',         '=MIN(B15;B23)',        true],
    [78, 'sem_exec_comercial',        '=B32',                 true],
    [79, 'sem_exec_operaciones',      '=MIN(B42;B50)',        true]
  ];

  estructura.forEach(function(item) {
    var fila = item[0]; var etiq = item[1];
    var valor = item[2]; var esForm = item[3];
    if (etiq !== '') hoja.getRange(fila, 1).setValue(etiq);
    if (valor !== '') {
      if (esForm) hoja.getRange(fila, 2).setFormula(valor);
      else hoja.getRange(fila, 2).setValue(valor);
    }
  });

  SpreadsheetApp.flush();

  [1,7,17,25,34,44,52,59,69].forEach(function(f) {
    var r = hoja.getRange(f,1,1,2);
    r.setBackground('#2C3E50');
    r.setFontColor('#FFFFFF');
    r.setFontWeight('bold');
  });

  [8,9,10,11,13,18,19,20,21,26,27,28,29,31,35,36,37,40,41,48].forEach(function(f) {
    hoja.getRange(f,2).setNumberFormat('$#,##0');
  });

  [22,39].forEach(function(f) {
    hoja.getRange(f,2).setNumberFormat('0.00%');
  });

  [14,30,38,45,46,47,53,54,55,56,57,60,61,62,63,64,65,67,70,71,72,73,74].forEach(function(f) {
    hoja.getRange(f,2).setNumberFormat('#,##0');
  });

  [12,49].forEach(function(f) {
    hoja.getRange(f,2).setNumberFormat('#,##0.00');
  });

  hoja.getRange(5,2).setNumberFormat('DD/MM/YYYY');
  hoja.setColumnWidth(1,240);
  hoja.setColumnWidth(2,160);

  Logger.log('✅ CALC_GERENCIAL v1.0 construida');
  Logger.log('Ejecuta validarCALC_GERENCIAL() para verificar');
}
function verificarSemaforosEjecutivos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cg = ss.getSheetByName('CALC_GERENCIAL');
  Logger.log('=== SEMÁFOROS EJECUTIVOS ===');
  Logger.log('sem_exec_finanzas:    ' + cg.getRange('B77').getValue() + 
             ' — ' + cg.getRange('A77').getValue());
  Logger.log('sem_exec_comercial:   ' + cg.getRange('B78').getValue() + 
             ' — ' + cg.getRange('A78').getValue());
  Logger.log('sem_exec_operaciones: ' + cg.getRange('B79').getValue() + 
             ' — ' + cg.getRange('A79').getValue());
  Logger.log('Esperado: 3/3/3 con datos actuales');
}
function crearPlanCuentasBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('MSTR_PLAN_CUENTAS');

  if (!hoja) {
    SpreadsheetApp.getUi().alert('Hoja MSTR_PLAN_CUENTAS no encontrada');
    return;
  }

  Logger.log('=== CREANDO PLAN DE CUENTAS BASE VANTIS ===');

  var cfg = ss.getSheetByName('CONFIG_SISTEMA');
  var empresa = cfg.getRange('A2').getValue();
  var usuario = Session.getActiveUser().getEmail() || 'sistema';
  var fecha = new Date();

  // Verificar cuentas existentes para no duplicar
  var filActual = hoja.getLastRow();
  var existentes = {};
  if (filActual >= 2) {
    hoja.getRange(2,1,filActual-1,1).getValues().forEach(function(f) {
      if (f[0] !== '') existentes[String(f[0])] = true;
    });
  }

  // Estructura: [id_cuenta, nombre_cuenta, tipo_cuenta, id_padre, descripcion, codigo_sii, nivel, imputa_directo]
  var cuentas = [
    // ACTIVO
    ['1',     'ACTIVO',                    'ACTIVO',           '',     'Activo total',                          '', 'GRUPO',  false],
    ['1.1',   'Activo Circulante',         'ACTIVO',           '1',    'Activos de corto plazo',                '', 'CLASE',  false],
    ['1.1.1', 'Caja y Bancos',             'ACTIVO',           '1.1',  'Efectivo disponible',                   '', 'CUENTA', true],
    ['1.1.2', 'Clientes por Cobrar',       'ACTIVO',           '1.1',  'Cuentas por cobrar a clientes (CxC)',   '', 'CUENTA', true],
    ['1.1.3', 'IVA Crédito Fiscal',        'ACTIVO',           '1.1',  'IVA pagado en compras, recuperable',    '', 'CUENTA', true],
    ['1.1.4', 'Existencias',               'ACTIVO',           '1.1',  'Inventario de mercaderías',             '', 'CUENTA', true],
    ['1.2',   'Activo Fijo',               'ACTIVO',           '1',    'Activos de largo plazo',                '', 'CLASE',  false],
    ['1.2.1', 'Equipos y Muebles',         'ACTIVO',           '1.2',  'Activo fijo operacional',                '', 'CUENTA', true],

    // PASIVO
    ['2',     'PASIVO',                    'PASIVO',           '',     'Pasivo total',                          '', 'GRUPO',  false],
    ['2.1',   'Pasivo Circulante',         'PASIVO',           '2',    'Obligaciones de corto plazo',           '', 'CLASE',  false],
    ['2.1.1', 'Proveedores por Pagar',     'PASIVO',           '2.1',  'Cuentas por pagar a proveedores (CxP)', '', 'CUENTA', true],
    ['2.1.2', 'IVA Débito Fiscal',         'PASIVO',           '2.1',  'IVA cobrado en ventas, a declarar',     '', 'CUENTA', true],
    ['2.1.3', 'Retenciones por Pagar',     'PASIVO',           '2.1',  'Retenciones de honorarios pendientes',  '', 'CUENTA', true],

    // PATRIMONIO
    ['3',     'PATRIMONIO',                'PATRIMONIO',       '',     'Patrimonio total',                      '', 'GRUPO',  false],
    ['3.1',   'Capital',                   'PATRIMONIO',       '3',    'Capital de la empresa',                 '', 'CLASE',  false],
    ['3.1.1', 'Capital Social',            'PATRIMONIO',       '3.1',  'Aporte de capital de los socios',       '', 'CUENTA', true],

    // EGRESOS — expandir las existentes
    ['5.2.2', 'Arriendo',                  'EGRESOS',          '5.2',  'Arriendo de local u oficina',           '', 'CUENTA', true],
    ['5.2.3', 'Servicios Básicos',         'EGRESOS',          '5.2',  'Luz, agua, gas, internet',               '', 'CUENTA', true],
    ['5.2.4', 'Suministros y Materiales',  'EGRESOS',          '5.2',  'Insumos y materiales de oficina',       '', 'CUENTA', true],
    ['5.2.5', 'Honorarios',                'EGRESOS',          '5.2',  'Pago a prestadores de servicios (BH)',  '', 'CUENTA', true],
    ['5.2.6', 'Gastos Generales',          'EGRESOS',          '5.2',  'Gastos operacionales varios',           '', 'CUENTA', true]
  ];

  var filas = [];
  cuentas.forEach(function(c) {
    if (existentes[c[0]]) {
      Logger.log('Ya existe, se omite: ' + c[0] + ' - ' + c[1]);
      return;
    }
    filas.push([
      c[0],        // id_cuenta
      empresa,     // id_empresa
      c[1],        // nombre_cuenta
      c[2],        // tipo_cuenta
      c[3],        // id_cuenta_padre
      c[4],        // descripcion
      c[5],        // codigo_sii
      c[6],        // nivel
      c[7],        // imputa_directo
      true,        // activo
      fecha,       // creado_en
      usuario      // creado_por
    ]);
  });

  if (filas.length === 0) {
    Logger.log('No hay cuentas nuevas para agregar — todas ya existen');
    return;
  }

  var primeraFilaVacia = hoja.getLastRow() + 1;
  hoja.getRange(primeraFilaVacia, 1, filas.length, 12).setValues(filas);

  SpreadsheetApp.flush();
  Logger.log('✅ ' + filas.length + ' cuentas nuevas agregadas al Plan de Cuentas');
  Logger.log('Total cuentas en el sistema: ' + (existentes && Object.keys(existentes).length + filas.length));
}