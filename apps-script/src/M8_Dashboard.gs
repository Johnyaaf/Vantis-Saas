// ============================================================
// VANTIS ERP — M8_Dashboard.gs
// Servidor de aplicación HTML
// ============================================================

var VANTIS_SHEET_ID = '11dKJAoJxoJ0sjDxW4TiIVPgg8XvieGo4vnCVUh-wfv4';

function _getSS() {
  return SpreadsheetApp.openById(VANTIS_SHEET_ID);
}

function abrirVantis() {
  var url = 'https://script.google.com/macros/s/AKfycbxxA6k7uVG4QOOB5Ur46HYICH0mEDRv5b5YzQ5dZ8fHxe_q3aJ7xWTvNM4hGxbZlJ7ilw/exec';
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank");google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo VANTIS ERP...');
}

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('vantis_app')
    .setTitle('VANTIS ERP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getDatosContexto() {
  try {
    var ss  = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var cg  = ss.getSheetByName('CALC_GERENCIAL');

    if (!cg) throw new Error('Hoja CALC_GERENCIAL no encontrada');
    if (!cfg) throw new Error('Hoja CONFIG_SISTEMA no encontrada');

    var vPeriodo = cg.getRange('B2').getValue();
    var periodoStr = '';
    if (vPeriodo instanceof Date) {
      periodoStr = vPeriodo.getFullYear() + '-' + String(vPeriodo.getMonth()+1).padStart(2,'0');
    } else {
      periodoStr = String(vPeriodo || '—');
    }

    return {
      nombreEmpresa: String(cg.getRange('B4').getValue() || 'Sin nombre'),
      idEmpresa:     String(cg.getRange('B3').getValue() || 'EMP-01'),
      periodo:       periodoStr,
      fechaCalculo:  cg.getRange('B5').getDisplayValue() || '—',
      usuario:       'Administrador',
      version:       'v1.0'
    };
  } catch(e) {
    Logger.log('getDatosContexto ERROR: ' + e.message);
    return {
      nombreEmpresa: 'VANTIS ERP',
      idEmpresa:     '—',
      periodo:       '—',
      fechaCalculo:  '—',
      usuario:       'Administrador',
      version:       'v1.0'
    };
  }
}

function getDashboardData() {
  try {
    var ss      = _getSS();
    var cg      = ss.getSheetByName('CALC_GERENCIAL');
    var cf      = ss.getSheetByName('CALC_FINANCIERO');
    var cs      = ss.getSheetByName('CALC_SUPPLY');
    var ctx     = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hPrd    = ss.getSheetByName('MSTR_PRODUCTOS');
    var filPrd  = _ultimaFilaReal(hPrd);

    var skus = [];
    if (filPrd >= 2) {
      var pd = hPrd.getRange(2,1,filPrd-1,28).getValues();
      for (var i=0; i<pd.length; i++) {
        var f = pd[i];
        if (f[0]==='' || f[27]!==true || f[19]!==true) continue;
        var st = Number(f[11])||0;
        var mn = Number(f[13])||0;
        if (st<=mn || st===0) skus.push({nombre:String(f[3]||f[2]).substring(0,22), dias:st===0?0:999});
      }
    }

    var meta = 0;
    var hM = ss.getSheetByName('DB_METAS');
    if (hM && hM.getLastRow()>1) {
      var md = hM.getRange(2,1,hM.getLastRow()-1,10).getValues();
      for (var j=0; j<md.length; j++) {
        if (md[j][0]!=='' && String(md[j][3])===empresa && String(md[j][4])==='VENTAS') {
          meta = Number(md[j][5])||0; break;
        }
      }
    }

    var resultado = {};
    resultado.periodo            = ctx.periodo;
    resultado.fechaCalculo       = ctx.fechaCalculo;
    resultado.scoreTotal         = Number(cg.getRange('B64').getValue());
    resultado.estadoGeneral      = Number(cg.getRange('B65').getValue());
    resultado.semLiquidez        = Number(cg.getRange('B15').getValue());
    resultado.semResultado       = Number(cg.getRange('B23').getValue());
    resultado.semVentas          = Number(cg.getRange('B32').getValue());
    resultado.semCobranza        = Number(cg.getRange('B42').getValue());
    resultado.semInventario      = Number(cg.getRange('B50').getValue());
    resultado.semExecFinanzas    = Number(cg.getRange('B77').getValue());
    resultado.semExecComercial   = Number(cg.getRange('B78').getValue());
    resultado.semExecOperaciones = Number(cg.getRange('B79').getValue());
    resultado.flujoNeto          = Number(cg.getRange('B20').getValue());
    resultado.liquidez30d        = Number(cg.getRange('B11').getValue());
    resultado.ratioLiquidez      = Number(cg.getRange('B12').getValue());
    resultado.diasCaja           = Number(cg.getRange('B14').getValue());
    resultado.ingresosNeto       = Number(cg.getRange('B18').getValue());
    resultado.egresosNeto        = Number(cg.getRange('B19').getValue());
    resultado.margenPct          = Number(cg.getRange('B22').getValue());
    resultado.ventasMes          = Number(cg.getRange('B26').getValue());
    resultado.ticketPromedio     = Number(cg.getRange('B29').getValue());
    resultado.documentosMes      = Number(cg.getRange('B30').getValue());
    resultado.metaVentas         = meta;
    resultado.cxcTotal           = Number(cg.getRange('B35').getValue());
    resultado.cxcVencida         = Number(cg.getRange('B36').getValue());
    resultado.cxc030             = Number(cf.getRange('B19').getValue());
    resultado.cxc3160            = Number(cf.getRange('B20').getValue());
    resultado.cxcMas60           = Number(cf.getRange('B21').getValue());
    resultado.cxpTotal           = Number(cg.getRange('B40').getValue());
    resultado.cxpPresion         = Number(cg.getRange('B41').getValue());
    resultado.stockTotal         = Number(cs.getRange('B25').getValue());
    resultado.alertas            = Number(cg.getRange('B53').getValue());
    resultado.accionesUrgentes   = Number(cg.getRange('B74').getValue());
    resultado.accLiquidez        = Number(cg.getRange('B70').getValue());
    resultado.accCobranza        = Number(cg.getRange('B71').getValue());
    resultado.accInventario      = Number(cg.getRange('B72').getValue());
    resultado.accResultado       = Number(cg.getRange('B73').getValue());
    resultado.skusQuiebre        = skus.slice(0,3);

    var hVentas  = ss.getSheetByName('DB_VENTAS');
    var filVentas = _ultimaFilaReal(hVentas);
    var periodos  = {};

    if (filVentas >= 2) {
      var vDatos = hVentas.getRange(2,1,filVentas-1,28).getValues();
      vDatos.forEach(function(f) {
        if (f[0] === '') return;
        if (String(f[27]) === 'ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var anio = fecha.getFullYear();
        var mes  = String(fecha.getMonth()+1).padStart(2,'0');
        var per  = anio + '-' + mes;
        if (!periodos[per]) periodos[per] = 0;
        periodos[per] += Number(f[19]) || 0;
      });
    }

    var claves = Object.keys(periodos).sort();
    var ultimos12 = claves.slice(-12);

    resultado.chartVentas = {
      labels:  ultimos12,
      valores: ultimos12.map(function(p){ return periodos[p] || 0; })
    };

    return resultado;

  } catch(e) {
    Logger.log('getDashboardData ERROR: ' + e.message);
    return null;
  }
}

function diagnosticarM8() {
  Logger.log('=== DIAGNÓSTICO M8_Dashboard ===');
  try {
    var ss = _getSS();
    Logger.log('SS OK: ' + ss.getName());
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    Logger.log('CONFIG_SISTEMA: ' + (cfg ? 'OK' : 'NO ENCONTRADA'));
    var cg = ss.getSheetByName('CALC_GERENCIAL');
    Logger.log('CALC_GERENCIAL: ' + (cg ? 'OK filas=' + cg.getLastRow() : 'NO ENCONTRADA'));
    Logger.log('_ultimaFilaReal disponible: ' + (typeof _ultimaFilaReal === 'function'));
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
  }
}

function testDatosContexto() {
  var resultado = getDatosContexto();
  Logger.log('=== TEST getDatosContexto ===');
  Logger.log('resultado es null: ' + (resultado === null));
}

function testDashboardData() {
  var resultado = getDashboardData();
  Logger.log('=== TEST getDashboardData ===');
  Logger.log('resultado es null: ' + (resultado === null));
}

function auditarFuentesFinanzas() {
  var ss = _getSS();
  Logger.log('=== AUDITORÍA FUENTES MÓDULO FINANZAS ===');
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  [[8,'ingresos_neto'],[10,'egresos_neto'],[12,'flujo_neto']].forEach(function(k) {
    Logger.log('  B'+k[0]+' '+k[1]+': '+cf.getRange(k[0],2).getValue());
  });
}

function getFinanzasData() {
  try {
    var ss  = _getSS();
    var cf  = ss.getSheetByName('CALC_FINANCIERO');
    var cg  = ss.getSheetByName('CALC_GERENCIAL');
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var anulados = _getDocumentosAnulados();

    var hI   = ss.getSheetByName('DB_INGRESOS');
    var hE   = ss.getSheetByName('DB_EGRESOS');
    var hM   = ss.getSheetByName('DB_METAS');
    var filI = _ultimaFilaReal(hI);
    var filE = _ultimaFilaReal(hE);

    var semLiq = Number(cg.getRange('B15').getValue());
    var semRes = Number(cg.getRange('B23').getValue());
    var semCob = Number(cg.getRange('B42').getValue());
    var riesgo = Math.min(semLiq, semRes, semCob);

    var periodoActual = ctx.periodo;
    var parts = periodoActual.split('-');
    var y = Number(parts[0]);
    var m = Number(parts[1]);
    m--; if (m===0) { m=12; y--; }
    var periodoAnterior = y+'-'+String(m).padStart(2,'0');

    // ── UNA SOLA PASADA SOBRE DB_INGRESOS ──
    var periodosI = {};
    var tablaI = {};
    if (filI >= 2) {
      hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        if (anulados.ventas[String(f[12]||'')]) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        periodosI[per] = (periodosI[per]||0) + Number(f[7]);
        if (per === periodoActual) {
          var key = String(f[4])+'|'+String(f[6]);
          tablaI[key] = (tablaI[key]||0) + Number(f[7]);
        }
      });
    }

    // ── UNA SOLA PASADA SOBRE DB_EGRESOS ──
    var periodosE = {};
    var cuentasE = {};
    var cuentasEMA = {};
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='') return;
        if (anulados.compras[String(f[15]||'')]) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        periodosE[per] = (periodosE[per]||0) + Number(f[7]);
        var key  = String(f[4])+'|'+String(f[6]);
        if (per === periodoActual)   cuentasE[key]   = (cuentasE[key]||0)   + Number(f[7]);
        if (per === periodoAnterior) cuentasEMA[key] = (cuentasEMA[key]||0) + Number(f[7]);
      });
    }

    var todosPerI = Object.keys(periodosI).sort();
    var todosPerE = Object.keys(periodosE).sort();
    var todosPer  = todosPerI.concat(todosPerE)
      .filter(function(v,i,a){ return a.indexOf(v)===i; }).sort();
    var ultimos12 = todosPer.slice(-12);
    var evolucion = {
      labels:   ultimos12,
      ingresos: ultimos12.map(function(p){ return periodosI[p]||0; }),
      egresos:  ultimos12.map(function(p){ return periodosE[p]||0; })
    };

    var totalE = Number(cf.getRange('B10').getValue()) || 1;
    var top5Items = Object.keys(cuentasE).map(function(k) {
      var p   = k.split('|');
      var mn  = cuentasE[k];
      var mMA = cuentasEMA[k] || 0;
      return {
        cuenta:    p[0],
        desc:      p[1],
        monto:     mn,
        pct:       (mn/totalE) || 0,
        variacion: mMA > 0 ? (mn-mMA)/mMA : null
      };
    }).sort(function(a,b){ return b.monto-a.monto; });

    var top5     = top5Items.slice(0,5);
    var sumTop5  = top5.reduce(function(s,x){ return s+x.monto; }, 0);
    var otrosMonto = (Number(cf.getRange('B10').getValue())||0) - sumTop5;
    if (otrosMonto > 0) {
      top5.push({ cuenta:'—', desc:'Otros', monto: otrosMonto,
                  pct: otrosMonto/totalE, variacion: null });
    }

    var totalI = Number(cf.getRange('B8').getValue()) || 1;
    var detalleIngresos = Object.keys(tablaI).map(function(k) {
      var p = k.split('|');
      return { cuenta: p[0], desc: p[1], monto: tablaI[k], pct: (tablaI[k]/totalI) || 0 };
    }).sort(function(a,b){ return b.monto-a.monto; });

    var metaVentas = 0;
    if (hM && _ultimaFilaReal(hM) >= 2) {
      hM.getRange(2,1,_ultimaFilaReal(hM)-1,10).getValues().forEach(function(f) {
        if (f[0]!=='' && String(f[3])===empresa && String(f[4])==='VENTAS') {
          metaVentas = Number(f[5])||0;
        }
      });
    }

    var cajaPeriodo     = Number(cf.getRange('B42').getValue());
    var cxc030          = Number(cf.getRange('B19').getValue());
    var cxpPresion      = Number(cf.getRange('B38').getValue());
    var egresosNeto     = Number(cf.getRange('B10').getValue());
    var flujoProyectado = cajaPeriodo + cxc030 - cxpPresion;
    var diasCobertura   = egresosNeto > 0
      ? Math.round(flujoProyectado / (egresosNeto/30))
      : 999;

    var resultado = {};
    resultado.periodo          = ctx.periodo;
    resultado.fechaCalculo     = ctx.fechaCalculo;
    resultado.cajaPeriodo      = cajaPeriodo;
    resultado.cajaReal         = _calcularCajaReal(empresa, ctx, anulados);
    resultado.ingresosNeto     = Number(cf.getRange('B8').getValue());
    resultado.egresosNeto      = egresosNeto;
    resultado.flujoNeto        = Number(cf.getRange('B12').getValue());
    resultado.margenPct        = Number(cf.getRange('B14').getValue());
    resultado.liquidez30d      = Number(cf.getRange('B45').getValue());
    resultado.ratioLiquidez    = Number(cf.getRange('B46').getValue());
    resultado.diasCaja         = Number(cf.getRange('B48').getValue());
    resultado.riesgo           = riesgo;
    resultado.cxcTotal         = Number(cf.getRange('B17').getValue());
    resultado.cxcVencida       = Number(cf.getRange('B18').getValue());
    resultado.cxc030           = cxc030;
    resultado.cxc3160          = Number(cf.getRange('B20').getValue());
    resultado.cxcMas60         = Number(cf.getRange('B21').getValue());
    resultado.cxpTotal         = Number(cf.getRange('B30').getValue());
    resultado.cxpPresion       = cxpPresion;
    resultado.flujoProyectado  = flujoProyectado;
    resultado.diasCobertura    = diasCobertura;
    resultado.metaVentas       = metaVentas;
    resultado.metaPct          = metaVentas > 0 ? resultado.ingresosNeto/metaVentas : 0;
    resultado.evolucion        = evolucion;
    resultado.top5Egresos      = top5;
    resultado.detalleIngresos  = detalleIngresos;
    return resultado;

  } catch(e) {
    Logger.log('getFinanzasData ERROR: ' + e.stack);
    return { ok:false, errorReal: e.message, stack: e.stack };
  }
}

function getComercialData() {
  try {
    var ss  = _getSS();
    var cc  = ss.getSheetByName('CALC_COMERCIAL');
    var ctx = getDatosContexto();
    var empresa  = ctx.idEmpresa;
    var periodoActual = ctx.periodo;

    var hV  = ss.getSheetByName('DB_VENTAS');
    var hC  = ss.getSheetByName('MSTR_CLIENTES');
    var hP  = ss.getSheetByName('MSTR_PRODUCTOS');
    var hM  = ss.getSheetByName('DB_METAS');
    var filV = _ultimaFilaReal(hV);
    var filC = _ultimaFilaReal(hC);
    var filP = _ultimaFilaReal(hP);

    var mapClientes = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,10).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var nombre = String(f[5]||'').trim();
        var rut    = String(f[4]||'').trim();
        var esNombre = nombre !== '' && isNaN(Number(nombre)) && nombre.length > 2;
        mapClientes[String(f[0])] = esNombre ? { nombre: nombre, rut: rut } : null;
      });
    }

    var mapProductos = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,10).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapProductos[String(f[0])] = String(f[3]||f[2]||f[0]);
      });
    }

    var metaVentas = 0;
    if (hM && _ultimaFilaReal(hM) >= 2) {
      hM.getRange(2,1,_ultimaFilaReal(hM)-1,10).getValues().forEach(function(f) {
        if (f[0]!=='' && String(f[3])===empresa && String(f[4])==='VENTAS') {
          metaVentas = Number(f[5])||0;
        }
      });
    }

    var periodos     = {};
    var porTipo      = {};
    var porProducto  = {};
    var porCliente   = {};

    var partsP = periodoActual.split('-');
    var yAct   = Number(partsP[0]);
    var mAct   = Number(partsP[1]);

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var fecha  = f[10];
        if (!(fecha instanceof Date)) return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;
        var emp = String(f[2]||'');
        if (emp !== empresa) return;

        var anio = fecha.getFullYear();
        var mes  = fecha.getMonth()+1;
        var per  = anio+'-'+String(mes).padStart(2,'0');
        var neto = Number(f[19])||0;
        var cant = Number(f[16])||0;
        var mrg  = Number(f[23])||0;
        var tipo = String(f[3]||'SIN TIPO');
        var idCli = String(f[12]||'');
        var idPrd = String(f[13]||'');

        if (!periodos[per]) periodos[per] = 0;
        periodos[per] += neto;

        if (anio === yAct && mes === mAct) {
          if (!porTipo[tipo]) porTipo[tipo] = { monto:0, docs:0 };
          porTipo[tipo].monto += neto;
          porTipo[tipo].docs  += 1;

          if (!porProducto[idPrd]) porProducto[idPrd] = { monto:0, cant:0, margen:0 };
          porProducto[idPrd].monto  += neto;
          porProducto[idPrd].cant   += cant;
          porProducto[idPrd].margen += mrg;

          if (idCli !== 'CLI-GENERICO') {
            if (!porCliente[idCli]) porCliente[idCli] = { monto:0, docs:0 };
            porCliente[idCli].monto += neto;
            porCliente[idCli].docs  += 1;
          }
        }
      });
    }

    var clavesPer = Object.keys(periodos).sort();
    var ult12     = clavesPer.slice(-12);
    var evolucion = {
      labels:  ult12,
      ventas:  ult12.map(function(p){ return periodos[p]||0; }),
      meta:    metaVentas
    };

    evolucion.tendencia = evolucion.ventas.map(function(v,i,arr) {
      if (i < 2) return null;
      return Math.round((arr[i]+arr[i-1]+arr[i-2])/3);
    });

    var topProductos = Object.keys(porProducto).map(function(id) {
      var d   = porProducto[id];
      var pct = d.monto > 0 ? d.margen/d.monto : 0;
      var nombre = mapProductos[id] || id;
      return { id:id, nombre:nombre, cant:d.cant, monto:d.monto, margenPct:pct };
    }).sort(function(a,b){ return b.margenPct - a.margenPct; }).slice(0,5);

    var topClientes = Object.keys(porCliente).map(function(id) {
      var d   = porCliente[id];
      var obj = mapClientes[id];
      var label, sublabel, sinReg;
      if (obj) {
        label    = obj.nombre;
        sublabel = obj.rut ? obj.rut : '';
        sinReg   = false;
      } else {
        label    = id.substring(0,14)+'...';
        sublabel = 'Sin registrar';
        sinReg   = true;
      }
      return { id:id, label:label, sublabel:sublabel, sinReg:sinReg, monto:d.monto, docs:d.docs };
    }).sort(function(a,b){ return b.monto - a.monto; }).slice(0,5);

    var tiposDoc = Object.keys(porTipo).map(function(t) {
      var d = porTipo[t];
      return { tipo:t, monto:d.monto, docs:d.docs, ticket: d.docs>0 ? Math.round(d.monto/d.docs) : 0 };
    }).sort(function(a,b){ return b.monto - a.monto; });

    var resultado = {};
    resultado.periodo        = ctx.periodo;
    resultado.fechaCalculo   = ctx.fechaCalculo;
    resultado.ventasNeto     = Number(cc.getRange('B7').getValue());
    resultado.ticketPromedio = Number(cc.getRange('B14').getValue());
    resultado.margenBrutoPct = Number(cc.getRange('B11').getValue());
    resultado.documentosMes  = Number(cc.getRange('B13').getValue());
    resultado.ventasHoy      = Number(cc.getRange('B19').getValue());
    resultado.docsHoy        = Number(cc.getRange('B21').getValue());
    resultado.metaVentas     = metaVentas;
    resultado.metaPct        = metaVentas > 0 ? resultado.ventasNeto/metaVentas : 0;
    resultado.evolucion      = evolucion;
    resultado.topProductos   = topProductos;
    resultado.topClientes    = topClientes;
    resultado.tiposDoc       = tiposDoc;
    return resultado;

  } catch(e) {
    Logger.log('getComercialData ERROR: ' + e.message);
    return null;
  }
}

function getInventarioData() {
  try {
    var ss  = _getSS();
    var cs  = ss.getSheetByName('CALC_SUPPLY');
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var hP  = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);

    var hM  = ss.getSheetByName('DB_MOVIMIENTOS');
    var filM = _ultimaFilaReal(hM);

    var hV  = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);

    var ultimaVentaPorProducto = {};
    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var idProd = String(f[13]);
        if (!ultimaVentaPorProducto[idProd] || fecha > ultimaVentaPorProducto[idProd]) {
          ultimaVentaPorProducto[idProd] = fecha;
        }
      });
    }
    var hoy = new Date();

    var productos = [];
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='' || f[27]!==true) return;
        if (String(f[1]) !== empresa) return;
        var id     = String(f[0]);
        var stock  = Number(f[11])||0;
        var minimo = Number(f[13])||0;
        var critico= Number(f[15])||0;
        var costo  = Number(f[7])||0;
        var precio = Number(f[8])||0;
        var estado = 3;
        if (critico > 0 && stock <= critico) estado = 1;
        else if (minimo > 0 && stock < minimo) estado = 2;
        else if (stock === 0) estado = 1;

        var ultimaVenta = ultimaVentaPorProducto[id];
        var diasSinVenta = ultimaVenta
          ? Math.floor((hoy - ultimaVenta) / (1000*60*60*24))
          : null;

        var sugerido = Math.max(minimo - stock, 0);

        productos.push({
          id: id,
          sku: String(f[2]),
          nombre: String(f[3]),
          stock: stock,
          minimo: minimo,
          critico: critico,
          costo: costo,
          precio: precio,
          valor: stock*costo,
          estado: estado,
          diasSinVenta: diasSinVenta,
          sugerido: sugerido
        });
      });
    }
    productos.sort(function(a,b){ return b.valor - a.valor; });

    var periodosEnt = {};
    var periodosSal = {};
    if (filM >= 2) {
      hM.getRange(2,1,filM-1,12).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var tipo = String(f[7]||'');
        if (tipo === 'APERTURA') return;
        var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        var impacto = Number(f[9])||0;
        if (impacto > 0) periodosEnt[per] = (periodosEnt[per]||0) + impacto;
        if (impacto < 0) periodosSal[per] = (periodosSal[per]||0) + Math.abs(impacto);
      });
    }
    var todosPer = Object.keys(periodosEnt).concat(Object.keys(periodosSal))
      .filter(function(v,i,a){ return a.indexOf(v)===i; }).sort();
    var ult12 = todosPer.slice(-12);
    var evolucion = {
      labels:   ult12,
      entradas: ult12.map(function(p){ return periodosEnt[p]||0; }),
      salidas:  ult12.map(function(p){ return periodosSal[p]||0; })
    };

    var resultado = {};
    resultado.periodo          = ctx.periodo;
    resultado.fechaCalculo     = ctx.fechaCalculo;
    resultado.productosActivos = Number(cs.getRange('B7').getValue());
    resultado.valorCosto       = Number(cs.getRange('B12').getValue());
    resultado.valorVenta       = Number(cs.getRange('B13').getValue());
    resultado.criticos         = Number(cs.getRange('B10').getValue());
    resultado.sinStock         = Number(cs.getRange('B9').getValue());
    resultado.stockTotal       = Number(cs.getRange('B25').getValue());
    resultado.entradasPeriodo  = Number(cs.getRange('B17').getValue());
    resultado.salidasPeriodo   = Number(cs.getRange('B18').getValue());
    resultado.diasRotacion     = resultado.salidasPeriodo > 0
      ? Math.round(resultado.stockTotal / (resultado.salidasPeriodo/30))
      : 999;
    resultado.productos        = productos;
    resultado.evolucion        = evolucion;
    return resultado;

  } catch(e) {
    Logger.log('getInventarioData ERROR: ' + e.message);
    return null;
  }
}

function getCobranzaData() {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var mapa = _getMapaSaldosCxC();

    var hC   = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var mapClientes = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var nombre = String(f[5]||'').trim();
        var rut    = String(f[4]||'').trim();
        var esNombre = nombre !== '' && isNaN(Number(nombre)) && nombre.length > 2;
        mapClientes[String(f[0])] = { nombre: esNombre ? nombre : null, rut: rut };
      });
    }

    var hoy = new Date();
    var documentos = [];
    var porCliente = {};
    var sumaPonderada = 0;
    var sumaSaldos = 0;
    var totalCxC = 0, cxcVencida = 0, cxc030 = 0, cxc3160 = 0, cxcMas60 = 0;
    var docsPendientes = 0, docsVencidos = 0;

    Object.keys(mapa).forEach(function(k) {
      var m = mapa[k];
      if (m.idEmpresa !== empresa) return;
      if (m.saldoVivo <= 0) return;

      var cliObj = mapClientes[m.idCliente];
      var nombreCli = cliObj && cliObj.nombre ? cliObj.nombre : m.idCliente.substring(0,14)+'...';
      var sinReg = !(cliObj && cliObj.nombre);

      documentos.push({
        cliente: nombreCli, sinReg: sinReg, doc: m.numeroDoc,
        emision: m.emision instanceof Date ? m.emision.getTime() : null,
        vence: m.vence instanceof Date ? m.vence.getTime() : null,
        dias: m.dias, saldo: m.saldoVivo, vencido: m.dias < 0
      });

      totalCxC += m.saldoVivo;
      docsPendientes++;
      if (m.dias < 0) { cxcVencida += m.saldoVivo; docsVencidos++; }
      else if (m.dias <= 30) cxc030 += m.saldoVivo;
      else if (m.dias <= 60) cxc3160 += m.saldoVivo;
      else cxcMas60 += m.saldoVivo;

      if (m.idCliente !== 'CLI-GENERICO') {
        if (!porCliente[m.idCliente]) porCliente[m.idCliente] = { saldo:0, docs:0 };
        porCliente[m.idCliente].saldo += m.saldoVivo;
        porCliente[m.idCliente].docs  += 1;
      }
      if (m.vence instanceof Date) {
        var diasVenc = Math.floor((m.vence - hoy)/(1000*60*60*24));
        sumaPonderada += diasVenc * m.saldoVivo;
        sumaSaldos    += m.saldoVivo;
      }
    });

    documentos.sort(function(a,b){ return b.saldo - a.saldo; });

    var topDeudores = Object.keys(porCliente).map(function(id) {
      var d = porCliente[id];
      var cliObj = mapClientes[id];
      var nombre = cliObj && cliObj.nombre ? cliObj.nombre : id.substring(0,14)+'...';
      var sinReg = !(cliObj && cliObj.nombre);
      var pctCartera = totalCxC > 0 ? d.saldo/totalCxC : 0;
      return { nombre:nombre, sinReg:sinReg, saldo:d.saldo, docs:d.docs, pctCartera:pctCartera };
    }).sort(function(a,b){ return b.saldo - a.saldo; }).slice(0,5);

    var dsoPonderado = sumaSaldos > 0 ? Math.round(sumaPonderada/sumaSaldos) : 0;
    var plazoPromedio = 30;
    var clsDSO = dsoPonderado <= plazoPromedio+5 ? 3 : dsoPonderado <= plazoPromedio+15 ? 2 : 1;
    var concentracionMax = topDeudores.length ? topDeudores[0].pctCartera : 0;

    var flujoEsperado30d = 0;
    documentos.forEach(function(doc) {
      if (doc.vence) {
        var diasParaVencer = Math.floor((doc.vence - hoy.getTime())/(1000*60*60*24));
        if (diasParaVencer >= 0 && diasParaVencer <= 30) flujoEsperado30d += doc.saldo;
      }
    });

    var resultado = {};
    resultado.flujoEsperado30d = flujoEsperado30d;
    resultado.periodo        = ctx.periodo;
    resultado.fechaCalculo   = ctx.fechaCalculo;
    resultado.cxcTotal       = totalCxC;
    resultado.cxcVencida     = cxcVencida;
    resultado.cxc030         = cxc030;
    resultado.cxc3160        = cxc3160;
    resultado.cxcMas60       = cxcMas60;
    resultado.docsPendientes = docsPendientes;
    resultado.docsVencidos   = docsVencidos;
    resultado.dso            = dsoPonderado;
    resultado.dsoSemaforo    = clsDSO;
    resultado.plazoPromedio  = plazoPromedio;
    resultado.topDeudores    = topDeudores;
    resultado.concentracionMax = concentracionMax;
    resultado.documentos     = documentos;
    return resultado;

  } catch(e) {
    Logger.log('getCobranzaData ERROR: ' + e.message);
    return null;
  }
}

function getProveedoresData() {
  try {
    var ss  = _getSS();
    var cf  = ss.getSheetByName('CALC_FINANCIERO');
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var mapa = _getMapaSaldosCxP();

    var hP   = ss.getSheetByName('MSTR_PROVEEDORES');
    var filP = _ultimaFilaReal(hP);
    var mapProveedores = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var nombre = String(f[5]||'').trim();
        var rut    = String(f[4]||'').trim();
        var esNombre = nombre !== '' && isNaN(Number(nombre)) && nombre.length > 2;
        mapProveedores[String(f[0])] = { nombre: esNombre ? nombre : null, rut: rut };
      });
    }

    var hoy = new Date();
    var documentos = [];
    var porProveedor = {};
    var sumaPonderada = 0;
    var sumaSaldos = 0;
    var totalCxP=0, cxpVencida=0, cxpUrgente=0, cxp630=0, cxpMas30=0;
    var docsPendientes=0, docsVencidos=0;

    Object.keys(mapa).forEach(function(k) {
      var m = mapa[k];
      if (m.idEmpresa !== empresa) return;
      if (m.saldoVivo <= 0) return;

      var provObj  = mapProveedores[m.idProveedor];
      var nombreProv = provObj && provObj.nombre ? provObj.nombre : 'Proveedor sin registrar';
      var rutProv = provObj && provObj.nombre ? provObj.rut : m.idProveedor;
      var sinReg = !(provObj && provObj.nombre);
      var urgente = m.dias >= 0 && m.dias <= 5;
      var vencido = m.dias < 0;

      documentos.push({
        proveedor: nombreProv, rutProveedor: rutProv, sinReg: sinReg, doc: m.numeroDoc,
        emision: m.emision instanceof Date ? m.emision.getTime() : null,
        vence: m.vence instanceof Date ? m.vence.getTime() : null,
        dias: m.dias, saldo: m.saldoVivo, vencido: vencido, urgente: urgente
      });

      totalCxP += m.saldoVivo;
      docsPendientes++;
      if (vencido) { cxpVencida += m.saldoVivo; docsVencidos++; }
      if (urgente || vencido) cxpUrgente += m.saldoVivo;
      else if (m.dias <= 30) cxp630 += m.saldoVivo;
      else cxpMas30 += m.saldoVivo;

      if (!porProveedor[m.idProveedor]) porProveedor[m.idProveedor] = { saldo:0, docs:0 };
      porProveedor[m.idProveedor].saldo += m.saldoVivo;
      porProveedor[m.idProveedor].docs  += 1;

      if (m.vence instanceof Date) {
        var diasParaPagar = Math.floor((m.vence - hoy)/(1000*60*60*24));
        sumaPonderada += diasParaPagar * m.saldoVivo;
        sumaSaldos    += m.saldoVivo;
      }
    });

    documentos.sort(function(a,b){ return a.dias - b.dias; });

    var topProveedores = Object.keys(porProveedor).map(function(id) {
      var d = porProveedor[id];
      var provObj = mapProveedores[id];
      var nombre = provObj && provObj.nombre ? provObj.nombre : 'Proveedor sin registrar';
      var rut = provObj && provObj.nombre ? provObj.rut : id;
      var sinReg = !(provObj && provObj.nombre);
      var pctCartera = totalCxP > 0 ? d.saldo/totalCxP : 0;
      return { id:id, nombre:nombre, rut:rut, sinReg:sinReg, saldo:d.saldo, docs:d.docs, pctCartera:pctCartera };
    }).sort(function(a,b){ return b.saldo - a.saldo; }).slice(0,5);

    var dpoPonderado = sumaSaldos > 0 ? Math.round(sumaPonderada/sumaSaldos) : 0;
    var plazoPromedio = 30;
    var dpoSaludable = dpoPonderado <= plazoPromedio + 10;
    var concentracionMax = topProveedores.length ? topProveedores[0].pctCartera : 0;
    var presion5dias = 0;
    documentos.forEach(function(doc) { if (doc.urgente || doc.vencido) presion5dias += doc.saldo; });

    var cajaActual = Number(cf.getRange('B42').getValue());
    var faltante = presion5dias - cajaActual;

    var resultado = {};
    resultado.periodo        = ctx.periodo;
    resultado.fechaCalculo   = ctx.fechaCalculo;
    resultado.cxpTotal       = totalCxP;
    resultado.cxpVencida     = cxpVencida;
    resultado.cxpUrgente     = cxpUrgente;
    resultado.cxp630         = cxp630;
    resultado.cxpMas30       = cxpMas30;
    resultado.docsPendientes = docsPendientes;
    resultado.docsVencidos   = docsVencidos;
    resultado.dpo            = dpoPonderado;
    resultado.dpoSaludable   = dpoSaludable;
    resultado.plazoPromedio  = plazoPromedio;
    resultado.topProveedores = topProveedores;
    resultado.concentracionMax = concentracionMax;
    resultado.documentos     = documentos;
    resultado.presion5dias   = presion5dias;
    resultado.cajaActual     = cajaActual;
    resultado.faltanteCaja   = faltante;
    return resultado;

  } catch(e) {
    Logger.log('getProveedoresData ERROR: ' + e.message);
    return null;
  }
}

function getAlertasData() {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hoy = new Date();

    var alertas = [];

    var hP   = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var ultimaVentaPorProducto = {};

    var hV   = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var idProd = String(f[13]);
        if (!ultimaVentaPorProducto[idProd] || fecha > ultimaVentaPorProducto[idProd]) {
          ultimaVentaPorProducto[idProd] = fecha;
        }

        var margenPct = Number(f[24])||0;
        if (margenPct < 0) {
          alertas.push({
            nivel: 1, tiempo:'HOY', tipo:'COMERCIAL', modulo:'comercial',
            texto: 'Venta '+String(f[4])+' con margen negativo ('+(margenPct*100).toFixed(1)+'%)',
            sugerido: 'Revisar precio del producto o costo de venta',
            impactoCaja: 0
          });
        }
      });
    }

    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='' || f[27]!==true) return;
        if (String(f[1]) !== empresa) return;
        var idProd = String(f[0]);
        var nombre  = String(f[3]);
        var stock   = Number(f[11])||0;
        var minimo  = Number(f[13])||0;
        var critico = Number(f[15])||0;

        if (stock === 0) {
          alertas.push({
            nivel: 1, tiempo:'HOY', tipo:'STOCK', modulo:'inventario',
            texto: nombre + ' está sin stock',
            sugerido: 'Comprar ' + Math.max(minimo,1) + ' unidades',
            impactoCaja: 0
          });
        } else if (critico > 0 && stock <= critico) {
          alertas.push({
            nivel: 1, tiempo:'HOY', tipo:'STOCK', modulo:'inventario',
            texto: nombre + ' bajo nivel crítico (' + stock + ' unid.)',
            sugerido: 'Comprar ' + Math.max(minimo-stock,1) + ' unidades',
            impactoCaja: 0
          });
        } else if (minimo > 0 && stock < minimo) {
          alertas.push({
            nivel: 2, tiempo:'SEMANA', tipo:'STOCK', modulo:'inventario',
            texto: nombre + ' bajo stock mínimo (' + stock + ' unid.)',
            sugerido: 'Comprar ' + (minimo-stock) + ' unidades',
            impactoCaja: 0
          });
        }

        var ultimaVenta = ultimaVentaPorProducto[idProd];
        var diasSinVenta = ultimaVenta ? Math.floor((hoy - ultimaVenta)/(1000*60*60*24)) : null;
        if (stock > 0 && diasSinVenta !== null && diasSinVenta >= 90) {
          alertas.push({
            nivel: 3, tiempo:'PLANIFICAR', tipo:'STOCK', modulo:'inventario',
            texto: nombre + ' sin ventas hace ' + diasSinVenta + ' días — capital inmovilizado',
            sugerido: 'Evaluar liquidación o promoción del producto',
            impactoCaja: 0
          });
        }
      });
    }

    var mapaCxC = _getMapaSaldosCxC();
    var cxcVencidaTotal = 0;
    var cxcMas60Total = 0;
    var porClienteTmp = {};
    var ultimaVentaPorCliente = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var idCli = String(f[12]);
        if (!ultimaVentaPorCliente[idCli] || fecha > ultimaVentaPorCliente[idCli]) {
          ultimaVentaPorCliente[idCli] = fecha;
        }
      });
    }

    Object.keys(mapaCxC).forEach(function(k) {
      var m = mapaCxC[k];
      if (m.idEmpresa !== empresa || m.saldoVivo <= 0) return;
      if (m.dias < 0) cxcVencidaTotal += m.saldoVivo;
      if (m.dias < -60) cxcMas60Total += m.saldoVivo;
      porClienteTmp[m.idCliente] = (porClienteTmp[m.idCliente]||0) + m.saldoVivo;
    });

    if (cxcVencidaTotal > 0) {
      alertas.push({
        nivel: cxcVencidaTotal > 500000 ? 1 : 2,
        tiempo: cxcVencidaTotal > 500000 ? 'HOY' : 'SEMANA',
        tipo: 'COBRANZA', modulo: 'cobranza',
        texto: 'Cartera vencida por ' + fmtMoneyServer(cxcVencidaTotal),
        sugerido: 'Contactar clientes con documentos vencidos',
        impactoCaja: cxcVencidaTotal
      });
    }

    if (cxcMas60Total > 0) {
      alertas.push({
        nivel: 1, tiempo:'HOY', tipo:'COBRANZA', modulo:'cobranza',
        texto: 'Cartera con mora superior a 60 días: ' + fmtMoneyServer(cxcMas60Total),
        sugerido: 'Evaluar inicio de cobranza prejudicial',
        impactoCaja: cxcMas60Total
      });
    }

    var maxClienteSaldo = 0;
    Object.keys(porClienteTmp).forEach(function(id) {
      if (porClienteTmp[id] > maxClienteSaldo) maxClienteSaldo = porClienteTmp[id];
    });
    var cf = ss.getSheetByName('CALC_FINANCIERO');
    var cxcTotal = Number(cf.getRange('B17').getValue());
    if (cxcTotal > 0 && (maxClienteSaldo/cxcTotal) > 0.5) {
      alertas.push({
        nivel: 2, tiempo:'SEMANA', tipo:'COBRANZA', modulo:'cobranza',
        texto: 'Un cliente concentra ' + Math.round((maxClienteSaldo/cxcTotal)*100) + '% de la cartera por cobrar',
        sugerido: 'Diversificar cartera de clientes a futuro',
        impactoCaja: 0
      });
    }

    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var mapClientesNombre = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        var nombre = String(f[5]||'').trim();
        var esNombre = nombre !== '' && isNaN(Number(nombre)) && nombre.length > 2;
        mapClientesNombre[String(f[0])] = esNombre ? nombre : null;
      });
    }
    Object.keys(ultimaVentaPorCliente).forEach(function(idCli) {
      var diasInactivo = Math.floor((hoy - ultimaVentaPorCliente[idCli])/(1000*60*60*24));
      if (diasInactivo >= 60) {
        var nombreCli = mapClientesNombre[idCli] || 'Cliente sin registrar';
        alertas.push({
          nivel: 2, tiempo:'SEMANA', tipo:'COMERCIAL', modulo:'comercial',
          texto: nombreCli + ' inactivo hace ' + diasInactivo + ' días',
          sugerido: 'Contactar para reactivar la relación comercial',
          impactoCaja: 0
        });
      }
    });

    var mapaCxP = _getMapaSaldosCxP();
    var cxpUrgenteTotal = 0;
    var porProveedorTmp = {};
    Object.keys(mapaCxP).forEach(function(k) {
      var m = mapaCxP[k];
      if (m.idEmpresa !== empresa || m.saldoVivo <= 0) return;
      if ((m.dias >= 0 && m.dias <= 5) || m.dias < 0) cxpUrgenteTotal += m.saldoVivo;
      porProveedorTmp[m.idProveedor] = (porProveedorTmp[m.idProveedor]||0) + m.saldoVivo;
    });
    if (cxpUrgenteTotal > 0) {
      alertas.push({
        nivel: 1, tiempo:'HOY', tipo:'PAGOS', modulo:'proveedores',
        texto: 'Pagos urgentes por ' + fmtMoneyServer(cxpUrgenteTotal) + ' en próximos 5 días',
        sugerido: 'Verificar caja disponible antes de la fecha de pago',
        impactoCaja: -cxpUrgenteTotal
      });
    }

    var cxpTotal = Number(cf.getRange('B30').getValue());
    var maxProvSaldo = 0;
    Object.keys(porProveedorTmp).forEach(function(id) {
      if (porProveedorTmp[id] > maxProvSaldo) maxProvSaldo = porProveedorTmp[id];
    });
    if (cxpTotal > 0 && (maxProvSaldo/cxpTotal) > 0.5) {
      alertas.push({
        nivel: 2, tiempo:'SEMANA', tipo:'PAGOS', modulo:'proveedores',
        texto: 'Un proveedor concentra ' + Math.round((maxProvSaldo/cxpTotal)*100) + '% de las cuentas por pagar',
        sugerido: 'Buscar proveedor alternativo para reducir dependencia',
        impactoCaja: 0
      });
    }

    alertas.sort(function(a,b) {
      if (a.nivel !== b.nivel) return a.nivel - b.nivel;
      return Math.abs(b.impactoCaja) - Math.abs(a.impactoCaja);
    });

    var hBatch = ss.getSheetByName('BATCH_RESULTADOS');
    var ultimoBatch = null;
    var batchFallido = false;
    if (hBatch) {
      var filB = _ultimaFilaReal(hBatch);
      if (filB >= 2) {
        var ultimaFila = hBatch.getRange(filB,1,1,10).getValues()[0];
        var fechaBatch = ultimaFila[2];
        var estadoBatch = String(ultimaFila[5]||'');
        ultimoBatch = fechaBatch instanceof Date ? fechaBatch.getTime() : null;
        batchFallido = estadoBatch !== 'EXITOSO';
      }
    }

    var resultado = {};
    resultado.alertas = alertas;
    resultado.totalCriticas = alertas.filter(function(a){ return a.nivel===1; }).length;
    resultado.totalStock = alertas.filter(function(a){ return a.tipo==='STOCK'; }).length;
    resultado.totalCobranza = alertas.filter(function(a){ return a.tipo==='COBRANZA'; }).length;
    resultado.totalPagos = alertas.filter(function(a){ return a.tipo==='PAGOS'; }).length;
    resultado.ultimoBatch = ultimoBatch;
    resultado.batchFallido = batchFallido;
    return resultado;

  } catch(e) {
    Logger.log('getAlertasData ERROR: ' + e.message);
    return null;
  }
}

function fmtMoneyServer(n) {
  return '$' + Math.round(n).toLocaleString('es-CL');
}

function _limpiarNaN(obj) {
  var json = JSON.stringify(obj, function(key, val) {
    return (typeof val === 'number' && !isFinite(val)) ? null : val;
  });
  return JSON.parse(json);
}

function getFichaCliente(idCliente) {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var hC   = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var datosCliente = null;

    if (filC >= 2) {
      var rangoC = hC.getRange(2,1,filC-1,23).getValues();
      for (var i=0; i<rangoC.length; i++) {
        if (String(rangoC[i][0]) === idCliente) {
          datosCliente = rangoC[i];
          break;
        }
      }
    }

    if (!datosCliente) return null;

    var ficha = {
      id: String(datosCliente[0]),
      razonSocial: String(datosCliente[5]||''),
      nombreFantasia: String(datosCliente[6]||''),
      rut: String(datosCliente[4]||''),
      giro: String(datosCliente[7]||''),
      email: String(datosCliente[8]||''),
      telefono: String(datosCliente[9]||''),
      direccion: String(datosCliente[10]||''),
      comuna: String(datosCliente[11]||''),
      condicionPago: String(datosCliente[14]||''),
      limiteCredito: Number(datosCliente[15])||0,
      activo: datosCliente[17] === true
    };

    var mapaCxC = _getMapaSaldosCxC();
    var documentos = [];

    Object.keys(mapaCxC).forEach(function(k) {
      var m = mapaCxC[k];
      if (m.idEmpresa !== empresa) return;
      if (m.idCliente !== idCliente) return;
      if (m.saldoVivo <= 0) return;

      documentos.push({
        idCxc: m.idCxc,
        doc: m.numeroDoc,
        tipoDoc: m.tipoDoc,
        emision: m.emision instanceof Date ? m.emision.getTime() : null,
        vence: m.vence instanceof Date ? m.vence.getTime() : null,
        monto: m.montoOriginal,
        pagado: m.sumaCobros,
        saldo: m.montoOriginal,
        dias: m.dias,
        estado: m.estado,
        abono: m.sumaCobros + m.sumaNC,
        saldoFinal: m.saldoVivo
      });
    });

    var saldoTotal = documentos.reduce(function(s,d){ return s + Math.max(d.saldoFinal,0); }, 0);

    var hV   = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var historialCompras = [];
    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[12])!==idCliente) return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;
        historialCompras.push({
          fecha: f[10] instanceof Date ? f[10].getTime() : null,
          doc: String(f[4]),
          tipoDoc: String(f[3]),
          monto: Number(f[19])||0
        });
      });
    }
    historialCompras.sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    var historialPagos = [];
    var docPorIdCxc = {};
    Object.keys(mapaCxC).forEach(function(k) {
      var m = mapaCxC[k];
      docPorIdCxc[m.idCxc] = m.numeroDoc;
    });
    var hCobHist = ss.getSheetByName('DB_COBROS_CXC');
    var filCobHist = _ultimaFilaReal(hCobHist);
    if (filCobHist >= 2) {
      hCobHist.getRange(2,1,filCobHist-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[4])!==idCliente) return;
        historialPagos.push({
          idCobro: String(f[0]),
          fecha: f[5] instanceof Date ? f[5].getTime() : null,
          doc: docPorIdCxc[String(f[3])] || String(f[3]),
          monto: Number(f[7])||0,
          anulado: f[12]===true,
          motivoAnulacion: String(f[13]||''),
          anuladoPor: String(f[14]||''),
          fechaAnulacion: f[15] instanceof Date ? f[15].getTime() : null
        });
      });
    }
    historialPagos.sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    var resultado = {};
    resultado.ficha = ficha;
    resultado.documentos = documentos;
    resultado.saldoTotal = saldoTotal;
    resultado.historialCompras = historialCompras;
    resultado.historialPagos = historialPagos;
    return resultado;

  } catch(e) {
    Logger.log('getFichaCliente ERROR: ' + e.message);
    return null;
  }
}

function getFichaProveedor(idProveedor) {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var hP   = ss.getSheetByName('MSTR_PROVEEDORES');
    var filP = _ultimaFilaReal(hP);
    var datosProveedor = null;

    if (filP >= 2) {
      var rangoP = hP.getRange(2,1,filP-1,24).getValues();
      for (var i=0; i<rangoP.length; i++) {
        if (String(rangoP[i][0]) === idProveedor) {
          datosProveedor = rangoP[i];
          break;
        }
      }
    }

    if (!datosProveedor) return null;

    var ficha = {
      id: String(datosProveedor[0]),
      razonSocial: String(datosProveedor[5]||''),
      nombreFantasia: String(datosProveedor[6]||''),
      rut: String(datosProveedor[4]||''),
      giro: String(datosProveedor[7]||''),
      email: String(datosProveedor[8]||''),
      telefono: String(datosProveedor[9]||''),
      condicionPago: String(datosProveedor[10]||''),
      plazoEntrega: Number(datosProveedor[11])||0,
      activo: datosProveedor[18] === true
    };

    var mapaCxP = _getMapaSaldosCxP();
    var documentos = [];

    Object.keys(mapaCxP).forEach(function(k) {
      var m = mapaCxP[k];
      if (m.idEmpresa !== empresa) return;
      if (m.idProveedor !== idProveedor) return;
      if (m.saldoVivo <= 0) return;

      documentos.push({
        idCxp: m.idCxp,
        doc: m.numeroDoc,
        tipoDoc: m.tipoDoc,
        emision: m.emision instanceof Date ? m.emision.getTime() : null,
        vence: m.vence instanceof Date ? m.vence.getTime() : null,
        monto: m.montoOriginal,
        pagado: m.sumaPagos,
        saldo: m.montoOriginal,
        dias: m.dias,
        estado: m.estado,
        abono: m.sumaPagos,
        saldoFinal: m.saldoVivo
      });
    });

    var saldoTotal = documentos.reduce(function(s,d){ return s + Math.max(d.saldoFinal,0); }, 0);

    var hC   = ss.getSheetByName('DB_COMPRAS');
    var filCom = _ultimaFilaReal(hC);
    var historialCompras = [];
    if (filCom >= 2) {
      hC.getRange(2,1,filCom-1,30).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[3])!==idProveedor) return;
        var estado = String(f[22]||'');
        if (estado === 'ANULADO') return;
        historialCompras.push({
          fecha: f[6] instanceof Date ? f[6].getTime() : null,
          doc: String(f[1]),
          monto: Number(f[17])||0
        });
      });
    }
    historialCompras.sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    var historialPagos = [];
    var docPorIdCxp = {};
    Object.keys(mapaCxP).forEach(function(k) {
      var m = mapaCxP[k];
      docPorIdCxp[m.idCxp] = m.numeroDoc;
    });
    var hPagHist = ss.getSheetByName('DB_PAGOS_CXP');
    var filPagHist = _ultimaFilaReal(hPagHist);
    if (filPagHist >= 2) {
      hPagHist.getRange(2,1,filPagHist-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[4])!==idProveedor) return;
        historialPagos.push({
          idPago: String(f[0]),
          fecha: f[5] instanceof Date ? f[5].getTime() : null,
          doc: docPorIdCxp[String(f[3])] || String(f[3]),
          monto: Number(f[7])||0,
          anulado: f[12]===true,
          motivoAnulacion: String(f[13]||''),
          anuladoPor: String(f[14]||''),
          fechaAnulacion: f[15] instanceof Date ? f[15].getTime() : null
        });
      });
    }
    historialPagos.sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    var resultado = {};
    resultado.ficha = ficha;
    resultado.documentos = documentos;
    resultado.saldoTotal = saldoTotal;
    resultado.historialCompras = historialCompras;
    resultado.historialPagos = historialPagos;
    return resultado;

  } catch(e) {
    Logger.log('getFichaProveedor ERROR: ' + e.message);
    return null;
  }
}

function buscarClientes(termino) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);

    if (!termino || termino.trim() === '') return [];

    var terminoNorm = _normalizarTexto(termino).replace(/[.\-]/g,'');
    var resultados = [];

    if (filC >= 2) {
      hC.getRange(2,1,filC-1,23).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        if (f[17] !== true) return;
        var nombre = String(f[5]||'');
        var rut    = String(f[4]||'');
        var nombreNorm = _normalizarTexto(nombre);
        var rutNorm    = _normalizarTexto(rut).replace(/[.\-]/g,'');
        if (nombreNorm.indexOf(terminoNorm) !== -1 || rutNorm.indexOf(terminoNorm) !== -1) {
          resultados.push({ id:String(f[0]), nombre:nombre, rut:rut, condicionPago:String(f[14]||'') });
        }
      });
    }
    return resultados.slice(0,10);
  } catch(e) {
    Logger.log('buscarClientes ERROR: ' + e.message);
    return [];
  }
}

function buscarProductos(termino) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var tasaIva = Number(ss.getSheetByName('CONFIG_SISTEMA').getRange('L2').getValue()) || 0.19;
    var cfgCom = getConfigComercial();
    var mostrarCosto = !cfgCom || cfgCom.mostrarCostoProductos;

    if (!termino || termino.trim() === '') return [];

    var terminoNorm = _normalizarTexto(termino);
    var resultados = [];

    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        if (f[27] !== true) return;
        var sku    = String(f[2]||'');
        var nombre = String(f[3]||'');
        var skuNorm    = _normalizarTexto(sku);
        var nombreNorm = _normalizarTexto(nombre);
        if (nombreNorm.indexOf(terminoNorm) !== -1 || skuNorm.indexOf(terminoNorm) !== -1) {
          var precioNeto = Number(f[8])||0;
          resultados.push({
            id: String(f[0]),
            sku: sku,
            nombre: nombre,
            unidad: String(f[6]||'UN'),
            precioNeto: precioNeto,
            precioConIva: Math.round(precioNeto * (1+tasaIva)),
            costo: mostrarCosto ? (Number(f[7])||0) : null,
            stock: Number(f[11])||0
          });
        }
      });
    }
    return resultados.slice(0,10);
  } catch(e) {
    Logger.log('buscarProductos ERROR: ' + e.message);
    return [];
  }
}

function _normalizarTexto(texto) {
  return String(texto)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function procesarNuevaVenta(datos) {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var cfgCom = getConfigComercial();

    if (!datos.cliente || !datos.cliente.id) {
      return { ok:false, error:'Debes seleccionar un cliente' };
    }
    var esNDSimple = datos.tipoDoc === 'NOTA_DEBITO' && datos.esCargoSimple;

    if (!esNDSimple) {
      if (!datos.lineas || !datos.lineas.length) {
        return { ok:false, error:'Debes agregar al menos una línea de producto' };
      }
      for (var i=0; i<datos.lineas.length; i++) {
        if (!datos.lineas[i].productoId) {
          return { ok:false, error:'Hay una línea sin producto seleccionado' };
        }
        if (!datos.lineas[i].cantidad || datos.lineas[i].cantidad <= 0) {
          return { ok:false, error:'Hay una línea con cantidad inválida' };
        }
        var descLinea = Number(datos.lineas[i].descuentoPct)||0;
        if (cfgCom && descLinea > cfgCom.descuentoMaxPct && !cfgCom.autorizarDescuentoAlto) {
          return { ok:false, error:'El descuento de '+descLinea+'% supera el máximo permitido ('+cfgCom.descuentoMaxPct+'%). Ajusta el descuento o solicita autorización en Configuración.' };
        }
      }
    } else {
      if (!datos.glosaCargo || String(datos.glosaCargo).trim() === '') {
        return { ok:false, error:'Debes ingresar una descripción del cargo' };
      }
      if (!datos.montoCargo || datos.montoCargo <= 0) {
        return { ok:false, error:'El monto debe ser mayor a cero' };
      }
    }
    var checkPeriodo = _validarPeriodoNoBloqueado(datos.fecha);
    if (!checkPeriodo.ok) return { ok:false, error: checkPeriodo.error };

    if (datos.tipoDoc === 'NOTA_CREDITO' && !datos.docOrigen) {
      return { ok:false, error:'Debes seleccionar la factura origen de la Nota de Crédito' };
    }
    if (datos.cliente.id === 'CLI-GENERICO') {
      if (datos.condicionPago !== 'EFECTIVO' && datos.condicionPago !== 'TARJETA') {
        return { ok:false, error:'El Cliente Genérico solo acepta pago Efectivo o Tarjeta' };
      }
    }

    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var mapaProductos = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
        if (f[0]==='') return;
        mapaProductos[String(f[0])] = { fila: idx+2, stock: Number(f[11])||0, costo: Number(f[7])||0, nombre:String(f[3]) };
      });
    }
    if (!esNDSimple) {
      for (var j=0; j<datos.lineas.length; j++) {
        var prod = mapaProductos[datos.lineas[j].productoId];
        if (!prod) return { ok:false, error:'Producto no encontrado: ' + datos.lineas[j].productoId };
        var bloqueaSinStock = !cfgCom || cfgCom.bloquearVentaSinStock;
        if (bloqueaSinStock && datos.lineas[j].cantidad > prod.stock) {
          return { ok:false, error:'Stock insuficiente para "' + prod.nombre + '". Disponible: ' + prod.stock };
        }
      }
    }

    var fechaVenta = new Date(datos.fecha);
    var tasaIva = Number(ss.getSheetByName('CONFIG_SISTEMA').getRange('L2').getValue()) || 0.19;
    var idDocumento = 'DOC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS');
    var numeroDoc = (datos.tipoDoc === 'BOLETA' ? 'BOL-' : datos.tipoDoc === 'FACTURA' ? 'FAC-' :
                      datos.tipoDoc === 'NOTA_CREDITO' ? 'NC-' : 'ND-') +
                     Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

    var esEsContado = (datos.condicionPago === 'CONTADO' || datos.condicionPago === 'EFECTIVO' || datos.condicionPago === 'TARJETA');

    var hV = ss.getSheetByName('DB_VENTAS');
    var filasVenta = [];
    var totalNetoDoc = 0;
    var totalIvaDoc = 0;
    var totalBrutoDoc = 0;

    if (esNDSimple) {
      var subtotalNetoND = Math.round(Number(datos.montoCargo));
      var subtotalIvaND = Math.round(subtotalNetoND * tasaIva);
      var subtotalBrutoND = subtotalNetoND + subtotalIvaND;
      totalNetoDoc = subtotalNetoND;
      totalIvaDoc = subtotalIvaND;
      totalBrutoDoc = subtotalBrutoND;

      var filaND = new Array(34).fill('');
      filaND[0]  = generarID('VTA');
      filaND[1]  = idDocumento;
      filaND[2]  = empresa;
      filaND[3]  = datos.tipoDoc;
      filaND[4]  = numeroDoc;
      filaND[6]  = datos.glosaCargo;
      filaND[10] = fechaVenta;
      filaND[11] = ctx.periodo;
      filaND[12] = datos.cliente.id;
      if (datos.tipoDoc === 'NOTA_CREDITO') filaND[9] = datos.docOrigen;
      filaND[19] = subtotalNetoND;
      filaND[20] = subtotalIvaND;
      filaND[21] = subtotalBrutoND;
      filaND[25] = 'CLP';
      filaND[26] = 1;
      filaND[27] = 'VIGENTE';
      filaND[30] = new Date();
      filaND[31] = usuario;
      filasVenta.push(filaND);
    } else {
    datos.lineas.forEach(function(linea) {
      var prod = mapaProductos[linea.productoId];
      var subtotalNeto = Math.round(linea.cantidad * linea.precioNeto * (1 - (linea.descuentoPct||0)/100));
      var subtotalIva = Math.round(subtotalNeto * tasaIva);
      var subtotalBruto = subtotalNeto + subtotalIva;
      var costoTotal = linea.cantidad * prod.costo;
      var margenLinea = subtotalNeto - costoTotal;
      var margenPct = subtotalNeto > 0 ? margenLinea/subtotalNeto : 0;

      totalNetoDoc += subtotalNeto;
      totalIvaDoc += subtotalIva;
      totalBrutoDoc += subtotalBruto;

      var idLinea = generarID('VTA');
      var fila = new Array(34).fill('');
      fila[0]  = idLinea;
      fila[1]  = idDocumento;
      fila[2]  = empresa;
      fila[3]  = datos.tipoDoc;
      fila[4]  = numeroDoc;
      fila[9]  = '';
      fila[10] = fechaVenta;
      fila[11] = ctx.periodo;
      fila[12] = datos.cliente.id;
      fila[13] = linea.productoId;
      if (datos.tipoDoc === 'NOTA_CREDITO') fila[9] = datos.docOrigen;
      fila[16] = linea.cantidad;
      fila[17] = linea.precioNeto;
      fila[18] = (linea.descuentoPct||0)/100;
      fila[19] = subtotalNeto;
      fila[20] = subtotalIva;
      fila[21] = subtotalBruto;
      fila[22] = prod.costo;
      fila[23] = margenLinea;
      fila[24] = margenPct;
      fila[25] = 'CLP';
      fila[26] = 1;
      fila[27] = 'VIGENTE';
      fila[30] = new Date();
      fila[31] = usuario;
      filasVenta.push(fila);

      hP.getRange(prod.fila, 12).setValue(prod.stock - linea.cantidad);

      var hMov = ss.getSheetByName('DB_MOVIMIENTOS');
      var filaMov = _ultimaFilaReal(hMov) + 1;
      hMov.getRange(filaMov, 1, 1, 12).setValues([[
        generarID('MOV'), empresa, fechaVenta, ctx.periodo, linea.productoId,
        '', '', 'SALIDA_VENTA', linea.cantidad, -linea.cantidad, idDocumento, usuario
      ]]);
    });
    }

    var primeraFilaV = _ultimaFilaReal(hV) + 1;
    hV.getRange(primeraFilaV, 1, filasVenta.length, 34).setValues(filasVenta);

    var hI = ss.getSheetByName('DB_INGRESOS');
    var filaIngreso = _ultimaFilaReal(hI) + 1;
    hI.getRange(filaIngreso, 1, 1, 15).setValues([[
      generarID('ING'), empresa, fechaVenta, ctx.periodo, '4.1.1', '',
      (esNDSimple ? 'ND ' : 'Venta ') + numeroDoc + ' — ' + datos.cliente.nombre,
      totalNetoDoc, totalIvaDoc, totalBrutoDoc, numeroDoc, datos.tipoDoc,
      idDocumento, new Date(), usuario
    ]]);

    var generoCxC = false;
    if (!esEsContado) {
      var diasPlazo = datos.condicionPago === '30_DIAS' ? 30 : datos.condicionPago === '60_DIAS' ? 60 : 0;
      var fechaVenc = new Date(fechaVenta.getTime() + diasPlazo*24*60*60*1000);
      var hCxC = ss.getSheetByName('DB_CXC');
      var filaCxC = _ultimaFilaReal(hCxC) + 1;
      hCxC.getRange(filaCxC, 1, 1, 21).setValues([[
        generarID('CXC'), empresa, datos.cliente.id, '', numeroDoc, datos.tipoDoc,
        fechaVenta, fechaVenc, totalBrutoDoc, 0, totalBrutoDoc, diasPlazo, '0_30_DIAS',
        'PENDIENTE', '', '', '', '', '', new Date(), usuario
      ]]);
      generoCxC = true;
    }

    SpreadsheetApp.flush();

    return {
      ok: true,
      numeroDoc: numeroDoc,
      cliente: datos.cliente.nombre,
      totalNeto: totalNetoDoc,
      totalIva: totalIvaDoc,
      totalBruto: totalBrutoDoc,
      generoCxC: generoCxC
    };

  } catch(e) {
    Logger.log('procesarNuevaVenta ERROR: ' + e.message);
    return { ok:false, error: 'Error del sistema: ' + e.message };
  }
}
function validarRUT(rut) {
  rut = String(rut).replace(/[.\-]/g, '').toUpperCase();
  if (rut.length < 2) return false;
  var cuerpo = rut.slice(0, -1);
  var dv = rut.slice(-1);
  var suma = 0;
  var multiplo = 2;
  for (var i = cuerpo.length - 1; i >= 0; i--) {
    suma += Number(cuerpo[i]) * multiplo;
    multiplo = multiplo === 7 ? 2 : multiplo + 1;
  }
  var resto = 11 - (suma % 11);
  var dvCalculado = resto === 11 ? '0' : resto === 10 ? 'K' : String(resto);
  return dvCalculado === dv;
}

function buscarProveedores(termino) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hP = ss.getSheetByName('MSTR_PROVEEDORES');
    var filP = _ultimaFilaReal(hP);

    if (!termino || termino.trim() === '') return [];

    var terminoNorm = _normalizarTexto(termino).replace(/[.\-]/g,'');
    var resultados = [];

    if (filP >= 2) {
      hP.getRange(2,1,filP-1,24).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        if (f[18] !== true) return;
        var nombre = String(f[5]||'');
        var rut    = String(f[4]||'');
        var nombreNorm = _normalizarTexto(nombre);
        var rutNorm    = _normalizarTexto(rut).replace(/[.\-]/g,'');
        if (nombreNorm.indexOf(terminoNorm) !== -1 || rutNorm.indexOf(terminoNorm) !== -1) {
          resultados.push({
            id:String(f[0]), nombre:nombre, rut:rut,
            condicionPago:String(f[10]||'')
          });
        }
      });
    }
    return resultados.slice(0,10);
  } catch(e) {
    Logger.log('buscarProveedores ERROR: ' + e.message);
    return [];
  }
}

function buscarCuentasContables(termino) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hPC = ss.getSheetByName('MSTR_PLAN_CUENTAS');
    var filPC = hPC.getLastRow();

    if (!termino || termino.trim() === '') return [];

    var terminoNorm = _normalizarTexto(termino);
    var resultados = [];

    hPC.getRange(2,1,filPC-1,9).getValues().forEach(function(f) {
      if (f[0]==='') return;
      if (String(f[8]) !== '' && f[8] !== true) {} // activo, validación laxa
      var nivel = String(f[7]||'');
      if (nivel !== 'CUENTA') return; // solo cuentas imputables, no grupos
      if (String(f[1]) !== empresa) return;
      var nombre = String(f[2]||'');
      var id = String(f[0]||'');
      var nombreNorm = _normalizarTexto(nombre);
      var idNorm = _normalizarTexto(id);
      if (nombreNorm.indexOf(terminoNorm) !== -1 || idNorm.indexOf(terminoNorm) !== -1) {
        resultados.push({ id: id, nombre: nombre, tipo: String(f[3]||'') });
      }
    });
    return resultados.slice(0,10);
  } catch(e) {
    Logger.log('buscarCuentasContables ERROR: ' + e.message);
    return [];
  }
}
function procesarNuevaCompra(datos) {
  try {
    var ss  = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var usuario = Session.getActiveUser().getEmail() || 'sistema';

    // ── VALIDACIONES GENERALES ──
    if (!datos.proveedor || !datos.proveedor.id) {
      return { ok:false, error:'Debes seleccionar un proveedor o prestador' };
    }
    if (!datos.numeroDoc || String(datos.numeroDoc).trim() === '') {
      return { ok:false, error:'Debes ingresar el número de documento' };
    }

    var checkPeriodo = _validarPeriodoNoBloqueado(datos.fecha);
    if (!checkPeriodo.ok) return { ok:false, error: checkPeriodo.error };

    var tasaIva = Number(ss.getSheetByName('CONFIG_SISTEMA').getRange('L2').getValue()) || 0.19;
    var fechaCompra = new Date(datos.fecha);
    var esEsContado = (datos.condicionPago === 'CONTADO');

    var prefijo = datos.esMercaderia ? 'COM-' :
                  datos.tipoDoc === 'BOLETA_HONORARIOS' ? 'HON-' : 'GAS-';
    var idDocumento = prefijo + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS');

    var totalNetoDoc = 0;
    var totalIvaDoc = 0;
    var totalBrutoDoc = 0;
    var totalRetencion = 0;

    var esNDSimpleCompra = datos.tipoDoc === 'NOTA_DEBITO' && datos.esCargoSimple;

    // ── CASO C: NOTA DE DÉBITO — CARGO SIMPLE ──
    if (esNDSimpleCompra) {
      if (!datos.glosaCargo || String(datos.glosaCargo).trim() === '') {
        return { ok:false, error:'Debes ingresar una descripción del cargo' };
      }
      if (!datos.montoCargo || datos.montoCargo <= 0) {
        return { ok:false, error:'El monto debe ser mayor a cero' };
      }

      totalNetoDoc = Math.round(Number(datos.montoCargo));
      totalIvaDoc = datos.afectoIva ? Math.round(totalNetoDoc * tasaIva) : 0;
      totalBrutoDoc = totalNetoDoc + totalIvaDoc;

      var hEND = ss.getSheetByName('DB_EGRESOS');
      var filaEgrND = _ultimaFilaReal(hEND) + 1;
      hEND.getRange(filaEgrND, 1, 1, 18).setValues([[
        generarID('EGR'), empresa, fechaCompra, ctx.periodo, '5.2.6', datos.centroCosto || '',
        datos.glosaCargo + ' — ' + datos.proveedor.nombre,
        totalNetoDoc, totalIvaDoc, totalBrutoDoc, 'CLP', 1, totalBrutoDoc,
        datos.numeroDoc, datos.tipoDoc, idDocumento, new Date(), usuario
      ]]);

    // ── CASO A: COMPRA DE MERCADERÍA ──
    } else if (datos.esMercaderia) {
      if (!datos.lineas || !datos.lineas.length) {
        return { ok:false, error:'Debes agregar al menos una línea de producto' };
      }
      for (var i=0; i<datos.lineas.length; i++) {
        if (!datos.lineas[i].productoId) {
          return { ok:false, error:'Hay una línea sin producto seleccionado' };
        }
        if (!datos.lineas[i].cantidad || datos.lineas[i].cantidad <= 0) {
          return { ok:false, error:'Hay una línea con cantidad inválida' };
        }
      }

      var hP = ss.getSheetByName('MSTR_PRODUCTOS');
      var filP = _ultimaFilaReal(hP);
      var mapaProductos = {};
      if (filP >= 2) {
        hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
          if (f[0]==='') return;
          mapaProductos[String(f[0])] = { fila: idx+2, stock: Number(f[11])||0, costo: Number(f[7])||0, nombre:String(f[3]) };
        });
      }

      var hC = ss.getSheetByName('DB_COMPRAS');
      var filasCompra = [];

      datos.lineas.forEach(function(linea) {
        var prod = mapaProductos[linea.productoId];
        if (!prod) return;
        var subtotalNeto = Math.round(linea.cantidad * linea.costoUnitario * (1 - (linea.descuentoPct||0)/100));
        var subtotalIva = Math.round(subtotalNeto * tasaIva);
        var subtotalBruto = subtotalNeto + subtotalIva;

        totalNetoDoc += subtotalNeto;
        totalIvaDoc += subtotalIva;
        totalBrutoDoc += subtotalBruto;

        var fila = new Array(30).fill('');
        fila[0]  = generarID('COM');
        fila[1]  = idDocumento;
        fila[2]  = empresa;
        fila[3]  = datos.proveedor.id;
        fila[4]  = linea.productoId;
        fila[5]  = datos.centroCosto || '';
        fila[6]  = fechaCompra;
        fila[7]  = ctx.periodo;
        fila[10] = datos.numeroDoc;
        fila[11] = linea.cantidad;
        fila[12] = linea.cantidad;
        fila[13] = linea.costoUnitario;
        fila[14] = (linea.descuentoPct||0)/100;
        fila[15] = subtotalNeto;
        fila[16] = subtotalIva;
        fila[17] = subtotalBruto;
        fila[18] = 'CLP';
        fila[19] = 1;
        fila[20] = subtotalBruto;
        fila[22] = 'RECIBIDA';
        fila[25] = datos.urlDocumento || '';
        fila[26] = new Date();
        fila[27] = usuario;
        filasCompra.push(fila);

        // Actualizar stock y costo promedio ponderado (CPP)
        var stockNuevo = prod.stock + linea.cantidad;
        var costoNuevo;
        if (stockNuevo > 0) {
          costoNuevo = Math.round(
            (prod.stock * prod.costo + linea.cantidad * linea.costoUnitario) / stockNuevo
          );
        } else {
          costoNuevo = prod.costo;
        }
        hP.getRange(prod.fila, 12).setValue(stockNuevo);
        hP.getRange(prod.fila, 8).setValue(costoNuevo);

        // Movimiento de entrada
        var hMov = ss.getSheetByName('DB_MOVIMIENTOS');
        var filaMov = _ultimaFilaReal(hMov) + 1;
        hMov.getRange(filaMov, 1, 1, 12).setValues([[
          generarID('MOV'), empresa, fechaCompra, ctx.periodo, linea.productoId,
          '', '', 'ENTRADA_COMPRA', linea.cantidad, linea.cantidad, idDocumento, usuario
        ]]);
      });

      var primeraFilaC = _ultimaFilaReal(hC) + 1;
      hC.getRange(primeraFilaC, 1, filasCompra.length, 30).setValues(filasCompra);

      // Egreso contable — cuenta Existencias
      var hE = ss.getSheetByName('DB_EGRESOS');
      var filaEgr = _ultimaFilaReal(hE) + 1;
      hE.getRange(filaEgr, 1, 1, 18).setValues([[
        generarID('EGR'), empresa, fechaCompra, ctx.periodo, '1.1.4', datos.centroCosto || '',
        'Compra mercadería ' + datos.numeroDoc + ' — ' + datos.proveedor.nombre,
        totalNetoDoc, totalIvaDoc, totalBrutoDoc, 'CLP', 1, totalBrutoDoc,
        datos.numeroDoc, datos.tipoDoc, idDocumento, new Date(), usuario
      ]]);

    // ── CASO B: GASTO / ACTIVO / BOLETA HONORARIOS ──
    } else {
      if (!datos.cuentaContable || !datos.cuentaContable.id) {
        return { ok:false, error:'Debes seleccionar la cuenta contable' };
      }
      if (!datos.glosa || String(datos.glosa).trim() === '') {
        return { ok:false, error:'Debes ingresar una glosa/descripción' };
      }
      if (!datos.montoNeto || datos.montoNeto <= 0) {
        return { ok:false, error:'El monto debe ser mayor a cero' };
      }

      totalNetoDoc = Math.round(Number(datos.montoNeto));

      if (datos.tipoDoc === 'BOLETA_HONORARIOS') {
        var pctRetencion = Number(ss.getSheetByName('CONFIG_SISTEMA').getRange('Y2').getValue()) || 0.1525;
        totalRetencion = Math.round(totalNetoDoc * pctRetencion);
        totalBrutoDoc = totalNetoDoc - totalRetencion;
        totalIvaDoc = 0;
      } else {
        totalIvaDoc = datos.afectoIva ? Math.round(totalNetoDoc * tasaIva) : 0;
        totalBrutoDoc = totalNetoDoc + totalIvaDoc;
      }

      var hE2 = ss.getSheetByName('DB_EGRESOS');
      var filaEgr2 = _ultimaFilaReal(hE2) + 1;
      hE2.getRange(filaEgr2, 1, 1, 18).setValues([[
        generarID('EGR'), empresa, fechaCompra, ctx.periodo, datos.cuentaContable.id, datos.centroCosto || '',
        datos.glosa + ' — ' + datos.proveedor.nombre,
        totalNetoDoc, totalIvaDoc, totalBrutoDoc, 'CLP', 1, totalBrutoDoc,
        datos.numeroDoc, datos.tipoDoc, idDocumento, new Date(), usuario
      ]]);

      // Si hay retención, registrar pasivo
      if (totalRetencion > 0) {
        var hE3 = ss.getSheetByName('DB_EGRESOS');
        // La retención se registra como nota informativa en notas, no como egreso adicional aquí
      }
    }

    // ── ESCRIBIR DB_CXP — solo si NO es contado ──
    var generoCxP = false;
    if (!esEsContado) {
      var diasPlazo = datos.condicionPago === '30_DIAS' ? 30 : datos.condicionPago === '60_DIAS' ? 60 : 0;
      var fechaVenc = new Date(fechaCompra.getTime() + diasPlazo*24*60*60*1000);
      var hCxP = ss.getSheetByName('DB_CXP');
      var filaCxP = _ultimaFilaReal(hCxP) + 1;
      var montoCxP = totalRetencion > 0 ? totalBrutoDoc : totalBrutoDoc;
      hCxP.getRange(filaCxP, 1, 1, 20).setValues([[
        generarID('CXP'), empresa, datos.proveedor.id, datos.centroCosto || '', datos.numeroDoc, datos.tipoDoc,
        fechaCompra, fechaVenc, montoCxP, 0, 'CLP', 1, montoCxP, diasPlazo, '0_30_DIAS',
        'PENDIENTE', '', '', new Date(), usuario
      ]]);
      generoCxP = true;
    }

    SpreadsheetApp.flush();

    return {
      ok: true,
      numeroDoc: datos.numeroDoc,
      proveedor: datos.proveedor.nombre,
      totalNeto: totalNetoDoc,
      totalIva: totalIvaDoc,
      totalRetencion: totalRetencion,
      totalBruto: totalBrutoDoc,
      generoCxP: generoCxP
    };

  } catch(e) {
    Logger.log('procesarNuevaCompra ERROR: ' + e.message);
    return { ok:false, error: 'Error del sistema: ' + e.message };
  }
}
function buscarFacturasCliente(idCliente) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);

    if (!idCliente) return [];

    var facturasPorDoc = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        if (String(f[2]) !== empresa) return;
        if (String(f[12]) !== idCliente) return;
        var tipoDoc = String(f[3]||'');
        if (tipoDoc !== 'FACTURA') return;
        var estado = String(f[27]||'');
        if (estado === 'ANULADO') return;

        var numeroDoc = String(f[4]);
        var bruto = Number(f[21])||0;
        var fecha = f[10];

        if (!facturasPorDoc[numeroDoc]) {
          facturasPorDoc[numeroDoc] = {
            doc: numeroDoc,
            monto: 0,
            fecha: fecha instanceof Date ? fecha.getTime() : null
          };
        }
        facturasPorDoc[numeroDoc].monto += bruto;
      });
    }

    var resultado = Object.keys(facturasPorDoc).map(function(k) {
      return facturasPorDoc[k];
    }).sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    return resultado.slice(0,15);
  } catch(e) {
    Logger.log('buscarFacturasCliente ERROR: ' + e.message);
    return [];
  }
}

function buscarFacturasProveedor(idProveedor) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var hC = ss.getSheetByName('DB_COMPRAS');
    var filC = _ultimaFilaReal(hC);

    if (!idProveedor) return [];

    var facturasPorDoc = {};

    if (filC >= 2) {
      hC.getRange(2,1,filC-1,30).getValues().forEach(function(f) {
        if (f[0]==='') return;
        if (String(f[2]) !== empresa) return;
        if (String(f[3]) !== idProveedor) return;
        var numeroDoc = String(f[10]||'');
        if (!numeroDoc) return;
        var estado = String(f[22]||'');
        if (estado === 'ANULADO') return;

        var bruto = Number(f[20])||0;
        var fecha = f[6];

        if (!facturasPorDoc[numeroDoc]) {
          facturasPorDoc[numeroDoc] = {
            doc: numeroDoc,
            monto: 0,
            fecha: fecha instanceof Date ? fecha.getTime() : null
          };
        }
        facturasPorDoc[numeroDoc].monto += bruto;
      });
    }

    // También buscar facturas de gastos/activos en DB_EGRESOS (no son mercadería)
    var hE = ss.getSheetByName('DB_EGRESOS');
    var filE = _ultimaFilaReal(hE);
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='') return;
        if (String(f[1]) !== empresa) return;
        var numeroDoc = String(f[13]||'');
        var tipoDoc = String(f[14]||'');
        if (tipoDoc !== 'FACTURA' || !numeroDoc) return;
        if (facturasPorDoc[numeroDoc]) return; // ya está desde DB_COMPRAS

        var bruto = Number(f[12])||0;
        var fecha = f[2];
        facturasPorDoc[numeroDoc] = {
          doc: numeroDoc,
          monto: bruto,
          fecha: fecha instanceof Date ? fecha.getTime() : null
        };
      });
    }

    var resultado = Object.keys(facturasPorDoc).map(function(k) {
      return facturasPorDoc[k];
    }).sort(function(a,b){ return (b.fecha||0)-(a.fecha||0); });

    return resultado.slice(0,15);
  } catch(e) {
    Logger.log('buscarFacturasProveedor ERROR: ' + e.message);
    return [];
  }
}
function procesarCliente(datos) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var hC = ss.getSheetByName('MSTR_CLIENTES');

    if (!datos.razonSocial || String(datos.razonSocial).trim() === '') {
      return { ok:false, error:'La razón social es obligatoria' };
    }
    if (!datos.rut || String(datos.rut).trim() === '') {
      return { ok:false, error:'El RUT es obligatorio' };
    }
    if (!validarRUT(datos.rut)) {
      return { ok:false, error:'El RUT ingresado no es válido (dígito verificador incorrecto)' };
    }

    var rutNormalizado = String(datos.rut).replace(/[.\-]/g,'').toUpperCase();
    rutNormalizado = rutNormalizado.substring(0, rutNormalizado.length-1);

    var esEdicion = !!datos.id;
    var fechaAhora = new Date();

    if (esEdicion) {
      // Buscar la fila exacta del cliente
      var filC = _ultimaFilaReal(hC);
      var filaEncontrada = -1;
      var datosExistentes = hC.getRange(2,1,filC-1,23).getValues();
      for (var i=0; i<datosExistentes.length; i++) {
        if (String(datosExistentes[i][0]) === datos.id) { filaEncontrada = i+2; break; }
      }
      if (filaEncontrada === -1) return { ok:false, error:'Cliente no encontrado para editar' };

      hC.getRange(filaEncontrada, 4).setValue(rutNormalizado);
      hC.getRange(filaEncontrada, 5).setValue(String(datos.rut));
      hC.getRange(filaEncontrada, 6).setValue(datos.razonSocial);
      hC.getRange(filaEncontrada, 7).setValue(datos.nombreFantasia || '');
      hC.getRange(filaEncontrada, 8).setValue(datos.giro || '');
      hC.getRange(filaEncontrada, 9).setValue(datos.email || '');
      hC.getRange(filaEncontrada, 10).setValue(datos.telefono || '');
      hC.getRange(filaEncontrada, 11).setValue(datos.direccion || '');
      hC.getRange(filaEncontrada, 12).setValue(datos.comuna || '');
      hC.getRange(filaEncontrada, 15).setValue(datos.condicionPago || 'CONTADO');
      hC.getRange(filaEncontrada, 16).setValue(Number(datos.limiteCredito)||0);
      hC.getRange(filaEncontrada, 22).setValue(fechaAhora);
      hC.getRange(filaEncontrada, 23).setValue(usuario);

      SpreadsheetApp.flush();
      return { ok:true, id: datos.id, accion:'editado' };

    } else {
      // Verificar RUT duplicado
      var filC2 = _ultimaFilaReal(hC);
      if (filC2 >= 2) {
        var existentes = hC.getRange(2,1,filC2-1,5).getValues();
        for (var j=0; j<existentes.length; j++) {
          if (existentes[j][0]==='') continue;
          if (String(existentes[j][3]) === rutNormalizado) {
            return { ok:false, error:'Ya existe un cliente registrado con ese RUT' };
          }
        }
      }

      var idNuevo = generarID('CLI');
      var fila = new Array(23).fill('');
      fila[0]  = idNuevo;
      fila[1]  = empresa;
      fila[2]  = 'EMPRESA';
      fila[3]  = rutNormalizado;
      fila[4]  = String(datos.rut);
      fila[5]  = datos.razonSocial;
      fila[6]  = datos.nombreFantasia || '';
      fila[7]  = datos.giro || '';
      fila[8]  = datos.email || '';
      fila[9]  = datos.telefono || '';
      fila[10] = datos.direccion || '';
      fila[11] = datos.comuna || '';
      fila[14] = datos.condicionPago || 'CONTADO';
      fila[15] = Number(datos.limiteCredito)||0;
      fila[17] = true;
      fila[19] = fechaAhora;
      fila[20] = usuario;

      var primeraFilaVacia = _ultimaFilaReal(hC) + 1;
      hC.getRange(primeraFilaVacia, 1, 1, 23).setValues([fila]);
      SpreadsheetApp.flush();

      return { ok:true, id: idNuevo, accion:'creado' };
    }

  } catch(e) {
    Logger.log('procesarCliente ERROR: ' + e.message);
    return { ok:false, error: 'Error del sistema: ' + e.message };
  }
}

function procesarProveedor(datos) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var hP = ss.getSheetByName('MSTR_PROVEEDORES');

    if (!datos.razonSocial || String(datos.razonSocial).trim() === '') {
      return { ok:false, error:'La razón social es obligatoria' };
    }
    if (!datos.rut || String(datos.rut).trim() === '') {
      return { ok:false, error:'El RUT es obligatorio' };
    }
    if (!validarRUT(datos.rut)) {
      return { ok:false, error:'El RUT ingresado no es válido (dígito verificador incorrecto)' };
    }

    var rutNormalizado = String(datos.rut).replace(/[.\-]/g,'').toUpperCase();
    rutNormalizado = rutNormalizado.substring(0, rutNormalizado.length-1);

    var esEdicion = !!datos.id;
    var fechaAhora = new Date();

    if (esEdicion) {
      var filP = _ultimaFilaReal(hP);
      var filaEncontrada = -1;
      var datosExistentes = hP.getRange(2,1,filP-1,24).getValues();
      for (var i=0; i<datosExistentes.length; i++) {
        if (String(datosExistentes[i][0]) === datos.id) { filaEncontrada = i+2; break; }
      }
      if (filaEncontrada === -1) return { ok:false, error:'Proveedor no encontrado para editar' };

      hP.getRange(filaEncontrada, 4).setValue(rutNormalizado);
      hP.getRange(filaEncontrada, 5).setValue(String(datos.rut));
      hP.getRange(filaEncontrada, 6).setValue(datos.razonSocial);
      hP.getRange(filaEncontrada, 7).setValue(datos.nombreFantasia || '');
      hP.getRange(filaEncontrada, 8).setValue(datos.giro || '');
      hP.getRange(filaEncontrada, 9).setValue(datos.email || '');
      hP.getRange(filaEncontrada, 10).setValue(datos.telefono || '');
      hP.getRange(filaEncontrada, 11).setValue(datos.condicionPago || 'CONTADO');
      hP.getRange(filaEncontrada, 12).setValue(Number(datos.plazoEntrega)||0);
      hP.getRange(filaEncontrada, 23).setValue(fechaAhora);
      hP.getRange(filaEncontrada, 24).setValue(usuario);

      SpreadsheetApp.flush();
      return { ok:true, id: datos.id, accion:'editado' };

    } else {
      var filP2 = _ultimaFilaReal(hP);
      if (filP2 >= 2) {
        var existentes = hP.getRange(2,1,filP2-1,5).getValues();
        for (var j=0; j<existentes.length; j++) {
          if (existentes[j][0]==='') continue;
          if (String(existentes[j][3]) === rutNormalizado) {
            return { ok:false, error:'Ya existe un proveedor registrado con ese RUT' };
          }
        }
      }

      var idNuevo = generarID('PRV');
      var fila = new Array(24).fill('');
      fila[0]  = idNuevo;
      fila[1]  = empresa;
      fila[2]  = 'EMPRESA';
      fila[3]  = rutNormalizado;
      fila[4]  = String(datos.rut);
      fila[5]  = datos.razonSocial;
      fila[6]  = datos.nombreFantasia || '';
      fila[7]  = datos.giro || '';
      fila[8]  = datos.email || '';
      fila[9]  = datos.telefono || '';
      fila[10] = datos.condicionPago || 'CONTADO';
      fila[11] = Number(datos.plazoEntrega)||0;
      fila[12] = 'CLP';
      fila[13] = false;
      fila[18] = true;
      fila[20] = fechaAhora;
      fila[21] = usuario;

      var primeraFilaVacia = _ultimaFilaReal(hP) + 1;
      hP.getRange(primeraFilaVacia, 1, 1, 24).setValues([fila]);
      SpreadsheetApp.flush();

      return { ok:true, id: idNuevo, accion:'creado' };
    }

  } catch(e) {
    Logger.log('procesarProveedor ERROR: ' + e.message);
    return { ok:false, error: 'Error del sistema: ' + e.message };
  }
}
function listarClientes() {
  var ss = _getSS();
  var ctx = getDatosContexto();
  var empresa = ctx.idEmpresa;
  var h = ss.getSheetByName('MSTR_CLIENTES');
  var fil = _ultimaFilaReal(h);
  var lista = [];
  if (fil >= 2) {
    h.getRange(2,1,fil-1,23).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      lista.push({ id:String(f[0]), nombre:String(f[5]||''), rut:String(f[4]||''), activo: f[17]===true });
    });
  }
  return lista.sort(function(a,b){ return a.nombre.localeCompare(b.nombre); });
}

function listarProveedores() {
  var ss = _getSS();
  var ctx = getDatosContexto();
  var empresa = ctx.idEmpresa;
  var h = ss.getSheetByName('MSTR_PROVEEDORES');
  var fil = _ultimaFilaReal(h);
  var lista = [];
  if (fil >= 2) {
    h.getRange(2,1,fil-1,24).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      lista.push({ id:String(f[0]), nombre:String(f[5]||''), rut:String(f[4]||''), activo: f[18]===true });
    });
  }
  return lista.sort(function(a,b){ return a.nombre.localeCompare(b.nombre); });
}
function _calcularSaldoCxC(idCxc) {
  var ss = _getSS();
  var hCxC = ss.getSheetByName('DB_CXC');
  var filCxC = _ultimaFilaReal(hCxC);
  var docOriginal = null;

  if (filCxC >= 2) {
    var datos = hCxC.getRange(2,1,filCxC-1,21).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === idCxc) {
        docOriginal = { numeroDoc: String(datos[i][4]), montoOriginal: Number(datos[i][8])||0, idCliente: String(datos[i][2]) };
        break;
      }
    }
  }
  if (!docOriginal) return null;

  var sumaCobros = 0;
  var hCob = ss.getSheetByName('DB_COBROS_CXC');
  var filCob = _ultimaFilaReal(hCob);
  if (filCob >= 2) {
    hCob.getRange(2,1,filCob-1,18).getValues().forEach(function(f) {
      if (f[0]==='') return;
      if (String(f[3]) === idCxc && f[12] !== true) sumaCobros += Number(f[7])||0;
    });
  }

  var sumaNC = 0;
  if (filCxC >= 2) {
    hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
      if (f[0]==='') return;
      if (String(f[5]) === 'NOTA_CREDITO' && String(f[3]) === docOriginal.numeroDoc) sumaNC += Number(f[8])||0;
    });
  }

  var saldo = docOriginal.montoOriginal - sumaCobros - sumaNC;
  return { saldo: Math.max(saldo,0), montoOriginal: docOriginal.montoOriginal, idCliente: docOriginal.idCliente, numeroDoc: docOriginal.numeroDoc };
}

function listarMediosPago() {
  var ss = _getSS();
  var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
  var fil = _ultimaFilaReal(h);
  var lista = [];
  if (fil >= 2) {
    h.getRange(2,1,fil-1,5).getValues().forEach(function(f) {
      if (f[0]==='' || f[3]!==true) return;
      lista.push({ id:String(f[0]), nombre:String(f[2]) });
    });
  }
  return lista.sort(function(a,b){ return a.nombre.localeCompare(b.nombre); });
}

function registrarCobroCxC(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!datos.idCxc) return { ok:false, error:'Falta documento de referencia' };
    var monto = Number(datos.monto);
    if (!monto || monto <= 0) return { ok:false, error:'El monto debe ser mayor a cero' };
    if (!datos.idMedioPago) return { ok:false, error:'Debes seleccionar un medio de pago' };
    var fechaCobro = new Date(datos.fechaCobro);
    if (fechaCobro > new Date()) return { ok:false, error:'La fecha de cobro no puede ser futura' };

    var checkPeriodo = _validarPeriodoNoBloqueado(fechaCobro);
    if (!checkPeriodo.ok) return { ok:false, error: checkPeriodo.error };

    var info = _calcularSaldoCxC(datos.idCxc);
    if (!info) return { ok:false, error:'Documento no encontrado' };
    if (monto > info.saldo) return { ok:false, error:'Saldo insuficiente. Disponible: $' + info.saldo.toLocaleString('es-CL') };

    var ss = _getSS();
    var h = ss.getSheetByName('DB_COBROS_CXC');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();
    var idCobro = generarID('COB');

    var fila = new Array(18).fill('');
    fila[0]  = idCobro;
    fila[1]  = Utilities.getUuid();
    fila[2]  = ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[3]  = datos.idCxc;
    fila[4]  = info.idCliente;
    fila[5]  = fechaCobro;
    fila[6]  = ahora;
    fila[7]  = monto;
    fila[8]  = datos.idMedioPago;
    fila[9]  = datos.referenciaPago || '';
    fila[10] = false;
    fila[11] = '';
    fila[12] = false;
    fila[13] = '';
    fila[14] = '';
    fila[15] = '';
    fila[16] = ahora;
    fila[17] = usuario;

    var primeraFilaVacia = _ultimaFilaReal(h) + 1;
    h.getRange(primeraFilaVacia, 1, 1, 18).setValues([fila]);
    SpreadsheetApp.flush();

    var saldoNuevo = info.saldo - monto;
    actualizarSaldosVivosEnCalc();
    return { ok:true, idCobro:idCobro, saldoRestante:saldoNuevo, pagadoTotal: saldoNuevo<=0 };

  } catch(e) {
    Logger.log('registrarCobroCxC ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function anularCobroCxC(idCobro, motivo) {
  try {
    if (!motivo || motivo.trim()==='') return { ok:false, error:'El motivo es obligatorio' };
    var ss = _getSS();
    var h = ss.getSheetByName('DB_COBROS_CXC');
    var fil = _ultimaFilaReal(h);
    if (fil < 2) return { ok:false, error:'Sin registros' };

    var datos = h.getRange(2,1,fil-1,18).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === idCobro) {
        if (datos[i][12] === true) return { ok:false, error:'Este cobro ya fue anulado anteriormente' };
        var filaReal = i+2;
        h.getRange(filaReal, 13).setValue(true);
        h.getRange(filaReal, 14).setValue(motivo);
        h.getRange(filaReal, 15).setValue(Session.getActiveUser().getEmail() || 'sistema');
        h.getRange(filaReal, 16).setValue(new Date());
        SpreadsheetApp.flush();
        actualizarSaldosVivosEnCalc();
        return { ok:true };
      }
    }
    return { ok:false, error:'Cobro no encontrado' };
  } catch(e) {
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function getDetalleDocumentoVenta(numeroDoc) {
  try {
    var ss = _getSS();
    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var lineas = [];
    var cabecera = null;

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[4])!==numeroDoc) return;
        if (!cabecera) {
          cabecera = {
            idDocumento: String(f[1]),
            tipoDoc: String(f[3]),
            numeroDoc: String(f[4]),
            fecha: f[10] instanceof Date ? f[10].getTime() : null,
            estado: String(f[27]||''),
            idCliente: String(f[12]||'')
          };
        }
        if (f[13]) {
          lineas.push({
            productoId: String(f[13]),
            cantidad: Number(f[16])||0,
            precioNeto: Number(f[17])||0,
            descuentoPct: Number(f[18])||0,
            subtotalNeto: Number(f[19])||0,
            subtotalIva: Number(f[20])||0,
            subtotalBruto: Number(f[21])||0
          });
        } else if (String(f[6]||'') !== '') {
          cabecera.glosaCargo = String(f[6]);
          cabecera.montoCargo = Number(f[19])||0;
        }
      });
    }

    var mapProd = {};
    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapProd[String(f[0])] = String(f[3]||f[2]);
      });
    }
    lineas.forEach(function(l) { l.nombre = mapProd[l.productoId] || l.productoId; });

    if (!cabecera) return null;
    cabecera.lineas = lineas;
    cabecera.totalNeto = lineas.reduce(function(s,l){return s+l.subtotalNeto;},0) + (cabecera.montoCargo||0);
    cabecera.totalIva = lineas.reduce(function(s,l){return s+l.subtotalIva;},0);
    cabecera.totalBruto = lineas.reduce(function(s,l){return s+l.subtotalBruto;},0) + (cabecera.montoCargo||0);
    return cabecera;

  } catch(e) {
    Logger.log('getDetalleDocumentoVenta ERROR: ' + e.message);
    return null;
  }
}
function getDetalleDocumentoCompra(numeroDoc) {
  try {
    var ss = _getSS();
    var hC = ss.getSheetByName('DB_COMPRAS');
    var filC = _ultimaFilaReal(hC);
    var lineas = [];
    var cabecera = null;

    if (filC >= 2) {
      hC.getRange(2,1,filC-1,30).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[10])!==numeroDoc) return;
        if (!cabecera) {
          cabecera = {
            idDocumento: String(f[1]),
            tipoDoc: 'COMPRA',
            numeroDoc: String(f[10]),
            fecha: f[6] instanceof Date ? f[6].getTime() : null,
            estado: String(f[22]||''),
            idProveedor: String(f[3]||'')
          };
        }
        lineas.push({
          productoId: String(f[4]),
          cantidad: Number(f[11])||0,
          costoUnitario: Number(f[13])||0,
          descuentoPct: Number(f[14])||0,
          subtotalNeto: Number(f[15])||0,
          subtotalIva: Number(f[16])||0,
          subtotalBruto: Number(f[17])||0
        });
      });
    }

    if (!cabecera) {
      var hE = ss.getSheetByName('DB_EGRESOS');
      var filE = _ultimaFilaReal(hE);
      if (filE >= 2) {
        hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
          if (f[0]==='' || String(f[13])!==numeroDoc) return;
          cabecera = {
            tipoDoc: String(f[14]||'GASTO'),
            numeroDoc: String(f[13]),
            fecha: f[2] instanceof Date ? f[2].getTime() : null,
            estado: 'VIGENTE',
            glosaCargo: String(f[6]||''),
            cuenta: String(f[4]||''),
            montoCargo: Number(f[7])||0,
            ivaCargo: Number(f[8])||0,
            totalCargo: Number(f[9])||0
          };
        });
      }
    }

    var mapProd = {};
    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapProd[String(f[0])] = String(f[3]||f[2]);
      });
    }
    lineas.forEach(function(l) { l.nombre = mapProd[l.productoId] || l.productoId; });

    if (!cabecera) return null;
    cabecera.lineas = lineas;
    cabecera.totalNeto = lineas.length ? lineas.reduce(function(s,l){return s+l.subtotalNeto;},0) : (cabecera.montoCargo||0);
    cabecera.totalIva = lineas.length ? lineas.reduce(function(s,l){return s+l.subtotalIva;},0) : (cabecera.ivaCargo||0);
    cabecera.totalBruto = lineas.length ? lineas.reduce(function(s,l){return s+l.subtotalBruto;},0) : (cabecera.totalCargo||0);
    return cabecera;

  } catch(e) {
    Logger.log('getDetalleDocumentoCompra ERROR: ' + e.message);
    return null;
  }
}
function _calcularSaldoCxP(idCxp) {
  var ss = _getSS();
  var hCxP = ss.getSheetByName('DB_CXP');
  var filCxP = _ultimaFilaReal(hCxP);
  var docOriginal = null;

  if (filCxP >= 2) {
    var datos = hCxP.getRange(2,1,filCxP-1,20).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === idCxp) {
        docOriginal = { numeroDoc: String(datos[i][4]), montoOriginal: Number(datos[i][8])||0, idProveedor: String(datos[i][2]) };
        break;
      }
    }
  }
  if (!docOriginal) return null;

  var sumaPagos = 0;
  var hPag = ss.getSheetByName('DB_PAGOS_CXP');
  var filPag = _ultimaFilaReal(hPag);
  if (filPag >= 2) {
    hPag.getRange(2,1,filPag-1,18).getValues().forEach(function(f) {
      if (f[0]==='') return;
      if (String(f[3]) === idCxp && f[12] !== true) sumaPagos += Number(f[7])||0;
    });
  }

  var saldo = docOriginal.montoOriginal - sumaPagos;
  return { saldo: Math.max(saldo,0), montoOriginal: docOriginal.montoOriginal, idProveedor: docOriginal.idProveedor, numeroDoc: docOriginal.numeroDoc };
}

function registrarPagoCxP(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!datos.idCxp) return { ok:false, error:'Falta documento de referencia' };
    var monto = Number(datos.monto);
    if (!monto || monto <= 0) return { ok:false, error:'El monto debe ser mayor a cero' };
    if (!datos.idMedioPago) return { ok:false, error:'Debes seleccionar un medio de pago' };
    var fechaPago = new Date(datos.fechaPago);
    if (fechaPago > new Date()) return { ok:false, error:'La fecha de pago no puede ser futura' };

    var checkPeriodo = _validarPeriodoNoBloqueado(fechaPago);
    if (!checkPeriodo.ok) return { ok:false, error: checkPeriodo.error };

    var info = _calcularSaldoCxP(datos.idCxp);
    if (!info) return { ok:false, error:'Documento no encontrado' };
    if (monto > info.saldo) return { ok:false, error:'Saldo insuficiente. Disponible: $' + info.saldo.toLocaleString('es-CL') };

    var ss = _getSS();
    var h = ss.getSheetByName('DB_PAGOS_CXP');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();
    var idPago = generarID('PAG');

    var fila = new Array(18).fill('');
    fila[0]  = idPago;
    fila[1]  = Utilities.getUuid();
    fila[2]  = ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[3]  = datos.idCxp;
    fila[4]  = info.idProveedor;
    fila[5]  = fechaPago;
    fila[6]  = ahora;
    fila[7]  = monto;
    fila[8]  = datos.idMedioPago;
    fila[9]  = datos.referenciaPago || '';
    fila[10] = false;
    fila[11] = '';
    fila[12] = false;
    fila[13] = '';
    fila[14] = '';
    fila[15] = '';
    fila[16] = ahora;
    fila[17] = usuario;

    var primeraFilaVacia = _ultimaFilaReal(h) + 1;
    h.getRange(primeraFilaVacia, 1, 1, 18).setValues([fila]);
    SpreadsheetApp.flush();

    var saldoNuevo = info.saldo - monto;
    actualizarSaldosVivosEnCalc();
    return { ok:true, idPago:idPago, saldoRestante:saldoNuevo, pagadoTotal: saldoNuevo<=0 };

  } catch(e) {
    Logger.log('registrarPagoCxP ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function anularPagoCxP(idPago, motivo) {
  try {
    if (!motivo || motivo.trim()==='') return { ok:false, error:'El motivo es obligatorio' };
    var ss = _getSS();
    var h = ss.getSheetByName('DB_PAGOS_CXP');
    var fil = _ultimaFilaReal(h);
    if (fil < 2) return { ok:false, error:'Sin registros' };

    var datos = h.getRange(2,1,fil-1,18).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === idPago) {
        if (datos[i][12] === true) return { ok:false, error:'Este pago ya fue anulado anteriormente' };
        var filaReal = i+2;
        h.getRange(filaReal, 13).setValue(true);
        h.getRange(filaReal, 14).setValue(motivo);
        h.getRange(filaReal, 15).setValue(Session.getActiveUser().getEmail() || 'sistema');
        h.getRange(filaReal, 16).setValue(new Date());
        SpreadsheetApp.flush();
        actualizarSaldosVivosEnCalc();
        return { ok:true };
      }
    }
    return { ok:false, error:'Pago no encontrado' };
  } catch(e) {
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function _getMapaSaldosCxC() {
  var ss = _getSS();
  var hCxC = ss.getSheetByName('DB_CXC');
  var filCxC = _ultimaFilaReal(hCxC);
  var mapa = {};
  var hoy = new Date();
  hoy.setHours(0,0,0,0);

  if (filCxC >= 2) {
    hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
      if (f[0]==='') return;
      var idCxc = String(f[0]);
      mapa[idCxc] = {
        idCxc: idCxc,
        idEmpresa: String(f[1]),
        idCliente: String(f[2]),
        numeroDoc: String(f[4]),
        tipoDoc: String(f[5]),
        emision: f[6],
        vence: f[7],
        montoOriginal: Number(f[8])||0,
        sumaCobros: 0,
        sumaNC: 0
      };
    });
  }

  // Sumar NC aplicadas (id_referencia = numeroDoc del documento original)
  if (filCxC >= 2) {
    hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[5]) !== 'NOTA_CREDITO') return;
      var docRef = String(f[17]||''); // id_referencia
      Object.keys(mapa).forEach(function(k) {
        if (mapa[k].numeroDoc === docRef) mapa[k].sumaNC += Number(f[8])||0;
      });
    });
  }

  // Sumar cobros no anulados
  var hCob = ss.getSheetByName('DB_COBROS_CXC');
  var filCob = _ultimaFilaReal(hCob);
  if (filCob >= 2) {
    hCob.getRange(2,1,filCob-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || f[12]===true) return;
      var idCxc = String(f[3]);
      if (mapa[idCxc]) mapa[idCxc].sumaCobros += Number(f[7])||0;
    });
  }

  // Calcular saldo vivo, días, tramo, estado
  Object.keys(mapa).forEach(function(k) {
    var m = mapa[k];
    m.saldoVivo = Math.max(m.montoOriginal - m.sumaCobros - m.sumaNC, 0);
    if (m.vence instanceof Date) {
      var v = new Date(m.vence); v.setHours(0,0,0,0);
      m.dias = Math.round((v - hoy) / (1000*60*60*24));
    } else {
      m.dias = 0;
    }
    if (m.dias < 0) m.tramo = 'VENCIDO';
    else if (m.dias <= 30) m.tramo = '0_30_DIAS';
    else if (m.dias <= 60) m.tramo = '31_60_DIAS';
    else m.tramo = 'MAS_60_DIAS';
    m.estado = m.saldoVivo <= 0 ? 'PAGADO' : 'PENDIENTE';
  });

  return mapa;
}

function _getMapaSaldosCxP() {
  var ss = _getSS();
  var hCxP = ss.getSheetByName('DB_CXP');
  var filCxP = _ultimaFilaReal(hCxP);
  var mapa = {};
  var hoy = new Date();
  hoy.setHours(0,0,0,0);

  if (filCxP >= 2) {
    hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f) {
      if (f[0]==='') return;
      var idCxp = String(f[0]);
      mapa[idCxp] = {
        idCxp: idCxp,
        idEmpresa: String(f[1]),
        idProveedor: String(f[2]),
        numeroDoc: String(f[4]),
        tipoDoc: String(f[5]),
        emision: f[6],
        vence: f[7],
        montoOriginal: Number(f[8])||0,
        sumaPagos: 0
      };
    });
  }

  var hPag = ss.getSheetByName('DB_PAGOS_CXP');
  var filPag = _ultimaFilaReal(hPag);
  if (filPag >= 2) {
    hPag.getRange(2,1,filPag-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || f[12]===true) return;
      var idCxp = String(f[3]);
      if (mapa[idCxp]) mapa[idCxp].sumaPagos += Number(f[7])||0;
    });
  }

  Object.keys(mapa).forEach(function(k) {
    var m = mapa[k];
    m.saldoVivo = Math.max(m.montoOriginal - m.sumaPagos, 0);
    if (m.vence instanceof Date) {
      var v = new Date(m.vence); v.setHours(0,0,0,0);
      m.dias = Math.round((v - hoy) / (1000*60*60*24));
    } else {
      m.dias = 0;
    }
    if (m.dias < 0) m.tramo = 'VENCIDO';
    else if (m.dias <= 5) m.tramo = '0_5_DIAS';
    else if (m.dias <= 30) m.tramo = '6_30_DIAS';
    else m.tramo = 'MAS_30_DIAS';
    m.estado = m.saldoVivo <= 0 ? 'PAGADO' : 'PENDIENTE';
  });

  return mapa;
}
function actualizarSaldosVivosEnCalc() {
  var ss = _getSS();
  var cf = ss.getSheetByName('CALC_FINANCIERO');
  var ctx = getDatosContexto();
  var empresa = ctx.idEmpresa;

  var mapaCxC = _getMapaSaldosCxC();
  var totalCxC=0, cxcVencida=0, cxc030=0, cxc3160=0, cxcMas60=0, docsP=0, docsV=0;
  Object.keys(mapaCxC).forEach(function(k) {
    var m = mapaCxC[k];
    if (m.idEmpresa !== empresa || m.saldoVivo <= 0) return;
    totalCxC += m.saldoVivo;
    docsP++;
    if (m.dias < 0) { cxcVencida += m.saldoVivo; docsV++; }
    else if (m.dias <= 30) cxc030 += m.saldoVivo;
    else if (m.dias <= 60) cxc3160 += m.saldoVivo;
    else cxcMas60 += m.saldoVivo;
  });

  var mapaCxP = _getMapaSaldosCxP();
  var totalCxP=0, cxpVencida=0, cxpUrgente=0, cxp630=0, cxpMas30=0, docsPP=0, docsVP=0;
  Object.keys(mapaCxP).forEach(function(k) {
    var m = mapaCxP[k];
    if (m.idEmpresa !== empresa || m.saldoVivo <= 0) return;
    totalCxP += m.saldoVivo;
    docsPP++;
    if (m.dias < 0) { cxpVencida += m.saldoVivo; docsVP++; }
    if ((m.dias >= 0 && m.dias <= 5) || m.dias < 0) cxpUrgente += m.saldoVivo;
    else if (m.dias <= 30) cxp630 += m.saldoVivo;
    else cxpMas30 += m.saldoVivo;
  });

  cf.getRange('B17').setValue(totalCxC);
  cf.getRange('B18').setValue(cxcVencida);
  cf.getRange('B19').setValue(cxc030);
  cf.getRange('B20').setValue(cxc3160);
  cf.getRange('B21').setValue(cxcMas60);
  cf.getRange('B22').setValue(docsP);
  cf.getRange('B23').setValue(docsV);

  cf.getRange('B30').setValue(totalCxP);
  cf.getRange('B31').setValue(cxpVencida);
  cf.getRange('B32').setValue(cxpUrgente);
  cf.getRange('B33').setValue(cxp630);
  cf.getRange('B34').setValue(cxpMas30);
  cf.getRange('B35').setValue(docsPP);
  cf.getRange('B36').setValue(docsVP);

  cf.getRange('B38').setValue(cxpVencida + cxpUrgente);

  var anulados = _getDocumentosAnulados();
  var ingresosNeto = 0, egresosNeto = 0;
  var periodoActual = ctx.periodo;

  var hI = ss.getSheetByName('DB_INGRESOS');
  var filI = _ultimaFilaReal(hI);
  if (filI >= 2) {
    hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      if (anulados.ventas[String(f[12]||'')]) return;
      var fecha = f[2];
      if (!(fecha instanceof Date)) return;
      var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
      if (per === periodoActual) ingresosNeto += Number(f[7])||0;
    });
  }

  var hE = ss.getSheetByName('DB_EGRESOS');
  var filE = _ultimaFilaReal(hE);
  if (filE >= 2) {
    hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      if (anulados.compras[String(f[15]||'')]) return;
      var fecha = f[2];
      if (!(fecha instanceof Date)) return;
      var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
      if (per === periodoActual) egresosNeto += Number(f[7])||0;
    });
  }

  cf.getRange('B8').setValue(ingresosNeto);
  cf.getRange('B10').setValue(egresosNeto);

  SpreadsheetApp.flush();
  Logger.log('✅ Saldos vivos actualizados en CALC_FINANCIERO');
}
function _calcularCajaReal(empresa, ctx, anuladosPrevios) {
  var ss = _getSS();
  var anulados = anuladosPrevios || _getDocumentosAnulados();

  var ventasContado = 0;
  var hV = ss.getSheetByName('DB_VENTAS');
  var filV = _ultimaFilaReal(hV);
  if (filV >= 2) {
    hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[2])!==empresa) return;
      if (String(f[27]||'')==='ANULADO') return;
      var idCli = String(f[12]||'');
      // Contado = no generó CxC (no existe fila en DB_CXC con id_referencia = id_documento)
    });
  }

  // Ventas contado: sumamos directo desde DB_INGRESOS cruzando con DB_CXC (si no generó CxC, es contado)
  var idsConCxC = {};
  var hCxC = ss.getSheetByName('DB_CXC');
  var filCxC = _ultimaFilaReal(hCxC);
  if (filCxC >= 2) {
    hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
      if (f[0]==='') return;
      idsConCxC[String(f[17])] = true; // id_referencia = id_documento de venta
    });
  }

  var hI = ss.getSheetByName('DB_INGRESOS');
  var filI = _ultimaFilaReal(hI);
  if (filI >= 2) {
    hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      var idOrigen = String(f[12]||'');
      if (anulados.ventas[idOrigen]) return;
      if (!idsConCxC[idOrigen]) ventasContado += Number(f[7])||0;
    });
  }

  var cobrosCxC = 0;
  var hCob = ss.getSheetByName('DB_COBROS_CXC');
  var filCob = _ultimaFilaReal(hCob);
  if (filCob >= 2) {
    hCob.getRange(2,1,filCob-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || f[12]===true) return;
      cobrosCxC += Number(f[7])||0;
    });
  }

  var idsConCxP = {};
  var hCxP = ss.getSheetByName('DB_CXP');
  var filCxP = _ultimaFilaReal(hCxP);
  if (filCxP >= 2) {
    hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f) {
      if (f[0]==='') return;
      idsConCxP[String(f[4])] = true; // id_documento del egreso/compra
    });
  }

  var comprasContado = 0;
  var hE = ss.getSheetByName('DB_EGRESOS');
  var filE = _ultimaFilaReal(hE);
  if (filE >= 2) {
    hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[1])!==empresa) return;
      var idOrigen = String(f[15]||'');
      if (anulados.compras[idOrigen]) return;
      if (!idsConCxP[String(f[13])]) comprasContado += Number(f[9])||0;
    });
  }

  var pagosCxP = 0;
  var hPag = ss.getSheetByName('DB_PAGOS_CXP');
  var filPag = _ultimaFilaReal(hPag);
  if (filPag >= 2) {
    hPag.getRange(2,1,filPag-1,18).getValues().forEach(function(f) {
      if (f[0]==='' || f[12]===true) return;
      pagosCxP += Number(f[7])||0;
    });
  }

  return Math.round(ventasContado + cobrosCxC - comprasContado - pagosCxP);
}
function _getDocumentosAnulados() {
  var ss = _getSS();
  var ventas = {};
  var compras = {};

  var hV = ss.getSheetByName('DB_VENTAS');
  var filV = _ultimaFilaReal(hV);
  if (filV >= 2) {
    hV.getRange(2,1,filV-1,28).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[27])!=='ANULADO') return;
      ventas[String(f[1])] = true; // id_documento
    });
  }

  var hC = ss.getSheetByName('DB_COMPRAS');
  var filC = _ultimaFilaReal(hC);
  if (filC >= 2) {
    hC.getRange(2,1,filC-1,23).getValues().forEach(function(f) {
      if (f[0]==='' || String(f[22])!=='ANULADO') return;
      compras[String(f[1])] = true; // id_documento
    });
  }

  return { ventas: ventas, compras: compras };
}

function getClientePorId(idCliente) {
  var ss = _getSS();
  var h = ss.getSheetByName('MSTR_CLIENTES');
  var fil = _ultimaFilaReal(h);
  if (fil < 2) return null;
  var datos = h.getRange(2,1,fil-1,15).getValues();
  for (var i=0; i<datos.length; i++) {
    if (String(datos[i][0]) === idCliente) {
      return { id: idCliente, nombre: String(datos[i][5]||''), rut: String(datos[i][4]||''), condicionPago: String(datos[i][14]||'') };
    }
  }
  return null;
}

function getProveedorPorId(idProveedor) {
  var ss = _getSS();
  var h = ss.getSheetByName('MSTR_PROVEEDORES');
  var fil = _ultimaFilaReal(h);
  if (fil < 2) return null;
  var datos = h.getRange(2,1,fil-1,15).getValues();
  for (var i=0; i<datos.length; i++) {
    if (String(datos[i][0]) === idProveedor) {
      return { id: idProveedor, nombre: String(datos[i][5]||''), rut: String(datos[i][4]||''), condicionPago: String(datos[i][10]||'') };
    }
  }
  return null;
}
function getLineasParaNC(numeroDoc, tipo) {
  try {
    var ss = _getSS();
    var lineas = [];
    var cliente = null;

    if (tipo === 'venta') {
      var hV = ss.getSheetByName('DB_VENTAS');
      var filV = _ultimaFilaReal(hV);
      var hP = ss.getSheetByName('MSTR_PRODUCTOS');
      var filP = _ultimaFilaReal(hP);
      var mapProd = {};
      if (filP >= 2) {
        hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
          if (f[0]==='') return;
          mapProd[String(f[0])] = String(f[3]||f[2]);
        });
      }
      var mapStock = {};
      if (filP >= 2) {
        hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
          if (f[0]==='') return;
          mapStock[String(f[0])] = Number(f[11])||0;
        });
      }
      if (filV >= 2) {
        hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
          if (f[0]==='' || String(f[4])!==numeroDoc || !f[13]) return;
          cliente = String(f[12]);
          lineas.push({
            productoId: String(f[13]),
            nombre: mapProd[String(f[13])] || String(f[13]),
            cantidadVendida: Number(f[16])||0,
            precioNeto: Number(f[17])||0,
            stockActual: mapStock[String(f[13])] || 0
          });
        });
      }
    } else {
      var hC = ss.getSheetByName('DB_COMPRAS');
      var filC = _ultimaFilaReal(hC);
      var hP2 = ss.getSheetByName('MSTR_PRODUCTOS');
      var filP2 = _ultimaFilaReal(hP2);
      var mapProd2 = {};
      if (filP2 >= 2) {
        hP2.getRange(2,1,filP2-1,28).getValues().forEach(function(f) {
          if (f[0]==='') return;
          mapProd2[String(f[0])] = String(f[3]||f[2]);
        });
      }
      if (filC >= 2) {
        hC.getRange(2,1,filC-1,30).getValues().forEach(function(f) {
          if (f[0]==='' || String(f[10])!==numeroDoc) return;
          cliente = String(f[3]);
          lineas.push({
            productoId: String(f[4]),
            nombre: mapProd2[String(f[4])] || String(f[4]),
            cantidadVendida: Number(f[11])||0,
            precioNeto: Number(f[13])||0
          });
        });
      }
    }

    return { lineas: lineas, idEntidad: cliente };

  } catch(e) {
    Logger.log('getLineasParaNC ERROR: ' + e.message);
    return { lineas: [], idEntidad: null };
  }
}
function verificarConflictoNC(numeroDoc, tipo) {
  try {
    var ss = _getSS();
    if (tipo === 'venta') {
      var hCxC = ss.getSheetByName('DB_CXC');
      var filCxC = _ultimaFilaReal(hCxC);
      var idCxc = null;
      if (filCxC >= 2) {
        hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
          if (String(f[4]) === numeroDoc) idCxc = String(f[0]);
        });
      }
      if (!idCxc) return { tieneConflicto: false };
      var info = _calcularSaldoCxC(idCxc);
      if (!info) return { tieneConflicto: false };
      var totalCobrado = info.montoOriginal - info.saldo;
      return { tieneConflicto: totalCobrado > 0, totalCobrado: totalCobrado, idCxc: idCxc, idEntidad: info.idCliente };
    } else {
      var hCxP = ss.getSheetByName('DB_CXP');
      var filCxP = _ultimaFilaReal(hCxP);
      var idCxp = null;
      if (filCxP >= 2) {
        hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f) {
          if (String(f[4]) === numeroDoc) idCxp = String(f[0]);
        });
      }
      if (!idCxp) return { tieneConflicto: false };
      var infoP = _calcularSaldoCxP(idCxp);
      if (!infoP) return { tieneConflicto: false };
      var totalPagado = infoP.montoOriginal - infoP.saldo;
      return { tieneConflicto: totalPagado > 0, totalCobrado: totalPagado, idCxp: idCxp, idEntidad: infoP.idProveedor };
    }
  } catch(e) {
    Logger.log('verificarConflictoNC ERROR: ' + e.message);
    return { tieneConflicto: false };
  }
}

function crearAnticipo(datos) {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('DB_ANTICIPOS');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();
    var idAnticipo = generarID('ANT');

    var fila = new Array(17).fill('');
    fila[0]  = idAnticipo;
    fila[1]  = ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[2]  = datos.idEntidad;
    fila[3]  = datos.tipoEntidad;
    fila[4]  = datos.origenId;
    fila[5]  = datos.monto;
    fila[6]  = 0;
    fila[7]  = ahora;
    fila[8]  = datos.motivo || 'Saldo a favor por NC';
    fila[9]  = 'DISPONIBLE';
    fila[10] = '';
    fila[11] = '';
    fila[12] = false;
    fila[13] = '';
    fila[14] = '';
    fila[15] = ahora;
    fila[16] = usuario;

    var primeraFilaVacia = _ultimaFilaReal(h) + 1;
    h.getRange(primeraFilaVacia, 1, 1, 17).setValues([fila]);
    SpreadsheetApp.flush();

    return { ok:true, idAnticipo: idAnticipo };
  } catch(e) {
    Logger.log('crearAnticipo ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function devolverDineroCliente(datos) {
  try {
    var ss = _getSS();
    var monto = Number(datos.monto);
    if (!monto || monto <= 0) return { ok:false, error:'Monto inválido' };

    var hE = ss.getSheetByName('DB_EGRESOS');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date(datos.fecha || new Date());
    var idEgreso = generarID('EGR');

    var fila = new Array(18).fill('');
    fila[0]  = idEgreso;
    fila[1]  = ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[2]  = ahora;
    fila[4]  = '2.1.1';
    fila[6]  = 'Devolución dinero a cliente — ' + (datos.motivo || '');
    fila[7]  = monto;
    fila[8]  = 0;
    fila[9]  = monto;
    fila[10] = 'CLP';
    fila[11] = 1;
    fila[13] = '';
    fila[14] = 'DEVOLUCION_CLIENTE';
    fila[16] = ahora;
    fila[17] = usuario;

    var primeraFilaVacia = _ultimaFilaReal(hE) + 1;
    hE.getRange(primeraFilaVacia, 1, 1, 18).setValues([fila]);
    SpreadsheetApp.flush();

    return { ok:true, idEgreso: idEgreso };
  } catch(e) {
    Logger.log('devolverDineroCliente ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function devolverDineroProveedor(datos) {
  try {
    var ss = _getSS();
    var monto = Number(datos.monto);
    if (!monto || monto <= 0) return { ok:false, error:'Monto inválido' };

    var hI = ss.getSheetByName('DB_INGRESOS');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date(datos.fecha || new Date());
    var idIngreso = generarID('ING');

    var fila = new Array(15).fill('');
    fila[0]  = idIngreso;
    fila[1]  = ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[2]  = ahora;
    fila[4]  = '1.1.1';
    fila[6]  = 'Devolución dinero de proveedor — ' + (datos.motivo || '');
    fila[7]  = monto;
    fila[8]  = 0;
    fila[9]  = monto;
    fila[11] = 'DEVOLUCION_PROVEEDOR';
    fila[13] = ahora;
    fila[14] = usuario;

    var primeraFilaVacia = _ultimaFilaReal(hI) + 1;
    hI.getRange(primeraFilaVacia, 1, 1, 15).setValues([fila]);
    SpreadsheetApp.flush();

    return { ok:true, idIngreso: idIngreso };
  } catch(e) {
    Logger.log('devolverDineroProveedor ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function getAnticipoDisponible(idEntidad, tipoEntidad) {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('DB_ANTICIPOS');
    var fil = _ultimaFilaReal(h);
    var total = 0;
    if (fil >= 2) {
      h.getRange(2,1,fil-1,17).getValues().forEach(function(f) {
        if (f[0]==='' || f[12]===true) return;
        if (String(f[2])!==idEntidad || String(f[3])!==tipoEntidad) return;
        if (String(f[9])==='DEVUELTO') return;
        total += (Number(f[5])||0) - (Number(f[6])||0);
      });
    }
    return Math.max(total, 0);
  } catch(e) {
    Logger.log('getAnticipoDisponible ERROR: ' + e.message);
    return 0;
  }
}

function aplicarAnticipoCxC(idCxc, monto) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var info = _calcularSaldoCxC(idCxc);
    if (!info) return { ok:false, error:'Documento no encontrado' };
    if (monto > info.saldo) return { ok:false, error:'Monto supera el saldo pendiente' };

    var disponible = getAnticipoDisponible(info.idCliente, 'CLIENTE');
    if (monto > disponible) return { ok:false, error:'Anticipo insuficiente. Disponible: $' + disponible.toLocaleString('es-CL') };

    var ss = _getSS();
    var h = ss.getSheetByName('DB_ANTICIPOS');
    var fil = _ultimaFilaReal(h);
    var restante = monto;
    var datos = h.getRange(2,1,fil-1,17).getValues();
    for (var i=0; i<datos.length && restante>0; i++) {
      if (datos[i][0]==='' || datos[i][12]===true) continue;
      if (String(datos[i][2])!==info.idCliente || String(datos[i][3])!=='CLIENTE') continue;
      if (String(datos[i][9])==='DEVUELTO') continue;
      var disp = (Number(datos[i][5])||0) - (Number(datos[i][6])||0);
      if (disp <= 0) continue;
      var usar = Math.min(disp, restante);
      h.getRange(i+2, 7).setValue((Number(datos[i][6])||0) + usar);
      if (usar >= disp) h.getRange(i+2, 10).setValue('AGOTADO');
      restante -= usar;
    }

    var hCob = ss.getSheetByName('DB_COBROS_CXC');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();
    var idCobro = generarID('COB');
    var fila = new Array(18).fill('');
    fila[0]=idCobro; fila[1]=Utilities.getUuid(); fila[2]=ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[3]=idCxc; fila[4]=info.idCliente; fila[5]=ahora; fila[6]=ahora; fila[7]=monto;
    fila[8]='ANTICIPO'; fila[9]=''; fila[10]=false; fila[11]=''; fila[12]=false; fila[13]='';
    fila[14]=''; fila[15]=''; fila[16]=ahora; fila[17]=usuario;
    var primeraFilaVacia = _ultimaFilaReal(hCob) + 1;
    hCob.getRange(primeraFilaVacia, 1, 1, 18).setValues([fila]);

    SpreadsheetApp.flush();
    actualizarSaldosVivosEnCalc();
    return { ok:true, idCobro: idCobro };
  } catch(e) {
    Logger.log('aplicarAnticipoCxC ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function aplicarAnticipoCxP(idCxp, monto) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var info = _calcularSaldoCxP(idCxp);
    if (!info) return { ok:false, error:'Documento no encontrado' };
    if (monto > info.saldo) return { ok:false, error:'Monto supera el saldo pendiente' };

    var disponible = getAnticipoDisponible(info.idProveedor, 'PROVEEDOR');
    if (monto > disponible) return { ok:false, error:'Anticipo insuficiente. Disponible: $' + disponible.toLocaleString('es-CL') };

    var ss = _getSS();
    var h = ss.getSheetByName('DB_ANTICIPOS');
    var fil = _ultimaFilaReal(h);
    var restante = monto;
    var datos = h.getRange(2,1,fil-1,17).getValues();
    for (var i=0; i<datos.length && restante>0; i++) {
      if (datos[i][0]==='' || datos[i][12]===true) continue;
      if (String(datos[i][2])!==info.idProveedor || String(datos[i][3])!=='PROVEEDOR') continue;
      if (String(datos[i][9])==='DEVUELTO') continue;
      var disp = (Number(datos[i][5])||0) - (Number(datos[i][6])||0);
      if (disp <= 0) continue;
      var usar = Math.min(disp, restante);
      h.getRange(i+2, 7).setValue((Number(datos[i][6])||0) + usar);
      if (usar >= disp) h.getRange(i+2, 10).setValue('AGOTADO');
      restante -= usar;
    }

    var hPag = ss.getSheetByName('DB_PAGOS_CXP');
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();
    var idPago = generarID('PAG');
    var fila = new Array(18).fill('');
    fila[0]=idPago; fila[1]=Utilities.getUuid(); fila[2]=ss.getSheetByName('CONFIG_SISTEMA').getRange('A2').getValue();
    fila[3]=idCxp; fila[4]=info.idProveedor; fila[5]=ahora; fila[6]=ahora; fila[7]=monto;
    fila[8]='ANTICIPO'; fila[9]=''; fila[10]=false; fila[11]=''; fila[12]=false; fila[13]='';
    fila[14]=''; fila[15]=''; fila[16]=ahora; fila[17]=usuario;
    var primeraFilaVacia = _ultimaFilaReal(hPag) + 1;
    hPag.getRange(primeraFilaVacia, 1, 1, 18).setValues([fila]);

    SpreadsheetApp.flush();
    actualizarSaldosVivosEnCalc();
    return { ok:true, idPago: idPago };
  } catch(e) {
    Logger.log('aplicarAnticipoCxP ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
function _validarPeriodoNoBloqueado(fecha) {
  try {
    var ss = _getSS();
    var hC = ss.getSheetByName('DB_CIERRES');
    var filC = hC.getLastRow();
    if (filC < 2) return { ok: true };

    var f = fecha instanceof Date ? fecha : new Date(fecha);
    var periodo = f.getFullYear() + '-' + String(f.getMonth()+1).padStart(2,'0');

    var cerrado = false;
    hC.getRange(2,1,filC-1,13).getValues().forEach(function(row) {
      if (row[0]==='') return;
      if (String(row[2]) === periodo && String(row[12]) === 'CERRADO') cerrado = true;
    });

    if (cerrado) {
      return { ok: false, error: 'El período ' + periodo + ' está cerrado. No se pueden registrar transacciones en este período.' };
    }
    return { ok: true };
  } catch(e) {
    Logger.log('_validarPeriodoNoBloqueado ERROR: ' + e.message);
    return { ok: true }; // fail-open para no bloquear el sistema por error de validación
  }
}

function listarCierres() {
  var ss = _getSS();
  var h = ss.getSheetByName('DB_CIERRES');
  var fil = h.getLastRow();
  var lista = [];
  if (fil >= 2) {
    h.getRange(2,1,fil-1,13).getValues().forEach(function(f) {
      if (f[0]==='') return;
      lista.push({
        idCierre: String(f[0]), periodo: String(f[2]),
        fechaCierre: f[3] instanceof Date ? f[3].getTime() : null,
        usuario: String(f[4]), ventasNetas: Number(f[5])||0,
        utilidadNeta: Number(f[8])||0, estado: String(f[12])
      });
    });
  }
  return lista.sort(function(a,b){ return b.periodo.localeCompare(a.periodo); });
}

function cerrarPeriodo(periodo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = _getSS();
    var hC = ss.getSheetByName('DB_CIERRES');
    var filC = hC.getLastRow();

    if (filC >= 2) {
      var yaExiste = hC.getRange(2,1,filC-1,13).getValues().some(function(f) {
        return String(f[2]) === periodo && String(f[12]) === 'CERRADO';
      });
      if (yaExiste) return { ok:false, error:'Este período ya está cerrado' };
    }

    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var anulados = _getDocumentosAnulados();

    var ventasNetas = 0, costoVenta = 0;
    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (per !== periodo) return;
        ventasNetas += Number(f[19])||0;
        costoVenta += (Number(f[16])||0) * (Number(f[22])||0);
      });
    }

    var gastosOperacionales = 0;
    var hE = ss.getSheetByName('DB_EGRESOS');
    var filE = _ultimaFilaReal(hE);
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        if (anulados.compras[String(f[15]||'')]) return;
        if (String(f[4])==='1.1.4') return; // excluir mercadería, no es gasto operacional
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var per = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (per !== periodo) return;
        gastosOperacionales += Number(f[7])||0;
      });
    }

    var utilidadNeta = ventasNetas - costoVenta - gastosOperacionales;
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var ahora = new Date();

    var fila = [
      generarID('CIE'), empresa, periodo, ahora, usuario,
      ventasNetas, costoVenta, gastosOperacionales, utilidadNeta,
      0, 0, 0, 'CERRADO'
    ];
    var primeraFilaVacia = hC.getLastRow() + 1;
    hC.getRange(primeraFilaVacia, 1, 1, 13).setValues([fila]);
    SpreadsheetApp.flush();

    return { ok:true, periodo: periodo, utilidadNeta: utilidadNeta };

  } catch(e) {
    Logger.log('cerrarPeriodo ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function reabrirPeriodo(periodo, clave) {
  try {
    var ss = _getSS();
    var claveReal = ss.getSheetByName('CONFIG_SISTEMA').getRange('AA2').getValue();
    if (String(clave) !== String(claveReal)) {
      return { ok:false, error:'Clave de administrador incorrecta' };
    }

    var hC = ss.getSheetByName('DB_CIERRES');
    var filC = hC.getLastRow();
    var encontrado = false;
    if (filC >= 2) {
      var datos = hC.getRange(2,1,filC-1,13).getValues();
      for (var i=0; i<datos.length; i++) {
        if (String(datos[i][2])===periodo && String(datos[i][12])==='CERRADO') {
          hC.getRange(i+2, 13).setValue('REABIERTO');
          encontrado = true;
        }
      }
    }
    if (!encontrado) return { ok:false, error:'No se encontró un cierre activo para este período' };

    SpreadsheetApp.flush();
    return { ok:true };

  } catch(e) {
    Logger.log('reabrirPeriodo ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function getEERRData(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;
    var anulados = _getDocumentosAnulados();

    var hPC = ss.getSheetByName('MSTR_PLAN_CUENTAS');
    var filPC = hPC.getLastRow();
    var mapSubtipo = {};
    hPC.getRange(2,1,filPC-1,17).getValues().forEach(function(f) {
      if (f[0]==='') return;
      mapSubtipo[String(f[0])] = String(f[12]||'NO_APLICA');
    });

    var ventasNetas = 0, costoVenta = 0;
    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        ventasNetas += Number(f[19])||0;
        costoVenta += (Number(f[16])||0) * (Number(f[22])||0);
      });
    }

    var gastosOperacionales = 0, gastosFinancieros = 0, otrosGastos = 0;
    var hE = ss.getSheetByName('DB_EGRESOS');
    var filE = _ultimaFilaReal(hE);
    var cuentasSinClasificar = {};
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        if (anulados.compras[String(f[15]||'')]) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var cuenta = String(f[4]);
        var subtipo = mapSubtipo[cuenta] || 'NO_APLICA';
        var monto = Number(f[7])||0;
        if (subtipo === 'GASTO_OPERACIONAL') gastosOperacionales += monto;
        else if (subtipo === 'GASTO_FINANCIERO') gastosFinancieros += monto;
        else if (subtipo === 'OTRO_GASTO') otrosGastos += monto;
        else if (subtipo === 'NO_APLICA') { /* activo, no es gasto */ }
        else cuentasSinClasificar[cuenta] = (cuentasSinClasificar[cuenta]||0) + monto;
      });
    }

    var utilidadBruta = ventasNetas - costoVenta;
    var resultadoOperacional = utilidadBruta - gastosOperacionales;
    var otrosResultados = -gastosFinancieros - otrosGastos;
    var utilidadNeta = resultadoOperacional + otrosResultados;

    var margenBrutoPct = ventasNetas > 0 ? utilidadBruta/ventasNetas : 0;
    var margenOperacionalPct = ventasNetas > 0 ? resultadoOperacional/ventasNetas : 0;
    var margenNetoPct = ventasNetas > 0 ? utilidadNeta/ventasNetas : 0;

    return {
      periodo: per,
      ventasNetas: ventasNetas,
      costoVenta: costoVenta,
      utilidadBruta: utilidadBruta,
      gastosOperacionales: gastosOperacionales,
      resultadoOperacional: resultadoOperacional,
      gastosFinancieros: gastosFinancieros,
      otrosGastos: otrosGastos,
      utilidadNeta: utilidadNeta,
      margenBrutoPct: margenBrutoPct,
      margenOperacionalPct: margenOperacionalPct,
      margenNetoPct: margenNetoPct,
      alertaCuentasSinClasificar: Object.keys(cuentasSinClasificar).length > 0 ? cuentasSinClasificar : null
    };

  } catch(e) {
    Logger.log('getEERRData ERROR: ' + e.message);
    return null;
  }
}
function getFlujoEfectivoHistorico(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;
    var anulados = _getDocumentosAnulados();

    var ventasContado = 0, cobrosCxC = 0, anticiposRecibidos = 0;
    var comprasContado = 0, pagosCxP = 0, devolucionesEntregadas = 0, devolucionesRecibidas = 0;

    var idsConCxC = {};
    var hCxC = ss.getSheetByName('DB_CXC');
    var filCxC = _ultimaFilaReal(hCxC);
    if (filCxC >= 2) {
      hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
        if (f[0]==='') return;
        idsConCxC[String(f[17])] = true;
      });
    }

    var hI = ss.getSheetByName('DB_INGRESOS');
    var filI = _ultimaFilaReal(hI);
    if (filI >= 2) {
      hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idOrigen = String(f[12]||'');
        var tipoDoc = String(f[11]||'');
        if (anulados.ventas[idOrigen]) return;
        if (tipoDoc === 'DEVOLUCION_PROVEEDOR') { devolucionesRecibidas += Number(f[7])||0; return; }
        if (!idsConCxC[idOrigen]) ventasContado += Number(f[7])||0;
      });
    }

    var idsConCxP = {};
    var hCxP = ss.getSheetByName('DB_CXP');
    var filCxP = _ultimaFilaReal(hCxP);
    if (filCxP >= 2) {
      hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f) {
        if (f[0]==='') return;
        idsConCxP[String(f[4])] = true;
      });
    }

    var hE = ss.getSheetByName('DB_EGRESOS');
    var filE = _ultimaFilaReal(hE);
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idOrigen = String(f[15]||'');
        var tipoDoc = String(f[14]||'');
        if (anulados.compras[idOrigen]) return;
        if (tipoDoc === 'DEVOLUCION_CLIENTE') { devolucionesEntregadas += Number(f[7])||0; return; }
        if (!idsConCxP[idOrigen]) comprasContado += Number(f[7])||0;
      });
    }

    var hCob = ss.getSheetByName('DB_COBROS_CXC');
    var filCob = _ultimaFilaReal(hCob);
    if (filCob >= 2) {
      hCob.getRange(2,1,filCob-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || f[12]===true) return;
        var fecha = f[5];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        if (String(f[8])==='ANTICIPO') { anticiposRecibidos += Number(f[7])||0; return; }
        cobrosCxC += Number(f[7])||0;
      });
    }

    var hPag = ss.getSheetByName('DB_PAGOS_CXP');
    var filPag = _ultimaFilaReal(hPag);
    if (filPag >= 2) {
      hPag.getRange(2,1,filPag-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || f[12]===true) return;
        var fecha = f[5];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        if (String(f[8])==='ANTICIPO') return;
        pagosCxP += Number(f[7])||0;
      });
    }

    var totalEntradas = ventasContado + cobrosCxC + anticiposRecibidos + devolucionesRecibidas;
    var totalSalidas = comprasContado + pagosCxP + devolucionesEntregadas;
    var flujoNeto = totalEntradas - totalSalidas;

    return {
      periodo: per,
      ventasContado: ventasContado,
      cobrosCxC: cobrosCxC,
      anticiposRecibidos: anticiposRecibidos,
      devolucionesRecibidas: devolucionesRecibidas,
      totalEntradas: totalEntradas,
      comprasContado: comprasContado,
      pagosCxP: pagosCxP,
      devolucionesEntregadas: devolucionesEntregadas,
      totalSalidas: totalSalidas,
      flujoNeto: flujoNeto
    };

  } catch(e) {
    Logger.log('getFlujoEfectivoHistorico ERROR: ' + e.message);
    return null;
  }
}
function getIndicadoresGerenciales(periodo) {
  try {
    var cobranza = getCobranzaData();
    var proveedores = getProveedoresData();
    var inventario = getInventarioData();
    var eerr = getEERRData(periodo);

    if (!cobranza || !proveedores || !inventario || !eerr) {
      return null;
    }

    var dso = cobranza.dso;
    var dpo = proveedores.dpo;
    var diasInventario = inventario.diasRotacion;
    var ccc = diasInventario + dso - dpo;

    return {
      periodo: eerr.periodo,
      dso: dso,
      dpo: dpo,
      diasInventario: diasInventario,
      ccc: ccc,
      ebitdaAproximado: eerr.resultadoOperacional,
      margenBrutoPct: eerr.margenBrutoPct,
      margenOperacionalPct: eerr.margenOperacionalPct,
      margenNetoPct: eerr.margenNetoPct,
      ventasNetas: eerr.ventasNetas
    };

  } catch(e) {
    Logger.log('getIndicadoresGerenciales ERROR: ' + e.message);
    return null;
  }
}
function getFlujoCajaMultiMes(meses) {
  try {
    var n = meses || 6;
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var anulados = _getDocumentosAnulados();

    var hoy = new Date();
    var periodos = [];
    for (var i = n-1; i >= 0; i--) {
      var fecha = new Date(hoy.getFullYear(), hoy.getMonth()-i, 1);
      periodos.push(fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0'));
    }

    var mapa = {};
    periodos.forEach(function(p) {
      mapa[p] = { periodo:p, ventasContado:0, cobrosCxC:0, anticiposRecibidos:0, devolucionesRecibidas:0,
        totalEntradas:0, comprasContado:0, pagosCxP:0, devolucionesEntregadas:0, totalSalidas:0, flujoNeto:0 };
    });

    var idsConCxC = {};
    var hCxC = ss.getSheetByName('DB_CXC');
    var filCxC = _ultimaFilaReal(hCxC);
    if (filCxC >= 2) {
      hCxC.getRange(2,1,filCxC-1,21).getValues().forEach(function(f) {
        if (f[0]==='') return;
        idsConCxC[String(f[17])] = true;
      });
    }

    var idsConCxP = {};
    var hCxP = ss.getSheetByName('DB_CXP');
    var filCxP = _ultimaFilaReal(hCxP);
    if (filCxP >= 2) {
      hCxP.getRange(2,1,filCxP-1,20).getValues().forEach(function(f) {
        if (f[0]==='') return;
        idsConCxP[String(f[4])] = true;
      });
    }

    var hI = ss.getSheetByName('DB_INGRESOS');
    var filI = _ultimaFilaReal(hI);
    if (filI >= 2) {
      hI.getRange(2,1,filI-1,15).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (!mapa[p]) return;
        var idOrigen = String(f[12]||'');
        var tipoDoc = String(f[11]||'');
        if (anulados.ventas[idOrigen]) return;
        if (tipoDoc === 'DEVOLUCION_PROVEEDOR') { mapa[p].devolucionesRecibidas += Number(f[7])||0; return; }
        if (!idsConCxC[idOrigen]) mapa[p].ventasContado += Number(f[7])||0;
      });
    }

    var hE = ss.getSheetByName('DB_EGRESOS');
    var filE = _ultimaFilaReal(hE);
    if (filE >= 2) {
      hE.getRange(2,1,filE-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[1])!==empresa) return;
        var fecha = f[2];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (!mapa[p]) return;
        var idOrigen = String(f[15]||'');
        var tipoDoc = String(f[14]||'');
        if (anulados.compras[idOrigen]) return;
        if (tipoDoc === 'DEVOLUCION_CLIENTE') { mapa[p].devolucionesEntregadas += Number(f[7])||0; return; }
        if (!idsConCxP[idOrigen]) mapa[p].comprasContado += Number(f[7])||0;
      });
    }

    var hCob = ss.getSheetByName('DB_COBROS_CXC');
    var filCob = _ultimaFilaReal(hCob);
    if (filCob >= 2) {
      hCob.getRange(2,1,filCob-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || f[12]===true) return;
        var fecha = f[5];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (!mapa[p]) return;
        if (String(f[8])==='ANTICIPO') { mapa[p].anticiposRecibidos += Number(f[7])||0; return; }
        mapa[p].cobrosCxC += Number(f[7])||0;
      });
    }

    var hPag = ss.getSheetByName('DB_PAGOS_CXP');
    var filPag = _ultimaFilaReal(hPag);
    if (filPag >= 2) {
      hPag.getRange(2,1,filPag-1,18).getValues().forEach(function(f) {
        if (f[0]==='' || f[12]===true) return;
        var fecha = f[5];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (!mapa[p]) return;
        if (String(f[8])==='ANTICIPO') return;
        mapa[p].pagosCxP += Number(f[7])||0;
      });
    }

    periodos.forEach(function(p) {
      var m = mapa[p];
      m.totalEntradas = m.ventasContado + m.cobrosCxC + m.anticiposRecibidos + m.devolucionesRecibidas;
      m.totalSalidas = m.comprasContado + m.pagosCxP + m.devolucionesEntregadas;
      m.flujoNeto = m.totalEntradas - m.totalSalidas;
    });

    return periodos.map(function(p){ return mapa[p]; });

  } catch(e) {
    Logger.log('getFlujoCajaMultiMes ERROR: ' + e.message);
    return [];
  }
}

function getRentabilidadClientes(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;

    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var porCliente = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idCli = String(f[12]||'');
        if (idCli === 'CLI-GENERICO') return;
        if (!porCliente[idCli]) porCliente[idCli] = { ventas:0, margen:0, docs:{} };
        porCliente[idCli].ventas += Number(f[19])||0;
        porCliente[idCli].margen += Number(f[23])||0;
        porCliente[idCli].docs[String(f[1])] = true;
      });
    }

    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var mapCli = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapCli[String(f[0])] = String(f[5]||f[0]);
      });
    }

    var lista = Object.keys(porCliente).map(function(id) {
      var c = porCliente[id];
      return {
        id: id,
        nombre: mapCli[id] || id.substring(0,14)+'...',
        ventas: c.ventas,
        margen: c.margen,
        margenPct: c.ventas > 0 ? c.margen/c.ventas : 0,
        documentos: Object.keys(c.docs).length
      };
    }).sort(function(a,b){ return b.margen - a.margen; });

    return { periodo: per, clientes: lista };

  } catch(e) {
    Logger.log('getRentabilidadClientes ERROR: ' + e.message);
    return null;
  }
}
function getRentabilidadProductos(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;

    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var porProducto = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idProd = String(f[13]||'');
        if (!idProd) return;
        if (!porProducto[idProd]) porProducto[idProd] = { ventas:0, margen:0, cantidad:0 };
        porProducto[idProd].ventas += Number(f[19])||0;
        porProducto[idProd].margen += Number(f[23])||0;
        porProducto[idProd].cantidad += Number(f[16])||0;
      });
    }

    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var mapProd = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapProd[String(f[0])] = String(f[3]||f[2]||f[0]);
      });
    }

    var lista = Object.keys(porProducto).map(function(id) {
      var p = porProducto[id];
      return {
        id: id,
        nombre: mapProd[id] || id,
        ventas: p.ventas,
        margen: p.margen,
        margenPct: p.ventas > 0 ? p.margen/p.ventas : 0,
        cantidad: p.cantidad
      };
    }).sort(function(a,b){ return b.margen - a.margen; });

    return { periodo: per, productos: lista };

  } catch(e) {
    Logger.log('getRentabilidadProductos ERROR: ' + e.message);
    return null;
  }
}
function getRentabilidadClientes(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;

    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var porCliente = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idCli = String(f[12]||'');
        if (idCli === 'CLI-GENERICO') return;
        if (!porCliente[idCli]) porCliente[idCli] = { ventas:0, margen:0, docs:{} };
        porCliente[idCli].ventas += Number(f[19])||0;
        porCliente[idCli].margen += Number(f[23])||0;
        porCliente[idCli].docs[String(f[1])] = true;
      });
    }

    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var mapCli = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,15).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapCli[String(f[0])] = String(f[5]||f[0]);
      });
    }

    var lista = Object.keys(porCliente).map(function(id) {
      var c = porCliente[id];
      return {
        id: id,
        nombre: mapCli[id] || id.substring(0,14)+'...',
        ventas: c.ventas,
        margen: c.margen,
        margenPct: c.ventas > 0 ? c.margen/c.ventas : 0,
        documentos: Object.keys(c.docs).length
      };
    }).sort(function(a,b){ return b.margen - a.margen; });

    return { periodo: per, clientes: lista };

  } catch(e) {
    Logger.log('getRentabilidadClientes ERROR: ' + e.message);
    return null;
  }
}

function getRentabilidadProductos(periodo) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;
    var per = periodo || ctx.periodo;

    var hV = ss.getSheetByName('DB_VENTAS');
    var filV = _ultimaFilaReal(hV);
    var porProducto = {};

    if (filV >= 2) {
      hV.getRange(2,1,filV-1,34).getValues().forEach(function(f) {
        if (f[0]==='' || String(f[2])!==empresa) return;
        if (String(f[27])==='ANULADO') return;
        var fecha = f[10];
        if (!(fecha instanceof Date)) return;
        var p = fecha.getFullYear()+'-'+String(fecha.getMonth()+1).padStart(2,'0');
        if (p !== per) return;
        var idProd = String(f[13]||'');
        if (!idProd) return;
        if (!porProducto[idProd]) porProducto[idProd] = { ventas:0, margen:0, cantidad:0 };
        porProducto[idProd].ventas += Number(f[19])||0;
        porProducto[idProd].margen += Number(f[23])||0;
        porProducto[idProd].cantidad += Number(f[16])||0;
      });
    }

    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var mapProd = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapProd[String(f[0])] = String(f[3]||f[2]||f[0]);
      });
    }

    var lista = Object.keys(porProducto).map(function(id) {
      var p = porProducto[id];
      return {
        id: id,
        nombre: mapProd[id] || id,
        ventas: p.ventas,
        margen: p.margen,
        margenPct: p.ventas > 0 ? p.margen/p.ventas : 0,
        cantidad: p.cantidad
      };
    }).sort(function(a,b){ return b.margen - a.margen; });

    return { periodo: per, productos: lista };

  } catch(e) {
    Logger.log('getRentabilidadProductos ERROR: ' + e.message);
    return null;
  }
}
function getConsolidadoFinanciero(periodo) {
  try {
    var eerr = getEERRData(periodo);
    var finanzas = getFinanzasData();
    var cobranza = getCobranzaData();
    var inventario = getInventarioData();
    var rentClientes = getRentabilidadClientes(periodo);
    var rentProductos = getRentabilidadProductos(periodo);

    if (!eerr || !finanzas) return null;

    var sinMovimiento = [];
    var valorInmovilizado = 0;
    var valorTotalInventario = inventario ? inventario.valorCosto : 0;
    if (inventario && inventario.productos) {
      inventario.productos.forEach(function(p) {
        if (p.diasSinVenta !== null && p.diasSinVenta !== undefined && p.diasSinVenta >= 90) {
          sinMovimiento.push({ nombre: p.nombre, diasSinVenta: p.diasSinVenta, valor: p.valor });
          valorInmovilizado += p.valor;
        }
      });
    }
    sinMovimiento.sort(function(a,b){ return b.valor - a.valor; });
    var pctInmovilizado = valorTotalInventario > 0 ? valorInmovilizado/valorTotalInventario : 0;

    var alertas = [];
    if (finanzas.flujoProyectado < 0) {
      alertas.push({ nivel:'critico', mensaje: 'Flujo de caja proyectado a 30 días es negativo: ' + fmtMoneyServer(finanzas.flujoProyectado) });
    }
    if (cobranza && cobranza.concentracionMax > 0.5) {
      alertas.push({ nivel:'advertencia', mensaje: 'Un cliente concentra ' + Math.round(cobranza.concentracionMax*100) + '% de la cartera por cobrar' });
    }
    if (eerr.margenBrutoPct < 0.15) {
      alertas.push({ nivel:'critico', mensaje: 'Margen bruto del período es bajo: ' + (eerr.margenBrutoPct*100).toFixed(1) + '%' });
    }
    if (pctInmovilizado > 0.2) {
      alertas.push({ nivel:'advertencia', mensaje: Math.round(pctInmovilizado*100) + '% del valor de inventario está inmovilizado (sin ventas en 90+ días)' });
    }
    if (cobranza && cobranza.cxcVencida > 0 && cobranza.cxcTotal > 0 && (cobranza.cxcVencida/cobranza.cxcTotal) > 0.15) {
      alertas.push({ nivel:'critico', mensaje: 'Cartera vencida representa ' + Math.round((cobranza.cxcVencida/cobranza.cxcTotal)*100) + '% del total por cobrar' });
    }

    return {
      periodo: eerr.periodo,
      resumenEjecutivo: {
        ventasTotales: eerr.ventasNetas,
        costoVentas: eerr.costoVenta,
        margenBruto: eerr.utilidadBruta,
        gastosOperacionales: eerr.gastosOperacionales,
        utilidadOperacional: eerr.resultadoOperacional,
        utilidadNeta: eerr.utilidadNeta,
        cajaDisponible: finanzas.cajaReal,
        cuentasPorCobrar: finanzas.cxcTotal,
        cuentasPorPagar: finanzas.cxpTotal
      },
      eerr: eerr,
      flujoCaja: {
        cajaActual: finanzas.cajaPeriodo,
        cxc30d: finanzas.cxc030,
        cxpExigible: finanzas.cxpPresion,
        flujoProyectado: finanzas.flujoProyectado
      },
      rentClientes: rentClientes ? rentClientes.clientes.slice(0,5) : [],
      rentProductos: rentProductos ? rentProductos.productos.slice(0,5) : [],
      inventarioCritico: {
        sinMovimiento: sinMovimiento.slice(0,10),
        stockInmovilizadoValorizado: valorInmovilizado,
        pctDelInventarioTotal: pctInmovilizado
      },
      alertas: alertas
    };

  } catch(e) {
    Logger.log('getConsolidadoFinanciero ERROR: ' + e.message);
    return null;
  }
}
function procesarCargaMasivaClientesCSV(csvTexto) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var lineas = csvTexto.split(/\r?\n/).filter(function(l){ return l.trim() !== ''; });
    if (lineas.length < 2) return { ok:false, error:'El archivo no tiene datos (solo encabezado o vacío)' };

    function parsearLineaCSV(linea) {
      var resultado = [];
      var actual = '';
      var dentroComillas = false;
      for (var i=0; i<linea.length; i++) {
        var c = linea[i];
        if (c === '"') {
          if (dentroComillas && linea[i+1] === '"') { actual += '"'; i++; }
          else dentroComillas = !dentroComillas;
        } else if (c === ';' && !dentroComillas) {
          resultado.push(actual); actual = '';
        } else {
          actual += c;
        }
      }
      resultado.push(actual);
      return resultado;
    }

    var filas = lineas.slice(1).map(parsearLineaCSV);

    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var rutsExistentes = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,5).getValues().forEach(function(f) {
        if (f[0]==='') return;
        rutsExistentes[String(f[3])] = String(f[0]);
      });
    }

    var creados = 0, omitidos = 0, errores = [];
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var fechaAhora = new Date();
    var filasNuevas = [];

    filas.forEach(function(f, idx) {
      if (!f[0] || String(f[0]).trim()==='') return;
      var razonSocial = String(f[0]).trim();
      var rut = String(f[1]||'').trim();

      if (!rut) { errores.push('Fila '+(idx+2)+': sin RUT, omitida'); return; }
      if (!validarRUT(rut)) { errores.push('Fila '+(idx+2)+' ('+razonSocial+'): RUT inválido, omitida'); return; }

      var rutNorm = rut.replace(/[.\-]/g,'').toUpperCase();
      rutNorm = rutNorm.substring(0, rutNorm.length-1);

      if (rutsExistentes[rutNorm]) { omitidos++; return; }

      var idNuevo = generarID('CLI');
      var fila = new Array(23).fill('');
      fila[0]=idNuevo; fila[1]=empresa; fila[2]='EMPRESA'; fila[3]=rutNorm; fila[4]=rut;
      fila[5]=razonSocial; fila[6]=''; fila[7]=String(f[2]||''); fila[8]=String(f[3]||'');
      fila[9]=String(f[4]||''); fila[10]=String(f[5]||''); fila[11]=String(f[6]||'');
      fila[14]=String(f[7]||'CONTADO'); fila[15]=Number(f[8])||0; fila[17]=true;
      fila[19]=fechaAhora; fila[20]=usuario;
      filasNuevas.push(fila);
      rutsExistentes[rutNorm] = idNuevo;
      creados++;
    });

    if (filasNuevas.length) {
      var primeraFilaVacia = _ultimaFilaReal(hC) + 1;
      hC.getRange(primeraFilaVacia, 1, filasNuevas.length, 23).setValues(filasNuevas);
      SpreadsheetApp.flush();
    }

    return { ok:true, creados: creados, omitidos: omitidos, errores: errores };

  } catch(e) {
    Logger.log('procesarCargaMasivaClientesCSV ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function procesarCargaMasivaProveedoresCSV(csvTexto) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var lineas = csvTexto.split(/\r?\n/).filter(function(l){ return l.trim() !== ''; });
    if (lineas.length < 2) return { ok:false, error:'El archivo no tiene datos (solo encabezado o vacío)' };

    function parsearLineaCSV(linea) {
      var resultado = [];
      var actual = '';
      var dentroComillas = false;
      for (var i=0; i<linea.length; i++) {
        var c = linea[i];
        if (c === '"') {
          if (dentroComillas && linea[i+1] === '"') { actual += '"'; i++; }
          else dentroComillas = !dentroComillas;
        } else if (c === ';' && !dentroComillas) {
          resultado.push(actual); actual = '';
        } else {
          actual += c;
        }
      }
      resultado.push(actual);
      return resultado;
    }

    var filas = lineas.slice(1).map(parsearLineaCSV);

    var hP = ss.getSheetByName('MSTR_PROVEEDORES');
    var filP = _ultimaFilaReal(hP);
    var rutsExistentes = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,5).getValues().forEach(function(f) {
        if (f[0]==='') return;
        rutsExistentes[String(f[3])] = String(f[0]);
      });
    }

    var creados = 0, omitidos = 0, errores = [];
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var fechaAhora = new Date();
    var filasNuevas = [];

    filas.forEach(function(f, idx) {
      if (!f[0] || String(f[0]).trim()==='') return;
      var razonSocial = String(f[0]).trim();
      var rut = String(f[1]||'').trim();

      if (!rut) { errores.push('Fila '+(idx+2)+': sin RUT, omitida'); return; }
      if (!validarRUT(rut)) { errores.push('Fila '+(idx+2)+' ('+razonSocial+'): RUT inválido, omitida'); return; }

      var rutNorm = rut.replace(/[.\-]/g,'').toUpperCase();
      rutNorm = rutNorm.substring(0, rutNorm.length-1);

      if (rutsExistentes[rutNorm]) { omitidos++; return; }

      var idNuevo = generarID('PRV');
      var fila = new Array(24).fill('');
      fila[0]=idNuevo; fila[1]=empresa; fila[2]='EMPRESA'; fila[3]=rutNorm; fila[4]=rut;
      fila[5]=razonSocial; fila[6]=''; fila[7]=String(f[2]||''); fila[8]=String(f[3]||'');
      fila[9]=String(f[4]||''); fila[10]=String(f[5]||'CONTADO'); fila[11]=Number(f[6])||0;
      fila[12]='CLP'; fila[13]=false; fila[18]=true;
      fila[20]=fechaAhora; fila[21]=usuario;
      filasNuevas.push(fila);
      rutsExistentes[rutNorm] = idNuevo;
      creados++;
    });

    if (filasNuevas.length) {
      var primeraFilaVacia = _ultimaFilaReal(hP) + 1;
      hP.getRange(primeraFilaVacia, 1, filasNuevas.length, 24).setValues(filasNuevas);
      SpreadsheetApp.flush();
    }

    return { ok:true, creados: creados, omitidos: omitidos, errores: errores };

  } catch(e) {
    Logger.log('procesarCargaMasivaProveedoresCSV ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function descargarPlantillaProveedoresCSV() {
  var headers = ['razon_social','rut','giro','email','telefono','condicion_pago','plazo_entrega_dias'];
  var ejemplos = [
    ['Proveedor Ejemplo SpA','76.111.222-3','Fabricación e importación','contacto@proveedor.cl','+56911112222','CONTADO',0],
    ['Insumos del Sur Limitada','77.333.444-5','Venta de insumos','ventas@insumos.cl','+56933334444','30_DIAS',7]
  ];
  var filas = [headers].concat(ejemplos);
  var csv = filas.map(function(fila) {
    return fila.map(function(celda) {
      var val = String(celda);
      if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1) {
        val = '"' + val.replace(/"/g,'""') + '"';
      }
      return val;
    }).join(';');
  }).join('\n');
  return csv;
}

function descargarPlantillaSaldosInventarioCSV() {
  var headers = ['sku','cantidad','costo_unitario'];
  var ejemplos = [
    ['PRD-001',50,1000],
    ['PRD-002',0,0]
  ];
  var filas = [headers].concat(ejemplos);
  var csv = filas.map(function(fila) {
    return fila.map(function(celda) {
      var val = String(celda);
      if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1) {
        val = '"' + val.replace(/"/g,'""') + '"';
      }
      return val;
    }).join(';');
  }).join('\n');
  return csv;
}

function descargarPlantillaSaldosCxCCSV() {
  var headers = ['rut_cliente','numero_documento','fecha_emision','fecha_vencimiento','monto'];
  var ejemplos = [
    ['76.123.456-7','FAC-0001','2026-05-01','2026-05-31',150000],
    ['77.987.654-3','FAC-0002','2026-06-10','2026-07-10',320000]
  ];
  var filas = [headers].concat(ejemplos);
  var csv = filas.map(function(fila) {
    return fila.map(function(celda) {
      var val = String(celda);
      if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1) {
        val = '"' + val.replace(/"/g,'""') + '"';
      }
      return val;
    }).join(';');
  }).join('\n');
  return csv;
}

function descargarPlantillaSaldosCxPCSV() {
  var headers = ['rut_proveedor','numero_documento','fecha_emision','fecha_vencimiento','monto'];
  var ejemplos = [
    ['76.111.222-3','FAC-9001','2026-05-05','2026-06-04',280000],
    ['77.333.444-5','FAC-9002','2026-06-15','2026-07-15',95000]
  ];
  var filas = [headers].concat(ejemplos);
  var csv = filas.map(function(fila) {
    return fila.map(function(celda) {
      var val = String(celda);
      if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1) {
        val = '"' + val.replace(/"/g,'""') + '"';
      }
      return val;
    }).join(';');
  }).join('\n');
  return csv;
}

function procesarCargaSaldosInventarioCSV(csvTexto) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var lineas = csvTexto.split(/\r?\n/).filter(function(l){ return l.trim() !== ''; });
    if (lineas.length < 2) return { ok:false, error:'El archivo no tiene datos (solo encabezado o vacío)' };

    function parsearLineaCSV(linea) {
      var resultado = [];
      var actual = '';
      var dentroComillas = false;
      for (var i=0; i<linea.length; i++) {
        var c = linea[i];
        if (c === '"') {
          if (dentroComillas && linea[i+1] === '"') { actual += '"'; i++; }
          else dentroComillas = !dentroComillas;
        } else if (c === ';' && !dentroComillas) {
          resultado.push(actual); actual = '';
        } else {
          actual += c;
        }
      }
      resultado.push(actual);
      return resultado;
    }

    var filas = lineas.slice(1).map(parsearLineaCSV);

    var hP = ss.getSheetByName('MSTR_PRODUCTOS');
    var filP = _ultimaFilaReal(hP);
    var mapaSku = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,28).getValues().forEach(function(f, idx) {
        if (f[0]==='') return;
        mapaSku[String(f[2])] = { fila: idx+2, stockActual: Number(f[11])||0 };
      });
    }

    var procesados = 0, errores = [];
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var fechaAhora = new Date();
    var hMov = ss.getSheetByName('DB_MOVIMIENTOS');

    filas.forEach(function(f, idx) {
      if (!f[0] || String(f[0]).trim()==='') return;
      var sku = String(f[0]).trim();
      var cantidad = Number(f[1])||0;
      var costo = Number(f[2])||0;

      var prod = mapaSku[sku];
      if (!prod) { errores.push('Fila '+(idx+2)+': SKU "'+sku+'" no existe en el catálogo'); return; }
      if (cantidad <= 0) { errores.push('Fila '+(idx+2)+': cantidad inválida'); return; }

      hP.getRange(prod.fila, 12).setValue(prod.stockActual + cantidad);
      hP.getRange(prod.fila, 8).setValue(costo);

      var filaMov = _ultimaFilaReal(hMov) + 1;
      hMov.getRange(filaMov, 1, 1, 12).setValues([[
        generarID('MOV'), empresa, fechaAhora, ctx.periodo, sku,
        '', '', 'APERTURA', cantidad, cantidad, 'CARGA_INICIAL', usuario
      ]]);

      procesados++;
    });

    SpreadsheetApp.flush();
    return { ok:true, procesados: procesados, errores: errores };

  } catch(e) {
    Logger.log('procesarCargaSaldosInventarioCSV ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function procesarCargaSaldosCxCCSV(csvTexto) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var lineas = csvTexto.split(/\r?\n/).filter(function(l){ return l.trim() !== ''; });
    if (lineas.length < 2) return { ok:false, error:'El archivo no tiene datos (solo encabezado o vacío)' };

    function parsearLineaCSV(linea) {
      var resultado = [];
      var actual = '';
      var dentroComillas = false;
      for (var i=0; i<linea.length; i++) {
        var c = linea[i];
        if (c === '"') {
          if (dentroComillas && linea[i+1] === '"') { actual += '"'; i++; }
          else dentroComillas = !dentroComillas;
        } else if (c === ';' && !dentroComillas) {
          resultado.push(actual); actual = '';
        } else {
          actual += c;
        }
      }
      resultado.push(actual);
      return resultado;
    }

    var filas = lineas.slice(1).map(parsearLineaCSV);

    var hC = ss.getSheetByName('MSTR_CLIENTES');
    var filC = _ultimaFilaReal(hC);
    var mapaRut = {};
    if (filC >= 2) {
      hC.getRange(2,1,filC-1,4).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapaRut[String(f[3])] = String(f[0]);
      });
    }

    var procesados = 0, errores = [];
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var hCxC = ss.getSheetByName('DB_CXC');

    filas.forEach(function(f, idx) {
      if (!f[0] || String(f[0]).trim()==='') return;
      var rut = String(f[0]).trim().replace(/[.\-]/g,'').toUpperCase();
      rut = rut.substring(0, rut.length-1);
      var numeroDoc = String(f[1]||'').trim();
      var monto = Number(f[4])||0;

      var idCliente = mapaRut[rut];
      if (!idCliente) { errores.push('Fila '+(idx+2)+': cliente con RUT "'+f[0]+'" no existe'); return; }
      if (!numeroDoc) { errores.push('Fila '+(idx+2)+': falta número de documento'); return; }
      if (monto <= 0) { errores.push('Fila '+(idx+2)+': monto inválido'); return; }

      var fechaEm = new Date(f[2]);
      var fechaVc = new Date(f[3]);
      var diasPlazo = Math.round((fechaVc - fechaEm)/(1000*60*60*24));

      var filaCxC = _ultimaFilaReal(hCxC) + 1;
      hCxC.getRange(filaCxC, 1, 1, 21).setValues([[
        generarID('CXC'), empresa, idCliente, '', numeroDoc, 'SALDO_INICIAL',
        fechaEm, fechaVc, monto, 0, monto, diasPlazo, '0_30_DIAS',
        'PENDIENTE', '', '', '', '', '', new Date(), usuario
      ]]);
      procesados++;
    });

    SpreadsheetApp.flush();
    return { ok:true, procesados: procesados, errores: errores };

  } catch(e) {
    Logger.log('procesarCargaSaldosCxCCSV ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function procesarCargaSaldosCxPCSV(csvTexto) {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var empresa = ctx.idEmpresa;

    var lineas = csvTexto.split(/\r?\n/).filter(function(l){ return l.trim() !== ''; });
    if (lineas.length < 2) return { ok:false, error:'El archivo no tiene datos (solo encabezado o vacío)' };

    function parsearLineaCSV(linea) {
      var resultado = [];
      var actual = '';
      var dentroComillas = false;
      for (var i=0; i<linea.length; i++) {
        var c = linea[i];
        if (c === '"') {
          if (dentroComillas && linea[i+1] === '"') { actual += '"'; i++; }
          else dentroComillas = !dentroComillas;
        } else if (c === ';' && !dentroComillas) {
          resultado.push(actual); actual = '';
        } else {
          actual += c;
        }
      }
      resultado.push(actual);
      return resultado;
    }

    var filas = lineas.slice(1).map(parsearLineaCSV);

    var hP = ss.getSheetByName('MSTR_PROVEEDORES');
    var filP = _ultimaFilaReal(hP);
    var mapaRut = {};
    if (filP >= 2) {
      hP.getRange(2,1,filP-1,4).getValues().forEach(function(f) {
        if (f[0]==='') return;
        mapaRut[String(f[3])] = String(f[0]);
      });
    }

    var procesados = 0, errores = [];
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var hCxP = ss.getSheetByName('DB_CXP');

    filas.forEach(function(f, idx) {
      if (!f[0] || String(f[0]).trim()==='') return;
      var rut = String(f[0]).trim().replace(/[.\-]/g,'').toUpperCase();
      rut = rut.substring(0, rut.length-1);
      var numeroDoc = String(f[1]||'').trim();
      var monto = Number(f[4])||0;

      var idProveedor = mapaRut[rut];
      if (!idProveedor) { errores.push('Fila '+(idx+2)+': proveedor con RUT "'+f[0]+'" no existe'); return; }
      if (!numeroDoc) { errores.push('Fila '+(idx+2)+': falta número de documento'); return; }
      if (monto <= 0) { errores.push('Fila '+(idx+2)+': monto inválido'); return; }

      var fechaEm = new Date(f[2]);
      var fechaVc = new Date(f[3]);
      var diasPlazo = Math.round((fechaVc - fechaEm)/(1000*60*60*24));

      var filaCxP = _ultimaFilaReal(hCxP) + 1;
      hCxP.getRange(filaCxP, 1, 1, 20).setValues([[
        generarID('CXP'), empresa, idProveedor, '', numeroDoc, 'SALDO_INICIAL',
        fechaEm, fechaVc, monto, 0, 'CLP', 1, monto, diasPlazo, '0_30_DIAS',
        'PENDIENTE', '', '', new Date(), usuario
      ]]);
      procesados++;
    });

    SpreadsheetApp.flush();
    return { ok:true, procesados: procesados, errores: errores };

  } catch(e) {
    Logger.log('procesarCargaSaldosCxPCSV ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}
function descargarPlantillaClientesCSV() {
  var headers = ['razon_social','rut','giro','email','telefono','direccion','comuna','condicion_pago','limite_credito'];
  var ejemplos = [
    ['Comercial Ejemplo Limitada','76.123.456-7','Venta al por mayor','contacto@ejemplo.cl','+56912345678','Av. Siempre Viva 123','Santiago','CONTADO',0],
    ['Distribuidora Modelo SpA','77.987.654-3','Distribución productos','ventas@modelo.cl','+56987654321','Calle Falsa 456','Providencia','30_DIAS',500000]
  ];
  var filas = [headers].concat(ejemplos);
  var csv = filas.map(function(fila) {
    return fila.map(function(celda) {
      var val = String(celda);
      if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1) {
        val = '"' + val.replace(/"/g,'""') + '"';
      }
      return val;
    }).join(';');
  }).join('\n');
  return csv;
}
// ============================================================
// MÓDULO CONFIGURACIÓN
// ============================================================

function getConfigEmpresa() {
  try {
    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    var v = cfg.getRange(2,1,1,cfg.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = h.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      razonSocial:    String(val('nombre_empresa')||''),
      nombreFantasia: String(val('nombre_fantasia')||''),
      rut:            String(val('rut_empresa')||''),
      giro:           String(val('giro_empresa')||''),
      direccion:      String(val('direccion')||''),
      ciudad:         String(val('comuna')||''),
      region:         String(val('region')||''),
      telefono:       String(val('telefono')||''),
      email:          String(val('email_admin')||''),
      logoUrl:        String(val('logo_url')||''),
      moneda:         String(val('moneda_base')||'CLP'),
      zonaHoraria:    String(val('zona_horaria')||'America/Santiago')
    };
  } catch(e) {
    Logger.log('getConfigEmpresa ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigEmpresa(datos) {
  try {
    if (!datos.razonSocial || String(datos.razonSocial).trim()==='') {
      return { ok:false, error:'La razón social es obligatoria' };
    }
    if (!datos.rut || String(datos.rut).trim()==='') {
      return { ok:false, error:'El RUT es obligatorio' };
    }
    if (!validarRUT(datos.rut)) {
      return { ok:false, error:'El RUT ingresado no es válido' };
    }

    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = h.indexOf(nombreCol);
      if (idx>=0) cfg.getRange(2, idx+1).setValue(valor);
    }

    setVal('nombre_empresa', datos.razonSocial);
    setVal('nombre_fantasia', datos.nombreFantasia||'');
    setVal('rut_empresa', datos.rut);
    setVal('giro_empresa', datos.giro||'');
    setVal('direccion', datos.direccion||'');
    setVal('comuna', datos.ciudad||'');
    setVal('region', datos.region||'');
    setVal('telefono', datos.telefono||'');
    setVal('email_admin', datos.email||'');
    setVal('logo_url', datos.logoUrl||'');
    setVal('moneda_base', datos.moneda||'CLP');
    setVal('zona_horaria', datos.zonaHoraria||'America/Santiago');

    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('guardarConfigEmpresa ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigTributario() {
  try {
    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    var v = cfg.getRange(2,1,1,cfg.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = h.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      tasaIva:          Number(val('tasa_iva'))||0.19,
      retencionBH:      Number(val('retencion_bh_pct'))||0.1525,
      moneda:           String(val('moneda_base')||'CLP'),
      decimales:        Number(val('decimales'))||0,
      formatoFecha:     String(val('formato_fecha')||'dd-MM-yyyy')
    };
  } catch(e) {
    Logger.log('getConfigTributario ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigTributario(datos) {
  try {
    var iva = Number(datos.tasaIva);
    if (isNaN(iva) || iva < 0 || iva > 1) {
      return { ok:false, error:'El IVA debe ser un valor entre 0 y 1 (ej: 0.19 para 19%)' };
    }
    var ret = Number(datos.retencionBH);
    if (isNaN(ret) || ret < 0 || ret > 1) {
      return { ok:false, error:'La retención debe ser un valor entre 0 y 1' };
    }

    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = h.indexOf(nombreCol);
      if (idx>=0) cfg.getRange(2, idx+1).setValue(valor);
    }

    setVal('tasa_iva', iva);
    setVal('retencion_bh_pct', ret);
    setVal('retencion_bh_actualizado', new Date());
    setVal('decimales', Number(datos.decimales)||0);
    setVal('formato_fecha', datos.formatoFecha||'dd-MM-yyyy');

    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('guardarConfigTributario ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigComercial() {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_COMERCIAL');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    var v = h.getRange(2,1,1,h.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = headers.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      metaMensualDefault:        Number(val('meta_mensual_default'))||0,
      metaAnual:                 Number(val('meta_anual'))||0,
      ticketObjetivo:            Number(val('ticket_objetivo'))||0,
      margenBrutoObjetivo:       Number(val('margen_bruto_objetivo'))||0.25,
      crecimientoMensualEsperado:Number(val('crecimiento_mensual_esperado'))||0.05,
      ventaMinimaDiaria:         Number(val('venta_minima_diaria'))||0,
      objetivoClientesNuevos:    Number(val('objetivo_clientes_nuevos'))||0,
      diasCreditoDefault:        Number(val('dias_credito_default'))||30,
      limiteCreditoDefault:      Number(val('limite_credito_default'))||0,
      descuentoMaxPct:           Number(val('descuento_max_pct'))||0,
      listaPrecioDefault:        String(val('lista_precio_default')||'GENERAL'),
      bloquearVentaSinStock:     val('bloquear_venta_sin_stock')!==false,
      autorizarDescuentoAlto:    val('autorizar_descuento_alto')===true,
      permitirModificarPrecio:   val('permitir_modificar_precio')!==false,
      mostrarCostoProductos:     val('mostrar_costo_productos')!==false
    };
  } catch(e) {
    Logger.log('getConfigComercial ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigComercial(datos) {
  try {
    var descMax = Number(datos.descuentoMaxPct);
    if (isNaN(descMax) || descMax < 0 || descMax > 100) {
      return { ok:false, error:'El descuento máximo debe ser un valor entre 0 y 100' };
    }

    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_COMERCIAL');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = headers.indexOf(nombreCol);
      if (idx>=0) h.getRange(2, idx+1).setValue(valor);
    }

    setVal('meta_mensual_default', Number(datos.metaMensualDefault)||0);
    setVal('meta_anual', Number(datos.metaAnual)||0);
    setVal('ticket_objetivo', Number(datos.ticketObjetivo)||0);
    setVal('margen_bruto_objetivo', Number(datos.margenBrutoObjetivo)||0.25);
    setVal('crecimiento_mensual_esperado', Number(datos.crecimientoMensualEsperado)||0.05);
    setVal('venta_minima_diaria', Number(datos.ventaMinimaDiaria)||0);
    setVal('objetivo_clientes_nuevos', Number(datos.objetivoClientesNuevos)||0);
    setVal('dias_credito_default', Number(datos.diasCreditoDefault)||30);
    setVal('limite_credito_default', Number(datos.limiteCreditoDefault)||0);
    setVal('descuento_max_pct', descMax);
    setVal('lista_precio_default', datos.listaPrecioDefault||'GENERAL');
    setVal('bloquear_venta_sin_stock', datos.bloquearVentaSinStock===true);
    setVal('autorizar_descuento_alto', datos.autorizarDescuentoAlto===true);
    setVal('permitir_modificar_precio', datos.permitirModificarPrecio===true);
    setVal('mostrar_costo_productos', datos.mostrarCostoProductos===true);

    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('guardarConfigComercial ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigInventario() {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_INVENTARIO');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    var v = h.getRange(2,1,1,h.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = headers.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      stockCriticoDefault:   Number(val('stock_critico_default'))||5,
      politicaCosteo:        String(val('politica_costeo')||'CPP'),
      permiteStockNegativo:  val('permite_stock_negativo')===true,
      diasReposicionDefault: Number(val('dias_reposicion_default'))||7
    };
  } catch(e) {
    Logger.log('getConfigInventario ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigInventario(datos) {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_INVENTARIO');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = headers.indexOf(nombreCol);
      if (idx>=0) h.getRange(2, idx+1).setValue(valor);
    }

    setVal('stock_critico_default', Number(datos.stockCriticoDefault)||5);
    setVal('permite_stock_negativo', datos.permiteStockNegativo===true);
    setVal('dias_reposicion_default', Number(datos.diasReposicionDefault)||7);
    // politica_costeo NO se modifica desde UI en esta fase (FIFO deshabilitado)

    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('guardarConfigInventario ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigMediosPago() {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var lista = [];
    if (fil >= 2) {
      h.getRange(2,1,fil-1,6).getValues().forEach(function(f, idx) {
        if (f[0]==='') return;
        lista.push({
          id: String(f[0]),
          nombre: String(f[2]),
          activo: f[3]===true,
          orden: Number(f[4])||0,
          esDefault: f[5]===true,
          fila: idx+2
        });
      });
    }
    return lista.sort(function(a,b){ return a.orden - b.orden; });
  } catch(e) {
    Logger.log('getConfigMediosPago ERROR: ' + e.message);
    return [];
  }
}

function crearMedioPago(nombre) {
  try {
    if (!nombre || String(nombre).trim()==='') {
      return { ok:false, error:'El nombre es obligatorio' };
    }
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var maxOrden = 0;
    if (fil >= 2) {
      h.getRange(2,5,fil-1,1).getValues().forEach(function(f) {
        var o = Number(f[0])||0;
        if (o > maxOrden) maxOrden = o;
      });
    }
    var idNuevo = 'MP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    var primeraFilaVacia = fil + 1;
    h.getRange(primeraFilaVacia, 1, 1, 6).setValues([[
      idNuevo, 'EMP-01', String(nombre).trim(), true, maxOrden+1, false
    ]]);
    SpreadsheetApp.flush();
    return { ok:true, id: idNuevo };
  } catch(e) {
    Logger.log('crearMedioPago ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function editarMedioPago(id, nombre) {
  try {
    if (!nombre || String(nombre).trim()==='') {
      return { ok:false, error:'El nombre es obligatorio' };
    }
    if (id === 'ANTICIPO') {
      return { ok:false, error:'El medio de pago Anticipo es interno del sistema y no puede editarse' };
    }
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var datos = h.getRange(2,1,fil-1,6).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === id) {
        h.getRange(i+2, 3).setValue(String(nombre).trim());
        SpreadsheetApp.flush();
        return { ok:true };
      }
    }
    return { ok:false, error:'Medio de pago no encontrado' };
  } catch(e) {
    Logger.log('editarMedioPago ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function toggleActivoMedioPago(id) {
  try {
    if (id === 'ANTICIPO') {
      return { ok:false, error:'El medio de pago Anticipo es interno del sistema y no puede desactivarse' };
    }
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var datos = h.getRange(2,1,fil-1,6).getValues();
    for (var i=0; i<datos.length; i++) {
      if (String(datos[i][0]) === id) {
        var estadoActual = datos[i][3] === true;
        h.getRange(i+2, 4).setValue(!estadoActual);
        SpreadsheetApp.flush();
        return { ok:true, activo: !estadoActual };
      }
    }
    return { ok:false, error:'Medio de pago no encontrado' };
  } catch(e) {
    Logger.log('toggleActivoMedioPago ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function marcarMedioPagoDefault(id) {
  try {
    if (id === 'ANTICIPO') {
      return { ok:false, error:'El medio de pago Anticipo no puede ser el medio por defecto' };
    }
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var datos = h.getRange(2,1,fil-1,6).getValues();
    for (var i=0; i<datos.length; i++) {
      h.getRange(i+2, 6).setValue(String(datos[i][0]) === id);
    }
    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('marcarMedioPagoDefault ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function reordenarMediosPago(idsEnOrden) {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('MSTR_MEDIOS_PAGO');
    var fil = _ultimaFilaReal(h);
    var datos = h.getRange(2,1,fil-1,6).getValues();
    idsEnOrden.forEach(function(id, idx) {
      for (var i=0; i<datos.length; i++) {
        if (String(datos[i][0]) === id) {
          h.getRange(i+2, 5).setValue(idx+1);
          break;
        }
      }
    });
    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('reordenarMediosPago ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigSeguridad() {
  try {
    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    var v = cfg.getRange(2,1,1,cfg.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = h.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      requiereConfirmacionAnulacion: val('requiere_confirmacion_anulacion')===true
    };
  } catch(e) {
    Logger.log('getConfigSeguridad ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigSeguridad(datos) {
  try {
    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = h.indexOf(nombreCol);
      if (idx>=0) cfg.getRange(2, idx+1).setValue(valor);
    }
    setVal('requiere_confirmacion_anulacion', datos.requiereConfirmacionAnulacion===true);
    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('guardarConfigSeguridad ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function cambiarClaveAdmin(claveActual, claveNueva) {
  try {
    if (!claveNueva || String(claveNueva).trim().length < 4) {
      return { ok:false, error:'La nueva clave debe tener al menos 4 caracteres' };
    }
    var ss = _getSS();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    var idx = h.indexOf('clave_admin_cierre');
    var claveReal = cfg.getRange(2, idx+1).getValue();

    if (String(claveActual) !== String(claveReal)) {
      return { ok:false, error:'La clave actual es incorrecta' };
    }

    cfg.getRange(2, idx+1).setValue(String(claveNueva));
    SpreadsheetApp.flush();
    return { ok:true };
  } catch(e) {
    Logger.log('cambiarClaveAdmin ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigAlertas() {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_ALERTAS');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    var v = h.getRange(2,1,1,h.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = headers.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    return {
      liquidezMinima:     Number(val('liquidez_minima'))||1.5,
      margenMinimo:       Number(val('margen_minimo'))||0.15,
      stockCriticoUmbral: Number(val('stock_critico_umbral'))||5,
      cxcVencidaMax:      Number(val('cxc_vencida_max'))||500000,
      cxpUrgenteMax:      Number(val('cxp_urgente_max'))||500000,
      scoreMinimo:        Number(val('score_minimo'))||60
    };
  } catch(e) {
    Logger.log('getConfigAlertas ERROR: ' + e.message);
    return null;
  }
}

function guardarConfigAlertas(datos) {
  try {
    var ss = _getSS();
    var h = ss.getSheetByName('CONFIG_ALERTAS');
    var headers = h.getRange(1,1,1,h.getLastColumn()).getValues()[0];
    function setVal(nombreCol, valor) {
      var idx = headers.indexOf(nombreCol);
      if (idx>=0) h.getRange(2, idx+1).setValue(valor);
    }

    setVal('liquidez_minima', Number(datos.liquidezMinima)||1.5);
    setVal('margen_minimo', Number(datos.margenMinimo)||0.15);
    setVal('stock_critico_umbral', Number(datos.stockCriticoUmbral)||5);
    setVal('cxc_vencida_max', Number(datos.cxcVencidaMax)||500000);
    setVal('cxp_urgente_max', Number(datos.cxpUrgenteMax)||500000);
    setVal('score_minimo', Number(datos.scoreMinimo)||60);

    SpreadsheetApp.flush();
    return { ok:true, nota:'Estos umbrales aún no afectan los semáforos del Dashboard — esa integración se hará en una fase posterior' };
  } catch(e) {
    Logger.log('guardarConfigAlertas ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigRespaldos() {
  try {
    var ss = _getSS();
    var archivo = DriveApp.getFileById(ss.getId());
    return {
      ultimaActualizacion: archivo.getLastUpdated().getTime()
    };
  } catch(e) {
    Logger.log('getConfigRespaldos ERROR: ' + e.message);
    return { ultimaActualizacion: null };
  }
}

function exportarRespaldoCompleto() {
  try {
    var ss = _getSS();
    var nombreCopia = 'VANTIS_Respaldo_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var copia = ss.copy(nombreCopia);
    return { ok:true, url: copia.getUrl(), nombre: nombreCopia };
  } catch(e) {
    Logger.log('exportarRespaldoCompleto ERROR: ' + e.message);
    return { ok:false, error:'Error del sistema: ' + e.message };
  }
}

function getConfigSistemaInfo() {
  try {
    var ss = _getSS();
    var ctx = getDatosContexto();
    var cfg = ss.getSheetByName('CONFIG_SISTEMA');
    var h = cfg.getRange(1,1,1,cfg.getLastColumn()).getValues()[0];
    var v = cfg.getRange(2,1,1,cfg.getLastColumn()).getValues()[0];
    function val(nombreCol) { var idx = h.indexOf(nombreCol); return idx>=0 ? v[idx] : ''; }

    var archivo = DriveApp.getFileById(ss.getId());
    var alertas = getAlertasData();

    return {
      version: String(val('version_esquema')||'1.0.0'),
      ultimaActualizacion: archivo.getLastUpdated().getTime(),
      empresaActiva: ctx.nombreEmpresa,
      periodoActivo: ctx.periodo,
      estadoSistema: alertas && alertas.batchFallido ? 'Atención' : 'OK',
      baseDatos: 'Google Sheets'
    };
  } catch(e) {
    Logger.log('getConfigSistemaInfo ERROR: ' + e.message);
    return null;
  }
}