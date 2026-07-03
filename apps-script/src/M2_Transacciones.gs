/**
 * ============================================================
 * M2_Transacciones.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Motor transaccional del sistema.
 * Procesa ventas, compras, movimientos de stock, CxC y CxP.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Módulos:      M0_Setup, M3_Maestros, M4_Auditoria
 *   Hojas:        DB_VENTAS, DB_COMPRAS, DB_MOVIMIENTOS,
 *                 DB_CXC, DB_CXP, MSTR_PRODUCTOS,
 *                 MSTR_TIPOS_CAMBIO
 *   APIs Google:  LockService, SpreadsheetApp
 *   Triggers:     onEdit (instalado por M0)
 * ============================================================
 */

// ============================================================
// FUNCIÓN: onEdit
// Trigger principal de transacciones
// Se ejecuta en cada edición del archivo
// ============================================================

/**
 * Detecta la hoja editada y ejecuta el proceso correspondiente.
 * Solo procesa filas de datos (fila >= 2) con campos obligatorios.
 *
 * @param {Object} e - Evento de edición de Google Sheets
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var hoja   = e.range.getSheet();
    var nombre = hoja.getName();
    var fila   = e.range.getRow();
    var col    = e.range.getColumn();

    // Solo procesar filas de datos — no encabezados
    if (fila < 2) return;

    Logger.log('onEdit: Hoja=' + nombre + ' Fila=' + fila + ' Col=' + col);

    switch (nombre) {
      case HOJAS.DB_VENTAS:
        procesarVenta(e);
        break;
      case HOJAS.DB_COMPRAS:
        _detectarAccionCompra(e);
        break;
      case HOJAS.DB_CXC:
        _detectarPagoCxC(e);
        break;
      case HOJAS.DB_CXP:
        _detectarPagoCxP(e);
        break;
      default:
        // Hoja no transaccional — ignorar
        break;
    }

  } catch (e2) {
    Logger.log('ERROR onEdit: ' + e2.message);
  }
}

// ============================================================
// FUNCIÓN: procesarVenta
// Registra y valida una línea de venta en DB_VENTAS
// ============================================================

/**
 * Procesa una línea de venta cuando el usuario completa los datos.
 * Valida período, cliente, producto y stock.
 * Actualiza stock y genera CxC si corresponde.
 *
 * Los campos obligatorios para disparar el proceso son:
 * - fecha_venta (col K)
 * - id_cliente (col M)
 * - id_producto (col N)
 * - cantidad (col Q)
 * - precio_unitario_neto (col R)
 *
 * @param {Object} e - Evento de edición
 */
function procesarVenta(e) {
  try {
    var hoja = e.range.getSheet();
    var fila = e.range.getRow();

    // Leer toda la fila en una sola llamada
    var datos = hoja.getRange(fila, 1, 1, 34).getValues()[0];

    // Verificar que los campos obligatorios están completos
    var fechaVenta   = datos[10]; // K — fecha_venta
    var idCliente    = String(datos[12] || '').trim(); // M
    var idProducto   = String(datos[13] || '').trim(); // N
    var cantidad     = Number(datos[16]); // Q
    var precioUnit   = Number(datos[17]); // R
    var idLinea      = String(datos[0] || '').trim();  // A — id_linea_venta

    // Si ya tiene ID significa que ya fue procesada
    if (idLinea !== '') return;

    // Verificar campos mínimos para procesar
    if (!fechaVenta || !idCliente || !idProducto || cantidad <= 0 || precioUnit <= 0) {
      return; // Fila incompleta — esperar
    }

    // VALIDACIÓN 1: Período abierto
    if (!validarPeriodoAbierto(true)) {
      SpreadsheetApp.getUi().alert(
        '⛔ Sin período activo\n\n' +
        'No existe un período contable ABIERTO.\n' +
        'Ir a ERP → Administración → Gestionar períodos.'
      );
      return;
    }

    // VALIDACIÓN 2: Cliente existe y está activo
    var validCliente = _validarCliente(idCliente);
    if (!validCliente.valido) {
      SpreadsheetApp.getUi().alert('❌ Cliente inválido\n\n' + validCliente.motivo);
      return;
    }

    // VALIDACIÓN 3: Producto existe y está activo
    var validProducto = _validarProducto(idProducto);
    if (!validProducto.valido) {
      SpreadsheetApp.getUi().alert('❌ Producto inválido\n\n' + validProducto.motivo);
      return;
    }

    // VALIDACIÓN 4: Stock suficiente (solo si maneja stock)
    var descuento  = Number(datos[18]) || 0; // S — descuento_pct
    var manejaStock = validProducto.manejaStock;

    if (manejaStock) {
      var validStock = _validarStock(idProducto, cantidad);
      if (!validStock.valido) {
        SpreadsheetApp.getUi().alert('❌ Stock insuficiente\n\n' + validStock.motivo);
        return;
      }
    }

    // LOCK — evitar concurrencia
    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      // Obtener datos adicionales
      var config       = getConfig();
      var idEmpresa    = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
      var periodo      = getPeriodoActual();
      var ahora        = new Date();
      var usuario      = Session.getActiveUser().getEmail() || 'sistema';
      var costoMomento = _capturarCostoMomento(idProducto);
      var tipoDoc      = String(datos[3] || 'BOLETA').trim(); // D
      var numDoc       = String(datos[4] || '').trim();       // E
      idLinea = generarID('VTA');

      // Escribir campos del sistema en la fila
      hoja.getRange(fila, 1).setValue(idLinea);          // A id_linea_venta
      hoja.getRange(fila, 3).setValue(idEmpresa);        // C id_empresa
      hoja.getRange(fila, 23).setValue(costoMomento);    // W costo_unitario_momento
      hoja.getRange(fila, 28).setValue(ESTADOS.VIGENTE); // AB estado
      hoja.getRange(fila, 31).setValue(ahora);           // AE creado_en
      hoja.getRange(fila, 32).setValue(usuario);         // AF creado_por

      SpreadsheetApp.flush();

      // Actualizar stock si el producto lo maneja
      if (manejaStock) {
        var idBodega = String(datos[9] || '').trim(); // J id_bodega
        if (!idBodega) {
          var bodegaDef = validarBodegaDefecto();
          idBodega = bodegaDef.valida ? bodegaDef.idBodega : '';
        }
        actualizarStock(
          idProducto,
          cantidad,
          ESTADOS.MOV_SALIDA_VENTA,
          idBodega,
          idLinea,
          'VENTA'
        );
      }

      // Crear CxC si no es contado
      var condicionPago = validCliente.condicionPago;
      if (condicionPago && condicionPago !== 'CONTADO') {
        var subtotalNeto  = cantidad * precioUnit * (1 - descuento);
        var tasa          = config ? Number(config.tasa_iva) : 0.19;
        var aplIva        = validProducto.aplicaIva;
        var iva           = aplIva ? Math.round(subtotalNeto * tasa) : 0;
        var totalBruto    = subtotalNeto + iva;
        var idDocumento   = String(datos[1] || idLinea).trim(); // B id_documento

        crearCxCDesdeVenta(
          idDocumento,
          idCliente,
          Math.round(totalBruto),
          condicionPago,
          new Date(fechaVenta)
        );
      }

      lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    // Registrar en auditoría
    registrarLog(HOJAS.DB_VENTAS, 'CREAR', idLinea, 'id_producto', idProducto);
    Logger.log('procesarVenta: ' + idLinea + ' procesada — Producto: ' + idProducto + ' Cant: ' + cantidad);

  } catch (e) {
    Logger.log('ERROR procesarVenta: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: actualizarStock
// Actualiza stock_actual en MSTR_PRODUCTOS y registra movimiento
// ============================================================

/**
 * Actualiza el stock de un producto y registra el movimiento
 * en DB_MOVIMIENTOS. Usa LockService para evitar concurrencia.
 *
 * @param {string} idProducto    - ID del producto
 * @param {number} cantidad      - Cantidad del movimiento
 * @param {string} tipoMovimiento - Tipo según ESTADOS.MOV_*
 * @param {string} idBodega      - Bodega origen/destino
 * @param {string} idReferencia  - ID del documento origen
 * @param {string} tipoRef       - Tipo de referencia: 'VENTA', 'COMPRA', 'AJUSTE'
 * @returns {boolean} true si se actualizó correctamente
 */
function actualizarStock(idProducto, cantidad, tipoMovimiento, idBodega, idReferencia, tipoRef) {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var hojaProd    = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    var hojaMovs    = ss.getSheetByName(HOJAS.DB_MOVIMIENTOS);

    if (!hojaProd || !hojaMovs) {
      Logger.log('ERROR actualizarStock: Hojas no encontradas');
      return false;
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      // Buscar la fila del producto
      var filaProd = _buscarFilaPorId(hojaProd, idProducto, COL.PRD_ID_PRODUCTO);
      if (filaProd === -1) {
        lock.releaseLock();
        Logger.log('ERROR actualizarStock: Producto no encontrado: ' + idProducto);
        return false;
      }

      // Leer stock actual
      var stockActual = Number(
        hojaProd.getRange(filaProd, COL.PRD_STOCK_ACTUAL + 1).getValue()
      );
      var costoUnit   = Number(
        hojaProd.getRange(filaProd, COL.PRD_COSTO_UNITARIO + 1).getValue()
      );

      // Calcular impacto en stock
      var impacto     = _calcularImpactoStock(tipoMovimiento, cantidad);
      var nuevoStock  = stockActual + impacto;

      // Verificar que no quede negativo
      if (nuevoStock < 0) {
        lock.releaseLock();
        Logger.log('ERROR actualizarStock: Stock resultante negativo: ' + nuevoStock);
        return false;
      }

      // Actualizar stock en MSTR_PRODUCTOS
      hojaProd.getRange(filaProd, COL.PRD_STOCK_ACTUAL + 1).setValue(nuevoStock);

      // Registrar movimiento en DB_MOVIMIENTOS
      var config    = getConfig();
      var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
      var ahora     = new Date();
      var periodo   = getPeriodoActual() || '';
      var usuario   = Session.getActiveUser().getEmail() || 'sistema-automatico';
      var idMov     = generarID('MOV');

      var filaMov = new Array(16).fill('');
      filaMov[0]  = idMov;           // A id_movimiento
      filaMov[1]  = idEmpresa;       // B id_empresa
      filaMov[2]  = ahora;           // C fecha_movimiento
      filaMov[3]  = '';        // D sys_periodo — ARRAYFORMULA
      filaMov[4]  = idProducto;      // E id_producto
      filaMov[5]  = idBodega;        // F id_bodega_origen
      filaMov[6]  = idBodega;        // G id_bodega_destino
      filaMov[7]  = tipoMovimiento;  // H tipo_movimiento
      filaMov[8]  = cantidad;        // I cantidad
      filaMov[9]  = '';        // J sys_impacto_stock — ARRAYFORMULA
      filaMov[10] = costoUnit;       // K costo_unitario_momento
      filaMov[11] = idReferencia;    // L id_referencia
      filaMov[12] = tipoRef;         // M tipo_referencia
      filaMov[13] = '';              // N motivo_ajuste
      filaMov[14] = ahora;           // O creado_en
      filaMov[15] = usuario;         // P creado_por

      _escribirFilaSegura(hojaMovs, filaMov, HOJAS.DB_MOVIMIENTOS);
lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.MSTR_PRODUCTOS, 'STOCK', idProducto, 'stock_actual',
      tipoMovimiento + ':' + impacto);
    Logger.log('actualizarStock: ' + idProducto + ' | ' + tipoMovimiento +
      ' | Cantidad: ' + cantidad + ' | Impacto: ' + impacto);
    return true;

  } catch (e) {
    Logger.log('ERROR actualizarStock: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: crearCxCDesdeVenta
// Crea automáticamente el registro en DB_CXC
// ============================================================

/**
 * Crea un registro en DB_CXC cuando se confirma una venta a crédito.
 * Solo se ejecuta si la condición de pago no es CONTADO.
 *
 * @param {string} idDocumento    - ID del documento de venta
 * @param {string} idCliente      - ID del cliente
 * @param {number} montoTotal     - Monto total con IVA
 * @param {string} condicionPago  - Condición de pago del cliente
 * @param {Date}   fechaVenta     - Fecha de la venta
 */
function crearCxCDesdeVenta(idDocumento, idCliente, montoTotal, condicionPago, fechaVenta) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_CXC);

    if (!hoja) {
      Logger.log('ERROR crearCxCDesdeVenta: DB_CXC no encontrada');
      return false;
    }

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var usuario   = Session.getActiveUser().getEmail() || 'sistema';
    var idCxC     = generarID('CXC');

    // Calcular fecha de vencimiento
    var fechaVenc = _calcularFechaVencimiento(fechaVenta || ahora, condicionPago);

    var fila = new Array(21).fill('');
    fila[0]  = idCxC;              // A id_cxc
    fila[1]  = idEmpresa;          // B id_empresa
    fila[2]  = idCliente;          // C id_cliente
    fila[3]  = '';                 // D id_cc
    fila[4]  = idDocumento;        // E id_documento
    fila[5]  = 'VENTA';            // F tipo_documento
    fila[6]  = fechaVenta || ahora;// G fecha_emision
    fila[7]  = fechaVenc;          // H fecha_vencimiento
    fila[8]  = Math.round(montoTotal); // I monto_original
    fila[9]  = 0;                  // J monto_pagado
    fila[10] = ''; // K calc_saldo_pendiente — ARRAYFORMULA no escribir
    fila[11] = ''; // L calc_dias_vencimiento — ARRAYFORMULA no escribir
    fila[12] = ''; // M calc_tramo_aging — ARRAYFORMULA no escribir
    fila[13] = ESTADOS.PENDIENTE;  // N estado
    fila[14] = '';                 // O fecha_ultimo_pago
    fila[15] = '';                 // P gestionado_por
    fila[16] = '';                 // Q url_documento
    fila[17] = idDocumento;        // R id_referencia
    fila[18] = '';                 // S notas
    fila[19] = ahora;              // T creado_en
    fila[20] = usuario;            // U creado_por

    _escribirFilaSegura(hoja, fila, HOJAS.DB_CXC);

registrarLog(HOJAS.DB_CXC, 'CREAR', idCxC, 'id_cliente', idCliente);
    Logger.log('crearCxCDesdeVenta: ' + idCxC + ' creada — Cliente: ' +
      idCliente + ' Monto: ' + montoTotal);
    return true;

  } catch (e) {
    Logger.log('ERROR crearCxCDesdeVenta: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _validarCliente
// ============================================================

/**
 * Verifica que un cliente existe y está activo.
 *
 * @param {string} idCliente - ID del cliente a validar
 * @returns {Object} { valido, motivo, condicionPago }
 */
function _validarCliente(idCliente) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_CLIENTES);

    if (!hoja) {
      return { valido: false, motivo: 'Hoja MSTR_CLIENTES no encontrada', condicionPago: '' };
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { valido: false, motivo: 'No hay clientes registrados', condicionPago: '' };
    }

    var datos = hoja.getRange(2, 1, ultimaFila - 1, 18).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var id     = String(fila[COL_CLI.ID]).trim();
      var activo = fila[COL_CLI.ACTIVO];

      if (id === idCliente) {
        if (!activo) {
          return { valido: false, motivo: 'Cliente inactivo: ' + idCliente, condicionPago: '' };
        }
        return {
          valido:        true,
          motivo:        'OK',
          condicionPago: String(fila[COL_CLI.CONDICION_PAGO] || 'CONTADO')
        };
      }
    }

    return { valido: false, motivo: 'Cliente no encontrado: ' + idCliente, condicionPago: '' };

  } catch (e) {
    Logger.log('ERROR _validarCliente: ' + e.message);
    return { valido: false, motivo: 'Error: ' + e.message, condicionPago: '' };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _validarProducto
// ============================================================

/**
 * Verifica que un producto existe y está activo.
 *
 * @param {string} idProducto - ID del producto a validar
 * @returns {Object} { valido, motivo, manejaStock, aplicaIva, costoUnitario }
 */
function _validarProducto(idProducto) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);

    if (!hoja) {
      return { valido: false, motivo: 'Hoja MSTR_PRODUCTOS no encontrada' };
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { valido: false, motivo: 'No hay productos registrados' };
    }

    var datos = hoja.getRange(2, 1, ultimaFila - 1, 29).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var id     = String(fila[COL.PRD_ID_PRODUCTO]).trim();
      var activo = fila[COL.PRD_ACTIVO];

      if (id === idProducto) {
        if (!activo) {
          return { valido: false, motivo: 'Producto inactivo: ' + idProducto };
        }
        return {
          valido:        true,
          motivo:        'OK',
          manejaStock:   fila[COL.PRD_MANEJA_STOCK] !== false,
          aplicaIva:     fila[COL.PRD_APLICA_IVA] !== false,
          costoUnitario: Number(fila[COL.PRD_COSTO_UNITARIO]) || 0
        };
      }
    }

    return { valido: false, motivo: 'Producto no encontrado: ' + idProducto };

  } catch (e) {
    Logger.log('ERROR _validarProducto: ' + e.message);
    return { valido: false, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _validarStock
// ============================================================

/**
 * Verifica que hay stock suficiente para la cantidad solicitada.
 *
 * @param {string} idProducto - ID del producto
 * @param {number} cantidad   - Cantidad requerida
 * @returns {Object} { valido, motivo, stockActual }
 */
function _validarStock(idProducto, cantidad) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);

    if (!hoja) {
      return { valido: false, motivo: 'Hoja MSTR_PRODUCTOS no encontrada', stockActual: 0 };
    }

    var fila   = _buscarFilaPorId(hoja, idProducto, COL.PRD_ID_PRODUCTO);
    if (fila === -1) {
      return { valido: false, motivo: 'Producto no encontrado', stockActual: 0 };
    }

    var stock = Number(hoja.getRange(fila, COL.PRD_STOCK_ACTUAL + 1).getValue());

    if (stock < cantidad) {
      return {
        valido:      false,
        motivo:      'Stock insuficiente. Disponible: ' + stock + ' | Solicitado: ' + cantidad,
        stockActual: stock
      };
    }

    return { valido: true, motivo: 'OK', stockActual: stock };

  } catch (e) {
    Logger.log('ERROR _validarStock: ' + e.message);
    return { valido: false, motivo: 'Error: ' + e.message, stockActual: 0 };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _capturarCostoMomento
// ============================================================

/**
 * Lee el costo unitario actual del producto para el snapshot.
 * Este valor queda fijo en la línea de venta — no cambia aunque
 * el costo del producto cambie en el futuro.
 *
 * @param {string} idProducto - ID del producto
 * @returns {number} Costo unitario actual o 0 si no se encuentra
 */
function _capturarCostoMomento(idProducto) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    if (!hoja) return 0;

    var fila = _buscarFilaPorId(hoja, idProducto, COL.PRD_ID_PRODUCTO);
    if (fila === -1) return 0;

    return Number(hoja.getRange(fila, COL.PRD_COSTO_UNITARIO + 1).getValue()) || 0;

  } catch (e) {
    Logger.log('ERROR _capturarCostoMomento: ' + e.message);
    return 0;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _calcularImpactoStock
// ============================================================

/**
 * Calcula si un movimiento suma o resta stock.
 * Entradas: positivo. Salidas: negativo. Traslados: 0.
 *
 * @param {string} tipoMovimiento - Tipo según ESTADOS.MOV_*
 * @param {number} cantidad       - Cantidad del movimiento
 * @returns {number} Impacto en stock (positivo o negativo)
 */
function _calcularImpactoStock(tipoMovimiento, cantidad) {
  var salidas = [
    ESTADOS.MOV_SALIDA_VENTA,
    ESTADOS.MOV_AJUSTE_NEGATIVO,
    ESTADOS.MOV_DEV_PROVEEDOR
  ];

  if (tipoMovimiento === ESTADOS.MOV_TRASLADO) return 0;
  if (salidas.indexOf(tipoMovimiento) !== -1) return -Math.abs(cantidad);
  return Math.abs(cantidad);
}

// ============================================================
// FUNCIÓN PRIVADA: _calcularFechaVencimiento
// ============================================================

/**
 * Calcula la fecha de vencimiento según la condición de pago.
 *
 * @param {Date}   fechaBase      - Fecha de inicio del plazo
 * @param {string} condicionPago  - Condición: 30_DIAS, 60_DIAS, etc.
 * @returns {Date} Fecha de vencimiento calculada
 */
function _calcularFechaVencimiento(fechaBase, condicionPago) {
  try {
    var dias = 0;
    var mapa = {
      '15_DIAS': 15,
      '30_DIAS': 30,
      '45_DIAS': 45,
      '60_DIAS': 60,
      '90_DIAS': 90
    };

    dias = mapa[condicionPago] || 30;

    var fecha = new Date(fechaBase);
    fecha.setDate(fecha.getDate() + dias);
    return fecha;

  } catch (e) {
    Logger.log('ERROR _calcularFechaVencimiento: ' + e.message);
    var fallback = new Date();
    fallback.setDate(fallback.getDate() + 30);
    return fallback;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _detectarAccionCompra
// ============================================================

/**
 * Detecta si el evento en DB_COMPRAS corresponde a una nueva OC
 * o a una recepción de mercadería.
 *
 * @param {Object} e - Evento de edición
 */
function _detectarAccionCompra(e) {
  try {
    var hoja  = e.range.getSheet();
    var fila  = e.range.getRow();
    var datos = hoja.getRange(fila, 1, 1, 28).getValues()[0];

    var idLinea       = String(datos[COL.OC_ID_LINEA] || '').trim();
    var cantRecibida  = Number(datos[COL.OC_CANTIDAD_RECIBIDA]);
    var estadoOC      = String(datos[COL.OC_ESTADO_OC] || '').trim();

    // Si ya tiene ID y tiene cantidad recibida → es una recepción
    if (idLinea !== '' && cantRecibida > 0 &&
        estadoOC !== ESTADOS.OC_RECIBIDA &&
        estadoOC !== ESTADOS.OC_ANULADA) {
      procesarRecepcionOC(e);
      return;
    }

    // Si no tiene ID → es una OC nueva
    if (idLinea === '') {
      procesarCompra(e);
    }

  } catch (e2) {
    Logger.log('ERROR _detectarAccionCompra: ' + e2.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _detectarPagoCxC
// ============================================================

/**
 * Detecta si el evento en DB_CXC corresponde a un pago.
 * Solo procesa si cambió la columna monto_pagado (J = col 10).
 *
 * @param {Object} e - Evento de edición
 */
function _detectarPagoCxC(e) {
  try {
    var col = e.range.getColumn();
    // Solo procesar si el cambio fue en monto_pagado (col J = 10)
    if (col !== 10) return;

    var hoja     = e.range.getSheet();
    var fila     = e.range.getRow();
    var datos    = hoja.getRange(fila, 1, 1, 21).getValues()[0];
    var idCxC    = String(datos[0] || '').trim();
    var montoOri = Number(datos[8]);
    var montoPag = Number(datos[9]);
    var estado   = String(datos[13] || '').trim();

    if (!idCxC || montoOri <= 0) return;

    // Actualizar estado según el pago
    var nuevoEstado = montoPag >= montoOri ? ESTADOS.PAGADO : ESTADOS.ABONADO;

    if (nuevoEstado !== estado) {
      hoja.getRange(fila, 14).setValue(nuevoEstado); // N estado
      hoja.getRange(fila, 15).setValue(new Date());   // O fecha_ultimo_pago
      SpreadsheetApp.flush();

      registrarLog(HOJAS.DB_CXC, 'PAGO', idCxC, 'estado', nuevoEstado);
      Logger.log('_detectarPagoCxC: ' + idCxC + ' → ' + nuevoEstado);
    }

  } catch (e2) {
    Logger.log('ERROR _detectarPagoCxC: ' + e2.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _detectarPagoCxP
// ============================================================

/**
 * Detecta si el evento en DB_CXP corresponde a un pago.
 * Solo procesa si cambió la columna monto_pagado (J = col 10).
 *
 * @param {Object} e - Evento de edición
 */
function _detectarPagoCxP(e) {
  try {
    var col = e.range.getColumn();
    // Solo procesar si el cambio fue en monto_pagado (col J = 10)
    if (col !== 10) return;

    var hoja     = e.range.getSheet();
    var fila     = e.range.getRow();
    var datos    = hoja.getRange(fila, 1, 1, 21).getValues()[0];
    var idCxP    = String(datos[0] || '').trim();
    var montoOri = Number(datos[8]);
    var montoPag = Number(datos[9]);
    var estado   = String(datos[15] || '').trim();

    if (!idCxP || montoOri <= 0) return;

    var nuevoEstado = montoPag >= montoOri ? ESTADOS.PAGADO : ESTADOS.ABONADO;

    if (nuevoEstado !== estado) {
      hoja.getRange(fila, 16).setValue(nuevoEstado); // P estado
      hoja.getRange(fila, 17).setValue(new Date());   // Q fecha_ultimo_pago
      SpreadsheetApp.flush();

      registrarLog(HOJAS.DB_CXP, 'PAGO', idCxP, 'estado', nuevoEstado);
      Logger.log('_detectarPagoCxP: ' + idCxP + ' → ' + nuevoEstado);
    }

  } catch (e2) {
    Logger.log('ERROR _detectarPagoCxP: ' + e2.message);
  }
}

// ============================================================
// FIN DE M2_Transacciones.gs — PARTE 1
// ============================================================
// ============================================================
// M2_Transacciones.gs — PARTE 2
// ============================================================

// ============================================================
// FUNCIÓN: procesarCompra
// Registra una nueva línea de OC en DB_COMPRAS
// ============================================================

/**
 * Procesa una nueva orden de compra cuando el usuario completa
 * los datos en DB_COMPRAS.
 *
 * Campos obligatorios para disparar el proceso:
 * - id_proveedor (col D)
 * - id_producto (col E)
 * - fecha_emision (col G)
 * - cantidad_pedida (col L)
 * - precio_unitario_neto (col N)
 *
 * @param {Object} e - Evento de edición
 */
function procesarCompra(e) {
  try {
    var hoja  = e.range.getSheet();
    var fila  = e.range.getRow();
    var datos = hoja.getRange(fila, 1, 1, 28).getValues()[0];

    // Verificar que no tiene ID aún
    var idLinea = String(datos[COL.OC_ID_LINEA] || '').trim();
    if (idLinea !== '') return;

    // Verificar campos obligatorios
    var idProveedor  = String(datos[COL.OC_ID_PROVEEDOR] || '').trim();
    var idProducto   = String(datos[COL.OC_ID_PRODUCTO] || '').trim();
    var fechaEmision = datos[COL.OC_FECHA_EMISION];
    var cantPedida   = Number(datos[COL.OC_CANTIDAD_PEDIDA]);
    var precioUnit   = Number(datos[COL.OC_PRECIO_UNIT_NETO]);

    if (!idProveedor || !idProducto || !fechaEmision ||
        cantPedida <= 0 || precioUnit <= 0) {
      return; // Fila incompleta
    }

    // VALIDACIÓN 1: Período abierto
    if (!validarPeriodoAbierto(true)) {
      SpreadsheetApp.getUi().alert(
        '⛔ Sin período activo\n\n' +
        'No existe un período contable ABIERTO.'
      );
      return;
    }

    // VALIDACIÓN 2: Proveedor existe y está activo
    var validPrv = _validarProveedor(idProveedor);
    if (!validPrv.valido) {
      SpreadsheetApp.getUi().alert('❌ Proveedor inválido\n\n' + validPrv.motivo);
      return;
    }

    // VALIDACIÓN 3: Producto existe y está activo
    var validPrd = _validarProducto(idProducto);
    if (!validPrd.valido) {
      SpreadsheetApp.getUi().alert('❌ Producto inválido\n\n' + validPrd.motivo);
      return;
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      var config    = getConfig();
      var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
      var ahora     = new Date();
      var usuario   = Session.getActiveUser().getEmail() || 'sistema';
      var nuevoId   = generarID('OC');

      // Escribir campos del sistema
      hoja.getRange(fila, 1).setValue(nuevoId);              // A id_linea_oc
      hoja.getRange(fila, 3).setValue(idEmpresa);            // C id_empresa
      hoja.getRange(fila, 23).setValue(ESTADOS.OC_EMITIDA);  // W estado_oc
      hoja.getRange(fila, 27).setValue(ahora);               // AA creado_en
      hoja.getRange(fila, 28).setValue(usuario);             // AB creado_por

      SpreadsheetApp.flush();
      lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.DB_COMPRAS, 'CREAR', nuevoId, 'id_proveedor', idProveedor);
    Logger.log('procesarCompra: ' + nuevoId + ' creada — Proveedor: ' +
      idProveedor + ' Producto: ' + idProducto);

  } catch (e) {
    Logger.log('ERROR procesarCompra: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: procesarRecepcionOC
// Registra la recepción de una OC y actualiza el stock
// ============================================================

/**
 * Procesa la recepción de mercadería de una OC.
 * Actualiza stock, registra el costo actual y crea CxP si corresponde.
 *
 * Se dispara cuando el usuario ingresa cantidad_recibida (col M)
 * en una OC que ya tiene ID y no está RECIBIDA ni ANULADA.
 *
 * @param {Object} e - Evento de edición
 */
function procesarRecepcionOC(e) {
  try {
    var hoja  = e.range.getSheet();
    var fila  = e.range.getRow();
    var datos = hoja.getRange(fila, 1, 1, 28).getValues()[0];

    var idLinea      = String(datos[COL.OC_ID_LINEA] || '').trim();
    var idOC         = String(datos[COL.OC_ID_OC] || '').trim();
    var idProveedor  = String(datos[COL.OC_ID_PROVEEDOR] || '').trim();
    var idProducto   = String(datos[COL.OC_ID_PRODUCTO] || '').trim();
    var cantPedida   = Number(datos[COL.OC_CANTIDAD_PEDIDA]);
    var cantRecibida = Number(datos[COL.OC_CANTIDAD_RECIBIDA]);
    var precioUnit   = Number(datos[COL.OC_PRECIO_UNIT_NETO]);
    var descuento    = Number(datos[COL.OC_DESCUENTO_PCT]) || 0;
    var moneda       = String(datos[COL.OC_MONEDA_ORIGEN] || 'CLP').trim();
    var tipoCambio   = Number(datos[COL.OC_TIPO_CAMBIO]) || 1;
    var idBodega     = String(datos[COL.OC_ID_BODEGA_DESTINO] || '').trim();

    if (!idLinea || cantRecibida <= 0) return;

    // VALIDACIÓN: cantidad recibida no puede superar la pedida
    if (cantRecibida > cantPedida) {
      SpreadsheetApp.getUi().alert(
        '❌ Cantidad inválida\n\n' +
        'La cantidad recibida (' + cantRecibida + ') no puede\n' +
        'superar la cantidad pedida (' + cantPedida + ').'
      );
      hoja.getRange(fila, COL.OC_CANTIDAD_RECIBIDA + 1).clearContent();
      return;
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      var ahora   = new Date();
      var usuario = Session.getActiveUser().getEmail() || 'sistema';

      // Determinar nuevo estado de la OC
      var nuevoEstado = cantRecibida >= cantPedida
        ? ESTADOS.OC_RECIBIDA
        : ESTADOS.OC_PARCIAL;

      // Actualizar estado y fecha
      hoja.getRange(fila, 10).setValue(ahora);        // J fecha_entrega_real
      hoja.getRange(fila, 23).setValue(nuevoEstado);  // W estado_oc
      hoja.getRange(fila, 29).setValue(ahora);        // AC modificado_en
      hoja.getRange(fila, 30).setValue(usuario);      // AD modificado_por

      SpreadsheetApp.flush();
      lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    // Actualizar stock si el producto lo maneja
    var validPrd = _validarProducto(idProducto);
    if (validPrd.valido && validPrd.manejaStock) {
      if (!idBodega) {
        var bodegaDef = validarBodegaDefecto();
        idBodega = bodegaDef.valida ? bodegaDef.idBodega : '';
      }
      actualizarStock(
        idProducto,
        cantRecibida,
        ESTADOS.MOV_ENTRADA_COMPRA,
        idBodega,
        idLinea,
        'COMPRA'
      );

      // Actualizar costo unitario en MSTR_PRODUCTOS
      _actualizarCostoProducto(idProducto, precioUnit, moneda, tipoCambio);
    }

    // Registrar en historial de precios
    _registrarHistorialPrecio(idProducto, idProveedor, precioUnit, moneda, tipoCambio);

    // Crear CxP si no es contado
    var validPrv = _validarProveedor(idProveedor);
    if (validPrv.valido && validPrv.condicionPago !== 'CONTADO') {
      var subtotal   = cantRecibida * precioUnit * (1 - descuento);
      var config     = getConfig();
      var tasa       = config ? Number(config.tasa_iva) : 0.19;
      var iva        = Math.round(subtotal * tasa);
      var totalBruto = Math.round(subtotal + iva);

      crearCxPDesdeCompra(
        idOC || idLinea,
        idProveedor,
        totalBruto,
        validPrv.condicionPago,
        new Date()
      );
    }

    registrarLog(HOJAS.DB_COMPRAS, 'RECEPCION', idLinea, 'cantidad_recibida',
      String(cantRecibida));
    Logger.log('procesarRecepcionOC: ' + idLinea + ' recibida — Cant: ' + cantRecibida);

  } catch (e) {
    Logger.log('ERROR procesarRecepcionOC: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: crearCxPDesdeCompra
// Crea automáticamente el registro en DB_CXP
// ============================================================

/**
 * Crea un registro en DB_CXP cuando se recepciona una compra a crédito.
 *
 * @param {string} idDocumento   - ID de la OC
 * @param {string} idProveedor   - ID del proveedor
 * @param {number} montoTotal    - Monto total con IVA
 * @param {string} condicionPago - Condición de pago del proveedor
 * @param {Date}   fechaEmision  - Fecha de la recepción
 */
function crearCxPDesdeCompra(idDocumento, idProveedor, montoTotal, condicionPago, fechaEmision) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_CXP);
    if (!hoja) {
      Logger.log('ERROR crearCxPDesdeCompra: DB_CXP no encontrada');
      return false;
    }

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var usuario   = Session.getActiveUser().getEmail() || 'sistema';
    var idCxP     = generarID('CXP');
    var fechaVenc = _calcularFechaVencimiento(fechaEmision || ahora, condicionPago);

    var fila = new Array(20).fill('');
    fila[0]  = idCxP;                     // A id_cxp
    fila[1]  = idEmpresa;                 // B id_empresa
    fila[2]  = idProveedor;               // C id_proveedor
    fila[3]  = '';                        // D id_cc
    fila[4]  = idDocumento;               // E id_documento
    fila[5]  = 'COMPRA';                  // F tipo_documento
    fila[6]  = fechaEmision || ahora;     // G fecha_emision
    fila[7]  = fechaVenc;                 // H fecha_vencimiento
    fila[8]  = Math.round(montoTotal);    // I monto_original
    fila[9]  = 0;                         // J monto_pagado
    fila[10] = 'CLP';                     // K moneda_origen
    fila[11] = 1;                         // L tipo_cambio_momento
    fila[12] = ''; // M calc_saldo_pendiente — ARRAYFORMULA
    fila[13] = ''; // N calc_dias_vencimiento — ARRAYFORMULA
    fila[14] = ''; // O calc_tramo_aging — ARRAYFORMULA
    fila[15] = ESTADOS.PENDIENTE;         // P estado
    fila[16] = '';                        // Q fecha_ultimo_pago
    fila[17] = '';                        // R notas
    fila[18] = ahora;                     // S creado_en
    fila[19] = usuario;                   // T creado_por

    _escribirFilaSegura(hoja, fila, HOJAS.DB_CXP);

    registrarLog(HOJAS.DB_CXP, 'CREAR', idCxP, 'id_proveedor', idProveedor);
    Logger.log('crearCxPDesdeCompra: ' + idCxP + ' creada — Proveedor: ' +
      idProveedor + ' Monto: ' + montoTotal);
    return true;

  } catch (e) {
    Logger.log('ERROR crearCxPDesdeCompra: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: registrarIngreso
// Registra un ingreso financiero en DB_INGRESOS
// ============================================================

function registrarIngreso(datos) {
  try {
    if (!datos || !datos.id_cuenta || !datos.monto_neto) {
      return { exito: false, id: null, motivo: 'id_cuenta y monto_neto son obligatorios' };
    }

    var validCuenta = validarCuentaContable(datos.id_cuenta);
    if (!validCuenta.valida) {
      return { exito: false, id: null, motivo: 'Cuenta inválida: ' + validCuenta.motivo };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_INGRESOS);
    if (!hoja) {
      return { exito: false, id: null, motivo: 'Hoja DB_INGRESOS no encontrada' };
    }

    var config     = getConfig();
    var idEmpresa  = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var tasa       = config ? Number(config.tasa_iva) : 0.19;
    var ahora      = new Date();
    var usuario    = Session.getActiveUser().getEmail() || 'sistema';
    // periodo calculado por ARRAYFORMULA — no se escribe por código
    var idIngreso  = generarID('ING');
    var montoNeto  = Number(datos.monto_neto);
    var iva        = datos.aplica_iva !== false ? Math.round(montoNeto * tasa) : 0;
    var montoBruto = montoNeto + iva;

    var fila = new Array(15).fill('');
    fila[0]  = idIngreso;
    fila[1]  = idEmpresa;
    fila[2]  = datos.fecha || ahora;
    fila[3]  = ''; // sys_periodo — calculado por ARRAYFORMULA
    fila[4]  = datos.id_cuenta;
    fila[5]  = datos.id_cc || '';
    fila[6]  = datos.descripcion || '';
    fila[7]  = montoNeto;
    fila[8]  = iva;
    fila[9]  = montoBruto;
    fila[10] = datos.numero_documento || '';
    fila[11] = datos.tipo_documento || '';
    fila[12] = datos.id_origen || '';
    fila[13] = ahora;
    fila[14] = usuario;

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      _escribirFilaSegura(hoja, fila, HOJAS.DB_INGRESOS);
      lock.releaseLock();
    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.DB_INGRESOS, 'CREAR', idIngreso, 'id_cuenta', datos.id_cuenta);
    Logger.log('registrarIngreso: ' + idIngreso + ' — Cuenta: ' +
      datos.id_cuenta + ' Monto: ' + montoNeto);
    return { exito: true, id: idIngreso, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR registrarIngreso: ' + e.message);
    return { exito: false, id: null, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: registrarEgreso
// Registra un egreso financiero en DB_EGRESOS
// ============================================================

function registrarEgreso(datos) {
  try {
    if (!datos || !datos.id_cuenta || !datos.monto_neto) {
      return { exito: false, id: null, motivo: 'id_cuenta y monto_neto son obligatorios' };
    }

    var validCuenta = validarCuentaContable(datos.id_cuenta);
    if (!validCuenta.valida) {
      return { exito: false, id: null, motivo: 'Cuenta inválida: ' + validCuenta.motivo };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_EGRESOS);
    if (!hoja) {
      return { exito: false, id: null, motivo: 'Hoja DB_EGRESOS no encontrada' };
    }

    var config     = getConfig();
    var idEmpresa  = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var tasa       = config ? Number(config.tasa_iva) : 0.19;
    var ahora      = new Date();
    var usuario    = Session.getActiveUser().getEmail() || 'sistema';
    // periodo calculado por ARRAYFORMULA — no se escribe por código
    var idEgreso   = generarID('EGR');
    var montoNeto  = Number(datos.monto_neto);
    var iva        = datos.aplica_iva !== false ? Math.round(montoNeto * tasa) : 0;
    var montoBruto = montoNeto + iva;
    var moneda     = datos.moneda || 'CLP';
    var tipoCambio = Number(datos.tipo_cambio) || 1;
    var montoClp   = moneda === 'CLP' ? montoBruto : Math.round(montoBruto * tipoCambio);

    var fila = new Array(18).fill('');
    fila[0]  = idEgreso;
    fila[1]  = idEmpresa;
    fila[2]  = datos.fecha || ahora;
    fila[3]  = ''; // sys_periodo — calculado por ARRAYFORMULA
    fila[4]  = datos.id_cuenta;
    fila[5]  = datos.id_cc || '';
    fila[6]  = datos.descripcion || '';
    fila[7]  = montoNeto;
    fila[8]  = iva;
    fila[9]  = montoBruto;
    fila[10] = moneda;
    fila[11] = tipoCambio;
    fila[12] = ''; // sys_monto_clp — calculado por ARRAYFORMULA
    fila[13] = datos.numero_documento || '';
    fila[14] = datos.tipo_documento || '';
    fila[15] = datos.id_origen || '';
    fila[16] = ahora;
    fila[17] = usuario;

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      _escribirFilaSegura(hoja, fila, HOJAS.DB_EGRESOS);
      lock.releaseLock();
    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.DB_EGRESOS, 'CREAR', idEgreso, 'id_cuenta', datos.id_cuenta);
    Logger.log('registrarEgreso: ' + idEgreso + ' — Monto: ' + montoNeto);
    return { exito: true, id: idEgreso, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR registrarEgreso: ' + e.message);
    return { exito: false, id: null, motivo: 'Error: ' + e.message };
  }
}



// ============================================================
// FUNCIÓN: procesarAnulacion
// Anula una venta o compra y revierte el stock
// ============================================================

/**
 * Anula un documento de venta o compra.
 * Cambia el estado a ANULADO y revierte el movimiento de stock.
 *
 * @param {string} tipo        - 'VENTA' o 'COMPRA'
 * @param {string} idDocumento - ID del documento a anular
 * @param {string} motivo      - Motivo de la anulación
 * @returns {boolean} true si se anuló correctamente
 */
function procesarAnulacion(tipo, idDocumento, motivo) {
  try {
    var ui = SpreadsheetApp.getUi();

    if (!tipo || !idDocumento || !motivo) {
      ui.alert('❌ Faltan datos\n\nTipo, ID de documento y motivo son obligatorios.');
      return false;
    }

    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var nombreHoja = tipo === 'VENTA' ? HOJAS.DB_VENTAS : HOJAS.DB_COMPRAS;
    var hoja       = ss.getSheetByName(nombreHoja);

    if (!hoja) {
      Logger.log('ERROR procesarAnulacion: Hoja no encontrada: ' + nombreHoja);
      return false;
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return false;

    // Buscar todas las líneas del documento
    var colId     = tipo === 'VENTA' ? COL.VTA_ID_DOCUMENTO + 1 : COL.OC_ID_OC + 1;
    var colEstado = tipo === 'VENTA' ? COL.VTA_ESTADO + 1 : COL.OC_ESTADO_OC + 1;
    var colMotivo = tipo === 'VENTA' ? 29 : 24; // AC en ventas, X en compras
    var numCols   = tipo === 'VENTA' ? 34 : 30;

    var datos = hoja.getRange(2, 1, ultimaFila - 1, numCols).getValues();
    var lineasAnuladas = 0;

    for (var i = 0; i < datos.length; i++) {
      var fila    = datos[i];
      var idDoc   = String(fila[tipo === 'VENTA' ? COL.VTA_ID_DOCUMENTO : COL.OC_ID_OC] || '');
      var estado  = String(fila[tipo === 'VENTA' ? COL.VTA_ESTADO : COL.OC_ESTADO_OC] || '');

      if (idDoc !== idDocumento) continue;
      if (estado === ESTADOS.ANULADO || estado === ESTADOS.OC_ANULADA) continue;

      var filaReal = i + 2;

      // Cambiar estado a ANULADO
      var estadoAnulado = tipo === 'VENTA' ? ESTADOS.ANULADO : ESTADOS.OC_ANULADA;
      hoja.getRange(filaReal, colEstado).setValue(estadoAnulado);
      hoja.getRange(filaReal, colMotivo).setValue(motivo);

      // Revertir stock si corresponde
      var idProducto = String(
        fila[tipo === 'VENTA' ? COL.VTA_ID_PRODUCTO : COL.OC_ID_PRODUCTO] || ''
      );
      var cantidad = Number(
        fila[tipo === 'VENTA' ? COL.VTA_CANTIDAD : COL.OC_CANTIDAD_RECIBIDA] || 0
      );

      if (idProducto && cantidad > 0) {
        var validPrd = _validarProducto(idProducto);
        if (validPrd.valido && validPrd.manejaStock) {
          // Revertir: si era venta → devolución cliente, si era compra → devolución proveedor
          var tipoReversion = tipo === 'VENTA'
            ? ESTADOS.MOV_DEV_CLIENTE
            : ESTADOS.MOV_DEV_PROVEEDOR;

          actualizarStock(
            idProducto,
            cantidad,
            tipoReversion,
            '',
            idDocumento,
            'ANULACION'
          );
        }
      }

      lineasAnuladas++;
    }

    SpreadsheetApp.flush();

    if (lineasAnuladas === 0) {
      ui.alert('❌ Documento no encontrado o ya está anulado: ' + idDocumento);
      return false;
    }

    registrarLog(nombreHoja, 'ANULAR', idDocumento, 'motivo_anulacion', motivo);
    Logger.log('procesarAnulacion: ' + idDocumento + ' anulado — ' +
      lineasAnuladas + ' líneas');
    return true;

  } catch (e) {
    Logger.log('ERROR procesarAnulacion: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: obtenerTipoCambio
// Obtiene el tipo de cambio para una moneda y fecha
// ============================================================

/**
 * Busca el tipo de cambio más reciente para una moneda dada.
 * Si no encuentra uno exacto para la fecha busca el más cercano.
 *
 * @param {string} moneda - Código de moneda: 'USD', 'EUR'
 * @param {Date}   fecha  - Fecha para la que se busca el TC
 * @returns {number} Tipo de cambio en CLP o 1 si es CLP
 */
function obtenerTipoCambio(moneda, fecha) {
  try {
    if (!moneda || moneda === 'CLP') return 1;

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.MSTR_TIPOS_CAMBIO);

    if (!hoja) {
      Logger.log('obtenerTipoCambio: Hoja MSTR_TIPOS_CAMBIO no encontrada');
      return 1;
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      Logger.log('obtenerTipoCambio: No hay tipos de cambio registrados');
      return 1;
    }

    // Leer columnas: moneda(C=2), fecha(D=3), valor_clp(E=4)
    var datos       = hoja.getRange(2, 1, ultimaFila - 1, 5).getValues();
    var fechaBusq   = fecha ? new Date(fecha) : new Date();
    var tcEncontrado = 0;
    var fechaMasProx = null;

    for (var i = 0; i < datos.length; i++) {
      var fila       = datos[i];
      var monedaFila = String(fila[2]).trim().toUpperCase();
      var fechaFila  = new Date(fila[3]);
      var valorFila  = Number(fila[4]);

      if (monedaFila !== moneda.toUpperCase()) continue;
      if (valorFila <= 0) continue;

      // Buscar el más cercano a la fecha solicitada sin superar
      if (fechaFila <= fechaBusq) {
        if (!fechaMasProx || fechaFila > fechaMasProx) {
          fechaMasProx  = fechaFila;
          tcEncontrado  = valorFila;
        }
      }
    }

    if (tcEncontrado > 0) {
      Logger.log('obtenerTipoCambio: ' + moneda + ' = ' + tcEncontrado);
      return tcEncontrado;
    }

    Logger.log('ADVERTENCIA obtenerTipoCambio: No se encontró TC para ' +
      moneda + ' — usando 1');
    return 1;

  } catch (e) {
    Logger.log('ERROR obtenerTipoCambio: ' + e.message);
    return 1;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _validarProveedor
// ============================================================

/**
 * Verifica que un proveedor existe y está activo.
 *
 * @param {string} idProveedor - ID del proveedor
 * @returns {Object} { valido, motivo, condicionPago }
 */
function _validarProveedor(idProveedor) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PROVEEDORES);

    if (!hoja) {
      return { valido: false, motivo: 'Hoja MSTR_PROVEEDORES no encontrada', condicionPago: '' };
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { valido: false, motivo: 'No hay proveedores registrados', condicionPago: '' };
    }

    var datos = hoja.getRange(2, 1, ultimaFila - 1, 19).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var id     = String(fila[COL_PRV.ID]).trim();
      var activo = fila[COL_PRV.ACTIVO];

      if (id === idProveedor) {
        if (!activo) {
          return {
            valido: false,
            motivo: 'Proveedor inactivo: ' + idProveedor,
            condicionPago: ''
          };
        }
        return {
          valido:        true,
          motivo:        'OK',
          condicionPago: String(fila[COL_PRV.CONDICION_PAGO] || 'CONTADO')
        };
      }
    }

    return {
      valido: false,
      motivo: 'Proveedor no encontrado: ' + idProveedor,
      condicionPago: ''
    };

  } catch (e) {
    Logger.log('ERROR _validarProveedor: ' + e.message);
    return { valido: false, motivo: 'Error: ' + e.message, condicionPago: '' };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _actualizarCostoProducto
// ============================================================

/**
 * Actualiza el costo unitario de un producto cuando se recepciona
 * una compra. Convierte a CLP si la compra fue en otra moneda.
 *
 * @param {string} idProducto  - ID del producto
 * @param {number} precioUnit  - Precio unitario de la compra
 * @param {string} moneda      - Moneda de la compra
 * @param {number} tipoCambio  - Tipo de cambio al momento de la compra
 */
function _actualizarCostoProducto(idProducto, precioUnit, moneda, tipoCambio) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    if (!hoja) return;

    var fila = _buscarFilaPorId(hoja, idProducto, COL.PRD_ID_PRODUCTO);
    if (fila === -1) return;

    // Convertir a CLP si es necesario
    var costoClp = moneda === 'CLP'
      ? precioUnit
      : Math.round(precioUnit * (tipoCambio || 1));

    hoja.getRange(fila, COL.PRD_COSTO_UNITARIO + 1).setValue(costoClp);
    SpreadsheetApp.flush();

    Logger.log('_actualizarCostoProducto: ' + idProducto +
      ' costo actualizado a ' + costoClp + ' CLP');

  } catch (e) {
    Logger.log('ERROR _actualizarCostoProducto: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _registrarHistorialPrecio
// ============================================================

/**
 * Registra el precio de compra en HIST_PRECIOS para
 * seguimiento histórico de evolución de precios por proveedor.
 *
 * @param {string} idProducto  - ID del producto
 * @param {string} idProveedor - ID del proveedor
 * @param {number} precioNeto  - Precio neto de la compra
 * @param {string} moneda      - Moneda de la compra
 * @param {number} tipoCambio  - Tipo de cambio al momento
 */
function _registrarHistorialPrecio(idProducto, idProveedor, precioNeto, moneda, tipoCambio) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.HIST_PRECIOS);
    if (!hoja) return;

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var periodo   = getPeriodoActual() || '';
    var precioClp = moneda === 'CLP'
      ? precioNeto
      : Math.round(precioNeto * (tipoCambio || 1));

    var fila = new Array(11).fill('');
    fila[0]  = generarID('HPC');  // A id_hist_precio
    fila[1]  = idEmpresa;         // B id_empresa
    fila[2]  = idProducto;        // C id_producto
    fila[3]  = idProveedor;       // D id_proveedor
    fila[4]  = ahora;             // E fecha_registro
    fila[5]  = periodo;           // F sys_periodo
    fila[6]  = precioNeto;        // G precio_neto
    fila[7]  = moneda;            // H moneda
    fila[8]  = precioClp;         // I precio_clp
    fila[9]  = '';                // J calc_variacion_pct — ARRAYFORMULA
    fila[10] = ahora;             // K creado_en

    _escribirFilaSegura(hoja, fila, HOJAS.HIST_PRECIOS);
    Logger.log('_registrarHistorialPrecio: Precio registrado para ' + idProducto);

  } catch (e) {
    Logger.log('ERROR _registrarHistorialPrecio: ' + e.message);
  }
}

// ============================================================
// MAPA DE COLUMNAS CON ARRAYFORMULA — índice base 0
// ============================================================
var AF_COLS = {
  'DB_INGRESOS':    [3],
  'DB_EGRESOS':     [3, 12],
  'DB_CXC':         [10, 11, 12],
  'DB_CXP':         [12, 13, 14],
  'DB_MOVIMIENTOS': [3, 9],
  'HIST_PRECIOS':   [],
  'DB_VENTAS':      [11, 19, 20, 21, 23, 24],
  'DB_COMPRAS':     [7, 15, 16, 17, 20],
  'DB_CAJA':        [3, 11, 13],
  'DB_PRESUPUESTO': [4],
  'DB_METAS':       [4],
  'HIST_STOCK':     [],
  'BATCH_RESULTADOS': []
};

function _getPrimeraFilaVaciaReal(hoja) {
  try {
    var ultimaFila = hoja.getMaxRows();
    if (ultimaFila <= 2) return 3;
    var colA = hoja.getRange(3, 1, ultimaFila - 2, 1).getValues();
    for (var i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] !== '' && colA[i][0] !== null) {
        return i + 4;
      }
    }
    return 3;
  } catch (e) {
    Logger.log('ERROR _getPrimeraFilaVaciaReal: ' + e.message);
    return 3;
  }
}

function _construirBloques(totalCols, colsFormula) {
  var bloques      = [];
  var inicioBloque = null;
  for (var c = 0; c < totalCols; c++) {
    var esFormula = colsFormula.indexOf(c) !== -1;
    if (!esFormula) {
      if (inicioBloque === null) inicioBloque = c;
    } else {
      if (inicioBloque !== null) {
        bloques.push({ inicio: inicioBloque, fin: c - 1 });
        inicioBloque = null;
      }
    }
  }
  if (inicioBloque !== null) {
    bloques.push({ inicio: inicioBloque, fin: totalCols - 1 });
  }
  return bloques;
}

function _escribirFilaSegura(hoja, datos, nombreHoja) {
  try {
    var colsFormula = AF_COLS[nombreHoja] || [];
    var filaDestino = _getPrimeraFilaVaciaReal(hoja);
    var bloques     = _construirBloques(datos.length, colsFormula);
    bloques.forEach(function(bloque) {
      var valores = [];
      for (var c = bloque.inicio; c <= bloque.fin; c++) {
        valores.push(datos[c] !== undefined ? datos[c] : '');
      }
      hoja.getRange(filaDestino, bloque.inicio + 1, 1, valores.length)
          .setValues([valores]);
    });
    SpreadsheetApp.flush();
    return filaDestino;
  } catch (e) {
    Logger.log('ERROR _escribirFilaSegura [' + nombreHoja + ']: ' + e.message);
    throw e;
  }
}
// ============================================================
// FIN DE M2_Transacciones.gs — PARTE 2
// ============================================================
