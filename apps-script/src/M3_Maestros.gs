/**
 * ============================================================
 * M3_Maestros.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Gestiona creación, modificación y baja de
 * registros maestros. Valida duplicados, normaliza RUTs
 * y ejecuta soft delete.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Módulos:      M0_Setup, M4_Auditoria
 *   Hojas:        MSTR_CLIENTES, MSTR_PROVEEDORES, MSTR_PRODUCTOS,
 *                 MSTR_BODEGAS, MSTR_PLAN_CUENTAS, MSTR_TABLAS_REF
 *   APIs Google:  SpreadsheetApp, LockService
 * ============================================================
 */

// ============================================================
// ÍNDICES DE COLUMNAS MAESTROS — base 0
// ============================================================

var COL_CLI = {
  ID:               0,   // A id_cliente
  ID_EMPRESA:       1,   // B id_empresa
  TIPO:             2,   // C tipo_cliente
  RUT_NORM:         3,   // D rut_normalizado
  RUT:              4,   // E rut_cliente
  RAZON_SOCIAL:     5,   // F razon_social
  NOMBRE_FANTASIA:  6,   // G nombre_fantasia
  GIRO:             7,   // H giro
  EMAIL:            8,   // I email
  TELEFONO:         9,   // J telefono
  DIRECCION:        10,  // K direccion
  COMUNA:           11,  // L comuna
  REGION:           12,  // M region
  ID_SEGMENTO:      13,  // N id_segmento
  CONDICION_PAGO:   14,  // O condicion_pago
  LIMITE_CREDITO:   15,  // P limite_credito
  ID_VENDEDOR:      16,  // Q id_vendedor
  ACTIVO:           17,  // R activo
  FECHA_INACT:      18,  // S fecha_inactivacion
  CREADO_EN:        19,  // T creado_en
  CREADO_POR:       20,  // U creado_por
  MODIFICADO_EN:    21,  // V modificado_en
  MODIFICADO_POR:   22   // W modificado_por
};

var COL_PRV = {
  ID:               0,   // A id_proveedor
  ID_EMPRESA:       1,   // B id_empresa
  TIPO:             2,   // C tipo_proveedor
  RUT_NORM:         3,   // D rut_normalizado
  RUT:              4,   // E rut_proveedor
  RAZON_SOCIAL:     5,   // F razon_social
  NOMBRE_FANTASIA:  6,   // G nombre_fantasia
  GIRO:             7,   // H giro
  EMAIL:            8,   // I email
  TELEFONO:         9,   // J telefono
  CONDICION_PAGO:   10,  // K condicion_pago
  PLAZO_ENTREGA:    11,  // L plazo_entrega_dias
  MONEDA:           12,  // M moneda_habitual
  CRITICO:          13,  // N critico
  SCORE_PRECIO:     14,  // O score_precio
  SCORE_CALIDAD:    15,  // P score_calidad
  SCORE_CUMPL:      16,  // Q score_cumplimiento
  CALC_SCORE:       17,  // R calc_score_global
  ACTIVO:           18,  // S activo
  FECHA_INACT:      19,  // T fecha_inactivacion
  CREADO_EN:        20,  // U creado_en
  CREADO_POR:       21,  // V creado_por
  MODIFICADO_EN:    22,  // W modificado_en
  MODIFICADO_POR:   23   // X modificado_por
};

// ============================================================
// FUNCIÓN: normalizarRUT
// Limpia y formatea un RUT chileno
// ============================================================

/**
 * Normaliza un RUT chileno al formato estándar XX.XXX.XXX-X.
 * También retorna el RUT sin puntos ni guión para comparaciones.
 *
 * @param {string} rut - RUT en cualquier formato
 * @returns {Object} { formateado, normalizado, dv, valido }
 */
function normalizarRUT(rut) {
  try {
    if (!rut || rut === '') {
      return { formateado: '', normalizado: '', dv: '', valido: false };
    }

    // Limpiar caracteres no válidos
    var rutLimpio = String(rut).replace(/[^0-9kK]/g, '').toUpperCase();

    if (rutLimpio.length < 2) {
      return { formateado: '', normalizado: '', dv: '', valido: false };
    }

    // Separar dígito verificador
    var dv     = rutLimpio.slice(-1);
    var numero = rutLimpio.slice(0, -1);

    if (numero.length === 0) {
      return { formateado: '', normalizado: '', dv: '', valido: false };
    }

    // Validar dígito verificador
    var valido = _validarDVRUT(numero, dv);

    // Formatear con puntos y guión
    var formateado = _formatearNumeroRUT(numero) + '-' + dv;

    return {
      formateado:   formateado,
      normalizado:  numero,
      dv:           dv,
      valido:       valido
    };

  } catch (e) {
    Logger.log('ERROR normalizarRUT: ' + e.message);
    return { formateado: '', normalizado: '', dv: '', valido: false };
  }
}

// ============================================================
// FUNCIÓN: verificarDuplicadoRUT
// Verifica que no existe otro registro con el mismo RUT
// ============================================================

/**
 * Busca si ya existe un registro con el mismo RUT normalizado.
 *
 * @param {string} tipo       - 'CLIENTE' o 'PROVEEDOR'
 * @param {string} rutNorm    - RUT normalizado sin puntos ni guión
 * @param {string} idExcluir  - ID a excluir de la búsqueda (para ediciones)
 * @returns {boolean} true si hay duplicado
 */
function verificarDuplicadoRUT(tipo, rutNorm, idExcluir) {
  try {
    if (!rutNorm || rutNorm === '') return false;

    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var nombreHoja = tipo === 'CLIENTE'
      ? HOJAS.MSTR_CLIENTES
      : HOJAS.MSTR_PROVEEDORES;
    var colRut     = tipo === 'CLIENTE' ? COL_CLI.RUT_NORM : COL_PRV.RUT_NORM;
    var colId      = 0;

    var hoja       = ss.getSheetByName(nombreHoja);
    if (!hoja) return false;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return false;

    var datos = hoja.getRange(2, 1, ultimaFila - 1, colRut + 1).getValues();

    for (var i = 0; i < datos.length; i++) {
      var idFila  = String(datos[i][colId]);
      var rutFila = String(datos[i][colRut]).trim();

      if (idExcluir && idFila === idExcluir) continue;
      if (rutFila === String(rutNorm).trim()) {
        Logger.log('verificarDuplicadoRUT: Duplicado encontrado en fila ' + (i + 2));
        return true;
      }
    }

    return false;

  } catch (e) {
    Logger.log('ERROR verificarDuplicadoRUT: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: verificarDuplicadoSKU
// Verifica que no existe otro producto con el mismo SKU
// ============================================================

/**
 * Busca si ya existe un producto con el mismo SKU.
 *
 * @param {string} sku        - SKU a verificar
 * @param {string} idExcluir  - ID producto a excluir (para ediciones)
 * @returns {boolean} true si hay duplicado
 */
function verificarDuplicadoSKU(sku, idExcluir) {
  try {
    if (!sku || sku === '') return false;

    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var hoja       = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    if (!hoja) return false;

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return false;

    var datos = hoja.getRange(2, 1, ultimaFila - 1,
      COL.PRD_SKU + 1).getValues();

    for (var i = 0; i < datos.length; i++) {
      var idFila  = String(datos[i][COL.PRD_ID_PRODUCTO]);
      var skuFila = String(datos[i][COL.PRD_SKU]).trim();

      if (idExcluir && idFila === idExcluir) continue;
      if (skuFila.toLowerCase() === String(sku).trim().toLowerCase()) {
        Logger.log('verificarDuplicadoSKU: SKU duplicado encontrado en fila ' + (i + 2));
        return true;
      }
    }

    return false;

  } catch (e) {
    Logger.log('ERROR verificarDuplicadoSKU: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: validarCuentaContable
// Verifica que una cuenta existe, está activa y es imputable
// ============================================================

/**
 * Valida que una cuenta contable puede recibir transacciones.
 * Solo cuentas de nivel 3 con imputa_directo=TRUE son válidas.
 *
 * @param {string} idCuenta - Código de cuenta. Ej: '4.1.1'
 * @returns {Object} { valida, motivo }
 */
function validarCuentaContable(idCuenta) {
  try {
    if (!idCuenta || idCuenta === '') {
      return { valida: false, motivo: 'Código de cuenta vacío' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PLAN_CUENTAS);
    if (!hoja) {
      return { valida: false, motivo: 'Hoja MSTR_PLAN_CUENTAS no encontrada' };
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { valida: false, motivo: 'Plan de cuentas vacío' };
    }

    // Leer columnas: id_cuenta(A), tipo(D), nivel(H), imputa_directo(I), activo(J)
    var datos = hoja.getRange(2, 1, ultimaFila - 1, 10).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila    = datos[i];
      var id      = String(fila[0]).trim();
      var nivel   = Number(fila[7]);
      var imputa  = fila[8];
      var activo  = fila[9];

      if (id === String(idCuenta).trim()) {
        if (!activo) {
          return { valida: false, motivo: 'Cuenta inactiva' };
        }
        if (nivel !== 3) {
          return { valida: false, motivo: 'Solo se pueden usar cuentas de nivel 3' };
        }
        if (!imputa) {
          return { valida: false, motivo: 'La cuenta no permite imputación directa' };
        }
        return { valida: true, motivo: 'OK' };
      }
    }

    return { valida: false, motivo: 'Cuenta no encontrada: ' + idCuenta };

  } catch (e) {
    Logger.log('ERROR validarCuentaContable: ' + e.message);
    return { valida: false, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: validarBodegaDefecto
// Verifica que existe exactamente una bodega marcada como defecto
// ============================================================

/**
 * Verifica que hay exactamente una bodega marcada como defecto.
 * Es un requisito del sistema para procesar movimientos de stock.
 *
 * @returns {Object} { valida, idBodega, motivo }
 */
function validarBodegaDefecto() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_BODEGAS);
    if (!hoja) {
      return { valida: false, idBodega: null, motivo: 'Hoja MSTR_BODEGAS no encontrada' };
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { valida: false, idBodega: null, motivo: 'No hay bodegas registradas' };
    }

    // Leer columnas: id(A), es_bodega_defecto(F=5), activo(G=6)
    var datos          = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();
    var bodegasDefecto = [];

    for (var i = 0; i < datos.length; i++) {
      var fila      = datos[i];
      var id        = String(fila[0]);
      var esDefault = fila[5];
      var activo    = fila[6];

      if (esDefault === true && activo === true) {
        bodegasDefecto.push(id);
      }
    }

    if (bodegasDefecto.length === 0) {
      return { valida: false, idBodega: null, motivo: 'No hay bodega marcada como defecto' };
    }
    if (bodegasDefecto.length > 1) {
      return {
        valida:   false,
        idBodega: null,
        motivo:   'Hay ' + bodegasDefecto.length + ' bodegas marcadas como defecto. Solo debe haber una.'
      };
    }

    return { valida: true, idBodega: bodegasDefecto[0], motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR validarBodegaDefecto: ' + e.message);
    return { valida: false, idBodega: null, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: validarSoftDelete
// Verifica que un registro puede darse de baja sin conflictos
// ============================================================

/**
 * Verifica si un maestro puede ser dado de baja.
 * Un cliente no puede darse de baja si tiene CxC pendiente.
 * Un proveedor no puede darse de baja si tiene CxP pendiente.
 * Un producto no puede darse de baja si tiene stock > 0.
 *
 * @param {string} tipo - 'CLIENTE', 'PROVEEDOR' o 'PRODUCTO'
 * @param {string} id   - ID del registro
 * @returns {Object} { puede, motivo }
 */
function validarSoftDelete(tipo, id) {
  try {
    if (!tipo || !id) {
      return { puede: false, motivo: 'Tipo o ID no especificado' };
    }

    switch (tipo) {
      case 'CLIENTE':
        return _verificarSaldoPendienteCxC(id);
      case 'PROVEEDOR':
        return _verificarSaldoPendienteCxP(id);
      case 'PRODUCTO':
        return _verificarStockActivo(id);
      default:
        return { puede: true, motivo: 'Tipo no requiere validación especial' };
    }

  } catch (e) {
    Logger.log('ERROR validarSoftDelete: ' + e.message);
    return { puede: false, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: darDeBajaMaestro
// Marca un registro como activo=FALSE sin borrarlo
// ============================================================

/**
 * Ejecuta el soft delete de un registro maestro.
 * Cambia activo=FALSE y registra la fecha de inactivación.
 * Nunca borra filas — el historial se conserva siempre.
 *
 * @param {string} tipo   - 'CLIENTE', 'PROVEEDOR' o 'PRODUCTO'
 * @param {string} id     - ID del registro a dar de baja
 * @param {string} motivo - Motivo de la baja
 * @returns {boolean} true si se ejecutó correctamente
 */
function darDeBajaMaestro(tipo, id, motivo) {
  try {
    // Validar que puede darse de baja
    var validacion = validarSoftDelete(tipo, id);
    if (!validacion.puede) {
      SpreadsheetApp.getUi().alert(
        '❌ No se puede dar de baja\n\n' +
        validacion.motivo
      );
      return false;
    }

    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var nombreHoja, colActivo, colFechaInact, colMotivoInact;

    switch (tipo) {
      case 'CLIENTE':
        nombreHoja    = HOJAS.MSTR_CLIENTES;
        colActivo     = COL_CLI.ACTIVO + 1;       // base 1 para getRange
        colFechaInact = COL_CLI.FECHA_INACT + 1;
        colMotivoInact = null; // MSTR_CLIENTES no tiene motivo_inactivacion
        break;
      case 'PROVEEDOR':
        nombreHoja    = HOJAS.MSTR_PROVEEDORES;
        colActivo     = COL_PRV.ACTIVO + 1;
        colFechaInact = COL_PRV.FECHA_INACT + 1;
        colMotivoInact = null;
        break;
      case 'PRODUCTO':
        nombreHoja    = HOJAS.MSTR_PRODUCTOS;
        colActivo     = COL.PRD_ACTIVO + 1;       // col AB = 28 base 1
        colFechaInact = 31;                        // col AE fecha_inactivacion
        colMotivoInact = 32;                       // col AF motivo_inactivacion
        break;
      default:
        Logger.log('darDeBajaMaestro: Tipo no reconocido: ' + tipo);
        return false;
    }

    var hoja       = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      Logger.log('darDeBajaMaestro: Hoja no encontrada: ' + nombreHoja);
      return false;
    }

    // Buscar la fila del registro
    var fila = _buscarFilaPorId(hoja, id, 0);
    if (fila === -1) {
      SpreadsheetApp.getUi().alert('❌ Registro no encontrado: ' + id);
      return false;
    }

    // Usar LockService para evitar ediciones concurrentes
    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      // Marcar como inactivo
      hoja.getRange(fila, colActivo).setValue(false);
      hoja.getRange(fila, colFechaInact).setValue(new Date());

      // Registrar motivo si la columna existe
      if (colMotivoInact) {
        hoja.getRange(fila, colMotivoInact).setValue(motivo || 'Baja manual');
      }

      SpreadsheetApp.flush();
      lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    // Registrar en auditoría
    registrarLog(nombreHoja, 'BAJA', id, 'activo', 'FALSE');
    registrarLog(nombreHoja, 'BAJA', id, 'motivo_inactivacion', motivo || 'Baja manual');

    Logger.log('darDeBajaMaestro: ' + tipo + ' ' + id + ' dado de baja correctamente');
    return true;

  } catch (e) {
    Logger.log('ERROR darDeBajaMaestro: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Error al dar de baja: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: crearCliente
// Crea un nuevo cliente en MSTR_CLIENTES
// ============================================================

/**
 * Crea un nuevo registro de cliente con todas las validaciones.
 * Valida RUT duplicado, normaliza el RUT y genera el ID.
 *
 * @param {Object} datos - Objeto con los datos del cliente
 * @returns {Object} { exito, id, motivo }
 */
function crearCliente(datos) {
  try {
    var ui = SpreadsheetApp.getUi();

    // Validaciones básicas
    if (!datos.rut || !datos.razon_social) {
      return { exito: false, id: null, motivo: 'RUT y razón social son obligatorios' };
    }

    // Normalizar RUT
    var rutInfo = normalizarRUT(datos.rut);
    if (!rutInfo.valido) {
      return { exito: false, id: null, motivo: 'RUT inválido: ' + datos.rut };
    }

    // Verificar duplicado
    if (verificarDuplicadoRUT('CLIENTE', rutInfo.normalizado, null)) {
      return {
        exito:  false,
        id:     null,
        motivo: 'Ya existe un cliente con RUT ' + rutInfo.formateado
      };
    }

    // Obtener configuración
    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var usuario   = Session.getActiveUser().getEmail() || 'sistema';

    // Generar ID
    var idCliente = generarID('CLI');

    // Construir fila completa
    var fila = [
      idCliente,                          // A id_cliente
      idEmpresa,                          // B id_empresa
      datos.tipo_cliente || 'OTRO',       // C tipo_cliente
      rutInfo.normalizado,                // D rut_normalizado
      rutInfo.formateado,                 // E rut_cliente
      String(datos.razon_social).trim(),  // F razon_social
      datos.nombre_fantasia || '',        // G nombre_fantasia
      datos.giro || '',                   // H giro
      datos.email || '',                  // I email
      datos.telefono || '',               // J telefono
      datos.direccion || '',              // K direccion
      datos.comuna || '',                 // L comuna
      datos.region || '',                 // M region
      datos.id_segmento || '',            // N id_segmento
      datos.condicion_pago || 'CONTADO', // O condicion_pago
      datos.limite_credito || 0,          // P limite_credito
      datos.id_vendedor || '',            // Q id_vendedor
      true,                               // R activo
      '',                                 // S fecha_inactivacion
      ahora,                              // T creado_en
      usuario,                            // U creado_por
      '',                                 // V modificado_en
      ''                                  // W modificado_por
    ];

    // Escribir en la hoja
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.MSTR_CLIENTES);
    if (!hoja) {
      return { exito: false, id: null, motivo: 'Hoja MSTR_CLIENTES no encontrada' };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      hoja.appendRow(fila);
      SpreadsheetApp.flush();
      lock.releaseLock();
    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.MSTR_CLIENTES, 'CREAR', idCliente, 'rut_cliente', rutInfo.formateado);
    Logger.log('crearCliente: ' + idCliente + ' creado — ' + datos.razon_social);

    return { exito: true, id: idCliente, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR crearCliente: ' + e.message);
    return { exito: false, id: null, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: crearProveedor
// Crea un nuevo proveedor en MSTR_PROVEEDORES
// ============================================================

/**
 * Crea un nuevo registro de proveedor con todas las validaciones.
 *
 * @param {Object} datos - Objeto con los datos del proveedor
 * @returns {Object} { exito, id, motivo }
 */
function crearProveedor(datos) {
  try {
    if (!datos.rut || !datos.razon_social) {
      return { exito: false, id: null, motivo: 'RUT y razón social son obligatorios' };
    }

    var rutInfo = normalizarRUT(datos.rut);
    if (!rutInfo.valido) {
      return { exito: false, id: null, motivo: 'RUT inválido: ' + datos.rut };
    }

    if (verificarDuplicadoRUT('PROVEEDOR', rutInfo.normalizado, null)) {
      return {
        exito:  false,
        id:     null,
        motivo: 'Ya existe un proveedor con RUT ' + rutInfo.formateado
      };
    }

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var usuario   = Session.getActiveUser().getEmail() || 'sistema';
    var idPrv     = generarID('PRV');

    var fila = [
  idPrv,                                // 0  → A id_proveedor
  idEmpresa,                            // 1  → B id_empresa
  datos.tipo_proveedor || 'OTRO',       // 2  → C tipo_proveedor
  rutInfo.normalizado,                  // 3  → D rut_normalizado
  rutInfo.formateado,                   // 4  → E rut_proveedor
  String(datos.razon_social).trim(),    // 5  → F razon_social
  datos.nombre_fantasia || '',          // 6  → G nombre_fantasia
  datos.giro || '',                     // 7  → H giro
  datos.email || '',                    // 8  → I email
  datos.telefono || '',                 // 9  → J telefono
  datos.condicion_pago || 'CONTADO',   // 10 → K condicion_pago
  datos.plazo_entrega_dias || 0,        // 11 → L plazo_entrega_dias
  datos.moneda_habitual || 'CLP',       // 12 → M moneda_habitual
  datos.critico || false,               // 13 → N critico
  datos.score_precio || 0,              // 14 → O score_precio
  datos.score_calidad || 0,             // 15 → P score_calidad
  datos.score_cumplimiento || 0,        // 16 → Q score_cumplimiento
  // índice 17 = R calc_score_global — ARRAYFORMULA — no se incluye en escritura
  true,                                 // 17 → S activo (col 19)
  '',                                   // 18 → T fecha_inactivacion (col 20)
  ahora,                                // 19 → U creado_en (col 21)
  usuario,                              // 20 → V creado_por (col 22)
  '',                                   // 21 → W modificado_en (col 23)
  ''                                    // 22 → X modificado_por (col 24)
];

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.MSTR_PROVEEDORES);
    if (!hoja) {
      return { exito: false, id: null, motivo: 'Hoja MSTR_PROVEEDORES no encontrada' };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    // REEMPLAZAR este bloque dentro de crearProveedor:

var lock = LockService.getScriptLock();
lock.waitLock(SISTEMA.TIMEOUT_LOCK);

try {
  var nuevaFila = _ultimaFilaReal(hoja) + 1;
  
  // Bloque 1: cols A-Q (índices 0-16) — antes de ARRAYFORMULA en R
  hoja.getRange(nuevaFila, 1, 1, 17).setValues([fila.slice(0, 17)]);
  
  // Bloque 2: col S (activo) — índice 17 del array → col 19 de la hoja
  // CORRECCIÓN H18: activo estaba en fila[17] pero nunca se escribía
  hoja.getRange(nuevaFila, 19, 1, 1).setValues([[fila[17]]]);
  
  // Bloque 3: cols T-X (índices 18-22) — fecha_inact, creado_en, creado_por, mod_en, mod_por
  hoja.getRange(nuevaFila, 20, 1, 5).setValues([fila.slice(18, 23)]);
  
  SpreadsheetApp.flush();
  lock.releaseLock();
} catch (eLock) {
  lock.releaseLock();
  throw eLock;
}

    registrarLog(HOJAS.MSTR_PROVEEDORES, 'CREAR', idPrv, 'rut_proveedor', rutInfo.formateado);
    Logger.log('crearProveedor: ' + idPrv + ' creado — ' + datos.razon_social);

    return { exito: true, id: idPrv, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR crearProveedor: ' + e.message);
    return { exito: false, id: null, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN: crearProducto
// Crea un nuevo producto en MSTR_PRODUCTOS
// ============================================================

/**
 * Crea un nuevo registro de producto con todas las validaciones.
 * Genera el movimiento de apertura de stock si stock_inicial > 0.
 *
 * @param {Object} datos - Objeto con los datos del producto
 * @returns {Object} { exito, id, motivo }
 */
function crearProveedor(datos) {
  try {
    if (!datos.rut || !datos.razon_social) {
      return { exito: false, id: null, motivo: 'RUT y razón social son obligatorios' };
    }

    var rutInfo = normalizarRUT(datos.rut);
    if (!rutInfo.valido) {
      return { exito: false, id: null, motivo: 'RUT inválido: ' + datos.rut };
    }

    if (verificarDuplicadoRUT('PROVEEDOR', rutInfo.normalizado, null)) {
      return {
        exito:  false,
        id:     null,
        motivo: 'Ya existe un proveedor con RUT ' + rutInfo.formateado
      };
    }

    var config    = getConfig();
    var idEmpresa = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;
    var ahora     = new Date();
    var usuario   = Session.getActiveUser().getEmail() || 'sistema';
    var idPrv     = generarID('PRV');

    var fila = [
      idPrv,                               // 0  → A id_proveedor
      idEmpresa,                           // 1  → B id_empresa
      datos.tipo_proveedor || 'OTRO',      // 2  → C tipo_proveedor
      rutInfo.normalizado,                 // 3  → D rut_normalizado
      rutInfo.formateado,                  // 4  → E rut_proveedor
      String(datos.razon_social).trim(),   // 5  → F razon_social
      datos.nombre_fantasia || '',         // 6  → G nombre_fantasia
      datos.giro || '',                    // 7  → H giro
      datos.email || '',                   // 8  → I email
      datos.telefono || '',                // 9  → J telefono
      datos.condicion_pago || 'CONTADO',  // 10 → K condicion_pago
      datos.plazo_entrega_dias || 0,       // 11 → L plazo_entrega_dias
      datos.moneda_habitual || 'CLP',      // 12 → M moneda_habitual
      datos.critico || false,              // 13 → N critico
      datos.score_precio || 0,             // 14 → O score_precio
      datos.score_calidad || 0,            // 15 → P score_calidad
      datos.score_cumplimiento || 0,       // 16 → Q score_cumplimiento
      // índice 17 = R calc_score_global — ARRAYFORMULA — no se escribe
      true,                                // 17 → S activo (col 19)
      '',                                  // 18 → T fecha_inactivacion (col 20)
      ahora,                               // 19 → U creado_en (col 21)
      usuario,                             // 20 → V creado_por (col 22)
      '',                                  // 21 → W modificado_en (col 23)
      ''                                   // 22 → X modificado_por (col 24)
    ];

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJAS.MSTR_PROVEEDORES);
    if (!hoja) {
      return { exito: false, id: null, motivo: 'Hoja MSTR_PROVEEDORES no encontrada' };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(SISTEMA.TIMEOUT_LOCK);

    try {
      var nuevaFila = _ultimaFilaReal(hoja) + 1;

      // Bloque 1: A-Q (índices 0-16) — antes de ARRAYFORMULA en R(18)
      hoja.getRange(nuevaFila, 1, 1, 17).setValues([fila.slice(0, 17)]);

      // Bloque 2: S (activo) — índice 17 del array → col 19 de la hoja
      hoja.getRange(nuevaFila, 19, 1, 1).setValues([[fila[17]]]);

      // Bloque 3: T-X (índices 18-22) → cols 20-24
      hoja.getRange(nuevaFila, 20, 1, 5).setValues([fila.slice(18, 23)]);

      SpreadsheetApp.flush();
      lock.releaseLock();

    } catch (eLock) {
      lock.releaseLock();
      throw eLock;
    }

    registrarLog(HOJAS.MSTR_PROVEEDORES, 'CREAR', idPrv, 'rut_proveedor', rutInfo.formateado);
    Logger.log('crearProveedor: ' + idPrv + ' creado — ' + datos.razon_social);

    return { exito: true, id: idPrv, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR crearProveedor: ' + e.message);
    return { exito: false, id: null, motivo: 'Error: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _registrarMovimientoApertura
// ============================================================

/**
 * Registra el movimiento de APERTURA cuando se crea un producto
 * con stock inicial mayor que cero.
 *
 * @param {string} idProducto    - ID del producto recién creado
 * @param {number} stockInicial  - Cantidad inicial de stock
 * @param {number} costoUnitario - Costo unitario al momento de apertura
 * @param {string} idBodega      - Bodega donde se registra el stock
 */
function _registrarMovimientoApertura(idProducto, stockInicial, costoUnitario, idBodega) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var hoja    = ss.getSheetByName(HOJAS.DB_MOVIMIENTOS);
    if (!hoja) {
      Logger.log('_registrarMovimientoApertura: DB_MOVIMIENTOS no encontrada');
      return;
    }

    var ahora   = new Date();
    var periodo = getPeriodoActual() || '';
    var usuario = Session.getActiveUser().getEmail() || 'sistema';
    var config  = getConfig();
    var idEmp   = config ? config.id_empresa : SISTEMA.EMPRESA_DEFAULT;

    var fila = new Array(16).fill('');
    fila[0]  = generarID('MOV');       // A id_movimiento
    fila[1]  = idEmp;                  // B id_empresa
    fila[2]  = ahora;                  // C fecha_movimiento
    fila[4]  = idProducto;             // E id_producto
    fila[5]  = idBodega;               // F id_bodega_origen
    fila[6]  = idBodega;               // G id_bodega_destino
    fila[7]  = ESTADOS.MOV_APERTURA;  // H tipo_movimiento
    fila[8]  = stockInicial;           // I cantidad
    fila[10] = costoUnitario;          // K costo_unitario_momento
    fila[11] = idProducto;             // L id_referencia
    fila[12] = 'APERTURA';             // M tipo_referencia
    fila[13] = 'Stock inicial';        // N motivo_ajuste
    fila[14] = ahora;                  // O creado_en
    fila[15] = usuario;                // P creado_por

    var nuevaFila = _ultimaFilaReal(hoja) + 1;
    // Escribir columnas A-C (índices 0-2, antes de ARRAYFORMULA en D=sys_periodo)
    hoja.getRange(nuevaFila, 1, 1, 3).setValues([fila.slice(0, 3)]);
    // Escribir columnas E-I (índices 4-8, entre ARRAYFORMULA D y J=sys_impacto_stock)
    hoja.getRange(nuevaFila, 5, 1, 5).setValues([fila.slice(4, 9)]);
    // Escribir columnas K-P (índices 10-15, después de ARRAYFORMULA en J)
    hoja.getRange(nuevaFila, 11, 1, 6).setValues([fila.slice(10, 16)]);
    SpreadsheetApp.flush();

    registrarLog(HOJAS.DB_MOVIMIENTOS, 'CREAR', fila[0], 'tipo_movimiento', 'APERTURA');
    Logger.log('_registrarMovimientoApertura: Movimiento APERTURA creado para ' + idProducto);

  } catch (e) {
    Logger.log('ERROR _registrarMovimientoApertura: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarSaldoPendienteCxC
// ============================================================

/**
 * Verifica que un cliente no tiene CxC pendiente de cobro.
 * Si tiene saldo pendiente no se puede dar de baja.
 *
 * @param {string} idCliente - ID del cliente a verificar
 * @returns {Object} { puede, motivo }
 */
function _verificarSaldoPendienteCxC(idCliente) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_CXC);
    if (!hoja) return { puede: true, motivo: 'OK' };

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { puede: true, motivo: 'OK' };

    // Leer columnas: id_cliente(C=2), calc_saldo_pendiente(K=10), estado(N=13)
    var datos = hoja.getRange(2, 1, ultimaFila - 1, 14).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila       = datos[i];
      var cliId      = String(fila[2]);
      var saldo      = Number(fila[10]);
      var estado     = String(fila[13]);

      if (cliId === idCliente &&
          saldo > 0 &&
          estado !== ESTADOS.PAGADO &&
          estado !== ESTADOS.INCOBRABLE) {
        return {
          puede:  false,
          motivo: 'El cliente tiene CxC pendiente por $' + saldo.toLocaleString('es-CL')
        };
      }
    }

    return { puede: true, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR _verificarSaldoPendienteCxC: ' + e.message);
    return { puede: false, motivo: 'Error verificando CxC: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarSaldoPendienteCxP
// ============================================================

/**
 * Verifica que un proveedor no tiene CxP pendiente de pago.
 *
 * @param {string} idProveedor - ID del proveedor a verificar
 * @returns {Object} { puede, motivo }
 */
function _verificarSaldoPendienteCxP(idProveedor) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.DB_CXP);
    if (!hoja) return { puede: true, motivo: 'OK' };

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { puede: true, motivo: 'OK' };

    // Leer columnas: id_proveedor(C=2), calc_saldo_pendiente(M=12), estado(P=15)
    var datos = hoja.getRange(2, 1, ultimaFila - 1, 16).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila   = datos[i];
      var prvId  = String(fila[2]);
      var saldo  = Number(fila[12]);
      var estado = String(fila[15]);

      if (prvId === idProveedor &&
          saldo > 0 &&
          estado !== ESTADOS.PAGADO) {
        return {
          puede:  false,
          motivo: 'El proveedor tiene CxP pendiente por $' + saldo.toLocaleString('es-CL')
        };
      }
    }

    return { puede: true, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR _verificarSaldoPendienteCxP: ' + e.message);
    return { puede: false, motivo: 'Error verificando CxP: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarStockActivo
// ============================================================

/**
 * Verifica que un producto no tiene stock antes de darlo de baja.
 *
 * @param {string} idProducto - ID del producto
 * @returns {Object} { puede, motivo }
 */
function _verificarStockActivo(idProducto) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.MSTR_PRODUCTOS);
    if (!hoja) return { puede: true, motivo: 'OK' };

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { puede: true, motivo: 'OK' };

    var datos = hoja.getRange(2, 1, ultimaFila - 1,
      COL.PRD_STOCK_ACTUAL + 1).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila  = datos[i];
      var id    = String(fila[COL.PRD_ID_PRODUCTO]);
      var stock = Number(fila[COL.PRD_STOCK_ACTUAL]);

      if (id === idProducto && stock > 0) {
        return {
          puede:  false,
          motivo: 'El producto tiene stock de ' + stock + ' unidades. Ajusta el stock a 0 antes de dar de baja.'
        };
      }
    }

    return { puede: true, motivo: 'OK' };

  } catch (e) {
    Logger.log('ERROR _verificarStockActivo: ' + e.message);
    return { puede: false, motivo: 'Error verificando stock: ' + e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _buscarFilaPorId
// ============================================================

/**
 * Busca la fila de un registro por su ID en la columna indicada.
 *
 * @param {Sheet}  hoja   - Objeto hoja de Google Sheets
 * @param {string} id     - ID a buscar
 * @param {number} colIdx - Índice base 0 de la columna del ID
 * @returns {number} Número de fila (base 1) o -1 si no se encuentra
 */
function _buscarFilaPorId(hoja, id, colIdx) {
  try {
    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return -1;

    var datos = hoja.getRange(2, colIdx + 1, ultimaFila - 1, 1).getValues();

    for (var i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(id).trim()) {
        return i + 2; // +2 porque índice base 0 + fila de encabezado
      }
    }

    return -1;

  } catch (e) {
    Logger.log('ERROR _buscarFilaPorId: ' + e.message);
    return -1;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _validarDVRUT
// ============================================================

/**
 * Valida el dígito verificador de un RUT chileno.
 *
 * @param {string} numero - Número del RUT sin DV
 * @param {string} dv     - Dígito verificador
 * @returns {boolean} true si el DV es correcto
 */
function _validarDVRUT(numero, dv) {
  try {
    var suma    = 0;
    var mul     = 2;
    var numStr  = String(numero).split('').reverse();

    for (var i = 0; i < numStr.length; i++) {
      suma += parseInt(numStr[i]) * mul;
      mul   = mul === 7 ? 2 : mul + 1;
    }

    var resto  = suma % 11;
    var dvCalc = resto === 0 ? '0' : resto === 1 ? 'K' : String(11 - resto);

    return dvCalc === String(dv).toUpperCase();

  } catch (e) {
    Logger.log('ERROR _validarDVRUT: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _formatearNumeroRUT
// ============================================================

/**
 * Formatea el número de un RUT con puntos separadores de miles.
 *
 * @param {string} numero - Número sin puntos
 * @returns {string} Número formateado con puntos
 */
function _formatearNumeroRUT(numero) {
  try {
    return String(numero).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  } catch (e) {
    return numero;
  }
}
// ============================================================
// FUNCIÓN PRIVADA: _ultimaFilaReal
// Busca la última fila con dato real en columna A (id_producto)
// Columna A es exclusivamente manual — nunca tiene ARRAYFORMULA
// ============================================================

/**
 * Retorna la última fila con dato real en columna A.
 * Lee solo columna A porque es campo de ID manual — garantiza
 * que nunca hay fórmulas calculadas que confundan el resultado.
 * 
 * Corrección H07: antes leía desde fila 3 y podía detectar
 * ARRAYFORMULAs de columnas AC/AD como "datos reales".
 * Ahora lee exclusivamente columna A desde fila 2.
 *
 * @param {Sheet} hoja - Objeto hoja de Google Sheets
 * @returns {number} Número de última fila con dato en col A (base 1)
 */
function _ultimaFilaReal(hoja) {
  try {
    var maxFila = hoja.getMaxRows();
    
    // Leer solo columna A desde fila 2
    // Columna A = id_* — siempre manual, nunca ARRAYFORMULA
    var datos = hoja.getRange(2, 1, maxFila - 1, 1).getValues();
    
    // Buscar desde abajo hacia arriba el primer valor no vacío
    for (var i = datos.length - 1; i >= 0; i--) {
      var valor = datos[i][0];
      if (valor !== '' && valor !== null && valor !== undefined) {
        return i + 2; // +2: índice base 0 + fila de encabezado
      }
    }
    
    // Sin datos reales — el próximo registro va a fila 2
    return 1;
    
  } catch (e) {
    Logger.log('ERROR _ultimaFilaReal [' + hoja.getName() + ']: ' + e.message);
    return hoja.getLastRow(); // Fallback seguro
  }
}
// ============================================================
// FIN DE M3_Maestros.gs
// ============================================================
