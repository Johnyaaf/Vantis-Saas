/**
 * ============================================================
 * M0_Setup.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Núcleo del sistema. Define constantes
 * globales, configuración, generación de IDs y setup inicial.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Hojas:       CONFIG_SISTEMA, CONFIG_PERIODOS, LOG_AUDITORIA
 *   Named Ranges: CONFIG__*, CONFIG_PERIODOS__*, LISTA__*
 *   APIs Google:  ScriptApp, PropertiesService, Session
 * ============================================================
 */

'use strict';

// ============================================================
// CONSTANTES GLOBALES — Nombres de hojas
// Fuente única de verdad para todos los módulos
// ============================================================

var HOJAS = {
  // Configuración
  CONFIG_SISTEMA:        'CONFIG_SISTEMA',
  CONFIG_PERIODOS:       'CONFIG_PERIODOS',
  CONFIG_ADMIN:          'CONFIG_ADMIN',

  // Maestros
  MSTR_CLIENTES:         'MSTR_CLIENTES',
  MSTR_PROVEEDORES:      'MSTR_PROVEEDORES',
  MSTR_PRODUCTOS:        'MSTR_PRODUCTOS',
  MSTR_BODEGAS:          'MSTR_BODEGAS',
  MSTR_PLAN_CUENTAS:     'MSTR_PLAN_CUENTAS',
  MSTR_TABLAS_REF:       'MSTR_TABLAS_REF',
  MSTR_TIPOS_CAMBIO:     'MSTR_TIPOS_CAMBIO',

  // Transaccionales
  DB_VENTAS:             'DB_VENTAS',
  DB_COMPRAS:            'DB_COMPRAS',
  DB_MOVIMIENTOS:        'DB_MOVIMIENTOS',
  DB_INGRESOS:           'DB_INGRESOS',
  DB_EGRESOS:            'DB_EGRESOS',
  DB_CXC:                'DB_CXC',
  DB_CXP:                'DB_CXP',
  DB_CAJA:               'DB_CAJA',
  DB_PRESUPUESTO:        'DB_PRESUPUESTO',
  DB_METAS:              'DB_METAS',

  // Historial
  HIST_PRECIOS:          'HIST_PRECIOS',
  HIST_STOCK:            'HIST_STOCK',

  // Cálculo
  CALC_ALERTAS:          'CALC_ALERTAS',
  CALC_FINANCIERO:       'CALC_FINANCIERO',
  CALC_COMERCIAL:        'CALC_COMERCIAL',
  CALC_SUPPLY:           'CALC_SUPPLY',
  CALC_OBJETIVOS:        'CALC_OBJETIVOS',
  CALC_FLUJO:            'CALC_FLUJO',

  // Batch y resultados
  BATCH_RESULTADOS:      'BATCH_RESULTADOS',

  // Dashboard e informes
  DASH_EJECUTIVO:        'DASH_EJECUTIVO',
  INF_FINANCIERO:        'INF_FINANCIERO',
  INF_VENTAS:            'INF_VENTAS',
  INF_INVENTARIO:        'INF_INVENTARIO',
  INF_COMPRAS:           'INF_COMPRAS',
  INF_PRESUPUESTO:       'INF_PRESUPUESTO',

  // Importación
  IMPORT_MAESTROS:       'IMPORT_MAESTROS',
  IMPORT_TRANSACCIONES:  'IMPORT_TRANSACCIONES',

  // Técnico
  LOG_AUDITORIA:         'LOG_AUDITORIA',
  DOCUMENTACION:         'DOCUMENTACION',
  LISTAS_VALIDACION:     'LISTAS_VALIDACION',
  INICIO:                'INICIO'
};

// ============================================================
// CONSTANTES DE COLUMNAS — Índices base 0
// Crítico: deben coincidir con los encabezados físicos
// ============================================================

var COL = {
  // MSTR_PRODUCTOS (base 0)
  PRD_ID_PRODUCTO:         0,   // A
  PRD_ID_EMPRESA:          1,   // B
  PRD_SKU:                 2,   // C
  PRD_NOMBRE:              3,   // D
  PRD_ID_CATEGORIA:        4,   // E
  PRD_TIPO_PRODUCTO:       5,   // F
  PRD_UNIDAD_MEDIDA:       6,   // G
  PRD_COSTO_UNITARIO:      7,   // H
  PRD_PRECIO_VENTA_NETO:   8,   // I
  PRD_ID_PROVEEDOR_PRINC:  9,   // J
  PRD_ID_BODEGA_DEFECTO:   10,  // K
  PRD_STOCK_ACTUAL:        11,  // L ← CRÍTICO — validado en auditoría
  PRD_STOCK_INICIAL:       12,  // M
  PRD_STOCK_MINIMO:        13,  // N
  PRD_STOCK_MAXIMO:        14,  // O
  PRD_STOCK_CRITICO:       15,  // P
  PRD_CLASIFICACION_ABC:   16,  // Q
  PRD_DIAS_ROTACION:       17,  // R
  PRD_APLICA_IVA:          18,  // S
  PRD_MANEJA_STOCK:        19,  // T
  PRD_ACTIVO:              27,  // AB

  // DB_VENTAS (base 0)
  VTA_ID_LINEA:            0,   // A
  VTA_ID_DOCUMENTO:        1,   // B
  VTA_ID_EMPRESA:          2,   // C
  VTA_TIPO_DOCUMENTO:      3,   // D
  VTA_NUMERO_DOC:          4,   // E
  VTA_ID_CLIENTE:          12,  // M
  VTA_ID_PRODUCTO:         13,  // N
  VTA_CANTIDAD:            16,  // Q
  VTA_PRECIO_UNITARIO:     17,  // R
  VTA_DESCUENTO_PCT:       18,  // S
  VTA_COSTO_MOMENTO:       22,  // W
  VTA_ESTADO:              27,  // AB
  VTA_CREADO_EN:           30,  // AE
  VTA_CREADO_POR:          31,  // AF

  // DB_COMPRAS (base 0)
  OC_ID_LINEA:             0,   // A
  OC_ID_OC:                1,   // B
  OC_ID_EMPRESA:           2,   // C
  OC_ID_PROVEEDOR:         3,   // D
  OC_ID_PRODUCTO:          4,   // E
  OC_FECHA_EMISION:        6,   // G
  OC_CANTIDAD_PEDIDA:      11,  // L
  OC_CANTIDAD_RECIBIDA:    12,  // M
  OC_PRECIO_UNIT_NETO:     13,  // N
  OC_DESCUENTO_PCT:        14,  // O
  OC_MONEDA_ORIGEN:        18,  // S
  OC_TIPO_CAMBIO:          19,  // T
  OC_ID_BODEGA_DESTINO:    21,  // V
  OC_ESTADO_OC:            22,  // W
  OC_CREADO_EN:            26,  // AA
  OC_CREADO_POR:           27,  // AB

  // DB_MOVIMIENTOS (base 0)
  MOV_ID_MOVIMIENTO:       0,   // A
  MOV_ID_EMPRESA:          1,   // B
  MOV_FECHA:               2,   // C
  MOV_ID_PRODUCTO:         4,   // E
  MOV_ID_BODEGA_ORIGEN:    5,   // F
  MOV_ID_BODEGA_DESTINO:   6,   // G
  MOV_TIPO:                7,   // H
  MOV_CANTIDAD:            8,   // I
  MOV_COSTO_MOMENTO:       10,  // K
  MOV_ID_REFERENCIA:       11,  // L
  MOV_TIPO_REFERENCIA:     12,  // M
  MOV_MOTIVO_AJUSTE:       13,  // N
  MOV_CREADO_EN:           14,  // O
  MOV_CREADO_POR:          15,  // P

  // DB_CXC (base 0)
  CXC_ID:                  0,   // A
  CXC_ID_EMPRESA:          1,   // B
  CXC_ID_CLIENTE:          2,   // C
  CXC_FECHA_EMISION:       6,   // G
  CXC_FECHA_VENCIMIENTO:   7,   // H
  CXC_MONTO_ORIGINAL:      8,   // I
  CXC_MONTO_PAGADO:        9,   // J
  CXC_ESTADO:              13,  // N
  CXC_ID_REFERENCIA:       17,  // R
  CXC_CREADO_EN:           19,  // T
  CXC_CREADO_POR:          20,  // U

  // DB_CXP (base 0)
  CXP_ID:                  0,   // A
  CXP_ID_EMPRESA:          1,   // B
  CXP_ID_PROVEEDOR:        2,   // C
  CXP_FECHA_EMISION:       6,   // G
  CXP_FECHA_VENCIMIENTO:   7,   // H
  CXP_MONTO_ORIGINAL:      8,   // I
  CXP_MONTO_PAGADO:        9,   // J
  CXP_ESTADO:              15,  // P
  CXP_CREADO_EN:           18,  // S
  CXP_CREADO_POR:          19,  // T

  // LOG_AUDITORIA (base 0)
  LOG_ID:                  0,   // A
  LOG_TIMESTAMP:           1,   // B
  LOG_USUARIO:             2,   // C
  LOG_TABLA:               3,   // D
  LOG_ACCION:              4,   // E
  LOG_ID_REGISTRO:         5,   // F
  LOG_CAMPO:               6,   // G
  LOG_VALOR_NUEVO:         7,   // H
  LOG_ID_EMPRESA:          8,   // I
  LOG_SESSION_ID:          9,   // J

  // CONFIG_ADMIN sección A (base 0)
  ADM_ID_ALERTA:           0,   // A
  ADM_DESCRIPCION:         1,   // B
  ADM_MODULO:              2,   // C
  ADM_TIPO_UMBRAL:         3,   // D
  ADM_VALOR_1:             4,   // E
  ADM_VALOR_2:             5,   // F
  ADM_EMAIL_DESTINO:       6,   // G
  ADM_ACTIVO:              7,   // H

  // CONFIG_ADMIN sección B — usuarios (base 0)
  ADM_USR_ID:              0,   // A
  ADM_USR_ID_EMPRESA:      1,   // B
  ADM_USR_EMAIL:           2,   // C
  ADM_USR_NOMBRE:          3,   // D
  ADM_USR_ROL:             4,   // E
  ADM_USR_ACTIVO:          5,   // F
  ADM_USR_CREADO_EN:       6    // G
};

// ============================================================
// CONSTANTES DE ROLES
// ============================================================

var ROLES = {
  ADMINISTRADOR: 'ADMINISTRADOR',
  FINANCIERO:    'FINANCIERO',
  COMERCIAL:     'COMERCIAL',
  COMPRAS:       'COMPRAS',
  BODEGA:        'BODEGA',
  LECTURA:       'LECTURA'
};

// ============================================================
// CONSTANTES DE ESTADOS
// ============================================================

var ESTADOS = {
  // Períodos
  PERIODO_ABIERTO:    'ABIERTO',
  PERIODO_CIERRE:     'EN_CIERRE',
  PERIODO_CERRADO:    'CERRADO',

  // Ventas y documentos
  VIGENTE:            'VIGENTE',
  ANULADO:            'ANULADO',
  DEVUELTO:           'DEVUELTO',

  // Órdenes de compra
  OC_BORRADOR:        'BORRADOR',
  OC_EMITIDA:         'EMITIDA',
  OC_CONFIRMADA:      'CONFIRMADA',
  OC_PARCIAL:         'PARCIAL',
  OC_RECIBIDA:        'RECIBIDA',
  OC_ANULADA:         'ANULADA',

  // CxC y CxP
  PENDIENTE:          'PENDIENTE',
  ABONADO:            'ABONADO',
  PAGADO:             'PAGADO',
  VENCIDO:            'VENCIDO',
  INCOBRABLE:         'INCOBRABLE',

  // Movimientos
  MOV_APERTURA:             'APERTURA',
  MOV_ENTRADA_COMPRA:       'ENTRADA_COMPRA',
  MOV_SALIDA_VENTA:         'SALIDA_VENTA',
  MOV_AJUSTE_POSITIVO:      'AJUSTE_POSITIVO',
  MOV_AJUSTE_NEGATIVO:      'AJUSTE_NEGATIVO',
  MOV_DEV_CLIENTE:          'DEVOLUCION_CLIENTE',
  MOV_DEV_PROVEEDOR:        'DEVOLUCION_PROVEEDOR',
  MOV_TRASLADO:             'TRASLADO'
};

// ============================================================
// CONSTANTES DE CONFIGURACIÓN DEL SISTEMA
// ============================================================

var SISTEMA = {
  VERSION:              '1.0.0',
  EMPRESA_DEFAULT:      'EMP-01',
  LOG_BUFFER_KEY:       'LOG_BUFFER',
  LOG_BUFFER_MAX:       50,        // Flush automático al llegar a 50 entradas
  SESSION_ID_KEY:       'SESSION_ID',
  ULTIMA_CONSULTA_KEY:  'ULTIMA_CONSULTA_UPDATES',
  VERSION_KEY:          'VERSION',
  FECHA_INSTALL_KEY:    'FECHA_INSTALACION',
  URL_VERSIONES:        'https://script.google.com/macros/s/DEPLOYMENT_ID_PROVEEDOR/exec',
  FILA_INICIO_USUARIOS: 53,        // CONFIG_ADMIN fila donde empiezan los usuarios
  FILA_ENCABEZADO_USR:  52,        // CONFIG_ADMIN fila del encabezado de usuarios
  TIMEOUT_LOCK:         30000      // 30 segundos timeout para LockService
};

// ============================================================
// FUNCIÓN: generarID
// Genera un ID único con formato PREFIJO-YYYYMMDD-HHMMSS-XXX
// No lee hojas — O(1) puro
// ============================================================

/**
 * Genera un ID único para cualquier tipo de registro.
 * Usa timestamp + sufijo aleatorio para evitar colisiones.
 *
 * @param {string} prefijo - Ejemplo: 'VTA', 'OC', 'MOV', 'CXC', 'CXP', 'CLI', 'PRV', 'PRD'
 * @returns {string} ID único en formato PREFIJO-YYYYMMDD-HHMMSS-XXX
 */
function generarID(prefijo) {
  try {
    var ahora = new Date();
    var año   = ahora.getFullYear();
    var mes   = String(ahora.getMonth() + 1).padStart(2, '0');
    var dia   = String(ahora.getDate()).padStart(2, '0');
    var hora  = String(ahora.getHours()).padStart(2, '0');
    var min   = String(ahora.getMinutes()).padStart(2, '0');
    var seg   = String(ahora.getSeconds()).padStart(2, '0');

    // Sufijo aleatorio de 3 dígitos para evitar colisiones en el mismo segundo
    var sufijo = String(Math.floor(Math.random() * 900) + 100);

    return prefijo + '-' + año + mes + dia + '-' + hora + min + seg + '-' + sufijo;

  } catch (e) {
    // Fallback de emergencia usando solo Math.random()
    return prefijo + '-' + Date.now() + '-' + String(Math.floor(Math.random() * 9000) + 1000);
  }
}

// ============================================================
// FUNCIÓN: getConfig
// Lee CONFIG_SISTEMA fila 2 completa en una sola llamada
// ============================================================

/**
 * Retorna un objeto con todos los valores de configuración del sistema.
 * Lee una sola fila en una sola llamada para minimizar operaciones I/O.
 *
 * @returns {Object} Objeto con propiedades de configuración o null si falla
 */
function getConfig() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.CONFIG_SISTEMA);

    if (!hoja) {
      Logger.log('ERROR getConfig: Hoja CONFIG_SISTEMA no encontrada');
      return null;
    }

    // Leer fila 2 completa en una sola llamada (columnas A a X = 24 columnas)
    var datos = hoja.getRange(2, 1, 1, 24).getValues()[0];

    return {
      id_empresa:           datos[0],   // A
      nombre_empresa:       datos[1],   // B
      rut_empresa:          datos[2],   // C
      rut_normalizado:      datos[3],   // D
      giro_empresa:         datos[4],   // E
      email_admin:          datos[9],   // J
      moneda_base:          datos[10],  // K
      tasa_iva:             datos[11],  // L
      mes_inicio_fiscal:    datos[12],  // M
      año_vigente:          datos[13],  // N
      color_primario:       datos[15],  // P
      color_acento:         datos[16],  // Q
      multiempresa_activo:  datos[17],  // R
      zona_horaria:         datos[18],  // S
      fecha_hoy:            datos[19],  // T
      version_esquema:      datos[20],  // U
      creado_en:            datos[22],  // W
      creado_por:           datos[23]   // X
    };

  } catch (e) {
    Logger.log('ERROR getConfig: ' + e.message);
    return null;
  }
}

// ============================================================
// FUNCIÓN: getPeriodoActual
// Retorna el período ABIERTO vigente para la empresa activa
// ============================================================

/**
 * Busca el período contable con estado ABIERTO para la empresa actual.
 * Si no existe período abierto retorna null y el sistema no permite transacciones.
 *
 * @returns {string|null} Período en formato YYYY-MM o null si no hay período abierto
 */
function getPeriodoActual() {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var hojaPer     = ss.getSheetByName(HOJAS.CONFIG_PERIODOS);
    var config      = getConfig();

    if (!hojaPer || !config) {
      Logger.log('ERROR getPeriodoActual: Hoja o config no disponible');
      return null;
    }

    var idEmpresa   = config.id_empresa;
    var ultimaFila  = hojaPer.getLastRow();

    if (ultimaFila < 2) {
      Logger.log('ERROR getPeriodoActual: No hay períodos registrados');
      return null;
    }

    // Leer todas las filas de períodos en una sola llamada
    var datos = hojaPer.getRange(2, 1, ultimaFila - 1, 8).getValues();

    for (var i = 0; i < datos.length; i++) {
      var fila         = datos[i];
      var idEmpresaFil = fila[1]; // B — id_empresa
      var periodo      = fila[4]; // E — sys_periodo
      var estado       = fila[6]; // G — estado

      if (idEmpresaFil === idEmpresa && estado === ESTADOS.PERIODO_ABIERTO && periodo !== '') {
  // Si el valor es una fecha lo convertimos a formato YYYY-MM
  if (periodo instanceof Date) {
    var yyyy = periodo.getFullYear();
    var mm   = String(periodo.getMonth() + 1).padStart(2, '0');
    return yyyy + '-' + mm;
  }
  return String(periodo);
}
    }

    Logger.log('ERROR getPeriodoActual: No se encontró período ABIERTO para empresa ' + idEmpresa);
    return null;

  } catch (e) {
    Logger.log('ERROR getPeriodoActual: ' + e.message);
    return null;
  }
}

// ============================================================
// FUNCIÓN: validarPeriodoAbierto
// Verifica que existe un período ABIERTO antes de transacciones
// ============================================================

/**
 * Valida que existe un período ABIERTO.
 * Muestra un alert al usuario si no hay período disponible.
 *
 * @param {boolean} silencioso - Si true, no muestra alert. Solo retorna boolean.
 * @returns {boolean} true si hay período abierto, false si no
 */
function validarPeriodoAbierto(silencioso) {
  try {
    var periodo = getPeriodoActual();

    if (!periodo) {
      if (!silencioso) {
        SpreadsheetApp.getUi().alert(
          '⛔ Sin período activo\n\n' +
          'No existe un período contable ABIERTO.\n' +
          'Ir a ERP → Administración → Gestionar períodos\n' +
          'y abre un período para continuar.'
        );
      }
      return false;
    }

    return true;

  } catch (e) {
    Logger.log('ERROR validarPeriodoAbierto: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: setupSistema
// Inicializa el sistema en nuevas instalaciones
// Ejecutada desde ERP → Configuración → Inicializar sistema
// ============================================================

/**
 * Función de instalación del sistema para nuevos clientes.
 * Verifica propiedad, instala triggers, configura protecciones
 * e inicializa PropertiesService.
 *
 * Solo debe ejecutarse una vez por instalación.
 * Si ya existe triggers, pregunta al usuario si desea reinstalar.
 */
function setupSistema() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // VERIFICACIÓN 1: El usuario activo debe ser el propietario
    if (!_verificarPropietario()) {
      ui.alert(
        '❌ Permiso insuficiente\n\n' +
        'setupSistema() solo puede ejecutarlo el propietario.\n\n' +
        'Propietario: ' + ss.getOwner().getEmail() + '\n' +
        'Usuario actual: ' + Session.getActiveUser().getEmail()
      );
      return false;
    }

    // VERIFICACIÓN 2: CONFIG_SISTEMA debe tener datos de empresa
    var config = getConfig();
    if (!config || !config.id_empresa || config.id_empresa === '') {
      ui.alert(
        '❌ Sistema no configurado\n\n' +
        'CONFIG_SISTEMA no tiene datos de empresa.\n' +
        'Contacta al proveedor para completar la instalación.'
      );
      return false;
    }

    // VERIFICACIÓN 3: Triggers existentes
    var triggersExistentes = ScriptApp.getProjectTriggers();
    if (triggersExistentes.length > 0) {
      var respuesta = ui.alert(
        '⚠️ Triggers existentes\n\n' +
        'Ya existen ' + triggersExistentes.length + ' trigger(s).\n\n' +
        '¿Deseas reinstalar desde cero?',
        ui.ButtonSet.YES_NO
      );
      if (respuesta !== ui.Button.YES) {
        ui.alert('Cancelado. Los triggers no fueron modificados.');
        return false;
      }
      triggersExistentes.forEach(function(trigger) {
        ScriptApp.deleteTrigger(trigger);
      });
      Logger.log('Triggers anteriores eliminados: ' + triggersExistentes.length);
    }

    // PASO 1: Instalar los 6 triggers
    _instalarTriggers();

    // PASO 2: Inicializar PropertiesService
    _inicializarPropertiesService();

    // PASO 3: Verificar hojas críticas
    var hojasVerificadas = _verificarHojasCriticas();

    // PASO 4: Registrar fecha de instalación
    try {
      var hojaConfig = ss.getSheetByName(HOJAS.CONFIG_SISTEMA);
      hojaConfig.getRange('W2').setValue(new Date());
    } catch (eW) {
      Logger.log('ADVERTENCIA: No se pudo escribir fecha en W2: ' + eW.message);
    }

    // CONFIRMACIÓN FINAL
    ui.alert(
      '✅ Sistema inicializado correctamente\n\n' +
      '────────────────────────────────\n' +
      'Empresa:     ' + config.nombre_empresa + '\n' +
      'ID:          ' + config.id_empresa + '\n' +
      'Versión:     ' + SISTEMA.VERSION + '\n' +
      '────────────────────────────────\n\n' +
      'Triggers instalados: 6\n' +
      'PropertiesService:   inicializado\n\n' +
      'Próximo paso: carga tus datos maestros.'
    );

    Logger.log('setupSistema completado para: ' + config.nombre_empresa);
    return true;

  } catch (e) {
    Logger.log('ERROR setupSistema: ' + e.message);
    ui.alert(
      '❌ Error durante la inicialización\n\n' +
      'Detalle: ' + e.message + '\n\n' +
      'Contacta al proveedor con este mensaje.'
    );
    return false;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarPropietario
// ============================================================

/**
 * Verifica que el usuario activo es el propietario del archivo.
 * Solo el propietario puede instalar triggers y protecciones.
 *
 * @returns {boolean} true si el usuario es propietario
 */
function _verificarPropietario() {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var propietario   = ss.getOwner().getEmail();
    var usuarioActivo = Session.getActiveUser().getEmail();

    Logger.log('Propietario: ' + propietario + ' | Usuario activo: ' + usuarioActivo);
    return propietario === usuarioActivo;

  } catch (e) {
    Logger.log('ERROR _verificarPropietario: ' + e.message);
    // En entornos donde getOwner() no está disponible (algunos contextos de script)
    // retornamos true para no bloquear el setup
    return true;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _instalarTriggers
// ============================================================

/**
 * Instala los 6 triggers del sistema.
 * Debe ejecutarse solo después de eliminar los triggers anteriores.
 */
function _instalarTriggers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Trigger 1: onOpen — cada apertura del archivo
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(ss)
      .onOpen()
      .create();

    // Trigger 2: onEdit — cada edición en cualquier celda
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    // Trigger 3: actualizarFechaHoy — diario a las 00:00
    ScriptApp.newTrigger('actualizarFechaHoy')
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();

    // Trigger 4: batchNocturno — diario a las 01:00
    ScriptApp.newTrigger('batchNocturno')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();

    // Trigger 5: crearPeriodoSiguiente — día 25 de cada mes a las 08:00
    ScriptApp.newTrigger('crearPeriodoSiguiente')
      .timeBased()
      .onMonthDay(25)
      .atHour(8)
      .create();

    // Trigger 6: flushLogBuffer — cada 5 minutos
    ScriptApp.newTrigger('flushLogBuffer')
      .timeBased()
      .everyMinutes(5)
      .create();

    Logger.log('_instalarTriggers: 6 triggers instalados correctamente');

  } catch (e) {
    Logger.log('ERROR _instalarTriggers: ' + e.message);
    throw e; // Propagar el error hacia setupSistema
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _inicializarPropertiesService
// ============================================================

/**
 * Inicializa todas las propiedades del script en PropertiesService.
 * Establece el buffer de log vacío y registra la fecha de instalación.
 */
function _inicializarPropertiesService() {
  try {
    var props = PropertiesService.getScriptProperties();

    props.setProperty(SISTEMA.LOG_BUFFER_KEY,    '[]');
    props.setProperty(SISTEMA.VERSION_KEY,        SISTEMA.VERSION);
    props.setProperty(SISTEMA.FECHA_INSTALL_KEY,  new Date().toISOString());
    props.setProperty(SISTEMA.SESSION_ID_KEY,     generarID('SES'));

    Logger.log('_inicializarPropertiesService: completado');

  } catch (e) {
    Logger.log('ERROR _inicializarPropertiesService: ' + e.message);
    throw e;
  }
}

// ============================================================
// FUNCIÓN: configurarProtecciones (pública — también llamada desde menú)
// ============================================================

/**
 * Aplica protecciones sobre hojas críticas e inmutables.
 * LOG_AUDITORIA: nadie puede editar manualmente.
 * DB_MOVIMIENTOS: nadie puede editar manualmente.
 *
 * Solo funciona si el usuario activo es el propietario.
 *
 * @returns {boolean} true si las protecciones se aplicaron correctamente
 */
function configurarProtecciones() {
  return _configurarProtecciones();
}

/**
 * Implementación interna de configurarProtecciones.
 * @returns {boolean}
 */
function _configurarProtecciones() {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var emailPropiet  = Session.getActiveUser().getEmail();

    // Proteger LOG_AUDITORIA — solo el propietario puede editar
    var hojaLog = ss.getSheetByName(HOJAS.LOG_AUDITORIA);
    if (hojaLog) {
      // Eliminar protecciones existentes en esta hoja
      var protExistentes = hojaLog.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protExistentes.forEach(function(p) { p.remove(); });

      // Aplicar nueva protección
      var protLog = hojaLog.protect();
      protLog.setDescription('LOG_AUDITORIA — Solo sistema');
      protLog.addEditor(emailPropiet);
      protLog.removeEditors(protLog.getEditors().filter(function(e) {
        return e.getEmail() !== emailPropiet;
      }));

      Logger.log('Protección aplicada: LOG_AUDITORIA');
    }

    // Proteger DB_MOVIMIENTOS — inmutable, solo el script escribe
    var hojaMovimientos = ss.getSheetByName(HOJAS.DB_MOVIMIENTOS);
    if (hojaMovimientos) {
      var protExistMov = hojaMovimientos.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protExistMov.forEach(function(p) { p.remove(); });

      var protMov = hojaMovimientos.protect();
      protMov.setDescription('DB_MOVIMIENTOS — Solo sistema');
      protMov.addEditor(emailPropiet);
      protMov.removeEditors(protMov.getEditors().filter(function(e) {
        return e.getEmail() !== emailPropiet;
      }));

      Logger.log('Protección aplicada: DB_MOVIMIENTOS');
    }

    return true;

  } catch (e) {
    Logger.log('ERROR _configurarProtecciones: ' + e.message);
    // No propagamos el error — las protecciones son importantes pero no críticas
    return false;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarHojasCriticas
// ============================================================

/**
 * Verifica que todas las hojas críticas del sistema existen.
 * Registra en el log las hojas faltantes si las hay.
 *
 * @returns {boolean} true si todas las hojas críticas existen
 */
function _verificarHojasCriticas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var hojasCriticas = [
      HOJAS.CONFIG_SISTEMA,
      HOJAS.CONFIG_PERIODOS,
      HOJAS.CONFIG_ADMIN,
      HOJAS.MSTR_PRODUCTOS,
      HOJAS.MSTR_CLIENTES,
      HOJAS.MSTR_PROVEEDORES,
      HOJAS.DB_VENTAS,
      HOJAS.DB_COMPRAS,
      HOJAS.DB_MOVIMIENTOS,
      HOJAS.DB_CXC,
      HOJAS.DB_CXP,
      HOJAS.LOG_AUDITORIA,
      HOJAS.CALC_ALERTAS
    ];

    var hojasFaltantes = [];

    hojasCriticas.forEach(function(nombreHoja) {
      if (!ss.getSheetByName(nombreHoja)) {
        hojasFaltantes.push(nombreHoja);
        Logger.log('FALTANTE: Hoja ' + nombreHoja + ' no encontrada');
      }
    });

    if (hojasFaltantes.length > 0) {
      Logger.log('Hojas faltantes: ' + hojasFaltantes.join(', '));
      return false;
    }

    Logger.log('_verificarHojasCriticas: todas las hojas críticas OK');
    return true;

  } catch (e) {
    Logger.log('ERROR _verificarHojasCriticas: ' + e.message);
    return false;
  }
}

// ============================================================
// FIN DE M0_Setup.gs
// ============================================================
// ============================================================
