/**
 * ============================================================
 * M1_Sesion.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Maneja la sesión del usuario. Construye
 * el menú ERP según el rol. Controla visibilidad de hojas.
 * Es el punto de entrada del sistema en cada apertura.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Módulos:      M0_Setup, M4_Auditoria
 *   Hojas:        CONFIG_ADMIN
 *   APIs Google:  Session, SpreadsheetApp.getUi()
 *   Triggers:     onOpen (instalado por M0)
 * ============================================================
 */

// ============================================================
// FUNCIÓN: onOpen
// Trigger principal — se ejecuta cada vez que se abre el archivo
// ============================================================

/**
 * Punto de entrada del sistema.
 * Se ejecuta automáticamente cuando cualquier usuario abre el archivo.
 * Detecta el rol del usuario y construye el menú correspondiente.
 */
function onOpen() {
  try {
    // Obtener email del usuario activo
    var email = _getUsuarioActivo();
    Logger.log('onOpen: Usuario activo: ' + email);

    // Detectar rol del usuario
    var rol = getRolUsuario(email);
    Logger.log('onOpen: Rol detectado: ' + rol);

    // Aplicar visibilidad de hojas según rol
    aplicarVisibilidadHojas(rol);

    // Construir menú según rol
    construirMenu(rol);

    // Verificar actualizaciones silenciosamente (máximo 1 vez por semana)
    _verificarActualizacionesSemanal();

    Logger.log('onOpen: Sistema listo para: ' + email + ' [' + rol + ']');

  } catch (e) {
    Logger.log('ERROR onOpen: ' + e.message);
    // Aunque falle el onOpen construimos el menú mínimo
    // para que el usuario pueda acceder a configuración
    try {
      _construirMenuMinimo();
    } catch (e2) {
      Logger.log('ERROR onOpen menu mínimo: ' + e2.message);
    }
  }
}

// ============================================================
// FUNCIÓN: getRolUsuario
// Busca el email en CONFIG_ADMIN y retorna el rol asignado
// ============================================================

/**
 * Busca el email del usuario en CONFIG_ADMIN sección B
 * y retorna el rol que tiene asignado.
 *
 * Si el email no está registrado retorna LECTURA por defecto.
 * Si CONFIG_ADMIN no existe o falla retorna LECTURA por seguridad.
 *
 * @param {string} email - Email del usuario a buscar
 * @returns {string} Rol del usuario: ADMINISTRADOR, FINANCIERO,
 *                   COMERCIAL, COMPRAS, BODEGA o LECTURA
 */
function getRolUsuario(email) {
  try {
    if (!email || email === '') {
      Logger.log('getRolUsuario: Email vacío — asignando LECTURA');
      return ROLES.LECTURA;
    }

    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var hojaAdmin = ss.getSheetByName(HOJAS.CONFIG_ADMIN);

    if (!hojaAdmin) {
      Logger.log('getRolUsuario: CONFIG_ADMIN no encontrada — asignando LECTURA');
      return ROLES.LECTURA;
    }

    var ultimaFila = hojaAdmin.getLastRow();

    // Los usuarios empiezan en SISTEMA.FILA_INICIO_USUARIOS (53)
    if (ultimaFila < SISTEMA.FILA_INICIO_USUARIOS) {
      Logger.log('getRolUsuario: No hay usuarios registrados — asignando LECTURA');
      return ROLES.LECTURA;
    }

    var filas     = ultimaFila - SISTEMA.FILA_INICIO_USUARIOS + 1;
    var datos     = hojaAdmin.getRange(
      SISTEMA.FILA_INICIO_USUARIOS, 1, filas, 7
    ).getValues();

    // Buscar el email en la columna C (índice 2)
    for (var i = 0; i < datos.length; i++) {
      var fila          = datos[i];
      var emailFila     = String(fila[COL.ADM_USR_EMAIL]).trim().toLowerCase();
      var activoFila    = fila[COL.ADM_USR_ACTIVO];
      var rolFila       = String(fila[COL.ADM_USR_ROL]).trim();

      if (emailFila === email.toLowerCase() && activoFila === true) {
        Logger.log('getRolUsuario: ' + email + ' → ' + rolFila);
        return rolFila || ROLES.LECTURA;
      }
    }

    Logger.log('getRolUsuario: ' + email + ' no encontrado — asignando LECTURA');
    return ROLES.LECTURA;

  } catch (e) {
    Logger.log('ERROR getRolUsuario: ' + e.message);
    return ROLES.LECTURA;
  }
}

// ============================================================
// FUNCIÓN: aplicarVisibilidadHojas
// Muestra u oculta hojas según el rol del usuario
// ============================================================

/**
 * Controla qué hojas puede ver el usuario según su rol.
 * Las hojas técnicas y de configuración solo las ve el ADMINISTRADOR.
 *
 * @param {string} rol - Rol del usuario activo
 */
function aplicarVisibilidadHojas(rol) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Definir hojas visibles por rol
    var hojasSiempreVisibles = [
      HOJAS.INICIO,
      HOJAS.DASH_EJECUTIVO
    ];

    var hojasPorRol = {
      ADMINISTRADOR: [
        HOJAS.CONFIG_SISTEMA, HOJAS.CONFIG_PERIODOS, HOJAS.CONFIG_ADMIN,
        HOJAS.MSTR_CLIENTES, HOJAS.MSTR_PROVEEDORES, HOJAS.MSTR_PRODUCTOS,
        HOJAS.MSTR_BODEGAS, HOJAS.MSTR_PLAN_CUENTAS, HOJAS.MSTR_TABLAS_REF,
        HOJAS.MSTR_TIPOS_CAMBIO,
        HOJAS.DB_VENTAS, HOJAS.DB_COMPRAS, HOJAS.DB_MOVIMIENTOS,
        HOJAS.DB_INGRESOS, HOJAS.DB_EGRESOS, HOJAS.DB_CXC, HOJAS.DB_CXP,
        HOJAS.DB_CAJA, HOJAS.DB_PRESUPUESTO, HOJAS.DB_METAS,
        HOJAS.HIST_PRECIOS, HOJAS.HIST_STOCK,
        HOJAS.CALC_FINANCIERO, HOJAS.CALC_COMERCIAL, HOJAS.CALC_SUPPLY,
        HOJAS.CALC_OBJETIVOS, HOJAS.CALC_FLUJO, HOJAS.CALC_ALERTAS,
        HOJAS.BATCH_RESULTADOS,
        HOJAS.INF_FINANCIERO, HOJAS.INF_VENTAS, HOJAS.INF_INVENTARIO,
        HOJAS.INF_COMPRAS, HOJAS.INF_PRESUPUESTO,
        HOJAS.LOG_AUDITORIA, HOJAS.DOCUMENTACION
      ],
      FINANCIERO: [
        HOJAS.DB_INGRESOS, HOJAS.DB_EGRESOS, HOJAS.DB_CXC, HOJAS.DB_CXP,
        HOJAS.DB_CAJA, HOJAS.DB_PRESUPUESTO, HOJAS.DB_METAS,
        HOJAS.MSTR_CLIENTES, HOJAS.MSTR_PROVEEDORES, HOJAS.MSTR_PLAN_CUENTAS,
        HOJAS.INF_FINANCIERO, HOJAS.INF_PRESUPUESTO,
        HOJAS.CALC_FINANCIERO, HOJAS.CALC_FLUJO, HOJAS.CALC_OBJETIVOS
      ],
      COMERCIAL: [
        HOJAS.DB_VENTAS, HOJAS.DB_CXC,
        HOJAS.MSTR_CLIENTES, HOJAS.MSTR_PRODUCTOS,
        HOJAS.INF_VENTAS,
        HOJAS.CALC_COMERCIAL
      ],
      COMPRAS: [
        HOJAS.DB_COMPRAS, HOJAS.DB_CXP,
        HOJAS.MSTR_PROVEEDORES, HOJAS.MSTR_PRODUCTOS,
        HOJAS.INF_COMPRAS,
        HOJAS.CALC_SUPPLY,
        HOJAS.HIST_PRECIOS
      ],
      BODEGA: [
        HOJAS.DB_MOVIMIENTOS,
        HOJAS.MSTR_PRODUCTOS, HOJAS.MSTR_BODEGAS,
        HOJAS.INF_INVENTARIO,
        HOJAS.HIST_STOCK
      ],
      LECTURA: [
        HOJAS.INF_FINANCIERO, HOJAS.INF_VENTAS, HOJAS.INF_INVENTARIO,
        HOJAS.INF_COMPRAS, HOJAS.INF_PRESUPUESTO
      ]
    };

    // Obtener lista de hojas permitidas para este rol
    var hojasPermitidas = hojasSiempreVisibles.concat(
      hojasPorRol[rol] || hojasPorRol[ROLES.LECTURA]
    );

    // Aplicar visibilidad a cada hoja
    var todasLasHojas = ss.getSheets();
    todasLasHojas.forEach(function(hoja) {
      var nombre = hoja.getName();

      // LISTAS_VALIDACION siempre oculta — es hoja técnica
      if (nombre === HOJAS.LISTAS_VALIDACION) {
        if (!hoja.isSheetHidden()) hoja.hideSheet();
        return;
      }

      // Hojas de importación solo para ADMINISTRADOR
      if (nombre === HOJAS.IMPORT_MAESTROS ||
          nombre === HOJAS.IMPORT_TRANSACCIONES) {
        if (rol !== ROLES.ADMINISTRADOR) {
          if (!hoja.isSheetHidden()) hoja.hideSheet();
        } else {
          if (hoja.isSheetHidden()) hoja.showSheet();
        }
        return;
      }

      // Aplicar visibilidad según permisos del rol
      if (hojasPermitidas.indexOf(nombre) !== -1) {
        if (hoja.isSheetHidden()) hoja.showSheet();
      } else {
        if (!hoja.isSheetHidden()) hoja.hideSheet();
      }
    });

    Logger.log('aplicarVisibilidadHojas: Visibilidad aplicada para rol ' + rol);

  } catch (e) {
    Logger.log('ERROR aplicarVisibilidadHojas: ' + e.message);
    // No detener el flujo si falla la visibilidad
  }
}

// ============================================================
// FUNCIÓN: construirMenu
// Construye el menú ERP según el rol del usuario
// ============================================================

/**
 * Construye el menú ERP en la barra de herramientas de Sheets.
 * Cada rol ve solo las opciones que le corresponden.
 *
 * @param {string} rol - Rol del usuario activo
 */
function construirMenu(rol) {
  try {
    var ui   = SpreadsheetApp.getUi();
    var menu = ui.createMenu('🏢 ERP');

    // ── VANTIS ERP ──────────────────────────────────────────
    menu.addItem('🚀 Abrir VANTIS ERP', 'abrirVantis');
    menu.addSeparator();

    // ── VENTAS ──────────────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR || rol === ROLES.COMERCIAL) {
      menu.addSubMenu(
        ui.createMenu('Ventas')
          .addItem('Nueva venta manual', 'abrirFormularioVenta')
          .addItem('Registrar pago CxC', 'registrarPagoCxC')
          .addItem('Ver CxC pendiente', 'mostrarCxCPendiente')
          .addSeparator()
          .addItem('Anular documento', 'procesarAnulacionVenta')
      );
    }

    // ── COMPRAS ─────────────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR || rol === ROLES.COMPRAS) {
      menu.addSubMenu(
        ui.createMenu('Compras')
          .addItem('Nueva orden de compra', 'abrirFormularioOC')
          .addItem('Registrar recepción', 'procesarRecepcionOC')
          .addItem('Registrar pago CxP', 'registrarPagoCxP')
          .addItem('Ver CxP pendiente', 'mostrarCxPPendiente')
          .addSeparator()
          .addItem('Anular orden de compra', 'procesarAnulacionOC')
      );
    }

    // ── INVENTARIO ──────────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR ||
        rol === ROLES.BODEGA ||
        rol === ROLES.COMPRAS) {
      menu.addSubMenu(
        ui.createMenu('Inventario')
          .addItem('Ajuste de stock', 'procesarAjusteStock')
          .addItem('Traslado entre bodegas', 'procesarTraslado')
          .addItem('Ver stock crítico', 'mostrarStockCritico')
      );
    }

    // ── FINANZAS ────────────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR || rol === ROLES.FINANCIERO) {
      menu.addSubMenu(
        ui.createMenu('Finanzas')
          .addItem('Registrar ingreso', 'registrarIngreso')
          .addItem('Registrar egreso', 'registrarEgreso')
          .addItem('Registrar pago CxC', 'registrarPagoCxC')
          .addItem('Registrar pago CxP', 'registrarPagoCxP')
      );
    }

    // ── ADMINISTRACIÓN ──────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR) {
      menu.addSubMenu(
        ui.createMenu('Administración')
          .addItem('Gestionar clientes', 'abrirMaestroClientes')
          .addItem('Gestionar proveedores', 'abrirMaestroProveedores')
          .addItem('Gestionar productos', 'abrirMaestroProductos')
          .addItem('Gestionar bodegas', 'abrirMaestroBodegas')
          .addSeparator()
          .addItem('Gestionar usuarios', 'abrirGestionUsuarios')
          .addItem('Gestionar períodos', 'abrirGestionPeriodos')
      );
    }

    // ── CONFIGURACIÓN ────────────────────────────────────────
    if (rol === ROLES.ADMINISTRADOR) {
      menu.addSubMenu(
        ui.createMenu('Configuración')
          .addItem('Inicializar sistema', 'setupSistema')
          .addItem('Verificar actualizaciones', 'checkForUpdates')
          .addItem('Ejecutar batch manual', 'batchNocturno')
          .addItem('Vaciar buffer de log', 'flushLogBuffer')
          .addItem('Limpiar log antiguo', 'limpiarLogAntiguoDesdeMenu')
          .addSeparator()
          .addItem('Configurar protecciones', 'configurarProtecciones')
      );
    }

    // ── AYUDA — visible para todos los roles ─────────────────
    menu.addSeparator();
    menu.addItem('ℹ️ Información del sistema', 'mostrarInfoSistema');

    menu.addToUi();
    Logger.log('construirMenu: Menú construido para rol ' + rol);

  } catch (e) {
    Logger.log('ERROR construirMenu: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: mostrarInfoSistema
// Muestra versión instalada, empresa y fecha de instalación
// ============================================================

/**
 * Muestra un alert con la información del sistema.
 * Disponible para todos los roles desde el menú ERP.
 */
function mostrarInfoSistema() {
  try {
    var config  = getConfig();
    var props   = PropertiesService.getScriptProperties();
    var version = props.getProperty(SISTEMA.VERSION_KEY) || SISTEMA.VERSION;
    var fechaIn = props.getProperty(SISTEMA.FECHA_INSTALL_KEY) || 'no registrada';
    var buffer  = getEstadoBuffer();
    var email   = _getUsuarioActivo();
    var rol     = getRolUsuario(email);

    SpreadsheetApp.getUi().alert(
      'ℹ️ Información del sistema\n\n' +
      '────────────────────────────────\n' +
      'Empresa:            ' + (config ? config.nombre_empresa : 'no disponible') + '\n' +
      'RUT:                ' + (config ? config.rut_empresa : 'no disponible') + '\n' +
      'Versión:            ' + version + '\n' +
      'Fecha instalación:  ' + fechaIn + '\n' +
      '────────────────────────────────\n' +
      'Usuario actual:     ' + email + '\n' +
      'Rol asignado:       ' + rol + '\n' +
      '────────────────────────────────\n' +
      'Buffer de log:      ' + buffer.entradas + '/' + buffer.maximo + ' entradas\n' +
      '────────────────────────────────\n\n' +
      'Soporte: contactar al proveedor del sistema.'
    );

  } catch (e) {
    Logger.log('ERROR mostrarInfoSistema: ' + e.message);
    SpreadsheetApp.getUi().alert('Error al obtener información del sistema: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: checkForUpdates
// Verifica si hay una versión más reciente disponible
// ============================================================

/**
 * Consulta el servidor del proveedor para verificar actualizaciones.
 * Puede llamarse manualmente desde el menú o automáticamente
 * una vez por semana desde onOpen().
 *
 * @param {boolean} silencioso - Si true no muestra alert cuando está actualizado
 */
function checkForUpdates(silencioso) {
  try {
    var props       = PropertiesService.getScriptProperties();
    var versionAct  = props.getProperty(SISTEMA.VERSION_KEY) || SISTEMA.VERSION;

    // Verificar si ya consultamos hoy
    var ultimaConsulta = props.getProperty(SISTEMA.ULTIMA_CONSULTA_KEY);
    var hoy            = new Date().toDateString();

    if (silencioso && ultimaConsulta === hoy) {
      Logger.log('checkForUpdates: Ya se consultó hoy — omitiendo');
      return null;
    }

    // Si la URL del servidor no está configurada, no consultar
    if (SISTEMA.URL_VERSIONES.indexOf('DEPLOYMENT_ID_PROVEEDOR') !== -1) {
      Logger.log('checkForUpdates: URL del servidor no configurada — omitiendo');
      if (!silencioso) {
        SpreadsheetApp.getUi().alert(
          'ℹ️ Verificación de actualizaciones\n\n' +
          'Versión instalada: ' + versionAct + '\n\n' +
          'El servidor de actualizaciones no está configurado.\n' +
          'Contacta al proveedor para verificar manualmente.'
        );
      }
      return null;
    }

    // Consultar servidor del proveedor
    var resp = UrlFetchApp.fetch(SISTEMA.URL_VERSIONES + '?action=version', {
      muteHttpExceptions: true,
      deadline: 10
    });

    props.setProperty(SISTEMA.ULTIMA_CONSULTA_KEY, hoy);

    if (resp.getResponseCode() !== 200) {
      Logger.log('checkForUpdates: Servidor no disponible — código ' + resp.getResponseCode());
      return null;
    }

    var data = JSON.parse(resp.getContentText());

    // Comparar versiones
    var hayActualizacion = _versionEsMenor(versionAct, data.version);

    if (!hayActualizacion) {
      if (!silencioso) {
        SpreadsheetApp.getUi().alert(
          '✅ Sistema actualizado\n\n' +
          'Versión instalada: ' + versionAct + '\n' +
          'No hay actualizaciones disponibles.'
        );
      }
      return { actualizado: true };
    }

    // Hay actualización disponible
    SpreadsheetApp.getUi().alert(
      (data.urgente ? '🔴 ACTUALIZACIÓN URGENTE\n\n' : '📦 Actualización disponible\n\n') +
      'Versión instalada:   ' + versionAct + '\n' +
      'Versión disponible:  ' + data.version + '  [' + data.tipo + ']\n' +
      'Fecha:               ' + data.fecha_release + '\n\n' +
      'Qué incluye:\n' + data.changelog + '\n\n' +
      'Contactar al proveedor para aplicar la actualización.'
    );

    registrarLog('SISTEMA', 'ACTUALIZAR', 'VERSION', 'version_disponible', data.version);
    return { actualizado: false, versionDisponible: data.version };

  } catch (e) {
    Logger.log('ERROR checkForUpdates: ' + e.message);
    if (!silencioso) {
      SpreadsheetApp.getUi().alert('No se pudo verificar actualizaciones: ' + e.message);
    }
    return null;
  }
}

// ============================================================
// FUNCIONES DE NAVEGACIÓN A MAESTROS
// Abren la hoja correspondiente y posicionan el cursor
// ============================================================

function abrirMaestroClientes() {
  _navegarAHoja(HOJAS.MSTR_CLIENTES, 'Maestro de Clientes');
}

function abrirMaestroProveedores() {
  _navegarAHoja(HOJAS.MSTR_PROVEEDORES, 'Maestro de Proveedores');
}

function abrirMaestroProductos() {
  _navegarAHoja(HOJAS.MSTR_PRODUCTOS, 'Maestro de Productos');
}

function abrirMaestroBodegas() {
  _navegarAHoja(HOJAS.MSTR_BODEGAS, 'Maestro de Bodegas');
}

function abrirGestionUsuarios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJAS.CONFIG_ADMIN);
  if (hoja) {
    ss.setActiveSheet(hoja);
    hoja.setActiveRange(hoja.getRange(SISTEMA.FILA_INICIO_USUARIOS, 1));
    SpreadsheetApp.getUi().alert(
      'ℹ️ Gestión de usuarios\n\n' +
      'Los usuarios se registran desde la fila ' + SISTEMA.FILA_INICIO_USUARIOS + '.\n\n' +
      'Columnas: id_usuario | id_empresa | email | nombre | rol | activo | creado_en\n\n' +
      'Roles disponibles:\n' +
      '• ADMINISTRADOR\n• FINANCIERO\n• COMERCIAL\n• COMPRAS\n• BODEGA\n• LECTURA'
    );
  }
}

function abrirGestionPeriodos() {
  _navegarAHoja(HOJAS.CONFIG_PERIODOS, 'Gestión de Períodos');
}

function abrirFormularioVenta() {
  _navegarAHoja(HOJAS.DB_VENTAS, 'Registro de Ventas');
}

function abrirFormularioOC() {
  _navegarAHoja(HOJAS.DB_COMPRAS, 'Registro de Órdenes de Compra');
}

function mostrarCxCPendiente() {
  _navegarAHoja(HOJAS.DB_CXC, 'Cuentas por Cobrar');
}

function mostrarCxPPendiente() {
  _navegarAHoja(HOJAS.DB_CXP, 'Cuentas por Pagar');
}

function mostrarStockCritico() {
  _navegarAHoja(HOJAS.MSTR_PRODUCTOS, 'Stock de Productos');
}

function registrarIngreso() {
  _navegarAHoja(HOJAS.DB_INGRESOS, 'Registro de Ingresos');
}

function registrarEgreso() {
  _navegarAHoja(HOJAS.DB_EGRESOS, 'Registro de Egresos');
}

function registrarPagoCxC() {
  _navegarAHoja(HOJAS.DB_CXC, 'Registro de Pago CxC');
}

function registrarPagoCxP() {
  _navegarAHoja(HOJAS.DB_CXP, 'Registro de Pago CxP');
}

function procesarAjusteStock() {
  _navegarAHoja(HOJAS.DB_MOVIMIENTOS, 'Ajuste de Stock');
}

function procesarTraslado() {
  _navegarAHoja(HOJAS.DB_MOVIMIENTOS, 'Traslado entre Bodegas');
}

function procesarAnulacionVenta() {
  _navegarAHoja(HOJAS.DB_VENTAS, 'Anulación de Venta');
}

function procesarAnulacionOC() {
  _navegarAHoja(HOJAS.DB_COMPRAS, 'Anulación de OC');
}

function limpiarLogAntiguoDesdeMenu() {
  var resultado = limpiarLogAntiguo(6);
  SpreadsheetApp.getUi().alert(
    resultado
      ? '✅ Log limpiado correctamente.\nSe conservaron los últimos 6 meses.'
      : '❌ Error al limpiar el log. Revisa el registro de ejecución.'
  );
}

// ============================================================
// FUNCIÓN PRIVADA: _getUsuarioActivo
// ============================================================

/**
 * Retorna el email del usuario activo.
 * @returns {string} Email o string vacío si no está disponible
 */
function _getUsuarioActivo() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    Logger.log('ERROR _getUsuarioActivo: ' + e.message);
    return '';
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarAcceso
// ============================================================

/**
 * Verifica si un rol tiene permiso para ejecutar una acción.
 * Usado por M2 y M3 para validar permisos antes de operar.
 *
 * @param {string} rol           - Rol a verificar
 * @param {Array}  rolesPermitidos - Lista de roles que pueden ejecutar
 * @returns {boolean} true si tiene acceso
 */
function _verificarAcceso(rol, rolesPermitidos) {
  try {
    return rolesPermitidos.indexOf(rol) !== -1;
  } catch (e) {
    Logger.log('ERROR _verificarAcceso: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _navegarAHoja
// ============================================================

/**
 * Navega a una hoja específica y posiciona el cursor en la
 * primera fila vacía disponible para ingreso de datos.
 *
 * @param {string} nombreHoja - Nombre de la hoja destino
 * @param {string} titulo     - Título descriptivo para el log
 */
function _navegarAHoja(nombreHoja, titulo) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(nombreHoja);

    if (!hoja) {
      SpreadsheetApp.getUi().alert(
        '❌ Hoja no encontrada\n\n' +
        'La hoja "' + nombreHoja + '" no existe en el sistema.'
      );
      return;
    }

    // Mostrar la hoja si está oculta
    if (hoja.isSheetHidden()) hoja.showSheet();

    // Activar la hoja
    ss.setActiveSheet(hoja);

    // Posicionar en la primera fila vacía (después de encabezados)
    var ultimaFila = hoja.getLastRow();
    var filaDestino = Math.max(ultimaFila + 1, 2);
    hoja.setActiveRange(hoja.getRange(filaDestino, 1));

    Logger.log('_navegarAHoja: Navegado a ' + titulo + ' (' + nombreHoja + ') fila ' + filaDestino);

  } catch (e) {
    Logger.log('ERROR _navegarAHoja ' + nombreHoja + ': ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _construirMenuMinimo
// ============================================================

/**
 * Construye un menú mínimo de emergencia cuando onOpen falla.
 * Permite al usuario acceder a configuración básica.
 */
function _construirMenuMinimo() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏢 ERP')
    .addItem('ℹ️ Información del sistema', 'mostrarInfoSistema')
    .addItem('Inicializar sistema', 'setupSistema')
    .addToUi();
  Logger.log('_construirMenuMinimo: Menú mínimo construido');
}

// ============================================================
// FUNCIÓN PRIVADA: _verificarActualizacionesSemanal
// ============================================================

/**
 * Verifica actualizaciones silenciosamente una vez por semana.
 * Se llama desde onOpen() sin interrumpir al usuario.
 */
function _verificarActualizacionesSemanal() {
  try {
    var props        = PropertiesService.getScriptProperties();
    var ultimaVerif  = props.getProperty('ULTIMA_VERIFICACION_SEMANAL');
    var ahora        = new Date();
    var diasDesde    = ultimaVerif
      ? (ahora - new Date(ultimaVerif)) / (1000 * 60 * 60 * 24)
      : 8;

    if (diasDesde >= 7) {
      props.setProperty('ULTIMA_VERIFICACION_SEMANAL', ahora.toISOString());
      checkForUpdates(true);
    }

  } catch (e) {
    Logger.log('ERROR _verificarActualizacionesSemanal: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _versionEsMenor
// ============================================================

/**
 * Compara dos versiones semánticas.
 * Retorna true si v1 es menor que v2.
 *
 * @param {string} v1 - Versión actual. Ej: '1.0.0'
 * @param {string} v2 - Versión a comparar. Ej: '1.0.1'
 * @returns {boolean} true si v1 < v2
 */
function _versionEsMenor(v1, v2) {
  try {
    var p1 = String(v1).split('.').map(Number);
    var p2 = String(v2).split('.').map(Number);
    for (var i = 0; i < 3; i++) {
      if ((p1[i] || 0) < (p2[i] || 0)) return true;
      if ((p1[i] || 0) > (p2[i] || 0)) return false;
    }
    return false;
  } catch (e) {
    Logger.log('ERROR _versionEsMenor: ' + e.message);
    return false;
  }
}

// ============================================================
// FIN DE M1_Sesion.gs
// ============================================================
