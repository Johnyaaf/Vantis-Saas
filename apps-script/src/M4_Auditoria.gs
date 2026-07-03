/**
 * ============================================================
 * M4_Auditoria.gs — ERP PYME Inteligente v1.0
 * ============================================================
 * Responsabilidad: Registra todas las acciones del sistema
 * en LOG_AUDITORIA usando un buffer en PropertiesService
 * para minimizar escrituras directas a la hoja.
 *
 * Autor: ERP PYME
 * Versión: 1.0.0
 * Compatibilidad: Apps Script V8
 *
 * DEPENDENCIAS:
 *   Módulos:      M0_Setup (HOJAS, COL, SISTEMA)
 *   Hojas:        LOG_AUDITORIA
 *   APIs Google:  PropertiesService, Session, SpreadsheetApp
 *   Triggers:     flushLogBuffer (cada 5 minutos — instalado por M0)
 * ============================================================
 */

// ============================================================
// FUNCIÓN: registrarLog
// Agrega una entrada al buffer de log en PropertiesService
// ============================================================

/**
 * Registra una acción del sistema en el buffer de auditoría.
 * No escribe directamente en la hoja — usa buffer para eficiencia.
 * El buffer se vacía automáticamente cada 5 minutos por trigger,
 * o inmediatamente cuando alcanza SISTEMA.LOG_BUFFER_MAX entradas.
 *
 * @param {string} tabla       - Nombre de la hoja afectada. Ej: 'DB_VENTAS'
 * @param {string} accion      - Acción ejecutada. Ej: 'CREAR', 'MODIFICAR', 'ELIMINAR'
 * @param {string} idRegistro  - ID del registro afectado. Ej: 'VTA-20250615-001'
 * @param {string} campo       - Campo modificado. Ej: 'stock_actual'
 * @param {string} valorNuevo  - Nuevo valor del campo
 */
function registrarLog(tabla, accion, idRegistro, campo, valorNuevo) {
  try {
    var props      = PropertiesService.getScriptProperties();
    var bufferJSON = props.getProperty(SISTEMA.LOG_BUFFER_KEY);

    // Inicializar buffer si no existe
    var buffer = [];
    if (bufferJSON) {
      try {
        buffer = JSON.parse(bufferJSON);
      } catch (parseErr) {
        Logger.log('ADVERTENCIA registrarLog: Buffer corrupto, reiniciando. ' + parseErr.message);
        buffer = [];
      }
    }

    // Detectar usuario según contexto de ejecución
    var usuario = _detectarUsuario();

    // Obtener o generar session ID
    var sessionId = props.getProperty(SISTEMA.SESSION_ID_KEY) || generarID('SES');

    // Obtener id_empresa desde PropertiesService o usar default
    var idEmpresa = SISTEMA.EMPRESA_DEFAULT;
    try {
      var configEmpresa = props.getProperty('ID_EMPRESA');
      if (configEmpresa) idEmpresa = configEmpresa;
    } catch (e) {
      // Usar default si falla
    }

    // Construir entrada de log
    var entrada = [
      generarID('LOG'),           // A — id_log
      new Date().toISOString(),   // B — timestamp
      usuario,                    // C — usuario
      tabla,                      // D — tabla
      accion,                     // E — accion
      String(idRegistro || ''),   // F — id_registro
      String(campo || ''),        // G — campo
      String(valorNuevo || ''),   // H — valor_nuevo
      idEmpresa,                  // I — id_empresa
      sessionId                   // J — session_id
    ];

    // Agregar al buffer
    buffer.push(entrada);

    // Guardar buffer actualizado
    props.setProperty(SISTEMA.LOG_BUFFER_KEY, JSON.stringify(buffer));

    // Flush automático si se alcanza el límite máximo
    if (buffer.length >= SISTEMA.LOG_BUFFER_MAX) {
      Logger.log('registrarLog: Buffer lleno (' + buffer.length + '), ejecutando flush automático');
      flushLogBuffer();
    }

  } catch (e) {
    // El log nunca debe detener la operación principal
    // Solo registra el error internamente
    Logger.log('ERROR registrarLog: ' + e.message);
  }
}

// ============================================================
// FUNCIÓN: flushLogBuffer
// Vacía el buffer y escribe en LOG_AUDITORIA con setValues()
// Llamada por trigger cada 5 minutos y por flush automático
// ============================================================

/**
 * Escribe todas las entradas pendientes del buffer en LOG_AUDITORIA.
 * Usa setValues() para escribir todas las filas de una sola vez.
 * Es el único punto donde el código escribe en LOG_AUDITORIA.
 *
 * Esta función es llamada por:
 * - El trigger time-based cada 5 minutos
 * - registrarLog() cuando el buffer alcanza el máximo
 * - El menú ERP → Configuración → Vaciar buffer de log
 */
function flushLogBuffer() {
  try {
    var props      = PropertiesService.getScriptProperties();
    var bufferJSON = props.getProperty(SISTEMA.LOG_BUFFER_KEY);

    // Si no hay buffer o está vacío, no hacer nada
    if (!bufferJSON) {
      Logger.log('flushLogBuffer: Buffer vacío, nada que escribir');
      return;
    }

    var buffer = [];
    try {
      buffer = JSON.parse(bufferJSON);
    } catch (parseErr) {
      Logger.log('ERROR flushLogBuffer: Buffer corrupto — ' + parseErr.message);
      // Limpiar buffer corrupto
      props.setProperty(SISTEMA.LOG_BUFFER_KEY, '[]');
      return;
    }

    if (!buffer || buffer.length === 0) {
      Logger.log('flushLogBuffer: Buffer vacío, nada que escribir');
      return;
    }

    Logger.log('flushLogBuffer: Escribiendo ' + buffer.length + ' entradas en LOG_AUDITORIA');

    // Obtener la hoja de auditoría
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var hojaLog = ss.getSheetByName(HOJAS.LOG_AUDITORIA);

    if (!hojaLog) {
      Logger.log('ERROR flushLogBuffer: Hoja LOG_AUDITORIA no encontrada');
      return;
    }

    // Encontrar la primera fila vacía después de los datos existentes
    var ultimaFila = hojaLog.getLastRow();
    var filaInicio = ultimaFila + 1;

    // Verificar que no excedemos el límite de la hoja
    var filasDisponibles = hojaLog.getMaxRows() - ultimaFila;
    if (filasDisponibles < buffer.length) {
      Logger.log('ADVERTENCIA flushLogBuffer: Espacio insuficiente en LOG_AUDITORIA. ' +
        'Disponibles: ' + filasDisponibles + ' | Necesarias: ' + buffer.length);
      // Escribir solo las que caben
      buffer = buffer.slice(0, filasDisponibles);
    }

    if (buffer.length === 0) {
      Logger.log('ADVERTENCIA flushLogBuffer: LOG_AUDITORIA lleno, no se pueden escribir más entradas');
      props.setProperty(SISTEMA.LOG_BUFFER_KEY, '[]');
      return;
    }

    // Escribir todas las entradas de una sola vez con setValues()
    hojaLog.getRange(filaInicio, 1, buffer.length, 10).setValues(buffer);

    // Limpiar el buffer después de escribir exitosamente
    props.setProperty(SISTEMA.LOG_BUFFER_KEY, '[]');

    Logger.log('flushLogBuffer: ' + buffer.length + ' entradas escritas correctamente desde fila ' + filaInicio);

  } catch (e) {
    Logger.log('ERROR flushLogBuffer: ' + e.message);
    // No limpiar el buffer si hubo error — intentar en el próximo ciclo
  }
}

// ============================================================
// FUNCIÓN: limpiarLogAntiguo
// Archiva entradas de log más antiguas que N meses
// Llamada desde el menú ERP o por el batch nocturno
// ============================================================

/**
 * Elimina las entradas de LOG_AUDITORIA más antiguas que el
 * número de meses indicado. Protege las entradas recientes.
 *
 * @param {number} mesesRetener - Número de meses a conservar (default: 6)
 */
function limpiarLogAntiguo(mesesRetener) {
  try {
    var meses = mesesRetener || 6;
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var hoja  = ss.getSheetByName(HOJAS.LOG_AUDITORIA);

    if (!hoja) {
      Logger.log('ERROR limpiarLogAntiguo: Hoja LOG_AUDITORIA no encontrada');
      return false;
    }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila <= 1) {
      Logger.log('limpiarLogAntiguo: No hay entradas que limpiar');
      return true;
    }

    // Calcular fecha límite
    var fechaLimite = new Date();
    fechaLimite.setMonth(fechaLimite.getMonth() - meses);

    // Leer columna de timestamp (columna B = índice 2)
    var timestamps = hoja.getRange(2, 2, ultimaFila - 1, 1).getValues();

    // Encontrar la primera fila que está dentro del período a conservar
    var primeraFilaAConservar = -1;
    for (var i = 0; i < timestamps.length; i++) {
      var fechaEntrada = new Date(timestamps[i][0]);
      if (fechaEntrada >= fechaLimite) {
        primeraFilaAConservar = i + 2; // +2 porque índice base 0 + fila encabezado
        break;
      }
    }

    if (primeraFilaAConservar === -1) {
      Logger.log('limpiarLogAntiguo: Todas las entradas son antiguas — eliminando todo');
      hoja.getRange(2, 1, ultimaFila - 1, 10).clearContent();
      registrarLog('LOG_AUDITORIA', 'LIMPIAR', 'TODAS', 'timestamp', 'Limpieza completa ejecutada');
      return true;
    }

    if (primeraFilaAConservar === 2) {
      Logger.log('limpiarLogAntiguo: No hay entradas antiguas que limpiar');
      return true;
    }

    // Eliminar filas antiguas (desde fila 2 hasta primeraFilaAConservar - 1)
    var filasAEliminar = primeraFilaAConservar - 2;
    hoja.deleteRows(2, filasAEliminar);

    Logger.log('limpiarLogAntiguo: ' + filasAEliminar + ' entradas antiguas eliminadas');
    registrarLog('LOG_AUDITORIA', 'LIMPIAR', filasAEliminar + '_FILAS', 'timestamp',
      'Limpieza de entradas anteriores a ' + fechaLimite.toISOString());

    return true;

  } catch (e) {
    Logger.log('ERROR limpiarLogAntiguo: ' + e.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: getEstadoBuffer
// Retorna información sobre el estado actual del buffer
// Útil para diagnóstico desde el menú ERP
// ============================================================

/**
 * Retorna el estado actual del buffer de log.
 * Usado para diagnóstico y monitoreo del sistema.
 *
 * @returns {Object} Estado del buffer con conteo y última entrada
 */
function getEstadoBuffer() {
  try {
    var props      = PropertiesService.getScriptProperties();
    var bufferJSON = props.getProperty(SISTEMA.LOG_BUFFER_KEY);
    var buffer     = bufferJSON ? JSON.parse(bufferJSON) : [];

    return {
      entradas:       buffer.length,
      maximo:         SISTEMA.LOG_BUFFER_MAX,
      porcentaje:     Math.round((buffer.length / SISTEMA.LOG_BUFFER_MAX) * 100),
      ultimaEntrada:  buffer.length > 0 ? buffer[buffer.length - 1][1] : 'ninguna',
      bufferLleno:    buffer.length >= SISTEMA.LOG_BUFFER_MAX
    };

  } catch (e) {
    Logger.log('ERROR getEstadoBuffer: ' + e.message);
    return { entradas: 0, error: e.message };
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _detectarUsuario
// Detecta el usuario según el contexto de ejecución
// ============================================================

/**
 * Detecta quién está ejecutando el código.
 * En triggers automáticos Session.getActiveUser() retorna vacío,
 * por lo que se usa 'sistema-automatico' como identificador.
 *
 * @returns {string} Email del usuario o 'sistema-automatico'
 */
function _detectarUsuario() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email !== '') {
      return email;
    }
    return 'sistema-automatico';
  } catch (e) {
    return 'sistema-automatico';
  }
}

// ============================================================
// FUNCIÓN PRIVADA: _generarSessionId
// Genera o recupera el ID de sesión actual
// ============================================================

/**
 * Retorna el ID de sesión actual desde PropertiesService.
 * Si no existe lo genera y lo guarda.
 *
 * @returns {string} ID de sesión en formato SES-XXXXXXXX
 */
function _generarSessionId() {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sessionId = props.getProperty(SISTEMA.SESSION_ID_KEY);

    if (!sessionId) {
      sessionId = generarID('SES');
      props.setProperty(SISTEMA.SESSION_ID_KEY, sessionId);
    }

    return sessionId;
  } catch (e) {
    Logger.log('ERROR _generarSessionId: ' + e.message);
    return 'SES-ERROR';
  }
}

// ============================================================
// FIN DE M4_Auditoria.gs
// ============================================================
