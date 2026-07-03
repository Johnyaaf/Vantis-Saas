# Analisis de migracion VANTIS ERP a SaaS

## Objetivo corregido

El producto final no debe depender de Google Sheets, Google Apps Script, `SpreadsheetApp`, `google.script.run`, `PropertiesService`, `ScriptApp`, `HtmlService` ni triggers de Google.

Los archivos Apps Script quedan solo como referencia funcional para entender reglas de negocio, pantallas, datos y procesos. La implementacion real debe vivir en:

- `backend`: FastAPI, PostgreSQL, SQLAlchemy, Alembic, jobs y autenticacion.
- `frontend`: app web SaaS, idealmente React/Vite/TypeScript.
- `apps-script`: fuente historica, no ejecutable en produccion.

## Estado del frontend

La carpeta `frontend` actual fue eliminada. Tambien se elimino el `package.json` raiz que habia sido creado solo para levantar ese frontend temporal.

## Modulos analizados

### M0_Setup.gs

Responsabilidad original:

- Define constantes globales de hojas.
- Define indices de columnas.
- Define roles, estados y configuracion del sistema.
- Genera IDs.
- Lee configuracion del sistema.
- Valida periodo abierto.
- Instala triggers.
- Configura protecciones de hojas.

Funciones principales:

- `generarID`
- `getConfig`
- `getPeriodoActual`
- `validarPeriodoAbierto`
- `setupSistema`
- `configurarProtecciones`

Migracion SaaS:

- Las constantes de hojas pasan a tablas/modelos SQL.
- Los indices de columnas desaparecen; se reemplazan por columnas nombradas.
- `generarID` debe reemplazarse por UUIDs o secuencias por documento.
- `getConfig` pasa a tabla `tenant_config`.
- `getPeriodoActual` pasa a tabla `periodos_contables`.
- `validarPeriodoAbierto` pasa a una dependencia/servicio backend antes de crear ventas, compras, pagos o movimientos.
- Los triggers de Apps Script pasan a jobs backend programados.

### M1_Sesion.gs

Responsabilidad original:

- Controla apertura del sistema.
- Construye menu por rol.
- Verifica usuario activo.
- Aplica visibilidad de hojas.
- Expone accesos rapidos a acciones operativas.

Funciones relevantes:

- `onOpen`
- `construirMenu`
- `getRolUsuario`
- `aplicarVisibilidadHojas`
- `_verificarAcceso`
- `_getUsuarioActivo`

Migracion SaaS:

- El menu pasa al frontend.
- Los permisos pasan a JWT + roles + permisos backend.
- La visibilidad de hojas desaparece; se reemplaza por rutas protegidas y permisos por modulo.
- `Session.getActiveUser()` se reemplaza por usuario autenticado en base de datos.

### M2_Transacciones.gs

Responsabilidad original:

- Motor transaccional.
- Procesa ventas, compras, recepcion de OC, movimientos de stock, CxC, CxP, ingresos y egresos.
- Valida cliente, proveedor, producto, stock y periodo.
- Captura costo al momento de la venta.
- Crea cuentas por cobrar/pagar.

Funciones clave:

- `procesarVenta`
- `procesarCompra`
- `procesarRecepcionOC`
- `actualizarStock`
- `crearCxCDesdeVenta`
- `crearCxPDesdeCompra`
- `registrarIngreso`
- `registrarEgreso`
- `procesarAnulacion`
- `_validarStock`
- `_capturarCostoMomento`

Migracion SaaS:

- Este es el nucleo del backend.
- Debe migrarse antes que dashboard avanzado.
- Requiere transacciones SQL reales con commit/rollback.
- Las escrituras en varias hojas deben convertirse en una operacion atomica:
  - venta
  - lineas de venta
  - movimiento de stock
  - cuenta por cobrar
  - log de auditoria
- Las validaciones deben vivir en servicios backend, no en frontend.

### M3_Maestros.gs

Responsabilidad original:

- CRUD de maestros.
- Clientes, proveedores, productos, bodegas, plan de cuentas y tablas de referencia.
- Normaliza RUT.
- Valida duplicados.
- Soft delete.

Funciones clave:

- `crearCliente`
- `crearProveedor`
- `darDeBajaMaestro`
- `normalizarRUT`
- `verificarDuplicadoRUT`
- `verificarDuplicadoSKU`
- `validarSoftDelete`

Migracion SaaS:

- Debe mapearse a modulos backend REST:
  - `/api/clientes`
  - `/api/proveedores`
  - `/api/productos`
  - `/api/bodegas`
  - `/api/cuentas-contables`
  - `/api/tablas-referencia`
- RUT debe validarse en backend y tambien en frontend como ayuda visual.
- Soft delete debe ser columna `activo` o `deleted_at`.
- Duplicados deben reforzarse con constraints SQL por tenant.

### M4_Auditoria.gs

Responsabilidad original:

- Registra acciones en `LOG_AUDITORIA`.
- Usa buffer en `PropertiesService`.
- Hace flush periodico hacia hoja.

Funciones clave:

- `registrarLog`
- `flushLogBuffer`
- `limpiarLogAntiguo`
- `getEstadoBuffer`

Migracion SaaS:

- Reemplazar por tabla `audit_logs`.
- No usar buffer de Google.
- Se puede registrar directo en base de datos dentro de la misma transaccion.
- Para alto volumen, usar cola o tarea async despues.
- Debe guardar: tenant, usuario, modulo, accion, entidad, id_entidad, cambios JSON, timestamp, ip/opcional.

### M5_Batch.gs

Responsabilidad original:

- Procesos nocturnos.
- ABC productos/clientes.
- Snapshot de stock.
- Score de proveedores.
- Alertas.
- Emails.
- Creacion de periodo siguiente.

Funciones clave:

- `batchNocturno`
- `batchABCProductos`
- `batchABCClientes`
- `generarSnapshotStock`
- `batchScoreCumplimientoProveedores`
- `detectarAlertas`
- `enviarAlertasEmail`
- `crearPeriodoSiguiente`

Migracion SaaS:

- Pasa a jobs backend:
  - APScheduler, Celery, RQ, Dramatiq o cron externo.
- Resultados van a tablas:
  - `batch_runs`
  - `stock_snapshots`
  - `alertas`
  - `proveedor_scores`
  - `clasificacion_abc`
- Emails deben salir por proveedor real: SMTP, SendGrid, Resend, SES, etc.

### M6_CalcMotores.gs

Responsabilidad original:

- Construye hojas calculadas:
  - `CALC_SUPPLY`
  - `CALC_FINANCIERO`
  - `CALC_COMERCIAL`
  - `CALC_GERENCIAL`
- Valida estructuras y semaforos.
- Crea plan de cuentas base.

Funciones clave:

- `crearCALC_SUPPLY`
- `crearCALC_FINANCIERO`
- `crearCALC_COMERCIAL`
- `crearCALC_GERENCIAL`
- `validarCALC_SUPPLY`
- `validarCALC_FINANCIERO`
- `validarCALC_GERENCIAL`

Migracion SaaS:

- No crear hojas calculadas.
- Convertir a servicios de reporting.
- Algunas metricas pueden calcularse por SQL views o materialized views.
- Los indicadores gerenciales pueden exponerse como endpoints:
  - `/api/dashboard`
  - `/api/reportes/financiero`
  - `/api/reportes/comercial`
  - `/api/reportes/inventario`
  - `/api/reportes/supply`

### M7_Utils.gs

Responsabilidad original:

- Herramientas de diagnostico, auditoria, correccion y mantenimiento.
- Crea plantillas.
- Limpia filas fantasma.
- Corrige formulas y validaciones.
- Contiene muchas funciones temporales/de soporte.

Funciones detectadas: 142.

Conclusion:

- No debe migrarse completo.
- Hay que clasificarlo en tres grupos:
  - utilidades descartables de Google Sheets
  - scripts de migracion puntual
  - funcionalidades reales que deben pasar a backend

Ejemplos descartables:

- `limpiarFilasFantasmaSeguro`
- funciones de validaciones de hojas
- funciones de correccion de formulas
- diagnosticos de rangos/celdas

Ejemplos potencialmente utiles:

- plantillas CSV
- auditorias de integridad
- diagnosticos de datos
- setup de estructuras base

### M8_Dashboard.gs

Responsabilidad original:

- Servidor de datos para la app HTML.
- Abre la app.
- Entrega datos al dashboard.
- Entrega datos a finanzas, comercial, inventario, cobranza, proveedores y alertas.
- Ejecuta acciones desde la UI: cobros, pagos, anticipos, cierres, configuracion y respaldos.

Funciones clave:

- `getDashboardData`
- `getFinanzasData`
- `getComercialData`
- `getInventarioData`
- `getCobranzaData`
- `getProveedoresData`
- `getAlertasData`
- `listarClientes`
- `listarProveedores`
- `registrarCobroCxC`
- `registrarPagoCxP`
- `cerrarPeriodo`
- `reabrirPeriodo`
- `getEERRData`
- `getFlujoCajaMultiMes`
- `getRentabilidadClientes`
- `guardarConfigEmpresa`
- `guardarConfigTributario`
- `guardarConfigComercial`
- `guardarConfigInventario`
- `exportarRespaldoCompleto`

Migracion SaaS:

- Este modulo define la mayoria de endpoints de frontend.
- No debe migrarse como archivo unico.
- Debe dividirse en routers:
  - `dashboard`
  - `finanzas`
  - `comercial`
  - `inventario`
  - `cobranza`
  - `proveedores`
  - `reportes`
  - `configuracion`
  - `respaldos`

### vantis_app.html

Responsabilidad original:

- Interfaz de usuario monolitica.
- Contiene HTML, CSS y JavaScript juntos.
- Llama a Apps Script con `google.script.run`.
- Maneja navegacion, modales, tablas, formularios, reportes y configuracion.

Funciones detectadas: 201.
Llamadas `google.script.run`: 87.

Conclusion:

- No debe usarse como frontend final.
- Sirve como referencia visual y funcional.
- Debe reimplementarse como app SaaS.

Migracion frontend recomendada:

- React + TypeScript + Vite.
- React Router.
- TanStack Query.
- React Hook Form + Zod.
- Tablas con TanStack Table.
- Componentes de layout, sidebar, topbar, modales, formularios y reportes.

## Modelo de datos inferido

Tablas principales que deben existir en PostgreSQL:

- `tenants`
- `usuarios`
- `usuario_tenant`
- `roles`
- `permisos`
- `periodos_contables`
- `clientes`
- `proveedores`
- `productos`
- `bodegas`
- `plan_cuentas`
- `tablas_referencia`
- `tipos_cambio`
- `ventas`
- `venta_lineas`
- `compras`
- `compra_lineas`
- `movimientos_stock`
- `cuentas_por_cobrar`
- `cuentas_por_pagar`
- `ingresos`
- `egresos`
- `caja_movimientos`
- `presupuestos`
- `metas`
- `historial_precios`
- `stock_snapshots`
- `alertas`
- `batch_runs`
- `audit_logs`
- `medios_pago`
- `anticipos`
- `cierres_periodo`
- `config_empresa`
- `config_tributaria`
- `config_comercial`
- `config_inventario`
- `config_alertas`

## Arquitectura backend recomendada

Estructura sugerida:

```txt
backend/app/
  core/
    config.py
    database.py
    security.py
    dependencies.py
  models/
    tenant.py
    auth.py
    maestros.py
    transacciones.py
    finanzas.py
    inventario.py
    auditoria.py
    configuracion.py
  modules/
    auth/
    tenants/
    clientes/
    proveedores/
    productos/
    bodegas/
    ventas/
    compras/
    inventario/
    cobranza/
    pagos/
    finanzas/
    dashboard/
    reportes/
    configuracion/
    auditoria/
    jobs/
```

## Arquitectura frontend recomendada

Estructura sugerida:

```txt
frontend/
  src/
    app/
      router.tsx
      providers.tsx
    api/
      client.ts
      auth.ts
      clientes.ts
      dashboard.ts
      ventas.ts
      compras.ts
      inventario.ts
      finanzas.ts
    components/
      layout/
      ui/
      forms/
      tables/
      charts/
      modals/
    modules/
      auth/
      dashboard/
      clientes/
      proveedores/
      productos/
      ventas/
      compras/
      inventario/
      cobranza/
      finanzas/
      reportes/
      configuracion/
```

## Prioridad de migracion

1. Base SaaS:
   - frontend vacio nuevo
   - login
   - registro tenant
   - layout privado
   - JWT

2. Maestros:
   - clientes
   - proveedores
   - productos
   - bodegas
   - plan de cuentas

3. Transacciones:
   - ventas
   - compras
   - stock
   - CxC
   - CxP
   - ingresos
   - egresos

4. Dashboard:
   - KPIs base
   - alertas
   - cobranza
   - inventario
   - comercial

5. Finanzas/reportes:
   - EERR
   - flujo de caja
   - rentabilidad
   - consolidado financiero

6. Jobs:
   - batch nocturno
   - ABC
   - snapshots
   - alertas por email
   - cierre/apertura de periodos

## Riesgos principales

- El sistema original depende demasiado de indices de columnas. En SaaS eso debe convertirse a modelos tipados.
- Muchas reglas estan mezcladas con UI y hoja de calculo.
- `M7_Utils.gs` contiene mucho codigo operacional temporal; migrarlo completo seria un error.
- `vantis_app.html` es demasiado monolitico para mantenerlo como base final.
- Los procesos transaccionales deben ser atomicos en PostgreSQL para evitar inconsistencias.
- Hay que decidir estrategia multi-tenant: schema por tenant o tenant_id por tabla.

## Recomendacion tecnica

No intentar convertir Apps Script linea por linea.

La mejor ruta es reconstruir el producto con arquitectura SaaS, usando estos archivos como especificacion funcional. El orden correcto es:

1. Diseñar base de datos.
2. Completar backend por modulos.
3. Crear frontend SaaS desde cero.
4. Migrar primero maestros.
5. Migrar despues transacciones.
6. Finalmente migrar dashboard, calculos y batch.
