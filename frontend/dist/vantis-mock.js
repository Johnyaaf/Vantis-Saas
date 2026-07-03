(function () {
  document.addEventListener("DOMContentLoaded", function () {
    const icons = {
      dashboard: "🏠",
      finanzas: "💰",
      comercial: "📈",
      transacciones: "🧾",
      inventario: "🏭",
      cobranza: "🤝",
      proveedores: "🏢",
      alertas: "🔔",
      "carga-masiva": "📥",
      cierres: "🔒",
      configuracion: "⚙️",
      reportes: "📊",
    };
    document.querySelectorAll(".sitem").forEach(function (item) {
      const icon = item.querySelector(".sicon");
      if (icon && icons[item.dataset.mod]) icon.textContent = icons[item.dataset.mod];
    });
  });

  const ctx = {
    nombreEmpresa: "Comercializadora Zentra Spa",
    periodo: "2026-06",
    fechaCalculo: "01/07/2026",
    usuario: "Administrador",
  };

  const productos = [
    { sku: "SKU-011", nombre: "Toner Compatible HP LaserJet", stock: 18, stockMin: 4, costo: 18900, valor: 340200, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-016", nombre: "Cable Electrico 2.5mm x100m", stock: 12, stockMin: 3, costo: 28000, valor: 336000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-017", nombre: "Pintura Latex Blanca 4L", stock: 35, stockMin: 8, costo: 8900, valor: 311500, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-018", nombre: "Cemento Gris 25kg", stock: 50, stockMin: 15, costo: 5200, valor: 260000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-005", nombre: "Cafe Molido 250g", stock: 70, stockMin: 20, costo: 2900, valor: 203000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-003", nombre: "Detergente Liquido 3L", stock: 60, stockMin: 15, costo: 3200, valor: 192000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-014", nombre: "Mascarillas KN95 x10", stock: 40, stockMin: 10, costo: 3200, valor: 128000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-004", nombre: "Papel Bond A4 500 hojas", stock: 45, stockMin: 10, costo: 2800, valor: 126000, diasSinVenta: 999, estado: "OK" },
    { sku: "SKU-019", nombre: "Servicio Instalacion Electrica", stock: 0, stockMin: 0, costo: 0, valor: 0, diasSinVenta: 999, estado: "SIN_STOCK" },
    { sku: "SKU-020", nombre: "Servicio Mantencion Mensual", stock: 0, stockMin: 0, costo: 0, valor: 0, diasSinVenta: 999, estado: "SIN_STOCK" },
  ];

  const cxcDocumentos = [
    { cliente: "Industrias Metalicas Atacama", documento: "FAC-20260607", emision: "28-06-2026", vence: "28-07-2026", dias: 26, saldo: 4200000, estado: "OK" },
    { cliente: "Centro Automotriz Velocidad", documento: "FAC-20260617", emision: "01-06-2026", vence: "01-07-2026", dias: -1, saldo: 3400000, estado: "VENCIDO" },
    { cliente: "Constructora Pacifico Ltda.", documento: "FAC-20260602", emision: "11-06-2026", vence: "11-07-2026", dias: 9, saldo: 3200000, estado: "OK" },
    { cliente: "Vina Santa Elena SpA", documento: "FAC-20260610", emision: "02-06-2026", vence: "04-07-2026", dias: 2, saldo: 2900000, estado: "RIESGO" },
    { cliente: "Farmacia Bienestar Ltda.", documento: "FAC-20260604", emision: "21-06-2026", vence: "21-07-2026", dias: 19, saldo: 2450000, estado: "OK" },
  ];

  const cxpDocumentos = [
    { proveedor: "Lubricantes y Combustibles SpA", documento: "FAC-20260616", emision: "01-06-2026", vence: "01-07-2026", dias: -1, saldo: 4200000, estado: "MORA" },
    { proveedor: "Servicios Logisticos Expres SpA", documento: "FAC-20260610", emision: "04-06-2026", vence: "04-07-2026", dias: 2, saldo: 3800000, estado: "URGENTE" },
    { proveedor: "Comercial Andina SpA", documento: "FAC-20260603", emision: "21-06-2026", vence: "21-07-2026", dias: 19, saldo: 3100000, estado: "PLANIFICADO" },
    { proveedor: "Materiales de Construccion Andes", documento: "FAC-20260614", emision: "07-05-2026", vence: "16-06-2026", dias: -16, saldo: 2650000, estado: "MORA" },
    { proveedor: "Importadora Pacific Trade SpA", documento: "FAC-20260601", emision: "11-06-2026", vence: "11-07-2026", dias: 9, saldo: 2400000, estado: "PLANIFICADO" },
  ];

  const responses = {
    getDatosContexto: () => ctx,
    getDashboardData: () => ({
      ...ctx,
      scoreTotal: 52,
      estadoGeneral: "CRITICO",
      flujoNeto: 0,
      diasCaja: 999,
      ratioLiquidez: 1.11,
      ventasMes: 0,
      cxcVencida: 7190000,
      margenPct: 0,
      semExecFinanzas: 2,
      semExecComercial: 3,
      semExecOperaciones: 2,
      semLiquidez: 2,
      semResultado: 1,
      semVentas: 2,
      ventas: [],
      ingresosEgresos: [],
      agingCxc: { alDia: 24890000, vencido: 7190000 },
      skusRiesgo: [],
    }),
    getFinanzasData: () => ({
      ...ctx,
      cajaReal: 0,
      ingresosNeto: 0,
      egresosNeto: 0,
      margenPct: 0,
      liquidez30d: 2520000,
      diasCaja: 999,
      ratioLiquidez: 1.11,
      flujoNeto: 2520000,
      cxcTotal: 32080000,
      cxc030: 24890000,
      cxc3160: 0,
      cxcMas60: 0,
      cxcVencida: 7190000,
      metaVentas: 0,
      metaPct: 0,
      evolucion: [],
      detalleIngresos: [],
      top5Egresos: [],
    }),
    getComercialData: () => ({
      ...ctx,
      ventasNeto: 0,
      ticketPromedio: 0,
      margenBrutoPct: 0,
      documentosMes: 0,
      ventasHoy: 0,
      docsHoy: 0,
      metaVentas: 0,
      metaPct: 0,
      tiposDoc: [],
      evolucion: [],
      topProductos: [],
      topClientes: [],
    }),
    getInventarioData: () => ({
      ...ctx,
      productosActivos: 18,
      valorCosto: 2796900,
      criticos: 0,
      sinStock: 0,
      diasRotacion: 999,
      productos,
      evolucion: [],
    }),
    getCobranzaData: () => ({
      ...ctx,
      cxcTotal: 32080000,
      cxcVencida: 10590000,
      docsPendientes: 20,
      docsVencidos: 7,
      dso: 4,
      dsoSemaforo: "verde",
      plazoPromedio: 35,
      flujoEsperado30d: 21490000,
      cxc030: 21490000,
      cxc3160: 0,
      cxcMas60: 0,
      concentracionMax: 13,
      topDeudores: cxcDocumentos.map((d) => ({ cliente: d.cliente, saldo: d.saldo, pct: 13 })),
      documentos: cxcDocumentos,
    }),
    getProveedoresData: () => ({
      ...ctx,
      cxpTotal: 32870000,
      cxpVencida: 10320000,
      cxpUrgente: 16250000,
      cxp630: 16620000,
      cxpMas30: 0,
      docsPendientes: 20,
      docsVencidos: 6,
      dpo: 6,
      dpoSaludable: true,
      plazoPromedio: 30,
      presion5dias: 16250000,
      cajaActual: 0,
      faltanteCaja: 16250000,
      concentracionMax: 13,
      topProveedores: cxpDocumentos.map((d) => ({ proveedor: d.proveedor, saldo: d.saldo, pct: 13 })),
      documentos: cxpDocumentos,
    }),
    getAlertasData: () => ({
      totalCriticas: 4,
      totalStock: 2,
      totalCobranza: 1,
      totalPagos: 1,
      ultimoBatch: "OK",
      batchFallido: false,
      alertas: [
        { severidad: "CRITICA", tipo: "Proveedores", titulo: "Pagos urgentes por $16.250.000 en proximos 5 dias", sugerido: "Verificar caja disponible antes de la fecha de pago", destino: "proveedores" },
        { severidad: "CRITICA", tipo: "Cobranza", titulo: "Cartera vencida por $10.590.000", sugerido: "Contactar clientes con documentos vencidos", destino: "cobranza" },
        { severidad: "CRITICA", tipo: "Inventario", titulo: "Servicio Instalacion Electrica esta sin stock", sugerido: "Comprar 1 unidades", destino: "inventario" },
        { severidad: "CRITICA", tipo: "Inventario", titulo: "Servicio Mantencion Mensual esta sin stock", sugerido: "Comprar 1 unidades", destino: "inventario" },
      ],
    }),
    listarCierres: () => [],
    listarClientes: () => [],
    listarProveedores: () => [],
    listarMediosPago: () => [
      { id: "EFECTIVO", nombre: "Efectivo", activo: true, esDefault: true },
      { id: "TRANSFERENCIA", nombre: "Transferencia", activo: true, esDefault: false },
      { id: "CHEQUE", nombre: "Cheque", activo: true, esDefault: false },
      { id: "TDEBITO", nombre: "Tarjeta Debito", activo: true, esDefault: false },
      { id: "TCREDITO", nombre: "Tarjeta Credito", activo: true, esDefault: false },
      { id: "ANTICIPO", nombre: "Anticipo / Saldo a favor", activo: true, esDefault: false },
    ],
    getEERRData: () => ({
      periodo: ctx.periodo,
      ventasNetas: 0,
      costoVenta: 0,
      utilidadBruta: 0,
      gastosOperacionales: 0,
      gastosFinancieros: 0,
      otrosGastos: 0,
      resultadoOperacional: 0,
      utilidadNeta: 0,
      margenBrutoPct: 0,
      margenOperacionalPct: 0,
      margenNetoPct: 0,
      alertaCuentasSinClasificar: false,
    }),
    guardarConfigEmpresa: () => ({ ok: true }),
    guardarConfigTributario: () => ({ ok: true }),
    guardarConfigComercial: () => ({ ok: true }),
    guardarConfigInventario: () => ({ ok: true }),
    guardarConfigSeguridad: () => ({ ok: true }),
    guardarConfigAlertas: () => ({ ok: true }),
    crearMedioPago: () => ({ ok: true }),
    editarMedioPago: () => ({ ok: true }),
    toggleActivoMedioPago: () => ({ ok: true }),
    marcarMedioPagoDefault: () => ({ ok: true }),
    reordenarMediosPago: () => ({ ok: true }),
    crearAnticipo: () => ({ ok: true }),
  };

  const run = {
    _success: null,
    _failure: null,
    withSuccessHandler(callback) {
      this._success = callback;
      return this;
    },
    withFailureHandler(callback) {
      this._failure = callback;
      return this;
    },
  };

  Object.keys(responses).forEach((name) => {
    run[name] = function (...args) {
      const success = this._success;
      const failure = this._failure;
      this._success = null;
      this._failure = null;
      setTimeout(() => {
        try {
          const value = responses[name](...args);
          if (success) success(JSON.parse(JSON.stringify(value)));
        } catch (error) {
          if (failure) failure(error);
          else console.error(error);
        }
      }, 40);
      return this;
    };
  });

  window.google = { script: { run } };
})();
