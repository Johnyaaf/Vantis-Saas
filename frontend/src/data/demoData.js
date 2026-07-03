export const company = {
  name: "Comercializadora Zentra Spa",
  rut: "78.149.043-1",
  period: "2026-06",
  updated: "01/07/2026",
  role: "Administrador",
};

export const emptySession = { access_token: "", usuario: null, tenant: null };

export const demoSession = {
  access_token: "demo-token",
  usuario: { nombre: "Administrador", email: "demo@vantis.local", rol: "admin" },
  tenant: { nombre: company.name, plan: "full" },
  demo: true,
};

export const navGroups = [
  ["PRINCIPAL", [["dashboard", "Dashboard"]]],
  [
    "GESTION",
    [
      ["finanzas", "Finanzas"],
      ["comercial", "Comercial"],
      ["transacciones", "Transacciones"],
      ["inventario", "Inventario"],
      ["cobranza", "Cobranza"],
      ["proveedores", "Proveedores"],
      ["alertas", "Alertas"],
    ],
  ],
  [
    "SISTEMA",
    [
      ["reportes", "Reportes"],
      ["carga", "Carga Masiva"],
      ["cierres", "Cierres"],
      ["config", "Configuracion"],
    ],
  ],
];

export const products = [
  ["SKU-011", "Toner Compatible HP LaserJet", 18, 18900],
  ["SKU-016", "Cable Electrico 2.5mm x100m", 12, 28000],
  ["SKU-017", "Pintura Latex Blanca 4L", 35, 8900],
  ["SKU-018", "Cemento Gris 25kg", 50, 5200],
  ["SKU-005", "Cafe Molido 250g", 70, 2900],
  ["SKU-003", "Detergente Liquido 3L", 60, 3200],
  ["SKU-014", "Mascarillas KN95 x10", 40, 3200],
  ["SKU-004", "Papel Bond A4 500 hojas", 45, 2800],
  ["SKU-010", "Casco de Seguridad Blanco", 25, 4500],
  ["SKU-015", "Grasa Lubricante Industrial 1kg", 22, 4800],
  ["SKU-007", "Escoba Industrial Cerdas Duras", 30, 3500],
  ["SKU-013", "Alcohol Gel 500ml", 95, 1100],
  ["SKU-001", "Aceite Vegetal Premium 1L", 85, 1200],
  ["SKU-006", "Azucar Blanca 1kg", 150, 590],
  ["SKU-002", "Arroz Grano Largo 1kg", 120, 680],
  ["SKU-008", "Bolsas de Basura 80L x50", 90, 890],
  ["SKU-009", "Guantes de Seguridad Talla M", 55, 1200],
  ["SKU-012", "Cinta Adhesiva 48mm x50m", 80, 680],
  ["SKU-019", "Servicio Instalacion Electrica", 0, 0],
  ["SKU-020", "Servicio Mantencion Mensual", 0, 0],
];

export const cxcDocs = [
  ["Industrias Metalicas Atacama", "FAC-20260607", "28-06-2026", "28-07-2026", "+26", 4200000, "ok"],
  ["Centro Automotriz Velocidad", "FAC-20260617", "01-06-2026", "01-07-2026", "-1", 3400000, "bad"],
  ["Constructora Pacifico Ltda.", "FAC-20260602", "11-06-2026", "11-07-2026", "+9", 3200000, "ok"],
  ["Vina Santa Elena SpA", "FAC-20260610", "02-06-2026", "04-07-2026", "+2", 2900000, "warn"],
  ["Farmacia Bienestar Ltda.", "FAC-20260604", "21-06-2026", "21-07-2026", "+19", 2450000, "ok"],
  ["Servicios Electricos del Sur", "FAC-20260614", "02-05-2026", "16-06-2026", "-16", 2100000, "bad"],
  ["Distribuidora Norte SpA", "FAC-20260601", "06-06-2026", "06-07-2026", "+4", 1850000, "warn"],
  ["Ferreteria Los Andes SpA", "FAC-20260609", "03-06-2026", "03-07-2026", "+1", 1680000, "warn"],
  ["Maderas del Sur SpA", "FAC-20260616", "17-04-2026", "06-06-2026", "-26", 1560000, "bad"],
  ["Agropecuaria Los Boldos Ltda.", "FAC-20260613", "12-05-2026", "21-06-2026", "-11", 1340000, "bad"],
];

export const cxpDocs = [
  ["Materiales de Construccion Andes", "FAC-20260614", "07-05-2026", "16-06-2026", "-16", 2650000, "late"],
  ["Plasticos y Polimeros Ltda.", "FAC-20260613", "17-05-2026", "21-06-2026", "-11", 890000, "late"],
  ["Confecciones Textil Sur SpA", "BOL-20260615", "27-05-2026", "23-06-2026", "-9", 420000, "late"],
  ["Alimentos del Campo SpA", "FAC-20260612", "22-05-2026", "26-06-2026", "-6", 1200000, "late"],
  ["Lubricantes y Combustibles SpA", "FAC-20260616", "01-06-2026", "01-07-2026", "-1", 4200000, "late"],
  ["Servicios Logisticos Expres SpA", "FAC-20260610", "04-06-2026", "04-07-2026", "+2", 3800000, "bad"],
  ["Comercial Andina SpA", "FAC-20260603", "21-06-2026", "21-07-2026", "+19", 3100000, "warn"],
  ["Importadora Pacific Trade SpA", "FAC-20260601", "11-06-2026", "11-07-2026", "+9", 2400000, "warn"],
  ["Empaques y Envases Ltda.", "FAC-20260607", "29-06-2026", "29-07-2026", "+27", 2200000, "warn"],
];

export const reportsTabs = [
  "Estado de Resultados",
  "Flujo de Caja",
  "Flujo de Efectivo",
  "Indicadores",
  "Consolidado",
  "Rentabilidad",
];

export function money(value) {
  return `$${Number(value).toLocaleString("es-CL")}`;
}
