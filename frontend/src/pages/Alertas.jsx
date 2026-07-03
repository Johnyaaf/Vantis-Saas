import React from "react";
import { Kpi, PageTitle } from "../components/ui";

export function Alertas() {
  const rows = [
    ["Pagos urgentes por $16.250.000 en proximos 5 dias", "Proveedores", "Ir a Proveedores"],
    ["Cartera vencida por $10.590.000", "Cobranza", "Ir a Cobranza"],
    ["Servicio Instalacion Electrica esta sin stock", "Inventario", "Ir a Inventario"],
    ["Servicio Mantencion Mensual esta sin stock", "Inventario", "Ir a Inventario"],
  ];
  return (
    <>
      <PageTitle title="Alertas" subtitle="4 alerta(s) activa(s)" />
      <section className="kpi-grid five">
        <Kpi tone="red" label="CRITICAS" value="4" detail="Requieren accion hoy" />
        <Kpi tone="orange" label="STOCK" value="2" detail="Inventario" />
        <Kpi tone="orange" label="COBRANZA" value="1" detail="Cuentas por cobrar" />
        <Kpi tone="orange" label="PAGOS" value="1" detail="Cuentas por pagar" />
        <Kpi label="TOTAL" value="4" detail="Todas las alertas" />
      </section>
      <section className="alerts">
        <h3>HOY (4)</h3>
        {rows.map(([title, area, action]) => (
          <article key={title}>
            <span>!</span>
            <div><b>CRITICA</b><h2>{title}</h2><p>{area}</p><small>Sugerido: revisar y resolver durante el dia</small></div>
            <button>{action}</button>
          </article>
        ))}
      </section>
    </>
  );
}
