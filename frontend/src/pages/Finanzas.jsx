import React from "react";
import { Banner, BarList, Kpi, PageTitle } from "../components/ui";
import { company } from "../data/demoData";

export function Finanzas() {
  return (
    <>
      <PageTitle title="Finanzas" subtitle={`Flujo de caja y liquidez - Periodo ${company.period}`} />
      <section className="kpi-grid six">
        <Kpi tone="green" label="CAJA REAL" value="$0" detail="Efectivo disponible hoy" />
        <Kpi label="INGRESOS NETOS" value="$0" detail={`Periodo ${company.period}`} />
        <Kpi tone="red" label="EGRESOS NETOS" value="$0" detail={`Periodo ${company.period}`} />
        <Kpi label="MARGEN OPERACIONAL" value="0,0%" detail="$0 neto" />
        <Kpi tone="orange" label="LIQUIDEZ 30D" value="$2.520.000" detail="Ratio 1.11" />
        <Kpi tone="green" label="COBERTURA CAJA" value="999 dias" detail="Solo caja - Sin cobrar CxC" />
      </section>
      <section className="finance-empty">
        <h3>EVOLUCION FLUJO</h3>
        <p>Acumulado historico de flujo.<br />Disponible en el proximo cierre de periodo.</p>
      </section>
      <section className="finance-bottom">
        <BarList title="PROYECCION 30 DIAS" rows={[["Caja actual", 0, "gray"], ["CxC 30 dias", 24890000, "green"], ["CxP exigible", 22370000, "red"]]} />
        <BarList title="AGING CXC" rows={[["0-30 dias", 21490000, "green"], ["31-60 dias", 0, "gray"], ["+60 dias", 0, "gray"], ["Vencida", 7190000, "red"]]} />
      </section>
      <Banner text="Sin meta de ventas configurada - configura objetivos para activar seguimiento" action="Configurar" />
    </>
  );
}
