import React from "react";
import { Banner, ChartCard, Kpi, MiniPanel, PageTitle } from "../components/ui";
import { company } from "../data/demoData";

export function Dashboard() {
  return (
    <>
      <PageTitle title="Dashboard Ejecutivo" subtitle={`Vision general completa - Periodo ${company.period}`} />
      <section className="score-card">
        <span>SCORE GERENCIAL</span>
        <div><i style={{ width: "52%" }} /></div>
        <strong>52/100</strong>
        <b>CRITICO</b>
        <em>Finanzas</em>
        <em>Comercial</em>
        <em>Operaciones</em>
      </section>
      <section className="kpi-grid five">
        <Kpi tone="orange" label="FLUJO NETO" value="$0" detail="Margen 0,0%" />
        <Kpi label="LIQUIDEZ" value="999 dias" detail="Ratio 1.11" />
        <Kpi tone="red" label="VENTAS DEL MES" value="$0" detail="Ticket $0" />
        <Kpi tone="red" label="CXC VENCIDO" value="$7.190.000" detail="Riesgo 22,4% de cartera" />
        <Kpi tone="orange" label="MARGEN OPERACIONAL" value="0,0%" detail="$0 neto" />
      </section>
      <section className="dash-grid">
        <ChartCard title="EVOLUCION DE VENTAS" large />
        <div className="stack">
          <ChartCard title="INGRESOS VS EGRESOS" />
          <ChartCard title="AGING CXC - AL DIA VS VENCIDO" bars />
          <MiniPanel title="TOP SKUS EN RIESGO DE QUIEBRE" text="Sin productos en riesgo de quiebre" />
        </div>
      </section>
      <Banner text="Acelerar cobranza de facturas vencidas y +60 dias libera $7.190.000" action="Ir a Cobranza" />
    </>
  );
}
