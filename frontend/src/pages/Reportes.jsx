import React, { useState } from "react";
import { DataTable, Kpi, PageTitle, Tabs } from "../components/ui";
import { company, reportsTabs } from "../data/demoData";

export function Reportes() {
  const [tab, setTab] = useState("Estado de Resultados");
  return (
    <>
      <PageTitle title="Reportes Financieros" subtitle={`Periodo ${company.period}`} />
      <Tabs tabs={reportsTabs} active={tab} onChange={setTab} />
      <ReportBody tab={tab} />
    </>
  );
}

function ReportBody({ tab }) {
  if (tab === "Indicadores") {
    return (
      <ReportFrame title="INDICADORES GERENCIALES">
        <section className="report-cards">
          <Kpi label="DIAS DE INVENTARIO" value="999 dias" detail="El inventario tarda en rotar." />
          <Kpi label="DSO - DIAS DE COBRO" value="4 dias" detail="La cobranza esta bajo control." />
          <Kpi label="DPO - DIAS DE PAGO" value="6 dias" detail="Hay margen para negociar plazos." />
          <Kpi tone="red" label="CCC - CICLO DE CAJA" value="997 dias" detail="El ciclo de caja es largo." />
        </section>
      </ReportFrame>
    );
  }
  if (tab === "Consolidado") {
    return (
      <ReportFrame title="CONSOLIDADO FINANCIERO Y GERENCIAL">
        <section className="report-cards">
          <Kpi label="VENTAS TOTALES" value="$0" />
          <Kpi tone="green" label="UTILIDAD NETA" value="$0" />
          <Kpi tone="green" label="CAJA REAL (NETA)" value="$0" />
          <Kpi label="CUENTAS POR COBRAR" value="$32.080.000" />
          <Kpi label="CUENTAS POR PAGAR" value="$32.870.000" />
        </section>
        <p className="report-alert">Critico - Cartera vencida representa 33% del total por cobrar</p>
      </ReportFrame>
    );
  }
  if (tab === "Rentabilidad") {
    return (
      <ReportFrame title="RENTABILIDAD DETALLADA">
        <DataTable title="RENTABILIDAD POR CLIENTE (0)" columns={["Cliente", "Ventas", "Margen", "Margen %", "Docs"]} rows={[["Sin datos del periodo", "", "", "", ""]]} />
        <DataTable title="RENTABILIDAD POR PRODUCTO (0)" columns={["Producto", "Ventas", "Margen", "Margen %", "Cantidad"]} rows={[["Sin datos del periodo", "", "", "", ""]]} />
      </ReportFrame>
    );
  }
  return (
    <ReportFrame title={tab.toUpperCase()}>
      <table className="report-table">
        <tbody>{["Ventas Netas", "Costo de Venta", "Utilidad Bruta", "Gastos Operacionales", "Resultado Operacional", "UTILIDAD NETA DEL PERIODO"].map((row) => <tr key={row}><td>{row}</td><td>$0</td></tr>)}</tbody>
      </table>
    </ReportFrame>
  );
}

function ReportFrame({ title, children }) {
  return <section className="report"><button>Exportar PDF</button><h3>{title}</h3><h2>{company.name}</h2><p>Mes de {company.period} - Cifras en pesos chilenos (CLP)</p>{children}</section>;
}
