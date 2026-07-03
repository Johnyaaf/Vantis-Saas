import React from "react";
import { Banner, DataTable, Kpi, PageTitle } from "../components/ui";
import { company } from "../data/demoData";

export function Comercial() {
  return (
    <>
      <PageTitle title="Comercial" subtitle={`Ventas y rentabilidad comercial - Periodo ${company.period}`} />
      <section className="kpi-grid five">
        <Kpi label="VENTAS NETAS" value="$0" detail={`Periodo ${company.period}`} />
        <Kpi label="TICKET PROMEDIO" value="$0" detail="Por documento" />
        <Kpi tone="orange" label="MARGEN BRUTO COM." value="0,0%" detail="Solo producto - Sin gastos op." />
        <Kpi label="DOCUMENTOS MES" value="0" detail="Emitidos periodo" />
        <Kpi label="META VS REAL" value="Sin meta" detail="Configurar" />
      </section>
      <section className="commercial-layout">
        <div>
          <h3>POR TIPO DOCUMENTO</h3>
          <p>Sin ventas del periodo</p>
        </div>
        <div className="commercial-bottom">
          <DataTable title="TOP PRODUCTOS - MAS RENTABLES" columns={["#", "Producto", "Unid", "Monto", "Margen"]} rows={[["", "Sin datos del periodo", "", "", ""]]} compact />
          <DataTable title="TOP CLIENTES DEL PERIODO" columns={["#", "Cliente", "Monto", "Docs"]} rows={[["", "Sin datos del periodo", "", ""]]} compact />
        </div>
      </section>
      <Banner text={`Sin ventas registradas hoy - ${company.updated}`} />
    </>
  );
}
