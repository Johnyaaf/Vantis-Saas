import React from "react";
import { DataTable, PageTitle } from "../components/ui";
import { company } from "../data/demoData";

export function Cierres() {
  return (
    <>
      <PageTitle title="Cierre de Periodos" subtitle="Congela las cifras del mes - irreversible sin clave de administrador" />
      <section className="close-page"><h3>PERIODO ACTUAL</h3><strong>{company.period}</strong><button>Cerrar Periodo {company.period}</button></section>
      <DataTable title="HISTORIAL DE CIERRES" columns={["Periodo", "Fecha Cierre", "Usuario", "Ventas Netas", "Utilidad Neta", "Estado"]} rows={[["Sin cierres registrados", "", "", "", "", ""]]} />
    </>
  );
}
