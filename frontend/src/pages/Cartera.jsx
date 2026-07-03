import React from "react";
import { Banner, BarList, DataTable, Dot, Kpi, PageTitle } from "../components/ui";
import { company, cxcDocs, cxpDocs, money } from "../data/demoData";

export function Cartera({ type }) {
  const isCxc = type === "cxc";
  const docs = isCxc ? cxcDocs : cxpDocs;
  return (
    <>
      <PageTitle title={isCxc ? "Cobranza" : "Proveedores"} subtitle={`${isCxc ? "Cuentas por cobrar y gestion de cartera" : "Cuentas por pagar y gestion de obligaciones"} - Periodo ${company.period}`} />
      <section className="kpi-grid five">
        <Kpi label={isCxc ? "CXC TOTAL" : "CXP TOTAL"} value={money(isCxc ? 32080000 : 32870000)} detail={isCxc ? "Cartera pendiente" : "Obligaciones pendientes"} />
        <Kpi tone="red" label={isCxc ? "CXC VENCIDA" : "CXP VENCIDA"} value={money(isCxc ? 10590000 : 10320000)} detail="En mora" />
        <Kpi label="DOCS PENDIENTES" value="20" detail={isCxc ? "Por cobrar" : "Por pagar"} />
        <Kpi tone="red" label="DOCS VENCIDOS" value={isCxc ? "7" : "6"} detail="En mora" />
        <Kpi tone="green" label={isCxc ? "DSO (DIAS COBRO PROM.)" : "DPO (DIAS PAGO PROM.)"} value={isCxc ? "4 dias" : "6 dias"} detail={isCxc ? "Meta <= 35 dias" : "Plazo compra: 30 dias"} />
      </section>
      <section className="split">
        <BarList title="AGING DETALLADO" rows={isCxc ? [["0-30 dias", 21490000, "green"], ["31-60 dias", 0, "gray"], ["+60 dias", 0, "gray"], ["Vencida", 10590000, "red"]] : [["Urgente 0-5d", 16250000, "red"], ["Proximo 6-30d", 16620000, "orange"], ["Planificado +30d", 0, "gray"], ["Mora", 10320000, "black"]]} />
        <DataTable title={isCxc ? "TOP DEUDORES" : "TOP PROVEEDORES"} compact columns={["#", isCxc ? "Cliente" : "Proveedor", "Saldo", "%"]} rows={docs.slice(0, 5).map((doc, index) => [index + 1, doc[0], money(doc[5]), `${13 - index}%`])} />
      </section>
      <DataTable title={isCxc ? "DOCUMENTOS POR COBRAR" : "DOCUMENTOS POR PAGAR"} columns={[isCxc ? "Cliente" : "Proveedor", "Documento", "Emision", "Vence", "Dias", "Saldo", "Estado"]} rows={docs.map((doc) => [doc[0], doc[1], doc[2], doc[3], doc[4], money(doc[5]), <Dot good={doc[6] === "ok"} warn={doc[6] === "warn"} />])} />
      <Banner text={isCxc ? "$10.590.000 en 7 documento(s) vencido(s) - Contactar hoy para evitar mora" : "ALERTA: $16.250.000 vence en 5 dias - Caja actual $0 - Faltan $16.250.000"} action={isCxc ? "Ver Vencidos" : ""} />
    </>
  );
}
