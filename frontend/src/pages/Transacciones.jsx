import React from "react";
import { PageTitle } from "../components/ui";

export function Transacciones() {
  const actions = [
    ["Nueva Venta", "Boleta, Factura, Nota de Credito o Nota de Debito a clientes"],
    ["Nueva Compra / Gasto", "Compra de mercaderia, gastos, activos o Boleta de Honorarios"],
    ["Ver Clientes", "Busca y abre la ficha de un cliente existente"],
    ["Ver Proveedores", "Busca y abre la ficha de un proveedor existente"],
  ];
  return (
    <>
      <PageTitle title="Transacciones" subtitle="Registra ventas, compras y gastos del negocio" />
      <section className="action-grid">
        {actions.map(([title, text]) => <article key={title}><span>*</span><h2>{title}</h2><p>{text}</p></article>)}
      </section>
    </>
  );
}
