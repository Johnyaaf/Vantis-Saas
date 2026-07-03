import React, { useMemo } from "react";
import { BarList, DataTable, Dot, Kpi, PageTitle } from "../components/ui";
import { company, money, products } from "../data/demoData";

export function Inventario() {
  const total = useMemo(() => products.reduce((sum, [, , stock, cost]) => sum + stock * cost, 0), []);
  return (
    <>
      <PageTitle title="Inventario" subtitle={`Stock y valorizacion - Periodo ${company.period}`} />
      <section className="kpi-grid five">
        <Kpi label="PRODUCTOS ACTIVOS" value="18" detail="En catalogo" />
        <Kpi label="VALOR INVENTARIO" value={money(total)} detail="A costo" />
        <Kpi tone="green" label="PRODUCTOS CRITICOS" value="0" detail="Bajo nivel critico" />
        <Kpi tone="green" label="SIN STOCK" value="0" detail="Quiebre actual" />
        <Kpi label="ROTACION" value="999 dias" detail="Cobertura stock actual" />
      </section>
      <BarList title="STOCK POR PRODUCTO" rows={products.slice(0, 9).map((product) => [product[1], product[2], "green"])} />
      <DataTable
        title="FICHA DE PRODUCTOS"
        columns={["#", "SKU", "Nombre", "Stock", "Min.", "Costo", "Valor", "Dias sin venta", "Estado"]}
        rows={products.map((product, index) => [index + 1, product[0], product[1], product[2], index % 4 === 0 ? 4 : 10, money(product[3]), money(product[2] * product[3]), "Sin ventas", <Dot good={product[2] > 0} />])}
      />
    </>
  );
}
