import React from "react";
import { PageTitle } from "../components/ui";

export function Carga() {
  const blocks = [
    "Clientes",
    "Proveedores",
    "Productos",
    "Saldos Iniciales - Inventario",
    "Saldos Iniciales - Cuentas por Cobrar",
    "Saldos Iniciales - Cuentas por Pagar",
  ];
  return (
    <>
      <PageTitle title="Carga Masiva Inicial" subtitle="Descarga la plantilla CSV, completala y subela de vuelta" />
      <section className="upload-grid">
        {blocks.map((block) => <article key={block}><h2>{block}</h2><p>Catalogo o deuda inicial para migracion.</p><div><button>Descargar Plantilla</button><button>Subir CSV</button></div></article>)}
      </section>
    </>
  );
}
