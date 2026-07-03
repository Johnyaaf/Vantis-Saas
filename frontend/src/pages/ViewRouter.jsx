import React from "react";
import { Alertas } from "./Alertas";
import { Cartera } from "./Cartera";
import { Carga } from "./Carga";
import { Cierres } from "./Cierres";
import { Comercial } from "./Comercial";
import { Config } from "./Config";
import { Dashboard } from "./Dashboard";
import { Finanzas } from "./Finanzas";
import { Inventario } from "./Inventario";
import { Reportes } from "./Reportes";
import { Transacciones } from "./Transacciones";

export function ViewRouter({ activeView }) {
  if (activeView === "dashboard") return <Dashboard />;
  if (activeView === "finanzas") return <Finanzas />;
  if (activeView === "comercial") return <Comercial />;
  if (activeView === "transacciones") return <Transacciones />;
  if (activeView === "inventario") return <Inventario />;
  if (activeView === "cobranza") return <Cartera type="cxc" />;
  if (activeView === "proveedores") return <Cartera type="cxp" />;
  if (activeView === "alertas") return <Alertas />;
  if (activeView === "reportes") return <Reportes />;
  if (activeView === "carga") return <Carga />;
  if (activeView === "cierres") return <Cierres />;
  if (activeView === "config") return <Config />;
  return <Dashboard />;
}
