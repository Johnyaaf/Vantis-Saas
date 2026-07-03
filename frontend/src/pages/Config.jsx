import React, { useState } from "react";
import { DataTable, Dot, Kpi } from "../components/ui";
import { company } from "../data/demoData";

const configTabs = ["Empresa", "Tributario", "Comercial", "Medios de Pago", "Inventario", "Seguridad", "Alertas", "Respaldos", "Sistema"];

export function Config() {
  const [tab, setTab] = useState("Empresa");
  return (
    <section className="config-shell">
      <aside>
        <h2>Configuracion</h2>
        <p>Centro de administracion VANTIS</p>
        {configTabs.map((item) => <button key={item} className={tab === item ? "active" : ""} onClick={() => setTab(item)}>{item}</button>)}
      </aside>
      <div className="config-body"><ConfigBody tab={tab} /></div>
    </section>
  );
}

function ConfigBody({ tab }) {
  if (tab === "Medios de Pago") {
    return <DataTable title="MEDIOS DE PAGO" columns={["Orden", "Nombre", "Activo", "Default", "Acciones"]} rows={["Efectivo", "Transferencia", "Cheque", "Tarjeta Debito", "Tarjeta Credito", "Anticipo / Saldo a favor"].map((method, index) => ["Subir/Bajar", method, <Dot good />, index === 0 ? "Default" : "Marcar", "Editar  Desactivar"])} />;
  }
  if (tab === "Sistema") {
    return <section className="config-cards"><Kpi label="VERSION ERP" value="1.0.0" detail="VANTIS ERP" /><Kpi label="EMPRESA ACTIVA" value={company.name} /><Kpi label="PERIODO ACTIVO" value={company.period} /><Kpi label="BASE DE DATOS" value="PostgreSQL" detail="Migracion SaaS" /></section>;
  }
  if (tab === "Respaldos") {
    return <section><h3>RESPALDOS</h3><h2>02-07-2026, 8:55:45 p. m.</h2><button className="blue-btn">Exportar Respaldo Completo</button><p className="muted">Restaurar respaldo disponible en proxima fase</p></section>;
  }
  return <FormMock tab={tab} />;
}

function FormMock({ tab }) {
  const fields = tab === "Empresa"
    ? ["Razon social *", "RUT *", "Nombre de fantasia", "Giro", "Direccion", "Ciudad", "Region", "Telefono", "Correo", "Logo (URL)", "Moneda", "Zona horaria"]
    : ["Parametro principal", "Valor objetivo", "Umbral minimo", "Limite maximo", "Politica", "Estado"];
  return (
    <form className="settings-form">
      <h3>{tab.toUpperCase()}</h3>
      {fields.map((field, index) => <label key={field}>{field}<input defaultValue={index === 0 ? company.name : index === 1 ? company.rut : ""} /></label>)}
      <button className="save-btn">Guardar cambios</button>
    </form>
  );
}
