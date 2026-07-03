import React from "react";
import { company, money } from "../data/demoData";

export function PageTitle({ title, subtitle }) {
  return (
    <div className="page-title">
      <div>
        <h1>{title}</h1>
        <p>{subtitle}</p>
      </div>
      <span>Actualizado: {company.updated}</span>
    </div>
  );
}

export function Kpi({ tone = "blue", label, value, detail }) {
  return (
    <article className={`kpi ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
      <small>{detail}</small>
    </article>
  );
}

export function ChartCard({ title, large, bars }) {
  return (
    <article className={`chart-card ${large ? "large" : ""}`}>
      <h3>{title}</h3>
      {bars ? (
        <div className="hbars">
          <p><i style={{ width: "90%" }} />Al dia</p>
          <p><i className="red" style={{ width: "28%" }} />Vencido</p>
        </div>
      ) : (
        <div className="grid-lines" />
      )}
    </article>
  );
}

export function MiniPanel({ title, text }) {
  return <article className="mini-panel"><h3>{title}</h3><p>{text}</p></article>;
}

export function Banner({ text, action }) {
  return <section className="banner"><strong>{text}</strong>{action ? <button>{action}</button> : null}</section>;
}

export function BarList({ title, rows }) {
  const max = Math.max(...rows.map((row) => Number(row[1])), 1);
  return (
    <section className="bar-list">
      <h3>{title}</h3>
      {rows.map(([label, value, tone]) => (
        <p key={label}>
          <span>{label}</span>
          <i><b className={tone} style={{ width: `${Math.max((Number(value) / max) * 100, value ? 8 : 0)}%` }} /></i>
          <strong>{typeof value === "number" && value > 1000 ? money(value) : `${value} unid.`}</strong>
        </p>
      ))}
    </section>
  );
}

export function DataTable({ title, columns, rows, compact }) {
  return (
    <section className={`data-table ${compact ? "compact" : ""}`}>
      <h3>{title}</h3>
      <table>
        <thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
        <tbody>{rows.map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={cellIndex}>{cell}</td>)}</tr>)}</tbody>
      </table>
    </section>
  );
}

export function Tabs({ tabs, active, onChange }) {
  return <div className="tabs">{tabs.map((tab) => <button key={tab} className={active === tab ? "active" : ""} onClick={() => onChange(tab)}>{tab}</button>)}</div>;
}

export function Dot({ good, warn }) {
  return <span className={`dot ${good ? "good" : warn ? "warn" : "bad"}`} />;
}
