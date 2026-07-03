import React from "react";
import { company } from "../data/demoData";

export function Topbar({ onSignOut, demo }) {
  return (
    <header className="erp-topbar">
      <div className="top-meta"><span>EMPRESA</span><strong>{company.name}</strong></div>
      <div className="top-meta"><span>PERIODO</span><strong>{company.period}</strong></div>
      <div className="top-meta"><span>ACTUALIZADO</span><strong>{company.updated}</strong></div>
      <div className="top-user">
        <span className="live-dot">LIVE</span>
        {demo ? <span>Demo local</span> : null}
        <button onClick={onSignOut}>{company.role}</button>
      </div>
    </header>
  );
}
