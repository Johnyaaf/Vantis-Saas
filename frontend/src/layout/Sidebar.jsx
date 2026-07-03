import React from "react";
import { navGroups } from "../data/demoData";

export function Sidebar({ activeView, onChange }) {
  return (
    <aside className="side">
      <div className="side-brand"><span></span><strong>VANTIS</strong></div>
      {navGroups.map(([group, items]) => (
        <nav key={group} className="side-group">
          <p>{group}</p>
          {items.map(([id, label]) => (
            <button key={id} className={activeView === id ? "active" : ""} onClick={() => onChange(id)}>
              <span className="nav-icon">{label.slice(0, 1)}</span>{label}{id === "alertas" ? <b>0</b> : null}
            </button>
          ))}
        </nav>
      ))}
      <small>VANTIS ERP v1.0</small>
    </aside>
  );
}
