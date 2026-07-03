import React, { useState } from "react";
import { login, registerTenant } from "../api/auth";
import { ApiError } from "../api/client";

export function AuthScreen({ view, onViewChange, onLogin, onDemo }) {
  return (
    <main className="auth-page">
      <section className="auth-copy">
        <span className="brand-pill">VANTIS ERP</span>
        <h1>ERP SaaS para operar una pyme chilena con control real.</h1>
        <p>Frontend migrado como aplicacion SaaS: sin Google Sheets, sin Apps Script y preparado para conectar al backend FastAPI.</p>
        <div className="auth-metrics"><span>FastAPI</span><span>PostgreSQL</span><span>Multi-tenant</span></div>
      </section>
      <section className="auth-card">
        <div className="auth-tabs">
          <button className={view === "login" ? "active" : ""} onClick={() => onViewChange("login")}>Ingresar</button>
          <button className={view === "register" ? "active" : ""} onClick={() => onViewChange("register")}>Registrar empresa</button>
        </div>
        {view === "login" ? <LoginForm onLogin={onLogin} /> : <RegisterForm onDone={() => onViewChange("login")} />}
        <button className="demo-button" type="button" onClick={onDemo}>Entrar y ver ERP demo</button>
      </section>
    </main>
  );
}

function LoginForm({ onLogin }) {
  const [form, setForm] = useState({ email: "", password: "" });
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setError("");
    try {
      onLogin(await login(form));
    } catch (err) {
      setError(err instanceof ApiError ? err.detail : "Backend sin base de datos activa. Usa el modo demo para ver el ERP.");
    }
  }

  return (
    <form className="form" onSubmit={submit}>
      <h2>Ingresar</h2>
      <label>Email<input type="email" required value={form.email} onChange={(event) => setForm({ ...form, email: event.target.value })} /></label>
      <label>Contrasena<input type="password" required value={form.password} onChange={(event) => setForm({ ...form, password: event.target.value })} /></label>
      {error ? <p className="error">{error}</p> : null}
      <button className="primary-button">Ingresar</button>
    </form>
  );
}

function RegisterForm({ onDone }) {
  const [form, setForm] = useState({ nombre: "", email: "", password: "", rut_empresa: "", nombre_empresa: "", plan: "full" });
  const [message, setMessage] = useState("");

  async function submit(event) {
    event.preventDefault();
    try {
      await registerTenant(form);
      setMessage("Empresa registrada. Ya puedes iniciar sesion.");
      setTimeout(onDone, 900);
    } catch {
      setMessage("Registro real requiere PostgreSQL activo. Puedes revisar la app con el modo demo.");
    }
  }

  return (
    <form className="form" onSubmit={submit}>
      <h2>Registrar empresa</h2>
      <label>Nombre administrador<input required value={form.nombre} onChange={(event) => setForm({ ...form, nombre: event.target.value })} /></label>
      <label>Email<input type="email" required value={form.email} onChange={(event) => setForm({ ...form, email: event.target.value })} /></label>
      <label>Contrasena<input type="password" required value={form.password} onChange={(event) => setForm({ ...form, password: event.target.value })} /></label>
      <label>RUT empresa<input required value={form.rut_empresa} onChange={(event) => setForm({ ...form, rut_empresa: event.target.value })} /></label>
      <label>Razon social<input required value={form.nombre_empresa} onChange={(event) => setForm({ ...form, nombre_empresa: event.target.value })} /></label>
      <button className="primary-button">Crear empresa</button>
      {message ? <p className="success">{message}</p> : null}
    </form>
  );
}
