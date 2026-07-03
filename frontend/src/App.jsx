import React, { useState } from "react";
import { AuthScreen } from "./auth/AuthScreen";
import { Sidebar } from "./layout/Sidebar";
import { Topbar } from "./layout/Topbar";
import { ViewRouter } from "./pages/ViewRouter";
import { demoSession, emptySession } from "./data/demoData";

function readSession() {
  try {
    return JSON.parse(localStorage.getItem("vantis.session")) ?? emptySession;
  } catch {
    return emptySession;
  }
}

export function App() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("react") !== "1") {
    return <iframe className="legacy-frame" title="VANTIS ERP migrado" src="/legacy-vantis.html" />;
  }

  const [session, setSession] = useState(readSession);
  const [authView, setAuthView] = useState("login");
  const [activeView, setActiveView] = useState("dashboard");

  function saveSession(next) {
    localStorage.setItem("vantis.session", JSON.stringify(next));
    setSession(next);
  }

  function signOut() {
    localStorage.removeItem("vantis.session");
    setSession(emptySession);
  }

  if (!session.access_token) {
    return <AuthScreen view={authView} onViewChange={setAuthView} onLogin={saveSession} onDemo={() => saveSession(demoSession)} />;
  }

  return (
    <div className="erp-shell">
      <Sidebar activeView={activeView} onChange={setActiveView} />
      <div className="erp-workspace">
        <Topbar onSignOut={signOut} demo={session.demo} />
        <main className="erp-page">
          <ViewRouter activeView={activeView} />
        </main>
      </div>
    </div>
  );
}
