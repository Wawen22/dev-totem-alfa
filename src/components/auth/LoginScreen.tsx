import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../../auth/authConfig";
import "./login.css";

export function LoginScreen() {
  const { instance } = useMsal();

  const handleLogin = async () => {
    await instance.loginRedirect(loginRequest);
  };

  return (
    <div className="login-shell">
      <div className="login-backdrop">
        <div className="login-grid-lines" />
        <div className="login-orb orb-left" />
        <div className="login-orb orb-right" />
      </div>

      <div className="login-content">
        <section className="login-hero">
          <div className="hero-chip">Microsoft 365 • SPA</div>
          <h1>Accesso sicuro con Entra ID</h1>
          <p className="hero-lead">
            Avvia il tuo progetto serverless: autenticazione SSO, dati su SharePoint e logiche Power Automate, tutto dal frontend.
          </p>

          <div className="hero-meta">
            <div className="meta-card">
              <span className="meta-label">Identità</span>
              <span className="meta-value">Microsoft Entra ID</span>
            </div>
            <div className="meta-card">
              <span className="meta-label">Database</span>
              <span className="meta-value">SharePoint Lists via Graph</span>
            </div>
            <div className="meta-card">
              <span className="meta-label">Automazioni</span>
              <span className="meta-value">Power Automate HTTP</span>
            </div>
          </div>

          <ul className="hero-list">
            <li><span className="list-dot" />Single Sign-On enterprise con PKCE</li>
            <li><span className="list-dot" />CRUD diretto verso Microsoft Graph</li>
            <li><span className="list-dot" />Flow HTTP trigger per email/OTP</li>
          </ul>
        </section>

        <section className="login-panel">
          <div className="panel-head">
            <span className="panel-badge">Accesso</span>
            <h2>Entra con il tuo account Microsoft</h2>
            <p>SSO aziendale, token sicuri in cache locale, nessuna password memorizzata.</p>
          </div>

          <button onClick={handleLogin} className="login-btn" aria-label="Accedi con Microsoft">
            <span className="m365-mark" aria-hidden="true">
              <span className="tile" />
              <span className="tile" />
              <span className="tile" />
              <span className="tile" />
            </span>
            <span className="btn-text">
              <strong>Accedi con Microsoft</strong>
              <small>Entra ID • MSAL • PKCE</small>
            </span>
            <span className="btn-arrow" aria-hidden="true">→</span>
          </button>

          <div className="panel-footer">
            <div className="foot-chip">
              <span className="chip-dot" />Sessione gestita da Azure; token per Graph inclusi.
            </div>
            <p className="foot-note">Configura l'env e premi login per avviare il blueprint.</p>
          </div>
        </section>
      </div>
    </div>
  );
}
