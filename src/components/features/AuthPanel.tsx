import { useMsal } from "@azure/msal-react";
import { AccountBadge } from "../auth/AccountBadge";
import { LoginButton } from "../auth/LoginButton";
import { LogoutButton } from "../auth/LogoutButton";

export function AuthPanel() {
  const { accounts } = useMsal();
  const isAuthenticated = accounts.length > 0;

  return (
    <div className="panel">
      <h3>Autenticazione Entra ID</h3>
      <p>MSAL React gestisce login, sessione e silent token refresh per chiamare Graph.</p>
      <div className="button-row" style={{ marginBottom: 12 }}>
        {isAuthenticated ? <LogoutButton /> : <LoginButton />}
        <AccountBadge />
      </div>
      <div className="card-highlight">
        <div><strong>Flusso</strong></div>
        <div>1. Login redirect (PKCE)</div>
        <div>2. acquireTokenSilent() per Graph</div>
        <div>3. Token in cache localStorage</div>
      </div>
    </div>
  );
}
