import { protectedResources } from "../../auth/authConfig";

export function SetupChecklist() {
  const envClientId = import.meta.env.VITE_CLIENT_ID || "--";
  const envTenantId = import.meta.env.VITE_TENANT_ID || "--";
  const envRedirect = import.meta.env.VITE_REDIRECT_URI || "--";
  const envSiteId = protectedResources.sharePointSiteId || "--";
  const envListId = protectedResources.sharePointListId || "--";
  const envAccessiId = import.meta.env.VITE_ACCESSI_LIST_ID || "--";
  const envSettingsId = import.meta.env.VITE_SETTINGS_LIST_ID || "--";
  const envFlowUrl = protectedResources.powerAutomateFlowUrl || "--";
  const envSiteUrl = import.meta.env.VITE_SHAREPOINT_SITE_URL || "https://portal.office.com/sharepoint";

  return (
    <div className="panel">
      <h3>Setup completo (Entra ID + SharePoint Visitatori + Power Automate)</h3>
      <p>Segui i passaggi del blueprint kiosk: stessa guida del README, pronta da copiare nel tenant. Compila l'env e verifica ogni step.</p>

      <div className="cta-row">
        <a className="cta-link" href="https://portal.azure.com" target="_blank" rel="noreferrer">Portal Entra ID ↗</a>
        <a className="cta-link" href="https://developer.microsoft.com/graph/graph-explorer" target="_blank" rel="noreferrer">Graph Explorer ↗</a>
        <a className="cta-link" href="https://make.powerautomate.com" target="_blank" rel="noreferrer">Power Automate ↗</a>
        <a className="cta-link" href={envSiteUrl} target="_blank" rel="noreferrer">SharePoint site ↗</a>
      </div>

      <div className="card-highlight">
        <strong>Env essenziali (copia in .env.local)</strong>
        <div className="code-block">
          VITE_CLIENT_ID={envClientId}<br />
          VITE_TENANT_ID={envTenantId}<br />
          VITE_REDIRECT_URI={envRedirect}<br />
          VITE_SHAREPOINT_SITE_ID={envSiteId}<br />
          VITE_SHAREPOINT_LIST_ID={envListId}<br />
          VITE_ACCESSI_LIST_ID={envAccessiId}<br />
          VITE_SETTINGS_LIST_ID={envSettingsId}<br />
          VITE_PA_FLOW_URL={envFlowUrl}
        </div>
      </div>

      <div className="list">
        <div className="list-item">
          <strong>1) Entra ID – App Registration (SPA)</strong>
          <span>Portal Azure → Entra ID → App registrations → New registration.</span>
          <span>Name: Kiosk-Visitatori (o cliente); Account type: Single tenant.</span>
          <span>Redirect URI: SPA → http://localhost:5173 (aggiungi prod dopo).</span>
          <span>Overview → copia Application (client) ID → VITE_CLIENT_ID, Directory (tenant) ID → VITE_TENANT_ID.</span>
          <span>API Permissions → Microsoft Graph → Delegated: User.Read, Sites.Read.All, Sites.ReadWrite.All → Grant admin consent.</span>
          <span><strong>App role Totem.Admin (opzionale)</strong>: App roles → Create app role (Totem Admin / Totem.Admin) → Enterprise applications → Users and groups → assegna; verifica claim roles su jwt.ms.</span>
        </div>
        <div className="list-item">
          <strong>2) SharePoint (lista Visitatori)</strong>
          <span>Crea/riusa un sito SharePoint dedicato.</span>
          <span>Crea lista “Visitatori” con colonne: Title (ID), Nome, Cognome, Email, Azienda, Categoria (Choice), Stato (Choice).</span>
          <span>Recupera Site ID via Graph Explorer: GET https://graph.microsoft.com/v1.0/sites/&lt;dominio&gt;:/sites/&lt;sito&gt; → campo id → VITE_SHAREPOINT_SITE_ID.</span>
          <span>Recupera List ID (GUID tra %7B e %7D dall’URL impostazioni lista) → VITE_SHAREPOINT_LIST_ID.</span>
          <span>Permessi: operatori Read su Visitatori; se scrivono, Contribute/Edit.</span>
        </div>
        <div className="list-item">
          <strong>3) Power Automate (opzionale)</strong>
          <span>Instant cloud flow con trigger “When an HTTP request is received”.</span>
          <span>Schema JSON: azione send|resend|otp (vedi docs), aggiungi azioni email nei case.</span>
          <span>Salva l’HTTP POST URL → VITE_PA_FLOW_URL.</span>
        </div>
        <div className="list-item">
          <strong>4) Variabili & avvio</strong>
          <span>Copia .env.example → .env.local e incolla i valori sopra.</span>
          <span>npm install && npm run dev, poi login con un utente del tenant.</span>
          <span>Usa “Crea item demo” per validare i permessi sulla lista Visitatori.</span>
        </div>
      </div>
      <div className="card-highlight" style={{ marginTop: 12 }}>
        <div><strong>Env correnti (runtime)</strong></div>
        <div className="code-block">
          <div>siteId: {envSiteId}</div>
          <div>visitatori listId: {envListId}</div>
          <div>accessi listId: {envAccessiId}</div>
          <div>settings listId: {envSettingsId}</div>
          <div>power automate: {envFlowUrl}</div>
        </div>
      </div>
    </div>
  );
}
