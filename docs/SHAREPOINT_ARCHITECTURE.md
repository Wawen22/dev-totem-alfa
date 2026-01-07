# Architettura e Integrazione SharePoint - Totem Alfa

## Panoramica Architetturale

Totem Alfa è una Single Page Application (SPA) realizzata con **React** e **Vite**, progettata per funzionare come interfaccia touch verticale. L'applicazione non possiede un backend proprietario (Node.js, Python, ecc.), ma si appoggia interamente all'ecosistema **Microsoft 365** sfruttando **Microsoft Graph API**.

### Flusso dei Dati
1.  **Autenticazione**: L'app utilizza MSAL (Microsoft Authentication Library) per autenticare l'utente/device contro Azure Active Directory.
2.  **Accesso Dati**: Una volta ottenuto il token di accesso, l'app interroga Microsoft Graph (`https://graph.microsoft.com/v1.0`).
3.  **Storage**: I dati risiedono su SharePoint Online (Liste per l'inventario, Document Library per i PDF).

---

## Integrazione SharePoint & Graph API

### 1. Configurazione Azure AD (App Registration)
Per permettere all'applicazione di autenticarsi e accedere ai dati, è mandatorio registrarla su Microsoft Entra ID (ex Azure AD).

**Creazione App Registration:**
1. Accedere al [Portale Azure](https://portal.azure.com/).
2. Navigare in **Microsoft Entra ID** > **App registrations** > **New registration**.
3. **Name**: Inserire un nome (es. "Totem Alfa App").
4. **Supported account types**: Selezionare "Accounts in this organizational directory only" (Single Tenant).
5. **Redirect URI**:
   - Platform: **Single-page application (SPA)**.
   - URI: `http://localhost:5173` (per sviluppo) e l'URL di produzione quando disponibile.
6. Cliccare su **Register**.
7. Copiare i valori **Application (client) ID** e **Directory (tenant) ID** nel file `.env.local`.

**Configurazione Permessi (API Permissions):**
1. Nel menu dell'app registrata, andare su **API permissions**.
2. Cliccare **Add a permission** > **Microsoft Graph** > **Delegated permissions**.
3. Aggiungere i seguenti permessi (scopes):
   - `User.Read` (Profilo utente base)
   - `Sites.Read.All` (Lettura liste SharePoint)
   - `Files.Read.All` (Lettura documenti e drive)
   - `email`, `openid`, `profile` (Standard OIDC)
4. Cliccare **Add permissions**.
5. **Importante**: Cliccare sul pulsante **Grant admin consent for [Nome Organizzazione]**.
   - *Nota*: Senza il consenso amministrativo, agli utenti potrebbe essere bloccato l'accesso o richiesto un consenso individuale che l'UI del totem potrebbe non gestire elegantemente.

### 2. Autenticazione Lato Client (MSAL)
Il file `src/auth/authConfig.ts` configura l'istanza MSAL.
- **Client ID & Tenant ID**: Identificano l'applicazione registrata su Azure AD.
- **Scopes**: Permessi richiesti (`Sites.Read.All`, `Files.Read.All`, `User.Read`).

L'hook personalizzato `hooks/useAuthenticatedGraphClient.ts` si occupa di:
1.  Verificare se esiste un account attivo.
2.  Richiedere un token (silenziosamente se possibile).
3.  Istanziare e restituire un `Client` del Microsoft Graph SDK pronto all'uso.

### 3. Recupero Dati Inventario (Liste)
Il servizio `src/services/sharePointService.ts` gestisce le operazioni CRUD sulle liste SharePoint.
- **Metodo**: `getListItems`
- **Endpoint Graph**: `/sites/{site-id}/lists/{list-id}/items`
- **Espansione Fields**: I dati reali delle colonne SharePoint si trovano nella proprietà `fields`. La query Graph include `expand=fields` per recuperarli tutti in una sola chiamata.

### 4. Recupero e Visualizzazione Documenti (Drive)
Il componente `src/components/features/DocumentBrowser.tsx` implementa un browser di file completo.

#### A. Navigazione Cartelle
- **Concetto**: SharePoint gestisce i file in "Document Libraries", che Graph espone come "Drives".
- **Endpoint**: 
  - Root: `/sites/{site-id}/drives/{drive-id}/root/children`
  - Sottocartelle: `/sites/{site-id}/drives/{drive-id}/root:/{path-relativo}:/children`
- **Logica**: L'app naviga ricorsivamente costruendo il path. Una cartella è distinta da un file dalla presenza della proprietà `folder`.

#### B. Visualizzazione PDF (Blob & Object URL)
Per evitare che i browser scarichino forzatamente i PDF invece di visualizzarli, utilizziamo un approccio basato su Blob:
1.  **Recupero URL Download**: Ogni file restituito da Graph ha una proprietà `@microsoft.graph.downloadUrl`.
2.  **Fetch del Contenuto**: L'app effettua una `fetch()` su quell'URL (che è un link temporaneo diretto ad Azure Storage/SharePoint).
3.  **Creazione Blob**: La risposta viene convertita in un oggetto `Blob` (Binary Large Object).
4.  **Object URL**: Viene generato un URL locale temporaneo (`blob:http://localhost...`) tramite `URL.createObjectURL(blob)`.
5.  **iframe**: Questo URL viene passato come `src` di un `iframe`, forzando il browser a renderizzare il PDF internamente.

---

## Guida alla Configurazione (.env.local)

Il file `.env.local` contiene gli ID univoci necessari a Graph per trovare le risorse corrette. Ecco come recuperarli utilizzando **Graph Explorer** (https://developer.microsoft.com/graph/graph-explorer).

### 1. Recuperare `VITE_SHAREPOINT_SITE_ID`
ID univoco della Site Collection.
- **Azione**: Copia l'URL del tuo sito SharePoint (es. `https://contoso.sharepoint.com/sites/NomeSito`).
- **Query Graph Explorer**: 
  ```http
  GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/NomeSito
  ```
- **Risultato**: Cerca il campo `id` nella risposta JSON.
  - *Formato*: `hostname,guid-sito,guid-web`

### 2. Recuperare `VITE_SHAREPOINT_DRIVE_ID` (Documenti)
ID della libreria documenti specifica (es. "Documenti Condivisi" o una libreria custom).
- **Prerequisito**: Avere il Site ID dal passo 1.
- **Query Graph Explorer**:
  ```http
  GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
  ```
- **Logica**: La risposta è una lista di tutte le Document Library del sito. Cerca quella con `name` o `webUrl` corrispondente alla tua cartella target.
- **Risultato**: Copia il valore `id` di quell'oggetto.

### 3. Recuperare gli ID delle Liste (`VITE_FORGIATI_LIST_ID`, ecc.)
- **Prerequisito**: Avere il Site ID.
- **Query Graph Explorer**:
  ```http
  GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists
  ```
- **Logica**: Cerca nell'array dei risultati l'oggetto con `displayName` corrispondente alla lista desiderata (es. "1_FORGIATI").
- **Risultato**: Copia il valore `id`.

### 4. Configurare `VITE_DOCS_ROOT_PATH`
Percorso relativo opzionale per iniziare la navigazione in una sottocartella specifica.
- **Esempio**: Se la Document Library è "Documenti" e vuoi partire da "Qualità/Sicurezza", imposta questa variabile a `Qualità/Sicurezza`.
- **Nota**: Usare `/` come separatore. Non serve iniziare con `/`.

---

## Struttura Nomi Colonne (Architecture Decision Record)
SharePoint assegna nomi interni alle colonne (es. `field_1`, `field_2`) che differiscono dai nomi visualizzati. Il mapping è hardcodato nei file di configurazione (`src/config/*.ts`) per evitare chiamate API extra di "discovery" ad ogni avvio, migliorando le performance.

- **File Mapping**: `src/config/forgiatiColumns.ts`, `src/config/tubiColumns.ts`
- **Manutenzione**: Se una colonna cambia in SharePoint, è necessario aggiornare manualmente questi file verificando i nuovi nomi interni via Graph Explorer (`/items?expand=fields`).
