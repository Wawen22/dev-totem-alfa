# Totem Alfa – Implementation Plan

## 1) Obiettivi
- Totem fullscreen verticale per operatori ALFA ENG (touch-friendly, testo grande).
- Auth SSO Microsoft (MSAL) + permessi Graph `Sites.ReadWrite.All`.
- CRUD su liste SharePoint: 1_FORGIATI (oggi), 2_ORING-HNBR, 2_ORING-NBR, 3_TUBI.
- Futuro: trigger Power Automate per notifiche / email / logica.

## 2) Architettura
- **Frontend**: React 19 + Vite 7 + TypeScript.
- **Auth**: MSAL browser + MSAL React (`loginRedirect`, cache `localStorage`).
- **Data Layer**: Graph SDK v3 + `SharePointService` (listItems, createItem, updateItem se serve via patch) parametrizzato da SiteId + ListId.
- **Config**: `.env.local` (usa gli stessi valori del KIOSK, cambia solo i GUID delle liste magazzino).
- **UI**: layout verticale con tiles di sezione + pannelli tabellari, palette chiara, pulsanti XL.
- **Ruoli (App Role)**: App role `Totem.Admin` lato Entra ID per mostrare il CTA "Pannello Admin" solo agli utenti/gruppi assegnati.

## 3) Dati – Lista 1_FORGIATI
- Colonne di business in `src/config/forgiatiColumns.ts` (etichette già inserite).
- Dopo aver collegato la lista, apri il pannello FORGIATI e controlla le "Chiavi SharePoint rilevate" (debug). Copia quegli internal name dentro `forgiatiColumns.ts` al posto dei placeholder.
- Se servono default value o mapping (es. formattazione numeri Ø), aggiungi una funzione di formatter nel componente `ForgiatiPanel`.

## 4) Passi operativi (oggi)
1. **ID Lista**: crea/recupera il GUID di 1_FORGIATI e inseriscilo in `.env.local` come `VITE_FORGIATI_LIST_ID` (oltre a `VITE_SHAREPOINT_SITE_ID`).
2. **Scopes**: verifica che l'app registrata abbia `Sites.ReadWrite.All` con admin consent.
3. **Run dev**: `npm install && npm run dev` dentro `totem-alfa-app`, login Microsoft, verifica popolamento tabella FORGIATI e chiavi debug.
4. **Allinea colonne**: aggiorna `forgiatiColumns.ts` con gli internal name reali.
5. **Abilitare edit**: aggiungi azioni (es. pulsanti +/- su giacenza) usando `SharePointService` (metodo `createItem` già pronto; per update usa patch su `/fields`).
6. **Ruolo Admin (Totem.Admin)**:
	- In **App roles** crea `Totem Admin` (Value `Totem.Admin`, Allowed member types: Users/Groups, descrizione: "Consente accesso al pannello amministrazione totem").
	- In **Enterprise applications** assegna utenti/gruppi al ruolo `Totem Admin`.
	- Re-login e verifica su https://jwt.ms che nel claim `roles` compaia `Totem.Admin`. Solo con questo claim compare il CTA "Pannello Admin" in basso a destra.

## 5) Iterazioni successive
- **Sezioni mancanti**: duplicare `ForgiatiPanel` per ORING-HNBR, ORING-NBR, TUBI con relativi env `VITE_ORING_*` e colonne dedicate.
- **Admin/Settings**: collegare `VITE_SETTINGS_LIST_ID` per feature flag (modalità QR/EMAIL ecc.).
- **Power Automate**: usare `VITE_PA_FLOW_URL` per invocare flow HTTP (es. notifica su variazioni giacenza).
- **UI**: aggiungere gesture touch (scroll inertia, pulsanti a prova di guanti) e modalità high-contrast.
- **Analytics**: log degli eventi UI e audit variazioni stock in SharePoint list dedicata o Application Insights.

## 6) QA rapido
- Login/Logout: utente senza permessi su sito deve ricevere 403 informativo.
- Dati: la tabella deve mostrare almeno 1 item e le chiavi debug devono essere coerenti con la lista.
- Performance: massimo 50-100 item per fetch; valutare pagination/filters se la lista cresce.

## 7) Deploy
- Build: `npm run build` → output statico in `dist/`.
- Hosting suggerito: Azure Static Web Apps o Storage static website; configura redirect URI per produzione in Entra ID.
