# Totem Alfa – SPA verticale per magazzino (React + Vite + MSAL)

Totem touch fullscreen per operatori ALFA ENG: login Microsoft (SSO), consultazione e aggiornamento liste SharePoint (1_FORGIATI, 2_ORING-HNBR, 2_ORING-NBR, 3_TUBI), integrazione futura con Power Automate. L'interfaccia è progettata per schermi verticali 4K con pulsanti e testo molto leggibili.

## Cosa contiene già
- Bootstrap completo MSAL (login redirect + active account) con React 19 + Vite 7.
- Hook `useAuthenticatedGraphClient` per ottenere un client Graph autenticato.
- Service `SharePointService` per leggere/scrivere liste SharePoint via Graph.
- Pannelli inventario **1_FORGIATI** e **3_TUBI** con ricerca client-side, paging (100/200), colonne sticky e sort per data recente.
- Selezione/multiselezione articoli con limite a 10 elementi e carrello riepilogativo espandibile (titolo, N° Bolla, Colata) con rimozione/svuota.
- Flusso **Aggiorna Giacenza**: card touch-friendly per gli articoli selezionati, campi modificabili per FORGIATI/TUBI, salvataggio diretto su SharePoint via Graph (PATCH `/fields`).

## Struttura (principali)
```
totem-alfa-app/
├── .env.example                  # variabili ambiente (usa stessi valori del KIOSK, cambia solo GUID liste)
├── src/
│   ├── App.tsx                   # shell verticale + pannelli FORGIATI/TUBI + carrello selezioni
│   ├── styles.css                # tema totem (tipografia grande, pulsanti XL)
│   ├── auth/authConfig.ts        # config MSAL + scopes
│   ├── hooks/useAuthenticatedGraphClient.ts
│   ├── services/sharePointService.ts
│   ├── hooks/useCachedList.ts    # cache in RAM 5 minuti per liste
│   ├── config/forgiatiColumns.ts # mappa colonne -> etichette (aggiorna con internal name SharePoint)
│   └── config/tubiColumns.ts     # mappa colonne TUBI (field_1...)
└── README.md
```

## Setup rapido
1. `cd totem-alfa-app`
2. `cp .env.example .env.local` e inserisci i valori (puoi riusare quelli del KIOSK, cambiando solo i GUID delle liste magazzino).
3. `npm install`
4. `npm run dev` e fai login con account aziendale.

## Variabili d'ambiente (chiave)
- `VITE_CLIENT_ID`, `VITE_TENANT_ID`, `VITE_REDIRECT_URI`: dati registrazione SPA Entra ID.
- `VITE_SHAREPOINT_SITE_ID`, `VITE_SHAREPOINT_SITE_URL`: sito SharePoint con le liste magazzino.
- `VITE_FORGIATI_LIST_ID`, `VITE_ORING_HNBR_LIST_ID`, `VITE_ORING_NBR_LIST_ID`, `VITE_TUBI_LIST_ID`: GUID liste (fornisci almeno FORGIATI per partire).
- `VITE_SETTINGS_LIST_ID`: opzionale, per salvare configurazioni runtime (modalità auth, feature flag).
- `VITE_PA_FLOW_URL`: opzionale, endpoint Power Automate.
- `VITE_ADMIN_EMAILS`, `VITE_ADMIN_GROUP_IDS`: opzionali per pannello admin futuro.

## Note su SharePoint (FORGIATI/TUBI)
- Nel file `src/config/forgiatiColumns.ts` trovi le etichette per FORGIATI; aggiorna i campi `field` con gli **internal name** reali (usare chiavi `fields` restituite da Graph).
- Nel file `src/config/tubiColumns.ts` mappa le colonne TUBI (`field_1`, `field_2`, ...); sostituisci con gli internal name effettivi.
- In questa fase eseguiamo solo **read**; gli hook per **update/patch** sono già supportati dal `SharePointService` e possono essere agganciati a pulsanti di modifica/aggiornamento giacenze.
  - Il flusso `Aggiorna Giacenza` utilizza `SharePointService.updateItem` per patchare i campi consentiti:
    - **FORGIATI**: `field_24` (Commessa), `field_23` (Data prelievo), `field_22` (Giacenza Q.tà), `field_25` (Note), `field_21` (Giacenza mm da barra).
    - **TUBI**: `field_21` (Data ultimo prelievo), `field_25` (Commessa), `field_20` (Giacenza Tubo intero), `field_19` (Giacenza mm contabile).

## Come funziona Aggiorna Giacenza
1. Seleziona fino a 10 articoli nei pannelli FORGIATI/TUBI (carrello espandibile in alto).
2. Premi **Aggiorna Giacenza** nel carrello per aprire la vista a card.
3. Modifica solo i campi consentiti (colorazione per tipologia: blu FORGIATI, verde TUBI) e premi **Aggiorna** per salvare su SharePoint.
4. Messaggi di stato (success/error) mostrano l’esito del salvataggio; il carrello resta invariato finché non lo svuoti.

## Prossimi step suggeriti
1) Inserisci in `.env.local` i GUID di `VITE_FORGIATI_LIST_ID` e `VITE_TUBI_LIST_ID` (oltre a `VITE_SHAREPOINT_SITE_ID`).  
2) Avvia l'app e verifica popolamento FORGIATI/TUBI; allinea `forgiatiColumns.ts` e `tubiColumns.ts` con gli internal name reali.  
3) Abilita i tab ORING duplicando il pannello TUBI/forgiati e creando `oringColumns.ts` + env `VITE_ORING_*`.  
4) Collega il carrello a un Flow Power Automate o a una lista ordini/prelievi e valuta stampa PDF/QR del riepilogo.  
5) Implementa azioni di update (es. +/- giacenza, prelievo) usando `SharePointService` (patch su `/fields`).

## Deploy
- In dev basta `npm run dev`.
- Per prod `npm run build` e deploy su Azure Static Web Apps o hosting statico a scelta.

Se qualcosa non funziona controlla: consent Graph (Sites.ReadWrite.All), GUID liste, account attivo MSAL, e che l'utente abbia permessi sul sito SharePoint.
