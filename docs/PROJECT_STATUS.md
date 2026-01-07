# Totem Alfa – Documentazione Progetto & Stato Avanzamento

**Ultimo Aggiornamento:** 5 Gennaio 2026
**Versione:** 0.2 (Alpha Funzionante)

## 1. Panoramica
Applicazione Totem touch-friendly verticale per la gestione dell'inventario ALFA ENG. Integrata con Microsoft SharePoint via Microsoft Graph API. L'app è progettata per essere utilizzata su schermi verticali (1080x1920) in magazzino.

## 2. Funzionalità Implementate

### A. Dashboard Principale
- **Design Moderno:** Entry point con Hero banner e widget Data/Ora 3D in tempo reale.
- **Navigazione:** Menu principale a card verticali ("pulsantoni") per accesso rapido.
- **Sezioni:**
  - **Giacenza Magazzino:** Attiva e funzionante.
  - **Documentazione:** Attiva e Funzionante (File Browser & PDF Viewer).
  - **Sito Web:** Placeholder disabilitato per sviluppi futuri.

### B. Gestione Inventario (Giacenza Magazzino)
- **Liste Supportate:**
  1. **1_FORGIATI:** Implementata completa.
  2. **3_TUBI:** Implementata completa.
  3. **2_ORING-HNBR / NBR:** Pulsanti presenti ma disabilitati (Next Step).
- **Tabella Dati Evoluta:**
  - **Scroll Orizzontale & Verticale:** Ottimizzato per touch, senza impaginazione forzata (caricamento full dataset).
  - **Sticky Columns:** Codice Articolo e Quantità sempre visibili durante lo scorrimento.
  - **Zebra Striping:** Righe a colori alterni per leggibilità.
  - **Ordinamento:** Automatico per data (dal più recente).
  - **Formattazione:** Date convertite automaticamente da seriali Excel o ISO a formato leggibile "dd/MM/yyyy".

### D. Selezione articoli e Carrello
- **Multiselezione:** Checkbox su FORGIATI e TUBI con limite massimo di 10 articoli totali.
- **Carrello riepilogo:** Pannello espandibile con tipo lista, titolo articolo, N° Bolla e Colata, azioni di rimozione singola e svuota tutto; colorazione riga per tipologia (blu Forgiati, verde Tubi).
- **Guard-rail:** Messaggio di warning quando si supera il limite di selezione.

### E. Aggiorna Giacenza (Nuovo)
- **Card verticali 2-per-riga:** Layout touch-friendly con strip colore per tipologia (Forgiati blu, Tubi verde) e campi consultivi etichettati (ordine, fornitore, DN, ecc.).
- **Campi modificabili con salvataggio SharePoint:** PATCH via Graph su `/fields`:
  - Forgiati: Commessa (`field_24`), Data Prelievo (`field_23`), Giacenza Q.tà (`field_22`), Note (`field_25`), Giacenza mm da barra (`field_21`).
  - Tubi: Data ultimo prelievo (`field_21`), Commessa (`field_25`), Giacenza BIB tagliati (`field_20`), Giacenza mm contabile (`field_19`).
- **Messaggistica stato:** Alert success/error dopo il salvataggio; il carrello resta editabile finché non viene svuotato.

### C. Performance & Architettura
- **Caching Intelligente (`useCachedList`):**
  - I dati scaricati vengono salvati in RAM per 5 minuti.
  - Il cambio scheda (es. da Forgiati a Tubi) è istantaneo se i dati sono in cache.
  - Pulsante "Aggiorna" per forzare il refresh dai server Microsoft.
- **Download Totale:** Scaricamento batch di tutto l'inventario (4000+ righe) all'avvio per permettere filtri e ricerche locali istantanee.

## 3. Configurazione Tecnica

### Setup Variabili d'Ambiente (.env.local)
Assicurarsi che le seguenti chiavi siano configurate:
```bash
VITE_SHAREPOINT_SITE_ID=...
VITE_FORGIATI_LIST_ID=...
VITE_TUBI_LIST_ID=...
# VITE_ORING_HNBR_LIST_ID=... (Futuro)
# VITE_ORING_NBR_LIST_ID=... (Futuro)
```

### Struttura Colonne
Le configurazioni delle colonne sono statiche per performance e stabilità:
- `src/config/forgiatiColumns.ts`: Mapping per la lista Forgiati.
- `src/config/tubiColumns.ts`: Mapping per la lista Tubi (nomi interni come `field_1`, `field_2`...).

## 4. Prossimi Passi (Roadmap)

1. **Liste ORING:**
   - Recuperare GUID liste ORING.
   - Creare `src/config/oringColumns.ts`.
   - Duplicare/Adattare `TubiPanel` in un `OringPanel` generico riutilizzabile.

2. **Dettaglio & Modifica:**
   - Implementare modale/drawer al click sulla riga per vedere tutti i 28+ campi.
   - Aggiungere pulsanti per modifica rapida quantità (+/-) o prelievo.

3. **Checkout/integrazione carrello:**
  - Collegare il carrello a un Flow Power Automate o a una lista di ordini/prelievi dedicata.
  - Aggiungere stampa PDF/QR dell'elenco selezionato.

4. **Sezioni Extra Dashboard:**
  - **COMPLETATO** Implementare visualizzatore PDF per "Documentazione".
  - Collegare iframe o link esterno per "Sito Web".

## 5. Repository
Il progetto è ospitato su GitHub: `git@github.com:TEL-CO/totem-alfa.git`
Branch principale: `main`
