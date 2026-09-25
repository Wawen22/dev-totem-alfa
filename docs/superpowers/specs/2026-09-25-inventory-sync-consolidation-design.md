# Consolidamento sincronizzazione inventario — Design

**Stato:** approvato come indirizzo operativo il 25 settembre 2026; nessuna operazione sui dati è autorizzata senza conferma esplicita per il singolo magazzino.

## Obiettivo

Rendere affidabile la sincronizzazione fra i file Excel dei magazzini, le liste SharePoint e il Totem: nessuna modifica deve essere persa in caso di conflitto, duplicato, schema non compatibile o errore di rete.

Nella bonifica iniziale, Excel è la fonte autorevole. Dopo il completamento e la verifica di un magazzino, nessun lato prevale automaticamente: una differenza è un conflitto da mostrare e risolvere.

## Decisioni confermate

- Si procede un magazzino alla volta.
- L'utente ha eseguito un backup locale dei file Excel di magazzino.
- Prima di una scrittura, creazione o eliminazione SharePoint, Codex presenta il perimetro dell'operazione e chiede conferma.
- Una chiave logica è `CODICE + LOTTO` (il campo SharePoint `IdentLotto` corrisponde a `LOTTO`). Le chiavi mancanti, duplicate o ambigue non vengono scritte.
- I record presenti soltanto su SharePoint non vengono eliminati automaticamente durante la sincronizzazione ordinaria.
- La pulizia di duplicati identici o record vuoti richiede sempre una conferma separata e un elenco degli ID interessati.

### Eccezioni approvate: TUBO-MECCANICO, TUBI e FORGIATI

Il 25 settembre 2026 l'utente ha richiesto e confermato che il pulsante **Aggiorna Totem da Excel** per `4_TUBO-MECCANICO`, `3_TUBI` e `1_FORGIATI` ripristini il comportamento storico Excel-autorevole: con chiave `CODICE + LOTTO` univoca, i campi gestiti diversi vengono aggiornati in SharePoint dal file Excel. Il PATCH include soltanto le colonne realmente presenti nella tabella Excel, così non può cancellare campi assenti dal file. Restano invariati i blocchi per chiavi duplicate o ambigue, righe fuori tabella, schema non compatibile, errori consecutivi e l'assenza di eliminazioni automatiche. Questa eccezione non modifica la politica ordinaria degli altri magazzini.

## Modello operativo

### 1. Verifica non distruttiva

Per il magazzino selezionato, leggere file Excel e lista SharePoint e produrre un report con:

- nome file, tabella e intervallo effettivo Excel;
- colonne Excel e nomi interni SharePoint richiesti;
- conteggio righe valide su entrambi i lati;
- chiavi `CODICE + LOTTO` assenti, duplicate o ambigue;
- righe Excel valorizzate fuori dalla tabella;
- differenze per campo e record presenti su un solo lato.

Questa fase non modifica SharePoint né Excel.

### 2. Bonifica iniziale controllata (Excel autorevole)

Solo dopo una conferma specifica dell'utente, la lista SharePoint viene portata alla fotografia del file Excel verificato:

1. salvare nel report l'elenco degli item SharePoint prima dell'operazione e i relativi ID;
2. creare o aggiornare esclusivamente i record con una chiave Excel valida e non duplicata;
3. non eliminare al primo passaggio i record SharePoint non presenti in Excel: etichettarli come `da rivedere`;
4. proporre un secondo elenco di eventuali eliminazioni, limitato ai duplicati realmente identici e agli item completamente vuoti; l'esecuzione richiede una seconda conferma;
5. rileggere entrambe le fonti e produrre un report finale di convergenza.

Non è ammesso un "cancella tutta la lista e reimporta" senza il report di verifica e la conferma specifica: cancellerebbe ID, cronologia e informazioni non esportate nel file.

### 3. Sincronizzazione ordinaria protetta

Il pulsante **Aggiorna Totem da Excel** usa il seguente comportamento:

| Situazione | Azione |
| --- | --- |
| Chiave valida presente soltanto in Excel | Crea in SharePoint |
| Stessa chiave e stessi valori gestiti | Non modifica |
| Stessa chiave, differenze soltanto nell'identificativo tecnico del lotto | Allinea l'identificativo |
| Stessa chiave, almeno un valore aziendale diverso | Segnala conflitto; non sovrascrive |
| Chiave duplicata/ambigua o riga fuori tabella | Salta e segnala |
| Record presente soltanto in SharePoint | Segnala; non elimina |

Il pulsante **Aggiorna Excel da Totem** deve seguire la stessa politica: nessuna sovrascrittura silenziosa. Se un flusso non soddisfa questa regola, resta disabilitato o viene esplicitamente marcato come bonifica iniziale Excel-autorevole.

## Ordine di consolidamento

1. **TUBO MECCANICO** — primo intervento: già adotta i conflitti non distruttivi ma necessita della validazione preventiva dello schema SharePoint e della stessa gestione degli errori di TUBI/FORGIATI.
2. **TUBI** — verifica con casi di regressione e allineamento del reporting.
3. **FORGIATI** — verifica con casi di regressione e allineamento del reporting.
4. **FLANGE** — migrazione dal comportamento attuale Excel-prioritario al comportamento ordinario protetto; la bonifica iniziale resta Excel-prioritaria soltanto con conferma.
5. **ORING-HNBR, ORING-NBR, SPARK GUPS, FILO&FLUSSO** — definizione e attivazione graduale del flusso Excel -> Totem, ciascuno con la propria verifica e approvazione.

Un magazzino entra nello stato **Consolidato** solo dopo: build riuscita, test automatizzati, verifica non distruttiva pulita, bonifica iniziale approvata (se necessaria), rilettura post-operazione e test utente indicato nel report.

## Requisiti tecnici comuni

- Validare tutte le colonne SharePoint richieste prima della prima scrittura.
- Normalizzare testo, numeri e date prima di confrontare i campi.
- Registrare per ogni esecuzione: data, direzione, magazzino, esito, conteggi e dettaglio dei record saltati/conflittuali.
- Interrompere la sincronizzazione dopo errori consecutivi di scrittura, conservando il dettaglio delle righe già processate.
- Aggiornare la vista e invalidare la cache dopo una scrittura riuscita.
- Tutte le eliminazioni devono essere azioni separate, esplicitamente confermate e limitate agli item indicati nel report.

## Registro di avanzamento

| Data | Magazzino | Attività | Stato | Nota |
| --- | --- | --- | --- | --- |
| 2026-09-25 | FLANGE | Verifica iniziale live di file, tabella e schema SharePoint | completata | file `11_FLANGE .xlsx`, tabella `tblFlange`, intervallo `Foglio1!A1:W61`: 60 righe dati e 23 colonne; tutti i nomi interni SharePoint richiesti sono presenti. Nessun record è stato modificato. |
| 2026-09-25 | FLANGE | Protezioni da consolidare | da correggere | il flusso è già Excel-autorevole e testato, ma manca il preflight dello schema, lo stop dopo 3 errori consecutivi e il match deve essere limitato a `CODICE + LOTTO`, senza fallback sul solo `CODICE`. |
| 2026-09-25 | TUBI | Verifica iniziale di file, tabella e schema SharePoint | completata | file `3_TUBI.xlsx`, tabella `tblTUBI`, intervallo `MAG_TUBI_COOP_2012!A1:AB1116`: 1.115 righe dati e 28 colonne; tutti i nomi interni SharePoint richiesti sono presenti |
| 2026-09-25 | TUBI | Correzioni di robustezza | completato | `Prezzo metro` ora usa `field_23` (distinto da `Prezzo kg/mt` / `field_22`) sia nel flusso principale sia nel pannello admin; Excel -> SharePoint ora interrompe dopo 3 errori consecutivi. Test, TypeScript e build riusciti; nessun record è stato modificato. |
| 2026-09-25 | TUBI | Ripristino Excel-autorevole commit `bf3ed2e` | completato | Su chiave `CODICE + LOTTO` univoca, le differenze aziendali vengono aggiornate da Excel con PATCH limitato alle colonne presenti nel file; duplicati, ambiguità, righe fuori tabella e record solo SharePoint restano protetti. Test utente riuscito su `TCAB093` / `CODICE SAM`. |
| 2026-09-25 | FORGIATI | Verifica iniziale live di file, tabella e schema SharePoint | completata | file `1_FORGIATI.xlsx`, tabella `tblFORGIATI`, intervallo `'mag forg'!A1:AC4165`: 4.164 righe dati e 29 colonne; tutti i nomi interni SharePoint richiesti sono presenti. Nessun record è stato modificato. |
| 2026-09-25 | FORGIATI | Ripristino Excel-autorevole commit `57f9201` | completato | Test utente riuscito: 2 aggiornati e 4.131 invariati. Il report ha segnalato 12 duplicati Excel, 19 saltati e 200 record solo SharePoint / duplicati; nessuno è stato eliminato o modificato automaticamente. |
| 2026-09-25 | TUBO MECCANICO | Consolidamento tecnico commit `d09dff6` | completato | preflight schema, stop dopo 3 errori consecutivi e report record solo SharePoint aggiunti; test utente riuscito |
| 2026-09-25 | TUBO MECCANICO | Correzione schema SharePoint approvata | completato | lista Alfa `4_TUBO-MECCANICO`: creata colonna testo `IdentLotto`, nome visibile `LOTTO`; nessun record modificato |
| 2026-09-25 | TUBO MECCANICO | Ripristino Excel-autorevole commit `efcdfd8` | completato | differenze su chiave univoca aggiornate da Excel; PATCH limitato alle colonne realmente presenti; test utente riuscito |
| 2026-09-25 | Tutti | Bonifica iniziale SharePoint da Excel | non avviata | richiederà conferma per ciascun magazzino |

## Criteri di accettazione per ciascun magazzino

1. Una nuova riga Excel con chiave valida crea esattamente un item SharePoint.
2. Una riga identica non genera scritture.
3. Una modifica concorrente sullo stesso campo genera un conflitto senza alterare nessuno dei due valori.
4. Duplicati, chiavi mancanti e righe fuori tabella non generano scritture.
5. Il report espone tutti i record non riconciliati.
6. Il test utente finale dimostra una modifica del lato ammesso dal flusso, il salvataggio cloud e la convergenza della vista Totem dopo refresh.

## Limite noto e passo successivo

I pulsanti attuali sono sincronizzazioni manuali nel browser. Una sincronizzazione automatica anche a browser chiuso richiede un servizio persistente con storico/versioni (come descritto in `docs/SYNC_ARCHITECTURE.md`). Il presente lavoro consolida prima la sicurezza dei flussi manuali e della bonifica; l'automazione persistente è una fase successiva, separata.
