# Gestione ORING FKM ed EPDM

Integrazione nel TOTEM con la stessa gestione operativa di HNBR: ricerca, selezione, carrello, aggiornamento giacenze, ordini, prenotazioni e CRUD nel pannello Admin. Le liste sono indipendenti.

## SharePoint

Tenant: ALFA (alfaengnet.onmicrosoft.com). Sito: https://alfaengnet.sharepoint.com/sites/ALFA. Liste da creare: **2_ORING-FKM** e **2_ORING-EPDM**.

Fonte: TOTEM ALFA / MAGAZZINO / MAGAZZINI ATTIVI / 2_ORING.xlsx, fogli FKM ed EPDM. Importazione iniziale preparata il 22/10/2026: **26 record FKM** (righe 22–47) e **30 record EPDM** (righe 9–38). Le 14 posizioni FKM senza articolo non vengono importate.

Il codice FKM61400 compare in due posizioni diverse (23 e 26), conservate come record distinti. Gli ultimi due EPDM non hanno codice Alfa: SharePoint richiede Title, quindi si usa l’etichetta esplicita «Senza codice (cliente 76100)» / «Senza codice (cliente 76000)», senza inventare codici Alfa. La riga originale rimane identificata da RigaExcel. RigaExcel conserva la provenienza e permette di verificare ogni record senza deduplicare per Title.

Le celle senza intestazione R/T di EPDM vengono conservate come ValoreOriginaleR e NoteOriginali. Non sono interpretate come prenotazioni. Date Excel convertite in ISO, numeri conservati come numeri, celle vuote omesse senza trasformarle in zero. Nessuna modifica al file Excel sorgente.

Schema comune: Title (testo, colonna standard SharePoint per il codice Alfa), più:

| Nome interno | Etichetta | Tipo |
| --- | --- | --- |
| `field_0` | Posizione | text |
| `field_2` | Codice cliente | text |
| `field_3` | Colore Mescola | text |
| `field_4` | Ø Int. | number |
| `field_5` | Ø Est. | number |
| `field_6` | Ø Corda | number |
| `field_7` | Materiale | text |
| `field_8` | Durezza | text |
| `field_9` | Fluidi | text |
| `field_12` | Commessa / Standard | text |
| `field_13` | Design temperature | text |
| `field_14` | Prezzo unitario | number |
| `field_15` | Giacenza | number |
| `field_16` | Prenotazione fornitore | number |
| `field_17` | Data prelievo / aggiornamento | dateTime |
| `field_18` | NR Ordine | text |
| `field_19` | Data ordine | dateTime |
| `CodiceSAM` | Codice SAM | text |
| `QuantitaOrdinata` | Q.tà ordinata | number |
| `ValoreOriginaleR` | Valore originale colonna R | text |
| `NoteOriginali` | Note originali | text |
| `RigaExcel` | Riga Excel origine | number |

## Configurazione

GUID creati e già configurati in `.env.local` in questo checkout DEV:

| Variabile | GUID |
| --- | --- |
| `VITE_ORING_FKM_LIST_ID` | `6d32aab0-1068-40b9-8633-dc868a928538` |
| `VITE_ORING_EPDM_LIST_ID` | `98e87950-64df-490b-8547-ff8f0c0ad12f` |

La build Vercel configura questi due GUID pubblici tramite `vercel.json`, senza dipendere da `.env.local`. Le eventuali variabili già impostate sul progetto Vercel hanno precedenza. Per altri ambienti impostare queste variabili prima di ricostruire il frontend.

## Sincronizzazione

Le modifiche ad articoli FKM ed EPDM già presenti in Excel vengono salvate sia sulle liste SharePoint sia nei rispettivi fogli di `2_ORING.xlsx`, dal Pannello Admin e da Aggiorna Giacenza. Il percorso predefinito è `MAGAZZINO/MAGAZZINI ATTIVI/2_ORING.xlsx` nella libreria `TOTEM ALFA`; eventuali personalizzazioni usano `VITE_ORING_EXCEL_PATH`, `VITE_ORING_EXCEL_DRIVE_NAME` o `VITE_ORING_EXCEL_DRIVE_ID`.

Prima di scrivere vengono verificate intestazioni, identità originale del record e celle di destinazione. La scrittura riguarda solo le celle modificate, preservando formule, formati e modifiche Excel estranee all'operazione. I valori vengono riletti per verificare la scrittura. Se si modifica un diametro di cui dipende una formula, il risultato calcolato viene riportato anche su SharePoint. La modifica diretta di una cella calcolata viene rifiutata, indicando di modificare i valori da cui dipende.

I codici ripetuti vengono distinti con i dati originali e, quando necessario, la posizione e la riga di provenienza; un abbinamento ambiguo interrompe il salvataggio. Il codice originale viene usato anche in caso di rinomina. Le etichette “Senza codice” non vengono scritte come codici Alfa nel foglio.

La prenotazione EPDM viene salvata in una colonna `PRENOTAZIONE FORNITORE`, aggiunta dopo l'intervallo utilizzato alla prima valorizzazione. Le celle originali senza intestazione R/T restano distinte. Un errore Excel mostra uno stato di errore anche se SharePoint è già stato aggiornato; riprovare il salvataggio verifica e completa la scrittura. Salvare dall'Admin senza nuove modifiche permette di riallineare a Excel il record SharePoint dopo un precedente salvataggio parziale.

Questo flusso riguarda le modifiche dei record esistenti; la creazione e l'eliminazione dall'Admin rimangono operazioni sulle liste SharePoint. Non viene attivata una sincronizzazione automatica Excel → SharePoint. L'eventuale flusso Power Automate riceve anche le nuove sorgenti e i flag hasOringFkm/hasOringEpdm e non viene invocato se la sincronizzazione del carrello non è completata.

## Attivazione

Creazione e importazione eseguite il 22/10/2026, dopo conferma esplicita dell'utente:

- [2_ORING-FKM](https://alfaengnet.sharepoint.com/sites/ALFA/Lists/2_ORINGFKM): 26 record, giacenza totale 880.
- [2_ORING-EPDM](https://alfaengnet.sharepoint.com/sites/ALFA/Lists/2_ORINGEPDM): 30 record, giacenza totale 320.
- Tutte le 22 colonne personalizzate verificate per nome interno e tipo.
- Rilettura di ogni record da SharePoint e confronto di ogni campo importato con Excel: nessuna differenza.
- File Excel sorgente non modificato.

La verifica del frontend copre ricerca, selezione indipendente dei codici ripetuti, carrello misto FKM/EPDM, aggiornamento giacenze (incluso mantenimento del valore a video), instradamento sulle liste corrette, operazioni Admin di lettura/modifica/creazione/eliminazione e cache aggiornata. Le prove interattive sono eseguite in un browser con Graph simulato in memoria e rete esterna bloccata, senza scritture di prova sulle liste reali, a 1365×900 e 1080×1920. La verifica SharePoint usa invece le liste reali.

La build è generata in `dist/`; la pubblicazione del TOTEM è gestita dall'integrazione Git di Vercel dopo il push.

Verifica della correzione Excel: 68 test automatici superati, inclusi 17 test su mapping, aggiornamento puntuale, formule, date, celle vuote, righe spostate, duplicati, errori parziali e retry. Prove interattive con Graph simulato: modifiche di colore e giacenza in entrambi i materiali, verifica delle celle rilette e conservazione di tutte le formule originali, errore Excel visibile e retry riuscito. Non sono state eseguite scritture di prova sui dati reali del cliente.
