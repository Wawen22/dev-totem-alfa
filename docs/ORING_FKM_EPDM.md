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

Come HNBR e NBR, questi magazzini salvano le modifiche sulle liste SharePoint. Non è presente una sincronizzazione automatica bidirezionale con Excel per gli ORING esistenti; questa integrazione non la attiva. L'importazione iniziale è separata dai successivi aggiornamenti operativi. L'eventuale flusso Power Automate riceve anche le nuove sorgenti e i flag hasOringFkm/hasOringEpdm.

## Attivazione

Creazione e importazione eseguite il 22/10/2026, dopo conferma esplicita dell'utente:

- [2_ORING-FKM](https://alfaengnet.sharepoint.com/sites/ALFA/Lists/2_ORINGFKM): 26 record, giacenza totale 880.
- [2_ORING-EPDM](https://alfaengnet.sharepoint.com/sites/ALFA/Lists/2_ORINGEPDM): 30 record, giacenza totale 320.
- Tutte le 22 colonne personalizzate verificate per nome interno e tipo.
- Rilettura di ogni record da SharePoint e confronto di ogni campo importato con Excel: nessuna differenza.
- File Excel sorgente non modificato.

La verifica del frontend copre ricerca, selezione indipendente dei codici ripetuti, carrello misto FKM/EPDM, aggiornamento giacenze (incluso mantenimento del valore a video), instradamento sulle liste corrette, operazioni Admin di lettura/modifica/creazione/eliminazione e cache aggiornata. Le prove interattive sono eseguite in un browser con Graph simulato in memoria e rete esterna bloccata, senza scritture di prova sulle liste reali, a 1365×900 e 1080×1920. La verifica SharePoint usa invece le liste reali.

La build è generata in `dist/`; la pubblicazione del TOTEM è gestita dall'integrazione Git di Vercel dopo il push.

Verifica finale: 51 test automatici superati, controllo TypeScript e build Vite riusciti.
