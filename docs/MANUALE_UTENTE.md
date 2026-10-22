# 📖 Manuale Utente - App Totem ALFA ENG

Guida all'utilizzo dell'applicazione Totem Magazzino per la gestione dell'inventario.

---

## 🔑 Accesso all'Applicazione

L'accesso è integrato con l'ecosistema Microsoft 365 aziendale:

- **Link Applicazione:** Totem ALFA ENG
- **Modalità:** Effettuare il login utilizzando il proprio account Microsoft aziendale

> 💡 **Nota:** L'autenticazione avviene tramite Microsoft 365. Utilizzare le stesse credenziali dell'account aziendale.

---

## 🏠 Dashboard Principale

Dopo l'accesso, viene visualizzata la **Dashboard** con tre opzioni principali:

| Opzione | Descrizione |
|---------|-------------|
| **Giacenza Magazzino** | Consulta l'inventario e gli ordini in tempo reale |
| **Consulta Documentazione** | Manuali, procedure e schede tecniche |
| **Visualizza Sito Web** | Accesso rapido al portale aziendale (www.alfa-eng.net) |

---

## 📦 Giacenza Magazzino

La sezione **Giacenza Magazzino** permette di consultare e aggiornare l'inventario. L'inventario è organizzato in **5 categorie**:

### Categorie Inventario

| Tab | Descrizione |
|-----|-------------|
| **1_FORGIATI** | Articoli forgiati |
| **2_ORING-HNBR** | O-ring in materiale HNBR |
| **2_ORING-NBR** | O-ring in materiale NBR |
| **2_ORING-FKM** | O-ring in materiale FKM |
| **2_ORING-EPDM** | O-ring in materiale EPDM |
| **3_TUBI** | Tubi e tubazioni |
| **6_SPARK GUPS** | Spark Gups |

### Funzionalità di Ricerca

- 🔍 **Barra di ricerca:** Cerca in tutte le colonne contemporaneamente
- 📄 **Paginazione:** Visualizza 100 o 200 elementi per pagina
- 🔄 **Pulsante "Aggiorna":** Ricarica i dati dall'inventario

### Selezione Articoli

1. Clicca sulla **checkbox** a sinistra della riga per selezionare un articolo
2. Per articoli con più lotti (FORGIATI e TUBI), clicca sull'icona 📦 per gestire i lotti
3. Massimo **10 articoli** selezionabili contemporaneamente

---

## 🛒 Carrello e Aggiornamento Giacenza

### Visualizzazione Carrello

Gli articoli selezionati appaiono nel **Carrello** in basso nella schermata:

- Visualizza il numero di articoli selezionati (es. "3 / 10 articoli selezionati")
- Possibilità di **comprimere/espandere** il carrello
- **Svuota articoli** per rimuovere tutti gli elementi
- **Rimuovi** singoli articoli con il pulsante ✕

### Aggiornamento Giacenza

1. Seleziona gli articoli desiderati dal magazzino
2. Clicca su **"Aggiorna Giacenza"**
3. Modifica i campi disponibili per ciascun articolo
4. Clicca **"Aggiorna"** per salvare le modifiche

---

## 📝 Campi Modificabili per Tipologia

### FORGIATI

| Campo | Descrizione |
|-------|-------------|
| Commessa | Numero commessa di riferimento |
| Data prelievo | Data del prelievo |
| Giacenza Q.tà | Quantità in giacenza |
| Giacenza mm da barra | Lunghezza in mm da barra |
| Note | Annotazioni libere |

### TUBI

| Campo | Descrizione |
|-------|-------------|
| Data ultimo prelievo | Data dell'ultimo prelievo |
| Commessa | Numero commessa |
| Giacenza Tubo Intero mm | Lunghezza tubo intero |
| Giacenza Contab mm | Giacenza contabile (accetta somme, es. "100+20+10") |

### ORING-HNBR

| Campo | Descrizione |
|-------|-------------|
| Commessa | Numero commessa |
| Giacenza | Quantità in giacenza |
| Prenotazione | Quantità prenotata |
| Data prelievo | Data del prelievo |
| Nr Ordine | Numero ordine |
| Data ordine | Data dell'ordine |

### ORING-NBR

| Campo | Descrizione |
|-------|-------------|
| Commessa | Numero commessa |
| Q.tà minima | Quantità minima |
| Giacenza | Quantità in giacenza |
| Prenotazione | Quantità prenotata |
| Data prelievo | Data del prelievo |

### SPARK-GUPS

| Campo | Descrizione |
|-------|-------------|
| Lotto | Identificativo lotto |
| Codice SAM | Codice SAM |
| Tipologia Articolo | Tipologia dell'articolo |
| N. Ordine | Numero ordine |
| Data Ordine | Data dell'ordine |
| Fornitore | Nome fornitore |
| Quantità Ordinata | Quantità ordinata |
| N. Bolla | Numero bolla |
| Data Consegna | Data consegna |
| Giacenza | Quantità in giacenza |
| Data Ultimo Prelievo | Data ultimo prelievo |
| Prezzo unitario | Prezzo unitario |
| Commessa | Numero commessa |

---

## 📦 Gestione Lotti (FORGIATI e TUBI)

Per gli articoli FORGIATI e TUBI, è possibile gestire i lotti associati:

### Visualizzare i Lotti

1. Clicca sull'icona 📦 accanto al codice articolo
2. Si apre una finestra con tutti i lotti disponibili per quell'articolo

### Selezionare un Lotto

- Ogni lotto è identificato da una **lettera progressiva** (A, B, C...)
- Clicca **"Aggiungi"** per selezionare il lotto desiderato

### Creare un Nuovo Lotto

1. Dalla finestra di gestione lotti, compila i campi:
   - **Nuovo N° COLATA** (obbligatorio - deve essere diverso dai lotti esistenti)
   - **Nuovo N° ORDINE** (facoltativo)
   - **Codice SAM** (facoltativo)
   - **Data ordine** (facoltativo)
2. Clicca **"Crea lotto"**

> ⚠️ **Attenzione:** Il sistema genera automaticamente un identificativo lotto progressivo (A, B, C...). Il nuovo lotto viene creato duplicando l'articolo esistente con i nuovi dati inseriti.

---

## 🛠️ Pannello Admin e Gestione Inventario

Gli utenti autorizzati visualizzeranno un pulsante **"Pannello Admin"** in basso a destra nella schermata principale. Questo tasto abilita le funzioni di gestione inventario avanzate.

### 👥 Utenti Abilitati al Pannello Admin

L'accesso al Pannello Admin è riservato ai seguenti profili:

- Cataldo Ruppi
- Hamza Errais
- Laura Erpici
- Maria Rosa Pedrazzi
- Marius Bejinaru
- Mirco Masini
- Simone Borghi

### 📋 Funzionalità Disponibili

All'interno della sezione **Gestione Inventario**, è possibile operare direttamente sul SharePoint senza dover intervenire sui file Excel:

#### 1. Modifica Articoli

- Seleziona un articolo dalla lista
- Modifica i campi desiderati nel form di modifica
- Clicca **"Salva modifiche"**
- Le modifiche vengono sincronizzate automaticamente su SharePoint e Excel

#### 2. Inserimento Nuovi Articoli

- Clicca su **"+ Nuovo articolo"**
- Compila i campi del form (il campo **Codice/Title** è obbligatorio)
- Clicca **"Crea nuovo"**
- Il nuovo articolo viene aggiunto a SharePoint e Excel

#### 3. Eliminazione Articoli

- Clicca sull'icona 🗑️ accanto all'articolo da eliminare
- Conferma l'eliminazione
- L'articolo viene rimosso da SharePoint e Excel

### Vantaggi del Pannello Admin

Questa sezione di **GESTIONE INVENTARIO** nel Pannello Admin mira a:

- ✅ Centralizzare l'operatività in un'unica interfaccia
- ✅ Eliminare la necessità di intervenire direttamente sui file Excel
- ✅ Garantire l'unicità e la consistenza dei dati
- ✅ Sincronizzare automaticamente le modifiche tra SharePoint e Excel

---

## 📁 Consulta Documentazione

La sezione **Consulta Documentazione** permette di:

- Navigare nelle cartelle documentali
- Visualizzare PDF e documenti senza uscire dall'applicazione
- Consultare manuali, procedure e schede tecniche

### Come Utilizzare

1. Dalla Dashboard, clicca su **"Consulta Documentazione"**
2. Naviga nelle cartelle cliccando sui nomi
3. Clicca su un file per visualizzarlo
4. Usa il pulsante **"← Indietro"** per tornare alla cartella precedente

---

## 🌐 Visualizza Sito Web

La sezione **Visualizza Sito Web** permette di accedere al portale aziendale:

- Visualizza il sito **www.alfa-eng.net** direttamente nell'app
- Non è necessario aprire un browser esterno
- Usa il pulsante **"← Indietro"** per tornare alla Dashboard

---

## ❓ Domande Frequenti

### Come faccio a selezionare più articoli?

Clicca sulla checkbox a sinistra di ogni riga. Puoi selezionare fino a 10 articoli contemporaneamente.

### Come creo un nuovo lotto per un articolo esistente?

Per FORGIATI e TUBI: clicca sull'icona 📦, poi compila i campi per il nuovo lotto e clicca "Crea lotto".

### Perché non vedo il pulsante "Pannello Admin"?

Il Pannello Admin è visibile solo agli utenti autorizzati. Se hai bisogno di accesso, contatta uno degli amministratori.

### Le modifiche vengono salvate automaticamente?

No, le modifiche devono essere confermate cliccando su "Aggiorna" (per la giacenza) o "Salva modifiche" (nel Pannello Admin).

### Come torno alla Dashboard?

Clicca sul pulsante **"← Indietro"** o sul logo ALFA ENG in alto a sinistra.

---

## 📞 Supporto

Per assistenza tecnica o problemi con l'applicazione, contattare:

- **Team IT** tramite i canali aziendali
- **Amministratori del sistema** (vedi lista utenti abilitati)

---

*Ultimo aggiornamento: Febbraio 2026*
*Versione documento: 1.0*

### ORING-FKM e ORING-EPDM

I nuovi magazzini consentono ricerca, selezione nel carrello, modifica di giacenza, prenotazione, commessa, numero e data ordine e data prelievo, come HNBR. Il Pannello Admin consente di inserire, modificare ed eliminare articoli in ciascuna lista. Posizione e riga Excel distinguono i record con lo stesso codice. Gli EPDM senza codice Alfa sono indicati con “Senza codice (cliente …)”; l’amministratore può completarne il codice.

### Pulsanti di sincronizzazione ORING

Per HNBR, NBR, FKM ed EPDM, seleziona il magazzino nel Pannello Admin e usa:

- **Aggiorna Totem da Excel** per riportare nella lista i valori del foglio, inclusi gli articoli nuovi.
- **Aggiorna Excel da Totem** per riportare nel foglio i valori della lista, aggiungendo gli articoli mancanti e conservando le formule.

Attendi il risultato prima di eseguire altre operazioni. Il riepilogo mostra aggiornati, creati, invariati e casi da verificare; il pulsante del dettaglio permette di consultarli. Gli elementi presenti solo nella destinazione non vengono eliminati. Se un codice ha un abbinamento ambiguo, la relativa riga non viene modificata e il risultato viene indicato come parziale.
