# Sincronizzazione inventario Excel / Totem

## Problema osservato

Il Totem legge le liste SharePoint. Excel e le liste sono modificabili separatamente.
I pulsanti manuali hanno finora confrontato soltanto i valori correnti e hanno
scritto righe complete sulla destinazione. Quando entrambi i lati cambiano,
questa operazione può cancellare una modifica valida. La chiave di lotto
`CODICE + IdentLotto` non ha inoltre un vincolo di unicità condiviso.

## Protezione immediata nel client

I sync manuali per FORGIATI e TUBI aggiungono i record assenti, ma segnalano
le righe esistenti con valori divergenti senza sovrascriverle. Il sync
Excel -> SharePoint di TUBO-MECCANICO applica la stessa regola. Una chiave
duplicata o ambigua viene saltata e mostrata nel dettaglio della
sincronizzazione. Il vecchio sync completo SharePoint -> Excel di
TUBO-MECCANICO è sospeso, perché scriveva l'intera riga senza confronto.

Questa protezione evita alcune perdite di dati, ma non è la sincronizzazione
automatica definitiva: le modifiche a righe esistenti richiedono una decisione
esplicita finché non esiste uno storico condiviso. I salvataggi diretti da
Totem e da Admin, e i sync generici di altre liste, vanno migrati allo stesso
motore prima di poter garantire la convergenza completa.

## Architettura definitiva richiesta

1. Ogni riga Excel e ogni item SharePoint ricevono un `SyncId` immutabile.
   Durante la migrazione si associa per `CODICE + IdentLotto` solo quando il
   match è univoco; i duplicati vanno in una coda di risoluzione. Il `SyncId`
   deve essere univoco sul lato SharePoint. Il codice e la lettera lotto
   rimangono dati modificabili, non identificatori tecnici.
2. Un servizio di sincronizzazione, in esecuzione anche quando il browser è
   chiuso, conserva per ogni `SyncId` lo snapshot dell'ultima sincronizzazione
   e la versione di entrambi i lati. Un trigger sulla modifica del file Excel
   e un trigger sulla lista SharePoint mettono il `SyncId` in una coda. Una
   scansione periodica recupera eventuali eventi persi.
3. Per ogni campo si confrontano `base`, Excel corrente e SharePoint corrente.
   Se cambia un solo lato, si propaga quel campo. Se cambiano campi diversi,
   si uniscono. Se lo stesso campo cambia su entrambi i lati, si registra un
   conflitto visibile; nessun valore viene sostituito automaticamente.
4. Gli aggiornamenti SharePoint usano la versione/ETag letta in precedenza;
   se cambia nel frattempo, l'operazione riparte da una nuova lettura. Per
   Excel, un unico worker per file limita scritture concorrenti, rilegge la
   riga prima di scriverla e ripete il confronto. Le operazioni devono essere
   idempotenti, con retry e log dei fallimenti.
5. Il Totem mostra stato della sincronizzazione e conflitti. Le azioni
   manuali diventano "Verifica ora" e "Risolvi conflitto", non importazioni
   che scelgono implicitamente un lato vincente.

## Condizioni per il rilascio

- Fare backup di liste e file Excel e inventariare le colonne effettive.
- Eseguire una migrazione iniziale con report di chiavi duplicate, lotti
  mancanti e righe Excel fuori tabella. Risolvere queste anomalie prima di
  abilitare la sincronizzazione automatica.
- Provare: modifica solo Excel, solo Totem, campi diversi, stesso campo,
  nuovo lotto da entrambi i lati, retry dopo errore, due operatori concorrenti,
  trigger duplicato e browser chiuso.
- Misurare il ritardo tra modifica e convergenza; la sincronizzazione sarà
  eventuale, non istantanea durante errori di rete o conflitti reali.
