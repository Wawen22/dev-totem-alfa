# Consolidamento TUBO MECCANICO Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rendere il flusso manuale **Aggiorna Totem da Excel** di TUBO MECCANICO non distruttivo e verificabile allo stesso livello di TUBI e FORGIATI, quindi produrre una verifica iniziale senza modificare dati.

**Architecture:** Il flusso continuerà a leggere Excel e SharePoint in browser tramite Microsoft Graph, ma prima di ogni scrittura verificherà lo schema della lista. Un piccolo modulo puro conterrà i controlli condivisi e sarà coperto da test Vitest; `App.tsx` applicherà il preflight, preparerà i payload in base al tipo di colonna e fermerà il run dopo tre errori consecutivi. La successiva verifica del magazzino sarà read-only via SharePoint MCP e terminerà con una richiesta di conferma prima della bonifica Excel-autorevole.

**Tech Stack:** React 19, TypeScript 5.9, Vite 7, Vitest 4, Microsoft Graph, SharePoint.

**Spec:** `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md`

## Global Constraints

- Excel è autorevole esclusivamente nella bonifica iniziale approvata esplicitamente dall'utente.
- Nel sync ordinario, una differenza di valore aziendale è un conflitto e non autorizza alcuna sovrascrittura.
- La chiave di riconciliazione è `CODICE + LOTTO` (`Title + IdentLotto`).
- Duplicati, chiavi ambigue, righe fuori dalla tabella e record presenti soltanto su SharePoint non generano eliminazioni automatiche.
- Nessuna chiamata MCP di creazione, aggiornamento o eliminazione SharePoint è eseguita in questo piano senza una conferma successiva e specifica dell'utente.
- Il registro `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md` va aggiornato al termine di ogni attività completata.

## Review Focus

- Lista SharePoint con un campo interno rinominato o mancante: il sync deve fallire prima della prima scrittura; coperto in Task 2 e Task 3.
- Tre rifiuti Graph consecutivi: il run deve terminare e mantenere il dettaglio dell'ultima riga fallita; coperto in Task 2 e Task 3.
- Un errore isolato fra due operazioni riuscite: il contatore deve azzerarsi e il run può continuare; coperto in Task 2 e Task 3.
- Record presente solo SharePoint: deve comparire nel report e non essere eliminato; coperto in Task 3 e nel test utente di Task 4.
- Date, numeri e valori nulli: in creazione SharePoint devono essere inviati soltanto i campi ammessi dal loro tipo; coperto nel test utente di Task 4.

---

### Task 1: Aggiungere l'infrastruttura di test TypeScript

**Files:**
- Modify: `package.json`
- Modify: `package-lock.json`
- Create: `vitest.config.ts`

**Interfaces:**
- Produces: script `npm test` per eseguire i file `src/**/*.test.ts` in ambiente Node.
- Produces: script `npm run test:run` per la CI e per la verifica non interattiva.

- [ ] **Step 1: Installare la dipendenza di test compatibile con l'ambiente corrente**

Run:

```bash
npm install --save-dev vitest@^4.1.6
```

Node è `v22.17.0` e il progetto usa Vite `7.2.x`, quindi soddisfa i prerequisiti della versione Vitest 4 scelta.

- [ ] **Step 2: Aggiungere gli script di test in `package.json`**

Inserire nella sezione `scripts`:

```json
"test": "vitest",
"test:run": "vitest run"
```

- [ ] **Step 3: Creare la configurazione Node per i test puri**

```ts
// vitest.config.ts
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
    passWithNoTests: true,
  },
});
```

- [ ] **Step 4: Verificare che Vitest individui la configurazione**

Run: `npm run test:run`

Expected: termina con codice `0`; prima del Task 2 è accettabile il messaggio che non sono stati trovati test.

- [ ] **Step 5: Commit**

```bash
git add package.json package-lock.json vitest.config.ts
git commit -m "test: aggiunge Vitest per i controlli sync"
```

### Task 2: Estrarre e testare le guardie di sincronizzazione

**Files:**
- Create: `src/services/syncGuards.ts`
- Create: `src/services/syncGuards.test.ts`

**Interfaces:**
- Produces: `getMissingRequiredColumns(existingColumnNames: Iterable<string>, requiredColumnNames: readonly string[]): string[]`.
- Produces: `recordWriteOutcome(previousConsecutiveFailures: number, succeeded: boolean): number`.
- Produces: `MAX_CONSECUTIVE_SYNC_WRITE_FAILURES`, valore `3`.
- Consumed by: `handleSyncTuboMeccanicoExcelToSharePoint` in `src/App.tsx`.

- [ ] **Step 1: Scrivere i test che descrivono i controlli**

```ts
import { describe, expect, it } from "vitest";
import {
  MAX_CONSECUTIVE_SYNC_WRITE_FAILURES,
  getMissingRequiredColumns,
  recordWriteOutcome,
} from "./syncGuards";

describe("getMissingRequiredColumns", () => {
  it("restituisce soltanto i nomi interni assenti", () => {
    expect(getMissingRequiredColumns(["Title", "IdentLotto", "field_1"], ["Title", "IdentLotto", "field_1", "field_3"]))
      .toEqual(["field_3"]);
  });

  it("elimina i requisiti ripetuti", () => {
    expect(getMissingRequiredColumns(["Title"], ["Title", "Title"]))
      .toEqual([]);
  });

  it("espone sia il campo tecnico sia il campo data mancanti", () => {
    expect(getMissingRequiredColumns(["Title"], ["Title", "IdentLotto", "field_3"]))
      .toEqual(["IdentLotto", "field_3"]);
  });
});

describe("recordWriteOutcome", () => {
  it("azzera gli errori dopo una scrittura riuscita", () => {
    expect(recordWriteOutcome(2, true)).toBe(0);
  });

  it("raggiunge la soglia di arresto al terzo errore consecutivo", () => {
    const failures = recordWriteOutcome(recordWriteOutcome(1, false), false);
    expect(failures).toBe(MAX_CONSECUTIVE_SYNC_WRITE_FAILURES);
  });
});
```

- [ ] **Step 2: Eseguire il test e verificare il fallimento iniziale**

Run: `npm run test:run -- src/services/syncGuards.test.ts`

Expected: FAIL perché il modulo `./syncGuards` non esiste.

- [ ] **Step 3: Implementare il modulo puro**

```ts
export const MAX_CONSECUTIVE_SYNC_WRITE_FAILURES = 3;

export function getMissingRequiredColumns(
  existingColumnNames: Iterable<string>,
  requiredColumnNames: readonly string[]
): string[] {
  const existing = new Set(existingColumnNames);
  return Array.from(new Set(requiredColumnNames)).filter((name) => !existing.has(name));
}

export function recordWriteOutcome(
  previousConsecutiveFailures: number,
  succeeded: boolean
): number {
  return succeeded ? 0 : previousConsecutiveFailures + 1;
}
```

- [ ] **Step 4: Eseguire i test**

Run: `npm run test:run -- src/services/syncGuards.test.ts`

Expected: 5 test PASS.

- [ ] **Step 5: Commit**

```bash
git add src/services/syncGuards.ts src/services/syncGuards.test.ts
git commit -m "test: copre guardie per sync inventario"
```

### Task 3: Applicare le guardie a TUBO MECCANICO

**Files:**
- Modify: `src/App.tsx:1017-1038,7963-8248`

**Interfaces:**
- Consumes: `SharePointService.listColumns(listId)`, `prepareFieldsForSharePointCreate(fields, columns)`, `getMissingRequiredColumns`, `recordWriteOutcome`, `MAX_CONSECUTIVE_SYNC_WRITE_FAILURES`.
- Produces: un `SyncResult` che contiene sezioni `updated`, `created`, `unchanged`, `skipped`, `duplicates` e `sharepoint-only`.

- [ ] **Step 1: Aggiungere il preflight prima dell'indicizzazione SharePoint**

Subito dopo `listItems` nel gestore TUBO MECCANICO, ottenere le colonne e rifiutare il run se manca anche un solo nome interno richiesto:

```ts
const [spItems, sharePointColumns] = await Promise.all([
  sharepointService.listItems<Record<string, unknown>>(tuboMeccanicoListId),
  sharepointService.listColumns(tuboMeccanicoListId),
]);
const requiredSyncColumns = [
  "Title",
  "IdentLotto",
  ...TUBO_MECCANICO_SHAREPOINT_TEXT_FIELDS,
  ...TUBO_MECCANICO_SHAREPOINT_NUMERIC_FIELDS,
  ...TUBO_MECCANICO_SHAREPOINT_DATE_FIELDS,
];
const missingSyncColumns = getMissingRequiredColumns(
  sharePointColumns.map((column) => column.name),
  requiredSyncColumns
);
if (missingSyncColumns.length > 0) {
  throw new Error(
    `Colonne SharePoint TUBO-MECCANICO mancanti o con nome interno diverso: ${missingSyncColumns.join(", ")}`
  );
}
```

- [ ] **Step 2: Preparare i campi prima di ogni creazione SharePoint**

Sostituire la creazione diretta di `record.fields` con:

```ts
const createFields = prepareFieldsForSharePointCreate(record.fields, sharePointColumns);
const createdItem = await sharepointService.createItem<Record<string, unknown>>(
  tuboMeccanicoListId,
  createFields
);
```

Usare `createFields` anche come fallback per costruire la mappa comparabile dell'item appena creato.

- [ ] **Step 3: Fermare la scrittura al terzo errore consecutivo**

Inizializzare `let consecutiveWriteErrors = 0` prima del ciclo. Dopo ogni iterazione senza eccezione assegnare:

```ts
consecutiveWriteErrors = recordWriteOutcome(consecutiveWriteErrors, true);
```

Nel `catch`, dopo aver aggiunto il record a `skippedLabels`, assegnare:

```ts
consecutiveWriteErrors = recordWriteOutcome(consecutiveWriteErrors, false);
if (consecutiveWriteErrors >= MAX_CONSECUTIVE_SYNC_WRITE_FAILURES) {
  throw new Error(
    `Sincronizzazione interrotta dopo ${MAX_CONSECUTIVE_SYNC_WRITE_FAILURES} errori SharePoint consecutivi. Ultimo record: ${record.title} (${record.identLotto}). ${message}`
  );
}
```

- [ ] **Step 4: Segnalare senza eliminare i record solo SharePoint**

Dopo il ciclo, costruire `sharePointOnlyLabels` dagli item non presenti in `usedItemIds`, usando `getResolvedTuboMeccanicoIdentLotto`. Aggiungere la sezione `sharepoint-only` ai dettagli e includerne il conteggio nel messaggio. `success` deve essere `false` se esistono tali record, perché richiedono verifica manuale, ma non deve chiamare `deleteItem`.

- [ ] **Step 5: Eseguire test, typecheck e build**

Run:

```bash
npm run test:run
npx tsc -b
npm run build
```

Expected: tutti con exit code `0`.

- [ ] **Step 6: Commit**

```bash
git add src/App.tsx src/services/syncGuards.ts src/services/syncGuards.test.ts
git commit -m "fix: consolida sync TUBO MECCANICO"
```

### Task 4: Verifica utente dopo il deploy, senza bonifica dati

**Files:**
- Modify: `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md`

**Interfaces:**
- Consumes: pulsante Admin **Aggiorna Totem da Excel** per TUBO MECCANICO.
- Produces: esito utente documentato nel registro di avanzamento.

- [ ] **Step 1: Deploy della build che contiene il commit di Task 3**

Attendere che il deployment Vercel associato al commit sia `Ready`; non eseguire il sync da Excel durante la build o su una versione precedente.

- [ ] **Step 2: Eseguire il test utente su una riga di prova non ambigua**

1. Aprire il file TUBO MECCANICO in Excel desktop/web.
2. Modificare un campo non chiave di una riga di prova e salvare; attendere che il salvataggio SharePoint/OneDrive sia concluso.
3. In Totem, aprire **Pannello Admin → 4_TUBO-MECCANICO → Aggiorna Totem da Excel**.
4. Verificare che il report mostri `Conflitto Excel / Totem` e che il valore Totem non venga sovrascritto.
5. Ripristinare il valore Excel al valore Totem e salvarlo; rieseguire il pulsante e verificare `Nessuna differenza rilevata`.

Questo test non richiede modifiche SharePoint manuali: il primo run deve essere soltanto conflittuale e il secondo invariato.

- [ ] **Step 3: Registrare l'esito senza inserire dati sensibili**

Aggiornare la riga TUBO MECCANICO nel registro del documento di specifica con: data, commit deployato, esito dei test e valore `consolidato` o `da correggere`.

- [ ] **Step 4: Commit del registro**

```bash
git add docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md
git commit -m "docs: registra verifica sync TUBO MECCANICO"
```

### Task 5: Audit iniziale TUBO MECCANICO, solo lettura

**Files:**
- Modify: `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md`

**Interfaces:**
- Consumes: connessione MCP esplicita del tenant Alfa, file Excel TUBO MECCANICO, tabella Excel configurata e lista SharePoint configurata.
- Produces: report di preflight con conteggi, schema, chiavi duplicate/ambigue, record solo Excel, record solo SharePoint e conflitti.

- [ ] **Step 1: Collegare e selezionare esplicitamente il profilo Alfa**

Usare il MCP locale SharePoint per elencare i profili e selezionare il tenant Alfa prima di ogni lettura. Non usare un tenant implicito.

- [ ] **Step 2: Scoprire le risorse senza modificarle**

Leggere: libreria, percorso e tabella Excel TUBO MECCANICO; colonne e item della lista SharePoint TUBO MECCANICO. Verificare che corrispondano alle variabili `VITE_TUBO_MECCANICO_*` senza stampare segreti.

- [ ] **Step 3: Produrre il riepilogo da presentare all'utente**

Il riepilogo deve contenere esattamente:

```text
Excel: <righe valide>, <duplicati>, <righe fuori tabella>
SharePoint: <item>, <duplicati>, <record solo SharePoint>
Confronto: <chiavi coincidenti>, <solo Excel>, <conflitti>
Azioni proposte: <creazioni>, <aggiornamenti Excel-autorevoli>, <nessuna eliminazione>, <elementi da revisione>
```

Non chiamare `createItem`, `updateItem` o `deleteItem`.

- [ ] **Step 4: Aggiornare il registro e chiedere la conferma operativa**

Aggiungere i conteggi al registro, poi chiedere all'utente una conferma che indichi: nome magazzino, numero di creazioni, numero di aggiornamenti e che non verrà eseguita alcuna eliminazione. Solo questa conferma abiliterà una futura bonifica iniziale.

- [ ] **Step 5: Commit del registro audit**

```bash
git add docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md
git commit -m "docs: registra audit TUBO MECCANICO"
```

## Self-review

- **Copertura spec:** preflight, protezione conflitti, chiavi, errori, record solo SharePoint, test e audit read-only sono assegnati ai Task 1–5. La bonifica che scrive dati è intenzionalmente fuori piano, perché richiede il report e una conferma futura.
- **Nessun placeholder:** ogni task indica file, interfacce, comandi e comportamento atteso.
- **Coerenza tipi:** `getMissingRequiredColumns`, `recordWriteOutcome` e `MAX_CONSECUTIVE_SYNC_WRITE_FAILURES` sono definiti in Task 2 e consumati in Task 3.
- **Review focus:** ogni voce è collegata al test di Task 2/3 oppure alla verifica esplicita di Task 4.
