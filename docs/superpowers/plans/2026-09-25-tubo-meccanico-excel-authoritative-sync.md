# TUBO MECCANICO Excel-autorevole Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fare sì che `Aggiorna Totem da Excel` aggiorni i campi diversi della lista `4_TUBO-MECCANICO` dal file Excel, come accadeva prima del commit di protezione `0e5e05d`.

**Architecture:** La chiave resta `CODICE + LOTTO`. Per una chiave univoca già presente, una differenza nei campi gestiti genera un `PATCH` con i valori dell'Excel e viene riportata come aggiornamento; per chiavi duplicate, ambigue, righe fuori tabella e item solo SharePoint il comportamento protetto resta invariato. Il comportamento viene incapsulato in una guardia pura testata prima dell'uso nel gestore React.

**Tech Stack:** React 19, TypeScript 5.9, Vite 7, Vitest 4, Microsoft Graph.

**Spec:** `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md`

## Global Constraints

- Ambito: solo `TUBO-MECCANICO`, direzione Excel -> SharePoint.
- L'utente ha confermato esplicitamente il 25 settembre 2026 che per questo magazzino Excel è autorevole.
- Nessuna eliminazione automatica; record solo SharePoint restano nel report.
- Chiavi duplicate/ambigue e righe fuori dalla tabella non vengono scritte.
- Con schema incompleto il run fallisce prima della prima scrittura; dopo tre errori consecutivi si interrompe.
- Il commit viene pubblicato su `origin/master` come richiesto dall'utente.

## Review Focus

- Differenza solo in `CODICE SAM`: deve produrre un aggiornamento, non un conflitto saltato.
- Riga identica: non deve generare scritture.
- Solo `LOTTO` mancante: deve continuare ad aggiornare solo l'identificativo.
- Duplicati Excel o SharePoint: nessun `PATCH` né creazione.
- Un valore Excel vuoto deve restare incluso nel payload autoritativo per poter svuotare il campo SharePoint corrispondente.

### Task 1: Rendere autoritativo il sync di TUBO-MECCANICO

**Files:**
- Modify: `src/services/syncGuards.ts`
- Modify: `src/services/syncGuards.test.ts`
- Modify: `src/App.tsx:8172-8226`
- Modify: `docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md`

**Interfaces:**
- Produces: `shouldApplyExcelAuthoritativeUpdate(changedBusinessFields: number, identityNeedsAlignment: boolean): boolean` e `buildExcelAuthoritativePatch(fields, comparedFieldKeys, hasBusinessChanges)`.
- Consumes: `diffComparableTubiFieldMaps`, `SharePointService.updateItem`, `record.fields` e il preflight schema già presente.

- [ ] **Step 1: Scrivere i test di comportamento in `src/services/syncGuards.test.ts`**

```ts
import { shouldApplyExcelAuthoritativeUpdate } from "./syncGuards";

it("applica Excel quando un campo aziendale è diverso", () => {
  expect(shouldApplyExcelAuthoritativeUpdate(1, false)).toBe(true);
});

it("non scrive quando riga e identificativo lotto sono già allineati", () => {
  expect(shouldApplyExcelAuthoritativeUpdate(0, false)).toBe(false);
});

it("allinea l'identificativo lotto anche senza differenze aziendali", () => {
  expect(shouldApplyExcelAuthoritativeUpdate(0, true)).toBe(true);
});

it("include i vuoti espliciti ma non i campi assenti dall'Excel", () => {
  expect(buildExcelAuthoritativePatch(
    { Title: "TMA001", IdentLotto: "A", field_1: null, field_2: "559-13 R0", field_5: null },
    ["Title", "field_1", "field_2"],
    true
  )).toEqual({ Title: "TMA001", IdentLotto: "A", field_1: null, field_2: "559-13 R0" });
});
```

- [ ] **Step 2: Eseguire il test RED**

Run: `npm run test:run -- src/services/syncGuards.test.ts`

Expected: fallisce perché `shouldApplyExcelAuthoritativeUpdate` non è esportata.

- [ ] **Step 3: Implementare la guardia pura**

```ts
export function shouldApplyExcelAuthoritativeUpdate(
  changedBusinessFields: number,
  identityNeedsAlignment: boolean
): boolean {
  return changedBusinessFields > 0 || identityNeedsAlignment;
}
```

- [ ] **Step 4: Usarla nel ramo di item TUBO-MECCANICO esistente**

Calcolare `identityNeedsAlignment` dal valore `IdentLotto` memorizzato. Quando la guardia restituisce `true`, costruire prima il payload con `buildExcelAuthoritativePatch`: se cambiano campi aziendali seleziona soltanto i campi che l'Excel ha effettivamente mappato, inclusi gli eventuali vuoti espliciti; se differisce soltanto il lotto invia solo `IdentLotto`. Quindi chiamare:

```ts
await sharepointService.updateItem<Record<string, unknown>>(
  tuboMeccanicoListId,
  currentRecord.item.id,
  updateFields
);
```

Poi aggiornare la cache locale dell'item, la mappa comparabile, il contatore `updated` e il dettaglio con i campi modificati. Aggiungere `field_2` alle colonne testo gestite e le varianti delle intestazioni reali `GR. MAT.1`, `Giacenza Totale (mm)` e `Giacenza x Tagliati (mm)` alla mappa Excel. Quando la guardia restituisce `false`, incrementare `unchanged`. Non inviare un PATCH per rami duplicati o ambigui.

- [ ] **Step 5: Eseguire GREEN e le verifiche complete**

Run:

```bash
npm run test:run
npx tsc -b
npm run build
```

Expected: test, TypeScript e build terminano con codice `0`.

- [ ] **Step 6: Aggiornare il registro, commit e pubblicazione**

```bash
git add src/services/syncGuards.ts src/services/syncGuards.test.ts src/App.tsx docs/superpowers/specs/2026-09-25-inventory-sync-consolidation-design.md docs/superpowers/plans/2026-09-25-tubo-meccanico-excel-authoritative-sync.md
git commit -m "fix: rende Excel autorevole per tubo meccanico"
git push origin master
```

## Self-review

- Copertura: l'unica modifica di comportamento richiesta è coperta dal Task 1; schema, duplicati, righe fuori tabella, report e stop dopo tre errori restano nel gestore esistente.
- Nessun placeholder: ogni comando, file e ramo da modificare è indicato.
- Coerenza: la guardia è definita nel Task 1 e consumata nello stesso task.
- Review focus: i primi tre casi sono testati direttamente dalla guardia; duplicati e valori vuoti sono verificati dal percorso esistente e dal test utente post-deploy.
