# Fix SharePoint Sync Overwrite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Impedire che Power Automate sovrascriva i dati aggiornati dall'utente quando l'aggiornamento del file Excel fallisce, e invalidare la cache locale dopo ogni salvataggio riuscito.

**Architecture:** Tre modifiche chirurgiche: (1) bloccare la chiamata a Power Automate se almeno un aggiornamento Excel fallisce, impostando lo stato a "error" anziché "success"; (2) esportare una funzione `clearCacheKeys` da `useCachedList.ts` per invalidare le voci di cache specifiche; (3) chiamare `clearCacheKeys` e svuotare carrello/editValues dopo un salvataggio completamente riuscito.

**Tech Stack:** React, TypeScript, Microsoft Graph API, SharePoint, Power Automate

---

## Contesto del Bug

Il flusso attuale quando l'aggiornamento Excel fallisce:

```
STEP 1: updateItem su SharePoint list  →  ✅ OK  (giacenza = 150)
STEP 2: updateWorkbook su Excel        →  ❌ FALLISCE (API non trovata)
STEP 3: flowService.invoke()           →  ✅ Viene chiamato comunque!
         └─ Power Automate legge Excel → giacenza = 100 (vecchio)
         └─ Power Automate scrive su SharePoint list → giacenza = 100 ← DANNO
```

---

## File da modificare

| File | Modifica |
|------|----------|
| `src/hooks/useCachedList.ts` | Esportare `clearCacheKeys(keys: string[]): void` |
| `src/App.tsx` | Fix handleSave: blocco PA su errore Excel + clearCache + clearCart |

---

### Task 1: Esportare `clearCacheKeys` da `useCachedList.ts`

**File:**
- Modify: `src/hooks/useCachedList.ts`

- [ ] **Step 1: Aggiungere la funzione esportata subito dopo la dichiarazione di `globalCache`**

In `src/hooks/useCachedList.ts`, dopo la riga 7 (`const CACHE_DURATION = ...`), aggiungere:

```typescript
export function clearCacheKeys(keys: string[]): void {
  keys.forEach((key) => {
    delete globalCache[key];
  });
}
```

Il file modificato sarà:
```typescript
// Cache globale che persiste finché l'app non viene ricaricata (F5)
const globalCache: Record<string, { data: any[]; timestamp: number }> = {};
const CACHE_DURATION = 5 * 60 * 1000; // 5 minuti di validità cache default

export function clearCacheKeys(keys: string[]): void {
  keys.forEach((key) => {
    delete globalCache[key];
  });
}
```

- [ ] **Step 2: Verificare che il file compili**

```bash
cd "/home/rnebili/Progetti/TEL&CO/MICROSOFT/[TLCO-PRODUCTS]/ALFA ENG/totem-alfa/DEV"
npx tsc --noEmit 2>&1 | head -30
```

Expected: nessun errore relativo a `useCachedList.ts`

---

### Task 2: Importare `clearCacheKeys` in `App.tsx`

**File:**
- Modify: `src/App.tsx` (riga 892)

- [ ] **Step 1: Aggiornare l'import di `useCachedList`**

Cercare la riga 892 di `src/App.tsx`:
```typescript
import { useCachedList } from "./hooks/useCachedList";
```

Sostituirla con:
```typescript
import { useCachedList, clearCacheKeys } from "./hooks/useCachedList";
```

---

### Task 3: Fix `handleSave` — bloccare PA su errore Excel e gestire post-salvataggio

**File:**
- Modify: `src/App.tsx` (righe 5589–5600)

Questo è il cuore del fix. Il codice attuale (righe 5589–5600):

```typescript
      const baseMessage = excelErrors.length > 0
        ? `Giacenza aggiornata; Excel non aggiornato: ${excelErrors.join(" | ")}`
        : "Giacenza aggiornata con successo.";
      setSaveStatus("success");
      setSaveMessage(baseMessage);

      try {
        await flowService.invoke(flowPayload);
      } catch (flowErr: any) {
        console.error("Errore Power Automate", flowErr);
        setSaveMessage(`${baseMessage}; notifica Teams non inviata: ${flowErr?.message || "errore flow"}`);
      }
```

- [ ] **Step 1: Sostituire il blocco righe 5589–5600 con la logica corretta**

```typescript
      if (excelErrors.length > 0) {
        // Excel non aggiornato: NON notificare Power Automate per evitare
        // che legga dati obsoleti da Excel e sovrascriva SharePoint.
        setSaveStatus("error");
        setSaveMessage(`Giacenza aggiornata su SharePoint, ma Excel non aggiornato: ${excelErrors.join(" | ")} — Power Automate NON notificato per evitare sovrascrittura dati.`);
        return;
      }

      // Salvataggio completamente riuscito: invalidare cache e pulire carrello
      const updatedSources = Array.from(new Set(cartList.map((i) => i.source)));
      const cacheKeysToInvalidate = updatedSources.map((source) => {
        switch (source) {
          case "FORGIATI": return "forgiati";
          case "TUBI": return "tubi";
          case "ORING-HNBR": return "oring-hnbr";
          case "ORING-NBR": return "oring-nbr";
          case "SPARK-GUPS": return "spark-gups";
          case "FILO-FLUSSO": return "filo-flusso";
          case "TUBO-MECCANICO": return "tubo-meccanico";
          default: return null;
        }
      }).filter((k): k is string => k !== null);
      clearCacheKeys(cacheKeysToInvalidate);
      setCartItems({});
      setEditValues({});

      setSaveStatus("success");
      setSaveMessage("Giacenza aggiornata con successo.");

      try {
        await flowService.invoke(flowPayload);
      } catch (flowErr: any) {
        console.error("Errore Power Automate", flowErr);
        setSaveMessage(`Giacenza aggiornata con successo; notifica Teams non inviata: ${flowErr?.message || "errore flow"}`);
      }
```

Nota: la riga `const sources = ...` già presente a riga 5531 definisce `sources` per il payload. La nuova variabile si chiama `updatedSources` per evitare conflitti.

- [ ] **Step 2: Verificare che `setCartItems` e `setEditValues` siano nelle dipendenze del `useCallback` di `handleSave`**

Cercare la riga ~5606 (la chiusura del `useCallback`):
```typescript
  }, [
    accounts,
    cartList,
    editValues,
    flowService,
    ...
```

Aggiungere `setCartItems` e `setEditValues` all'array di dipendenze se non già presenti. Poiché sono state create con `useState`, React garantisce referenze stabili — non causano re-render extra. Verificare che siano nell'array.

---

### Task 4: Build e verifica finale

- [ ] **Step 1: Compilare TypeScript senza errori**

```bash
cd "/home/rnebili/Progetti/TEL&CO/MICROSOFT/[TLCO-PRODUCTS]/ALFA ENG/totem-alfa/DEV"
npx tsc --noEmit 2>&1
```

Expected: nessun errore

- [ ] **Step 2: Build di produzione**

```bash
cd "/home/rnebili/Progetti/TEL&CO/MICROSOFT/[TLCO-PRODUCTS]/ALFA ENG/totem-alfa/DEV"
npm run build 2>&1 | tail -20
```

Expected: build completata senza errori

- [ ] **Step 3: Commit**

```bash
cd "/home/rnebili/Progetti/TEL&CO/MICROSOFT/[TLCO-PRODUCTS]/ALFA ENG/totem-alfa/DEV"
git add src/hooks/useCachedList.ts src/App.tsx
git commit -m "fix: blocca Power Automate se Excel fallisce, invalida cache post-salvataggio

- Quando l'aggiornamento Excel fallisce, il salvataggio viene marcato come
  errore e Power Automate NON viene notificato. Questo impedisce che PA
  legga dati obsoleti da Excel e li riscrva su SharePoint, sovrascrivendo
  le modifiche dell'utente.
- Dopo un salvataggio completamente riuscito, le voci di cache rilevanti
  vengono invalidate e il carrello svuotato, garantendo che i dati
  visualizzati siano sempre aggiornati.

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

---

## Riepilogo impatto

| Scenario | Prima del fix | Dopo il fix |
|----------|---------------|-------------|
| Excel update fallisce | Status: success, PA notificato → sovrascrittura | Status: error, PA bloccato → dati al sicuro |
| Excel update riesce | Cache non invalidata, carrello rimane | Cache invalidata, carrello svuotato |
| Ricaricamento dopo save | Dati vecchi dalla cache | Dati freschi da SharePoint |
