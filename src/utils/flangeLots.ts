import { flangeColumns } from "../config/flangeColumns";
import type { SharePointListItem } from "../types/sharepoint";

export const effectiveFlangeLot = (value: unknown): string => String(value ?? "").trim().toUpperCase() || "A";

export function assertUniqueFlangeLots(items: SharePointListItem<Record<string, unknown>>[]): void {
  const used = new Set<string>();
  for (const item of items) {
    const title = String(item.fields.Title ?? "").trim().toUpperCase();
    if (!title) continue;
    const label = effectiveFlangeLot(item.fields.IdentLotto);
    const key = JSON.stringify([title, label]);
    if (used.has(key)) throw new Error(`Identificativo lotto ambiguo: ${title} (${label}). Correggerlo nel Pannello Admin prima di sincronizzare.`);
    used.add(key);
  }
}

export function flangeLotLabels(items: SharePointListItem<Record<string, unknown>>[]): Map<string, string> {
  const labels = new Map<string, string>();
  // Historical blank identifiers represent A. Never silently rename duplicate
  // persisted lots: their identity must be corrected explicitly in the admin.
  for (const item of items) labels.set(item.id, effectiveFlangeLot(item.fields.IdentLotto));
  return labels;
}

export function nextFlangeLot(existing: unknown[]): string {
  const used = new Set(existing.map((value) => String(value ?? "").trim().toUpperCase()));
  // Excel-style letters also support codes with more than 26 lots.
  for (let index = 1; ; index++) {
    let value = index;
    let label = "";
    while (value > 0) {
      value--;
      label = String.fromCharCode(65 + value % 26) + label;
      value = Math.floor(value / 26);
    }
    if (!used.has(label)) return label;
  }
}

export function buildFlangeLotPayload(
  source: Record<string, unknown>,
  options: { identLotto: string; colata: string; ordine: string; codiceSam: string; dataOrdine: string | null }
): Record<string, unknown> {
  const payload: Record<string, unknown> = {};
  const reset = new Set(["IdentLotto", "Quantita", "GiacenzaMm", "NumeroOrdine", "DataOrdine", "CodiceSAM", "NumeroColata", "NumeroBolla", "NumeroCertificato", "DataConsegna", "Commessa", "Commessa2"]);
  for (const column of flangeColumns) {
    if (column.hidden || reset.has(column.field)) continue;
    const value = source[column.field];
    if (value !== null && value !== undefined && value !== "") payload[column.field] = value;
  }
  return {
    ...payload,
    IdentLotto: options.identLotto,
    NumeroColata: options.colata.trim(),
    NumeroOrdine: options.ordine.trim(),
    CodiceSAM: options.codiceSam.trim(),
    DataOrdine: options.dataOrdine,
    Quantita: 0,
    GiacenzaMm: 0,
  };
}
