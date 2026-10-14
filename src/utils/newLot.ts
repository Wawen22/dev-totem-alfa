import { forgiatiColumns, type ForgiatoColumn } from "../config/forgiatiColumns";
import { tubiColumns } from "../config/tubiColumns";
import { tuboMeccanicoColumns } from "../config/tuboMeccanicoColumns";
import { flangeColumns } from "../config/flangeColumns";

export type LotKind = "FORGIATI" | "TUBI" | "TUBO-MECCANICO" | "FLANGE";
const columns: Record<LotKind, ForgiatoColumn[]> = { FORGIATI: forgiatiColumns, TUBI: tubiColumns, "TUBO-MECCANICO": tuboMeccanicoColumns, FLANGE: flangeColumns };
const excluded = new Set(["Title", "IdentLotto", "Modified", "Created", "field_26", "PrezzoEuroKg", "Archiviato"]);

export function newLotColumns(kind: LotKind): ForgiatoColumn[] {
  return columns[kind].filter((column) => !column.computed && !excluded.has(column.field));
}

// Only identity and explicitly entered values: no predecessor data or reset defaults.
export function buildNewLotPayload(kind: LotKind, title: string, identLotto: string, values: Record<string, unknown>): Record<string, unknown> {
  const payload: Record<string, unknown> = { Title: title.trim(), IdentLotto: identLotto.trim().toUpperCase() };
  for (const column of newLotColumns(kind)) {
    const raw = values[column.field];
    if (raw === null || raw === undefined || String(raw).trim() === "") continue;
    if (column.type === "date") {
      const date = new Date(String(raw));
      if (!Number.isFinite(date.getTime())) throw new Error(`Data non valida: ${column.label}`);
      payload[column.field] = date.toISOString();
    } else if (column.type === "number") {
      const number = Number(String(raw).trim().replace(",", "."));
      if (!Number.isFinite(number)) throw new Error(`Numero non valido: ${column.label}`);
      payload[column.field] = number;
    } else payload[column.field] = String(raw).trim();
  }
  return payload;
}
