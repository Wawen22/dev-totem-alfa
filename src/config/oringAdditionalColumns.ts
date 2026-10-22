import type { ForgiatoColumn } from "./forgiatiColumns";
import { oringHnbrColumns } from "./oringHnbrColumns";

export type AdditionalOringKind = "ORING-FKM" | "ORING-EPDM";
export const isAdditionalOring = (kind: string): kind is AdditionalOringKind =>
  kind === "ORING-FKM" || kind === "ORING-EPDM";
export const usesHnbrFields = (kind: string): boolean =>
  kind === "ORING-HNBR" || isAdditionalOring(kind);

// Keep the operational field mapping shared with HNBR. The source worksheets
// differ: their initial import maps by header, never by SharePoint field order.
export const oringAdditionalColumns: ForgiatoColumn[] = [
  ...oringHnbrColumns.filter((column) => !["field_10", "field_11"].includes(column.field)),
  { field: "CodiceSAM", label: "Codice SAM" },
  { field: "QuantitaOrdinata", label: "Q.tà ordinata", type: "number" },
  { field: "ValoreOriginaleR", label: "Valore originale colonna R" },
  { field: "NoteOriginali", label: "Note originali" },
  { field: "RigaExcel", label: "Riga Excel origine", type: "number" },
];

export const getOringItemTitle = (fields: Record<string, unknown>): string =>
  String(fields.Title || "").trim() || `Senza codice (cliente ${fields.field_2 || "-"})`;
