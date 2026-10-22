import type { AdditionalOringKind } from "../config/oringAdditionalColumns";

type ColumnType = "text" | "number" | "dateTime";
export const additionalOringSchema: { name: string; displayName: string; type: ColumnType }[] = [
  { name: "field_0", displayName: "Posizione", type: "text" },
  { name: "field_2", displayName: "Codice cliente", type: "text" },
  { name: "field_3", displayName: "Colore Mescola", type: "text" },
  { name: "field_4", displayName: "Ø Int.", type: "number" },
  { name: "field_5", displayName: "Ø Est.", type: "number" },
  { name: "field_6", displayName: "Ø Corda", type: "number" },
  { name: "field_7", displayName: "Materiale", type: "text" },
  { name: "field_8", displayName: "Durezza", type: "text" },
  { name: "field_9", displayName: "Fluidi", type: "text" },
  { name: "field_12", displayName: "Commessa / Standard", type: "text" },
  { name: "field_13", displayName: "Design temperature", type: "text" },
  { name: "field_14", displayName: "Prezzo unitario", type: "number" },
  { name: "field_15", displayName: "Giacenza", type: "number" },
  { name: "field_16", displayName: "Prenotazione fornitore", type: "number" },
  { name: "field_17", displayName: "Data prelievo / aggiornamento", type: "dateTime" },
  { name: "field_18", displayName: "NR Ordine", type: "text" },
  { name: "field_19", displayName: "Data ordine", type: "dateTime" },
  { name: "CodiceSAM", displayName: "Codice SAM", type: "text" },
  { name: "QuantitaOrdinata", displayName: "Q.tà ordinata", type: "number" },
  { name: "ValoreOriginaleR", displayName: "Valore originale colonna R", type: "text" },
  { name: "NoteOriginali", displayName: "Note originali", type: "text" },
  { name: "RigaExcel", displayName: "Riga Excel origine", type: "number" },
];

const normalize = (value: unknown) => String(value ?? "").trim().toLowerCase();
const isBlank = (value: unknown) => value === null || value === undefined || (typeof value === "string" && value.trim() === "");
const parseNumber = (value: unknown): number => {
  const result = typeof value === "number" ? value : Number(String(value).trim().replace(",", "."));
  if (!Number.isFinite(result)) throw new Error(`Numero Excel non valido: ${value}`);
  return result;
};
const parseDate = (value: unknown): string => {
  if (typeof value === "number") return new Date(Math.round((value - 25569) * 86400000)).toISOString();
  const text = String(value).trim();
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  const iso = match ? `${match[3].length === 2 ? `20${match[3]}` : match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}T00:00:00.000Z` : text;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) throw new Error(`Data Excel non valida: ${value}`);
  return date.toISOString();
};

/** Read-only preparation: preserve source rows, blanks, and repeated codes.
 * Row numbers refer to usedRange A1, as returned by the MCP worksheet reader.
 * EPDM's unlabelled R/T cells are retained without inventing their meaning.
 */
export function prepareOringImport(kind: AdditionalOringKind, values: unknown[][]) {
  const headerIndex = values.findIndex((row) => row.some((cell) => normalize(cell) === "codice alfa"));
  if (headerIndex < 0) throw new Error(`Intestazioni ${kind} non trovate`);
  const header = values[headerIndex].map(normalize);
  const find = (...aliases: string[]) => {
    const index = header.findIndex((cell) => aliases.includes(cell));
    if (index < 0) throw new Error(`Colonna ${aliases.join(" / ")} mancante in ${kind}`);
    return index;
  };
  const map: Record<string, number> = {
    Title: find("codice alfa"), CodiceSAM: find("codice sam"), field_2: find("codice cliente"),
    field_3: find("colore mescola"), field_4: find("ø int."), field_5: find("ø est."),
    field_6: find("ø corda"), field_7: find("materiale"), field_8: find("durezza"),
    field_9: find("fluidi", "fluido"), field_12: find("commessa", "standard"),
    field_13: find("design temperature"), field_14: find("prezzo unitario"),
    field_15: find("giacenza"), field_18: find("nr ordine", "ordine"),
    field_19: find("data ordine", "data"), QuantitaOrdinata: find("q.ta'"),
  };
  if (kind === "ORING-FKM") {
    map.field_0 = find("posizione");
    map.field_16 = find("prenotazione fornitore");
    map.field_17 = find("data pr./agg.");
  } else {
    // "Data" is the order date in L; the update date is explicitly in S.
    if (header[18] !== "data" || header[17] !== "" || header[19] !== "") {
      throw new Error("Struttura delle colonne finali EPDM cambiata: verificare R/S/T");
    }
    map.field_17 = 18;
    map.ValoreOriginaleR = 17;
    map.NoteOriginali = 19;
  }
  const types = new Map(additionalOringSchema.map((column) => [column.name, column.type]));
  return values.slice(headerIndex + 1).flatMap((row, index) => {
    if (isBlank(row[map.Title]) && isBlank(row[map.field_2])) return [];
    const fields: Record<string, unknown> = { RigaExcel: headerIndex + index + 2 };
    for (const [field, column] of Object.entries(map)) {
      const value = row[column];
      if (isBlank(value)) continue;
      fields[field] = types.get(field) === "number" ? parseNumber(value)
        : types.get(field) === "dateTime" ? parseDate(value) : String(value);
    }
    // SharePoint's standard Title column is required on these lists. Use an
    // explicit absence label, never fabricate an Alfa product code.
    if (!fields.Title) fields.Title = `Senza codice (cliente ${fields.field_2 || "-"})`;
    return [{ fields }];
  });
}
