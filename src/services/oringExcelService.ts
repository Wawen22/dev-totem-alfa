import { AdditionalOringKind, isAdditionalOring } from "../config/oringAdditionalColumns";
import { SharePointService, WorkbookWorksheetRange } from "./sharePointService";

type Fields = Record<string, unknown>;
const blank = (v: unknown) => v === null || v === undefined || String(v).trim() === "";
const key = (v: unknown) => String(v ?? "").trim().toUpperCase();
const headerKey = (v: unknown) => String(v ?? "").trim().toLowerCase();
const numeric = new Set(["field_4", "field_5", "field_6", "field_14", "field_15", "field_16", "QuantitaOrdinata"]);
const dates = new Set(["field_17", "field_19"]);
const originalTitle = (v: unknown) => /^Senza codice \(cliente .*\)$/.test(String(v)) ? "" : String(v ?? "");

export function oringFieldsEqual(field: string, a: unknown, b: unknown): boolean {
  if (blank(a) && blank(b)) return true;
  if (blank(a) || blank(b)) return false;
  if (numeric.has(field)) return Number(a) === Number(b);
  if (dates.has(field)) return excelDate(a) === excelDate(b);
  return String(a).replace(/\r\n/g, "\n") === String(b).replace(/\r\n/g, "\n");
}

function excelDate(value: unknown): number {
  if (typeof value === "number") return Math.floor(value);
  const time = Date.parse(String(value));
  if (!Number.isFinite(time)) throw new Error(`Data ORING non valida: ${value}`);
  return Math.floor(time / 86400000) + 25569;
}

function excelValue(field: string, value: unknown): string | number {
  if (blank(value)) return ""; // Graph null means "ignore", not "clear".
  if (dates.has(field)) return excelDate(value);
  if (numeric.has(field)) {
    const number = Number(value);
    if (!Number.isFinite(number)) throw new Error(`Valore numerico non valido per ${field}`);
    return number;
  }
  return field === "Title" ? originalTitle(value) : String(value);
}

export function changedOringFields(proposed: Fields, original: Fields): Fields {
  return Object.fromEntries(Object.entries(proposed).filter(([field, value]) =>
    field !== "RigaExcel" && !oringFieldsEqual(field, value, original[field])
  ));
}

function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
  return name;
}

export function oringWorksheetLayout(kind: AdditionalOringKind, range: WorkbookWorksheetRange) {
  const header = range.values.findIndex(row => row.some(cell => headerKey(cell) === "codice alfa"));
  if (header < 0) throw new Error(`Intestazioni del foglio ${kind} non riconosciute`);
  const names = range.values[header].map(headerKey);
  const find = (...aliases: string[]) => {
    const matches = names.flatMap((name, i) => aliases.includes(name) ? [i] : []);
    if (matches.length !== 1) throw new Error(`Colonna Excel assente o ambigua: ${aliases.join(" / ")}`);
    return matches[0];
  };
  const columns: Record<string, number> = {
    Title: find("codice alfa"), CodiceSAM: find("codice sam"), field_2: find("codice cliente"),
    field_3: find("colore mescola"), field_4: find("ø int."), field_5: find("ø est."), field_6: find("ø corda"),
    field_7: find("materiale"), field_8: find("durezza"), field_9: find("fluidi", "fluido"),
    field_12: find("commessa", "standard"), field_13: find("design temperature"), field_14: find("prezzo unitario"),
    field_15: find("giacenza"), field_18: find("nr ordine", "ordine"), QuantitaOrdinata: find("q.ta'"),
  };
  if (kind === "ORING-FKM") {
    columns.field_0 = find("posizione");
    columns.field_16 = find("prenotazione fornitore");
    columns.field_17 = find("data pr./agg.");
    columns.field_19 = find("data ordine");
  } else {
    // The original EPDM template repeats "Data" in L and S. Validate the
    // original layout before assigning meanings to its unlabelled R/T cells.
    if (range.columnIndex !== 0 || names[11] !== "data" || names[18] !== "data" || names[17] !== "" || names[19] !== "") {
      throw new Error("Struttura EPDM cambiata: verificare le colonne L/R/S/T");
    }
    columns.field_19 = 11;
    columns.field_17 = 18;
    columns.ValoreOriginaleR = 17;
    columns.NoteOriginali = 19;
    const reservation = names.indexOf("prenotazione fornitore");
    if (reservation >= 0) columns.field_16 = reservation;
  }
  return { header, columns };
}

export function matchOringWorksheetRow(kind: AdditionalOringKind, range: WorkbookWorksheetRange, original: Fields): number {
  const { header, columns } = oringWorksheetLayout(kind, range);
  const identity = ["Title", "field_2", "CodiceSAM"].filter(field => !blank(field === "Title" ? originalTitle(original[field]) : original[field]));
  if (identity.length === 0) throw new Error("Identità dell'articolo ORING mancante");
  let candidates = range.values.flatMap((row, index) => index > header && identity.every(field =>
    key(row[columns[field]]) === key(field === "Title" ? originalTitle(original[field]) : original[field])
  ) ? [index] : []);
  // Original standard/commessa distinguishes FKM61400's two source rows.
  if (candidates.length > 1) {
    for (const field of ["field_12", "field_14"]) {
      if (blank(original[field])) continue;
      const narrowed = candidates.filter(index => oringFieldsEqual(field, range.values[index][columns[field]], original[field]));
      if (narrowed.length > 0) candidates = narrowed;
      if (candidates.length === 1) break;
    }
  }
  if (candidates.length > 1 && kind === "ORING-FKM" && !blank(original.field_0)) {
    const narrowed = candidates.filter(index => key(range.values[index][columns.field_0]) === key(original.field_0));
    // A position alone can be renumbered by the worksheet formula; require
    // the original physical row as well before accepting this final fallback.
    candidates = narrowed.filter(index => range.rowIndex + index + 1 === Number(original.RigaExcel));
  }
  if (candidates.length !== 1) throw new Error(candidates.length ? "Riga Excel ORING ambigua: nessuna scrittura eseguita" : "Riga Excel ORING non trovata: verificare codice, codice cliente e Codice SAM");
  return candidates[0];
}

export class OringSaveError extends Error {
  constructor(message: string, public readonly sharePointSaved: boolean) { super(message); }
}

/** Prepare every targeted cell before mutating either source. Never PATCH a
 * whole row: formulas, formatting and unrelated Excel edits stay untouched.
 * The original SP fields are also used when the user renames an identity key.
 */
export async function saveOringWithExcel(
  service: SharePointService,
  kind: AdditionalOringKind,
  listId: string,
  itemId: string,
  original: Fields,
  proposed: Fields
): Promise<Fields> {
  if (!isAdditionalOring(kind)) throw new Error("Magazzino ORING non supportato");
  const path = (import.meta.env.VITE_ORING_EXCEL_PATH || "MAGAZZINO/MAGAZZINI ATTIVI/2_ORING.xlsx").trim();
  const driveName = (import.meta.env.VITE_ORING_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "TOTEM ALFA").trim();
  const driveId = import.meta.env.VITE_ORING_EXCEL_DRIVE_ID || await service.getDriveIdByName(driveName);
  if (!driveId) throw new Error(`Libreria Excel ${driveName} non trovata`);
  const file = await service.getDriveItemByPath(path, driveId);
  const sessionId = await service.createWorkbookSessionByItemId(file.id, { persistChanges: true }, driveId);
  const sheet = kind.slice(6);
  let sharePointSaved = false;
  let failed = false;
  try {
    const range = await service.readWorkbookWorksheetByItemId(file.id, sheet, { sessionId }, driveId);
    const { columns } = oringWorksheetLayout(kind, range);
    const rowIndex = matchOringWorksheetRow(kind, range, original);
    const rowNumber = range.rowIndex + rowIndex + 1;
    const patch = changedOringFields(proposed, original);
    // Clicking Save without new edits is also an explicit retry of a previous
    // SharePoint-only save; compare all mapped values with Excel in that case.
    const intent = Object.keys(patch).length ? patch : proposed;
    let reservationHeader: { address: string; value: string } | undefined;
    if (kind === "ORING-EPDM" && columns.field_16 === undefined && !blank(intent.field_16)) {
      // R contains unrelated original values. Reserve a dedicated column after
      // the entire used range, without repurposing or overwriting source cells.
      columns.field_16 = Math.max(20, ...range.values.map(row => row.length));
      const { header } = oringWorksheetLayout(kind, range);
      reservationHeader = { address: `${columnName(range.columnIndex + columns.field_16)}${range.rowIndex + header + 1}`, value: "PRENOTAZIONE FORNITORE" };
    }
    const writes: { field: string; column: number; value: string | number; address: string }[] = [];
    for (const [field, value] of Object.entries(intent)) {
      if (field === "RigaExcel") continue;
      const column = columns[field];
      if (column === undefined) {
        if (!blank(value)) throw new Error(`Il campo ${field} non ha una colonna nel foglio ${sheet}: nessuna scrittura eseguita`);
        continue;
      }
      const next = excelValue(field, value);
      if (oringFieldsEqual(field, next, range.values[rowIndex]?.[column])) continue;
      if (String(range.formulas?.[rowIndex]?.[column] ?? "").startsWith("=")) {
        if (field in patch) throw new Error(`La cella ${sheet}!${columnName(range.columnIndex + column)}${rowNumber} contiene una formula: modifica i valori da cui dipende`);
        continue;
      }
      writes.push({ field, column, value: next, address: `${columnName(range.columnIndex + column)}${rowNumber}` });
    }
    // All validation above precedes the SharePoint update.
    if (Object.keys(patch).length) await service.updateItem(listId, itemId, patch);
    sharePointSaved = true;
    if (reservationHeader) await service.updateWorkbookCellByAddress(file.id, sheet, reservationHeader.address, reservationHeader.value, { sessionId }, driveId);
    for (const write of writes) await service.updateWorkbookCellByAddress(file.id, sheet, write.address, write.value, {
      sessionId, ...(dates.has(write.field) ? { numberFormat: "dd/mm/yyyy" } : {}),
    }, driveId);
    const after = await service.readWorkbookWorksheetByItemId(file.id, sheet, { sessionId }, driveId);
    for (const write of writes) {
      const actual = after.values[rowNumber - 1 - after.rowIndex]?.[range.columnIndex + write.column - after.columnIndex];
      if (!oringFieldsEqual(write.field, write.value, actual)) throw new Error(`Verifica Excel fallita su ${sheet}!${write.address}`);
    }
    const result: Fields = { ...original, ...patch, RigaExcel: rowNumber };
    const derived: Fields = { RigaExcel: rowNumber };
    for (const [field, column] of Object.entries(columns)) {
      if (!String(range.formulas?.[rowIndex]?.[column] ?? "").startsWith("=")) continue;
      const actual = after.values[rowNumber - 1 - after.rowIndex]?.[range.columnIndex + column - after.columnIndex];
      if (actual === undefined || (typeof actual === "string" && actual.startsWith("#"))) throw new Error(`Formula Excel non valida nel campo ${field}`);
      const typed = field === "field_0" ? String(actual) : actual;
      if (!oringFieldsEqual(field, result[field], typed)) derived[field] = typed;
    }
    await service.updateItem(listId, itemId, derived);
    return { ...result, ...derived };
  } catch (error) {
    failed = true;
    const message = error instanceof Error ? error.message : "Errore Excel";
    throw new OringSaveError(sharePointSaved
      ? `SharePoint aggiornato, ma sincronizzazione Excel non completata: ${message}. Riprova il salvataggio.`
      : message, sharePointSaved);
  } finally {
    try {
      await service.closeWorkbookSessionByItemId(file.id, sessionId, driveId);
    } catch (error) {
      if (!failed) throw new OringSaveError("Scritture eseguite, ma chiusura della sessione Excel non confermata. Riprova il salvataggio per verificarle.", sharePointSaved);
      console.warn("Chiusura sessione Excel ORING non riuscita", error);
    }
  }
}
