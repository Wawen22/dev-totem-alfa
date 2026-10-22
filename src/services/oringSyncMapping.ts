import { OringKind } from "../config/oringSync";
import { SharePointColumnDefinition, WorkbookWorksheetRange } from "./sharePointService";
import { oringWorksheetLayout } from "./oringExcelService";

export type OringFields = Record<string, unknown>;
export const blankOringValue = (value: unknown) => value == null || String(value).trim() === "";
export const oringKey = (value: unknown) => String(value ?? "").trim().toUpperCase();
export const oringCode = (value: unknown) => /^Senza codice \(cliente .*\)$/.test(String(value)) ? "" : String(value ?? "");
export const excelColumnName = (index: number): string => {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
  return name;
};

export function oringSyncLayout(kind: OringKind, range: WorkbookWorksheetRange) {
  if (kind !== "ORING-NBR") return oringWorksheetLayout(kind === "ORING-HNBR" ? "ORING-FKM" : kind, range);
  const normalize = (v: unknown) => String(v ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  const header = range.values.findIndex(row => row.some(cell => normalize(cell) === "codice alfa"));
  if (header < 0) throw new Error("Intestazioni NBR non trovate");
  const names = range.values[header].map(normalize);
  const find = (...aliases: string[]) => {
    const matches = names.flatMap((name, index) => aliases.includes(name) ? [index] : []);
    if (matches.length !== 1) throw new Error(`Colonna NBR assente o ambigua: ${aliases.join(" / ")}`);
    return matches[0];
  };
  if (range.columnIndex !== 0 || names[9] !== "fluidi" || names[10] || names[11]) throw new Error("Struttura delle colonne Fluidi NBR cambiata");
  return { header, columns: {
    Title: find("codice alfa"), field_1: find("codice cliente"), field_2: find("colore mescola"),
    field_3: find("ø int."), field_4: find("ø est."), field_5: find("ø corda"),
    field_6: find("materiale"), field_7: find("durezza"), field_8: find("fluidi"), field_9: 10, field_10: 11,
    field_11: find("commessa"), field_12: find("design temperature"), field_13: find("prezzo unitario"),
    field_14: find("q.tà minima"), field_15: find("prenotazione"), field_16: find("giacenza"), field_17: find("data pr./agg."),
    // These Excel columns have no counterpart in the legacy NBR SharePoint list.
    CodiceSAM: find("codice sam"),
  } };
}

export const customerField = (kind: OringKind) => kind === "ORING-NBR" ? "field_1" : "field_2";
export const standardField = (kind: OringKind) => kind === "ORING-NBR" ? "field_11" : "field_12";
export const stockField = (kind: OringKind) => kind === "ORING-NBR" ? "field_16" : "field_15";

function calendarDate(value: unknown, fromExcel: boolean): string {
  if (typeof value === "number" || /^\d+(\.\d+)?$/.test(String(value))) {
    const date = new Date((Math.floor(Number(value)) - 25569) * 86400000);
    if (!Number.isFinite(date.getTime())) throw new Error(`Data Excel non valida: ${value}`);
    return date.toISOString().slice(0, 10);
  }
  const text = String(value).trim();
  const italian = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})$/);
  if (italian) {
    const year = italian[3].length === 2 ? `20${italian[3]}` : italian[3];
    const iso = `${year}-${italian[2].padStart(2, "0")}-${italian[1].padStart(2, "0")}`;
    const date = new Date(`${iso}T00:00:00Z`);
    if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== iso) throw new Error(`Data Excel non valida: ${value}`);
    return iso;
  }
  const date = new Date(text);
  if (!Number.isFinite(date.getTime())) throw new Error(`Data non valida: ${value}`);
  // Legacy date-only SharePoint fields may be midnight Rome encoded as 23:00
  // or 22:00 UTC on the preceding day. Do not export that preceding UTC day.
  if (!fromExcel && /T/.test(text)) {
    const parts = new Intl.DateTimeFormat("en-CA", { timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(date);
    const get = (type: string) => parts.find(part => part.type === type)!.value;
    return `${get("year")}-${get("month")}-${get("day")}`;
  }
  return date.toISOString().slice(0, 10);
}

export function valueForSharePoint(value: unknown, column: SharePointColumnDefinition, fromExcel = true): unknown {
  if (blankOringValue(value)) return null;
  if (column.dateTime) return `${calendarDate(value, fromExcel)}T00:00:00.000Z`;
  if (column.number || column.currency) {
    const number = typeof value === "number" ? value : Number(String(value).trim().replace(",", "."));
    if (!Number.isFinite(number)) throw new Error(`Numero non valido in ${column.name}: ${value}`);
    return number;
  }
  return String(value).replace(/\r\n/g, "\n");
}

export function valueForOringExcel(value: unknown, column: SharePointColumnDefinition): string | number {
  if (blankOringValue(value)) return "";
  if (column.name === "Title") return oringCode(value);
  if (column.dateTime) return Date.parse(`${calendarDate(value, false)}T00:00:00Z`) / 86400000 + 25569;
  return valueForSharePoint(value, column, false) as string | number;
}

export function sameOringSyncValue(a: unknown, b: unknown, column: SharePointColumnDefinition, bFromExcel = false): boolean {
  if (column.name === "Title") { a = oringCode(a); b = oringCode(b); }
  if (blankOringValue(a) && blankOringValue(b)) return true;
  if (blankOringValue(a) || blankOringValue(b)) return false;
  return valueForSharePoint(a, column, false) === valueForSharePoint(b, column, bFromExcel);
}

/** Mutual, one-to-one matching prevents two source records from overwriting
 * the same destination. Ambiguity is a result to review, never "first wins".
 */
export function matchOringSyncRecords(kind: OringKind, excel: OringFields[], sharepoint: OringFields[]) {
  const customer = customerField(kind), standard = standardField(kind);
  const score = (source: OringFields, target: OringFields): number => {
    const code = oringKey(oringCode(source.Title)), otherCode = oringKey(oringCode(target.Title));
    const client = oringKey(source[customer]), otherClient = oringKey(target[customer]);
    const sameCode = !!code && code === otherCode;
    const sameClient = !!client && client === otherClient;
    const samePosition = (kind === "ORING-HNBR" || kind === "ORING-FKM")
      && !blankOringValue(source.field_0) && !blankOringValue(target.field_0)
      && oringKey(source.field_0) === oringKey(target.field_0);
    if (!sameCode && !sameClient) return 0;
    if (client && otherClient && client !== otherClient && !(sameCode && samePosition)) return 0;
    let result = sameCode ? 50 + (sameClient ? 10 : 0) : 20;
    if (source.CodiceSAM && target.CodiceSAM && oringKey(source.CodiceSAM) === oringKey(target.CodiceSAM)) result += 5;
    if (kind === "ORING-HNBR" || kind === "ORING-FKM") {
      if (!blankOringValue(source.field_0) && !blankOringValue(target.field_0)) {
        if (samePosition) result += 40;
      }
    }
    if (oringKey(source[standard]) === oringKey(target[standard])) result += 3;
    return result;
  };
  const scores = excel.map(source => sharepoint.map(target => score(source, target)));
  const best = (values: number[]) => {
    const maximum = Math.max(0, ...values);
    return maximum ? values.flatMap((value, index) => value === maximum ? [index] : []) : [];
  };
  const excelBest = scores.map(best);
  const spBest = sharepoint.map((_, index) => best(scores.map(row => row[index])));
  const matches = new Map<number, number>();
  excelBest.forEach((targets, index) => {
    if (targets.length === 1 && spBest[targets[0]].length === 1 && spBest[targets[0]][0] === index) matches.set(index, targets[0]);
  });
  // Repeated source rows must not become duplicate new SharePoint records.
  const logicalKey = (fields: OringFields) => JSON.stringify([
    oringKey(oringCode(fields.Title)), oringKey(fields[customer]), oringKey(fields.field_0), oringKey(fields[standard]),
  ]);
  const counts = (rows: OringFields[]) => rows.reduce((map, row) => map.set(logicalKey(row), (map.get(logicalKey(row)) || 0) + 1), new Map<string, number>());
  const excelCounts = counts(excel), spCounts = counts(sharepoint);
  const newExcel = excel.map((_, index) => index).filter(index => excelBest[index].length === 0 && excelCounts.get(logicalKey(excel[index])) === 1);
  const newSharePoint = sharepoint.map((_, index) => index).filter(index => spBest[index].length === 0 && spCounts.get(logicalKey(sharepoint[index])) === 1);
  const matchedSp = new Set(matches.values());
  return {
    matches, newExcel, newSharePoint,
    ambiguousExcel: excel.map((_, i) => i).filter(i => !matches.has(i) && !newExcel.includes(i)),
    ambiguousSharePoint: sharepoint.map((_, i) => i).filter(i => !matchedSp.has(i) && !newSharePoint.includes(i)),
  };
}
