import { isOringKind, OringKind, oringSheetName } from "../config/oringSync";
import { SharePointService, SharePointColumnDefinition, WorkbookWorksheetRange } from "./sharePointService";
import { SyncDetailItem, SyncDetailKey, SyncResult } from "../types/sync";
import { clearCacheKeys } from "../hooks/useCachedList";
import { MAX_CONSECUTIVE_SYNC_WRITE_FAILURES } from "./syncGuards";
import {
  blankOringValue as blank, customerField, excelColumnName, matchOringSyncRecords,
  OringFields, oringCode, oringSyncLayout, sameOringSyncValue, stockField,
  valueForOringExcel, valueForSharePoint,
} from "./oringSyncMapping";

export type OringSyncDirection = "excel-to-totem" | "totem-to-excel";
type ExcelRecord = { fields: OringFields; row: number; index: number };
type CellWrite = { field: string; address: string; value: string | number; format?: string };
type RowPlan = { fields: OringFields; itemId: string; row: number; source?: ExcelRecord; created: boolean; expandTable: boolean; writes: CellWrite[]; formulas: { field: string; address: string; formula: string }[] };

function recordsFromRange(kind: OringKind, range: WorkbookWorksheetRange, schema: SharePointColumnDefinition[]) {
  const layout = oringSyncLayout(kind, range);
  const byName = new Map(schema.filter(c => !c.readOnly && !c.hidden).map(c => [c.name, c]));
  const required = ["Title", customerField(kind), stockField(kind)];
  const missing = required.filter(field => !byName.has(field) || layout.columns[field] === undefined);
  if (missing.length) throw new Error(`Colonne ORING mancanti: ${missing.join(", ")}`);
  const mapped = Object.entries(layout.columns).filter(([field]) => byName.has(field));
  const records: ExcelRecord[] = [];
  const invalid: SyncDetailItem[] = [];
  range.values.forEach((row, index) => {
    if (index <= layout.header) return;
    if (blank(row[layout.columns.Title]) && blank(row[layout.columns[customerField(kind)]])) return;
    const rowNumber = range.rowIndex + index + 1;
    try {
      const fields = Object.fromEntries(mapped.map(([field, column]) => [field, valueForSharePoint(row[column], byName.get(field)!)]));
      if (!fields.Title) fields.Title = `Senza codice (cliente ${fields[customerField(kind)] || "-"})`;
      if (byName.has("RigaExcel")) fields.RigaExcel = rowNumber;
      records.push({ fields, row: rowNumber, index });
    } catch (error) {
      invalid.push({ label: `Riga Excel ${rowNumber}`, detail: error instanceof Error ? error.message : "Dati non validi" });
    }
  });
  return { ...layout, byName, mapped, records, invalid };
}

/** Source-authoritative manual synchronization. Neither direction deletes
 * destination-only records. Every matching decision is made before writes.
 */
export async function syncOringWarehouse(
  service: SharePointService,
  kind: OringKind,
  listId: string | undefined,
  direction: OringSyncDirection,
  onProgress?: (message: string) => void
): Promise<SyncResult> {
  const detailMap: Record<SyncDetailKey, SyncDetailItem[]> = { updated: [], created: [], unchanged: [], skipped: [], duplicates: [], "sharepoint-only": [] };
  const add = (key: SyncDetailKey, fields: OringFields, detail: string, reference?: string) =>
    detailMap[key].push({ label: String(fields.Title || "(senza codice)"), code: String(fields.Title || ""), reference, detail });
  const label = direction === "excel-to-totem" ? "Excel → Totem" : "Totem → Excel";
  let session: string | undefined, fileId = "", driveId: string | undefined;
  let failure: string | undefined;
  try {
    if (!isOringKind(kind) || !listId) throw new Error(`Lista ${kind} non configurata`);
    const driveName = import.meta.env.VITE_ORING_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "TOTEM ALFA";
    driveId = import.meta.env.VITE_ORING_EXCEL_DRIVE_ID || await service.getDriveIdByName(driveName) || undefined;
    if (!driveId) throw new Error(`Libreria ${driveName} non trovata`);
    const path = import.meta.env.VITE_ORING_EXCEL_PATH || "MAGAZZINO/MAGAZZINI ATTIVI/2_ORING.xlsx";
    const sheet = oringSheetName(kind);
    onProgress?.(`Lettura ${sheet} e lista ${kind}...`);
    const file = await service.getDriveItemByPath(path, driveId); fileId = file.id;
    session = await service.createWorkbookSessionByItemId(fileId, { persistChanges: true }, driveId);
    const [range, items, schema] = await Promise.all([
      service.readWorkbookWorksheetByItemId(fileId, sheet, { sessionId: session }, driveId),
      service.listItems<OringFields>(listId), service.listColumns(listId),
    ]);
    const parsed = recordsFromRange(kind, range, schema);
    const { records, mapped, columns, byName } = parsed;
    detailMap.skipped.push(...parsed.invalid);
    // With malformed Excel records, absence is unproven: never append/create
    // records on the opposite side based on that incomplete view.
    const matching = matchOringSyncRecords(kind, records.map(r => r.fields), items.map(i => i.fields));
    for (const index of matching.ambiguousExcel) add("duplicates", records[index].fields, "Abbinamento Excel/SharePoint ambiguo: nessuna scrittura", `Riga Excel ${records[index].row}`);
    for (const index of matching.ambiguousSharePoint) add("sharepoint-only", items[index].fields, "Record SharePoint ambiguo: nessuna scrittura", `ID ${items[index].id}`);
    let consecutiveFailures = 0;
    const writeFailed = (fields: OringFields, error: unknown, reference?: string) => {
      add("skipped", fields, error instanceof Error ? error.message : "Errore scrittura", reference);
      if (++consecutiveFailures >= MAX_CONSECUTIVE_SYNC_WRITE_FAILURES) throw new Error(`Sincronizzazione interrotta dopo ${consecutiveFailures} errori consecutivi`);
    };
    if (direction === "excel-to-totem") {
      for (let index = 0; index < records.length; index++) {
        const record = records[index], target = matching.matches.get(index);
        if (matching.ambiguousExcel.includes(index)) continue;
        if (target === undefined && parsed.invalid.length) {
          add("skipped", record.fields, "Creazione sospesa: correggere prima le righe Excel non valide"); continue;
        }
        try {
          if (target === undefined) {
            await service.createItem(listId, record.fields);
            add("created", record.fields, "Creato da Excel", `Riga Excel ${record.row}`);
          } else {
            const current = items[target];
            const patch = Object.fromEntries(Object.entries(record.fields).filter(([field, value]) => {
              const column = byName.get(field);
              return column && !sameOringSyncValue(current.fields[field], value, column);
            }));
            if (Object.keys(patch).length) {
              await service.updateItem(listId, current.id, patch);
              add("updated", record.fields, `Aggiornati ${Object.keys(patch).length} campi da Excel`, `Riga Excel ${record.row}`);
            } else add("unchanged", record.fields, "Già allineato a Excel", `Riga Excel ${record.row}`);
          }
          consecutiveFailures = 0;
        } catch (error) { writeFailed(record.fields, error, `Riga Excel ${record.row}`); }
        onProgress?.(`${label} ${kind}: ${index + 1}/${records.length} righe esaminate...`);
      }
      for (const index of matching.newSharePoint) add("sharepoint-only", items[index].fields, "Presente solo nel Totem: conservato senza eliminazioni", `ID ${items[index].id}`);
    } else {
      if (!items.length) throw new Error("Lista SharePoint vuota: Excel non modificato");
      const inverse = new Map([...matching.matches].map(([excelIndex, spIndex]) => [spIndex, excelIndex]));
      const plans: RowPlan[] = [];
      const occupied = new Set(records.map(record => record.row));
      let nextRow = range.rowIndex + range.values.length + 1;
      let reservationHeader: { address: string; value: string } | undefined;
      if (kind === "ORING-EPDM" && columns.field_16 === undefined && items.some(item => !blank(item.fields.field_16))) {
        columns.field_16 = Math.max(20, ...range.values.map(row => row.length));
        mapped.push(["field_16", columns.field_16]);
        reservationHeader = { address: `${excelColumnName(range.columnIndex + columns.field_16)}${range.rowIndex + parsed.header + 1}`, value: "PRENOTAZIONE FORNITORE" };
      }
      const template = records[records.length - 1];
      const templateFormulas = template ? (range.formulas[template.index] || []).flatMap((formula, col) => {
        if (!String(formula).startsWith("=")) return [];
        return [{ col, formula: range.formulasR1C1?.[template.index]?.[col] }];
      }) : [];
      for (let index = 0; index < items.length; index++) {
        if (matching.ambiguousSharePoint.includes(index)) continue;
        const item = items[index], matched = inverse.get(index);
        const source = matched === undefined ? undefined : records[matched];
        try {
          if (!source && parsed.invalid.length) throw new Error("Aggiunta Excel sospesa: esistono righe Excel non valide");
          if (!source && blank(oringCode(item.fields.Title)) && blank(item.fields[customerField(kind)])) throw new Error("Codice e cliente mancanti: impossibile aggiungere la riga Excel");
          let row = source?.row;
          let expandTable = false;
          if (row === undefined) {
            const emptyIndex = range.values.findIndex((values, i) => i > parsed.header && !occupied.has(range.rowIndex + i + 1)
              && values.every((value, col) => blank(value) || (col === columns.field_0 && String(range.formulas[i]?.[col]).startsWith("="))));
            if (emptyIndex >= 0) row = range.rowIndex + emptyIndex + 1;
            else { row = nextRow++; expandTable = kind === "ORING-HNBR" || kind === "ORING-FKM"; }
          }
          occupied.add(row);
          const localIndex = row - range.rowIndex - 1;
          const formulas: RowPlan["formulas"] = [];
          if (!source) for (const formula of templateFormulas) {
            if (typeof formula.formula !== "string" || !formula.formula.startsWith("=")) throw new Error("Formula R1C1 del modello non disponibile: riga non aggiunta");
            formulas.push({ field: mapped.find(([, col]) => col === formula.col)?.[0] || "", address: `${excelColumnName(range.columnIndex + formula.col)}${row}`, formula: formula.formula });
          }
          const writes: CellWrite[] = [];
          for (const [field, col] of mapped) {
            const address = `${excelColumnName(range.columnIndex + col)}${row}`;
            if (String(range.formulas[localIndex]?.[col] || "").startsWith("=") || formulas.some(formula => formula.address === address)) continue;
            const column = byName.get(field)!;
            const value = valueForOringExcel(item.fields[field], column);
            if (!sameOringSyncValue(item.fields[field], range.values[localIndex]?.[col], column, true)) writes.push({ field, address, value, ...(column.dateTime ? { format: "dd/mm/yyyy" } : {}) });
          }
          plans.push({ fields: item.fields, itemId: item.id, row, source, created: !source, expandTable, writes, formulas });
        } catch (error) { add("skipped", item.fields, error instanceof Error ? error.message : "Pianificazione Excel fallita"); }
      }
      // Refuse to append into a formatted table whose actual end differs from
      // the used range: there may be footnotes/data beyond the table.
      if (plans.some(plan => plan.expandTable)) {
        const table = kind === "ORING-HNBR" ? "tblORING" : "tblORING3";
        const body = await service.getWorkbookTableDataBodyRangeByItemId(fileId, table, driveId);
        const end = body.address.match(/:?[A-Z]+(\d+)$/)?.[1];
        if (!end || Number(end) !== range.rowIndex + range.values.length) throw new Error("Fine tabella Excel non coerente: verificare le righe aggiuntive prima di sincronizzare");
      }
      if (reservationHeader && plans.length) await service.updateWorkbookCellByAddress(fileId, sheet, reservationHeader.address, reservationHeader.value, { sessionId: session }, driveId);
      let processed = 0;
      for (const plan of plans) {
        try {
          if (plan.expandTable) await service.appendWorkbookTableRowByItemId(fileId, kind === "ORING-HNBR" ? "tblORING" : "tblORING3", Array(Math.max(...range.values.map(row => row.length))).fill(null), { sessionId: session }, driveId);
          for (const write of plan.writes) await service.updateWorkbookCellByAddress(fileId, sheet, write.address, write.value, { sessionId: session, numberFormat: write.format }, driveId);
          for (const formula of plan.formulas) await service.updateWorkbookFormulaByAddress(fileId, sheet, formula.address, formula.formula, { sessionId: session }, driveId);
          // Verify each record while the session is open. Rows partially
          // written on a failed attempt will be matched on the next retry.
          const after = plan.writes.length || plan.formulas.length || plan.expandTable
            ? await service.readWorkbookWorksheetByItemId(fileId, sheet, { sessionId: session }, driveId)
            : range;
          const actualRow = plan.row - after.rowIndex - 1;
          for (const write of plan.writes) {
            const col = columns[write.field] + range.columnIndex - after.columnIndex;
            if (!sameOringSyncValue(plan.fields[write.field], after.values[actualRow]?.[col], byName.get(write.field)!, true)) throw new Error(`Verifica Excel fallita in ${sheet}!${write.address}`);
          }
          const derived: OringFields = {};
          if (byName.has("RigaExcel")) derived.RigaExcel = plan.row;
          for (const [field, col] of mapped) {
            const actualCol = col + range.columnIndex - after.columnIndex;
            const formula = after.formulas?.[actualRow]?.[actualCol];
            if (!String(formula || "").startsWith("=")) continue;
            const actual = valueForSharePoint(after.values[actualRow]?.[actualCol], byName.get(field)!);
            if (!sameOringSyncValue(plan.fields[field], actual, byName.get(field)!)) derived[field] = actual;
          }
          if (Object.keys(derived).some(field => !sameOringSyncValue(plan.fields[field], derived[field], byName.get(field)!))) await service.updateItem(listId, plan.itemId, derived);
          const section = plan.created ? "created" : plan.writes.length ? "updated" : "unchanged";
          add(section, plan.fields, plan.created ? "Riga aggiunta in Excel" : plan.writes.length ? `${plan.writes.length} celle aggiornate e verificate` : "Già allineato; formule conservate", `Riga Excel ${plan.row}`);
          consecutiveFailures = 0;
        } catch (error) { writeFailed(plan.fields, error, `Riga Excel ${plan.row}`); }
        onProgress?.(`${label} ${kind}: ${++processed}/${plans.length} righe, ${detailMap.updated.length} aggiornate, ${detailMap.created.length} create...`);
      }
      for (const index of matching.newExcel) add("skipped", records[index].fields, "Presente solo in Excel: conservato senza eliminazioni", `Riga Excel ${records[index].row}`);
    }
  } catch (error) { failure = error instanceof Error ? error.message : "Errore sincronizzazione ORING"; }
  finally {
    if (session && driveId) try { await service.closeWorkbookSessionByItemId(fileId, session, driveId); }
    catch (error) { failure = `${failure ? `${failure}. ` : ""}Chiusura sessione Excel non confermata: ${error instanceof Error ? error.message : "errore"}`; }
    clearCacheKeys([kind.toLowerCase(), `admin-${kind.toLowerCase()}`]);
  }
  const issues = detailMap.skipped.length + detailMap.duplicates.length + detailMap["sharepoint-only"].length;
  const success = !failure && issues === 0;
  const labels: Record<SyncDetailKey, string> = { updated: "Aggiornati", created: "Creati", unchanged: "Invariati", skipped: "Saltati / errori", duplicates: "Abbinamenti ambigui", "sharepoint-only": "Solo Totem / ambigui" };
  return {
    success,
    message: `${kind} ${label}: ${failure ? `errore — ${failure}` : success ? "completato" : "parziale: verificare il dettaglio"}. Aggiornati ${detailMap.updated.length}, creati ${detailMap.created.length}, invariati ${detailMap.unchanged.length}, da verificare ${issues}.`,
    details: (Object.keys(detailMap) as SyncDetailKey[]).filter(key => detailMap[key].length).map(key => ({ key, label: labels[key], items: detailMap[key] })),
  };
}
