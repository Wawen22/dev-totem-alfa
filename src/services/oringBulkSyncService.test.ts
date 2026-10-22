import { beforeEach, describe, expect, it, vi } from "vitest";
import { ORING_KINDS, OringKind } from "../config/oringSync";
import { SharePointColumnDefinition, SharePointService, WorkbookWorksheetRange } from "./sharePointService";
import { syncOringWarehouse } from "./oringBulkSyncService";
import { matchOringSyncRecords, oringSyncLayout, sameOringSyncValue, valueForOringExcel, valueForSharePoint } from "./oringSyncMapping";

vi.mock("../hooks/useCachedList", () => ({ clearCacheKeys: vi.fn() }));
const positionedHeader = ["Posizione", "Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "FLUIDI", "NR Ordine", "Data Ordine", "Q.TA'", "Commessa ", "Design temperature", "Prezzo unitario", "Giacenza", "PRENOTAZIONE FORNITORE", "DATA PR./AGG."];
const nbrHeader = ["Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola ", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "Fluidi", "", "", "Ordine", "Data", "Q.TA'", "Commessa ", "Design temperature", "Prezzo unitario", "Q.tà\nminima", "PRENOTAZIONE", "GIACENZA", "DATA PR./AGG."];
const epdmHeader = ["Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola ", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "Fluido", "Ordine", "Data", "Q.TA'", "STANDARD", "Design temperature", "Prezzo unitario", "GIACENZA", "", "DATA", ""];

function fixture(kind: OringKind = "ORING-EPDM") {
  const material = kind.slice(6), positioned = kind === "ORING-HNBR" || kind === "ORING-FKM", nbr = kind === "ORING-NBR";
  const header = positioned ? positionedHeader : nbr ? nbrHeader : epdmHeader;
  const headIndex = positioned ? 6 : 7;
  const row = positioned
    ? [kind === "ORING-HNBR" ? "A0" : 1, material + "100", "SAM100", 100, "NERO", 16, 20, 2, material, "SH90", "GAS", "OD1", "17/12/24", 50, "STD1", "-40 / 150", 3.5, 12, "2", 46091]
    : nbr ? ["NBR100", "SAM100", 100, "NERO", 16, 20, 2, "NBR", "SH90", "GAS", "OLIO", "COLD WATER", "OD1", "17/12/24", 50, "STD1", "-30 / 110", 3.5, 100, "2", 12, 46091]
      : ["EPDM100", "SAM100", 100, "NERO", 16, 20, 2, "EPDM", "SH90", "ACQUA", "OD1", "17/12/24", 50, "STD1", "-30 / 110", 3.5, 12, 3.53, 46091, "nota"];
  const values: unknown[][] = [...Array.from({ length: headIndex }, () => []), [...header], row];
  const formulas = values.map(row => [...row]);
  const formulasR1C1 = values.map(row => [...row]);
  const intColumn = positioned ? 5 : 4;
  formulas[headIndex + 1][intColumn] = positioned ? "=G8-2*H8" : "=F9-2*G9";
  formulasR1C1[headIndex + 1][intColumn] = "=RC[1]-2*RC[2]";
  if (kind === "ORING-FKM") { formulas[headIndex + 1][0] = "=ROW()-7"; formulasR1C1[headIndex + 1][0] = "=ROW()-7"; }
  const range: WorkbookWorksheetRange = { address: `${material}!A1:${nbr ? "V" : "T"}${values.length}`, rowIndex: 0, columnIndex: 0, values, formulas, formulasR1C1 };
  const layout = oringSyncLayout(kind, range);
  const numberFields = nbr ? ["field_3", "field_4", "field_5", "field_13", "field_14", "field_16"] : ["field_4", "field_5", "field_6", "field_14", "field_15", ...(kind === "ORING-HNBR" ? [] : ["field_16", "QuantitaOrdinata"])];
  const schema: SharePointColumnDefinition[] = Object.keys(layout.columns)
    .filter(field => (kind !== "ORING-HNBR" && !nbr) || !["CodiceSAM", "QuantitaOrdinata"].includes(field))
    .map(name => ({ name, ...(numberFields.includes(name) ? { number: {} } : name === "field_17" || (positioned && name === "field_19") ? { dateTime: {} } : { text: {} }) }));
  if (kind === "ORING-FKM" || kind === "ORING-EPDM") schema.push({ name: "RigaExcel", number: {} });
  if (kind === "ORING-EPDM") schema.push({ name: "field_16", number: {} });
  const fields = Object.fromEntries(schema.filter(col => layout.columns[col.name] !== undefined).map(col => [col.name, valueForSharePoint(row[layout.columns[col.name]], col)]));
  if (schema.some(col => col.name === "RigaExcel")) fields.RigaExcel = values.length;
  const items = [{ id: "1", fields }];
  const cell = (address: string) => { const match = address.match(/^([A-Z]+)(\d+)$/)!; return { row: Number(match[2]) - 1, col: [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1 }; };
  const recalculate = () => {
    for (let r = headIndex + 1; r < values.length; r++) for (let c = 0; c < (formulasR1C1[r] || []).length; c++) {
      const f = formulasR1C1[r][c];
      if (f === "=RC[1]-2*RC[2]") values[r][c] = Number(values[r][c + 1] || 0) - 2 * Number(values[r][c + 2] || 0);
      if (f === "=ROW()-7") values[r][c] = r + 1 - 7;
    }
  };
  const service = {
    getDriveIdByName: vi.fn().mockResolvedValue("drive"), getDriveItemByPath: vi.fn().mockResolvedValue({ id: "file" }),
    createWorkbookSessionByItemId: vi.fn().mockResolvedValue("session"), closeWorkbookSessionByItemId: vi.fn().mockResolvedValue(undefined),
    readWorkbookWorksheetByItemId: vi.fn(async (_file?: string, _sheet?: string) => { recalculate(); return structuredClone(range); }),
    listItems: vi.fn(async () => structuredClone(items)), listColumns: vi.fn(async () => schema),
    updateItem: vi.fn(async (_list, id, patch) => { Object.assign(items.find(item => item.id === id)!.fields, patch); }),
    createItem: vi.fn(async (_list, fields) => { const item = { id: String(items.length + 1), fields }; items.push(item); return item; }),
    updateWorkbookCellByAddress: vi.fn(async (_file, _sheet, address, value) => {
      const { row, col } = cell(address); values[row] ||= []; formulas[row] ||= []; formulasR1C1[row] ||= [];
      values[row][col] = value; formulas[row][col] = value; formulasR1C1[row][col] = value;
    }),
    updateWorkbookFormulaByAddress: vi.fn(async (_file, _sheet, address, formula) => {
      const { row, col } = cell(address); formulas[row] ||= []; formulasR1C1[row] ||= [];
      formulas[row][col] = formula; formulasR1C1[row][col] = formula; recalculate();
    }),
    appendWorkbookTableRowByItemId: vi.fn(async () => { values.push(Array(header.length).fill(null)); formulas.push([]); formulasR1C1.push([]); }),
    getWorkbookTableDataBodyRangeByItemId: vi.fn(async () => ({ address: `${material}!A${headIndex + 2}:${nbr ? "V" : "T"}${values.length}`, rowCount: values.length - headIndex - 1 })),
    deleteItem: vi.fn(),
  };
  return { kind, service, typed: service as unknown as SharePointService, range, items, layout, schema, headIndex };
}

beforeEach(() => { vi.stubEnv("VITE_ORING_EXCEL_DRIVE_ID", ""); });

describe.each(ORING_KINDS)("sincronizzazione %s", kind => {
  it("Excel → Totem aggiorna i campi della lista corretta, incluse celle vuote", async () => {
    const f = fixture(kind), stock = kind === "ORING-NBR" ? "field_16" : "field_15";
    f.items[0].fields[stock] = 99;
    const result = await syncOringWarehouse(f.typed, kind, "list", "excel-to-totem");
    expect(result.success).toBe(true);
    expect(f.items[0].fields[stock]).toBe(12);
    expect(f.service.updateItem).toHaveBeenCalledWith("list", "1", expect.objectContaining({ [stock]: 12 }));
    expect(f.service.updateWorkbookCellByAddress).not.toHaveBeenCalled();
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });

  it("Totem → Excel scrive la giacenza nella colonna specifica preservando le formule", async () => {
    const f = fixture(kind), stock = kind === "ORING-NBR" ? "field_16" : "field_15";
    f.items[0].fields[stock] = 0;
    const before = structuredClone(f.range.formulas);
    const result = await syncOringWarehouse(f.typed, kind, "list", "totem-to-excel");
    expect(result.success).toBe(true);
    const address = kind === "ORING-NBR" ? "U9" : kind === "ORING-EPDM" ? "Q9" : "R8";
    expect(f.service.updateWorkbookCellByAddress).toHaveBeenCalledWith("file", kind === "ORING-NBR" ? "NBR " : kind.slice(6), address, 0, expect.any(Object), "drive");
    for (let r = 0; r < before.length; r++) for (let c = 0; c < before[r].length; c++) if (String(before[r][c]).startsWith("=")) expect(f.range.formulas[r][c]).toBe(before[r][c]);
    expect(f.service.deleteItem).not.toHaveBeenCalled();
  });

  it("Excel → Totem crea i record mancanti e non li duplica al secondo passaggio", async () => {
    const f = fixture(kind); f.items.length = 0;
    expect((await syncOringWarehouse(f.typed, kind, "list", "excel-to-totem")).success).toBe(true);
    expect(f.service.createItem).toHaveBeenCalledTimes(1);
    expect((await syncOringWarehouse(f.typed, kind, "list", "excel-to-totem")).success).toBe(true);
    expect(f.service.createItem).toHaveBeenCalledTimes(1);
  });

  it("Totem → Excel aggiunge i record mancanti, copia le formule e resta idempotente", async () => {
    const f = fixture(kind), customer = kind === "ORING-NBR" ? "field_1" : "field_2";
    f.items.push({ id: "2", fields: { ...f.items[0].fields, Title: kind.slice(6) + "200", [customer]: "200", field_0: kind === "ORING-HNBR" ? "A1" : "2", RigaExcel: null } });
    const first = await syncOringWarehouse(f.typed, kind, "list", "totem-to-excel");
    expect(first.success, first.message).toBe(true);
    expect(first.details?.find(section => section.key === "created")?.items).toHaveLength(1);
    expect(f.service.updateWorkbookFormulaByAddress).toHaveBeenCalled();
    const length = f.range.values.length;
    const second = await syncOringWarehouse(f.typed, kind, "list", "totem-to-excel");
    expect(second.success, second.message).toBe(true);
    expect(f.range.values).toHaveLength(length);
    expect(second.details?.find(section => section.key === "created")).toBeUndefined();
  });
});

describe("controlli sincronizzazione ORING", () => {
  it("rispetta lo spazio finale NBR e mantiene le tre colonne Fluidi separate", async () => {
    const f = fixture("ORING-NBR"); f.items.length = 0;
    await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(f.service.readWorkbookWorksheetByItemId.mock.calls[0][1]).toBe("NBR ");
    expect(f.items[0].fields).toMatchObject({ field_8: "GAS", field_9: "OLIO", field_10: "COLD WATER", field_14: 100 });
    expect(f.items[0].fields).not.toHaveProperty("RigaExcel");
    expect(f.items[0].fields).not.toHaveProperty("CodiceSAM");
  });

  it("mantiene il giorno italiano delle date SharePoint anche durante l'ora legale", () => {
    const column = { name: "field_19", dateTime: {} };
    expect(sameOringSyncValue("2024-12-16T23:00:00Z", "17/12/24", column, true)).toBe(true);
    expect(sameOringSyncValue("2025-03-30T22:00:00Z", "31/03/25", column, true)).toBe(true);
    expect(valueForOringExcel("2024-12-16T23:00:00Z", column)).toBe(Date.parse("2024-12-17") / 86400000 + 25569);
  });

  it("distinte posizioni HNBR/FKM restano record distinti dello stesso codice", () => {
    for (const kind of ["ORING-HNBR", "ORING-FKM"] as const) {
      const rows = [{ Title: "SAME", field_2: "100", field_0: "1", field_12: "STD" }, { Title: "SAME", field_2: "100", field_0: "2", field_12: null }];
      const result = matchOringSyncRecords(kind, rows, [...rows].reverse());
      expect([...result.matches]).toEqual([[0, 1], [1, 0]]);
    }
  });

  it("le righe duplicate ambigue non diventano nuovi elementi né sovrascrivono il primo", async () => {
    const f = fixture(); f.items.push({ id: "2", fields: { ...f.items[0].fields } });
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(result.success).toBe(false);
    expect(f.service.updateItem).not.toHaveBeenCalled();
    expect(f.service.createItem).not.toHaveBeenCalled();
  });

  it("non divide automaticamente un vecchio elemento HNBR con posizione 23-26", () => {
    const excel = [{ Title: "HNBR61400", field_2: "61400", field_0: "23", field_12: "STD" }, { Title: "HNBR61400", field_2: "61400", field_0: "26", field_12: "" }];
    const result = matchOringSyncRecords("ORING-HNBR", excel, [{ Title: "HNBR61400", field_2: "61400", field_0: "23-26", field_12: "COMMESSE" }]);
    expect(result.matches.size).toBe(0);
    expect(result.newExcel).toEqual([]);
    expect(result.ambiguousExcel).toEqual([0, 1]);
  });

  it("riconosce un cliente corretto in Excel se codice e posizione coincidono", () => {
    const result = matchOringSyncRecords("ORING-HNBR", [{ Title: "HNBR3675", field_2: "37", field_0: "A3" }], [{ Title: "HNBR3675", field_2: "3675", field_0: "A3" }]);
    expect(result.matches.get(0)).toBe(0);
  });

  it("preserva le righe solo SharePoint durante Excel → Totem", async () => {
    const f = fixture(); f.items.push({ id: "2", fields: { Title: "EXTRA", field_2: "200" } });
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(result.success).toBe(false);
    expect(result.details?.find(section => section.key === "sharepoint-only")?.items).toHaveLength(1);
    expect(f.service.deleteItem).not.toHaveBeenCalled();
  });

  it("non converte le etichette EPDM senza codice in codici Excel", async () => {
    const f = fixture(); f.range.values[8][0] = ""; f.items[0].fields.Title = "Senza codice (cliente 100)";
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "totem-to-excel");
    expect(result.success).toBe(true);
    expect(f.range.values[8][0]).toBe("");
  });

  it("non sovrascrive SAM/Ordine/Quantità NBR che non esistono nella lista storica", async () => {
    const f = fixture("ORING-NBR"); f.items[0].fields.field_16 = 99;
    const before = f.range.values[8].slice();
    await syncOringWarehouse(f.typed, f.kind, "list", "totem-to-excel");
    for (const col of [1, 12, 13, 14]) expect(f.range.values[8][col]).toBe(before[col]);
  });

  it("non annulla prenotazioni EPDM quando il vecchio foglio non ha ancora la colonna", async () => {
    const f = fixture(); f.items[0].fields.field_16 = 5;
    await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(f.items[0].fields.field_16).toBe(5);
    await syncOringWarehouse(f.typed, f.kind, "list", "totem-to-excel");
    expect(f.range.values[7][20]).toBe("PRENOTAZIONE FORNITORE");
    expect(f.range.values[8][20]).toBe(5);
    expect(f.range.values[8][17]).toBe(3.53);
  });

  it("rifiuta intestazioni o schema mancanti prima di scrivere", async () => {
    const f = fixture(); f.range.values[7][16] = "RINOMINATA";
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(result.success).toBe(false);
    expect(f.service.updateItem).not.toHaveBeenCalled();
    expect(f.service.createItem).not.toHaveBeenCalled();
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });

  it("un errore di scrittura o verifica non viene presentato come successo", async () => {
    const f = fixture(); f.items[0].fields.field_15 = 99;
    f.service.updateWorkbookCellByAddress.mockResolvedValue(undefined);
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "totem-to-excel");
    expect(result.success).toBe(false);
    expect(result.details?.find(section => section.key === "skipped")?.items[0].detail).toContain("Verifica Excel fallita");
  });

  it("un errore di chiusura sessione mantiene l'esito di errore", async () => {
    const f = fixture(); f.service.closeWorkbookSessionByItemId.mockRejectedValue(new Error("SESSION"));
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(result.success).toBe(false); expect(result.message).toContain("Chiusura sessione");
  });

  it("interrompe le scritture dopo tre errori consecutivi mantenendo il log", async () => {
    const f = fixture(); f.items.length = 0;
    for (let index = 1; index < 4; index++) {
      const row = [...f.range.values[8]]; row[0] = "EPDM" + (index + 100); row[2] = index + 100;
      f.range.values.push(row);
    }
    f.service.createItem.mockRejectedValue(new Error("Graph non disponibile"));
    const result = await syncOringWarehouse(f.typed, f.kind, "list", "excel-to-totem");
    expect(result.success).toBe(false);
    expect(result.message).toContain("interrotta dopo 3 errori");
    expect(f.service.createItem).toHaveBeenCalledTimes(3);
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });
});
