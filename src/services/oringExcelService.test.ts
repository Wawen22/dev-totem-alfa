import { beforeEach, describe, expect, it, vi } from "vitest";
import { matchOringWorksheetRow, saveOringWithExcel } from "./oringExcelService";
import { SharePointService, WorkbookWorksheetRange } from "./sharePointService";

const headers = {
  FKM: ["Posizione", "Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "FLUIDI", "NR Ordine", "Data Ordine", "Q.TA'", "Commessa ", "Design temperature", "Prezzo unitario", "Giacenza", "PRENOTAZIONE FORNITORE", "DATA PR./AGG."],
  EPDM: ["Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola ", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "Fluido", "Ordine", "Data", "Q.TA'", "STANDARD", "Design temperature", "Prezzo unitario", "GIACENZA", "", "DATA", ""],
};

function fixture(material: "FKM" | "EPDM" = "EPDM") {
  const fkm = material === "FKM";
  const headerRow = fkm ? 6 : 7;
  const record = fkm
    ? [15, "FKM6775", "SAM1", 6775, "ROSSO", 196.34, 207, 5.33, "FKM", "SH90", "GAS", "OD1", "", "", "STD", "-40 / 220", 8.62, 60, "", 46294]
    : ["EPDM2287", "SAM1", 2287, "BLU", 72.75, 76.31, 1.78, "EPDM", "SH90", "ACQUA", "OD1", "", "", "STD", "-30 / 110", 10.58, 20, 3.53, 46300, "nota"];
  const values = [...Array.from({ length: headerRow }, () => []), [...headers[material]], record];
  const formulas = values.map(row => [...row]);
  formulas[headerRow + 1][fkm ? 5 : 4] = fkm ? "=G8-2*H8" : "=F9-2*G9";
  if (fkm) formulas[headerRow + 1][0] = "=ROW()-7";
  const range: WorkbookWorksheetRange = { address: `${material}!A1:T${values.length}`, rowIndex: 0, columnIndex: 0, values, formulas };
  const original = { Title: fkm ? "FKM6775" : "EPDM2287", CodiceSAM: "SAM1", field_2: fkm ? "6775" : "2287", field_0: fkm ? "15" : undefined, field_3: fkm ? "ROSSO" : "BLU", field_4: fkm ? 196.34 : 72.75, field_5: fkm ? 207 : 76.31, field_6: fkm ? 5.33 : 1.78, field_15: fkm ? 60 : 20, RigaExcel: values.length };
  const calls: string[] = [];
  const service = {
    getDriveIdByName: vi.fn().mockResolvedValue("drive"), getDriveItemByPath: vi.fn().mockResolvedValue({ id: "file" }),
    createWorkbookSessionByItemId: vi.fn().mockResolvedValue("session"),
    closeWorkbookSessionByItemId: vi.fn().mockResolvedValue(undefined),
    readWorkbookWorksheetByItemId: vi.fn(async () => structuredClone(range)),
    updateItem: vi.fn(async () => { calls.push("SharePoint"); }),
    updateWorkbookCellByAddress: vi.fn(async (_file, _sheet, address: string, value, _options?: { numberFormat?: string }) => {
      calls.push(address);
      const [, letters, row] = address.match(/^([A-Z]+)(\d+)$/)!;
      const column = [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      values[Number(row) - 1][column] = value;
      formulas[Number(row) - 1][column] = value;
      if (column === (fkm ? 6 : 5)) values[headerRow + 1][fkm ? 5 : 4] = Number(value) - 2 * Number(values[headerRow + 1][fkm ? 7 : 6]);
    }),
  };
  return { service, typed: service as unknown as SharePointService, range, original, calls };
}

beforeEach(() => { vi.stubEnv("VITE_ORING_EXCEL_DRIVE_ID", ""); });

describe("scrittura puntuale ORING su Excel", () => {
  it.each(["FKM", "EPDM"] as const)("salva colore %s nella cella corretta preservando le formule", async material => {
    const f = fixture(material);
    const before = structuredClone(f.range.formulas);
    await saveOringWithExcel(f.typed, `ORING-${material}`, "list", "1", f.original, { ...f.original, field_3: "BLU2" });
    expect(f.service.updateWorkbookCellByAddress).toHaveBeenCalledTimes(1);
    expect(f.calls).toContain(material === "FKM" ? "E8" : "D9");
    for (let r = 0; r < before.length; r++) for (let c = 0; c < before[r].length; c++) {
      if (String(before[r][c]).startsWith("=")) expect(f.range.formulas[r][c]).toBe(before[r][c]);
    }
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalledWith("file", "session", "drive");
  });

  it("cambia giacenza senza sovrascrivere un colore modificato direttamente in Excel", async () => {
    const f = fixture(); f.range.values[8][3] = "COLORE RECENTE";
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_15: 0 });
    expect(f.range.values[8][3]).toBe("COLORE RECENTE");
    expect(f.service.updateWorkbookCellByAddress.mock.calls[0].slice(1, 4)).toEqual(["EPDM", "Q9", 0]);
  });

  it("svuota le celle con stringa vuota e scrive date Excel numeriche", async () => {
    const f = fixture();
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_3: null, field_19: "2026-10-22T00:00:00.000Z" });
    expect(f.range.values[8][3]).toBe("");
    const dateCall = f.service.updateWorkbookCellByAddress.mock.calls.find(c => c[2] === "L9")!;
    expect(dateCall[3]).toBe(Date.parse("2026-10-22") / 86400000 + 25569);
    expect(dateCall[4]).toMatchObject({ numberFormat: "dd/mm/yyyy" });
  });

  it("aggiorna i valori calcolati SharePoint dopo una modifica del diametro esterno", async () => {
    const f = fixture();
    const result = await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_5: 80 });
    expect(result.field_4).toBeCloseTo(80 - 2 * 1.78);
    expect(f.range.formulas[8][4]).toBe("=F9-2*G9");
    expect(f.service.updateWorkbookCellByAddress).toHaveBeenCalledTimes(1);
  });

  it("rifiuta la modifica diretta di una formula prima di scrivere SharePoint", async () => {
    const f = fixture();
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_4: 99 })).rejects.toThrow("contiene una formula");
    expect(f.service.updateItem).not.toHaveBeenCalled();
    expect(f.service.updateWorkbookCellByAddress).not.toHaveBeenCalled();
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });

  it("usa l'identità originale quando si rinomina il codice", async () => {
    const f = fixture();
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { Title: "EPDM-NUOVO" });
    expect(f.range.values[8][0]).toBe("EPDM-NUOVO");
  });

  it("ritrova una riga spostata e non usa ciecamente RigaExcel", () => {
    const f = fixture(); f.range.values.splice(8, 0, []);
    expect(matchOringWorksheetRow("ORING-EPDM", f.range, f.original)).toBe(9);
  });

  it("distingue codici FKM ripetuti tramite la commessa originale, rifiutando ambiguità", () => {
    const f = fixture("FKM"); const other = [...f.range.values[7]]; other[0] = 26; other[14] = "ALTRO";
    f.range.values.push(other);
    expect(matchOringWorksheetRow("ORING-FKM", f.range, { ...f.original, field_12: "ALTRO" })).toBe(8);
    expect(() => matchOringWorksheetRow("ORING-FKM", f.range, { ...f.original, field_0: undefined })).toThrow("ambigua");
  });

  it("riconosce i record EPDM senza codice Alfa tramite cliente e SAM", async () => {
    const f = fixture(); f.range.values[8][0] = ""; f.range.values[8][1] = "";
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", { ...f.original, Title: "Senza codice (cliente 2287)", CodiceSAM: "" }, { field_15: 7 });
    expect(f.range.values[8][0]).toBe("");
    expect(f.range.values[8][16]).toBe(7);
  });

  it("conserva R/T di EPDM e aggiunge una colonna dedicata alla prenotazione", async () => {
    const f = fixture();
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_16: 3 });
    expect(f.range.values[7][20]).toBe("PRENOTAZIONE FORNITORE");
    expect(f.range.values[8][20]).toBe(3);
    expect(f.range.values[8][17]).toBe(3.53);
    expect(f.range.values[8][19]).toBe("nota");
  });

  it("se Excel fallisce segnala il salvataggio parziale e permette di riprovare", async () => {
    const f = fixture(); f.service.updateWorkbookCellByAddress.mockRejectedValueOnce(new Error("Excel bloccato"));
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_3: "BLU2" })).rejects.toMatchObject({ sharePointSaved: true, message: expect.stringContaining("SharePoint aggiornato") });
    await saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", { ...f.original, field_3: "BLU2" }, { field_3: "BLU2" });
    expect(f.range.values[8][3]).toBe("BLU2");
  });

  it("verifica il valore riletto e non considera riuscito un PATCH ignorato", async () => {
    const f = fixture(); f.service.updateWorkbookCellByAddress.mockResolvedValue(undefined);
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_15: 99 })).rejects.toThrow("Verifica Excel fallita");
  });

  it("un errore di chiusura non nasconde il primo errore di validazione", async () => {
    const f = fixture(); f.service.closeWorkbookSessionByItemId.mockRejectedValue(new Error("close"));
    const warning = vi.spyOn(console, "warn").mockImplementation(() => {});
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_4: 99 })).rejects.toThrow("contiene una formula");
    warning.mockRestore();
  });

  it("chiude le sessioni anche dopo un errore nella rilettura e segnala il fallimento", async () => {
    const f = fixture(); f.service.readWorkbookWorksheetByItemId.mockRejectedValueOnce(new Error("lettura negata"));
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_3: "BLU2" })).rejects.toMatchObject({ sharePointSaved: false });
    expect(f.service.updateItem).not.toHaveBeenCalled();
    expect(f.service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });

  it("se la sessione non si chiude non mostra un successo non confermato", async () => {
    const f = fixture(); f.service.closeWorkbookSessionByItemId.mockRejectedValue(new Error("close"));
    await expect(saveOringWithExcel(f.typed, "ORING-EPDM", "list", "1", f.original, { field_3: "BLU2" })).rejects.toMatchObject({ sharePointSaved: true, message: expect.stringContaining("chiusura") });
  });

  it("usa gli endpoint worksheet e PATCH cella con la stessa sessione Graph", async () => {
    const response = fixture().range;
    const request = { header: vi.fn().mockReturnThis(), get: vi.fn().mockResolvedValue(response), patch: vi.fn().mockResolvedValue(undefined) };
    const api = vi.fn(() => request);
    const service = new SharePointService(async () => ({ api }) as any, "site");
    expect(await service.readWorkbookWorksheetByItemId("file", "EPDM", { sessionId: "session" }, "drive")).toBe(response);
    expect(api).toHaveBeenCalledWith("/sites/site/drives/drive/items/file/workbook/worksheets/EPDM/usedRange(valuesOnly=true)");
    await service.updateWorkbookCellByAddress("file", "EPDM", "D9", "BLU2", { sessionId: "session" }, "drive");
    expect(api).toHaveBeenCalledWith("/sites/site/drives/drive/items/file/workbook/worksheets/EPDM/range(address='D9')");
    expect(request.header).toHaveBeenCalledWith("workbook-session-id", "session");
    expect(request.patch).toHaveBeenCalledWith({ values: [["BLU2"]] });
  });
});
