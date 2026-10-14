import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import type { SharePointService } from "./services/sharePointService";
import { NewLotFields } from "./components/features/NewLotFields";

vi.mock("./auth/authConfig", () => ({ msalConfig: {}, loginRequest: {} }));
vi.mock("@azure/msal-browser", async (importOriginal) => ({
  ...await importOriginal<typeof import("@azure/msal-browser")>(),
  PublicClientApplication: class {
    initialize() { return Promise.resolve(); }
    getAllAccounts() { return []; }
    addEventCallback() {}
  },
}));
vi.mock("./hooks/useAuthenticatedGraphClient", () => ({ useAuthenticatedGraphClient: () => vi.fn() }));
vi.mock("./hooks/useCachedList", () => ({
  clearCacheKeys: vi.fn(),
  useCachedList: () => ({ data: [
    { id: "1", fields: { Title: "FLCA001", IdentLotto: "A", GiacenzaMm: 24 } },
    { id: "2", fields: { Title: "FLCA001", IdentLotto: "B", GiacenzaMm: 12 } },
    { id: "3", fields: { Title: "FLCA002", IdentLotto: "A", Archiviato: true } },
  ], loading: false, error: null, refresh: vi.fn() }),
}));

let app: typeof import("./App");
beforeAll(async () => { app = await import("./App"); });
beforeEach(() => {
  vi.stubEnv("VITE_SHAREPOINT_SITE_ID", "site");
  vi.stubEnv("VITE_FLANGE_LIST_ID", "flange-list");
  vi.stubEnv("VITE_FLANGE_EXCEL_PATH", "flange.xlsx");
  vi.stubEnv("VITE_FLANGE_EXCEL_TABLE", "tblFlange");
  vi.stubEnv("VITE_FLANGE_EXCEL_DRIVE_ID", "drive");
});

describe("interfaccia flange", () => {
  it("presenta campi nuovi vuoti, incluso Schedula, con i dati compilabili manualmente", () => {
    const html = renderToStaticMarkup(createElement(NewLotFields, { kind: "TUBI", values: {}, onChange: vi.fn() }));
    expect(html).toContain("Schedula");
    expect(html).toContain("FORNITORE");
    expect(html).toContain("Prezzo kg/mt");
    expect(html).not.toMatch(/value="[^"]+"/);
    expect(html).not.toContain("Data/ora modifica");
  });
  it("raggruppa per codice e offre selezione e gestione lotti con lo stile degli altri magazzini", () => {
    const html = renderToStaticMarkup(createElement(app.FlangePanel, { selectedItems: {}, onToggle: vi.fn(), selectionLimitReached: false }));
    expect(html).toContain('class="search-box"');
    expect(html).toContain('type="checkbox"');
    expect(html).toContain("Gestisci lotti FLCA001");
    expect(html).toContain("Lotti (2)");
    expect(html).toContain("GIACENZA (Numero Pezzi)");
    expect(html).not.toContain("Prezzo €/Kg");
    expect(html).not.toContain("FLCA002");
  });
  it("usa giacenza e prezzo della flangia nella pagina di aggiornamento", () => {
    const html = renderToStaticMarkup(createElement(app.StockUpdatePage, {
      items: [{ key: "FLANGE-1", source: "FLANGE", itemId: "1", title: "FLCA001", lottoProg: "A", fields: { GiacenzaMm: 24, PrezzoCad: 289, PrezzoEuroKg: 99, Commessa: "AF4870" } }],
      values: {}, onChange: vi.fn(), onSave: vi.fn(), saving: false, status: "idle", message: null, onBackToInventory: vi.fn(),
    }));
    expect(html).toContain("Giacenza (Numero Pezzi)");
    expect(html).toContain('value="24"');
    expect(html).toContain("Prezzo CAD");
    expect(html).toContain('value="289"');
    expect(html).not.toContain("Giacenza Tubo Intero");
    expect(html).not.toContain("Prezzo €/Kg");
  });
  it("mostra Schedula tra i campi modificabili dell'aggiornamento Tubi", () => {
    const html = renderToStaticMarkup(createElement(app.StockUpdatePage, {
      items: [{ key: "TUBI-1", source: "TUBI", itemId: "1", title: "TCAB105", fields: { Schedula: "SCH 40" } }],
      values: {}, onChange: vi.fn(), onSave: vi.fn(), saving: false, status: "idle", message: null, onBackToInventory: vi.fn(),
    }));
    expect(html).toContain("Schedula");
    expect(html).toContain('value="SCH 40"');
  });
});

function serviceFixture() {
  const service = {
    getDriveItemByPath: vi.fn().mockResolvedValue({ id: "file" }),
    listWorkbookTableColumnsByItemId: vi.fn().mockResolvedValue(["CODICE", "LOTTO", "GIACENZA (mm)", "PREZZO CAD", "Prezzo €/Kg", "COMMESSA"]),
    listWorkbookTableRowsByItemId: vi.fn().mockResolvedValue([{ index: 0, values: [["FLCA001", "A", 24, 289, 99, "OLD"]] }]),
    createWorkbookSessionByItemId: vi.fn().mockResolvedValue("session"),
    closeWorkbookSessionByItemId: vi.fn().mockResolvedValue(undefined),
    appendWorkbookTableRowByItemId: vi.fn().mockResolvedValue(undefined),
    getWorkbookTableDataBodyRangeByItemId: vi.fn().mockResolvedValue({ address: "Flange!A2:F2", rowCount: 1 }),
    updateWorkbookRangeByAddress: vi.fn().mockResolvedValue(undefined),
  };
  return { service, typed: service as unknown as SharePointService };
}

describe("salvataggio Excel flange", () => {
  it("riconosce l'intestazione reale GIACENZA(Nr) in entrambe le direzioni", () => {
    expect(app.buildFlangeExcelRow(["CODICE", "LOTTO", "GIACENZA(Nr)"], { Title: "FLCA044", IdentLotto: "A", GiacenzaMm: 7 })).toEqual(["FLCA044", "A", 7]);
    expect(app.buildFlangePayloadFromExcelRow(["CODICE", "LOTTO", "GIACENZA(Nr)"], ["FLCA044", "A", 7])).toMatchObject({ Title: "FLCA044", GiacenzaMm: 7 });
    expect(app.buildFlangePayloadFromExcelRow(["CODICE", "LOTTO"], ["FLCA048", ""])).toMatchObject({ Title: "FLCA048", IdentLotto: "A" });
  });
  it("aggiorna il lotto selezionato mantenendo prezzo/kg storico e campi non modificati", async () => {
    const { service, typed } = serviceFixture();
    await app.writeFlangeExcel(typed, { Title: "FLCA001", IdentLotto: "A", GiacenzaMm: 7, PrezzoCad: 300, Commessa: "NEW" });
    expect(service.updateWorkbookRangeByAddress.mock.calls[0][3]).toEqual([["FLCA001", "A", 7, 300, 99, "NEW"]]);
    expect(service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });
  it("aggiunge il nuovo lotto senza aggiornare quello preesistente", async () => {
    const { service, typed } = serviceFixture();
    await app.writeFlangeExcel(typed, { Title: "FLCA001", IdentLotto: "B", GiacenzaMm: 0, PrezzoCad: 289 }, true);
    expect(service.appendWorkbookTableRowByItemId.mock.calls[0][2]).toEqual(["FLCA001", "B", 0, 289, "", ""]);
    expect(service.updateWorkbookRangeByAddress).not.toHaveBeenCalled();
  });
  it("rifiuta la modifica se Excel contiene solo un altro lotto", async () => {
    const { service, typed } = serviceFixture();
    await expect(app.writeFlangeExcel(typed, { Title: "FLCA001", IdentLotto: "B", GiacenzaMm: 0 })).rejects.toThrow("Riga Excel non trovata");
    expect(service.updateWorkbookRangeByAddress).not.toHaveBeenCalled();
    expect(service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });
  it("chiude la sessione anche se la scrittura Excel fallisce", async () => {
    const { service, typed } = serviceFixture();
    service.appendWorkbookTableRowByItemId.mockRejectedValue(new Error("Excel bloccato"));
    await expect(app.writeFlangeExcel(typed, { Title: "FLCA001", IdentLotto: "B" }, true)).rejects.toThrow("Excel bloccato");
    expect(service.closeWorkbookSessionByItemId).toHaveBeenCalled();
  });
  it("non crea A se Excel contiene già il lotto storico vuoto", async () => {
    const { service, typed } = serviceFixture();
    service.listWorkbookTableRowsByItemId.mockResolvedValue([{ index: 0, values: [["FLCA001", "", 24, 289, 99, "OLD"]] }]);
    await expect(app.writeFlangeExcel(typed, { Title: "FLCA001", IdentLotto: "A" }, true)).rejects.toThrow("esiste già");
    expect(service.appendWorkbookTableRowByItemId).not.toHaveBeenCalled();
  });
});

describe("sincronizzazione Schedula Tubi", () => {
  it("usa i nomi delle colonne dopo l'inserimento senza spostare SP e GRADO", () => {
    const columns = ["CODICE", "IdentLotto", "SP", "Schedula", "GRADO", "GRADO MATERIALE 2"];
    const values = ["TCAB105", "A", "6.35", "SCH 40", "API 5L X52", "X52"];
    const parsed = app.buildTubiPayloadFromExcelRow(columns, values);
    expect(parsed.fields).toMatchObject({ field_10: "6.35", Schedula: "SCH 40", field_11: "API 5L X52", GRADOMATERIALE2: "X52" });
    expect(app.buildTubiExcelRow(columns, parsed.fields)).toEqual(values);
  });
});
