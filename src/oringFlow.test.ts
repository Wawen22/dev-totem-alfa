import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { oringAdditionalColumns } from "./config/oringAdditionalColumns";
import { additionalOringSchema } from "./utils/oringInitialImport";
import { ORING_ADDITIONAL_FIELDS, AdminPanel } from "./components/features/AdminPanel";

vi.mock("./auth/authConfig", () => ({ msalConfig: {}, loginRequest: {} }));
vi.mock("@azure/msal-browser", async (original) => ({
  ...await original<typeof import("@azure/msal-browser")>(),
  PublicClientApplication: class {
    initialize() { return Promise.resolve(); }
    getAllAccounts() { return []; }
    addEventCallback() {}
  },
}));
vi.mock("./hooks/useAuthenticatedGraphClient", () => ({ useAuthenticatedGraphClient: () => vi.fn() }));
vi.mock("./hooks/useCachedList", () => ({
  clearCacheKeys: vi.fn(),
  useCachedList: vi.fn((_service, listId) => ({
    data: listId === "fkm" ? [
      { id: "1", fields: { Title: "FKM61400", field_0: "23", field_15: 60 } },
      { id: "2", fields: { Title: "FKM61400", field_0: "26" } },
    ] : [{ id: "1", fields: { field_2: "76100", field_15: 0 } }],
    loading: false, error: null, refresh: vi.fn(),
  })),
}));

let app: typeof import("./App");
beforeAll(async () => { app = await import("./App"); });
beforeEach(() => {
  vi.stubEnv("VITE_SHAREPOINT_SITE_ID", "site");
  vi.stubEnv("VITE_ORING_FKM_LIST_ID", "fkm");
  vi.stubEnv("VITE_ORING_EPDM_LIST_ID", "epdm");
});

describe("gestione FKM ed EPDM", () => {
  it("rende disponibili entrambi i magazzini nella navigazione e nell'Admin", () => {
    const nav = renderToStaticMarkup(createElement(app.NavigationTabs, { activeTab: "oring-fkm", onTabChange: vi.fn() }));
    const admin = renderToStaticMarkup(createElement(AdminPanel, { siteId: "site", oringFkmListId: "fkm", oringEpdmListId: "epdm" }));
    for (const html of [nav, admin]) {
      expect(html).toContain("2_ORING-FKM");
      expect(html).toContain("2_ORING-EPDM");
    }
  });

  it("conserva due record FKM dello stesso codice e la loro selezione indipendente", () => {
    const html = renderToStaticMarkup(createElement(app.OringPanel, { kind: "ORING-FKM", selectedItems: { "ORING-FKM-2": { key: "ORING-FKM-2", source: "ORING-FKM", itemId: "2", title: "FKM61400", fields: {} } }, onToggle: vi.fn(), selectionLimitReached: false }));
    expect(html).toContain("2 risultati");
    expect(html.match(/type="checkbox"/g)).toHaveLength(2);
    expect(html.match(/checked=""/g)).toHaveLength(1);
    expect(html).toContain("Codice SAM");
  });

  it("mostra il codice cliente per EPDM senza codice e applica il limite di selezione", () => {
    const html = renderToStaticMarkup(createElement(app.OringPanel, { kind: "ORING-EPDM", selectedItems: {}, onToggle: vi.fn(), selectionLimitReached: true }));
    expect(html).toContain("Senza codice (cliente 76100)");
    expect(html).toMatch(/type="checkbox"[^>]*disabled=""/);
  });

  it.each(["ORING-FKM", "ORING-EPDM"] as const)("usa i campi ORING nella pagina giacenze per %s", (source) => {
    const html = renderToStaticMarkup(createElement(app.StockUpdatePage, {
      items: [{ key: `${source}-1`, source, itemId: "1", title: "ORING TEST", fields: { field_15: 36, field_16: 2, field_12: "AF1", field_18: "OD1", field_19: "2026-10-22T00:00:00Z" } }],
      values: {}, onChange: vi.fn(), onSave: vi.fn(), saving: false, status: "idle", message: null, onBackToInventory: vi.fn(),
    }));
    expect(html).toContain(source);
    expect(html).toContain('value="36"');
    expect(html).toContain('value="2"');
    expect(html).toContain('value="OD1"');
    expect(html).toContain('value="2026-10-22"');
    expect(html).not.toContain("Giacenza Tubo Intero");
  });

  it.each(["ORING-HNBR", "ORING-FKM", "ORING-EPDM"])("salva %s sulla propria lista senza confondere i campi NBR", (source) => {
    const ids = { "ORING-HNBR": "hnbr", "ORING-FKM": "fkm", "ORING-EPDM": "epdm" };
    const result = app.buildOringStockUpdate(source, { type: "ORING-FKM", title: "TEST", giacenza: "0", prenotazione: "2", commessa: "AF1", nrOrdine: "OD1", dataOrdine: "2026-10-22", dataPrelievo: "" }, ids);
    expect(result.listId).toBe(ids[source]);
    expect(result.payload).toEqual({ field_12: "AF1", field_15: 0, field_16: 2, field_17: null, field_18: "OD1", field_19: "2026-10-22T00:00:00.000Z" });
  });

  it("rifiuta una lista nuova non configurata senza usare quella HNBR", () => {
    expect(() => app.buildOringStockUpdate("ORING-FKM", { type: "ORING-FKM", title: "T" }, { "ORING-HNBR": "hnbr" })).toThrow("List ID non configurato per ORING-FKM");
  });

  it("allinea tipi e campi di tabella, modulo Admin e schema SharePoint", () => {
    const schema = new Map(additionalOringSchema.map((column) => [column.name, column.type]));
    for (const column of oringAdditionalColumns) expect(column.field === "Title" || schema.has(column.field)).toBe(true);
    for (const field of ORING_ADDITIONAL_FIELDS) {
      if (field.key === "Title") continue;
      const type = schema.get(field.key);
      expect(type).toBeDefined();
      if (type === "number") expect(["number", "decimal"]).toContain(field.type);
      if (type === "dateTime") expect(field.type).toBe("date");
    }
  });
});
