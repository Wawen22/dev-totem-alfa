import { describe, expect, it } from "vitest";
import { buildNewLotPayload, newLotColumns, type LotKind } from "./newLot";

describe("creazione manuale lotti", () => {
  for (const kind of ["FORGIATI", "TUBI", "TUBO-MECCANICO", "FLANGE"] as LotKind[]) {
    it(`${kind}: un modulo vuoto salva solo codice e identificativo`, () => {
      expect(buildNewLotPayload(kind, " TEST001 ", "b", {})).toEqual({ Title: "TEST001", IdentLotto: "B" });
      expect(buildNewLotPayload(kind, "TEST001", "B", { Modified: "2026-01-01", Title: "OTHER", IdentLotto: "A", Archiviato: true })).toEqual({ Title: "TEST001", IdentLotto: "B" });
    });
  }
  it("consente di compilare anche caratteristiche, quantità e prezzo", () => {
    expect(buildNewLotPayload("FLANGE", "FLCA044", "B", { DN: "50", PrezzoCad: "12,5", DataOrdine: "2026-10-14", NumeroColata: "123" })).toEqual({ Title: "FLCA044", IdentLotto: "B", DN: 50, PrezzoCad: 12.5, DataOrdine: "2026-10-14T00:00:00.000Z", NumeroColata: "123" });
    expect(newLotColumns("TUBI").some((column) => column.field === "Schedula")).toBe(true);
  });
});
