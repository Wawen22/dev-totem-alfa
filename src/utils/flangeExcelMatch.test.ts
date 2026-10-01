import { describe, expect, it } from "vitest";
import { matchFlangeExcelRow } from "./flangeExcelMatch";

describe("corrispondenza Excel dei lotti flange", () => {
  const rows = [{ index: 0, values: [["FLCA001", "A"]] }];
  it("non usa il lotto A quando si aggiorna un nuovo lotto B", () => {
    expect(matchFlangeExcelRow(rows, 0, 1, "FLCA001", "B")).toBeNull();
  });
  it("trova il lotto esatto tra più righe dello stesso codice", () => {
    expect(matchFlangeExcelRow([...rows, { index: 1, values: [["FLCA001", "B"]] }], 0, 1, "flca001", " b ")).toBe(1);
  });
  it("riconosce il lotto storico vuoto come A", () => {
    expect(matchFlangeExcelRow([{ index: 2, values: [["FLCA001", ""]] }], 0, 1, "FLCA001", "A")).toBe(2);
    expect(matchFlangeExcelRow(rows, 0, 1, "FLCA001", "")).toBe(0);
  });
  it("non scrive righe ambigue o tabelle prive di identificativo lotto", () => {
    expect(matchFlangeExcelRow([...rows, { index: 1, values: [["FLCA001", "A"]] }], 0, 1, "FLCA001", "A")).toBeNull();
    expect(matchFlangeExcelRow(rows, 0, null, "FLCA001", "B")).toBeNull();
    const historicalDuplicate = [...rows, { index: 1, values: [["FLCA001", ""]] }];
    expect(matchFlangeExcelRow(historicalDuplicate, 0, 1, "FLCA001", "A")).toBeNull();
    expect(matchFlangeExcelRow(historicalDuplicate, 0, 1, "FLCA001", "")).toBeNull();
  });
});
