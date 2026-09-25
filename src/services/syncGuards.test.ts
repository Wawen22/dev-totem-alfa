import { describe, expect, it } from "vitest";
import {
  buildExcelAuthoritativePatch,
  MAX_CONSECUTIVE_SYNC_WRITE_FAILURES,
  getMissingRequiredColumns,
  recordWriteOutcome,
  shouldApplyExcelAuthoritativeUpdate,
} from "./syncGuards";

describe("getMissingRequiredColumns", () => {
  it("restituisce soltanto i nomi interni assenti", () => {
    expect(
      getMissingRequiredColumns(
        ["Title", "IdentLotto", "field_1"],
        ["Title", "IdentLotto", "field_1", "field_3"]
      )
    ).toEqual(["field_3"]);
  });

  it("elimina i requisiti ripetuti", () => {
    expect(getMissingRequiredColumns(["Title"], ["Title", "Title"])).toEqual([]);
  });

  it("espone sia il campo tecnico sia il campo data mancanti", () => {
    expect(
      getMissingRequiredColumns(["Title"], ["Title", "IdentLotto", "field_3"])
    ).toEqual(["IdentLotto", "field_3"]);
  });
});

describe("recordWriteOutcome", () => {
  it("azzera gli errori dopo una scrittura riuscita", () => {
    expect(recordWriteOutcome(2, true)).toBe(0);
  });

  it("raggiunge la soglia di arresto al terzo errore consecutivo", () => {
    const failures = recordWriteOutcome(recordWriteOutcome(1, false), false);
    expect(failures).toBe(MAX_CONSECUTIVE_SYNC_WRITE_FAILURES);
  });
});

describe("shouldApplyExcelAuthoritativeUpdate", () => {
  it("applica Excel quando un campo aziendale è diverso", () => {
    expect(shouldApplyExcelAuthoritativeUpdate(1, false)).toBe(true);
  });

  it("non scrive quando riga e identificativo lotto sono già allineati", () => {
    expect(shouldApplyExcelAuthoritativeUpdate(0, false)).toBe(false);
  });

  it("allinea l'identificativo lotto anche senza differenze aziendali", () => {
    expect(shouldApplyExcelAuthoritativeUpdate(0, true)).toBe(true);
  });
});

describe("buildExcelAuthoritativePatch", () => {
  it("include un vuoto esplicito ma non i campi assenti dalla tabella Excel", () => {
    expect(
      buildExcelAuthoritativePatch(
        {
          Title: "TMA001",
          IdentLotto: "A",
          field_1: null,
          field_2: "559-13 R0",
          field_5: null,
        },
        ["Title", "field_1", "field_2"],
        true
      )
    ).toEqual({
      Title: "TMA001",
      IdentLotto: "A",
      field_1: null,
      field_2: "559-13 R0",
    });
  });

  it("allinea solo il lotto quando non ci sono differenze aziendali", () => {
    expect(
      buildExcelAuthoritativePatch(
        { Title: "TMA001", IdentLotto: "A", field_1: "test1" },
        ["Title", "field_1"],
        false
      )
    ).toEqual({ IdentLotto: "A" });
  });
});
