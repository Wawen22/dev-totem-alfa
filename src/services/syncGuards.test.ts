import { describe, expect, it } from "vitest";
import {
  MAX_CONSECUTIVE_SYNC_WRITE_FAILURES,
  getMissingRequiredColumns,
  recordWriteOutcome,
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
