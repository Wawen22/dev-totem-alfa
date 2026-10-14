import { describe, expect, it } from "vitest";
import { isArchivedFlange } from "./flangeArchive";

describe("flange archiviate", () => {
  it("mantiene operativi gli articoli senza flag e riconosce gli archiviati", () => {
    expect(isArchivedFlange({})).toBe(false);
    expect(isArchivedFlange({ Archiviato: false })).toBe(false);
    expect(isArchivedFlange({ Archiviato: "false" })).toBe(false);
    expect(isArchivedFlange({ Archiviato: true })).toBe(true);
    expect(isArchivedFlange({ Archiviato: "true" })).toBe(true);
  });
});
