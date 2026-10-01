import { describe, expect, it } from "vitest";
import { flangeColumns } from "./flangeColumns";

describe("campi del magazzino flange", () => {
  it("esprime la giacenza in numero pezzi mantenendo il nome interno SharePoint", () => {
    expect(flangeColumns.find((column) => column.field === "GiacenzaMm")?.label).toBe("GIACENZA (Numero Pezzi)");
  });
  it("mostra soltanto il prezzo CAD", () => {
    const visible = flangeColumns.filter((column) => !column.hidden);
    expect(visible.some((column) => column.field === "PrezzoCad")).toBe(true);
    expect(visible.some((column) => column.field === "PrezzoEuroKg")).toBe(false);
  });
});
