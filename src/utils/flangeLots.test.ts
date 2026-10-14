import { describe, expect, it } from "vitest";
import { assertUniqueFlangeLots, buildFlangeLotPayload, flangeLotLabels, nextFlangeLot } from "./flangeLots";

describe("nuovi lotti flange", () => {
  it("blocca la sincronizzazione di due record con la stessa identità storica", () => {
    expect(() => assertUniqueFlangeLots([{ id: "1", fields: { Title: "FLCA001", IdentLotto: "" } }, { id: "2", fields: { Title: "flca001", IdentLotto: "A" } }])).toThrow("Identificativo lotto ambiguo");
    expect(() => assertUniqueFlangeLots([{ id: "1", fields: { Title: "FLCA001", IdentLotto: "A" } }, { id: "2", fields: { Title: "FLCA002", IdentLotto: "A" } }])).not.toThrow();
  });
  it("mantiene l'identità dei lotti storici e duplicati senza rinominarli", () => {
    const labels = flangeLotLabels([{ id: "1", fields: {} }, { id: "2", fields: { IdentLotto: "A" } }, { id: "3", fields: { IdentLotto: "a" } }]);
    expect([...labels.entries()]).toEqual([["1", "A"], ["2", "A"], ["3", "A"]]);
  });
  it("evita identificativi già usati, anche con spazi e minuscole", () => {
    expect(nextFlangeLot([" a ", "B", null])).toBe("C");
    expect(nextFlangeLot(Array.from({ length: 26 }, (_, index) => String.fromCharCode(65 + index)))).toBe("AA");
  });
  it("non copia nessun dato del lotto originale, tranne il codice articolo", () => {
    const source = { Title: "FLCA046", DN: 24, SP: 11, PrezzoCad: 289, GiacenzaMm: 7, Quantita: 12, NumeroBolla: "123", NumeroCertificato: "CERT", DataConsegna: "2026-10-01", Commessa: "AF4870", PrezzoEuroKg: 99, Modified: "2026-10-01", id: "5" };
    const payload = buildFlangeLotPayload(source, { identLotto: "C", colata: " nuova ", ordine: "666", codiceSam: "SAM", dataOrdine: null });
    expect(payload).toEqual({ Title: "FLCA046", IdentLotto: "C", NumeroColata: "nuova", NumeroOrdine: "666", CodiceSAM: "SAM" });
    expect(source.GiacenzaMm).toBe(7);
  });
});
