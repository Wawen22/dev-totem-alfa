import { describe, expect, it } from "vitest";
import { prepareOringImport } from "./oringInitialImport";

const fkmHeader = ["Posizione", "Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "FLUIDI", "NR Ordine", "Data Ordine", "Q.TA'", "Commessa ", "Design temperature", "Prezzo unitario", "Giacenza", "PRENOTAZIONE FORNITORE", "DATA PR./AGG."];
const epdmHeader = ["Codice Alfa ", "Codice SAM", "Codice cliente ", "Colore Mescola ", "Ø Int.", "Ø Est.", "Ø Corda ", "Materiale", "Durezza", "Fluido", "Ordine", "Data", "Q.TA'", "STANDARD", "Design temperature", "Prezzo unitario", "GIACENZA", "", "DATA", ""];

describe("importazione iniziale ORING", () => {
  it("mappa FKM, conserva i codici ripetuti e scarta solo le posizioni vuote", () => {
    const row = [23, "FKM61400", "OR035497-0533FKM90", 61400, "NERO", 354.99, 365.65, 5.33, "FKM AED LT", "SH 90", "SOUR GAS", "OD-12", "17/12/24", 100, "STD.ALFA", "-40 / 220", 11.49, 60, 2, 46294];
    const items = prepareOringImport("ORING-FKM", [[], fkmHeader, [1], row, [26, ...row.slice(1)]]);
    expect(items).toHaveLength(2);
    expect(items[0].fields).toMatchObject({ Title: "FKM61400", field_0: "23", field_2: "61400", field_4: 354.99, field_12: "STD.ALFA", field_14: 11.49, field_15: 60, field_16: 2, field_18: "OD-12", field_19: "2024-12-17T00:00:00.000Z", QuantitaOrdinata: 100, CodiceSAM: "OR035497-0533FKM90", RigaExcel: 4 });
    expect(items[1].fields).toMatchObject({ Title: "FKM61400", field_0: "26", RigaExcel: 5 });
    expect(items[0].fields.field_17).toBe(new Date((46294 - 25569) * 86400000).toISOString());
  });

  it("mappa EPDM senza scambiare giacenza, prenotazione e colonne senza etichetta", () => {
    const row = ["EPDM 61600", "OR040526-0533", 61600, "NERO + TRATTO BLU", 405.26, 415.92, 5.33, "EPDM", "SH 90", "ACQUA POTABILE", "O-10", "17/12/24", 10, "16 PN16-25", "-30 / 110", 10.68, 36, 3.53, 46300, "4X5,33"];
    const [item] = prepareOringImport("ORING-EPDM", [epdmHeader, row]);
    expect(item.fields).toMatchObject({ Title: "EPDM 61600", field_2: "61600", field_15: 36, field_14: 10.68, field_18: "O-10", field_19: "2024-12-17T00:00:00.000Z", ValoreOriginaleR: "3.53", NoteOriginali: "4X5,33" });
    expect(item.fields.field_16).toBeUndefined();
    expect(item.fields.field_17).toBe(new Date((46300 - 25569) * 86400000).toISOString());
  });

  it("preserva zero, celle vuote e righe senza codice senza creare codici Alfa", () => {
    const row = ["", "", 76100, "NERO", 0, "", "", "EPDM", "SH 90", "ACQUA", "", "", "", "36 PN16-25", "", "", 0];
    const [item] = prepareOringImport("ORING-EPDM", [epdmHeader, row]);
    expect(item.fields.Title).toBe("Senza codice (cliente 76100)");
    expect(item.fields.field_4).toBe(0);
    expect(item.fields.field_15).toBe(0);
    expect(item.fields.field_14).toBeUndefined();
    expect(item.fields.field_17).toBeUndefined();
  });

  it("rifiuta strutture cambiate o valori numerici non validi prima di scrivere", () => {
    expect(() => prepareOringImport("ORING-FKM", [["Codice Alfa"]])).toThrow("Colonna");
    const row = [1, "FKM1", "", 1, "", "errato"];
    expect(() => prepareOringImport("ORING-FKM", [fkmHeader, row])).toThrow("Numero Excel non valido");
    const changed = [...epdmHeader]; changed[17] = "Prenotazione";
    expect(() => prepareOringImport("ORING-EPDM", [changed])).toThrow("Struttura");
  });
});
