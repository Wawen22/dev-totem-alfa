import { ForgiatoColumn } from "./forgiatiColumns";

// Colonne per lista 2_ORING-NBR (mappate sugli internal name SharePoint)
export const oringNbrColumns: ForgiatoColumn[] = [
  { field: "Title", label: "Title / Codice" },
  { field: "field_1", label: "Codice cliente" },
  { field: "field_2", label: "Colore Mescola" },
  { field: "field_3", label: "Ø Int." },
  { field: "field_4", label: "Ø Est." },
  { field: "field_5", label: "Ø Corda" },
  { field: "field_6", label: "Materiale" },
  { field: "field_7", label: "Durezza" },
  { field: "field_8", label: "Fluidi" },
  { field: "field_9", label: "Colonna1" },
  { field: "field_10", label: "Colonna2" },
  { field: "field_11", label: "Commessa" },
  { field: "field_12", label: "Design temperature" },
  { field: "field_13", label: "Prezzo unitario" },
  { field: "field_14", label: "Q.tà minima" },
  { field: "field_15", label: "Prenotazione", type: "number" },
  { field: "field_16", label: "Giacenza", type: "number" },
  { field: "field_17", label: "Data prelievo", type: "date" },
];
