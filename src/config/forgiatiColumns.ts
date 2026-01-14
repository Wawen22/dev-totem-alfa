export interface ForgiatoColumn {
  field: string;
  label: string;
  width?: string;
  type?: 'text' | 'date' | 'number';
}

// Aggiorna i valori `field` con gli internal name reali della lista SharePoint.
export const forgiatiColumns: ForgiatoColumn[] = [
  { field: "Title", label: "Codice / Title", width: "200px" },
  { field: "field_1", label: "N° ordine" },
  { field: "field_2", label: "Data Ord", type: "date" },
  { field: "field_3", label: "Fornitore" },
  { field: "field_4", label: "Pos" },
  { field: "field_5", label: "Q.tà" },
  { field: "field_6", label: "DN" },
  { field: "field_7", label: "Classe" },
  { field: "field_8", label: "No. Disegno - Particolare" },
  { field: "field_9", label: "Grado Materiale" },
  { field: "field_10", label: "N° Bolla." },
  { field: "field_11", label: "Data consegna", type: "date" },
  { field: "field_12", label: "N° cert" },
  { field: "field_13", label: "N° colata" },
  { field: "field_14", label: "Tipo Certificazione" },
  { field: "field_15", label: "Prez. C/D €" },
  { field: "field_16", label: "Ø Est.(mm)" },
  { field: "field_17", label: "Ø Int.(mm)" },
  { field: "field_18", label: "H Altez. (mm)" },
  { field: "field_19", label: "Anello - Disco" },
  { field: "field_20", label: "Grezzo - Sgrossato" },
  { field: "field_21", label: "Giacenza mm lunghezza da barra" },
  { field: "field_22", label: "Giacenza Q.tà" },
  { field: "field_23", label: "Data prelievo", type: "date" },
  { field: "field_24", label: "Commessa" },
  { field: "field_25", label: "Note" },
  { field: "field_26", label: "Data/ora modifica", type: "date" },
  { field: "CodiceSAM", label: "Codice SAM" },
];
