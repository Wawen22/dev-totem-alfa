import { ForgiatoColumn } from "./forgiatiColumns";

// Internal names verified on the SharePoint list 11_FLANGE.
export const flangeColumns: ForgiatoColumn[] = [
  { field: "Title", label: "CODICE", width: "150px" },
  { field: "IdentLotto", label: "LOTTO" },
  { field: "CodiceSAM", label: "CODICE SAM" },
  { field: "NumeroOrdine", label: "N° ORDINE" },
  { field: "DataOrdine", label: "DATA OD", type: "date" },
  { field: "Fornitore", label: "FORNITORE" },
  { field: "Quantita", label: "Q.tà.", type: "number" },
  { field: "DN", label: "DN", type: "number" },
  { field: "SP", label: "SP", type: "number" },
  { field: "Grado1", label: "GRADO 1" },
  { field: "Grado2", label: "GRADO 2" },
  { field: "Norma", label: "NORMA" },
  { field: "Classe", label: "CLASSE" },
  { field: "Tipo", label: "TYPE" },
  { field: "NumeroBolla", label: "N° Bolla" },
  { field: "DataConsegna", label: "DATA CONSEGNA", type: "date" },
  { field: "NumeroCertificato", label: "N° CERT." },
  { field: "NumeroColata", label: "N° COLATA" },
  { field: "PrezzoCad", label: "PREZZO CAD", type: "number" },
  { field: "GiacenzaMm", label: "GIACENZA (mm)", type: "number" },
  { field: "Commessa", label: "COMMESSA" },
  { field: "PrezzoEuroKg", label: "Prezzo €/Kg", type: "number" },
  { field: "Commessa2", label: "COMMESSA2" },
  { field: "Modified", label: "Data/ora modifica", type: "date", hidden: true },
];
