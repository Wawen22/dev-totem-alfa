export interface FiloFlussoColumn {
  field: string;
  label: string;
  width?: string;
  type?: 'text' | 'date' | 'number';
}

export const filoFlussoColumns: FiloFlussoColumn[] = [
  { field: 'Title', label: 'Descrizione', width: '300px' },
  { field: 'field_1', label: 'COD.SAM' },
  { field: 'field_2', label: 'COD.ESAB' },
  { field: 'field_3', label: 'LOTTO' },
  { field: 'field_4', label: 'CONFEZ. DA KG' },
  { field: 'field_5', label: 'PREZZO AL KG 2025' },
  { field: 'field_6', label: 'PREZZO AL KG 2026' },
  { field: 'field_7', label: "Q.TA' ACQ. 2025" },
  { field: 'field_8', label: "Q.TA' ACQ. 2026" },
  { field: 'field_9', label: 'GIACENZA' }
];
