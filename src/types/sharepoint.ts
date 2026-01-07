export interface SharePointListItem<TFields = Record<string, unknown>> {
  id: string;
  fields: TFields & { Title?: string };
}

export interface VisitorFields {
  Title?: string; // ID univoco o codice visitatore
  Nome?: string;
  Cognome?: string;
  Email?: string;
  Azienda?: string;
  Categoria?: string;
  Stato?: string;
  Created?: string;
}
