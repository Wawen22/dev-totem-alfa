export const isArchivedFlange = (fields: Record<string, unknown>): boolean =>
  fields.Archiviato === true || fields.Archiviato === 1 || ["true", "1", "si", "sì"].includes(String(fields.Archiviato ?? "").trim().toLowerCase());
