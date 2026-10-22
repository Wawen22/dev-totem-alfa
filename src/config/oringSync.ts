export const ORING_KINDS = ["ORING-HNBR", "ORING-NBR", "ORING-FKM", "ORING-EPDM"] as const;
export type OringKind = typeof ORING_KINDS[number];
export const isOringKind = (kind: string): kind is OringKind =>
  (ORING_KINDS as readonly string[]).includes(kind);

// The legacy NBR worksheet really has a trailing space in its name.
export const oringSheetName = (kind: OringKind): string => {
  const material = kind.slice(6);
  return import.meta.env[`VITE_ORING_${material}_EXCEL_SHEET`] || (material === "NBR" ? "NBR " : material);
};
