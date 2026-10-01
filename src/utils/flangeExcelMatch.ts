import { effectiveFlangeLot } from "./flangeLots";

type ExcelRow = { index: number; values: unknown[][] };

export function matchFlangeExcelRow(rows: ExcelRow[], titleIndex: number | null, lotIndex: number | null, code: string, lot: string): number | null {
  if (titleIndex === null || lotIndex === null) return null;
  const normalize = (value: unknown) => String(value ?? "").trim().toUpperCase();
  const matches: number[] = [];
  const target = effectiveFlangeLot(lot);
  for (const row of rows) {
    const values = row.values[0] || [];
    if (normalize(values[titleIndex]) !== normalize(code)) continue;
    if (effectiveFlangeLot(values[lotIndex]) === target) matches.push(row.index);
  }
  return matches.length === 1 ? matches[0] : null;
}
