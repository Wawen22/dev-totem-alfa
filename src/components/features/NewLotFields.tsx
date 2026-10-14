import { newLotColumns, type LotKind } from "../../utils/newLot";

export function NewLotFields({ kind, values, onChange, disabled, exclude = [] }: {
  kind: LotKind;
  values: Record<string, string>;
  onChange: (field: string, value: string) => void;
  disabled?: boolean;
  exclude?: string[];
}) {
  return <div className="field-grid" style={{ marginTop: 12, gridColumn: "1 / -1" }}>
    {newLotColumns(kind).filter((column) => !exclude.includes(column.field)).map((column) => (
      <label className="field" key={column.field}>
        <span>{column.label}</span>
        <input type={column.type === "date" ? "date" : "text"} inputMode={column.type === "number" ? "decimal" : undefined}
          value={values[column.field] ?? ""} disabled={disabled} onChange={(event) => onChange(column.field, event.target.value)} />
      </label>
    ))}
  </div>;
}
