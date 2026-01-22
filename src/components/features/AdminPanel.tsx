import React, { useCallback, useEffect, useMemo, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";
import { SharePointService } from "../../services/sharePointService";
import { useCachedList } from "../../hooks/useCachedList";
import { formatSharePointDate } from "../../utils/dateUtils";
import { SharePointListItem } from "../../types/sharepoint";

type ListKind = "FORGIATI" | "TUBI" | "ORING-HNBR" | "ORING-NBR";

type FieldType = "text" | "number" | "date" | "textarea";

type FieldConfig = {
  key: string;
  label: string;
  type?: FieldType;
  required?: boolean;
  placeholder?: string;
  writable?: boolean; // false -> visualizzato ma non inviato a SharePoint
  sendToSharePoint?: boolean; // false -> mantiene editabile ma non invia a SharePoint
};

type FormState = Record<string, string>;

type AdminPanelProps = {
  siteId?: string;
  forgiatiListId?: string;
  tubiListId?: string;
  oringHnbrListId?: string;
  oringNbrListId?: string;
};

const FORGIATI_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Codice / Title", required: true, placeholder: "Es. FOR-001" },
  { key: "field_1", label: "N° Ordine" },
  { key: "field_2", label: "Data Ordine", type: "date" },
  { key: "field_10", label: "N° Bolla" },
  { key: "field_13", label: "N° Colata" },
  { key: "field_24", label: "Commessa" },
  { key: "field_23", label: "Data prelievo" },
  { key: "field_22", label: "Giacenza Q.tà" },
  { key: "field_21", label: "Giacenza mm da barra" },
  { key: "field_25", label: "Note", type: "textarea", placeholder: "Annotazioni libere" },
  { key: "field_3", label: "Fornitore" },
  { key: "field_4", label: "Pos" },
  { key: "field_5", label: "Q.tà" },
  { key: "field_6", label: "DN" },
  { key: "field_7", label: "Classe" },
  { key: "field_8", label: "No. Disegno - Particolare" },
  { key: "field_9", label: "Grado Materiale" },
  { key: "field_11", label: "Data consegna" },
  { key: "field_12", label: "N° cert" },
  { key: "field_14", label: "Tipo Certificazione" },
  { key: "field_15", label: "Prez. C/D €" },
  { key: "field_16", label: "Ø Est.(mm)" },
  { key: "field_17", label: "Ø Int.(mm)" },
  { key: "field_18", label: "H Altez. (mm)" },
  { key: "field_19", label: "Anello - Disco" },
  { key: "field_20", label: "Grezzo - Sgrossato" },
  { key: "field_26", label: "Data/ora modifica", type: "date", writable: false },
  { key: "CodiceSAM", label: "Codice SAM" },
  { key: "LottoProgressivo", label: "Lotto Progressivo", writable: false },
];

const TUBI_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Codice / Title", required: true, placeholder: "Es. TUB-001" },
  { key: "CodiceSAM", label: "Codice SAM" },
  { key: "field_1", label: "TIPO" },
  { key: "field_2", label: "N° Ordine" },
  { key: "field_3", label: "Data Ordine", type: "date" },
  { key: "field_4", label: "Fornitore" },
  { key: "field_5", label: "P." },
  { key: "field_6", label: "Q.tà" },
  { key: "field_7", label: "Lungh. Tubo (metro)" },
  { key: "field_8", label: "DN (\")" },
  { key: "field_9", label: "DN (mm)" },
  { key: "field_10", label: "SP" },
  { key: "field_11", label: "GRADO" },
  { key: "field_12", label: "PSL1 / PSL2" },
  { key: "field_13", label: "PED" },
  { key: "field_14", label: "HIC" },
  { key: "field_15", label: "N° Bolla" },
  { key: "field_18", label: "N° Colata" },
  { key: "field_25", label: "No. Commessa" },
  { key: "field_21", label: "Data ultimo prelievo" },
  { key: "field_19", label: "Giacenza contab. (mm)", placeholder: "es. 100+20" },
  { key: "field_20", label: "Giacenza non tagliato (mm)" },
  { key: "field_16", label: "Data consegna" },
  { key: "field_22", label: "Prezzo kg/mt" },
  { key: "field_23", label: "Prezzo metro" },
  { key: "field_24", label: "Acquistato dal Curatore" },
  { key: "field_26", label: "Esubero", writable: false },
  { key: "field_27", label: "RITARDO" },
  { key: "LottoProgressivo", label: "Lotto Progressivo", writable: false },
];

const ORING_HNBR_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Codice / Title", required: true, placeholder: "Es. ORH-001" },
  { key: "field_0", label: "Posizione" },
  { key: "field_2", label: "Codice cliente" },
  { key: "field_3", label: "Colore Mescola" },
  { key: "field_4", label: "Ø Int." },
  { key: "field_5", label: "Ø Est." },
  { key: "field_6", label: "Ø Corda" },
  { key: "field_7", label: "Materiale" },
  { key: "field_8", label: "Durezza" },
  { key: "field_9", label: "FLUIDI" },
  { key: "field_10", label: "Colonna1" },
  { key: "field_11", label: "Colonna2" },
  { key: "field_12", label: "Commessa" },
  { key: "field_13", label: "Design temperature" },
  { key: "field_14", label: "Prezzo unitario" },
  { key: "field_15", label: "Giacenza", type: "number" },
  { key: "field_16", label: "Prenotazione", type: "number" },
  { key: "field_17", label: "Data prelievo", type: "date" },
  { key: "field_18", label: "NR Ordine" },
  { key: "field_19", label: "Data ordine", type: "date" },
];

const ORING_NBR_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Codice / Title", required: true, placeholder: "Es. ORN-001" },
  { key: "field_1", label: "Codice cliente" },
  { key: "field_2", label: "Colore Mescola" },
  { key: "field_3", label: "Ø Int." },
  { key: "field_4", label: "Ø Est." },
  { key: "field_5", label: "Ø Corda" },
  { key: "field_6", label: "Materiale" },
  { key: "field_7", label: "Durezza" },
  { key: "field_8", label: "Fluidi" },
  { key: "field_9", label: "Colonna1" },
  { key: "field_10", label: "Colonna2" },
  { key: "field_11", label: "Commessa" },
  { key: "field_12", label: "Design temperature" },
  { key: "field_13", label: "Prezzo unitario" },
  { key: "field_14", label: "Q.tà minima" },
  { key: "field_15", label: "Prenotazione", type: "number" },
  { key: "field_16", label: "Giacenza", type: "number" },
  { key: "field_17", label: "Data prelievo", type: "date" },
];

const toStr = (val: unknown): string => {
  if (val === null || val === undefined) return "";
  return String(val);
};

const getTimeValue = (val: unknown): number => {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") {
    // Excel serial
    return (val - 25569) * 86400 * 1000;
  }
  if (typeof val === "string") {
    const trimmed = val.trim();
    if (/^\d+$/.test(trimmed)) {
      const serial = Number(trimmed);
      return (serial - 25569) * 86400 * 1000;
    }
  }
  const parsed = new Date(val as any);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
};

const toInputDate = (val: unknown): string => {
  const t = getTimeValue(val);
  if (!t) return "";
  return new Date(t).toISOString().slice(0, 10);
};

const toIsoOrNull = (val?: string): string | null => {
  if (!val) return null;
  const d = new Date(val);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
};

const toNumberOrNull = (val?: string): number | null => {
  if (val === undefined || val === null || val === "") return null;
  const parsed = Number(val);
  return Number.isFinite(parsed) ? parsed : null;
};

const formatLottoProg = (val: string | undefined | null) => {
  const str = val ? String(val) : "";
  return str ? str.toUpperCase() : "A";
};

const normalizeExcelKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const TUBI_DATE_FIELDS = new Set(["field_3", "field_16", "field_21", "Modified", "Created"]);

const toExcelSerialDate = (val: unknown): number | "" => {
  const t = getTimeValue(val);
  if (!t) return "";
  return Math.round(t / 86400000 + 25569);
};

const toExcelCellValue = (fieldKey: string | null, value: unknown): string | number | boolean | null => {
  if (!fieldKey) return "";
  if (TUBI_DATE_FIELDS.has(fieldKey)) {
    const serial = toExcelSerialDate(value);
    return serial === "" ? "" : serial;
  }
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
};

const buildTubiExcelColumnMap = () => {
  const map = new Map<string, string>();
  map.set(normalizeExcelKey("CODICE"), "Title");
  map.set(normalizeExcelKey("CODICE SAM"), "CodiceSAM");
  map.set(normalizeExcelKey("NORDINE"), "field_2");
  map.set(normalizeExcelKey("N ORDINE"), "field_2");
  map.set(normalizeExcelKey("DATA OD"), "field_3");
  map.set(normalizeExcelKey("DATA ORDINE"), "field_3");
  map.set(normalizeExcelKey("FORNITORE"), "field_4");
  map.set(normalizeExcelKey("PRODUTTORE"), "field_5");
  map.set(normalizeExcelKey("QTA"), "field_6");
  map.set(normalizeExcelKey("QTA."), "field_6");
  map.set(normalizeExcelKey("LUNGH TUBO METRO"), "field_7");
  map.set(normalizeExcelKey("LUNGH TUBO METRI"), "field_7");
  map.set(normalizeExcelKey("LUNGH TUBO M"), "field_7");
  map.set(normalizeExcelKey("DN"), "field_8");
  map.set(normalizeExcelKey("DNMM"), "field_9");
  map.set(normalizeExcelKey("DN MM"), "field_9");
  map.set(normalizeExcelKey("SP"), "field_10");
  map.set(normalizeExcelKey("GRADO"), "field_11");
  map.set(normalizeExcelKey("PSL1PSL2"), "field_12");
  map.set(normalizeExcelKey("PED"), "field_13");
  map.set(normalizeExcelKey("HIC"), "field_14");
  map.set(normalizeExcelKey("NBOLLA"), "field_15");
  map.set(normalizeExcelKey("N BOLLA"), "field_15");
  map.set(normalizeExcelKey("DATA CONSEGNA"), "field_16");
  map.set(normalizeExcelKey("DATA CONSEGNA."), "field_16");
  map.set(normalizeExcelKey("N CERT"), "field_17");
  map.set(normalizeExcelKey("N CERT."), "field_17");
  map.set(normalizeExcelKey("N COLATA"), "field_18");
  map.set(normalizeExcelKey("GIACENZAMM PER CONTABILITA"), "field_19");
  map.set(normalizeExcelKey("GIACENZAMM CONTABILE"), "field_19");
  map.set(normalizeExcelKey("GIACENZAMM X GESTIONE MAGAZZINO"), "field_20");
  map.set(normalizeExcelKey("GIACENZAMM NON TAGLIATO"), "field_20");
  map.set(normalizeExcelKey("DATA ULTIMO PRELIEVO"), "field_21");
  map.set(normalizeExcelKey("PREZZO KGMT"), "field_22");
  map.set(normalizeExcelKey("PREZZO METRO"), "field_23");
  map.set(normalizeExcelKey("ACQUISTATO DAL CURATORE"), "field_24");
  map.set(normalizeExcelKey("NO COMMESSA"), "field_25");
  map.set(normalizeExcelKey("N COMMESSA"), "field_25");
  map.set(normalizeExcelKey("ESUBERO"), "field_26");
  map.set(normalizeExcelKey("RITARDO"), "field_27");
  map.set(normalizeExcelKey("IDENTLOTTO"), "IdentLotto");
  map.set(normalizeExcelKey("IDENT LOTTO"), "IdentLotto");
  return map;
};

const TUBI_EXCEL_COLUMN_MAP = buildTubiExcelColumnMap();

const buildTubiExcelRow = (excelColumns: string[], fields: Record<string, unknown>) => {
  const fieldKeyLookup = new Map<string, string>();
  Object.keys(fields).forEach((key) => {
    fieldKeyLookup.set(normalizeExcelKey(key), key);
  });

  return excelColumns.map((columnName) => {
    const normalized = normalizeExcelKey(columnName || "");
    const fieldKey =
      TUBI_EXCEL_COLUMN_MAP.get(normalized) || fieldKeyLookup.get(normalized) || null;
    const rawValue = fieldKey ? fields[fieldKey] : null;
    return toExcelCellValue(fieldKey, rawValue);
  });
};

const getExcelColumnIndex = (excelColumns: string[], fieldKey: string) => {
  const target = normalizeExcelKey(fieldKey);
  for (let i = 0; i < excelColumns.length; i++) {
    const normalized = normalizeExcelKey(excelColumns[i] || "");
    const mapped = TUBI_EXCEL_COLUMN_MAP.get(normalized);
    if (mapped && normalizeExcelKey(mapped) === target) {
      return i;
    }
  }
  return null;
};

const findExcelRowIndex = (
  rows: Array<{ index: number; values: Array<Array<unknown>> }>,
  excelColumns: string[],
  options: { codice: string; identLotto: string }
) => {
  const titleIdx = getExcelColumnIndex(excelColumns, "Title");
  const identIdx = getExcelColumnIndex(excelColumns, "IdentLotto");
  if (titleIdx === null || identIdx === null) return null;

  const targetCodice = normalizeExcelKey(options.codice || "");
  const targetIdent = normalizeExcelKey(options.identLotto || "");

  for (const row of rows) {
    const rowValues = row.values?.[0] || [];
    const codiceVal = normalizeExcelKey(String(rowValues[titleIdx] ?? ""));
    const identValRaw = String(rowValues[identIdx] ?? "");
    const identVal = normalizeExcelKey(identValRaw);
    const identMatches =
      identVal === targetIdent ||
      (!identVal && targetIdent === "a");
    if (codiceVal === targetCodice && identMatches) {
      return row.index;
    }
  }

  return null;
};

const parseExcelAddress = (address: string) => {
  const [sheetPartRaw, rangePartRaw] = address.split("!");
  const sheetPart = sheetPartRaw || "";
  const rangePart = rangePartRaw || "";
  const sheetName = sheetPart.replace(/^'+|'+$/g, "");
  const [startCellRaw, endCellRaw] = rangePart.split(":");
  const parseCell = (cell: string) => {
    const match = cell.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;
    return { col: match[1].toUpperCase(), row: Number(match[2]) };
  };
  const startCell = parseCell(startCellRaw || "");
  const endCell = parseCell(endCellRaw || startCellRaw || "");
  if (!startCell) return null;
  return {
    sheetName,
    startCol: startCell.col,
    startRow: startCell.row,
    endCol: endCell ? endCell.col : startCell.col,
    endRow: endCell ? endCell.row : startCell.row,
  };
};

const buildRowRangeAddress = (
  dataBodyAddress: string,
  rowIndex: number,
  rowCount?: number
) => {
  const parsed = parseExcelAddress(dataBodyAddress);
  if (!parsed) return null;
  const effectiveRowCount = rowCount ?? Math.max(0, parsed.endRow - parsed.startRow + 1);
  if (rowIndex < 0 || rowIndex >= effectiveRowCount) return null;
  const rowNumber = parsed.startRow + rowIndex;
  return {
    sheetName: parsed.sheetName,
    address: `${parsed.startCol}${rowNumber}:${parsed.endCol}${rowNumber}`,
  };
};

const buildFormFromItem = (
  item: SharePointListItem<Record<string, unknown>> | null,
  fields: FieldConfig[]
): FormState => {
  const form: FormState = {};
  fields.forEach((field) => {
    if (field.key === "IdentLotto") {
      const raw =
        (item?.fields as Record<string, unknown> | undefined)?.LottoProgressivo ??
        (item?.fields as Record<string, unknown> | undefined)?.IdentLotto;
      form[field.key] = toStr(raw);
      return;
    }
    const value = item?.fields ? (item.fields as Record<string, unknown>)[field.key] : undefined;
    if (field.type === "date") {
      form[field.key] = toInputDate(value);
      return;
    }
    form[field.key] = toStr(value);
  });
  return form;
};

const normalizePayload = (form: FormState, fields: FieldConfig[]): Record<string, unknown> => {
  const payload: Record<string, unknown> = {};
  fields.forEach((field) => {
    if (field.writable === false || field.sendToSharePoint === false) return;
    const raw = form[field.key];
    if (field.type === "number") {
      payload[field.key] = toNumberOrNull(raw);
      return;
    }
    if (field.type === "date") {
      payload[field.key] = toIsoOrNull(raw);
      return;
    }
    if (field.key === "Title") {
      payload.Title = raw?.trim() ? raw.trim() : null;
      return;
    }
    payload[field.key] = raw?.trim() ? raw : null;
  });
  return payload;
};

const LIST_OPTIONS: { kind: ListKind; label: string }[] = [
  { kind: "FORGIATI", label: "1_FORGIATI" },
  { kind: "ORING-HNBR", label: "2_ORING-HNBR" },
  { kind: "ORING-NBR", label: "2_ORING-NBR" },
  { kind: "TUBI", label: "3_TUBI" },
];

const getFieldSet = (kind: ListKind) => {
  if (kind === "FORGIATI") return FORGIATI_FIELDS;
  if (kind === "TUBI") return TUBI_FIELDS;
  if (kind === "ORING-HNBR") return ORING_HNBR_FIELDS;
  return ORING_NBR_FIELDS;
};

const getListNoun = (kind: ListKind) => {
  if (kind === "FORGIATI") return "forgiato";
  if (kind === "TUBI") return "tubo";
  if (kind === "ORING-HNBR") return "oring HNBR";
  return "oring NBR";
};

const getSortDateKey = (kind: ListKind) => {
  if (kind === "FORGIATI") return "field_2";
  if (kind === "TUBI") return "field_3";
  if (kind === "ORING-HNBR") return "field_19";
  return "field_17";
};

const getSearchKeys = (kind: ListKind) => {
  if (kind === "FORGIATI") return ["Title", "field_10", "field_13"];
  if (kind === "TUBI") return ["Title", "field_15", "field_18"];
  if (kind === "ORING-HNBR") return ["Title", "field_18", "field_12", "field_2"];
  return ["Title", "field_11", "field_1"];
};

const getSummaryMeta = (
  kind: ListKind,
  fields: Record<string, unknown>,
  options: { lottoProg?: string } = {}
) => {
  if (kind === "FORGIATI") {
    return [
      { label: "Bolla", value: toStr(fields.field_10) || "-" },
      { label: "Colata", value: toStr(fields.field_13) || "-" },
      { label: "Giacenza", value: toStr(fields.field_22) || "-" },
    ];
  }
  if (kind === "TUBI") {
    return [
      { label: "Bolla", value: toStr(fields.field_15) || "-" },
      { label: "Colata", value: toStr(fields.field_18) || "-" },
      { label: "Giacenza", value: toStr(fields.field_19) || "-" },
      { label: "Lotto", value: formatLottoProg(options.lottoProg || toStr(fields.LottoProgressivo) || "A") },
    ];
  }
  if (kind === "ORING-HNBR") {
    return [
      { label: "Ordine", value: toStr(fields.field_18) || "-" },
      { label: "Commessa", value: toStr(fields.field_12) || "-" },
      { label: "Giacenza", value: toStr(fields.field_15) || "-" },
    ];
  }
  return [
    { label: "Commessa", value: toStr(fields.field_11) || "-" },
    { label: "Prenotazione", value: toStr(fields.field_15) || "-" },
    { label: "Giacenza", value: toStr(fields.field_16) || "-" },
  ];
};

export function AdminPanel({ siteId, forgiatiListId, tubiListId, oringHnbrListId, oringNbrListId }: AdminPanelProps) {
  const getClient = useAuthenticatedGraphClient();
  const tubiExcelPath = (import.meta.env.VITE_TUBI_EXCEL_PATH || "").trim();
  const tubiExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const tubiExcelFilename = (import.meta.env.VITE_SP_TUBI_FILENAME || import.meta.env.VITE_TUBI_EXCEL_FILE || "").trim();
  const tubiExcelTable = (import.meta.env.VITE_TUBI_EXCEL_TABLE || "tblTUBI").trim();
  const tubiExcelDriveIdEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_ID || "").trim();
  const tubiExcelDriveNameEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const [tubiExcelDriveId, setTubiExcelDriveId] = useState<string | null>(tubiExcelDriveIdEnv || null);

  const service = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);

  const {
    data: forgiatiItems,
    loading: forgiatiLoading,
    error: forgiatiError,
    refresh: refreshForgiati,
  } = useCachedList<Record<string, unknown>>(service, forgiatiListId, "admin-forgiati");

  const { data: tubiItems, loading: tubiLoading, error: tubiError, refresh: refreshTubi } = useCachedList<Record<string, unknown>>(
    service,
    tubiListId,
    "admin-tubi"
  );

  const {
    data: oringHnbrItems,
    loading: oringHnbrLoading,
    error: oringHnbrError,
    refresh: refreshOringHnbr,
  } = useCachedList<Record<string, unknown>>(service, oringHnbrListId, "admin-oring-hnbr");

  const {
    data: oringNbrItems,
    loading: oringNbrLoading,
    error: oringNbrError,
    refresh: refreshOringNbr,
  } = useCachedList<Record<string, unknown>>(service, oringNbrListId, "admin-oring-nbr");

  const [activeList, setActiveList] = useState<ListKind>("FORGIATI");
  const [search, setSearch] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<FormState>(() => buildFormFromItem(null, FORGIATI_FIELDS));
  const [createForm, setCreateForm] = useState<FormState>(() => buildFormFromItem(null, FORGIATI_FIELDS));
  const [updateStatus, setUpdateStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [updateMessage, setUpdateMessage] = useState<string | null>(null);
  const [createStatus, setCreateStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [createMessage, setCreateMessage] = useState<string | null>(null);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [isCreateOpen, setIsCreateOpen] = useState(false);

  const activeItems =
    activeList === "FORGIATI"
      ? forgiatiItems
      : activeList === "TUBI"
      ? tubiItems
      : activeList === "ORING-HNBR"
      ? oringHnbrItems
      : oringNbrItems;
  const activeLoading =
    activeList === "FORGIATI"
      ? forgiatiLoading
      : activeList === "TUBI"
      ? tubiLoading
      : activeList === "ORING-HNBR"
      ? oringHnbrLoading
      : oringNbrLoading;
  const activeError =
    activeList === "FORGIATI"
      ? forgiatiError
      : activeList === "TUBI"
      ? tubiError
      : activeList === "ORING-HNBR"
      ? oringHnbrError
      : oringNbrError;
  const activeListId =
    activeList === "FORGIATI"
      ? forgiatiListId
      : activeList === "TUBI"
      ? tubiListId
      : activeList === "ORING-HNBR"
      ? oringHnbrListId
      : oringNbrListId;
  const activeRefresh =
    activeList === "FORGIATI"
      ? refreshForgiati
      : activeList === "TUBI"
      ? refreshTubi
      : activeList === "ORING-HNBR"
      ? refreshOringHnbr
      : refreshOringNbr;
  const fieldSet = getFieldSet(activeList);

  const tubiGroupedRows = useMemo(() => {
    const map = new Map<string, SharePointListItem<Record<string, unknown>>[]>();
    (tubiItems || []).forEach((item) => {
      const title = toStr((item.fields as Record<string, unknown>).Title) || "(senza codice)";
      const existing = map.get(title) || [];
      map.set(title, [...existing, item]);
    });
    return Array.from(map.entries()).map(([title, items]) => ({ title, items }));
  }, [tubiItems]);

  const tubiProgressiveMap = useMemo(() => {
    const map = new Map<string, string>();
    const normalizeLetter = (val: string | null | undefined) => {
      if (!val) return null;
      const cleaned = String(val).trim().toLowerCase();
      if (/^[a-z]$/.test(cleaned)) return cleaned;
      return null;
    };
    const getCreatedKey = (item: SharePointListItem<Record<string, unknown>>) => {
      const created = getTimeValue((item.fields as Record<string, unknown>).Created);
      if (created) return created;
      const idNum = Number(item.id);
      return Number.isFinite(idNum) ? idNum : 0;
    };

    tubiGroupedRows.forEach((group) => {
      const sorted = [...group.items].sort((a, b) => getCreatedKey(a) - getCreatedKey(b));
      const used = new Set<string>();
      sorted.forEach((item) => {
        const existing = normalizeLetter((item.fields as Record<string, unknown>).LottoProgressivo as string | undefined);
        let letter = existing;
        if (!letter || used.has(letter)) {
          let code = "a".charCodeAt(0);
          while (used.has(String.fromCharCode(code))) {
            code++;
          }
          letter = String.fromCharCode(code);
        }
        used.add(letter);
        map.set(item.id, letter);
      });
    });

    return map;
  }, [tubiGroupedRows]);

  const sortedItems = useMemo(() => {
    const items = [...activeItems];
    const dateKey = getSortDateKey(activeList);
    items.sort(
      (a, b) =>
        getTimeValue((b.fields as Record<string, unknown>)[dateKey]) -
        getTimeValue((a.fields as Record<string, unknown>)[dateKey])
    );
    return items;
  }, [activeItems, activeList]);

  const filteredItems = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return sortedItems;
    return sortedItems.filter((item) => {
      const fields = item.fields as Record<string, unknown>;
      return getSearchKeys(activeList).some((key) => toStr(fields[key]).toLowerCase().includes(term));
    });
  }, [sortedItems, search, activeList]);

  useEffect(() => {
    setPage(1);
  }, [search, activeList, sortedItems.length]);

  const totalPages = Math.max(1, Math.ceil(filteredItems.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const start = (currentPage - 1) * pageSize;
  const visibleItems = filteredItems.slice(start, start + pageSize);

  // Selezione automatica primo elemento quando cambia lista o filtro
  useEffect(() => {
    const inFiltered = selectedId && filteredItems.some((item) => item.id === selectedId);
    if (!inFiltered) {
      setSelectedId(visibleItems[0]?.id ?? null);
    }
  }, [filteredItems, visibleItems, selectedId]);

  // Aggiorna form di modifica quando cambia elemento selezionato o tab
  useEffect(() => {
    const current = filteredItems.find((item) => item.id === selectedId) ?? null;
    setEditForm(buildFormFromItem(current, fieldSet));
  }, [selectedId, filteredItems, fieldSet]);

  // Reset form di creazione quando cambio tab
  useEffect(() => {
    setCreateForm(buildFormFromItem(null, fieldSet));
    setUpdateStatus("idle");
    setUpdateMessage(null);
    setCreateStatus("idle");
    setCreateMessage(null);
    setIsCreateOpen(false);
  }, [activeList, fieldSet]);

  const handleFieldChange = (formSetter: React.Dispatch<React.SetStateAction<FormState>>) =>
    (key: string, value: string) => {
      formSetter((prev) => ({ ...prev, [key]: value }));
    };

  const handleUpdate = useCallback(async () => {
    if (!service) {
      setUpdateStatus("error");
      setUpdateMessage("Servizio SharePoint non configurato");
      return;
    }
    if (!activeListId) {
      setUpdateStatus("error");
      setUpdateMessage("List ID mancante per la lista selezionata");
      return;
    }
    if (!selectedId) {
      setUpdateStatus("error");
      setUpdateMessage("Seleziona un elemento da modificare");
      return;
    }

    const fields = getFieldSet(activeList);
    const payload = normalizePayload(editForm, fields);
    if (!payload.Title) {
      setUpdateStatus("error");
      setUpdateMessage("Il campo Title è obbligatorio");
      return;
    }

    setUpdateStatus("saving");
    setUpdateMessage(null);

    try {
      await service.updateItem<Record<string, unknown>>(activeListId, selectedId, payload);
      let excelError: string | null = null;
      if (activeList === "TUBI") {
        try {
          const resolvedPath = tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
          let resolvedDriveId = tubiExcelDriveId;
          if (!resolvedDriveId && tubiExcelDriveNameEnv) {
            resolvedDriveId = await service.getDriveIdByName(tubiExcelDriveNameEnv);
            setTubiExcelDriveId(resolvedDriveId);
          }
          if (!resolvedDriveId && tubiExcelDriveNameEnv) {
            throw new Error(`Libreria "${tubiExcelDriveNameEnv}" non trovata`);
          }
          if (!resolvedPath || !tubiExcelTable) {
            throw new Error("Percorso Excel o tabella non configurati");
          }

          const currentItem = filteredItems.find((item) => item.id === selectedId) || null;
          const identRaw =
            (currentItem?.fields as any)?.LottoProgressivo ||
            (currentItem?.id ? tubiProgressiveMap.get(currentItem.id) : undefined);
          const identLotto = formatLottoProg(identRaw || "A");
          const lookupTitle = toStr((currentItem?.fields as any)?.Title) || editForm.Title || "";

          const driveItem = await service.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await service.listWorkbookTableColumnsByItemId(driveItem.id, tubiExcelTable, resolvedDriveId || undefined);
          const rows = await service.listWorkbookTableRowsByItemId(driveItem.id, tubiExcelTable, resolvedDriveId || undefined);
          const rowIndex = findExcelRowIndex(rows, columns, {
            codice: lookupTitle,
            identLotto,
          });

          if (rowIndex === null) {
            throw new Error(`Riga Excel non trovata per ${editForm.Title || "codice"} (${identLotto})`);
          }

          const dataBodyRange = await service.getWorkbookTableDataBodyRangeByItemId(
            driveItem.id,
            tubiExcelTable,
            resolvedDriveId || undefined
          );
          const rowRange = buildRowRangeAddress(dataBodyRange.address, rowIndex, dataBodyRange.rowCount);
          if (!rowRange) {
            throw new Error("Impossibile calcolare la riga Excel da aggiornare");
          }

          const excelFields = {
            ...(currentItem?.fields || {}),
            ...payload,
            IdentLotto: identLotto,
          } as Record<string, unknown>;
          const rowValues = buildTubiExcelRow(columns, excelFields);
          const sessionId = await service.createWorkbookSessionByItemId(
            driveItem.id,
            { persistChanges: true },
            resolvedDriveId || undefined
          );
          try {
            await service.updateWorkbookRangeByAddress(
              driveItem.id,
              rowRange.sheetName,
              rowRange.address,
              [rowValues],
              { sessionId },
              resolvedDriveId || undefined
            );
          } finally {
            try {
              await service.closeWorkbookSessionByItemId(
                driveItem.id,
                sessionId,
                resolvedDriveId || undefined
              );
            } catch (closeErr) {
              console.warn("Errore chiusura sessione Excel TUBI (admin)", closeErr);
            }
          }
        } catch (excelErr: any) {
          console.error("Errore aggiornamento Excel TUBI (admin)", excelErr);
          excelError = excelErr?.message || "Errore aggiornamento Excel";
        }
      }

      setUpdateStatus("success");
      setUpdateMessage(
        excelError
          ? `Elemento aggiornato; Excel non aggiornato: ${excelError}`
          : "Elemento aggiornato con successo"
      );
      activeRefresh();
    } catch (err: any) {
      setUpdateStatus("error");
      setUpdateMessage(err?.message || "Errore durante il salvataggio");
    }
  }, [
    service,
    activeListId,
    selectedId,
    editForm,
    activeRefresh,
    activeList,
    filteredItems,
    tubiExcelPath,
    tubiExcelFolder,
    tubiExcelFilename,
    tubiExcelTable,
    tubiExcelDriveId,
    tubiExcelDriveNameEnv,
  ]);

  const handleCreate = useCallback(async () => {
    if (!service) {
      setCreateStatus("error");
      setCreateMessage("Servizio SharePoint non configurato");
      return;
    }
    if (!activeListId) {
      setCreateStatus("error");
      setCreateMessage("List ID mancante per la lista selezionata");
      return;
    }

    const fields = getFieldSet(activeList);
    const payload = normalizePayload(createForm, fields);
    if (!payload.Title) {
      setCreateStatus("error");
      setCreateMessage("Il campo Title è obbligatorio");
      return;
    }

    setCreateStatus("saving");
    setCreateMessage(null);

    try {
      await service.createItem<Record<string, unknown>>(activeListId, payload);
      setCreateStatus("success");
      setCreateMessage("Nuovo elemento creato");
      setCreateForm(buildFormFromItem(null, fields));
      setIsCreateOpen(false);
      activeRefresh();
    } catch (err: any) {
      setCreateStatus("error");
      const detail = err?.body?.error?.message || err?.message;
      setCreateMessage(detail || "Errore durante la creazione");
    }
  }, [service, activeListId, createForm, activeRefresh, activeList]);

  const renderField = (field: FieldConfig, form: FormState, onChange: (key: string, val: string) => void) => {
    const commonProps = {
      value: form[field.key] ?? "",
      onChange: (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => onChange(field.key, e.target.value),
      placeholder: field.placeholder,
      disabled: field.writable === false,
    };

    if (field.type === "textarea") {
      return (
        <label key={field.key} className="field full">
          <span>
            {field.label}
            {field.required ? " *" : ""}
          </span>
          <textarea rows={3} {...commonProps} />
        </label>
      );
    }

    return (
      <label key={field.key} className="field">
        <span>
          {field.label}
          {field.required ? " *" : ""}
        </span>
        <input
          type={field.type === "number" ? "number" : field.type === "date" ? "date" : "text"}
          {...commonProps}
        />
      </label>
    );
  };

  const activeSummary = (item: SharePointListItem<Record<string, unknown>>) => {
    const fields = item.fields as Record<string, unknown>;
    const lottoProg = item.id ? tubiProgressiveMap.get(item.id) : undefined;
    return {
      title: toStr(fields.Title) || "-",
      meta: getSummaryMeta(activeList, fields, { lottoProg }),
      modified:
        formatSharePointDate(fields.Modified) || formatSharePointDate(fields["field_26"] || fields["Modified"]),
    };
  };

  const selectedItem = filteredItems.find((item) => item.id === selectedId) || null;

  return (
    <div className="admin-panel">
      <div className="admin-header">
        <div>
          <p className="eyebrow" style={{ marginBottom: 4 }}>
            Gestione inventario
          </p>
          <h2 style={{ margin: 0 }}>Pannello Admin</h2>
          <p className="muted" style={{ marginTop: 6 }}>
            Modifica o inserisci articoli di FORGIATI, TUBI e ORING senza uscire dal totem.
          </p>
        </div>
        <div className="admin-header-actions">
          <span className="pill ghost" style={{ alignSelf: "flex-start", marginTop: 6 }}>
            Ruolo Totem.Admin attivo
          </span>
          <button className="btn primary" type="button" onClick={() => setIsCreateOpen(true)}>
            + Nuovo articolo
          </button>
        </div>
      </div>

      <div className="admin-switch">
        {LIST_OPTIONS.map(({ kind, label }) => (
          <button
            key={kind}
            className={`admin-switch__btn ${activeList === kind ? "active" : ""}`}
            onClick={() => setActiveList(kind)}
            type="button"
          >
            {label}
          </button>
        ))}
      </div>

      <div className="panel">
        <div className="panel-content admin-grid">
          <div className="admin-column">
            <div className="admin-toolbar">
              <div className="admin-search">
                <span role="img" aria-label="search">
                  🔍
                </span>
                <input
                  type="search"
                  placeholder="Cerca per codice o campi principali"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <div className="pill ghost">{filteredItems.length} risultati</div>
              <label className="admin-page-size">
                <span>Per pagina</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setPage(1);
                  }}
                >
                  {[20, 50, 100].map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </label>
              <div className="admin-pager">
                <button
                  className="icon-btn"
                  aria-label="Pagina precedente"
                  onClick={() => setPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  type="button"
                >
                  ←
                </button>
                <span>
                  {currentPage} / {totalPages}
                </span>
                <button
                  className="icon-btn"
                  aria-label="Pagina successiva"
                  onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  type="button"
                >
                  →
                </button>
              </div>
              <button className="btn secondary" onClick={() => activeRefresh()} disabled={activeLoading} type="button">
                {activeLoading ? "Aggiorno..." : "Aggiorna"}
              </button>
              {activeError && <span className="pill danger" style={{ background: "#fee2e2", color: "#b91c1c" }}>{activeError}</span>}
              {!activeListId && (
                <span className="pill warning" style={{ background: "#fff4e5", color: "#c2410c" }}>
                  Configurare l'ID della lista {activeList}
                </span>
              )}
            </div>

            <div className="admin-list">
              {filteredItems.length === 0 && (
                <div className="muted" style={{ padding: "12px 6px" }}>
                  Nessun elemento trovato. Verifica il filtro o aggiorna.
                </div>
              )}
              {visibleItems.map((item) => {
                const summary = activeSummary(item);
                const isSelected = selectedId === item.id;
                return (
                  <button
                    key={item.id}
                    className={`admin-list-item ${isSelected ? "selected" : ""}`}
                    onClick={() => setSelectedId(item.id)}
                    type="button"
                  >
                    <div className="admin-list-item__title">{summary.title}</div>
                    <div className="admin-list-item__meta">
                      {summary.meta.map((entry) => (
                        <span key={entry.label} className="pill ghost">
                          {entry.label} {entry.value}
                        </span>
                      ))}
                      {summary.modified && <span className="pill ghost">Mod. {summary.modified}</span>}
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="admin-column" style={{ gap: 16 }}>
            <div className="admin-card">
              <div className="admin-card__header">
                <div>
                  <p className="eyebrow" style={{ marginBottom: 4 }}>
                    Modifica elemento selezionato
                  </p>
                  <h3 style={{ margin: 0 }}>{selectedItem ? activeSummary(selectedItem).title : "Seleziona un elemento"}</h3>
                </div>
                <div className="pill ghost">{activeList}</div>
              </div>

              <div className="admin-form-grid">
                {fieldSet.map((field) => renderField(field, editForm, handleFieldChange(setEditForm)))}
              </div>

              {updateMessage && (
                <div className={`alert ${updateStatus === "success" ? "success" : updateStatus === "error" ? "error" : "warning"}`}>
                  {updateMessage}
                </div>
              )}

              <div className="admin-actions">
                <button className="btn secondary" onClick={() => setEditForm(buildFormFromItem(selectedItem, fieldSet))} type="button">
                  Ripristina valori
                </button>
                <button className="btn primary" onClick={handleUpdate} disabled={updateStatus === "saving"} type="button">
                  {updateStatus === "saving" ? "Salvo..." : "Salva modifiche"}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {isCreateOpen && (
        <div className="modal-backdrop blur" role="dialog" aria-modal="true">
          <div className="modal" style={{ maxWidth: 900, width: "95%" }}>
            <div className="modal__header">
                <div>
                  <p className="eyebrow" style={{ marginBottom: 4 }}>Inserisci nuovo articolo</p>
                  <h3 style={{ margin: 0 }}>Nuovo {getListNoun(activeList)}</h3>
                </div>
              <button className="icon-btn" aria-label="Chiudi" onClick={() => setIsCreateOpen(false)} type="button">✕</button>
            </div>
            <div className="modal__body">
              <div className="admin-form-grid">
                {fieldSet.map((field) => renderField(field, createForm, handleFieldChange(setCreateForm)))}
              </div>
              {createMessage && (
                <div className={`alert ${createStatus === "success" ? "success" : createStatus === "error" ? "error" : "warning"}`} style={{ marginTop: 12 }}>
                  {createMessage}
                </div>
              )}
            </div>
            <div className="modal__footer">
              <button className="btn secondary" onClick={() => setCreateForm(buildFormFromItem(null, fieldSet))} type="button">
                Pulisci campi
              </button>
              <button className="btn primary" onClick={handleCreate} disabled={createStatus === "saving"} type="button">
                {createStatus === "saving" ? "Creo..." : "Crea nuovo"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
