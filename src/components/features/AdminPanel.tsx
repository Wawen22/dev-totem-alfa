import React, { useCallback, useEffect, useMemo, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";
import { SharePointService } from "../../services/sharePointService";
import { useCachedList } from "../../hooks/useCachedList";
import { formatSharePointDate } from "../../utils/dateUtils";
import { SharePointListItem } from "../../types/sharepoint";

type ListKind = "FORGIATI" | "TUBI" | "ORING-HNBR" | "ORING-NBR" | "SPARK-GUPS";

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
  sparkGupsListId?: string;
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
  { key: "GRADOMATERIALE2", label: "Grado Materiale 2" },
  { key: "field_11", label: "Data consegna", type: "date" },
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
  { key: "GRADOMATERIALE2", label: "Grado materiale 2" },
  { key: "field_12", label: "PSL1 / PSL2" },
  { key: "field_13", label: "PED" },
  { key: "field_14", label: "HIC" },
  { key: "field_15", label: "N° Bolla" },
  { key: "field_17", label: "N° Certificato" },
  { key: "field_18", label: "N° Colata" },
  { key: "field_25", label: "No. Commessa" },
  { key: "field_21", label: "Data ultimo prelievo" },
  { key: "field_19", label: "Giacenza contab. (mm)", placeholder: "es. 100+20" },
  { key: "field_20", label: "Giacenza non tagliato (mm)" },
  { key: "field_16", label: "Data consegna", type: "date" },
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

const SPARK_GUPS_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Title", required: true, placeholder: "Es. SPK-001" },
  { key: "field_1", label: "Lotto" },
  { key: "field_2", label: "Codice SAM" },
  { key: "field_3", label: "Tipologia Articolo" },
  { key: "field_4", label: "N. Ordine" },
  { key: "field_5", label: "Data Ordine" },
  { key: "field_6", label: "Fornitore" },
  { key: "field_7", label: "Quantità Ordinata" },
  { key: "field_8", label: "N. Bolla" },
  { key: "field_9", label: "Data Consegna" },
  { key: "field_10", label: "Giacenza" },
  { key: "field_11", label: "Data Ultimo Prelievo" },
  { key: "field_12", label: "Prezzo unitario" },
  { key: "field_13", label: "Commessa" },
];

const toStr = (val: unknown): string => {
  if (val === null || val === undefined) return "";
  return String(val);
};

const compareTubiTitle = (a: string, b: string) => {
  const left = a.trim().toUpperCase();
  const right = b.trim().toUpperCase();
  if (!left && !right) return 0;
  if (!left) return 1;
  if (!right) return -1;
  const matchLeft = left.match(/^([A-Z]+)(\d+)$/);
  const matchRight = right.match(/^([A-Z]+)(\d+)$/);
  if (matchLeft && matchRight) {
    const prefixCompare = matchLeft[1].localeCompare(matchRight[1], "it", { numeric: true });
    if (prefixCompare !== 0) return prefixCompare;
    const leftNum = Number(matchLeft[2]);
    const rightNum = Number(matchRight[2]);
    if (Number.isFinite(leftNum) && Number.isFinite(rightNum)) {
      return leftNum - rightNum;
    }
  }
  return left.localeCompare(right, "it", { numeric: true });
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
    const itMatch = trimmed.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
    if (itMatch) {
      const day = Number(itMatch[1]);
      const month = Number(itMatch[2]);
      let year = Number(itMatch[3]);
      if (year < 100) year += 2000;
      const date = new Date(year, month - 1, day);
      if (
        date.getFullYear() === year &&
        date.getMonth() === month - 1 &&
        date.getDate() === day
      ) {
        return date.getTime();
      }
    }
  }
  const parsed = new Date(val as any);
  return Number.isNaN(parsed.getTime()) ? 0 : parsed.getTime();
};

const toInputDate = (val: unknown): string => {
  const t = getTimeValue(val);
  if (!t) return "";
  return new Date(t).toLocaleDateString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
};

const toIsoOrNull = (val?: string): string | null => {
  if (!val) return null;
  const trimmed = val.trim();
  if (!trimmed) return null;
  const itMatch = trimmed.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (itMatch) {
    const day = Number(itMatch[1]);
    const month = Number(itMatch[2]);
    let year = Number(itMatch[3]);
    if (year < 100) year += 2000;
    const yyyy = String(year).padStart(4, "0");
    const mm = String(month).padStart(2, "0");
    const dd = String(day).padStart(2, "0");
    const d = new Date(`${yyyy}-${mm}-${dd}`);
    return Number.isNaN(d.getTime()) ? null : d.toISOString();
  }
  const d = new Date(trimmed);
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
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const TUBI_DATE_FIELDS = new Set(["field_3", "field_16", "field_21", "Modified", "Created"]);
const FORGIATI_DATE_FIELDS = new Set(["field_2", "field_11", "field_23", "Modified", "Created"]);
const SPARK_DATE_FIELDS = new Set<string>();

const toExcelDateSerial = (val: unknown): number | "" => {
  const t = getTimeValue(val);
  if (!t) return "";
  const d = new Date(t);
  const utcMidnight = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return utcMidnight / 86400000 + 25569;
};

const toExcelCellValueForDateFields = (
  dateFields: Set<string>,
  fieldKey: string | null,
  value: unknown
): string | number | boolean | null => {
  if (!fieldKey) return "";
  if (dateFields.has(fieldKey)) {
    return toExcelDateSerial(value);
  }
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
};

const toExcelCellValue = (fieldKey: string | null, value: unknown): string | number | boolean | null => {
  return toExcelCellValueForDateFields(TUBI_DATE_FIELDS, fieldKey, value);
};

const toForgiatiExcelCellValue = (fieldKey: string | null, value: unknown): string | number | boolean | null => {
  return toExcelCellValueForDateFields(FORGIATI_DATE_FIELDS, fieldKey, value);
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
  map.set(normalizeExcelKey("GRADO MATERIALE 2"), "GRADOMATERIALE2");
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

const buildForgiatiExcelColumnMap = () => {
  const map = new Map<string, string>();
  map.set(normalizeExcelKey("CODICE"), "Title");
  map.set(normalizeExcelKey("CODICE SAM"), "CodiceSAM");
  map.set(normalizeExcelKey("NORDINE"), "field_1");
  map.set(normalizeExcelKey("N ORDINE"), "field_1");
  map.set(normalizeExcelKey("DATA ORD"), "field_2");
  map.set(normalizeExcelKey("DATA ORDINE"), "field_2");
  map.set(normalizeExcelKey("FORNITORE"), "field_3");
  map.set(normalizeExcelKey("POS"), "field_4");
  map.set(normalizeExcelKey("QTA"), "field_5");
  map.set(normalizeExcelKey("QTA."), "field_5");
  map.set(normalizeExcelKey("DN"), "field_6");
  map.set(normalizeExcelKey("CLASSE"), "field_7");
  map.set(normalizeExcelKey("NO DISEGNO PARTICOLARE"), "field_8");
  map.set(normalizeExcelKey("NO. DISEGNO - PARTICOLARE"), "field_8");
  map.set(normalizeExcelKey("GRADO MATERIALE"), "field_9");
  map.set(normalizeExcelKey("GRADO MATERIALE 2"), "GRADOMATERIALE2");
  map.set(normalizeExcelKey("N BOLLA"), "field_10");
  map.set(normalizeExcelKey("N BOLLA."), "field_10");
  map.set(normalizeExcelKey("DATA CONSEGNA"), "field_11");
  map.set(normalizeExcelKey("N CERT"), "field_12");
  map.set(normalizeExcelKey("N CERT."), "field_12");
  map.set(normalizeExcelKey("N COLATA"), "field_13");
  map.set(normalizeExcelKey("TIPO CERTIFICAZIONE"), "field_14");
  map.set(normalizeExcelKey("PREZ C D"), "field_15");
  map.set(normalizeExcelKey("PREZ C/D"), "field_15");
  map.set(normalizeExcelKey("Ø EST MM"), "field_16");
  map.set(normalizeExcelKey("O EST MM"), "field_16");
  map.set(normalizeExcelKey("Ø INT MM"), "field_17");
  map.set(normalizeExcelKey("O INT MM"), "field_17");
  map.set(normalizeExcelKey("H ALTEZ MM"), "field_18");
  map.set(normalizeExcelKey("ANELLO DISCO"), "field_19");
  map.set(normalizeExcelKey("GREZZO SGROSSATO"), "field_20");
  map.set(normalizeExcelKey("GIACENZA MM LUNGHEZZA DA BARRA"), "field_21");
  map.set(normalizeExcelKey("GIACENZA QTA"), "field_22");
  map.set(normalizeExcelKey("GIACENZA Q.TA"), "field_22");
  map.set(normalizeExcelKey("GIACENZA Q.TÀ"), "field_22");
  map.set(normalizeExcelKey("DATA PRELIEVO"), "field_23");
  map.set(normalizeExcelKey("COMMESSA"), "field_24");
  map.set(normalizeExcelKey("NOTE"), "field_25");
  map.set(normalizeExcelKey("IDENTLOTTO"), "IdentLotto");
  map.set(normalizeExcelKey("IDENT LOTTO"), "IdentLotto");
  return map;
};

const FORGIATI_EXCEL_COLUMN_MAP = buildForgiatiExcelColumnMap();

const buildSparkExcelColumnMap = () => {
  const map = new Map<string, string>();
  map.set(normalizeExcelKey("CODICE"), "Title");
  map.set(normalizeExcelKey("TITLE"), "Title");
  map.set(normalizeExcelKey("LOTTO"), "field_1");
  map.set(normalizeExcelKey("CODICE SAM"), "field_2");
  map.set(normalizeExcelKey("TIPOLOGIA ARTICOLO"), "field_3");
  map.set(normalizeExcelKey("N ORDINE"), "field_4");
  map.set(normalizeExcelKey("N. ORDINE"), "field_4");
  map.set(normalizeExcelKey("DATA ORDINE"), "field_5");
  map.set(normalizeExcelKey("FORNITORE"), "field_6");
  map.set(normalizeExcelKey("QUANTITA ORDINATA"), "field_7");
  map.set(normalizeExcelKey("QUANTITÀ ORDINATA"), "field_7");
  map.set(normalizeExcelKey("N BOLLA"), "field_8");
  map.set(normalizeExcelKey("N. BOLLA"), "field_8");
  map.set(normalizeExcelKey("DATA CONSEGNA"), "field_9");
  map.set(normalizeExcelKey("GIACENZA"), "field_10");
  map.set(normalizeExcelKey("DATA ULTIMO PRELIEVO"), "field_11");
  map.set(normalizeExcelKey("PREZZO UNITARIO"), "field_12");
  map.set(normalizeExcelKey("COMMESSA"), "field_13");
  return map;
};

const SPARK_EXCEL_COLUMN_MAP = buildSparkExcelColumnMap();

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

const buildForgiatiExcelRow = (excelColumns: string[], fields: Record<string, unknown>) => {
  const fieldKeyLookup = new Map<string, string>();
  Object.keys(fields).forEach((key) => {
    fieldKeyLookup.set(normalizeExcelKey(key), key);
  });

  return excelColumns.map((columnName) => {
    const normalized = normalizeExcelKey(columnName || "");
    const fieldKey =
      FORGIATI_EXCEL_COLUMN_MAP.get(normalized) || fieldKeyLookup.get(normalized) || null;
    const rawValue = fieldKey ? fields[fieldKey] : null;
    return toForgiatiExcelCellValue(fieldKey, rawValue);
  });
};

const buildSparkExcelRow = (excelColumns: string[], fields: Record<string, unknown>) => {
  const fieldKeyLookup = new Map<string, string>();
  Object.keys(fields).forEach((key) => {
    fieldKeyLookup.set(normalizeExcelKey(key), key);
  });

  return excelColumns.map((columnName) => {
    const normalized = normalizeExcelKey(columnName || "");
    const fieldKey =
      SPARK_EXCEL_COLUMN_MAP.get(normalized) || fieldKeyLookup.get(normalized) || null;
    const rawValue = fieldKey ? fields[fieldKey] : null;
    return toExcelCellValueForDateFields(SPARK_DATE_FIELDS, fieldKey, rawValue);
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

const getForgiatiExcelColumnIndex = (excelColumns: string[], fieldKey: string) => {
  const target = normalizeExcelKey(fieldKey);
  for (let i = 0; i < excelColumns.length; i++) {
    const normalized = normalizeExcelKey(excelColumns[i] || "");
    const mapped = FORGIATI_EXCEL_COLUMN_MAP.get(normalized);
    if (mapped && normalizeExcelKey(mapped) === target) {
      return i;
    }
  }
  return null;
};

const getSparkExcelColumnIndex = (excelColumns: string[], fieldKey: string) => {
  const target = normalizeExcelKey(fieldKey);
  for (let i = 0; i < excelColumns.length; i++) {
    const normalized = normalizeExcelKey(excelColumns[i] || "");
    const mapped = SPARK_EXCEL_COLUMN_MAP.get(normalized);
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

const findForgiatiExcelRowIndex = (
  rows: Array<{ index: number; values: Array<Array<unknown>> }>,
  excelColumns: string[],
  options: { codice: string; identLotto: string }
) => {
  const titleIdx = getForgiatiExcelColumnIndex(excelColumns, "Title");
  const identIdx = getForgiatiExcelColumnIndex(excelColumns, "IdentLotto");
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

const findSparkExcelRowIndex = (
  rows: Array<{ index: number; values: Array<Array<unknown>> }>,
  excelColumns: string[],
  options: { codice: string; lotto?: string }
) => {
  const titleIdx = getSparkExcelColumnIndex(excelColumns, "Title");
  if (titleIdx === null) return null;
  const lottoIdx = getSparkExcelColumnIndex(excelColumns, "field_1");

  const targetCodice = normalizeExcelKey(options.codice || "");
  const targetLotto = normalizeExcelKey(options.lotto || "");

  for (const row of rows) {
    const rowValues = row.values?.[0] || [];
    const codiceVal = normalizeExcelKey(String(rowValues[titleIdx] ?? ""));
    if (codiceVal !== targetCodice) continue;
    if (lottoIdx === null) return row.index;
    const lottoVal = normalizeExcelKey(String(rowValues[lottoIdx] ?? ""));
    if (lottoVal === targetLotto || (!lottoVal && !targetLotto)) {
      return row.index;
    }
  }

  return null;
};

const findLastRowIndexByCodice = (
  rows: Array<{ index: number; values: Array<Array<unknown>> }>,
  excelColumns: string[],
  codice: string,
  getColumnIndex: (columns: string[], fieldKey: string) => number | null
) => {
  const titleIdx = getColumnIndex(excelColumns, "Title");
  if (titleIdx === null) return null;
  const targetCodice = normalizeExcelKey(codice || "");
  let lastIndex: number | null = null;
  rows.forEach((row) => {
    const rowValues = row.values?.[0] || [];
    const codiceVal = normalizeExcelKey(String(rowValues[titleIdx] ?? ""));
    if (codiceVal === targetCodice) {
      if (lastIndex === null || row.index > lastIndex) {
        lastIndex = row.index;
      }
    }
  });
  return lastIndex;
};

const getNextLottoProg = (items: Array<{ id: string }>, map: Map<string, string>, lastUsedLetter?: string | null) => {
  let maxCode = 64; // one before 'A'
  items.forEach((itm) => {
    const currentProg = map.get(itm.id);
    if (currentProg) {
      const prog = formatLottoProg(currentProg);
      const code = prog.charCodeAt(0);
      if (code >= 65 && code <= 90 && code > maxCode) {
        maxCode = code;
      }
    }
  });
  if (lastUsedLetter) {
    const letterCode = formatLottoProg(lastUsedLetter).charCodeAt(0);
    if (letterCode >= 65 && letterCode <= 90 && letterCode > maxCode) {
      maxCode = letterCode;
    }
  }
  const nextCode = Math.min(90, maxCode + 1);
  return String.fromCharCode(nextCode < 65 ? 65 : nextCode);
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
  { kind: "SPARK-GUPS", label: "6_SPARK GUPS" },
];

const getFieldSet = (kind: ListKind) => {
  if (kind === "FORGIATI") return FORGIATI_FIELDS;
  if (kind === "TUBI") return TUBI_FIELDS;
  if (kind === "ORING-HNBR") return ORING_HNBR_FIELDS;
  if (kind === "SPARK-GUPS") return SPARK_GUPS_FIELDS;
  return ORING_NBR_FIELDS;
};

const getListNoun = (kind: ListKind) => {
  if (kind === "FORGIATI") return "forgiato";
  if (kind === "TUBI") return "tubo";
  if (kind === "ORING-HNBR") return "oring HNBR";
  if (kind === "SPARK-GUPS") return "spark gups";
  return "oring NBR";
};

const getSortDateKey = (kind: ListKind) => {
  if (kind === "FORGIATI") return "field_2";
  if (kind === "TUBI") return "field_3";
  if (kind === "ORING-HNBR") return "field_19";
  if (kind === "SPARK-GUPS") return "field_5";
  return "field_17";
};

const getSearchKeys = (kind: ListKind) => {
  if (kind === "FORGIATI") return ["Title", "field_10", "field_13"];
  if (kind === "TUBI") return ["Title", "field_15", "field_18"];
  if (kind === "ORING-HNBR") return ["Title", "field_18", "field_12", "field_2"];
  if (kind === "SPARK-GUPS") return ["Title", "field_1", "field_4", "field_8"];
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
      { label: "Lotto", value: formatLottoProg(options.lottoProg || toStr(fields.LottoProgressivo) || "A") },
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
  if (kind === "SPARK-GUPS") {
    return [
      { label: "Lotto", value: toStr(fields.field_1) || "-" },
      { label: "Ordine", value: toStr(fields.field_4) || "-" },
      { label: "Bolla", value: toStr(fields.field_8) || "-" },
      { label: "Giacenza", value: toStr(fields.field_10) || "-" },
    ];
  }
  return [
    { label: "Commessa", value: toStr(fields.field_11) || "-" },
    { label: "Prenotazione", value: toStr(fields.field_15) || "-" },
    { label: "Giacenza", value: toStr(fields.field_16) || "-" },
  ];
};

export function AdminPanel({
  siteId,
  forgiatiListId,
  tubiListId,
  oringHnbrListId,
  oringNbrListId,
  sparkGupsListId,
}: AdminPanelProps) {
  const getClient = useAuthenticatedGraphClient();
  const forgiatiExcelPath = (import.meta.env.VITE_FORGIATI_EXCEL_PATH || "").trim();
  const forgiatiExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const forgiatiExcelFilename = (import.meta.env.VITE_SP_FORGIATI_FILENAME || import.meta.env.VITE_FORGIATI_EXCEL_FILE || "").trim();
  const forgiatiExcelTable = (import.meta.env.VITE_FORGIATI_EXCEL_TABLE || "tblFORGIATI").trim();
  const forgiatiExcelDriveIdEnv = (import.meta.env.VITE_FORGIATI_EXCEL_DRIVE_ID || "").trim();
  const forgiatiExcelDriveNameEnv = (import.meta.env.VITE_FORGIATI_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const [forgiatiExcelDriveId, setForgiatiExcelDriveId] = useState<string | null>(forgiatiExcelDriveIdEnv || null);
  const tubiExcelPath = (import.meta.env.VITE_TUBI_EXCEL_PATH || "").trim();
  const tubiExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const tubiExcelFilename = (import.meta.env.VITE_SP_TUBI_FILENAME || import.meta.env.VITE_TUBI_EXCEL_FILE || "").trim();
  const tubiExcelTable = (import.meta.env.VITE_TUBI_EXCEL_TABLE || "tblTUBI").trim();
  const tubiExcelDriveIdEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_ID || "").trim();
  const tubiExcelDriveNameEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const [tubiExcelDriveId, setTubiExcelDriveId] = useState<string | null>(tubiExcelDriveIdEnv || null);
  const sparkExcelPath = (import.meta.env.VITE_SPARK_GUPS_EXCEL_PATH || "").trim();
  const sparkExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const sparkExcelFilename = (import.meta.env.VITE_SP_SPARK_GUPS_FILENAME || import.meta.env.VITE_SPARK_GUPS_EXCEL_FILE || "").trim();
  const sparkExcelTable = (import.meta.env.VITE_SPARK_GUPS_EXCEL_TABLE || "tblSPARKGUPS").trim();
  const sparkExcelDriveIdEnv = (import.meta.env.VITE_SPARK_GUPS_EXCEL_DRIVE_ID || "").trim();
  const sparkExcelDriveNameEnv = (import.meta.env.VITE_SPARK_GUPS_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const [sparkExcelDriveId, setSparkExcelDriveId] = useState<string | null>(sparkExcelDriveIdEnv || null);

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

  const {
    data: sparkGupsItems,
    loading: sparkGupsLoading,
    error: sparkGupsError,
    refresh: refreshSparkGups,
  } = useCachedList<Record<string, unknown>>(service, sparkGupsListId, "admin-spark-gups");

  const [activeList, setActiveList] = useState<ListKind>("FORGIATI");
  const [search, setSearch] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<FormState>(() => buildFormFromItem(null, FORGIATI_FIELDS));
  const [createForm, setCreateForm] = useState<FormState>(() => buildFormFromItem(null, FORGIATI_FIELDS));
  const [updateStatus, setUpdateStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [updateMessage, setUpdateMessage] = useState<string | null>(null);
  const [deleteStatus, setDeleteStatus] = useState<"idle" | "deleting" | "success" | "error">("idle");
  const [deleteMessage, setDeleteMessage] = useState<string | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);
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
      : activeList === "SPARK-GUPS"
      ? sparkGupsItems
      : oringNbrItems;
  const activeLoading =
    activeList === "FORGIATI"
      ? forgiatiLoading
      : activeList === "TUBI"
      ? tubiLoading
      : activeList === "ORING-HNBR"
      ? oringHnbrLoading
      : activeList === "SPARK-GUPS"
      ? sparkGupsLoading
      : oringNbrLoading;
  const activeError =
    activeList === "FORGIATI"
      ? forgiatiError
      : activeList === "TUBI"
      ? tubiError
      : activeList === "ORING-HNBR"
      ? oringHnbrError
      : activeList === "SPARK-GUPS"
      ? sparkGupsError
      : oringNbrError;
  const activeListId =
    activeList === "FORGIATI"
      ? forgiatiListId
      : activeList === "TUBI"
      ? tubiListId
      : activeList === "ORING-HNBR"
      ? oringHnbrListId
      : activeList === "SPARK-GUPS"
      ? sparkGupsListId
      : oringNbrListId;
  const activeRefresh =
    activeList === "FORGIATI"
      ? refreshForgiati
      : activeList === "TUBI"
      ? refreshTubi
      : activeList === "ORING-HNBR"
      ? refreshOringHnbr
      : activeList === "SPARK-GUPS"
      ? refreshSparkGups
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
    if (activeList === "TUBI" || activeList === "FORGIATI") {
      items.sort((a, b) => {
        const left = toStr((a.fields as Record<string, unknown>).Title);
        const right = toStr((b.fields as Record<string, unknown>).Title);
        return compareTubiTitle(right, left);
      });
      return items;
    }
    const dateKey = getSortDateKey(activeList);
    items.sort(
      (a, b) =>
        getTimeValue((b.fields as Record<string, unknown>)[dateKey]) -
        getTimeValue((a.fields as Record<string, unknown>)[dateKey])
    );
    return items;
  }, [activeItems, activeList]);

  const forgiatiGroupedRows = useMemo(() => {
    const map = new Map<string, SharePointListItem<Record<string, unknown>>[]>();
    (forgiatiItems || []).forEach((item) => {
      const title = toStr((item.fields as Record<string, unknown>).Title) || "(senza codice)";
      const existing = map.get(title) || [];
      map.set(title, [...existing, item]);
    });
    return Array.from(map.entries()).map(([title, items]) => ({ title, items }));
  }, [forgiatiItems]);

  const forgiatiProgressiveMap = useMemo(() => {
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

    forgiatiGroupedRows.forEach((group) => {
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
  }, [forgiatiGroupedRows]);

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
    setDeleteStatus("idle");
    setDeleteMessage(null);
    setDeletingId(null);
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
      if (activeList === "TUBI" || activeList === "FORGIATI" || activeList === "SPARK-GUPS") {
        try {
          const isForgiati = activeList === "FORGIATI";
          const isSpark = activeList === "SPARK-GUPS";
          const resolvedPath = isForgiati
            ? forgiatiExcelPath || (forgiatiExcelFolder && forgiatiExcelFilename ? `${forgiatiExcelFolder}/${forgiatiExcelFilename}` : "")
            : isSpark
            ? sparkExcelPath || (sparkExcelFolder && sparkExcelFilename ? `${sparkExcelFolder}/${sparkExcelFilename}` : "")
            : tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
          let resolvedDriveId = isForgiati ? forgiatiExcelDriveId : isSpark ? sparkExcelDriveId : tubiExcelDriveId;
          const driveNameEnv = isForgiati ? forgiatiExcelDriveNameEnv : isSpark ? sparkExcelDriveNameEnv : tubiExcelDriveNameEnv;
          if (!resolvedDriveId && driveNameEnv) {
            resolvedDriveId = await service.getDriveIdByName(driveNameEnv);
            if (isForgiati) {
              setForgiatiExcelDriveId(resolvedDriveId);
            } else if (isSpark) {
              setSparkExcelDriveId(resolvedDriveId);
            } else {
              setTubiExcelDriveId(resolvedDriveId);
            }
          }
          if (!resolvedDriveId && driveNameEnv) {
            throw new Error(`Libreria "${driveNameEnv}" non trovata`);
          }
          const tableName = isForgiati ? forgiatiExcelTable : isSpark ? sparkExcelTable : tubiExcelTable;
          if (!resolvedPath || !tableName) {
            throw new Error("Percorso Excel o tabella non configurati");
          }

          const currentItem = filteredItems.find((item) => item.id === selectedId) || null;
          const identRaw =
            (currentItem?.fields as any)?.LottoProgressivo ||
            (currentItem?.id
              ? isForgiati
                ? forgiatiProgressiveMap.get(currentItem.id)
                : tubiProgressiveMap.get(currentItem.id)
              : undefined);
          const identLotto = formatLottoProg(identRaw || "A");
          const lookupTitle = toStr((currentItem?.fields as any)?.Title) || editForm.Title || "";

          const driveItem = await service.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await service.listWorkbookTableColumnsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);
          const rows = await service.listWorkbookTableRowsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);
          const rowIndex = isForgiati
            ? findForgiatiExcelRowIndex(rows, columns, { codice: lookupTitle, identLotto })
            : isSpark
            ? findSparkExcelRowIndex(rows, columns, { codice: lookupTitle, lotto: toStr(editForm.field_1) })
            : findExcelRowIndex(rows, columns, { codice: lookupTitle, identLotto });

          if (rowIndex === null) {
            throw new Error(
              `Riga Excel non trovata per ${editForm.Title || "codice"}${isSpark ? "" : ` (${identLotto})`}`
            );
          }

          const dataBodyRange = await service.getWorkbookTableDataBodyRangeByItemId(
            driveItem.id,
            tableName,
            resolvedDriveId || undefined
          );
          const rowRange = buildRowRangeAddress(dataBodyRange.address, rowIndex, dataBodyRange.rowCount);
          if (!rowRange) {
            throw new Error("Impossibile calcolare la riga Excel da aggiornare");
          }

          const excelFields = {
            ...(currentItem?.fields || {}),
            ...payload,
            ...(isSpark ? {} : { IdentLotto: identLotto }),
          } as Record<string, unknown>;
          const rowValues = isForgiati
            ? buildForgiatiExcelRow(columns, excelFields)
            : isSpark
            ? buildSparkExcelRow(columns, excelFields)
            : buildTubiExcelRow(columns, excelFields);
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
              console.warn("Errore chiusura sessione Excel (admin)", closeErr);
            }
          }
        } catch (excelErr: any) {
          console.error("Errore aggiornamento Excel (admin)", excelErr);
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
    sparkExcelPath,
    sparkExcelFolder,
    sparkExcelFilename,
    sparkExcelTable,
    sparkExcelDriveId,
    sparkExcelDriveNameEnv,
  ]);

  const handleDelete = useCallback(async (item: SharePointListItem<Record<string, unknown>>) => {
    if (!service) {
      setDeleteStatus("error");
      setDeleteMessage("Servizio SharePoint non configurato");
      return;
    }
    if (!activeListId) {
      setDeleteStatus("error");
      setDeleteMessage("List ID mancante per la lista selezionata");
      return;
    }

    const fields = item.fields as Record<string, unknown>;
    const title = toStr(fields.Title) || "elemento";
    let lottoHint = "";
    if (activeList === "FORGIATI" || activeList === "TUBI") {
      const identRaw =
        (fields as any)?.LottoProgressivo ||
        (item.id
          ? activeList === "FORGIATI"
            ? forgiatiProgressiveMap.get(item.id)
            : tubiProgressiveMap.get(item.id)
          : undefined);
      lottoHint = ` (lotto ${formatLottoProg(identRaw || "A")})`;
    } else if (activeList === "SPARK-GUPS") {
      const lotto = toStr(fields.field_1);
      lottoHint = lotto ? ` (lotto ${lotto})` : "";
    }

    if (!window.confirm(`Eliminare ${getListNoun(activeList)} "${title}"${lottoHint}?`)) {
      return;
    }

    setDeleteStatus("deleting");
    setDeleteMessage(null);
    setDeletingId(item.id);

    try {
      await service.deleteItem(activeListId, item.id);
      let excelError: string | null = null;
      if (activeList === "TUBI" || activeList === "FORGIATI" || activeList === "SPARK-GUPS") {
        try {
          const isForgiati = activeList === "FORGIATI";
          const isSpark = activeList === "SPARK-GUPS";
          const resolvedPath = isForgiati
            ? forgiatiExcelPath || (forgiatiExcelFolder && forgiatiExcelFilename ? `${forgiatiExcelFolder}/${forgiatiExcelFilename}` : "")
            : isSpark
            ? sparkExcelPath || (sparkExcelFolder && sparkExcelFilename ? `${sparkExcelFolder}/${sparkExcelFilename}` : "")
            : tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
          let resolvedDriveId = isForgiati ? forgiatiExcelDriveId : isSpark ? sparkExcelDriveId : tubiExcelDriveId;
          const driveNameEnv = isForgiati ? forgiatiExcelDriveNameEnv : isSpark ? sparkExcelDriveNameEnv : tubiExcelDriveNameEnv;
          if (!resolvedDriveId && driveNameEnv) {
            resolvedDriveId = await service.getDriveIdByName(driveNameEnv);
            if (isForgiati) {
              setForgiatiExcelDriveId(resolvedDriveId);
            } else if (isSpark) {
              setSparkExcelDriveId(resolvedDriveId);
            } else {
              setTubiExcelDriveId(resolvedDriveId);
            }
          }
          if (!resolvedDriveId && driveNameEnv) {
            throw new Error(`Libreria "${driveNameEnv}" non trovata`);
          }
          const tableName = isForgiati ? forgiatiExcelTable : isSpark ? sparkExcelTable : tubiExcelTable;
          if (!resolvedPath || !tableName) {
            throw new Error("Percorso Excel o tabella non configurati");
          }

          const identRaw =
            (fields as any)?.LottoProgressivo ||
            (item.id
              ? isForgiati
                ? forgiatiProgressiveMap.get(item.id)
                : tubiProgressiveMap.get(item.id)
              : undefined);
          const identLotto = formatLottoProg(identRaw || "A");
          const lookupTitle = toStr(fields.Title) || "";
          const sparkLotto = toStr(fields.field_1);

          const driveItem = await service.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await service.listWorkbookTableColumnsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);
          const rows = await service.listWorkbookTableRowsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);
          const rowIndex = isForgiati
            ? findForgiatiExcelRowIndex(rows, columns, { codice: lookupTitle, identLotto })
            : isSpark
            ? findSparkExcelRowIndex(rows, columns, { codice: lookupTitle, lotto: sparkLotto })
            : findExcelRowIndex(rows, columns, { codice: lookupTitle, identLotto });

          if (rowIndex === null) {
            throw new Error(
              `Riga Excel non trovata per ${lookupTitle || "codice"}${isSpark ? "" : ` (${identLotto})`}`
            );
          }

          const dataBodyRange = await service.getWorkbookTableDataBodyRangeByItemId(
            driveItem.id,
            tableName,
            resolvedDriveId || undefined
          );
          const rowRange = buildRowRangeAddress(dataBodyRange.address, rowIndex, dataBodyRange.rowCount);
          if (!rowRange) {
            throw new Error("Impossibile calcolare la riga Excel da eliminare");
          }

          const sessionId = await service.createWorkbookSessionByItemId(
            driveItem.id,
            { persistChanges: true },
            resolvedDriveId || undefined
          );
          try {
            try {
              await service.deleteWorkbookTableRowByIndex(
                driveItem.id,
                tableName,
                rowIndex,
                { sessionId },
                resolvedDriveId || undefined
              );
            } catch (deleteErr) {
              await service.deleteWorkbookRangeByAddress(
                driveItem.id,
                rowRange.sheetName,
                rowRange.address,
                { sessionId, shift: "Up" },
                resolvedDriveId || undefined
              );
            }
          } finally {
            try {
              await service.closeWorkbookSessionByItemId(
                driveItem.id,
                sessionId,
                resolvedDriveId || undefined
              );
            } catch (closeErr) {
              console.warn("Errore chiusura sessione Excel (admin delete)", closeErr);
            }
          }
        } catch (excelErr: any) {
          console.error("Errore eliminazione Excel (admin)", excelErr);
          excelError = excelErr?.message || "Errore eliminazione Excel";
        }
      }

      setDeleteStatus("success");
      setDeleteMessage(
        excelError
          ? `Elemento eliminato; Excel non aggiornato: ${excelError}`
          : "Elemento eliminato con successo"
      );
      if (selectedId === item.id) {
        setSelectedId(null);
      }
      activeRefresh();
    } catch (err: any) {
      setDeleteStatus("error");
      setDeleteMessage(err?.message || "Errore durante l'eliminazione");
    } finally {
      setDeletingId(null);
    }
  }, [
    service,
    activeListId,
    activeList,
    activeRefresh,
    forgiatiExcelPath,
    forgiatiExcelFolder,
    forgiatiExcelFilename,
    forgiatiExcelTable,
    forgiatiExcelDriveId,
    forgiatiExcelDriveNameEnv,
    tubiExcelPath,
    tubiExcelFolder,
    tubiExcelFilename,
    tubiExcelTable,
    tubiExcelDriveId,
    tubiExcelDriveNameEnv,
    sparkExcelPath,
    sparkExcelFolder,
    sparkExcelFilename,
    sparkExcelTable,
    sparkExcelDriveId,
    sparkExcelDriveNameEnv,
    tubiProgressiveMap,
    forgiatiProgressiveMap,
    selectedId,
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
      let excelError: string | null = null;
      if (activeList === "TUBI" || activeList === "FORGIATI" || activeList === "SPARK-GUPS") {
        try {
          const isForgiati = activeList === "FORGIATI";
          const isSpark = activeList === "SPARK-GUPS";
          const resolvedPath = isForgiati
            ? forgiatiExcelPath || (forgiatiExcelFolder && forgiatiExcelFilename ? `${forgiatiExcelFolder}/${forgiatiExcelFilename}` : "")
            : isSpark
            ? sparkExcelPath || (sparkExcelFolder && sparkExcelFilename ? `${sparkExcelFolder}/${sparkExcelFilename}` : "")
            : tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
          let resolvedDriveId = isForgiati ? forgiatiExcelDriveId : isSpark ? sparkExcelDriveId : tubiExcelDriveId;
          const driveNameEnv = isForgiati ? forgiatiExcelDriveNameEnv : isSpark ? sparkExcelDriveNameEnv : tubiExcelDriveNameEnv;
          if (!resolvedDriveId && driveNameEnv) {
            resolvedDriveId = await service.getDriveIdByName(driveNameEnv);
            if (isForgiati) {
              setForgiatiExcelDriveId(resolvedDriveId);
            } else if (isSpark) {
              setSparkExcelDriveId(resolvedDriveId);
            } else {
              setTubiExcelDriveId(resolvedDriveId);
            }
          }
          if (!resolvedDriveId && driveNameEnv) {
            throw new Error(`Libreria "${driveNameEnv}" non trovata`);
          }
          const tableName = isForgiati ? forgiatiExcelTable : isSpark ? sparkExcelTable : tubiExcelTable;
          if (!resolvedPath || !tableName) {
            throw new Error("Percorso Excel o tabella non configurati");
          }

          const titleValue = toStr(payload.Title) || toStr(createForm.Title);
          const driveItem = await service.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await service.listWorkbookTableColumnsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);
          const rows = await service.listWorkbookTableRowsByItemId(driveItem.id, tableName, resolvedDriveId || undefined);

          let rowValues: Array<string | number | boolean | null> = [];
          let insertIndex: number | undefined = undefined;
          if (isSpark) {
            rowValues = buildSparkExcelRow(columns, payload);
            const insertAfter = findLastRowIndexByCodice(rows, columns, titleValue, getSparkExcelColumnIndex);
            insertIndex = insertAfter !== null ? insertAfter + 1 : undefined;
          } else {
            const baseItems = isForgiati ? forgiatiItems : tubiItems;
            const map = isForgiati ? forgiatiProgressiveMap : tubiProgressiveMap;
            const sameTitle = (baseItems || []).filter((item) =>
              normalizeExcelKey(toStr((item.fields as Record<string, unknown>).Title)) ===
              normalizeExcelKey(titleValue)
            );
            const nextProg = formatLottoProg(getNextLottoProg(sameTitle, map));
            const excelFields = {
              ...payload,
              IdentLotto: nextProg,
            } as Record<string, unknown>;
            const identIndex = isForgiati
              ? getForgiatiExcelColumnIndex(columns, "IdentLotto")
              : getExcelColumnIndex(columns, "IdentLotto");
            if (identIndex === null) {
              throw new Error("Colonna IdentLotto non trovata nella tabella Excel");
            }
            rowValues = isForgiati
              ? buildForgiatiExcelRow(columns, excelFields)
              : buildTubiExcelRow(columns, excelFields);
            const insertAfter = findLastRowIndexByCodice(
              rows,
              columns,
              titleValue,
              isForgiati ? getForgiatiExcelColumnIndex : getExcelColumnIndex
            );
            insertIndex = insertAfter !== null ? insertAfter + 1 : undefined;
          }

          const sessionId = await service.createWorkbookSessionByItemId(
            driveItem.id,
            { persistChanges: true },
            resolvedDriveId || undefined
          );
          try {
            try {
              await service.appendWorkbookTableRowByItemId(
                driveItem.id,
                tableName,
                rowValues,
                { sessionId, index: insertIndex },
                resolvedDriveId || undefined
              );
            } catch (appendErr) {
              if (insertIndex !== undefined) {
                console.warn("Inserimento Excel (admin) in posizione fallito, riprovo in coda", appendErr);
                await service.appendWorkbookTableRowByItemId(
                  driveItem.id,
                  tableName,
                  rowValues,
                  { sessionId },
                  resolvedDriveId || undefined
                );
              } else {
                throw appendErr;
              }
            }
          } finally {
            try {
              await service.closeWorkbookSessionByItemId(
                driveItem.id,
                sessionId,
                resolvedDriveId || undefined
              );
            } catch (closeErr) {
              console.warn("Errore chiusura sessione Excel (admin create)", closeErr);
            }
          }
        } catch (excelErr: any) {
          console.error("Errore creazione Excel (admin)", excelErr);
          excelError = excelErr?.message || "Errore creazione Excel";
        }
      }
      setCreateStatus("success");
      setCreateMessage(
        excelError
          ? `Nuovo elemento creato; Excel non aggiornato: ${excelError}`
          : "Nuovo elemento creato"
      );
      setCreateForm(buildFormFromItem(null, fields));
      setIsCreateOpen(false);
      activeRefresh();
    } catch (err: any) {
      setCreateStatus("error");
      const detail = err?.body?.error?.message || err?.message;
      setCreateMessage(detail || "Errore durante la creazione");
    }
  }, [
    service,
    activeListId,
    createForm,
    activeRefresh,
    activeList,
    forgiatiExcelPath,
    forgiatiExcelFolder,
    forgiatiExcelFilename,
    forgiatiExcelTable,
    forgiatiExcelDriveId,
    forgiatiExcelDriveNameEnv,
    tubiExcelPath,
    tubiExcelFolder,
    tubiExcelFilename,
    tubiExcelTable,
    tubiExcelDriveId,
    tubiExcelDriveNameEnv,
    sparkExcelPath,
    sparkExcelFolder,
    sparkExcelFilename,
    sparkExcelTable,
    sparkExcelDriveId,
    sparkExcelDriveNameEnv,
    forgiatiItems,
    tubiItems,
    forgiatiProgressiveMap,
    tubiProgressiveMap,
  ]);

  const renderField = (field: FieldConfig, form: FormState, onChange: (key: string, val: string) => void) => {
    const isDateField = field.type === "date";
    const commonProps = {
      value: form[field.key] ?? "",
      onChange: (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => onChange(field.key, e.target.value),
      placeholder: field.placeholder ?? (isDateField ? "gg/mm/aaaa" : undefined),
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
          type={field.type === "number" ? "number" : "text"}
          inputMode={isDateField ? "numeric" : undefined}
          {...commonProps}
        />
      </label>
    );
  };

  const activeSummary = (item: SharePointListItem<Record<string, unknown>>) => {
    const fields = item.fields as Record<string, unknown>;
    const lottoProg =
      activeList === "TUBI"
        ? item.id
          ? tubiProgressiveMap.get(item.id)
          : undefined
        : activeList === "FORGIATI"
        ? item.id
          ? forgiatiProgressiveMap.get(item.id)
          : undefined
        : undefined;
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
              {deleteMessage && (
                <div className={`alert ${deleteStatus === "success" ? "success" : deleteStatus === "error" ? "error" : "warning"}`}>
                  {deleteMessage}
                </div>
              )}
              {filteredItems.length === 0 && (
                <div className="muted" style={{ padding: "12px 6px" }}>
                  Nessun elemento trovato. Verifica il filtro o aggiorna.
                </div>
              )}
              {visibleItems.map((item) => {
                const summary = activeSummary(item);
                const isSelected = selectedId === item.id;
                return (
                  <div
                    key={item.id}
                    className={`admin-list-item ${isSelected ? "selected" : ""}`}
                    onClick={() => setSelectedId(item.id)}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" || e.key === " ") {
                        e.preventDefault();
                        setSelectedId(item.id);
                      }
                    }}
                  >
                    <div className="admin-list-item__head">
                      <div className="admin-list-item__title">{summary.title}</div>
                      <button
                        className="icon-btn danger"
                        type="button"
                        aria-label={`Elimina ${summary.title}`}
                        title="Elimina"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDelete(item);
                        }}
                        disabled={deleteStatus === "deleting" && deletingId === item.id}
                      >
                        🗑️
                      </button>
                    </div>
                    <div className="admin-list-item__meta">
                      {summary.meta.map((entry) => (
                        <span key={entry.label} className="pill ghost">
                          {entry.label} {entry.value}
                        </span>
                      ))}
                      {summary.modified && <span className="pill ghost">Mod. {summary.modified}</span>}
                    </div>
                  </div>
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
