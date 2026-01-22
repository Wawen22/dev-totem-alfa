import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal, MsalProvider } from "@azure/msal-react";
import { EventType, PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./auth/authConfig";
import { useAuthenticatedGraphClient } from "./hooks/useAuthenticatedGraphClient";
import { SharePointService } from "./services/sharePointService";
import { forgiatiColumns, ForgiatoColumn } from "./config/forgiatiColumns";
import { tubiColumns } from "./config/tubiColumns";
import { oringHnbrColumns } from "./config/oringHnbrColumns";
import { oringNbrColumns } from "./config/oringNbrColumns";
import { SharePointListItem } from "./types/sharepoint";
import { formatSharePointDate } from "./utils/dateUtils";
import { WebsiteViewer } from "./components/features/WebsiteViewer";
import { DocumentBrowser } from "./components/features/DocumentBrowser";
import { useIsAdmin } from "./hooks/useIsAdmin";
import { AdminPanel } from "./components/features/AdminPanel";
import { PowerAutomateService } from "./services/powerAutomateService";

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance
  .initialize()
  .then(() => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }

    msalInstance.addEventCallback((event) => {
      if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
        const payload = event.payload as any;
        msalInstance.setActiveAccount(payload.account);
      }
    });
  })
  .catch((error) => console.error("MSAL init error", error));

type ForgiatoItem = SharePointListItem<Record<string, unknown>>;
type TubiItem = SharePointListItem<Record<string, unknown>>;

type CartItem = {
  key: string;
  source: "FORGIATI" | "TUBI" | "ORING-HNBR" | "ORING-NBR";
  itemId: string;
  title: string;
  bolla?: unknown;
  colata?: unknown;
  lottoProg?: string;
  fields: Record<string, unknown>;
};

type EditableItemState = {
  type: CartItem["source"];
  title: string;
  commessa?: string;
  dataPrelievo?: string;
  giacenzaQt?: string;
  note?: string;
  giacenzaBarra?: string;
  dataUltimoPrelievo?: string;
  giacenzaBib?: string;
  giacenzaContab?: string;
  giacenza?: string;
  prenotazione?: string;
  nrOrdine?: string;
  dataOrdine?: string;
  qtaMinima?: string;
  bolla?: string;
  colata?: string;
  raw?: Record<string, unknown>;
};

const getTimeValue = (val: unknown): number => {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") {
    // Excel serial
    return (val - 25569) * 86400 * 1000;
  }
  const d = new Date(String(val));
  return isNaN(d.getTime()) ? 0 : d.getTime();
};

const toInputDate = (val: unknown): string => {
  const t = getTimeValue(val);
  if (!t) return "";
  const iso = new Date(t).toISOString();
  return iso.slice(0, 10);
};

const toStr = (val: unknown): string => {
  if (val === null || val === undefined) return "";
  return String(val);
};

const toIsoOrNull = (val: string | undefined): string | null => {
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d.toISOString();
};

const toNumberOrNull = (val: string | undefined): number | null => {
  if (val === undefined || val === null || val === "") return null;
  const n = Number(val);
  return Number.isFinite(n) ? n : null;
};

// Sum numbers written as a "+" separated expression (e.g. "100+20+10").
const sumAdditiveExpression = (val: string | undefined): number | null => {
  if (val === undefined || val === null) return null;
  const cleaned = val.replace(/\s+/g, "");
  if (!cleaned) return null;

  const parts = cleaned.split("+").filter((p) => p.length > 0);
  if (parts.length === 0) return null;

  let total = 0;
  for (const part of parts) {
    const num = Number(part.replace(",", "."));
    if (!Number.isFinite(num)) {
      return null;
    }
    total += num;
  }

  return total;
};

const formatCellValue = (value: unknown, type?: "text" | "date" | "number") => {
  if (type === "date") {
    return formatSharePointDate(value);
  }
  if (value === null || value === undefined) return "";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
};

const formatLottoProg = (val: string | undefined | null) => {
  const str = val ? String(val) : "";
  return str ? str.toUpperCase() : "A";
};

const normalizeExcelKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const buildExcelColumnFieldMap = (columns: ForgiatoColumn[]) => {
  const map = new Map<string, string>();
  columns.forEach((col) => {
    map.set(normalizeExcelKey(col.field), col.field);
    map.set(normalizeExcelKey(col.label), col.field);
  });
  return map;
};

const tubiExcelColumnFieldMap = (() => {
  const map = buildExcelColumnFieldMap(tubiColumns);
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
  map.set(normalizeExcelKey("LOTTO PROGRESSIVO"), "LottoProgressivo");
  return map;
})();

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

const buildTubiExcelRow = (excelColumns: string[], fields: Record<string, unknown>) => {
  const fieldKeyLookup = new Map<string, string>();
  Object.keys(fields).forEach((key) => {
    fieldKeyLookup.set(normalizeExcelKey(key), key);
  });

  return excelColumns.map((columnName) => {
    const normalized = normalizeExcelKey(columnName || "");
    const fieldKey =
      tubiExcelColumnFieldMap.get(normalized) || fieldKeyLookup.get(normalized) || null;
    const rawValue = fieldKey ? fields[fieldKey] : null;
    return toExcelCellValue(fieldKey, rawValue);
  });
};

const getExcelColumnIndex = (excelColumns: string[], fieldKey: string) => {
  const target = normalizeExcelKey(fieldKey);
  for (let i = 0; i < excelColumns.length; i++) {
    const normalized = normalizeExcelKey(excelColumns[i] || "");
    const mapped = tubiExcelColumnFieldMap.get(normalized);
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

const extractProgLetter = (val: string | undefined | null): string | null => {
  if (!val) return null;
  const match = String(val).trim().match(/[a-z]/i);
  return match ? match[0] : null;
};

const buildColataPlaceholder = (lottoProg?: string | null) => {
  const prefix = lottoProg ? `${formatLottoProg(lottoProg)}-` : "";
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  const timestamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  return `${prefix}${timestamp}`;
};

function DateTimeWidget() {
  const [date, setDate] = useState(new Date());

  useEffect(() => {
    const timer = setInterval(() => setDate(new Date()), 60000); // Update every minute
    return () => clearInterval(timer);
  }, []);

  const dateStr = date.toLocaleDateString("it-IT", {
    weekday: "long",
    day: "numeric",
    month: "long",
  });
  
  const timeStr = date.toLocaleTimeString("it-IT", {
    hour: "2-digit",
    minute: "2-digit",
  });

  return (
    <div className="datetime-widget">
      <div className="time">{timeStr}</div>
      <div className="date">{dateStr}</div>
    </div>
  );
}

interface HeroProps {
  title: string;
  subtitle?: string;
  onBack?: () => void;
}

function Hero({ title, subtitle, onBack }: HeroProps) {
  return (
    <header className="hero">
      <div className="hero-content">
        {onBack && (
          <button className="back-btn" onClick={onBack} aria-label="Indietro">
            ←
          </button>
        )}
        <div className="hero-brand">
          <img src="/imgs/logo.png" alt="Alfa Eng Logo" className="hero-logo" />
          <div className="hero-text">
            <p className="eyebrow">Totem Magazzino</p>
            <h1>{title}</h1>
            {subtitle && <p className="muted">{subtitle}</p>}
          </div>
        </div>
      </div>
      <DateTimeWidget />
    </header>
  );
}

function Dashboard({ onNavigate }: { onNavigate: (view: 'dashboard' | 'inventory' | 'update-stock' | 'website' | 'docs') => void }) {
  return (
    <div className="dashboard-container">
      <button 
        className="dashboard-card primary" 
        onClick={() => onNavigate('inventory')}
      >
        <div className="card-content">
          <h3>Giacenza Magazzino</h3>
          <p>Consulta inventario e ordini in tempo reale</p>
        </div>
        <span className="arrow">→</span>
      </button>

      <button 
        className="dashboard-card primary" 
        onClick={() => onNavigate('docs')}
      >
        <div className="card-content">
          <h3>Consulta Documentazione</h3>
          <p>Manuali, procedure e schede tecniche</p>
        </div>
        <span className="arrow">→</span>
      </button>

      <button 
        className="dashboard-card primary" 
        onClick={() => onNavigate('website')}
      >
        <div className="card-content">
          <h3>Visualizza Sito Web</h3>
          <p>Accesso rapido al portale aziendale</p>
        </div>
        <span className="arrow">→</span>
      </button>
    </div>
  );
}

import { useCachedList } from "./hooks/useCachedList";

// ... (imports remain)

// ... (previous components remain)

interface SelectionStateProps {
  selectedItems: Record<string, CartItem>;
  onToggle: (item: CartItem, isSelected: boolean) => void;
  selectionLimitReached: boolean;
}

function ForgiatiPanel({ selectedItems, onToggle, selectionLimitReached }: SelectionStateProps) {
  const getClient = useAuthenticatedGraphClient();
  const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
  const listId = import.meta.env.VITE_FORGIATI_LIST_ID;

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [lotSelection, setLotSelection] = useState<{
    title: string;
    items: ForgiatoItem[];
  } | null>(null);
  const [newLotColata, setNewLotColata] = useState("");
  const [newLotOrdine, setNewLotOrdine] = useState("");
  const [newLotDataOrdine, setNewLotDataOrdine] = useState("");
  const [newLotCodiceSam, setNewLotCodiceSam] = useState("");
  const [lotSaveStatus, setLotSaveStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [lotSaveMessage, setLotSaveMessage] = useState<string | null>(null);
  const QTA_RESET_VALUE = 0;
  const GIACENZA_QTA_RESET_VALUE = 0;
  const NOTE_RESET_VALUE = "";

  const service = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);

  const { data: rawRows, loading, error, refresh } = useCachedList<Record<string, unknown>>(
    service,
    listId,
    "forgiati"
  );

  const rows = useMemo(() => {
    const sorted = [...rawRows];
    sorted.sort((a, b) => {
      const valA = (a.fields as any).field_2;
      const valB = (b.fields as any).field_2;
      return getTimeValue(valB) - getTimeValue(valA);
    });
    return sorted;
  }, [rawRows]);

  const groupedRows = useMemo(() => {
    const map = new Map<string, ForgiatoItem[]>();
    rows.forEach((item) => {
      const title = toStr((item.fields as Record<string, unknown>).Title) || "(senza codice)";
      const existing = map.get(title) || [];
      map.set(title, [...existing, item]);
    });
    return Array.from(map.entries()).map(([title, items]) => ({
      title,
      items,
      representative: items[0],
    }));
  }, [rows]);

  useEffect(() => {
    if (!lotSelection) return;
    const updated = groupedRows.find((group) => group.title === lotSelection.title);
    if (!updated) return;
    const currentIds = lotSelection.items.map((itm) => itm.id).join(",");
    const nextIds = updated.items.map((itm) => itm.id).join(",");
    if (currentIds !== nextIds) {
      setLotSelection(updated);
    }
  }, [groupedRows, lotSelection]);

  const forgiatiProgressiveMap = useMemo(() => {
    const map = new Map<string, string>();

    const normalizeLetter = (val: string | null | undefined) => {
      if (!val) return null;
      const cleaned = String(val).trim().toLowerCase();
      if (/^[a-z]$/.test(cleaned)) return cleaned;
      return null;
    };

    const getCreatedKey = (item: ForgiatoItem) => {
      const created = getTimeValue((item.fields as any).Created);
      if (created) return created;
      const idNum = Number(item.id);
      return Number.isFinite(idNum) ? idNum : 0;
    };

    groupedRows.forEach((group) => {
      const sorted = [...group.items].sort((a, b) => getCreatedKey(a) - getCreatedKey(b));
      const used = new Set<string>();

      sorted.forEach((item) => {
        const existing = normalizeLetter((item.fields as any).LottoProgressivo);
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
  }, [groupedRows]);

  const toCartItemForgiati = useCallback(
    (item: ForgiatoItem): CartItem => ({
      key: `FORGIATI-${item.id}`,
      source: "FORGIATI",
      itemId: item.id,
      title: toStr((item.fields as Record<string, unknown>).Title) || "-",
      bolla: (item.fields as Record<string, unknown>).field_10,
      colata: (item.fields as Record<string, unknown>).field_13,
      lottoProg: forgiatiProgressiveMap.get(item.id) || "a",
      fields: item.fields as Record<string, unknown>,
    }),
    [forgiatiProgressiveMap]
  );

  const formatCell = (value: unknown, type?: 'text' | 'date' | 'number') => {
    if (type === 'date') {
      return formatSharePointDate(value);
    }
    if (value === null || value === undefined) return "";
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  };

  const visibleColumns = forgiatiColumns;

  const normalizedRows = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return groupedRows;
    return groupedRows.filter((group) =>
      group.items.some((item) =>
        forgiatiColumns.some((col) => {
          const v = (item.fields as Record<string, unknown>)[col.field];
          if (v === null || v === undefined) return false;
          let valStr = String(v);
          if (col.type === "date") {
            valStr = formatSharePointDate(v);
          }
          return valStr.toLowerCase().includes(term);
        })
      )
    );
  }, [groupedRows, search]);

  const totalPages = Math.max(1, Math.ceil(normalizedRows.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageSliceStart = (currentPage - 1) * pageSize;
  const visibleRows = normalizedRows.slice(pageSliceStart, pageSliceStart + pageSize);

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
    setPage(1);
  };

  const handleClearSearch = () => {
    setSearch("");
    setPage(1);
  };

  const handlePrev = () => setPage((p) => Math.max(1, p - 1));
  const handleNext = () => setPage((p) => Math.min(totalPages, p + 1));

  const existingColate = useMemo(() => {
    const set = new Set<string>();
    if (!lotSelection) return set;
    lotSelection.items.forEach((itm) => {
      const val = toStr((itm.fields as any).field_13).trim().toLowerCase();
      if (val) set.add(val);
    });
    return set;
  }, [lotSelection]);

  useEffect(() => {
    if (lotSelection) {
      const nextProg = getNextLottoProg(lotSelection.items, forgiatiProgressiveMap);
      setNewLotColata(buildColataPlaceholder(nextProg));
      setNewLotOrdine("");
      setNewLotDataOrdine("");
      setNewLotCodiceSam("");
      setLotSaveStatus("idle");
      setLotSaveMessage(null);
    }
  }, [lotSelection, forgiatiProgressiveMap]);

  const buildLotClonePayload = (
    item: ForgiatoItem,
    options: {
      colata: string;
      ordine?: string;
      dataOrdine?: string | null;
      codiceSam?: string;
      qtaReset: number | string;
      giacenzaQtReset: number | string;
      noteReset: string;
      dataConsegnaReset: string | null;
    }
  ) => {
    const { colata, ordine, dataOrdine, codiceSam, qtaReset, giacenzaQtReset, noteReset, dataConsegnaReset } = options;
    const sourceFields = (item.fields || {}) as Record<string, unknown>;
    const payload: Record<string, unknown> = {};

    const normalizeDateValue = (val: unknown): string | null => {
      const t = getTimeValue(val);
      if (!t) return null;
      const iso = new Date(t).toISOString();
      return iso;
    };

    const normalizeTextValue = (val: unknown): string | null => {
      if (val === null || val === undefined) return null;
      const trimmed = String(val).trim();
      return trimmed ? trimmed : null;
    };

    const ordineValue = ordine !== undefined ? ordine : toStr((sourceFields as any).field_1);
    const dataOrdineValue = dataOrdine !== undefined ? dataOrdine : normalizeDateValue((sourceFields as any).field_2);
    const qtaValue = String(qtaReset);
    const giacenzaQtValue = String(giacenzaQtReset);
    const noteValue = noteReset;
    const codiceSamValue = normalizeTextValue(codiceSam);
    const dataConsegnaValue = dataConsegnaReset !== undefined ? normalizeDateValue(dataConsegnaReset) : normalizeDateValue((sourceFields as any).field_11);

    Object.entries(sourceFields).forEach(([key, value]) => {
      if (
        key === "id" ||
        key === "ContentType" ||
        key === "ContentTypeId" ||
        key === "Created" ||
        key === "Modified" ||
        key === "AuthorLookupId" ||
        key === "EditorLookupId" ||
        key === "ComplianceAssetId" ||
        key === "GUID" ||
        key === "Attachments" ||
        key === "AppAuthor" ||
        key === "AppEditor" ||
        key === "AppAuthorLookupId" ||
        key === "AppEditorLookupId" ||
        key === "Edit" ||
        key === "ItemChildCount" ||
        key === "FolderChildCount" ||
        key === "_UIVersionString" ||
        key === "UIVersionString" ||
        key === "_ModerationStatus" ||
        key === "_ModerationComments" ||
        key === "_ComplianceFlags" ||
        key === "_ComplianceTag" ||
        key === "_ComplianceTagUserId" ||
        key === "_ComplianceTagWrittenTime" ||
        key === "_IsRecord" ||
        key === "_Level" ||
        key === "_Version" ||
        key === "field_26" ||
        key.startsWith("LinkTitle") ||
        key.startsWith("_Compliance") ||
        key.startsWith("@") ||
        key.startsWith("odata")
      ) {
        return;
      }

      if (key === "field_13") {
        payload[key] = colata;
        return;
      }

      if (key === "field_1") {
        payload[key] = ordineValue;
        return;
      }

      if (key === "field_2") {
        payload[key] = dataOrdineValue;
        return;
      }

      if (key === "field_5") {
        payload[key] = qtaValue;
        return;
      }

      if (key === "field_22") {
        payload[key] = giacenzaQtValue;
        return;
      }

      if (key === "field_25") {
        payload[key] = noteValue;
        return;
      }

      if (key === "CodiceSAM") {
        payload[key] = codiceSam !== undefined ? codiceSamValue : value;
        return;
      }

      if (key === "field_11") {
        // Skip - will be handled after if needed
        return;
      }

      if (key === "LottoProgressivo") {
        payload[key] = null;
        return;
      }

      payload[key] = value;
    });

    if (!("field_13" in payload)) payload.field_13 = colata;
    if (!("field_1" in payload)) payload.field_1 = ordineValue;
    if (!("field_2" in payload)) payload.field_2 = dataOrdineValue;
    if (!("field_5" in payload)) payload.field_5 = qtaValue;
    if (!("field_22" in payload)) payload.field_22 = giacenzaQtValue;
    if (!("field_25" in payload)) payload.field_25 = noteValue;
    // Omit field_11 (Data consegna) entirely - let SharePoint set default or null

    // Remove null/undefined/empty string values to avoid SharePoint errors
    Object.keys(payload).forEach((k) => {
      const v = payload[k];
      if (v === null || v === undefined || v === "") {
        delete payload[k];
      }
    });

    return payload;
  };

  const handleCreateLot = async () => {
    if (!service || !listId || !lotSelection) {
      setLotSaveStatus("error");
      setLotSaveMessage("Servizio non disponibile per la creazione del lotto.");
      return;
    }

    const normalized = newLotColata.trim();
    if (!normalized) return;
    const normalizedLower = normalized.toLowerCase();
    if (existingColate.has(normalizedLower)) {
      setLotSaveStatus("error");
      setLotSaveMessage("Inserisci un N° colata diverso dai lotti esistenti.");
      return;
    }

    const baseItem = lotSelection.items[0];
    if (!baseItem) return;

    setLotSaveStatus("saving");
    setLotSaveMessage(null);
    try {
      const ordineOverride = newLotOrdine.trim();
      const dataOrdineInput = newLotDataOrdine.trim();
      const payload = buildLotClonePayload(baseItem, {
        colata: normalized,
        ordine: ordineOverride,
        dataOrdine: dataOrdineInput ? toIsoOrNull(dataOrdineInput) : null,
        codiceSam: newLotCodiceSam.trim(),
        qtaReset: QTA_RESET_VALUE,
        giacenzaQtReset: GIACENZA_QTA_RESET_VALUE,
        noteReset: NOTE_RESET_VALUE,
        dataConsegnaReset: null,
      });
      await service.createItem<Record<string, unknown>>(listId, payload);
      setLotSaveStatus("success");
      setLotSaveMessage("Nuovo lotto creato. Aggiorno la lista...");
      const fallbackProg = getNextLottoProg(lotSelection.items, forgiatiProgressiveMap, extractProgLetter(normalized));
      setNewLotColata(buildColataPlaceholder(fallbackProg));
      setNewLotOrdine("");
      setNewLotDataOrdine("");
      setNewLotCodiceSam("");
      await refresh();
    } catch (err: any) {
      setLotSaveStatus("error");
      setLotSaveMessage(err?.message || "Errore durante la creazione del lotto.");
    }
  };

  const normalizedNewLot = newLotColata.trim();
  const normalizedNewLotLower = normalizedNewLot.toLowerCase();
  const isDuplicateLot = normalizedNewLot.length > 0 && existingColate.has(normalizedNewLotLower);
  const canCreateLot = Boolean(normalizedNewLot) && !isDuplicateLot && lotSaveStatus !== "saving";

  return (
    <div className="panel inventory-panel">
      <div className="panel-content">
        {error && <div className="alert error">{error}</div>}
        {!listId && (
          <div className="alert warning">Inserisci il GUID della lista FORGIATI in .env.local (VITE_FORGIATI_LIST_ID).</div>
        )}

        <div className="toolbar">
          <div className="search-box">
            <input
              type="search"
              placeholder="Cerca in tutte le colonne..."
              value={search}
              onChange={handleSearchChange}
            />
            {search && <button className="icon-btn" onClick={handleClearSearch} aria-label="Clear search">✕</button>}
          </div>
          
          <div className="pager-info">
             <button className="btn secondary" onClick={() => refresh()} disabled={loading} style={{ padding: '8px 12px', fontSize: '13px' }}>
              {loading ? "..." : "Aggiorna"}
            </button>
            <span className="pill">{normalizedRows.length} risultati</span>
            <label className="page-size">
              <span>Per pagina</span>
              <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}>
                {[100, 200].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </label>
            <div className="pager">
              <button className="icon-btn" onClick={handlePrev} disabled={currentPage === 1}>←</button>
              <span>{currentPage} / {totalPages}</span>
              <button className="icon-btn" onClick={handleNext} disabled={currentPage === totalPages}>→</button>
            </div>
          </div>
        </div>

        <div className="table-scroll">
          <table className="inventory-table">
            <thead>
              <tr>
                <th scope="col" className="selection-col" aria-label="Seleziona">
                  <span className="selection-col__icon" aria-hidden="true"></span>
                </th>
                {visibleColumns.map((col) => (
                  <th
                    scope="col"
                    className={col.field === "Title" ? "sticky-col" : undefined}
                    style={col.width ? { width: col.width } : undefined}
                  >
                    <span>{col.label}</span>
                    <small>{col.field}</small>
                  </th>
                ))}
                <th scope="col" style={{ width: 160 }}>Colate</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.length === 0 && (
                <tr>
                  <td colSpan={visibleColumns.length + 2} className="muted" style={{ textAlign: "center" }}>
                    {loading ? "Carico dati..." : "Nessun elemento trovato"}
                  </td>
                </tr>
              )}
              {visibleRows.map((group) => {
                const representative = group.representative;
                const selectedCount = group.items.filter((itm) => selectedItems[`FORGIATI-${itm.id}`]).length;

                const handleGroupCheck = () => {
                  if (group.items.length === 1) {
                    const singleItem = toCartItemForgiati(group.items[0]);
                    onToggle(singleItem, Boolean(selectedItems[singleItem.key]));
                    return;
                  }
                  setLotSelection(group);
                };

                return (
                  <tr key={group.title}>
                    <td className="selection-col">
                      <input
                        type="checkbox"
                        aria-label="Seleziona riga"
                        checked={selectedCount > 0}
                        disabled={selectionLimitReached && selectedCount === 0 && group.items.length === 1}
                        onChange={handleGroupCheck}
                      />
                    </td>
                    {visibleColumns.map((col) => {
                      const isTitle = col.field === "Title";
                      const className = isTitle ? "sticky-col" : undefined;
                      const content = formatCell((representative.fields as Record<string, unknown>)[col.field], col.type);

                      if (isTitle) {
                        return (
                          <th scope="row" key={col.field} className={className}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "space-between" }}>
                              <span>{content}</span>
                              <button
                                className="icon-btn"
                                aria-label="Gestisci lotti"
                                title="Gestisci lotti"
                                onClick={() => setLotSelection(group)}
                                type="button"
                              >
                                📦
                              </button>
                            </div>
                          </th>
                        );
                      }

                      return (
                        <td key={col.field} className={className}>
                          {content}
                        </td>
                      );
                    })}
                    <td>
                      <div className="colata-badges">
                        {(() => {
                          const sortedLots = [...group.items].sort(
                            (a, b) => getTimeValue((b.fields as any).field_23) - getTimeValue((a.fields as any).field_23)
                          );
                          const latest = sortedLots[0];
                          const moreCount = Math.max(0, sortedLots.length - 1);
                          const latestProg = latest ? forgiatiProgressiveMap.get(latest.id) || "a" : "a";
                          const latestColata = latest ? toStr((latest.fields as any).field_13) || "-" : "-";
                          return (
                            <>
                              <span className="pill ghost" style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                                <span>{latestColata}</span>
                                <span className="lotto-chip">{formatLottoProg(latestProg)}</span>
                              </span>
                              {moreCount > 0 && (
                                <span className="pill ghost" style={{ fontSize: 11, padding: "2px 6px" }}>+{moreCount}</span>
                              )}
                            </>
                          );
                        })()}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {lotSelection && (
        <div className="modal-backdrop blur" role="dialog" aria-modal="true">
          <div className="modal">
            <div className="modal__header">
              <div>
                <p className="eyebrow" style={{ marginBottom: 4 }}>Seleziona lotto</p>
                <h3 style={{ margin: 0 }}>{lotSelection.title}</h3>
              </div>
              <button className="icon-btn" aria-label="Chiudi" onClick={() => setLotSelection(null)} type="button">✕</button>
            </div>
            <div className="modal__body">
              <div className="table-scroll modal-table">
                <table className="inventory-table">
                  <thead>
                    <tr>
                      <th scope="col">Ident. Lotto</th>
                      <th scope="col">N° COLATA</th>
                      <th scope="col">N° ORDINE</th>
                      <th scope="col">Data ordine</th>
                      <th scope="col">Giacenza Q.tà</th>
                      <th scope="col" style={{ width: 140 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {lotSelection.items.map((item) => {
                      const cartItem = toCartItemForgiati(item);
                      const alreadySelected = Boolean(selectedItems[cartItem.key]);
                      const disableAdd = selectionLimitReached && !alreadySelected;
                      const lotProg = forgiatiProgressiveMap.get(item.id) || "a";
                      return (
                        <tr key={item.id}>
                          <td><span className="lotto-chip">{formatLottoProg(lotProg)}</span></td>
                          <td>{toStr((item.fields as any).field_13) || "-"}</td>
                          <td>{toStr((item.fields as any).field_1) || "-"}</td>
                          <td>{formatSharePointDate((item.fields as any).field_2)}</td>
                          <td>{toStr((item.fields as any).field_22) || "-"}</td>
                          <td>
                            <button
                              className={alreadySelected ? "btn secondary" : "btn primary"}
                              style={{ width: "100%", padding: "8px 10px", fontSize: 13 }}
                              onClick={() => {
                                onToggle(cartItem, alreadySelected);
                                setLotSelection(null);
                              }}
                              disabled={disableAdd}
                              type="button"
                            >
                              {alreadySelected ? "Rimuovi" : "Aggiungi"}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="section-heading" style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <span className="pill ghost">Nuovo lotto</span>
                  <p className="muted" style={{ margin: "6px 0 0" }}>Duplico l&apos;articolo</p>
                </div>
              </div>
              <div className="field-grid" style={{ marginTop: 8, alignItems: "flex-end" }}>
                <label className="field" style={{ flex: 1 }}>
                  <span>Nuovo N° COLATA</span>
                  <input
                    type="text"
                    value={newLotColata}
                    onChange={(e) => {
                      setNewLotColata(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Inserisci un valore diverso dai lotti esistenti"
                  />
                </label>
                <label className="field">
                  <span>Nuovo N° ORDINE</span>
                  <input
                    type="text"
                    value={newLotOrdine}
                    onChange={(e) => {
                      setNewLotOrdine(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Facoltativo"
                  />
                </label>
                <label className="field">
                  <span>Codice SAM</span>
                  <input
                    type="text"
                    value={newLotCodiceSam}
                    onChange={(e) => {
                      setNewLotCodiceSam(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Facoltativo"
                  />
                </label>
                <label className="field">
                  <span>Data ordine</span>
                  <input
                    type="date"
                    value={newLotDataOrdine}
                    onChange={(e) => {
                      setNewLotDataOrdine(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                  />
                </label>
                <button
                  className="btn primary"
                  style={{ minWidth: 160, height: 44 }}
                  onClick={handleCreateLot}
                  disabled={!canCreateLot}
                  type="button"
                >
                  {lotSaveStatus === "saving" ? "Creo..." : "Crea lotto"}
                </button>
              </div>
              {newLotColata && isDuplicateLot && (
                <p className="muted" style={{ marginTop: 6 }}>Valore già presente tra le colate esistenti.</p>
              )}
              {lotSaveMessage && (
                <div className={`alert ${lotSaveStatus === "success" ? "success" : lotSaveStatus === "error" ? "error" : "warning"}`} style={{ marginTop: 10 }}>
                  {lotSaveMessage}
                </div>
              )}
              {selectionLimitReached && <p className="muted" style={{ marginTop: 10 }}>Limite massimo di 10 articoli in carrello.</p>}
            </div>
            <div className="modal__footer">
              <button className="btn secondary" onClick={() => setLotSelection(null)} type="button">Chiudi</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function OringHnbrPanel({ selectedItems, onToggle, selectionLimitReached }: SelectionStateProps) {
  const getClient = useAuthenticatedGraphClient();
  const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
  const listId = import.meta.env.VITE_ORING_HNBR_LIST_ID;

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);

  const service = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);

  const { data: rawRows, loading, error, refresh } = useCachedList<Record<string, unknown>>(
    service,
    listId,
    "oring-hnbr"
  );

  const rows = useMemo(() => {
    const sorted = [...rawRows];
    sorted.sort((a, b) => {
      const valA = (a.fields as any).field_19;
      const valB = (b.fields as any).field_19;
      return getTimeValue(valB) - getTimeValue(valA);
    });
    return sorted;
  }, [rawRows]);

  const visibleColumns = oringHnbrColumns;

  const normalizedRows = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return rows;
    return rows.filter((item) =>
      visibleColumns.some((col) => {
        const v = (item.fields as Record<string, unknown>)[col.field];
        if (v === null || v === undefined) return false;

        let valStr = String(v);
        if (col.type === "date") {
          valStr = formatSharePointDate(v);
        }

        return valStr.toLowerCase().includes(term);
      })
    );
  }, [rows, search, visibleColumns]);

  const totalPages = Math.max(1, Math.ceil(normalizedRows.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageSliceStart = (currentPage - 1) * pageSize;
  const visibleRows = normalizedRows.slice(pageSliceStart, pageSliceStart + pageSize);

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
    setPage(1);
  };

  const handleClearSearch = () => {
    setSearch("");
    setPage(1);
  };

  const handlePrev = () => setPage((p) => Math.max(1, p - 1));
  const handleNext = () => setPage((p) => Math.min(totalPages, p + 1));

  return (
    <div className="panel inventory-panel">
      <div className="panel-content">
        {error && <div className="alert error">{error}</div>}
        {!listId && (
          <div className="alert warning">Inserisci il GUID della lista ORING-HNBR in .env.local (VITE_ORING_HNBR_LIST_ID).</div>
        )}

        <div className="toolbar">
          <div className="search-box">
            <input
              type="search"
              placeholder="Cerca in tutte le colonne..."
              value={search}
              onChange={handleSearchChange}
            />
            {search && <button className="icon-btn" onClick={handleClearSearch} aria-label="Clear search">✕</button>}
          </div>
          
          <div className="pager-info">
             <button className="btn secondary" onClick={() => refresh()} disabled={loading} style={{ padding: '8px 12px', fontSize: '13px' }}>
              {loading ? "..." : "Aggiorna"}
            </button>
            <span className="pill">{normalizedRows.length} risultati</span>
            <label className="page-size">
              <span>Per pagina</span>
              <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}>
                {[100, 200].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </label>
            <div className="pager">
              <button className="icon-btn" onClick={handlePrev} disabled={currentPage === 1}>←</button>
              <span>{currentPage} / {totalPages}</span>
              <button className="icon-btn" onClick={handleNext} disabled={currentPage === totalPages}>→</button>
            </div>
          </div>
        </div>

        <div className="table-scroll">
          <table className="inventory-table">
            <thead>
              <tr>
                <th scope="col" className="selection-col" aria-label="Seleziona">
                  <span className="selection-col__icon" aria-hidden="true"></span>
                </th>
                {visibleColumns.map((col) => (
                  <th
                    scope="col"
                    key={col.field}
                    className={col.field === "Title" ? "sticky-col" : undefined}
                    style={col.width ? { width: col.width } : undefined}
                  >
                    <span>{col.label}</span>
                    <small>{col.field}</small>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 && (
                <tr>
                  <td colSpan={visibleColumns.length + 1} className="muted" style={{ textAlign: "center" }}>
                    {loading ? "Carico dati..." : "Nessun elemento trovato"}
                  </td>
                </tr>
              )}
              {visibleRows.map((item) => {
                const cartItem: CartItem = {
                  key: `ORING-HNBR-${item.id}`,
                  source: "ORING-HNBR",
                  itemId: item.id,
                  title: ((item.fields as Record<string, unknown>).Title as string) || "-",
                  fields: item.fields as Record<string, unknown>,
                };
                const checked = Boolean(selectedItems[cartItem.key]);

                return (
                  <tr key={item.id}>
                    <td className="selection-col">
                      <input
                        type="checkbox"
                        aria-label="Seleziona riga"
                        checked={checked}
                        disabled={!checked && selectionLimitReached}
                        onChange={() => onToggle(cartItem, checked)}
                      />
                    </td>
                    {visibleColumns.map((col) => {
                      const isTitle = col.field === "Title";
                      const className = isTitle ? "sticky-col" : undefined;
                      const content = formatCellValue((item.fields as Record<string, unknown>)[col.field], col.type);

                      if (isTitle) {
                        return (
                          <th scope="row" key={col.field} className={className}>
                            {content}
                          </th>
                        );
                      }

                      return (
                        <td key={col.field} className={className}>
                          {content}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function OringNbrPanel({ selectedItems, onToggle, selectionLimitReached }: SelectionStateProps) {
  const getClient = useAuthenticatedGraphClient();
  const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
  const listId = import.meta.env.VITE_ORING_NBR_LIST_ID;

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);

  const service = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);

  const { data: rawRows, loading, error, refresh } = useCachedList<Record<string, unknown>>(
    service,
    listId,
    "oring-nbr"
  );

  const rows = useMemo(() => {
    const sorted = [...rawRows];
    sorted.sort((a, b) => {
      const valA = (a.fields as any).field_17;
      const valB = (b.fields as any).field_17;
      return getTimeValue(valB) - getTimeValue(valA);
    });
    return sorted;
  }, [rawRows]);

  const visibleColumns = oringNbrColumns;

  const normalizedRows = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return rows;
    return rows.filter((item) =>
      visibleColumns.some((col) => {
        const v = (item.fields as Record<string, unknown>)[col.field];
        if (v === null || v === undefined) return false;

        let valStr = String(v);
        if (col.type === "date") {
          valStr = formatSharePointDate(v);
        }

        return valStr.toLowerCase().includes(term);
      })
    );
  }, [rows, search, visibleColumns]);

  const totalPages = Math.max(1, Math.ceil(normalizedRows.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageSliceStart = (currentPage - 1) * pageSize;
  const visibleRows = normalizedRows.slice(pageSliceStart, pageSliceStart + pageSize);

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
    setPage(1);
  };

  const handleClearSearch = () => {
    setSearch("");
    setPage(1);
  };

  const handlePrev = () => setPage((p) => Math.max(1, p - 1));
  const handleNext = () => setPage((p) => Math.min(totalPages, p + 1));

  return (
    <div className="panel inventory-panel">
      <div className="panel-content">
        {error && <div className="alert error">{error}</div>}
        {!listId && (
          <div className="alert warning">Inserisci il GUID della lista ORING-NBR in .env.local (VITE_ORING_NBR_LIST_ID).</div>
        )}

        <div className="toolbar">
          <div className="search-box">
            <input
              type="search"
              placeholder="Cerca in tutte le colonne..."
              value={search}
              onChange={handleSearchChange}
            />
            {search && <button className="icon-btn" onClick={handleClearSearch} aria-label="Clear search">✕</button>}
          </div>
          
          <div className="pager-info">
             <button className="btn secondary" onClick={() => refresh()} disabled={loading} style={{ padding: '8px 12px', fontSize: '13px' }}>
              {loading ? "..." : "Aggiorna"}
            </button>
            <span className="pill">{normalizedRows.length} risultati</span>
            <label className="page-size">
              <span>Per pagina</span>
              <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}>
                {[100, 200].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </label>
            <div className="pager">
              <button className="icon-btn" onClick={handlePrev} disabled={currentPage === 1}>←</button>
              <span>{currentPage} / {totalPages}</span>
              <button className="icon-btn" onClick={handleNext} disabled={currentPage === totalPages}>→</button>
            </div>
          </div>
        </div>

        <div className="table-scroll">
          <table className="inventory-table">
            <thead>
              <tr>
                <th scope="col" className="selection-col" aria-label="Seleziona">
                  <span className="selection-col__icon" aria-hidden="true"></span>
                </th>
                {visibleColumns.map((col) => (
                  <th
                    scope="col"
                    key={col.field}
                    className={col.field === "Title" ? "sticky-col" : undefined}
                    style={col.width ? { width: col.width } : undefined}
                  >
                    <span>{col.label}</span>
                    <small>{col.field}</small>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.length === 0 && (
                <tr>
                  <td colSpan={visibleColumns.length + 1} className="muted" style={{ textAlign: "center" }}>
                    {loading ? "Carico dati..." : "Nessun elemento trovato"}
                  </td>
                </tr>
              )}
              {visibleRows.map((item) => {
                const cartItem: CartItem = {
                  key: `ORING-NBR-${item.id}`,
                  source: "ORING-NBR",
                  itemId: item.id,
                  title: ((item.fields as Record<string, unknown>).Title as string) || "-",
                  fields: item.fields as Record<string, unknown>,
                };
                const checked = Boolean(selectedItems[cartItem.key]);

                return (
                  <tr key={item.id}>
                    <td className="selection-col">
                      <input
                        type="checkbox"
                        aria-label="Seleziona riga"
                        checked={checked}
                        disabled={!checked && selectionLimitReached}
                        onChange={() => onToggle(cartItem, checked)}
                      />
                    </td>
                    {visibleColumns.map((col) => {
                      const isTitle = col.field === "Title";
                      const className = isTitle ? "sticky-col" : undefined;
                      const content = formatCellValue((item.fields as Record<string, unknown>)[col.field], col.type);

                      if (isTitle) {
                        return (
                          <th scope="row" key={col.field} className={className}>
                            {content}
                          </th>
                        );
                      }

                      return (
                        <td key={col.field} className={className}>
                          {content}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function TubiPanel({ selectedItems, onToggle, selectionLimitReached }: SelectionStateProps) {
  const getClient = useAuthenticatedGraphClient();
  const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
  const listId = import.meta.env.VITE_TUBI_LIST_ID;
  const tubiExcelPath = (import.meta.env.VITE_TUBI_EXCEL_PATH || "").trim();
  const tubiExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const tubiExcelFilename = (import.meta.env.VITE_SP_TUBI_FILENAME || import.meta.env.VITE_TUBI_EXCEL_FILE || "").trim();
  const tubiExcelTable = (import.meta.env.VITE_TUBI_EXCEL_TABLE || "tblTUBI").trim();
  const tubiExcelDriveIdEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_ID || "").trim();
  const tubiExcelDriveNameEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const tubiExcelDriveIdRef = useRef<string | null>(tubiExcelDriveIdEnv || null);

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [lotSelection, setLotSelection] = useState<{
    title: string;
    items: TubiItem[];
  } | null>(null);
  const [newLotColata, setNewLotColata] = useState("");
  const [newLotOrdine, setNewLotOrdine] = useState("");
  const [newLotDataOrdine, setNewLotDataOrdine] = useState("");
  const [newLotCodiceSam, setNewLotCodiceSam] = useState("");
  const [lotSaveStatus, setLotSaveStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [lotSaveMessage, setLotSaveMessage] = useState<string | null>(null);
  const GIACENZA_RESET_VALUE = 500;

  const service = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);

  const { data: rawRows, loading, error, refresh } = useCachedList<Record<string, unknown>>(
    service,
    listId,
    "tubi"
  );

  const rows = useMemo(() => {
    const sorted = [...rawRows];
    sorted.sort((a, b) => {
      const valA = (a.fields as any).field_3;
      const valB = (b.fields as any).field_3;
      return getTimeValue(valB) - getTimeValue(valA);
    });
    return sorted;
  }, [rawRows]);

  const groupedRows = useMemo(() => {
    const map = new Map<string, TubiItem[]>();
    rows.forEach((item) => {
      const title = toStr((item.fields as Record<string, unknown>).Title) || "(senza codice)";
      const existing = map.get(title) || [];
      map.set(title, [...existing, item]);
    });
    return Array.from(map.entries()).map(([title, items]) => ({
      title,
      items,
      representative: items[0],
    }));
  }, [rows]);

  useEffect(() => {
    if (!lotSelection) return;
    const updated = groupedRows.find((group) => group.title === lotSelection.title);
    if (!updated) return;

    const currentIds = lotSelection.items.map((itm) => itm.id).join(",");
    const nextIds = updated.items.map((itm) => itm.id).join(",");
    if (currentIds !== nextIds) {
      setLotSelection(updated);
    }
  }, [groupedRows, lotSelection]);

  const tubiProgressiveMap = useMemo(() => {
    const map = new Map<string, string>();

    const normalizeLetter = (val: string | null | undefined) => {
      if (!val) return null;
      const cleaned = String(val).trim().toLowerCase();
      if (/^[a-z]$/.test(cleaned)) return cleaned;
      return null;
    };

    const getCreatedKey = (item: TubiItem) => {
      const created = getTimeValue((item.fields as any).Created);
      if (created) return created;
      const idNum = Number(item.id);
      return Number.isFinite(idNum) ? idNum : 0;
    };

    groupedRows.forEach((group) => {
      const sorted = [...group.items].sort((a, b) => getCreatedKey(a) - getCreatedKey(b));
      const used = new Set<string>();

      sorted.forEach((item) => {
        const existing = normalizeLetter((item.fields as any).LottoProgressivo);
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
  }, [groupedRows]);

  const visibleColumns = tubiColumns;

  const normalizedRows = useMemo(() => {
    const term = search.trim().toLowerCase();
    if (!term) return groupedRows;
    return groupedRows.filter((group) =>
      group.items.some((item) =>
        visibleColumns.some((col) => {
          const v = (item.fields as Record<string, unknown>)[col.field];
          if (v === null || v === undefined) return false;

          let valStr = String(v);
          if (col.type === "date") {
            valStr = formatSharePointDate(v);
          }

          return valStr.toLowerCase().includes(term);
        })
      )
    );
  }, [groupedRows, search, visibleColumns]);

  const totalPages = Math.max(1, Math.ceil(normalizedRows.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageSliceStart = (currentPage - 1) * pageSize;
  const visibleRows = normalizedRows.slice(pageSliceStart, pageSliceStart + pageSize);

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
    setPage(1);
  };

  const handleClearSearch = () => {
    setSearch("");
    setPage(1);
  };

  const handlePrev = () => setPage((p) => Math.max(1, p - 1));
  const handleNext = () => setPage((p) => Math.min(totalPages, p + 1));

  const toCartItem = (item: TubiItem): CartItem => ({
    key: `TUBI-${item.id}`,
    source: "TUBI",
    itemId: item.id,
    title: toStr((item.fields as Record<string, unknown>).Title) || "-",
    bolla: (item.fields as Record<string, unknown>).field_15,
    colata: (item.fields as Record<string, unknown>).field_18,
    lottoProg: tubiProgressiveMap.get(item.id) || "a",
    fields: item.fields as Record<string, unknown>,
  });

  const handleSelectLot = (item: TubiItem) => {
    const cartItem = toCartItem(item);
    const alreadySelected = Boolean(selectedItems[cartItem.key]);
    onToggle(cartItem, alreadySelected);
    setLotSelection(null);
  };

  const existingColate = useMemo(() => {
    const set = new Set<string>();
    if (!lotSelection) return set;
    lotSelection.items.forEach((itm) => {
      const val = toStr((itm.fields as any).field_18).trim().toLowerCase();
      if (val) set.add(val);
    });
    return set;
  }, [lotSelection]);

  useEffect(() => {
    if (lotSelection) {
      const baseItem = lotSelection.items[0];
      const baseFields = (baseItem?.fields || {}) as any;
      const nextProg = getNextLottoProg(lotSelection.items, tubiProgressiveMap);
      setNewLotColata(buildColataPlaceholder(nextProg));
      setNewLotOrdine("");
      setNewLotDataOrdine(toInputDate(baseFields?.field_3));
      setNewLotCodiceSam("");
      setLotSaveStatus("idle");
      setLotSaveMessage(null);
    }
  }, [lotSelection, tubiProgressiveMap]);

  const buildLotClonePayload = (
    item: TubiItem,
    options: {
      colata: string;
      ordine?: string;
      dataOrdine?: string | null;
      codiceSam?: string;
      giacenzaNonTagliato: number | string;
      giacenzaContabile: number | string;
    }
  ) => {
    const { colata, ordine, dataOrdine, codiceSam, giacenzaNonTagliato, giacenzaContabile } = options;
    const sourceFields = (item.fields || {}) as Record<string, unknown>;
    const payload: Record<string, unknown> = {};

    const normalizeDateValue = (val: unknown): string | null => {
      const t = getTimeValue(val);
      if (!t) return null;
      return new Date(t).toISOString();
    };

    const normalizeTextValue = (val: unknown): string | null => {
      if (val === null || val === undefined) return null;
      const trimmed = String(val).trim();
      return trimmed ? trimmed : null;
    };

    const ordineValue = normalizeTextValue(ordine !== undefined ? ordine : (sourceFields as any).field_2);
    const dataOrdineValue = dataOrdine !== undefined ? dataOrdine : normalizeDateValue((sourceFields as any).field_3);
    const dataConsegnaValue = normalizeDateValue((sourceFields as any).field_16);
    const dataUltimoPrelievoValue = normalizeDateValue((sourceFields as any).field_21);
    const codiceSamValue = normalizeTextValue(codiceSam);
    const giacenzaContabValue = normalizeTextValue(giacenzaContabile);
    const giacenzaNonTagliatoValue = normalizeTextValue(giacenzaNonTagliato);

    Object.entries(sourceFields).forEach(([key, value]) => {
      if (
        key === "id" ||
        key === "ContentType" ||
        key === "ContentTypeId" ||
        key === "Created" ||
        key === "Modified" ||
        key === "AuthorLookupId" ||
        key === "EditorLookupId" ||
        key === "ComplianceAssetId" ||
        key === "GUID" ||
        key === "Attachments" ||
        key === "AppAuthor" ||
        key === "AppEditor" ||
        key === "AppAuthorLookupId" ||
        key === "AppEditorLookupId" ||
        key === "Edit" ||
        key === "ItemChildCount" ||
        key === "FolderChildCount" ||
        key === "_UIVersionString" ||
        key === "UIVersionString" ||
        key === "_ModerationStatus" ||
        key === "_ModerationComments" ||
        key === "_ComplianceFlags" ||
        key === "_ComplianceTag" ||
        key === "_ComplianceTagUserId" ||
        key === "_ComplianceTagWrittenTime" ||
        key === "_IsRecord" ||
        key === "_Level" ||
        key === "_Version" ||
        key.startsWith("LinkTitle") ||
        key.startsWith("_Compliance") ||
        key.startsWith("@") ||
        key.startsWith("odata")
      ) {
        return;
      }

      if (key === "field_18") {
        payload[key] = normalizeTextValue(colata);
        return;
      }

      if (key === "field_2") {
        payload[key] = ordineValue;
        return;
      }

      if (key === "field_3") {
        payload[key] = dataOrdineValue;
        return;
      }

      if (key === "field_16") {
        payload[key] = dataConsegnaValue;
        return;
      }

      if (key === "field_19") {
        payload[key] = giacenzaContabValue;
        return;
      }

      if (key === "field_20") {
        payload[key] = giacenzaNonTagliatoValue;
        return;
      }

      if (key === "field_21") {
        payload[key] = dataUltimoPrelievoValue;
        return;
      }

      if (key === "CodiceSAM") {
        payload[key] = codiceSamValue;
        return;
      }

      if (key === "LottoProgressivo") {
        payload[key] = null;
        return;
      }

      payload[key] = normalizeTextValue(value);
    });

    if (!("field_18" in payload)) payload.field_18 = normalizeTextValue(colata);
    if (!("field_2" in payload)) payload.field_2 = ordineValue;
    if (!("field_3" in payload)) payload.field_3 = dataOrdineValue;
    if (!("field_19" in payload)) payload.field_19 = giacenzaContabValue;
    if (!("field_20" in payload)) payload.field_20 = giacenzaNonTagliatoValue;
    if (!("CodiceSAM" in payload)) payload.CodiceSAM = null;

    return payload;
  };

  const handleCreateLot = async () => {
    if (!service || !listId || !lotSelection) {
      setLotSaveStatus("error");
      setLotSaveMessage("Servizio non disponibile per la creazione del lotto.");
      return;
    }

    const normalized = newLotColata.trim();
    if (!normalized) return;
    const normalizedLower = normalized.toLowerCase();
    if (existingColate.has(normalizedLower)) {
      setLotSaveStatus("error");
      setLotSaveMessage("Inserisci un N° colata diverso dai lotti esistenti.");
      return;
    }

    const baseItem = lotSelection.items[0];
    if (!baseItem) return;

    setLotSaveStatus("saving");
    setLotSaveMessage(null);
    let payload: Record<string, unknown> | null = null;
    try {
      const ordineOverride = newLotOrdine.trim();
      const dataOrdineInput = newLotDataOrdine.trim();
      payload = buildLotClonePayload(baseItem, {
        colata: normalized,
        ordine: ordineOverride ? ordineOverride : undefined,
        dataOrdine: dataOrdineInput ? toIsoOrNull(dataOrdineInput) : undefined,
        codiceSam: newLotCodiceSam.trim() ? newLotCodiceSam : undefined,
        giacenzaNonTagliato: GIACENZA_RESET_VALUE,
        giacenzaContabile: GIACENZA_RESET_VALUE,
      });
      const newLotProg = formatLottoProg(
        getNextLottoProg(lotSelection.items, tubiProgressiveMap)
      );
      await service.createItem<Record<string, unknown>>(listId, payload);
      let excelError: string | null = null;
      const resolvedPath = tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
      let resolvedDriveId = tubiExcelDriveIdRef.current;
      if (!resolvedDriveId && tubiExcelDriveNameEnv) {
        resolvedDriveId = await service.getDriveIdByName(tubiExcelDriveNameEnv);
        tubiExcelDriveIdRef.current = resolvedDriveId;
      }
      if (!resolvedDriveId && tubiExcelDriveNameEnv) {
        excelError = `Libreria "${tubiExcelDriveNameEnv}" non trovata`;
      } else if (resolvedPath && tubiExcelTable) {
        try {
          const driveItem = await service.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await service.listWorkbookTableColumnsByItemId(driveItem.id, tubiExcelTable, resolvedDriveId || undefined);
          const identIndex = getExcelColumnIndex(columns, "IdentLotto");
          if (identIndex === null) {
            throw new Error("Colonna IdentLotto non trovata nella tabella Excel");
          }
          const rowValues = buildTubiExcelRow(columns, {
            ...payload,
            IdentLotto: newLotProg,
          });
          const sessionId = await service.createWorkbookSessionByItemId(
            driveItem.id,
            { persistChanges: true },
            resolvedDriveId || undefined
          );
          try {
            await service.appendWorkbookTableRowByItemId(
              driveItem.id,
              tubiExcelTable,
              rowValues,
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
              console.warn("Errore chiusura sessione Excel TUBI", closeErr);
            }
          }
        } catch (err: any) {
          console.error("Errore aggiornamento Excel TUBI", {
            error: err,
            filePath: resolvedPath,
            driveId: resolvedDriveId,
            tableName: tubiExcelTable,
            payload,
          });
          excelError = err?.message || "Errore aggiornamento Excel";
        }
      } else {
        excelError = "Percorso Excel o tabella non configurati";
      }
      setLotSaveStatus("success");
      setLotSaveMessage(
        excelError
          ? `Nuovo lotto creato su SharePoint. Excel non aggiornato: ${excelError}`
          : "Nuovo lotto creato e registrato su Excel. Aggiorno la lista..."
      );
      const fallbackProg = getNextLottoProg(lotSelection.items, tubiProgressiveMap);
      setNewLotColata(buildColataPlaceholder(fallbackProg));
      setNewLotOrdine("");
      setNewLotCodiceSam("");
      await refresh();
    } catch (err: any) {
      console.error("Errore creazione lotto TUBI", {
        error: err,
        listId,
        payload,
      });
      setLotSaveStatus("error");
      setLotSaveMessage(err?.message || "Errore durante la creazione del lotto.");
    }
  };

  const normalizedNewLot = newLotColata.trim();
  const normalizedNewLotLower = normalizedNewLot.toLowerCase();
  const isDuplicateLot = normalizedNewLot.length > 0 && existingColate.has(normalizedNewLotLower);
  const canCreateLot = Boolean(normalizedNewLot) && !isDuplicateLot && lotSaveStatus !== "saving";

  return (
    <div className="panel inventory-panel">
      <div className="panel-content">
        {error && <div className="alert error">{error}</div>}
        {!listId && (
          <div className="alert warning">Inserisci il GUID della lista TUBI in .env.local (VITE_TUBI_LIST_ID).</div>
        )}

        <div className="toolbar">
          <div className="search-box">
            <input
              type="search"
              placeholder="Cerca in tutte le colonne..."
              value={search}
              onChange={handleSearchChange}
            />
            {search && <button className="icon-btn" onClick={handleClearSearch} aria-label="Clear search">✕</button>}
          </div>
          
          <div className="pager-info">
             <button className="btn secondary" onClick={() => refresh()} disabled={loading} style={{ padding: '8px 12px', fontSize: '13px' }}>
              {loading ? "..." : "Aggiorna"}
            </button>
            <span className="pill">{normalizedRows.length} risultati</span>
            <label className="page-size">
              <span>Per pagina</span>
              <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}>
                {[100, 200].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </label>
            <div className="pager">
              <button className="icon-btn" onClick={handlePrev} disabled={currentPage === 1}>←</button>
              <span>{currentPage} / {totalPages}</span>
              <button className="icon-btn" onClick={handleNext} disabled={currentPage === totalPages}>→</button>
            </div>
          </div>
        </div>

        <div className="table-scroll">
          <table className="inventory-table">
            <thead>
              <tr>
                <th scope="col" className="selection-col" aria-label="Seleziona">
                  <span className="selection-col__icon" aria-hidden="true"></span>
                </th>
                {visibleColumns.map((col) => (
                  <th
                    scope="col"
                    key={col.field}
                    className={col.field === "Title" ? "sticky-col" : undefined}
                    style={col.width ? { width: col.width } : undefined}
                  >
                    <span>{col.label}</span>
                    <small>{col.field}</small>
                  </th>
                ))}
                <th scope="col" style={{ width: 160 }}>Colate</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.length === 0 && (
                <tr>
                  <td colSpan={visibleColumns.length + 2} className="muted" style={{ textAlign: "center" }}>
                    {loading ? "Carico dati..." : "Nessun elemento trovato"}
                  </td>
                </tr>
              )}
              {visibleRows.map((group) => {
                const representative = group.representative;
                const selectedCount = group.items.filter((itm) => selectedItems[`TUBI-${itm.id}`]).length;
                const lotLabel = `${group.items.length} colata${group.items.length === 1 ? "" : "e"}`;

                const handleGroupCheck = () => {
                  if (group.items.length === 1) {
                    const cartItem = toCartItem(group.items[0]);
                    onToggle(cartItem, Boolean(selectedItems[cartItem.key]));
                    return;
                  }
                  setLotSelection(group);
                };

                return (
                  <tr key={group.title}>
                    <td className="selection-col">
                      <input
                        type="checkbox"
                        aria-label="Seleziona riga"
                        checked={selectedCount > 0}
                        disabled={selectionLimitReached && selectedCount === 0 && group.items.length === 1}
                        onChange={handleGroupCheck}
                      />
                    </td>
                    {visibleColumns.map((col) => {
                      const isTitle = col.field === "Title";
                      const className = isTitle ? "sticky-col" : undefined;
                      const content = formatCellValue((representative.fields as Record<string, unknown>)[col.field], col.type);

                      if (isTitle) {
                        return (
                          <th scope="row" key={col.field} className={className}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "space-between" }}>
                              <span>{content}</span>
                              <button
                                className="icon-btn"
                                aria-label="Gestisci lotti"
                                title="Gestisci lotti"
                                onClick={() => setLotSelection(group)}
                                type="button"
                              >
                                📦
                              </button>
                            </div>
                          </th>
                        );
                      }

                      return (
                        <td key={col.field} className={className}>
                          {content}
                        </td>
                      );
                    })}
                    <td>
                      <div className="colata-badges" style={{ alignItems: "center" }}>
                        {(() => {
                          const sortedLots = [...group.items].sort(
                            (a, b) => getTimeValue((b.fields as any).field_21) - getTimeValue((a.fields as any).field_21)
                          );
                          const latest = sortedLots[0];
                          const moreCount = Math.max(0, sortedLots.length - 1);
                          const latestProg = latest ? tubiProgressiveMap.get(latest.id) || "a" : "a";
                          const latestColata = latest ? toStr((latest.fields as any).field_18) || "-" : "-";
                          return (
                            <>
                              <span className="pill ghost" style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                                <span>{latestColata}</span>
                                <span className="lotto-chip">{formatLottoProg(latestProg)}</span>
                              </span>
                              {moreCount > 0 && (
                                <span className="pill ghost" style={{ fontSize: 11, padding: "2px 6px" }}>+{moreCount}</span>
                              )}
                            </>
                          );
                        })()}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {lotSelection && (
        <div className="modal-backdrop blur" role="dialog" aria-modal="true">
          <div className="modal">
            <div className="modal__header">
              <div>
                <p className="eyebrow" style={{ marginBottom: 4 }}>Seleziona lotto</p>
                <h3 style={{ margin: 0 }}>{lotSelection.title}</h3>
              </div>
              <button className="icon-btn" aria-label="Chiudi" onClick={() => setLotSelection(null)} type="button">✕</button>
            </div>
            <div className="modal__body">
              <div className="table-scroll modal-table">
                <table className="inventory-table">
                  <thead>
                    <tr>
                      <th scope="col">Ident. Lotto</th>
                      <th scope="col">N° COLATA</th>
                      <th scope="col">N° ORDINE</th>
                      <th scope="col">DATA ORDINE</th>
                      <th scope="col">Giacenza Non Tagliato</th>
                      <th scope="col" style={{ width: 140 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {lotSelection.items.map((item) => {
                      const cartItem = toCartItem(item);
                      const alreadySelected = Boolean(selectedItems[cartItem.key]);
                      const disableAdd = selectionLimitReached && !alreadySelected;
                      const lotProg = tubiProgressiveMap.get(item.id) || "a";
                      return (
                        <tr key={item.id}>
                          <td><span className="lotto-chip">{formatLottoProg(lotProg)}</span></td>
                          <td>{toStr((item.fields as any).field_18) || "-"}</td>
                          <td>{toStr((item.fields as any).field_2) || "-"}</td>
                          <td>{formatSharePointDate((item.fields as any).field_3)}</td>
                          <td>{toStr((item.fields as any).field_20) || "-"}</td>
                          <td>
                            <button
                              className={alreadySelected ? "btn secondary" : "btn primary"}
                              style={{ width: "100%", padding: "8px 10px", fontSize: 13 }}
                              onClick={() => handleSelectLot(item)}
                              disabled={disableAdd}
                              type="button"
                            >
                              {alreadySelected ? "Rimuovi" : "Aggiungi"}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="section-heading" style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <span className="pill ghost">Nuovo lotto</span>
                  <p className="muted" style={{ margin: "6px 0 0" }}>Duplico l&apos;articolo</p>
                </div>
              </div>
              <div className="field-grid" style={{ marginTop: 8, alignItems: "flex-end" }}>
                <label className="field" style={{ flex: 1 }}>
                  <span>Nuovo N° COLATA</span>
                  <input
                    type="text"
                    value={newLotColata}
                    onChange={(e) => {
                      setNewLotColata(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Inserisci un valore diverso dai lotti esistenti"
                  />
                </label>
                <label className="field">
                  <span>Nuovo N° ORDINE</span>
                  <input
                    type="text"
                    value={newLotOrdine}
                    onChange={(e) => {
                      setNewLotOrdine(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Facoltativo"
                  />
                </label>
                <label className="field">
                  <span>Codice SAM</span>
                  <input
                    type="text"
                    value={newLotCodiceSam}
                    onChange={(e) => {
                      setNewLotCodiceSam(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                    placeholder="Facoltativo"
                  />
                </label>
                <label className="field">
                  <span>Data ordine</span>
                  <input
                    type="date"
                    value={newLotDataOrdine}
                    onChange={(e) => {
                      setNewLotDataOrdine(e.target.value);
                      setLotSaveStatus("idle");
                      setLotSaveMessage(null);
                    }}
                  />
                </label>
                <button
                  className="btn primary"
                  style={{ minWidth: 160, height: 44 }}
                  onClick={handleCreateLot}
                  disabled={!canCreateLot}
                  type="button"
                >
                  {lotSaveStatus === "saving" ? "Creo..." : "Crea lotto"}
                </button>
              </div>
              {newLotColata && isDuplicateLot && (
                <p className="muted" style={{ marginTop: 6 }}>Valore già presente tra le colate esistenti.</p>
              )}
              {lotSaveMessage && (
                <div className={`alert ${lotSaveStatus === "success" ? "success" : lotSaveStatus === "error" ? "error" : "warning"}`} style={{ marginTop: 10 }}>
                  {lotSaveMessage}
                </div>
              )}
              {selectionLimitReached && <p className="muted" style={{ marginTop: 10 }}>Limite massimo di 10 articoli in carrello.</p>}
            </div>
            <div className="modal__footer">
              <button className="btn secondary" onClick={() => setLotSelection(null)} type="button">Chiudi</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function buildEditableState(item: CartItem): EditableItemState {
  const fields = item.fields || {};

  if (item.source === "FORGIATI") {
    return {
      type: "FORGIATI",
      title: item.title || "-",
      commessa: toStr((fields as any).field_24),
      dataPrelievo: toInputDate((fields as any).field_23),
      giacenzaQt: toStr((fields as any).field_22),
      note: toStr((fields as any).field_25),
      giacenzaBarra: toStr((fields as any).field_21),
      bolla: item.bolla ? String(item.bolla) : "",
      colata: item.colata ? String(item.colata) : "",
      raw: fields,
    };
  }

  if (item.source === "ORING-HNBR") {
    return {
      type: "ORING-HNBR",
      title: item.title || "-",
      commessa: toStr((fields as any).field_12),
      dataPrelievo: toInputDate((fields as any).field_17),
      giacenza: toStr((fields as any).field_15),
      prenotazione: toStr((fields as any).field_16),
      nrOrdine: toStr((fields as any).field_18),
      dataOrdine: toInputDate((fields as any).field_19),
      raw: fields,
    };
  }

  if (item.source === "ORING-NBR") {
    return {
      type: "ORING-NBR",
      title: item.title || "-",
      commessa: toStr((fields as any).field_11),
      giacenza: toStr((fields as any).field_16),
      prenotazione: toStr((fields as any).field_15),
      dataPrelievo: toInputDate((fields as any).field_17),
      qtaMinima: toStr((fields as any).field_14),
      raw: fields,
    };
  }

  return {
    type: "TUBI",
    title: item.title || "-",
    dataUltimoPrelievo: toInputDate((fields as any).field_21),
    commessa: toStr((fields as any).field_25),
    giacenzaBib: toStr((fields as any).field_20),
    giacenzaContab: toStr((fields as any).field_19),
    bolla: item.bolla ? String(item.bolla) : "",
    colata: item.colata ? String(item.colata) : "",
    raw: fields,
  };
}

function StockUpdatePage({
  items,
  values,
  onChange,
  onSave,
  saving,
  status,
  message,
  onBackToInventory,
}: {
  items: CartItem[];
  values: Record<string, EditableItemState>;
  onChange: (key: string, field: keyof EditableItemState, value: string) => void;
  onSave: () => void;
  saving: boolean;
  status: "idle" | "saving" | "success" | "error";
  message: string | null;
  onBackToInventory: () => void;
}) {
  if (items.length === 0) {
    return (
      <div className="panel">
        <div className="panel-content">
          <p className="muted">Non hai ancora selezionato articoli.</p>
          <button className="btn primary" onClick={onBackToInventory} type="button">
            ← Torna alla selezione
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="update-page">
      <div className="update-actions">
        <div>
          <p className="eyebrow" style={{ marginBottom: 6 }}>Carrello</p>
          <h2 style={{ margin: 0 }}>{items.length} articoli da aggiornare</h2>
          <p className="muted">Rivedi e modifica solo i campi indicati per ciascuna tipologia.</p>
        </div>
        <div className="action-row">
          <button className="btn secondary" onClick={onBackToInventory} type="button">
            ← Torna alla selezione
          </button>
          <button className="btn primary" onClick={onSave} disabled={saving} type="button">
            {saving ? "Salvo..." : "Aggiorna"}
          </button>
        </div>
      </div>

      {message && (
        <div className={`alert ${status === "success" ? "success" : status === "error" ? "error" : "warning"}`}>
          {message}
        </div>
      )}

      <div className="update-grid">
        {items.map((item) => {
          const cardValues = values[item.key] || buildEditableState(item);
          const isForgiati = item.source === "FORGIATI";
          const isOring = item.source === "ORING-HNBR";
          const isOringNbr = item.source === "ORING-NBR";
          const raw = (cardValues.raw || {}) as Record<string, unknown>;

          const infoItems = isForgiati
            ? [
                { label: "N° Ordine", value: raw["field_1"] },
                { label: "Posizione", value: raw["field_4"] },
                { label: "Q.tà orig", value: raw["field_5"] },
                { label: "DN", value: raw["field_6"] },
                { label: "Classe", value: raw["field_7"] },
                { label: "Materiale", value: raw["field_9"] },
                { label: "No. Disegno - Particolare", value: raw["field_8"] },
                { label: "Codice SAM", value: raw["CodiceSAM"] },
              ]
            : isOring
            ? [
                { label: "Posizione", value: raw["field_0"] },
                { label: "Codice cliente", value: raw["field_2"] },
                { label: "Colore Mescola", value: raw["field_3"] },
                { label: "Ø Int.", value: raw["field_4"] },
                { label: "Ø Est.", value: raw["field_5"] },
                { label: "Ø Corda", value: raw["field_6"] },
                { label: "Materiale", value: raw["field_7"] },
                { label: "Durezza", value: raw["field_8"] },
                { label: "Giacenza", value: raw["field_15"] },
                { label: "Prenotazione", value: raw["field_16"] },
              ]
            : isOringNbr
            ? [
                { label: "Codice cliente", value: raw["field_1"] },
                { label: "Colore Mescola", value: raw["field_2"] },
                { label: "Ø Int.", value: raw["field_3"] },
                { label: "Ø Est.", value: raw["field_4"] },
                { label: "Ø Corda", value: raw["field_5"] },
                { label: "Materiale", value: raw["field_6"] },
                { label: "Durezza", value: raw["field_7"] },
                { label: "Fluidi", value: raw["field_8"] },
                { label: "Giacenza", value: raw["field_16"] },
                { label: "Prenotazione", value: raw["field_15"] },
              ]
            : [
                { label: "N° Ordine", value: raw["field_2"] },
                { label: "Codice SAM", value: raw["CodiceSAM"] },
                { label: "Fornitore", value: raw["field_4"] },
                { label: "DATA ORDINE", value: formatSharePointDate(raw["field_3"]) },
                { label: "Q.tà", value: raw["field_6"] },
                { label: "Lung. tubo mm", value: raw["field_7"] },
                { label: "DN", value: raw["field_8"] },
                { label: "DN mm", value: raw["field_9"] },
                { label: "SP", value: raw["field_10"] },
                { label: "Grado", value: raw["field_11"] },
                { label: "N Cert", value: raw["field_17"] },
                { label: "DATA CONSEGNA", value: formatSharePointDate(raw["field_16"]) },
              ];

          return (
            <div
              key={item.key}
              className={`stock-card ${
                isForgiati
                  ? "stock-card--forgiati"
                  : isOring
                  ? "stock-card--oring"
                  : isOringNbr
                  ? "stock-card--oring-nbr"
                  : "stock-card--tubi"
              }`}
            >
              <div className="stock-card__head">
                <div className="pill pill-tight">{item.source}</div>
                <div>
                  <div className="stock-card__title">{cardValues.title || item.title || "-"}</div>
                  <div className="stock-card__meta">
                    <span>{cardValues.bolla ? `Bolla ${cardValues.bolla}` : "Bolla -"}</span>
                    <span className="bullet">•</span>
                    <span>{cardValues.colata ? `Colata ${cardValues.colata}` : "Colata -"}</span>
                    <span className="bullet">•</span>
                    <span>
                      Ident. lotto <span className="lotto-chip">{formatLottoProg(item.lottoProg)}</span>
                    </span>
                  </div>
                </div>
              </div>

              <div className="info-grid">
                {infoItems.map((info) => (
                  <div key={`${item.key}-${info.label}`} className="info-chip">
                    <span className="info-label">{info.label}</span>
                    <span className="info-value">{toStr(info.value) || "-"}</span>
                  </div>
                ))}
              </div>

              <div
                className={`stock-card__body ${
                  isForgiati ? "is-forgiati" : isOring ? "is-oring" : isOringNbr ? "is-oring-nbr" : "is-tubi"
                }`}
              >
                <div className="section-heading">
                  <span className="pill ghost">Campi modificabili</span>
                </div>
                {isForgiati ? (
                  <>
                    <div className="field-grid">
                      <label className="field">
                        <span>Commessa</span>
                        <input
                          type="text"
                          value={cardValues.commessa || ""}
                          onChange={(e) => onChange(item.key, "commessa", e.target.value)}
                        />
                      </label>
                      <label className="field">
                        <span>Data prelievo</span>
                        <input
                          type="date"
                          value={cardValues.dataPrelievo || ""}
                          onChange={(e) => onChange(item.key, "dataPrelievo", e.target.value)}
                        />
                      </label>
                      <label className="field">
                        <span>Giacenza Q.tà</span>
                        <input
                          type="number"
                          inputMode="decimal"
                          value={cardValues.giacenzaQt || ""}
                          onChange={(e) => onChange(item.key, "giacenzaQt", e.target.value)}
                        />
                      </label>
                      <label className="field">
                        <span>Giacenza mm da barra</span>
                        <input
                          type="number"
                          inputMode="decimal"
                          value={cardValues.giacenzaBarra || ""}
                          onChange={(e) => onChange(item.key, "giacenzaBarra", e.target.value)}
                        />
                      </label>
                    </div>
                    <label className="field full">
                      <span>Note</span>
                      <textarea
                        rows={2}
                        value={cardValues.note || ""}
                        onChange={(e) => onChange(item.key, "note", e.target.value)}
                      />
                    </label>
                  </>
                ) : isOring ? (
                  <div className="field-grid">
                    <label className="field">
                      <span>Commessa</span>
                      <input
                        type="text"
                        value={cardValues.commessa || ""}
                        onChange={(e) => onChange(item.key, "commessa", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Giacenza</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.giacenza || ""}
                        onChange={(e) => onChange(item.key, "giacenza", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Prenotazione</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.prenotazione || ""}
                        onChange={(e) => onChange(item.key, "prenotazione", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Data prelievo</span>
                      <input
                        type="date"
                        value={cardValues.dataPrelievo || ""}
                        onChange={(e) => onChange(item.key, "dataPrelievo", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Nr Ordine</span>
                      <input
                        type="text"
                        value={cardValues.nrOrdine || ""}
                        onChange={(e) => onChange(item.key, "nrOrdine", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Data ordine</span>
                      <input
                        type="date"
                        value={cardValues.dataOrdine || ""}
                        onChange={(e) => onChange(item.key, "dataOrdine", e.target.value)}
                      />
                    </label>
                  </div>
                ) : isOringNbr ? (
                  <div className="field-grid">
                    <label className="field">
                      <span>Commessa</span>
                      <input
                        type="text"
                        value={cardValues.commessa || ""}
                        onChange={(e) => onChange(item.key, "commessa", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Q.tà minima</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.qtaMinima || ""}
                        onChange={(e) => onChange(item.key, "qtaMinima", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Giacenza</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.giacenza || ""}
                        onChange={(e) => onChange(item.key, "giacenza", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Prenotazione</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.prenotazione || ""}
                        onChange={(e) => onChange(item.key, "prenotazione", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Data prelievo</span>
                      <input
                        type="date"
                        value={cardValues.dataPrelievo || ""}
                        onChange={(e) => onChange(item.key, "dataPrelievo", e.target.value)}
                      />
                    </label>
                  </div>
                ) : (
                  <div className="field-grid">
                    <label className="field">
                      <span>Data ultimo prelievo</span>
                      <input
                        type="date"
                        value={cardValues.dataUltimoPrelievo || ""}
                        onChange={(e) => onChange(item.key, "dataUltimoPrelievo", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Commessa</span>
                      <input
                        type="text"
                        value={cardValues.commessa || ""}
                        onChange={(e) => onChange(item.key, "commessa", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Giacenza Tubo Intero mm</span>
                      <input
                        type="number"
                        inputMode="decimal"
                        value={cardValues.giacenzaBib || ""}
                        onChange={(e) => onChange(item.key, "giacenzaBib", e.target.value)}
                      />
                    </label>
                    <label className="field">
                      <span>Giacenza Contab mm (somma accetta +)</span>
                      <input
                        type="text"
                        inputMode="decimal"
                        placeholder="es. 100+20+10"
                        value={cardValues.giacenzaContab || ""}
                        onChange={(e) => onChange(item.key, "giacenzaContab", e.target.value)}
                      />
                    </label>
                  </div>
                )}

              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function NavigationTabs({
  activeTab,
  onTabChange,
}: {
  activeTab: "forgiati" | "oring-hnbr" | "oring-nbr" | "tubi";
  onTabChange: (id: "forgiati" | "oring-hnbr" | "oring-nbr" | "tubi") => void;
}) {
  const tabs = [
    { id: "forgiati", label: "1_FORGIATI", enabled: true },
    { id: "oring-hnbr", label: "2_ORING-HNBR", enabled: true },
    { id: "oring-nbr", label: "2_ORING-NBR", enabled: true },
    { id: "tubi", label: "3_TUBI", enabled: true },
  ] as const;

  return (
    <nav className="nav-tabs">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          className={`nav-tab ${activeTab === tab.id ? "active" : tab.enabled ? "" : "disabled"}`}
          disabled={!tab.enabled}
          onClick={() => tab.enabled && onTabChange(tab.id)}
        >
          {tab.label}
        </button>
      ))}
    </nav>
  );
}

function AuthenticatedShell() {
  const { accounts } = useMsal();
  const getClient = useAuthenticatedGraphClient();
  const siteId = import.meta.env.VITE_SHAREPOINT_SITE_ID;
  const forgiatiListId = import.meta.env.VITE_FORGIATI_LIST_ID;
  const oringHnbrListId = import.meta.env.VITE_ORING_HNBR_LIST_ID;
  const oringNbrListId = import.meta.env.VITE_ORING_NBR_LIST_ID;
  const tubiListId = import.meta.env.VITE_TUBI_LIST_ID;
  const tubiExcelPath = (import.meta.env.VITE_TUBI_EXCEL_PATH || "").trim();
  const tubiExcelFolder = (import.meta.env.VITE_SP_FOLDER_PATH || import.meta.env.VITE_EXCEL_FOLDER_PATH || "").trim();
  const tubiExcelFilename = (import.meta.env.VITE_SP_TUBI_FILENAME || import.meta.env.VITE_TUBI_EXCEL_FILE || "").trim();
  const tubiExcelTable = (import.meta.env.VITE_TUBI_EXCEL_TABLE || "tblTUBI").trim();
  const tubiExcelDriveIdEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_ID || "").trim();
  const tubiExcelDriveNameEnv = (import.meta.env.VITE_TUBI_EXCEL_DRIVE_NAME || import.meta.env.VITE_SP_LIBRARY_NAME || "").trim();
  const tubiExcelDriveIdRef = useRef<string | null>(tubiExcelDriveIdEnv || null);
  const sharepointService = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);
  const flowService = useMemo(() => new PowerAutomateService(), []);
  const [view, setView] = useState<'dashboard' | 'inventory' | 'update-stock' | 'website' | 'docs' | 'admin'>('dashboard');
  const [activeTab, setActiveTab] = useState<"forgiati" | "oring-hnbr" | "oring-nbr" | "tubi">("forgiati");
  const [cartItems, setCartItems] = useState<Record<string, CartItem>>({});
  const [selectionMessage, setSelectionMessage] = useState<string | null>(null);
  const [editValues, setEditValues] = useState<Record<string, EditableItemState>>({});
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [isCartOpen, setIsCartOpen] = useState(true);
  const isAdmin = useIsAdmin();

  const shellClass = useMemo(() => {
    if (view === "update-stock") return "totem-shell update-mode";
    if (view === "admin") return "totem-shell admin-mode";
    if (view === "docs") return "totem-shell docs-mode";
    return "totem-shell";
  }, [view]);

  useEffect(() => {
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }, [accounts]);

  useEffect(() => {
    if (view === "admin" && !isAdmin) {
      setView("dashboard");
    }
  }, [view, isAdmin]);

  const handleBack = () => {
    setActiveTab("forgiati");
    setView('dashboard');
  };

  const selectionLimitReached = Object.keys(cartItems).length >= 10;

  const adminCta = isAdmin && view !== 'admin' ? (
    <button className="admin-cta" onClick={() => setView('admin')} type="button">
      Pannello Admin
    </button>
  ) : null;

  const handleToggleSelection = useCallback((item: CartItem, isSelected: boolean) => {
    setCartItems((prev) => {
      if (isSelected) {
        const { [item.key]: _, ...rest } = prev;
        setSelectionMessage(null);
        return rest;
      }

      if (Object.keys(prev).length >= 10) {
        setSelectionMessage("Puoi selezionare massimo 10 articoli.");
        return prev;
      }

      setSelectionMessage(null);
      return { ...prev, [item.key]: item };
    });
  }, []);

  const handleRemoveFromCart = useCallback((key: string) => {
    setCartItems((prev) => {
      const { [key]: _, ...rest } = prev;
      return rest;
    });
    setSelectionMessage(null);
  }, []);

  const handleClearCart = useCallback(() => {
    setCartItems({});
    setSelectionMessage(null);
  }, []);

  const cartList = useMemo(() => Object.values(cartItems), [cartItems]);

  const getCartOrderNumber = (item: CartItem): string => {
    const raw = (item.fields || {}) as Record<string, unknown>;
    if (item.source === "FORGIATI") return toStr((raw as any).field_1);
    if (item.source === "TUBI") return toStr((raw as any).field_2);
    if (item.source === "ORING-HNBR") return toStr((raw as any).field_18);
    return "";
  };

  useEffect(() => {
    if (cartList.length > 0) {
      setIsCartOpen(true);
    } else {
      setIsCartOpen(false);
    }
  }, [cartList]);

  useEffect(() => {
    setEditValues((prev) => {
      const next: Record<string, EditableItemState> = {};
      cartList.forEach((item) => {
        next[item.key] = prev[item.key] ? { ...prev[item.key] } : buildEditableState(item);
      });
      return next;
    });
  }, [cartList]);

  const handleEditChange = useCallback(
    (key: string, field: keyof EditableItemState, value: string) => {
      setEditValues((prev) => {
        const base = prev[key] || (cartItems[key] ? buildEditableState(cartItems[key]) : undefined);
        if (!base) return prev;
        const next: EditableItemState = { ...base, [field]: value };

        if (base.type === "TUBI" && field === "giacenzaContab") {
          const sum = sumAdditiveExpression(value);
          if (sum !== null) {
            next.giacenzaBib = String(sum);
          }
        }

        return {
          ...prev,
          [key]: next,
        };
      });
    },
    [cartItems]
  );

  useEffect(() => {
    setSaveStatus("idle");
    setSaveMessage(null);
  }, [view]);

  const handleSave = useCallback(async () => {
    if (!sharepointService) {
      setSaveStatus("error");
      setSaveMessage("Configurazione SharePoint mancante.");
      return;
    }

    if (cartList.length === 0) {
      setSaveStatus("error");
      setSaveMessage("Nessun articolo nel carrello da aggiornare.");
      return;
    }

    setSaveStatus("saving");
    setSaveMessage(null);

    try {
      const updates = cartList.map((item) => {
        const values = editValues[item.key] || buildEditableState(item);
        const isForgiati = item.source === "FORGIATI";
        const isOring = item.source === "ORING-HNBR";
        const isOringNbr = item.source === "ORING-NBR";
        const listId = isForgiati ? forgiatiListId : isOring ? oringHnbrListId : isOringNbr ? oringNbrListId : tubiListId;
        if (!listId) {
          throw new Error(`List ID non configurato per ${item.source}.`);
        }

        const contabSum = sumAdditiveExpression(values.giacenzaContab);

        const payload = isForgiati
          ? {
              field_24: values.commessa ?? null,
              field_23: toIsoOrNull(values.dataPrelievo),
              field_22: toNumberOrNull(values.giacenzaQt),
              field_25: values.note ?? null,
              field_21: toNumberOrNull(values.giacenzaBarra),
            }
          : isOring
          ? {
              field_12: values.commessa ?? null,
              field_15: toNumberOrNull(values.giacenza),
              field_16: toNumberOrNull(values.prenotazione),
              field_17: toIsoOrNull(values.dataPrelievo),
              field_18: values.nrOrdine ?? null,
              field_19: toIsoOrNull(values.dataOrdine),
            }
          : isOringNbr
          ? {
              field_11: values.commessa ?? null,
              field_14: toNumberOrNull(values.qtaMinima),
              field_16: toNumberOrNull(values.giacenza),
              field_15: toNumberOrNull(values.prenotazione),
              field_17: toIsoOrNull(values.dataPrelievo),
            }
          : {
              field_21: toIsoOrNull(values.dataUltimoPrelievo),
              field_25: values.commessa ?? null,
              field_20: contabSum ?? toNumberOrNull(values.giacenzaBib),
              field_19: values.giacenzaContab ?? null,
            };

        return { item, listId, values, payload };
      });

      await Promise.all(
        updates.map(({ item, listId, payload }) =>
          sharepointService.updateItem<Record<string, unknown>>(listId, item.itemId, payload)
        )
      );

      let excelUpdateError: string | null = null;
      const tubiUpdates = updates.filter((update) => update.item.source === "TUBI");
      if (tubiUpdates.length > 0) {
        try {
          const resolvedPath = tubiExcelPath || (tubiExcelFolder && tubiExcelFilename ? `${tubiExcelFolder}/${tubiExcelFilename}` : "");
          let resolvedDriveId = tubiExcelDriveIdRef.current;
          if (!resolvedDriveId && tubiExcelDriveNameEnv) {
            resolvedDriveId = await sharepointService.getDriveIdByName(tubiExcelDriveNameEnv);
            tubiExcelDriveIdRef.current = resolvedDriveId;
          }
          if (!resolvedDriveId && tubiExcelDriveNameEnv) {
            throw new Error(`Libreria "${tubiExcelDriveNameEnv}" non trovata`);
          }
          if (!resolvedPath || !tubiExcelTable) {
            throw new Error("Percorso Excel o tabella non configurati");
          }

          const driveItem = await sharepointService.getDriveItemByPath(resolvedPath, resolvedDriveId || undefined);
          const columns = await sharepointService.listWorkbookTableColumnsByItemId(driveItem.id, tubiExcelTable, resolvedDriveId || undefined);
          const codiceIndex = getExcelColumnIndex(columns, "Title");
          const identIndex = getExcelColumnIndex(columns, "IdentLotto");
          if (codiceIndex === null || identIndex === null) {
            throw new Error("Colonne CODICE/IdentLotto non trovate nella tabella Excel");
          }
          const rows = await sharepointService.listWorkbookTableRowsByItemId(driveItem.id, tubiExcelTable, resolvedDriveId || undefined);
          const sessionId = await sharepointService.createWorkbookSessionByItemId(
            driveItem.id,
            { persistChanges: true },
            resolvedDriveId || undefined
          );
          const missing: string[] = [];
          const targetCodici = new Set(tubiUpdates.map((update) => normalizeExcelKey(update.item.title)));
          try {
            for (const row of rows) {
              const rowValues = row.values?.[0] || [];
              const codiceVal = normalizeExcelKey(String(rowValues[codiceIndex] ?? ""));
              if (!targetCodici.has(codiceVal)) continue;
              const identVal = normalizeExcelKey(String(rowValues[identIndex] ?? ""));
              if (identVal) continue;
              const nextValues = [...rowValues];
              nextValues[identIndex] = "A";
              await sharepointService.updateWorkbookTableRowByIndex(
                driveItem.id,
                tubiExcelTable,
                row.index,
                nextValues as Array<string | number | boolean | null>,
                { sessionId },
                resolvedDriveId || undefined
              );
            }
            for (const update of tubiUpdates) {
              const identLotto = formatLottoProg(update.item.lottoProg);
              const rowIndex = findExcelRowIndex(rows, columns, {
                codice: update.item.title,
                identLotto,
              });
              if (rowIndex === null) {
                missing.push(`${update.item.title} (${identLotto})`);
                continue;
              }
              const excelFields = {
                ...(update.item.fields || {}),
                ...update.payload,
                IdentLotto: identLotto,
              } as Record<string, unknown>;
              const rowValues = buildTubiExcelRow(columns, excelFields);
              await sharepointService.updateWorkbookTableRowByIndex(
                driveItem.id,
                tubiExcelTable,
                rowIndex,
                rowValues,
                { sessionId },
                resolvedDriveId || undefined
              );
            }
          } finally {
            try {
              await sharepointService.closeWorkbookSessionByItemId(
                driveItem.id,
                sessionId,
                resolvedDriveId || undefined
              );
            } catch (closeErr) {
              console.warn("Errore chiusura sessione Excel TUBI", closeErr);
            }
          }

          if (missing.length > 0) {
            throw new Error(`Righe non trovate: ${missing.join(", ")}`);
          }
        } catch (excelErr: any) {
          console.error("Errore aggiornamento Excel TUBI", excelErr);
          excelUpdateError = excelErr?.message || "Errore aggiornamento Excel";
        }
      }

      const account = accounts[0];
      const sources = Array.from(new Set(cartList.map((i) => i.source)));
      const flowPayload = {
        action: "AggiornaGiacenza",
        user: account?.username || account?.name || "unknown",
        displayName: account?.name || null,
        totalItems: cartList.length,
        sources,
        hasForgiati: sources.includes("FORGIATI"),
        hasTubi: sources.includes("TUBI"),
        hasOringHnbr: sources.includes("ORING-HNBR"),
        hasOringNbr: sources.includes("ORING-NBR"),
        items: cartList.map((item) => {
          const values = editValues[item.key] || buildEditableState(item);
          const isTubi = item.source === "TUBI";
          const isForgiati = item.source === "FORGIATI";
          return {
            key: item.key,
            itemId: item.itemId,
            title: item.title,
            source: item.source,
            commessa: values.commessa ?? null,
            giacenza: values.giacenza ?? values.giacenzaQt ?? values.giacenzaBib ?? null,
            prenotazione: values.prenotazione ?? null,
            note: values.note ?? null,
            dataUltimoPrelievo: isTubi ? values.dataUltimoPrelievo ?? null : null,
            giacenzaTuboIntero: isTubi ? values.giacenzaBib ?? null : null,
            giacenzaContabMm: isTubi ? values.giacenzaContab ?? null : null,
            dataPrelievo: isForgiati ? values.dataPrelievo ?? null : null,
            giacenzaQt: isForgiati ? values.giacenzaQt ?? null : null,
            giacenzaBarra: isForgiati ? values.giacenzaBarra ?? null : null,
          };
        }),
        timestamp: new Date().toISOString(),
      };

      const baseMessage = excelUpdateError
        ? `Giacenza aggiornata; Excel non aggiornato: ${excelUpdateError}`
        : "Giacenza aggiornata con successo.";
      setSaveStatus("success");
      setSaveMessage(baseMessage);

      try {
        await flowService.invoke(flowPayload);
      } catch (flowErr: any) {
        console.error("Errore Power Automate", flowErr);
        setSaveMessage(`${baseMessage}; notifica Teams non inviata: ${flowErr?.message || "errore flow"}`);
      }
    } catch (err: any) {
      console.error("Errore aggiornamento", err);
      setSaveStatus("error");
      setSaveMessage(err?.message || "Errore durante il salvataggio.");
    }
  }, [accounts, cartList, editValues, flowService, forgiatiListId, oringHnbrListId, oringNbrListId, sharepointService, tubiListId]);

  const renderActivePanel = () => {
    const selectionProps: SelectionStateProps = {
      selectedItems: cartItems,
      onToggle: handleToggleSelection,
      selectionLimitReached,
    };

    switch (activeTab) {
      case "oring-hnbr":
        return <OringHnbrPanel {...selectionProps} />;
      case "oring-nbr":
        return <OringNbrPanel {...selectionProps} />;
      case "tubi":
        return <TubiPanel {...selectionProps} />;
      default:
        return <ForgiatiPanel {...selectionProps} />;
    }
  };

  if (view === 'dashboard') {
    return (
      <>
        <div className={shellClass}>
          <Hero 
            title="Benvenuto in Totem Alfa" 
            subtitle="Seleziona un'operazione per iniziare"
          />
          <main className="dashboard-main">
            <Dashboard onNavigate={setView} />
          </main>
        </div>
        {adminCta}
      </>
    );
  }

  if (view === 'website') {
    return (
      <>
        <div className={shellClass}>
          <Hero 
            title="Sito Web Aziendale" 
            onBack={() => setView('dashboard')}
          />
          <main className="dashboard-main" style={{ padding: '0 2rem 2rem' }}>
            <WebsiteViewer url="https://www.alfa-eng.net/" />
          </main>
        </div>
        {adminCta}
      </>
    );
  }

  if (view === 'docs') {
    return (
      <>
        <div className={shellClass}>
          <Hero 
            title="Consulta Documentazione" 
            subtitle="Naviga cartelle e apri PDF senza uscire dal totem"
            onBack={() => setView('dashboard')}
          />
          <main className="dashboard-main" style={{ 
            padding: '0 2rem 2rem', 
            maxWidth: '1400px', 
            margin: '0 auto', 
            width: '100%',
            overflowY: 'auto',
            alignItems: 'flex-start',
            display: 'block', // Override flex centering for this view 
            height: '100%' 
          }}>
            <DocumentBrowser
              siteId={siteId}
              driveId={import.meta.env.VITE_SHAREPOINT_DRIVE_ID}
              initialPath={import.meta.env.VITE_DOCS_ROOT_PATH || ""}
            />
          </main>
        </div>
        {adminCta}
      </>
    );
  }

  if (view === 'admin') {
    return (
      <>
        <div className={shellClass}>
          <Hero
            title="Pannello Admin"
            subtitle="Accesso riservato agli amministratori"
            onBack={() => setView('dashboard')}
          />
          <main
            className="dashboard-main"
            style={{
              padding: '0 2rem 2rem',
              width: '100%',
              maxWidth: '1500px',
              margin: '0 auto',
              alignItems: 'flex-start',
              display: 'block',
            }}
          >
            <AdminPanel
              siteId={siteId}
              forgiatiListId={forgiatiListId}
              tubiListId={tubiListId}
              oringHnbrListId={oringHnbrListId}
              oringNbrListId={oringNbrListId}
            />
          </main>
        </div>
        {adminCta}
      </>
    );
  }

  if (view === 'update-stock') {
    return (
      <>
        <div className={shellClass}>
          <Hero 
            title="Aggiorna Giacenza" 
            onBack={() => setView('inventory')}
          />
          <StockUpdatePage
            items={cartList}
            values={editValues}
            onChange={handleEditChange}
            onSave={handleSave}
            saving={saveStatus === "saving"}
            status={saveStatus}
            message={saveMessage}
            onBackToInventory={() => setView('inventory')}
          />
        </div>
        {adminCta}
      </>
    );
  }

  return (
    <>
      <div className={shellClass}>
        <Hero 
          title="Giacenza Magazzino" 
          onBack={handleBack}
        />
        <NavigationTabs activeTab={activeTab} onTabChange={setActiveTab} />
        {selectionMessage && (
          <div className="alert warning" style={{ marginTop: 6 }}>
            {selectionMessage}
          </div>
        )}
        {cartList.length > 0 && (
          <div className="cart-panel">
            <div className="cart-header">
              <div>
                <p className="eyebrow" style={{ marginBottom: 6 }}>Carrello</p>
                <h3>{cartList.length} / 10 articoli selezionati</h3>
              </div>
              <div className="cart-actions">
                <button
                  className="btn secondary"
                  style={{ padding: "10px 12px" }}
                  onClick={() => setIsCartOpen((v) => !v)}
                  type="button"
                >
                  {isCartOpen ? "Comprimi ▲" : "Espandi ▼"}
                </button>
                <button className="btn danger-ghost" onClick={handleClearCart} type="button">
                  🗑 Svuota articoli
                </button>
                <button className="btn primary" onClick={() => setView('update-stock')} type="button">
                  Aggiorna Giacenza
                </button>
              </div>
            </div>
            {isCartOpen && (
              <div className="cart-table-wrapper">
                <table className="cart-table">
                  <thead>
                    <tr>
                      <th scope="col">TIPO</th>
                      <th scope="col">ARTICOLO</th>
                      <th scope="col">N° Ordine</th>
                      <th scope="col">Colata</th>
                      <th scope="col">Ident. Lotto</th>
                      <th scope="col" style={{ width: 64 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {cartList.map((item) => (
                      <tr
                        key={item.key}
                        className={`cart-row ${
                          item.source === "FORGIATI"
                            ? "cart-row-forgiati"
                            : item.source === "TUBI"
                            ? "cart-row-tubi"
                            : item.source === "ORING-HNBR"
                            ? "cart-row-oring"
                            : "cart-row-oring-nbr"
                        }`}
                      >
                        <td>{item.source}</td>
                        <td>{item.title || "-"}</td>
                        <td>{getCartOrderNumber(item) || "-"}</td>
                        <td>{item.colata ? String(item.colata) : "-"}</td>
                        <td><span className="lotto-chip">{formatLottoProg(item.lottoProg)}</span></td>
                        <td>
                          <button
                            className="icon-btn danger"
                            aria-label="Rimuovi articolo"
                            onClick={() => handleRemoveFromCart(item.key)}
                            type="button"
                          >
                            ✕
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
        <main className="panels-grid">
          {renderActivePanel()}
        </main>
      </div>
      {adminCta}
    </>
  );
}

function LoginScreen() {
  const { instance } = useMsal();

  const handleLogin = () => instance.loginRedirect(loginRequest);

  return (
    <div className="login-screen">
      <div className="login-card">
        <p className="eyebrow">ALFA ENG</p>
        <h1>Entra nel Totem Magazzino</h1>
        <p className="muted">Accedi con l'account Microsoft aziendale per consultare e aggiornare le liste.</p>
        <button className="btn primary" onClick={handleLogin}>Accedi con Microsoft</button>
      </div>
    </div>
  );
}

function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedTemplate>
        <AuthenticatedShell />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <LoginScreen />
      </UnauthenticatedTemplate>
    </MsalProvider>
  );
}

export default App;
