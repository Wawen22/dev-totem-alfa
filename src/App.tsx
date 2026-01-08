import React, { useCallback, useEffect, useMemo, useState } from "react";
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

  const forgiatiProgressiveMap = useMemo(() => {
    const map = new Map<string, string>();

    const normalizeLetter = (val: string | null | undefined) => {
      if (!val) return null;
      const cleaned = String(val).trim().toLowerCase();
      if (/^[a-z]$/.test(cleaned)) return cleaned;
      return null;
    };

    groupedRows.forEach((group) => {
      const sorted = [...group.items].sort((a, b) => getTimeValue((b.fields as any).field_23) - getTimeValue((a.fields as any).field_23));
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
                    <td>
                      <div className="colata-badges">
                        {(() => {
                          const sortedLots = [...group.items].sort(
                            (a, b) => getTimeValue((b.fields as any).field_23) - getTimeValue((a.fields as any).field_23)
                          );
                          const latest = sortedLots[0];
                          const moreCount = Math.max(0, sortedLots.length - 1);
                          const latestProg = latest ? forgiatiProgressiveMap.get(latest.id) || "a" : "a";
                          const latestLabel = latest ? `${toStr((latest.fields as any).field_13) || "-"} (${latestProg})` : "-";
                          return (
                            <>
                              <span className="pill ghost">{latestLabel}</span>
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
        <div className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal">
            <div className="modal__header">
              <div>
                <p className="eyebrow" style={{ marginBottom: 4 }}>Seleziona colata</p>
                <h3 style={{ margin: 0 }}>{lotSelection.title}</h3>
              </div>
              <button className="icon-btn" aria-label="Chiudi" onClick={() => setLotSelection(null)} type="button">✕</button>
            </div>
            <div className="modal__body">
              <div className="table-scroll modal-table">
                <table className="inventory-table">
                  <thead>
                    <tr>
                      <th scope="col">N° COLATA</th>
                      <th scope="col">Ident. Lotto</th>
                      <th scope="col">N° BOLLA</th>
                      <th scope="col">Data prelievo</th>
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
                          <td>{toStr((item.fields as any).field_13) || "-"}</td>
                          <td>{lotProg}</td>
                          <td>{toStr((item.fields as any).field_10) || "-"}</td>
                          <td>{formatSharePointDate((item.fields as any).field_23)}</td>
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

  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [lotSelection, setLotSelection] = useState<{
    title: string;
    items: TubiItem[];
  } | null>(null);

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

  const tubiProgressiveMap = useMemo(() => {
    const map = new Map<string, string>();

    const normalizeLetter = (val: string | null | undefined) => {
      if (!val) return null;
      const cleaned = String(val).trim().toLowerCase();
      if (/^[a-z]$/.test(cleaned)) return cleaned;
      return null;
    };

    groupedRows.forEach((group) => {
      const sorted = [...group.items].sort((a, b) => getTimeValue((b.fields as any).field_21) - getTimeValue((a.fields as any).field_21));
      const used = new Set<string>();

      sorted.forEach((item, idx) => {
        const existing = normalizeLetter((item.fields as any).LottoProgressivo);
        let letter = existing;
        if (!letter || used.has(letter)) {
          // Assign next available letter starting from 'a' when empty or duplicate
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

  const handleOpenLot = (group: { title: string; items: TubiItem[] }) => {
    if (group.items.length === 1) {
      const cartItem = toCartItem(group.items[0]);
      onToggle(cartItem, Boolean(selectedItems[cartItem.key]));
      return;
    }
    setLotSelection(group);
  };

  const handleSelectLot = (item: TubiItem) => {
    const cartItem = toCartItem(item);
    const alreadySelected = Boolean(selectedItems[cartItem.key]);
    onToggle(cartItem, alreadySelected);
    setLotSelection(null);
  };

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
                    <td>
                      <div className="colata-badges">
                        {(() => {
                          const sortedLots = [...group.items].sort(
                            (a, b) => getTimeValue((b.fields as any).field_21) - getTimeValue((a.fields as any).field_21)
                          );
                          const latest = sortedLots[0];
                          const moreCount = Math.max(0, sortedLots.length - 1);
                          const latestProg = latest ? tubiProgressiveMap.get(latest.id) || "a" : "a";
                          const latestLabel = latest ? `${toStr((latest.fields as any).field_18) || "-"} (${latestProg})` : "-";
                          return (
                            <>
                              <span className="pill ghost">{latestLabel}</span>
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
        <div className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal">
            <div className="modal__header">
              <div>
                <p className="eyebrow" style={{ marginBottom: 4 }}>Seleziona colata</p>
                <h3 style={{ margin: 0 }}>{lotSelection.title}</h3>
              </div>
              <button className="icon-btn" aria-label="Chiudi" onClick={() => setLotSelection(null)} type="button">✕</button>
            </div>
            <div className="modal__body">
              <div className="table-scroll modal-table">
                <table className="inventory-table">
                  <thead>
                    <tr>
                      <th scope="col">N° COLATA</th>
                      <th scope="col">Ident. Lotto</th>
                      <th scope="col">N° BOLLA</th>
                      <th scope="col">DATA ULT. PREL.</th>
                      <th scope="col">Giacenza Contab.</th>
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
                          <td>{toStr((item.fields as any).field_18) || "-"}</td>
                          <td>{lotProg}</td>
                          <td>{toStr((item.fields as any).field_15) || "-"}</td>
                          <td>{formatSharePointDate((item.fields as any).field_21)}</td>
                          <td>{toStr((item.fields as any).field_19) || "-"}</td>
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
                    <span>{item.lottoProg ? `Ident. lotto ${item.lottoProg}` : "Ident. lotto a"}</span>
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
  const sharepointService = useMemo(() => {
    if (!siteId) return null;
    return new SharePointService(getClient, siteId);
  }, [getClient, siteId]);
  const [view, setView] = useState<'dashboard' | 'inventory' | 'update-stock' | 'website' | 'docs' | 'admin'>('dashboard');
  const [activeTab, setActiveTab] = useState<"forgiati" | "oring-hnbr" | "oring-nbr" | "tubi">("forgiati");
  const [cartItems, setCartItems] = useState<Record<string, CartItem>>({});
  const [selectionMessage, setSelectionMessage] = useState<string | null>(null);
  const [editValues, setEditValues] = useState<Record<string, EditableItemState>>({});
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "success" | "error">("idle");
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [isCartOpen, setIsCartOpen] = useState(true);

  useEffect(() => {
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }, [accounts]);

  const activeAccount = accounts[0];
  const roles = useMemo(() => {
    const claims = (activeAccount?.idTokenClaims || {}) as Record<string, unknown>;
    const claimRoles = claims["roles"];
    if (Array.isArray(claimRoles)) return claimRoles.map((r) => String(r));
    if (typeof claimRoles === "string") return [claimRoles];
    return [];
  }, [activeAccount]);

  const isAdmin = roles.includes("Totem.Admin");

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
      await Promise.all(
        cartList.map((item) => {
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

          return sharepointService.updateItem<Record<string, unknown>>(listId, item.itemId, payload);
        })
      );

      setSaveStatus("success");
      setSaveMessage("Giacenza aggiornata con successo.");
    } catch (err: any) {
      console.error("Errore aggiornamento", err);
      setSaveStatus("error");
      setSaveMessage(err?.message || "Errore durante il salvataggio.");
    }
  }, [cartList, editValues, forgiatiListId, oringHnbrListId, oringNbrListId, sharepointService, tubiListId]);

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
        <div className="totem-shell">
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
        <div className="totem-shell">
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
        <div className="totem-shell">
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
        <div className="totem-shell">
          <Hero
            title="Pannello Admin"
            subtitle="Accesso riservato agli amministratori"
            onBack={() => setView('dashboard')}
          />
          <main className="dashboard-main">
            <div className="panel">
              <div className="panel-content">
                <p className="eyebrow" style={{ marginBottom: 8 }}>Stato</p>
                <p className="muted">Area admin in allestimento. Questo pulsante serve a verificare il ruolo Totem.Admin.</p>
                <button className="btn secondary" onClick={() => setView('dashboard')} type="button">
                  Torna alla dashboard
                </button>
              </div>
            </div>
          </main>
        </div>
        {adminCta}
      </>
    );
  }

  if (view === 'update-stock') {
    return (
      <>
        <div className="totem-shell update-mode">
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
      <div className="totem-shell">
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
                      <th scope="col">N° Bolla</th>
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
                        <td>{item.bolla ? String(item.bolla) : "-"}</td>
                        <td>{item.colata ? String(item.colata) : "-"}</td>
                        <td>{item.lottoProg ? String(item.lottoProg) : "a"}</td>
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
