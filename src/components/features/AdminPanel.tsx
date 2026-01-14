import React, { useCallback, useEffect, useMemo, useState } from "react";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";
import { SharePointService } from "../../services/sharePointService";
import { useCachedList } from "../../hooks/useCachedList";
import { formatSharePointDate } from "../../utils/dateUtils";
import { SharePointListItem } from "../../types/sharepoint";

type ListKind = "FORGIATI" | "TUBI";

type FieldType = "text" | "number" | "date" | "textarea";

type FieldConfig = {
  key: string;
  label: string;
  type?: FieldType;
  required?: boolean;
  placeholder?: string;
};

type FormState = Record<string, string>;

type AdminPanelProps = {
  siteId?: string;
  forgiatiListId?: string;
  tubiListId?: string;
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
  { key: "field_26", label: "Data/ora modifica" },
  { key: "CodiceSAM", label: "Codice SAM" },
  { key: "LottoProgressivo", label: "Lotto Progressivo" },
];

const TUBI_FIELDS: FieldConfig[] = [
  { key: "Title", label: "Codice / Title", required: true, placeholder: "Es. TUB-001" },
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
  { key: "field_26", label: "Esubero" },
  { key: "field_27", label: "RITARDO" },
  { key: "LottoProgressivo", label: "Lotto Progressivo" },
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

const buildFormFromItem = (
  item: SharePointListItem<Record<string, unknown>> | null,
  fields: FieldConfig[]
): FormState => {
  const form: FormState = {};
  fields.forEach((field) => {
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

const getFieldSet = (kind: ListKind) => (kind === "FORGIATI" ? FORGIATI_FIELDS : TUBI_FIELDS);

export function AdminPanel({ siteId, forgiatiListId, tubiListId }: AdminPanelProps) {
  const getClient = useAuthenticatedGraphClient();

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

  const activeItems = activeList === "FORGIATI" ? forgiatiItems : tubiItems;
  const activeLoading = activeList === "FORGIATI" ? forgiatiLoading : tubiLoading;
  const activeError = activeList === "FORGIATI" ? forgiatiError : tubiError;
  const activeListId = activeList === "FORGIATI" ? forgiatiListId : tubiListId;
  const activeRefresh = activeList === "FORGIATI" ? refreshForgiati : refreshTubi;
  const fieldSet = getFieldSet(activeList);

  const sortedItems = useMemo(() => {
    const items = [...activeItems];
    const dateKey = activeList === "FORGIATI" ? "field_2" : "field_3";
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
      const title = toStr(fields.Title).toLowerCase();
      const bolla = toStr(fields[activeList === "FORGIATI" ? "field_10" : "field_15"]).toLowerCase();
      const colata = toStr(fields[activeList === "FORGIATI" ? "field_13" : "field_18"]).toLowerCase();
      return title.includes(term) || bolla.includes(term) || colata.includes(term);
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
      setUpdateStatus("success");
      setUpdateMessage("Elemento aggiornato con successo");
      activeRefresh();
    } catch (err: any) {
      setUpdateStatus("error");
      setUpdateMessage(err?.message || "Errore durante il salvataggio");
    }
  }, [service, activeListId, selectedId, editForm, activeRefresh, activeList]);

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
    return {
      title: toStr(fields.Title) || "-",
      bolla: toStr(fields[activeList === "FORGIATI" ? "field_10" : "field_15"]) || "-",
      colata: toStr(fields[activeList === "FORGIATI" ? "field_13" : "field_18"]) || "-",
      giacenza: toStr(fields[activeList === "FORGIATI" ? "field_22" : "field_19"]) || "-",
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
            Modifica o inserisci articoli di FORGIATI e TUBI senza uscire dal totem.
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
        {(["FORGIATI", "TUBI"] as ListKind[]).map((kind) => (
          <button
            key={kind}
            className={`admin-switch__btn ${activeList === kind ? "active" : ""}`}
            onClick={() => setActiveList(kind)}
            type="button"
          >
            {kind === "FORGIATI" ? "1_FORGIATI" : "3_TUBI"}
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
                  placeholder="Cerca per codice, bolla o colata"
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
                      <span className="pill ghost">Bolla {summary.bolla}</span>
                      <span className="pill ghost">Colata {summary.colata}</span>
                      <span className="pill ghost">Giacenza {summary.giacenza}</span>
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
                <h3 style={{ margin: 0 }}>Nuovo {activeList === "FORGIATI" ? "forgiato" : "tubo"}</h3>
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
