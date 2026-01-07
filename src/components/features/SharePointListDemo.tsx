import { useEffect, useMemo, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { protectedResources } from "../../auth/authConfig";
import { useAuthenticatedGraphClient } from "../../hooks/useAuthenticatedGraphClient";
import { SharePointService } from "../../services/sharePointService";
import { SharePointListItem, VisitorFields } from "../../types/sharepoint";

export function SharePointListDemo() {
  const listId = protectedResources.sharePointListId;
  const { accounts } = useMsal();
  const getClient = useAuthenticatedGraphClient();
  const service = useMemo(() => {
    if (!protectedResources.sharePointSiteId) return null;
    return new SharePointService(getClient, protectedResources.sharePointSiteId);
  }, [getClient, protectedResources.sharePointSiteId]);

  const [items, setItems] = useState<SharePointListItem<VisitorFields>[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadItems = async () => {
    if (!listId || !service) return;
    setLoading(true);
    setError(null);
    try {
      const response = await service.listItems<VisitorFields>(listId);
      setItems(response);
    } catch (err: any) {
      setError(err?.message || "Errore nel recupero dati");
    } finally {
      setLoading(false);
    }
  };

  const seedDemoItem = async () => {
    if (!listId || !service) return;
      setLoading(true);
      setError(null);
      try {
        await service.createItem(listId, {
          Title: `VIS-${Math.floor(Math.random() * 99999)}`,
          Nome: "Mario",
          Cognome: "Rossi",
          Email: "mario.rossi@example.com",
          Azienda: "Contoso",
          Categoria: "VISITATORE",
          Stato: "Attivo",
        });
        await loadItems();
      } catch (err: any) {
        setError(err?.message || "Errore nella creazione item");
      } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (accounts.length && listId && service) {
      loadItems();
    }
  }, [accounts, listId, service]);

  return (
    <div className="panel">
      <h3>SharePoint + Graph</h3>
      <p>Collega la lista <strong>Visitatori</strong> e verifica CRUD via Microsoft Graph.</p>
      {!protectedResources.sharePointSiteId && (
        <div className="card-highlight">Configura VITE_SHAREPOINT_SITE_ID per connetterti al sito.</div>
      )}
      {protectedResources.sharePointSiteId && !listId && (
        <div className="card-highlight">Configura VITE_SHAREPOINT_LIST_ID per abilitare il test.</div>
      )}
      <div className="button-row" style={{ marginBottom: 12 }}>
        <button className="btn" onClick={loadItems} disabled={loading || !listId || !service} type="button">Ricarica</button>
        <button className="btn secondary" onClick={seedDemoItem} disabled={loading || !listId || !service} type="button">Crea item demo</button>
        <span className="tag">{loading ? "Loading..." : `${items.length} items`}</span>
      </div>
      {error && <div className="card-highlight" style={{ borderColor: "#ef4444", color: "#fca5a5" }}>{error}</div>}
      <div className="list">
        {items.map((item) => (
          <div className="list-item" key={item.id}>
            <strong>{item.fields.Title || "(ID mancante)"} • {item.fields.Nome || "-"} {item.fields.Cognome || ""}</strong>
            <span>{item.fields.Email || "Email non presente"}</span>
            <span>{item.fields.Azienda || "Azienda non presente"} • {item.fields.Categoria || "Categoria?"}</span>
            <span>{item.fields.Stato || "Stato?"}</span>
          </div>
        ))}
        {!items.length && !loading && listId && (
          <div className="card-highlight">Nessun dato ancora. Usa "Crea item demo" o inserisci dati dalla UI SharePoint.</div>
        )}
      </div>
    </div>
  );
}
