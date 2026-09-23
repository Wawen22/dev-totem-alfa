import { useState, useEffect, useCallback } from "react";
import { SharePointService } from "../services/sharePointService";
import { SharePointListItem } from "../types/sharepoint";

// Cache globale che persiste finché l'app non viene ricaricata (F5)
const globalCache: Record<string, { data: any[]; timestamp: number }> = {};
const cacheAliases: Record<string, string> = {};
const CACHE_DURATION = 5 * 60 * 1000; // 5 minuti di validità cache default

export function clearCacheKeys(keys: string[]): void {
  keys.forEach((key) => {
    delete globalCache[key];
    const aliasedKey = cacheAliases[key];
    if (aliasedKey) delete globalCache[aliasedKey];
  });
}

export function useCachedList<T = any>(
  service: SharePointService | null,
  listId: string | undefined,
  listNameKey: string // Chiave univoca per la cache (es. 'forgiati', 'tubi')
) {
  const [data, setData] = useState<SharePointListItem<T>[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchData = useCallback(
    async (forceRefresh = false) => {
      if (!service) {
        setError("Servizio SharePoint non inizializzato");
        return;
      }
      if (!listId) {
        setError(`Config mancante per lista ${listNameKey}`);
        return;
      }

      // 1. Controllo Cache (se non forziamo l'aggiornamento)
      // The same SharePoint list can be rendered by the operational and admin
      // panels. Cache by its immutable list ID so the two views can never show
      // different snapshots of the same data.
      const cacheKey = listId || listNameKey;
      cacheAliases[listNameKey] = cacheKey;
      const cached = globalCache[cacheKey];
      const now = Date.now();

      if (!forceRefresh && cached && (now - cached.timestamp < CACHE_DURATION)) {
        console.log(`[CACHE] Utilizzo dati in cache per ${listNameKey}`);
        setData(cached.data);
        setLoading(false);
        return;
      }

      // 2. Fetch da Rete
      try {
        setLoading(true);
        setError(null);
        console.log(`[NETWORK] Scaricamento dati per ${listNameKey}...`);
        
        const result = await service.listItems<T>(listId);
        
        // 3. Aggiornamento Cache
        globalCache[cacheKey] = {
          data: result,
          timestamp: Date.now(),
        };

        setData(result);
      } catch (err) {
        const message = err instanceof Error ? err.message : "Errore imprevisto";
        setError(message);
      } finally {
        setLoading(false);
      }
    },
    [service, listId, listNameKey]
  );

  // Caricamento iniziale
  useEffect(() => {
    fetchData(false);
  }, [fetchData]);

  return { data, loading, error, refresh: () => fetchData(true) };
}
