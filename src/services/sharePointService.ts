import { Client } from "@microsoft/microsoft-graph-client";
import { SharePointListItem } from "../types/sharepoint";

export class SharePointService {
  private siteId: string;

  constructor(private readonly getClient: () => Promise<Client>, siteId?: string) {
    this.siteId = siteId || import.meta.env.VITE_SHAREPOINT_SITE_ID;
    if (!this.siteId) {
      throw new Error("VITE_SHAREPOINT_SITE_ID mancante");
    }
  }

  async listItems<TFields = Record<string, unknown>>(listId: string): Promise<SharePointListItem<TFields>[]> {
    if (!listId) throw new Error("listId mancante");
    const client = await this.getClient();
    let allItems: SharePointListItem<TFields>[] = [];
    const initialPath = `/sites/${this.siteId}/lists/${listId}/items`;
    let requestPath: string | undefined = initialPath;
    const visitedPages = new Set<string>();
    let pageCount = 0;

    // Graph returns an opaque, complete @odata.nextLink. Never recreate it or
    // reapply query options: doing so can skip older pages of a large list.
    while (requestPath) {
      if (visitedPages.has(requestPath)) {
        throw new Error(`Paginazione Graph ciclica rilevata per la lista ${listId}`);
      }
      visitedPages.add(requestPath);

      const request = client.api(requestPath);
      const response = requestPath === initialPath
        ? await request.expand("fields").top(999).orderby("createdDateTime desc").get()
        : await request.get();
      const items = (response.value || []).map((item: any) => ({
        id: item.id,
        fields: item.fields,
      }));
      allItems = allItems.concat(items);
      pageCount += 1;
      requestPath = typeof response["@odata.nextLink"] === "string"
        ? response["@odata.nextLink"]
        : undefined;
    }

    // Deduplicate by ID to ensure clean data
    const uniqueItems = new Map<string, SharePointListItem<TFields>>();
    allItems.forEach(item => uniqueItems.set(item.id, item));
    
    const result = Array.from(uniqueItems.values());
    console.info(`[GRAPH] Lista ${listId}: ${result.length} elementi su ${pageCount} pagina/e`);
    return result;
  }

  async createItem<TFields extends Record<string, unknown>>(listId: string, fields: TFields): Promise<SharePointListItem<TFields>> {
    if (!listId) throw new Error("listId mancante");
    const client = await this.getClient();
    try {
      const response = await client
        .api(`/sites/${this.siteId}/lists/${listId}/items`)
        .post({ fields });

      return { id: response.id, fields: response.fields };
    } catch (err: any) {
      // Il GraphError espone il corpo della risposta in `err.body` (di solito una
      // stringa JSON). Lo interpretiamo per estrarre codice/messaggio/dettagli e
      // logghiamo l'errore completo insieme ai campi inviati per diagnosticare i
      // rifiuti di tipo colonna (es. "General exception while processing").
      let parsedBody: any = err?.body;
      if (typeof parsedBody === "string") {
        try {
          parsedBody = JSON.parse(parsedBody);
        } catch {
          /* body non JSON: lasciamo la stringa originale */
        }
      }
      const graphError = parsedBody?.error;
      const innerMsg =
        graphError?.message ||
        parsedBody?.["odata.error"]?.message?.value ||
        err?.message;
      // Riepilogo tipo di ogni campo inviato (utile per capire quale colonna
      // SharePoint rifiuta: es. una data/numero/lookup con tipo inatteso).
      const fieldTypes = Object.fromEntries(
        Object.entries(fields as Record<string, unknown>).map(([k, v]) => [
          k,
          v === null ? "null" : Array.isArray(v) ? "array" : typeof v,
        ])
      );
      console.error("Errore createItem SharePoint", {
        listId,
        statusCode: err?.statusCode,
        code: graphError?.code || err?.code,
        message: innerMsg,
        innerError: graphError?.innerError,
        body: parsedBody,
        fieldTypes,
        fieldsJson: JSON.stringify(fields),
        fields,
      });
      const msg = innerMsg || "Errore creazione elemento";
      throw new Error(msg);
    }
  }

  async updateItem<TFields extends Record<string, unknown>>(listId: string, itemId: string, fields: Partial<TFields>): Promise<void> {
    if (!listId) throw new Error("listId mancante");
    if (!itemId) throw new Error("itemId mancante");
    const client = await this.getClient();

    await client
      .api(`/sites/${this.siteId}/lists/${listId}/items/${itemId}/fields`)
      .patch(fields);
  }

  async deleteItem(listId: string, itemId: string): Promise<void> {
    if (!listId) throw new Error("listId mancante");
    if (!itemId) throw new Error("itemId mancante");
    const client = await this.getClient();
    await client
      .api(`/sites/${this.siteId}/lists/${listId}/items/${itemId}`)
      .delete();
  }

  async listColumns(listId: string): Promise<Array<{ name: string; displayName?: string; columnGroup?: string }>> {
    if (!listId) throw new Error("listId mancante");
    const client = await this.getClient();
    const response = await client
      .api(`/sites/${this.siteId}/lists/${listId}/columns`)
      .select("name,displayName,columnGroup")
      .get();

    return response.value || [];
  }

  async listDrives(): Promise<Array<{ id: string; name?: string }>> {
    const client = await this.getClient();
    const response = await client.api(`/sites/${this.siteId}/drives`).get();
    return response.value || [];
  }

  async getDriveIdByName(name: string): Promise<string | null> {
    if (!name) return null;
    const drives = await this.listDrives();
    const normalized = name.trim().toLowerCase();
    const match = drives.find((drive) => (drive.name || "").trim().toLowerCase() === normalized);
    return match?.id || null;
  }

  private encodePathSegments(path: string) {
    return path
      .split("/")
      .map((segment) => encodeURIComponent(segment))
      .join("/");
  }

  private buildDriveBase(driveId?: string) {
    return driveId
      ? `/sites/${this.siteId}/drives/${driveId}`
      : `/sites/${this.siteId}/drive`;
  }

  async getDriveItemByPath(filePath: string, driveId?: string): Promise<{ id: string; name?: string; webUrl?: string }> {
    if (!filePath) throw new Error("filePath mancante");
    const client = await this.getClient();
    const cleanPath = filePath.replace(/^\/+/, "");
    const encodedPath = this.encodePathSegments(cleanPath);
    const base = this.buildDriveBase(driveId);
    const response = await client
      .api(`${base}/root:/${encodedPath}`)
      .select("id,name,webUrl")
      .get();

    return {
      id: response.id,
      name: response.name,
      webUrl: response.webUrl,
    };
  }

  async listWorkbookTableColumnsByPath(filePath: string, tableName: string, driveId?: string): Promise<string[]> {
    if (!filePath) throw new Error("filePath mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const cleanPath = filePath.replace(/^\/+/, "");
    const encodedPath = this.encodePathSegments(cleanPath);
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    const response = await client
      .api(`${base}/root:/${encodedPath}:/workbook/tables/${encodedTable}/columns`)
      .get();

    return (response.value || []).map((col: { name?: string }) => col.name).filter(Boolean);
  }

  async listWorkbookTableColumnsByItemId(itemId: string, tableName: string, driveId?: string): Promise<string[]> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    const response = await client
      .api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/columns`)
      .get();

    return (response.value || []).map((col: { name?: string }) => col.name).filter(Boolean);
  }

  async listWorkbookTableRowsByItemId(
    itemId: string,
    tableName: string,
    driveId?: string
  ): Promise<Array<{ index: number; values: Array<Array<unknown>> }>> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    let rows: Array<{ index: number; values: Array<Array<unknown>> }> = [];
    let nextLink: string | undefined = `${base}/items/${itemId}/workbook/tables/${encodedTable}/rows`;

    while (nextLink) {
      const response = await client.api(nextLink).get();
      const value = response.value || [];
      rows = rows.concat(
        value.map((row: { index: number; values: Array<Array<unknown>> }) => ({
          index: row.index,
          values: row.values,
        }))
      );
      nextLink = response["@odata.nextLink"];
    }

    return rows;
  }

  async getWorkbookTableDataBodyRangeByItemId(
    itemId: string,
    tableName: string,
    driveId?: string
  ): Promise<{ address: string; rowCount?: number; columnCount?: number }> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    const response = await client
      .api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/dataBodyRange`)
      .get();

    return {
      address: response.address,
      rowCount: response.rowCount,
      columnCount: response.columnCount,
    };
  }

  async createWorkbookSessionByItemId(
    itemId: string,
    options: { persistChanges?: boolean } = {},
    driveId?: string
  ): Promise<string> {
    if (!itemId) throw new Error("itemId mancante");
    const client = await this.getClient();
    const base = this.buildDriveBase(driveId);
    const response = await client
      .api(`${base}/items/${itemId}/workbook/createSession`)
      .post({ persistChanges: options.persistChanges ?? true });

    return response.id;
  }

  async closeWorkbookSessionByItemId(itemId: string, sessionId: string, driveId?: string): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!sessionId) throw new Error("sessionId mancante");
    const client = await this.getClient();
    const base = this.buildDriveBase(driveId);
    await client
      .api(`${base}/items/${itemId}/workbook/closeSession`)
      .header("workbook-session-id", sessionId)
      .post({});
  }

  async appendWorkbookTableRowByPath(
    filePath: string,
    tableName: string,
    values: Array<string | number | boolean | null>,
    driveId?: string
  ): Promise<void> {
    if (!filePath) throw new Error("filePath mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const cleanPath = filePath.replace(/^\/+/, "");
    const encodedPath = this.encodePathSegments(cleanPath);
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    await client
      .api(`${base}/root:/${encodedPath}:/workbook/tables/${encodedTable}/rows/add`)
      .post({ values: [values] });
  }

  async appendWorkbookTableRowByItemId(
    itemId: string,
    tableName: string,
    values: Array<string | number | boolean | null>,
    options: { sessionId?: string; index?: number } = {},
    driveId?: string
  ): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    let req = client.api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/rows/add`);
    if (options.sessionId) {
      req = req.header("workbook-session-id", options.sessionId);
    }
    const body: { values: Array<Array<string | number | boolean | null>>; index?: number } = { values: [values] };
    if (options.index !== undefined && Number.isFinite(options.index)) {
      body.index = options.index;
    }
    await req.post(body);
  }

  async updateWorkbookRangeByAddress(
    itemId: string,
    sheetName: string,
    address: string,
    values: Array<Array<string | number | boolean | null>>,
    options: { sessionId?: string } = {},
    driveId?: string
  ): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!sheetName) throw new Error("sheetName mancante");
    if (!address) throw new Error("address mancante");
    const client = await this.getClient();
    const base = this.buildDriveBase(driveId);
    const sheetSegment = encodeURIComponent(sheetName);
    let req = client.api(`${base}/items/${itemId}/workbook/worksheets/${sheetSegment}/range(address='${address}')`);
    if (options.sessionId) {
      req = req.header("workbook-session-id", options.sessionId);
    }
    await req.patch({ values });
  }

  async getWorkbookRangeValuesByAddress(
    itemId: string,
    sheetName: string,
    address: string,
    options: { sessionId?: string } = {},
    driveId?: string
  ): Promise<Array<Array<unknown>>> {
    if (!itemId) throw new Error("itemId mancante");
    if (!sheetName) throw new Error("sheetName mancante");
    if (!address) throw new Error("address mancante");
    const client = await this.getClient();
    const base = this.buildDriveBase(driveId);
    const sheetSegment = encodeURIComponent(sheetName);
    let req = client.api(`${base}/items/${itemId}/workbook/worksheets/${sheetSegment}/range(address='${address}')`);
    if (options.sessionId) {
      req = req.header("workbook-session-id", options.sessionId);
    }
    const response = await req.get();
    return response.values || [];
  }

  async updateWorkbookTableRowByIndex(
    itemId: string,
    tableName: string,
    rowIndex: number,
    values: Array<string | number | boolean | null>,
    options: { sessionId?: string } = {},
    driveId?: string
  ): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    if (!Number.isFinite(rowIndex)) throw new Error("rowIndex non valido");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    let firstError: any = null;
    try {
      let req = client.api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/rows/${rowIndex}`);
      if (options.sessionId) {
        req = req.header("workbook-session-id", options.sessionId);
      }
      await req.patch({ values: [values] });
      return;
    } catch (err) {
      firstError = err;
    }

    try {
      let req = client.api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/rows/itemAt(index=${rowIndex})`);
      if (options.sessionId) {
        req = req.header("workbook-session-id", options.sessionId);
      }
      await req.patch({ values: [values] });
    } catch (fallbackErr: any) {
      const firstMsg = firstError?.body?.error?.message || firstError?.message || "errore aggiornamento riga";
      const fallbackMsg = fallbackErr?.body?.error?.message || fallbackErr?.message || "errore fallback aggiornamento riga";
      throw new Error(`${firstMsg}; fallback itemAt: ${fallbackMsg}`);
    }
  }

  async deleteWorkbookTableRowByIndex(
    itemId: string,
    tableName: string,
    rowIndex: number,
    options: { sessionId?: string } = {},
    driveId?: string
  ): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!tableName) throw new Error("tableName mancante");
    if (!Number.isFinite(rowIndex)) throw new Error("rowIndex non valido");
    const client = await this.getClient();
    const encodedTable = encodeURIComponent(tableName);
    const base = this.buildDriveBase(driveId);
    let req = client.api(`${base}/items/${itemId}/workbook/tables/${encodedTable}/rows/${rowIndex}`);
    if (options.sessionId) {
      req = req.header("workbook-session-id", options.sessionId);
    }
    await req.delete();
  }

  async deleteWorkbookRangeByAddress(
    itemId: string,
    sheetName: string,
    address: string,
    options: { sessionId?: string; shift?: "Up" | "Left" } = {},
    driveId?: string
  ): Promise<void> {
    if (!itemId) throw new Error("itemId mancante");
    if (!sheetName) throw new Error("sheetName mancante");
    if (!address) throw new Error("address mancante");
    const client = await this.getClient();
    const base = this.buildDriveBase(driveId);
    const sheetSegment = encodeURIComponent(sheetName);
    let req = client.api(`${base}/items/${itemId}/workbook/worksheets/${sheetSegment}/range(address='${address}')/delete`);
    if (options.sessionId) {
      req = req.header("workbook-session-id", options.sessionId);
    }
    await req.post({ shift: options.shift || "Up" });
  }
}
