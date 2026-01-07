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
    
    // Initial request with max page size (999)
    let nextLink: string | undefined = undefined;
    
    // First call
    let response = await client
      .api(`/sites/${this.siteId}/lists/${listId}/items`)
      .expand("fields")
      .top(999) 
      .orderby("createdDateTime desc")
      .get();

    while (response) {
      const items = (response.value || []).map((item: any) => ({
        id: item.id,
        fields: item.fields,
      }));
      allItems = allItems.concat(items);

      nextLink = response["@odata.nextLink"];
      if (nextLink) {
        response = await client.api(nextLink).get();
      } else {
        response = null;
      }
    }

    // Deduplicate by ID to ensure clean data
    const uniqueItems = new Map<string, SharePointListItem<TFields>>();
    allItems.forEach(item => uniqueItems.set(item.id, item));
    
    return Array.from(uniqueItems.values());
  }

  async createItem<TFields extends Record<string, unknown>>(listId: string, fields: TFields): Promise<SharePointListItem<TFields>> {
    if (!listId) throw new Error("listId mancante");
    const client = await this.getClient();

    const response = await client
      .api(`/sites/${this.siteId}/lists/${listId}/items`)
      .post({ fields });

    return { id: response.id, fields: response.fields };
  }

  async updateItem<TFields extends Record<string, unknown>>(listId: string, itemId: string, fields: Partial<TFields>): Promise<void> {
    if (!listId) throw new Error("listId mancante");
    if (!itemId) throw new Error("itemId mancante");
    const client = await this.getClient();

    await client
      .api(`/sites/${this.siteId}/lists/${listId}/items/${itemId}/fields`)
      .patch(fields);
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
}
