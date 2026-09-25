import { createCipheriv, createDecipheriv, createHash, randomBytes } from "node:crypto";
import { createServer, type Server } from "node:http";
import { chmod, mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { McpServer } from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import * as z from "zod/v4";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const stateDirectory = join(homedir(), ".sharepoint-multitenant-mcp");
const tokenStorePath = join(stateDirectory, "connections.enc.json");
const masterKeyPath = join(stateDirectory, "master.key");

type TenantAppConfig = {
  id: string;
  label: string;
  tenantId: string;
  clientId: string;
  scopes: string[];
};
type Config = { redirectPort: number; clients: TenantAppConfig[] };
type Connection = {
  id: string;
  clientProfileId: string;
  tenantId: string;
  accountId: string;
  displayName: string;
  username: string;
  refreshToken: string;
};
type StoredConnections = { activeConnectionId?: string; connections: Record<string, Connection> };
type PendingSignIn = { state: string; verifier: string; server: Server; startedAt: string };

let pendingSignIn: PendingSignIn | undefined;

function result(value: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(value, null, 2) }] };
}

function failure(error: unknown) {
  const message = error instanceof Error ? error.message : String(error);
  return { content: [{ type: "text" as const, text: message }], isError: true };
}

type SharePointColumnType = "text" | "number" | "dateTime" | "boolean";

function sharePointColumnDefinition(column: {
  name: string;
  displayName?: string;
  description?: string;
  type: SharePointColumnType;
}) {
  return {
    name: column.name,
    ...(column.displayName ? { displayName: column.displayName } : {}),
    ...(column.description ? { description: column.description } : {}),
    [column.type]: column.type === "dateTime" ? { format: "dateOnly" } : {},
  };
}

function normalizedHeader(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
}

function normalizedValue(value: unknown) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLocaleUpperCase("it-IT");
}

function dateOnlyValue(value: unknown) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30 + Math.floor(value)));
    return date.toISOString().slice(0, 10);
  }
  const text = String(value).trim();
  const italian = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
  if (italian) return `${italian[3]}-${italian[2].padStart(2, "0")}-${italian[1].padStart(2, "0")}`;
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  throw new Error(`Invalid date value '${text}'.`);
}

function numberValue(value: unknown) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value).trim().replace(/\s/g, "");
  const comma = raw.lastIndexOf(",");
  const dot = raw.lastIndexOf(".");
  const text = comma >= 0 && dot >= 0
    ? comma > dot
      ? raw.replace(/\./g, "").replace(",", ".")
      : raw.replace(/,/g, "")
    : raw.replace(",", ".");
  const number = Number(text);
  if (!Number.isFinite(number)) throw new Error(`Invalid numeric value '${String(value)}'.`);
  return number;
}

async function ensureStateDirectory() {
  await mkdir(stateDirectory, { recursive: true, mode: 0o700 });
  await chmod(stateDirectory, 0o700).catch(() => undefined);
}

async function getMasterKey() {
  await ensureStateDirectory();
  try {
    return await readFile(masterKeyPath);
  } catch {
    const key = randomBytes(32);
    await writeFile(masterKeyPath, key, { mode: 0o600 });
    await chmod(masterKeyPath, 0o600).catch(() => undefined);
    return key;
  }
}

async function loadConnections(): Promise<StoredConnections> {
  try {
    const key = await getMasterKey();
    const encrypted = JSON.parse(await readFile(tokenStorePath, "utf8")) as { iv: string; tag: string; ciphertext: string };
    const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(encrypted.iv, "base64"));
    decipher.setAuthTag(Buffer.from(encrypted.tag, "base64"));
    const plaintext = Buffer.concat([decipher.update(Buffer.from(encrypted.ciphertext, "base64")), decipher.final()]);
    return JSON.parse(plaintext.toString("utf8")) as StoredConnections;
  } catch (error: unknown) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return { connections: {} };
    throw new Error("Unable to read the local connection store. Do not delete its master.key while connections exist.");
  }
}

async function saveConnections(store: StoredConnections) {
  const key = await getMasterKey();
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(JSON.stringify(store), "utf8"), cipher.final()]);
  const payload = JSON.stringify({ iv: iv.toString("base64"), tag: cipher.getAuthTag().toString("base64"), ciphertext: ciphertext.toString("base64") });
  const temporaryPath = `${tokenStorePath}.${process.pid}.tmp`;
  await writeFile(temporaryPath, payload, { mode: 0o600 });
  await chmod(temporaryPath, 0o600).catch(() => undefined);
  await rename(temporaryPath, tokenStorePath);
}

async function loadConfig(): Promise<Config> {
  const configPath = process.env.SHAREPOINT_MCP_CONFIG ?? join(stateDirectory, "config.json");
  let raw: unknown;
  try {
    raw = JSON.parse(await readFile(configPath, "utf8"));
  } catch {
    throw new Error(`Create ${configPath} from config.example.json and set clientId before connecting a tenant.`);
  }
  const parsed = z.object({
    redirectPort: z.number().int().min(1024).max(65535).default(43123),
    clients: z.array(z.object({
      id: z.string().regex(/^[a-z0-9-]+$/, "Client profile id may contain only lowercase letters, digits, and hyphens."),
      label: z.string().min(1),
      tenantId: z.string().uuid("Use the Directory (tenant) ID from this customer's Entra app registration."),
      clientId: z.string().uuid("Use the Application (client) ID from this customer's Entra app registration."),
      scopes: z.array(z.string().min(1)).min(1),
    })).min(1),
  }).safeParse(raw);
  if (!parsed.success) throw new Error(`Invalid MCP configuration: ${parsed.error.issues.map((issue) => issue.message).join(", ")}`);
  if (new Set(parsed.data.clients.map((client) => client.id)).size !== parsed.data.clients.length) throw new Error("Every client profile must have a unique id.");
  return parsed.data;
}

function clientProfile(config: Config, clientProfileId: string) {
  const profile = config.clients.find((candidate) => candidate.id === clientProfileId);
  if (!profile) throw new Error(`Client profile '${clientProfileId}' is not configured locally.`);
  return profile;
}

function authority(profile: TenantAppConfig) {
  return `https://login.microsoftonline.com/${encodeURIComponent(profile.tenantId)}/oauth2/v2.0`;
}

function decodeJwt(token: string): Record<string, unknown> {
  const payload = token.split(".")[1];
  if (!payload) throw new Error("Microsoft did not return a valid ID token.");
  return JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as Record<string, unknown>;
}

function redirectUri(config: Config) {
  return `http://localhost:${config.redirectPort}/callback`;
}

async function exchangeCode(config: Config, profile: TenantAppConfig, code: string, verifier: string) {
  const response = await fetch(`${authority(profile)}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: profile.clientId,
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri(config),
      code_verifier: verifier,
      scope: ["openid", "profile", "offline_access", ...profile.scopes].join(" "),
    }),
  });
  const body = await response.json() as Record<string, unknown>;
  if (!response.ok) throw new Error(`Microsoft sign-in failed: ${String(body.error_description ?? body.error ?? response.status)}`);
  if (typeof body.refresh_token !== "string" || typeof body.id_token !== "string") throw new Error("Microsoft did not return a refresh token. Ensure offline_access is enabled.");
  return body as { refresh_token: string; id_token: string };
}

async function refreshAccessToken(connection: Connection, config: Config) {
  const profile = clientProfile(config, connection.clientProfileId);
  const response = await fetch(`${authority(profile)}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: profile.clientId,
      grant_type: "refresh_token",
      refresh_token: connection.refreshToken,
      scope: ["openid", "profile", "offline_access", ...profile.scopes].join(" "),
    }),
  });
  const body = await response.json() as Record<string, unknown>;
  if (!response.ok || typeof body.access_token !== "string") throw new Error(`Microsoft token refresh failed for ${connection.username}. Reconnect this tenant. ${String(body.error_description ?? body.error ?? "")}`);
  if (typeof body.refresh_token === "string") {
    const store = await loadConnections();
    const saved = store.connections[connection.id];
    if (saved) {
      saved.refreshToken = body.refresh_token;
      await saveConnections(store);
    }
  }
  return body.access_token;
}

async function activeConnection(connectionId?: string) {
  const store = await loadConnections();
  const id = connectionId ?? store.activeConnectionId;
  if (!id) throw new Error("No tenant is selected. Connect a tenant, then select it.");
  const connection = store.connections[id];
  if (!connection) throw new Error("The selected tenant connection no longer exists. Choose another connection.");
  return connection;
}

async function graphRequest(path: string, init: RequestInit = {}, connectionId?: string) {
  const [connection, config] = await Promise.all([activeConnection(connectionId), loadConfig()]);
  const accessToken = await refreshAccessToken(connection, config);
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", ...(init.headers ?? {}) },
  });
  const text = await response.text();
  let body: unknown = text;
  try { body = text ? JSON.parse(text) : {}; } catch { /* Keep non-JSON Graph responses readable. */ }
  if (!response.ok) throw new Error(`Microsoft Graph ${response.status}: ${typeof body === "string" ? body : JSON.stringify(body)}`);
  return body;
}

async function startSignIn(clientProfileId: string) {
  if (pendingSignIn) throw new Error("A tenant sign-in is already waiting for its browser callback. Finish it or restart the MCP server.");
  const config = await loadConfig();
  const profile = clientProfile(config, clientProfileId);
  const state = randomBytes(24).toString("base64url");
  const verifier = randomBytes(48).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  const url = new URL(`${authority(profile)}/authorize`);
  url.search = new URLSearchParams({
    client_id: profile.clientId,
    response_type: "code",
    redirect_uri: redirectUri(config),
    response_mode: "query",
    scope: ["openid", "profile", "offline_access", ...profile.scopes].join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account",
  }).toString();

  const server = createServer(async (request, response) => {
    const callback = new URL(request.url ?? "/", redirectUri(config));
    if (callback.pathname !== "/callback") {
      response.writeHead(404).end("Not found");
      return;
    }
    if (callback.searchParams.get("state") !== state || !callback.searchParams.get("code")) {
      response.writeHead(400, { "Content-Type": "text/html" }).end("<h1>Sign-in failed</h1><p>Invalid or expired sign-in state. Return to Codex and start again.</p>");
      return;
    }
    try {
      const tokens = await exchangeCode(config, profile, callback.searchParams.get("code")!, verifier);
      const claims = decodeJwt(tokens.id_token);
      const tenantId = String(claims.tid ?? "");
      const accountId = String(claims.oid ?? claims.sub ?? "");
      if (!tenantId || !accountId) throw new Error("Microsoft did not identify the tenant and account.");
      if (tenantId !== profile.tenantId) throw new Error(`The signed-in account belongs to tenant ${tenantId}, not the configured customer tenant ${profile.tenantId}.`);
      const connection: Connection = {
        id: `${profile.id}:${accountId}`,
        clientProfileId: profile.id,
        tenantId,
        accountId,
        displayName: String(claims.name ?? claims.preferred_username ?? "Microsoft user"),
        username: String(claims.preferred_username ?? claims.email ?? accountId),
        refreshToken: tokens.refresh_token,
      };
      const store = await loadConnections();
      store.connections[connection.id] = connection;
      store.activeConnectionId = connection.id;
      await saveConnections(store);
      response.writeHead(200, { "Content-Type": "text/html" }).end(`<h1>Connected</h1><p>${connection.displayName} is now connected. You can close this tab and return to Codex.</p>`);
      pendingSignIn?.server.close();
      pendingSignIn = undefined;
    } catch (error) {
      response.writeHead(500, { "Content-Type": "text/html" }).end(`<h1>Sign-in failed</h1><p>${error instanceof Error ? error.message : String(error)}</p>`);
    }
  });
  await new Promise<void>((resolve, reject) => server.once("error", reject).listen(config.redirectPort, "127.0.0.1", resolve));
  pendingSignIn = { state, verifier, server, startedAt: new Date().toISOString() };
  return { authorizationUrl: url.toString(), callback: redirectUri(config), expiresNote: "Keep this Codex session open until the browser reports Connected." };
}

const server = new McpServer({ name: "sharepoint-multitenant", version: "0.1.0" });

server.registerTool("list_client_profiles", { description: "List the locally configured customer app registrations. No tokens or secrets are returned.", inputSchema: z.object({}) }, async () => {
  try {
    const config = await loadConfig();
    return result(config.clients.map(({ id, label, tenantId, scopes }) => ({ id, label, tenantId, scopes })));
  } catch (error) { return failure(error); }
});

server.registerTool("connect_tenant", { description: "Start Microsoft sign-in for one configured customer app registration. Open the returned URL in a browser.", inputSchema: z.object({ clientProfileId: z.string().min(1) }) }, async ({ clientProfileId }) => {
  try { return result(await startSignIn(clientProfileId)); } catch (error) { return failure(error); }
});

server.registerTool("list_connections", { description: "List locally connected Microsoft tenant accounts and indicate which one is active.", inputSchema: z.object({}) }, async () => {
  try {
    const store = await loadConnections();
    return result(Object.values(store.connections).map(({ refreshToken: _refreshToken, ...connection }) => ({ ...connection, active: connection.id === store.activeConnectionId })));
  } catch (error) { return failure(error); }
});

server.registerTool("select_connection", { description: "Select the connected tenant account to use for subsequent Microsoft Graph operations.", inputSchema: z.object({ connectionId: z.string().min(1) }) }, async ({ connectionId }) => {
  try {
    const store = await loadConnections();
    if (!store.connections[connectionId]) throw new Error("Unknown connectionId. Call list_connections first.");
    store.activeConnectionId = connectionId;
    await saveConnections(store);
    return result({ activeConnectionId: connectionId });
  } catch (error) { return failure(error); }
});

server.registerTool("disconnect_tenant", { description: "Delete one tenant's encrypted local refresh token. This does not sign the person out of Microsoft in other browser sessions.", annotations: { destructiveHint: true }, inputSchema: z.object({ connectionId: z.string().min(1), confirm: z.literal(true).describe("Must be true after the user explicitly confirms disconnecting this tenant.") }) }, async ({ connectionId }) => {
  try {
    const store = await loadConnections();
    const connection = store.connections[connectionId];
    if (!connection) throw new Error("Unknown connectionId.");
    delete store.connections[connectionId];
    if (store.activeConnectionId === connectionId) store.activeConnectionId = Object.keys(store.connections)[0];
    await saveConnections(store);
    return result({ disconnected: { id: connection.id, tenantId: connection.tenantId, username: connection.username }, activeConnectionId: store.activeConnectionId ?? null });
  } catch (error) { return failure(error); }
});

server.registerTool("search_sites", { description: "Search sites in the active SharePoint tenant.", inputSchema: z.object({ query: z.string().min(1), connectionId: z.string().optional() }) }, async ({ query, connectionId }) => {
  try { return result(await graphRequest(`/sites?search=${encodeURIComponent(query)}`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("list_document_libraries", { description: "List document libraries in a SharePoint site.", inputSchema: z.object({ siteId: z.string().min(1), connectionId: z.string().optional() }) }, async ({ siteId, connectionId }) => {
  try { return result(await graphRequest(`/sites/${encodeURIComponent(siteId)}/drives`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("list_library_folder", { description: "List files and folders in the root or a folder of a SharePoint document library.", inputSchema: z.object({ driveId: z.string().min(1), itemId: z.string().optional(), connectionId: z.string().optional() }) }, async ({ driveId, itemId, connectionId }) => {
  try {
    const path = itemId ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children` : `/drives/${encodeURIComponent(driveId)}/root/children`;
    return result(await graphRequest(path, {}, connectionId));
  } catch (error) { return failure(error); }
});

server.registerTool("list_excel_tables", { description: "List named Excel tables in a workbook stored in a SharePoint document library.", inputSchema: z.object({ driveId: z.string().min(1), itemId: z.string().min(1), connectionId: z.string().optional() }) }, async ({ driveId, itemId, connectionId }) => {
  try { return result(await graphRequest(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/tables`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("read_excel_table", { description: "Read the values and address of an Excel table in a SharePoint workbook.", inputSchema: z.object({ driveId: z.string().min(1), itemId: z.string().min(1), tableIdOrName: z.string().min(1), connectionId: z.string().optional() }) }, async ({ driveId, itemId, tableIdOrName, connectionId }) => {
  try { return result(await graphRequest(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/tables/${encodeURIComponent(tableIdOrName)}/range`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("add_excel_table_column", { description: "Insert a named column in an Excel table. First show the target and index to the user, then call only after explicit confirmation.", annotations: { destructiveHint: true }, inputSchema: z.object({ driveId: z.string().min(1), itemId: z.string().min(1), tableIdOrName: z.string().min(1), name: z.string().min(1), index: z.number().int().nonnegative(), values: z.array(z.array(z.unknown())).optional(), connectionId: z.string().optional(), confirm: z.literal(true).describe("Must be true only after the user has explicitly confirmed this exact Excel change.") }) }, async ({ driveId, itemId, tableIdOrName, name, index, values, connectionId }) => {
  try {
    return result(await graphRequest(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/tables/${encodeURIComponent(tableIdOrName)}/columns/add`, { method: "POST", body: JSON.stringify({ name, index, ...(values ? { values } : {}) }) }, connectionId));
  } catch (error) { return failure(error); }
});

server.registerTool("list_sharepoint_lists", { description: "List SharePoint lists in a site.", inputSchema: z.object({ siteId: z.string().min(1), connectionId: z.string().optional() }) }, async ({ siteId, connectionId }) => {
  try { return result(await graphRequest(`/sites/${encodeURIComponent(siteId)}/lists`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("list_sharepoint_list_columns", { description: "List the column definitions of a SharePoint list.", inputSchema: z.object({ siteId: z.string().min(1), listId: z.string().min(1), connectionId: z.string().optional() }) }, async ({ siteId, listId, connectionId }) => {
  try { return result(await graphRequest(`/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`, {}, connectionId)); } catch (error) { return failure(error); }
});

server.registerTool("create_sharepoint_list", {
  description: "Create a generic SharePoint list with typed columns. State the exact tenant, site, list and schema first; requires explicit confirmation.",
  annotations: { destructiveHint: true },
  inputSchema: z.object({
    siteId: z.string().min(1),
    displayName: z.string().min(1),
    description: z.string().optional(),
    columns: z.array(z.object({
      name: z.string().regex(/^[A-Za-z][A-Za-z0-9_]*$/, "Use an internal name beginning with a letter and containing letters, numbers, or underscores."),
      displayName: z.string().min(1).optional(),
      description: z.string().optional(),
      type: z.enum(["text", "number", "dateTime", "boolean"]),
    })).max(60),
    connectionId: z.string().optional(),
    confirm: z.literal(true).describe("Must be true only after the user has explicitly confirmed this exact list and schema creation."),
  }),
}, async ({ siteId, displayName, description, columns, connectionId }) => {
  try {
    const names = columns.map((column) => column.name.toLowerCase());
    if (new Set(names).size !== names.length) throw new Error("Every SharePoint column name must be unique.");
    const body = {
      displayName,
      ...(description ? { description } : {}),
      columns: columns.map(sharePointColumnDefinition),
      list: { template: "genericList" },
    };
    return result(await graphRequest(`/sites/${encodeURIComponent(siteId)}/lists`, { method: "POST", body: JSON.stringify(body) }, connectionId));
  } catch (error) { return failure(error); }
});

server.registerTool("create_sharepoint_list_column", { description: "Create a text, number, or yes/no column in a SharePoint list. Requires Sites.Manage.All and explicit confirmation.", annotations: { destructiveHint: true }, inputSchema: z.object({ siteId: z.string().min(1), listId: z.string().min(1), name: z.string().regex(/^[A-Za-z][A-Za-z0-9_]*$/, "Use an internal name beginning with a letter and containing letters, numbers, or underscores."), displayName: z.string().min(1).optional(), type: z.enum(["text", "number", "boolean"]), description: z.string().optional(), connectionId: z.string().optional(), confirm: z.literal(true).describe("Must be true only after the user has explicitly confirmed this exact SharePoint schema change.") }) }, async ({ siteId, listId, name, displayName, type, description, connectionId }) => {
  try {
    const body = sharePointColumnDefinition({ name, displayName, description, type });
    return result(await graphRequest(`/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`, { method: "POST", body: JSON.stringify(body) }, connectionId));
  } catch (error) { return failure(error); }
});

server.registerTool("create_sharepoint_list_items", {
  description: "Create multiple items in an existing SharePoint list. Use only for a reviewed import; the result reports every failed row without deleting prior successful rows.",
  annotations: { destructiveHint: true },
  inputSchema: z.object({
    siteId: z.string().min(1),
    listId: z.string().min(1),
    items: z.array(z.object({ fields: z.record(z.string(), z.unknown()) })).min(1).max(100),
    connectionId: z.string().optional(),
    confirm: z.literal(true).describe("Must be true only after the user has explicitly confirmed this exact item import."),
  }),
}, async ({ siteId, listId, items, connectionId }) => {
  const created: Array<{ index: number; id: string }> = [];
  const failed: Array<{ index: number; error: string }> = [];
  for (const [index, item] of items.entries()) {
    try {
      const response = await graphRequest(
        `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
        { method: "POST", body: JSON.stringify({ fields: item.fields }) },
        connectionId
      ) as { id?: string };
      created.push({ index, id: String(response.id ?? "") });
    } catch (error) {
      failed.push({ index, error: error instanceof Error ? error.message : String(error) });
    }
  }
  return result({ requested: items.length, created, failed });
});

server.registerTool("import_excel_table_to_sharepoint_list", {
  description: "Import non-empty rows from a reviewed Excel table into an empty SharePoint list. Converts selected date and number columns, skips duplicate identity keys, and never updates existing items.",
  annotations: { destructiveHint: true },
  inputSchema: z.object({
    driveId: z.string().min(1),
    itemId: z.string().min(1),
    tableIdOrName: z.string().min(1),
    siteId: z.string().min(1),
    listId: z.string().min(1),
    columnMap: z.record(z.string(), z.string().regex(/^[A-Za-z][A-Za-z0-9_]*$/)),
    dateFields: z.array(z.string()).default([]),
    numberFields: z.array(z.string()).default([]),
    identityFields: z.array(z.string()).min(1),
    connectionId: z.string().optional(),
    confirm: z.literal(true).describe("Must be true only after the user has explicitly confirmed this exact Excel-to-SharePoint initial import."),
  }),
}, async ({ driveId, itemId, tableIdOrName, siteId, listId, columnMap, dateFields, numberFields, identityFields, connectionId }) => {
  try {
    const range = await graphRequest(
      `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/tables/${encodeURIComponent(tableIdOrName)}/range`,
      {},
      connectionId
    ) as { values?: unknown[][] };
    const [headers, ...rows] = range.values ?? [];
    if (!headers || !Array.isArray(headers)) throw new Error("The Excel table has no header row.");

    const mappedHeaders = headers.map((header) => columnMap[normalizedHeader(header)]);
    if (!mappedHeaders.includes("Title")) throw new Error("The columnMap must map the Excel code column to SharePoint Title.");
    const allowedDateFields = new Set(dateFields);
    const allowedNumberFields = new Set(numberFields);
    const candidates: Array<{ sourceRow: number; fields: Record<string, unknown>; identity: string }> = [];
    const duplicateIdentities = new Set<string>();
    const knownIdentities = new Set<string>();
    const skippedEmpty: number[] = [];

    rows.forEach((row, rowIndex) => {
      const fields: Record<string, unknown> = {};
      mappedHeaders.forEach((field, columnIndex) => {
        if (!field) return;
        const raw = row[columnIndex];
        if (allowedDateFields.has(field)) fields[field] = dateOnlyValue(raw);
        else if (allowedNumberFields.has(field)) fields[field] = numberValue(raw);
        else fields[field] = raw === null || raw === undefined || raw === "" ? null : String(raw).trim();
      });
      const title = String(fields.Title ?? "").trim();
      if (!title) {
        skippedEmpty.push(rowIndex + 2);
        return;
      }
      fields.Title = title;
      const identity = identityFields.map((field) => normalizedValue(fields[field])).join("::");
      if (knownIdentities.has(identity)) {
        duplicateIdentities.add(identity);
        return;
      }
      knownIdentities.add(identity);
      candidates.push({ sourceRow: rowIndex + 2, fields, identity });
    });

    const importRows = candidates.filter((candidate) => !duplicateIdentities.has(candidate.identity));
    const created: Array<{ sourceRow: number; id: string }> = [];
    const failed: Array<{ sourceRow: number; error: string }> = [];
    for (const candidate of importRows) {
      try {
        const response = await graphRequest(
          `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`,
          { method: "POST", body: JSON.stringify({ fields: candidate.fields }) },
          connectionId
        ) as { id?: string };
        created.push({ sourceRow: candidate.sourceRow, id: String(response.id ?? "") });
      } catch (error) {
        failed.push({ sourceRow: candidate.sourceRow, error: error instanceof Error ? error.message : String(error) });
      }
    }
    return result({
      excelRows: rows.length,
      requested: importRows.length,
      created,
      skippedEmptyRows: skippedEmpty,
      skippedDuplicateIdentities: Array.from(duplicateIdentities),
      failed,
    });
  } catch (error) { return failure(error); }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("SharePoint Multitenant MCP running over stdio");
}

void main();
