# SharePoint Multitenant MCP

Local Codex plugin for working with SharePoint and Excel files across separate Microsoft 365 customer tenants. Each customer has an independent single-tenant Entra app registration. Authentication is delegated: the MCP can never do more than the Microsoft account currently connected.

## What it supports

- Separate app registrations, connections, selection, and local disconnection for multiple Entra tenants.
- Search SharePoint sites, browse document libraries and folders, and inspect Excel tables.
- Read table values and add an Excel table column after an explicit confirmation.

The source is ready for additional Graph tools (SharePoint list items, list columns, cell/range updates) without changing the authentication design.

## One-time Microsoft Entra setup

Create one **single-tenant app registration for each customer tenant**. Do not reuse a client ID between customers:

1. In that customer's Microsoft Entra admin center, register an application with **Accounts in this organizational directory only**.
2. Add `http://localhost:43123/callback` as a redirect URI under **Mobile and desktop applications**.
3. Add delegated Microsoft Graph permissions: `User.Read`, `Files.ReadWrite`, `Sites.ReadWrite.All`, and, only if you need to change SharePoint list columns, `Sites.Manage.All`.
4. A tenant administrator grants consent for that customer only.

`Sites.Manage.All` is intentionally broad. Do not request it for installations that only edit Excel files in document libraries.

## Local setup

```bash
cd plugins/sharepoint-multitenant
npm install
npm run build
mkdir -p ~/.sharepoint-multitenant-mcp
cp config.example.json ~/.sharepoint-multitenant-mcp/config.json
chmod 600 ~/.sharepoint-multitenant-mcp/config.json
```

Copy a `clients[]` entry for every customer into `~/.sharepoint-multitenant-mcp/config.json`, with that customer's **Directory (tenant) ID** and **Application (client) ID**. Neither value is a client secret; this local app uses Authorization Code + PKCE. This local configuration and encrypted per-user refresh tokens stay outside this repository, so they survive plugin upgrades and are never shared with colleagues. Never commit tokens or a client secret.

Install the plugin in Codex from this local plugin folder, then start a new Codex chat. Ask Codex to connect a tenant. The `connect_tenant` tool returns a Microsoft sign-in URL; open it in your browser and then use `list_connections`.

## Account switching and logout

- `list_client_profiles`: shows the locally configured customer applications.
- `connect_tenant`: starts sign-in only for the chosen customer app and saves its account separately.
- `select_connection`: chooses the tenant used by subsequent tools.
- `disconnect_tenant`: deletes the encrypted local refresh token for one tenant. It does not log the account out of unrelated Microsoft browser sessions.

## Security model

Write tools always require `confirm: true`, after the user has approved the exact target and change. Tokens are encrypted at rest with a locally generated key and files are created with user-only permissions where the operating system supports them. For a team distribution, publish this folder in a private repository and let each colleague use their own local connection store.
