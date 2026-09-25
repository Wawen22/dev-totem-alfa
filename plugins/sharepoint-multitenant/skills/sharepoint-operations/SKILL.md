---
name: sharepoint-operations
description: Work with the user's connected Microsoft SharePoint tenant through the local SharePoint Multitenant MCP.
---

# SharePoint Multitenant operations

Use this MCP only after the user explicitly identifies or selects the intended tenant. Start by calling `list_connections` and, when needed, `select_connection`.

For Excel workbooks, discover the document library and file before reading a table. Before any modifying tool, state the precise tenant, site, file or list, and proposed change. Call a write tool only after the user has explicitly confirmed it; pass `confirm: true` only for that approved operation.

Never infer that two customer tenants are interchangeable. A user may only access resources permitted to the connected Microsoft account.
