# Architecture

## Components

```
+---------------------------- Microsoft 365 tenant ----------------------------+
|                                                                              |
|  SharePoint site: /sites/cosmos                                    |
|                                                                              |
|   Home.aspx                                                                  |
|     +--------------------------------------------------+                     |
|     | SML Cosmos Intranet web part (SPFx + React)  |                     |
|     |   hero + quick links + events + video + news     |                     |
|     +--------------------------------------------------+                     |
|            | SPHttpClient (REST)              ^                              |
|            v                                  |                              |
|   Lists/CompanyNews  -------------------------+                              |
|   Lists/CompanyEvents ------------------------+                              |
|   Lists/QuickLinks (optional content store)                                  |
|                                                                              |
|  Tenant App Catalog: sml-cosmos-intranet.sppkg                           |
+------------------------------------------------------------------------------+
```

## Why this split

- **SharePoint site + lists** hold content and permissions. Business users can edit
  news and events without touching code.
- **SPFx web part** owns presentation and client-side logic. It runs inside
  SharePoint with automatic SSO and theme integration.
- **PnP PowerShell template** makes the site reproducible across tenants
  (dev/test/prod) instead of manual clicks.

## Data flow

1. `deploy.ps1` creates the communication site and applies `site-template.xml`.
2. The SPFx package is uploaded to the app catalog and installed on the site.
3. The web part calls:
   ```
   GET /sites/cosmos/_api/web/lists/getbytitle('Company News')/items
       ?$select=Id,Title,NewsSummary,NewsLink,NewsImage,Modified
       &$orderby=Modified desc&$top=6

   GET /sites/cosmos/_api/web/lists/getbytitle('Company Events')/items
       ?$select=Id,Title,EventDate,EventEndDate,EventLocation,EventDescription,RegistrationLink
       &$orderby=EventDate asc&$top=20
   ```
4. Events are filtered client-side to today/future and limited to 6 cards.
5. The featured video is rendered from the YouTube URL configured in the property
   pane (`youtube.com/watch`, `youtu.be`, `shorts` and `embed` URLs are supported).

No Microsoft Graph permission is required because requests are scoped to the
current site.

## Branding

- Red + light green palette (`#dc2626` primary red, `#86efac` light green)
  applied through local CSS variables in `CompanyIntranet.module.scss`.
- Inline SVG brand mark (`BrandLogo.tsx`) so the web part works before a logo
  is configured.
- Optional custom logo via the **Company logo URL** web part property
  (upload to **Site Assets** — see `docs/backend.md`).
- Theme-aware fallbacks are kept minimal; the web part intentionally keeps the
  SML Cosmos identity regardless of the SharePoint theme.

## Upcoming events sources

| Source | Transport | Permission |
| --- | --- | --- |
| Signed-in user's Outlook calendar | Microsoft Graph `GET /me/calendarView` | `Calendars.Read` (approved in API access) |
| Company events list | SharePoint REST `/_api/web/lists/getbytitle(...)/items` | Site permissions |

Selected by the **Upcoming events source** property (`calendar`, `company`, `both`).

## Extension points

| Need | Approach |
| --- | --- |
| Read news from the site's Pages library instead of a list | Use `/_api/sitepages` or Microsoft Graph `sitePage` endpoints |
| Pull data from other sites | Add a `WebPartContext.msGraphClientFactory` call and approve Graph permissions in `package-solution.json` |
| Approvals / notifications | Power Automate flow on the `Company News` or `Company Events` list |
| Org chart / people search | Microsoft Graph `users` + `@microsoft/microsoft-graph-client` |
| Multi-language UI | Add locale files under `spfx/src/webparts/companyIntranet/loc/` |
| CI/CD | `npm run package` in GitHub Actions or Azure DevOps, then `Add-PnPApp` |
| Multiple videos | Replace the single `youtubeUrl` property with a `Video Library` list |

## Permissions

The current web part only needs the signed-in user's SharePoint permissions on the
site. If you add Microsoft Graph calls, declare the scopes in
`config/package-solution.json` under `webApiPermissionRequests` and approve them in
the SharePoint admin center.
