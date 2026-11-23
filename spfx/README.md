# SML Cosmos Intranet - SPFx web part

SharePoint Framework (SPFx 1.24) solution containing the **SML Cosmos Intranet**
web part. It renders a red-branded hero, quick links, upcoming events, a featured
YouTube video and company news from SharePoint lists.

## Prerequisites

- Node.js 22.14 LTS (recommended). The SPFx 1.24 build toolchain supports
  Node 18.17.1+, 20.11+ and 22.14+; Node 24/26 are not supported yet.
- A Microsoft 365 tenant with SharePoint Online.
- Permission to deploy tenant-scoped or site-scoped apps.

## Install and run locally

```bash
nvm use        # reads .nvmrc (Node 22.14)
npm install
npm run serve
```

Then open `https://<tenant>.sharepoint.com/_layouts/workbench.aspx` and add the
**SML Cosmos Intranet** web part.

To point the workbench at your tenant, edit `config/serve.json` and replace
`{tenantDomain}` with your tenant domain.

## Build the deployment package

```bash
npm run package
```

The package is written to `sharepoint/solution/sml-cosmos-intranet.sppkg`.

Upload it to the tenant App Catalog (or a site App Catalog) and add the app to
the intranet site. The web part then appears in the web part toolbox.

## Web part properties

| Property | Description |
| --- | --- |
| Hero title | Main heading in the hero banner |
| Hero subtitle | Supporting text in the hero banner |
| News list title | SharePoint list with news items (default `Company News`) |
| Events list title | SharePoint list with upcoming events (default `Company Events`) |
| YouTube link | Watch/share/shorts/embed URL; leave empty to hide the video |
| Video caption | Optional caption under the player |
| Quick links (JSON) | Optional array of `{ "label", "url", "icon" }` |

## List schemas

### Company News

| Column | Type | Purpose |
| --- | --- | --- |
| `Title` | Single line of text | News headline |
| `NewsSummary` | Multiple lines of text | Short summary |
| `NewsLink` | Hyperlink | Target page or external URL |
| `NewsImage` | Single line of text | Card image URL |

### Company Events

| Column | Type | Purpose |
| --- | --- | --- |
| `Title` | Single line of text | Event name |
| `EventDate` | Date and time | Start date/time |
| `EventEndDate` | Date and time | Optional end date/time |
| `EventLocation` | Single line of text | Venue |
| `EventDescription` | Multiple lines of text | Event details |
| `RegistrationLink` | Hyperlink | Optional registration URL |

The `provisioning/` folder creates both lists and seeds sample data.
