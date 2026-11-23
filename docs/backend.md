# Backend connections — what to connect

Everything the web part shows comes from one of the sources below. You need to
**create the lists** (already done by `provisioning/`), **approve one Graph
permission**, and optionally **upload a logo**.

## 1. SharePoint lists (content you maintain)

Created automatically by `provisioning/deploy.ps1`:

| List | Columns | Powers |
| --- | --- | --- |
| `Company News` | `Title`, `NewsSummary`, `NewsLink`, `NewsImage` | Company news cards |
| `Company Events` | `Title`, `EventDate`, `EventEndDate`, `EventLocation`, `EventDescription`, `RegistrationLink` | Company events section |
| `Quick Links` | `Title`, `LinkUrl`, `LinkIcon` | Optional content-managed links |

Manage them here (replace the site URL):

```
https://YOURTENANT.sharepoint.com/sites/YOURSITE/Lists/CompanyNews
https://YOURTENANT.sharepoint.com/sites/YOURSITE/Lists/CompanyEvents
https://YOURTENANT.sharepoint.com/sites/YOURSITE/Lists/QuickLinks
```

> To show quick links from the list instead of the property pane JSON, that list
> read is a small code change — ask me and I'll wire it.

## 2. Microsoft Graph — the signed-in user's Outlook calendar

This is what makes **"Upcoming events" show the user's real meeting invites**.

| Item | Value |
| --- | --- |
| Endpoint | `GET /me/calendarView?startDateTime=...&endDateTime=...` |
| Requested scope | `Calendars.Read` (delegated) |
| Lookahead | 30 days |
| Declared in | `spfx/config/package-solution.json` → `webApiPermissionRequests` |

Each invite is shown with a **"My calendar"** pill, time, location, body preview and
an **Open invite** link to Outlook.

### Approval required (do this once)

1. Build and upload the `.sppkg` to the App Catalog.
2. SharePoint admin center → **Advanced** (or **More features**) → **API access**.
3. Find the pending request for **Microsoft Graph / Calendars.Read**.
4. Click **Approve**.

Until it is approved, the events card shows a friendly message. You can still
switch **Upcoming events source** to *Company events list* and it will work.

PowerShell alternative:

```powershell
Connect-PnPOnline -Url "https://YOURTENANT-admin.sharepoint.com" -Interactive
Approve-PnPSiteCollectionAppCatalogPermission -Scope Calendars.Read -Id <permission-guid>
```

### Switch the source

In the web part properties → **Content & events → Upcoming events source**:

| Value | Behaviour |
| --- | --- |
| `My calendar (Outlook invites)` | Graph only |
| `Company events list` | SharePoint list only |
| `Both` | Merged, sorted by date, top 6 (default) |

## 3. Company logo

Two options — both work with the same web part property.

### Option A (recommended): Site Assets

1. Site → **Site contents → Site Assets → New folder** `Brand` (or use the root).
2. **Upload** your logo (PNG/SVG, square works best, `logo.png`).
3. Click the file → **Copy link** (or copy the URL from the address bar).
   Expected shape:
   ```
   https://YOURTENANT.sharepoint.com/sites/YOURSITE/SiteAssets/Brand/logo.png
   ```
4. Edit the web part → **Branding → Company logo URL** → paste → save.

The logo is rendered inside a white rounded chip over the hero gradient, so a
**transparent PNG or SVG** looks best. Recommended: at least 200×200 px.

> If the image shows a red X, the URL is wrong or the file is not shared with
> "Everyone who can view this site". Re-copy the link.

### Option B: package it into the app

Put the file at `spfx/src/webparts/companyIntranet/assets/logo.png` and reference it
from `BrandLogo.tsx`. Requires a rebuild/redeploy to change the logo.

Leave **Company logo URL** empty and the built-in **SML Cosmos orbit mark** is used.

## 4. YouTube video

| Item | Value |
| --- | --- |
| Data | Just a URL pasted in the property pane |
| Supported | `watch?v=`, `youtu.be/`, `/embed/`, `/shorts/`, or an 11-char id |
| Embed | `https://www.youtube-nocookie.com/embed/<id>` |
| Empty | The section hides and shows a placeholder in edit mode |

Backend alternative: store the URL in a SharePoint list setting and read it in
`onInit()` — useful if marketing changes it often. Ask if you want that.

## 5. Theme / colour

The palette lives at the top of
`spfx/src/webparts/companyIntranet/components/CompanyIntranet.module.scss`:

```scss
--c-red:          #dc2626;   /* primary red */
--c-red-dark:     #991b1b;
--c-red-deep:     #450a0a;   /* hero gradient start */
--c-green:        #22c55e;   /* accent green */
--c-green-light:  #86efac;   /* light green */
--c-green-soft:   #dcfce7;
```

Change those values to re-brand the whole hero, buttons, pills and cards.

## Connectivity checklist

| # | Step | Who |
| --- | --- | --- |
| 1 | Build + upload `.sppkg` to App Catalog | App Catalog owner |
| 2 | Approve **Calendars.Read** in API access | SharePoint admin |
| 3 | Install app on the site | Site owner |
| 4 | Create lists (or run `deploy.ps1`) | Site owner |
| 5 | Upload logo to **Site Assets**, paste URL | Site owner |
| 6 | Paste YouTube URL, publish page | Site owner |
