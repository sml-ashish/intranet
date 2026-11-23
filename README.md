# SML Cosmos Intranet

SharePoint Online company intranet for **SML Cosmos — "Your Digital Universe."**
Built with SharePoint Framework (SPFx 1.24 + React) and a PnP PowerShell provisioning template.

- `spfx/` — SPFx solution with the **SML Cosmos** web part.
- `provisioning/` — creates the communication site, lists and sample content.
- `docs/` — architecture and deployment notes.

## What it builds

A space-themed, animated SharePoint **Communication site** with:

- An animated **hero** (gradient, floating glow, twinkling stars) with today's date,
  the site tagline and CTA buttons.
- **Quick links** grid with staggered reveal and hover animations.
- **Upcoming events** that can come from the signed-in user's **Outlook calendar**
  (Microsoft Graph), the `Company Events` list, or both.
- **Featured video** section that plays a YouTube link configured in the property pane.
- **Company news** cards from the `Company News` list, with image zoom on hover.
- **Logo slot** for your company logo (or a built-in SVG mark).
- Skeleton shimmer loaders and reveal-on-scroll for every section.
- `prefers-reduced-motion` support for accessibility.

## Prerequisites

- Microsoft 365 tenant with SharePoint Online.
- Node.js 22.14 LTS (recommended). SPFx 1.24 supports Node 18.17.1+, 20.11+ and
  22.14+; Node 24/26 are not supported by the build toolchain yet.
- PowerShell 7+ with the PnP.PowerShell module:
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```
- Permission to create sites and deploy apps in the tenant.

## Quick start

1. **Build the web part**

   ```bash
   cd spfx
   nvm use            # reads .nvmrc (Node 22.14)
   npm install
   npm run package
   ```

   Output: `spfx/sharepoint/solution/sml-cosmos-intranet.sppkg`.

2. **Create the site and apply the template**

   ```powershell
   cd provisioning
   ./deploy.ps1 `
     -TenantUrl "https://contoso.sharepoint.com" `
     -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
     -Owner     "admin@contoso.onmicrosoft.com"
   ```

3. **Deploy the app and add the web part** (same script, after packaging)

   ```powershell
   ./deploy.ps1 `
     -TenantUrl "https://contoso.sharepoint.com" `
     -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
     -Owner     "admin@contoso.onmicrosoft.com" `
     -SppkgPath "../spfx/sharepoint/solution/sml-cosmos-intranet.sppkg" `
     -AddWebPart
   ```

## Configure the web part

After adding the web part to `Home.aspx`, edit it:

| Property | Value |
| --- | --- |
| Hero title | `SML Cosmos` |
| Hero tagline | `Your Digital Universe.` |
| Company logo URL | Paste the logo URL from **Site Assets** (optional) |
| News list title | `Company News` |
| Company events list title | `Company Events` |
| Upcoming events source | My calendar / Company events / Both |
| YouTube link | Paste the video URL (watch/share/shorts/embed) |
| Video caption | Optional caption under the player |
| Quick links (JSON) | Optional array of `{ "label", "url", "icon" }` |

## Preview before deploying

You can see the finished design in your browser before anything touches
SharePoint:

```bash
cd spfx
nvm use
npm run preview     # builds preview/dist
python3 -m http.server 4399 --bind 127.0.0.1 --directory preview/dist
```

Then open **http://localhost:4399** — the real web part runs with demo news,
events (calendar + company) and animations, plus a "Preview mode" banner.

Optional query parameters:

```
http://localhost:4399/?video=https://youtu.be/VIDEO_ID   # test the video section
http://localhost:4399/?logo=https://.../logo.png         # test your company logo
```

The preview switches to demo content automatically when it detects `localhost`,
a `file://` page, or the SPFx `debugManifestsFile` workbench parameter — on a
deployed site it always uses live SharePoint data.

## Local development

```bash
cd spfx
nvm use
npm install
npm run serve
```

Open `https://<tenant>.sharepoint.com/_layouts/workbench.aspx` and add the web part.

## Docs

- `docs/backend.md` — everything to connect from the backend (lists, Graph, logo).
- `docs/deployment.md` — deployment runbook and App Catalog permissions.
- `docs/architecture.md` — design and extension points.
