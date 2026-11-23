# Deployment and app permissions

## Where files go

| Artifact | Location | Notes |
| --- | --- | --- |
| `sml-cosmos-intranet.sppkg` | **Tenant App Catalog** (or site collection App Catalog) | One package serves every site. Re-upload here for updates. |
| The intranet site | SharePoint site URL you choose, e.g. `https://contoso.sharepoint.com/sites/cosmos` | Created once. The package is not stored here. |

So: **same App Catalog location, same site URL.** Creating a "new site" means a new
site URL, but the `.sppkg` still goes to the same App Catalog.

## First-time deployment

1. Build the package:
   ```bash
   cd spfx
   nvm use
   npm run package
   ```
   Output: `spfx/sharepoint/solution/sml-cosmos-intranet.sppkg`.

2. Create the site and content:
   ```powershell
   cd provisioning
   ./deploy.ps1 `
     -TenantUrl "https://contoso.sharepoint.com" `
     -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
     -Owner     "admin@contoso.onmicrosoft.com" `
     -SkipAppDeployment
   ```

3. Admin deploys the app:
   ```powershell
   ./deploy.ps1 `
     -TenantUrl "https://contoso.sharepoint.com" `
     -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
     -Owner     "admin@contoso.onmicrosoft.com" `
     -SppkgPath "../spfx/sharepoint/solution/sml-cosmos-intranet.sppkg" `
     -AddWebPart
   ```

4. Open `Home.aspx`, edit the web part, set the YouTube link, and publish.

## Updating the package

1. Bump `solution.version` in `spfx/config/package-solution.json`
   (e.g. `1.0.0.0` → `1.0.1.0`).
2. Run `npm run package`.
3. Re-upload the `.sppkg` to the **same App Catalog** and overwrite the app.

## How to stop asking the admin every time

You have three options, from simplest to most automated.

### Option 1 — Become an App Catalog owner (recommended)

Ask the SharePoint admin to add you as an **owner of the App Catalog site**. Then you
can upload/overwrite `.sppkg` files yourself, without full tenant admin rights.

Admin steps:
1. SharePoint admin center → **More features → Apps → Open** (App Catalog).
2. App Catalog site → **Site permissions** → **Site owners** → **Add owners**.
3. Add your account.

### Option 2 — Site collection app catalog

Ask the admin to enable a **site collection App Catalog** for the intranet site and
make you a site collection admin. You can then upload packages scoped to that site.

Admin command:
```powershell
Add-PnPSiteCollectionAppCatalog -Site "https://contoso.sharepoint.com/sites/cosmos"
```

### Option 3 — CI/CD with an Entra app (no human upload)

For a repeatable pipeline, the admin creates an Entra app registration with
SharePoint **Sites.Selected** permission and grants it access to the App Catalog
site. Store the certificate in GitHub/Azure DevOps, then the pipeline runs
`Add-PnPApp` on every release. Admin consent is needed once.

## Email template to request App Catalog access

> Subject: Request — App Catalog owner access for SML Cosmos intranet
>
> Hi [Admin],
>
> We are deploying the SML Cosmos intranet SharePoint Framework package
> (`sml-cosmos-intranet.sppkg`). To avoid raising a request for every release,
> could you please:
>
> 1. Add me as an **owner of the tenant App Catalog site** (SharePoint admin center
>    → More features → Apps → App Catalog → Site permissions → Site owners).
> 2. Confirm the intranet site URL: `https://[tenant].sharepoint.com/sites/cosmos`.
>
> The package only uses the signed-in user's SharePoint permissions on the intranet
> site; it does not request Microsoft Graph scopes. I will keep the package version
> bumped on each release and upload it to the same App Catalog.
>
> Thanks,
> [Your name]

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| "NodeJS version ... does not meet the requirements" | Run `nvm use` (Node 22.14) before `npm run package`. |
| Web part missing from toolbox | App not installed on the site, or page not republished. |
| "Unable to load the news/events list" | List name in the web part properties does not match the site list. |
| YouTube section empty | No valid YouTube URL in the property pane. |
| Package upload rejected | Bump `solution.version` and re-upload. |
