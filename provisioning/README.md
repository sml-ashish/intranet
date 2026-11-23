# Provisioning

`deploy.ps1` creates the SML Cosmos intranet communication site and applies
`site-template.xml`.

## Prerequisites

- PowerShell 7+
- PnP.PowerShell module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## Usage

Create the site and apply the template:

```powershell
./deploy.ps1 `
  -TenantUrl "https://contoso.sharepoint.com" `
  -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
  -Owner     "admin@contoso.onmicrosoft.com"
```

Create the site, deploy the SPFx package and add the web part to Home.aspx:

```powershell
./deploy.ps1 `
  -TenantUrl "https://contoso.sharepoint.com" `
  -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
  -Owner     "admin@contoso.onmicrosoft.com" `
  -SppkgPath "../spfx/sharepoint/solution/sml-cosmos-intranet.sppkg" `
  -AddWebPart
```

## What the template creates

| Artifact | Type | Purpose |
| --- | --- | --- |
| `Company News` | Custom list | News cards on the landing page |
| `Company Events` | Custom list | Upcoming events section |
| `Quick Links` | Custom list | Optional content-managed links |
| `Company Stats` | Custom list | Key statistics for the landing page |
| `Featured Apps` | Custom list | Featured application tiles |
| `Departments` | Custom list | Department directory |
| `Client Logos` | Picture library | Staff-uploaded client logos |
| `NewsSummary` | Site column (Note) | Short summary on the news card |
| `NewsLink` | Site column (URL) | Card target |
| `NewsImage` | Site column (Text) | Card image URL |
| `EventDate` / `EventEndDate` | Site columns (DateTime) | Event schedule |
| `EventLocation` | Site column (Text) | Venue |
| `EventDescription` | Site column (Note) | Event details |
| `RegistrationLink` | Site column (URL) | Optional registration URL |
| `LinkUrl` / `LinkIcon` | Site columns | Quick links data |
| `Value` / `Icon` | Site columns (Text) | Stat value and Fluent UI icon name |
| `InHero` | Site column (Boolean, display name "In hero") | Show the stat in the hero section |
| `Description` | Site column (Note, plain text) | App / department summary |
| `Url` | Site column (URL) | App / department link |
| `Head` | Site column (Text) | Department head / contact |

The template is idempotent and can be re-applied to update fields and lists.
`deploy.ps1` also seeds sample events when the `Company Events` list is empty,
plus sample stats (`Company Stats`), apps (`Featured Apps`) and departments
(`Departments`) when those lists are empty. `Client Logos` is left empty so staff
can upload the real client logos.

## Notes

- `Add-PnPPageWebPart` uses the web part manifest id
  `816780be-0219-4cf1-826d-3cb7a307de01` (see
  `spfx/src/webparts/companyIntranet/CompanyIntranetWebPart.manifest.json`).
- If the page already has the web part, remove the `-AddWebPart` switch or delete the
  existing instance first.
- For app-only authentication, pass `-ClientId` and use a certificate via PnP
  environment settings instead of `-Interactive`.
