<#
.SYNOPSIS
    Creates the Company Intranet communication site and applies the PnP site template.

.DESCRIPTION
    Requires the PnP.PowerShell module (PowerShell 7+ recommended):
        Install-Module PnP.PowerShell -Scope CurrentUser

    The script is idempotent: it reuses the site if it already exists and
    re-applies the template (lists, fields and sample rows).

.EXAMPLE
    ./deploy.ps1 `
        -TenantUrl "https://contoso.sharepoint.com" `
        -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
        -Owner     "admin@contoso.onmicrosoft.com"

.EXAMPLE
    ./deploy.ps1 `
        -TenantUrl "https://contoso.sharepoint.com" `
        -SiteUrl   "https://contoso.sharepoint.com/sites/cosmos" `
        -Owner     "admin@contoso.onmicrosoft.com" `
        -SppkgPath "../spfx/sharepoint/solution/sml-cosmos-intranet.sppkg" `
        -AddWebPart
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $TenantUrl,

    [Parameter(Mandatory = $true)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [string] $Owner,

    [string] $SiteTitle = "SML Cosmos Intranet",
    [string] $SiteDescription = "SML Cosmos company intranet: news, events, people and tools.",

    [string] $TemplatePath = (Join-Path $PSScriptRoot "site-template.xml"),

    [string] $SppkgPath,

    [string] $ClientId,

    [switch] $AddWebPart,

    [switch] $SkipAppDeployment
)

$ErrorActionPreference = "Stop"

# Web part manifest id from spfx/src/webparts/companyIntranet/CompanyIntranetWebPart.manifest.json
$WebPartId = "816780be-0219-4cf1-826d-3cb7a307de01"

function Connect-ToUrl {
    param([string] $Url)

    $connection = @{
        Url         = $Url
        Interactive = $true
    }

    if ($ClientId) {
        $connection.ClientId = $ClientId
    }

    Connect-PnPOnline @connection
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}

if (-not (Test-Path $TemplatePath)) {
    throw "Template not found: $TemplatePath"
}

Write-Host "Connecting to tenant $TenantUrl..." -ForegroundColor Cyan
Connect-ToUrl -Url $TenantUrl

$existingSite = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue

if (-not $existingSite) {
    Write-Host "Creating communication site $SiteUrl..." -ForegroundColor Cyan
    New-PnPSite -Type CommunicationSite `
        -Title $SiteTitle `
        -Url $SiteUrl `
        -Owner $Owner `
        -Description $SiteDescription `
        -Wait
} else {
    Write-Host "Site already exists, reusing $SiteUrl." -ForegroundColor Yellow
}

Write-Host "Connecting to site $SiteUrl..." -ForegroundColor Cyan
Connect-ToUrl -Url $SiteUrl

Write-Host "Applying site template $TemplatePath..." -ForegroundColor Cyan
Invoke-PnPSiteTemplate -Path $TemplatePath

function Add-SampleEvents {
    $eventsList = Get-PnPList -Identity "Company Events" -ErrorAction SilentlyContinue
    if (-not $eventsList) {
        return
    }

    $existingItems = @(Get-PnPListItem -List "Company Events" -PageSize 1)
    if ($existingItems.Count -gt 0) {
        Write-Host "Company Events already has items, skipping sample events." -ForegroundColor Yellow
        return
    }

    Write-Host "Seeding sample company events..." -ForegroundColor Cyan
    $today = (Get-Date).Date

    $samples = @(
        @{
            Title            = "Monthly Town Hall"
            EventDate        = $today.AddDays(7).AddHours(10)
            EventEndDate     = $today.AddDays(7).AddHours(11)
            EventLocation    = "Head Office Auditorium"
            EventDescription = "Company update from the leadership team, followed by an open Q&A."
        },
        @{
            Title            = "Safety Awareness Workshop"
            EventDate        = $today.AddDays(14).AddHours(9).AddMinutes(30)
            EventEndDate     = $today.AddDays(14).AddHours(12)
            EventLocation    = "Training Room 2"
            EventDescription = "Refresher on site safety, lifting plans and incident reporting."
        },
        @{
            Title            = "SML Cosmos Constellation Day"
            EventDate        = $today.AddDays(30).AddHours(17)
            EventEndDate     = $today.AddDays(30).AddHours(20)
            EventLocation    = "Head Office Grounds"
            EventDescription = "Celebrate our people, projects and milestones together."
        }
    )

    foreach ($sample in $samples) {
        Add-PnPListItem -List "Company Events" -Values $sample | Out-Null
    }
}

Add-SampleEvents

function Add-SampleStats {
    $statsList = Get-PnPList -Identity "Company Stats" -ErrorAction SilentlyContinue
    if (-not $statsList) {
        return
    }

    $existingItems = @(Get-PnPListItem -List "Company Stats" -PageSize 1)
    if ($existingItems.Count -gt 0) {
        Write-Host "Company Stats already has items, skipping sample stats." -ForegroundColor Yellow
        return
    }

    Write-Host "Seeding sample company stats..." -ForegroundColor Cyan

    $samples = @(
        @{
            Title  = "Fleet utilisation"
            Value  = "92%"
            Icon   = "Truck"
            InHero = $true
        },
        @{
            Title  = "Active sites"
            Value  = "14"
            Icon   = "Location"
            InHero = $true
        },
        @{
            Title  = "Days without incident"
            Value  = "213"
            Icon   = "Shield"
            InHero = $true
        },
        @{
            Title  = "Fleet size"
            Value  = "250+"
            Icon   = "Truck"
            InHero = $false
        },
        @{
            Title  = "Years of operation"
            Value  = "38"
            Icon   = "Clock"
            InHero = $false
        },
        @{
            Title  = "People"
            Value  = "1,200+"
            Icon   = "People"
            InHero = $false
        },
        @{
            Title  = "Projects delivered"
            Value  = "5,400"
            Icon   = "Star"
            InHero = $false
        }
    )

    try {
        foreach ($sample in $samples) {
            Add-PnPListItem -List "Company Stats" -Values $sample | Out-Null
        }
    } catch {
        Write-Host "Seeding Company Stats failed, continuing: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Add-SampleFeaturedApps {
    $appsList = Get-PnPList -Identity "Featured Apps" -ErrorAction SilentlyContinue
    if (-not $appsList) {
        return
    }

    $existingItems = @(Get-PnPListItem -List "Featured Apps" -PageSize 1)
    if ($existingItems.Count -gt 0) {
        Write-Host "Featured Apps already has items, skipping sample apps." -ForegroundColor Yellow
        return
    }

    Write-Host "Seeding sample featured apps..." -ForegroundColor Cyan

    $samples = @(
        @{
            Title = "Safety Portal"
            Icon  = "Shield"
            Url   = "/Lists/QuickLinks"
        },
        @{
            Title = "Crane Booking"
            Icon  = "Calendar"
            Url   = "/Lists/QuickLinks"
        },
        @{
            Title = "HR Self-Service"
            Icon  = "People"
            Url   = "/Lists/QuickLinks"
        },
        @{
            Title = "IT Support"
            Icon  = "Phone"
            Url   = "/Lists/QuickLinks"
        }
    )

    try {
        foreach ($sample in $samples) {
            Add-PnPListItem -List "Featured Apps" -Values $sample | Out-Null
        }
    } catch {
        Write-Host "Seeding Featured Apps failed, continuing: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Add-SampleDepartments {
    $departmentsList = Get-PnPList -Identity "Departments" -ErrorAction SilentlyContinue
    if (-not $departmentsList) {
        return
    }

    $existingItems = @(Get-PnPListItem -List "Departments" -PageSize 1)
    if ($existingItems.Count -gt 0) {
        Write-Host "Departments already has items, skipping sample departments." -ForegroundColor Yellow
        return
    }

    Write-Host "Seeding sample departments..." -ForegroundColor Cyan

    $samples = @(
        @{
            Title       = "Operations"
            Description = "Fleet operations, logistics and site dispatch."
            Icon        = "Truck"
            Head        = "ops@smlcosmos.com"
            Url         = "/Lists/QuickLinks"
        },
        @{
            Title       = "Projects & Engineering"
            Description = "Project delivery, engineering and planning."
            Icon        = "Globe"
            Head        = "projects@smlcosmos.com"
            Url         = "/Lists/QuickLinks"
        },
        @{
            Title       = "Safety & Compliance"
            Description = "Safety management, audits and regulatory compliance."
            Icon        = "Shield"
            Head        = "safety@smlcosmos.com"
            Url         = "/Lists/QuickLinks"
        }
    )

    try {
        foreach ($sample in $samples) {
            Add-PnPListItem -List "Departments" -Values $sample | Out-Null
        }
    } catch {
        Write-Host "Seeding Departments failed, continuing: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Add-SampleStats
Add-SampleFeaturedApps
Add-SampleDepartments

Write-Host "Ensuring the home page exists..." -ForegroundColor Cyan
$homePage = Get-PnPPage -Identity "Home" -ErrorAction SilentlyContinue

if (-not $homePage) {
    Add-PnPPage -Name "Home" -LayoutType Home -Title "Home"
}

Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Home.aspx"

if (-not $SkipAppDeployment -and $SppkgPath) {
    if (-not (Test-Path $SppkgPath)) {
        throw "SPFx package not found: $SppkgPath. Run 'npm run package' in the spfx folder first."
    }

    Write-Host "Deploying SPFx package to the tenant app catalog..." -ForegroundColor Cyan
    Connect-ToUrl -Url $TenantUrl

    $app = Add-PnPApp -Path $SppkgPath -Scope Tenant -Overwrite -Publish
    Write-Host "Installing the app on $SiteUrl..." -ForegroundColor Cyan

    Connect-ToUrl -Url $SiteUrl
    Install-PnPApp -Identity $app.Id

    if ($AddWebPart) {
        Write-Host "Adding the Company Intranet web part to Home.aspx..." -ForegroundColor Cyan
        Add-PnPPageWebPart -Page "Home" -ComponentId $WebPartId -Section 1 -Column 1
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Site: $SiteUrl"
Write-Host "Next: open the site, add the 'Company Intranet' web part to Home.aspx and publish the page."
