<#
.SYNOPSIS
    Provisions the Site Migration Tracker lists, columns, indexes, views and configuration
    on any SharePoint Online site, driven by a JSON schema file.

.DESCRIPTION
    - Reusable: point it at any site and pass a department code / business group.
    - Idempotent: safe to run again. Existing lists and columns are kept; missing ones are added,
      indexes and choice values are brought in line with the schema.
    - Nothing is ever deleted.

.PARAMETER SiteUrl
    Target site, e.g. https://tenant.sharepoint.com/sites/SPDM-CA-Tracking-DEV

.PARAMETER ClientId
    Entra ID app (client) ID used by PnP PowerShell to sign in. If omitted, the
    ENTRAID_APP_ID / ENTRAID_CLIENT_ID environment variable is used.

.PARAMETER SchemaPath
    Path to the schema JSON. Defaults to site-migration-tracker.schema.json next to this script.

.PARAMETER DeptCode
    Department code used as the reference number prefix, e.g. CA.

.PARAMETER BusinessGroup
    Business group value used to filter the source workbook, e.g. "SLF Canada".

.PARAMETER ListPrefix
    Optional prefix for list URLs and titles, used when more than one department shares a site.

.PARAMETER SeedConfig
    Adds the default Configuration rows (only rows whose key does not already exist).

.EXAMPLE
    .\Deploy-SiteMigrationTracker.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/SPDM-CA-Tracking-DEV" `
        -ClientId "00000000-0000-0000-0000-000000000000" -DeptCode CA -BusinessGroup "SLF Canada" `
        -EnterpriseTeamEmail "enterprise.team@company.com" -SharedMailbox "migration@company.com" `
        -SupportTeamEmail "sp-canada-support@company.com" -SeedConfig

.NOTES
    Requires PowerShell 7.4+ and PnP.PowerShell 2.x or 3.x.
    The signed-in account needs Owner (Full Control) on the target site.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $SiteUrl,
    [string] $ClientId = ($env:ENTRAID_APP_ID ?? $env:ENTRAID_CLIENT_ID),
    [string] $SchemaPath = (Join-Path $PSScriptRoot 'site-migration-tracker.schema.json'),
    [string] $DeptCode = 'CA',
    [string] $BusinessGroup = 'SLF Canada',
    [string] $EnterpriseTeamEmail = '',
    [string] $SharedMailbox = '',
    [string] $SupportTeamEmail = '',
    [string] $ListPrefix = '',
    [switch] $SeedConfig
)

$ErrorActionPreference = 'Stop'
$script:Summary = [ordered]@{ ListsCreated = 0; FieldsCreated = 0; FieldsUpdated = 0; ViewsCreated = 0; ConfigRowsAdded = 0; Warnings = 0 }

# ----------------------------------------------------------------------------- helpers
function Write-Step($msg)  { Write-Host "  . $msg" -ForegroundColor Gray }
function Write-Done($msg)  { Write-Host "  + $msg" -ForegroundColor Green }
function Write-Skip($msg)  { Write-Host "  - $msg" -ForegroundColor DarkGray }
function Write-Warn2($msg) { Write-Host "  ! $msg" -ForegroundColor Yellow; $script:Summary.Warnings++ }

function ConvertTo-XmlText([string]$s) { [System.Security.SecurityElement]::Escape($s) }

function Get-Prop($obj, [string]$name, $default = $null) {
    if ($obj.PSObject.Properties.Name -contains $name -and $null -ne $obj.$name) { return $obj.$name }
    return $default
}

function New-FieldXml($f) {
    $name    = $f.name
    $display = ConvertTo-XmlText $f.display
    $type    = $f.type
    $req     = if (Get-Prop $f 'required' $false) { 'TRUE' } else { 'FALSE' }
    $desc    = ConvertTo-XmlText (Get-Prop $f 'description' '')
    $common  = "ID='{$([guid]::NewGuid())}' Name='$name' StaticName='$name' DisplayName='$display' Required='$req' Description='$desc'"
    $default = Get-Prop $f 'default'
    $defXml  = if ($null -ne $default) { "<Default>$(ConvertTo-XmlText $default)</Default>" } else { '' }

    switch ($type) {
        'Text'     { "<Field Type='Text' $common MaxLength='$(Get-Prop $f 'maxLength' 255)'>$defXml</Field>" }
        'Note'     {
            $append = if (Get-Prop $f 'appendOnly' $false) { "AppendOnly='TRUE'" } else { '' }
            "<Field Type='Note' $common NumLines='6' RichText='FALSE' $append />"
        }
        'Number'   { "<Field Type='Number' $common Decimals='$(Get-Prop $f 'decimals' 0)'>$defXml</Field>" }
        'DateTime' { "<Field Type='DateTime' $common Format='$(Get-Prop $f 'format' 'DateOnly')' />" }
        'Boolean'  { "<Field Type='Boolean' $common>$defXml</Field>" }
        'User'     { "<Field Type='User' $common UserSelectionMode='PeopleOnly' Mult='FALSE' />" }
        'URL'      { "<Field Type='URL' $common Format='Hyperlink' />" }
        'Choice'   {
            $choices = ($f.choices | ForEach-Object { "<CHOICE>$(ConvertTo-XmlText $_)</CHOICE>" }) -join ''
            "<Field Type='Choice' $common Format='Dropdown' FillInChoice='FALSE'><CHOICES>$choices</CHOICES>$defXml</Field>"
        }
        default    { throw "Unsupported field type '$type' for field '$name'." }
    }
}

# ----------------------------------------------------------------------------- lists
function Confirm-List($l) {
    $listUrl   = "Lists/$ListPrefix$($l.name)"
    $listTitle = if ($ListPrefix) { "$ListPrefix $($l.title)" } else { $l.title }

    $list = Get-PnPList -Identity $listUrl -ErrorAction SilentlyContinue
    if (-not $list) {
        # Create with a space-free title so the URL is clean, then set the friendly title.
        $list = New-PnPList -Title "$ListPrefix$($l.name)" -Url $listUrl -Template GenericList -OnQuickLaunch:([bool](Get-Prop $l 'onQuickLaunch' $false))
        $script:Summary.ListsCreated++
        Write-Done "List created: $listUrl"
    } else {
        Write-Skip "List exists: $listUrl"
    }

    Set-PnPList -Identity $list -Title $listTitle -Description (Get-Prop $l 'description' '') -EnableVersioning $true -MajorVersions 50 | Out-Null

    $titleDisplay = Get-Prop $l 'titleDisplayName'
    if ($titleDisplay) { Set-PnPField -List $list -Identity 'Title' -Values @{ Title = $titleDisplay } | Out-Null }

    return (Get-PnPList -Identity $listUrl)
}

# ----------------------------------------------------------------------------- fields
function Confirm-Field($list, $f) {
    $field = Get-PnPField -List $list -Identity $f.name -ErrorAction SilentlyContinue
    if (-not $field) {
        Add-PnPFieldFromXml -List $list -FieldXml (New-FieldXml $f) | Out-Null
        $script:Summary.FieldsCreated++
        Write-Done "Field created: $($f.name) ($($f.type))"
        $field = Get-PnPField -List $list -Identity $f.name
    }

    # Read current settings from the field's schema XML (works for every field type and PnP version).
    [xml]$fx       = $field.SchemaXml
    $isIndexed     = $fx.Field.Indexed -eq 'TRUE'
    $isUnique      = $fx.Field.EnforceUniqueValues -eq 'TRUE'

    # Bring indexes in line with the schema (unique requires indexed first).
    $wantIndexed = [bool](Get-Prop $f 'indexed' $false) -or [bool](Get-Prop $f 'unique' $false)
    $wantUnique  = [bool](Get-Prop $f 'unique' $false)

    if ($wantIndexed -and -not $isIndexed) {
        Set-PnPField -List $list -Identity $f.name -Values @{ Indexed = $true } | Out-Null
        $script:Summary.FieldsUpdated++
        Write-Done "Indexed: $($f.name)"
    }
    if ($wantUnique -and -not $isUnique) {
        Set-PnPField -List $list -Identity $f.name -Values @{ EnforceUniqueValues = $true } | Out-Null
        $script:Summary.FieldsUpdated++
        Write-Done "Unique values enforced: $($f.name)"
    }

    # Add any choice values that are in the schema but missing on the column (never removes values).
    if ($f.type -eq 'Choice') {
        $field   = Get-PnPField -List $list -Identity $f.name
        [xml]$cx = $field.SchemaXml
        $current = @($cx.Field.CHOICES.CHOICE | Where-Object { $_ })
        $missing = @($f.choices | Where-Object { $current -notcontains $_ })
        if ($missing.Count -gt 0) {
            try {
                Set-PnPField -List $list -Identity $f.name -Values @{ Choices = [string[]]($current + $missing) } | Out-Null
                $script:Summary.FieldsUpdated++
                Write-Done "Choices added to $($f.name): $($missing -join ', ')"
            } catch {
                Write-Warn2 "Could not update choices on $($f.name): $($_.Exception.Message)"
            }
        }
    }
}

function Set-DefaultViewFields($list, $l) {
    $fields = @('LinkTitle') + @($l.fields | Where-Object { Get-Prop $_ 'inDefaultView' $false } | ForEach-Object { $_.name })
    $defaultView = Get-PnPView -List $list | Where-Object { $_.DefaultView }
    if ($defaultView) {
        Set-PnPView -List $list -Identity $defaultView.Id -Fields $fields | Out-Null
        Write-Step "Default view columns set ($($fields.Count))"
    }
}

# ----------------------------------------------------------------------------- views
function Confirm-View($list, $v) {
    $existing = Get-PnPView -List $list -Identity $v.title -ErrorAction SilentlyContinue
    if ($existing) { Write-Skip "View exists: $($v.title)"; return }
    Add-PnPView -List $list -Title $v.title -Fields $v.fields -Query $v.query -RowLimit (Get-Prop $v 'rowLimit' 100) -Paged | Out-Null
    $script:Summary.ViewsCreated++
    Write-Done "View created: $($v.title)"
}

# ----------------------------------------------------------------------------- config seed
function Add-ConfigSeed($schema) {
    $listUrl = "Lists/$($ListPrefix)Configuration"
    $tokens = @{
        '{DeptCode}'            = $DeptCode
        '{BusinessGroup}'       = $BusinessGroup
        '{EnterpriseTeamEmail}' = $EnterpriseTeamEmail
        '{SharedMailbox}'       = $SharedMailbox
        '{SupportTeamEmail}'    = $SupportTeamEmail
    }
    foreach ($row in $schema.configSeed) {
        $key = $row.key
        $caml = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$(ConvertTo-XmlText $key)</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>"
        $exists = Get-PnPListItem -List $listUrl -Query $caml
        if ($exists) { Write-Skip "Config exists: $key"; continue }

        $value = [string]$row.value
        foreach ($t in $tokens.Keys) { $value = $value.Replace($t, $tokens[$t]) }
        if ($value -match '^\{.+\}$' -or [string]::IsNullOrWhiteSpace($value)) {
            Write-Warn2 "Config '$key' has no value yet - fill it in the Configuration list."
        }

        Add-PnPListItem -List $listUrl -Values @{
            Title             = $key
            ConfigValue       = $value
            Category          = $row.category
            ConfigDescription = $row.description
            IsActive          = $true
        } | Out-Null
        $script:Summary.ConfigRowsAdded++
        Write-Done "Config added: $key"
    }
}

# ============================================================================= main
Write-Host "`nSite Migration Tracker - provisioning" -ForegroundColor Cyan
Write-Host "Site:   $SiteUrl"
Write-Host "Schema: $SchemaPath"
Write-Host "Dept:   $DeptCode ($BusinessGroup)$(if ($ListPrefix) { "  Prefix: $ListPrefix" })`n"

if (-not (Test-Path $SchemaPath)) { throw "Schema file not found: $SchemaPath" }
$schema = Get-Content -Path $SchemaPath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Host "Schema version $($schema.schemaVersion): $($schema.lists.Count) lists`n"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw "No ClientId. Pass -ClientId or set the ENTRAID_APP_ID environment variable."
}

$logPath = Join-Path $PSScriptRoot ("provision-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logPath | Out-Null

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    Write-Done "Connected as $((Get-PnPWeb -Includes CurrentUser).CurrentUser.Email)"

    foreach ($l in $schema.lists) {
        Write-Host "`n[$($l.title)]" -ForegroundColor Cyan
        $list = Confirm-List $l
        foreach ($f in $l.fields) { Confirm-Field $list $f }
        Set-DefaultViewFields $list $l
        foreach ($v in @($l.views)) { if ($v) { Confirm-View $list $v } }
    }

    if ($SeedConfig) {
        Write-Host "`n[Configuration seed]" -ForegroundColor Cyan
        Add-ConfigSeed $schema
    }

    Write-Host "`nDone." -ForegroundColor Cyan
    $script:Summary.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-16} {1}" -f $_.Key, $_.Value) }
}
catch {
    Write-Host "`nFAILED: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
finally {
    Stop-Transcript | Out-Null
    Write-Host "`nLog: $logPath"
}
