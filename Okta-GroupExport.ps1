<#
.SYNOPSIS
  Exports all members of a single Okta group (by name or ID) into a CSV, including any number of custom profile attributes.

.DESCRIPTION
  • Prompts for: Okta Org URL, API token, group (name or 00g… ID), custom attributes, output path.
  • Columns: userId, username, activated, created, lastLogin + any custom attributes.
  • Joins array-typed attributes with a semicolon.
  • Shows a progress bar as it pages through Okta (200 records per call).
  • Requires PowerShell 5.1+ or PowerShell 7+.
  • Handles pagination, stopping when Okta returns no rel="next" link or an empty page.

.PARAMETER Org
  (Optional) The full Okta org URL, e.g. https://acme.okta.com.
  If omitted, the script prompts for it (or uses $Env:OKTA_ORG).

.PARAMETER Token
  (Optional) A minimally read-only API token (with okta.users.read & okta.groups.read).
  If omitted, the script prompts (or uses $Env:OKTA_TOKEN).

.PARAMETER Group
  (Optional) The group name (partial match) or 00g… ID. If omitted, the script will prompt.

.PARAMETER ProfileAttrs
  (Optional) A string array of profile-attribute names (e.g. groupSettings,CostCenter).
  If omitted, the script prompts for them; otherwise those values are used.

.PARAMETER CsvPath
  (Optional) Full path for the CSV. If omitted, the script prompts and defaults to "OktaGroup_<timestamp>.csv" in the working directory.

.EXAMPLE
  # Fully interactive (no args):
  .\OktaGroupExport.ps1

.EXAMPLE
  # Non-interactive (all args provided):
  $env:OKTA_ORG   = "https://acme.okta.com"
  $env:OKTA_TOKEN = "00aBcDeF..."
  .\OktaGroupExport.ps1 `
    -Group "Working Group" `
    -ProfileAttrs "groupSettings","CostCenter" `
    -CsvPath "C:\oktaExports\workgroup.csv"

.NOTES
  • This script is designed to run entirely locally: no external modules or binaries.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $Org,

    [Parameter(Mandatory = $false)]
    [string] $Token,

    [Parameter(Mandatory = $false)]
    [string] $Group,

    [Parameter(Mandatory = $false)]
    [string[]] $ProfileAttrs,

    [Parameter(Mandatory = $false)]
    [string] $CsvPath
)

#region Helper functions

function Read-NonEmpty {
    param(
        [string] $Prompt,
        [switch] $Secure
    )
    do {
        if ($Secure) {
            $v = Read-Host $Prompt -AsSecureString |
                 ConvertFrom-SecureString -AsPlainText
        }
        else {
            $v = Read-Host $Prompt
        }
        if (-not $v) {
            Write-Host "  ⇒ value required." -ForegroundColor Red
        }
    } until ($v)
    return $v
}

function Invoke-OktaRest {
    param(
        [string] $Uri,
        [string] $Token
    )

    $headers = @{ Authorization = "SSWS $Token" }
    try {
        $result = Invoke-RestMethod `
                   -Method Get `
                   -Uri $Uri `
                   -Headers $headers `
                   -ErrorAction Stop `
                   -ResponseHeadersVariable respHdrs
        # Preserve the Link header for pagination
        $script:LastLinkHeader = $respHdrs['Link']
        return $result
    }
    catch {
        throw "Request failed: $($_.Exception.Message)"
    }
}

function Get-NextLink {
    param(
        [string] $LinkHeader
    )
    if (-not $LinkHeader) { return $null }

    # Match only the rel="next" URL segment
    if ($LinkHeader -match '<([^>]+)>;\s*rel="next"') {
        return $matches[1]
    }
    return $null
}

#endregion

trap {
    Write-Host "`n⚠️  Script aborted: $($_.Exception.Message)" -ForegroundColor Red
    break
}

Write-Host "`n=== Okta Group Exporter ===`n" -ForegroundColor Cyan

#region 1. Resolve Org & Token

if (-not $Org) {
    if ($Env:OKTA_ORG) {
        $Org = $Env:OKTA_ORG
    }
    else {
        $Org = Read-NonEmpty "Okta Org URL (https://org.okta.com):"
    }
}
# Remove any trailing slash
$Org = $Org.TrimEnd('/')

if (-not $Token) {
    if ($Env:OKTA_TOKEN) {
        $Token = $Env:OKTA_TOKEN
    }
    else {
        $Token = Read-NonEmpty "API token (will not echo):" -Secure
    }
}

#endregion

#region 2. Resolve Group ID & Name

if (-not $Group) {
    $Group = Read-NonEmpty "Group name *or* groupID to export:"
}

try {
    if ($Group -notmatch '^00g') {
        $escaped   = [uri]::EscapeDataString($Group)
        $searchUri = "$Org/api/v1/groups?q=$escaped"
        $match     = Invoke-OktaRest $searchUri $Token | Select-Object -First 1

        if (-not $match) {
            throw "No group found matching '$Group'."
        }

        $groupId   = $match.id
        $groupName = $match.profile.name
    }
    else {
        $groupId   = $Group
        $groupName = $Group
    }

    Write-Host "`n✔ Using group: $groupName ($groupId)" -ForegroundColor Green
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

#endregion

#region 3. Prompt for Custom Attributes (if none provided)

if (-not $ProfileAttrs) {
    $attrsInput = Read-Host "Custom profile attributes (comma-sep, blank for none):"
    if ($attrsInput) {
        $ProfileAttrs = $attrsInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    else {
        # If blank, leave $ProfileAttrs as $null or empty array
        $ProfileAttrs = @()
    }
}

#endregion

#region 4. Build Output Path

if (-not $CsvPath) {
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $default = "OktaGroup_${ts}.csv"
    $input   = Read-Host "Save CSV as (default $default)"
    if ($input) {
        $CsvPath = $input
    }
    else {
        $CsvPath = $default
    }
}

#endregion

#region 5. Fetch Members & Stream to CSV

# 5.1 Prepare and write header row
$baseCols   = @('userId','username','activated','created','lastLogin')
$headerCols = $baseCols + $ProfileAttrs
$headerLine = $headerCols -join ','

# Write or overwrite the CSV file with the header
$headerLine | Out-File -FilePath $CsvPath -Encoding UTF8

# 5.2 Initialize paging
$count   = 0
$pageUri = "$Org/api/v1/groups/$groupId/users?limit=200"

Write-Progress -Activity "Fetching users" -Status "0 users so far…" -PercentComplete 0

while ($pageUri) {
    try {
        $page = Invoke-OktaRest $pageUri $Token
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    # If Okta ever returns an empty array, break out
    if ($null -eq $page -or $page.Count -eq 0) {
        break
    }

    # 5.3 Build a PSCustomObject array for this page
    $objects = foreach ($u in $page) {
        $obj = [ordered]@{
            userId    = $u.id
            username  = $u.profile.login
            activated = $u.activated
            created   = $u.created
            lastLogin = $u.lastLogin
        }
        foreach ($a in $ProfileAttrs) {
            $val = $u.profile.$a
            if    ($val -is [array]) { $val = $val -join ';' }
            elseif (-not $val)       { $val = '' }
            $obj[$a] = $val
        }
        [pscustomobject]$obj
    }

    # 5.4 Append these rows to the CSV (no header on subsequent calls)
    $objects | Export-Csv -NoTypeInformation -Append -Path $CsvPath -Encoding UTF8

    # 5.5 Update progress
    $count += $page.Count
    Write-Progress -Activity "Fetching users" `
                   -Status  "$count users so far…" `
                   -PercentComplete 0

    # 5.6 Get next link (if any)
    $pageUri = Get-NextLink $script:LastLinkHeader
}
Write-Progress -Activity "Fetching users" -Completed

# If no rows were written beyond the header, warn and exit.
if ($count -eq 0) {
    Write-Host "No users found in the group." -ForegroundColor Yellow
    exit 0
}

#endregion

#region 6. Completion Message

try {
    Write-Host "`n✅  CSV written to '$CsvPath'  ($count users)" -ForegroundColor Green
}
catch {
    Write-Host "`n❌  Could not write CSV: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

#endregion