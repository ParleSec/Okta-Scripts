<#
.SYNOPSIS
  Export Okta group members to CSV with interactive attribute selection.

.DESCRIPTION
  Exports Okta group members with:
  - Interactive attribute selection (checkbox-style)
  - Automatic attribute discovery
  - Custom attribute support
  - Column ordering matching Okta admin console

.PARAMETER Org
  Okta org (e.g., "acme" or "https://acme.okta.com")

.PARAMETER Token  
  API token with user/group read permissions

.PARAMETER Group
  Group name or ID (00g...)

.PARAMETER QuickExport
  Skip selection and use Okta default columns

.PARAMETER Output
  Output file path (defaults to GroupName_timestamp.csv)

.EXAMPLE
  .\OktaGroupExport.ps1
  
.EXAMPLE
  .\OktaGroupExport.ps1 -Group "Sales" -QuickExport
#>

[CmdletBinding()]
param(
    [string]$Org = $env:OKTA_ORG,
    [string]$Token = $env:OKTA_TOKEN,
    [string]$Group,
    [switch]$QuickExport,
    [string]$Output
)

# Column display names
$ColumnNames = @{
    id = "User Id"; status = "Status"; login = "Username"; email = "Primary email"
    firstName = "First name"; lastName = "Last name"; displayName = "Display name"
    secondEmail = "Secondary email"; primaryPhone = "Primary phone"; mobilePhone = "Mobile phone"
    title = "Title"; department = "Department"; manager = "Manager"; employeeNumber = "Employee ID"
    created = "Created date"; activated = "Activated date"; lastLogin = "Last login date"
    lastUpdated = "Last updated"; passwordChanged = "Password changed"; statusChanged = "Status changed date"
}

# Default selections (matching Okta export)
$DefaultSelected = @('id','status','login','firstName','lastName','email')

function Invoke-OktaApi {
    param($Uri)
    $headers = @{ Authorization = "SSWS $Token"; Accept = "application/json" }
    try {
        $response = Invoke-WebRequest -Uri $Uri -Headers $headers -UseBasicParsing -ErrorAction Stop
        $script:NextLink = if ($response.Headers.Link -match '<([^>]+)>;\s*rel="next"') { $matches[1] }
        return ($response.Content | ConvertFrom-Json)
    } catch {
        throw if ($_.Exception.Response.StatusCode -eq 401) { "Invalid API token" } else { $_.Exception.Message }
    }
}

function Show-AttributeMenu {
    param($Available, $Selected)
    
    Clear-Host
    Write-Host "`nSELECT ATTRIBUTES TO EXPORT" -ForegroundColor Cyan
    Write-Host "===========================" -ForegroundColor Cyan
    Write-Host "Use numbers to toggle, 'ALL', 'NONE', or Enter to continue`n"
    
    $menu = @{}
    $i = 1
    
    # Group attributes by type
    $groups = [ordered]@{
        "Basic Info" = @('id','status','login','email','firstName','lastName','displayName')
        "Contact" = @('primaryPhone','mobilePhone','secondEmail')
        "Organization" = @('title','department','manager','employeeNumber')
        "Dates" = @('created','activated','lastLogin','lastUpdated','passwordChanged','statusChanged')
        "Custom" = @()
    }
    
    foreach ($groupName in $groups.Keys) {
        $attrs = $groups[$groupName] | Where-Object { $_ -in $Available -or $groupName -eq "Custom" }
        if ($attrs.Count -eq 0 -and $groupName -ne "Custom") { continue }
        
        Write-Host "`n$groupName`:" -ForegroundColor Yellow
        
        if ($groupName -eq "Custom") {
            # Show other discovered attributes
            $shown = $groups.Values | ForEach-Object { $_ }
            $custom = $Available | Where-Object { $_ -notin $shown }
            foreach ($attr in $custom) {
                $checked = if ($attr -in $Selected) { "[X]" } else { "[ ]" }
                $display = if ($ColumnNames[$attr]) { $ColumnNames[$attr] } else { $attr }
                Write-Host ("  {0} {1,2}. {2}" -f $checked, $i, $display)
                $menu[$i] = $attr
                $i++
            }
        } else {
            foreach ($attr in $attrs) {
                $checked = if ($attr -in $Selected) { "[X]" } else { "[ ]" }
                $display = $ColumnNames[$attr]
                Write-Host ("  {0} {1,2}. {2}" -f $checked, $i, $display)
                $menu[$i] = $attr
                $i++
            }
        }
    }
    
    return $menu
}

# Main
Clear-Host
Write-Host "OKTA GROUP EXPORT TOOL" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan

# Get inputs
if (!$Org) { $Org = Read-Host "`nOkta org (e.g., 'acme' or 'acme.okta.com')" }
$OktaUrl = $Org.Trim()
if ($OktaUrl -notmatch '^https?://') {
    $OktaUrl = if ($OktaUrl -match '\.okta\.com') { "https://$OktaUrl" } else { "https://$OktaUrl.okta.com" }
}

if (!$Token) { 
    $secure = Read-Host "API token" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
}

if (!$Group) { $Group = Read-Host "Group name or ID" }

# Find group
Write-Host "`nSearching..." -ForegroundColor Gray
if ($Group -match '^00g') {
    $groupData = Invoke-OktaApi "$OktaUrl/api/v1/groups/$Group"
    $groupId = $Group
    $groupName = $groupData.profile.name
} else {
    $groups = Invoke-OktaApi "$OktaUrl/api/v1/groups?q=$([uri]::EscapeDataString($Group))&limit=10"
    if ($groups.Count -eq 0) { throw "No groups found" }
    
    if ($groups.Count -eq 1) {
        $groupId = $groups[0].id
        $groupName = $groups[0].profile.name
    } else {
        Write-Host "`nMultiple groups found:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host "  $($i+1). $($groups[$i].profile.name)"
        }
        $choice = Read-Host "Select number"
        $selected = $groups[[int]$choice - 1]
        $groupId = $selected.id
        $groupName = $selected.profile.name
    }
}
Write-Host "✓ Group: $groupName" -ForegroundColor Green

# Discover available attributes
Write-Host "Discovering attributes..." -ForegroundColor Gray
$sample = Invoke-OktaApi "$OktaUrl/api/v1/groups/$groupId/users?limit=10"
$available = @('id','status','created','activated','lastLogin','lastUpdated','passwordChanged','statusChanged')
if ($sample.Count -gt 0) {
    $available += @($sample[0].profile.PSObject.Properties.Name)
    # Check for more profile attributes across all samples
    foreach ($user in $sample | Select-Object -Skip 1) {
        $user.profile.PSObject.Properties.Name | Where-Object { $_ -notin $available } | ForEach-Object { $available += $_ }
    }
}
$available = $available | Select-Object -Unique | Sort-Object

# Select attributes
if ($QuickExport) {
    $selectedAttrs = $DefaultSelected
} else {
    $selected = [System.Collections.ArrayList]@($DefaultSelected)
    $menu = Show-AttributeMenu -Available $available -Selected $selected
    
    while ($true) {
        $input = Read-Host "`nToggle"
        if ([string]::IsNullOrWhiteSpace($input)) { break }
        
        switch ($input.ToUpper()) {
            'ALL' { 
                $selected.Clear()
                $selected.AddRange($menu.Values)
                Write-Host "All selected" -ForegroundColor Green
            }
            'NONE' { 
                $selected.Clear()
                Write-Host "All cleared" -ForegroundColor Yellow
            }
            default {
                $nums = $input -split ',' | ForEach-Object { 
                    if ($_ -match '^\d+$') { [int]$_ }
                    elseif ($_ -match '^(\d+)-(\d+)$') { [int]$matches[1]..[int]$matches[2] }
                } | Where-Object { $_ }
                
                foreach ($num in $nums) {
                    if ($menu.ContainsKey($num)) {
                        if ($menu[$num] -in $selected) { $selected.Remove($menu[$num]) }
                        else { [void]$selected.Add($menu[$num]) }
                    }
                }
            }
        }
        
        $menu = Show-AttributeMenu -Available $available -Selected $selected
    }
    
    # Add custom attributes
    Write-Host "`nAdd custom attributes? (comma-separated, or Enter to skip)" -ForegroundColor Yellow
    Write-Host "Example: costCenter,building,customField1" -ForegroundColor Gray
    $custom = Read-Host "Custom"
    if ($custom) {
        $custom -split ',' | ForEach-Object { 
            $attr = $_.Trim()
            if ($attr -and $attr -notin $selected) { [void]$selected.Add($attr) }
        }
    }
    
    $selectedAttrs = $selected
}

if ($selectedAttrs.Count -eq 0) { throw "No attributes selected" }

# Order attributes logically
$ordered = @()
$order = @('id','firstName','lastName','displayName','login','email','status','title','department','manager',
           'primaryPhone','mobilePhone','employeeNumber','created','activated','lastLogin')
foreach ($attr in $order) {
    if ($attr -in $selectedAttrs) { $ordered += $attr }
}
foreach ($attr in $selectedAttrs) {
    if ($attr -notin $ordered) { $ordered += $attr }
}
$selectedAttrs = $ordered

# Set output file
if (!$Output) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $Output = "$($groupName -replace '[^\w]','_')_$timestamp.csv"
}

# Build headers
$headers = @()
foreach ($attr in $selectedAttrs) {
    $headers += if ($ColumnNames[$attr]) { $ColumnNames[$attr] } else { $attr }
}

# Export users
Write-Host "`nExporting $($selectedAttrs.Count) attributes..." -ForegroundColor Gray
$csv = [System.Text.StringBuilder]::new()
[void]$csv.AppendLine(($headers | ForEach-Object { "`"$_`"" }) -join ',')

$count = 0
$pageUrl = "$OktaUrl/api/v1/groups/$groupId/users?limit=200"
$startTime = Get-Date

while ($pageUrl) {
    $users = Invoke-OktaApi $pageUrl
    if ($users.Count -eq 0) { break }
    
    foreach ($user in $users) {
        $row = @()
        foreach ($attr in $selectedAttrs) {
            $value = if ($attr -in @('id','status','created','activated','lastLogin','lastUpdated','passwordChanged','statusChanged')) {
                $user.$attr
            } else {
                $user.profile.$attr
            }
            
            if ($null -eq $value) { $value = '' }
            elseif ($value -is [datetime]) { $value = $value.ToString('yyyy-MM-dd HH:mm:ss') }
            elseif ($value -is [array]) { $value = $value -join ';' }
            else { $value = $value.ToString() }
            
            if ($value -match '[",\r\n]') { $value = "`"$($value -replace '"','""')`"" }
            $row += $value
        }
        [void]$csv.AppendLine($row -join ',')
        $count++
    }
    
    Write-Progress -Activity "Exporting users" -Status "$count users" -PercentComplete -1
    $pageUrl = $script:NextLink
}

# Save file
[System.IO.File]::WriteAllText($Output, $csv.ToString(), [System.Text.Encoding]::UTF8)
Write-Progress -Activity "Exporting users" -Completed

# Summary
$duration = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
Write-Host "`n✓ EXPORT COMPLETE" -ForegroundColor Green
Write-Host "  File: $Output"
Write-Host "  Users: $count"
Write-Host "  Columns: $($selectedAttrs.Count)"
Write-Host "  Time: ${duration}s"
Write-Host "  Size: $([Math]::Round((Get-Item $Output).Length / 1KB, 1))KB"
