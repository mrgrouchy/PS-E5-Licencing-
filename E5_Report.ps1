# Requires: Microsoft.Graph & ActiveDirectory Modules
Import-Module ActiveDirectory

Write-Host "🔄 Fetching license SKUs from Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All"
$skus = Get-MgSubscribedSku

# Target SKUs
$targetSkuMap = @{
    "ENTERPRISEPREMIUM" = ""
    "SPE_E5" = ""
}
foreach ($sku in $skus) {
    if ($targetSkuMap.ContainsKey($sku.SkuPartNumber)) {
        $targetSkuMap[$sku.SkuPartNumber] = $sku.SkuId
    }
}
$skuIdToName = @{}
foreach ($sku in $skus) {
    $skuIdToName[$sku.SkuId] = $sku.SkuPartNumber
}
$targetSkuIds = $targetSkuMap.Values | Where-Object { $_ -ne "" }

if ($targetSkuIds.Count -eq 0) {
    Write-Host "❌ No matching E5 SKUs found." -ForegroundColor Red
    return
}

Write-Host "✅ Target SKUs found:`n$($targetSkuMap.GetEnumerator() | ForEach-Object { \" - $($_.Key): $($_.Value)\" })" -ForegroundColor Green

Write-Host "📥 Retrieving all Entra ID users..." -ForegroundColor Cyan
$allUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,AssignedLicenses,Country,CreatedDateTime,EmployeeType"

Write-Host "🔎 Filtering Entra ID users with target E5 licenses..." -ForegroundColor Cyan
$licensedUsers = $allUsers | Where-Object {
    $_.AssignedLicenses.SkuId | Where-Object { $targetSkuIds -contains $_ }
}

$uniqueUsers = $licensedUsers | Sort-Object UserPrincipalName -Unique
$totalUsers = $uniqueUsers.Count
Write-Host "✅ Total unique licensed users: $totalUsers" -ForegroundColor Green

# Get on-prem AD users with lastLogonTimestamp
Write-Host "🔄 Getting on-prem AD lastLogonTimestamp..." -ForegroundColor Cyan
$adUsers = Get-ADUser -Filter * -Properties lastLogonTimestamp,UserPrincipalName | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = $_.UserPrincipalName
        LastLogon         = if ($_.lastLogonTimestamp) { [DateTime]::FromFileTime($_.lastLogonTimestamp) } else { $null }
    }
}
$adUserMap = @{}
foreach ($entry in $adUsers) {
    if (-not $adUserMap.ContainsKey($entry.UserPrincipalName)) {
        $adUserMap[$entry.UserPrincipalName] = $entry.LastLogon
    }
}

# Define allowed employee types
$allowedEmployeeTypes = @("Employee", "FTE", "Permanent")

# Track users with unexpected employee types
$invalidEmployeeTypes = @()

$report = @()
$counter = 0
foreach ($user in $uniqueUsers) {
    $percent = [math]::Round((($counter / $totalUsers) * 100), 0)
    Write-Progress -Activity "Generating E5 License User Report (AD-based)" `
                   -Status "Processing $($user.DisplayName) ($counter of $totalUsers)" `
                   -PercentComplete $percent

    # On-prem AD lastLogonTimestamp
    $adLogonDate = $null
    $adLogonDisplay = "Not Found"
    $daysSinceSignIn = "N/A"

    if ($adUserMap.ContainsKey($user.UserPrincipalName)) {
        $adLogonDate = $adUserMap[$user.UserPrincipalName]
        if ($adLogonDate) {
            $adLogonDisplay = $adLogonDate.ToString("yyyy-MM-dd HH:mm:ss")
            $daysSinceSignIn = [math]::Round(((Get-Date) - $adLogonDate).TotalDays, 0)
        }
    }

    # Get all assigned SKU names
    $allUserSkuNames = @()
    foreach ($license in $user.AssignedLicenses) {
        if ($skuIdToName.ContainsKey($license.SkuId)) {
            $allUserSkuNames += $skuIdToName[$license.SkuId]
        }
    }
    $allSkuLabel = ($allUserSkuNames -join ", ")

    # Get only matching E5 SKU names
    $e5SkuNames = $allUserSkuNames | Where-Object { $targetSkuMap.ContainsKey($_) }
    $skuLabel = ($e5SkuNames -join ", ")

    $createdDate = if ($user.CreatedDateTime) { ([datetime]$user.CreatedDateTime).ToString("yyyy-MM-dd") } else { "Unknown" }
    $employeeType = if ($user.EmployeeType) { $user.EmployeeType } else { "Unknown" }

    # Replace EmployeeType with Email if not allowed
    if ($allowedEmployeeTypes -notcontains $employeeType) {
        $invalidEmployeeTypes += $user.UserPrincipalName
        $employeeType = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
    }

    $report += [PSCustomObject]@{
        DisplayName         = $user.DisplayName
        UserPrincipalName   = $user.UserPrincipalName
        Email               = $user.Mail
        Country             = $user.Country
        LicenseSku          = $skuLabel
        AllAssignedSkus     = $allSkuLabel
        LastSignIn          = $adLogonDisplay
        DaysSinceLastSignIn = $daysSinceSignIn
        CreatedDate         = $createdDate
        EmployeeType        = $employeeType
    }
    $counter++
}

Write-Progress -Activity "Generating E5 License User Report (AD-based)" -Completed

$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$exportPath = "AD_E5_Users_Report_$timestamp.csv"
$report | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "✅ Done. Report saved to: $exportPath" -ForegroundColor Green

if ($invalidEmployeeTypes.Count -gt 0) {
    Write-Host "⚠️ Users with invalid employee types:" -ForegroundColor Yellow
    $invalidEmployeeTypes | ForEach-Object { Write-Host " - $_" }
}
