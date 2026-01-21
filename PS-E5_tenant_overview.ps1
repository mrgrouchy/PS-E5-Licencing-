# Requires: Microsoft.Graph, ActiveDirectory, & ExchangeOnlineManagement Modules
Import-Module ActiveDirectory, ExchangeOnlineManagement

Write-Host "🔄 Building E5 License Landscape Report..." -ForegroundColor Cyan

# Target SKUs only (as requested)
$targetSkuMap = @{
    "ENTERPRISEPREMIUM" = ""
    "SPE_E5" = ""
}

Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"
$skus = Get-MgSubscribedSku

# Map target SKUs only
$targetSkuIds = @()
foreach ($sku in $skus) {
    if ($targetSkuMap.ContainsKey($sku.SkuPartNumber)) {
        $targetSkuMap[$sku.SkuPartNumber] = $sku.SkuId
        $targetSkuIds += $sku.SkuId
    }
}

if ($targetSkuIds.Count -eq 0) {
    Write-Host "❌ No target SKUs (ENTERPRISEPREMIUM/SPE_E5) found!" -ForegroundColor Red
    return
}

Write-Host "✅ Target SKUs: $($targetSkuMap.GetEnumerator() | ForEach-Object { '$($_.Key)' })" -ForegroundColor Green

Write-Host "📥 Fetching data..." -ForegroundColor Cyan

# All Entra ID users
$allUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,Country,CreatedDateTime,EmployeeType,UserType,AccountEnabled,SignInActivity,AssignedLicenses"

# Exchange mailboxes
Connect-ExchangeOnline -ShowBanner:$false
$exchangeMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails,PrimarySmtpAddress | 
    Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
$mailboxMap = @{}
foreach ($mbx in $exchangeMailboxes) {
    $key = $mbx.UserPrincipalName ?? $mbx.PrimarySmtpAddress
    if ($key) { $mailboxMap[$key] = $mbx.RecipientTypeDetails }
}
Disconnect-ExchangeOnline -Confirm:$false

# Service principals
$servicePrincipals = Get-MgServicePrincipal -All -Property "DisplayName,AppId,AccountEnabled,CreatedDateTime"

$totalUsers = $allUsers.Count
Write-Host "✅ $totalUsers users + $($servicePrincipals.Count) service principals" -ForegroundColor Green

# On-prem AD lastLogon
Write-Host "🔄 AD lastLogon data..." -ForegroundColor Cyan
$adUsers = Get-ADUser -Filter * -Properties lastLogonTimestamp,UserPrincipalName | ForEach-Object {
    [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName; LastLogon = if ($_.lastLogonTimestamp) { [DateTime]::FromFileTime($_.lastLogonTimestamp) } }
}
$adUserMap = @{}
foreach ($entry in $adUsers) {
    if (-not $adUserMap[$entry.UserPrincipalName]) { $adUserMap[$entry.UserPrincipalName] = $entry.LastLogon }
}

# Build report
Write-Host "🔎 Analyzing E5 licenses..." -ForegroundColor Cyan
$report = @()

foreach ($user in $allUsers) {
    # Target E5 licenses only
    $targetLicenses = foreach ($lic in $user.AssignedLicenses) {
        if ($targetSkuIds -contains $lic.SkuId) { $targetSkuMap[[string]$skuIdToName[$lic.SkuId]] }
    }
    $licenseStr = if ($targetLicenses) { ($targetLicenses | Select-Object -Unique) -join ", " } else { "❌ No Target SKUs" }
    $hasTargetSku = [bool]$targetLicenses

    # Mailbox classification
    $mailboxType = "👤 Regular User"
    $keys = @($user.UserPrincipalName, $user.Mail)
    foreach ($key in $keys) {
        if ($key -and $mailboxMap[$key]) {
            $rt = $mailboxMap[$key]
            $mailboxType = switch ($rt) {
                "SharedMailbox" { "🔸 Shared Mailbox" }
                "RoomMailbox" { "🏢 Room Mailbox" }
                "EquipmentMailbox" { "⚙️ Equipment Mailbox" }
                default { "👤 User Mailbox" }
            }
            break
        }
    }

    # Sign-ins
    $entraSignIn = if ($user.SignInActivity?.LastSignInDateTime) { 
        [datetime]$user.SignInActivity.LastSignInDateTime | Get-Date -Format "yyyy-MM-dd HH:mm" 
    } else { "Never" }
    $adSignIn = $adUserMap[$user.UserPrincipalName]
    $adSignInStr = if ($adSignIn) { $adSignIn.ToString("yyyy-MM-dd HH:mm") } else { "Not Found" }
    $daysSinceAD = if ($adSignIn) { [math]::Round((Get-Date - $adSignIn).TotalDays) } else { "N/A" }

    # Status
    $accountStatus = if ($user.AccountEnabled) { "✅ Enabled" } else { "❌ Disabled" }
    $licenseStatus = if ($hasTargetSku) { "✅ E5 Licensed" } else { "❌ No E5" }
    $created = if ($user.CreatedDateTime) { [datetime]$user.CreatedDateTime | Get-Date -Format "yyyy-MM-dd" } else { "Unknown" }

    $report += [PSCustomObject]@{
        ObjectType        = "User"
        MailboxType       = $mailboxType
        DisplayName       = $user.DisplayName
        UPN                = $user.UserPrincipalName
        Email              = $user.Mail
        AccountStatus     = $accountStatus
        LicenseStatus     = $licenseStatus
        TargetSKUs        = $licenseStr
        Entra_LastSignIn  = $entraSignIn
        AD_LastSignIn     = $adSignInStr
        DaysSince_AD      = $daysSinceAD
        Created           = $created
        Country           = $user.Country ?? ""
        EmployeeType      = $user.EmployeeType ?? "Unknown"
    }
}

# Service Principals (no E5 licenses)
foreach ($sp in $servicePrincipals) {
    $report += [PSCustomObject]@{
        ObjectType        = "ServicePrincipal"
        MailboxType       = "N/A"
        DisplayName       = $sp.DisplayName
        UPN                = $sp.AppId
        Email              = ""
        AccountStatus     = if ($sp.AccountEnabled) { "✅ Enabled" } else { "❌ Disabled" }
        LicenseStatus     = "N/A"
        TargetSKUs        = "N/A"
        Entra_LastSignIn  = "N/A"
        AD_LastSignIn     = "N/A"
        DaysSince_AD      = "N/A"
        Created           = if ($sp.CreatedDateTime) { [datetime]$sp.CreatedDateTime | Get-Date -Format "yyyy-MM-dd" } else { "Unknown" }
        Country           = ""
        EmployeeType      = "Service Account"
    }
}

# Stats & Export
$e5Users = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" }).Count
$disabledE5 = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.AccountStatus -eq "❌ Disabled" }).Count
$e5Shared = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.MailboxType -like "*Shared*" }).Count

Write-Host "📊 E5 STATS:" -ForegroundColor Cyan
Write-Host "   E5 Users: $e5Users" -ForegroundColor Green
Write-Host "   Disabled E5: $disabledE5 ⚠️" -ForegroundColor Yellow
Write-Host "   E5 Shared MB: $e5Shared ⚠️" -ForegroundColor Yellow

$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$exportPath = "E5_License_Landscape_$timestamp.csv"
$report | Sort-Object LicenseStatus desc, AccountStatus, MailboxType, DisplayName | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "✅ E5 REPORT SAVED: $exportPath" -ForegroundColor Green
Write-Host "   Target SKUs: ENTERPRISEPREMIUM + SPE_E5 only" -ForegroundColor Cyan
