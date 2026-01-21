# Requires: Microsoft.Graph, ActiveDirectory, & ExchangeOnlineManagement Modules
Import-Module ExchangeOnlineManagement

Write-Host "🔄 Building E5 License Landscape Report..." -ForegroundColor Cyan

# Target SKUs ONLY (ENTERPRISEPREMIUM + SPE_E5)
$targetSkuMap = @{
    "ENTERPRISEPREMIUM" = ""
    "SPE_E5" = ""
}

Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"
$skus = Get-MgSubscribedSku

# Map target SKUs
$targetSkuIds = @()
$skuIdToName = @{}
foreach ($sku in $skus) {
    $skuIdToName[$sku.SkuId] = $sku.SkuPartNumber
    if ($targetSkuMap.ContainsKey($sku.SkuPartNumber)) {
        $targetSkuMap[$sku.SkuPartNumber] = $sku.SkuId
        $targetSkuIds += $sku.SkuId
    }
}

if ($targetSkuIds.Count -eq 0) {
    Write-Host "❌ No target SKUs (ENTERPRISEPREMIUM/SPE_E5) found!" -ForegroundColor Red
    return
}

Write-Host "✅ Target SKUs mapped: $($targetSkuMap.Keys -join ', ')" -ForegroundColor Green

Write-Host "📥 Fetching comprehensive data..." -ForegroundColor Cyan

# All Entra ID users with sign-in activity
$allUsers = Get-MgUser -All -Property "DisplayName,UserPrincipalName,Mail,Country,CreatedDateTime,EmployeeType,UserType,AccountEnabled,SignInActivity,AssignedLicenses"

# Exchange mailboxes for classification
Write-Host "  → Exchange mailboxes..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
$exchangeMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails,PrimarySmtpAddress |
    Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress
$mailboxMap = @{}
foreach ($mbx in $exchangeMailboxes) {
    $key = $mbx.UserPrincipalName ?? $mbx.PrimarySmtpAddress
    if ($key) { $mailboxMap[$key] = $mbx.RecipientTypeDetails }
}
Disconnect-ExchangeOnline -Confirm:$false

# Service principals (non-human app identities)
Write-Host "  → Service principals..." -ForegroundColor Cyan
$servicePrincipals = Get-MgServicePrincipal -All -Property "DisplayName,AppId,AccountEnabled,CreatedDateTime"

$totalUsers = $allUsers.Count
Write-Host "✅ Retrieved: $totalUsers users + $($servicePrincipals.Count) service principals" -ForegroundColor Green

# On-prem AD lastLogon (backup for hybrid)
Write-Host "🔄 On-prem AD lastLogon (backup)..." -ForegroundColor Cyan
$adUsers = Get-ADUser -Filter * -Properties lastLogonTimestamp,UserPrincipalName | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = $_.UserPrincipalName
        LastLogon = if ($_.lastLogonTimestamp) { [DateTime]::FromFileTime($_.lastLogonTimestamp) }
    }
}
$adUserMap = @{}
foreach ($entry in $adUsers) {
    if (-not $adUserMap.ContainsKey($entry.UserPrincipalName)) {
        $adUserMap[$entry.UserPrincipalName] = $entry.LastLogon
    }
}

# Build COMPLETE report
Write-Host "🔎 Building E5 license report (Entra sign-ins prioritized)..." -ForegroundColor Cyan
$report = @()
$userCounter = 0

foreach ($user in $allUsers) {
    $percent = [math]::Round(($userCounter / $totalUsers) * 100)
    Write-Progress -Activity "Processing Users" -Status "$($userCounter+1)/$totalUsers" -PercentComplete $percent

    # TARGET E5 SKUs ONLY
    $targetLicenses = foreach ($lic in $user.AssignedLicenses) {
        if ($targetSkuIds -contains $lic.SkuId) { $targetSkuMap[$skuIdToName[$lic.SkuId]] }
    }
    $targetSkuStr = if ($targetLicenses) { ($targetLicenses | Select-Object -Unique) -join ", " } else { "❌ No E5 SKUs" }
    $hasE5 = [bool]$targetLicenses

    # Mailbox classification
    $mailboxType = "👤 Regular User"
    $keys = @($user.UserPrincipalName, $user.Mail)
    foreach ($key in $keys) {
        if ($key -and $mailboxMap.ContainsKey($key)) {
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

    # ENTRA ID Sign-In (PRIMARY - cloud M365 activity)
    $entraLast = $user.SignInActivity?.LastSignInDateTime
    $entraStr = if ($entraLast) {
        [datetime]$entraLast | Get-Date -Format "yyyy-MM-dd HH:mm"
    } else { "Never" }
    $daysEntra = if ($entraLast) {
        [math]::Round((Get-Date - [datetime]$entraLast).TotalDays, 1)
    } else { "N/A" }

    # AD Backup
    $adLast = $adUserMap[$user.UserPrincipalName]
    $adStr = if ($adLast) { $adLast.ToString("yyyy-MM-dd HH:mm") } else { "N/A" }

    # Combined LastSignIn (Entra > AD)
    $primarySignIn = if ($entraStr -ne "Never") {
        "☁️ Entra: $entraStr (${daysEntra}days)"
    } elseif ($adStr -ne "N/A") {
        "🏠 AD: $adStr"
    } else {
        "🚫 Never"
    }

    # Status flags
    $accountStatus = if ($user.AccountEnabled) { "✅ Enabled" } else { "❌ Disabled" }
    $licenseStatus = if ($hasE5) { "✅ E5 Licensed" } else { "❌ No E5" }

    $createdDate = if ($user.CreatedDateTime) {
        [datetime]$user.CreatedDateTime | Get-Date -Format "yyyy-MM-dd"
    } else { "Unknown" }

    $report += [PSCustomObject]@{
        ObjectType        = "User"
        MailboxType       = $mailboxType
        DisplayName       = $user.DisplayName
        UPN                = $user.UserPrincipalName
        Email              = $user.Mail ?? ""
        AccountStatus     = $accountStatus
        LicenseStatus     = $licenseStatus
        Target_E5_SKUs    = $targetSkuStr
        LastSignIn        = $primarySignIn
        Entra_LastSignIn  = $entraStr
        DaysSince_Entra   = $daysEntra
        AD_LastSignIn     = $adStr
        CreatedDate       = $createdDate
        Country           = $user.Country ?? ""
        EmployeeType      = $user.EmployeeType ?? "Unknown"
        UserType          = $user.UserType ?? "Member"
    }
    $userCounter++
}

# Service Principals
Write-Host "➕ Adding Service Principals..." -ForegroundColor Cyan
foreach ($sp in $servicePrincipals) {
    $report += [PSCustomObject]@{
        ObjectType        = "ServicePrincipal"
        MailboxType       = "N/A"
        DisplayName       = $sp.DisplayName
        UPN                = $sp.AppId
        Email              = ""
        AccountStatus     = if ($sp.AccountEnabled) { "✅ Enabled" } else { "❌ Disabled" }
        LicenseStatus     = "N/A"
        Target_E5_SKUs    = "N/A"
        LastSignIn        = "N/A"
        Entra_LastSignIn  = "N/A"
        DaysSince_Entra   = "N/A"
        AD_LastSignIn     = "N/A"
        CreatedDate       = if ($sp.CreatedDateTime) {
            [datetime]$sp.CreatedDateTime | Get-Date -Format "yyyy-MM-dd"
        } else { "Unknown" }
        Country           = ""
        EmployeeType      = "Service Account"
        UserType          = "App"
    }
}

Write-Progress -Activity "Complete" -Completed

# FINAL STATS
$e5Total = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" }).Count
$disabledE5 = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.AccountStatus -eq "❌ Disabled" }).Count
$e5Shared = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.MailboxType -like "*Shared*" }).Count
$inactiveE5_90d = ($report | Where-Object {
    $_.LicenseStatus -eq "✅ E5 Licensed" -and
    [double]($_.DaysSince_Entra -replace "N/A", "0") -gt 90
}).Count

Write-Host "`n📊 E5 LICENSE LANDSCAPE SUMMARY:" -ForegroundColor Cyan
Write-Host "   Total E5 Users: $e5Total" -ForegroundColor Green
Write-Host "   ❌ Disabled E5: $disabledE5" -ForegroundColor Red
Write-Host "   🔸 E5 Shared MB: $e5Shared" -ForegroundColor Magenta
Write-Host "   😴 E5 Inactive 90+ days: $inactiveE5_90d" -ForegroundColor Yellow

# EXPORT
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$exportPath = "E5_License_Report_$timestamp.csv"
$report | Sort-Object @{Expression='LicenseStatus'; Descending=$true}, AccountStatus, MailboxType, DisplayName |
    Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ FULL REPORT SAVED: $exportPath" -ForegroundColor Green
Write-Host "   Columns: Target_E5_SKUs | AccountStatus | MailboxType | LastSignIn (Entra+AD)" -ForegroundColor Cyan
