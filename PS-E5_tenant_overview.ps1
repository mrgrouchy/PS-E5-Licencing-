# Requires: Microsoft.Graph, ActiveDirectory, & ExchangeOnlineManagement Modules
Import-Module ExchangeOnlineManagement

Write-Host "🔄 Building E5 License Landscape Report..." -ForegroundColor Cyan

# Target SKUs ONLY (ENTERPRISEPREMIUM + SPE_E5)
$targetSkuMap = @{
    "ENTERPRISEPREMIUM" = ""
    "SPE_E5"            = ""
}

Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All"
$skus = Get-MgSubscribedSku

# Map target SKUs
$targetSkuIds = @()
$skuIdToName  = @{}

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
Write-Host "📥 Fetching data..." -ForegroundColor Cyan

# All Entra ID users with sign-in activity
$allUsers = Get-MgUser -All -Property `
    "DisplayName,UserPrincipalName,Mail,Country,CreatedDateTime,EmployeeType,UserType,AccountEnabled,SignInActivity,AssignedLicenses"

# Exchange mailboxes for classification (WITH RecipientTypeDetails in report)
Write-Host "  → Exchange mailboxes..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

$exchangeMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails,PrimarySmtpAddress |
    Select-Object DisplayName, UserPrincipalName, RecipientTypeDetails, PrimarySmtpAddress

# Build lookup: key = UPN or PrimarySMTP; value = mailbox object
$mailboxMap = @{}
foreach ($mbx in $exchangeMailboxes) {
    $key1 = $mbx.UserPrincipalName
    $key2 = $mbx.PrimarySmtpAddress

    if ($key1) { $mailboxMap[$key1.ToLower()] = $mbx }
    if ($key2) { $mailboxMap[$key2.ToLower()] = $mbx }
}

Disconnect-ExchangeOnline -Confirm:$false

$totalUsers = $allUsers.Count
Write-Host "✅ Retrieved: $totalUsers users and $($exchangeMailboxes.Count) mailboxes" -ForegroundColor Green

# Build report
Write-Host "🔎 Building E5 license report..." -ForegroundColor Cyan
$report      = @()
$userCounter = 0

foreach ($user in $allUsers) {
    $percent = [math]::Round(($userCounter / $totalUsers) * 100)
    Write-Progress -Activity "Processing Users" -Status "$($userCounter+1)/$totalUsers" -PercentComplete $percent

    # TARGET E5 SKUs ONLY
    $targetLicenses = foreach ($lic in $user.AssignedLicenses) {
        if ($targetSkuIds -contains $lic.SkuId) { $targetSkuMap[$skuIdToName[$lic.SkuId]] }
    }

    $targetSkuStr = if ($targetLicenses) { ($targetLicenses | Select-Object -Unique) -join ", " } else { "❌ No E5 SKUs" }
    $hasE5        = [bool]$targetLicenses

    # Mailbox classification via Exchange RecipientTypeDetails
    $mailboxType        = "None / No mailbox"
    $recipientTypeRaw   = "N/A"

    $keys = @($user.UserPrincipalName, $user.Mail) |
            Where-Object { $_ -and $_.Trim() -ne "" } |
            ForEach-Object { $_.ToLower() }

    foreach ($key in $keys) {
        if ($mailboxMap.ContainsKey($key)) {
            $mbx = $mailboxMap[$key]
            $recipientTypeRaw = $mbx.RecipientTypeDetails

            $mailboxType = switch ($mbx.RecipientTypeDetails) { # most common values[web:55][web:57]
                "SharedMailbox"     { "🔸 Shared Mailbox" }
                "RoomMailbox"       { "🏢 Room Mailbox" }
                "EquipmentMailbox"  { "⚙️ Equipment Mailbox" }
                "DiscoveryMailbox"  { "🔍 Discovery Mailbox" }
                default             { "👤 User Mailbox" }
            }
            break
        }
    }

    # Entra sign-in only
    $entraLast = $user.SignInActivity?.LastSignInDateTime
    $entraStr  = if ($entraLast) {
        [datetime]$entraLast | Get-Date -Format "yyyy-MM-dd HH:mm"
    } else { "Never" }

    $daysEntra = if ($entraLast) {
        [math]::Round((Get-Date - [datetime]$entraLast).TotalDays, 1)
    } else { "N/A" }

    # Primary sign-in string
    $primarySignIn = if ($entraStr -ne "Never") {
        "☁️ Entra: $entraStr (${daysEntra}days)"
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
        ObjectType             = "User"
        MailboxType            = $mailboxType
        RecipientTypeDetails   = $recipientTypeRaw      # RAW Exchange RecipientTypeDetails
        DisplayName            = $user.DisplayName
        UPN                    = $user.UserPrincipalName
        Email                  = $user.Mail ?? ""
        AccountStatus          = $accountStatus
        LicenseStatus          = $licenseStatus
        Target_E5_SKUs         = $targetSkuStr
        LastSignIn             = $primarySignIn
        Entra_LastSignIn       = $entraStr
        DaysSince_Entra        = $daysEntra
        AD_LastSignIn          = "N/A"                  # AD removed, keep column for now
        CreatedDate            = $createdDate
        Country                = $user.Country ?? ""
        EmployeeType           = $user.EmployeeType ?? "Unknown"
        UserType               = $user.UserType ?? "Member"
    }

    $userCounter++
}

Write-Progress -Activity "Complete" -Completed

# FINAL STATS
$e5Total    = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" }).Count
$disabledE5 = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.AccountStatus -eq "❌ Disabled" }).Count
$e5Shared   = ($report | Where-Object { $_.LicenseStatus -eq "✅ E5 Licensed" -and $_.RecipientTypeDetails -eq "SharedMailbox" }).Count
$inactiveE5_90d = ($report | Where-Object {
    $_.LicenseStatus -eq "✅ E5 Licensed" -and
    [double]($_.DaysSince_Entra -replace "N/A", "0") -gt 90
}).Count

Write-Host "`n📊 E5 LICENSE LANDSCAPE SUMMARY:" -ForegroundColor Cyan
Write-Host "   Total E5 Users: $e5Total" -ForegroundColor Green
Write-Host "   ❌ Disabled E5: $disabledE5" -ForegroundColor Red
Write-Host "   🔸 E5 Shared MB (by RecipientTypeDetails): $e5Shared" -ForegroundColor Magenta
Write-Host "   😴 E5 Inactive 90+ days: $inactiveE5_90d" -ForegroundColor Yellow

# EXPORT
$timestamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$exportPath = "E5_License_Report_$timestamp.csv"

$report |
    Sort-Object @{Expression='LicenseStatus'; Descending=$true}, AccountStatus, MailboxType, DisplayName |
    Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ FULL REPORT SAVED: $exportPath" -ForegroundColor Green
Write-Host "   Columns include: RecipientTypeDetails, MailboxType, Target_E5_SKUs, LastSignIn" -ForegroundColor Cyan
