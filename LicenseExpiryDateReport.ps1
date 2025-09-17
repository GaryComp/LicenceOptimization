Param(
    [switch]$Trial,
    [switch]$Free,
    [switch]$Purchased,
    [switch]$Expired,
    [switch]$Active,
    [switch]$CreateSession,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# ------------------------- Connect to Microsoft Graph -------------------------
Function Connect-ToGraph {
    if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
        Write-Host "Microsoft Graph PowerShell SDK not found. Installing..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber -Force
    }

    if ($CreateSession) {
        Disconnect-MgGraph -Force
    }

    Write-Host "`n🔌 Connecting to Microsoft Graph..."
    if ($TenantId -and $ClientId -and $CertificateThumbprint) {
        Microsoft.Graph.Authentication\Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    }
    else {
        Microsoft.Graph.Authentication\Connect-MgGraph -Scopes "Directory.Read.All" -NoWelcome
    }
}

# Call the connection helper
Connect-ToGraph

# ------------------------- Output Paths -------------------------
$TimeStamp = Get-Date -Format "yyyy-MMM-dd-ddd__hh-mm_tt"
$ExportCSV = ".\LicenseExpiryReport_$TimeStamp.csv"

# ------------------------- Load Friendly Names -------------------------
$FriendlyNameHash = @{}
if (Test-Path ".\Product_names_and_service_plan_identifiers.csv") {
    Import-Csv ".\Product_names_and_service_plan_identifiers.csv" | ForEach-Object {
        if ($_.SkuPartNumber -and $_.Product_Display_Name) {
            $FriendlyNameHash[$_.SkuPartNumber.Trim()] = $_.Product_Display_Name.Trim()
        }
    }
}
else {
    Write-Host "❌ No friendly name mapping file (Product_names_and_service_plan_identifiers.csv) found in current directory." -ForegroundColor Red
    exit
}

# ------------------------- Determine Filters -------------------------
$ShowAll = -not ($Trial -or $Free -or $Purchased -or $Expired -or $Active)

# ------------------------- Retrieve License Info -------------------------
Write-Host "`n📦 Retrieving subscribed SKUs..." -ForegroundColor Cyan
$Skus = Get-MgSubscribedSku -All

# Lifecycle info (try v1.0, fallback to beta)
$lifecycleUriV1   = "https://graph.microsoft.com/v1.0/directory/subscriptions"
$lifecycleUriBeta = "https://graph.microsoft.com/beta/directory/subscriptions"
$lifecycleInfo = $null

try {
    $lifecycleInfo = (Invoke-MgGraphRequest -Uri $lifecycleUriV1 -Method GET -ErrorAction Stop).value
}
catch {
    Write-Host "⚠️ v1.0 lifecycle endpoint failed, falling back to beta endpoint" -ForegroundColor Yellow
    $lifecycleInfo = (Invoke-MgGraphRequest -Uri $lifecycleUriBeta -Method GET -ErrorAction Stop).value
}

# ------------------------- Process Results -------------------------
$Results = @()
foreach ($Sku in $Skus) {
    $SkuId = $Sku.SkuId
    $SkuPartNumber = $Sku.SkuPartNumber
    $FriendlyName = $FriendlyNameHash[$SkuPartNumber] ?? $SkuPartNumber
    $Consumed = ($Sku.ConsumedUnits | Select-Object -First 1)

    $Lifecycle = $lifecycleInfo | Where-Object { $_.skuId -eq $SkuId }
    if (-not $Lifecycle) { continue }

    $Created    = ($Lifecycle.createdDateTime | Select-Object -First 1)
    $Status     = ($Lifecycle.status | Select-Object -First 1)
    $Total      = ($Lifecycle.totalLicenses | Select-Object -First 1)
    $ExpiryDate = ($Lifecycle.nextLifecycleDateTime | Select-Object -First 1)

    # Ensure numeric subtraction works
    $Remaining = 0
    if ($Total -and $Consumed -is [int]) {
        $Remaining = $Total - $Consumed
    }

    # Subscription type
    if ($SkuPartNumber -like "*Free*" -and -not $ExpiryDate) {
        $Type = "Free"
    }
    elseif (-not $ExpiryDate) {
        $Type = "Trial"
    }
    else {
        $Type = "Purchased"
    }

    # Subscribed date
    if ($Created) {
        try {
            $SubscribedDate = [datetime]$Created
            $SubscribedAgo = (New-TimeSpan -Start $SubscribedDate -End (Get-Date)).Days
            $SubscribedFriendly = if ($SubscribedAgo -eq 0) { "Today" } else { "$SubscribedAgo days ago" }
            $SubscribedString = "$SubscribedDate ($SubscribedFriendly)"
        }
        catch {
            $SubscribedString = "Invalid date"
        }
    }
    else {
        $SubscribedString = "Unknown"
    }

    # Expiry date
    if ($ExpiryDate) {
        try {
            $ExpiryDateTime = [datetime]$ExpiryDate
            $DaysRemaining = (New-TimeSpan -Start (Get-Date) -End $ExpiryDateTime).Days
            switch ($Status) {
                "Enabled"   { $ExpiryNote = "Will expire in $DaysRemaining days" }
                "Warning"   { $ExpiryNote = "Expired. Will suspend in $DaysRemaining days" }
                "Suspended" { $ExpiryNote = "Expired. Will delete in $DaysRemaining days" }
                "LockedOut" { $ExpiryNote = "Subscription is locked. Contact Microsoft." }
                default     { $ExpiryNote = "Unknown status" }
            }
        }
        catch {
            $ExpiryNote = "Invalid expiry date"
            $ExpiryDateTime = "-"
        }
    }
    else {
        $ExpiryNote = "Never Expires"
        $ExpiryDateTime = "-"
    }

    # Apply filters
    $Include = $false
    if ($ShowAll) { $Include = $true }
    if ($Trial -and $Type -eq "Trial") { $Include = $true }
    if ($Free -and $Type -eq "Free") { $Include = $true }
    if ($Purchased -and $Type -eq "Purchased") { $Include = $true }
    if ($Expired -and $Status -ne "Enabled") { $Include = $true }
    if ($Active -and $Status -eq "Enabled") { $Include = $true }

    if ($Include) {
        $Results += [PSCustomObject]@{
            "Subscription Name"                             = $SkuPartNumber
            "Friendly Subscription Name"                    = $FriendlyName
            "Subscribed Date"                               = $SubscribedString
            "Total Units"                                   = $Total
            "Consumed Units"                                = $Consumed
            "Remaining Units"                               = $Remaining
            "Subscription Type"                             = $Type
            "License Expiry Date / Next Lifecycle Activity" = $ExpiryDateTime
            "Friendly Expiry Date"                          = $ExpiryNote
            "Status"                                        = $Status
            "SKU Id"                                        = $SkuId
        }
    }
}

# ------------------------- Export -------------------------
if ($Results.Count -gt 0) {
    $Results | Export-Csv -Path $ExportCSV -NoTypeInformation
    Write-Host "`n✅ Report saved to:" -NoNewline; Write-Host " $ExportCSV" -ForegroundColor Cyan
    Write-Host "📦 $($Results.Count) subscriptions included.`n"
}
else {
    Write-Host "⚠️ No subscriptions matched the given filters." -ForegroundColor Yellow
}
