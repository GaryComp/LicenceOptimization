Param
(
    [Parameter(Mandatory = $false)]
    [string]$UserNamesFile,
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

    Write-Host "Connected to Microsoft Graph successfully." -ForegroundColor Green
}

# ------------------------- Export User License Details -------------------------
Function Export-UserLicenseDetails {
    param (
        [Parameter(Mandatory=$true)][object]$User,
        [string]$ExportCSV,
        [object]$SkuMapping,
        [object]$ServiceMapping
    )

    $UPN = $User.UserPrincipalName
    $DisplayName = $User.DisplayName
    $Country = if ($User.Country) { $User.Country } else { "-" }

    Write-Progress -Activity "Exported user count: $Global:LicensedUserCount" -Status "Processing: $UPN"

    $SKUs = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction SilentlyContinue
    if (!$SKUs) { return }

    foreach ($Sku in $SKUs) {
        $SkuId = $Sku.SkuPartNumber.Trim()
        $FriendlyLicense = $SkuId
        if ($SkuMapping) {
            $match = $SkuMapping | Where-Object { $_.SkuPartNumber -and ($_.SkuPartNumber.Trim().ToLower() -eq $SkuId.ToLower()) }
            if ($match) { $FriendlyLicense = $match.Product_Display_Name }
        }

        foreach ($Service in $Sku.ServicePlans) {
            $ServiceName = $Service.ServicePlanName.Trim()
            $FriendlyService = $ServiceName
            if ($ServiceMapping) {
                $match = $ServiceMapping | Where-Object { $_.Service_Plan_Name -and ($_.Service_Plan_Name.Trim().ToLower() -eq $ServiceName.ToLower()) }
                if ($match) { $FriendlyService = $match.ServicePlanDisplayName }
            }

            [PSCustomObject]@{
                DisplayName                = $DisplayName
                UserPrincipalName          = $UPN
                Country                    = $Country
                LicenseSkuPartNumber       = $SkuId
                LicenseFriendlyName        = $FriendlyLicense
                ServicePlanName            = $ServiceName
                ServicePlanFriendlyName    = $FriendlyService
                ProvisioningStatus         = $Service.ProvisioningStatus
            } | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
        }
    }
}

# ------------------------- Close Connection -------------------------
Function Close-Connection {
    Disconnect-MgGraph | Out-Null
    Exit
}

# ------------------------- Main -------------------------
Function main {
    Connect-ToGraph
    Write-Host "`nNote: For best results, run this in a fresh PowerShell window." -ForegroundColor Yellow

    $timestamp = (Get-Date -Format "yyyy-MMM-dd-ddd hh-mm tt").ToString()
    $ExportCSV = ".\DetailedO365UserLicenseReport_$timestamp.csv"

    # Try to load mapping file (optional)
    $SkuMapping = $null
    $ServiceMapping = $null
    if (Test-Path "Product_names_and_service_plan_identifiers.csv") {
        try {
            $csvData = Import-Csv -Path ".\Product_names_and_service_plan_identifiers.csv"
            if ($csvData | Get-Member -Name "SkuPartNumber" -ErrorAction SilentlyContinue) {
                $SkuMapping = $csvData | Where-Object { $_.SkuPartNumber -and $_.Product_Display_Name }
            }
            if ($csvData | Get-Member -Name "Service_Plan_Name" -ErrorAction SilentlyContinue) {
                $ServiceMapping = $csvData | Where-Object { $_.Service_Plan_Name -and $_.ServicePlanDisplayName }
            }
        } catch {
            Write-Host "⚠️ Failed to parse mapping file, continuing without friendly names." -ForegroundColor Yellow
        }
    }

    $Global:LicensedUserCount = 0

    if ($UserNamesFile) {
        $UserNames = Import-Csv -Header "UserPrincipalName" $UserNamesFile
        foreach ($item in $UserNames) {
            $user = Get-MgUser -UserId $item.UserPrincipalName -ErrorAction SilentlyContinue
            if ($user) {
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
                if ($licenseDetails) {
                    Export-UserLicenseDetails -User $user -ExportCSV $ExportCSV -SkuMapping $SkuMapping -ServiceMapping $ServiceMapping
                    $Global:LicensedUserCount++
                }
            }
        }
    } else {
        Get-MgUser -All | ForEach-Object {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $_.Id -ErrorAction SilentlyContinue
            if ($licenseDetails) {
                Export-UserLicenseDetails -User $_ -ExportCSV $ExportCSV -SkuMapping $SkuMapping -ServiceMapping $ServiceMapping
                $Global:LicensedUserCount++
            }
        }
    }

    if (Test-Path -Path $ExportCSV) {
        Write-Host "`n✅ Detailed report available at: $ExportCSV" -ForegroundColor Cyan
        Write-Host "📦 $Global:LicensedUserCount users processed." -ForegroundColor Green
    } else {
        Write-Host "⚠️ No data found." 
    }
    Close-Connection
}

main
