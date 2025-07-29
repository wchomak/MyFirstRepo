param (
    [switch]$WhatIf  # Run as: .\SetLicenseAppsBulk.ps1 -WhatIf
)

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "User.ReadWrite.All", "Organization.Read.All"

# Define the license SKU
$targetSkuPartNumber = "Microsoft_365_E3_(no_Teams)"
$auditLog = @()

# Get the license SKU object
$sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $targetSkuPartNumber }

if (-not $sku) {
    Write-Host "License SKU not found: $targetSkuPartNumber" -ForegroundColor Red
    return
}

$skuId = $sku.SkuId
Write-Host "Found SKU: $targetSkuPartNumber ($skuId)" -ForegroundColor Cyan

# Define service plans to EXCLUDE
$excludedServiceNames = @(
    "KAIZALA_O365_P3",         # Microsoft Kaizala Pro
    "Deskless",                # Microsoft StaffHub
    "STREAM_O365_E3",          # Microsoft Stream
    "MCOSTANDARD",             # Microsoft Teams & Skype
    "SWAY",                    # Sway
    "VIVAENGAGE_CORE",         # Viva Engage
    "VIVA_LEARNING_SEEDED",    # Viva Learning
    "YAMMER_ENTERPRISE"        # Yammer
)

# Get the ServicePlanId values to disable (flattened correctly)
$disabledPlans = $sku.ServicePlans |
    Where-Object { $_.ServicePlanName -in $excludedServiceNames } |
    ForEach-Object { $_.ServicePlanId }

# Get all users and include ID
$allUsers = Get-MgUser -All -Property "Id", "UserPrincipalName", "DisplayName", "AssignedLicenses"
$licensedUsers = $allUsers | Where-Object { $_.AssignedLicenses.SkuId -contains $skuId }

# Select first two users (adjust or remove limit as needed)
$targetUsers = $licensedUsers | Select-Object -First 150

if (-not $targetUsers) {
    Write-Host "No users found with license $targetSkuPartNumber." -ForegroundColor Yellow
    return
}

# Loop through users
foreach ($user in $targetUsers) {
    # Fetch full license details
    $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id | Where-Object { $_.SkuId -eq $skuId }

    # Check if any excluded plans are still enabled
    $enabledExcludedPlans = $licenseDetails.ServicePlans | Where-Object {
        $_.ProvisioningStatus -eq "Success" -and
        $excludedServiceNames -contains $_.ServicePlanName
    }

    $logEntry = [PSCustomObject]@{
        Timestamp         = (Get-Date).ToString("s")
        UserPrincipalName = $user.UserPrincipalName
        DisplayName       = $user.DisplayName
        Status            = ""
        Error             = ""
    }

    if (-not $enabledExcludedPlans) {
        Write-Host "⏭ Skipping $($user.UserPrincipalName) — all excluded services already disabled." -ForegroundColor DarkGray
        $logEntry.Status = "Skipped - Already Compliant"
        $auditLog += $logEntry
        continue
    }

    if ($WhatIf) {
        Write-Host "🔎 [WHATIF] Would update license for $($user.UserPrincipalName), disabling:" -ForegroundColor Yellow
        $enabledExcludedPlans.ServicePlanName | ForEach-Object { Write-Host " - $_" }
        $logEntry.Status = "WhatIf"
    }
    else {
        Write-Host "⚙ Updating license for $($user.UserPrincipalName) ..." -ForegroundColor Green

        $licenseAssignment = @{
            AddLicenses = @(
                @{
                    SkuId = $sku.SkuId
                    DisabledPlans = $disabledPlans
                }
            )
            RemoveLicenses = @()
        }

        try {
            Set-MgUserLicense -UserId $user.Id -BodyParameter $licenseAssignment
            Write-Host "✔ Successfully updated $($user.UserPrincipalName)" -ForegroundColor Cyan
            $logEntry.Status = "Success"
        }
        catch {
            Write-Host "❌ Failed to update $($user.UserPrincipalName): $_" -ForegroundColor Red
            $logEntry.Status = "Failure"
            $logEntry.Error = $_.Exception.Message
        }
    }

    $auditLog += $logEntry
}

# Export audit log to Desktop
$auditPath = [System.IO.Path]::Combine(
    [Environment]::GetFolderPath("Desktop"),
    "License_AuditLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)
$auditLog | Export-Csv -Path $auditPath -NoTypeInformation
Write-Host "`n📄 Audit log saved to:`n$auditPath" -ForegroundColor Yellow
