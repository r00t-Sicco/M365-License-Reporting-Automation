# ============================================================
# Microsoft 365 License Reporting Automation
# Public / Sanitized Example
#
# ============================================================
# USER CONFIGURATION
# ============================================================
#
# Before running this script, replace the values below with
# information from your Microsoft 365 environment.
#
# TENANT ID
#   Microsoft Entra admin center:
#   Identity > Overview > Tenant ID
#
# APPLICATION (CLIENT) ID
#   Microsoft Entra admin center:
#   App registrations > Your application > Overview
#   Copy "Application (client) ID"
#
# CERTIFICATE THUMBPRINT
#   This is the thumbprint of the certificate configured for
#   the Entra application.
#
#   To view certificates installed on this computer:
#
#   Get-ChildItem Cert:\LocalMachine\My |
#       Select-Object Subject, Thumbprint, HasPrivateKey, NotAfter
#
# ORGANIZATION NAME
#   Must match the Microsoft 365 organization display name.
#   The script verifies this after connecting to Graph to help
#   prevent accidentally running against the wrong tenant.
#
# SENDER
#   Mailbox used to send the completed report.
#
# RECIPIENTS
#   One or more addresses that should receive the report.
#
# IMPORTANT:
#   Never commit private keys, PFX files, passwords, client
#   secrets, production tenant IDs, certificate thumbprints,
#   or customer-specific information to a public repository.
#
# ============================================================

$OrganizationName = "Contoso Ltd."

$TenantId = "<YOUR-TENANT-ID>"
$ClientId = "<YOUR-APPLICATION-CLIENT-ID>"
$CertificateThumbprint = "<YOUR-CERTIFICATE-THUMBPRINT>"

$Sender = "m365reports@contoso.com"

$Recipients = @(
    "itadmin@contoso.com",
    "reports@contoso.com"
)

$BasePath = "C:\M365-LicenseReporter"

$InactiveThresholdDays = 90

# ============================================================
# Logging / Scheduled Task Error Handling
# ============================================================

$BasePath  = "C:\\M365-LicenseReporter"
$LogFolder = Join-Path $BasePath "Logs"
New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
$LogTimestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$LogPath = Join-Path $LogFolder "M365-License-Report-$LogTimestamp.log"
$TranscriptStarted = $false
$RunSucceeded = $false

try {
    Start-Transcript -Path $LogPath -Append | Out-Null
    $TranscriptStarted = $true
    Write-Host ""
    Write-Host "==============================================="
    Write-Host " Microsoft 365 License Reporter"
    Write-Host " Run started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host "==============================================="
    Write-Host ""

# ============================================================
# Contoso Ltd.
# Microsoft 365 Monthly License Report
# ============================================================

$ErrorActionPreference = "Stop"

$ReportRoot = "C:\M365-LicenseReporter\Reports"
$TempRoot   = "C:\M365-LicenseReporter\Temp"

$ReportDate = Get-Date
$ReportMonth = $ReportDate.ToString("yyyy-MM")

$MonthlyFolder = Join-Path $ReportRoot $ReportMonth

New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
New-Item -ItemType Directory -Path $MonthlyFolder -Force | Out-Null

Write-Host ""
Write-Host "==============================================="
Write-Host " Contoso Ltd. - M365 License Reporter"
Write-Host "==============================================="
Write-Host ""

# ------------------------------------------------------------
# Load Microsoft Graph modules
# ------------------------------------------------------------

Write-Host "[1/7] Loading Microsoft Graph modules..."

Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.32.0
Import-Module Microsoft.Graph.Users -RequiredVersion 2.32.0
Import-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion 2.32.0
Import-Module Microsoft.Graph.Reports -RequiredVersion 2.32.0

# ------------------------------------------------------------
# Connect to Microsoft Graph
# ------------------------------------------------------------

Write-Host "[2/7] Connecting to Microsoft Graph..."

Connect-MgGraph `
    -TenantId "<YOUR-TENANT-ID>" `
    -ClientId "<YOUR-APPLICATION-CLIENT-ID>" `
    -CertificateThumbprint "<YOUR-CERTIFICATE-THUMBPRINT>" `
    -NoWelcome

$Organization = Get-MgOrganization | Select-Object -First 1

Write-Host "      Connected to: $($Organization.DisplayName)"

if ($Organization.DisplayName -ne "Contoso Ltd.") {
    throw "Wrong Microsoft 365 tenant. Expected Contoso Ltd.."
}

# ------------------------------------------------------------
# License definitions
# ------------------------------------------------------------

$LicenseNames = @{
    "EXCHANGESTANDARD"         = "Exchange Online Plan 1"
    "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
    "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
}

$PaidSkuNames = @(
    "EXCHANGESTANDARD",
    "O365_BUSINESS_PREMIUM",
    "O365_BUSINESS_ESSENTIALS"
)

# ------------------------------------------------------------
# Get tenant licenses
# ------------------------------------------------------------

Write-Host "[3/7] Collecting Microsoft 365 licenses..."

$Skus = Get-MgSubscribedSku -All

$SkuLookup = @{}

foreach ($Sku in $Skus) {
    $SkuLookup[$Sku.SkuId.ToString()] = $Sku.SkuPartNumber
}

$LicenseSummary = foreach ($Sku in $Skus) {

    if ($Sku.SkuPartNumber -in $PaidSkuNames) {

        [PSCustomObject]@{
            License   = $LicenseNames[$Sku.SkuPartNumber]
            Purchased = $Sku.PrepaidUnits.Enabled
            Assigned  = $Sku.ConsumedUnits
            Available = $Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits
        }
    }
}

# ------------------------------------------------------------
# Get users and assigned paid licenses
# ------------------------------------------------------------

Write-Host "[4/7] Collecting licensed users..."

$Users = Get-MgUser -All -Property `
    "id,displayName,userPrincipalName,accountEnabled,assignedLicenses,userType"

$UserLicenseReport = foreach ($User in $Users) {

    $PaidLicenses = @(
        foreach ($License in $User.AssignedLicenses) {

            $SkuId = $License.SkuId.ToString()

            if ($SkuLookup.ContainsKey($SkuId)) {

                $SkuPartNumber = $SkuLookup[$SkuId]

                if ($SkuPartNumber -in $PaidSkuNames) {
                    $LicenseNames[$SkuPartNumber]
                }
            }
        }
    )

    if ($PaidLicenses.Count -gt 0) {

        [PSCustomObject]@{
            Name           = $User.DisplayName
            Email          = $User.UserPrincipalName
            License        = ($PaidLicenses -join ", ")
            AccountEnabled = $User.AccountEnabled
        }
    }
}

# ------------------------------------------------------------
# Download Microsoft 365 activity
# ------------------------------------------------------------

Write-Host "[5/7] Collecting 90-day Microsoft 365 activity..."

$ActivityFile = Join-Path $TempRoot "M365-Activity.csv"

Get-MgReportOffice365ActiveUserDetail `
    -Period "D90" `
    -OutFile $ActivityFile

$Activity = Import-Csv $ActivityFile

$ReportRefreshDate = [datetime](
    $Activity |
    Select-Object -First 1 -ExpandProperty 'Report Refresh Date'
)

$ActivityLookup = @{}

foreach ($Row in $Activity) {

    $Dates = @(
        $Row.'Exchange Last Activity Date'
        $Row.'OneDrive Last Activity Date'
        $Row.'SharePoint Last Activity Date'
        $Row.'Teams Last Activity Date'
    ) |
    Where-Object { $_ } |
    ForEach-Object { [datetime]$_ }

    if ($Dates.Count -gt 0) {

        $LastActivity = $Dates |
            Sort-Object -Descending |
            Select-Object -First 1

        $DaysInactive = ($ReportRefreshDate - $LastActivity).Days
    }
    else {
        $LastActivity = $null
        $DaysInactive = $null
    }

    $Email = $Row.'User Principal Name'.ToLower()

    $ActivityLookup[$Email] = [PSCustomObject]@{
        LastActivity = $LastActivity
        DaysInactive = $DaysInactive
    }
}

# ------------------------------------------------------------
# Build master report
# ------------------------------------------------------------

Write-Host "[6/7] Analyzing licensed accounts..."

$MasterReport = foreach ($User in $UserLicenseReport) {

    $EmailKey = $User.Email.ToLower()

    if ($ActivityLookup.ContainsKey($EmailKey)) {
        $UserActivity = $ActivityLookup[$EmailKey]
        $LastActivity = $UserActivity.LastActivity
        $DaysInactive = $UserActivity.DaysInactive
    }
    else {
        $LastActivity = $null
        $DaysInactive = $null
    }

    if ($User.AccountEnabled -eq $false) {
        $Status = "REVIEW - Disabled account is licensed"
    }
    elseif ($null -ne $DaysInactive -and $DaysInactive -ge 90) {
        $Status = "REVIEW - Inactive 90+ days"
    }
    elseif ($null -eq $LastActivity) {
        $Status = "REVIEW - No M365 activity recorded"
    }
    else {
        $Status = "OK"
    }

    [PSCustomObject]@{
        Name           = $User.Name
        Email          = $User.Email
        License        = $User.License
        AccountEnabled = $User.AccountEnabled
        LastActivity   = $LastActivity
        DaysInactive   = $DaysInactive
        Status         = $Status
    }
}

# ------------------------------------------------------------
# Summary
# ------------------------------------------------------------

$ReviewAccounts = @(
    $MasterReport |
    Where-Object {$_.Status -ne "OK"}
)

$DisabledLicensed = @(
    $MasterReport |
    Where-Object {$_.AccountEnabled -eq $false}
)

$InactiveAccounts = @(
    $MasterReport |
    Where-Object {
        $_.AccountEnabled -eq $true -and
        $null -ne $_.DaysInactive -and
        $_.DaysInactive -ge 90
    }
)

Write-Host "[7/7] Exporting report data..."

$LicenseSummaryPath = Join-Path $MonthlyFolder "License-Summary.csv"
$LicensedUsersPath  = Join-Path $MonthlyFolder "Licensed-Users.csv"
$ReviewPath         = Join-Path $MonthlyFolder "Accounts-To-Review.csv"

$LicenseSummary |
    Sort-Object License |
    Export-Csv $LicenseSummaryPath -NoTypeInformation

$MasterReport |
    Sort-Object Name |
    Export-Csv $LicensedUsersPath -NoTypeInformation

$ReviewAccounts |
    Sort-Object DaysInactive -Descending |
    Export-Csv $ReviewPath -NoTypeInformation

Write-Host ""
Write-Host "==============================================="
Write-Host " REPORT COMPLETE"
Write-Host "==============================================="
Write-Host ""
Write-Host "Activity data through: $($ReportRefreshDate.ToString('MMMM d, yyyy'))"
Write-Host "Paid licensed accounts: $($MasterReport.Count)"
Write-Host "Accounts requiring review: $($ReviewAccounts.Count)"
Write-Host "Disabled but licensed: $($DisabledLicensed.Count)"
Write-Host "Inactive 90+ days: $($InactiveAccounts.Count)"
Write-Host ""
Write-Host "License Summary:"
$LicenseSummary | Format-Table -AutoSize

Write-Host ""
Write-Host "Accounts Requiring Review:"
$ReviewAccounts |
    Sort-Object DaysInactive -Descending |
    Format-Table Name,License,LastActivity,DaysInactive,Status -AutoSize

Write-Host ""
Write-Host "Files saved to:"
Write-Host $MonthlyFolder
Write-Host ""
# ============================================================
# Generate HTML Report
# ============================================================

Write-Host "Generating HTML report..."

$HtmlReportPath = Join-Path $MonthlyFolder "M365-License-Report-$ReportMonth.html"

# Month represented by Microsoft's activity data
$ReportDisplayMonth = $ReportRefreshDate.ToString("MMMM yyyy")
$ActivityThrough    = $ReportRefreshDate.ToString("MMMM d, yyyy")

# -------------------------------
# License Summary Rows
# -------------------------------

$LicenseRows = ""

foreach ($Item in ($LicenseSummary | Sort-Object License)) {

    $LicenseRows += @"
        <tr>
            <td>$($Item.License)</td>
            <td>$($Item.Purchased)</td>
            <td>$($Item.Assigned)</td>
            <td>$($Item.Available)</td>
        </tr>
"@
}

# -------------------------------
# Licensed User Rows
# -------------------------------

$UserRows = ""

foreach ($User in ($MasterReport | Sort-Object Name)) {

    if ($User.LastActivity) {
        $LastActivityText = $User.LastActivity.ToString("MM/dd/yyyy")
    }
    else {
        $LastActivityText = "No activity recorded"
    }

    if ($User.Status -eq "OK") {
        $StatusClass = "ok"
        $StatusText  = "Active"
    }
    else {
        $StatusClass = "review"
        $StatusText  = "Review"
    }

    $UserRows += @"
        <tr>
            <td>$($User.Name)</td>
            <td>$($User.Email)</td>
            <td>$($User.License)</td>
            <td>$LastActivityText</td>
            <td class="$StatusClass">$StatusText</td>
        </tr>
"@
}

# -------------------------------
# Review Rows
# -------------------------------

$ReviewRows = ""

foreach ($User in ($ReviewAccounts | Sort-Object DaysInactive -Descending)) {

    if ($User.LastActivity) {
        $LastActivityText = $User.LastActivity.ToString("MM/dd/yyyy")
    }
    else {
        $LastActivityText = "No activity recorded"
    }

    if ($User.AccountEnabled -eq $false) {
        $Reason = "Disabled account is still licensed"
    }
    elseif ($null -eq $User.LastActivity) {
        $Reason = "No Microsoft 365 activity recorded"
    }
    else {
        $Reason = "Inactive for $($User.DaysInactive) days"
    }

    $ReviewRows += @"
        <tr>
            <td>$($User.Name)</td>
            <td>$($User.License)</td>
            <td>$LastActivityText</td>
            <td>$Reason</td>
        </tr>
"@
}

# -------------------------------
# HTML Template
# -------------------------------

$Html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">

<title>Contoso Ltd. - Microsoft 365 License Report</title>

<style>

body {
    font-family: "Segoe UI", Arial, sans-serif;
    margin: 0;
    background: #f4f6f8;
    color: #222;
}

.container {
    max-width: 1200px;
    margin: 30px auto;
    background: white;
    padding: 40px;
    box-shadow: 0 2px 10px rgba(0,0,0,.08);
}

h1 {
    margin-bottom: 4px;
    font-size: 28px;
}

.subtitle {
    color: #666;
    font-size: 16px;
    margin-bottom: 5px;
}

.refresh {
    color: #888;
    font-size: 13px;
    margin-bottom: 30px;
}

.summary {
    display: flex;
    gap: 15px;
    margin-bottom: 35px;
}

.card {
    flex: 1;
    background: #f5f7fa;
    border: 1px solid #dde2e7;
    border-radius: 6px;
    padding: 18px;
}

.card-number {
    font-size: 28px;
    font-weight: 600;
}

.card-label {
    margin-top: 5px;
    color: #666;
    font-size: 13px;
}

h2 {
    margin-top: 35px;
    font-size: 19px;
    border-bottom: 2px solid #ddd;
    padding-bottom: 8px;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 15px;
    font-size: 13px;
}

th {
    text-align: left;
    background: #f0f2f5;
    padding: 10px;
    border-bottom: 2px solid #ccc;
}

td {
    padding: 9px 10px;
    border-bottom: 1px solid #e4e4e4;
}

.ok {
    font-weight: 600;
}

.review {
    font-weight: 600;
}

.review-table {
    border: 1px solid #ddd;
}

.note {
    margin-top: 30px;
    font-size: 12px;
    color: #777;
}

@media print {

    body {
        background: white;
    }

    .container {
        box-shadow: none;
        margin: 0;
        max-width: none;
    }

    tr {
        page-break-inside: avoid;
    }
}

</style>

</head>

<body>

<div class="container">

<h1>Contoso Ltd.</h1>

<div class="subtitle">
Microsoft 365 License Report - $ReportDisplayMonth
</div>

<div class="refresh">
Microsoft 365 activity data through $ActivityThrough
</div>


<div class="summary">

    <div class="card">
        <div class="card-number">$($MasterReport.Count)</div>
        <div class="card-label">Paid Licensed Accounts</div>
    </div>

    <div class="card">
        <div class="card-number">$($ReviewAccounts.Count)</div>
        <div class="card-label">Accounts to Review</div>
    </div>

    <div class="card">
        <div class="card-number">$($DisabledLicensed.Count)</div>
        <div class="card-label">Disabled + Licensed</div>
    </div>

    <div class="card">
        <div class="card-number">$($InactiveAccounts.Count)</div>
        <div class="card-label">Inactive 90+ Days</div>
    </div>

</div>


<h2>License Summary</h2>

<table>

<thead>
<tr>
    <th>License</th>
    <th>Purchased</th>
    <th>Assigned</th>
    <th>Available</th>
</tr>
</thead>

<tbody>
$LicenseRows
</tbody>

</table>


<h2>Accounts to Review</h2>

<table class="review-table">

<thead>
<tr>
    <th>Name</th>
    <th>License</th>
    <th>Last Activity</th>
    <th>Reason</th>
</tr>
</thead>

<tbody>
$ReviewRows
</tbody>

</table>


<h2>Licensed Accounts</h2>

<table>

<thead>
<tr>
    <th>Name</th>
    <th>Email</th>
    <th>License</th>
    <th>Last Activity</th>
    <th>Status</th>
</tr>
</thead>

<tbody>
$UserRows
</tbody>

</table>


<div class="note">
This report is generated automatically from Microsoft 365 licensing and usage data.
Accounts marked for review should be verified before any licensing changes are made.
</div>

</div>

</body>
</html>
"@

$Html | Out-File $HtmlReportPath -Encoding UTF8

Write-Host ""
Write-Host "HTML report created:"
Write-Host $HtmlReportPath
Write-Host ""

# ============================================================
# Generate PDF Report
# ============================================================

Write-Host "Generating PDF report..."

$PdfReportPath = Join-Path $MonthlyFolder "M365-License-Report-$ReportMonth.pdf"

$EdgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
)

$EdgePath = $EdgePaths |
    Where-Object { Test-Path $_ } |
    Select-Object -First 1

if (-not $EdgePath) {
    throw "Microsoft Edge could not be found. PDF generation cannot continue."
}

$HtmlFullPath = (Resolve-Path $HtmlReportPath).Path
$HtmlUri = [System.Uri]::new($HtmlFullPath).AbsoluteUri

$EdgeArguments = @(
    "--headless"
    "--disable-gpu"
    "--no-pdf-header-footer"
    "--print-to-pdf=$PdfReportPath"
    $HtmlUri
)

$Process = Start-Process `
    -FilePath $EdgePath `
    -ArgumentList $EdgeArguments `
    -Wait `
    -PassThru `
    -WindowStyle Hidden

if ($Process.ExitCode -ne 0) {
    throw "Edge PDF generation failed with exit code $($Process.ExitCode)."
}

if (-not (Test-Path $PdfReportPath)) {
    throw "PDF generation completed but the PDF file was not found."
}

Write-Host ""
Write-Host "PDF report created:"
Write-Host $PdfReportPath
Write-Host ""

# ============================================================
# Email PDF Report
# ============================================================

Write-Host "Emailing PDF report..."

$Sender = "m365reports@contoso.com"

$Recipients = @(
    "itadmin@contoso.com",
    "reports@contoso.com"
)

$PdfBytes  = [System.IO.File]::ReadAllBytes($PdfReportPath)
$PdfBase64 = [System.Convert]::ToBase64String($PdfBytes)
$PdfName   = [System.IO.Path]::GetFileName($PdfReportPath)

$EmailBody = @"
<p>Hello,</p>

<p>Attached is the Contoso Ltd. Microsoft 365 License Report for
<strong>$ReportDisplayMonth</strong>.</p>

<p><strong>Summary:</strong></p>

<ul>
    <li>Paid licensed accounts: $($MasterReport.Count)</li>
    <li>Accounts requiring review: $($ReviewAccounts.Count)</li>
    <li>Disabled but licensed: $($DisabledLicensed.Count)</li>
    <li>Inactive 90+ days: $($InactiveAccounts.Count)</li>
</ul>

<p>Microsoft 365 activity data is current through $ActivityThrough.</p>

<p>Please see the attached report for full licensing details and accounts requiring review.</p>

<p>This report was generated automatically.</p>
"@

$Mail = @{
    message = @{
        subject = "Contoso Ltd. - Microsoft 365 License Report - $ReportDisplayMonth"

        body = @{
            contentType = "HTML"
            content     = $EmailBody
        }

        toRecipients = @(
    foreach ($Recipient in $Recipients) {
        @{
            emailAddress = @{
                address = $Recipient
            }
        }
    }
)

        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = $PdfName
                contentType   = "application/pdf"
                contentBytes  = $PdfBase64
            }
        )
    }

    saveToSentItems = $true
}

Invoke-MgGraphRequest `
    -Method POST `
    -Uri "https://graph.microsoft.com/v1.0/users/$Sender/sendMail" `
    -Body ($Mail | ConvertTo-Json -Depth 10) `
    -ContentType "application/json"

Write-Host ""
Write-Host "Report emailed successfully."
Write-Host "From: $Sender"
Write-Host "To:   $($Recipients -join ', ')"
Write-Host ""

    $RunSucceeded = $true
    Write-Host ""
    Write-Host "==============================================="
    Write-Host " AUTOMATION COMPLETE"
    Write-Host " Run completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    if ($PdfReportPath) { Write-Host " Report: $PdfReportPath" }
    Write-Host " Log:    $LogPath"
    Write-Host "==============================================="
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "==============================================="
    Write-Host " AUTOMATION FAILED"
    Write-Host " Time:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host " Error: $($_.Exception.Message)"
    Write-Host "==============================================="
    Write-Host ""
    Write-Error $_
}
finally {
    if ($TranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
    }
}

if ($RunSucceeded) { exit 0 } else { exit 1 }
