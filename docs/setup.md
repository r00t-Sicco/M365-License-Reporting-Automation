# Deployment Guide

This guide covers deployment of the Microsoft 365 License Reporting Automation.

The automation uses:

- Microsoft Graph application permissions
- Microsoft Entra ID app registration
- Certificate-based authentication
- Exchange Online Application RBAC
- PowerShell
- Microsoft Edge for PDF generation
- Windows Task Scheduler

> This repository uses example values. Replace all placeholders with values from your own environment.

---

## 1. Requirements

The automation host requires:

- Windows 10/11 or Windows Server
- Windows PowerShell 5.1 or PowerShell 7
- Microsoft Edge
- Internet access to Microsoft 365 services
- Microsoft Graph PowerShell modules
- A certificate containing a private key
- An Entra ID application registration
- Appropriate Microsoft Graph application permissions

The machine should remain available when the scheduled report is configured to run.

---

## 2. Install Microsoft Graph PowerShell

Install the required Microsoft Graph modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope AllUsers
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope AllUsers
Install-Module Microsoft.Graph.Users -Scope AllUsers
Install-Module Microsoft.Graph.Reports -Scope AllUsers
```

Verify installation:

```powershell
Get-Module Microsoft.Graph* -ListAvailable |
    Select-Object Name,Version,ModuleBase
```

---

## 3. Create the Entra ID Application

Open the **Microsoft Entra admin center**.

Navigate to:

```text
Identity
└── Applications
    └── App registrations
        └── New registration
```

Example name:

```text
M365 License Reporter
```

For a reporting system intended for one organization, select:

```text
Accounts in this organizational directory only
```

Register the application.

From the application's **Overview** page, record:

```text
Application (client) ID
Directory (tenant) ID
```

These values will later be entered into the PowerShell script.

---

## 4. Configure Microsoft Graph Permissions

Open:

```text
App registrations
└── M365 License Reporter
    └── API permissions
        └── Add a permission
            └── Microsoft Graph
                └── Application permissions
```

Configure the permissions required by the reporting script:

```text
User.Read.All
Organization.Read.All
LicenseAssignment.Read.All
Reports.Read.All
```

If automated email delivery is enabled, the application also requires:

```text
Mail.Send
```

These are **Application permissions**, not Delegated permissions.

After configuring the permissions, select:

```text
Grant admin consent
```

Review permissions before granting tenant-wide consent.

---

## 5. Create the Authentication Certificate

On the Windows computer that will execute the automation, open PowerShell as Administrator.

Example:

```powershell
$Cert = New-SelfSignedCertificate `
    -Subject "CN=M365 License Reporter" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)
```

Display the certificate information:

```powershell
$Cert |
    Select-Object Subject,Thumbprint,NotAfter
```

The **thumbprint** will be used by the reporting script.

---

## 6. Export the Public Certificate

Export the public portion of the certificate:

```powershell
Export-Certificate `
    -Cert $Cert `
    -FilePath "C:\M365-LicenseReporter\M365-License-Reporter.cer"
```

The `.cer` file contains the public certificate and can be uploaded to the Entra application.

Navigate to:

```text
Entra ID
└── App registrations
    └── M365 License Reporter
        └── Certificates & secrets
            └── Certificates
                └── Upload certificate
```

Upload the `.cer` file.

> The private key must remain protected on the automation host. Never upload a `.pfx` file or private key to a public repository.

---

## 7. Configure the Script

Edit:

```text
src\M365-LicenseReport.ps1
```

Update the configuration section:

```powershell
$OrganizationName = "Test Organization"

$TenantId = "<YOUR-TENANT-ID>"
$ClientId = "<YOUR-APPLICATION-CLIENT-ID>"
$CertificateThumbprint = "<YOUR-CERTIFICATE-THUMBPRINT>"

$Sender = "m365reports@example.com"

$Recipients = @(
    "itadmin@example.com"
)

$BasePath = "C:\M365-LicenseReporter"
$InactiveThresholdDays = 90
```

### Configuration Reference

| Variable | Purpose |
|---|---|
| `$OrganizationName` | Expected Microsoft 365 organization name |
| `$TenantId` | Microsoft Entra tenant identifier |
| `$ClientId` | Application registration client identifier |
| `$CertificateThumbprint` | Authentication certificate installed on the host |
| `$Sender` | Mailbox used for automated report delivery |
| `$Recipients` | Report recipients |
| `$BasePath` | Local working directory |
| `$InactiveThresholdDays` | Inactivity threshold used for review |

---

## 8. Test Certificate Authentication

Test the application authentication before scheduling the script:

```powershell
Connect-MgGraph `
    -TenantId "<YOUR-TENANT-ID>" `
    -ClientId "<YOUR-APPLICATION-CLIENT-ID>" `
    -CertificateThumbprint "<YOUR-CERTIFICATE-THUMBPRINT>" `
    -NoWelcome
```

Verify the authentication context:

```powershell
Get-MgContext |
    Format-List TenantId,ClientId,AuthType,CertificateThumbprint
```

A successful application-authenticated connection should report:

```text
AuthType : AppOnly
```

---

## 9. Restrict Automated Email with Exchange Application RBAC

Microsoft Graph `Mail.Send` application permission can provide broad application-level mail capability.

Exchange Online Application RBAC can be used to restrict the reporting application's email access to the designated reporting mailbox.

Install Exchange Online PowerShell if required:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

Connect:

```powershell
Connect-ExchangeOnline
```

### Register the Service Principal in Exchange

Locate the Entra enterprise application's service principal Object ID.

Then register it with Exchange Online:

```powershell
New-ServicePrincipal `
    -AppId "<YOUR-APPLICATION-CLIENT-ID>" `
    -ObjectId "<YOUR-SERVICE-PRINCIPAL-OBJECT-ID>" `
    -DisplayName "M365 License Reporter"
```

> The service principal Object ID is different from the application/client ID.

---

## 10. Create the Mailbox Resource Scope

Create a management scope containing only the mailbox authorized to send reports:

```powershell
New-ManagementScope `
    -Name "M365 License Reporter - Reporting Mailbox Only" `
    -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'm365reports@example.com'"
```

Verify which mailbox matches the scope:

```powershell
Get-Recipient `
    -RecipientPreviewFilter (
        Get-ManagementScope "M365 License Reporter - Reporting Mailbox Only"
    ).RecipientFilter |
    Select-Object DisplayName,PrimarySmtpAddress
```

Only the intended reporting mailbox should be returned.

---

## 11. Assign Application Mail.Send

Assign the Exchange application role:

```powershell
New-ManagementRoleAssignment `
    -Name "M365 License Reporter - Mail Send" `
    -Role "Application Mail.Send" `
    -App "<YOUR-APPLICATION-CLIENT-ID>" `
    -CustomResourceScope "M365 License Reporter - Reporting Mailbox Only"
```

Test authorization:

```powershell
Test-ServicePrincipalAuthorization `
    -Identity "<YOUR-APPLICATION-CLIENT-ID>" `
    -Resource "m365reports@example.com"
```

The authorized reporting mailbox should show:

```text
InScope : True
```

Testing another mailbox should show:

```text
InScope : False
```

This provides an additional security boundary around automated email delivery.

---

## 12. Deploy the Script

Create:

```text
C:\M365-LicenseReporter
```

Place the production copy of the script at:

```text
C:\M365-LicenseReporter\M365-LicenseReport.ps1
```

The script will maintain directories for generated reports, logs, and temporary files.

Example:

```text
C:\M365-LicenseReporter\
│
├── M365-LicenseReport.ps1
├── Reports\
├── Logs\
└── Temp\
```

---

## 13. Test a Complete Run

Before creating the scheduled task, execute the script manually:

```powershell
Set-Location "C:\M365-LicenseReporter"

.\M365-LicenseReport.ps1
```

Confirm that the run:

1. Authenticates to Microsoft Graph
2. Retrieves license information
3. Retrieves licensed users
4. Downloads Microsoft 365 activity information
5. Analyzes accounts
6. Creates CSV exports
7. Generates the HTML report
8. Generates the PDF report
9. Sends the report email
10. Writes an execution log

Review the generated files before enabling unattended execution.

---

## 14. Configure Windows Task Scheduler

Open **Task Scheduler** and select:

```text
Create Task
```

### General

Configure the task to run using an account with access to:

- The authentication certificate/private key
- The reporting directory
- Required PowerShell modules

For unattended execution, configure:

```text
Run whether user is logged on or not
```

and:

```text
Run with highest privileges
```

when required by the environment.

### Trigger

Create the desired recurring schedule.

Example:

```text
Monthly
Day: 1
Time: 8:00 AM
```

### Action

**Program/script**

```text
powershell.exe
```

**Add arguments**

```text
-NoProfile -ExecutionPolicy Bypass -File "C:\M365-LicenseReporter\M365-LicenseReport.ps1"
```

**Start in**

```text
C:\M365-LicenseReporter
```

Save the task.

---

## 15. Test the Scheduled Task

Do not wait for the first scheduled date.

Right-click the task and select:

```text
Run
```

Verify:

- Task Scheduler reports successful execution
- A new execution log is created
- Report files are generated
- The PDF opens successfully
- The email is delivered
- The attachment is present
- The sender is the expected reporting mailbox

---

## Logging and Troubleshooting

Execution logs are written under:

```text
C:\M365-LicenseReporter\Logs
```

When troubleshooting an unattended failure, check:

1. The latest automation log
2. Task Scheduler history/result code
3. Certificate expiration
4. Certificate private-key availability
5. Microsoft Graph permissions
6. Application admin consent
7. Exchange Application RBAC authorization
8. PowerShell module availability
9. Network connectivity
10. PDF generation / Edge availability

---

## Certificate Renewal

Certificates expire.

Periodically check:

```powershell
Get-ChildItem Cert:\LocalMachine\My |
    Select-Object Subject,Thumbprint,HasPrivateKey,NotAfter
```

Before the authentication certificate expires:

1. Create a replacement certificate
2. Upload its public certificate to the Entra application
3. Install the certificate/private key on the automation host
4. Update `$CertificateThumbprint`
5. Test authentication
6. Test a complete report run
7. Remove the old certificate after successful migration

---

## Security Notes

- Do not store administrator passwords in the script.
- Do not commit `.pfx` files or private keys.
- Do not publish production credentials.
- Protect access to the automation host.
- Restrict application permissions to those required.
- Use Exchange Application RBAC to constrain application mail access where appropriate.
- Review application permissions periodically.
- Monitor certificate expiration.
- Keep PowerShell modules and the automation host maintained.

---

## Production Safety

This automation is designed as a **reporting and review system**.

It identifies potentially unused Microsoft 365 licenses but does **not** automatically:

- Remove licenses
- Disable users
- Delete accounts
- Modify mailboxes
- Change Microsoft 365 configuration

Any remediation remains an administrator decision.
