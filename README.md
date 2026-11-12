# PowerShell scripts

Collection of PowerShell scripts for **Microsoft 365**, **Microsoft Entra ID**, and a few **Windows** admin tasks.

Scripts are grouped by product so you can open a folder and see only what you need. Replace placeholders before you run anything.

## Scripts

| Folder | What you will find |
|---|---|
| [Exchange-Online](Exchange-Online/) | Retention hold, address lists, room mailbox auto-release, empty shared mailboxes |
| [Microsoft-Entra](Microsoft-Entra/) | MFA authentication-methods report, remove app permissions |
| [SharePoint-Online](SharePoint-Online/) | Bulk site-collection admins and SharePoint version-history cleanup |
| [Teams](Teams/) | Export teams and members (optionally with SharePoint) |
| [OneDrive](OneDrive/) | List files shared via OneDrive |
| [Windows](Windows/) | Remote lock, system language, copy files, local Mail duplicate cleanup |

Each folder has its own `README.md` with a one-line description per script.

## Prerequisites

Install the module you need, then connect before running the related scripts:

| Workload | Module | Connect |
|---|---|---|
| Exchange Online | `ExchangeOnlineManagement` | `Connect-ExchangeOnline` |
| Microsoft Graph / Entra | `Microsoft.Graph` | `Connect-MgGraph` |
| SharePoint Online | `Microsoft.Online.SharePoint.PowerShell` | `Connect-SPOService` |
| Teams | `MicrosoftTeams` | `Connect-MicrosoftTeams` |

Some older scripts still use `MSOnline`, `AzureADPreview`, or raw Graph REST. Prefer interactive Graph / Exchange modules when you update a script.

## Placeholders

Several scripts ship with sample values. Replace them with your tenant data:

- `TuoClientID`, `TuoClientSecret`, `TuoTenantID`
- `[Your Mailbox adress here]`, `youremail@provider.com`
- `example.com`, `user1@example.com`
- `contoso.sharepoint.com`, `https://example-admin.sharepoint.com`

Do not commit real client secrets.

## Example

```powershell
.\SharePoint-Online\Bulk-Add-TenantSiteAdmins.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -AdminAccounts @("admin@contoso.com") `
  -WhatIf
```

## License

See [LICENSE.txt](LICENSE.txt).
