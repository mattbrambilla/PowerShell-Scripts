# SharePoint Online

Requires `Microsoft.Online.SharePoint.PowerShell` (`Connect-SPOService`).

| Script | Purpose |
|---|---|
| `Bulk-Add-TenantSiteAdmins.ps1` | Add one or more accounts as site collection administrators on team and communication sites. Supports `-WhatIf` and CSV logging. |
| `Manage-SharePointVersionHistory.ps1` | Analyze SharePoint version history usage, estimate the storage footprint, and optionally delete older file versions with `-DeleteAllVersions`. |
