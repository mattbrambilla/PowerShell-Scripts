# Exchange Online

Requires [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement). Connect with `Connect-ExchangeOnline` before running (unless the script already does it).

| Script | Purpose |
|---|---|
| `Enable-RetentionPolicy.ps1` | Enable retention hold on a mailbox. Replace `[Your Mailbox adress here]`. |
| `New-AddressList.ps1` | Create an address list filtered by company. |
| `New-GlobalAddressList.ps1` | Create a global address list (GAL) filtered by company. |
| `Get-TeamsRoomAutoReleaseStatus.ps1` | List room mailboxes and their calendar auto-release status; includes sample enable/disable commands. |
| `Get-EmptySharedMailboxes.ps1` | Find shared mailboxes with no Full Access delegates besides `NT AUTHORITY\SELF`. |
