# Exchange Online

Requires [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement). Connect with `Connect-ExchangeOnline` before running (unless the script already does it).

| Script | Purpose |
|---|---|
| `enableRetentionPolicy.ps1` | Enable retention hold on a mailbox. Replace `[Your Mailbox adress here]`. |
| `createAdressList.ps1` | Create an address list filtered by company. |
| `createGlobalAdressList.ps1` | Create a global address list (GAL) filtered by company. |
| `teamsRoomAutoReleaseStatus.ps1` | List room mailboxes and their calendar auto-release status; includes sample enable/disable commands. |
| `List empty shared mailboxes (Read and manage).ps1` | Find shared mailboxes with no Full Access delegates besides `NT AUTHORITY\SELF`. |
