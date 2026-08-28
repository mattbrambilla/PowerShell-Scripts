# AGENTS.md — Powershell-Scirpts

Script PowerShell per amministrazione Microsoft 365 / Entra ID.
Remote: `https://github.com/mattbrambilla/Powershell-Scirpts.git` (branch: `master`)

## Attenzione

- **README.md** (inglese) è l'indice pubblico del repo; ogni cartella prodotto ha un README con l'elenco script.
- Script contengono segnaposto da sostituire prima dell'esecuzione:
  - `TuoClientID`, `TuoClientSecret`, `TuoTenantID`
  - `[Your Mailbox adress here]`
  - `contoso.sharepoint.com`, `admin1contoso@example.com`
- Mix di commenti in italiano e inglese.

## Struttura directory

| Directory/file | Contenuto |
|---|---|
| `Exchange-Online/` | Retention policy, address list, GAL, room mailbox auto-release, shared mailbox vuote |
| `Microsoft-Entra/` | Report metodi MFA, rimozione permessi app |
| `SharePoint-Online/` | Aggiunta admin di sito in blocco |
| `Teams/` | Export team e membri |
| `OneDrive/` | File condivisi OneDrive |
| `Windows/` | Blocco remoto PC, lingua sistema, copia file, pulizia duplicati Mail |
| `templates/` | Boilerplate README/CHANGELOG (non è la documentazione del progetto) |

## Moduli PowerShell richiesti

Installare se mancanti, poi connettersi prima di eseguire script:

- `ExchangeOnlineManagement` → `Connect-ExchangeOnline`
- `Microsoft.Graph` (o `Microsoft.Graph.Authentication`) → `Connect-MgGraph`
- `Microsoft.Online.SharePoint.PowerShell` → `Connect-SPOService`

## Convenzioni stile

- Comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.LINK`, `.NOTES`, `.CHANGELOG`)
- Variabili in cima. Colori: `cyan` (sistema), `green` (processo), `red` (errore), `yellow` (avviso)
- Selezione interattiva con `Out-GridView -PassThru`
- Nessun test, linter, o CI configurato

## Commit

Messaggi in inglese. Conventional Commits preferiti:

| Prefisso | Uso |
|---|---|
| `feat:` | Nuovo script o funzionalità |
| `refactor:` | Miglioramenti/refactoring |
| `fix:` | Correzione bug |
| *(nessun prefisso)* | Frasi descrittive semplici (es. `Add script to ...`, `Remove ...`) |
