# PRA Cloud Mailbox

When Active Directory, Exchange and Entra Connect are all lost, gives every user and shared mailbox a mailbox in Exchange Online, and rolls everything back when the infrastructure is rebuilt. PowerShell 7.5+ engine, guided WPF console. Author: Nicolas Fabert. Version 2.1.0.

This folder contains everything needed to run the tool: `Invoke-PraCloudMailbox.ps1`, its modules and the WPF console, the configuration and samples, the report template, the SQLite library and the two guides (user guide and administrator guide). Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.
>
> The `Install-Module` commands in this documentation use `-Force`, so they also update or reinstall a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one (as administrator for `-Scope AllUsers`), run `Uninstall-Module <ModuleName> -AllVersions -Force`, then run the `Install-Module` command again.

## Requirements

- PowerShell 7.5 or later (`pwsh.exe`) for every action, Collect included, and for the console. Windows PowerShell 5.1 is not supported; the 1.x releases are the version that ran Collect in Windows PowerShell 5.1.
- Collect: a domain-joined Windows server that reaches an Exchange server with Kerberos (the remote Exchange PowerShell endpoint, `http://<server>/PowerShell/`) and an account that reads recipients and permissions. Set `Exchange.Server` in the configuration.
- Check, Convert, Recover: `ExchangeOnlineManagement` 3.10.1 or later and `Microsoft.Graph.Authentication` 2.41.1 or later, installed for PowerShell 7, and an app registration with a certificate (or an administrator signing in).
- Windows Terminal for emoji and colours.
- To read the HTML report: a current desktop Microsoft Edge or Google Chrome (the CSV with the same name works everywhere).

## Quick start

1. Check the download: compare the hash of the zip with its `.sha256` file (`(Get-FileHash .\PraCloudMailbox-2.1.0.zip -Algorithm SHA256).Hash`), extract it, and check that `Get-AuthenticodeSignature .\module\Test-PackageIntegrity.ps1` says `Valid`.
2. Unblock the files (box above) and install the two modules in an administrator PowerShell 7: `Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.41.1 -Scope AllUsers -Force`, then `Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.1 -Scope AllUsers -Force`.
3. Fill in the configuration: `notepad .\config\PraCloudMailbox.config.psd1` (`Exchange.Server`, tenant, app registration or interactive sign-in, licences, case hold, Entra Connect).
4. Schedule the daily Collect with `pwsh.exe`, before the disaster. In an administrator PowerShell 7 on the Collect computer, one command checks the account (a domain user that Exchange accepts) and registers the task:

   ```powershell
   .\Invoke-PraCloudMailbox.ps1 -Action Schedule                  # Preview: checks, registers nothing
   .\Invoke-PraCloudMailbox.ps1 -Action Schedule -Mode Apply      # asks for the account and its password, then registers the task
   ```

   The task runs `pwsh.exe -NoProfile -NonInteractive -File "<full path>\Invoke-PraCloudMailbox.ps1" -Action Collect -Mode Apply -Force` (an absolute path): it writes a snapshot and its `.sha256` file.
5. Protect the snapshot (below), then copy it with its `.sha256` file to the cloud admin server.
6. Check, then the disaster, then the rollback:

   ```powershell
   pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check                         # is every object ready? (nothing is written)
   pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply           # give every object its cloud mailbox
   pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch 1d5ac20b   # roll a batch back
   pwsh -STA -File .\Invoke-PraCloudMailbox.ps1 -Gui                             # the guided console
   ```

`Check -Batch <id>` follows every object of a batch, also the ones Convert never started, and never says "provisioned" while the batch is not Complete. An object found in Microsoft Entra ID by its UPN only, with another mailbox GUID, is refused (`ENTRA_WRONG_OBJECT`). A run interrupted by Ctrl+C ends with `RESULT: FAIL`, and a Convert is not resumed once a Recover rolled it back. Recover uses the licence mode of the batch (`Licensing.Users.Mode`).

Every command can be narrowed: `-Identity`, `-CsvPath`, `-SearchBase`, `-Group`, `-Scope`, `-MaxObjects`. Without `-Mode Apply` nothing is changed.

What Convert and Recover handle: user mailboxes, shared mailboxes and, if you turn it on, native rooms and equipment. They do **not** recreate contacts, distribution groups, mail-enabled security groups, dynamic distribution groups or Microsoft 365 groups. Collect can keep contacts, distribution groups and dynamic distribution groups in the snapshot, for reference only; Microsoft 365 groups are not collected. The members of the groups you name are used only to choose the objects of a run with `-Group`. Calendar settings are restored only for the native rooms and equipment of the `CloudResource` mode.

## Protect the snapshot

After the disaster the snapshot file is the **only copy** of the directory data. Back it up every day and **externalise the backup**: off-site, out of reach of the domain that can be compromised (offline, immutable, or separate credentials), in several generations. Restrict the access list of the file and of its backups (no share open to *Domain Users*). Copy the `.sha256` file with the snapshot: Check, Convert and Recover verify the file against it, and a **changed** file stops the run: restore it from the external backup (`-AcceptUnverifiedSnapshot` only when you know why). Without a `.sha256` file a Convert or a Recover in Apply mode stops too. A batch stays tied to its snapshot: Recover and `Check -Batch` stop when it is missing, and Collect never deletes the snapshot of a batch that is not rolled back. A Collect that lost its mailboxes (`Collect.MaxDropPercent`) is refused, so the good snapshot stays. The journal, `data\PraCloudMailbox-journal.db`, is not covered by the hash because Convert and Recover write it: back it up during the operations.

## Package integrity

The tool runs only from a package signed by its owner. `module\Test-PackageIntegrity.ps1` and `PackageIntegrity.psd1` must carry a valid, timestamped signature of `0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6`; every PowerShell file must be signed by the same owner; every stable file (modules, XAML, template, sample files, SQLite library, README, licence, change log, guides, images) must match the signed manifest. The entry script and each module check this first. A changed, missing, duplicated, unexpected or reparse-point file stops the tool before any import, window or operation. There is no unsigned mode, switch or environment bypass; do not disable your execution policy or this check to clear an error.

Edit only `config\PraCloudMailbox.config.psd1` and your own CSV lists in `config\`: they are data, outside the check, like the generated `data`, `logs` and `reports` folders. A signature detects a change; it does not stop the file system from being written.

Output folders: `Store.Path`, `Store.BackupFolder`, `Store.JournalPath`, `Logging.Folder` and `Report.Folder` must be either a folder outside the package or, inside the package, the canonical `data`, `logs` and `reports` folder (a sub-folder of its own canonical folder is accepted). Any other folder inside the package is refused with a clear message before a file is written, because the integrity check would reject the unexpected files at the next start. Keep the script of `EntraConnect.ScriptPath` outside the package.

## Content

| Item | Role |
|---|---|
| `Invoke-PraCloudMailbox.ps1` | Entry script. |
| `module\` | Modules, WPF console (`PRA.Gui*.ps1`, `PRA.Gui.xaml`) and the integrity verifier. |
| `config\` | Configuration to fill in, the CSV sample and the Entra Connect script sample. |
| `templates\` | HTML report templates: the report of a run and the global report of a batch. |
| `lib\sqlite\` | System.Data.SQLite (SQLite, public domain) for PowerShell 7. |
| `docs\` | User guide and administrator guide, Markdown and self-contained HTML, with images. |
| `PraCloudMailbox.psd1`, `PraCloudMailbox.psm1` | Module manifest and root module (loaded by the script). |
| `PackageIntegrity.psd1` | Owner-signed manifest of the stable files. |
| `README.md`, `LICENSE`, `THIRD-PARTY-NOTICES.md` | This readme, the MIT licence, the notices of the bundled library. |

## Documentation

- [User guide](docs/PraCloudMailbox-UserGuide.md) - the guided path in plain words, every task in the console and on the command line; also `docs/PraCloudMailbox-UserGuide.html`, a single file to open locally
- [Administrator guide](docs/PraCloudMailbox-Guide.md) - everything else, including the exact minimum permissions; also `docs/PraCloudMailbox-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/PraCloudMailbox

License: [MIT](LICENSE).

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

The tool changes production objects in Entra ID and Exchange Online: run the preview, review the plan, keep the journal and test the whole cycle in a lab before using it in a real disaster.