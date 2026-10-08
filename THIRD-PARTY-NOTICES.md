# Third-party notices

The folder `lib\sqlite` contains unmodified binaries of **System.Data.SQLite 1.0.119.0** (SQLite 3.46.1), taken from the
nuget.org packages `Stub.System.Data.SQLite.Core.NetFramework` 1.0.119 (`lib\net46`, `build\net46\x64`) and
`Stub.System.Data.SQLite.Core.NetStandard` 1.0.119 (`lib\netstandard2.0`, `runtimes\win-x64\native`), checked by SHA-256.

| File | Used by | SHA-256 |
|---|---|---|
| `net46\System.Data.SQLite.dll` | Windows PowerShell 5.1 (Collect) | `77DAC4E1AA63161D6AFF363030A801945D6B8D5B823A6B460B2C2BE9EC38D3B7` |
| `net46\x64\SQLite.Interop.dll` | Windows PowerShell 5.1 (Collect) | `1111916DC329A13BD627B2CD90C9B2263DE9923FD0BB6059C69C52332F360C37` |
| `netstandard2.0\System.Data.SQLite.dll` | PowerShell 7 (Check, Convert, Recover) | `BA07F6DABA18C815A7506E13031E19F2A6025858DCA0D16FF71F98CF65A95208` |
| `netstandard2.0\x64\SQLite.Interop.dll` | PowerShell 7 (Check, Convert, Recover) | `1111916DC329A13BD627B2CD90C9B2263DE9923FD0BB6059C69C52332F360C37` |

**License**: System.Data.SQLite and SQLite are in the **public domain** — https://system.data.sqlite.org and
https://www.sqlite.org/copyright.html.

The project also uses the PowerShell modules `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement`
(installed by the administrator from the PowerShell Gallery, not bundled), the Exchange Management Shell of the
Exchange servers, the ADSync module of the Entra Connect server, and Microsoft Edge (already installed with Windows)
only to render the images of the guides.