# Configuration

Win32Forge reads its configuration from `Config\config.json`. This file is created by `Setup-Win32Forge.ps1`. To configure manually, copy `Config\config.example.json` to `Config\config.json` and edit the values.

> `config.json` is excluded from the git repository — it contains your Tenant ID and personal paths.

---

## Full reference

```json
{
  "AuthMethod": "MicrosoftGraphCLI",
  "TenantID": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "ClientID": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
  "DefaultOutputPath": "C:\\IntunePackages\\Output",
  "DocumentationPath": "C:\\IntunePackages\\Docs",
  "IntuneWinAppUtilPath": "C:\\Win32Forge\\Tools\\IntuneWinAppUtil.exe",
  "DefaultTemplate": "PSADT-Required",
  "VerboseLogging": false,
  "LogPath": "C:\\Win32Forge\\Logs\\win32forge.log"
}
```

| Field | Type | Description |
| --- | --- | --- |
| `AuthMethod` | string | `MicrosoftGraphCLI` or `CustomApp` — see [[Authentication]] |
| `TenantID` | string | Your Entra ID tenant ID |
| `ClientID` | string | Leave as default for Graph CLI; replace with your app registration client ID for CustomApp |
| `DefaultOutputPath` | string | Folder where `.intunewin` packages are saved after packaging |
| `DocumentationPath` | string | Folder where generated Markdown app docs are written |
| `IntuneWinAppUtilPath` | string | Full path to `IntuneWinAppUtil.exe` — downloaded by setup |
| `DefaultTemplate` | string | Template filename (without `.json`) applied to new uploads unless overridden per-app |
| `VerboseLogging` | bool | `true` to enable structured log file output |
| `LogPath` | string | Full path to the log file — required when `VerboseLogging` is `true` |

---

## Changing settings after setup

You can edit config values at any time through the **Settings** button in the main Win32Forge window, or by editing `Config\config.json` directly. Restart Win32Forge after manually editing the file.

---

## Verbose logging

When `VerboseLogging` is `true`, Win32Forge appends structured entries to `LogPath` for:

- App packaging (IntuneWinAppUtil arguments, exit codes, output)
- Graph API calls (method, URL, auth method used)
- App upload (name, version, detection type, assignment)
- Errors with stack traces

Log entries use the format:

```text
[2026-04-18 14:32:01] [INFO ] [Invoke-ProcessApp] Starting: 7-Zip 24.09
[2026-04-18 14:32:03] [DEBUG] [Invoke-TenantGraphRequest] Graph POST https://graph.microsoft.com/...
[2026-04-18 14:32:08] [INFO ] [Add-IntuneApplication] Upload complete: App ID = xxxxxxxx
```

The log file is appended to — it is never overwritten. New sessions are separated by a header line.
