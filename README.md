# Win32Forge

**Upload, Automate & Document Win32 Apps in Intune**

Win32Forge is a free, open source PowerShell 7 GUI tool for packaging, uploading, and documenting Win32 applications in Microsoft Intune. Built around templates and PSADT (PSAppDeployToolkit) v4, it removes the repetitive manual work from Intune app management.

> **No warranty.** This tool is provided free of charge and without any warranty of any kind. Use at your own risk.  
> Built and maintained by [modernworkspacehub.com](https://modernworkspacehub.com)

---

## Features

- **One-click packaging & upload** — wraps IntuneWinAppUtil.exe and the IntuneWin32App PowerShell module into a guided GUI workflow
- **Template system** — define install commands, assignments, return codes, and deployment settings once in a JSON template; reuse across every app
- **PSADT v4 support** — auto-detects app name, version, and commands from a PSAppDeployToolkit package; enforces silent deployment mode
- **Bulk upload manager** — queue multiple apps in a grid, assign individual templates and logos, upload them all in one run
- **Automatic documentation** — generates a Markdown doc for each uploaded app with metadata, detection rules, assignments, and return codes
- **Settings wizard** — configure tenant, authentication method, paths, default template, and verbose logging from within the GUI
- **Verbose logging** — optional structured log file for troubleshooting packaging and Graph API operations
- **Both auth methods** — Microsoft Graph CLI (no app registration needed) or a custom Entra ID app registration

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 7 (`pwsh.exe`) | **Not** Windows PowerShell 5.1 |
| [IntuneWin32App module](https://github.com/MSEndpointMgr/IntuneWin32App) | Installed by `Setup-Win32Forge.ps1` |
| [Microsoft.Graph.Authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph) | Installed by `Setup-Win32Forge.ps1` |
| IntuneWinAppUtil.exe | Downloaded by `Setup-Win32Forge.ps1` |
| Intune Administrator (or equivalent) permissions in your tenant | Required for app upload |

---

## Quick Start

### 1. Clone or download

```powershell
git clone https://github.com/your-username/Win32Forge.git
cd Win32Forge
```

### 2. Run setup (once)

Open **PowerShell 7** (`pwsh.exe`) and run:

```powershell
.\Setup-Win32Forge.ps1
```

The setup script will:
- Install required PowerShell modules (PS7 module path)
- Download `IntuneWinAppUtil.exe` from Microsoft
- Walk you through creating `Config\config.json`
- Test authentication against your tenant

### 3. Launch Win32Forge

```powershell
.\Invoke-Win32Forge.ps1
```

> **Important:** Always launch with `pwsh.exe` (PowerShell 7), not `powershell.exe` (Windows PowerShell 5.1). Modules installed during setup are placed in the PS7 module path and will not be found by PS5.1.

---

## Configuration

`Config\config.json` is created by `Setup-Win32Forge.ps1`. To configure manually, copy `Config\config.example.json` to `Config\config.json` and fill in your values.

| Field | Description |
|---|---|
| `AuthMethod` | `MicrosoftGraphCLI` or `CustomApp` |
| `TenantID` | Your Entra ID tenant ID |
| `ClientID` | Leave as default for Graph CLI; replace with your app registration client ID for CustomApp |
| `DefaultOutputPath` | Where `.intunewin` packages are saved |
| `DocumentationPath` | Where Markdown app docs are written |
| `IntuneWinAppUtilPath` | Path to `IntuneWinAppUtil.exe` (downloaded by setup) |
| `DefaultTemplate` | Template name (without `.json`) to apply when no per-app template is set |
| `VerboseLogging` | `true` / `false` — enable structured log file |
| `LogPath` | Full path to the log file (required when `VerboseLogging` is `true`) |

### Authentication methods

**MicrosoftGraphCLI** (recommended for most users)  
Uses the Microsoft Graph Command Line Tools public client application. No app registration required. Prompts for interactive browser login on first use.

**CustomApp**  
Uses your own Entra ID app registration. Requires the following delegated permissions:
- `DeviceManagementApps.ReadWrite.All`
- `Group.Read.All` (for group-based assignments)

---

## Templates

Templates live in the `Templates\` folder as JSON files. Each template defines deployment settings that are applied when an app is uploaded. The `DefaultTemplate` in `config.json` is used unless overridden per-app in the GUI.

### Included templates

| Template | Description |
|---|---|
| `PSADT-Required.json` | PSADT v4 app, deployed as **Required** to **All Devices** |
| `PSADT-Available.json` | PSADT v4 app, published as **Available** in Company Portal for **All Users** |
| `PSADT-Groups.json` | PSADT v4 app, deployed as **Required** to a specific **Azure AD group** (edit group name/ID before use) |
| `Generic-Required.json` | Standard Win32 app (no PSADT wrapper), deployed as **Required** to **All Devices** |
| `Generic-Available.json` | Standard Win32 app (no PSADT wrapper), published as **Available** for **All Users** |

### Template structure

```json
{
  "TemplateName": "PSADT-Required",
  "IsPSADT": true,
  "InstallCommandLine": "Invoke-AppDeployToolkit.exe -DeployMode Silent",
  "UninstallCommandLine": "Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent",
  "InstallExperience": "system",
  "RestartBehavior": "suppress",
  "Architecture": "x64",
  "MinimumSupportedWindowsRelease": "W10_2004",
  "MaximumInstallationTimeInMinutes": 60,
  "AllowAvailableUninstall": false,
  "ReturnCodes": [ ... ],
  "Assignment": {
    "Type": "AllDevices",
    "Intent": "required",
    "Notification": "showAll"
  }
}
```

For group-based assignments, set `"Type": "Group"` and provide a `Groups` array:

```json
"Assignment": {
  "Type": "Group",
  "Groups": [
    {
      "GroupName": "YOUR-GROUP-NAME",
      "GroupID": "YOUR-GROUP-OBJECT-ID",
      "Intent": "required",
      "Notification": "showAll",
      "FilterName": "",
      "FilterID": "",
      "FilterIntent": "include"
    }
  ]
}
```

---

## PSADT Support

When `IsPSADT` is `true` in the template, Win32Forge will:

1. Locate the `Invoke-AppDeployToolkit.ps1` script inside the source folder
2. Read the `$appName` and `$appVersion` variables to auto-populate the Intune app name and version
3. Set install/uninstall commands to the PSADT silent deployment convention

Source folders must be a valid PSAppDeployToolkit v4 package structure with `Invoke-AppDeployToolkit.ps1` at the root.

---

## Folder Structure

```
Win32Forge\
├── Invoke-Win32Forge.ps1       # Main entry point — launch this
├── Setup-Win32Forge.ps1        # One-time setup script
├── Assets\
│   └── logo.png                # Tool logo (add your own — not included in repo)
├── Config\
│   ├── config.example.json     # Example configuration — copy to config.json
│   └── config.json             # Your configuration (not in repo — created by setup)
├── Docs\                       # Generated app documentation (not in repo)
├── Private\                    # Internal PowerShell functions
├── Templates\                  # JSON deployment templates
└── Tools\
    └── IntuneWinAppUtil.exe    # Downloaded by setup (not in repo)
```

---

## Contributing

Contributions are welcome. Please open an issue or pull request on GitHub.

- Report bugs or suggest features via [GitHub Issues](../../issues)
- All pull requests should target the `main` branch
- Keep changes focused — one feature or fix per PR

---

## License

MIT — see [LICENSE](LICENSE) for full terms.

---

## Disclaimer

Win32Forge is a free, community tool provided **without warranty of any kind**, express or implied. It is not affiliated with or endorsed by Microsoft. Use of this tool against your Intune tenant is entirely at your own risk. Always test in a non-production environment first.

Built with ❤️ at [modernworkspacehub.com](https://modernworkspacehub.com)
