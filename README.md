# Win32Forge

Upload, Automate & Document Win32 Apps in Intune

Win32Forge is a free, open source PowerShell 7 GUI tool for packaging, uploading, and documenting Win32 applications in Microsoft Intune. It is built around a JSON template system and has deep support for [PSAppDeployToolkit (PSADT) v4](https://psappdeploytoolkit.com/), removing the repetitive manual work from Intune app management.

> **No warranty.** This tool is provided free of charge and without any warranty of any kind. Use at your own risk.  
> Built and maintained by [modernworkspacehub.com](https://modernworkspacehub.com)

---

## Built on IntuneWin32App

Win32Forge uses the **[IntuneWin32App](https://github.com/MSEndpointMgr/IntuneWin32App)** PowerShell module by [MSEndpointMgr](https://msendpointmgr.com) as its backend for all Intune app creation, detection rules, requirement rules, and assignments. A huge thanks to the MSEndpointMgr team for building and maintaining that module — Win32Forge would not be possible without it.

---

## Features

### Single app upload

Upload one app at a time through a guided, tabbed form covering:

- **App metadata** — display name, version, publisher, description, owner, notes, information URL, privacy URL, and app category (categories are loaded live from your tenant)
- **Commands** — install and uninstall command lines, install context (system or user), device restart behaviour
- **Detection method** — choose from PowerShell script, registry key, MSI product code, or file/folder existence/version checks
- **Requirement rules** — target architecture (x64, x86, ARM64, or any combination), minimum Windows version, and optional additional requirement rules (script, registry, or file based)
- **Assignment** — All Devices, All Users, specific Entra ID group(s) with per-group intent and notification, or no assignment. Intune assignment filters (loaded from your tenant) can be applied to any assignment type
- **Logo** — attach a PNG or JPG app icon for the Company Portal tile

When you select a source folder, Win32Forge automatically scans it and pre-fills two fields if it finds them:

- **Detection script** — any `.ps1` file with "detection" in its name in the **root** of the source folder is auto-set as the PowerShell detection script (only if no detection method has been set yet)
- **Logo** — the first PNG, JPG, or JPEG found in the **root** of the source folder is auto-set as the app logo (only if no logo has been set yet)

A confirmation prompt is shown each time so you can review or override before uploading.

### Template system

Templates are JSON files in the `Templates\` folder. They define the deployment defaults for an app — install commands, assignment type, return codes, architecture, restart behaviour, and more. Select a template per app in both the single upload form and the bulk manager, and override it on a per-app basis as needed.

Templates are created and edited using the built-in **Template Editor** (the Templates button in the main window) — no manual JSON editing required. The templates included in this repo are examples to get you started; you are expected to build your own to match your environment.

### PSADT v4 support

When a template has `IsPSADT` enabled, Win32Forge scans the source folder for `Invoke-AppDeployToolkit.ps1` and extracts the app's metadata — display name, version, publisher, and author — directly from the script variables. The install and uninstall commands in a PSADT template are the **PSADT framework's own deployment commands** (`Invoke-AppDeployToolkit.exe -DeployMode Silent`), not commands specific to the underlying app installer. PSADT handles the actual install logic internally.

### Bulk upload manager

The bulk manager is a full app catalogue editor. Each row in the grid represents one app and exposes the same fields available in the single upload form — source folder, template, display name, version, publisher, setup file, install/uninstall commands, description, information URL, privacy URL, logo, detection method, and assignment. Using templates reduces the number of fields you need to fill per row; the template can be changed per row independently of the global default.

The same auto-detection applies here too: when a source folder is set, Win32Forge scans for a detection script and logo and pre-fills them if found.

Additional features:

- **Scan a folder** to auto-discover multiple app packages at once
- **Edit any row** in the full single-app form for detailed detection and assignment configuration
- **Import/export** the entire queue as JSON for repeatable deployments
- **Right-click context menu** for per-row actions (edit, delete, upload now)
- Uploads run sequentially with live status per row — errors are captured and displayed without stopping the rest of the queue

### Automatic documentation

After every successful upload, Win32Forge writes a Markdown document to your configured docs folder containing: app metadata, packaging details, install/uninstall commands, detection method (including script content if applicable), requirement rules, assignment details with filter information, return codes, the Intune app ID, and a direct link to the app in the Intune portal.

Example documentation files are included in the `Docs\` folder so you can see what the output looks like before uploading your first app.

### In-app settings

All configuration is managed through the **Settings** button in the main Win32Forge window — no need to edit `config.json` by hand. You can update your tenant ID, auth method, output and docs paths, default template, and verbose logging. The default template set here is used for all new uploads but can still be overridden per app in the single upload form, or per row in the bulk manager. IntuneWinAppUtil.exe can also be re-downloaded directly from the Settings window.

### Verbose logging

Optional structured log file capturing packaging operations, Graph API calls, upload details, and errors with stack traces — useful for troubleshooting in larger environments.

### Headless bulk mode

Run unattended batch uploads by passing a JSON file directly:

```powershell
pwsh .\Invoke-Win32Forge.ps1 -BulkFile "C:\apps\upload-queue.json"
```

---

## Prerequisites

| Requirement | Notes |
| --- | --- |
| **PowerShell 7** (`pwsh.exe`) | **Not** Windows PowerShell 5.1 |
| [IntuneWin32App module](https://github.com/MSEndpointMgr/IntuneWin32App) | Installed automatically by `Setup-Win32Forge.ps1` |
| [Microsoft.Graph.Authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph) | Installed automatically by `Setup-Win32Forge.ps1` |
| IntuneWinAppUtil.exe | Downloaded automatically by `Setup-Win32Forge.ps1` |
| Intune Administrator (or equivalent) permissions | Required to upload and assign apps |

---

## Quick Start

### 1. Clone or download

```powershell
git clone https://github.com/durrante/Win32Forge.git
cd Win32Forge
```

Or download the ZIP from the [Releases page](https://github.com/durrante/Win32Forge/releases) and extract it.

### 2. Run setup (once)

Open **PowerShell 7** (`pwsh.exe`) and run:

```powershell
pwsh .\Setup-Win32Forge.ps1
```

The setup script will install required modules, download IntuneWinAppUtil.exe, and walk you through creating `Config\config.json`.

### 3. Launch Win32Forge

```powershell
pwsh .\Invoke-Win32Forge.ps1
```

> **Important:** Always launch with `pwsh.exe` (PowerShell 7), not `powershell.exe` (Windows PowerShell 5.1). Modules installed during setup go to the PS7 module path and will not be found by PS5.1.

---

## Configuration

`Config\config.json` is created by `Setup-Win32Forge.ps1`. To configure manually, copy `Config\config.example.json` to `Config\config.json` and fill in your values.

| Field | Description |
| --- | --- |
| `AuthMethod` | `MicrosoftGraphCLI` or `CustomApp` |
| `TenantID` | Your Entra ID tenant ID |
| `ClientID` | Leave as default for Graph CLI; replace with your app registration client ID for CustomApp |
| `DefaultOutputPath` | Where `.intunewin` packages are saved |
| `DocumentationPath` | Where Markdown app docs are written |
| `IntuneWinAppUtilPath` | Full path to `IntuneWinAppUtil.exe` |
| `DefaultTemplate` | Template filename (without `.json`) used when no per-app template is set |
| `VerboseLogging` | `true` / `false` — enable structured log file |
| `LogPath` | Full path to the log file (required when `VerboseLogging` is `true`) |

### Authentication methods

**MicrosoftGraphCLI** (recommended)  
Uses the Microsoft Graph Command Line Tools public client app. No app registration required. Prompts for interactive browser login per session.

**CustomApp**  
Uses your own Entra ID app registration. Required delegated permissions:

| Permission | Purpose |
| --- | --- |
| `DeviceManagementApps.ReadWrite.All` | Upload and assign Win32 apps |
| `DeviceManagementConfiguration.Read.All` | Load Intune assignment filters (optional — filters won't load if missing) |
| `Group.Read.All` | Search and resolve Entra ID groups for assignments |

---

## Templates

Templates live in `Templates\` as JSON files and define the deployment defaults for an app. Select a template per upload; edit or create templates using the built-in Template Editor.

### Included templates

| Template | Assignment | Intent | PSADT |
| --- | --- | --- | --- |
| `PSADT-Required` | All Devices | Required | Yes |
| `PSADT-Available` | All Users | Available | Yes |
| `PSADT-Groups` | Specific Entra ID group (placeholder — edit before use) | Required | Yes |
| `Generic-Required` | All Devices | Required | No |
| `Generic-Available` | All Users | Available | No |

---

## PSADT Support

When `IsPSADT` is enabled on a template, Win32Forge scans the source folder for `Invoke-AppDeployToolkit.ps1` and reads the `$appVendor`, `$appName`, `$appVersion`, and author variables to auto-populate the Intune app metadata fields.

The install and uninstall commands in a PSADT-enabled template are the **PSADT toolkit commands**, not commands specific to the underlying installer:

```text
Install:   Invoke-AppDeployToolkit.exe -DeployMode Silent
Uninstall: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent
```

PSADT itself handles calling the actual installer — these commands are how Intune triggers the toolkit. The install logic for your specific application lives inside `Invoke-AppDeployToolkit.ps1`.

Source folders must be a valid PSAppDeployToolkit v4 structure with `Invoke-AppDeployToolkit.ps1` at the root.

---

## Folder Structure

```text
Win32Forge\
├── Invoke-Win32Forge.ps1       # Main entry point — launch this
├── Setup-Win32Forge.ps1        # One-time setup script
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

- Report bugs or suggest features via [GitHub Issues](https://github.com/durrante/Win32Forge/issues)
- All pull requests should target the `main` branch
- Keep changes focused — one feature or fix per PR

---

## License

MIT — see [LICENSE](LICENSE) for full terms.

---

## Disclaimer

Win32Forge is a free, community tool provided **without warranty of any kind**, express or implied. It is not affiliated with or endorsed by Microsoft. Use of this tool against your Intune tenant is entirely at your own risk. Always test in a non-production environment first.

Built with ❤️ at [modernworkspacehub.com](https://modernworkspacehub.com)
