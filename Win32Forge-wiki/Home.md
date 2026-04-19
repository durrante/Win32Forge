# Win32Forge Wiki

**Upload, Automate & Document Win32 Apps in Intune**

Win32Forge is a free, open source PowerShell 7 GUI tool for packaging, uploading, and documenting Win32 applications in Microsoft Intune. It is built around a template system and has deep support for PSAppDeployToolkit (PSADT) v4.

> Built and maintained at [modernworkspacehub.com](https://modernworkspacehub.com) — provided without warranty.

---

## Wiki Pages

| Page | Description |
| --- | --- |
| [[Installation]] | Prerequisites, setup steps, first launch |
| [[Authentication]] | MicrosoftGraphCLI vs custom app registration |
| [[Configuration]] | All config.json fields explained |
| [[Templates]] | Template system — structure, fields, examples |
| [[PSADT Support]] | PSADT v4 package auto-detection |
| [[Bulk Upload]] | Uploading multiple apps in one run |
| [[Troubleshooting]] | Common issues and fixes |

---

## Quick Start

```powershell
# 1. Clone the repo
git clone https://github.com/durrante/Win32Forge.git
cd Win32Forge

# 2. Run setup (PowerShell 7 only)
pwsh .\Setup-Win32Forge.ps1

# 3. Launch
pwsh .\Invoke-Win32Forge.ps1
```

See [[Installation]] for full details.

---

## Key Concepts

- **Templates** — JSON files that define how an app is deployed (assignment type, install commands, return codes). Pick one per app upload, or set a default in config.
- **PSADT support** — Point Win32Forge at a PSAppDeployToolkit v4 package folder and it reads the app name, version, and commands automatically.
- **Bulk manager** — Queue a grid of apps with individual templates and logos, then upload them all in one run.
- **Auto-documentation** — After each successful upload, Win32Forge writes a Markdown doc with all app metadata, detection rules, assignments, and return codes.
