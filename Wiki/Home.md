# Win32Forge Wiki

Upload, Automate & Document Win32 Apps in Intune

Win32Forge is a free, open source PowerShell 7 GUI tool for packaging, uploading, and documenting Win32 applications in Microsoft Intune. It is built around a template system and has deep support for PSAppDeployToolkit (PSADT) v4.

> Built and maintained at [modernworkspacehub.com](https://modernworkspacehub.com) — provided without warranty.

---

## Built on IntuneWin32App

Win32Forge uses the **[IntuneWin32App](https://github.com/MSEndpointMgr/IntuneWin32App)** PowerShell module by [MSEndpointMgr](https://msendpointmgr.com) as its backend for all Intune app creation, detection rules, requirement rules, and assignments. A huge thanks to the MSEndpointMgr team for building and maintaining that module.

---

## Wiki Pages

| Page | Description |
| --- | --- |
| [[Installation]] | Prerequisites, setup steps, first launch |
| [[Authentication]] | MicrosoftGraphCLI vs custom app registration |
| [[Configuration]] | All config.json fields explained |
| [[Templates]] | Template system — structure, fields, examples |
| [[PSADT Support]] | PSADT v4 package metadata auto-detection |
| [[Bulk Upload]] | Managing and uploading multiple apps in one run |
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

- **Templates** — JSON files defining deployment settings (assignment, commands, return codes, architecture). Select one per app; create and edit them in the built-in Template Editor.
- **PSADT support** — Point Win32Forge at a PSAppDeployToolkit v4 package and it reads `$appName`, `$appVersion`, `$appVendor`, and author from the script automatically. The install commands in a PSADT template are the PSADT framework's own commands — PSADT handles calling your actual installer internally.
- **Bulk manager** — A full catalogue editor: queue apps in a grid with per-row fields for every Intune property, scan a folder to discover packages, import/export JSON, and upload sequentially with live status.
- **Auto-documentation** — After each successful upload, Win32Forge writes a Markdown doc with all metadata, detection rules, assignments, the Intune App ID, and a direct portal link. Example documentation files are included in the `Docs\` folder for reference.
