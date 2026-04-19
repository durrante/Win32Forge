# Installation

## Prerequisites

| Requirement | Notes |
| --- | --- |
| **PowerShell 7** (`pwsh.exe`) | **Not** Windows PowerShell 5.1. [Download here](https://github.com/PowerShell/PowerShell/releases) |
| **Internet access** | Required during setup to download modules and IntuneWinAppUtil.exe |
| **Intune Administrator** (or equivalent) | Required to upload and assign apps |

> The setup script installs all PowerShell module dependencies automatically. You do not need to install them manually.

---

## Step 1 — Clone or download

```powershell
git clone https://github.com/durrante/Win32Forge.git
```

Or download the ZIP from the [Releases page](https://github.com/durrante/Win32Forge/releases) and extract it.

---

## Step 2 — Run setup

Open **PowerShell 7** (`pwsh.exe`) and run:

```powershell
cd Win32Forge
pwsh .\Setup-Win32Forge.ps1
```

> **Important:** Run `pwsh.exe`, not `powershell.exe`. Modules installed by PS7 go to the PS7 module path and will not be found if you launch the tool with Windows PowerShell 5.1.

The setup script will guide you through:

1. **Module installation** — installs `IntuneWin32App` and `Microsoft.Graph.Authentication` into the PS7 module path
2. **IntuneWinAppUtil.exe download** — downloads Microsoft's Win32 Content Prep Tool and places it in `Tools\`
3. **Authentication method** — choose `MicrosoftGraphCLI` (recommended, no app registration needed) or `CustomApp` (your own Entra ID app)
4. **Tenant configuration** — enter your Entra ID Tenant ID
5. **Folder paths** — output folder for `.intunewin` packages, documentation folder for generated Markdown docs
6. **Default template** — the template applied to new uploads unless overridden per-app
7. **Authentication test** — launches a browser login to confirm everything works

Configuration is saved to `Config\config.json`.

---

## Step 3 — Launch Win32Forge

```powershell
pwsh .\Invoke-Win32Forge.ps1
```

Win32Forge loads the config, connects to your tenant, and opens the main window.

---

## Updating

To update to a newer release:

1. Pull the latest code (or download and extract the new release ZIP)
2. Re-run `pwsh .\Setup-Win32Forge.ps1` — it will update modules and re-download IntuneWinAppUtil.exe if needed
3. Your existing `Config\config.json` is preserved

---

## Folder structure after setup

```
Win32Forge\
├── Invoke-Win32Forge.ps1
├── Setup-Win32Forge.ps1
├── Assets\
│   └── logo.png              ← add your own logo here (optional)
├── Config\
│   ├── config.example.json
│   └── config.json           ← created by setup
├── Docs\                     ← generated app docs written here
├── Private\                  ← internal functions (do not edit)
├── Templates\                ← your deployment templates
└── Tools\
    └── IntuneWinAppUtil.exe  ← downloaded by setup
```
