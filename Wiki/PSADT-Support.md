# PSADT Support

Win32Forge has first-class support for [PSAppDeployToolkit (PSADT) v4](https://psappdeploytoolkit.com/). When a template has `"IsPSADT": true`, Win32Forge reads the PSADT package to auto-populate the Intune app metadata fields — saving you from typing the same details that already exist in the script.

---

## What "PSADT support" actually means

It is important to understand what Win32Forge does and does not do with a PSADT package:

**What it does:** reads metadata from `Invoke-AppDeployToolkit.ps1` to auto-fill the Intune app record.

**What it does not do:** scan your PSADT script to find the install commands for the underlying application.

The install and uninstall commands in a PSADT-enabled template are the **PSADT framework's own deployment commands**:

```text
Install:   Invoke-AppDeployToolkit.exe -DeployMode Silent
Uninstall: Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent
```

These are how Intune triggers the PSADT toolkit in silent mode. PSADT itself is responsible for calling your actual application installer — that logic lives inside `Invoke-AppDeployToolkit.ps1`. Win32Forge does not read or modify it.

---

## What gets auto-detected

Win32Forge locates `Invoke-AppDeployToolkit.ps1` in the source folder and reads these variables via regex:

| Variable in script | Maps to Intune field |
| --- | --- |
| `$appVendor` | Publisher |
| `$appName` | Display Name |
| `$appVersion` | Version |
| `$appScriptAuthor` / `$appOwner` | Owner |

Author/owner is detected from several patterns including `# Author:`, `# Owner:`, `# Created by:`, and `.NOTES` blocks.

---

## Expected folder structure

Your source folder must be a valid PSAppDeployToolkit v4 package:

```text
MyApp_1.0\
├── Invoke-AppDeployToolkit.exe      ← PSADT v4 launcher (used as setup file)
├── Invoke-AppDeployToolkit.ps1      ← main script (Win32Forge reads metadata from here)
├── AppDeployToolkit\
│   └── ...
└── Files\
    └── setup.exe                    ← your actual installer (called by the script)
```

> PSADT v3 packages use `Deploy-Application.ps1` — Win32Forge can detect v3 packages but PSADT v4 is recommended.

---

## Upload flow with PSADT

1. Browse to the root of the PSADT package folder in the upload form
2. Win32Forge detects `Invoke-AppDeployToolkit.ps1` and switches to PSADT mode
3. Display name, version, publisher, and author fields are auto-filled — you can still override them
4. Install and uninstall commands are set from the template (read-only in PSADT mode)
5. Win32Forge packages the entire folder with IntuneWinAppUtil.exe
6. The `.intunewin` is uploaded to Intune with the populated metadata

---

## Detection rules with PSADT

Win32Forge does not generate detection rules automatically for PSADT packages — you still configure these in the upload form. Common approaches:

- **Registry key** — check for an uninstall key under `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`
- **File/folder existence** — check for the main executable in `Program Files`
- **MSI product code** — use the MSI product code if the underlying installer is an MSI

---

## PSADT v4 resources

- [PSAppDeployToolkit documentation](https://psappdeploytoolkit.com/docs)
- [PSAppDeployToolkit GitHub](https://github.com/PSAppDeployToolkit/PSAppDeployToolkit)
- [PSADT guides on modernworkspacehub.com](https://modernworkspacehub.com)
