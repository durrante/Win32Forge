# PSADT Support

Win32Forge has first-class support for [PSAppDeployToolkit (PSADT) v4](https://psappdeploytoolkit.com/). When a template has `"IsPSADT": true`, Win32Forge reads the PSADT package and auto-populates the app name and version without any manual input.

---

## What gets auto-detected

When you point Win32Forge at a PSADT v4 source folder, it locates `Invoke-AppDeployToolkit.ps1` and reads:

| Variable | Maps to |
| --- | --- |
| `$appName` | Intune **App name** |
| `$appVersion` | Intune **Version** |
| `$appVendor` | Intune **Publisher** (if present) |

The install and uninstall commands are set automatically from the template:
```
Invoke-AppDeployToolkit.exe -DeployMode Silent
Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent
```

---

## Expected folder structure

Your source folder must be a valid PSAppDeployToolkit v4 package:

```
MyApp_1.0\
├── Invoke-AppDeployToolkit.exe      ← PSADT v4 launcher
├── Invoke-AppDeployToolkit.ps1      ← main script (Win32Forge reads $appName etc. from here)
├── AppDeployToolkit\
│   └── ...
└── Files\
    └── setup.exe                    ← your installer
```

> PSADT v3 packages used `Deploy-Application.ps1` instead of `Invoke-AppDeployToolkit.ps1`. PSADT v3 is not currently supported.

---

## Upload flow with PSADT

1. In the upload form, browse to the root of the PSADT package folder
2. If the selected template has `IsPSADT: true`, Win32Forge scans for `Invoke-AppDeployToolkit.ps1`
3. App name and version fields are filled in automatically — you can still override them
4. Win32Forge calls `IntuneWinAppUtil.exe` with the package folder as the source, producing a `.intunewin` file
5. The `.intunewin` is uploaded to Intune using the populated metadata

---

## Detection rules with PSADT

Win32Forge does not generate detection rules automatically — you still set these in the upload form. Common approaches for PSADT-deployed apps:

- **Registry key** — check for an uninstall key under `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`
- **File/folder existence** — check for the main executable in `Program Files`
- **MSI product code** — use the MSI product code if the underlying installer is an MSI

---

## PSADT v4 resources

- [PSAppDeployToolkit documentation](https://psappdeploytoolkit.com/docs)
- [PSAppDeployToolkit GitHub](https://github.com/PSAppDeployToolkit/PSAppDeployToolkit)
- [PSADT guides on modernworkspacehub.com](https://modernworkspacehub.com)
