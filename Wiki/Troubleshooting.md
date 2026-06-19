# Troubleshooting

---

## The tool won't launch / "module not found" errors

**Cause:** Win32Forge was launched with Windows PowerShell 5.1 (`powershell.exe`) instead of PowerShell 7 (`pwsh.exe`). Modules installed by setup go to the PS7 module path, which PS5.1 cannot see.

**Fix:** Always launch with:
```powershell
pwsh .\Invoke-Win32Forge.ps1
```

If you are running from the Windows Run dialog or a shortcut, make sure it points to `pwsh.exe`, not `powershell.exe`.

---

## Setup-Win32Forge.ps1 fails to install modules

**Cause:** PowerShell module installation requires Administrator rights, or the PSGallery is untrusted.

**Fix:** Run PowerShell 7 as Administrator, then run setup again. If prompted about an untrusted repository, type `Y` to accept.

---

## "Graph API call failed" / authentication errors

**Possible causes and fixes:**

1. **Token expired** — close and re-launch Win32Forge. A new browser login prompt will appear.
2. **Wrong tenant ID** — open Settings and verify `TenantID` matches your Entra ID tenant.
3. **Missing API permissions (CustomApp)** — go to Entra ID → App registrations → your app → API permissions and confirm all three required permissions are granted with admin consent. See [[Authentication]].
4. **Conditional Access blocking the login** — check if your tenant has CA policies that block the Microsoft Graph Command Line Tools app. If so, switch to `CustomApp` auth with your own app registration.

---

## Assignment filters don't appear

**Cause:** The account used to sign in does not have the `DeviceManagementConfiguration.Read.All` permission.

**Fix:** Add `DeviceManagementConfiguration.Read.All` as a delegated permission on your app registration (CustomApp), or ensure your account has Intune Administrator role. Filters are optional — all other features work without them.

---

## App name / version not auto-detected for PSADT packages

**Cause:** Win32Forge looks for `Invoke-AppDeployToolkit.ps1` at the root of the source folder (PSADT v4). If your package uses the older `Deploy-Application.ps1` (PSADT v3), auto-detection will not work.

**Fix:** Migrate your packages to PSADT v4, or manually enter the app name and version in the upload form.

---

## Logo is rejected / "unsupported format"

Win32Forge only accepts **PNG, JPG, and JPEG** logos. ICO, BMP, and other formats are not supported by Intune's Win32 app logo API.

Convert your logo to PNG before using it. A 512×512 or 300×300 PNG works well.

---

## IntuneWinAppUtil.exe not found

**Fix:** Re-run `Setup-Win32Forge.ps1` — it will download the tool and place it at the path configured in `IntuneWinAppUtilPath`. Alternatively, download it manually from [Microsoft's GitHub](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) and place it at the path in your config.

---

## Upload succeeds but app doesn't appear in Intune

**Cause:** Intune app sync can take a few minutes. Refresh the Apps list in the Intune portal after 2–3 minutes.

If the app still doesn't appear, check the verbose log (enable `VerboseLogging` in Settings) for Graph API error responses during the upload.

---

## Enable verbose logging for diagnostics

In the Settings window, tick **Verbose Logging** and set a log file path, then reproduce the issue. The log will contain detailed Graph API calls, responses, and any error stack traces.

See [[Configuration]] for full logging details.

---

## Windows 11 24H2 (or 23H2) rejected as minimum OS

**Symptom:** `Cannot validate argument on parameter 'MinimumSupportedWindowsRelease'`.

**Status:** Fixed in 1.1.0. Win32Forge patches the installed IntuneWin32App module on launch to add `W11_23H2` and `W11_24H2`. If you still see this, fully close and re-launch Win32Forge so the patched module is reloaded.

---

## A category isn't applied / "ModelValidationFailure ... property named 'id'"

**Cause:** The category name didn't resolve to a category that exists in your tenant. In earlier builds, category names containing a space (e.g. *Data management*) also failed to resolve and could fail the whole upload.

**Status:** Fixed in 1.1.0 — category resolution is now done locally and reliably, and an unknown category is skipped with a warning rather than failing the upload. To apply a category, make sure its name **exactly matches** an existing Intune app category (Intune → Apps → Categories). Re-launch Win32Forge after upgrading so the module patch is active.

---

## Large app upload floods "SAS Uri renewal ... segment 'deviceAppManagement'" warnings

**Cause:** A bug in the module's mid-upload SAS-token renewal that triggers on large packages.

**Status:** Fixed in 1.1.0 (module patch applied on launch). If you see it, re-launch Win32Forge so the patch is active.

---

## Packaging fails: "Could not find a part of the path" on a deep folder

**Cause:** `IntuneWinAppUtil.exe` is limited to 260-character paths (`MAX_PATH`). Deeply-nested packages on a long source path (e.g. under OneDrive) can exceed it.

**Status:** Win32Forge 1.1.0 automatically retries packaging through a short directory junction, which resolves almost all cases. If a package is so deep that it *still* exceeds 260 chars from a short root, either enable Windows long-path support (`LongPathsEnabled`) or move the source closer to the drive root.

---

## Upload error mentions a date like "13/04/2026 was not recognized as a valid DateTime"

**Cause:** A locale (e.g. en-GB) date-parsing bug in the module on non-US systems.

**Status:** Fixed in 1.1.0 (module patch applied on launch).

---

## About the automatic module patches

On every launch, Win32Forge applies small, idempotent compatibility patches to your locally-installed IntuneWin32App module (the four issues above). They are safe to re-run and don't modify the upstream gallery package. **After upgrading Win32Forge, fully re-launch it** so the patches are (re)applied and the module is reloaded.

---

## Still stuck?

Open an issue on [GitHub](https://github.com/durrante/Win32Forge/issues) and include:
- The error message
- PowerShell version (`$PSVersionTable.PSVersion`)
- IntuneWin32App module version (`Get-Module IntuneWin32App -ListAvailable | Select Version`)
- Relevant lines from the verbose log (with any sensitive IDs redacted)
