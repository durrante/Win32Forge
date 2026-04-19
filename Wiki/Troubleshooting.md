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

## Still stuck?

Open an issue on [GitHub](https://github.com/durrante/Win32Forge/issues) and include:
- The error message
- PowerShell version (`$PSVersionTable.PSVersion`)
- IntuneWin32App module version (`Get-Module IntuneWin32App -ListAvailable | Select Version`)
- Relevant lines from the verbose log (with any sensitive IDs redacted)
