# Google Chrome

| Field | Value |
|-------|-------|
| Display Name | Google Chrome |
| Version | Latest |
| Publisher | Google |
| Description | Google Chrome is a free, fast, and secure web browser developed by Google, launched in 2008. It is the world's most popular browser, used to access websites and web applications across Windows, macOS, Linux, Android, and iOS. It is known for its minimalist design, high-speed performance, and integration with Google services. |
| Author | Alex Durrant |
| Notes | PSADT v4 package (Chrome) |
| Categories | Productivity |
| Install Context | System |
| Information URL | [https://www.google.com/intl/en_uk/chrome/?_gl=1*1l44kpk*_up*MQ..*_ga*MjExNzgwMTQ5MS4xNzc2NTE2OTc3*_ga_B7W0ZKZYDK*czE3NzY1MTY5NzckbzEkZzAkdDE3NzY1MTY5NzckajYwJGwwJGgw&gclid=CjwKCAjw14zPBhAuEiwAP3-Ebx8FMcwU3RWh12XPPCq5D0JLl5tERoN9E1BqKLsxO-ggjLs3CqNNpRoCMRwQAvD_BwE&gclsrc=aw.ds&gbraid=0AAAAAoY3CA6EoaYkvMvHfPw0QqIzUTz6w](https://www.google.com/intl/en_uk/chrome/?_gl=1*1l44kpk*_up*MQ..*_ga*MjExNzgwMTQ5MS4xNzc2NTE2OTc3*_ga_B7W0ZKZYDK*czE3NzY1MTY5NzckbzEkZzAkdDE3NzY1MTY5NzckajYwJGwwJGgw&gclid=CjwKCAjw14zPBhAuEiwAP3-Ebx8FMcwU3RWh12XPPCq5D0JLl5tERoN9E1BqKLsxO-ggjLs3CqNNpRoCMRwQAvD_BwE&gclsrc=aw.ds&gbraid=0AAAAAoY3CA6EoaYkvMvHfPw0QqIzUTz6w) |
| Privacy URL | [https://policies.google.com/privacy?hl=en&_gl=1*n6l92w*_up*MQ..*_ga*NjAyMTE0NzU3LjE3NzY1MTY5NTY.*_ga_B7W0ZKZYDK*czE3NzY1MTY5NTUkbzEkZzAkdDE3NzY1MTY5NTYkajU5JGwwJGgw](https://policies.google.com/privacy?hl=en&_gl=1*n6l92w*_up*MQ..*_ga*NjAyMTE0NzU3LjE3NzY1MTY5NTY.*_ga_B7W0ZKZYDK*czE3NzY1MTY5NTUkbzEkZzAkdDE3NzY1MTY5NTYkajU5JGwwJGgw) |
| Intune App ID | `740b8852-825c-44ad-889c-3e2b4ef7e3bc` |
| Uploaded | 2026-04-18 13:56:43 |
| Template | PSADT-Autopilot |

[View in Intune Portal](https://intune.microsoft.com/#blade/Microsoft_Intune_Apps/SettingsMenu/0/appId/740b8852-825c-44ad-889c-3e2b4ef7e3bc)

---

## Packaging

| Field | Value |
|-------|-------|
| Source Folder | `D:\OneDrive\OneDrive - Alexdu\Documents\GitHub\PSADT\v4\GoogleChrome_Evergreen_PSADT` |
| Setup File | `Invoke-AppDeployToolkit.exe` |
| .intunewin | `C:\Users\me\Downloads\IntunePackage\Google_Chrome_Latest_PSADT.intunewin`  (9.57 MB) |
| Logo | `C:\Users\me\Downloads\IntuneDocs\Google Chrome_Logo.png` |

---

## Commands

> **PSADT Package** (v4)
> Install and uninstall commands use the PSAppDeployToolkit framework.
> Silent mode is enforced; the toolkit handles all UI suppression and logging.

| Command | Value |
|---------|-------|
| Install | `Invoke-AppDeployToolkit.exe -DeployMode Auto` |
| Uninstall | `Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Auto` |

---

## Detection Method

**PowerShell Script**: `GoogleChrome_DetectionScript.ps1`  
- Enforce signature check: False  
- Run as 32-bit: False

**Script Content:**

```powershell
<#
.SYNOPSIS
    Customised Win32App Detection Script
.DESCRIPTION
    This script identifies if a specific software, defined by its display name, is installed on the system.
    It checks the uninstall keys in the registry and reports back.
.EXAMPLE
    $TargetSoftware = 'Firefox'  # Searches for an uninstall key with the display name 'Firefox'
#>

# Define the name of the software to search for
$TargetSoftware = 'Google Chrome'

# Function to fetch uninstall keys from the registry
function Fetch-UninstallKeys {
    [CmdletBinding()]
    param (
        [string]$TargetName
    )

    # Continue on error
    $ErrorActionPreference = 'Continue'

    # Define uninstall registry paths
    $registryPaths = @(
        "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "registry::HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    # Initialize software list
    $softwareList = @()

    # Loop through each registry path to find software
    foreach ($path in $registryPaths) {
        $softwareList += Get-ChildItem $path | Get-ItemProperty | Where-Object { $_.DisplayName } | Sort-Object DisplayName
    }

    # Filter software list based on target name
    if ($TargetName) {
        $softwareList | Where-Object { $_.DisplayName -like "*$TargetName*" }
    } else {
        $softwareList | Sort-Object DisplayName -Unique
    }
}

# Main script logic
$DetectedSoftware = Fetch-UninstallKeys -TargetName $TargetSoftware

# Check if software is installed and output result
if ($DetectedSoftware) {
    Write-Host "$TargetSoftware is installed."
    exit 0
} else {
    Write-Host "$TargetSoftware is NOT installed."
    exit 1
}
```

---

## Requirements

- Architecture: **x64arm64**  
- Minimum Windows: **W11_23H2**

---

## Assignment

**Group Assignment**

| Group | Group ID | Intent | Notification | Filter | Filter ID | Filter Intent |
|-------|----------|--------|--------------|--------|-----------|---------------|
| All Company | `5837c301-1779-4370-b365-86410ec22313` | required | showAll | test | `0cd8920f-436e-40d1-94ea-b131d030a33e` | include |
| All Users | `6d1e151e-d618-45d8-8495-bc6a5bd8b004` | required | showAll | test | `0cd8920f-436e-40d1-94ea-b131d030a33e` | include |

---

## Return Codes

| Code | Type |
|------|------|
| 0 | success |
| 1707 | success |
| 3010 | softReboot |
| 1641 | hardReboot |
| 1618 | retry |

---

_Generated by Intune Win32 App Uploader on 2026-04-18 13:56:43_
