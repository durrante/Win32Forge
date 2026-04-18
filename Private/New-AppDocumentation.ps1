<#
.SYNOPSIS
    Generates a Markdown documentation file for a deployed Intune Win32 application.

.DESCRIPTION
    Creates a per-app .md file in the documentation folder containing:
      - Application details (name, version, publisher, description, author, categories)
      - Packaging info (source folder, setup file, .intunewin path, logo path)
      - Install and uninstall commands, install context, with PSADT note if applicable
      - Detection method summary (includes script content for Script type)
      - Requirement rules (including any additional requirement script)
      - Assignment details:
          Groups  — name, AAD object ID, intent, notification, filter name/ID/intent
          Flat    — type, intent, notification, filter name/ID
      - Return codes table if custom codes are configured
      - Information URL / Privacy URL if present
      - Intune App ID, portal link, and upload timestamp

    The doc file is named: <DisplayName>_<Version>_<YYYYMMDD>.md
    Logo is copied alongside the doc file; its path is noted (not embedded as inline image).
#>

function New-AppDocumentation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$AppConfig,

        [Parameter(Mandatory)]
        [PSCustomObject]$IntuneApp,

        [Parameter(Mandatory)]
        [string]$DocumentationPath,

        [string]$IntunewinPath = ''
    )

    New-Item -ItemType Directory -Path $DocumentationPath -Force | Out-Null

    $safeName    = $AppConfig.DisplayName -replace '[\\/:*?"<>|]', '_'
    $safeVersion = ($AppConfig.Version ?? 'NoVersion') -replace '[\\/:*?"<>|]', '_'
    $dateStr     = Get-Date -Format 'yyyyMMdd'
    $docFileName = "${safeName}_${safeVersion}_${dateStr}.md"
    $docPath     = Join-Path $DocumentationPath $docFileName

    #region Logo
    $logoNote = '_No logo provided_'
    if ($AppConfig.LogoPath -and (Test-Path $AppConfig.LogoPath)) {
        $logoExt      = [System.IO.Path]::GetExtension($AppConfig.LogoPath)
        $logoDestName = "${safeName}_Logo${logoExt}"
        $logoDest     = Join-Path $DocumentationPath $logoDestName
        Copy-Item -Path $AppConfig.LogoPath -Destination $logoDest -Force
        $logoNote     = "``$logoDest``"
    }
    #endregion

    #region Optional summary fields
    $author      = if ($AppConfig.Owner)       { $AppConfig.Owner }       else { '-' }
    $description = if ($AppConfig.Description) { $AppConfig.Description } else { '-' }
    $infoUrl     = if ($AppConfig.InformationURL) { "[$($AppConfig.InformationURL)]($($AppConfig.InformationURL))" } else { '-' }
    $privUrl     = if ($AppConfig.PrivacyURL)     { "[$($AppConfig.PrivacyURL)]($($AppConfig.PrivacyURL))" }         else { '-' }
    $installCtx  = if ($AppConfig.InstallContext) { $AppConfig.InstallContext } else { 'System' }

    $categories = if ($AppConfig.Categories -and @($AppConfig.Categories).Count -gt 0) {
        (@($AppConfig.Categories) -join ', ')
    } else { '-' }
    #endregion

    #region Detection summary
    $det = $AppConfig.Detection
    $detSummary = switch ($det.Type) {
        'Script' {
            $scriptName = Split-Path $det.ScriptPath -Leaf
            $scriptContent = ''
            if ($det.ScriptPath -and (Test-Path $det.ScriptPath)) {
                $raw = Get-Content $det.ScriptPath -Raw -ErrorAction SilentlyContinue
                if ($raw) {
                    $scriptContent = "`n`n**Script Content:**`n`n``````powershell`n$($raw.TrimEnd())`n``````"
                }
            }
            "**PowerShell Script**: ``$scriptName``  `n" +
            "- Enforce signature check: $($det.EnforceSignatureCheck)  `n" +
            "- Run as 32-bit: $($det.RunAs32Bit)$scriptContent"
        }
        'MSI' {
            $verLine = if ($det.ProductVersion) { "`n- Version: $($det.ProductVersionOperator) $($det.ProductVersion)" } else { '' }
            "**MSI Product Code**: ``$($det.ProductCode)``$verLine"
        }
        'Registry' {
            $valueLine = if ($det.ValueName) { "`n- Value name: $($det.ValueName)" } else { '' }
            $opLine    = if ($det.Value)     { "`n- Operator / Value: $($det.Operator) ``$($det.Value)``" } else { '' }
            "**Registry**: ``$($det.KeyPath)``  `n" +
            "- Detection type: $($det.DetectionType)$valueLine$opLine  `n" +
            "- Check 32-bit: $($det.Check32BitOn64System)"
        }
        'File' {
            $opLine = if ($det.Value) { "`n- Operator / Value: $($det.Operator) ``$($det.Value)``" } else { '' }
            "**File/Folder**: ``$($det.Path)\$($det.FileOrFolder)``  `n" +
            "- Detection type: $($det.DetectionType)$opLine  `n" +
            "- Check 32-bit: $($det.Check32BitOn64System)"
        }
        default { "Unknown ($($det.Type))" }
    }
    #endregion

    #region Assignment summary
    $asg = $AppConfig.Assignment
    $asgSummary = if (-not $asg -or $asg.Type -eq 'None') {
        '_Not configured_'
    } elseif ($asg.Type -eq 'Group') {
        $groups = @($asg.Groups)
        if ($groups.Count -gt 0) {
            $rows = @($groups | ForEach-Object {
                $grp     = $_
                $gName   = if ($grp -is [hashtable]) { $grp.GroupName   ?? $grp.DisplayName ?? 'Unknown' }  else { [string]($grp.GroupName   ?? $grp.DisplayName ?? 'Unknown') }
                $gId     = if ($grp -is [hashtable]) { $grp.GroupID     ?? $grp.id          ?? '' }          else { [string]($grp.GroupID     ?? $grp.id          ?? '') }
                $gInt    = if ($grp -is [hashtable]) { $grp.Intent      ?? 'required' }                      else { [string]($grp.Intent      ?? 'required') }
                $gNotif  = if ($grp -is [hashtable]) { $grp.Notification ?? 'showAll' }                      else { [string]($grp.Notification ?? 'showAll') }
                $gFiltN  = if ($grp -is [hashtable]) { $grp.FilterName  ?? '' }                              else { [string]($grp.FilterName  ?? '') }
                $gFiltId = if ($grp -is [hashtable]) { $grp.FilterID    ?? '' }                              else { [string]($grp.FilterID    ?? '') }
                $gFiltI  = if ($grp -is [hashtable]) { $grp.FilterIntent ?? 'include' }                      else { [string]($grp.FilterIntent ?? 'include') }

                $hasFilter = $gFiltN -and $gFiltN -ne '(No filter)'
                $fName   = if ($hasFilter) { $gFiltN }               else { '-' }
                $fId     = if ($hasFilter -and $gFiltId) { "``$gFiltId``" } else { '-' }
                $fIntent = if ($hasFilter) { $gFiltI }               else { '-' }
                $gIdCell = if ($gId) { "``$gId``" } else { '-' }

                "| $gName | $gIdCell | $gInt | $gNotif | $fName | $fId | $fIntent |"
            })
            "**Group Assignment**`n`n" +
            "| Group | Group ID | Intent | Notification | Filter | Filter ID | Filter Intent |`n" +
            "|-------|----------|--------|--------------|--------|-----------|---------------|`n" +
            "$($rows -join "`n")"
        } else {
            '**Group** _(no groups configured)_'
        }
    } else {
        $fPart = '-'
        $fIdPart = '-'
        if ($asg.FilterID) {
            $fName   = if ($asg.FilterName) { $asg.FilterName } else { $asg.FilterID }
            $fPart   = "$fName (``$($asg.FilterID)``)"
            $fIdPart = $asg.FilterIntent ?? 'include'
        }
        "| Type | Intent | Notification | Filter | Filter Intent |`n" +
        "|------|--------|--------------|--------|---------------|`n" +
        "| $($asg.Type) | $($asg.Intent ?? 'required') | $($asg.Notification ?? 'showAll') | $fPart | $fIdPart |"
    }
    #endregion

    #region Requirements
    $reqSummary  = "- Architecture: **$($AppConfig.Architecture ?? 'x64')**  `n"
    $reqSummary += "- Minimum Windows: **$($AppConfig.MinimumSupportedWindowsRelease ?? 'W10_2004')**"
    if ($AppConfig.RequirementScript) {
        $rsName = Split-Path $AppConfig.RequirementScript.ScriptPath -Leaf
        $reqSummary += "  `n- Additional script: ``$rsName``"
    }
    #endregion

    #region Return codes table
    $rcSection = ''
    $rcList = @($AppConfig.ReturnCodes)
    if ($rcList.Count -gt 0) {
        $rcRows = @($rcList | ForEach-Object {
            $rc   = $_
            $code = if ($rc -is [hashtable]) { $rc.ReturnCode ?? $rc.returnCode } else { $rc.ReturnCode ?? $rc.returnCode }
            $type = if ($rc -is [hashtable]) { $rc.Type ?? $rc.type ?? 'success' } else { $rc.Type ?? $rc.type ?? 'success' }
            "| $code | $type |"
        })
        $rcSection = @"

---

## Return Codes

| Code | Type |
|------|------|
$($rcRows -join "`n")
"@
    }
    #endregion

    #region PSADT note
    $psadtNote = ''
    if ($AppConfig.IsPSADT) {
        $psadtVer = 'v4'
        if ($AppConfig.SourceFolder -and
            -not (Test-Path (Join-Path $AppConfig.SourceFolder 'Invoke-AppDeployToolkit.exe'))) {
            $psadtVer = 'v3'
        }
        $psadtNote = @"

> **PSADT Package** ($psadtVer)
> Install and uninstall commands use the PSAppDeployToolkit framework.
> Silent mode is enforced; the toolkit handles all UI suppression and logging.

"@
    }
    #endregion

    #region .intunewin info
    $intunewinSection = if ($IntunewinPath -and (Test-Path $IntunewinPath)) {
        "``$IntunewinPath``  ($('{0:N2}' -f ((Get-Item $IntunewinPath).Length / 1MB)) MB)"
    } else { '_Not recorded_' }
    #endregion

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $appId     = $IntuneApp.id ?? '_Unknown_'
    $appUrl    = "https://intune.microsoft.com/#blade/Microsoft_Intune_Apps/SettingsMenu/0/appId/$appId"

    $markdown = @"
# $($AppConfig.DisplayName)

| Field | Value |
|-------|-------|
| Display Name | $($AppConfig.DisplayName) |
| Version | $($AppConfig.Version ?? '-') |
| Publisher | $($AppConfig.Publisher ?? '-') |
| Description | $description |
| Author | $author |
| Notes | $($AppConfig.Notes ?? '-') |
| Categories | $categories |
| Install Context | $installCtx |
| Information URL | $infoUrl |
| Privacy URL | $privUrl |
| Intune App ID | ``$appId`` |
| Uploaded | $timestamp |
| Template | $($AppConfig.Template ?? '-') |

[View in Intune Portal]($appUrl)

---

## Packaging

| Field | Value |
|-------|-------|
| Source Folder | ``$($AppConfig.SourceFolder)`` |
| Setup File | ``$($AppConfig.SetupFile)`` |
| .intunewin | $intunewinSection |
| Logo | $logoNote |

---

## Commands
$psadtNote
| Command | Value |
|---------|-------|
| Install | ``$($AppConfig.InstallCommandLine)`` |
| Uninstall | ``$($AppConfig.UninstallCommandLine)`` |

---

## Detection Method

$detSummary

---

## Requirements

$reqSummary

---

## Assignment

$asgSummary
$rcSection

---

---

> Generated by **[Win32Forge](https://modernworkspacehub.com)** on ${timestamp}
>
> Win32Forge is a free, open source tool provided **without warranty** of any kind — use at your own risk.
> Visit [modernworkspacehub.com](https://modernworkspacehub.com) for more Intune resources and guides.
"@

    $markdown | Set-Content -Path $docPath -Encoding UTF8
    Write-Host "  [OK] Documentation saved: $docPath" -ForegroundColor Green

    return $docPath
}
