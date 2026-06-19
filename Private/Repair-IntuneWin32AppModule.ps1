# Win32Forge v1.1.0  |  https://github.com/durrante/Win32Forge  |  MIT  |  Release history: CHANGELOG.md
<#
.SYNOPSIS
    Patches all installed IntuneWin32App module versions for compatibility fixes.

.DESCRIPTION
    Finds every installed copy of the IntuneWin32App module and applies locale-fix patches so
    that the tool works correctly on non-US systems (e.g. en-GB where day > 12 breaks
    InvariantCulture DateTime parsing).

    Patches for module version 1.3.x (Windows PowerShell 5.1 path):
      Patch 1 — New-IntuneWin32AppRequirementRule.ps1
          Adds W11_23H2, W11_24H2, arm64, x64arm64, AllWithARM64 support.

      Patch 2 — Add-IntuneWin32App.ps1
          Normalises ExpiresOn in the Begin block before arithmetic, so en-GB locales
          (day > 12) don't produce an InvariantCulture string-parse failure.

      Patch 3 — Invoke-AzureStorageBlobUpload.ps1 (1.3.x)
          Same locale fix for AccessToken.ExpiresOn in the upload chunk loop.

    Patches for module version 1.4.x / 1.5.x (PowerShell 7 path):
      Patch A — Private\Invoke-AzureStorageBlobUpload.ps1
          Replaces [DateTimeOffset]::Parse(ExpiresOn.ToString(), InvariantCulture)
          with a direct .ToUniversalTime() call — no string round-trip, no locale issue.

      Patch B — Public\Test-AccessToken.ps1
          Same fix for the identical pattern in Test-AccessToken.

    Safe to call multiple times — skips any patch already applied.
    Returns $true if at least one patch was applied, $false if all were already up to date.
#>

function Repair-IntuneWin32AppModule {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $allModules = Get-Module IntuneWin32App -ListAvailable -ErrorAction SilentlyContinue
    if (-not $allModules) {
        Write-Warning 'Repair-IntuneWin32AppModule: IntuneWin32App module not found — skipping patches.'
        return $false
    }

    $anyPatched = $false

    foreach ($mod in $allModules) {
        $moduleBase = $mod.ModuleBase
        $version    = $mod.Version
        Write-Verbose "Repair-IntuneWin32AppModule: checking version $version at $moduleBase"

        # ──────────────────────────────────────────────────────────────────────
        # Determine generation: 1.3.x uses one code style; 1.4+/1.5+ uses another
        # ──────────────────────────────────────────────────────────────────────
        $isModern = ($version.Major -ge 2) -or ($version.Major -eq 1 -and $version.Minor -ge 4)

        if ($isModern) {
            # ══════════════════════════════════════════════════════════════════
            # Patch A — Invoke-AzureStorageBlobUpload.ps1 (1.4+/1.5+)
            #   [DateTimeOffset]::Parse(ExpiresOn.ToString(), InvariantCulture, …)
            #   fails on en-GB when day > 12 (month 13 = invalid).
            #   Fix: use ExpiresOn.ToUniversalTime() directly.
            # ══════════════════════════════════════════════════════════════════
            $blobFile = Join-Path $moduleBase 'Private\Invoke-AzureStorageBlobUpload.ps1'
            if (-not (Test-Path $blobFile)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $blobFile not found — skipping Patch A."
            }
            else {
                $blobContent = Get-Content $blobFile -Raw
                # Guard: already patched if the Win32Forge sentinel assignment is present.
                if ($blobContent -match '\$_eo\s*=\s*\$Global:AccessToken\.ExpiresOn') {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch A already applied.'
                }
                else {
                    # Match the original block regardless of indentation or line-endings (CRLF vs LF).
                    # NOTE: an earlier version used a literal here-string with .Replace(), which silently
                    # no-op'd when the module file used CRLF and this script used LF (or vice versa) yet
                    # still reported success. Use a whitespace-tolerant regex and only claim success if
                    # the content actually changed.
                    $patternA = '(?s)#\s*Convert ExpiresOn to DateTimeOffset in UTC\s*\r?\n\s*' +
                                '\$ExpiresOnUTC\s*=\s*\[DateTimeOffset\]::Parse\(\s*' +
                                '\$Global:AccessToken\.ExpiresOn\.ToString\(\)\s*,\s*' +
                                '\[System\.Globalization\.CultureInfo\]::InvariantCulture\s*,\s*' +
                                '\[System\.Globalization\.DateTimeStyles\]::AssumeUniversal\s*' +
                                '\)\.ToUniversalTime\(\)'
                    $newStr = @'
        # Patch A (Win32Forge): avoid locale-specific ToString() + InvariantCulture Parse,
        # which fails when day > 12 (e.g. en-GB '18/06/2026' parsed as month 18 = invalid).
        # ExpiresOn is a DateTimeOffset — convert to UTC directly without string round-trip.
        $_eo = $Global:AccessToken.ExpiresOn
        $ExpiresOnUTC = if ($_eo -is [System.DateTimeOffset]) {
            $_eo.ToUniversalTime()
        } elseif ($_eo -is [datetime]) {
            [System.DateTimeOffset]::new([datetime]::SpecifyKind($_eo, [System.DateTimeKind]::Utc), [System.TimeSpan]::Zero)
        } else {
            [System.DateTimeOffset]::UtcNow.AddHours(1)
        }
'@
                    # Use a MatchEvaluator so '$' sequences in $newStr are treated literally,
                    # not as regex replacement back-references.
                    $patchedContent = [regex]::Replace($blobContent, $patternA, { param($m) $newStr })
                    if ($patchedContent -ne $blobContent) {
                        try {
                            Set-Content -Path $blobFile -Value $patchedContent -Encoding UTF8 -Force
                            Write-Verbose "Repair-IntuneWin32AppModule: Patch A applied to $blobFile"
                            $anyPatched = $true
                        }
                        catch {
                            Write-Warning "Repair-IntuneWin32AppModule: could not write Patch A — $_"
                        }
                    }
                    else {
                        Write-Verbose 'Repair-IntuneWin32AppModule: target for Patch A not found in Invoke-AzureStorageBlobUpload.ps1 — skipping (format changed or already patched).'
                    }
                }
            }

            # ══════════════════════════════════════════════════════════════════
            # Patch B — Test-AccessToken.ps1 (1.4+/1.5+)
            #   Same [DateTimeOffset]::Parse(ExpiresOn.ToString(), InvariantCulture) bug.
            # ══════════════════════════════════════════════════════════════════
            $tokenFile = Join-Path $moduleBase 'Public\Test-AccessToken.ps1'
            if (-not (Test-Path $tokenFile)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $tokenFile not found — skipping Patch B."
            }
            else {
                $tokenContent = Get-Content $tokenFile -Raw
                if ($tokenContent -match '_eo2\b.*is \[System\.DateTimeOffset\]') {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch B already applied.'
                }
                else {
                    $oldStrB = '            # Convert ExpiresOn to DateTimeOffset in UTC' + "`r`n" +
                               '            $ExpiresOnUTC = [DateTimeOffset]::Parse($Global:AccessToken.ExpiresOn.ToString(), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()'
                    $oldStrB_lf = '            # Convert ExpiresOn to DateTimeOffset in UTC' + "`n" +
                               '            $ExpiresOnUTC = [DateTimeOffset]::Parse($Global:AccessToken.ExpiresOn.ToString(), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()'
                    $newStrB = @'
            # Patch B (Win32Forge): avoid locale-specific ToString() + InvariantCulture Parse.
            $_eo2 = $Global:AccessToken.ExpiresOn
            $ExpiresOnUTC = if ($_eo2 -is [System.DateTimeOffset]) {
                $_eo2.ToUniversalTime()
            } elseif ($_eo2 -is [datetime]) {
                [System.DateTimeOffset]::new([datetime]::SpecifyKind($_eo2, [System.DateTimeKind]::Utc), [System.TimeSpan]::Zero)
            } else {
                [System.DateTimeOffset]::UtcNow.AddHours(1)
            }
'@
                    if ($tokenContent -match [regex]::Escape('$Global:AccessToken.ExpiresOn.ToString()')) {
                        $patchedTokenContent = $tokenContent.Replace($oldStrB, $newStrB)
                        if ($patchedTokenContent -eq $tokenContent) {
                            # Try LF line endings
                            $patchedTokenContent = $tokenContent.Replace($oldStrB_lf, $newStrB)
                        }
                        if ($patchedTokenContent -ne $tokenContent) {
                            try {
                                Set-Content -Path $tokenFile -Value $patchedTokenContent -Encoding UTF8 -Force
                                Write-Verbose "Repair-IntuneWin32AppModule: Patch B applied to $tokenFile"
                                $anyPatched = $true
                            }
                            catch {
                                Write-Warning "Repair-IntuneWin32AppModule: could not write Patch B — $_"
                            }
                        }
                        else {
                            Write-Verbose 'Repair-IntuneWin32AppModule: could not match Patch B target text exactly — skipping (already patched or format changed).'
                        }
                    }
                    else {
                        Write-Verbose 'Repair-IntuneWin32AppModule: target for Patch B not found in Test-AccessToken.ps1 — skipping.'
                    }
                }
            }

            # ══════════════════════════════════════════════════════════════════
            # Patch C — New-IntuneWin32AppRequirementRule.ps1 (1.4+/1.5+)
            #   The shipped 1.5.0 ValidateSet for MinimumSupportedWindowsRelease
            #   stops at W11_22H2, so selecting Windows 11 23H2 / 24H2 fails with
            #   "Cannot validate argument on parameter 'MinimumSupportedWindowsRelease'".
            #   Add the missing releases to the ValidateSet AND their API-value
            #   mappings in $OperatingSystemTable. Matches MSEndpointMgr PR #219:
            #     W11_23H2 -> Windows11_23H2 ,  W11_24H2 -> Windows11_24H2
            # ══════════════════════════════════════════════════════════════════
            $reqFileModern = Join-Path $moduleBase 'Public\New-IntuneWin32AppRequirementRule.ps1'
            if (-not (Test-Path $reqFileModern)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $reqFileModern not found — skipping Patch C."
            }
            else {
                $reqContentModern = Get-Content $reqFileModern -Raw
                if ($reqContentModern -match 'W11_24H2') {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch C already applied (W11_24H2 present).'
                }
                else {
                    $patchedReq = $reqContentModern

                    # 1) Extend the ValidateSet. The ValidateSet entry for W11_22H2 is
                    #    immediately followed by ")]" (closing the set + the attribute),
                    #    which uniquely distinguishes it from the hashtable entry below.
                    $patchedReq = $patchedReq -replace
                        '"W11_22H2"(\s*)\)\]',
                        '"W11_22H2", "W11_23H2", "W11_24H2"$1)]'

                    # 2) Add the two API-value mappings after the W11_22H2 table entry,
                    #    preserving the original line's indentation (captured as $1).
                    $patchedReq = [regex]::Replace(
                        $patchedReq,
                        '(?m)^(\s*)"W11_22H2"\s*=\s*"Windows11_22H2".*$',
                        '$1"W11_22H2" = "Windows11_22H2"' + "`r`n" +
                        '$1"W11_23H2" = "Windows11_23H2"' + "`r`n" +
                        '$1"W11_24H2" = "Windows11_24H2"'
                    )

                    if (($patchedReq -ne $reqContentModern) -and
                        ($patchedReq -match 'W11_24H2"\s*\)\]|"W11_24H2", ') -and
                        ($patchedReq -match '"W11_24H2"\s*=\s*"Windows11_24H2"')) {
                        try {
                            Set-Content -Path $reqFileModern -Value $patchedReq -Encoding UTF8 -Force
                            Write-Verbose "Repair-IntuneWin32AppModule: Patch C applied to $reqFileModern"
                            $anyPatched = $true
                        }
                        catch {
                            Write-Warning "Repair-IntuneWin32AppModule: could not write Patch C — $_"
                        }
                    }
                    else {
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch C targets not found (ValidateSet/table format changed) — skipping.'
                    }
                }
            }

            # ══════════════════════════════════════════════════════════════════
            # Patch D — Invoke-AzureStorageBlobUploadRenew.ps1 (1.4+/1.5+)
            #   The SAS-URI renewal (triggered mid-upload on large packages) builds
            #   "deviceAppManagement/$Resource/..." but $Resource (the $FilesUri) ALREADY
            #   starts with "deviceAppManagement/", producing a doubled segment and a
            #   "BadRequest: Resource not found for the segment 'deviceAppManagement'" error
            #   on every renewal — flooding the log and risking failed uploads of large apps.
            #   Sibling Wait-IntuneWin32AppFileProcessing uses $Resource directly; do the same.
            # ══════════════════════════════════════════════════════════════════
            $renewFile = Join-Path $moduleBase 'Private\Invoke-AzureStorageBlobUploadRenew.ps1'
            if (-not (Test-Path $renewFile)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $renewFile not found — skipping Patch D."
            }
            else {
                $renewContent = Get-Content $renewFile -Raw
                if ($renewContent -notmatch 'deviceAppManagement/\$\(\$Resource\)') {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch D already applied.'
                }
                else {
                    # Literal string replacements (within single lines — CRLF/LF agnostic).
                    $patchedRenew = $renewContent.Replace(
                        '-Resource "deviceAppManagement/$($Resource)/renewUpload"',
                        '-Resource "$($Resource)/renewUpload"').Replace(
                        '-Resource "deviceAppManagement/$($Resource)"',
                        '-Resource "$($Resource)"')
                    if ($patchedRenew -ne $renewContent) {
                        try {
                            Set-Content -Path $renewFile -Value $patchedRenew -Encoding UTF8 -Force
                            Write-Verbose "Repair-IntuneWin32AppModule: Patch D applied to $renewFile"
                            $anyPatched = $true
                        }
                        catch {
                            Write-Warning "Repair-IntuneWin32AppModule: could not write Patch D — $_"
                        }
                    }
                    else {
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch D target text not found (format changed) — skipping.'
                    }
                }
            }

            # ══════════════════════════════════════════════════════════════════
            # Patch E — Add-IntuneWin32App.ps1 category lookup (1.4+/1.5+)
            #   The category lookup used a server-side OData $filter on displayName with a
            #   URL-encoded value. That is unreliable for names containing spaces (e.g.
            #   'Data management' never matched), and an unmatched result returns the response
            #   wrapper whose null id then failed the whole app-create with:
            #     "ModelValidationFailure: A null value was found for the property named 'id'".
            #   Fix: fetch ALL categories and match locally (case-insensitive) — deterministic,
            #   no encoding involved — and only use a category that actually has an id.
            # ══════════════════════════════════════════════════════════════════
            $addAppFileModern = Join-Path $moduleBase 'Public\Add-IntuneWin32App.ps1'
            if (-not (Test-Path $addAppFileModern)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $addAppFileModern not found — skipping Patch E."
            }
            else {
                $addContentModern = Get-Content $addAppFileModern -Raw
                if ($addContentModern -match 'mobileAppCategories"\s*-ErrorAction\s*"Stop"\s*\|\s*Where-Object') {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch E already applied.'
                }
                else {
                    # Replace the whole category query line (any encoding variant) with a deterministic
                    # fetch-all + local match. MatchEvaluator keeps the replacement string literal.
                    $localMatch = '$Category = Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceAppManagement/mobileAppCategories" -ErrorAction "Stop" | Where-Object { $_.displayName -eq $CategoryNameItem } | Select-Object -First 1'
                    $patchedAdd = [regex]::Replace(
                        $addContentModern,
                        '\$Category = Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceAppManagement/mobileAppCategories\?[^\r\n]*?-ErrorAction "Stop"',
                        { param($m) $localMatch })
                    # Only use the category when it resolved to a real id (idempotent: a no-op if an
                    # earlier patch already added the guard).
                    $patchedAdd = $patchedAdd.Replace(
                        'if ($null -ne $Category) {',
                        'if ($null -ne $Category -and $Category.id) {')
                    if (($patchedAdd -ne $addContentModern) -and
                        ($patchedAdd -match 'mobileAppCategories"\s*-ErrorAction\s*"Stop"\s*\|\s*Where-Object')) {
                        try {
                            Set-Content -Path $addAppFileModern -Value $patchedAdd -Encoding UTF8 -Force
                            Write-Verbose "Repair-IntuneWin32AppModule: Patch E applied to $addAppFileModern"
                            $anyPatched = $true
                        }
                        catch {
                            Write-Warning "Repair-IntuneWin32AppModule: could not write Patch E — $_"
                        }
                    }
                    else {
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch E target not found (format changed) — skipping.'
                    }
                }
            }
        }
        else {
            # ══════════════════════════════════════════════════════════════════
            # Legacy 1.3.x patches (Windows PowerShell 5.1 module path)
            # ══════════════════════════════════════════════════════════════════

            # Patch 1 — New-IntuneWin32AppRequirementRule.ps1
            $requirementRuleFile = Join-Path $moduleBase 'Public\New-IntuneWin32AppRequirementRule.ps1'
            if (-not (Test-Path $requirementRuleFile)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $requirementRuleFile not found — skipping Patch 1."
            }
            else {
                $reqContent = Get-Content $requirementRuleFile -Raw
                $alreadyPatched1 = ($reqContent -match 'W11_23H2' -and
                                    $reqContent -notmatch 'W11_25H2' -and
                                    $reqContent -match 'x64arm64')
                if ($alreadyPatched1) {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch 1 already applied.'
                }
                else {
                    Write-Verbose "Repair-IntuneWin32AppModule: applying Patch 1 — $requirementRuleFile"
                    $patchedContent = @'
function New-IntuneWin32AppRequirementRule {
    <#
    .SYNOPSIS
        Construct a new requirement rule as an optional requirement for Add-IntuneWin32App cmdlet.

    .DESCRIPTION
        Construct a new requirement rule as an optional requirement for Add-IntuneWin32App cmdlet.

    .PARAMETER Architecture
        Specify the architecture as a requirement for the Win32 app.
        Supported values: x64, x86, arm64, x64x86, AllWithARM64.

    .PARAMETER MinimumSupportedWindowsRelease
        Specify the minimum supported Windows release version as a requirement for the Win32 app.
        Supported values: W10_1607 through W11_24H2.

    .PARAMETER MinimumFreeDiskSpaceInMB
        Specify the minimum free disk space in MB as a requirement for the Win32 app.

    .PARAMETER MinimumMemoryInMB
        Specify the minimum required memory in MB as a requirement for the Win32 app.

    .PARAMETER MinimumNumberOfProcessors
        Specify the minimum number of required logical processors as a requirement for the Win32 app.

    .PARAMETER MinimumCPUSpeedInMHz
        Specify the minimum CPU speed in Mhz (as an integer) as a requirement for the Win32 app.

    .NOTES
        Author:      Nickolaj Andersen
        Contact:     @NickolajA
        Created:     2020-01-27
        Updated:     2025-12-07 (patched by Win32Forge for W11_24H2 + ARM64 support)

        Version history:
        1.0.0 - (2020-01-27) Function created
        1.0.7 - (2025-12-07) Added ARM64, x64x86, AllWithARM64; added OS values through W11_24H2;
                              W11_25H2 excluded — Intune API does not accept it yet
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [parameter(Mandatory = $true, HelpMessage = "Specify the architecture as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("x64", "x86", "arm64", "x64x86", "x64arm64", "AllWithARM64")]
        [string]$Architecture,

        [parameter(Mandatory = $true, HelpMessage = "Specify the minimum supported Windows release version as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("W10_1607", "W10_1703", "W10_1709", "W10_1803", "W10_1809", "W10_1903", "W10_1909",
                     "W10_2004", "W10_20H2", "W10_21H1", "W10_21H2", "W10_22H2",
                     "W11_21H2", "W11_22H2", "W11_23H2", "W11_24H2")]
        [Alias("MinimumSupportedOperatingSystem")]
        [string]$MinimumSupportedWindowsRelease,

        [parameter(Mandatory = $false, HelpMessage = "Specify the minimum free disk space in MB as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [int]$MinimumFreeDiskSpaceInMB,

        [parameter(Mandatory = $false, HelpMessage = "Specify the minimum required memory in MB as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [int]$MinimumMemoryInMB,

        [parameter(Mandatory = $false, HelpMessage = "Specify the minimum number of required logical processors as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [int]$MinimumNumberOfProcessors,

        [parameter(Mandatory = $false, HelpMessage = "Specify the minimum CPU speed in Mhz (as an integer) as a requirement for the Win32 app.")]
        [ValidateNotNullOrEmpty()]
        [int]$MinimumCPUSpeedInMHz
    )
    Process {
        $ArchitectureTable = @{
            "x64"          = "x64"
            "x86"          = "x86"
            "arm64"        = "arm64"
            "x64x86"       = "x64,x86"
            "x64arm64"     = "x64,arm64"
            "AllWithARM64" = "x64,x86,arm64"
        }

        $OperatingSystemTable = @{
            "W10_1607" = "1607"
            "W10_1703" = "1703"
            "W10_1709" = "1709"
            "W10_1803" = "1803"
            "W10_1809" = "1809"
            "W10_1903" = "1903"
            "W10_1909" = "1909"
            "W10_2004" = "2004"
            "W10_20H2" = "20H2"
            "W10_21H1" = "21H1"
            "W10_21H2" = "Windows10_21H2"
            "W10_22H2" = "Windows10_22H2"
            "W11_21H2" = "Windows11_21H2"
            "W11_22H2" = "Windows11_22H2"
            "W11_23H2" = "Windows11_23H2"
            "W11_24H2" = "Windows11_24H2"
        }

        $RequirementRule = [ordered]@{
            "allowedArchitectures"           = $ArchitectureTable[$Architecture]
            "applicableArchitectures"        = "none"
            "minimumSupportedWindowsRelease" = $OperatingSystemTable[$MinimumSupportedWindowsRelease]
        }

        if ($PSBoundParameters["MinimumFreeDiskSpaceInMB"])  { $RequirementRule.Add("minimumFreeDiskSpaceInMB",  $MinimumFreeDiskSpaceInMB)  }
        if ($PSBoundParameters["MinimumMemoryInMB"])          { $RequirementRule.Add("minimumMemoryInMB",          $MinimumMemoryInMB)          }
        if ($PSBoundParameters["MinimumNumberOfProcessors"]) { $RequirementRule.Add("minimumNumberOfProcessors", $MinimumNumberOfProcessors)  }
        if ($PSBoundParameters["MinimumCPUSpeedInMHz"])       { $RequirementRule.Add("minimumCpuSpeedInMHz",       $MinimumCPUSpeedInMHz)       }

        return $RequirementRule
    }
}
'@
                    try {
                        Set-Content -Path $requirementRuleFile -Value $patchedContent -Encoding UTF8 -Force
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch 1 applied.'
                        $anyPatched = $true
                    }
                    catch { Write-Warning "Repair-IntuneWin32AppModule: could not write Patch 1 — $_" }
                }
            }

            # Patch 2 — Add-IntuneWin32App.ps1 Begin block (1.3.x)
            $addAppFile = Join-Path $moduleBase 'Public\Add-IntuneWin32App.ps1'
            if (-not (Test-Path $addAppFile)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $addAppFile not found — skipping Patch 2."
            }
            else {
                $addLines = Get-Content $addAppFile
                $alreadyPatched2 = $addLines -match '_expiresOn'
                if ($alreadyPatched2) {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch 2 already applied.'
                }
                else {
                    $targetIdx = -1
                    for ($ln = 0; $ln -lt $addLines.Count; $ln++) {
                        if ($addLines[$ln] -match 'TokenLifeTime\s*=\s*\(\s*\$Global:AuthenticationHeader\.ExpiresOn') {
                            $targetIdx = $ln
                            break
                        }
                    }
                    if ($targetIdx -lt 0) {
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch 2 target not found — skipping.'
                    }
                    else {
                        $indent = ''
                        if ($addLines[$targetIdx] -match '^(\s+)') { $indent = $Matches[1] }
                        $replacement = @(
                            "$indent# Patch 2 (Win32Forge): normalise ExpiresOn so non-US locales (e.g. en-GB) don't"
                            "$indent# produce a culture-formatted string that InvariantCulture DateTime.Parse rejects."
                            "${indent}`$_expiresOn = `$Global:AuthenticationHeader.ExpiresOn"
                            "${indent}if (`$_expiresOn -is [string]) {"
                            "${indent}    try   { `$_expiresOn = [datetime]::Parse(`$_expiresOn, [System.Globalization.CultureInfo]::CurrentCulture) }"
                            "${indent}    catch { `$_expiresOn = [datetime]::UtcNow.AddHours(1) }"
                            "${indent}} elseif (`$_expiresOn -is [System.DateTimeOffset]) {"
                            "${indent}    `$_expiresOn = `$_expiresOn.UtcDateTime"
                            "${indent}} elseif (`$_expiresOn -isnot [datetime]) {"
                            "${indent}    `$_expiresOn = [datetime]::UtcNow.AddHours(1)"
                            "${indent}}"
                            "${indent}`$TokenLifeTime = (`$_expiresOn - (Get-Date).ToUniversalTime()).Minutes"
                        )
                        $newLines = [System.Collections.Generic.List[string]]::new()
                        for ($ln = 0; $ln -lt $addLines.Count; $ln++) {
                            if ($ln -eq $targetIdx) { $newLines.AddRange([string[]]$replacement) }
                            else                    { $newLines.Add($addLines[$ln]) }
                        }
                        try {
                            Set-Content -Path $addAppFile -Value $newLines -Encoding UTF8 -Force
                            Write-Verbose 'Repair-IntuneWin32AppModule: Patch 2 applied.'
                            $anyPatched = $true
                        }
                        catch { Write-Warning "Repair-IntuneWin32AppModule: could not write Patch 2 — $_" }
                    }
                }
            }

            # Patch 3 — Invoke-AzureStorageBlobUpload.ps1 (1.3.x)
            $blobFile3 = Join-Path $moduleBase 'Private\Invoke-AzureStorageBlobUpload.ps1'
            if (-not (Test-Path $blobFile3)) {
                Write-Verbose "Repair-IntuneWin32AppModule: $blobFile3 not found — skipping Patch 3."
            }
            else {
                $blobLines = Get-Content $blobFile3
                $alreadyPatched3 = $blobLines -match '_atExpiry'
                if ($alreadyPatched3) {
                    Write-Verbose 'Repair-IntuneWin32AppModule: Patch 3 already applied.'
                }
                else {
                    $targetIdx3 = -1
                    for ($ln = 0; $ln -lt $blobLines.Count; $ln++) {
                        if ($blobLines[$ln] -match 'TokenExpiresMinutes\s*=\s*\(\s*\$Global:AccessToken\.ExpiresOn') {
                            $targetIdx3 = $ln; break
                        }
                    }
                    if ($targetIdx3 -lt 0) {
                        Write-Verbose 'Repair-IntuneWin32AppModule: Patch 3 target not found — skipping.'
                    }
                    else {
                        $indent3 = ''
                        if ($blobLines[$targetIdx3] -match '^(\s+)') { $indent3 = $Matches[1] }
                        $replacement3 = @(
                            "$indent3# Patch 3 (Win32Forge): normalise AccessToken.ExpiresOn before arithmetic."
                            "${indent3}`$_atExpiry = if (`$Global:AccessToken -and `$Global:AccessToken.ExpiresOn) {"
                            "${indent3}    `$e = `$Global:AccessToken.ExpiresOn"
                            "${indent3}    if (`$e -is [System.DateTimeOffset])  { `$e.UtcDateTime }"
                            "${indent3}    elseif (`$e -is [datetime])            { `$e }"
                            "${indent3}    else { try { [datetime]::Parse([string]`$e, [System.Globalization.CultureInfo]::CurrentCulture) } catch { [datetime]::UtcNow.AddHours(1) } }"
                            "${indent3}} else { [datetime]::UtcNow.AddHours(1) }"
                            "${indent3}`$TokenExpiresMinutes = (`$_atExpiry - `$UTCDateTime).Minutes"
                        )
                        $newBlobLines = [System.Collections.Generic.List[string]]::new()
                        for ($ln = 0; $ln -lt $blobLines.Count; $ln++) {
                            if ($ln -eq $targetIdx3) { $newBlobLines.AddRange([string[]]$replacement3) }
                            else                     { $newBlobLines.Add($blobLines[$ln]) }
                        }
                        try {
                            Set-Content -Path $blobFile3 -Value $newBlobLines -Encoding UTF8 -Force
                            Write-Verbose 'Repair-IntuneWin32AppModule: Patch 3 applied.'
                            $anyPatched = $true
                        }
                        catch { Write-Warning "Repair-IntuneWin32AppModule: could not write Patch 3 — $_" }
                    }
                }
            }
        }
    }

    return $anyPatched
}
