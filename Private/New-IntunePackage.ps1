# Win32Forge v1.1.0  |  https://github.com/durrante/Win32Forge  |  MIT  |  Release history: CHANGELOG.md
<#
.SYNOPSIS
    Forces OneDrive / Files On-Demand placeholders in a source tree to be downloaded locally.

.DESCRIPTION
    IntuneWinAppUtil.exe reads every source file with a plain FileStream. When a file is a
    dehydrated cloud placeholder (OneDrive "Files On-Demand"), that read goes through the cloud
    filter driver, and when the driver rejects it the tool dies with an opaque fatal error:

        System.IO.IOException: The cloud operation is invalid.
           at System.IO.FileStream.ReadCore(...)
           at ...ZipUtil.CreateEntryFromFile(...)
        ERROR  File '<output>.intunewin' has failed to be generated

    "The cloud operation is invalid" is Win32 error 362 (ERROR_CLOUD_FILE_INVALID_REQUEST).
    It is far more likely when the source is reached through a reparse point — which is exactly
    what the MAX_PATH junction workaround below creates — because the placeholder is then being
    recalled through a path the sync engine did not hand out.

    The fix is to hydrate the files BEFORE packaging, always against the real source path, never
    through the junction. Reading a single byte from a placeholder makes OneDrive download the
    whole file; the attributes are re-checked afterwards and anything still dehydrated is read in
    full as a fallback.

.PARAMETER SourceFolder
    The real source folder (NOT a junction to it).

.PARAMETER Force
    Read every file in full rather than only the ones flagged as placeholders. Used as a retry
    after IntuneWinAppUtil.exe has already failed with a cloud error, both to catch files whose
    attributes lie and to identify exactly which file cannot be read.

.OUTPUTS
    PSCustomObject with Checked, Hydrated, Failed (int) and FailedFiles (string[]).
#>
function Invoke-SourceHydration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder,

        [switch]$Force
    )

    # FILE_ATTRIBUTE_OFFLINE 0x1000 | RECALL_ON_OPEN 0x40000 | RECALL_ON_DATA_ACCESS 0x400000.
    # [System.IO.FileAttributes] only names Offline, so compare the raw bits.
    $placeholderMask = 0x00441000

    $checked     = 0
    $hydrated    = 0
    $failedFiles = [System.Collections.Generic.List[string]]::new()
    $pinQueue    = [System.Collections.Generic.List[object]]::new()
    $cloudErrorPattern = 'cloud operation is invalid|cloud file provider|cloud operation was not completed'

    try {
        $enumOpts = [System.IO.EnumerationOptions]::new()
        $enumOpts.RecurseSubdirectories = $true
        $enumOpts.IgnoreInaccessible    = $true
        # Default is Hidden|System — those files still go into the package, so scan them too.
        $enumOpts.AttributesToSkip      = [System.IO.FileAttributes]0

        $files = ([System.IO.DirectoryInfo]::new($SourceFolder)).EnumerateFiles('*', $enumOpts)
    }
    catch {
        Write-ToolLog "Cloud-placeholder scan could not enumerate '$SourceFolder' — $($_.Exception.Message)" -Level WARN
        return [pscustomobject]@{ Checked = 0; Hydrated = 0; Failed = 0; FailedFiles = @() }
    }

    $announced = $false

    # EnumerateFiles is lazy — a missing/disappearing root or an unreadable directory throws on
    # MoveNext, i.e. from inside this foreach rather than from the call above. Wrap the whole loop
    # so hydration degrades to "partial scan + warning" instead of aborting the package run.
    try {
        foreach ($file in $files) {
            $checked++

            $isPlaceholder = ([int]$file.Attributes -band $placeholderMask) -ne 0
            if (-not $isPlaceholder -and -not $Force) { continue }

            if (-not $announced) {
                $what = if ($Force) { 'Verifying every source file is readable' } else { 'Downloading cloud-only (OneDrive) files' }
                Write-Host "  [*] $what — this can take a while on a large package..." -ForegroundColor Yellow
                Write-ToolLog "Cloud hydration pass started (Force=$([bool]$Force)) on '$SourceFolder'"
                $announced = $true
            }

            try {
                # Reading one byte is enough to make the sync engine recall the whole file.
                $fs = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open,
                                             [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try {
                    if ($fs.Length -gt 0) {
                        $probe = [byte[]]::new(1)
                        $null  = $fs.Read($probe, 0, 1)

                        # Some providers only recall the range that was touched. If the file still
                        # reports as a placeholder, stream the rest to force a full hydration.
                        $stillPlaceholder = ([int][System.IO.File]::GetAttributes($file.FullName) -band $placeholderMask) -ne 0
                        if ($Force -or $stillPlaceholder) {
                            $fs.Position = 0
                            $fs.CopyTo([System.IO.Stream]::Null)
                        }
                    }
                }
                finally { $fs.Dispose() }

                if ($isPlaceholder) { $hydrated++ }
            }
            catch {
                if ("$($_.Exception.Message)" -match $cloudErrorPattern) {
                    # Recall-on-read was rejected — queue for the pin fallback below.
                    $pinQueue.Add([pscustomobject]@{ File = $file; Error = $_.Exception.Message })
                }
                else {
                    $failedFiles.Add("$($file.FullName) — $($_.Exception.Message)")
                    Write-ToolLog "Could not hydrate/read '$($file.FullName)' — $($_.Exception.Message)" -Level WARN
                }
            }
        }
    }
    catch {
        Write-ToolLog "Cloud-placeholder scan stopped early after $checked file(s) in '$SourceFolder' — $($_.Exception.Message)" -Level WARN
    }

    # Pin fallback. With LongPathsEnabled=0, OneDrive rejects an on-read recall for any path over
    # 260 chars ("The cloud operation is invalid"), even via a \\?\ prefix. Setting the Pinned
    # attribute ("Always keep on this device") makes the sync engine download the file through its
    # own background path instead, which does work. Pin, wait for the download, then clear the pin
    # again — clearing Pinned (without setting Unpinned) leaves the file local, so it stays readable
    # for packaging while the user's OneDrive settings end up as they were.
    if ($pinQueue.Count -gt 0) {
        $pinnedBit   = 0x00080000   # FILE_ATTRIBUTE_PINNED
        $unpinnedBit = 0x00100000   # FILE_ATTRIBUTE_UNPINNED
        Write-Host "  [*] OneDrive refused $($pinQueue.Count) direct download$(if ($pinQueue.Count -ne 1) { 's' }) (usually paths over 260 chars) — asking OneDrive to sync them instead..." -ForegroundColor Yellow
        Write-ToolLog "Pin fallback for $($pinQueue.Count) file(s) whose on-read recall was rejected"

        $pending = [System.Collections.Generic.List[object]]::new()
        $totalBytes = 0
        foreach ($item in $pinQueue) {
            $path = $item.File.FullName
            try {
                $orig = [int][System.IO.File]::GetAttributes($path)
                $item | Add-Member -NotePropertyName WasPinned -NotePropertyValue (($orig -band $pinnedBit) -ne 0)
                $new  = ($orig -bor $pinnedBit) -band (-bnot $unpinnedBit)
                # [Enum]::ToObject — a plain cast rejects the Pinned bit, which has no enum member.
                [System.IO.File]::SetAttributes($path, [Enum]::ToObject([System.IO.FileAttributes], $new))
                $pending.Add($item)
                $totalBytes += $item.File.Length
            }
            catch {
                $failedFiles.Add("$path — $($item.Error) (pin fallback also failed: $($_.Exception.Message))")
                Write-ToolLog "Could not pin '$path' — $($_.Exception.Message)" -Level WARN
            }
        }

        # Allow a generous base plus time proportional to the download size (~1 s per MB).
        $deadline = (Get-Date).AddSeconds(90 + [math]::Ceiling($totalBytes / 1MB))
        $waiting  = @($pending)
        while ($waiting.Count -gt 0 -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 1500
            $waiting = @($waiting | Where-Object {
                try { ([int][System.IO.File]::GetAttributes($_.File.FullName) -band $placeholderMask) -ne 0 } catch { $true }
            })
        }

        $waitingPaths = [System.Collections.Generic.HashSet[string]]::new([string[]]@($waiting | ForEach-Object { $_.File.FullName }), [System.StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $pending) {
            $path = $item.File.FullName
            $stillPlaceholder = $waitingPaths.Contains($path)
            if ($stillPlaceholder) {
                $failedFiles.Add("$path — $($item.Error) (OneDrive did not finish downloading it after pinning)")
                Write-ToolLog "Pin fallback timed out for '$path'" -Level WARN
            }
            else {
                $hydrated++
                Write-ToolLog "Pin fallback downloaded '$path'" -Level DEBUG
            }
            # Restore the user's pin state (only if we changed it). Never set Unpinned — that would
            # dehydrate the file again before IntuneWinAppUtil.exe gets to read it.
            if (-not $item.WasPinned) {
                try {
                    $cur = [int][System.IO.File]::GetAttributes($path)
                    [System.IO.File]::SetAttributes($path, [Enum]::ToObject([System.IO.FileAttributes], ($cur -band (-bnot $pinnedBit))))
                }
                catch { Write-ToolLog "Could not restore pin state on '$path' — $($_.Exception.Message)" -Level WARN }
            }
        }
    }

    if ($hydrated -gt 0) {
        Write-Host "  [OK] $hydrated cloud-only file$(if ($hydrated -ne 1) { 's' }) downloaded locally." -ForegroundColor Green
    }
    if ($announced -or $hydrated -gt 0) {
        Write-ToolLog "Cloud hydration pass finished: checked=$checked hydrated=$hydrated failed=$($failedFiles.Count)"
    }
    else {
        Write-ToolLog "Cloud hydration pass: no placeholders among $checked file(s) in '$SourceFolder'." -Level DEBUG
    }

    return [pscustomobject]@{
        Checked     = $checked
        Hydrated    = $hydrated
        Failed      = $failedFiles.Count
        FailedFiles = $failedFiles.ToArray()
    }
}

<#
.SYNOPSIS
    Creates a .intunewin package from a source folder using IntuneWinAppUtil.exe.

.DESCRIPTION
    Wraps either:
      1. The IntuneWin32App module's New-IntuneWin32AppPackage cmdlet (preferred)
      2. IntuneWinAppUtil.exe directly (fallback)

    Returns the full path to the generated .intunewin file.
#>

function New-IntunePackage {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$SourceFolder,

        [Parameter(Mandatory)]
        [string]$SetupFile,

        [Parameter(Mandatory)]
        [string]$OutputFolder,

        # Path to IntuneWinAppUtil.exe - read from config if not supplied
        [string]$IntuneWinAppUtilPath = ''
    )

    # Ensure output folder exists
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

    # Resolve the utility path
    if (-not $IntuneWinAppUtilPath -or -not (Test-Path $IntuneWinAppUtilPath)) {
        # Try the module's bundled copy first
        $moduleBase = (Get-Module IntuneWin32App -ListAvailable | Select-Object -First 1).ModuleBase
        $moduleTool = Join-Path $moduleBase 'Bin\IntuneWinAppUtil.exe'
        if (Test-Path $moduleTool) {
            $IntuneWinAppUtilPath = $moduleTool
        }
        else {
            # Try the tool directory alongside this script
            $localTool = Join-Path $PSScriptRoot '..\Tools\IntuneWinAppUtil.exe'
            if (Test-Path $localTool) {
                $IntuneWinAppUtilPath = (Resolve-Path $localTool).Path
            }
        }
    }

    $setupFileFull = Join-Path $SourceFolder $SetupFile
    if (-not (Test-Path $setupFileFull)) {
        throw "Setup file not found: $setupFileFull"
    }

    Write-Host "  [*] Packaging: $SourceFolder" -ForegroundColor Yellow
    Write-Host "      Setup file: $SetupFile" -ForegroundColor Gray
    Write-Host "      Output:     $OutputFolder" -ForegroundColor Gray
    Write-ToolLog "IntuneWinAppUtil: SourceFolder='$SourceFolder'  SetupFile='$SetupFile'  Output='$OutputFolder'  Tool='$IntuneWinAppUtilPath'" -Level DEBUG

    # Call IntuneWinAppUtil.exe directly.
    # We skip the New-IntuneWin32AppPackage cmdlet — it wraps the same exe but without
    # output redirection, which can cause the process to block when run from a WPF host.
    $intunewinPath = $null

    if (-not $IntuneWinAppUtilPath -or -not (Test-Path $IntuneWinAppUtilPath)) {
        throw "IntuneWinAppUtil.exe not found. Run Setup-Win32Forge.ps1 to download it, or set the path in Config\config.json."
    }

    # Runs IntuneWinAppUtil.exe against a source folder. ProcessStartInfo with redirected
    # stdout/stderr prevents output-buffer freeze when run from a WPF host.
    function Invoke-PackageExe {
        param([string]$Source)
        $psi                        = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName               = $IntuneWinAppUtilPath
        $psi.Arguments              = "-c `"$Source`" -s `"$SetupFile`" -o `"$OutputFolder`" -q"
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $p = [System.Diagnostics.Process]::new()
        $p.StartInfo = $psi
        $p.Start() | Out-Null
        # Read output asynchronously to prevent deadlock if a buffer fills
        $o = $p.StandardOutput.ReadToEndAsync()
        $e = $p.StandardError.ReadToEndAsync()
        $p.WaitForExit()
        $o.Wait(); $e.Wait()
        return [pscustomobject]@{ ExitCode = $p.ExitCode; Stdout = $o.Result.Trim(); Stderr = $e.Result.Trim() }
    }

    # Pre-emptively pull down any OneDrive placeholders. IntuneWinAppUtil.exe cannot read a
    # dehydrated file and dies with an unrecoverable "The cloud operation is invalid" mid-zip,
    # so this has to happen before the first attempt — and against the real path, because the
    # junction fallback below makes recall through a reparse point even more likely to fail.
    $hydration = Invoke-SourceHydration -SourceFolder $SourceFolder
    if ($hydration.Failed -gt 0) {
        Write-Host "  [!] $($hydration.Failed) source file$(if ($hydration.Failed -ne 1) { 's' }) could not be read — packaging may fail." -ForegroundColor Yellow
    }

    # Signatures that identify the two known IntuneWinAppUtil.exe failure modes.
    $longPathSignature = 'Could not find a part of the path|DirectoryNotFoundException|PathTooLong|filename or extension is too long'
    $cloudSignature    = 'The cloud operation is invalid|cloud file provider|ERROR_CLOUD_FILE|CloudFile|cloud operation was not completed'

    $activeSource = $SourceFolder
    $junction     = $null
    $madeJunction = $false

    try {
        $result = Invoke-PackageExe -Source $activeSource

        # IntuneWinAppUtil.exe is a .NET Framework tool with the classic 260-char MAX_PATH limit.
        # A deep source tree (e.g. a PSADT payload under a long OneDrive path) can exceed it and fail
        # with "DirectoryNotFoundException: Could not find a part of the path". When we see that
        # signature, retry once via a short directory junction so the paths the tool opens are short.
        $looksLikeLongPath = ($result.ExitCode -ne 0) -and
            (("$($result.Stdout)`n$($result.Stderr)") -match $longPathSignature)
        if ($looksLikeLongPath) {
            $junctionRoot = Join-Path $env:SystemDrive 'W32F'
            $junction     = Join-Path $junctionRoot ([guid]::NewGuid().ToString('N').Substring(0, 8))
            try {
                New-Item -ItemType Directory -Path $junctionRoot -Force -ErrorAction Stop | Out-Null
                New-Item -ItemType Junction -Path $junction -Target $SourceFolder -ErrorAction Stop | Out-Null
                $madeJunction = $true
                $activeSource = $junction
                Write-Host "  [!] Source path exceeds the 260-char limit — retrying via short path ($junction)..." -ForegroundColor Yellow
                Write-ToolLog "Long-path failure detected; retrying package via junction '$junction' -> '$SourceFolder'" -Level WARN
                $result = Invoke-PackageExe -Source $activeSource
            }
            catch {
                Write-ToolLog "Could not create short-path junction for retry — $($_.Exception.Message)" -Level ERROR
            }
        }

        # A cloud error means a placeholder slipped past the pre-emptive pass — either its
        # attributes were stale, or it was only partially recalled. Read every file in full and
        # retry once. The hydration always runs against the REAL source, never the junction.
        $looksLikeCloud = ($result.ExitCode -ne 0) -and
            (("$($result.Stdout)`n$($result.Stderr)") -match $cloudSignature)
        if ($looksLikeCloud) {
            Write-Host '  [!] OneDrive blocked a file read — forcing a full download of the source, then retrying...' -ForegroundColor Yellow
            Write-ToolLog "Cloud-file failure detected; forcing full hydration of '$SourceFolder' before retry" -Level WARN

            $forced = Invoke-SourceHydration -SourceFolder $SourceFolder -Force
            if ($forced.Failed -gt 0) {
                # Name the offending files — IntuneWinAppUtil.exe never says which file it choked on.
                foreach ($f in $forced.FailedFiles) { Write-ToolLog "  unreadable: $f" -Level ERROR }
                $sample = ($forced.FailedFiles | Select-Object -First 3) -join '; '
                throw "$($forced.Failed) file(s) in the source folder cannot be read because OneDrive will not make them available offline. Right-click the source folder in Explorer and choose 'Always keep on this device', wait for the sync to finish, then try again. If the paths are over 260 characters, enabling Windows long path support (LongPathsEnabled) or moving the folder to a shorter path also fixes this. First failures: $sample"
            }

            $result = Invoke-PackageExe -Source $activeSource
        }
    }
    finally {
        # IMPORTANT: delete the reparse point only (non-recursive) so the real source is untouched.
        if ($madeJunction) {
            try { [System.IO.Directory]::Delete($junction, $false) }
            catch { Write-ToolLog "Could not remove junction '$junction' — $($_.Exception.Message)" -Level WARN }
        }
    }

    Write-ToolLog "IntuneWinAppUtil.exe exited: code=$($result.ExitCode)" -Level DEBUG
    if ($result.Stdout) { Write-ToolLog "  stdout: $($result.Stdout)" -Level DEBUG }
    if ($result.Stderr) { Write-ToolLog "  stderr: $($result.Stderr)" -Level $(if ($result.ExitCode -ne 0) { 'ERROR' } else { 'WARN' }) }

    if ($result.ExitCode -ne 0) {
        $errText = if ($result.Stderr) { $result.Stderr } else { $result.Stdout }

        # Give the cloud failure an actionable message instead of the raw .NET stack trace.
        if (("$($result.Stdout)`n$($result.Stderr)") -match $cloudSignature) {
            throw "IntuneWinAppUtil.exe could not read the source files because OneDrive refused to make them available offline (`"The cloud operation is invalid`"). Right-click '$SourceFolder' in Explorer and choose 'Always keep on this device', wait for the sync to complete, then try again (for paths over 260 characters, enabling Windows long path support also fixes this). Full output: $errText"
        }

        throw "IntuneWinAppUtil.exe failed (exit $($result.ExitCode))$(if ($errText) {": $errText"})"
    }

    # Locate the generated .intunewin file
    if (-not $intunewinPath) {
        $baseSetupName = [System.IO.Path]::GetFileNameWithoutExtension($SetupFile)
        $intunewinPath = Get-ChildItem -Path $OutputFolder -Filter '*.intunewin' |
                         Where-Object { $_.BaseName -eq $baseSetupName } |
                         Sort-Object LastWriteTime -Descending |
                         Select-Object -First 1 -ExpandProperty FullName

        # Fallback: just take the newest .intunewin in the output folder
        if (-not $intunewinPath) {
            $intunewinPath = Get-ChildItem -Path $OutputFolder -Filter '*.intunewin' |
                             Sort-Object LastWriteTime -Descending |
                             Select-Object -First 1 -ExpandProperty FullName
        }
    }

    if (-not $intunewinPath -or -not (Test-Path $intunewinPath)) {
        throw "Package was not created. No .intunewin file found in: $OutputFolder"
    }

    Write-Host "  [OK] Package created: $intunewinPath" -ForegroundColor Green
    Write-ToolLog "Package created: '$intunewinPath'  ($('{0:N2}' -f ((Get-Item $intunewinPath).Length / 1MB)) MB)"
    return $intunewinPath
}

<#
.SYNOPSIS
    Updates the inner content filename and Detection.xml FileName element inside a .intunewin ZIP.

.DESCRIPTION
    IntuneWinAppUtil.exe always names the encrypted inner payload "IntunePackage.intunewin"
    regardless of the source or output filename. The IntuneWin32App module reads that name from
    Detection.xml and uses it as the filename that appears in the Intune portal.

    This function rewrites the .intunewin ZIP to:
      - Rename IntuneWinPackage/Contents/IntunePackage.intunewin → IntuneWinPackage/Contents/<DesiredName>
      - Update <FileName> in IntuneWinPackage/Metadata/Detection.xml to match

    Call this after renaming the outer .intunewin file so that Intune shows the correct name.

.PARAMETER IntunewinPath
    Full path to the .intunewin file to patch (modified in place).

.PARAMETER DesiredName
    The filename to set inside the ZIP, e.g. "MyApp_1.0_PSADT.intunewin".
    Usually the leaf name of the renamed outer file.
#>
function Update-IntunewinPackageName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IntunewinPath,

        [Parameter(Mandatory)]
        [string]$DesiredName
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempPath = $IntunewinPath + '.patching'

    $srcStream = $null
    $srcZip    = $null
    $dstStream = $null
    $dstZip    = $null

    try {
        $srcStream = [System.IO.File]::OpenRead($IntunewinPath)
        $srcZip    = [System.IO.Compression.ZipArchive]::new($srcStream, [System.IO.Compression.ZipArchiveMode]::Read)
        $dstStream = [System.IO.File]::Create($tempPath)
        $dstZip    = [System.IO.Compression.ZipArchive]::new($dstStream, [System.IO.Compression.ZipArchiveMode]::Create)

        foreach ($srcEntry in $srcZip.Entries) {
            # Map the source entry name to the destination entry name
            $dstName = $srcEntry.FullName
            if ($dstName -eq 'IntuneWinPackage/Contents/IntunePackage.intunewin') {
                $dstName = "IntuneWinPackage/Contents/$DesiredName"
            }

            $dstEntry = $dstZip.CreateEntry($dstName, [System.IO.Compression.CompressionLevel]::NoCompression)
            $dstEntry.LastWriteTime = $srcEntry.LastWriteTime

            $inStream  = $srcEntry.Open()
            $outStream = $dstEntry.Open()

            if ($srcEntry.FullName -eq 'IntuneWinPackage/Metadata/Detection.xml') {
                # Patch the FileName element so the Intune portal shows the correct name
                $reader  = [System.IO.StreamReader]::new($inStream, [System.Text.Encoding]::UTF8)
                $xml     = $reader.ReadToEnd()
                $xml     = $xml -replace '<FileName>[^<]*</FileName>', "<FileName>$DesiredName</FileName>"
                $bytes   = [System.Text.Encoding]::UTF8.GetBytes($xml)
                $outStream.Write($bytes, 0, $bytes.Length)
            }
            else {
                $inStream.CopyTo($outStream)
            }

            $outStream.Dispose()
            $inStream.Dispose()
        }

        $dstZip.Dispose();    $dstZip    = $null
        $dstStream.Dispose(); $dstStream = $null
        $srcZip.Dispose();    $srcZip    = $null
        $srcStream.Dispose(); $srcStream = $null

        Remove-Item $IntunewinPath -Force
        Move-Item   $tempPath      $IntunewinPath

        Write-Verbose "Update-IntunewinPackageName: inner filename updated to '$DesiredName'."
    }
    catch {
        if ($dstZip)    { try { $dstZip.Dispose()    } catch {} }
        if ($dstStream) { try { $dstStream.Dispose() } catch {} }
        if ($srcZip)    { try { $srcZip.Dispose()    } catch {} }
        if ($srcStream) { try { $srcStream.Dispose() } catch {} }
        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        throw "Update-IntunewinPackageName: failed to patch '$IntunewinPath' — $_"
    }
}
