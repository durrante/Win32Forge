# Win32Forge v1.1.0  |  https://github.com/durrante/Win32Forge  |  MIT  |  Release history: CHANGELOG.md
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

    $result = Invoke-PackageExe -Source $SourceFolder

    # IntuneWinAppUtil.exe is a .NET Framework tool with the classic 260-char MAX_PATH limit.
    # A deep source tree (e.g. a PSADT payload under a long OneDrive path) can exceed it and fail
    # with "DirectoryNotFoundException: Could not find a part of the path". When we see that
    # signature, retry once via a short directory junction so the paths the tool opens are short.
    $looksLikeLongPath = ($result.ExitCode -ne 0) -and
        (("$($result.Stdout)`n$($result.Stderr)") -match 'Could not find a part of the path|DirectoryNotFoundException|PathTooLong|filename or extension is too long')
    if ($looksLikeLongPath) {
        $junctionRoot = Join-Path $env:SystemDrive 'W32F'
        $junction     = Join-Path $junctionRoot ([guid]::NewGuid().ToString('N').Substring(0, 8))
        $madeJunction = $false
        try {
            New-Item -ItemType Directory -Path $junctionRoot -Force -ErrorAction Stop | Out-Null
            New-Item -ItemType Junction -Path $junction -Target $SourceFolder -ErrorAction Stop | Out-Null
            $madeJunction = $true
            Write-Host "  [!] Source path exceeds the 260-char limit — retrying via short path ($junction)..." -ForegroundColor Yellow
            Write-ToolLog "Long-path failure detected; retrying package via junction '$junction' -> '$SourceFolder'" -Level WARN
            $result = Invoke-PackageExe -Source $junction
        }
        catch {
            Write-ToolLog "Could not create short-path junction for retry — $($_.Exception.Message)" -Level ERROR
        }
        finally {
            # IMPORTANT: delete the reparse point only (non-recursive) so the real source is untouched.
            if ($madeJunction) {
                try { [System.IO.Directory]::Delete($junction, $false) }
                catch { Write-ToolLog "Could not remove junction '$junction' — $($_.Exception.Message)" -Level WARN }
            }
        }
    }

    Write-ToolLog "IntuneWinAppUtil.exe exited: code=$($result.ExitCode)" -Level DEBUG
    if ($result.Stdout) { Write-ToolLog "  stdout: $($result.Stdout)" -Level DEBUG }
    if ($result.Stderr) { Write-ToolLog "  stderr: $($result.Stderr)" -Level $(if ($result.ExitCode -ne 0) { 'ERROR' } else { 'WARN' }) }

    if ($result.ExitCode -ne 0) {
        $errText = if ($result.Stderr) { $result.Stderr } else { $result.Stdout }
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
