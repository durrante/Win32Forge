# Win32Forge v1.1.0  |  https://github.com/durrante/Win32Forge  |  MIT  |  Release history: CHANGELOG.md
<#
.SYNOPSIS
    Reads an optional per-app metadata file from the root of a source folder.

.DESCRIPTION
    Win32Forge auto-detects a metadata file (root of the source folder only, like the
    logo and detection-script auto-scan) and maps its contents onto the app fields:

        Description   -> Description
        URL           -> Information URL   (aliases: Info URL, Information URL)
        Privacy URL   -> Privacy URL        (alias:  Privacy)
        Category      -> Categories         (alias:  Categories; comma/line separated)

    Two file formats are supported (checked in this order):

      metadata.json — a flat object, e.g.
          {
            "Description": "…",
            "URL": "https://…",
            "PrivacyURL": "https://…",
            "Categories": ["Productivity", "PDF"]
          }

      metadata.txt — a section file. A line that is exactly a known header (optionally
        followed by a colon) starts a section; every following line belongs to that
        section until the next header. Example:

          Description:
          **My App** is a …
          ## Key Features
          * …

          URL:
          https://vendor.com/product

          Privacy URL:
          https://vendor.com/privacy

          Category:
          Productivity, Utilities

        Only known headers start a section, so a stray "Note:" or "Licensing note:"
        line inside the description is treated as description text, not a new section.

    Returns a PSCustomObject with Description / InformationURL / PrivacyURL / Categories
    / SourceFile, or $null when no metadata file exists or it contained nothing usable.
#>

function Get-AppMetadata {
    [CmdletBinding()]
    [OutputType([psobject])]
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder
    )

    if (-not $SourceFolder -or -not (Test-Path $SourceFolder -PathType Container)) { return $null }

    # Normalise any value into a clean string[] of categories (array, or comma/semicolon/newline separated).
    function ConvertTo-CategoryList {
        param($Value)
        if ($null -eq $Value) { return @() }
        $items = if ($Value -is [System.Array]) { $Value } else { ([string]$Value) -split '[,;\r\n]' }
        return @($items | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' })
    }

    $result = [ordered]@{
        Description    = $null
        InformationURL = $null
        PrivacyURL     = $null
        Categories     = @()
        SourceFile     = $null
    }

    $jsonFile = Join-Path $SourceFolder 'metadata.json'
    $txtFile  = Join-Path $SourceFolder 'metadata.txt'

    if (Test-Path $jsonFile -PathType Leaf) {
        try {
            $j = Get-Content -LiteralPath $jsonFile -Raw -ErrorAction Stop | ConvertFrom-Json
            $desc = $j.Description ?? $j.description
            $info = $j.InformationURL ?? $j.InfoURL ?? $j.URL ?? $j.Url ?? $j.url ?? $j.informationUrl
            $priv = $j.PrivacyURL ?? $j.Privacy ?? $j.privacyUrl ?? $j.privacy
            $cats = $j.Categories ?? $j.Category ?? $j.category ?? $j.categories

            if ($desc) { $result.Description    = ([string]$desc).Trim() }
            if ($info) { $result.InformationURL = ([string]$info).Trim() }
            if ($priv) { $result.PrivacyURL     = ([string]$priv).Trim() }
            $result.Categories = ConvertTo-CategoryList $cats
            $result.SourceFile = $jsonFile
        }
        catch {
            Write-ToolLog "Get-AppMetadata: failed to parse '$jsonFile' — $($_.Exception.Message)" -Level WARN
            return $null
        }
    }
    elseif (Test-Path $txtFile -PathType Leaf) {
        try {
            $lines = Get-Content -LiteralPath $txtFile -ErrorAction Stop
        }
        catch {
            Write-ToolLog "Get-AppMetadata: failed to read '$txtFile' — $($_.Exception.Message)" -Level WARN
            return $null
        }

        # Map a (lowercased, trimmed) header label to a canonical bucket name.
        $headerMap = @{
            'description'        = 'Description'
            'url'                = 'InformationURL'
            'info url'           = 'InformationURL'
            'information url'    = 'InformationURL'
            'infourl'            = 'InformationURL'
            'informationurl'     = 'InformationURL'
            'privacy url'        = 'PrivacyURL'
            'privacy'            = 'PrivacyURL'
            'privacyurl'         = 'PrivacyURL'
            'category'           = 'Categories'
            'categories'         = 'Categories'
        }

        $buckets = @{}
        $current = $null
        foreach ($line in $lines) {
            # A header line is "<label>" or "<label>: <optional inline value>" where <label>
            # contains only letters/spaces and resolves to a known field.
            $m = [regex]::Match($line.Trim(), '^(?<label>[A-Za-z][A-Za-z ]*?)\s*:\s*(?<rest>.*)$')
            $bucket = if ($m.Success) { $headerMap[$m.Groups['label'].Value.Trim().ToLower()] } else { $null }

            if ($bucket) {
                $current = $bucket
                if (-not $buckets.ContainsKey($current)) {
                    $buckets[$current] = [System.Collections.Generic.List[string]]::new()
                }
                $rest = $m.Groups['rest'].Value
                if ($rest.Trim() -ne '') { $buckets[$current].Add($rest) }
            }
            elseif ($current) {
                $buckets[$current].Add($line)
            }
            # Lines before the first known header are ignored.
        }

        # Trim leading/trailing blank lines from a bucket and join into one block.
        function Join-Block {
            param([System.Collections.Generic.List[string]]$Lines)
            if (-not $Lines) { return $null }
            $arr = @($Lines)
            $s = 0; $e = $arr.Count - 1
            while ($s -le $e -and [string]::IsNullOrWhiteSpace($arr[$s])) { $s++ }
            while ($e -ge $s -and [string]::IsNullOrWhiteSpace($arr[$e])) { $e-- }
            if ($s -gt $e) { return $null }
            return ($arr[$s..$e] -join "`r`n")
        }

        if ($buckets.ContainsKey('Description')) { $result.Description = Join-Block $buckets['Description'] }

        if ($buckets.ContainsKey('InformationURL')) {
            $result.InformationURL = @($buckets['InformationURL'] | ForEach-Object { $_.Trim() } | Where-Object { $_ }) | Select-Object -First 1
        }
        if ($buckets.ContainsKey('PrivacyURL')) {
            $result.PrivacyURL = @($buckets['PrivacyURL'] | ForEach-Object { $_.Trim() } | Where-Object { $_ }) | Select-Object -First 1
        }
        if ($buckets.ContainsKey('Categories')) {
            $result.Categories = ConvertTo-CategoryList (@($buckets['Categories']) -join "`n")
        }

        $result.SourceFile = $txtFile
    }
    else {
        return $null
    }

    # Nothing usable parsed — behave as if there were no metadata file.
    if (-not $result.Description -and -not $result.InformationURL -and
        -not $result.PrivacyURL -and @($result.Categories).Count -eq 0) {
        return $null
    }

    return [pscustomobject]$result
}
