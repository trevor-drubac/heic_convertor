<#
    Convert-Heic.ps1
    - Verifies ImageMagick (magick) availability + version (+ path)
    - Checks HEIC support
    - Validates each HEIC file before conversion
    - Converts all .heic to PNG (iCCP-safe) and JPG
    - Prompts interactively for JPG quality and overwrite behaviour
    - Shows a progress bar during batch conversion
    - Creates an incremented logfile: heic_convertor1.log, heic_convertor2.log, ...
#>

[CmdletBinding()]
param(
  [switch]$Recurse,
  [ValidateRange(0,100)]
  [int]$JpgQuality = -1   # -1 means "not supplied — ask interactively"
)

$ErrorActionPreference = 'Stop'   # surface errors instead of silently swallowing them
$script:LogPath = $null
$script:ExitCode = 0              # track overall exit code

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-NextLogPath {
    param(
        [string]$BaseName  = "heic_convertor",
        [string]$Extension = ".log",
        [string]$Directory = (Get-Location).Path
    )
    for ($i = 1; $i -lt 100000; $i++) {
        $candidate = Join-Path $Directory ("{0}{1}{2}" -f $BaseName, $i, $Extension)
        if (-not (Test-Path $candidate)) { return $candidate }
    }
    throw "Could not find an available log filename."
}

function Initialize-Log {
    $script:LogPath = Get-NextLogPath
    $header = @(
        "=== HEIC → PNG+JPG Converter Log ==="
        "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Working Dir: $((Get-Location).Path)"
        ""
    ) -join [Environment]::NewLine
    $header | Out-File -FilePath $script:LogPath -Encoding utf8 -Force
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO',
        [switch]$NoHost
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
    if (-not $NoHost) { Write-Host $line }
    Add-Content -Path $script:LogPath -Value $line -Encoding utf8
}

# ---------------------------------------------------------------------------
# Initialise log
# ---------------------------------------------------------------------------
Initialize-Log
Write-Log "=== HEIC → PNG + JPG Batch Converter ==="

# ---------------------------------------------------------------------------
# Interactive: ask for JPG quality if not supplied via parameter
# ---------------------------------------------------------------------------
if ($JpgQuality -eq -1) {
    do {
        $raw = Read-Host "Enter JPG quality (0-100) [default: 92]"
        if ($raw -eq '') {
            $JpgQuality = 92
            break
        }
        $parsed = 0
        if ([int]::TryParse($raw, [ref]$parsed) -and $parsed -ge 0 -and $parsed -le 100) {
            $JpgQuality = $parsed
        } else {
            Write-Host "  Invalid value '$raw'. Please enter a number between 0 and 100."
        }
    } while ($JpgQuality -eq -1)
}

Write-Log "Parameters: Recurse=$Recurse, JpgQuality=$JpgQuality"

# ---------------------------------------------------------------------------
# Find ImageMagick
# ---------------------------------------------------------------------------
$magickCmd = $null
try {
    $magickCmd = Get-Command magick -ErrorAction SilentlyContinue
    if (-not $magickCmd) {
        $magickCmd = Get-Command magick.exe -ErrorAction SilentlyContinue
    }
} catch {
    Write-Log "Error while searching for 'magick' on PATH: $_" 'ERROR'
}

if (-not $magickCmd) {
    Write-Log "ImageMagick 'magick' was not found on PATH." 'ERROR'
    Write-Log "Install with: winget install --id ImageMagick.ImageMagick -e" 'INFO'
    Write-Log "Then reopen PowerShell and rerun this script." 'INFO'
    exit 1
}

Write-Log "Magick path: $($magickCmd.Source)"

# ---------------------------------------------------------------------------
# Show ImageMagick version
# ---------------------------------------------------------------------------
try {
    $versionOutput = & $magickCmd.Source -version 2>&1
    if ($versionOutput) {
        Write-Log "Magick version output:" 'INFO'
        ($versionOutput -split "`r?`n") | ForEach-Object { Write-Log $_ -NoHost }
        $verLine = ($versionOutput -split "`r?`n" | Select-Object -First 1)
        if ($verLine) { Write-Log "Detected: $verLine" }
    }
} catch {
    Write-Log "Could not read ImageMagick version: $_" 'WARN'
}

# ---------------------------------------------------------------------------
# Check HEIC support
# ---------------------------------------------------------------------------
$heicSupported = $false
try {
    $fmtList = & $magickCmd.Source identify -list format 2>&1
    if ($fmtList) {
        $heicSupported = [bool]($fmtList | Select-String -SimpleMatch 'HEIC')
    } else {
        Write-Log "'identify -list format' returned no output — HEIC support status unknown." 'WARN'
    }
} catch {
    Write-Log "identify -list format failed: $_ — continuing anyway, conversion may still succeed." 'WARN'
}

if ($heicSupported) {
    Write-Log "HEIC support: PRESENT"
} else {
    Write-Log "HEIC support: NOT DETECTED (install 'HEIF Image Extensions' from Microsoft Store, then restart PowerShell)." 'WARN'
}

# ---------------------------------------------------------------------------
# Collect HEIC files
# ---------------------------------------------------------------------------
$gciParams = @{ Filter = '*.heic' }
if ($Recurse) { $gciParams.Recurse = $true }

$files = Get-ChildItem @gciParams
if (-not $files) {
    Write-Log "No .heic files found in: $((Get-Location).Path)" 'WARN'
    Write-Log "Log saved to: $script:LogPath"
    exit 0
}

Write-Log "Found $($files.Count) .heic file(s)."

# ---------------------------------------------------------------------------
# Interactive: overwrite protection
# ---------------------------------------------------------------------------
$overwriteAll  = $false
$skipExisting  = $false

$anyExist = $files | Where-Object {
    (Test-Path (Join-Path $_.DirectoryName "$($_.BaseName).png")) -or
    (Test-Path (Join-Path $_.DirectoryName "$($_.BaseName).jpg"))
}

if ($anyExist) {
    Write-Host ""
    Write-Host "Some output files (.png / .jpg) already exist."
    do {
        Write-Host "  [O] Overwrite all existing files"
        Write-Host "  [S] Skip files that already exist"
        Write-Host "  [A] Ask me for each file"
        $owChoice = (Read-Host "Your choice [O/S/A]").Trim().ToUpper()
    } while ($owChoice -notin @('O','S','A'))

    switch ($owChoice) {
        'O' { $overwriteAll = $true;  Write-Log "Overwrite mode: overwrite all"  }
        'S' { $skipExisting = $true;  Write-Log "Overwrite mode: skip existing"  }
        'A' {                         Write-Log "Overwrite mode: ask per file"   }
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Helper: decide whether to write a specific output path
# Returns $true if the conversion should proceed, $false to skip.
# ---------------------------------------------------------------------------
function Confirm-Overwrite {
    param([string]$OutputPath)
    if (-not (Test-Path $OutputPath)) { return $true }  # file does not exist — always proceed
    if ($overwriteAll)                { return $true }
    if ($skipExisting)                { return $false }
    # Per-file prompt
    do {
        $ans = (Read-Host "  '$OutputPath' exists. Overwrite? [Y/N]").Trim().ToUpper()
    } while ($ans -notin @('Y','N'))
    return ($ans -eq 'Y')
}

# ---------------------------------------------------------------------------
# Counters
# ---------------------------------------------------------------------------
[int]$pngOk      = 0
[int]$pngFail    = 0
[int]$pngSkipped = 0
[int]$jpgOk      = 0
[int]$jpgFail    = 0
[int]$jpgSkipped = 0
[int]$integrityFail = 0
[int]$current    = 0

Write-Log "Starting conversion…"

foreach ($f in $files) {
    $current++
    $in   = $f.FullName
    $dir  = $f.DirectoryName
    $base = $f.BaseName
    $png  = Join-Path $dir "$base.png"
    $jpg  = Join-Path $dir "$base.jpg"

    # Progress bar
    Write-Progress `
        -Activity "Converting HEIC files" `
        -Status   "[$current/$($files.Count)] $($f.Name)" `
        -PercentComplete ([int](($current - 1) / $files.Count * 100))

    Write-Log "Processing ($current/$($files.Count)): $in"

    # -----------------------------------------------------------------------
    # File integrity check — verify the HEIC is readable by ImageMagick
    # -----------------------------------------------------------------------
    try {
        $identifyOut = & $magickCmd.Source identify "$in" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Integrity check FAILED for '$in' (exit $LASTEXITCODE): $identifyOut" 'ERROR'
            $integrityFail++
            $script:ExitCode = 1
            continue   # skip this file entirely
        }
    } catch {
        Write-Log "Integrity check threw an exception for '$in': $_" 'ERROR'
        $integrityFail++
        $script:ExitCode = 1
        continue
    }

    # -----------------------------------------------------------------------
    # PNG conversion
    # -----------------------------------------------------------------------
    if (Confirm-Overwrite $png) {
        try {
            $pngOut = & $magickCmd.Source "$in" -strip -colorspace sRGB -define png:exclude-chunk=iCCP "$png" 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path $png)) {
                $pngOk++
                Write-Log "  PNG ok: $png"
            } else {
                $pngFail++
                $script:ExitCode = 1
                Write-Log "  PNG FAILED (exit $LASTEXITCODE): $pngOut" 'ERROR'
            }
        } catch {
            $pngFail++
            $script:ExitCode = 1
            Write-Log "  PNG conversion threw an exception: $_" 'ERROR'
        }
    } else {
        $pngSkipped++
        Write-Log "  PNG skipped (already exists): $png"
    }

    # -----------------------------------------------------------------------
    # JPG conversion
    # -----------------------------------------------------------------------
    if (Confirm-Overwrite $jpg) {
        try {
            $jpgOut = & $magickCmd.Source "$in" -strip -colorspace sRGB -quality $JpgQuality "$jpg" 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path $jpg)) {
                $jpgOk++
                Write-Log "  JPG ok: $jpg"
            } else {
                $jpgFail++
                $script:ExitCode = 1
                Write-Log "  JPG FAILED (exit $LASTEXITCODE): $jpgOut" 'ERROR'
            }
        } catch {
            $jpgFail++
            $script:ExitCode = 1
            Write-Log "  JPG conversion threw an exception: $_" 'ERROR'
        }
    } else {
        $jpgSkipped++
        Write-Log "  JPG skipped (already exists): $jpg"
    }
}

Write-Progress -Activity "Converting HEIC files" -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$elapsed = [int]$stopwatch.Elapsed.TotalSeconds
Write-Log "=== Summary ==="
Write-Log "Files processed : $($files.Count)"
Write-Log "Integrity fails : $integrityFail"
Write-Log "PNG  : $pngOk ok, $pngFail failed, $pngSkipped skipped"
Write-Log "JPG  : $jpgOk ok, $jpgFail failed, $jpgSkipped skipped (quality=$JpgQuality)"
Write-Log "Elapsed: ${elapsed}s"
Write-Log "Log saved to: $script:LogPath"
Write-Log "Done."

exit $script:ExitCode
