<#
    Convert-Heic.ps1
    - Verifies ImageMagick (magick) availability + version (+ path)
    - Checks HEIC support
    - Converts all .heic to PNG (iCCP-safe) and JPG
    - Creates an incremented logfile: heic_convertor1.log, heic_convertor2.log, ...
#>

[CmdletBinding()]
param(
  [switch]$Recurse,
  [int]$JpgQuality = 92
)

$ErrorActionPreference = 'Continue'
$script:LogPath = $null

function Get-NextLogPath {
    param(
        [string]$BaseName = "heic_convertor",
        [string]$Extension = ".log",
        [string]$Directory = (Get-Location).Path
    )
    for ($i = 1; $i -lt 100000; $i++) {
        $candidate = Join-Path $Directory ("{0}{1}{2}" -f $BaseName, $i, $Extension)
        if (-not (Test-Path $candidate)) {
            return $candidate
        }
    }
    throw "Could not find an available log filename."
}

function Initialize-Log {
    $script:LogPath = Get-NextLogPath
    # Create file with UTF-8 encoding and a header
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
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
    if (-not $NoHost) { Write-Host $line }
    Add-Content -Path $script:LogPath -Value $line -Encoding utf8
}

Initialize-Log
Write-Log "=== HEIC → PNG + JPG Batch Converter ==="
Write-Log "Parameters: Recurse=$Recurse, JpgQuality=$JpgQuality"

# --- Find ImageMagick (magick) ---
$magickCmd = $null
try {
  $magickCmd = Get-Command magick -ErrorAction SilentlyContinue
  if (-not $magickCmd) {
    $magickCmd = Get-Command magick.exe -ErrorAction SilentlyContinue
  }
} catch { }

if (-not $magickCmd) {
  Write-Log "ImageMagick 'magick' was not found on PATH." 'ERROR'
  Write-Log "Install with: winget install --id ImageMagick.ImageMagick -e" 'INFO'
  Write-Log "Then reopen PowerShell and rerun this script." 'INFO'
  exit 1
}

Write-Log "Magick path: $($magickCmd.Source)"
# --- Show version ---
try {
  $versionOutput = & $magickCmd.Source -version 2>&1
  if ($versionOutput) {
    Write-Log "Magick version output:" 'INFO'
    ($versionOutput -split "`r?`n") | ForEach-Object { Write-Log $_ -NoHost }  # log only (avoid cluttering host too much)
    # Also write a single line to host summarizing:
    $verLine = ($versionOutput -split "`r?`n" | Select-Object -First 1)
    if ($verLine) { Write-Log "Detected: $verLine" }
  }
} catch {
  Write-Log "Could not read ImageMagick version: $_" 'WARN'
}

# --- Check HEIC support ---
$heicSupported = $false
try {
  $fmtList = & $magickCmd.Source identify -list format 2>$null
  if ($fmtList) {
    $heicSupported = [bool](($fmtList | Select-String -SimpleMatch 'HEIC'))
  }
} catch {
  Write-Log "identify -list format failed: $_" 'WARN'
}

if ($heicSupported) {
  Write-Log "HEIC support: PRESENT"
} else {
  Write-Log "HEIC support: NOT DETECTED (install 'HEIF Image Extensions' from Microsoft Store, then restart PowerShell)." 'WARN'
}

# --- Collect files ---
$gciParams = @{ Filter = '*.heic' }
if ($Recurse) { $gciParams.Recurse = $true }

$files = Get-ChildItem @gciParams
if (-not $files) {
  Write-Log "No .heic files found in: $((Get-Location).Path)" 'WARN'
  Write-Log "Log saved to: $script:LogPath"
  exit 0
}

# --- Counters ---
[int]$pngOk = 0
[int]$pngFail = 0
[int]$jpgOk = 0
[int]$jpgFail = 0

Write-Log "Found $($files.Count) .heic file(s). Starting conversion…"

foreach ($f in $files) {
  $in   = $f.FullName
  $dir  = $f.DirectoryName
  $base = $f.BaseName
  $png  = Join-Path $dir "$base.png"
  $jpg  = Join-Path $dir "$base.jpg"

  Write-Log "Processing: $in"

  # --- PNG (exclude iCCP to avoid 'Incorrect data in iCCP' errors) ---
  & $magickCmd.Source "$in" -strip -colorspace sRGB -define png:exclude-chunk=iCCP "$png"
  if ($LASTEXITCODE -eq 0 -and (Test-Path "$png")) {
    $pngOk++
    Write-Log "✓ PNG created: $png"
  } else {
    $pngFail++
    Write-Log "! PNG conversion failed for: $in" 'WARN'
  }

  # --- JPG (always create) ---
  & $magickCmd.Source "$in" -strip -colorspace sRGB -quality $JpgQuality "$jpg"
  if ($LASTEXITCODE -eq 0 -and (Test-Path "$jpg")) {
    $jpgOk++
    Write-Log "✓ JPG created: $jpg"
  } else {
    $jpgFail++
    Write-Log "! JPG conversion failed for: $in" 'WARN'
  }
}

Write-Log "=== Summary ==="
Write-Log "PNG: $pngOk ok, $pngFail failed"
Write-Log "JPG: $jpgOk ok, $jpgFail failed (quality=$JpgQuality)"
Write-Log "Log saved to: $script:LogPath"
Write-Log ("Elapsed: {0}s" -f [int]([System.Diagnostics.Stopwatch]::StartNew().Elapsed.TotalSeconds)) -NoHost
Write-Log "Done."
