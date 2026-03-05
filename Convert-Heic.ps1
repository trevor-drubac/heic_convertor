<#
    Convert-Heic.ps1
    - Subtle color scheme + clear section dividers
    - Start prompt: (Y) run all, (N) exit, (C) step-by-step per section
    - Optional interactive picker (-Interactive) for Format & Quality via WinForms GUI
    - Supports -WhatIf and -DryRun
    - Auto-install ImageMagick via winget on Windows when missing (A/M prompt)
    - Overwrite protection: ask user (overwrite all / skip all / ask per file)
    - File integrity check before conversion
    - Parallel conversion via -ThrottleLimit (requires PowerShell 7+)
    - Exit code propagation
    - Incremented log (heic_convertorN.log) written to target root
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
  [Parameter(Position = 0)]
  [string]$Path = ".",
  [switch]$Recurse,
  [ValidateSet('PNG', 'JPG', 'Both')]
  [string]$Format = 'Both',
  [ValidateRange(1, 100)]
  [int]$Quality = 92,
  [switch]$DryRun,
  [switch]$Interactive,
  [ValidateRange(1, 32)]
  [int]$ThrottleLimit = 1
)

$ErrorActionPreference = 'Stop'
$script:ExitCode = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ---------------------------------------------------------------------------
# PowerShell version check
# Runs immediately so every subsequent code path can branch on $script:PS7Plus.
#
# PS 7+ path  - parallel conversion (ForEach-Object -Parallel), full feature set
# PS 5.1 path - sequential conversion only, all other features identical
# ---------------------------------------------------------------------------
$script:PS7Plus     = $PSVersionTable.PSVersion.Major -ge 7
$script:UseParallel = $script:PS7Plus -and ($ThrottleLimit -gt 1)

if ($script:PS7Plus) {
  Write-Host ("PowerShell {0} detected -- PS 7+ path: parallel conversion available (ThrottleLimit={1})." `
    -f $PSVersionTable.PSVersion, $ThrottleLimit) -ForegroundColor DarkCyan
} else {
  Write-Host ("PowerShell {0} detected -- PS 5.1 path: sequential conversion only." `
    -f $PSVersionTable.PSVersion) -ForegroundColor DarkCyan
  if ($ThrottleLimit -gt 1) {
    Write-Host ("  -ThrottleLimit {0} ignored: parallel processing requires PowerShell 7+." `
      -f $ThrottleLimit) -ForegroundColor DarkYellow
  }
}

# ---------------------------------------------------------------------------
# Cross-version Windows detector
# ---------------------------------------------------------------------------
function Test-IsWindows {
  try { if ($PSVersionTable.PSEdition -eq 'Desktop') { return $true } } catch {}
  try { if ($IsWindows) { return $true } } catch {}
  if ($env:OS -eq 'Windows_NT') { return $true }
  if ([System.Environment]::OSVersion.Platform -eq 'Win32NT') { return $true }
  return $false
}

# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------
function Write-Color {
  param(
    [Parameter(Mandatory)] [string]$Text,
    [ValidateSet('Info', 'Warn', 'Error', 'Success', 'Header', 'Action', 'Path', 'Dry')]
    [string]$Style = 'Info',
    [switch]$NoNewLine
  )
  $map = @{
    Info    = 'White'
    Warn    = 'DarkYellow'
    Error   = 'Red'
    Success = 'Green'
    Header  = 'DarkCyan'
    Action  = 'Cyan'
    Path    = 'Gray'
    Dry     = 'DarkGray'
  }
  $color = $map[$Style]
  if ($NoNewLine) { Write-Host $Text -ForegroundColor $color -NoNewline }
  else            { Write-Host $Text -ForegroundColor $color }
}

function Divider([string]$Title) {
  Write-Color ("`n------------- {0} -------------" -f $Title) -Style Header
}

function Get-Timestamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }

# ---------------------------------------------------------------------------
# Log helpers
# ---------------------------------------------------------------------------
function Get-NextLogPath([string]$Root) {
  $n = 1
  while ($true) {
    $p = Join-Path $Root ("heic_convertor{0}.log" -f $n)
    if (-not (Test-Path $p)) { return $p }
    $n++
    if ($n -gt 9999) { throw "Too many log files already present in $Root." }
  }
}

$RootFull = (Resolve-Path -LiteralPath $Path).Path
$LogPath  = Get-NextLogPath -Root $RootFull

function Write-Log {
  param(
    [string]$Message,
    [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'HEADER', 'ACTION', 'PATH', 'DRY')]
    [string]$Level = 'INFO'
  )
  $line = "[{0}] [{1}] {2}" -f (Get-Timestamp), $Level, $Message
  $line | Out-File -FilePath $LogPath -Append -Encoding utf8
  switch ($Level) {
    'INFO'    { Write-Color $Message -Style Info    }
    'WARN'    { Write-Color $Message -Style Warn    }
    'ERROR'   { Write-Color $Message -Style Error   }
    'SUCCESS' { Write-Color $Message -Style Success }
    'HEADER'  { Write-Color $Message -Style Header  }
    'ACTION'  { Write-Color $Message -Style Action  }
    'PATH'    { Write-Color $Message -Style Path    }
    'DRY'     { Write-Color $Message -Style Dry     }
  }
}

# ---------------------------------------------------------------------------
# Start-prompt logic  (Y / N / C)
# ---------------------------------------------------------------------------
$StepMode = $false

function Ask-InitialDecision {
  while ($true) {
    Write-Color "Continue? (Y = run all, N = exit, C = step-by-step)" -Style Info
    $resp = (Read-Host "[Y/N/C]").Trim().ToUpperInvariant()
    switch ($resp) {
      'Y'     { return 'ALL'  }
      'N'     { return 'EXIT' }
      'C'     { return 'STEP' }
      default { Write-Color "Please type Y, N, or C." -Style Warn }
    }
  }
}

function Ask-RunSection([string]$Title) {
  while ($true) {
    Write-Color ("Run section '{0}'? (Y = run, N = exit)" -f $Title) -Style Info
    $resp = (Read-Host "[Y/N]").Trim().ToUpperInvariant()
    switch ($resp) {
      'Y'     { return $true  }
      'N'     { return $false }
      default { Write-Color "Please type Y or N." -Style Warn }
    }
  }
}

function Run-Section([string]$Title, [ScriptBlock]$Block) {
  Divider $Title
  if ($StepMode) {
    if (-not (Ask-RunSection $Title)) {
      Write-Log "User chose to exit during section '$Title'." 'WARN'
      exit 0
    }
  }
  & $Block
}

# ---------------------------------------------------------------------------
# Interactive GUI (WinForms) + text fallback
# ---------------------------------------------------------------------------
function Show-InteractiveDialog {
  param(
    [string]$InitialFormat  = 'Both',
    [int]   $InitialQuality = 92
  )
  $result = $null
  try {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop

    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = "HEIC Converter Options"
    $form.StartPosition = "CenterScreen"
    $form.Size          = New-Object System.Drawing.Size(380, 190)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox   = $false
    $form.MinimizeBox   = $false
    $form.TopMost       = $true

    $lblFormat          = New-Object System.Windows.Forms.Label
    $lblFormat.Text     = "Output format:"
    $lblFormat.Location = New-Object System.Drawing.Point(15, 20)
    $lblFormat.AutoSize = $true

    $cmbFormat               = New-Object System.Windows.Forms.ComboBox
    $cmbFormat.DropDownStyle = "DropDownList"
    [void]$cmbFormat.Items.AddRange(@("PNG", "JPG", "Both"))
    $cmbFormat.SelectedItem  = $InitialFormat
    $cmbFormat.Location      = New-Object System.Drawing.Point(140, 16)
    $cmbFormat.Size          = New-Object System.Drawing.Size(200, 24)

    $lblQuality          = New-Object System.Windows.Forms.Label
    $lblQuality.Text     = "JPEG quality (1-100):"
    $lblQuality.Location = New-Object System.Drawing.Point(15, 55)
    $lblQuality.AutoSize = $true

    $numQuality          = New-Object System.Windows.Forms.NumericUpDown
    $numQuality.Minimum  = 1
    $numQuality.Maximum  = 100
    $numQuality.Value    = [decimal]$InitialQuality
    $numQuality.Location = New-Object System.Drawing.Point(200, 52)
    $numQuality.Size     = New-Object System.Drawing.Size(60, 24)

    $handler = { $numQuality.Enabled = ($cmbFormat.SelectedItem -ne "PNG") }
    $cmbFormat.add_SelectedIndexChanged($handler)
    & $handler

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Location     = New-Object System.Drawing.Point(160, 100)

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location     = New-Object System.Drawing.Point(245, 100)

    $form.Controls.AddRange(@($lblFormat, $cmbFormat, $lblQuality, $numQuality, $btnOK, $btnCancel))

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $result = [PSCustomObject]@{
        Format  = [string]$cmbFormat.SelectedItem
        Quality = [int]$numQuality.Value
      }
    }
  } catch {
    Write-Log "WinForms UI not available: $($_.Exception.Message) - falling back to text prompts." 'WARN'
  }
  return $result
}

function Prompt-ForOptions([string]$CurrentFormat, [int]$CurrentQuality) {
  $f = Read-Host "Output format (PNG/JPG/Both) [default: $CurrentFormat]"
  if ([string]::IsNullOrWhiteSpace($f)) { $f = $CurrentFormat }
  $f = $f.ToUpperInvariant()
  if ($f -notin @('PNG', 'JPG', 'BOTH')) {
    Write-Log "Invalid format '$f', using '$CurrentFormat'." 'WARN'
    $f = $CurrentFormat
  }

  $q      = Read-Host "JPEG quality 1-100 [default: $CurrentQuality]"
  $parsed = 0
  if ([string]::IsNullOrWhiteSpace($q)) {
    $parsed = $CurrentQuality
  } elseif (-not ([int]::TryParse($q, [ref]$parsed)) -or $parsed -lt 1 -or $parsed -gt 100) {
    Write-Log "Invalid quality '$q', using $CurrentQuality." 'WARN'
    $parsed = $CurrentQuality
  }
  return [PSCustomObject]@{ Format = $f; Quality = $parsed }
}

# ---------------------------------------------------------------------------
# Initial prompt
# ---------------------------------------------------------------------------
Divider "Start"
$decision = Ask-InitialDecision
switch ($decision) {
  'EXIT' { Write-Log "User chose to exit." 'WARN'; exit 0 }
  'STEP' { $StepMode = $true; Write-Log "Step-by-step mode enabled." 'INFO' }
  default { Write-Log "Running all sections." 'INFO' }
}

# ---------------------------------------------------------------------------
# Section: Startup & Options
# ---------------------------------------------------------------------------
Run-Section "Startup & Options" {
  if ($Interactive) {
    $sel = Show-InteractiveDialog -InitialFormat $Format -InitialQuality $Quality
    if (-not $sel) { $sel = Prompt-ForOptions -CurrentFormat $Format -CurrentQuality $Quality }
    if ($sel) { $script:Format = $sel.Format; $script:Quality = $sel.Quality }
  } else {
    $script:Format  = $Format
    $script:Quality = $Quality
  }

  Write-Log "HEIC -> PNG/JPG Converter (PowerShell)" 'HEADER'
  Write-Log ("PowerShell: {0} | Path: {1}" -f $PSVersionTable.PSVersion, $(if ($script:PS7Plus) { 'PS 7+  (parallel capable)' } else { 'PS 5.1 (sequential only)' })) 'INFO'
  Write-Log ("Root      : {0}" -f $RootFull) 'PATH'
  Write-Log ("Recurse   : {0}" -f $Recurse)  'INFO'
  Write-Log ("Format    : {0}" -f $script:Format)  'INFO'
  Write-Log ("Quality   : {0}" -f $script:Quality) 'INFO'
  Write-Log ("DryRun    : {0}  WhatIf: {1}" -f $DryRun, ($WhatIfPreference -eq $true)) 'DRY'
  $script:StartTime = Get-Date
}

# ---------------------------------------------------------------------------
# Section: ImageMagick Setup
# ---------------------------------------------------------------------------
Run-Section "ImageMagick Setup" {
  function Get-Magick {
    foreach ($name in @('magick.exe', 'magick')) {
      $cmd = Get-Command $name -ErrorAction SilentlyContinue
      if ($cmd) { return $cmd.Path }
    }
    $paths = @()
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
      if ($root) {
        $paths += Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match '^ImageMagick' } |
                  ForEach-Object { Join-Path $_.FullName 'magick.exe' }
      }
    }
    $paths += '/usr/local/bin/magick', '/usr/bin/magick'
    foreach ($p in $paths) { if (Test-Path $p) { return $p } }
    return $null
  }

  function Prompt-InstallMagick {
    while ($true) {
      Write-Color "ImageMagick not found. (A) auto-install via winget, (M) manual install?" -Style Info
      $resp = (Read-Host "[A/M]").Trim().ToUpperInvariant()
      switch ($resp) {
        'A' {
          if (-not (Test-IsWindows)) {
            Write-Log "Auto-install is only supported on Windows. Choose (M) for manual." 'WARN'
            continue
          }
          $wg = Get-Command winget -ErrorAction SilentlyContinue
          if (-not $wg) {
            Write-Log "winget not available. Install 'App Installer' from the Microsoft Store, then retry." 'ERROR'
            continue
          }
          if ($DryRun -or $WhatIfPreference) {
            Write-Log "(dry) Would run: winget install ImageMagick.Q16-HDRI" 'DRY'
            return $false
          }
          Write-Log "Running: winget install ImageMagick.Q16-HDRI" 'ACTION'
          $proc = Start-Process -FilePath $wg.Source `
                                -ArgumentList @('install', 'ImageMagick.Q16-HDRI') `
                                -Wait -PassThru -NoNewWindow
          if ($proc.ExitCode -ne 0) {
            Write-Log "winget exited with code $($proc.ExitCode). Installation may have failed." 'ERROR'
            Write-Log "Try running PowerShell as Administrator, then choose (A) again." 'WARN'
            return $false
          }
          Write-Log "Installation finished. Re-detecting 'magick'..." 'INFO'
          return $true
        }
        'M' {
          Write-Log "Download ImageMagick from https://imagemagick.org/script/download.php" 'INFO'
          [void](Read-Host "Press Enter to exit...")
          exit 1
        }
        default { Write-Color "Please type A or M." -Style Warn }
      }
    }
  }

  $script:Magick = Get-Magick
  if (-not $script:Magick) {
    Write-Log "ImageMagick 'magick' not found." 'ERROR'
    $installed = Prompt-InstallMagick
    if ($installed) {
      $script:Magick = Get-Magick
      if (-not $script:Magick) {
        Write-Log "'magick' still not found after install. Close and reopen PowerShell (PATH refresh), then rerun." 'ERROR'
        [void](Read-Host "Press Enter to exit...")
        exit 1
      }
    } elseif ($DryRun -or $WhatIfPreference) {
      Write-Log "Dry/WhatIf mode: continuing without magick." 'DRY'
    } else {
      [void](Read-Host "Press Enter to exit...")
      exit 1
    }
  }

  if ($script:Magick) {
    try {
      $ver = & $script:Magick -version 2>&1
      Write-Log ("Using: {0}" -f $script:Magick) 'PATH'
      Write-Log ("magick -version:`n{0}" -f ($ver -join [Environment]::NewLine)) 'INFO'
    } catch {
      Write-Log "Failed to run 'magick -version': $_" 'WARN'
    }
  }
}

# ---------------------------------------------------------------------------
# Section: HEIC Support Check
# ---------------------------------------------------------------------------
Run-Section "HEIC Support Check" {
  try {
    $formats = (& $script:Magick -list format 2>&1) -join "`n"
    if ($formats -match '(?i)\bHEIC\b' -or $formats -match '(?i)\bHEIF\b') {
      Write-Log "HEIC/HEIF support: PRESENT" 'SUCCESS'
    } else {
      Write-Log "HEIC/HEIF not listed. Conversions may fail (install libheif / HEIF Image Extensions from Microsoft Store)." 'WARN'
    }
  } catch {
    Write-Log "Failed to query 'magick -list format': $_ - continuing anyway." 'WARN'
  }
}

# ---------------------------------------------------------------------------
# Section: File Discovery
# ---------------------------------------------------------------------------
Run-Section "File Discovery" {
  $searchParams = @{ Path = $RootFull; Filter = '*.heic'; File = $true; ErrorAction = 'SilentlyContinue' }
  if ($Recurse) { $searchParams['Recurse'] = $true }
  $script:Files = Get-ChildItem @searchParams
  if (-not $script:Files -or $script:Files.Count -eq 0) {
    Write-Log ("No .heic files found in: {0}" -f $RootFull) 'WARN'
    Write-Log ("Log saved to: {0}" -f $LogPath) 'PATH'
    exit 0
  }
  Write-Log ("Found {0} .heic file(s)." -f $script:Files.Count) 'HEADER'
}

# ---------------------------------------------------------------------------
# Section: Overwrite Check
# ---------------------------------------------------------------------------
Run-Section "Overwrite Check" {
  $script:OverwriteAll = $false
  $script:SkipExisting = $false

  $anyExist = $script:Files | Where-Object {
    ($script:Format -in @('PNG', 'Both') -and (Test-Path (Join-Path $_.DirectoryName "$($_.BaseName).png"))) -or
    ($script:Format -in @('JPG', 'Both') -and (Test-Path (Join-Path $_.DirectoryName "$($_.BaseName).jpg")))
  }

  if ($anyExist) {
    Write-Log "Some output files already exist." 'WARN'
    do {
      Write-Color "  [O] Overwrite all existing files" -Style Info
      Write-Color "  [S] Skip files that already exist" -Style Info
      Write-Color "  [A] Ask me for each file" -Style Info
      $owChoice = (Read-Host "Your choice [O/S/A]").Trim().ToUpperInvariant()
    } while ($owChoice -notin @('O', 'S', 'A'))

    switch ($owChoice) {
      'O' { $script:OverwriteAll = $true; Write-Log "Overwrite mode: overwrite all" 'INFO' }
      'S' { $script:SkipExisting = $true; Write-Log "Overwrite mode: skip existing" 'INFO' }
      'A' {                               Write-Log "Overwrite mode: ask per file"  'INFO' }
    }
  }
}

# ---------------------------------------------------------------------------
# Helper: confirm before overwriting an existing output file
# ---------------------------------------------------------------------------
function Confirm-Overwrite([string]$OutputPath) {
  if (-not (Test-Path $OutputPath)) { return $true }
  if ($script:OverwriteAll)         { return $true }
  if ($script:SkipExisting)         { return $false }
  do {
    $ans = (Read-Host ("  '{0}' exists. Overwrite? [Y/N]" -f (Split-Path -Leaf $OutputPath))).Trim().ToUpperInvariant()
  } while ($ans -notin @('Y', 'N'))
  return ($ans -eq 'Y')
}

# ---------------------------------------------------------------------------
# Section: Conversion
# ---------------------------------------------------------------------------
Run-Section "Conversion" {
  $script:pngOk      = 0; $script:pngFail    = 0; $script:pngSkipped = 0
  $script:jpgOk      = 0; $script:jpgFail    = 0; $script:jpgSkipped = 0
  $script:intFail    = 0

  # $script:UseParallel is resolved at script start from the PS version check.
  # PS 7+ path  : ForEach-Object -Parallel with ThrottleLimit
  # PS 5.1 path : sequential ForEach-Object
  $useParallel = $script:UseParallel

  # Per-file overwrite prompting is incompatible with parallel execution.
  # If the user chose ask-per-file, switch to skip-existing automatically.
  if ($useParallel -and -not $script:OverwriteAll -and -not $script:SkipExisting) {
    Write-Log "Per-file overwrite prompting is not supported in parallel mode. Existing outputs will be skipped." 'WARN'
    $script:SkipExisting = $true
  }

  if ($useParallel) {
    Write-Log ("PS 7+ path: starting parallel conversion (ThrottleLimit={0})..." -f $ThrottleLimit) 'ACTION'
  } else {
    Write-Log ("PS {0} path: starting sequential conversion..." -f $PSVersionTable.PSVersion) 'ACTION'
  }

  # ---------------------------------------------------------------------------
  # Shared progress counter (synchronized hashtable, safe across runspaces)
  # ---------------------------------------------------------------------------
  $progress = [hashtable]::Synchronized(@{ Done = 0 })
  $total    = $script:Files.Count

  # ---------------------------------------------------------------------------
  # Scriptblock shared by both paths - runs per file, returns a result object.
  # All values come in via parameters so the block works both inline (sequential)
  # and inside ForEach-Object -Parallel (where $using: variables are needed).
  # ---------------------------------------------------------------------------
  $convertFile = {
    param(
      $File, $Magick, $Fmt, $Qual, $Dry, $OwAll, $SkipEx, $Progress, $Total
    )

    $log     = [System.Collections.Generic.List[string]]::new()
    $ts      = { "[{0}]" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
    $addLog  = { param($msg, $lvl = 'INFO')
      $line = "{0} [{1}] {2}" -f (& $ts), $lvl, $msg
      $log.Add($line)
      $styleMap = @{ INFO='White'; WARN='DarkYellow'; ERROR='Red'; SUCCESS='Green'
                     ACTION='Cyan'; PATH='Gray'; DRY='DarkGray' }
      $color = if ($styleMap.ContainsKey($lvl)) { $styleMap[$lvl] } else { 'White' }
      Write-Host $line -ForegroundColor $color
    }

    $confirmOw = { param($out)
      if (-not (Test-Path $out)) { return $true }
      if ($OwAll)  { return $true }
      if ($SkipEx) { return $false }
      return $false   # fallback: skip (parallel mode can't prompt)
    }

    $r = [PSCustomObject]@{
      LogLines   = $null
      pngOk = 0; pngFail = 0; pngSkipped = 0
      jpgOk = 0; jpgFail = 0; jpgSkipped = 0
      intFail = 0; ExitCode = 0
    }

    $done = [System.Threading.Interlocked]::Increment([ref]$Progress.Done)
    Write-Progress -Activity "Converting HEIC files" `
                   -Status   ("[$done/$Total] {0}" -f $File.Name) `
                   -PercentComplete ([int](($done - 1) / $Total * 100))

    & $addLog ("Processing [$done/$Total]: {0}" -f $File.FullName) 'PATH'

    # Integrity check
    try {
      $identOut = & $Magick identify $File.FullName 2>&1
      if ($LASTEXITCODE -ne 0) {
        & $addLog ("  Integrity FAILED (exit $LASTEXITCODE): {0}" -f ($identOut -join ' ')) 'ERROR'
        $r.intFail = 1; $r.ExitCode = 1; $r.LogLines = $log.ToArray(); return $r
      }
    } catch {
      & $addLog ("  Integrity check exception: {0}" -f $_) 'ERROR'
      $r.intFail = 1; $r.ExitCode = 1; $r.LogLines = $log.ToArray(); return $r
    }

    $png = Join-Path $File.DirectoryName ($File.BaseName + '.png')
    $jpg = Join-Path $File.DirectoryName ($File.BaseName + '.jpg')

    # PNG conversion
    if ($Fmt -in @('PNG', 'Both')) {
      if (& $confirmOw $png) {
        if ($Dry) {
          & $addLog ("  (dry) Would create PNG: {0}" -f $png) 'DRY'; $r.pngOk++
        } else {
          try {
            $out = & $Magick $File.FullName -colorspace sRGB -strip -define png:exclude-chunk=iCCP $png 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path $png)) {
              & $addLog ("  PNG ok: {0}" -f $png) 'SUCCESS'; $r.pngOk++
            } else {
              & $addLog ("  PNG FAILED (exit $LASTEXITCODE): {0}" -f ($out -join ' ')) 'ERROR'
              $r.pngFail++; $r.ExitCode = 1
            }
          } catch {
            & $addLog ("  PNG exception: {0}" -f $_) 'ERROR'; $r.pngFail++; $r.ExitCode = 1
          }
        }
      } else {
        & $addLog ("  PNG skipped (exists): {0}" -f $png) 'INFO'; $r.pngSkipped++
      }
    }

    # JPG conversion
    if ($Fmt -in @('JPG', 'Both')) {
      if (& $confirmOw $jpg) {
        if ($Dry) {
          & $addLog ("  (dry) Would create JPG: {0} (q={1})" -f $jpg, $Qual) 'DRY'; $r.jpgOk++
        } else {
          try {
            $out = & $Magick $File.FullName -colorspace sRGB -strip -quality $Qual $jpg 2>&1
            if ($LASTEXITCODE -eq 0 -and (Test-Path $jpg)) {
              & $addLog ("  JPG ok: {0}" -f $jpg) 'SUCCESS'; $r.jpgOk++
            } else {
              & $addLog ("  JPG FAILED (exit $LASTEXITCODE): {0}" -f ($out -join ' ')) 'ERROR'
              $r.jpgFail++; $r.ExitCode = 1
            }
          } catch {
            & $addLog ("  JPG exception: {0}" -f $_) 'ERROR'; $r.jpgFail++; $r.ExitCode = 1
          }
        }
      } else {
        & $addLog ("  JPG skipped (exists): {0}" -f $jpg) 'INFO'; $r.jpgSkipped++
      }
    }

    $r.LogLines = $log.ToArray()
    return $r
  }

  # ---------------------------------------------------------------------------
  # Execute: parallel or sequential
  # ---------------------------------------------------------------------------
  # NOTE: intentionally NOT named $args - that is a PS automatic variable
  # (unbound positional params). Using $args inside a scriptblock resets it,
  # which would make $args.Total = $null => [int]$null = 0 => divide by zero.
  $cbArgs = @{
    Magick   = $script:Magick
    Fmt      = $script:Format
    Qual     = $script:Quality
    Dry      = $DryRun.IsPresent
    OwAll    = $script:OverwriteAll
    SkipEx   = $script:SkipExisting
    Progress = $progress
    Total    = $total
  }

  # ---------------------------------------------------------------------------
  # PS 7+ path: parallel execution via ForEach-Object -Parallel
  # PS 5.1 path: sequential execution via ForEach-Object
  # Both paths call the same $convertFile scriptblock and return identical
  # result objects; aggregation below is path-independent.
  # ---------------------------------------------------------------------------
  if ($useParallel) {
    # --- PS 7+ path ---
    $results = $script:Files | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
      $cb   = $using:convertFile
      $a    = $using:cbArgs
      & $cb -File $_ -Magick $a.Magick -Fmt $a.Fmt -Qual $a.Qual -Dry $a.Dry `
             -OwAll $a.OwAll -SkipEx $a.SkipEx -Progress $a.Progress -Total $a.Total
    }
  } else {
    # --- PS 5.1 path ---
    $results = $script:Files | ForEach-Object {
      & $convertFile -File $_ -Magick $cbArgs.Magick -Fmt $cbArgs.Fmt -Qual $cbArgs.Qual `
                     -Dry $cbArgs.Dry -OwAll $cbArgs.OwAll -SkipEx $cbArgs.SkipEx `
                     -Progress $cbArgs.Progress -Total $cbArgs.Total
    }
  }

  Write-Progress -Activity "Converting HEIC files" -Completed

  # ---------------------------------------------------------------------------
  # Aggregate results: write collected log lines then sum counters
  # ---------------------------------------------------------------------------
  foreach ($r in $results) {
    foreach ($line in $r.LogLines) {
      $line | Out-File -FilePath $LogPath -Append -Encoding utf8
    }
    $script:pngOk      += $r.pngOk
    $script:pngFail    += $r.pngFail
    $script:pngSkipped += $r.pngSkipped
    $script:jpgOk      += $r.jpgOk
    $script:jpgFail    += $r.jpgFail
    $script:jpgSkipped += $r.jpgSkipped
    $script:intFail    += $r.intFail
    if ($r.ExitCode -ne 0) { $script:ExitCode = 1 }
  }
}

# ---------------------------------------------------------------------------
# Section: Summary
# ---------------------------------------------------------------------------
Run-Section "Summary" {
  $elapsed = [int]$stopwatch.Elapsed.TotalSeconds
  Write-Log ("Files found    : {0}" -f $script:Files.Count) 'INFO'
  Write-Log ("Integrity fails: {0}" -f $script:intFail) 'INFO'
  if ($script:Format -in @('PNG', 'Both')) {
    Write-Log ("PNG : {0} ok, {1} failed, {2} skipped" -f $script:pngOk, $script:pngFail, $script:pngSkipped) 'INFO'
  }
  if ($script:Format -in @('JPG', 'Both')) {
    Write-Log ("JPG : {0} ok, {1} failed, {2} skipped (quality={3})" -f $script:jpgOk, $script:jpgFail, $script:jpgSkipped, $script:Quality) 'INFO'
  }
  Write-Log ("Elapsed        : {0}s" -f $elapsed) 'INFO'
  Write-Log ("Log saved to   : {0}" -f $LogPath) 'PATH'
  Write-Log "Done." 'SUCCESS'
}

exit $script:ExitCode
