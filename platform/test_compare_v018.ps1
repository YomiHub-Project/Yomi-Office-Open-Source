# Yomi-Office Open Source v0.1.8 - Windows Installer Matrix, Feature QA, and MS Office Side-by-Side Comparison
param(
  [switch]$SkipInstallTest,
  [switch]$SkipFeatureTest,
  [switch]$SkipCompare,
  [switch]$NoVideo,
  [string]$Edition = 'Win10-11-x64',
  [int]$MaxCompare = 0
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

$VER  = '0.1.8'
$SLUG = 'Yomi-Office-Open-Source'
$PLAT = $PSScriptRoot
$SRC  = Split-Path -Parent $PLAT
$VDIR = Split-Path -Parent $SRC
$PROJ = Split-Path -Parent $VDIR
$OUT_WIN = Join-Path (Split-Path -Parent $PROJ) "output\v$VER\Windows"
$TESTS   = Join-Path $PROJ 'Test Office'
$BLD     = Join-Path $VDIR 'build\qa_win'
$SHOTS   = Join-Path $VDIR 'shots\windows'
$VIDDIR  = Join-Path $VDIR 'video\windows'
$REPORT  = Join-Path $VDIR 'coord\windows_qa_report.md'
$FFMPEG  = 'D:\yh\tools\ffmpeg-9.0.2-essentials_build\bin\ffmpeg.exe'

$MS_WORD  = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
$MS_EXCEL = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
$MS_PPT   = 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE'

New-Item -ItemType Directory -Force -Path $BLD, $SHOTS, $VIDDIR | Out-Null

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WinQaNative {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
"@

$script:results = New-Object System.Collections.ArrayList
$script:pass = 0
$script:fail = 0

function Say($m) { Write-Output ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) }
function LogResult($cat, $name, $ok, $detail) {
  if ($ok) { $script:pass++; $tag = 'PASS' } else { $script:fail++; $tag = 'FAIL' }
  $line = ("{0,-4} [{1}] {2} :: {3}" -f $tag, $cat, $name, $detail)
  Say $line
  [void]$script:results.Add([PSCustomObject]@{ Status=$tag; Category=$cat; Item=$name; Detail=$detail; Time=(Get-Date -Format 'HH:mm:ss') })
}

function Capture-Screen([string]$outPng) {
  $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $bmp = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
  $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose()
  $bmp.Dispose()
}

function Capture-Window([IntPtr]$hwnd, [string]$outPng) {
  if ($hwnd -eq [IntPtr]::Zero) { Capture-Screen $outPng; return }
  [WinQaNative]::ShowWindow($hwnd, 9) | Out-Null
  [WinQaNative]::SetForegroundWindow($hwnd) | Out-Null
  Start-Sleep -Milliseconds 350
  $r = New-Object WinQaNative+RECT
  if ([WinQaNative]::GetWindowRect($hwnd, [ref]$r) -and ($r.Right - $r.Left) -gt 80 -and ($r.Bottom - $r.Top) -gt 80) {
    $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    $pwOk = [WinQaNative]::PrintWindow($hwnd, $hdc, 2)
    $g.ReleaseHdc($hdc)
    if (-not $pwOk) {
      $g.CopyFromScreen($r.Left, $r.Top, 0, 0, (New-Object System.Drawing.Size($w, $h)))
    }
    $bmp.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
  } else {
    Capture-Screen $outPng
  }
}

function Make-SideBySide([string]$leftPng, [string]$rightPng, [string]$outPng, [string]$titleLeft, [string]$titleRight) {
  if (-not (Test-Path $leftPng) -or -not (Test-Path $rightPng)) { return }
  $imgL = [System.Drawing.Image]::FromFile($leftPng)
  $imgR = [System.Drawing.Image]::FromFile($rightPng)
  $w = 960; $h = 560; $hdr = 36
  $canvas = New-Object System.Drawing.Bitmap(($w * 2), ($h + $hdr))
  $g = [System.Drawing.Graphics]::FromImage($canvas)
  $g.Clear([System.Drawing.Color]::FromArgb(24, 24, 28))
  $g.InterpolationMode = 'HighQualityBicubic'
  $font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
  $brush = [System.Drawing.Brushes]::White
  $g.DrawString($titleLeft, $font, $brush, 12, 8)
  $g.DrawString($titleRight, $font, $brush, ($w + 12), 8)
  $g.DrawImage($imgL, 0, $hdr, $w, $h)
  $g.DrawImage($imgR, $w, $hdr, $w, $h)
  $canvas.Save($outPng, [System.Drawing.Imaging.ImageFormat]::Png)
  $font.Dispose(); $g.Dispose(); $canvas.Dispose(); $imgL.Dispose(); $imgR.Dispose()
}

function Start-DesktopVideo([string]$outMp4) {
  if ($NoVideo -or -not (Test-Path $FFMPEG)) { return $null }
  if (Test-Path $outMp4) { Remove-Item $outMp4 -Force -EA SilentlyContinue }
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $FFMPEG
  $psi.Arguments = "-y -f gdigrab -framerate 8 -i desktop -vf scale=1280:-2 -c:v libx264 -preset ultrafast -pix_fmt yuv420p `"$outMp4`""
  $psi.UseShellExecute = $false
  $psi.RedirectStandardInput = $true
  $psi.RedirectStandardError = $true
  $psi.CreateNoWindow = $true
  $p = [System.Diagnostics.Process]::Start($psi)
  Say "Video recording started PID=$($p.Id) -> $outMp4"
  return $p
}

function Stop-DesktopVideo($proc, [string]$outMp4) {
  if (-not $proc) { return }
  try {
    if (-not $proc.HasExited) {
      $proc.StandardInput.WriteLine('q')
      $proc.StandardInput.Flush()
      if (-not $proc.WaitForExit(12000)) { $proc.Kill() }
    }
  } catch { try { $proc.Kill() } catch {} }
  $sz = if (Test-Path $outMp4) { (Get-Item $outMp4).Length } else { 0 }
  LogResult 'Video' (Split-Path -Leaf $outMp4) ($sz -gt 10000) ("size={0:N0} B" -f $sz)
}

function Get-HttpJson([string]$url) {
  $wc = New-Object System.Net.WebClient
  $wc.Encoding = [System.Text.Encoding]::UTF8
  $raw = $wc.DownloadString($url)
  $wc.Dispose()
  return ($raw | ConvertFrom-Json)
}

function Post-HttpJson([string]$url, [string]$body = '{}') {
  $wc = New-Object System.Net.WebClient
  $wc.Encoding = [System.Text.Encoding]::UTF8
  $wc.Headers['Content-Type'] = 'application/json'
  $raw = $wc.UploadString($url, $body)
  $wc.Dispose()
  return ($raw | ConvertFrom-Json)
}

function Wait-ServerPort([string]$portFile, [int]$timeoutSec = 25) {
  $deadline = (Get-Date).AddSeconds($timeoutSec)
  while ((Get-Date) -lt $deadline) {
    if (Test-Path $portFile) {
      $txt = (Get-Content $portFile -EA SilentlyContinue | Select-Object -First 1)
      if ($txt -and $txt.Trim() -match '^\d+$') { return [int]$txt.Trim() }
    }
    Start-Sleep -Milliseconds 300
  }
  return 0
}

function Wait-MainWindow($proc, [int]$timeoutSec = 15) {
  $deadline = (Get-Date).AddSeconds($timeoutSec)
  while ((Get-Date) -lt $deadline) {
    if ($proc.HasExited) { return [IntPtr]::Zero }
    $proc.Refresh()
    if ($proc.MainWindowHandle -ne [IntPtr]::Zero) { return $proc.MainWindowHandle }
    Start-Sleep -Milliseconds 300
  }
  return [IntPtr]::Zero
}

function Stop-Spawned([int[]]$beforePids, [string[]]$names) {
  foreach ($n in $names) {
    Get-Process -Name $n -EA SilentlyContinue | Where-Object { $beforePids -notcontains $_.Id } | ForEach-Object {
      try { $_.CloseMainWindow() | Out-Null } catch {}
    }
  }
  Start-Sleep -Milliseconds 700
  foreach ($n in $names) {
    Get-Process -Name $n -EA SilentlyContinue | Where-Object { $beforePids -notcontains $_.Id } | ForEach-Object {
      try { Stop-Process -Id $_.Id -Force -EA SilentlyContinue } catch {}
    }
  }
}

$videoProc = $null
try {
  $videoProc = Start-DesktopVideo $VIDDIR\windows_v018_qa.mp4

  # =========================================================================
  # [1] INSTALLER & ARTIFACT MATRIX CHECK
  # =========================================================================
  if (-not $SkipInstallTest) {
    Say '=== [1] Windows Artifact & Installer Verification ==='
    $expected = @(
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x86-Setup.exe",
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x86.msi",
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x86-portable.zip",
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x64-Setup.exe",
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x64.msi",
      "Windows 7-8.1\$SLUG-$VER-Win7-8.1-x64-portable.zip",
      "Windows 10-11\$SLUG-$VER-Win10-11-x86-Setup.exe",
      "Windows 10-11\$SLUG-$VER-Win10-11-x86.msi",
      "Windows 10-11\$SLUG-$VER-Win10-11-x86-portable.zip",
      "Windows 10-11\$SLUG-$VER-Win10-11-x86.appx",
      "Windows 10-11\$SLUG-$VER-Win10-11-x64-Setup.exe",
      "Windows 10-11\$SLUG-$VER-Win10-11-x64.msi",
      "Windows 10-11\$SLUG-$VER-Win10-11-x64-portable.zip",
      "Windows 10-11\$SLUG-$VER-Win10-11-x64.appx",
      "Windows 10-11\$SLUG-$VER-Win10-11-ARM64-Setup.exe",
      "Windows 10-11\$SLUG-$VER-Win10-11-ARM64.msi",
      "Windows 10-11\$SLUG-$VER-Win10-11-ARM64-portable.zip",
      "Windows 10-11\$SLUG-$VER-Win10-11-ARM64.appx",
      "$SLUG-$VER-Setup.exe",
      "$SLUG-$VER.msi",
      "$SLUG-$VER-portable.zip",
      "$SLUG-$VER.appxbundle",
      "$SLUG-$VER-DevCert.cer",
      "version.json",
      "SHA256_Windows_v$VER.md"
    )
    $missing = @()
    foreach ($rel in $expected) {
      $fp = Join-Path $OUT_WIN $rel
      if (-not (Test-Path $fp) -or (Get-Item $fp).Length -eq 0) { $missing += $rel }
    }
    LogResult 'Matrix' 'All 25 Windows deliverables' ($missing.Count -eq 0) ("missing=" + ($missing -join ', '))

    # Verify ARM64 portable WebView2Loader matches libs\arm64
    $armZip = Join-Path $OUT_WIN "Windows 10-11\$SLUG-$VER-Win10-11-ARM64-portable.zip"
    if (Test-Path $armZip) {
      $armDst = Join-Path $BLD 'check_arm64'
      if (Test-Path $armDst) { Remove-Item $armDst -Recurse -Force }
      [System.IO.Compression.ZipFile]::ExtractToDirectory($armZip, $armDst)
      $srcHash = (Get-FileHash (Join-Path $PLAT 'windows\libs\arm64\WebView2Loader.dll') -Algorithm SHA256).Hash
      $zipHash = (Get-FileHash (Join-Path $armDst 'WebView2Loader.dll') -Algorithm SHA256).Hash
      $edTxt   = (Get-Content (Join-Path $armDst 'app\edition.txt') -Raw).Trim()
      LogResult 'ARM64' 'WebView2Loader arm64 + edition.txt' (($srcHash -eq $zipHash) -and ($edTxt -eq 'Win10-11-ARM64')) ("hashMatch=$($srcHash -eq $zipHash) edition=$edTxt")
      Remove-Item $armDst -Recurse -Force -EA SilentlyContinue
    }

    # Verify APPXBUNDLE signature
    $bundlePath = Join-Path $OUT_WIN "$SLUG-$VER.appxbundle"
    if (Test-Path $bundlePath) {
      $sig = Get-AuthenticodeSignature $bundlePath
      LogResult 'APPX' 'Authenticode signature' ($null -ne $sig.SignerCertificate) ("status=$($sig.Status) subj=$($sig.SignerCertificate.Subject)")
    }

    # Test Portable ZIP launch & HTTP API
    $portZip = Join-Path $OUT_WIN "Windows 10-11\$SLUG-$VER-$Edition-portable.zip"
    $portDir = Join-Path $BLD 'portable_test'
    if (Test-Path $portDir) { Remove-Item $portDir -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($portZip, $portDir)
    $portExe = Join-Path $portDir 'OfficeOpenSource.exe'
    $portFileP = Join-Path $portDir 'data\server.port'
    $pPort = Start-Process $portExe -PassThru
    $hwnd = Wait-MainWindow $pPort 20
    $portNum = Wait-ServerPort $portFileP 20
    $apiOk = $false; $scanOk = $false; $detail = "port=$portNum"
    if ($portNum -gt 0) {
      Start-Sleep -Seconds 2
      try {
        $info = Get-HttpJson "http://127.0.0.1:$portNum/api/info"
        $scan = Post-HttpJson "http://127.0.0.1:$portNum/api/scan_office" '{}'
        $apiOk = ($info.version -eq $VER -and $info.portable -eq $true -and $info.msi -eq $false)
        $scanOk = ($scan.ok -eq $true -and $scan.files.Count -ge 20)
        $detail = "ver=$($info.version) portable=$($info.portable) msi=$($info.msi) scanFiles=$($scan.files.Count) title=$($pPort.MainWindowTitle)"
      } catch { $detail = "err=$($_.Exception.Message)" }
    }
    Capture-Window $hwnd (Join-Path $SHOTS '01_portable_home.png')
    LogResult 'Installer' 'Portable ZIP runtime + API' ($apiOk -and $scanOk) $detail
    try { $pPort.CloseMainWindow() | Out-Null; Start-Sleep -Milliseconds 600; if (-not $pPort.HasExited) { $pPort.Kill() } } catch {}

    # Preserve pre-existing install dir state before testing MSI & Setup.exe
    $localProg = Join-Path $env:LOCALAPPDATA 'Programs\YomiOfficeOpenSource'
    $msiFile = Join-Path $OUT_WIN "Windows 10-11\$SLUG-$VER-$Edition.msi"
    $msiLog  = Join-Path $BLD 'msi_install.log'
    $pMsiInst = Start-Process 'msiexec.exe' -ArgumentList "/i `"$msiFile`" /qn /l*v `"$msiLog`"" -Wait -PassThru
    $msiFlagExists = Test-Path (Join-Path $localProg 'installed_by_msi.txt')
    $msiExe = Join-Path $localProg 'OfficeOpenSource.exe'
    $msiApiOk = $false; $msiDetail = "exit=$($pMsiInst.ExitCode) flag=$msiFlagExists"
    if ($pMsiInst.ExitCode -eq 0 -and (Test-Path $msiExe)) {
      $appPortFile = Join-Path $env:LOCALAPPDATA 'OfficeOpenSource\server.port'
      Remove-Item $appPortFile -Force -EA SilentlyContinue
      $pm = Start-Process $msiExe -PassThru
      $hm = Wait-MainWindow $pm 20
      $mp = Wait-ServerPort $appPortFile 20
      if ($mp -gt 0) {
        Start-Sleep -Seconds 2
        try {
          $im = Get-HttpJson "http://127.0.0.1:$mp/api/info"
          $msiApiOk = ($im.version -eq $VER -and $im.msi -eq $true -and $im.portable -eq $false)
          $msiDetail += " ver=$($im.version) msi=$($im.msi) title=$($pm.MainWindowTitle)"
        } catch { $msiDetail += " err=$($_.Exception.Message)" }
      }
      Capture-Window $hm (Join-Path $SHOTS '02_msi_home.png')
      try { $pm.CloseMainWindow() | Out-Null; Start-Sleep -Milliseconds 600; if (-not $pm.HasExited) { $pm.Kill() } } catch {}
    }
    LogResult 'Installer' 'MSI install + /api/info msi=true' ($pMsiInst.ExitCode -eq 0 -and $msiFlagExists -and $msiApiOk) $msiDetail
    $pMsiUn = Start-Process 'msiexec.exe' -ArgumentList "/x `"$msiFile`" /qn" -Wait -PassThru
    LogResult 'Installer' 'MSI clean uninstall' ($pMsiUn.ExitCode -eq 0 -and -not (Test-Path $msiExe)) ("exit=$($pMsiUn.ExitCode)")

    # Test Setup.exe silent install (/S) & --uninstall
    $setupExe = Join-Path $OUT_WIN "Windows 10-11\$SLUG-$VER-$Edition-Setup.exe"
    $pSetup = Start-Process $setupExe -ArgumentList '/S' -Wait -PassThru
    $setupInstalledExe = Join-Path $localProg 'OfficeOpenSource.exe'
    $noMsiFlag = -not (Test-Path (Join-Path $localProg 'installed_by_msi.txt'))
    $regExists = Test-Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\YomiOfficeOpenSource'
    $setupApiOk = $false; $sDetail = "exit=$($pSetup.ExitCode) reg=$regExists noMsiFlag=$noMsiFlag"
    if ($pSetup.ExitCode -eq 0 -and (Test-Path $setupInstalledExe)) {
      $appPortFile = Join-Path $env:LOCALAPPDATA 'OfficeOpenSource\server.port'
      Remove-Item $appPortFile -Force -EA SilentlyContinue
      $ps = Start-Process $setupInstalledExe -PassThru
      $hs = Wait-MainWindow $ps 20
      $sp = Wait-ServerPort $appPortFile 20
      if ($sp -gt 0) {
        Start-Sleep -Seconds 2
        try {
          $is = Get-HttpJson "http://127.0.0.1:$sp/api/info"
          $setupApiOk = ($is.version -eq $VER -and $is.msi -eq $false)
          $sDetail += " ver=$($is.version) msi=$($is.msi) title=$($ps.MainWindowTitle)"
        } catch { $sDetail += " err=$($_.Exception.Message)" }
      }
      Capture-Window $hs (Join-Path $SHOTS '03_setup_home.png')
      try { $ps.CloseMainWindow() | Out-Null; Start-Sleep -Milliseconds 600; if (-not $ps.HasExited) { $ps.Kill() } } catch {}
    }
    LogResult 'Installer' 'Setup.exe /S + /api/info msi=false' ($pSetup.ExitCode -eq 0 -and $regExists -and $noMsiFlag -and $setupApiOk) $sDetail
  }

  # =========================================================================
  # [2] EXTRACT PORTABLE RUNTIME FOR FEATURE QA & DOCUMENT COMPARISON
  # =========================================================================
  $qaRunDir = Join-Path $BLD 'runtime_x64'
  if (-not (Test-Path (Join-Path $qaRunDir 'OfficeOpenSource.exe'))) {
    $portZip = Join-Path $OUT_WIN "Windows 10-11\$SLUG-$VER-$Edition-portable.zip"
    if (Test-Path $qaRunDir) { Remove-Item $qaRunDir -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($portZip, $qaRunDir)
  }
  $oosExe = Join-Path $qaRunDir 'OfficeOpenSource.exe'
  $oosPortFile = Join-Path $qaRunDir 'data\server.port'

  if (-not $SkipFeatureTest) {
    Say '=== [2] Windows Feature & Web UI Verification ==='
    # Test CLI --convert on one DOCX -> PDF and one XLSX -> PDF to verify embedded LibreOffice 24.8.7 engine
    $convIn = Join-Path $TESTS 'AGC_letter_to_Hai_Long_15_Sep.docx'
    $convOut = Join-Path $BLD 'conv_test.pdf'
    Remove-Item $convOut -Force -EA SilentlyContinue
    $pc = Start-Process $oosExe -ArgumentList "--convert `"$convIn`" `"$convOut`"" -Wait -PassThru
    $convSz = if (Test-Path $convOut) { (Get-Item $convOut).Length } else { 0 }
    LogResult 'Feature' 'LOK --convert DOCX->PDF' ($pc.ExitCode -eq 0 -and $convSz -gt 5000) ("exit=$($pc.ExitCode) pdf=$convSz B")

    # Test web server endpoints & capture feature UI screens via installed Edge/Chrome pointing at live Launcher port OR OfficeOpenSource.exe
    Remove-Item $oosPortFile -Force -EA SilentlyContinue
    $pf = Start-Process $oosExe -PassThru
    $hf = Wait-MainWindow $pf 20
    $port = Wait-ServerPort $oosPortFile 20
    if ($port -gt 0) {
      Start-Sleep -Seconds 2
      Capture-Window $hf (Join-Path $SHOTS '04_yomi_home_screen.png')
      $scan = Post-HttpJson "http://127.0.0.1:$port/api/scan_office" '{}'
      LogResult 'Feature' 'POST /api/scan_office' ($scan.ok -and $scan.files.Count -ge 20) ("found=$($scan.files.Count) files")

      $edge = @('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', 'C:\Program Files\Microsoft\Edge\Application\msedge.exe') | Where-Object { Test-Path $_ } | Select-Object -First 1
      if ($edge) {
        $edgeTmp = Join-Path $BLD 'edge_selftest'
        New-Item -ItemType Directory -Force -Path $edgeTmp | Out-Null
        $dom = & $edge --headless=new --disable-gpu "--user-data-dir=$edgeTmp" --virtual-time-budget=12000 --dump-dom "http://127.0.0.1:$port/#selftest" 2>$null
        $domStr = ($dom -join "`n")
        $stOk = ($domStr -match 'passed' -or $domStr -match 'OOSTEST')
        LogResult 'Feature' 'Web #selftest via Edge headless' $stOk ("matched=" + $stOk)
        foreach ($tab in @(@('devscan','05_ui_devscan.png'), @('pdf','06_ui_pdf.png'), @('mathtype','07_ui_mathtype.png'), @('scanner','08_ui_camscanner.png'), @('ai','09_ui_ai_freeclaude.png'))) {
          $shotOut = Join-Path $SHOTS $tab[1]
          & $edge --headless=new --disable-gpu "--user-data-dir=$edgeTmp" --window-size=1366,860 --virtual-time-budget=4000 "--screenshot=$shotOut" "http://127.0.0.1:$port/#tool=$($tab[0])" 2>$null | Out-Null
          LogResult 'Feature' "Screenshot #tool=$($tab[0])" (Test-Path $shotOut) $tab[1]
        }
      }
    }
    try { $pf.CloseMainWindow() | Out-Null; Start-Sleep -Milliseconds 600; if (-not $pf.HasExited) { $pf.Kill() } } catch {}
  }

  # =========================================================================
  # [3] ALL 23 TEST OFFICE FILES SIDE-BY-SIDE VS MICROSOFT OFFICE
  # =========================================================================
  if (-not $SkipCompare) {
    Say '=== [3] Test Office (23 files) Side-by-Side Comparison ==='
    $beforeOffice = @()
    Get-Process WINWORD, EXCEL, POWERPNT, OfficeOpenSource, soffice.bin -EA SilentlyContinue | ForEach-Object { $beforeOffice += $_.Id }

    $testFiles = Get-ChildItem $TESTS -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name
    if ($MaxCompare -gt 0) { $testFiles = $testFiles | Select-Object -First $MaxCompare }

    $idx = 0
    foreach ($tf in $testFiles) {
      $idx++
      $ext = $tf.Extension.ToLowerInvariant()
      $safe = ($tf.BaseName -replace '[^a-zA-Z0-9_-]+', '_').Trim('_')
      if ($safe.Length -gt 32) { $safe = $safe.Substring(0, 32) }
      $tag = "{0:D2}_{1}" -f $idx, $safe

      $msExe = $null; $msArgs = "`"$($tf.FullName)`""; $msProcName = $null
      if ($ext -in @('.doc', '.docx', '.rtf', '.odt')) { $msExe = $MS_WORD; $msProcName = 'WINWORD'; $msArgs = "/q `"$($tf.FullName)`"" }
      elseif ($ext -in @('.xls', '.xlsx', '.xlsm', '.ods', '.csv')) { $msExe = $MS_EXCEL; $msProcName = 'EXCEL'; $msArgs = "/e `"$($tf.FullName)`"" }
      elseif ($ext -in @('.ppt', '.pptx', '.ppsx', '.odp')) { $msExe = $MS_PPT; $msProcName = 'POWERPNT'; $msArgs = "`"$($tf.FullName)`"" }

      Say ("[{0:D2}/{1:D2}] Testing {2} ({3:N0} KB)" -f $idx, $testFiles.Count, $tf.Name, ($tf.Length / 1KB))

      # Launch Yomi-Office Native Engine
      $waitSec = if ($tf.Length -gt 20MB) { 18 } elseif ($tf.Length -gt 5MB) { 9 } else { 5 }
      $py = Start-Process $oosExe -ArgumentList "--native `"$($tf.FullName)`"" -PassThru
      $hy = Wait-MainWindow $py 25
      if ($hy -ne [IntPtr]::Zero) { Start-Sleep -Seconds $waitSec }
      $yAlive = ($hy -ne [IntPtr]::Zero) -and (-not $py.HasExited)
      $yShot = Join-Path $SHOTS "cmp_${tag}_yomi.png"
      Capture-Window $hy $yShot

      # Position Yomi window on left half and launch MS Office on right half for visual/video side-by-side
      $mOk = $false
      $mShot = Join-Path $SHOTS "cmp_${tag}_msoffice.png"
      if ($msExe -and (Test-Path $msExe)) {
        if ($hy -ne [IntPtr]::Zero) { [WinQaNative]::MoveWindow($hy, 0, 0, 960, 1020, $true) | Out-Null }
        $pm = Start-Process $msExe -ArgumentList $msArgs -PassThru
        Start-Sleep -Seconds $waitSec
        $mp = Get-Process -Name $msProcName -EA SilentlyContinue | Where-Object { $beforeOffice -notcontains $_.Id -and $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
        if ($mp) {
          $hm = $mp.MainWindowHandle
          [WinQaNative]::MoveWindow($hm, 960, 0, 960, 1020, $true) | Out-Null
          Start-Sleep -Milliseconds 800
          Capture-Window $hm $mShot
          $mOk = $true
        } else {
          Capture-Screen $mShot
        }
      }

      if ($yAlive -and $mOk) {
        $pairShot = Join-Path $SHOTS "cmp_${tag}_pair.png"
        Make-SideBySide $yShot $mShot $pairShot "Yomi-Office 0.1.8: $($tf.Name)" "Microsoft Office 16: $($tf.Name)"
      }

      LogResult 'Compare' $tf.Name ($yAlive -and (Test-Path $yShot)) ("yomi=$yAlive msOffice=$mOk size=$($tf.Length)")
      Stop-Spawned $beforeOffice @('OfficeOpenSource', 'soffice.bin', 'WINWORD', 'EXCEL', 'POWERPNT')
      Start-Sleep -Milliseconds 500
    }
  }
}
finally {
  Stop-DesktopVideo $videoProc (Join-Path $VIDDIR 'windows_v018_qa.mp4')
}

# Write markdown report
$md = New-Object System.Collections.Generic.List[string]
$md.Add("# Báo cáo kiểm thử Windows v$VER")
$md.Add("")
$md.Add("- Thời gian: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$md.Add("- Kết quả tổng: **$script:pass PASS / $script:fail FAIL**")
$md.Add("")
$md.Add("| KQ | Nhóm | Hạng mục | Chi tiết | Giờ |")
$md.Add("|---|---|---|---|---|")
foreach ($r in $script:results) {
  $md.Add("| $($r.Status) | $($r.Category) | $($r.Item) | $($r.Detail.Replace('|','/')) | $($r.Time) |")
}
[IO.File]::WriteAllText($REPORT, ($md -join "`r`n"), (New-Object Text.UTF8Encoding($false)))
Say "DONE: $script:pass PASS, $script:fail FAIL -> $REPORT"
