# Yomi-Office Open Source v0.1.8 - Windows full build matrix
#   Engine: LibreOffice 24.8.7 LTS core on every Windows edition + embedded WebView2 + unified native UI
#   Editions:
#     1. Win7-8.1-x86    (core win7-x86, bundle universal) -> Setup.exe, .msi, portable.zip
#     2. Win7-8.1-x64    (core win7-x64, bundle universal) -> Setup.exe, .msi, portable.zip
#     3. Win10-11-x86    (core win7-x86, bundle modern)    -> Setup.exe, .msi, portable.zip, .appx
#     4. Win10-11-x64    (core win7-x64, bundle modern)    -> Setup.exe, .msi, portable.zip, .appx
#     5. Win10-11-ARM64  (core w24-arm64, bundle modern)   -> Setup.exe, .msi, portable.zip, .appx
#     + Universal (All-AnyCPU) -> Setup.exe, .msi, portable.zip
#     + Signed .appxbundle + DevCert.cer + version.json + SHA256_Windows_v0.1.8.md
param(
  [switch]$SkipWeb,
  [switch]$SkipAppx,
  [switch]$SkipMsi,
  [string]$OnlyEdition
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$VER  = '0.1.8'
$VERQ = '0.1.8.0'
$APP_SLUG = 'Yomi-Office-Open-Source'
$APP_NAME = 'Yomi-Office Open Source'

# Dynamic paths (no hardcoded absolute project root)
$PLAT = $PSScriptRoot
$SRC  = Split-Path -Parent $PLAT
$VDIR = Split-Path -Parent $SRC
$PROJ = Split-Path -Parent $VDIR

$DISTS = @{
  universal = Join-Path $SRC 'dist'
  modern    = Join-Path $SRC 'dist-modern'
  legacy    = Join-Path $SRC 'dist-legacy'
}

$OUT_ROOT = Join-Path (Split-Path -Parent $PROJ) 'output'
$REL      = Join-Path $OUT_ROOT "v$VER"
$REL_WIN  = Join-Path $REL 'Windows'
$BLD      = Join-Path $VDIR 'build'
$LO       = Join-Path $PROJ 'v016\lo'

$KIT = 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64'
$CSC_ROSLYN = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe'
$CSC_FX = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$CSC = if (Test-Path $CSC_ROSLYN) { $CSC_ROSLYN } else { $CSC_FX }
$WIX = 'D:\yh\tools\wix314'

function Say($m) { Write-Output ("[{0}] {1}" -f (Get-Date -Format HH:mm:ss), $m) }

New-Item -ItemType Directory -Force -Path $OUT_ROOT, $REL, $REL_WIN, $BLD | Out-Null

# [0] Web App bundles
if (-not $SkipWeb) {
  Say '== [0] Building Web App bundles (universal, modern, legacy)'
  $env:NODE_PATH = Join-Path $PROJ 'v005\deps\node_modules'
  Push-Location $SRC
  try {
    Say '   -> building profile: universal'
    & node tools/build.js --profile=universal --out=dist
    if ($LASTEXITCODE -ne 0) { throw "Web build failed (universal)" }
    Say '   -> building profile: modern'
    & node tools/build.js --profile=modern --out=dist-modern
    if ($LASTEXITCODE -ne 0) { throw "Web build failed (modern)" }
    Say '   -> building profile: legacy'
    & node tools/build.js --profile=legacy --out=dist-legacy
    if ($LASTEXITCODE -ne 0) { throw "Web build failed (legacy)" }
  } finally {
    Pop-Location
  }
}

foreach ($k in @('universal', 'modern')) {
  if (-not (Test-Path (Join-Path $DISTS[$k] 'index.html'))) {
    throw "FATAL: dist $k missing ($($DISTS[$k]))"
  }
}

$ico = Join-Path $VDIR 'logo\app.ico'
if (-not (Test-Path $ico)) { throw "FATAL: app.ico missing at $ico" }
$logo = Join-Path $VDIR 'logo\Yomi-Office Open Source.png'
if (-not (Test-Path $logo)) { $logo = Join-Path $VDIR 'logo\office-logo-icon.png' }
if (-not (Test-Path $logo)) { throw "FATAL: logo PNG missing in $VDIR\logo" }

function Build-Launcher($cpu, $loaderArch, $dest) {
  if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  $winDir = Join-Path $PLAT 'windows'
  $libDir = Join-Path $winDir 'libs'

  & $CSC /nologo /target:winexe "/platform:$cpu" /optimize+ /unsafe /codepage:65001 /utf8output "/win32icon:$ico" `
     /r:System.dll /r:System.Core.dll /r:System.Web.Extensions.dll /r:System.Security.dll /r:Microsoft.CSharp.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll `
     "/r:$libDir\Microsoft.Web.WebView2.WinForms.dll" "/r:$libDir\Microsoft.Web.WebView2.Core.dll" `
     "/out:$dest\OfficeOpenSource.exe" `
     "$winDir\Launcher.cs" "$winDir\LokView.cs" "$winDir\LanShare.cs" "$winDir\OfficeUi.cs" "$winDir\Ribbon.cs" `
     "$winDir\Chrome.cs" "$winDir\Backstage.cs" "$winDir\NativeDocForm.cs" "$winDir\RibbonDefs.cs" `
     "$winDir\Dialogs.cs" "$winDir\AiPane.cs" "$winDir\UnifiedUi.cs" "$winDir\UnifiedSettings.cs" `
     "/resource:$SRC\src\core\ui_layout.json,ui_layout.json" `
     "/resource:$SRC\src\core\ui_settings.json,ui_settings.json" `
     /nowarn:1690,0162,0108
  if ($LASTEXITCODE -ne 0) { throw "csc Launcher failed for platform=$cpu (exit=$LASTEXITCODE)" }

  Copy-Item "$winDir\app.config" "$dest\OfficeOpenSource.exe.config" -Force
  Copy-Item "$libDir\Microsoft.Web.WebView2.WinForms.dll" "$dest\" -Force
  Copy-Item "$libDir\Microsoft.Web.WebView2.Core.dll" "$dest\" -Force
  Copy-Item $ico "$dest\app.ico" -Force

  if ($loaderArch -eq 'all') {
    Copy-Item "$libDir\x64\WebView2Loader.dll" "$dest\" -Force
    foreach ($a in @('x86', 'x64', 'arm64')) {
      New-Item -ItemType Directory -Force -Path "$dest\runtimes\win-$a\native" | Out-Null
      Copy-Item "$libDir\$a\WebView2Loader.dll" "$dest\runtimes\win-$a\native\" -Force
    }
  } else {
    Copy-Item "$libDir\$loaderArch\WebView2Loader.dll" "$dest\" -Force
    New-Item -ItemType Directory -Force -Path "$dest\runtimes\win-$loaderArch\native" | Out-Null
    Copy-Item "$libDir\$loaderArch\WebView2Loader.dll" "$dest\runtimes\win-$loaderArch\native\" -Force
  }
  return $true
}

function Add-BasicTools($pb) {
  if (-not (Test-Path $pb)) { return }
  & robocopy "$PLAT\windows\basic\OOS" "$pb\OOS" /E /NFL /NDL /NJH /NJS /NP | Out-Null
  if ($LASTEXITCODE -ge 8) { throw "robocopy BasicTools failed ($LASTEXITCODE)" }
  foreach ($x in @('script.xlc', 'dialog.xlc')) {
    $f = "$pb\$x"
    if (Test-Path $f) {
      $t = [IO.File]::ReadAllText($f)
      if ($t -notmatch 'library:name="OOS"') {
        $t = $t.Replace('</library:libraries>', ' <library:library library:name="OOS" library:link="false"/>' + "`n" + '</library:libraries>')
        [IO.File]::WriteAllText($f, $t, (New-Object Text.UTF8Encoding $false))
      }
    }
  }
}

$UpgradeCodes = @{
  'Win7-8.1-x86'   = 'A8E4B7C1-5F20-4D31-9B10-018000000101'
  'Win7-8.1-x64'   = 'A8E4B7C1-5F20-4D31-9B10-018000000102'
  'Win10-11-x86'   = 'A8E4B7C1-5F20-4D31-9B10-018000000201'
  'Win10-11-x64'   = 'A8E4B7C1-5F20-4D31-9B10-018000000202'
  'Win10-11-ARM64' = 'A8E4B7C1-5F20-4D31-9B10-018000000203'
  'All-AnyCPU'     = 'A8E4B7C1-5F20-4D31-9B10-018000000900'
}

function Patch-HeatPerUser([string]$wxsPath) {
  [xml]$xml = Get-Content -Path $wxsPath -Encoding UTF8
  $nsUri = 'http://schemas.microsoft.com/wix/2006/wi'
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace('w', $nsUri)

  $md5 = [System.Security.Cryptography.MD5]::Create()
  foreach ($comp in $xml.SelectNodes('//w:Component', $ns)) {
    $cid = $comp.GetAttribute('Id')
    if ($comp.GetAttribute('Guid') -eq '*') {
      $hash = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("YomiOOS:" + $cid))
      $comp.SetAttribute('Guid', (New-Object guid (,$hash)).ToString('B').ToUpper())
    }
    foreach ($f in $comp.SelectNodes('w:File', $ns)) {
      $f.RemoveAttribute('KeyPath')
    }
    $reg = $xml.CreateElement('RegistryValue', $nsUri)
    $reg.SetAttribute('Root', 'HKCU')
    $reg.SetAttribute('Key', 'Software\YomiOfficeOpenSource\MSI\Files')
    $reg.SetAttribute('Name', $cid)
    $reg.SetAttribute('Type', 'integer')
    $reg.SetAttribute('Value', '1')
    $reg.SetAttribute('KeyPath', 'yes')
    [void]$comp.AppendChild($reg)
  }
  $md5.Dispose()

  # Ensure every harvested directory has a RemoveFolder on uninstall (ICE64)
  $frag = $xml.SelectSingleNode('//w:Fragment', $ns)
  $dirRef = $xml.CreateElement('DirectoryRef', $nsUri)
  $dirRef.SetAttribute('Id', 'INSTALLFOLDER')
  $rmComp = $xml.CreateElement('Component', $nsUri)
  $rmComp.SetAttribute('Id', 'cmpRemoveHarvestedDirs')
  $rmComp.SetAttribute('Guid', [guid]::NewGuid().ToString('B').ToUpper())

  $rmInst = $xml.CreateElement('RemoveFolder', $nsUri)
  $rmInst.SetAttribute('Id', 'rm_INSTALLFOLDER')
  $rmInst.SetAttribute('Directory', 'INSTALLFOLDER')
  $rmInst.SetAttribute('On', 'uninstall')
  [void]$rmComp.AppendChild($rmInst)

  $dirCount = 0
  foreach ($d in $xml.SelectNodes('//w:Directory', $ns)) {
    $did = $d.GetAttribute('Id')
    if ([string]::IsNullOrEmpty($did)) { continue }
    $dirCount++
    $rf = $xml.CreateElement('RemoveFolder', $nsUri)
    $rf.SetAttribute('Id', "rm_$did")
    $rf.SetAttribute('Directory', $did)
    $rf.SetAttribute('On', 'uninstall')
    [void]$rmComp.AppendChild($rf)
  }

  $rmKey = $xml.CreateElement('RegistryValue', $nsUri)
  $rmKey.SetAttribute('Root', 'HKCU')
  $rmKey.SetAttribute('Key', 'Software\YomiOfficeOpenSource\MSI')
  $rmKey.SetAttribute('Name', 'DirCleanup')
  $rmKey.SetAttribute('Type', 'integer')
  $rmKey.SetAttribute('Value', '1')
  $rmKey.SetAttribute('KeyPath', 'yes')
  [void]$rmComp.AppendChild($rmKey)

  [void]$dirRef.AppendChild($rmComp)
  [void]$frag.AppendChild($dirRef)

  $cg = $xml.SelectSingleNode('//w:ComponentGroup[@Id="AppFiles"]', $ns)
  if ($cg) {
    $cref = $xml.CreateElement('ComponentRef', $nsUri)
    $cref.SetAttribute('Id', 'cmpRemoveHarvestedDirs')
    [void]$cg.AppendChild($cref)
  }

  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $settings.Indent = $true
  $xw = [System.Xml.XmlWriter]::Create($wxsPath, $settings)
  $xml.Save($xw)
  $xw.Close()
}

function Build-Msi($tag, $wixArch, $instDir, $msiOut) {
  if ($SkipMsi) { return }
  $heat   = Join-Path $WIX 'heat.exe'
  $candle = Join-Path $WIX 'candle.exe'
  $light  = Join-Path $WIX 'light.exe'
  foreach ($tool in @($heat, $candle, $light)) {
    if (-not (Test-Path $tool)) { throw "WiX tool not found: $tool" }
  }

  $work = Join-Path $BLD "msi_work_$tag"
  if (Test-Path $work) { Remove-Item $work -Recurse -Force }
  $stage = Join-Path $work 'stage'
  $obj   = Join-Path $work 'obj'
  New-Item -ItemType Directory -Force -Path $stage, $obj | Out-Null

  & robocopy $instDir $stage /E /NFL /NDL /NJH /NJS /NP /MT:16 | Out-Null
  if ($LASTEXITCODE -ge 8) { throw "robocopy stage for MSI failed ($LASTEXITCODE)" }
  [IO.File]::WriteAllText((Join-Path $stage 'installed_by_msi.txt'), "$APP_NAME $VER ($tag)`r`n", (New-Object Text.UTF8Encoding($false)))

  $filesWxs = Join-Path $work 'files.wxs'
  & $heat dir $stage -nologo -ag -srd -sreg -sfrag -cg AppFiles -dr INSTALLFOLDER -var var.StageDir -out $filesWxs
  if ($LASTEXITCODE -ne 0) { throw "heat.exe failed for $tag (exit=$LASTEXITCODE)" }

  Patch-HeatPerUser $filesWxs

  $upCode = $UpgradeCodes[$tag]
  if (-not $upCode) { throw "Missing UpgradeCode for $tag" }
  $smGuid = [guid]::NewGuid().ToString('D').ToUpper()
  $dtGuid = [guid]::NewGuid().ToString('D').ToUpper()

  $productWxs = Join-Path $work 'product.wxs'
  $wxs = @"
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="$APP_NAME" Language="1033" Codepage="65001" Version="$VERQ"
           Manufacturer="Yomi-Office Open Source Project" UpgradeCode="$upCode">
    <Package InstallerVersion="405" Compressed="yes" InstallScope="perUser" InstallPrivileges="limited"
             Description="$APP_NAME $VER ($tag)" Comments="$APP_NAME - Word, Excel, PowerPoint, PDF, A.I" />
    <MajorUpgrade DowngradeErrorMessage="Da co phien ban moi hon cua $APP_NAME duoc cai dat." AllowSameVersionUpgrades="yes" />
    <MediaTemplate EmbedCab="yes" CompressionLevel="medium" />
    <Icon Id="AppIcon.ico" SourceFile="$ico" />
    <Property Id="ARPPRODUCTICON" Value="AppIcon.ico" />
    <Property Id="ARPHELPLINK" Value="https://github.com/YomiHub-Project/Yomi-Office-Open-Source" />
    <Property Id="ARPNOMODIFY" Value="1" />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="LocalAppDataFolder">
        <Directory Id="LocalProgramsFolder" Name="Programs">
          <Directory Id="INSTALLFOLDER" Name="YomiOfficeOpenSource" />
        </Directory>
      </Directory>
      <Directory Id="ProgramMenuFolder" />
      <Directory Id="DesktopFolder" />
    </Directory>

    <DirectoryRef Id="LocalProgramsFolder">
      <Component Id="LocalProgramsDirCleanup" Guid="D8E4B7C1-5F20-4D31-9B10-018000000099">
        <RemoveFolder Id="rm_LocalProgramsFolder" Directory="LocalProgramsFolder" On="uninstall" />
        <RegistryValue Root="HKCU" Key="Software\YomiOfficeOpenSource\MSI" Name="LocalProgramsCleanup" Type="integer" Value="1" KeyPath="yes" />
      </Component>
    </DirectoryRef>

    <DirectoryRef Id="ProgramMenuFolder">
      <Component Id="StartMenuShortcut" Guid="$smGuid">
        <Shortcut Id="AppStartMenuShortcut"
                  Name="$APP_NAME"
                  Description="$APP_NAME - Word, Excel, PowerPoint, PDF"
                  Target="[INSTALLFOLDER]OfficeOpenSource.exe"
                  WorkingDirectory="INSTALLFOLDER"
                  Icon="AppIcon.ico" />
        <RegistryValue Root="HKCU" Key="Software\YomiOfficeOpenSource\MSI" Name="StartMenuShortcut" Type="integer" Value="1" KeyPath="yes" />
      </Component>
    </DirectoryRef>

    <DirectoryRef Id="DesktopFolder">
      <Component Id="DesktopShortcut" Guid="$dtGuid">
        <Shortcut Id="AppDesktopShortcut"
                  Name="$APP_NAME"
                  Description="$APP_NAME - Word, Excel, PowerPoint, PDF"
                  Target="[INSTALLFOLDER]OfficeOpenSource.exe"
                  WorkingDirectory="INSTALLFOLDER"
                  Icon="AppIcon.ico" />
        <RegistryValue Root="HKCU" Key="Software\YomiOfficeOpenSource\MSI" Name="DesktopShortcut" Type="integer" Value="1" KeyPath="yes" />
      </Component>
    </DirectoryRef>

    <Feature Id="ProductFeature" Title="$APP_NAME" Level="1">
      <ComponentGroupRef Id="AppFiles" />
      <ComponentRef Id="LocalProgramsDirCleanup" />
      <ComponentRef Id="StartMenuShortcut" />
      <ComponentRef Id="DesktopShortcut" />
    </Feature>

    <CustomAction Id="RegisterFileAssocs"
                  Directory="INSTALLFOLDER"
                  ExeCommand="&quot;[INSTALLFOLDER]OfficeOpenSource.exe&quot; --register"
                  Execute="deferred"
                  Impersonate="yes"
                  Return="ignore" />

    <InstallExecuteSequence>
      <Custom Action="RegisterFileAssocs" After="InstallFiles">NOT Installed</Custom>
    </InstallExecuteSequence>
  </Product>
</Wix>
"@
  [IO.File]::WriteAllText($productWxs, $wxs, (New-Object Text.UTF8Encoding($false)))

  & $candle -nologo -arch $wixArch "-dStageDir=$stage" -out "$obj\product.wixobj" $productWxs
  if ($LASTEXITCODE -ne 0) { throw "candle.exe failed for $tag product.wxs (exit=$LASTEXITCODE)" }
  & $candle -nologo -arch $wixArch "-dStageDir=$stage" -out "$obj\files.wixobj" $filesWxs
  if ($LASTEXITCODE -ne 0) { throw "candle.exe failed for $tag files.wxs (exit=$LASTEXITCODE)" }

  if (Test-Path $msiOut) { Remove-Item $msiOut -Force }
  # Suppress ICE60 (LO DLL language column warnings), ICE61 (AllowSameVersionUpgrades), ICE91 (per-user ProgramMenu/Desktop/LocalAppData warnings)
  & $light -nologo -sice:ICE60 -sice:ICE61 -sice:ICE91 -out $msiOut "$obj\product.wixobj" "$obj\files.wixobj"
  if ($LASTEXITCODE -ne 0) { throw "light.exe failed for $tag (exit=$LASTEXITCODE)" }
}

function Build-Windows($edition, $arch, $cpu, $wixArch, $dist, $folder, $loDir) {
  $tag = "$edition-$arch"
  $lb = Join-Path $BLD "launcher_$tag"
  $loaderArch = if ($arch -eq 'x86') { 'x86' } elseif ($arch -eq 'ARM64') { 'arm64' } elseif ($arch -eq 'AnyCPU') { 'all' } else { 'x64' }
  $ok = Build-Launcher $cpu $loaderArch $lb

  $inst = Join-Path $BLD "layout_$tag"
  if (Test-Path $inst) { Remove-Item $inst -Recurse -Force }
  New-Item -ItemType Directory -Force -Path $inst | Out-Null
  Copy-Item "$lb\*" $inst -Recurse
  Copy-Item $dist "$inst\app" -Recurse

  if ($loDir -and ((Test-Path "$loDir\program\sofficeapp.dll") -or (Test-Path "$loDir\program\mergedlo.dll"))) {
    & robocopy $loDir "$inst\lo" /E /XF updater.exe update_service.exe updater.ini updatecheckuilo.dll /NFL /NDL /NJH /NJS /NP /MT:16 | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy LO core failed ($LASTEXITCODE)" }
    Write-Host "   + LibreOffice core from $loDir"
    Add-BasicTools "$inst\lo\presets\basic"
  } else {
    Write-Host "   ! no LibreOffice core for $tag ($loDir)"
  }

  Set-Content "$inst\app\edition.txt" "$tag" -Encoding ASCII

  # Portable ZIP
  $port = Join-Path $BLD "portable_$tag"
  if (Test-Path $port) { Remove-Item $port -Recurse -Force }
  Copy-Item $inst $port -Recurse
  Set-Content "$port\portable.flag" 'Portable mode: settings and data are stored in .\data' -Encoding UTF8
  $os = if ($edition -eq 'Win7-8.1') { 'Windows 7 SP1 / 8 / 8.1' } elseif ($edition -eq 'All') { 'Windows 7 SP1 / 8 / 8.1 / 10 / 11' } else { 'Windows 10 / 11' }
  $readme = "$APP_NAME $VER - ban Native Hybrid $tag (khong can cai dat)`r`n`r`nHe dieu hanh: $os ($arch)`r`n1. Giai nen toan bo thu muc vao USB hoac o dia bat ky.`r`n2. Chay OfficeOpenSource.exe (Chay truc tiep tren Native Runtime cua ung dung).`r`n3. Du lieu va cai dat duoc luu trong thu muc data ben canh.`r`n`r`nHoat dong hoan toan ngoai tuyen."
  [IO.File]::WriteAllText("$port\README.txt", $readme, (New-Object Text.UTF8Encoding($true)))
  $zip = Join-Path $folder "$APP_SLUG-$VER-$tag-portable.zip"
  if (Test-Path $zip) { Remove-Item $zip -Force }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::CreateFromDirectory($port, $zip, [System.IO.Compression.CompressionLevel]::Optimal, $false)

  # Setup.exe
  $payloadGz = Join-Path $BLD "payload_$tag.gz"
  & "$BLD\Pack.exe" $inst $payloadGz | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "Pack.exe failed for $tag" }
  $setup = Join-Path $folder "$APP_SLUG-$VER-$tag-Setup.exe"
  & $CSC /nologo /target:winexe "/platform:$cpu" /optimize+ /codepage:65001 /utf8output "/win32icon:$ico" `
     /r:System.Windows.Forms.dll /r:System.Drawing.dll "/resource:$payloadGz,Payload" "/out:$setup" "$PLAT\windows\Setup.cs"
  if ($LASTEXITCODE -ne 0) { throw "Setup.cs compile failed for $tag" }

  # MSI
  $msi = Join-Path $folder "$APP_SLUG-$VER-$tag.msi"
  Build-Msi $tag $wixArch $inst $msi

  $msiLen = if (Test-Path $msi) { (Get-Item $msi).Length } else { 0 }
  Say ("{0,-18} launcher={1} setup={2:N0} msi={3:N0} portable={4:N0}" -f $tag, $ok, (Get-Item $setup).Length, $msiLen, (Get-Item $zip).Length)
  return $inst
}

# == [1] Pack tool
Say '== [1] Pack tool'
& $CSC_FX /nologo /target:exe /platform:anycpu /optimize+ /codepage:65001 "/out:$BLD\Pack.exe" "$PLAT\windows\Pack.cs"
if ($LASTEXITCODE -ne 0) { throw "Pack.cs compile failed" }

$lay = @{}
$w7  = Join-Path $REL_WIN 'Windows 7-8.1'
$w10 = Join-Path $REL_WIN 'Windows 10-11'
New-Item -ItemType Directory -Force -Path $w7, $w10 | Out-Null

if (-not $OnlyEdition -or $OnlyEdition -eq 'Win7-8.1-x86') {
  Say '== [2a] Win7-8.1-x86'
  $lay['Win7-8.1-x86'] = Build-Windows 'Win7-8.1' 'x86' 'x86' 'x86' $DISTS.universal $w7 "$LO\trim\win7-x86"
}
if (-not $OnlyEdition -or $OnlyEdition -eq 'Win7-8.1-x64') {
  Say '== [2b] Win7-8.1-x64'
  $lay['Win7-8.1-x64'] = Build-Windows 'Win7-8.1' 'x64' 'x64' 'x64' $DISTS.universal $w7 "$LO\trim\win7-x64"
}
if (-not $OnlyEdition -or $OnlyEdition -eq 'Win10-11-x86') {
  Say '== [3a] Win10-11-x86'
  $lay['Win10-11-x86'] = Build-Windows 'Win10-11' 'x86' 'x86' 'x86' $DISTS.modern $w10 "$LO\trim\win7-x86"
}
if (-not $OnlyEdition -or $OnlyEdition -eq 'Win10-11-x64') {
  Say '== [3b] Win10-11-x64'
  $lay['Win10-11-x64'] = Build-Windows 'Win10-11' 'x64' 'x64' 'x64' $DISTS.modern $w10 "$LO\trim\win7-x64"
}
if (-not $OnlyEdition -or $OnlyEdition -eq 'Win10-11-ARM64') {
  Say '== [3c] Win10-11-ARM64'
  # ARM64: .NET Framework 4.8.1 runs AnyCPU natively on Windows 11 ARM64; WiX 3.14 uses -arch x64
  $lay['Win10-11-ARM64'] = Build-Windows 'Win10-11' 'ARM64' 'anycpu' 'x64' $DISTS.modern $w10 "$LO\trim\w24-arm64"
}

if (-not $OnlyEdition) {
  Say '== [4] Universal (All-AnyCPU) -> output\v0.1.8\Windows'
  $uni = Build-Windows 'All' 'AnyCPU' 'anycpu' 'x86' $DISTS.universal $BLD "$LO\trim\win7-x64"
  Copy-Item "$BLD\$APP_SLUG-$VER-All-AnyCPU-Setup.exe"     (Join-Path $REL_WIN "$APP_SLUG-$VER-Setup.exe") -Force
  Copy-Item "$BLD\$APP_SLUG-$VER-All-AnyCPU-portable.zip"  (Join-Path $REL_WIN "$APP_SLUG-$VER-portable.zip") -Force
  if (-not $SkipMsi -and (Test-Path "$BLD\$APP_SLUG-$VER-All-AnyCPU.msi")) {
    Copy-Item "$BLD\$APP_SLUG-$VER-All-AnyCPU.msi"         (Join-Path $REL_WIN "$APP_SLUG-$VER.msi") -Force
  }
}

if (-not $SkipAppx -and -not $OnlyEdition) {
  Say '== [5] APPX x86/x64/arm64 + APPXBUNDLE (Windows 10 1809+ / 11)'
  Add-Type -AssemblyName System.Drawing
  $pfx = Join-Path $PROJ 'scripts\dev_cert\OfficeOpenSourceDev.pfx'
  $pfxPass = 'OOS-dev-0.0.4'
  $bdir = Join-Path $BLD 'appx_bundle'
  if (Test-Path $bdir) { Remove-Item $bdir -Recurse -Force }
  New-Item -ItemType Directory -Force -Path $bdir | Out-Null

  function Make-Asset($w, $h, $dest) {
    $bmp = New-Object Drawing.Bitmap($w, $h)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.Clear([Drawing.Color]::Transparent)
    $g.InterpolationMode = 'HighQualityBicubic'
    $img = [Drawing.Image]::FromFile($logo)
    $s = [math]::Min($w, $h) * 0.8
    $g.DrawImage($img, [float](($w - $s) / 2), [float](($h - $s) / 2), [float]$s, [float]$s)
    $img.Dispose()
    $bmp.Save($dest, [Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
  }

  $types = '.docx .docm .dotx .xlsx .xlsm .xltx .pptx .pptm .ppsx .potx .pdf .csv .md .odt .ods .odp .rtf .epub .fb2 .tex .latex .mml .eqx .mathtype' -split ' '
  $ft = ($types | ForEach-Object { "<uap:FileType>$_</uap:FileType>" }) -join ''

  $appxSpecs = @(
    @{ appxArch = 'x86';   layKey = 'Win10-11-x86';   tag = 'Win10-11-x86' },
    @{ appxArch = 'x64';   layKey = 'Win10-11-x64';   tag = 'Win10-11-x64' },
    @{ appxArch = 'arm64'; layKey = 'Win10-11-ARM64'; tag = 'Win10-11-ARM64' }
  )

  foreach ($sp in $appxSpecs) {
    $arch = $sp.appxArch
    $l = Join-Path $BLD "appx_$arch"
    if (Test-Path $l) { Remove-Item $l -Recurse -Force }
    New-Item -ItemType Directory -Force -Path "$l\Assets" | Out-Null
    Copy-Item "$($lay[$sp.layKey])\*" $l -Recurse
    Make-Asset 50 50 "$l\Assets\StoreLogo.png"
    Make-Asset 150 150 "$l\Assets\Square150x150Logo.png"
    Make-Asset 44 44 "$l\Assets\Square44x44Logo.png"
    Make-Asset 310 150 "$l\Assets\Wide310x150Logo.png"
    Make-Asset 620 300 "$l\Assets\SplashScreen.png"

    $m = @"
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10" xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10" xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities" IgnorableNamespaces="uap rescap">
  <Identity Name="OfficeOpenSource" Publisher="CN=OfficeOpenSourceDev" Version="$VERQ" ProcessorArchitecture="$arch" />
  <Properties><DisplayName>$APP_NAME</DisplayName><PublisherDisplayName>Yomi-Office Open Source Project</PublisherDisplayName><Logo>Assets\StoreLogo.png</Logo><Description>Bộ ứng dụng văn phòng mã nguồn mở: Word, Excel, PowerPoint, PDF, MathType, A.I — hoạt động ngoại tuyến</Description></Properties>
  <Resources><Resource Language="vi-VN" /></Resources>
  <Dependencies><TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.26100.0" /></Dependencies>
  <Capabilities><rescap:Capability Name="runFullTrust" /><Capability Name="internetClient" /><DeviceCapability Name="microphone" /><DeviceCapability Name="webcam" /></Capabilities>
  <Applications><Application Id="OfficeOpenSource" Executable="OfficeOpenSource.exe" EntryPoint="Windows.FullTrustApplication">
    <uap:VisualElements DisplayName="$APP_NAME" Description="$APP_NAME $VER" BackgroundColor="transparent" Square150x150Logo="Assets\Square150x150Logo.png" Square44x44Logo="Assets\Square44x44Logo.png">
      <uap:DefaultTile Wide310x150Logo="Assets\Wide310x150Logo.png" ShortName="Yomi-Office" /><uap:SplashScreen Image="Assets\SplashScreen.png" BackgroundColor="#185ABD" />
    </uap:VisualElements>
    <Extensions><uap:Extension Category="windows.fileTypeAssociation"><uap:FileTypeAssociation Name="officedocs"><uap:DisplayName>Tài liệu Yomi-Office</uap:DisplayName><uap:SupportedFileTypes>$ft</uap:SupportedFileTypes></uap:FileTypeAssociation></uap:Extension></Extensions>
  </Application></Applications>
</Package>
"@
    [IO.File]::WriteAllText("$l\AppxManifest.xml", $m, (New-Object Text.UTF8Encoding($true)))
    $ax = Join-Path $bdir "$APP_SLUG-$VER-$($sp.tag).appx"
    & "$KIT\makeappx.exe" pack /o /h SHA256 /d $l /p $ax | Select-Object -Last 3
    if ($LASTEXITCODE -ne 0) { throw "makeappx pack failed for $arch" }
    & "$KIT\signtool.exe" sign /fd SHA256 /f $pfx /p $pfxPass $ax | Select-Object -Last 1
    if ($LASTEXITCODE -ne 0) { throw "signtool failed for $ax" }
    Copy-Item $ax $w10 -Force
    Say ("appx $arch bytes=" + (Get-Item $ax).Length)
  }

  $bundle = Join-Path $REL_WIN "$APP_SLUG-$VER.appxbundle"
  if (Test-Path $bundle) { Remove-Item $bundle -Force }
  & "$KIT\makeappx.exe" bundle /o /d $bdir /p $bundle /bv $VERQ | Select-Object -Last 2
  if ($LASTEXITCODE -ne 0) { throw "makeappx bundle failed" }
  & "$KIT\signtool.exe" sign /fd SHA256 /f $pfx /p $pfxPass $bundle | Select-Object -Last 1
  if ($LASTEXITCODE -ne 0) { throw "signtool failed for bundle" }
  Copy-Item $bundle $w10 -Force
  Say "appxbundle bytes=$((Get-Item $bundle).Length)"

  $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2($pfx, $pfxPass)
  $cerPath = Join-Path $REL_WIN "$APP_SLUG-$VER-DevCert.cer"
  [IO.File]::WriteAllBytes($cerPath, $cert.Export('Cert'))
  Copy-Item $cerPath $w10 -Force
}

# == [6] version.json & SHA256_Windows_v0.1.8.md
Say '== [6] Writing version.json & SHA256_Windows_v0.1.8.md'
$files = [ordered]@{}
$escVer = [regex]::Escape($VER)
Get-ChildItem $REL_WIN -Recurse -File | Where-Object { $_.Name -match "^$APP_SLUG-" } | Sort-Object Name | ForEach-Object {
  $n = $_.Name
  if ($n -eq "$APP_SLUG-$VER-Setup.exe")        { $files['setup'] = $n }
  elseif ($n -eq "$APP_SLUG-$VER.msi")          { $files['msi'] = $n }
  elseif ($n -eq "$APP_SLUG-$VER-portable.zip") { $files['portable'] = $n }
  elseif ($n -eq "$APP_SLUG-$VER.appxbundle")   { $files['appxbundle'] = $n }
  elseif ($n -match "^$APP_SLUG-$escVer-(.+)-Setup\.exe$")    { $files[$Matches[1] + '-setup'] = $n }
  elseif ($n -match "^$APP_SLUG-$escVer-(.+)-portable\.zip$") { $files[$Matches[1] + '-portable'] = $n }
  elseif ($n -match "^$APP_SLUG-$escVer-(.+)\.msi$")          { $files[$Matches[1] + '-msi'] = $n }
  elseif ($n -match "^$APP_SLUG-$escVer-(.+)\.appx$")         { $files[$Matches[1] + '-appx'] = $n }
}

$notesFile = Join-Path $PROJ "RELEASE_NOTES_v$VER.md"
$notes = if (Test-Path $notesFile) { ((Get-Content $notesFile -Encoding UTF8 | Select-Object -First 40) -join "`n") } else { "$APP_NAME $VER" }
$manifest = [ordered]@{
  app     = $APP_NAME
  version = $VER
  date    = (Get-Date -Format 'yyyy-MM-dd')
  notes   = $notes
  files   = $files
}
[IO.File]::WriteAllText((Join-Path $REL_WIN 'version.json'), ($manifest | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))

$shaLines = @("# SHA256 Checksums — $APP_NAME v$VER (Windows)", "", "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "", "| File | Size (bytes) | SHA256 |", "|---|---:|---|")
Get-ChildItem $REL_WIN -Recurse -File | Where-Object { $_.Name -ne 'SHA256_Windows_v0.1.8.md' } | Sort-Object FullName | ForEach-Object {
  $relPath = $_.FullName.Substring($REL_WIN.Length + 1).Replace('\', '/')
  $hash = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
  $shaLines += "| $relPath | $($_.Length) | ``$hash`` |"
}
[IO.File]::WriteAllText((Join-Path $REL_WIN "SHA256_Windows_v$VER.md"), ($shaLines -join "`r`n"), (New-Object Text.UTF8Encoding($false)))

Say "BUILD WINDOWS v$VER DONE ($($files.Count) manifest entries)"
