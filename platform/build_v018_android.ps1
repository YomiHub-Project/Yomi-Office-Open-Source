# Yomi-Office Open Source v0.1.8 — build giao diện web (3 profile) + 9 APK Android (3 phiên bản Android × 3 CPU)
# Phần Windows do Antigravity làm trong build_v018_windows.ps1. Chạy qua v018\runner.ps1 (job) hoặc trực tiếp.
param([switch]$SkipWeb, [switch]$SkipAndroid, [string]$Suffix = '')   # -Suffix '-a': build web vào dist-a… riêng cho Android, không đụng dist của bản Windows
$ErrorActionPreference = 'Continue'
$VER = '0.1.8'
$SRC = Split-Path -Parent $PSScriptRoot                     # ...\v018\src
$PROJ = Split-Path -Parent (Split-Path -Parent $SRC)        # gốc dự án (không viết cứng đường dẫn)
$OUT = Join-Path (Split-Path -Parent $PROJ) "output\v$VER"; $AOUT = "$OUT\Android"
New-Item -ItemType Directory -Force -Path $AOUT | Out-Null
function Say($m) { Write-Output ("[{0}] {1}" -f (Get-Date -Format HH:mm:ss), $m) }
Say "PROJ=$PROJ OUT=$OUT"

if (-not $SkipWeb) {
  $env:NODE_PATH = "$PROJ\v005\deps\node_modules"
  Push-Location $SRC
  foreach ($p in @(@('universal','dist'), @('modern','dist-modern'), @('legacy','dist-legacy'))) {
    Say "web build $($p[0]) -> $($p[1])"; $t = Get-Date
    node tools/build.js --profile=$($p[0]) --out=$($p[1])$Suffix 2>&1 | Select-Object -Last 8
    Say ("  exit={0} {1:n0}s index={2}" -f $LASTEXITCODE, ((Get-Date) - $t).TotalSeconds, (Test-Path "$($p[1])$Suffix\index.html"))
  }
  Pop-Location
  if (-not $Suffix) { Set-Content "$PROJ\v018\coord\WEB_DONE.txt" (Get-Date -Format o) -Encoding UTF8 }   # báo cho build Windows (Antigravity)
}

if (-not $SkipAndroid) {
  $env:JAVA_HOME = 'D:\yh\tools\jdk'; $env:ANDROID_HOME = 'D:\yh\tools\android-sdk'; $env:ANDROID_SDK_ROOT = $env:ANDROID_HOME; $env:GRADLE_USER_HOME = 'D:\yh\gradle-home'
  $env:PATH = "$env:JAVA_HOME\bin;$env:PATH"; $env:GRADLE_OPTS = '-Xmx6g -Dorg.gradle.jvmargs=-Xmx6g'
  $and = "$SRC\platform\android"
  # giao diện cho từng flavor (android8 dùng src\universal\assets theo sourceSets)
  $map = @{ legacy = 'dist-legacy'; universal = 'dist'; modern = 'dist-modern' }
  $ed = @{ legacy = 'Android6-7'; universal = 'Android8-9'; modern = 'Android10plus' }
  foreach ($fl in $map.Keys) {
    $www = "$and\app\src\$fl\assets\www"; if (Test-Path $www) { Remove-Item $www -Recurse -Force }
    New-Item -ItemType Directory -Force -Path (Split-Path $www) | Out-Null
    Copy-Item "$SRC\$($map[$fl])$Suffix" $www -Recurse
    Set-Content "$www\edition.txt" $ed[$fl] -Encoding ASCII
  }
  # v0.1.8: KHÔNG chép mipmap từ app\ Flutter cũ nữa (sẽ đè logo Yomi mới trong res\mipmap-*)
  Set-Content "$and\local.properties" ("sdk.dir=" + ($env:ANDROID_HOME -replace '\\', '\\')) -Encoding ASCII
  $gradle = Get-ChildItem 'D:\yh\gradle-home\wrapper\dists\gradle-8.3-all' -Recurse -Filter gradle.bat | Select-Object -First 1 -Expand FullName
  $tasks = 'assembleLegacyRelease', 'assembleAndroid8Release', 'assembleModernRelease'
  Push-Location $and; $t = Get-Date
  & $gradle --no-daemon --offline @tasks 2>&1 | Select-Object -Last 40
  $ge = $LASTEXITCODE
  if ($ge -ne 0) { Say "offline build failed ($ge), retry online"; & $gradle --no-daemon @tasks 2>&1 | Select-Object -Last 40; $ge = $LASTEXITCODE }
  Pop-Location
  Say ("gradle exit={0} ({1:n0} phút)" -f $ge, ((Get-Date) - $t).TotalMinutes)
  $names = @{ legacy = 'Android6-7'; android8 = 'Android8-9'; modern = 'Android10plus' }
  $aapt = Get-ChildItem 'D:\yh\tools\android-sdk\build-tools' -Recurse -Filter aapt.exe | Sort-Object FullName -Descending | Select-Object -First 1 -Expand FullName
  Get-ChildItem $AOUT -Filter *.apk -EA SilentlyContinue | Remove-Item -Force
  foreach ($fl in 'legacy', 'android8', 'modern') {
    foreach ($abi in 'arm64-v8a', 'armeabi-v7a', 'x86_64') {
      $apk = Get-ChildItem "$and\app\build\outputs\apk\$fl\release" -Filter "*$abi*.apk" -EA SilentlyContinue | Select-Object -First 1
      if (-not $apk) { Say "APK $fl/$abi KHÔNG TÌM THẤY"; continue }
      $dst = "$AOUT\Yomi-Office-Open-Source-$VER-$($names[$fl])-$abi.apk"; Copy-Item $apk.FullName $dst -Force
      $b = & $aapt dump badging $dst 2>&1 | Select-String "^(package|sdkVersion|native-code|application-label):"
      Say ("APK {0,-55} {1,6:n1} MB | {2}" -f (Split-Path $dst -Leaf), ((Get-Item $dst).Length / 1MB), (($b | ForEach-Object { $_.Line }) -join ' | '))
    }
  }
}
Say 'XONG'
