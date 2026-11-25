# Yomi-Office Open Source v0.1.8 — kiểm thử Android trên thiết bị thật (-Serial) hoặc emulator (-Avd)
# Tự chọn APK theo API + CPU của thiết bị, quay video (scrcpy → .mkv → .mp4), chụp ảnh, chạy #selftest, mở tệp trong "Test Office",
# thử tab "Trên thiết bị", Trợ lý A.I, DeX. Kết quả: v018\test\android_<Tag>_*.png, v018\video\android\*.mp4, log PASS/FAIL.
param([string]$Serial = '', [string]$Avd = '', [string]$Tag = '', [int]$MaxFiles = 24, [switch]$Dex, [switch]$Tv)
$ErrorActionPreference = 'Continue'
$VER = '0.1.8'
$SRC = Split-Path -Parent $PSScriptRoot; $PROJ = Split-Path -Parent (Split-Path -Parent $SRC)
$T = "$PROJ\v018\shots\android"; $V = "$PROJ\v018\video\android"; $AOUT = Join-Path (Split-Path -Parent $PROJ) "output\v$VER\Android"
New-Item -ItemType Directory -Force -Path $T, $V | Out-Null
$ASDK = Join-Path $env:LOCALAPPDATA 'Android\Sdk'; $adb = "$ASDK\platform-tools\adb.exe"; $emu = "$ASDK\emulator\emulator.exe"
if (-not (Test-Path $adb)) { $adb = 'D:\yh\tools\android-sdk\platform-tools\adb.exe' }
$scrcpy = Get-ChildItem 'D:\yh\tools\scrcpy' -Recurse -Filter scrcpy.exe -EA SilentlyContinue | Select-Object -First 1 -Expand FullName
$ffmpeg = 'D:\yh\tools\ffmpeg-9.0.2-essentials_build\bin\ffmpeg.exe'
$pkg = 'com.officeopensource.office_open_source'
$env:ANDROID_SDK_ROOT = $ASDK; $env:ANDROID_HOME = $ASDK; $env:ANDROID_AVD_HOME = "$env:USERPROFILE\.android\avd"
foreach ($pv in 'HTTP_PROXY','HTTPS_PROXY','http_proxy','https_proxy','ALL_PROXY','all_proxy') { Remove-Item "Env:$pv" -EA SilentlyContinue }
if (-not $Tag) { $Tag = if ($Serial) { 'device' } else { ($Avd -replace '[^A-Za-z0-9]', '') } }
function Say($m) { Write-Output ("[{0}] {1}" -f (Get-Date -Format HH:mm:ss), $m) }
$pass = 0; $fail = 0; function Check($n, $ok, $i) { if ($ok) { $script:pass++; Say "PASS  $n $i" } else { $script:fail++; Say "FAIL  $n $i" } }
function A { $list = @($args)
  $q = ($list | ForEach-Object { $a = [string]$_; if ($a -match '[\s"]') { '"' + ($a -replace '"', '\"') + '"' } else { $a } }) -join ' '
  $psi = New-Object Diagnostics.ProcessStartInfo; $psi.FileName = $adb; $psi.Arguments = "-s $($script:dev) $q"; $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
  $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
  $p = [Diagnostics.Process]::Start($psi); $o = $p.StandardOutput.ReadToEndAsync(); $e = $p.StandardError.ReadToEndAsync()
  if (-not $p.WaitForExit(300000)) { try { $p.Kill() } catch {}; return '' }; return (($o.Result + $e.Result) -split "`r?`n" | Where-Object { $_ -ne '' }) }
function Shot($name) { $f = "$T\android_${Tag}_$name.png"; A shell screencap -p /data/local/tmp/oos_shot.png | Out-Null; A pull /data/local/tmp/oos_shot.png $f | Out-Null; return (Test-Path $f) }

& $adb start-server | Out-Null
$emuProc = $null
if ($Serial) { $script:dev = $Serial }
else {
  Get-Process emulator*, qemu* -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 3
  $cfg = "$env:USERPROFILE\.android\avd\$Avd.avd\config.ini"
  if (Test-Path $cfg) { $c = @(Get-Content $cfg | Where-Object { $_ -notmatch '^disk\.dataPartition\.size' }); $c += 'disk.dataPartition.size=6G'; Set-Content $cfg $c -Encoding ASCII }
  $emuProc = Start-Process $emu -ArgumentList @('-avd', $Avd, '-no-audio', '-no-snapshot', '-no-boot-anim', '-gpu', 'swiftshader_indirect', '-memory', $(if ($Avd -match 'API23') { '2048' } else { '4096' })) -PassThru -RedirectStandardError "$T\emu_${Tag}_err.log" -RedirectStandardOutput "$T\emu_${Tag}_out.log"
  $script:dev = $null
  for ($i = 0; $i -lt 60 -and -not $script:dev; $i++) { Start-Sleep 4; $l = (& $adb devices) | Where-Object { $_ -match '^(emulator-\d+)\s+device' } | Select-Object -First 1; if ($l) { $script:dev = ($l -split '\s+')[0] } }
  Check "emulator start ($Avd)" ([bool]$script:dev) $script:dev
  if (-not $script:dev) { exit 1 }
  $booted = $false; $dl = (Get-Date).AddMinutes($(if ($Avd -match 'API23') { 30 } else { 12 }))
  while (-not $booted -and (Get-Date) -lt $dl) { Start-Sleep 6; $booted = (((A shell getprop sys.boot_completed) -join '').Trim() -eq '1') }
  Check 'boot completed' $booted ''
  if (-not $booted) { exit 2 }
  Start-Sleep 25; A shell input keyevent 82 | Out-Null
}
$rel = ((A shell getprop ro.build.version.release) -join '').Trim(); $sdk = [int](((A shell getprop ro.build.version.sdk) -join '').Trim())
$abis = ((A shell getprop ro.product.cpu.abilist) -join '').Trim() -split ','; $model = ((A shell getprop ro.product.model) -join '').Trim()
Say "Thiết bị: $model — Android $rel (API $sdk) — CPU $($abis -join ',')"
$edVer = if ($sdk -ge 29) { 'Android10plus' } elseif ($sdk -ge 26) { 'Android8-9' } else { 'Android6-7' }
$abi = $abis | Where-Object { @('arm64-v8a', 'x86_64', 'armeabi-v7a') -contains $_ } | Select-Object -First 1   # theo thứ tự CPU của thiết bị (emulator x86_64 ưu tiên x86_64, không dùng arm64 giả lập)
$Apk = "$AOUT\Yomi-Office-Open-Source-$VER-$edVer-$abi.apk"
Check "chọn bộ cài phù hợp" (Test-Path $Apk) (Split-Path $Apk -Leaf)
if (-not (Test-Path $Apk)) { exit 3 }
A shell settings put global window_animation_scale 0 | Out-Null; A shell settings put global transition_animation_scale 0 | Out-Null; A shell settings put global animator_duration_scale 0 | Out-Null

# ---- quay video toàn bộ quá trình (scrcpy ghi .mkv: không hỏng khi dừng giữa chừng)
$rec = "$V\android_${Tag}_v$VER.mkv"; $recProc = $null
if (-not $Serial) {   # emulator: trình quay có sẵn của emulator, tối đa 180 s/lần → quay nhiều đoạn 175 s rồi ghép bằng ffmpeg
  $segDir = "$env:TEMP\oos_seg_$Tag";   # KHÔNG có dấu cách: lệnh console emulator cắt đường dẫn tại dấu cách
  Remove-Item $segDir -Recurse -Force -EA SilentlyContinue; New-Item -ItemType Directory -Force -Path $segDir | Out-Null
  Set-Content "$segDir\run.flag" '1'
  $recJob = Start-Job -ArgumentList $adb, $script:dev, $segDir -ScriptBlock { param($adb, $dev, $dir)
    $i = 0; while (Test-Path "$dir\run.flag") { $i++; $f = "$dir\seg{0:000}.webm" -f $i
      & $adb -s $dev emu screenrecord start --time-limit 175 $f | Out-Null
      for ($k = 0; $k -lt 176 -and (Test-Path "$dir\run.flag"); $k++) { Start-Sleep 1 }
      & $adb -s $dev emu screenrecord stop | Out-Null; Start-Sleep 1 } }
  Start-Sleep 4; $rec = "$V\android_${Tag}_v$VER.mp4"
  Say "đang quay video emulator (từng đoạn 175 s) → $segDir"
} elseif ($scrcpy) { Remove-Item $rec -EA SilentlyContinue; $recProc = Start-Process $scrcpy -ArgumentList @('-s', $script:dev, '--no-playback', '--no-audio', '--max-size', '1280', '--record', $rec) -PassThru -WindowStyle Hidden; Start-Sleep 4 }
if ($Serial) { Check 'bắt đầu quay video' ([bool]$recProc -and -not $recProc.HasExited) $rec }

# ---- cài đặt
$t0 = Get-Date; $r1 = (A install -r $Apk) -join ' '
Check "cài đặt $(Split-Path $Apk -Leaf) ($([math]::Round((Get-Item $Apk).Length/1MB,1)) MB)" ($r1 -match 'Success') ("{0} ({1:n0}s)" -f $r1, ((Get-Date) - $t0).TotalSeconds)
$dp = (A shell dumpsys package $pkg) -join "`n"
Check 'phiên bản 0.1.8' ($dp -match "versionName=$([regex]::Escape($VER))") ([regex]::Match($dp, 'versionName=\S+').Value + ' ' + [regex]::Match($dp, 'versionCode=\d+').Value + ' ' + [regex]::Match($dp, 'primaryCpuAbi=\S+').Value)
$lbl = (A shell cmd package resolve-activity --brief $pkg) -join ' '
if ($sdk -ge 30) { A shell appops set $pkg MANAGE_EXTERNAL_STORAGE allow | Out-Null } else { A shell pm grant $pkg android.permission.READ_EXTERNAL_STORAGE | Out-Null; A shell pm grant $pkg android.permission.WRITE_EXTERNAL_STORAGE | Out-Null }

# ---- đưa các tệp kiểm thử lên máy (thư mục Documents\YomiTest)
$tf = Get-ChildItem "$PROJ\Test Office" -File | Where-Object { $_.Name -notmatch '^~\$' } | Select-Object -First $MaxFiles
A shell mkdir -p /sdcard/Documents/YomiTest | Out-Null
$i = 0; foreach ($f in $tf) { $i++; $safe = ('t{0:00}{1}' -f $i, $f.Extension.ToLower()); A push $f.FullName "/sdcard/Documents/YomiTest/$safe" | Out-Null }
A shell am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file:///sdcard/Documents/YomiTest | Out-Null
Say "đã đưa $($tf.Count) tệp lên /sdcard/Documents/YomiTest"

# ---- màn hình chính
A logcat -c | Out-Null
A shell am start -W -n "$pkg/.MainActivity" | Out-Null; Start-Sleep 30
Check 'ảnh: màn hình chính' (Shot '01_home') ''
$title = (A shell dumpsys activity activities) -join ' '
# ---- tab "Trên thiết bị" (Feature Check #28)
A shell am start -W -n "$pkg/.MainActivity" --es hash 'tool=devscan' | Out-Null; Start-Sleep 35
Check 'ảnh: quét tệp Office trên thiết bị' (Shot '02_devscan') ''
$lg = (A logcat -d) -join "`n"; $m = [regex]::Match($lg, 'OOSDEVSCAN \| (\d+) files \| word=(\d+) excel=(\d+) ppt=(\d+)')
Check 'quét tệp Office: tìm thấy tệp kiểm thử' ($m.Success -and [int]$m.Groups[1].Value -ge [math]::Min($tf.Count, 1)) $m.Value
# ---- Trợ lý A.I (Free Claude Code có trong danh sách)
A shell am start -W -n "$pkg/.MainActivity" --es hash 'tool=ai' | Out-Null; Start-Sleep 20
Check 'ảnh: Trợ lý A.I' (Shot '03_ai') ''
# ---- tự kiểm tra (#selftest): Word / Excel / PowerPoint / PDF / VBA / định dạng / v0.1.8
A logcat -c | Out-Null
A shell am start -W -n "$pkg/.MainActivity" --es hash 'selftest' | Out-Null
$done = $null; for ($k = 0; $k -lt 60 -and -not $done; $k++) { Start-Sleep 5; $lg = (A logcat -d) -join "`n"; $mm = [regex]::Match($lg, 'OOSTEST DONE \| (\d+) passed, (\d+) failed'); if ($mm.Success) { $done = $mm } }
$lines = ([regex]::Matches($lg, 'OOSTEST (PASS|FAIL) \| [^\r\n"]+') | ForEach-Object { $_.Value }) | Select-Object -Unique
$lines | ForEach-Object { Say "   $_" }
Check 'selftest hoàn tất' ([bool]$done) $(if ($done) { $done.Value } else { 'không thấy OOSTEST DONE' })
Check 'selftest không có mục FAIL' ($done -and [int]$done.Groups[2].Value -eq 0) $(if ($done) { $done.Value })
Shot '04_selftest' | Out-Null
# ---- mở từng tệp Office thật bằng Intent VIEW (engine Native LibreOfficeKit)
$mime = @{ '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; '.doc' = 'application/msword'; '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; '.xlsm' = 'application/vnd.ms-excel.sheet.macroEnabled.12'; '.pptx' = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }
$i = 0
foreach ($f in $tf) {
  $i++; $safe = ('t{0:00}{1}' -f $i, $f.Extension.ToLower()); $mt = $mime[$f.Extension.ToLower()]; if (-not $mt) { continue }
  A shell am force-stop $pkg | Out-Null; A logcat -c | Out-Null; $t1 = Get-Date
  A shell am start -W -a android.intent.action.VIEW -d "file:///sdcard/Documents/YomiTest/$safe" -t $mt -n "$pkg/.MainActivity" | Out-Null
  $ok = $false; for ($k = 0; $k -lt 24 -and -not $ok; $k++) { Start-Sleep 5; $act = (A shell dumpsys activity activities) -join ' '; $ok = $act -match 'NativeDocActivity|DocActivity' }
  Start-Sleep 12
  $crash = (A logcat -d -b crash) -join ' '
  $nm = ('{0:00}_{1}' -f (10 + $i), ($f.BaseName -replace '[^\w\-]', '_').Substring(0, [math]::Min(30, ($f.BaseName -replace '[^\w\-]', '_').Length)))
  Shot $nm | Out-Null
  Check ("mở tệp {0:00}: {1}" -f $i, $f.Name) ($ok -and -not ($crash -match $pkg)) ("{0:n0}s" -f ((Get-Date) - $t1).TotalSeconds)
}
# ---- DeX / màn hình TV (giao diện desktop của ứng dụng ở 1920x1080)
if ($Dex) {
  A shell am force-stop $pkg | Out-Null
  A shell wm size 1920x1080 | Out-Null; A shell wm density 160 | Out-Null; Start-Sleep 4
  A shell am start -W -n "$pkg/.MainActivity" --es hash 'dex=tvhome' | Out-Null; Start-Sleep 30
  Check 'ảnh: DeX / TV desktop' (Shot '40_dex_desktop') ''
  A shell wm size reset | Out-Null; A shell wm density reset | Out-Null
}
# ---- Trình chiếu lên TV (màn hình phụ 1920x1080 giả lập qua overlay display — giống cắm HDMI/Miracast)
if ($Tv) {
  A shell am force-stop $pkg | Out-Null
  # emulator: thêm màn hình phụ THẬT (như cắm TV/HDMI); máy thật: overlay display
  if (-not $Serial) { A emu multidisplay add 1 1920 1080 240 0 | Out-Null } else { A shell settings put global overlay_display_devices 1920x1080/240 | Out-Null }; Start-Sleep 8
  # không phụ thuộc ngôn ngữ máy (OneUI tiếng Việt in "Lớp phủ #1"): đếm mDisplayId trong dumpsys display
  $dd = (A shell dumpsys display) -join "`n"
  $logical = @([regex]::Matches($dd, 'mDisplayId=(\d+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
  $ids = @([regex]::Matches(((A shell dumpsys SurfaceFlinger --display-id) -join "`n"), 'Display (\d+)') | ForEach-Object { $_.Groups[1].Value })
  Check 'có màn hình phụ (TV)' ($logical.Count -ge 2 -or $ids.Count -ge 2) ("logical: " + ($logical -join ',') + " | SF: " + ($ids -join ','))
  A shell am start -W -n "$pkg/.MainActivity" --es hash 'dex=presenter' | Out-Null; Start-Sleep 30
  Check 'ảnh: điện thoại khi trình chiếu' (Shot '50_tv_phone') ''
  if ($ids.Count -ge 2) { A shell screencap -d $ids[-1] -p /data/local/tmp/oos_tv.png | Out-Null; $f = "$T\android_${Tag}_51_tv_screen.png"; A pull /data/local/tmp/oos_tv.png $f | Out-Null; Check 'ảnh: màn hình TV' (Test-Path $f) '' }
  A shell am start -W -n "$pkg/.MainActivity" --es hash 'dex=tvhome' | Out-Null; Start-Sleep 25
  Check 'ảnh: giao diện TV home' (Shot '52_tv_home') ''
  if (-not $Serial) { A emu multidisplay del 1 | Out-Null } else { A shell settings delete global overlay_display_devices | Out-Null }
}
$crash = (A logcat -d -b crash) -join ' '; Check 'không văng ứng dụng (toàn bộ phiên)' (-not ($crash -match $pkg)) ''
# ---- dừng quay emulator, đổi sang .mp4
if (-not $Serial) { Remove-Item "$segDir\run.flag" -EA SilentlyContinue; Wait-Job $recJob -Timeout 20 | Out-Null; Stop-Job $recJob -EA SilentlyContinue; A emu screenrecord stop | Out-Null; Start-Sleep 4
  $segs = Get-ChildItem $segDir -Filter *.webm | Where-Object { $_.Length -gt 10KB } | Sort-Object Name
  ($segs | ForEach-Object { "file '" + $_.FullName.Replace("'", "'\''") + "'" }) | Set-Content "$segDir\list.txt" -Encoding ASCII
  & $ffmpeg -y -loglevel error -f concat -safe 0 -i "$segDir\list.txt" -c:v libx264 -preset veryfast -crf 30 -pix_fmt yuv420p -vf "scale='min(720,iw)':-2" $rec 2>&1 | Out-Null
  Check 'video kiểm thử' ((Test-Path $rec) -and (Get-Item $rec).Length -gt 100KB) ("{0} ({1} đoạn, {2:n1} MB)" -f (Split-Path $rec -Leaf), $segs.Count, ((Get-Item $rec -EA SilentlyContinue).Length / 1MB))
  if (Test-Path $rec) { Remove-Item $segDir -Recurse -Force -EA SilentlyContinue } }
# ---- dừng quay, đổi sang .mp4
if ($recProc) { try { Stop-Process -Id $recProc.Id -EA SilentlyContinue } catch {}; Start-Sleep 3
  $mp4 = [IO.Path]::ChangeExtension($rec, '.mp4'); & $ffmpeg -y -loglevel error -i $rec -c copy $mp4 2>&1 | Out-Null
  Check 'video kiểm thử' ((Test-Path $mp4) -and (Get-Item $mp4).Length -gt 100KB) ("{0} ({1:n1} MB)" -f (Split-Path $mp4 -Leaf), ((Get-Item $mp4 -EA SilentlyContinue).Length / 1MB)) }
Say "KẾT QUẢ $Tag ($model, Android $rel): $pass đạt, $fail lỗi"
if ($emuProc) { Get-Process emulator*, qemu* -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue }
