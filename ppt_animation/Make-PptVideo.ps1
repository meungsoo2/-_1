<#
  열려 있는 PowerPoint 를 COM 으로 조작해서
   1. 원본은 두고 "_애니메이션.pptx" 로 다른 이름 저장
   2. 모든 슬라이드 페이드 전환 + 자동 넘김
   3. 겹치거나 맞닿은 도형/사진/글자를 그룹으로 묶어 위->아래 페이드인
   4. 1번 슬라이드에 배경음악 (모든 슬라이드에 걸쳐 재생, 반복, 아이콘 숨김, 자동 시작)
   5. "_동영상.mp4" 1080p 내보내기

  실행 (PowerShell):
    powershell -ExecutionPolicy Bypass -File .\Make-PptVideo.ps1
    powershell -ExecutionPolicy Bypass -File .\Make-PptVideo.ps1 -Music "C:\음악\bgm.mp3"
#>
param(
    [string]$Source = "C:\박은영PDF\박은영동영상만들기261013.pptx",
    [string]$Music = "",
    [double]$Transition = 0.4,   # 슬라이드 전환(초)
    [double]$AnimDuration = 0.4, # 그룹 하나 나타나는 시간(초)
    [double]$AnimGap = 0.15,     # 그룹 사이 간격(초)
    [double]$Hold = 2,           # 애니메이션 끝난 뒤 머무는 시간(초)
    [double]$GroupMargin = 5,    # 이 거리(pt) 안이면 같은 그룹
    [double]$MusicVolume = 0.5,  # 0 ~ 1
    [int]$VideoHeight = 1080,
    [switch]$NoVideo
)
$ErrorActionPreference = "Stop"

# ---- PowerPoint 상수 ----
$msoTrue = -1; $msoFalse = 0
$msoMedia = 16; $msoPlaceholder = 14
$msoAnimEffectFade = 10; $msoAnimEffectMediaPlay = 83
$msoAnimTriggerOnPageClick = 1; $msoAnimTriggerWithPrevious = 2; $msoAnimTriggerAfterPrevious = 3
$ppEffectFadeSmoothly = 3849
$ppSaveAsOpenXMLPresentation = 24
$MusicName = "BackgroundMusic"

# ---- 1. PowerPoint 연결 + 다른 이름으로 저장 ----
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $null
foreach ($p in $ppt.Presentations) {
    if ($p.FullName -ieq $Source) { $pres = $p }
}
if (-not $pres) {
    if (-not (Test-Path -LiteralPath $Source)) { throw "파일이 없습니다: $Source" }
    Write-Host "열려 있지 않아서 파일을 엽니다: $Source"
    $pres = $ppt.Presentations.Open($Source)
}

$dir = Split-Path -Parent $Source
$base = [IO.Path]::GetFileNameWithoutExtension($Source)
$animPath = Join-Path $dir ($base + "_애니메이션.pptx")
$videoPath = Join-Path $dir ($base + "_동영상.mp4")

$pres.SaveAs($animPath, $ppSaveAsOpenXMLPresentation)
Write-Host "[1] 다른 이름으로 저장: $animPath (원본은 그대로)"

$slideW = $pres.PageSetup.SlideWidth
$slideH = $pres.PageSetup.SlideHeight

function Test-Animate($shp) {
    if ($shp.Visible -eq $msoFalse) { return $false }
    if ($shp.Type -eq $msoMedia) { return $false }
    if ($shp.Width * $shp.Height -ge 0.8 * $slideW * $slideH) { return $false }   # 배경
    if ($shp.Type -eq $msoPlaceholder -and $shp.HasTextFrame -eq $msoTrue) {
        if ($shp.TextFrame.TextRange.Text.Trim().Length -eq 0) { return $false }  # 빈 자리표시자
    }
    return $true
}

# ---- 2, 3. 슬라이드별 전환 + 그룹 애니메이션 ----
$totalGroups = 0
foreach ($sld in $pres.Slides) {
    $seq = $sld.TimeLine.MainSequence
    for ($i = $seq.Count; $i -ge 1; $i--) { $seq.Item($i).Delete() }

    # 대상 개체 수집
    $items = @()
    for ($i = 1; $i -le $sld.Shapes.Count; $i++) {
        $shp = $sld.Shapes.Item($i)
        if (Test-Animate $shp) {
            $items += [pscustomobject]@{ Shape = $shp; L = $shp.Left; T = $shp.Top
                R = $shp.Left + $shp.Width; B = $shp.Top + $shp.Height; G = $items.Count }
        }
    }

    # 겹치거나 맞닿은 개체를 같은 그룹으로 (union-find)
    $par = @(0..([Math]::Max($items.Count - 1, 0)))
    function Find-Root([int]$x) { while ($par[$x] -ne $x) { $x = $par[$x] }; return $x }
    for ($i = 0; $i -lt $items.Count; $i++) {
        for ($j = $i + 1; $j -lt $items.Count; $j++) {
            $a = $items[$i]; $b = $items[$j]
            if ($a.L -le $b.R + $GroupMargin -and $b.L -le $a.R + $GroupMargin -and
                $a.T -le $b.B + $GroupMargin -and $b.T -le $a.B + $GroupMargin) {
                $par[(Find-Root $i)] = (Find-Root $j)
            }
        }
    }
    for ($i = 0; $i -lt $items.Count; $i++) { $items[$i].G = Find-Root $i }

    # 그룹 순서: 위 -> 아래, 같은 줄(20pt 이내)이면 왼쪽 -> 오른쪽
    $groups = $items | Group-Object G | ForEach-Object {
        [pscustomobject]@{ Top = ($_.Group | Measure-Object T -Minimum).Minimum
                           Left = ($_.Group | Measure-Object L -Minimum).Minimum
                           Members = $_.Group }
    }
    $groups = @($groups | Sort-Object @{ Expression = { [Math]::Round($_.Top / 20) } }, Left)

    $g = 0
    foreach ($grp in $groups) {
        $g++
        $first = $true
        foreach ($m in $grp.Members) {
            if ($first) {
                $eff = $seq.AddEffect($m.Shape, $msoAnimEffectFade, 0, $msoAnimTriggerAfterPrevious, -1)
                $eff.Timing.TriggerDelayTime = $(if ($g -eq 1) { 0.1 } else { $AnimGap })
                $first = $false
            } else {
                $eff = $seq.AddEffect($m.Shape, $msoAnimEffectFade, 0, $msoAnimTriggerWithPrevious, -1)
            }
            $eff.Timing.Duration = $AnimDuration
        }
    }
    $totalGroups += $groups.Count

    $n = $groups.Count
    $t = 0.1 + $n * $AnimDuration + [Math]::Max($n - 1, 0) * $AnimGap + $Hold
    $tr = $sld.SlideShowTransition
    $tr.EntryEffect = $ppEffectFadeSmoothly
    $tr.Duration = $Transition
    $tr.AdvanceOnClick = $msoFalse
    $tr.AdvanceOnTime = $msoTrue
    $tr.AdvanceTime = $t
    Write-Host ("  슬라이드 {0}: 그룹 {1}개, {2:0.0}초" -f $sld.SlideIndex, $n, $t)
}
Write-Host "[2,3] 전환 + 그룹 애니메이션 적용 완료 (총 그룹 $totalGroups 개)"

# ---- 4. 배경음악 ----
$slide1 = $pres.Slides.Item(1)
if (-not $Music) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "배경음악 파일 선택 (취소하면 음악 없이 진행)"
    $dlg.Filter = "음악 파일|*.mp3;*.wav;*.m4a;*.wma|모든 파일|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Music = $dlg.FileName }
}

if ($Music) {
    for ($i = $slide1.Shapes.Count; $i -ge 1; $i--) {
        if ($slide1.Shapes.Item($i).Name -eq $MusicName) { $slide1.Shapes.Item($i).Delete() }
    }
    $snd = $slide1.Shapes.AddMediaObject2($Music, $msoFalse, $msoTrue, 10, 10, 30, 30)
    $snd.Name = $MusicName
    $snd.MediaFormat.Volume = $MusicVolume
    $snd.MediaFormat.Muted = $msoFalse
    $snd.MediaFormat.FadeInDuration = 2000
    $snd.MediaFormat.FadeOutDuration = 2000

    # "백그라운드에서 재생" 과 같은 설정: 맨 앞, 이전 효과와 함께(자동 시작), 끝까지, 반복, 숨김
    $seq = $slide1.TimeLine.MainSequence
    for ($i = $seq.Count; $i -ge 1; $i--) {
        if ($seq.Item($i).Shape.Name -eq $MusicName) { $seq.Item($i).Delete() }
    }
    $eff = $seq.AddEffect($snd, $msoAnimEffectMediaPlay, 0, $msoAnimTriggerWithPrevious, 1)
    $eff.Timing.TriggerDelayTime = 0
    $ps = $eff.EffectInformation.PlaySettings
    $ps.StopAfterSlides = 999
    $ps.LoopUntilStopped = $msoTrue
    $ps.HideWhileNotPlaying = $msoTrue
    $ps.PauseAnimation = $msoFalse

    # ---- 재생 설정 확인 ----
    $chk = $seq.Item(1)
    $ok = ($chk.Shape.Name -eq $MusicName) -and
          ($chk.Timing.TriggerType -ne $msoAnimTriggerOnPageClick) -and
          ($chk.EffectInformation.PlaySettings.StopAfterSlides -ge $pres.Slides.Count) -and
          ($snd.MediaFormat.Muted -eq $msoFalse) -and ($snd.MediaFormat.Length -gt 0)
    Write-Host "[4] 배경음악 확인"
    Write-Host ("    파일: {0}" -f $Music)
    Write-Host ("    길이: {0:0.0}초, 볼륨: {1}, 음소거: {2}" -f ($snd.MediaFormat.Length / 1000), $snd.MediaFormat.Volume, ($snd.MediaFormat.Muted -eq $msoTrue))
    Write-Host ("    애니메이션 1번째: {0} / 시작: {1} / 재생 슬라이드 수: {2} / 반복: {3}" -f `
        $chk.Shape.Name, $(if ($chk.Timing.TriggerType -eq $msoAnimTriggerWithPrevious) { "이전 효과와 함께(자동)" } else { "클릭할 때(!)" }), `
        $chk.EffectInformation.PlaySettings.StopAfterSlides, ($chk.EffectInformation.PlaySettings.LoopUntilStopped -eq $msoTrue))
    if ($ok) { Write-Host "    => 재생 설정 정상" -ForegroundColor Green }
    else { Write-Host "    => 재생 설정에 문제가 있습니다. 위 내용을 Claude 에게 알려 주세요." -ForegroundColor Red }
} else {
    Write-Host "[4] 음악 선택 안 함 - 건너뜀"
}

$pres.Save()

# ---- 5. 동영상 내보내기 ----
if ($NoVideo) { Write-Host "[5] -NoVideo 지정 - 동영상 건너뜀"; return }
Write-Host "[5] 동영상 만드는 중... (슬라이드 수에 따라 몇 분 걸립니다)"
$pres.CreateVideo($videoPath, $true, 5, $VideoHeight, 30, 100)
# 0=None 1=InProgress 2=Queued 3=Done 4=Failed
do {
    Start-Sleep -Seconds 2
    $st = $pres.CreateVideoStatus
    Write-Host -NoNewline "."
} while ($st -eq 1 -or $st -eq 2)
Write-Host ""
if ($st -eq 3) { Write-Host "[5] 완료: $videoPath" -ForegroundColor Green }
else { Write-Host "[5] 동영상 만들기 실패 (상태 $st)" -ForegroundColor Red }
