<#
  Excel(26 입학 설명회 영상.xlsx)의 시트별 도형/사진/글자를 PowerPoint 슬라이드로 옮기고
  슬라이드마다 최대 4개 그룹(배경·도형·사진·글자)으로 묶어 화려한 등장/퇴장 애니메이션,
  자동 넘김(총 50초), 배경음악(mp3)을 넣은 PPT를 만든다.

  실행:  powershell -ExecutionPolicy Bypass -File .\make_ppt_animation.ps1
  동영상까지:  ... -File .\make_ppt_animation.ps1 -ExportVideo
#>
param(
    [string]$ExcelPath    = 'C:\박은영PDF\26 입학 설명회 영상.xlsx',
    [string]$MusicPath    = '',            # 비우면 같은 폴더의 첫 mp3 사용
    [string]$OutputPath   = '',            # 비우면 "<엑셀이름>_애니메이션.pptx"
    [double]$TotalSeconds = 50,
    [switch]$IncludeCells,                 # 셀 내용(표/글자)도 배경 그림으로 함께 넣기
    [switch]$ExportVideo                   # 저장 후 MP4(1080p)도 생성
)

$ErrorActionPreference = 'Stop'

# ---- Office 상수 ----
$msoTrue = -1; $msoFalse = 0
$msoGroup = 6; $msoPicture = 13; $msoLinkedPicture = 11; $msoTextBox = 17
$msoComment = 4; $msoFormControl = 8
$ppLayoutBlank = 12
$trigWithPrevious = 2
$fxAppear = 1; $fxFly = 2; $fxFade = 10; $fxSpiral = 15; $fxSwivel = 19
$fxWheel = 21; $fxZoom = 23; $fxBounce = 26; $fxFloat = 30; $fxGrowTurn = 31
$dirUp = 1; $dirRight = 2; $dirDown = 3; $dirLeft = 4
$ppEffectFadeSmoothly = 3849

# 그룹 종류별 등장 효과 후보 (슬라이드마다 돌아가며 사용해서 매번 다른 느낌)
$EntrancePool = @{
    bg    = @(@{Id=$fxFade}, @{Id=$fxZoom}, @{Id=$fxWheel})
    shape = @(@{Id=$fxFly;Dir=$dirLeft}, @{Id=$fxSpiral}, @{Id=$fxWheel}, @{Id=$fxFly;Dir=$dirRight})
    photo = @(@{Id=$fxSwivel}, @{Id=$fxBounce}, @{Id=$fxZoom}, @{Id=$fxFly;Dir=$dirRight})
    text  = @(@{Id=$fxFly;Dir=$dirDown}, @{Id=$fxFloat}, @{Id=$fxGrowTurn}, @{Id=$fxFly;Dir=$dirLeft})
}
$ExitPool = @(@{Id=$fxFade}, @{Id=$fxZoom}, @{Id=$fxFly;Dir=$dirUp}, @{Id=$fxFly;Dir=$dirLeft})
$GroupOrder = @('bg','shape','photo','text')

# ---- 파일 확인 ----
if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "엑셀 파일이 없습니다: $ExcelPath" }
$folder = Split-Path -LiteralPath $ExcelPath -Parent
if (-not $MusicPath) {
    $mp3 = Get-ChildItem -LiteralPath $folder -Filter *.mp3 | Select-Object -First 1
    if ($mp3) { $MusicPath = $mp3.FullName }
}
if (-not $OutputPath) {
    $OutputPath = Join-Path $folder (([IO.Path]::GetFileNameWithoutExtension($ExcelPath)) + '_애니메이션.pptx')
}
Write-Host "엑셀 : $ExcelPath"
Write-Host "음악 : $(if ($MusicPath) { $MusicPath } else { '(mp3 없음 - 음악 생략)' })"
Write-Host "출력 : $OutputPath"

# ---- 열려 있는 Excel 사용 (없으면 새로 열기) ----
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $xl = New-Object -ComObject Excel.Application; $xl.Visible = $true }
$wb = $null
foreach ($w in $xl.Workbooks) { if ($w.FullName -ieq $ExcelPath) { $wb = $w; break } }
if (-not $wb) { $wb = $xl.Workbooks.Open($ExcelPath) }

# ---- PowerPoint 새 프레젠테이션 (16:9) ----
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $msoTrue
$pres = $pp.Presentations.Add($msoTrue)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540
$SW = 960.0; $SH = 540.0; $Margin = 20.0

function Invoke-Paste($slide) {
    for ($t = 0; $t -lt 8; $t++) {
        try { return $slide.Shapes.Paste() } catch { Start-Sleep -Milliseconds 400 }
    }
    throw '붙여넣기에 실패했습니다 (클립보드).'
}

function Get-Items($range) {
    $list = @(); for ($i = 1; $i -le $range.Count; $i++) { $list += $range.Item($i) }; return $list
}

function Scale-Fonts($shp, [double]$k) {
    if ($shp.Type -eq $msoGroup) {
        for ($i = 1; $i -le $shp.GroupItems.Count; $i++) { Scale-Fonts $shp.GroupItems.Item($i) $k }
        return
    }
    try {
        if ($shp.HasTextFrame -and $shp.TextFrame2.HasText) {
            $runs = $shp.TextFrame2.TextRange.Runs()
            for ($i = 1; $i -le $runs.Count; $i++) { $r = $runs.Item($i); $r.Font.Size = [Math]::Max(6, $r.Font.Size * $k) }
        }
    } catch {}
}

function Test-HasPicture($shp) {
    if ($shp.Type -eq $msoPicture -or $shp.Type -eq $msoLinkedPicture) { return $true }
    if ($shp.Type -eq $msoGroup) {
        for ($i = 1; $i -le $shp.GroupItems.Count; $i++) { if (Test-HasPicture $shp.GroupItems.Item($i)) { return $true } }
    }
    try { if ($shp.Fill.Type -eq 6) { return $true } } catch {}   # 그림 채우기 도형
    return $false
}

function Get-OverlapRatio($a, $b) {   # a 면적 중 b와 겹치는 비율
    $w = [Math]::Min($a.Left + $a.Width, $b.Left + $b.Width) - [Math]::Max($a.Left, $b.Left)
    $h = [Math]::Min($a.Top + $a.Height, $b.Top + $b.Height) - [Math]::Max($a.Top, $b.Top)
    if ($w -le 0 -or $h -le 0) { return 0 }
    $area = [Math]::Max(1, $a.Width * $a.Height)
    return ($w * $h) / $area
}

function Test-IsText($shp) {
    try {
        if ($shp.Type -eq $msoTextBox) { return $true }
        if ($shp.HasTextFrame -and $shp.TextFrame2.HasText -and -not $shp.Fill.Visible) { return $true }
    } catch {}
    return $false
}

function Add-Fx($slide, $shape, $fx, [double]$delay, [double]$dur, [bool]$isExit) {
    $seq = $slide.TimeLine.MainSequence
    try { $e = $seq.AddEffect($shape, $fx.Id, 0, $trigWithPrevious) }
    catch { $e = $seq.AddEffect($shape, $fxFade, 0, $trigWithPrevious) }
    if ($fx.Dir) { try { $e.EffectParameters.Direction = $fx.Dir } catch {} }
    if ($isExit) { $e.Exit = $msoTrue }
    $e.Timing.Duration = $dur
    $e.Timing.TriggerDelayTime = $delay
    return $e
}

# ---- 시트 → 슬라이드 ----
$sheets = @()
foreach ($ws in $wb.Worksheets) {
    if ($ws.Visible -ne -1) { continue }
    $idx = @()
    for ($i = 1; $i -le $ws.Shapes.Count; $i++) {
        $t = $ws.Shapes.Item($i).Type
        if ($t -ne $msoComment -and $t -ne $msoFormControl) { $idx += $i }
    }
    $hasCells = $xl.WorksheetFunction.CountA($ws.UsedRange) -gt 0
    if ($idx.Count -gt 0 -or $hasCells) { $sheets += @{ Ws = $ws; Idx = $idx; Cells = $hasCells } }
}
if ($sheets.Count -eq 0) { throw '옮길 내용(도형/셀)이 있는 시트가 없습니다.' }

$N = $sheets.Count
$SlideDur = $TotalSeconds / $N
$Trans = 0.6
Write-Host "슬라이드 $N 장, 장당 약 $([Math]::Round($SlideDur,1))초"

$slideNo = 0
foreach ($info in $sheets) {
    $slideNo++
    $ws = $info.Ws
    Write-Host "[$slideNo/$N] 시트 '$($ws.Name)' 처리 중..."
    $slide = $pres.Slides.Add($slideNo, $ppLayoutBlank)
    $ws.Activate()

    $pasted = @()
    # 셀 내용: 도형이 없거나 -IncludeCells 이면 그림으로 넣음
    if ($info.Cells -and ($info.Idx.Count -eq 0 -or $IncludeCells)) {
        $ws.UsedRange.CopyPicture(1, -4147)
        $pasted += Get-Items (Invoke-Paste $slide)
    }
    if ($info.Idx.Count -gt 0) {
        $ws.Shapes.Range([object[]]$info.Idx).Select()
        $xl.Selection.Copy()
        Start-Sleep -Milliseconds 300
        $pasted += Get-Items (Invoke-Paste $slide)
    }

    # 전체를 슬라이드 크기에 맞춤 (상대 위치 유지)
    if ($pasted.Count -gt 1) {
        $tmp = $slide.Shapes.Range([object[]]($pasted | ForEach-Object { $_.Name })).Group()
    } else { $tmp = $pasted[0] }
    $k = [Math]::Min(($SW - 2*$Margin) / $tmp.Width, ($SH - 2*$Margin) / $tmp.Height)
    $tmp.LockAspectRatio = $msoTrue
    $tmp.Width = $tmp.Width * $k
    $tmp.Left = ($SW - $tmp.Width) / 2
    $tmp.Top  = ($SH - $tmp.Height) / 2
    if ([Math]::Abs($k - 1) -gt 0.05) { Scale-Fonts $tmp $k }
    if ($pasted.Count -gt 1) { $items = Get-Items ($tmp.Ungroup()) } else { $items = @($tmp) }

    $j = 0
    foreach ($s in $items) { $j++; $s.Name = "S${slideNo}_$j" }

    # ---- 4개 그룹으로 분류 ----
    $cat = @{ bg = @(); shape = @(); photo = @(); text = @() }
    $photos = @($items | Where-Object { Test-HasPicture $_ })
    $photoNames = @($photos | ForEach-Object { $_.Name })
    foreach ($s in $items) {
        if ($photoNames -contains $s.Name) { $cat.photo += $s; continue }
        $onPhoto = $false
        foreach ($p in $photos) {
            if ($p.Width * $p.Height -lt $s.Width * $s.Height * 0.8) { continue }   # 자기보다 작은 사진엔 안 붙임
            if ((Get-OverlapRatio $s $p) -ge 0.5) { $onPhoto = $true; break }
        }
        if ($onPhoto) { $cat.photo += $s }                                     # 사진에 붙은 테두리/설명
        elseif (Test-IsText $s) { $cat.text += $s }
        elseif ($s.Width * $s.Height -ge $SW * $SH * 0.25) { $cat.bg += $s }   # 큰 배경 도형
        else { $cat.shape += $s }
    }

    $groups = @()
    foreach ($key in $GroupOrder) {
        $members = $cat[$key]
        if ($members.Count -eq 0) { continue }
        if ($members.Count -gt 1) {
            $g = $slide.Shapes.Range([object[]]($members | ForEach-Object { $_.Name })).Group()
        } else { $g = $members[0] }
        $g.Name = "G${slideNo}_$key"
        $groups += @{ Key = $key; Shape = $g }
    }
    Write-Host ("    그룹: " + (($groups | ForEach-Object { "$($_.Key)" }) -join ', '))

    # ---- 애니메이션 (모두 '이전 효과와 함께' + 지연시간으로 절대 시점 지정) ----
    $inDur = 0.9; $inGap = 0.45; $outDur = 0.6; $outGap = 0.15
    $gi = 0
    foreach ($g in $groups) {
        $pool = $EntrancePool[$g.Key]
        $fx = $pool[($slideNo + $gi) % $pool.Count]
        [void](Add-Fx $slide $g.Shape $fx (0.2 + $gi * $inGap) $inDur $false)
        $gi++
    }
    $inEnd = 0.2 + ($groups.Count - 1) * $inGap + $inDur
    $outStart = $SlideDur - $Trans - 0.2 - ($groups.Count - 1) * $outGap - $outDur
    if ($slideNo -lt $N -and $outStart -gt $inEnd + 1.0) {
        $gi = 0
        foreach ($g in ($groups[($groups.Count-1)..0])) {   # 나중에 나온 그룹부터 퇴장
            $fx = $ExitPool[($slideNo + $gi) % $ExitPool.Count]
            [void](Add-Fx $slide $g.Shape $fx ($outStart + $gi * $outGap) $outDur $true)
            $gi++
        }
    }

    # ---- 전환 + 자동 넘김 ----
    $tr = $slide.SlideShowTransition
    try { $tr.EntryEffect = $ppEffectFadeSmoothly } catch {}
    try { $tr.Duration = $Trans } catch {}
    $tr.AdvanceOnClick = $msoFalse
    $tr.AdvanceOnTime = $msoTrue
    $tr.AdvanceTime = [Math]::Max(1, $SlideDur - $Trans)
}

# ---- 배경음악: 1번 슬라이드에서 시작해 끝까지 재생 ----
if ($MusicPath) {
    $s1 = $pres.Slides.Item(1)
    $m = $s1.Shapes.AddMediaObject2($MusicPath, $msoFalse, $msoTrue, -60, -60, 40, 40)
    $m.Name = 'BGM'
    $ps = $m.AnimationSettings.PlaySettings
    $ps.PlayOnEntry = $msoTrue
    $ps.HideWhileNotPlaying = $msoTrue
    $ps.LoopUntilStopped = $msoTrue
    $ps.StopAfterSlides = 999
    try {
        $mf = $m.MediaFormat
        $mf.FadeInDuration = 500
        $mf.FadeOutDuration = 2500
        if ($mf.Length -gt $TotalSeconds * 1000) { $mf.EndPoint = [int]($TotalSeconds * 1000) }
    } catch {}
    # 음악 재생 효과를 맨 앞으로, 슬라이드 시작과 동시에
    $seq = $s1.TimeLine.MainSequence
    for ($i = 1; $i -le $seq.Count; $i++) {
        $e = $seq.Item($i)
        if ($e.Shape.Name -eq 'BGM') { $e.MoveTo(1); $e.Timing.TriggerType = $trigWithPrevious; $e.Timing.TriggerDelayTime = 0; break }
    }
}

$pres.SaveAs($OutputPath)
Write-Host "저장 완료: $OutputPath"

# ---- (선택) MP4 내보내기 ----
if ($ExportVideo) {
    $mp4 = [IO.Path]::ChangeExtension($OutputPath, '.mp4')
    Write-Host "동영상 만드는 중: $mp4"
    $pres.CreateVideo($mp4, $true, 5, 1080, 30, 90)
    while ($pres.CreateVideoStatus -eq 1 -or $pres.CreateVideoStatus -eq 2) { Start-Sleep -Seconds 2; Write-Host -NoNewline '.' }
    Write-Host ''
    if ($pres.CreateVideoStatus -eq 3) { Write-Host "동영상 완료: $mp4" } else { Write-Host '동영상 만들기에 실패했습니다.' }
}
