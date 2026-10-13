<#
  열려 있는 PPT(26 입학 설명회 영상.pptx)를 복사해서
  슬라이드마다 최대 4개 그룹(배경·도형·사진·글자)으로 묶어 화려한 등장/퇴장 애니메이션,
  자동 넘김(총 50초), 배경음악(같은 폴더의 mp3)을 넣은 새 PPT를 만든다. 원본은 건드리지 않는다.

  실행:        powershell -ExecutionPolicy Bypass -File .\make_ppt_animation.ps1
  동영상까지:  ... -File .\make_ppt_animation.ps1 -ExportVideo
#>
param(
    [string]$SourcePath   = 'C:\박은영PDF\26 입학 설명회 영상.pptx',
    [string]$MusicPath    = '',            # 비우면 같은 폴더의 첫 mp3 사용
    [string]$OutputPath   = '',            # 비우면 "<원본이름>_애니메이션.pptx"
    [double]$TotalSeconds = 50,
    [switch]$ExportVideo                   # 저장 후 MP4(1080p)도 생성
)

$ErrorActionPreference = 'Stop'

# ---- Office 상수 ----
$msoTrue = -1; $msoFalse = 0
$msoGroup = 6; $msoPicture = 13; $msoLinkedPicture = 11; $msoTextBox = 17
$msoMedia = 16; $msoPlaceholder = 14
$trigWithPrevious = 2
$fxFly = 2; $fxFade = 10; $fxSpiral = 15; $fxSwivel = 19
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

# ---- 원본 PPT 찾기: 열려 있는 PowerPoint → 지정 경로 → 폴더 검색 ----
$DefaultFolder = 'C:\박은영PDF'
try { $pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $pp = New-Object -ComObject PowerPoint.Application }
$pp.Visible = $msoTrue

$src = $null
foreach ($p in $pp.Presentations) { if ($p.FullName -ieq $SourcePath) { $src = $p; break } }
if (-not $src) {
    foreach ($p in $pp.Presentations) { if ($p.Name -like '*입학*' -and $p.Name -notlike '*_애니메이션*') { $src = $p; break } }
}
if (-not $src) {
    if (-not (Test-Path -LiteralPath $SourcePath)) {
        $found = Get-ChildItem -LiteralPath $DefaultFolder -Filter *.pptx -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -like '*입학*' -and $_.Name -notlike '*_애니메이션*' } | Select-Object -First 1
        if (-not $found) {
            Write-Host "[$DefaultFolder] 폴더의 파일 목록:"
            Get-ChildItem -LiteralPath $DefaultFolder -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "   $($_.Name)" }
            throw "원본 PPT를 찾지 못했습니다. PowerPoint에서 파일을 열어 둔 채로 다시 실행하세요."
        }
        $SourcePath = $found.FullName
    }
    $src = $pp.Presentations.Open($SourcePath)
}

$folder = if ($src.Path -and $src.Path -notmatch '^https?:') { $src.Path } else { $DefaultFolder }
if (-not $MusicPath) {
    $mp3 = Get-ChildItem -LiteralPath $folder -Filter *.mp3 -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($mp3) { $MusicPath = $mp3.FullName }
}
if (-not $OutputPath) {
    $OutputPath = Join-Path $folder (([IO.Path]::GetFileNameWithoutExtension($src.Name)) + '_애니메이션.pptx')
}
Write-Host "원본 : $($src.FullName)"
Write-Host "음악 : $(if ($MusicPath) { $MusicPath } else { '(mp3 없음 - 음악 생략)' })"
Write-Host "출력 : $OutputPath"

# 같은 이름의 결과 파일이 열려 있으면 닫기
foreach ($p in @($pp.Presentations)) { if ($p.FullName -ieq $OutputPath) { $p.Close() } }

# 원본은 그대로 두고 복사본에 작업
$src.SaveCopyAs($OutputPath)
$pres = $pp.Presentations.Open($OutputPath)
$SW = [double]$pres.PageSetup.SlideWidth
$SH = [double]$pres.PageSetup.SlideHeight

function Test-HasPicture($shp) {
    if ($shp.Type -eq $msoPicture -or $shp.Type -eq $msoLinkedPicture) { return $true }
    if ($shp.Type -eq $msoGroup) {
        for ($i = 1; $i -le $shp.GroupItems.Count; $i++) { if (Test-HasPicture $shp.GroupItems.Item($i)) { return $true } }
    }
    try { if ($shp.Type -eq $msoPlaceholder -and $shp.PlaceholderFormat.ContainedType -eq $msoPicture) { return $true } } catch {}
    try { if ($shp.Fill.Type -eq 6) { return $true } } catch {}   # 그림으로 채운 도형
    return $false
}

function Get-OverlapRatio($a, $b) {   # a 면적 중 b와 겹치는 비율
    $w = [Math]::Min($a.Left + $a.Width, $b.Left + $b.Width) - [Math]::Max($a.Left, $b.Left)
    $h = [Math]::Min($a.Top + $a.Height, $b.Top + $b.Height) - [Math]::Max($a.Top, $b.Top)
    if ($w -le 0 -or $h -le 0) { return 0 }
    return ($w * $h) / [Math]::Max(1, $a.Width * $a.Height)
}

function Test-IsText($shp) {
    try {
        if ($shp.Type -eq $msoTextBox) { return $true }
        if ($shp.HasTextFrame -and $shp.TextFrame2.HasText -and -not $shp.Fill.Visible) { return $true }
        if ($shp.Type -eq $msoPlaceholder -and $shp.HasTextFrame -and $shp.TextFrame2.HasText) { return $true }
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
}

# ---- 대상 슬라이드 (숨긴 슬라이드 제외) ----
$slides = @()
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $s = $pres.Slides.Item($i)
    if ($s.SlideShowTransition.Hidden -ne $msoTrue) { $slides += $s }
}
$N = $slides.Count
if ($N -eq 0) { throw '슬라이드가 없습니다.' }
$SlideDur = $TotalSeconds / $N
$Trans = 0.6
Write-Host "슬라이드 $N 장, 장당 약 $([Math]::Round($SlideDur,1))초"

$slideNo = 0
foreach ($slide in $slides) {
    $slideNo++
    Write-Host "[$slideNo/$N] 슬라이드 처리 중..."

    # 기존 애니메이션 지우기
    $seq = $slide.TimeLine.MainSequence
    for ($i = $seq.Count; $i -ge 1; $i--) { $seq.Item($i).Delete() }

    # 대상 도형 (오디오/비디오, 빈 개체 틀 제외). 이름이 겹치지 않게 정리
    $items = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $s = $slide.Shapes.Item($i)
        if ($s.Type -eq $msoMedia) { continue }
        try { if ($s.Type -eq $msoPlaceholder -and $s.HasTextFrame -and -not $s.TextFrame2.HasText) { continue } } catch {}
        $s.Name = "S${slideNo}_$i"
        $items += $s
    }

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

    # 그룹으로 묶기. 개체 틀/표처럼 묶을 수 없는 건 묶지 않고 같은 타이밍으로 함께 움직임
    $groups = @()
    foreach ($key in $GroupOrder) {
        $members = @($cat[$key])
        if ($members.Count -eq 0) { continue }
        $targets = $members
        if ($members.Count -gt 1) {
            try {
                $g = $slide.Shapes.Range([object[]]($members | ForEach-Object { $_.Name })).Group()
                $g.Name = "G${slideNo}_$key"
                $targets = @($g)
            } catch {
                $groupable = @($members | Where-Object { $_.Type -ne $msoPlaceholder -and -not $_.HasTable })
                $rest      = @($members | Where-Object { $_.Type -eq $msoPlaceholder -or $_.HasTable })
                if ($groupable.Count -gt 1) {
                    try {
                        $g = $slide.Shapes.Range([object[]]($groupable | ForEach-Object { $_.Name })).Group()
                        $g.Name = "G${slideNo}_$key"
                        $targets = @($g) + $rest
                    } catch { $targets = $members }
                }
            }
        }
        $groups += @{ Key = $key; Targets = $targets }
    }
    Write-Host ("    그룹: " + (($groups | ForEach-Object { "$($_.Key)($($_.Targets.Count))" }) -join ', '))

    # ---- 애니메이션 (모두 '이전 효과와 함께' + 지연시간으로 시점 지정) ----
    $inDur = 0.9; $inGap = 0.45; $outDur = 0.6; $outGap = 0.15
    $gi = 0
    foreach ($g in $groups) {
        $pool = $EntrancePool[$g.Key]
        $fx = $pool[($slideNo + $gi) % $pool.Count]
        foreach ($t in $g.Targets) { Add-Fx $slide $t $fx (0.2 + $gi * $inGap) $inDur $false }
        $gi++
    }
    if ($groups.Count -gt 0) {
        $inEnd = 0.2 + ($groups.Count - 1) * $inGap + $inDur
        $outStart = $SlideDur - $Trans - 0.2 - ($groups.Count - 1) * $outGap - $outDur
        if ($slideNo -lt $N -and $outStart -gt $inEnd + 1.0) {
            $gi = 0
            foreach ($g in ($groups[($groups.Count-1)..0])) {   # 나중에 나온 그룹부터 퇴장
                $fx = $ExitPool[($slideNo + $gi) % $ExitPool.Count]
                foreach ($t in $g.Targets) { Add-Fx $slide $t $fx ($outStart + $gi * $outGap) $outDur $true }
                $gi++
            }
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

# ---- 배경음악: 첫 슬라이드에서 시작해 끝까지 재생 ----
if ($MusicPath) {
    $s1 = $slides[0]
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

$pres.Save()
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
