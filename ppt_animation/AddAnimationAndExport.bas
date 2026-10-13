Attribute VB_Name = "AddAnimationAndExport"
' ============================================================
'  깔끔한 애니메이션 자동 적용 + MP4 동영상 내보내기
'  - 슬라이드 전환: 부드럽게 페이드
'  - 개체 등장: 서로 겹치거나 맞닿은 도형+사진+글자를 한 그룹으로 묶어, 위 -> 아래 순서로 페이드
'  - 슬라이드를 꽉 채우는 배경 그림은 애니메이션 없음
'  - 각 슬라이드 표시 시간 = 애니메이션 끝난 후 SLIDE_HOLD 초
' ============================================================
Option Explicit

' ---- 원하는 대로 조정하세요 ----
Private Const SLIDE_HOLD As Single = 2       ' 애니메이션 끝난 뒤 머무는 시간(초)
Private Const ANIM_DURATION As Single = 0.4  ' 그룹 하나가 나타나는 시간(초)
Private Const ANIM_GAP As Single = 0.15      ' 그룹 사이 간격(초)
Private Const TRANSITION_SEC As Single = 0.4 ' 슬라이드 전환 시간(초)
Private Const GROUP_MARGIN As Single = 5     ' 이 거리(pt) 안에서 겹치거나 맞닿으면 같은 그룹 (클수록 크게 묶임)
Private Const VIDEO_HEIGHT As Long = 1080     ' 1080 = Full HD, 720 = HD
Private Const ANIM_SUFFIX As String = "_애니메이션" ' 새 PPT 파일 이름 뒤에 붙는 말
Private Const MUSIC_VOLUME As Single = 0.5   ' 배경음악 볼륨 (0 ~ 1)
Private Const MUSIC_FADE_SEC As Single = 2   ' 음악 시작/끝 페이드 시간(초)
Private Const VIDEO_SUFFIX As String = "_동영상"    ' 동영상 파일 이름 뒤에 붙는 말
' ------------------------------

Public Sub ApplyCleanAnimations()
    Dim sld As Slide
    Dim pres As Presentation
    Dim baseName As String, newPath As String
    Set pres = ActivePresentation

    ' 원본을 보호하기 위해 다른 이름으로 먼저 저장 (예: 원본이름_애니메이션.pptx)
    If pres.Path = "" Then
        MsgBox "먼저 파일을 저장해 주세요.", vbExclamation
        Exit Sub
    End If
    baseName = Left(pres.Name, InStrRev(pres.Name, ".") - 1)
    If Right(baseName, Len(ANIM_SUFFIX)) <> ANIM_SUFFIX Then
        newPath = pres.Path & "\" & baseName & ANIM_SUFFIX & ".pptx"
        pres.SaveAs FileName:=newPath, FileFormat:=ppSaveAsOpenXMLPresentation
    End If

    For Each sld In pres.Slides
        ApplyToSlide sld, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    Next sld

    SetupMusicPlayback pres   ' 배경음악이 있으면 재생 설정 다시 걸기
    pres.Save
    MsgBox "애니메이션 적용 완료! (" & pres.Slides.Count & "개 슬라이드)" & vbCrLf & _
           "저장된 파일: " & pres.FullName & vbCrLf & "(원본 파일은 그대로입니다)" & vbCrLf & _
           "F5로 확인해 보신 뒤 ExportVideo 를 실행하세요.", vbInformation
End Sub

Private Sub ApplyToSlide(sld As Slide, slideW As Single, slideH As Single)
    Dim seq As Sequence
    Dim i As Long, j As Long, n As Long, g As Long, k As Long
    Dim idx() As Long, par() As Long
    Dim L() As Single, T() As Single, R() As Single, B() As Single
    Dim gTop() As Single, gLeft() As Single, gId() As Long, gOrder() As Long
    Dim nG As Long, tmp As Long, first As Boolean
    Dim shp As Shape
    Dim eff As Effect
    Dim totalTime As Single

    ' 1) 기존 애니메이션 제거 (지저분함 방지)
    Set seq = sld.TimeLine.MainSequence
    For i = seq.Count To 1 Step -1
        seq.Item(i).Delete
    Next i

    ' 2) 애니메이션 대상 개체 수집 + 위치 기록
    n = 0
    ReDim idx(1 To sld.Shapes.Count + 1)
    ReDim L(1 To sld.Shapes.Count + 1): ReDim T(1 To sld.Shapes.Count + 1)
    ReDim R(1 To sld.Shapes.Count + 1): ReDim B(1 To sld.Shapes.Count + 1)
    ReDim par(1 To sld.Shapes.Count + 1)
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If ShouldAnimate(shp, slideW, slideH) Then
            n = n + 1
            idx(n) = i
            L(n) = shp.Left: T(n) = shp.Top
            R(n) = shp.Left + shp.Width: B(n) = shp.Top + shp.Height
            par(n) = n
        End If
    Next i
    If n = 0 Then GoTo SetTransition

    ' 3) 서로 겹치거나 맞닿은 도형/사진/글자를 한 그룹으로 묶기
    For i = 1 To n - 1
        For j = i + 1 To n
            If L(i) <= R(j) + GROUP_MARGIN And L(j) <= R(i) + GROUP_MARGIN And _
               T(i) <= B(j) + GROUP_MARGIN And T(j) <= B(i) + GROUP_MARGIN Then
                par(FindRoot(par, i)) = FindRoot(par, j)
            End If
        Next j
    Next i

    ' 4) 그룹별 위치(가장 위/왼쪽) 계산
    nG = 0
    ReDim gId(1 To n): ReDim gTop(1 To n): ReDim gLeft(1 To n): ReDim gOrder(1 To n)
    For i = 1 To n
        k = FindRoot(par, i)
        g = 0
        For j = 1 To nG
            If gId(j) = k Then g = j: Exit For
        Next j
        If g = 0 Then
            nG = nG + 1: g = nG
            gId(g) = k: gTop(g) = T(i): gLeft(g) = L(i): gOrder(g) = g
        Else
            If T(i) < gTop(g) Then gTop(g) = T(i)
            If L(i) < gLeft(g) Then gLeft(g) = L(i)
        End If
    Next i

    ' 5) 그룹 순서: 위 -> 아래, 같은 줄이면 왼쪽 -> 오른쪽
    For i = 1 To nG - 1
        For j = i + 1 To nG
            If GroupBefore(gTop(gOrder(j)), gLeft(gOrder(j)), gTop(gOrder(i)), gLeft(gOrder(i))) Then
                tmp = gOrder(i): gOrder(i) = gOrder(j): gOrder(j) = tmp
            End If
        Next j
    Next i

    ' 6) 그룹 단위 페이드 인 (그룹 안의 개체는 동시에)
    For g = 1 To nG
        first = True
        For i = 1 To n
            If FindRoot(par, i) = gId(gOrder(g)) Then
                If first Then
                    Set eff = seq.AddEffect(Shape:=sld.Shapes(idx(i)), effectId:=msoAnimEffectFade, _
                                            trigger:=msoAnimTriggerAfterPrevious)
                    If g = 1 Then
                        eff.Timing.TriggerDelayTime = 0.1
                    Else
                        eff.Timing.TriggerDelayTime = ANIM_GAP
                    End If
                    first = False
                Else
                    Set eff = seq.AddEffect(Shape:=sld.Shapes(idx(i)), effectId:=msoAnimEffectFade, _
                                            trigger:=msoAnimTriggerWithPrevious)
                End If
                eff.Timing.Duration = ANIM_DURATION
            End If
        Next i
    Next g

SetTransition:
    ' 7) 슬라이드 전환 + 자동 넘김 시간
    totalTime = 0.1 + nG * ANIM_DURATION + IIf(nG > 1, (nG - 1) * ANIM_GAP, 0) + SLIDE_HOLD
    With sld.SlideShowTransition
        .EntryEffect = ppEffectFadeSmoothly
        .Duration = TRANSITION_SEC
        .AdvanceOnClick = msoFalse
        .AdvanceOnTime = msoTrue
        .AdvanceTime = totalTime
    End With
End Sub

Private Function FindRoot(par() As Long, ByVal x As Long) As Long
    Do While par(x) <> x
        x = par(x)
    Loop
    FindRoot = x
End Function

Private Function GroupBefore(aTop As Single, aLeft As Single, bTop As Single, bLeft As Single) As Boolean
    If Abs(aTop - bTop) <= 20 Then
        GroupBefore = (aLeft < bLeft)
    Else
        GroupBefore = (aTop < bTop)
    End If
End Function

Private Function ShouldAnimate(shp As Shape, slideW As Single, slideH As Single) As Boolean
    ShouldAnimate = False
    If Not shp.Visible Then Exit Function
    If shp.Type = msoMedia Then Exit Function   ' 배경음악 아이콘 제외

    ' 슬라이드의 80% 이상을 덮는 배경 개체는 제외
    If shp.Width * shp.Height >= 0.8 * slideW * slideH Then Exit Function

    ' 비어 있는 자리표시자(텍스트 없음) 제외
    If shp.Type = msoPlaceholder Then
        If shp.HasTextFrame Then
            If Len(Trim(shp.TextFrame.TextRange.Text)) = 0 Then Exit Function
        End If
    End If

    ShouldAnimate = True
End Function

' 배경음악 넣기: ApplyCleanAnimations 다음에 실행하세요
Public Sub AddBackgroundMusic()
    Dim pres As Presentation
    Dim fd As FileDialog
    Dim musicPath As String
    Dim snd As Shape
    Dim seq As Sequence
    Dim i As Long

    Set pres = ActivePresentation

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "배경음악 파일을 선택하세요 (mp3, wav, m4a)"
    fd.Filters.Clear
    fd.Filters.Add "음악 파일", "*.mp3;*.wav;*.m4a;*.wma"
    If fd.Show <> -1 Then Exit Sub
    musicPath = fd.SelectedItems(1)

    ' 이미 넣은 배경음악이 있으면 지우고 새로 넣기
    For i = pres.Slides(1).Shapes.Count To 1 Step -1
        If pres.Slides(1).Shapes(i).Name = "BackgroundMusic" Then pres.Slides(1).Shapes(i).Delete
    Next i

    ' 1번 슬라이드에 넣고 (파일에 포함), 화면 밖에 숨김
    Set snd = pres.Slides(1).Shapes.AddMediaObject2(musicPath, msoFalse, msoTrue, -100, -100, 30, 30)
    snd.Name = "BackgroundMusic"

    With snd.MediaFormat
        .Volume = MUSIC_VOLUME
        .FadeInDuration = MUSIC_FADE_SEC * 1000
        .FadeOutDuration = MUSIC_FADE_SEC * 1000
    End With

    SetupMusicPlayback pres

    pres.Save
    MsgBox "배경음악을 넣었습니다." & vbCrLf & "F5로 들어 보신 뒤 ExportVideo 를 실행하세요.", vbInformation
End Sub

' 배경음악 자동 재생 설정 (클릭 없이 시작, 끝까지 재생, 반복)
Private Sub SetupMusicPlayback(pres As Presentation)
    Dim snd As Shape
    Dim seq As Sequence
    Dim eff As Effect
    Dim i As Long

    For i = 1 To pres.Slides(1).Shapes.Count
        If pres.Slides(1).Shapes(i).Name = "BackgroundMusic" Then Set snd = pres.Slides(1).Shapes(i)
    Next i
    If snd Is Nothing Then Exit Sub

    ' 이 음악에 걸린 기존 효과(클릭 대기 등) 모두 제거
    Set seq = pres.Slides(1).TimeLine.MainSequence
    For i = seq.Count To 1 Step -1
        If seq.Item(i).Shape.Name = "BackgroundMusic" Then seq.Item(i).Delete
    Next i

    ' 맨 앞(Index 1)에 "이전 효과와 함께" 재생 효과 추가 -> 슬라이드 시작 즉시 재생
    Set eff = seq.AddEffect(Shape:=snd, effectId:=msoAnimEffectMediaPlay, _
                            trigger:=msoAnimTriggerWithPrevious, Index:=1)
    With eff.EffectInformation.PlaySettings
        .StopAfterSlides = 999        ' 마지막 슬라이드까지 계속
        .LoopUntilStopped = msoTrue   ' 음악이 짧으면 반복
        .HideWhileNotPlaying = msoTrue
        .PauseAnimation = msoFalse
    End With
End Sub

Public Sub ExportVideo()
    Dim outPath As String
    Dim pres As Presentation
    Set pres = ActivePresentation

    If pres.Path = "" Then
        MsgBox "먼저 파일을 저장해 주세요.", vbExclamation
        Exit Sub
    End If

    Dim baseName As String
    baseName = Left(pres.Name, InStrRev(pres.Name, ".") - 1)
    If Right(baseName, Len(ANIM_SUFFIX)) = ANIM_SUFFIX Then baseName = Left(baseName, Len(baseName) - Len(ANIM_SUFFIX))
    outPath = pres.Path & "\" & baseName & VIDEO_SUFFIX & ".mp4"

    ' UseTimingsAndNarrations=True -> 위에서 설정한 시간대로 녹화
    pres.CreateVideo FileName:=outPath, _
                     UseTimingsAndNarrations:=True, _
                     DefaultSlideDuration:=5, _
                     VertResolution:=VIDEO_HEIGHT, _
                     FramesPerSecond:=30, _
                     Quality:=100

    MsgBox "동영상 만드는 중입니다. 화면 아래 상태 표시줄에서 진행률을 확인하세요." & vbCrLf & _
           "완료되면 다음 위치에 저장됩니다:" & vbCrLf & outPath, vbInformation
End Sub
