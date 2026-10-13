Attribute VB_Name = "AddAnimationAndExport"
' ============================================================
'  깔끔한 애니메이션 자동 적용 + MP4 동영상 내보내기
'  - 슬라이드 전환: 부드럽게 페이드 (0.7초)
'  - 개체 등장: 위 -> 아래, 왼쪽 -> 오른쪽 순서로 "나타내기(페이드)"
'  - 슬라이드를 꽉 채우는 배경 그림은 애니메이션 없음
'  - 각 슬라이드 표시 시간 = 애니메이션 끝난 후 SLIDE_HOLD 초
' ============================================================
Option Explicit

' ---- 원하는 대로 조정하세요 ----
Private Const SLIDE_HOLD As Single = 3        ' 애니메이션 끝난 뒤 머무는 시간(초)
Private Const ANIM_DURATION As Single = 0.6   ' 개체 하나가 나타나는 시간(초)
Private Const ANIM_GAP As Single = 0.3        ' 개체 사이 간격(초)
Private Const TRANSITION_SEC As Single = 0.7  ' 슬라이드 전환 시간(초)
Private Const VIDEO_HEIGHT As Long = 1080     ' 1080 = Full HD, 720 = HD
' ------------------------------

Public Sub ApplyCleanAnimations()
    Dim sld As Slide
    Dim pres As Presentation
    Set pres = ActivePresentation

    For Each sld In pres.Slides
        ApplyToSlide sld, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
    Next sld

    MsgBox "애니메이션 적용 완료! (" & pres.Slides.Count & "개 슬라이드)" & vbCrLf & _
           "F5로 확인해 보신 뒤 ExportVideo 를 실행하세요.", vbInformation
End Sub

Private Sub ApplyToSlide(sld As Slide, slideW As Single, slideH As Single)
    Dim seq As Sequence
    Dim i As Long, j As Long, n As Long
    Dim idx() As Long, tmp As Long
    Dim shp As Shape
    Dim eff As Effect
    Dim totalTime As Single

    ' 1) 기존 애니메이션 제거 (지저분함 방지)
    Set seq = sld.TimeLine.MainSequence
    For i = seq.Count To 1 Step -1
        seq.Item(i).Delete
    Next i

    ' 2) 애니메이션 대상 개체 수집
    n = 0
    ReDim idx(1 To sld.Shapes.Count + 1)
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        If ShouldAnimate(shp, slideW, slideH) Then
            n = n + 1
            idx(n) = i
        End If
    Next i

    ' 3) 위 -> 아래, 왼쪽 -> 오른쪽 순으로 정렬
    For i = 1 To n - 1
        For j = i + 1 To n
            If IsBefore(sld.Shapes(idx(j)), sld.Shapes(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    ' 4) 페이드 인 효과 추가 (모두 자동 진행 = 클릭 불필요)
    For i = 1 To n
        Set eff = seq.AddEffect(Shape:=sld.Shapes(idx(i)), _
                                effectId:=msoAnimEffectFade, _
                                trigger:=msoAnimTriggerAfterPrevious)
        eff.Timing.Duration = ANIM_DURATION
        If i = 1 Then
            eff.Timing.TriggerDelayTime = 0.3
        Else
            eff.Timing.TriggerDelayTime = ANIM_GAP
        End If
    Next i

    ' 5) 슬라이드 전환 + 자동 넘김 시간
    totalTime = 0.3 + n * ANIM_DURATION + IIf(n > 1, (n - 1) * ANIM_GAP, 0) + SLIDE_HOLD
    With sld.SlideShowTransition
        .EntryEffect = ppEffectFadeSmoothly
        .Duration = TRANSITION_SEC
        .AdvanceOnClick = msoFalse
        .AdvanceOnTime = msoTrue
        .AdvanceTime = totalTime
    End With
End Sub

Private Function ShouldAnimate(shp As Shape, slideW As Single, slideH As Single) As Boolean
    ShouldAnimate = False
    If Not shp.Visible Then Exit Function

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

Private Function IsBefore(a As Shape, b As Shape) As Boolean
    ' 같은 줄(세로 차이 20pt 이내)이면 왼쪽 먼저, 아니면 위쪽 먼저
    If Abs(a.Top - b.Top) <= 20 Then
        IsBefore = (a.Left < b.Left)
    Else
        IsBefore = (a.Top < b.Top)
    End If
End Function

Public Sub ExportVideo()
    Dim outPath As String
    Dim pres As Presentation
    Set pres = ActivePresentation

    If pres.Path = "" Then
        MsgBox "먼저 파일을 저장해 주세요.", vbExclamation
        Exit Sub
    End If

    outPath = pres.Path & "\" & Left(pres.Name, InStrRev(pres.Name, ".") - 1) & "_animated.mp4"

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
