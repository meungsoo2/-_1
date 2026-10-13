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
