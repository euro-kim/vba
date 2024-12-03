Attribute VB_Name = "CanvasChangeSubs"
Sub ChangeFillColor()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
    
    ' black, red, blue, purple, green
    Dim colors(1 To 5) As Long
    colors(1) = RGB(0, 0, 0)
    colors(2) = RGB(255, 0, 0)
    colors(3) = RGB(0, 112, 192)
    colors(4) = RGB(112, 48, 160)
    colors(5) = RGB(0, 176, 80)
    
    Dim old_color As Long
    old_color = Selection.ChildShapeRange.Fill.ForeColor.RGB
    
    Dim i As Integer
    Dim hit As Boolean
    hit = False
    i = 0
    
    On Error Resume Next
    Do Until hit = True
        i = i + 1
        If Selection.ChildShapeRange.Fill.Transparency = 1 Then
            Selection.ChildShapeRange.Fill.Transparency = 0
            Selection.ChildShapeRange.Fill.ForeColor.RGB = colors(1)
            hit = True
        ElseIf i = UBound(colors) Then
            Selection.ChildShapeRange.Fill.Transparency = 1
            hit = True
        ElseIf old_color = colors(i) Then
            Selection.ChildShapeRange.Fill.ForeColor.RGB = colors(i + 1)
            hit = True
        End If
    Loop
End Sub
Sub ChangeFillPattern()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
    
    Dim patterns(1 To 5) As Long
    patterns(1) = msoPattern70Percent
    patterns(2) = msoPatternDownwardDiagonal
    patterns(3) = msoPatternUpwardDiagonal
    patterns(4) = msoPatternSolidDiamond
    patterns(5) = msoPattern10Percent

    
    Dim old_pattern As Long
    old_pattern = Selection.ChildShapeRange.Fill.Pattern
    
    Dim i As Integer
    Dim hit As Boolean
    hit = False
    i = 0
    
    On Error Resume Next
    Do Until hit = True
        i = i + 1
        If old_pattern = msoPatternMixed Then
            Selection.ChildShapeRange.Fill.Patterned patterns(1)
            hit = True
        ElseIf i = UBound(patterns) Then
            Selection.ChildShapeRange.Fill.Solid
            hit = True
        ElseIf old_pattern = patterns(i) Then
            Selection.ChildShapeRange.Fill.Patterned patterns(i + 1)
            hit = True
        End If
    Loop
End Sub
Sub ChangeLineColor()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
     
    Dim colors(1 To 5) As Long
    ' black, red, blue, purple, green, black
    colors(1) = RGB(0, 0, 0)
    colors(2) = RGB(255, 0, 0)
    colors(3) = RGB(0, 112, 192)
    colors(4) = RGB(112, 48, 160)
    colors(5) = RGB(0, 176, 80)
    
    Dim old_color As Long
    old_color = Selection.ChildShapeRange.Line.ForeColor.RGB
        
    Dim i As Integer
    Dim hit As Boolean
    hit = False
    i = 0

    Do While Not hit
        i = i + 1
        If Selection.ChildShapeRange.Line.Transparency = 1 Then
            Selection.ChildShapeRange.Line.Transparency = 0
            Selection.ChildShapeRange.Line.ForeColor.RGB = colors(1)
            hit = True
        ElseIf i = UBound(colors) Then
            Selection.ChildShapeRange.Line.Transparency = 1
            hit = True
        ElseIf old_color = colors(i) Then
            Selection.ChildShapeRange.Line.ForeColor.RGB = colors(i + 1)
            hit = True
        End If
    Loop
End Sub
Sub ChangeDashStyle()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
    
    Dim arr(0 To 4) As MsoLineDashStyle
    arr(0) = msoLineSolid
    arr(1) = msoLineDash
    arr(2) = msoLineSquareDot
    arr(3) = msoLineRoundDot
    arr(4) = msoLineSolid
    
    Dim hit As Boolean
    Dim i As Integer
    i = 0
    
    'Example of Use of Do Until Loop, instead of using For Next
    Do Until hit = True
        If Selection.ChildShapeRange.Line.DashStyle = arr(i) Then
            Selection.ChildShapeRange.Line.DashStyle = arr(i + 1)
            hit = True
        End If
        If i = UBound(arr) Then
            Selection.ChildShapeRange.Line.DashStyle = arr(0)
            hit = True
        End If
        i = i + 1
    Loop
    
End Sub

Sub ChangeArrow()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
    
    BeginArrowheadStyle = Selection.ChildShapeRange.Line.BeginArrowheadStyle
    EndArrowheadStyle = Selection.ChildShapeRange.Line.EndArrowheadStyle
    
    If BeginArrowheadStyle = msoArrowheadNone And EndArrowheadStyle = msoArrowheadNone Then
            Selection.ChildShapeRange.Line.BeginArrowheadStyle = msoArrowheadOval
            Selection.ChildShapeRange.Line.EndArrowheadStyle = msoArrowheadOval
            Exit Sub
    End If
    
    If BeginArrowheadStyle = msoArrowheadOval And EndArrowheadStyle = msoArrowheadOval Then
        Selection.ChildShapeRange.Line.BeginArrowheadStyle = msoArrowheadOval
        Selection.ChildShapeRange.Line.EndArrowheadStyle = msoArrowheadNone
        Exit Sub
    End If
    
    If BeginArrowheadStyle = msoArrowheadOval And EndArrowheadStyle = msoArrowheadNone Then
        Selection.ChildShapeRange.Line.BeginArrowheadStyle = msoArrowheadNone
        Selection.ChildShapeRange.Line.EndArrowheadStyle = msoArrowheadOval
        Exit Sub
    End If
    
    If BeginArrowheadStyle = msoArrowheadNone And EndArrowheadStyle = msoArrowheadOval Then
        Selection.ChildShapeRange.Line.BeginArrowheadStyle = msoArrowheadNone
        Selection.ChildShapeRange.Line.EndArrowheadStyle = msoArrowheadNone
        Exit Sub
    End If

End Sub
