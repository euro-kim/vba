Attribute VB_Name = "CanvasFunctions"
'MsoShapeType 1 is MsoAutoShape
'MsoShapeType 7 is MsoFormControl, wdSelectionInlineShape
'MsoShapeType 8 is wdSelectionShapes
'MsoShapeType 9 is MsoLine
'MsoShapeType 20is Canvas

Public Function IsShape() As Boolean
    On Error Resume Next
    If Selection.Type = 8 Then
        IsShape = True
        Exit Function
    End If
    IsShape = False
End Function
Public Function IsInlineShape() As Boolean
    On Error Resume Next
    If Selection.Type = wdSelectionInlineShape Then
        IsInlineShape = True
        Exit Function
    End If
    IsInlineShape = False
End Function
Public Function IsCanvasShape() As Boolean
    On Error Resume Next
    If Selection.ShapeRange.Count = 0 Then
        IsCanvasShape = False
        Exit Function
    End If
    If Selection.ShapeRange(1).Type = msoCanvas Then
        IsCanvasShape = True
        Exit Function
    End If
    IsCanvasShape = False
End Function
Public Function IsChildShapeRange() As Boolean
    'checks if a shape inside a canvas has been selected.
    On Error Resume Next
    If Selection.ChildShapeRange.Count > 0 Then
        IsChildShapeRange = True
        Exit Function
    End If
    IsChildShapeRange = False
End Function
Public Function IsInlineCanvasShape() As Boolean
    IsInlineCanvasShape = False
    If IsCanvasShape = True Then
        IsInlineCanvasShape = True
        Exit Function
    ElseIf IsShape = True Then
        MsgBox "Shape has been selected, but it is not a Canvas"
        Exit Function
    ElseIf IsInlineShape = True Then
        MsgBox "InlineShape has been selected, but it is not a Canvas"
        Exit Function
    Else
        MsgBox "Please select a Canvas"
    End If
End Function
Public Function CanvasName() As String
    Dim canvasShape As Shape
    Dim innerShape As Shape

        ' Iterate over all shapes within the CanvasShape
        For Each canvasShape In Selection.ShapeRange
            CanvasName = canvasShape.Name
            Exit Function
        Next
End Function
Function CountShapes() As Integer
    'counts the number of shapes
    Dim Counter As Double
    Counter = 0
    Dim outerShape As Shape
    Dim innerShape As Shape
    
    If IsChildShapeRange = True Then
        'if Shapes inside the canvas have been selected, count the selected shapes
        For Each innerShape In Selection.ChildShapeRange
            Counter = Counter + 1
        Next
    ElseIf IsCanvasShape = True Then
        'if canvas has been selected, count the Shapes inside the canvas
        For Each innerShape In Selection.ShapeRange(1).CanvasItems
            Counter = Counter + 1
        Next
    Else
        'if normal shapes have been selected, count the selected shapes
        '이부분 작동안함
        For Each outerShape In Selection.ShapeRange.ParentGroup
            Counter = Counter + 1
        Next
    End If
CountShapes = Counter
End Function
