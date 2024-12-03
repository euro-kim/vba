Attribute VB_Name = "CanvasAddSubs"
Sub AddDot()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If
    
    dot_width = 0.2
    dot_height = 0.2
    
    Dim str As String
    
    Dim xval As Double
    Dim yval As Double
    Dim zval As Double
    
    Dim transformed_xval As Double
    Dim transformed_yval As Double
    
    Dim left_value As Double
    Dim top_value As Double
    
    'if a shape is selected in a xyplane, dot is added to the edges of the shape
    If IsChildShapeRange = True Then
        'cor(x,y)
        Dim cor(1 To 4, 1 To 2) As Double
        cor(1, 1) = Selection.ChildShapeRange.left: cor(1, 2) = Selection.ChildShapeRange.top
        cor(2, 1) = cor(1, 1) + Selection.ChildShapeRange.width: cor(2, 2) = cor(1, 2)
        cor(3, 1) = cor(1, 1): cor(3, 2) = cor(1, 2) + Selection.ChildShapeRange.height
        cor(4, 1) = cor(1, 1) + Selection.ChildShapeRange.width: cor(4, 2) = cor(1, 2) + Selection.ChildShapeRange.height

        For i = 1 To 4
            Set newDot = Selection.ShapeRange(1).CanvasItems.AddShape( _
                Type:=msoShapeOval, _
                left:=cor(i, 1) - CentimetersToPoints(dot_width / 2), _
                top:=cor(i, 2) - CentimetersToPoints(dot_height / 2), _
                width:=CentimetersToPoints(dot_width), _
                height:=CentimetersToPoints(dot_height))
            With newDot
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Visible = msoFalse
            End With
        Next
        Exit Sub
    End If

    If CanvasName = "space" Or CanvasName = "space2" Then
        'adds a dot to xyz space or xyz space2

        str = InputBox("Dot Generator", "Plese Input the (x,y,z) Coordinates", "(0,0,0)")
        If str = "" Then
                Exit Sub
        End If
        replacedStr = Replace(str, "(", "")
        replacedStr = Replace(replacedStr, ")", "")
        SplitStr = Split(replacedStr, ",")
        
        xval = CDbl(Trim(SplitStr(0)))
        yval = CDbl(Trim(SplitStr(1)))
        zval = CDbl(Trim(SplitStr(2)))
        
        If CanvasName = "space" Then
            Dim new_space As New xyzspace
            new_space.initialize
            
            'the transformed values are in centimeters.
            transformed_xval = new_space.xspacetoplane(xval, yval, zval)
            transformed_yval = new_space.yspacetoplane(xval, yval, zval)
            
            left_value = new_space.xcor(transformed_xval - (dot_width / 2))
            top_value = new_space.ycor(transformed_yval + (dot_height / 2))
        End If
        
        If CanvasName = "space2" Then
            Dim new_space2 As New xyzspace2
            new_space2.initialize
            
            left_value = new_space2.xcor(xval, yval, zval) - CentimetersToPoints(dot_width / 2)
            top_value = new_space2.ycor(xval, yval, zval) - CentimetersToPoints(dot_height / 2)
        End If

    Else
        'adds a dot to xyplane
        
        str = InputBox("Dot Generator", "Plese Input the (x,y) Coordinates", "(0,0)")
        If str = "" Then
                Exit Sub
        End If
        replacedStr = Replace(str, "(", "")
        replacedStr = Replace(replacedStr, ")", "")
        SplitStr = Split(replacedStr, ",")
        
        xval = CDbl(Trim(SplitStr(0)))
        yval = CDbl(Trim(SplitStr(1)))
        
        transformed_xval = xval
        transformed_yval = yval
    
        Dim new_plane As New xyplane
        new_plane.initialize
        
        left_value = new_plane.xcor(transformed_xval - (dot_width / 2))
        top_value = new_plane.ycor(transformed_yval + (dot_height / 2))

    End If
    
    Set newDot = Selection.ShapeRange(1).CanvasItems.AddShape( _
        Type:=msoShapeOval, _
        left:=left_value, _
        top:=top_value, _
        width:=CentimetersToPoints(dot_width), _
        height:=CentimetersToPoints(dot_height))
    With newDot
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Visible = msoFalse
        .Select
    End With
    
End Sub
Sub AddLine()
    If IsInlineCanvasShape = False Then
        Exit Sub
    End If

    Dim beg_str, end_str As String
    Dim beg_xval As Double
    Dim end_xval As Double
    Dim beg_yval As Double
    Dim end_yval As Double
    Dim beg_zval As Double
    Dim end_zval As Double
    
    Dim newLine As Shape
        
    beg_str = ""
    end_str = ""
    If CanvasName = "space" Then
        
        beg_str = InputBox("Line Generator", "Plese Input the beginning (x,y,z) Coordinates", "(2,1,0)")
        end_str = InputBox("Line Generator", "Plese Input the ending (x,y,z) Coordinates", "(3,2,1)")
        If beg_str = "" <> end_str = "" Then
                Exit Sub
        End If
        
        beg_replacedStr = Replace(beg_str, "(", "")
        beg_replacedStr = Replace(beg_replacedStr, ")", "")
        beg_SplitStr = Split(beg_replacedStr, ",")
        
        beg_xval = CDbl(Trim(beg_SplitStr(0)))
        beg_yval = CDbl(Trim(beg_SplitStr(1)))
        beg_zval = CDbl(Trim(beg_SplitStr(2)))
        
        end_replacedStr = Replace(end_str, "(", "")
        end_replacedStr = Replace(end_replacedStr, ")", "")
        end_SplitStr = Split(end_replacedStr, ",")
        
        end_xval = CDbl(Trim(end_SplitStr(0)))
        end_yval = CDbl(Trim(end_SplitStr(1)))
        end_zval = CDbl(Trim(end_SplitStr(2)))
        
        Dim new_space As New xyzspace
        new_space.initialize
        
        Set newLine = Selection.ShapeRange(1).CanvasItems.AddLine( _
             new_space.xcor(new_space.xspacetoplane(beg_xval, beg_yval, beg_zval)), _
             new_space.ycor(new_space.yspacetoplane(beg_xval, beg_yval, beg_zval)), _
             new_space.xcor(new_space.xspacetoplane(end_xval, end_yval, end_zval)), _
             new_space.ycor(new_space.yspacetoplane(end_xval, end_yval, end_zval)))
             
    Else
        Dim new_plane As New xyplane
        'if a Dot is selected, perpandicular foot is added
        If IsChildShapeRange = True Then
            If Selection.ChildShapeRange.AutoShapeType = 9 Then
                left_value = Selection.ChildShapeRange.left
                top_value = Selection.ChildShapeRange.top
                dot_width = 0.2
                dot_height = 0.2
    
                new_plane.initialize
                Set xfoot = Selection.ShapeRange(1).CanvasItems.AddLine( _
                     left_value + CentimetersToPoints(dot_width / 2), _
                     new_plane.ycor(0), _
                     left_value + CentimetersToPoints(dot_width / 2), _
                     top_value + CentimetersToPoints(dot_height / 2))
                     
                Set yfoot = Selection.ShapeRange(1).CanvasItems.AddLine( _
                     new_plane.xcor(0), _
                     top_value + CentimetersToPoints(dot_height / 2), _
                     left_value + CentimetersToPoints(dot_width / 2), _
                     top_value + CentimetersToPoints(dot_height / 2))
                
                With xfoot.Line
                    .DashStyle = msoLineRoundDot
                    .Weight = 1
                    .ForeColor.RGB = RGB(0, 0, 0)
                End With
        
                With yfoot.Line
                    .DashStyle = msoLineRoundDot
                    .Weight = 1
                    .ForeColor.RGB = RGB(0, 0, 0)
                End With
                Exit Sub
            End If
        End If
        
        beg_str = InputBox("Line Generator", "Plese Input the beginning (x,y) Coordinates", "(3,0)")
        end_str = InputBox("Line Generator", "Plese Input the ending (x,y) Coordinates", "(0,3)")
        If beg_str = "" <> end_str = "" Then
                Exit Sub
        End If
        On Error Resume Next
        beg_replacedStr = Replace(beg_str, "(", "")
        beg_replacedStr = Replace(beg_replacedStr, ")", "")
        beg_SplitStr = Split(beg_replacedStr, ",")
        
        beg_xval = CDbl(Trim(beg_SplitStr(0)))
        beg_yval = CDbl(Trim(beg_SplitStr(1)))
        
        end_replacedStr = Replace(end_str, "(", "")
        end_replacedStr = Replace(end_replacedStr, ")", "")
        end_SplitStr = Split(end_replacedStr, ",")
        
        end_xval = CDbl(Trim(end_SplitStr(0)))
        end_yval = CDbl(Trim(end_SplitStr(1)))
        
        new_plane.initialize
        
        Set newLine = Selection.ShapeRange(1).CanvasItems.AddLine( _
             new_plane.xcor(beg_xval), _
             new_plane.ycor(beg_yval), _
             new_plane.xcor(end_xval), _
             new_plane.ycor(end_yval))
    
    End If
 
    With newLine.Line
        .DashStyle = msoLineSolid
        .Weight = 1
        .ForeColor.RGB = RGB(0, 0, 0)
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadNone
         End With
    newLine.Select
    
End Sub
Sub AddVertexShape()
    Dim shp As Shape
    Dim innerShape As Shape
    Dim vertex As Shape
    
    Num = CountShapes
    Dim points() As Single
    
    ReDim points(1 To Num + 1, 1 To 2) As Single
    dot_width = 0.2
    dot_height = 0.2
    Dim left_value As Single
    Dim top_value As Single
    
    Counter = 0
    For Each innerShape In Selection.ChildShapeRange
        Counter = Counter + 1
        'for some reason scaling of 20 is necessary.
        left_value = 20 * innerShape.left
        top_value = 20 * innerShape.top
        points(Counter, 1) = left_value + CentimetersToPoints(dot_width / 2)
        points(Counter, 2) = top_value + CentimetersToPoints(dot_height / 2)
        innerShape.Delete
    Next
    points(Num + 1, 1) = points(1, 1)
    points(Num + 1, 2) = points(1, 2)
    

    Set shp = Selection.ShapeRange(1).CanvasItems.AddPolyline(points)
    
    With shp.Line
        .Visible = msoTrue ' Make the outline visible
        .Transparency = 1 ' Set outline to transparent
    End With

    With shp.Fill
        .Visible = msoTrue ' Make the fill visible
        .ForeColor.RGB = RGB(255, 0, 0) ' Set fill color (red)
    End With
    
    'Add the vertices again
    dot_width = 0.2
    dot_height = 0.2
    i = 0
    Do Until i = Num
    i = i + 1
        Set newDot = Selection.ShapeRange(1).CanvasItems.AddShape( _
            Type:=msoShapeOval, _
            left:=points(i, 1) - CentimetersToPoints(dot_width / 2), _
            top:=points(i, 2) - CentimetersToPoints(dot_height / 2), _
            width:=CentimetersToPoints(dot_width), _
            height:=CentimetersToPoints(dot_height))
        With newDot
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Visible = msoFalse
        End With
    Loop
    
End Sub

Sub AddCurve()
' IndifferenceCurve Macro
    Dim new_plane As New xyplane
    new_plane.initialize
    
    Dim points(0 To 3, 1 To 2) As Single
     points(0, 1) = 1:     points(0, 2) = 5
     points(1, 1) = 2:     points(1, 2) = 10
     points(2, 1) = 5:     points(2, 2) = 15
     points(3, 1) = 8:     points(3, 2) = 10
 
    Dim temp As Double
    For i = 0 To 3
        'convert Single -> Double -> Function -> Double -> Single
        temp = points(i, 1)
        points(i, 1) = CSng(new_plane.xcor(temp))
        temp = points(i, 2)
        points(i, 2) = CSng(new_plane.ycor(temp))
    Next
 
    Set newCurve = ActiveDocument.Shapes(Selection.ShapeRange.Name).CanvasItems.AddCurve(points)
    With newCurve.Line
        .Weight = 1
        .ForeColor.RGB = RGB(0, 0, 0)
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadNone
    End With
    newCurve.Select
End Sub
