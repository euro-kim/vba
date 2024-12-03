VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xyzspace2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private private_width, private_height As Double
Private private_origin As Double
Private private_info As String

Public Property Let width(InCentimeters As Double)
    private_width = InCentimeters
End Property
Public Property Get width() As Double
    width = private_width
End Property
Public Property Let height(InCentimeters As Double)
    private_height = InCentimeters
End Property
Public Property Get height() As Double
    height = private_height
End Property
Public Property Let Info(InputString As String)
    private_info = InputString
End Property
Public Property Get Info() As String
    Info = private_info
End Property
Public Property Let Origin(InCentimeters As Double)
    private_origin = InCentimeters
End Property
Public Property Get Origin() As Double
    Origin = private_origin
End Property
Public Function xcor(xvar As Double, yvar As Double, zvar As Double) As Double
    xcor = CentimetersToPoints(Me.width / 2) + CentimetersToPoints(xvar) - (1 / Sqr(2)) * CentimetersToPoints(yvar)
End Function
Public Function ycor(xvar As Double, yvar As Double, zvar As Double) As Double
    ycor = CentimetersToPoints(Me.height / 2) + (1 / Sqr(2)) * CentimetersToPoints(yvar) - CentimetersToPoints(zvar)
End Function
Public Sub initialize()
    Me.width = 7.9
    Me.height = 6
    'the origin: (3.95,3)
End Sub

Public Sub build()
    If IsInlineShape = True Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    End If
    

    
    Dim newCanvas As Shape
    Dim objRange As Range
    Dim objEq As OMath

    Set newCanvas = ActiveDocument.Shapes.AddCanvas( _
        left:=Selection.Information(wdHorizontalPositionRelativeToPage) - Me.width, _
        top:=Selection.Information(wdVerticalPositionRelativeToPage), _
        width:=CentimetersToPoints(Me.width), _
        height:=CentimetersToPoints(Me.height))

    With newCanvas
        .Name = "space2"
        .WrapFormat.Type = wdWrapInline
        .WrapFormat.DistanceTop = CentimetersToPoints(0)
        .WrapFormat.DistanceBottom = CentimetersToPoints(0)
        .WrapFormat.AllowOverlap = False
    End With
    
    Dim newCanvasShapes As CanvasShapes
    Set newCanvasShapes = newCanvas.CanvasItems
    
    'Adds the X axis
     Set Xaxis = newCanvasShapes.AddLine( _
        CentimetersToPoints(0.3), _
        CentimetersToPoints(Me.height / 2), _
        CentimetersToPoints(Me.width - 0.3), _
        CentimetersToPoints(Me.height / 2))
    Xaxis.Name = "xaxis"
    
     With Xaxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
        .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadStyle = msoArrowheadNone
     End With
     
     'Adds the Y axis
     Set Yaxis = newCanvasShapes.AddLine( _
        CentimetersToPoints(Me.width - 0.7) - CentimetersToPoints(1.9) / 2, _
        CentimetersToPoints(0.7), _
        CentimetersToPoints(0.7) + CentimetersToPoints(1.9) / 2, _
        CentimetersToPoints(Me.height - 0.7))
    
    Yaxis.Name = "yaxis"

     With Yaxis.Line

        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
        .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadStyle = msoArrowheadNone
     End With

     Set Zaxis = newCanvasShapes.AddLine( _
        CentimetersToPoints(Me.width / 2), _
        CentimetersToPoints(Me.height - 0.3), _
        CentimetersToPoints(Me.width / 2), _
        CentimetersToPoints(0.3))

    Zaxis.Name = "zaxis"
    
     With Zaxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
        .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadStyle = msoArrowheadNone
     End With

    TextboxSize = CentimetersToPoints(1)
    
    'Creating the box for x
    Set Xbox = newCanvasShapes.AddShape( _
        msoShapeRectangle, _
        CentimetersToPoints(Me.width - 0.3), _
        CentimetersToPoints(Me.height / 2), _
        TextboxSize, _
        TextboxSize)
        
    With Xbox
        .Name = "xbox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.TextRange = ""
        .TextFrame.TextRange.Font.ColorIndex = wdBlack
    End With

     'Creating math equation form of x
    Set objRange = Xbox.TextFrame.TextRange
    objRange.Text = "x"
    Set objRange = Xbox.TextFrame.TextRange.OMaths.Add(objRange)
    
        'Creating the box for y
    Set Ybox = newCanvasShapes.AddShape( _
        msoShapeRectangle, _
        CentimetersToPoints(1), _
        CentimetersToPoints(Me.height) - (1 / Sqr(2)) * CentimetersToPoints(1), _
        TextboxSize, _
        TextboxSize)
        
    With Ybox
        .Name = "ybox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.TextRange = ""
        .TextFrame.TextRange.Font.ColorIndex = wdBlack
    End With

     'Creating math equation form of y
    Set objRange = Ybox.TextFrame.TextRange
    objRange.Text = "y"
    Set objRange = Ybox.TextFrame.TextRange.OMaths.Add(objRange)
    
            'Creating the box for z
    Set Zbox = newCanvasShapes.AddShape( _
        msoShapeRectangle, _
        CentimetersToPoints(Me.width / 2) - CentimetersToPoints(0.3), _
        CentimetersToPoints(0), _
        TextboxSize, _
        TextboxSize)
        
    With Zbox
        .Name = "zbox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.TextRange = ""
        .TextFrame.TextRange.Font.ColorIndex = wdBlack
    End With

     'Creating math equation form of z
    Set objRange = Zbox.TextFrame.TextRange
    objRange.Text = "z"
    Set objRange = Zbox.TextFrame.TextRange.OMaths.Add(objRange)
    
    Set newLine1 = newCanvasShapes.AddLine( _
         xcor(0.3, 0, 0), _
         ycor(0.3, 0, 0), _
         xcor(0.3, 0.3, 0), _
         ycor(0.3, 0.3, 0))
    newLine1.Line.ForeColor.RGB = RGB(0, 0, 0)
    newLine1.Name = "angle_xy_ver"
    newLine1.Line.Weight = 1
    
    Set newLine2 = newCanvasShapes.AddLine( _
         xcor(0, 0.3, 0), _
         ycor(0, 0.3, 0), _
         xcor(0.3, 0.3, 0), _
         ycor(0.3, 0.3, 0))
    newLine2.Line.ForeColor.RGB = RGB(0, 0, 0)
    newLine2.Name = "angle_xy_hor"
    newLine2.Line.Weight = 1
         
    Set newLine3 = newCanvasShapes.AddLine( _
         xcor(0.3, 0, 0), _
         ycor(0.3, 0, 0), _
         xcor(0.3, 0, 0.3), _
         ycor(0.3, 0, 0.3))
    newLine3.Line.ForeColor.RGB = RGB(0, 0, 0)
    newLine3.Name = "angle_xz_ver"
    newLine3.Line.Weight = 1
    
    Set newLine4 = newCanvasShapes.AddLine( _
         xcor(0, 0, 0.3), _
         ycor(0, 0, 0.3), _
         xcor(0.3, 0, 0.3), _
         ycor(0.3, 0, 0.3))
    newLine4.Line.ForeColor.RGB = RGB(0, 0, 0)
    newLine4.Name = "angle_xz_hor"
    newLine4.Line.Weight = 1
    

End Sub








