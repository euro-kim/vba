VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xyplane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private private_width, private_height As Double
Private private_axis_edge_distance, private_axis_edge_distance_beg, private_axis_edge_distance_end As Double
Private private_info As String
Private private_origin As Double
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
'canvas_edge, axis_beg distance
Public Property Let AxisEdgeDistanceBeg(InCentimeters As Double)
    private_axis_edge_distance_beg = InCentimeters
End Property
Public Property Get AxisEdgeDistanceBeg() As Double
    AxisEdgeDistanceBeg = private_axis_edge_distance_beg
End Property
'axis_end, canvas_edge distance
Public Property Let AxisEdgeDistanceEnd(InCentimeters As Double)
    private_axis_edge_distance_end = InCentimeters
End Property
Public Property Get AxisEdgeDistanceEnd() As Double
    AxisEdgeDistanceEnd = private_axis_edge_distance_end
End Property
Public Property Let Origin(InCentimeters As Double)
    private_origin = InCentimeters
End Property
Public Property Get Origin() As Double
    Origin = private_origin
End Property
Public Function xcor(InCentimeters As Double) As Double
    xcor = CentimetersToPoints(Me.Origin + InCentimeters)
End Function
Public Function ycor(InCentimeters As Double) As Double
    ycor = CentimetersToPoints(Me.height - Origin - InCentimeters)
End Function
Public Sub initialize()
    Me.width = 7.9
    Me.height = 6
    'origin=(0.8,6-0.8=5.2)
    Origin = 0.8
    Me.AxisEdgeDistanceBeg = 0.3
    Me.AxisEdgeDistanceEnd = 0.5
End Sub
Public Sub build()
    'Dim currentDateTime As String
    'Dim yearStr As String, monthStr As String, dayStr As String
    'Dim hourStr As String, minuteStr As String, secondStr As String
    ' Get current date and time
    'currentDateTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    ' Extract year, month, day, hour, and minute
    'yearStr = Format(Now, "yyyy")
    'monthStr = Format(Now, "mm")
    'dayStr = Format(Now, "dd")
    'hourStr = Format(Now, "hh")
    'minuteStr = Format(Now, "nn")
    'secondStr = Format(Now, "ss")

    ' Combine the components into the desired string format
    'Dim dateString As String
    'dateString = yearStr & "-" & monthStr & "-" & dayStr & " " & hourStr & ":" & minuteStr & ":" & secondStr
    
    If IsInlineShape = True Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    End If
    
    Dim newCanvas As Shape

    Set newCanvas = ActiveDocument.Shapes.AddCanvas( _
        left:=Selection.Information(wdHorizontalPositionRelativeToPage) - Me.width, _
        top:=Selection.Information(wdVerticalPositionRelativeToPage), _
        width:=CentimetersToPoints(Me.width), _
        height:=CentimetersToPoints(Me.height))

    With newCanvas
        .Name = "plane"
        .WrapFormat.Type = wdWrapInline
        .WrapFormat.DistanceTop = CentimetersToPoints(0)
        .WrapFormat.DistanceBottom = CentimetersToPoints(0)
        .WrapFormat.AllowOverlap = False
    End With
    
    Dim newCanvasShapes As CanvasShapes
    Set newCanvasShapes = newCanvas.CanvasItems
    
    TextboxSize = CentimetersToPoints(1)

    'Adds the X axis
     Set Xaxis = newCanvasShapes.AddLine( _
        CentimetersToPoints(Me.AxisEdgeDistanceBeg), _
        CentimetersToPoints(Me.height - Me.Origin), _
        CentimetersToPoints(Me.width - Me.AxisEdgeDistanceEnd), _
        CentimetersToPoints(Me.height - Me.Origin))
    
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
        CentimetersToPoints(Me.Origin), _
        CentimetersToPoints(Me.height - Me.AxisEdgeDistanceBeg), _
        CentimetersToPoints(Me.Origin), _
        CentimetersToPoints(Me.AxisEdgeDistanceEnd))
    
    Yaxis.Name = "yaxis"

     With Yaxis.Line
        .DashStyle = msoLineSolid
        .Weight = 1.5
        .ForeColor.RGB = RGB(0, 0, 0)
        .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadStyle = msoArrowheadNone
     End With
     
    'Creating the box for x
    Set Xbox = newCanvasShapes.AddShape( _
        msoShapeRectangle, _
        CentimetersToPoints(Me.width - Me.AxisEdgeDistanceEnd) - TextboxSize / 2, _
        CentimetersToPoints(Me.height - Me.Origin), _
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
    Dim objRange As Range
    Dim objEq As OMath
    Set objRange = Xbox.TextFrame.TextRange
    objRange.Text = "x"
    Set objRange = Xbox.TextFrame.TextRange.OMaths.Add(objRange)
    'Set objEq = objRange.OMaths(1)
    
    'Creating the box for y
     Set Ybox = newCanvasShapes.AddShape( _
     msoShapeRectangle, _
     CentimetersToPoints(0), _
     CentimetersToPoints(Me.AxisEdgeDistanceEnd) - TextboxSize / 2, _
     TextboxSize, _
     TextboxSize)
     
     With Ybox
        .Name = "ybox"
        .TextFrame.TextRange = ""
        .TextFrame.TextRange.Font.ColorIndex = wdBlack
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
     End With
       'Creating math equation form of y

    Dim objRange2 As Range
    Dim objEq2 As OMath
    Set objRange2 = Ybox.TextFrame.TextRange
    objRange2.Text = "y"
    Set objRange2 = Selection.OMaths.Add(objRange2)
    'Set objEq2 = objRange.OMaths(1)
End Sub






