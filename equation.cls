VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "equation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private private_string As String
Private private_base As String
Private private_subscript As String
Private private_exponent As String
Public Property Let str(InputString As String)
    If Len(InputString) = 0 Then
        private_string = ""
        Exit Property
    End If

    GreekStr = InputString
    
    Dim LowerGreekAlphabets() As String, UpperGreekAlphabets() As String
    LowerGreekAlphabets = Split(ChrW(945) & "," & ChrW(946) & "," & ChrW(947) & "," & ChrW(948) & "," & ChrW(949) & "," & ChrW(950) & "," & ChrW(951) & "," & ChrW(952) & "," & ChrW(953) & "," & ChrW(954) & "," & ChrW(955) & "," & ChrW(956) & "," & ChrW(957) & "," & ChrW(958) & "," & ChrW(959) & "," & ChrW(960) & "," & ChrW(961) & "," & ChrW(962) & "," & ChrW(963) & "," & ChrW(964) & "," & ChrW(965) & "," & ChrW(966) & "," & ChrW(967) & "," & ChrW(969), ",")
    UpperGreekAlphabets = Split(ChrW(913) & "," & ChrW(914) & "," & ChrW(915) & "," & ChrW(916) & "," & ChrW(917) & "," & ChrW(918) & "," & ChrW(919) & "," & ChrW(920) & "," & ChrW(921) & "," & ChrW(922) & "," & ChrW(923) & "," & ChrW(924) & "," & ChrW(925) & "," & ChrW(926) & "," & ChrW(927) & "," & ChrW(928) & "," & ChrW(929) & "," & ChrW(931) & "," & ChrW(932) & "," & ChrW(933) & "," & ChrW(934) & "," & ChrW(935) & "," & ChrW(936) & "," & ChrW(937), ",")
    
    Dim LowerEnglishAlphabets() As String, UpperEnglishAlphabets() As String
    LowerEnglishAlphabets = Split("aa,bb,gg,dd,ee,zz,hh,cc,ii,kk,ll,mm,nn,qq,oo,pp,rr,ss,tt,uu,ff,xx,jj,ww", ",")
    UpperEnglishAlphabets = Split("AA,BB,GG,DD,EE,ZZ,HH,CC,II,KK,LL,MM,NN,QQ,OO,PP,RR,SS,TT,UU,FF,XX,JJ,WW", ",")
    
    Dim i As Integer
    For i = 0 To 23
        GreekStr = Replace(GreekStr, LowerEnglishAlphabets(i), LowerGreekAlphabets(i))
        GreekStr = Replace(GreekStr, UpperEnglishAlphabets(i), UpperGreekAlphabets(i))
    Next i
    ChangeStr = GreekStr
    
    private_string = ChangeStr
    
    commaCount = Len(ChangeStr) - Len(Replace(ChangeStr, ",", ""))
    SplitStr = Split(ChangeStr, ",")
    
    If commaCount = 2 Then
        private_exponent = "^" & SplitStr(2)
        private_subscript = "_" & SplitStr(1)
    ElseIf commaCount = 1 Then
        private_exponent = ""
        private_subscript = "_" & SplitStr(1)
    Else
        private_exponent = ""
        private_subscript = ""
    End If
    
    Bar = ChrW(&HAF)
    hat = ChrW(&H302)
    vec = ChrW(&H20D7)
    under = ChrW(&H252C)
    
    If InStr(SplitStr(0), "bar") > 0 Then
        BaseStr = Bar & Replace(SplitStr(0), "bar", "")
    ElseIf InStr(SplitStr(0), "hat") > 0 Then
        BaseStr = Replace(SplitStr(0), "hat", "") & hat
    ElseIf InStr(SplitStr(0), "vec") > 0 Then
        BaseStr = Replace(SplitStr(0), "vec", "") & vec
    ElseIf InStr(SplitStr(0), "mat") > 0 Then
        TrimStr = Trim(InputBox("Form:  row*column", "matrix generator", ""))
        If InStr(TrimStr, ",") > 0 Then
            temprivate_string = Split(TrimStr, ",")
        ElseIf InStr(TrimStr, "*") > 0 Then
            temprivate_string = Split(TrimStr, "*")
        Else
            MsgBox "No Matirx Input"
            BaseStr = SplitStr(0)
        End If
        BaseStr = Replace(SplitStr(0), "mat", "") & under & "(" & temprivate_string(0) & "*" & temprivate_string(1) & ")"
    Else
        BaseStr = SplitStr(0)
    End If
    
    private_base = BaseStr
End Property
Public Property Get str() As String
    str = private_string
End Property
Public Property Get base() As String
    base = private_base
End Property
Public Property Get subscript() As String
    subscript = private_subscript
End Property
Public Property Get exponent() As String
    exponent = private_exponent
End Property
Public Sub build()
    
    'creating the math form
    Dim buildStr As String
    If Me.str = "" Then
        buildStr = ""
    Else
        If Me.exponent = "" And Me.subscript = "" Then
            If left(str, 1) = "(" And Right(str, 1) = ")" Then
                buildStr = Mid(Me.base, 2, Len(Me.base) - 2)
            End If
            buildStr = Me.base
        Else
            If Me.exponent = "" Then
                    buildStr = Me.base & Me.subscript
            Else
                    buildStr = Me.base & Me.subscript & Me.exponent
            End If
        End If
    End If
    
    'building the math forn
    If IsInlineShape = True Then
        'creating a box
        Dim new_canvas As New xyplane
        Dim EqBox As Shape
        Set EqBox = Selection.ShapeRange(1).CanvasItems.AddShape(msoShapeRectangle, new_canvas.xcor(3.5), new_canvas.ycor(3), CentimetersToPoints(0.5), CentimetersToPoints(0.5))
        With EqBox
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            .TextFrame.TextRange = ""
            .TextFrame.TextRange.Font.ColorIndex = wdBlack
            .TextFrame.MarginTop = 0
            .TextFrame.MarginBottom = 0
            .TextFrame.MarginLeft = 0
            .TextFrame.MarginRight = 0
            .TextFrame.AutoSize = True
            .TextFrame.TextRange.Font.Name = "Cambria Math"
        End With
               
        Dim objRange As Range
        Dim objEq As OMath
        Set objRange = EqBox.TextFrame.TextRange
        objRange.Text = buildStr
        Set objRange = EqBox.TextFrame.TextRange.OMaths.Add(objRange)
    
    Else
        Set objRange = Selection.Range
        objRange.Text = buildStr
        Set objRange = Selection.OMaths.Add(objRange)
        Set objEq = objRange.OMaths(1)
        objEq.BuildUp
    End If


End Sub
Public Sub initialize()
    TrimStr = Trim(InputBox("Form:  base Type, subscript(optional), exponent(optional)", "math form generator", ""))
    Me.str = TrimStr
End Sub

