VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Private private_row, private_column As Integer
 Private private_str As String
Public Property Let row(value1 As Integer)
    private_row = value1
End Property
Public Property Let column(value2 As Integer)
    private_column = value2
End Property
Public Property Get row() As Integer
    row = private_row
End Property
Public Property Get column() As Integer
    column = private_column
End Property
Public Property Let str(InputString As String)
    'бс(11&12&13@21&22&23@31&32&33)
    
    If InputString = "" Then
        Dim EmptyStr As String
        EmptyStr = ""
        Dim j As Integer
        For j = 1 To Me.row * Me.column - 1
            EmptyStr = EmptyStr & ","
        Next
        SplitStr = Split(EmptyStr, ",")
    Else
        SplitStr = Split(InputString, ",")
        If Not UBound(SplitStr) - LBound(SplitStr) + 1 = Me.row * Me.column Then
            MsgBox "Number of elements does not match row*colum"
        End If
    End If
    
    Dim temp_str As String
    temp_str = ""
    
    Dim i As Integer
    For i = 0 To UBound(SplitStr)
        If (i + 1) Mod Me.column > 0 Then
            If i > 0 Then
                temp_str = temp_str & SplitStr(i) & "&"
            Else
                temp_str = "бс(" & SplitStr(i) & "&"
            End If
        Else
            If i < UBound(SplitStr) Then
                temp_str = temp_str & SplitStr(i) & "@"
            Else
                temp_str = temp_str & SplitStr(i) & ")"
            End If
        End If
    Next
    
    'бс(@@)
    If Me.column = 1 Then
    temp_str = "бс("
    For i = 1 To Me.row - 1
        temp_str = temp_str & "@"
    Next
    temp_str = temp_str & ")"
    End If
    
    private_str = temp_str
    
End Property

Public Property Get str() As String
    str = private_str
End Property

Public Sub initialize()
    TrimStr = Trim(InputBox("Form:  row*column", "matrix generator", ""))
    If InStr(TrimStr, ",") > 0 Then
        SplitStr = Split(TrimStr, ",")
    ElseIf InStr(TrimStr, "*") > 0 Then
        SplitStr = Split(TrimStr, "*")
    Else
        Exit Sub
    End If
    
    Me.row = CInt(SplitStr(0))
    Me.column = CInt(SplitStr(1))
    Me.str = ""
End Sub
Public Sub build()
    Selection.TypeText (Me.str)
End Sub

