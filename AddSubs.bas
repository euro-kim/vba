Attribute VB_Name = "AddSubs"
Sub Addxyplane()
    Dim new_xyplane As New xyplane
    new_xyplane.initialize
    new_xyplane.build
End Sub
Sub Addxyzspace()
    Dim new_xyzplane As New xyzspace
    new_xyzplane.initialize
    new_xyzplane.build
End Sub
Sub Addxyzspace2()
    Dim new_xyzplane2 As New xyzspace2
    new_xyzplane2.initialize
    new_xyzplane2.build
End Sub
Sub Addmatrix()
    Dim new_matrix As New matrix
    new_matrix.initialize
    new_matrix.build
End Sub
Sub Addequation()
    Dim new_equation As New equation
    new_equation.build
    ActiveWindow.IMEMode = wdIMEModeAlpha
End Sub
Sub GetMostRecentPNGFile()
    Dim folderPath As String
    Dim fileName As String
    Dim mostRecentFile As String
    Dim mostRecentDate As Date
    Dim fileDate As Date
    
    ' Set the folder path
    folderPath = "C:\Users\HOME\OneDrive\»çÁø\pdf\" ' Replace with your folder path
    
    ' Initialize variables
    mostRecentDate = DateSerial(1900, 1, 1)
    
    ' Loop through each file in the folder
    fileName = Dir(folderPath & "*.png")
    Do While fileName <> ""
        ' Get the creation date of the file
        fileDate = FileDateTime(folderPath & fileName)
        
        ' Check if the file is more recent than the current most recent file
        If fileDate > mostRecentDate Then
            mostRecentDate = fileDate
            mostRecentFile = fileName
        End If
        
        ' Get the next file
        fileName = Dir
    Loop
    
 If mostRecentFile <> "" Then
        ' Insert the most recent PNG file as an image in the Word document

        Selection.Range.InlineShapes.AddPicture fileName:=folderPath & mostRecentFile, LinkToFile:=False, SaveWithDocument:=True
        
        ' Move the selection to the end of the document
        Selection.EndKey Unit:=wdStory

    Else
        ' No PNG files found in the folder
        MsgBox "No PNG files found in the folder."
    End If
End Sub


