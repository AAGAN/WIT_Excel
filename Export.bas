Attribute VB_Name = "Module5"
Sub Export()
 Dim flag As Boolean
Dim i As Integer
Dim strPath As String
strPath = Application.GetSaveAsFilename(FileFilter:= _
 "Text Files (*.txt), *.txt", Title:="Save Location")
 If strPath <> "False" Then
    'open the file for writing
    Open strPath For Output As #4
     
    flag = True
    i = 1
    
    'keeps going until the end of the file is reacheed
    While flag = True
        'check if the current cell has data in it
        If Cells(i, 1) <> "" Then
        'write the data to the file
           Print #4, Cells(i, 1).Value; vbTab; Cells(i, 2).Value
            'go to next cell
           i = i + 1
        Else
           'if the last row has been reached exit the loop
            flag = False
        End If
    Wend
    'close the file
    Close #4
End If
End Sub

