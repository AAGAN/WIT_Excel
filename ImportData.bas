Attribute VB_Name = "Module2"
Sub ImportData()
'This has to open the new file
'If a file is selected then the drop down box choice needs to be selected
'This file creates the grid
'Referencing the cell in sheet 1, the new file is queries until the bucket is matched
'This density is placed in the corresponding cell

Dim FilePath As String
Dim intChoice As Integer
Dim ResultStr As String
Dim CountLines As Integer
Dim LineFromFile As String


Dim magicNumber As Integer 'Number from combobox
Dim startingRow As Integer
Dim startingColumn As Integer
    
Dim tmpNumber As Integer
Dim tmpRow As Integer
Dim tmpColumn As Integer
Dim densityIndex As Integer
    
Dim numberCounter As Integer
   
Dim comboInput As String
    
Dim direction As Byte
Dim worksheetName As String
worksheetName = "Results " & ThisWorkbook.Sheets.Count - 2

 Dim directory As String
 
 'Set the directory path
 directory = "C:\"


'Options for the import dialog
Application.FileDialog(msoFileDialogOpen).InitialFileName = directory
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
Application.FileDialog(msoFileDialogOpen).Filters.Clear
Application.FileDialog(msoFileDialogOpen).Filters.Add "Files", "*.csv;*.txt"

intChoice = Application.FileDialog(msoFileDialogOpen).Show


If intChoice <> 0 Then ' The file has been selected and will be opened
    Worksheets("Results").Unprotect password:="vette"
    'Insert a new worksheet
    CreateSheet (worksheetName)
    
    copyInputData (worksheetName)
    
    With Worksheets("Input").Shapes("Drop Down 7").ControlFormat
        
        comboInput = .List(.ListIndex)
     
    End With
    
    'Set initial Data based on the ComboBox
    If comboInput = "8x8" Then
        magicNumber = 8
        startingRow = 19
        startingColumn = 10
    ElseIf comboInput = "9x9" Then
        magicNumber = 9
        startingRow = 20
        startingColumn = 11
    ElseIf comboInput = "10x10" Then
        magicNumber = 10
        startingRow = 21
        startingColumn = 12
    End If
     Worksheets(worksheetName).Range("C12:L21").Clear
    
    Call ConditionalFormatting(startingRow, startingColumn, worksheetName)
    
    numberCounter = 1
    tmpRow = startingRow
    tmpColumn = startingColumn
    tmpNumber = Worksheets("Input").Cells(tmpRow, tmpColumn).Value
    direction = 0
    
   FilePath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
  
   
   'Importing Data
   'LineItems (0) 'Bucket **
   'LineItems (1) 'Density **
   'LineItems (2) 'Date
   'LineItems (3) 'Time
   'LineItems (4) 'Depth
   'LineItems (5) 'Weight
   'LineItems (6) 'Temp
    
    Open FilePath For Input As #1
    Do While Seek(1) <= LOF(1)
    Line Input #1, ResultStr
    CountLines = CountLines + 1
    Loop
    Close (1)
    
    CountLines = CountLines - 6
   While numberCounter < magicNumber * magicNumber + 1
        
        Open FilePath For Input As #1
        row_number = -6
        
               
         Do While (CountLines > row_number)
        
              Line Input #1, LineFromFile
             
              LineItems = Split(LineFromFile, vbTab)
              
               
              If tmpNumber = LineItems(0) Then 'The bucket number from the first sheet = the bucket number in thefile
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn) = LineItems(1)
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).Borders.LineStyle = xlContinuous
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).Borders.Weight = xlMedium
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).HorizontalAlignment = xlCenter
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).VerticalAlignment = xlCenter
                numberCounter = numberCounter + 1
                 
                 
                 
                If direction = 0 Then 'Filling numbers to the left
                    tmpColumn = tmpColumn - 1
                    If tmpColumn < 3 Then
                        direction = 1
                        tmpColumn = 3
                        tmpRow = tmpRow - 1
                    End If
                    tmpNumber = Worksheets("Input").Cells(tmpRow, tmpColumn).Value
                ElseIf direction = 1 Then 'Filling numbers to the right
                    tmpColumn = tmpColumn + 1
                    If tmpColumn > 2 + magicNumber Then
                        direction = 0
                        tmpColumn = 2 + magicNumber
                        tmpRow = tmpRow - 1
                    End If
                    tmpNumber = Worksheets("Input").Cells(tmpRow, tmpColumn).Value
                End If
                Close #1
                Exit Do
                
                
              Else
              
              
                row_number = row_number + 1
                 If row_number = CountLines Then
                    Close #1
                    
                
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).Borders.LineStyle = xlContinuous
                Worksheets(worksheetName).Cells(tmpRow, tmpColumn).Borders.Weight = xlMedium
                     
                  numberCounter = numberCounter + 1
                  If direction = 0 Then 'Filling numbers to the left
                    tmpColumn = tmpColumn - 1
                    If tmpColumn < 3 Then
                        direction = 1
                        tmpColumn = 3
                        tmpRow = tmpRow - 1
                    End If
                    tmpNumber = Worksheets("Input").Cells(tmpRow, tmpColumn).Value
                ElseIf direction = 1 Then 'Filling numbers to the right
                    tmpColumn = tmpColumn + 1
                    If tmpColumn > 2 + magicNumber Then
                        direction = 0
                        tmpColumn = 2 + magicNumber
                        tmpRow = tmpRow - 1
                    End If
                    tmpNumber = Worksheets("Input").Cells(tmpRow, tmpColumn).Value
                End If
                     
                    Exit Do
                End If
              End If
         Loop
         
        
                
   Wend
    
  LockWorkBook (worksheetName)
    
End If
End Sub


Sub CreateSheet(s As String)

Sheets("Results").Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
ActiveSheet.Name = s


End Sub

Sub ConditionalFormatting(r As Integer, c As Integer, s As String)

    Sheets(s).Select
    Range(Cells(12, 3), Cells(r, c)).Select
    
  
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0", Formula2:="=0.0149"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(192, 0, 0)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
       Formula1:="=0.015", Formula2:="=0.0199"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 102)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.020", Formula2:="=0.0249"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(146, 208, 80)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
       Formula1:="=0.025", Formula2:="=0.029"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 176, 80)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.03", Formula2:="=0.049"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 112, 192)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.049"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(112, 48, 160)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
         Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=LEN(TRIM(C12))=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 255)
        .TintAndShade = 0
    End With
        Selection.FormatConditions(1).StopIfTrue = False
  
    

    
End Sub


Sub copyInputData(s As String)
's is the new workbook

Dim sprinkler As String
Dim coverage As String
Dim flow As String
Dim recess As String
Dim duration As String
Dim dte As String
Dim parPer As Integer
Dim note As String

'Get the information from the Input worksheet
sprinkler = Worksheets("Input").Cells(3, "E").Value
flow = Worksheets("Input").Cells(5, "E").Value
recess = Worksheets("Input").Cells(6, "E").Value
duration = Worksheets("Input").Cells(7, "E").Value
dte = Worksheets("Input").Cells(3, "K").Value
note = Worksheets("Input").Cells(23, "E").Value
parPer = Worksheets("Input").Cells(11, "D").Value

'Get the text from the combobox
With Worksheets("Input").Shapes("Drop Down 7").ControlFormat
        coverage = .List(.ListIndex)
End With

'Get the checkbox value


'Write the information to the new worksheet
Worksheets(s).Cells(3, "E") = sprinkler
Worksheets(s).Cells(4, "E") = coverage
Worksheets(s).Cells(5, "E") = flow
Worksheets(s).Cells(6, "E") = recess
Worksheets(s).Cells(7, "E") = duration
Worksheets(s).Cells(3, "K") = dte
Worksheets(s).Cells(23, "E") = note

Sheets(s).Select
If parPer = 1 Then 'Parallel is selected
    Cells(9, "C").Select
    With Selection.Interior
        .Color = RGB(146, 208, 80)
    
    End With
ElseIf parPer = 2 Then
    Cells(9, "E").Select
    With Selection.Interior
        .Color = RGB(146, 208, 80)
     End With
    
End If
    





End Sub


Sub LockWorkBook(s As String)

Dim password As String
password = "vette"

Worksheets(s).Protect password:=password
Worksheets("Results").Protect password:=password

End Sub

