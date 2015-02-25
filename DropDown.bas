Attribute VB_Name = "Module1"
Sub DropDown7_Change()

    Dim magicNumber As Integer
    Dim startingRow As Integer
    Dim startingColumn As Integer
    
    Dim tmpNumber As Integer
    Dim tmpRow As Integer
    Dim tmpColumn As Integer
    
    Dim numberCounter As Integer
    
    Dim comboInput As String
    
    Dim direction As Byte
    
    'Get the ComboBox Value
    With Worksheets("Input").Shapes("Drop Down 7").ControlFormat
        
       comboInput = .List(.ListIndex)
     
    End With
    
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
    
    'First clear the cells values
    Worksheets("Input").Range("C12:L21").Clear
    
    'Set the Cells Values
    numberCounter = 1
    tmpRow = startingRow
    tmpColumn = startingColumn
    tmpNumber = 1
    direction = 0
    
    'Loop through and make the cell values
    '1 Will be at the magicNumber*magicNumberCell
    
    'Direction will be used to control the snake order
    '0 means left and 1 means right
    
    'Fill Snake order
     While numberCounter < magicNumber * magicNumber + 1
    
        Worksheets("Input").Cells(tmpRow, tmpColumn) = tmpNumber
        Worksheets("Input").Cells(tmpRow, tmpColumn).Borders.LineStyle = xlContinuous
        Worksheets("Input").Cells(tmpRow, tmpColumn).Borders.Weight = xlMedium
        Worksheets("Input").Cells(tmpRow, tmpColumn).HorizontalAlignment = xlCenter
        Worksheets("Input").Cells(tmpRow, tmpColumn).VerticalAlignment = xlCenter
        
        
        numberCounter = numberCounter + 1
        tmpNumber = tmpNumber + 1
        
        If direction = 0 Then 'Filling numbers to the left
            tmpColumn = tmpColumn - 1
            If tmpColumn < 3 Then
                direction = 1
                tmpColumn = 3
                tmpRow = tmpRow - 1
            End If
        ElseIf direction = 1 Then 'Filling numbers to the right
         tmpColumn = tmpColumn + 1
            If tmpColumn > 2 + magicNumber Then
                direction = 0
                tmpColumn = 2 + magicNumber
                tmpRow = tmpRow - 1
            End If
        End If
        
    Wend
     

End Sub

