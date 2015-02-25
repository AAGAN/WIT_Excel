Attribute VB_Name = "Module3"
Sub ConvertSnakeToAccelTest()
Application.ScreenUpdating = False 'Stops screen from jumping around while searching - will leave loading icon while running
Dim CurrentRow, CurrentColumn, RowSize, ColumnSize, ColumnOffset, i, j, k As Integer
Dim CurrentWorksheet, TargetWorksheet, CurrentValue As String

RowSize = Range("l6").Value                       'quantity of rows of pans
ColumnSize = Range("p6").Value                   'quantity of columns of pans

CurrentWorksheet = ActiveSheet.Name              'Defines which worksheet you are on
TargetWorksheet = "Export Array"
Sheets("Export Array").Cells.ClearContents
Worksheets(TargetWorksheet).Activate             'go to export worksheet
ActiveSheet.Cells(1, 1).Select                   'select target cell
ActiveCell.Value = "Project Number="
ActiveCell.Offset(1, 0).Value = "Project Name="
ActiveCell.Offset(2, 0).Value = "Test Number="
ActiveCell.Offset(3, 0).Value = "Test Description="
ActiveCell.Offset(4, 0).Value = "Date/Time="
ActiveCell.Offset(5, 0).Value = "Bucket #"
ActiveCell.Offset(5, 1).Value = " Density(gpm/ft^2)"
ActiveCell.Offset(6, 0).Select
Worksheets(CurrentWorksheet).Activate            'come back to current worksheet

k = 0                                            'row count in target worksheet
ColumnOffset = 1

For i = RowSize To 1 Step -1                                            'loop by rows
  ColumnOffset = -ColumnOffset                                          'every next row has opposite direction
  CurrentRow = ActiveCell.Row                                           'row# of current active cell
  For j = ColumnSize To 1 Step -1                                       'loop by columns
    CurrentColumn = ActiveCell.Column                                   'column# of current active cell
    CurrentValue = ActiveCell.Value                                     'take the value of current cell
    Worksheets(TargetWorksheet).Activate                                'go to export worksheet
    k = k + 1                                                           'next row in target worksheet
    If k > (RowSize * ColumnSize) Then
       ActiveWorkbook.Save                                              'safety just in case
       Application.Quit
    End If
    ActiveSheet.Cells((k + 6), 2).Select                                'select target cell
    ActiveCell.Value = CurrentValue                                     'assign value to target cell
    ActiveCell.Offset(0, -1).Value = k
    Worksheets(CurrentWorksheet).Activate                               'come back to current worksheet
    If j > 1 Then ActiveCell.Offset(0, ColumnOffset).Select             'select next current cell on current worksheet
  Next j
ActiveSheet.Cells((CurrentRow - 1), CurrentColumn).Select               'select next row cell up
Next i
End Sub

