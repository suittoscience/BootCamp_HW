Sub EasyWallStreet()

'Set variables
Dim x As Double
Dim Total As Double
Dim TotalV As Double


'Headings
    Cells(1, 9).Value = Cells(1, 1).Value
    Cells(1, 10).Value = "Total Stock Value"
    
'Create loop
    x = 2
    Cells(x, 9).Value = Cells(x, 1).Value

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To LastRow

If Cells(i, 1).Value = Cells(x, 9).Value Then

TotalV = TotalV + Cells(i, 7).Value

     Else
     
Cells(x, 10).Value = TotalV

TotalV = Cells(i, 7).Value

x = x + 1
Cells(x, 9).Value = Cells(i, 1).Value

End If
    
    Next i

Cells(x, 10).Value = TotalV
    
'Resize the columns I through J

Columns("I:J").EntireColumn.AutoFit

Cells(1, 1).Select

End Sub

