Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("H1").Select
ActiveCell.FormulaR1C1 = "=RC[-5]"
Range("H" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-5]/R2C3-1"
 Range("I1").Select
ActiveCell.FormulaR1C1 = "=RC[-5]"
Range("I" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-5]/R2C4-1"
 Range("J1").Select
ActiveCell.FormulaR1C1 = "=RC[-5]"
Range("J" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-5]/R2C5-1"
Next
Range("H1:J" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$H$1:$J$" & num)
End Sub
