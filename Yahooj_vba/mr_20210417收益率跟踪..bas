Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("I1").Select
ActiveCell.FormulaR1C1 = "=RC[-6]"
Range("I" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-6]/R2C3-1"
 Range("J1").Select
ActiveCell.FormulaR1C1 = "=RC[-6]"
Range("J" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-6]/R2C4-1"
 Range("K1").Select
ActiveCell.FormulaR1C1 = "=RC[-6]"
Range("K" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-6]/R2C5-1"
 Range("L1").Select
ActiveCell.FormulaR1C1 = "=RC[-6]"
Range("L" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-6]/R2C6-1"
Next
Range("I1:L" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$I$1:$L$" & num)
End Sub
