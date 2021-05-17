Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("K1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("K" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C3-1"
 Range("L1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("L" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C4-1"
 Range("M1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("M" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C5-1"
 Range("N1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("N" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C6-1"
 Range("O1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("O" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C7-1"
 Range("P1").Select
ActiveCell.FormulaR1C1 = "=RC[-8]"
Range("P" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-8]/R2C8-1"
Next
Range("K1:P" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$K$1:$P$" & num)
End Sub
