Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("O1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("O" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C3-1"
 Range("P1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("P" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C4-1"
 Range("Q1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("Q" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C5-1"
 Range("R1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("R" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C6-1"
 Range("S1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("S" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C7-1"
 Range("T1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("T" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C8-1"
 Range("U1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("U" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C9-1"
 Range("V1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("V" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C10-1"
 Range("W1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("W" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C11-1"
 Range("X1").Select
ActiveCell.FormulaR1C1 = "=RC[-12]"
Range("X" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-12]/R2C12-1"
Next
Range("O1:X" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$O$1:$X$" & num)
End Sub
