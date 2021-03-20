Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("N1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("N" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C3-1"
 Range("O1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("O" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C4-1"
 Range("P1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("P" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C5-1"
 Range("Q1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("Q" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C6-1"
 Range("R1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("R" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C7-1"
 Range("S1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("S" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C8-1"
 Range("T1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("T" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C9-1"
 Range("U1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("U" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C10-1"
 Range("V1").Select
ActiveCell.FormulaR1C1 = "=RC[-11]"
Range("V" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-11]/R2C11-1"
Next
Range("N1:V" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$N$1:$V$" & num)
End Sub
