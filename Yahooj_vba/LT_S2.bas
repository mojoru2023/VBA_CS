Public num As Integer
Sub s()
Dim i As Integer
For i = 2 To 9999
If Range("C" & i) = "" Then
num = i
Exit For
End If
 Range("S1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("S" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C3-1"
 Range("T1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("T" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C4-1"
 Range("U1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("U" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C5-1"
 Range("V1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("V" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C6-1"
 Range("W1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("W" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C7-1"
 Range("X1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("X" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C8-1"
 Range("Y1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("Y" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C9-1"
 Range("Z1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("Z" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C10-1"
 Range("AA1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AA" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C11-1"
 Range("AB1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AB" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C12-1"
 Range("AC1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AC" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C13-1"
 Range("AD1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AD" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C14-1"
 Range("AE1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AE" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C15-1"
 Range("AF1").Select
ActiveCell.FormulaR1C1 = "=RC[-16]"
Range("AF" & i).Select 
 ActiveCell.FormulaR1C1 = "=RC[-16]/R2C16-1"
Next
Range("S1:AF" & num).Select
ActiveSheet.Shapes.AddChart2(227, xlLine).Select
ActiveChart.SetSourceData Source:=Range("Sheet1!$S$1:$AF$" & num)
End Sub
