Attribute VB_Name = "Ä£¿é1"
Public num As Integer


Sub ºê2()
'
' ºê2 ºê
'

'

For i = 2 To 19000
    If Range("a" & i) = "" Then
        Exit For
    End If
    


Range("V1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("V" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C3-1"


Range("W1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("W" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C4-1"


Range("X1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("X" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C5-1"



Range("Y1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("Y" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C6-1"



Range("Z1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("Z" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C7-1"


Range("AA1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AA" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C8-1"




Range("AB1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AB" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C9-1"




Range("AC1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AC" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C10-1"





Range("AD1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AD" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C11-1"


Range("AE1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AE" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C12-1"

Range("AF1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AF" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C13-1"

Range("AG1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AG" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C14-1"
Range("AH1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AH" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C15-1"

Range("AI1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AI" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C16-1"

Range("AJ1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AJ" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C17-1"
Range("AK1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AK" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C18-1"

Range("AL1").Select
ActiveCell.FormulaR1C1 = "=RC[-19]"
Range("AL" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-19]/R2C19-1"




Next

        
    ActiveCell.Offset(-65, -28).Range("A1:V" & i).Select
    ActiveCell.Offset(0, -3).Range("A1").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet2!$V$1:$AL$" & i)

    
End Sub




