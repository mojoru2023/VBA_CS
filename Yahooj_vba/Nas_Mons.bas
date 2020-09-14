Attribute VB_Name = "模块1"
'设置一个全局变量'
Public num As Integer

Sub 宏2()
'
' 宏2 宏
'

'


Dim i As Integer

For i = 2 To 1900
    If Range("a" & i) = "" Then
        num = i
        
        Exit For
    End If
    



Range("T1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("T" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C3-1"



Range("U1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("U" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C4-1"
    
    
Range("V1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("V" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C5-1"

Range("W1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("W" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C6-1"
    
Range("X1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("X" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C7-1"
    
Range("Y1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("Y" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C8-1"
    
Range("Z1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("Z" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C9-1"

Range("AA1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AA" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C10-1"

Range("AB1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AB" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C11-1"

Range("AC1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AC" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C12-1"

Range("AD1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AD" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C13-1"

Range("AE1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AE" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C14-1"

Range("AF1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AF" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C15-1"

Range("AG1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AG" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C16-1"


Range("AH1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AH" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C17-1"

Range("AI1").Select
ActiveCell.FormulaR1C1 = "=RC[-17]"
Range("AI" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-17]/R2C18-1"



Next
    Range("T1:AI" & num).Select
    Range("U5").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$T$1:$AI$" & num)

    
    
    
End Sub

