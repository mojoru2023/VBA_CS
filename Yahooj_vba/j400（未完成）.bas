Attribute VB_Name = "模块1"
'设置一个全局变量'
Public num As Integer

Sub 宏4()
Attribute 宏4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏4 宏
'

'

'先选定起始单元格'




Dim i As Integer

For i = 2 To 1900
    If Range("a" & i) = "" Then
        num = i
        
        Exit For
    End If
    
    

    
    Range("AA1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("AA" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C3-1"



Range("AB1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("AB" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C4-1"




Range("Ac" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C5-1"
Range("Ac1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ad" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C6-1"
Range("Ad1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ae" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C7-1"
Range("Ae1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Af" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C8-1"
Range("Af1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ag" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C9-1"
Range("Ag1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ah" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C10-1"
Range("Ah1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ai" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C11-1"
Range("Ai1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Aj" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C12-1"
Range("Aj1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ak" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C13-1"
Range("Ak1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Al" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C14-1"
Range("Al1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Am" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C15-1"
Range("Am1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("An" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C16-1"
Range("An1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ao" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C17-1"
Range("Ao1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ap" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C18-1"
Range("Ap1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Aq" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C19-1"
Range("Aq1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Ar" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C20-1"
Range("Ar1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("As" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C21-1"
Range("As1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("At" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C22-1"
Range("At1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Au" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C23-1"
Range("Au1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("Av" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C24-1"
Range("Av1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
    
    
    
    
Next


'作图的动作'
    ActiveCell.Offset(-28, -28).Range("A1:V" & num).Select
    ActiveCell.Offset(-14, -24).Range("A1").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$AA$1:$AV$" & num)

    
End Sub

