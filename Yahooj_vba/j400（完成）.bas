Attribute VB_Name = "ģ��1"

Public num As Integer
Sub D()
Attribute D.VB_ProcData.VB_Invoke_Func = "q\n14"
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



Range("AC1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AC" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C5-1"

Range("AD1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AD" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C6-1"

Range("AE1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AE" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C7-1"

Range("AF1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AF" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C8-1"

Range("AG1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AG" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C9-1"

Range("AH1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AH" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C10-1"

Range("AI1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AI" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C11-1"

Range("AJ1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AJ" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C12-1"

Range("AK1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AK" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C13-1"

Range("AL1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AL" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C14-1"

Range("AM1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AM" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C15-1"

Range("AN1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AN" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C16-1"

Range("AO1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AO" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C17-1"

Range("AP1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AP" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C18-1"

Range("AQ1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AQ" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C19-1"

Range("AR1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AR" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C20-1"

Range("AS1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AS" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C21-1"

Range("AT1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"

Range("AT" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C22-1"

Range("AU1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AU" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C23-1"

Range("AV1").Select
ActiveCell.FormulaR1C1 = "=RC[-24]"
Range("AV" & i).Select
ActiveCell.FormulaR1C1 = "=RC[-24]/R2C24-1"

    
    
    
    
Next


    ActiveCell.Offset(-65, -28).Range("A1:V" & i).Select
    ActiveCell.Offset(-7, -19).Range("A1").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$AA$1:$AV$" & i)
End Sub
