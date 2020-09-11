Attribute VB_Name = "模块11"
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
    
    
Range("S1").Select


    '第一列的净值计算'
    
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    Range("S" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C3-1"
   
    
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("T" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C4-1"
    
    Range("u1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("u" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C5-1"
    
    
    
    
    
        Range("v1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("v" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C6-1"
    
    
    
        
        Range("w1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("w" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C7-1"
    
    
            Range("x1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("x" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C8-1"
    
    
                Range("y1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("y" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C9-1"
    
    
    
               Range("z1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("z" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C10-1"
    
    
                   Range("aa1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("aa" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C11-1"
    
    
                       Range("ab1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("ab" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C12-1"
    
        
                       Range("ac1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("ac" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C13-1"
    
    
    
        
                       Range("ad1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("ad" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C14-1"
    
    
        
                       Range("ae1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("ae" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C15-1"
    
    
        
                       Range("af1").Select
    ActiveCell.FormulaR1C1 = "=RC[-16]"
    
    Range("af" & i).Select
    ActiveCell.FormulaR1C1 = "=RC[-16]/R2C16-1"
    
    
  
    
    
    
    
    
    
Next


'作图的动作'
    ActiveCell.Offset(-3, -3).Range("A1:B" & num).Select
    ActiveCell.Offset(0, -3).Range("A1").Activate
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$S$1:$af$" & num)

    
End Sub

