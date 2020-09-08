Attribute VB_Name = "친욥1"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 브1 브
'

'
Range("B:B").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!B:B")
ActiveWindow.SmallScroll Down:=-9



Range("C:C").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!C:C")
        Range("D:D").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!D:D")
        Range("E:E").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!E:E")
        Range("F:F").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!F:F")
        Range("G:G").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!G:G")
        Range("H:H").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!H:H")
        Range("I:I").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!I:I")
        Range("J:J").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!J:J")







End Sub
