Sub g()
 Range("B:B").Select  
  ActiveSheet.Shapes.AddChart2(227, xlLine).Select 
   ActiveChart.SetSourceData Source:=Range("Sheet1!B:B")
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
   

End Sub
