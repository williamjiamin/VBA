Sub addValuetoEnd()
    Worksheets("Sheet1").Activate
    Range("A1").End(xlDown).Offset(1, 0).Select
    
    ActiveCell.Value = ActiveCell.Offset(-1, 0).Value + 1
    ActiveCell.Offset(0, 1).Value = "护目镜"
    ActiveCell.Offset(0, 2).Value = "200"
    ActiveCell.Offset(0, 1).Value = "国产"
    ActiveCell.Offset(0, 1).Value = "1/1/2022"
    
    
End Sub
