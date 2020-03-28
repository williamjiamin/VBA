Sub addValuetoEnd()
    Worksheets("Sheet1").Activate
    Range("A1").End(xlDown).Offset(1, 0).Select
    
    ActiveCell.Value = ActiveCell.Offset(-1, 0).Value + 1
    ActiveCell.Offset(0, 1).Value = "护目镜"
    ActiveCell.Offset(0, 2).Value = "200"
    ActiveCell.Offset(0, 3).Value = "国产"
    ActiveCell.Offset(0, 4).Value = "1/1/2022"
    
    
End Sub


Sub ColSelection()
    Range("A3", Range("A1").End(xlDown)).Select
    Selection.Font.Bold = True
End Sub

Sub ColSelection_v2()
    Range("A3", Range("A1").End(xlDown)).Font.Bold = True
End Sub

Sub ColSelection_v3()
    Range("B3", Range("B1").End(xlDown)).Font.Bold = True
End Sub

'注意这种跳转有问题

Sub ColSelection_v4()
    Range("C3", Range("C2").End(xlDown)).Font.Bold = True
End Sub

Sub ColSelection_v5()
    Range("A3", Range("A3").End(xlDown).End(xlToRight)).Select
    Selection.Interior.Color = rgbAliceBlue
    
End Sub

Sub ColSelection_v6()
    Range("A3", Range("A3").End(xlDown).End(xlToRight)).Font.Bold = True
End Sub


Sub CopyEntireArea()
    Worksheets("Sheet1").Activate
    Range("A1").CurrentRegion.Copy
    
    'Worksheets("Sheet2").Activate
    'Range("A1").PasteSpecial
     Worksheets("Sheet2").Range("A1").PasteSpecial
     Worksheets("Sheet2").Range("A1").PasteSpecial xlPasteColumnWidths
    
End Sub
