Option Explicit

Sub changeFormatWithoutWith()
        Range("F2:F23").Font.Color = rgbBlue
        Range("F2:F23").Font.Size = 20
End Sub



Sub changeFormatWithWith()
        With Worksheets("Sheet1").Range("F2", Worksheets("Sheet1").Range("F1").End(xlDown))
        .Font.Color = rgbBlue
        .Font.Size = 20
        .Interior.Color = rgbRed
        End With
        
        With Worksheets("Sheet1").Range("E2", Worksheets("Sheet1").Range("E1").End(xlDown))
        .Font.Color = rgbBlue
        .Font.Size = 20
        .Interior.Color = rgbAquamarine
        .NumberFormat = "dddd dd mm yyyy"
        
        Worksheets("Sheet1").Cells.Interior.Color = .Interior.Color
        
        End With
        
        
        
End Sub
