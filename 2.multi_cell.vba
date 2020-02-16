Sub SelectMultipleCells()
    ' select cell first and change color
    
    Range("A1:D1").Select
    Selection.Interior.Color = rgbDarkBlue
    
    'change color without selecting cell first
    Range("A1:D1").Interior.Color = rgbWhite
    
    'change font color without selecting cell first
    Range("A1:D1").Font.Color = rgbRed
    
End Sub

