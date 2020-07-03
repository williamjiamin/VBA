Option Explicit

Sub Change_Range_Of_Cell()
        Dim NamesCells As Range
        
        Set NamesCells = Range("B2:B7")
        'Set NamesCells = Range("C2", Range("C2").End(xlDown))
        
        NamesCells.Font.Color = rgbBlue
        
End Sub
