Sub ReletivePosi()
        Worksheets("Sheet1").Activate
        Range("A1").Select
        ActiveCell.End(xlDown).Select
        'ActiveCell.End(xlToRight).Select
        ActiveCell.Offset(1, 0).Select
        
End Sub

Sub ReletivePosiQuick()
         Worksheets("Sheet1").Activate
         Range("A1").End(xlDown).Offset(1, 0).Select
         
End Sub
