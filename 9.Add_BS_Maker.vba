Sub add_something_new()
        Worksheets("新资产").Activate
        Range("B2").End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = "制表人：乐学偶得(公众号：乐学Fintech)"
        
        MsgBox "已经加好了制表人哦~"
End Sub
