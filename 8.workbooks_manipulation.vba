'referring to workbook
Sub referring_to_workbook()
    Workbooks("乐学Fintech数据汇报工作簿2.xlsx").Activate
    Workbooks("乐学偶得数据统计工作簿1.xlsx").Activate
End Sub


'referring to workbook by index(根据你打开的顺序进行index)

Sub referring_to_workbook_by_index()
    Workbooks(1).Activate
    Workbooks(2).Activate
    
End Sub


'逐一关闭正在进行的workbook
Sub close_activated_workbook()
    ActivateWorkbook.Close
    ActivateWorkbook.Close True
End Sub
