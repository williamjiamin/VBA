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
    ActiveWorkbook.Close
    'ActiveWorkbook.Close True
End Sub

'创建/打开一个全新的workbook

Sub open_a_new_workbook()
    Workbooks.Add
End Sub


'特定路径的workbook打开

Sub open_workbook_by_dir()
    Workbooks.Open "C:\Users\yons\Desktop\aaa_test001.xlsx"
End Sub



'xlsx为excel的默认类型
'https://docs.microsoft.com/zh-cn/office/vba/api/overview/

Sub change_file_type()
    Workbooks.Add
    'ActiveWorkbook.Close
    'mov->m4v
    ActiveWorkbook.SaveAs "C:\Users\yons\Desktop\lexueoude.xlsm", xlOpenXMLWorkbookMacroEnabled
End Sub

