Option Explicit

Public new_balance_sheet_maker As String
Public new_balance_sheet_date As Date
Public bs_id As Integer

Sub GetBossInput()
        new_balance_sheet_maker = InputBox("领导啊！请在这里输入制表人信息就可以啦！-->")
        new_balance_sheet_date = InputBox("领导啊！请在这里输入制表日期就可以啦！-->")
        
        Call WriteToList
        
End Sub


Sub WriteToList()

        Worksheets("新资产").Activate
        Range("B2").End(xlDown).Offset(1, 0).Select
        bs_id = ActiveCell.Offset(-1, -1).Value
        bs_id = bs_id + 1
      
        ActiveCell.Offset(0, -1).Value = bs_id
        ActiveCell.Value = new_balance_sheet_maker
        ActiveCell.Offset(0, 1).Value = new_balance_sheet_date
        
        MsgBox new_balance_sheet_maker & "已经添加到资产负债表中了哦~"

End Sub


