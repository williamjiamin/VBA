Option Explicit

Sub add_something_new_using_variable()
        
        Dim new_balance_sheet_maker As String
        Dim new_balance_sheet_date As Date
        Dim bs_id As Integer
        '1.直接通过代码中的信息对变量进行输入
        
        'new_balance_sheet_maker = "制表人：lexueoude.com(公众号：乐学Fintech)"
        
        '2.通过inputbox进行交互式的输入
        new_balance_sheet_maker = InputBox("领导啊！请在这里输入制表人信息就可以啦！-->")
        new_balance_sheet_date = InputBox("领导啊！请在这里输入制表日期就可以啦！-->")
        
        Worksheets("新资产").Activate
        Range("B2").End(xlDown).Offset(1, 0).Select
        '3.通过各类操作获取原来excel中任意区域的信息，并做相关计算后填入
        bs_id = ActiveCell.Offset(-1, -1).Value
        bs_id = bs_id + 1
        '先Dim声明之后，按下Ctrl+空格，调出自动填充(如果调不出来，可以调整输入法为英文后再试)
        ActiveCell.Offset(0, -1).Value = bs_id
        ActiveCell.Value = new_balance_sheet_maker
        ActiveCell.Offset(0, 1).Value = new_balance_sheet_date
        
        MsgBox new_balance_sheet_maker & "已经添加到资产负债表中了哦~"
End Sub

