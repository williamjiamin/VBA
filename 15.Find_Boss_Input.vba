Option Explicit

Sub Find_Boss_Input()

        Dim The_Name_Boss_Want_To_Find As String
        Dim The_Name_Cell As Range
        
        The_Name_Boss_Want_To_Find = InputBox("老板呀，请输入您想查找的资产科目名称")
        
        Set The_Name_Cell = _
               Range("B2", Range("B2").End(xlDown)).Find(The_Name_Boss_Want_To_Find)
        
        MsgBox The_Name_Cell.Value & "找到了，在这里：" & The_Name_Cell.Address
        
        
End Sub
