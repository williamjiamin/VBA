Option Explicit

Sub AddDetails()
        Dim Name As String
        Dim strMakeDate As String
        Dim datMakeDate As Date
        Dim Num As Integer
        
        Name = InputBox("请输入制表人信息：")
        strMakeDate = InputBox("请输入报表制作日期：")
        
        If strMakeDate = "" Then
            MsgBox "您没有输入日期"
            Exit Sub
        End If
        
        datMakeDate = CDate(strMakeDate)
        
        Num = InputBox("请输入报表编号：")
        
        Range("B1").End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = Num
        ActiveCell.Offset(0, 1).Value = Name
        ActiveCell.Offset(0, 2).Value = datMakeDate
        
End Sub
