Option Explicit

Sub GetYourName()
        Range("A1").Value = InputBox("请输入您的名字~", "必填选项")
End Sub

Sub GetYourNameUsingVariable()
        Dim YourName As String
        YourName = InputBox("请输入您的名字~", "必填选项")
        MsgBox "你好呀:" & YourName
End Sub


Sub GetYourNameUsingVariableV2()
        Dim YourName As String
        YourName = InputBox("请输入您的名字~", "必填选项")
        
        If YourName = "" Then
            MsgBox "您似乎并没有输入任何名字鸭~", vbExclamation
        
        Else
            MsgBox "你好呀:" & YourName
        End If
            
End Sub
