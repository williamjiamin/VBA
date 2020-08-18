Option Explicit

Sub ApplicationInputBox()
        Dim Name As String
        Dim MakerNum As Integer
        Dim MakeDate As Date
        
        Name = Application.InputBox("请输入您的姓名：")
        MakerNum = Application.InputBox(Prompt:="请输入编号", Type:=1)

        
        Range("B2").End(xlDown).Offset(1, 0).Value = Name
        Range("B2").End(xlDown).Offset(0, -1).Value = MakerNum

End Sub

Sub EnterFormula()
    Dim OurFormula As String
    Dim FormulaCell As Range
    
    OurFormula = Application.InputBox(Prompt:="请输入自定义公式：", Type:=0, Default:="=SUM(")
    
    Set FormulaCell = Application.InputBox(Prompt:="请选中应用公式的range：", Type:=8)
    
   FormulaCell.FormulaLocal = OurFormula
End Sub




Sub CopyAndPasteData()
    Dim CopyRange As Range
    Dim PasteRange As Range
    
    Set CopyRange = Application.InputBox(Prompt:="请拖选需要复制的内容：", Type:=8)
    Set PasteRange = Application.InputBox(Prompt:="请点击选择需要粘贴的位置：", Type:=8)
    
    CopyRange.Copy PasteRange
    
End Sub


