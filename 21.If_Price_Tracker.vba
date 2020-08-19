Option Explicit

Sub priceTracker()
        Dim Price As Integer
        Dim Msg As String
        Dim Index_Num As Integer
        
        Range("G4").Select
        Price = ActiveCell.Value
        Index_Num = ActiveCell.Offset(0, -2).Value
    
        
'      If Price < 10 Then
'           Msg = "可以买入"
'       Else
'           Msg = "等好时机买入"
'       End If
        
        
'        If Price < 10 Then Msg = "可以买入" Else Msg = "等好时机买入"

 '   If Price < 5 Then
 '       Msg = "赶紧买买买"
 '   ElseIf Price < 10 Then
 '       Msg = "考虑考虑，可以买一点试试"
 '   ElseIf Price < 15 Then
 '       Msg = "有点风险了，最好不买"
 '   Else
 '       Msg = "股价太高了，坚决不能买"
 '
 '    End If
 
 'nest if 会导致逻辑关系非常复杂，建议可以直接使用上方的if elseif else完成
 
'     If Price < 5 Then
'        Msg = "赶紧买买买"
'     Else
'            If Price < 10 Then
'                Msg = "考虑考虑，可以买一点试试"
'            Else
'                If Price < 15 Then
'                    Msg = "有点风险了，最好不买"
'                Else
'                    Msg = "股价太高了，坚决不能买"
'                End If
'            End If
'     End If



    If Price < 10 And Index_Num < 5 Then
          Msg = "可以买入"
      Else
          Msg = "等好时机买入"
    End If

        MsgBox " 您选中的股票当天价格为 ：" & Price & " 目前的建议是 " & Msg
End Sub
