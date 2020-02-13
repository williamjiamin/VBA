Sub SelectSingleCell()
   '以下这个sub里面的方法，本质上都是模拟普通操作，先进行选中，后进行value输入
   
    '先激活/选中某一个工作簿（如果工作簿已经储存了，需要写上后缀）
    Workbooks("工作簿2.xlsx").Activate
    
    '先激活/选中某一个sheet
    
    Worksheets("Sheet1").Activate
    
    '方法1 VBA内用的比较多
    
    Range("A1").Select
    ActiveCell.Value = "商品编号"
    
    Range("B1").Select
    ActiveCell.Value = "商品名称"
    
    Range("C1").Select
    ActiveCell.Value = "商品库存"
    
    Range("E1").Select
    ActiveCell.Value = "最后一次进货日期"
    
    Range("A2").Select
    ActiveCell.Value = "001"
    
    '方法2 容易搞混，记住先行数，再列数，从1开始计数，混合开发中可用
    
    Cells(1, 4).Select
    ActiveCell.Value = "商品货源"
    
    Cells(2, 3).Select
    ActiveCell.Value = 999
    
    Cells(2, 4).Select
    ActiveCell.Value = "国产"
    
    
    
    '方法3 可看作方法1的升级版
    [B2].Select
    ActiveCell.Value = "口罩"
    
    [E2].Select
    ActiveCell.Value = #2/9/2020#
    
    
    
    
End Sub


Sub InputValueWithoutSelecting()
        Range("A3").Value = 2
        Range("B3").Value = "防护服"
        Range("C3").Value = 888
        Range("D3").Value = "进口"
        Range("E3").Value = #8/18/2020#
End Sub


Sub InputValueWithoutSelecting_V2()
        Workbooks("工作簿3").Worksheets("Sheet1").Range("A3").Value = 2
        Workbooks("工作簿3").Worksheets("Sheet1").Range("B3").Value = "防护服"
        Workbooks("工作簿3").Worksheets("Sheet1").Range("C3").Value = 888
        Workbooks("工作簿3").Worksheets("Sheet1").Range("D3").Value = "进口"
        Workbooks("工作簿3").Worksheets("Sheet1").Range("E3").Value = #8/18/2020#
End Sub











