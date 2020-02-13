Sub SelectSingleCell()
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
