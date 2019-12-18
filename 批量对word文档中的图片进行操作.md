# 批量对word文档中的图片进行操作

由于老婆大人有排版需求，需要对word文档中的图片进行批量修改大小以及修改下图片外观，于是翻看了word中VBA的代码帮助，找网上的牛人代码，遂实现了批量修改图片大小，以及给图片添加框线的功能，特此记录。

但是后来无论怎么找代码帮助，也没有找到该如何修改对齐方式为居中，暂时留个空白，以后再来补。

```vb
Sub picture()
Dim tol_shaps As Integer
Dim i As Integer
tol_shaps = ActiveDocument.InlineShapes.Count
'MsgBox tol_shaps

For i = 1 To tol_shaps
    With ActiveDocument.InlineShapes(i)
    '以下是调整图像的高度和宽度
    .ScaleHeight = 80 '将高度修改为原来图形的120%
    .ScaleWidth = 80  '将宽度修改为原来图形的120%
    

    
    
    
    '以下是修改图形的四个边框线，Borders属性
       With .Borders(wdBorderTop)
           .LineStyle = wdLineStyleSingle
           .LineWidth = wdLineWidth050pt
       End With
       With .Borders(wdBorderBottom)
           .LineStyle = wdLineStyleSingle
           .LineWidth = wdLineWidth050pt
       End With
       With .Borders(wdBorderLeft)
           .LineStyle = wdLineStyleSingle
           .LineWidth = wdLineWidth050pt
       End With
       With .Borders(wdBorderRight)
           .LineStyle = wdLineStyleSingle
           .LineWidth = wdLineWidth050pt
       End With
    End With
Next i
End Sub
```

