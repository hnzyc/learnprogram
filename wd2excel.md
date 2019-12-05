# excel和word中的数据互相写入的代码

在前面解决实际项目合同批量生成的过程中，最后有一个资料清单和分配表处，一直没有解决，因为我还无法解决从excel表格数据直接写入word对应的表格中，恰好今天就在公众号[VBA编程学习与实践](javascript:void(0);)中就看到了这个帖子，赶紧把代码抄下来，像笑来老师倡导的那样，多读程序，只有输入足够多了，才会慢慢的增加输出。

具体代码如下：

```vbscript
Sub ExcelTableToWord()
    Dim WdApp As Object
    Dim objTable As Object
    Dim objDoc As Object
    Dim strPath As String
    Dim arr As Variant, brr As Variant
    Dim k As Long, x As Long, y As Long
    Dim i As Long, j As Long, Clny As Long
    On Error Resume Next
    Set WdApp = CreateObject("Word.Application")
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Word文件", "*.doc*", 1
        '只显示word文件
        .AllowMultiSelect = False
        '禁止多选文件
        If .Show Then strPath = .SelectedItems(1) Else Exit Sub
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    arr = [a1].CurrentRegion
    'excel表格数据读入数组arr
    Set objDoc = WdApp.documents.Open(strPath)
    '后台打开用户选定的word文档
    For Each objTable In objDoc.tables
    '遍历word中的表格
            x = objTable.Rows.Count'注意这里的.Count方法，可以用在前面代码里，这样就不用每次手动调整最大循环次数了
        y = objTable.Columns.Count
        For j = 1 To y
        '遍历表格的标题行，默认标题处于第一行
            If Application.Clean(objTable.Cell(1, j).Range.Text) = arr(1, j) Then
            '如果标题行一致，则将excel表数据写入word
                For i = 2 To x
                    With objTable.Cell(i, j).Range
                        .Text = ""
                        .Text = arr(i, j)
                    End With
                Next
            End If
        Next
    Next
    objDoc.Close True: WdApp.Quit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set objDoc = Nothing
    Set WdApp = Nothing
    MsgBox "处理完成。"
End Sub
```

接下来，就是word内容写入excel的代码了

```vbscript
Sub GetWordTable()
    '读取word中的表格数据到excel
    Dim WdApp As Object
    Dim objTable As Object
    Dim objDoc As Object
    Dim strPath As String
    Dim shtEach As Worksheet
    Dim shtSelect As Worksheet
    Dim k As Long, x As Long, y As Long
    Dim i As Long, j As Long
    Dim brr As Variant
    On Error Resume Next
    Set WdApp = CreateObject("Word.Application")
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Word文件", "*.doc*", 1
        '只显示word文件
        .AllowMultiSelect = False
        '禁止多选文件
        If .Show Then strPath = .SelectedItems(1) Else Exit Sub
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set shtSelect = ActiveSheet
    '当前表赋值变量shtSelect，方便代码运行完成后叶落归根回到开始的地方
    For Each shtEach In Worksheets
    '删除当前工作表以外的所有工作表
        If shtEach.Name <> shtSelect.Name Then shtEach.Delete
    Next
    shtSelect.Name = "EH看见星光"
    '这句代码不是无聊，作用在于……你猜……
    '……其实是避免下面的程序工作表名称重复
    Set objDoc = WdApp.documents.Open(strPath)
    '后台打开用户选定的word文档
    For Each objTable In objDoc.tables
    '遍历文档中的每个表格
        k = k + 1
        Worksheets.Add after:=Worksheets(Worksheets.Count)
        '新建工作表
        ActiveSheet.Name = k & "表"
        objTable.Range.Copy
        '整表复制
        ActiveSheet.Paste
        'word表粘贴到excel，保留word表的格式
        '整表复制的方法无法避免身份证之类数据的变形,如果有这样的数据，最好使用如下单元格遍历
        x = objTable.Rows.Count
        'table的行数
        y = objTable.Columns.Count
        'table的列数
        ReDim brr(1 To x, 1 To y)
        '以下遍历行列，数据写入数组brr
        For i = 1 To x
            For j = 1 To y
                brr(i, j) = "'" & Application.Clean(objTable.Cell(i, j).Range.Text)
                'Clean函数清除制表符等
                '半角单引号将数据统一转换为文本格式，避免身份证等数值变形
            Next
        Next
        With [a1].Resize(x, y)
            .Value = brr
            '数据写入Excel工作表
            .Borders.LineStyle = 1
            '添加边框线
        End With
    Next
    shtSelect.Select
    objDoc.Close: WdApp.Quit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set objDoc = Nothing
    Set WdApp = Nothing
    MsgBox "共获取：" & k & "张表格的数据。"
End Sub
```

