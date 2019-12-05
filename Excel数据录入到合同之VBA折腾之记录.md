# Excel数据录入到合同之VBA折腾之记录

## 一、起因

国庆节前后，我们有大量的平浙信的项目反复做，但是这里面有个痛点，就是每次都需要进行一个财务测算，每期产品的底层资产从几笔到几百笔不等，测算的精度要求很高，整个过程几百上千个数据不能有一丁点的错误，而且，这些数据都需要按照一整套合同，对应的填写到不同的位置，大概我数了下，有60个左右的数据，有些重复填写十几二十次，分散在一套合同12份文件的上百处。

我最开始做的时候，好在项目只有一个，时间也比较充足，那么每完成一次测算并把数据更新到合同中，反复复核的话，一个人一整天的时间差不多够用。

后来，我首先在excel层面把word里需要的数据统一整理了一遍，然后逐个往word文档中复制粘贴，时间上缩短到了三个个小时完成一套合同。但是同时我们的业务量也增加了，过去的两周时间里同时推进了11个项目，没办法，整整耗尽了三个人的时间精力，就这样，还是被人发现这里或那里的错误。

所以，我觉得这样下去是不行的，不说效率问题，这些错误就无法容忍。于是我开始折腾，如何把excel数据直接导入到word文档中指定的位置。

## 二、初步解决方案——邮件合并

最简单的方式就是邮件合并了，但是我当然首先想到的解决方案是VBA ，也在网上找到了热浪老师的经典代码（http://club.excelhome.net/thread-477904-1-1.html )，无奈，时间紧，任务急，测试了几次还不是很完美，只能先使用邮件合并项目落地再说。

于是乎，我在帮助安臻项目挂牌的间隙，一个个的替换域，做完了一套邮件合并的合同模板，但是我发现这些模板有个问题，不能保留相对路径，只要我复制整个文件夹内容换个地方，所有的数据源都需要重新链接一次，有的是每次打开都需要更新，而且，导出断开链接的文档需要编辑单个文档的方式，另存为，修改文件名，好复杂。当然啦，相比之前的，至少保证了不会出错，不会遗漏，时间也已经大大缩短，平浙信11号全套合同，我熟练操作，花了5分钟。

## 三、终极解决方案——VBA

周日凌晨4点，睡不着了，还是折腾代码把，这次把热浪老师做的几个文档反复比对，终于发现了玄机，之前一直实现不了的原因是没有全部替换，结合热浪老师的代码《工程付款申请》里填写文字的代码（最后比对发现，这段代码才是全部替换）：

` '填写文字数据
         With .Selection.Find
            For j = 12 To 1 Step -1 '从大到小，防止字符串序号低位与高位串扰
               Str1 = "数据" & Format(j)
               Str2 = Sheets(数据表名).Cells(i, j + 1)
               .Text = Str1 '查找到指定字符串
               .Replacement.Text = Str2 '替换字符串
              .Execute Replace:=wdReplaceAll '全部替换
            Next j
         End With`

整体的代码使用的是《将Excel数据对应写入已做好的Word模板的指定位置(分发)》，但是这里面填写文字数据的代码没有使用全部替换，导致我一开始无论如何都不完美，这部分代码对应的是：

` For j = 25 To 1 Step -1 '填写文字数据,防止字符串序号低位与高位串扰
           Str1 = "数据" & Format(j, "000")
           Str2 = Sheets("数据").Cells(i, j + 1)
           .Selection.HomeKey Unit:=wdStory '光标置于文件首
           If .Selection.Find.Execute(Str1) Then '查找到指定字符串
              .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
              .Selection.Text = Str2 '替换字符串
           End If
        Next j`

最终，我修改代码以后，把这部分的代码修改为目前使用的代码：

` .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 25
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With`

以下是全部的代码：

`Private Sub CommandButton输出合同到Word文件_Click()
   Dim Word对象 As New Word.Application, 当前路径, 导出文件名, 导出路径文件名, i, j
   Dim Str1, Str2
   当前路径 = ThisWorkbook.Path
   最后行号 = Sheets("数据").Range("B65536").End(xlUp).Row
   判断 = 0
   For i = 3 To 最后行号
      导出文件名 = "信托合同：6份，每份骑缝章1各，签署页公章和法人章各1个，附件6预留印鉴公章和法人各1个"
      FileCopy 当前路径 & "\信托合同：6份，每份骑缝章1各，签署页公章和法人章各1个，附件6预留印鉴公章和法人各1个.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      导出路径文件名 = 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      With Word对象
        .Documents.Open 导出路径文件名
        .Visible = False
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 25
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        'For j = 1 To 3 '填写表格数据
         '  .ActiveDocument.Tables(1).Cell(2, j).Range = Sheets("数据").Cells(i, j + 6)
          ' .ActiveDocument.Tables(1).Cell(4, j).Range = Sheets("数据").Cells(i, j + 9)
        'Next j
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
        '.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter '设置位置在页脚
       ' Str1 = "数据007"
       ' Str2 = Sheets("数据2").Cells(2, 1)
        '.Selection.HomeKey Unit:=wdStory '光标置于文件首
       ' If .Selection.Find.Execute(Str1) Then '查找到指定字符串
        '   .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
         '  .Selection.Text = Str2 '替换字符串
        'End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
   If 判断 = 0 Then
      i = MsgBox("已输出到 Word 文件！", 0 + 48 + 256 + 0, "提示：")
   End If
End Sub`

----

`Private Sub CommandButton输出通知到Word文件_Click()

End Sub`

## 三、继续优化

以上，已经实现了对于单个合同文本，自合同模板复制后，重新命名，并且对于其中的数据进行全部替换了。但是，我想实现的终极目的是：

1. 自动读取项目名称（数据001），简历以该项目名称的文件夹
2. 将合同模板全套复制到新文件夹下（不需要修改名称）
3. 遍历该文件夹，查找所有的数据（001——048），全部替换

所以，接下来，继续折腾，直至完美。

### 1、自动创建文件夹名

折腾失败

### 2、转而用最笨的方法解决问题

把第一个 信托合同制作的代码，一次向下复制十次，把信托合同对应修改为其他合同，然后代码完成，看起来长的不行，难看的要死，可是我就是搞不定嘛，最后就变成了终极代码的复制再复制，不过代码复制，虽然难看，但是功能还是实现了的。



