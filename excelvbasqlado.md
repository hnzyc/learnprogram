## Excel VBA+ADO+SQL入门教程002：简单认识ADO

本文记录一下EXCEL VBA+ADO+SQL入门教程，不知道有没有用，学习一下先：

教程来源：https://mp.weixin.qq.com/s/O-clp1rAqU338ZlQPngftg

**对VBA无感或无基础者，本章可跳过，并不影响之后的SQL学习。**



示例文件下载

链接: https://pan.baidu.com/s/10Czj_9LtAKIVN5YatIdnAQ

提取码: sesv

## 1.

ADO (ActiveX Data Objects，ActiveX数据对象）是微软提出的应用程序接口，用以实现访问关系或非关系数据库中的数据……更多概念信息请自行咨询百度君，无赖脸。

之所以要学习ADO，一个原因是ADO自身的一些属性和方法对于数据处理是极其有益的；而首要原因是，在EXCEL VBA中，一般只有通过ADO，才可以使用强大的SQL查询语言访问外部数据源，进而查、改、增、删外部数据源中的数据。

后面这话延伸在具体编程操作上，就形成了四步走发展战略……

**1.VBA引用ADO类库。**

**2.ADO建立对数据源的链接。**

**3.ADO执行SQL语言。**

**4.VBA处理\**SQL查询结果。\****

嗯，这就好比你先找个女（男）朋友，然后谈恋爱，最后才能结婚……

**
**

## 2.

在VBA中引用ADO类库一般有两种方式。

一种是前期绑定。

所谓前期绑定，是指在VBE中手工勾选引用Microsoft ADO相关类库。

在Excel中，按<Alt+F11>快捷键打开VBA编辑窗口，依次单击【工具】→【引用】，打开【引用-VBAProject】对话框。在【可使用的引用】列表框中，勾选“Microsoft ActiveX Data Objects 2.8 Library”库，**或**“Microsoft ActiveX Data Objects 6.1 Library”库，单击【确定】按钮关闭对话框。

![img](https://mmbiz.qpic.cn/mmbiz_jpg/SbWgux809jX4xP2EaE2usqll0qusxmttc5j4MSaceJwjUpXibf3XLmUtSmNkh1ce01qMIzNS4aiaOfvvtd7KE9ag/640?wx_fmt=jpeg&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)      

另一种是使用代码后期绑定。

```vbscript
Sub 后期绑定()    Dim cnn As Object    
    Set cnn = CreateObject("adodb.connection")
End Sub
```

两种方式的主要区别是，前期绑定后，在代码编辑过程中，VBE的“自动列出成员”功能，可以提供ADO的属性和方法，这便于代码快捷、准确的编写，但当他人的Excel工作簿并没有手工前期绑定ADO类库时，相关代码将无法运行；因此后期代码绑定ADO的通用性会更强些，它不需要手工绑定相关类库。

星光俺老油……老江湖的经验是，代码编写及调试时，使用前期绑定，代码完善后，再修改为后期绑定发布使用。

## 3.

不论我们使用SQL语言对数据源作何操作，都得首先使用ADO创建并打开一个由VBA到数据源的链接；这就好比得先修路，才能使用汽车运输货物。

在VBA中，我们通常使用ADO的Connection.Open语句来显式建立一个到数据源的链接。

Connection.Open语法如下：

connection.Open ConnectionString, UserID, Password, Options

**ConnectionString可选，字符串，包含连接信息。**

UserID可选，字符串，包含建立连接时所使用用户名。

Password可选，字符串，包含建立连接时所使用密码。

Options可选，决定该方法是在连接建立之后（异步）还是连接建立之前（同步）返回，默认是同步，adAsyncConnect是异步。

……**语法看起来似乎很复杂？****不必烦扰，现在，对我们而言，重点只是****大体了解一下****参数ConnectionString，也就是连接字符串****。**虽然不同的数据库或文件有不同的连接字符串，但常用的数据库或文件的连接字符串均是固定的。

举个例子，如果将代码所在的Excel（2016版）作为一个外部数据源建立链接，代码如下：

```vbscript
Sub Mycnn()    
    Dim cnn As Object    '定义变量    
    Set cnn = CreateObject("adodb.connection")    '后期绑定ADO    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=yes;IMEX=0';Data Source=" & ThisWorkbook.FullName    '建立链接    
    cnn.Close    '关闭链接    
    Set cnn = Nothing    '释放内存
End Sub
```

说一下上面代码连接字符串中各关键字（字体加粗部分）的意思。

**Provider**是Connection 对象提供者名称的字符串值，03版Excel是“Microsoft.jet.OLEDB.4.0”，其它版本可以使用“Microsoft.ACE.OLEDB.12.0”；

**Extended Properties**是Excel版本号及其它相关信息，03版本是Excel 8.0，其它版本可以使用Excel 12.0。

其中**HDR项**是引用工作表是否有标题行，默认值HDR=Yes，意思是引用表的第一行是标题行，标题只能一行，不能多行，亦不能存在合并单元格。HDR=no，意思是引用表不存在标题行，也就是说第一行开始就是数据记录了；此时，相关字段名在SQL语句中可以使用f加序列号表示，第1列字段名是f1，第2列字段名是f2，其余以此类推，f是英文field(字段)的缩写。

**IMEX项**是汇入模式，默认为0（只读模式），1是只写，2是可读写。当参数设置为1时，除了只写，还有默认全部记录数据类型为文本的用途，关于这一点及其限制前提我们以后再谈。

**Data Source**是数据来源工作薄的完整路径。

VBA代码Application.Version可以获取计算机的Excel版本号，因此以下代码兼顾了03及各高级版本Excel的情况

```vbscript
Sub Mycnn3()    
    Dim cnn As Object    
    Dim strPath As String    
    Dim str_cnn As String    
    Set cnn = CreateObject("adodb.connection")    
    strPath = ThisWorkbook.FullName    
    If Application.Version < 12 Then        
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & strPath    
    Else        
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & strPath    
    End If    
    cnn.Open str_cnn    
    cnn.Close    
    Set cnn = Nothing
End sub
```

最后，需要提醒大家的是，链接是一种昂贵的资源（官方语），因此在代码运行完毕后，请养成关闭链接（cnn.Close）并释放内存（Set cnn = Nothing）的好习惯。

 **本节小贴士：**

3.1，连接字符串中各关键字的对应值可能和大小写有关，这是因为不同数据库的要求可能不一样，但通常来说，关键字和大小写无关，例如Provider，可以写成provider或者PROVIDER。不过，虽然关键字和大小写无关，但和拼写正确与否……当然是有关的！（想啥呢哥们？）当手打的连接字符串代码运行出错时，建议先复制正确的运行，再仔细核对个人错漏之处。

3.2，连接字符串中各关键字之间使用英文分号（;）间隔，例如（关键字1=值1;关键字2=值2;关键字3=值3……），另外，任何包含分号、单引号或双引号的值必须用双引号引起来，由于在VBA中连接字符串的外层已经存在了一个双引号，因此通常使用英文单引号进行转义，例如上例中的Extended Properties=**'**Excel 12.0;HDR=yes;IMEX=2**'**，抄写时，千万别漏了英文单引号哦。

3.3，星光俺掐指一算，算出相当一部分童鞋英语水平堪忧，想来拼写这段英文连接字符串错漏百出是很有可能的，因此特呈上锦囊一份，参见下图。别问我这图是哪来的，如果不几道，佛山无银脚，出门右拐重看第一章吧~

如果这锦囊您也不想用——其实收藏本帖，用到时打开帖子复制粘贴相关代码就可以了——嘿嘿，木错，这才是最常用的一招。

![img](https://mmbiz.qpic.cn/mmbiz_jpg/SbWgux809jX4xP2EaE2usqll0qusxmttX6sDwWQ8iaR2olo1ICopx9WVmQmH6JKexfM5DZ7TNpViar97jXyhS6tA/640?wx_fmt=jpeg&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)      

## 4.

聊完了如何绑定ADO以及建立与数据源的链接……

最后说下如何使用ADO执行SQL语句。

我们可以使用ADO的Connection对象或Recordset、Commannd执行SQL语句；详细内容我们放到ADO部分再讲；这里大家只需要先了解Connection对象的Execute方法就可以了。

这是一个最常用的VBA+ADO+SQL套路化查询代码，通常，我们只需要修改SQL语言以及放置查询结果的单元格位置。

```vbscript
Sub DoSql_Execute1()    
    Dim cnn As Object, rst As Object    
    Dim strPath As String, str_cnn As String, strSQL As String    
    Dim i As Long    
    Set cnn = CreateObject("adodb.connection")    
    '以上是第一步，后期绑定ADO    
    '    
    strPath = ThisWorkbook.FullName    
    If Application.Version < 12 Then        
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & strPath    
    Else        
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & strPath    
    End If    
    cnn.Open str_cnn    
    '以上是第二步，建立链接    
    '    
    strSQL = "SELECT 姓名,成绩 FROM [Sheet1$] WHERE 成绩>=80"    
    'SQL语句，查询Sheet1表成绩大于80……姓名和成绩的记录    
    Set rst = cnn.Execute(strSQL)    
    'cnn.Execute()执行SQL语句，始终得到一个新的记录集rst    
    '以上是第三步，编写并使用SQL语句   
    '    
    [d:e].ClearContents    '清空[d:e]区域的值    
    For i = 0 To rst.Fields.Count - 1    
        '利用fields属性获取所有字段名，fields包含了当前记录有关的所有字段,fields.count得到字段的数量    
        '由于Fields.Count下标为0，又从0开始遍历，因此总数-1        
        Cells(1, i + 4) = rst.Fields(i).Name    
    Next    
    Range("d2").CopyFromRecordset rst    
    '使用单元格对象的CopyFromRecordset方法将rst内容复制到D2单元格为左上角的单元格区域    
    '以上是第四步，将SQL查询结果和字段名写入表格指定区域    
    '    
    cnn.Close    '关闭链接    
    Set cnn = Nothing    '释放内存
End Sub
```

呵，总结一下：

**对于新手而言**，本章的重点是了解VBA执行SQL的操作过程，以及懂得复制**第4节的**代码执行SQL语句，仅此而已，其它？看过就算，大概过一眼，留个印象，以后再见面好说话也就行了。