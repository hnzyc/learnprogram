# **FSO之文件及文件夹操作及获取相关属性等**

本文摘自excelhome论坛liulang0808的[文章](http://club.excelhome.net/forum.php?mod=viewthread&tid=1174170&extra=&authorid=238368&page=1)，主要是太经典了，怕将来不开源了，记录下来，以备常常学习和借用。

总体的代码框架如下：

```vbscript
Sub 按钮1_Click()
   Application.ScreenUpdating = False
    Set fso =CreateObject("Scripting.FileSystemObject")
'    此处根据具体操作添加代码
   Application.ScreenUpdating = True
End Sub
```

## 一、文件有关的操作

1. 判断文件是否存在

FileExists方法用于判断指定的文件是否存在，若存在则返回True。其语法为：
fso.FileExists(Filepath)
Filepath为文件完整路径，String类型，不能包含有通配符。如果用户有充分的权限，Filepath可以是网络路径或共享名
示例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    strfile = Application.InputBox("请输入文件的完整名称:", "请输入文件的完整名称:", , , , , , 2)
    If fso.fileexists(strfile) Then
        MsgBox strfile & " :存在"
    Else
        MsgBox strfile & " :不存在"
    End If
    Application.ScreenUpdating = True
End Sub
```

2. 移动文件

MoveFile方法用来移动文件，将文件从一个文件夹移动到另一个文件夹。其语法为：
FSO.MoveFile source,destination
参数source必需，指定要移动的文件的路径，String类型。参数destination必需，指定文件移动操作中的目标位置的路径，String类型。
如果Source包含通配符或者destination以路径分隔符结尾，则认为destination是一个路径，否则认为destination的最后一部分是文件名。
如果目标文件已经存在，则将出现一个错误。
source可以包含通配符，但只能出现在它的最后一部分中。
destination参数不能包含通配符。
source或destination可以是相对路径或绝对路径，可以是网络路径或共享名。
MoveFile方法在开始操作前先解析source和destination这两个参数。
实例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourcefile = ThisWorkbook.Path & "\txt\*" '将txt文件下所有文件移走，首先需要确认是相关文件时针的存在

    destinationfolder = ThisWorkbook.Path & "\tt\" '注意路径格式“tt\”，后面的“\”

    fso.movefile sourcefile, destinationfolder
    Application.ScreenUpdating = True
End Sub
```

3. 拷贝文件

CopyFile方法用来复制文件，将文件从一个文件夹复制到另一个文件夹。其语法为：
fso.CopyFile Source,Destination [,OverwriteFiles]
参数Source必需，指定要复制的文件的路径和名称，String类型。参数Destination必需，代表复制文件的目标路径和文件名（可选），String类型。参数OverwriteFiles可选，表示是否覆盖一个现有文件的标志，True表示覆盖，False表示不覆盖，Boolean类型，默认值为True。
参数source中源路径可以是绝对路径或相对路径，源文件名可包含通配符但源路径不能。在参数Destination中不能包含通配符。
如果目标路径或文件设置为只读，则无论OverwriteFiles参数的值如何，都将无法完成CopyFile方法。如果参数OverwriteFiles设置为False且Destination指定的文件已经存在，则会产生一个运行时错误“文件已经存在”。如果在复制多个文件时出现错误，CopyFile方法将立即停止复制操作，该方法不具有撤销错误前文件复制操作的返回功能。如果用户有充分的权限，那么source或destination可以是网络路径或共享名。 CopyFile方法可以复制一个保存在特定文件夹中的文件。如果文件夹本身有包含文件的子文件夹，则使用CopyFile方法不能复制这些文件，应该使用CopyFolder方法。
具体实例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourcefile = ThisWorkbook.Path & "\txt\*" '将txt文件下所有文件拷贝走，首先需要确认是相关文件时针的存在
    destinationfolder = ThisWorkbook.Path & "\tt\" '注意此处不同于movefile，后面的“\”可以省略，只要确实存在该文件夹
    fso.copyfile sourcefile, destinationfolder
    Application.ScreenUpdating = True
End Sub
```

4. 删除文件

DeleteFile方法删除指定的一个或多个文件。其语法为：
fso.DeleteFile FileSpec[,Force]
参数FileSpec必需，代表要删除的单个文件或多个文件的名称和路径，String类型，可以在路径的最后部分包含通配符，可以为相对路径或绝对路径。如果在FileSpec中只有文件名，则认为该文件在应用程序的当前驱动器和文件夹中。参数Force可选，如果将其设置为True，则忽略文件的只读标志并删除该文件，Boolean类型，默认值为False。
如果指定要删除的文件已经打开，该方法将失败并出现一个“Permission Denied”错误。如果找不到指定的文件，则该方法失败。
如果在删除多个文件的过程中出现错误，DeleteFile方法将立即停止删除操作，即不能删除余下的文件部分。该方法不具有撤销产生错误前文件删除操作的返回功能。
如果用户有充分的权限，源路径或目标路径可以是网络路径或共享名。
注意：DeleteFile方法永久性地删除文件，并不把这些文件移到回收站中。
示例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    strfile = Application.InputBox("请输入文件的完整名称:", "请输入文件的完整名称:", , , , , , 2)
    fso.deletefile strfile
    Application.ScreenUpdating = True
End Sub
```

5. GetFile方法

GetFile方法用来返回一个File对象。
其语法为：fso.GetFile (FilePath)
参数FilePath必需，指定路径和文件名，String类型。可以是绝对路径或相对路径。如果FilePath是一个共享名或网络路径，GetFile确认该驱动器或共享是File对象创建进程的一部分。如果参数FilePath指定的路径的任何部分不能连接或不存在，就会产生错误。
GetFile方法返回的是File对象，而不是TextStream对象。File对象不是打开的文件，主要是用来完成如复制或移动文件和询问文件的属性之类的方法。尽管不能对File对象进行写或读操作，但可以使用File对象的OpenAsTextStream方法获得TextStream对象。
要获得所需的FilePath字符串，首先应该使用GetAbsolutePathName方法。如果FilePath包含网络驱动器或共享，可以在调用GetFile方法之前用DriveExists方法来检验所需的驱动器是否可用。
因为在FilePath指定的文件不存在时会产生错误，所以应该在调用GetFile之前调用FileExists方法确定文件是否存在。
必须用Set语句将File对象赋给一个局部对象变量。
**具体实例见下面的属性代码**

6. 文件的各种属性

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    strfile = Application.InputBox("请输入文件的完整名称:", "请输入文件的完整名称:", , , , , , 2)
    Set objfile = fso.GetFile(strfile)
    If fso.fileexists(strfile) Then
      
        sReturn = "文件属性： " & objfile.Attributes & vbCrLf
         
        sReturn = sReturn & "文件创建日期： " & objfile.DateCreated & vbCrLf
         
        sReturn = sReturn & "文件修改日期： " & objfile.DateLastModified & vbCrLf
         
        sReturn = sReturn & "文件大小 " & FormatNumber(objfile.Size / 1024, -1)
         
        sReturn = sReturn & "Kb" & vbCrLf
         
        sReturn = sReturn & "文件类型： " & objfile.Type & vbCrLf

        MsgBox sReturn

    Else
        MsgBox strfile & " :不存在"
    End If
    Application.ScreenUpdating = True
End Sub
```

## 二、文件夹操作

1. 判断文件夹是否存在

FolderExists方法可以判断指定的文件夹是否存在，若存在则返回True。其语法为：
fso.FolderExists(FolderSpec)
参数FolderSpec指定文件夹的完整路径，String类型，不能包含通配符。
如果用户有充分的权限，FolderSpec可以是网络路径或共享名，例如：
If fso.FileExists ("\\NTSERV1\d$\TestPath\") Then

示例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    strfile = Application.InputBox("请输入文件的完整名称:", "请输入文件的完整名称:", , , , , , 2)
    If fso.fileexists(strfile) Then
        MsgBox strfile & " :存在"
    Else
        MsgBox strfile & " :不存在"
    End If
    Application.ScreenUpdating = True
End Sub
```

2. 移动

MoveFolder方法用来移动文件夹，将文件夹及其文件和子文件夹一起从某个位置移动到另一个位置。其语法为：
fso.MoveFolder source,destination
参数Source指定要移动的文件夹的路径，String类型。参数destination指定文件夹移动操作中目标位置的路径，String类型。
Source必须以通配符或非路径分隔符结束，可以使用通配符，但必须出现在最后一部分中。destination不能使用通配符。除非不允许使用通配符，否则源文件夹中所有的子文件夹和文件都被复制到destination指定的位置，也就是说MoveFolder方法是递归的。
如果destination用路径分隔符结束或者source用通配符结束，MoveFolder就认为source中指定的文件夹存在于destination中。例如，假设有如下文件夹结构：
MoveFolder "C:\Rootone\*","C:\RootTow\"
产生如下文件夹结构：
MoveFolder "C:\Rootone","C:\RootTwo\"
产生如下文件夹结构：
Source和destination可以为绝对路径或相对路径，可以为网络路径或共享名。
MoveFile方法在开始操作前先解析source和destination这两个参数。

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = ThisWorkbook.Path & "\tt"
    dfolder = ThisWorkbook.Path & "\txt\"
    If Not fso.folderexists(sfolder) Then
        MsgBox sfolder & " :不存在"
        Exit Sub
    End If
   
    If Not fso.folderexists(dfolder) Then
        MsgBox dfolder & " :不存在"
        Exit Sub
    End If
    fso.movefolder sfolder, dfolder
    Application.ScreenUpdating = True
End Sub
```

3. 拷贝

CopyFolder方法用于复制文件夹，即将一个文件夹的内容（包括其子文件夹）复制到其他位置。其语法为：
fso.CopyFolder Source,Destination[,OverwriteFiles]
参数Source必需，指定要复制的文件夹的路径和文件夹名，String类型，必须使用通配符或者非路径分隔符来结束。参数Destination必需，指定文件夹复制操作的目标文件夹的路径，String类型。参数OverwriteFiles可选，表示是否被覆盖一个现有文件的标志，True表示覆盖，False表示不覆盖，Boolean类型。
通配符只能在参数Source中使用，但是只能放在最后的组件中。在参数Destination中不能使用通配符。
除非不允许使用通配符，否则就可以把源文件夹中的所有子文件夹和文件都复制到Destination指定的文件夹中，也就是说CopyFolder方法是递归的。
如果参数Destination以一个路径分隔符结束或者参数Source以一个通配符结束，CopyFolder方法就认为参数Source中的指定的文件夹存在于参数Destination中，否则就创建这样一个文件夹。例如，假设有如下的文件夹结构：
CopyFolder "C:\Rootone\*","C:\RootTwo"
产生如下的文件夹结构：
CopyFolder "C:\Rootone","C:\RootTwo\"
产生如下的文件夹结构：
如果参数Destination指定的目标路径或任意文件被设置成只读属性，则不论OverwriteFiles的值如何，CopyFolder方法者将失效。
如果OverwriterFiles设置为False，而参数Source指定的源文件夹或任何文件存在于参数Destination中，将产生运行时错误“文件已经存在”。
如果在复制多个文件夹时出现错误，CopyFolder方法立即停止复制操作，不再复制余下要复制的文件。该方法不具有撤销产生错误前文件复制操作的返回功能。
如果用户有充分的权限，source或destination都可以是网络路径或共享名，例如：
CopyFolder "C:\Rootone","\\NTSERV1\d$\RootTwo\"

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = ThisWorkbook.Path & "\tt"
    dfolder = ThisWorkbook.Path & "\txt\"
    If Not fso.folderexists(sfolder) Then
        MsgBox sfolder & " :不存在"
        Exit Sub
    End If
   
    If Not fso.folderexists(dfolder) Then
        MsgBox dfolder & " :不存在"
        Exit Sub
    End If
    fso.copyfolder sfolder, dfolder
    Application.ScreenUpdating = True
End Sub
```

4. 删除文件夹

DeleteFolder方法用于删除指定的文件夹及其所有的文件和子文件夹。其语法为：
fso.DeleteFolder FileSpec[,Force]
参数FileSpec必需，指定要删除的文件夹的名称和路径，String类型。在参数FileSpec中，可以在路径的最后部分包含通配符，但不能用路径分隔符结束，可以为相对路径或绝对路径。
参数Force可选，Boolean类型，如果设置为True，将忽略文件的只读标志并删除这个文件。默认为False。如果参数Force设置为False并且文件夹中的任意一个文件为只读，则该方法将失败。如果找不到指定的文件夹，则该方法失败。
如果指定的文件夹中有文件已经打开，则不能完成删除操作，且产生一个“Permisson Denied”错误。DeleteFolder方法删除指定文件夹中的所有内容，包括其他文件夹及其内容。
如果在删除多个文件或文件夹时出现错误，DeleteFolder方法将立即停止删除操作，即不能删除余下的文件夹或文件。该方法不具有撤销产生错误前文件夹删除操作的返回功能。
DeleteFolder方法永久性删除文件夹，并不把它们移到回收站中。
如果用户有充分的权限，源路径和目标路径可以是网络路径或共享名，例如：
DeleteFolder "\\RootTest"

示例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = ThisWorkbook.Path & "\txt\tt"
    If Not fso.folderexists(sfolder) Then
        MsgBox sfolder & " :不存在"
        Exit Sub
    End If
   
    fso.deletefolder sfolder
    Application.ScreenUpdating = True
End Sub
```

5. 创建文件夹

CreateFolder方法用于在指定的路径下创建一个新文件夹，并返回其Folder对象。其语法为：
fso.CreateFolder (Path)
参数Path必需，为一个返回要创建的新文件夹名的表达式，String类型。Path指定的路径可以是相对路径也可以是绝对路径，如果没有指定路径则使用当前驱动器和目录作为路径。在新的文件夹名中不能使用通配符。
如果参数Path指定的路径为只读，则CreateFolder方法将失败；如果参数Path指定的文件夹已经存在，就会产生运行时错误“文件已经存在”。如果用户有充分的权限，则参数Path可以指定为网络路径或共享名，例如：
Fso.CreateFolder "\\NTSERV1\RootTest\newFolder"
示例如下：

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = ThisWorkbook.Path & "\thisfolder"
    If fso.folderexists(sfolder) Then
        MsgBox sfolder & " :已经存在"
        Exit Sub
    End If
   
    fso.CreateFolder sfolder
    Application.ScreenUpdating = True
End Sub
```

6. GetAbsolutePathName方法

将相对路径转变为一个全限定路径（包括驱动器名），返回一个字符串，包含一个给定的路径说明的绝对路径。其语法为：
fso.GetAbsolutePathName (Path)
参数Path必需，代表路径说明，String类型。
“.”返回当前文件夹的驱动器名和完整路径。“..”返回当前文件夹的父文件夹的驱动器名和路径。“filename”返回当前文件夹中的文件的驱动器名、路径及文件名。
所有相对路径名均以当前文件夹为基准。
如果没有明确地提供驱动器作为Path的一部分，就以当前驱动器作为Path参数中的驱动器。在Path中可以包含任意个通配符。
对于映射网络驱动器和共享而言，这种方法不能返回完整的网络地址，而是返回全限定的本地路径和本地驱动器名。
GetAbsolutePathName不能检验指定路径中是否存在某个给定的文件或文件夹

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = "thisfolder"
    If fso.folderexists(sfolder) Then
        MsgBox sfolder & " :已经存在"
        Exit Sub
    End If
   
    str1 = fso.GetAbsolutePathName(sfolder)
    MsgBox sfolder & "  ：的绝对路径为： " & str1
    Application.ScreenUpdating = True
End Sub
```

8. GetParentFolderName方法

返回给定路径中最后部分前的文件夹名，其语法为：
fso.GetParentFolderName (Path)
参数Path必需，指定路径说明，String类型。
如果从Path中不能确定父文件夹名，就返回一个零长字符串（””）。Path可以为相对路径或绝对路径。可以是网络驱动器或共享。
GetParentFolderName方法不能检验Path的某个部分是否存在。
GetParentFolderName方法认为Path中不属于驱动器说明的那部分字符串除了最后一部分外余下的字符串就是父文件夹。除此之外它不做任何其他检测，更像是一个字符串解析和处理例程而不是与对象处理有关的例程。

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    sfolder = ThisWorkbook.Path & "\tt\"
    If Not fso.folderexists(sfolder) Then
        MsgBox sfolder & " :不存在"
        Exit Sub
    End If
   
    str1 = fso.GetParentFolderName(sfolder)
    MsgBox sfolder & "  ：父路径： " & str1
    Application.ScreenUpdating = True
End Sub
```

9. GetSpecialFolder方法

GetSpecialFolder方法返回操作系统文件夹路径，其中0代表Windows文件夹，1代表System（系统）文件夹，2代表Temp（临时）文件夹。其语法为：
fso.GetSpecialFolder (SpecialFolder)
参数SpecialFolder必需，为特殊的文件夹常数，表示三种特殊系统文件夹中其中一个的值。
可以使用Set语句将Folder对象赋给一个局部对象变量，但是如果只对检索特殊的文件夹感兴趣，就可以使用下列语句来实现：
sPath=fso.GetSpecialFolder (iFolderConst)
或：
sPath=fso.GetSpecialFolder (iFolderConst).Path
由于Path属性是Folder对象的缺省属性，所认第一个语句有效。因为不是给一个对象变量赋值，所以赋给sPath的值是缺省的Path属性值，而不是对象引用。
示例

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False
    Dim strWindowsFolder As String
    Dim strSystemFolder As String
    Dim strTempFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    strWindowsFolder = fso.GetSpecialFolder(0)
    strSystemFolder = fso.GetSpecialFolder(1)
    strTempFolder = fso.GetSpecialFolder(2)
    MsgBox strWindowsFolder & vbCrLf & strSystemFolder & vbCrLf _
    & strTempFolder, vbInformation + vbOKOnly, "Special Folders"
    Application.ScreenUpdating = True
End Sub
```

10. GetFolder方法

GetFolder方法返回Folder对象。其语法为：
fso.GetFolder (FolderPath)
参数FolderPath必需，指定所需文件夹的路径，String类型，可以为相对路径或绝对路径。如果FolderPath是共享名或网络路径，GetFolder确认该驱动器或共享是File对象创建进程的一部分。如果FolderPath的任何部分不能连接或不存在，就会产生一个错误。
要获得所需的Path字符串，首先应该使用GetAbsolutePathName方法。如果FolderPath包含一个网络驱动器或共享，可以在调用GetFolder方法之前使用DriveExists方法确认指定的驱动器是否可用。由于GetFolder方法要求FolderPath是一个有效文件夹的路径，所以应调用FolderExists方法来检验FolderPath是否存在。
必须使用Set语句将Folder对象赋给一个局部对象变量。
具体实例见楼下 属性获取

```vbscript
Sub 按钮1_Click()
    Application.ScreenUpdating = False

    Dim sReturn As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder1 = fso.GetFolder(ThisWorkbook.Path & "\")
    sReturn = "文件夹属性： " & folder1.Attributes & vbCrLf
    '获取最近一次访问的时间
    sReturn = sReturn & "创建时间： " & folder1.Datecreated & vbCrLf
    sReturn = sReturn & "最后访问时间： is " & folder1.DateLastAccessed & vbCrLf
    '获取最后一次修改的时间
    sReturn = sReturn & "最后修改时间： " & folder1.DateLastModified & vbCrLf
    '获取文件夹的大小
    sReturn = sReturn & "文件夹大小： " & FormatNumber(folder1.Size / 1024, 0)
    sReturn = sReturn & "Kb" & vbCrLf
    '判断文件或文件夹类型
    sReturn = sReturn & "类型为： " & folder1.Type & vbCrLf
    MsgBox sReturn
    Application.ScreenUpdating = True
End Sub
```

## 三、另一牛人的代码- **跟我学 【喜迎2015立春】遍历文件夹(含子文件夹)方法 ABC**

以下内容摘自另一牛人香川群子的[文章](http://club.excelhome.net/thread-1185089-1-1.html)，其中FSO操作代码简洁，带有弹出框选择目标位置，也摘录下来：

首先要介绍，在VBA代码运行以后，调用【目标文件夹】的方法：

① 微软Excel VBA 默认选择文件夹的Dialog对话框

```vbscript
Sub ListFilesTest()
    With Application.FileDialog(msoFileDialogFolderPicker) '运行后出现标准的选择文件夹对话框        
        If .Show Then myPath = .SelectedItems(1) Else Exit Sub '如选中则返回=-1 / 取消未选则返回=0 
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & "" 
    '返回的是选中目标文件夹的绝对路径,但除了本地C盘、D盘会以"C:"形式返回外，其余路径无""需要自己添加 
End Sub
```

② 视窗浏览器界面选择目标文件夹

```vbscript
Sub ListFilesTest()
    Set myFolder = CreateObject("Shell.Application").BrowseForFolder(0, "GetFolder", 0)
    If Not myFolder Is Nothing Then myPath$ = myFolder.Items.Item.Path Else MsgBox "Folder not Selected": Exit Sub
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    '同样返回的是选中目标文件夹的绝对路径,但除了本地C盘、D盘会以"C:"形式返回外，其余路径无""需要添加 

End Sub
```

这两种选择目标文件夹的方法，总的效果应该都不错。
方法-1 默认Dialog对话框左侧栏有桌面、我的文档等快捷方式，也比较符合一般人的使用习惯。
优点是，本层文件夹内的子文件夹全部以大图标方式列出（也可以改为列表）看起来较为轻松。
缺点是，如果有多层子文件夹，需要一层一层地点下去……似乎比较累一点。

与此相对、方法-2 是浏览器形式，点击+号可以展开、点击-号可以折叠。
因此也有很多人特别喜欢这一种的，尤其是有多层子文件夹时很方便。

1. 仅列出目标文件夹中所有文件。（不包括 子文件夹、不包括子文件夹中的文件）

```vbscript
Sub ListFilesTest()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    '以上选择目标文件夹以得到路径myPath

    MsgBox ListFiles(myPath)    '调用FSO的ListFiles过程返回目标文件夹下的所有文件名
    
End Sub

Function ListFiles(myPath$)
   Set fso = CreateObject("Scripting.FileSystemObject") '打开FSO脚本、建立FSO对象实例
   For Each f In fso.GetFolder(myPath).Files  '用FSO方法遍历指定文件夹内所有文件
      i = i + 1: s = s & vbCr & f.Name            '逐个列出文件名并统计文件个数 i
   Next
   ListFiles = i & " Files:" & s  '返回所有文件名的合并字符串
End Function
```

知识介绍：
Set fso = CreateObject("Scripting.FileSystemObject")
建立FSO 即【文件系统对象】的实例。

这以后，即可简单、直接地引用fso的各种属性（有时间可以自己慢慢研究）

For Each f In fso.GetFolder(myPath).Files
'用FSO方法遍历指定文件夹内所有文件

fso.GetFolder(myPath) 是指对于路径myPath，使用FSO对象方法得到其文件夹.GetFolder属性
然后，对于这个指定的目标文件夹，继续返回其所有文件的属性、即.Files
完整的部分为： fso.GetFolder(myPath).Files

然后，对于这个所有文件的集合即 fso.GetFolder(myPath).Files
通过For……Each循环就可以遍历其中每一个文件了。

具体地，For Each f In 中的f变量，即为每一个文件。
循环中，可以使用f的各种属性。 f.Name只是其中的一种属性=文件名。

2. 仅列出目标文件夹中所有子文件夹名。（不包括目标文件夹中文件、不包括子文件夹中的文件或子文件夹）

```vbscript
Sub ListFilesTest()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    
    MsgBox ListFolders(myPath)
    
End Sub
Function ListFolders(myPath$)
   Set fso = CreateObject("Scripting.FileSystemObject")
   For Each f In fso.GetFolder(myPath).SubFolders
      j = j + 1: t = t & vbCr & f.Name
   Next
   ListFolders = j & " Folders:" & t
End Function
```

和楼上的代码ListFiles相比，差异很小，仅在于：
fso.GetFolder(myPath).Files
fso.GetFolder(myPath).SubFolders

即，把目标文件夹fso.GetFolder(myPath)的属性，
有.Files 所有文件、改为 .SubFolders 所有子文件夹

3. 遍历目标文件夹内所有文件、以及所有子文件夹中的所有文件

```vbscript
Sub ListFilesTest()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    
    [a:a] = ""                    '清空A列
    Call ListAllFso(myPath)   '调用FSO遍历子文件夹的递归过程
    
End Sub

Function ListAllFso(myPath$) '用FSO方法遍历并列出所有文件和文件夹名的【递归过程】
    Set fld = CreateObject("Scripting.FileSystemObject").GetFolder(myPath)
    '用FSO方法得到当前路径的文件夹对象实例 注意这里的【当前路径myPath是个递归变量】

    For Each f In fld.Files  '遍历当前文件夹内所有【文件.Files】
        [a65536].End(3).Offset(1) = f.Name '在A列逐个列出文件名
    Next

    For Each fd In fld.SubFolders  '遍历当前文件夹内所有【子文件夹.SubFolders】
        [a65536].End(3).Offset(1) = " " & fd.Name & ""  '在A列逐个列出子文件夹名
        Call ListAllFso(fd.Path)       '注意此时的路径变量已经改变为【子文件夹的路径fd.Path】
        '注意重点在这里： 继续向下调用递归过程【遍历子文件夹内所有文件文件夹对象】
    Next
End Function

```

由于很多初学者不太能理解递归算法的过程而产生畏难、抵触情绪，
所以下面避开递归，而采用字典记录中间结果的方法，同样来达到遍历所所有子文件的目的(不过个人觉得还不如递归呢)：

```vbscript
Sub ListFilesTest()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    
    MsgBox "List Files:" & vbCr & Join(ListAllFsoDic(myPath), vbCr)
    MsgBox "List SubFolders:" & vbCr & Join(ListAllFsoDic(myPath, 1), vbCr)
End Sub

Function ListAllFsoDic(myPath$, Optional k = 0) '使用2个字典但无需递归的遍历过程
    Dim i&, j&
    Set d1 = CreateObject("Scripting.Dictionary") '字典d1记录子文件夹的绝对路径名    
    Set d2 = CreateObject("Scripting.Dictionary") '字典d2记录文件名 （文件夹和文件分开处理）

    d1(myPath) = ""           '以当前路径myPath作为起始记录，以便开始循环检查
 
    Set fso = CreateObject("Scripting.FileSystemObject")
    Do While i < d1.Count
    '当字典1文件夹中有未遍历处理的key存在时进行Do循环 直到 i=d1.Count即所有子文件夹都已处理时停止

        kr = d1.Keys '取出文件夹中所有的key即所有子文件夹路径 （注意每次都要更新）
        For Each f In fso.GetFolder(kr(i)).Files '遍历该子文件夹中所有文件 （注意仅从新的kr(i) 开始）
            j = j + 1: d2(j) = f.Name
           '把该子文件夹内的所有文件名作为字典Item项加入字典d2 (为防止文件重名不能用key属性)
        Next

        i = i + 1 '已经处理过的子文件夹数目 i +1 （避免下次产生重复处理）
        For Each fd In fso.GetFolder(kr(i - 1)).SubFolders '遍历该文件夹中所有新的子文件夹
            d1(fd.Path) = " " & fd.Name & "" 
            '把新的子文件夹路径存入字典d1以便在下一轮循环中处理
        Next
    Loop

    If k Then ListAllFsoDic = d1.Keys Else ListAllFsoDic = d2.Items
    '如果参数=1则列出字典d1中所有子文件夹的路径名 (如使用d1.Items则仅列出子文件夹名称不含路径)
    '如果参数=0则默认列出字典d2中Items即所有文件名 

End Function

```

4. 作为本帖的特色，介绍使用VBA语句直接调用Dos中Dir命令来搜寻文件名的方法：(个人感觉不如FSO)

```vbscript
Sub ListFilesDos()
    Set myFolder = CreateObject("Shell.Application").BrowseForFolder(0, "GetFolder", 0)
    If Not myFolder Is Nothing Then myPath$ = myFolder.Items.Item.Path Else MsgBox "Folder not Selected": Exit Sub
    
    myFile$ = InputBox("Filename", "Find File", ".xl")
    '在这里输入需要指定的关键字，可以是文件名的一部分，或指定文件类型如 ".xl"
    tms = Timer
    With CreateObject("Wscript.Shell") 'VBA调用Dos命令
        ar = Split(.exec("cmd /c dir /a-d /b /s " & Chr(34) & myPath & Chr(34)).StdOut.ReadAll, vbCrLf) '所有文档含子文件夹
        '指定Dos中Dir命令的开关然后提取结果 为指定文件夹以及所含子文件夹内的所有文件的含路径全名。
        s = "from " & UBound(ar) & " Files by Search time: " & Format(Timer - tms, " 0.00s") & " in: " & myPath
        '记录Dos中执行Dir命令的耗时
        tms = Timer: ar = Filter(ar, myFile) '然后开始按指定关键词进行筛选。可筛选文件名或文件类型
        Application.StatusBar = Format(Timer - tms, "0.00s") & " Find " & UBound(ar) + IIf(myFile = "", 0, 1) & " Files " & s
        '在Excel状态栏上显示执行结果以及耗时
    End With
    [a:a] = "": If UBound(ar) > -1 Then [a2].Resize(1 + UBound(ar)) = WorksheetFunction.Transpose(ar)
    '清空A列，然后输出结果
End Sub
```

呵呵，Dos命令不仅简洁，而且高效。  

追加更正：提去文件个数统计 提取文件结果的数组ar是下标 0开始的1维数组，元素个数应该=UBound(ar)+1 【此处修正+1为ar(0)】 但实际未产生筛选时的文件结果数=UBound(ar) 无需+1 【因为Dos提取时Dir最后1个""也在结果之中】 而当指定筛选参数myFile不为空时，即产生实际筛选以后的数组ar中会排除最后的那个"",所以筛选后的统计文件结果数=UBound(ar) + 1

关于Dos中Dir命令的开关问题：

【提取文档】
.Exec("cmd /c dir /a-d /b " ………Dir返回指定文件夹下【不包括子文件夹】的所有文档名（不含文件夹）
.Exec("cmd /c dir /a-d /b /s " ………Dir返回指定文件夹下【包括子文件夹】在内的所有文档名（不含文件夹）

其中， /s 即 是否包含 SubFolder的意思
而 /a-d 是文件对象中排除文件夹目录(-d)只剩下文档的意思。

【提取文件夹】
.Exec("cmd /c dir /a-a /b " ………Dir返回指定文件夹下【不包括子文件夹】内的所有子文件夹名（不含文档）
.Exec("cmd /c dir /a-a /b /s " ………Dir返回指定文件夹下【包括子文件夹】内的所有子文件夹名（不含文档）
而 /a-a 是文件对象中排除文档(-a)只剩下文件夹目录的意思。

【提取文档和文件夹】
.Exec("cmd /c dir /b " ………Dir返回指定文件夹下【不包括子文件夹】的所有【文档名】和【文件夹名】
.Exec("cmd /c dir /b /s " ………Dir返回指定文件夹下【包括子文件夹】的所有【文档名】和【文件夹名】


呵呵，以上6种的开关组合就足够了。
补充：Dos Dir开关的帮助文件：

显示目录中的文件和子目录列表。

DIR [drive:][path][filename] [/A[[:]attributes]] [/B] [/C] [/D] [/L] [/N]
 [/O[[:]sortorder]] [/P] [/Q] [/S] [/T[[:]timefield]] [/W] [/X] [/4]

 [drive:][path][filename]
         指定要列出的驱动器、目录和/或文件。

 /A      显示具有指定属性的文件。
 attributes  D 目录          R 只读文件
          H 隐藏文件        A 准备存档的文件
          S 系统文件       - 表示“否”的前缀
/B      使用空格式(没有标题信息或摘要)。
 /C      在文件大小中显示千位数分隔符。这是默认值。用 /-C 来
         停用分隔符显示。
 /D      跟宽式相同，但文件是按栏分类列出的。
 /L      用小写。
 /N      新的长列表格式，其中文件名在最右边。
 /O      用分类顺序列出文件。
 sortorder  N 按名称(字母顺序)   S 按大小(从小到大)
          E 按扩展名(字母顺序)  D 按日期/时间(从先到后)
          G 组目录优先       - 颠倒顺序的前缀
 /P      在每个信息屏幕后暂停。
 /Q      显示文件所有者。
 /S      显示指定目录和所有子目录中的文件。
 /T      控制显示或用来分类的时间字符域。
 timefield  C 创建时间
         A 上次访问时间
         W 上次写入的时间
 /W      用宽列表格式。
 /X      显示为非 8dot3 文件名产生的短名称。格式是 /N 的格式，
         短名称插在长名称前面。如果没有短名称，在其位置则
         显示空白。
 /4      用四位数字显示年

可以在 DIRCMD 环境变量中预先设定开关。通过添加前缀 - (破折号)
来替代预先设定的开关。例如，/-W。

前面的Dir代码，是两个Do循环嵌套使用，
一边检查当前文件夹内的子文件夹，一边检查当前文件夹内的文件。


其实，Dir方法也可以这么写代码：
① 检查并列出所有子文件夹
② 然后根据需要遍历所有子文件夹中的文件

即，两个Do循环是分开来的。
但是、第2次的Do循环需要外套For循环遍历所有已知子文件夹。

```vbscript
Sub ListFilesDir()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    
    MsgBox Join(ListAllDir(myPath), vbCr) 'GetAllSubFolder's File
    MsgBox Join(ListAllDir(myPath, 1), vbCr) 'GetThisFolder's File
    
    MsgBox Join(ListAllDir(myPath, -1), vbCr) 'GetThisFolder's SubFolder
    MsgBox Join(ListAllDir(myPath, -2), vbCr) 'GetAllSubFolder
    
    MsgBox Join(ListAllDir(myPath, , "tst"), vbCr) 'GetAllSubFolder's SpecialFile
    MsgBox Join(ListAllDir(myPath, 1, "tst"), vbCr) 'GetThisFolder's SpecialFile
End Sub

Function ListAllDir(myPath$, Optional sb& = 0, Optional SpFile$ = "")
    Dim i&, j&, k&, myFile$
    ReDim fld(0)
    
    fld(0) = myPath
    On Error Resume Next
    Do
        myFile = Dir(fld(i), vbDirectory)
        Do While myFile <> ""
            If myFile <> "." And myFile <> ".." Then
                If (GetAttr(fld(i) & myFile) And vbDirectory) = vbDirectory Then
                    If Err.Number Then Err.Clear Else j = j + 1: ReDim Preserve fld(j): fld(j) = fld(i) & myFile & ""
                End If
            End If
            myFile = Dir
        Loop
        If sb Mod 2 Then Exit Do Else i = i + 1
    Loop Until i > UBound(fld)
    If sb < 0 And Len(SpFile) = 0 Then ListAllDir = fld: Exit Function
    '以上为止，遍历检查并列出指定目标文件夹中、所有的子文件夹。
    
    '以下为遍历已获得的子文件夹数组fld 然后Dir循环检查其中所有的文件
    ReDim file(0)
    For i = 0 To UBound(fld)
        myFile = Dir(fld(i), vbDirectory)
        Do While myFile <> ""
            If myFile <> "." And myFile <> ".." Then
                If Not (GetAttr(fld(i) & myFile) And vbDirectory) = vbDirectory Then
                    If SpFile = "" Then
                        file(k) = myFile: k = k + 1: ReDim Preserve file(k)
                    Else
                        If InStr(myFile, SpFile) Then file(k) = myFile: k = k + 1: ReDim Preserve file(k)
                    End If
                End If
            End If
            myFile = Dir
        Loop
    Next
    ListAllDir = file
End Function

```

一般说，还是第1种两个Do嵌套的方法好……虽然代码中需要同时处理文件夹和文件名，但Do循环比较高效一些。

第2种方法也并非全无是处。
当处理文件为重点时，以第2种方法比较好。

Dos版 加入Dir各种参数以后的完整代码：

```vbscript
Sub ListFilesDos()
    myMode& = Val(InputBox("Search Mode:-3 To 3", "Find File", 0)) '指定Dos Dir的查找开关、返回模式
    '奇数为不含子文件夹、偶数为含子文件夹 / 负数为目录、正数为文档 / >1为文档及目录
    
    If myMode > -3 Then
        myFile$ = InputBox("Part of Filename or Filetype as "".xl""", "Find File", ".xl")
        '输入指定关键字，可以是文件(文档和目录)名称中的任意部分，或指定文件类型如 ".xl"
    
        Set myFolder = CreateObject("Shell.Application").BrowseForFolder(0, "GetFolder", 0)
        If Not myFolder Is Nothing Then myPath$ = myFolder.Items.Item.path Else MsgBox "Folder not Selected": Exit Sub
        '浏览列表指定查找目录
    End If
    tms = Timer
    With CreateObject("Wscript.Shell") 'VBA调用Dos命令
　　cmdStr = Choose(myMode + 4, "/? ", "/a:d /b /s ", "/a:d /b ", "/a:a /b /s ", "/a:a /b ", "/b /s ", "/b ", "/a:a /o:e /o:n /s ", "/a:a /o:e /o:n ", "/a:d /o:e /o:n /s ", "/a:d /o:e /o:n ")
        ar = Split(.exec("cmd /c dir " & cmdStr & Chr(34) & myPath & Chr(34)).StdOut.ReadAll, vbCrLf)
        '指定Dos中Dir命令的开关然后提取结果 为指定文件夹以及所含子文件夹内的所有文件的含路径全名。
        
        s = UBound(ar) & " Files by Search time: " & Format(Timer - tms, " 0.00s") & " in: " & myPath
        Application.StatusBar = " Find " & s: tms = Timer '记录Dos中执行Dir命令的耗时 并在Excel状态栏上显示
        If myFile <> "" Then '如指定了匹配关键字则
            ar = Filter(ar, myFile) '按指定关键词myFile进行筛选。可筛选文件名或文件类型、然后在Excel状态栏上显示结果
            Application.StatusBar = Format(Timer - tms, "0.00s") & " Find " & 1 + UBound(ar) & " Files from " & s
        End If
    End With
    [a:a] = "": If UBound(ar) > -1 Then [a2].Resize(1 + UBound(ar)) = WorksheetFunction.Transpose(ar)
'    清空A列，然后输出结果
End Sub
```

为大家看得清楚明白，把各种开关写成Select形式：

​     Select Case myMode '根据开关模式设置Dos Dir的开关参数
​        Case -3
​          cmdStr = "cmd /c dir /?" '列出Dir各个参数开关的帮助文件
​        Case -2
​          cmdStr = "cmd /c dir /a-a /b /s " & Chr(34) & myPath & Chr(34) '目录不含文档[/a-a]含子文件夹
​        Case -1
​          cmdStr = "cmd /c dir /a-a /b " & Chr(34) & myPath & Chr(34) '目录不含文档[/a-a](不含子文件夹)
​        Case 0
​          cmdStr = "cmd /c dir /a-d /b /s " & Chr(34) & myPath & Chr(34) '文档不含目录[/a-d]含子文件夹
​        Case 1
​          cmdStr = "cmd /c dir /a-d /b " & Chr(34) & myPath & Chr(34) '文档不含目录[/a-d](不含子文件夹)
​        Case 2
​          cmdStr = "cmd /c dir /b /s " & Chr(34) & myPath & Chr(34) '所有文档及目录含子文件夹
​        Case 3
​          cmdStr = "cmd /c dir /a-d /b " & Chr(34) & myPath & Chr(34) '所有文档及目录(不含子文件夹)
​     End Select

但实际代码中用Choose语句简化。

5. FSO 递归方法实现各种指定搜寻的完整代码：

```vbscript

Dim jg(), k&, tms# '因为是递归，所以事先指定存放结果的公用变量数组jg以及计数器k和起始时间tms
Sub ListFilesFso()
    sb& = InputBox("Search Type: AllFiles=0/Files=1/Folder=-1/All Folder=-2", "Find Files", 0) '选定返回模式
    SpFile$ = InputBox("匹配文件名或文件类型", "Find Files", ".xl") '指定匹配要求，留空则匹配全部
    If SpFile Like ".*" Then SpFile = LCase(SpFile) & "*" '如果指定了文件类型则一律转换为大写字母方便比较
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then myPath$ = .SelectedItems(1) Else Exit Sub
    End With
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    
    ReDim jg(65535, 3)
    jg(0, 0) = "Ext": jg(0, 1) = IIf(sb < 0, IIf(Len(SpFile), "Filename", "No"), "Filename")
    jg(0, 2) = "Folder": jg(0, 3) = "Path"
    '定义存放文件名结果的数组jg 、并写入标题
    tms = Timer: k = 0: Call ListAllFso(myPath, sb, SpFile) '调用递归过程检查指定文件夹及其子文件夹
    If sb < 0 And Len(SpFile) = 0 Then Application.StatusBar = "Get " & k & " Folders."
    [a1].CurrentRegion = "": [a1].Resize(k + 1, 4) = jg: [a1].CurrentRegion.AutoFilter Field:=1
    '输出结果到工作表，并启用筛选模式
End Sub

Function ListAllFso(myPath$, Optional sb& = 0, Optional SpFile$ = "") '递归检查子文件夹的过程代码
    Set fld = CreateObject("Scripting.FileSystemObject").GetFolder(myPath)
    On Error Resume Next
    If sb >= 0 Or Len(SpFile) Then '如果模式为0或1、或指定了匹配文件要求，则遍历各个文件
        For Each f In fld.Files '用FSO方法遍历文件.Files
            t = False '匹配状态初始化
            n = InStrRev(f.Name, "."): fnm = Left(f.Name, n - 1): x = LCase(Mid(f.Name, n))
            If Err.Number Then Err.Clear
            
            If SpFile = " " Then 'Space 如果匹配要求为空则匹配全部
                t = True
            ElseIf SpFile Like ".*" Then '如果匹配要求为文件类型则
                If x Like SpFile Then t = True '当文件符合文件类型要求时匹配，否则不匹配
            Else '否则为需要匹配文件名称中的一部分
                If InStr(fnm, SpFile) Then t = True '如果匹配则状态为True
            End If
            If t Then k = k + 1: jg(k, 0) = x: jg(k, 1) = "'" & fnm: jg(k, 2) = fld.Name: jg(k, 3) = fld.Path
        Next
        Application.StatusBar = Format(Timer - tms, "0.0s") & " Get " & k & " Files , Searching in Folder ... " & fld.Path
    End If
    
    For Each fd In fld.SubFolders '然后遍历检查所有子文件夹.SubFolders
        If sb < 0 And Len(SpFile) = 0 Then k = k + 1: jg(k, 0) = "fld": jg(k, 1) = k: jg(k, 2) = fd.Name: jg(k, 3) = fld.Path
        If sb Mod 2 = 0 Then Call ListAllFso(fd.Path, sb, SpFile)
    Next
End Function

```

6. 定义的写法——补课

关于变量类型缩写的快速记忆：

! = Single 单精度小数……因为 ! 笔画只是1竖单笔画，所以记住为【单精度】
\# = Double 双精度小数 …因为 # 笔画是2横2竖，所以记住为【双精度】
@ = Currency 货币型4位小数 …现实中大家也常用@符号代表价格、单价，所以记住为【货币型小数】
$ = String 文本字符串 …因为 String第1个字母是 S 所以记住为【美元s=String 文本字符串】

% = Integer 整数 ……因为 % 是百分比符号我们把它联想为较少的整数【整型数值】
& = Long 整数 ……因为 & 可以看做是Long首字母L的花体字 所以记住为【长整型数值】

呵呵，这样稍稍动脑筋记忆一下，以后就可以简单使用了。
比如这样子：
Dim i&, j&, k&, l&, l1&, l2&, m&, n&, s$, w1$, w2$

如果很正规地写，成为： 
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim l1 As Long
  Dim l2 As Long
  Dim m As Long
  Dim n As Long

  Dim s As String
  Dim w1 As String
  Dim w2 As String

这样就会很长。

或者写在一起时，横向会很长也不方便
  Dim i As Long, j As Long, k As Long, l As Long, l1 As Long, l2 As Long, m As Long, n As Long
  Dim s As String, w1 As String, w2 As String

…………
以上只是个人习惯而已。

但是，新手千万不要这样子：
  Dim i, j, k, l, l1, l2, m, n As Long
  Dim s, w1, w2 As String


这样做，只有最后一个蓝色的变量被正确定义了变量类型，
其它的都会被作为Variant变量使用……或许不影响使用，但至少违背了作者的初衷。所以不好。

如果需要操作文件以及文件内的各个工作表Sheet，那么当然首先要打开该文件。

```vbscript
Function ListFiles(myPath$)
   Set Fso = CreateObject("Scripting.FileSystemObject")
   For Each f In Fso.GetFolder(myPath).Files
      Workbooks.Open (f) '打开文件
      For Each sh In ActiveWorkbook.Sheets '遍历该文件的所有工作表
        sh.Activate '激活工作表
        
      Next
      ActiveWorkbook.Close '关闭文件
   Next
End Function

```



【Dir 使用方法】

myPath = "c:\"  '首先设定目标文件夹，注意末尾必须是【\】文件夹符号。

myFile = Dir(myPath, vbDirectory)  '第一次使用Dir函数时，必须完整输入路径和检索要求。
                                 ' 如果直接使用Dir不带参数则会报错。

Do While myFile <> ""  '开始Do不定循环、直至在本文件夹内没有找到文档/文件夹而返回空白时停止。

  If myFile <> "." And myFile <> ".." Then
     '此If判断为忽略 当前文件夹"."以及忽略上级文件夹".."

​     If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
​        '接下来的If判断是：通过二进制的位比较计算结果= vbDirectory 来判断这是一个文件夹。
​        Debug.Print myFile      '判断为文件夹时的处理
​     Else '否则为文档
​        Debug.Print myFile     '判断为文档时的处理。
​     End If
  End If
  myFile = Dir   '继续调用【不带路径参数的Dir函数】 这样就能得到下一个搜寻结果。
Loop

## 四、本人的代码

1. 之前XX项目折腾从excel到word，批量替换多个变量，遍历文件夹，但是开始自己还不会，只好在代码内复制粘贴，主要是为了满足功能，就这已经帮我们团队节省了80%以上的人力，先把丑陋的代码放在这里，命名v1.1：

```vbscript
Private Sub CommandButton1_Click()

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
            For j = 1 To 48
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
    
   For i = 3 To 最后行号
      导出文件名 = "本期区块链应收款清单：4份，总行公章+法人章+骑缝"
      FileCopy 当前路径 & "\本期区块链应收款清单：4份，总行公章+法人章+骑缝.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      导出路径文件名 = 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      With Word对象
        .Documents.Open 导出路径文件名
        .Visible = False
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 48
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
       
   For i = 3 To 最后行号
      导出文件名 = "风险说明书：6份，总行公章、法人章+骑缝"
      FileCopy 当前路径 & "\风险说明书：6份，总行公章、法人章+骑缝.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      导出路径文件名 = 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      With Word对象
        .Documents.Open 导出路径文件名
        .Visible = False
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 48
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
       
   For i = 3 To 最后行号
      导出文件名 = "信托说明书"
      FileCopy 当前路径 & "\信托说明书.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      导出路径文件名 = 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      With Word对象
        .Documents.Open 导出路径文件名
        .Visible = False
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 48
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
       
   For i = 3 To 最后行号
        导出文件名 = "资产服务协议：5份，总行公章、法人章+骑缝"
      FileCopy 当前路径 & "\资产服务协议：5份，总行公章、法人章+骑缝.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      导出路径文件名 = 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
      With Word对象
        .Documents.Open 导出路径文件名
        .Visible = False
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        With .Selection.Find '填写文字数据
            For j = 1 To 48
             Str1 = "数据" & Format(j, "000")
             Str2 = Sheets("数据").Cells(i, j + 1)
             .Text = Str1
             .Replacement.Text = Str2 '替换字符串
             .Execute Replace:=wdReplaceAll '全部替换
            Next j
        End With
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
       
   For i = 3 To 最后行号
      导出文件名 = "资产交割确认书：5份，总行公章+法人章"
      FileCopy 当前路径 & "\资产交割确认书：5份，总行公章+法人章.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
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
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
       
   For i = 3 To 最后行号
      导出文件名 = "交割确认函：5份，总行公章+法人章"
      FileCopy 当前路径 & "\交割确认函：5份，总行公章+法人章.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
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
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i

    
   For i = 3 To 最后行号
      导出文件名 = "信托收益权转让登记表：3份，总行公章"
      FileCopy 当前路径 & "\信托收益权转让登记表：3份，总行公章.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
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
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
  
   For i = 3 To 最后行号
      导出文件名 = "信托资金保管合同"
      FileCopy 当前路径 & "\信托资金保管合同.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
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
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
   For i = 3 To 最后行号
      导出文件名 = "财产权受益权转让合同：5份，总行公章、法人章_骑缝"
      FileCopy 当前路径 & "\财产权受益权转让合同：5份，总行公章、法人章_骑缝.doc", 当前路径 & "\" & 导出文件名 & "(" & Sheets("数据").Range("B" & i) & ").doc"
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
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
        Str1 = "数据001"
        Str2 = Sheets("数据2").Cells(i, 2)
        .Selection.HomeKey Unit:=wdStory '光标置于文件首
        If .Selection.Find.Execute(Str1) Then '查找到指定字符串
           .Selection.Font.Color = wdColorAutomatic '字符为自动颜色
           .Selection.Text = Str2 '替换字符串
        End If
      End With
      Word对象.Documents.Save
      Word对象.Quit
      Set Word对象 = Nothing
   Next i
   
   If 判断 = 0 Then
      i = MsgBox("已输出到 Word 文件！", 0 + 48 + 256 + 0, "提示：")
   End If
End Sub
Private Sub CommandButton输出通知到Word文件_Click()

End Sub


```

2. 恰好这次有个客户反馈word模板目录有问题，于是经过几天李笑来老师的自学之书的指点，加上对python的进一步理解，对编程算是有了一点点进步，结合前面几位牛人的代码，我算是完成了代码的升级，所以命名2.0，代码如下：

```vbscript
Private Sub CommandButton1_Click()

    Dim Str1, Str2
    Dim j
    Dim wdApp As Word.Application
    
    Set myFolder = CreateObject("Shell.Application").BrowseForFolder(0, "GetFolder", 0)
    If Not myFolder Is Nothing Then myPath$ = myFolder.Items.Item.Path Else MsgBox "Folder not Selected": Exit Sub
    If Right(myPath, 1) <> "" Then myPath = myPath & ""
    '同样返回的是选中目标文件夹的绝对路径,但除了本地C盘、D盘会以"C:"形式返回外，其余路径无""需要添加
    判断 = 0
    Application.ScreenUpdating = False
    
    Set fld = CreateObject("Scripting.FileSystemObject").GetFolder(myPath) '设置FSO实例
       
    For Each f In fld.files '遍历文件夹
     '以下代码实现每个文件的查找替换
        Set wdApp = New Word.Application
        If Right(f.Name, 4) = ".doc" And InStr(f.Name, "$") = 0 Then
        With wdApp
            .Documents.Open (f) '打开文件，进行操作
            .Visible = False
            .Selection.HomeKey Unit:=wdStory '光标置于文件首
            With .Selection.Find '填写文字数据
                For j = 1 To 49
                    Str1 = "数据" & Format(j, "000")
                    Str2 = Sheets("数据").Cells(3, j + 1)
                    .Text = Str1
                    .Replacement.Text = Str2 '替换字符串
                    .Execute Replace:=wdReplaceAll '全部替换
                Next j
            End With
        
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置位置在页眉
                Str1 = "数据001"
                Str2 = Sheets("数据2").Cells(3, 2)
                .Selection.HomeKey Unit:=wdStory '光标置于文件首
      
                 If .Selection.Find.Execute(Str1) Then '查找到指定字符串
                    .Selection.Text = Str2 '替换字符串
                 End If
             
            .Documents.Save
            .Quit
        End With
        End If
        Set wdApp = Nothing
    Next
    Application.ScreenUpdating = True
    If 判断 = 0 Then
        i = MsgBox("已输出到 Word 文件！", 0 + 49 + 256 + 0, "提示：")
    End If
End Sub
Private Sub CommandButton输出通知到Word文件_Click()

End Sub
```

3. **继续优化** 代码，实现通用的批量查找替换

其实很容易，就把上面代码第26行修改为如下即可：

```vbscript
Str1 = Sheets("数据").Cells(2, j + 1)
```

修改后的使用方法：

>1. 选择需要批量替换的文件夹
>2. 注意指定文档为.doc的文档，这个可以通过修改代码来改为所有文件
>3. 在excel输入项里，第二行填写旧信息，第三行填写需要替换的信息
>4. 点击按钮即可，设置最多的是49项，可以通过修改代码增加或减少

