## Python 办公小助手：读取 PDF 中表格并重命名

原创： TED [TEDxPY](javascript:void(0);) *11月1日*

日常工作中，我们或多或少都会接触到 Excel 表格、Word 文档和 PDF 文件。偶尔来个处理文件的任务，几个快捷键操作一下——搞定！但是，偏偏有些烦人的工作，操作繁琐且数据复杂，更要命的是耗时间，吭哧吭哧一下午却难出几个成果。



此时如果我们掌握些 Python 编程的技巧，整理下文件处理的流程通过编码来实现，不仅省时省力省心，还可以精进编码技术。今天我们就通过一个 PDF 处理的实例来演示下 Python 助力办公的过程。



上周朋友提了个 PDF 处理的问题，要求如下：



![img](https://mmbiz.qpic.cn/mmbiz_jpg/ib9fOiakpb83FWJ9JD6nN5al5gswhJOKJo0kkDYanW7AATeXa498t2I5icdP9kll43n8A39KhG4CraicDOicygPGIZA/640?wx_fmt=jpeg&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



大致整理下，这问题和把大象装冰箱一样要分三步：



1. 读取 PDF 中的表格内容
2. 在表格内容中提取特定数据
3. 以特定数据对文件重命名



此时面向 Python 默默许愿：要是 Python 中有现成的模块可以直接读取 PDF 中的表格就好了！



心愿达成！确实有个 tabula 模块可以直接解析 PDF 中的表格：

> tabula-py is a simple Python wrapper of tabula-java, which can read table of PDF. You can read tables from PDF and convert into pandas's DataFrame. tabula-py also enables you to convert a PDF file into CSV/TSV/JSON file.
>
> https://pypi.org/project/tabula-py/

如上所述， tabula-py 是 tabula-java 的一个封装模块，可以将 PDF 中的表格数据转化为 pandas 的 DataFrame 格式。



注意，安装 tabula-py 时命令是 pip install tabula-py，但导入时是 import tabula。



此外，该模块由于是对 tabula-java 的封装依赖 java，需要安装 java 才能正常调用。并且由最终转化得到的数据格式也可以看出，此模块也依赖 pandas 和 numpy，需要自行导入。



- 

```
详细链接：https://pypi.org/project/tabula-py/
```



安装好 tabula-py，我们也准备一份 PDF 文件（demo.pdf）用于演示代码：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvMmcbdvdd7JrgFcpAS41CVZV5ITBy8Ng1ib7A1ojPHxqnbHmnqLlCqfg/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



现在任务更清晰了：读取 demo.pdf 文件中的 “批号（款号）”数据：

- 

```
"批号（款号）"："DRDY173131441HHDKD QWOEP23"
```

最终将这一串批号数据当作名字给 PDF 重命名，生成 DRDY...EP23.pdf 文件。



------



如果你能坚持看到这里，我准备向你推荐下 jupyter notebook。因为它可以按代码块执行，上下代码块之间变量可以共用，同时会直接显示代码块运行结果。拿它用来做代码及运行结果展示非常好用——下文记录的过程就是通过它运行代码截图所得。



\1. 首先，导入 tabula，使用其函数读取 PDF 中的表格数据：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRv2ZECvpN5iaMicMKTHvSF1GwnrdOqc3np0NsZ8pDRhgXpicVYoFsPQ0mow/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



由所得结果大致可以看出，我们想要的批号数据是在第二列。



\2. 之前提到读到的 PDF 表格数据是 DataFrame 格式，可以用 help 函数确认下：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvU7yoCHSn7Wiaq0mwDQlf5RABqzBlgPGsROHTaSgJQwe29xjP6aYAGcQ/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



\3. 由表格数据中提取其每一列的名称：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvmqpdUrB1Pn1F3XxBOI1NseGKMO2AAn4wf996diaZolBB7XLiaQkyZGbw/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



\4. 根据目测分析，批号位于第二列，所以提取第二列名字：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvYUtZ60hy3LHCmVZHGHQ7S6neqwOl9vZ8j5peUPuhb0bQAk0ydUuRPQ/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



\5. 通过 DataFrame["列名称"] 来定位到该列具体数据：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRv9iaV4UdoFZC4r6iak4NtvgPpO6BxLIKSsQHxwrwHw8ianoFJVCVJfUsicQ/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



\6. 通过 for 循环逐一打印此列数据，提取其中“批号”数据：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvNBC3HL8iboohvaR4ATqwtJJmh3HLSsxO120uFXcRsKGacIuWJyTKrrg/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



这里直接采用的是 *"批号" in 字符串* 的语法，倘若数据字符串中含有“批号”二字就会被筛选出，最终我们也如愿拿到了“批号数据”并赋值给 target 变量。



\7. 拿到了“批号”数据，我们只选取字母数字拼接的数据串。接下来采用正则表达式，按照批号数据格式中只包含大写字母、数字以及中间会夹杂空格，制定匹配模式进行匹配提取：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvDia4llJcM58JPhYnoLyadstssZZLbcGIhC4m8PthR1WJBtG61Y0iaj8g/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



最终我们拿到了批号数据串赋值给 result 变量。



\8. 最终我们利用 os 模块将文件夹内的 “demo.pdf” 重命名为 result 所代表的批号数据串.pdf ：



![img](https://mmbiz.qpic.cn/mmbiz_png/ib9fOiakpb83H3XzD6fB6wlOFhKelLBRRvvtRUiaJUHjvUiaviaiaMmYfLwb6CicmNHoLV3vdSvv0ZVRLSHVtCyY7uEFA/640?wx_fmt=png&tp=webp&wxfrom=5&wx_lazy=1&wx_co=1)



注意，这里的 f"{变量}字符串内容" 是格式化字符串的形式。



至此，我们完成了对单份 PDF 处理的完整流程。接下来我们可以多试几份不同 PDF 寻找共同的提取批号数据的规律，将其整理成连贯、简洁的最终版代码：



- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 
- 

```
#!/usr/bin/env python# encoding: utf-8# @Time : 2019-10-24 21:39__author__ = 'Ted'import tabulaimport reimport os# 将提取单一 PDF 文件内批号数据的过程定义成 get_target("pdf名称") 函数，最终函数将数据返回def get_target(filename):    df = tabula.read_pdf(filename)    pattern = r'[A-Z0-9]+[\s]*[A-Z0-9]*'    for item_sub in df[df.columns[1]]:        if "批号" in str(item_sub):            result = re.search(pattern,item_sub).group()            return result    return Falseif __name__=="__main__":    # 获取 PDF 所在文件夹    folder = "test"    # os 模块定位到该文件夹    os.chdir(folder)    # 获取文件夹内文件列表    pdflist = os.listdir()    # 打印该文件列表    print(pdflist)    # 对文件列表 for 循环处理    for item in pdflist:        # 如果该文件名称最后四位是 .pdf 或 .PDF,即我们要找的 PDF 文件        if item[-4:] in [".pdf",".PDF"] :            # 对该文件进行提取批号函数操作，将批号数据赋值给 new_name            new_name = get_target(item)            # 如果不为空，即获取到了批号数据            if new_name:                # 对文件进行重命名操作                os.rename(item,f"{new_name}.pdf")    print("重命名成功！")
```



如果我们有大量 PDF 文件都要提取文件内的批号数据进行重命名，可以将其放到同一个文件夹中，然后只要在最终代码中修改 folder = "文件夹名称"，运行代码等待几秒，便可微微一笑任务搞定了。



以上，感谢阅读～

阅读 230

 在看5

![img](http://wx.qlogo.cn/mmopen/omObnLHphGZCTvqWJtA46ASGBG9JgDH7mibsbPoC98E3GibMUttfo6GD7wibibw6jtB3oBx9RZWnBkpkDz8HqaAIdalWBLkZOxyx/132)

写下你的留言

**精选留言**

-  2

    置顶**TED**

    ![img](http://wx.qlogo.cn/mmopen/omObnLHphGZXuNqqpej9uhEPrOicXfKmOyOdlgq4mVMnJTc66tqXlknwzJAqFLlo55YySHFeUK0ENsBQ3uzrt97VgZfKoNxibm/96)

    

    代码、展示用的 PDF 以及 Jupyter notebook 文件已上传 GitHub，链接如下： https://github.com/pengfexue2/pdf_dataframe_rename.git

-  2

    **聆听，逝去的流年**

    ![img](http://wx.qlogo.cn/mmopen/anblvjPKYbNiaIHkYia4ra0gEbrg4wnmeAGhl7cgr4NbLtc0nL4jwOW1twUhbB5MzSjIDqI5zzMpdM0mSMYHC8dV8uxLl5xqKb/96)

    

    解决了我的一个大烦恼

     2

    作者

    本文的甲方

-  

    **GS**

    ![img](http://wx.qlogo.cn/mmopen/anblvjPKYbP4xHBqzE3y2tnwHfykaDg8TLPianHPkFNiag2JYEJdmPkY1mia8XFNR76namNSY2UVTdzBBCBtfxPSQ/96)

    

    大佬牛逼啊

     

    作者

     花式赞美 感谢支持

-  

    **-**

    ![img](http://wx.qlogo.cn/mmopen/ScZjzbdmZkq1ctpFsEd8Y48Vrv9bCq7TX5DCvwdR3DRTajYREU4ic3cwIZWQMIkV58vCJvcyYEczfZcpwxEB6Q4QkxM66jfAr/96)

    

    牛逼！

     

    作者

     没有没有

-  

    **℡小程か张**

    ![img](http://wx.qlogo.cn/mmopen/iaeVvKGrB1ZOQBNkaMTpyVYesfX0z4GMMOQNqoicOGiczezic4NWS19SkicGzicxQFOu4IHLByHawicSQE9WlmNjFjGseTynp0Yu0ur/96)

    

    牛批

     

    作者

    捧场张

-  

    **Leo Debugging🇨🇳🇭🇰**

    ![img](http://wx.qlogo.cn/mmopen/PiajxSqBRaEKj96f9y2Jnjp5WezfawQXqOUIqfucrAsxT5my32MSibLE4opqHc8hGcnSB2iacP0PgY9qayzfMM60A/96)

    

    ୧( ⁼̴̶̤̀ω⁼̴̶̤́ )૭

     

    作者

    忠实读者