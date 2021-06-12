import xlwings as xw
import win32com.client as win32
import  os
# 首先定义所有需要用到的文件路径
file_path = '台账模板.xlsx' # 分配表文件所在路径，由于我将文件都拷贝到当前文件夹，所以直接写文件名即可
sheet_name = '模板' # 需要拆分的sheet名称，根据实际情况确定
file_path_investors = r'D:\learnpython\excel_examples\投资者信息统计.xlsx' # 投资者信息所在路径
report = r'D:\learnpython\excel_examples\XX项目受托报告202106.pdf'


#打开excel，读取分配总表数据
app = xw.App(visible = False, add_book = False) # 启动Excel程序
workbook = app.books.open(file_path) # 打开分配表来源工作簿
worksheet = workbook.sheets[sheet_name] # 选中要拆分的数据所在sheet
value = worksheet.range('B38').expand('table').value # 读取需要拆分的所有数据，B38单元格是数据的起始单元格，根据实际情况修改
data = dict() # 创建一个空字典，用于按要求分类储存数据
#打开投资者信息数据
workbook_1 = app.books.open(file_path_investors) # 打开投资者信息工作簿
worksheet_1 = workbook_1.sheets['Sheet1'] # 选中提取的数据所在sheet
value_1 = worksheet_1.range('A1').expand('table').value # 读取需要投资者信息数据
# 拆分分配数据
for i in range(len(value)): # 按行遍历工作表数据
    client_name = value[i][0] # 获取分类依据的数据，两个[]的数字代表行和列，从0开始
    if client_name not in data: #判断字典中是否已有当前客户名称
        data[client_name] = [] # 若没有，就创建一个当前客户名称的空列表，用来储存当前行数据
    data[client_name].append(value[i]) # 将当前行数据追加到当前行客户名称对应的列表中

for key,value in data.items(): # 按客户名称遍历分类后的数据
    new_workbook = xw.books.add() # 新建工作簿
    new_worksheet = new_workbook.sheets.add(key) # 在新工作簿中新建工作表，并且命名为客户名称
    new_worksheet['A1'].value = worksheet['B37:H37'].value # 将要拆分的列标题复制到第一行,根据实际情况修改
    new_worksheet['A2'].value = value # 将当前客户名称下的数据复制到新建的工作表
    new_worksheet.autofit() #自动调整格式    
    # 保存为新工作簿，命名为客户名称，这里也可以加上当前日期或其他内容，根据实际需要调整    
    new_workbook.save(r'D:\learnpython\excel_examples\分配表\{}.xlsx'.format(key)) 
    print('{}已拆分完毕'.format(key))

    # 匹配投资者收件箱，发送邮件
    for j in range(len(value_1)): # 对投资者信息遍历
        if value_1[j][0] in key: # 投资者名称如果和分配表文件名一致
            # 发送邮件
            outlook = win32.Dispatch('Outlook.Application') #初始化outlook
            mail = outlook.CreateItem(0) # 固定写法
            
            mail.To = value_1[j][4] # 邮箱收件人名称就是邮箱
            mail.CC = ''
            mail.Subject = 'XX项目分配邮件通知'
            mail.HtmlBody = """
            <p><b>尊敬的投资者：</b></p>
            <p>您好，请查收贵司本次投资分配明细表，以及受托报告，详见附件。</p>
            <p>如有任何问题，请及时联系我们。</p>
            <font color="gray" size="2">赵运超</font><br>
            <font color="gray" size="2">手机：18924588577/13902319827</font><br/>
            """
            path1 = r'D:\learnpython\excel_examples\分配表\{}.xlsx'.format(key)
            path2 = report
             
            mail.Attachments.Add(path1) # 附件1，分配表
            mail.Attachments.Add(path2) # 附件2，受托报告

        #     mail.Send()#发送# 应先在邮箱通讯录中建立邮件组，直接输入邮件组名字可以发送
            mail.Save() #保存草稿
            print('{}邮件已发送'.format(key))
#             print('找不到投资者{}，请确保分配表中的投资者名称与信息表中完全一致'.format(key))
    
app.quit() # 退出Excel程序

print("*"*27)
print('未提示已发送的，请确保分配表中的投资者名称与信息表中完全一致')
print("已全部拆分完毕,并完成邮件发送")
print("*"*27)