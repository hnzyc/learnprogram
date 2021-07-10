import xlwings as xw
import win32com.client as win32
import  os


def split_excel(file_path,sheet_name,init_range,col=0):
    """
    读取分配表，并根据设置拆分
    自动保存到指定文件夹
    """     

    #打开excel，读取分配总表数据
    app = xw.App(visible = False, add_book = False) # 启动Excel程序
    workbook = app.books.open(file_path) # 打开分配表来源工作簿
    worksheet = workbook.sheets[sheet_name] # 选中要拆分的数据所在sheet
    value = worksheet.range(init_range).expand('table').value # 读取需要拆分的所有数据，B37单元格是数据的起始单元格，根据实际情况修改
    
    data = dict() # 创建一个空字典，用于按要求分类储存数据
    # 拆分分配数据
    for i in range(len(value)): # 按行遍历工作表数据
        client_name = value[i][col] # 获取分类依据的数据，两个[]的数字代表行和列，从0开始
        if client_name not in data: #判断字典中是否已有当前客户名称
            data[client_name] = [] # 若没有，就创建一个当前客户名称的空列表，用来储存当前行数据
        data[client_name].append(value[i]) # 将当前行数据追加到当前行客户名称对应的列表中
    
    app.quit()
    return data
    
def save_excel(data,save_path,title_range):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    title = data[title_range]
    app = xw.App(visible = False, add_book = False) # 启动Excel程序
    for key,value in data.items(): # 按客户名称遍历分类后的数据
        workbook = app.books.add() # 新建工作簿
        
        if key != title_range:            
            worksheet = workbook.sheets.add(key) # 在新工作簿中新建工作表，并且命名为客户名称 
            worksheet['A1'].value = title # 将要拆分的列标题复制到第一行,根据实际情况修改           
            worksheet['A2'].value = value # 将当前客户名称下的数据复制到新建的工作表
            worksheet.autofit() #自动调整格式  
            worksheet.range('B2').number_format = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ' #根据目前的分配表分别设置
            worksheet.range('D2:G2').number_format = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ' # 若后续有变化
            worksheet.range('C2').number_format = '0.00%'                                               # 需手动单独调整


    #保存为新工作簿，命名为客户名称，这里也可以加上当前日期或其他内容，根据实际需要调整    
        workbook.save(os.path.join(save_path,'{}.xlsx'.format(key)))
        print('{}已拆分完毕'.format(key))
    app.quit()  

def send_mail(file_path,mail_sub,mail_body,attach_report,save_path,mail_to,sheet_name,col=0):
    """读取投资者信息文件，例如D:\examples\投资者信息统计.xlsx
    投资者分类依据所在列col,默认为1
    投资者邮箱所在列mail_to，注意第一列设为0，第2列设为1，依此类推，默认是4；
    如果是最后一列，可以设置为-1，倒数第二列，设置为-2，依此类推
    邮件主题mail_sub，例如'XX项目分配邮件通知'
    邮件内容mail_body，默认使用html格式，见示例
    附件1为分配表
    附件2为受托报告
    """
    app = xw.App(visible = False, add_book = False) # 启动Excel程序
    #打开投资者信息数据
    workbook = app.books.open(file_path) # 打开投资者信息工作簿
    worksheet = workbook.sheets[sheet_name] # 选中提取的数据所在sheet
    value = worksheet.range('A1').expand('table').value # 读取需要投资者信息数据
     # 匹配投资者收件箱，发送邮件
    outlook = win32.Dispatch('Outlook.Application') #初始化outlook
    for key in os.listdir(save_path): # 遍历拆分分配表文件夹内文件名
        inv_name = os.path.splitext(key)[0]
        for i in range(len(value)): # 对投资者信息遍历           
            if value[i][col] == inv_name: # 投资者名称如果和分配表文件名一致
                mail = outlook.CreateItem(0) # 固定写法 
                # 发送邮件
                mail.To = value[i][-1] # 邮箱收件人名称就是邮箱
                mail.Subject = mail_sub
                mail.HtmlBody = mail_body

                path1 = os.path.join(save_path,'{}.xlsx'.format(inv_name))
                path2 = attach_report

                mail.Attachments.Add(path1) # 附件1，分配表
                if path2:
                    mail.Attachments.Add(path2) # 附件2，受托报告

                print('{}邮件已发送'.format(inv_name))
                mail.Send() # 发送邮件    
    app.quit()      

def save_mail(file_path,mail_sub,mail_body,attach_report,save_path,mail_to,sheet_name,col=0):
    """
    读取投资者信息文件，例如D:\examples\投资者信息统计.xlsx
    投资者分类依据所在列col,默认为1
    投资者邮箱所在列mail_to，注意第一列设为0，第2列设为1，依此类推，默认是4；
    如果是最后一列，可以设置为-1，倒数第二列，设置为-2，依此类推
    邮件主题mail_sub，例如'XX项目分配邮件通知'
    邮件内容mail_body，默认使用html格式，见示例
    附件1为分配表,默认配置
    附件2为受托报告，默认配置，也可不设置，则不附带附件2
    """
    app = xw.App(visible = False, add_book = False) # 启动Excel程序
    #打开投资者信息数据
    workbook = app.books.open(file_path) # 打开投资者信息工作簿
    worksheet = workbook.sheets[sheet_name] # 选中提取的数据所在sheet
    value = worksheet.range('A1').expand('table').value # 读取需要投资者信息数据
     # 匹配投资者收件箱，发送邮件
    outlook = win32.Dispatch('Outlook.Application') #初始化outlook
    for key in os.listdir(save_path): # 遍历拆分分配表文件夹内文件名
        inv_name = os.path.splitext(key)[0]
        for i in range(len(value)): # 对投资者信息遍历  
            mail_failure = []         
            if value[i][col] == inv_name: # 投资者名称如果和分配表文件名一致
                mail = outlook.CreateItem(0) # 固定写法 
                # 发送邮件
                mail.To = value[i][mail_to] # 邮箱收件人名称就是邮箱
                mail.Subject = mail_sub
                mail.HtmlBody = mail_body

                path1 = os.path.join(save_path,'{}.xlsx'.format(inv_name))
                path2 = attach_report

                mail.Attachments.Add(path1) # 附件1，分配表
                if path2:
                    mail.Attachments.Add(path2) # 附件2，受托报告

                print('{}邮件已保存，请查看草稿箱'.format(inv_name))
                mail.Save() #保存草稿  
    app.quit()        