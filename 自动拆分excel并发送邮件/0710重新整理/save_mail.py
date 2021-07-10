import xlwings as xw
import win32com.client as win32
import  os
import work
from config import global_config #编写config模块，设置全局配置信息

# 首先定义所有需要用到的文件路径
file_path = global_config.get('send_mail','file_path') # 读取投资者信息表文件所在路径
sheet_name = global_config.get('send_mail','sheet_name')
col = global_config.getint('send_mail','col')
mail_to = global_config.getint('send_mail','mail_to')
mail_sub = global_config.get('send_mail','mail_sub')
mail_body = global_config.get('send_mail','mail_body')
attach_report = global_config.get('send_mail','attach_report')
save_path = global_config.get('excel_split','save_path')

if __name__ == '__main__':
    print('现在开始保存邮件，请稍候……')
    print('***'*20)
    work.save_mail(file_path,mail_sub,mail_body,attach_report,save_path,mail_to,sheet_name,col)
