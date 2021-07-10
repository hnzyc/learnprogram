import xlwings as xw
import win32com.client as win32
import  os
import work
from config import global_config #编写config模块，设置全局配置信息

# 首先定义所有需要用到的文件路径
file_path = global_config.get('excel_split','file_path') # 读取分配表文件所在路径
sheet_name = global_config.get('excel_split','sheet_name') # 需要拆分的sheet名称，根据实际情况确定
init_range = global_config.getRaw('excel_split','init_range') 
col = global_config.getint('excel_split','col')
title_range = global_config.get('excel_split','title_range')
save_path = global_config.get('excel_split','save_path')

if __name__ == '__main__':
    
    print('现在开始拆分分配表数据，请稍候……')
    print('***'*20)
    data = work.split_excel(file_path,sheet_name,init_range,col)
    work.save_excel(data,save_path,title_range)
    print('*'*60)