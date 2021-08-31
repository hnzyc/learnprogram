import os
import work
from config import global_config #编写config模块，设置全局配置信息

 # 转让登记表文件夹路径
path = global_config.get('register_path','path')
investors_info = os.path.join(path,global_config.get('register_path','investors_info'))


if __name__ == '__main__':
    print('如转让登记表为docx，则逐个转换，请稍等……')
    # work.word2pdf(path) # 转让登记表全部转换为pdf
    print('现在逐个提取转入登记表投资者信息，请稍候……')
    data = work.cust_info_list(path) #转让登记表统计信息
    data.to_excel(investors_info) #保存转让登记表信息为excel文件
    print('已全部提取完成，请打开转入登记表所在文件夹查看')
  