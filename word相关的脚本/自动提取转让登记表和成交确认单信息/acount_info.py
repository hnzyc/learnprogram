import os
import work
from config import global_config #编写config模块，设置全局配置信息

# 结算单文件夹路径
path1 =  global_config.get('account_path','path1')
# 结算单保存的文件名
investors_info1 = os.path.join(path1,global_config.get('account_path','investors_info1'))


if __name__ == '__main__':
      
    print('正在逐个提取结算单信息，请稍候……')
    data1 = work.cust_info_list2(path1) # 结算单信息统计
    data1.to_excel(investors_info1) # 保存结算单信息到excel表
    print('已全部提取完成，请打开转入结算单所在文件夹查看')