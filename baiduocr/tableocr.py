# coding=utf-8
# 作者：赵运超
# 2021-08-06
# 批量修改文件名称
# 参考来源：https://www.cnblogs.com/mrlayfolk/p/12630128.html

import os
from aip import AipOcr #调取百度AI接口所需库
import requests
import time
import glob #文件名模式匹配，不用遍历整个目录判断每个文件是不是符合
import sys
import tkinter as tk
from tkinter import filedialog # 图形化对话框
from PIL import Image   #处理图片的库

 
# 新建AipOcr
""" 你的 APPID AK SK """
APP_ID = '24597675'
API_KEY = 'non3MelzmCk4uWeRKLojMyTi'
SECRET_KEY = 'fsDb2Cp13mcYb6kUfv5UUM0L2oEe3e2f'
 
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

#调整原始图片
def convertimg(picfile, outdir):

    img = Image.open(picfile)
    width, height = img.size
    while (width * height > 4000000):  # 该数值压缩后的图片大约两百多k
        width = width // 2   # '//'表示整数除法，例5//2=2
        height = height // 2
    new_img = img.resize((width, height), Image.BILINEAR)   #重置图片大小和质量
    '''
    Image.NEAREST ：低质量；Image.BILINEAR：双线性；Image.BICUBIC ：三次样条插值；Image.ANTIALIAS：高质量
    '''
    new_img.save(os.path.join(outdir, os.path.basename(picfile))) #新图片保存在outdir即tmp目录下。os.path.basename(path)表示返回文件名。tmp/pic.jpg

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

"""文件下载"""
def file_download(url, file_path):
    res = requests.get(url)
    with open(file_path, 'wb') as f:
        f.write(res.content)

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    data_dir = filedialog.askdirectory(title = '请选择图片文件夹') + '/'
    result_dir = filedialog.askdirectory(title = '请选择输出文件夹') + '/'
    num = 0
    for name in os.listdir(data_dir):
        print('{0}:{1}正在处理'.format(num+1, name.split('.')[0]))
        image = get_file_content(os.path.join(data_dir, name))
        res = client.tableRecognitionAsync(image)
        # print(res)
        if 'error_code' in res.keys():
            print('Error! error_code:', res['error_code'])
            sys.exit()
        req_id = res['result'][0]['request_id']#获取识别ID号

        for count in range(1,20): #OCR识别也需要一定时间，设定10秒内每隔1秒查询一次
            res = client.getTableRecognitionResult(req_id) #通过ID获取表格文件XLS地址
            print(res['result']['ret_msg'])
            if res['result']['ret_msg'] == '已完成':
                break #云端处理完毕，成功获取表格文件下载地址，跳出循环
            else:
                time.sleep(3)

        url = res['result']['result_data']
        # print(url)
        xls_name = name.split('.')[0] + '.xls'
        file_download(url, os.path.join(result_dir, xls_name))
        num += 1
        print('{0}:{1}下载完成。'.format(num, xls_name))
        time.sleep(1)