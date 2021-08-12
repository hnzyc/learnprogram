# coding=utf-8
# 作者：赵运超
# 2021-08-07
# 批量修改文件名称
# 参考来源：https://github.com/Baidu-AIP/QuickStart/blob/master/OCR/main.py
# 可以实现简单对话框选择路径，也可以使用输入路径方式来实现

import sys
import json
import base64
import os
# import glob
import tkinter as tk
from tkinter import filedialog # 图形化对话框


# 保证兼容python2以及python3
IS_PY3 = sys.version_info.major == 3
if IS_PY3:
    from urllib.request import urlopen
    from urllib.request import Request
    from urllib.error import URLError
    from urllib.parse import urlencode
    from urllib.parse import quote_plus
else:
    import urllib2
    from urllib import quote_plus
    from urllib2 import urlopen
    from urllib2 import Request
    from urllib2 import URLError
    from urllib import urlencode

# 防止https证书校验不正确
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

API_KEY = 'non3MelzmCk4uWeRKLojMyTi'

SECRET_KEY = 'fsDb2Cp13mcYb6kUfv5UUM0L2oEe3e2f'


OCR_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"


"""  TOKEN start """
TOKEN_URL = 'https://aip.baidubce.com/oauth/2.0/token'


"""
    获取token
"""
def fetch_token():
    params = {'grant_type': 'client_credentials',
              'client_id': API_KEY,
              'client_secret': SECRET_KEY}
    post_data = urlencode(params)
    if (IS_PY3):
        post_data = post_data.encode('utf-8')
    req = Request(TOKEN_URL, post_data)
    try:
        f = urlopen(req, timeout=5)
        result_str = f.read()
    except URLError as err:
        print(err)
    if (IS_PY3):
        result_str = result_str.decode()


    result = json.loads(result_str)

    if ('access_token' in result.keys() and 'scope' in result.keys()):
        if not 'brain_all_scope' in result['scope'].split(' '):
            print ('please ensure has check the  ability')
            exit()
        return result['access_token']
    else:
        print ('please overwrite the correct API_KEY and SECRET_KEY')
        exit()

"""
    读取文件
"""
def read_file(image_path):
    f = None
    try:
        f = open(image_path, 'rb')
        return f.read()
    except:
        print('read image file fail')
        return None
    finally:
        if f:
            f.close()


"""
    调用远程服务
"""
def request(url, data):
    req = Request(url, data.encode('utf-8'))
    has_error = False
    try:
        f = urlopen(req)
        result_str = f.read()
        if (IS_PY3):
            result_str = result_str.decode()
        return result_str
    except  URLError as err:
        print(err)


if __name__ == '__main__':

    # 获取access token
    token = fetch_token()

    # 拼接通用文字识别高精度url
    image_url = OCR_URL + "?access_token=" + token

    text = ""
    # 图形化选择路径
    # root = tk.Tk()
    # root.withdraw()
    # data_dir = filedialog.askdirectory(title = '请选择图片文件夹') + '/'
    # result_dir = filedialog.askdirectory(title = '请选择输出文件夹') + '/'
    data_dir = r'D:\learnpython\baiduocr\picture'
    result_dir = r'D:\learnpython\baiduocr\tmp'
    

    # if not os.path.exists(result_dir):
    #     os.mkdir(result_dir) #如果输出文件不存在，创建输出文件
    num = 0
    for picfile in os.listdir(data_dir):  #遍历调整后的图片存放的文件夹
        print('{0}:{1}正在处理'.format(num+1, picfile.split('.')[0]))
        # print(picfile)
        text_name = picfile.split('.')[0] + '.txt'
        outfile = os.path.join(result_dir,text_name)
        # print(outfile)
    # 读取书籍页面图片
    # file_content = read_file('./picture/3B2E79DE27AE48CF8F731575AAEBC22D.jpg')
        file_content = read_file(os.path.join(data_dir,picfile)) # 需要注意的是，要读取完整路径才可以正常识别
        # print(file_content)
    # 调用文字识别服务
        result = request(image_url, urlencode({'image': base64.b64encode(file_content)}))
        # print(result)
    # 解析返回结果
        result_json = json.loads(result)
        # 如果需要将前述多个内容拼接到一个文档，只需要将下述循环升级一档即可
        for words_result in result_json["words_result"]:
            text = text + words_result["words"] + '\n'
        with open(outfile, "w", encoding='utf-8') as f:
            f.write(str(text))
            f.close()
        text = '' #需要清空内容
        print('{0}:{1}处理完成。'.format(num, text_name))  