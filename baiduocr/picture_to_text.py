#-*- coding = uft-8 -*-
#@Time : 2020/5/20 10:15 上午
#@Author : J哥
#@File : pic_to_text.py

"""
利用百度api实现图片文本识别
"""

import glob  #文件名模式匹配，不用遍历整个目录判断每个文件是不是符合
import os  #操作文件的库
from os import path
from aip import AipOcr  #调取百度AI接口所需库
from PIL import Image   #处理图片的库

'''
picfile:原始图片路径
outdir：输出图片路径
'''

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
    new_img.save(path.join(outdir, os.path.basename(picfile))) #新图片保存在outdir即tmp目录下。os.path.basename(path)表示返回文件名。tmp/pic.jpg


# 利用百度api识别文本，并保存提取图片中的文字
def baiduOCR(picfile, outfile):
    filename = path.basename(picfile) #将图片路径赋值给filename

    APP_ID = '24597675'  # 刚才获取的 ID，下同
    API_KEY = 'non3MelzmCk4uWeRKLojMyTi'
    SECRECT_KEY = 'fsDb2Cp13mcYb6kUfv5UUM0L2oEe3e2f'
    client = AipOcr(APP_ID, API_KEY, SECRECT_KEY)

    i = open(picfile, 'rb') #以二进制只读方式打开文件

    '''
    "r" 以读方式打开，只能读文件 ， 如果文件不存在，会发生异常      
    "w"  以写方式打开，只能写文件， 如果文件不存在，创建该文件；如果文件已存在，先清空，再打开文件   
    "rb" 以二进制读方式打开，只能读文件 ， 如果文件不存在，会发生异常      
    "wb" 以二进制写方式打开，只能写文件， 如果文件不存在，创建该文件；如果文件已存在，先清空，再打开文件
    "a+": 附加读写方式打开
    '''

    img = i.read()  #读取图片
    print("正在识别图片：\t" + filename)

    '''
    \t表示缩进，相当于按一下Tab键
    \n表示换行，相当于按一下Enter键
    \n\t表示换行加缩进
    '''

    message = client.basicGeneral(img)  # 通用文字识别，每天50000次免费
    #message = client.basicAccurate(img)   # 通用文字高精度识别，每天500次免费
    print("识别成功！")
    i.close() #关闭文件

# 文本识别结果输出为txt
    with open(outfile, 'a+') as fo:
        fo.writelines("+" * 60 + '\n')   #分隔线，60个+表示
        fo.writelines("识别图片：\t" + filename + "\n" * 2)  #正在识别的图片名，并空两行
        fo.writelines("文本内容：\n")
        # 输出文本内容
        for text in message.get('words_result'):  #words_result识别结果数组，类型为array[]
            #print(text)
            fo.writelines(text.get('words') + '\n')  #words识别结果字符串
        fo.writelines('\n' * 2)
    print("文本导出成功！")


if __name__ == "__main__":

    outfile = 'export.txt'  #输出文件
    outdir = './tmp'  #临时文件
    if path.exists(outfile):
        os.remove(outfile)  #如果输出文件已存在，删除
    if not path.exists(outdir):
        os.mkdir(outdir) #如果输出文件不存在，创建输出文件
    print("压缩过大的图片...")
    # 首先对过大的图片进行压缩，以提高识别速度，将压缩的图片保存到临时文件夹中
    for picfile in glob.glob("./picture/*"):  #遍历原始图片存放的文件夹
        convertimg(picfile, outdir)  #调整原始图片
    print("图片识别...")
    for picfile in glob.glob("./tmp/*"):  #遍历调整后的图片存放的文件夹
        baiduOCR(picfile, outfile)  #识别图片文本
        os.remove(picfile)  #识别后删除图片
    print('图片文本提取结束！文本输出结果位于 %s 文件中。' % outfile)
    os.removedirs(outdir) #递归删除目录，即如果子文件夹成功删除, removedirs()才尝试它们的父文件夹
