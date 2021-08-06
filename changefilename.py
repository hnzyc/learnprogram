# coding=utf-8
# 作者：赵运超
# 2021-08-06
# 批量修改文件名称

import os

# 文件所在路径
path = input('请输入需要修改的文件所在文件夹路径：')
old = input('请输入需要替换的内容：')
new = input('请输入需要替换后的内容：')
# 遍历文件夹中所有的文件
for file in os.listdir(path):
    print(file)
    new_name = file.replace(old,new)
    os.rename(os.path.join(path,file),os.path.join(path,new_name))

for file in os.listdir(path):
    print(file)   