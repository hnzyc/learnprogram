from docx2pdf import convert
import os
import pdfplumber
import re
import pandas as pd

def pdf2info(file):
    """
    输入转让登记表的文件名，
    从该文件提取投资者信息，
    返回投资者信息列表
    """   
    with pdfplumber.open(os.path.join(file)) as pdf:
            page01 = pdf.pages[0] #指定页码
            table1 = page01.extract_table()#提取单个表格
            basic = table1[11][0]
            basic_l=basic.split('\n')
            count_name = basic_l[4].split('：')[1]
            count_bank = basic_l[5].split('：')[1]
            count_num = basic_l[6].split('：')[1]
            cus_name = table1[2][3]
            cus_address = table1[5][3]
            cus_persons = table1[7][3]
            cus_phones = table1[9][3]
            cus_mails = table1[10][3].replace("\n", "")
            # product = table1[0][1]
            list = [cus_name,cus_persons,cus_mails,cus_phones,count_name,count_bank,count_num,cus_address]
            # try:
            #     product_info = basic_info(product)[1]
            # except:
            #     list = [cus_name,cus_persons,cus_mails,cus_phones,count_name,count_bank,count_num,cus_address]
            # else:
            #     list = [cus_name,product_info,cus_persons,cus_mails,cus_phones,count_name,count_bank,count_num,cus_address]

    return list

def pdf2info2(file):
    """
    输入结算单的文件名，
    从该文件提取投资者信息，
    返回投资者信息结算信息列表
    """   
    with pdfplumber.open(os.path.join(file)) as pdf:
            page01 = pdf.pages[0] #指定页码
            table1 = page01.extract_table()#提取单个表格
            
            product_num = table1[2][1] # 资产代码
            product_name = table1[2][3].replace('\n','') # 资产全称
            product_quant = table1[3][1] # 交易规模
            product_price = table1[3][3] # 成交金额
            cus_name_2 = table1[12][1] # 机构全称
            cus_count_name = table1[13][1] # 账户全称
            cus_count_num = table1[14][1] # 账户账号
                        
            list = [product_num,product_name,product_quant,product_price,cus_name_2,cus_count_name,cus_count_num]
    return list

def basic_info(reg_str):
    """正则提取【】内容，并返回一个列表"""
    regex = r"[（【](.*?)[）】]"
    matches = re.findall(regex, reg_str)
    return matches 

def word2pdf(files_path):
    """对于给定的文件夹，转换文件夹内所有docx文件为同名的pdf文件"""
    files = os.listdir(files_path)
    for f in files:
        if f.endswith('docx') and '转让登记表' in f:
            f_new=os.path.splitext(f)[0]+".pdf"
            convert(os.path.join(files_path,f), os.path.join(files_path,f_new))
            
def cust_info_list(files_path):
    """对于给定的文件夹，将文件夹内转让登记表pdf文件信息提取出来
    返回pandas的DataFrame
    """
    col = []
    files = os.listdir(files_path)
    customer_info_list=[]
    for f in files:
        if f.endswith('pdf') and '转让登记表' in f:
            file = os.path.join(files_path,f)
            customer_info = pdf2info(file)
            customer_info_list.append(customer_info)
    # col = ['投资者全称','联系人','联系邮箱','联系人手机','投资者账户名称','开户行','账号','地址']
    if len(customer_info)==8:
        col = ['投资者全称','联系人','联系邮箱','联系人手机','投资者账户名称','开户行','账号','地址']        
    else:
        col = ['投资者全称','购买份额层级','联系人','联系邮箱','联系人手机','投资者账户名称','开户行','账号','地址']        
    data = pd.DataFrame(customer_info_list,columns=col)
    data = data.drop_duplicates()

    return data

def cust_info_list2(files_path):
    """对于给定的文件夹，将文件夹内结算单pdf文件信息提取出来
    返回pandas的DataFrame
    """
    files = os.listdir(files_path)
    customer_info_list=[]
    for f in files:
        if f.endswith('pdf') and '结算单' in f:
            file = os.path.join(files_path,f)
            customer_info = pdf2info2(file)
            customer_info_list.append(customer_info)

    customer_info_list

    col = ['资产代码','资产全称','交易规模','成交金额','机构全称','账户全称','银登账号']
    data = pd.DataFrame(customer_info_list,columns=col)
    
    return data