#用于修改文件名称


import os
import re


os.chdir(r'E:\新网工作\支付工作\支付收费对账单\2019年2月对账单\2019年2月存管开票信息')
files = os.listdir('.')

for filename in files:
     ma=re.split('1月',filename)
     newname=ma[0]+'2月'+ma[1]
     os.rename(filename,newname)
