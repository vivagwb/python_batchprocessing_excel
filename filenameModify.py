#用于修改文件名称


import os
import re


os.chdir(r'C:\Users\Administrator\Desktop\存管3月对账单')
files = os.listdir('.')

for filename in files:
     ma=re.split('2月',filename)
     newname=ma[0]+'3月'+ma[1]
     os.rename(filename,newname)
