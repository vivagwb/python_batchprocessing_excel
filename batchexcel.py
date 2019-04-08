#批量修改excel指定单元格

import xlwings as  xw
import os


os.chdir(r'C:\Users\Administrator\Desktop\存管3月对账单')
print(os.getcwd())
files=os.listdir('.')

modifycellname_1='A4'#修改申请日期
modifycontent_1='2019年3月充值服务费'#申请日期


modifycellname_2='A9'#修改开票事由的日期
modifycontent_2='对账区间：3月1日-3月31日'


#modifycellname_3='C18'#修改备注信息的日期
#modifycontent_3='2019年2月充值服务费'


sheetname='1.开具增值税发票登记表'

for  file in files:
          exlapp=xw.App(visible=False,add_book=False)
          exlapp.display_alerts=False
          exlapp.screen_updating=False
          wb=exlapp.books.open(file)
          wb.sheets[sheetname].range(modifycellname_1).value=modifycontent_1
          wb.sheets[sheetname].range(modifycellname_2).value = modifycontent_2
          #wb.sheets[sheetname].range(modifycellname_3).value = modifycontent_3
          print(file+'修改成功')
          wb.save()
          wb.close()
          exlapp.quit()
