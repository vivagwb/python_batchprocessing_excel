#用于读取对账单生成分润对账单


import xlwings as  xw
import os

import time


billpath=r'E:\新网工作\支付工作\支付收费对账单\2019年2月对账单\存管2月对账单'
fillpath=r'E:\新网工作\支付工作\支付收费对账单\2019年2月对账单\2019年2月懒猫对账模板.xlsx'


#确定商户有本月有几个收费项目
cellname_num="A12:A16"


cellname_bill_1='B12'
# cellname_bill_2='C17'
# cellname_bill_3='C6'
cellname_fill_1='B2'#填充商户名称
cellname_fill_2='C2'#填写产品类型
cellname_fill_3='D2'#填写成功笔数
cellname_fill_4='E2'#填写成功金额

sheet_bill='1'
sheet_fill="模板"



def  getbill(billpath,cellname_num,cellname_bill_1,sheetname):
    os.chdir(billpath)
    billfiles=os.listdir('.')
    billallcontent=[]

    #获取账单金额数据
    for  file in billfiles:
          exlapp=xw.App(visible=False,add_book=False)
          exlapp.display_alerts=False
          exlapp.screen_updating=False
          wb=exlapp.books.open(file)
          listnum=wb.sheets[sheet_bill].range(cellname_num).value
          paynum=listnum.index("实际应收合计")


          #获取公司名称
          billlistname=wb.sheets[sheet_bill].range(cellname_bill_1).value




          #获取需获取单元格的整个范围
          cellend='E'+str(12+paynum-1)  #获取终止范围
          cellrange='B12'+":"+cellend


          #获取收费项目内容，包括业务类型、笔数及金额

          billcontent=wb.sheets[sheet_bill].range(cellrange).value

          if isinstance(billcontent[0],list):
              billallcontent.extend(billcontent)
          else:
              billallcontent.append(billcontent)

          print(billlistname+"处理完毕")

          wb.save()
          wb.close()
          exlapp.quit()
          #time.sleep(1)
    return billallcontent

billallcontent=getbill(billpath,cellname_num,cellname_bill_1,sheet_bill)

print(billallcontent)



fillexl=xw.App(visible=False,add_book=False)
fillexl.display_alerts=False
fillexl.screen_updating=False
fillwb=fillexl.books.open(fillpath)
for i,bill in enumerate(billallcontent):
    #print(str(i)+bill+'\n')
    print(bill[0],bill[1],bill[2],bill[3])

    fillwb.sheets[sheet_fill].range(cellname_fill_1).value=bill[0]
    fillwb.sheets[sheet_fill].range(cellname_fill_2).value=bill[1]
    fillwb.sheets[sheet_fill].range(cellname_fill_3).value=bill[2]
    fillwb.sheets[sheet_fill].range(cellname_fill_4).value=bill[3]


    cellname_fill_1=cellname_fill_1[0]+str(int(cellname_fill_1[1])+i+1)
    cellname_fill_2 = cellname_fill_2[0] + str(int(cellname_fill_2[1]) + i + 1)
    cellname_fill_3 = cellname_fill_3[0] + str(int(cellname_fill_3[1]) + i + 1)
    cellname_fill_4 = cellname_fill_4[0] + str(int(cellname_fill_4[1]) + i + 1)
    print(bill[0]+"处理完毕"+'\n')
fillwb.save()

fillwb.close()
fillexl.quit()


