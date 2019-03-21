import os


billnamelist=os.listdir(r'E:\新网工作\支付工作\支付收费对账单\2019年2月对账单\存管2月对账单')   #对账单路径
invoicenamelist=os.listdir(r'E:\新网工作\支付工作\支付收费对账单\2019年2月对账单\2019年2月存管开票信息')   #开票申请单路径


billlist=[]
invoicelist=[]

template=[]
template_1=[]

for name in invoicenamelist:
    namesplit=name.split('开')[0]
    print(namesplit)
    invoicelist.append(namesplit)

# print(invoicelist)
# print('\n')


for name in billnamelist:
    namesplit=name.split('2月')[0]
    print(namesplit)
    billlist.append(namesplit)

# print(billlist)
# print('\n')

for name in billlist:
    if name in invoicelist:
        pass

    else:
        #print(name+'无开票模板')
        template.append(name)

print('无开票模板：\n')
print(template)


for name in invoicelist:
    if name in billlist:
        pass
    else:
        #print(name+"该商户本月无交易")
        template_1.append(name)

print('该商户伍交易：\n')
print(template_1)