import xlwings as xw

#chapter : python to excel
#新建工作簿方式：

wb = xw.Book(r'E:\PycharmProjects\xlwingsLearning\xltemplates\template1.xlsx') #直接新建一个临时工作簿

# wb = xw.books.add() # 执行报错
#
# #xw.Book('Book1')
#
# #引用活动App
# app = xw.apps.active  #引用活动app
#
# #引用活动工作簿
#
# wb1 = xw.books.active
#
# wb2 = app.books.active
#
# #活动工作表
# sht1 = xw.sheets.active
# sht2 = wb.sheets.active
#
# #在活动应用程序的活动工作簿的活动表上
# xw.Range('A1')
#
# #单元格表示法
#
# xw.Range('A1')
#
# xw.Range((1,1))  #相当于A1
#
# xw.Range('A1:C3')
#
# xw.Range((1,1),(3,3))  #相当于引用A1:C3
#
# xw.Range('NamedRange')  #引用命名区域
#
# xw.Range(xw.Range('A1'),xw.Range('C3')) #xw.Range嵌套


#print(xw.apps.keys())

sht = wb.sheets(2)

sht.range('A1').value = 3

#Range索引/切片
rng = sht.range('A2:F2')

print(rng[0,0].value)

print(rng[1].value)

print(rng[:,3:].value)

print(rng[0,:].value)









