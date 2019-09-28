import xlwings as xw
# wb = xw.Book()  #创建一个新的工作簿
# wb = xw.Book('111.xlsx')  #连接到当前工作目录中的现有文件

# 在Windows上：使用原始字符串来转义反斜杠
wb = xw.Book(r'E:\PycharmProjects\xlwingsLearning\xltemplates\template1.xlsx')

#实例化工作表对象
sht = wb.sheets['Sheet1']

#在Range内读取单元格值，写入值
sht.range('A1').value = '250'
sht.range('B1').value = '8'
#可以直接输入公式
sht.range('C1').value = '=A1/B1'
a1_value = sht.range('A1').value
b1_value = sht.range('B1').value
c1_value = sht.range('C1').value
print(a1_value)
print(b1_value)
#直接得到公式的值
print(c1_value)

#二、直接与当前活动工作部交互,如果找不到活动的工作簿，会报错。
xw.Range('D1').value = 'ActiveSheetRange'
print(xw.Range('D1').value)