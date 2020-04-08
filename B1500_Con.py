import openpyxl

def write07excel(path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = '2007测试表'

    value = [["名称", "价格", "出版社", "语言"],
             ["如何高效读懂一本书", "22.3", "机械工业出版社", "中文"],
             ["暗时间", "32.4", "人民邮电出版社", "中文"],
             ["拆掉思维里的墙", "26.7", "机械工业出版社", "中文"]]
    for i in range(0, 4):
        for j in range(0, len(value[i])):
            sheet.cell(row=i+1, column=j+1, value=str(value[i][j]))
    wb.save(path)
    print("写入数据成功！")

def read07excel(path):
    wb = openpyxl.load_workbook(path)
    sheet = wb["2007测试表"]

    for row in sheet.rows:
        for cell in row:
            print(cell.value, "\t", end="")
        print()
+9

k=8.62e-5 #k/q
esio2=3.9 
esic=9.7
e0=8.85e-12 #(F/m)
q=1.6e-19
#------------------输入--------------------
Save_file='C:\Users\haojilong\Desktop\O1\RX\Dit.xlsx'


