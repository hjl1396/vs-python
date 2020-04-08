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
Save_file= 'C:/Users/hjl13/Desktop/O1/RX/Dit.xlsx'
print("请输入参数（用空格隔开）")
print("温度（K）,氧化层电容（pF）,串联电阻（Ω）,串联电感（uH）,测试半径（um）,请输入RX文件所在位置,        掺杂浓度（cm-3）")
print(r'350          46.293          24.0348       7.703         164.93  C:\Users\haojilong\Desktop\O1\RX    7.5214e15,JL1 ')
print("(若对应参数无需修改输入0)")

inp_data1=[350,46.293,24.0348,7.703,164.93,r'C:\Users\haojilong\Desktop\O1\RX',7.5214e15,r'JL1']

inp_data = input().split()

for i in range(0, len(inp_data)):
    if(inp_data[i]!='0'):
        inp_data1[i]=inp_data[i]

for i in range(0, len(inp_data1)):
    print(inp_data1[i], "\n", end="")