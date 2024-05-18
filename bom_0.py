import xlrd 
import xlwt 

# 读取模板加载模板数据到新建的文档之中 
data = xlrd.open_workbook(r'模板.xls') 
table = data.sheets()[0] 
print(table) 
 
nRows = table.nrows 
nCols = table.ncols 
 
# 读取Excel表格前七行数据并存储在列表rowsTable 
rowsTable = [] 
for num in range(8): 
 rowsTable.append(table.row_values(num))

 # 打开数据文件，将数据加载入上一步新建的文档 
print("请输入要处理的BOM文档路径：") 
bomURL = input('>') 
data1 = xlrd.open_workbook(bomURL) 
table1 = data1.sheets()[0] 
 
nRows1 = table1.nrows 
nCols1 = table1.ncols 
 
colsTable = [] 
for num1 in range(nCols1): 
 colsTable.append(table1.col_values(num1))


 # 创建一个Excel 
workBook = xlwt.Workbook(encoding='utf-8', style_compression=0) 
sheet = workBook.add_sheet('test', cell_overwrite_ok=True) 
# sheet.write(0, 0, cell_A1) 
 
# 将模板中的前八行数据填入新建文档 
for num1 in range(8): 
 ColsNum = 0 
 while ColsNum < nCols: 
    for cell_value1 in rowsTable[num1]: 
        sheet.write(num1, ColsNum, cell_value1) 
 if len(rowsTable[num1][ColsNum]) != 0: 
    ColsNum += 1 
# print("%d" % ColsNum) 
 else: 
    ColsNum = 27 
 break


# 将第1列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 2, colsTable[0][num2]) 
 
# 将第2列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 21, colsTable[1][num2]) 
 
# 将第3列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 0, colsTable[2][num2]) 
 
# 将第4列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 1, colsTable[3][num2]) 
 
# 将第5列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 4, colsTable[4][num2]) 
 
# 将第6列的数据填入新建文档 
for num2 in range(nRows1): 
 sheet.write(num2+8, 6, colsTable[5][num2])