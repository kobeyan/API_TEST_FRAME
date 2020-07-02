
import os
import xlrd

excel_path = os.path.join(os.path.dirname(os.path.abspath( __file__)),"data\\test_data.xlsx")

workbook = xlrd.open_workbook(excel_path)
sheet = workbook.sheet_by_name("Sheet1")


# 1. 读取第二列的所有数据
col2_data = sheet.col_values(1,0,5)
print(col2_data)

# 2.编写一个方法，参数为单元格的坐标(x,y),如果给的坐标是合并的单元格，输出此单元格是合并的，否则输入是普通单元格
def ismerge(x,y):

    merged = sheet.merged_cells
    for site in merged:
        if x >= site[0] and x < site[1]:
            if y >= site[2] and y < site[3]:
                cellvalue = sheet.cell_value(site[0],site[2])
                return ("此单元格是合并单元格， 值为：",cellvalue)
    return ("此单元格是普通单元格,  值为：",sheet.cell_value(x,y))

# 3.读取完成情况，进行降序排序
valuelist = []
for i in range(1,sheet.nrows):
        valuetuple = ismerge(i,3)
        print(valuetuple)
        valuelist.append(valuetuple[1])
print(sorted(valuelist,reverse=True))
