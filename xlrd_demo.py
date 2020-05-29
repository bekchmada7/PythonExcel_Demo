'''
    @使用xlrd工程对现有XLS文件进行读取操作;
    @author da
    @2020-05-27        
'''
import xlrd
#打开文件，获得文件workbook对象
wd = xlrd.open_workbook("E:\天津精益线束项目\线束_天津现场调研_20181122\LMS项目\P4.项目上线\徐水二期精准BOM测试数据0303.xlsx")
#获得Excel文件所有sheet页签
st = wd.sheet_names()
print("sheet_names====",st)
#根据索引获得sheet页签名称
worksheet = wd.sheet_by_index(0)
print("sheet by index ====",worksheet)
#根据指定名称获得sheet页签名称
worksheet = wd.sheet_by_name('Sheet1')
print("sheet_by_name====",worksheet)
#数组方式返回sheet值
st_name0 = wd.sheet_names()[0]
print("wd.sheet_names()[0]====",st_name0)

#获取表的姓名
name = worksheet.name  
print(name) 
#获取该表总行数
nrows = worksheet.nrows  
print("获取该表总行数====",nrows)  
#获取该表总列数
ncols = worksheet.ncols  
print("获取该表总列数====",ncols)

#通过获取表的行数循环输出 输出表格每一行数据

#for i in range(nrows):
#    print(worksheet.row_values(i))
    
#获取某一行信息

A1 = worksheet.cell(0,0).value.encode('UTF-8')
print("A1的值是======",A1)
B1 = worksheet.cell(0,1).value
print("B1的值是======",B1)