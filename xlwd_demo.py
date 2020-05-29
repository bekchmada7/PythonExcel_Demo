'''
    @使用xlwt工程对Excel文件进行写入操作
    @author da
    @2020-05-27
'''
import xlwt
#创建一个XLS文件
book = xlwt.Workbook(encoding='UTF-8',style_compression=0)
#创建一个sheet页签并命名
sheet = book.add_sheet('基础信息',cell_overwrite_ok=True)
#直接粗暴的写入某一单元格值
sheet.write(0,0,'姓名')
sheet.write(1,0,'张三')
'''
    @通过循环遍历写入xls表值
'''
sheet2 = book.add_sheet('城市信息',cell_overwrite_ok=True)
A1 = ['北京市', '天津市', '河北省', '山西省', '内蒙古自治区']
A2 = [1000,500,300,200,100]
A3 = [2700,1800,1500,1300,1000]
A4 = ['省份','指数','人均收入']
#写入第一列
for i in range(0,len(A1)):
    sheet2.write(i+1, 0,A1[i])
#写入第二列
for i in range(0,len(A2)):
    sheet2.write(i+1, 1,A2[i])
#写入第三列
for i in range(0,len(A3)):
    sheet2.write(i+1,2,A3[i])
#写入表头
for i in range(0,len(A4)):
    sheet2.write(0, i,A4[i])

book.save('demo.xls')