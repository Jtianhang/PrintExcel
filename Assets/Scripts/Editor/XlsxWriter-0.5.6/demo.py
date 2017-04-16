
# -*- coding: utf-8-*-
import xlsxwriter
import sys
import os

#带生成excel的文本参数传递进来，这里取到它的目录 
folder,path=os.path.split(sys.argv[1])
#在他的同级目录下生成.xlsx
xlsxPath = folder + '/demo.xlsx'

#准备开始生成excel
workbook = xlsxwriter.Workbook(xlsxPath)
worksheet = workbook.add_worksheet()
row = 0

#读取文本，按照♀符号隔开的顺序
file = open(sys.argv[1], 'r')
lines = [ x.strip().split('♀') for x in file.readlines() ]

#循环遍历，把内容写进excel中
for line in lines[1:]:
	#这里要注意，因为设置文本的格式是UTF-8不然会报错喔
    worksheet.write(row, 0, unicode(line[0] , 'utf-8'))
    worksheet.write(row, 1, unicode(line[1] , 'utf-8'))
    row += 1

workbook.close()


