# -*- coding: utf-8 -*- 
# Python3 
# 含部件名的合格证

import xlsxwriter
import load

filename = input('文件名：') + '.xlsx'
workbook = xlsxwriter.Workbook(filename)  #创建XLSX文件
xiangmu = ['产品名称:', '部件名称:', '产品货号:', '产品编号:', '生产批号:', '灭菌批号:', '检验日期:', '检验员号:']  #项目列表

#设置格式----------------------------
format_cop = workbook.add_format()
format_cop.set_font_size(12)
format_cop.set_align('center')
format_cop.set_align('vcenter')
format_cop.set_top()
format_cop.set_left()
format_cop.set_right()
format_cop.set_font_name('宋体')

format_title = workbook.add_format()
format_title.set_font_size(16)
format_title.set_bold()
format_title.set_align('center')
format_title.set_align('vcenter')
format_title.set_left()
format_title.set_right()
format_title.set_font_name('宋体')

format_text = workbook.add_format()
format_text.set_font_size(10.5)
format_text.set_align('center')
format_text.set_align('vcenter')
format_text.set_left()
format_text.set_font_name('宋体')

format_text_u = workbook.add_format()
format_text_u.set_font_size(10.5)
format_text_u.set_bottom()
format_text_u.set_align('center')
format_text_u.set_align('vcenter')
format_text_u.set_font_name('宋体')

format_b = workbook.add_format()
format_b.set_font_size(10.5)
format_b.set_right()
format_b.set_align('center')
format_b.set_align('vcenter')
format_b.set_font_name('宋体')

format_b2 = workbook.add_format()
format_b2.set_font_size(10.5)
format_b2.set_right()
format_b2.set_left()
format_b2.set_bottom()
format_b2.set_align('center')
format_b2.set_align('vcenter')
format_b2.set_font_name('宋体')
#设置格式结束----------------------------

'''
#获取合格证批次数量，便于生成相同数量的工作表
pici = int(input('合格证批次数量：'))  

for i in range(pici):
	print('----输入第%s批信息，还有%s批----' % (i + 1, pici - i - 1))
	pinming = input('产品名称：')
	bujian = input('部件名称：')
	guige = input('产品货号：')
	pihao = input('生产批号：')
	miejun = input('灭菌批号：')
	jianyanriqi = input('检验日期：')
	startnum = int(input('起始编号：(纯数字，默认为4位数字如0001)'))
	endnum = int(input('结束编号：(纯数字，默认为4位数字如0001)'))
	qianzhui = input('编号前有无前缀？如BAM/C之类，如没有直接回车') #以上获取合格证信息

'''
#载入数据
data_in = load.data_list
qianzui = ''
print(data_in)

for i in range(len(data_in)):
	#print(data_in[i][7])
	#生成工作表，用批号加序号作为工作表名
	worksheet = workbook.add_worksheet(str(i)) 
	
	#设置页边距
	worksheet.set_margins(0.6, 0.6, 1.5, 1.5)


	#生成合格证的计数值,初始值为1
	col = 1  


	#生成合格证需用到的行数
	r = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]


	#判断合格证生成的数量，数量未达到时无限循环
	#while col <= (data_in[i][7]): 
	while col <= data_in[i][7] - data_in[i][6] + 1: 
		
		#用到的列数
		c = [0, 1, 2, 3]
		

		#每一行生成三张合格证
		for n in range(3): 

			#如果合格证数量已达到，跳出循环
			#if col > (data_in[i][7]):
			if col > data_in[i][7] - data_in[i][6] + 1: 
				break
			
			#合格证上的序列号，默认为4位数字
			sn = str( col).zfill(4) 
			
			#设置合格证各行高，各列宽
			worksheet.set_row(r[0], 24)
			worksheet.set_row(r[1], 24)
			worksheet.set_row(r[2], 24)
			worksheet.set_row(r[3], 24)
			worksheet.set_row(r[4], 24)
			worksheet.set_row(r[5], 24)
			worksheet.set_row(r[6], 24)
			worksheet.set_row(r[7], 24)
			worksheet.set_row(r[8], 24)
			worksheet.set_row(r[9], 24)
			worksheet.set_row(r[10], 15)
			worksheet.set_row(r[11], 5)

			worksheet.set_column(c[0], c[0], 9)
			worksheet.set_column(c[1], c[1], 20)
			worksheet.set_column(c[2], c[2], 0.54)
			worksheet.set_column(c[3], c[3], 1) 
		
			
			#填写合格证信息
			worksheet.merge_range(chr(65 + c[0]) + str(r[1]) + ':' + chr(67 + c[0]) + str(r[1]),'四川大学生物材料工程研究中心',format_cop) #A1:C1
			worksheet.merge_range(chr(65 + c[0]) + str(r[2]) + ':' + chr(67 + c[0]) + str(r[2]),'产品合格证',format_title)
			worksheet.merge_range(chr(67 + c[0]) + str(r[3]) + ':' + chr(67 + c[0]) + str(r[10]),' ',format_b)
			worksheet.merge_range(chr(65 + c[0]) + str(r[11]) + ':' + chr(67 + c[0]) + str(r[11]),' ',format_b2)
			
			worksheet.write_column(chr(65 + c[0]) + str(r[3]), xiangmu, format_text)
			worksheet.write(chr(66 + c[0]) + str(r[3]), data_in[i][0], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[4]), data_in[i][1], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[5]), data_in[i][2], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[6]), sn, format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[7]), data_in[i][3], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[8]), data_in[i][4], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[9]), data_in[i][5], format_text_u)
			worksheet.write(chr(66 + c[0]) + str(r[10]), ' ', format_text_u)

			#每生成一张合格证，计数+1，同时列数+4，准备生成下一张
			col += 1 

			for num in range(len(c)):
				c[num] += 4

			
		#一排生成3张合格证后，各行数+12，跳往下一行继续生成合格证
		for num in range(len(r)):
			r[num] += 12

workbook.close()

