#python3 不含部件名的合格证
import xlsxwriter

filename = input('文件名：') + '.xlsx'
workbook = xlsxwriter.Workbook(filename)  #创建XLSX文件
xiangmu = ['产品名称:', '规格型号:', '产品编号:', '生产批号:', '灭菌批号:', '检验日期:', '检验员号:']  #项目列表

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



pici = int(input('合格证批次数量：'))  #获取合格证批次数量，便于生成同数量的工作表


for i in range(pici):
	print('----输入第%s批信息，还有%s批----' % (i + 1, pici - i - 1))
	pinming = input('产品名称：')
	#bujian = input('部件名称：')
	guige = input('规格型号：')
	pihao = input('生产批号：')
	miejun = input('灭菌批号：')
	jianyanriqi = input('检验日期：')
	startnum = int(input('起始编号：(纯数字，默认为4位数字如0001)'))
	endnum = int(input('结束编号：(纯数字，默认为4位数字如0001)'))
	qianzhui = input('编号前有无前缀？如BAM/C之类，如没有直接回车') #以上获取合格证信息


	worksheet = workbook.add_worksheet(pihao + '_' + str(i)) #生成工作表，用批号加序号作为工作表名
	worksheet.set_margins(0.6, 0.6, 1.5, 1.5)


	col = 1  #生成合格证的计数值


	#生成合格证需用到的行数
	r_0 = 0
	r_1 = 1
	r_2 = 2
	r_3 = 3
	r_4 = 4
	r_5 = 5
	r_6 = 6
	r_7 = 7
	r_8 = 8
	r_9 = 9
	r_10 = 10
	#r_11 = 11




	while col <= (endnum - startnum + 1): #判断合格证生成的数量，数量未达到时无限循环
		
		#用到的列数
		c_0 = 0
		c_1 = 1
		c_2 = 2
		c_3 = 3


		for i in range(3): #每一行生成三张合格证

			if col > (endnum - startnum + 1):  #如果合格证数量已达到，跳出循环
				break
			sn = str(int(startnum) + col - 1).zfill(4) #合格证上的序列号，默认为4位数字

			worksheet.set_row(r_0, 24)
			worksheet.set_row(r_1, 24)
			worksheet.set_row(r_2, 24)
			worksheet.set_row(r_3, 24)
			worksheet.set_row(r_4, 24)
			worksheet.set_row(r_5, 24)
			worksheet.set_row(r_6, 24)
			worksheet.set_row(r_7, 24)
			worksheet.set_row(r_8, 24)
			worksheet.set_row(r_9, 15)
			worksheet.set_row(r_10, 5)
			#worksheet.set_row(r_11, 5)

			worksheet.set_column(c_0, c_0, 9)
			worksheet.set_column(c_1, c_1, 20)
			worksheet.set_column(c_2, c_2, 0.54)
			worksheet.set_column(c_3, c_3, 1) #以上设置合格证各行高，各列宽
		
			

			#填写合格证信息
			worksheet.merge_range(chr(65 + c_0) + str(r_1) + ':' + chr(67 + c_0) + str(r_1),'公司名称',format_cop) #A1:C1
			worksheet.merge_range(chr(65 + c_0) + str(r_2) + ':' + chr(67 + c_0) + str(r_2),'产品合格证',format_title)
			worksheet.merge_range(chr(67 + c_0) + str(r_3) + ':' + chr(67 + c_0) + str(r_9),' ',format_b)
			worksheet.merge_range(chr(65 + c_0) + str(r_10) + ':' + chr(67 + c_0) + str(r_10),' ',format_b2)
			
			worksheet.write_column(chr(65 + c_0) + str(r_3), xiangmu, format_text)
			worksheet.write(chr(66 + c_0) + str(r_3), pinming, format_text_u)
			#worksheet.write(chr(66 + c_0) + str(r_4), bujian, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_4), guige, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_5), qianzhui + sn, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_6), pihao, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_7), miejun, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_8), jianyanriqi, format_text_u)
			worksheet.write(chr(66 + c_0) + str(r_9), ' ', format_text_u)

			col += 1 #每生成一张合格证，计数+1，同时列数+4，准备生成下一张
			
			c_0 += 4
			c_1 += 4
			c_2 += 4
			c_3 += 4
			
		#一排生成3张合格证后，各行数+12，跳往下一行继续生成合格证
		r_0 += 11
		r_1 += 11
		r_2 += 11
		r_3 += 11
		r_4 += 11
		r_5 += 11
		r_6 += 11
		r_7 += 11
		r_8 += 11
		r_9 += 11
		r_10 += 11
		#r_11 += 12

workbook.close()
