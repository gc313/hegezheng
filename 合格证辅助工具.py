import math
while 1:
	lie = 8#每列页数
	print('-------------------')
	print('每列页数为%s' % lie)

	sta = int(input('合格证起始编号：'))

	fin = int(input('合格证结束编号：'))
	con = fin - sta + 1

	tar = (math.floor(con / 72)) * lie

	print('从第%s页后开始修改' % tar)
