# '
# ' 宏7 宏
# '
#
# '
#
# '重新创建一个列表'
# Sheets.Add
# After := ActiveSheet
#
# Range("A1").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!RC[365]"
# ActiveCell.Offset(0, 1).Range("A1").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!R[1]C[364]"
# ActiveCell.Offset(0, 1).Range("A1").Select
# ActiveCell.FormulaR1C1 = "=MAX(Sheet1!R[1]C[363]:Sheet1!R[554]C[363])"
# ActiveCell.Offset(0, 1).Range("A1").Select
# ActiveCell.FormulaR1C1 = "=MIN(Sheet1!R[1]C[362]:Sheet1!R[554]C[362])"
# ActiveCell.Offset(0, 1).Range("A1").Select
# ActiveCell.FormulaR1C1 = "=STDEV(Sheet1!R[332]C[361]:Sheet1!R[443]C[361])"
# ActiveCell.Offset(1, 0).Range("A1").Select
#
# Range("A2").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!R[-1]C[366]"
# Range("B2").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!RC[365]"
# Range("C2").Select
# ActiveCell.FormulaR1C1 = "=MAX(Sheet1!RC[364]:Sheet1!R[553]C[364])"
# Range("D2").Select
# ActiveCell.FormulaR1C1 = "=MIN(sheet1!RC[363]:Sheet1!R[553]C[363])"
# Range("E2").Select
# ActiveCell.FormulaR1C1 = "=STDEV(Sheet1!R[331]C[362]:Sheet1!R[442]C[362])"
# Range("E3").Select
# End
# Sub
# '
# ' 宏7 宏
# '
#
# '
#
# '重新创建一个列表'
# Sheets.Add
# After := ActiveSheet
# nb-nc
# a1-a2
# 365  366
# 363是总列数  365是nb在sheet1中的第365列的意思,nc就是366呗！一次类推
# 标题

# 1 - 363
#
# Range("A1").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!R[0]C[365]"
#
# Range("A2").Select
# ActiveCell.FormulaR1C1 = "=Sheet1!R[-1]C[366]"
#{1} 为0可以
title_str = 'Range("A{0}").Select\nActiveCell.FormulaR1C1 = "=Sheet1!R[{1}]C[{2}]"'
# 434是最后一个
#{0} 最大数字-n_num
#{1} 363+n_num
last_Str1 = 'Range("b1").Select'
last_Str2 = 'ActiveCell.FormulaR1C1 = "=Sheet1!R[{0}]C[{1}]\nActiveCell.Offset(1, 0).Range("A1").Select'
for n_num in range(1,364):
    _0_num = n_num
    _1_num = 1-n_num
    _2_num = 364+n_num

    last_n1 = 66-n_num
    last_n2 = 363+n_num
    print(title_str.format(_0_num,_1_num,_2_num))
    print(last_Str2.format(last_n1,last_n2))

# 已经把标题和最后一个值搞定了。后面就一个一个搞定！因为编译过程太长所以，做成5个函数单独去计算吧！




