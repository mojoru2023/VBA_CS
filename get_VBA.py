


# IW ---SN
# 3----254
import string
import time

f_alist=[]
alphabet =string.ascii_uppercase
str_1 = ' Range("{0}1").Select\nActiveCell.FormulaR1C1 = "=RC[-254]"\nRange("{0}" & i).Select \n ActiveCell.FormulaR1C1 = "=RC[-254]/R2C{1}-1"'
for i1 in alphabet:
    for i2 in alphabet:
        f_al =i1+i2
        f_alist.append(f_al)

start_index=f_alist.index("IW")
end_index=f_alist.index("SN")
for a_i,num in zip(f_alist[start_index:end_index+1],range(3,255)):
    f_code=str_1.format(a_i,num)


    try:
        with open('p_vba.txt', 'a') as file_handle:
            # .txt可以不自己新建,代码会自动新建
            file_handle.write(f_code + ",")  # 写入
            file_handle.write('\n')  # 有时放在循环里面需要自动转行，不然会覆盖上一条数据
            print("{0} 整理完毕".format("p_Vba.txt"))
    except:
        pass


