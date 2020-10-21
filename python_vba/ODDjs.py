


# IW ---SN
# 3----254
import string
import time

f_alist=[]
alphabet =string.ascii_uppercase
head_str='Public num As Integer\nSub s()\nDim i As Integer\nFor i = 2 To 9999\nIf Range("C" & i) = "" Then\nnum = i\nExit For\nEnd If'
tail_str='Next\nRange("{0}1:{1}" & num).Select\nActiveSheet.Shapes.AddChart2(227, xlLine).Select\nActiveChart.SetSourceData Source:=Range("Sheet1!${0}$1:${1}$" & num)\nEnd Sub'
str_1 = ' Range("{0}1").Select\nActiveCell.FormulaR1C1 = "=RC[-{2}]"\nRange("{0}" & i).Select \n ActiveCell.FormulaR1C1 = "=RC[-{2}]/R2C{1}-1"'


for i1 in alphabet:
    for i2 in alphabet:
        f_al =i1+i2
        f_alist.append(f_al)

start_index=f_alist.index("AN") # 修改处1
end_index=f_alist.index("BV") # 修改处2
len_str =end_index-start_index+3
def into_file(f_name,item_name):
    try:
        with open('{0}.bas'.format(f_name), 'a') as file_handle:

            file_handle.write(item_name + "")  # 写入
            file_handle.write('\n')  # 有时放在循环里面需要自动转行，不然会覆盖上一条数据
            print("{0} 整理完毕".format(f_name))
    except:
        pass




if __name__=="__main__":


    into_file("ODDJS",head_str)
    for a_i,num in zip(f_alist[start_index:end_index+1],range(3,end_index-start_index+3+1)):

        f_code=str_1.format(a_i,num,len_str)
        into_file("ODDJS",f_code)

    into_file("ODDJS", tail_str.format(f_alist[start_index],f_alist[end_index]))



