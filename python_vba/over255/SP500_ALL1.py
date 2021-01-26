


# IW ---SN
# 3----254
import string
import time

f_alist=[]
alphabet =string.ascii_uppercase
head_str='Public num As Integer\nSub s()\nDim i As Integer\nFor i = 2 To 9999\nIf Range("C" & i) = "" Then\nnum = i\nExit For\nEnd If'

'作图代码'


tail_next = '\nNext'
tail_images = 'Range("{0}1:{1}" & num).Select\nActiveSheet.Shapes.AddChart2(227, xlLine).Select\nActiveChart.SetSourceData Source:=Range("Sheet1!${0}$1:${1}$" & num)'
tail_end = '\nEnd Sub'

str_1 = ' Range("{0}1").Select\nActiveCell.FormulaR1C1 = "=RC[-{2}]"\nRange("{0}" & i).Select \n ActiveCell.FormulaR1C1 = "=RC[-{2}]/R2C{1}-1"'


for i1 in alphabet:
    for i2 in alphabet:
        f_al =i1+i2
        f_alist.append(f_al)



for i1 in alphabet:
    for i2 in alphabet:
        for i3 in alphabet:

            f_al =i1+i2+i3
            f_alist.append(f_al)


start_index=f_alist.index("KD") # 修改处1
end_index=f_alist.index("VB") # 修改处2
len_str =end_index-start_index+3
print(len_str)
def into_file(f_name,item_name):
    try:
        with open('{0}.bas'.format(f_name), 'a') as file_handle:

            file_handle.write(item_name + "")  # 写入
            file_handle.write('\n')  # 有时放在循环里面需要自动转行，不然会覆盖上一条数据
            print("{0} 整理完毕".format(f_name))
    except:
        pass




if __name__=="__main__":
    for i1 in alphabet:
        for i2 in alphabet:
            for i3 in alphabet:
                f_al = i1 + i2 + i3
                f_alist.append(f_al)
    print(f_alist[start_index],f_alist[end_index])
    start_index = f_alist.index("KD")  # 修改处1
    end_index = f_alist.index("VB")  # 修改处2
    # NB WQ WR AAX
    # print(f_alist[start_index],f_alist[start_index+249],f_alist[start_index+250],f_alist[end_index])

    into_file("SP500_ALL1",head_str)
    for a_i,num in zip(f_alist[start_index:end_index+1],range(3,end_index-start_index+3+1)):

        f_code=str_1.format(a_i,num,len_str)
        into_file("SP500_ALL1",f_code)
    into_file("SP500_ALL1",tail_next)

    into_file("SP500_ALL1", tail_images.format(f_alist[start_index],f_alist[start_index+100]))
    into_file("SP500_ALL1", tail_images.format(f_alist[start_index+101],f_alist[start_index+200]))
    into_file("SP500_ALL1", tail_images.format(f_alist[start_index+201],f_alist[end_index]))
    into_file("SP500_ALL1",tail_end)



