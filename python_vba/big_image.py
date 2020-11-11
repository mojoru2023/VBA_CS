


# IW ---SN
# 3----254
import string
import time


alphabet =string.ascii_uppercase
head_str='Sub g()'
tail_str='End Sub'
str_1 = ' Range("{0}:{0}").Select  \n  ActiveSheet.Shapes.AddChart2(227, xlLine).Select \n   ActiveChart.SetSourceData Source:=Range("Sheet1!{0}:{0}")'




def into_file(f_name,item_name):
    try:
        with open('{0}.bas'.format(f_name), 'a') as file_handle:

            file_handle.write(item_name + "")  # 写入
            file_handle.write('\n')  # 有时放在循环里面需要自动转行，不然会覆盖上一条数据
            print("{0} 整理完毕".format(f_name))
    except:
        pass



if __name__=="__main__":
    f_alist = []
    for i1 in alphabet:
        f_alist.append(i1)
    for i1 in alphabet:
        for i2 in alphabet:
            f_al = i1 + i2
            f_alist.append(f_al)
    start_index = f_alist.index("B")  # 修改处1
    end_index = f_alist.index("AY")  # 修改处2
    len_str = end_index - start_index + 1
    into_file("BIG_IMAGES", head_str.format(f_alist[start_index],f_alist[end_index]))


    for a_i in f_alist[start_index:end_index + 1]:


        f_code=str_1.format(a_i)
        into_file("BIG_IMAGES",f_code)

    into_file("BIG_IMAGES", tail_str.format(f_alist[start_index],f_alist[end_index]))



