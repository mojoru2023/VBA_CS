


# IW ---SN
# 3----254
import string
import time

f_alist=[]
alphabet =string.ascii_uppercase
title_str1='    Range("A1").Select\n   ActiveCell.FormulaR1C1 = "=Sheet1!RC[256]"\n    ActiveCell.Offset(1, 0).Range("A1").Select'
title_str2 ='    ActiveCell.FormulaR1C1 = "=Sheet1!R[{0}]C[{1}]"\n    ActiveCell.Offset(1, 0).Range("A1").Select'

max_str1= ' Range("b1").Select'
max_str2= 'ActiveCell.FormulaR1C1 = "=max(Sheet1!R[{0}]C[{1}]:Sheet1!R[{2}]C[{1}])"\n   ActiveCell.Offset(1, 0).Range("A1").Select'


min_str1= ' Range("c1").Select'
min_str2= '    ActiveCell.FormulaR1C1 = "=MIN(Sheet1!R[{0}]C[{1}]:Sheet1!R[{2}]C[{1}])"\n    ActiveCell.Offset(1, 0).Range("A1").Select'



stdev_str1='Range("d1").Select'
stdev_str2='    ActiveCell.FormulaR1C1 = "=STDEV(Sheet1!R[{0}]C[{2}]:Sheet1!R[{1}]C[{2}])"\n    ActiveCell.Offset(1, 0).Range("A1").Select'

tail_str='\nEnd Sub'


for i1 in alphabet:
    for i2 in alphabet:
        f_al =i1+i2
        f_alist.append(f_al)

start_index=f_alist.index("IW") # 修改处1
end_index=f_alist.index("SN") # 修改处2
len_str =end_index-start_index
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
    title_list=[title_str1]
    max_list= [max_str1]
    min_list= [min_str1]
    max_ = []
    min_ =[]
    stdev_list=[stdev_str1]
    getNewSheet=["sub sm() \nSheets.Add After:=ActiveSheet"]

    for item in range(-len_str-1,2):
        max_.append(item)
    for item in range(-len_str-2,2):
        min_.append(item)
    max_.reverse()
    min_.reverse()
    print(len(max_))



    # into_file("ZGG_",head_str)
    for nu in (range(1,end_index-start_index+1)):
        a1= -nu
        a2=66+nu
        f_title_str = title_str2.format(a1,a2)
        title_list.append(f_title_str)

    #最大完成
    for nu,num_max in zip(range(1,end_index-start_index+2),max_):

        # {0} num_max
        # {1} 64+num
        # {2} 666-num
        #
        # 最大值666
        # 反向从1开始，0  -1

        max_1 = num_max
        max_2 = 64+nu
        max_3 = 9999-nu
        f_max_str =max_str2.format(max_1,max_2,max_3)
        #print(f_max_str)
        max_list.append(f_max_str)


    for nu,num_min in zip(range(1,end_index-start_index+2),min_):

# 反向从1开始 0 -1,最大是666
# {0} num_min
# {1} 63+num
# {2}
        min_1 = num_min
        min_2 = 63+nu
        min_3 = 9999-nu

        f_min_str =min_str2.format(min_1,min_2,min_3)
       # print(f_min_str)
        min_list.append(f_min_str)




        #最后一个数 367 267
# {0} 367-num
# {1} 367-num-100
# {2} 62+num
    for nu in range(1,end_index-start_index+2):
        stdev1=999-nu
        stdev2=999-nu-767 # 230-999
        stdev3=62+nu
        f_stdev_str =stdev_str2.format(stdev1,stdev2,stdev3)
        stdev_list.append(f_stdev_str)
    NewSheet_title_max_min_stdev = getNewSheet+title_list + max_list + min_list+stdev_list
    for i in NewSheet_title_max_min_stdev:
        into_file("bigJS_new_mmsl",i)
    into_file("bigJS_new_mmsl",tail_str)














