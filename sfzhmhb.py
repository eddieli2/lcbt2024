import pandas as pd
import os
import sys
df_sfz=pd.read_excel("C:\\lishan\\venv_and_code\\lcbt2024\\data\\处理结果.xlsx",dtype=object)
df_sfz_mb=pd.read_excel("C:\\lishan\\venv_and_code\\lcbt2024\\data\\2023年度个贷业务利差补贴明细表.xlsx",dtype=object)
df_sfz_mb=df_sfz_mb.reindex(['客户身份证号码', '合同号',  '客户姓名名称', '贷款业务品种', '贷款发放额度', '户籍所在地', '户籍地址',
           '利率浮动值', '是否执行先收后返', '贷款执行利率', '2022年12月末', '2023年1月末', '2023年2月末',
           '2023年3月末', '2023年4月末', '2023年5月末', '2023年6月末', '2023年7月末', '2023年8月末',
           '2023年9月末', '2023年10月末', '2023年11月末', '2023年12月末', '全年贷款平均余额',
           '上年度的贷款利差补贴', '实际获得金额', '贷款品种五级', '额度合同号'])
i=0  #创建一个计数器，用于记录处理到第几个人了

'''
完成了身份证号码的列表的构造，该列表完成了身份证号码的去重处理。这是一个思维难点。
'''
ser=df_sfz["客户身份证号码"].unique()
ls_sfz=ser.tolist()

'''
将中间结果按身份证号码进行数据分割分别输出excel
'''
#for sfzhm in ls_sfz:
#    tf=df_sfz["客户身份证号码"]==sfzhm
#    df_sfz2=df_sfz.loc[tf==True,:]
#    with pd.ExcelWriter("C:\\lishan\\venv_and_code\\lcbt2024\\data\\output\\split\\{}.xlsx".format(sfzhm)) as writer:
#        df_sfz2.to_excel(writer, sheet_name="身份证号码",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
#构造创建文件的路径，使用变量创建文件,特别注意format方法的使用

'''
定义一个函数，他的作用是，按照每行数据的类型不同选择使用不同的算法进行处理
'''
def proc_row(tuple_indivi,df_sfz2):
    global df_sfz_mb
    for row in range(tuple_indivi[0]):    #遍历每一行
        dict_hb.clear()
        sum=0
        cell_value=df_sfz2.iloc[row, 0]
        if type(cell_value)==str:
            for col in range(tuple_indivi[1]):   #遍历每一列，该行数据为字符串类型。
                    dict_hb.add(df_sfz2.iloc[row,col])
            ls_hb=list(dict_hb)
            str_hb=",".join(ls_hb)
            col_end=tuple_indivi[1]
            df_sfz2.iloc[row,col_end]=str_hb #将每行的元素都放入一个集合中，再将集合转换为字符串写入对应的“合并”单元格中。
        else:
            for col in range(tuple_indivi[1]):   #遍历每一列，该行数据为数字类型。
                    sum=sum+df_sfz2.iloc[row,col]
            col_end = tuple_indivi[1]
            df_sfz2.iloc[row, col_end] = sum
    df_sfz2=df_sfz2.T
    df_sfz_mb.iloc[i,:]=df_sfz2.loc["合并",:]


'''
按身份证号码分离每个人的数据
'''
#for sfzhm in ls_sfz:   #遍历每一个身份证号码
sfzhm="35262719810104301X"
tf=df_sfz["客户身份证号码"]==sfzhm   #构造用于数据筛选的series,将符合条件的行全部筛选出来了，并且标记为True
df_sfz2=df_sfz.loc[tf==True,:]
df_sfz2=df_sfz2.T
df_sfz2=df_sfz2.reindex(['客户身份证号码', '合同号',  '客户姓名名称', '贷款业务品种', '贷款发放额度', '户籍所在地', '户籍地址',
       '利率浮动值', '是否执行先收后返', '贷款执行利率', '2022年12月末', '2023年1月末', '2023年2月末',
       '2023年3月末', '2023年4月末', '2023年5月末', '2023年6月末', '2023年7月末', '2023年8月末',
       '2023年9月末', '2023年10月末', '2023年11月末', '2023年12月末', '全年贷款平均余额',
       '上年度的贷款利差补贴', '实际获得金额', '贷款品种五级', '额度合同号'])   #重新排列索引，将'客户身份证号码'放在'合同号'前面
tuple_indivi=df_sfz2.shape  #df的shape方法返回的是一个元组类型数据。
df_sfz2["合并"]="NaN"  #添加一个新的列用于存放每行合并后的数据
dict_hb=set()   #定义一个空字典，用于存放每一行的所有数据，之所以使用字典是因为他有去重的功能。到下一行先清空这个字典再装入数据。用于数据类型为字符串的行。
sum=0  #用于数据类型为数字类型的行进行加总计算。
proc_row(tuple_indivi, df_sfz2)  #调用函数，处理当前这个人的数据
i=i+1






