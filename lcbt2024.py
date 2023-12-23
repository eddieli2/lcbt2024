'''
1:df_mb表示目标表格的意思
2:df_mb2表示df_mb的一个副本
'''
import pandas as pd
import openpyxl
'''
将每个月的数据excle读入python程序中，存放在df变量中；
将输出表格（仅包含表结构）也读入到python中，存放在变量df_mb中
'''
df_202212=pd.read_excel("202212即期.xlsx",dtype=object)

df_202301=pd.read_excel("202301即期.xlsx",dtype=object)

df_202302=pd.read_excel("202302即期.xlsx",dtype=object)

df_202303=pd.read_excel("202303即期.xlsx",dtype=object)

df_202304=pd.read_excel("202304即期.xlsx",dtype=object)

df_202305=pd.read_excel("202305即期.xlsx",dtype=object)

df_202306=pd.read_excel("202306即期.xlsx",dtype=object)

df_202307=pd.read_excel("202307即期.xlsx",dtype=object)

df_202308=pd.read_excel("202308即期.xlsx",dtype=object)

df_202309=pd.read_excel("202309即期.xlsx",dtype=object)

df_202310=pd.read_excel("202310即期.xlsx",dtype=object)

df_202311=pd.read_excel("202311即期.xlsx",dtype=object)

#df_202312=pd.read_excel("202312即期.xlsx",dtype=object)

df_mb=pd.read_excel("2023年度个贷业务利差补贴明细表.xlsx")

'''
先将df_mb的序号0到9的列使用当年12月的对应数据填入
'''
df_mb["合同号"]=df_202311["合同号"]

df_mb["客户身份证号码"]=df_202311["证件号码"]

df_mb["客户姓名名称"]=df_202311["姓名"]

df_mb["贷款业务品种"]=df_202311["贷款品种2级"]

df_mb["贷款发放额度"]=df_202311["发放金额"]

df_mb["户籍所在地"]=df_202311["户籍所在地"]

df_mb["户籍地址"]=df_202311["户籍地址"]

df_mb["利率浮动值"]=df_202311["利率浮动值"]

df_mb["是否执行先收后返"]="否"

df_mb["贷款执行利率"]=df_202311["执行利率"]

'''
将每月数据excel中的余额填入到df_mb中对应的列中。
'''
df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202212, how="left", on=["合同号"])
df_mb["2022年12月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202301, how="left", on=["合同号"])
df_mb["2023年1月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202302, how="left", on=["合同号"])
df_mb["2023年2月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202303, how="left", on=["合同号"])
df_mb["2023年3月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202304, how="left", on=["合同号"])
df_mb["2023年4月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202305, how="left", on=["合同号"])
df_mb["2023年5月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202306, how="left", on=["合同号"])
df_mb["2023年6月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202307, how="left", on=["合同号"])
df_mb["2023年7月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202308, how="left", on=["合同号"])
df_mb["2023年8月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202309, how="left", on=["合同号"])
df_mb["2023年9月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202310, how="left", on=["合同号"])
df_mb["2023年10月末"]=df_mb2["余额"]

df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202311, how="left", on=["合同号"])
df_mb["2023年11月末"]=df_mb2["余额"]

# df_mb2=df_mb.copy()
# df_mb2=df_mb2.merge(df_202312, how="left", on=["合同号"])
# df_mb["2023年12月末"]=df_mb2["余额"]

'''
将df_mb中的所有nan数字0替换
'''
df_mb.fillna(0,inplace=True)

'''
上年末，2022年12月的源表里“余额”为0的行删除，先筛选出2022年12月“余额”为0的合同号，
将筛选出来的合同号与2023年12月源表中的合同号对比，匹配上的行删除，也就是在目标表中删除。
'''
tf=df_202212["余额"]>0
df_0=df_202212.loc[tf==False,:]
ls_0=df_0["合同号"].tolist()
df_mb=df_mb.set_index("合同号")
df_mb=df_mb.drop(index=ls_0)
df_mb.reset_index()

'''
计算全年贷款平均余额
'''
df_mb["全年贷款平均余额"]=(df_mb["2022年12月末"]*0.5+df_mb["2023年1月末"]+df_mb["2023年2月末"]+df_mb["2023年3月末"]+df_mb["2023年4月末"]+\
df_mb["2023年5月末"]+df_mb["2023年6月末"]+df_mb["2023年7月末"]+df_mb["2023年8月末"]+df_mb["2023年9月末"]+df_mb["2023年10月末"]+df_mb["2023年11月末"]+\
df_mb["2023年12月末"]*0.5)/12

'''
计算上年度的贷款利差补贴
'''
df_mb["上年度的贷款利差补贴"]=df_mb["全年贷款平均余额"]*0.02

'''
实际获得金额
'''
df_mb["实际获得金额"]=df_mb["上年度的贷款利差补贴"]

'''
贷款品种五级
12月的数据到了，源数据要取12月的。
'''
df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202311, how="left", on=["合同号"])
df_mb.loc[:,"贷款品种五级"]=df_mb2["贷款品种5级"].tolist()


'''
额度合同号
12月的数据到了，源数据要取12月的。
'''
df_mb2=df_mb.copy()
df_mb2=df_mb2.merge(df_202311, how="left", on=["合同号"])
df_mb.loc[:,"额度合同号"]=df_mb2["对应额度合同号"].tolist()

with pd.ExcelWriter("处理结果.xlsx") as writer:
     df_mb.to_excel(writer, sheet_name="基于合同号",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
#     df_mb_c8r.to_excel(writer, sheet_name="基于身份证号码",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
#     df_mb_zj2_sfzcf.to_excel(writer, sheet_name="8个特殊人",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")

"""
按身份证号码进行汇总
"""



# df_mb_sfz = df_mb.copy()
# df_mb_sfz = df_mb_sfz.drop(columns="合同号")
# df_mb_sfz = df_mb_sfz.groupby("客户身份证号码").sum()  #共525人


# """
# 重复数据处理
# """
# df_mb_zj = df_mb[["客户身份证号码","客户姓名名称","贷款业务品种","户籍所在地","户籍地址","是否执行先收后返","贷款品种五级"]]
#
# ser = df_mb_zj["客户身份证号码"].duplicated(keep=False)
#
# df_mb_zj2_zcxyc = df_mb_zj.loc[ser==False, :]
# df_mb_zj2_zcxyc = df_mb_zj2_zcxyc.sort_values("客户身份证号码") #230人只出现了一次
#
# df_mb_zj2_ycys = df_mb_zj.loc[ser==True, :]  #包含所有列都重复的行和部分列重复的行（部分不重复）
# ser2 = df_mb_zj2_ycys.duplicated()  #所有列都重复的行只保留第一次出现（以下3行）
# df_mb_zj2_ycys =  df_mb_zj2_ycys.loc[ser2==False,:]
# df_mb_zj2_ycys = df_mb_zj2_ycys.sort_values("客户身份证号码")  #305-18=287人
#
# ser3 = df_mb_zj2_ycys["客户身份证号码"].duplicated(keep=False)
# df_mb_zj2_sfzcf = df_mb_zj2_ycys.loc[ser3==True,:]  #这个很重要啊，一共涉及8个人,这8个人因为产品上的重复占了18行。
# df_mb_zj2_sfzbcf = df_mb_zj2_ycys.loc[ser3==False,:]  #305-18=287人
#
# """为表“基于身份证号码”添加非数字字段"""
# #230人+287人 信息合并上来
# df_zcxyc_hb = pd.concat([df_mb_zj2_zcxyc,df_mb_zj2_sfzbcf],axis=0)
# df_zcxyc_hb["客户身份证号码"] = "_" + df_zcxyc_hb["客户身份证号码"]
# df_mb_sfz = df_mb_sfz.reset_index()
# df_mb_sfz["客户身份证号码"] = "_" + df_mb_sfz["客户身份证号码"]
# df_mb_c8r = df_mb_sfz.merge(df_zcxyc_hb,how="left",on=["客户身份证号码"])


"""
对处理结果进行存储
"""
# df_mb["合同号"] = df_mb["合同号"].astype(np.int64).astype(str)
# df_mb_zj2_sfzcf = df_mb_zj2_sfzcf.sort_values("客户身份证号码")
# with pd.ExcelWriter("处理结果.xlsx") as writer:
#     df_mb.to_excel(writer, sheet_name="基于合同号",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
#     df_mb_c8r.to_excel(writer, sheet_name="基于身份证号码",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
#     df_mb_zj2_sfzcf.to_excel(writer, sheet_name="8个特殊人",startrow=0, startcol=0, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
