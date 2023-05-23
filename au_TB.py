'''
实现TB自动化
审计时
step0:底稿完备
step1:copy last year's TB and rename it with new year flag
step2:手工copy去年数,run macros in VB（清空TB本年数据）
step3: change the parameters in this code and run
'''

import re
import time
import warnings

warnings.filterwarnings("ignore")
import xlwings as xw
import numpy as np
import pandas as pd
from functools import reduce
import math

# alin with
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
# 显示所有列
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', 200)
pd.set_option('max_colwidth', 100)
pd.set_option('display.width', 125)

df_seq = pd.read_excel('D:\work\\202206\\序时账202206.xls', dtype="str")
# print("序时账:\n", df_seq.head())
# 制造科目编号与科目名称字典
df_bs = pd.read_excel('D:\work\\202206\余额表202206.xls')  # balance_statement
print("科目名称:\n", df_bs.tail())

dict_acc = {i: v for i, v in zip(df_bs['科目编码'], df_bs['科目名称'])}
print("科目名称字典:\n", dict_acc)
print(type(dict_acc["资产小计"]))
print(np.isnan(dict_acc["资产小计"]))
print(re.match("\d+", '收入费用小计'))
# dict_acc_rever  = {key: value for key, value in dict_acc .items() if  re.match("\d+",key)}
# dict_acc_rever  = {value: key for key, value in dict_acc .items() if not (isinstance(value, float) and np.isnan(value))}
dict_acc_rever = {value: key for key, value in dict_acc.items() if (isinstance(value, str))}
# dict_acc_rever  = {value: key for key, value in dict_acc .items() if not (math.isnan(value))} # 错误在没有一个一个地运行
print("名称代码字典:\n", dict_acc_rever)
pat = re.compile(r"\d{4}\b")
lst_acc_tier1 = [v for k, v in dict_acc.items() if pat.match(k)]  # 1级科目
print("lst_acc_tier1:\n", lst_acc_tier1)

'''首先要能够解析excel中单元格的含义与位置标志提取'''
tb = "典当行业报表及附注202206.xlsm"
df_tb_bs = pd.read_excel(f'D:\work\\202206\{tb}', sheet_name="B-典当", header=4, index_col=0)
df_tb_bs["审前数"] = [0 for i in range(len(df_tb_bs))]
df_tb_bs["审前数.1"] = [0 for i in range(len(df_tb_bs))]
df_tb_bs.index = list(map(lambda x: x.strip() if type(x) != float else "", df_tb_bs.index))  # delete spaces in string
df_tb_bs["负债和股东权益"] = list(
    map(lambda x: x.strip() if type(x) != float else "", df_tb_bs["负债和股东权益"]))  # delete spaces
print(df_tb_bs.head(10))
df_tb_ic = pd.read_excel(f'D:\work\\202206\{tb}', sheet_name="P-典当", header=5, index_col=0)
df_tb_ic["审前数"] = [0 for i in range(len(df_tb_ic))]
df_tb_ic.index = list(map(lambda x: x.strip() if type(x) != float else "", df_tb_ic.index))  # delete spaces in string

print("df_tb_ic：\n", df_tb_ic)

# print(df_tb_bs.loc["短期投资","负债和股东权益"])
# print(type(df_tb_bs.loc["短期投资","负债和股东权益"]))
# print(len(df_tb_bs.loc["短期投资","负债和股东权益"]))
# print("2:\n",df_tb_bs.head(10))


# for app in xw.apps:
for book in xw.books:
    if tb in book.name:
        wb1 = book
        break
    else:
        wb1 = xw.Book(f'D:\work\\202206\\{tb}')
        time.sleep(5)
print(wb1)
ws_bs = wb1.sheets("B-典当");
ws_ic = wb1.sheets("P-典当");
ws_head = wb1.sheets("表头")

# 表头也可以automation
# year = re.findall('20\d{2}', tb)[0]
# print(year)
# ws_head.range("B5").value = f"AUDIT PERIOD:   {year}.12.31"
# ws_head.range("B6").value = f"截止到{year}-12-31"
# ws_head.range("B7").value = f"{year}年度"
# ws_head.range("B8").value = f"AUDIT PERIOD: {year}年"
# ws_head.range("B9").value = f"{year}年12月31日"
# wb1.save()


# 以字典方式存储
# 获取tb中的科目名称，每个名称对应一个在原excel的位置标志；
# creat a map: from df_bs to df_tb_bs to wb.ws.cell

df_bs_tier1 = df_bs[df_bs["科目编码"].str.match("\d{4}\\b")]
df_bs_tier1 = df_bs_tier1.set_index("科目名称")
# print(df_bs_tier1)
# creat a dict, if the acc name can't be found in the df_tb_bs
dict_unknown_acc = {"现金": "货币资金", "银行存款": "货币资金", "固定资产": "固定资产原价",
                    "累计折旧": "减：累计折旧"}  # 无法在余额表科目与报表科目对应的项目
dict_unknown_acc_ic = {"会费收入": "主营业务收入", "业务活动成本": "日常费用", "其他费用": "财务费用"}

# print("df_tb_bs0:\n",df_tb_bs)
for ind in df_bs_tier1.index:
    print(ind)
    if int(dict_acc_rever[ind][:1]) < 4:
        closing_b = df_bs_tier1.loc[ind, "期末借方"] + df_bs_tier1.loc[ind, "期末贷方"]
    else:
        closing_b = df_bs_tier1.loc[ind, "本期发生借方"]

    if closing_b != 0:
        print(ind, int(dict_acc_rever[ind][:1]), closing_b, )
        if ind in df_tb_bs.index:
            df_tb_bs.loc[ind, "审前数"] = df_tb_bs.loc[ind, "审前数"] + closing_b
        elif ind in list(df_tb_bs["负债和股东权益"]):
            print("in liability group")
            ind_in_TB = df_tb_bs["负债和股东权益"][df_tb_bs["负债和股东权益"].values == ind].index
            df_tb_bs.loc[ind_in_TB, "审前数.1"] = df_tb_bs.loc[ind_in_TB, "审前数.1"] + closing_b
        elif ind in list(df_tb_ic.index):
            df_tb_ic.loc[ind, "审前数"] = df_tb_ic.loc[ind, "审前数"] + closing_b
            print("already in ic_statements", df_tb_ic.loc[ind, "审前数"])
        elif ind in dict_unknown_acc.keys():
            df_tb_bs.loc[dict_unknown_acc[ind], "审前数"] = df_tb_bs.loc[dict_unknown_acc[ind], "审前数"] + closing_b
        elif ind in dict_unknown_acc_ic.keys():
            df_tb_ic.loc[dict_unknown_acc_ic[ind], "审前数"] = df_tb_ic.loc[
                                                                   dict_unknown_acc_ic[ind], "审前数"] + closing_b
        else:
            continue
# print("df_tb_bs:\n",df_tb_bs)
# print("df_tb_ic:\n",df_tb_ic)

# 录入TB
row_number = df_tb_bs.index.get_loc('货币资金')
col_number = df_tb_bs.columns.get_loc('审前数.1')
print(row_number, col_number)

# get the index of cell with numbers in "unaudit" column
loc_assets = [[i, 0] for i in range(len(df_tb_bs["审前数"])) if df_tb_bs["审前数"].iloc[i] != 0]
print(loc_assets)
loc_lia = [[i, 8] for i in range(len(df_tb_bs["审前数.1"])) if df_tb_bs["审前数.1"].iloc[i] != 0]
print(loc_lia)
loc_ic = [[i, 0] for i in range(len(df_tb_ic["审前数"])) if df_tb_ic["审前数"].iloc[i] != 0]
print(loc_ic)

"""
赋值前先对原表进行清空,由macro\clear.bas 执行，又通过VSCODE来编写和执行。
"""
print("'货币资金'在py_df中坐标[2,0],excel中坐标[7,1]或[B8](先列，且从1开始计数)")
# 不同报表，需要清空的单元坐标可能不一样
# ws_bs.range('B1:B17').value = None;ws_bs.range('B1:B17').value = None
# print("df_tb_bs2:\n",df_tb_bs)
for i in loc_assets:
    print(df_tb_bs.iloc[i[0], i[1]])
    ws_bs[i[0] + 5, i[1] + 1].value = df_tb_bs.iloc[i[0], i[1]]
# ws_bs[0, 1].value = 100
for i in loc_lia:
    ws_bs[i[0] + 5, i[1] + 1].value = df_tb_bs.iloc[i[0], i[1]]
for i in loc_ic:
    ws_ic[i[0] + 6, i[1] + 1].value = df_tb_ic.iloc[i[0], i[1]]

'''
对资产情况表二中货币资金明细的填写
step1: clear , done in macor "clearSheet"
step2: call function to realize copy 
'''
def copy_df_excel(str_seek, df, sheet):
    """
    给定一个str参数，“科目_初或末”就能实现从original到destination的拷贝
    give an account and it's time,copy original data in a DF to excel.rng
    :param str2:
    :param df:
    :param rng:
    :return:
    """

    str_seek = str_seek.split("_")
    str_seek2 = str_seek[1]
    str_seek3 = [i for i in df.columns if re.search(f"\w+{str_seek2}\w+", i)] # 含有末字的“年末数”或“期末数”字样
    if df.loc[str_seek[0], str_seek3[0]] != 0:
        value = df.loc[str_seek[0], str_seek3[0]]
    else:
        value = df.loc[str_seek[0], str_seek3[0]]
    print(value)

    # unused
    def FindRowCol(Sheet, RowOrCol, KeyWord):
        try:
            if RowOrCol == 'Row':
                Cell_Address = Sheet.api.Cells.Find(What=KeyWord, After=Sheet.api.Cells(Sheet.api.Rows.Count,
                                                                                        Sheet.api.Columns.Count),
                                                    LookAt=xw.constants.LookAt.xlWhole,
                                                    LookIn=xw.constants.FindLookIn.xlFormulas,
                                                    SearchDirection=xw.constants.SearchDirection.xlNext,
                                                    MatchCase=False).Row
            elif RowOrCol == 'Col':
                Cell_Address = Sheet.api.Cells.Find(What=KeyWord, After=Sheet.api.Cells(Sheet.api.Rows.Count,
                                                                                        Sheet.api.Columns.Count),
                                                    LookAt=xw.constants.LookAt.xlWhole,
                                                    LookIn=xw.constants.FindLookIn.xlFormulas,
                                                    SearchDirection=xw.constants.SearchDirection.xlNext,
                                                    MatchCase=False).Column
        except:
            Cell_Address = 0
        return Cell_Address    # 



    lst_possible_col = ["期" + str_seek[1] + "数", "年" + str_seek[1] + "数", "期" + str_seek[1] + "余额",
                        "年" + str_seek[1] + "余额"]
    for i in lst_possible_col:
        anchor = sheet.range("A1:X100").api.Find(i)
        if anchor:
            destination_cell_col = re.findall("\$(.+?)\$", anchor.Address)[0]
            break

    print()
    lst_possible_row = [" " * i + str_seek[0] for i in range(7)]
    print(lst_possible_row)
    for i in lst_possible_row:
        anchor2 = sheet.range("A1:X100").api.Find(i)
        if anchor2:
            print(anchor2.Address)
            destination_cell_row = re.findall("\$(\d+)", anchor2.Address)[0]
            break
    print(destination_cell_row, destination_cell_col)
    sheet.range(destination_cell_col+destination_cell_row).value=value

copy_df_excel("现金_末", df_bs_tier1, wb1.sheets("表二"))
copy_df_excel("银行存款_末", df_bs_tier1, wb1.sheets("表二"))





wb1.save()
