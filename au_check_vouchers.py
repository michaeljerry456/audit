"""
实现对”工作底稿“的编写
根据序时账，按规则选举样本后录入底稿中
暂未完成（以后年度）：
    期初期末数据的填列；


"""
import re
import time
import warnings

warnings.filterwarnings("ignore")
import xlwings as xw
import numpy as np
import pandas as pd
from functools import reduce

# alin with
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
# 显示所有列
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', 200)
pd.set_option('max_colwidth', 100)
pd.set_option('display.width', 125)


df_seq = pd.read_excel('D:\work\\202206\\序时账202206.xls', dtype="str")
print("序时账:\n", df_seq.head())
# 制造科目编号与科目名称字典
df_bs = pd.read_excel('D:\work\\202206\余额表202206.xls')  # balance_statement
print("科目名称:\n", df_bs.head())

dict_acc = {i: v for i, v in zip(df_bs['科目编码'], df_bs['科目名称'])}
print("科目名称字典:\n", dict_acc)
pat = re.compile(r"\d{4}\b")
lst_acc_tier1 = [v for k, v in dict_acc.items() if pat.match(k)]  # 1级科目
print("lst_acc_tier1:\n", lst_acc_tier1)


# 根据科目名称抽查凭证形成DF
def sel_L_tran(df,df_seq, acc="", quan=3):
    # 将序时账化简为拟选择对象
    df = df[df['科目名称'] == acc]
    # print(1, "\n", df)
    df = df.sort_values(['方向', '金额'], ascending=False)
    df = df.groupby(['方向']).head(quan)  # 借贷方各取3个
    # print(2,"\n",df)
    df = df[~(df['摘要'] == '期末结转')]
    print(2, "\n", df)

    lst_opp_acc = []
    for i in range(len(df)):
        # df_add=df_seq[(df_seq['日期']==df['日期'].iloc[i]) and (df_seq['凭证号数']==df['凭证号数'].iloc[i])]
        df_add = df_seq[(df_seq['日期'] == df['日期'].iloc[i])][(df_seq['凭证号数'] == df['凭证号数'].iloc[i])]
        print("df_add", df_add)
        # df=pd.concat([df,df_add])
        print(3, "\n", df)
        df_add = df_add[df_add['方向'] != df['方向'].iloc[i]]
        print("df_add",df_add)
        opp_acc = list(map(lambda x: x[:4] if len(x) > 4 else x, df_add['科目编码']))
        opp_acc = ','.join(list(map(lambda x: dict_acc[x], list(set(opp_acc)))))
        # 会费收入22.9-6# 对应5501科目，在科目余额表中找不到，所以报错
        lst_opp_acc.append(opp_acc)
        print(opp_acc)
    df['opp_acc'] = lst_opp_acc
    df.rename(columns={'外币': '金额_dr', '金额': '金额_cr'}, inplace=True)
    print(df)
    print("按底稿格式调整df")
    df['金额_dr'] = df.apply(lambda x: x['金额_cr'] if x['方向'] == '借' else x['金额_dr'], axis=1)
    df['金额_cr'] = df.apply(lambda x: x['金额_cr'] if x['方向'] == '贷' else 0, axis=1)
    df = df[['日期', "凭证号数", "摘要", "科目名称", "opp_acc", "金额_dr", "金额_cr"]]
    df.insert(1, "cate", [np.nan] * len(df))

    return df


# 业务活动成本

working_paper = '典当协会审计底稿202206'

'''操作EXCEL'''
'''此等方法控制一个已经打开的excel'''
# for app in xw.apps:
for book in xw.books:
    if working_paper in book.name:
        wb1 = book
        break
    else:
        wb1 = xw.Book(f'D:\work\\202206\\{working_paper}.xlsx')
        time.sleep(5)
print(wb1)


# app = xw.apps[57316]
# wb1 = app.books('典当协会审计底稿_2023.xlsx')

# app = xw.App(visible=True, add_book=False)
# wb1 = app.books.open('D:\典当协会审计底稿-2021_3.xlsx')


def input_wp(df_seq, df_bs, df_wp, check_acc=""):
    num_row = np.where(df_wp == '日期')[0][0]
    num_row2 = np.where(df_wp == '审计说明：')[0][0]
    num_col = np.where(df_wp == '备注')[1][0]

    print(num_row, num_row2, num_col)
    # todo 是否为参数
    ws1 = wb1.sheets['CP-' + check_acc]
    print(ws1.name)
    # 在xw中，+1，表示df取表默认第一行为索引，再+1，df默认第一行为0行，再+1，“日期”本身为合并2行，再+1，要执行的xw单元格所在行
    ws1.range((num_row + 4, 2), (num_row2 + 1, num_col - 5)).clear()  #

    # 如果本科目本期无发生额则返回
    num_row3 = np.where(df_bs['科目名称'] == check_acc)[0][0]
    if int(df_bs['本期发生借方'].iloc[num_row3]) + int(df_bs['本期发生贷方'].iloc[num_row3]) == 0:
        return
    else:
        if int(df_bs['本期发生借方'].iloc[num_row3]) > int(df_bs['本期发生贷方'].iloc[num_row3]):
            sel_side = '本期发生借方'
        else:
            sel_side = '本期发生贷方'

    # 选择抽凭科目
    # print(df_bs)
    # opp_dict_acc = {v: k for k, v in dict_acc.items()}  # 反转字典key and value
    # code = opp_dict_acc[check_acc]
    # print("checking acc_tier_1:",code)
    code = next((k for k, v in dict_acc.items() if v == check_acc), None)
    print("selected account:", code)
    df_bs_tier2 = df_bs[df_bs['科目编码'].str.match(f"\\b{code}")]  # 该科目的二级明细科目余额表
    # print("df_bs_tier2\n", df_bs_tier2)
    df_bs_tier2 = df_bs_tier2.sort_values([sel_side], ascending=[False])
    print("df_bs_tier2\n", df_bs_tier2)
    selected_acc = list(df_bs_tier2[sel_side])
    sum_acc = reduce(lambda x, y: x + y, selected_acc)
    if sum_acc > selected_acc[0] * 1.8:  # 如果有明细科目,选择发生额较大的明细
        s = selected_acc[1]
        m = 1
        while s < selected_acc[0] * 0.6:
            m += 1
            s = s + selected_acc[m]
        selected_acc = [df_bs_tier2['科目名称'].iloc[i + 1] for i in range(m)]
    else:
        selected_acc = [df_bs_tier2['科目名称'].iloc[0]]
    selected_acc = list(map(lambda x: x.strip(), selected_acc))
    print("selected_acc",selected_acc)

    # 执行抽凭,实现了多个二级明细的抽查
    quan = 3  # 单方向抽凭数量
    df_sel_L_tran = pd.DataFrame()
    for acc in selected_acc:
        df_temp = sel_L_tran(df_seq,df_seq, acc=acc, quan=quan)
        df_sel_L_tran = pd.concat([df_sel_L_tran, df_temp])

    print(f"以下为{check_acc}的抽凭内容：\n")
    print(df_sel_L_tran)

    # app.quit()
    # 抽查底稿中如果行数不够增加行数
    # len(df_sel_L_tran) 是抽查的数量 +2 是富裕，(num_row2-num_row-2)+2 是实际空白行数，num_row本身是合并单元
    if len(df_sel_L_tran) + 2 - (num_row2 - num_row - 2) > 0:
        _ = [ws1.api.Rows(num_row2 + i).Insert() for i in range(len(df_sel_L_tran) + 2 - (num_row2 - num_row - 2))]

    #     for i in range(len(df_sel_L_tran) + 2 - (num_row2 - num_row - 2)):
    #         ws1.api.Rows(num_row2).Insert()

    # 录入数据
    for i, v in enumerate(df_sel_L_tran.index):
        # print(list(df_sel_L_tran.loc[v]))
        ws1.range(num_row + 4 + i, 2).value = list(df_sel_L_tran.loc[v])

    wb1.save()
    time.sleep(2)
    # return

print(lst_acc_tier1[15:])
for acc in lst_acc_tier1[15:]: # thresh
    print(acc)
    if acc in ['累计折旧','非限定性净资产']: continue
    df_wp = pd.read_excel(f'D:\work\\2023\\{working_paper}.xlsx', sheet_name=f'CP-{acc}')
    input_wp(df_seq, df_bs, df_wp, check_acc=acc)
    print("over")

wb1.close()
