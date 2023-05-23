

import xlwings as xw
import pandas as pd;import re
df_bs = pd.read_excel('D:\work\\2023\余额表2022.xls')  # balance_statement
print("科目名称:\n", df_bs.tail())
df_bs_tier1 = df_bs[df_bs["科目编码"].str.match("\d{4}\\b")]
df_bs_tier1 = df_bs_tier1.set_index("科目名称")
str_seek = "现金_末";str_seek=str_seek.split("_")
str_seek2=str_seek[1]
str_seek3=[i for i in df_bs_tier1.columns if re.search(f"\w+{str_seek2}\w+",i)]
if df_bs_tier1.loc[str_seek[0],str_seek3[0]]!=0:
    value = df_bs_tier1.loc[str_seek[0],str_seek3[0]]
else: value = df_bs_tier1.loc[str_seek[0],str_seek3[0]]
print(value)



# def FindRowCol(Sheet, RowOrCol, KeyWord):
#     try:
#         if RowOrCol == 'Row':
#             Cell_Address = Sheet.api.Cells.Find(What=KeyWord, After=Sheet.api.Cells(Sheet.api.Rows.Count, Sheet.api.Columns.Count), LookAt=xw.constants.LookAt.xlWhole,
#                                                     LookIn=xw.constants.FindLookIn.xlFormulas, SearchDirection=xw.constants.SearchDirection.xlNext, MatchCase=False).Row
#         elif RowOrCol == 'Col':
#             Cell_Address = Sheet.api.Cells.Find(What=KeyWord, After=Sheet.api.Cells(Sheet.api.Rows.Count, Sheet.api.Columns.Count), LookAt=xw.constants.LookAt.xlWhole,
#                                                     LookIn=xw.constants.FindLookIn.xlFormulas, SearchDirection=xw.constants.SearchDirection.xlNext, MatchCase=False).Column
#     except:
#         Cell_Address = 0
#     return Cell_Address


print("----------------")
tb = "典当行业报表及附注2022_2.xlsm"
wb1=xw.Book(f'D:\work\\2023\\{tb}')
ws_sheet2 = wb1.sheets("表二")

lst_possible_col = ["期"+str_seek[1]+"数","年"+str_seek[1]+"数","期"+str_seek[1]+"余额","年"+str_seek[1]+"余额"]
print(lst_possible_col)
for i in lst_possible_col:
    print(i)
    # anchor_col = FindRowCol(ws_sheet2, "Col", i)
    anchor = ws_sheet2.range("A1:X100").api.Find(i)
    print(anchor)
    if anchor:
        destination_cell_col=re.findall("\$(.+?)\$",anchor.Address)[0]
        break

print(destination_cell_col)

lst_possible_row = [" "*i+str_seek[0] for i in range(7)]
print(lst_possible_row)

for i in lst_possible_row:
    # print(i)
    # anchor_col = FindRowCol(ws_sheet2, "Col", i)
    anchor2 = ws_sheet2.range("A1:X100").api.Find(i)
    # print(anchor2)
    if anchor2:
        destination_cell_row=re.findall("\$(\d+?)",anchor2.Address)[0]
        break
print(destination_cell_row)