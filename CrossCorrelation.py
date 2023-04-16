import pandas as pd
import numpy as np
import xlwings as xw
comparefile1 = 'C:/Users/ArvindSanala/PycharmProjects/pythonProject/WeightStatusReports/WeightInputReport20230131-KW 07.xlsx'
comparefile2 = 'C:/Users/ArvindSanala/PycharmProjects/pythonProject/WeightStatusReports/WeightInputReport20230206-KW 07.xlsx'
df1 = pd.read_excel(comparefile1)
df2 = pd.read_excel(comparefile2)
df2 = df2.reset_index()
df2.head(3)
df_diff =pd.merge(df1,df2,how="outer", indicator="Exist")
df_diff = df_diff.query("Exist != 'both'")
df_highlight =df_diff.query("Exist == 'right_only'")
highlight_rows =df_highlight['index'].tolist()
highlight_rows = [int(row) for row in highlight_rows]
first_row_in_excel = 2
highlight_rows =[x + first_row_in_excel for x in highlight_rows]
with xw.App(visible=False) as app:
 updated_wb = app.books.open(comparefile2)
 updated_ws = updated_wb.sheets(1)
 rng = updated_ws.used_range
 print(f"Used Range: {rng.address}")
 for row in rng.rows:
   if row.row in highlight_rows:
      row.color = (255, 71, 76)
 updated_wb.save("finalcomp.xlsx")
