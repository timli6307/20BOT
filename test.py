"""import pandas as pd

# 讀取 Excel 檔案，並將其轉換為 DataFrame
df = pd.read_excel('friendList.xlsx')

# 更改 DataFrame 中的值
df.at[len(df),0] = input('input data: ')

# 將更改後的 DataFrame 寫回到 Excel 檔案
df.to_excel('friendList.xlsx', index=False)
"""
import pandas as pd

# 讀取 Excel 檔案，並將其轉換為 DataFrame
df = pd.read_excel('friendList.xlsx')

# 將資料表中的空白值填補為上一個非空白值
df.fillna(method='ffill', inplace=True)

# 將整理好的資料寫回到 Excel 檔案中
df.to_excel('', index=False)