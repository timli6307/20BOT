import pandas as pd

# 讀取Excel檔案
df1 = pd.read_excel('test.xlsx', sheet_name='工作表1')
df2 = pd.read_excel('test.xlsx', sheet_name='工作表2')

# 比較兩個sheet的差異
diff = df1.compare(df2)

# 如果有差異，將Sheet2的資料改寫成Sheet1的資料
if not diff.empty:
    df2.update(df1)

    # 將結果寫入新的Excel檔案
    with pd.ExcelWriter('new_file.xlsx') as writer:
        df2.to_excel(writer, sheet_name='Sheet2', index=False)