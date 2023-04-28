import pandas as pd
df1 = pd.read_excel("test.xlsx", sheet_name= '工作表1')
df2 = pd.read_excel('compare.xlsx', sheet_name='工作表1')
def check_data():
    global df1
    diff = df1.compare(df2)#照寫
    if not diff.empty:
        df1_over_limit = df1.query('碳排放 >= 30')
        message = ""
        for i, row_data in df1_over_limit.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        df = pd.read_excel('friendList.xlsx',sheet_name= 'Sheet1')
        for i ,row_data in df.iterrows():
            print(message)
        
        # 將 DataFrame 儲存到 Excel 文件中
        with pd.ExcelWriter('compare.xlsx') as writer:
            df1.to_excel(writer, sheet_name='工作表1', index=False)
