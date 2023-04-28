import requests
import json
import schedule
import time
import pandas as pd
import Tokens as t
from flask import Flask, request, abort
from linebot import (
    LineBotApi,
    WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError,
    LineBotApiError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,FollowEvent,UnfollowEvent,
    TemplateSendMessage, ButtonsTemplate, PostbackTemplateAction,PostbackEvent,MessageTemplateAction,
    PostbackAction,
    QuickReply, QuickReplyButton, MessageAction
)

app = Flask(__name__)

line_bot_api = LineBotApi(t.Token)
handler = WebhookHandler(t.Key)

@app.route("/callback", methods=['POST'])
def callback():
    # 取得 Line Server 回傳的 Header 資料
    signature = request.headers['X-Line-Signature']
    # 取得 POST 資料
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    try:
        # 驗證資料是否正確
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

#使用者加入Bot後 傳送歡迎訊息 訊息要改成歡迎+自我介紹+功能說明
@handler.add(FollowEvent)
def handle_follow(event):
    user_id = event.source.user_id
    message = TextSendMessage(text="歡迎加入Line 廢水監控機器人！\n輸入「所有」查看所有工廠數據\n輸入「超標工廠」查看所有超標工廠數據\n輸入「menu」查看各縣市工廠數據")
    line_bot_api.push_message(user_id, message)
    df = pd.read_excel('friendList.xlsx',sheet_name= 'Sheet1')
    if user_id not in df['user_id'].tolist():
        # 加入新的好友# 取得新加入好友的user_id

        # 檢查是否已經加入過好友
        if user_id in df['user_id'].values:
            print('user_id已存在')
        else:
            # 將user_id新增至Excel
            new_user = {'user_id': user_id}
            df = df.append(new_user, ignore_index=True)

            # 儲存更新後的Excel檔案
            df.to_excel('friendList.xlsx', index=False)
            print('已新增user_id:', user_id)

#刪除好友後將user_id從excel中移除
@handler.add(UnfollowEvent)
def handle_unfollow(event):
    user_id = event.source.user_id
    # 讀取現有的Excel檔案
    df = pd.read_excel('friendList.xlsx')

    # 檢查user_id是否存在
    if user_id not in df['user_id'].values:
        print('user_id不存在')
    else:
        # 找出要刪除的user_id所在的索引位置
        index = df[df['user_id'] == user_id].index

        # 刪除該索引位置的資料
        df.drop(index, inplace=True)

        # 重新編號索引
        df.reset_index(drop=True, inplace=True)

        # 儲存更新後的Excel檔案
        df.to_excel('friendList.xlsx', index=False)
        print('已刪除user_id:', user_id)
        
df1 = pd.read_excel("test.xlsx", sheet_name= '工作表1')
df2 = pd.read_excel('compare.xlsx', sheet_name='工作表1')

# 當接收到 MessageEvent (文字訊息) 時的處理函式
@handler.add(MessageEvent, message=TextMessage)#根據取得數據需要去做修改
def handle_message(event):
    user_message = event.message.text
    if'所有' in  event.message.text: 
        message = ""
        for i ,row_data in df1.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))

    if '超標工廠' in event.message.text:
        df1_over_limit = df1.query('碳排放 >= 30')
        message = ""
        for i, row_data in df1_over_limit.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        if message:
            line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))
    buttons_template = None
    if event.message.text == "menu":
        buttons_template = ButtonsTemplate(
            title="請選擇",
            text="您要選擇全部工廠還是選擇縣市？",
            actions=[
                PostbackTemplateAction(label="全部工廠", data="action=show_all_factories"),
                PostbackTemplateAction(label="選擇縣市",data = "action=select_city")
            ]
        )
        template_message = TemplateSendMessage(
            alt_text="請在手機上開啟此訊息",
            template=buttons_template
        )
        line_bot_api.reply_message(event.reply_token, template_message)

@handler.add(PostbackEvent)
def handle_postback(event):
    if event.postback.data == "action=show_all_factories":
        message = ""
        for i ,row_data in df1.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))

    elif event.postback.data == 'action=select_city':
        buttons_template = ButtonsTemplate(
            title="請選擇",
            text="您要選擇哪個縣市？",
            actions=[
                PostbackTemplateAction(label="台北市",data= '台北'),
                PostbackTemplateAction(label="新北市",data= '新北'),
                PostbackTemplateAction(label="桃園市",data= '桃園'),
                PostbackTemplateAction(label="台中市",data= '台中')
            ]
        )
        template_message = TemplateSendMessage(
            alt_text="請在手機上開啟此訊息",
            template=buttons_template
        )
        line_bot_api.reply_message(event.reply_token, template_message)
    else:
        message = ""
        ci = event.postback.data
        for i ,row_data in df1.iterrows():
            if ci == row_data['城市']:
                message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))

def check_data():
    diff = df1.compare(df2)#照寫
    if not diff.empty:
        df1_over_limit = df1.query('碳排放 >= 30')
        message = ""
        for i, row_data in df1_over_limit.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n城市: {row_data['城市']}\n-------------\n"
        df = pd.read_excel('friendList.xlsx',sheet_name= 'Sheet1')
        for i ,row_data in df.iterrows():
            line_bot_api.push_message(row_data['user_id'], TextSendMessage(text=message))
        
        # 將 DataFrame 儲存到 Excel 文件中
        with pd.ExcelWriter('compare.xlsx') as writer:
            df1.to_excel(writer, sheet_name='工作表1', index=False)
schedule.every(10).seconds.do(check_data)

if __name__ == "__main__":
    app.run()
    schedule.run.pending()
    time.sleep(1)