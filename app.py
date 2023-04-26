import requests
import json
import pandas as pd
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
    MessageEvent, TextMessage, TextSendMessage,FollowEvent,
    TemplateSendMessage, ButtonsTemplate, PostbackTemplateAction,PostbackEvent,MessageTemplateAction,
    PostbackAction,
    QuickReply, QuickReplyButton, MessageAction
)

app = Flask(__name__)

# 填入你的 LINE 機器人 Channel Access Token
line_bot_api = LineBotApi('n2zf4JW8D/HFwIAL3ekM1/nF4fNEZBKMuSRBgcV/EZsDQCjeNm9PidcivDpnGUdtowQ/0mUtO0wz7+Vlbpr9eRFN+fAOffCmA6qUY6I+yE5JR208KfnqgxFhzh9mTxClVR7XaRE9AHmMz+0p+GEQrgdB04t89/1O/w1cDnyilFU=')

# 填入你的 LINE 機器人 Channel Secret
handler = WebhookHandler('0bb28ec8efa362b936c89246c51ec1c9')

# 設定 Webhook URL
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

# 讀取 Excel 檔案中的資料
df = pd.read_excel("test.xlsx", sheet_name= '工作表1')

num_rows, num_cols = df.shape

#使用者加入Bot後 傳送歡迎訊息 訊息要改成歡迎+自我介紹+功能說明
@handler.add(FollowEvent)
def handle_follow(event):
    user_id = event.source.user_id
    message = TextSendMessage(text="歡迎加入我的 LINE Bot！")
    line_bot_api.push_message(user_id, message)

# 當接收到 MessageEvent (文字訊息) 時的處理函式
@handler.add(MessageEvent, message=TextMessage)#根據取得數據需要去做修改
def handle_message(event):
    # 取得使用者傳來的文字
    
    user_message = event.message.text

    if'所有' in  event.message.text: 
        message = ""
        for i ,row_data in df.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n-------------\n"
        line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))

    if '超標工廠' in event.message.text:
        # 過濾 DataFrame 中排放量大於等於 30 的工廠數據
        df_over_limit = df.query('碳排放 >= 30')
        # 設置 message 變數為空字符串
        message = ""
        # 遍歷 df_over_limit 中的每行數據
        for i, row_data in df_over_limit.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n排放量: {row_data['碳排放']}\n-------------\n"
        # 如果 message 不為空，則將其推送到使用者的 LINE 帳號中
        if message:
            line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))
    buttons_template = None
    if event.message.text == "menu":
        # create a button template message
        buttons_template = ButtonsTemplate(
            title="請選擇",
            text="您要選擇全部工廠還是選擇縣市？",
            actions=[
                PostbackTemplateAction(label="全部工廠", data="action=show_all_factories" ,text= '選擇工廠'),
                PostbackTemplateAction(label="選擇縣市",data = "action=select_city" ,text= '選擇縣市')
            ]
        )
        template_message = TemplateSendMessage(
            alt_text="請在手機上開啟此訊息",
            template=buttons_template
        )
        line_bot_api.reply_message(event.reply_token, template_message)
city = ['台北市','新北市']

@handler.add(PostbackEvent)
def handle_postback(event):
    if event.postback.data == "action=show_all_factories":
        message = ""
        for i ,row_data in df.iterrows():
            message += f"工廠名稱: {row_data['工廠']}\n碳排放: {row_data['碳排放']}\n超標: {row_data['超標']}\n-------------\n"
        line_bot_api.push_message(event.source.user_id, TextSendMessage(text=message))


    elif event.postback.data == 'action=select_city':
        buttons_template = ButtonsTemplate(
            title="請選擇",
            text="您要選擇哪個縣市？",
            actions=[
                PostbackTemplateAction(label="台北市",data= 'a',text= '台北市'),
                PostbackTemplateAction(label="新北市",data= 'b',text= '新北市'),
                PostbackTemplateAction(label="桃園市",data= 'c',text= '桃園市'),
                PostbackTemplateAction(label="台中市",data= 'd',text= '台中市')
            ]
        )
        template_message = TemplateSendMessage(
            alt_text="請在手機上開啟此訊息",
            template=buttons_template
        )
        line_bot_api.reply_message(event.reply_token, template_message)



    else:
        ci = event.postback.data
        message = ""
        for i ,row_data in df.iterrows():
            line_bot_api.reply_message(event.reply_token,TextSendMessage(text = f"{row_data['城市']}"))
    #line_bot_api.reply_message(event.reply_token, TextSendMessage(text= getMsg(event.postback))
if __name__ == "__main__":
    # 執行 Flask Server
    app.run()



"""elif event.message.text == '選擇縣市':
        buttons_template = ButtonsTemplate(
            title="請選擇",
            text="您要選擇哪個縣市？",
            actions=[
                PostbackTemplateAction(label="台北", data="台北" ,text= '台北'),
                PostbackTemplateAction(label="新北",data = "新北" ,text= '新北')
            ]
        )
        template_message = TemplateSendMessage(
            alt_text="請在手機上開啟此訊息",
            template=buttons_template
        )"""