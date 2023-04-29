import serial  # 引用pySerial模組
def water():
    COM_PORT = 'COM3'    # 指定通訊埠名稱
    BAUD_RATES = 9600    # 設定傳輸速率
    ser = serial.Serial(COM_PORT, BAUD_RATES)   # 初始化序列通訊埠
    try:
        message = ""
        while ser.in_waiting:          # 若收到序列資料…
            #data_raw =   # 讀取一行
            data_raw = ser.readline()  # 用預設的UTF-8解碼
            data = data_raw.decode()
            print('原本接收到的資料',data_raw)
            print('接收到的資料：', data)# test
            """if data[0] == '0':
                message = '水庫A的水高於水庫B'
            elif data[0] == '1':
                message = '水庫A的水低於水庫B'
            elif data[0] == '2':
                message = '水庫A的水等於水庫B'"""
    except KeyboardInterrupt:
        ser.close()    # 清除序列通訊物件
        print('exit')
while True:
    water()