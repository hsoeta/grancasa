import pyautogui as pag
import time
import datetime
import pyperclip as pc
import schedule

##PC版のラインを使用する想定で作成. mouse.pyで目的の座標を取得して編集してください.
def cycle(pattern):
    if pattern == 1:
        x, y = 100, 200 ##最初にクリックしたい場所 : クリップマーク (=ファイル送信)
    elif pattern == 2:
        x, y = 300, 400　##2番目にクリックしたい場所 : 送信したいファイル
    elif pattern == 3:
        x, y = 500, 600　##3番目にクリックしたい場所 : "開く"ボタン
    elif pattern == 4:
        x, y = 700, 800　##4番目にクリックしたい場所 : "メッセージを入力"と表示されている領域ならどこでもOK
    else:
        raise ValueError("Invalid pattern")
    time.sleep(1)
    pag.click(x, y)

def paste():
    time.sleep(2)
    pag.hotkey("command", "v")
    time.sleep(2)
    pag.press("enter")

def send():
    #ファイル選択メニュー
    cycle(1)
    #送信ファイル選択
    cycle(2)
    #決定ボタン
    cycle(3)
    #メッセージ入力フィールド
    cycle(4)
    time.sleep(1)
    pc.copy('本日の宿泊者予定です。')
    paste()
    pc.copy(plan_today)
    paste()
    pc.copy(prob_yester)
    paste()
    pc.copy(next_person)
    paste()

def do_something():
    print("ジョブを実行しました。")

print("                              ")
send_time    = input("送信予定時刻       : ")
exe_time     = datetime.datetime.strptime(send_time, "%H:%M")
plan_today   = input("宿泊プランについて : ")
prob_yester  = input("報告事項について   : ")
next_person  = input("この後も自分なのか : ")
print("                              ")
print("------------------------------")
print("                              ")
print("{0}:{1} に実行予定です。" .format(exe_time.hour, exe_time.minute))

print("          ")

#schedule.every().day.at(exe_time.strftime("%H:%M")).do(do_something)

schedule.every().day.at(exe_time.strftime("%H:%M")).do(send)

while True:
    schedule.run_pending()
    time.sleep(1)
