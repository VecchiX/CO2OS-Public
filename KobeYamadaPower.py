## 本ツールについて
##
## (1)次のコマンドでEXE化して配布すること
##  pyinstaller --onefile --exclude-module numpy --exclude-module pandas KobeYamadaPower.py
##
## (2)「CONFIG(神戸山田).xlsx」と一緒に配布すること

import time
import datetime
import pyautogui as pa
import signal
import os

print("\n\n神戸山田発電量 2024.04.24版\n")

## 起動画面選択 ##
while True:
    print("起動する画面を次の中から選んでください")
    print("1.数値確認のための画面")
    print("2.一覧監視画面(アラート確認へ)")
    print("3.グラフ")

    user_choice = input("番号を入力してください (1, 2, 3): ")

    if user_choice in ["1", "2", "3"]:
        break
    else:
        print("無効な選択です。1～3の数字を入力してください。")

user_choice = int(user_choice)

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from FUNCTIONS import scroll

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from CONFIGxlsx import dictionary as Conf
from CONFIGxlsx import Get_Info

#### メイン ####
now = datetime.datetime.now()
print(now.strftime("%Y/%m/%d %H:%M"))


pa.press("esc")
time.sleep(5)

Site_Message =""
fullpath = []

# try:

## Excelからの情報取得 ##
Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
print(Site_Message)

# 選択肢に応じて読み込むシートを切り替える
Conf_Sheet_name=["","発電量","一覧監視画面","グラフ"]
ret = Get_Info(Conf_Sheet_name[user_choice], "CONFIG(神戸山田).xlsx")
if ret == False:
    raise Exception


## 監視用Webサイトに接続 ##
Site_Message = '＜Laplace Systemにログイン＞'
print(Site_Message)
driver = initialize_driver()
Open_Site(Conf['URL'])


## 監視用Webサイトにログイン ##
Site_Message = '＜ログイン＞'
print(Site_Message)
Login(Conf, "username", "password", "", By.NAME, By.NAME)


## 監視用画面を開く ##
Site_Message = '＜監視用画面を開く＞'
print(Site_Message)
num_win = 0
for key, url in Conf.items():
    if key.find('URL') == 0:
        if key != 'URL':
            driver.execute_script('window.open("'+url+'");')
            num_win += 1 
            time.sleep(2)

dt = datetime.datetime.now()
dt += datetime.timedelta(hours=-6)
dt2 = dt + datetime.timedelta(days=-6)

# 選択肢に応じた処理
for num in range(num_win):
    handle = driver.window_handles[num + 1]
    driver.switch_to.window(handle)

    # 数値確認のとき
    if user_choice == 1:
        Click_Element('1日','//input[@id="sizeDayId"]', 1)
        obj = driver.find_element(By.XPATH, '//input[@value="monthly"]')
        if not obj.is_selected():               
            raise(Exception)
        del obj

        obj = driver.find_element(By.XPATH, '//select[@name="startYear"]')
        select = Select(obj)
        select.select_by_value(str(dt.year))

        obj = driver.find_element(By.XPATH, '//select[@name="startMonth"]')
        select = Select(obj)
        select.select_by_value(str(dt.month))
    
        ## ↓↓　メモリ不足になってしまうときはコメントアウト

        time.sleep(3)
        Click_Element('データ表示', '//*[@id="set"]', 10)

        scroll()

        iframe = driver.find_element(By.TAG_NAME,'iframe')
        driver.switch_to.frame(iframe)
        if dt.day > 12:
            dtext = str(dt2.month) + '/' + str(dt2.day)
            obj = driver.find_element(By.XPATH, '//td[text()="' + dtext + '"]')
            driver.execute_script('arguments[0].scrollIntoView();', obj)

        ## ↑↑　メモリ不足になってしまうときはコメントアウト

driver.switch_to.window(handle)

os.kill(driver.service.process.pid,signal.SIGTERM)