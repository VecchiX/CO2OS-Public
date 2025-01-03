import datetime
import pyautogui as pa

import time
import sys

from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

from FUNCTIONS import Login, Click_Element, initialize_driver, Open_Site
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By

driver = None

## 監視画面設定 ##
def Site_Initialize(Config_Sheet_name=''):
    global driver

    print("\n\nGPM共通処理 Ver.2024.12.13.磯部\n")

    try:

        now = datetime.datetime.now()
        print(now.strftime("%Y/%m/%d %H:%M"))

        pa.press("esc")
        time.sleep(5)

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)
        args = sys.argv
        print(args)
        count = len(args) - 1

        # GPM情報取得 
        if Config_Sheet_name == '':
            ret = Get_Info(args[1])
        else:
            ret = Get_Info(Config_Sheet_name)

        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception
                

        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['URL'])


        ## 監視用Webサイトにログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)

        Login(Conf, "_username", "_password", "submit_show", By.NAME, By.NAME)
        time.sleep(35)

        ## GPM スマート分析画面を表示 ##
        Site_Message = "＜スマート分析画面を表示＞"
        print(Site_Message)

        driver.get(Conf['スマート分析URL'])
        time.sleep(45) #45


        ## 日本語/英語チェック ##
        Site_Message = "＜言語チェック＞"
        print(Site_Message)
        elm = driver.find_element(By.XPATH, '(//span[@class="navbar-brand ng-binding"])')
        lang_chk = elm.text

        if lang_chk.find("スマート分析") == -1:

            Site_Message = "＜日本語に変更＞"
            print(Site_Message)

            # 言語指定
            Click_Element('User', '//span[@class="user-info ng-binding"]', 2)
            Click_Element('Setting', '//a[contains(@href, "javascript")]/span[@class="hidden-desktop"]', 2)

            elm = driver.find_element(By.ID,'userLocale')
            elm.click()
            select = Select(elm)
            select.select_by_value('string:ja')
            time.sleep(5)

            Click_Element('Save', '(//button[contains(text(),"Save")])[1]', 10)

    except Exception as e:
        if Site_Message:
            Site_Message += "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"

        print(f"【GPM共通処理異常】：" + Site_Message + "\r\n" )
    
    finally:
        return driver


## 日時指定 ##
def SetDate(date_from):
    global driver

    try:
        # 過去日が指定されたら、指定された日の12:00を選択する
        if date_from != "":
            Site_Message = "＜日付指定＞"
            print(Site_Message)
            Click_Element('日付指定','//input[@class="form-control form-control-sm"]', 2)

            # 年月指定
            while True:
                elm = driver.find_element(By.XPATH, '(//th[@class="datepicker-switch"])[1]')
                date_text = elm.text
                date_text = date_text.split()
                month_text = date_text[0].replace('月', '')
                month_text = month_text.zfill(2)
                date_text = date_text[1] + '-' + month_text

                if (date_text in date_from):
                    break
                else:
                    Click_Element('月指定', '(//th[@class="prev"])[1]', 2)

            # 日指定
            day = str(int(date_from.split('-')[2]))
            Click_Element('日指定', '//td[@class="day" and text()="' + day + '"]', 2)

            # 過去日指定時は、時刻(12:00)固定
            Click_Element('時刻指定', '//select[@ng-model="ctrl.hour"]', 2)
            Click_Element('12:00', '//option[@value="number:12"]', 2)

    except Exception as e:
        if Site_Message:
            Site_Message += "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"

        print(f"【GPM共通処理異常】：" + Site_Message + "\r\n" )
        raise 


## ログアウト ##
def logout():
    global driver

    try:
        Site_Message = '＜ログアウト＞'
        print(Site_Message)

        Click_Element('ユーザ名', '//span[@class="user-info ng-binding"]', 2)
        Click_Element('ログアウト', '//a[contains(@href, "logout")]/span[@class="hidden-desktop"]', 2)
        time.sleep(5)

    except Exception as e:
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        print(f"【GPM共通処理異常】：" + Site_Message + "\r\n" )
        raise 
