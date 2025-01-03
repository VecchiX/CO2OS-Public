import time
from datetime import datetime, timedelta
import pyautogui as pa

from selenium.webdriver.common.by import By

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element 
import glob
import os
from io import BytesIO

from PIL import Image
from SEND_EMAIL import send_email as SendEmail
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

def main():
    now = datetime.now()
    print('\n\nSGET 千葉ニュータウン Ver.2024.12.13磯部\n')
    print(now.strftime('%Y/%m/%d %H:%M'))

    pa.press('esc')
    time.sleep(5)

    Site_Message = ''
    PCS_Message = ''
    Alart_Message = ''
    fullpath = []

    try:

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # 千葉NT情報
        ret = Get_Info('千葉NT')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception


        ## 監視用Webサイトに接続 ##
        Site_Message = '監視用Webサイトに接続'
        print(Site_Message)

        # 「ON」または「OFF(old)」のとき正常に画像が処理される
        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['URL'])


        ## 監視用Webサイトにログイン ##
        Site_Message = '監視用Webサイトにログイン'
        print(Site_Message)
        Login(Conf, "USER_ID", "PASSWORD","", By.NAME, By.NAME)
        time.sleep(10)

        # ログイン確認（ログインできていれば、これはエラーにならない）
        elm = driver.find_element(By.ID, 'userNameLabel')

        ## スクリーンショットを保存 ##
        Site_Message = 'スクリーンショットを保存'
        print(Site_Message)

        #保存・削除ファイル名を設定
        file_name = 'ChibaNT*.jpg'
        folder_name='c:\\Users\\' + os.getlogin() + '\\Desktop\\'
        fullpath = folder_name + file_name

        #古い画像ファイルを削除する
        for png_name in glob.glob(fullpath):
            if os.path.isfile(png_name)==True:
                os.remove(png_name)
                time.sleep(1)

        #保存ファイル名に日付を設定する
        dt_now = datetime.now()
        if dt_now.hour < 12:
            AM_PM = 'AM'
        else:
            AM_PM ='PM'

        fullpath = fullpath.replace('*', str(dt_now.date()) + AM_PM)

        #スクショ保存
        w = driver.execute_script('return document.body.scrollWidth')
        h = driver.execute_script('return document.body.scrollHeight')
        driver.set_window_size(w, h)
        screenshot = driver.get_screenshot_as_png()
        image = Image.open(BytesIO(screenshot)).convert('RGB')
        image.save(fullpath, 'JPEG', quality=25)


        ## アラート確認 ##
        Site_Message = 'アラート確認'
        print(Site_Message)

        # アラート情報の取得
        elm = driver.find_element(By.ID, 'headerAlarmDateLabel')
        alart_date = datetime.strptime(elm.text, '%Y/%m/%d %H:%M:%S')
        alart_text = driver.find_element(By.ID, 'messageTextFieldLabel').text
        alart_status = driver.find_element(By.ID, 'headerRecordSatatusLabel').text
        alart_grade = driver.find_element(By.ID, 'headerGradeLabel').text 

        # 現在時刻とアラート日時の比較
        alart_flg = False
        dt_yesterday = dt_now - timedelta(days=1)
        if dt_now.hour <= 20:
            # 現在時刻が21:00以前の場合、前日21:00以降と本日分のアラートが注意対象
            if alart_date.date() == dt_yesterday.date() and alart_date.hour >= 21:
                alart_flg = True
            elif alart_date.date() == dt_now.date():
                alart_flg = True
        else:
            # 現在時刻が21:00の以降、本日9:00以降のアラートが注意対象
            if alart_date.date() == dt_now.date() and alart_date.hour >= 9:
                alart_flg = True

        if alart_flg:
            Alart_Message = (
                "\n\n"+ alart_date.strftime('%Y/%m/%d %H:%M:%S')
                + " にアラート状態が変化しました。\n"
                + f"アラートレベル：{alart_grade} ({alart_status})\n"
                + f"アラート内容　：{alart_text}\n\n"
                + "手動で監視画面を開き、「アラーム」画面より「警報」を確認してください。\n"
            )

            Conf['件名'] = '【重要】' + Conf['件名'] + '【アラート状態に変化あり】'


        ## PCS出力確認(0.0kWを検索) ##
        Site_Message = 'PCS出力確認(午前中のみ)'
        print(Site_Message)

        if AM_PM == 'AM':
            elms = driver.find_elements(By.XPATH, '//*[text()="0.0"]')
            if elms:
                PCS_Message = (
                    "\n\n停止しているPCSが存在します。\n"
                    + "チェックシートへ「×」を記入して、エスカレーションする必要があります。\n")

                if '重要' not in Conf['件名']:
                    Conf['件名'] = '【重要】' + Conf['件名']
                Conf['件名'] += '【PCS停止】'


        ## メニュー画面を開いてログアウト ##
        Site_Message = 'メニュー画面を開いてログアウト'
        print(Site_Message)

        Click_Element('メニュー','image4165-2', 10, By.ID)
        Click_Element('ログアウト', '//a[text()="ログアウト"]', 10)
        Site_Message = '監視用Webサイトのアクセスに成功しました'

    except:
        ## ツール異常終了時 ##
        if '監視用Webサイト' in Site_Message:
            Site_Message += (
                'できませんでした。\n\n'
                + '至急、手動でサイトにログインできるか確認してください。')
            Conf['件名'] = '【重要】' + Conf['件名'] + '【ログイン失敗】'

        elif Site_Message:
            Site_Message = (
                '【ツール異常終了】\r\n'
                + Site_Message + 'できませんでした。\n\n')

        else:
            Site_Message = "【ツール異常終了】\r\n想定外の問題が発生しました。"

    finally:
        print(Site_Message)
        driver.close()

    ## メール配信 ##
    path = []

    if fullpath :
        path = [fullpath]

    Site_Message += PCS_Message + Alart_Message

    SendEmail(Conf, Site_Message, path)

    print('完了')
    time.sleep(30)
    driver.quit()

####  実行  ####
if __name__ == '__main__':
    main()
