## 本ツールについて
##
## ・次のコマンドでEXE化することができる。
##  pyinstaller --onefile HIKARI_NTT_Checker.py
## ・EXEファイルを配布する場合は、「CONFIG.xlsx」も一緒に配布する。
## ・ExEファイルと「CONFIG.xlsx」を同じフォルダに格納し、
## 　EXEファイルを実行する

import time
from datetime import datetime
import pyautogui as pa
import os

from selenium.webdriver.common.by import By

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

#####   転送状態確認
def TransferStatus(driver):
    Status = ""

    # 無条件転送であるか確認
    radio = driver.find_element(By.ID, "r1")
    if not radio.is_selected():

        # 転送停止であるか確認
        radio = driver.find_element(By.ID, "stop")
        if radio.is_selected():
            Status = "転送停止"
        else:
            Status = "選択状態異常"
        
    return Status

#####   メイン
def main():

    print('\n\nひかり電話 設定確認 Ver.2024.12.13磯部\n')

    pa.press('esc')
    time.sleep(5)

    Mail_Body = ''
    Site_Message = ''
    current_date = datetime.now()
    Status = ''

    try:
        ## Excelからの情報取得 ##
        Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
        print(Site_Message)

        ret = Get_Info('ひかり電話確認')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 電話番号履歴ファイルからのデータ取得
        documents_folder = os.path.expanduser("~\\Documents")
        file_path = os.path.join(documents_folder, Conf['ファイル名'])
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as file:
                lines = file.readlines()
        else:
            lines = []
        
        lines = [line.strip() for line in lines]

        # 電話番号と名前の辞書を作成する
        phone_numbers = Conf['電話番号'].split()
        names = Conf['名前'].split()
        phone_book = dict(zip(phone_numbers, names))

        # 名前関数を定義
        def get_name_by_num(phone_number):
            name = phone_book.get(phone_number, "該当する名前が見つかりません")
            return name

        # driver初期化
        driver = initialize_driver(Conf['画面表示'])


        ## サイトに接続 ##
        Site_Message = '＜ひかり電話 設定サイトに接続＞'
        print(Site_Message)
        Open_Site(Conf['URL'])


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)
        Click_Element('ログイン', '//input[@type="image"]', 5)
        Login(Conf, 'phoneNo', 'pass')

        # パスワード有効期限確認    
        elms = driver.find_elements(By.CLASS_NAME, 'mes_error')
        if elms:
            # ボタンをクリック
            Click_Element('パスワード有効期限', '//input[@type="image"]', 5)

        # ボイスワープ
        Click_Element('ボイスワープ', '//input[@title="ボイスワープ"]', 5)


        ## サービス開始／停止確認 ##
        Site_Message = '＜サービス開始／停止確認＞'
        print(Site_Message)
        Click_Element('ボイスワープ', '//input[@title="サービス開始／停止"]', 5)
        Status = TransferStatus(driver=driver)

        ## 転送番号取得 ##
        Site_Message = '＜転送番号取得＞'
        print(Site_Message)
        name = ''

        if Status =='':
            Click_Element('サービスメニューへ', 'img_b0_al_top', 5, By.CLASS_NAME)
            Click_Element('転送先電話番号設定', '//input[@title="転送先電話番号設定"]', 5)

            elms = driver.find_elements(By.ID, 'listValue')
            numbers = driver.find_elements(By.XPATH, '//input[@type="text"]')

            for index, elm in enumerate(elms):
                if elm.is_selected():
                    elm = numbers[index]
                    phone_number = elm.get_attribute('value')
                    name = get_name_by_num(phone_number)
                    Status = phone_number + ":" + name


        ## 名前のチェック ##
        Site_Message = '＜名前のチェック＞'
        print(Site_Message)
        if name == '':
            name = Status

        lines.append(name)
        if "見つかりません" in name:
            # 電話番号リストに該当しないとき
            Conf['件名'] += '【警告】該当なし'
        elif "異常" in name:
            Conf['件名'] += '【警告】選択状態異常'
        elif len(lines) >=3 and all(line == name for line in lines[-3:]):
            # 3回以上繰り返しているとき
            Conf['件名'] += name + 'さん 【警告】3回以上繰り返し'
        else:
            Conf['件名'] += name + 'さん'

        lines = lines[-4:]

        # 電話番号履歴ファイルの更新処理
        with open(file_path, "w", encoding="utf-8") as file:
            file.writelines(line + "\n" for line in lines)


        ## ログアウト ##
        Site_Message = '＜ログアウト＞'
        print(Site_Message)
        Click_Element('ログアウト', '//input[@title="ログアウト"]', 5)


        ## メッセージを編集 ##
        Site_Message = '＜メッセージを編集＞'
        print(Site_Message)

        Mail_Body = Status + '\n\n'

        Site_Message = (
            'ひかり電話の設定確認に成功しました。\n'
            + current_date.strftime('%Y/%m/%d %H:%M') +' 時点の情報です。\n\n' 
        )

    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message + "\r\n"

    finally:        
        ## メール配信 ##
        Mail_Body = Site_Message + Mail_Body
        print(Mail_Body)

        send_email(Conf, Mail_Body, [])

        driver.close()
        print('完了')
        time.sleep(5)
        driver.quit()


#####   実行
if __name__ == '__main__':
    main()