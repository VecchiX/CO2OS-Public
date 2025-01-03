import time
from datetime import datetime
import pyautogui as pa

from selenium.webdriver.common.by import By

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

import unicodedata

#####   全角を半角に変換
def zen_to_han(text):
    return unicodedata.normalize('NFKC', text)

#####   メイン
def main():

    print('\n\nHNGK(ecoめがね)アラート検出 Ver.2024.12.13磯部\n')

    pa.press('esc')
    time.sleep(5)

    Mail_Body = ''
    Site_Message = ''
    current_date = datetime.now()
    ID = ''

    try:
        ## Excelからの情報取得 ##
        Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
        print(Site_Message)

        ret = Get_Info('HNGK_ecoめがね')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # ID、パスワードの編集
        IDs = Conf['ログインIDs'].split(',')
        PWDs = Conf['パスワードs'].split(',')

        # driver初期化
        driver = initialize_driver(Conf['画面表示'])


        ## ID、パスワード毎のループ処理 ##
        for ID, PWD in zip(IDs, PWDs):

            Conf['ログインID'] = ID
            Conf['パスワード'] = PWD

            ## 監視用Webサイトに接続 ##
            Site_Message = '＜監視用Webサイトに接続＞'
            print(Site_Message)
            Open_Site(Conf['ログインURL'])


            ## ログイン ##
            Site_Message = '＜ログイン＞'
            print(Site_Message)
            Login(Conf, 'mailaddress', 'password', '//*[@id="member_form"]/div[2]/a', byBTN=By.XPATH)
            time.sleep(10)


            ## サブメニュー表示 ##
            Site_Message = '＜サブメニュー表示＞'
            print(Site_Message)
            Click_Element('サブメニュー', 'username', 10, By.ID)
            Click_Element('商品一覧', '//a[text()="商品一覧"]', 10)


            ## 異常発電所の番号を取得 ##
            Site_Message = '＜異常発電所の番号を取得＞'
            print(Site_Message)

            # 「×」の発電所を取得する
            elms = driver.find_elements(By.XPATH, '//*[@src="/productview/img/icon_cross.png"]/parent::td/preceding-sibling::td[@class="name change-display"]/span')

            # 発電所の番号を取得する
            for elm in elms:    
                plant_no = elm.get_attribute("title")
                if plant_no =='':
                    plant_no = elm.text

                Mail_Body += plant_no +"\n"

            ## ログアウト
            Site_Message = '＜ログアウト＞'
            print(Site_Message)

            Click_Element('サブメニュー', 'username', 10, By.ID)
            Click_Element('ログアウト', '//a[@class="logout"]', 10)

            # アラートを取得し、OKする
            alert = driver.switch_to.alert
            alert.accept()


        ## メッセージを編集 ##
        Site_Message = '＜メッセージを編集＞'
        print(Site_Message)

        if Mail_Body != "":
            Mail_Body = zen_to_han(Mail_Body)
            Mail_Body = '異常が確認された発電所の番号は、以下のとおりです。\n\n' + Mail_Body
        else:
            Mail_Body = '発電所に異常はありませんでした。'

        Site_Message = (
            '監視用Webサイトのアクセスに成功しました。\n'
            + current_date.strftime('%Y/%m/%d %H:%M') +' 時点のデータです。\n\n' 
        )

    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "「" + ID + "」にて【ツール異常終了】\r\n" + Site_Message + "\r\n"

    finally:        
        ## メール配信 ##
        Mail_Body = Site_Message + Mail_Body
        print(Mail_Body)

        send_email(Conf, Mail_Body, [])

        driver.close()
        print('完了')
        time.sleep(30)
        driver.quit()


#####   実行
if __name__ == '__main__':
    main()