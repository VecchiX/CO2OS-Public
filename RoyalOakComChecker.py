import time
from datetime import datetime, timedelta
import glob
import os.path
import pyautogui as pa
import pandas as pd

from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By

from FUNCTIONS import initialize_driver, Open_Site, Login
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


def main():

    print('\n\nロイヤルオーク Communication Checker Ver.2024.12.13磯部\n\n')

    pa.press('esc')
    time.sleep(5)

    Site_Message=''
    Comm_Message=''
    total = 0
    No_Communication = False
    No_Product = False

    try:

        ## Excelからの情報取得 ##
        Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
        print(Site_Message)

        ret = Get_Info('ロイヤルオーク')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # データ日時設定
        if Conf['処理日時']=='':
            dt_now = datetime.now() + timedelta(minutes=-1)
        else:
            dt_string = Conf['処理日時']
            dt_now = datetime.strptime(dt_string, "%Y-%m-%d %H:%M")


        ## 過去分ファイル削除 ##
        Site_Message = '＜過去分のファイルを削除＞'
        print(Site_Message)

        #ファイル削除
        file_list = glob.glob('c:\\Users\\' + os.getlogin() + '\\Downloads\\Rawdata_*.csv')

        for file in file_list:
            if os.path.isfile(file):
                os.remove(file)
                time.sleep(5) #5


        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['ログインURL'])


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)
        Login(Conf, "login_user_id", "login_user_pass", "login_submitter")
        time.sleep(10)


        ## ファイル出力 ##
        Site_Message = '＜ファイル出力＞'
        print(Site_Message)

        #最新計測値
        driver.find_element(By.XPATH,"//*[@id='menu_area']/div/a[2]").click()
        time.sleep(10)

        # 日付
        dt_string = dt_now.strftime("%Y/%m/%d")
        date_input = driver.find_element(By.XPATH, '//input[@name="dojo.date1"]')
        date_input.clear()
        date_input.send_keys(dt_string)

        # 時間
        hour = str(dt_now.hour)
        h1=driver.find_elements(By.CLASS_NAME,"hour ")
        select = Select(h1[0])
        select.select_by_value(hour)
        time.sleep(1)

        select = Select(h1[1])
        select.select_by_value(hour)
        time.sleep(1)

        # 分
        minute = str(dt_now.minute)
        n1=driver.find_elements(By.CLASS_NAME,"minute ")
        select = Select(n1[0])
        select.select_by_value(minute)
        time.sleep(1)

        select = Select(n1[1])
        select.select_by_value(minute)
        time.sleep(1)

        #download
        driver.find_element(By.XPATH,"//*[@id='CSV_data']/button").click()
        time.sleep(10)


        ## ファイル読み込み ##
        Site_Message = '＜ファイル読み込み＞'
        print(Site_Message)

        file_list = glob.glob('c:\\Users\\' + os.getlogin() + '\\Downloads\\Rawdata_*.csv')

        for file in file_list:
            if os.path.isfile(file):
                csvfile=file

        print(csvfile)
        print('\n')

        df = pd.read_csv(csvfile, encoding="utf-8")

        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        # 有効電力のデータを抽出する
        df = df[df["Parameter"]=="有効電力"]

        # メッセージ編集
        PCS_Names = Conf['PCSリスト'].split(',')

        for PCS_Name in PCS_Names:
            row_data = df[df["Device Name"]==PCS_Name]
            if row_data.empty:
                row_string = PCS_Name + ", 無通信"
                No_Communication = True
            else:
                row_data = row_data[["Device Name", "Measured Value", "Unit"]]
                value = row_data["Measured Value"].values[0]
                row_string = (
                    row_data["Device Name"].values[0] + 
                    "  " + f"{value:>10,.1f}" + 
                    row_data["Unit"].values[0]
                )

                if value==0:
                    No_Product = True
                
                total += value

            Comm_Message += row_string + '\n'

        comments = []
        if total==0:
            comments.append('未稼働') 
            Comm_Message = "発電所全体が稼働していないようです。\n\n" + Comm_Message
        elif No_Product:
            comments.append('出力ゼロ')
            Comm_Message = "稼働していないPCSが存在しています。\n\n" + Comm_Message

        if No_Communication:
            comments.append('無通信')
            Comm_Message = "無通信状態のPCSが存在しています。\n" + Comm_Message

        if comments:
            Conf['件名'] = "発電監視警報：" + Conf['件名'] + "（" +'・'.join(comments) + "）"


        Site_Message = (
            '監視用Webサイトのアクセスに成功しました。\n'
            + dt_now.strftime('%Y/%m/%d %H:%M') +' 時点のデータです。\n\n' 
        )
        print(Comm_Message)

    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message += "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message

    finally:
        print(Site_Message)
        driver.close()

        ## メール配信 ##       
        Mail_Body= Site_Message + Comm_Message
        send_email(Conf, Mail_Body, [])

        print('完了')
        time.sleep(30)
        driver.quit()


################################ メール配信 ################################
if __name__ == '__main__':
    main()