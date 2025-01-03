import time
from datetime import datetime
import pyautogui as pa
import pandas as pd

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import initialize_driver, Open_Site, Login, read_html_NEW
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

def main():
    print('\n\n福島富岡Checker Ver.2024.12.13磯部\n\n')

    pa.press('esc')
    time.sleep(5)

    Site_Message=''
    Energy_Message=''

    try:

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        ret = Get_Info('福島富岡')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        #処理日
        if Conf['処理日']=='':
            dt = datetime.today()
        else:
            dt = datetime.strptime(Conf['処理日'], "%Y-%m-%d")        

        startYear=str(dt.year)
        startMonth=str(dt.month)
        startDay=str(dt.day)

        #メール(CC)の退避(電力異常時のみCCにも送る)
        mail_cc = Conf['CC']
        Conf['CC'] = ''


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)

        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['ログインURL'])

        print(Site_Message)
        Login(Conf, "username", "password", "", By.NAME, By.NAME)

        # ログイン確認（ログインできていれば、これはエラーにならない）
        elm = driver.find_element(By.CLASS_NAME, 'mypageTitle')


        ## 監視画面接続 ##
        Site_Message = '＜監視画面接続＞'
        print(Site_Message)

        base_url=Conf['ベースURL']
        login_url = f'{base_url}startYear={startYear}&startMonth={startMonth}&startDay={startDay}'
        Open_Site(login_url)

        # テーブル表示待ち
        table = WebDriverWait(driver, 90).until(
            EC.visibility_of_element_located((By.XPATH, '//div[@id="ledgerSheetArea"]/table'))
        )   
        if not table.text.strip():
            Site_Message = Site_Message+"(タイムアウト発生)"


        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        # テーブル読み込み
        table_html = table.get_attribute('outerHTML')
        df = read_html_NEW(table_html, header=None)[0]
        # df = pd.read_html(table.get_attribute('outerHTML'), header=None)[0]

        # 1列目を数値に変換
        df[0] = df[0].str.extract('(\d+)').astype(float)

        # 8時から17時までのデータを抽出
        df_selected = df[(df[0] >= 8) & (df[0] <= 15)]

        # 列1を削除
        df_selected = df_selected.drop(columns=[1])

        # 1列目の値を08～09時、09時～10時、10時～11時のように書式設定して揃える
        df_selected[0] = df_selected[0].apply(lambda x: f"{int(x):02d}時～{int(x)+1:02d}時：")

        # データの状態（発電量ゼロ、無通信の有無)
        Energy_Message = dt.strftime('%Y年%m月%d日') + 'のデータです。\n'
        
        zero_data_exists = any(df_selected[2] == 0.0)
        if zero_data_exists:
            Energy_Message += '・発電量ゼロの時間帯があります。\n'

        empty_data_exists = any(df_selected[2].isnull())
        if empty_data_exists:
            Energy_Message += '・無通信の時間帯があります。\n'

        # 売電量（書式を設定する）
        df_selected[2] = df_selected[2].apply(lambda x: '{:,.1f} kW/h'.format(x) if pd.notnull(x) else "-- kW/h")
        Energy_Message += '\n' + df_selected.to_string(index=False, header=False)

        # 電力異常時のみCCにもメール配信
        if zero_data_exists or empty_data_exists:
            Conf['件名'] += '（電力量異常）'
            Conf['CC'] = mail_cc

        Site_Message = '監視用Webサイトのアクセスに成功しました。\n' 
        print(Energy_Message)

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
        Mail_Body = Site_Message + '\n' + Energy_Message + '\n\n' + 'CO2OS Monitoring Team'
        send_email(Conf, Mail_Body, [])

        print('完了')
        time.sleep(30)
        driver.quit()

#####   実行
if __name__ == '__main__':
    main()