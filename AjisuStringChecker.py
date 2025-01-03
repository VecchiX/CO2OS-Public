import time
import datetime
import pyautogui as pa
# import pandas as pd
# from io import StringIO

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import initialize_driver, Open_Site, Login, read_html_NEW
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


#### メイン ####
def main():

    now = datetime.datetime.now()
    print('\n\n阿知須String Checker Ver.2024.12.13磯部\n')
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)

    # 変数設定
    Site_Message=''
    Zero_String_Message=''

    try:
        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # GPM情報取得 
        ret = Get_Info("阿知須String")
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)

        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['ログインURL'])

        print(Site_Message)
        Login(Conf, "username", "password", "", By.NAME, By.NAME)

        # ログイン確認（ログインできていれば、これはエラーにならない）
        elm = driver.find_element(By.CLASS_NAME, 'mypageTitle')


        ## ストリング電流取得 ##
        Site_Message = '＜ストリング電流取得＞'
        print(Site_Message)
        Open_Site(Conf['データ画面URL'])

        # テーブル表示待ち
        tables = WebDriverWait(driver, 90).until(
            EC.presence_of_all_elements_located((By.ID, 'stringSheet[]'))
        )   
        if not any(table.text.strip() for table in tables):
            Site_Message = "タイムアウト発生"


        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        Zero_String_Message='\n'
        No_Production = False

        for i, table in enumerate(tables):

            table_html = table.get_attribute('outerHTML')
            df = read_html_NEW(table_html, header=None)[0]
            
            # テーブル番号の書式設定
            table_number = f"{i + 1:02}"

            # 未接続String番号の取得
            No_connect_list = Conf['未接続String'+table_number]

            # 列名1～9の値を取得
            for col in range(1, 10):
                col_name = f"{col}"
                col_value = df[col_name].values[0]

                comment = ''
                if col_name in No_connect_list:
                    if col_value == 0:
                        comment = '未接続'
                    else:
                        comment = 'ノイズ'
                elif col_value == 0:
                    comment = 'No Production'
                    No_Production = True
                
                if col_value == '--':
                    comment = 'No Production'
                    No_Production = True
                    col_value = "-----.-"
                else:
                    col_value = f"{col_value: 7.1f}"
                
                Zero_String_Message += f"JB {table_number}-{col_name.ljust(2)} {col_value} A\t" + comment + '\n'

        if No_Production:
            Conf['件名'] += '（電流ゼロ）'

        print(Zero_String_Message)

        Site_Message = '監視用Webサイトのアクセスに成功しました\n' 
        Zero_String_Message = now.strftime("%Y/%m/%d %H:%M") + "時点のデータです\n\n" + Zero_String_Message

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

        ####    メール配信
        Mail_Body = Site_Message + Zero_String_Message
        send_email(Conf, Mail_Body, [])

        print('完了')    
        time.sleep(30)
        driver.quit()            


#####   実行
if __name__ == "__main__":
    main()
