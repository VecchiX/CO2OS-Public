import time
import datetime
import pyautogui as pa

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import initialize_driver, Open_Site, Login
from SEND_EMAIL import send_email   
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

import numpy as np


#### メイン ####
def main():
    now = datetime.datetime.now()
    print('\n\n坂東String Checker Ver.2024.12.13磯部\n')
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
        ret = Get_Info("坂東String")
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


        # 値の編集
        p=driver.find_elements(By.ID, 'statusValue[]')
        n=len(p)

        #欠番対応
        q = [''] * 198

        for i in range(196):
            if i < 9:
                q[i] = p[i].text
            else:
                q[i+2] = p[i].text

        q[9] = '0.0'
        q[10] = '0.0'    

        #配列に入れ直し (PCS, string)
        a=np.zeros((18,11))

        for i in range(n+2):

            if q[i]=="--":
                q[i]="0"

            a[i//11][i%11]=float(q[i])    
                    
        #未接続リスト -1:未接続
        zero_string=np.zeros((18,11))

        for i in range(1):
            zero_string[i][9]=-1
            zero_string[i][10]=-1
                
        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        Zero_String_Message='PCS\tstring\t電流[A]\tコメント\n'

        Zero_FLG = False
        No_Product = True        
        for y in range(18):
            for x in range(11):
                comment = ''
                if (a[y][x]==0) and (zero_string[y][x]==-1):
                    # 未接続のストリング、かつ電流0のとき
                    comment = '未接続'
                elif (a[y][x]==0) and (zero_string[y][x]!=-1):
                    comment = 'No Production'
                    # 一カ所でも未稼働があれば、ゼロフラグをTrueとする
                    Zero_FLG = True
                elif (a[y][x]!=0) and (zero_string[y][x]==-1):
                    # 未接続のストリングに、電流が観測された場合
                    comment = 'ノイズ'
                else:
                    # 一カ所でも稼働していれば、未稼働フラグをFalseとする
                    No_Product = False

                Zero_String_Message += str(y+1) + '\t' + str(x+1) + '\t' + str(a[y][x]) + '\t' + comment + '\n' 

        # メール件名の編集
        if No_Product:
            Conf['件名'] += '（未稼働）'
        elif Zero_FLG:
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

        ## メール配信 ##
        Mail_Body = Site_Message + Zero_String_Message
        send_email(Conf, Mail_Body, [])

        print('完了')    
        time.sleep(30)
        driver.quit()

####    実行
if __name__ == "__main__":
    main()
