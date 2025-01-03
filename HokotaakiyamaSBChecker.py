import time
import datetime
import pyautogui as pa

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import initialize_driver, Login, Open_Site
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


#### メイン ####
def main():
    now = datetime.datetime.now()
    print('\n\n鉾田秋山太陽光発電所 接続箱リスト Ver.2024.12.13磯部\n')
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)

    # 変数設定
    Site_Message=''
    Zero_StringBox_Message=''


    try:
        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # GPM情報取得 
        ret = Get_Info("鉾田秋山SB")
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


        #接続箱電流をチェック
        d=driver.find_elements(By.ID,'statusValue[]')
        time.sleep(10)

        ## ストリング電流取得 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        Zero_StringBox_Message ='SB\t[A]\t\n'
        Zero_FLG = False
        total = 0

        for i in range(24):

            # 接続箱番号(本当は誤りがあるらしい)
            Zero_StringBox_Message += str(i+1)+ '\t'

            # 接続箱電流(数値または「--」のようだ)
            f = d[i].text

            if f.replace('.', '').isnumeric():
                f=float(f)
                Zero_StringBox_Message += "{:>5.1f}".format(f)
            else:
                Zero_StringBox_Message += "{:<5s}".format(f) 
                f=0.0

            total = total + f

            # 接続箱電流がゼロの場合
            if f==0.0:
                Zero_StringBox_Message += '\t' + '電流ゼロ'
                Zero_FLG = True

            # 次の接続箱へ
            Zero_StringBox_Message +='\n'

            if i==11:
                print('\n')
                Zero_StringBox_Message += '\n'

        # メールタイトル編集
        if total==0:
            Conf['件名'] += '（未稼働）'
        elif Zero_FLG:
            Conf['件名'] += '（電流ゼロ）'

        # 処理終了
        print(Zero_StringBox_Message)

        Site_Message = '監視用Webサイトのアクセスに成功しました\n' 
        Zero_StringBox_Message = now.strftime("%Y/%m/%d %H:%M") + "時点のデータです\n\n" + Zero_StringBox_Message


    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message

    finally:
        print(Site_Message)
        driver.close()

        ## メール配信 ##
        Mail_Body=Site_Message + Zero_StringBox_Message
        send_email(Conf, Mail_Body, [])

        print('完了')    
        time.sleep(30)
        driver.quit()            


####    実行
if __name__ == "__main__":
    main()
