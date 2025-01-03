import time
import datetime
import pyautogui as pa

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

import os
import sys
import glob
import pandas as pd

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


def main():
    print('\n君津ソーラーパーク Ver.2024.05.15磯部\n')

    pa.press('esc')
    time.sleep(5)

    Site_Message=''
    AllSection_Message=''
    plant_max = 0
    plant_min = 0

    try:

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        args = sys.argv
        ret = Get_Info(args[1])
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 区画データの存在確認用
        kukaku_text = Conf['区画']
        kukaku_list = kukaku_text.split(',')

        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver('new')
        Open_Site(Conf['ログインURL'])


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)
        Login(Conf, "login-id", "login-password")


        ## ダウンロードページ ##
        Site_Message = '＜ダウンロードページ＞'
        print(Site_Message)
        driver.get(Conf['ダウンロードURL'])
        time.sleep(10)

        ## ダウンロード ##
        Site_Message = '＜ダウンロード＞'
        print(Site_Message)

        #期間
        dt = datetime.datetime.today()

        # 日付取得
        d = dt.day 

        # ファイル名設定
        filename_date = 'C:\\Users\\' + os.getlogin() + '\\Downloads\\min_' + dt.strftime("%Y-%m-%d")

        Section_Message=''
        Exists_Zero = False

        dropdown1 = driver.find_element(By.NAME,'target')
        select1 = Select(dropdown1)
        elements = driver.find_elements(By.XPATH, '//*[@name="target"]//option')
        for element in elements:

            Section_Message = ""
            value = element.get_attribute("value")
            text = element.text
            select1.select_by_value(value)
            time.sleep(1)

            # 日付指定
            dropdown2=driver.find_element(By.NAME,'fromDD')
            select2 = Select(dropdown2)
            select2.select_by_value(str(d))
            time.sleep(1)

            dropdown3=driver.find_element(By.NAME,'toDD')
            select3 = Select(dropdown3)
            select3.select_by_value(str(d))
            time.sleep(1)

            # csv削除
            csvname = filename_date + '*.csv'

            for csv_name in glob.glob(csvname):
                if os.path.isfile(csv_name):
                    os.remove(csv_name)
                    time.sleep(2)

            # ダウンロード
            Click_Element('ダウンロード', 'csvdownload',10, By.NAME)

            oldcsvname=filename_date + '.csv'

            if os.path.isfile(oldcsvname):
                kukaku = text.replace('君津ソーラーパーク','').replace('発電所', '')
                kukaku_no = kukaku.replace('区画', '')
                kukaku_list.remove(kukaku_no)

                print(kukaku + "処理中")
            
                newcsvname = filename_date + '_' + kukaku +'.csv'

                os.rename(oldcsvname,newcsvname)
                time.sleep(1) #5
                
                # CSVを読み込む
                df = pd.read_csv(newcsvname, encoding = "shift-jis")

                # データ絞り込み
                filterd_columns = [col for col in df.columns if '最大瞬時電力' in col]
                result_df = df[filterd_columns].iloc[-1]

                #最大瞬時電力 総量、系統
                for column_name, value in result_df.items():

                    if not pd.isna(value):
                        column_name = column_name.replace('最大瞬時電力(kW)', '')
                        if column_name == '総量':
                            column_name = '全体'
                            if plant_max < value:
                                plant_max = value

                            if plant_min ==0:
                                plant_min = value
                            elif value < plant_min:
                                plant_min = value

                        column_name = kukaku + '_' + column_name
                        Section_Message += f"{column_name:<12}{value:>6.2f} kW"

                        if value < 0.05:
                            Section_Message += '\tたぶん発電していません'
                            Exists_Zero = True

                        Section_Message +='\n'

                if Section_Message!='':
                    Section_Message +='\n'

                AllSection_Message += Section_Message
            else:
                AllSection_Message = 'csvファイルが存在しません\n'


        # 電力ゼロ
        comments = []
        if Exists_Zero:
            comments.append('電力　ゼロ') 

        # 存在すべき区画データが存在しない場合
        if kukaku_list:
            kukaku_Message = '区画' + ', '.join(kukaku_list) + 'のデータが存在しません。\n\n'
            AllSection_Message = kukaku_Message + AllSection_Message
            comments.append('区画データ無')

        # 区画の電力差が大きい場合
        if plant_max != 0:
            if plant_min/plant_max < 0.7:
                AllSection_Message = '区画間の電力差が大きいです\n\n' + AllSection_Message
                comments.append('電力差　大') 

        if comments:
            Conf['件名'] += "（" +'・'.join(comments) + "）"

        Site_Message = '監視用Webサイトのアクセスに成功しました。\n' 
        Site_Message += datetime.datetime.now().strftime("%Y/%m/%d %H:%M") + "時点のデータです。\n\n"
        print(AllSection_Message)

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
        Mail_Body = Site_Message + '\n' + AllSection_Message
        send_email(Conf, Mail_Body, [])
        print('完了')
        time.sleep(30)
        driver.quit()    


#####   実行
if __name__ == '__main__':
    main()
