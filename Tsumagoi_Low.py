import os
import glob
import time
from datetime import datetime, date, timedelta
import pandas as pd
import pyautogui as pa

from selenium.webdriver.common.by import By

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


def main():

    print('\n\n嬬恋太陽光発電所（低圧） Ver.2024.05.15磯部\n')

    pa.press('esc')
    time.sleep(5)

    Mail_Body = ''
    Site_Message = ''
    Sensor_Message = ''
    sensor1 = 0.0
    sensor2 = 0.0
    sensor3 = 0.0


    try:
        ## Excelからの情報取得 ##
        Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
        print(Site_Message)

        ret = Get_Info('嬬恋低圧')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # データ日時設定
        current_date = Conf['処理日時']
        if current_date=='':
            # 通常は、現在時刻の1時間前
            current_day = date.today().strftime("%Y-%m-%d")
            current_hour = datetime.now() + timedelta(hours=-1)
            current_hour = current_hour.strftime('%H') + ':00'
            current_date = datetime.strptime(f'{current_day} {current_hour}', "%Y-%m-%d %H:%M")
        else:
            # 過去に知事が指定されたとき
            current_date = Conf['処理日時']
            current_date = datetime.strptime(f'{current_date}', "%Y-%m-%d %H:%M")


        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver('new')
        Open_Site(Conf['ログインURL'])


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)
        Login(Conf, 'mailaddress', 'password', '//*[@id="member_form"]/div[2]/a', byBTN=By.XPATH)
        time.sleep(10)


        ## CSVダウンロード ##
        Site_Message = '＜CSVダウンロード＞'
        print(Site_Message)

        #ダイアログ
        Click_Element('ダイアログ','div._btn_alpha.btn_csv', 5, By.CSS_SELECTOR)

        #センサーごと
        Click_Element('センサーごと','._btn.tab_cencer.tab', 5, By.CSS_SELECTOR)

        #時間
        Click_Element('時間','div._btn.btn_csv_select_day', 5, By.CSS_SELECTOR)

        #日付指定
        elms = driver.find_elements(By.XPATH,'//span[contains(@class,"csv_day")]/input[@class="in_date"]')
        for elm in elms:
            elm.clear()
            elm.send_keys(current_date.strftime('%Y/%m/%d'))

        #csv削除
        csvname='C:\\Users\\' + os.getlogin() + '\\Downloads\\eco_megane_*.csv'

        for existing_csvname in glob.glob(csvname):
            if os.path.isfile(existing_csvname) == True:
                os.remove(existing_csvname)
                time.sleep(1)

        #ダウンロード
        Click_Element('時間','._btn.btn_dl', 10, By.CSS_SELECTOR)


        ## CSV読み込み、修正、データ抽出 ##
        Site_Message = '＜CSV読み込み・データ抽出＞'
        print(Site_Message)

        csvname=(
            'C:\\Users\\' + os.getlogin() + '\\Downloads\\eco_megane_'
            + current_date.strftime('%Y%m%d') + '-' 
            + current_date.strftime('%Y%m%d') + '_hour_s_00005823004.csv'
        )

        # CSVファイル内の余剰な「,」を取り除く
        with open(csvname, 'r') as file:
            lines = [line.replace(",,,,", "") for line in file]
        with open(csvname, 'w') as file:
            file.writelines(lines)

        # CSVファイル読み込み
        df = pd.read_csv(csvname, encoding = "shift-jis", skipinitialspace=True)

        # データ抽出
        power_threshold=0.1
        df = df[df["データ計測日"]==current_date.strftime('%Y/%m/%d %H:%M')]

        # センサー01
        row = df[df['センサー番号'] == 'センサー01']
        sensor1=row['売電電力量(kWh)'].values[0]
        Mail_Body = f'センサー01：{sensor1:>8.4f} kWh\n'

        if sensor1 < power_threshold:
            Sensor_Message='センサー01系統\t' + str(sensor1) + ' kWh\tPCS-CとPCS-Dの両方が、たぶん停止しています'+ '\n'
            
        # センサー02
        row = df[df['センサー番号'] == 'センサー02']
        sensor2=row['売電電力量(kWh)'].values[0]
        Mail_Body += f'センサー02：{sensor2:>8.4f} kWh\n'

        if sensor2 < power_threshold:
            Sensor_Message += 'センサー02系統\t' + str(sensor2) + ' kWh\tPCS-AとPCS-Bの両方が、たぶん停止しています'+ '\n'

        # センサー03
        row = df[df['センサー番号'] == 'センサー03']
        sensor3=row['売電電力量(kWh)'].values[0]
        Mail_Body += f'センサー03：{sensor3:>8.4f} kWh\n'

        if sensor3 < power_threshold:
            Sensor_Message=Sensor_Message + 'センサー03系統\t' + str(sensor3) + ' kWh\t\tPCS-Eが、たぶん停止しています'+ '\n'

        if len(Sensor_Message) != 0:
            Sensor_Message=Sensor_Message + '\n'

        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        power_ratio_high=2.25
        power_ratio_low=1.75

        sensor = sensor3
        if sensor==0:
            sensor=0.01   #Oで割らないため

        if (sensor1 / sensor < power_ratio_low) or (power_ratio_high < sensor1 / sensor):
            Sensor_Message += 'センサー01系統と03系統の比率が異常です（2.0前後が正常）\n'

            if sensor3 == 0:
                Sensor_Message += 'センサー01系統/03系統： --.--\n\n'
            else:    
                Sensor_Message += 'センサー01系統/03系統：' + str(format(sensor1/sensor,'.2f')) + '\n\n'
            
            if (power_threshold <= sensor1) and (sensor1 / sensor < 1.0):
                Sensor_Message += 'PCS-CもしくはPCS-Dが、たぶん停止しています' + '\n\n\n'
                    
        if (sensor2 / sensor < power_ratio_low) or (power_ratio_high < sensor2 / sensor):
            Sensor_Message += 'センサー02系統と03系統の比率が異常です（2.0前後が正常）\n'
            
            if sensor3 == 0:
                Sensor_Message += 'センサー02系統/03系統： --.--\n\n'
            else:    
                Sensor_Message += 'センサー02系統/03系統：' + str(format(sensor2/sensor,'.2f')) + '\n\n'
            
            if (power_threshold <= sensor2) and (sensor2 / sensor < 1.0):
                Sensor_Message += 'PCS-AもしくはPCS-Bが、たぶん停止しています' + '\n\n\n'


        Site_Message = (
            '監視用Webサイトのアクセスに成功しました。\n'
            + current_date.strftime('%Y/%m/%d %H:%M') +' 時点のデータです。\n\n' 
        )

    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message += "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message

    finally:
        driver.close()
        
        ## メール配信 ##
        if Sensor_Message != '':              #電力量異常
            Mail_Body += '\n' + Sensor_Message
            Conf['件名'] += '（電力量異常）'

        Mail_Body = Site_Message + Mail_Body
        print(Mail_Body)

        send_email(Conf, Mail_Body, [])

        print('完了')
        time.sleep(30)
        driver.quit()


#####   実行
if __name__ == '__main__':
    main()