import time
import datetime
import pyautogui as pa
import pandas as pd
from math import isnan

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import Click_Element, initialize_driver, Open_Site, Login, read_html_NEW

from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


#### メイン ####
def main():
    now = datetime.datetime.now()
    print("\n\n東かがわPCSチェック Ver.2024.12.13.磯部\n")
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)


    Power_Message = ''
    PCS_Message = ''
    Site_Message = ''
    current_hour = now.strftime('%H:00').lstrip('0')

    try:
        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # GPM情報取得 
        ret = Get_Info("東かがわPCSチェック")

        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 時刻確認 6:00以前ならば前日データを取得する ##
        Site_Message = '時刻確認'
        print(Site_Message)
        
        previous_day = False
        date = now.date()
        if now.time() < datetime.datetime.strptime('06:00', '%H:%M').time():
            previous_day = True
            date = now + datetime.timedelta(days=-1)


        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)

        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['ログインURL'])

        #ポップアップを消す
        elm = driver.find_elements(By.ID,"onetrust-accept-btn-handler")
        if 0 < len(elm): 
            elm[0].click()
        else:
            print('None')
        time.sleep(10)


        ## ログイン ##
        Site_Message = '＜ログイン＞'
        Click_Element("ログイン", '//a[contains(@id,"Login")]', 10)
        print(Site_Message)
        Login(Conf, "username", "password")


        ## 発電所選択 ##
        Site_Message = '＜発電所選択＞'
        print(Site_Message)

        driver.get(Conf['システム一覧URL'])
        time.sleep(10)
        Click_Element('東かがわ', '//a[contains(text(),"' + Conf['件名'] + '")]', 10)
        time.sleep(10)


        ## 時間遅れ電力 ##
        Site_Message = '＜時間遅れ電力の取得＞'
        print(Site_Message)

        power=driver.find_element(By.CLASS_NAME,"mainValueAmount").text
        unit =driver.find_element(By.CLASS_NAME,"mainValueUnit").text
        delay=driver.find_element(By.CLASS_NAME,"widgetSubHead").text
        time.sleep(10)

        if isinstance(power, (int, float)):
            if float(power) < 1.0:
                Power_Message=power + ' ' + unit + '\t' + delay + '\tたぶん、発電していなかった（遅延データだから）\n'
            else:
                Power_Message=power + ' ' + unit + '\t' + delay + '\n'
        else:
            Power_Message = "時間遅れ電力データは取得できませんでした\n"

        print(Power_Message + '\n')


        ## 詳細画面オープン ##
        Site_Message = '＜詳細画面オープン＞'
        print(Site_Message)

        Click_Element('解析', '//a[@title="解析"]', 20)
        Click_Element('すべてのデバイスの選択', '//label[contains(@for,"SelectAll")]', 20)
        Click_Element('太陽光発電システム全体', '//label[contains(@for,"PlantSelection")]', 15)
        Click_Element('詳細', '//span[contains(@id,"ChartDetail")]', 15)

        if previous_day:
            Click_Element('前日',"PC_pagebtn", 15, By.CLASS_NAME)

        ## テーブル読み込み ##
        Site_Message = '＜テーブル読み込み＞'
        print(Site_Message)

        # テーブル表示待ち
        table = WebDriverWait(driver, 90).until(
            EC.visibility_of_element_located((By.XPATH, '//table[contains(@id,"ChartDetail")]'))
        )   
        if not table.text.strip():
            Site_Message += "(タイムアウト発生)"

        table_html = table.get_attribute('outerHTML')
        table = read_html_NEW(table_html, header=None)[0]
        # table = pd.read_html(table.get_attribute('outerHTML'), header=None)[0]
        df = pd.DataFrame(table.values[1:], columns=table.iloc[0])
        df.columns.values[0] = "時刻"
        suffix_to_remove = ' 電力 平均値 [kW]'
        df.columns = [col.replace(suffix_to_remove, '') for col in df.columns]

        ## データ編集 ##
        Site_Message = '＜データ編集＞'
        print(Site_Message)

        # 現在時、または前日12時のデータを取得
        last_record = df[df['時刻']==current_hour]
        if previous_day:
            last_record =  df[df['時刻']=='12:00']

        # 発電量書式設定
        No_Product = False
        total = 0
        for column, value in zip(last_record.columns[1:], last_record.values[0, 1:]):
            value = float(value)
            total = total + value
            Message = "{:<10} {:>8.2f} kW  ".format(column, value)

            Comment = ''
            if value < 1.0 or isnan(value) :
                Comment = "\tNo Production"
                No_Product = True
                
            PCS_Message += Message + Comment +'\n'

        # メールタイトル、メッセージ編集
        if total == 0:
            Conf['件名'] += "（未稼働）"
            PCS_Message = "全PCSが稼働していないようです。\n\n" + PCS_Message 
        elif No_Product:
            Conf['件名'] = "【重要】" + Conf['件名'] + "【PCS停止アラート】" 
            PCS_Message = "PCSが停止しているようです。至急確認してください。\n\n" + PCS_Message 


        ## ログアウト
        Click_Element('ユーザ名', "ctl00_Header_lblUserName", 5, By.ID)
        Click_Element('ログアウト', "ctl00_Header_hylLogout", 15, By.ID)

        Site_Message = '監視用Webサイトのアクセスに成功しました\n' 
        PCS_Message = "以下、" + date.strftime("%Y/%m/%d") +" " + last_record['時刻'].values[0]+ "時点のPCS毎の発電量です。\n\n"+PCS_Message

        print(PCS_Message)
        print(Site_Message)

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

        ## メール送信 ##
        Mail_Body=Site_Message + '\n' + Power_Message + '\n' + PCS_Message
        send_email(Conf, Mail_Body, [])

        print('完了')    
        time.sleep(10)
        driver.quit()


####    実行
if __name__ == "__main__":
    main()
