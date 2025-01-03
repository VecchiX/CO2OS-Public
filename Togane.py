import time
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains

import pyautogui as pa

from FUNCTIONS import initialize_driver, Open_Site, Login, Click_Element
from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


def main():
    Version = '東金市上布田太陽光発電所 Ver.2024.12.13磯部'
    print('\n' + Version + '\n')

    pa.press('esc')
    time.sleep(5)

    try:

        Site_Message=''
        Power_Message=''

        ## Excelからの情報取得 ##
        Site_Message = '＜CONFIGファイル(EXCEL)に接続＞'
        print(Site_Message)

        ret = Get_Info('東金')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception


        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)

        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['ログインURL'])

        ## ログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)

        #ログアウト(ログインしたままの場合に対応する)
        elms = driver.find_elements(By.ID, "exitUser")
        if elms:
            elms[0].click()
            time.sleep(5)

        #ログイン（ID, PASS入力）
        Login(Conf, "userCode", "userPwd", "btnLogin", byPASS=By.NAME)

        #スライドバーをdrag
        verify = driver.find_element(By.ID,"verify")
        a = verify.size
        dw = int(a["width"])

        print("スライドバー幅：" + str(dw))

        v=driver.find_element(By.ID,"nc_1_n1z")
        actions = ActionChains(driver)
        actions.click_and_hold(v).move_by_offset(dw,0).perform()
        actions.release().perform()
        time.sleep(10)


        ## 数値取得 ##
        Site_Message = '＜数値取得＞'
        print(Site_Message)

        # 現在の電力
        power=driver.find_element(By.XPATH,'//*[@filedname="data.detail.pac"]')
        Power_Message = '現在の電力 : ' + power.text + 'kW\n\n'
        time.sleep(5)

        # PCS電力量を表示
        Click_Element('PCS電力量', '//*[@id="inverterModal"]', 5)

        # デバイス名を取得・編集
        device_names=driver.find_elements(By.CLASS_NAME,'deviceNameSpan')
        PCS_Name = []
        for device_name in device_names:
            device_name = device_name.text
            print(device_name)
            if device_name != '':
                if device_name[-1] == '1':
                    PCS_Name.append('PCS-3') 
                elif device_name[-1] == '2':
                    PCS_Name.append('PCS-2') 
                elif device_name[-1] == '3':
                    PCS_Name.append('PCS-1') 

        # PCS発電量
        PCS_Value = []
        device_values = driver.find_elements(By.CLASS_NAME, "progress-bar")
        for device_value in device_values:
            device_value = device_value.text
            if device_value != "":
                PCS_Value.append(device_value)

        # PCS_Nameの順にソート
        sorted_data = sorted(zip(PCS_Name, PCS_Value), key=lambda x: x[0])

        # ソートされたデータから各リストを再構築
        PCS_Name, PCS_Value = zip(*sorted_data)

        ## メッセージ編集 ##
        Site_Message = '＜メッセージ編集＞'
        print(Site_Message)

        # PCS毎の発電量
        for i in range(len(PCS_Name)):
            Power_Message += PCS_Name[i] + ' : ' + PCS_Value[i] + '\n'

        # 最大値、最小値
        numeric_value=[float(value.replace('kWh', '')) for value in PCS_Value]
        pmax=max(numeric_value)
        pmin=min(numeric_value)

        print("max : " + str(pmax) + "kWh    min : " + str(pmin)+ "kWh\n")

        if pmin < pmax * 0.7:
            Power_Message += '\nPCS間の電力量差が大きいです!'
            Conf['件名'] += '（異常）'


        ## ログアウト ##
        Site_Message = '＜ログアウト＞'
        print(Site_Message)

        Click_Element('クローズ', '//*[@class="close"]', 5)
        Click_Element('ユーザーサークル', '//*[@class="fa fa-user-circle"]', 5)
        Click_Element('ログアウト', '//*[@class="lgoout"]', 5)

        Site_Message = (
            '監視用Webサイトのアクセスに成功しました。\n'
            + datetime.today().strftime('%Y/%m/%d %H:%M') +' 時点のデータです。\n\n' 
        )

    except:
        ## ツール異常終了時 ##
        if Site_Message:
            Site_Message += "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message

    finally:
        print('\n')
        print(Site_Message)
        print(Power_Message)

        ## メール配信 ##
        Mail_Body = Version + '\n\n' + Site_Message + '\n' + Power_Message
        send_email(Conf, Mail_Body, [])

        print('完了')
        time.sleep(30)
        driver.quit()


####    実行
if __name__ == '__main__':
    main()