import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import Click_Element, read_html_NEW

from SEND_EMAIL import send_email
from GPM_COMMON_PROCESS import Site_Initialize
from GPM_COMMON_PROCESS import logout
from GPM_COMMON_PROCESS import SetDate

from CONFIGxlsx import dictionary as Conf

#### メイン ####
def main():

    Site_Message = ''
    PCS_Power_Message = ''
    ACTIVE_POWER ='有効電力'

    try:
        print("\n\nGPM_PCSチェッカー Ver.2024.12.13.磯部\n")

        ## サイト初期化 ##
        driver = Site_Initialize('')

        ## 日付指定 ##
        SetDate(Conf['処理日'])

        ## データ指定 ##
        Site_Message = '＜データ指定＞'
        print(Site_Message)
        Click_Element('項目', '//select[@ng-model="ctrl.typology"]', 2)
        Click_Element('インバータ', '//option[contains(@label, "INVERTER")]', 5)
        Click_Element('要素', '(//span[contains(text(),"要素")])[2]', 2)
        elm = driver.find_element(By.XPATH, '(//div[@class="col-xs-6"]/input)[2]')
        elm.send_keys(ACTIVE_POWER)
        time.sleep(3)
        Click_Element(ACTIVE_POWER, '//label/span[contains(text(), " ' + ACTIVE_POWER + ' ")]', 2)
        Click_Element('OK', '(//button[@ng-disabled="!ctrl.btnOk"])[2]', 5)

        # テーブル表示待ち
        table = WebDriverWait(driver, 90).until(
            EC.visibility_of_element_located((By.XPATH, '//table'))
        )   
        if not table.text.strip():
            Site_Message =Site_Message+"(タイムアウト発生)"
            raise Exception()

        ## 件数指定 4番目(100)を選択 ##
        elm=driver.find_element(By.ID,'perPageSelect')
        elm.click()
        select = Select(elm)
        select.select_by_index(3)
        time.sleep(5)

        ## データ処理 ##
        Site_Message = '＜データ処理＞'
        print(Site_Message)

        # テーブルを取得
        table = driver.find_element(By.XPATH, '//table')

        # Pandasでテーブルを読み込む
        # 木更津はカンマと小数点を修正する必要がある
        table_html = table.get_attribute('outerHTML')
        if Conf['英数名'] =="Kisarazu":
            df = read_html_NEW(table_html, decimal=',', thousands='.')[0]
        else:
            df = read_html_NEW(table_html)[0]

        # 最初の時刻と同じデータだけを抽出する
        first_time = df.loc[0, '日付']
        first_records = df[df['日付'] == first_time].sort_values(by='デバイス')

        # 最大値と最小値を取得
        ACTIVE_POWER = ACTIVE_POWER + " kW"
        power_max = first_records[ACTIVE_POWER].max()
        power_min = first_records[ACTIVE_POWER].min()

        #メッセージ作成
        message_title = Conf['件名']
        if power_min <= 0 and power_max <= 0:
            PCS_Power_Message = '全PCSが稼働していないようです。\n\n'
            message_title += "(未稼働)"
        elif power_min <= 0:
            PCS_Power_Message = 'PCSが停止しているようです。至急確認してください。\n\n'
            message_title = "【重要】" + message_title + "【PCS停止アラート】"
        elif power_min / power_max <= Conf['しきい値'] / 100:
            PCS_Power_Message = 'ぱっとしないPCSがあります\n\n'
            message_title += "(有効電力異常)"
        else:
            message_title += "(正常)"

        Conf['件名'] = message_title

        # 出力値リストの編集
        for index, row in first_records.iterrows():

            PCS_Power_Message += "{:<15} {:>10.3f} KW  ".format(row['デバイス'], row[ACTIVE_POWER])

            label = ""
            ratio_text = "{:>7}".format("--.--")        
            if power_max > 0:
                ratio = row[ACTIVE_POWER] / power_max * 100
                ratio_text = "{:>7.2f}".format(ratio)

                if ratio < Conf['しきい値']:
                    label = " " + Conf['しきい値コメント']

            PCS_Power_Message += ratio_text + " % "
            
            if power_max > 0:
                if row[ACTIVE_POWER] == power_max:
                    label += " (max)"
                elif row[ACTIVE_POWER] == power_min:
                    label += " (min)"

            PCS_Power_Message += label +'\n'

        print(PCS_Power_Message)

        ## ログアウト ##
        logout()

        Site_Message = '監視用Webサイトのアクセスに成功しました\n' + first_time +' 時点のデータです。'

    except Exception as e:
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        print(f"【ツール異常終了】\r\n{e}\r\n")
        Site_Message = "【ツール異常終了】" + Site_Message

    finally:
        print(Site_Message)
        driver.close()

        ## メール配信 ##
        send_email(Conf, Site_Message + "\n\n" + PCS_Power_Message, [])

        print('完了')
        time.sleep(10)
        driver.quit()

####    実行
if __name__ == "__main__":
    main()