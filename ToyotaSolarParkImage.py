import time
import datetime
import pyautogui as pa

from selenium.webdriver.common.by import By

import glob
import os
from io import BytesIO
import base64

from FUNCTIONS import initialize_driver, Open_Site, Click_Element, extract_numeric, read_html_NEW
from PIL import Image

from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf

#### スクショ保存 ####
def save_screenshot(driver, file_path, is_full_size=False):
    # スクリーンショット設定
    screenshot_config = {
        # Trueの場合スクロールで隠れている箇所も含める、Falseの場合表示されている箇所のみ
        "captureBeyondViewport": is_full_size,
    }

    # スクリーンショットを取得し、jpgで保存
    base64_image = driver.execute_cdp_cmd("Page.captureScreenshot", screenshot_config)
    image_data = base64.b64decode(base64_image["data"])
    image_stream = BytesIO(image_data)
    image = Image.open(image_stream).convert('RGB')
    image.save(file_path, 'JPEG', quality=25)

    time.sleep(2)


#### メイン ####
def main():
    now = datetime.datetime.now()
    print("\n\n日本ルツボ 豊田ソーラーパーク 画像添付 Ver.2024.12.13磯部\n")
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)

    Site_Message = ""
    Power_Message = ""
    fullpath = []
    path=[]

    try:

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # 基本情報取得    
        ret = Get_Info('日本ルツボ豊田画像')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 現在時刻
        date_from = datetime.datetime.now()

        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)

        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['URL'])


        ## 過去分のファイルを削除 ##
        Site_Message = '＜過去分のファイルを削除＞'
        print(Site_Message)
        file_name = '日本ルツボ豊田_*.jpg'
        folder_name='c:\\Users\\' + os.getlogin() + '\\Desktop\\'
        fullpath = folder_name + file_name

        #古い画像ファイルを削除する
        for png_name in glob.glob(fullpath):
            if os.path.isfile(png_name)==True:
                os.remove(png_name)
                time.sleep(1)


        ## 発電量取得 ##
        Site_Message = '＜発電量取得＞'
        print(Site_Message)

        # テーブルを取得
        table = driver.find_element(By.CLASS_NAME, 'GT-EvaluationTable')

        # テーブルのヘッダーを取得
        header_elements = table.find_elements(By.XPATH, './/div[@class="GT-EvaluationTable__HeaderRowItem"]//a')
        column_names = [header.text for header in header_elements]

        # テーブルの行ヘッダーを取得
        row_header_elements = table.find_elements(By.XPATH, './/div[@class="GT-EvaluationTable__HeaderColItemInner"]')
        row_names = [header.text.strip() for header in row_header_elements]

        # Pandasでテーブルを読み込む
        table_html = table.get_attribute('outerHTML')
        df = read_html_NEW(table_html)[0]

        # カラム名、行名を設定
        df.columns = column_names
        df.index = row_names

        print(df)

        # 一番右の列を取得
        values = df.iloc[:, -1]
        column_name = df.columns[-1]

        # 発電量が0の場合、前日のデータとする
        if values['ソーラーパーク発電量'] == '0.00kWh':
            values = df.iloc[:, -2]
            column_name = df.columns[-2]
        
        values = values.apply(extract_numeric)

        # BANK毎の発電量を取得
        BANK1 = values['BANK1 発電量']
        BANK2 = values['BANK2 発電量']

        # 発電量メッセージの編集
        Power_Message = (
            f"\n\n{column_name}の発電情報は次の通りです。\n\n"
            + f"全天日射強度〔W/㎡〕：{values['全天日射強度(W/㎡)']}\n"
            + f"発電量合計　 〔kWh〕：{values['ソーラーパーク発電量']}\n"
            + f"BANK1 発電量〔kWh〕：{BANK1}\n"
            + f"BANK2 発電量〔kWh〕：{BANK2}\n" )

        # BANK1とBANK2の差が20%以上の場合、アラートメッセージを追加
        if BANK1 + BANK2 and abs(BANK1 - BANK2) / ((BANK1 + BANK2) / 2) >= 0.1:
            Power_Message += (
                "\n【注意】\nBANK1とBANK2の発電量の差が大きい(10%以上)です。\n"
                + "30分の画像より、一方のBANKが停止していないか確認してください。\n"
                + "※出力制御の可能性もあります。出力制御情報は社外でも参照可能です。\n")
            Conf['件名'] = "【重要】" + Conf['件名'] +"（発電量比率異常）"

        # BANK1とBANK2が共に停止している場合、アラートメッセージを追加
        if BANK1 <=0 and BANK2 <=0:
            Power_Message += "\nBANK1、BANK2共に未稼働です\n"
            Conf['件名'] += "（未稼働）"

        print(Power_Message)

        ## スクリーンショットを保存 ##
        Site_Message = '＜スクリーンショットを保存＞'
        print(Site_Message)

        # 日毎の画像を保存
        driver.set_window_size(1920, 1080)
        time.sleep(2)
        save_screenshot(driver, fullpath.replace('*', '日'), is_full_size=True)

        # 30分毎の画像を保存
        Click_Element('30分', '(//input[@type="radio" and @value="0"])[1]', 3)
        driver.set_window_size(3440, 1080)
        Click_Element('左端へ', '(//button[@class="GT-DateBar__DatepickerButton"])[1]', 3)
        Click_Element('右端へ', '(//button[@class="GT-DateBar__DatepickerButton"])[6]', 3)
        save_screenshot(driver, fullpath.replace('*', '30分'), is_full_size=True)

        # 処理完了
        Site_Message = (
            '監視用Webサイトのアクセスに成功しました\n' 
            + date_from.strftime("%Y年%m月%d日 %H:%M") 
            +' 時点の画像データです。')

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
    if fullpath:
        path = glob.glob(fullpath)

    send_email(Conf, Site_Message + Power_Message, path)

    print('完了')    
    time.sleep(10)
    driver.quit()

####    実行
if __name__ == "__main__":
    main()
