import time
import datetime
import pyautogui as pa

import glob
import os

from FUNCTIONS import initialize_driver, Open_Site, Click_Element

from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


#### メイン ####
def main():
    now = datetime.datetime.now()
    print("\n\n日本ルツボ豊田CSV Ver.2024.12.13磯部\n")
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)

    Site_Message =""
    download_files =  os.path.expanduser("~\\Downloads") +  '\\GT_ソーラーパーク*.CSV'

    try:

        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # 基本情報取得    
        ret = Get_Info('日本ルツボ豊田CSV')
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 開始日
        # CONFIGファイルから取得 ／ 前日一日分
        date_from = datetime.datetime.now()
        if Conf['開始日'] =='':
            date_from = date_from + datetime.timedelta(days=-1)
            date_from = date_from.strftime('%Y-%m-%d')
        else:
            date_from = Conf['開始日']

        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['URL'])


        ## ダウンロードフォルダ内の「PCS_～」ファイルを削除j
        Site_Message = '＜過去分のファイルを削除＞'
        print(Site_Message)

        for filename in glob.glob(download_files):
            os.remove(filename)

        ## 前日分に移動 ##
        Site_Message = "＜前日分に移動＞"
        print(Site_Message)
        Click_Element('前日分', '(//button[@class="GT-DateBar__DatepickerButton"])[3]', 2)

        ## CSV出力 ##
        Site_Message = '＜CSV出力＞'
        print(Site_Message)

        Click_Element('CSV出力', '//button[text()="CSV出力"]', 5)

        # 最後に開いたウィンドウ・タブに切り替える
        original_window = driver.current_window_handle
        driver.switch_to.window(driver.window_handles[len(driver.window_handles) - 1])
        Click_Element("グループ", '//button[text()="グループ"]', 30)
        Click_Element("キャンセル", '//button[text()="キャンセル"]', 3)

        # もとのウィンドウに切り替える
        driver.switch_to.window(original_window)
    
        Site_Message = '監視用Webサイトのアクセスに成功しました'

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

        ##  メール配信 ##
        path = []

        # メールに添付するCSVファイルを指定する
        if download_files:
            path = glob.glob(download_files)

        send_email(Conf, Site_Message, path)

        print('完了')    
        time.sleep(10)
        driver.quit()

####    メール配信
if __name__ == "__main__":
    main()
