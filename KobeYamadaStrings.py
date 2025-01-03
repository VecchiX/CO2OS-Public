import time
import datetime
import pyautogui as pa
import random
import string

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import glob
import os
import zipfile
import pandas as pd

from FUNCTIONS import initialize_driver, Open_Site, Click_Element, Login
from FUNCTIONS import f_move, f_newfile

from SEND_EMAIL import send_email
from CONFIGxlsx import Get_Info
from CONFIGxlsx import dictionary as Conf


#### メイン ####
def main():
    now = datetime.datetime.now()
    print("\n\n神戸山田Strings Ver.2024.12.13磯部\n")
    print(now.strftime("%Y/%m/%d %H:%M"))

    pa.press("esc")
    time.sleep(5)

    Site_Message =""
    download_directory = os.path.expanduser("~\\Downloads")
    output_file_path = ""

    try:
        ## Excelからの情報取得 ##
        Site_Message = 'CONFIGファイル(EXCEL)に接続'
        print(Site_Message)

        # 神戸山田Stirng情報 
        ret = Get_Info("神戸山田Strings")
        if ret == False:
            raise Exception

        # Yahooアカウント
        ret = Get_Info("メール送信設定")
        if ret == False:
            raise Exception

        # 開始日、終了日
        # CONFIGファイルから取得 ／ 前日一日分
        date_from = datetime.datetime.now()
        if Conf['開始日'] =='':
            date_from += datetime.timedelta(days=-1)
            date_from = date_from.strftime('%Y-%m-%d')
            date_to = date_from
        else:
            date_from = Conf['開始日']
            date_to = Conf['終了日']

        ## 監視用Webサイトに接続 ##
        Site_Message = '＜監視用Webサイトに接続＞'
        print(Site_Message)
        driver = initialize_driver(Conf['画面表示'])
        Open_Site(Conf['URL'])


        ## 監視用Webサイトにログイン ##
        Site_Message = '＜ログイン＞'
        print(Site_Message)
        Login(Conf, "username", "value", "submitDataverify")
        time.sleep(20)

        ## ダウンロードフォルダ内の「PCS_～」ファイルを削除j
        Site_Message = '＜過去分のファイルを削除＞'
        print(Site_Message)

        for filename in glob.glob(download_directory + '\\PCS???_*_*.xlsx'):
            os.remove(filename)

        # # 発電所名のファイルは削除しないことになったのでコメントアウトしました。
        # for filename in glob.glob(download_directory + '\\' + Conf['英数名'] + '????-??-??.xlsx'):
        #     os.remove(filename)

        # ## 言語切替
        # Site_Message = "＜言語切替＞"
        # print(Site_Message)
        # Click_Element('言語設定', '//span[@id="refr.mm.locale"]', 2)
        # Click_Element('日本語', '//span[@title="日本語"]', 10)

        ## 「設備管理」をクリック ##
        Site_Message = "＜画面切替(設備管理)＞"
        print(Site_Message)
        Click_Element('設備管理', '//span[@title ="設備管理"]', 10)
        Click_Element('クッキーポリシー', '//*[@id="cookie-policy"]/i',2)

        ## データ指定 ##
        Site_Message = '＜データ指定・検索＞'
        print(Site_Message)
        Click_Element('設備タイプ', '(//div[@class ="ant-form-item-control-input-content"])[1]', 10)
        Click_Element('PCS', '(//div[@class ="ant-select-item-option-content"][text() ="PCS"])[1]', 10)
        Click_Element('検索', '//button[text()="検索"]', 10)

        ## 件数指定 ##
        Site_Message = '＜処理件数指定＞'
        print(Site_Message)

        Click_Element('ページサイズ', '//div[@aria-label="ページサイズ"]', 2)
        Click_Element('300件', '//div[@title ="300 件 / ページ"]',10)

        #### ループ処理 ####
        Counter = 0
        while True:
            Counter += 1
            Counter_txt = str(Counter).zfill(3)

            ## エクスポート ##
            Site_Message = '＜エクスポート＞'
            print(Site_Message)

            # 設備状態を全件チェック
            time.sleep(20)
            elm = driver.find_element(By.XPATH, '(//input[@type ="checkbox"])[1]')
            if not elm.is_selected():
                elm.click()
                time.sleep(5)
            # Click_Element('全件チェック', '(//input[@type ="checkbox"])[1]', 5)
            
            Click_Element('性能データのエクスポート', '//button[text()="性能データのエクスポート"]', 30)

            ## PVxx電流選択 ##
            Site_Message = '＜PVxx電流選択＞'
            print(Site_Message)

            # 選択解除
            elm = driver.find_element(By.XPATH, '//span[text()="すべて選択"]/parent::label//input')
            if  elm.is_selected():
                elm.click()
                time.sleep(5)
            
            # PV1～10の電流
            elms = driver.find_elements(By.XPATH, '//span[contains(text(),"入力電流")]/preceding-sibling::span')
            i = 1
            for elm in elms:
                elm.click()
                time.sleep(3)
                i += 1
                if i > 10:
                    break

            # OK
            Click_Element('OK', '//span[text()="OK"]/parent::button', 5)

            ## ダウンロード ##
            Site_Message = '＜ダウンロード＞'
            print(Site_Message)

            # ダウンロードタスクの名前の設定
            DL_NAME = "Py" + date_from + "-" + Counter_txt
            DL_NAME += ''.join(random.choices(string.ascii_uppercase, k=3))
            # ※末尾にランダムな3文字を付与して、名前の重複を防ぐ

            elm = driver.find_element(By.XPATH, '//input[@id="taskName"]')
            elm.send_keys(Keys.CONTROL + "a")
            elm.send_keys(Keys.DELETE)
            time.sleep(1)
            elm.send_keys(DL_NAME)

            # 期間指定
            Click_Element('期間(カスタム)','//input[@value="bySelfDefine"]', 3)
            elm = driver.find_element(By.XPATH, '//input[@placeholder="開始日付"]')
            elm.click()
            elm.send_keys(date_from)
            elm.send_keys(Keys.ENTER)

            elm = driver.find_element(By.XPATH, '//input[@placeholder="終了日付"]')
            elm.click()
            elm.send_keys(date_to)
            elm.send_keys(Keys.ENTER)

            Click_Element('1h', '//input[@value="3600"]', 1)
            Click_Element('結合', '//input[@value="BY_DEVTYPE"]', 1)
            Click_Element('OK', '//span[text()="OK"]/parent::button', 50)
            Click_Element(
                'ダウンロード'
                , '//div[@title="' + DL_NAME + '"]/parent::div/parent::div//a[@title="ダウンロード"]'
                , 10)
            Click_Element('確定', '//button[@class="ant-btn ant-btn-primary"]', 5)
            Click_Element(
                '削除'
                , '//div[@title="' + DL_NAME + '"]/parent::div/parent::div//a[@title="削除"]'
                , 5)
            Click_Element('確定', '//button[@class="ant-btn ant-btn-primary"]', 5)

            ## ダウンロードしたファイルの操作 ##
            Site_Message = '＜ダウンロードしたファイルの操作＞'
            print(Site_Message)

            # ダウンロードディレクトリ内の最新のファイルを取得
            latest_file = f_newfile(download_directory)

            # ZIPファイルを展開
            zipf = zipfile.ZipFile(latest_file.fullpath)
            zipf.extractall(download_directory)
            time.sleep(2)
            zipf.close()
            os.remove(latest_file.fullpath)

            # 展開したファイルをDownLoadフォルダの先頭に移動
            latest_file = f_newfile(download_directory)
            f_move(latest_file.fullpath, download_directory)
            os.removedirs(latest_file.fullpath)
            latest_file = f_newfile(download_directory)
            os.rename(
                latest_file.fullpath
                , latest_file.fullpath.replace('PCS_', 'PCS'+ Counter_txt + '_'))

            ## 次のページへ ##
            Site_Message = '＜次のページへ＞'
            print(Site_Message)

            Click_Element('ダウンロード画面クローズ', '//button[@aria-label="Close"]', 2)

            # 「次のページへ」がクリックできる限り、ループする
            elm = driver.find_element(By.XPATH, '//span[@aria-label="right"]/parent::button')
            if elm.is_enabled():
                Click_Element('次のページへ', '//span[@aria-label="right"]/parent::button', 2)
            else:
                print('終了')
                break


        ## ファイルを結合 ##
        excel_files = [f for f in os.listdir(download_directory) if f.endswith('.xlsx') and f.startswith('PCS')]

        # すべてのExcelファイルを読み込んで結合
        dfs = []
        for file in excel_files:
            file_path = os.path.join(download_directory, file)
            df = pd.read_excel(file_path, skiprows=3)  # 3行目までスキップ
            dfs.append(df)

        # 結合したデータフレームを新しいExcelファイルとして保存
        output_file_path = os.path.join(download_directory, Conf['英数名'] + date_from +'.xlsx')
        merged_df = pd.concat(dfs, ignore_index=True)
        merged_df.to_excel(output_file_path, index=False, engine='openpyxl')

        # 結合したデータを読み込む
        merged_df = pd.read_excel(output_file_path)

        # 15列目以降を削除
        merged_df = merged_df.iloc[:, :15]

        # 結合したデータフレームを上書き保存
        merged_df.to_excel(output_file_path, index=False, engine='openpyxl')

        ## ログアウト ##
        Site_Message = '＜ログアウト＞'
        print(Site_Message)

        Click_Element('ユーザ名', '//*[@id="refr.mm.user"]/span', 2)
        Click_Element('ログアウト', '//*[@id="action-0-0"]/span[2]', 2)
        Click_Element('はい', '//*[text()="はい"]/parent::button', 5)
        time.sleep(5)

        Site_Message =  '神戸山田Webサイトのアクセスに成功しました。\r\n\r\n'\
                        '添付されたExcelファイルは500kB前後の容量になります。\r\n'\
                        '不要となったメール(もしくは添付ファイル)を削除することをお勧めします。'

    except:
        if Site_Message:
            Site_Message = Site_Message + "できませんでした。"
        else:
            Site_Message = "想定外の問題が発生しました。"
        
        Site_Message = "【ツール異常終了】\r\n" + Site_Message

    finally:
        print(Site_Message)
        driver.close()

        ## メール配信 ##
        if output_file_path == "":
            path = {}
        else:
            path = glob.glob(output_file_path)
        
        send_email(Conf, Site_Message, path)

        print('完了')    
        time.sleep(10)
        driver.quit()

####    実行
if __name__ == "__main__":
    main()
