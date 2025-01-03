import time
import pandas as pd

from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from FUNCTIONS import Click_Element, read_html_NEW
from SEND_EMAIL import send_email
from GPM_COMMON_PROCESS import Site_Initialize, logout, SetDate
from CONFIGxlsx import dictionary as Conf

#### メイン ####
def main():
    
    Site_Message = ''
    String_Message = ''
    DIRECT_CURRENT = '直流電流ストリング'
    zero_couter = 0
    zero_Str_info_bak=''
    zero_Str_info=''

    try:

        print("\n\n黒川Stringsチェッカー Ver.2024.12.13.磯部\n")

        ## サイト初期化 ##
        driver = Site_Initialize("黒川Stringsチェック")


        ## データ指定 ##
        Site_Message = '＜データ指定・検索＞'
        print(Site_Message)

        date_from = Conf['処理日']
        SetDate(date_from)

        Click_Element('項目', '//select[@ng-model="ctrl.typology"]', 2)
        Click_Element('ストリングス', '//option[contains(@label, "STRINGBOX")]', 5)
        Click_Element('要素全選択', '(//span[contains(text(),"要素")])[2]', 2)

        # パラメータに「直流電流ストリング」を含むものを検索する
        elm = driver.find_element(By.XPATH, '(//input[@ng-model="ctrl.search.name"])[4]')
        elm.send_keys(DIRECT_CURRENT)
        time.sleep(2)

        # 直流電流ストリング 01、02、…　のループ処理
        elms = driver.find_elements(By.XPATH, '//span[contains(text(),"' + DIRECT_CURRENT + '") and @class]')

        # 空のデータフレームを作成
        result_df = pd.DataFrame()


        for elm in elms:
            elm.click()
            Click_Element('OK', '(//button[contains(text(),"OK")])[2]', 2)

            # テーブル表示待ち
            table = WebDriverWait(driver, 90).until(
                EC.visibility_of_element_located((By.XPATH, '//table'))
            )   
            if not table.text.strip():
                Site_Message =Site_Message+"(タイムアウト発生)"
                raise Exception()

            ## 件数指定 10番目(750)を選択 ##
            if result_df.empty:
                elm2=driver.find_element(By.ID,'perPageSelect')
                elm2.click()
                select = Select(elm2)
                select.select_by_index(9)
                time.sleep(2)

            elm.click()

            ## データ処理 ##
            Site_Message = '＜データ処理＞'
            print(Site_Message)

            # テーブルを取得
            table = driver.find_element(By.XPATH, '//table')

            # Pandasでテーブルを読み込む
            table_html = table.get_attribute('outerHTML')
            df = read_html_NEW(table_html)[0]
            # df = pd.read_html(table.get_attribute('outerHTML'))[0]

            # 最初の時刻と同じデータだけを抽出する
            first_time = df.loc[0, '日付']
            df = df[df['日付'] == first_time].sort_values(by='デバイス')

            # 列名に「直流電流ストリング」が含まれる列を検索
            column_with_current = df.columns[df.columns.str.contains(DIRECT_CURRENT)]

            if not column_with_current.empty:
                # 列名の変更
                df.columns = [col.replace(column_with_current[0], DIRECT_CURRENT) for col in df.columns]

                # デバイス列の値の変更
                column_number = column_with_current[0].replace(DIRECT_CURRENT,'').replace("（A) A", "")

                df['デバイス'] = df['デバイス'] + f'.{column_number}'

            # 結合
            result_df = pd.concat([result_df, df], ignore_index=True)
            result_df = result_df.sort_values(by='デバイス')

            if column_number == "15":
                break


        # 出力値リストの編集
        String_Message = ""
        Zero_Message = ""

        for index, row in result_df.iterrows():

            Message = "{:<15} {:>9.4f} A  ".format(row["デバイス"], row[DIRECT_CURRENT])

            if row['デバイス'] in Conf['未接続Strings']:
                Message += "未接続"
            elif row[DIRECT_CURRENT] <=0.01:
                Message += "電流ほぼゼロ"
                zero_Str_info = row["デバイス"][0:8]
                if zero_Str_info == zero_Str_info_bak:
                    Zero_Message += ',' + row["デバイス"][9:12]
                else:
                    Zero_Message += '\n' + row["デバイス"]
                zero_couter +=1 
                zero_Str_info_bak = zero_Str_info

            String_Message += Message +'\n'


        ## ログアウト ##
        logout()
        Site_Message = '監視用Webサイトのアクセスに成功しました\n' + first_time +' 時点のデータです。'

        ## メッセージ編集 ##
        if zero_couter > 0:
            Zero_Message = "電流ほぼゼロのストリング"+ Zero_Message

        String_Message = Zero_Message + "\n\n以下、全ストリング電流\n" + String_Message
        print(String_Message)

        if Zero_Message == "":
            Conf['件名'] += "（正常）"
        else:
            Conf['件名'] += "（電流ゼロ）"

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
        send_email(Conf, Site_Message + "\n\n" + String_Message, [])

        print('完了')    
        time.sleep(10)
        driver.quit()


####    実行
if __name__ == "__main__":
    main()
