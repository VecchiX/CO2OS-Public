import time
import os
import glob
import shutil
import re
import pandas as pd
from io import StringIO

from selenium import webdriver
from selenium.webdriver.common.by import By
# seleniumのバージョンアップにより、
# 自動的に最適なChromeDriverがインストールされるようになったため
# コメントアウトしました
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import ElementClickInterceptedException

print("\n\n 共通関数 Ver.2024.12.13.磯部\n")

driver = None   # ドライバー初期化


## ChromeDriver更新 ##
def initialize_driver(headless_mode=None):
    global driver

    if driver is None:

        Site_Message = '＜ChromeDriverの取得・設定＞'
        print(Site_Message)

        options = webdriver.ChromeOptions()
        # options = webdriver.EdgeOptions()
        options.use_chromium = True     # Chronmiumベースとする

        #ヘッドレスモード設定
        if headless_mode == 'OFF':
            # 画面操作は正確であるが、画面ハードコピーの際、画像の欠けが発生する
            options.add_argument('--headless=new')
            # options.add_argument('--headless')  # 一時避難的な変更
        elif headless_mode == 'OFF(old)':
            # 画面操作は不安定であるが、画面ハードコピーは正確に実施される
            # options.add_argument('--headless')
            options.add_argument('--headless=old')  # 一時避難的な変更
        
        # その他optionの設定
        # 画面のサイズを設定
        options.add_argument('--window-size=1920,1080')
        options.add_experimental_option("excludeSwitches", ["enable-logging", 'enable-automation', 'load-extension'])

        prefs = {
            'profile.default_content_setting_values.notifications' : 2,                 # 通知ポップアップを無効
            'credentials_enable_service' : False,                                       # パスワード保存のポップアップを無効
            'profile.password_manager_enabled' : False,                                 # パスワード保存のポップアップを無効
            'download.prompt_for_download': False,                                      # ダウンロードのpromptを表示しない
        }
        options.add_experimental_option('prefs', prefs)
        options.add_experimental_option("detach", True) # 終了時も開いたまま
        options.add_argument("--start-maximized")        # 最大化
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        options.add_argument("--disable-blink-features=AutomationControlled")

        # seleniumのバージョンアップにより、自動的に最適なChromeDriverがインストールされるようになった
        # driver = webdriver.Chrome(service =Service(ChromeDriverManager().install()), options =options)
        driver = webdriver.Chrome(options =options)
        driver.set_page_load_timeout(60)

        return driver


## 監視用Webサイトに接続 ##
def Open_Site(url):
    global driver

    driver.get(url)

    #最大化・最前面
    driver.minimize_window()
    time.sleep(3)
    driver.maximize_window()
    time.sleep(3)


## 要素をクリック ##
def Click_Element(message, xpath_str, second, by=By.XPATH):
    global driver
    message = '「' + message + '」クリック'
    print(message)
    element = driver.find_element(by, xpath_str)
    time.sleep(1)

    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)
    finally:
        time.sleep(second - 1)
        return element


## ログイン ##
def Login(Conf, obj_ID, objPASS, objBTN="", byID=By.ID, byPASS=By.ID, byBTN=By.ID):
    global driver
    elm = driver.find_element(byID, obj_ID)
    elm.clear()
    elm.send_keys(Conf['ログインID'])

    elm = driver.find_element(byPASS, objPASS)
    elm.clear()
    elm.send_keys(Conf['パスワード'])
    time.sleep(5)

    if objBTN == "":
        elm.submit()
        time.sleep(10)

    else:
        Click_Element('ログイン', objBTN, 10, byBTN)


## 画面を下方にスクロール ##
def scroll():
    while 1:
        html01=driver.page_source
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        html02=driver.page_source
        if html01!=html02:
            html01=html02
        else:
            break


## StringIOでHTMLコンテンツをラップし、pd.read_html を呼び出す共通関数 ##
def read_html_NEW(html_content, **kwargs):
    # Parameters:
    #     html_content (str): HTMLの文字列
    #     **kwargs: pd.read_htmlに渡す追加のキーワード引数

    # Returns:
    #     list of DataFrames: 読み込まれたHTMLテーブルのDataFrameリスト

    # return pd.read_html(html_content, **kwargs)
    return pd.read_html(StringIO(html_content), **kwargs)


## フォルダ内のファイルを移動 ##
def f_move(path_from, path_to):

    list_file_name =  os.listdir(path_from)
    for i_file_name in list_file_name:
        join_path = os.path.join(path_from, i_file_name)
        move_path = os.path.join(path_to, i_file_name)

        if os.path.isfile(join_path):
            shutil.move(join_path, move_path)


## フォルダ内の最新のファイル(フォルダ)を検索 ##
class FileInfo:
    def __init__(self, filename, fullpath):
        self.filename = filename
        self.fullpath = fullpath

def f_newfile(directory):
    list_of_files = glob.glob(os.path.join(directory, '*'))
    latest_file = max(list_of_files, key=os.path.getctime)
    filename = os.path.basename(latest_file)
    file_info = FileInfo(filename=filename, fullpath=latest_file)
    return file_info


## 数値文字列より、単位を取り除いて数値部分のみを抽出する ##
def extract_numeric(value):
    match = re.search(r'[\d.]+', value)
    return float(match.group()) if match else None

