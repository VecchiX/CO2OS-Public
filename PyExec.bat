@echo off

rem 環境変数の設定
set PY_ENV=CO2OS

rem 文字コードをUTF-8に指定（パスに日本語がなければ不要）
chcp 65001

rem 仮想環境をActivateするための特殊なバッチファイルを起動
call %USERPROFILE%\Anaconda3\Scripts\activate.bat
 
rem 仮想環境をActivate
call activate %PY_ENV%
 
rem 作業フォルダを指定
echo %USERPROFILE%
cd %USERPROFILE%\Documents\Python\CO2OS\千葉NT
 
rem python.exeでスクリプトを実行
START /B %USERPROFILE%\anaconda3\envs\%PY_ENV%\python.exe %1.py >>.\log\%1%2.log %2

set PY_ENV= 
rem コマンドプロンプトの画面を残す場合（残さない場合不要）
rem pause