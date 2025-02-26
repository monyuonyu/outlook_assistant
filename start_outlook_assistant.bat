@echo off

title Outlook秘書アシスタント
echo Outlook秘書アシスタントを起動しています...

:: Pythonの仮想環境をアクティベート
call .venv\Scripts\activate.bat

:: API KEYを環境変数として設定
set ANTHROPIC_API_KEY=YOYR_API_KEY

:: カスタム設定でスクリプトを実行
python outlook_assistant.py ^
  --emails 15 ^
  --days 14 ^
  --priority-keywords 至急,重要,緊急 ^
  --working-hours 0830 1730 ^
  --focus-time 10 12 ^
  --report-style detailed ^
  --api-key %ANTHROPIC_API_KEY%

:: 実行後にウィンドウを開いたままにする
pause
