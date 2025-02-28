@echo off
chcp 65001 > nul

title Outlook秘書アシスタント
echo Outlook秘書アシスタントを起動しています...

:: Pythonの仮想環境をアクティベート（環境に合わせて修正してください）
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
) else if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
) else (
    echo 仮想環境が見つかりません。
    echo まず、以下のコマンドで仮想環境を作成してください：
    echo python -m venv .venv
    echo その後、pip install -r requirements.txt を実行してください。
    pause
    exit /b 1
)

:: API KEYを環境変数として設定
set ANTHROPIC_API_KEY=YOUR_API_KEY_HERE

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
