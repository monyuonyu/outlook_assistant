@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

title Outlook秘書アシスタント
echo Outlook秘書アシスタントを起動しています...

:: パスの設定
set "PYTHON_HOME=%~dp0python-embedded"
set "PATH=%PYTHON_HOME%;%PATH%"

:: 組み込みPythonの存在確認
if not exist "%PYTHON_HOME%\python.exe" (
    echo エラー: 組み込みPythonが見つかりません。
    echo python-embeddedフォルダが正しく配置されているか確認してください。
    pause
    exit /b 1
)

echo.
echo 設定内容:
echo - メール取得数: 15通
echo - 期間: 14日
echo - 優先キーワード: 至急,重要,緊急
echo - 勤務時間: 8:30-17:30
echo - 集中時間: 10:00-12:00
echo.

:: スクリプトの実行
"%PYTHON_HOME%\python.exe" outlook_assistant.py ^
    --emails 15 ^
    --days 14 ^
    --priority-keywords 至急,重要,緊急 ^
    --working-hours 0830 1730 ^
    --focus-time 10 12 ^
    --report-style detailed ^
    --api-key YOUR_ANTHROPIC_API_KEY

if %errorlevel% neq 0 (
    echo エラー: スクリプトの実行中にエラーが発生しました。
)

echo.
echo 処理が完了しました。
pause
