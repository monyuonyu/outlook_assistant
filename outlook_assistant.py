#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Outlook Assistant

このスクリプトはMicrosoft Outlookから未読メールとカレンダー予定を取得し、
Claude APIを使用して秘書レポートを生成します。日々の情報管理を効率化し、
重要なメールや予定の管理を支援します。

使用方法:
python outlook_assistant.py [オプション]

オプション:
--emails N        : 取得する未読メール数 (デフォルト: 10)
--days N          : 取得する予定の日数 (デフォルト: 7)
--api-key KEY     : Anthropic API Key
--api-version VER : Anthropic API バージョン
--priority-domains DOMAINS : 優先ドメインをカンマ区切りで指定
--priority-keywords KEYWORDS : 優先キーワードをカンマ区切りで指定
--working-hours START END : 勤務時間 (例: --working-hours 9 18)
--focus-time START END : 集中作業時間 (例: --focus-time 10 12)
--report-style STYLE : レポートスタイル (detailed/concise)
"""

import sys
import time
import traceback

from config import create_arg_parser, load_settings, API_KEY, API_VERSION
from outlook_client import OutlookClient
from claude_client import ClaudeClient

def main():
    """メイン関数"""
    parser = create_arg_parser()
    args = parser.parse_args()
    
    print("Outlook秘書アシスタント")
    print("=" * 40)
    
    try:
        # API KEYの検証と設定
        api_key = args.api_key or API_KEY
        api_version = args.api_version or API_VERSION
        
        if not api_key or api_key == "YOUR_ANTHROPIC_API_KEY":
            print("Anthropic API Keyが設定されていません。")
            api_key = input("API Keyを入力するか、--api-keyオプションで指定してください: ")
            if not api_key:
                print("API Keyが指定されていないため終了します。")
                return
        
        # 設定の読み込み
        settings = load_settings(args)
        
        # Outlookクライアントの初期化
        outlook = OutlookClient()
        
        # メールと予定の取得
        print(f"\n1. 未読メールを最大{args.emails}件取得します...")
        try:
            emails_data = outlook.get_unread_emails(args.emails)
            print(f"  {len(emails_data)}件のメールを取得しました。")
        except Exception as e:
            print(f"  メール処理中にエラー: {e}")
            emails_data = []
        
        print(f"\n2. 今後{args.days}日間の予定を取得します...")
        try:
            events_data = outlook.get_calendar_events(args.days)
            print(f"  {len(events_data)}件の予定を取得しました。")
        except Exception as e:
            print(f"  予定処理中にエラー: {e}")
            events_data = []
            
        # Claudeクライアントの初期化
        claude = ClaudeClient(api_key, api_version)
            
        # プロンプトの作成とAPI呼び出し
        print("\n3. 秘書アシスタント用のプロンプトを作成しています...")
        prompt = claude.create_prompt(emails_data, events_data, settings)
        
        print("\n4. Claude APIを呼び出しています...")
        print("  APIからの応答を待っています...")
        response = claude.call_api(prompt)
        
        # 結果の保存と表示
        print("\n5. 秘書レポートを保存しています...")
        report_path = claude.save_response(response)
        
        print(f"\n秘書レポートを保存しました: {report_path}")
        print("\n========= 秘書レポート =========")
        print(response[:1000] + ("..." if len(response) > 1000 else ""))
        print("================================")
        
        print("\nレポート全文は保存されたファイルで確認できます。")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        traceback.print_exc()
    
    finally:
        print("\n5秒後に終了します...")
        time.sleep(5)

if __name__ == "__main__":
    main()
