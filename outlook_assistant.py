#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Outlook Assistant

このスクリプトはMicrosoft Outlookから未読メールとカレンダー予定を取得し、
Claude APIを使用して秘書レポートを生成します。日々の情報管理を効率化し、
重要なメールや予定の管理を支援します。

必要なライブラリ:
- win32com.client (pywin32)
- requests
- その他標準ライブラリ

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

import win32com.client
import sys
import time
import traceback
import pythoncom
from datetime import datetime, timedelta
import os
import requests
import argparse

# Anthropic API用の設定
API_KEY = "YOUR_ANTHROPIC_API_KEY"  # 実際の使用時に自分のAPI keyを設定
API_URL = "https://api.anthropic.com/v1/messages"
API_VERSION = "2023-06-01"  # 最新のバージョン

def get_unread_emails(max_emails=10):
    """
    Outlookから未読メールを最大指定件数取得する
    
    Args:
        max_emails (int): 取得する未読メールの最大件数
    
    Returns:
        list: 未読メールのリスト
    """
    emails_data = []
    
    try:
        # Outlookのインスタンスを取得
        print(f"Outlookに接続中...")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except pythoncom.com_error as e:
            print(f"Outlookへの接続に失敗しました: {e}")
            return emails_data
        
        # 受信トレイを取得
        print("受信トレイにアクセス中...")
        try:
            inbox = namespace.GetDefaultFolder(6)  # 6は受信トレイを表す定数
        except Exception as e:
            print(f"受信トレイの取得に失敗しました: {e}")
            return emails_data
            
        # 未読メールの件数を確認
        try:
            unread_count = inbox.UnReadItemCount
            print(f"未読メール数: {unread_count}")
            if unread_count == 0:
                print("未読メールはありません。")
                return emails_data
        except Exception as e:
            print(f"未読メール数の取得に失敗しました: {e}")
            
        # 未読メールだけをフィルタリング
        try:
            filter_string = "[Unread]=True"
            unread_items = inbox.Items.Restrict(filter_string)
            unread_items.Sort("[ReceivedTime]", True)  # 受信日時の降順でソート
        except Exception as e:
            print(f"未読メールのフィルタリングに失敗しました: {e}")
            return emails_data
            
        # 取得する件数を決定
        try:
            emails_to_process = min(max_emails, unread_items.Count)
            print(f"処理する未読メール: {emails_to_process}件")
        except Exception as e:
            print(f"未読メール数の計算に失敗しました: {e}")
            return emails_data
        
        # 未読メールを処理
        print("未読メールを取得中...")
        for i in range(1, emails_to_process + 1):
            try:
                email = unread_items.Item(i)
                
                # メールデータを辞書に格納
                email_data = {
                    "id": i,
                    "subject": getattr(email, "Subject", ""),
                    "sender": getattr(email, "SenderName", ""),
                    "sender_email": getattr(email, "SenderEmailAddress", ""),
                    "received_time": str(getattr(email, "ReceivedTime", "")),
                    "body": getattr(email, "Body", "").strip(),
                    "has_attachments": getattr(email, "Attachments", None) is not None and getattr(email, "Attachments.Count", 0) > 0,
                }
                
                # 添付ファイル情報の取得
                if email_data["has_attachments"]:
                    attachments_info = []
                    try:
                        for j in range(1, email.Attachments.Count + 1):
                            attachment = email.Attachments.Item(j)
                            attachments_info.append({
                                "filename": attachment.FileName,
                                "size": attachment.Size
                            })
                        email_data["attachments"] = attachments_info
                    except Exception as e:
                        print(f"  メール {i} の添付ファイル情報取得中にエラー: {e}")
                        email_data["attachments"] = []
                
                # メールデータをリストに追加
                emails_data.append(email_data)
                print(f"  メール {i} の情報を取得しました: {email_data['subject']}")
                
            except Exception as e:
                print(f"  メール {i} の処理中にエラー: {e}")
                traceback.print_exc()
        
        return emails_data
        
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
        return emails_data

def get_calendar_events(days_ahead=7):
    """
    Outlookカレンダーから指定日数先までの予定を取得する
    
    Args:
        days_ahead (int): 何日先までの予定を取得するか
    
    Returns:
        list: 予定のリスト
    """
    events_data = []
    
    try:
        # Outlookのインスタンスを取得
        print(f"Outlookに接続中...")
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except pythoncom.com_error as e:
            print(f"Outlookへの接続に失敗しました: {e}")
            return events_data
        
        # カレンダーフォルダを取得
        print("カレンダーフォルダにアクセス中...")
        try:
            calendar = namespace.GetDefaultFolder(9)  # 9はカレンダーを表す定数
        except Exception as e:
            print(f"カレンダーフォルダの取得に失敗しました: {e}")
            return events_data
            
        # 日付範囲を設定
        start_date = datetime.now()
        end_date = start_date + timedelta(days=days_ahead)
        
        # 日付をOutlookのフィルタリング用の文字列形式に変換
        start_date_str = start_date.strftime("%m/%d/%Y %H:%M %p")
        end_date_str = end_date.strftime("%m/%d/%Y %H:%M %p")
        
        # 指定期間内の予定を取得
        try:
            filter_string = f"[Start] >= '{start_date_str}' AND [End] <= '{end_date_str}'"
            appointments = calendar.Items.Restrict(filter_string)
            appointments.Sort("[Start]")  # 開始日時でソート
            print(f"期間内の予定数: {appointments.Count}")
        except Exception as e:
            print(f"予定の取得に失敗しました: {e}")
            traceback.print_exc()
            return events_data
            
        # 各予定を処理
        print(f"予定を取得中... ({start_date.strftime('%Y/%m/%d')} から {end_date.strftime('%Y/%m/%d')} まで)")
        
        if appointments.Count == 0:
            print("  指定期間内に予定はありません。")
        else:
            for i in range(1, appointments.Count + 1):
                try:
                    appointment = appointments.Item(i)
                    
                    # 予定データを辞書に格納
                    event_data = {
                        "id": i,
                        "subject": getattr(appointment, "Subject", ""),
                        "start": str(getattr(appointment, "Start", "")),
                        "end": str(getattr(appointment, "End", "")),
                        "location": getattr(appointment, "Location", ""),
                        "body": getattr(appointment, "Body", "").strip(),
                        "organizer": getattr(appointment, "Organizer", ""),
                        "is_recurring": getattr(appointment, "IsRecurring", False),
                        "is_all_day_event": getattr(appointment, "AllDayEvent", False),
                        "importance": str(getattr(appointment, "Importance", "")),
                        "sensitivity": str(getattr(appointment, "Sensitivity", "")),
                        "meeting_status": str(getattr(appointment, "MeetingStatus", "")),
                    }
                    
                    # 必須参加者と任意参加者の取得
                    try:
                        if hasattr(appointment, "RequiredAttendees"):
                            event_data["required_attendees"] = appointment.RequiredAttendees
                        if hasattr(appointment, "OptionalAttendees"):
                            event_data["optional_attendees"] = appointment.OptionalAttendees
                    except Exception as e:
                        print(f"  予定 {i} の参加者情報取得中にエラー: {e}")
                    
                    # Teams/Zoomなどの会議URLを抽出する試み
                    try:
                        body = event_data["body"].lower()
                        meeting_url = None
                        
                        # Teams会議リンクの検索
                        if "teams.microsoft.com" in body:
                            import re
                            teams_pattern = r'https://teams\.microsoft\.com/l/meetup-join/[^\s<>"]+'
                            teams_match = re.search(teams_pattern, body)
                            if teams_match:
                                meeting_url = teams_match.group(0)
                        
                        # Zoom会議リンクの検索
                        elif "zoom.us" in body:
                            import re
                            zoom_pattern = r'https://[a-zA-Z0-9-.]+zoom\.us/[^\s<>"]+'
                            zoom_match = re.search(zoom_pattern, body)
                            if zoom_match:
                                meeting_url = zoom_match.group(0)
                                
                        if meeting_url:
                            event_data["meeting_url"] = meeting_url
                    except Exception as e:
                        # 会議URL抽出のエラーは無視
                        pass
                    
                    # 予定データをリストに追加
                    events_data.append(event_data)
                    
                    # 簡易情報を表示
                    start_time = appointment.Start.strftime("%Y/%m/%d %H:%M") if hasattr(appointment, "Start") else "不明"
                    print(f"  予定 {i}: {start_time} - {event_data['subject']}")
                    
                except Exception as e:
                    print(f"  予定 {i} の処理中にエラー: {e}")
                    traceback.print_exc()
        
        return events_data
        
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
        return events_data

def create_assistant_prompt(emails_data, events_data, settings=None):
    """
    メールと予定データからClaudeに送るプロンプトを作成する
    
    Args:
        emails_data (list): 未読メールのリスト
        events_data (list): 予定のリスト
        settings (dict): カスタム設定
        
    Returns:
        str: Claudeに送るプロンプト
    """
    today = datetime.now().strftime("%Y年%m月%d日")
    
    # デフォルト設定
    default_settings = {
        "priority_domains": ["example.com", "important-client.com"],  # 優先ドメイン
        "priority_keywords": ["至急", "重要", "期限", "緊急"],  # 優先キーワード
        "working_hours": {"start": 9, "end": 18},  # 勤務時間
        "focus_time": {"start": 10, "end": 12},  # 集中作業時間
        "report_style": "detailed"  # レポートスタイル（detailed/concise）
    }
    
    # 設定のマージ
    if settings:
        for key, value in settings.items():
            if key in default_settings and value:
                default_settings[key] = value
    
    settings = default_settings
    
    prompt = f"""
あなたは経験豊富なエグゼクティブアシスタントです。今日は{today}です。
以下の情報を基に、効率的なタスク管理と意思決定をサポートするレポートを作成してください。

# 分析のためのガイドライン
- 優先ドメイン: {', '.join(settings['priority_domains'])}
- 優先キーワード: {', '.join(settings['priority_keywords'])}
- 勤務時間: {settings['working_hours']['start']}時～{settings['working_hours']['end']}時
- 集中作業時間: {settings['focus_time']['start']}時～{settings['focus_time']['end']}時
- レポートスタイル: {settings['report_style']}

# 未読メール（最新{len(emails_data)}件）
"""

    # メール情報の追加
    if not emails_data:
        prompt += "未読メールはありません。\n\n"
    else:
        for i, email in enumerate(emails_data, 1):
            prompt += f"""
## メール {i}: {email['subject']}
- 送信者: {email['sender']} ({email['sender_email']})
- 受信日時: {email['received_time']}
- 添付ファイル: {'あり' if email.get('has_attachments', False) else 'なし'}

{email['body'][:500]}{"..." if len(email['body']) > 500 else ""}

"""

    # 予定情報の追加
    prompt += f"\n# 今後{len(events_data)}件の予定\n"
    
    if not events_data:
        prompt += "予定はありません。\n\n"
    else:
        # 日付ごとにグループ化
        events_by_date = {}
        for event in events_data:
            try:
                start_datetime = datetime.strptime(event['start'].split('+')[0].split('.')[0], '%Y-%m-%d %H:%M:%S')
                date_str = start_datetime.strftime('%Y年%m月%d日(%a)')
                
                if date_str not in events_by_date:
                    events_by_date[date_str] = []
                    
                events_by_date[date_str].append(event)
            except Exception:
                # 日付解析に失敗した場合はスキップ
                continue
        
        # 日付順に予定を表示
        for date_str, day_events in sorted(events_by_date.items()):
            prompt += f"\n## {date_str}\n"
            
            for event in sorted(day_events, key=lambda x: x['start']):
                try:
                    start_time = datetime.strptime(event['start'].split('+')[0].split('.')[0], '%Y-%m-%d %H:%M:%S').strftime('%H:%M')
                    end_time = datetime.strptime(event['end'].split('+')[0].split('.')[0], '%Y-%m-%d %H:%M:%S').strftime('%H:%M')
                    
                    prompt += f"""
- {start_time}～{end_time} {event['subject']}
  場所: {event['location']}
  {"オンライン会議URL: " + event.get('meeting_url', 'なし') if 'meeting_url' in event else ""}
  {"参加者: " + event.get('required_attendees', '記載なし') if 'required_attendees' in event else ""}
"""
                except Exception:
                    # 時間解析に失敗した場合は基本情報のみ
                    prompt += f"- {event['subject']}\n"

    # 指示を追加
    prompt += """
# 分析とレポート作成の指示

## 1. メール分析
- メールを「緊急対応」「今日中に対応」「週内対応」「情報のみ」に分類してください
- 各カテゴリのメールについて簡潔な要約と推奨アクションを提示してください
- 特に優先ドメインや優先キーワードを含むメールに注目してください

## 2. 予定分析
- 今後の予定を時系列で整理し、準備が必要なものを特定してください
- 予定と予定の間の移動時間や準備時間を考慮した現実的なスケジュールを提案してください
- 集中作業時間を確保できるよう、予定の調整案があれば提示してください

## 3. タスク管理
- メールと予定から抽出した具体的なタスクリストを作成してください
- 各タスクに優先度（高/中/低）と対応期限を設定してください
- 「今日必ず完了すべきこと」のショートリスト（3項目以内）を提示してください

## 4. 今後の計画
- 今日から1週間の効率的な業務計画を提案してください
- 重要な締め切りや準備が必要なイベントを強調してください
- 週末までに完了すべき主要タスクを特定してください

秘書としての経験と判断力を活かし、意思決定を支援する具体的で実用的なアドバイスを提供してください。
情報の重要度に応じて簡潔にまとめ、すぐに行動に移せる形式で提示してください。

必ずMarkdown形式でレポートを作成してください。見出し、箇条書き、強調などのMarkdown記法を適切に活用して、読みやすく構造化されたレポートにしてください。
"""

    return prompt

def call_claude_api(prompt):
    """
    ClaudeのAPIを呼び出す関数
    
    Args:
        prompt (str): 送信するプロンプト
        
    Returns:
        str: Claudeからの応答
    """
    headers = {
        "x-api-key": API_KEY,
        "anthropic-version": API_VERSION,
        "content-type": "application/json"
    }
    
    data = {
        "model": "claude-3-7-sonnet-20250219",  # 最新のモデル
        "max_tokens": 4000,
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ],
        "system": "あなたは優秀な秘書です。メールや予定表の情報を整理し、優先順位をつけてMarkdown形式で簡潔にまとめてください。重要なタスクを特定し、具体的なアドバイスを提供してください。Markdown形式の見出し、箇条書き、強調などを活用して、読みやすく構造化されたレポートを作成してください。"
    }
    
    try:
        response = requests.post(API_URL, headers=headers, json=data)
        response.raise_for_status()  # エラーチェック
        
        result = response.json()
        if "content" in result and len(result["content"]) > 0 and "text" in result["content"][0]:
            response_text = result["content"][0]["text"]
            
            # 重複部分を検出して削除する処理
            lines = response_text.split('\n')
            unique_lines = []
            seen = set()
            
            for line in lines:
                line_key = line.strip()
                if line_key and line_key not in seen:
                    seen.add(line_key)
                    unique_lines.append(line)
            
            return '\n'.join(unique_lines)
        else:
            return "APIからの応答で予期しない形式が返されました。"
            
    except Exception as e:
        error_details = f"{str(e)}"
        if hasattr(e, 'response') and e.response:
            try:
                error_details += f" - Response: {e.response.text}"
            except:
                pass
        return f"APIの呼び出し中にエラーが発生しました: {error_details}"

def save_assistant_response(response):
    """
    Claudeからの応答をMarkdown形式で保存する関数
    
    Args:
        response (str): 保存する応答内容
        
    Returns:
        str: 保存したファイルのパス
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"assistant_report_{timestamp}.md"
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(response)
    
    return os.path.abspath(filename)

def parse_domain_list(domain_str):
    """カンマ区切りのドメインリストを解析する"""
    if not domain_str:
        return None
    return [d.strip() for d in domain_str.split(',') if d.strip()]

def parse_keyword_list(keyword_str):
    """カンマ区切りのキーワードリストを解析する"""
    if not keyword_str:
        return None
    return [k.strip() for k in keyword_str.split(',') if k.strip()]

def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description='Outlook情報をClaudeに送って秘書レポートを作成',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument('--emails', type=int, default=10, help='取得する未読メールの最大件数')
    parser.add_argument('--days', type=int, default=7, help='何日先までの予定を取得するか')
    parser.add_argument('--api-key', type=str, help='Anthropic API Key（設定済みの場合は不要）')
    parser.add_argument('--api-version', type=str, help='Anthropic API Version（設定済みの場合は不要）')
    parser.add_argument('--priority-domains', type=str, help='優先ドメインをカンマ区切りで指定（例: company.com,client.com）')
    parser.add_argument('--priority-keywords', type=str, help='優先キーワードをカンマ区切りで指定（例: 至急,重要,期限）')
    parser.add_argument('--working-hours', type=int, nargs=2, metavar=('START', 'END'), help='勤務時間（例: 9 18）')
    parser.add_argument('--focus-time', type=int, nargs=2, metavar=('START', 'END'), help='集中作業時間（例: 10 12）')
    parser.add_argument('--report-style', type=str, choices=['detailed', 'concise'], default='detailed', help='レポートスタイル')
    
    args = parser.parse_args()
    
    print("Outlook秘書アシスタント")
    print("=" * 40)
    
    try:
        # API KEYの設定
        global API_KEY, API_VERSION
        if args.api_key:
            API_KEY = args.api_key
        
        if args.api_version:
            API_VERSION = args.api_version
        
        if not API_KEY or API_KEY == "YOUR_ANTHROPIC_API_KEY":
            print("Anthropic API Keyが設定されていません。")
            api_key = input("API Keyを入力するか、--api-keyオプションで指定してください: ")
            if api_key:
                API_KEY = api_key
            else:
                print("API Keyが指定されていないため終了します。")
                return
        
        # 設定の準備
        settings = {}
        if args.priority_domains:
            settings["priority_domains"] = parse_domain_list(args.priority_domains)
        if args.priority_keywords:
            settings["priority_keywords"] = parse_keyword_list(args.priority_keywords)
        if args.working_hours:
            settings["working_hours"] = {"start": args.working_hours[0], "end": args.working_hours[1]}
        if args.focus_time:
            settings["focus_time"] = {"start": args.focus_time[0], "end": args.focus_time[1]}
        if args.report_style:
            settings["report_style"] = args.report_style
        
        # メールと予定の取得
        print(f"\n1. 未読メールを最大{args.emails}件取得します...")
        # 未読メール取得
        try:
            emails_data = get_unread_emails(args.emails)
            print(f"  {len(emails_data)}件のメールを取得しました。")
        except Exception as e:
            print(f"  メール処理中にエラー: {e}")
            emails_data = []
        
        # 予定取得
        print(f"\n2. 今後{args.days}日間の予定を取得します...")
        try:
            events_data = get_calendar_events(args.days)
            print(f"  {len(events_data)}件の予定を取得しました。")
        except Exception as e:
            print(f"  予定処理中にエラー: {e}")
            events_data = []
            
        # プロンプトの作成
        print("\n3. 秘書アシスタント用のプロンプトを作成しています...")
        prompt = create_assistant_prompt(emails_data, events_data, settings)
        
        # Claudeの呼び出し
        print("\n4. Claude APIを呼び出しています...")
        print("  APIからの応答を待っています...")
        response = call_claude_api(prompt)
        
        # 結果の保存と表示
        print("\n5. 秘書レポートを保存しています...")
        report_path = save_assistant_response(response)
        
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
