#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Claude APIクライアントモジュール
"""

import requests
from datetime import datetime
import os

class ClaudeClient:
    def __init__(self, api_key, api_version):
        """
        初期化
        
        Args:
            api_key (str): Anthropic API Key
            api_version (str): API バージョン
        """
        self.api_key = api_key
        self.api_version = api_version
        self.api_url = "https://api.anthropic.com/v1/messages"

    def create_prompt(self, emails_data, events_data, settings=None):
        """
        プロンプトを作成
        
        Args:
            emails_data (list): 未読メールのリスト
            events_data (list): 予定のリスト
            settings (dict): カスタム設定
            
        Returns:
            str: Claudeに送るプロンプト
        """
        today = datetime.now().strftime("%Y年%m月%d日")
        
        # デフォルト設定を取得（settings引数が無効な場合に使用）
        from config import DEFAULT_SETTINGS
        settings = settings or DEFAULT_SETTINGS
        
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
                        prompt += f"- {event['subject']}\n"

        # 分析指示を追加
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

    def call_api(self, prompt):
        """
        Claude APIを呼び出す
        
        Args:
            prompt (str): 送信するプロンプト
            
        Returns:
            str: Claudeからの応答
        """
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": self.api_version,
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
            response = requests.post(self.api_url, headers=headers, json=data)
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

    def save_response(self, response):
        """
        応答をファイルに保存
        
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
