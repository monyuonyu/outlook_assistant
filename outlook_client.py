#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Outlookクライアントモジュール
"""

import win32com.client
import pythoncom
import traceback
from datetime import datetime, timedelta
import re

class OutlookClient:
    def __init__(self):
        """Outlookクライアントの初期化"""
        self.outlook = None
        self.namespace = None
        
    def connect(self):
        """Outlookに接続"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except pythoncom.com_error as e:
            print(f"Outlookへの接続に失敗しました: {e}")
            return False

    def get_unread_emails(self, max_emails=10):
        """未読メールを取得"""
        emails_data = []
        
        try:
            if not self.connect():
                return emails_data
                
            # 受信トレイを取得
            print("受信トレイにアクセス中...")
            try:
                inbox = self.namespace.GetDefaultFolder(6)  # 6は受信トレイを表す定数
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
                    email_data = self._process_email(email, i)
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

    def _process_email(self, email, index):
        """メールオブジェクトからデータを抽出"""
        email_data = {
            "id": index,
            "subject": getattr(email, "Subject", ""),
            "sender": getattr(email, "SenderName", ""),
            "sender_email": getattr(email, "SenderEmailAddress", ""),
            "received_time": str(getattr(email, "ReceivedTime", "")),
            "body": getattr(email, "Body", "").strip(),
            "has_attachments": getattr(email, "Attachments", None) is not None and getattr(email, "Attachments.Count", 0) > 0,
        }
        
        # 添付ファイル情報の取得
        if email_data["has_attachments"]:
            email_data["attachments"] = self._get_attachments_info(email)
        
        return email_data

    def _get_attachments_info(self, email):
        """メールの添付ファイル情報を取得"""
        attachments_info = []
        try:
            for j in range(1, email.Attachments.Count + 1):
                attachment = email.Attachments.Item(j)
                attachments_info.append({
                    "filename": attachment.FileName,
                    "size": attachment.Size
                })
        except Exception as e:
            print(f"  添付ファイル情報取得中にエラー: {e}")
        return attachments_info

    def get_calendar_events(self, days_ahead=7):
        """予定を取得"""
        events_data = []
        
        try:
            if not self.connect():
                return events_data
                
            # カレンダーフォルダを取得
            print("カレンダーフォルダにアクセス中...")
            try:
                calendar = self.namespace.GetDefaultFolder(9)  # 9はカレンダーを表す定数
            except Exception as e:
                print(f"カレンダーフォルダの取得に失敗しました: {e}")
                return events_data
                
            # 日付範囲を設定
            start_date = datetime.now()
            end_date = start_date + timedelta(days=days_ahead)
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
                        event_data = self._process_appointment(appointment, i)
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

    def _process_appointment(self, appointment, index):
        """予定オブジェクトからデータを抽出"""
        event_data = {
            "id": index,
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
        
        # 参加者情報の取得
        self._add_attendees_info(appointment, event_data)
        
        # オンライン会議URLの抽出
        meeting_url = self._extract_meeting_url(event_data["body"])
        if meeting_url:
            event_data["meeting_url"] = meeting_url
            
        return event_data

    def _add_attendees_info(self, appointment, event_data):
        """予定の参加者情報を取得"""
        try:
            if hasattr(appointment, "RequiredAttendees"):
                event_data["required_attendees"] = appointment.RequiredAttendees
            if hasattr(appointment, "OptionalAttendees"):
                event_data["optional_attendees"] = appointment.OptionalAttendees
        except Exception as e:
            print(f"  参加者情報取得中にエラー: {e}")

    def _extract_meeting_url(self, body):
        """本文からオンライン会議URLを抽出"""
        try:
            body = body.lower()
            
            # Teams会議リンクの検索
            if "teams.microsoft.com" in body:
                teams_pattern = r'https://teams\.microsoft\.com/l/meetup-join/[^\s<>"]+'
                teams_match = re.search(teams_pattern, body)
                if teams_match:
                    return teams_match.group(0)
            
            # Zoom会議リンクの検索
            elif "zoom.us" in body:
                zoom_pattern = r'https://[a-zA-Z0-9-.]+zoom\.us/[^\s<>"]+'
                zoom_match = re.search(zoom_pattern, body)
                if zoom_match:
                    return zoom_match.group(0)
                    
            return None
            
        except Exception:
            return None
