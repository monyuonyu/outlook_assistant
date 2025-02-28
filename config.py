#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
設定管理モジュール
"""

import argparse

# デフォルト設定
DEFAULT_SETTINGS = {
    "priority_domains": ["example.com", "important-client.com"],
    "priority_keywords": ["至急", "重要", "期限", "緊急"],
    "working_hours": {"start": 9, "end": 18},
    "focus_time": {"start": 10, "end": 12},
    "report_style": "detailed"
}

# Anthropic API用の設定
API_KEY = "YOUR_ANTHROPIC_API_KEY"
API_URL = "https://api.anthropic.com/v1/messages"
API_VERSION = "2023-06-01"

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

def create_arg_parser():
    """コマンドライン引数パーサーを作成"""
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
    return parser

def load_settings(args):
    """コマンドライン引数から設定を読み込む"""
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
    
    # デフォルト設定とマージ
    merged_settings = DEFAULT_SETTINGS.copy()
    for key, value in settings.items():
        if value:
            merged_settings[key] = value
    
    return merged_settings
