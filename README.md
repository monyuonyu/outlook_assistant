# Outlook Assistant

**Outlook Assistant** はMicrosoft Outlookから未読メールとカレンダー予定を自動的に取得し、Claude AIを使用して秘書レポートを生成するPythonツールです。日々の情報管理を効率化し、重要なメールや予定の優先順位付けを支援します。

## 特徴

- 📧 Outlookの未読メールを自動取得・分析
- 📅 カレンダーの今後の予定を自動取得・整理
- 🔍 重要度に基づいてメールを自動分類（緊急/今日/週内/情報のみ）
- ✅ 優先度付きのタスクリストを自動生成
- 📊 効率的な1日および週間の計画を提案
- 📝 美しいMarkdown形式でレポートを出力
- ⚙️ 優先ドメイン、キーワード、勤務時間などをカスタマイズ可能

## 必要条件

- Python 3.7以上
- Microsoft Outlook（インストール済み）
- Anthropic API Key（Claude AIへのアクセス用）
- Windows OS（pywin32の依存関係のため）

## インストール

```bash
# 必要なパッケージをインストール
pip install -r requirements.txt
```

## 使用方法

### 基本的な使用方法

```bash
python outlook_assistant.py --api-key YOUR_ANTHROPIC_API_KEY
```

### 詳細なオプション

```bash
python outlook_assistant.py --emails 15 --days 10 --api-key YOUR_API_KEY --priority-domains company.com,client.com --priority-keywords 至急,重要,期限 --working-hours 0830 1730 --focus-time 10 12 --report-style detailed
```

### コマンドラインオプション

| オプション | 説明 | デフォルト値 |
|------------|------|------------|
| `--emails N` | 取得する未読メール数 | 10 |
| `--days N` | 取得する予定の日数 | 7 |
| `--api-key KEY` | Anthropic API Key | - |
| `--api-version VER` | Anthropic API バージョン | 2023-06-01 |
| `--priority-domains DOMAINS` | 優先ドメインをカンマ区切りで指定 | example.com,important-client.com |
| `--priority-keywords KEYWORDS` | 優先キーワードをカンマ区切りで指定 | 至急,重要,期限,緊急 |
| `--working-hours START END` | 勤務時間 | 9 18 |
| `--focus-time START END` | 集中作業時間 | 10 12 |
| `--report-style STYLE` | レポートスタイル (detailed/concise) | detailed |

## 出力

スクリプトを実行すると、以下の出力が生成されます：

- **assistant_report_{timestamp}.md** - Markdownフォーマットの秘書レポート

## トラブルシューティング

- **Outlookに接続できない**: Outlookがインストールされ、起動していることを確認してください。
- **API呼び出しエラー**: API Keyが正しく設定されているか確認してください。
