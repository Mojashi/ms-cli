# ms-cli 使い方

Microsoft Teams チャット + Outlook メールを操作するCLI。
Graph APIが使えない環境向けに、Teams/Outlook内部APIを利用。

## セットアップ

### 前提条件
- Node.js 22+
- Chrome で Teams (`teams.microsoft.com`) にログイン済み
- macOS (Chrome Cookie復号にKeychainアクセスが必要)

### インストール
```bash
cd ~/repos/ms-cli
pnpm install
```

PATHに登録済み (`~/.local/bin/ms-cli` → `~/repos/ms-cli/bin/ms-cli.mjs`)。
どこからでも `ms-cli` で実行可能。

## 認証

### 初回ログイン
```bash
ms-cli auth login
```
1. Chrome Cookie DB から `skypetoken_asm` を自動抽出
2. Cookie がなければ Puppeteer で Chrome を起動してログインフロー実行
3. トークンは `~/.ms-cli/config.json` に保存

### 手動トークン指定
```bash
ms-cli auth login --token <skypetoken_asm値>
ms-cli auth login --token <token> --refresh-token <MSAL refresh token>
```

### ステータス確認
```bash
ms-cli auth status
```
各トークンの有効期限を表示。

### トークンリフレッシュ
```bash
ms-cli auth refresh
```
MSAL refresh token を使って skypetoken を再取得。

## Teams チャット

### チャット一覧
```bash
ms-cli chat list                    # 直近20件
ms-cli chat list -n 50              # 50件表示
ms-cli chat list -u                 # 未読のみ
ms-cli chat list -t chat            # 1:1/グループチャットのみ
ms-cli chat list -t channel         # チャネルのみ
ms-cli chat list -t meeting         # 会議チャットのみ
```

出力にはタイプ別アイコン、未読マーカー `[NEW]`、最新メッセージプレビューが含まれる。
各チャットの `id:` 行がメッセージ読み取り等に使うID。

### メッセージ読み取り
```bash
ms-cli chat read <conversation-id>          # 直近20件
ms-cli chat read <conversation-id> -n 50    # 50件
ms-cli chat read <conversation-id> --json   # JSON出力
```

未読メッセージには `[NEW]` タグが表示される。

### メッセージ送信
```bash
ms-cli chat send <conversation-id> "メッセージ本文"
```

### 既読にする
```bash
ms-cli chat mark-read <conversation-id>
```

## Outlook メール

Outlook REST API v2.0 を使用。MSAL refresh token から自動的にOutlookトークンを取得・更新する。

### メール一覧
```bash
ms-cli mail list                    # 受信トレイ 直近15件
ms-cli mail list -n 30              # 30件表示
ms-cli mail list -u                 # 未読のみ
ms-cli mail list -f sentitems       # 送信済みフォルダ
```

`[NEW]` マーカーで未読を表示。添付ファイルありは `[+]`、重要度高は `[!]`。

### メール本文を読む
```bash
ms-cli mail read <message-id>              # テキスト表示
ms-cli mail read <message-id> --json       # JSON出力
```

`mail list` や `mail search` の出力に含まれる `id:` をそのまま使う。

### メール検索
```bash
ms-cli mail search "キーワード"              # 10件
ms-cli mail search "from:田中" -n 20        # 20件
```

### 下書き作成
```bash
ms-cli mail draft --to user@example.com -s "件名" -b "本文"
ms-cli mail draft --to user@example.com --cc boss@example.com -s "件名" -b "本文"
ms-cli mail draft --to user@example.com -s "件名" -b "<p>HTML本文</p>" --html
ms-cli mail draft --to user@example.com -s "件名" -b "本文" --importance High
```

下書きフォルダの確認: `ms-cli mail list -f drafts`

### 返信の下書き
```bash
ms-cli mail reply <message-id> -b "返信本文"          # 全員返信（デフォルト）
ms-cli mail reply <message-id> -b "返信本文" --no-all # 送信者のみに返信
```

件名 (`RE: ...`)・宛先・引用本文は自動セット。作成後 `mail send <id>` で送信。

### 下書きを送信
```bash
ms-cli mail send <message-id>
```

`mail draft` で返される `id:` や `mail list -f drafts` のIDを指定。

### メール作成→即送信
```bash
ms-cli mail compose --to user@example.com -s "件名" -b "本文"
ms-cli mail compose --to user@example.com --cc cc@example.com -s "件名" -b "本文"
```

下書きを経由せず直接送信する。

### Touch ID 認証

以下の送信系コマンドは実行時に **Touch ID (指紋認証)** が必須:
- `ms-cli mail send` (下書き送信)
- `ms-cli mail compose` (即送信)
- `ms-cli chat send` (Teams メッセージ送信)

Claude Code 経由で実行された場合もTouch IDダイアログが表示され、ユーザーが指紋で承認しない限り送信されない。

## 予定表 (カレンダー)

Outlook CalendarView API を使用。

### 今日の予定
```bash
ms-cli cal today
```

### 今後の予定一覧
```bash
ms-cli cal list                     # 7日分
ms-cli cal list -d 14               # 14日分
ms-cli cal list -d 3 -n 50          # 3日分、最大50件
```

日付ごとにグループ表示。ステータス色分け: `[Busy]`=赤, `[Tentative]`=黄, `[Free]`=グレー。

### 予定の詳細
```bash
ms-cli cal read <event-id>          # 出席者・場所・本文表示
ms-cli cal read <event-id> --json   # JSON出力
```

### 他ユーザーのスケジュール確認
```bash
ms-cli cal schedule user1@example.com user2@example.com
ms-cli cal schedule user1@example.com -d 2026-03-01              # 日付指定
ms-cli cal schedule user1@example.com --start-hour 10 --end-hour 18  # 時間帯指定
```

30分刻みのビジュアルバーで空き/仮/ビジー/OOFを表示。

### 空きスロット検索
```bash
ms-cli cal find-slot user1@example.com user2@example.com             # 60分、5営業日
ms-cli cal find-slot user1@example.com --duration 30 --days 3        # 30分、3営業日
ms-cli cal find-slot user1@example.com --start-hour 10 --end-hour 15 # 10-15時の範囲
ms-cli cal find-slot user1@example.com --start-date 2026-03-01       # 開始日指定
```

全参加者が空いている共通スロットを自動検索。土日はスキップ。

## Claude Code との連携

このCLIはClaude CodeがBashツール経由で呼び出すことを想定。

```
# 例: Claude Codeに「未読チャットを確認して」と頼むと内部的に:
ms-cli chat list -u

# 例: 「山田さんからのメール探して」
ms-cli mail search "山田"

# 例: 特定チャットの最新メッセージを読む
ms-cli chat read 19:abc123@thread.v2 -n 5

# 例: 「今日の予定教えて」
ms-cli cal today

# 例: 「来週の予定は？」
ms-cli cal list -d 7

# 例: 「田中さんにメール下書きして」
ms-cli mail draft --to tanaka@example.com -s "件名" -b "本文"
```

## 設定ファイル

`~/.ms-cli/config.json`:
```json
{
  "skypeToken": "eyJ...",
  "refreshToken": "1.AWs...",
  "outlookToken": "eyJ...",
  "tenantId": "<your-tenant-id>",
  "region": "<region>",
  "chatServiceHost": "<region>.ng.msg.teams.microsoft.com"
}
```

## トークン有効期限

| トークン | 有効期間 | 自動更新 |
|---------|---------|---------|
| skypetoken_asm | ~24h | `auth refresh` で手動更新 / Cookie再抽出 |
| Outlook Bearer | ~1h | `mail` コマンド実行時に自動更新 |
| MSAL refresh token | 長期間 | AADトークン取得時にローテーション |

## ファイル構成

```
~/repos/ms-cli/
├── bin/ms-cli.mjs          # CLI エントリポイント (PATH用)
├── src/
│   ├── index.ts            # commander サブコマンド定義
│   ├── api.ts              # Teams Chat API クライアント
│   ├── outlook-api.ts      # Outlook REST API クライアント
│   ├── auth.ts             # トークン管理・リフレッシュ
│   ├── config.ts           # 設定永続化 (~/.ms-cli/config.json)
│   ├── cookie-extractor.ts # Chrome Cookie DB復号
│   └── browser-login.ts    # Puppeteer ブラウザログイン
├── RESEARCH.md             # API調査結果
└── USAGE.md                # このファイル
```
