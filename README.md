# ms-cli

Microsoft Teams チャット・Outlook メール・カレンダーをターミナルから操作する CLI ツール。

## 機能

- **Teams チャット** — 一覧・閲覧・送信・既読マーク
- **Outlook メール** — 一覧・閲覧・検索・下書き・送信・返信・添付ファイル
- **カレンダー** — 今日の予定・一覧・詳細・複数ユーザーのスケジュール表示・空きスロット検索
- **SharePoint / OneDrive** — サイト・ファイルの検索、ダウンロード、アップロード
- **Microsoft Forms** — フォーム、質問、回答の参照
- **Touch ID** — 送信系操作は指紋認証が必須（Claude Code 経由でも安全）

## インストール

### バイナリ (推奨)

[Releases](../../releases) から macOS 向けバイナリをダウンロード:

```bash
# Apple Silicon
curl -L -o ms-cli https://github.com/Mojashi/ms-cli/releases/latest/download/ms-cli-darwin-arm64

# Intel Mac
curl -L -o ms-cli https://github.com/Mojashi/ms-cli/releases/latest/download/ms-cli-darwin-x64

chmod +x ms-cli
mv ms-cli /usr/local/bin/
```

### ソースから

```bash
git clone https://github.com/Mojashi/ms-cli.git
cd ms-cli
bun install
bun build src/index.ts --compile --outfile ms-cli
mv ms-cli /usr/local/bin/
```

## セットアップ

```bash
ms-cli auth login
```

macOS の専用 WebView が開き、Microsoft 365 の通常のログイン画面が表示されます。認証には Authorization Code Flow + PKCE を使い、Device Code Flow は使用しません。

取得した refresh token と各 API の access token は `~/.ms-cli/config.json` に保存されます。ファイルは所有ユーザーだけが読み書きできる mode `0600` で作成されます。

### 複数テナント (ゲスト含む)

`ms-cli auth login` すると、**所属する全テナント（ゲスト招待先含む）を自動検出して登録**します。
ゲストテナントへも 1 回のログインで利用可能になり、テナントを意識せず操作できます。

```bash
ms-cli auth login        # ログイン → 全テナントを自動検出・登録
ms-cli auth sync         # ログイン済みの状態で、未登録テナントを再検出
ms-cli auth list         # 登録済みアカウント一覧 (* = 現在)
ms-cli auth use <ref>    # 既定アカウントを切替 (key / tenantId / テナント名の一部で指定可)
ms-cli auth remove <ref> # アカウントを削除
```

**テナント横断の挙動** (基本的に `auth use` は不要):

- `chat list` — 全テナントのチャットを**混在・新しい順**で表示 (各行に `[テナント名]` を付記)
- `chat read` / `mail read` 等の ID 指定 — どのテナントの ID かを**自動判定**
- `mail` / `cal` — メールボックスを持つテナント (メンバー) のみ対象。ゲストは自動スキップ
- `-a, --account <ref>` で特定テナントに限定。`MS_CLI_ACCOUNT=<ref>` で一時上書きも可

> 仕組み: 1 つの refresh token を各テナントへ redeem し、Teams の `tenantsv2` API で所属テナントを列挙しています。

## 使い方

```bash
# Teams
ms-cli chat list                 # チャット一覧
ms-cli chat list -u              # 未読のみ
ms-cli chat read <id>            # メッセージ閲覧
ms-cli chat send <id> "Hello"    # メッセージ送信 (Touch ID)

# メール
ms-cli mail list                 # 受信トレイ
ms-cli mail list -u              # 未読のみ
ms-cli mail read <id>            # メール本文
ms-cli mail search "keyword"     # 検索
ms-cli mail draft --to user@example.com -s "件名" -b "本文"
ms-cli mail send <id>            # 下書き送信 (Touch ID)

# カレンダー
ms-cli cal today                 # 今日の予定
ms-cli cal list -d 7             # 7日分
ms-cli cal schedule user1@example.com user2@example.com
ms-cli cal find-slot user1@example.com --duration 30
```

詳細は [USAGE.md](USAGE.md) を参照。

## Claude Code との連携

このCLIは [Claude Code](https://claude.com/claude-code) の Bash ツール経由で呼び出すことを想定しています。

```
「未読チャットを確認して」     → ms-cli chat list -u
「山田さんからのメール探して」  → ms-cli mail search "山田"
「今日の予定教えて」           → ms-cli cal today
```

送信系コマンドは Touch ID が必須のため、Claude Code が勝手にメッセージを送信することはありません。

## 設定ファイル

`~/.ms-cli/config.json`:

| フィールド        | 説明                                             |
| ----------------- | ------------------------------------------------ |
| `skypeToken`      | Teams 内部 JWT (ログイン時に自動設定)            |
| `refreshToken`    | OAuth リフレッシュトークン (ログイン時に自動設定) |
| `graphToken`      | Microsoft Graph access token                     |
| `outlookToken`    | Outlook access token                             |
| `formsToken`      | Microsoft Forms access token                     |
| `tenantId`        | Azure AD テナント ID (ログイン時に自動検出)      |
| `region`          | リージョン (ログイン時に自動検出)                |
| `chatServiceHost` | Teams Chat API ホスト (ログイン時に自動設定)     |

refresh token は Graph、Outlook、Forms、Teams チャット用 token の自動更新に利用されます。`auth remove` を実行すると、対象アカウントの token を含む設定が削除されます。

## License

MIT
