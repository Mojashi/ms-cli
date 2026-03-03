# ms-cli

Microsoft Teams チャット・Outlook メール・カレンダーをターミナルから操作する CLI ツール。

Graph API が使えない環境向けに、Teams / Outlook 内部 API を利用しています。macOS 専用。

## 機能

- **Teams チャット** — 一覧・閲覧・送信・既読マーク
- **Outlook メール** — 一覧・閲覧・検索・下書き・送信・返信・添付ファイル
- **カレンダー** — 今日の予定・一覧・詳細・複数ユーザーのスケジュール表示・空きスロット検索
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

### 1. Client ID を取得

Teams Web Client が使用している OAuth Client ID をブラウザから取得します。

1. [teams.microsoft.com](https://teams.microsoft.com) にログイン
2. DevTools を開く (F12)
3. **Application** → **Local Storage** → `https://teams.microsoft.com`
4. キー名に `client_id` を含む MSAL 関連エントリを探す（例: `login.microsoftonline.com-...`のキー内の JSON）
5. `client_id` の値（UUID 形式）をコピー

### 2. セットアップ & ログイン

```bash
ms-cli auth setup
```

Client ID の入力を求められるので、上で取得した値を貼り付けます。続けて Device Code Flow でブラウザ認証が始まります。

トークンは `~/.ms-cli/config.json` に保存されます。

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
| `clientId`        | OAuth アプリケーション ID (必須)                 |
| `skypeToken`      | Teams 内部 JWT (ログイン時に自動設定)            |
| `refreshToken`    | MSAL リフレッシュトークン (ログイン時に自動設定) |
| `tenantId`        | Azure AD テナント ID (ログイン時に自動検出)      |
| `region`          | リージョン (ログイン時に自動検出)                |
| `chatServiceHost` | Teams Chat API ホスト (ログイン時に自動設定)     |

## License

MIT
