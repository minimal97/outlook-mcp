# Outlook カレンダー MCP サーバー

AX0 から Outlook カレンダーを読み書きする自作MCPサーバー。
個人Microsoftアカウント（Outlook.com）対応。デバイスコード認証。

---

## ファイル構成

```
outlook-mcp/
├── README.md              # この取説
└── outlook_mcp_server.py  # MCPサーバー本体
```

実ファイル: `[scriptsDirectory]/outlook_mcp_server.py`
MCP設定: `~/.claude.json`（プロジェクト `[projectDirectory]`）

---

## セットアップ手順（初回のみ）

### 1. Azure アプリ登録

1. [portal.azure.com](https://portal.azure.com) → 「アプリの登録」→「新規登録」
2. サポートされるアカウントの種類: **個人用 Microsoft アカウントのみ**
3. リダイレクトURI: `http://localhost`

### 2. API権限を追加

「APIのアクセス許可」→「Microsoft Graph」→「委任されたアクセス許可」
- `Calendars.ReadWrite` を追加

### 3. パブリッククライアントフローを有効化

「認証」→「詳細設定」→「パブリック クライアント フローを許可する」→ **はい**

> ※ クライアントシークレットは不要（デバイスコード認証のため）

### 4. 環境構築

```bash
uv venv ~/.venv/outlook-mcp
uv pip install --python ~/.venv/outlook-mcp/bin/python msal requests mcp
```

### 5. MCP登録

```bash
claude mcp add outlook-mcp \
  -e OUTLOOK_CLIENT_ID="<クライアントID>" \
  -- ~/.venv/outlook-mcp/bin/python [scriptsDirectory]/outlook_mcp_server.py
```

---

## 初回認証

```
start_auth    → URLとコードが返ってくる
              → https://microsoft.com/link でサインイン
complete_auth → 成功したらトークン保存
```

以降はトークンが自動更新されるので再認証不要。

---

## 使えるツール

| ツール | 説明 |
|--------|------|
| `start_auth` | 認証開始（URL/コードを返す） |
| `complete_auth` | 認証完了・トークン保存 |
| `get_calendar_events` | 予定一覧取得（`days`で何日先まで） |
| `create_calendar_event` | 予定追加（`subject`, `start`, `end`） |

---

## コストについて

**Graph API + アプリ登録は無料。ずっと。**
MCPサーバーはローカル動作なのでAzureの課金は一切発生しない。

---

## トークンの保存先

- `~/.outlook_mcp_tokens.json` — 認証トークン（自動更新）
- `~/.outlook_mcp_flow.json` — 認証フロー一時ファイル（認証後削除）
