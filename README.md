# bsky-follow (GAS + Sheets)

Bluesky API でキーワード検索し、投稿ユーザー一覧をスプレッドシートに同期、
シート上の `action` 列からフォロー/アンフォローを自動実行する最小構成です。

## 使い方

1. スプレッドシートを作成し、`拡張機能 → Apps Script` でこのリポジトリの `Code.js` と `appsscript.json` を貼り付けます。
2. スクリプトエディタで `setup()` を実行します（初回のみ）。
3. `Config` シートに下記を入力します。
   - `IDENTIFIER`: Bluesky のハンドル
   - `APP_PASSWORD`: アプリパスワード
   - `PDS_HOST`: 例 `https://bsky.social`（未入力なら既定で bsky.social）
   - `SEARCH_QUERY`: 検索キーワード
   - `SEARCH_LIMIT`: 1回の検索件数（1〜100）
   - `MAX_PAGES`: ページ数（1〜20）
   - `SEARCH_SORT`: `latest` or `top`（任意）
   - `SEARCH_LANG`: 言語コード（例 `ja`、任意）
   - `WEBHOOK_TOKEN`: Web App で呼ぶ場合のトークン（任意）
4. メニュー `Bsky Follow → Search & sync` を実行。
   - `Posts` と `Users` シートが更新されます。
   - `Posts.post_uri` は Bluesky の投稿URLになります。
   - `Users.user_url` は Bluesky のユーザーページURLになります。
5. `Users` シートの `action` 列に `follow` / `unfollow` を入れて
   `Bsky Follow → Apply follow/unfollow` を実行。

## デプロイ（Web App）

1. Apps Script 画面で `デプロイ → 新しいデプロイ`。
2. 種類は `ウェブアプリ`、`実行ユーザー: 自分`、`アクセス: 自分のみ`。
3. 公開URLに `?action=search` もしくは `?action=apply` を付けて呼び出します。
   - `WEBHOOK_TOKEN` を設定した場合は `?token=...` も付与してください。

## 注意

- アンフォローは `follow_uri` が入っている行のみ対応します（このツールでフォローした行）。
- 1回の実行で大量の操作をするとレート制限の可能性があります。必要なら `MAX_PAGES` や実行間隔を調整してください。
