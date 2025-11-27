# ACS API 利用メモ

## 広告枠 取得（複数） `/media_space/search`

- 初期ソート: 登録日時の降順。
- 主なクエリパラメーター:
  - `id` (string, <= 32)
  - `user` (string, <= 32) — アフィリエイター
  - `media` (string, <= 32) — メディア
  - `name` (string, <= 255) — 広告枠名
  - `tag` (string) — 配信タグ
  - `opens` (int, 0/1) — 公開ステータス
  - `parent_use_state` (int, 0/1) — 利用ステータス
  - `edit_unix` (int32) — 最終編集日時
  - `regist_unix` (int32) — 登録日時
- ヘッダー: `X-Auth-Token: {accessKey}:{secretKey}`
- 実装箇所: `registerMedia.gs` の `listMediaSpaces` で、全ページを走査して広告枠を取得しています。【F:registerMedia.gs†L505-L522】

## 広告・メディアに関する既存取得処理

- 広告（プロモーション）と広告主のマスタ更新: `updateMasterFromAPI.gs` の `updateMasterFromAPI`
  が `/advertiser/search` と `/promotion/search` を全件取得して「マスタ」シートを更新します。【F:updateMasterFromAPI.gs†L1-L62】
- メディア一覧取得: `registerMedia.gs` の `listActiveMediaByAffiliate` で `/media/search` を呼び出し、
  指定アフィリエイターの有効なメディアをフィルタリングしています。【F:registerMedia.gs†L479-L502】
