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

## 提携申請 API `/promotion_apply/*`

### 共通の呼び出し方

- HTTP ヘッダー: `X-Auth-Token: {accessKey}:{secretKey}`（管理者アカウントの API キーを連結）。
- `POST` / `PUT` は `Content-Type: application/json` を必ず付与し、Body に JSON を送ります。
- クエリパラメーターを使う `GET` は URL エンコードした文字列を付与します。
- 広告 ID・メディア ID はローカルのマスタシートではなく、毎回 `/promotion/search`・`/media/search` などの API 応答から取得した値を使います。

### 登録 `POST /promotion_apply/regist`

- Body (application/json)
  - `media` (string, required, <= 32) — メディア ID
  - `promotion` (string, required, <= 32) — 広告 ID
  - `state` (int, optional, 0 or 1) — 承認状態（0: 未承認, 1: 承認）。省略時は承認(1)で送る。
- cURL 例（登録）:
  ```bash
  curl -X POST "https://<acs-host>/promotion_apply/regist" \
    -H "X-Auth-Token: ${ACCESS_KEY}:${SECRET_KEY}" \
    -H "Content-Type: application/json" \
    -d '{"media":"<media-id>","promotion":"<promotion-id>","state":1}'
  ```
- 主な利用箇所: `affiliateAutomation.gs` の `ensurePromotionApplication` が重複チェック後に登録し、`applyAffiliatePartnerships.gs` の `registerPromotionApplication` も一括申請で使用します。【F:affiliateAutomation.gs†L330-L384】【F:applyAffiliatePartnerships.gs†L409-L477】

### 複数取得 `GET /promotion_apply/search`

- Query Parameters
  - `id`, `user`, `media`, `advertiser`, `promotion` (string, <= 32)
  - `state` (int, 0/1/2/3) — 0: 未承認, 1: 承認, 2: 保留, 3: 却下
- cURL 例（検索）:
  ```bash
  curl -G "https://<acs-host>/promotion_apply/search" \
    -H "X-Auth-Token: ${ACCESS_KEY}:${SECRET_KEY}" \
    --data-urlencode "media=<media-id>" \
    --data-urlencode "promotion=<promotion-id>"
  ```
- 主な利用箇所: `affiliateAutomation.gs` の `findExistingPromotionApplication` で重複申請を判定し、`registerMedia.gs` の `findPromotionApplications` でも送信済み判定に用いています。【F:affiliateAutomation.gs†L386-L418】【F:registerMedia.gs†L602-L627】

### 単一取得 `GET /promotion_apply/info`

- Query Parameter: `id` (uuid, required) — 提携申請 ID
- cURL 例（1件取得）:
  ```bash
  curl -G "https://<acs-host>/promotion_apply/info" \
    -H "X-Auth-Token: ${ACCESS_KEY}:${SECRET_KEY}" \
    --data-urlencode "id=<apply-id>"
  ```
- 返却例: `{ "record": { "id": "string", "user": "string", "media": "string", "advertiser": "string", "promotion": "string", "state": 0 } }`

### 編集 `PUT /promotion_apply/edit`

- Body (application/json)
  - `state` (int, 0/1/2/3) — 承認状態（0: 未承認, 1: 承認, 2: 保留, 3: 却下）
- cURL 例（更新）:
  ```bash
  curl -X PUT "https://<acs-host>/promotion_apply/edit" \
    -H "X-Auth-Token: ${ACCESS_KEY}:${SECRET_KEY}" \
    -H "Content-Type: application/json" \
    -d '{"state":1}'
  ```
- ステータス変更が必要な場合に使用します。
