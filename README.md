# Creative スプレッドシート連携アプリの表ガイド

このリポジトリでは、スプレッドシートに貼り付け可能な複数行の入力表を複数のアプリで共有しています。特に **PromotionApplyApp** と **MediaRegisterApp** はほぼ同じ列構成を持ちながら DOM 構造が微妙に異なり、それが貼り付け可否に影響する事例がありました。本書では、今後同種の表を実装するときに踏まえるべき共通ルールをまとめます。

## 既存実装の要点

- **PromotionApplyApp** は各セルの実入力要素（`input` / `select`）自体に `data-column-key` を付与し、貼り付け処理は `event.target.dataset.columnKey` をそのまま開始列として扱います。【F:PromotionApplyApp.html†L976-L1044】
- **MediaRegisterApp** はセルを `.cell-input-control` でラップし、選択チェックボックスも同じセル内に含まれるため、`data-column-key` を持たないラッパーやチェックボックスから貼り付けが始まることがあります。そのため、貼り付け開始列はイベント経路上で最寄りの `data-column-key` を持つ要素をたどって解決し、値の書き込みは `.cell-select-checkbox` を除外した実入力要素を `findColumnInput` で探索して行います。【F:MediaRegisterApp.html†L848-L856】【F:MediaRegisterApp.html†L1388-L1423】

## 表実装のガイドライン

1. **必ず列キーを露出させる**: ユーザーがクリック・フォーカス・貼り付けする要素（通常は `input`/`select`）に `data-column-key` を直接付与します。ラッパーを追加する場合でも、フォーカスを受ける要素かラッパーのどちらかに列キーを持たせてください。PromotionApply 型のシンプルな構造であればこのルールだけで十分です。【F:PromotionApplyApp.html†L976-L1044】
2. **ラップ構造では入力探索ヘルパーを併用する**: ラッパーやチェックボックスが同居する MediaRegister 型の構造では、`findColumnInput` のように「ラッパー内の実入力」を探し、選択用チェックボックス（`.cell-select-checkbox`）を除外するヘルパーを用いて値読み書きを行ってください。`data-column-key` をラッパーにも付けておくと貼り付け開始列の推定が安定します。【F:MediaRegisterApp.html†L848-L856】【F:MediaRegisterApp.html†L1388-L1423】
3. **貼り付けイベントは表全体で捕捉する**: `tableBody.addEventListener('paste', handlePaste, true);` のようにバブリング前に表全体で `paste` を拾い、開始行・開始列を決めたうえでクリップボードの行列データを流し込む既存パターンを再利用してください。開始列が判別できない場合に早期 return すると値が無視されるため、イベント経路上で列キーを見つけるロジックを必ず組み込みます。【F:MediaRegisterApp.html†L1370-L1423】
4. **行データの収集も同じヘルパーで揃える**: バリデーションや送信処理で表データを集約する際は、貼り付け時と同じ入力探索ロジックを使って値を取得します。これにより、DOM 構造の違いにかかわらず未入力検出が一貫します。【F:MediaRegisterApp.html†L1369-L1386】

## 新規表を追加する際のチェックリスト

- [ ] すべての列に一意な `data-column-key` を割り当てたか。
- [ ] フォーカスされる要素（またはその直近のラッパー）に `data-column-key` が付いているか。
- [ ] 選択専用チェックボックスには `data-column-key` を付けず、値入力とは区別しているか。
- [ ] クリップボード貼り付け処理で、イベント経路をたどって開始列を解決しているか。
- [ ] 値の読み書き・バリデーションで、ラップ構造を考慮した入力探索ヘルパーを再利用しているか。

この方針に従うことで、PromotionApplyApp と MediaRegisterApp のように DOM 構造が異なる場合でも、貼り付けや未入力検出の挙動を揃えられます。
