/**
 * Webアプリの入り口となる doGet 関数です。
 * HTML テンプレートにアクション定義を渡して描画します。
 */
function doGet() {
  // 定義済みアクションをクライアント側へ受け渡す
  const template = HtmlService.createTemplateFromFile('MainSite');
  template.actionsJson = JSON.stringify(getWebActionDefinitions());
  return template
    .evaluate()
    .setTitle('Creative Operations Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Web アクションの最小限の情報を抽出して返します。
 * 直接 HTML で扱いやすい構造にすることで保守性を高めています。
 */
function getWebActionDefinitions() {
  return WEB_ACTIONS.map(function(action) {
    return {
      id: action.id,
      group: action.group,
      name: action.name,
      description: action.description,
      handler: action.handler,
      fields: action.fields || []
    };
  });
}

/**
 * クライアントからリクエストされたアクションを実際の処理にディスパッチします。
 * actionId に紐づく handler を安全に呼び出し、結果をそのまま返却します。
 */
function runWebAction(actionId, formValues) {
  if (!actionId) {
    throw new Error('アクションIDが指定されていません。');
  }

  const action = WEB_ACTIONS.find(function(item) {
    return item.id === actionId;
  });
  if (!action) {
    throw new Error('指定されたアクションが見つかりません: ' + actionId);
  }

  const handlerName = action.handler;
  const handler = globalThis[handlerName];
  if (typeof handler !== 'function') {
    throw new Error('実行対象の関数が定義されていません: ' + handlerName);
  }

  const fields = action.fields || [];
  const values = formValues || {};
  const args = fields.map(function(field) {
    const value = values[field.id];
    if (!field.optional && (!value && value !== 0)) {
      throw new Error('必須項目が未入力です: ' + field.label);
    }
    // 未入力の場合は null を渡すことで、既存処理の分岐を書き換えなくても良いようにしています。
    return value === '' ? null : value;
  });

  // 引数が無い場合は apply を通さず直接実行し、エラーハンドリングを簡潔に保ちます。
  return args.length ? handler.apply(null, args) : handler();
}

/**
 * HTML ファイルをインクルードするためのヘルパーです。
 * 必要に応じてテンプレート内から呼び出してください。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
