/**
 * Webアプリの入り口となる doGet 関数です。
 * HTML テンプレートにアクション定義を渡して描画します。
 */
function doGet() {
  // 定義済みアクションをクライアント側へ受け渡す
  const template = HtmlService.createTemplateFromFile('MainSite');
  template.actionsJson = JSON.stringify(getWebActionDefinitions());
  template.logoUrl = getLogoUrlFromSheet();
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
  return getWebActionConfigList().map(function(action) {
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

  const action = getWebActionConfigList().find(function(item) {
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

/**
 * シートに格納されたロゴ画像を取得して data URL もしくは公開 URL として返します。
 * 取得に失敗した場合は空文字を返し、フロント側でフォールバック表示を行います。
 *
 * @return {string}
 */
function getLogoUrlFromSheet() {
  const SPREADSHEET_ID = '1f22F3tSeK3PNndceAVmEeQPlDx48O4BCAid1HroJsuw';
  const SHEET_NAME = 'シート1';
  const TARGET_RANGE = 'A1';

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }

    const range = sheet.getRange(TARGET_RANGE);
    const value = range.getValue();

    // 新しい CellImage API で画像が格納されている場合
    if (value && typeof value === 'object') {
      if (typeof value.getBlob === 'function') {
        const blob = value.getBlob();
        if (blob) {
          const contentType = blob.getContentType() || 'image/png';
          const base64 = Utilities.base64Encode(blob.getBytes());
          return 'data:' + contentType + ';base64,' + base64;
        }
      }

      if (typeof value.getSourceUrl === 'function') {
        const sourceUrl = value.getSourceUrl();
        if (sourceUrl) {
          return sourceUrl;
        }
      }
    }

    // =IMAGE("URL") 形式のセルから URL を抽出
    const formulaUrl = extractImageUrlFromFormula(range.getFormula());
    if (formulaUrl) {
      return formulaUrl;
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (trimmed && /^https?:\/\//i.test(trimmed)) {
        return trimmed;
      }
    }
  } catch (error) {
    console.error('ロゴ画像の取得に失敗しました: ' + error);
  }

  return '';
}

/**
 * =IMAGE 関数の数式から画像 URL を取り出します。
 *
 * @param {string} formula
 * @return {string}
 */
function extractImageUrlFromFormula(formula) {
  if (!formula) {
    return '';
  }

  const doubleQuoteMatch = formula.match(/=IMAGE\(\s*"([^"]+)"/i);
  if (doubleQuoteMatch && doubleQuoteMatch[1]) {
    return doubleQuoteMatch[1];
  }

  const singleQuoteMatch = formula.match(/=IMAGE\(\s*'([^']+)'/i);
  if (singleQuoteMatch && singleQuoteMatch[1]) {
    return singleQuoteMatch[1];
  }

  return '';
}
