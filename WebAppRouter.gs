/**
 * Webアプリの入り口となる doGet 関数です。
 * HTML テンプレートにアクション定義を渡して描画します。
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  Logger.log('doGet params: %s', JSON.stringify(params));

  try {
    if ((params.view || '') === 'media-register') {
      return renderMediaRegisterPage();
    }
    if ((params.view || '') === 'promotion-apply') {
      return renderPromotionApplyPage();
    }
    return renderPortalPage(params);
  } catch (err) {
    Logger.log('doGet error: %s', (err && err.stack) || err);
    return HtmlService.createHtmlOutput(
        '<h1>サーバー側エラー</h1><pre>' +
        (err && (err.stack || err.toString())) +
        '</pre>'
      )
      .setTitle('OTONARI API エラー');
  }
}

function renderPortalPage(params) {
  params = params || {};
  const template = HtmlService.createTemplateFromFile('MainSite');
  template.actionsJson = getWebActionDefinitions();
  template.logoUrl = getLogoUrlFromSheet();
  template.selectedActionId = params.action || '';
  template.selectedView = params.view || '';
  template.baseUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('OTONARI API ポータル')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function renderMediaRegisterPage() {
  const template = HtmlService.createTemplateFromFile('MediaRegister');
  template.portalUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('メディア登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function renderPromotionApplyPage() {
  const template = HtmlService.createTemplateFromFile('PromotionApply');
  template.portalUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('提携申請登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Web アクションの最小限の情報を抽出して返します。
 * 直接 HTML で扱いやすい構造にすることで保守性を高めています。
 */
function getWebActionDefinitions() {
  const configList = getWebActionConfigList();
  if (!Array.isArray(configList)) {
    console.warn('Webアクション設定の取得結果が配列ではありません。安全のため空配列を返します。');
    return [];
  }

  return configList
    .filter(function(action) {
      return action && typeof action === 'object';
    })
    .map(function(action) {
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

  const capturedLogs = [];
  const restoreLogging = captureExecutionLogs(capturedLogs);

  try {
    // 引数が無い場合は apply を通さず直接実行し、エラーハンドリングを簡潔に保ちます。
    const result = args.length ? handler.apply(null, args) : handler();
    const embeddedLogs = result && Array.isArray(result.logs) ? result.logs : [];
    return { result: result, logs: capturedLogs.concat(embeddedLogs) };
  } finally {
    restoreLogging();
  }
}

/**
 * Logger.log / console.* に出力される内容を配列に蓄積し、呼び出し元へ返却できるようにします。
 * 実行後は必ず restore 関数で元の状態に戻してください。
 */
function captureExecutionLogs(buffer) {
  const target = Array.isArray(buffer) ? buffer : [];
  const originalLoggerLog = Logger.log;
  const originalConsole = {
    log: console.log || function() {},
    info: console.info || console.log || function() {},
    warn: console.warn || console.log || function() {},
    error: console.error || console.log || function() {},
  };

  const append = function(level, args) {
    try {
      const message = Array.prototype.slice.call(args).map(function(part) {
        if (part === null || part === undefined) {
          return '';
        }
        if (typeof part === 'string') {
          return part;
        }
        try {
          return JSON.stringify(part);
        } catch (e) {
          return String(part);
        }
      }).join(' ');
      if (message) {
        target.push({ level: level || 'info', message: message });
      }
    } catch (ignored) {
      // ログ出力の失敗は処理継続を優先し、握りつぶします。
    }
  };

  Logger.log = function() {
    append('info', arguments);
    return originalLoggerLog.apply(Logger, arguments);
  };

  console.log = function() {
    append('info', arguments);
    return originalConsole.log.apply(console, arguments);
  };
  console.info = function() {
    append('info', arguments);
    return originalConsole.info.apply(console, arguments);
  };
  console.warn = function() {
    append('warning', arguments);
    return originalConsole.warn.apply(console, arguments);
  };
  console.error = function() {
    append('error', arguments);
    return originalConsole.error.apply(console, arguments);
  };

  return function restoreLogging() {
    Logger.log = originalLoggerLog;
    console.log = originalConsole.log;
    console.info = originalConsole.info;
    console.warn = originalConsole.warn;
    console.error = originalConsole.error;
  };
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
      const blobDataUrl = convertBlobToDataUrl(value);
      if (blobDataUrl) {
        return blobDataUrl;
      }

      if (typeof value.getSourceUrl === 'function') {
        const sourceUrl = value.getSourceUrl();
        if (sourceUrl) {
          return sourceUrl;
        }
      }
    }

    // セル上に配置された画像（オーバーセル画像）を探索
    const images = sheet.getImages && sheet.getImages();
    if (images && typeof images.forEach === 'function') {
      for (var i = 0; i < images.length; i++) {
        var image = images[i];
        if (!image) {
          continue;
        }

        var anchorCell = typeof image.getAnchorCell === 'function'
          ? image.getAnchorCell()
          : null;

        if (!anchorCell) {
          var anchorRow = typeof image.getAnchorRow === 'function' ? image.getAnchorRow() : null;
          var anchorColumn = typeof image.getAnchorColumn === 'function' ? image.getAnchorColumn() : null;
          if (anchorRow && anchorColumn) {
            anchorCell = sheet.getRange(anchorRow, anchorColumn);
          }
        }

        if (anchorCell && anchorCell.getA1Notation && anchorCell.getA1Notation() === TARGET_RANGE) {
          var overImageDataUrl = convertBlobToDataUrl(image);
          if (overImageDataUrl) {
            return overImageDataUrl;
          }
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

/**
 * Blob を持つ可能性があるオブジェクトから data URL を生成します。
 *
 * @param {Object} blobHolder
 * @return {string}
 */
function convertBlobToDataUrl(blobHolder) {
  if (!blobHolder || typeof blobHolder.getBlob !== 'function') {
    return '';
  }

  var blob = blobHolder.getBlob();
  if (!blob) {
    return '';
  }

  var contentType = blob.getContentType() || 'image/png';
  var base64 = Utilities.base64Encode(blob.getBytes());
  return 'data:' + contentType + ';base64,' + base64;
}
