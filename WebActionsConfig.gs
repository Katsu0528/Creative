/**
 * Webアプリ上で実行可能な処理の定義一覧です。
 * 新しい処理を追加する際は、ここに項目を追加してください。
 */
const WEB_ACTION_DEFINITIONS = [
  {
    id: 'adcore',
    group: '広告・提携ワークフロー',
    name: '広告登録フローを実行',
    description: '入稿用シートの設定に従って広告登録／素材追加／提携申請をまとめて実行します。',
    handler: 'AdCore',
    fields: []
  },
  {
    id: 'registerPromotionItem',
    group: '広告・提携ワークフロー',
    name: '商品テーブル登録',
    description: '広告ID・商品名・URLを指定して商品テーブルを作成します。',
    handler: 'registerPromotionItem',
    fields: [
      {
        id: 'promotionId',
        label: '広告ID',
        type: 'text',
        placeholder: '例: 123456',
        optional: false
      },
      {
        id: 'itemName',
        label: '商品名',
        type: 'text',
        placeholder: '例: サンプル商品A',
        optional: false
      },
      {
        id: 'itemUrl',
        label: '商品URL',
        type: 'url',
        placeholder: 'https://example.com',
        optional: false
      }
    ]
  },
  {
    id: 'registerMedia',
    group: '広告・提携ワークフロー',
    name: 'メディア登録',
    description: '複数のメディア情報を貼り付けて一括登録します。',
    handler: 'registerMediaFromWeb',
    fields: []
  },
  {
    id: 'promotionApplySheet',
    group: '広告・提携ワークフロー',
    name: '提携申請登録',
    description: '広告名とメディア名を貼り付けて提携申請をまとめて送信します。',
    handler: 'registerPromotionApplicationsFromWeb',
    fields: []
  },
  {
    id: 'summarizeAdsFromFolder',
    group: '集計・レポート',
    name: 'フォルダ内広告レポート集計',
    description: 'Driveフォルダ内のファイルを集計し、マスターシートに取り込みます。',
    handler: 'summarizeAdsFromFolder',
    fields: []
  },
  {
    id: 'summarizeConfirmedResultsByAffiliate',
    group: '集計・レポート',
    name: '確定成果レポート作成',
    description: '対象期間の確定成果を取得し「受領」シートへ集計します。',
    handler: 'summarizeConfirmedResultsByAffiliate',
    fields: []
  },
  {
    id: 'summarizeAgencyAds',
    group: '集計・レポート',
    name: '代理店向けレポート作成',
    description: '代理店向けの成果レポートを作成し、必要に応じてシート名で出力先を指定します。',
    handler: 'summarizeAgencyAds',
    fields: [
      {
        id: 'targetSheetName',
        label: '出力先シート名（任意）',
        type: 'text',
        placeholder: '未入力の場合はデフォルトのシートを使用',
        optional: true
      }
    ]
  },
  {
    id: 'parseMultiFormatData',
    group: '集計・レポート',
    name: '多様な明細データの解析',
    description: '貼り付けたテキストデータから必要な列を抽出し「抽出結果」シートに整形します。',
    handler: 'parseMultiFormatData',
    fields: []
  },
  {
    id: 'copyNextMonthSheets',
    group: '集計・レポート',
    name: '翌月シート作成',
    description: 'テンプレートを元に翌月分のシートを作成します。',
    handler: 'copyNextMonthSheets',
    fields: []
  },
  {
    id: 'createNextMonthAndSummarize',
    group: '集計・レポート',
    name: '翌月シート作成＋集計',
    description: '翌月分のシートを作成した後、集計処理まで一度に実行します。',
    handler: 'createNextMonthAndSummarize',
    fields: []
  },
  {
    id: 'processSeikaChanges',
    group: '集計・レポート',
    name: '成果変更データの処理',
    description: '「成果変更用」シートの内容を元にAPIから対象レコードを取得・整形します。',
    handler: 'processSeikaChanges',
    fields: []
  },
  {
    id: 'updateAdvertiserIds',
    group: 'マスタ・メンテナンス',
    name: '広告主IDの自動補完',
    description: '「クライアント情報」シートの広告主名からIDを検索し、ID列を更新します。',
    handler: 'updateAdvertiserIds',
    fields: []
  },
  {
    id: 'updateMasterFromAPI',
    group: 'マスタ・メンテナンス',
    name: 'マスタ情報更新',
    description: 'APIから最新のマスタ情報を取得してシートへ反映します。',
    handler: 'updateMasterFromAPI',
    fields: []
  },
  {
    id: 'cleanupSheets',
    group: 'マスタ・メンテナンス',
    name: '一時シートの削除',
    description: 'スプレッドシート内の「(SJIS)action_log_raw_」で始まるシートを削除します。',
    handler: 'cleanupSheets',
    fields: []
  },
  {
    id: 'downloadCsvDlShiftJis',
    group: 'マスタ・メンテナンス',
    name: 'CSVダウンロード（Shift-JIS）',
    description: 'ダウンロードしたCSVをShift-JISに変換し、ユーザーへ返却します。',
    handler: 'downloadCsvDlShiftJis',
    fields: []
  },
  {
    id: 'generateChatMessages',
    group: '通知・コミュニケーション',
    name: 'Google Chat通知文生成',
    description: '案件情報を元にGoogle Chatへ投稿するメッセージ文を生成します。',
    handler: 'generateChatMessages',
    fields: []
  }
];

/**
 * Web アクション定義の配列を返します。
 * Apps Script のファイル読み込み順に依存しないよう、関数経由で参照します。
 */
function getWebActionConfigList() {
  return WEB_ACTION_DEFINITIONS;
}

// 既存コードとの後方互換性のためにグローバル変数へも代入しておく
if (typeof WEB_ACTIONS === 'undefined') {
  globalThis.WEB_ACTIONS = WEB_ACTION_DEFINITIONS;
}
