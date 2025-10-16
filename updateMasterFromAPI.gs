function updateMasterFromAPI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("マスタ") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("マスタ");
  sheet.clearContents();
  // G列:グロス単価、H列:ネット単価をヘッダーに追加
  sheet.getRange(1, 1, 1, 8).setValues([
    ["表示名", "会社名", "氏名", "広告主ID", "広告名", "広告ID", "グロス単価", "ネット単価"]
  ]);

  const apiConfig = getApiConfig();

  const advertiserUrl = `${apiConfig.baseUrl}/advertiser/search`;
  const promotionUrl = `${apiConfig.baseUrl}/promotion/search`;

  try {
    const advertiserList = callAllPagesAPI(advertiserUrl, apiConfig.headers);
    const promotionList = callAllPagesAPI(promotionUrl, apiConfig.headers);

    // 広告主情報をIDでマッピング
    const advertiserMap = {};
    advertiserList.forEach(ad => {
      advertiserMap[ad.id] = {
        company: ad.company || "",
        name: ad.name || ""
      };
    });

    const output = [];

    promotionList.forEach(promo => {
      const advId = promo.advertiser;
      const promoName = promo.name || "";
      const promoId = promo.id;

      // グロス単価・ネット単価（なければ空文字）
      const gross = promo.gross_action_cost !== undefined ? promo.gross_action_cost : "";
      const net = promo.net_action_cost !== undefined ? promo.net_action_cost : "";

      const adv = advertiserMap[advId];
      if (!adv) return;

      const company = adv.company;
      const personal = adv.name;
      const displayName = `${company} ${personal}`.trim();

      // G列:グロス単価、H列:ネット単価を追加
      output.push([displayName, company, personal, advId, promoName, promoId, gross, net]);
    });

    if (output.length > 0) {
      sheet.getRange(2, 1, output.length, 8).setValues(output);
    }

    Logger.log(`✅ 完了：${output.length} 件を更新しました`);
  } catch (e) {
    Logger.log("❌ エラー: " + e);
  }
}

function callAllPagesAPI(baseUrl, headers) {
  const allRecords = [];
  let offset = 0;
  const limit = 100;

  while (true) {
    const url = `${baseUrl}?offset=${offset}&limit=${limit}`;
    const options = {
      method: 'get',
      headers,
      muteHttpExceptions: true
    };

    let response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (error) {
      Logger.log(`❌ API通信エラー: ${error} (offset: ${offset})`);
      break;
    }

    const code = response.getResponseCode();
    const body = response.getContentText();
    if (code !== 200) {
      Logger.log(`❌ APIエラー: ${code} at offset ${offset} body: ${body}`);
      if (Number(code) === 401) {
        throw new Error('API認証に失敗しました。アクセスキーとシークレットキーを確認してください。');
      }
      throw new Error(`APIリクエストに失敗しました (status: ${code}, offset: ${offset}).`);
    }

    let records;
    try {
      const data = JSON.parse(body);
      records = normalizeRecords(data.records);
    } catch (error) {
      Logger.log(`❌ JSON解析エラー: ${error} body: ${body}`);
      break;
    }

    if (!records.length) {
      break;
    }

    allRecords.push(...records);
    if (records.length < limit) {
      break;
    }

    offset += records.length;
  }

  return allRecords;
}

function getApiConfig() {
  const props = PropertiesService.getScriptProperties();
  let baseUrl =
    getCleanProperty(props, ['OTONARI_BASE_URL', 'https://otonari-asp.com/api/v1/m']) ||
    'https://otonari-asp.com/api/v1/m';
  baseUrl = baseUrl.replace(/\/+$/, '');

  const accessKey =
    getCleanProperty(props, ['OTONARI_ACCESS_KEY', 'agqnoournapf']) || 'agqnoournapf';
  const secretKey =
    getCleanProperty(props, ['OTONARI_SECRET_KEY', '1kvu9dyv1alckgocc848socw']) ||
    '1kvu9dyv1alckgocc848socw';

  if (!accessKey || !secretKey) {
    throw new Error('APIのアクセスキーまたはシークレットキーが設定されていません。');
  }

  return {
    baseUrl,
    headers: { 'X-Auth-Token': `${accessKey}:${secretKey}` }
  };
}

function normalizeRecords(records) {
  if (!records) {
    return [];
  }

  if (Array.isArray(records)) {
    return records;
  }

  if (records.id) {
    return [records];
  }

  const normalized = [];
  for (const key in records) {
    if (Object.prototype.hasOwnProperty.call(records, key) && records[key]) {
      normalized.push(records[key]);
    }
  }
  return normalized;
}

function getCleanProperty(props, keys) {
  for (const key of keys) {
    const value = props.getProperty(key);
    if (value && value.trim()) {
      return value.trim();
    }
  }
  return null;
}
