function updateMasterFromAPI() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    Logger.log('[updateMasterFromAPI] ❌ Active spreadsheet not found.');
    return;
  }

  const sheet = spreadsheet.getSheetByName('マスタ') || spreadsheet.insertSheet('マスタ');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 8).setValues([
    ['表示名', '会社名', '氏名', '広告主ID', '広告名', '広告ID', 'グロス単価', 'ネット単価']
  ]);

  const accessKey = 'agqnoournapf';
  const secretKey = '1kvu9dyv1alckgocc848socw';
  const token = accessKey + ':' + secretKey;

  const advertiserUrl = 'https://otonari-asp.com/api/v1/m/advertiser/search';
  const promotionUrl = 'https://otonari-asp.com/api/v1/m/promotion/search';

  try {
    Logger.log('[updateMasterFromAPI] Fetching advertisers...');
    const advertiserList = callAllPagesAPI(advertiserUrl, token, 'advertiser');
    Logger.log('[updateMasterFromAPI] advertiser count=' + advertiserList.length);

    Logger.log('[updateMasterFromAPI] Fetching promotions...');
    const promotionList = callAllPagesAPI(promotionUrl, token, 'promotion');
    Logger.log('[updateMasterFromAPI] promotion count=' + promotionList.length);

    const advertiserMap = {};
    advertiserList.forEach(function (ad) {
      if (!ad || !ad.id) {
        Logger.log('[updateMasterFromAPI] Skipping advertiser without id: ' + JSON.stringify(ad));
        return;
      }
      advertiserMap[ad.id] = {
        company: ad.company || '',
        name: ad.name || ''
      };
    });

    const output = [];
    var missingAdvertiserCount = 0;

    promotionList.forEach(function (promo) {
      if (!promo) {
        Logger.log('[updateMasterFromAPI] Skipping undefined promotion record');
        return;
      }
      const advId = promo.advertiser;
      if (!advId) {
        Logger.log('[updateMasterFromAPI] Promotion without advertiser id: ' + JSON.stringify(promo));
        return;
      }

      const adv = advertiserMap[advId];
      if (!adv) {
        missingAdvertiserCount++;
        Logger.log('[updateMasterFromAPI] Advertiser not found for promotion. advertiserId=' + advId + ' promotionId=' + promo.id);
        return;
      }

      const displayName = (adv.company + ' ' + adv.name).trim();
      const promoName = promo.name || '';
      const gross = promo.gross_action_cost !== undefined ? promo.gross_action_cost : '';
      const net = promo.net_action_cost !== undefined ? promo.net_action_cost : '';

      output.push([
        displayName,
        adv.company,
        adv.name,
        advId,
        promoName,
        promo.id || '',
        gross,
        net
      ]);
    });

    if (missingAdvertiserCount) {
      Logger.log('[updateMasterFromAPI] Promotions skipped due to missing advertiser: ' + missingAdvertiserCount);
    }

    if (output.length > 0) {
      sheet.getRange(2, 1, output.length, 8).setValues(output);
      Logger.log('[updateMasterFromAPI] ✅ 更新完了 rows=' + output.length);
    } else {
      Logger.log('[updateMasterFromAPI] ⚠️ 出力対象がありませんでした。');
    }
  } catch (error) {
    Logger.log('[updateMasterFromAPI] ❌ エラー: ' + error + '\n' + (error && error.stack ? error.stack : ''));
  }
}

function callAllPagesAPI(baseUrl, token, label) {
  const allRecords = [];
  let offset = 0;
  const limit = 100;

  while (true) {
    const url = baseUrl + '?offset=' + offset + '&limit=' + limit;
    Logger.log('[callAllPagesAPI] Request start label=' + (label || '') + ' offset=' + offset + ' limit=' + limit + ' url=' + url);
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'X-Auth-Token': token },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      const content = response.getContentText();

      if (code !== 200) {
        Logger.log('[callAllPagesAPI] ❌ HTTP ' + code + ' label=' + (label || '') + ' body=' + content);
        break;
      }

      let data;
      try {
        data = JSON.parse(content);
      } catch (parseError) {
        Logger.log('[callAllPagesAPI] ❌ JSON parse error label=' + (label || '') + ' error=' + parseError + ' body=' + content);
        break;
      }

      if (!data || !data.records) {
        Logger.log('[callAllPagesAPI] ⚠️ No records field label=' + (label || '') + ' response=' + content);
        break;
      }

      Logger.log('[callAllPagesAPI] Retrieved ' + data.records.length + ' records label=' + (label || '') + ' offset=' + offset);
      allRecords.push.apply(allRecords, data.records);

      if (data.records.length < limit) {
        Logger.log('[callAllPagesAPI] Completed label=' + (label || '') + ' total=' + allRecords.length);
        break;
      }

      offset += limit;
    } catch (fetchError) {
      Logger.log('[callAllPagesAPI] ❌ Fetch error label=' + (label || '') + ' error=' + fetchError + '\n' + (fetchError && fetchError.stack ? fetchError.stack : ''));
      break;
    }
  }

  return allRecords;
}
