var VENDING_SURVEY_SPREADSHEET_ID = '1xkg8vNscpcWTA6GA0VPxGTJCAH6LyvsYhq7VhOlDcXg';
var VENDING_SUMMARY_SHEET_NAME = '集計';
var VENDING_INFO_SHEET_NAME = '情報';
var VENDING_IMAGE_FOLDER_ID = '1AJd4BTFTVrLNep44PDz1AuwSF_5TFxdx';
var VENDING_LINEUP_FOLDER_ID = '18fA4HRavIBTM2aPL-OqVaWhjRRgBhlKg';

function getVendingSurveyOptions() {
  var infoSheet = getVendingSheet_(VENDING_INFO_SHEET_NAME);
  var values = infoSheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return { options: [] };
  }

  var headers = values.shift();
  var headerIndex = createHeaderIndex_(headers);

  return {
    options: values
      .filter(function(row) {
        return row && row.length;
      })
      .map(function(row) {
        var name = getValueByHeader_(row, headerIndex, 'メーカー');
        var featureText = getValueByHeader_(row, headerIndex, '特徴');
        var availability = (getValueByHeader_(row, headerIndex, '可否') || '').trim();
        var imageId = getValueByHeader_(row, headerIndex, '画像ID') || '';
        var imageFileName = getValueByHeader_(row, headerIndex, '画像ファイル名') || '';
        var imageUrl = resolveImageUrl_(name, imageId, imageFileName);
        var lineupFolderId = getValueByHeader_(row, headerIndex, 'ラインアップフォルダID') || '';

        return {
          name: name,
          description: featureText,
          features: splitFeatures_(featureText),
          availability: availability,
          availabilityLevel: getAvailabilityLevel_(availability),
          imageUrl: imageUrl,
          lineupUrl:
            ScriptApp.getService().getUrl() +
            '?view=vending-lineup&maker=' +
            encodeURIComponent(name) +
            (lineupFolderId ? '&folderId=' + encodeURIComponent(lineupFolderId) : ''),
        };
      })
      .filter(function(option) {
        return option.name;
      }),
  };
}

function submitVendingSurveyResponse(selectedMaker, requestText) {
  var maker = (selectedMaker || '').trim();
  var request = (requestText || '').trim();
  if (!maker) {
    throw new Error('自販機メーカーを選択してください。');
  }

  var sheet = getVendingSheet_(VENDING_SUMMARY_SHEET_NAME);
  var email = '';
  try {
    var activeUser = Session.getActiveUser();
    email = activeUser ? activeUser.getEmail() : '';
  } catch (e) {
    email = '';
  }

  var rowValues = [new Date(), email, maker, request];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowValues.length).setValues([rowValues]);

  return { message: '送信が完了しました。ご協力ありがとうございます。', email: email };
}

function getVendingLineup(manufacturer, optionalFolderId) {
  var maker = (manufacturer || '').trim();
  var lineupFolderId = (optionalFolderId || '').trim();
  var baseFolder = DriveApp.getFolderById(lineupFolderId || VENDING_LINEUP_FOLDER_ID);

  var makerFolder = findChildFolderByName_(baseFolder, maker) || baseFolder;
  var categories = [];
  var categoryIterator = makerFolder.getFolders();
  while (categoryIterator.hasNext()) {
    var categoryFolder = categoryIterator.next();
    categories.push(buildProductCategory_(categoryFolder.getName(), categoryFolder));
  }

  if (!categories.length) {
    categories.push(buildProductCategory_('ラインアップ', makerFolder));
  }

  return {
    manufacturer: maker || 'ラインアップ',
    categories: categories,
  };
}

function buildProductCategory_(categoryName, folder) {
  var items = [];
  var fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    if (file.isTrashed()) continue;
    items.push({
      name: file.getName(),
      price: extractPriceFromName_(file.getName()),
      imageUrl: convertToImageUrl_(file.getId()),
    });
  }

  return { name: categoryName, items: items };
}

function extractPriceFromName_(fileName) {
  if (!fileName) return '';
  var match = fileName.match(/([0-9]{2,}(?:[,\.]?[0-9]+)?)/);
  if (match) {
    return match[1].replace(',', '');
  }
  return '';
}

function convertToImageUrl_(fileId) {
  return fileId ? 'https://lh3.googleusercontent.com/d/' + fileId : '';
}

function resolveImageUrl_(makerName, explicitId, fileName) {
  if (explicitId) {
    return convertToImageUrl_(explicitId);
  }

  var folder = DriveApp.getFolderById(VENDING_IMAGE_FOLDER_ID);
  var candidates = [];
  var fileIterator = folder.getFiles();
  var lowerName = (makerName || '').toLowerCase();
  var lowerFileName = (fileName || '').toLowerCase();

  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    if (file.isTrashed()) continue;
    var currentName = file.getName();
    var lowerCurrentName = currentName.toLowerCase();
    if (
      (lowerName && lowerCurrentName.indexOf(lowerName) !== -1) ||
      (lowerFileName && lowerCurrentName.indexOf(lowerFileName) !== -1)
    ) {
      candidates.push(convertToImageUrl_(file.getId()));
    }
  }

  if (candidates.length) {
    return candidates[0];
  }

  // Fallback: first file in the folder
  var fallbackIterator = folder.getFiles();
  return fallbackIterator.hasNext() ? convertToImageUrl_(fallbackIterator.next().getId()) : '';
}

function splitFeatures_(text) {
  var raw = (text || '').split(/\n|、|\/|・|\r/).map(function(part) {
    return (part || '').trim();
  });
  return raw.filter(function(item) {
    return item;
  });
}

function getAvailabilityLevel_(availabilityText) {
  var normalized = (availabilityText || '').replace(/\s+/g, '').toLowerCase();
  if (!normalized) return 'unknown';
  if (normalized.indexOf('不可') !== -1 || normalized.indexOf('決定') !== -1) {
    return 'unavailable';
  }
  if (normalized.indexOf('可能') !== -1) {
    return 'available';
  }
  return 'unknown';
}

function createHeaderIndex_(headers) {
  var index = {};
  (headers || []).forEach(function(header, idx) {
    if (!header) return;
    index[String(header).trim()] = idx;
  });
  return index;
}

function getValueByHeader_(row, index, headerName) {
  var pos = index[headerName];
  return pos === 0 || pos ? row[pos] : '';
}

function getVendingSheet_(name) {
  var spreadsheet = SpreadsheetApp.openById(VENDING_SURVEY_SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  return sheet;
}

function findChildFolderByName_(parent, keyword) {
  if (!keyword) return null;
  var target = keyword.toLowerCase();
  var iterator = parent.getFolders();
  while (iterator.hasNext()) {
    var folder = iterator.next();
    if (folder.getName().toLowerCase().indexOf(target) !== -1) {
      return folder;
    }
  }
  return null;
}
