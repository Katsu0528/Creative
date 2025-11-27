const BASE_API_URL = 'https://otonari-asp.com/api/v1/m';
const DEFAULT_ACCESS_KEY = 'agqnoournapf';
const DEFAULT_SECRET_KEY = '1kvu9dyv1alckgocc848socw';

function getApiConfigForPromotionRegistration() {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = (props.getProperty('OTONARI_BASE_URL') || BASE_API_URL).replace(/\/+$/, '');
  const accessKey = props.getProperty('OTONARI_ACCESS_KEY') || DEFAULT_ACCESS_KEY;
  const secretKey = props.getProperty('OTONARI_SECRET_KEY') || DEFAULT_SECRET_KEY;
  const authToken = `${accessKey}:${secretKey}`;

  return {
    baseUrl: baseUrl,
    headers: {
      'X-Auth-Token': authToken,
    },
  };
}

function buildApiUrl(baseUrl, path, query) {
  const normalizedBase = (baseUrl || '').replace(/\/+$/, '');
  const normalizedPath = String(path || '').replace(/^\/+/, '');
  const searchParams = query && typeof query === 'object' ? new URLSearchParams(query) : null;
  const queryString = searchParams && Array.from(searchParams.keys()).length
    ? `?${searchParams.toString()}`
    : '';
  return `${normalizedBase}/${normalizedPath}${queryString}`;
}

function fetchApiJson(path, query) {
  const config = getApiConfigForPromotionRegistration();
  const url = buildApiUrl(config.baseUrl, path, query);
  const response = UrlFetchApp.fetch(url, { method: 'get', headers: config.headers });
  return JSON.parse(response.getContentText());
}

function searchAdvertiserById(advertiserId) {
  return fetchApiJson('/advertiser/search', { id: advertiserId });
}

function searchPromotionById(promotionId) {
  return fetchApiJson('/promotion/search', { id: promotionId });
}

function registerPromotionItem(promotionId, itemName, itemUrl) {
  const config = getApiConfigForPromotionRegistration();
  const headers = Object.assign({}, config.headers, { 'Content-Type': 'application/json' });
  const payload = {
    promotion: promotionId,
    name: itemName,
    url_type: ['via_system'],
    url: itemUrl,
    display_url: itemUrl,
  };
  const response = UrlFetchApp.fetch(
    buildApiUrl(config.baseUrl, '/promotion_item/regist'),
    { method: 'post', headers: headers, payload: JSON.stringify(payload) }
  );
  return JSON.parse(response.getContentText());
}
