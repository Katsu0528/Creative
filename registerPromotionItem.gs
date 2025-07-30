// Utility functions to interact with advertiser and promotion APIs
// and to register a new promotion item. This demonstrates how to
// combine existing working endpoints.

const BASE_API_URL = 'https://otonari-asp.com/api/v1/m';

function searchAdvertiserById(advId) {
  const props = PropertiesService.getScriptProperties();
  const accessKey = 'agqnoournapf';
  const secretKey = '1kvu9dyv1alckgocc848socw';
  const headers = { 'X-Auth-Token': accessKey + ':' + secretKey };
  const url = `${BASE_API_URL}/advertiser/search?id=${encodeURIComponent(advId)}`;
  const response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
  return JSON.parse(response.getContentText());
}

function searchPromotionById(promoId) {
  const props = PropertiesService.getScriptProperties();
  const accessKey = props.getProperty('agqnoournapf');
  const secretKey = props.getProperty('1kvu9dyv1alckgocc848socw');
  const headers = { 'X-Auth-Token': accessKey + ':' + secretKey };
  const url = `${BASE_API_URL}/promotion/search?id=${encodeURIComponent(promoId)}`;
  const response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
  return JSON.parse(response.getContentText());
}

function registerPromotionItem(promotionId, itemName, itemUrl) {
  const props = PropertiesService.getScriptProperties();
  const accessKey = props.getProperty('agqnoournapf');
  const secretKey = props.getProperty('1kvu9dyv1alckgocc848socw');
  const headers = {
    'X-Auth-Token': accessKey + ':' + secretKey,
    'Content-Type': 'application/json'
  };
  const payload = {
    promotion: promotionId,
    name: itemName,
    url_type: ['via_system'],
    url: itemUrl,
    display_url: itemUrl
  };
  const response = UrlFetchApp.fetch(
    `${BASE_API_URL}/promotion_item/regist`,
    { method: 'post', headers: headers, payload: JSON.stringify(payload) }
  );
  return JSON.parse(response.getContentText());
}
