/**
 * Meta (Facebook/Instagram) Ads データ取得スクリプト
 *
 * Meta Marketing API v21.0 からキャンペーン別日次データを取得し、
 * Google Sheet の meta_ads_data タブに書き込む。
 *
 * 設定方法:
 *   1. Script Properties に以下を設定:
 *      - META_ACCESS_TOKEN: 長期アクセストークン
 *      - META_AD_ACCOUNT_ID: 広告アカウントID（act_なし）
 *   2. トリガーで毎日 fetchMetaAdsData を実行
 */

var SPREADSHEET_ID = "1GbB23Qzf_lhErGiWCcAJz1Yqk_UUloNatWgBpXilGkc";
var SHEET_NAME = "meta_ads_data";
var API_VERSION = "v21.0";
var FX_RATE_USD_JPY = 150;
var START_DATE = "2026-06-01";

function fetchMetaAdsData() {
  var props = PropertiesService.getScriptProperties();
  var accessToken = props.getProperty("META_ACCESS_TOKEN");
  var adAccountId = props.getProperty("META_AD_ACCOUNT_ID");

  if (!accessToken || !adAccountId) {
    Logger.log("ERROR: META_ACCESS_TOKEN or META_AD_ACCOUNT_ID not set in Script Properties");
    return;
  }

  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 10).setValues([
      ["date", "campaign", "impressions", "clicks", "cost", "conversions", "ctr", "cpc", "cvr", "cpa"]
    ]);
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  var yesterday = getYesterday();
  var allRows = [];
  var url = buildApiUrl(adAccountId, accessToken, START_DATE, yesterday);

  while (url) {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());

    if (json.error) {
      Logger.log("API Error: " + JSON.stringify(json.error));
      return;
    }

    var data = json.data || [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var date = row.date_start;
      var campaignName = row.campaign_name || "";
      var impressions = parseInt(row.impressions || 0);
      var clicks = parseInt(row.clicks || 0);
      var spendUsd = parseFloat(row.spend || 0);
      var cost = Math.round(spendUsd * FX_RATE_USD_JPY);

      var conversions = 0;
      if (row.actions) {
        for (var j = 0; j < row.actions.length; j++) {
          if (row.actions[j].action_type === "offsite_conversion.fb_pixel_lead" ||
              row.actions[j].action_type === "lead" ||
              row.actions[j].action_type === "offsite_conversion") {
            conversions += parseInt(row.actions[j].value || 0);
          }
        }
      }

      var ctr = impressions > 0 ? Math.round(clicks / impressions * 10000) / 100 : 0;
      var cpc = clicks > 0 ? Math.round(cost / clicks) : 0;
      var cvr = clicks > 0 ? Math.round(conversions / clicks * 10000) / 100 : 0;
      var cpa = conversions > 0 ? Math.round(cost / conversions) : 0;

      allRows.push([date, campaignName, impressions, clicks, cost, conversions, ctr, cpc, cvr, cpa]);
    }

    // ページネーション
    url = (json.paging && json.paging.next) ? json.paging.next : null;
  }

  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, 10).setValues(allRows);
  }

  Logger.log(allRows.length + " rows written to " + SHEET_NAME);
}

function buildApiUrl(adAccountId, accessToken, since, until) {
  var baseUrl = "https://graph.facebook.com/" + API_VERSION + "/act_" + adAccountId + "/insights";
  var params = [
    "fields=date_start,campaign_name,impressions,clicks,spend,actions",
    "time_range=" + encodeURIComponent('{"since":"' + since + '","until":"' + until + '"}'),
    "time_increment=1",
    "level=campaign",
    "limit=500",
    "access_token=" + accessToken
  ];
  return baseUrl + "?" + params.join("&");
}

function getYesterday() {
  var now = new Date();
  now.setDate(now.getDate() - 1);
  return Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
}
