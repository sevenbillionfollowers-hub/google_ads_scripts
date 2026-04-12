var SHEET_URL        = 'https://docs.google.com/spreadsheets/d/108CnrRFLlMxmrS-9PQwUN1-JypODUkin3c_0UXlTV2o/edit';
var DATE_WINDOW_DAYS = 2;
var CAMPAIGNS_TAB    = 'Campaigns';
var SEARCH_TERMS_TAB = 'SearchTerms';

var CAMPAIGN_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'status', 'channel_type', 'final_url',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value', 'cpa', 'roas'
];

var SEARCH_TERM_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'ad_group_name', 'search_term',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value'
];

function getSheetUrl() {
  if (SHEET_URL.indexOf('REPLACE_ME') !== -1) {
    throw new Error(
      'SHEET_URL is not configured — edit main.js in the hosting repo ' +
      'and replace REPLACE_ME with the target Google Sheet URL.'
    );
  }
  return SHEET_URL;
}

function runReports(ss) {
  var account = AdsApp.currentAccount();
  var ctx = {
    accountId:    account.getCustomerId(),
    accountName:  account.getName(),
    timezone:     account.getTimeZone(),
    dateRange:    computeDateRange(DATE_WINDOW_DAYS, account.getTimeZone()),
    campaignUrls: fetchCampaignFinalUrls()
  };

  Logger.log(
    'runReports → ' + ctx.accountName + ' (' + ctx.accountId + ') ' +
    'tz=' + ctx.timezone + ' ' +
    'window=' + ctx.dateRange.start + '..' + ctx.dateRange.end
  );

  writeCampaigns(ss, ctx);
  writeSearchTerms(ss, ctx);

  Logger.log('runReports → done');
}

function fetchCampaignFinalUrls() {
  var urls = {};
  var query =
    'SELECT ' +
      'campaign.id, ' +
      'ad_group_ad.ad.final_urls ' +
    'FROM ad_group_ad ' +
    "WHERE ad_group_ad.status = 'ENABLED' " +
      "AND ad_group.status = 'ENABLED' " +
      "AND campaign.status = 'ENABLED'";

  var iter = AdsApp.search(query);
  while (iter.hasNext()) {
    var r = iter.next();
    var cid = r.campaign.id;
    if (urls[cid]) continue;
    var finalUrls = (r.adGroupAd && r.adGroupAd.ad && r.adGroupAd.ad.finalUrls) || [];
    if (finalUrls.length > 0) {
      urls[cid] = finalUrls[0];
    }
  }
  return urls;
}

function writeCampaigns(ss, ctx) {
  var sheet = ss.getSheetByName(CAMPAIGNS_TAB) || ss.insertSheet(CAMPAIGNS_TAB);
  ensureHeaders(sheet, CAMPAIGN_HEADERS);

  var query =
    'SELECT ' +
      'segments.date, ' +
      'campaign.id, ' +
      'campaign.name, ' +
      'campaign.status, ' +
      'campaign.advertising_channel_type, ' +
      'metrics.impressions, ' +
      'metrics.clicks, ' +
      'metrics.cost_micros, ' +
      'metrics.ctr, ' +
      'metrics.average_cpc, ' +
      'metrics.conversions, ' +
      'metrics.conversions_value ' +
    'FROM campaign ' +
    "WHERE segments.date BETWEEN '" + ctx.dateRange.start + "' AND '" + ctx.dateRange.end + "'";

  var iter = AdsApp.search(query);
  var rows = [];
  while (iter.hasNext()) {
    var r = iter.next();
    var cost            = Number(r.metrics.costMicros || 0) / 1e6;
    var avgCpc          = Number(r.metrics.averageCpc || 0) / 1e6;
    var conversions     = Number(r.metrics.conversions || 0);
    var conversionValue = Number(r.metrics.conversionsValue || 0);
    var cpa  = conversions > 0 ? cost / conversions : 0;
    var roas = cost > 0 ? conversionValue / cost : 0;

    rows.push([
      r.segments.date,
      ctx.accountName,
      ctx.accountId,
      r.campaign.id,
      r.campaign.name,
      r.campaign.status,
      r.campaign.advertisingChannelType,
      ctx.campaignUrls[r.campaign.id] || '',
      Number(r.metrics.impressions || 0),
      Number(r.metrics.clicks || 0),
      cost,
      Number(r.metrics.ctr || 0),
      avgCpc,
      conversions,
      conversionValue,
      cpa,
      roas
    ]);
  }

  Logger.log('Campaigns rows collected: ' + rows.length);
  upsertRows(sheet, CAMPAIGN_HEADERS, rows, ctx.accountId, ctx.dateRange);
}

function writeSearchTerms(ss, ctx) {
  var sheet = ss.getSheetByName(SEARCH_TERMS_TAB) || ss.insertSheet(SEARCH_TERMS_TAB);
  ensureHeaders(sheet, SEARCH_TERM_HEADERS);

  var query =
    'SELECT ' +
      'segments.date, ' +
      'campaign.id, ' +
      'campaign.name, ' +
      'ad_group.id, ' +
      'ad_group.name, ' +
      'search_term_view.search_term, ' +
      'metrics.impressions, ' +
      'metrics.clicks, ' +
      'metrics.cost_micros, ' +
      'metrics.ctr, ' +
      'metrics.average_cpc, ' +
      'metrics.conversions, ' +
      'metrics.conversions_value ' +
    'FROM search_term_view ' +
    "WHERE segments.date BETWEEN '" + ctx.dateRange.start + "' AND '" + ctx.dateRange.end + "'";

  var iter = AdsApp.search(query);
  var rows = [];
  while (iter.hasNext()) {
    var r = iter.next();
    var cost   = Number(r.metrics.costMicros || 0) / 1e6;
    var avgCpc = Number(r.metrics.averageCpc || 0) / 1e6;

    rows.push([
      r.segments.date,
      ctx.accountName,
      ctx.accountId,
      r.campaign.id,
      r.campaign.name,
      r.adGroup.name,
      r.searchTermView.searchTerm,
      Number(r.metrics.impressions || 0),
      Number(r.metrics.clicks || 0),
      cost,
      Number(r.metrics.ctr || 0),
      avgCpc,
      Number(r.metrics.conversions || 0),
      Number(r.metrics.conversionsValue || 0)
    ]);
  }

  Logger.log('Search-term rows collected: ' + rows.length);
  upsertRows(sheet, SEARCH_TERM_HEADERS, rows, ctx.accountId, ctx.dateRange);
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var dateColIdx = headers.indexOf('date') + 1;
    if (dateColIdx > 0) {
      sheet.getRange(2, dateColIdx, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    }
    return;
  }

  var currentHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  for (var h = 0; h < headers.length; h++) {
    if (String(currentHeader[h]) !== headers[h]) {
      throw new Error(
        'Header mismatch on tab "' + sheet.getName() + '" col ' + (h + 1) +
        ': expected "' + headers[h] + '", got "' + currentHeader[h] + '". ' +
        'Delete the tab and re-run to regenerate it.'
      );
    }
  }
}

function upsertRows(sheet, headers, newRows, accountId, dateRange) {
  var dateCol = headers.indexOf('date');
  var acctCol = headers.indexOf('account_id');
  if (dateCol < 0 || acctCol < 0) {
    throw new Error('upsertRows: headers must include date + account_id');
  }

  var sheetTz = sheet.getParent().getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();
  var keepRows = [];

  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < existing.length; i++) {
      var row = existing[i];
      var sameAccount = String(row[acctCol]) === String(accountId);
      var rowDate = toDateStr(row[dateCol], sheetTz);
      var inWindow = rowDate >= dateRange.start && rowDate <= dateRange.end;
      if (!(sameAccount && inWindow)) {
        keepRows.push(row);
      }
    }
  }

  for (var j = 0; j < newRows.length; j++) {
    newRows[j][dateCol] = toDateStr(newRows[j][dateCol], sheetTz);
  }

  var finalRows = keepRows.concat(newRows);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  }
}

function toDateStr(v, tz) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v);
}

function computeDateRange(days, tz) {
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var endDate   = addDays(todayStr, -1);
  var startDate = addDays(endDate, -(days - 1));
  return { start: startDate, end: endDate };
}

function addDays(dateStr, delta) {
  var parts = dateStr.split('-');
  var d = new Date(Date.UTC(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10),
    12, 0, 0
  ));
  d.setUTCDate(d.getUTCDate() + delta);
  var y = d.getUTCFullYear();
  var m = d.getUTCMonth() + 1;
  var day = d.getUTCDate();
  return y + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
}
