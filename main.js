var SHEET_URL        = 'https://docs.google.com/spreadsheets/d/108CnrRFLlMxmrS-9PQwUN1-JypODUkin3c_0UXlTV2o/edit';
var DATE_WINDOW_DAYS = 3;
var CAMPAIGNS_TAB    = 'Campaigns';
var SEARCH_TERMS_TAB = 'SearchTerms';

var CAMPAIGN_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'status', 'channel_type', 'final_url',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value', 'cpa', 'roas',
  'currency_code'
];

var CAMPAIGN_KEY_COLS = ['account_id', 'campaign_id', 'date'];

var SEARCH_TERM_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'ad_group_name', 'keyword', 'search_term',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value',
  'currency_code'
];

var SEARCH_TERM_KEY_COLS = ['account_id', 'campaign_id', 'ad_group_name', 'keyword', 'search_term', 'date'];

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
    currencyCode: account.getCurrencyCode(),
    timezone:     account.getTimeZone(),
    dateRange:    computeDateRange(DATE_WINDOW_DAYS, account.getTimeZone()),
    campaignUrls: fetchCampaignFinalUrls()
  };

  Logger.log(
    'runReports → ' + ctx.accountName + ' (' + ctx.accountId + ') ' +
    'currency=' + ctx.currencyCode + ' ' +
    'tz=' + ctx.timezone + ' ' +
    'window=' + ctx.dateRange.start + '..' + ctx.dateRange.end
  );

  writeCampaigns(ss, ctx);
  writeSearchTerms(ss, ctx);

  Logger.log('runReports → done');
}

function fetchCampaignFinalUrls() {
  // Prefer ENABLED active ads for URL attribution. Falling back to PAUSED only
  // when no ENABLED ad exists avoids a REMOVED ad's URL overwriting a live one.
  var enabledUrls = collectFinalUrls(
    "WHERE ad_group_ad.status = 'ENABLED' " +
      "AND ad_group.status = 'ENABLED' " +
      "AND campaign.status = 'ENABLED'"
  );
  var pausedUrls = collectFinalUrls(
    "WHERE ad_group_ad.status IN ('ENABLED', 'PAUSED') " +
      "AND ad_group.status IN ('ENABLED', 'PAUSED') " +
      "AND campaign.status IN ('ENABLED', 'PAUSED')"
  );

  var urls = {};
  for (var cid in pausedUrls) if (pausedUrls.hasOwnProperty(cid)) urls[cid] = pausedUrls[cid];
  for (var cid2 in enabledUrls) if (enabledUrls.hasOwnProperty(cid2)) urls[cid2] = enabledUrls[cid2];
  return urls;
}

function collectFinalUrls(whereClause) {
  // Sort so we always pick the smallest ad id per campaign → deterministic
  // URL attribution across runs even when a campaign has multiple ads.
  var query =
    'SELECT ' +
      'campaign.id, ' +
      'ad_group_ad.ad.id, ' +
      'ad_group_ad.ad.final_urls ' +
    'FROM ad_group_ad ' +
    whereClause + ' ' +
    'ORDER BY campaign.id, ad_group_ad.ad.id';

  var urls = {};
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
      roas,
      ctx.currencyCode
    ]);
  }

  Logger.log('Campaigns rows collected: ' + rows.length);
  upsertRows(sheet, CAMPAIGN_HEADERS, CAMPAIGN_KEY_COLS, rows, ctx.accountId, ctx.dateRange);
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
      'segments.keyword.info.text, ' +
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

    var keywordText =
      (r.segments && r.segments.keyword && r.segments.keyword.info && r.segments.keyword.info.text) || '';

    rows.push([
      r.segments.date,
      ctx.accountName,
      ctx.accountId,
      r.campaign.id,
      r.campaign.name,
      r.adGroup.name,
      keywordText,
      r.searchTermView.searchTerm,
      Number(r.metrics.impressions || 0),
      Number(r.metrics.clicks || 0),
      cost,
      Number(r.metrics.ctr || 0),
      avgCpc,
      Number(r.metrics.conversions || 0),
      Number(r.metrics.conversionsValue || 0),
      ctx.currencyCode
    ]);
  }

  Logger.log('Search-term rows collected: ' + rows.length);
  upsertRows(sheet, SEARCH_TERM_HEADERS, SEARCH_TERM_KEY_COLS, rows, ctx.accountId, ctx.dateRange);
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

  var existingCount = sheet.getLastColumn();
  var overlap = Math.min(existingCount, headers.length);
  var currentHeader = sheet.getRange(1, 1, 1, overlap).getValues()[0];

  // Existing prefix must match exactly — catches actual corruption.
  for (var h = 0; h < overlap; h++) {
    if (String(currentHeader[h]) !== headers[h]) {
      throw new Error(
        'Header mismatch on tab "' + sheet.getName() + '" col ' + (h + 1) +
        ': expected "' + headers[h] + '", got "' + currentHeader[h] + '". ' +
        'Delete the tab and re-run to regenerate it.'
      );
    }
  }

  // We've added trailing columns to the expected header — extend the sheet
  // in place so existing rows are preserved; blanks fill the new cells
  // until downstream rewrites them.
  if (headers.length > existingCount) {
    var extra = headers.slice(existingCount);
    sheet.getRange(1, existingCount + 1, 1, extra.length).setValues([extra]);
  }
}

function upsertRows(sheet, headers, keyCols, newRows, accountId, dateRange) {
  var dateCol = headers.indexOf('date');
  var acctCol = headers.indexOf('account_id');
  if (dateCol < 0 || acctCol < 0) {
    throw new Error('upsertRows: headers must include date + account_id');
  }

  var keyIdx = [];
  for (var kc = 0; kc < keyCols.length; kc++) {
    var idx = headers.indexOf(keyCols[kc]);
    if (idx < 0) throw new Error('upsertRows: unknown key column ' + keyCols[kc]);
    keyIdx.push(idx);
  }

  var sheetTz = sheet.getParent().getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();

  // Normalize incoming dates before key comparison so string keys match.
  for (var j = 0; j < newRows.length; j++) {
    newRows[j][dateCol] = toDateStr(newRows[j][dateCol], sheetTz);
  }

  // Build set of unique keys from newRows so we only replace rows that were
  // re-fetched this run. In-window rows absent from the new batch (e.g. a
  // campaign that had zero impressions today and is omitted by GAQL) are
  // preserved — otherwise history silently disappears.
  var newKeys = {};
  for (var n = 0; n < newRows.length; n++) {
    newKeys[makeKey(newRows[n], keyIdx)] = true;
  }

  var keepRows = [];
  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < existing.length; i++) {
      var row = existing[i];
      // Normalize date on existing row too so key comparison is apples-to-apples.
      row[dateCol] = toDateStr(row[dateCol], sheetTz);

      var sameAccount = String(row[acctCol]) === String(accountId);
      var inWindow = row[dateCol] >= dateRange.start && row[dateCol] <= dateRange.end;
      if (sameAccount && inWindow && newKeys[makeKey(row, keyIdx)]) {
        continue; // will be replaced by newRows
      }
      keepRows.push(row);
    }
  }

  var finalRows = keepRows.concat(newRows);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  }
}

function makeKey(row, keyIdx) {
  var parts = [];
  for (var i = 0; i < keyIdx.length; i++) {
    parts.push(String(row[keyIdx[i]]));
  }
  return parts.join('\u241F'); // SYMBOL FOR UNIT SEPARATOR — won't appear in any real field
}

function toDateStr(v, tz) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v);
}

function computeDateRange(days, tz) {
  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var endDate   = todayStr;
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
