var SHEET_URL         = 'https://docs.google.com/spreadsheets/d/108CnrRFLlMxmrS-9PQwUN1-JypODUkin3c_0UXlTV2o/edit';
var DATE_WINDOW_DAYS  = 3;
var CAMPAIGNS_TAB     = 'Campaigns';
var SEARCH_TERMS_TAB  = 'SearchTerms';
var KEYWORDS_TAB      = 'Keywords';
var CHANGE_EVENTS_TAB = 'ChangeEvents';

// change_event.change_date_time → `date` is the YYYY-MM-DD slice so upsertRows'
// window + account_id gates work unchanged. The Google-side LIMIT is 10000 per
// query and the WHERE filter is mandatory (the API rejects unbounded scans).
var CHANGE_EVENT_LIMIT = 10000;

var CAMPAIGN_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'status', 'channel_type', 'final_url',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value', 'cpa', 'roas',
  'currency_code',
  'primary_status', 'primary_status_reasons',
  'last_updated',
  'daily_budget', 'target_cpa', 'bidding_strategy_type',
  'account_timezone',
  'campaign_geo',
  // Trailing additions (ensureHeaders auto-appends these to existing sheets):
  // advertising_channel_sub_type + the campaign's conversion goals (JSON of
  // {category, origin, biddable} from the campaign_conversion_goal resource).
  // The derived "marketing objective" label is computed downstream in Laravel.
  'channel_sub_type', 'conversion_goals',
  // metrics.invalid_clicks (count Google filtered as invalid — not billed) +
  // the paired metrics.invalid_click_rate (invalid ÷ (invalid + valid) clicks).
  'invalid_clicks', 'invalid_click_rate'
];

var CAMPAIGN_KEY_COLS = ['account_id', 'campaign_id', 'date'];

var SEARCH_TERM_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'ad_group_name', 'keyword', 'search_term',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value',
  'currency_code',
  'last_updated'
];

var SEARCH_TERM_KEY_COLS = ['account_id', 'campaign_id', 'ad_group_name', 'keyword', 'search_term', 'date'];

var KEYWORD_HEADERS = [
  'date', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
  'ad_group_id', 'ad_group_name', 'criterion_id', 'keyword', 'match_type',
  // Keyword quality metrics (date-segmented via metrics.historical_*):
  'quality_score', 'ad_relevance', 'landing_page_experience', 'expected_ctr',
  'impressions', 'clicks', 'cost', 'ctr', 'avg_cpc',
  'conversions', 'conversion_value',
  'currency_code',
  'last_updated'
];

// criterion_id is unique per keyword within an ad group; combined with
// account/campaign/ad_group it's globally unique. Including `date` keeps the
// in-window replacement scoped correctly across runs (same as the other tabs).
var KEYWORD_KEY_COLS = ['account_id', 'campaign_id', 'ad_group_id', 'criterion_id', 'date'];

var CHANGE_EVENT_HEADERS = [
  'date', 'change_date_time', 'account_name', 'account_id',
  'change_event_resource_name',
  'change_resource_type', 'resource_change_operation',
  'changed_resource_name',
  'campaign_resource_name', 'ad_group_resource_name',
  'campaign_id', 'campaign_name',
  'user_email', 'client_type',
  'changed_fields_json', 'old_resource_json', 'new_resource_json',
  'last_updated'
];

// (account_id, change_event_resource_name) is globally unique — Google's event
// resource_name embeds `{date_time}~{order}`. Including `date` keeps the
// upsertRows in-window replacement scoped correctly across runs.
var CHANGE_EVENT_KEY_COLS = ['account_id', 'change_event_resource_name', 'date'];

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
  // One ISO8601 timestamp per runReports() call — every row written this run
  // gets the same `last_updated` value, so Laravel can trust the column to
  // reflect "when did the Ads Script actually run", not "when did the Sheet
  // row last get upserted" (the Laravel command stamps its own column too).
  var ctx = {
    accountId:    account.getCustomerId(),
    accountName:  account.getName(),
    currencyCode: account.getCurrencyCode(),
    timezone:     account.getTimeZone(),
    dateRange:    computeDateRange(DATE_WINDOW_DAYS, account.getTimeZone()),
    campaignUrls: fetchCampaignFinalUrls(),
    campaignGeo:  fetchCampaignGeoTargets(),
    conversionGoals: fetchCampaignConversionGoals(),
    runTimestamp: new Date().toISOString()
  };

  Logger.log(
    'runReports → ' + ctx.accountName + ' (' + ctx.accountId + ') ' +
    'currency=' + ctx.currencyCode + ' ' +
    'tz=' + ctx.timezone + ' ' +
    'window=' + ctx.dateRange.start + '..' + ctx.dateRange.end
  );

  writeCampaigns(ss, ctx);
  writeSearchTerms(ss, ctx);
  writeKeywords(ss, ctx);
  writeChangeEvents(ss, ctx);

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

// Resolve each campaign's location TARGETING to a country code → { campaignId:
// 'AU' | 'US' | … | 'MULTI' }. This is the campaign's geo targeting (what
// Google Ads serves to), NOT the account timezone — the Campaigns dashboard
// warns when it diverges from the landing's intended geo.
//
// Two-step: (1) collect each campaign's positive (non-negative) LOCATION
// criteria as geo_target_constant resource names, (2) resolve those constants
// to ISO country codes. Sub-country targets (region/city) still carry the
// country_code of the country they belong to, so the reduction works for
// country-, region-, and city-level targeting alike. A campaign with no
// location criteria (targets everywhere) is simply absent → '' downstream.
//
// Wrapped whole: any GAQL/permission error must degrade to an empty map and
// leave the column blank rather than abort writeCampaigns and the whole run.
function fetchCampaignGeoTargets() {
  try {
    var gtcByCampaign = {};   // campaignId → { geoTargetConstantResourceName: true }
    var allGtc = {};          // geoTargetConstantResourceName → true
    var critQuery =
      'SELECT ' +
        'campaign.id, ' +
        'campaign_criterion.location.geo_target_constant, ' +
        'campaign_criterion.negative ' +
      'FROM campaign_criterion ' +
      "WHERE campaign_criterion.type = 'LOCATION' " +
        "AND campaign.status != 'REMOVED'";

    var it = AdsApp.search(critQuery);
    while (it.hasNext()) {
      var r = it.next();
      var crit = r.campaignCriterion || {};
      // Negative (excluded) locations don't define where the campaign serves.
      if (crit.negative === true) continue;
      var gtc = crit.location && crit.location.geoTargetConstant;
      if (!gtc) continue;
      var cid = String(r.campaign.id);
      if (!gtcByCampaign[cid]) gtcByCampaign[cid] = {};
      gtcByCampaign[cid][gtc] = true;
      allGtc[gtc] = true;
    }

    var countryByGtc = resolveGeoTargetCountryCodes(Object.keys(allGtc));

    var geoByCampaign = {};
    for (var c in gtcByCampaign) {
      if (!gtcByCampaign.hasOwnProperty(c)) continue;
      var countries = {};
      for (var g in gtcByCampaign[c]) {
        if (!gtcByCampaign[c].hasOwnProperty(g)) continue;
        var cc = countryByGtc[g];
        if (cc) countries[cc] = true;
      }
      var list = [];
      for (var k in countries) if (countries.hasOwnProperty(k)) list.push(k);
      geoByCampaign[c] = list.length === 1 ? list[0] : (list.length > 1 ? 'MULTI' : '');
    }
    Logger.log('Campaign geo targets resolved: ' + Object.keys(geoByCampaign).length + ' campaigns');
    return geoByCampaign;
  } catch (e) {
    Logger.log('fetchCampaignGeoTargets failed (leaving campaign_geo blank): ' + e);
    return {};
  }
}

// Map geo_target_constant resource names → ISO-3166 alpha-2 country code.
// geo_target_constant.country_code is the country a target belongs to for
// every target_type (Country/Region/City/…), so one lookup covers them all.
// Filter by the numeric id (parsed from the "geoTargetConstants/{id}" resource
// name) — the canonical, broadly-supported GAQL form — then key results back to
// the original resource name. Chunked IN-clauses keep each query within limits.
function resolveGeoTargetCountryCodes(resourceNames) {
  var out = {};
  if (!resourceNames || !resourceNames.length) return out;

  var ids = [];
  var resourceById = {};
  for (var i = 0; i < resourceNames.length; i++) {
    var rn = String(resourceNames[i]);
    var parts = rn.split('/');
    var id = parts[parts.length - 1];
    if (id) {
      ids.push(id);
      resourceById[id] = rn;
    }
  }

  var CHUNK = 500;
  for (var k = 0; k < ids.length; k += CHUNK) {
    var chunk = ids.slice(k, k + CHUNK);
    var q =
      'SELECT ' +
        'geo_target_constant.id, ' +
        'geo_target_constant.country_code ' +
      'FROM geo_target_constant ' +
      'WHERE geo_target_constant.id IN (' + chunk.join(', ') + ')';
    var it = AdsApp.search(q);
    while (it.hasNext()) {
      var r = it.next();
      var gtc = r.geoTargetConstant || {};
      var gid = (gtc.id !== null && gtc.id !== undefined) ? String(gtc.id) : null;
      if (gid && gtc.countryCode && resourceById[gid]) {
        out[resourceById[gid]] = String(gtc.countryCode).toUpperCase();
      }
    }
  }
  return out;
}

// Resolve each campaign's conversion goals → { campaignId: [{category, origin,
// biddable}, …] }. `campaign_conversion_goal` is an attributes-only resource
// (no metrics, no date segmentation) — it lists which conversion-goal
// categories a campaign counts/optimizes for. The biddable categories are the
// signal Laravel uses to derive a "marketing objective" label downstream.
//
// The `.campaign` field is a resource name (customers/{cid}/campaigns/{id});
// we parse the trailing id rather than join campaign.id, so a join quirk can't
// abort the whole run. Wrapped whole: any GAQL/permission error degrades to an
// empty map and leaves the column blank — mirrors fetchCampaignGeoTargets().
function fetchCampaignConversionGoals() {
  try {
    var query =
      'SELECT ' +
        'campaign_conversion_goal.campaign, ' +
        'campaign_conversion_goal.category, ' +
        'campaign_conversion_goal.origin, ' +
        'campaign_conversion_goal.biddable ' +
      'FROM campaign_conversion_goal';

    var goalsByCampaign = {};
    var it = AdsApp.search(query);
    while (it.hasNext()) {
      var r = it.next();
      var g = r.campaignConversionGoal || {};
      var resName = String(g.campaign || '');
      var parts = resName.split('/');
      var cid = parts[parts.length - 1];
      if (!cid) continue;

      var cat = normEnumToken(g.category);
      var origin = normEnumToken(g.origin);
      if (!goalsByCampaign[cid]) goalsByCampaign[cid] = [];
      goalsByCampaign[cid].push({
        category: cat,
        origin: origin,
        biddable: g.biddable === true
      });
    }
    Logger.log('Campaign conversion goals resolved: ' + Object.keys(goalsByCampaign).length + ' campaigns');
    return goalsByCampaign;
  } catch (e) {
    Logger.log('fetchCampaignConversionGoals failed (leaving conversion_goals blank): ' + e);
    return {};
  }
}

// Normalize a Google enum token: UNSPECIFIED/UNKNOWN/blank → '' (so Laravel
// stores it as "absent" rather than noise). Otherwise pass the uppercase token.
function normEnumToken(v) {
  var s = String(v || '').toUpperCase();
  if (s === '' || s === 'UNSPECIFIED' || s === 'UNKNOWN') return '';
  return s;
}

// QualityScoreBucket enum (historical_*_quality_score) → keep only the three
// meaningful buckets; everything else (UNKNOWN/UNSPECIFIED/blank) → ''.
function normQualityBucket(v) {
  var s = String(v || '').toUpperCase();
  if (s === 'BELOW_AVERAGE' || s === 'AVERAGE' || s === 'ABOVE_AVERAGE') return s;
  return '';
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
      'campaign.primary_status, ' +
      'campaign.primary_status_reasons, ' +
      'campaign.advertising_channel_type, ' +
      'campaign.advertising_channel_sub_type, ' +
      'campaign.bidding_strategy_type, ' +
      'campaign.target_cpa.target_cpa_micros, ' +
      'campaign.maximize_conversions.target_cpa_micros, ' +
      'campaign_budget.amount_micros, ' +
      'metrics.impressions, ' +
      'metrics.clicks, ' +
      'metrics.cost_micros, ' +
      'metrics.ctr, ' +
      'metrics.average_cpc, ' +
      'metrics.conversions, ' +
      'metrics.conversions_value, ' +
      'metrics.invalid_clicks, ' +
      'metrics.invalid_click_rate ' +
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

    var biddingType = r.campaign.biddingStrategyType || '';
    if (biddingType === 'UNSPECIFIED' || biddingType === 'UNKNOWN') biddingType = '';

    // TARGET_CPA puts the target on campaign.target_cpa.target_cpa_micros;
    // MAXIMIZE_CONVERSIONS may carry an optional cap on
    // campaign.maximize_conversions.target_cpa_micros. Coalesce.
    var tcpaMicros = (r.campaign.targetCpa && r.campaign.targetCpa.targetCpaMicros) ||
                     (r.campaign.maximizeConversions && r.campaign.maximizeConversions.targetCpaMicros) || 0;
    var tcpa = Number(tcpaMicros) > 0 ? Number(tcpaMicros) / 1e6 : '';

    var budget = Number((r.campaignBudget && r.campaignBudget.amountMicros) || 0) / 1e6;

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
      ctx.currencyCode,
      r.campaign.primaryStatus || '',
      JSON.stringify(r.campaign.primaryStatusReasons || []),
      ctx.runTimestamp,
      budget,
      tcpa,
      biddingType,
      ctx.timezone,
      ctx.campaignGeo[r.campaign.id] || '',
      normEnumToken(r.campaign.advertisingChannelSubType),
      JSON.stringify(ctx.conversionGoals[r.campaign.id] || []),
      Number(r.metrics.invalidClicks || 0),
      Number(r.metrics.invalidClickRate || 0)
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
      ctx.currencyCode,
      ctx.runTimestamp
    ]);
  }

  Logger.log('Search-term rows collected: ' + rows.length);
  upsertRows(sheet, SEARCH_TERM_HEADERS, SEARCH_TERM_KEY_COLS, rows, ctx.accountId, ctx.dateRange);
}

function writeKeywords(ss, ctx) {
  var sheet = ss.getSheetByName(KEYWORDS_TAB) || ss.insertSheet(KEYWORDS_TAB);
  ensureHeaders(sheet, KEYWORD_HEADERS);

  // keyword_view is the keyword (ad_group_criterion) level — quality metrics
  // live here, NOT on search_term_view. The historical_* metrics are
  // date-segmentable (the plain ad_group_criterion.quality_info.* snapshot is
  // not, and would repeat today's value across every date row).
  var query =
    'SELECT ' +
      'segments.date, ' +
      'campaign.id, ' +
      'campaign.name, ' +
      'ad_group.id, ' +
      'ad_group.name, ' +
      'ad_group_criterion.criterion_id, ' +
      'ad_group_criterion.keyword.text, ' +
      'ad_group_criterion.keyword.match_type, ' +
      'metrics.historical_quality_score, ' +
      'metrics.historical_creative_quality_score, ' +
      'metrics.historical_landing_page_quality_score, ' +
      'metrics.historical_search_predicted_ctr, ' +
      'metrics.impressions, ' +
      'metrics.clicks, ' +
      'metrics.cost_micros, ' +
      'metrics.ctr, ' +
      'metrics.average_cpc, ' +
      'metrics.conversions, ' +
      'metrics.conversions_value ' +
    'FROM keyword_view ' +
    "WHERE segments.date BETWEEN '" + ctx.dateRange.start + "' AND '" + ctx.dateRange.end + "' " +
      "AND ad_group_criterion.type = 'KEYWORD'";

  var iter = AdsApp.search(query);
  var rows = [];
  while (iter.hasNext()) {
    var r = iter.next();
    var crit = r.adGroupCriterion || {};
    var kw = crit.keyword || {};
    var cost   = Number(r.metrics.costMicros || 0) / 1e6;
    var avgCpc = Number(r.metrics.averageCpc || 0) / 1e6;

    // historical_quality_score is int64 1–10; 0/absent means Google has no QS
    // for this keyword/date — emit '' so Laravel stores NULL, not 0.
    var qs = Number(r.metrics.historicalQualityScore || 0);

    rows.push([
      r.segments.date,
      ctx.accountName,
      ctx.accountId,
      r.campaign.id,
      r.campaign.name,
      r.adGroup.id,
      r.adGroup.name,
      crit.criterionId,
      kw.text || '',
      normEnumToken(kw.matchType),
      qs >= 1 ? qs : '',
      normQualityBucket(r.metrics.historicalCreativeQualityScore),
      normQualityBucket(r.metrics.historicalLandingPageQualityScore),
      normQualityBucket(r.metrics.historicalSearchPredictedCtr),
      Number(r.metrics.impressions || 0),
      Number(r.metrics.clicks || 0),
      cost,
      Number(r.metrics.ctr || 0),
      avgCpc,
      Number(r.metrics.conversions || 0),
      Number(r.metrics.conversionsValue || 0),
      ctx.currencyCode,
      ctx.runTimestamp
    ]);
  }

  Logger.log('Keyword rows collected: ' + rows.length);
  upsertRows(sheet, KEYWORD_HEADERS, KEYWORD_KEY_COLS, rows, ctx.accountId, ctx.dateRange);
}

function writeChangeEvents(ss, ctx) {
  var sheet = ss.getSheetByName(CHANGE_EVENTS_TAB) || ss.insertSheet(CHANGE_EVENTS_TAB);
  ensureHeaders(sheet, CHANGE_EVENT_HEADERS);

  // change_event is datetime-segmented, not date-segmented like the metrics
  // resources. Use the same rolling window in account timezone; LIMIT +
  // ORDER BY are both mandated by Google's change_event API.
  var query =
    'SELECT ' +
      'change_event.resource_name, ' +
      'change_event.change_date_time, ' +
      'change_event.change_resource_type, ' +
      'change_event.change_resource_name, ' +
      'change_event.resource_change_operation, ' +
      'change_event.changed_fields, ' +
      'change_event.client_type, ' +
      'change_event.user_email, ' +
      'change_event.old_resource, ' +
      'change_event.new_resource, ' +
      'change_event.campaign, ' +
      'change_event.ad_group, ' +
      'campaign.id, ' +
      'campaign.name ' +
    'FROM change_event ' +
    "WHERE change_event.change_date_time BETWEEN '" + ctx.dateRange.start + " 00:00:00' AND '" + ctx.dateRange.end + " 23:59:59' " +
    'ORDER BY change_event.change_date_time DESC ' +
    'LIMIT ' + CHANGE_EVENT_LIMIT;

  var iter = AdsApp.search(query);
  var rows = [];
  while (iter.hasNext()) {
    var r = iter.next();
    var ce = r.changeEvent || {};
    var dt = ce.changeDateTime || '';
    var date = dt.length >= 10 ? dt.substring(0, 10) : '';

    rows.push([
      date,
      dt,
      ctx.accountName,
      ctx.accountId,
      ce.resourceName || '',
      ce.changeResourceType || '',
      ce.resourceChangeOperation || '',
      ce.changeResourceName || '',
      ce.campaign || '',
      ce.adGroup || '',
      (r.campaign && r.campaign.id) ? String(r.campaign.id) : '',
      (r.campaign && r.campaign.name) ? r.campaign.name : '',
      ce.userEmail || '',
      ce.clientType || '',
      JSON.stringify(ce.changedFields || {}),
      JSON.stringify(ce.oldResource || {}),
      JSON.stringify(ce.newResource || {}),
      ctx.runTimestamp
    ]);
  }

  Logger.log('Change-event rows collected: ' + rows.length);
  if (rows.length >= CHANGE_EVENT_LIMIT) {
    Logger.log('change_event LIMIT ' + CHANGE_EVENT_LIMIT + ' reached — possible truncation');
  }
  upsertRows(sheet, CHANGE_EVENT_HEADERS, CHANGE_EVENT_KEY_COLS, rows, ctx.accountId, ctx.dateRange);
}

// Columns whose values are text but happen to look like something Sheets'
// "Automatic" format would auto-coerce — `date` → Date object, `last_updated`
// → Date object (ISO8601 with `T`/`Z` is close enough that some locales parse
// it). Coercion breaks the downstream gviz CSV contract: the Laravel sync
// expects the raw string we wrote, not a locale-formatted Date readback.
// `change_date_time` carries Google's `YYYY-MM-DD HH:MM:SS` form — same risk.
var TEXT_FORMATTED_COLUMNS = ['date', 'last_updated', 'change_date_time'];

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    applyTextFormats(sheet, headers, 0);
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
    // Format newly-added text columns before any data is written, otherwise
    // the first setValues() with an ISO8601 string lands in a cell that's
    // still "Automatic" and may get coerced to Date before we can reformat.
    applyTextFormats(sheet, headers, existingCount);
  }
}

// Apply `@` (plain text) format to any TEXT_FORMATTED_COLUMNS present at
// header index >= skipBefore. Also reformats existing data cells so stale
// Date-object values in a pre-format sheet get preserved as strings on next
// read rather than coerced at read time.
function applyTextFormats(sheet, headers, skipBefore) {
  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;
  for (var i = 0; i < TEXT_FORMATTED_COLUMNS.length; i++) {
    var idx = headers.indexOf(TEXT_FORMATTED_COLUMNS[i]);
    if (idx < 0 || idx < skipBefore) continue;
    sheet.getRange(2, idx + 1, maxRows - 1, 1).setNumberFormat('@');
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
