var SHEET_URL         = 'https://docs.google.com/spreadsheets/d/108CnrRFLlMxmrS-9PQwUN1-JypODUkin3c_0UXlTV2o/edit';
var DATE_WINDOW_DAYS  = 3;
var CAMPAIGNS_TAB     = 'Campaigns';
var SEARCH_TERMS_TAB  = 'SearchTerms';
var KEYWORDS_TAB      = 'Keywords';
var CHANGE_EVENTS_TAB = 'ChangeEvents';

// How many days of history each tab keeps. The ONLY consumer of this sheet is
// the Laravel sync, which pushes `where A >= now-7days` down to gviz — it never
// reads older rows (history lives in the DB). Before this constant existed the
// tabs grew forever and EVERY run of EVERY account rewrote the whole accumulated
// history, which is what pushed a single write phase past STALE_LOCK_MS and
// produced the `refreshSheetLock: lost ownership` aborts. 21 days is 3x the
// consumer's lookback, so the sync can be down for two weeks without data loss.
var RETENTION_DAYS = 21;

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
  'invalid_clicks', 'invalid_click_rate',
  // metrics.average_target_cpa_micros — the traffic-weighted target CPA the bid
  // strategy actually optimized toward over the write window (NOT the static
  // `target_cpa` setting). Resolved by fetchCampaignAvgTargetCpa() as one
  // window-aggregated value per campaign.
  'avg_target_cpa'
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

var DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

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
  // dateRange is needed both as ctx.dateRange and as the window for the avg
  // target CPA metric query, so resolve it into a local first.
  var dateRange = computeDateRange(DATE_WINDOW_DAYS, account.getTimeZone());
  var ctx = {
    accountId:    account.getCustomerId(),
    accountName:  account.getName(),
    currencyCode: account.getCurrencyCode(),
    timezone:     account.getTimeZone(),
    dateRange:    dateRange,
    // Everything older than this is deleted from every tab (see RETENTION_DAYS).
    pruneBefore:  addDays(dateRange.end, -(RETENTION_DAYS - 1)),
    // Resolved once instead of once per upsertRows call — it is a Sheets round
    // trip and it never changes mid-run.
    sheetTz:      ss.getSpreadsheetTimeZone(),
    campaignUrls: fetchCampaignFinalUrls(),
    campaignGeo:  fetchCampaignGeoTargets(),
    conversionGoals: fetchCampaignConversionGoals(),
    avgTargetCpa: fetchCampaignAvgTargetCpa(dateRange),
    runTimestamp: new Date().toISOString()
  };

  Logger.log(
    'runReports → ' + ctx.accountName + ' (' + ctx.accountId + ') ' +
    'currency=' + ctx.currencyCode + ' ' +
    'tz=' + ctx.timezone + ' ' +
    'window=' + ctx.dateRange.start + '..' + ctx.dateRange.end + ' ' +
    'retention>=' + ctx.pruneBefore
  );

  // PHASE 1 — COLLECT. Every GAQL query runs OUTSIDE the mutex. None of it
  // touches the shared spreadsheet, and it is the slowest part of a run (four
  // paged Ads API scans). Holding a fleet-wide lock across it used to multiply
  // the critical section by the full report-fetch time for all N accounts.
  var batches = [
    { tab: CAMPAIGNS_TAB,     headers: CAMPAIGN_HEADERS,     keys: CAMPAIGN_KEY_COLS,     rows: collectCampaigns(ctx),    expectTz: ctx.timezone },
    { tab: SEARCH_TERMS_TAB,  headers: SEARCH_TERM_HEADERS,  keys: SEARCH_TERM_KEY_COLS,  rows: collectSearchTerms(ctx),  expectTz: null },
    { tab: KEYWORDS_TAB,      headers: KEYWORD_HEADERS,      keys: KEYWORD_KEY_COLS,      rows: collectKeywords(ctx),     expectTz: null },
    { tab: CHANGE_EVENTS_TAB, headers: CHANGE_EVENT_HEADERS, keys: CHANGE_EVENT_KEY_COLS, rows: collectChangeEvents(ctx), expectTz: null }
  ];

  // NOTHING TO WRITE ⇒ NEVER TAKE THE LOCK. The cost of the write phase is a
  // function of the SHARED SHEET's size, not of this account's data: upsertRows
  // scans the whole key-column prefix of every tab (measured 2026-09-02:
  // 893,075 cells across the four tabs, 85% of it SearchTerms) before it looks
  // at `newRows`. So an account with zero rows used to queue for the fleet-wide
  // mutex and then hold it for ~25 s to write nothing — measured on account
  // 971-477-6033, which has never produced a single row in any tab: 3 s of its
  // own GAQL work, 246 s waiting for the lock, then a 17 s hold before it was
  // stomped mid-run.
  //
  // Such accounts are invisible to every census the Sheet supports (they
  // contribute no rows, so they appear in no tab and in no DB table), which is
  // why the fleet's true lock demand was being under-counted.
  //
  // The ONLY thing a zero-row run still accomplishes is the retention prune,
  // and that is fleet-shared housekeeping — 71-89 accounts with real rows do it
  // every hour. The single degenerate case is an entirely idle fleet, which by
  // definition has nothing to prune.
  var hasWork = false;
  for (var w = 0; w < batches.length; w++) {
    if (batches[w].rows.length > 0) { hasWork = true; break; }
  }
  if (!hasWork) {
    Logger.log('runReports → no rows in any tab; skipping the write phase entirely ' +
      '(lock not acquired, retention prune left to accounts that have rows)');
    return;
  }

  // PHASE 2 — WRITE. All N accounts write to ONE shared spreadsheet on
  // independent hourly schedules. Serialize the write phase behind a
  // sheet-anchored mutex: a per-script LockService cannot help, because every
  // account runs main.js as its OWN Apps Script project, so the lock must live
  // in the one resource they all share — the spreadsheet itself.
  //
  // upsertRows() heartbeats the lock around every Sheets round trip it makes
  // (see refreshSheetLock), so STALE_LOCK_MS measures "time since the holder
  // last made progress" at round-trip granularity rather than at write-phase
  // granularity. The old code only heartbeat BETWEEN the four phases, so one
  // slow phase looked abandoned, a waiter reclaimed a live lock, and the holder
  // aborted with `lost ownership … (reclaimed as stale)`.
  var lockUuid = acquireSheetLock(ss);
  try {
    for (var i = 0; i < batches.length; i++) {
      var b = batches[i];
      refreshSheetLock(ss, lockUuid, true);
      var sheet = ss.getSheetByName(b.tab) || ss.insertSheet(b.tab);
      var headerWidth = ensureHeaders(sheet, b.headers);
      ensureTextFormats(ss, lockUuid, sheet, b.headers);
      upsertRows(ss, lockUuid, sheet, b.headers, headerWidth, b.keys, b.rows, ctx, b.expectTz);
    }
    // Final ownership assertion. Without it a reclaim during the LAST phase was
    // never detected — nothing checked the lock after writeChangeEvents.
    refreshSheetLock(ss, lockUuid, true);
  } finally {
    releaseSheetLock(ss, lockUuid);
  }

  Logger.log('runReports → done');
}

// ---------------------------------------------------------------------------
// Ads API collectors — pure GAQL, no Sheets access, all run outside the mutex.
// ---------------------------------------------------------------------------

function fetchCampaignFinalUrls() {
  // ONE scan of ad_group_ad, not two. The old code ran an ENABLED-only query
  // and then an ENABLED|PAUSED query whose result set is a strict superset of
  // the first, doubling the ad scan for every account on every run. Sorting by
  // (campaign.id, ad.id) keeps the pick deterministic across runs; a fully
  // ENABLED ad still beats a PAUSED one so a paused ad's URL can never
  // overwrite a live one.
  //
  // DELIBERATELY NOT wrapped in try/catch, unlike the other ctx fetches. Those
  // degrade to a blank column because a blank is harmless downstream. This one
  // is not: `GoogleAdsSyncStatsCommand` treats the Sheet as authoritative for
  // `final_url`/`final_domain` and writes them on every upsert, so emitting a
  // blank would NULL the stored URL for every campaign in the account. Letting
  // the run die is the self-healing failure — the next hourly run re-pulls the
  // whole 3-day window and the Sheet keeps its previous, correct values.
  {
    var query =
      'SELECT ' +
        'campaign.id, ' +
        'campaign.status, ' +
        'ad_group.status, ' +
        'ad_group_ad.status, ' +
        'ad_group_ad.ad.id, ' +
        'ad_group_ad.ad.final_urls ' +
      'FROM ad_group_ad ' +
      "WHERE ad_group_ad.status IN ('ENABLED', 'PAUSED') " +
        "AND ad_group.status IN ('ENABLED', 'PAUSED') " +
        "AND campaign.status IN ('ENABLED', 'PAUSED') " +
      'ORDER BY campaign.id, ad_group_ad.ad.id';

    var best = {};
    var iter = AdsApp.search(query);
    while (iter.hasNext()) {
      var r = iter.next();
      var finalUrls = (r.adGroupAd && r.adGroupAd.ad && r.adGroupAd.ad.finalUrls) || [];
      if (finalUrls.length === 0) continue;

      var cid = r.campaign.id;
      var fullyEnabled =
        r.adGroupAd.status === 'ENABLED' &&
        r.adGroup.status === 'ENABLED' &&
        r.campaign.status === 'ENABLED';

      var cur = best[cid];
      if (!cur || (fullyEnabled && !cur.enabled)) {
        best[cid] = { url: finalUrls[0], enabled: fullyEnabled };
      }
    }

    var urls = {};
    for (var c in best) if (best.hasOwnProperty(c)) urls[c] = best[c].url;
    return urls;
  }
}

// Resolve each campaign's location TARGETING to a country code → { campaignId:
// 'AU' | 'US' | … | 'MULTI' }. This is the campaign's geo targeting (what
// Google Ads serves to), NOT the account timezone — the Campaigns dashboard
// warns when it diverges from the landing's intended geo.
//
// Two-step: (1) collect each campaign's positive (non-negative, non-REMOVED)
// LOCATION criteria as geo_target_constant resource names, (2) resolve those
// constants to ISO country codes. Sub-country targets (region/city) still carry
// the country_code of the country they belong to, so the reduction works for
// country-, region-, and city-level targeting alike. A campaign with no
// location criteria (targets everywhere) is simply absent → '' downstream.
//
// campaign_criterion.status is SELECTED and filtered in JS rather than pushed
// into the WHERE clause: a rejected WHERE would blow up the whole query and the
// catch below would silently blank campaign_geo for every campaign, whereas an
// absent field just leaves `status` undefined and behaves exactly as before.
//
// Wrapped whole: any GAQL/permission error must degrade to an empty map and
// leave the column blank rather than abort the campaign write and the whole run.
function fetchCampaignGeoTargets() {
  try {
    var gtcByCampaign = {};   // campaignId → { geoTargetConstantResourceName: true }
    var allGtc = {};          // geoTargetConstantResourceName → true
    var critQuery =
      'SELECT ' +
        'campaign.id, ' +
        'campaign_criterion.status, ' +
        'campaign_criterion.location.geo_target_constant, ' +
        'campaign_criterion.negative ' +
      'FROM campaign_criterion ' +
      "WHERE campaign_criterion.type = 'LOCATION' " +
        "AND campaign.status != 'REMOVED'";

    var it = AdsApp.search(critQuery);
    while (it.hasNext()) {
      var r = it.next();
      var crit = r.campaignCriterion || {};
      // Negative (excluded) locations don't define where the campaign serves,
      // and a REMOVED criterion no longer targets anything — counting it made
      // a campaign that was re-targeted from one country to another report the
      // union of both, i.e. a false 'MULTI'.
      if (crit.negative === true) continue;
      if (crit.status === 'REMOVED') continue;
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

// Resolve each campaign's AVERAGE target CPA over the write window →
// { campaignId: number (account currency) }. Unlike campaign.target_cpa (the
// current *setting*, already captured per-row from
// campaign.target_cpa.target_cpa_micros), metrics.average_target_cpa_micros is
// the traffic-weighted target the bid strategy actually optimized toward across
// the period — folding in device bid adjustments, ad-group target overrides,
// and any mid-window changes to the target. This is the value Google surfaces
// as the "Avg. target CPA" column/header.
//
// We deliberately keep segments.date in the WHERE filter ONLY (not in SELECT),
// so Google returns one window-aggregated row per campaign — matching the
// header figure for the date range, and side-stepping any segments.date
// selectability constraint on the metric.
//
// Wrapped whole: any GAQL/permission/compat error degrades to an empty map and
// leaves the column blank rather than abort the campaign write and the whole
// run — mirrors fetchCampaignGeoTargets()/fetchCampaignConversionGoals().
function fetchCampaignAvgTargetCpa(dateRange) {
  try {
    var query =
      'SELECT ' +
        'campaign.id, ' +
        'metrics.average_target_cpa_micros ' +
      'FROM campaign ' +
      "WHERE segments.date BETWEEN '" + dateRange.start + "' AND '" + dateRange.end + "'";

    var out = {};
    var it = AdsApp.search(query);
    while (it.hasNext()) {
      var r = it.next();
      var micros = (r.metrics && r.metrics.averageTargetCpaMicros) || 0;
      var v = Number(micros) > 0 ? Number(micros) / 1e6 : 0;
      // Only campaigns running a tCPA-style strategy report a value; the rest
      // come back 0/absent → leave them out so the column stores NULL, not 0.
      if (v > 0) out[String(r.campaign.id)] = v;
    }
    Logger.log('Campaign avg target CPA resolved: ' + Object.keys(out).length + ' campaigns');
    return out;
  } catch (e) {
    Logger.log('fetchCampaignAvgTargetCpa failed (leaving avg_target_cpa blank): ' + e);
    return {};
  }
}

function collectCampaigns(ctx) {
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
      Number(r.metrics.invalidClickRate || 0),
      // Window-aggregated avg target CPA (same value on every day-row for a
      // campaign; the Laravel sync collapses to one parent row per campaign).
      ctx.avgTargetCpa[r.campaign.id] || ''
    ]);
  }

  Logger.log('Campaigns rows collected: ' + rows.length);
  return rows;
}

function collectSearchTerms(ctx) {
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
  return rows;
}

function collectKeywords(ctx) {
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
  return rows;
}

function collectChangeEvents(ctx) {
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
  return rows;
}

// ---------------------------------------------------------------------------
// Sheet headers + grid capacity
// ---------------------------------------------------------------------------

// Columns pinned to `@` (plain text) so Sheets' "Automatic" format cannot
// reinterpret what we write.
//
// - `date` / `last_updated` / `change_date_time` would be coerced to Date
//   objects, breaking the gviz CSV contract the Laravel sync reads and the
//   string key comparison upsertRows does on read-back.
// - The free-text Ads fields matter for a different reason: setValues() parses
//   a leading `=` as a FORMULA. A real search term like `=free games to play`
//   lands as `#NAME?`, permanently corrupting that row's upsert key so the row
//   is re-appended every hour and never matched again.
// The three columns the PREVIOUS main.js already pinned. They are useless as a
// "has this tab been formatted?" probe — every live tab already has them set,
// so probing one would short-circuit ensureTextFormats() and the columns added
// below would never actually be applied in production.
var LEGACY_TEXT_COLUMNS = ['date', 'last_updated', 'change_date_time'];

var TEXT_FORMATTED_COLUMNS = [
  'date', 'last_updated', 'change_date_time',
  'account_id', 'account_name', 'campaign_name', 'ad_group_name',
  'keyword', 'search_term', 'final_url', 'user_email',
  'primary_status_reasons', 'conversion_goals',
  'change_event_resource_name', 'changed_resource_name',
  'campaign_resource_name', 'ad_group_resource_name',
  'changed_fields_json', 'old_resource_json', 'new_resource_json'
];

// Keep this much spare grid below the last data row. A fresh insertSheet() is
// 1000x26, and deleteRows() (retention pruning) SHRINKS getMaxRows() — without
// slack, an account still running the previous main.js (which sizes its write
// off its own read, not off the grid) could ask for a range past the last row
// and throw for the duration of a staggered rollout.
var GRID_ROW_SLACK = 2000;

// Returns the effective header width so upsertRows doesn't have to re-read it.
function ensureHeaders(sheet, headers) {
  // A fresh insertSheet() has 26 columns and CAMPAIGN_HEADERS is 31 wide — the
  // header write itself used to throw "out of bounds" on a regenerated tab,
  // which is exactly what the "Delete the tab and re-run" error below tells an
  // operator to do.
  ensureColumnCapacity(sheet, headers.length);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    applyTextFormats(sheet, headers, 0, 2, sheet.getMaxRows());
    return headers.length;
  }

  // Measure the HEADER ROW, not getLastColumn(). getLastColumn() is a
  // whole-sheet high-water mark that clearContent() never resets, so a single
  // stray cell parked to the right of the headers used to trip the fail-closed
  // width guard below and wedge every account permanently.
  var existingCount = liveHeaderWidth(sheet);
  var overlap = Math.min(existingCount, headers.length);
  var currentHeader = overlap > 0 ? sheet.getRange(1, 1, 1, overlap).getValues()[0] : [];

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
    applyTextFormats(sheet, headers, existingCount, 2, sheet.getMaxRows());
    return headers.length;
  }

  return existingCount;
}

// Width of the header ROW: index of its last non-blank cell.
function liveHeaderWidth(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var width = 0;
  for (var i = 0; i < hdr.length; i++) {
    if (String(hdr[i]) !== '') width = i + 1;
  }
  return width;
}

function ensureColumnCapacity(sheet, needed) {
  var maxCols = sheet.getMaxColumns();
  if (needed <= maxCols) return;
  sheet.insertColumnsAfter(maxCols, needed - maxCols);
}

// Grow the grid so `neededLastRow` is addressable, plus slack. Newly inserted
// rows inherit no number format, so the text columns are pinned on the new band
// (only the new band — re-pinning the whole column on every grow would be a
// large formatting op inside the mutex for no benefit).
function ensureRowCapacity(sheet, neededLastRow, headers) {
  var maxRows = sheet.getMaxRows();
  if (neededLastRow + GRID_ROW_SLACK <= maxRows) return;
  var firstNew = maxRows + 1;
  sheet.insertRowsAfter(maxRows, (neededLastRow + GRID_ROW_SLACK) - maxRows);
  applyTextFormats(sheet, headers, 0, firstNew, sheet.getMaxRows());
}

// Self-healing format pin. The live tabs already carry every header, so the
// create/extend paths in ensureHeaders() never fire on them — without this, the
// columns added to TEXT_FORMATTED_COLUMNS would never actually be pinned in
// production and the leading-`=` formula hazard would stay live forever.
//
// Steady-state cost is ONE single-cell read per tab per run: sample the number
// format of the last data row in the first pinned column, and only re-pin when
// it isn't already '@'.
function ensureTextFormats(ss, lockUuid, sheet, headers) {
  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;
  // Probe a column this version ADDED, never one the previous version already
  // pinned — otherwise every live tab reports "@" and returns early.
  var sentinel = -1;
  for (var i = 0; i < TEXT_FORMATTED_COLUMNS.length; i++) {
    var name = TEXT_FORMATTED_COLUMNS[i];
    if (indexOfStr(LEGACY_TEXT_COLUMNS, name) >= 0) continue;
    var idx = headers.indexOf(name);
    if (idx >= 0) { sentinel = idx; break; }
  }
  if (sentinel < 0) return;
  var probeRow = Math.max(2, Math.min(sheet.getLastRow(), maxRows));
  if (sheet.getRange(probeRow, sentinel + 1).getNumberFormat() === '@') return;
  Logger.log(sheet.getName() + ': pinning text columns to plain-text format (' +
    (maxRows - 1) + ' rows, chunked)');
  applyTextFormatsChunked(ss, lockUuid, sheet, headers, 0, 2, maxRows);
}

// Rows per setNumberFormat() call in the full-tab re-pin. This is the ONE
// unbounded formatting op in the file and it sits inside the mutex, ahead of
// upsertRows.
//
// Un-chunked it was the largest single operation any run could perform with NO
// heartbeat inside it: on the live SearchTerms tab (2026-09-02) that is 8
// pinned columns x ~96,500 rows = 772,144 cells in one call. A holder is
// reclaimed as stale after STALE_LOCK_MS (4 min) WITHOUT A HEARTBEAT — not
// after 4 min of no progress — so a single op that outruns that horizon hands
// the lock to a waiter while the holder is still alive and still writing. That
// is the precondition for the `lost ownership … (reclaimed as stale)` abort,
// and it is the only operation in the run big enough to reach it on its own.
var TEXT_FORMAT_CHUNK_ROWS = 10000;

// applyTextFormats() over [fromRow, toRow] in bands, asserting + refreshing
// lock ownership between bands. Ownership is re-checked BEFORE each band, so a
// run that already lost the lock stops re-formatting a tab another account is
// writing instead of finishing the whole pass first.
function applyTextFormatsChunked(ss, lockUuid, sheet, headers, skipBefore, fromRow, toRow) {
  for (var start = fromRow; start <= toRow; start += TEXT_FORMAT_CHUNK_ROWS) {
    var end = Math.min(start + TEXT_FORMAT_CHUNK_ROWS - 1, toRow);
    refreshSheetLock(ss, lockUuid, false);
    applyTextFormats(sheet, headers, skipBefore, start, end);
  }
}

// Apply `@` (plain text) format to any TEXT_FORMATTED_COLUMNS present at
// header index >= skipBefore, over rows [fromRow, toRow]. One getRangeList()
// call covers every column, so this stays a single Sheets round trip no matter
// how many columns are pinned.
function applyTextFormats(sheet, headers, skipBefore, fromRow, toRow) {
  if (toRow < fromRow) return;
  var a1 = [];
  for (var i = 0; i < TEXT_FORMATTED_COLUMNS.length; i++) {
    var idx = headers.indexOf(TEXT_FORMATTED_COLUMNS[i]);
    if (idx < 0 || idx < skipBefore) continue;
    var letter = colLetter(idx + 1);
    a1.push(letter + fromRow + ':' + letter + toRow);
  }
  if (a1.length === 0) return;
  sheet.getRangeList(a1).setNumberFormat('@');
}

// Array.prototype.indexOf on strings — spelled out because this file targets
// the oldest engine an Ads Script may still be compiled on.
function indexOfStr(list, needle) {
  for (var i = 0; i < list.length; i++) if (list[i] === needle) return i;
  return -1;
}

function colLetter(n) {
  var s = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ---------------------------------------------------------------------------
// Sheet-anchored cross-account mutex
// ---------------------------------------------------------------------------
//
// Apps Script LockService is scoped per project, and every Google Ads account
// runs main.js as its OWN project, so a ScriptLock/DocumentLock does NOT
// serialize the N accounts that share this one Sheet. The only thing they all
// share is the spreadsheet — so the lock lives in it: `_lock!A1` holds
// `{uuid}@{ts}`. The uuid is the lock's identity (stable for the whole hold);
// the ts is a heartbeat the holder bumps as it makes progress. A lock whose
// heartbeat is older than STALE_LOCK_MS is treated as abandoned (a run that
// crashed or hit the 30-min cap) and reclaimed, so a dead run can never wedge
// the fleet.
//
// TOKEN FORMAT IS FROZEN at `{uuid}@{ts}`. stub.js re-fetches main.js from
// GitHub on every run, so for up to an hour after a push the fleet runs MIXED
// versions against this one cell. Any format change would make the two
// populations mutually invisible and guarantee the concurrent-write corruption
// this mutex exists to prevent.
//
// ROUND-ALIGNED CLAIM. The old acquire was read → decide → blind setValue →
// sleep(250-750ms) → re-read. Because that settle was the same order of
// magnitude as a Sheets write round trip, two contenders could each confirm
// their own uuid before the other's write landed, and BOTH would proceed —
// unserialized, which is the corruption mode the mutex exists to stop.
// SpreadsheetApp offers no conditional write, so instead of racing on delay we
// align on the wall clock: a claim may only be issued in the first
// CLAIM_WINDOW_MS of a CLAIM_ROUND_MS round, and every claimant confirms at
// CLAIM_CONFIRM_MS into that SAME round. Every claim of round R has therefore
// committed before any confirm read of round R happens, so all contenders
// observe the same final value and exactly one sees its own uuid.
//
// INVARIANT: LOCK_ACQUIRE_TIMEOUT_MS MUST exceed STALE_LOCK_MS, so a waiter
// either wins when the holder releases or reclaims a genuinely dead lock — it
// can never run out of patience while the lock is still un-reclaimable.
var LOCK_SHEET = '_lock';
var STALE_LOCK_MS = 240000;            // 4 min with NO heartbeat ⇒ holder is dead → reclaim
var LOCK_ACQUIRE_TIMEOUT_MS = 900000;  // 15 min patient wait (well inside the 30-min script cap)
var LOCK_HEARTBEAT_MIN_MS = 20000;     // don't re-write the cell more often than this
var CLAIM_ROUND_MS = 6000;
var CLAIM_WINDOW_MS = 1500;
var CLAIM_CONFIRM_MS = 4500;
var BACKOFF_BASE_MS = 1500;
var BACKOFF_CAP_MS = 15000;

// Last time WE wrote the heartbeat. main.js is eval'd fresh per run, so this is
// per-run state.
var _lockLastBeatAt = 0;

// uuid half of a `{uuid}@{ts}` lock cell (or '' for an empty/garbage cell).
function lockUuidOf(cellValue) {
  return String(cellValue || '').split('@')[0];
}

// ts half. A cell whose ts is missing or non-numeric (manual edit, a coerced
// value, a half-written cell) returns 0 = "infinitely old", so it is reclaimed
// on the next attempt instead of wedging every account forever.
function lockTsOf(cellValue) {
  var parts = String(cellValue || '').split('@');
  var ts = parts.length > 1 ? Number(parts[1]) : NaN;
  return (isFinite(ts) && ts > 0) ? ts : 0;
}

// Decorrelated exponential backoff. The old flat 1-3s poll meant every waiter
// hammered the shared document a few times a second for up to 8 minutes — the
// contention itself slowed the holder's writes — and gave an unlucky account no
// increasing chance of winning.
function nextBackoff(prev) {
  var next = Math.floor(prev + Math.random() * prev * 2);
  if (next > BACKOFF_CAP_MS) next = BACKOFF_CAP_MS - Math.floor(Math.random() * 2000);
  return next;
}

function acquireSheetLock(ss) {
  var meta = ss.getSheetByName(LOCK_SHEET) || ss.insertSheet(LOCK_SHEET);
  var cell = meta.getRange('A1');
  var uuid = Utilities.getUuid();
  var startedAt = new Date().getTime();
  var backoff = BACKOFF_BASE_MS;
  var attempt = 0;
  var cur = '';

  var missedWindows = 0;

  while (true) {
    // Timeout is checked FIRST, so no path through the loop — including the
    // round-alignment retry below — can spin past the deadline.
    if ((new Date().getTime() - startedAt) > LOCK_ACQUIRE_TIMEOUT_MS) {
      throw new Error('acquireSheetLock: gave up after ' + attempt + ' attempts / ' +
        Math.round((new Date().getTime() - startedAt) / 1000) + 's (lock held by "' + cur + '")');
    }

    attempt++;
    cur = String(cell.getValue() || '');
    var now = new Date().getTime();

    if (cur === '' || (now - lockTsOf(cur)) > STALE_LOCK_MS) {
      var offset = now % CLAIM_ROUND_MS;
      // If the read round trip itself keeps overrunning the claim window we
      // would never claim at all, so after a few misses claim unaligned — that
      // degrades to the old probabilistic settle rather than never acquiring.
      if (offset > CLAIM_WINDOW_MS && missedWindows < 5) {
        // Too late in this round to claim safely. Sleep to the next round
        // boundary and re-enter the loop — the fresh read at the top both
        // re-checks the lock and lands us inside the next claim window.
        missedWindows++;
        Utilities.sleep(CLAIM_ROUND_MS - offset + 10);
        continue;
      }
      var roundStart = now - offset;
      cell.setValue(uuid + '@' + now);
      SpreadsheetApp.flush();

      // From here our token is in the cell but we do not yet own it. If
      // anything throws before we either win or retract (a Sheets quota error
      // in the confirm read, say), the token would sit there un-owned and
      // freeze the whole fleet until STALE_LOCK_MS elapsed. Retract on the way
      // out so a transient error costs one round, not four minutes.
      try {

      // Sleep EXACTLY to this round's shared confirm instant — no floor. The
      // floor is what would break the invariant: if our claim write committed
      // after the round's confirm instant, the other claimants of this round
      // already read a value that did not contain it, so one of them may have
      // confirmed ITSELF the winner. Confirming late would then produce two
      // holders — the very race the round alignment exists to remove. Fail the
      // claim instead, and retract our token by re-stamping it with an
      // already-expired ts so the cell is immediately reclaimable rather than
      // wedged for STALE_LOCK_MS. `uuid@1` still parses as {uuid}@{ts} for BOTH
      // the new lockTsOf() and the previous main.js's Number(split('@')[1]),
      // so mixed-version accounts both read it as a long-stale lock.
      var settle = (roundStart + CLAIM_CONFIRM_MS) - new Date().getTime();
      if (settle <= 0) {
        if (lockUuidOf(cell.getValue()) === uuid) {
          cell.setValue(uuid + '@1');
          SpreadsheetApp.flush();
        }
      } else {
        Utilities.sleep(settle);
        if (lockUuidOf(cell.getValue()) === uuid) {
          _lockLastBeatAt = new Date().getTime();
          Logger.log('Sheet lock acquired after ' + attempt + ' attempt(s) / ' +
            Math.round((_lockLastBeatAt - startedAt) / 1000) + 's');
          return uuid;
        }
      }

      } catch (claimErr) {
        try {
          if (lockUuidOf(cell.getValue()) === uuid) {
            cell.setValue(uuid + '@1');
            SpreadsheetApp.flush();
          }
        } catch (retractErr) { /* best effort — stale reclaim is the backstop */ }
        throw claimErr;
      }

      missedWindows = 0;   // we did get to claim; the alignment path is healthy
    }

    backoff = nextBackoff(backoff);
    Utilities.sleep(backoff);
  }
}

// Assert we still own the lock, and bump the heartbeat if it's due.
//
// The READ always happens, so this doubles as a pre-write ownership assertion:
// callers invoke it immediately BEFORE each destructive Sheets operation, so a
// run that lost the lock stops before corrupting rows rather than after (the
// old code only checked at phase boundaries, i.e. after the damage).
// The WRITE is throttled to LOCK_HEARTBEAT_MIN_MS (or forced) so heartbeating
// at round-trip granularity doesn't itself become the dominant cost.
function refreshSheetLock(ss, uuid, force) {
  if (!uuid) return;
  var meta = ss.getSheetByName(LOCK_SHEET);
  // A vanished _lock tab means we are no longer serialized against anyone.
  // Continuing to write would be exactly the unlocked concurrent write this
  // whole mechanism exists to prevent, so fail instead of silently proceeding.
  if (!meta) {
    throw new Error('refreshSheetLock: ' + LOCK_SHEET +
      ' tab disappeared mid-run — aborting rather than writing unserialized.');
  }
  var cell = meta.getRange('A1');
  if (lockUuidOf(cell.getValue()) !== uuid) {
    throw new Error('refreshSheetLock: lost ownership of ' + LOCK_SHEET +
      '!A1 mid-run (reclaimed as stale) — aborting to avoid concurrent-write corruption.');
  }
  var now = new Date().getTime();
  if (!force && (now - _lockLastBeatAt) < LOCK_HEARTBEAT_MIN_MS) return;
  cell.setValue(uuid + '@' + now);
  SpreadsheetApp.flush();
  _lockLastBeatAt = now;
}

function releaseSheetLock(ss, uuid) {
  if (!uuid) return;
  var meta = ss.getSheetByName(LOCK_SHEET);
  if (!meta) return;
  var cell = meta.getRange('A1');
  if (lockUuidOf(cell.getValue()) === uuid) {
    cell.clearContent();
    SpreadsheetApp.flush();
  }
}

// ---------------------------------------------------------------------------
// Upsert
// ---------------------------------------------------------------------------

// Retention deletes are issued as contiguous blocks; cap how many blocks one
// run may issue so a heavily fragmented first pass can't stretch the lock hold.
// Whatever is left over is deleted by the next run.
var MAX_PRUNE_BLOCKS_PER_RUN = 25;

// Purely diagnostic: log a warning if our in-place updates fragment into more
// than this many contiguous blocks (each block costs one round trip). There is
// deliberately no "rewrite the whole span" fallback — that would rewrite the
// interleaved rows of OTHER accounts, re-introducing exactly the cross-account
// clobber the mutex exists to prevent.
var MAX_UPDATE_BLOCKS = 40;

// In-place upsert.
//
// The previous implementation was a whole-tab read → clearContent → write-all,
// per tab, per account, per run — O(entire accumulated history) even though a
// run only ever changes its own ~150 rows in a 3-day window. With N accounts on
// an hourly cadence that is the direct cause of the production
// `refreshSheetLock: lost ownership` abort (one phase outran STALE_LOCK_MS),
// and it also made the tab visibly TORN to the gviz reader: the Laravel sync
// polls every 10 minutes and could observe the tab between the clear and the
// rewrite (measured: SearchTerms row count reading 1177 / 3296 / 5990 / 7382
// seconds apart while accounts rotated through).
//
// Now: read only the KEY columns, locate each incoming row's existing position,
// overwrite those rows in place in contiguous blocks, append the rest, and
// delete anything past the retention horizon. Rows are never mass-cleared, so
// a concurrent reader can never see a truncated tab, and the payload per run is
// proportional to the write window instead of to all history.
function upsertRows(ss, lockUuid, sheet, headers, headerWidth, keyCols, newRows, ctx, expectAccountTimezone) {
  var name = sheet.getName();
  var dateCol = headers.indexOf('date');
  var acctCol = headers.indexOf('account_id');
  if (dateCol < 0 || acctCol < 0) {
    throw new Error('upsertRows: headers must include date + account_id');
  }

  // FAIL CLOSED on header-width skew. If the live Sheet's HEADER ROW is wider
  // than the columns this (possibly CDN-stale) main.js knows about, writing a
  // short row would strand the trailing columns from whatever campaign
  // physically preceded it — the cross-campaign column smear this guard exists
  // to prevent. Refuse to write rather than corrupt; the account self-heals on
  // its next run once it fetches the current main.js.
  if (headerWidth > headers.length) {
    throw new Error('upsertRows: tab "' + name + '" has ' + headerWidth +
      ' header columns but this main.js knows ' + headers.length +
      ' — refusing to write a short row into a wide sheet (CDN-stale code). ' +
      'Re-run after the new main.js propagates.');
  }

  var keyIdx = [];
  var indexWidth = Math.max(dateCol, acctCol) + 1;
  for (var kc = 0; kc < keyCols.length; kc++) {
    var idx = headers.indexOf(keyCols[kc]);
    if (idx < 0) throw new Error('upsertRows: unknown key column ' + keyCols[kc]);
    keyIdx.push(idx);
    if (idx + 1 > indexWidth) indexWidth = idx + 1;
  }

  var accountId = String(ctx.accountId);
  var tz = ctx.sheetTz;

  // Normalize incoming dates before key comparison so string keys match, and
  // collapse any duplicate keys inside this batch itself (last wins — the same
  // rule the Laravel sync applies when it upserts).
  var byKey = {};
  var order = [];
  // Lowest `date` among the rows we are about to write. No sheet row older than
  // this can share a key with an incoming row, because `date` is part of every
  // key set — which is what lets scanIndex skip reading their key columns.
  // Derived from the rows themselves rather than from ctx.dateRange so the
  // claim is self-evident and cannot drift if the collectors ever widen.
  var windowStart = null;
  for (var j = 0; j < newRows.length; j++) {
    newRows[j][dateCol] = toDateStr(newRows[j][dateCol], tz);
    var nd = newRows[j][dateCol];
    if (windowStart === null || nd < windowStart) windowStart = nd;
    var nk = makeKey(newRows[j], keyIdx);
    if (!byKey.hasOwnProperty(nk)) order.push(nk);
    byKey[nk] = newRows[j];
  }
  // FAIL SAFE: the whole optimization rests on `date` being part of the key.
  // All four key sets include it today; if one ever stops, fall back to the
  // full-width scan rather than silently missing matches (which would append a
  // duplicate every hour).
  var dateIsKey = indexOfStr(keyCols, 'date') >= 0;

  var scan = scanIndex(sheet, indexWidth, dateCol, acctCol, keyIdx, accountId, tz, ctx.pruneBefore, dateIsKey, windowStart);
  var pruned = 0;

  if (scan.drop.length > 0) {
    var del = deleteRowBlocks(ss, lockUuid, sheet, scan.drop);
    pruned = del.deleted;
    if (del.skippedBlocks > 0) {
      Logger.log(name + ': retention prune capped at ' + MAX_PRUNE_BLOCKS_PER_RUN +
        ' blocks — ' + del.skippedBlocks + ' block(s) deferred to the next run');
    }
    // Row numbers below every deletion have shifted; re-read the (now smaller)
    // key index rather than trying to arithmetically adjust them.
    refreshSheetLock(ss, lockUuid, false);
    scan = scanIndex(sheet, indexWidth, dateCol, acctCol, keyIdx, accountId, tz, ctx.pruneBefore, dateIsKey, windowStart);
  }

  var pos = scan.pos;
  var updates = [];
  var appends = [];
  for (var o = 0; o < order.length; o++) {
    var k = order[o];
    if (pos.hasOwnProperty(k)) {
      updates.push([pos[k], byKey[k]]);
    } else {
      appends.push(byKey[k]);
    }
  }
  updates.sort(function (a, b) { return a[0] - b[0]; });

  // [startRow, count] ranges this run wrote. Every row inside them is OURS —
  // this function never writes over a row belonging to another account, which
  // is what makes the self-check below an exact assertion.
  var written = [];

  if (updates.length > 0) {
    // One setValues per contiguous run of rows. Each account's rows for a tab
    // are written as one batch and then updated in place, so in the steady
    // state this is a single block; the previous main.js also kept a run's
    // window rows contiguous (it re-appended them together), so even a
    // staggered rollout does not fragment them.
    var blocks = groupContiguous(updates);
    if (blocks.length > MAX_UPDATE_BLOCKS) {
      Logger.log(name + ': WARNING — ' + updates.length + ' updates fragmented into ' +
        blocks.length + ' blocks; expect one round trip each');
    }
    for (var b = 0; b < blocks.length; b++) {
      refreshSheetLock(ss, lockUuid, false);
      sheet.getRange(blocks[b].start, 1, blocks[b].rows.length, headers.length).setValues(blocks[b].rows);
      written.push([blocks[b].start, blocks[b].rows.length]);
    }
  }

  if (appends.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    if (startRow < 2) startRow = 2;
    ensureRowCapacity(sheet, startRow + appends.length - 1, headers);
    refreshSheetLock(ss, lockUuid, false);
    sheet.getRange(startRow, 1, appends.length, headers.length).setValues(appends);
    written.push([startRow, appends.length]);
  } else {
    // Keep headroom below the last row even on a pure-update run, so an account
    // still on the previous main.js (which sizes its write off its own read)
    // always has grid to write into after a prune shrank getMaxRows().
    ensureRowCapacity(sheet, sheet.getLastRow(), headers);
  }

  // Post-write self-check — cheap insurance against a mutex failure or a still-
  // unpatched account writing without the lock. Re-read exactly the ranges we
  // wrote and assert they carry THIS account's id, and — for the Campaigns tab
  // — this account's timezone. A mismatch means another account's write landed
  // on our rows; throw loudly so the run fails visibly instead of shipping a
  // contaminated row.
  //
  // Read-side cost control, WITHOUT weakening the assertion: it used to re-read
  // every written block at FULL header width, one round trip each — on the
  // measured account 12 round trips (7.2 s) inside the mutex, to look at two
  // columns. Now it reads only as far as the columns it checks, and merges
  // blocks whose gap is cheaper than the round trip it saves. Rows inside a
  // merged gap belong to other accounts and are skipped by the `isWritten`
  // filter; every row we actually wrote is still checked, exactly as before.
  if (written.length > 0) {
    SpreadsheetApp.flush();
    var tzCol = headers.indexOf('account_timezone');
    var checkWidth = Math.max(acctCol, tzCol) + 1;
    var isWritten = {};
    var writtenRows = [];
    for (var wi = 0; wi < written.length; wi++) {
      for (var wn = 0; wn < written[wi][1]; wn++) {
        var wr = written[wi][0] + wn;
        isWritten[wr] = true;
        writtenRows.push(wr);
      }
    }
    writtenRows.sort(function (a, b) { return a - b; });
    var checkSpans = coalesceRowBlocks(groupContiguousRows(writtenRows), checkWidth);
    for (var w = 0; w < checkSpans.length; w++) {
      var back = sheet.getRange(checkSpans[w][0], 1, checkSpans[w][1], checkWidth).getValues();
      for (var i = 0; i < back.length; i++) {
        var rowNo = checkSpans[w][0] + i;
        if (!isWritten[rowNo]) continue;
        if (String(back[i][acctCol]) !== accountId) {
          throw new Error('upsertRows self-check FAILED on "' + name +
            '" row ' + rowNo + ': account_id="' + back[i][acctCol] +
            '" expected "' + accountId + '" — concurrent-write contamination.');
        }
        if (expectAccountTimezone && tzCol >= 0 &&
            String(back[i][tzCol]) !== String(expectAccountTimezone)) {
          throw new Error('upsertRows self-check FAILED on "' + name +
            '" row ' + rowNo + ': account_timezone="' + back[i][tzCol] +
            '" expected "' + expectAccountTimezone + '".');
        }
      }
    }
  }

  Logger.log(name + ': ' + updates.length + ' updated in place, ' + appends.length +
    ' appended, ' + pruned + ' pruned (< ' + ctx.pruneBefore + ')');
}

// Read the key columns of every data row and classify it:
//   drop — past the retention horizon, unparseable date, or an older duplicate
//          of one of OUR keys (every key set includes account_id, so only our
//          own rows can ever collide)
//   pos  — key → sheet row number, for our account's surviving rows
// Walks backwards so the LAST occurrence of a key is the one kept, matching the
// last-wins rule the downstream sync uses.
// Minimum cells a narrow first pass must save before it is worth the extra
// round trips of the second pass. On the live tabs (2026-09-02) only
// SearchTerms clears it — 94,518 rows x (8 - 3) = 472,590 cells saved, against
// ~12 small range reads for this account's in-window rows. Campaigns (2,015),
// ChangeEvents (5,019) and Keywords (64,860) stay on the single wide read,
// where the second pass would cost more than it saves. Keywords crosses over on
// its own once the tab passes ~40k rows.
var TWO_PASS_MIN_CELLS_SAVED = 200000;

function scanIndex(sheet, indexWidth, dateCol, acctCol, keyIdx, accountId, tz, pruneBefore, dateIsKey, windowStart) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { drop: [], pos: {} };

  // The prefix that carries date + account_id — everything the retention
  // classification and the account filter need. The key columns beyond it are
  // only needed for rows that can actually match an incoming key.
  var narrowWidth = Math.max(dateCol, acctCol) + 1;
  var saved = (lastRow - 1) * (indexWidth - narrowWidth);
  if (dateIsKey && saved >= TWO_PASS_MIN_CELLS_SAVED) {
    return scanIndexTwoPass(sheet, lastRow, indexWidth, narrowWidth, dateCol, acctCol,
      keyIdx, accountId, tz, pruneBefore, windowStart);
  }

  var index = sheet.getRange(2, 1, lastRow - 1, indexWidth).getValues();
  var drop = [];
  var pos = {};
  for (var i = index.length - 1; i >= 0; i--) {
    var rowNo = i + 2;
    var d = toDateStr(index[i][dateCol], tz);
    if (!DATE_RE.test(d) || d < pruneBefore) { drop.push(rowNo); continue; }
    if (String(index[i][acctCol]) !== accountId) continue;
    // Write the NORMALIZED date back before keying. Sheets hands back a Date
    // object for any date cell that predates the '@' pin (and, on a tab built
    // by the previous main.js, for every row past the 1000 it formatted at
    // creation). Keying off the raw value would stringify that as
    // "Fri Aug 07 2026 …", which can never equal the incoming "2026-08-07" —
    // so the row would never be matched and a duplicate would be appended
    // every single hour.
    index[i][dateCol] = d;
    var k = makeKey(index[i], keyIdx);
    if (pos.hasOwnProperty(k)) { drop.push(rowNo); continue; }
    pos[k] = rowNo;
  }
  drop.sort(function (a, b) { return a - b; });
  return { drop: drop, pos: pos };
}

// Same contract as scanIndex, in two reads instead of one wide one.
//
// PASS 1 reads only the date + account_id prefix of the whole tab. That is
// everything the retention classification needs, and everything needed to tell
// our rows from the other ~150 accounts' rows.
//
// PASS 2 reads the full key width for OUR rows only, and only those dated at or
// after `windowStart` (the lowest date we are about to write). Rows older than
// that cannot match an incoming key because `date` is part of every key set, so
// their key columns are dead weight — measured on SearchTerms: 127 rows in ~12
// blocks, against 94,518 rows read at full width before.
//
// DELIBERATE BEHAVIOUR CHANGE: duplicate keys are now garbage-collected only
// within the write window, not across the whole tab. Out-of-window duplicates
// survive until retention ages them out. Accepted because (a) a clean read of
// the live tab on 2026-09-02 found zero duplicates across 94,518 rows, (b) the
// only consumer already dedupes defensively before upserting
// (GoogleAdsSyncStatsCommand, "Defensive dedupe by the Ads-Script's intended
// uniqueness key"), and (c) an in-window duplicate — the kind an interleaved
// write actually produces — is still dropped on the very next run.
function scanIndexTwoPass(sheet, lastRow, indexWidth, narrowWidth, dateCol, acctCol,
                          keyIdx, accountId, tz, pruneBefore, windowStart) {
  var narrow = sheet.getRange(2, 1, lastRow - 1, narrowWidth).getValues();
  var drop = [];
  var mine = [];
  for (var i = 0; i < narrow.length; i++) {
    var rowNo = i + 2;
    var d = toDateStr(narrow[i][dateCol], tz);
    if (!DATE_RE.test(d) || d < pruneBefore) { drop.push(rowNo); continue; }
    if (String(narrow[i][acctCol]) !== accountId) continue;
    // windowStart === null means this tab has nothing to write this run — there
    // is no key to match, so pass 2 is skipped entirely and only the retention
    // classification above survives.
    if (windowStart === null || d < windowStart) continue;
    mine.push(rowNo);
  }

  var pos = {};
  if (mine.length > 0) {
    // Membership set of the rows we own. Coalesced spans deliberately include
    // rows belonging to OTHER accounts, and those must never reach makeKey():
    // a foreign key landing in `pos` is harmless on its own (our incoming keys
    // carry our account_id and could never match it), but a foreign DUPLICATE
    // would push another account's row number into `drop` and delete it.
    var isOurs = {};
    for (var m = 0; m < mine.length; m++) isOurs[mine[m]] = true;

    var spans = coalesceRowBlocks(groupContiguousRows(mine), indexWidth);
    // Walk spans and rows in DESCENDING order so the LAST occurrence of a key
    // is the one kept and earlier ones are dropped — the same last-wins rule
    // the wide scan applies by iterating its array backwards.
    for (var b = spans.length - 1; b >= 0; b--) {
      var vals = sheet.getRange(spans[b][0], 1, spans[b][1], indexWidth).getValues();
      for (var r = vals.length - 1; r >= 0; r--) {
        var rn = spans[b][0] + r;
        if (!isOurs[rn]) continue;
        // Normalize the date before keying, for the same reason the wide scan
        // does: Sheets hands back a Date object for any cell that predates the
        // '@' pin, and its stringification can never equal "2026-08-07".
        vals[r][dateCol] = toDateStr(vals[r][dateCol], tz);
        var k = makeKey(vals[r], keyIdx);
        if (pos.hasOwnProperty(k)) { drop.push(rn); continue; }
        pos[k] = rn;
      }
    }
  }

  drop.sort(function (a, b) { return a - b; });
  return { drop: drop, pos: pos };
}

// Cost of ONE Sheets round trip, expressed in cells. Measured on the live doc
// 2026-09-02 from account 371-906-0364's own log: the SearchTerms phase took
// 33 s for a 283,554-cell narrow scan plus 4 round trips across each of 12
// blocks ⇒ ~68,000 cells/s and ~0.60 s per round trip ⇒ ~40,000 cells.
//
// This is the number that makes block-wise reading pay or not pay, and getting
// it wrong is what made the first cut of the two-pass scan a WASH: it saved
// 472,590 cells (6.9 s) and spent one round trip per block (12 x 0.60 s =
// 7.2 s). Break-even was 11.6 blocks against a fleet average of 12.1.
var SHEET_ROUND_TRIP_CELLS = 40000;

// Merge adjacent row blocks whose gap is cheaper to read than the round trip it
// would save. READ PATHS ONLY — never use this to widen a setValues(), because
// the rows inside a gap belong to OTHER accounts and writing over them is the
// exact cross-account clobber the mutex exists to prevent. Callers must filter
// the rows they actually own back out of the returned spans.
//
// On the measured account this collapses 12 blocks to 2: the gap histogram is
// 1, 1, 2, 2, 13, 10, 5, 3, 938, 1613 and 8362 rows, so everything but the last
// merges for 20,704 extra cells (0.3 s) in place of 10 round trips (6.0 s).
function coalesceRowBlocks(blocks, width) {
  if (blocks.length < 2) return blocks;
  var out = [];
  var cur = [blocks[0][0], blocks[0][1]];
  for (var i = 1; i < blocks.length; i++) {
    var gap = blocks[i][0] - (cur[0] + cur[1]);
    if (gap * width < SHEET_ROUND_TRIP_CELLS) {
      cur[1] = blocks[i][0] + blocks[i][1] - cur[0];
    } else {
      out.push(cur);
      cur = [blocks[i][0], blocks[i][1]];
    }
  }
  out.push(cur);
  return out;
}

// Ascending row numbers → [[startRow, count], …] contiguous blocks.
function groupContiguousRows(rows) {
  var blocks = [];
  var start = rows[0];
  var prev = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === prev + 1) { prev = rows[i]; continue; }
    blocks.push([start, prev - start + 1]);
    start = rows[i];
    prev = rows[i];
  }
  blocks.push([start, prev - start + 1]);
  return blocks;
}

// Delete the given (ascending) row numbers as contiguous blocks, bottom-up so
// earlier row numbers stay valid. deleteRows() moves no data over the wire, so
// this is far cheaper than rewriting the tab to omit the rows.
function deleteRowBlocks(ss, lockUuid, sheet, rows) {
  var blocks = [];
  var start = rows[0];
  var prev = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === prev + 1) { prev = rows[i]; continue; }
    blocks.push([start, prev - start + 1]);
    start = rows[i];
    prev = rows[i];
  }
  blocks.push([start, prev - start + 1]);

  var deleted = 0;
  var issued = 0;
  var b = blocks.length - 1;
  for (; b >= 0; b--) {
    if (issued >= MAX_PRUNE_BLOCKS_PER_RUN) break;
    // Re-assert ownership before EVERY deletion, not once before the loop —
    // a deleteRows is destructive and there can be up to 25 of them.
    refreshSheetLock(ss, lockUuid, false);
    sheet.deleteRows(blocks[b][0], blocks[b][1]);
    deleted += blocks[b][1];
    issued++;
  }
  return { deleted: deleted, skippedBlocks: b + 1 };
}

// [[rowNo, values], …] (sorted) → [{start, rows: [values, …]}, …]
function groupContiguous(updates) {
  var blocks = [];
  var i = 0;
  while (i < updates.length) {
    var start = updates[i][0];
    var rows = [updates[i][1]];
    var j = i;
    while (j + 1 < updates.length && updates[j + 1][0] === updates[j][0] + 1) {
      j++;
      rows.push(updates[j][1]);
    }
    blocks.push({ start: start, rows: rows });
    i = j + 1;
  }
  return blocks;
}

function makeKey(row, keyIdx) {
  var parts = [];
  for (var i = 0; i < keyIdx.length; i++) {
    parts.push(String(row[keyIdx[i]]));
  }
  // SYMBOL FOR UNIT SEPARATOR — won't appear in any real field. Kept as an
  // escape, not a literal: main.js is fetched raw over HTTP and eval()'d, so
  // the separator must not depend on the response's charset being honoured.
  return parts.join("\u241F");
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
