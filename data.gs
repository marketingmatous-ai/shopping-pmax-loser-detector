/**
 * data.gs — Nacitani dat z Google Ads API (GAQL)
 *
 * Vsechny funkce pouzivaji AdsApp.search() s automatickou paginaci.
 * Data jsou vracena v normalizovanem tvaru (ne raw GAQL response).
 *
 * Vsechny cost jsou v hlavni jednotce (NE micros), prevadi se v tomto souboru.
 *
 * Pouziva:
 *  - shopping_performance_view — pro Shopping + PMAX product-level data
 *  - campaign — pro campaign-level stats (CTR baseline per campaign)
 *  - customer — pro account currency
 */

var DataLayer = (function () {
  /**
   * Hlavni query: vraci vsechny produkty v SHOPPING + PMAX kampanich za lookback period.
   * Aplikuje filtering pres brand/rest regex v JS (ne v GAQL — RE2 ma omezenou sadu).
   *
   * @param config CONFIG objekt
   * @returns { products: [...], accountBaseline: {...}, perCampaignBaseline: {...}, excludedCounts: {...} }
   */
  function fetchAllData(config) {
    var endDate = new Date();
    var startDate = Utils.addDays(endDate, -config.lookbackDays);

    Logger.log('INFO: Nacitam data za obdobi ' + Utils.formatDate(startDate) + ' az ' + Utils.formatDate(endDate));

    // === Product-level data ===
    var products = fetchProducts(startDate, endDate, config);

    Logger.log('INFO: Nacteno ' + products.length + ' radku product-level dat (pred filtrem).');

    // === Klasifikuj kampane do MAIN/BRAND/REST bucketu (drive se filtrovalo ven — nyni drzime vse pro split metriky) ===
    var classified = classifyCampaignTypes(products, config);
    var mainOnly = [];
    for (var ci = 0; ci < classified.rows.length; ci++) {
      if (classified.rows[ci].campaignType === 'MAIN') mainOnly.push(classified.rows[ci]);
    }
    Logger.log('INFO: Po klasifikaci campaign type: ' + mainOnly.length + ' MAIN, ' + classified.counts.brand + ' brand, ' + classified.counts.rest + ' rest, ' + classified.counts.paused + ' paused.');

    // === Agregace po item_id se split metrikami (main/brand/rest/total) ===
    var aggregated = aggregateByItemIdSplit(classified.rows);
    Logger.log('INFO: Aggregated ' + aggregated.length + ' unique item_ids with split metrics.');

    // === Account-level baseline (z MAIN produktu) ===
    var accountBaseline = computeAccountBaseline(mainOnly);
    Logger.log('INFO: Account baseline: CPC=' + Utils.formatNumber(accountBaseline.avgCpc, 2) + ', CVR=' + Utils.safePctFormat(accountBaseline.cvr * 100) + ', CTR=' + Utils.safePctFormat(accountBaseline.avgCtr * 100));

    // === Per-campaign CTR baseline (pro volitelny scope, MAIN only) ===
    var perCampaignBaseline = {};
    if (config.ctrBaselineScope === 'campaign') {
      perCampaignBaseline = computePerCampaignBaseline(mainOnly);
    }

    // === Currency z customer API ===
    var currency = fetchAccountCurrency();

    // === First-click dates (pro rising star / age gate) ===
    var firstClickDates = fetchFirstClickDates(aggregated, startDate, config);

    // === Product prices (mapovano z shopping_product resource) ===
    var productPrices = fetchProductPrices(aggregated);
    Logger.log('INFO: Nactene ceny pro ' + Object.keys(productPrices).length + ' produktu z ' + aggregated.length + ' agregovanych.');

    // === Impression shares (z shopping_product_view — shopping_performance_view to nepodporuje) ===
    var impressionShares = fetchImpressionShares(aggregated, startDate, endDate, config);

    // === YoY stats (volitelne) ===
    var lastYearStats = {};
    if (config.enableYoYSeasonalityCheck) {
      lastYearStats = fetchLastYearStats(startDate, endDate, aggregated, config);
    }

    // === Previous period data (pro RISING/DECLINING detekci) ===
    var previousPeriodResult = fetchPreviousPeriodProducts(startDate, endDate, config);
    var previousPeriodRaw = previousPeriodResult.rows;
    var previousPeriodFailed = previousPeriodResult.failed;
    Logger.log('INFO: Previous period: ' + previousPeriodRaw.length + ' rows loaded.');
    var previousPeriodByItemId = aggregatePreviousPeriodByItemId(previousPeriodRaw, config);

    // === Enrich aggregated produkty o previous period reference + impression share ===
    for (var pi = 0; pi < aggregated.length; pi++) {
      var itemKey = aggregated[pi].itemId;
      var itemKeyLower = String(itemKey).toLowerCase();
      var prev = previousPeriodByItemId[itemKeyLower];
      if (prev) {
        aggregated[pi].main_metrics_previous = prev.main_metrics;
        aggregated[pi].brand_metrics_previous = prev.brand_metrics;
        aggregated[pi].rest_metrics_previous = prev.rest_metrics;
        aggregated[pi].total_metrics_previous = prev.total_metrics;
      } else {
        aggregated[pi].main_metrics_previous = null;
        aggregated[pi].brand_metrics_previous = null;
        aggregated[pi].rest_metrics_previous = null;
        aggregated[pi].total_metrics_previous = null;
      }

      if (impressionShares[itemKeyLower] !== undefined) {
        aggregated[pi].searchImpressionShare = impressionShares[itemKeyLower];
      }
    }

    return {
      products: aggregated,
      productsRawRowCount: products.length,       // Raw product-campaign rows z GAQL (pred filtrem)
      productsKeptRowCount: mainOnly.length,      // MAIN rows — odpovida chovani pred split refaktorem
      accountBaseline: accountBaseline,
      perCampaignBaseline: perCampaignBaseline,
      currency: currency,
      firstClickDates: firstClickDates,
      productPrices: productPrices,
      impressionShares: impressionShares,
      lastYearStats: lastYearStats,
      previousPeriodByItemId: previousPeriodByItemId,
      previousPeriodFailed: previousPeriodFailed,
      lookbackStart: startDate,
      lookbackEnd: endDate,
      excludedCounts: classified.counts
    };
  }

  /**
   * Fetchne produkty za predchozi obdobi stejne delky (prev_end = start - 1, prev_start = prev_end - lookbackDays).
   * Pouziva se pro RISING / DECLINING detekci (period-over-period comparison).
   */
  function fetchPreviousPeriodProducts(startDate, endDate, config) {
    var prevEnd = Utils.addDays(startDate, -1);
    var prevStart = Utils.addDays(prevEnd, -config.lookbackDays);

    Logger.log('INFO: Fetching previous period: ' + Utils.formatDate(prevStart) + ' to ' + Utils.formatDate(prevEnd));

    var channelsFilter = config.analyzeChannels.map(function (c) { return "'" + c + "'"; }).join(', ');

    var query =
      'SELECT ' +
      '  segments.product_item_id, ' +
      '  metrics.clicks, ' +
      '  metrics.impressions, ' +
      '  metrics.cost_micros, ' +
      '  metrics.conversions, ' +
      '  metrics.conversions_value, ' +
      '  campaign.name, ' +
      '  campaign.advertising_channel_type, ' +
      '  campaign.status ' +
      'FROM shopping_performance_view ' +
      'WHERE segments.date BETWEEN "' + Utils.formatDate(prevStart) + '" AND "' + Utils.formatDate(prevEnd) + '" ' +
      '  AND campaign.advertising_channel_type IN (' + channelsFilter + ') ' +
      '  AND campaign.status != "PAUSED"';

    var rows = [];
    var failed = false;
    var errorMessage = null;
    try {
      var iterator = AdsApp.search(query);
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.segments.productItemId;
        if (!itemId) continue;
        rows.push({
          itemId: itemId,
          campaignName: (row.campaign && row.campaign.name) ? row.campaign.name : '',
          clicks: Utils.safeParseNumber(row.metrics.clicks, 0),
          impressions: Utils.safeParseNumber(row.metrics.impressions, 0),
          cost: Utils.microsToMajor(Utils.safeParseNumber(row.metrics.costMicros, 0)),
          conversions: Utils.safeParseNumber(row.metrics.conversions, 0),
          conversionValue: Utils.safeParseNumber(row.metrics.conversionsValue, 0)
        });
      }
    } catch (e) {
      failed = true;
      errorMessage = e.message;
      Logger.log('ERROR: Previous period query failed — RISING/DECLINING detection disabled for this run: ' + e.message);
    }

    return { rows: rows, failed: failed, errorMessage: errorMessage };
  }

  /**
   * Agreguje previous-period syrova data do per-item map so split bucketsy.
   * Vraci { itemId: { main_metrics, brand_metrics, rest_metrics, total_metrics } }.
   */
  function aggregatePreviousPeriodByItemId(rows, config) {
    var groups = {};

    function emptyMetrics() {
      return { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0 };
    }

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = String(r.itemId).toLowerCase();

      if (!groups[key]) {
        groups[key] = {
          main_metrics: emptyMetrics(),
          brand_metrics: emptyMetrics(),
          rest_metrics: emptyMetrics()
        };
      }

      var type;
      if (Utils.isBrandCampaign(r.campaignName, config.brandCampaignPattern)) {
        type = 'brand_metrics';
      } else if (Utils.isRestCampaign(r.campaignName, config.restCampaignPattern)) {
        type = 'rest_metrics';
      } else {
        type = 'main_metrics';
      }

      var b = groups[key][type];
      b.clicks += r.clicks;
      b.impressions += r.impressions;
      b.cost += r.cost;
      b.conversions += r.conversions;
      b.conversionValue += r.conversionValue;
    }

    for (var k in groups) {
      if (!groups.hasOwnProperty(k)) continue;
      var g = groups[k];
      g.total_metrics = {
        clicks: g.main_metrics.clicks + g.brand_metrics.clicks + g.rest_metrics.clicks,
        impressions: g.main_metrics.impressions + g.brand_metrics.impressions + g.rest_metrics.impressions,
        cost: g.main_metrics.cost + g.brand_metrics.cost + g.rest_metrics.cost,
        conversions: g.main_metrics.conversions + g.brand_metrics.conversions + g.rest_metrics.conversions,
        conversionValue: g.main_metrics.conversionValue + g.brand_metrics.conversionValue + g.rest_metrics.conversionValue
      };

      // Compute derived metrics on all buckets
      var bucketKeys = ['main_metrics', 'brand_metrics', 'rest_metrics', 'total_metrics'];
      for (var bi = 0; bi < bucketKeys.length; bi++) {
        var b = g[bucketKeys[bi]];
        b.ctr = Utils.safeDiv(b.clicks, b.impressions, 0);
        b.cvr = Utils.safeDiv(b.conversions, b.clicks, 0);
        b.roas = Utils.safeDiv(b.conversionValue, b.cost, 0);
        b.avgCpc = Utils.safeDiv(b.cost, b.clicks, 0);
        b.pno = b.conversionValue > 0 ? (b.cost / b.conversionValue * 100) : 0;
      }
    }

    return groups;
  }

  /**
   * Fetchne search_impression_share per produkt ze shopping_product_view.
   * shopping_performance_view tuto metriku nenabizi — musime doplnit separatni query.
   * Vraci mapu { lowercase_item_id: impression_share_decimal }.
   */
  function fetchImpressionShares(products, startDate, endDate, config) {
    var result = {};
    if (!products || products.length === 0) return result;

    // GOTCHA: metrics.search_impression_share NELZE kombinovat s impressions/clicks/cost
    // v jednom SELECTu (Google Ads API to odmitne s "undefined" error).
    // Proto je IS v separatni query, ne jako soucast fetchProducts.
    //
    // Resource: shopping_performance_view (NE shopping_product_view — ten neexistuje).
    var channelsFilter = config.analyzeChannels.map(function (c) { return "'" + c + "'"; }).join(', ');

    try {
      var query =
        'SELECT ' +
        '  segments.product_item_id, ' +
        '  metrics.search_impression_share, ' +
        '  campaign.advertising_channel_type ' +
        'FROM shopping_performance_view ' +
        'WHERE segments.date BETWEEN "' + Utils.formatDate(startDate) + '" AND "' + Utils.formatDate(endDate) + '" ' +
        '  AND campaign.advertising_channel_type IN (' + channelsFilter + ')';

      var iterator = AdsApp.search(query);
      var count = 0;
      // Agregace per item_id: bereme MAX z kampani (konzervativni — pokud produkt ma
      // nekde vysoky IS, neni "lost opportunity"). Aritmeticky prumer zkresluje u
      // produktu s velkymi variancemi mezi kampanemi (1 kampan IS=0.9, 1 kampan IS=0.1).
      // Impression-weighted average by byl idealni, ale IS nelze kombinovat s impressions
      // v jednom SELECTu GAQL — MAX je pragmaticky kompromis.
      var tmpMax = {}; // item_id → maxIS
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.segments && row.segments.productItemId ? row.segments.productItemId : null;
        if (!itemId) continue;
        var is = Utils.safeParseNumber(row.metrics && row.metrics.searchImpressionShare, 0);
        if (is <= 0) continue; // skip zero-IS radky (nerelevantni)
        var key = String(itemId).toLowerCase();
        if (tmpMax[key] === undefined || is > tmpMax[key]) {
          tmpMax[key] = is;
        }
      }
      var keys = Object.keys(tmpMax);
      for (var k = 0; k < keys.length; k++) {
        result[keys[k]] = tmpMax[keys[k]];
        count++;
      }
      Logger.log('INFO: Loaded search_impression_share for ' + count + ' products (MAX per item_id) from shopping_performance_view.');
    } catch (e) {
      Logger.log('WARN: Impression share fetch failed: ' + e.message + ' — LOST_OPPORTUNITY detection will skip.');
    }

    return result;
  }

  /**
   * Mapuje item_id na product_price (z shopping_product resource).
   *
   * shopping_performance_view NEMA cenu produktu — ta je v shopping_product.
   * item_id v shopping_product je UPPERCASE (napr. "NB 2414 KO"),
   * v shopping_performance_view je lowercase ("nb 2414 ko").
   * Mapujeme case-insensitive.
   *
   * @param products pole aggregovanych produktu
   * @returns mapa { lowercase_item_id: price_in_currency }
   */
  function fetchProductPrices(products) {
    var result = {};
    if (!products || products.length === 0) {
      return result;
    }
    try {
      // Query BEZ WHERE klauzule — shopping_product.status je enum, porovnani s literalem
      // UNSPECIFIED drive (bez uvozovek) mohlo selhat a vratit 0 radku. Mistto filtrovani
      // v GAQL pullujeme vsechno a filtrujeme v JS (price_micros > 0 = aktivni produkt).
      var query =
        'SELECT shopping_product.item_id, shopping_product.price_micros, shopping_product.currency_code ' +
        'FROM shopping_product';
      var iterator = AdsApp.search(query);
      var fetchedCount = 0;
      var skippedNoPrice = 0;
      var samplePairs = []; // prvnich 5 item_id:price pro diagnostiku
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.shoppingProduct && row.shoppingProduct.itemId ? row.shoppingProduct.itemId : null;
        if (!itemId) continue;
        var priceMicros = row.shoppingProduct.priceMicros ? Number(row.shoppingProduct.priceMicros) : 0;
        if (priceMicros <= 0) {
          skippedNoPrice++;
          continue;
        }
        // Klic = lowercase pro match s performance view (item_id v shopping_product je UPPERCASE)
        var lowerKey = String(itemId).toLowerCase();
        result[lowerKey] = priceMicros / 1000000;
        fetchedCount++;
        if (samplePairs.length < 5) {
          samplePairs.push(lowerKey + '=' + (priceMicros / 1000000).toFixed(0));
        }
      }
      Logger.log('INFO: fetchProductPrices — nacteno ' + fetchedCount + ' cen (skipped no_price: ' + skippedNoPrice + '). Sample: ' + samplePairs.join(', '));

      // Diagnostika: kolik z agreg. produktu ma cenu v feedu?
      var matched = 0;
      var unmatched = 0;
      var unmatchedSamples = [];
      for (var p = 0; p < products.length; p++) {
        var pid = String(products[p].itemId || '').toLowerCase();
        if (result[pid]) {
          matched++;
        } else {
          unmatched++;
          if (unmatchedSamples.length < 5) {
            unmatchedSamples.push(pid);
          }
        }
      }
      Logger.log('INFO: fetchProductPrices coverage — matched ' + matched + '/' + products.length + ' produktu. Unmatched sample: ' + unmatchedSamples.join(', '));
    } catch (e) {
      Logger.log('WARN: fetchProductPrices selhal: ' + e.message + ' — produkty bez derived ceny budou skip-nuty.');
    }
    return result;
  }

  /**
   * GAQL query pro product stats. Pouziva shopping_performance_view.
   * PMAX produkty jsou v tomto view dostupne pres filter advertising_channel_type.
   */
  function fetchProducts(startDate, endDate, config) {
    var channelsFilter = config.analyzeChannels.map(function (c) { return "'" + c + "'"; }).join(', ');

    // Pozn.: metrics.search_impression_share NENI v shopping_performance_view
    // (je jen v campaign/ad_group/keyword views). Bez ni ztraci low_ctr
    // sub-klasifikaci (irrelevant_keyword_match vs high_visibility_low_appeal),
    // ale zakladni low_ctr_audit flag stale funguje.
    var query =
      'SELECT ' +
      '  segments.product_item_id, ' +
      '  segments.product_title, ' +
      '  segments.product_brand, ' +
      '  segments.product_type_l1, ' +
      '  segments.product_custom_attribute0, ' +
      '  segments.product_custom_attribute1, ' +
      '  segments.product_custom_attribute2, ' +
      '  segments.product_custom_attribute3, ' +
      '  segments.product_custom_attribute4, ' +
      '  metrics.clicks, ' +
      '  metrics.impressions, ' +
      '  metrics.cost_micros, ' +
      '  metrics.conversions, ' +
      '  metrics.conversions_value, ' +
      '  campaign.id, ' +
      '  campaign.name, ' +
      '  campaign.status, ' +
      '  campaign.advertising_channel_type ' +
      'FROM shopping_performance_view ' +
      'WHERE segments.date BETWEEN "' + Utils.formatDate(startDate) + '" AND "' + Utils.formatDate(endDate) + '" ' +
      '  AND campaign.advertising_channel_type IN (' + channelsFilter + ')';

    var rows = [];
    var iterator;
    try {
      iterator = AdsApp.search(query);
    } catch (e) {
      Logger.log('ERROR: GAQL query failed: ' + e.message);
      throw new Error('GAQL query selhal: ' + e.message);
    }

    while (iterator.hasNext()) {
      var row = iterator.next();
      var parsed = parseRow(row);
      if (parsed) {
        rows.push(parsed);
      }
    }

    return rows;
  }

  /**
   * Parsuje jeden GAQL row do normalizovane struktury.
   */
  function parseRow(row) {
    var itemId = (row.segments && row.segments.productItemId) ? row.segments.productItemId : null;
    if (!itemId) {
      return null; // Radek bez item_id preskoc
    }

    var cost = Utils.microsToMajor(Utils.safeParseNumber(row.metrics.costMicros, 0));
    var clicks = Utils.safeParseNumber(row.metrics.clicks, 0);
    var impressions = Utils.safeParseNumber(row.metrics.impressions, 0);
    var conversions = Utils.safeParseNumber(row.metrics.conversions, 0);
    var conversionValue = Utils.safeParseNumber(row.metrics.conversionsValue, 0);

    return {
      itemId: itemId,
      productTitle: row.segments.productTitle || '',
      productBrand: row.segments.productBrand || '',
      productType: row.segments.productTypeL1 || '',
      customLabel0: row.segments.productCustomAttribute0 || '',
      customLabel1: row.segments.productCustomAttribute1 || '',
      customLabel2: row.segments.productCustomAttribute2 || '',
      customLabel3: row.segments.productCustomAttribute3 || '',
      customLabel4: row.segments.productCustomAttribute4 || '',
      clicks: clicks,
      impressions: impressions,
      cost: cost,
      conversions: conversions,
      conversionValue: conversionValue,
      searchImpressionShare: 0, // Neni v shopping_performance_view — fallback 0
      campaignId: row.campaign.id || '',
      campaignName: row.campaign.name || '',
      campaignStatus: row.campaign.status || '',
      channel: row.campaign.advertisingChannelType || ''
    };
  }

  /**
   * Klasifikuje produktove radky dle typu kampane (MAIN / BRAND / REST).
   * Paused kampane se vynechavaji uplne. Vraci { rows, counts }.
   * Oproti puvodnimu filterCampaigns drzi vsechny non-paused radky (potrebne
   * pro split metriky — main/brand/rest, ne jen main filter).
   */
  function classifyCampaignTypes(products, config) {
    var result = [];
    var counts = { main: 0, brand: 0, rest: 0, paused: 0 };

    for (var i = 0; i < products.length; i++) {
      var p = products[i];
      if (p.campaignStatus === 'PAUSED') {
        counts.paused++;
        continue;
      }
      if (Utils.isBrandCampaign(p.campaignName, config.brandCampaignPattern)) {
        p.campaignType = 'BRAND';
        counts.brand++;
      } else if (Utils.isRestCampaign(p.campaignName, config.restCampaignPattern)) {
        p.campaignType = 'REST';
        counts.rest++;
      } else {
        p.campaignType = 'MAIN';
        counts.main++;
      }
      result.push(p);
    }

    return { rows: result, counts: counts };
  }

  /**
   * Backward-compat wrapper. Vraci { kept, excluded } jen pro MAIN kampane.
   */
  function filterCampaigns(products, config) {
    var r = classifyCampaignTypes(products, config);
    var kept = [];
    for (var i = 0; i < r.rows.length; i++) {
      if (r.rows[i].campaignType === 'MAIN') kept.push(r.rows[i]);
    }
    return { kept: kept, excluded: r.counts };
  }

  /**
   * Agreguje radky (product × campaign) do per-item_id agregatu.
   * Metriky se scitaji; pripadne dalsi fields (campaign name) se uchovavaji
   * jako "primary" podle nejvetsi cost.
   */
  function aggregateByItemId(rows) {
    var groups = {};

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = r.itemId;

      if (!groups[key]) {
        groups[key] = {
          itemId: r.itemId,
          productTitle: r.productTitle,
          productBrand: r.productBrand,
          productType: r.productType,
          customLabel0: r.customLabel0,
          customLabel1: r.customLabel1,
          customLabel2: r.customLabel2,
          customLabel3: r.customLabel3,
          customLabel4: r.customLabel4,
          clicks: 0,
          impressions: 0,
          cost: 0,
          conversions: 0,
          conversionValue: 0,
          searchImpressionShareSum: 0,
          searchImpressionShareWeight: 0,
          primaryCampaignId: r.campaignId,
          primaryCampaignName: r.campaignName,
          primaryCampaignCost: r.cost,
          channel: r.channel,
          campaigns: []
        };
      }

      var g = groups[key];
      g.clicks += r.clicks;
      g.impressions += r.impressions;
      g.cost += r.cost;
      g.conversions += r.conversions;
      g.conversionValue += r.conversionValue;

      // Weighted average pro search_impression_share (weight = impressions)
      if (r.searchImpressionShare > 0 && r.impressions > 0) {
        g.searchImpressionShareSum += r.searchImpressionShare * r.impressions;
        g.searchImpressionShareWeight += r.impressions;
      }

      // Primary campaign = ta s nejvetsi cost
      if (r.cost > g.primaryCampaignCost) {
        g.primaryCampaignId = r.campaignId;
        g.primaryCampaignName = r.campaignName;
        g.primaryCampaignCost = r.cost;
        g.channel = r.channel;
      }

      g.campaigns.push({
        id: r.campaignId,
        name: r.campaignName,
        cost: r.cost,
        clicks: r.clicks,
        conversions: r.conversions
      });
    }

    // Finalizuj weighted avg
    var result = [];
    for (var k in groups) {
      if (groups.hasOwnProperty(k)) {
        var item = groups[k];
        item.searchImpressionShare = Utils.safeDiv(item.searchImpressionShareSum, item.searchImpressionShareWeight, 0);
        delete item.searchImpressionShareSum;
        delete item.searchImpressionShareWeight;
        delete item.primaryCampaignCost;

        // Pocitane derived metriky
        item.ctr = Utils.safeDiv(item.clicks, item.impressions, 0);
        item.cvr = Utils.safeDiv(item.conversions, item.clicks, 0);
        item.roas = Utils.safeDiv(item.conversionValue, item.cost, 0);
        item.avgCpc = Utils.safeDiv(item.cost, item.clicks, 0);
        item.actualPno = Utils.safeDiv(item.cost, item.conversionValue, 0) * 100;

        result.push(item);
      }
    }

    return result;
  }

  /**
   * Agreguje radky (product x campaign) do per-item_id agregatu se split metrikami.
   * Kazdy produkt ma {main_metrics, brand_metrics, rest_metrics, total_metrics}
   * + backward-compat mirror MAIN metrik na top-level (clicks, cost, ctr...) tak,
   * aby existujici classifier kod fungoval dal.
   */
  function aggregateByItemIdSplit(rows) {
    var groups = {};

    function emptyMetrics() {
      return { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0 };
    }

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = r.itemId;

      if (!groups[key]) {
        groups[key] = {
          itemId: r.itemId,
          productTitle: r.productTitle,
          productBrand: r.productBrand,
          productType: r.productType,
          customLabel0: r.customLabel0,
          customLabel1: r.customLabel1,
          customLabel2: r.customLabel2,
          customLabel3: r.customLabel3,
          customLabel4: r.customLabel4,
          main_metrics: emptyMetrics(),
          brand_metrics: emptyMetrics(),
          rest_metrics: emptyMetrics(),
          primaryCampaignName: r.campaignName,
          primaryCampaignId: r.campaignId,
          primaryCampaignCost: r.cost,
          channel: r.channel,
          campaigns: []
        };
      }

      var g = groups[key];
      var bucketKey = String(r.campaignType).toLowerCase() + '_metrics';
      var bucket = g[bucketKey];
      if (!bucket) {
        Logger.log('WARN: aggregateByItemIdSplit — unknown campaignType "' + r.campaignType + '" for item ' + r.itemId + ', row skipped.');
        continue;
      }

      bucket.clicks += r.clicks;
      bucket.impressions += r.impressions;
      bucket.cost += r.cost;
      bucket.conversions += r.conversions;
      bucket.conversionValue += r.conversionValue;

      if (r.cost > g.primaryCampaignCost) {
        g.primaryCampaignId = r.campaignId;
        g.primaryCampaignName = r.campaignName;
        g.primaryCampaignCost = r.cost;
        g.channel = r.channel;
      }

      g.campaigns.push({
        id: r.campaignId, name: r.campaignName, type: r.campaignType,
        cost: r.cost, clicks: r.clicks, conversions: r.conversions
      });
    }

    var result = [];
    var bucketKeys = ['main_metrics', 'brand_metrics', 'rest_metrics'];
    for (var k in groups) {
      if (!groups.hasOwnProperty(k)) continue;
      var item = groups[k];

      item.total_metrics = {
        clicks: item.main_metrics.clicks + item.brand_metrics.clicks + item.rest_metrics.clicks,
        impressions: item.main_metrics.impressions + item.brand_metrics.impressions + item.rest_metrics.impressions,
        cost: item.main_metrics.cost + item.brand_metrics.cost + item.rest_metrics.cost,
        conversions: item.main_metrics.conversions + item.brand_metrics.conversions + item.rest_metrics.conversions,
        conversionValue: item.main_metrics.conversionValue + item.brand_metrics.conversionValue + item.rest_metrics.conversionValue
      };

      var allBuckets = bucketKeys.concat(['total_metrics']);
      for (var bi = 0; bi < allBuckets.length; bi++) {
        var b = item[allBuckets[bi]];
        b.ctr = Utils.safeDiv(b.clicks, b.impressions, 0);
        b.cvr = Utils.safeDiv(b.conversions, b.clicks, 0);
        b.roas = Utils.safeDiv(b.conversionValue, b.cost, 0);
        b.avgCpc = Utils.safeDiv(b.cost, b.clicks, 0);
        b.pno = b.conversionValue > 0 ? (b.cost / b.conversionValue * 100) : 0;
      }

      // Backward compat: expose MAIN metrics at top level (so existing classifier code keeps working)
      item.clicks = item.main_metrics.clicks;
      item.impressions = item.main_metrics.impressions;
      item.cost = item.main_metrics.cost;
      item.conversions = item.main_metrics.conversions;
      item.conversionValue = item.main_metrics.conversionValue;
      item.ctr = item.main_metrics.ctr;
      item.cvr = item.main_metrics.cvr;
      item.roas = item.main_metrics.roas;
      item.avgCpc = item.main_metrics.avgCpc;
      item.actualPno = item.main_metrics.pno;
      item.searchImpressionShare = 0; // Populated later from separate query

      // Multi-campaign info (user dotaz: "co kdyz je produkt ve vice kampanich?")
      // Primary campaign = ta s nejvyssim cost. Tady shrnujeme kolik celkem kampani
      // produkt mel a kolik % cost tvori primary kampan.
      var uniqueCampaignNames = {};
      for (var ci = 0; ci < item.campaigns.length; ci++) {
        uniqueCampaignNames[item.campaigns[ci].name] = true;
      }
      item.campaignsCount = Object.keys(uniqueCampaignNames).length;
      item.primaryCampaignSharePct = item.total_metrics.cost > 0
        ? (item.primaryCampaignCost / item.total_metrics.cost * 100)
        : 0;
      // Top 3 kampani dle cost (pro rychly overview v DETAIL tabu)
      var sortedCampaigns = item.campaigns.slice().sort(function (a, b) {
        return b.cost - a.cost;
      });
      var topNames = [];
      var seen = {};
      for (var si = 0; si < sortedCampaigns.length && topNames.length < 3; si++) {
        var nm = sortedCampaigns[si].name;
        if (!seen[nm]) {
          seen[nm] = true;
          topNames.push(nm);
        }
      }
      item.topCampaigns = topNames.join(' | ');

      result.push(item);
    }

    return result;
  }

  /**
   * Spocita account-level baseline napric vsemi non-brand non-rest non-paused produkty.
   */
  function computeAccountBaseline(rows) {
    var total = {
      clicks: 0,
      impressions: 0,
      cost: 0,
      conversions: 0,
      conversionValue: 0
    };

    for (var i = 0; i < rows.length; i++) {
      total.clicks += rows[i].clicks;
      total.impressions += rows[i].impressions;
      total.cost += rows[i].cost;
      total.conversions += rows[i].conversions;
      total.conversionValue += rows[i].conversionValue;
    }

    return {
      totalClicks: total.clicks,
      totalImpressions: total.impressions,
      totalCost: total.cost,
      totalConversions: total.conversions,
      totalConversionValue: total.conversionValue,
      avgCpc: Utils.safeDiv(total.cost, total.clicks, 0),
      cvr: Utils.safeDiv(total.conversions, total.clicks, 0),
      avgCtr: Utils.safeDiv(total.clicks, total.impressions, 0),
      avgRoas: Utils.safeDiv(total.conversionValue, total.cost, 0),
      avgPno: total.conversionValue > 0 ? (total.cost / total.conversionValue * 100) : 0,
      avgAov: Utils.safeDiv(total.conversionValue, total.conversions, 0),
      rowCount: rows.length
    };
  }

  /**
   * Per-campaign CTR baseline. Pouziva se pokud config.ctrBaselineScope === 'campaign'.
   */
  function computePerCampaignBaseline(rows) {
    var groups = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var key = r.campaignId;
      if (!groups[key]) {
        groups[key] = { clicks: 0, impressions: 0, campaignName: r.campaignName };
      }
      groups[key].clicks += r.clicks;
      groups[key].impressions += r.impressions;
    }

    var result = {};
    for (var k in groups) {
      if (groups.hasOwnProperty(k)) {
        result[k] = {
          campaignName: groups[k].campaignName,
          ctr: Utils.safeDiv(groups[k].clicks, groups[k].impressions, 0),
          impressions: groups[k].impressions,
          clicks: groups[k].clicks
        };
      }
    }

    return result;
  }

  /**
   * Currency kodu z customer API.
   */
  function fetchAccountCurrency() {
    try {
      var query = 'SELECT customer.currency_code FROM customer';
      var iterator = AdsApp.search(query);
      if (iterator.hasNext()) {
        var row = iterator.next();
        return row.customer.currencyCode || 'CZK';
      }
    } catch (e) {
      Logger.log('WARN: Nepodarilo se zjistit currency: ' + e.message);
    }
    return 'CZK';
  }

  /**
   * First-click date per item_id. Pouziva se pro age gate (rising star).
   * Pouziva GAQL nad delsim obdobim (2× lookback) pro heuristiku.
   *
   * DULEZITE: Filtr musi odpovidat hlavnimu query (shopping_performance_view pro
   * analyzu), jinak by brand/rest/paused kampan data mohly nastavit drivejsi
   * first-click datum u produktu, ktery je do non-brand Shopping kampan novy.
   */
  function fetchFirstClickDates(products, currentStartDate, config) {
    var result = {};
    if (!products || products.length === 0) {
      return result;
    }

    // Rozsah: 2× lookback
    var extendedStart = Utils.addDays(currentStartDate, -config.lookbackDays);
    var endDate = new Date();

    var channelsFilter = config.analyzeChannels.map(function (c) { return "'" + c + "'"; }).join(', ');

    var query =
      'SELECT ' +
      '  segments.product_item_id, ' +
      '  segments.date, ' +
      '  metrics.clicks, ' +
      '  campaign.name, ' +
      '  campaign.advertising_channel_type, ' +
      '  campaign.status ' +
      'FROM shopping_performance_view ' +
      'WHERE segments.date BETWEEN "' + Utils.formatDate(extendedStart) + '" AND "' + Utils.formatDate(endDate) + '" ' +
      '  AND metrics.clicks > 0 ' +
      '  AND campaign.advertising_channel_type IN (' + channelsFilter + ') ' +
      '  AND campaign.status != "PAUSED"';
    // Poznamka: bez ORDER BY — sortujeme v JS (YYYY-MM-DD lexicographic sort).
    // ORDER BY na velke datasety je pomaly a muze prekrocit 30-min limit.

    try {
      var iterator = AdsApp.search(query);
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.segments.productItemId;
        if (!itemId) {
          continue;
        }
        var campaignName = (row.campaign && row.campaign.name) ? row.campaign.name : '';
        // Vyfiltruj brand + rest kampane (REGEXP_MATCH v GAQL ma omezenou RE2 syntax,
        // konzistenci s filterCampaigns dosahujeme stejnymi regexy v JS)
        if (Utils.isBrandCampaign(campaignName, config.brandCampaignPattern)) {
          continue;
        }
        if (Utils.isRestCampaign(campaignName, config.restCampaignPattern)) {
          continue;
        }
        var dateStr = row.segments.date;
        if (!result[itemId] || dateStr < result[itemId]) {
          // YYYY-MM-DD lexicographic comparison = chronological
          result[itemId] = dateStr;
        }
      }
    } catch (e) {
      Logger.log('WARN: First-click date query selhal: ' + e.message + ' — produkty budou mit age=null');
    }

    return result;
  }

  /**
   * YoY stats — porovnani se stejnym obdobim pred rokem.
   */
  function fetchLastYearStats(currentStartDate, currentEndDate, products, config) {
    var result = {};
    if (!products || products.length === 0) {
      return result;
    }

    var lastYearStart = Utils.addDays(currentStartDate, -365);
    var lastYearEnd = Utils.addDays(currentEndDate, -365);

    var channelsFilter = config.analyzeChannels.map(function (c) { return "'" + c + "'"; }).join(', ');

    var query =
      'SELECT ' +
      '  segments.product_item_id, ' +
      '  metrics.clicks, ' +
      '  metrics.impressions, ' +
      '  metrics.cost_micros, ' +
      '  metrics.conversions, ' +
      '  metrics.conversions_value, ' +
      '  campaign.name, ' +
      '  campaign.advertising_channel_type ' +
      'FROM shopping_performance_view ' +
      'WHERE segments.date BETWEEN "' + Utils.formatDate(lastYearStart) + '" AND "' + Utils.formatDate(lastYearEnd) + '" ' +
      '  AND campaign.advertising_channel_type IN (' + channelsFilter + ')';

    try {
      var iterator = AdsApp.search(query);
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.segments.productItemId;
        if (!itemId) {
          continue;
        }
        // Vyfiltruj brand + rest kampane (konzistence s hlavni analyzou)
        var campaignName = (row.campaign && row.campaign.name) ? row.campaign.name : '';
        if (Utils.isBrandCampaign(campaignName, config.brandCampaignPattern)) {
          continue;
        }
        if (Utils.isRestCampaign(campaignName, config.restCampaignPattern)) {
          continue;
        }
        if (!result[itemId]) {
          result[itemId] = { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0 };
        }
        result[itemId].clicks += Utils.safeParseNumber(row.metrics.clicks, 0);
        result[itemId].impressions += Utils.safeParseNumber(row.metrics.impressions, 0);
        result[itemId].cost += Utils.microsToMajor(Utils.safeParseNumber(row.metrics.costMicros, 0));
        result[itemId].conversions += Utils.safeParseNumber(row.metrics.conversions, 0);
        result[itemId].conversionValue += Utils.safeParseNumber(row.metrics.conversionsValue, 0);
      }
      Logger.log('INFO: YoY data: ' + Object.keys(result).length + ' produktu melo data pred rokem.');
    } catch (e) {
      Logger.log('WARN: YoY query selhal (asi account < 1 rok): ' + e.message);
    }

    return result;
  }

  /**
   * Vraci mapu item_id → last lifecycle entry (z LIFECYCLE_LOG tabu).
   * Pouziva se pro detectTransition v classifier.
   */
  function fetchPreviousLifecycle(outputSheetId) {
    var result = {};
    try {
      var ss = SpreadsheetApp.openById(outputSheetId);
      var sheet = ss.getSheetByName('LIFECYCLE_LOG');
      if (!sheet) {
        return result;
      }
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return result;
      }

      // Detekce layoutu:
      //   - Novy s dashboardem: header na row 9, data od row 10
      //   - Stary bez dashboardu: header na row 1, data od row 2
      var r9 = String(sheet.getRange(9, 1).getValue()).replace(/^\s+|\s+$/g, '');
      var r1 = String(sheet.getRange(1, 1).getValue()).replace(/^\s+|\s+$/g, '');
      var headerRow, dataStartRow;
      if (r9 === 'run_date') {
        headerRow = 9;
        dataStartRow = 10;
      } else if (r1 === 'run_date') {
        headerRow = 1;
        dataStartRow = 2;
      } else {
        Logger.log('WARN: fetchPreviousLifecycle — nelze detekovat layout, history prazdna.');
        return result;
      }

      if (lastRow < dataStartRow) {
        return result;
      }

      var lastCol = sheet.getLastColumn();
      var data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
      var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

      var idxItemId = headers.indexOf('item_id');
      var idxRunDate = headers.indexOf('run_date');
      var idxLabel = headers.indexOf('current_label');
      var idxCampaign = headers.indexOf('current_campaign');
      var idxRuns = headers.indexOf('runs_since_first_flag');

      if (idxItemId < 0 || idxRunDate < 0) {
        Logger.log('WARN: fetchPreviousLifecycle — nelze najit kolonky item_id nebo run_date.');
        return result;
      }

      // Najdi posledni radek per item_id — porovnani podle NORMALIZED date stringu
      // (Sheets muze vratit Date objekt nebo locale-formatted string)
      for (var i = 0; i < data.length; i++) {
        var itemId = data[i][idxItemId];
        if (!itemId) {
          continue;
        }
        var rawDate = data[i][idxRunDate];
        var dateStr = Utils.normalizeDate(rawDate);
        if (!result[itemId] || dateStr > result[itemId].runDate) {
          result[itemId] = {
            runDate: dateStr, // uz normalized ISO
            label: data[i][idxLabel] || '',
            campaign: data[i][idxCampaign] || '',
            runsSinceFirstFlag: idxRuns >= 0 ? Utils.safeParseNumber(data[i][idxRuns], 0) : 0
          };
        }
      }
    } catch (e) {
      Logger.log('WARN: Nacitani LIFECYCLE_LOG historie selhalo: ' + e.message);
    }
    return result;
  }

  /**
   * Read existing ACTIONS tab and extract manual columns per item_id.
   * Used for preserving manual input (action_taken, action_date, consultant_note)
   * between runs. Returns empty object if tab doesn't exist.
   *
   * Returns: { item_id: { action_taken, action_date, consultant_note } }
   */
  function readExistingActions(outputSheetId) {
    var result = {};
    try {
      var ss = SpreadsheetApp.openById(outputSheetId);
      var sheet = ss.getSheetByName('ACTIONS');
      if (!sheet || sheet.getLastRow() < 2) return result;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();

      // ACTIONS tab ma dashboard panel nahore — header tabulky je na dynamic row
      // (zavisi na poctu insights v panelu). Scanner najde row s 'priority_rank' v col 1.
      var scanLimit = Math.min(lastRow, 60);
      var scanValues = sheet.getRange(1, 1, scanLimit, 1).getValues();
      var headerRow = -1;
      for (var si = 0; si < scanValues.length; si++) {
        var cellVal = String(scanValues[si][0] || '').replace(/^\s+|\s+$/g, '');
        if (cellVal === 'priority_rank') {
          headerRow = si + 1; // 1-indexed
          break;
        }
      }
      if (headerRow < 0) {
        Logger.log('WARN: ACTIONS tab — nelze najit header row (priority_rank). Skip preserve.');
        return result;
      }
      if (lastRow <= headerRow) {
        return result; // jen header, zadna data
      }

      var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
      var data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

      var idxItemId = headers.indexOf('item_id');
      var idxAction = headers.indexOf('action_taken');
      var idxDate = headers.indexOf('action_date');
      var idxNote = headers.indexOf('consultant_note');

      if (idxItemId < 0) {
        Logger.log('WARN: ACTIONS tab missing item_id column — skip preserve.');
        return result;
      }

      for (var i = 0; i < data.length; i++) {
        var itemId = data[i][idxItemId];
        if (!itemId) continue;
        // Normalize action_date (Date objekt → ISO string, jinak normalize)
        var actionDateRaw = idxDate >= 0 ? data[i][idxDate] : '';
        var actionDate = '';
        if (actionDateRaw instanceof Date) {
          actionDate = Utils.formatDate(actionDateRaw);
        } else if (actionDateRaw) {
          var normalized = Utils.normalizeDate(actionDateRaw);
          actionDate = normalized || String(actionDateRaw);
        }
        result[itemId] = {
          action_taken: idxAction >= 0 ? (data[i][idxAction] || '') : '',
          action_date: actionDate,
          consultant_note: idxNote >= 0 ? (data[i][idxNote] || '') : ''
        };
      }
      Logger.log('INFO: Read ' + Object.keys(result).length + ' ACTIONS rows for manual preserve.');
    } catch (e) {
      Logger.log('WARN: readExistingActions failed: ' + e.message);
    }
    return result;
  }

  /**
   * Read existing PRODUCT_TIMELINE tab and return full rows keyed by item_id.
   * Used for:
   *   - Preserving first_flag_date (captured once)
   *   - Preserving categories_history (chronological chain)
   *   - Preserving manual columns (latest_action, latest_action_date, latest_note)
   *
   * Returns: { item_id: { ...all columns as object with header-keyed values... } }
   */
  function readExistingProductTimeline(outputSheetId) {
    var result = {};
    try {
      var ss = SpreadsheetApp.openById(outputSheetId);
      var sheet = ss.getSheetByName('PRODUCT_TIMELINE');
      if (!sheet || sheet.getLastRow() < 2) return result;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

      // Normalize datum field names — pokud je hodnota Date objekt nebo locale string,
      // prevedeme na ISO YYYY-MM-DD pro konzistentni porovnavani mezi runy.
      var dateFields = { 'first_flag_date': true, 'latest_action_date': true, 'kpi_before_date': true };

      for (var i = 0; i < data.length; i++) {
        var rowObj = {};
        for (var j = 0; j < headers.length; j++) {
          var colName = headers[j];
          var val = data[i][j];
          if (dateFields[colName] && val) {
            val = Utils.normalizeDate(val) || '';
          }
          rowObj[colName] = val;
        }
        var itemId = rowObj.item_id;
        if (!itemId) continue;
        result[itemId] = rowObj;
      }
      Logger.log('INFO: Read ' + Object.keys(result).length + ' PRODUCT_TIMELINE rows.');
    } catch (e) {
      Logger.log('WARN: readExistingProductTimeline failed: ' + e.message);
    }
    return result;
  }

  return {
    fetchAllData: fetchAllData,
    fetchPreviousLifecycle: fetchPreviousLifecycle,
    readExistingActions: readExistingActions,
    readExistingProductTimeline: readExistingProductTimeline,
    classifyCampaignTypes: classifyCampaignTypes,
    aggregateByItemIdSplit: aggregateByItemIdSplit
  };
})();
