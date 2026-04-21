# Shopping/PMAX Loser Detector v2 — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Evolve v1 detector (labeling tool) into v2 insights + campaign optimization tool with 3 new categories (RISING, DECLINING, LOST_OPPORTUNITY), split scope metrics (main vs brand vs rest), per-product history tracking, semi-manual intervention logging, and 9-tab sheet with DASHBOARD trends.

**Architecture:** Google Ads Script (ES5 JavaScript), single-account deploy. Data from `shopping_performance_view` (classification) + `shopping_product` (prices) + `shopping_product_view` (impression share). Output to Google Sheet via `SpreadsheetApp`. 9 tabs: DASHBOARD, FEED_UPLOAD, ACTIONS, PRODUCT_TIMELINE, DETAIL, LIFECYCLE_LOG, WEEKLY_SNAPSHOT, CONFIG, README.

**Tech Stack:** Google Ads Scripts API (JS ES5), GAQL queries, SpreadsheetApp, Node.js (unit test runner for pure classifier functions), Python (MCP data validation).

**Design reference:** `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/docs/2026-04-21-v2-redesign-design.md`

---

## Testing strategy

- **Pure functions** (classifier logic, utils) → unit tests via Node.js runner in `/tmp/loser-test/` (pattern from v1, 46 existing tests). Add ~25 new tests for RISING/DECLINING/LOST_OPPORTUNITY and preserve mechanism.
- **GAS-specific functions** (data fetching, sheet writing) → manual dry-run in Google Ads UI + MCP query comparison for data accuracy.
- **End-to-end** → 2 dry-runs on kabelka.cz account (empty sheet + filled sheet with manual input).

## Commit strategy

Projekt není v git repo. Místo `git commit` na konci každé fáze:
1. Run `bash build-combined.sh` to regenerate `combined.gs`
2. Run `node --check /tmp/c-test.js` for syntax validation
3. Update `znalosti/_learning-log.md` po každé dokončené fázi

---

## PHASE 1: Data fetching rozšíření (previous period, split metrics, SIS)

### Task 1.1: Add `fetchPreviousPeriodProducts` to data.gs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs` (add function after `fetchProducts`)

**Step 1: Write the code**

Add new function that fetches product metrics from previous period (N days before lookback):

```javascript
/**
 * Fetch product stats from previous period (pre-lookback window).
 * Same schema as fetchProducts but for different date range.
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
    '  campaign.name ' +
    'FROM shopping_performance_view ' +
    'WHERE segments.date BETWEEN "' + Utils.formatDate(prevStart) + '" AND "' + Utils.formatDate(prevEnd) + '" ' +
    '  AND campaign.advertising_channel_type IN (' + channelsFilter + ')';

  var rows = [];
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
    Logger.log('WARN: Previous period query failed: ' + e.message);
  }

  return rows;
}
```

**Step 2: Update `fetchAllData` to include previous period**

Add after `var firstClickDates = ...`:

```javascript
var previousPeriodRaw = fetchPreviousPeriodProducts(startDate, endDate, config);
Logger.log('INFO: Previous period: ' + previousPeriodRaw.length + ' rows loaded.');
```

Add to return object:
```javascript
previousPeriodRaw: previousPeriodRaw,
```

**Step 3: Rebuild & validate syntax**

Run: `bash /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/build-combined.sh && cp /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/combined.gs /tmp/c-test.js && node --check /tmp/c-test.js && echo "✅ Syntax OK"`

Expected output: `✅ Syntax OK`

---

### Task 1.2: Refactor `filterCampaigns` to track MAIN / BRAND / REST buckets

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs` (replace `filterCampaigns`)

**Step 1: Replace function**

```javascript
/**
 * Classifies each row as MAIN / BRAND / REST based on campaign name patterns.
 * Returns all rows (no filtering) with added `campaignType` field.
 * Paused campaigns are STILL excluded (no value in including stale data).
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
```

**Step 2: Update `fetchAllData` to use new function**

Replace `var filtered = filterCampaigns(products, config);` with:
```javascript
var classified = classifyCampaignTypes(products, config);
```

Update downstream references from `filtered.kept` to `classified.rows` and `filtered.excluded` to `classified.counts`.

**Step 3: Keep old `filterCampaigns` as thin wrapper for backward compat**

After new function add:
```javascript
function filterCampaigns(products, config) {
  var r = classifyCampaignTypes(products, config);
  // Backward compat: only MAIN in "kept"
  var kept = r.rows.filter(function (p) { return p.campaignType === 'MAIN'; });
  return { kept: kept, excluded: r.counts };
}
```

**Step 4: Rebuild & syntax check**

Run: `bash /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/build-combined.sh && cp /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/combined.gs /tmp/c-test.js && node --check /tmp/c-test.js`

Expected: No syntax errors.

---

### Task 1.3: Rewrite `aggregateByItemId` to produce split metrics

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs` (replace `aggregateByItemId`)

**Step 1: Write the new code**

```javascript
/**
 * Aggregate rows per item_id into split metrics (MAIN / BRAND / REST / TOTAL).
 * Each product ends up with 4 metric buckets.
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
    var bucket = g[r.campaignType.toLowerCase() + '_metrics'];
    bucket.clicks += r.clicks;
    bucket.impressions += r.impressions;
    bucket.cost += r.cost;
    bucket.conversions += r.conversions;
    bucket.conversionValue += r.conversionValue;

    // Primary = campaign with highest cost (any type)
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

  // Compute TOTAL (main + brand + rest) + derived metrics
  var result = [];
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

    // Derived per bucket
    ['main_metrics', 'brand_metrics', 'rest_metrics', 'total_metrics'].forEach(function (bucketKey) {
      var b = item[bucketKey];
      b.ctr = Utils.safeDiv(b.clicks, b.impressions, 0);
      b.cvr = Utils.safeDiv(b.conversions, b.clicks, 0);
      b.roas = Utils.safeDiv(b.conversionValue, b.cost, 0);
      b.avgCpc = Utils.safeDiv(b.cost, b.clicks, 0);
      b.pno = b.conversionValue > 0 ? (b.cost / b.conversionValue * 100) : 0;
    });

    // Backward compat: expose MAIN metrics at top level (for v1 code paths)
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

    result.push(item);
  }

  return result;
}
```

**Step 2: Wire it up in `fetchAllData`**

Replace `var aggregated = aggregateByItemId(filtered.kept);` with:
```javascript
var aggregated = aggregateByItemIdSplit(classified.rows);
Logger.log('INFO: Aggregated ' + aggregated.length + ' unique item_ids with split metrics.');
```

**Step 3: Rebuild & syntax check**

Expected: No syntax errors.

---

### Task 1.4: Aggregate previous period into per-item map

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs`

**Step 1: Add aggregation function**

```javascript
/**
 * Aggregate previous period rows per item_id, split by campaign type.
 * Returns: { itemId: { main_metrics, brand_metrics, rest_metrics, total_metrics } }
 */
function aggregatePreviousPeriodByItemId(rows, config) {
  var groups = {};

  function emptyMetrics() {
    return { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0 };
  }

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var key = r.itemId;

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

  // Compute TOTAL + derived
  for (var k in groups) {
    if (!groups.hasOwnProperty(k)) continue;
    var g = groups[k];
    g.total_metrics = {
      clicks: g.main_metrics.clicks + g.brand_metrics.clicks + g.rest_metrics.clicks,
      cost: g.main_metrics.cost + g.brand_metrics.cost + g.rest_metrics.cost,
      conversions: g.main_metrics.conversions + g.brand_metrics.conversions + g.rest_metrics.conversions,
      conversionValue: g.main_metrics.conversionValue + g.brand_metrics.conversionValue + g.rest_metrics.conversionValue
    };
  }

  return groups;
}
```

**Step 2: Wire up in `fetchAllData`**

After `var previousPeriodRaw = fetchPreviousPeriodProducts(...)` add:

```javascript
var previousPeriodByItemId = aggregatePreviousPeriodByItemId(previousPeriodRaw, config);
```

And in return object:
```javascript
previousPeriodByItemId: previousPeriodByItemId,
```

**Step 3: Rebuild & syntax check**

---

### Task 1.5: Add `fetchImpressionShares` for LOST_OPPORTUNITY

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs`

**Step 1: Write code**

```javascript
/**
 * Fetch search_impression_share per item_id from shopping_product_view.
 * This metric is NOT in shopping_performance_view but IS in shopping_product_view.
 * Returns: { item_id_lowercase: impression_share_0_to_1 }
 */
function fetchImpressionShares(products, config) {
  var result = {};
  if (!products || products.length === 0) return result;

  try {
    var query =
      'SELECT ' +
      '  shopping_product.item_id, ' +
      '  metrics.search_impression_share ' +
      'FROM shopping_product_view ' +
      'WHERE segments.date DURING LAST_30_DAYS';

    var iterator = AdsApp.search(query);
    var count = 0;
    while (iterator.hasNext()) {
      var row = iterator.next();
      var itemId = row.shoppingProduct && row.shoppingProduct.itemId ? row.shoppingProduct.itemId : null;
      if (!itemId) continue;
      var is = Utils.safeParseNumber(row.metrics && row.metrics.searchImpressionShare, 0);
      result[String(itemId).toLowerCase()] = is;
      count++;
    }
    Logger.log('INFO: Loaded search_impression_share for ' + count + ' products from shopping_product_view.');
  } catch (e) {
    Logger.log('WARN: Impression share fetch failed: ' + e.message + ' — LOST_OPPORTUNITY detection will skip.');
  }

  return result;
}
```

**Step 2: Wire up in `fetchAllData`**

Add after `var productPrices = fetchProductPrices(aggregated);`:
```javascript
var impressionShares = fetchImpressionShares(aggregated, config);
```

In return object:
```javascript
impressionShares: impressionShares,
```

**Step 3: Enrich aggregated products with IS**

Before return, loop through aggregated:
```javascript
for (var p = 0; p < aggregated.length; p++) {
  var key = String(aggregated[p].itemId).toLowerCase();
  if (impressionShares[key] !== undefined) {
    aggregated[p].searchImpressionShare = impressionShares[key];
  }
}
```

**Step 4: Rebuild & syntax check**

**Step 5: Validate via MCP**

Verify that `shopping_product_view` returns SIS data:
```
mcp__google-ads__query_gaql customer_id=6746098877
  query="SELECT shopping_product.item_id, metrics.search_impression_share FROM shopping_product_view WHERE segments.date DURING LAST_30_DAYS LIMIT 5"
```

Expected: 5 rows with `searchImpressionShare` values 0-1.

---

### Task 1.6: Enrich products with `previousPeriodByItemId` reference

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs`

**Step 1: In `fetchAllData`, enrich each aggregated product**

Before return, loop through aggregated:
```javascript
for (var p = 0; p < aggregated.length; p++) {
  var prev = previousPeriodByItemId[aggregated[p].itemId];
  if (prev) {
    aggregated[p].main_metrics_previous = prev.main_metrics;
    aggregated[p].brand_metrics_previous = prev.brand_metrics;
    aggregated[p].rest_metrics_previous = prev.rest_metrics;
    aggregated[p].total_metrics_previous = prev.total_metrics;
  } else {
    aggregated[p].main_metrics_previous = null;
    aggregated[p].total_metrics_previous = null;
  }
}
```

**Step 2: Rebuild & syntax check**

---

## PHASE 2: Nové klasifikátory (RISING, DECLINING, LOST_OPPORTUNITY)

### Task 2.1: Add `classifyRising` function

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/classifier.gs`
- Test: `/tmp/loser-test/test-runner.js` (add new cases)

**Step 1: Write failing tests in test-runner.js**

Add after existing YoY tests:

```javascript
console.log('\n━━━ RISING CLASSIFICATION ━━━');

assertPartial('Rising: growth +60%, 5 conv both periods', function () {
  var product = makeProduct({
    total_metrics: { conversions: 8, conversionValue: 1600 },
    total_metrics_previous: { conversions: 5, conversionValue: 1000 }
  });
  var config = makeConfig({ risingGrowthThreshold: 50, minConversionsForTrendCompare: 3 });
  return Classifier.classifyRising(product, config);
}, { matched: true, reason: 'growth' });

assertPartial('Rising: growth +150%, strong_growth reason', function () {
  var product = makeProduct({
    total_metrics: { conversions: 10, conversionValue: 2500 },
    total_metrics_previous: { conversions: 4, conversionValue: 1000 }
  });
  var config = makeConfig({ risingGrowthThreshold: 50, minConversionsForTrendCompare: 3 });
  return Classifier.classifyRising(product, config);
}, { matched: true, reason: 'strong_growth' });

assertPartial('Rising: insufficient previous conv → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 10, conversionValue: 2000 },
    total_metrics_previous: { conversions: 2, conversionValue: 400 }
  });
  var config = makeConfig({ risingGrowthThreshold: 50, minConversionsForTrendCompare: 3 });
  return Classifier.classifyRising(product, config);
}, { matched: false });

assertPartial('Rising: growth < threshold → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 5, conversionValue: 1100 },
    total_metrics_previous: { conversions: 5, conversionValue: 1000 }
  });
  var config = makeConfig({ risingGrowthThreshold: 50, minConversionsForTrendCompare: 3 });
  return Classifier.classifyRising(product, config);
}, { matched: false });
```

**Step 2: Update makeProduct helper to include split metrics**

```javascript
function makeProduct(overrides) {
  var p = {
    itemId: 'test-item-1', productTitle: 'Test', productBrand: 'Test',
    // ... existing fields ...
    main_metrics: { clicks: 200, impressions: 5000, cost: 100, conversions: 2, conversionValue: 200, ctr: 0.04, cvr: 0.01, roas: 2.0, pno: 50 },
    brand_metrics: { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0, ctr: 0, cvr: 0, roas: 0, pno: 0 },
    rest_metrics: { clicks: 0, impressions: 0, cost: 0, conversions: 0, conversionValue: 0, ctr: 0, cvr: 0, roas: 0, pno: 0 },
    total_metrics: { clicks: 200, impressions: 5000, cost: 100, conversions: 2, conversionValue: 200, ctr: 0.04, cvr: 0.01, roas: 2.0, pno: 50 },
    main_metrics_previous: null,
    total_metrics_previous: null,
    // existing top-level props pokračují ...
  };
  // existing override logic
}
```

**Step 3: Run tests, expect FAIL**

Run: `cd /tmp/loser-test && node test-runner.js 2>&1 | tail -20`

Expected: 4 failures ("Classifier.classifyRising is not a function").

**Step 4: Implement `classifyRising` in classifier.gs**

```javascript
/**
 * Detect RISING products (revenue growth vs previous period).
 */
function classifyRising(product, config) {
  var current = product.total_metrics;
  var previous = product.total_metrics_previous;

  if (!previous || !current) return { matched: false };
  if (current.conversions < config.minConversionsForTrendCompare) return { matched: false };
  if (previous.conversions < config.minConversionsForTrendCompare) return { matched: false };
  if (previous.conversionValue <= 0) return { matched: false };

  var growthPct = (current.conversionValue - previous.conversionValue) / previous.conversionValue * 100;

  if (growthPct < config.risingGrowthThreshold) return { matched: false };

  var reason = growthPct >= 100 ? 'strong_growth' : 'growth';

  return {
    matched: true,
    label: 'RISING',
    reason: reason,
    tier: '',
    growthPct: growthPct,
    suggestedAction: 'Early scaling kandidát: zvýšit budget / vydělit do vlastní asset group / zvýšit tROAS ceiling.'
  };
}
```

Add to return object: `classifyRising: classifyRising`.

**Step 5: Run tests, expect PASS**

Expected: 4 new tests pass.

**Step 6: Rebuild combined.gs**

Run: `bash /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/build-combined.sh`

Expected: combined.gs updated, syntax OK.

---

### Task 2.2: Add `classifyDeclining` function

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/classifier.gs`
- Test: `/tmp/loser-test/test-runner.js`

**Step 1: Write tests**

```javascript
console.log('\n━━━ DECLINING CLASSIFICATION ━━━');

assertPartial('Declining: drop 40%, 5 conv both periods', function () {
  var product = makeProduct({
    total_metrics: { conversions: 4, conversionValue: 600 },
    total_metrics_previous: { conversions: 5, conversionValue: 1000 }
  });
  var config = makeConfig({ decliningDropThreshold: 30, minConversionsForTrendCompare: 3 });
  return Classifier.classifyDeclining(product, config);
}, { matched: true, reason: 'decline' });

assertPartial('Declining: drop 60% → critical_decline', function () {
  var product = makeProduct({
    total_metrics: { conversions: 3, conversionValue: 400 },
    total_metrics_previous: { conversions: 5, conversionValue: 1000 }
  });
  var config = makeConfig({ decliningDropThreshold: 30, minConversionsForTrendCompare: 3 });
  return Classifier.classifyDeclining(product, config);
}, { matched: true, reason: 'critical_decline' });

assertPartial('Declining: small drop → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 5, conversionValue: 850 },
    total_metrics_previous: { conversions: 5, conversionValue: 1000 }
  });
  var config = makeConfig({ decliningDropThreshold: 30, minConversionsForTrendCompare: 3 });
  return Classifier.classifyDeclining(product, config);
}, { matched: false });
```

**Step 2: Run tests, expect FAIL**

**Step 3: Implement `classifyDeclining`**

```javascript
function classifyDeclining(product, config) {
  var current = product.total_metrics;
  var previous = product.total_metrics_previous;

  if (!previous || !current) return { matched: false };
  if (current.conversions < config.minConversionsForTrendCompare) return { matched: false };
  if (previous.conversions < config.minConversionsForTrendCompare) return { matched: false };
  if (previous.conversionValue <= 0) return { matched: false };

  var growthPct = (current.conversionValue - previous.conversionValue) / previous.conversionValue * 100;

  if (growthPct > -config.decliningDropThreshold) return { matched: false };

  var reason = growthPct <= -50 ? 'critical_decline' : 'decline';

  return {
    matched: true,
    label: 'DECLINING',
    reason: reason,
    tier: '',
    growthPct: growthPct,
    suggestedAction: 'Investigate: zkontroluj cenu vs konkurence, stav skladu, sezónnost, nedávné změny feed dat.'
  };
}
```

Add to exports.

**Step 4: Run tests, expect PASS**

**Step 5: Rebuild**

---

### Task 2.3: Add `classifyLostOpportunity` function

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/classifier.gs`
- Test: `/tmp/loser-test/test-runner.js`

**Step 1: Write tests**

```javascript
console.log('\n━━━ LOST_OPPORTUNITY CLASSIFICATION ━━━');

assertPartial('LostOpp: rentabilní + low IS → FLAG', function () {
  var product = makeProduct({
    total_metrics: { conversions: 8, conversionValue: 3200, cost: 500, pno: 15 },
    searchImpressionShare: 0.3
  });
  var config = makeConfig({ lostOpportunityMinConv: 5, lostOpportunityMaxPnoMultiplier: 0.8, lostOpportunityMaxImpressionShare: 0.5 });
  return Classifier.classifyLostOpportunity(product, config);
}, { matched: true });

assertPartial('LostOpp: high IS → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 8, conversionValue: 3200, cost: 500, pno: 15 },
    searchImpressionShare: 0.7
  });
  var config = makeConfig({ lostOpportunityMinConv: 5, lostOpportunityMaxPnoMultiplier: 0.8, lostOpportunityMaxImpressionShare: 0.5 });
  return Classifier.classifyLostOpportunity(product, config);
}, { matched: false });

assertPartial('LostOpp: too few conv → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 3, conversionValue: 1200, cost: 180, pno: 15 },
    searchImpressionShare: 0.3
  });
  var config = makeConfig({ lostOpportunityMinConv: 5, lostOpportunityMaxPnoMultiplier: 0.8, lostOpportunityMaxImpressionShare: 0.5 });
  return Classifier.classifyLostOpportunity(product, config);
}, { matched: false });

assertPartial('LostOpp: not rentable (PNO too high) → skip', function () {
  var product = makeProduct({
    total_metrics: { conversions: 8, conversionValue: 2000, cost: 500, pno: 25 },
    searchImpressionShare: 0.3
  });
  var config = makeConfig({ targetPnoPct: 20, lostOpportunityMinConv: 5, lostOpportunityMaxPnoMultiplier: 0.8, lostOpportunityMaxImpressionShare: 0.5 });
  return Classifier.classifyLostOpportunity(product, config);
}, { matched: false });
```

**Step 2: Run tests, expect FAIL**

**Step 3: Implement**

```javascript
function classifyLostOpportunity(product, config) {
  var total = product.total_metrics;
  if (!total) return { matched: false };
  if (total.conversions < config.lostOpportunityMinConv) return { matched: false };

  var pnoLimit = config.targetPnoPct * config.lostOpportunityMaxPnoMultiplier;
  if (total.pno > pnoLimit) return { matched: false };
  if (total.pno <= 0) return { matched: false };

  var is = product.searchImpressionShare;
  if (is === null || is === undefined) return { matched: false };
  if (is > config.lostOpportunityMaxImpressionShare) return { matched: false };

  return {
    matched: true,
    label: 'LOST_OPPORTUNITY',
    reason: 'low_is_high_roas',
    tier: '',
    impressionShare: is,
    suggestedAction: 'Rentabilní produkt s nízkou visibility — zvýšit bid, samostatná asset group, nebo zahrnout do top-priority kampaně.'
  };
}
```

Add to exports.

**Step 4: Run tests, expect PASS**

**Step 5: Rebuild**

---

### Task 2.4: Add new CONFIG parameters with validation

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/_config.gs`
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/config.gs`
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/main.gs` (CONFIG tab rows)

**Step 1: Add to `_config.gs` CONFIG object**

Add after `lowCtrSkipIfProfitableMinConv`:

```javascript
  // === TREND DETECTION (RISING/DECLINING) ===
  risingGrowthThreshold:         50,    // Growth >= X% = RISING
  decliningDropThreshold:        30,    // Drop >= X% = DECLINING
  minConversionsForTrendCompare:  3,    // Min conv obou periodách

  // === LOST_OPPORTUNITY ===
  lostOpportunityMinConv:           5,
  lostOpportunityMaxPnoMultiplier:  0.8,  // PNO <= target × N
  lostOpportunityMaxImpressionShare: 0.5, // IS < N

  // === EFFECTIVENESS ===
  effectivenessMinDaysSinceAction:    14,
  restCampaignEfficientThreshold:     0.2,  // Rest cost <= N × before_main
```

**Step 2: Add validation in `config.gs`**

```javascript
if (!isNumberInRange(config.risingGrowthThreshold, 10, 500)) {
  errors.push('risingGrowthThreshold musi byt 10-500 (ted: ' + config.risingGrowthThreshold + ')');
}
if (!isNumberInRange(config.decliningDropThreshold, 10, 95)) {
  errors.push('decliningDropThreshold musi byt 10-95 (ted: ' + config.decliningDropThreshold + ')');
}
if (!isIntegerInRange(config.minConversionsForTrendCompare, 1, 100)) {
  errors.push('minConversionsForTrendCompare musi byt integer 1-100');
}
if (!isIntegerInRange(config.lostOpportunityMinConv, 1, 100)) {
  errors.push('lostOpportunityMinConv musi byt integer 1-100');
}
if (!isNumberInRange(config.lostOpportunityMaxPnoMultiplier, 0.1, 2.0)) {
  errors.push('lostOpportunityMaxPnoMultiplier musi byt 0.1-2.0');
}
if (!isNumberInRange(config.lostOpportunityMaxImpressionShare, 0.05, 1.0)) {
  errors.push('lostOpportunityMaxImpressionShare musi byt 0.05-1.0');
}
if (!isIntegerInRange(config.effectivenessMinDaysSinceAction, 1, 180)) {
  errors.push('effectivenessMinDaysSinceAction musi byt 1-180');
}
if (!isNumberInRange(config.restCampaignEfficientThreshold, 0.0, 1.0)) {
  errors.push('restCampaignEfficientThreshold musi byt 0.0-1.0');
}
```

**Step 3: Add to CONFIG tab in setupOutputSheet (main.gs)**

Najdi `['▸ POKROCILE', '', ''],` sekci a před ni přidej:

```javascript
    ['', '', ''],
    ['▸ TREND DETECTION (RISING/DECLINING)', '', ''],
    ['risingGrowthThreshold', 50, 'Growth ≥ X% = RISING (default 50% medium)'],
    ['decliningDropThreshold', 30, 'Drop ≥ X% = DECLINING (default 30% medium)'],
    ['minConversionsForTrendCompare', 3, 'Min conversions v obou periodách'],
    ['', '', ''],
    ['▸ LOST_OPPORTUNITY', '', ''],
    ['lostOpportunityMinConv', 5, 'Min conversions pro rentability claim'],
    ['lostOpportunityMaxPnoMultiplier', 0.8, 'PNO ≤ target × N (výrazně rentabilní)'],
    ['lostOpportunityMaxImpressionShare', 0.5, 'IS < N (Google málo zobrazuje)'],
    ['', '', ''],
    ['▸ EFFECTIVENESS', '', ''],
    ['effectivenessMinDaysSinceAction', 14, 'Dní před vyhodnocením účinnosti'],
    ['restCampaignEfficientThreshold', 0.2, 'Rest cost ≤ N × before_main pro "efficient"'],
```

**Step 4: Rebuild & syntax check**

---

### Task 2.5: Wire up new classifiers in `classifyProduct` pipeline

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/classifier.gs`

**Step 1: Update `classifyProduct` to call new classifiers**

After existing `lowCtr` classification, add:

```javascript
    // === KLASIFIKACE 3: RISING ===
    var rising = classifyRising(product, config);

    // === KLASIFIKACE 4: DECLINING ===
    var declining = classifyDeclining(product, config);

    // === KLASIFIKACE 5: LOST_OPPORTUNITY ===
    var lostOpp = classifyLostOpportunity(product, config);
```

**Step 2: Update `mergeResults` for 5-way priority**

Replace existing `mergeResults` with:

```javascript
function mergeResults(loser, lowCtr, rising, declining, lostOpp, config) {
  var primary = '';
  var secondary = [];
  var reason = '';
  var tier = '';
  var note = '';
  var suggestedAction = '';
  var growthPct = null;

  // Priority: LOSER > LOW_CTR > DECLINING > LOST_OPP > RISING
  if (loser.matched) {
    primary = loser.label;
    reason = loser.reason;
    tier = loser.tier;
    note = loser.note;
    if (lowCtr.matched) secondary.push(lowCtr.label);
    if (declining.matched) secondary.push('DECLINING');
    suggestedAction = lowCtr.suggestedAction || '';
  } else if (lowCtr.matched) {
    primary = lowCtr.label;
    reason = lowCtr.reason;
    note = '';
    if (declining.matched) secondary.push('DECLINING');
    suggestedAction = lowCtr.suggestedAction || '';
  } else if (declining.matched) {
    primary = declining.label;
    reason = declining.reason;
    growthPct = declining.growthPct;
    suggestedAction = declining.suggestedAction || '';
  } else if (lostOpp.matched) {
    primary = lostOpp.label;
    reason = lostOpp.reason;
    suggestedAction = lostOpp.suggestedAction || '';
  } else if (rising.matched) {
    primary = rising.label;
    reason = rising.reason;
    growthPct = rising.growthPct;
    suggestedAction = rising.suggestedAction || '';
  } else if (lowCtr.skipped_profitable) {
    note = lowCtr.note || '';
  }

  return {
    primaryLabel: primary,
    secondaryFlags: secondary,
    reasonCode: reason,
    tier: tier,
    note: note,
    suggestedAction: suggestedAction,
    growthPct: growthPct
  };
}
```

**Step 3: Update `classifyProduct` to call new mergeResults**

```javascript
var merged = mergeResults(loser, lowCtr, rising, declining, lostOpp, config);
```

Add to result:
```javascript
result.growthPct = merged.growthPct;
```

And update result schema (line ~64) to include `growthPct: null`.

**Step 4: Run all tests**

Run: `cd /tmp/loser-test && node test-runner.js 2>&1 | tail -30`

Expected: all 46 original + new ~11 pass.

**Step 5: Rebuild**

---

## PHASE 3: Preserve mechanismus pro manual input

### Task 3.1: Add `readExistingActions` to data.gs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs`

**Step 1: Write code**

```javascript
/**
 * Read existing ACTIONS tab and extract manual columns per item_id.
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
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

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
      result[itemId] = {
        action_taken: idxAction >= 0 ? (data[i][idxAction] || '') : '',
        action_date: idxDate >= 0 ? (data[i][idxDate] || '') : '',
        consultant_note: idxNote >= 0 ? (data[i][idxNote] || '') : ''
      };
    }
    Logger.log('INFO: Read ' + Object.keys(result).length + ' ACTIONS rows for manual preserve.');
  } catch (e) {
    Logger.log('WARN: readExistingActions failed: ' + e.message);
  }
  return result;
}
```

Export it: `readExistingActions: readExistingActions`.

**Step 2: Wire up in main.gs**

After `previousLifecycleMap = DataLayer.fetchPreviousLifecycle(...)`:
```javascript
var existingActions = CONFIG.dryRun ? {} : DataLayer.readExistingActions(CONFIG.outputSheetId);
```

**Step 3: Rebuild & syntax**

---

### Task 3.2: Add `readExistingProductTimeline` to data.gs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/data.gs`

**Step 1: Write similar function for PRODUCT_TIMELINE**

```javascript
/**
 * Read existing PRODUCT_TIMELINE tab.
 * Returns: { item_id: { first_flag_date, total_runs_flagged, categories_history,
 *                        kpi_before_*, latest_action, latest_action_date, latest_note } }
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

    for (var i = 0; i < data.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = data[i][j];
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
```

Export.

**Step 2: Wire up in main.gs**

```javascript
var existingTimeline = CONFIG.dryRun ? {} : DataLayer.readExistingProductTimeline(CONFIG.outputSheetId);
```

**Step 3: Rebuild**

---

## PHASE 4: Nové taby (ACTIONS, PRODUCT_TIMELINE, WEEKLY_SNAPSHOT)

### Task 4.1: Implement `writeActionsTab` in output.gs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Add new function**

```javascript
/**
 * Write ACTIONS tab — all flagged/insight products in one filterable view.
 * Preserves manual columns from existingActions map.
 */
function writeActionsTab(ss, classified, existingActions, config, summary) {
  var sheet = getOrCreateSheet(ss, 'ACTIONS');
  sheet.clearContents();

  var headers = [
    'priority_rank', 'category', 'item_id', 'product_title', 'product_price',
    'current_campaign', 'tier', 'reason_code',
    'main_clicks', 'main_impressions', 'main_cost', 'main_conv', 'main_pno_pct', 'main_ctr_pct',
    'total_clicks', 'total_cost', 'total_conv', 'total_roas', 'brand_share_pct',
    'growth_pct', 'wasted_spend', 'recommended_action',
    'days_since_first_flag', 'transition_status', 'secondary_flags',
    'action_taken', 'action_date', 'consultant_note'
  ];

  // Filter flagged only, sort by priority
  var flagged = classified.filter(function (c) { return c.primaryLabel && c.primaryLabel.length > 0; });

  // Compute priority_rank
  flagged.sort(function (a, b) {
    return computePriorityScore(b) - computePriorityScore(a);
  });

  var data = [headers];
  for (var i = 0; i < flagged.length; i++) {
    var c = flagged[i];
    var manual = existingActions[c.itemId] || {};

    var totalConv = (c.total_metrics && c.total_metrics.conversions) || 0;
    var brandConv = (c.brand_metrics && c.brand_metrics.conversions) || 0;
    var brandSharePct = totalConv > 0 ? (brandConv / totalConv * 100) : 0;

    data.push([
      i + 1,                              // priority_rank
      c.primaryLabel,                     // category
      c.itemId,
      c.productTitle || '',
      roundNumber(c.productPrice, 2),
      c.primaryCampaignName || '',
      c.tier || '',
      c.reasonCode || '',
      c.main_metrics ? c.main_metrics.clicks : c.clicks,
      c.main_metrics ? c.main_metrics.impressions : c.impressions,
      roundNumber(c.main_metrics ? c.main_metrics.cost : c.cost, 2),
      c.main_metrics ? c.main_metrics.conversions : c.conversions,
      roundNumber(c.main_metrics ? c.main_metrics.pno : c.actualPno, 2),
      roundNumber((c.main_metrics ? c.main_metrics.ctr : c.ctr) * 100, 3),
      c.total_metrics ? c.total_metrics.clicks : 0,
      roundNumber(c.total_metrics ? c.total_metrics.cost : 0, 2),
      c.total_metrics ? c.total_metrics.conversions : 0,
      roundNumber(c.total_metrics ? c.total_metrics.roas : 0, 2),
      roundNumber(brandSharePct, 1),
      c.growthPct !== null && c.growthPct !== undefined ? roundNumber(c.growthPct, 1) : '',
      roundNumber(c.wastedSpend, 2),
      c.suggestedAction || '',
      computeDaysSinceFirstFlag(c, summary),
      c.transitionType || 'NEW_FLAG',
      (c.secondaryFlags || []).join(', '),
      manual.action_taken || '',
      manual.action_date || '',
      manual.consultant_note || ''
    ]);
  }

  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');

  // Apply category-based background colors
  applyCategoryColors(sheet, flagged, headers.length);

  Logger.log('INFO: ACTIONS — written ' + flagged.length + ' rows.');
}

function computePriorityScore(c) {
  if (c.primaryLabel === 'loser_rest') return c.wastedSpend || 0;
  if (c.primaryLabel === 'DECLINING') return Math.abs(c.growthPct || 0) * 10;
  if (c.primaryLabel === 'LOST_OPPORTUNITY') return (c.total_metrics && c.total_metrics.conversionValue) || 0;
  if (c.primaryLabel === 'RISING') return (c.growthPct || 0) * 10;
  return (c.impressions || 0) / 100; // LOW_CTR
}

function computeDaysSinceFirstFlag(c, summary) {
  // Placeholder — bude wired z PRODUCT_TIMELINE v Task 4.4
  return '';
}

function applyCategoryColors(sheet, flagged, numCols) {
  var colors = {
    'loser_rest': '#fde8e8',
    'low_ctr_audit': '#fff8d4',
    'DECLINING': '#ffe0cc',
    'RISING': '#d4f5d4',
    'LOST_OPPORTUNITY': '#d4e4f5'
  };
  for (var i = 0; i < flagged.length; i++) {
    var color = colors[flagged[i].primaryLabel];
    if (color) {
      sheet.getRange(i + 2, 1, 1, numCols).setBackground(color);
    }
  }
}
```

Export as part of Output.

**Step 2: Rebuild & syntax check**

---

### Task 4.2: Implement `writeProductTimelineTab`

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Add function**

```javascript
/**
 * Write PRODUCT_TIMELINE tab — per-product history, upsert logic.
 * Merges new data with existing timeline (preserve history + manual columns).
 */
function writeProductTimelineTab(ss, classified, existingTimeline, existingActions, config, runDate) {
  var sheet = getOrCreateSheet(ss, 'PRODUCT_TIMELINE');

  var headers = [
    'item_id', 'product_title', 'first_flag_date', 'total_runs_flagged',
    'categories_history', 'current_status',
    // Before snapshot (only captured once, when first flagged)
    'kpi_before_cost', 'kpi_before_conv', 'kpi_before_pno', 'kpi_before_roas', 'kpi_before_ctr',
    // Current snapshots
    'kpi_current_total_cost', 'kpi_current_total_conv', 'kpi_current_total_pno', 'kpi_current_total_roas', 'kpi_current_total_ctr',
    'kpi_current_main_cost', 'kpi_current_main_conv', 'kpi_current_main_pno',
    'kpi_current_rest_cost', 'kpi_current_rest_conv',
    'kpi_current_brand_cost', 'kpi_current_brand_conv',
    // Deltas
    'delta_cost_pct', 'delta_conv_pct', 'delta_pno_pct', 'delta_roas_pct',
    'effectiveness_score', 'days_since_action',
    'latest_action', 'latest_action_date', 'latest_note'
  ];

  // Collect all item_ids: current flagged + previously in timeline
  var itemIds = {};
  classified.forEach(function (c) { if (c.primaryLabel) itemIds[c.itemId] = true; });
  for (var k in existingTimeline) { if (existingTimeline.hasOwnProperty(k)) itemIds[k] = true; }

  var data = [headers];
  var runDateStr = Utils.formatDate(runDate);

  for (var itemId in itemIds) {
    if (!itemIds.hasOwnProperty(itemId)) continue;

    var current = findClassifiedById(classified, itemId);
    var existing = existingTimeline[itemId];
    var manual = existingActions[itemId] || {};

    var row = buildTimelineRow(itemId, current, existing, manual, runDateStr, config);
    data.push(row);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');

  Logger.log('INFO: PRODUCT_TIMELINE — written ' + (data.length - 1) + ' rows (upsert).');
}

function findClassifiedById(classified, itemId) {
  for (var i = 0; i < classified.length; i++) {
    if (classified[i].itemId === itemId) return classified[i];
  }
  return null;
}

function buildTimelineRow(itemId, current, existing, manual, runDateStr, config) {
  var isFlagged = current && current.primaryLabel;
  var firstFlagDate = (existing && existing.first_flag_date) || (isFlagged ? runDateStr : '');
  var totalRunsFlagged = (existing && existing.total_runs_flagged) || 0;
  if (isFlagged) totalRunsFlagged = totalRunsFlagged + 1;

  var categoriesHistory = existing && existing.categories_history ? String(existing.categories_history) : '';
  if (isFlagged) {
    var newCat = current.primaryLabel + (current.tier ? ':' + current.tier : '');
    if (categoriesHistory.indexOf(newCat) === -1) {
      categoriesHistory = categoriesHistory ? (categoriesHistory + ' → ' + newCat) : newCat;
    }
  } else if (existing && existing.current_status === 'FLAGGED') {
    categoriesHistory = categoriesHistory + ' → RESOLVED';
  }

  var currentStatus = isFlagged ? 'FLAGGED' : (existing && existing.first_flag_date ? 'RESOLVED' : 'STABLE');

  // KPI before: snapshot from existing if present, else capture now
  var kpiBefore = {
    cost: (existing && existing.kpi_before_cost) || (isFlagged && current && current.total_metrics_previous ? current.total_metrics_previous.cost : ''),
    conv: (existing && existing.kpi_before_conv) || (isFlagged && current && current.total_metrics_previous ? current.total_metrics_previous.conversions : ''),
    pno: (existing && existing.kpi_before_pno) || (isFlagged && current && current.total_metrics_previous ? (current.total_metrics_previous.cost / (current.total_metrics_previous.conversionValue || 1) * 100) : ''),
    roas: (existing && existing.kpi_before_roas) || (isFlagged && current && current.total_metrics_previous ? (current.total_metrics_previous.conversionValue / (current.total_metrics_previous.cost || 1)) : ''),
    ctr: (existing && existing.kpi_before_ctr) || ''
  };

  // Current snapshots
  var t = current && current.total_metrics ? current.total_metrics : { cost: 0, conversions: 0, pno: 0, roas: 0, ctr: 0 };
  var m = current && current.main_metrics ? current.main_metrics : { cost: 0, conversions: 0, pno: 0 };
  var r = current && current.rest_metrics ? current.rest_metrics : { cost: 0, conversions: 0 };
  var b = current && current.brand_metrics ? current.brand_metrics : { cost: 0, conversions: 0 };

  // Deltas
  var deltaCost = kpiBefore.cost > 0 ? ((t.cost - kpiBefore.cost) / kpiBefore.cost * 100) : '';
  var deltaConv = kpiBefore.conv > 0 ? ((t.conversions - kpiBefore.conv) / kpiBefore.conv * 100) : '';
  var deltaPno = kpiBefore.pno > 0 ? ((t.pno - kpiBefore.pno) / kpiBefore.pno * 100) : '';
  var deltaRoas = kpiBefore.roas > 0 ? ((t.roas - kpiBefore.roas) / kpiBefore.roas * 100) : '';

  // Effectiveness
  var effectivenessScore = computeEffectivenessScore(manual, kpiBefore, t, runDateStr, config);
  var daysSinceAction = computeDaysSinceAction(manual, runDateStr);

  return [
    itemId,
    current ? (current.productTitle || '') : (existing && existing.product_title) || '',
    firstFlagDate, totalRunsFlagged, categoriesHistory, currentStatus,
    roundNumber(kpiBefore.cost, 2), roundNumber(kpiBefore.conv, 2), roundNumber(kpiBefore.pno, 2), roundNumber(kpiBefore.roas, 2), roundNumber(kpiBefore.ctr, 3),
    roundNumber(t.cost, 2), roundNumber(t.conversions, 2), roundNumber(t.pno, 2), roundNumber(t.roas, 2), roundNumber((t.ctr || 0) * 100, 3),
    roundNumber(m.cost, 2), roundNumber(m.conversions, 2), roundNumber(m.pno, 2),
    roundNumber(r.cost, 2), roundNumber(r.conversions, 2),
    roundNumber(b.cost, 2), roundNumber(b.conversions, 2),
    typeof deltaCost === 'number' ? roundNumber(deltaCost, 1) : '',
    typeof deltaConv === 'number' ? roundNumber(deltaConv, 1) : '',
    typeof deltaPno === 'number' ? roundNumber(deltaPno, 1) : '',
    typeof deltaRoas === 'number' ? roundNumber(deltaRoas, 1) : '',
    effectivenessScore,
    daysSinceAction,
    manual.action_taken || '',
    manual.action_date || '',
    manual.consultant_note || ''
  ];
}

function computeEffectivenessScore(manual, kpiBefore, kpiCurrent, runDateStr, config) {
  if (!manual.action_date) return 'N/A';
  var days = computeDaysSinceAction(manual, runDateStr);
  if (days < config.effectivenessMinDaysSinceAction) return 'PENDING';
  if (kpiBefore.cost <= 0) return 'N/A';

  var deltaCostPct = (kpiCurrent.cost - kpiBefore.cost) / kpiBefore.cost * 100;
  var deltaRoasPct = kpiBefore.roas > 0 ? ((kpiCurrent.roas - kpiBefore.roas) / kpiBefore.roas * 100) : 0;

  if (deltaCostPct <= -30 && deltaRoasPct >= -10) return '+';
  if (deltaCostPct > -10 || deltaRoasPct < -30) return '-';
  return '=';
}

function computeDaysSinceAction(manual, runDateStr) {
  if (!manual.action_date) return '';
  var parts = String(manual.action_date).split('-');
  if (parts.length !== 3) return '';
  var actionDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  var runDateParts = runDateStr.split('-');
  var runDate = new Date(parseInt(runDateParts[0], 10), parseInt(runDateParts[1], 10) - 1, parseInt(runDateParts[2], 10));
  return Utils.daysBetween(actionDate, runDate);
}
```

Export.

**Step 2: Rebuild & syntax check**

---

### Task 4.3: Implement `appendWeeklySnapshot`

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Write function**

```javascript
/**
 * Append 1 row to WEEKLY_SNAPSHOT per run.
 * Data: account-level KPI, flagged counts, wasted spend, transitions.
 */
function appendWeeklySnapshot(ss, summary, effectiveness, classified, runDate) {
  var sheet = getOrCreateSheet(ss, 'WEEKLY_SNAPSHOT');

  if (sheet.getLastRow() === 0) {
    var headers = [
      'run_date', 'week_id', 'account_cost_total', 'account_clicks', 'account_conversions',
      'account_conv_value', 'account_roas', 'account_pno_pct', 'account_ctr_pct',
      'flagged_count_total', 'flagged_loser_rest', 'flagged_low_ctr',
      'flagged_declining', 'flagged_rising', 'flagged_lost_opp',
      'wasted_spend_total', 'resolved_this_run', 're_flagged_this_run',
      'label_application_rate_pct'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  }

  var weekId = computeWeekId(runDate);
  var categoryCounts = {
    loser_rest: 0, low_ctr_audit: 0, DECLINING: 0, RISING: 0, LOST_OPPORTUNITY: 0
  };
  classified.forEach(function (c) {
    if (categoryCounts[c.primaryLabel] !== undefined) categoryCounts[c.primaryLabel]++;
  });

  var row = [
    Utils.formatDate(runDate), weekId,
    roundNumber(summary.accountBaseline.totalCost, 2),
    summary.accountBaseline.totalClicks,
    roundNumber(summary.accountBaseline.totalConversions, 2),
    roundNumber(summary.accountBaseline.totalConversionValue, 2),
    roundNumber(summary.accountBaseline.avgRoas, 2),
    roundNumber(summary.accountBaseline.avgPno, 2),
    roundNumber(summary.accountBaseline.avgCtr * 100, 3),
    summary.flags.totalFlagged,
    categoryCounts.loser_rest, categoryCounts.low_ctr_audit,
    categoryCounts.DECLINING, categoryCounts.RISING, categoryCounts.LOST_OPPORTUNITY,
    roundNumber(summary.flags.totalWastedSpend, 2),
    (effectiveness && effectiveness.transitions && effectiveness.transitions.RESOLVED) || 0,
    (effectiveness && effectiveness.transitions && effectiveness.transitions.RE_FLAGGED) || 0,
    effectiveness && effectiveness.applicationRate !== null ? roundNumber(effectiveness.applicationRate, 1) : ''
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  Logger.log('INFO: WEEKLY_SNAPSHOT appended (week ' + weekId + ').');
}

function computeWeekId(date) {
  var y = date.getFullYear();
  var firstDayOfYear = new Date(y, 0, 1);
  var pastDaysOfYear = (date - firstDayOfYear) / 86400000;
  var weekNum = Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
  return y + '-W' + (weekNum < 10 ? '0' : '') + weekNum;
}
```

Export.

**Step 2: Wire up in main.gs**

Po existing `Output.writeAll(...)` v main.gs změnit na rozšířenou verzi (viz Task 4.4).

**Step 3: Rebuild & syntax**

---

### Task 4.4: Update `Output.writeAll` to orchestrate new tabs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Update writeAll signature & body**

```javascript
function writeAll(outputSheetId, classified, summary, effectiveness, config, runDate,
                   existingActions, existingTimeline) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(outputSheetId);
  } catch (e) {
    throw new Error('Nepodarilo se otevrit output sheet: ' + e.message);
  }

  Logger.log('INFO: Writing to sheet "' + ss.getName() + '"');

  writeFeedUploadTab(ss, classified, config);
  writeActionsTab(ss, classified, existingActions || {}, config, summary);
  writeProductTimelineTab(ss, classified, existingTimeline || {}, existingActions || {}, config, runDate);
  writeDetailTab(ss, classified, config);
  writeSummaryTab(ss, summary, effectiveness, config, runDate);
  appendLifecycleLogTab(ss, classified, runDate);
  appendWeeklySnapshot(ss, summary, effectiveness, classified, runDate);

  Logger.log('INFO: Output written.');
}
```

**Step 2: Update main.gs call**

```javascript
Output.writeAll(CONFIG.outputSheetId, classified, summary, effectiveness, CONFIG, runDate,
                existingActions, existingTimeline);
```

**Step 3: Rebuild & syntax**

---

### Task 4.5: Update setupOutputSheet to create new tabs

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/main.gs`

**Step 1: Add ACTIONS, PRODUCT_TIMELINE, WEEKLY_SNAPSHOT tab creation**

V `setupOutputSheet` po existujících tabs přidat placeholders (budou populated při prvním runu):

```javascript
  // === Tab: ACTIONS (placeholder) ===
  var actionsSheet = ss.insertSheet('ACTIONS');
  actionsSheet.getRange(1, 1).setValue('(naplni se po prvnim runu main())');

  // === Tab: PRODUCT_TIMELINE (placeholder) ===
  var timelineSheet = ss.insertSheet('PRODUCT_TIMELINE');
  timelineSheet.getRange(1, 1).setValue('(naplni se po prvnim runu main() — per-produkt historie)');

  // === Tab: WEEKLY_SNAPSHOT (placeholder) ===
  var snapshotSheet = ss.insertSheet('WEEKLY_SNAPSHOT');
  snapshotSheet.getRange(1, 1).setValue('(naplni se po kazdem runu — 1 radek per tyden pro trendy)');
```

Vložit mezi FEED_UPLOAD a DETAIL tab creation.

**Step 2: Update README tab content**

Doplnit popis nových tabů v readmeContent array.

**Step 3: Rebuild**

---

## PHASE 5: DASHBOARD rozšíření (trends, REST health, brand insights)

### Task 5.1: Add weekly trend sparklines do writeSummaryTab

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Add helper to read last N weekly snapshots**

```javascript
function readWeeklyTrends(ss, maxWeeks) {
  var sheet = ss.getSheetByName('WEEKLY_SNAPSHOT');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(2, lastRow - maxWeeks + 1);
  var numRows = lastRow - startRow + 1;
  var lastCol = sheet.getLastColumn();

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

  return data.map(function (row) {
    var obj = {};
    for (var i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
    return obj;
  });
}
```

**Step 2: Add trend section do writeSummaryTab**

After existing FLAGGED sekci:

```javascript
    rows.push(['▸ WEEKLY TRENDS (posledních 8 týdnů)']);
    var trends = readWeeklyTrends(ss, 8);
    if (trends.length < 2) {
      rows.push(['(čekáme na více dat — potřeba min 2 týdny)', '', '']);
    } else {
      rows.push(['Week', 'ROAS', 'PNO%', 'Wasted', 'Flagged']);
      trends.forEach(function (t) {
        rows.push([t.week_id, roundNumber(t.account_roas, 2),
                    roundNumber(t.account_pno_pct, 1),
                    roundNumber(t.wasted_spend_total, 0),
                    t.flagged_count_total]);
      });
    }
    rows.push([]);
```

**Step 3: Rebuild & validate**

---

### Task 5.2: Add REST CAMPAIGN HEALTH section

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Add helper**

```javascript
function computeRestCampaignHealth(classified, config) {
  var counts = { total: 0, efficient: 0, acceptable: 0, wasteful: 0 };
  var totalRestCost = 0;

  classified.forEach(function (c) {
    if (!c.rest_metrics || c.rest_metrics.cost <= 0) return;
    counts.total++;
    totalRestCost += c.rest_metrics.cost;

    // Need kpi_before_main_cost from timeline — placeholder: compare to current main
    var mainCostBefore = c.main_metrics ? c.main_metrics.cost : 1;
    var ratio = mainCostBefore > 0 ? (c.rest_metrics.cost / mainCostBefore) : 0;

    if (ratio <= config.restCampaignEfficientThreshold) counts.efficient++;
    else if (ratio <= 0.5) counts.acceptable++;
    else counts.wasteful++;
  });

  return { counts: counts, totalRestCost: totalRestCost };
}
```

**Step 2: Add section in writeSummaryTab**

```javascript
    rows.push(['▸ REST CAMPAIGN HEALTH']);
    var restHealth = computeRestCampaignHealth(classified, config);
    rows.push(['Produktů v rest:', restHealth.counts.total]);
    rows.push(['Rest cost celkem:', roundNumber(restHealth.totalRestCost, 2) + ' ' + summary.currency]);
    rows.push(['  Efficient (<20% pre-cost):', restHealth.counts.efficient]);
    rows.push(['  Acceptable (<50%):', restHealth.counts.acceptable]);
    rows.push(['  ⚠ Wasteful (>50%):', restHealth.counts.wasteful]);
    rows.push([]);
```

**Step 3: Rebuild**

---

### Task 5.3: Add BRAND INSIGHTS section

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Helper**

```javascript
function computeBrandInsights(classified) {
  var brandOnlySellers = [];
  var brandDependent = [];

  classified.forEach(function (c) {
    if (!c.total_metrics || c.total_metrics.conversions < 1) return;

    var mainConv = c.main_metrics ? c.main_metrics.conversions : 0;
    var brandConv = c.brand_metrics ? c.brand_metrics.conversions : 0;

    if (mainConv === 0 && brandConv > 0) {
      brandOnlySellers.push({ itemId: c.itemId, revenue: c.brand_metrics.conversionValue });
    }

    var brandShare = c.total_metrics.conversionValue > 0
      ? (c.brand_metrics ? c.brand_metrics.conversionValue : 0) / c.total_metrics.conversionValue
      : 0;
    if (brandShare > 0.5) {
      brandDependent.push({ itemId: c.itemId, share: brandShare });
    }
  });

  var brandOnlyRevenue = brandOnlySellers.reduce(function (s, b) { return s + (b.revenue || 0); }, 0);

  return {
    brandOnlySellers: brandOnlySellers,
    brandDependent: brandDependent,
    brandOnlyCount: brandOnlySellers.length,
    brandOnlyRevenue: brandOnlyRevenue
  };
}
```

**Step 2: Add section in writeSummaryTab**

```javascript
    rows.push(['▸ BRAND INSIGHTS']);
    var bi = computeBrandInsights(classified);
    rows.push(['Brand-only sellers (conv jen přes brand):', bi.brandOnlyCount]);
    rows.push(['Revenue z brand-only:', roundNumber(bi.brandOnlyRevenue, 2) + ' ' + summary.currency]);
    rows.push(['Brand-dependent (>50% conv přes brand):', bi.brandDependent.length]);
    if (bi.brandOnlyCount > 0) {
      rows.push(['Insight pro klienta:', 'Tyto produkty prodávají jen přes brand — kandidáti na marketing awareness/promo.']);
    }
    rows.push([]);
```

**Step 3: Rebuild**

---

## PHASE 6: Effectiveness logika

### Task 6.1: Wire effectiveness score do DASHBOARD

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/output.gs`

**Step 1: Read PRODUCT_TIMELINE for aggregate metrics**

```javascript
function computeAggregateEffectiveness(ss, config) {
  var sheet = ss.getSheetByName('PRODUCT_TIMELINE');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var idxScore = headers.indexOf('effectiveness_score');
  var idxDeltaCost = headers.indexOf('delta_cost_pct');
  var idxDeltaRoas = headers.indexOf('delta_roas_pct');
  if (idxScore < 0) return null;

  var counts = { '+': 0, '=': 0, '-': 0, 'PENDING': 0, 'N/A': 0 };
  var deltaCostSum = 0, deltaCostCount = 0;
  var deltaRoasSum = 0, deltaRoasCount = 0;

  for (var i = 0; i < data.length; i++) {
    var score = data[i][idxScore];
    if (counts[score] !== undefined) counts[score]++;

    if (score === '+' || score === '=' || score === '-') {
      var dc = parseFloat(data[i][idxDeltaCost]);
      var dr = parseFloat(data[i][idxDeltaRoas]);
      if (!isNaN(dc)) { deltaCostSum += dc; deltaCostCount++; }
      if (!isNaN(dr)) { deltaRoasSum += dr; deltaRoasCount++; }
    }
  }

  return {
    counts: counts,
    avgDeltaCost: deltaCostCount > 0 ? deltaCostSum / deltaCostCount : null,
    avgDeltaRoas: deltaRoasCount > 0 ? deltaRoasSum / deltaRoasCount : null,
    totalEvaluated: counts['+'] + counts['='] + counts['-']
  };
}
```

**Step 2: Add to writeSummaryTab EFFECTIVENESS section**

```javascript
    var aggEff = computeAggregateEffectiveness(ss, config);
    if (aggEff && aggEff.totalEvaluated > 0) {
      rows.push(['▸ EFFECTIVENESS (přes všechny resolved produkty)']);
      rows.push(['Produkty s "+" (úspěšná intervence):', aggEff.counts['+']]);
      rows.push(['Produkty s "=" (smíšený výsledek):', aggEff.counts['=']]);
      rows.push(['Produkty s "-" (intervence škodí):', aggEff.counts['-']]);
      rows.push(['PENDING (<14 dní od akce):', aggEff.counts['PENDING']]);
      rows.push(['N/A (akce nebyla zaznamenána):', aggEff.counts['N/A']]);
      if (aggEff.avgDeltaCost !== null) {
        rows.push(['Avg delta cost (evaluated):', Utils.safePctFormat(aggEff.avgDeltaCost)]);
      }
      if (aggEff.avgDeltaRoas !== null) {
        rows.push(['Avg delta ROAS (evaluated):', Utils.safePctFormat(aggEff.avgDeltaRoas)]);
      }
      rows.push([]);
    }
```

**Step 3: Rebuild**

---

## PHASE 7: README + CONFIG doc update

### Task 7.1: Update README.md with new categories and tuning tips

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/README.md`

**Step 1: Add section "Nové kategorie (v2)"**

Přidat po existujícím popisu LOSER_REST / LOW_CTR:

```markdown
### RISING (v2) — growth momentum

Produkty s revenue růstem ≥ 50% vs předchozí stejně dlouhý lookback. Min 3 konverze v obou periodách. **Akce:** early scaling — zvýšit budget, vydělit do vlastní asset group.

### DECLINING (v2) — early warning

Produkty s poklesem revenue ≥ 30% vs předchozí lookback. **Akce:** investigate — cena vs konkurence, sklad, sezonnost.

### LOST_OPPORTUNITY (v2) — skalovatelné

Rentabilní produkty (conv ≥ 5, PNO ≤ target×0.8) s nízkým Impression Share (<0.5). **Akce:** zvýšit bid, top-priority kampaň.
```

**Step 2: Update CONFIG parameters table**

Přidat sekci s novými parametry (risingGrowthThreshold, atd.).

**Step 3: Update tuning tipy**

Rozšířit o:
- Více RISING flagů: sniž `risingGrowthThreshold` na 25
- Míň DECLINING: zvyš `decliningDropThreshold` na 50
- Víc LOST_OPPORTUNITY: zvyš `lostOpportunityMaxImpressionShare` na 0.7

**Step 4: Rebuild (jen README, ne combined.gs)**

---

### Task 7.2: Update README tab in sheet (setupOutputSheet)

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/main.gs`

**Step 1: Add new categories description in readmeContent**

Přidat sekci "▸ KATEGORIE v2" s popisem RISING/DECLINING/LOST_OPP.

**Step 2: Update "▸ TABY" to include new tabs**

Přidat ACTIONS, PRODUCT_TIMELINE, WEEKLY_SNAPSHOT.

**Step 3: Rebuild**

---

## PHASE 8: Testing

### Task 8.1: Run all unit tests

**Step 1: Run**

```bash
cd /tmp/loser-test && node test-runner.js 2>&1 | tail -30
```

Expected: All 46 original + ~11 new = 57 tests passing.

**Step 2: If failures, fix classifiers**

Debug in classifier.gs based on diff output.

---

### Task 8.2: Dry-run on kabelka.cz

**Step 1: Copy combined.gs to Google Ads UI**

In user's script for kabelka.cz — paste combined.gs.

**Step 2: Set CONFIG**

- `outputSheetId: ''` (or new sheet for v2)
- `dryRun: true`
- Rest as defaults (will be set from CONFIG tab).

**Step 3: Run Preview**

**Step 4: Verify logs**

Expected log messages:
- `Loaded search_impression_share for X products`
- `Previous period: X rows loaded`
- `Aggregated X unique item_ids with split metrics`
- `Classification funnel...`
- Flag counts reasonable (based on MCP validation)

**Step 5: If OK → set dryRun=false → Run live**

---

### Task 8.3: Verify ACTIONS tab on Kabelka

**Step 1: Open sheet ACTIONS tab**

**Step 2: Verify schema**

- Headers present: priority_rank, category, item_id, ..., action_taken, action_date, consultant_note
- Rows sorted by priority_rank asc
- Category colors applied

**Step 3: Spot-check 5 random rows vs MCP**

For 5 random flagged products, verify via MCP:
- main_cost, main_conv matches MCP query
- total_cost includes brand + rest

---

### Task 8.4: Test manual input preservation (2 runs)

**Step 1: Manually edit ACTIONS tab**

- Pick 3 products, fill `action_taken = "label applied"`, `action_date = today`, `consultant_note = "test"`.

**Step 2: Run script again**

**Step 3: Verify preservation**

- Open ACTIONS: those 3 products still have manual values filled.
- Open PRODUCT_TIMELINE: `latest_action`, `latest_action_date`, `latest_note` for those products = manual values.

---

### Task 8.5: Validate via MCP — RISING/DECLINING accuracy

**Step 1: From sheet, pick 5 RISING products**

**Step 2: For each, query MCP**

```
SELECT segments.product_item_id, metrics.conversions, metrics.conversions_value
FROM shopping_performance_view
WHERE segments.product_item_id = 'X'
  AND segments.date BETWEEN 'current_start' AND 'current_end'
```

Plus same for previous period.

**Step 3: Verify growth_pct matches script calculation ± 5%**

---

### Task 8.6: Update learning log

**Files:**
- Modify: `/Users/matousnovy/Documents/PPC/znalosti/_learning-log.md`

**Step 1: Add entry**

```markdown
### [2026-04-21] [Skripty/GAS] Shopping/PMAX Loser Detector v2 — major redesign
**Klient:** obecné (první deploy na kabelka.cz)
**Kontext:** Skript v1 byl "labeling tool". V2 rozšířen o pozitivní insights (RISING, LOST_OPPORTUNITY), early warnings (DECLINING), per-produkt historii s effectiveness tracking, split scope metrics (main vs brand vs rest), a semi-manuální interventions logging.
**Zjištění:**
- Shopping_product_view je jediný zdroj search_impression_share na per-produkt úrovni
- Brand kampaně musí být zahrnuty do insights (konverze je konverze), ale vyloučeny z LOSER klasifikace (= optimalizace main kampaní)
- Preserve manual columns vyžaduje read-before-write pattern
- Weekly snapshot tab umožňuje WoW/MoM trend analýzu bez drahých queries
**Akce:** V2 deployed. Design doc: /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/docs/2026-04-21-v2-redesign-design.md. Implementation plan: same folder /plans/.
**Poznatky pro budoucí skripty:**
- Split scope patterny (main/brand/rest) jsou užitečný abstraction pro každý Shopping analyzer
- Semi-manuální tracking (automatic + manual note columns) je robustní způsob engagement konzultanta
```

---

## Completion criteria

- [ ] All 8 phases completed
- [ ] Unit tests: 57+ passing
- [ ] Dry-run on kabelka.cz: no errors, reasonable flag counts
- [ ] Live run on kabelka.cz: 7+ tabs populated correctly
- [ ] Manual input preservation verified (2 runs)
- [ ] MCP data accuracy validated for 5 random products
- [ ] README.md + README tab updated
- [ ] Learning log entry added

**Target:** 40-60 total flagged products (all 5 categories) for kabelka.cz with new defaults.
