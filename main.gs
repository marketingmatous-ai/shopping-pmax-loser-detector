/**
 * main.gs — Entry point pro Shopping/PMAX Loser Detector
 *
 * Obsahuje main() orchestraci, setupOutputSheet() helper a buildSummary().
 * CONFIG objekt je v _config.gs — prvni modul v combined.gs (viditelny hned
 * po otevreni v Google Ads Scripts editoru, aby ho user nemusel hledat).
 */

/**
 * Main entry point. Google Ads spousti tuto funkci.
 *
 * AUTO-SETUP: pokud CONFIG.outputSheetId je prazdny, skript automaticky
 * spousti setupOutputSheet() — vytvori novy Google Sheet ve tvem Drive.
 * Po dokonceni zkopiruj ID z logu do CONFIG.outputSheetId a pust main() znovu.
 */
function main() {
  var runDate = new Date();

  try {
    Logger.log('========================================');
    Logger.log('Shopping/PMAX Loser Detector — Run start');
    Logger.log('Date: ' + Utils.formatDate(runDate));
    Logger.log('========================================');

    // === AUTO-SETUP: pokud outputSheetId neni nastaveny, vytvor sheet a skonci ===
    if (!CONFIG.outputSheetId || String(CONFIG.outputSheetId).length < 20) {
      Logger.log('INFO: CONFIG.outputSheetId neni nastaveny — spoustim AUTO-SETUP mode.');
      Logger.log('');
      setupOutputSheet();
      Logger.log('');
      Logger.log('========================================');
      Logger.log('AUTO-SETUP dokoncen.');
      Logger.log('DALSI KROK:');
      Logger.log('  1. Zkopiruj ID vyse (radek "ID: ...")');
      Logger.log('  2. Paste do CONFIG.outputSheetId v tomto skriptu');
      Logger.log('  3. Save a pust main() znovu');
      Logger.log('========================================');
      return;
    }

    // === 1. VALIDACE CONFIGU ===
    // POZN: Drive CONFIG je VZDY zdroj pravdy = CONFIG objekt v tomto skriptu (nahore).
    // Sheet CONFIG tab je pouze informativni (read-only snapshot runtime hodnot).
    // Output.refreshConfigTab() zapise aktualne pouzite hodnoty do sheet tabu po kazdem runu.
    Config.validate(CONFIG);

    // === 2. NACTENI DAT ===
    var data = DataLayer.fetchAllData(CONFIG);

    if (!data.accountBaseline || data.accountBaseline.totalClicks === 0) {
      throw new Error('Account nema zadna Shopping/PMAX data v lookback period. Zkontroluj, ze (1) Shopping/PMAX kampane jsou aktivni, (2) lookbackDays neni prilis male, (3) datum account neni prilis novy.');
    }

    // === 3. NACTENI LIFECYCLE HISTORIE (pro transitions) ===
    var previousLifecycleMap = {};
    if (CONFIG.enableHistoryDedup && !CONFIG.dryRun) {
      previousLifecycleMap = DataLayer.fetchPreviousLifecycle(CONFIG.outputSheetId);
      Logger.log('INFO: Nacteno ' + Object.keys(previousLifecycleMap).length + ' previous lifecycle entries.');
    }

    // === NACTENI EXISTUJICICH ACTIONS A PRODUCT_TIMELINE (preserve manual columns) ===
    var existingActions = {};
    var existingTimeline = {};
    if (!CONFIG.dryRun) {
      existingActions = DataLayer.readExistingActions(CONFIG.outputSheetId);
      existingTimeline = DataLayer.readExistingProductTimeline(CONFIG.outputSheetId);
    }

    // === 4. KLASIFIKACE KAZDEHO PRODUKTU ===
    Logger.log('INFO: Zacinam klasifikaci ' + data.products.length + ' produktu...');
    var classified = [];
    var funnelStats = {
      rawRows: data.productsRawRowCount || 0,             // Pred filtrem brand/rest/paused
      brandExcluded: data.excludedCounts.brand,
      restExcluded: data.excludedCounts.rest,
      pausedExcluded: data.excludedCounts.paused,
      keptRows: data.productsKeptRowCount || 0,           // Po filtru, pred agregaci
      afterAggregation: data.products.length,             // Unikatni item_id po agregaci
      tooYoung: 0,
      insufficientData: 0,
      dataQualityIssues: 0,
      classified: 0
    };

    for (var i = 0; i < data.products.length; i++) {
      var product = data.products[i];
      var firstClickDate = data.firstClickDates[product.itemId] || null;
      var prevLifecycle = previousLifecycleMap[product.itemId] || null;

      var result = Classifier.classifyProduct(
        product,
        CONFIG,
        data.accountBaseline,
        data.perCampaignBaseline,
        firstClickDate,
        data.lastYearStats,
        prevLifecycle,
        runDate,
        data.productPrices
      );

      // Funnel stats
      if (result.status === 'NEW_PRODUCT_RAMP_UP') {
        funnelStats.tooYoung++;
      } else if (result.status === 'INSUFFICIENT_DATA') {
        funnelStats.insufficientData++;
      } else if (result.status === 'DATA_QUALITY_ISSUE') {
        funnelStats.dataQualityIssues++;
      } else {
        funnelStats.classified++;
      }

      classified.push(result);
    }

    // === 5. EFFECTIVENESS KPIs ===
    var effectiveness = Classifier.computeEffectiveness(classified, previousLifecycleMap, null);

    // === 6. SUMMARY AGREGAT ===
    var summary = buildSummary(classified, data, funnelStats, runDate);

    // === 7. LOG PREHLED ===
    logSummary(summary, effectiveness);

    if (CONFIG.dryRun) {
      Logger.log('INFO: DRY RUN — nic se nezapsalo do sheetu.');
      Logger.log('INFO: Classified vysledku: ' + classified.length);
      Logger.log('INFO: Flagged (primary_label nenil): ' + summary.flags.totalFlagged);
      return;
    }

    // === 8. ZAPIS DO SHEETU ===
    Output.writeAll(CONFIG.outputSheetId, classified, summary, effectiveness, CONFIG, runDate, existingActions, existingTimeline);

    // === 9. EMAIL REPORT ===
    var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.outputSheetId;
    Output.sendEmailReport(CONFIG, summary, effectiveness, sheetUrl);

    Logger.log('========================================');
    Logger.log('Run dokoncen uspesne.');
    Logger.log('========================================');
  } catch (e) {
    Logger.log('FATAL ERROR: ' + e.message);
    Logger.log('Stack: ' + (e.stack || '(no stack)'));

    // Pokud je configured admin email, poslat error alert
    if (CONFIG.adminEmail) {
      try {
        var emails = Utils.validateEmails(CONFIG.adminEmail);
        if (emails.length > 0) {
          MailApp.sendEmail(
            emails.join(','),
            '[ALERT] Shopping/PMAX Loser Detector FAILED',
            'Skript selhal s chybou:\n\n' +
            e.message + '\n\n' +
            'Stack:\n' + (e.stack || '(no stack)') + '\n\n' +
            'Check Google Ads Scripts logs pro vice detailu.'
          );
        }
      } catch (emailErr) {
        Logger.log('WARN: Nelze odeslat error email: ' + emailErr.message);
      }
    }

    throw e; // Re-throw, aby Google Ads zaznamenal failure
  }
}

/**
 * Sestavi summary objekt pro SUMMARY tab a email.
 */
function buildSummary(classified, data, funnelStats, runDate) {
  var flags = {
    totalFlagged: 0,
    loserRestTotal: 0,
    lowCtrTotal: 0,
    overlap: 0,
    loserByTier: { zero_conv: 0, low_volume: 0, mid_volume: 0, high_volume: 0 },
    lowCtrByReason: { irrelevant_keyword_match: 0, high_visibility_low_appeal: 0, low_ctr_general: 0 },
    totalWastedSpend: 0,
    wastedSpendPctOfTotal: 0
  };

  // Agregaty pro vizualni shrnuti
  var flaggedTotalCost = 0;
  var flaggedTotalConversions = 0;
  var flaggedTotalConvValue = 0;
  var flaggedTotalClicks = 0;
  var flaggedTotalImpressions = 0;

  for (var i = 0; i < classified.length; i++) {
    var c = classified[i];
    if (!c.primaryLabel) {
      continue;
    }
    flags.totalFlagged++;
    flags.totalWastedSpend += c.wastedSpend || 0;
    flaggedTotalCost += c.cost || 0;
    flaggedTotalConversions += c.conversions || 0;
    flaggedTotalConvValue += c.conversionValue || 0;
    flaggedTotalClicks += c.clicks || 0;
    flaggedTotalImpressions += c.impressions || 0;

    if (c.primaryLabel === CONFIG.labelLoserRestValue) {
      flags.loserRestTotal++;
      if (c.tier && flags.loserByTier[c.tier] !== undefined) {
        flags.loserByTier[c.tier]++;
      }
    } else if (c.primaryLabel === CONFIG.labelLowCtrValue) {
      flags.lowCtrTotal++;
      if (c.reasonCode && flags.lowCtrByReason[c.reasonCode] !== undefined) {
        flags.lowCtrByReason[c.reasonCode]++;
      }
    }

    if (c.secondaryFlags && c.secondaryFlags.length > 0) {
      flags.overlap++;
    }
  }

  flags.wastedSpendPctOfTotal = data.accountBaseline.totalCost > 0
    ? (flags.totalWastedSpend / data.accountBaseline.totalCost) * 100
    : 0;

  // Top losers + top low-CTR
  var losers = [];
  var lowCtrs = [];
  for (var j = 0; j < classified.length; j++) {
    var r = classified[j];
    if (r.primaryLabel === CONFIG.labelLoserRestValue) {
      losers.push(r);
    } else if (r.primaryLabel === CONFIG.labelLowCtrValue) {
      lowCtrs.push(r);
    }
  }
  var topLosers = Utils.sortDesc(losers, function (x) { return x.wastedSpend; }).slice(0, 10);
  var topLowCtr = Utils.sortDesc(lowCtrs, function (x) { return x.impressions; }).slice(0, 10);

  var accountInfo = getAccountInfo();

  // Doplnit lookbackDays do funnelStats (pro extrapolace)
  funnelStats.lookbackDays = CONFIG.lookbackDays;

  // === NEW PRODUCTS STATISTICS (age-failed classification = recently added) ===
  var newProducts = computeNewProductsSummary(classified);

  return {
    accountName: accountInfo.name,
    customerId: accountInfo.customerId,
    currency: data.currency,
    runDate: runDate,
    lookbackStart: data.lookbackStart,
    lookbackEnd: data.lookbackEnd,
    accountBaseline: data.accountBaseline,
    perCampaignBaseline: data.perCampaignBaseline,
    funnel: funnelStats,
    flags: flags,
    topLosers: topLosers,
    topLowCtr: topLowCtr,
    flaggedTotalCost: flaggedTotalCost,
    flaggedTotalConversions: flaggedTotalConversions,
    flaggedTotalConvValue: flaggedTotalConvValue,
    flaggedTotalClicks: flaggedTotalClicks,
    flaggedTotalImpressions: flaggedTotalImpressions,
    newProducts: newProducts,
    config: {
      minProductAgeDays: CONFIG.minProductAgeDays,
      targetPnoPct: CONFIG.targetPnoPct
    }
  };
}

/**
 * Agregate statistika novych produktu (status=NEW_PRODUCT_RAMP_UP).
 * Identifikuje "rising star candidates" — mlade produkty s prvnimi signalmi uspechu.
 */
function computeNewProductsSummary(classified) {
  var total = 0;
  var ageSum = 0;
  var ageMin = null;
  var ageMax = null;
  var withGmcFeedPrice = 0;
  var withClicks10Plus = 0;
  var withFirstConv = 0;
  var totalNewCost = 0;
  var totalNewConv = 0;
  var totalNewConvValue = 0;
  var totalNewClicks = 0;
  var totalNewImpressions = 0;

  // Rising star candidates: mladé + 10+ clicks + 1+ conv (sorted by ROAS desc)
  var risingCandidates = [];

  for (var i = 0; i < classified.length; i++) {
    var c = classified[i];
    if (c.status !== 'NEW_PRODUCT_RAMP_UP') continue;

    total++;
    if (c.ageDays !== null && c.ageDays !== undefined) {
      ageSum += c.ageDays;
      if (ageMin === null || c.ageDays < ageMin) ageMin = c.ageDays;
      if (ageMax === null || c.ageDays > ageMax) ageMax = c.ageDays;
    }
    if (c.priceSource === 'gmc_feed') withGmcFeedPrice++;
    if ((c.clicks || 0) >= 10) withClicks10Plus++;
    if ((c.conversions || 0) >= 1) withFirstConv++;

    totalNewCost += c.cost || 0;
    totalNewConv += c.conversions || 0;
    totalNewConvValue += c.conversionValue || 0;
    totalNewClicks += c.clicks || 0;
    totalNewImpressions += c.impressions || 0;

    if ((c.clicks || 0) >= 10 && (c.conversions || 0) >= 1) {
      risingCandidates.push({
        itemId: c.itemId,
        productTitle: c.productTitle,
        ageDays: c.ageDays,
        clicks: c.clicks,
        conversions: c.conversions,
        cost: c.cost,
        conversionValue: c.conversionValue,
        roas: c.roas,
        actualPno: c.actualPno,
        priceSource: c.priceSource,
        productPrice: c.productPrice
      });
    }
  }

  // Sort rising candidates by ROAS desc
  risingCandidates = Utils.sortDesc(risingCandidates, function (x) { return x.roas || 0; });
  var topRising = risingCandidates.slice(0, 10);

  return {
    total: total,
    avgAge: total > 0 ? ageSum / total : 0,
    minAge: ageMin,
    maxAge: ageMax,
    withGmcFeedPrice: withGmcFeedPrice,
    withClicks10Plus: withClicks10Plus,
    withFirstConv: withFirstConv,
    totalCost: totalNewCost,
    totalConversions: totalNewConv,
    totalConversionValue: totalNewConvValue,
    totalClicks: totalNewClicks,
    totalImpressions: totalNewImpressions,
    topRisingCandidates: topRising
  };
}

/**
 * Loguje summary do Google Ads Scripts logs.
 */
function logSummary(summary, effectiveness) {
  Logger.log('========== SUMMARY ==========');
  Logger.log('Account: ' + summary.accountName + ' (' + summary.customerId + ')');
  Logger.log('Lookback: ' + Utils.formatDate(summary.lookbackStart) + ' — ' + Utils.formatDate(summary.lookbackEnd));
  Logger.log('Account cost: ' + summary.accountBaseline.totalCost + ' ' + summary.currency);
  Logger.log('Account CVR: ' + Utils.safePctFormat(summary.accountBaseline.cvr * 100));
  Logger.log('');
  Logger.log('FUNNEL (campaign bucket split — brand/rest se OddELuji, ne vylucuji):');
  Logger.log('  raw rows (vse):          ' + summary.funnel.rawRows);
  Logger.log('  separated to brand:      ' + summary.funnel.brandExcluded + ' (do brand_metrics, zachovano)');
  Logger.log('  separated to rest:       ' + summary.funnel.restExcluded + ' (do rest_metrics, zachovano)');
  Logger.log('  ignored paused:          ' + summary.funnel.pausedExcluded);
  Logger.log('  kept for classification (main): ' + ((summary.funnel.keptRows !== undefined) ? summary.funnel.keptRows : 'N/A'));
  Logger.log('  unique item_ids after agg: ' + ((summary.funnel.afterAggregation !== undefined) ? summary.funnel.afterAggregation : 'N/A'));
  Logger.log('');
  Logger.log('FUNNEL (per-product gates):');
  Logger.log('  young (< ' + summary.config.minProductAgeDays + ' dni):    ' + summary.funnel.tooYoung);
  Logger.log('  insufficient data:       ' + summary.funnel.insufficientData);
  Logger.log('  data quality issues:     ' + summary.funnel.dataQualityIssues);
  Logger.log('  classified:              ' + summary.funnel.classified);
  var funnelSum = (summary.funnel.tooYoung || 0) + (summary.funnel.insufficientData || 0) + (summary.funnel.dataQualityIssues || 0) + (summary.funnel.classified || 0);
  Logger.log('  SUM (should match unique item_ids): ' + funnelSum);
  Logger.log('');
  Logger.log('FLAGGED: total=' + summary.flags.totalFlagged + ', loser=' + summary.flags.loserRestTotal + ', low_ctr=' + summary.flags.lowCtrTotal + ', overlap=' + summary.flags.overlap);
  Logger.log('Wasted spend: ' + summary.flags.totalWastedSpend + ' ' + summary.currency + ' (' + Utils.safePctFormat(summary.flags.wastedSpendPctOfTotal) + ' total cost)');

  if (effectiveness) {
    Logger.log('');
    Logger.log('TRANSITIONS: NEW=' + (effectiveness.transitions.NEW_FLAG || 0) + ', REPEATED=' + (effectiveness.transitions.REPEATED || 0) + ', RESOLVED=' + (effectiveness.transitions.RESOLVED || 0) + ', RE_FLAGGED=' + (effectiveness.transitions.RE_FLAGGED || 0));
    if (effectiveness.applicationRate !== null) {
      Logger.log('Label application rate: ' + Utils.safePctFormat(effectiveness.applicationRate));
    }
  }

  if (summary.newProducts && summary.newProducts.total > 0) {
    var np = summary.newProducts;
    Logger.log('');
    Logger.log('NOVE PRODUKTY (< ' + summary.config.minProductAgeDays + ' dni, neklasifikovane):');
    Logger.log('  total: ' + np.total + ', avg age: ' + (np.avgAge || 0).toFixed(1) + ' dni (range: ' + np.minAge + '-' + np.maxAge + ')');
    Logger.log('  s feed cenou: ' + np.withGmcFeedPrice + '/' + np.total);
    Logger.log('  s 10+ kliky:  ' + np.withClicks10Plus);
    Logger.log('  s 1+ konv:    ' + np.withFirstConv + ' (rising star candidates)');
    if (np.topRisingCandidates && np.topRisingCandidates.length > 0) {
      Logger.log('  Top 3 rising star: ' + np.topRisingCandidates.slice(0, 3).map(function (r) {
        return r.itemId + ' (ROAS ' + (r.roas || 0).toFixed(2) + ', age ' + r.ageDays + 'd)';
      }).join(', '));
    }
  }
  Logger.log('=============================');
}

/**
 * Vraci zakladni info o aktualnim accountu.
 */
function getAccountInfo() {
  try {
    var ci = AdsApp.currentAccount();
    return {
      name: ci.getName() || '(unnamed)',
      customerId: ci.getCustomerId() || '(unknown)'
    };
  } catch (e) {
    return { name: '(unknown)', customerId: '(unknown)' };
  }
}

/**
 * JEDNORAZOVY SETUP — vytvori novy Google Sheet s prepravenou strukturou
 * (4 taby, headers, vysvetleni) a vypise do logu jeho ID + URL.
 *
 * Spusteni:
 *   1. V Google Ads Scripts UI vyber z dropdownu tuto funkci (setupOutputSheet)
 *      misto main() a klikni "Run" (ne Preview — zapisuje novy soubor).
 *   2. V logu najdes URL + ID noveho sheetu.
 *   3. Zkopiruj ID do CONFIG.outputSheetId v main.gs.
 *   4. Otevri sheet v Drive a poprve mu nastav sdileni (pokud potrebujes).
 *   5. Prepni na main() a pokracuj beznym workflow.
 *
 * Sheet se vytvori pod Google accountem, pod kterym skript bezi
 * (= tvoje osobni Drive slozka "My Drive").
 */
/**
 * Sestavi rows pro CONFIG tab z CONFIG objektu (runtime values).
 * Pouziva se v setupOutputSheet (prvni run) i v kazdem runu (refresh).
 */
function buildConfigTabRows(cfg) {
  return [
    ['Parametr', 'Hodnota', 'Popis'],
    ['ZÁKLADNÍ', '', ''],
    ['targetPnoPct', cfg.targetPnoPct, 'Cílové PNO v % (např. 25 = 25% nákladů z revenue)'],
    ['lookbackDays', cfg.lookbackDays, 'Okno analýzy ve dnech (7–365)'],
    ['adminEmail', cfg.adminEmail || '', 'Email pro notifikace (volitelné)'],
    ['', '', ''],
    ['LABELY (do custom_label_N v GMC / Mergado)', '', ''],
    ['customLabelIndex', cfg.customLabelIndex, 'Číslo labelu 0–4 (co klient nepoužívá v GMC)'],
    ['labelLoserRestValue', cfg.labelLoserRestValue, 'Hodnota zapsaná do custom_label pro losery'],
    ['labelLowCtrValue', cfg.labelLowCtrValue, 'Hodnota zapsaná do custom_label pro low-CTR'],
    ['labelHealthyValue', cfg.labelHealthyValue || '(vypnuto)', 'Hodnota pro zdravé produkty (status=ok, bez flagu). "" = nezapisovat'],
    ['', '', ''],
    ['CAMPAIGN BUCKET SPLIT (regex v názvu kampaně)', '', ''],
    ['brandCampaignPattern', cfg.brandCampaignPattern, 'Brand kampaně → brand_metrics bucket (zachováno pro insights, NE v klasifikaci)'],
    ['restCampaignPattern', cfg.restCampaignPattern, 'Rest kampaně → rest_metrics bucket (pro RESOLVED tracking, NE v klasifikaci)'],
    ['analyzeChannels', (cfg.analyzeChannels || []).join(','), 'Typy kampaní (comma-separated)'],
    ['', '', ''],
    ['SAMPLE SIZE GATE (ochrana proti false-positive)', '', ''],
    ['minClicksAbsolute', cfg.minClicksAbsolute, 'Absolutní minimum kliků před klasifikací'],
    ['minExpectedConvFloor', cfg.minExpectedConvFloor, 'Min expected conversions (clicks × account CVR)'],
    ['', '', ''],
    ['RISING STAR PROTECTION', '', ''],
    ['minProductAgeDays', cfg.minProductAgeDays, 'Produkty mladší N dní se neevaluují'],
    ['', '', ''],
    ['LOSER TIER THRESHOLDS', '', ''],
    ['tierLowVolumeMax', cfg.tierLowVolumeMax, 'Conv 1–N = low volume tier'],
    ['tierMidVolumeMax', cfg.tierMidVolumeMax, 'Conv N+1 až M = mid volume tier'],
    ['pnoMultiplierZeroConv', cfg.pnoMultiplierZeroConv, '0 conv: spend ≥ X × expected CPA'],
    ['pnoMultiplierLowVol', cfg.pnoMultiplierLowVol, '1–3 conv: PNO ≥ X × target'],
    ['pnoMultiplierMidVol', cfg.pnoMultiplierMidVol, '4–10 conv: PNO ≥ X × target'],
    ['pnoMultiplierHighVol', cfg.pnoMultiplierHighVol, '11+ conv: PNO ≥ X × target (chrání volume)'],
    ['', '', ''],
    ['LOW CTR DETEKCE', '', ''],
    ['ctrBaselineScope', cfg.ctrBaselineScope, '"account" nebo "campaign"'],
    ['minImpressionsLowCtr', cfg.minImpressionsLowCtr, 'Min impressions za lookback (hlavní floor, 300-1000 podle velikosti účtu)'],
    ['minClicksLowCtr', cfg.minClicksLowCtr, 'Min kliků (0 = bez limitu; zvyš pokud dostáváš false positives)'],
    ['ctrThresholdMultiplier', cfg.ctrThresholdMultiplier, 'Produkt s CTR < X × baseline = flag (doporučeno 0.6–0.8)'],
    ['lowCtrSkipIfProfitableMinConv', cfg.lowCtrSkipIfProfitableMinConv, 'Min konverzí pro skip rentabilních (produkt s N+ konv a PNO<=target*1.1 neflaggovat)'],
    ['', '', ''],
    ['▸ TREND DETECTION (RISING/DECLINING)', '', ''],
    ['risingGrowthThreshold', cfg.risingGrowthThreshold, 'Growth ≥ X% = RISING (default 50% medium)'],
    ['decliningDropThreshold', cfg.decliningDropThreshold, 'Drop ≥ X% = DECLINING (default 30% medium)'],
    ['minConversionsForTrendCompare', cfg.minConversionsForTrendCompare, 'Min conversions v obou periodách'],
    ['', '', ''],
    ['▸ LOST_OPPORTUNITY', '', ''],
    ['lostOpportunityMinConv', cfg.lostOpportunityMinConv, 'Min conversions pro rentability claim'],
    ['lostOpportunityMaxPnoMultiplier', cfg.lostOpportunityMaxPnoMultiplier, 'PNO ≤ target × N (výrazně rentabilní)'],
    ['lostOpportunityMaxImpressionShare', cfg.lostOpportunityMaxImpressionShare, 'IS < N (Google málo zobrazuje)'],
    ['', '', ''],
    ['▸ EFFECTIVENESS', '', ''],
    ['effectivenessMinDaysSinceAction', cfg.effectivenessMinDaysSinceAction, 'Dní před vyhodnocením účinnosti'],
    ['restCampaignEfficientThreshold', cfg.restCampaignEfficientThreshold, 'Rest cost ≤ N × before_main pro "efficient"'],
    ['', '', ''],
    ['POKROČILÉ', '', ''],
    ['enableYoYSeasonalityCheck', cfg.enableYoYSeasonalityCheck, 'YoY porovnání (pokud >1 rok dat)'],
    ['enableHistoryDedup', cfg.enableHistoryDedup, 'Tracking v LIFECYCLE_LOG tabu'],
    ['historyDedupDays', cfg.historyDedupDays, 'Dedup okno v dnech'],
    ['dryRun', cfg.dryRun, 'TRUE = jen loguje, nezapisuje do sheetu'],
    ['groupByParentId', cfg.groupByParentId, 'TRUE = agreguj varianty na parent_id'],
    ['includeProductTitles', cfg.includeProductTitles, 'TRUE = zapsat product_title do DETAIL tabu'],
    ['maxRowsDetailTab', cfg.maxRowsDetailTab, 'Sheet size guard']
  ];
}

function setupOutputSheet() {
  var accountInfo = getAccountInfo();
  var sheetName = 'Shopping-PMAX Loser Detector — ' + accountInfo.name + ' (' + accountInfo.customerId + ')';

  var ss = SpreadsheetApp.create(sheetName);
  var ssId = ss.getId();
  var ssUrl = ss.getUrl();

  // === Tab 1: FEED_UPLOAD ===
  var feedSheet = ss.getSheets()[0];
  feedSheet.setName('FEED_UPLOAD');
  feedSheet.getRange(1, 1, 1, 2).setValues([['id', 'custom_label_X']]);
  feedSheet.getRange(2, 1).setValue('(naplni se po prvnim runu main())');
  feedSheet.setFrozenRows(1);

  // === Tab 2: DETAIL ===
  var detailSheet = ss.insertSheet('DETAIL');
  detailSheet.getRange(1, 1).setValue('(naplni se po prvnim runu main() — vsechny flagged produkty s full trace dat)');

  // === Tab 3: SUMMARY ===
  var summarySheet = ss.insertSheet('SUMMARY');
  summarySheet.getRange(1, 1).setValue('(naplni se po prvnim runu main() — dashboard, effectiveness KPIs, config audit trail)');

  // === Tab 4: LIFECYCLE_LOG ===
  var lifecycleHeaders = [
    'run_date',
    'item_id',
    'current_label',
    'previous_label',
    'transition_type',
    'current_campaign',
    'previous_campaign',
    'campaign_moved',
    'cost_30d',
    'conversions_30d',
    'pno_30d_pct',
    'roas_30d',
    'ctr_30d_pct',
    'reason_code',
    'tier',
    'runs_since_first_flag',
    'notes'
  ];
  var lifecycleSheet = ss.insertSheet('LIFECYCLE_LOG');
  lifecycleSheet.getRange(1, 1, 1, lifecycleHeaders.length).setValues([lifecycleHeaders]);
  lifecycleSheet.setFrozenRows(1);

  // === ACTIONS tab (placeholder — populated on first main() run) ===
  var actionsSheet = ss.insertSheet('ACTIONS');
  actionsSheet.getRange(1, 1).setValue('(naplni se po prvnim runu main())');

  // === PRODUCT_TIMELINE tab (placeholder) ===
  var timelineSheet = ss.insertSheet('PRODUCT_TIMELINE');
  timelineSheet.getRange(1, 1).setValue('(naplni se po prvnim runu main() — per-produkt historie)');

  // === WEEKLY_SNAPSHOT tab (placeholder) ===
  var snapshotSheet = ss.insertSheet('WEEKLY_SNAPSHOT');
  snapshotSheet.getRange(1, 1).setValue('(naplni se po kazdem runu — 1 radek per tyden pro trendy)');

  // === MONTHLY_REVIEW tab (placeholder) ===
  var monthlyReviewSheet = ss.insertSheet('MONTHLY_REVIEW');
  monthlyReviewSheet.getRange(1, 1).setValue('(naplni se po kazdem runu — mesicni agregace z LIFECYCLE_LOG)');

  // === Tab 5: CONFIG (zobrazuje AKTUALNI hodnoty z CONFIG objektu v kodu) ===
  // POZN: Tento tab je pouze informacni — editace hodnot zde NEMA zadny ucinek.
  // Zdroj pravdy je CONFIG objekt na zacatku tohoto skriptu (editor v Google Ads).
  // Tab se automaticky aktualizuje pri kazdem runu (viz Output.refreshConfigTab).
  var configSheet = ss.insertSheet('CONFIG');
  var configRows = buildConfigTabRows(CONFIG);
  configSheet.getRange(1, 1, configRows.length, 3).setValues(configRows);
  configSheet.setFrozenRows(1);
  configSheet.setColumnWidth(1, 240);
  configSheet.setColumnWidth(2, 200);
  configSheet.setColumnWidth(3, 520);

  // Header row
  configSheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1a73e8')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  configSheet.setRowHeight(1, 32);

  // Sekční headers — nacházíme je podle toho že sloupec B a C jsou prázdné
  for (var cr = 0; cr < configRows.length; cr++) {
    var label = configRows[cr][0];
    var val = configRows[cr][1];
    var desc = configRows[cr][2];
    // Sekce má neprázdný text v sloupci A ale prázdné hodnota i popis
    if (label && val === '' && desc === '' && cr > 0) {
      var sheetRow = cr + 1; // 1-indexed
      configSheet.getRange(sheetRow, 1, 1, 3)
        .setBackground('#e8f0fe')
        .setFontWeight('bold')
        .setFontColor('#1a73e8')
        .setFontSize(10);
    }
  }

  // Vertical borders + alternate row shading pro přehlednost
  configSheet.getRange(2, 1, configRows.length - 1, 3)
    .setBorder(null, true, null, true, true, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // Value column středové zarovnání + hodnoty bold
  configSheet.getRange(2, 2, configRows.length - 1, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // Popisy drobnějším fontem, šedé
  configSheet.getRange(2, 3, configRows.length - 1, 1)
    .setFontSize(9)
    .setFontColor('#555555');

  // === README tab (user-facing instrukce) ===
  var readmeSheet = ss.insertSheet('README');
  var selfCopyUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/copy';
  var readmeContent = [
    ['Shopping/PMAX Loser Detector — Output Sheet'],
    [''],
    ['Tento sheet je output interniho nastroje pro identifikaci neefektivnich produktu v Shopping a PMAX kampanich.'],
    [''],
    ['▸ POUZIT JAKO TEMPLATE PRO DALSIHO KLIENTA'],
    ['Klikni na tento odkaz → Google vytvori tvou vlastni kopii s plnym layoutem:'],
    [selfCopyUrl],
    ['Pak vloz ID kopie (cast mezi /d/ a /edit) do CONFIG.outputSheetId v skriptu.'],
    [''],
    ['▸ TABY'],
    ['FEED_UPLOAD', 'Ready-to-upload CSV pro GMC Supplemental Feed nebo Mergado import.'],
    ['DETAIL', 'Plny trace dat per produkt (raw metriky, gate trace, reason codes, transitions).'],
    ['SUMMARY', 'Dashboard — account baseline, funnel, flagged breakdown, IMPACT vizualizace, top 10.'],
    ['LIFECYCLE_LOG', 'Append-only timeline per produkt (transitions: NEW_FLAG/REPEATED/RESOLVED/RE_FLAGGED).'],
    ['ACTIONS', 'Vsechny flagged produkty v jednom filterable tabu, s manual sloupci action_taken/date/note.'],
    ['PRODUCT_TIMELINE', 'Per-produkt historie + effectiveness score + manual columns.'],
    ['WEEKLY_SNAPSHOT', 'Append-only weekly KPI pro trend analyzu v DASHBOARD.'],
    ['CONFIG', 'Editovatelne parametry (targetPnoPct, lookbackDays, thresholds atd.).'],
    [''],
    ['▸ KATEGORIE v2'],
    ['RISING', 'Revenue rust ≥ 50% vs predchozi lookback (min 3 conv v obou periodach). Akce: early scaling — zvysit budget, vydelit do vlastni asset group.'],
    ['DECLINING', 'Pokles revenue ≥ 30% vs predchozi lookback. Akce: investigate — cena vs konkurence, sklad, sezonnost.'],
    ['LOST_OPPORTUNITY', 'Rentabilni produkty (conv ≥ 5, PNO ≤ target×0.8) s nizkym Impression Share (<0.5). Akce: zvysit bid, top-priority kampan.'],
    [''],
    ['▸ EDITACE PARAMETRU'],
    ['Vsechny parametry jdou upravit v tabu CONFIG — skript je nacita pred kazdym runem.'],
    ['Po zmene hodnot v CONFIG tabu pust skript znovu (viz "Run Now" nize) nebo pockej na scheduled run.'],
    [''],
    ['▸ RUN NOW (tlacitko v tomto sheetu)'],
    ['Chces tlacitko "Run Now" primo zde v sheetu? Pridej bound Apps Script:'],
    ['  1. V tomto sheetu: Rozsireni (Extensions) → Apps Script'],
    ['  2. Vlozi nasledujici kod a ulozit (Cmd+S):'],
    [''],
    ['  function onOpen() {'],
    ['    SpreadsheetApp.getUi()'],
    ['      .createMenu("🔧 Loser Detector")'],
    ['      .addItem("Otevrit Google Ads Scripts editor", "openScriptEditor")'],
    ['      .addItem("Jak spustit run", "showHowToRun")'],
    ['      .addToUi();'],
    ['  }'],
    [''],
    ['  function openScriptEditor() {'],
    ['    var url = "https://ads.google.com/aw/bulk/scripts";'],
    ['    var html = "<script>window.open(\\"" + url + "\\", \\"_blank\\");google.script.host.close();</script>";'],
    ['    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(10).setHeight(10), "Otevirám...");'],
    ['  }'],
    [''],
    ['  function showHowToRun() {'],
    ['    SpreadsheetApp.getUi().alert("Run workflow",'],
    ['      "1. Otevri Google Ads Scripts UI.\\n" +'],
    ['      "2. Vyber skript Shopping PMAX Loser Detector.\\n" +'],
    ['      "3. Klikni Preview (dry-run test) nebo Run (production).\\n" +'],
    ['      "4. Pockej 3-8 min, pak obnov tento sheet.",'],
    ['      SpreadsheetApp.getUi().ButtonSet.OK);'],
    ['  }'],
    [''],
    ['  3. Ulozit, pak v menu sheetu uvidis polozku 🔧 Loser Detector.'],
    ['  4. Google Ads Scripts API neumoznuje spustit skript primo z jineho sheetu kvuli zabezpeceni.'],
    ['     Tlacitko otevre Google Ads Scripts editor, kde klines Preview nebo Run manualne.'],
    [''],
    ['  Scheduled run: Loser detector se standardne spousti kazdy tyden v pondeli 07:00 (CET) automaticky.'],
    ['  Nova data jsou ve sheetu do 10 minut po startu runu.'],
    [''],
    ['▸ WORKFLOW'],
    ['1. Skript beha weekly, vygeneruje flagged produkty do tabu FEED_UPLOAD.'],
    ['2. Konzultant exportuje FEED_UPLOAD jako CSV → upload do GMC Supplemental Feed nebo Mergado.'],
    ['3. custom_label se aplikuje na produkty v GMC.'],
    ['4. V Google Ads listing group ruzne rest kampane excluduje produkty s labelem "loser_rest".'],
    ['5. Po 14 dnech skript checkne LIFECYCLE_LOG a oznaci produkty jako RESOLVED pokud byly presunuty do rest.'],
    [''],
    ['▸ TUNING TIPY PER KLIENT (edit v CONFIG tabu)'],
    [''],
    ['🔼 VIC FLAGU (skript je prilis mirny):'],
    ['  • minImpressionsLowCtr: sniz na 300 (chytne i mensi produkty s nizsi viditelnosti)'],
    ['  • ctrThresholdMultiplier: zvys na 0.8 (chytne i mirny podprumer)'],
    ['  • pnoMultiplierLowVol: sniz z 1.5 na 1.3 (strict loser tier)'],
    [''],
    ['🔽 MENE FLAGU (false positives nebo konzervativnejsi vystup):'],
    ['  • minImpressionsLowCtr: zvys na 700-1000 (jen signifikantni produkty)'],
    ['  • ctrThresholdMultiplier: sniz na 0.5 (extremne nizke CTR)'],
    ['  • minClicksAbsolute: zvys na 50-100 (vic dat pro loser verdict)'],
    [''],
    ['💰 FLEXIBILNEJSI RENTABILITA (setri i produkty s malo konverzemi):'],
    ['  • lowCtrSkipIfProfitableMinConv: sniz na 1 nebo 2 (1 conv s nizkym PNO = skip)'],
    ['  • pnoMultiplierLowVol: zvys na 2.0 (low volume tier je volnejsi)'],
    [''],
    ['🏢 VELKY ACCOUNT (>50k SKU):'],
    ['  • lookbackDays: 60-90 (vic dat per produkt)'],
    ['  • minImpressionsLowCtr: 1000 (jen signifikantni produkty)'],
    [''],
    ['🏪 MALY ACCOUNT / NIZKE VOLUME:'],
    ['  • lookbackDays: 30 kvuli responsivnosti'],
    ['  • minImpressionsLowCtr: 300'],
    ['  • minClicksAbsolute: 20'],
    [''],
    ['▸ UPOZORNENI'],
    ['Tento sheet obsahuje technicka data pro PPC specialisty. Neklikej na "Share" a nesdilej s klienty primo.'],
    ['Pro klienta pripravuj zjednodusenou verzi (vytah + akcni kroky).'],
    [''],
    ['Vygenerovano: ' + (function () {
      var now = new Date();
      var pad = function (n) { return ('0' + n).slice(-2); };
      return Utils.formatDate(now) + ' ' + pad(now.getHours()) + ':' + pad(now.getMinutes());
    })()],
    ['Account: ' + accountInfo.name + ' (' + accountInfo.customerId + ')']
  ];

  // Pad rows to same length
  var maxCols = 2;
  for (var i = 0; i < readmeContent.length; i++) {
    while (readmeContent[i].length < maxCols) {
      readmeContent[i].push('');
    }
  }
  readmeSheet.getRange(1, 1, readmeContent.length, maxCols).setValues(readmeContent);
  readmeSheet.setColumnWidth(1, 280);
  readmeSheet.setColumnWidth(2, 760);

  // Stylovani: header (row 1) = large, bold
  readmeSheet.getRange(1, 1, 1, maxCols).merge()
    .setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8')
    .setFontFamily('Montserrat');
  readmeSheet.setRowHeight(1, 36);

  // Najdi radek s copy URL a udelej z nej hyperlink
  for (var rIdx = 0; rIdx < readmeContent.length; rIdx++) {
    if (readmeContent[rIdx][0] === selfCopyUrl) {
      var linkCell = readmeSheet.getRange(rIdx + 1, 1, 1, maxCols);
      linkCell.merge();
      var linkText = '🔗 ' + selfCopyUrl;
      var linkRt = SpreadsheetApp.newRichTextValue()
        .setText(linkText)
        .setLinkUrl(0, linkText.length, selfCopyUrl)
        .build();
      linkCell.setRichTextValue(linkRt)
        .setFontColor('#1a73e8').setFontWeight('bold').setFontSize(11)
        .setBackground('#e8f0fe').setHorizontalAlignment('center');
      readmeSheet.setRowHeight(rIdx + 1, 32);
    }

    // Section headery (zacinaji "▸ ")
    if (typeof readmeContent[rIdx][0] === 'string' && readmeContent[rIdx][0].indexOf('▸') === 0) {
      readmeSheet.getRange(rIdx + 1, 1, 1, maxCols).merge()
        .setFontSize(12).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#1a73e8').setFontFamily('Montserrat')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
      readmeSheet.setRowHeight(rIdx + 1, 28);
    }
  }

  // Přesuň README jako první tab
  ss.setActiveSheet(readmeSheet);
  ss.moveActiveSheet(1);

  // Log info
  var copyUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/copy';
  Logger.log('========================================================');
  Logger.log('NOVY OUTPUT SHEET VYTVOREN');
  Logger.log('========================================================');
  Logger.log('Nazev: ' + sheetName);
  Logger.log('ID:    ' + ssId);
  Logger.log('URL:   ' + ssUrl);
  Logger.log('');
  Logger.log('📋 SHARE URL PRO KOPII (pro nasazeni na dalsi klienty jako template):');
  Logger.log('  ' + copyUrl);
  Logger.log('  (Kdo klikne na tento odkaz, dostane Google vyzvu "Vytvorit kopii".)');
  Logger.log('');
  Logger.log('DALSI KROK:');
  Logger.log('  1. Zkopiruj ID vyse (dlouhy retezec po /d/ v URL).');
  Logger.log('  2. Otevri skript v Google Ads Scripts editoru.');
  Logger.log('  3. Uprav CONFIG.outputSheetId = "' + ssId + '";');
  Logger.log('  4. Pred prvnim runem: zkontroluj CONFIG.targetPnoPct, lookbackDays,');
  Logger.log('     customLabelIndex a ostatni parametry.');
  Logger.log('  5. Pust main() s CONFIG.dryRun=true pro test, potom false pro production.');
  Logger.log('========================================================');

  return { id: ssId, url: ssUrl, copyUrl: copyUrl, name: sheetName };
}
