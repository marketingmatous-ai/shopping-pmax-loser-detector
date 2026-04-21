/**
 * classifier.gs — Core klasifikacni logika
 *
 * Vsechny funkce jsou pure (bez side effects) — berou product + config + baseline,
 * vraceji strukturovany vysledek. Snadno testovatelne.
 *
 * Pipeline per produkt:
 *   1. checkAgeGate (rising star <30 dni)
 *   2. checkSampleSizeGate (dost kliks?)
 *   3. resolveProductPrice (derived vs fallback AOV)
 *   4. classifyLoser (multi-tier PNO podle volume)
 *   5. classifyLowCtr (account/campaign baseline + IS kontext)
 *   6. mergeResults (primary + secondary + reason)
 *   7. detectTransition (vs lifecycle history)
 */

var Classifier = (function () {

  /**
   * Hlavni pipeline pro jeden produkt. Kombinuje vsechny gaty + klasifikace.
   * @param product ze data.gs (agregovany per item_id)
   * @param config CONFIG
   * @param accountBaseline z data.gs
   * @param perCampaignBaseline optional, pouziva se pri ctrBaselineScope='campaign'
   * @param firstClickDate string YYYY-MM-DD nebo null
   * @param lastYearData optional, pro YoY signal
   * @param previousLifecycle entry z LIFECYCLE_LOG
   * @param runDate Date
   * @param productPrices optional mapa { item_id_lowercase: price }
   * @returns strukturovany vysledek
   */
  function classifyProduct(product, config, accountBaseline, perCampaignBaseline, firstClickDate, lastYearData, previousLifecycle, runDate, productPrices) {
    var result = {
      itemId: product.itemId,
      productTitle: product.productTitle,
      productBrand: product.productBrand,
      productType: product.productType,
      campaignName: product.primaryCampaignName,
      campaignId: product.primaryCampaignId,
      channel: product.channel,
      campaignsCount: product.campaignsCount || 1,
      primaryCampaignSharePct: product.primaryCampaignSharePct || 100,
      topCampaigns: product.topCampaigns || product.primaryCampaignName || '',
      // raw metrics
      clicks: product.clicks,
      impressions: product.impressions,
      cost: product.cost,
      conversions: product.conversions,
      conversionValue: product.conversionValue,
      ctr: product.ctr,
      cvr: product.cvr,
      roas: product.roas,
      actualPno: product.actualPno,
      searchImpressionShare: product.searchImpressionShare,
      // gate trace
      passedAgeGate: false,
      passedSampleGate: false,
      passedYoyCheck: true,
      ageDays: null,
      firstClickDate: firstClickDate || null,
      // threshold trace
      productPrice: 0,
      priceSource: 'none',
      expectedCpa: 0,
      expectedConversions: 0,
      minClicksRequired: 0,
      // classification
      primaryLabel: '',
      secondaryFlags: [],
      reasonCode: '',
      tier: '',
      note: '',
      suggestedAction: '',
      growthPct: null,
      // additional
      wastedSpend: 0,
      wastedSpendPct: 0,
      yoySignal: 'no_yoy_data',
      status: 'ok', // ok / NEW_PRODUCT_RAMP_UP / INSUFFICIENT_DATA / DATA_QUALITY_ISSUE / SKIPPED_DUP
      // transition info (doplni se v detectTransition)
      transitionType: '',
      previousLabel: '',
      previousCampaign: '',
      campaignMoved: false,
      runsSinceFirstFlag: 0,
      daysInCurrentLabel: 0
    };

    // === PRODUCT PRICE RESOLUTION (PRED vsemi gates — pro diagnostiku) ===
    // Zapisujeme priceSource pro VSECHNY produkty (i age-failed, insufficient atd.),
    // aby v DETAIL tabu bylo videt, jestli produkt ma cenu ve feedu, i kdyz jeste
    // neni klasifikovany. To zlepsi diagnostiku novych produktu (RAMP_UP).
    var priceResult = resolveProductPrice(product, accountBaseline, productPrices);
    result.productPrice = priceResult.price;
    result.priceSource = priceResult.source;

    // === YoY SEASONALITY SIGNAL (take pro vsechny, i age-failed) ===
    result.yoySignal = computeYoYSignal(product, lastYearData);

    // === AGE GATE (rising star) ===
    var ageResult = checkAgeGate(product, firstClickDate, runDate, config);
    result.ageDays = ageResult.ageDays;
    result.passedAgeGate = ageResult.passed;
    if (!ageResult.passed) {
      result.status = 'NEW_PRODUCT_RAMP_UP';
      applyTransition(result, previousLifecycle, runDate);
      return result;
    }

    // === SAMPLE SIZE GATE ===
    var sampleResult = checkSampleSizeGate(product, accountBaseline, priceResult.price, config);
    result.expectedCpa = sampleResult.expectedCpa;
    result.expectedConversions = sampleResult.expectedConversions;
    result.minClicksRequired = sampleResult.minClicksRequired;
    result.passedSampleGate = sampleResult.passed;
    if (!sampleResult.passed) {
      result.status = 'INSUFFICIENT_DATA';
      applyTransition(result, previousLifecycle, runDate);
      return result;
    }

    // === DATA QUALITY CHECK ===
    if (product.conversions > 0 && product.conversionValue <= 0) {
      result.status = 'DATA_QUALITY_ISSUE';
      result.note = 'conversions>0 ale conversionValue<=0 — tracking issue, preskakuji klasifikaci';
      applyTransition(result, previousLifecycle, runDate);
      return result;
    }

    // === KLASIFIKACE 1: LOSER_REST ===
    var loser = classifyLoser(product, config, sampleResult.expectedCpa, priceResult.source);
    // Pokud zero-conv byl skipnut kvuli chybejici cene, zaznam v note pro diagnostiku
    if (loser.skippedReason === 'price_unavailable') {
      result.note = 'zero-conv tier SKIP: price_source=unavailable (ani gmc_feed ani derived)';
    }

    // === KLASIFIKACE 2: LOW_CTR_AUDIT ===
    var ctrBaseline = config.ctrBaselineScope === 'campaign'
      ? getCampaignCtr(perCampaignBaseline, product.primaryCampaignId, accountBaseline.avgCtr)
      : accountBaseline.avgCtr;
    var lowCtr = classifyLowCtr(product, ctrBaseline, config);

    // === KLASIFIKACE 3: RISING ===
    var rising = classifyRising(product, config);

    // === KLASIFIKACE 4: DECLINING ===
    var declining = classifyDeclining(product, config);

    // === KLASIFIKACE 5: LOST_OPPORTUNITY ===
    var lostOpp = classifyLostOpportunity(product, config);

    // === MERGE ===
    var merged = mergeResults(loser, lowCtr, rising, declining, lostOpp, config);
    result.primaryLabel = merged.primaryLabel;
    result.secondaryFlags = merged.secondaryFlags;
    result.reasonCode = merged.reasonCode;
    result.tier = merged.tier;
    result.note = merged.note || result.note;
    result.suggestedAction = merged.suggestedAction || '';
    result.growthPct = merged.growthPct;

    // === WASTED SPEND RANKING ===
    result.wastedSpend = computeWastedSpend(product, config);
    result.wastedSpendPct = Utils.safeDiv(result.wastedSpend, product.cost, 0) * 100;

    // === TRANSITION DETECTION ===
    applyTransition(result, previousLifecycle, runDate);

    return result;
  }

  // === GATES ===

  function checkAgeGate(product, firstClickDate, runDate, config) {
    if (!firstClickDate) {
      // Pokud nemame first_click_date, predpokladame, ze produkt je "stary dost"
      // (byl v accountu mimo sledovane okno). Nepoznasluje nic, pripadne loguje.
      return { passed: true, ageDays: null, reason: 'no_first_click_data' };
    }
    // Normalize date — muze byt ISO string, Date objekt, nebo locale string (US/CS).
    // Parse YYYY-MM-DD jako local midnight (ne UTC, jinak TZ offset zkresluje daysBetween)
    var isoStr = Utils.normalizeDate(firstClickDate);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(isoStr)) {
      return { passed: true, ageDays: null, reason: 'invalid_first_click_date_format' };
    }
    var parts = isoStr.split('-');
    var fcd = new Date(
      parseInt(parts[0], 10),
      parseInt(parts[1], 10) - 1,
      parseInt(parts[2], 10)
    );
    var ageDays = Utils.daysBetween(fcd, runDate);
    if (ageDays === null || ageDays < config.minProductAgeDays) {
      return { passed: false, ageDays: ageDays, reason: 'too_young' };
    }
    return { passed: true, ageDays: ageDays, reason: 'ok' };
  }

  function checkSampleSizeGate(product, accountBaseline, productPrice, config) {
    var expectedCpa = productPrice * (config.targetPnoPct / 100);
    var expectedConversions = product.clicks * accountBaseline.cvr;

    // Edge case: pokud account baseline nema data, min_clicks = fallback na absolutni
    if (accountBaseline.cvr <= 0 || accountBaseline.avgCpc <= 0) {
      return {
        passed: product.clicks >= config.minClicksAbsolute,
        expectedCpa: expectedCpa,
        expectedConversions: 0,
        minClicksRequired: config.minClicksAbsolute,
        reason: 'account_no_baseline_data'
      };
    }

    // Sample gate pouziva jen:
    // - minClicksAbsolute (hardcoded floor, typicky 30)
    // - minExpectedConvFloor / CVR (statisticky: "kolik kliku potrebujeme pro ocekavanou konverzi")
    //
    // Price-scaled cost check (2× expected_CPA) se dela az v classifyLoser() pro zero-conv tier —
    // tam je to relevantni (jde o spend threshold, ne click count).
    var minClicks = Math.max(
      config.minClicksAbsolute,
      Math.ceil(Utils.safeDiv(config.minExpectedConvFloor, accountBaseline.cvr, config.minClicksAbsolute))
    );

    return {
      passed: product.clicks >= minClicks,
      expectedCpa: expectedCpa,
      expectedConversions: expectedConversions,
      minClicksRequired: minClicks,
      reason: product.clicks >= minClicks ? 'ok' : 'insufficient_clicks'
    };
  }

  // === PRODUCT PRICE RESOLUTION ===

  /**
   * Resolve product price. Priority:
   *   1. shopping_product.price_micros (presna cena z feedu) — lookup z productPrices mapy
   *   2. derived z conversion_value / conversions (pro produkty s prodeji)
   *   3. UNAVAILABLE (NIKDY avg AOV) — zero-conv produkty bez gmc_feed ceny se nemaji
   *      klasifikovat, protoze bez skutecne ceny je threshold zcela nepresny.
   *
   * Pozn.: avg_AOV fallback zpusobuje false-negatives: levne produkty pod avg dostavaly
   * nerealisticky vysoky zero-conv threshold → nezachytily se jako losery. Proto ho
   * odstranujeme — produkt bez price se neflagguje v zero-conv tieru, ale v low/mid/high
   * volume je OK (tam je aktualPno = cost/value × 100 nezavisle na price).
   */
  function resolveProductPrice(product, accountBaseline, productPrices) {
    // 1. Prime cena z shopping_product resource (GMC feed)
    if (productPrices && product.itemId) {
      var key = String(product.itemId).toLowerCase();
      if (productPrices[key] && productPrices[key] > 0) {
        return {
          price: productPrices[key],
          source: 'gmc_feed'
        };
      }
    }
    // 2. Derived AOV z conversion value (pro produkty s prodeji)
    if (product.conversions >= 1 && product.conversionValue > 0) {
      return {
        price: product.conversionValue / product.conversions,
        source: 'derived'
      };
    }
    // 3. Neni cena dostupna — zero-conv produkty SKIP klasifikaci (ne avg_AOV fallback)
    return {
      price: 0,
      source: 'unavailable'
    };
  }

  // === YoY SIGNAL ===

  function computeYoYSignal(product, lastYearData) {
    if (!lastYearData || !lastYearData[product.itemId]) {
      return 'no_yoy_data';
    }
    var lastYear = lastYearData[product.itemId];
    if (lastYear.conversions >= 5 && product.conversions < 2) {
      return 'possibly_seasonal_decline';
    }
    if (lastYear.conversions < 2 && product.conversions < 2) {
      return 'no_seasonal_pattern';
    }
    return 'stable_yoy';
  }

  // === CLASSIFY LOSER (multi-tier PNO) ===

  function classifyLoser(product, config, expectedCpa, priceSource) {
    var actualPno = product.actualPno; // v %
    var targetPno = config.targetPnoPct;
    var conv = product.conversions;

    // ZERO-CONV tier: conversions < 1 (strict zero + fractional z attribution modelu)
    // Reason: produkt s 0.14 / 0.5 "konverze" z attribution splitu nema zadnou skutecnou
    // prodejni udalost — kazda fractional conv je touchpoint, ne prodej. Proto se hodnoti
    // stejnym cost gate jako strict zero.
    //
    // KRITICKE: zero-conv klasifikace vyzaduje SKUTECNOU cenu produktu (gmc_feed).
    // Bez nej neni threshold spolehlivy — produkt se NEKLASIFIKUJE (status = price_unavailable).
    // Derived price neni mozna (conv == 0), avg_AOV fallback jsme odstranili.
    if (conv < 1) {
      if (priceSource === 'unavailable' || !priceSource || expectedCpa <= 0) {
        return {
          matched: false,
          skippedReason: 'price_unavailable'
        };
      }
      if (product.cost >= config.pnoMultiplierZeroConv * expectedCpa) {
        return {
          matched: true,
          label: config.labelLoserRestValue,
          reason: conv === 0 ? 'zero_sales_high_spend' : 'fractional_conv_high_spend',
          tier: 'zero_conv',
          note: conv > 0 ? 'Fractional conv ' + conv.toFixed(2) + ' (attribution split) — bez realneho prodeje' : ''
        };
      }
      return { matched: false };
    }

    // LOW VOLUME tier: 1 <= conv <= tierLowVolumeMax (default 3)
    if (conv >= 1 && conv <= config.tierLowVolumeMax) {
      if (actualPno >= config.pnoMultiplierLowVol * targetPno) {
        return {
          matched: true,
          label: config.labelLoserRestValue,
          reason: 'low_volume_high_pno',
          tier: 'low_volume',
          note: ''
        };
      }
      return { matched: false };
    }

    // MID VOLUME tier: tierLowVolumeMax < conv <= tierMidVolumeMax (default 4-10)
    if (conv > config.tierLowVolumeMax && conv <= config.tierMidVolumeMax) {
      if (actualPno >= config.pnoMultiplierMidVol * targetPno) {
        return {
          matched: true,
          label: config.labelLoserRestValue,
          reason: 'mid_volume_high_pno',
          tier: 'mid_volume',
          note: ''
        };
      }
      return { matched: false };
    }

    // HIGH VOLUME tier: conv > tierMidVolumeMax (11+)
    if (conv > config.tierMidVolumeMax) {
      if (actualPno >= config.pnoMultiplierHighVol * targetPno) {
        return {
          matched: true,
          label: config.labelLoserRestValue,
          reason: 'high_volume_extreme_pno',
          tier: 'high_volume',
          note: 'POZOR: produkt PRISPIVA k volume — manualni overeni'
        };
      }
      return { matched: false };
    }

    return { matched: false };
  }

  // === CLASSIFY LOW CTR (s IS kontextem) ===

  function classifyLowCtr(product, ctrBaseline, config) {
    // Gate 1: dost impressions (floor proti fluktuacim)
    if (product.impressions < config.minImpressionsLowCtr) {
      return { matched: false };
    }
    // Gate 2 (optional): min clicks floor.
    // Default 0 — pokud ma produkt dost impressions (1000+) s 0-3 kliky,
    // je to jasny signal ze produkt neni atraktivni (CTR << baseline).
    // User muze zvysit pokud dostava false positives u produktu s extremne malo kliky.
    var minClicksLowCtr = config.minClicksLowCtr !== undefined ? config.minClicksLowCtr : 0;
    if (minClicksLowCtr > 0 && product.clicks < minClicksLowCtr) {
      return { matched: false };
    }
    if (ctrBaseline <= 0) {
      return { matched: false };
    }

    if (product.ctr >= config.ctrThresholdMultiplier * ctrBaseline) {
      return { matched: false };
    }

    // === RENTABILITY GATE ===
    // Pokud produkt PRODAVA A JE RENTABILNI, low CTR neni problem —
    // Google ho dobre optimalizuje i pri nizkem CTR (Smart Bidding, PMAX AI).
    // Flagging takoveho produktu by vedl k zbytecnemu vylouceni rentabilniho SKU.
    //
    // Kriteria "rentabilni":
    //   1. Ma dost konverzi (conv >= lowCtrSkipIfProfitableMinConv, default 3)
    //   2. PNO je pod target (mirne pretahnout OK, 10% buffer)
    var minConvForProfitable = config.lowCtrSkipIfProfitableMinConv !== undefined
      ? config.lowCtrSkipIfProfitableMinConv
      : 3;
    var pnoBufferMultiplier = 1.1; // 10% buffer nad target — stale rentabilni
    if (product.conversions >= minConvForProfitable &&
        product.actualPno > 0 &&
        product.actualPno <= config.targetPnoPct * pnoBufferMultiplier) {
      // Produkt prodava rentabilne — neflagovat, ale mohl by dale zlepsit performance
      // (tohle zacne zaznamenat jako "potential improvement" v DETAIL notes, ale bez flagu)
      return {
        matched: false,
        skipped_profitable: true,
        note: 'Rentabilni pres low CTR (conv=' + product.conversions.toFixed(1) +
              ', PNO=' + product.actualPno.toFixed(0) + '% <= target). ' +
              'Audit CTR (fotka/title) by mohl dale zvysit volume.'
      };
    }

    // CTR je pod threshold — zjisti reason code podle IS (pokud je k dispozici).
    // Pozn.: search_impression_share neni v shopping_performance_view,
    // takze v praxi se pouziva jen low_ctr_general. IS-based reasony zustavaji
    // pro pripad, kdy bychom v budoucnu fetchnuli IS z jineho view (campaign-level).
    var reason = 'low_ctr_general';
    if (product.searchImpressionShare > 0 && product.searchImpressionShare < 0.3) {
      reason = 'irrelevant_keyword_match';
    } else if (product.searchImpressionShare > 0.7) {
      reason = 'high_visibility_low_appeal';
    }

    // Suggested action text per reason code
    var suggestedAction = '';
    if (reason === 'irrelevant_keyword_match') {
      suggestedAction = 'Audit product_title a description — produkt se mozna zobrazuje na irelevantnich dotazech. Zvazit upravu title/description + negative keywords.';
    } else if (reason === 'high_visibility_low_appeal') {
      suggestedAction = 'Produkt ma dostatek zobrazeni ale nikdo neklika. Audit: kvalita hlavni fotky, title (keywords, delka, atraktivita), cena vs konkurence, stav skladu.';
    } else {
      suggestedAction = 'General audit: fotka (kvalita/uhel), title (keywords/length), cena (konkurence), skladovost, product_type/category match.';
    }

    return {
      matched: true,
      label: config.labelLowCtrValue,
      reason: reason,
      tier: '',
      note: '',
      suggestedAction: suggestedAction
    };
  }

  function getCampaignCtr(perCampaignBaseline, campaignId, fallback) {
    if (perCampaignBaseline && perCampaignBaseline[campaignId]) {
      return perCampaignBaseline[campaignId].ctr;
    }
    return fallback;
  }

  // === CLASSIFY RISING (trend detection — revenue growth) ===

  /**
   * Detect RISING products — revenue growth vs previous period.
   * Requires min conversions in both periods for statistical validity.
   */
  function classifyRising(product, config) {
    // POZN: Pouzivame main_metrics (ne total) pro konzistenci s LOSER_REST.
    // Brand scale nebo sale-only produkt by jinak vytvoril false RISING:
    //   Priklad: produkt prodal 10k v brand, 50 v main (prev) vs 100 v main (curr).
    //   total growth by byl ~2% (neflag), ale main growth 100% = spravne RISING.
    // Trend se ma posuzovat jen podle main kampani (kde probiha optimalizace).
    var current = product.main_metrics;
    var previous = product.main_metrics_previous;

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
      suggestedAction: 'Early scaling kandidat: zvysit budget / vydelit do vlastni asset group / zvysit tROAS ceiling.'
    };
  }

  // === CLASSIFY DECLINING (trend detection — revenue drop) ===

  /**
   * Detect DECLINING products — revenue drop vs previous period.
   */
  function classifyDeclining(product, config) {
    // POZN: main_metrics (stejne jako RISING — viz komentar vyse pro zduvodneni).
    var current = product.main_metrics;
    var previous = product.main_metrics_previous;

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
      suggestedAction: 'Investigate: zkontroluj cenu vs konkurence, stav skladu, sezonnost, nedavne zmeny feed dat.'
    };
  }

  // === CLASSIFY LOST OPPORTUNITY (profitable + low IS) ===

  /**
   * Detect LOST_OPPORTUNITY — profitable products with low impression share.
   */
  function classifyLostOpportunity(product, config) {
    var total = product.total_metrics;
    if (!total) return { matched: false };
    if (total.conversions < config.lostOpportunityMinConv) return { matched: false };

    if (total.pno === null || total.pno === undefined) return { matched: false };
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
      suggestedAction: 'Rentabilni produkt s nizkou visibility — zvysit bid, samostatna asset group, nebo zahrnout do top-priority kampane.'
    };
  }

  // === MERGE RESULTS ===

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
      if (declining.matched) {
        secondary.push('DECLINING');
        growthPct = declining.growthPct;
      }
      suggestedAction = lowCtr.suggestedAction || '';
    } else if (lowCtr.matched) {
      primary = lowCtr.label;
      reason = lowCtr.reason;
      note = '';
      if (declining.matched) {
        secondary.push('DECLINING');
        growthPct = declining.growthPct;
      }
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

  // === WASTED SPEND ===

  function computeWastedSpend(product, config) {
    var targetRoas = config.targetPnoPct > 0 ? (100 / config.targetPnoPct) : 0;
    if (targetRoas <= 0) {
      return 0;
    }
    var actualRoas = product.roas;
    var achievement = Utils.safeDiv(actualRoas, targetRoas, 0);
    if (achievement >= 1) {
      return 0;
    }
    return product.cost * (1 - achievement);
  }

  // === TRANSITION DETECTION ===

  function applyTransition(result, previousLifecycle, runDate) {
    var prev = previousLifecycle || null;
    var currentLabel = result.primaryLabel;
    var currentCampaign = result.campaignName;

    if (!prev) {
      // Nemame predchozi zaznam
      if (currentLabel) {
        result.transitionType = 'NEW_FLAG';
        result.runsSinceFirstFlag = 1;
        result.daysInCurrentLabel = 0;
      } else {
        result.transitionType = 'NO_CHANGE';
        result.runsSinceFirstFlag = 0;
      }
      return;
    }

    var prevLabel = prev.label || '';
    var prevCampaign = prev.campaign || '';
    result.previousLabel = prevLabel;
    result.previousCampaign = prevCampaign;
    result.campaignMoved = (prevCampaign !== currentCampaign) && (prevCampaign !== '');

    if (currentLabel === '' && prevLabel === '') {
      result.transitionType = 'NO_CHANGE';
      result.runsSinceFirstFlag = prev.runsSinceFirstFlag || 0;
      return;
    }

    if (currentLabel === prevLabel && !result.campaignMoved) {
      result.transitionType = 'REPEATED';
      result.runsSinceFirstFlag = (prev.runsSinceFirstFlag || 0) + 1;
      result.note = (result.note ? result.note + ' | ' : '') + 'REPEATED — label zrejme nebyl aplikovan';
      return;
    }

    if (currentLabel === '' && prevLabel !== '') {
      // Produkt vypadl z flagu. Rozlisujeme 2 pripady:
      //   a) Produkt se fyzicky presunul do REST kampane → RESOLVED (klient aplikoval label)
      //   b) Produkt se zlepsil sam (PNO klesl, CTR zvysil, atd.) → UN_FLAGGED (organicky)
      // Detekce: porovnat currentCampaign proti REST pattern + campaignMoved flag.
      var isInRestCampaign = false;
      try {
        // Presun do rest detekujeme podle:
        //   1. campaignMoved = true (prev != current)
        //   2. current campaign name matchuje rest pattern
        if (result.campaignMoved && currentCampaign) {
          // Pokusi se cist globalni CONFIG, jinak fallback na (?i)REST
          var restPattern = (typeof CONFIG !== 'undefined' && CONFIG.restCampaignPattern) ? CONFIG.restCampaignPattern : '(?i)REST';
          isInRestCampaign = Utils.safeRegexMatch(restPattern, currentCampaign);
        }
      } catch (e) {
        isInRestCampaign = false;
      }

      if (isInRestCampaign) {
        result.transitionType = 'RESOLVED';
        result.note = (result.note ? result.note + ' | ' : '') + 'RESOLVED — přesunut do rest kampaně (úspěšný zásah)';
      } else {
        result.transitionType = 'UN_FLAGGED';
        result.note = (result.note ? result.note + ' | ' : '') + 'UN_FLAGGED — produkt se zlepšil sám, zůstal v main kampani';
      }
      result.runsSinceFirstFlag = prev.runsSinceFirstFlag || 0;
      return;
    }

    if (prevLabel === '' && currentLabel !== '') {
      // Drive mel nejaky transition zaznam (= prev existuje) a ted je flaggovany.
      // Protoze NO_CHANGE se neuklada do LIFECYCLE_LOG, prev existing znamena,
      // ze produkt byl predtim flagged a pak RESOLVED → ted se vraci.
      result.transitionType = 'RE_FLAGGED';
      result.runsSinceFirstFlag = (prev.runsSinceFirstFlag || 0) + 1;
      result.note = (result.note ? result.note + ' | ' : '') + 'RE_FLAGGED — vratil se po resolve';
      return;
    }

    if (prevLabel !== currentLabel && prevLabel !== '' && currentLabel !== '') {
      result.transitionType = 'CATEGORY_CHANGE';
      result.runsSinceFirstFlag = (prev.runsSinceFirstFlag || 0) + 1;
      return;
    }

    // Default
    result.transitionType = currentLabel ? 'NEW_FLAG' : 'NO_CHANGE';
    result.runsSinceFirstFlag = currentLabel ? 1 : 0;
  }

  // === EFFECTIVENESS (agregat pro SUMMARY tab) ===

  /**
   * Pocita effectiveness KPIs z lifecycle_log + current run.
   * @param allResults classifikovane produkty z tohoto runu
   * @param previousLifecycleMap mapa item_id → last entry
   * @param lifecycleFullHistory optional — raw historie vsech radku z LIFECYCLE_LOG (pro pre/post analysis)
   * @returns objekt s KPIs
   */
  function computeEffectiveness(allResults, previousLifecycleMap, lifecycleFullHistory) {
    var transitions = {
      NEW_FLAG: 0,
      REPEATED: 0,
      RESOLVED: 0,        // presun do rest kampane (uspesny zasah klienta)
      UN_FLAGGED: 0,      // produkt se zlepsil sam (organicky), bez zasahu
      RE_FLAGGED: 0,
      CATEGORY_CHANGE: 0,
      NO_CHANGE: 0
    };

    var repeatedCountWarning = 0;

    for (var i = 0; i < allResults.length; i++) {
      var r = allResults[i];
      if (transitions[r.transitionType] !== undefined) {
        transitions[r.transitionType]++;
      }
      if (r.transitionType === 'REPEATED' && r.runsSinceFirstFlag >= 2) {
        repeatedCountWarning++;
      }
    }

    // Application rate: z previous lifecycle entries se label, kolik je ted resolved
    var labeledLastRun = 0;
    var resolvedThisRun = 0;
    if (previousLifecycleMap) {
      for (var itemId in previousLifecycleMap) {
        if (previousLifecycleMap.hasOwnProperty(itemId)) {
          var prev = previousLifecycleMap[itemId];
          if (prev.label && prev.label.length > 0) {
            labeledLastRun++;
            // Najdi odpovidajici current result
            var current = findResultByItemId(allResults, itemId);
            if (!current || current.primaryLabel === '') {
              resolvedThisRun++;
            }
          }
        }
      }
    }

    var applicationRate = labeledLastRun > 0 ? (resolvedThisRun / labeledLastRun) * 100 : null;

    return {
      transitions: transitions,
      repeatedWarning: repeatedCountWarning,
      applicationRate: applicationRate,
      labeledLastRun: labeledLastRun,
      resolvedThisRun: resolvedThisRun
      // NOTE: weekly KPI trend a pre/post cost delta by vyzadovalo pristup k account-level stats
      // napric casem — implementace v budoucnu, pokud bude potreba. Zatim je to minimalni MVP.
    };
  }

  function findResultByItemId(results, itemId) {
    for (var i = 0; i < results.length; i++) {
      if (results[i].itemId === itemId) {
        return results[i];
      }
    }
    return null;
  }

  return {
    classifyProduct: classifyProduct,
    computeEffectiveness: computeEffectiveness,
    classifyRising: classifyRising,
    classifyDeclining: classifyDeclining,
    classifyLostOpportunity: classifyLostOpportunity
  };
})();
