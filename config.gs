/**
 * config.gs — Validace CONFIG objektu
 *
 * Kontroluje vsechna pole CONFIG v main.gs a throws s jasnym popisem,
 * co je spatne. Filozofie: "fail fast and loud" — radsi padnout hned
 * s jasnou chybou, nez tise generovat garbage output.
 *
 * Volana z main() pred zacatkem pipeliny.
 */

var Config = (function () {
  /**
   * Nacte CONFIG hodnoty z tabu "CONFIG" v output Google Sheetu
   * a prepise hodnoty v runtime CONFIG objektu.
   *
   * Tab CONFIG ma strukturu:
   *   A: parametr_nazev | B: hodnota | C: popis
   *
   * Pokud tab neexistuje nebo je prazdny, vrati originalni config bez zmen.
   * Pokud hodnota v sheetu je prazdna, nechava originalni default.
   *
   * @param config  vychozi CONFIG objekt
   * @returns upraveny CONFIG s hodnotami ze sheetu
   */
  function loadFromSheet(config) {
    if (!config.outputSheetId || String(config.outputSheetId).length < 20) {
      return config; // Bez outputSheetId nelze cist ze sheetu
    }
    try {
      var ss = SpreadsheetApp.openById(config.outputSheetId);
      var sheet = ss.getSheetByName('CONFIG');
      if (!sheet) {
        Logger.log('INFO: Tab CONFIG ve sheetu neexistuje — pouzivam defaulty z main.gs.');
        return config;
      }
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return config;
      }
      var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      var overridden = [];
      for (var i = 0; i < data.length; i++) {
        var key = String(data[i][0] || '').replace(/^\s+|\s+$/g, '');
        var rawValue = data[i][1];
        if (!key || key.indexOf('#') === 0 || key.indexOf('▸') === 0) {
          continue; // Skip prazdne a nadpisove radky
        }
        if (rawValue === '' || rawValue === null || rawValue === undefined) {
          continue; // Skip prazdne hodnoty — zachovej default
        }
        // Type coercion podle typu default hodnoty
        if (config.hasOwnProperty(key)) {
          var defaultVal = config[key];
          var coerced = coerceValue(rawValue, defaultVal);
          config[key] = coerced;
          overridden.push(key + '=' + coerced);
        }
      }
      if (overridden.length > 0) {
        Logger.log('INFO: CONFIG tab v sheetu prepsal ' + overridden.length + ' hodnot: ' + overridden.join(', '));
      }
    } catch (e) {
      Logger.log('WARN: Nacteni CONFIG tabu selhalo: ' + e.message + ' — pouzivam defaulty z main.gs.');
    }
    return config;
  }

  /**
   * Koerce value podle typu default hodnoty.
   */
  function coerceValue(rawValue, defaultVal) {
    if (typeof defaultVal === 'boolean') {
      return Utils.safeParseBool(rawValue, defaultVal);
    }
    if (typeof defaultVal === 'number') {
      return Utils.safeParseNumber(rawValue, defaultVal);
    }
    if (Object.prototype.toString.call(defaultVal) === '[object Array]') {
      // Array — ocekavame comma-separated string
      var s = String(rawValue);
      return s.split(',').map(function (x) { return x.replace(/^\s+|\s+$/g, ''); }).filter(function (x) { return x.length > 0; });
    }
    return String(rawValue);
  }

  /**
   * Validuje CONFIG objekt. Throws pri jakekoli nesrovnalosti.
   * @param config CONFIG objekt z main.gs
   */
  function validate(config) {
    var errors = [];

    // === ZAKLADNI ===
    if (!isNumberInRange(config.targetPnoPct, 5, 100)) {
      errors.push('targetPnoPct musi byt cislo mezi 5 a 100 (ted: ' + config.targetPnoPct + ')');
    }
    if (!isNumberInRange(config.lookbackDays, 7, 365)) {
      errors.push('lookbackDays musi byt cislo mezi 7 a 365 (ted: ' + config.lookbackDays + ')');
    }
    if (!config.outputSheetId || typeof config.outputSheetId !== 'string' || config.outputSheetId.length < 20) {
      errors.push('outputSheetId je povinny string (Google Sheets ID, obvykle 40+ znaku)');
    }
    if (config.adminEmail && typeof config.adminEmail === 'string' && config.adminEmail.length > 0) {
      var validEmails = Utils.validateEmails(config.adminEmail);
      if (validEmails.length === 0) {
        errors.push('adminEmail obsahuje pouze nevalidni adresy');
      }
    }

    // === LABEL KONFIGURACE ===
    if (!isIntegerInRange(config.customLabelIndex, 0, 4)) {
      errors.push('customLabelIndex musi byt integer 0-4 (ted: ' + config.customLabelIndex + ')');
    }
    if (!config.labelLoserRestValue || typeof config.labelLoserRestValue !== 'string') {
      errors.push('labelLoserRestValue musi byt neprazdny string');
    }
    if (!config.labelLowCtrValue || typeof config.labelLowCtrValue !== 'string') {
      errors.push('labelLowCtrValue musi byt neprazdny string');
    }
    if (config.labelLoserRestValue === config.labelLowCtrValue) {
      errors.push('labelLoserRestValue a labelLowCtrValue nesmi byt stejne');
    }
    // labelHealthyValue je optional — muze byt '' (vypnuto) nebo neprazdny string
    if (config.labelHealthyValue !== undefined && config.labelHealthyValue !== '') {
      if (typeof config.labelHealthyValue !== 'string') {
        errors.push('labelHealthyValue musi byt string nebo "" (vypnuto)');
      } else {
        if (config.labelHealthyValue === config.labelLoserRestValue) {
          errors.push('labelHealthyValue a labelLoserRestValue nesmi byt stejne');
        }
        if (config.labelHealthyValue === config.labelLowCtrValue) {
          errors.push('labelHealthyValue a labelLowCtrValue nesmi byt stejne');
        }
      }
    }

    // === ITEM_ID CASE OVERRIDE ===
    if (config.itemIdCaseOverride !== undefined) {
      var validCaseModes = ['auto', 'upper', 'lower', 'preserve'];
      if (validCaseModes.indexOf(config.itemIdCaseOverride) === -1) {
        errors.push('itemIdCaseOverride musi byt "auto", "upper", "lower" nebo "preserve" (ted: "' + config.itemIdCaseOverride + '")');
      }
    }

    // === CAMPAIGN FILTERING ===
    if (!isValidRegex(config.brandCampaignPattern)) {
      errors.push('brandCampaignPattern neni validni regex: "' + config.brandCampaignPattern + '"');
    }
    if (!isValidRegex(config.restCampaignPattern)) {
      errors.push('restCampaignPattern neni validni regex: "' + config.restCampaignPattern + '"');
    }
    if (!config.analyzeChannels || !isArray(config.analyzeChannels) || config.analyzeChannels.length === 0) {
      errors.push('analyzeChannels musi byt neprazdne pole (napr. ["SHOPPING", "PERFORMANCE_MAX"])');
    } else {
      var validChannels = ['SHOPPING', 'PERFORMANCE_MAX'];
      for (var i = 0; i < config.analyzeChannels.length; i++) {
        if (validChannels.indexOf(config.analyzeChannels[i]) === -1) {
          errors.push('analyzeChannels obsahuje neplatny kanal: "' + config.analyzeChannels[i] + '" (povolene: ' + validChannels.join(', ') + ')');
        }
      }
    }

    // === SAMPLE SIZE GATE ===
    if (!isNumberInRange(config.minClicksAbsolute, 10, 10000)) {
      errors.push('minClicksAbsolute musi byt 10-10000 (ted: ' + config.minClicksAbsolute + ')');
    }
    if (!isNumberInRange(config.minExpectedConvFloor, 0.1, 50)) {
      errors.push('minExpectedConvFloor musi byt 0.1-50 (ted: ' + config.minExpectedConvFloor + ')');
    }
    // cpaClicksMultiplier: deprecated, odstraneno — price-scaled threshold se pouziva
    // jen v classifyLoser (cost check pro zero-conv tier)

    // === RISING STAR ===
    if (!isNumberInRange(config.minProductAgeDays, 0, 180)) {
      errors.push('minProductAgeDays musi byt 0-180 (ted: ' + config.minProductAgeDays + ')');
    }

    // === LOSER TIER ===
    if (!isIntegerInRange(config.tierLowVolumeMax, 1, 100)) {
      errors.push('tierLowVolumeMax musi byt integer 1-100 (ted: ' + config.tierLowVolumeMax + ')');
    }
    if (!isIntegerInRange(config.tierMidVolumeMax, 2, 1000)) {
      errors.push('tierMidVolumeMax musi byt integer 2-1000 (ted: ' + config.tierMidVolumeMax + ')');
    }
    if (config.tierMidVolumeMax <= config.tierLowVolumeMax) {
      errors.push('tierMidVolumeMax musi byt vetsi nez tierLowVolumeMax');
    }
    if (!isNumberInRange(config.pnoMultiplierZeroConv, 1, 10)) {
      errors.push('pnoMultiplierZeroConv musi byt 1-10 (ted: ' + config.pnoMultiplierZeroConv + ')');
    }
    if (!isNumberInRange(config.pnoMultiplierLowVol, 1, 10)) {
      errors.push('pnoMultiplierLowVol musi byt 1-10 (ted: ' + config.pnoMultiplierLowVol + ')');
    }
    if (!isNumberInRange(config.pnoMultiplierMidVol, 1, 10)) {
      errors.push('pnoMultiplierMidVol musi byt 1-10 (ted: ' + config.pnoMultiplierMidVol + ')');
    }
    if (!isNumberInRange(config.pnoMultiplierHighVol, 1, 10)) {
      errors.push('pnoMultiplierHighVol musi byt 1-10 (ted: ' + config.pnoMultiplierHighVol + ')');
    }

    // === LOW CTR ===
    if (config.ctrBaselineScope !== 'account' && config.ctrBaselineScope !== 'campaign') {
      errors.push('ctrBaselineScope musi byt "account" nebo "campaign" (ted: "' + config.ctrBaselineScope + '")');
    }
    if (!isNumberInRange(config.minImpressionsLowCtr, 100, 1000000)) {
      errors.push('minImpressionsLowCtr musi byt 100-1000000 (ted: ' + config.minImpressionsLowCtr + ')');
    }
    // minClicksLowCtr: 0 = bez limitu (vsechny produkty s dostatkem impressions), 1-10000 = minimum kliku
    if (config.minClicksLowCtr !== undefined && !isNumberInRange(config.minClicksLowCtr, 0, 10000)) {
      errors.push('minClicksLowCtr musi byt 0-10000 (ted: ' + config.minClicksLowCtr + ')');
    }
    if (!isNumberInRange(config.ctrThresholdMultiplier, 0.1, 1.0)) {
      errors.push('ctrThresholdMultiplier musi byt 0.1-1.0 (ted: ' + config.ctrThresholdMultiplier + ')');
    }
    if (config.lowCtrSkipIfProfitableMinConv !== undefined &&
        !isNumberInRange(config.lowCtrSkipIfProfitableMinConv, 0, 100)) {
      errors.push('lowCtrSkipIfProfitableMinConv musi byt 0-100 (ted: ' + config.lowCtrSkipIfProfitableMinConv + ')');
    }

    // === TREND DETECTION (RISING / DECLINING) ===
    if (!isNumberInRange(config.risingGrowthThreshold, 10, 500)) {
      errors.push('risingGrowthThreshold musi byt 10-500 (ted: ' + config.risingGrowthThreshold + ')');
    }
    if (!isNumberInRange(config.decliningDropThreshold, 10, 95)) {
      errors.push('decliningDropThreshold musi byt 10-95 (ted: ' + config.decliningDropThreshold + ')');
    }
    if (!isIntegerInRange(config.minConversionsForTrendCompare, 1, 100)) {
      errors.push('minConversionsForTrendCompare musi byt integer 1-100');
    }

    // === LOST_OPPORTUNITY ===
    if (!isIntegerInRange(config.lostOpportunityMinConv, 1, 100)) {
      errors.push('lostOpportunityMinConv musi byt integer 1-100');
    }
    if (!isNumberInRange(config.lostOpportunityMaxPnoMultiplier, 0.1, 2.0)) {
      errors.push('lostOpportunityMaxPnoMultiplier musi byt 0.1-2.0');
    }
    if (!isNumberInRange(config.lostOpportunityMaxImpressionShare, 0.05, 1.0)) {
      errors.push('lostOpportunityMaxImpressionShare musi byt 0.05-1.0');
    }

    // === EFFECTIVENESS ===
    if (!isIntegerInRange(config.effectivenessMinDaysSinceAction, 1, 180)) {
      errors.push('effectivenessMinDaysSinceAction musi byt 1-180');
    }
    if (!isNumberInRange(config.restCampaignEfficientThreshold, 0.0, 1.0)) {
      errors.push('restCampaignEfficientThreshold musi byt 0.0-1.0');
    }

    // === HISTORICAL DEDUP ===
    if (typeof config.enableHistoryDedup !== 'boolean') {
      errors.push('enableHistoryDedup musi byt boolean (true/false)');
    }
    if (config.enableHistoryDedup && !isNumberInRange(config.historyDedupDays, 1, 90)) {
      errors.push('historyDedupDays musi byt 1-90 (ted: ' + config.historyDedupDays + ')');
    }

    // === BOOLS ===
    if (typeof config.dryRun !== 'boolean') {
      errors.push('dryRun musi byt boolean (true/false)');
    }
    if (typeof config.enableYoYSeasonalityCheck !== 'boolean') {
      errors.push('enableYoYSeasonalityCheck musi byt boolean');
    }
    if (typeof config.groupByParentId !== 'boolean') {
      errors.push('groupByParentId musi byt boolean');
    }
    if (typeof config.includeProductTitles !== 'boolean') {
      errors.push('includeProductTitles musi byt boolean');
    }

    // === ADVANCED ===
    if (!isNumberInRange(config.maxRowsDetailTab, 100, 500000)) {
      errors.push('maxRowsDetailTab musi byt 100-500000 (ted: ' + config.maxRowsDetailTab + ')');
    }

    // === Overovani dostupnosti output sheetu ===
    try {
      var sheet = SpreadsheetApp.openById(config.outputSheetId);
      // OK, muze otevrit
      Logger.log('INFO: Output sheet "' + sheet.getName() + '" je dostupny.');
    } catch (e) {
      errors.push('outputSheetId neni otevirateln (Sheet ID "' + config.outputSheetId + '"): ' + e.message + '. Overte, ze: (1) ID je spravne, (2) sheet je sdileny s uzivatelem spoustejicim skript.');
    }

    if (errors.length > 0) {
      throw new Error(
        'CONFIG validation failed s ' + errors.length + ' chybami:\n' +
        '  - ' + errors.join('\n  - ') +
        '\n\nOpravte CONFIG v main.gs a pust te skript znovu.'
      );
    }

    Logger.log('INFO: CONFIG validace uspesna.');
    return true;
  }

  // === HELPERS ===

  function isNumberInRange(value, min, max) {
    return typeof value === 'number' && !isNaN(value) && isFinite(value) && value >= min && value <= max;
  }

  function isIntegerInRange(value, min, max) {
    return isNumberInRange(value, min, max) && Math.floor(value) === value;
  }

  function isValidRegex(pattern) {
    if (!pattern || typeof pattern !== 'string') {
      return false;
    }
    try {
      // Parsuj leading (?flags) syntax stejne jako Utils.safeRegexMatch
      // (JavaScript RegExp nepodporuje inline flags nativne)
      var flags = '';
      var cleanPattern = String(pattern);
      var flagMatch = cleanPattern.match(/^\(\?([imsuy]+)\)/);
      if (flagMatch) {
        flags = flagMatch[1];
        cleanPattern = cleanPattern.substring(flagMatch[0].length);
      }
      new RegExp(cleanPattern, flags);
      return true;
    } catch (e) {
      return false;
    }
  }

  function isArray(value) {
    return Object.prototype.toString.call(value) === '[object Array]';
  }

  return {
    validate: validate,
    loadFromSheet: loadFromSheet
  };
})();
