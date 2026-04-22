/**
 * ============================================================================
 * Shopping/PMAX Loser Detector — COMBINED FILE (pro Google Ads Scripts UI)
 * ============================================================================
 *
 * Tento soubor je automaticky sestaveny z 7 modulu:
 *   _config (TVOJE KONFIGURACE), utils, config (validation), data,
 *   classifier, output, main.
 *
 * Google Ads Scripts ma jediny editor — nelze pridat multi-file structure.
 * Proto celou kombinaci paste do jednoho skriptu v UI.
 *
 * ⬇⬇⬇ NAHORE SOUBORU JE CONFIG — UPRAVUJ TAM ⬇⬇⬇
 *
 * PRVNI RUN workflow:
 *   1. Paste combined.gs do Google Ads Scripts editoru
 *   2. Nech CONFIG.outputSheetId prazdny ('')
 *   3. Pust main() (Nahled/Preview)
 *   4. V logu najdi ID a URL vytvoreneho sheetu
 *   5. Paste ID do CONFIG.outputSheetId a nastav targetPnoPct, atd.
 *   6. Pust znovu s dryRun=true pro test
 *   7. Po verifikaci prepni dryRun=false a pust ostry run
 *
 * DOKUMENTACE:
 *   /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/README.md
 *   /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/DEPLOYMENT-GUIDE.md
 *
 * VYGENEROVANO: 2026-04-22 21:41
 * BUILD SCRIPT: build-combined.sh
 * ============================================================================
 */


/**
 * ============================================================================
 * 🔧 KONFIGURACE — TUTO SEKCI UPRAV PRO KAZDEHO KLIENTA
 * ============================================================================
 *
 * Vsechny parametry jsou v CONFIG objektu nize. Uprav podle potreb klienta,
 * uloz (Cmd+S) a spust skript (Nahled / Spustit).
 *
 * ─────────────────────────────────────────────────────────────
 * 📋 PRVNI SETUP — 2 MOZNOSTI
 * ─────────────────────────────────────────────────────────────
 *
 *  A) AUTO-SETUP (doporuceno pro prvni deploy):
 *     1. Nech CONFIG.outputSheetId prazdny ('')
 *     2. V dropdownu "Vybrat funkci" nahore zvol "setupOutputSheet"
 *     3. Klikni "Spustit" (ne Nahled — zapisuje novy soubor do Drive)
 *     4. V logu najdes ID + URL noveho sheetu
 *     5. Zkopiruj ID do CONFIG.outputSheetId nize
 *     6. Prepni dropdown zpet na "main" a pokracuj
 *
 *  B) KOPIE HOTOVEHO TEMPLATE (doporuceno — rychle + bezpecne):
 *     1. Otevri tento odkaz v prohlizeci pro vytvoreni kopie cisteho template:
 *
 *        👉  https://docs.google.com/spreadsheets/d/1BPmB00tXlc7Jq5sdNrrMUXo7OFQp1eoTYaSBLTRYkTA/copy
 *
 *     2. Google zobrazi dialog "Vytvorit kopii" — potvrd
 *     3. Kopii prejmenuj podle klienta (napr. "Loser Detector — newclient.cz")
 *     4. Zkopiruj ID kopie z URL (cast mezi /d/ a /edit)
 *     5. Vloz ID do CONFIG.outputSheetId nize a pust main()
 *
 *     Template URL vyse je CISTY sheet bez klientskych dat — bezpecny
 *     pro verejne sdileni. Obsahuje jen layout (README, CONFIG tab,
 *     placeholder zpravy).
 *
 * ─────────────────────────────────────────────────────────────
 * 🚀 DAL WORKFLOW (po setupu)
 * ─────────────────────────────────────────────────────────────
 *
 *   1. Uprav targetPnoPct, brandCampaignPattern atd. podle klienta
 *   2. Pust main() s dryRun=true pro test
 *   3. Po verifikaci logu prepni dryRun=false a pust ostry run
 *   4. V Google Ads Scripts UI nastav Schedule (napr. weekly)
 *
 * REFERENCE:
 *   GitHub: https://github.com/marketingmatous-ai/shopping-pmax-loser-detector
 *   Issues: https://github.com/marketingmatous-ai/shopping-pmax-loser-detector/issues
 * ============================================================================
 */

var CONFIG = {
  // === ZÁKLADNÍ ===
  targetPnoPct:           30,       // Cilove PNO v % (napr. 30 = 30% nakladu z revenue)
  lookbackDays:           30,       // Okno analyzy (doporuceno 30-90)
  outputSheetId:          '',       // Google Sheets ID — NECH PRAZDNE pri prvnim runu
                                    // (auto-setup vytvori novy sheet a vypise ID do logu)
  adminEmail:             '',       // Email pro notifikace ('' = bez emailu)

  // === LABEL KONFIGURACE ===
  customLabelIndex:       2,        // Cislo labelu 0-4 (uzivatel zvoli, co nepouziva v GMC)
  labelLoserRestValue:    'loser_rest',     // Hodnota zapsana do custom_label
  labelLowCtrValue:       'low_ctr_audit',  // Hodnota pro low-CTR kategorii
  labelHealthyValue:      'healthy',        // Produkty co prosly revizi (status='ok', bez flagu).
                                            // Zapise se do FEED_UPLOAD spolu s flagged.
                                            // Pouziti: v rest kampani filter `custom_label_N != healthy`
                                            // aby zdrave produkty zustaly jen v main kampanich.
                                            // '' (prazdny string) = nezapisovat (opt-out).

  // === ITEM_ID CASE (pro match s GMC) ===
  // Google Ads vraci item_id lowercase, ale GMC ukladá UPPERCASE (typicky).
  // Skript mapuje z shopping_product (canonical zdroj), pro unmatched produkty
  // (disapproved, stazene z feedu) pouzije heuristiku / override.
  //   'auto'     — detekuj dominantni case (>=70% produktu) z shopping_product (DEFAULT)
  //   'upper'    — vynutit UPPERCASE pro vsechny (doporuceno pokud GMC ma IDs velkymi pismeny)
  //   'lower'    — vynutit lowercase pro vsechny
  //   'preserve' — nic neupravovat (vrati to, co dostaneme z Google Ads)
  itemIdCaseOverride:     'auto',

  // === CAMPAIGN FILTERING ===
  brandCampaignPattern:   '(?i)BRD',        // Brand kampane — VYLOUCENE z analyzy
                                            // (uprav podle naming konvence klienta, napr. BRA, BRAND)
  restCampaignPattern:    '(?i)REST',       // Rest kampane — ignorovane
  analyzeChannels:        ['SHOPPING', 'PERFORMANCE_MAX'],  // Typy kampani

  // === SAMPLE SIZE GATE (ochrana proti false-positive) ===
  minClicksAbsolute:      30,       // Absolutni minimum kliku pred klasifikaci
  minExpectedConvFloor:   1,        // Min expected conversions (clicks × account CVR)

  // === RISING STAR PROTECTION ===
  minProductAgeDays:      30,       // Produkty mladsi N dni se neevaluuji

  // === LOSER TIER THRESHOLDS ===
  tierLowVolumeMax:       3,        // Conv 1-3 = low volume tier
  tierMidVolumeMax:       10,       // Conv 4-10 = mid volume tier
  pnoMultiplierZeroConv:  2.0,      // 0 conv → spend ≥ 2× expected CPA
  pnoMultiplierLowVol:    1.5,      // 1-3 conv → PNO ≥ 1.5× target
  pnoMultiplierMidVol:    2.0,      // 4-10 conv → PNO ≥ 2.0× target
  pnoMultiplierHighVol:   3.0,      // 11+ conv → PNO ≥ 3.0× target (chrani volume)

  // === LOW CTR DETEKCE ===
  ctrBaselineScope:       'account',  // 'account' (doporuceno) nebo 'campaign'
  minImpressionsLowCtr:   500,        // Min impressions pro low-CTR detekci
  minClicksLowCtr:        0,          // Min kliku pro low-CTR (0 = vsechny; 1000 imp stacila jako signal)
  ctrThresholdMultiplier: 0.7,        // Produkt s CTR < 0.7× baseline = flag
  lowCtrSkipIfProfitableMinConv: 3,   // Rentabilni produkty (conv>=N a PNO<=target*1.1) se neflaggují

  // === TREND DETECTION (RISING/DECLINING) ===
  risingGrowthThreshold:         50,    // Growth >= X% = RISING
  decliningDropThreshold:        30,    // Drop >= X% = DECLINING
  minConversionsForTrendCompare:  3,    // Min conv obou periodach

  // === LOST_OPPORTUNITY ===
  lostOpportunityMinConv:             5,
  lostOpportunityMaxPnoMultiplier:    0.8,  // PNO <= target × N
  lostOpportunityMaxImpressionShare:  0.5,  // IS < N

  // === EFFECTIVENESS ===
  effectivenessMinDaysSinceAction:    14,
  restCampaignEfficientThreshold:     0.2,  // Rest cost <= N × before_main

  // === SEASONALITY ===
  enableYoYSeasonalityCheck: true,  // YoY porovnani z historical data

  // === HISTORICAL DEDUP ===
  enableHistoryDedup:     true,     // Tracking v LIFECYCLE_LOG tabu
  historyDedupDays:       14,

  // === DRY RUN ===
  dryRun:                 false,    // true = jen loguje, nezapisuje do sheetu

  // === PRODUKT AGREGACE ===
  groupByParentId:        false,    // true = agreguj varianty na parent_id

  // === ADVANCED ===
  includeProductTitles:   true,     // Privacy flag — titles v output sheetu
  maxRowsDetailTab:       50000     // Sheet size guard
};

// ============================================================================
// ⬇⬇⬇ KOD SKRIPTU NIZE — NEMUSIS UPRAVOVAT ⬇⬇⬇
// ============================================================================

/**
 * utils.gs — Pomocne funkce pro Shopping/PMAX Loser Detector
 *
 * Obsahuje:
 *  - Safe math (div by zero, NaN handling)
 *  - Regex s try/catch
 *  - Formatovani cisel a procent s cappingem extremnich hodnot
 *  - Validace emailu
 *  - Prevody jednotek (micros → koruny)
 *  - Datove utility (formatovani datumu, rozdil dnu)
 *
 * Vsechny funkce jsou pure (bez side effects) a obsahuji null/edge handling.
 */

var Utils = (function () {
  /**
   * Bezpecny podil. Pokud delitel je 0 nebo delenec je null/undefined, vraci fallback.
   */
  function safeDiv(numerator, denominator, fallback) {
    if (fallback === undefined) {
      fallback = 0;
    }
    if (denominator === 0 || denominator === null || denominator === undefined) {
      return fallback;
    }
    if (numerator === null || numerator === undefined) {
      return fallback;
    }
    var result = numerator / denominator;
    if (!isFinite(result)) {
      return fallback;
    }
    return result;
  }

  /**
   * Formatovani procent s cappingem. NaN / Infinity → "∞", >9999 → "9999%+".
   */
  function safePctFormat(value, decimals) {
    if (decimals === undefined) {
      decimals = 2;
    }
    if (value === null || value === undefined || isNaN(value)) {
      return '-';
    }
    if (!isFinite(value)) {
      return '∞';
    }
    if (value > 9999) {
      return '9999%+';
    }
    if (value < -9999) {
      return '-9999%';
    }
    return value.toFixed(decimals) + '%';
  }

  /**
   * Formatovani penez. Prevod z micros na hlavni jednotku (1 Kc = 1_000_000 micros).
   */
  function microsToMajor(micros) {
    if (micros === null || micros === undefined || isNaN(micros)) {
      return 0;
    }
    return micros / 1000000;
  }

  /**
   * Formatovani cisla s tisici oddelovacem (pro vypsani do sheetu a emailu).
   */
  function formatNumber(n, decimals) {
    if (decimals === undefined) {
      decimals = 2;
    }
    if (n === null || n === undefined || isNaN(n)) {
      return '-';
    }
    return n.toFixed(decimals);
  }

  /**
   * Bezpecny regex match s podporou inline flagu (?i), (?m), (?s), (?u).
   *
   * JavaScript RegExp nativne nepodporuje inline flags jako `(?i)pattern` —
   * flags musi byt predane jako druhy argument `new RegExp(pattern, 'i')`.
   * Tento wrapper parsuje leading `(?flags)` a prelozi je do JS flags syntaxi.
   *
   * Priklad: "(?i)BRD" → new RegExp("BRD", "i") — case-insensitive match.
   *
   * Pokud pattern je invalidni, vraci false + loguje warning.
   */
  function safeRegexMatch(pattern, text) {
    if (!pattern || !text) {
      return false;
    }
    try {
      var flags = '';
      var cleanPattern = String(pattern);
      // Extract leading (?xyz) flags — ECMAScript regex nemaji inline flag support,
      // musime je prelozit do flags argumentu.
      var flagMatch = cleanPattern.match(/^\(\?([imsuy]+)\)/);
      if (flagMatch) {
        flags = flagMatch[1];
        cleanPattern = cleanPattern.substring(flagMatch[0].length);
      }
      var re = new RegExp(cleanPattern, flags);
      return re.test(text);
    } catch (e) {
      Logger.log('WARN: Invalid regex pattern "' + pattern + '": ' + e.message);
      return false;
    }
  }

  /**
   * Validuje comma-separated email string, vraci list validnich emailu.
   */
  function validateEmails(emailStr) {
    if (!emailStr) {
      return [];
    }
    var parts = emailStr.split(',');
    var valid = [];
    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    for (var i = 0; i < parts.length; i++) {
      var trimmed = parts[i].replace(/^\s+|\s+$/g, '');
      if (emailRegex.test(trimmed)) {
        valid.push(trimmed);
      } else if (trimmed.length > 0) {
        Logger.log('WARN: Invalid email address skipped: "' + trimmed + '"');
      }
    }
    return valid;
  }

  /**
   * Formatovani datumu jako YYYY-MM-DD (pro GAQL a sheety).
   */
  function formatDate(date) {
    if (!date) {
      return '';
    }
    var y = date.getFullYear();
    var m = String(date.getMonth() + 1);
    var d = String(date.getDate());
    if (m.length === 1) {
      m = '0' + m;
    }
    if (d.length === 1) {
      d = '0' + d;
    }
    return y + '-' + m + '-' + d;
  }

  /**
   * Normalizuje libovolny datumovy input (Date objekt, ISO string, US format)
   * na YYYY-MM-DD format. Pouziva se pro dedup klice v LIFECYCLE_LOG —
   * Google Sheets si muze datum prepnout mezi formatami (ISO ↔ locale),
   * a string comparison by jinak selhavala.
   *
   * Priklady:
   *   Date(2026, 3, 21)  → "2026-04-21"
   *   "2026-04-21"       → "2026-04-21"
   *   "4/21/2026"        → "2026-04-21"  (US format)
   *   "21.4.2026"        → "2026-04-21"  (CS format)
   *   null / undefined   → ""
   */
  function normalizeDate(raw) {
    if (raw === null || raw === undefined || raw === '') {
      return '';
    }
    if (raw instanceof Date) {
      return formatDate(raw);
    }
    var s = String(raw).replace(/^\s+|\s+$/g, '');
    // Uz ISO format (YYYY-MM-DD) — vrati prvnich 10 znaku
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      return s.substring(0, 10);
    }
    // US format: M/D/YYYY nebo MM/DD/YYYY
    var usMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (usMatch) {
      var yr = usMatch[3];
      var mn = usMatch[1].length === 1 ? '0' + usMatch[1] : usMatch[1];
      var dy = usMatch[2].length === 1 ? '0' + usMatch[2] : usMatch[2];
      return yr + '-' + mn + '-' + dy;
    }
    // CS format: D.M.YYYY nebo DD.MM.YYYY
    var csMatch = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (csMatch) {
      var yr2 = csMatch[3];
      var mn2 = csMatch[2].length === 1 ? '0' + csMatch[2] : csMatch[2];
      var dy2 = csMatch[1].length === 1 ? '0' + csMatch[1] : csMatch[1];
      return yr2 + '-' + mn2 + '-' + dy2;
    }
    // Unknown format — return as-is (falls back to string comparison)
    return s;
  }

  /**
   * Pocet dnu mezi dvema datumy (date2 - date1, v kalendarnich dnech).
   */
  function daysBetween(date1, date2) {
    if (!date1 || !date2) {
      return null;
    }
    var msPerDay = 1000 * 60 * 60 * 24;
    var diff = date2.getTime() - date1.getTime();
    return Math.floor(diff / msPerDay);
  }

  /**
   * Vrati Date objekt posunuty o N dnu oproti referenci.
   */
  function addDays(date, days) {
    var result = new Date(date.getTime());
    result.setDate(result.getDate() + days);
    return result;
  }

  /**
   * Bezpecna konverze stringu na cislo. Pro prazdne / neparsovatelne vraci fallback.
   */
  function safeParseNumber(value, fallback) {
    if (fallback === undefined) {
      fallback = 0;
    }
    if (value === null || value === undefined || value === '') {
      return fallback;
    }
    var n = Number(value);
    if (isNaN(n) || !isFinite(n)) {
      return fallback;
    }
    return n;
  }

  /**
   * Bezpecna konverze stringu na bool. Akceptuje "TRUE"/"FALSE"/"true"/"1"/"0"/bool.
   */
  function safeParseBool(value, fallback) {
    if (fallback === undefined) {
      fallback = false;
    }
    if (value === null || value === undefined || value === '') {
      return fallback;
    }
    if (typeof value === 'boolean') {
      return value;
    }
    var s = String(value).toLowerCase().replace(/^\s+|\s+$/g, '');
    if (s === 'true' || s === '1' || s === 'yes') {
      return true;
    }
    if (s === 'false' || s === '0' || s === 'no') {
      return false;
    }
    return fallback;
  }

  /**
   * Zkratka: urcuje, zda kampan je brandova dle configu.
   */
  function isBrandCampaign(campaignName, brandPattern) {
    return safeRegexMatch(brandPattern, campaignName);
  }

  /**
   * Zkratka: urcuje, zda kampan je restova dle configu.
   */
  function isRestCampaign(campaignName, restPattern) {
    return safeRegexMatch(restPattern, campaignName);
  }

  /**
   * Agreguje pole objektu podle klice (reduce).
   * @param items pole objektu
   * @param keyFn funkce vracejici klic pro agregaci
   * @param reducer funkce (acc, item) → acc
   * @param initialValue initial value per key
   */
  function groupBy(items, keyFn, reducer, initialValue) {
    var groups = {};
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var key = keyFn(item);
      if (groups[key] === undefined) {
        groups[key] = typeof initialValue === 'function' ? initialValue() : JSON.parse(JSON.stringify(initialValue));
      }
      groups[key] = reducer(groups[key], item);
    }
    return groups;
  }

  /**
   * Vrati poli razene dle numerickeho pole (desc).
   */
  function sortDesc(items, fieldFn) {
    return items.slice().sort(function (a, b) {
      return fieldFn(b) - fieldFn(a);
    });
  }

  return {
    safeDiv: safeDiv,
    safePctFormat: safePctFormat,
    microsToMajor: microsToMajor,
    formatNumber: formatNumber,
    safeRegexMatch: safeRegexMatch,
    validateEmails: validateEmails,
    formatDate: formatDate,
    normalizeDate: normalizeDate,
    daysBetween: daysBetween,
    addDays: addDays,
    safeParseNumber: safeParseNumber,
    safeParseBool: safeParseBool,
    isBrandCampaign: isBrandCampaign,
    isRestCampaign: isRestCampaign,
    groupBy: groupBy,
    sortDesc: sortDesc
  };
})();

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

    // === Product prices + canonical item_ids (mapovano z shopping_product resource) ===
    var productPricesResult = fetchProductPrices(aggregated);
    var productPrices = productPricesResult.prices;
    var itemIdCanonical = productPricesResult.canonical;
    Logger.log('INFO: Nactene ceny pro ' + Object.keys(productPrices).length + ' produktu z ' + aggregated.length + ' agregovanych.');
    Logger.log('INFO: Canonical item_ids (pro GMC upload) nacteno: ' + Object.keys(itemIdCanonical).length);

    // Detekuj dominantni case z canonical mapy (pro fallback u produktu co neni v shopping_product).
    // Pokud >=70% produktu ma pismenka UPPERCASE, aplikujeme upper na unmatched (a vice versa).
    // Duvod: disapproved / stazene produkty nejsou v shopping_product, ale jsou v GMC pod UPPERCASE
    // (protoze klient systematicky pouziva UPPERCASE). Bez heuristiky by GMC upload selhal.
    //
    // CONFIG.itemIdCaseOverride umoznuje vynutit case manualne ('upper'/'lower'/'preserve').
    // Default 'auto' = heuristika.
    var caseStats = detectDominantCase(itemIdCanonical);
    var override = config.itemIdCaseOverride || 'auto';
    var dominantCase; // 'upper' | 'lower' | null
    if (override === 'upper' || override === 'lower') {
      dominantCase = override;
      Logger.log('INFO: item_id case — CONFIG OVERRIDE = ' + override +
                 ' (stats: upper=' + caseStats.upperPct + '% / lower=' + caseStats.lowerPct + '% / mixed=' + caseStats.mixedPct + '% z ' + caseStats.total + ')');
    } else if (override === 'preserve') {
      dominantCase = null;
      Logger.log('INFO: item_id case — preserve (zadny fallback). Stats: upper=' + caseStats.upperPct + '% / lower=' + caseStats.lowerPct + '% / mixed=' + caseStats.mixedPct + '% z ' + caseStats.total);
    } else {
      dominantCase = caseStats.decision;
      Logger.log('INFO: item_id case — auto-detect: ' + (dominantCase || 'mixed/unknown') +
                 ' (stats: upper=' + caseStats.upperPct + '% / lower=' + caseStats.lowerPct + '% / mixed=' + caseStats.mixedPct + '% z ' + caseStats.total + ')');
      if (!dominantCase && caseStats.total >= 10) {
        Logger.log('HINT: Pokud GMC feed upload selhava pro unmatched produkty, nastav CONFIG.itemIdCaseOverride = "upper" (nebo "lower") pro vynuceni.');
      }
    }

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

    // === Enrich aggregated produkty o previous period reference + impression share + canonical ID ===
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

      // Canonical item_id pro GMC upload.
      // Priority:
      //  1) shopping_product mapping (zdroj pravdy — presne jak je v GMC, vcetne mixed-case "Print" atd.)
      //  2) dominantCase heuristika (pro disapproved — jen pokud override aktivni nebo auto detekce souhlasi)
      //  3) fallback: itemKey jak prisel z performance view (lowercase)
      //
      // hasCanonicalFromGmc = true znamena, ze produkt je aktualne v GMC (mapping overen).
      // Pokud false, produkt byl pravdepodobne stazen z feedu / disapproved → v GMC neni, upload selze.
      if (itemIdCanonical[itemKeyLower]) {
        aggregated[pi].itemIdCanonical = itemIdCanonical[itemKeyLower];
        aggregated[pi].hasCanonicalFromGmc = true;
      } else if (dominantCase === 'upper') {
        aggregated[pi].itemIdCanonical = String(itemKey).toUpperCase();
        aggregated[pi].hasCanonicalFromGmc = false;
      } else if (dominantCase === 'lower') {
        aggregated[pi].itemIdCanonical = String(itemKey).toLowerCase();
        aggregated[pi].hasCanonicalFromGmc = false;
      } else {
        aggregated[pi].itemIdCanonical = itemKey;
        aggregated[pi].hasCanonicalFromGmc = false;
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
   * Detekuje dominantni case v canonical item_id mape (UPPER / LOWER / null).
   * Pouziva se jako fallback pro produkty, ktere nejsou v shopping_product (disapproved,
   * stazene z feedu, atd.) — pokud je vetsina produktu UPPERCASE, predpoklada se, ze
   * unmatched produkty budou v GMC take UPPERCASE.
   *
   * Vraci { decision: 'upper' | 'lower' | null, upperPct, lowerPct, mixedPct, total }.
   * Rozhodnuti: >=70% produktu ma dany case → dominantni. Jinak null.
   */
  function detectDominantCase(canonical) {
    var upperCount = 0;
    var lowerCount = 0;
    var mixedCount = 0;
    var total = 0;
    for (var k in canonical) {
      if (!canonical.hasOwnProperty(k)) continue;
      var v = canonical[k];
      // Analyza jen pismen (cislice a mezery case nemaji)
      var letters = String(v).replace(/[^A-Za-z]/g, '');
      if (!letters) continue;
      total++;
      if (letters === letters.toUpperCase()) {
        upperCount++;
      } else if (letters === letters.toLowerCase()) {
        lowerCount++;
      } else {
        mixedCount++;
      }
    }
    var result = {
      decision: null,
      upperPct: total > 0 ? Math.round(upperCount / total * 100) : 0,
      lowerPct: total > 0 ? Math.round(lowerCount / total * 100) : 0,
      mixedPct: total > 0 ? Math.round(mixedCount / total * 100) : 0,
      total: total
    };
    if (total < 10) return result; // malo dat pro spolehlivou detekci
    if (upperCount / total >= 0.7) result.decision = 'upper';
    else if (lowerCount / total >= 0.7) result.decision = 'lower';
    return result;
  }

  /**
   * Mapuje item_id na product_price (z shopping_product resource)
   * a zaroven vraci mapu lowercase → canonical (UPPERCASE z GMC) item_id.
   *
   * shopping_performance_view NEMA cenu produktu — ta je v shopping_product.
   * item_id v shopping_product je UPPERCASE (napr. "NB 2414 KO" = jak je v GMC),
   * v shopping_performance_view je lowercase ("nb 2414 ko").
   * Interne mapujeme case-insensitive (lowercase klic), ale pro upload do GMC
   * potrebujeme zachovat canonical UPPERCASE variantu.
   *
   * @param products pole aggregovanych produktu
   * @returns { prices: {lowercase_item_id: price}, canonical: {lowercase_item_id: original_item_id} }
   */
  function fetchProductPrices(products) {
    var result = { prices: {}, canonical: {} };
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
      var canonicalCount = 0;
      var skippedNoPrice = 0;
      var samplePairs = []; // prvnich 5 item_id:price pro diagnostiku
      while (iterator.hasNext()) {
        var row = iterator.next();
        var itemId = row.shoppingProduct && row.shoppingProduct.itemId ? row.shoppingProduct.itemId : null;
        if (!itemId) continue;
        var lowerKey = String(itemId).toLowerCase();
        // Canonical mapa — vzdy (i pro produkty bez ceny, aby slo doplnit case pro GMC upload)
        if (!result.canonical[lowerKey]) {
          result.canonical[lowerKey] = String(itemId);
          canonicalCount++;
        }
        var priceMicros = row.shoppingProduct.priceMicros ? Number(row.shoppingProduct.priceMicros) : 0;
        if (priceMicros <= 0) {
          skippedNoPrice++;
          continue;
        }
        result.prices[lowerKey] = priceMicros / 1000000;
        fetchedCount++;
        if (samplePairs.length < 5) {
          samplePairs.push(lowerKey + '=' + (priceMicros / 1000000).toFixed(0));
        }
      }
      Logger.log('INFO: fetchProductPrices — nacteno ' + fetchedCount + ' cen + ' + canonicalCount + ' canonical IDs (skipped no_price: ' + skippedNoPrice + '). Sample: ' + samplePairs.join(', '));

      // Diagnostika: kolik z agreg. produktu ma cenu v feedu?
      var matched = 0;
      var unmatched = 0;
      var unmatchedSamples = [];
      for (var p = 0; p < products.length; p++) {
        var pid = String(products[p].itemId || '').toLowerCase();
        if (result.prices[pid]) {
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
      itemIdCanonical: product.itemIdCanonical || product.itemId,  // verze z shopping_product (pro GMC upload)
      hasCanonicalFromGmc: !!product.hasCanonicalFromGmc,            // true = produkt aktualne v GMC feedu
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

/**
 * output.gs — Zapis vysledku do Google Sheetu + email notifikace
 *
 * 4 taby:
 *   1. FEED_UPLOAD    — ready-to-upload CSV format (id, custom_label_N)
 *   2. DETAIL         — rozhodovaci data pro specialisty (trace vsech gate, reason codes)
 *   3. SUMMARY        — dashboard + effectiveness KPIs
 *   4. LIFECYCLE_LOG  — per-produkt timeline (append-only, jen transitions)
 *
 * Idempotence: FEED_UPLOAD, DETAIL, SUMMARY pouzivaji clearContents() + fresh write.
 * Jen LIFECYCLE_LOG je append-only.
 */

var Output = (function () {
  var MAX_CELLS_PER_TAB_LIMIT = 9500000; // 10M celkove limit, ponecham buffer
  var BATCH_SIZE = 5000; // Radku na jeden setValues() call

  /**
   * Prida (nebo refresh) filtr na tab pres vsechny data radky + sloupce.
   * Pokud filtr uz existuje, smaze ho a vytvori novy (po clearContents by jinak
   * mohl odkazovat na neexistujici range).
   *
   * @param sheet Google Sheet object
   * @param lastRow posledni radek s datem (vcetne headeru)
   * @param lastCol posledni sloupec s datem
   */
  function addFilterToSheet(sheet, lastRow, lastCol) {
    if (!sheet || lastRow < 2 || lastCol < 1) {
      return;
    }
    try {
      var existing = sheet.getFilter();
      if (existing) {
        existing.remove();
      }
      sheet.getRange(1, 1, lastRow, lastCol).createFilter();
    } catch (e) {
      Logger.log('WARN: Nepodarilo se pridat filtr na tab "' + sheet.getName() + '": ' + e.message);
    }
  }

  /**
   * Hlavni entry — zapise vsechny taby + posle email.
   */
  function writeAll(outputSheetId, classified, summary, effectiveness, config, runDate, existingActions, existingTimeline) {
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
    appendLifecycleLogTab(ss, classified, runDate);
    appendWeeklySnapshot(ss, summary, effectiveness, classified, runDate);
    // MONTHLY_REVIEW ctene z LIFECYCLE_LOG, musi byt PO appendu lifecyclu
    writeMonthlyReviewTab(ss, runDate);
    // SUMMARY zapisovan naposledy, aby Phase 5 sekce mohly cist z WEEKLY_SNAPSHOT (vcetne aktualniho tydne)
    writeSummaryTab(ss, summary, classified, effectiveness, config, runDate);
    // CONFIG tab refresh — aktualne pouzite hodnoty (nikoli hardcoded z initial setupu)
    refreshConfigTab(ss, config);

    Logger.log('INFO: Output written.');
  }

  /**
   * CONFIG tab — refresh z runtime config objektu pri kazdem runu.
   * Pouziva globalni helper buildConfigTabRows() z main.gs.
   * Tim zajistime, ze hodnoty v sheet CONFIG tabu odpovidaji tem,
   * ktere skript skutecne pouzil pri tomto runu.
   */
  function refreshConfigTab(ss, config) {
    var sheet = ss.getSheetByName('CONFIG');
    if (!sheet) {
      Logger.log('WARN: CONFIG tab neexistuje — preskakuji refresh.');
      return;
    }
    if (typeof buildConfigTabRows !== 'function') {
      Logger.log('WARN: buildConfigTabRows() neni dostupna — preskakuji refresh.');
      return;
    }
    var rows = buildConfigTabRows(config);
    sheet.clearContents();
    // Clear existing conditional formats
    try { sheet.clearConditionalFormatRules(); } catch (cfe) { /* ok */ }

    // === HLAVIČKA (banner nad daty) ===
    // Vytvorime vlastni banner row 1, data pak zacnou od row 2
    var totalRows = rows.length + 3; // +3 pro banner + info + blank
    var headerBanner = sheet.getRange(1, 1, 1, 3);
    headerBanner.merge();
    headerBanner.setValue('⚙️  KONFIGURACE SKRIPTU (runtime hodnoty z posledního běhu)')
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat');
    sheet.setRowHeight(1, 36);

    // Info řádek — jak se CONFIG používá
    var infoBanner = sheet.getRange(2, 1, 1, 3);
    infoBanner.merge();
    infoBanner.setValue('ℹ️  Tento tab je POUZE INFORMATIVNÍ. Zdroj pravdy = CONFIG objekt v kódu skriptu (Google Ads → Skripty → editor). Úpravy zde nemají vliv.')
      .setBackground('#fff4e5').setFontColor('#a56200').setFontSize(10).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheet.setRowHeight(2, 32);

    // Prazdny radek
    sheet.setRowHeight(3, 8);

    // === HEADER tabulky (row 4) ===
    sheet.getRange(4, 1, 1, 3).setValues([['Parametr', 'Hodnota', 'Popis']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#174ea6')
      .setFontSize(11).setVerticalAlignment('middle').setFontFamily('Montserrat')
      .setBorder(true, true, true, true, null, null, '#0b3d91', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(4, 30);

    // === DATA (row 5 a niz) ===
    // rows[0] je hlavicka (Parametr, Hodnota, Popis) — skipneme ji
    var dataRows = rows.slice(1);
    if (dataRows.length > 0) {
      sheet.getRange(5, 1, dataRows.length, 3).setValues(dataRows);

      // Styling per radek
      for (var cr = 0; cr < dataRows.length; cr++) {
        var label = dataRows[cr][0];
        var val = dataRows[cr][1];
        var desc = dataRows[cr][2];
        var sheetRow = cr + 5; // 5 = offset kvuli banneru + headeru

        if (label && (val === '' || val === null || val === undefined) && (desc === '' || desc === null || desc === undefined)) {
          // Sekcni header (group label jako "ZÁKLADNÍ", "CAMPAIGN FILTERING" atd.)
          sheet.getRange(sheetRow, 1, 1, 3)
            .merge()
            .setBackground('#e8f0fe')
            .setFontWeight('bold')
            .setFontColor('#1a73e8')
            .setFontSize(11)
            .setFontFamily('Montserrat')
            .setHorizontalAlignment('left')
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, null, null, '#a5c3f0', SpreadsheetApp.BorderStyle.SOLID);
          sheet.setRowHeight(sheetRow, 28);
        } else {
          // Data radek
          // Alternating rows pro citelnost
          var bgColor = (cr % 2 === 0) ? '#ffffff' : '#fafafa';
          sheet.getRange(sheetRow, 1, 1, 3).setBackground(bgColor)
            .setBorder(null, null, true, null, null, null, '#e8e8e8', SpreadsheetApp.BorderStyle.SOLID);

          // Sloupec 1 (param name): monospace, bold
          sheet.getRange(sheetRow, 1).setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(10).setFontColor('#202124');
          // Sloupec 2 (value): bold, center, colored dle typu
          var valCell = sheet.getRange(sheetRow, 2);
          valCell.setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setFontFamily('Roboto Mono');
          if (typeof val === 'boolean') {
            valCell.setFontColor(val ? '#1e8e3e' : '#ea4335').setBackground(val ? '#e6f4ea' : '#fce8e6');
          } else if (typeof val === 'number') {
            valCell.setFontColor('#1a73e8');
          } else if (val === '') {
            valCell.setValue('—').setFontColor('#9aa0a6');
          } else {
            valCell.setFontColor('#202124');
          }
          // Sloupec 3 (description): italic, gray
          sheet.getRange(sheetRow, 3).setFontSize(10).setFontColor('#5f6368').setFontStyle('italic').setWrap(true);
          sheet.setRowHeight(sheetRow, 26);
        }
      }
    }

    // Frozen header (row 4 = parametr/hodnota/popis)
    sheet.setFrozenRows(4);

    // Column widths — dostatek prostoru pro dlouhe popisy
    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 520);

    // Conditional format: zvyraznit specialni hodnoty v col B
    try {
      var valuesRange = sheet.getRange(5, 2, dataRows.length, 1);
      var rules = [
        // dryRun = true → oranzovy warning
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=AND($A5="dryRun",$B5=TRUE)')
          .setBackground('#fff4e5').setFontColor('#a56200').setBold(true)
          .setRanges([valuesRange]).build()
      ];
      sheet.setConditionalFormatRules(rules);
    } catch (cfrE) { /* ok */ }
  }

  /**
   * FEED_UPLOAD tab — jen id + custom_label_{N}, ready pro upload do GMC / Mergado.
   */
  function writeFeedUploadTab(ss, classified, config) {
    var sheetName = 'FEED_UPLOAD';
    var sheet = getOrCreateSheet(ss, sheetName);
    sheet.clearContents();

    var headerName = 'custom_label_' + config.customLabelIndex;
    var data = [['id', headerName]];
    var flaggedCount = 0;
    var healthyCount = 0;

    // DULEZITE: Pouzivame itemIdCanonical (z shopping_product = presne jak je v GMC vcetne mixed case).
    // shopping_performance_view vraci lowercase item_id, ale GMC ukladá typicky UPPERCASE (nebo mixed case).
    // Bez teto konverze GMC supplemental feed nematchne produkty a labely se neaplikuji.
    //
    // SKIP: Produkty bez canonical z GMC (hasCanonicalFromGmc=false) se NEZAPISUJI — pravdepodobne
    // byly stazeny z feedu (disapproved/deleted) a GMC je nezna. Zapsani by vedlo k chybe "Nabidka neexistuje".

    var skippedNotInGmc = 0;

    // Prvni prochod: flagged produkty (zachovava prioritni razeni)
    for (var i = 0; i < classified.length; i++) {
      var c = classified[i];
      if (c.primaryLabel && c.primaryLabel.length > 0) {
        if (!c.hasCanonicalFromGmc) {
          skippedNotInGmc++;
          continue;
        }
        data.push([c.itemIdCanonical || c.itemId, c.primaryLabel]);
        flaggedCount++;
      }
    }

    // Druhy prochod: healthy produkty (pokud je labelHealthyValue nastaveny)
    // Kriterium: status='ok' (prosel vsemi gates) + bez primaryLabel (klasifikator ho nevyhodil).
    // Zombie (INSUFFICIENT_DATA / NEW_PRODUCT_RAMP_UP / DATA_QUALITY_ISSUE) nedostavaji label.
    if (config.labelHealthyValue && config.labelHealthyValue.length > 0) {
      for (var j = 0; j < classified.length; j++) {
        var h = classified[j];
        if (h.status === 'ok' && (!h.primaryLabel || h.primaryLabel.length === 0)) {
          if (!h.hasCanonicalFromGmc) {
            skippedNotInGmc++;
            continue;
          }
          data.push([h.itemIdCanonical || h.itemId, config.labelHealthyValue]);
          healthyCount++;
        }
      }
    }

    if (skippedNotInGmc > 0) {
      Logger.log('INFO: FEED_UPLOAD — skipnuto ' + skippedNotInGmc + ' produktu (nejsou v shopping_product, pravdepodobne stazene z GMC feedu).');
    }

    if (data.length === 1) {
      Logger.log('INFO: FEED_UPLOAD — zadny produkt nebyl flaggovany ani healthy.');
      sheet.getRange(1, 1, 1, 2).setValues(data);
      return;
    }

    writeDataInBatches(sheet, data);
    Logger.log('INFO: FEED_UPLOAD — zapsano ' + flaggedCount + ' flagged' +
               (healthyCount > 0 ? ' + ' + healthyCount + ' healthy' : '') +
               ' = ' + (data.length - 1) + ' radku.');
  }

  /**
   * Write ACTIONS tab — all flagged products in one filterable view.
   * Preserves manual columns from existingActions map (action_taken, action_date, consultant_note).
   */
  function writeActionsTab(ss, classified, existingActions, config, summary) {
    var sheet = getOrCreateSheet(ss, 'ACTIONS');
    sheet.clearContents();
    try { sheet.clearConditionalFormatRules(); } catch (ce) { /* ok */ }
    // Unmerge vsechny existing merges (clearContents je nerusi, kolize pri novych merge)
    try {
      var maxCols = Math.max(28, sheet.getMaxColumns());
      var maxRows = Math.max(50, sheet.getMaxRows());
      sheet.getRange(1, 1, Math.min(maxRows, 50), Math.min(maxCols, 28)).breakApart();
    } catch (be) { /* ok */ }

    // === COLLECT FLAGGED ===
    var flagged = [];
    for (var fi = 0; fi < classified.length; fi++) {
      if (classified[fi].primaryLabel && classified[fi].primaryLabel.length > 0) {
        flagged.push(classified[fi]);
      }
    }
    flagged.sort(function (a, b) {
      return computePriorityScore(b) - computePriorityScore(a);
    });

    // === SOUHRNNE KPI + INSIGHTS PANEL (nad tabulkou) ===
    var currency = summary && summary.currency ? summary.currency : 'CZK';
    var categoryCounts = { loser_rest: 0, low_ctr_audit: 0, DECLINING: 0, RISING: 0, LOST_OPPORTUNITY: 0 };
    var totalWasted = 0, totalCost = 0;
    var newFlags = 0, repeated = 0, repeatedWarning = 0, reFlagged = 0;
    var missingManualAction = 0;

    for (var ci2 = 0; ci2 < flagged.length; ci2++) {
      var cc = flagged[ci2];
      if (categoryCounts[cc.primaryLabel] !== undefined) categoryCounts[cc.primaryLabel]++;
      totalWasted += (cc.wastedSpend || 0);
      totalCost += (cc.cost || 0);
      if (cc.transitionType === 'NEW_FLAG') newFlags++;
      if (cc.transitionType === 'REPEATED') repeated++;
      if (cc.transitionType === 'REPEATED' && cc.runsSinceFirstFlag >= 2) repeatedWarning++;
      if (cc.transitionType === 'RE_FLAGGED') reFlagged++;
      var manualForCheck = existingActions[cc.itemId] || {};
      if (!manualForCheck.action_taken || String(manualForCheck.action_taken).trim() === '') {
        missingManualAction++;
      }
    }

    // BANNER nadpis
    var aRow = 1;
    sheet.getRange(aRow, 1, 1, 12).merge()
      .setValue('🎯  AKČNÍ SEZNAM (priorita 1 = nejvyšší)')
      .setBackground('#c5221f').setFontColor('#ffffff').setFontSize(14).setFontWeight('bold')
      .setFontFamily('Montserrat').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(aRow, 36); aRow++;

    // INFO řádek
    sheet.getRange(aRow, 1, 1, 12).merge()
      .setValue('ℹ️  Seřazeno podle priority (wasted_spend pro losery, growth/drop pro RISING/DECLINING). Manuální sloupce (action_taken / action_date / consultant_note) se PŘEPISUJÍ NOVÝMI HODNOTAMI — skript je zachovává mezi běhy.')
      .setBackground('#fff4e5').setFontColor('#a56200').setFontSize(10).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheet.setRowHeight(aRow, 32); aRow++;

    sheet.setRowHeight(aRow, 8); aRow++; // spacer

    // 4 KPI karty (side by side, row 4-6)
    function writeActionsKpi(startCol, colSpan, label, value, subtext, bgColor, valueColor) {
      var r = sheet.getRange(aRow, startCol, 3, colSpan);
      r.merge();
      r.setBackground(bgColor)
        .setBorder(true, true, true, true, null, null, valueColor, SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheet.getRange(aRow, startCol).setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(
          label + '\n\n' + value + '\n' + subtext
        ).setTextStyle(0, label.length, SpreadsheetApp.newTextStyle()
          .setFontSize(9).setBold(true).setForegroundColor(valueColor).build())
          .setTextStyle(label.length + 2, label.length + 2 + String(value).length, SpreadsheetApp.newTextStyle()
            .setFontSize(20).setBold(true).setForegroundColor(valueColor).build())
          .setTextStyle(label.length + 2 + String(value).length + 1, (label + '\n\n' + value + '\n' + subtext).length, SpreadsheetApp.newTextStyle()
            .setFontSize(10).setItalic(true).setForegroundColor(valueColor).build())
        .build()
      );
      sheet.getRange(aRow, startCol).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    }

    writeActionsKpi(1, 3,
      'CELKEM OZNAČENO',
      formatInt(flagged.length),
      newFlags + ' nových · ' + repeated + ' opakovaných',
      '#fce4e4', '#c5221f');
    writeActionsKpi(4, 3,
      'MARNÝ SPEND (k řešení)',
      formatMoney(totalWasted, currency),
      totalCost > 0 ? roundNumber(totalWasted / totalCost * 100, 1) + '% z nákladů na označené' : '—',
      '#fff4e5', '#a56200');
    writeActionsKpi(7, 3,
      'BEZ MANUÁLNÍ AKCE',
      formatInt(missingManualAction),
      'chybí zápis v action_taken',
      '#e8f0fe', '#1a73e8');
    writeActionsKpi(10, 3,
      'WARNING (≥ 2 běhy)',
      formatInt(repeatedWarning),
      'label nebyl aplikován',
      repeatedWarning > 0 ? '#fce4e4' : '#eaf7ea',
      repeatedWarning > 0 ? '#c5221f' : '#1e8e3e');

    for (var kpiR = 0; kpiR < 3; kpiR++) sheet.setRowHeight(aRow + kpiR, 28);
    aRow += 3;
    sheet.setRowHeight(aRow, 8); aRow++; // spacer

    // BREAKDOWN PER KATEGORIE
    sheet.getRange(aRow, 1, 1, 12).merge()
      .setValue('ROZDĚLENÍ PODLE KATEGORIE')
      .setBackground('#5f6368').setFontColor('#ffffff').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(aRow, 26); aRow++;

    // 6 kategorii × 2 sloupce = 12 sloupcu (match s layoutem ACTIONS)
    // Kazda kategorie ma vlastni par merged bunek pro header + value
    // Tim se vyrovna alignment — cisla jsou VZDY pod svou kategorii (bez ohledu na delku textu)
    var cats = [
      ['⚠️ loser_rest', categoryCounts.loser_rest, '#fce8e6', '#c5221f'],
      ['🔵 low_ctr_audit', categoryCounts.low_ctr_audit, '#fef7e0', '#b06000'],
      ['📉 DECLINING', categoryCounts.DECLINING, '#fce8e6', '#b31412'],
      ['📈 RISING', categoryCounts.RISING, '#e6f4ea', '#1e8e3e'],
      ['💎 LOST_OPPORTUNITY', categoryCounts.LOST_OPPORTUNITY, '#e8f0fe', '#1a73e8'],
      ['CELKEM', flagged.length, '#f1f3f4', '#202124']
    ];

    // Header row (kategorie labely)
    for (var ci = 0; ci < cats.length; ci++) {
      var startCol = ci * 2 + 1;
      sheet.getRange(aRow, startCol, 1, 2).merge()
        .setValue(cats[ci][0])
        .setBackground('#fafafa').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, false, true, null, null, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(aRow, 26); aRow++;

    // Value row (cisla)
    for (var ci2 = 0; ci2 < cats.length; ci2++) {
      var startCol2 = ci2 * 2 + 1;
      sheet.getRange(aRow, startCol2, 1, 2).merge()
        .setValue(formatInt(cats[ci2][1]))
        .setBackground(cats[ci2][2])
        .setFontColor(cats[ci2][3])
        .setFontSize(18).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, true, true, true, null, null, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(aRow, 36); aRow++;
    sheet.setRowHeight(aRow, 8); aRow++; // spacer

    // INSIGHTS — kratke actionable poznamky (prvnich pet relevant situations)
    var insights = [];
    if (repeatedWarning > 0) {
      insights.push('⚠️  ' + repeatedWarning + ' produktů je označeno ≥ 2 běhy bez změny. Label pravděpodobně nebyl aplikován v GMC/Mergado.');
    }
    if (reFlagged > 0) {
      insights.push('🔄  ' + reFlagged + ' produktů se vrátilo jako označené po předchozím vyřešení. Zvaž jiný přístup — neúspěšná intervence.');
    }
    if (missingManualAction > 0 && flagged.length > 0) {
      var missingPct = Math.round(missingManualAction / flagged.length * 100);
      insights.push('📝  ' + missingPct + '% produktů nemá vyplněný action_taken. Zapiš co jsi udělal (excluded, label, bid-decrease…).');
    }
    if (categoryCounts.LOST_OPPORTUNITY > 0) {
      insights.push('💎  ' + categoryCounts.LOST_OPPORTUNITY + ' rentabilních produktů má nízký impression share — zvaž bid-boost nebo dedikovanou kampaň.');
    }
    if (categoryCounts.RISING > 0) {
      insights.push('📈  ' + categoryCounts.RISING + ' produktů rychle roste. Monitoruj je — mohou být dobrými kandidáty na škálování.');
    }
    if (insights.length === 0) {
      insights.push('✅  Žádná kritická upozornění. Pokračuj podle priority níže.');
    }

    sheet.getRange(aRow, 1, 1, 12).merge()
      .setValue('💡  DOPORUČENÍ A UPOZORNĚNÍ')
      .setBackground('#e8f0fe').setFontColor('#1a73e8').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(aRow, 26); aRow++;
    for (var ins = 0; ins < insights.length; ins++) {
      sheet.getRange(aRow, 1, 1, 12).merge()
        .setValue(insights[ins])
        .setBackground('#ffffff').setFontSize(10).setFontColor('#202124')
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
      sheet.setRowHeight(aRow, 22); aRow++;
    }
    sheet.setRowHeight(aRow, 16); aRow++; // spacer

    // === HLAVICKA TABULKY ===
    var headerStartRow = aRow;

    var headers = [
      'priority_rank', 'category', 'item_id', 'product_title', 'product_price',
      'current_campaign', 'tier', 'reason_code',
      'main_clicks', 'main_impressions', 'main_cost', 'main_conv', 'main_pno_pct', 'main_ctr_pct',
      'total_clicks', 'total_cost', 'total_conv', 'total_roas', 'brand_share_pct',
      'growth_pct', 'wasted_spend', 'recommended_action',
      'days_since_first_flag', 'transition_status', 'secondary_flags',
      'action_taken', 'action_date', 'consultant_note'
    ];

    var data = [headers];
    for (var i = 0; i < flagged.length; i++) {
      var c = flagged[i];
      var manual = existingActions[c.itemId] || {};

      var totalConv = (c.total_metrics && c.total_metrics.conversions) || 0;
      var brandConv = (c.brand_metrics && c.brand_metrics.conversions) || 0;
      var brandSharePct = totalConv > 0 ? (brandConv / totalConv * 100) : 0;

      var mainClicks = c.main_metrics ? c.main_metrics.clicks : c.clicks;
      var mainImpressions = c.main_metrics ? c.main_metrics.impressions : c.impressions;
      var mainCost = c.main_metrics ? c.main_metrics.cost : c.cost;
      var mainConv = c.main_metrics ? c.main_metrics.conversions : c.conversions;
      var mainPno = c.main_metrics ? c.main_metrics.pno : c.actualPno;
      var mainCtr = c.main_metrics ? c.main_metrics.ctr : c.ctr;

      data.push([
        i + 1,
        c.primaryLabel,
        c.itemId,
        c.productTitle || '',
        roundNumber(c.productPrice, 2),
        c.primaryCampaignName || c.campaignName || '',
        c.tier || '',
        c.reasonCode || '',
        mainClicks,
        mainImpressions,
        roundNumber(mainCost, 2),
        roundNumber(mainConv, 2),
        roundNumber(mainPno, 2),
        roundNumber((mainCtr || 0) * 100, 3),
        c.total_metrics ? c.total_metrics.clicks : 0,
        roundNumber(c.total_metrics ? c.total_metrics.cost : 0, 2),
        roundNumber(c.total_metrics ? c.total_metrics.conversions : 0, 2),
        roundNumber(c.total_metrics ? c.total_metrics.roas : 0, 2),
        roundNumber(brandSharePct, 1),
        c.growthPct !== null && c.growthPct !== undefined ? roundNumber(c.growthPct, 1) : '',
        roundNumber(c.wastedSpend, 2),
        c.suggestedAction || '',
        '',  // days_since_first_flag — populated from PRODUCT_TIMELINE in future iteration
        c.transitionType || 'NEW_FLAG',
        (c.secondaryFlags || []).join(', '),
        manual.action_taken || '',
        manual.action_date || '',
        manual.consultant_note || ''
      ]);
    }

    if (data.length === 1) {
      sheet.getRange(headerStartRow, 1, 1, headers.length).setValues(data);
      Logger.log('INFO: ACTIONS — zadne flagged produkty (souhrn panel vygenerovan).');
      return;
    }

    sheet.getRange(headerStartRow, 1, data.length, headers.length).setValues(data);
    sheet.setFrozenRows(headerStartRow); // zmrazit vcetne panelu + hlavicky
    sheet.getRange(headerStartRow, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff').setFontSize(10);
    sheet.setRowHeight(headerStartRow, 30);

    // Force text format pro action_date sloupec (aby Sheets nepreformatoval ISO na locale)
    var idxActionDate = headers.indexOf('action_date');
    if (idxActionDate >= 0 && data.length > 1) {
      sheet.getRange(headerStartRow + 1, idxActionDate + 1, data.length - 1, 1).setNumberFormat('@');
    }

    // Applikuj barvy per kategorie (na data pod headerem)
    applyCategoryColorsOffset(sheet, flagged, headers.length, headerStartRow);
    // Filtr — od header row dolu
    addFilterToSheet(sheet, headerStartRow + data.length - 1, headers.length);

    // Scroll vzdycky dolu na tabulku (user okamzite vidi co resit — ne panely nahore)
    // ... zvazit zda je to vhodne. Pro MVP necháme výchozí chování.

    Logger.log('INFO: ACTIONS — written ' + flagged.length + ' rows (hlavicka + ' + insights.length + ' insights).');
  }

  /**
   * applyCategoryColors — s offsetem (data zacinaji od prozadene row, ne od 1).
   */
  function applyCategoryColorsOffset(sheet, flagged, numCols, startRow) {
    var colors = {
      'loser_rest': '#fde8e8',
      'low_ctr_audit': '#fff8d4',
      'DECLINING': '#ffe0cc',
      'RISING': '#d4f5d4',
      'LOST_OPPORTUNITY': '#d4e4f5'
    };
    for (var i = 0; i < flagged.length; i++) {
      var c = flagged[i];
      var color = colors[c.primaryLabel];
      if (color) {
        // startRow = header row, data od startRow + 1
        sheet.getRange(startRow + 1 + i, 1, 1, numCols).setBackground(color);
      }
    }
  }

  function computePriorityScore(c) {
    if (c.primaryLabel === 'loser_rest') return c.wastedSpend || 0;
    if (c.primaryLabel === 'DECLINING') return Math.abs(c.growthPct || 0) * 10;
    if (c.primaryLabel === 'LOST_OPPORTUNITY') return (c.total_metrics && c.total_metrics.conversionValue) || 0;
    if (c.primaryLabel === 'RISING') return (c.growthPct || 0) * 10;
    return (c.impressions || 0) / 100;
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

  /**
   * Write PRODUCT_TIMELINE tab — per-product history, upsert logic.
   * Preserves history (first_flag_date, categories_history) + manual columns from existingTimeline + existingActions.
   */
  function writeProductTimelineTab(ss, classified, existingTimeline, existingActions, config, runDate) {
    var sheet = getOrCreateSheet(ss, 'PRODUCT_TIMELINE');

    var headers = [
      'item_id', 'product_title', 'first_flag_date', 'total_runs_flagged',
      'categories_history', 'current_status',
      'kpi_before_cost', 'kpi_before_conv', 'kpi_before_pno', 'kpi_before_roas', 'kpi_before_ctr',
      'kpi_current_total_cost', 'kpi_current_total_conv', 'kpi_current_total_pno', 'kpi_current_total_roas', 'kpi_current_total_ctr',
      'kpi_current_main_cost', 'kpi_current_main_conv', 'kpi_current_main_pno',
      'kpi_current_rest_cost', 'kpi_current_rest_conv',
      'kpi_current_brand_cost', 'kpi_current_brand_conv',
      'delta_cost_pct', 'delta_conv_pct', 'delta_pno_pct', 'delta_roas_pct',
      'effectiveness_score', 'days_since_action',
      'latest_action', 'latest_action_date', 'latest_note'
    ];

    // Collect item_ids from both current flagged and previously in timeline
    var itemIds = {};
    for (var ci = 0; ci < classified.length; ci++) {
      if (classified[ci].primaryLabel) itemIds[classified[ci].itemId] = true;
    }
    for (var k in existingTimeline) {
      if (existingTimeline.hasOwnProperty(k)) itemIds[k] = true;
    }

    var runDateStr = Utils.formatDate(runDate);
    var data = [headers];

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
    // Force text format pro date sloupce (first_flag_date = col 3, latest_action_date = col 31)
    // aby Sheets nepreformatovalo ISO na locale
    if (data.length > 1) {
      sheet.getRange(2, 3, data.length - 1, 1).setNumberFormat('@'); // first_flag_date
      var idxLatestActionDate = headers.indexOf('latest_action_date');
      if (idxLatestActionDate >= 0) {
        sheet.getRange(2, idxLatestActionDate + 1, data.length - 1, 1).setNumberFormat('@');
      }
    }
    addFilterToSheet(sheet, data.length, headers.length);

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
    // Normalize existing.first_flag_date (Date objekt / locale string → ISO)
    var existingFlagDate = existing && existing.first_flag_date ? Utils.normalizeDate(existing.first_flag_date) : '';
    var firstFlagDate = existingFlagDate || (isFlagged ? runDateStr : '');
    var totalRunsFlagged = (existing && existing.total_runs_flagged) || 0;
    if (isFlagged) totalRunsFlagged = Number(totalRunsFlagged) + 1;

    var categoriesHistory = existing && existing.categories_history ? String(existing.categories_history) : '';
    if (isFlagged) {
      var newCat = current.primaryLabel + (current.tier ? ':' + current.tier : '');
      if (categoriesHistory.indexOf(newCat) === -1) {
        categoriesHistory = categoriesHistory ? (categoriesHistory + ' \u2192 ' + newCat) : newCat;
      }
    } else if (existing && existing.current_status === 'FLAGGED') {
      categoriesHistory = categoriesHistory + ' \u2192 RESOLVED';
    }

    var currentStatus = isFlagged ? 'FLAGGED' : (existing && existing.first_flag_date ? 'RESOLVED' : 'STABLE');

    // KPI before: snapshot from existing (captured once on first flag) or from previous period on first-flag moment
    var kpiBeforeCost = (existing && existing.kpi_before_cost !== undefined && existing.kpi_before_cost !== '') ? existing.kpi_before_cost : '';
    var kpiBeforeConv = (existing && existing.kpi_before_conv !== undefined && existing.kpi_before_conv !== '') ? existing.kpi_before_conv : '';
    var kpiBeforePno = (existing && existing.kpi_before_pno !== undefined && existing.kpi_before_pno !== '') ? existing.kpi_before_pno : '';
    var kpiBeforeRoas = (existing && existing.kpi_before_roas !== undefined && existing.kpi_before_roas !== '') ? existing.kpi_before_roas : '';
    var kpiBeforeCtr = (existing && existing.kpi_before_ctr !== undefined && existing.kpi_before_ctr !== '') ? existing.kpi_before_ctr : '';

    // If first time flagged, capture snapshot from total_metrics_previous
    if (isFlagged && (!existing || !existing.first_flag_date) && current.total_metrics_previous) {
      var prev = current.total_metrics_previous;
      if (kpiBeforeCost === '') kpiBeforeCost = roundNumber(prev.cost || 0, 2);
      if (kpiBeforeConv === '') kpiBeforeConv = roundNumber(prev.conversions || 0, 2);
      if (kpiBeforePno === '' && prev.conversionValue > 0) kpiBeforePno = roundNumber((prev.cost / prev.conversionValue * 100) || 0, 2);
      if (kpiBeforeRoas === '' && prev.cost > 0) kpiBeforeRoas = roundNumber((prev.conversionValue / prev.cost) || 0, 2);
    }

    // Current snapshots
    var t = current && current.total_metrics ? current.total_metrics : { cost: 0, conversions: 0, pno: 0, roas: 0, ctr: 0 };
    var m = current && current.main_metrics ? current.main_metrics : { cost: 0, conversions: 0, pno: 0 };
    var r = current && current.rest_metrics ? current.rest_metrics : { cost: 0, conversions: 0 };
    var b = current && current.brand_metrics ? current.brand_metrics : { cost: 0, conversions: 0 };

    // Deltas (only computable if we have numeric kpiBefore)
    var beforeCostNum = Number(kpiBeforeCost);
    var beforeConvNum = Number(kpiBeforeConv);
    var beforePnoNum = Number(kpiBeforePno);
    var beforeRoasNum = Number(kpiBeforeRoas);

    var deltaCost = (beforeCostNum > 0 && !isNaN(beforeCostNum)) ? roundNumber((t.cost - beforeCostNum) / beforeCostNum * 100, 1) : '';
    var deltaConv = (beforeConvNum > 0 && !isNaN(beforeConvNum)) ? roundNumber((t.conversions - beforeConvNum) / beforeConvNum * 100, 1) : '';
    var deltaPno = (beforePnoNum > 0 && !isNaN(beforePnoNum)) ? roundNumber((t.pno - beforePnoNum) / beforePnoNum * 100, 1) : '';
    var deltaRoas = (beforeRoasNum > 0 && !isNaN(beforeRoasNum)) ? roundNumber((t.roas - beforeRoasNum) / beforeRoasNum * 100, 1) : '';

    var effectivenessScore = computeEffectivenessScore(manual, { cost: beforeCostNum, roas: beforeRoasNum }, t, runDateStr, config);
    var daysSinceAction = computeDaysSinceAction(manual, runDateStr);

    return [
      itemId,
      (current && current.productTitle) || (existing && existing.product_title) || '',
      firstFlagDate,
      totalRunsFlagged,
      categoriesHistory,
      currentStatus,
      kpiBeforeCost,
      kpiBeforeConv,
      kpiBeforePno,
      kpiBeforeRoas,
      kpiBeforeCtr,
      roundNumber(t.cost, 2),
      roundNumber(t.conversions, 2),
      roundNumber(t.pno, 2),
      roundNumber(t.roas, 2),
      roundNumber((t.ctr || 0) * 100, 3),
      roundNumber(m.cost, 2),
      roundNumber(m.conversions, 2),
      roundNumber(m.pno, 2),
      roundNumber(r.cost, 2),
      roundNumber(r.conversions, 2),
      roundNumber(b.cost, 2),
      roundNumber(b.conversions, 2),
      deltaCost,
      deltaConv,
      deltaPno,
      deltaRoas,
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
    if (days === '' || days < config.effectivenessMinDaysSinceAction) return 'PENDING';
    if (!kpiBefore.cost || kpiBefore.cost <= 0 || isNaN(kpiBefore.cost)) return 'N/A';

    var deltaCostPct = (kpiCurrent.cost - kpiBefore.cost) / kpiBefore.cost * 100;
    var deltaRoasPct = (kpiBefore.roas && kpiBefore.roas > 0 && !isNaN(kpiBefore.roas))
      ? ((kpiCurrent.roas - kpiBefore.roas) / kpiBefore.roas * 100)
      : 0;

    if (deltaCostPct <= -30 && deltaRoasPct >= -10) return '+';
    if (deltaCostPct > -10 || deltaRoasPct < -30) return '-';
    return '=';
  }

  function computeDaysSinceAction(manual, runDateStr) {
    if (!manual.action_date) return '';
    // Normalize action_date — muze byt Date objekt / ISO / US / CS format
    var isoDate = Utils.normalizeDate(manual.action_date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return '';
    var parts = isoDate.split('-');
    var actionDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    var runDateParts = runDateStr.split('-');
    var runDate = new Date(parseInt(runDateParts[0], 10), parseInt(runDateParts[1], 10) - 1, parseInt(runDateParts[2], 10));
    return Utils.daysBetween(actionDate, runDate);
  }

  /**
   * Upsert 1 row per week_id to WEEKLY_SNAPSHOT with account-level KPI + flag counts.
   * Pri opakovanem spusteni v tydenu se existujici radek prepisuje (nikoli appenduje).
   * Tim zabranime duplikacim pri dev/testing runech. Tydni trendy grafy ctou
   * vzdy jen 1 radek per tyden = konzistentni view.
   */
  function appendWeeklySnapshot(ss, summary, effectiveness, classified, runDate) {
    var sheet = getOrCreateSheet(ss, 'WEEKLY_SNAPSHOT');

    var headers = [
      'run_date', 'week_id', 'account_cost_total', 'account_clicks', 'account_conversions',
      'account_conv_value', 'account_roas', 'account_pno_pct', 'account_ctr_pct',
      'flagged_count_total', 'flagged_loser_rest', 'flagged_low_ctr',
      'flagged_declining', 'flagged_rising', 'flagged_lost_opp',
      'wasted_spend_total', 'resolved_this_run', 're_flagged_this_run',
      'label_application_rate_pct'
    ];

    // Check row 1: pokud neni header (napr. placeholder z setupOutputSheet), prepsat.
    // Placeholder "(naplni se...)" vs prazdny sheet vs stara verze bez headeru.
    var r1Val = sheet.getLastRow() >= 1 ? String(sheet.getRange(1, 1).getValue()).trim() : '';
    if (r1Val !== 'run_date') {
      // Salvage existing data (rows 2+ pokud vypadaji jako skutecne snapshot data)
      var salvaged = [];
      if (sheet.getLastRow() >= 2) {
        var candidateData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        for (var si = 0; si < candidateData.length; si++) {
          var col1 = candidateData[si][0];
          var col2 = String(candidateData[si][1] || '').trim();
          if (col1 && col2 && /^\d{4}-W\d{1,2}$/.test(col2)) {
            salvaged.push(candidateData[si]);
          }
        }
      }
      // Zkus salvage i z row 1 (pokud neni placeholder)
      if (sheet.getLastRow() >= 1 && r1Val && r1Val.substring(0, 1) !== '(') {
        var r1Data = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
        var r1Col2 = String(r1Data[1] || '').trim();
        if (/^\d{4}-W\d{1,2}$/.test(r1Col2)) {
          salvaged.unshift(r1Data);
        }
      }
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
        .setBackground('#174ea6').setFontColor('#ffffff');
      if (salvaged.length > 0) {
        sheet.getRange(2, 1, salvaged.length, headers.length).setValues(salvaged);
      }
    }

    var weekId = computeWeekId(runDate);
    var categoryCounts = {
      loser_rest: 0, low_ctr_audit: 0, DECLINING: 0, RISING: 0, LOST_OPPORTUNITY: 0
    };
    for (var ci = 0; ci < classified.length; ci++) {
      var label = classified[ci].primaryLabel;
      if (categoryCounts[label] !== undefined) categoryCounts[label]++;
    }

    var row = [
      Utils.formatDate(runDate), weekId,
      roundNumber(summary.accountBaseline.totalCost, 2),
      summary.accountBaseline.totalClicks,
      roundNumber(summary.accountBaseline.totalConversions, 2),
      roundNumber(summary.accountBaseline.totalConversionValue, 2),
      roundNumber(summary.accountBaseline.avgRoas, 2),
      roundNumber(summary.accountBaseline.avgPno, 2),
      roundNumber((summary.accountBaseline.avgCtr || 0) * 100, 3),
      summary.flags.totalFlagged,
      categoryCounts.loser_rest, categoryCounts.low_ctr_audit,
      categoryCounts.DECLINING, categoryCounts.RISING, categoryCounts.LOST_OPPORTUNITY,
      roundNumber(summary.flags.totalWastedSpend, 2),
      (effectiveness && effectiveness.transitions && effectiveness.transitions.RESOLVED) || 0,
      (effectiveness && effectiveness.transitions && effectiveness.transitions.RE_FLAGGED) || 0,
      (effectiveness && effectiveness.applicationRate !== null && effectiveness.applicationRate !== undefined)
        ? roundNumber(effectiveness.applicationRate, 1) : ''
    ];

    var lastRow = sheet.getLastRow();
    // Upsert logika: pokud existuje radek s week_id = current, prepis ho.
    // Jinak append novy radek.
    // POZN: Scannujeme VSECHNY radky (ne jen [2..lastRow-1]), protoze header
    // checkuje sheet.getLastRow() === 0 — pokud uz jsou v sheetu data ale
    // chybi header (napr. ze setupOutputSheet placeholderu), getLastRow > 0
    // ale radek 1 muze byt data. String() wrappuje pro bezpecne porovnani
    // (Apps Script getValues() muze vracet Date misto stringu).
    var targetRow = lastRow + 1;
    var action = 'appended';
    if (lastRow >= 1) {
      // Naskenuj VSECHNY radky v sloupci 2 (week_id)
      var existingWeeks = sheet.getRange(1, 2, lastRow, 1).getValues();
      var weekIdStr = String(weekId).trim();
      // Iterujeme od konce — pokud by byly duplikaty z pre-fix bugu, overwritneme posledni
      for (var ei = existingWeeks.length - 1; ei >= 0; ei--) {
        var cellStr = String(existingWeeks[ei][0] || '').trim();
        if (cellStr === weekIdStr && cellStr.length > 0) {
          targetRow = ei + 1; // +1 protoze range zacinal od row 1
          action = 'upserted';
          break;
        }
      }
    }

    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    // Force text format pro run_date (col 1) a week_id (col 2)
    sheet.getRange(targetRow, 1, 1, 2).setNumberFormat('@');

    // Filtr refresh jen pri append — upsert nemeni range dat
    if (action === 'appended') {
      addFilterToSheet(sheet, targetRow, row.length);
    }

    Logger.log('INFO: WEEKLY_SNAPSHOT ' + action + ' (week ' + weekId + ', row ' + targetRow + ').');
  }

  function computeWeekId(date) {
    var y = date.getFullYear();
    var firstDayOfYear = new Date(y, 0, 1);
    var pastDaysOfYear = (date - firstDayOfYear) / 86400000;
    var weekNum = Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
    return y + '-W' + (weekNum < 10 ? '0' : '') + weekNum;
  }

  /**
   * DETAIL tab — plny trace pro specialisty.
   */
  function writeDetailTab(ss, classified, config) {
    var sheetName = 'DETAIL';
    var sheet = getOrCreateSheet(ss, sheetName);
    sheet.clearContents();

    var headers = [
      'item_id',
      'primary_label',
      'secondary_flags',
      'reason_code',
      'tier',
      'status',
      'transition_type',
      'runs_since_first_flag',
      'flag_status'
    ];
    if (config.includeProductTitles) {
      headers.push('product_title');
    }
    headers = headers.concat([
      'product_brand',
      'product_type',
      'campaign_name',
      'campaigns_count',
      'primary_campaign_share_pct',
      'top_campaigns',
      'previous_campaign',
      'campaign_moved',
      'channel',
      'first_click_date',
      'age_days',
      'clicks',
      'impressions',
      'cost',
      'conversions',
      'conversion_value',
      'actual_PNO_pct',
      'actual_ROAS',
      'CTR_pct',
      'search_impression_share',
      'product_price',
      'price_source',
      'expected_CPA',
      'expected_conversions',
      'min_clicks_required',
      'passed_age_gate',
      'passed_sample_gate',
      'yoy_signal',
      'wasted_spend',
      'wasted_spend_pct',
      'note',
      'suggested_action'
    ]);

    var data = [headers];

    var maxRows = Math.min(classified.length, config.maxRowsDetailTab);
    for (var i = 0; i < maxRows; i++) {
      var c = classified[i];
      var flagStatus = c.transitionType === 'REPEATED' && c.runsSinceFirstFlag >= 2
        ? 'REPEATED_WARNING'
        : (c.transitionType === 'RE_FLAGGED' ? 'RE_FLAGGED_AFTER_RESOLVE' : (c.primaryLabel ? 'FLAGGED' : 'OK'));

      var row = [
        c.itemId,
        c.primaryLabel,
        c.secondaryFlags.join(', '),
        c.reasonCode,
        c.tier,
        c.status,
        c.transitionType,
        c.runsSinceFirstFlag,
        flagStatus
      ];
      if (config.includeProductTitles) {
        row.push(c.productTitle);
      }
      row = row.concat([
        c.productBrand,
        c.productType,
        c.campaignName,
        c.campaignsCount || 1,
        roundNumber(c.primaryCampaignSharePct || 100, 1),
        c.topCampaigns || c.campaignName,
        c.previousCampaign || '',
        c.campaignMoved,
        c.channel,
        c.firstClickDate,
        c.ageDays === null ? '' : c.ageDays,
        c.clicks,
        c.impressions,
        roundNumber(c.cost, 2),
        c.conversions,
        roundNumber(c.conversionValue, 2),
        roundNumber(c.actualPno, 2),
        roundNumber(c.roas, 4),
        roundNumber(c.ctr * 100, 4),
        roundNumber(c.searchImpressionShare * 100, 2),
        roundNumber(c.productPrice, 2),
        c.priceSource,
        roundNumber(c.expectedCpa, 2),
        roundNumber(c.expectedConversions, 2),
        c.minClicksRequired,
        c.passedAgeGate,
        c.passedSampleGate,
        c.yoySignal,
        roundNumber(c.wastedSpend, 2),
        roundNumber(c.wastedSpendPct, 2),
        c.note || '',
        c.suggestedAction || ''
      ]);
      data.push(row);
    }

    if (classified.length > config.maxRowsDetailTab) {
      data.push([
        '[TRUNCATED] ' + (classified.length - config.maxRowsDetailTab) + ' dalsich radku vynecano (maxRowsDetailTab limit)'
      ]);
    }

    writeDataInBatches(sheet, data);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2); // zmrazit item_id + primary_label (lepsi skim pri scrollovani)
    sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff').setFontSize(10);
    sheet.setRowHeight(1, 30);
    addFilterToSheet(sheet, data.length, data[0].length);

    // Podminene formatovani — zvyraznit problematicke/vyborne hodnoty
    // POZN: aplikujeme jen kdyz je dost radku (>1), jinak ConditionalFormat hodi error
    if (data.length > 1) {
      try {
        applyDetailConditionalFormat(sheet, data[0], data.length);
      } catch (e) {
        Logger.log('WARN: Conditional format DETAIL failed: ' + e.message);
      }
    }

    Logger.log('INFO: DETAIL — zapsano ' + (data.length - 1) + ' radku.');
  }

  /**
   * Aplikuje podminene formatovani na DETAIL tab:
   *   - actual_PNO_pct > 60 → cervena (hodne ztratove)
   *   - actual_PNO_pct 30-60 → oranzova (hranicni)
   *   - actual_ROAS >= 5 → zelena (vyborne)
   *   - primary_label != '' → jemny zluty background na cele radce
   */
  function applyDetailConditionalFormat(sheet, headers, totalRows) {
    var rules = sheet.getConditionalFormatRules();
    // Clear existing (idempotentnost)
    sheet.clearConditionalFormatRules();

    // Helper: najdi index sloupce podle jmena
    function colIdx(name) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i] === name) return i + 1; // 1-indexed
      }
      return -1;
    }

    var pnoCol = colIdx('actual_PNO_pct');
    var roasCol = colIdx('actual_ROAS');
    var priceSourceCol = colIdx('price_source');

    var newRules = [];

    // PNO > 60 = red
    if (pnoCol > 0) {
      var pnoRange = sheet.getRange(2, pnoCol, totalRows - 1, 1);
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(60)
          .setBackground('#fce4e4').setFontColor('#c5221f').setBold(true)
          .setRanges([pnoRange]).build()
      );
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(30, 60)
          .setBackground('#fff4e5').setFontColor('#a56200')
          .setRanges([pnoRange]).build()
      );
    }

    // ROAS >= 5 = green, 3-5 = light green
    if (roasCol > 0) {
      var roasRange = sheet.getRange(2, roasCol, totalRows - 1, 1);
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(5)
          .setBackground('#d4f5d4').setFontColor('#1e6e1e').setBold(true)
          .setRanges([roasRange]).build()
      );
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(3, 5)
          .setBackground('#eaf7ea').setFontColor('#2e7d32')
          .setRanges([roasRange]).build()
      );
    }

    // price_source = unavailable → zlute varovani
    if (priceSourceCol > 0) {
      var psRange = sheet.getRange(2, priceSourceCol, totalRows - 1, 1);
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('unavailable')
          .setBackground('#fff4e5').setFontColor('#a56200').setItalic(true)
          .setRanges([psRange]).build()
      );
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('gmc_feed')
          .setFontColor('#1e6e1e')
          .setRanges([psRange]).build()
      );
    }

    sheet.setConditionalFormatRules(newRules);
  }

  /**
   * SUMMARY tab — vizualni dashboard s KPI cards, grafy, barvami a formatovanim.
   */
  function writeSummaryTab(ss, summary, classified, effectiveness, config, runDate) {
    var sheetName = 'SUMMARY';
    var sheet = getOrCreateSheet(ss, sheetName);

    // Uklid: odstran stare grafy pred novym runem (idempotence)
    var existingCharts = sheet.getCharts();
    for (var ec = 0; ec < existingCharts.length; ec++) {
      sheet.removeChart(existingCharts[ec]);
    }
    sheet.clear();

    // Nastav sirky sloupcu pro hezky vzhled
    // Column A: dlouhe labely (napr. "LOSER_REST — low_volume (1–3 conv + PNO ≥ 1.5× target)")
    sheet.setColumnWidth(1, 360);
    sheet.setColumnWidth(2, 170);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 180);
    sheet.setColumnWidth(7, 180);

    // Barvy
    var COLOR_PRIMARY = '#1a73e8';
    var COLOR_PRIMARY_LIGHT = '#e8f0fe';
    var COLOR_DANGER = '#d93025';
    var COLOR_DANGER_LIGHT = '#fce8e6';
    var COLOR_SUCCESS = '#188038';
    var COLOR_SUCCESS_LIGHT = '#e6f4ea';
    var COLOR_WARNING = '#f29900';
    var COLOR_WARNING_LIGHT = '#fef7e0';
    var COLOR_GRAY = '#5f6368';
    var COLOR_GRAY_LIGHT = '#f8f9fa';

    // Precalc hodnoty
    var flaggedTotalCost = summary.flaggedTotalCost || 0;
    var flaggedTotalConversions = summary.flaggedTotalConversions || 0;
    var flaggedTotalConvValue = summary.flaggedTotalConvValue || 0;
    var pctFlaggedCost = summary.accountBaseline.totalCost > 0
      ? (flaggedTotalCost / summary.accountBaseline.totalCost) * 100 : 0;
    var pctFlaggedProducts = summary.funnel.afterAggregation > 0
      ? (summary.flags.totalFlagged / summary.funnel.afterAggregation) * 100 : 0;
    var pctWastedOfTotal = summary.accountBaseline.totalCost > 0
      ? (summary.flags.totalWastedSpend / summary.accountBaseline.totalCost) * 100 : 0;
    var lookbackDays = summary.funnel.lookbackDays || config.lookbackDays || 30;
    var yearlySavings = summary.flags.totalWastedSpend * (365 / lookbackDays);
    var currency = summary.currency;

    var row = 1;

    // ============================================================
    // HLAVIČKA (řádky 1-3)
    // ============================================================
    sheet.getRange(row, 1, 1, 7).merge()
      .setValue('SHOPPING/PMAX LOSER DETECTOR')
      .setBackground(COLOR_PRIMARY)
      .setFontColor('#ffffff')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 42);
    row++;

    sheet.getRange(row, 1, 1, 7).merge()
      .setValue(summary.accountName + '  •  ' + summary.customerId + '  •  ' + Utils.formatDate(summary.lookbackStart) + ' — ' + Utils.formatDate(summary.lookbackEnd) + ' (' + lookbackDays + ' dnů)  •  Měna: ' + currency)
      .setBackground('#174ea6')
      .setFontColor('#ffffff')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 28);
    row++;

    // Legenda barev + quick nav (diskretní proužky pod hlavičkou)
    sheet.getRange(row, 1, 1, 7).merge()
      .setValue('🔵 Metrika účtu    🟠 Označené / varování    🔴 Ztráty / kritické    🟢 Vyřešeno / výborné    ⚪ Neutrální / pozadí')
      .setBackground('#f8f9fa')
      .setFontColor('#5f6368')
      .setFontSize(10)
      .setFontStyle('italic')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 24);
    row++;

    // Quick nav placeholder — vyplni se na konci s klikatelnymi hyperlinks
    var quickNavRow = row;
    sheet.getRange(row, 1, 1, 7).merge()
      .setBackground('#e8f0fe')
      .setFontColor('#1a73e8')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 26);
    row++;

    row++; // prázdný řádek

    // Track pozice sekci pro hyperlinks (zacina tady, tracking napojen pred kazdou sekci)
    var sectionPositions = {
      kpi: 6  // KPI karty jsou nad quick navem (hlavni KPI zacina row 6)
    };

    // ============================================================
    // KPI KARTY — redesign (4 karty, každá 4 řádky: label / value / value / sub)
    // Lepší typografická hierarchie — label malý, hodnota velká, sub střední
    // ============================================================
    var kpiRow = row;

    // Helper: napiš KPI kartu se strukturovaným layoutem
    function writeKpiCard(startCol, colSpan, label, value, subtext, bgColor, borderColor, labelColor, valueColor) {
      // Label (1 row, 9pt, uppercase, subtle)
      var labelRange = sheet.getRange(kpiRow, startCol, 1, colSpan);
      labelRange.merge();
      labelRange.setValue(label)
        .setBackground(bgColor)
        .setFontColor(labelColor)
        .setFontSize(9)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      // Value (2 rows, 22pt, bold)
      var valueRange = sheet.getRange(kpiRow + 1, startCol, 2, colSpan);
      valueRange.merge();
      valueRange.setValue(value)
        .setBackground(bgColor)
        .setFontColor(valueColor)
        .setFontSize(22)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      // Sub (1 row, 10pt, italic)
      var subRange = sheet.getRange(kpiRow + 3, startCol, 1, colSpan);
      subRange.merge();
      subRange.setValue(subtext)
        .setBackground(bgColor)
        .setFontColor(labelColor)
        .setFontSize(10)
        .setFontStyle('italic')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      // Tlustý rámeček kolem celé karty
      sheet.getRange(kpiRow, startCol, 4, colSpan)
        .setBorder(true, true, true, true, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }

    // Karta 1: CELKOVÉ NÁKLADY (account)
    writeKpiCard(
      1, 2,
      'CELKOVÉ NÁKLADY ÚČTU',
      formatMoney(summary.accountBaseline.totalCost, currency),
      summary.accountBaseline.totalConversions.toFixed(0) + ' konverzí  •  ROAS ' + roundNumber(summary.accountBaseline.avgRoas, 2),
      COLOR_PRIMARY_LIGHT, COLOR_PRIMARY, COLOR_PRIMARY, COLOR_PRIMARY
    );

    // Karta 2: NÁKLADY NA OZNAČENÉ
    writeKpiCard(
      3, 2,
      'NÁKLADY NA OZNAČENÉ PRODUKTY',
      formatMoney(flaggedTotalCost, currency),
      summary.flags.totalFlagged + ' produktů  •  ' + roundNumber(pctFlaggedCost, 1) + '% z účtu',
      COLOR_WARNING_LIGHT, COLOR_WARNING, '#a56200', '#a56200'
    );

    // Karta 3: MARNÝ SPEND
    writeKpiCard(
      5, 2,
      'MARNÝ SPEND (vs cílové ROAS)',
      formatMoney(summary.flags.totalWastedSpend, currency),
      roundNumber(pctWastedOfTotal, 1) + '% z celkových nákladů',
      COLOR_DANGER_LIGHT, COLOR_DANGER, COLOR_DANGER, COLOR_DANGER
    );

    // Karta 4: POTENCIÁLNÍ ÚSPORA / rok
    writeKpiCard(
      7, 1,
      'POTENCIÁLNÍ ÚSPORA / ROK',
      formatMoney(yearlySavings, currency),
      'roční extrapolace',
      COLOR_SUCCESS_LIGHT, COLOR_SUCCESS, COLOR_SUCCESS, COLOR_SUCCESS
    );

    // Row heights pro KPI karty
    sheet.setRowHeight(kpiRow, 20);       // label
    sheet.setRowHeight(kpiRow + 1, 26);   // value (1/2)
    sheet.setRowHeight(kpiRow + 2, 22);   // value (2/2)
    sheet.setRowHeight(kpiRow + 3, 22);   // sub

    row = kpiRow + 4;
    row++; // blank
    row++; // extra blank pro vzdušnost

    // ============================================================
    // KLASIFIKAČNÍ FUNNEL
    // ============================================================
    sectionPositions.funnel = row;
    row = writeSectionHeader(sheet, row, 'KLASIFIKAČNÍ TRYCHTÝŘ', COLOR_PRIMARY);

    var healthyProductsCount = Math.max(0, (summary.funnel.classified || 0) - (summary.flags.totalFlagged || 0));
    // Spocitej min clicks prahy pro edukativni insight
    var minClicksThresholdCalc = Math.max(
      config.minClicksAbsolute || 30,
      Math.ceil((config.minExpectedConvFloor || 1) / Math.max(summary.accountBaseline.cvr || 0.001, 0.001))
    );

    var funnelData = [
      ['Řádků produkt-kampaň (před filtrem)', formatInt(summary.funnel.rawRows), ''],
      ['Vyloučeno — brand kampaně (' + config.brandCampaignPattern + ')', formatInt(summary.funnel.brandExcluded), pct(summary.funnel.brandExcluded, summary.funnel.rawRows)],
      ['Vyloučeno — rest kampaně (' + config.restCampaignPattern + ')', formatInt(summary.funnel.restExcluded), pct(summary.funnel.restExcluded, summary.funnel.rawRows)],
      ['Vyloučeno — pozastavené kampaně', formatInt(summary.funnel.pausedExcluded), pct(summary.funnel.pausedExcluded, summary.funnel.rawRows)],
      ['Zbývá po filtru (před agregací)', formatInt(summary.funnel.keptRows), pct(summary.funnel.keptRows, summary.funnel.rawRows)],
      ['Po agregaci (unikátních item_id)', formatInt(summary.funnel.afterAggregation), ''],
      ['Přeskočeno — nové produkty (< ' + config.minProductAgeDays + ' dní)', formatInt(summary.funnel.tooYoung), pct(summary.funnel.tooYoung, summary.funnel.afterAggregation)],
      ['Přeskočeno — málo dat (< ' + minClicksThresholdCalc + ' kliků)', formatInt(summary.funnel.insufficientData), pct(summary.funnel.insufficientData, summary.funnel.afterAggregation)],
      ['Přeskočeno — problém s kvalitou dat', formatInt(summary.funnel.dataQualityIssues), pct(summary.funnel.dataQualityIssues, summary.funnel.afterAggregation)],
      ['🔍 Klasifikováno (celkem)', formatInt(summary.funnel.classified), pct(summary.funnel.classified, summary.funnel.afterAggregation)],
      ['    ├─ ⚠️ Označené (problémové)', formatInt(summary.flags.totalFlagged), pct(summary.flags.totalFlagged, summary.funnel.classified)],
      ['    └─ ✅ Zdravé (bez problémů)', formatInt(healthyProductsCount), pct(healthyProductsCount, summary.funnel.classified)]
    ];
    row = writeKeyValueTable(sheet, row, funnelData, COLOR_GRAY_LIGHT);

    // Informační insight box — vysvětlí user proč je tolik produktů v insufficient data
    var insufPct = summary.funnel.afterAggregation > 0
      ? Math.round(summary.funnel.insufficientData / summary.funnel.afterAggregation * 100)
      : 0;
    if (insufPct >= 20) {
      sheet.getRange(row, 1, 1, 7).merge()
        .setValue('ℹ️  ' + insufPct + '% produktů (' + formatInt(summary.funnel.insufficientData) + ') nemá dost dat pro klasifikaci — mají < ' + minClicksThresholdCalc + ' kliků za ' + (summary.funnel.lookbackDays || config.lookbackDays) + ' dní. Je to přirozené pro Shopping longtail (90% produktů = nízká frekvence). Threshold zabraňuje false positives (např. "1 konverze mezi 10 kliky" není statistická událost). Pro agresivnější zachycení: sniž minClicksAbsolute v CONFIG.')
        .setBackground('#e8f0fe').setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic')
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
      sheet.setRowHeight(row, 60);
      row++;
    }
    row++; // blank

    // ============================================================
    // FLAGGED BREAKDOWN + PIE CHART
    // ============================================================
    row = writeSectionHeader(sheet, row, 'ROZDĚLENÍ OZNAČENÝCH PRODUKTŮ', COLOR_WARNING);

    // LOSER_REST tiers
    var loserByTier = summary.flags.loserByTier;
    var tierData = [
      ['LOSER_REST — zero_conv (0 konv + high spend)', loserByTier.zero_conv || 0, pct(loserByTier.zero_conv || 0, summary.flags.totalFlagged)],
      ['LOSER_REST — low_volume (1–3 conv + PNO ≥ 1.5× target)', loserByTier.low_volume || 0, pct(loserByTier.low_volume || 0, summary.flags.totalFlagged)],
      ['LOSER_REST — mid_volume (4–10 conv + PNO ≥ 2.0× target)', loserByTier.mid_volume || 0, pct(loserByTier.mid_volume || 0, summary.flags.totalFlagged)],
      ['LOSER_REST — high_volume (11+ conv + PNO ≥ 3.0× target)', loserByTier.high_volume || 0, pct(loserByTier.high_volume || 0, summary.flags.totalFlagged)],
      ['LOW_CTR_AUDIT — irrelevant_keyword_match', summary.flags.lowCtrByReason.irrelevant_keyword_match || 0, pct(summary.flags.lowCtrByReason.irrelevant_keyword_match || 0, summary.flags.totalFlagged)],
      ['LOW_CTR_AUDIT — high_visibility_low_appeal', summary.flags.lowCtrByReason.high_visibility_low_appeal || 0, pct(summary.flags.lowCtrByReason.high_visibility_low_appeal || 0, summary.flags.totalFlagged)],
      ['LOW_CTR_AUDIT — low_ctr_general', summary.flags.lowCtrByReason.low_ctr_general || 0, pct(summary.flags.lowCtrByReason.low_ctr_general || 0, summary.flags.totalFlagged)],
      ['Overlap (loser + low_ctr)', summary.flags.overlap, pct(summary.flags.overlap, summary.flags.totalFlagged)]
    ];
    row = writeKeyValueTable(sheet, row, tierData, COLOR_WARNING_LIGHT);

    // Insert PIE CHART — umisten MIMO data (sloupec I+) aby neprekryval text
    if (summary.flags.totalFlagged > 0) {
      var chartDataStartRow = row + 1;
      sheet.getRange(chartDataStartRow, 9, 1, 2).setValues([['Kategorie', 'Počet']]);
      var chartRows = [
        ['Ztrátové — 0 konv', loserByTier.zero_conv || 0],
        ['Ztrátové — nízký objem', loserByTier.low_volume || 0],
        ['Ztrátové — střední objem', loserByTier.mid_volume || 0],
        ['Ztrátové — vysoký objem', loserByTier.high_volume || 0],
        ['Nízké CTR', summary.flags.lowCtrTotal]
      ];
      sheet.getRange(chartDataStartRow + 1, 9, chartRows.length, 2).setValues(chartRows);
      try { sheet.hideColumns(9, 2); } catch (he) { /* optional */ }

      var pieChart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange(chartDataStartRow, 9, chartRows.length + 1, 2))
        // Pozice: column 9 = mimo datovou oblast (A-G)
        .setPosition(kpiRow + 5, 9, 0, 0)
        .setOption('title', 'Rozdělení označených produktů podle kategorie')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('pieHole', 0.4)
        .setOption('colors', [COLOR_DANGER, '#ea8600', COLOR_WARNING, '#9aa0a6', '#1a73e8'])
        .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
        .setOption('width', 520)
        .setOption('height', 320)
        .build();
      sheet.insertChart(pieChart);
    }
    row++; // blank

    // ============================================================
    // IMPACT TABULKA (detailní rozpad)
    // ============================================================
    row = writeSectionHeader(sheet, row, 'DOPAD — KOLIK INVESTUJEME DO OZNAČENÝCH PRODUKTŮ', COLOR_DANGER);

    var impactData = [
      ['Počet označených / klasifikováno', formatInt(summary.flags.totalFlagged) + ' / ' + formatInt(summary.funnel.classified), roundNumber(pctFlaggedProducts, 1) + '%'],
      ['Náklady na označené / celkové náklady', formatMoney(flaggedTotalCost, currency) + ' / ' + formatMoney(summary.accountBaseline.totalCost, currency), roundNumber(pctFlaggedCost, 1) + '%'],
      ['Marný spend / celkové náklady', formatMoney(summary.flags.totalWastedSpend, currency) + ' / ' + formatMoney(summary.accountBaseline.totalCost, currency), roundNumber(pctWastedOfTotal, 1) + '%'],
      ['Marný spend / náklady na označené', formatMoney(summary.flags.totalWastedSpend, currency) + ' / ' + formatMoney(flaggedTotalCost, currency), flaggedTotalCost > 0 ? roundNumber((summary.flags.totalWastedSpend / flaggedTotalCost) * 100, 1) + '%' : '–'],
      ['Konverze z označených / celkové konverze', roundNumber(flaggedTotalConversions, 1) + ' / ' + roundNumber(summary.accountBaseline.totalConversions, 1), summary.accountBaseline.totalConversions > 0 ? roundNumber((flaggedTotalConversions / summary.accountBaseline.totalConversions) * 100, 1) + '%' : '–'],
      ['Revenue z označených / celkové revenue', formatMoney(flaggedTotalConvValue, currency) + ' / ' + formatMoney(summary.accountBaseline.totalConversionValue, currency), summary.accountBaseline.totalConversionValue > 0 ? roundNumber((flaggedTotalConvValue / summary.accountBaseline.totalConversionValue) * 100, 1) + '%' : '–'],
      ['📅 Potenciální roční úspora (extrapolace)', formatMoney(yearlySavings, currency), 'marný × 365 / ' + lookbackDays + ' dní']
    ];
    row = writeKeyValueTable(sheet, row, impactData, COLOR_DANGER_LIGHT);

    // BAR CHART — Account cost breakdown (flagged vs rest) — mimo data (sloupec I+)
    if (summary.accountBaseline.totalCost > 0) {
      var barDataStartRow = row + 1;
      var nonFlaggedCost = summary.accountBaseline.totalCost - flaggedTotalCost;
      sheet.getRange(barDataStartRow, 11, 3, 2).setValues([
        ['Kategorie', 'Náklady (' + currency + ')'],
        ['Označené produkty', flaggedTotalCost],
        ['Ostatní produkty', nonFlaggedCost]
      ]);
      try { sheet.hideColumns(11, 2); } catch (he2) { /* optional */ }

      var barChart = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange(barDataStartRow, 11, 3, 2))
        // Pozice: column 9 = mimo datovou oblast
        .setPosition(row - 6, 9, 0, 0)
        .setOption('title', 'Rozpad nákladů: Označené vs Ostatní')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('colors', [COLOR_DANGER, '#34a853'])
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', 220)
        .setOption('hAxis', { format: 'short' })
        .build();
      sheet.insertChart(barChart);
    }
    row++; // blank

    // ============================================================
    // ACCOUNT BASELINE
    // ============================================================
    row = writeSectionHeader(sheet, row, 'ZÁKLADNÍ METRIKY ÚČTU (mimo brand/rest kampaně)', COLOR_GRAY);

    var baselineData = [
      ['Celkové náklady', formatMoney(summary.accountBaseline.totalCost, currency), ''],
      ['Celkem kliků', formatInt(summary.accountBaseline.totalClicks), ''],
      ['Celkem zobrazení', formatInt(summary.accountBaseline.totalImpressions), ''],
      ['Celkem konverzí', formatInt(Math.round(summary.accountBaseline.totalConversions)), ''],
      ['Celková hodnota konverzí', formatMoney(summary.accountBaseline.totalConversionValue, currency), ''],
      ['Průměrné CVR účtu', Utils.safePctFormat(summary.accountBaseline.cvr * 100), ''],
      ['Průměrné CTR účtu', Utils.safePctFormat(summary.accountBaseline.avgCtr * 100), ''],
      ['Průměrné CPC', formatMoney(summary.accountBaseline.avgCpc, currency), ''],
      ['Průměrné ROAS', roundNumber(summary.accountBaseline.avgRoas, 2), ''],
      ['Průměrné PNO', Utils.safePctFormat(summary.accountBaseline.avgPno), ''],
      ['Průměrná hodnota objednávky (AOV)', formatMoney(summary.accountBaseline.avgAov, currency), '']
    ];
    row = writeKeyValueTable(sheet, row, baselineData, COLOR_GRAY_LIGHT);
    row++; // blank

    // ============================================================
    // TRANSITIONS
    // ============================================================
    if (effectiveness) {
      row = writeSectionHeader(sheet, row, 'ZMĚNY V TOMTO BĚHU (přechody stavů)', COLOR_SUCCESS);
      // Pozn.: hodnoty prevadime na formatInt (string) aby sheet neaplikoval % format
      // z predchozich tabulek (ROZDĚLENÍ). Jinak by se ze 32 stalo 3200%.
      var transData = [
        ['🆕 NEW_FLAG — Poprvé označený produkt', formatInt(effectiveness.transitions.NEW_FLAG || 0), ''],
        ['🔁 REPEATED — Stále označený (label nebyl aplikován?)', formatInt(effectiveness.transitions.REPEATED || 0), ''],
        ['✅ RESOLVED — Přesunut do rest kampaně (úspěšný zásah)', formatInt(effectiveness.transitions.RESOLVED || 0), ''],
        ['🌱 UN_FLAGGED — Produkt se zlepšil sám (bez zásahu)', formatInt(effectiveness.transitions.UN_FLAGGED || 0), ''],
        ['⚠️ RE_FLAGGED — Vrátil se po vyřešení', formatInt(effectiveness.transitions.RE_FLAGGED || 0), ''],
        ['🔄 CATEGORY_CHANGE — Změnil kategorii', formatInt(effectiveness.transitions.CATEGORY_CHANGE || 0), ''],
        ['⚠️ REPEATED_WARNING — Označený ≥ 2 běhy bez změny', formatInt(effectiveness.repeatedWarning || 0), '']
      ];
      row = writeKeyValueTable(sheet, row, transData, COLOR_SUCCESS_LIGHT);
      if (effectiveness.applicationRate !== null) {
        var appRate = effectiveness.applicationRate || 0;
        // Barva podle hodnoty: >= 50% zelena, 20-50% oranzova, < 20% cervena
        var rateColor = appRate >= 50 ? COLOR_SUCCESS : (appRate >= 20 ? COLOR_WARNING : COLOR_DANGER);
        var rateBackground = appRate >= 50 ? COLOR_SUCCESS_LIGHT : (appRate >= 20 ? COLOR_WARNING_LIGHT : COLOR_DANGER_LIGHT);
        sheet.getRange(row, 1).setValue('Míra aplikace labelů:').setFontWeight('bold').setBackground(rateBackground);
        sheet.getRange(row, 2).setValue(Utils.safePctFormat(appRate))
          .setFontWeight('bold').setFontColor(rateColor).setBackground(rateBackground).setFontSize(13);
        sheet.getRange(row, 3).setValue('(' + effectiveness.resolvedThisRun + ' vyřešeno / ' + effectiveness.labeledLastRun + ' označeno v minulém běhu)')
          .setBackground(rateBackground).setFontStyle('italic');
        sheet.setRowHeight(row, 26);
        row++;
      }
      row++; // blank
    }

    // ============================================================
    // TOP 10 LOSERS
    // ============================================================
    sectionPositions.topLosers = row;
    row = writeSectionHeader(sheet, row, 'TOP 10 ZTRÁTOVÝCH PRODUKTŮ (dle marného spendu)', COLOR_DANGER);

    var topLosers = summary.topLosers || [];
    var losersTableStartRow = row;
    var losersHeader = ['item_id', 'reason_code', 'tier', 'cost', 'conversions', 'actual_PNO_%', 'wasted_spend'];
    sheet.getRange(row, 1, 1, 7).setValues([losersHeader])
      .setBackground(COLOR_GRAY_LIGHT).setFontWeight('bold').setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    row++;
    if (topLosers.length === 0) {
      sheet.getRange(row, 1, 1, 7).merge().setValue('Žádný LOSER_REST produkt nebyl označen.').setFontStyle('italic').setFontColor(COLOR_GRAY).setHorizontalAlignment('center');
      row++;
    } else {
      for (var li = 0; li < topLosers.length && li < 10; li++) {
        var ll = topLosers[li];
        sheet.getRange(row, 1, 1, 7).setValues([[
          ll.itemId, ll.reasonCode, ll.tier,
          roundNumber(ll.cost, 2), ll.conversions,
          roundNumber(ll.actualPno, 2), roundNumber(ll.wastedSpend, 2)
        ]]).setBorder(null, true, null, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
        // Reset formatu pro text sloupce (item_id, reason_code, tier) — jinak dedi z predchozich tabulek
        sheet.getRange(row, 1, 1, 3).setNumberFormat('@');
        sheet.getRange(row, 5).setNumberFormat('0.00');  // conversions (numeric)
        // Zvýrazni wasted_spend červenou
        sheet.getRange(row, 7).setFontColor(COLOR_DANGER).setFontWeight('bold');
        sheet.getRange(row, 4).setNumberFormat('#,##0.00 "' + currency + '"');
        sheet.getRange(row, 7).setNumberFormat('#,##0.00 "' + currency + '"');
        sheet.getRange(row, 6).setNumberFormat('0.00"%"');
        if (li % 2 === 1) {
          sheet.getRange(row, 1, 1, 7).setBackground('#fafafa');
        }
        row++;
      }

      // SCATTER PLOT — Cost vs wasted spend (top losers)
      // X = cost, Y = wasted_spend, kazdy produkt = bod
      // Pozice: vpravo od tabulky (sloupec I+)
      var scatterStart = row + 1;
      var scatterData = [['Cost (' + currency + ')', 'Marný spend (' + currency + ')']];
      for (var si = 0; si < topLosers.length && si < 10; si++) {
        scatterData.push([roundNumber(topLosers[si].cost, 0), roundNumber(topLosers[si].wastedSpend, 0)]);
      }
      sheet.getRange(scatterStart, 15, scatterData.length, 2).setValues(scatterData);
      try { sheet.hideColumns(15, 2); } catch (he3) { /* ok */ }

      var scatterChart = sheet.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(sheet.getRange(scatterStart, 15, scatterData.length, 2))
        .setPosition(losersTableStartRow, 9, 0, 0)
        .setOption('title', 'Top 10 ztrátových: Náklad vs Marný spend')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('hAxis', { title: 'Náklady na produkt (' + currency + ')', textStyle: { fontSize: 9 } })
        .setOption('vAxis', { title: 'Marný spend (' + currency + ')', textStyle: { fontSize: 9 } })
        .setOption('colors', [COLOR_DANGER])
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', 320)
        .setOption('pointSize', 10)
        .build();
      sheet.insertChart(scatterChart);
    }
    row++; // blank
    row++; // extra

    // ============================================================
    // TOP 10 LOW-CTR
    // ============================================================
    row = writeSectionHeader(sheet, row, 'TOP 10 NÍZKÉHO CTR (dle zobrazení)', COLOR_WARNING);

    var topLowCtr = summary.topLowCtr || [];
    var lowCtrHeader = ['item_id', 'reason_code', 'impressions', 'CTR_%', 'baseline_CTR_%', 'IS_%'];
    sheet.getRange(row, 1, 1, 6).setValues([lowCtrHeader])
      .setBackground(COLOR_GRAY_LIGHT).setFontWeight('bold').setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    row++;
    if (topLowCtr.length === 0) {
      sheet.getRange(row, 1, 1, 6).merge().setValue('Žádný LOW_CTR produkt nebyl flagged.').setFontStyle('italic').setFontColor(COLOR_GRAY).setHorizontalAlignment('center');
      row++;
    } else {
      for (var lci = 0; lci < topLowCtr.length && lci < 10; lci++) {
        var lc = topLowCtr[lci];
        var baselinePct = (summary.accountBaseline.avgCtr || 0) * 100;
        if (config.ctrBaselineScope === 'campaign' && summary.perCampaignBaseline && summary.perCampaignBaseline[lc.campaignId]) {
          baselinePct = summary.perCampaignBaseline[lc.campaignId].ctr * 100;
        }
        sheet.getRange(row, 1, 1, 6).setValues([[
          lc.itemId, lc.reasonCode, lc.impressions,
          roundNumber(lc.ctr * 100, 3), roundNumber(baselinePct, 3), roundNumber(lc.searchImpressionShare * 100, 1)
        ]]).setBorder(null, true, null, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
        // Explicitni format (override z predchozi tabulky, ktera meta CZK)
        sheet.getRange(row, 3).setNumberFormat('#,##0');                    // impressions
        sheet.getRange(row, 4).setNumberFormat('0.000"%"').setFontColor(COLOR_DANGER).setFontWeight('bold'); // CTR_%
        sheet.getRange(row, 5).setNumberFormat('0.000"%"');                 // baseline_CTR_%
        sheet.getRange(row, 6).setNumberFormat('0.0"%"');                   // IS_%
        if (lci % 2 === 1) {
          sheet.getRange(row, 1, 1, 6).setBackground('#fafafa');
        }
        row++;
      }
    }
    row++; // blank

    // ============================================================
    // WEEKLY TRENDS (TABULKA + LINE CHART — trend na prvni pohled)
    // ============================================================
    sectionPositions.trends = row;
    row = writeSectionHeader(sheet, row, 'TÝDENNÍ TRENDY (posledních 8 týdnů)', COLOR_PRIMARY);
    var trends = readWeeklyTrends(ss, 8);
    if (trends.length < 2) {
      sheet.getRange(row, 1, 1, 7).merge()
        .setValue('(čekáme na více dat — potřeba min 2 týdny)')
        .setFontStyle('italic').setFontColor(COLOR_GRAY).setHorizontalAlignment('center');
      row++;
    } else {
      // TABULKA (kompaktni, 5 sloupcu)
      var trendsHeader = ['Týden', 'ROAS', 'PNO %', 'Marný spend', 'Označené'];
      var trendsTableStartRow = row;
      sheet.getRange(row, 1, 1, 5).setValues([trendsHeader])
        .setBackground(COLOR_GRAY_LIGHT).setFontWeight('bold')
        .setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
      row++;

      // Sber dat pro chart
      var chartDataRows = [['Týden', 'ROAS', 'PNO %', 'Marný spend (tis. ' + currency + ')']];

      for (var ti = 0; ti < trends.length; ti++) {
        var t = trends[ti];
        sheet.getRange(row, 1, 1, 5).setValues([[
          t.week_id,
          roundNumber(t.account_roas, 2),
          roundNumber(t.account_pno_pct, 1),
          roundNumber(t.wasted_spend_total, 0),
          t.flagged_count_total
        ]]).setBorder(null, true, null, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(row, 2).setNumberFormat('0.00');
        sheet.getRange(row, 3).setNumberFormat('0.0"%"');
        sheet.getRange(row, 4).setNumberFormat('#,##0 "' + currency + '"');
        sheet.getRange(row, 5).setNumberFormat('#,##0');
        if (ti % 2 === 1) {
          sheet.getRange(row, 1, 1, 5).setBackground('#fafafa');
        }

        // Chart data — wasted v tisicich pro lepsi skladani s ROAS (ne v CZK)
        chartDataRows.push([
          t.week_id,
          roundNumber(t.account_roas, 2),
          roundNumber(t.account_pno_pct, 1),
          roundNumber(t.wasted_spend_total / 1000, 1)
        ]);

        row++;
      }

      // LINE CHART — pozice vpravo od tabulky (mimo data area)
      var chartSourceStart = row + 1;
      sheet.getRange(chartSourceStart, 13, chartDataRows.length, 4).setValues(chartDataRows);
      try { sheet.hideColumns(13, 4); } catch (he) { /* ok */ }

      var trendsChart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sheet.getRange(chartSourceStart, 13, chartDataRows.length, 4))
        .setPosition(trendsTableStartRow, 9, 0, 0)
        .setOption('title', 'Vývoj klíčových metrik v čase')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('colors', [COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER])
        .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
        .setOption('width', 520)
        .setOption('height', 280)
        .setOption('hAxis', { title: 'Týden', textStyle: { fontSize: 9 } })
        .setOption('vAxis', { textStyle: { fontSize: 9 } })
        .setOption('pointSize', 5)
        .setOption('curveType', 'function')
        .build();
      sheet.insertChart(trendsChart);
    }
    row++; // blank
    row++; // extra

    // ============================================================
    // VÝVOJ OZNAČENÍ V ČASE — stacked area chart per kategorie
    // Odpovidá: "Resíme problémy, nebo přibývají?"
    // ============================================================
    if (trends.length >= 2) {
      sectionPositions.labelsEvolution = row;
      row = writeSectionHeader(sheet, row, 'VÝVOJ OZNAČENÍ V ČASE (řešíme, nebo přibývá?)', COLOR_PRIMARY);

      // Data pro stacked area — per kategorie per tyden
      var labelsEvoData = [['Týden', 'loser_rest', 'low_ctr_audit', 'DECLINING', 'RISING', 'LOST_OPPORTUNITY']];
      for (var tle = 0; tle < trends.length; tle++) {
        var te = trends[tle];
        labelsEvoData.push([
          te.week_id || '?',
          te.flagged_loser_rest || 0,
          te.flagged_low_ctr || 0,
          te.flagged_declining || 0,
          te.flagged_rising || 0,
          te.flagged_lost_opp || 0
        ]);
      }

      var evoChartStart = row + 1;
      sheet.getRange(evoChartStart, 19, labelsEvoData.length, 6).setValues(labelsEvoData);
      try { sheet.hideColumns(19, 6); } catch (he5) { /* ok */ }

      var evoChart = sheet.newChart()
        .setChartType(Charts.ChartType.AREA)
        .addRange(sheet.getRange(evoChartStart, 19, labelsEvoData.length, 6))
        .setPosition(row, 1, 0, 0)
        .setOption('title', 'Kolik produktů je označeno podle kategorie (týden co týden)')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('isStacked', true)
        .setOption('colors', [COLOR_DANGER, COLOR_WARNING, '#ea8600', COLOR_SUCCESS, '#1a73e8'])
        .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
        .setOption('width', 960)
        .setOption('height', 320)
        .setOption('hAxis', { title: 'Týden', textStyle: { fontSize: 9 } })
        .setOption('vAxis', { title: 'Počet označených', textStyle: { fontSize: 9 } })
        .setOption('areaOpacity', 0.7)
        .build();
      sheet.insertChart(evoChart);

      // Reserve 18 rows for chart space
      row += 18;

      // NEW vs RESOLVED chart — kolik pridavame vs res
      var flowChartRows = [['Týden', 'Nově označené', 'Vyřešené (přesun do rest)']];
      for (var tfl = 0; tfl < trends.length; tfl++) {
        var tf = trends[tfl];
        flowChartRows.push([
          tf.week_id || '?',
          0, // TODO: potrebujeme z WEEKLY_SNAPSHOT pridat sloupec 'new_flags_this_run'
          tf.resolved_this_run || 0
        ]);
      }

      // Sub-section header: Příliv vs odliv
      sheet.getRange(row, 1, 1, 7).merge()
        .setValue('💡  Jak to číst: Plocha roste = přibývají problémy. Klesá = úspěšně je řešíme. Stabilní = rovnováha mezi novými a vyřešenými.')
        .setBackground('#e8f0fe').setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic')
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      sheet.setRowHeight(row, 26);
      row++;
      row++; // blank
    }

    // ============================================================
    // VYKON PODLE TYPU KAMPANE (tabulka + column chart)
    // Agreguje per-kampan: kolik produktu tam je, kolik cost, wasted, flagged
    // Dava vedet: kde se peníze ztrácejí nejvic? (napr. Assets Akce a slevy)
    // ============================================================
    var campaignPerf = computeCampaignPerformance(classified);
    if (campaignPerf.rows.length > 0) {
      sectionPositions.campaigns = row;
      row = writeSectionHeader(sheet, row, 'VÝKON PODLE KAMPANĚ (kde se peníze ztrácejí)', COLOR_WARNING);

      var campPerfHeader = ['Kampaň', 'Produktů', 'Označených', 'Náklady', 'Marný spend', 'Marný %'];
      var campPerfTableStart = row;
      sheet.getRange(row, 1, 1, 6).setValues([campPerfHeader])
        .setBackground(COLOR_GRAY_LIGHT).setFontWeight('bold')
        .setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
      row++;

      var chartRowsCamp = [['Kampaň', 'Marný spend', 'Náklady']];

      for (var cpi = 0; cpi < campaignPerf.rows.length && cpi < 10; cpi++) {
        var cp = campaignPerf.rows[cpi];
        var wastedPct = cp.cost > 0 ? (cp.wastedSpend / cp.cost) * 100 : 0;
        sheet.getRange(row, 1, 1, 6).setValues([[
          cp.campaignName,
          cp.totalProducts,
          cp.flaggedProducts,
          roundNumber(cp.cost, 0),
          roundNumber(cp.wastedSpend, 0),
          roundNumber(wastedPct, 1)
        ]]).setBorder(null, true, null, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(row, 2).setNumberFormat('#,##0');
        sheet.getRange(row, 3).setNumberFormat('#,##0');
        sheet.getRange(row, 4).setNumberFormat('#,##0 "' + currency + '"');
        sheet.getRange(row, 5).setNumberFormat('#,##0 "' + currency + '"').setFontColor(COLOR_DANGER).setFontWeight('bold');
        sheet.getRange(row, 6).setNumberFormat('0.0"%"');
        // Barevny indikator marneho spendu %
        if (wastedPct >= 10) {
          sheet.getRange(row, 6).setBackground('#fce4e4').setFontColor(COLOR_DANGER).setFontWeight('bold');
        } else if (wastedPct >= 5) {
          sheet.getRange(row, 6).setBackground('#fff4e5').setFontColor('#a56200');
        } else {
          sheet.getRange(row, 6).setBackground('#eaf7ea').setFontColor(COLOR_SUCCESS);
        }
        if (cpi % 2 === 1) {
          sheet.getRange(row, 1, 1, 5).setBackground('#fafafa');
        }
        chartRowsCamp.push([cp.campaignName.slice(0, 40), roundNumber(cp.wastedSpend, 0), roundNumber(cp.cost, 0)]);
        row++;
      }

      // COLUMN CHART — Wasted spend per kampan (vedle tabulky)
      var campChartDataStart = row + 1;
      sheet.getRange(campChartDataStart, 17, chartRowsCamp.length, 3).setValues(chartRowsCamp);
      try { sheet.hideColumns(17, 3); } catch (he4) { /* ok */ }

      var campChart = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(sheet.getRange(campChartDataStart, 17, chartRowsCamp.length, 3))
        .setPosition(campPerfTableStart, 9, 0, 0)
        .setOption('title', 'Marný spend a náklady podle kampaně')
        .setOption('titleTextStyle', { fontSize: 13, bold: true })
        .setOption('colors', [COLOR_DANGER, '#5f6368'])
        .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
        .setOption('width', 520)
        .setOption('height', 320)
        .setOption('hAxis', { textStyle: { fontSize: 8 }, slantedText: true, slantedTextAngle: 30 })
        .setOption('vAxis', { format: 'short', textStyle: { fontSize: 9 } })
        .build();
      sheet.insertChart(campChart);

      row++; // blank
      row++; // extra
    }

    // ============================================================
    // EFFECTIVENESS (aggregate pres vsechny resolved produkty)
    // ============================================================
    var aggEff = computeAggregateEffectiveness(ss);
    if (aggEff && aggEff.totalEvaluated > 0) {
      row = writeSectionHeader(sheet, row, 'ÚČINNOST ZÁSAHŮ (přes všechny vyřešené produkty)', COLOR_SUCCESS);
      var effData = [
        ['Produkty s + (úspěšná intervence)', formatInt(aggEff.counts['+']), pct(aggEff.counts['+'], aggEff.totalEvaluated)],
        ['Produkty s = (smíšený výsledek)', formatInt(aggEff.counts['=']), pct(aggEff.counts['='], aggEff.totalEvaluated)],
        ['Produkty s − (intervence škodí)', formatInt(aggEff.counts['-']), pct(aggEff.counts['-'], aggEff.totalEvaluated)],
        ['⌛ PENDING (< 14 dní od vyřešení)', formatInt(aggEff.counts['PENDING']), ''],
        ['N/A (bez dostatečných dat)', formatInt(aggEff.counts['N/A']), ''],
        ['Průměrná změna nákladů', Utils.safePctFormat(aggEff.avgDeltaCost), ''],
        ['Průměrná změna ROAS', Utils.safePctFormat(aggEff.avgDeltaRoas), '']
      ];
      row = writeKeyValueTable(sheet, row, effData, COLOR_SUCCESS_LIGHT);
      row++; // blank
    }

    // ============================================================
    // REST CAMPAIGN HEALTH
    // ============================================================
    row = writeSectionHeader(sheet, row, 'ZDRAVÍ REST KAMPANÍ (účinnost přesunu loserů)', COLOR_WARNING);
    var restHealth = computeRestCampaignHealth(classified, config);
    var efficientThresholdPct = Math.round((config.restCampaignEfficientThreshold || 0.2) * 100);
    var restData = [
      ['Produktů v rest:', restHealth.counts.total, ''],
      ['Rest cost celkem:', formatMoney(restHealth.totalRestCost, currency), ''],
      ['  Efficient (≤ ' + efficientThresholdPct + '% pre-cost)', restHealth.counts.efficient, pct(restHealth.counts.efficient, restHealth.counts.total)],
      ['  Acceptable (≤ 50%)', restHealth.counts.acceptable, pct(restHealth.counts.acceptable, restHealth.counts.total)],
      ['  ⚠ Wasteful (> 50%)', restHealth.counts.wasteful, pct(restHealth.counts.wasteful, restHealth.counts.total)]
    ];
    row = writeKeyValueTable(sheet, row, restData, COLOR_WARNING_LIGHT);
    row++; // blank

    // ============================================================
    // BRAND INSIGHTS
    // ============================================================
    row = writeSectionHeader(sheet, row, 'VHLEDY Z BRAND KAMPANÍ', COLOR_PRIMARY);
    var bi = computeBrandInsights(classified);
    var brandData = [
      ['Brand-only sellers (conv jen přes brand):', bi.brandOnlyCount, ''],
      ['Revenue z brand-only:', formatMoney(bi.brandOnlyRevenue, currency), ''],
      ['Brand-dependent (>50% conv přes brand):', bi.brandDependent.length, '']
    ];
    row = writeKeyValueTable(sheet, row, brandData, COLOR_PRIMARY_LIGHT);
    if (bi.brandOnlyCount > 0) {
      sheet.getRange(row, 1, 1, 7).merge()
        .setValue('💡 Insight: Tyto produkty prodávají jen přes brand — kandidáti na marketing awareness / promo.')
        .setFontSize(10).setFontStyle('italic').setFontColor(COLOR_PRIMARY)
        .setBackground(COLOR_PRIMARY_LIGHT).setWrap(true).setHorizontalAlignment('center');
      sheet.setRowHeight(row, 28);
      row++;
    }
    row++; // blank

    // ============================================================
    // NOVÉ PRODUKTY (< minProductAgeDays dní, chráněné před klasifikací)
    // ============================================================
    if (summary.newProducts && summary.newProducts.total > 0) {
      var np = summary.newProducts;
      var minAgeDays = (summary.config && summary.config.minProductAgeDays) || 30;
      sectionPositions.newProducts = row;
      row = writeSectionHeader(sheet, row, 'NOVÉ PRODUKTY (< ' + minAgeDays + ' dní — v ramp-up)', COLOR_PRIMARY);

      var avgAgeStr = (np.avgAge || 0).toFixed(1) + ' dní';
      var ageRangeStr = (np.minAge !== null ? np.minAge : '?') + ' – ' + (np.maxAge !== null ? np.maxAge : '?') + ' dní';
      var newProductsData = [
        ['Celkem nových produktů', np.total, ''],
        ['Průměrné stáří', avgAgeStr, ''],
        ['Rozsah stáří', ageRangeStr, ''],
        ['S cenou ve feedu', np.withGmcFeedPrice, pct(np.withGmcFeedPrice, np.total)],
        ['Získaly 10+ kliků', np.withClicks10Plus, pct(np.withClicks10Plus, np.total)],
        ['Získaly 1+ konverze', np.withFirstConv, pct(np.withFirstConv, np.total) + ' (rising star candidates)'],
        ['Celkový cost', formatMoney(np.totalCost, currency), ''],
        ['Celkový revenue', formatMoney(np.totalConversionValue, currency), ''],
        ['Souhrnné ROAS', np.totalCost > 0 ? (np.totalConversionValue / np.totalCost).toFixed(2) : '—', '']
      ];
      row = writeKeyValueTable(sheet, row, newProductsData, COLOR_PRIMARY_LIGHT);

      // Rising star candidates table (top 10)
      if (np.topRisingCandidates && np.topRisingCandidates.length > 0) {
        sheet.getRange(row, 1, 1, 7).merge()
          .setValue('🌟 Top ' + Math.min(10, np.topRisingCandidates.length) + ' Rising Star Candidates (10+ kliků + 1+ konverze)')
          .setFontSize(11).setFontWeight('bold').setFontColor(COLOR_PRIMARY)
          .setBackground(COLOR_PRIMARY_LIGHT).setHorizontalAlignment('center');
        sheet.setRowHeight(row, 26);
        row++;

        var risingHeader = ['item_id', 'age (dny)', 'clicks', 'conv', 'cost', 'revenue', 'ROAS'];
        sheet.getRange(row, 1, 1, 7).setValues([risingHeader])
          .setBackground(COLOR_GRAY_LIGHT).setFontWeight('bold')
          .setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
        row++;

        for (var ri = 0; ri < np.topRisingCandidates.length && ri < 10; ri++) {
          var rc = np.topRisingCandidates[ri];
          sheet.getRange(row, 1, 1, 7).setValues([[
            rc.itemId,
            rc.ageDays,
            rc.clicks,
            roundNumber(rc.conversions, 2),
            roundNumber(rc.cost, 2),
            roundNumber(rc.conversionValue, 2),
            roundNumber(rc.roas, 2)
          ]]).setBorder(null, true, null, true, true, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
          sheet.getRange(row, 5).setNumberFormat('#,##0.00 "' + currency + '"');
          sheet.getRange(row, 6).setNumberFormat('#,##0.00 "' + currency + '"');
          // Zvyrazni ROAS zelene pokud >= 3
          if ((rc.roas || 0) >= 3) {
            sheet.getRange(row, 7).setFontColor(COLOR_SUCCESS).setFontWeight('bold');
          }
          if (ri % 2 === 1) {
            sheet.getRange(row, 1, 1, 7).setBackground('#fafafa');
          }
          row++;
        }

        sheet.getRange(row, 1, 1, 7).merge()
          .setValue('💡 Insight: Tyto produkty slibně startují. Po dosažení ' + minAgeDays + ' dní budou plně klasifikovány. Zvaž bid-boost nebo dedikovanou kampaň.')
          .setFontSize(10).setFontStyle('italic').setFontColor(COLOR_PRIMARY)
          .setBackground(COLOR_PRIMARY_LIGHT).setWrap(true).setHorizontalAlignment('center');
        sheet.setRowHeight(row, 28);
        row++;
      }
      row++; // blank
    }

    // ============================================================
    // CONFIG AUDIT TRAIL (menší, diskrétní)
    // ============================================================
    row = writeSectionHeader(sheet, row, 'POUŽITÁ KONFIGURACE (pro audit)', COLOR_GRAY);
    var configData = [
      ['targetPnoPct', config.targetPnoPct],
      ['lookbackDays', config.lookbackDays],
      ['minClicksAbsolute', config.minClicksAbsolute],
      ['minExpectedConvFloor', config.minExpectedConvFloor],
      ['minProductAgeDays', config.minProductAgeDays],
      ['brandCampaignPattern', config.brandCampaignPattern],
      ['restCampaignPattern', config.restCampaignPattern],
      ['customLabelIndex', config.customLabelIndex],
      ['ctrBaselineScope', config.ctrBaselineScope],
      ['dryRun', config.dryRun]
    ];
    for (var ci = 0; ci < configData.length; ci++) {
      sheet.getRange(row, 1).setValue(configData[ci][0]).setFontSize(9).setFontColor(COLOR_GRAY);
      sheet.getRange(row, 2).setValue(configData[ci][1]).setFontSize(9).setFontColor(COLOR_GRAY).setFontFamily('Roboto Mono');
      row++;
    }
    row++;

    // Disclaimer
    sheet.getRange(row, 1, 1, 7).merge()
      .setValue('⚠️ Attribution disclaimer: Metriky reflektují default Last-Click model Scripts API. Některé hodnoty se mohou lehce lišit od Google Ads UI při jiném attribution modelu.')
      .setFontSize(9).setFontStyle('italic').setFontColor(COLOR_GRAY)
      .setBackground('#fafafa').setWrap(true).setHorizontalAlignment('center');
    sheet.setRowHeight(row, 36);
    row++;

    // ============================================================
    // QUICK NAV — plain text prehled sekci (bez hyperlinku)
    // Hyperlinky byly odstraneny — klik v Google Sheets (RichText i
    // HYPERLINK) nescrolloval spolehlive v ramci tabu. User scrolluje
    // rucne. Radek tu ponechavame jako vizualni TOC (obsah).
    // ============================================================
    try {
      sheet.getRange(quickNavRow, 1).setValue(
        '📊 KPI  │  🔽 Trychtýř  │  📉 Top ztrátové  │  📈 Trendy  │  📊 Vývoj štítků  │  🎯 Kampaně  │  🌟 Nové produkty'
      );
    } catch (qne) {
      Logger.log('WARN: Quick nav write failed: ' + qne.message);
    }

    Logger.log('INFO: SUMMARY — zapsano ' + row + ' radku s vizualizacemi.');
  }

  /**
   * Precte poslednich N radku z WEEKLY_SNAPSHOT tabu a vrati pole objektu
   * {header: value}. Pokud sheet chybi nebo ma < 2 radky, vraci prazdne pole.
   */
  function readWeeklyTrends(ss, maxWeeks) {
    var sheet = ss.getSheetByName('WEEKLY_SNAPSHOT');
    if (!sheet || sheet.getLastRow() < 2) return [];

    var lastRow = sheet.getLastRow();
    var startRow = Math.max(2, lastRow - maxWeeks + 1);
    var numRows = lastRow - startRow + 1;
    var lastCol = sheet.getLastColumn();

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

    var result = [];
    for (var r = 0; r < data.length; r++) {
      var obj = {};
      for (var i = 0; i < headers.length; i++) {
        obj[headers[i]] = data[r][i];
      }
      result.push(obj);
    }
    return result;
  }

  /**
   * REST campaign health — rozdeli produkty v rest kampanich podle efficiency.
   * efficient: rest_cost / main_cost <= config.restCampaignEfficientThreshold (default 0.2)
   * acceptable: <= 0.5
   * wasteful: > 0.5
   */
  /**
   * Agreguje vykon per kampan: kolik produktu, kolik oznacenych, kolik cost, wasted spend.
   * Pouziva primaryCampaignName — kazdy produkt spadne do primary kampane.
   * Vraci top 10 kampani serazenych podle wasted spend desc.
   */
  function computeCampaignPerformance(classified) {
    var campaignMap = {}; // name → { totalProducts, flaggedProducts, cost, wastedSpend }

    for (var i = 0; i < classified.length; i++) {
      var c = classified[i];
      var key = c.campaignName || '(bez kampaně)';
      if (!campaignMap[key]) {
        campaignMap[key] = {
          campaignName: key,
          totalProducts: 0,
          flaggedProducts: 0,
          cost: 0,
          wastedSpend: 0
        };
      }
      campaignMap[key].totalProducts++;
      if (c.primaryLabel) campaignMap[key].flaggedProducts++;
      campaignMap[key].cost += (c.cost || 0);
      campaignMap[key].wastedSpend += (c.wastedSpend || 0);
    }

    var rows = [];
    for (var k in campaignMap) {
      if (campaignMap.hasOwnProperty(k)) {
        rows.push(campaignMap[k]);
      }
    }
    // Sort podle wasted_spend desc
    rows.sort(function (a, b) { return b.wastedSpend - a.wastedSpend; });
    return { rows: rows };
  }

  function computeRestCampaignHealth(classified, config) {
    var counts = { total: 0, efficient: 0, acceptable: 0, wasteful: 0 };
    var totalRestCost = 0;
    var efficientThreshold = (config && config.restCampaignEfficientThreshold) || 0.2;

    for (var i = 0; i < classified.length; i++) {
      var c = classified[i];
      if (!c.rest_metrics || c.rest_metrics.cost <= 0) continue;
      counts.total++;
      totalRestCost += c.rest_metrics.cost;

      // Proxy: srovnani aktualni main_cost vs rest_cost. Idealne by to bylo
      // kpi_before_main_cost z timeline, ale pro MVP staci current main.
      var mainCost = c.main_metrics ? c.main_metrics.cost : 0;
      var ratio = mainCost > 0 ? (c.rest_metrics.cost / mainCost) : 999;

      if (ratio <= efficientThreshold) counts.efficient++;
      else if (ratio <= 0.5) counts.acceptable++;
      else counts.wasteful++;
    }

    return { counts: counts, totalRestCost: totalRestCost };
  }

  /**
   * Brand insights — najde brand-only sellers a brand-dependent produkty.
   * brand_only_sellers: main_conv = 0 AND brand_conv > 0 (produkt se prodava jen pres brand)
   * brand_dependent: brand_share z total revenue > 0.5
   */
  function computeBrandInsights(classified) {
    var brandOnlySellers = [];
    var brandDependent = [];

    for (var i = 0; i < classified.length; i++) {
      var c = classified[i];
      if (!c.total_metrics || c.total_metrics.conversions < 1) continue;

      var mainConv = c.main_metrics ? c.main_metrics.conversions : 0;
      var brandConv = c.brand_metrics ? c.brand_metrics.conversions : 0;
      var brandRev = c.brand_metrics ? c.brand_metrics.conversionValue : 0;
      var totalRev = c.total_metrics.conversionValue;

      if (mainConv === 0 && brandConv > 0) {
        brandOnlySellers.push({ itemId: c.itemId, revenue: brandRev });
      }

      var brandShare = totalRev > 0 ? (brandRev / totalRev) : 0;
      if (brandShare > 0.5) {
        brandDependent.push({ itemId: c.itemId, share: brandShare });
      }
    }

    var brandOnlyRevenue = 0;
    for (var j = 0; j < brandOnlySellers.length; j++) {
      brandOnlyRevenue += brandOnlySellers[j].revenue || 0;
    }

    return {
      brandOnlySellers: brandOnlySellers,
      brandDependent: brandDependent,
      brandOnlyCount: brandOnlySellers.length,
      brandOnlyRevenue: brandOnlyRevenue
    };
  }

  /**
   * Aggregate effectiveness pres cely PRODUCT_TIMELINE tab — pocet produktu v kazdem
   * skore (+ / = / - / PENDING / N/A) + prumerne delta_cost_pct / delta_roas_pct
   * pres evaluated produkty (+/=/-). Vraci null pokud sheet chybi nebo je prazdny.
   */
  function computeAggregateEffectiveness(ss) {
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

  /**
   * Zapise section header (merged, s barvou) a vrati dalsi row.
   */
  function writeSectionHeader(sheet, row, title, color) {
    sheet.getRange(row, 1, 1, 7).merge()
      .setValue('  ▸  ' + title)
      .setBackground(color)
      .setFontColor('#ffffff')
      .setFontSize(13)
      .setFontWeight('bold')
      .setFontFamily('Montserrat')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(row, 34);
    return row + 1;
  }

  /**
   * Zapise key-value tabulku (3 sloupce: label, value, extra).
   * Vrati dalsi row.
   */
  function writeKeyValueTable(sheet, row, data, bgColor) {
    if (data.length === 0) return row;
    var values = [];
    for (var d = 0; d < data.length; d++) {
      values.push([data[d][0], data[d][1], data[d][2] || '']);
    }
    sheet.getRange(row, 1, data.length, 3).setValues(values)
      .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    // Alternating row colors
    for (var rr = 0; rr < data.length; rr++) {
      if (rr % 2 === 0) {
        sheet.getRange(row + rr, 1, 1, 3).setBackground('#ffffff');
      } else {
        sheet.getRange(row + rr, 1, 1, 3).setBackground('#fafafa');
      }
    }

    // Label column — normal text
    sheet.getRange(row, 1, data.length, 1).setFontSize(10);
    // Value column — bold + reset formatu na plain (jinak by dedil % nebo CZK z predchozi tabulky)
    sheet.getRange(row, 2, data.length, 1)
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setNumberFormat('@'); // text = neaplikovat zadny implicitni format
    // Extra column — smaller, gray
    sheet.getRange(row, 3, data.length, 1).setFontSize(9).setFontColor('#5f6368').setHorizontalAlignment('right').setNumberFormat('@');

    return row + data.length;
  }

  /**
   * Formatuje penize s thousand separators a currency suffix.
   */
  function formatMoney(n, currency) {
    if (n === null || n === undefined || isNaN(n)) return '–';
    var rounded = Math.round(n);
    var parts = String(rounded).split('');
    var out = [];
    for (var i = parts.length - 1, c = 0; i >= 0; i--, c++) {
      if (c > 0 && c % 3 === 0) out.unshift(' ');
      out.unshift(parts[i]);
    }
    return out.join('') + ' ' + (currency || 'Kč');
  }

  function formatInt(n) {
    if (n === null || n === undefined || isNaN(n)) return '–';
    var rounded = Math.round(n);
    var parts = String(rounded).split('');
    var out = [];
    for (var i = parts.length - 1, c = 0; i >= 0; i--, c++) {
      if (c > 0 && c % 3 === 0) out.unshift(' ');
      out.unshift(parts[i]);
    }
    return out.join('');
  }

  function pct(num, denom) {
    if (!denom || denom === 0) return '';
    return roundNumber((num / denom) * 100, 1) + '%';
  }

  /**
   * Vyrovna pocty sloupcu napric radky (pro setValues()).
   */
  function normalizeRowLengths(rows) {
    var maxLen = 1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].length > maxLen) {
        maxLen = rows[i].length;
      }
    }
    for (var j = 0; j < rows.length; j++) {
      while (rows[j].length < maxLen) {
        rows[j].push('');
      }
    }
  }

  /**
   * LIFECYCLE_LOG — dashboard nahore + append-only data s dedup.
   *
   * Layout:
   *   Row 1    — Title banner
   *   Row 2    — Subtitle (total transitions, unique products, runs)
   *   Row 3    — spacer
   *   Row 4-5  — 6 KPI karet (NEW_FLAG / REPEATED / RESOLVED / UN_FLAGGED / RE_FLAGGED / CAT_CHANGE)
   *   Row 6    — spacer
   *   Row 7    — Info "Tento mesic vs predchozi mesic"
   *   Row 8    — spacer
   *   Row 9    — Data header (run_date, item_id, ...)
   *   Row 10+  — Data (append-only + dedup)
   *
   * Dashboard se rebuildue pri kazdem runu z aktualnich dat (cte VSECHNA data).
   * Dedup: per (run_date, item_id, transition_type) — nove radky co match
   * existing key preskocime (same-day retry chrani pred duplicity).
   */
  function appendLifecycleLogTab(ss, classified, runDate) {
    var sheetName = 'LIFECYCLE_LOG';
    var sheet = getOrCreateSheet(ss, sheetName);
    var todayStr = Utils.formatDate(runDate);

    var DATA_START_ROW = 10;
    var HEADER_ROW = 9;
    var headers = [
      'run_date', 'item_id', 'current_label', 'previous_label', 'transition_type',
      'current_campaign', 'previous_campaign', 'campaign_moved',
      'cost_30d', 'conversions_30d', 'pno_30d_pct', 'roas_30d', 'ctr_30d_pct',
      'reason_code', 'tier', 'runs_since_first_flag', 'notes'
    ];
    var NUM_COLS = headers.length;

    // === 1. DETECT EXISTING LAYOUT a NAČTI data ===
    // Dva scenare:
    //   a) Novy layout — data od row DATA_START_ROW, header na HEADER_ROW
    //   b) Stary layout — data od row 2, header na row 1
    //   c) Prazdny sheet — data nejsou
    var lastRow = sheet.getLastRow();
    var existingData = [];
    if (lastRow >= 2) {
      // Detekce: cti row HEADER_ROW A1 — pokud match 'run_date', novy layout
      var r9Value = sheet.getRange(HEADER_ROW, 1).getValue();
      var r1Value = sheet.getRange(1, 1).getValue();
      var dataFromRow;
      if (String(r9Value).trim() === 'run_date') {
        dataFromRow = DATA_START_ROW;
      } else if (String(r1Value).trim() === 'run_date') {
        dataFromRow = 2; // stary layout — migration
      } else {
        dataFromRow = null;
      }

      if (dataFromRow !== null && lastRow >= dataFromRow) {
        var raw = sheet.getRange(dataFromRow, 1, lastRow - dataFromRow + 1, NUM_COLS).getValues();
        // Filter prazdnych radku + DEDUP existing data (ochrana proti historickym duplicitam
        // z pre-fix verze, kdy Sheets locale prepnul ISO -> US format a dedup match selhal).
        var seenKeys = {};
        var cleanupSkipped = 0;
        for (var ri = 0; ri < raw.length; ri++) {
          if (raw[ri][0] === '' || raw[ri][0] === null) continue;
          // Normalize date na ISO pred dedup
          var dNorm = Utils.normalizeDate(raw[ri][0]);
          raw[ri][0] = dNorm; // prepis hodnotu na ISO (konzistentni format do dat)
          var itemNorm = String(raw[ri][1] || '').trim();
          var transNorm = String(raw[ri][4] || '').trim();
          var seenKey = dNorm + '|' + itemNorm + '|' + transNorm;
          if (seenKeys[seenKey]) {
            cleanupSkipped++;
            continue;
          }
          seenKeys[seenKey] = true;
          existingData.push(raw[ri]);
        }
        if (cleanupSkipped > 0) {
          Logger.log('INFO: LIFECYCLE_LOG — cleanup: ' + cleanupSkipped + ' historickych duplicit odstraneno (date format mismatch).');
        }
      }
    }

    // === 2. DEDUP SET (pro porovnani s new transitions) ===
    var existingKeys = {};
    for (var ei = 0; ei < existingData.length; ei++) {
      var dateStr = Utils.normalizeDate(existingData[ei][0]);
      var itemIdEx = String(existingData[ei][1] || '').trim();
      var transEx = String(existingData[ei][4] || '').trim();
      if (dateStr && itemIdEx && transEx) {
        existingKeys[dateStr + '|' + itemIdEx + '|' + transEx] = true;
      }
    }

    // === 3. FILTER NOVE RADKY (transitions != NO_CHANGE + not in dedup set) ===
    var newRows = [];
    var skippedDup = 0;
    for (var i = 0; i < classified.length; i++) {
      var c = classified[i];
      if (c.transitionType === 'NO_CHANGE' || c.transitionType === '') {
        continue;
      }
      var key = todayStr + '|' + c.itemId + '|' + c.transitionType;
      if (existingKeys[key]) {
        skippedDup++;
        continue;
      }
      newRows.push([
        todayStr,
        c.itemId,
        c.primaryLabel || '',
        c.previousLabel || '',
        c.transitionType,
        c.campaignName || '',
        c.previousCampaign || '',
        c.campaignMoved,
        roundNumber(c.cost, 2),
        c.conversions,
        roundNumber(c.actualPno, 2),
        roundNumber(c.roas, 4),
        roundNumber(c.ctr * 100, 4),
        c.reasonCode || '',
        c.tier || '',
        c.runsSinceFirstFlag || 0,
        c.note || ''
      ]);
    }

    // === 4. COMBINE: existing + new ===
    var allData = existingData.concat(newRows);

    // === 5. SPOCTI STATS pro dashboard ===
    var stats = computeLifecycleStats(allData, runDate);

    // === 6. REBUILD SHEET (clear + write dashboard + header + data) ===
    sheet.clearContents();
    try { sheet.clearConditionalFormatRules(); } catch (cfe) { /* ok */ }
    try {
      var eFilter = sheet.getFilter();
      if (eFilter) eFilter.remove();
    } catch (fe) { /* ok */ }

    // Unmerge vsechno (cleanup z predchozich layoutu)
    try { sheet.getRange(1, 1, HEADER_ROW, NUM_COLS).breakApart(); } catch (be) { /* ok */ }

    // 6a. DASHBOARD
    writeLifecycleDashboard(sheet, stats, NUM_COLS);

    // 6b. HEADER (row 9)
    sheet.getRange(HEADER_ROW, 1, 1, NUM_COLS).setValues([headers])
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#174ea6')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, null, null, '#0b3d91', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(HEADER_ROW, 28);
    sheet.setFrozenRows(HEADER_ROW);

    // 6c. DATA (row 10+)
    if (allData.length > 0) {
      sheet.getRange(DATA_START_ROW, 1, allData.length, NUM_COLS).setValues(allData);
      // Force text format pro run_date sloupec (col 1) — aby Google Sheets
      // nepreformatovalo ISO na locale string (4/21/2026 misto 2026-04-21).
      // Bez tohoto: auto-format podle locale uzivatele, nekonzistentni napric runy.
      sheet.getRange(DATA_START_ROW, 1, allData.length, 1).setNumberFormat('@');

      // Filter na data range
      try {
        sheet.getRange(HEADER_ROW, 1, allData.length + 1, NUM_COLS).createFilter();
      } catch (fre) { /* ok */ }
    }

    Logger.log('INFO: LIFECYCLE_LOG — rebuilt with dashboard, ' + newRows.length + ' novych' +
               (skippedDup > 0 ? ' (' + skippedDup + ' duplicit preskocenych)' : '') +
               ', total ' + allData.length + ' transitions.');
  }

  /**
   * Spocti statistiky z LIFECYCLE_LOG pro dashboard.
   *
   * Vraci:
   *   - lastRun: transitions z NEJNOVEJSIHO run_date (actionable pro tydenni review)
   *   - currentMonth / prevMonth: transitions za cely mesic (pro mesicni review)
   *   - total / uniqueItemCount / runCount: kumulativni historie (kontext v subtitle)
   *
   * Dashboard KPI karty pouzivaji lastRun (ne kumulativni totals) — jinak by cisla
   * rostla do nekonecna pri pravidelnem spousteni a ztratila actionable vyznam.
   */
  function computeLifecycleStats(data, runDate) {
    var stats = {
      total: data.length,
      transitions: { NEW_FLAG: 0, REPEATED: 0, RESOLVED: 0, UN_FLAGGED: 0, RE_FLAGGED: 0, CATEGORY_CHANGE: 0 },
      lastRun: { NEW_FLAG: 0, REPEATED: 0, RESOLVED: 0, UN_FLAGGED: 0, RE_FLAGGED: 0, CATEGORY_CHANGE: 0 },
      uniqueItems: {},
      runDates: {},
      currentMonth: { NEW_FLAG: 0, RESOLVED: 0, RE_FLAGGED: 0 },
      prevMonth: { NEW_FLAG: 0, RESOLVED: 0, RE_FLAGGED: 0 }
    };

    var cmStr = Utils.formatDate(runDate).substring(0, 7); // YYYY-MM
    var pmDate = new Date(runDate.getFullYear(), runDate.getMonth() - 1, 1);
    var pmStr = Utils.formatDate(pmDate).substring(0, 7);

    // Prvni prochod: najit NEJNOVEJSI date v data (latest run)
    var latestDate = '';
    for (var li = 0; li < data.length; li++) {
      var dStr = Utils.normalizeDate(data[li][0]);
      if (dStr && dStr > latestDate) latestDate = dStr;
    }

    for (var i = 0; i < data.length; i++) {
      var d = data[i];
      var dateStr = Utils.normalizeDate(d[0]);
      var monthKey = dateStr.substring(0, 7);
      var transition = String(d[4] || '').trim();
      var itemId = String(d[1] || '').trim();

      if (stats.transitions[transition] !== undefined) stats.transitions[transition]++;
      if (itemId) stats.uniqueItems[itemId] = true;
      if (dateStr) stats.runDates[dateStr] = true;

      // Last run transitions (actionable)
      if (dateStr === latestDate && stats.lastRun[transition] !== undefined) {
        stats.lastRun[transition]++;
      }

      if (monthKey === cmStr && stats.currentMonth[transition] !== undefined) {
        stats.currentMonth[transition]++;
      }
      if (monthKey === pmStr && stats.prevMonth[transition] !== undefined) {
        stats.prevMonth[transition]++;
      }
    }

    stats.uniqueItemCount = Object.keys(stats.uniqueItems).length;
    stats.runCount = Object.keys(stats.runDates).length;
    stats.currentMonthStr = cmStr;
    stats.prevMonthStr = pmStr;
    stats.latestRunDate = latestDate;
    return stats;
  }

  /**
   * Zapise dashboard (rows 1-8) do LIFECYCLE_LOG tabu.
   */
  function writeLifecycleDashboard(sheet, stats, numCols) {
    // Row 1: Title
    sheet.getRange(1, 1, 1, numCols).merge()
      .setValue('📜  LIFECYCLE_LOG — historie přechodů stavů produktů')
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat');
    sheet.setRowHeight(1, 36);

    // Row 2: Subtitle — info o POSLEDNIM RUNU + kumulativni kontext vpravo
    var subtitleText = '📅 Poslední run: ' + (stats.latestRunDate || '—') +
                       '  ·  Historie: ' + formatInt(stats.total) + ' transitions, ' +
                       formatInt(stats.uniqueItemCount) + ' unikátních produktů, ' +
                       formatInt(stats.runCount) + ' runs';
    sheet.getRange(2, 1, 1, numCols).merge()
      .setValue(subtitleText)
      .setBackground('#e8f0fe').setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(2, 22);

    // Row 3: spacer
    sheet.setRowHeight(3, 8);

    // Rows 4-5: 6 KPI karet — POSLEDNI RUN (ne kumulativni historie!)
    // Kumulativni totals by rostly do nekonecna pri pravidelnem tydnem spousteni.
    // "Poslední run" = actionable pohled pro tydenni review.
    var kpiCats = [
      ['🆕 NEW_FLAG', stats.lastRun.NEW_FLAG, '#e8f0fe', '#1a73e8'],
      ['🔁 REPEATED', stats.lastRun.REPEATED, '#fff4e5', '#a56200'],
      ['✅ RESOLVED', stats.lastRun.RESOLVED, '#e6f4ea', '#1e8e3e'],
      ['🌱 UN_FLAGGED', stats.lastRun.UN_FLAGGED, '#e6f4ea', '#1e8e3e'],
      ['⚠️ RE_FLAGGED', stats.lastRun.RE_FLAGGED, '#fce8e6', '#c5221f'],
      ['🔄 CAT_CHANGE', stats.lastRun.CATEGORY_CHANGE, '#f1f3f4', '#5f6368']
    ];

    // Header row (row 4)
    for (var ci = 0; ci < kpiCats.length; ci++) {
      sheet.getRange(4, ci * 2 + 1, 1, 2).merge()
        .setValue(kpiCats[ci][0])
        .setBackground('#fafafa').setFontSize(9).setFontWeight('bold').setFontColor('#5f6368')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, false, true, null, null, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(4, 22);

    // Value row (row 5)
    for (var ci2 = 0; ci2 < kpiCats.length; ci2++) {
      sheet.getRange(5, ci2 * 2 + 1, 1, 2).merge()
        .setValue(formatInt(kpiCats[ci2][1]))
        .setBackground(kpiCats[ci2][2]).setFontColor(kpiCats[ci2][3])
        .setFontSize(18).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, true, true, true, null, null, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    }
    sheet.setRowHeight(5, 36);

    // Row 6: spacer
    sheet.setRowHeight(6, 8);

    // Row 7: Tento mesic vs predchozi mesic
    var cm = stats.currentMonth;
    var pm = stats.prevMonth;
    var monthCompText = '📅 Tento měsíc (' + stats.currentMonthStr + '): ' +
      cm.NEW_FLAG + ' nových, ' + cm.RESOLVED + ' vyřešených, ' + cm.RE_FLAGGED + ' vrácených  │  ' +
      'Předchozí měsíc (' + stats.prevMonthStr + '): ' +
      pm.NEW_FLAG + ' nových, ' + pm.RESOLVED + ' vyřešených, ' + pm.RE_FLAGGED + ' vrácených';
    sheet.getRange(7, 1, 1, numCols).merge()
      .setValue(monthCompText)
      .setBackground('#fff8e1').setFontColor('#5f6368').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheet.setRowHeight(7, 24);

    // Row 8: spacer
    sheet.setRowHeight(8, 10);
  }

  /**
   * MONTHLY_REVIEW — mesicni agregace transitions z LIFECYCLE_LOG.
   * 1 radek per mesic, fresh rebuild pri kazdem runu (idempotent).
   *
   * Umoznuje manazerovi/klientovi rychle vyhodnotit:
   *   - Kolik novych loseru pribylo tento mesic
   *   - Kolik bylo VYRESENO (rest kampan) = uspesne zasahy
   *   - Kolik se vratilo (RE_FLAGGED) = opakovane problemy
   *   - Label application rate = efektivita operativy
   */
  function writeMonthlyReviewTab(ss, runDate) {
    var sheetName = 'MONTHLY_REVIEW';
    var sheet = getOrCreateSheet(ss, sheetName);

    var lifecycleSheet = ss.getSheetByName('LIFECYCLE_LOG');
    if (!lifecycleSheet) {
      Logger.log('WARN: MONTHLY_REVIEW — LIFECYCLE_LOG tab neexistuje, skip.');
      return;
    }

    var lastRow = lifecycleSheet.getLastRow();
    if (lastRow < 2) {
      // Prazdny log — jen header
      sheet.clearContents();
      var emptyHeaders = ['month', 'new_flags', 'repeated', 'resolved', 'un_flagged', 're_flagged', 'category_change', 'unique_products_touched', 'application_rate_pct'];
      sheet.getRange(1, 1, 1, emptyHeaders.length).setValues([emptyHeaders]);
      sheet.getRange(1, 1, 1, emptyHeaders.length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      return;
    }

    // Detekce kde zacinaji data v LIFECYCLE_LOG:
    //   - Novy layout (s dashboardem) — header na row 9, data od row 10
    //   - Stary layout (bez dashboardu) — header na row 1, data od row 2
    var r9 = String(lifecycleSheet.getRange(9, 1).getValue()).trim();
    var r1 = String(lifecycleSheet.getRange(1, 1).getValue()).trim();
    var dataStart;
    if (r9 === 'run_date') {
      dataStart = 10;
    } else if (r1 === 'run_date') {
      dataStart = 2;
    } else {
      Logger.log('WARN: MONTHLY_REVIEW — nelze detekovat LIFECYCLE_LOG layout, skip.');
      return;
    }

    if (lastRow < dataStart) {
      sheet.clearContents();
      var emptyH = ['month', 'new_flags', 'repeated', 'resolved', 'un_flagged', 're_flagged', 'category_change', 'unique_products_touched', 'application_rate_pct'];
      sheet.getRange(1, 1, 1, emptyH.length).setValues([emptyH]);
      sheet.getRange(1, 1, 1, emptyH.length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      return;
    }

    // Nacti LIFECYCLE_LOG data: run_date (col 1), item_id (col 2), transition_type (col 5)
    var data = lifecycleSheet.getRange(dataStart, 1, lastRow - dataStart + 1, 5).getValues();

    // Agregace per YYYY-MM
    var buckets = {}; // { 'YYYY-MM': { transitions: {...}, products: Set } }
    for (var i = 0; i < data.length; i++) {
      // Normalize datum (Date objekt, ISO string nebo US locale format)
      var dateStr = Utils.normalizeDate(data[i][0]);
      var itemId = String(data[i][1] || '').trim();
      var transition = String(data[i][4] || '').trim();
      // Validace: dateStr musi byt YYYY-MM-DD format po normalizaci
      if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;

      var monthKey = dateStr.substring(0, 7); // YYYY-MM
      if (!buckets[monthKey]) {
        buckets[monthKey] = {
          transitions: { NEW_FLAG: 0, REPEATED: 0, RESOLVED: 0, UN_FLAGGED: 0, RE_FLAGGED: 0, CATEGORY_CHANGE: 0 },
          products: {}
        };
      }
      if (buckets[monthKey].transitions[transition] !== undefined) {
        buckets[monthKey].transitions[transition]++;
      }
      if (itemId) {
        buckets[monthKey].products[itemId] = true;
      }
    }

    // Serad mesice DESC (nejnovejsi nahore)
    var monthKeys = Object.keys(buckets).sort(function (a, b) { return a < b ? 1 : -1; });

    // Fresh write
    sheet.clearContents();
    try { sheet.clearConditionalFormatRules(); } catch (cfe) { /* ok */ }

    var headers = [
      'month',
      'new_flags',
      'repeated',
      'resolved',
      'un_flagged',
      're_flagged',
      'category_change',
      'unique_products_touched',
      'application_rate_pct'
    ];

    var rows = [headers];
    for (var mi = 0; mi < monthKeys.length; mi++) {
      var mk = monthKeys[mi];
      var b = buckets[mk];
      var uniqueCount = Object.keys(b.products).length;

      // Application rate = RESOLVED / (flagged_v_predchozim_mesici).
      // Pro zjednoduseni: RESOLVED / (NEW_FLAG + REPEATED + RE_FLAGGED) tohoto mesice × 100.
      var totalFlaggedActivity = b.transitions.NEW_FLAG + b.transitions.REPEATED + b.transitions.RE_FLAGGED;
      var appRate = totalFlaggedActivity > 0
        ? Math.round((b.transitions.RESOLVED / totalFlaggedActivity) * 1000) / 10
        : 0;

      rows.push([
        mk,
        b.transitions.NEW_FLAG,
        b.transitions.REPEATED,
        b.transitions.RESOLVED,
        b.transitions.UN_FLAGGED,
        b.transitions.RE_FLAGGED,
        b.transitions.CATEGORY_CHANGE,
        uniqueCount,
        appRate
      ]);
    }

    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    // Force text format pro month sloupec (col 1) aby Sheets nepreformatoval
    if (rows.length > 1) {
      sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat('@');
    }

    // Header styling
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#1a73e8')
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 30);

    // Formatovani
    if (rows.length > 1) {
      // Month sloupec: bold, monospace
      sheet.getRange(2, 1, rows.length - 1, 1)
        .setFontWeight('bold')
        .setFontFamily('Roboto Mono');

      // Cislovane sloupce (2-8): center
      sheet.getRange(2, 2, rows.length - 1, 7)
        .setHorizontalAlignment('center');

      // Application rate sloupec (9): center + % suffix
      var appRateRange = sheet.getRange(2, 9, rows.length - 1, 1);
      appRateRange.setHorizontalAlignment('center')
        .setNumberFormat('0.0"%"');

      // Conditional formatting:
      //   - RESOLVED > 0 → zelene (col 4)
      //   - RE_FLAGGED > 0 → cervene (col 6)
      //   - application_rate >= 50 → zelene; < 20 → cervene; jinak zlute (col 9)
      var resolvedRange = sheet.getRange(2, 4, rows.length - 1, 1);
      var reFlaggedRange = sheet.getRange(2, 6, rows.length - 1, 1);
      var appRange = sheet.getRange(2, 9, rows.length - 1, 1);

      var rules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground('#d4edda').setFontColor('#155724')
          .setRanges([resolvedRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground('#f8d7da').setFontColor('#721c24')
          .setRanges([reFlaggedRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(50)
          .setBackground('#d4edda').setFontColor('#155724').setBold(true)
          .setRanges([appRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(20, 49.9)
          .setBackground('#fff3cd').setFontColor('#856404')
          .setRanges([appRange]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(20)
          .setBackground('#f8d7da').setFontColor('#721c24')
          .setRanges([appRange]).build()
      ];
      sheet.setConditionalFormatRules(rules);
    }

    // Column widths
    sheet.setColumnWidth(1, 100);  // month
    for (var cw = 2; cw <= 9; cw++) {
      sheet.setColumnWidth(cw, 140);
    }

    // Filter pro snadne serazeni/filtering
    try {
      var existingFilter = sheet.getFilter();
      if (existingFilter) existingFilter.remove();
      if (rows.length > 1) {
        sheet.getRange(1, 1, rows.length, headers.length).createFilter();
      }
    } catch (fe) { /* ok */ }

    Logger.log('INFO: MONTHLY_REVIEW — zapsano ' + (rows.length - 1) + ' mesicu.');
  }

  /**
   * Batch write — pro velke datasety rozdeli na chunks po 5000 radku.
   */
  function writeDataInBatches(sheet, data) {
    if (data.length === 0) {
      return;
    }
    var cols = data[0].length;
    var totalRows = data.length;

    // Check sheet size limit.
    // clearContents() maze hodnoty ale ne allocated rows — pro idempotent taby
    // je rozhodujici jen velikost noveho datasetu, ne existing sheet capacity.
    var newCells = totalRows * cols;
    if (newCells > MAX_CELLS_PER_TAB_LIMIT) {
      Logger.log('WARN: Dataset prekracuje ' + MAX_CELLS_PER_TAB_LIMIT + ' cell limit — truncating.');
      totalRows = Math.floor(MAX_CELLS_PER_TAB_LIMIT / cols);
      if (totalRows < 1) {
        throw new Error('Dataset je prilis velky — i jeden row prekracuje cell limit.');
      }
    }

    var writtenRows = 0;
    while (writtenRows < totalRows) {
      var batchSize = Math.min(BATCH_SIZE, totalRows - writtenRows);
      var batchData = data.slice(writtenRows, writtenRows + batchSize);
      // Vyrovnej delky sloupcu v batch
      normalizeRowLengths(batchData);
      sheet.getRange(writtenRows + 1, 1, batchSize, batchData[0].length).setValues(batchData);
      writtenRows += batchSize;
    }
  }

  function getOrCreateSheet(ss, sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  }

  function roundNumber(n, decimals) {
    if (n === null || n === undefined || isNaN(n) || !isFinite(n)) {
      return '';
    }
    var factor = Math.pow(10, decimals);
    return Math.round(n * factor) / factor;
  }

  /**
   * Vytvori progress bar z UNICODE blocks.
   * 0-100 pct → string s 20 znaky (█ full, ░ empty).
   */
  function makeBar(pct, width) {
    if (width === undefined) width = 20;
    if (pct === null || pct === undefined || isNaN(pct) || !isFinite(pct)) {
      return '─'.repeat(width);
    }
    pct = Math.max(0, Math.min(100, pct));
    var filled = Math.round((pct / 100) * width);
    var empty = width - filled;
    var bar = '';
    for (var i = 0; i < filled; i++) bar += '█';
    for (var j = 0; j < empty; j++) bar += '░';
    return bar;
  }

  /**
   * Email report — interni format pro specialistu.
   */
  function sendEmailReport(config, summary, effectiveness, sheetUrl) {
    if (!config.adminEmail) {
      return;
    }
    var emails = Utils.validateEmails(config.adminEmail);
    if (emails.length === 0) {
      Logger.log('WARN: Zadne validni email adresy — email se neposila.');
      return;
    }

    var subject = '[PPC Loser Detector] ' + summary.accountName + ' — ' + Utils.formatDate(summary.runDate);

    var body = 'Account: ' + summary.accountName + ' (' + summary.customerId + ')\n';
    body += 'Period: ' + Utils.formatDate(summary.lookbackStart) + ' to ' + Utils.formatDate(summary.lookbackEnd) + ' (' + config.lookbackDays + ' days)\n';
    body += 'Sheet: ' + sheetUrl + '\n\n';

    body += '=== FUNNEL ===\n';
    body += 'Raw rows:              ' + summary.funnel.rawRows + '\n';
    body += 'Brand excluded:        ' + summary.funnel.brandExcluded + '\n';
    body += 'Rest excluded:         ' + summary.funnel.restExcluded + '\n';
    body += 'Paused excluded:       ' + summary.funnel.pausedExcluded + '\n';
    body += 'Too young (<' + config.minProductAgeDays + 'd):     ' + summary.funnel.tooYoung + '\n';
    body += 'Insufficient data:     ' + summary.funnel.insufficientData + '\n';
    body += 'Data quality issues:   ' + summary.funnel.dataQualityIssues + '\n';
    body += 'Classified:            ' + summary.funnel.classified + '\n\n';

    body += '=== FLAGGED ===\n';
    body += 'LOSER_REST:    ' + summary.flags.loserRestTotal + '\n';
    body += '  zero_conv:   ' + (summary.flags.loserByTier.zero_conv || 0) + '\n';
    body += '  low_volume:  ' + (summary.flags.loserByTier.low_volume || 0) + '\n';
    body += '  mid_volume:  ' + (summary.flags.loserByTier.mid_volume || 0) + '\n';
    body += '  high_volume: ' + (summary.flags.loserByTier.high_volume || 0) + '\n';
    body += 'LOW_CTR:       ' + summary.flags.lowCtrTotal + '\n';
    body += 'Overlap:       ' + summary.flags.overlap + '\n';
    body += 'Wasted spend:  ' + roundNumber(summary.flags.totalWastedSpend, 2) + ' ' + summary.currency + ' (' + Utils.safePctFormat(summary.flags.wastedSpendPctOfTotal) + ' total cost)\n\n';

    if (effectiveness) {
      body += '=== TRANSITIONS ===\n';
      body += 'NEW_FLAG:      ' + (effectiveness.transitions.NEW_FLAG || 0) + '\n';
      body += 'REPEATED:      ' + (effectiveness.transitions.REPEATED || 0) + '\n';
      body += 'RESOLVED:      ' + (effectiveness.transitions.RESOLVED || 0) + '\n';
      body += 'RE_FLAGGED:    ' + (effectiveness.transitions.RE_FLAGGED || 0) + '\n';
      if (effectiveness.applicationRate !== null) {
        body += '\nLabel application rate: ' + Utils.safePctFormat(effectiveness.applicationRate) + '\n';
        body += '  (' + effectiveness.resolvedThisRun + ' resolved / ' + effectiveness.labeledLastRun + ' labeled last run)\n';
      }
      body += '\n';
    }

    body += '=== TOP WASTED SPEND ===\n';
    var top = summary.topLosers || [];
    for (var i = 0; i < top.length && i < 5; i++) {
      body += (i + 1) + '. ' + top[i].itemId + ' — ' + roundNumber(top[i].wastedSpend, 2) + ' ' + summary.currency + ' — ' + top[i].reasonCode + '\n';
    }
    body += '\n';

    body += '=== TOP LOW-CTR ===\n';
    var topC = summary.topLowCtr || [];
    for (var j = 0; j < topC.length && j < 5; j++) {
      body += (j + 1) + '. ' + topC[j].itemId + ' — CTR ' + roundNumber(topC[j].ctr * 100, 3) + '% — IS ' + roundNumber(topC[j].searchImpressionShare * 100, 1) + '% — ' + topC[j].reasonCode + '\n';
    }
    body += '\nAttribution: Scripts API default = Last-Click.\n';

    try {
      MailApp.sendEmail(emails.join(','), subject, body);
      Logger.log('INFO: Email odeslan na: ' + emails.join(','));
    } catch (e) {
      Logger.log('WARN: Email odeslani selhalo: ' + e.message);
    }
  }

  return {
    writeAll: writeAll,
    writeFeedUploadTab: writeFeedUploadTab,
    writeActionsTab: writeActionsTab,
    writeProductTimelineTab: writeProductTimelineTab,
    writeDetailTab: writeDetailTab,
    writeSummaryTab: writeSummaryTab,
    appendLifecycleLogTab: appendLifecycleLogTab,
    appendWeeklySnapshot: appendWeeklySnapshot,
    writeMonthlyReviewTab: writeMonthlyReviewTab,
    sendEmailReport: sendEmailReport
  };
})();

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
  Logger.log('FUNNEL (campaign filter):');
  Logger.log('  raw rows (pre-filter):   ' + summary.funnel.rawRows);
  Logger.log('  excluded brand:          ' + summary.funnel.brandExcluded);
  Logger.log('  excluded rest:           ' + summary.funnel.restExcluded);
  Logger.log('  excluded paused:         ' + summary.funnel.pausedExcluded);
  Logger.log('  kept (main enabled):     ' + ((summary.funnel.keptRows !== undefined) ? summary.funnel.keptRows : 'N/A'));
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
    ['CAMPAIGN FILTERING (regex v názvu kampaně)', '', ''],
    ['brandCampaignPattern', cfg.brandCampaignPattern, 'Brand kampaně se vyloučí z analýzy'],
    ['restCampaignPattern', cfg.restCampaignPattern, 'Rest kampaně se ignorují'],
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
