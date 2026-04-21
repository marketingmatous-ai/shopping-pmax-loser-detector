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
