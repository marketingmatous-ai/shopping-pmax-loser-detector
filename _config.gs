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

  // === CAMPAIGN BUCKET SEPARATION ===
  // Kampane se rozdeluji do 3 bucketu (main / brand / rest) podle regex matchingu
  // na nazev kampane. Klasifikace (LOSER / LOW_CTR / RISING / DECLINING / LOST_OPP)
  // se pocita JEN z main_metrics, aby brand spike / rest spend nezkreslily thresholdy.
  // Brand a rest data jsou ale zachovana v bucketech (brand_metrics, rest_metrics)
  // pro dashboard insights, effectiveness tracking a lifecycle transitions (RESOLVED).
  brandCampaignPattern:   '(?i)BRD',        // Regex na brand kampane (uprav podle naming konvence)
                                            // Napr. '(?i)BRD' / '(?i)BRA' / '(?i)BRAND'
                                            // '' (prazdny) nebo '^$' = zadny brand bucket
  restCampaignPattern:    '(?i)REST',       // Regex na rest kampane (kam presouvame losery)
                                            // '^$' = zadna rest struktura (zadny RESOLVED tracking)
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
