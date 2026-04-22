# Shopping/PMAX Loser Detector

**Google Ads Script pro identifikaci neefektivních produktů a trend detekci v Shopping + PMAX kampaních.**

Produkčně nasazovaný nástroj pro PPC specialisty. Analyzuje výkon per-produkt, klasifikuje do 5 kategorií, trackuje lifecycle v čase, generuje akční seznam v Google Sheetu.

---

## 🚀 Quick start (pro kolegy — nasazení na klienta)

Celý skript je v jednom souboru `combined.gs`. Stačí pastnout do Google Ads Scripts editoru.

**3 kroky:**

1. **Vytvoř kopii output sheet template:**
   👉 [https://docs.google.com/spreadsheets/d/1BPmB00tXlc7Jq5sdNrrMUXo7OFQp1eoTYaSBLTRYkTA/copy](https://docs.google.com/spreadsheets/d/1BPmB00tXlc7Jq5sdNrrMUXo7OFQp1eoTYaSBLTRYkTA/copy)

   Přejmenuj kopii podle klienta, zkopíruj ID z URL (část mezi `/d/` a `/edit`).

2. **V Google Ads účtu klienta → Tools & Settings → Bulk Actions → Scripts → +**
   Paste obsah `combined.gs` z tohoto repa ([raw link](https://raw.githubusercontent.com/marketingmatous-ai/shopping-pmax-loser-detector/main/combined.gs)).

3. **Uprav CONFIG na začátku skriptu:**
   - `outputSheetId: '<ID z kroku 1>'`
   - `targetPnoPct: 30` (cílové PNO klienta)
   - `brandCampaignPattern: '(?i)BRD'` (regex na brand kampaně klienta)

   Ulož → Spustit.

Detailní návod: [DEPLOYMENT-GUIDE.md](./DEPLOYMENT-GUIDE.md)

---

## 📋 Obsah

1. [Co skript dělá](#co-skript-dělá)
2. [Architektura](#architektura)
3. [Klasifikační pravidla](#klasifikační-pravidla)
4. [Output struktura (10 tabů)](#output-struktura-10-tabů)
5. [Transition tracking](#transition-tracking-lifecycle)
6. [Dashboard & vizualizace](#dashboard--vizualizace)
7. [Setup — 2 možnosti](#setup--2-možnosti)
8. [CONFIG parametry](#config-parametry)
9. [Logika celého skriptu](#logika-celého-skriptu)
10. [Známé limitace](#známé-limitace)
11. [Troubleshooting](#troubleshooting)

---

## Co skript dělá

### Vstup
- Google Ads účet se **Shopping** a/nebo **Performance Max** kampaněmi
- Google Sheet kam zapisovat výstupy (vytvoří skript sám, nebo kopií z template)

### Proces (každý run)
1. **Fetch dat** — 5 separátních GAQL queries (produkty, previous period, impression share, first-click dates, YoY)
2. **Klasifikace kampaní** — main / brand / rest / paused (podle regex patternů)
3. **Agregace per item_id** — každý produkt má 4 sady metrik: main / brand / rest / total
4. **Pipeline gates** — age gate, sample size gate, data quality check
5. **5 klasifikátorů** — LOSER_REST / LOW_CTR_AUDIT / RISING / DECLINING / LOST_OPPORTUNITY
6. **Merge & priority** — primární label + secondary flags
7. **Transition detection** — porovnání s LIFECYCLE_LOG (NEW / REPEATED / RESOLVED / UN_FLAGGED / RE_FLAGGED / CATEGORY_CHANGE)
8. **Effectiveness scoring** — pre/post metriky pro vyřešené produkty
9. **Zápis** do 9 tabů Google Sheetu + optional email report

### Výstup
- **FEED_UPLOAD** tab — ready-to-import CSV pro GMC/Mergado
- **ACTIONS** tab — top-priority akční seznam pro specialisty
- **SUMMARY** tab — dashboard s KPI, grafy, trendy
- **DETAIL** tab — full trace per produkt
- **LIFECYCLE_LOG** tab — historie transitions
- **PRODUCT_TIMELINE** tab — per-produkt effectiveness score
- **WEEKLY_SNAPSHOT** tab — týdenní KPI pro dlouhodobé trendy
- **CONFIG** tab — runtime hodnoty (informativní)
- **README** tab — instrukce pro uživatele sheetu

---

## Architektura

```
combined.gs (~5867 řádků, build z 7 modulů)
  ├─ _config.gs     → CONFIG objekt (editace per klient)
  ├─ utils.gs       → Helpers (formatDate, addDays, safeRegex, ...)
  ├─ config.gs      → Validace CONFIG + deprecated loadFromSheet
  ├─ data.gs        → GAQL queries + agregace
  ├─ classifier.gs  → 5 klasifikátorů + transition detection
  ├─ output.gs      → Zápis do 9 tabů + vizualizace
  └─ main.gs        → Orchestrace + setupOutputSheet()

Build: bash build-combined.sh → generuje combined.gs
```

**Jazyk:** JavaScript ES5 (Google Ads Scripts constraint).

**Proč monolit combined.gs:** Google Ads Scripts editor má jediný `.gs` soubor. Zdroj pravdy jsou modulární soubory (snadná editace), do UI se pastuje combined.

---

## Klasifikační pravidla

### 🔴 LOSER_REST — multi-tier PNO klasifikace

Produkty co utrácejí bez návratnosti. 4 tiery podle volume konverzí:

| Tier | Volume | Threshold | Reason code |
|------|--------|-----------|-------------|
| **zero_conv** | conv < 1 | spend ≥ `pnoMultiplierZeroConv` × expected_CPA (default **2.0×**) | `zero_sales_high_spend` nebo `fractional_conv_high_spend` |
| **low_volume** | 1–3 conv | actualPno ≥ `pnoMultiplierLowVol` × target_pno (default **1.5×**) | `low_volume_high_pno` |
| **mid_volume** | 4–10 conv | actualPno ≥ `pnoMultiplierMidVol` × target_pno (default **2.0×**) | `mid_volume_high_pno` |
| **high_volume** | 11+ conv | actualPno ≥ `pnoMultiplierHighVol` × target_pno (default **3.0×**) | `high_volume_extreme_pno` |

**Proč různé multiplier per tier:** High-volume produkty přispívají k objemu i když mají vyšší PNO — chceme je flaggovat jen při extrémním odklonu. Low-volume produkty jsou volatilní → přísnější threshold.

**Důležité:**
- `actualPno` = `(main_cost / main_conversion_value) × 100` — **jen main metrics**, brand data neovlivní
- `expected_CPA` = `product_price × (target_pno / 100)`
- **zero_conv VYŽADUJE skutečnou cenu produktu** (`gmc_feed` nebo `derived`). Bez ní skript SKIP (raději neklasifikovat než klasifikovat nespolehlivě)

### 🟡 LOW_CTR_AUDIT — kontextové detekce

Produkty co se zobrazují, ale málokdo na ně klikne.

| Gate | Threshold |
|------|-----------|
| Impressions | ≥ `minImpressionsLowCtr` (default 500) |
| Clicks | ≥ `minClicksLowCtr` (default 0 = bez limitu) |
| CTR | < `ctrThresholdMultiplier` × baseline CTR (default 0.7×) |

**Skip pokud rentabilní:** produkt s conv ≥ `lowCtrSkipIfProfitableMinConv` (default 3) a PNO ≤ target × 1.1 se **neflaguje** (rentabilní produkt s "nízkým" CTR je stále dobrý byznys).

**Reason code (podle Impression Share):**
- `irrelevant_keyword_match` — IS < 0.3 + nízké CTR = feed matchuje špatné dotazy
- `high_visibility_low_appeal` — IS > 0.7 + nízké CTR = produkt se zobrazuje, ale nikdo neklikne (cena/foto/stock)
- `low_ctr_general` — medium IS = manuální audit

**Baseline scope** (`ctrBaselineScope`):
- `account` (default) — průměr přes celý účet (jednodušší, méně šumu)
- `campaign` — průměr per kampaň (vhodnější pro diverzifikované účty)

### 📈 RISING — growth detection

Produkty s výrazným růstem revenue vs. předchozí období.

| Gate | Threshold |
|------|-----------|
| Growth % | ≥ `risingGrowthThreshold` (default 50%) |
| Min conversions | ≥ `minConversionsForTrendCompare` v **obou** obdobích (default 3) |

**Growth formula:** `((current_value − previous_value) / previous_value) × 100`

**Reason:** `strong_growth` (≥ 100%) nebo `growth` (50–99%)

**Používá main_metrics** (ne total), aby brand spike nezpůsobil false RISING.

**Akce:** Early scaling kandidáti — zvýšit budget / vydělit do vlastní asset group.

### 📉 DECLINING — drop detection

Produkty s poklesem revenue ≥ `decliningDropThreshold` (default 30%).

| Gate | Threshold |
|------|-----------|
| Drop % | ≥ 30% |
| Min conversions | ≥ 3 v obou obdobích |

**Reason:** `critical_decline` (≥ 50% drop) nebo `decline` (30–49%).

**Akce:** Investigace — cena vs konkurence, skladové zásoby, sezónnost, stránka produktu.

### 💎 LOST_OPPORTUNITY — rentabilní + nedostatečně zobrazované

Produkty které prodávají výborně, ale Google je málo zobrazuje.

| Gate | Threshold |
|------|-----------|
| Conversions | ≥ `lostOpportunityMinConv` (default 5) |
| PNO | ≤ target × `lostOpportunityMaxPnoMultiplier` (default 0.8× = 20% lepší než target) |
| Search Impression Share | < `lostOpportunityMaxImpressionShare` (default 0.5 = 50%) |

**Reason:** `low_is_high_roas` (high ROAS, low visibility)

**Akce:** Zvýšit bid ceiling, vydělit do dedikované kampaně, zvýšit daily budget.

### Priority při overlapu

Produkt může splňovat víc kategorií. Pořadí priorit:

```
LOSER_REST > LOW_CTR_AUDIT > DECLINING > LOST_OPPORTUNITY > RISING
```

`primaryLabel` = nejdůležitější kategorie, ostatní jdou do `secondaryFlags`.

---

## Output struktura (10 tabů)

### 1. README
Dokumentace pro uživatele sheetu — odkaz na template copy, popis tabů, instrukce.

### 2. FEED_UPLOAD
Ready-to-import CSV formát:
- `id` — product item_id
- `custom_label_N` — hodnota:
  - `loser_rest` — ztrátové produkty (všechny 4 tiery)
  - `low_ctr_audit` — nízké CTR
  - `DECLINING` — pokles tržeb
  - `RISING` — růst
  - `LOST_OPPORTUNITY` — rentabilní + málo zobrazované
  - `healthy` — produkty, které prošly revizí (status=ok, bez flagu)

Použití: Upload jako Supplemental Feed do GMC, nebo import do Mergada.

**Tipy pro nastavení v Google Ads:**

- **Rest kampaň (filter loserů):** `custom_label_N = loser_rest` → jen losery jdou do rest
- **Alternativně exclude healthy z rest:** `custom_label_N != healthy` → zdravé produkty zůstávají v main
- **Main kampaň scaling:** filtruj `custom_label_N = RISING` pro vlastní asset group / listing group s vyšším budgetem
- **Audit kampaň:** `custom_label_N = low_ctr_audit` nebo `DECLINING` → ruční review produktového feedu

**Opt-out healthy label:** Nastav `CONFIG.labelHealthyValue = ''` (prázdný string) → healthy produkty se nezapisují, feed obsahuje jen flagged (chování před touto verzí).

### 3. DETAIL
Full trace per produkt (41 sloupců):
- Identifikace: item_id, primary_label, secondary_flags, reason_code, tier, status, transition_type
- Titulky & kampaně: product_title, product_brand, product_type, campaign_name, **campaigns_count**, **primary_campaign_share_pct**, **top_campaigns**
- Věk: first_click_date, age_days
- Metriky: clicks, impressions, cost, conversions, conversion_value, actual_PNO_pct, actual_ROAS, CTR_pct, search_impression_share
- Cena: product_price, price_source (gmc_feed / derived / unavailable)
- Gate trace: expected_CPA, expected_conversions, min_clicks_required, passed_age_gate, passed_sample_gate
- Signály: yoy_signal, wasted_spend, wasted_spend_pct, note

**Podmíněné formátování:**
- actual_PNO_pct > 60 → červené pozadí
- actual_PNO_pct 30–60 → oranžové
- actual_ROAS ≥ 5 → zelené
- price_source = unavailable → žluté (warning)
- price_source = gmc_feed → zelený text

**Frozen 2 sloupce** (item_id + primary_label) + filtry na všech sloupcích.

### 4. SUMMARY
**Hlavní dashboard.** Obsahuje:

- **Hlavička** — account name, customer ID, lookback period
- **Legenda barev** + **Quick nav** (klikatelné hyperlinky na sekce)
- **4 KPI karty:** CELKOVÉ NÁKLADY / NÁKLADY NA OZNAČENÉ / MARNÝ SPEND / POTENCIÁLNÍ ÚSPORA
- **KLASIFIKAČNÍ TRYCHTÝŘ** — funnel od raw rows po klasifikované (s ├─ označené / └─ zdravé bullety)
- **Informační insight box** — vysvětlení proč X% produktů nemá dost dat
- **ROZDĚLENÍ OZNAČENÝCH PRODUKTŮ** + **pie chart** per tier
- **DOPAD** — náklady vs revenue breakdown + **bar chart**
- **ZÁKLADNÍ METRIKY ÚČTU** — total cost, clicks, conv, CVR, CTR, ROAS, PNO, AOV
- **ZMĚNY V TOMTO BĚHU** — transitions count + **Míra aplikace labelů** (barevně podle hodnoty)
- **TOP 10 ZTRÁTOVÝCH** + **scatter plot** (cost × marný spend)
- **TOP 10 NÍZKÉHO CTR**
- **TÝDENNÍ TRENDY** — tabulka + **line chart** (ROAS, PNO %, Marný spend)
- **VÝVOJ OZNAČENÍ V ČASE** — **stacked area chart** (řešíme nebo přibývá?)
- **VÝKON PODLE KAMPANĚ** — tabulka + **column chart** (kde peníze utíkají)
- **ÚČINNOST ZÁSAHŮ** — +/=/− breakdown přes historii
- **ZDRAVÍ REST KAMPANÍ** — efficient / acceptable / wasteful
- **VHLEDY Z BRAND KAMPANÍ** — brand-only sellers
- **NOVÉ PRODUKTY** (< 30 dní) + **rising star candidates** tabulka
- **POUŽITÁ KONFIGURACE** — config audit trail

### 5. LIFECYCLE_LOG
**Dashboard nahoře + append-only data dole.** Každý řádek dat = jeden run × produkt × stav transition.

**Dashboard (rows 1-8):**
- Row 1: Title banner
- Row 2: Kumulativní info (`N transitions · M unikátních produktů · K runs`)
- Rows 4-5: **6 KPI karet** (NEW_FLAG / REPEATED / RESOLVED / UN_FLAGGED / RE_FLAGGED / CAT_CHANGE)
- Row 7: Tento měsíc vs předchozí měsíc (nové, vyřešené, vrácené)

**Data (row 10+):**
Sloupce: run_date, item_id, current_label, previous_label, transition_type, current_campaign, previous_campaign, campaign_moved, cost_30d, conversions_30d, pno_30d_pct, roas_30d, ctr_30d_pct, reason_code, tier, runs_since_first_flag, notes

**Dedup:** per `(run_date, item_id, transition_type)` — při opakovaném spuštění v jednom dni se duplicity přeskakují. Logger vypíše count `skippedDup`.

**Jen transitions != NO_CHANGE** se ukládají (tab nebobtnává).

### 5b. MONTHLY_REVIEW
**Měsíční agregace z LIFECYCLE_LOG** — 1 řádek per měsíc pro rychlý management overview.

Sloupce: month | new_flags | repeated | resolved | un_flagged | re_flagged | category_change | unique_products_touched | **application_rate_pct**

**Fresh rebuild při každém runu** (idempotentní z historie). Seřazeno DESC (nejnovější měsíc nahoře).

**Conditional formatting:**
- `resolved > 0` → zelené pozadí (úspěšné zásahy)
- `re_flagged > 0` → červené (vrátilo se po vyřešení)
- `application_rate_pct`: ≥50% zeleně · 20-49% žlutě · <20% červeně

**Formule application rate:** `RESOLVED / (NEW_FLAG + REPEATED + RE_FLAGGED) × 100` — zjednodušený ukazatel "kolik % flagovaných bylo úspěšně vyřešeno".

**Use case:** Začátek každého měsíce → otevři MONTHLY_REVIEW → zkontroluj application_rate předchozího měsíce. Nízké číslo = labely se neaplikují v GMC/Mergado.

### 6. ACTIONS
**Akční seznam** pro PPC specialistu. Nahoře má **souhrnný panel**:

- **4 KPI karty:** CELKEM OZNAČENO / MARNÝ SPEND / BEZ MANUÁLNÍ AKCE / WARNING ≥ 2 běhy
- **ROZDĚLENÍ PODLE KATEGORIE**
- **💡 DOPORUČENÍ A UPOZORNĚNÍ** — dynamické insights (REPEATED_WARNING, RE_FLAGGED, missing action_taken, atd.)

Pod panelem tabulka všech flagged produktů:
- priority_rank, category, item_id, product_title, product_price, current_campaign, tier, reason_code
- Main metriky: main_clicks, main_impressions, main_cost, main_conv, main_pno_pct, main_ctr_pct
- Total metriky: total_clicks, total_cost, total_conv, total_roas, brand_share_pct
- growth_pct, wasted_spend, recommended_action
- days_since_first_flag, transition_status, secondary_flags
- **Manuální sloupce** (preserve mezi runy): action_taken, action_date, consultant_note

Barevné pozadí per kategorie (loser_rest = červená, LOW_CTR = žlutá, DECLINING = oranžová, RISING = zelená, LOST_OPP = modrá).

### 7. PRODUCT_TIMELINE
Per-produkt (upsert — jeden řádek per item_id) historie:
- Current status, first_flag_date, last_flag_date
- kpi_before snapshot (když poprvé flagged)
- Current kpi vs before delta
- **Effectiveness score** (+ / = / − / PENDING / N/A)
- Days since action, days in current label
- Manual columns (preserve)

Effectiveness score logic:
- `+` — delta_cost < 0 AND (delta_roas > 0 OR resolved_to_rest)
- `=` — smíšený výsledek
- `−` — intervence škodí (cost roste, ROAS padá)
- `PENDING` — méně než `effectivenessMinDaysSinceAction` (default 14) od zásahu
- `N/A` — nemáme dost dat pro vyhodnocení

### 8. WEEKLY_SNAPSHOT
Append-only (1 řádek per týden). Ukládá:
- account_cost_total, account_clicks, account_conversions, account_conv_value
- account_roas, account_pno_pct, account_ctr_pct
- flagged_count_total, flagged_loser_rest, flagged_low_ctr, flagged_declining, flagged_rising, flagged_lost_opp
- wasted_spend_total
- resolved_this_run, re_flagged_this_run
- label_application_rate_pct

**Použití:** Dlouhodobý tracking pro DASHBOARD trendy (line chart + stacked area v SUMMARY).

### 9. CONFIG
**Informativní snapshot** runtime hodnot CONFIG. Barevně kódovaný:
- Boolean hodnoty (zelená TRUE / červená FALSE)
- Numerické hodnoty (modré)
- Sekcní headery (modré pozadí)
- dryRun = TRUE → žlutý warning

**Nelze editovat** — zdroj pravdy je CONFIG objekt v kódu.

---

## Transition tracking (lifecycle)

Každý run skript porovnává aktuální klasifikaci s předchozí (z LIFECYCLE_LOG) a přiřadí jeden z 7 transition states:

| Transition | Podmínka | Význam |
|------------|----------|--------|
| **NEW_FLAG** 🆕 | bez prev + current má label | Poprvé označený |
| **REPEATED** 🔁 | current label = prev label + neměnil kampaň | Stále označený — label pravděpodobně nebyl aplikován |
| **REPEATED_WARNING** ⚠️ | REPEATED + runs ≥ 2 | **Kritické** — 2+ běhy bez akce |
| **RESOLVED** ✅ | prev měl label, current bez, **+ teď v rest kampani** | Úspěšný zásah (klient aplikoval label) |
| **UN_FLAGGED** 🌱 | prev měl label, current bez, zůstal v main | Produkt se zlepšil sám (organicky, bez zásahu) |
| **RE_FLAGGED** ⚠️ | prev bez, current má label | Vrátil se po vyřešení — neúspěšný zásah |
| **CATEGORY_CHANGE** 🔄 | prev label ≠ current label | Změnil kategorii (např. LOSER → LOW_CTR) |

**Label application rate:** `(RESOLVED count) / (labeled last run)` — měří jak efektivně se aplikují doporučení.

---

## Dashboard & vizualizace

### SUMMARY tab obsahuje 6 různých grafů:

1. **Pie chart** — Rozdělení označených podle kategorie
2. **Bar chart** — Rozpad nákladů (Označené vs Ostatní)
3. **Line chart** — Týdenní trendy ROAS/PNO/Marný spend
4. **Stacked area** — Vývoj označení v čase per kategorie
5. **Column chart** — Marný spend podle kampaně
6. **Scatter plot** — Top 10 ztrátových (cost × wasted)

### Interaktivita:
- **Filtry** na všech datových tabech (DETAIL, ACTIONS, LIFECYCLE_LOG, PRODUCT_TIMELINE, WEEKLY_SNAPSHOT)
- **Klikatelná Quick nav** pod hlavičkou SUMMARY (skok na sekce)
- **Frozen rows/columns** na klíčových tabech
- **Podmíněné formátování** (PNO/ROAS/IS/price_source)

---

## Setup — 2 možnosti

### A) AUTO-SETUP (první deploy)
1. Paste `combined.gs` do Google Ads Scripts editoru
2. Nech `CONFIG.outputSheetId: ''`
3. V dropdownu "Vybrat funkci" zvol `setupOutputSheet` → **Spustit**
4. Z logu zkopíruj ID nového sheetu → vlož do `CONFIG.outputSheetId`
5. Uprav CONFIG parametry podle klienta
6. Pust `main()` s `dryRun=true` → ověř log
7. Pust s `dryRun=false` → ostré nasazení

### B) KOPIE TEMPLATE (rychlé pro další klienty)
1. Klikni na copy URL v komentáři skriptu (řádek ~57):
   ```
   https://docs.google.com/spreadsheets/d/1BPmB00tXlc7Jq5sdNrrMUXo7OFQp1eoTYaSBLTRYkTA/copy
   ```
2. Google vytvoří kopii template v tvém Drive
3. Přejmenuj podle klienta
4. Zkopíruj ID kopie z URL → vlož do `CONFIG.outputSheetId`
5. Uprav CONFIG + pust

### Scheduling
Po prvním úspěšném runu v Google Ads Scripts UI nastav **Schedule → Weekly**. Doporučený den: pondělí ráno (data za předchozí týden kompletní).

---

## CONFIG parametry

### Základní
```javascript
targetPnoPct: 30           // Cílové PNO v % (nákladovost)
lookbackDays: 60           // Okno analýzy (doporučeno 30-90)
outputSheetId: '...'       // Google Sheet ID (povinné po prvním setupu)
adminEmail: ''             // Email pro notifikace (volitelné)
```

### Labels (do custom_label v GMC/Mergado)
```javascript
customLabelIndex: 2        // Číslo labelu 0-4
labelLoserRestValue: 'loser_rest'
labelLowCtrValue: 'low_ctr_audit'
labelHealthyValue: 'healthy'   // '' = nezapisovat healthy label
```

### Item_id case (pro match s GMC supplemental feed)
```javascript
itemIdCaseOverride: 'auto'
// 'auto'     — detekuj dominant case z shopping_product (DEFAULT, doporučeno)
// 'upper'    — vynutit UPPERCASE (pozor: rozbije mixed-case jako "Print")
// 'lower'    — vynutit lowercase
// 'preserve' — nic neupravovat (= lowercase z Google Ads)
```
Skript automaticky mapuje item_id ze `shopping_product` resource (= přesně jak v GMC, včetně mixed-case). Produkty co nejsou v aktuálním GMC feedu (stažené/disapproved) se **skipnou** z FEED_UPLOAD, aby nedošlo k "Nabídka neexistuje" chybě. Viz [DEPLOYMENT-GUIDE.md — Krok 8.4b](./DEPLOYMENT-GUIDE.md#84b-jak-skript-řeší-item_id-case-uppercase-vs-lowercase).

### Campaign filtering
```javascript
brandCampaignPattern: '(?i)BRD'      // Regex pro brand kampaně
restCampaignPattern: '(?i)REST'      // Regex pro rest kampaně
analyzeChannels: ['SHOPPING', 'PERFORMANCE_MAX']
```

### Sample size gate
```javascript
minClicksAbsolute: 30           // Minimum kliků pro klasifikaci
minExpectedConvFloor: 1         // Min expected conv (clicks × account_CVR)
```

### Rising star protection
```javascript
minProductAgeDays: 30           // Produkty < N dní se neevaluují
```

### LOSER tiers
```javascript
tierLowVolumeMax: 3             // 1-3 conv = low volume
tierMidVolumeMax: 10            // 4-10 conv = mid volume
pnoMultiplierZeroConv: 2.0      // 0 conv → spend ≥ 2× expected CPA
pnoMultiplierLowVol: 1.5        // 1-3 conv → PNO ≥ 1.5× target
pnoMultiplierMidVol: 2.0        // 4-10 conv → PNO ≥ 2.0× target
pnoMultiplierHighVol: 3.0       // 11+ conv → PNO ≥ 3.0× target
```

### LOW CTR detekce
```javascript
ctrBaselineScope: 'account'     // 'account' nebo 'campaign'
minImpressionsLowCtr: 500       // Min impressions
minClicksLowCtr: 0              // Min kliků (0 = bez limitu)
ctrThresholdMultiplier: 0.7     // CTR < 0.7× baseline = flag
lowCtrSkipIfProfitableMinConv: 3 // Neflag rentabilních (conv ≥ N + PNO ≤ target×1.1)
```

### Trend detection
```javascript
risingGrowthThreshold: 50                // Growth ≥ 50% = RISING
decliningDropThreshold: 30               // Drop ≥ 30% = DECLINING
minConversionsForTrendCompare: 3         // Min conv obou periodách
```

### LOST_OPPORTUNITY
```javascript
lostOpportunityMinConv: 5                 // Min conv pro rentability claim
lostOpportunityMaxPnoMultiplier: 0.8      // PNO ≤ target × 0.8 (výrazně rentabilní)
lostOpportunityMaxImpressionShare: 0.5    // IS < 0.5 (málo zobrazuje)
```

### Effectiveness
```javascript
effectivenessMinDaysSinceAction: 14       // Dní před vyhodnocením zásahu
restCampaignEfficientThreshold: 0.2       // Rest cost ≤ 0.2× before = efficient
```

### Ostatní
```javascript
enableYoYSeasonalityCheck: true           // YoY porovnání (pokud >1 rok dat)
enableHistoryDedup: true                  // Tracking transitions
historyDedupDays: 14                      // Same-day dedup okno
dryRun: false                             // TRUE = jen loguje
groupByParentId: false                    // TRUE = agreguj varianty
includeProductTitles: true                // Zapisovat product_title
maxRowsDetailTab: 50000                   // Sheet size guard
```

---

## Logika celého skriptu

### Data scope oddělení

Skript důsledně rozlišuje 4 "buckets" metrik per produkt:

| Bucket | Zdroj | Používá se pro |
|--------|-------|----------------|
| `main_metrics` | Non-brand, non-rest, non-paused | **Všechny klasifikátory** (LOSER, LOW_CTR, RISING, DECLINING, LOST_OPP) |
| `brand_metrics` | Brand kampaně (match pattern) | Brand insights (brand-only sellers) |
| `rest_metrics` | Rest kampaně (match pattern) | Rest campaign health, effectiveness pre/post |
| `total_metrics` | Všechny dohromady | Dashboard metriky (celkový pohled), informační |

**Rationale:** Brand a rest data by zkreslila klasifikaci. Brand spike není "skutečný růst produktu", rest data jsou již "vyřešené" produkty.

### Pipeline per produkt (classifyProduct)

```
┌──────────────────────────────────────────────────────┐
│ 1. Initialize result (default status='ok', price='') │
├──────────────────────────────────────────────────────┤
│ 2. Price resolution                                  │
│    → gmc_feed / derived / unavailable                │
├──────────────────────────────────────────────────────┤
│ 3. YoY signal                                        │
│    → stable_yoy / possibly_seasonal_decline / ...    │
├──────────────────────────────────────────────────────┤
│ 4. Age gate (< 30 dní?)                              │
│    → FAIL: status=NEW_PRODUCT_RAMP_UP, RETURN        │
├──────────────────────────────────────────────────────┤
│ 5. Sample size gate                                  │
│    → FAIL: status=INSUFFICIENT_DATA, RETURN          │
├──────────────────────────────────────────────────────┤
│ 6. Data quality check (conv > 0 ale value ≤ 0)       │
│    → FAIL: status=DATA_QUALITY_ISSUE, RETURN         │
├──────────────────────────────────────────────────────┤
│ 7. 5× klasifikátory (všechny na main_metrics)        │
│    ├─ classifyLoser            → 4 tiers             │
│    ├─ classifyLowCtr           → 3 reasons           │
│    ├─ classifyRising                                 │
│    ├─ classifyDeclining                              │
│    └─ classifyLostOpportunity                        │
├──────────────────────────────────────────────────────┤
│ 8. mergeResults                                      │
│    → priority: LOSER > LOW_CTR > DECLINING > ... > RISING │
├──────────────────────────────────────────────────────┤
│ 9. applyTransition (z LIFECYCLE_LOG prev)            │
│    → NEW/REPEATED/RESOLVED/UN_FLAGGED/RE_FLAGGED/... │
└──────────────────────────────────────────────────────┘
```

### Funnel invariant

```
raw_rows
  − brand_excluded
  − rest_excluded
  − paused_excluded
  = kept_rows
    → aggregate per item_id
    = afterAggregation (unique products)
      − tooYoung
      − insufficientData
      − dataQualityIssues
      = classified
        = flagged + healthy
```

**Invariant check** v logu: `young + insufficient + dq + classified = afterAggregation`. Pokud ne → bug v pipeline.

### Idempotence

Skript lze spustit vícekrát bez ztráty dat:
- **FEED_UPLOAD, DETAIL, SUMMARY** — clearContents + fresh write (overwrite)
- **LIFECYCLE_LOG** — append-only + same-day dedup (neduplikujem při retry)
- **PRODUCT_TIMELINE** — upsert per item_id (merge manual columns)
- **ACTIONS** — clearContents + preserve manual columns (action_taken / action_date / consultant_note)
- **WEEKLY_SNAPSHOT** — append 1 řádek per run

---

## Známé limitace

### 1. Impression Share aggregation = MAX (ne weighted average)

**Problém:** GAQL API nepovoluje kombinovat `metrics.search_impression_share` s `metrics.impressions` v jednom SELECTu. Proto nemůžeme spočítat impression-weighted average.

**Řešení:** Skript používá **MAX(IS)** per produkt (konzervativní). Rationale: pokud produkt má v jedné kampani vysoký IS, není "lost opportunity" i když v jiné kampani má nízký.

**Dopad:** Méně false-positive LOST_OPPORTUNITY flagů. Může ale schovat edge case produktů co mají high IS v jedné velké kampani, nízký ve všech ostatních.

### 2. Attribution — Last-Click default

Skript používá default attribution model Google Ads Scripts API = **Last-Click**. Některé hodnoty se mohou lehce lišit od Google Ads UI při jiném attribution modelu (Data-Driven, Linear, atd.).

### 3. YoY signal vyžaduje > 1 rok dat

Pokud účet existuje méně než 365 dní, YoY query selže → `yoy_signal: no_yoy_data` pro všechny produkty.

### 4. Fractional konverze (< 1)

Attribution model někdy přiřadí produktu 0.14 nebo 0.5 konverze (split přes touchpointy). Skript to hodnotí jako **zero_conv tier** (stejný threshold jako 0 conv). Reason code: `fractional_conv_high_spend`.

### 5. Shopping longtail

Typicky 80–90% produktů v e-commerce Shopping kampaních nemá dost dat pro klasifikaci (méně než threshold kliků). Skript to reportuje v SUMMARY jako "insufficient data". **Není to bug** — je to realita distribuce trafficu.

### 6. Template update při změně layoutu

Pokud se změní struktura skriptu (nové taby, sloupce, sekce), template sheet (`1BPmB00tXlc7...`) musí být ručně updatnut — spustit `setupOutputSheet()` na prázdném účtu znovu a nahradit ID v komentáři.

---

## Troubleshooting

### Error: "CONFIG validation failed: lookbackDays musi byt cislo mezi 7 a 365"
**Příčina:** CONFIG má `lookbackDays` mimo rozsah 7-365.  
**Oprava:** Nastav `lookbackDays: 30` (nebo jinou hodnotu v rozsahu).

### Error: "Account nema zadna Shopping/PMAX data v lookback period"
**Příčina:** Účet nemá aktivní Shopping/PMAX kampaně, nebo všechny jsou paused.  
**Oprava:** Zkontroluj že jsou aktivní kampaně typu SHOPPING nebo PERFORMANCE_MAX. Pokud ne, skript nelze použít.

### Sheet ukazuje "flagged = 0" i když produkty jsou v kampaních
**Příčina:** Pravděpodobně **všechny produkty** padly na **sample gate** (< 30 kliků za lookback).  
**Oprava:** 
- Zvaž delší `lookbackDays` (60 → 90)
- Zvaž nižší `minClicksAbsolute` (30 → 15), ale roste false-positive rate

### "WARN: YoY query selhal"
**Příčina:** Účet existuje < 365 dní nebo nemá data před rokem.  
**Oprava:** Není nutná — skript pokračuje bez YoY signálu (`yoy_signal: no_yoy_data`).

### CONFIG tab v sheetu ukazuje jiné hodnoty než v kódu
**Příčina:** Před 2026-04-21 skript četl CONFIG ze sheetu (`Config.loadFromSheet`). Tento bug byl opraven — zdroj pravdy je kód.  
**Oprava:** Refresh sheet po novém runu, CONFIG tab se přepíše aktuálními hodnotami.

### Grafy se překrývají s tabulkami
**Příčina:** Stará verze skriptu umísťovala grafy přes datovou oblast.  
**Oprava:** Aktuální verze (md5 `a0550272+`) umísťuje všechny grafy do sloupce I+ (mimo data). Pokud vidíš staré chování, paste aktuální combined.gs.

### TOP 10 LOW_CTR ukazuje "1 CZK" místo CTR %
**Příčina:** Google Sheets dědí setNumberFormat mezi buňkami. Fix v md5 `020d827d+`.  
**Oprava:** Paste aktuální combined.gs.

### RESOLVED count je vyšší než skutečné zásahy specialisty
**Příčina:** Před fix v md5 `78d295a5` skript nerozlišoval mezi RESOLVED (přesun do rest) a UN_FLAGGED (organické zlepšení).  
**Oprava:** Paste aktuální combined.gs. Po opravě uvidíš RESOLVED (skutečný zásah) vs UN_FLAGGED (sám se zlepšil).

---

## Reference

- **Repozitář:** https://github.com/marketingmatous-ai/shopping-pmax-loser-detector
- **Plán/spec:** `docs/` (v tomto repu)
- **Issues / bugreporty:** https://github.com/marketingmatous-ai/shopping-pmax-loser-detector/issues

### Lokální vývoj (pokud chceš skript dál upravovat)

Naklonuj repo a pracuj se zdrojovými moduly:

```bash
git clone https://github.com/marketingmatous-ai/shopping-pmax-loser-detector.git
cd shopping-pmax-loser-detector
```

**Build combined.gs** (zkompiluje 7 modulů do jednoho souboru pro GAS editor):

```bash
bash build-combined.sh
```

**Syntax check** (před commitem / deployem):

```bash
cp combined.gs /tmp/c.js && node --check /tmp/c.js && rm /tmp/c.js
```

**Copy do clipboardu** (macOS):

```bash
pbcopy < combined.gs
```

### Přispívání

Pull requesty vítány. Před odesláním:

1. Úpravy prováděj v modulárních `.gs` souborech (ne v `combined.gs`)
2. Spusť `bash build-combined.sh`
3. Ověř `node --check` na vygenerovaném `combined.gs`
4. Napiš stručný popis změny do commit message
