# Shopping/PMAX Loser Detector — v2 Redesign Design Doc

**Datum:** 2026-04-21
**Stav:** Approved (user schválil všech 5 sekcí)
**Účel:** Evoluce z "labeling tool" na "kompletní produktový insights + kampaň optimization tool"

---

## Context

Aktuální skript (v1) úspěšně detekuje LOSER_REST (4 tiers) a LOW_CTR_AUDIT. Konzultant potvrdil:

- Skript správně identifikuje produkty plýtvající budget
- Logika multi-tier PNO funguje (chrání high-volume produkty)
- Product price z GMC feedu je správně integrován
- Rentability gate u LOW_CTR šetří rentabilní produkty

Nicméně konzultant chce rozšířit o:

1. **Pozitivní insights** (nejen negative flagging) — kandidáti na scaling
2. **Trend detection** (momentum-based early warning / opportunity)
3. **Kompletní produktový view** — napříč všemi kampaněmi (main + brand + rest)
4. **Per-produkt historii** — tracking intervence a její účinnosti
5. **Manual input** — specialista zapíše akci, skript měří pre/post dopad
6. **DASHBOARD s trendy** — WoW/MoM vývoj account KPI

## Success criteria

- Konzultant v jednom sheetu vidí: co vyloučit z main, co scalovat, jak intervence pomáhají
- Skript rozlišuje purpose (optimalizace kampaní vs. produktové insights) a používá správný scope dat
- Manual poznámky specialisty se zachovávají mezi runs
- Effectiveness measurement funguje automatiky (pre/post delta KPI)
- Klient může dostat proaktivní doporučení ("tyto produkty prodávají jen přes brand — musíme promovat")

---

## Rozhodnutí po brainstormingu (schválené uživatelem)

### 1. Nové kategorie (schváleno)

| Kategorie | Scope | Kritérium |
|---|---|---|
| **RISING** 🆕 | TOTAL (main + brand + rest) | Revenue growth ≥ 50% vs previous period, min 3 conv obě periody |
| **DECLINING** 🆕 | TOTAL | Revenue drop ≥ 30% vs previous, min 3 conv obě periody |
| **LOST_OPPORTUNITY** 🆕 | TOTAL pro rentabilitu, MAIN pro IS | conv≥5 + PNO≤target×0.8 + IS<0.5 |

RISING/DECLINING thresholds **configurable v CONFIG** (default medium: 50% / 30%).

HERO kategorie **NE** (user odmítl — top produkty jsou viditelné v UI).

### 2. Split scope per účel (schváleno)

Zásadní rozhodnutí: **brand kampaně zahrnujeme pro produktové insights, ale NE pro flagování**.

| Účel | Scope | Kategorie |
|---|---|---|
| **Optimalizace main kampaní** | JEN MAIN (non-brand, non-rest) | LOSER_REST, LOW_CTR_AUDIT |
| **Produktové insights** | ALL (main + brand + rest) | RISING, DECLINING, LOST_OPPORTUNITY |
| **Effectiveness tracking** | ALL | PRODUCT_TIMELINE, effectiveness score |
| **Klasifikace campaign health** | split per type | REST campaign health insight |

**Rationale:** Konverze je konverze. Produkt prodávající jen přes brand je rentabilní z business pohledu, ale v main plýtvá → flagujeme pro main, ale tracking ho nesmí ignorovat.

### 3. Sheet structure (schváleno)

**9 tabů celkem:**

1. **DASHBOARD** — executive view + trendy + insights
2. **FEED_UPLOAD** — ready-to-upload CSV (unchanged)
3. **ACTIONS** 🆕 — všechny flagged/insight produkty v jednom tabu, filterable
4. **PRODUCT_TIMELINE** 🆕 — per-produkt historie, pre/post KPI, effectiveness score
5. **DETAIL** — full raw data (technická hloubka)
6. **LIFECYCLE_LOG** — append-only event log (transitions)
7. **WEEKLY_SNAPSHOT** 🆕 — weekly KPI aggregate pro trendy
8. **CONFIG** — parametry (editovatelné)
9. **README** — dokumentace + tuning tipy

### 4. Tracking (schváleno: semi-manuální)

- **Auto:** transitions, KPI deltas, effectiveness score, categories history
- **Manual sloupce v ACTIONS + PRODUCT_TIMELINE:** `action_taken`, `action_date`, `consultant_note`
- **Preserve mechanismus:** skript čte existing data před overwrite, merguje manual columns do nových řádků

---

## Architektura

### Data flow

```
┌─────────────────────────────────────────────────────────────┐
│ main()                                                       │
├─────────────────────────────────────────────────────────────┤
│  1. Load CONFIG (z CONFIG tabu, fallback na defaults)       │
│  2. Validate CONFIG                                          │
│  3. Auto-setup sheet pokud chybí outputSheetId              │
├─────────────────────────────────────────────────────────────┤
│  4. Fetch account baseline (30d default) — ALL campaigns    │
│  5. Fetch products CURRENT period — ALL campaigns            │
│  6. Fetch products PREVIOUS period 🆕 (pre-lookback)        │
│  7. Fetch product prices (shopping_product resource)        │
│  8. Fetch first_click_dates                                  │
│  9. Fetch YoY stats (volitelné)                             │
│  10. Fetch search_impression_share 🆕 (pro LOST_OPP)        │
│  11. Fetch existing ACTIONS 🆕 (preserve manual columns)    │
│  12. Fetch existing PRODUCT_TIMELINE 🆕                     │
│  13. Fetch existing LIFECYCLE_LOG                            │
│  14. Fetch existing WEEKLY_SNAPSHOT 🆕                      │
├─────────────────────────────────────────────────────────────┤
│  15. Per produkt: compute split metrics                      │
│      - main_metrics (non-brand, non-rest)                    │
│      - brand_metrics                                         │
│      - rest_metrics                                          │
│      - total_metrics (main + brand + rest)                   │
│  16. Per produkt classify:                                   │
│      a) Age gate (rising star) — na main data               │
│      b) Sample gate — na main data                           │
│      c) LOSER_REST (4 tiers) — main_metrics                 │
│      d) LOW_CTR_AUDIT — main_metrics                         │
│      e) RISING 🆕 — total, current vs previous              │
│      f) DECLINING 🆕 — total                                │
│      g) LOST_OPPORTUNITY 🆕 — total + main IS               │
│      h) Merge priority (LOSER > LOW_CTR > DECLINING > ...)  │
│      i) Detect transition                                    │
│  17. Compute effectiveness 🆕 per produkt (pre/post delta)  │
├─────────────────────────────────────────────────────────────┤
│  18. Write DASHBOARD                                         │
│  19. Write FEED_UPLOAD                                       │
│  20. Write ACTIONS 🆕 (preserve manual columns via merge)   │
│  21. Write PRODUCT_TIMELINE 🆕 (preserve manual, upsert)    │
│  22. Write DETAIL                                            │
│  23. Append LIFECYCLE_LOG (jen transitions)                 │
│  24. Append WEEKLY_SNAPSHOT 🆕                              │
│  25. Send email report                                       │
└─────────────────────────────────────────────────────────────┘
```

### Period-over-period fetching

- **Current period**: `today - lookbackDays` → `today`
- **Previous period**: `today - 2×lookbackDays` → `today - lookbackDays`
- **Implementace:** 2 separátní GAQL queries (jednodušší agregace než jedna delší query)

### Split metrics per product

```javascript
product = {
    itemId: 'nb 2414 ko',
    title: '...',
    price: 1790,
    main_metrics:   { clicks, cost, conv, conv_value, imp, ctr, pno, roas },
    brand_metrics:  { ... },
    rest_metrics:   { ... },
    total_metrics:  { clicks, cost, conv, conv_value, imp, ctr, pno, roas },
    // Previous period (pre-lookback)
    main_metrics_previous:   { ... },
    total_metrics_previous:  { ... },
    // Extra
    search_impression_share: 0.35,  // z shopping_product_view, per main
    first_click_date: '2025-08-15'
}
```

---

## Klasifikační logika

### LOSER_REST (unchanged, jen data scope změněn)

Používá **`main_metrics`**. Logika 4 tiers zachována.

### LOW_CTR_AUDIT (unchanged, jen data scope)

Používá **`main_metrics`**. Rentability gate, data sufficiency — zachováno.

### RISING 🆕

```
Input: product.total_metrics, product.total_metrics_previous

IF total_metrics.conversions < minConvForTrendCompare (3):
    skip (insufficient current data)
IF total_metrics_previous.conversions < minConvForTrendCompare:
    skip (insufficient previous data)
IF total_metrics_previous.conv_value == 0:
    skip (baseline je 0)

growth_pct = (current_rev - previous_rev) / previous_rev × 100

IF growth_pct >= risingGrowthThreshold (default 50%):
    category = 'RISING'
    reason_code = 'strong_growth' (if >100%) or 'growth' (50-100%)
    action = 'early_scaling_candidate'
```

### DECLINING 🆕

```
Stejný gate jako RISING, ale:

IF growth_pct <= -decliningDropThreshold (default -30%):
    category = 'DECLINING'
    reason_code = 'critical_decline' (if <-50%) or 'decline' (-30% až -50%)
    action = 'investigate_decline'
```

### LOST_OPPORTUNITY 🆕

```
Input: product.total_metrics, product.search_impression_share

IF total_metrics.conversions < lostOpportunityMinConv (5):
    skip (insufficient conv for rentability claim)
IF total_metrics.pno > targetPnoPct × lostOpportunityMaxPnoMultiplier (0.8):
    skip (není výrazně rentabilní)
IF product.search_impression_share > lostOpportunityMaxImpressionShare (0.5):
    skip (Google už dost zobrazuje)

category = 'LOST_OPPORTUNITY'
reason_code = 'low_is_high_roas'
action = 'increase_bid_or_dedicated_campaign'
```

### Priority při overlap

```
1. LOSER_REST (všechny tiers) → primary
2. LOW_CTR_AUDIT → primary (pokud není LOSER)
3. DECLINING → primary (pokud není 1-2) OR secondary flag
4. LOST_OPPORTUNITY → primary (pokud není 1-3)
5. RISING → primary (pokud není 1-4)
```

`secondary_flags` pole obsahuje všechny aplikovatelné kategorie (např. DECLINING + LOW_CTR současně).

---

## Sheet layouts

### Tab 3: ACTIONS 🆕

**Řazení:** priority_rank asc (1 = nejvyšší priorita).

**Priority formula:**

```
IF category == LOSER_REST:
    priority_score = wasted_spend
ELSE IF category == LOW_CTR:
    priority_score = missed_revenue_potential (imp × baseline_ctr × avg_aov × cvr)
ELSE IF category == DECLINING:
    priority_score = revenue_lost (previous_rev - current_rev)
ELSE IF category == RISING:
    priority_score = growth_potential (current_rev × growth_pct)
ELSE IF category == LOST_OPPORTUNITY:
    priority_score = scaling_potential (rev × (1 - is))
```

**Sloupce:**

| # | Sloupec | Auto/Manual | Popis |
|---|---|---|---|
| 1 | `priority_rank` | Auto | 1..N |
| 2 | `category` | Auto | LOSER_REST / LOW_CTR / DECLINING / RISING / LOST_OPP |
| 3 | `item_id` | Auto | |
| 4 | `product_title` | Auto | Z GMC |
| 5 | `product_price` | Auto | Z shopping_product |
| 6 | `current_campaign` | Auto | Primary (po cost) |
| 7 | `tier` | Auto | zero_conv / low_volume / mid / high / empty |
| 8 | `reason_code` | Auto | Detail důvodu |
| 9-14 | `main_clicks`, `main_impressions`, `main_cost`, `main_conv`, `main_pno_pct`, `main_ctr_pct` | Auto | MAIN metriky (primárně pro flagging) |
| 15-18 | `total_clicks`, `total_cost`, `total_conv`, `total_roas` | Auto | TOTAL overview |
| 19 | `brand_share_pct` | Auto | % total conv z brand kampaní |
| 20 | `growth_pct` | Auto | Pro RISING/DECLINING; empty pro ostatní |
| 21 | `wasted_spend` | Auto | Pro negative kategorie |
| 22 | `recommended_action` | Auto | Text doporučení |
| 23 | `days_since_first_flag` | Auto | |
| 24 | `transition_status` | Auto | NEW / REPEATED / RESOLVED / RE_FLAGGED |
| 25 | `secondary_flags` | Auto | Comma-separated dalších kategorií |
| **26** | ⚡ `action_taken` | **Manual** | "label applied", "fotka změněna", atd. |
| **27** | ⚡ `action_date` | **Manual** | YYYY-MM-DD |
| **28** | ⚡ `consultant_note` | **Manual** | Volný text |

**Styling:** barevné pozadí per category (🟥 LOSER, 🟧 DECLINING, 🟨 LOW_CTR, 🟩 RISING, 🟦 LOST_OPP).

### Tab 4: PRODUCT_TIMELINE 🆕

**Upsert logika:** 1 řádek per item_id. Přepis existujícího (auto columns), preserve manual.

**Sloupce:**

| # | Sloupec | Auto/Manual | Popis |
|---|---|---|---|
| 1 | `item_id` | Auto | PK |
| 2 | `product_title` | Auto | |
| 3 | `first_flag_date` | Auto | Datum první flagace |
| 4 | `total_runs_flagged` | Auto | Count |
| 5 | `categories_history` | Auto | "LOSER_REST → RESOLVED → RE_FLAGGED → LOW_CTR" |
| 6 | `current_status` | Auto | FLAGGED / RESOLVED / STABLE |
| 7-11 | `kpi_before_cost/conv/pno/roas/ctr` | Auto | 30d před first_flag_date (TOTAL) |
| 12-16 | `kpi_current_cost/conv/pno/roas/ctr` | Auto | Aktuální (TOTAL) |
| 17-21 | `kpi_current_main_*` | Auto | Split MAIN |
| 22-26 | `kpi_current_rest_*` | Auto | Split REST |
| 27-31 | `kpi_current_brand_*` | Auto | Split BRAND |
| 32-36 | `delta_*_pct` | Auto | Delta vs before (TOTAL) |
| 37 | `effectiveness_score` | Auto | `+` / `=` / `-` / `N/A` / `PENDING` |
| 38 | `days_since_action` | Auto | (current_date - action_date) |
| **39** | ⚡ `latest_action` | **Manual** | Z ACTIONS merge |
| **40** | ⚡ `latest_action_date` | **Manual** | |
| **41** | ⚡ `latest_note` | **Manual** | |

### Tab 7: WEEKLY_SNAPSHOT 🆕

Append-only, 1 řádek per run.

**Sloupce:**

| # | Sloupec | Popis |
|---|---|---|
| 1 | `run_date` | |
| 2 | `week_id` | "2026-W17" |
| 3 | `account_cost_total` | |
| 4 | `account_conversions` | |
| 5 | `account_conv_value` | |
| 6 | `account_roas` | |
| 7 | `account_pno_pct` | |
| 8 | `account_ctr_pct` | |
| 9 | `flagged_count_total` | |
| 10-14 | `flagged_count_<category>` | Per category |
| 15 | `wasted_spend_total` | |
| 16 | `resolved_this_run` | |
| 17 | `re_flagged_this_run` | |
| 18 | `label_application_rate_pct` | |

DASHBOARD použije posledních 8 řádků pro weekly trend charts.

---

## Effectiveness measurement

### 3 úrovně

1. **Per-produkt:** `effectiveness_score` v PRODUCT_TIMELINE
2. **Run-level:** transitions counts v DASHBOARD
3. **Historical:** weekly trends z WEEKLY_SNAPSHOT

### Effectiveness score logika

```
IF action_date is empty:
    score = 'N/A' (action nebyla zaznamenána)
ELSE IF days_since_action < 14:
    score = 'PENDING' (brzy na vyhodnocení, min 2 týdny pro signal)
ELSE:
    delta_cost_pct = (current_total_cost - before_total_cost) / before_total_cost × 100
    delta_roas_pct = (current_total_roas - before_total_roas) / before_total_roas × 100

    IF delta_cost_pct <= -30 AND delta_roas_pct >= -10:
        score = '+' (cost klesl ≥30%, ROAS stabilní/vyšší) ✅
    ELSE IF delta_cost_pct > -10 OR delta_roas_pct < -30:
        score = '-' (cost nesedí nebo ROAS významně horší) ⚠
    ELSE:
        score = '=' (smíšený signál, inconclusive)
```

### REST campaign health insight

Pro každý produkt v REST kampani (`current_status = RESOLVED`):

```
rest_efficiency = rest_current_cost / kpi_before_main_cost

IF rest_efficiency < 0.2:
    rest_health = 'efficient' (rest má <20% původního main spendu)
ELSE IF rest_efficiency < 0.5:
    rest_health = 'acceptable'
ELSE:
    rest_health = 'wasteful' ⚠ (rest plýtvá podobně jako původní main — špatně nastavená)
```

DASHBOARD ukáže count `wasteful` rest produktů s klickable listem.

### Brand insights

**Brand-only sellers:**
```
IF main_conversions == 0 AND brand_conversions > 0:
    tag: 'brand_only_seller'
    insight_for_client: "tyto produkty prodávají jen přes brand — marketing potřebuje zvednout awareness"
```

**Brand-dependent products:**
```
brand_share = brand_conv_value / total_conv_value

IF brand_share > 0.5:
    tag: 'brand_dependent'
    insight_for_client: "produkt stojí na brand awareness — scaling brand kampaně pomůže"
```

Počítáme a zobrazujeme v DASHBOARD sekci "BRAND INSIGHTS".

---

## Preserve mechanismus pro manual input

### Algoritmus

```
PŘED WRITE do ACTIONS:
1. readExistingActions(sheet) → Map<item_id, {action_taken, action_date, consultant_note}>
2. Pro každý nový řádek:
     manualData = existingMap[row.item_id] || {}
     row.action_taken = manualData.action_taken || ''
     row.action_date = manualData.action_date || ''
     row.consultant_note = manualData.consultant_note || ''
3. clearContents()
4. setValues(newRows)

PŘED WRITE do PRODUCT_TIMELINE:
1. readExistingTimeline(sheet) → Map<item_id, fullRow>
2. Pro každý flagged produkt (+ resolved kandidáti):
     existingRow = existingMap[item_id]
     newRow = {
         ...computedAutoFields,
         latest_action: existingRow?.latest_action || latestActionFromActionsTab,
         latest_action_date: existingRow?.latest_action_date || ...,
         latest_note: existingRow?.latest_note || ...
     }
3. Produkty existujici v map ale ne v current flag set:
     Zachovat řádek (upsert, ne delete), jen update kpi_current_* na aktuální data
     current_status = 'RESOLVED' nebo 'STABLE'
4. clearContents() + setValues(newRows)
```

### Edge cases

- **První run:** existingMap prázdný, manual columns prázdné. OK.
- **Konzultant rozbije schema (přidá sloupec):** detekujeme header mismatch, loguje warning, skip preserve (overwrite celého tabu). Konzultant uvidí warning v logu.
- **Sheet protected:** read fallback na prázdnou mapu, log warning.

---

## Nové CONFIG parametry

Přibudou do CONFIG tabu:

| Parametr | Default | Popis |
|---|---|---|
| `risingGrowthThreshold` | 50 | Min % growth pro RISING (medium) |
| `decliningDropThreshold` | 30 | Min % drop pro DECLINING |
| `minConversionsForTrendCompare` | 3 | Min conv obou periodách |
| `lostOpportunityMinConv` | 5 | Min conv pro LOST_OPP |
| `lostOpportunityMaxPnoMultiplier` | 0.8 | PNO ≤ target × N |
| `lostOpportunityMaxImpressionShare` | 0.5 | IS < N |
| `effectivenessMinDaysSinceAction` | 14 | Počet dní před vyhodnocením účinnosti |
| `restCampaignEfficientThreshold` | 0.2 | Rest cost ≤ N × before_main_cost |

---

## Implementační fáze (doporučené pořadí)

1. **Fáze 1: Data fetching rozšíření**
   - Previous period query
   - Split metrics per campaign type (main/brand/rest)
   - search_impression_share query

2. **Fáze 2: Nové klasifikace**
   - RISING, DECLINING, LOST_OPPORTUNITY classifier funkce
   - Merge priority update

3. **Fáze 3: Preserve mechanismus**
   - readExistingActions
   - readExistingTimeline
   - Merge logic

4. **Fáze 4: Nové taby**
   - ACTIONS writer (s barevným stylingem)
   - PRODUCT_TIMELINE writer (upsert)
   - WEEKLY_SNAPSHOT appender

5. **Fáze 5: DASHBOARD rozšíření**
   - Weekly trend sparklines
   - REST campaign health sekce
   - Brand insights sekce

6. **Fáze 6: Effectiveness logika**
   - computeEffectivenessScore
   - Rest efficiency classification
   - Brand-only / brand-dependent tagging

7. **Fáze 7: README + CONFIG doc update**
   - Nové parametry + tuning tipy
   - Vysvětlení nových kategorií

8. **Fáze 8: Testing**
   - Dry-run na Kabelce
   - Audit false positives v nových kategoriích
   - Validace preserve mechanismu (2 runy s manual input)

---

## Testing strategie

1. **Syntax check** — Node `--check` na combined.gs
2. **MCP validation** — ověřit očekávané results na kabelka.cz před live runem
3. **Dry-run** — `CONFIG.dryRun=true` pro první test
4. **Live 1st run** — vytvoří nové taby, očekáváme:
   - 7-10 LOSER_REST
   - 40-55 LOW_CTR_AUDIT
   - 3-8 RISING
   - 3-8 DECLINING
   - 2-5 LOST_OPPORTUNITY
5. **Manual input test** — konzultant vyplní `action_taken` u 3 produktů, spustí 2. run, ověří že se preservuje
6. **Effectiveness test** — po 14+ dnech s aplikovanými labely ověřit effectiveness_score

---

## Backward compatibility

- Existující output sheety (v1) nebudou kompatibilní. User musí **vytvořit nový sheet** (smazat starý nebo pustit `setupOutputSheet` znovu).
- CONFIG tab: nové parametry budou auto-vyplněné při novém setup. V existujících sheetech bude fallback na defaults.
- Data z LIFECYCLE_LOG z v1 zůstanou platná (schema kompatibilní).

---

## Open questions / deferred

- **STAGNANT kategorie** (0 imp za lookback, produkt v feedu) — odloženo pro v3. Vyžaduje integraci s GMC API pro seznam produktů ve feedu.
- **HERO kategorie** — user odmítl v brainstormingu.
- **Klient-facing PDF export** — skript jen do sheetu. Export do PDF/prezentace je out-of-scope (user dělá manually).
- **Alert hooks (Slack, Discord)** — zatím jen email. Webhook integrace odložena.

---

## Success measurement

- Po 4 týdnech používání: **label application rate** v DASHBOARD ≥ 70%
- Po 8 týdnech: průměrný **effectiveness_score** '+' u RESOLVED produktů ≥ 60%
- **Wasted spend** v DASHBOARD klesá WoW/MoM
- Konzultant může **proaktivně reportovat klientovi** 3-5 insights per měsíc z RISING / LOST_OPPORTUNITY / BRAND_INSIGHTS
