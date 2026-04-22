# Deployment Guide — Shopping/PMAX Loser Detector

Krok-za-krokem nasazení skriptu pro jeden Google Ads account. Opakuj tento proces per klient.

## Předpoklady

- Google Ads account s aktivními Shopping nebo PMAX kampaněmi
- Oprávnění "Standard access" nebo vyšší (pro Google Ads Scripts)
- Google účet pro ownership skriptu (doporučeno konzultant account, ne klient)

## Krok 1: Vytvoř nový skript v Google Ads UI

1. V Google Ads UI: **Tools & Settings → Bulk Actions → Scripts**
2. Klikni `+` pro vytvoření nového skriptu
3. Pojmenuj: `Shopping PMAX Loser Detector — [Klient]`

## Krok 2: Zkopíruj kód

Google Ads Scripts má **jediný editor** — všechny 6 modulů skriptu (`utils.gs`, `config.gs`, `data.gs`, `classifier.gs`, `output.gs`, `main.gs`) musí být v jednom souboru.

**Použij připravený `combined.gs` soubor z GitHub repa:**

Nejjednodušší cesta — raw link:

👉 [https://raw.githubusercontent.com/marketingmatous-ai/shopping-pmax-loser-detector/main/combined.gs](https://raw.githubusercontent.com/marketingmatous-ai/shopping-pmax-loser-detector/main/combined.gs)

1. Otevři odkaz v prohlížeči
2. **Vyber celý obsah** (Cmd+A), **zkopíruj** (Cmd+C)
3. V Google Ads Scripts editoru **smaž default obsah** a **paste** (Cmd+V)
4. Klikni **Save**

> ℹ️ Alternativně: naklonuj repo a použij lokální `combined.gs`:
> ```bash
> git clone https://github.com/marketingmatous-ai/shopping-pmax-loser-detector.git
> cd shopping-pmax-loser-detector
> pbcopy < combined.gs  # macOS — combined.gs je v clipboardu
> ```
>
> Pokud změníš zdrojové moduly (`*.gs`), regeneruj `combined.gs`:
> ```bash
> bash build-combined.sh
> ```

## Krok 3: Vytvoř output Google Sheet automaticky

Skript má helper funkci `setupOutputSheet()`, která:
- Vytvoří nový Google Sheet ve tvém Drive
- Nastaví 4 taby (FEED_UPLOAD, DETAIL, SUMMARY, LIFECYCLE_LOG) + README tab
- Nastaví headers a frozen rows
- Vypíše ID a URL do logu

**Postup:**

1. V Google Ads Scripts editoru najdi **dropdown vedle tlačítka "Preview"** — obsahuje seznam funkcí
2. Vyber **`setupOutputSheet`**
3. Klikni **Run** (ne Preview — chceme reálně vytvořit sheet)
4. Pokud skript požaduje authorization, schval ho:
   - "Review permissions" → vyber Google účet → "Advanced" → "Go to [Script Name] (unsafe)" → Allow
5. Po dokončení otevři **Logs** (`Cmd+Enter` nebo View → Logs)
6. Najdi řádky:
   ```
   NOVY OUTPUT SHEET VYTVOREN
   Nazev: Shopping-PMAX Loser Detector — [Klient] (123-456-7890)
   ID:    1ABC...xyz
   URL:   https://docs.google.com/spreadsheets/d/1ABC...xyz
   ```
7. **Zkopíruj ID** (dlouhý řetězec po `ID:`)

> ℹ️ Sheet se vytvoří v Drive Google účtu, pod kterým skript běží. Typicky tvůj konzultantský účet. Je to očekávané — skript má tak rovnou access pro write.

## Krok 4: Edit CONFIG v main.gs

Vrať se do editoru skriptu. Najdi na začátku (v části combined file po utils/config/data/classifier/output) `var CONFIG = {`.

### Povinné pole

```javascript
outputSheetId: '1ABC...xyz',     // <<< ID z Kroku 3
targetPnoPct:  30,               // Cílové PNO klienta (např. 30, 40, 50)
```

### Doporučená pole

```javascript
adminEmail: 'konzultant@email.cz',  // Kam posílat týdenní email report
lookbackDays: 30,                    // 30 default, zkus 60 pro low-volume účty
customLabelIndex: 2,                 // Ověř, že klient label_2 nepoužívá; jinak zvol jiné číslo
```

### Parametry k ověření podle klienta

```javascript
brandCampaignPattern: '(?i)BRD',    // Uprav podle naming conventions klienta
                                     // Např.: '(?i)(BRD|BRAND|ZNACKA)'
restCampaignPattern:  '(?i)REST',   // Pokud má klient rest kampaně jinak pojmenované
```

### Default hodnoty jsou industry-aligned

Ostatní parametry (sample size, tiers, thresholds) mají industry-sensible defaults. Upravuj jen pokud:

- **Malý account s málo konverzemi**: sniž `minExpectedConvFloor` z 3 na 2, zvyš `lookbackDays` na 60-90
- **Luxury / low-volume**: zvyš `lookbackDays` na 60-90, sniž `minImpressionsLowCtr` na 500
- **Velký account s >100K produkty**: zvyš `maxRowsDetailTab`, zvyš `minClicksAbsolute` na 200

**Save skript po každé změně.**

## Krok 5: Dry-run test

1. V CONFIG nastav:
   ```javascript
   dryRun: true,
   ```
2. V dropdownu zvol funkci **`main`**
3. Klikni **Preview** (oranžové tlačítko)
4. Počkej na dokončení (typicky 2-10 min podle velikosti účtu)
5. Zkontroluj **Logs**:
   - Žádné ERROR — OK
   - `INFO: DRY RUN — nic se nezapsalo do sheetu.`
   - Funnel counts dávají smysl?
   - Classified count vs flagged count — rozumný poměr (typicky 5-20% produktů flagged)

### Časté chyby a řešení

| Error | Řešení |
|---|---|
| `CONFIG validation failed` | Oprav CONFIG podle popisu v chybě |
| `outputSheetId neni otevirateln` | Zkontroluj, že jsi zkopíroval správné ID ze Setup logů |
| `Account nema zadna Shopping/PMAX data` | Ověř aktivní kampaně; zvyš `lookbackDays` |
| Skript timeout (30 min) | Sniž `lookbackDays`, nastav `maxRowsDetailTab` menší |
| GAQL query failed | Možná changed API — nahlaš issue |

## Krok 6: Live run

1. V CONFIG nastav:
   ```javascript
   dryRun: false,
   ```
2. Klikni **Preview** pro dalˆí kontrolu logs
3. Pokud vše OK, klikni **Run**

## Krok 7: Verifikuj output

Otevři svůj Google Sheet (URL z Kroku 3). Měly by být 5 tabů (README + 4 datové):

### `FEED_UPLOAD` — sanity check

- Je tam řádek `id | custom_label_X`?
- Produkty mají smysluplné item_id (ne prázdné nebo s mezerami)?
- Pokud je prázdný (jen header), skript neflaggoval nic — buď dobrá zpráva, nebo thresholds příliš přísné

### `DETAIL` — audit false-positives

1. Vyber 10 náhodných flagged produktů (filter `primary_label != ''`)
2. Pro každý otevři Google Ads UI → Products tab → hledej podle item_id
3. Ověř:
   - Čísla (cost, conversions) souhlasí s UI?
   - Produkt je opravdu underperforming, nebo jde o edge case?
   - `reason_code` dává smysl?
4. **Target: ≤ 2/10 false-positive.** Pokud víc → zpřísni thresholds (viz README)

### `SUMMARY` — dashboard

- Account baseline je plausible?
- Funnel je rozumný? (90%+ produktů prošlo filtry — pokud <50%, něco je špatně)
- Flags breakdown — v rámci očekávání?
- Top 10 losers — souhlasí s tvou intuicí o účtu?

### `LIFECYCLE_LOG` — při první runu prázdný (jen header)

- První run má jen NEW_FLAG transitions
- Po 2-3 týdnech se začnou objevovat RESOLVED, REPEATED, RE_FLAGGED

## Krok 8: Aplikuj labels v GMC / Mergado

### 🥇 Doporučeno: GMC Supplemental Feed s Google Sheets (auto-refresh)

**Toto je produkční setup** — feed se automaticky aktualizuje po každém týdením runu skriptu bez manuálního uploadu.

#### 8.1 Přidej nový supplemental feed v GMC

1. V Merchant Center: **Marketing → Feeds → Supplemental feeds** (nebo **Products → Feeds**)
2. Klikni **+** (Add supplemental source)
3. **Zdroj dat:** `Google Tabulky` / Google Sheets
4. **Vyber existující Tabulku** (ne "Generate new")
5. Vlož URL output sheetu klienta:
   ```
   https://docs.google.com/spreadsheets/d/<ID_SHEETU>/edit
   ```
6. ⚠️ **KRITICKÉ:** v dialogu **"Název listu" vyber `FEED_UPLOAD`** (NE default první tab = README, GMC by pak četl chybné data).

#### 8.2 Nastav parametry feedu

- **Název feedu:** `Loser Detector — custom_label_N` (N = CONFIG.customLabelIndex)
- **Země prodeje:** dle klienta (CZ / SK / ...)
- **Jazyk:** dle klienta
- **Plán načítání (Fetch schedule):**
  - Frequency: **Denně**
  - Čas: **08:00** (pondělí ráno po týdenním runu skriptu)
  - Časové pásmo: **Europe/Prague**

#### 8.3 Přiřaď k primary feed

1. GMC → **Primary feeds** → klikni na primární feed klienta
2. Záložka **Settings / Pravidla** → sekce **Supplemental feeds**
3. **Add supplemental feed** → vyber náš nový feed
4. Uložit

#### 8.4 Ověření

Po 5-30 min otevři supplemental feed → **Processing**:

- ✅ Názvy atributů: `id`, `custom_label_N`
- ✅ Celkový počet aktualizovaných produktů: cca stejný jako `flagged + healthy` v Logger výstupu skriptu
- ✅ Shodující se produkty: ideálně = celkový (pokud 100% item_id existuje v primary)
- ❌ Pokud vidíš chyby typu "Neplatný atribut id" nebo "Příliš mnoho hodnot" → GMC čte špatný tab, zopakuj krok 8.1 s explicit výběrem `FEED_UPLOAD`

#### 8.5 Nastavit listing groups v Google Ads

V **rest kampani** (Shopping nebo PMAX):

- **Varianta A:** Listing group filter `custom_label_N = loser_rest` → jen losery do rest
- **Varianta B:** Listing group filter `custom_label_N != healthy` → všechno kromě zdravých do rest
- **Varianta C (rozšířená):** `custom_label_N IN (loser_rest, low_ctr_audit, DECLINING)` → negativní kategorie do rest, RISING a LOST_OPPORTUNITY zůstávají v main (scaling)

V **main kampani** obráceně — exclude `loser_rest` nebo include `healthy`:

- `custom_label_N = healthy` → jen ověřené zdravé produkty
- Nebo: `custom_label_N != loser_rest AND != low_ctr_audit` → vše kromě evidentně problémových

### 🥈 Alternativa: Manuální CSV upload (jednorázově / bez auto-refresh)

1. V sheetu otevři tab **FEED_UPLOAD**
2. `Soubor → Stáhnout → Hodnoty oddělené čárkou (.csv)` — stáhne JEN aktivní tab
3. GMC → supplemental feed → **Upload file** → vyber stažený CSV
4. Nevýhoda: nutno opakovat každý týden po runu skriptu

### 🥉 Přes Mergado

1. V Mergado vytvoř rule: "IF custom_label_N IS IN [list ze sheetu] THEN ..."
2. Uploaduj seznam item_id z FEED_UPLOAD tabu
3. Aktivuj rule

## Krok 9: Nastav weekly schedule

V Google Ads Scripts UI:

1. Klikni na skript v seznamu
2. Najdi sekci **Frequency**: Weekly, **Day**: Monday, **Time**: 07:00 (CET)
3. Save

Skript poběží každý týden. Pokud konfiguroval `adminEmail`, dostaneš report emailem.

## Krok 10: Po 4 týdnech — review effectiveness

Otevři SUMMARY tab:

- **Transitions this run**: rostou RESOLVED count?
- **Label application rate**: ideálně >70% (flagy se aplikují)
- **REPEATED_WARNING**: produkty >2 runy bez změny — label se neaplikoval (GMC issue?)
- **RE_FLAGGED**: produkty vracející se — thresholds příliš citlivé?

Pokud je application rate nízký:
- Ověř, že GMC Supplemental Feed je aktivní
- Ověř, že konzultant uploaduje CSV pravidelně
- Zvaž automatizaci uploadu

Zapiš findings do svého interního learning logu (nebo [otevři issue v repu](https://github.com/marketingmatous-ai/shopping-pmax-loser-detector/issues) pro týmové sdílení).

## Troubleshooting common scenarios

### Sheet nabobtnal >10M cells

- LIFECYCLE_LOG roste každým runem (jen transitions, ale u velkých účtů může být hodně)
- Archivuj starší záznamy do separátního sheetu (zatím ruční)
- Deferred: auto-archive >180 dní

### Skript běží >25 min

- Sniž `lookbackDays`
- Nastav `maxRowsDetailTab` menší
- Zkontroluj, zda account nemá neuvěřitelně velký feed (>1M produktů)

### Email limits

- Free Gmail: 100 emailů/den. Weekly schedule = 1 email/týden/klient — OK pro 5-10 klientů
- Workspace: 1500/den — dostatečně

### Setup sheet failed — "You do not have permission"

- Skript může vytvářet soubory v tvém Drive, ale potřebuje authorization
- Při prvním runu `setupOutputSheet` schval prompted permissions
- Pokud permission denied persists: prekontroluj, že Script běží pod správným Google účtem (ne klient account)

## Deployment checklist (per klient)

- [ ] Nový Google Ads Script vytvořen s name "Shopping PMAX Loser Detector — [Klient]"
- [ ] combined.gs zkopírován do editoru
- [ ] `setupOutputSheet()` spuštěno, sheet vytvořen, ID zkopírováno
- [ ] CONFIG.outputSheetId nastaven
- [ ] CONFIG.targetPnoPct nastaven per klient
- [ ] CONFIG.lookbackDays zvážen podle velikosti účtu
- [ ] CONFIG.customLabelIndex zvolen (nekonfliktní s existujícím GMC setupem klienta)
- [ ] brandCampaignPattern / restCampaignPattern ověřeny podle naming conventions
- [ ] CONFIG.adminEmail nastaven (volitelné)
- [ ] Dry-run test pass (žádné ERROR)
- [ ] Manuální audit 10 náhodných flagů (≤2/10 false-positive)
- [ ] Live run první
- [ ] Verify 5 tabů v sheetu (README + 4 datové)
- [ ] Apply labels v GMC / Mergado
- [ ] Weekly schedule nastaveno
- [ ] Po 4 týdnech: review effectiveness

## Tipy pro efektivní per-klient deploy

1. **Šablony CONFIG**: Uchovávej si lokálně vedle sebe CONFIG hodnoty pro každého klienta (sensible defaults + klientské parametry). Text file v klientově složce `/klienti/[klient]/ppc-loser-detector-config.txt`.
2. **Pojmenování sheetů**: Konzistentní prefix "Shopping-PMAX Loser Detector — " usnadní filtering v Drive.
3. **Scheduling timing**: Pondělí ráno (po víkendu — čerstvá data z víkendového trafficu).
4. **Attribution**: Pokud klient má custom attribution (GA4 data-driven), zvýš `minExpectedConvFloor` o 1-2 bod pro větší buffer.
