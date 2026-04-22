# Healthy label pro FEED_UPLOAD — design

**Datum:** 2026-04-21
**Verze:** 1.0
**Status:** Approved, ready for implementation

## Problém

FEED_UPLOAD tab obsahuje jen flagged produkty (loser_rest / low_ctr_audit / DECLINING / RISING / LOST_OPPORTUNITY). V rest kampani specialista nastavuje listing group filter `custom_label_N = loser_rest` aby jen losery dostaly rest traffic. Ale alternativní přístup — **vyloučit zdravé produkty z rest kampaně explicitně** — není možný, protože zdravé produkty nemají žádnou custom_label hodnotu.

Use case: specialista chce nastavit v rest kampani `custom_label_N != healthy` aby zdravé produkty byly garantovaně jen v main kampaních.

## Řešení

Přidat volitelný label pro zdravé produkty, které prošly revizí.

### Kritérium „prošel revizí"

Produkt dostane `healthy` label pokud:
- `c.status === 'ok'` — prošel všemi gates (age gate 30+ dní, sample size gate 30+ kliků, data quality check)
- `!c.primaryLabel` — klasifikátor ho nevyhodil jako problémový

Zombie produkty (INSUFFICIENT_DATA, NEW_PRODUCT_RAMP_UP) a data quality issues dostávají **žádný label** — nejsou ověřené ani jako dobré, ani jako špatné.

### Název štítku

Default: `healthy` — krátké, výstižné, dobře fungující v GMC feed rule kontextu.

Uživatel může přepsat v CONFIG.labelHealthyValue.

### Opt-out

Pokud uživatel nechce healthy v feedu (např. klient má jinou kampanní strukturu), může nastavit `CONFIG.labelHealthyValue = ''` (prázdný string). Skript pak healthy produkty vynechá — chování stejné jako před touto změnou.

## Implementace — 3 soubory

### 1. `_config.gs`

Přidat do sekce LABEL KONFIGURACE:

```javascript
labelHealthyValue:      'healthy',  // Produkty co prosly revizi (status='ok').
                                    // '' = nezapisovat do feedu (opt-out).
```

### 2. `config.gs` — validate()

Přidat:

```javascript
if (config.labelHealthyValue && typeof config.labelHealthyValue !== 'string') {
  errors.push('labelHealthyValue musi byt string nebo "" (vypnuto)');
}
if (config.labelHealthyValue && config.labelHealthyValue === config.labelLoserRestValue) {
  errors.push('labelHealthyValue a labelLoserRestValue nesmi byt stejne');
}
if (config.labelHealthyValue && config.labelHealthyValue === config.labelLowCtrValue) {
  errors.push('labelHealthyValue a labelLowCtrValue nesmi byt stejne');
}
```

### 3. `output.gs` — writeFeedUploadTab()

Rozšíření loopu:

```javascript
var flaggedCount = 0;
var healthyCount = 0;
for (var i = 0; i < classified.length; i++) {
  var c = classified[i];
  if (c.primaryLabel && c.primaryLabel.length > 0) {
    data.push([c.itemId, c.primaryLabel]);
    flaggedCount++;
  } else if (config.labelHealthyValue &&
             c.status === 'ok' &&
             (!c.primaryLabel || c.primaryLabel.length === 0)) {
    data.push([c.itemId, config.labelHealthyValue]);
    healthyCount++;
  }
}

// Logger zmenit:
Logger.log('INFO: FEED_UPLOAD — zapsano ' + flaggedCount + ' flagged' +
           (healthyCount > 0 ? ' + ' + healthyCount + ' healthy' : '') +
           ' = ' + (flaggedCount + healthyCount) + ' radku.');
```

**Pořadí řádků:** flagged nejdřív (zachovává původní priority ordering), pak healthy na konci.

### 4. README.md — sekce FEED_UPLOAD

Rozšířit popis + přidat instrukce pro rest kampaň filter setup.

## Co se NEmění

- ACTIONS tab (healthy tam nepatří — není co řešit)
- SUMMARY tab (klasifikační funnel zůstává)
- DETAIL tab (per-product trace beze změny)
- LIFECYCLE_LOG (healthy není transition, nevytváří záznam)
- MONTHLY_REVIEW (agregace transitions, healthy se neúčastní)
- PRODUCT_TIMELINE (sleduje jen flagged produkty)

## Backward compatibility

- Default `labelHealthyValue: 'healthy'` → healthy produkty se nově objeví ve feedu
- Pokud existující klient nechce změnu, stačí nastavit `''` — chování jako před verzí

## Validace úspěchu

1. Po deploy skript vypíše: `FEED_UPLOAD — zapsano 34 flagged + 63 healthy = 97 radku.`
2. V FEED_UPLOAD tabu řádky 2-35 obsahují flagged labels, řádky 36-98 mají `healthy`
3. Config tab v sheetu zobrazuje `labelHealthyValue: healthy`
4. V Google Ads lze v rest kampani nastavit filter `custom_label_N != healthy`

## Odhadovaný rozsah

- 1 CONFIG hodnota
- ~10 řádků validace
- ~8 řádků v writeFeedUploadTab
- README update 1 sekce
- Bez migrace dat (fresh run přepíše FEED_UPLOAD)
