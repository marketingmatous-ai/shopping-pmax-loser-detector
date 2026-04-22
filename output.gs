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
