function handleStock(e) {
  const sh = e.source.getActiveSheet();
  const ss = e.source;
  const turnedOn  = (e.value === "TRUE") || (e.value === true);
  const turnedOff = (e.value === "FALSE") || (e.value === false);
  const CLEAR_ON_UNCHECK = false;

  const sheetName = sh && typeof sh.getName === 'function' ? sh.getName() : 'Stock';

  function resolveLayoutName_(name) {
    const layout = SHEET_LAYOUT || {};
    if (layout[name]) return name;
    const normalizedName = normText_(name || '');
    const match = Object.keys(layout).find(key => normText_(key) === normalizedName);
    if (match) return match;
    if (layout.STOCK) return 'STOCK';
    if (layout.Stock) return 'Stock';
    return name || 'Stock';
  }

  const layoutName = resolveLayoutName_(sheetName);
  const HEADER_ROW = getSheetHeaderRow_(layoutName);
  const DATA_START_ROW = getSheetDataStartRow_(layoutName);

  // Ignore les edits au-dessus des données
  if (e.range.getRow() < DATA_START_ROW) {
    logDebug_('handleStock.skip', {
      ligne: e.range.getRow(),
      colonne: e.range.getColumn(),
      headerRow: HEADER_ROW,
      dataStartRow: DATA_START_ROW,
      feuille: sheetName
    });
    return;
  }

  ss.toast(
    `${sheetName} • Ligne ${e.range.getRow()} • Colonne ${e.range.getColumn()} • DataStart ${DATA_START_ROW} • Layout ${layoutName}`,
    'Diagnostic Stock',
    4
  );

  const stockHeaders = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];

  const resolver = makeHeaderResolver_(stockHeaders);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID      = colExact(HEADERS.STOCK.ID);
  const C_LABEL   = colExact(HEADERS.STOCK.LIBELLE)
    || colExact(HEADERS.STOCK.LIBELLE_ALT)
    || colExact(HEADERS.STOCK.LIBELLE_ALT2)
    || colExact(HEADERS.STOCK.ARTICLE)
    || colExact(HEADERS.STOCK.ARTICLE_ALT)
    || colExact(HEADERS.STOCK.ARTICLE_ALT2)
    || colWhere(h => h.includes('libell'))
    || colWhere(h => h.includes('article'))
    || 2;
  const C_OLD_SKU = colExact(HEADERS.STOCK.OLD_SKU);
  const C_SKU     = colExact(HEADERS.STOCK.SKU)
    || colExact(HEADERS.STOCK.REFERENCE)
    || colWhere(h => h.includes('sku')); // Compat titres variés
  const C_PRIX    = colExact(HEADERS.STOCK.PRIX_VENTE)
    || colExact(HEADERS.STOCK.PRIX_VENTE_ALT2)
    || colWhere(h => h.includes("prix") && h.includes("vente"));
  const C_TAILLE  = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const tailleHeaderLabel = getHeaderLabel_(resolver, C_TAILLE, HEADERS.STOCK.TAILLE);
  const C_LOT     = colExact(HEADERS.STOCK.LOT)
    || colExact(HEADERS.STOCK.LOT_ALT2)
    || colExact(HEADERS.STOCK.LOT_ALT3)
    || colWhere(h => h.includes('lot'));
  const C_DMS     = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK)
    || colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK_ALT2) // "MIS EN STOCK"
    || colWhere(h => h.includes('mise en stock'));

  const combinedMisCol = resolveCombinedMisEnLigneColumn_(resolver);
  const legacyMisCols = combinedMisCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyMisEnLigneColumn_(resolver);
  const useCombinedMisCol = !!combinedMisCol;

  const C_MIS     = useCombinedMisCol ? combinedMisCol : legacyMisCols.checkboxCol;
  const C_DMIS    = useCombinedMisCol ? combinedMisCol : legacyMisCols.dateCol;
  const combinedPubCol = resolveCombinedPublicationColumn_(resolver);
  const legacyPubCols = combinedPubCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyPublicationColumns_(resolver);
  const useCombinedPubCol = !!combinedPubCol;

  const combinedVenduCol = resolveCombinedVenduColumn_(resolver);
  const legacyVenduCols = combinedVenduCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyVenduColumns_(resolver);
  const useCombinedVenduCol = !!combinedVenduCol;

  const C_PUB     = useCombinedPubCol ? combinedPubCol : legacyPubCols.checkboxCol;
  const C_DPUB    = useCombinedPubCol ? combinedPubCol : legacyPubCols.dateCol;
  const C_VENDU   = useCombinedVenduCol ? combinedVenduCol : legacyVenduCols.checkboxCol;
  let   C_DVENTE  = useCombinedVenduCol ? combinedVenduCol : legacyVenduCols.dateCol;
  if (!C_DVENTE) C_DVENTE = colExact(HEADERS.STOCK.DATE_VENTE_ALT) || 10;
  const C_STAMPV  = colExact(HEADERS.STOCK.VENTE_EXPORTEE_LE);
  const C_VALIDE  = colExact(HEADERS.STOCK.VALIDER_SAISIE)
    || colExact(HEADERS.STOCK.VALIDER_SAISIE_ALT)
    || colWhere(h => h.includes("valider"));

  const C_REPUB = colExact(HEADERS.STOCK.REPUBLIE)
  || colExact(HEADERS.STOCK.REPUBLIE_ALT)
  || colExact(HEADERS.STOCK.REPUBLIE_ALT2)
  || colWhere(h => h.toLowerCase().includes('republ'));
  const C_BOOST = colExact(HEADERS.STOCK.BOOST)
    || colExact(HEADERS.STOCK.BOOST_ALT)
    || colWhere(h => h.toLowerCase().includes('boost'));

  if (!C_VENDU || !C_DVENTE) {
    const missing = [];
    if (!C_VENDU) missing.push('VENDU');
    if (!C_DVENTE) missing.push('DATE DE VENTE');
    const message = `Colonnes manquantes: ${missing.join(', ')} — layout ${layoutName}, headerRow ${HEADER_ROW}, dataStart ${DATA_START_ROW}`;
    logDebug_('handleStock.missingColumns', {
      message,
      layoutName,
      headerRow: HEADER_ROW,
      dataStartRow: DATA_START_ROW,
      headersPreview: stockHeaders.slice(0, 15)
    });
    ss.toast(message, 'Stock', 6);
    return;
  }

  const c = e.range.getColumn(), r = e.range.getRow();
  if (r === HEADER_ROW) return;

  logDebug_('handleStock', {
    ligne: r,
    colonne: c,
    enteteCourant: stockHeaders[c - 1],
    enteteRepublie: C_REPUB ? stockHeaders[(C_REPUB || 1) - 1] : '',
    enteteBoost: C_BOOST ? stockHeaders[(C_BOOST || 1) - 1] : '',
    ranges: { C_PUB, C_REPUB, C_VENDU, C_DVENTE }
  });

  if (C_REPUB && c === C_REPUB) {
  if (!turnedOn) return;

  const sold = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
  if (sold) {
    sh.getRange(r, C_REPUB).setValue(false);
    ss.toast('Impossible de republier un produit déjà vendu.', 'Stock', 5);
    logDebug_('handleStock.republie', { ligne: r, statut: 'refusé', raison: 'déjà vendu' });
    return;
  }

  if (!C_LABEL) {
    sh.getRange(r, C_REPUB).setValue(false);
    ss.toast('Colonne LIBELLÉ introuvable (badge republication).', 'Stock', 6);
    logDebug_('handleStock.republie', { ligne: r, statut: 'refusé', raison: 'libellé manquant' });
    return;
  }

  const labelCell = sh.getRange(r, C_LABEL);
  const currentLabel = String(labelCell.getValue() || '');
  labelCell.setValue(bumpRepublishBadge_(currentLabel));

  sh.getRange(r, C_REPUB).setValue(false);
  logDebug_('handleStock.republie', { ligne: r, statut: 'badge appliqué', ancienLibelle: currentLabel });
  return;
}

  if (C_BOOST && c === C_BOOST) {
    if (!turnedOn) return;

    const sold = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    if (sold) {
      sh.getRange(r, C_BOOST).setValue(false);
      ss.toast('Impossible de booster un produit déjà vendu.', 'Stock', 5);
      logDebug_('handleStock.boost', { ligne: r, statut: 'refusé', raison: 'déjà vendu' });
      return;
    }

    if (!C_LABEL) {
      sh.getRange(r, C_BOOST).setValue(false);
      ss.toast('Colonne LIBELLÉ introuvable (badge boost).', 'Stock', 6);
      logDebug_('handleStock.boost', { ligne: r, statut: 'refusé', raison: 'libellé manquant' });
      return;
    }

    const labelCell = sh.getRange(r, C_LABEL);
    const currentLabel = String(labelCell.getValue() || '');
    const boostedLabel = applyBoostBadge_(currentLabel);

    labelCell.setValue(boostedLabel);
    sh.getRange(r, C_BOOST).setValue(false);

    ss.toast('Badge boost appliqué.', 'Stock', 4);
    logDebug_('handleStock.boost', { ligne: r, statut: 'badge appliqué', ancienLibelle: currentLabel, nouveauLibelle: boostedLabel });
    return;
  }


  const chronoCols = {
    dms: C_DMS,
    dmis: C_DMIS,
    dpub: C_DPUB,
    dvente: C_DVENTE
  };

  const chronology = [
    { key: 'dms', column: chronoCols.dms, statusCol: 0 },
    { key: 'dmis', column: chronoCols.dmis, statusCol: C_MIS },
    { key: 'dpub', column: chronoCols.dpub, statusCol: C_PUB },
    { key: 'dvente', column: chronoCols.dvente, statusCol: C_VENDU }
  ];

  function findChronologyIndex_(key) {
    return chronology.findIndex(entry => entry.key === key);
  }

  function getChronoDate_(key) {
    const entry = chronology.find(item => item.key === key);
    if (!entry || !entry.column) {
      return null;
    }

    const value = sh.getRange(r, entry.column).getValue();
    return getDateOrNull_(value);
  }

  function computeDateFromPrevious_(key) {
    const now = new Date();
    const index = findChronologyIndex_(key);
    if (index <= 0) {
      return now;
    }

    const previous = chronology[index - 1];
    const baseDate = getChronoDate_(previous.key);
    if (baseDate && baseDate instanceof Date && !isNaN(baseDate)) {
      return baseDate.getTime() > now.getTime() ? new Date(baseDate.getTime()) : now;
    }

    return now;
  }

  function setChronoDateFromPrevious_(key) {
    const entry = chronology.find(item => item.key === key);
    if (!entry || !entry.column) {
      return null;
    }

    const nextDate = computeDateFromPrevious_(key);
    const cell = sh.getRange(r, entry.column);
    cell.setValue(nextDate);
    cell.setNumberFormat('dd/MM/yyyy');
    return nextDate;
  }

  function propagateForwardFrom_(key) {
    const startIndex = findChronologyIndex_(key);
    if (startIndex < 0) {
      return;
    }

    for (let i = startIndex + 1; i < chronology.length; i++) {
      const entry = chronology[i];
      if (!entry.column || !entry.statusCol) {
        continue;
      }

      const statusValue = sh.getRange(r, entry.statusCol).getValue();
      if (!isStatusActiveValue_(statusValue)) {
        continue;
      }

      setChronoDateFromPrevious_(entry.key);
    }
  }

  function getChronoKeyByColumn_(column) {
    const match = chronology.find(entry => entry.column === column);
    return match ? match.key : null;
  }

  function setCellToFallback_(col, fallback) {
    if (!col) return;
    const cell = sh.getRange(r, col);
    if (restorePreviousCellValue_(sh, r, col)) {
      return;
    }

    if (fallback === undefined || fallback === null || fallback === '') {
      cell.clearContent();
      return;
    }

    if (fallback instanceof Date && !isNaN(fallback.getTime())) {
      cell.setValue(fallback);
      return;
    }

    if (typeof fallback === 'number') {
      const parsed = getDateOrNull_(fallback);
      if (parsed) {
        cell.setValue(parsed);
        return;
      }
    }

    if (typeof fallback === 'string') {
      const parsed = getDateOrNull_(fallback);
      if (parsed) {
        cell.setValue(parsed);
        return;
      }
    }

    cell.setValue(fallback);
  }

  function revertCheckbox_(range, oldValue) {
    if (!range) return;
    let valueToSet = oldValue;
    if (oldValue === 'TRUE') valueToSet = true;
    if (oldValue === 'FALSE' || oldValue === undefined) valueToSet = false;
    if (valueToSet === null || valueToSet === '') {
      range.clearContent();
      return;
    }
    range.setValue(valueToSet);
  }

  function buildCheckboxRule_() {
    return SpreadsheetApp
      .newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();
  }

  function isCheckboxValidation_(validation) {
    return !!(validation
      && typeof validation.getCriteriaType === 'function'
      && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX);
  }

  function restoreCheckboxValidation_(range, previousValidation) {
    if (!range) return;
    if (previousValidation && isCheckboxValidation_(previousValidation)) {
      range.setDataValidation(previousValidation);
    } else {
      range.setDataValidation(buildCheckboxRule_());
    }
    if (range.getValue() === '' || range.getValue() === null) {
      range.setValue(false);
    }
  }

  function ensureChronologyOrRevert_(changedKey, fallback, checkboxInfo) {
    const result = enforceChronologicalDates_(sh, r, chronoCols);
    if (result.ok) {
      return true;
    }

    setCellToFallback_(chronoCols[changedKey], fallback);
    if (checkboxInfo && checkboxInfo.range) {
      revertCheckbox_(checkboxInfo.range, checkboxInfo.oldValue);
    }
    ss.toast(result.message || 'Ordre chronologique des dates invalide.', 'Stock', 6);
    return false;
  }

  // 0) Modification de SKU(ancienne nomenclature) → renumérotation (globale)
  if (c === C_OLD_SKU) {
    renumberStockByBrand_();
    return;
  }

  // 0bis) Modification du PRIX DE VENTE
  if (C_PRIX && c === C_PRIX) {
    const vendu = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    const priceCell = sh.getRange(r, C_PRIX);
    const priceValue = priceCell.getValue();
    const priceDisplay = priceCell.getDisplayValue();

    if (C_VALIDE) {
      const valCell = sh.getRange(r, C_VALIDE);
      const validation = valCell.getDataValidation();
      const isCheckbox = validation &&
        typeof validation.getCriteriaType === 'function' &&
        validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
      const allowInvalid = validation && typeof validation.getAllowInvalid === 'function' && validation.getAllowInvalid();
      const shouldEnable = (typeof priceValue === 'number') && !isNaN(priceValue) && priceValue > 0 && (!priceDisplay || priceDisplay.indexOf('⚠️') !== 0);

      if (shouldEnable) {
        if (!isCheckbox || allowInvalid) {
          const rule = SpreadsheetApp
            .newDataValidation()
            .requireCheckbox()
            .setAllowInvalid(false)
            .build();
          valCell.setDataValidation(rule);
          if (!isCheckbox) {
            valCell.setValue(false);
          }
        }
      } else {
        valCell.clearDataValidations();
        valCell.clearContent();
      }
    }

    if (!vendu) {
      // Rien n'est coché → pas d'alerte.
      clearPriceAlertIfAny_(sh, r, C_PRIX);
      return;
    }

    // La colonne VENDU est cochée → contrôle du prix
    ensureValidPriceOrWarn_(sh, r, C_PRIX);
    return;
  }

  const isCombinedMisCell = useCombinedMisCol && C_MIS && C_DMIS && C_MIS === C_DMIS;
  const isCombinedPubCell = useCombinedPubCol && C_PUB && C_DPUB && C_PUB === C_DPUB;
  const isCombinedVenduCell = useCombinedVenduCol && C_VENDU && C_DVENTE && C_VENDU === C_DVENTE;

  if (c === C_DMS
    || (!isCombinedMisCell && c === C_DMIS)
    || (!isCombinedPubCell && c === C_DPUB)
    || (!isCombinedVenduCell && c === C_DVENTE)) {
    const key = getChronoKeyByColumn_(c);
    if (!key) {
      return;
    }
    if (!ensureChronologyOrRevert_(key, e.oldValue)) {
      return;
    }
    propagateForwardFrom_(key);
    if (key !== 'dvente') {
      return;
    }
  }

  // 1) MIS EN LIGNE → horodate
  if (isCombinedMisCell && C_MIS && c === C_MIS) {
    const cell = sh.getRange(r, C_MIS);
    const headerLabel = stockHeaders[C_MIS - 1] || HEADERS.STOCK.MIS_EN_LIGNE;
    const previousValidation = cell.getDataValidation();
    const wasCheckbox = isCheckboxValidation_(previousValidation);
    const published = C_PUB ? isStatusActiveValue_(sh.getRange(r, C_PUB).getValue()) : false;
    const value = cell.getValue();
    const oldValue = e.oldValue;
    const oldValueDate = getDateOrNull_(oldValue);
    const checkboxInfo = !oldValueDate ? { range: cell, oldValue } : null;

    if (turnedOff) {
      if (published) {
        cell.setValue(true);
        ss.toast(
          `Impossible de décocher "${headerLabel}" tant que "${HEADERS.STOCK.PUBLIE}" est coché.`,
          'Stock',
          5
        );
        return;
      }
      if (!ensureChronologyOrRevert_('dmis', oldValue, checkboxInfo) && wasCheckbox) {
        restoreCheckboxValidation_(cell, previousValidation);
      }
      return;
    }

    if (turnedOn) {
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      setChronoDateFromPrevious_('dmis');
      const info = { range: cell, oldValue };
      if (!ensureChronologyOrRevert_('dmis', oldValue, info)) {
        restoreCheckboxValidation_(cell, previousValidation);
        return;
      }
      propagateForwardFrom_('dmis');
      return;
    }

    if (value === '' || value === null) {
      if (published) {
        setCellToFallback_(C_DMIS, oldValue);
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        ss.toast(
          `Impossible de vider "${headerLabel}" tant que "${HEADERS.STOCK.PUBLIE}" est coché.`,
          'Stock',
          5
        );
        return;
      }

      if (!ensureChronologyOrRevert_('dmis', oldValue, checkboxInfo)) {
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }

      restoreCheckboxValidation_(cell, previousValidation);
      cell.clearContent();
      return;
    }

    const parsedValue = getDateOrNull_(value);
    if (parsedValue) {
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      cell.setValue(parsedValue);
      cell.setNumberFormat('dd/MM/yyyy');
      if (!ensureChronologyOrRevert_('dmis', oldValue, checkboxInfo)) {
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }
      propagateForwardFrom_('dmis');
      return;
    }

    setCellToFallback_(C_DMIS, oldValue);
    if (oldValueDate) {
      cell.clearDataValidations();
      cell.setNumberFormat('dd/MM/yyyy');
    } else if (wasCheckbox) {
      restoreCheckboxValidation_(cell, previousValidation);
    }
    return;
  }

  if (isCombinedPubCell && C_PUB && c === C_PUB) {
    const cell = sh.getRange(r, C_PUB);
    const headerLabel = stockHeaders[C_PUB - 1] || HEADERS.STOCK.PUBLIE;
    const previousValidation = cell.getDataValidation();
    const wasCheckbox = isCheckboxValidation_(previousValidation);
    const sold = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    const value = cell.getValue();
    const oldValue = e.oldValue;
    const oldValueDate = getDateOrNull_(oldValue);
    const checkboxInfo = !oldValueDate ? { range: cell, oldValue } : null;

    if (turnedOff) {
      if (sold) {
        cell.setValue(true);
        ss.toast(
          `Impossible de décocher "${headerLabel}" lorsqu'une vente est cochée.`,
          'Stock',
          5
        );
        return;
      }

      if (!ensureChronologyOrRevert_('dpub', oldValue, checkboxInfo) && wasCheckbox) {
        restoreCheckboxValidation_(cell, previousValidation);
      }
      return;
    }

    if (turnedOn) {
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      setChronoDateFromPrevious_('dpub');
      const info = { range: cell, oldValue };
      if (!ensureChronologyOrRevert_('dpub', oldValue, info)) {
        restoreCheckboxValidation_(cell, previousValidation);
        return;
      }
      propagateForwardFrom_('dpub');
      return;
    }

    if (value === '' || value === null) {
      if (sold) {
        setCellToFallback_(C_DPUB, oldValue);
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        ss.toast(
          `Impossible de vider "${headerLabel}" lorsqu'une vente est cochée.`,
          'Stock',
          5
        );
        return;
      }

      if (!ensureChronologyOrRevert_('dpub', oldValue, checkboxInfo)) {
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }

      restoreCheckboxValidation_(cell, previousValidation);
      cell.clearContent();
      return;
    }

    const parsedValue = getDateOrNull_(value);
    if (parsedValue) {
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      cell.setValue(parsedValue);
      cell.setNumberFormat('dd/MM/yyyy');
      if (!ensureChronologyOrRevert_('dpub', oldValue, checkboxInfo)) {
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }
      propagateForwardFrom_('dpub');
      return;
    }

    setCellToFallback_(C_DPUB, oldValue);
    if (oldValueDate) {
      cell.clearDataValidations();
      cell.setNumberFormat('dd/MM/yyyy');
    } else if (wasCheckbox) {
      restoreCheckboxValidation_(cell, previousValidation);
    }
    return;
  }

  if (!isCombinedMisCell && C_MIS && C_DMIS && c === C_MIS) {
    const legacyHeaderLabel = stockHeaders[C_MIS - 1] || HEADERS.STOCK.MIS_EN_LIGNE_ALT;
    if (turnedOff) {
      if (C_PUB && isStatusActiveValue_(sh.getRange(r, C_PUB).getValue())) {
        sh.getRange(r, C_MIS).setValue(true);
        ss.toast(
          `Impossible de décocher "${legacyHeaderLabel}" tant que "${HEADERS.STOCK.PUBLIE}" est coché.`,
          'Stock',
          5
        );
        return;
      }
      if (!restorePreviousCellValue_(sh, r, C_DMIS) && CLEAR_ON_UNCHECK) {
        sh.getRange(r, C_DMIS).clearContent();
      }
      return;
    }

    if (turnedOn) {
      const cell = sh.getRange(r, C_DMIS);
      storePreviousCellValue_(sh, r, C_DMIS, cell.getValue());
      setChronoDateFromPrevious_('dmis');
      const checkboxInfo = { range: sh.getRange(r, C_MIS), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dmis', null, checkboxInfo)) {
        return;
      }
      propagateForwardFrom_('dmis');
      return;
    }

    return;
  }

  // 2) PUBLIÉ → horodate
  if (!isCombinedPubCell && C_PUB && C_DPUB && c === C_PUB) {
    if (turnedOff) {
      const vendu = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
      if (vendu) {
        sh.getRange(r, C_PUB).setValue(true);
        ss.toast(
          `Impossible de décocher "${HEADERS.STOCK.PUBLIE}" lorsqu'une vente est cochée.`,
          'Stock',
          5
        );
        return;
      }

      if (!restorePreviousCellValue_(sh, r, C_DPUB) && CLEAR_ON_UNCHECK) {
        sh.getRange(r, C_DPUB).clearContent();
      }
      return;
    }

    if (turnedOn) {
      const cell = sh.getRange(r, C_DPUB);
      storePreviousCellValue_(sh, r, C_DPUB, cell.getValue());
      setChronoDateFromPrevious_('dpub');
      const checkboxInfo = { range: sh.getRange(r, C_PUB), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dpub', null, checkboxInfo)) {
        return;
      }
      propagateForwardFrom_('dpub');
      return;
    }

    return;
  }

  // 3) VENDU → horodatage + alerte prix, déplacement seulement via "Valider la saisie"
  if (isCombinedVenduCell && C_VENDU && c === C_VENDU) {
    logDebug_('handleStock.vendu.combined', {
      ligne: r,
      colonne: c,
      C_VENDU,
      C_DVENTE,
      valeur: e.value,
      ancienneValeur: e.oldValue
    });
    ss.toast(
      `VENDU (combiné) • ligne ${r} • valeur ${e.value || '—'} • col VENDU ${C_VENDU} / DATE ${C_DVENTE}`,
      'Diagnostic Stock',
      5
    );
    const cell = sh.getRange(r, C_VENDU);
    const previousValidation = cell.getDataValidation();
    const wasCheckbox = isCheckboxValidation_(previousValidation);
    const value = cell.getValue();
    const oldValue = e.oldValue;
    const oldValueDate = getDateOrNull_(oldValue);
    const checkboxInfo = !oldValueDate ? { range: cell, oldValue } : null;

    if (turnedOn) {
      if (C_PRIX) {
        const priceCell = sh.getRange(r, C_PRIX);
        storePreviousCellValue_(sh, r, C_PRIX, priceCell.getValue());
      }

      if (wasCheckbox) {
        cell.clearDataValidations();
      }

      setChronoDateFromPrevious_('dvente');
      const info = { range: cell, oldValue };
      if (!ensureChronologyOrRevert_('dvente', oldValue, info)) {
        if (C_PRIX) {
          restorePreviousCellValue_(sh, r, C_PRIX);
        }
        restoreCheckboxValidation_(cell, previousValidation);
        return;
      }

      propagateForwardFrom_('dvente');
      const fallbackValue = oldValueDate ? new Date(oldValueDate.getTime()) : '';
      storePreviousCellValue_(sh, r, C_DVENTE, fallbackValue);
      ensureValidPriceOrWarn_(sh, r, C_PRIX);
      return;
    }

    if (turnedOff) {
      if (!restorePreviousCellValue_(sh, r, C_DVENTE)) {
        cell.clearContent();
      }

      if (C_PRIX) {
        restorePreviousCellValue_(sh, r, C_PRIX);
        clearPriceAlertIfAny_(sh, r, C_PRIX);
        const priceCell = sh.getRange(r, C_PRIX);
        priceCell.clearContent();
      } else {
        clearPriceAlertIfAny_(sh, r, C_PRIX);
      }

      if (C_VALIDE) {
        const valCell = sh.getRange(r, C_VALIDE);
        valCell.clearDataValidations();
        valCell.clearContent();
      }

      restoreCheckboxValidation_(cell, previousValidation);
      return;
    }

    if (value === '' || value === null) {
      if (!ensureChronologyOrRevert_('dvente', oldValue, checkboxInfo)) {
        if (C_PRIX) {
          restorePreviousCellValue_(sh, r, C_PRIX);
        }
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }

      restorePreviousCellValue_(sh, r, C_DVENTE);
      if (C_PRIX) {
        restorePreviousCellValue_(sh, r, C_PRIX);
        clearPriceAlertIfAny_(sh, r, C_PRIX);
        const priceCell = sh.getRange(r, C_PRIX);
        priceCell.clearContent();
      } else {
        clearPriceAlertIfAny_(sh, r, C_PRIX);
      }

      if (C_VALIDE) {
        const valCell = sh.getRange(r, C_VALIDE);
        valCell.clearDataValidations();
        valCell.clearContent();
      }

      restoreCheckboxValidation_(cell, previousValidation);
      cell.clearContent();
      return;
    }

    const parsedValue = getDateOrNull_(value);
    if (parsedValue) {
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      cell.setValue(parsedValue);
      cell.setNumberFormat('dd/MM/yyyy');
      if (!ensureChronologyOrRevert_('dvente', oldValue, checkboxInfo)) {
        if (C_PRIX) {
          restorePreviousCellValue_(sh, r, C_PRIX);
        }
        if (oldValueDate) {
          cell.clearDataValidations();
          cell.setNumberFormat('dd/MM/yyyy');
        } else if (wasCheckbox) {
          restoreCheckboxValidation_(cell, previousValidation);
        }
        return;
      }

      propagateForwardFrom_('dvente');
      const fallbackValue = oldValueDate ? new Date(oldValueDate.getTime()) : '';
      storePreviousCellValue_(sh, r, C_DVENTE, fallbackValue);
      ensureValidPriceOrWarn_(sh, r, C_PRIX);
      return;
    }

    setCellToFallback_(C_DVENTE, oldValue);
    if (C_PRIX) {
      ensureValidPriceOrWarn_(sh, r, C_PRIX);
    }
    if (oldValueDate) {
      cell.clearDataValidations();
      cell.setNumberFormat('dd/MM/yyyy');
    } else if (wasCheckbox) {
      restoreCheckboxValidation_(cell, previousValidation);
    }
    return;
  }

  if (!isCombinedVenduCell && C_VENDU && c === C_VENDU) {
    logDebug_('handleStock.vendu.legacy', {
      ligne: r,
      colonne: c,
      C_VENDU,
      C_DVENTE,
      valeur: e.value,
      ancienneValeur: e.oldValue
    });
    ss.toast(
      `VENDU (legacy) • ligne ${r} • valeur ${e.value || '—'} • col VENDU ${C_VENDU} / DATE ${C_DVENTE}`,
      'Diagnostic Stock',
      5
    );
    const dv = sh.getRange(r, C_DVENTE);

    if (turnedOn) {
      storePreviousCellValue_(sh, r, C_DVENTE, dv.getValue());
      if (C_PRIX) {
        const priceCell = sh.getRange(r, C_PRIX);
        storePreviousCellValue_(sh, r, C_PRIX, priceCell.getValue());
      }

      setChronoDateFromPrevious_('dvente');
      const checkboxInfo = { range: sh.getRange(r, C_VENDU), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dvente', null, checkboxInfo)) {
        if (C_PRIX) {
          restorePreviousCellValue_(sh, r, C_PRIX);
        }
        return;
      }
      propagateForwardFrom_('dvente');
      ensureValidPriceOrWarn_(sh, r, C_PRIX);
      return;
    }

    if (turnedOff) {
      if (!restorePreviousCellValue_(sh, r, C_DVENTE)) {
        dv.clearContent();
      }

      if (C_PRIX) {
        restorePreviousCellValue_(sh, r, C_PRIX);
        clearPriceAlertIfAny_(sh, r, C_PRIX);
        const priceCell = sh.getRange(r, C_PRIX);
        priceCell.clearContent();
      } else {
        clearPriceAlertIfAny_(sh, r, C_PRIX);
      }

      if (C_VALIDE) {
        const valCell = sh.getRange(r, C_VALIDE);
        valCell.clearDataValidations();
        valCell.clearContent();
      }
      return;
    }

    clearPriceAlertIfAny_(sh, r, C_PRIX);
    return;
  }

  // 5) Saisie directe d’une DATE DE VENTE → juste contrôle prix si VENDU coché
  if (!isCombinedVenduCell && c === C_DVENTE) {
    const val = sh.getRange(r, C_DVENTE).getValue();
    const isDate = val instanceof Date && !isNaN(val.getTime());
    if (!isDate) return;

    const vendu = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    if (!vendu) return;

    ensureValidPriceOrWarn_(sh, r, C_PRIX);
    propagateForwardFrom_('dpub');
    return;
  }

  // 6) "Valider la saisie" → déplacement vers Ventes si tout est OK
  if (C_VALIDE && c === C_VALIDE) {
    if (!turnedOn) {
      return;
    }

    const chronoCheck = enforceChronologicalDates_(sh, r, chronoCols, { requireAllDates: true });
    if (!chronoCheck.ok) {
      revertCheckbox_(e.range, e.oldValue);
      ss.toast(chronoCheck.message || 'Ordre chronologique des dates invalide.', 'Stock', 6);
      return;
    }

    const vendu = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    if (!vendu) {
      return;
    }

    if (!ensureValidPriceOrWarn_(sh, r, C_PRIX)) return;

    if (!C_TAILLE) {
      revertCheckbox_(e.range, e.oldValue);
      ss.toast('Colonne taille introuvable ("TAILLE" / "TAILLE DU COLIS").', 'Stock', 6);
      return;
    }

    const tailleValue = String(sh.getRange(r, C_TAILLE).getDisplayValue() || '').trim();
    if (!tailleValue) {
      revertCheckbox_(e.range, e.oldValue);
      ss.toast(`Indique la colonne ${tailleHeaderLabel} avant de valider.`, 'Stock', 6);
      return;
    }

    const lotValue = C_LOT ? String(sh.getRange(r, C_LOT).getDisplayValue() || '').trim() : '';
    const shippingLookup = buildShippingFeeLookup_(ss);
    if (!shippingLookup) {
      revertCheckbox_(e.range, e.oldValue);
      ss.toast('Impossible de calculer les frais de colissage : configure la feuille "Frais".', 'Stock', 6);
      return;
    }

    const fraisColis = shippingLookup(tailleValue, lotValue);
    if (!Number.isFinite(fraisColis)) {
      revertCheckbox_(e.range, e.oldValue);
      const lotMessage = lotValue ? ` / lot ${lotValue}` : '';
      ss.toast(`Frais de colissage introuvables pour la taille ${tailleValue}${lotMessage}.`, 'Stock', 6);
      return;
    }

    const perItemFee = computePerItemShippingFee_(fraisColis, lotValue);

    const valDate = sh.getRange(r, C_DVENTE).getValue();
    if (!(valDate instanceof Date) || isNaN(valDate.getTime())) return;

    const baseToDmsMap = buildBaseToStockDate_(ss);
    exportVente_(
      null,
      r,
      C_ID,
      C_LABEL,
      C_SKU,
      C_PRIX,
      C_DVENTE,
      C_STAMPV,
      baseToDmsMap,
      { shipping: { size: tailleValue, lot: lotValue, fee: perItemFee } }
    );
    return;
    }
  }

function applyBoostBadge_(label) {
  if (label === null || label === undefined) return label;

  const text = String(label);
  const trimmed = text.replace(/^\s+/, '');

  if (trimmed === '') return label;

  const BOOST_BADGE = '🚀';
  let rest = trimmed;

  if (rest.startsWith(BOOST_BADGE)) {
    rest = rest.slice(BOOST_BADGE.length);
    rest = rest.replace(/^\s+/, '');
  }

  return rest ? `${BOOST_BADGE} ${rest}` : BOOST_BADGE;
}

function bumpRepublishBadge_(label) {
  if (label === null || label === undefined) return label;

  const text = String(label);
  const trimmed = text.replace(/^\s+/, ''); // trimStart compatible

  if (trimmed === '') return label;

  const BADGES = ['🟩', '🟨', '🟧', '🟥'];

  // Détecter badge existant en début
  let currentIndex = -1;
  for (let i = 0; i < BADGES.length; i++) {
    if (trimmed.startsWith(BADGES[i])) {
      currentIndex = i;
      break;
    }
  }

  // Retirer badge existant + espaces éventuels après
  let rest = trimmed;
  if (currentIndex !== -1) {
    rest = trimmed.slice(BADGES[currentIndex].length);
    rest = rest.replace(/^\s+/, '');
  }

  // Calcul du prochain badge (bloqué sur 🟥)
  const nextIndex = currentIndex === -1 ? 0 : Math.min(currentIndex + 1, BADGES.length - 1);
  const nextBadge = BADGES[nextIndex];

  return rest ? `${nextBadge} ${rest}` : nextBadge;
}


function bulkSetStockStatusMisEnLigne() {
  bulkSetStockStatusForKey_({
    chronologyKey: 'dmis',
    statusLabel: HEADERS.STOCK.MIS_EN_LIGNE,
    getSelectionColumn: ctx => ctx.columns.mis,
    getCheckboxColumn: ctx => ctx.combinedFlags.dmis ? 0 : ctx.columns.mis,
    getDateColumn: ctx => ctx.columns.dmis,
    isCombined: ctx => ctx.combinedFlags.dmis
  });
}

function bulkSetStockStatusPublie() {
  bulkSetStockStatusForKey_({
    chronologyKey: 'dpub',
    statusLabel: HEADERS.STOCK.PUBLIE,
    getSelectionColumn: ctx => ctx.columns.pub,
    getCheckboxColumn: ctx => ctx.combinedFlags.dpub ? 0 : ctx.columns.pub,
    getDateColumn: ctx => ctx.columns.dpub,
    isCombined: ctx => ctx.combinedFlags.dpub
  });
}

function bulkSetStockStatusVendu() {
  bulkSetStockStatusForKey_({
    chronologyKey: 'dvente',
    statusLabel: HEADERS.STOCK.VENDU,
    getSelectionColumn: ctx => ctx.columns.vendu,
    getCheckboxColumn: ctx => ctx.combinedFlags.dvente ? 0 : ctx.columns.vendu,
    getDateColumn: ctx => ctx.columns.dvente,
    isCombined: ctx => ctx.combinedFlags.dvente
  });
}

function bulkSetStockStatusForKey_(config) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== 'Stock') {
    ss.toast('Ouvre la feuille "Stock" et sélectionne la colonne souhaitée.', 'Stock', 6);
    return;
  }

  const context = getStockStatusColumnContext_(sh);
  if (!context) {
    ss.toast('Impossible d\'identifier les colonnes de statut du stock.', 'Stock', 6);
    return;
  }

  const selectionColumn = config.getSelectionColumn(context);
  const dateColumn = config.getDateColumn(context);
  if (!selectionColumn || !dateColumn) {
    ss.toast(`Colonne introuvable pour "${config.statusLabel}".`, 'Stock', 6);
    return;
  }

  const checkboxColumn = typeof config.getCheckboxColumn === 'function'
    ? config.getCheckboxColumn(context)
    : selectionColumn;
  const isCombined = typeof config.isCombined === 'function'
    ? config.isCombined(context)
    : false;

  const selection = ss.getSelection();
  const rangeList = selection ? selection.getActiveRangeList() : null;
  let ranges = [];
  if (rangeList) {
    ranges = rangeList.getRanges();
  } else if (selection && selection.getActiveRange()) {
    ranges = [selection.getActiveRange()];
  }

  if (!ranges.length) {
    ss.toast(`Sélectionne les cases "${config.statusLabel}" à mettre à jour.`, 'Stock', 5);
    return;
  }

  const HEADER_ROW = 6;
  const DATA_START_ROW = HEADER_ROW + 1;

  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (!range) continue;
    if (range.getRow() < DATA_START_ROW) {
      ss.toast('La sélection doit être sur les lignes de données (à partir de la ligne 7).', 'Stock', 6);
      return;
    }
  }


  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (!range) continue;
    if (range.getNumColumns() !== 1 || range.getColumn() !== selectionColumn) {
      ss.toast(`La sélection doit être limitée à la colonne "${config.statusLabel}".`, 'Stock', 6);
      return;
    }
  }

  const chronology = context.chronology;
  const combinedFlags = context.combinedFlags || {};
  const lastColumn = sh.getLastColumn();
  let totalUpdated = 0;

  ranges.forEach(range => {
    const numRows = range.getNumRows();
    const startRow = range.getRow();
    if (numRows <= 0) {
      return;
    }

    const blockValues = sh.getRange(startRow, 1, numRows, lastColumn).getValues();
    const columnBuffers = new Map();
    const blockValidationClears = [];
    const blockFormats = [];

    function applyValue(column, rowOffset, value) {
      if (!column) return;
      let buffer = columnBuffers.get(column);
      if (!buffer) {
        buffer = [];
        for (let i = 0; i < numRows; i++) {
          buffer.push([blockValues[i][column - 1]]);
        }
        columnBuffers.set(column, buffer);
      }
      buffer[rowOffset][0] = value;
      blockValues[rowOffset][column - 1] = value;
    }

    let blockUpdated = false;

    const HEADER_ROW = 6;
    const DATA_START_ROW = HEADER_ROW + 1;

    for (let offset = 0; offset < numRows; offset++) {
      const rowIndex = startRow + offset;
      if (rowIndex < DATA_START_ROW) {
        continue; // ignore en-tête + lignes au-dessus
      }

      const rowValues = blockValues[offset];
      const rowDates = {};
      for (let i = 0; i < chronology.length; i++) {
        const entry = chronology[i];
        rowDates[entry.key] = entry.column
          ? getDateOrNull_(rowValues[entry.column - 1])
          : null;
      }

      if (rowDates[config.chronologyKey]) {
        continue;
      }

      const newDate = computeChronologyDateFromRow_(rowDates, chronology, config.chronologyKey);
      if (!newDate) {
        continue;
      }

      if (isCombined) {
        blockValidationClears.push({ rowIndex, column: dateColumn });
        applyValue(dateColumn, offset, newDate);
      } else {
        if (checkboxColumn) {
          applyValue(checkboxColumn, offset, true);
          rowValues[checkboxColumn - 1] = true;
        }
        applyValue(dateColumn, offset, newDate);
      }

      rowDates[config.chronologyKey] = newDate;
      rowValues[dateColumn - 1] = newDate;
      blockFormats.push({ rowIndex, column: dateColumn });

      propagateChronologyForwardForRow_(
        rowDates,
        chronology,
        config.chronologyKey,
        rowValues,
        offset,
        rowIndex,
        applyValue,
        blockFormats,
        blockValidationClears,
        combinedFlags
      );

      blockUpdated = true;
      totalUpdated++;
    }

    if (!blockUpdated) {
      return;
    }

    if (blockValidationClears.length) {
      blockValidationClears.forEach(target => {
        sh.getRange(target.rowIndex, target.column).clearDataValidations();
      });
    }

    columnBuffers.forEach((values, column) => {
      sh.getRange(startRow, column, numRows, 1).setValues(values);
    });

    if (blockFormats.length) {
      blockFormats.forEach(target => {
        sh.getRange(target.rowIndex, target.column).setNumberFormat('dd/MM/yyyy');
      });
    }
  });

  if (totalUpdated > 0) {
    ss.toast(`${totalUpdated} ligne(s) mises à jour pour "${config.statusLabel}".`, 'Stock', 5);
  } else {
    ss.toast(`Aucune nouvelle date à renseigner pour "${config.statusLabel}".`, 'Stock', 5);
  }
}

function getStockStatusColumnContext_(sheet) {
  if (!sheet) {
    return null;
  }

  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) {
    return null;
  }

  const HEADER_ROW = 6;
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];

  const resolver = makeHeaderResolver_(headers);
  if (!resolver) {
    return null;
  }

  const colExact = resolver.colExact.bind(resolver);
  const combinedMisCol = resolveCombinedMisEnLigneColumn_(resolver);
  const legacyMisCols = combinedMisCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyMisEnLigneColumn_(resolver);
  const useCombinedMisCol = !!combinedMisCol;

  const combinedPubCol = resolveCombinedPublicationColumn_(resolver);
  const legacyPubCols = combinedPubCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyPublicationColumns_(resolver);
  const useCombinedPubCol = !!combinedPubCol;

  const combinedVenduCol = resolveCombinedVenduColumn_(resolver);
  const legacyVenduCols = combinedVenduCol ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyVenduColumns_(resolver);
  const useCombinedVenduCol = !!combinedVenduCol;

  const C_MIS = useCombinedMisCol ? combinedMisCol : legacyMisCols.checkboxCol;
  const C_DMIS = useCombinedMisCol ? combinedMisCol : legacyMisCols.dateCol;
  const C_PUB = useCombinedPubCol ? combinedPubCol : legacyPubCols.checkboxCol;
  const C_DPUB = useCombinedPubCol ? combinedPubCol : legacyPubCols.dateCol;
  const C_VENDU = useCombinedVenduCol ? combinedVenduCol : legacyVenduCols.checkboxCol;
  let C_DVENTE = useCombinedVenduCol ? combinedVenduCol : legacyVenduCols.dateCol;
  if (!C_DVENTE) {
    C_DVENTE = colExact(HEADERS.STOCK.DATE_VENTE_ALT) || 0;
  }
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK) || 0;

  const chronology = [
    { key: 'dms', column: C_DMS, statusCol: 0 },
    { key: 'dmis', column: C_DMIS, statusCol: C_MIS },
    { key: 'dpub', column: C_DPUB, statusCol: C_PUB },
    { key: 'dvente', column: C_DVENTE, statusCol: C_VENDU }
  ];

  const combinedFlags = {
    dmis: useCombinedMisCol && C_MIS && C_DMIS && C_MIS === C_DMIS,
    dpub: useCombinedPubCol && C_PUB && C_DPUB && C_PUB === C_DPUB,
    dvente: useCombinedVenduCol && C_VENDU && C_DVENTE && C_VENDU === C_DVENTE
  };

  return {
    headers,
    resolver,
    columns: {
      mis: C_MIS,
      dmis: C_DMIS,
      pub: C_PUB,
      dpub: C_DPUB,
      vendu: C_VENDU,
      dvente: C_DVENTE,
      dms: C_DMS
    },
    chronology,
    combinedFlags
  };
}

function computeChronologyDateFromRow_(rowDates, chronology, key) {
  const index = chronology.findIndex(entry => entry.key === key);
  if (index < 0) {
    return new Date();
  }

  const now = new Date();
  if (index === 0) {
    return now;
  }

  const previous = chronology[index - 1];
  if (previous) {
    const baseDate = rowDates[previous.key];
    if (baseDate instanceof Date && !isNaN(baseDate.getTime())) {
      return baseDate.getTime() > now.getTime() ? new Date(baseDate.getTime()) : now;
    }
  }

  return now;
}

function propagateChronologyForwardForRow_(
  rowDates,
  chronology,
  startKey,
  rowValues,
  rowOffset,
  rowIndex,
  applyValue,
  formatTargets,
  validationClears,
  combinedFlags
) {
  const startIndex = chronology.findIndex(entry => entry.key === startKey);
  if (startIndex < 0) {
    return;
  }

  for (let i = startIndex + 1; i < chronology.length; i++) {
    const entry = chronology[i];
    if (!entry.column || !entry.statusCol) {
      continue;
    }

    const statusValue = rowValues[entry.statusCol - 1];
    if (!isStatusActiveValue_(statusValue)) {
      continue;
    }

    if (rowDates[entry.key]) {
      continue;
    }

    const nextDate = computeChronologyDateFromRow_(rowDates, chronology, entry.key);
    if (!nextDate) {
      continue;
    }

    rowDates[entry.key] = nextDate;
    rowValues[entry.column - 1] = nextDate;
    applyValue(entry.column, rowOffset, nextDate);
    formatTargets.push({ rowIndex, column: entry.column });
    if (combinedFlags[entry.key]) {
      validationClears.push({ rowIndex, column: entry.column });
    }
  }
}
