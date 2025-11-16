function handleStock(e) {
  const sh = e.source.getActiveSheet();
  const ss = e.source;
  const turnedOn  = (e.value === "TRUE") || (e.value === true);
  const turnedOff = (e.value === "FALSE") || (e.value === false);
  const CLEAR_ON_UNCHECK = false;

  const stockHeaders = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(stockHeaders);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID      = colExact(HEADERS.STOCK.ID);
  const C_LABEL   = colExact(HEADERS.STOCK.LIBELLE)
    || colExact(HEADERS.STOCK.LIBELLE_ALT)
    || colExact(HEADERS.STOCK.ARTICLE)
    || colExact(HEADERS.STOCK.ARTICLE_ALT)
    || colWhere(h => h.includes('libell'))
    || colWhere(h => h.includes('article'))
    || 2;
  const C_SKU     = colExact(HEADERS.STOCK.SKU) || colExact(HEADERS.STOCK.REFERENCE); // B/C
  const C_PRIX    = colExact(HEADERS.STOCK.PRIX_VENTE)
    || colWhere(h => h.includes("prix") && h.includes("vente"));
  const C_TAILLE  = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const tailleHeaderLabel = getHeaderLabel_(resolver, C_TAILLE, HEADERS.STOCK.TAILLE);
  const C_LOT     = colExact(HEADERS.STOCK.LOT)
    || colExact(HEADERS.STOCK.LOT_ALT)
    || colWhere(h => h.includes('lot'));
  const C_DMS     = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK);

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

  const c = e.range.getColumn(), r = e.range.getRow();

  try {
    applyStockListValidation_(sh, C_TAILLE, ['Petit', 'Moyen', 'Grand']);
    applyStockListValidation_(sh, C_LOT, ['2', '3', '4', '5']);
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

  function setChronoDateFromPrevious_(key) {
    const entry = chronology.find(item => item.key === key);
    if (!entry || !entry.column) {
      return null;
    }

    const nextDate = new Date();
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

  

  function hasValidPrice_() {
    if (!C_PRIX) return false;
    const priceCell = sh.getRange(r, C_PRIX);
    const value = priceCell.getValue();
    const display = priceCell.getDisplayValue();
    return (typeof value === 'number') && !isNaN(value) && value > 0 && (!display || display.indexOf('⚠️') !== 0);
  }

  function hasShippingSize_() {
    if (!C_TAILLE) return false;
    const val = String(sh.getRange(r, C_TAILLE).getDisplayValue() || '').trim();
    return !!val;
  }

  function refreshValiderCheckbox_() {
    if (!C_VALIDE) return;
    const valCell = sh.getRange(r, C_VALIDE);
    const validation = valCell.getDataValidation();
    const isCheckbox = isCheckboxValidation_(validation);
    const allowInvalid = validation && typeof validation.getAllowInvalid === 'function' && validation.getAllowInvalid();
    const shouldEnable = hasValidPrice_() && hasShippingSize_();

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
      return;
    }

    valCell.clearDataValidations();
    valCell.clearContent();
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

  // 0bis) Modification du PRIX DE VENTE
  if (C_PRIX && c === C_PRIX) {
    const vendu = C_VENDU ? isStatusActiveValue_(sh.getRange(r, C_VENDU).getValue()) : false;
    const priceCell = sh.getRange(r, C_PRIX);
    const priceValue = priceCell.getValue();
    const priceDisplay = priceCell.getDisplayValue();

    refreshValiderCheckbox_();

    if (!vendu) {
      // Rien n'est coché → pas d'alerte.
      clearPriceAlertIfAny_(sh, r, C_PRIX);
      return;
    }

    // La colonne VENDU est cochée → contrôle du prix
    ensureValidPriceOrWarn_(sh, r, C_PRIX);
    return;
  }

  if (C_TAILLE && c === C_TAILLE) {
    refreshValiderCheckbox_();
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

    if (oldValueDate && (turnedOn || turnedOff)) {
      cell.setValue(oldValueDate);
      cell.setNumberFormat('dd/MM/yyyy');
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      return;
    }

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
      restorePreviousCellValue_(sh, r, C_DMIS);
      restoreCheckboxValidation_(cell, previousValidation);
      return;
    }

    if (turnedOn) {
      const fallbackValue = oldValueDate ? new Date(oldValueDate.getTime()) : '';
      storePreviousCellValue_(sh, r, C_DMIS, fallbackValue);
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

    if (oldValueDate && (turnedOn || turnedOff)) {
      cell.setValue(oldValueDate);
      cell.setNumberFormat('dd/MM/yyyy');
      if (wasCheckbox) {
        cell.clearDataValidations();
      }
      return;
    }

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
      restorePreviousCellValue_(sh, r, C_DPUB);
      restoreCheckboxValidation_(cell, previousValidation);
      return;
    }

    if (turnedOn) {
      const fallbackValue = oldValueDate ? new Date(oldValueDate.getTime()) : '';
      storePreviousCellValue_(sh, r, C_DPUB, fallbackValue);
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
      refreshValiderCheckbox_();
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

      refreshValiderCheckbox_();

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
      refreshValiderCheckbox_();
      return;
    }

    setCellToFallback_(C_DVENTE, oldValue);
    if (C_PRIX) {
      ensureValidPriceOrWarn_(sh, r, C_PRIX);
    }
    refreshValiderCheckbox_();
    if (oldValueDate) {
      cell.clearDataValidations();
      cell.setNumberFormat('dd/MM/yyyy');
    } else if (wasCheckbox) {
      restoreCheckboxValidation_(cell, previousValidation);
    }
    return;
  }

  if (!isCombinedVenduCell && C_VENDU && c === C_VENDU) {
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
      refreshValiderCheckbox_();
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

      refreshValiderCheckbox_();
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
    e.range.setValue(new Date());
    e.range.setNumberFormat('dd/MM/yyyy');

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
  } finally {
    applySkuPaletteFormatting_(sh, C_SKU, C_LABEL);
  }
}

function applyStockListValidation_(sheet, col, values) {
  if (!sheet || !col || !Array.isArray(values) || !values.length) return;
  const lastRow = sheet.getMaxRows();
  if (lastRow < 2) return;
  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
}

function ensureStockSizeDropdowns() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) return;

  const stockSheet = ss.getSheetByName('Stock');
  if (!stockSheet) return;

  const stockHeaders = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(stockHeaders);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const sizeCol = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);

  applyStockListValidation_(stockSheet, sizeCol, ['Petit', 'Moyen', 'Grand']);
}
