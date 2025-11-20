function handleVentesReturn(e) {
  const ss = e && e.source;
  const sh = ss && ss.getActiveSheet();
  const range = e && e.range;
  if (!ss || !sh || sh.getName() !== 'Ventes' || !range) return;

  const row = range.getRow();
  if (row <= 1) return;

  const lastColumn = sh.getLastColumn();
  if (!lastColumn) return;

  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colReturn = resolver.colExact(HEADERS.VENTES.RETOUR)
    || resolver.colWhere(header => header.includes('retour'));
  if (!colReturn || range.getColumn() !== colReturn) return;

  const newValue = e.value !== undefined ? e.value : range.getValue();
  if (!isReturnFlagActive_(newValue)) {
    return;
  }

  const rowValues = sh.getRange(row, 1, 1, lastColumn).getValues()[0];
  const sale = buildSaleRecordFromVentesRow_(rowValues, resolver, row);
  if (!sale.dateVente) {
    clearReturnFlag_(range);
    ss.toast('Impossible de traiter le retour : date de vente introuvable.', 'Ventes', 8);
    return;
  }

  const labelForToast = sale.libelle || sale.sku || sale.id || `ligne ${row}`;

  try {
    const ledgerResult = copySaleToMonthlySheet_(
      ss,
      Object.assign({}, sale, { retour: true }),
      { updateExisting: true }
    );

    if (!ledgerResult || !ledgerResult.removed) {
      clearReturnFlag_(range);
      ss.toast(`Impossible de retirer ${labelForToast} du journal mensuel.`, 'Ventes', 8);
      return;
    }

    const returnedAt = new Date();
    const stockResult = restoreSaleToStock_(ss, Object.assign({}, sale, { returnedAt }));
    if (!stockResult.success) {
      copySaleToMonthlySheet_(ss, sale, { updateExisting: true });
      clearReturnFlag_(range);
      ss.toast(stockResult.message || `Impossible de recréer ${labelForToast} dans Stock.`, 'Ventes', 8);
      return;
    }

    sh.deleteRow(row);
    ss.toast(`Retour traité pour ${labelForToast}.`, 'Ventes', 5);
  } catch (err) {
    clearReturnFlag_(range);
    const message = err && err.message ? err.message : err;
    ss.toast(`Erreur lors du retour : ${message}`, 'Ventes', 10);
    if (typeof console !== 'undefined' && console.error) {
      console.error(err);
    }
  }
}

function isReturnFlagActive_(value) {
  if (value === true) return true;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === 'oui' || normalized === 'retour') {
      return true;
    }
  }
  if (value === 'TRUE') return true;
  return false;
}

function clearReturnFlag_(range) {
  if (!range) return;
  range.clearContent();
}

function isComptabiliseFlagActive_(value) {
  if (value === true) return true;
  if (value === 'TRUE') return true;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'oui' || normalized === 'true' || normalized === 'comptabilisé' || normalized === 'comptabilise') {
      return true;
    }
  }
  return false;
}

function buildSaleRecordFromVentesRow_(rowValues, resolver, rowNumber) {
  if (!rowValues || !resolver) return {};
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const COL_ID = colExact(HEADERS.VENTES.ID);
  const COL_SKU = colExact(HEADERS.VENTES.SKU);
  const COL_ARTICLE = colExact(HEADERS.VENTES.ARTICLE)
    || colExact(HEADERS.VENTES.ARTICLE_ALT)
    || colWhere(h => h.includes('article'));
  const COL_PRIX = colExact(HEADERS.VENTES.PRIX_VENTE)
    || colExact(HEADERS.VENTES.PRIX_VENTE_ALT)
    || colWhere(h => h.includes('prix'));
  const COL_DATE = colExact(HEADERS.VENTES.DATE_VENTE)
    || colWhere(h => h.includes('date') && h.includes('vente'));
  const COL_LOT = colExact(HEADERS.VENTES.LOT) || colWhere(h => h.includes('lot'));
  const COL_TAILLE = colExact(HEADERS.VENTES.TAILLE_COLIS)
    || colExact(HEADERS.VENTES.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const COL_COMPTA = colExact(HEADERS.VENTES.COMPTABILISE)
    || colWhere(h => h.toLowerCase().includes('compt'));

  const id = COL_ID ? rowValues[COL_ID - 1] : '';
  const sku = COL_SKU ? rowValues[COL_SKU - 1] : '';
  const libelle = COL_ARTICLE ? rowValues[COL_ARTICLE - 1] : '';
  const prixVenteRaw = COL_PRIX ? rowValues[COL_PRIX - 1] : '';
  const dateRaw = COL_DATE ? rowValues[COL_DATE - 1] : null;
  const lot = COL_LOT ? rowValues[COL_LOT - 1] : '';
  const taille = COL_TAILLE ? rowValues[COL_TAILLE - 1] : '';
  const comptabilise = COL_COMPTA ? isComptabiliseFlagActive_(rowValues[COL_COMPTA - 1]) : false;

  const prixVente = prixVenteRaw === '' || prixVenteRaw === null || prixVenteRaw === undefined
    ? ''
    : valueToNumber_(prixVenteRaw);

  return {
    id: id !== undefined && id !== null ? id : '',
    sku: sku !== undefined && sku !== null ? sku : '',
    libelle: libelle !== undefined && libelle !== null ? libelle : '',
    dateVente: getDateOrNull_(dateRaw),
    prixVente: Number.isFinite(prixVente) ? prixVente : '',
    prixAchat: '',
    margeBrute: '',
    coeffMarge: '',
    nbPieces: 1,
    lot,
    taille,
    comptabilise,
    sourceRowNumber: Number.isFinite(rowNumber) ? rowNumber : undefined
  };
}

function restoreSaleToStock_(ss, sale) {
  const stock = ss && ss.getSheetByName('Stock');
  if (!stock) {
    return { success: false, message: 'Feuille "Stock" introuvable.' };
  }

  const lastColumn = stock.getLastColumn();
  if (!lastColumn) {
    return { success: false, message: 'La feuille "Stock" ne contient aucune colonne.' };
  }

  const headers = stock.getRange(1, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID = colExact(HEADERS.STOCK.ID);
  const C_LABEL = colExact(HEADERS.STOCK.LIBELLE)
    || colExact(HEADERS.STOCK.LIBELLE_ALT)
    || colExact(HEADERS.STOCK.ARTICLE)
    || colExact(HEADERS.STOCK.ARTICLE_ALT)
    || colWhere(h => h.includes('libell'))
    || colWhere(h => h.includes('article'));
  const C_SKU = colExact(HEADERS.STOCK.SKU)
    || colExact(HEADERS.STOCK.REFERENCE)
    || colWhere(h => h.includes('sku'));
  const C_PRIX = colExact(HEADERS.STOCK.PRIX_VENTE)
    || colWhere(h => h.includes('prix') && h.includes('vente'));
  const C_TAILLE = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const C_LOT = colExact(HEADERS.STOCK.LOT) || colWhere(h => h.includes('lot'));
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK) || colWhere(h => h.includes('mise en stock'));
  const C_STAMPV = colExact(HEADERS.STOCK.VENTE_EXPORTEE_LE) || colWhere(h => h.includes('exporte'));
  const C_VALIDE = colExact(HEADERS.STOCK.VALIDER_SAISIE)
    || colExact(HEADERS.STOCK.VALIDER_SAISIE_ALT)
    || colWhere(h => h.includes('valider'));

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
  const C_DVENTE = useCombinedVenduCol ? combinedVenduCol : legacyVenduCols.dateCol;

  const rowValues = Array(lastColumn).fill('');
  const returnedAt = sale && sale.returnedAt instanceof Date && !isNaN(sale.returnedAt)
    ? sale.returnedAt
    : new Date();

  function setValue(column, value) {
    if (!column) return;
    rowValues[column - 1] = value;
  }

  function resetStatus(column, dateColumn) {
    if (column) {
      rowValues[column - 1] = '';
    }
    if (dateColumn && dateColumn !== column) {
      rowValues[dateColumn - 1] = '';
    }
  }

  setValue(C_ID, sale && sale.id !== undefined && sale.id !== null ? sale.id : '');
  setValue(C_SKU, sale && sale.sku !== undefined && sale.sku !== null ? sale.sku : '');
  setValue(C_LABEL, sale && sale.libelle !== undefined && sale.libelle !== null ? sale.libelle : '');

  if (C_PRIX && sale) {
    const priceValue = valueToNumber_(sale.prixVente);
    if (Number.isFinite(priceValue)) {
      setValue(C_PRIX, priceValue);
    }
  }

  setValue(C_TAILLE, sale && sale.taille !== undefined && sale.taille !== null ? sale.taille : '');
  setValue(C_LOT, sale && sale.lot !== undefined && sale.lot !== null ? sale.lot : '');

  if (C_DMS) {
    setValue(C_DMS, returnedAt);
  }

  if (C_STAMPV) {
    setValue(C_STAMPV, '');
  }

  if (C_VALIDE) {
    setValue(C_VALIDE, '');
  }

  resetStatus(C_MIS, C_DMIS);
  resetStatus(C_PUB, C_DPUB);
  resetStatus(C_VENDU, C_DVENTE);

  if (C_DVENTE && C_DVENTE !== C_VENDU) {
    setValue(C_DVENTE, '');
  }

  const targetRow = Math.max(2, stock.getLastRow() + 1);
  stock.getRange(targetRow, 1, 1, lastColumn).setValues([rowValues]);

  return { success: true, row: targetRow };
}

function exportValidatedSales_() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getSheetByName('Stock');
  if (!stock) {
    ss.toast('Feuille "Stock" introuvable.', 'Stock', 6);
    return;
  }

  const lastRow = stock.getLastRow();
  const lastColumn = stock.getLastColumn();
  if (lastColumn < 1) return;

  const headers = stock.getRange(1, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID = colExact(HEADERS.STOCK.ID);
  const C_LABEL = colExact(HEADERS.STOCK.LIBELLE)
    || colExact(HEADERS.STOCK.LIBELLE_ALT)
    || colExact(HEADERS.STOCK.ARTICLE)
    || colExact(HEADERS.STOCK.ARTICLE_ALT)
    || colWhere(h => h.includes('libell'))
    || colWhere(h => h.includes('article'))
    || 2;
  const C_SKU = colExact(HEADERS.STOCK.SKU) || colExact(HEADERS.STOCK.REFERENCE);
  const C_PRIX = colExact(HEADERS.STOCK.PRIX_VENTE) || colWhere(h => h.includes('prix') && h.includes('vente'));
  const C_TAILLE = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const tailleHeaderLabel = getHeaderLabel_(resolver, C_TAILLE, HEADERS.STOCK.TAILLE);
  const C_LOT = colExact(HEADERS.STOCK.LOT) || colWhere(h => h.includes('lot'));
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK);
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
  if (!C_DVENTE) C_DVENTE = colExact(HEADERS.STOCK.DATE_VENTE_ALT) || 10;
  const C_STAMPV = colExact(HEADERS.STOCK.VENTE_EXPORTEE_LE);
  const C_VALIDE = colExact(HEADERS.STOCK.VALIDER_SAISIE)
    || colExact(HEADERS.STOCK.VALIDER_SAISIE_ALT)
    || colWhere(h => h.includes('valider'));

  if (!C_VALIDE) {
    ss.toast('Colonne "VALIDER" introuvable dans Stock.', 'Stock', 6);
    return;
  }

  const validationsRange = stock.getRange(2, C_VALIDE, Math.max(1, stock.getMaxRows() - 1), 1);
  const allValidations = validationsRange.getDataValidations();
  const validationValues = validationsRange.getValues();

  let lastValidationRow = 0;
  for (let i = allValidations.length - 1; i >= 0; i--) {
    const validation = allValidations[i] && allValidations[i][0];
    const val = validationValues[i] && validationValues[i][0];
    if (isCheckboxValidation_(validation) || val === true || val === 'TRUE') {
      lastValidationRow = i + 1; // +1 because rows start at index 1 below header
      break;
    }
  }

  const columnHasCheckboxValidation = allValidations.some(row => isCheckboxValidation_(row && row[0]));

  const dataRowCount = Math.max(lastRow - 1, lastValidationRow);
  if (dataRowCount < 1) return;

  const dataRange = stock.getRange(2, 1, dataRowCount, lastColumn);
  const dataValues = dataRange.getValues();
  const validations = allValidations.slice(0, dataRowCount);
  const stampValues = C_STAMPV ? stock.getRange(2, C_STAMPV, dataRowCount, 1).getValues() : [];

  const baseToDmsMap = buildBaseToStockDate_(ss);
  const shippingLookup = buildShippingFeeLookup_(ss);
  if (!shippingLookup) {
    ss.toast('Impossible de calculer les frais de colissage : configure la feuille "Frais".', 'Stock', 6);
    return;
  }

  const rowsToExport = [];
  for (let i = 0; i < dataValues.length; i++) {
    const rowNumber = i + 2;
    const validation = validations[i] && validations[i][0];
    const hasCheckboxValidation = isCheckboxValidation_(validation)
      || (!validation && columnHasCheckboxValidation);
    if (!hasCheckboxValidation) {
      continue;
    }

    const checkboxValue = dataValues[i][C_VALIDE - 1];
    const isChecked = checkboxValue === true
      || checkboxValue === 'TRUE'
      || (typeof checkboxValue === 'string' && checkboxValue.trim().toLowerCase() === 'true');
    const alreadyExported = C_STAMPV && stampValues[i] && stampValues[i][0];
    if (alreadyExported) {
      continue;
    }

    const chronoCheck = enforceChronologicalDates_(stock, rowNumber, {
      dms: C_DMS,
      dmis: C_DMIS,
      dpub: C_DPUB,
      dvente: C_DVENTE
    }, { requireAllDates: true });
    if (!chronoCheck.ok) {
      ss.toast(chronoCheck.message || 'Ordre chronologique des dates invalide.', 'Stock', 6);
      continue;
    }

    if (!ensureValidPriceOrWarn_(stock, rowNumber, C_PRIX)) {
      continue;
    }

    if (!C_TAILLE) {
      ss.toast('Colonne taille introuvable ("TAILLE" / "TAILLE DU COLIS").', 'Stock', 6);
      continue;
    }

    const tailleValue = String(dataValues[i][C_TAILLE - 1] || '').trim();
    if (!tailleValue) {
      ss.toast(`Indique la colonne ${tailleHeaderLabel} avant de valider.`, 'Stock', 6);
      continue;
    }

    const lotValue = C_LOT ? String(dataValues[i][C_LOT - 1] || '').trim() : '';
    const fraisColis = shippingLookup(tailleValue, lotValue);
    if (!Number.isFinite(fraisColis)) {
      const lotMessage = lotValue ? ` / lot ${lotValue}` : '';
      ss.toast(`Frais de colissage introuvables pour la taille ${tailleValue}${lotMessage}.`, 'Stock', 6);
      continue;
    }

    const perItemFee = computePerItemShippingFee_(fraisColis, lotValue);
    const willCheckBox = !isChecked;
    rowsToExport.push({
      rowNumber,
      shipping: { size: tailleValue, lot: lotValue, fee: perItemFee },
      markCheckbox: willCheckBox
    });
  }

  if (!rowsToExport.length) return;

  const exportedRows = [];
  rowsToExport.forEach(entry => {
    if (entry.markCheckbox) {
      stock.getRange(entry.rowNumber, C_VALIDE).setValue(true);
    }

    const success = exportVente_(
      null,
      entry.rowNumber,
      C_ID,
      C_LABEL,
      C_SKU,
      C_PRIX,
      C_DVENTE,
      C_STAMPV,
      baseToDmsMap,
      { shipping: entry.shipping, skipDelete: true }
    );
    if (success) {
      exportedRows.push(entry.rowNumber);
    }
  });

  if (!exportedRows.length) return;

  exportedRows.sort((a, b) => b - a).forEach(rowNumber => {
    stock.deleteRow(rowNumber);
  });
}

function bulkValidateStockSelection() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getActiveSheet();
  if (!stock || stock.getName() !== 'Stock') {
    ss.toast('Ouvre la feuille "Stock" et sélectionne les cases à valider.', 'Stock', 6);
    return;
  }

  const lastRow = stock.getLastRow();
  if (lastRow < 2) {
    ss.toast('Aucune ligne à traiter dans "Stock".', 'Stock', 5);
    return;
  }

  const headers = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID = colExact(HEADERS.STOCK.ID);
  const C_LABEL = colExact(HEADERS.STOCK.LIBELLE)
    || colExact(HEADERS.STOCK.LIBELLE_ALT)
    || colExact(HEADERS.STOCK.ARTICLE)
    || colExact(HEADERS.STOCK.ARTICLE_ALT)
    || colWhere(h => h.includes('libell'))
    || colWhere(h => h.includes('article'))
    || 2;
  const C_SKU = colExact(HEADERS.STOCK.SKU)
    || colExact(HEADERS.STOCK.REFERENCE)
    || colWhere(h => h.includes('sku'));
  const C_PRIX = colExact(HEADERS.STOCK.PRIX_VENTE)
    || colWhere(h => h.includes('prix') && h.includes('vente'));
  const C_TAILLE = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const tailleHeaderLabel = getHeaderLabel_(resolver, C_TAILLE, HEADERS.STOCK.TAILLE);
  const C_LOT = colExact(HEADERS.STOCK.LOT) || colWhere(h => h.includes('lot'));
  const combinedVendu = resolveCombinedVenduColumn_(resolver);
  const legacyVendu = combinedVendu ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyVenduColumns_(resolver);
  const C_VENDU = combinedVendu || legacyVendu.checkboxCol || colExact(HEADERS.STOCK.VENDU_ALT);
  const C_DVENTE = combinedVendu || legacyVendu.dateCol || colExact(HEADERS.STOCK.DATE_VENTE_ALT) || 10;
  const C_STAMPV = colExact(HEADERS.STOCK.VENTE_EXPORTEE_LE) || colWhere(h => h.includes('exporte'));
  const C_VALIDE = colExact(HEADERS.STOCK.VALIDER_SAISIE)
    || colExact(HEADERS.STOCK.VALIDER_SAISIE_ALT)
    || colWhere(h => h.includes('valider'));
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK) || colWhere(h => h.includes('mise en stock'));
  const combinedMis = resolveCombinedMisEnLigneColumn_(resolver);
  const legacyMis = combinedMis ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyMisEnLigneColumn_(resolver);
  const C_DMIS = combinedMis || legacyMis.dateCol;
  const combinedPub = resolveCombinedPublicationColumn_(resolver);
  const legacyPub = combinedPub ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyPublicationColumns_(resolver);
  const C_DPUB = combinedPub || legacyPub.dateCol;

  if (!C_VALIDE) {
    ss.toast('Colonne "VALIDER" introuvable dans Stock.', 'Stock', 6);
    return;
  }
  if (!C_VENDU || !C_DVENTE || !C_PRIX) {
    ss.toast('Vérifie les en-têtes : VENDU, DATE DE VENTE et PRIX DE VENTE sont requis.', 'Stock', 6);
    return;
  }
  if (!C_TAILLE) {
    ss.toast('Colonne taille introuvable ("TAILLE" / "TAILLE DU COLIS").', 'Stock', 6);
    return;
  }

  const selection = ss.getSelection();
  const rangeList = selection ? selection.getActiveRangeList() : null;
  let ranges = [];
  if (rangeList) {
    ranges = rangeList.getRanges();
  } else if (selection && selection.getActiveRange()) {
    ranges = [selection.getActiveRange()];
  }

  if (!ranges.length) {
    ss.toast('Sélectionne les cases "VALIDER" à cocher.', 'Stock', 5);
    return;
  }

  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (!range) continue;
    if (range.getNumColumns() !== 1 || range.getColumn() !== C_VALIDE) {
      ss.toast('La sélection doit être limitée à la colonne "VALIDER".', 'Stock', 6);
      return;
    }
  }

  const validationsRange = stock.getRange(2, C_VALIDE, Math.max(1, stock.getMaxRows() - 1), 1);
  const columnValidations = validationsRange.getDataValidations();
  const columnHasCheckboxValidation = columnValidations.some(row => isCheckboxValidation_(row && row[0]));

  const baseToDmsMap = buildBaseToStockDate_(ss);
  const shippingLookup = buildShippingFeeLookup_(ss);
  if (!shippingLookup) {
    ss.toast('Impossible de calculer les frais de colissage : configure la feuille "Frais".', 'Stock', 6);
    return;
  }

  const lastColumn = stock.getLastColumn();
  const rowsToExport = [];

  ranges.forEach(range => {
    if (!range) return;
    const numRows = range.getNumRows();
    const startRow = range.getRow();

    const blockValues = stock.getRange(startRow, 1, numRows, lastColumn).getValues();

    for (let offset = 0; offset < numRows; offset++) {
      const rowNumber = startRow + offset;
      if (rowNumber <= 1) continue;

      const validationIndex = rowNumber - 2;
      const validation = validationIndex >= 0 && validationIndex < columnValidations.length
        ? columnValidations[validationIndex][0]
        : null;
      const hasCheckboxValidation = isCheckboxValidation_(validation) || (!validation && columnHasCheckboxValidation);
      if (!hasCheckboxValidation) {
        continue;
      }

      const rowValues = blockValues[offset];
      const alreadyExported = C_STAMPV && rowValues[C_STAMPV - 1];
      if (alreadyExported) {
        continue;
      }

      const vendu = C_VENDU ? isStatusActiveValue_(rowValues[C_VENDU - 1]) : false;
      if (!vendu) {
        continue;
      }

      const chronoCheck = enforceChronologicalDates_(stock, rowNumber, {
        dms: C_DMS,
        dmis: C_DMIS,
        dpub: C_DPUB,
        dvente: C_DVENTE
      }, { requireAllDates: true });
      if (!chronoCheck.ok) {
        ss.toast(chronoCheck.message || 'Ordre chronologique des dates invalide.', 'Stock', 6);
        continue;
      }

      if (!ensureValidPriceOrWarn_(stock, rowNumber, C_PRIX)) {
        continue;
      }

      const tailleValue = String(rowValues[C_TAILLE - 1] || '').trim();
      if (!tailleValue) {
        ss.toast(`Indique la colonne ${tailleHeaderLabel} avant de valider.`, 'Stock', 6);
        continue;
      }

      const lotValue = C_LOT ? String(rowValues[C_LOT - 1] || '').trim() : '';
      const fraisColis = shippingLookup(tailleValue, lotValue);
      if (!Number.isFinite(fraisColis)) {
        const lotMessage = lotValue ? ` / lot ${lotValue}` : '';
        ss.toast(`Frais de colissage introuvables pour la taille ${tailleValue}${lotMessage}.`, 'Stock', 6);
        continue;
      }

      const perItemFee = computePerItemShippingFee_(fraisColis, lotValue);
      const checkboxValue = rowValues[C_VALIDE - 1];
      const willCheck = checkboxValue !== true && checkboxValue !== 'TRUE'
        && !(typeof checkboxValue === 'string' && checkboxValue.trim().toLowerCase() === 'true');

      rowsToExport.push({
        rowNumber,
        shipping: { size: tailleValue, lot: lotValue, fee: perItemFee },
        markCheckbox: willCheck
      });
    }
  });

  if (!rowsToExport.length) {
    ss.toast('Aucune ligne prête dans la sélection.', 'Stock', 5);
    return;
  }

  const exportedRows = [];
  rowsToExport.forEach(entry => {
    if (entry.markCheckbox) {
      stock.getRange(entry.rowNumber, C_VALIDE).setValue(true);
    }

    const success = exportVente_(
      null,
      entry.rowNumber,
      C_ID,
      C_LABEL,
      C_SKU,
      C_PRIX,
      C_DVENTE,
      C_STAMPV,
      baseToDmsMap,
      { shipping: entry.shipping, skipDelete: true }
    );

    if (success) {
      exportedRows.push(entry.rowNumber);
    }
  });

  if (!exportedRows.length) {
    ss.toast('Aucune vente exportée depuis la sélection.', 'Stock', 5);
    return;
  }

  exportedRows.sort((a, b) => b - a).forEach(rowNumber => stock.deleteRow(rowNumber));
  ss.toast(`${exportedRows.length} ligne(s) ont été validées et déplacées.`, 'Stock', 5);
}

function exportVente_(e, row, C_ID, C_LABEL, C_SKU, C_PRIX, C_DVENTE, C_STAMPV, baseToDmsMap, options) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Stock");
  if (!sh) return false;

  const opts = options || {};
  const shipping = opts.shipping || null;
  const skipDelete = Boolean(opts.skipDelete);

  const ventes = ss.getSheetByName("Ventes") || ss.insertSheet("Ventes");
  if (ventes.getLastRow() === 0) {
    ventes.getRange(1, 1, 1, DEFAULT_VENTES_HEADERS.length).setValues([DEFAULT_VENTES_HEADERS]);
  }

  const headerWidth = Math.max(DEFAULT_VENTES_HEADERS.length, ventes.getLastColumn());
  const headerRange = ventes.getRange(1, 1, 1, headerWidth);
  const ventesHeaders = headerRange.getValues()[0];
  let headerMutated = false;
  for (let i = 0; i < DEFAULT_VENTES_HEADERS.length; i++) {
    if (!ventesHeaders[i]) {
      ventesHeaders[i] = DEFAULT_VENTES_HEADERS[i];
      headerMutated = true;
    }
  }
  if (headerMutated) {
    headerRange.setValues([ventesHeaders]);
  }

  const ventesResolver = makeHeaderResolver_(ventesHeaders);
  const ventesExact = ventesResolver.colExact.bind(ventesResolver);
  const ventesWhere = ventesResolver.colWhere.bind(ventesResolver);

  const COL_ID_VENTE    = ventesExact(HEADERS.VENTES.ID);
  const COL_DATE_VENTE  = ventesExact(HEADERS.VENTES.DATE_VENTE)
    || ventesWhere(h => h.includes('date') && h.includes('vente'));
  const COL_ARTICLE     = ventesExact(HEADERS.VENTES.ARTICLE)
    || ventesWhere(h => h.includes('article'))
    || ventesWhere(h => h.includes('libell'));
  const COL_SKU_VENTE   = ventesExact(HEADERS.VENTES.SKU);
  const COL_PRIX_VENTE  = ventesExact(HEADERS.VENTES.PRIX_VENTE)
    || ventesExact(HEADERS.VENTES.PRIX_VENTE_ALT)
    || ventesWhere(h => h.includes('prix') && h.includes('vente'));
  const COL_FRAIS_VENTE = ventesExact(HEADERS.VENTES.FRAIS_COLISSAGE)
    || ventesWhere(h => h.includes('frais') && h.includes('colis'));
  const COL_TAILLE_VENTE = ventesExact(HEADERS.VENTES.TAILLE_COLIS)
    || ventesExact(HEADERS.VENTES.TAILLE)
    || ventesWhere(isShippingSizeHeader_);
  const COL_LOT_VENTE   = ventesExact(HEADERS.VENTES.LOT)
    || ventesWhere(h => h.includes('lot'));
  const COL_COMPTA      = ventesExact(HEADERS.VENTES.COMPTABILISE)
    || ventesWhere(h => h.toLowerCase().includes('compt'));
  const COL_DELAI_IMM   = ventesExact(HEADERS.VENTES.DELAI_IMMOBILISATION)
    || ventesWhere(h => h.includes('immobilisation'));
  const COL_DELAI_ML    = ventesExact(HEADERS.VENTES.DELAI_MISE_EN_LIGNE)
    || ventesWhere(h => h.includes('mise en ligne'));
  const COL_DELAI_PUB   = ventesExact(HEADERS.VENTES.DELAI_PUBLICATION)
    || ventesWhere(h => h.includes('publication'));
  const COL_DELAI_VENTE = ventesExact(HEADERS.VENTES.DELAI_VENTE)
    || ventesWhere(h => h.includes('delai') && h.includes('vente'));
  const widthVentes     = Math.max(ventesHeaders.length, DEFAULT_VENTES_HEADERS.length, ventes.getLastColumn());

  const dateCell = sh.getRange(row, C_DVENTE);
  const dateV = dateCell.getValue();
  if (!(dateV instanceof Date) || isNaN(dateV)) return false;

  const idVal = C_ID ? sh.getRange(row, C_ID).getValue() : "";
  const label = C_LABEL ? sh.getRange(row, C_LABEL).getDisplayValue() : "";
  const sku   = C_SKU  ? sh.getRange(row, C_SKU).getDisplayValue() : "";
  const prix  = C_PRIX ? sh.getRange(row, C_PRIX).getValue() : "";

  const achatInfo = getAchatsRecordByIdOrSku_(ss, idVal, sku);
  const prixAchat = achatInfo && Number.isFinite(achatInfo.prixAchat) ? achatInfo.prixAchat : null;
  const dateReception = achatInfo && achatInfo.dateReception instanceof Date && !isNaN(achatInfo.dateReception)
    ? achatInfo.dateReception
    : null;

  const headersS = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const resolverS = makeHeaderResolver_(headersS);

  const C_DMS   = resolverS.colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK);
  const combinedMisForRow = resolveCombinedMisEnLigneColumn_(resolverS);
  const legacyMisForRow = combinedMisForRow ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyMisEnLigneColumn_(resolverS);
  const C_DMIS  = combinedMisForRow || legacyMisForRow.dateCol;
  const combinedPubForRow = resolveCombinedPublicationColumn_(resolverS);
  const legacyPubForRow = combinedPubForRow ? { checkboxCol: 0, dateCol: 0 } : resolveLegacyPublicationColumns_(resolverS);
  const C_DPUB  = combinedPubForRow || legacyPubForRow.dateCol;

  const chronoCheck = enforceChronologicalDates_(sh, row, {
    dms: C_DMS,
    dmis: C_DMIS,
    dpub: C_DPUB,
    dvente: C_DVENTE
  }, { requireAllDates: true });
  if (!chronoCheck.ok) {
    ss.toast(chronoCheck.message || 'Ordre chronologique des dates invalide.', 'Stock', 6);
    return false;
  }

  const dMiseLigne = C_DMIS ? sh.getRange(row, C_DMIS).getValue() : null;
  const dPub       = C_DPUB ? sh.getRange(row, C_DPUB).getValue() : null;

  let dMiseStock = C_DMS ? sh.getRange(row, C_DMS).getValue() : null;
  if (!(dMiseStock instanceof Date) || isNaN(dMiseStock)) {
    dMiseStock = null;
  }

  if (!dMiseStock && sku) {
    const base = extractSkuBase_(sku);
    if (base) {
      const map = baseToDmsMap || buildBaseToStockDate_(ss);
      dMiseStock = map[base] || null;
    }
  }

  function diffDays(toDate, fromDate) {
    if (!(toDate   instanceof Date) || isNaN(toDate))   return "";
    if (!(fromDate instanceof Date) || isNaN(fromDate)) return "";
    const ms = toDate.getTime() - fromDate.getTime();
    return Math.round(ms / (1000 * 60 * 60 * 24));
  }

  const delaiImm = diffDays(dateV, dMiseStock);
  const delaiML  = diffDays(dMiseLigne, dMiseStock);
  const delaiPub = diffDays(dPub, dMiseLigne);
  const delaiVte = diffDays(dateV, dPub);
  const delaiImmFromReception = diffDays(dateV, dateReception);
  const delaiVteFromMiseEnLigne = diffDays(dateV, dMiseLigne);

  const margeBrute = Number.isFinite(prix) && Number.isFinite(prixAchat) ? prix - prixAchat : '';
  const coeffMarge = Number.isFinite(prix) && Number.isFinite(prixAchat) && prixAchat !== 0
    ? prix / prixAchat
    : '';
  const nbPiecesVendu = 1;

  const start = Math.max(2, ventes.getLastRow() + 1);
  const newRow = Array(widthVentes).fill("");
  if (COL_ID_VENTE) newRow[COL_ID_VENTE - 1] = idVal;
  if (COL_DATE_VENTE) newRow[COL_DATE_VENTE - 1] = dateV;
  if (COL_ARTICLE) newRow[COL_ARTICLE - 1] = label;
  if (COL_SKU_VENTE) newRow[COL_SKU_VENTE - 1] = sku;
  if (COL_PRIX_VENTE) newRow[COL_PRIX_VENTE - 1] = prix;
  if (COL_FRAIS_VENTE && shipping && Number.isFinite(shipping.fee)) {
    newRow[COL_FRAIS_VENTE - 1] = shipping.fee;
  }
  if (COL_TAILLE_VENTE && shipping && shipping.size !== undefined) {
    newRow[COL_TAILLE_VENTE - 1] = shipping.size;
  }
  if (COL_LOT_VENTE && shipping && shipping.lot) {
    newRow[COL_LOT_VENTE - 1] = shipping.lot;
  }
  if (COL_DELAI_IMM) newRow[COL_DELAI_IMM - 1] = delaiImm;
  if (COL_DELAI_ML) newRow[COL_DELAI_ML - 1] = delaiML;
  if (COL_DELAI_PUB) newRow[COL_DELAI_PUB - 1] = delaiPub;
  if (COL_DELAI_VENTE) newRow[COL_DELAI_VENTE - 1] = delaiVte;

  ventes.getRange(start, 1, 1, newRow.length).setValues([newRow]);

  copySaleToMonthlySheet_(ss, {
    id: idVal,
    libelle: label,
    dateVente: dateV,
    prixVente: Number.isFinite(prix) ? prix : '',
    prixAchat: Number.isFinite(prixAchat) ? prixAchat : '',
    margeBrute: Number.isFinite(margeBrute) ? margeBrute : '',
    coeffMarge: Number.isFinite(coeffMarge) ? coeffMarge : '',
    nbPieces: nbPiecesVendu,
    sku
  });

  if (COL_COMPTA) {
    ventes.getRange(start, COL_COMPTA).setValue(true);
  }

  const ledgerName = getLedgerSheetNameForDate_(dateV);
  const ledgerSheet = ledgerName ? ss.getSheetByName(ledgerName) : null;
  if (ledgerSheet) {
    removeLedgerDuplicateSkus_(ledgerSheet);
  }

    const lastV = ventes.getLastRow();
    if (lastV > 2 && COL_DATE_VENTE) {
      ventes.getRange(2, 1, lastV - 1, ventes.getLastColumn()).sort([{column: COL_DATE_VENTE, ascending: false}]);
      ventes.getRange(2, COL_DATE_VENTE, lastV - 1, 1).setNumberFormat('dd/MM/yyyy');
    }

  if (shipping && Number.isFinite(shipping.fee)) {
    applyShippingFeeToAchats_(ss, idVal, shipping.fee);
  }

  if (C_STAMPV) sh.getRange(row, C_STAMPV).setValue(new Date());

  if (!skipDelete) {
    sh.deleteRow(row);
  }

  return true;
}

function getAchatsRecordByIdOrSku_(ss, idVal, sku) {
  const achats = ss.getSheetByName('Achats');
  if (!achats) return null;

  const lastRow = achats.getLastRow();
  if (lastRow < 2) return null;

  const headers = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);

  const colId = resolver.colExact(HEADERS.ACHATS.ID);
  const colRef = resolver.colExact(HEADERS.ACHATS.REFERENCE);
  const colPrix = resolver.colExact(HEADERS.ACHATS.PRIX_UNITAIRE_TTC);
  const colDate = resolver.colExact(HEADERS.ACHATS.DATE_LIVRAISON);

  const keyId = idVal !== undefined && idVal !== null && String(idVal).trim() !== ''
    ? normText_(idVal)
    : '';
  const keyRef = sku ? normText_(sku) : '';

  if (!keyId && !keyRef) return null;

  const data = achats.getRange(2, 1, lastRow - 1, achats.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const idMatch = keyId && colId ? normText_(row[colId - 1]) === keyId : false;
    const refMatch = keyRef && colRef ? normText_(row[colRef - 1]) === keyRef : false;
    if (!idMatch && !refMatch) continue;

    const prixCell = colPrix ? row[colPrix - 1] : null;
    const prixAchat = typeof prixCell === 'number' && Number.isFinite(prixCell) ? prixCell : null;
    const dateCell = colDate ? row[colDate - 1] : null;
    const dateReception = dateCell instanceof Date && !isNaN(dateCell) ? dateCell : null;

    return { prixAchat, dateReception };
  }

  return null;
}

function getLedgerSheetNameForDate_(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const month = date.getMonth();
  const year = date.getFullYear();
  return `Compta ${String(month + 1).padStart(2, '0')}-${year}`;
}

function copySaleToMonthlySheet_(ss, sale, options) {
  if (!sale || !(sale.dateVente instanceof Date) || isNaN(sale.dateVente)) {
    return { inserted: false, updated: false, sheetName: '' };
  }

  const opts = options || {};
  const allowUpdates = Boolean(opts.updateExisting);
  const isReturn = Boolean(sale.retour);

  const sheetName = getLedgerSheetNameForDate_(sale.dateVente);
  const month = sale.dateVente.getMonth();
  const year = sale.dateVente.getFullYear();
  const monthStart = new Date(year, month, 1);

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initializeMonthlyLedgerSheet_(sheet, monthStart);
  }

  ensureMonthlyLedgerSheet_(sheet, monthStart);

  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const dedupeKey = buildSaleDedupeKey_(sale);
  const rawSku = sale.sku !== undefined && sale.sku !== null ? String(sale.sku).trim() : '';
  const rawId = sale.id !== undefined && sale.id !== null ? String(sale.id).trim() : '';

  let existingRowNumber = 0;
  if (dedupeKey) {
    existingRowNumber = findLedgerRowByIdentifiers_(sheet, headersLen, sale, dedupeKey, allowUpdates);
  }

  if (!existingRowNumber && (rawSku || rawId)) {
    existingRowNumber = findLedgerRowByIdentifiers_(sheet, headersLen, sale, '', allowUpdates);
  }

  if (existingRowNumber && !allowUpdates) {
    const note = sheet.getRange(existingRowNumber, 1).getNote();
    const noteMatches = dedupeKey && note === dedupeKey;

    const rowId = MONTHLY_LEDGER_INDEX.ID >= 0
      ? sheet.getRange(existingRowNumber, MONTHLY_LEDGER_INDEX.ID + 1).getValue()
      : '';
    const rowSku = MONTHLY_LEDGER_INDEX.SKU >= 0
      ? sheet.getRange(existingRowNumber, MONTHLY_LEDGER_INDEX.SKU + 1).getValue()
      : '';

    const saleKey = buildIdSkuDuplicateKey_(rawId, rawSku);
    const rowKey = buildIdSkuDuplicateKey_(rowId, rowSku);
    const idSkuMatch = saleKey && rowKey && saleKey === rowKey;

    if (!noteMatches && idSkuMatch) {
      return { inserted: false, updated: false, skipped: true, sheetName };
    }
  }

  if (isReturn) {
    if (!existingRowNumber) {
      return { inserted: false, updated: false, removed: false, sheetName };
    }

    const weekNumber = findLedgerWeekNumberForRow_(sheet, existingRowNumber);
    sheet.deleteRow(existingRowNumber);
    if (weekNumber) {
      sortWeekRowsByDate_(sheet, weekNumber, headersLen);
      updateWeeklyTotals_(sheet, weekNumber, headersLen);
    }
    updateMonthlyTotals_(sheet, headersLen);
    updateLedgerResultRow_(sheet, headersLen);

    return { inserted: false, updated: false, removed: true, sheetName };
  }

  const weekRanges = computeMonthlyWeekRanges_(monthStart);
  const saleDayKey = dateToDayKey_(sale.dateVente);
  let weekIndex = weekRanges.findIndex(range => saleDayKey >= range.startKey && saleDayKey <= range.endKey);
  if (weekIndex < 0) {
    weekIndex = weekRanges.length - 1;
  }

  const targetWeekNumber = weekIndex + 1;
  let displacedWeekNumber = 0;

  if (existingRowNumber) {
    const currentWeekNumber = findLedgerWeekNumberForRow_(sheet, existingRowNumber);
    if (currentWeekNumber && currentWeekNumber !== targetWeekNumber) {
      displacedWeekNumber = currentWeekNumber;
      sheet.deleteRow(existingRowNumber);
      existingRowNumber = 0;
    }
  }

  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  const labelPrefix = `SEMAINE ${targetWeekNumber}`;
  const totalPrefix = `TOTAL VENTE SEMAINE ${targetWeekNumber}`;

  const labelIdx = labels.findIndex(v => v.toUpperCase().startsWith(labelPrefix));
  const totalIdx = labels.findIndex((v, idx) => idx > labelIdx && v.toUpperCase().startsWith(totalPrefix));
  if (labelIdx === -1 || totalIdx === -1) {
    return { inserted: false, updated: false };
  }

  const totalRowNumber = totalIdx + 1;
  let saleRowNumber = totalRowNumber;
  let insertedRow = false;
  let updatedRow = false;
  if (!existingRowNumber) {
    sheet.insertRows(totalRowNumber, 1);
    insertedRow = true;
  } else {
    saleRowNumber = existingRowNumber;
    updatedRow = true;
  }

  const saleRow = Array(headersLen).fill('');
  if (MONTHLY_LEDGER_INDEX.ID >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.ID] = rawId || '';
  }
  if (MONTHLY_LEDGER_INDEX.SKU >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.SKU] = rawSku || '';
  }
  if (MONTHLY_LEDGER_INDEX.LIBELLE >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.LIBELLE] = sale.libelle || '';
  }
  if (MONTHLY_LEDGER_INDEX.DATE_VENTE >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.DATE_VENTE] = sale.dateVente;
  }

  const prixVenteValue = (sale.prixVente === '' || sale.prixVente === null || sale.prixVente === undefined)
    ? NaN
    : (typeof sale.prixVente === 'number' ? sale.prixVente : valueToNumber_(sale.prixVente));
  if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0 && Number.isFinite(prixVenteValue)) {
    saleRow[MONTHLY_LEDGER_INDEX.PRIX_VENTE] = roundCurrency_(prixVenteValue);
  }

  const prixAchatValue = (sale.prixAchat === '' || sale.prixAchat === null || sale.prixAchat === undefined)
    ? NaN
    : (typeof sale.prixAchat === 'number' ? sale.prixAchat : valueToNumber_(sale.prixAchat));
  if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0 && Number.isFinite(prixAchatValue)) {
    saleRow[MONTHLY_LEDGER_INDEX.PRIX_ACHAT] = roundCurrency_(prixAchatValue);
  }

  const margeSource = sale.margeBrute;
  const margeValue = (margeSource === '' || margeSource === null || margeSource === undefined)
    ? NaN
    : (typeof margeSource === 'number' ? margeSource : valueToNumber_(margeSource));
  if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.MARGE_BRUTE] = Number.isFinite(margeValue) ? roundCurrency_(margeValue) : '';
  }

  const coeffSource = sale.coeffMarge;
  const coeffValue = (coeffSource === '' || coeffSource === null || coeffSource === undefined)
    ? NaN
    : (typeof coeffSource === 'number' ? coeffSource : valueToNumber_(coeffSource));
  if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
    saleRow[MONTHLY_LEDGER_INDEX.COEFF_MARGE] = Number.isFinite(coeffValue)
      ? Math.round(coeffValue * 100) / 100
      : '';
  }

  const piecesSource = sale.nbPieces;
  const piecesValue = (piecesSource === '' || piecesSource === null || piecesSource === undefined)
    ? NaN
    : (typeof piecesSource === 'number' ? piecesSource : valueToNumber_(piecesSource));
  if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
    if (Number.isFinite(piecesValue)) {
      saleRow[MONTHLY_LEDGER_INDEX.NB_PIECES] = piecesValue;
    } else if (piecesSource !== undefined && piecesSource !== null && piecesSource !== '') {
      saleRow[MONTHLY_LEDGER_INDEX.NB_PIECES] = piecesSource;
    }
  }

  sheet.getRange(saleRowNumber, 1, 1, headersLen).setValues([saleRow]);
  sheet.getRange(saleRowNumber, 1).setNote(dedupeKey || '');

  sortWeekRowsByDate_(sheet, targetWeekNumber, headersLen);
    updateWeeklyTotals_(sheet, targetWeekNumber, headersLen);
    if (displacedWeekNumber) {
      updateWeeklyTotals_(sheet, displacedWeekNumber, headersLen);
    }
    updateMonthlyTotals_(sheet, headersLen);
    updateLedgerResultRow_(sheet, headersLen);

    return { inserted: insertedRow, updated: updatedRow, removed: false, sheetName };
  }

function saleAlreadyInLedgers_(ledgerSheets, sale) {
  if (!Array.isArray(ledgerSheets) || !ledgerSheets.length || !sale) return false;
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const dedupeKey = buildSaleDedupeKey_(sale);
  const targetName = getLedgerSheetNameForDate_(sale.dateVente);

  const orderedSheets = [];
  if (targetName) {
    const targetSheet = ledgerSheets.find(sh => sh && sh.getName && sh.getName() === targetName);
    if (targetSheet) {
      orderedSheets.push(targetSheet);
    }
  }

  ledgerSheets.forEach(sh => {
    if (orderedSheets.indexOf(sh) === -1) {
      orderedSheets.push(sh);
    }
  });

  for (let i = 0; i < orderedSheets.length; i++) {
    const sheet = orderedSheets[i];
    if (!isMonthlyLedgerSheet_(sheet)) continue;
    if (sheet.getLastRow() <= 1) continue;

    const rowNumber = findLedgerRowByIdentifiers_(sheet, headersLen, sale, dedupeKey, false);
    if (rowNumber) return true;
  }

  return false;
}

function ensureMonthlyLedgerSheet_(sheet, monthStart) {
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell || firstCell === '') {
    initializeMonthlyLedgerSheet_(sheet, monthStart);
    return;
  }

  const headerWidth = sheet.getLastColumn();
  const headerValues = headerWidth > 0 ? sheet.getRange(1, 1, 1, headerWidth).getValues()[0] : [];
  if (!headerValues.length || headerValues[0] !== MONTHLY_LEDGER_HEADERS[0]) {
    // La feuille contient d'autres données : ne pas écraser.
    return;
  }

    synchronizeMonthlyLedgerHeaders_(sheet);
    ensureLedgerFeesTable_(sheet);
    applyMonthlySheetFormats_(sheet);
    applyLedgerFeesGuidance_(sheet);
    applyLedgerFeesValidation_(sheet);
    updateLedgerResultRow_(sheet, headersLen);
  }

function buildIdSkuDuplicateKey_(idValue, skuValue) {
  const idPart = idValue !== undefined && idValue !== null && String(idValue).trim() !== ''
    ? normText_(idValue)
    : '';
  const skuPart = skuValue !== undefined && skuValue !== null && String(skuValue).trim() !== ''
    ? normText_(skuValue)
    : '';

  if (!idPart || !skuPart) return '';
  return `${idPart}|${skuPart}`;
}

function removeLedgerDuplicateSkus_(sheet) {
  if (!sheet || !isMonthlyLedgerSheet_(sheet)) return { removed: 0, weeks: [] };
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const last = sheet.getLastRow();
  if (last <= 2) return { removed: 0, weeks: [] };

  const values = sheet.getRange(2, 1, last - 1, headersLen).getValues();
  const seen = {};
  const rowsToDelete = [];
  const affectedWeeks = new Set();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const label = String(values[i][0] || '').toUpperCase();
    if (label.startsWith('SEMAINE') || label.startsWith('TOTAL VENTE') || label.startsWith(LEDGER_RESULT_LABEL)) {
      continue;
    }

    const hasKeyInfo = [MONTHLY_LEDGER_INDEX.ID, MONTHLY_LEDGER_INDEX.SKU, MONTHLY_LEDGER_INDEX.LIBELLE]
      .filter(idx => idx >= 0)
      .some(idx => String(values[i][idx] || '').trim() !== '');
    if (!hasKeyInfo) continue;

    const key = buildIdSkuDuplicateKey_(
      MONTHLY_LEDGER_INDEX.ID >= 0 ? values[i][MONTHLY_LEDGER_INDEX.ID] : '',
      MONTHLY_LEDGER_INDEX.SKU >= 0 ? values[i][MONTHLY_LEDGER_INDEX.SKU] : ''
    );
    if (!key) continue;

    if (seen[key]) {
      rowsToDelete.push(rowNumber);
      const weekNumber = findLedgerWeekNumberForRow_(sheet, rowNumber);
      if (weekNumber) {
        affectedWeeks.add(weekNumber);
      }
    } else {
      seen[key] = rowNumber;
    }
  }

  if (!rowsToDelete.length) {
    return { removed: 0, weeks: [] };
  }

  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

    const weeks = Array.from(affectedWeeks);
    weeks.forEach(week => {
      sortWeekRowsByDate_(sheet, week, headersLen);
      updateWeeklyTotals_(sheet, week, headersLen);
    });
    updateMonthlyTotals_(sheet, headersLen);
    updateLedgerResultRow_(sheet, headersLen);

    return { removed: rowsToDelete.length, weeks };
  }

function buildSaleDedupeKey_(sale) {
  if (!sale) return '';

  if (sale.id !== undefined && sale.id !== null && String(sale.id).trim()) {
    return `ID:${String(sale.id).trim()}`;
  }

  const parts = [];
  if (sale.sku) {
    parts.push(`SKU:${normText_(sale.sku)}`);
  }

  if (sale.dateVente instanceof Date && !isNaN(sale.dateVente)) {
    const y = sale.dateVente.getFullYear();
    const m = String(sale.dateVente.getMonth() + 1).padStart(2, '0');
    const d = String(sale.dateVente.getDate()).padStart(2, '0');
    parts.push(`DATE:${y}-${m}-${d}`);
  }

  if (sale.lot) {
    parts.push(`LOT:${normText_(sale.lot)}`);
  }

  if (sale.taille) {
    parts.push(`SIZE:${normText_(sale.taille)}`);
  }

  const priceValue = sale.prixVente === '' || sale.prixVente === null || sale.prixVente === undefined
    ? NaN
    : (typeof sale.prixVente === 'number' ? sale.prixVente : valueToNumber_(sale.prixVente));
  if (Number.isFinite(priceValue)) {
    parts.push(`PRICE:${roundCurrency_(priceValue)}`);
  }

  if (Number.isFinite(sale.sourceRowNumber)) {
    parts.push(`ROW:${sale.sourceRowNumber}`);
  }

  if (!parts.length && sale.sourceRowNumber) {
    return `ROW:${sale.sourceRowNumber}`;
  }

  return parts.join('|');
}

function isMonthlyLedgerSheet_(sheet) {
  if (!sheet) return false;
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < headersLen) return false;
  const headers = sheet.getRange(1, 1, 1, headersLen).getValues()[0];
  for (let i = 0; i < headersLen; i++) {
    if (headers[i] !== MONTHLY_LEDGER_HEADERS[i]) {
      return false;
    }
  }
  return true;
}

function initializeMonthlyLedgerSheet_(sheet, monthStart) {
  sheet.clear();

  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  sheet.getRange(1, 1, 1, headersLen).setValues([MONTHLY_LEDGER_HEADERS]);
  ensureLedgerFeesTable_(sheet);
  sheet.setFrozenRows(1);

  const weekRanges = computeMonthlyWeekRanges_(monthStart);
  let row = 2;
  for (let i = 0; i < weekRanges.length; i++) {
    const range = weekRanges[i];
    const label = `SEMAINE ${i + 1} ${formatDateString_(range.start)} AU ${formatDateString_(range.end)}`;
    sheet.getRange(row, 1).setValue(label);
    row++;

    const weekTotalRow = Array(headersLen).fill('');
    weekTotalRow[0] = `TOTAL VENTE SEMAINE ${i + 1}`;
    sheet.getRange(row, 1, 1, headersLen)
      .setValues([weekTotalRow]);
    row++;

    sheet.getRange(row, 1, 1, headersLen).clearContent();
    row++;
  }

    const monthTotalRow = Array(headersLen).fill('');
    monthTotalRow[0] = `TOTAL VENTE MOIS`;
    sheet.getRange(row, 1, 1, headersLen)
      .setValues([monthTotalRow]);

    applyMonthlySheetFormats_(sheet);
    applyLedgerFeesGuidance_(sheet);
    applyLedgerFeesValidation_(sheet);
    updateLedgerResultRow_(sheet, headersLen);
  }

function applyMonthlySheetFormats_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (MONTHLY_LEDGER_INDEX.DATE_VENTE >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.DATE_VENTE + 1, maxRows, 1).setNumberFormat('dd/MM/yyyy');
  }
  if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.PRIX_VENTE + 1, maxRows, 1).setNumberFormat('#,##0.00');
  }
  if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.PRIX_ACHAT + 1, maxRows, 1).setNumberFormat('#,##0.00');
  }
  if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.MARGE_BRUTE + 1, maxRows, 1).setNumberFormat('#,##0.00');
  }
  if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.COEFF_MARGE + 1, maxRows, 1).setNumberFormat('0.00');
  }
  if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.NB_PIECES + 1, maxRows, 1).setNumberFormat('0');
  }
  if (LEDGER_FEES_COLUMNS.MONTANT > 0) {
    sheet.getRange(1, LEDGER_FEES_COLUMNS.MONTANT, maxRows, 1).setNumberFormat('#,##0.00');
  }
}

function applyLedgerFeesGuidance_(sheet) {
  if (LEDGER_FEES_COLUMNS.LIBELLE > 0) {
    sheet.getRange(1, LEDGER_FEES_COLUMNS.LIBELLE)
      .setNote('Renseignez ici le libellé précis du frais du mois (ex : location stand, commissions, etc.).');
  }
  if (LEDGER_FEES_COLUMNS.TYPE > 0) {
    sheet.getRange(1, LEDGER_FEES_COLUMNS.TYPE)
      .setNote('Sélectionnez le type de frais correspondant depuis la liste déroulante.');
  }
  if (LEDGER_FEES_COLUMNS.MONTANT > 0) {
    sheet.getRange(1, LEDGER_FEES_COLUMNS.MONTANT)
      .setNote('Saisissez le montant TTC du frais. La somme alimente automatiquement le coût de revient.');
  }
}

function applyLedgerFeesValidation_(sheet) {
  if (!LEDGER_FEES_COLUMNS.TYPE) return;
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LEDGER_FEE_TYPES, true)
    .setAllowInvalid(false)
    .build();
  const rows = Math.max(maxRows - 1, 1);
  sheet.getRange(2, LEDGER_FEES_COLUMNS.TYPE, rows, 1).setDataValidation(rule);
}

function findLedgerWeekNumberForRow_(sheet, rowNumber) {
  if (!sheet || !Number.isFinite(rowNumber) || rowNumber <= 1) {
    return 0;
  }

  const labels = sheet.getRange(1, 1, rowNumber, 1).getValues();
  for (let i = rowNumber - 1; i >= 0; i--) {
    const raw = String(labels[i][0] || '').trim();
    if (!raw) continue;
    const match = raw.match(/SEMAINE\s+(\d+)/i);
    if (match) {
      const parsed = parseInt(match[1], 10);
      return Number.isFinite(parsed) ? parsed : 0;
    }
  }

  return 0;
}

function findLedgerRowByIdentifiers_(sheet, headersLen, sale, dedupeKey, allowUpdates) {
  if (!sheet || !sale) return 0;
  const last = sheet.getLastRow();
  if (last <= 1) return 0;

  const allowDataMatch = Boolean(allowUpdates);
  const allowPriceMismatch = Boolean(allowUpdates);
  const requireExactIdSku = !allowUpdates;
  const normalizedSku = sale.sku ? normText_(sale.sku) : '';
  const skuIndex = MONTHLY_LEDGER_INDEX.SKU;
  const idIndex = MONTHLY_LEDGER_INDEX.ID;
  const saleIdSkuKey = requireExactIdSku ? buildIdSkuDuplicateKey_(sale.id, sale.sku) : '';

  const candidateKeys = [];
  if (dedupeKey) candidateKeys.push(dedupeKey);

  const legacyIdKey = sale.id !== undefined && sale.id !== null ? `ID:${String(sale.id).trim()}` : '';
  if (legacyIdKey && legacyIdKey !== dedupeKey) candidateKeys.push(legacyIdKey);

  const legacySkuKey = sale.sku ? `SKU:${normText_(sale.sku)}` : '';
  if (legacySkuKey && legacySkuKey !== dedupeKey) candidateKeys.push(legacySkuKey);

  if (candidateKeys.length) {
    const notes = sheet.getRange(2, 1, last - 1, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      const note = notes[i][0];
      if (!note || candidateKeys.indexOf(note) === -1) continue;
      const rowIndex = i + 2;

      const noteMatchesSourceRow = typeof dedupeKey === 'string' && dedupeKey.indexOf('ROW:') !== -1
        && note === dedupeKey;
      if (noteMatchesSourceRow) {
        return rowIndex;
      }

      if (ledgerRowMatchesSale_(sheet, headersLen, rowIndex, sale, {
        allowPriceMismatch,
        requireIdSkuPair: requireExactIdSku
      })) {
        return rowIndex;
      }
    }
  }

  const data = sheet.getRange(2, 1, last - 1, headersLen).getValues();
  let firstSkuMatchRow = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (normalizedSku && skuIndex >= 0) {
      const candidateSku = normText_(row[skuIndex]);
      const idSkuMatch = requireExactIdSku
        ? doesRowMatchIdSkuPair_(row, sale, idIndex, skuIndex, saleIdSkuKey)
        : true;

      if (candidateSku === normalizedSku && idSkuMatch) {
        if (!firstSkuMatchRow) firstSkuMatchRow = i + 2;

        if (ledgerRowDataMatchesSale_(row, sale, { allowPriceMismatch: true, requireIdSkuPair: requireExactIdSku })) {
          return i + 2;
        }
      }
    }

    if (allowDataMatch && ledgerRowDataMatchesSale_(row, sale, { allowPriceMismatch, requireIdSkuPair: requireExactIdSku })) {
      return i + 2;
    }
  }

  if (firstSkuMatchRow) {
    return firstSkuMatchRow;
  }

  return 0;
}

function ledgerRowMatchesSale_(sheet, headersLen, rowNumber, sale, options) {
  if (!sheet || !Number.isFinite(rowNumber)) return false;
  const rowValues = sheet.getRange(rowNumber, 1, 1, headersLen).getValues()[0];
  return ledgerRowDataMatchesSale_(rowValues, sale, options);
}

function ledgerRowDataMatchesSale_(row, sale, options) {
  if (!row || !sale) return false;

  const opts = options || {};
  const allowPriceMismatch = Boolean(opts.allowPriceMismatch);
  const requireIdSkuPair = Boolean(opts.requireIdSkuPair);

  const idIndex = MONTHLY_LEDGER_INDEX.ID;
  const skuIndex = MONTHLY_LEDGER_INDEX.SKU;
  const dateIndex = MONTHLY_LEDGER_INDEX.DATE_VENTE;
  const priceIndex = MONTHLY_LEDGER_INDEX.PRIX_VENTE;
  const lotIndex = MONTHLY_LEDGER_INDEX.LOT;
  const sizeIndex = MONTHLY_LEDGER_INDEX.TAILLE_COLIS;

  let matchedIdentifier = false;

  if (requireIdSkuPair) {
    const rowKey = buildIdSkuDuplicateKey_(
      idIndex >= 0 ? row[idIndex] : '',
      skuIndex >= 0 ? row[skuIndex] : ''
    );
    const saleKey = buildIdSkuDuplicateKey_(sale.id, sale.sku);
    if (!saleKey || !rowKey || rowKey !== saleKey) {
      return false;
    }
    matchedIdentifier = true;
  } else {
    if (idIndex >= 0 && sale.id !== undefined && sale.id !== null) {
      const candidateId = String(row[idIndex] || '').trim();
      if (candidateId !== String(sale.id).trim()) {
        return false;
      }
      matchedIdentifier = true;
    }

    const normalizedSku = sale.sku ? normText_(sale.sku) : '';
    if (normalizedSku && skuIndex >= 0) {
      const candidateSku = normText_(row[skuIndex]);
      if (candidateSku !== normalizedSku) {
        return false;
      }
      matchedIdentifier = true;
    } else if (normalizedSku) {
      return false;
    }
  }

  if (dateIndex >= 0 && sale.dateVente instanceof Date && !isNaN(sale.dateVente)) {
    const candidateDate = row[dateIndex];
    if (!(candidateDate instanceof Date) || isNaN(candidateDate)) {
      return false;
    }
    const sameDay = candidateDate.getFullYear() === sale.dateVente.getFullYear()
      && candidateDate.getMonth() === sale.dateVente.getMonth()
      && candidateDate.getDate() === sale.dateVente.getDate();
    if (!sameDay) {
      return false;
    }
  }

  if (lotIndex >= 0 && sale.lot) {
    const candidateLot = normText_(row[lotIndex]);
    if (candidateLot !== normText_(sale.lot)) {
      return false;
    }
  }

  if (sizeIndex >= 0 && sale.taille) {
    const candidateSize = normText_(row[sizeIndex]);
    if (candidateSize !== normText_(sale.taille)) {
      return false;
    }
  }

  if (priceIndex >= 0 && !allowPriceMismatch) {
    const salePrice = sale.prixVente === '' || sale.prixVente === null || sale.prixVente === undefined
      ? NaN
      : (typeof sale.prixVente === 'number' ? sale.prixVente : valueToNumber_(sale.prixVente));
    if (Number.isFinite(salePrice)) {
      const candidatePrice = row[priceIndex];
      if (!Number.isFinite(candidatePrice)) {
        return false;
      }
      if (Math.round(candidatePrice * 100) !== Math.round(salePrice * 100)) {
        return false;
      }
    }
  }

  return matchedIdentifier;
}

function doesRowMatchIdSkuPair_(row, sale, idIndex, skuIndex, prebuiltSaleKey) {
  if (!row || !sale) return false;
  const saleKey = prebuiltSaleKey || buildIdSkuDuplicateKey_(sale.id, sale.sku);
  if (!saleKey) return false;

  const rowKey = buildIdSkuDuplicateKey_(
    idIndex >= 0 ? row[idIndex] : '',
    skuIndex >= 0 ? row[skuIndex] : ''
  );

  return Boolean(rowKey) && rowKey === saleKey;
}

function sortWeekRowsByDate_(sheet, weekNumber, headersLen) {
  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  const labelPrefix = `SEMAINE ${weekNumber}`;
  const totalPrefix = `TOTAL VENTE SEMAINE ${weekNumber}`;
  const labelIdx = labels.findIndex(v => v.toUpperCase().startsWith(labelPrefix));
  const totalIdx = labels.findIndex((v, idx) => idx > labelIdx && v.toUpperCase().startsWith(totalPrefix));
  if (labelIdx === -1 || totalIdx === -1) return;

  const dataCount = totalIdx - labelIdx - 1;
  if (dataCount <= 1) return;

  const firstDataRow = labelIdx + 2;
  const dateColumn = MONTHLY_LEDGER_INDEX.DATE_VENTE >= 0 ? MONTHLY_LEDGER_INDEX.DATE_VENTE + 1 : 3;
  sheet.getRange(firstDataRow, 1, dataCount, headersLen).sort({ column: dateColumn, ascending: true });
}

function updateWeeklyTotals_(sheet, weekNumber, headersLen) {
  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  const labelPrefix = `SEMAINE ${weekNumber}`;
  const totalPrefix = `TOTAL VENTE SEMAINE ${weekNumber}`;
  const labelIdx = labels.findIndex(v => v.toUpperCase().startsWith(labelPrefix));
  const totalIdx = labels.findIndex((v, idx) => idx > labelIdx && v.toUpperCase().startsWith(totalPrefix));
  if (labelIdx === -1 || totalIdx === -1) return;

  const dataCount = totalIdx - labelIdx - 1;
  const totalRowNumber = totalIdx + 1;
  const totals = Array(headersLen).fill('');
  totals[0] = `TOTAL VENTE SEMAINE ${weekNumber}`;

  if (dataCount > 0) {
    const dataRange = sheet.getRange(labelIdx + 2, 1, dataCount, headersLen);
    const rows = dataRange.getValues();

    let sumPrixVente = 0;
    let sumPrixAchat = 0;
    let sumMarge = 0;
    let sumCoeff = 0;
    let sumPieces = 0;
    let countCoeff = 0;
    let rowCount = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const hasKeyInfo = [MONTHLY_LEDGER_INDEX.ID, MONTHLY_LEDGER_INDEX.SKU, MONTHLY_LEDGER_INDEX.LIBELLE]
        .filter(idx => idx >= 0)
        .some(idx => String(row[idx] || '').trim() !== '');
      if (!hasKeyInfo) {
        continue;
      }
      rowCount++;

      if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
        const prixCell = row[MONTHLY_LEDGER_INDEX.PRIX_VENTE];
        if (prixCell !== '' && prixCell !== null && prixCell !== undefined) {
          const prix = valueToNumber_(prixCell);
          if (Number.isFinite(prix)) sumPrixVente += prix;
        }
      }

      if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
        const prixCell = row[MONTHLY_LEDGER_INDEX.PRIX_ACHAT];
        if (prixCell !== '' && prixCell !== null && prixCell !== undefined) {
          const prix = valueToNumber_(prixCell);
          if (Number.isFinite(prix)) sumPrixAchat += prix;
        }
      }

      let marge = NaN;
      if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
        const margeCell = row[MONTHLY_LEDGER_INDEX.MARGE_BRUTE];
        if (margeCell !== '' && margeCell !== null && margeCell !== undefined) {
          marge = valueToNumber_(margeCell);
        }
      }
      if (Number.isFinite(marge)) sumMarge += marge;

      let coeff = NaN;
      if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
        const coeffCell = row[MONTHLY_LEDGER_INDEX.COEFF_MARGE];
        if (coeffCell !== '' && coeffCell !== null && coeffCell !== undefined) {
          coeff = valueToNumber_(coeffCell);
        }
      }
      if (Number.isFinite(coeff)) {
        sumCoeff += coeff;
        countCoeff++;
      }

      let pieces = NaN;
      if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
        const piecesCell = row[MONTHLY_LEDGER_INDEX.NB_PIECES];
        if (piecesCell !== '' && piecesCell !== null && piecesCell !== undefined) {
          pieces = valueToNumber_(piecesCell);
        }
      }
      if (Number.isFinite(pieces)) sumPieces += pieces;
    }

    if (rowCount > 0) {
      totals[0] = `TOTAL VENTE SEMAINE ${weekNumber} : ${rowCount}`;
      if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
        totals[MONTHLY_LEDGER_INDEX.PRIX_VENTE] = roundCurrency_(sumPrixVente);
      }
      if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
        totals[MONTHLY_LEDGER_INDEX.PRIX_ACHAT] = roundCurrency_(sumPrixAchat);
      }
      if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
        totals[MONTHLY_LEDGER_INDEX.MARGE_BRUTE] = roundCurrency_(sumMarge);
      }
      if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
        totals[MONTHLY_LEDGER_INDEX.COEFF_MARGE] = countCoeff
          ? Math.round((sumCoeff / countCoeff) * 100) / 100
          : '';
      }
      if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
        totals[MONTHLY_LEDGER_INDEX.NB_PIECES] = sumPieces;
      }
    }
  }

  sheet.getRange(totalRowNumber, 1, 1, headersLen).setValues([totals]);
  sheet.getRange(totalRowNumber, 1).setNote('');
}

function updateMonthlyTotals_(sheet, headersLen) {
  const summary = summarizeMonthlyLedgerData_(sheet, headersLen);
  if (!summary) return;

  const totals = Array(headersLen).fill('');
  totals[0] = summary.rowCount > 0
    ? `TOTAL VENTE MOIS : ${summary.rowCount}`
    : 'TOTAL VENTE MOIS';

  if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
    totals[MONTHLY_LEDGER_INDEX.PRIX_VENTE] = roundCurrency_(summary.sumPrixVente);
  }
  if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
    totals[MONTHLY_LEDGER_INDEX.PRIX_ACHAT] = roundCurrency_(summary.sumPrixAchat);
  }
  if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
    totals[MONTHLY_LEDGER_INDEX.MARGE_BRUTE] = roundCurrency_(summary.sumMarge);
  }
  if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
    totals[MONTHLY_LEDGER_INDEX.COEFF_MARGE] = summary.countCoeff
      ? Math.round((summary.sumCoeff / summary.countCoeff) * 100) / 100
      : '';
  }
  if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
    totals[MONTHLY_LEDGER_INDEX.NB_PIECES] = summary.sumPieces;
  }
  sheet.getRange(summary.monthRowNumber, 1, 1, headersLen).setValues([totals]);
  sheet.getRange(summary.monthRowNumber, 1).setNote('');
}

const LEDGER_RESULT_LABEL = 'RESULTAT';

function updateLedgerResultRow_(sheet, headersLen) {
  const summary = summarizeMonthlyLedgerData_(sheet, headersLen);
  if (!summary) return;

  const resultRowNumber = ensureLedgerResultRow_(sheet, headersLen, summary.monthRowNumber);
  if (!resultRowNumber) return;

  const totalFees = summary.sumFrais || 0;
  const totalCost = summary.sumPrixAchat + totalFees;
  const net = summary.sumPrixVente - totalCost;

  if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
    const revenueCell = sheet.getRange(resultRowNumber, MONTHLY_LEDGER_INDEX.PRIX_VENTE + 1);
    revenueCell.setValue(`Chiffre d'affaire : ${formatLedgerCurrencyLabel_(summary.sumPrixVente)}`);
    revenueCell.setNote('Somme des prix de vente du mois.');
  }
  if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
    const costCell = sheet.getRange(resultRowNumber, MONTHLY_LEDGER_INDEX.PRIX_ACHAT + 1);
    costCell.setValue(`Coût de revient : ${formatLedgerCurrencyLabel_(totalCost)}`);
    costCell.setNote('Coût de revient = somme des prix d\'achat + total des frais saisis (colonnes K à M).');
  }
  if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
    const profitCell = sheet.getRange(resultRowNumber, MONTHLY_LEDGER_INDEX.MARGE_BRUTE + 1);
    profitCell.setValue(`Bénéfice net : ${formatLedgerCurrencyLabel_(net)}`);
    profitCell.setNote('Bénéfice net = Chiffre d\'affaire - Coût de revient (frais inclus).');
  }
  if (LEDGER_FEES_COLUMNS.MONTANT > 0) {
    const feeCell = sheet.getRange(resultRowNumber, LEDGER_FEES_COLUMNS.MONTANT);
    feeCell.setValue(`Frais : ${formatLedgerCurrencyLabel_(totalFees)}`);
    feeCell.setNote('Total cumulé des montants de frais saisis dans le tableau des colonnes K à M.');
  }
}

function ensureLedgerResultRow_(sheet, headersLen, monthRowNumber) {
  if (!Number.isFinite(monthRowNumber)) return null;

  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  let resultIdx = labels.findIndex((value, idx) => idx > (monthRowNumber - 1)
    && value.toUpperCase().startsWith(LEDGER_RESULT_LABEL));

  if (resultIdx === -1) {
    sheet.insertRowsAfter(monthRowNumber, 1);
    resultIdx = monthRowNumber;
  }

  const resultRowNumber = resultIdx + 1;
  const rowValues = Array(headersLen).fill('');
  rowValues[0] = LEDGER_RESULT_LABEL;
  sheet.getRange(resultRowNumber, 1, 1, headersLen).setValues([rowValues]);
  sheet.getRange(resultRowNumber, 1).setNote('');
  return resultRowNumber;
}

function summarizeMonthlyLedgerData_(sheet, headersLen) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const labels = sheet.getRange(1, 1, lastRow, 1).getValues().map(row => String(row[0] || ''));
  const totalMonthIdx = labels.findIndex(v => v.toUpperCase().startsWith('TOTAL VENTE MOIS'));
  if (totalMonthIdx === -1) return null;

  const monthRowNumber = totalMonthIdx + 1;
  const dataHeight = monthRowNumber - 2;
  const summary = {
    monthRowNumber,
    rowCount: 0,
    sumPrixVente: 0,
    sumPrixAchat: 0,
    sumMarge: 0,
    sumCoeff: 0,
    countCoeff: 0,
    sumPieces: 0,
    sumFrais: 0
  };

  if (dataHeight <= 0) {
    return summary;
  }

  const values = sheet.getRange(2, 1, dataHeight, headersLen).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const labelText = String(row[0] || '').toUpperCase();
    if (labelText.startsWith('SEMAINE') || labelText.startsWith('TOTAL VENTE')) {
      continue;
    }

    const hasKeyInfo = [MONTHLY_LEDGER_INDEX.ID, MONTHLY_LEDGER_INDEX.SKU, MONTHLY_LEDGER_INDEX.LIBELLE]
      .filter(idx => idx >= 0)
      .some(idx => String(row[idx] || '').trim() !== '');
    if (!hasKeyInfo) {
      continue;
    }

    summary.rowCount++;

    if (MONTHLY_LEDGER_INDEX.PRIX_VENTE >= 0) {
      const prixCell = row[MONTHLY_LEDGER_INDEX.PRIX_VENTE];
      if (prixCell !== '' && prixCell !== null && prixCell !== undefined) {
        const prix = valueToNumber_(prixCell);
        if (Number.isFinite(prix)) summary.sumPrixVente += prix;
      }
    }

    if (MONTHLY_LEDGER_INDEX.PRIX_ACHAT >= 0) {
      const prixCell = row[MONTHLY_LEDGER_INDEX.PRIX_ACHAT];
      if (prixCell !== '' && prixCell !== null && prixCell !== undefined) {
        const prix = valueToNumber_(prixCell);
        if (Number.isFinite(prix)) summary.sumPrixAchat += prix;
      }
    }

    if (MONTHLY_LEDGER_INDEX.MARGE_BRUTE >= 0) {
      const margeCell = row[MONTHLY_LEDGER_INDEX.MARGE_BRUTE];
      if (margeCell !== '' && margeCell !== null && margeCell !== undefined) {
        const marge = valueToNumber_(margeCell);
        if (Number.isFinite(marge)) summary.sumMarge += marge;
      }
    }

    if (MONTHLY_LEDGER_INDEX.COEFF_MARGE >= 0) {
      const coeffCell = row[MONTHLY_LEDGER_INDEX.COEFF_MARGE];
      if (coeffCell !== '' && coeffCell !== null && coeffCell !== undefined) {
        const coeff = valueToNumber_(coeffCell);
        if (Number.isFinite(coeff)) {
          summary.sumCoeff += coeff;
          summary.countCoeff++;
        }
      }
    }

    if (MONTHLY_LEDGER_INDEX.NB_PIECES >= 0) {
      const piecesCell = row[MONTHLY_LEDGER_INDEX.NB_PIECES];
      if (piecesCell !== '' && piecesCell !== null && piecesCell !== undefined) {
        const pieces = valueToNumber_(piecesCell);
        if (Number.isFinite(pieces)) summary.sumPieces += pieces;
      }
    }
  }

  summary.sumFrais = computeLedgerFeesTotal_(sheet);
  return summary;
}

function formatLedgerCurrencyLabel_(value) {
  if (!Number.isFinite(value)) return '0,00';
  const rounded = Math.round(value * 100) / 100;
  if (typeof rounded.toLocaleString === 'function') {
    return rounded.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return rounded.toFixed(2);
}

function synchronizeMonthlyLedgerHeaders_(sheet) {
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateIdx = headerValues.indexOf('DATE DE VENTE');
  const prixVentePosition = MONTHLY_LEDGER_INDEX.PRIX_VENTE;
  const prixAchatPosition = MONTHLY_LEDGER_INDEX.PRIX_ACHAT;
  const prixColumnsAligned = prixVentePosition >= 0 && prixAchatPosition >= 0
    && headerValues[prixVentePosition] === MONTHLY_LEDGER_HEADERS[prixVentePosition]
    && headerValues[prixAchatPosition] === MONTHLY_LEDGER_HEADERS[prixAchatPosition];

  if (dateIdx >= 0 && !prixColumnsAligned) {
    sheet.insertColumnsAfter(dateIdx + 1, 2);
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < headersLen) {
    sheet.insertColumnsAfter(lastColumn, headersLen - lastColumn);
  }

  sheet.getRange(1, 1, 1, headersLen).setValues([MONTHLY_LEDGER_HEADERS]);
}

function ensureLedgerFeesTable_(sheet) {
  if (!LEDGER_FEES_TABLE || !LEDGER_FEES_TABLE.HEADERS || LEDGER_FEES_TABLE.HEADERS.length === 0) return;
  if (!Number.isFinite(LEDGER_FEES_TABLE.START_COLUMN) || LEDGER_FEES_TABLE.START_COLUMN <= 0) return;
  const requiredColumns = LEDGER_FEES_TABLE.START_COLUMN + LEDGER_FEES_TABLE.HEADERS.length - 1;
  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < requiredColumns) {
    sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  }
  sheet.getRange(1, LEDGER_FEES_TABLE.START_COLUMN, 1, LEDGER_FEES_TABLE.HEADERS.length)
    .setValues([LEDGER_FEES_TABLE.HEADERS]);
}

function computeLedgerFeesTotal_(sheet) {
  if (!LEDGER_FEES_COLUMNS.MONTANT) return 0;
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return 0;
  const rows = Math.max(maxRows - 1, 1);
  const values = sheet.getRange(2, LEDGER_FEES_COLUMNS.MONTANT, rows, 1).getValues();
  let total = 0;
  for (let i = 0; i < values.length; i++) {
    const raw = values[i][0];
    if (raw === '' || raw === null || raw === undefined) continue;
    const value = valueToNumber_(raw);
    if (Number.isFinite(value)) total += value;
  }
  return total;
}

function valueToNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : NaN;
  }
  if (typeof value === 'string') {
    const normalized = value.replace(/\s+/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  }
  return NaN;
}

function roundCurrency_(value) {
  if (!Number.isFinite(value)) return '';
  return Math.round(value * 100) / 100;
}

function computeMonthlyWeekRanges_(monthStart) {
  if (!(monthStart instanceof Date) || isNaN(monthStart)) return [];

  const ranges = [];
  const start = new Date(monthStart.getFullYear(), monthStart.getMonth(), 1);
  const monthEnd = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);

  let current = start;
  while (current <= monthEnd) {
    const weekStart = new Date(current);
    let weekEnd = new Date(weekStart);
    if (ranges.length === 0) {
      const dow = weekStart.getDay();
      const delta = dow === 0 ? 0 : 7 - dow;
      weekEnd.setDate(weekEnd.getDate() + delta);
    } else {
      weekEnd.setDate(weekEnd.getDate() + 6);
    }
    if (weekEnd > monthEnd) {
      weekEnd = new Date(monthEnd);
    }

    ranges.push({
      start: weekStart,
      end: weekEnd,
      startKey: dateToDayKey_(weekStart),
      endKey: dateToDayKey_(weekEnd)
    });

    current = new Date(weekEnd);
    current.setDate(current.getDate() + 1);
  }

  return ranges;
}

function dateToDayKey_(date) {
  if (!(date instanceof Date) || isNaN(date)) return NaN;
  const y = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();
  return (y * 10000) + (m * 100) + d;
}

function isCheckboxValidation_(validation) {
  if (!validation || typeof validation.getCriteriaType !== 'function') return false;
  const criteria = validation.getCriteriaType();
  return criteria === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
}

function formatDateString_(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

function formatMonthLabel_(year, monthIndex) {
  const safeMonth = Math.min(Math.max(monthIndex, 0), 11);
  return `${MONTH_NAMES_FR[safeMonth]} ${year}`;
}
