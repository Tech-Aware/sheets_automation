function exportVente_(e, row, C_ID, C_LABEL, C_SKU, C_PRIX, C_DVENTE, C_STAMPV, baseToDmsMap, options) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Stock");
  if (!sh) return;

  const opts = options || {};
  const shipping = opts.shipping || null;

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
  const COL_DELAI_IMM   = ventesExact(HEADERS.VENTES.DELAI_IMMOBILISATION)
    || ventesWhere(h => h.includes('immobilisation'));
  const COL_DELAI_ML    = ventesExact(HEADERS.VENTES.DELAI_MISE_EN_LIGNE)
    || ventesWhere(h => h.includes('mise en ligne'));
  const COL_DELAI_PUB   = ventesExact(HEADERS.VENTES.DELAI_PUBLICATION)
    || ventesWhere(h => h.includes('publication'));
  const COL_DELAI_VENTE = ventesExact(HEADERS.VENTES.DELAI_VENTE)
    || ventesWhere(h => h.includes('delai') && h.includes('vente'));
  const COL_RETOUR      = ventesExact(HEADERS.VENTES.RETOUR)
    || ventesWhere(h => h.includes('retour'));
  const widthVentes     = Math.max(ventesHeaders.length, DEFAULT_VENTES_HEADERS.length, ventes.getLastColumn());

  const dateCell = sh.getRange(row, C_DVENTE);
  const dateV = dateCell.getValue();
  if (!(dateV instanceof Date) || isNaN(dateV)) return;

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
    return;
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
  if (COL_RETOUR) newRow[COL_RETOUR - 1] = false;

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

  const lastV = ventes.getLastRow();
  if (lastV > 2 && COL_DATE_VENTE) {
    ventes.getRange(2, 1, lastV - 1, ventes.getLastColumn()).sort([{column: COL_DATE_VENTE, ascending: false}]);
    ventes.getRange(2, COL_DATE_VENTE, lastV - 1, 1).setNumberFormat('dd/MM/yyyy');
  }

  applySkuPaletteFormatting_(ventes, COL_SKU_VENTE, COL_ARTICLE);

  if (shipping && Number.isFinite(shipping.fee)) {
    applyShippingFeeToAchats_(ss, idVal, shipping.fee);
  }

  if (C_STAMPV) sh.getRange(row, C_STAMPV).setValue(new Date());

  sh.deleteRow(row);
}

function handleVentesReturn(e) {
  const ss = e && e.source;
  const sh = ss && ss.getActiveSheet();
  if (!ss || !sh || sh.getName() !== 'Ventes') {
    return;
  }

  const range = e.range;
  if (!range || range.getRow() <= 1) {
    return;
  }

  const lastColumn = sh.getLastColumn();
  if (!lastColumn) {
    return;
  }

  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (!headers.length) {
    return;
  }

  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);
  const colRetour = colExact(HEADERS.VENTES.RETOUR) || colWhere(h => h.includes('retour'));
  if (!colRetour || range.getColumn() !== colRetour) {
    return;
  }

  const newValue = sh.getRange(range.getRow(), colRetour).getValue();
  if (!isReturnFlagActive_(newValue)) {
    return;
  }

  try {
    const rowValues = sh.getRange(range.getRow(), 1, 1, lastColumn).getValues()[0];
    const saleRecord = buildSaleRecordFromVentesRow_(rowValues, resolver);
    if (!saleRecord.ok) {
      revertVentesReturnCell_(sh, range.getRow(), colRetour, e.oldValue);
      ss.toast(saleRecord.message || 'Impossible de lire la ligne de vente.', 'Ventes', 7);
      return;
    }

    const stockPlan = prepareStockReturnPlan_(ss, saleRecord);
    if (!stockPlan.ok) {
      revertVentesReturnCell_(sh, range.getRow(), colRetour, e.oldValue);
      ss.toast(stockPlan.message || 'Impossible de préparer la remise en stock.', 'Ventes', 7);
      return;
    }

    const ledgerResult = copySaleToMonthlySheet_(ss, Object.assign({}, saleRecord.sale, { retour: true }), { updateExisting: true });
    if (!ledgerResult || !ledgerResult.removed) {
      revertVentesReturnCell_(sh, range.getRow(), colRetour, e.oldValue);
      ss.toast('Impossible de retirer la vente de la comptabilité mensuelle.', 'Ventes', 7);
      return;
    }

    const stockResult = applyStockReturnPlan_(stockPlan);
    if (!stockResult.ok) {
      copySaleToMonthlySheet_(ss, saleRecord.sale, { updateExisting: true });
      revertVentesReturnCell_(sh, range.getRow(), colRetour, e.oldValue);
      ss.toast(stockResult.message || 'Retour en stock impossible : la ligne comptable a été rétablie.', 'Ventes', 7);
      return;
    }

    sh.deleteRow(range.getRow());
    const label = saleRecord.sale.sku || saleRecord.sale.libelle || 'vente';
    ss.toast(`Retour traité pour ${label}.`, 'Ventes', 5);
  } catch (err) {
    console.error(err);
    revertVentesReturnCell_(sh, range.getRow(), colRetour, e.oldValue);
    ss.toast('Erreur lors du traitement du retour.', 'Ventes', 7);
  }
}

function buildSaleRecordFromVentesRow_(rowValues, resolver) {
  if (!rowValues || !resolver) {
    return { ok: false, message: 'Ligne ou en-têtes "Ventes" introuvables.' };
  }

  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const colId = colExact(HEADERS.VENTES.ID);
  const colDate = colExact(HEADERS.VENTES.DATE_VENTE)
    || colWhere(h => h.includes('date') && h.includes('vente'));
  const colArticle = colExact(HEADERS.VENTES.ARTICLE)
    || colExact(HEADERS.VENTES.ARTICLE_ALT)
    || colWhere(h => h.includes('article'))
    || colWhere(h => h.includes('libell'));
  const colSku = colExact(HEADERS.VENTES.SKU);
  const colPrix = colExact(HEADERS.VENTES.PRIX_VENTE)
    || colExact(HEADERS.VENTES.PRIX_VENTE_ALT)
    || colWhere(h => h.includes('prix') && h.includes('vente'));
  const colTaille = colExact(HEADERS.VENTES.TAILLE_COLIS)
    || colExact(HEADERS.VENTES.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const colLot = colExact(HEADERS.VENTES.LOT) || colWhere(h => h.includes('lot'));
  const colFrais = colExact(HEADERS.VENTES.FRAIS_COLISSAGE) || colWhere(h => h.includes('frais'));

  const rawId = colId ? rowValues[colId - 1] : '';
  const rawSku = colSku ? rowValues[colSku - 1] : '';
  const rawLibelle = colArticle ? rowValues[colArticle - 1] : '';
  const rawPrix = colPrix ? rowValues[colPrix - 1] : '';
  const rawDate = colDate ? rowValues[colDate - 1] : null;

  const saleDate = getDateOrNull_(rawDate);
  if (!(saleDate instanceof Date) || isNaN(saleDate)) {
    return { ok: false, message: 'Date de vente manquante ou invalide.' };
  }

  const normalizedId = rawId !== undefined && rawId !== null ? String(rawId).trim() : '';
  const normalizedSku = rawSku !== undefined && rawSku !== null ? String(rawSku).trim() : '';
  if (!normalizedId && !normalizedSku) {
    return { ok: false, message: 'ID ou SKU requis pour traiter le retour.' };
  }

  const prixValue = valueToNumber_(rawPrix);
  const sale = {
    id: normalizedId || '',
    sku: normalizedSku || '',
    libelle: rawLibelle !== undefined && rawLibelle !== null ? String(rawLibelle) : '',
    dateVente: saleDate,
    prixVente: Number.isFinite(prixValue) ? prixValue : rawPrix,
    nbPieces: 1
  };

  const restockHints = {
    idValue: rawId,
    skuValue: rawSku,
    libelleValue: rawLibelle,
    prixVenteRaw: rawPrix,
    taille: colTaille ? rowValues[colTaille - 1] : '',
    lot: colLot ? rowValues[colLot - 1] : '',
    frais: colFrais ? rowValues[colFrais - 1] : ''
  };

  return { ok: true, sale, restockHints };
}

function prepareStockReturnPlan_(ss, saleRecord) {
  if (!saleRecord || !saleRecord.sale) {
    return { ok: false, message: 'Informations de vente indisponibles.' };
  }

  const stock = ss.getSheetByName('Stock');
  if (!stock) {
    return { ok: false, message: 'La feuille "Stock" est introuvable.' };
  }

  const lastColumn = stock.getLastColumn();
  if (!lastColumn) {
    return { ok: false, message: 'La feuille "Stock" ne contient aucun en-tête.' };
  }

  const headers = stock.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (!headers.length) {
    return { ok: false, message: 'Impossible de lire les en-têtes "Stock".' };
  }

  const resolver = makeHeaderResolver_(headers);
  const rowValues = Array(headers.length).fill('');
  const hints = saleRecord.restockHints || {};
  const now = new Date();

  const colId = resolver.colExact(HEADERS.STOCK.ID);
  if (colId) {
    rowValues[colId - 1] = hints.idValue !== undefined ? hints.idValue : saleRecord.sale.id;
  }

  const colLibelle = resolver.colExact(HEADERS.STOCK.LIBELLE)
    || resolver.colExact(HEADERS.STOCK.LIBELLE_ALT)
    || resolver.colExact(HEADERS.STOCK.ARTICLE)
    || resolver.colExact(HEADERS.STOCK.ARTICLE_ALT)
    || resolver.colWhere(h => h.includes('libell'));
  if (colLibelle) {
    const libelleValue = hints.libelleValue !== undefined && hints.libelleValue !== null
      ? hints.libelleValue
      : saleRecord.sale.libelle;
    rowValues[colLibelle - 1] = libelleValue || saleRecord.sale.sku || saleRecord.sale.id || '';
  }

  const colSku = resolver.colExact(HEADERS.STOCK.SKU)
    || resolver.colExact(HEADERS.STOCK.REFERENCE);
  if (colSku) {
    rowValues[colSku - 1] = hints.skuValue !== undefined ? hints.skuValue : saleRecord.sale.sku;
  }

  const colPrix = resolver.colExact(HEADERS.STOCK.PRIX_VENTE)
    || resolver.colWhere(h => h.includes('prix') && h.includes('vente'));
  if (colPrix) {
    const prixValue = saleRecord.sale.prixVente;
    if (Number.isFinite(prixValue)) {
      rowValues[colPrix - 1] = prixValue;
    } else if (hints.prixVenteRaw !== undefined) {
      rowValues[colPrix - 1] = hints.prixVenteRaw;
    }
  }

  const colTaille = resolver.colExact(HEADERS.STOCK.TAILLE_COLIS)
    || resolver.colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || resolver.colExact(HEADERS.STOCK.TAILLE)
    || resolver.colWhere(isShippingSizeHeader_);
  if (colTaille && hints.taille !== undefined) {
    rowValues[colTaille - 1] = hints.taille;
  }

  const colLot = resolver.colExact(HEADERS.STOCK.LOT) || resolver.colWhere(h => h.includes('lot'));
  if (colLot && hints.lot !== undefined) {
    rowValues[colLot - 1] = hints.lot;
  }

  const colStamp = resolver.colExact(HEADERS.STOCK.VENTE_EXPORTEE_LE);
  if (colStamp) {
    rowValues[colStamp - 1] = '';
  }

  const statusContext = getStockStatusColumnContext_(stock);
  const dateColumnsToFormat = [];

  function resetStatusColumns_(checkboxCol, dateCol, isCombined) {
    if (checkboxCol) {
      rowValues[checkboxCol - 1] = isCombined ? '' : false;
    }
    if (dateCol && (!isCombined || checkboxCol !== dateCol)) {
      rowValues[dateCol - 1] = '';
      dateColumnsToFormat.push(dateCol);
    }
  }

  if (statusContext && statusContext.columns) {
    if (statusContext.columns.dms) {
      rowValues[statusContext.columns.dms - 1] = now;
      dateColumnsToFormat.push(statusContext.columns.dms);
    }
    resetStatusColumns_(statusContext.columns.mis, statusContext.columns.dmis, statusContext.combinedFlags.dmis);
    resetStatusColumns_(statusContext.columns.pub, statusContext.columns.dpub, statusContext.combinedFlags.dpub);
    resetStatusColumns_(statusContext.columns.vendu, statusContext.columns.dvente, statusContext.combinedFlags.dvente);
  } else {
    const colDms = resolver.colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK);
    if (colDms) {
      rowValues[colDms - 1] = now;
      dateColumnsToFormat.push(colDms);
    }
  }

  return {
    ok: true,
    sheet: stock,
    rowValues,
    dateColumns: Array.from(new Set(dateColumnsToFormat)),
    skuColumn: colSku,
    labelColumn: colLibelle
  };
}

function applyStockReturnPlan_(plan) {
  if (!plan || !plan.sheet || !plan.rowValues) {
    return { ok: false, message: 'Plan de remise en stock invalide.' };
  }

  const sheet = plan.sheet;
  const lastRow = sheet.getLastRow();
  let targetRow;
  if (!lastRow) {
    sheet.insertRows(1, 1);
    targetRow = 1;
  } else {
    sheet.insertRowsAfter(lastRow, 1);
    targetRow = lastRow + 1;
  }

  sheet.getRange(targetRow, 1, 1, plan.rowValues.length).setValues([plan.rowValues]);
  if (plan.dateColumns && plan.dateColumns.length) {
    plan.dateColumns.forEach(col => {
      if (col > 0) {
        sheet.getRange(targetRow, col).setNumberFormat('dd/MM/yyyy');
      }
    });
  }

  if (plan.skuColumn || plan.labelColumn) {
    applySkuPaletteFormatting_(sheet, plan.skuColumn, plan.labelColumn);
  }

  return { ok: true };
}

function revertVentesReturnCell_(sheet, row, column, oldValue) {
  if (!sheet || !row || !column) {
    return;
  }
  const cell = sheet.getRange(row, column);
  if (oldValue === undefined || oldValue === null || oldValue === '') {
    cell.setValue(false);
    return;
  }
  if (oldValue === 'TRUE') {
    cell.setValue(true);
    return;
  }
  if (oldValue === 'FALSE') {
    cell.setValue(false);
    return;
  }
  cell.setValue(oldValue);
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

function copySaleToMonthlySheet_(ss, sale, options) {
  if (!sale || !(sale.dateVente instanceof Date) || isNaN(sale.dateVente)) {
    return { inserted: false, updated: false, removed: false };
  }

  const opts = options || {};
  const allowUpdates = Boolean(opts.updateExisting);
  const isReturn = isReturnFlagActive_(sale.retour);

  const month = sale.dateVente.getMonth();
  const year = sale.dateVente.getFullYear();
  const sheetName = `Compta ${String(month + 1).padStart(2, '0')}-${year}`;
  const monthStart = new Date(year, month, 1);

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initializeMonthlyLedgerSheet_(sheet, monthStart);
  }

  ensureMonthlyLedgerSheet_(sheet, monthStart);

  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const rawSku = sale.sku !== undefined && sale.sku !== null ? String(sale.sku).trim() : '';
  const skuKey = rawSku ? `SKU:${normText_(rawSku)}` : '';
  const rawId = sale.id !== undefined && sale.id !== null ? String(sale.id).trim() : '';
  const idKey = rawId ? `ID:${rawId}` : '';
  const dedupeKey = skuKey || idKey;

  let existingRowNumber = 0;
  if (dedupeKey) {
    const last = sheet.getLastRow();
    if (last > 1) {
      const notes = sheet.getRange(2, 1, last - 1, 1).getNotes();
      const matchIndex = notes.findIndex(row => row[0] === dedupeKey);
      if (matchIndex >= 0) {
        existingRowNumber = matchIndex + 2;
      }
    }
    if (existingRowNumber && !allowUpdates && !isReturn) {
      return { inserted: false, updated: false, removed: false };
    }
  }

  const weekRanges = computeMonthlyWeekRanges_(monthStart);
  const saleTime = sale.dateVente.getTime();
  let weekIndex = weekRanges.findIndex(range => saleTime >= range.start.getTime() && saleTime <= range.end.getTime());
  if (weekIndex < 0) {
    weekIndex = weekRanges.length - 1;
  }

  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  const labelPrefix = `SEMAINE ${weekIndex + 1}`;
  const totalPrefix = `TOTAL VENTE SEMAINE ${weekIndex + 1}`;

  const labelIdx = labels.findIndex(v => v.toUpperCase().startsWith(labelPrefix));
  const totalIdx = labels.findIndex((v, idx) => idx > labelIdx && v.toUpperCase().startsWith(totalPrefix));
  if (labelIdx === -1 || totalIdx === -1) {
    return { inserted: false, updated: false, removed: false };
  }

  const totalRowNumber = totalIdx + 1;
  let saleRowNumber = totalRowNumber;
  if (isReturn) {
    if (!allowUpdates || !dedupeKey || !existingRowNumber) {
      return { inserted: false, updated: false, removed: false };
    }

    sheet.deleteRow(existingRowNumber);
    updateWeeklyTotals_(sheet, weekIndex + 1, headersLen);
    updateMonthlyTotals_(sheet, headersLen);
    updateLedgerResultRow_(sheet, headersLen);
    applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
    ensureLedgerWeekHighlight_(sheet, headersLen);
    return { inserted: false, updated: false, removed: true };
  }

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

  sortWeekRowsByDate_(sheet, weekIndex + 1, headersLen);
  updateWeeklyTotals_(sheet, weekIndex + 1, headersLen);
  updateMonthlyTotals_(sheet, headersLen);
  updateLedgerResultRow_(sheet, headersLen);
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);

  return { inserted: insertedRow, updated: updatedRow, removed: false };
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
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);
  updateLedgerResultRow_(sheet, headersLen);
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
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);
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

    ranges.push({ start: weekStart, end: weekEnd });

    current = new Date(weekEnd);
    current.setDate(current.getDate() + 1);
  }

  return ranges;
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
