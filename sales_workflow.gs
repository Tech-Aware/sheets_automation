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

  ventes.getRange(start, 1, 1, newRow.length).setValues([newRow]);

  copySaleToMonthlySheet_(ss, {
    id: idVal,
    libelle: label,
    dateVente: dateV,
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

function copySaleToMonthlySheet_(ss, sale) {
  if (!sale || !(sale.dateVente instanceof Date) || isNaN(sale.dateVente)) {
    return false;
  }

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

  if (dedupeKey) {
    const last = sheet.getLastRow();
    if (last > 1) {
      const notes = sheet.getRange(2, 1, last - 1, 1).getNotes();
      const alreadyPresent = notes.some(row => row[0] === dedupeKey);
      if (alreadyPresent) {
        return false;
      }
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
    return false;
  }

  const totalRowNumber = totalIdx + 1;
  sheet.insertRows(totalRowNumber, 1);
  const saleRowNumber = totalRowNumber;

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
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);
  return true;
}

function ensureMonthlyLedgerSheet_(sheet, monthStart) {
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  const firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell || firstCell === '') {
    initializeMonthlyLedgerSheet_(sheet, monthStart);
    return;
  }

  const headerValues = sheet.getRange(1, 1, 1, headersLen).getValues()[0];
  if (headerValues[0] !== MONTHLY_LEDGER_HEADERS[0]) {
    // La feuille contient d'autres données : ne pas écraser.
    return;
  }

  applyMonthlySheetFormats_(sheet);
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);
}

function initializeMonthlyLedgerSheet_(sheet, monthStart) {
  sheet.clear();

  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  sheet.getRange(1, 1, 1, headersLen).setValues([MONTHLY_LEDGER_HEADERS]);
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
  applySkuPaletteFormatting_(sheet, MONTHLY_LEDGER_INDEX.SKU + 1, MONTHLY_LEDGER_INDEX.LIBELLE + 1);
  ensureLedgerWeekHighlight_(sheet, headersLen);
}

function applyMonthlySheetFormats_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (MONTHLY_LEDGER_INDEX.DATE_VENTE >= 0) {
    sheet.getRange(1, MONTHLY_LEDGER_INDEX.DATE_VENTE + 1, maxRows, 1).setNumberFormat('dd/MM/yyyy');
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
  const labels = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(row => String(row[0] || ''));
  const totalMonthIdx = labels.findIndex(v => v.toUpperCase().startsWith('TOTAL VENTE MOIS'));
  if (totalMonthIdx === -1) return;

  const monthRowNumber = totalMonthIdx + 1;
  const dataHeight = monthRowNumber - 2;
  const totals = Array(headersLen).fill('');
  totals[0] = 'TOTAL VENTE MOIS';

  if (dataHeight > 0) {
    const values = sheet.getRange(2, 1, dataHeight, headersLen).getValues();
    let countRows = 0;
    let sumMarge = 0;
    let sumCoeff = 0;
    let sumPieces = 0;
    let countCoeff = 0;

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

      countRows++;

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

    if (countRows > 0) {
      totals[0] = `TOTAL VENTE MOIS : ${countRows}`;
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

  sheet.getRange(monthRowNumber, 1, 1, headersLen).setValues([totals]);
  sheet.getRange(monthRowNumber, 1).setNote('');
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
