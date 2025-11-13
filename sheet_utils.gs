// --- utilitaires généraux ---

function escReg_(s){
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normText_(s) {
  return String(s || "")
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\/]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function makeHeaderResolver_(headers) {
  const row = Array.isArray(headers[0]) && headers.length === 1 ? headers[0] : headers;
  const normalized = row.map(normText_);
  return {
    headers: row,
    normalized,
    colExact(name) {
      return normalized.findIndex(h => h === normText_(name)) + 1;
    },
    colWhere(predicate) {
      const idx = normalized.findIndex(predicate);
      return idx >= 0 ? idx + 1 : 0;
    }
  };
}

function isShippingSizeHeader_(normalizedHeader) {
  if (!normalizedHeader) return false;
  if (!normalizedHeader.includes('taille')) return false;
  if (normalizedHeader.includes('colis')) return true;
  return normalizedHeader === 'taille';
}

function getHeaderLabel_(resolver, columnIndex, fallback) {
  if (!resolver || !columnIndex) {
    return fallback;
  }

  const headers = resolver.headers || [];
  const header = headers[columnIndex - 1];
  return header || fallback;
}

function makePrevDateKey_(sheet, row, col) {
  return ['prevDate', sheet.getSheetId(), row, col].join('|');
}

function storePreviousCellValue_(sheet, row, col, value) {
  const props = PropertiesService.getDocumentProperties();
  const key = makePrevDateKey_(sheet, row, col);
  let payload;
  if (value instanceof Date && !isNaN(value)) {
    payload = JSON.stringify({ type: 'date', value: value.getTime() });
  } else if (value === '' || value === null) {
    payload = JSON.stringify({ type: 'empty' });
  } else {
    payload = JSON.stringify({ type: 'value', value: value });
  }
  props.setProperty(key, payload);
}

function restorePreviousCellValue_(sheet, row, col) {
  const props = PropertiesService.getDocumentProperties();
  const key = makePrevDateKey_(sheet, row, col);
  const payload = props.getProperty(key);
  if (!payload) return false;

  let parsed;
  try {
    parsed = JSON.parse(payload);
  } catch (err) {
    props.deleteProperty(key);
    return false;
  }

  const cell = sheet.getRange(row, col);
  switch (parsed.type) {
    case 'date':
      cell.setValue(new Date(parsed.value));
      break;
    case 'empty':
      cell.clearContent();
      break;
    default:
      cell.setValue(parsed.value);
      break;
  }

  props.deleteProperty(key);
  return true;
}

function getDateOrNull_(value) {
  if (value instanceof Date) {
    const time = value.getTime();
    return isNaN(time) ? null : new Date(time);
  }

  if (typeof value === 'number') {
    const date = new Date(value);
    return isNaN(date.getTime()) ? null : date;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;

    const direct = new Date(trimmed);
    if (!isNaN(direct.getTime())) {
      return direct;
    }

    const slashMatch = trimmed.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})$/);
    if (slashMatch) {
      const day = parseInt(slashMatch[1], 10);
      const month = parseInt(slashMatch[2], 10) - 1;
      let year = parseInt(slashMatch[3], 10);
      if (year < 100) {
        year += year >= 70 ? 1900 : 2000;
      }
      const candidate = new Date(year, month, day);
      return isNaN(candidate.getTime()) ? null : candidate;
    }

    const dashMatch = trimmed.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
    if (dashMatch) {
      const year = parseInt(dashMatch[1], 10);
      const month = parseInt(dashMatch[2], 10) - 1;
      const day = parseInt(dashMatch[3], 10);
      const candidate = new Date(year, month, day);
      return isNaN(candidate.getTime()) ? null : candidate;
    }
  }

  return null;
}

function addDays_(date, days) {
  if (!(date instanceof Date)) {
    return null;
  }

  const time = date.getTime();
  if (isNaN(time)) {
    return null;
  }

  const clone = new Date(time);
  clone.setDate(clone.getDate() + days);
  return clone;
}

function getTomorrow_() {
  const date = new Date();
  date.setDate(date.getDate() + 1);
  return date;
}

function resolveCombinedPretPourStockColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.ACHATS.PRET_STOCK_COMBINED)
    || colWhere(h => h.includes('pret') && h.includes('mise en stock') && h.includes('date'));
  return combined || 0;
}

function resolveCombinedMisEnLigneColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.STOCK.MIS_EN_LIGNE)
    || colWhere(h => h.includes('mis en ligne') && h.includes('date'));
  if (combined) {
    return combined;
  }

  return 0;
}

function resolveLegacyMisEnLigneColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const checkboxCol = colExact(HEADERS.STOCK.MIS_EN_LIGNE_ALT)
    || colWhere(h => h.includes('mis') && h.includes('ligne') && !h.includes('date'));
  const dateCol = colExact(HEADERS.STOCK.DATE_MISE_EN_LIGNE_ALT)
    || colWhere(h => h.includes('date') && h.includes('mise en ligne'));

  return { checkboxCol, dateCol };
}

function resolveCombinedPublicationColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.STOCK.PUBLIE)
    || colWhere(h => h.includes('publi') && h.includes('date'));
  return combined || 0;
}

function resolveLegacyPublicationColumns_(resolver) {
  if (!resolver) return { checkboxCol: 0, dateCol: 0 };
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const checkboxCol = colExact(HEADERS.STOCK.PUBLIE_ALT)
    || colWhere(h => h.includes('publi') && !h.includes('date'));
  const dateCol = colExact(HEADERS.STOCK.DATE_PUBLICATION_ALT)
    || colWhere(h => h.includes('date') && h.includes('publi'));

  return { checkboxCol, dateCol };
}

function resolveCombinedVenduColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.STOCK.VENDU)
    || colWhere(h => h.includes('vendu') && h.includes('date'));
  return combined || 0;
}

function resolveLegacyVenduColumns_(resolver) {
  if (!resolver) return { checkboxCol: 0, dateCol: 0 };
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const checkboxCol = colExact(HEADERS.STOCK.VENDU_ALT)
    || colWhere(h => h.includes('vendu') && !h.includes('date'));
  const dateCol = colExact(HEADERS.STOCK.DATE_VENTE_ALT)
    || colWhere(h => h.includes('date') && (h.includes('vente') || h.includes('vendu')));

  return { checkboxCol, dateCol };
}

function isStatusActiveValue_(value) {
  if (value === true) {
    return true;
  }
  const date = getDateOrNull_(value);
  return !!date;
}

function toNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : NaN;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return NaN;

    const normalized = trimmed
      .replace(/\u00A0/g, '')
      .replace(/\s+/g, '')
      .replace(/,/g, '.');

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  if (value === null || value === undefined) {
    return NaN;
  }

  const coerced = Number(value);
  return Number.isFinite(coerced) ? coerced : NaN;
}

function parseLotCount_(lotValue) {
  if (lotValue === null || lotValue === undefined) {
    return NaN;
  }

  const direct = toNumber_(lotValue);
  if (Number.isFinite(direct)) {
    return direct;
  }

  const text = String(lotValue || '').trim();
  if (!text) {
    return NaN;
  }

  const match = text.match(/(\d+[\d.,]*)/);
  if (!match) {
    return NaN;
  }

  const parsed = toNumber_(match[1]);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function computePerItemShippingFee_(rawFee, lotValue) {
  if (!Number.isFinite(rawFee)) {
    return rawFee;
  }

  const lotCount = parseLotCount_(lotValue);
  if (Number.isFinite(lotCount) && lotCount > 1) {
    return rawFee / lotCount;
  }

  return rawFee;
}

function buildShippingFeeLookup_(ss) {
  const frais = ss && typeof ss.getSheetByName === 'function' ? ss.getSheetByName('Frais') : null;
  if (!frais) return null;

  const lastRow = frais.getLastRow();
  const lastCol = frais.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  const headers = frais.getRange(1, 1, 1, lastCol).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const COL_SIZE = colExact(HEADERS.STOCK.TAILLE_COLIS)
    || colExact(HEADERS.STOCK.TAILLE_COLIS_ALT)
    || colExact(HEADERS.STOCK.TAILLE)
    || colWhere(isShippingSizeHeader_);
  const COL_LOT = colExact('LOT') || colWhere(h => h.includes('lot'));
  const COL_FEE = colExact('FRAIS DE COLISSAGE')
    || colWhere(h => h.includes('frais') && h.includes('colis'))
    || colWhere(h => h.includes('frais') && h.includes('exped'));

  if (!COL_SIZE || !COL_FEE) return null;

  const sizeValues = frais.getRange(2, COL_SIZE, lastRow - 1, 1).getValues();
  const lotValues = COL_LOT ? frais.getRange(2, COL_LOT, lastRow - 1, 1).getValues() : null;
  const feeValues = frais.getRange(2, COL_FEE, lastRow - 1, 1).getValues();

  if (sizeValues.length !== feeValues.length) return null;

  const map = new Map();

  for (let i = 0; i < sizeValues.length; i++) {
    const sizeKey = normText_(sizeValues[i][0]);
    if (!sizeKey) continue;

    const lotRaw = lotValues ? lotValues[i][0] : null;
    const lotKey = Number.isFinite(parseLotCount_(lotRaw)) ? parseLotCount_(lotRaw) : null;
    const feeValue = feeValues[i][0];
    const fee = Number.isFinite(toNumber_(feeValue)) ? toNumber_(feeValue) : null;

    if (fee === null) continue;

    const key = lotKey ? `${sizeKey}__lot_${lotKey}` : sizeKey;
    map.set(key, fee);
  }

  return function lookup(size, lot) {
    if (!size) return null;
    const sizeKey = normText_(size);
    if (!sizeKey) return null;

    const lotCount = Number.isFinite(parseLotCount_(lot)) ? parseLotCount_(lot) : null;
    if (lotCount) {
      const key = `${sizeKey}__lot_${lotCount}`;
      if (map.has(key)) {
        return map.get(key);
      }
    }

    return map.get(sizeKey) || null;
  };
}

function applyShippingFeeToAchats_(ss, achatId, fee) {
  if (!Number.isFinite(fee)) {
    return;
  }

  const achats = ss.getSheetByName('Achats');
  if (!achats) {
    return;
  }

  const headers = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);

  const COL_ID = colExact(HEADERS.ACHATS.ID);
  const COL_FEE = colExact(HEADERS.ACHATS.FRAIS_COLISSAGE);
  const COL_TOTAL = colExact(HEADERS.ACHATS.TOTAL_TTC);
  if (!COL_ID || !COL_FEE || !COL_TOTAL) {
    return;
  }

  const last = achats.getLastRow();
  if (last < 2) {
    return;
  }

  const ids = achats.getRange(2, COL_ID, last - 1, 1).getValues();
  const fees = achats.getRange(2, COL_FEE, last - 1, 1).getValues();
  const totals = achats.getRange(2, COL_TOTAL, last - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    const idValue = ids[i][0];
    if (String(idValue) !== String(achatId)) {
      continue;
    }

    const feeCell = achats.getRange(i + 2, COL_FEE);
    feeCell.setValue(fee);

    const totalCell = achats.getRange(i + 2, COL_TOTAL);
    const baseTotal = toNumber_(totals[i][0]);
    if (Number.isFinite(baseTotal)) {
      totalCell.setValue(baseTotal + fee);
    }

    break;
  }
}

function updateAchatsTotalsWithFee_(achats, rowIndex, fee, cols) {
  if (!achats || !Number.isFinite(fee) || !cols) {
    return;
  }

  const { COL_TOTAL_TTC, COL_FRAIS_COLISSAGE } = cols;
  if (!COL_TOTAL_TTC || !COL_FRAIS_COLISSAGE) {
    return;
  }

  const feeCell = achats.getRange(rowIndex, COL_FRAIS_COLISSAGE);
  const currentFee = toNumber_(feeCell.getValue());
  const totalCell = achats.getRange(rowIndex, COL_TOTAL_TTC);
  const currentTotal = toNumber_(totalCell.getValue());

  feeCell.setValue(fee);
  if (Number.isFinite(currentTotal)) {
    const newTotal = currentTotal - (Number.isFinite(currentFee) ? currentFee : 0) + fee;
    totalCell.setValue(newTotal);
  }
}

function enforceChronologicalDates_(sheet, row, cols, options) {
  if (!sheet || !row || !cols) {
    return;
  }

  const {
    C_DMS,
    C_DML,
    C_DPUB,
    C_DVENTE
  } = cols;

  const {
    preventRegression,
    warnOnly,
    toastLabel
  } = options || {};

  const dmsCell = C_DMS ? sheet.getRange(row, C_DMS) : null;
  const dmlCell = C_DML ? sheet.getRange(row, C_DML) : null;
  const dpubCell = C_DPUB ? sheet.getRange(row, C_DPUB) : null;
  const dventeCell = C_DVENTE ? sheet.getRange(row, C_DVENTE) : null;

  const dms = dmsCell ? getDateOrNull_(dmsCell.getValue()) : null;
  const dml = dmlCell ? getDateOrNull_(dmlCell.getValue()) : null;
  const dpub = dpubCell ? getDateOrNull_(dpubCell.getValue()) : null;
  const dvente = dventeCell ? getDateOrNull_(dventeCell.getValue()) : null;

  const violations = [];

  function markViolation(cell, label) {
    if (!cell) return;
    violations.push(label);
    cell.setBackground('#f8d7da');
    cell.setFontColor('#721c24');
  }

  if (dml && dms && dml < dms) {
    markViolation(dmlCell, HEADER_LABELS.dmis);
    if (preventRegression) {
      const restored = restorePreviousCellValue_(sheet, row, C_DML);
      if (!restored) {
        dmlCell.setValue(dms);
      }
    }
  }

  if (dpub && dml && dpub < dml) {
    markViolation(dpubCell, HEADER_LABELS.dpub);
    if (preventRegression) {
      const restored = restorePreviousCellValue_(sheet, row, C_DPUB);
      if (!restored) {
        dpubCell.setValue(dml);
      }
    }
  }

  if (dvente && dpub && dvente < dpub) {
    markViolation(dventeCell, HEADER_LABELS.dvente);
    if (preventRegression) {
      const restored = restorePreviousCellValue_(sheet, row, C_DVENTE);
      if (!restored) {
        dventeCell.setValue(dpub);
      }
    }
  }

  if (!violations.length) {
    [dmsCell, dmlCell, dpubCell, dventeCell].forEach(cell => {
      if (cell) {
        cell.setBackground(null);
        cell.setFontColor(null);
      }
    });
    return;
  }

  const toastTitle = toastLabel || 'Chronologie invalide';
  const toastMessage = `${toastTitle} : ${violations.join(', ')}`;

  if (!warnOnly && typeof SpreadsheetApp !== 'undefined') {
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMessage, '⚠️ Avertissement', 5);
  }
}

function extractSkuBase_(sku) {
  if (!sku) return '';
  const match = String(sku).match(/^([A-Z]+)\d{2,}/i);
  return match ? match[1].toUpperCase() : '';
}

function extractSkuSuffix_(sku, expectedBase) {
  if (!sku) return '';
  const base = expectedBase || extractSkuBase_(sku);
  if (!base) return '';
  return String(sku).slice(base.length);
}

function computeSkuPaletteKey_(sku) {
  const base = extractSkuBase_(sku);
  return base ? base.slice(0, 3) : '';
}

function computeSkuPaletteColor_(sku) {
  const key = computeSkuPaletteKey_(sku);
  if (!key) return SKU_COLOR_DEFAULT;
  if (SKU_COLOR_OVERRIDES[key]) {
    return SKU_COLOR_OVERRIDES[key];
  }
  const index = key.charCodeAt(0) % SKU_COLOR_PALETTE.length;
  return SKU_COLOR_PALETTE[index] || SKU_COLOR_DEFAULT;
}

function applySkuPaletteFormatting_(sheet, skuCol, articleCol) {
  if (!sheet || !skuCol || !articleCol) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const skuRange = sheet.getRange(2, skuCol, lastRow - 1, 1);
  const skuValues = skuRange.getValues();
  const articleRange = sheet.getRange(2, articleCol, lastRow - 1, 1);
  const articleValues = articleRange.getValues();

  const backgrounds = [];
  const fontColors = [];

  for (let i = 0; i < skuValues.length; i++) {
    const sku = skuValues[i][0];
    if (!sku) {
      backgrounds.push(['#FFFFFF']);
      fontColors.push(['#000000']);
      continue;
    }

    const color = computeSkuPaletteColor_(sku);
    backgrounds.push([color.background]);
    fontColors.push([color.text]);

    const article = articleValues[i][0];
    if (article) {
      articleRange.getCell(i + 1, 1).setFontWeight('bold');
    }
  }

  skuRange.setBackgrounds(backgrounds);
  skuRange.setFontColors(fontColors);
}

function ensureLedgerWeekHighlight_(sheet, headersLen) {
  if (!sheet) return;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, headersLen);
  const rules = sheet.getConditionalFormatRules();
  const filtered = rules.filter(rule => {
    return rule.getDescription() !== LEDGER_WEEK_RULE_DESCRIPTION
      && rule.getDescription() !== LEDGER_MONTH_TOTAL_RULE_DESCRIPTION;
  });

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="Semaine"')
    .setBackground('#f0f0f0')
    .setBold(true)
    .setRanges([range])
    .setDescription(LEDGER_WEEK_RULE_DESCRIPTION)
    .build();

  const totalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="Total mensuel"')
    .setBackground('#d1ecf1')
    .setBold(true)
    .setRanges([range])
    .setDescription(LEDGER_MONTH_TOTAL_RULE_DESCRIPTION)
    .build();

  filtered.push(rule, totalRule);
  sheet.setConditionalFormatRules(filtered);
}

function buildBaseToStockDate_(ss) {
  const stock = ss && ss.getSheetByName('Stock');
  const achats = ss && ss.getSheetByName('Achats');
  if (!stock || !achats) return Object.create(null);

  const last = stock.getLastRow();
  if (last < 2) return Object.create(null);

  const headers = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_SKU = colExact(HEADERS.STOCK.SKU);
  const C_REF = colExact(HEADERS.STOCK.REFERENCE) || colWhere(h => h.includes('reference'));
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK) || colWhere(h => h.includes('mise en stock'));
  if (!C_SKU || !C_REF || !C_DMS) return Object.create(null);

  const skuValues = stock.getRange(2, C_SKU, last - 1, 1).getValues();
  const refValues = stock.getRange(2, C_REF, last - 1, 1).getValues();
  const dmsValues = stock.getRange(2, C_DMS, last - 1, 1).getValues();

  const map = Object.create(null);

  for (let i = 0; i < skuValues.length; i++) {
    const sku = skuValues[i][0];
    const ref = refValues[i][0];
    const dms = dmsValues[i][0];
    if (!sku || !ref) continue;

    const base = String(ref).trim();
    if (!base) continue;

    if (dms instanceof Date && !isNaN(dms)) {
      map[base] = dms;
    }
  }

  if (map && typeof map === 'object' && Object.keys(map).length) {
    return map;
  }

  const lastA = achats.getLastRow();
  if (lastA < 2) return Object.create(null);

  const headersA = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolverA = makeHeaderResolver_(headersA);
  const colWhereA = resolverA.colWhere.bind(resolverA);
  const colExactA = resolverA.colExact.bind(resolverA);

  const COL_REF = colExactA(HEADERS.ACHATS.REFERENCE) || colWhereA(h => h.includes('reference'));
  const COL_STP = colExactA(HEADERS.ACHATS.DATE_MISE_EN_STOCK)
    || colExactA(HEADERS.ACHATS.DATE_MISE_EN_STOCK_ALT)
    || colWhereA(h => h.includes('mis en stock'))
    || colWhereA(h => h.includes('mise en stock'));
  if (!COL_REF || !COL_STP) return Object.create(null);

  const refVals = achats.getRange(2, COL_REF, lastA - 1, 1).getValues();
  const stampVals = achats.getRange(2, COL_STP, lastA - 1, 1).getValues();

  const mapFallback = Object.create(null);
  for (let i = 0; i < refVals.length; i++) {
    const base = String(refVals[i][0] || "").trim();
    const dt = stampVals[i][0];
    if (!base) continue;
    if (dt instanceof Date && !isNaN(dt)) {
      mapFallback[base] = dt;
    }
  }
  return mapFallback;
}

function buildIdToSkuBaseMap_(ss) {
  const achats = ss && ss.getSheetByName('Achats');
  if (!achats) return Object.create(null);

  const last = achats.getLastRow();
  if (last < 2) return Object.create(null);

  const headers = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const COL_ID  = colExact(HEADERS.ACHATS.ID);
  const COL_REF = colExact(HEADERS.ACHATS.REFERENCE) || colWhere(h => h.includes('reference'));
  if (!COL_ID || !COL_REF) return Object.create(null);

  const ids  = achats.getRange(2, COL_ID, last - 1, 1).getValues();
  const refs = achats.getRange(2, COL_REF, last - 1, 1).getValues();

  const map = Object.create(null);
  for (let i = 0; i < ids.length; i++) {
    const idRaw = ids[i][0];
    const refRaw = refs[i][0];
    if (idRaw === null || idRaw === undefined || idRaw === '') continue;

    const key = String(idRaw);
    const base = String(refRaw || "").trim();
    if (base) {
      map[key] = base;
    }
  }

  return map;
}

// Vérifie le prix, colore la cellule si invalide, retourne true/false
function ensureValidPriceOrWarn_(sh, row, C_PRIX) {
  if (!C_PRIX) return false;
  const cell = sh.getRange(row, C_PRIX);
  const v = cell.getValue();

  // Prix numérique strictement positif
  if (typeof v === 'number' && !isNaN(v) && v > 0) {
    cell.setBackground(null);
    cell.setFontColor(null);
    return true;
  }

  const disp = cell.getDisplayValue();
  if (!disp || disp.indexOf('⚠️') === -1) {
    cell.setBackground('#0000FF');  // bleu fort
    cell.setFontColor('#FFFF00');   // jaune
    cell.setValue(`⚠️ Vous devez obligatoirement fournir un ${HEADERS.STOCK.PRIX_VENTE}`);
  }

  return false;
}

// Supprime l’alerte dans PRIX DE VENTE si c’est le message ⚠️
function clearPriceAlertIfAny_(sh, row, C_PRIX) {
  if (!C_PRIX) return;
  const cell = sh.getRange(row, C_PRIX);
  const disp = cell.getDisplayValue();

  if (disp && disp.indexOf('⚠️') === 0) {
    cell.clearContent();
    cell.setBackground(null);
    cell.setFontColor(null);
  }
}
