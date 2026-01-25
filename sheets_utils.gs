// --- utilitaires généraux ---

function escReg_(s){
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function logDebug_(contexte, details) {
  try {
    const prefix = contexte ? `[${contexte}] ` : '';
    const body = details === undefined || details === null
      ? ''
      : (typeof details === 'string' ? details : JSON.stringify(details));
    const message = `${prefix}${body}`;
    if (typeof console !== 'undefined' && console.log) {
      console.log(message);
    }
    if (typeof Logger !== 'undefined' && Logger.log) {
      Logger.log(message);
    }
  } catch (err) {
    // Ne bloque jamais le workflow pour un log
    try {
      const fallback = `[logDebug_ error] ${err && err.message ? err.message : err}`;
      if (typeof console !== 'undefined' && console.warn) {
        console.warn(fallback);
      }
    } catch (_ignored) {}
  }
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

function getSheetLayoutConfig_(sheetName) {
  if (!SHEET_LAYOUT) {
    return null;
  }

  const key = String(sheetName || '');
  return SHEET_LAYOUT[key] || SHEET_LAYOUT[key.toUpperCase()] || null;
}

function getSheetHeaderRow_(sheetName) {
  const layout = getSheetLayoutConfig_(sheetName);
  const headerRow = layout && Number(layout.HEADER_ROW);
  return headerRow && headerRow > 0 ? headerRow : 1;
}

function getSheetDataStartRow_(sheetName) {
  const layout = getSheetLayoutConfig_(sheetName);
  const dataRow = layout && Number(layout.DATA_START_ROW);
  if (dataRow && dataRow > 0) {
    return dataRow;
  }
  const headerRow = getSheetHeaderRow_(sheetName);
  return headerRow + 1;
}

function getSheetHeaders_(sheet, explicitName) {
  if (!sheet) return [];
  const name = explicitName || (sheet.getName && sheet.getName());
  const headerRow = getSheetHeaderRow_(name || '');
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return [];
  return sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
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
    || colExact(HEADERS.STOCK.MIS_EN_LIGNE_ALT)
    || colExact(HEADERS.STOCK.MIS_EN_LIGNE_ALT2)
    || colWhere(h => h.includes('mis en ligne'));
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
    || colExact(HEADERS.STOCK.MIS_EN_LIGNE_ALT2)
    || colWhere(h => h.includes('mis') && h.includes('ligne') && !h.includes('date'));
  const dateCol = colExact(HEADERS.STOCK.DATE_MISE_EN_LIGNE_ALT)
    || colExact(HEADERS.STOCK.DATE_MISE_EN_LIGNE_ALT2)
    || colWhere(h => h.includes('date') && h.includes('mise en ligne'));

  return { checkboxCol, dateCol };
}

function resolveCombinedPublicationColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.STOCK.PUBLIE)
    || colExact(HEADERS.STOCK.PUBLIE_ALT)
    || colExact(HEADERS.STOCK.PUBLIE_ALT2)
    || colWhere(h =>
      h.includes('publi') &&
      !h.includes('republ') // ✅ exclut REPUBLIÉ
    );

  return combined || 0;
}


function resolveLegacyPublicationColumns_(resolver) {
  if (!resolver) return { checkboxCol: 0, dateCol: 0 };
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const checkboxCol = colExact(HEADERS.STOCK.PUBLIE_ALT)
    || colExact(HEADERS.STOCK.PUBLIE_ALT2)
    || colWhere(h =>
      h.includes('publi') &&
      !h.includes('date') &&
      !h.includes('republ') // ✅ exclut REPUBLIÉ
    );

  const dateCol = colExact(HEADERS.STOCK.DATE_PUBLICATION_ALT)
    || colExact(HEADERS.STOCK.DATE_PUBLICATION_ALT2)
    || colWhere(h =>
      h.includes('date') &&
      h.includes('publi') &&
      !h.includes('republ') // ✅ exclut REPUBLIÉ
    );

  return { checkboxCol, dateCol };
}


function resolveCombinedVenduColumn_(resolver) {
  if (!resolver) return 0;
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const combined = colExact(HEADERS.STOCK.VENDU)
    || colExact(HEADERS.STOCK.VENDU_ALT)
    || colExact(HEADERS.STOCK.VENDU_ALT2)
    || colWhere(h => h.includes('vendu'));
  return combined || 0;
}

function resolveLegacyVenduColumns_(resolver) {
  if (!resolver) return { checkboxCol: 0, dateCol: 0 };
  const colExact = resolver.colExact ? resolver.colExact.bind(resolver) : () => 0;
  const colWhere = resolver.colWhere ? resolver.colWhere.bind(resolver) : () => 0;

  const checkboxCol = colExact(HEADERS.STOCK.VENDU_ALT)
    || colExact(HEADERS.STOCK.VENDU_ALT2)
    || colWhere(h => h.includes('vendu') && !h.includes('date'));
  const dateCol = colExact(HEADERS.STOCK.DATE_VENTE_ALT)
    || colExact(HEADERS.STOCK.DATE_VENTE_ALT2)
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
      .replace(/[€$£]/g, '')
      .replace(/,/g, '.')
      .replace(/\s+/g, '');

    const match = normalized.match(/-?\d+(?:\.\d+)?/);
    if (!match) return NaN;

    const parsed = Number(match[0]);
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
  const COL_LOT = colExact('LOT') || colWhere(h => h === 'lot');
  const COL_FEE = colExact('FRAIS DE COLISSAGE')
    || colWhere(h => h.includes('frais') && h.includes('colis'))
    || colWhere(h => h.includes('frais') && h.includes('exped'));

  const lotColumns = [];
  resolver.normalized.forEach((header, idx) => {
    if (!header.includes('lot')) return;
    const match = header.match(/(\d+[\d.,]*)/);
    const count = match ? toNumber_(match[1]) : NaN;
    if (Number.isFinite(count) && count > 0) {
      lotColumns.push({ col: idx + 1, count });
    }
  });

  if (!COL_SIZE || !COL_FEE) return null;

  const rows = frais.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = new Map();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sizeKey = normText_(row[COL_SIZE - 1]);
    if (!sizeKey) continue;

    const feeValue = row[COL_FEE - 1];
    const baseFee = Number.isFinite(toNumber_(feeValue)) ? toNumber_(feeValue) : null;
    const lotRaw = COL_LOT ? row[COL_LOT - 1] : null;
    const lotKey = Number.isFinite(parseLotCount_(lotRaw)) ? parseLotCount_(lotRaw) : null;

    if (baseFee !== null) {
      const key = lotKey ? `${sizeKey}__lot_${lotKey}` : sizeKey;
      map.set(key, baseFee);
    }

    lotColumns.forEach(({ col, count }) => {
      const value = row[col - 1];
      const fee = Number.isFinite(toNumber_(value)) ? toNumber_(value) : null;
      if (fee !== null) {
        map.set(`${sizeKey}__lot_${count}`, fee);
      }
    });
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
  const result = { ok: true, message: '', violations: [] };
  if (!sheet || !row || !cols) {
    result.ok = false;
    result.message = 'Contrôle chronologique indisponible (feuille ou colonnes manquantes).';
    return result;
  }

  const normalizedCols = {
    dms: cols.C_DMS || cols.dms || cols.DMS || cols.dateMiseEnStock || 0,
    dml: cols.C_DML || cols.C_DMIS || cols.dml || cols.dmis || cols.dateMiseEnLigne || 0,
    dpub: cols.C_DPUB || cols.dpub || cols.datePublication || 0,
    dvente: cols.C_DVENTE || cols.dvente || cols.dateVente || 0
  };

  const C_DMS = normalizedCols.dms;
  const C_DML = normalizedCols.dml;
  const C_DPUB = normalizedCols.dpub;
  const C_DVENTE = normalizedCols.dvente;

  const opts = options || {};
  const preventRegression = !!opts.preventRegression;
  const warnOnly = !!opts.warnOnly;
  const toastLabel = opts.toastLabel;
  const requireAllDates = !!opts.requireAllDates;

  const labels = {
    dms: (HEADER_LABELS && HEADER_LABELS.dms) || 'DATE DE MISE EN STOCK',
    dml: (HEADER_LABELS && HEADER_LABELS.dmis) || 'DATE DE MISE EN LIGNE',
    dpub: (HEADER_LABELS && HEADER_LABELS.dpub) || 'DATE DE PUBLICATION',
    dvente: (HEADER_LABELS && HEADER_LABELS.dvente) || 'DATE DE VENTE'
  };

  function getCell(col) {
    return col ? sheet.getRange(row, col) : null;
  }

  function getDateValue(cell) {
    if (!cell) return null;
    return getDateOrNull_(cell.getValue());
  }

  const dmsCell = getCell(C_DMS);
  const dmlCell = getCell(C_DML);
  const dpubCell = getCell(C_DPUB);
  const dventeCell = getCell(C_DVENTE);

  const dms = getDateValue(dmsCell);
  const dml = getDateValue(dmlCell);
  const dpub = getDateValue(dpubCell);
  const dvente = getDateValue(dventeCell);

  const violations = [];

  function markViolation(cell, message) {
    violations.push(message);
  }

  if (requireAllDates) {
    [
      { key: 'dms', cell: dmsCell, value: dms },
      { key: 'dml', cell: dmlCell, value: dml },
      { key: 'dpub', cell: dpubCell, value: dpub },
      { key: 'dvente', cell: dventeCell, value: dvente }
    ].forEach(entry => {
      if (!normalizedCols[entry.key]) return;
      if (entry.value instanceof Date && !isNaN(entry.value)) return;
      markViolation(entry.cell, `${labels[entry.key]} manquante`);
    });
  }

  if (dml && dms && dml < dms) {
    markViolation(dmlCell, `${labels.dml} < ${labels.dms}`);
    if (preventRegression && C_DML) {
      const restored = restorePreviousCellValue_(sheet, row, C_DML);
      if (!restored) {
        dmlCell.setValue(dms);
      }
    }
  }

  if (dpub && dml && dpub < dml) {
    markViolation(dpubCell, `${labels.dpub} < ${labels.dml}`);
    if (preventRegression && C_DPUB) {
      const restored = restorePreviousCellValue_(sheet, row, C_DPUB);
      if (!restored) {
        dpubCell.setValue(dml);
      }
    }
  }

  if (dvente && dpub && dvente < dpub) {
    markViolation(dventeCell, `${labels.dvente} < ${labels.dpub}`);
    if (preventRegression && C_DVENTE) {
      const restored = restorePreviousCellValue_(sheet, row, C_DVENTE);
      if (!restored) {
        dventeCell.setValue(dpub);
      }
    }
  }

  if (!violations.length) {
    return result;
  }

  const toastTitle = toastLabel || 'Chronologie invalide';
  const toastMessage = `${toastTitle} : ${violations.join(', ')}`;
  result.ok = false;
  result.message = toastMessage;
  result.violations = violations.slice();

  if (!warnOnly && typeof SpreadsheetApp !== 'undefined') {
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMessage, '⚠️ Avertissement', 5);
  }

  return result;
}

function normalizeSkuBase_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toUpperCase();
}

function extractSkuBase_(sku) {
  const normalized = normalizeSkuBase_(sku);
  if (!normalized) return '';
  const match = normalized.match(/^([A-Z]+)/);
  return match ? match[1] : '';
}

function extractSkuSuffix_(sku, expectedBase) {
  const normalizedSku = normalizeSkuBase_(sku);
  const base = normalizeSkuBase_(expectedBase) || extractSkuBase_(normalizedSku);
  if (!normalizedSku || !base || normalizedSku.indexOf(base) !== 0) {
    return null;
  }
  const remainder = normalizedSku.slice(base.length);
  if (!remainder) return null;
  const match = remainder.match(/^-?(\d+)/);
  if (!match) return null;
  const value = parseInt(match[1], 10);
  return Number.isFinite(value) && value > 0 ? value : null;
}

function ensureLegacyFormattingCleared_(ss) {
  if (!ss || typeof PropertiesService === 'undefined') {
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  if (!props) {
    return;
  }

  const cleanupKey = typeof LEGACY_FORMATTING_CLEANUP_KEY === 'string'
    ? LEGACY_FORMATTING_CLEANUP_KEY
    : 'LEGACY_FORMATTING_CLEARED_V1';

  if (props.getProperty(cleanupKey)) {
    return;
  }

  try {
    const sheets = ss.getSheets();
    if (Array.isArray(sheets)) {
      for (let i = 0; i < sheets.length; i++) {
        scrubLegacyConditionalFormatting_(sheets[i]);
      }
    }
    props.setProperty(cleanupKey, new Date().toISOString());
  } catch (err) {
    if (typeof console !== 'undefined' && console.warn) {
      console.warn('ensureLegacyFormattingCleared_ failed', err);
    }
  }
}

function scrubLegacyConditionalFormatting_(sheet) {
  if (!sheet || typeof sheet.getConditionalFormatRules !== 'function') {
    return;
  }

  const rules = sheet.getConditionalFormatRules();
  if (!rules || !rules.length) {
    return;
  }

  const filtered = rules.filter(rule => !shouldRemoveLegacyRule_(rule));
  if (filtered.length !== rules.length) {
    sheet.setConditionalFormatRules(filtered);
  }
}

function shouldRemoveLegacyRule_(rule) {
  if (!rule || typeof rule.getDescription !== 'function') {
    return false;
  }

  const description = String(rule.getDescription() || '').trim();
  if (!description) {
    return false;
  }

  if (typeof LEDGER_WEEK_RULE_DESCRIPTION === 'string'
    && description === LEDGER_WEEK_RULE_DESCRIPTION) {
    return true;
  }

  if (typeof LEDGER_MONTH_TOTAL_RULE_DESCRIPTION === 'string'
    && description === LEDGER_MONTH_TOTAL_RULE_DESCRIPTION) {
    return true;
  }

  return false;
}

function buildBaseToStockDate_(ss) {
  const stock = ss && ss.getSheetByName('Stock');
  const achats = ss && ss.getSheetByName('Achats');
  if (!stock || !achats) return Object.create(null);

  const last = stock.getLastRow();
  const dataStartRow = getSheetDataStartRow_('Stock');
  if (last < dataStartRow) return Object.create(null);

  const headers = getSheetHeaders_(stock, 'Stock');
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_SKU = colExact(HEADERS.STOCK.SKU);
  const C_REF = colExact(HEADERS.STOCK.REFERENCE) || colWhere(h => h.includes('reference'));
  const C_DMS = colExact(HEADERS.STOCK.DATE_MISE_EN_STOCK) || colWhere(h => h.includes('mise en stock'));
  if (!C_SKU || !C_REF || !C_DMS) return Object.create(null);

  const skuValues = stock.getRange(dataStartRow, C_SKU, last - dataStartRow + 1, 1).getValues();
  const refValues = stock.getRange(dataStartRow, C_REF, last - dataStartRow + 1, 1).getValues();
  const dmsValues = stock.getRange(dataStartRow, C_DMS, last - dataStartRow + 1, 1).getValues();

  const map = Object.create(null);

  for (let i = 0; i < skuValues.length; i++) {
    const sku = skuValues[i][0];
    const ref = refValues[i][0];
    const dms = dmsValues[i][0];
    if (!sku || !ref) continue;

    const base = normalizeSkuBase_(ref);
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
    const base = normalizeSkuBase_(refVals[i][0]);
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
    const base = normalizeSkuBase_(refRaw);
    if (base) {
      map[key] = base;
    }
  }

  return map;
}

/**
 * Construit un map ID -> GENRE(data) depuis la feuille Achats.
 * Utilisé pour déterminer la valeur par défaut de VINTED dans Stock.
 */
function buildIdToGenreMap_(ss) {
  const achats = ss && ss.getSheetByName('Achats');
  if (!achats) return Object.create(null);

  const last = achats.getLastRow();
  if (last < 2) return Object.create(null);

  const headers = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const COL_ID = colExact(HEADERS.ACHATS.ID);
  const COL_GENRE = colExact(HEADERS.ACHATS.GENRE_DATA)
    || colExact(HEADERS.ACHATS.GENRE_DATA_ALT)
    || colWhere(h => h.includes('genre') && h.includes('data'));

  if (!COL_ID || !COL_GENRE) return Object.create(null);

  const ids = achats.getRange(2, COL_ID, last - 1, 1).getValues();
  const genres = achats.getRange(2, COL_GENRE, last - 1, 1).getValues();

  const map = Object.create(null);
  for (let i = 0; i < ids.length; i++) {
    const idRaw = ids[i][0];
    const genreRaw = genres[i][0];
    if (idRaw === null || idRaw === undefined || idRaw === '') continue;

    const key = String(idRaw);
    const genre = String(genreRaw || '').trim();
    if (genre) {
      map[key] = genre;
    }
  }

  return map;
}

/**
 * Détermine la valeur VINTED par défaut en fonction du genre.
 * - "Femme" uniquement -> "Durin31"
 * - "Homme" uniquement -> "NivekDazar"
 * - Mixte (Homme/Femme ou Homme/ Femme) -> vide
 */
function getDefaultVintedFromGenre_(genre) {
  if (!genre) return '';

  const normalized = String(genre).toLowerCase().trim();

  // Vérifier si c'est mixte (contient les deux genres)
  const hasHomme = normalized.includes('homme');
  const hasFemme = normalized.includes('femme');

  if (hasHomme && hasFemme) {
    // Mixte -> laisser vide
    return '';
  }

  if (hasFemme && !hasHomme) {
    return 'Durin31';
  }

  if (hasHomme && !hasFemme) {
    return 'NivekDazar';
  }

  // Valeur non reconnue -> laisser vide
  return '';
}

// Vérifie le prix, colore la cellule si invalide, retourne true/false
function ensureValidPriceOrWarn_(sh, row, C_PRIX) {
  if (!C_PRIX) return false;
  const cell = sh.getRange(row, C_PRIX);
  const v = cell.getValue();

  // Prix numérique strictement positif
  if (typeof v === 'number' && !isNaN(v) && v > 0) {
    return true;
  }

  const disp = cell.getDisplayValue();
  if (!disp || disp.indexOf('⚠️') === -1) {
    cell.setValue(`⚠️ Vous devez obligatoirement fournir un ${HEADERS.STOCK.PRIX_VENTE}`);
  }

  return false;
}

// Supprime l'alerte dans PRIX DE VENTE si c'est le message ⚠️
function clearPriceAlertIfAny_(sh, row, C_PRIX) {
  if (!C_PRIX) return;
  const cell = sh.getRange(row, C_PRIX);
  const disp = cell.getDisplayValue();

  if (disp && disp.indexOf('⚠️') === 0) {
    cell.clearContent();
  }
}

// --- Gestion de la colonne VINTED avec liste déroulante ---

/**
 * Applique une validation de liste déroulante sur la colonne VINTED de la feuille Stock.
 * Pré-remplit également les valeurs VINTED vides en fonction du genre depuis Achats.
 * Cette fonction peut être appelée manuellement ou automatiquement.
 */
function applyVintedDropdownToStock() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getSheetByName('Stock');
  if (!stock) {
    ss.toast('Feuille "Stock" introuvable.', 'Erreur', 5);
    return;
  }

  const headerRow = getSheetHeaderRow_('Stock');
  const dataStartRow = getSheetDataStartRow_('Stock');
  const lastColumn = stock.getLastColumn();
  const lastRow = stock.getLastRow();

  if (!lastColumn) {
    ss.toast('La feuille Stock est vide.', 'Erreur', 5);
    return;
  }

  const headers = stock.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  // Chercher la colonne VINTED existante
  let C_VINTED = colExact(HEADERS.STOCK.VINTED)
    || colExact(HEADERS.STOCK.VINTED_ALT)
    || colWhere(h => h.includes('vinted'));

  // Si la colonne n'existe pas, la créer après la colonne LOT
  if (!C_VINTED) {
    const C_LOT = colExact(HEADERS.STOCK.LOT)
      || colExact(HEADERS.STOCK.LOT_ALT2)
      || colExact(HEADERS.STOCK.LOT_ALT3)
      || colWhere(h => h.includes('lot'));

    const C_VALIDER = colExact(HEADERS.STOCK.VALIDER_SAISIE)
      || colExact(HEADERS.STOCK.VALIDER_SAISIE_ALT)
      || colWhere(h => h.includes('valider'));

    // Insérer après LOT si existe, sinon avant VALIDER, sinon à la fin
    if (C_LOT) {
      stock.insertColumnAfter(C_LOT);
      C_VINTED = C_LOT + 1;
    } else if (C_VALIDER) {
      stock.insertColumnBefore(C_VALIDER);
      C_VINTED = C_VALIDER;
    } else {
      C_VINTED = lastColumn + 1;
    }

    // Ajouter le titre de la colonne
    stock.getRange(headerRow, C_VINTED).setValue(HEADERS.STOCK.VINTED);
  }

  // Récupérer les comptes Vinted disponibles
  const vintedAccounts = typeof VINTED_ACCOUNTS !== 'undefined' && Array.isArray(VINTED_ACCOUNTS)
    ? VINTED_ACCOUNTS
    : ['Durin31', 'NivekDazar'];

  // Créer la règle de validation avec liste déroulante
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vintedAccounts, true)
    .setAllowInvalid(false)
    .setHelpText('Sélectionne le compte Vinted sur lequel cet article a été vendu.')
    .build();

  // Appliquer la validation sur toutes les lignes de données
  if (lastRow >= dataStartRow) {
    const numRows = lastRow - dataStartRow + 1;
    stock.getRange(dataStartRow, C_VINTED, numRows, 1).setDataValidation(rule);

    // Pré-remplir les valeurs VINTED vides en fonction du genre depuis Achats
    const C_ID = colExact(HEADERS.STOCK.ID);
    if (C_ID) {
      const idToGenreMap = buildIdToGenreMap_(ss);
      const idValues = stock.getRange(dataStartRow, C_ID, numRows, 1).getValues();
      const vintedValues = stock.getRange(dataStartRow, C_VINTED, numRows, 1).getValues();

      const updates = [];
      for (let i = 0; i < numRows; i++) {
        const currentVinted = String(vintedValues[i][0] || '').trim();
        // Ne pré-remplir que si la cellule est vide
        if (!currentVinted) {
          const idValue = String(idValues[i][0] || '');
          const genre = idToGenreMap[idValue];
          const defaultVinted = getDefaultVintedFromGenre_(genre);
          updates.push([defaultVinted]);
        } else {
          updates.push([currentVinted]);
        }
      }

      // Appliquer les mises à jour en une seule opération
      stock.getRange(dataStartRow, C_VINTED, numRows, 1).setValues(updates);
    }
  }

  ss.toast(`Colonne VINTED configurée avec ${vintedAccounts.length} options.`, 'Stock', 5);
}

/**
 * Applique la validation VINTED sur une nouvelle ligne dans Stock.
 * Pré-remplit également la valeur VINTED en fonction du genre depuis Achats.
 * À appeler lors de l'ajout d'une ligne dans Stock.
 */
function applyVintedDropdownToRow_(sheet, row, idValue) {
  if (!sheet) return;

  const headerRow = getSheetHeaderRow_('Stock');
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_VINTED = colExact(HEADERS.STOCK.VINTED)
    || colExact(HEADERS.STOCK.VINTED_ALT)
    || colWhere(h => h.includes('vinted'));

  if (!C_VINTED) return;

  const vintedAccounts = typeof VINTED_ACCOUNTS !== 'undefined' && Array.isArray(VINTED_ACCOUNTS)
    ? VINTED_ACCOUNTS
    : ['Durin31', 'NivekDazar'];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vintedAccounts, true)
    .setAllowInvalid(false)
    .setHelpText('Sélectionne le compte Vinted sur lequel cet article a été vendu.')
    .build();

  const cell = sheet.getRange(row, C_VINTED);
  cell.setDataValidation(rule);

  // Pré-remplir la valeur VINTED en fonction du genre si un ID est fourni
  if (idValue) {
    const ss = sheet.getParent();
    const idToGenreMap = buildIdToGenreMap_(ss);
    const genre = idToGenreMap[String(idValue)];
    const defaultVinted = getDefaultVintedFromGenre_(genre);
    if (defaultVinted) {
      cell.setValue(defaultVinted);
    }
  }
}

/**
 * Applique une validation de liste déroulante sur la colonne VINTED de la feuille Ventes.
 * Cette fonction peut être appelée manuellement ou automatiquement.
 */
function applyVintedDropdownToVentes() {
  const ss = SpreadsheetApp.getActive();
  const ventes = ss.getSheetByName('Ventes');
  if (!ventes) {
    ss.toast('Feuille "Ventes" introuvable.', 'Erreur', 5);
    return;
  }

  const headerRow = getSheetHeaderRow_('Ventes');
  const dataStartRow = getSheetDataStartRow_('Ventes');
  const lastColumn = ventes.getLastColumn();
  const lastRow = ventes.getLastRow();

  if (!lastColumn) {
    ss.toast('La feuille Ventes est vide.', 'Erreur', 5);
    return;
  }

  const headers = ventes.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(headers);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  // Chercher la colonne VINTED existante
  let C_VINTED = colExact(HEADERS.VENTES.VINTED)
    || colExact(HEADERS.VENTES.VINTED_ALT)
    || colWhere(h => h.includes('vinted'));

  // Si la colonne n'existe pas, la créer après LOT et avant COMPTABILISE
  if (!C_VINTED) {
    const C_LOT = colExact(HEADERS.VENTES.LOT) || colWhere(h => h.includes('lot'));
    const C_COMPTA = colExact(HEADERS.VENTES.COMPTABILISE)
      || colWhere(h => h.toLowerCase().includes('compt'));

    if (C_LOT) {
      ventes.insertColumnAfter(C_LOT);
      C_VINTED = C_LOT + 1;
    } else if (C_COMPTA) {
      ventes.insertColumnBefore(C_COMPTA);
      C_VINTED = C_COMPTA;
    } else {
      C_VINTED = lastColumn + 1;
    }

    ventes.getRange(headerRow, C_VINTED).setValue(HEADERS.VENTES.VINTED);
  }

  // Récupérer les comptes Vinted disponibles
  const vintedAccounts = typeof VINTED_ACCOUNTS !== 'undefined' && Array.isArray(VINTED_ACCOUNTS)
    ? VINTED_ACCOUNTS
    : ['Durin31', 'NivekDazar'];

  // Créer la règle de validation avec liste déroulante
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vintedAccounts, true)
    .setAllowInvalid(true) // Dans Ventes, on autorise les valeurs invalides pour les anciennes lignes
    .setHelpText('Compte Vinted sur lequel cet article a été vendu.')
    .build();

  // Appliquer la validation sur toutes les lignes de données
  if (lastRow >= dataStartRow) {
    const numRows = lastRow - dataStartRow + 1;
    ventes.getRange(dataStartRow, C_VINTED, numRows, 1).setDataValidation(rule);

    // Pré-remplir les valeurs VINTED vides en fonction du genre depuis Achats
    const C_ID = colExact(HEADERS.VENTES.ID) || colWhere(h => h === 'id');
    if (C_ID) {
      const idToGenreMap = buildIdToGenreMap_(ss);
      const idValues = ventes.getRange(dataStartRow, C_ID, numRows, 1).getValues();
      const vintedValues = ventes.getRange(dataStartRow, C_VINTED, numRows, 1).getValues();

      const updates = [];
      for (let i = 0; i < numRows; i++) {
        const currentVinted = String(vintedValues[i][0] || '').trim();
        // Ne pré-remplir que si la cellule est vide
        if (!currentVinted) {
          const idValue = String(idValues[i][0] || '');
          const genre = idToGenreMap[idValue];
          const defaultVinted = getDefaultVintedFromGenre_(genre);
          updates.push([defaultVinted]);
        } else {
          updates.push([currentVinted]);
        }
      }

      // Appliquer les mises à jour en une seule opération
      ventes.getRange(dataStartRow, C_VINTED, numRows, 1).setValues(updates);
    }
  }

  ss.toast(`Colonne VINTED configurée dans Ventes avec ${vintedAccounts.length} options.`, 'Ventes', 5);
}

/**
 * Configure les colonnes VINTED dans Stock ET Ventes en une seule action.
 */
function configureVintedColumns() {
  applyVintedDropdownToStock();
  applyVintedDropdownToVentes();
}

/**
 * Recalcule les taxes et totaux pour toutes les feuilles Compta existantes.
 */
function recalculateAllLedgerTaxes() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const headersLen = MONTHLY_LEDGER_HEADERS.length;
  let updatedCount = 0;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const name = sheet.getName();

    // Vérifier si c'est une feuille Compta (format: "Compta MM-YYYY")
    if (name && name.match(/^Compta\s+\d{2}-\d{4}$/)) {
      updateLedgerResultRow_(sheet, headersLen);
      updatedCount++;
    }
  }

  ss.toast(`${updatedCount} feuille(s) Compta mise(s) à jour avec les taxes.`, 'Calcul des taxes', 5);
}

/**
 * Migre les feuilles Compta existantes vers la nouvelle structure avec colonne TAXES.
 * Ancienne structure: ID, SKU, LIBELLÉS, DATE DE VENTE, PRIX DE VENTE, PRIX D'ACHAT, MARGE BRUTE, COEFF MARGE, NBR PCS VENDU
 * Nouvelle structure: ID, SKU, LIBELLÉS, DATE DE VENTE, PRIX DE VENTE, TAXES, COUT DE REVIENT, BENEFICE NET, NBR PCS VENDU
 */
function migrateAllLedgerSheets() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  let migratedCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const name = sheet.getName();

    // Vérifier si c'est une feuille Compta (format: "Compta MM-YYYY")
    if (name && name.match(/^Compta\s+\d{2}-\d{4}$/)) {
      const result = migrateLedgerSheet_(sheet);
      if (result.migrated) {
        migratedCount++;
      } else {
        skippedCount++;
      }
    }
  }

  // Recalculer les totaux après migration
  if (migratedCount > 0) {
    recalculateAllLedgerTaxes();
  }

  ss.toast(`${migratedCount} feuille(s) migrée(s), ${skippedCount} déjà à jour.`, 'Migration Compta', 5);
}

/**
 * Migre une feuille Compta vers la nouvelle structure.
 * @param {Sheet} sheet - La feuille à migrer
 * @returns {Object} - { migrated: boolean, reason: string }
 */
function migrateLedgerSheet_(sheet) {
  if (!sheet) return { migrated: false, reason: 'no sheet' };

  const lastCol = sheet.getLastColumn();
  if (lastCol < 9) return { migrated: false, reason: 'not enough columns' };

  // Lire les en-têtes actuels
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Vérifier si c'est l'ancien format (colonne 6 = "PRIX D'ACHAT" au lieu de "TAXES")
  const col6Header = String(headers[5] || '').trim().toUpperCase();

  // Si la colonne 6 est déjà "TAXES", la feuille est déjà migrée
  if (col6Header === 'TAXES') {
    return { migrated: false, reason: 'already migrated' };
  }

  // Vérifier que c'est bien l'ancien format
  const isOldFormat = col6Header === "PRIX D'ACHAT" || col6Header === 'PRIX D\'ACHAT';
  if (!isOldFormat) {
    return { migrated: false, reason: 'unknown format: ' + col6Header };
  }

  // 1. Insérer une nouvelle colonne après PRIX DE VENTE (colonne 5 -> nouvelle colonne 6)
  sheet.insertColumnAfter(5);

  // 2. Définir l'en-tête de la nouvelle colonne TAXES
  sheet.getRange(1, 6).setValue('TAXES');

  // 3. Renommer les colonnes décalées
  // Colonne 7 (anciennement 6): PRIX D'ACHAT -> COUT DE REVIENT
  sheet.getRange(1, 7).setValue('COUT DE REVIENT');

  // Colonne 8 (anciennement 7): MARGE BRUTE -> BENEFICE NET
  sheet.getRange(1, 8).setValue('BENEFICE NET');

  // 4. Supprimer l'ancienne colonne COEFF MARGE (maintenant en position 9, anciennement 8)
  // La colonne 9 après insertion est l'ancien COEFF MARGE
  // On vérifie d'abord que c'est bien COEFF MARGE
  const newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col9Header = String(newHeaders[8] || '').trim().toUpperCase();
  if (col9Header === 'COEFF MARGE') {
    sheet.deleteColumn(9);
  }

  // 5. Formater la nouvelle colonne TAXES
  const maxRows = sheet.getMaxRows();
  sheet.getRange(1, 6, maxRows, 1).setNumberFormat('#,##0.00');

  return { migrated: true, reason: 'success' };
}
