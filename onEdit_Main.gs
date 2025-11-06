function onEdit(e) {
  const sh = e && e.source && e.source.getActiveSheet();
  if (!sh || !e.range || e.range.getRow() === 1) return;
  const name = sh.getName();
  if (name === "Achats") return handleAchats(e);
  if (name === "Stock")  return handleStock(e);
}

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

function enforceChronologicalDates_(sheet, row, cols, options) {
  const opts = options || {};
  const labels = Object.assign({
    dms: 'DATE DE MISE EN STOCK',
    dmis: 'DATE DE MISE EN LIGNE',
    dpub: 'DATE DE PUBLICATION',
    dvente: 'DATE DE VENTE'
  }, opts.labels);

  const order = [
    { key: 'dms',   col: cols && cols.dms },
    { key: 'dmis',  col: cols && cols.dmis },
    { key: 'dpub',  col: cols && cols.dpub },
    { key: 'dvente', col: cols && cols.dvente }
  ];

  const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Etc/GMT';
  const format = date => Utilities.formatDate(date, tz, 'dd/MM/yyyy');

  let previous = null;
  let previousKey = null;
  const values = Object.create(null);
  const missingKeys = [];

  for (let i = 0; i < order.length; i++) {
    const entry = order[i];
    if (!entry.col) continue;

    const cell = sheet.getRange(row, entry.col);
    const value = getDateOrNull_(cell.getValue());
    values[entry.key] = value;

    if (!value) {
      if (opts.requireAllDates) {
        missingKeys.push(entry.key);
      }
      continue;
    }

    if (previous && value.getTime() < previous.getTime()) {
      const labelPrev = labels[previousKey] || previousKey || 'date précédente';
      const labelCur = labels[entry.key] || entry.key || 'date suivante';
      return {
        ok: false,
        message: `${labelCur} (${format(value)}) ne peut pas être antérieure à ${labelPrev} (${format(previous)}).`,
        conflict: { earlier: previousKey, later: entry.key },
        values
      };
    }

    previous = value;
    previousKey = entry.key;
  }

  if (opts.requireAllDates && missingKeys.length > 0) {
    const missingLabels = missingKeys.map(key => labels[key] || key).join(', ');
    return {
      ok: false,
      message: `Impossible de valider : renseignez ${missingLabels}.`,
      missing: missingKeys,
      values
    };
  }

  return { ok: true, values };
}

function extractSkuBase_(sku) {
  const parts = String(sku || "").trim().split('-');
  if (parts.length < 2) return "";
  return parts.slice(0, parts.length - 1).join('-');
}

function buildBaseToStockDate_(ss) {
  const achats = ss && ss.getSheetByName('Achats');
  if (!achats) return Object.create(null);

  const lastA = achats.getLastRow();
  if (lastA < 2) return Object.create(null);

  const achatsHeaders = achats.getRange(1, 1, 1, achats.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(achatsHeaders);
  const colWhere = resolver.colWhere.bind(resolver);
  const colExact = resolver.colExact.bind(resolver);

  const COL_REF = colExact('reference') || colWhere(h => h.includes('reference'));
  const COL_STP = colWhere(h => h.includes('mis en stock')) || colWhere(h => h.includes('mise en stock'));
  if (!COL_REF || !COL_STP) return Object.create(null);

  const refVals = achats.getRange(2, COL_REF, lastA - 1, 1).getValues();
  const stampVals = achats.getRange(2, COL_STP, lastA - 1, 1).getValues();

  const map = Object.create(null);
  for (let i = 0; i < refVals.length; i++) {
    const base = String(refVals[i][0] || "").trim();
    const dt = stampVals[i][0];
    if (!base) continue;
    if (dt instanceof Date && !isNaN(dt)) {
      map[base] = dt;
    }
  }
  return map;
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

  const COL_ID  = colExact('id');
  const COL_REF = colExact('reference') || colWhere(h => h.includes('reference'));
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
    cell.setValue('⚠️ Vous devez obligatoirement fournir un prix de vente');
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

// === ACHATS → synchro + ajout lignes dans Stock ===
//
//  - Si on modifie F (REFERENCE) → met à jour le préfixe des SKU dans Stock.
//  - Si on modifie V (DATE DE MISE EN STOCK) → met à jour la date de mise en stock dans Stock.
//  - Si on coche U (PRÊT POUR MISE EN STOCK) → crée les lignes dans Stock.

function handleAchats(e) {
  const sh = e.source.getActiveSheet();
  const ss = e.source;
  const col = e.range.getColumn();
  const row = e.range.getRow();

  const achatsHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(achatsHeaders);
  const colWhere = resolver.colWhere.bind(resolver);
  const colExact = resolver.colExact.bind(resolver);

  const COL_ID   = colExact('id');
  const COL_ART  = colWhere(h => h.includes('article'));
  const COL_MAR  = colWhere(h => h.includes('marque'));
  const COL_GEN_DATA = colExact('genre(data)')
    || colExact('genre data')
    || colWhere(h => h.includes('genre') && h.includes('data'));
  const COL_GEN_LEGACY = colExact('genre') || colWhere(h => h.includes('genre'));
  const COL_GEN  = COL_GEN_DATA || (COL_GEN_LEGACY && COL_GEN_LEGACY !== COL_GEN_DATA ? COL_GEN_LEGACY : 0);
  const COL_REF  = colExact('reference') || colWhere(h => h.includes('reference'));
  const COL_DLIV = colWhere(h => h.includes('livraison'));
  const COL_QTY  = colWhere(h => h.includes('quantite') && (h.includes('recu') || h.includes('recue')));
  const COL_READY= colWhere(h => h.includes('pret') && h.includes('stock'));
  const COL_STP  = colWhere(h => h.includes('mis en stock')) || colWhere(h => h.includes('mise en stock'));

  // -------------------------
  // 0) MODIF REFERENCE (F) → MAJ PRÉFIXE DES SKU DANS STOCK
  // -------------------------
  if (COL_REF && col === COL_REF) {
    const oldBase = String(e.oldValue || "").trim();
    const newBase = String(e.value     || "").trim();

    // Si pas d’ancienne valeur ou pas de nouvelle, ou identiques → rien
    if (!oldBase || !newBase || oldBase === newBase) return;

    const stock = ss.getSheetByName("Stock");
    if (!stock) return;

    const headersS = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
    const resolverS = makeHeaderResolver_(headersS);
    const C_SKU_STOCK = resolverS.colExact("sku") || resolverS.colExact("reference");
    if (!C_SKU_STOCK) return;

    const lastS = stock.getLastRow();
    if (lastS < 2) return;

    const skuVals = stock.getRange(2, C_SKU_STOCK, lastS - 1, 1).getValues();
    const newSkuVals = [];

    const prefixOld = oldBase + '-';
    const prefixNew = newBase + '-';

    for (let i = 0; i < skuVals.length; i++) {
      let s = String(skuVals[i][0] || "").trim();
      if (!s) {
        newSkuVals.push([s]);
        continue;
      }

      // Cas standard: oldBase-numero → newBase-numero
      if (s.indexOf(prefixOld) === 0) {
        const suffix = s.substring(prefixOld.length); // garde le -numéro
        s = prefixNew + suffix;
        newSkuVals.push([s]);
        continue;
      }

      // Cas rare: SKU == oldBase seul
      if (s === oldBase) {
        newSkuVals.push([newBase]);
        continue;
      }

      // Autres cas → inchangé
      newSkuVals.push([s]);
    }

    stock.getRange(2, C_SKU_STOCK, lastS - 1, 1).setValues(newSkuVals);
    // On ne touche pas aux numéros, juste le préfixe.
    return;
  }

  // -------------------------
  // 1) ÉDITION DE LA COLONNE V (DATE DE MISE EN STOCK) → SYNC VERS STOCK
  // -------------------------
  if (COL_STP && col === COL_STP) {
    // Base de SKU (Achats!F)
    if (!COL_REF) return;
    const refBase = String(sh.getRange(row, COL_REF).getDisplayValue() || "").trim();
    if (!refBase) return;

    // Nouvelle valeur de date saisie en V (on accepte vraie Date ou texte dd/MM/yyyy)
    const cell = sh.getRange(row, COL_STP);
    const rawVal = cell.getValue();
    let dms = null;

    if (rawVal instanceof Date && !isNaN(rawVal)) {
      dms = rawVal;
    } else {
      const s = cell.getDisplayValue();
      if (s) {
        const m = s.match(/^\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\s*$/);
        if (m) {
          const d  = +m[1];
          const mo = +m[2];
          const y  = +(m[3].length === 2 ? ("20" + m[3]) : m[3]);
          dms = new Date(y, mo - 1, d);
        } else {
          dms = null;
        }
      } else {
        dms = null;
      }
    }

    const stock = ss.getSheetByName("Stock");
    if (!stock) return;

    const headersS = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
    const resolverS = makeHeaderResolver_(headersS);
    const C_SKU_STOCK = resolverS.colExact("sku") || resolverS.colExact("reference");
    const C_DMS_STOCK = resolverS.colExact("date de mise en stock"); // ta colonne E
    if (!C_SKU_STOCK || !C_DMS_STOCK) return;

    const lastS = stock.getLastRow();
    if (lastS < 2) return;

    const skuVals = stock.getRange(2, C_SKU_STOCK, lastS - 1, 1).getValues();
    const dmsVals = stock.getRange(2, C_DMS_STOCK, lastS - 1, 1).getValues();

    for (let i = 0; i < skuVals.length; i++) {
      const base = extractSkuBase_(skuVals[i][0]);
      if (!base) continue;

      if (base === refBase) {
        dmsVals[i][0] = dms; // peut être Date ou null
      }
    }

    stock.getRange(2, C_DMS_STOCK, lastS - 1, 1).setValues(dmsVals);
    return;
  }

  // -------------------------
  // 2) CASE U "PRÊT POUR MISE EN STOCK" → CREATION LIGNES DANS STOCK
  // -------------------------
  if (!COL_READY || col !== COL_READY) return; // pas U → on sort

  const turnedOn = (e.value === "TRUE") || (e.value === true);
  if (!turnedOn) return;

  if (!COL_STP) return;
  const stpCell = sh.getRange(row, COL_STP);
  const stockStampDisplay = stpCell.getDisplayValue();
  const stockStampRaw = stpCell.getValue();

  const achatId = COL_ID ? sh.getRange(row, COL_ID).getValue() : "";
  const article = COL_ART ? String(sh.getRange(row, COL_ART).getDisplayValue() || "").trim() : "";
  const marque  = COL_MAR ? String(sh.getRange(row, COL_MAR).getDisplayValue() || "").trim() : "";
  const genrePrimary = COL_GEN_DATA
    ? String(sh.getRange(row, COL_GEN_DATA).getDisplayValue() || "").trim()
    : "";
  const fallbackGenreCol = (!genrePrimary && COL_GEN_LEGACY && COL_GEN_LEGACY !== COL_GEN_DATA)
    ? COL_GEN_LEGACY
    : 0;
  const genreFallback = fallbackGenreCol
    ? String(sh.getRange(row, fallbackGenreCol).getDisplayValue() || "").trim()
    : "";
  const genre   = genrePrimary || genreFallback;
  if (!COL_REF || !COL_QTY) return;
  const skuBase = String(sh.getRange(row, COL_REF).getDisplayValue() || "").trim();
  const qty     = Number(sh.getRange(row, COL_QTY).getValue());
  if (!skuBase || !Number.isFinite(qty) || qty <= 0) return;

  // Date de livraison robuste
  if (!COL_DLIV) return;
  const raw = sh.getRange(row, COL_DLIV).getValue();
  let dateLiv;
  if (raw instanceof Date && !isNaN(raw)) {
    dateLiv = raw;
  } else {
    const s = sh.getRange(row, COL_DLIV).getDisplayValue();
    const m = s && s.match(/^\s*(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\s*$/);
    if (!m) return;
    const d = +m[1], mo = +m[2], y = +(m[3].length === 2 ? ("20"+m[3]) : m[3]);
    dateLiv = new Date(y, mo - 1, d);
  }

  const target = ss.getSheetByName("Stock");
  if (!target) return;

  // Repère dynamiquement les colonnes de Stock
  const headersStock = target.getRange(1, 1, 1, Math.max(4, target.getLastColumn())).getValues()[0];
  const resolverStock = makeHeaderResolver_(headersStock);

  const COL_ID_STOCK    = resolverStock.colExact('id');
  const COL_LABEL_STOCK = resolverStock.colWhere(h => h.includes('libell')) || resolverStock.colWhere(h => h.includes('article')) || 2;
  const COL_OLD_STOCK   = resolverStock.colExact('sku(ancienne nomenclature)');
  const COL_SKU_STOCK   = resolverStock.colExact('sku')
    || resolverStock.colExact('reference')
    || resolverStock.colWhere(h => h.includes('sku'))
    || 3;
  const COL_DATE_STOCK  = resolverStock.colWhere(h => h.includes('livraison')) || (COL_SKU_STOCK ? COL_SKU_STOCK + 1 : 0);
  const C_DMS_STOCK     = resolverStock.colExact('date de mise en stock'); // optionnel

  const base = skuBase;
  const label = `${article} ${marque} ${genre}`.trim();

  const lastExistingStockRow = target.getLastRow();
  let existingStockHasBase = false;
  let existingStockDms = null;
  if (lastExistingStockRow >= 2 && COL_SKU_STOCK) {
    const existingSkuValues = target.getRange(2, COL_SKU_STOCK, lastExistingStockRow - 1, 1).getValues();
    let existingDmsValues = null;
    if (C_DMS_STOCK) {
      existingDmsValues = target.getRange(2, C_DMS_STOCK, lastExistingStockRow - 1, 1).getValues();
    }
    let existingIdValues = null;
    if (COL_ID_STOCK) {
      existingIdValues = target.getRange(2, COL_ID_STOCK, lastExistingStockRow - 1, 1).getValues();
    }
    const prefix = `${base}-`;
    const achatIdKey = (achatId === null || achatId === undefined || achatId === '') ? '' : String(achatId);
    for (let i = 0; i < existingSkuValues.length; i++) {
      const rawSku = String(existingSkuValues[i][0] || "").trim();
      if (!rawSku || rawSku.indexOf(prefix) !== 0) continue;

      let idMatches = true;
      let storedIdKey = '';
      if (COL_ID_STOCK && existingIdValues) {
        const storedRaw = existingIdValues[i] && existingIdValues[i][0];
        storedIdKey = (storedRaw === null || storedRaw === undefined || storedRaw === '') ? '' : String(storedRaw);
        if (achatIdKey && storedIdKey) {
          idMatches = (storedIdKey === achatIdKey);
        }
      }

      if (!idMatches) {
        continue;
      }

      existingStockHasBase = true;

      if (existingDmsValues && !existingStockDms && COL_ID_STOCK && achatIdKey && storedIdKey && storedIdKey === achatIdKey) {
        const candidate = existingDmsValues[i][0];
        if (candidate instanceof Date && !isNaN(candidate)) {
          existingStockDms = candidate;
        }
      }
    }
  }

  if (existingStockHasBase) {
    if (COL_STP && !getDateOrNull_(stockStampRaw)) {
      const fallbackDms = existingStockDms || getDateOrNull_(stockStampDisplay);
      if (fallbackDms) {
        stpCell.setValue(fallbackDms);
      }
    }
    renumberStockByBrand_();
    return;
  }

  // Date de mise en stock : on fixe maintenant ET on la garde dans Achats!V
  let miseStockDate = getDateOrNull_(stockStampRaw);
  if (!miseStockDate) {
    miseStockDate = getDateOrNull_(stockStampDisplay);
  }
  if (!miseStockDate) {
    miseStockDate = new Date();
  }
  if (!(stockStampRaw instanceof Date) || isNaN(stockStampRaw)) {
    stpCell.setValue(miseStockDate);
  }

  const width = Math.max(target.getLastColumn(), COL_LABEL_STOCK || 0, COL_SKU_STOCK || 0, COL_DATE_STOCK || 0, COL_ID_STOCK || 0, COL_OLD_STOCK || 0);
  const rows = Array.from({length: qty}, () => Array(Math.max(1, width)).fill(""));

  for (let i = 0; i < rows.length; i++) {
    const rowValues = rows[i];
    if (COL_ID_STOCK) rowValues[COL_ID_STOCK - 1] = achatId;
    if (COL_LABEL_STOCK) rowValues[COL_LABEL_STOCK - 1] = label;
    if (COL_OLD_STOCK) rowValues[COL_OLD_STOCK - 1] = "";
    if (COL_SKU_STOCK) rowValues[COL_SKU_STOCK - 1] = `${base}-0`;
    if (COL_DATE_STOCK) rowValues[COL_DATE_STOCK - 1] = dateLiv;
  }

  const start = Math.max(2, target.getLastRow() + 1);
  target.getRange(start, 1, rows.length, rows[0].length).setValues(rows);

  if (C_DMS_STOCK) {
    target.getRange(start, C_DMS_STOCK, rows.length, 1).setValue(miseStockDate);
  }

  const lastS = target.getLastRow();
  if (lastS > 2 && COL_DATE_STOCK) {
    target.getRange(2, 1, lastS - 1, target.getLastColumn())
          .sort({ column: COL_DATE_STOCK, ascending: true });
    target.getRange(2, COL_DATE_STOCK, lastS - 1, 1).setNumberFormat("dd/MM/yyyy");
  }

  renumberStockByBrand_();
}

// === RE-NUMÉROTATION GLOBALE PAR BASE DE SKU AVEC OVERRIDE PAR B ===
//
// - Base = toutes les parties avant le dernier "-" du SKU actuel (col "SKU").
// - Si B (SKU ancienne) contient un nombre en fin, on utilise ce nombre pour ce produit.
// - Sinon, on numérote en continu pour cette base à partir de 1.
// - Pas de zéros en tête: suffixe = "1", "2", "3", ...
// - Paramètre onlyOld = true → on ne renumérote QUE les lignes où B est rempli.

function renumberStockByBrand_(onlyOld) {
  const ss = SpreadsheetApp.getActive();
  const stock  = ss.getSheetByName('Stock');
  if (!stock) return;

  const last = stock.getLastRow();
  if (last < 2) return;

  onlyOld = !!onlyOld;

  const stockHeaders = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(stockHeaders);
  const COL_ID    = resolver.colExact('id');
  const COL_OLD   = resolver.colExact("sku(ancienne nomenclature)") || resolver.colWhere(h => h.includes('ancienne')) || 2; // B
  const COL_NEW   = resolver.colExact("sku")
    || resolver.colExact("reference")
    || resolver.colWhere(h => h.includes('sku'))
    || 3; // C

  const width = Math.max(COL_NEW, COL_OLD, COL_ID || 0, stock.getLastColumn());
  const data = stock.getRange(2, 1, last - 1, width).getValues();

  const idToBase = COL_ID ? buildIdToSkuBaseMap_(ss) : null;
  const baseCounters = Object.create(null);
  const newSkuColValues = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    const oldSku = String(row[COL_OLD - 1] || "").trim();  // SKU ancienne
    let  curSku = String(row[COL_NEW - 1] || "").trim();   // SKU actuelle (base-0 ou autre)
    const idRaw = (COL_ID ? row[COL_ID - 1] : '') ;
    const idKey = idRaw === null || idRaw === undefined || idRaw === '' ? '' : String(idRaw);

    // Si on est en mode "onlyOld" et qu'il n'y a rien en B → on laisse le SKU tel quel
    if (onlyOld && !oldSku) {
      newSkuColValues.push([curSku]);
      continue;
    }

    const idBase = (idKey && idToBase && idToBase[idKey]) ? idToBase[idKey] : '';
    const curBase = extractSkuBase_(curSku);
    const oldBase = extractSkuBase_(oldSku);

    let base = curBase || oldBase || idBase;
    if (base && idBase && base !== idBase) {
      const curAligned = !curBase || curBase === idBase;
      const oldAligned = !oldBase || oldBase === idBase;
      if (curAligned && oldAligned) {
        base = idBase;
      }
    }

    if (!base) {
      newSkuColValues.push([curSku]);
      continue;
    }

    if (!curSku || curSku.indexOf(base + '-') !== 0) {
      curSku = base + '-0';
    }

    const counterKey = base;
    if (!Object.prototype.hasOwnProperty.call(baseCounters, counterKey)) {
      let initialCounter = 0;
      const curMatch = String(curSku).match(/-(\d+)\s*$/);
      if (curMatch) {
        const parsed = parseInt(curMatch[1], 10);
        if (!isNaN(parsed)) {
          initialCounter = parsed;
        }
      }
      baseCounters[counterKey] = initialCounter;
    }

    // extraction éventuelle du numéro dans SKU(ancienne)
    let overrideNum = null;
    if (oldSku) {
      const overrideBase = extractSkuBase_(oldSku);
      if (!overrideBase || overrideBase === base) {
        const m = String(oldSku).match(/(\d+)\s*$/);
        if (m) overrideNum = parseInt(m[1], 10);
      }
    }

    let suffix;
    if (overrideNum != null && Number.isFinite(overrideNum) && overrideNum > 0) {
      suffix = overrideNum;
      baseCounters[counterKey] = suffix;
    } else {
      suffix = baseCounters[counterKey] + 1;
      baseCounters[counterKey] = suffix;
    }

    const newSku = base + '-' + suffix; // sans padding
    newSkuColValues.push([newSku]);
  }

  stock.getRange(2, COL_NEW, newSkuColValues.length, 1).setValues(newSkuColValues);
}

// === STOCK → horodatages + validations / déplacement vers Ventes ===

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

  const C_ID      = colExact('id');
  const C_LABEL   = colWhere(h => h.includes('libell')) || colWhere(h => h.includes('article')) || 2;
  const C_OLD_SKU = colExact("sku(ancienne nomenclature)") || 2;
  const C_SKU     = colExact("sku") || colExact("reference"); // B/C
  const C_PRIX    = colWhere(h => h.includes("prix") && h.includes("vente")); // "PRIX DE VENTE"
  const C_DMS     = colExact("date de mise en stock");
  const C_MIS     = colExact("mis en ligne");
  const C_DMIS    = colExact("date de mise en ligne");
  const C_PUB     = colExact("publié");
  const C_DPUB    = colExact("date de publication");
  const C_VENDU   = colExact("vendu");
  let   C_DVENTE  = colExact("date de vente");
  if (!C_DVENTE) C_DVENTE = 10;
  const C_STAMPV  = colExact("vente exportee le");
  const C_VALIDE  = colWhere(h => h.includes("valider") && h.includes("saisie"));

  const c = e.range.getColumn(), r = e.range.getRow();

  const chronoCols = {
    dms: C_DMS,
    dmis: C_DMIS,
    dpub: C_DPUB,
    dvente: C_DVENTE
  };

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
    const vendu = C_VENDU ? (sh.getRange(r, C_VENDU).getValue() === true) : false;
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

  if (c === C_DMS || c === C_DMIS || c === C_DPUB || c === C_DVENTE) {
    const key = c === C_DMS ? 'dms' : (c === C_DMIS ? 'dmis' : (c === C_DPUB ? 'dpub' : 'dvente'));
    if (!ensureChronologyOrRevert_(key, e.oldValue)) {
      return;
    }
    if (c !== C_DVENTE) {
      return;
    }
  }

  // 1) MIS EN LIGNE → horodate
  if (C_MIS && C_DMIS && c === C_MIS) {
    if (turnedOff) {
      if (C_PUB && sh.getRange(r, C_PUB).getValue() === true) {
        sh.getRange(r, C_MIS).setValue(true);
        ss.toast('Impossible de décocher "MIS EN LIGNE" tant que "PUBLIÉ" est coché.', 'Stock', 5);
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
      cell.setValue(new Date());
      const checkboxInfo = { range: sh.getRange(r, C_MIS), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dmis', null, checkboxInfo)) {
        return;
      }
      return;
    }

    return;
  }

  // 2) PUBLIÉ → horodate
  if (C_PUB && C_DPUB && c === C_PUB) {
    if (turnedOff) {
      const vendu = C_VENDU ? (sh.getRange(r, C_VENDU).getValue() === true) : false;
      if (vendu) {
        sh.getRange(r, C_PUB).setValue(true);
        ss.toast('Impossible de décocher "PUBLIÉ" lorsqu\'une vente est cochée.', 'Stock', 5);
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
      cell.setValue(new Date());
      const checkboxInfo = { range: sh.getRange(r, C_PUB), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dpub', null, checkboxInfo)) {
        return;
      }
      return;
    }

    return;
  }

  // 3) VENDU → horodatage + alerte prix, déplacement seulement via "Valider la saisie"
  if (C_VENDU && c === C_VENDU) {
    const dv = sh.getRange(r, C_DVENTE);

    if (turnedOn) {
      storePreviousCellValue_(sh, r, C_DVENTE, dv.getValue());
      if (C_PRIX) {
        const priceCell = sh.getRange(r, C_PRIX);
        storePreviousCellValue_(sh, r, C_PRIX, priceCell.getValue());
      }

      let val = dv.getValue();
      if (!(val instanceof Date) || isNaN(val)) {
        dv.setValue(new Date());  // on date au moment du clic
      }
      const checkboxInfo = { range: sh.getRange(r, C_VENDU), oldValue: e.oldValue };
      if (!ensureChronologyOrRevert_('dvente', null, checkboxInfo)) {
        if (C_PRIX) {
          restorePreviousCellValue_(sh, r, C_PRIX);
        }
        return;
      }
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
        priceCell.setBackground(null);
        priceCell.setFontColor(null);
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
  if (c === C_DVENTE) {
    const val = sh.getRange(r, C_DVENTE).getValue();
    const isDate = val instanceof Date && !isNaN(val.getTime());
    if (!isDate) return;

    const vendu = C_VENDU ? (sh.getRange(r, C_VENDU).getValue() === true) : false;
    if (!vendu) return;

    ensureValidPriceOrWarn_(sh, r, C_PRIX);
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

    const vendu = C_VENDU ? (sh.getRange(r, C_VENDU).getValue() === true) : false;
    if (!vendu) {
      return;
    }

    if (!ensureValidPriceOrWarn_(sh, r, C_PRIX)) return;

    const valDate = sh.getRange(r, C_DVENTE).getValue();
    if (!(valDate instanceof Date) || isNaN(valDate.getTime())) return;

    const baseToDmsMap = buildBaseToStockDate_(ss);
    exportVente_(null, r, C_ID, C_LABEL, C_SKU, C_PRIX, C_DVENTE, C_STAMPV, baseToDmsMap);
    return;
  }
}

// Déplace la ligne de "Stock" vers "Ventes" (et calcule les délais)
function exportVente_(e, row, C_ID, C_LABEL, C_SKU, C_PRIX, C_DVENTE, C_STAMPV, baseToDmsMap) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Stock");
  if (!sh) return;

  const ventes = ss.getSheetByName("Ventes") || ss.insertSheet("Ventes");
  if (ventes.getLastRow() === 0) {
    ventes.getRange(1, 1, 1, 9).setValues([[
      "ID",
      "DATE DE VENTE",
      "ARTICLE",
      "SKU",
      "PRIX VENTE",
      "DÉLAI D'IMMOBILISATION",
      "DELAI DE MISE EN LIGNE",
      "DELAI DE PUBLICATION",
      "DELAI DE VENTE"
    ]]);
  }

  const ventesHeaders = ventes.getRange(1, 1, 1, Math.max(9, ventes.getLastColumn())).getValues()[0];
  const ventesResolver = makeHeaderResolver_(ventesHeaders);
  const ventesExact = ventesResolver.colExact.bind(ventesResolver);
  const ventesWhere = ventesResolver.colWhere.bind(ventesResolver);

  const COL_ID_VENTE    = ventesExact('id');
  const COL_DATE_VENTE  = ventesWhere(h => h.includes('date') && h.includes('vente'));
  const COL_ARTICLE     = ventesWhere(h => h.includes('article')) || ventesWhere(h => h.includes('libell'));
  const COL_SKU_VENTE   = ventesExact('sku');
  const COL_PRIX_VENTE  = ventesWhere(h => h.includes('prix') && h.includes('vente'));
  const COL_DELAI_IMM   = ventesWhere(h => h.includes('immobilisation'));
  const COL_DELAI_ML    = ventesWhere(h => h.includes('mise en ligne'));
  const COL_DELAI_PUB   = ventesWhere(h => h.includes('publication'));
  const COL_DELAI_VENTE = ventesWhere(h => h.includes('delai') && h.includes('vente'));
  const widthVentes     = Math.max(ventes.getLastColumn(), 9);

  const dateCell = sh.getRange(row, C_DVENTE);
  const dateV = dateCell.getValue();
  if (!(dateV instanceof Date) || isNaN(dateV)) return;

  const idVal = C_ID ? sh.getRange(row, C_ID).getValue() : "";
  const label = C_LABEL ? sh.getRange(row, C_LABEL).getDisplayValue() : "";
  const sku   = C_SKU  ? sh.getRange(row, C_SKU).getDisplayValue() : "";
  const prix  = C_PRIX ? sh.getRange(row, C_PRIX).getValue() : "";

  const headersS = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const resolverS = makeHeaderResolver_(headersS);

  const C_DMS   = resolverS.colExact("date de mise en stock");
  const C_DMIS  = resolverS.colExact("date de mise en ligne");
  const C_DPUB  = resolverS.colExact("date de publication");

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

  const start = Math.max(2, ventes.getLastRow() + 1);
  const newRow = Array(widthVentes).fill("");
  if (COL_ID_VENTE) newRow[COL_ID_VENTE - 1] = idVal;
  if (COL_DATE_VENTE) newRow[COL_DATE_VENTE - 1] = dateV;
  if (COL_ARTICLE) newRow[COL_ARTICLE - 1] = label;
  if (COL_SKU_VENTE) newRow[COL_SKU_VENTE - 1] = sku;
  if (COL_PRIX_VENTE) newRow[COL_PRIX_VENTE - 1] = prix;
  if (COL_DELAI_IMM) newRow[COL_DELAI_IMM - 1] = delaiImm;
  if (COL_DELAI_ML) newRow[COL_DELAI_ML - 1] = delaiML;
  if (COL_DELAI_PUB) newRow[COL_DELAI_PUB - 1] = delaiPub;
  if (COL_DELAI_VENTE) newRow[COL_DELAI_VENTE - 1] = delaiVte;

  ventes.getRange(start, 1, 1, newRow.length).setValues([newRow]);

  const lastV = ventes.getLastRow();
  if (lastV > 2 && COL_DATE_VENTE) {
    ventes.getRange(2, 1, lastV - 1, ventes.getLastColumn()).sort([{column: COL_DATE_VENTE, ascending: false}]);
    ventes.getRange(2, COL_DATE_VENTE, lastV - 1, 1).setNumberFormat('dd/MM/yyyy');
  }

  if (C_STAMPV) sh.getRange(row, C_STAMPV).setValue(new Date());

  sh.deleteRow(row);
}

// === MENU ===

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Maintenance')
    .addItem('Recalculer les SKU du Stock', 'recalcStock')
    .addItem('Mettre à jour les dates de mise en stock', 'syncMiseEnStockFromAchats')
    .addItem('Valider toutes les saisies prêtes', 'validateAllSales')
    .addItem('Trier les ventes (date décroissante)', 'sortVentesByDate')
    .addItem('Retirer du Stock les ventes importées', 'purgeStockFromVentes')
    .addToUi();
}

function sortVentesByDate() {
  const ss = SpreadsheetApp.getActive();
  const ventes = ss.getSheetByName('Ventes');
  if (!ventes) {
    ss.toast('Feuille Ventes introuvable', 'Tri des ventes', 5);
    return;
  }

  const lastRow = ventes.getLastRow();
  if (lastRow <= 2) {
    ss.toast('Aucune donnée à trier', 'Tri des ventes', 5);
    return;
  }

  const lastColumn = ventes.getLastColumn();
  const ventesHeaders = ventes.getRange(1, 1, 1, lastColumn).getValues()[0];
  const resolver = makeHeaderResolver_(ventesHeaders);
  const colDate = resolver.colWhere(h => h.includes('date') && h.includes('vente')) || 2;

  ventes
    .getRange(2, 1, lastRow - 1, lastColumn)
    .sort({ column: colDate, ascending: false });
  ventes.getRange(2, colDate, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy');

  ss.toast('Les ventes ont été triées par date décroissante.', 'Tri des ventes', 5);
}

// Recalcul des SKU uniquement dans Stock
function recalcStock() {
  const ss = SpreadsheetApp.getActive();
  const stock  = ss.getSheetByName('Stock');
  if (!stock) {
    SpreadsheetApp.getActive().toast('Feuille Stock introuvable', 'Recalcul SKU', 5);
    return;
  }

  const last = stock.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getActive().toast('Aucune ligne dans Stock', 'Recalcul SKU', 5);
    return;
  }

  const stockHeaders = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(stockHeaders);

  let C_DATE = resolver.colExact("date de livraison");
  if (!C_DATE) C_DATE = 4;

  const width = stock.getLastColumn();

  stock.getRange(2, 1, last - 1, width)
       .sort({ column: C_DATE, ascending: true });
  stock.getRange(2, C_DATE, last - 1, 1)
       .setNumberFormat('dd/MM/yyyy');

  SpreadsheetApp.getActive().toast(
    'Tri du stock terminé (aucune renumérotation de SKU effectuée).',
    'Recalcul SKU',
    5
  );
}

// Met à jour "DATE DE MISE EN STOCK" dans Stock à partir de Achats!V
function syncMiseEnStockFromAchats() {
  const ss = SpreadsheetApp.getActive();
  const achats = ss.getSheetByName('Achats');
  const stock  = ss.getSheetByName('Stock');
  if (!achats || !stock) {
    SpreadsheetApp.getActive().toast('Feuilles Achats/Stock introuvables', 'Mise à jour DMS', 5);
    return;
  }

  const lastA = achats.getLastRow();
  if (lastA < 2) {
    SpreadsheetApp.getActive().toast('Aucune donnée dans Achats', 'Mise à jour DMS', 5);
    return;
  }

  const mapBaseToDMS = buildBaseToStockDate_(ss);

  const lastS = stock.getLastRow();
  if (lastS < 2) {
    SpreadsheetApp.getActive().toast('Aucune donnée dans Stock', 'Mise à jour DMS', 5);
    return;
  }

  const headersS = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
  const resolverS = makeHeaderResolver_(headersS);

  const C_SKU  = resolverS.colExact("sku") || resolverS.colExact("reference");
  const C_DMS  = resolverS.colExact("date de mise en stock");
  if (!C_SKU || !C_DMS) {
    SpreadsheetApp.getActive().toast('Colonnes SKU ou "DATE DE MISE EN STOCK" introuvables', 'Mise à jour DMS', 5);
    return;
  }

  const skuVals = stock.getRange(2, C_SKU, lastS - 1, 1).getValues();
  const dmsRange = stock.getRange(2, C_DMS, lastS - 1, 1);
  const dmsVals = dmsRange.getValues();

  let updated = 0;
  let cleared = 0;
  for (let i = 0; i < skuVals.length; i++) {
    const base = extractSkuBase_(skuVals[i][0]);
    if (!base) continue;

    const dt = mapBaseToDMS[base];
    if (dt instanceof Date && !isNaN(dt)) {
      if (!(dmsVals[i][0] instanceof Date) || dmsVals[i][0].getTime() !== dt.getTime()) {
        dmsVals[i][0] = dt;
        updated++;
      }
    } else if (dmsVals[i][0]) {
      dmsVals[i][0] = null;
      cleared++;
    }
  }

  dmsRange.setValues(dmsVals);

  SpreadsheetApp.getActive().toast(
    `Dates de mise en stock mises à jour sur ${updated} ligne(s) et effacées sur ${cleared} ligne(s).`,
    'Mise à jour DMS',
    5
  );
}

function purgeStockFromVentes() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getSheetByName('Stock');
  const ventes = ss.getSheetByName('Ventes');

  if (!stock || !ventes) {
    ss.toast('Feuilles "Stock" ou "Ventes" introuvables.', 'Purge du stock', 6);
    return;
  }

  const stockLast = stock.getLastRow();
  const ventesLast = ventes.getLastRow();
  if (stockLast < 2) {
    ss.toast('Aucune ligne à traiter dans "Stock".', 'Purge du stock', 6);
    return;
  }
  if (ventesLast < 2) {
    ss.toast('Aucune vente disponible pour le rapprochement.', 'Purge du stock', 6);
    return;
  }

  const stockHeaders = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const ventesHeaders = ventes.getRange(1, 1, 1, ventes.getLastColumn()).getValues()[0];
  const stockResolver = makeHeaderResolver_(stockHeaders);
  const ventesResolver = makeHeaderResolver_(ventesHeaders);

  const C_STOCK_ID = stockResolver.colExact('id');
  const C_STOCK_SKU = stockResolver.colExact('sku') || stockResolver.colExact('reference');
  const C_VENTE_ID = ventesResolver.colExact('id');
  const C_VENTE_SKU = ventesResolver.colExact('sku');

  if (!C_STOCK_ID || !C_STOCK_SKU || !C_VENTE_ID || !C_VENTE_SKU) {
    ss.toast('Colonnes ID ou SKU introuvables dans Stock/Ventes.', 'Purge du stock', 8);
    return;
  }

  const ventesWidth = ventes.getLastColumn();
  const ventesValues = ventes.getRange(2, 1, ventesLast - 1, ventesWidth).getValues();
  const venteCounts = new Map();
  let ventesIgnorées = 0;

  function buildKey(idValue, skuValue) {
    const id = idValue === null || idValue === undefined ? '' : String(idValue).trim();
    const sku = skuValue === null || skuValue === undefined ? '' : String(skuValue).trim().toLowerCase();
    if (!id || !sku) return '';
    return id + '|' + sku;
  }

  for (let i = 0; i < ventesValues.length; i++) {
    const row = ventesValues[i];
    const key = buildKey(row[C_VENTE_ID - 1], row[C_VENTE_SKU - 1]);
    if (!key) {
      ventesIgnorées++;
      continue;
    }
    venteCounts.set(key, (venteCounts.get(key) || 0) + 1);
  }

  if (!venteCounts.size) {
    ss.toast('Aucun couple ID+SKU exploitable dans "Ventes".', 'Purge du stock', 6);
    return;
  }

  const stockWidth = stock.getLastColumn();
  const stockValues = stock.getRange(2, 1, stockLast - 1, stockWidth).getValues();
  const rowsToDelete = [];

  for (let i = 0; i < stockValues.length; i++) {
    const row = stockValues[i];
    const key = buildKey(row[C_STOCK_ID - 1], row[C_STOCK_SKU - 1]);
    const count = key ? venteCounts.get(key) : 0;
    if (count && count > 0) {
      rowsToDelete.push(i + 2);
      if (count === 1) {
        venteCounts.delete(key);
      } else {
        venteCounts.set(key, count - 1);
      }
    }
  }

  if (!rowsToDelete.length) {
    ss.toast('Aucune ligne du stock ne correspond aux ventes.', 'Purge du stock', 6);
    return;
  }

  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(row => stock.deleteRow(row));

  const restants = Array.from(venteCounts.values()).reduce((sum, val) => sum + val, 0);
  const messageParts = [`${rowsToDelete.length} ligne(s) supprimée(s) du Stock.`];
  if (ventesIgnorées) {
    messageParts.push(`${ventesIgnorées} vente(s) ignorée(s) (ID ou SKU manquant).`);
  }
  if (restants) {
    messageParts.push(`${restants} vente(s) sans correspondance dans le Stock.`);
  }

  ss.toast(messageParts.join(' '), 'Purge du stock', 8);
}

// Validation groupée
function validateAllSales() {
  const ss = SpreadsheetApp.getActive();
  const stock = ss.getSheetByName('Stock');
  if (!stock) {
    SpreadsheetApp.getUi().alert('Validation groupée', 'Feuille "Stock" introuvable.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const last = stock.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('Validation groupée', 'Aucune ligne à traiter dans "Stock".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const stockHeaders = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const resolver = makeHeaderResolver_(stockHeaders);
  const colExact = resolver.colExact.bind(resolver);
  const colWhere = resolver.colWhere.bind(resolver);

  const C_ID       = colExact('id');
  const C_LABEL    = colWhere(h => h.includes('libell')) || colWhere(h => h.includes('article')) || 2;
  const C_SKU      = colExact("sku") || colExact("reference");
  const C_PRIX     = colWhere(h => h.includes("prix") && h.includes("vente"));
  const C_DVENTE   = colExact("date de vente") || 10;
  const C_VENDU    = colExact("vendu");
  const C_VALIDATE = colWhere(h => h.includes("valider") && h.includes("saisie"));
  const C_DMS      = colWhere(h => h.includes("mise en stock"));
  const C_DMIS     = colExact("date de mise en ligne");
  const C_DPUB     = colExact("date de publication");

  if (!C_SKU || !C_PRIX || !C_DVENTE) {
    SpreadsheetApp.getUi().alert(
      'Validation groupée',
      'Colonnes SKU / PRIX DE VENTE / DATE DE VENTE introuvables. Vérifie les en-têtes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const lastCol = stock.getLastColumn();
  const data = stock.getRange(2, 1, last - 1, lastCol).getValues();
  const baseToDmsMap = buildBaseToStockDate_(ss);

  const ventes = ss.getSheetByName('Ventes') || ss.insertSheet('Ventes');
  if (ventes.getLastRow() === 0) {
    ventes.getRange(1,1,1,9).setValues([[
      "ID",
      "DATE DE VENTE",
      "ARTICLE",
      "SKU",
      "PRIX VENTE",
      "DÉLAI D'IMMOBILISATION",
      "DELAI DE MISE EN LIGNE",
      "DELAI DE PUBLICATION",
      "DELAI DE VENTE"
    ]]);
  }

  const ventesHeaders = ventes.getRange(1, 1, 1, Math.max(9, ventes.getLastColumn())).getValues()[0];
  const ventesResolver = makeHeaderResolver_(ventesHeaders);
  const ventesExact = ventesResolver.colExact.bind(ventesResolver);
  const ventesWhere = ventesResolver.colWhere.bind(ventesResolver);

  const COL_ID_VENTE    = ventesExact('id');
  const COL_DATE_VENTE  = ventesWhere(h => h.includes('date') && h.includes('vente'));
  const COL_ARTICLE     = ventesWhere(h => h.includes('article')) || ventesWhere(h => h.includes('libell'));
  const COL_SKU_VENTE   = ventesExact('sku');
  const COL_PRIX_VENTE  = ventesWhere(h => h.includes('prix') && h.includes('vente'));
  const COL_DELAI_IMM   = ventesWhere(h => h.includes('immobilisation'));
  const COL_DELAI_ML    = ventesWhere(h => h.includes('mise en ligne'));
  const COL_DELAI_PUB   = ventesWhere(h => h.includes('publication'));
  const COL_DELAI_VENTE = ventesWhere(h => h.includes('delai') && h.includes('vente'));
  const widthVentes     = Math.max(ventes.getLastColumn(), 9);

  const toAppend = [];
  const rowsToDel = [];
  let moved = 0;
  const invalidChronoRows = [];

  const chronoCols = {
    dms: C_DMS,
    dmis: C_DMIS,
    dpub: C_DPUB,
    dvente: C_DVENTE
  };

  const msPerDay = 1000 * 60 * 60 * 24;
  const daysDiff = (d2, d1) => {
    if (!(d2 instanceof Date) || isNaN(d2) || !(d1 instanceof Date) || isNaN(d1)) return "";
    return Math.round((d2.getTime() - d1.getTime()) / msPerDay);
  };

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const rowIndex = i + 2;

    const validateOk = C_VALIDATE ? (row[C_VALIDATE - 1] === true) : true;
    if (!validateOk) continue;

    const vendu = C_VENDU ? (row[C_VENDU - 1] === true) : false;
    if (!vendu) continue;

    const dateVente = row[C_DVENTE - 1];
    if (!(dateVente instanceof Date) || isNaN(dateVente)) continue;

    const prix = row[C_PRIX - 1];
    const prixOk = (typeof prix === 'number' && !isNaN(prix) && prix > 0);
    if (!prixOk) {
      ensureValidPriceOrWarn_(stock, rowIndex, C_PRIX);
      continue;
    }

    const chronoCheck = enforceChronologicalDates_(stock, rowIndex, chronoCols, { requireAllDates: true });
    if (!chronoCheck.ok) {
      if (C_VALIDATE) {
        stock.getRange(rowIndex, C_VALIDATE).setValue(false);
      }
      invalidChronoRows.push({ row: rowIndex, message: chronoCheck.message });
      continue;
    }

    const idValue = C_ID ? row[C_ID - 1] : '';
    const label = row[C_LABEL - 1];
    const sku   = C_SKU ? row[C_SKU - 1] : "";

    let dateMiseStock = C_DMS  ? row[C_DMS  - 1] : null;
    if (!(dateMiseStock instanceof Date) || isNaN(dateMiseStock)) {
      const base = extractSkuBase_(sku);
      if (base) {
        const dt = baseToDmsMap[base];
        dateMiseStock = (dt instanceof Date && !isNaN(dt)) ? dt : null;
      } else {
        dateMiseStock = null;
      }
    }
    const dateMiseLigne = C_DMIS ? row[C_DMIS - 1] : null;
    const datePub       = C_DPUB ? row[C_DPUB - 1] : null;

    const dImmobil = daysDiff(dateVente, dateMiseStock);
    const dLigne   = daysDiff(dateMiseLigne, dateMiseStock);
    const dPub     = daysDiff(datePub,  dateMiseLigne);
    const dVente   = daysDiff(dateVente, datePub);

    const newRow = Array(widthVentes).fill('');
    if (COL_ID_VENTE) newRow[COL_ID_VENTE - 1] = idValue;
    if (COL_DATE_VENTE) newRow[COL_DATE_VENTE - 1] = dateVente;
    if (COL_ARTICLE) newRow[COL_ARTICLE - 1] = label;
    if (COL_SKU_VENTE) newRow[COL_SKU_VENTE - 1] = sku;
    if (COL_PRIX_VENTE) newRow[COL_PRIX_VENTE - 1] = prix;
    if (COL_DELAI_IMM) newRow[COL_DELAI_IMM - 1] = dImmobil;
    if (COL_DELAI_ML) newRow[COL_DELAI_ML - 1] = dLigne;
    if (COL_DELAI_PUB) newRow[COL_DELAI_PUB - 1] = dPub;
    if (COL_DELAI_VENTE) newRow[COL_DELAI_VENTE - 1] = dVente;

    toAppend.push(newRow);

    rowsToDel.push(rowIndex);
    moved++;
  }

  if (toAppend.length > 0) {
    const startV = Math.max(2, ventes.getLastRow() + 1);
    ventes.getRange(startV, 1, toAppend.length, widthVentes).setValues(toAppend);

    const lastV = ventes.getLastRow();
    if (lastV > 2 && COL_DATE_VENTE) {
      ventes.getRange(2, 1, lastV - 1, ventes.getLastColumn())
            .sort([{column: COL_DATE_VENTE, ascending: false}]);
      ventes.getRange(2, COL_DATE_VENTE, lastV - 1, 1).setNumberFormat('dd/MM/yyyy');
    }

    rowsToDel.sort((a, b) => b - a);
    rowsToDel.forEach(r => stock.deleteRow(r));
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Validation groupée terminée',
    `${moved} ligne(s) ont été déplacées vers "Ventes".`,
    ui.ButtonSet.OK
  );

  if (invalidChronoRows.length > 0) {
    const first = invalidChronoRows[0];
    const extra = invalidChronoRows.length > 1 ? ` (et ${invalidChronoRows.length - 1} autre(s) ligne(s))` : '';
    const message = `Validation bloquée - ligne ${first.row}: ${first.message}${extra}`;
    SpreadsheetApp.getActive().toast(message, 'Validation groupée', 8);
  }
}
