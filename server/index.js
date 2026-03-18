'use strict';

const express  = require('express');
const multer   = require('multer');
const XLSX     = require('xlsx');
const path     = require('path');
const { calculatePrice, analyzeDevices } = require('./pricing-engine');

const app    = express();
const router = express.Router();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());

// Mount everything under /erpie so nginx proxy_pass /erpie/ → localhost:3001/erpie/ works
router.use(express.static(path.join(__dirname, '../client/public')));
app.use('/erpie', router);


// ═══════════════════════════════════════════════════════════════════════════
// UNIVERSAL DEVICE PARSER v2.0
// Handles: ARS Excel, PWC/Vendor Quote, Frankfurt CSV, Email text, Free text
// ═══════════════════════════════════════════════════════════════════════════

const BRANDS = ['apple','dell','hp','lenovo','microsoft','asus','acer',
                'fujitsu','toshiba','samsung','sony','lg','panasonic',
                'huawei','xiaomi','google','razer','msi','gigabyte'];

const BRAND_MODELS = {
  apple:     ['macbook','imac','mac mini','mac pro','mac studio','iphone','ipad','ipod'],
  dell:      ['latitude','inspiron','xps','precision','vostro','optiplex','alienware'],
  hp:        ['elitebook','probook','zbook','pavilion','envy','spectre','omen',
               'elitedesk','prodesk','z-workstation','chromebook','folio','revolve'],
  lenovo:    ['thinkpad','thinkcentre','ideapad','legion','yoga','tab'],
  microsoft: ['surface'],
};

const PRODUCT_TYPES = new Set([
  'notebook','laptop','desktop','allinone','all-in-one','workstation',
  'mobilephone','mobile','tablet','server','monitor',
  // with spaces stripped:
  'mobilephone','allinone',
]);

function cellStr(v) { return String(v ?? '').trim(); }

function normalizeProductType(s) {
  return s.toLowerCase().replace(/[\s\-\/]/g,'').replace('allone','allinone');
}

function isProductType(v) {
  return PRODUCT_TYPES.has(normalizeProductType(cellStr(v)));
}

function looksLikeBrand(v) {
  const s = cellStr(v).toLowerCase();
  return BRANDS.some(b => s === b || s.startsWith(b));
}

function looksLikeModel(v) {
  const s = cellStr(v).toLowerCase();
  if (s.length < 3) return false;
  // Has a known brand keyword
  for (const brand of BRANDS) {
    if (s.includes(brand)) return true;
  }
  // Has known model keyword
  const modelWords = ['latitude','elitebook','probook','thinkpad','thinkcentre',
    'macbook','surface','optiplex','inspiron','precision','vostro','zbook',
    'yoga','ideapad','legion','spectre','envy','pavilion','folio','revolve'];
  return modelWords.some(m => s.includes(m));
}

// ─── SPEC EXTRACTION ────────────────────────────────────────────────────────

function extractRAM(s) {
  // "32GB", "32 GB", "32GB DDR4", "Memory: 16GB"
  const m = s.match(/\b(4|8|16|32|64|128)\s*GB(?!\s*SSD)/i);
  return m ? m[1] + 'GB' : '';
}

function extractSSD(s) {
  // "512GB SSD", "SSD 512GB", "256 GB", "1TB SSD", "1024GB"
  const KNOWN_SSD_SIZES = new Set([120,128,240,256,480,512,960,1024,2048]);
  const ssdM = s.match(/(?:SSD\s*)?(\d+)\s*(GB|TB)\s*(?:SSD|M\.?2|NVMe|SATA)?/gi);
  if (!ssdM) return '';
  for (const match of ssdM) {
    const numM = match.match(/(\d+)\s*(GB|TB)/i);
    if (!numM) continue;
    const num = parseInt(numM[1]);
    const unit = numM[2].toUpperCase();
    const gb = unit === 'TB' ? num * 1024 : num;
    // Accept known SSD sizes directly
    if (KNOWN_SSD_SIZES.has(gb)) return gb + 'GB';
    // For non-standard sizes (600-999 range), only accept if SSD/NVMe/M.2 keyword is present
    // This prevents model numbers like 835, 855, 640, 650 from being treated as SSD
    if (gb >= 600 && gb < 1000) {
      if (/SSD|NVMe|M\.?2|SATA/i.test(match)) return gb + 'GB';
      continue; // skip — likely a model number
    }
    // Accept other sizes > 1000 (e.g. 2000GB)
    if (gb >= 100) return gb + 'GB';
  }
  return '';
}

function extractCPU(s) {
  // Intel: i3/i5/i7/i9 + model, AMD Ryzen, Apple M1/M2/M3
  const patterns = [
    /\b(Core\s*)?[iI][3579][-\s]?\d{4,5}[A-Z0-9]*/,
    /\bRyzen\s*[3579]\s*\d{4}[A-Z0-9]*/i,
    /\bM[123]\s*(Pro|Max|Ultra)?/,
    /\bCeleron\s*[A-Z0-9]+/i,
    /\bPentium\s*[A-Z0-9]+/i,
    /\bXeon\s*[A-Z0-9\-]+/i,
  ];
  for (const re of patterns) {
    const m = s.match(re);
    if (m) return m[0].trim();
  }
  // Gen hint: "I7-12th" → record as i7 gen12
  const genHint = s.match(/[iI]([3579])[-\s]?(\d+)(th|rd|nd|st)/i);
  if (genHint) return `i${genHint[1]}-${genHint[2]}th`;
  return '';
}

function extractGrade(s) {
  const m = s.match(/\b(A1|A2|A3|A|B1|B2|B3|B4|B|C1|C2|C|D)\b/);
  return m ? m[1] : '';
}

// ─── FORMAT DETECTORS ────────────────────────────────────────────────────────

function detectFormat(wb) {
  // 1. PlanBit ARS — has "Customer Inventory" tab or RITM-prefixed sheets
  if (wb.SheetNames.some(n => n.toLowerCase().includes('inventory')) ||
      wb.SheetNames.some(n => /^RITM/i.test(n.trim()))) {
    return 'ARS';
  }

  // 2. Check first/main sheet
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const nonEmpty = rows.filter(r => r.some(v => cellStr(v).length > 0));
    if (!nonEmpty.length) continue;

    // 3. PWC/Vendor quote: has Product+Brand+Model+Processor in header row
    for (let i = 0; i < Math.min(5, nonEmpty.length); i++) {
      const cells = nonEmpty[i].map(v => cellStr(v).toLowerCase());
      const hasProduct  = cells.some(c => c === 'product');
      const hasBrandCol = cells.some(c => c === 'brand');
      const hasModel    = cells.some(c => c === 'model');
      const hasProc     = cells.some(c => c.includes('processor') || c === 'p/n');
      if (hasProduct && hasBrandCol && hasModel && hasProc) {
        return 'VENDOR_QUOTE';
      }
    }

    // 4. Generic with headers (Frankfurt style, any CSV with model column)
    for (let i = 0; i < Math.min(10, nonEmpty.length); i++) {
      const cells = nonEmpty[i].map(v => cellStr(v).toLowerCase().replace(/[\s_\-]/g,''));
      const score = ['model','device','modelname','computername','description'].filter(k => cells.includes(k)).length
                  + ['serial','serialnumber','sn'].filter(k => cells.includes(k)).length;
      if (score >= 1) return 'GENERIC_HEADERS';
    }

    // 5. Data rows with product type in col0 (simplified ARS without inventory tab)
    const dataRows = nonEmpty.filter(r => isProductType(r[0]));
    if (dataRows.length >= 2) return 'ARS_SIMPLE';
  }

  return 'GENERIC_HEADERLESS';
}

// ─── PARSERS ─────────────────────────────────────────────────────────────────

function parseARS(wb) {
  let devices = [];

  // 1. Parse primary "Customer Inventory" sheet
  const invSheet = wb.SheetNames.find(n => n.toLowerCase().includes('inventory'));
  if (invSheet) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[invSheet], { header: 1, defval: '' });
    devices = parseARSRows(rows);
    console.log(`[parseARS] Inventory sheet "${invSheet}": ${devices.length} devices`);
  }

  // 2. Also parse any "Quote Request" sheet for mixed summary+detail (planbitPrice training)
  const quoteSheet = wb.SheetNames.find(n => /quote\s*request/i.test(n));
  if (quoteSheet) {
    const qRows = XLSX.utils.sheet_to_json(wb.Sheets[quoteSheet], { header: 1, defval: '' });
    const summaryPrices = extractSummaryPrices(qRows);
    if (Object.keys(summaryPrices).length > 0) {
      console.log(`[parseARS] Quote Request summary prices:`, JSON.stringify(summaryPrices));
      // Attach planbitPrice to matching devices
      for (const d of devices) {
        const modelKey = d.model.toLowerCase().replace(/\s+/g, ' ').trim();
        for (const [pattern, price] of Object.entries(summaryPrices)) {
          if (modelKey.includes(pattern) || pattern.includes(modelKey.split(' ').slice(-2).join(' '))) {
            d.planbitPrice = price;
            break;
          }
        }
      }
    }
  }

  // 3. Parse RITM-prefixed sheets as separate device lists (GENERIC_HEADERS style)
  const ritmSheets = wb.SheetNames.filter(n => /^RITM/i.test(n.trim()));
  if (ritmSheets.length > 0) {
    console.log(`[parseARS] Found ${ritmSheets.length} RITM sheet(s): ${ritmSheets.join(', ')}`);
    // Build a temporary workbook-like object with only RITM sheets
    const ritmWb = { SheetNames: ritmSheets, Sheets: {} };
    for (const n of ritmSheets) ritmWb.Sheets[n] = wb.Sheets[n];
    const ritmDevices = parseGenericHeaders(ritmWb);
    console.log(`[parseARS] RITM sheets parsed: ${ritmDevices.length} additional devices`);
    devices = devices.concat(ritmDevices);
  }

  return devices;
}

function parseARSSimple(wb) {
  // Like ARS but no dedicated inventory tab
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const devices = parseARSRows(rows);
    if (devices.length) return devices;
  }
  return [];
}

function parseARSRows(rows) {
  const summaryDevices = [];
  const detailDevices  = [];

  for (const row of rows) {
    const col0 = cellStr(row[0]);
    const col1 = cellStr(row[1]);
    const col2 = cellStr(row[2]);
    const col3 = cellStr(row[3]);
    const col4 = row[4];
    const col5 = row[5];

    // Summary: Product | Brand | Model | Processor | Qty | Price
    if (isProductType(col0) && col1.length > 0 && col2.length > 0 &&
        typeof col4 === 'number' && col4 >= 0) {
      const procStr = col3 + ' ' + col2; // combine processor + model for spec extraction
      const qty = Math.max(parseInt(col4) || 1, 1);
      const joepPrice = typeof col5 === 'number' ? col5 : null;

      for (let i = 0; i < qty; i++) {
        summaryDevices.push({
          model:     (col1 + ' ' + col2).trim(),
          cpu:       extractCPU(col3) || extractCPU(col2),
          ram:       extractRAM(procStr) || extractRAM(col2),
          ssd:       extractSSD(col3) || extractSSD(col2),
          joepPrice,
        });
      }
    }

    // Detail block: Product | (empty) | Full model | Serial
    if (isProductType(col0) && col1 === '' && looksLikeModel(col2) && col3.length > 0) {
      detailDevices.push({ model: col2, serial: col3 });
    }
  }

  // Merge detail serials into summary devices
  if (summaryDevices.length > 0 && detailDevices.length > 0) {
    detailDevices.forEach((d, i) => {
      if (summaryDevices[i]) {
        summaryDevices[i].serial = d.serial;
        if (d.model.length > summaryDevices[i].model.length) {
          summaryDevices[i].model = d.model;
        }
      }
    });
  }

  return summaryDevices.length ? summaryDevices : detailDevices;
}

// Extract summary prices from a Quote Request sheet
// Looks for rows with Product | Brand | Model | Processor | Qty | Price pattern
function extractSummaryPrices(rows) {
  const prices = {};
  for (const row of rows) {
    const col0 = cellStr(row[0]);
    const col1 = cellStr(row[1]);
    const col2 = cellStr(row[2]);
    const col4 = row[4];
    const col5 = row[5];
    // Match summary rows: Product type | Brand | Model | Processor | Qty(number) | Price(number)
    if (isProductType(col0) && col1.length > 0 && col2.length > 0 &&
        typeof col4 === 'number' && typeof col5 === 'number' && col5 > 0) {
      const modelKey = (col1 + ' ' + col2).toLowerCase().replace(/\s+/g, ' ').trim();
      prices[modelKey] = col5;
      console.log(`[extractSummaryPrices] "${modelKey}" → €${col5}`);
    }
  }
  return prices;
}

function parseVendorQuote(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });

    // Find header row
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase());
      if (cells.some(c => c === 'product') && cells.some(c => c === 'brand') && cells.some(c => c === 'model')) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;

    const header = rows[headerRowIdx].map(v => cellStr(v).toLowerCase());
    const col = (...names) => {
      for (const n of names) {
        const idx = header.findIndex(h => h.includes(n));
        if (idx >= 0) return idx;
      }
      return -1;
    };

    const productCol   = col('product');
    const brandCol     = col('brand');
    const modelCol     = col('model');
    const processorCol = col('processor');
    const hddCol       = col('hdd', 'storage', 'disk');
    const memCol       = col('mem', 'ram', 'memory');
    const qtyCol       = col('qty', 'quantity', 'units');

    const devices = [];
    for (const row of rows.slice(headerRowIdx + 1)) {
      if (!row.some(v => cellStr(v).length > 0)) continue;
      const product = cellStr(row[productCol] ?? '');
      if (!isProductType(product)) continue;

      const brand = cellStr(row[brandCol] ?? '');
      const model = cellStr(row[modelCol] ?? '');
      if (!brand && !model) continue;

      const procRaw = cellStr(row[processorCol] ?? '');
      const hddRaw  = cellStr(row[hddCol] ?? '');
      const memRaw  = cellStr(row[memCol] ?? '');

      // Qty: strip commas ("6,000" → 6000), cap for display
      const qtyRaw = qtyCol >= 0 ? String(row[qtyCol] ?? '').replace(/,/g, '') : '1';
      const qty = Math.min(parseInt(qtyRaw) || 1, 500);
      const expand = qty <= 20 ? qty : 1;

      const cpu = extractCPU(procRaw) || extractCPU(model);
      const ram = extractRAM(memRaw) || extractRAM(procRaw);
      const ssd = extractSSD(hddRaw) || extractSSD(procRaw);

      for (let i = 0; i < expand; i++) {
        devices.push({ model: (brand + ' ' + model).trim(), cpu, ram, ssd });
      }
    }
    if (devices.length) return devices;
  }
  return [];
}

function parseGenericHeaders(wb) {
  const results = [];

  // Sort sheets by row count (most data first)
  const sortedSheets = [...wb.SheetNames].sort((a, b) => {
    const rowsA = XLSX.utils.sheet_to_json(wb.Sheets[a], { header: 1, defval: '' }).length;
    const rowsB = XLSX.utils.sheet_to_json(wb.Sheets[b], { header: 1, defval: '' }).length;
    return rowsB - rowsA;
  });
  console.log(`[parseGenericHeaders] Sheet order: ${sortedSheets.map(n => n + '(' + XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }).length + ' rows)').join(', ')}`);

  for (const name of sortedSheets) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    if (!rows.length) continue;

    // Find header row
    let headerRowIdx = -1;
    const COL_KEYS = {
      brand:  ['manufacturer','brand','merk','vendor','make'],
      model:  ['model','device','modelname','computername','description','productname','name','item','assetname'],
      cpu:    ['cpu','processor','proc','cputype'],
      ram:    ['ram','memory','mem','ramsizeoutgoing'],
      ssd:    ['ssd','storage','hdd','disk','drive','hddsizeoutgoing'],
      grade:  ['grade','condition','quality','cosmeticcategory'],
      serial: ['serial','serialnumber','sn','assettag','asset'],
      qty:    ['qty','quantity','count','units'],
    };
    const allKeys = Object.values(COL_KEYS).flat();

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const cells = rows[i].map(v => cellStr(v).toLowerCase().replace(/[\s_\-\.]/g,''));
      if (cells.filter(c => allKeys.includes(c)).length >= 1) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;

    const header = rows[headerRowIdx].map(v => cellStr(v).toLowerCase().replace(/[\s_\-\.]/g,''));
    const colMap = {};
    for (const [field, keys] of Object.entries(COL_KEYS)) {
      colMap[field] = header.findIndex(h => keys.includes(h));
    }
    console.log(`[parseGenericHeaders] Sheet "${name}" colMap:`, JSON.stringify(colMap));

    const mappedIdxs = new Set(Object.values(colMap).filter(i => i >= 0));

    let lastModelIdx = -1; // track last row with a valid model (for pivot merge)

    for (const row of rows.slice(headerRowIdx + 1)) {
      if (!row.some(v => cellStr(v).length > 0)) continue;
      const model = colMap.model >= 0 ? cellStr(row[colMap.model]) : '';
      const hasModel = model && looksLikeModel(model);

      // Extract specs from this row
      const cpu   = colMap.cpu   >= 0 ? cellStr(row[colMap.cpu])   : '';
      const ram   = colMap.ram   >= 0 ? cellStr(row[colMap.ram])   : '';
      const ssd   = colMap.ssd   >= 0 ? cellStr(row[colMap.ssd])   : '';
      const grade = colMap.grade >= 0 ? cellStr(row[colMap.grade]) : '';
      const hasSpecs = cpu || ram || ssd || grade;

      if (!hasModel && !hasSpecs) continue;

      // Pivot merge: row has specs but no model → merge into previous model row
      if (!hasModel && hasSpecs && lastModelIdx >= 0) {
        const prev = results[lastModelIdx];
        if (cpu   && !prev.cpu)   prev.cpu   = cpu;
        if (ram   && !prev.ram)   prev.ram   = ram;
        if (ssd   && !prev.ssd)   prev.ssd   = ssd;
        if (grade && !prev.grade) prev.grade = grade;
        // Also check for serial in unmapped columns
        if (!prev.serial) {
          row.forEach((v, i) => {
            if (!mappedIdxs.has(i) && !prev.serial && /^[A-Z0-9]{6,}$/i.test(cellStr(v))) {
              prev.serial = cellStr(v);
            }
          });
        }
        console.log(`[pivot-merge] Merged spec row into "${prev.model}" (cpu=${cpu||'-'}, ram=${ram||'-'}, ssd=${ssd||'-'}, grade=${grade||'-'})`);
        continue;
      }
      if (!hasModel && hasSpecs && lastModelIdx < 0) {
        console.log(`[pivot-skip] Spec-only row skipped (no prior model): cpu=${cpu||'-'}, ram=${ram||'-'}, ssd=${ssd||'-'}`);
      }

      if (!hasModel) continue; // no model and no previous row to merge into

      // Combine brand + model if brand column exists
      const brandVal = colMap.brand >= 0 ? cellStr(row[colMap.brand]) : '';
      let fullModel = model;
      if (brandVal && !model.toLowerCase().startsWith(brandVal.toLowerCase())) {
        fullModel = brandVal + ' ' + model;
      }

      const d = { model: fullModel };
      if (cpu)   d.cpu   = cpu;
      if (ram)   d.ram   = ram;
      if (ssd)   d.ssd   = ssd;
      if (grade) d.grade = grade;

      // Auto-detect serial from unmapped columns
      if (!d.serial) {
        row.forEach((v, i) => {
          if (!mappedIdxs.has(i) && !d.serial && /^[A-Z0-9]{6,}$/i.test(cellStr(v))) {
            d.serial = cellStr(v);
          }
        });
      }

      results.push(d);
      lastModelIdx = results.length - 1;
    }
  }
  console.log(`[parseGenericHeaders] Parsed ${results.length} devices from ${wb.SheetNames.length} sheet(s)`);
  return results;
}

function parseGenericHeaderless(wb) {
  const results = [];

  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    const nonEmpty = rows.filter(r => r.some(v => cellStr(v).length > 0));
    if (nonEmpty.length < 2) continue;

    const numCols = Math.max(...nonEmpty.map(r => r.length), 0);

    // Score each column
    const scores = {};
    for (let c = 0; c < numCols; c++) {
      const vals = nonEmpty.map(r => cellStr(r[c])).filter(v => v.length > 0);
      if (!vals.length) continue;
      const r = n => n / vals.length;
      scores[c] = {
        model:  r(vals.filter(looksLikeModel).length),
        serial: r(vals.filter(v => /^[A-Z0-9]{6,}$/i.test(v)).length),
        ram:    r(vals.filter(v => /^(4|8|16|32|64|128)\s*GB?$/i.test(v)).length),
        ssd:    r(vals.filter(v => /^(128|240|256|480|512|960|1024|2048)\s*GB?$/i.test(v)).length),
        grade:  r(vals.filter(v => /^[A-D]\d?$/i.test(v)).length),
      };
    }

    const pick = (field, exclude = []) => {
      let best = -1, bestScore = 0.12;
      for (let c = 0; c < numCols; c++) {
        if (exclude.includes(c) || !scores[c]) continue;
        if ((scores[c][field] || 0) > bestScore) { best = c; bestScore = scores[c][field]; }
      }
      return best;
    };

    const modelCol  = pick('model');
    if (modelCol === -1) continue;
    const serialCol = pick('serial', [modelCol]);
    const ramCol    = pick('ram',    [modelCol, serialCol]);
    const ssdCol    = pick('ssd',    [modelCol, serialCol, ramCol]);
    const gradeCol  = pick('grade',  [modelCol, serialCol, ramCol, ssdCol]);

    const sheetResults = nonEmpty
      .filter(row => looksLikeModel(row[modelCol]))
      .map(row => {
        const d = { model: cellStr(row[modelCol]) };
        if (serialCol >= 0) d.serial = cellStr(row[serialCol]);
        if (ramCol    >= 0) d.ram    = cellStr(row[ramCol]);
        if (ssdCol    >= 0) d.ssd    = cellStr(row[ssdCol]);
        if (gradeCol  >= 0) d.grade  = cellStr(row[gradeCol]);
        return d;
      });

    results.push(...sheetResults);
  }
  return results;
}

// ─── TEXT INPUT PARSER ────────────────────────────────────────────────────────
function parseTextInput(text) {
  // Line-by-line email header filter — safe, doesn't eat device lines
  const EMAIL_JUNK = /^(summarize this email|inbox|re:|fwd:|from:|to me|to:|subject:|sent:|cc:|planbit|caas\s*[-–]?\s*$|feb |mar |jan |apr |may |jun |jul |aug |sep |oct |nov |dec |http)/i;
  const HAS_DEVICE = /\b(dell|hp|hewlett|lenovo|apple|macbook|thinkpad|latitude|elitebook|surface|asus|acer|fujitsu|toshiba|samsung|microsoft|iphone|ipad|galaxy|yoga|optiplex|probook|zbook|spectre|envy|pavilion)\b/i;

  const lines = text
    .split(/[\n;]+/)
    .map(l => l.trim())
    .filter(l => l.length > 4)
    .filter(l => !EMAIL_JUNK.test(l) || HAS_DEVICE.test(l));

  const devices = [];

  for (const line of lines) {
    if (!HAS_DEVICE.test(line)) continue;

    // Qty: "60x", "- 60x", "• 3×", "CAAS - 60x"
    const qtyMatch = line.match(/(?:^|CAAS\s*[-–]\s*)[\-\*\•\·]?\s*(\d+)\s*[xX×]/i);
    const qty = qtyMatch ? Math.min(parseInt(qtyMatch[1]), 999) : 1;
    const rest = qtyMatch ? line.slice(line.indexOf(qtyMatch[0]) + qtyMatch[0].length).trim() : line;

    // Model: everything before first / or , or spec keyword
    const modelStop = /[\/,]|\b(i[3579][-\s]|\d+\s*GB|\d+\s*TB|Gen\s*\d|M[123]\b|FHD|UHD|4K|\d+"|\d+inch)/i;
    const modelRaw = rest.split(modelStop)[0].trim().replace(/^[-\*\•\·\s]+/, '');

    if (!HAS_DEVICE.test(modelRaw) || modelRaw.length < 4) continue;

    const cpu   = extractCPU(rest);
    const ram   = extractRAM(rest);
    const ssd   = extractSSD(rest);
    const grade = extractGrade(rest);

    devices.push({ model: modelRaw, cpu, ram, ssd, grade });
  }

  // Deduplicate: if same model appears twice (e.g. "- 60x Dell" + "60x Dell..."),
  // merge into the one with most specs
  const merged = [];
  for (const d of devices) {
    const existing = merged.find(m => m.model.toLowerCase() === d.model.toLowerCase());
    if (existing) {
      // Keep whichever has more specs filled
      if (!existing.ram && d.ram) existing.ram = d.ram;
      if (!existing.ssd && d.ssd) existing.ssd = d.ssd;
      if (!existing.cpu && d.cpu) existing.cpu = d.cpu;
      if (!existing.grade && d.grade) existing.grade = d.grade;
    } else {
      merged.push(d);
    }
  }
  return merged;
}

// ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────
function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const format = detectFormat(wb);

  switch (format) {
    case 'ARS':             return parseARS(wb);
    case 'ARS_SIMPLE':     return parseARSSimple(wb);
    case 'VENDOR_QUOTE':   return parseVendorQuote(wb);
    case 'GENERIC_HEADERS':return parseGenericHeaders(wb);
    default:               return parseGenericHeaderless(wb);
  }
}


// ─── ROUTE 1: Single quote ───────────────────────────────────────────────────
router.post('/api/quote', (req, res) => {
  try {
    const result = calculatePrice(req.body);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 1b: AI-powered quote via openclaw ────────────────────────────────
const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileAsync = promisify(execFile);

// Resolve openclaw binary: honour env var, else try PATH, else known locations
const OPENCLAW_BIN = process.env.OPENCLAW_PATH || 'openclaw';

router.post('/api/ai-quote', async (req, res) => {
  try {
    const { brand, model, cpu, ram, storage, condition, keyboard, region, battery, quantity } = req.body;
    const specs = [brand, model, cpu, ram, storage, condition, keyboard ? 'Keyboard: ' + keyboard : '', region ? 'Region: ' + region : '', battery ? 'Battery: ' + battery : '', quantity ? 'Quantity: ' + quantity : ''].filter(Boolean).join(', ');
    console.log('[ai-quote] incoming specs:', specs);
    if (!specs) return res.status(400).json({ ok: false, error: 'No device specs provided' });

    const args = [
      'agent', '--agent', 'main',
      '-m', `Price this device: ${specs}`,
      '--json'
    ];
    console.log('[ai-quote] executing:', OPENCLAW_BIN, args.join(' '));

    const { stdout, stderr } = await execFileAsync(OPENCLAW_BIN, args, {
      timeout: 60000,
      env: { ...process.env, PATH: `${process.env.PATH}:/usr/local/bin:/opt/homebrew/bin` }
    });

    if (stderr) console.warn('[ai-quote] stderr:', stderr);
    console.log('[ai-quote] raw stdout (%d chars):', stdout.length, stdout.slice(0, 500));

    let parsed;
    try {
      parsed = JSON.parse(stdout);
    } catch (parseErr) {
      console.error('[ai-quote] JSON parse failed:', parseErr.message);
      console.error('[ai-quote] stdout was:', stdout.slice(0, 1000));
      return res.status(502).json({ ok: false, error: 'AI agent returned invalid JSON' });
    }

    console.log('[ai-quote] parsed keys:', Object.keys(parsed));
    console.log('[ai-quote] result keys:', parsed.result ? Object.keys(parsed.result) : 'NO result key');
    if (parsed.result?.payloads) {
      console.log('[ai-quote] payloads count:', parsed.result.payloads.length);
      if (parsed.result.payloads.length > 0) {
        console.log('[ai-quote] payloads[0]:', JSON.stringify(parsed.result.payloads[0]).slice(0, 300));
      }
    }

    // Try multiple response shapes: result.payloads[0].text → result.text → result → raw
    const text = parsed.result?.payloads?.[0]?.text
              || parsed.result?.text
              || (typeof parsed.result === 'string' ? parsed.result : null)
              || parsed.text
              || parsed.output;

    if (!text) {
      // Retry once — OpenClaw sometimes returns empty on first call after restart
      console.warn('[ai-quote] empty response, retrying once...');
      try {
        const { stdout: stdout2 } = await execFileAsync(OPENCLAW_BIN, args, {
          timeout: 60000,
          env: { ...process.env, PATH: `${process.env.PATH}:/usr/local/bin:/opt/homebrew/bin` }
        });
        const parsed2 = JSON.parse(stdout2);
        const text2 = parsed2.result?.payloads?.[0]?.text
                    || parsed2.result?.text
                    || (typeof parsed2.result === 'string' ? parsed2.result : null)
                    || parsed2.text
                    || parsed2.output;
        if (text2) {
          console.log('[ai-quote] retry success, returning %d chars', text2.length);
          return res.json({ ok: true, result: text2 });
        }
      } catch (retryErr) {
        console.error('[ai-quote] retry also failed:', retryErr.message);
      }
      console.error('[ai-quote] could not extract text. Full response:', JSON.stringify(parsed).slice(0, 1000));
      return res.status(502).json({ ok: false, error: 'No pricing result from AI agent — probeer het opnieuw', debug: Object.keys(parsed) });
    }

    console.log('[ai-quote] success, returning %d chars', text.length);
    res.json({ ok: true, result: text });
  } catch (err) {
    console.error('[ai-quote] error:', err.code || err.message);
    if (err.code === 'ENOENT') {
      return res.status(503).json({
        ok: false,
        error: `openclaw binary not found at "${OPENCLAW_BIN}". Set OPENCLAW_PATH env var to the correct path.`
      });
    }
    const status = err.killed ? 504 : 500;
    res.status(status).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 2: Batch file analysis ───────────────────────────────────────────
router.post('/api/analyze', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: 'No file uploaded' });
    const region  = req.body.region || 'EU';
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const format = detectFormat(wb);
    const devices = parseExcel(req.file.buffer);
    if (!devices.length) return res.status(400).json({ ok: false, error: 'No valid devices found in file' });
    const { results, summary } = analyzeDevices(devices, region);
    res.json({ ok: true, format, results, summary });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});


// ─── ROUTE 3: Text input analysis ────────────────────────────────────────────
router.post('/api/analyze-text', (req, res) => {
  try {
    const { text = '', region = 'EU' } = req.body;
    if (!text.trim()) return res.status(400).json({ ok: false, error: 'No text provided' });

    const devices = parseTextInput(text);
    if (!devices.length) return res.status(400).json({ ok: false, error: 'Geen apparaten herkend in de tekst. Probeer: "60x Dell Latitude 7430 i7 32GB 512GB"' });

    const { results, summary } = analyzeDevices(devices, region);
    res.json({ ok: true, results, summary });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 4: Generate HTML report ──────────────────────────────────────────
router.post('/api/report', (req, res) => {
  try {
    const { devices = [], summary = {}, dealName = 'ERPIE Deal' } = req.body;
    const html = generateReport(devices, summary, dealName);
    res.type('text/html').send(html);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── HTML REPORT GENERATOR ───────────────────────────────────────────────────
function statusColor(status) {
  if (status === 'GO')    return '#00c853';
  if (status === 'WATCH') return '#ffab00';
  return '#d50000';
}

function generateReport(devices, summary, dealName) {
  const date = new Date().toLocaleDateString('nl-NL', { day: '2-digit', month: 'long', year: 'numeric' });
  const rows = devices.map(d => `
    <tr>
      <td>${escHtml(d.model || '')}</td>
      <td>${escHtml(d.gen || '')}</td>
      <td>${d.ramGb || ''}GB</td>
      <td>${d.ssdGb || ''}GB</td>
      <td>${escHtml(d.grade || '')}</td>
      <td><span class="pill" style="background:${statusColor(d.status)}20;color:${statusColor(d.status)};border:1px solid ${statusColor(d.status)}">${d.status}</span></td>
      <td>€${(d.advisedPrice || 0).toLocaleString('nl-NL')}</td>
      <td>€${(d.priceLow || 0).toLocaleString('nl-NL')} – €${(d.priceHigh || 0).toLocaleString('nl-NL')}</td>
    </tr>`).join('');

  return `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<title>ERPIE Report – ${escHtml(dealName)}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #f0f4f8; color: #1a202c; }
  .header { background: linear-gradient(135deg, #0a0a1f 0%, #1a1a3e 100%); color: #fff; padding: 32px 48px; }
  .header h1 { font-size: 28px; font-weight: 700; color: #00d4ff; }
  .header p  { font-size: 13px; color: #a0aec0; margin-top: 4px; }
  .content { max-width: 1100px; margin: 32px auto; padding: 0 24px; }
  .cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
  .card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
  .card .label { font-size: 12px; color: #718096; text-transform: uppercase; letter-spacing: .05em; }
  .card .value { font-size: 28px; font-weight: 700; color: #1a202c; margin-top: 6px; }
  .card .sub   { font-size: 12px; color: #a0aec0; margin-top: 2px; }
  .pills { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
  .pill { padding: 4px 12px; border-radius: 999px; font-size: 13px; font-weight: 600; }
  .rec { background: #fff; border-left: 4px solid #00d4ff; padding: 16px 20px; border-radius: 8px; margin-bottom: 24px; font-size: 14px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
  thead th { background: #2d3748; color: #e2e8f0; padding: 12px 16px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .05em; }
  tbody tr:nth-child(even) { background: #f7fafc; }
  tbody td { padding: 10px 16px; font-size: 13px; border-bottom: 1px solid #e2e8f0; }
  tfoot td { padding: 12px 16px; font-weight: 700; background: #edf2f7; }
  .footer { text-align: center; padding: 32px; font-size: 11px; color: #a0aec0; }
</style>
</head>
<body>
<div class="header">
  <h1>PlanBit – ERPIE Price Report</h1>
  <p>${escHtml(dealName)} &nbsp;|&nbsp; ${date} &nbsp;|&nbsp; Powered by ERPIE PriceFinder v1.0</p>
</div>
<div class="content">
  <div class="cards">
    <div class="card">
      <div class="label">Total Assets</div>
      <div class="value">${summary.total || 0}</div>
      <div class="sub">devices analysed</div>
    </div>
    <div class="card">
      <div class="label">Advised Value</div>
      <div class="value">€${(summary.totalValue || 0).toLocaleString('nl-NL')}</div>
      <div class="sub">sum of ERP prices</div>
    </div>
    <div class="card">
      <div class="label">Average ERP</div>
      <div class="value">€${(summary.avgValue || 0).toLocaleString('nl-NL')}</div>
      <div class="sub">per device</div>
    </div>
    <div class="card">
      <div class="label">Bid Range</div>
      <div class="value" style="font-size:20px">€${(summary.bidLow || 0).toLocaleString('nl-NL')} – €${(summary.bidHigh || 0).toLocaleString('nl-NL')}</div>
      <div class="sub">suggested offer</div>
    </div>
  </div>
  <div class="pills">
    <span class="pill" style="background:#00c85320;color:#00c853;border:1px solid #00c853">GO: ${summary.goCount || 0}</span>
    <span class="pill" style="background:#ffab0020;color:#ffab00;border:1px solid #ffab00">WATCH: ${summary.watchCount || 0}</span>
    <span class="pill" style="background:#d5000020;color:#d50000;border:1px solid #d50000">NO-GO: ${summary.nogoCount || 0}</span>
  </div>
  <div class="rec">💡 <strong>Recommendation:</strong> ${escHtml(summary.recommendation || '')}</div>
  <table>
    <thead>
      <tr><th>Model</th><th>Gen</th><th>RAM</th><th>SSD</th><th>Grade</th><th>Status</th><th>ERP</th><th>Price Band</th></tr>
    </thead>
    <tbody>${rows}</tbody>
    <tfoot>
      <tr>
        <td colspan="6"><strong>TOTAAL (${summary.total || 0} devices)</strong></td>
        <td><strong>€${(summary.totalValue || 0).toLocaleString('nl-NL')}</strong></td>
        <td><strong>€${(summary.bidLow || 0).toLocaleString('nl-NL')} – €${(summary.bidHigh || 0).toLocaleString('nl-NL')}</strong></td>
      </tr>
    </tfoot>
  </table>
</div>
<div class="footer">
  ERPIE PriceFinder · PlanBit ITAD · Prijzen zijn indicatief op basis van marktdata.<br>
  Werkelijke opbrengst kan afwijken o.b.v. conditie, vraag en logistieke kosten.
</div>
</body>
</html>`;
}

function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── START ───────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8000;
app.listen(PORT, () => console.log(`ERPIE PriceFinder running on http://localhost:${PORT}`));
