'use strict';

// ─── BASE PRICES PER CPU GENERATION ─────────────────────────────────────────
const BASE_PRICES = {
  'Gen14': 450, 'Gen13': 380, 'Gen12': 310, 'Gen11': 245, 'Gen10': 200,
  'Gen9': 160,  'Gen8': 150,  'Gen7': 70,   'Gen6': 40,   'Gen5': 20,
  'Gen4': 10
};

// ─── WATCH CAPS ──────────────────────────────────────────────────────────────
const WATCH_CAPS = { 'Gen8': 120, 'Gen9': 140 };

// ─── RAM ADJUSTMENTS (baseline 8GB) ─────────────────────────────────────────
const RAM_ADJ = { 4: -40, 8: 0, 16: 30, 32: 70, 64: 130 };

// ─── SSD ADJUSTMENTS (baseline 256GB) ───────────────────────────────────────
const SSD_ADJ = { 128: -25, 256: 0, 512: 35, 1024: 70, 2048: 110 };

// ─── GRADE MULTIPLIERS ───────────────────────────────────────────────────────
const GRADE_MULT = {
  'A1': 1.40, 'A2': 1.25, 'A3': 1.15,
  'A': 1.00,
  'B': 0.90, 'B4': 0.85,
  'C': 0.70, 'C6': 0.60,
  'D': 0.40
};

// ─── BRAND TIER MULTIPLIERS ──────────────────────────────────────────────────
const BRAND_MULT_MAP = [
  { keywords: ['xps', 'macbook pro'],           mult: 1.15 },
  { keywords: ['zbook', 'x1 carbon'],           mult: 1.12 },
  { keywords: ['thinkpad p'],                   mult: 1.10 },
  { keywords: ['elitebook', 'latitude', 'thinkpad t', 'thinkpad x3', 'thinkpad x2', 'thinkpad x1', 'x360', 'x280', 'x270', 'x260', 'x390'], mult: 1.00 },
  { keywords: ['probook', 'thinkpad e'],        mult: 0.95 },
  { keywords: ['vostro'],                       mult: 0.92 },
  { keywords: ['ideapad', 'hp 250'],            mult: 0.90 },
];

// ─── MODEL DATABASE ──────────────────────────────────────────────────────────
const MODEL_DB = {
  // Dell Latitude
  'latitude 5410': 'Gen10', 'latitude 5420': 'Gen11', 'latitude 5430': 'Gen12', 'latitude 5440': 'Gen13',
  'latitude 5480': 'Gen7',  'latitude 5490': 'Gen8',  'latitude 5500': 'Gen8',  'latitude 5510': 'Gen10',
  'latitude 5520': 'Gen11', 'latitude 5530': 'Gen12',
  'latitude 7280': 'Gen7',  'latitude 7290': 'Gen8',  'latitude 7300': 'Gen8',  'latitude 7310': 'Gen10',
  'latitude 7320': 'Gen11', 'latitude 7330': 'Gen12', 'latitude 7400': 'Gen8',  'latitude 7410': 'Gen10',
  'latitude 7420': 'Gen11', 'latitude 7430': 'Gen12', 'latitude 7480': 'Gen7',  'latitude 7490': 'Gen8',

  // Dell XPS
  'xps 13 9310': 'Gen11', 'xps 13 9300': 'Gen10', 'xps 13 9380': 'Gen8',
  'xps 15 9500': 'Gen10', 'xps 15 9570': 'Gen8',

  // Dell OptiPlex (desktops)
  'optiplex 3050': 'Gen7', 'optiplex 3060': 'Gen8', 'optiplex 3070': 'Gen9', 'optiplex 3080': 'Gen10',
  'optiplex 5050': 'Gen7', 'optiplex 5060': 'Gen8', 'optiplex 5070': 'Gen9', 'optiplex 5080': 'Gen10',
  'optiplex 7050': 'Gen7', 'optiplex 7060': 'Gen8', 'optiplex 7070': 'Gen9', 'optiplex 7080': 'Gen10',

  // HP EliteBook
  'elitebook 830 g5': 'Gen8',  'elitebook 830 g6': 'Gen8',  'elitebook 830 g7': 'Gen10', 'elitebook 830 g8': 'Gen11',
  'elitebook 840 g5': 'Gen8',  'elitebook 840 g6': 'Gen8',  'elitebook 840 g7': 'Gen10', 'elitebook 840 g8': 'Gen11',
  'elitebook 840 g9': 'Gen12', 'elitebook 840 g10': 'Gen13',
  'elitebook 850 g5': 'Gen8',  'elitebook 850 g6': 'Gen8',  'elitebook 850 g7': 'Gen10', 'elitebook 850 g8': 'Gen11',
  // HP EliteBook AMD Ryzen
  'elitebook 835 g7': 'Gen10', 'elitebook 835 g8': 'Gen11',
  'elitebook 855 g7': 'Gen10', 'elitebook 855 g8': 'Gen11',
  'elitebook x360 1030 g2': 'Gen7', 'elitebook x360 1030 g3': 'Gen8', 'elitebook x360 1030 g4': 'Gen8',
  'elitebook x360 1030 g7': 'Gen10', 'elitebook x360 1030 g8': 'Gen11', 'elitebook x360 1030 g9': 'Gen12',
  'elitebook 1040 g9': 'Gen12', 'elitebook 1040 g10': 'Gen13',
  'x360 1030 g2': 'Gen7', 'x360 1030 g3': 'Gen8', 'x360 1030 g4': 'Gen8',
  'x360 1030 g7': 'Gen10', 'x360 1030 g8': 'Gen11', 'x360 1030 g9': 'Gen12',
  '1040 g9': 'Gen12', '1040 g10': 'Gen13',

  // HP ZBook
  'zbook firefly g7': 'Gen10', 'zbook firefly g8': 'Gen11', 'zbook firefly 14 g7': 'Gen10',
  'zbook firefly 14 g8': 'Gen11', 'zbook firefly 15 g7': 'Gen10', 'zbook firefly 15 g8': 'Gen11',
  'zbook fury g7': 'Gen10', 'zbook fury g8': 'Gen11', 'zbook fury 15 g7': 'Gen10',
  'zbook fury 15 g8': 'Gen11', 'zbook fury 17 g7': 'Gen10', 'zbook fury 17 g8': 'Gen11',
  'zbook power g7': 'Gen10', 'zbook power g8': 'Gen11',
  'zbook studio g7': 'Gen10', 'zbook studio g8': 'Gen11',

  // HP ProBook
  'probook 430 g5': 'Gen8',  'probook 430 g6': 'Gen8',  'probook 430 g7': 'Gen10',
  'probook 440 g5': 'Gen8',  'probook 440 g6': 'Gen8',  'probook 440 g7': 'Gen10',
  'probook 450 g5': 'Gen8',  'probook 450 g6': 'Gen8',  'probook 450 g7': 'Gen10',
  'probook 640 g4': 'Gen8',  'probook 640 g5': 'Gen8',
  'probook 650 g4': 'Gen8',  'probook 650 g5': 'Gen8',

  // HP 250
  'hp 250 g6': 'Gen6', '250 g6': 'Gen6',
  'hp 250 g7': 'Gen8', '250 g7': 'Gen8',

  // Lenovo ThinkPad T
  't14 gen 1': 'Gen10', 't14 gen 2': 'Gen11', 't14 gen 3': 'Gen12',
  't14s gen 1': 'Gen10', 't14s gen 2': 'Gen11', 't14s gen 3': 'Gen12',
  't460': 'Gen6', 't460s': 'Gen6',
  't470': 'Gen7', 't470s': 'Gen7',
  't480': 'Gen8', 't480s': 'Gen8',
  't490': 'Gen8', 't490s': 'Gen8',

  // Lenovo ThinkPad X
  'x1 carbon gen 5': 'Gen7', 'x1 carbon gen 6': 'Gen8', 'x1 carbon gen 7': 'Gen8',
  'x1 carbon gen 8': 'Gen10', 'x1 carbon gen 9': 'Gen11', 'x1 carbon gen 10': 'Gen12',
  'x260': 'Gen6', 'x270': 'Gen7', 'x280': 'Gen8', 'x390': 'Gen8',

  // Lenovo ThinkCentre (desktops)
  'm920q': 'Gen9', 'm720q': 'Gen8',

  // Toshiba
  'portege z30-a': 'Gen4', 'portege z30-b': 'Gen5', 'portege z30-c': 'Gen6',

  // Apple Intel
  'macbook pro 16 2019': 'Gen9',
  'macbook pro 15 2019': 'Gen9',
  'macbook pro 13 2020': 'Gen10',
};

// ─── APPLE SILICON DIRECT PRICES ─────────────────────────────────────────────
const APPLE_SILICON = {
  'macbook air m1': 450,
  'macbook air m2': 800,
  'macbook air m3': 1000,
  'macbook pro 14 m1 pro': 650,
  'macbook pro 16 m1 pro': 500,
};

// ─── APPLE INTEL DIRECT PRICES ───────────────────────────────────────────────
const APPLE_INTEL = {
  'macbook pro 16 2019': 390,
  'macbook pro 15 2019': 350,
  'macbook pro 13 2020': 350,
};

// ─── MODEL-SPECIFIC BASE PRICE OVERRIDES ─────────────────────────────────────
// These override the generation-based price when the model is matched exactly.
// Used for models whose market value differs significantly from gen-average.
const MODEL_BASE_PRICE = {
  // HP EliteBook AMD Ryzen (lower than Intel equivalents at same gen)
  'elitebook 835 g7': 180, 'elitebook 835 g8': 200,
  'elitebook 855 g7': 185, 'elitebook 855 g8': 210,
  // HP ZBook workstations (premium over standard laptops)
  'zbook firefly g7': 350, 'zbook firefly g8': 350,
  'zbook firefly 14 g7': 350, 'zbook firefly 14 g8': 350,
  'zbook firefly 15 g7': 350, 'zbook firefly 15 g8': 350,
  'zbook fury g7': 420, 'zbook fury g8': 420,
  'zbook fury 15 g7': 420, 'zbook fury 15 g8': 420,
  'zbook fury 17 g7': 420, 'zbook fury 17 g8': 420,
  'zbook power g7': 380, 'zbook power g8': 380,
  'zbook studio g7': 400, 'zbook studio g8': 400,
};

// ─── REGION ADJUSTMENTS ───────────────────────────────────────────────────────
const REGION_ADJ = { 'EU': 1.00, 'UK': 0.84, 'INTL': 0.85 };

// ─── NORMALISE MODEL NAME ─────────────────────────────────────────────────────
function normaliseModel(raw) {
  let s = (raw || '').toLowerCase().trim();
  s = s.replace(/hewlett[-\s]?packard|hp inc\./gi, 'hp');
  s = s.replace(/dell inc\./gi, 'dell');
  s = s.replace(/^lenovo\s+/i, '');
  s = s.replace(/\s+notebook\s+pc$/i, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

// ─── PARSE RAM ────────────────────────────────────────────────────────────────
function parseRam(val) {
  if (typeof val === 'number') return val;
  const m = String(val).match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 8;
}

// ─── PARSE SSD ────────────────────────────────────────────────────────────────
function parseSsd(val) {
  if (typeof val === 'number') return val;
  const s = String(val).toLowerCase();
  const m = s.match(/(\d+(?:\.\d+)?)\s*(tb|gb)?/);
  if (!m) return 256;
  let v = parseFloat(m[1]);
  if (m[2] === 'tb') v = Math.round(v * 1024);
  return v;
}

// ─── SNAP TO NEAREST SSD KEY ─────────────────────────────────────────────────
function snapSsd(gb) {
  const keys = [128, 256, 512, 1024, 2048];
  return keys.reduce((prev, k) => Math.abs(k - gb) < Math.abs(prev - gb) ? k : prev, 256);
}

// ─── SNAP TO NEAREST RAM KEY ─────────────────────────────────────────────────
function snapRam(gb) {
  const keys = [4, 8, 16, 32, 64];
  return keys.reduce((prev, k) => Math.abs(k - gb) < Math.abs(prev - gb) ? k : prev, 8);
}

// ─── TYPICAL CPUs PER GENERATION (for UI hints) ──────────────────────────────
const GEN_CPU_HINTS = {
  'Gen6':  ['i5-6200U', 'i5-6300U', 'i7-6500U', 'i7-6600U'],
  'Gen7':  ['i5-7200U', 'i5-7300U', 'i7-7500U', 'i7-7600U'],
  'Gen8':  ['i5-8250U', 'i5-8350U', 'i7-8550U', 'i7-8650U'],
  'Gen9':  ['i5-9300H', 'i7-9750H', 'i9-9880H'],
  'Gen10': ['i5-10210U', 'i5-10310U', 'i7-10510U', 'i7-10610U'],
  'Gen11': ['i5-1135G7', 'i5-1145G7', 'i7-1165G7', 'i7-1185G7'],
  'Gen12': ['i5-1235U', 'i5-1245U', 'i7-1255U', 'i7-1265U'],
  'Gen13': ['i5-1335U', 'i5-1345U', 'i7-1355U', 'i7-1365U'],
  'Gen14': ['i5-1435U', 'i5-1445U', 'i7-1455U', 'i7-1465U', 'Ultra 5 125U', 'Ultra 7 155U'],
};

// ─── CPU PARSING FALLBACK ─────────────────────────────────────────────────────
function genFromCpu(cpuStr) {
  if (!cpuStr) return null;
  const s = cpuStr.toLowerCase();

  // Intel Core iX-NNNN[suffix] — extract leading digit block
  // e.g. i5-1135G7 → "1135", i7-8650U → "8650", i9-14900HX → "14900"
  const m = s.match(/i[3579]-(\d+)/);
  if (m) {
    const digits = m[1];
    if (digits.length >= 5) {
      // 5+ digit: first 2 digits = gen (10, 11, 12, 13, 14)
      return `Gen${parseInt(digits.substring(0, 2), 10)}`;
    }
    // 4-digit: first 2 digits in 10-14 range → modern gen (e.g. 1135→Gen11, 1245→Gen12)
    // otherwise → first digit only (e.g. 8265→Gen8, 7600→Gen7)
    const firstTwo = parseInt(digits.substring(0, 2), 10);
    return (firstTwo >= 10 && firstTwo <= 14) ? `Gen${firstTwo}` : `Gen${parseInt(digits[0], 10)}`;
  }

  // AMD Ryzen XXXX → first digit maps: 3→Gen9, 4→Gen10, 5→Gen11, 6→Gen12, 7→Gen13, 8→Gen14
  const r = s.match(/ryzen\s+\d+\s+(\d{4})/);
  if (r) {
    const genMap = { 3: 9, 4: 10, 5: 11, 6: 12, 7: 13, 8: 14 };
    return `Gen${genMap[parseInt(r[1][0], 10)] || 10}`;
  }

  return null;
}

// ─── LOOK UP GENERATION FROM MODEL ───────────────────────────────────────────
function lookupGen(normModel) {
  // Exact match
  if (MODEL_DB[normModel]) return MODEL_DB[normModel];

  // Prefix match (longest wins)
  let best = null;
  let bestLen = 0;
  for (const key of Object.keys(MODEL_DB)) {
    if (normModel.includes(key) && key.length > bestLen) {
      best = MODEL_DB[key];
      bestLen = key.length;
    }
  }
  return best;
}

// ─── GET BRAND MULTIPLIER ─────────────────────────────────────────────────────
function getBrandMult(normModel) {
  for (const entry of BRAND_MULT_MAP) {
    for (const kw of entry.keywords) {
      if (normModel.includes(kw)) return entry.mult;
    }
  }
  return 1.00;
}

// ─── CLASSIFY STATUS ──────────────────────────────────────────────────────────
function classifyStatus(gen) {
  const n = parseInt(gen.replace('Gen', ''), 10);
  if (n >= 10) return 'GO';
  if (n >= 8)  return 'WATCH';
  return 'NO-GO';
}

// ─── BATTERY ADJUSTMENT ───────────────────────────────────────────────────────
function batteryAdj(battery) {
  const b = (battery || '').toLowerCase();
  if (b === 'missing') return -50;
  if (b === 'bad' || b === 'poor') return -35;
  return 0;
}

// ─── MAIN PRICE FUNCTION ──────────────────────────────────────────────────────
function calculatePrice(input) {
  const {
    model:   rawModel = '',
    ram:     rawRam   = 8,
    ssd:     rawSsd   = 256,
    grade:   rawGrade = 'A',
    battery: rawBat   = 'good',
    region:  rawRegion= 'EU',
    cpu:     rawCpu   = '',
    qty:     rawQty   = 1,
  } = input;

  const normModel = normaliseModel(rawModel);
  const ramGb  = snapRam(parseRam(rawRam));
  const ssdGb  = snapSsd(parseSsd(rawSsd));
  const grade  = (rawGrade || 'A').toUpperCase().trim();
  const region = (rawRegion || 'EU').toUpperCase().trim();
  const qty    = Math.max(1, parseInt(rawQty, 10) || 1);

  const reasoning = [];

  // ── Apple Silicon ────────────────────────────────────────────────────────────
  for (const [key, price] of Object.entries(APPLE_SILICON)) {
    if (normModel.includes(key)) {
      const gradeMult   = GRADE_MULT[grade] || GRADE_MULT['A'];
      const regionMult  = REGION_ADJ[region] || 1.00;
      const advised     = Math.round(price * gradeMult * regionMult);
      const low         = Math.round(advised * 0.85);
      const high        = Math.round(advised * 1.15);
      reasoning.push(`Apple Silicon direct price: €${price}`);
      reasoning.push(`Grade ${grade} ×${gradeMult}, Region ${region} ×${regionMult.toFixed(2)}`);
      return {
        model: rawModel, gen: 'Silicon', status: 'GO',
        advisedPrice: advised, priceLow: low, priceHigh: high,
        ramGb, ssdGb, grade, region,
        cpu: rawCpu || null, knownCpus: [],
        reasoning,
      };
    }
  }

  // ── Apple Intel ──────────────────────────────────────────────────────────────
  for (const [key, price] of Object.entries(APPLE_INTEL)) {
    if (normModel.includes(key)) {
      const gradeMult  = GRADE_MULT[grade] || GRADE_MULT['A'];
      const regionMult = REGION_ADJ[region] || 1.00;
      const advised    = Math.round(price * gradeMult * regionMult);
      const low        = Math.round(advised * 0.85);
      const high       = Math.round(advised * 1.15);
      reasoning.push(`Apple Intel direct price: €${price}`);
      reasoning.push(`Grade ${grade} ×${gradeMult}, Region ${region} ×${regionMult.toFixed(2)}`);
      return {
        model: rawModel, gen: 'Intel', status: 'GO',
        advisedPrice: advised, priceLow: low, priceHigh: high,
        ramGb, ssdGb, grade, region,
        cpu: rawCpu || null, knownCpus: [],
        reasoning,
      };
    }
  }

  // ── Generation lookup ────────────────────────────────────────────────────────
  let gen = lookupGen(normModel);
  if (!gen && rawCpu) {
    gen = genFromCpu(rawCpu);
    if (gen) reasoning.push(`Gen from CPU string: ${rawCpu} → ${gen}`);
  }
  // Accept pre-inferred gen from parser (e.g., ZONES_INVENTORY)
  if (!gen && input.gen && input.gen !== 'Unknown') {
    gen = input.gen;
    reasoning.push(`Gen inferred from model number: ${gen}`);
  }
  if (!gen) {
    reasoning.push(`Unknown model: "${rawModel}" — defaulting to Gen8`);
    gen = 'Gen8';
  } else if (!reasoning.some(r => r.includes('Gen'))) {
    reasoning.push(`Model matched: "${normModel}" → ${gen}`);
  }

  // Check for model-specific base price override
  let basePrice;
  let modelOverrideKey = null;
  for (const key of Object.keys(MODEL_BASE_PRICE)) {
    if (normModel.includes(key)) {
      if (!modelOverrideKey || key.length > modelOverrideKey.length) {
        modelOverrideKey = key;
      }
    }
  }
  if (modelOverrideKey) {
    basePrice = MODEL_BASE_PRICE[modelOverrideKey];
    reasoning.push(`Model-specific base price for "${modelOverrideKey}": €${basePrice} (overrides ${gen} default €${BASE_PRICES[gen] ?? 0})`);
  } else {
    basePrice = BASE_PRICES[gen] ?? 0;
    reasoning.push(`Base price ${gen}: €${basePrice}`);
  }

  const rAdj = RAM_ADJ[ramGb] ?? 0;
  const sAdj = SSD_ADJ[ssdGb] ?? 0;
  const bAdj = batteryAdj(rawBat);
  reasoning.push(`RAM ${ramGb}GB adj: €${rAdj}, SSD ${ssdGb}GB adj: €${sAdj}, Battery adj: €${bAdj}`);

  const priceGradeA = Math.max(0, basePrice + rAdj + sAdj + bAdj);

  const gradeMult = GRADE_MULT[grade] ?? GRADE_MULT['A'];
  const brandMult = getBrandMult(normModel);
  reasoning.push(`Grade ${grade} ×${gradeMult}, Brand mult ×${brandMult}`);

  const regionMult = REGION_ADJ[region] || 1.00;

  let advised = Math.round(priceGradeA * gradeMult * brandMult * regionMult);

  const status = classifyStatus(gen);

  // Apply WATCH caps
  if (status === 'WATCH' && WATCH_CAPS[gen] !== undefined) {
    const cap = Math.round(WATCH_CAPS[gen] * regionMult);
    if (advised > cap) {
      reasoning.push(`WATCH cap applied: €${advised} → €${cap}`);
      advised = cap;
    }
  }

  const low  = Math.round(advised * 0.85);
  const high = Math.round(advised * 1.15);

  return {
    model: rawModel, gen, status,
    advisedPrice: advised, priceLow: low, priceHigh: high,
    ramGb, ssdGb, grade, region,
    cpu: rawCpu || null,
    knownCpus: GEN_CPU_HINTS[gen] || [],
    reasoning,
  };
}

// ─── BATCH ANALYSIS ────────────────────────────────────────────────────────────
function analyzeDevices(devices, region = 'EU') {
  const results = devices.map(d => {
    try {
      return calculatePrice({ ...d, region: d.region || region });
    } catch (e) {
      return { model: d.model || 'Unknown', status: 'ERROR', advisedPrice: 0, error: e.message };
    }
  });

  const total  = results.length;
  const goList = results.filter(r => r.status === 'GO');
  const watchList = results.filter(r => r.status === 'WATCH');
  const nogoList  = results.filter(r => r.status === 'NO-GO');
  const totalValue = results.reduce((s, r) => s + (r.advisedPrice || 0), 0);
  const avgValue   = total ? Math.round(totalValue / total) : 0;

  // Bid range based on 80% of total advised (typical batch discount)
  const bidLow  = Math.round(totalValue * 0.70);
  const bidHigh = Math.round(totalValue * 0.85);

  return {
    results,
    summary: {
      total, goCount: goList.length, watchCount: watchList.length, nogoCount: nogoList.length,
      totalValue, avgValue, bidLow, bidHigh,
      recommendation: totalValue > 0
        ? (goList.length / total >= 0.6
          ? 'Strong portfolio — majority GO assets. Recommend competitive bid.'
          : nogoList.length / total >= 0.5
          ? 'Weak portfolio — majority NO-GO assets. Bid conservatively or pass.'
          : 'Mixed portfolio — significant WATCH assets. Evaluate logistics cost carefully.')
        : 'No pricing data available.',
    },
  };
}

module.exports = { calculatePrice, analyzeDevices, normaliseModel };
