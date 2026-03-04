'use strict';

const express  = require('express');
const multer   = require('multer');
const XLSX     = require('xlsx');
const path     = require('path');
const { calculatePrice, analyzeDevices } = require('./pricing-engine');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.static(path.join(__dirname, '../client/public')));

// ─── COLUMN NAME ALIASES ────────────────────────────────────────────────────
const COL_ALIASES = {
  model:   ['model', 'device', 'devicename', 'product', 'description', 'item',
            'name', 'assetname', 'computername', 'modelname', 'type'],
  cpu:     ['cpu', 'processor', 'proc'],
  ram:     ['ram', 'memory', 'mem'],
  ssd:     ['ssd', 'storage', 'hdd', 'disk', 'drive'],
  grade:   ['grade', 'condition', 'cond', 'quality'],
  battery: ['battery', 'bat'],
  serial:  ['serial', 'serialnumber', 'sn', 'asset', 'assettag', 'serialno'],
  qty:     ['qty', 'quantity', 'count', 'units'],
};

// Keywords used to detect model columns in headerless files
const MODEL_KEYWORDS = [
  'hp', 'dell', 'lenovo', 'apple', 'thinkpad', 'elitebook', 'latitude',
  'macbook', 'probook', 'optiplex', 'toshiba', 'portege', 'xps', 'ideapad',
  'thinkcentre', 'zbook', 'vostro', 'fujitsu', 'asus', 'acer', 'surface',
  'notebook', 'laptop', 'desktop', 'workstation', 'precision', 'inspiron',
  'pavilion', 'folio', 'revolve', 'spectre', 'envy', 'elite', 'book',
];

const SERIAL_RE = /^[A-Z0-9]{6,}$/i;
const RAM_RE    = /^(\d+)\s*gb?$/i;
const SSD_RE    = /^(\d+)\s*(gb?|tb?)$/i;
const GRADE_RE  = /^[A-D]\d?$/i;
const RAM_VALS  = new Set(['4','8','16','32','64','128']);
const SSD_VALS  = new Set(['128','256','512','1024','2048','240','480','960']);

function cellStr(v) { return String(v ?? '').trim(); }

function looksLikeModel(v) {
  const s = cellStr(v).toLowerCase();
  if (s.length < 4) return false;
  return MODEL_KEYWORDS.some(k => s.includes(k));
}

function resolveCol(headers, field) {
  const aliases = COL_ALIASES[field] || [field];
  for (const alias of aliases) {
    const found = headers.find(h =>
      h.toLowerCase().replace(/[\s_\-\.]/g, '') === alias.replace(/[\s_\-\.]/g, '')
    );
    if (found) return found;
  }
  return null;
}

function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const results = [];
  const aliasFlat = Object.values(COL_ALIASES).flat();

  // Process ALL sheets
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rawRows.length) continue;

    const nonEmptyRows = rawRows.filter(r => r.some(v => cellStr(v).length > 0));
    if (nonEmptyRows.length < 2) continue;

    // Find header row: scan first 10 rows for one with known column aliases
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(10, rawRows.length); i++) {
      const row = rawRows[i].map(cellStr);
      const matches = row.filter(h =>
        aliasFlat.includes(h.toLowerCase().replace(/[\s_\-\.]/g, ''))
      ).length;
      if (matches >= 1) { headerRowIdx = i; break; }
    }

    const sheetDevices = headerRowIdx >= 0
      ? parseWithHeaders(rawRows, headerRowIdx)
      : parseHeaderless(rawRows);

    results.push(...sheetDevices);
  }

  return results;
}

function parseWithHeaders(allRows, headerRowIdx) {
  const headerRow = allRows[headerRowIdx].map(cellStr);
  const dataRows  = allRows.slice(headerRowIdx + 1);

  const colMap = {};
  for (const field of Object.keys(COL_ALIASES)) {
    const aliases = COL_ALIASES[field];
    const idx = headerRow.findIndex(h =>
      aliases.includes(h.toLowerCase().replace(/[\s_\-\.]/g, ''))
    );
    colMap[field] = idx;
  }

  const mappedIdxs = new Set(Object.values(colMap).filter(i => i >= 0));

  return dataRows
    .filter(row => row.some(v => cellStr(v).length > 0))
    .map(row => {
      const d = {};
      for (const [field, idx] of Object.entries(colMap)) {
        if (idx >= 0) d[field] = cellStr(row[idx]);
      }
      if (!d.serial) {
        row.forEach((v, i) => {
          if (!mappedIdxs.has(i) && !d.serial && SERIAL_RE.test(cellStr(v))) {
            d.serial = cellStr(v);
          }
        });
      }
      return d;
    })
    .filter(d => d.model && looksLikeModel(d.model));
}

function parseHeaderless(rawRows) {
  const numCols = Math.max(...rawRows.map(r => r.length), 0);
  if (numCols === 0) return [];

  const scores = {};
  for (let c = 0; c < numCols; c++) {
    const vals = rawRows.map(r => cellStr(r[c])).filter(v => v.length > 0);
    if (!vals.length) continue;
    const r = n => n / vals.length;
    scores[c] = {
      model:  r(vals.filter(looksLikeModel).length),
      ram:    r(vals.filter(v => RAM_VALS.has(v) || RAM_RE.test(v)).length),
      ssd:    r(vals.filter(v => SSD_VALS.has(v) || SSD_RE.test(v)).length),
      grade:  r(vals.filter(v => GRADE_RE.test(v)).length),
      serial: r(vals.filter(v => SERIAL_RE.test(v)).length),
    };
  }

  const pick = (field, exclude = []) => {
    let best = -1, bestScore = 0.15;
    for (let c = 0; c < numCols; c++) {
      if (exclude.includes(c) || !scores[c]) continue;
      if (scores[c][field] > bestScore) { best = c; bestScore = scores[c][field]; }
    }
    return best;
  };

  const modelCol  = pick('model');
  if (modelCol === -1) return [];

  const serialCol = pick('serial', [modelCol]);
  const ramCol    = pick('ram',    [modelCol, serialCol]);
  const ssdCol    = pick('ssd',    [modelCol, serialCol, ramCol]);
  const gradeCol  = pick('grade',  [modelCol, serialCol, ramCol, ssdCol]);

  return rawRows
    .filter(row => row.some(v => cellStr(v).length > 0))
    .filter(row => looksLikeModel(row[modelCol]))
    .map(row => {
      const d = { model: cellStr(row[modelCol]) };
      if (serialCol >= 0) d.serial = cellStr(row[serialCol]);
      if (ramCol    >= 0) d.ram    = cellStr(row[ramCol]);
      if (ssdCol    >= 0) d.ssd    = cellStr(row[ssdCol]);
      if (gradeCol  >= 0) d.grade  = cellStr(row[gradeCol]);
      return d;
    })
    .filter(d => d.model);
}

// ─── ROUTE 1: Single quote ───────────────────────────────────────────────────
app.post('/api/quote', (req, res) => {
  try {
    const result = calculatePrice(req.body);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 2: Batch file analysis ───────────────────────────────────────────
app.post('/api/analyze', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: 'No file uploaded' });
    const region  = req.body.region || 'EU';
    const devices = parseExcel(req.file.buffer);
    if (!devices.length) return res.status(400).json({ ok: false, error: 'No valid devices found in file' });
    const { results, summary } = analyzeDevices(devices, region);
    res.json({ ok: true, results, summary });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 3: Generate HTML report ──────────────────────────────────────────
app.post('/api/report', (req, res) => {
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
      <tr>
        <th>Model</th><th>Gen</th><th>RAM</th><th>SSD</th><th>Grade</th><th>Status</th><th>ERP</th><th>Price Band</th>
      </tr>
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
