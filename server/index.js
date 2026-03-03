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

// ─── COLUMN NAME ALIASES ───────────────────────────────────────────────────────
const COL_ALIASES = {
  model:   ['model', 'device', 'product', 'description', 'item'],
  cpu:     ['cpu', 'processor', 'proc'],
  ram:     ['ram', 'memory', 'mem'],
  ssd:     ['ssd', 'storage', 'hdd', 'disk', 'drive'],
  grade:   ['grade', 'condition', 'cond', 'quality'],
  battery: ['battery', 'bat'],
  serial:  ['serial', 'serialnumber', 'sn', 'asset', 'assettag'],
  qty:     ['qty', 'quantity', 'count', 'units'],
};

function resolveCol(headers, field) {
  const aliases = COL_ALIASES[field] || [field];
  for (const alias of aliases) {
    const found = headers.find(h => h.toLowerCase().replace(/[\s_-]/g, '') === alias.replace(/[\s_-]/g, ''));
    if (found) return found;
  }
  return null;
}

function parseExcel(buffer) {
  const wb    = XLSX.read(buffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  if (!rows.length) return [];

  const headers = Object.keys(rows[0]);
  const colMap  = {};
  for (const field of Object.keys(COL_ALIASES)) {
    colMap[field] = resolveCol(headers, field);
  }

  return rows.map(row => {
    const device = {};
    for (const [field, col] of Object.entries(colMap)) {
      if (col) device[field] = row[col];
    }
    return device;
  }).filter(d => d.model);
}

// ─── ROUTE 1: Single quote ─────────────────────────────────────────────────────
app.post('/api/quote', (req, res) => {
  try {
    const result = calculatePrice(req.body);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(400).json({ ok: false, error: err.message });
  }
});

// ─── ROUTE 2: Batch file analysis ─────────────────────────────────────────────
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

// ─── ROUTE 3: Generate HTML report ────────────────────────────────────────────
app.post('/api/report', (req, res) => {
  try {
    const { devices = [], summary = {}, dealName = 'ERPIE Deal' } = req.body;
    const html = generateReport(devices, summary, dealName);
    res.type('text/html').send(html);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ─── HTML REPORT GENERATOR ─────────────────────────────────────────────────────
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
<title>ERPIE Report — ${escHtml(dealName)}</title>
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
  <h1>PlanBit — ERPIE Price Report</h1>
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
        <td colspan="6"><strong>TOTAAL (${(summary.total || 0)} devices)</strong></td>
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

// ─── START ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8000;
app.listen(PORT, () => console.log(`ERPIE PriceFinder running on http://localhost:${PORT}`));
