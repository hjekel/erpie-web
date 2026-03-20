'use strict';
const XLSX = require('xlsx');
const path = require('path');
const { analyzeDevices } = require('./pricing-engine');

// Import parseExcel and detectFormat from index.js
// We need to extract them — for now, replicate the chain
const indexPath = path.join(__dirname, 'index.js');

// Load the full server module's parser
const fs = require('fs');

// Test files with expected results
const TESTS = [
  {
    name: 'Freshfields Frankfurt (1974)',
    file: path.join(process.env.HOME, 'Downloads/1974 - Freshfields Frankfurth.xlsx'),
    expectedFormat: 'VENDOR_QUOTE',
    expectedDevices: [60, 70],  // range
    expectedERP: [5000, 10000],
  },
  {
    name: 'Freshfields Berlin (1970)',
    file: path.join(process.env.HOME, 'Downloads/1970 - Freshfields Berlin.xlsx'),
    expectedFormat: 'VENDOR_QUOTE',
    expectedDevices: [10, 20],
    expectedERP: [2000, 5000],
  },
  {
    name: 'Freshfields Manchester (1954)',
    file: path.join(process.env.HOME, 'Downloads/1954 - Freshfields Manchester.xlsx'),
    expectedFormat: 'VENDOR_QUOTE',
    expectedDevices: [10, 20],
    expectedERP: [2000, 5000],
  },
  {
    name: 'Freshfields London (1957)',
    file: path.join(process.env.HOME, 'Downloads/1957 Freshfields London.xlsx'),
    expectedFormat: 'VENDOR_QUOTE',
    expectedDevices: [30, 40],
    expectedERP: [4000, 7000],
  },
  {
    name: 'Freshfields Singapore (1962)',
    file: path.join(process.env.HOME, 'Downloads/1962 Zones Freshfields - Singapore.xlsx'),
    expectedFormat: 'VENDOR_QUOTE',
    expectedDevices: [15, 25],
    expectedERP: [2000, 5000],
  },
  {
    name: 'Mender High End Laptops',
    file: path.join(process.env.HOME, 'Downloads/Mender_High_End_Laptop_List .xlsx'),
    expectedFormat: 'GENERIC_HEADERS',
    expectedDevices: [1000, 1200],
    expectedERP: [150000, 350000],
  },
  {
    name: 'Sedgwick Amstelveen Inventory',
    file: path.join(process.env.HOME, 'Downloads/INVENTORY REPORT 2009 ZONES - Sedgwick Amstelveen Redeploy.xlsx'),
    expectedFormat: 'ZONES_INVENTORY',
    expectedDevices: [25, 40],  // unique model groups
    expectedERP: [3000, 7000],
  },
  {
    name: 'Cognizant Frankfurt',
    file: path.join(process.env.HOME, 'Downloads/Cognizant Frankfurth - Asset Details 190326.xlsx'),
    expectedFormat: 'GENERIC_HEADERS',
    expectedDevices: [450, 550],
    expectedERP: [70000, 100000],
  },
];

// We need to use the actual server's parseExcel — load it via require
// But index.js starts Express, so we extract just the parser functions
// Quick approach: use curl against the running server

const { execSync } = require('child_process');

console.log('╔══════════════════════════════════════════════════════════════╗');
console.log('║  ERPIE Universal Parser Test Suite                          ║');
console.log('╚══════════════════════════════════════════════════════════════╝');
console.log('');

let passed = 0, failed = 0, skipped = 0;

for (const test of TESTS) {
  if (!fs.existsSync(test.file)) {
    console.log(`⏭  ${test.name}: FILE NOT FOUND (${path.basename(test.file)})`);
    skipped++;
    continue;
  }

  try {
    const result = execSync(
      `curl -s -X POST "http://89.167.28.106/erpie/api/analyze" -F "file=@${test.file}" -F "region=EU"`,
      { timeout: 30000 }
    ).toString();

    const data = JSON.parse(result);
    if (!data.ok) {
      console.log(`❌ ${test.name}: ERROR — ${data.error}`);
      failed++;
      continue;
    }

    const format = data.format;
    const devices = data.results.length;
    const erp = data.results.reduce((s, r) => s + (r.advisedPrice || 0), 0);

    const formatOk = format === test.expectedFormat;
    const devicesOk = devices >= test.expectedDevices[0] && devices <= test.expectedDevices[1];
    const erpOk = erp >= test.expectedERP[0] && erp <= test.expectedERP[1];

    const status = formatOk && devicesOk && erpOk ? '✅' : '⚠️';
    if (status === '✅') passed++; else failed++;

    console.log(`${status} ${test.name}`);
    console.log(`   Format: ${format} ${formatOk ? '✓' : '✗ (expected ' + test.expectedFormat + ')'}`);
    console.log(`   Devices: ${devices} ${devicesOk ? '✓' : '✗ (expected ' + test.expectedDevices.join('-') + ')'}`);
    console.log(`   ERP: €${erp.toLocaleString()} ${erpOk ? '✓' : '✗ (expected €' + test.expectedERP.map(v => v.toLocaleString()).join('-') + ')'}`);
    console.log('');
  } catch (e) {
    console.log(`❌ ${test.name}: ${e.message.slice(0, 100)}`);
    failed++;
  }
}

console.log('────────────────────────────────────────────────────');
console.log(`Results: ${passed} passed, ${failed} failed, ${skipped} skipped`);
console.log('');
