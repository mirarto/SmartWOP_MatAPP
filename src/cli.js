#!/usr/bin/env node
const path = require('path');
const fs = require('fs');
const { parseXmlFile } = require('./parseXml');
const { generateTemplate } = require('./generateTemplate');
const { importXlsx } = require('./importXlsx');

const argv = process.argv.slice(2);
const cmd = argv[0];

if (cmd === 'parse') {
  const xmlPath = argv[1] || path.join(process.cwd(), 'materials_test.db');
  const out = argv[2] || path.join(process.cwd(), 'materials_parsed.json');
  try {
    const json = parseXmlFile(xmlPath);
    fs.writeFileSync(out, JSON.stringify(json, null, 2), 'utf8');
    console.log('Parsed JSON written to', out);
  } catch (err) {
    console.error('Error parsing:', err.message);
    process.exit(1);
  }

} else if (cmd === 'generate-template') {
  const jsonPath = argv[1] || path.join(process.cwd(), 'materials_parsed.json');
  const xlsxPath = argv[2] || path.join(process.cwd(), 'materials_template.xlsx');
  if (!fs.existsSync(jsonPath)) {
    console.error('JSON file not found:', jsonPath);
    process.exit(1);
  }
  generateTemplate(jsonPath, xlsxPath)
    .then(() => console.log('Template generated:', xlsxPath))
    .catch(err => {
      console.error('Error generating template:', err.message);
      process.exit(1);
    });

} else if (cmd === 'import-xlsx') {
  const xlsxPath = argv[1] || path.join(process.cwd(), 'materials_template.xlsx');
  const outDb = argv[2] || path.join(process.cwd(), 'materials_new.db');
  const originalDb = argv[3] || null;
  if (!fs.existsSync(xlsxPath)) {
    console.error('XLSX file not found:', xlsxPath);
    process.exit(1);
  }
  const force = argv.includes('--force');
  const reportOnly = argv.includes('--report');
  if (reportOnly) {
    // lazy-load report function from importXlsx module
  // allow optional --report-out <path>
  const outIndex = argv.indexOf('--report-out');
  let outPath = null;
  if (outIndex !== -1 && argv.length > outIndex + 1) outPath = argv[outIndex + 1];
  const brief = argv.includes('--report-brief');
  const full = argv.includes('--report-full');
    const { generateReport } = require('./importXlsx');
    generateReport(xlsxPath)
      .then(r => {
        console.log('--- XLSX report ---');
        console.log(`Materials: ${r.materials}`);
        console.log(`Panels: ${r.panels}`);
        console.log(`Layers: ${r.layers}`);
        console.log(`Textures: ${r.textures}`);
        console.log(`Edges: ${r.edges}`);
  if (r.duplicateMaterialNames.length) console.log('Duplicate material names: ' + JSON.stringify(r.duplicateMaterialNames, null, 2));
  if (r.duplicateMaterialIds.length) console.log('Duplicate material ids: ' + JSON.stringify(r.duplicateMaterialIds, null, 2));
  if (r.duplicatePanelNames.length) console.log('Duplicate panel names (per material): ' + JSON.stringify(r.duplicatePanelNames, null, 2));
  if (r.duplicateLayerNames.length) console.log('Duplicate layer names (per panel): ' + JSON.stringify(r.duplicateLayerNames, null, 2));
  if (r.duplicateTextures.length) console.log('Duplicate textures (material+position): ' + JSON.stringify(r.duplicateTextures, null, 2));
  if (r.duplicateEdgeNames.length) console.log('Duplicate edge names (per material): ' + JSON.stringify(r.duplicateEdgeNames, null, 2));
        if (brief) {
          console.log('\nBrief panels/layers/textures/edges per material:');
          Object.keys(r.panelsPerMaterial).slice(0,10).forEach(k => console.log(`  ${k}: panels=${r.panelsPerMaterial[k]}, layers=${Object.values(r.layersPerPanel).reduce((a,b)=>a+b,0)}`));
        } else {
          console.log('\nPanels per material (sample up to 10):');
          Object.keys(r.panelsPerMaterial).slice(0,10).forEach(k => console.log(`  ${k}: ${r.panelsPerMaterial[k]}`));
          console.log('\nLayers per panel (sample up to 10):');
          Object.keys(r.layersPerPanel).slice(0,10).forEach(k => console.log(`  ${k}: ${r.layersPerPanel[k]}`));
          console.log('\nTextures per material (sample up to 10):');
          Object.keys(r.texturesPerMaterial).slice(0,10).forEach(k => console.log(`  ${k}: top=${r.texturesPerMaterial[k].top}, bottom=${r.texturesPerMaterial[k].bottom}, other=${r.texturesPerMaterial[k].other}`));
          console.log('\nEdges per material (sample up to 10):');
          Object.keys(r.edgesPerMaterial).slice(0,10).forEach(k => console.log(`  ${k}: ${r.edgesPerMaterial[k]}`));
        }

            // write JSON if requested or default to ./reports/report-<ts>.json
            if (!outPath) {
              const ts = Date.now();
              const dir = path.join(process.cwd(), 'reports');
              if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
              outPath = path.join(dir, `report-${ts}.json`);
            }
        fs.writeFileSync(outPath, JSON.stringify(r, null, 2), 'utf8');
        console.log('\nReport saved to', outPath);
      })
      .catch(err => {
        console.error('Error generating report:', err.message);
        process.exit(1);
      });
    return;
  }

  importXlsx(xlsxPath, outDb, originalDb)
    .then((finalPath) => console.log('Import finished:', finalPath))
    .catch(err => {
      if (String(err.message).includes('Materials: duplicate') && !force) {
        console.error('\nImport aborted due to duplicate materials. Re-run with --force to override (not recommended).');
        process.exit(2);
      }
      console.error('Error importing:', err.message);
      process.exit(1);
    });

} else {
  console.log('Usage: node src/cli.js <command> [args]\n\nCommands:\n  parse [xmlPath] [outJson]        Parse XML .db to JSON\n  generate-template [jsonPath] [outXlsx]   Generate Excel template from JSON\n  import-xlsx [xlsxPath] [outDb] [originalDb]  Import changed XLSX and write .db (originalDb optional for backup)');
}


