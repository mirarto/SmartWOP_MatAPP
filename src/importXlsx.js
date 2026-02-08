const fs = require('fs');
const Excel = require('exceljs');
const { create } = require('xmlbuilder2');
const { randomUUID } = require('crypto');

function readSheetMap(workbook, sheetName) {
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return [];
  const rows = [];
  const header = [];
  ws.eachRow((row, rowNumber) => {
    const values = row.values; // values[1] is first cell
    if (rowNumber === 1) {
      for (let i = 1; i < values.length; i++) header[i] = values[i];
    } else {
      const obj = {};
      for (let i = 1; i < values.length; i++) {
        const key = header[i];
        if (!key) continue;
        obj[key] = values[i] === undefined ? '' : values[i];
      }
      // include original Excel row number to help locating issues
      obj.__row = rowNumber;
      rows.push(obj);
    }
  });
  return rows;
}

function ensureId(row, idKey, nameKey) {
  if (!row[idKey] || String(row[idKey]).trim() === '') {
    // generate new uuid and ensure it's stored
    row[idKey] = randomUUID();
  }
  // also ensure name exists
  if (!row[nameKey]) row[nameKey] = '';
}

async function importXlsx(xlsxPath, outDbPath, originalDbPath) {
  const wb = new Excel.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const materialsRows = readSheetMap(wb, 'Materials');
  const texturesRows = readSheetMap(wb, 'Textures');
  const panelsRows = readSheetMap(wb, 'Panels');
  const layersRows = readSheetMap(wb, 'Layers');
  const edgesRows = readSheetMap(wb, 'Edges');

  // Only validate material IDs/names: generate missing material IDs first, then check duplicates
  for (const r of materialsRows) ensureId(r, 'material_id', 'material_name');
  // check duplicate material_name and material_id
  const matNameMap = new Map();
  const matIdMap = new Map();
  const matNameDups = new Set();
  const matIdDups = new Set();
  for (const m of materialsRows) {
    const name = String(m.material_name || '').trim();
    const id = String(m.material_id || '').trim();
    if (name) {
      if (matNameMap.has(name)) matNameDups.add(name);
      matNameMap.set(name, (matNameMap.get(name) || 0) + 1);
    }
    if (id) {
      if (matIdMap.has(id)) matIdDups.add(id);
      matIdMap.set(id, (matIdMap.get(id) || 0) + 1);
    }
  }
  if (matNameDups.size > 0) {
    throw new Error('Materials: duplicate material_name values: ' + Array.from(matNameDups).join(', '));
  }
  if (matIdDups.size > 0) {
    throw new Error('Materials: duplicate material_id values: ' + Array.from(matIdDups).join(', '));
  }

  // Ensure IDs
  for (const r of materialsRows) ensureId(r, 'material_id', 'material_name');
  for (const r of panelsRows) ensureId(r, 'panel_id', 'panel_name');
  for (const r of layersRows) ensureId(r, 'layer_id', 'layer_name');
  for (const r of texturesRows) ensureId(r, 'texture_id', 'position');
  for (const r of edgesRows) ensureId(r, 'edge_id', 'name');

  // Group related
  const panelsByMaterial = {};
  for (const p of panelsRows) {
    const mid = p.material_id || '';
    panelsByMaterial[mid] = panelsByMaterial[mid] || [];
    panelsByMaterial[mid].push(p);
  }

  const layersByPanel = {};
  for (const L of layersRows) {
    const pid = L.panel_id || '';
    layersByPanel[pid] = layersByPanel[pid] || [];
    layersByPanel[pid].push(L);
  }

  const texturesByMaterial = {};
  for (const t of texturesRows) {
    const mid = t.material_id || '';
    texturesByMaterial[mid] = texturesByMaterial[mid] || {};
    // position top/bottom
    const pos = (t.position || '').toString().toLowerCase();
    if (pos) texturesByMaterial[mid][pos] = t;
  }

  const edgesByMaterial = {};
  for (const e of edgesRows) {
    const mid = e.material_id || '';
    edgesByMaterial[mid] = edgesByMaterial[mid] || [];
    edgesByMaterial[mid].push(e);
  }

  // Build materials array for xmlbuilder2
  const materialObjs = [];
  for (const m of materialsRows) {
    const mid = m.material_id;
    const mname = m.material_name || '';
    const details = {
      favorite: m.favorite || '',
      name: mname,
      type: m.type || '',
      rotatable: String(m.rotatable || ''),
      path: m.path || '',
      visual_effect: {
        reflect: m.reflect || '',
        rainbown: m.rainbown || '',
        specular: m.specular || '',
        shininess: m.shininess || '',
        glossiness: m.glossiness || '',
        opacity_min: m.opacity_min || '',
        opacity_max: m.opacity_max || ''
      }
    };

    const textures = {};
    const texs = texturesByMaterial[mid] || {};
  if (texs.top) textures.top = { image: texs.top.image || '', angle: texs.top.angle || '', fit_vertically: texs.top.fit_vertically || '' };
  if (texs.bottom) textures.bottom = { image: texs.bottom.image || '', angle: texs.bottom.angle || '', fit_vertically: texs.bottom.fit_vertically || '' };
  // include mirror if present
  if (texs.top && texs.top.mirror) textures.top.mirror = texs.top.mirror;
  if (texs.bottom && texs.bottom.mirror) textures.bottom.mirror = texs.bottom.mirror;

    const panelsList = panelsByMaterial[mid] || [];
    const panelObjs = panelsList.map(p => {
      const pid = p.panel_id;
      const layers = (layersByPanel[pid] || []).map(L => {
        const layerObj = {};
        if (L.layer_name) layerObj.name = L.layer_name;
        if (L.type) layerObj.type = L.type;
        if (L.supplier) layerObj.supplier = L.supplier;
        if (L.thickness) layerObj.thickness = { '#text': L.thickness, '@unit': L.thickness_unit || '' };
        if (L.length) layerObj.length = { '#text': L.length, '@unit': L.length_unit || '' };
        if (L.width) layerObj.width = { '#text': L.width, '@unit': L.width_unit || '' };
        if (L.price) layerObj.price = { '#text': L.price, '@unit': L.price_unit || '' };
        if (L.unprocessed_offset) layerObj.unprocessed_offset = { '#text': L.unprocessed_offset, '@unit': L.unprocessed_offset_unit || '' };
        if (L.outsize) layerObj.outsize = { '#text': L.outsize, '@unit': L.outsize_unit || '' };
        return layerObj;
      });

      const panelObj = {
        '@id': pid,
        name: p.panel_name || '',
        article: p.article || '',
        supplier: p.supplier || '',
        thickness: { '#text': p.thickness || '', '@unit': p.thickness_unit || '' },
        solid_base: { '#text': p.solid_base_name || '', '@id': p.solid_base_id || '' },
        layers: { layer: layers.length === 1 ? layers[0] : layers }
      };
      return panelObj;
    });

    const edgesList = (edgesByMaterial[mid] || []).map(e => {
        const obj = {
          name: e.name || '',
          article: e.article || '',
          supplier: e.supplier || '',
          thickness: { '#text': e.thickness || '', '@unit': e.thickness_unit || '' },
          factory_width: e.factory_width || '',
          price: e.price ? { '#text': e.price, '@unit': e.price_unit || '' } : undefined,
          width_min: { '#text': e.width_min || '', '@unit': e.width_min_unit || '' },
          width_max: { '#text': e.width_max || '', '@unit': e.width_max_unit || '' }
        };
    // assemble visual_effect for edges: only 'angle' is present in DB
    if (e.angle) obj.visual_effect = { angle: e.angle };
      if (e.edge_id) obj['@id'] = e.edge_id;
      return obj;
    });

    const matObj = {
      details: details,
      textures: textures,
      panels: { panel: panelObjs.length === 1 ? panelObjs[0] : panelObjs },
      edges: { edge: edgesList.length === 0 ? undefined : (edgesList.length === 1 ? edgesList[0] : edgesList) },
      '@id': mid
    };

    materialObjs.push(matObj);
  }

  const root = { materials: { '@version': '1.0', material: materialObjs } };

  // backup original if provided
  if (originalDbPath && fs.existsSync(originalDbPath)) {
    const bak = originalDbPath + '.bak.' + Date.now();
    fs.copyFileSync(originalDbPath, bak);
    console.log('Original DB backed up to', bak);
  }

  const doc = create(root);
  const xml = doc.end({ prettyPrint: true });
  fs.writeFileSync(outDbPath, xml, 'utf8');
  console.log('Written XML to', outDbPath);
}

// Generate a short report about the XLSX contents (counts, basic dup checks on Materials)
async function generateReport(xlsxPath) {
  const wb = new Excel.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  const materialsRows = readSheetMap(wb, 'Materials');
  const texturesRows = readSheetMap(wb, 'Textures');
  const panelsRows = readSheetMap(wb, 'Panels');
  const layersRows = readSheetMap(wb, 'Layers');
  const edgesRows = readSheetMap(wb, 'Edges');

  // basic counts
  const report = {
    materials: materialsRows.length,
    textures: texturesRows.length,
    panels: panelsRows.length,
    layers: layersRows.length,
    edges: edgesRows.length,
    duplicateMaterialNames: [],
    duplicateMaterialIds: [],
    panelsPerMaterial: {},
    layersPerPanel: {},
    texturesPerMaterial: {},
    edgesPerMaterial: {}
  };

  // check duplicates in materials and record row numbers
  const nameBuckets = new Map();
  const idBuckets = new Map();
  for (const m of materialsRows) {
    const name = String(m.material_name || '').trim();
    const id = String(m.material_id || '').trim();
    if (!name && m.__row) {
      // mark missing name as special duplicate entry
      report.duplicateMaterialNames.push({ name: null, rows: [m.__row] });
    }
    if (name) {
      if (!nameBuckets.has(name)) nameBuckets.set(name, []);
      nameBuckets.get(name).push({ row: m.__row, id });
    }
    if (id) {
      if (!idBuckets.has(id)) idBuckets.set(id, []);
      idBuckets.get(id).push({ row: m.__row, name });
    }
  }
  for (const [k, arr] of nameBuckets) if (arr.length > 1) report.duplicateMaterialNames.push({ name: k, rows: arr.map(x => x.row), ids: arr.map(x => x.id) });
  for (const [k, arr] of idBuckets) if (arr.length > 1) report.duplicateMaterialIds.push({ id: k, rows: arr.map(x => x.row), names: arr.map(x => x.name) });

  // panels per material
  for (const p of panelsRows) {
    const mid = p.material_id || '__MISSING__';
    report.panelsPerMaterial[mid] = (report.panelsPerMaterial[mid] || 0) + 1;
  }
  // panel name duplicates per material (with rows)
  report.duplicatePanelNames = [];
  const panelBuckets = new Map();
  for (const p of panelsRows) {
    const mid = p.material_id || '__MISSING__';
    const key = `${mid}::${String(p.panel_name || '').trim()}`;
    if (!panelBuckets.has(key)) panelBuckets.set(key, []);
    panelBuckets.get(key).push({ row: p.__row, panel_id: p.panel_id || '', row_full: p });
  }
  for (const [k, arr] of panelBuckets) {
    if (!arr[0]) continue;
    if (arr.length > 1) {
      const [mid, name] = k.split('::');
      report.duplicatePanelNames.push({ sheet: 'Panels', material_id: mid, panel_name: name || null, rows: arr.map(x => x.row), panel_ids: arr.map(x => x.panel_id), rows_full: arr.map(x => x.row_full) });
    }
  }
  // layers per panel
  for (const L of layersRows) {
    const pid = L.panel_id || '__MISSING__';
    report.layersPerPanel[pid] = (report.layersPerPanel[pid] || 0) + 1;
  }
  // layer name duplicates per panel
  report.duplicateLayerNames = [];
  const layerBuckets = new Map();
  for (const L of layersRows) {
    const pid = L.panel_id || '__MISSING__';
    const key = `${pid}::${String(L.layer_name || '').trim()}`;
    if (!layerBuckets.has(key)) layerBuckets.set(key, []);
    layerBuckets.get(key).push({ row: L.__row, layer_id: L.layer_id || '', row_full: L });
  }
  for (const [k, arr] of layerBuckets) {
    if (arr.length > 1) {
      const [pid, name] = k.split('::');
      report.duplicateLayerNames.push({ sheet: 'Layers', panel_id: pid, layer_name: name || null, rows: arr.map(x => x.row), layer_ids: arr.map(x => x.layer_id), rows_full: arr.map(x => x.row_full) });
    }
  }
  // textures per material by position
  for (const t of texturesRows) {
    const mid = t.material_id || '__MISSING__';
    report.texturesPerMaterial[mid] = report.texturesPerMaterial[mid] || { top: 0, bottom: 0, other: 0 };
    const pos = String(t.position || '').toLowerCase();
    if (pos === 'top') report.texturesPerMaterial[mid].top++;
    else if (pos === 'bottom') report.texturesPerMaterial[mid].bottom++;
    else report.texturesPerMaterial[mid].other++;
  }
  // texture duplicates per material+position
  report.duplicateTextures = [];
  const texBuckets = new Map();
  for (const t of texturesRows) {
    const mid = t.material_id || '__MISSING__';
    const pos = String(t.position || '').toLowerCase() || '__none__';
    const key = `${mid}::${pos}`;
    if (!texBuckets.has(key)) texBuckets.set(key, []);
    texBuckets.get(key).push({ row: t.__row, texture_id: t.texture_id || '', row_full: t });
  }
  for (const [k, arr] of texBuckets) {
    if (arr.length > 1) {
      const [mid, pos] = k.split('::');
      report.duplicateTextures.push({ sheet: 'Textures', material_id: mid, position: pos, rows: arr.map(x => x.row), texture_ids: arr.map(x => x.texture_id), rows_full: arr.map(x => x.row_full) });
    }
  }
  // edges per material
  for (const e of edgesRows) {
    const mid = e.material_id || '__MISSING__';
    report.edgesPerMaterial[mid] = (report.edgesPerMaterial[mid] || 0) + 1;
  }
  // edge name duplicates per material
  report.duplicateEdgeNames = [];
  const edgeBuckets = new Map();
  for (const e of edgesRows) {
    const mid = e.material_id || '__MISSING__';
    const key = `${mid}::${String(e.name || '').trim()}`;
    if (!edgeBuckets.has(key)) edgeBuckets.set(key, []);
    edgeBuckets.get(key).push({ row: e.__row, edge_id: e.edge_id || '', row_full: e });
  }
  for (const [k, arr] of edgeBuckets) {
    if (arr.length > 1) {
      const [mid, name] = k.split('::');
      report.duplicateEdgeNames.push({ sheet: 'Edges', material_id: mid, name: name || null, rows: arr.map(x => x.row), edge_ids: arr.map(x => x.edge_id), rows_full: arr.map(x => x.row_full) });
    }
  }

  // include sample rows (first up to 10) per sheet with name/id/row for quick inspection
  report.sample = {
    materials: materialsRows.slice(0, 10).map(r => ({ row: r.__row, material_id: r.material_id || '', material_name: r.material_name || '' })),
    panels: panelsRows.slice(0, 10).map(r => ({ row: r.__row, panel_id: r.panel_id || '', panel_name: r.panel_name || '', material_id: r.material_id || '' })),
    layers: layersRows.slice(0, 10).map(r => ({ row: r.__row, layer_id: r.layer_id || '', layer_name: r.layer_name || '', panel_id: r.panel_id || '' })),
    textures: texturesRows.slice(0, 10).map(r => ({ row: r.__row, texture_id: r.texture_id || '', position: r.position || '', material_id: r.material_id || '' })),
    edges: edgesRows.slice(0, 10).map(r => ({ row: r.__row, edge_id: r.edge_id || '', name: r.name || '', material_id: r.material_id || '' }))
  };

  return report;
}

module.exports = { importXlsx, generateReport };
