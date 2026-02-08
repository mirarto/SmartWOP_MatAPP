const fs = require('fs');
const Excel = require('exceljs');

function asArray(x) {
  if (x === undefined || x === null) return [];
  return Array.isArray(x) ? x : [x];
}

function getTextOrValue(node) {
  if (node === undefined || node === null) return '';
  if (typeof node === 'object') {
    if (node['#text'] !== undefined) return node['#text'];
    return '';
  }
  return node;
}

async function generateTemplate(jsonPath, xlsxPath) {
  const raw = fs.readFileSync(jsonPath, 'utf8');
  const data = JSON.parse(raw);
  const materialsNode = data.materials && data.materials.material ? data.materials.material : [];
  const materials = Array.isArray(materialsNode) ? materialsNode : [materialsNode];

  const wb = new Excel.Workbook();

  const mSheet = wb.addWorksheet('Materials');
  mSheet.columns = [
    { header: 'material_id', key: 'material_id', width: 40 },
    { header: 'material_name', key: 'material_name', width: 30 },
    { header: 'favorite', key: 'favorite', width: 8 },
    { header: 'type', key: 'type', width: 12 },
    { header: 'rotatable', key: 'rotatable', width: 10 },
    { header: 'path', key: 'path', width: 30 },
    { header: 'reflect', key: 'reflect', width: 10 },
    { header: 'rainbown', key: 'rainbown', width: 10 },
    { header: 'specular', key: 'specular', width: 10 },
    { header: 'shininess', key: 'shininess', width: 10 },
    { header: 'glossiness', key: 'glossiness', width: 10 },
    { header: 'opacity_min', key: 'opacity_min', width: 10 },
    { header: 'opacity_max', key: 'opacity_max', width: 10 }
  ];

  const textures = wb.addWorksheet('Textures');
  textures.columns = [
    { header: 'texture_id', key: 'texture_id', width: 40 },
    { header: 'material_id', key: 'material_id', width: 40 },
    { header: 'material_name', key: 'material_name', width: 30 },
    { header: 'position', key: 'position', width: 10 },
    { header: 'image', key: 'image', width: 60 },
    { header: 'angle', key: 'angle', width: 8 },
    { header: 'fit_vertically', key: 'fit_vertically', width: 12 }
    ,{ header: 'mirror', key: 'mirror', width: 10 }
  ];

  const panels = wb.addWorksheet('Panels');
  panels.columns = [
    { header: 'panel_id', key: 'panel_id', width: 40 },
    { header: 'material_id', key: 'material_id', width: 40 },
    { header: 'material_name', key: 'material_name', width: 30 },
    { header: 'panel_name', key: 'panel_name', width: 30 },
    { header: 'article', key: 'article', width: 20 },
    { header: 'supplier', key: 'supplier', width: 20 },
    { header: 'thickness', key: 'thickness', width: 10 },
    { header: 'thickness_unit', key: 'thickness_unit', width: 8 },
    { header: 'solid_base_id', key: 'solid_base_id', width: 40 },
    { header: 'solid_base_name', key: 'solid_base_name', width: 30 }
  ];

  const layers = wb.addWorksheet('Layers');
  layers.columns = [
    { header: 'layer_id', key: 'layer_id', width: 40 },
    { header: 'panel_id', key: 'panel_id', width: 40 },
    { header: 'panel_name', key: 'panel_name', width: 30 },
    { header: 'layer_name', key: 'layer_name', width: 30 },
    { header: 'thickness', key: 'thickness', width: 10 },
    { header: 'thickness_unit', key: 'thickness_unit', width: 8 },
    { header: 'type', key: 'type', width: 12 },
    { header: 'supplier', key: 'supplier', width: 15 },
    { header: 'length', key: 'length', width: 12 },
    { header: 'length_unit', key: 'length_unit', width: 8 },
    { header: 'width', key: 'width', width: 12 },
    { header: 'width_unit', key: 'width_unit', width: 8 },
    { header: 'price', key: 'price', width: 12 },
    { header: 'price_unit', key: 'price_unit', width: 8 },
    { header: 'unprocessed_offset', key: 'unprocessed_offset', width: 12 },
    { header: 'unprocessed_offset_unit', key: 'unprocessed_offset_unit', width: 8 },
    { header: 'outsize', key: 'outsize', width: 8 },
    { header: 'outsize_unit', key: 'outsize_unit', width: 8 }
  ];

  const edges = wb.addWorksheet('Edges');
  edges.columns = [
    { header: 'edge_id', key: 'edge_id', width: 40 },
    { header: 'material_id', key: 'material_id', width: 40 },
    { header: 'material_name', key: 'material_name', width: 30 },
    { header: 'name', key: 'name', width: 30 },
    { header: 'article', key: 'article', width: 30 },
    { header: 'supplier', key: 'supplier', width: 15 },
  { header: 'factory_width', key: 'factory_width', width: 12 },
  { header: 'angle', key: 'angle', width: 8 },
    { header: 'thickness', key: 'thickness', width: 10 },
    { header: 'thickness_unit', key: 'thickness_unit', width: 8 },
    { header: 'price', key: 'price', width: 12 },
    { header: 'price_unit', key: 'price_unit', width: 8 },
    { header: 'width_min', key: 'width_min', width: 10 },
    { header: 'width_min_unit', key: 'width_min_unit', width: 8 },
    { header: 'width_max', key: 'width_max', width: 10 },
    { header: 'width_max_unit', key: 'width_max_unit', width: 8 }
  ];

  for (const m of materials) {
    const mid = m['@_id'] || '';
    const mname = (m.details && m.details.name) || '';
    const det = m.details || {};
    const ve = det.visual_effect || {};

    mSheet.addRow({
      material_id: mid,
      material_name: mname,
      favorite: det.favorite || '',
      type: det.type || '',
      rotatable: det.rotatable || '',
      path: det.path || '',
      reflect: ve.reflect || '',
      rainbown: ve.rainbown || '',
      specular: ve.specular || '',
      shininess: ve.shininess || '',
      glossiness: ve.glossiness || '',
      opacity_min: ve.opacity_min || '',
      opacity_max: ve.opacity_max || ''
    });

    // textures
    if (m.textures) {
      for (const pos of ['top', 'bottom']) {
        if (m.textures[pos]) {
          const t = m.textures[pos];
          textures.addRow({
            texture_id: '',
            material_id: mid,
            material_name: mname,
            position: pos,
            image: getTextOrValue(t.image) || t.image || '',
            angle: getTextOrValue(t.angle) || t.angle || '',
            fit_vertically: (t.fit_vertically !== undefined) ? String(t.fit_vertically) : '',
            mirror: (t.mirror !== undefined) ? String(t.mirror) : ''
          });
        }
      }
    }

    // panels and layers
    if (m.panels && m.panels.panel) {
      const panelArr = asArray(m.panels.panel);
      for (const p of panelArr) {
        const pid = p['@_id'] || '';
        panels.addRow({
          panel_id: pid,
          material_id: mid,
          material_name: mname,
          panel_name: p.name || '',
          article: p.article || '',
          supplier: p.supplier || '',
          thickness: getTextOrValue(p.thickness) || (p.thickness && p.thickness['#text']) || '',
          thickness_unit: (p.thickness && p.thickness['@_unit']) || '' ,
          solid_base_id: (p.solid_base && p.solid_base['@_id']) || '',
          solid_base_name: (p.solid_base && p.solid_base['#text']) || p.solid_base || ''
        });

        if (p.layers && p.layers.layer) {
          const layerArr = asArray(p.layers.layer);
          for (const L of layerArr) {
            layers.addRow({
              layer_id: L['@_id'] || '',
              panel_id: pid,
              panel_name: p.name || '',
              layer_name: L.name || '',
                thickness: getTextOrValue(L.thickness) || '',
                thickness_unit: (L.thickness && L.thickness['@_unit']) || '',
              type: L.type || '',
              supplier: L.supplier || '',
              length: getTextOrValue(L.length) || '',
              length_unit: (L.length && L.length['@_unit']) || '',
              width: getTextOrValue(L.width) || '',
              width_unit: (L.width && L.width['@_unit']) || '',
              price: getTextOrValue(L.price) || '',
              price_unit: (L.price && L.price['@_unit']) || '',
                unprocessed_offset: getTextOrValue(L.unprocessed_offset) || '',
                unprocessed_offset_unit: (L.unprocessed_offset && L.unprocessed_offset['@_unit']) || '',
                outsize: getTextOrValue(L.outsize) || '',
                outsize_unit: (L.outsize && L.outsize['@_unit']) || ''
            });
          }
        }
      }
    }

    // edges
    if (m.edges && m.edges.edge) {
      const edgeArr = asArray(m.edges.edge);
      for (const e of edgeArr) {
        edges.addRow({
          edge_id: e['@_id'] || '',
          material_id: mid,
          material_name: mname,
          name: e.name || '',
          article: e.article || '',
          supplier: e.supplier || '',
          factory_width: getTextOrValue(e.factory_width) || '',
          factory_width_unit: (e.factory_width && e.factory_width['@_unit']) || '',
          angle: (e.visual_effect && e.visual_effect.angle) || '',
          thickness: getTextOrValue(e.thickness) || '',
          thickness_unit: (e.thickness && e.thickness['@_unit']) || '',
          price: getTextOrValue(e.price) || '',
          price_unit: (e.price && e.price['@_unit']) || '',
          width_min: getTextOrValue(e.width_min) || '',
          width_min_unit: (e.width_min && e.width_min['@_unit']) || '',
          width_max: getTextOrValue(e.width_max) || '',
          width_max_unit: (e.width_max && e.width_max['@_unit']) || ''
        });
      }
    }
  }

  await wb.xlsx.writeFile(xlsxPath);
}

module.exports = { generateTemplate };
