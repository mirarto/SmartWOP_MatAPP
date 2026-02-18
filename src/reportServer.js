const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { exec } = require('child_process');

const app = express();
app.use(express.json());

const reportsDir = path.join(process.cwd(), 'reports');

// import business logic to expose via HTTP
const { importXlsx, generateReport } = require('./importXlsx');
const { generateTemplate } = require('./generateTemplate');
const { parseXmlFile } = require('./parseXml');
const multer = require('multer');
const upload = multer({ dest: path.join(os.tmpdir(), 'matapp_uploads') });

function latestReportFile() {
  if (!fs.existsSync(reportsDir)) return null;
  const files = fs.readdirSync(reportsDir)
    .filter(f => f.startsWith('report-') && f.endsWith('.json'))
    .map(f => ({ f, t: fs.statSync(path.join(reportsDir, f)).mtimeMs }))
    .sort((a,b)=>b.t-a.t);
  return files.length ? path.join(reportsDir, files[0].f) : null;
}

app.get('/report', (req, res) => {
  const file = req.query.file || latestReportFile();
  if (!file) return res.status(404).json({ error: 'No report found' });
  const full = path.isAbsolute(file) ? file : path.join(process.cwd(), file);
  if (!fs.existsSync(full)) return res.status(404).json({ error: 'Report not found: ' + full });
  const json = JSON.parse(fs.readFileSync(full, 'utf8'));
  res.json(json);
});

// Open Excel and select sheet and row using PowerShell COM automation (Windows only)
app.post('/open', (req, res) => {
  const { filePath, sheetName, row } = req.body;
  if (!filePath || !sheetName || !row) return res.status(400).json({ error: 'filePath, sheetName and row required' });
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Excel file not found: ' + filePath });

  // Build PowerShell script and write to a temp file to avoid quoting issues
  const fileEsc = filePath.replace(/'/g, "''");
  const sheetEsc = sheetName.replace(/'/g, "''");
  const psLines = [];
  psLines.push(`$excel = New-Object -ComObject Excel.Application`);
  psLines.push(`$excel.Visible = $true`);
  psLines.push(`$wb = $excel.Workbooks.Open('${fileEsc}')`);
  psLines.push(`try { $ws = $wb.Worksheets.Item('${sheetEsc}') } catch { $ws = $wb.Worksheets.Item(1) }`);
  psLines.push(`$ws.Activate()`);
  psLines.push(`$rng = $ws.Range("A${row}")`);
  psLines.push(`$rng.Select()`);

  const psScript = psLines.join(os.EOL);
  const tmp = path.join(os.tmpdir(), `open_excel_${Date.now()}.ps1`);
  try {
    fs.writeFileSync(tmp, psScript, 'utf8');
  } catch (err) {
    return res.status(500).json({ error: 'Failed to write temp PowerShell script: ' + String(err) });
  }

  const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${tmp}"`;
  exec(cmd, { windowsHide: true }, (err, stdout, stderr) => {
    // best-effort cleanup of temp file
    try { fs.unlinkSync(tmp); } catch (e) {}
    if (err) return res.status(500).json({ error: String(err), stderr });
    res.json({ ok: true });
  });
});

// Native file/folder dialogs (Windows only). Return selected path as plain text in { path }
function runPSScript(script, callback) {
  const tmp = path.join(os.tmpdir(), `ps_dialog_${Date.now()}.ps1`);
  fs.writeFileSync(tmp, script, 'utf8');
  const cmd = `powershell -NoProfile -ExecutionPolicy Bypass -File "${tmp}"`;
  exec(cmd, { windowsHide: false }, (err, stdout, stderr) => {
    try { fs.unlinkSync(tmp); } catch (e) {}
    if (err) return callback(err, null, stderr);
    return callback(null, stdout ? String(stdout).trim() : '', stderr);
  });
}

app.get('/dialog/open-file', (req, res) => {
  // optional ?ext=xlsx or ext=db to set filter
  const ext = (req.query.ext || '').toLowerCase();
  let filter = "All files (*.*)|*.*";
  if (ext === 'xlsx') filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
  if (ext === 'db') filter = "DB files (*.db)|*.db|All files (*.*)|*.*";
  // build PowerShell script that forces the dialog to be topmost by creating a temporary hidden Form
  const script = [];
  script.push("Add-Type -AssemblyName System.Windows.Forms");
  script.push("$form = New-Object System.Windows.Forms.Form");
  script.push("$form.TopMost = $true");
  script.push("$form.Size = New-Object System.Drawing.Size(0,0)");
  script.push("$form.ShowInTaskbar = $false");
  script.push("$form.Opacity = 0");
  script.push("$ofd = New-Object System.Windows.Forms.OpenFileDialog");
  script.push(`$ofd.Filter = '${filter.replace(/'/g, "''")}'`);
  script.push("$ofd.Multiselect = $false");
  script.push("if ($ofd.ShowDialog($form) -eq 'OK') { Write-Output $ofd.FileName }");
  script.push("$form.Dispose()");
  runPSScript(script.join(os.EOL), (err, out, stderr) => {
    if (err) return res.status(500).json({ error: String(err), stderr });
    res.json({ path: out });
  });
});

app.get('/dialog/open-folder', (req, res) => {
  const script = [];
  script.push("Add-Type -AssemblyName System.Windows.Forms");
  script.push("$form = New-Object System.Windows.Forms.Form");
  script.push("$form.TopMost = $true");
  script.push("$form.Size = New-Object System.Drawing.Size(0,0)");
  script.push("$form.ShowInTaskbar = $false");
  script.push("$form.Opacity = 0");
  script.push("$f = New-Object System.Windows.Forms.FolderBrowserDialog");
  script.push("if ($f.ShowDialog($form) -eq 'OK') { Write-Output $f.SelectedPath }");
  script.push("$form.Dispose()");
  runPSScript(script.join(os.EOL), (err, out, stderr) => {
    if (err) return res.status(500).json({ error: String(err), stderr });
    res.json({ path: out });
  });
});

// Generate template via HTTP (calls generateTemplate(jsonPath, xlsxPath))
app.post('/generate-template', async (req, res) => {
  const { jsonPath, xlsxPath } = req.body || {};
  if (!jsonPath || !xlsxPath) return res.status(400).json({ error: 'jsonPath and xlsxPath required' });
  try {
    await generateTemplate(jsonPath, xlsxPath);
    res.json({ ok: true, xlsxPath });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

// Generate template directly from a .db (XML) file path
app.post('/api/generate-template-from-db', async (req, res) => {
  const { dbPath, xlsxPath } = req.body || {};
  if (!dbPath || !xlsxPath) return res.status(400).json({ error: 'dbPath and xlsxPath required' });
  if (!fs.existsSync(dbPath)) return res.status(404).json({ error: 'DB file not found: ' + dbPath });
  try {
    const parsed = parseXmlFile(dbPath);
    const tmpJson = path.join(os.tmpdir(), 'matapp_parsed_' + Date.now() + '.json');
    fs.writeFileSync(tmpJson, JSON.stringify(parsed, null, 2), 'utf8');
    await generateTemplate(tmpJson, xlsxPath);
    // cleanup tmp json
    try { fs.unlinkSync(tmpJson); } catch (e) {}
    res.json({ ok: true, xlsxPath });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

// Generate template from uploaded DB file and return XLSX as download
app.post('/api/generate-template-upload', upload.single('dbfile'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'dbfile required (multipart/form-data field name: dbfile)' });
  const uploadedDb = req.file.path;
  const tmpJson = path.join(os.tmpdir(), 'matapp_parsed_' + Date.now() + '.json');
  const tmpXlsx = path.join(os.tmpdir(), 'matapp_template_' + Date.now() + '.xlsx');
  try {
    const parsed = parseXmlFile(uploadedDb);
    fs.writeFileSync(tmpJson, JSON.stringify(parsed, null, 2), 'utf8');
    await generateTemplate(tmpJson, tmpXlsx);
    // send xlsx as attachment
    res.download(tmpXlsx, 'materials_template.xlsx', (err) => {
      // cleanup
      try { fs.unlinkSync(tmpJson); } catch (e) {}
      try { fs.unlinkSync(tmpXlsx); } catch (e) {}
      try { fs.unlinkSync(uploadedDb); } catch (e) {}
      if (err) console.error('Error sending generated xlsx', err);
    });
  } catch (err) {
    try { fs.unlinkSync(tmpJson); } catch (e) {}
    try { fs.unlinkSync(uploadedDb); } catch (e) {}
    res.status(500).json({ error: String(err) });
  }
});

// Upload XLSX and return a preview report (no commit)
app.post('/api/upload-xlsx-preview', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'file required (multipart/form-data field name: file)' });
  const uploaded = req.file.path;
  try {
    const report = await generateReport(uploaded);
    // optionally remove uploaded file after report generated
    try { fs.unlinkSync(uploaded); } catch (e) {}
    res.json({ ok: true, report });
  } catch (err) {
    try { fs.unlinkSync(uploaded); } catch (e) {}
    res.status(500).json({ error: String(err) });
  }
});

// Import XLSX via HTTP (calls importXlsx and generateReport), optional reportFolder
app.post('/import', async (req, res) => {
  const { xlsxPath, outDbPath, originalDbPath, reportFolder, reportFull } = req.body || {};
  if (!xlsxPath || !outDbPath) return res.status(400).json({ error: 'xlsxPath and outDbPath required' });
  try {
    await importXlsx(xlsxPath, outDbPath, originalDbPath);
    // generate report
    const report = await generateReport(xlsxPath);
    const folder = reportFolder ? String(reportFolder) : reportsDir;
    if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
    const reportPath = path.join(folder, `report-${Date.now()}.json`);
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), 'utf8');
    res.json({ ok: true, outDbPath, reportPath, report });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

// serve simple UI
app.use('/', express.static(path.join(__dirname, '..', 'ui')));

const port = process.env.PORT || 3000;
app.listen(port, ()=>console.log('Report server running on http://localhost:' + port));
