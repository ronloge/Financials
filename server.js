const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Simple logging
function log(level, message, data = null) {
  const timestamp = new Date().toISOString();
  const logEntry = `[${timestamp}] ${level.toUpperCase()}: ${message}`;
  console.log(logEntry + (data ? ` - ${JSON.stringify(data)}` : ''));
  
  // Write to log file
  const logFile = 'app.log';
  const logLine = logEntry + (data ? ` - ${JSON.stringify(data)}` : '') + '\n';
  fs.appendFileSync(logFile, logLine);
}

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.static('public'));
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage: storage,
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB limit
});

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/analyze', (req, res) => {
  try {
    log('info', 'Direct analysis request received');
    
    // Check if we have cached Excel data
    const excelDataFile = 'cached_excel_data.json';
    if (!fs.existsSync(excelDataFile)) {
      return res.status(400).json({ error: 'No cached Excel data found. Please upload Excel file first.' });
    }
    
    const allData = JSON.parse(fs.readFileSync(excelDataFile, 'utf8'));
    let config = loadConfig();
    
    // Override config with current GUI settings if provided
    if (req.body.thresholds) {
      config.thresholds = { ...config.thresholds, ...req.body.thresholds };
    }
    
    // Force reload exclusions cache to ensure we have latest data
    exclusionsCache = null;
    exclusionsLastModified = null;
    
    // Load cached engineers and SA filters from persistent files
    let cachedEngineers = null;
    let cachedSAs = null;
    
    if (fs.existsSync('engineers.txt')) {
      try {
        const content = fs.readFileSync('engineers.txt', 'utf8');
        cachedEngineers = content.split('\n')
          .map(line => line.trim().replace(/\s+/g, ' '))
          .filter(line => line && line.length > 1);
      } catch (error) {
        log('error', 'Error loading cached engineers', { error: error.message });
      }
    }
    
    if (fs.existsSync('solution_architects.txt')) {
      try {
        const content = fs.readFileSync('solution_architects.txt', 'utf8');
        cachedSAs = content.split('\n')
          .map(line => line.trim())
          .filter(line => line);
      } catch (error) {
        log('error', 'Error loading cached SAs', { error: error.message });
      }
    }
    
    // Load existing exclusions and filters
    const existingExclusions = loadExistingExclusions();
    const filters = { 
      exclusions: existingExclusions, 
      engineers: cachedEngineers, 
      solutionArchitects: cachedSAs 
    };
    
    log('info', 'Direct analysis filters loaded', { 
      exclusionsCount: existingExclusions.length,
      engineersCount: cachedEngineers ? cachedEngineers.length : 0,
      saCount: cachedSAs ? cachedSAs.length : 0
    });
    
    // Apply date filtering if enabled
    let filteredData = allData;
    if (config.project_filtering?.enable_date_filter) {
      filteredData = applyDateFilter(allData, config.project_filtering);
    }
    
    // Analysis with filters and config
    const analysis = analyzeData(filteredData, filters, config);
    
    log('info', 'Direct analysis completed successfully', { 
      consultants: analysis.consultants.length,
      solutionArchitects: analysis.solutionArchitects.length,
      totalProjects: analysis.totalProjects
    });
    
    res.json({
      success: true,
      analysis: analysis
    });
  } catch (error) {
    log('error', 'Direct analysis failed', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.post('/upload', upload.fields([
  { name: 'excelFiles', maxCount: 10 },
  { name: 'engineersFile', maxCount: 1 },
  { name: 'excludeFile', maxCount: 1 },
  { name: 'saFile', maxCount: 1 }
]), (req, res) => {
  try {
    log('info', 'Analysis request received', { 
      fileCount: req.files?.excelFiles?.length || 0,
      hasEngineersFile: !!req.files?.engineersFile,
      hasExcludeFile: !!req.files?.excludeFile,
      hasSAFile: !!req.files?.saFile
    });
    
    if (!req.files || !req.files.excelFiles) {
      log('error', 'No Excel files uploaded');
      return res.status(400).json({ error: 'No Excel files uploaded' });
    }

    let config = loadConfig();
    
    // Override config with current GUI settings if provided
    if (req.body.thresholds) {
      try {
        const thresholds = JSON.parse(req.body.thresholds);
        config.thresholds = { ...config.thresholds, ...thresholds };
      } catch (e) {
        log('error', 'Error parsing thresholds', { error: e.message });
      }
    }
    const isMultiFile = req.files.excelFiles.length > 1 && config.trending?.enable_trending;
    
    let allData = [];
    let filenames = [];
    
    // Process each Excel file
    for (const file of req.files.excelFiles) {
      try {
        log('info', 'Processing Excel file', { filename: file.filename, path: file.path, exists: fs.existsSync(file.path) });
        
        if (!fs.existsSync(file.path)) {
          log('error', 'Excel file not found', { path: file.path, filename: file.filename });
          return res.status(400).json({ error: `Excel file not found: ${file.filename}` });
        }
        
        const workbook = XLSX.readFile(file.path);
        if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
          log('error', 'Invalid Excel workbook', { filename: file.filename });
          return res.status(400).json({ error: `Invalid Excel file: ${file.filename}` });
        }
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (!data || data.length === 0) {
          log('error', 'Empty Excel file', { filename: file.filename });
          return res.status(400).json({ error: `Empty Excel file: ${file.filename}` });
        }
        
        const headerRowIndex = findHeaderRow(data);
        const headers = data[headerRowIndex];
        
        if (!headers || headers.length === 0) {
          log('error', 'No headers found in Excel file', { filename: file.filename, headerRowIndex });
          return res.status(400).json({ error: `No valid headers found in: ${file.filename}` });
        }
        
        const jsonData = data.slice(headerRowIndex + 1).map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
        
        if (isMultiFile) {
          // Add file identifier for trending
          jsonData.forEach(row => row._fileSource = file.filename);
        }
        
        allData = allData.concat(jsonData);
        filenames.push(file.filename);
        log('info', 'Excel file processed successfully', { filename: file.filename, rows: jsonData.length, headers: headers.length });
        
        // Cache Excel data for direct analysis
        fs.writeFileSync('cached_excel_data.json', JSON.stringify(allData, null, 2));
      } catch (fileError) {
        log('error', 'Error processing Excel file', { filename: file.filename, error: fileError.message, stack: fileError.stack });
        return res.status(400).json({ error: `Error processing ${file.filename}: ${fileError.message}` });
      }
    }

    // Load filter files and config
    const textInputs = {
      engineersText: req.body.engineersText,
      exclusionsText: req.body.exclusionsText,
      saText: req.body.saText
    };
    
    // Always load existing exclusions from CSV file first
    const existingExclusions = loadExistingExclusions();
    
    // Load other filters (skip uploaded exclusion file since we use CSV)
    const filters = loadFilterFiles(req.files, textInputs);
    filters.exclusions = [...filters.exclusions, ...existingExclusions];
    
    // Apply date filtering if enabled
    if (config.project_filtering?.enable_date_filter) {
      allData = applyDateFilter(allData, config.project_filtering);
    }
    
    // Analysis with filters and config
    const analysis = analyzeData(allData, filters, config);
    
    log('info', 'Analysis completed successfully', { 
      consultants: analysis.consultants.length,
      solutionArchitects: analysis.solutionArchitects.length,
      totalProjects: analysis.totalProjects
    });
    
    res.json({
      success: true,
      filenames: filenames,
      isMultiFile: isMultiFile,
      analysis: analysis
    });
  } catch (error) {
    log('error', 'Analysis failed', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.post('/config', (req, res) => {
  try {
    const config = req.body;
    fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
    res.json({ success: true, message: 'Configuration saved' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/config', (req, res) => {
  try {
    if (fs.existsSync('config.json')) {
      const config = JSON.parse(fs.readFileSync('config.json', 'utf8'));
      res.json(config);
    } else {
      res.json(getDefaultConfig());
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/exclusions', (req, res) => {
  try {
    const exclusions = loadExistingExclusions();
    res.json(exclusions);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.delete('/exclusions/cache', (req, res) => {
  try {
    exclusionsCache = null;
    exclusionsLastModified = null;
    engineersCache = null;
    engineersLastModified = null;
    saCache = null;
    saLastModified = null;
    log('info', 'All filter caches cleared');
    res.json({ success: true, message: 'All filter caches cleared' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/exclude', (req, res) => {
  try {
    const { consultant, project, reason } = req.body;
    
    if (!consultant || !project) {
      log('error', 'Missing consultant or project in exclusion request', req.body);
      return res.status(400).json({ error: 'Consultant and project are required' });
    }
    
    if (!reason || reason.trim().length < 2) {
      log('error', 'Missing or invalid reason in exclusion request', req.body);
      return res.status(400).json({ error: 'Reason is required and must be at least 2 characters' });
    }
    
    const exclusionFile = 'exclusions.csv';
    
    let content = '';
    if (fs.existsSync(exclusionFile)) {
      content = fs.readFileSync(exclusionFile, 'utf8');
    } else {
      content = 'Consultant,Project,Reason\n';
    }
    
    // Check if header needs updating
    const lines = content.split('\n');
    if (lines[0] && !lines[0].includes('Reason')) {
      lines[0] = 'Consultant,Project,Reason';
      content = lines.join('\n');
    }
    
    const newExclusion = `${consultant},${project},"${reason.replace(/"/g, '""')}"`;
    const existingLine = lines.find(line => line.includes(`${consultant},${project}`));
    
    if (!existingLine) {
      if (!content.endsWith('\n')) content += '\n';
      content += newExclusion + '\n';
      fs.writeFileSync(exclusionFile, content);
      
      // Update cache
      if (!exclusionsCache) exclusionsCache = [];
      exclusionsCache.push({ consultant, project, reason });
      exclusionsLastModified = new Date();
      
      log('info', 'Exclusion added', { consultant, project, reason });
    } else {
      log('info', 'Exclusion already exists', { consultant, project });
    }
    
    res.json({ success: true });
  } catch (error) {
    log('error', 'Error adding exclusion', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.delete('/exclude', (req, res) => {
  try {
    const { consultant, project } = req.body;
    
    if (!consultant || !project) {
      log('error', 'Missing consultant or project in exclusion removal request', req.body);
      return res.status(400).json({ error: 'Consultant and project are required' });
    }
    
    const exclusionFile = 'exclusions.csv';
    
    if (fs.existsSync(exclusionFile)) {
      let content = fs.readFileSync(exclusionFile, 'utf8');
      const targetLine = `${consultant},${project}`;
      const lines = content.split('\n');
      const filteredLines = lines.filter(line => !line.includes(targetLine));
      
      fs.writeFileSync(exclusionFile, filteredLines.join('\n'));
      
      // Update cache
      if (exclusionsCache) {
        exclusionsCache = exclusionsCache.filter(excl => 
          !(excl.consultant === consultant && excl.project === project)
        );
        exclusionsLastModified = new Date();
      }
      
      log('info', 'Exclusion removed', { consultant, project });
    } else {
      log('info', 'No exclusions file found', { consultant, project });
    }
    
    res.json({ success: true });
  } catch (error) {
    log('error', 'Error removing exclusion', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.post('/disqualify', (req, res) => {
  try {
    const { consultant } = req.body;
    
    if (!consultant) {
      return res.status(400).json({ error: 'Consultant name is required' });
    }
    
    const dqFile = 'disqualified.txt';
    let content = '';
    
    if (fs.existsSync(dqFile)) {
      content = fs.readFileSync(dqFile, 'utf8');
    }
    
    const lines = content.split('\n').filter(line => line.trim());
    if (!lines.includes(consultant)) {
      lines.push(consultant);
      fs.writeFileSync(dqFile, lines.join('\n') + '\n');
      log('info', 'Consultant disqualified', { consultant });
    }
    
    res.json({ success: true });
  } catch (error) {
    log('error', 'Error disqualifying consultant', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.delete('/disqualify', (req, res) => {
  try {
    const { consultant } = req.body;
    
    if (!consultant) {
      return res.status(400).json({ error: 'Consultant name is required' });
    }
    
    const dqFile = 'disqualified.txt';
    
    if (fs.existsSync(dqFile)) {
      const content = fs.readFileSync(dqFile, 'utf8');
      const lines = content.split('\n').filter(line => line.trim() && line !== consultant);
      fs.writeFileSync(dqFile, lines.join('\n') + (lines.length > 0 ? '\n' : ''));
      log('info', 'Consultant re-qualified', { consultant });
    }
    
    res.json({ success: true });
  } catch (error) {
    log('error', 'Error re-qualifying consultant', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.get('/disqualified', (req, res) => {
  try {
    const dqFile = 'disqualified.txt';
    const disqualified = [];
    
    if (fs.existsSync(dqFile)) {
      const content = fs.readFileSync(dqFile, 'utf8');
      content.split('\n').forEach(line => {
        if (line.trim()) disqualified.push(line.trim());
      });
    }
    
    res.json(disqualified);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/export', (req, res) => {
  try {
    const { format, tab, data } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided for export' });
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
    let filename, content;
    
    if (format === 'csv') {
      filename = `${tab || 'export'}-${timestamp}.csv`;
      content = generateCSV(data, tab);
      res.setHeader('Content-Type', 'text/csv');
    } else if (format === 'xlsx') {
      filename = `${tab || 'export'}-${timestamp}.xlsx`;
      content = generateXLSX(data, tab);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } else {
      return res.status(400).json({ error: 'Invalid format. Use csv or xlsx' });
    }
    
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(content);
    
  } catch (error) {
    log('error', 'Export failed', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

app.post('/export-all', (req, res) => {
  try {
    const { format, analysis } = req.body;
    
    if (!analysis) {
      return res.status(400).json({ error: 'No analysis data provided' });
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
    let filename, content;
    
    if (format === 'csv') {
      filename = `complete-analysis-${timestamp}.csv`;
      content = generateAllCSV(analysis);
      res.setHeader('Content-Type', 'text/csv');
    } else if (format === 'xlsx') {
      filename = `complete-analysis-${timestamp}.xlsx`;
      content = generateAllXLSX(analysis);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } else {
      return res.status(400).json({ error: 'Invalid format. Use csv or xlsx' });
    }
    
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(content);
    
  } catch (error) {
    log('error', 'Export all failed', { error: error.message });
    res.status(500).json({ error: error.message });
  }
});

function findHeaderRow(data) {
  // Look for row containing 'Resources Engaged' or similar headers
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const row = data[i];
    if (row && row.some(cell => 
      cell && typeof cell === 'string' && 
      (cell.toLowerCase().includes('resources') || cell.toLowerCase().includes('engaged'))
    )) {
      return i;
    }
  }
  return 13; // Default to row 14 (0-based index 13)
}

function loadFilterFiles(files, textInputs = {}) {
  const filters = {
    engineers: null,
    exclusions: [],
    solutionArchitects: null
  };
  
  try {
    // Engineers filter - combine cached file and text input
    let engineersList = [];
    
    // Load from uploaded file and cache it
    if (files.engineersFile && files.engineersFile[0]) {
      try {
        const uploadedFile = files.engineersFile[0];
        const persistentPath = 'engineers.txt';
        
        // Copy to persistent location
        fs.copyFileSync(uploadedFile.path, persistentPath);
        
        // Update cache
        const content = fs.readFileSync(persistentPath, 'utf8');
        engineersCache = content.split('\n')
          .map(line => line.trim().replace(/\s+/g, ' '))
          .filter(line => line && line.length > 1);
        engineersLastModified = new Date();
        
        engineersList = [...engineersCache];
        log('info', 'Engineers file cached', { count: engineersList.length });
      } catch (error) {
        log('error', 'Error caching engineers file', { error: error.message });
      }
    } else if (engineersCache) {
      // Use cached data
      engineersList = [...engineersCache];
    }
    
    // Add text input
    if (textInputs.engineersText && textInputs.engineersText.trim()) {
      const textEngineers = textInputs.engineersText.split('\n')
        .map(line => line.trim().replace(/\s+/g, ' '))
        .filter(line => line && line.length > 1);
      engineersList = [...engineersList, ...textEngineers];
    }
    
    if (engineersList.length > 0) {
      filters.engineers = [...new Set(engineersList)];
    }
    
    // Exclusions filter - only use text input (CSV file loaded separately)
    if (textInputs.exclusionsText && textInputs.exclusionsText.trim()) {
      const lines = textInputs.exclusionsText.split('\n').filter(line => line.trim());
      lines.forEach(line => {
        const parts = line.split(',');
        if (parts.length >= 2) {
          filters.exclusions.push({
            consultant: parts[0].trim(),
            project: parts[1].trim()
          });
        }
      });
    }
    
    // Solution Architects filter - combine cached file and text input
    let saList = [];
    
    // Load from uploaded file and cache it
    if (files.saFile && files.saFile[0]) {
      try {
        const uploadedFile = files.saFile[0];
        const persistentPath = 'solution_architects.txt';
        
        // Copy to persistent location
        fs.copyFileSync(uploadedFile.path, persistentPath);
        
        // Update cache
        const content = fs.readFileSync(persistentPath, 'utf8');
        saCache = content.split('\n')
          .map(line => line.trim())
          .filter(line => line);
        saLastModified = new Date();
        
        saList = [...saCache];
        log('info', 'SA file cached', { count: saList.length });
      } catch (error) {
        log('error', 'Error caching SA file', { error: error.message });
      }
    } else if (saCache) {
      // Use cached data
      saList = [...saCache];
    }
    
    // Add text input
    if (textInputs.saText && textInputs.saText.trim()) {
      const textSAs = textInputs.saText.split('\n')
        .map(line => line.trim())
        .filter(line => line);
      saList = [...saList, ...textSAs];
    }
    
    if (saList.length > 0) {
      filters.solutionArchitects = [...new Set(saList)];
    }
  } catch (error) {
    log('error', 'Error loading filter data', { error: error.message });
  }
  
  return filters;
}

function analyzeData(data, filters = {}, config = {}) {
  const successThreshold = config.thresholds?.success_threshold || 0.3;
  const consultants = {};
  const solutionArchitects = {};
  const customers = {};
  const allCustomers = {};
  const practiceProjects = new Set();
  
  // First pass: identify practice projects and process consultants/SAs
  data.forEach(row => {
    const budgetHrs = parseFloat(row['Budget Hrs']) || 0;
    const actualHrs = parseFloat(row['Total Hrs Posted']) || 0;
    const jobNumber = row['Job Number'] || '';
    let isPracticeProject = false;
    
    // Extract consultant data
    const resources = row['Resources Engaged'] || '';
    if (resources && budgetHrs > 0) {
      const consultantList = resources.split(',').map(name => {
        const normalized = name.trim().replace(/\s+/g, ' ');
        return normalized.length > 1 ? normalized : null;
      }).filter(name => name !== null);
      
      consultantList.forEach(consultant => {
        // Apply engineer filter if provided
        if (filters.engineers && filters.engineers.length > 0) {
          const normalizedConsultant = consultant.toLowerCase().trim();
          const found = filters.engineers.some(eng => {
            const normalizedEng = eng.toLowerCase().trim();
            return normalizedEng === normalizedConsultant || 
                   normalizedConsultant.includes(normalizedEng) ||
                   normalizedEng.includes(normalizedConsultant);
          });
          if (!found) {
            return;
          }
        }
        
        // Apply exclusions filter
        if (filters.exclusions && filters.exclusions.some(excl => 
          excl.consultant.toLowerCase().trim() === consultant.toLowerCase().trim() && 
          excl.project.trim() === jobNumber.trim()
        )) {
          return;
        }
        
        isPracticeProject = true;
        
        if (!consultants[consultant]) {
          consultants[consultant] = {
            projects: 0,
            totalHours: 0,
            withinBudget: 0,
            overBudget: 0,
            projectDetails: []
          };
        }
        
        consultants[consultant].projects++;
        consultants[consultant].totalHours += actualHrs;
        
        consultants[consultant].projectDetails.push({
          jobNumber: jobNumber,
          description: row['Job Description'] || '',
          customer: row['Customer'] || '',
          budgetHrs: budgetHrs,
          actualHrs: actualHrs,
          variance: ((actualHrs - budgetHrs) / budgetHrs * 100).toFixed(1),
          status: row['Project Status'] || '',
          completion: row['Project Complete %'] || 0
        });
        
        const variance = (actualHrs - budgetHrs) / budgetHrs;
        if (variance <= successThreshold) {
          consultants[consultant].withinBudget++;
        } else {
          consultants[consultant].overBudget++;
        }
      });
    }
    
    // Process Solution Architects
    const saField = row['Solution Architect'] || '';
    if (saField && saField.trim() && budgetHrs > 0) {
      const saList = saField.split(',').map(name => {
        const normalized = name.trim().replace(/\s+/g, ' ');
        return normalized.length > 1 ? normalized : null;
      }).filter(name => name !== null);
      
      saList.forEach(sa => {
        // Apply SA filter if provided
        if (filters.solutionArchitects && filters.solutionArchitects.length > 0) {
          const normalizedSA = sa.toLowerCase().trim();
          const found = filters.solutionArchitects.some(filterSA => {
            const normalizedFilterSA = filterSA.toLowerCase().trim();
            return normalizedFilterSA === normalizedSA || 
                   normalizedSA.includes(normalizedFilterSA) ||
                   normalizedFilterSA.includes(normalizedSA);
          });
          if (!found) {
            return;
          }
        }
        
        isPracticeProject = true;
        
        if (!solutionArchitects[sa]) {
          solutionArchitects[sa] = {
            projects: 0,
            totalBudgetedHours: 0,
            totalActualHours: 0,
            successfulProjects: 0,
            projectDetails: []
          };
        }
        
        solutionArchitects[sa].projects++;
        solutionArchitects[sa].totalBudgetedHours += budgetHrs;
        solutionArchitects[sa].totalActualHours += actualHrs;
        
        solutionArchitects[sa].projectDetails.push({
          jobNumber: jobNumber,
          description: row['Job Description'] || '',
          customer: row['Customer'] || '',
          budgetHrs: budgetHrs,
          actualHrs: actualHrs,
          variance: ((actualHrs - budgetHrs) / budgetHrs * 100).toFixed(1),
          status: row['Project Status'] || '',
          completion: row['Project Complete %'] || 0
        });
        
        const variance = (actualHrs - budgetHrs) / budgetHrs;
        if (variance <= successThreshold) {
          solutionArchitects[sa].successfulProjects++;
        }
      });
    }
    
    // Track practice projects
    if (isPracticeProject) {
      practiceProjects.add(jobNumber);
    }
  });
  
  // Second pass: process customer data
  data.forEach(row => {
    const budgetHrs = parseFloat(row['Budget Hrs']) || 0;
    const actualHrs = parseFloat(row['Total Hrs Posted']) || 0;
    const jobNumber = row['Job Number'] || '';
    const customerName = row['Customer'] || 'Unknown';
    
    if (customerName && customerName.trim() && budgetHrs > 0) {
      const variance = (actualHrs - budgetHrs) / budgetHrs;
      
      // Process ALL company projects
      if (!allCustomers[customerName]) {
        allCustomers[customerName] = {
          projects: 0,
          totalBudgetHrs: 0,
          totalActualHrs: 0,
          totalVariance: 0,
          withinBudget: 0,
          overBudget: 0,
          projectDetails: []
        };
      }
      
      allCustomers[customerName].projects++;
      allCustomers[customerName].totalBudgetHrs += budgetHrs;
      allCustomers[customerName].totalActualHrs += actualHrs;
      allCustomers[customerName].totalVariance += variance * 100;
      
      if (variance <= successThreshold) {
        allCustomers[customerName].withinBudget++;
      } else {
        allCustomers[customerName].overBudget++;
      }
      
      allCustomers[customerName].projectDetails.push({
        jobNumber: jobNumber,
        description: row['Job Description'] || '',
        budgetHrs: budgetHrs,
        actualHrs: actualHrs,
        variance: (variance * 100).toFixed(1),
        status: row['Project Status'] || '',
        completion: row['Project Complete %'] || 0,
        resources: row['Resources Engaged'] || '',
        solutionArchitect: row['Solution Architect'] || ''
      });
      
      // Process PRACTICE projects only
      if (practiceProjects.has(jobNumber)) {
        if (!customers[customerName]) {
          customers[customerName] = {
            projects: 0,
            totalBudgetHrs: 0,
            totalActualHrs: 0,
            totalVariance: 0,
            withinBudget: 0,
            overBudget: 0,
            projectDetails: []
          };
        }
        
        customers[customerName].projects++;
        customers[customerName].totalBudgetHrs += budgetHrs;
        customers[customerName].totalActualHrs += actualHrs;
        customers[customerName].totalVariance += variance * 100;
        
        if (variance <= successThreshold) {
          customers[customerName].withinBudget++;
        } else {
          customers[customerName].overBudget++;
        }
        
        customers[customerName].projectDetails.push({
          jobNumber: jobNumber,
          description: row['Job Description'] || '',
          budgetHrs: budgetHrs,
          actualHrs: actualHrs,
          variance: (variance * 100).toFixed(1),
          status: row['Project Status'] || '',
          completion: row['Project Complete %'] || 0,
          resources: row['Resources Engaged'] || '',
          solutionArchitect: row['Solution Architect'] || ''
        });
      }
    }
  });

  // Calculate metrics
  const consultantMetrics = Object.keys(consultants).map(name => {
    const consultant = consultants[name];
    const successRate = (consultant.withinBudget / consultant.projects) * 100;
    const efficiencyScore = (consultant.withinBudget / consultant.projects) * 100;
    
    return {
      name,
      projects: consultant.projects,
      hours: consultant.totalHours,
      successRate: successRate.toFixed(1),
      efficiencyScore: efficiencyScore.toFixed(1),
      withinBudget: consultant.withinBudget,
      overBudget: consultant.overBudget,
      projectDetails: consultant.projectDetails || []
    };
  });

  // Calculate SA metrics
  const saMetrics = Object.keys(solutionArchitects).map(name => {
    const sa = solutionArchitects[name];
    const successRate = (sa.successfulProjects / sa.projects) * 100;
    const variance = ((sa.totalActualHours - sa.totalBudgetedHours) / sa.totalBudgetedHours) * 100;
    
    return {
      name,
      projects: sa.projects,
      budgetedHours: sa.totalBudgetedHours,
      actualHours: sa.totalActualHours,
      successRate: successRate.toFixed(1),
      variance: variance.toFixed(1),
      projectDetails: sa.projectDetails || []
    };
  });

  // Calculate customer metrics for practice (filtered) projects
  const practiceCustomerMetrics = Object.keys(customers).map(name => {
    const customer = customers[name];
    const successRate = (customer.withinBudget / customer.projects) * 100;
    const avgVariance = customer.totalVariance / customer.projects;
    
    return {
      name,
      projects: customer.projects,
      successRate: successRate.toFixed(1),
      avgVariance: avgVariance.toFixed(1),
      totalBudgetHrs: customer.totalBudgetHrs,
      totalActualHrs: customer.totalActualHrs,
      withinBudget: customer.withinBudget,
      overBudget: customer.overBudget,
      projectDetails: customer.projectDetails || []
    };
  }).sort((a, b) => b.successRate - a.successRate);

  // Calculate customer metrics for all company projects
  const allCustomerMetrics = Object.keys(allCustomers).map(name => {
    const customer = allCustomers[name];
    const successRate = (customer.withinBudget / customer.projects) * 100;
    const avgVariance = customer.totalVariance / customer.projects;
    
    return {
      name,
      projects: customer.projects,
      successRate: successRate.toFixed(1),
      avgVariance: avgVariance.toFixed(1),
      totalBudgetHrs: customer.totalBudgetHrs,
      totalActualHrs: customer.totalActualHrs,
      withinBudget: customer.withinBudget,
      overBudget: customer.overBudget,
      projectDetails: customer.projectDetails || []
    };
  }).filter(c => c.projects >= 3).sort((a, b) => b.successRate - a.successRate);

  // Load disqualified consultants
  const disqualified = loadDisqualifiedConsultants();
  
  // Calculate Consultant of the Quarter (composite scoring)
  const consultantOfQuarter = Object.keys(consultants)
    .filter(name => !disqualified.includes(name)) // Exclude disqualified consultants
    .map(name => {
      const consultant = consultants[name];
      const successRate = (consultant.withinBudget / consultant.projects) * 100;
      const efficiencyScore = successRate; // Same calculation as success rate
      const projectCount = consultant.projects;
      const totalHours = consultant.totalHours;
      
      // Composite scoring (weighted)
      const successWeight = 0.4;
      const efficiencyWeight = 0.3;
      const volumeWeight = 0.2;
      const consistencyWeight = 0.1;
      
      // Normalize volume score (projects + hours)
      const maxProjects = Math.max(...Object.values(consultants).map(c => c.projects));
      const maxHours = Math.max(...Object.values(consultants).map(c => c.totalHours));
      const volumeScore = ((projectCount / maxProjects) * 50) + ((totalHours / maxHours) * 50);
      
      // Consistency score (inverse of variance in project performance)
      const projectVariances = consultant.projectDetails.map(p => Math.abs(parseFloat(p.variance)));
      const avgVariance = projectVariances.reduce((sum, v) => sum + v, 0) / projectVariances.length;
      const consistencyScore = Math.max(0, 100 - avgVariance);
      
      const compositeScore = (
        (successRate * successWeight) +
        (efficiencyScore * efficiencyWeight) +
        (volumeScore * volumeWeight) +
        (consistencyScore * consistencyWeight)
      );
      
      return {
        name,
        compositeScore: compositeScore.toFixed(1),
        successRate: successRate.toFixed(1),
        efficiencyScore: efficiencyScore.toFixed(1),
        volumeScore: volumeScore.toFixed(1),
        consistencyScore: consistencyScore.toFixed(1),
        projects: projectCount,
        hours: totalHours
      };
    }).sort((a, b) => b.compositeScore - a.compositeScore);

  // Calculate DAS+ Analysis
  const dasAnalysis = Object.keys(consultants).map(name => {
    const consultant = consultants[name];
    const dasScores = consultant.projectDetails.map(project => {
      const budgetRatio = project.actualHrs / project.budgetHrs;
      let completionRatio = project.completion || 0;
      
      // Completion normalization
      if (project.status && (project.status.toLowerCase().includes('closed') || project.status.toLowerCase().includes('complete'))) {
        completionRatio = 1.0;
      } else if (completionRatio >= 0.95) {
        completionRatio = 1.0;
      }
      
      const dasScore = 1 - Math.abs(budgetRatio - completionRatio);
      return Math.max(0, Math.min(1, dasScore));
    });
    
    const avgDAS = dasScores.reduce((sum, score) => sum + score, 0) / dasScores.length;
    const sortedScores = [...dasScores].sort((a, b) => a - b);
    const medianDAS = sortedScores[Math.floor(sortedScores.length / 2)];
    const lowCount = dasScores.filter(score => score < 0.75).length;
    const highCount = dasScores.filter(score => score >= 0.85).length;
    const reviewCount = dasScores.filter(score => score >= 0.3 && score <= 0.9).length;
    
    return {
      name,
      projects: consultant.projects,
      avgDAS: avgDAS.toFixed(3),
      medianDAS: medianDAS.toFixed(3),
      lowCount,
      highCount,
      reviewCount
    };
  }).sort((a, b) => b.avgDAS - a.avgDAS);

  // Calculate SA-Consultant Combinations (filtered SAs and Engineers only)
  const saCombinations = {};
  data.forEach(row => {
    const budgetHrs = parseFloat(row['Budget Hrs']) || 0;
    const actualHrs = parseFloat(row['Total Hrs Posted']) || 0;
    const resources = row['Resources Engaged'] || '';
    const saField = row['Solution Architect'] || '';
    
    if (budgetHrs > 0 && resources && saField) {
      const consultantList = resources.split(',').map(name => name.trim().replace(/\s+/g, ' ')).filter(name => name.length > 1);
      const saList = saField.split(',').map(name => name.trim().replace(/\s+/g, ' ')).filter(name => name.length > 1);
      
      consultantList.forEach(consultant => {
        // Apply engineer filter - only include consultants from engineers list
        if (filters.engineers && filters.engineers.length > 0) {
          const normalizedConsultant = consultant.toLowerCase().trim();
          const found = filters.engineers.some(eng => {
            const normalizedEng = eng.toLowerCase().trim();
            return normalizedEng === normalizedConsultant || 
                   normalizedConsultant.includes(normalizedEng) ||
                   normalizedEng.includes(normalizedConsultant);
          });
          if (!found) return;
        }
        
        saList.forEach(sa => {
          // Apply SA filter
          if (filters.solutionArchitects && filters.solutionArchitects.length > 0) {
            const normalizedSA = sa.toLowerCase().trim();
            const found = filters.solutionArchitects.some(filterSA => {
              const normalizedFilterSA = filterSA.toLowerCase().trim();
              return normalizedFilterSA === normalizedSA || 
                     normalizedSA.includes(normalizedFilterSA) ||
                     normalizedFilterSA.includes(normalizedSA);
            });
            if (!found) return;
          }
          
          const key = `${sa}|${consultant}`;
          if (!saCombinations[key]) {
            saCombinations[key] = { sa, consultant, projects: 0, successful: 0 };
          }
          saCombinations[key].projects++;
          const variance = (actualHrs - budgetHrs) / budgetHrs;
          if (variance <= successThreshold) {
            saCombinations[key].successful++;
          }
        });
      });
    }
  });
  
  const combinationMetrics = Object.values(saCombinations)
    .filter(combo => combo.projects >= 2)
    .map(combo => ({
      sa: combo.sa,
      consultant: combo.consultant,
      projects: combo.projects,
      successRate: ((combo.successful / combo.projects) * 100).toFixed(1)
    }))
    .sort((a, b) => b.successRate - a.successRate);

  // Calculate SA-Customer Success Rates (filtered SAs only)
  const saCustomerCombos = {};
  data.forEach(row => {
    const budgetHrs = parseFloat(row['Budget Hrs']) || 0;
    const actualHrs = parseFloat(row['Total Hrs Posted']) || 0;
    const customerName = row['Customer'] || '';
    const saField = row['Solution Architect'] || '';
    
    if (budgetHrs > 0 && customerName && saField) {
      const saList = saField.split(',').map(name => name.trim().replace(/\s+/g, ' ')).filter(name => name.length > 1);
      
      saList.forEach(sa => {
        // Apply SA filter
        if (filters.solutionArchitects && filters.solutionArchitects.length > 0) {
          const normalizedSA = sa.toLowerCase().trim();
          const found = filters.solutionArchitects.some(filterSA => {
            const normalizedFilterSA = filterSA.toLowerCase().trim();
            return normalizedFilterSA === normalizedSA || 
                   normalizedSA.includes(normalizedFilterSA) ||
                   normalizedFilterSA.includes(normalizedSA);
          });
          if (!found) return;
        }
        
        const key = `${sa}|${customerName}`;
        if (!saCustomerCombos[key]) {
          saCustomerCombos[key] = { sa, customer: customerName, projects: 0, successful: 0 };
        }
        saCustomerCombos[key].projects++;
        const variance = (actualHrs - budgetHrs) / budgetHrs;
        if (variance <= successThreshold) {
          saCustomerCombos[key].successful++;
        }
      });
    }
  });
  
  const saCustomerMetrics = Object.values(saCustomerCombos)
    .filter(combo => combo.projects >= 2)
    .map(combo => ({
      sa: combo.sa,
      customer: combo.customer,
      projects: combo.projects,
      successRate: ((combo.successful / combo.projects) * 100).toFixed(1)
    }))
    .sort((a, b) => b.successRate - a.successRate);

  // Add Risk Scoring to Customer Metrics
  const addRiskScoring = (customerList) => {
    return customerList.map(customer => {
      let riskScore = 0;
      
      // Success rate risk (40% weight)
      const successRate = parseFloat(customer.successRate);
      if (successRate < 50) riskScore += 4;
      else if (successRate < 70) riskScore += 2;
      else if (successRate < 85) riskScore += 1;
      
      // Project volume risk (20% weight)
      if (customer.projects >= 10) riskScore += 0;
      else if (customer.projects >= 5) riskScore += 1;
      else riskScore += 2;
      
      // Variance consistency risk (25% weight)
      const avgVariance = Math.abs(parseFloat(customer.avgVariance));
      if (avgVariance > 50) riskScore += 3;
      else if (avgVariance > 30) riskScore += 2;
      else if (avgVariance > 15) riskScore += 1;
      
      // Budget size risk (15% weight)
      const totalBudget = parseFloat(customer.totalBudgetHrs);
      if (totalBudget > 5000) riskScore += 1;
      else if (totalBudget > 2000) riskScore += 0;
      else riskScore += 1;
      
      let riskLevel, riskColor;
      if (riskScore <= 2) { riskLevel = 'Low'; riskColor = 'success'; }
      else if (riskScore <= 5) { riskLevel = 'Medium'; riskColor = 'warning'; }
      else if (riskScore <= 8) { riskLevel = 'High'; riskColor = 'danger'; }
      else { riskLevel = 'Critical'; riskColor = 'dark'; }
      
      return { ...customer, riskScore, riskLevel, riskColor };
    });
  };

  return {
    totalProjects: data.length,
    consultants: consultantMetrics.sort((a, b) => b.efficiencyScore - a.efficiencyScore),
    solutionArchitects: saMetrics.sort((a, b) => b.successRate - a.successRate),
    customers: {
      practice: addRiskScoring(practiceCustomerMetrics),
      company: addRiskScoring(allCustomerMetrics)
    },
    dasAnalysis: dasAnalysis,
    saCombinations: combinationMetrics,
    saCustomerAnalysis: saCustomerMetrics,
    consultantOfQuarter: consultantOfQuarter
  };
}

function getDefaultConfig() {
  return {
    thresholds: {
      efficiency_threshold: 0.15,
      success_threshold: 0.3,
      green_threshold: -0.1,
      yellow_threshold: 0.1,
      red_threshold: 0.3
    },
    trending: {
      enable_trending: false
    },
    project_filtering: {
      enable_date_filter: false,
      filter_type: "date",
      days_from_today: 365,
      specific_date: getCurrentQuarterStart(),
      exclude_closed_before_date: true
    }
  };
}

let configCache = null;
let configLastModified = null;
let exclusionsCache = null;
let exclusionsLastModified = null;
let engineersCache = null;
let engineersLastModified = null;
let saCache = null;
let saLastModified = null;

function loadConfig() {
  try {
    if (fs.existsSync('config.json')) {
      const stats = fs.statSync('config.json');
      if (!configCache || stats.mtime > configLastModified) {
        configCache = JSON.parse(fs.readFileSync('config.json', 'utf8'));
        configLastModified = stats.mtime;
      }
      return configCache;
    }
  } catch (error) {
    console.error('Error loading config:', error);
  }
  return getDefaultConfig();
}

function applyDateFilter(data, filterConfig) {
  if (!filterConfig.enable_date_filter) return data;
  
  let cutoffDate;
  if (filterConfig.filter_type === 'days') {
    cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - (filterConfig.days_from_today || 365));
  } else {
    cutoffDate = new Date(filterConfig.specific_date || getCurrentQuarterStart());
  }
  
  return data.filter(row => {
    const status = row['Project Status'] || '';
    
    // Always include open projects
    if (status.toLowerCase().includes('open') || status.toLowerCase().includes('active')) {
      return true;
    }
    
    // For closed projects, check end date
    const endDate = row['End Date'] || row['Project End Date'] || row['Completion Date'] || row['Close Date'];
    if (endDate) {
      const rowDate = new Date(endDate);
      return rowDate >= cutoffDate;
    }
    
    return true; // Include if no end date found
  });
}

function loadExistingExclusions() {
  const exclusionFile = 'exclusions.csv';
  
  if (fs.existsSync(exclusionFile)) {
    try {
      const stats = fs.statSync(exclusionFile);
      if (!exclusionsCache || stats.mtime > exclusionsLastModified) {
        const exclusions = [];
        const content = fs.readFileSync(exclusionFile, 'utf8');
        const lines = content.split('\n').filter(line => line.trim());
        
        // Skip header if it exists
        const startIndex = lines[0] && lines[0].toLowerCase().includes('consultant') ? 1 : 0;
        
        for (let i = startIndex; i < lines.length; i++) {
          const line = lines[i].trim();
          if (line) {
            // Parse CSV line properly handling quoted fields
            const parts = [];
            let current = '';
            let inQuotes = false;
            
            for (let j = 0; j < line.length; j++) {
              const char = line[j];
              if (char === '"') {
                if (inQuotes && line[j + 1] === '"') {
                  current += '"';
                  j++; // Skip next quote
                } else {
                  inQuotes = !inQuotes;
                }
              } else if (char === ',' && !inQuotes) {
                parts.push(current.trim());
                current = '';
              } else {
                current += char;
              }
            }
            parts.push(current.trim());
            
            if (parts.length >= 2 && parts[0].trim() && parts[1].trim()) {
              exclusions.push({
                consultant: parts[0].trim(),
                project: parts[1].trim(),
                reason: parts[2] ? parts[2].trim() : 'No reason provided'
              });
            }
          }
        }
        
        exclusionsCache = exclusions;
        exclusionsLastModified = stats.mtime;
        log('info', 'Loaded exclusions to cache', { count: exclusions.length });
      }
    } catch (error) {
      log('error', 'Error loading existing exclusions', { error: error.message });
      return [];
    }
  } else {
    exclusionsCache = [];
    exclusionsLastModified = null;
  }
  
  return exclusionsCache || [];
}

function loadDisqualifiedConsultants() {
  const dqFile = 'disqualified.txt';
  const disqualified = [];
  
  if (fs.existsSync(dqFile)) {
    try {
      const content = fs.readFileSync(dqFile, 'utf8');
      content.split('\n').forEach(line => {
        if (line.trim()) disqualified.push(line.trim());
      });
    } catch (error) {
      log('error', 'Error loading disqualified consultants', { error: error.message });
    }
  }
  
  return disqualified;
}

function getCurrentQuarterStart() {
  const now = new Date();
  const quarter = Math.floor((now.getMonth()) / 3);
  const quarterStart = new Date(now.getFullYear(), quarter * 3, 1);
  return quarterStart.toISOString().split('T')[0];
}

function generateCSV(data, tab) {
  let csv = '';
  
  switch (tab) {
    case 'consultants':
      csv = 'Consultant,Projects,Hours,Efficiency %,Success Rate %,Within Budget,Over Budget\n';
      data.forEach(row => {
        csv += `"${row.name}",${row.projects},${row.hours},${row.efficiencyScore},${row.successRate},${row.withinBudget},${row.overBudget}\n`;
      });
      break;
      
    case 'solutionArchitects':
      csv = 'Solution Architect,Projects,Budgeted Hours,Actual Hours,Success Rate %,Variance %\n';
      data.forEach(row => {
        csv += `"${row.name}",${row.projects},${row.budgetedHours},${row.actualHours},${row.successRate},${row.variance}\n`;
      });
      break;
      
    case 'customers':
      csv = 'Customer,Projects,Success Rate %,Avg Variance %,Total Budget Hrs,Risk Level\n';
      data.forEach(row => {
        csv += `"${row.name}",${row.projects},${row.successRate},${row.avgVariance},${row.totalBudgetHrs},${row.riskLevel}\n`;
      });
      break;
      
    case 'das':
      csv = 'Consultant,Projects,Avg DAS+,Median DAS+,High Performing,Needs Review,Low Performing\n';
      data.forEach(row => {
        csv += `"${row.name}",${row.projects},${row.avgDAS},${row.medianDAS},${row.highCount},${row.reviewCount},${row.lowCount}\n`;
      });
      break;
      
    case 'combinations':
      csv = 'Solution Architect,Consultant,Projects,Success Rate %\n';
      data.forEach(row => {
        csv += `"${row.sa}","${row.consultant}",${row.projects},${row.successRate}\n`;
      });
      break;
      
    case 'consultant-projects':
      csv = 'Job Number,Description,Customer,Budget Hours,Actual Hours,Variance %,Complete %,Status\n';
      data.forEach(row => {
        csv += `"${row['Job Number']}","${row.Description}","${row.Customer}",${row['Budget Hours']},${row['Actual Hours']},${row['Variance %']},${row['Complete %']},"${row.Status}"\n`;
      });
      break;
      
    default:
      csv = 'No data available\n';
  }
  
  return csv;
}

function generateXLSX(data, tab) {
  let worksheet;
  
  switch (tab) {
    case 'consultants':
      worksheet = XLSX.utils.json_to_sheet(data.map(row => ({
        'Consultant': row.name,
        'Projects': row.projects,
        'Hours': row.hours,
        'Efficiency %': row.efficiencyScore,
        'Success Rate %': row.successRate,
        'Within Budget': row.withinBudget,
        'Over Budget': row.overBudget
      })));
      break;
      
    case 'solutionArchitects':
      worksheet = XLSX.utils.json_to_sheet(data.map(row => ({
        'Solution Architect': row.name,
        'Projects': row.projects,
        'Budgeted Hours': row.budgetedHours,
        'Actual Hours': row.actualHours,
        'Success Rate %': row.successRate,
        'Variance %': row.variance
      })));
      break;
      
    case 'customers':
      worksheet = XLSX.utils.json_to_sheet(data.map(row => ({
        'Customer': row.name,
        'Projects': row.projects,
        'Success Rate %': row.successRate,
        'Avg Variance %': row.avgVariance,
        'Total Budget Hrs': row.totalBudgetHrs,
        'Risk Level': row.riskLevel
      })));
      break;
      
    case 'das':
      worksheet = XLSX.utils.json_to_sheet(data.map(row => ({
        'Consultant': row.name,
        'Projects': row.projects,
        'Avg DAS+': row.avgDAS,
        'Median DAS+': row.medianDAS,
        'High Performing': row.highCount,
        'Needs Review': row.reviewCount,
        'Low Performing': row.lowCount
      })));
      break;
      
    case 'combinations':
      worksheet = XLSX.utils.json_to_sheet(data.map(row => ({
        'Solution Architect': row.sa,
        'Consultant': row.consultant,
        'Projects': row.projects,
        'Success Rate %': row.successRate
      })));
      break;
      
    case 'consultant-projects':
      // Create worksheet without headers, then add data
      const wsData = data.map(row => Object.values(row));
      worksheet = XLSX.utils.aoa_to_sheet(wsData);
      break;
      
    default:
      worksheet = XLSX.utils.json_to_sheet([{ 'Error': 'No data available' }]);
  }
  
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, tab || 'Export');
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

function generateAllCSV(analysis) {
  let csv = '';
  
  // Consultants section
  csv += '=== CONSULTANT ANALYSIS ===\n';
  csv += 'Consultant,Projects,Hours,Efficiency %,Success Rate %,Within Budget,Over Budget\n';
  analysis.consultants.forEach(row => {
    csv += `"${row.name}",${row.projects},${row.hours},${row.efficiencyScore},${row.successRate},${row.withinBudget},${row.overBudget}\n`;
  });
  
  csv += '\n=== SOLUTION ARCHITECT ANALYSIS ===\n';
  csv += 'Solution Architect,Projects,Budgeted Hours,Actual Hours,Success Rate %,Variance %\n';
  if (analysis.solutionArchitects) {
    analysis.solutionArchitects.forEach(row => {
      csv += `"${row.name}",${row.projects},${row.budgetedHours},${row.actualHours},${row.successRate},${row.variance}\n`;
    });
  }
  
  csv += '\n=== PRACTICE CUSTOMER ANALYSIS ===\n';
  csv += 'Customer,Projects,Success Rate %,Avg Variance %,Total Budget Hrs,Risk Level\n';
  if (analysis.customers && analysis.customers.practice) {
    analysis.customers.practice.forEach(row => {
      csv += `"${row.name}",${row.projects},${row.successRate},${row.avgVariance},${row.totalBudgetHrs},${row.riskLevel}\n`;
    });
  }
  
  // Skip company customers to reduce size
  
  csv += '\n=== DAS+ ANALYSIS ===\n';
  csv += 'Consultant,Projects,Avg DAS+,Median DAS+,High Performing,Needs Review,Low Performing\n';
  if (analysis.dasAnalysis) {
    analysis.dasAnalysis.forEach(row => {
      csv += `"${row.name}",${row.projects},${row.avgDAS},${row.medianDAS},${row.highCount},${row.reviewCount},${row.lowCount}\n`;
    });
  }
  
  csv += '\n=== SA-CONSULTANT COMBINATIONS ===\n';
  csv += 'Solution Architect,Consultant,Projects,Success Rate %\n';
  if (analysis.saCombinations) {
    analysis.saCombinations.forEach(row => {
      csv += `"${row.sa}","${row.consultant}",${row.projects},${row.successRate}\n`;
    });
  }
  
  // Skip SA-Customer analysis to reduce size
  
  return csv;
}

function generateAllXLSX(analysis) {
  const workbook = XLSX.utils.book_new();
  
  // Consultants sheet (summary only)
  const consultantsWS = XLSX.utils.json_to_sheet(analysis.consultants.map(row => ({
    'Consultant': row.name,
    'Projects': row.projects,
    'Hours': row.hours,
    'Efficiency %': row.efficiencyScore,
    'Success Rate %': row.successRate,
    'Within Budget': row.withinBudget,
    'Over Budget': row.overBudget
  })));
  XLSX.utils.book_append_sheet(workbook, consultantsWS, 'Consultants');
  
  // Solution Architects sheet (summary only)
  if (analysis.solutionArchitects) {
    const saWS = XLSX.utils.json_to_sheet(analysis.solutionArchitects.map(row => ({
      'Solution Architect': row.name,
      'Projects': row.projects,
      'Budgeted Hours': row.budgetedHours,
      'Actual Hours': row.actualHours,
      'Success Rate %': row.successRate,
      'Variance %': row.variance
    })));
    XLSX.utils.book_append_sheet(workbook, saWS, 'Solution Architects');
  }
  
  // Practice Customers sheet (summary only)
  if (analysis.customers && analysis.customers.practice) {
    const practiceWS = XLSX.utils.json_to_sheet(analysis.customers.practice.map(row => ({
      'Customer': row.name,
      'Projects': row.projects,
      'Success Rate %': row.successRate,
      'Avg Variance %': row.avgVariance,
      'Total Budget Hrs': row.totalBudgetHrs,
      'Risk Level': row.riskLevel
    })));
    XLSX.utils.book_append_sheet(workbook, practiceWS, 'Practice Customers');
  }
  
  // DAS+ Analysis sheet
  if (analysis.dasAnalysis) {
    const dasWS = XLSX.utils.json_to_sheet(analysis.dasAnalysis.map(row => ({
      'Consultant': row.name,
      'Projects': row.projects,
      'Avg DAS+': row.avgDAS,
      'Median DAS+': row.medianDAS,
      'High Performing': row.highCount,
      'Needs Review': row.reviewCount,
      'Low Performing': row.lowCount
    })));
    XLSX.utils.book_append_sheet(workbook, dasWS, 'DAS+ Analysis');
  }
  
  // Company Customers sheet
  if (analysis.customers && analysis.customers.company) {
    const companyWS = XLSX.utils.json_to_sheet(analysis.customers.company.map(row => ({
      'Customer': row.name,
      'Projects': row.projects,
      'Success Rate %': row.successRate,
      'Avg Variance %': row.avgVariance,
      'Total Budget Hrs': row.totalBudgetHrs,
      'Risk Level': row.riskLevel
    })));
    XLSX.utils.book_append_sheet(workbook, companyWS, 'Company Customers');
  }
  
  // SA-Consultant Combinations sheet
  if (analysis.saCombinations) {
    const combosWS = XLSX.utils.json_to_sheet(analysis.saCombinations.map(row => ({
      'Solution Architect': row.sa,
      'Consultant': row.consultant,
      'Projects': row.projects,
      'Success Rate %': row.successRate
    })));
    XLSX.utils.book_append_sheet(workbook, combosWS, 'SA-Consultant Combos');
  }
  
  // SA-Customer Analysis sheet
  if (analysis.saCustomerAnalysis) {
    const saCustomerWS = XLSX.utils.json_to_sheet(analysis.saCustomerAnalysis.map(row => ({
      'Solution Architect': row.sa,
      'Customer': row.customer,
      'Projects': row.projects,
      'Success Rate %': row.successRate
    })));
    XLSX.utils.book_append_sheet(workbook, saCustomerWS, 'SA-Customer Analysis');
  }
  
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

app.listen(PORT, () => {
  log('info', `Server started on port ${PORT}`);
  console.log(`Server running on http://localhost:${PORT}`);
});