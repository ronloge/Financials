const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ storage: storage });

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/upload', upload.fields([
  { name: 'excelFile', maxCount: 1 },
  { name: 'engineersFile', maxCount: 1 },
  { name: 'excludeFile', maxCount: 1 },
  { name: 'saFile', maxCount: 1 }
]), (req, res) => {
  try {
    if (!req.files || !req.files.excelFile) {
      return res.status(400).json({ error: 'No Excel file uploaded' });
    }

    const workbook = XLSX.readFile(req.files.excelFile[0].path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Skip header rows (typically row 14 is the header)
    const headerRowIndex = findHeaderRow(data);
    const headers = data[headerRowIndex];
    const jsonData = data.slice(headerRowIndex + 1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    // Load filter files and config
    const filters = loadFilterFiles(req.files);
    const config = loadConfig();
    
    // Analysis with filters and config
    const analysis = analyzeData(jsonData, filters, config);
    
    res.json({
      success: true,
      filename: req.files.excelFile[0].filename,
      analysis: analysis
    });
  } catch (error) {
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

function loadFilterFiles(files) {
  const filters = {
    engineers: null,
    exclusions: [],
    solutionArchitects: null
  };
  
  try {
    if (files.engineersFile) {
      const engineersContent = fs.readFileSync(files.engineersFile[0].path, 'utf8');
      filters.engineers = engineersContent.split('\n')
        .map(line => line.trim().replace(/\s+/g, ' '))
        .filter(line => line && line.length > 1);
    }
    
    if (files.excludeFile) {
      const excludeContent = fs.readFileSync(files.excludeFile[0].path, 'utf8');
      const lines = excludeContent.split('\n');
      lines.slice(1).forEach(line => { // Skip header
        const parts = line.split(',');
        if (parts.length >= 2) {
          filters.exclusions.push({
            consultant: parts[0].trim(),
            project: parts[1].trim()
          });
        }
      });
    }
    
    if (files.saFile) {
      const saContent = fs.readFileSync(files.saFile[0].path, 'utf8');
      filters.solutionArchitects = saContent.split('\n').map(line => line.trim()).filter(line => line);
    }
  } catch (error) {
    console.error('Error loading filter files:', error);
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
          excl.consultant.toLowerCase() === consultant.toLowerCase() && 
          excl.project === jobNumber
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
  }).filter(c => c.projects >= 3).sort((a, b) => b.successRate - a.successRate);

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

  // Calculate Consultant of the Quarter (composite scoring)
  const consultantOfQuarter = Object.keys(consultants).map(name => {
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

  // Calculate SA-Consultant Combinations (filtered SAs only)
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
    project_filtering: {
      enable_date_filter: false,
      filter_type: "date",
      days_from_today: 365,
      specific_date: getCurrentQuarterStart(),
      exclude_closed_before_date: true
    }
  };
}

function loadConfig() {
  try {
    if (fs.existsSync('config.json')) {
      return JSON.parse(fs.readFileSync('config.json', 'utf8'));
    }
  } catch (error) {
    console.error('Error loading config:', error);
  }
  return getDefaultConfig();
}

function getCurrentQuarterStart() {
  const now = new Date();
  const quarter = Math.floor((now.getMonth()) / 3);
  const quarterStart = new Date(now.getFullYear(), quarter * 3, 1);
  return quarterStart.toISOString().split('T')[0];
}

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});