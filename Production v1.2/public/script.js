let currentAnalysis = null;
let performanceChart = null;
let autoRefreshInterval = null;

// Simple HTML sanitization
function sanitizeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
    loadExistingConfig();
    updateDateRangeDisplay();
    showMessage('Ready to analyze', 'success');
    
    // Configuration panel toggle
    document.getElementById('configPanel').addEventListener('shown.bs.collapse', () => {
        document.getElementById('configToggle').className = 'fas fa-chevron-up float-end';
    });
    document.getElementById('configPanel').addEventListener('hidden.bs.collapse', () => {
        document.getElementById('configToggle').className = 'fas fa-chevron-down float-end';
    });
});

function setupEventListeners() {
    try {
        // Data source selection
        document.querySelectorAll('input[name="dataSource"]').forEach(radio => {
            radio.addEventListener('change', toggleDataSource);
        });
        
        // File uploads with null checks
        const excelFile = document.getElementById('excelFile');
        const configFile = document.getElementById('configFile');
        if (excelFile) excelFile.addEventListener('change', handleFileUpload);
        if (configFile) configFile.addEventListener('change', handleConfigUpload);
        
        // Buttons with null checks
        const runAnalysisBtn = document.getElementById('runAnalysis');
        const viewExclusionsBtn = document.getElementById('viewExclusions');
        const connectSQLBtn = document.getElementById('connectSQL');
        const exportResultsBtn = document.getElementById('exportResults');
        const applyConfigBtn = document.getElementById('applyConfig');
        const refreshConfigBtn = document.getElementById('refreshConfig');
        if (runAnalysisBtn) runAnalysisBtn.addEventListener('click', runAnalysis);
        if (viewExclusionsBtn) viewExclusionsBtn.addEventListener('click', viewExclusions);
        if (connectSQLBtn) connectSQLBtn.addEventListener('click', connectToSQL);
        if (exportResultsBtn) exportResultsBtn.addEventListener('click', exportToCSV);
        if (applyConfigBtn) applyConfigBtn.addEventListener('click', handleConfigUpload);
        if (refreshConfigBtn) refreshConfigBtn.addEventListener('click', loadExistingConfig);
        
        // Configuration changes with individual update functions
        const thresholdElements = {
            'efficiencyThreshold': 'efficiency_threshold',
            'successThreshold': 'success_threshold', 
            'greenThreshold': 'green_threshold',
            'yellowThreshold': 'yellow_threshold',
            'redThreshold': 'red_threshold'
        };
        
        Object.entries(thresholdElements).forEach(([id, configKey]) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('input', () => updateSingleThreshold(configKey, parseFloat(el.value)));
        });
        
        const enableDateFilter = document.getElementById('enableDateFilter');
        const enableTrending = document.getElementById('enableTrending');
        const enableAutoRefresh = document.getElementById('enableAutoRefresh');
        if (enableDateFilter) enableDateFilter.addEventListener('change', () => {
            toggleDateFilter();
            updateDateFilterConfig();
        });
        if (enableTrending) enableTrending.addEventListener('change', () => {
            toggleTrending();
            updateTrendingConfig();
        });
        if (enableAutoRefresh) enableAutoRefresh.addEventListener('change', toggleAutoRefresh);
        
        // Date filter options
        document.querySelectorAll('input[name="filterType"]').forEach(radio => {
            radio.addEventListener('change', () => {
                updateDateRangeDisplay();
                updateDateFilterConfig();
            });
        });
        
        const daysFromToday = document.getElementById('daysFromToday');
        const specificDate = document.getElementById('specificDate');
        if (daysFromToday) daysFromToday.addEventListener('input', () => {
            updateDateRangeDisplay();
            updateDateFilterConfig();
        });
        if (specificDate) specificDate.addEventListener('change', () => {
            updateDateRangeDisplay();
            updateDateFilterConfig();
        });
    } catch (error) {
        console.error('Error setting up event listeners:', error);
        showMessage('Error initializing application', 'error');
    }
}

function handleFileUpload(event) {
    const files = event.target.files;
    if (files.length > 0) {
        document.getElementById('runAnalysis').disabled = false;
        const message = files.length === 1 ? `File selected: ${files[0].name}` : `${files.length} files selected for trending`;
        showMessage(message, 'success');
    }
}

async function runAnalysis() {
    const fileInput = document.getElementById('excelFile');
    const files = fileInput.files;
    
    if (files.length === 0) {
        showMessage('Please select Excel file(s) first', 'error');
        return;
    }

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
        formData.append('excelFiles', files[i]);
    }
    
    // Add filter files if selected
    const engineersFile = document.getElementById('engineersFile').files[0];
    const excludeFile = document.getElementById('excludeFile').files[0];
    const saFile = document.getElementById('saFile').files[0];
    
    if (engineersFile) formData.append('engineersFile', engineersFile);
    if (excludeFile) formData.append('excludeFile', excludeFile);
    if (saFile) formData.append('saFile', saFile);
    
    // Add text inputs
    const engineersText = document.getElementById('engineersText').value;
    const exclusionsText = document.getElementById('exclusionsText').value;
    const saText = document.getElementById('saText').value;
    
    if (engineersText.trim()) formData.append('engineersText', engineersText);
    if (exclusionsText.trim()) formData.append('exclusionsText', exclusionsText);
    if (saText.trim()) formData.append('saText', saText);

    try {
        showLoading(true);
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        
        if (result.success) {
            currentAnalysis = result.analysis;
            updateDashboard();
            updateConsultantsTable();
            updateSolutionArchitectsTable();
            updateCustomersTable();
            updateDASTable();
            updateCombinationsTable();
            
            // Enable export button
            const exportBtn = document.getElementById('exportResults');
            if (exportBtn) exportBtn.disabled = false;
            
            showMessage('Analysis completed successfully!', 'success');
        } else {
            showMessage('Error: ' + result.error, 'error');
        }
    } catch (error) {
        showMessage('Error uploading file: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

function showConsultantProjects(consultantName, projects) {
    let projectsHtml = `
        <div class="modal fade" id="projectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">📋 Projects for ${consultantName}</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="jobNumber">Job Number ↕</th>
                                        <th class="sortable" data-sort="description">Description ↕</th>
                                        <th class="sortable" data-sort="customer">Customer ↕</th>
                                        <th class="sortable" data-sort="budgetHrs">Budget Hrs ↕</th>
                                        <th class="sortable" data-sort="actualHrs">Actual Hrs ↕</th>
                                        <th class="sortable" data-sort="variance">Variance % ↕</th>
                                        <th class="sortable" data-sort="completion">Complete % ↕</th>
                                        <th class="sortable" data-sort="status">Status ↕</th>
                                        <th>Links</th>
                                        <th>Actions</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const greenThreshold = parseFloat(document.getElementById('greenThreshold').value) * 100;
        const redThreshold = parseFloat(document.getElementById('redThreshold').value) * 100;
        
        const rowClass = variance > redThreshold ? 'table-danger' : 
                        variance < greenThreshold ? 'table-success' : 'table-warning';
        
        const savantUrl = `https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo=${project.jobNumber}&isPmo=true`;
        const ssrsUrl = `https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber=${project.jobNumber}`;
        
        projectsHtml += `
            <tr class="${rowClass}">
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.customer}</td>
                <td>${project.budgetHrs.toLocaleString()}</td>
                <td>${project.actualHrs.toLocaleString()}</td>
                <td>${project.variance > 0 ? '+' : ''}${project.variance}%</td>
                <td>${(project.completion * 100).toFixed(0)}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
                <td>
                    <a href="${savantUrl}" target="_blank" class="btn btn-sm btn-outline-primary me-1">🔗 Savant</a>
                    <a href="${ssrsUrl}" target="_blank" class="btn btn-sm btn-outline-secondary">📄 SSRS</a>
                </td>
                <td>
                    <button class="btn btn-sm btn-outline-danger exclude-btn" data-consultant="${consultantName}" data-project="${project.jobNumber}">🚫 Exclude</button>
                </td>
            </tr>
        `;
    });
    
    projectsHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    const existingModal = document.getElementById('projectsModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('projectsModal'));
    modal.show();
    
    setTimeout(() => {
        addTableSorting();
        document.querySelectorAll('#projectsModal .exclude-btn').forEach(btn => {
            btn.addEventListener('click', handleExclusion);
        });
    }, 100);
}

async function handleExclusion(event) {
    const consultant = event.target.getAttribute('data-consultant');
    const project = event.target.getAttribute('data-project');
    
    if (confirm(`Exclude ${consultant} from project ${project}?`)) {
        try {
            const response = await fetch('/exclude', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ consultant, project })
            });
            
            if (response.ok) {
                event.target.textContent = '✓ Excluded';
                event.target.disabled = true;
                event.target.classList.remove('btn-outline-danger');
                event.target.classList.add('btn-success');
                showMessage(`${consultant} excluded from ${project}`, 'success');
            }
        } catch (error) {
            showMessage('Error saving exclusion: ' + error.message, 'error');
        }
    }
}

async function viewExclusions() {
    try {
        const response = await fetch('/exclusions');
        const exclusions = await response.json();
        
        let exclusionsHtml = `
            <div class="modal fade" id="exclusionsModal" tabindex="-1">
                <div class="modal-dialog modal-lg">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title">🚫 Current Exclusions</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                        </div>
                        <div class="modal-body">
                            <div class="table-responsive">
                                <table class="table table-striped">
                                    <thead>
                                        <tr>
                                            <th>Consultant</th>
                                            <th>Project</th>
                                            <th>Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody>
        `;
        
        exclusions.forEach(exclusion => {
            exclusionsHtml += `
                <tr>
                    <td>${exclusion.consultant}</td>
                    <td>${exclusion.project}</td>
                    <td>
                        <button class="btn btn-sm btn-outline-success remove-exclusion-btn" data-consultant="${exclusion.consultant}" data-project="${exclusion.project}">✓ Remove</button>
                    </td>
                </tr>
            `;
        });
        
        exclusionsHtml += `
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        const existingModal = document.getElementById('exclusionsModal');
        if (existingModal) existingModal.remove();
        
        document.body.insertAdjacentHTML('beforeend', exclusionsHtml);
        const modal = new bootstrap.Modal(document.getElementById('exclusionsModal'));
        modal.show();
        
        document.querySelectorAll('.remove-exclusion-btn').forEach(btn => {
            btn.addEventListener('click', removeExclusion);
        });
        
    } catch (error) {
        showMessage('Error loading exclusions: ' + error.message, 'error');
    }
}

async function removeExclusion(event) {
    const consultant = event.target.getAttribute('data-consultant');
    const project = event.target.getAttribute('data-project');
    
    if (confirm(`Remove exclusion for ${consultant} from project ${project}?`)) {
        try {
            const response = await fetch('/exclude', {
                method: 'DELETE',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ consultant, project })
            });
            
            if (response.ok) {
                event.target.closest('tr').remove();
                showMessage(`Exclusion removed for ${consultant}`, 'success');
            }
        } catch (error) {
            showMessage('Error removing exclusion: ' + error.message, 'error');
        }
    }
}

function showMessage(message, type) {
    const toast = document.createElement('div');
    toast.className = `alert alert-${type === 'error' ? 'danger' : 'success'} alert-dismissible fade show position-fixed`;
    toast.style.top = '20px';
    toast.style.right = '20px';
    toast.style.zIndex = '9999';
    toast.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    document.body.appendChild(toast);
    
    setTimeout(() => {
        if (toast.parentNode) {
            toast.parentNode.removeChild(toast);
        }
    }, 5000);
}

function updateDashboard() {
    if (!currentAnalysis) return;
    
    const totalHours = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.hours), 0);
    const avgEfficiency = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.efficiencyScore), 0) / currentAnalysis.consultants.length;
    
    document.getElementById('totalConsultants').innerHTML = `<div><strong>${currentAnalysis.consultants.length}</strong></div>`;
    document.getElementById('totalProjects').innerHTML = `<div><strong>${currentAnalysis.totalProjects}</strong></div>`;
    document.getElementById('totalHours').innerHTML = `<div><strong>${totalHours.toLocaleString()}</strong></div>`;
    document.getElementById('avgEfficiency').innerHTML = `<div><strong>${avgEfficiency.toFixed(1)}%</strong></div>`;
    
    // Update Performance Highlights
    updatePerformanceHighlights();
    
    // Update Performance Chart
    updatePerformanceChart();
}

function updateConsultantsTable() {
    if (!currentAnalysis) return;
    const tbody = document.querySelector('#consultantsTable tbody');
    tbody.innerHTML = '';
    currentAnalysis.consultants.forEach(consultant => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong><a href="#" class="consultant-link" data-consultant="${sanitizeHtml(consultant.name)}">${sanitizeHtml(consultant.name)}</a></strong></td>
            <td>${consultant.projects}</td>
            <td>${parseFloat(consultant.hours).toLocaleString()}</td>
            <td>${consultant.efficiencyScore}%</td>
            <td>${consultant.successRate}%</td>
            <td><span class="badge bg-success">${consultant.withinBudget}</span></td>
            <td><span class="badge bg-danger">${consultant.overBudget}</span></td>
        `;
        const consultantLink = row.querySelector('.consultant-link');
        consultantLink.addEventListener('click', (e) => {
            e.preventDefault();
            showConsultantProjects(consultant.name, consultant.projectDetails);
        });
        tbody.appendChild(row);
    });
}

function updateSolutionArchitectsTable() {
    if (!currentAnalysis || !currentAnalysis.solutionArchitects) return;
    
    const saTab = document.getElementById('solutionArchitects');
    if (!saTab) return;
    
    let saHtml = `
        <div class="card" style="border: none; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 12px; margin-top: 1rem;">
            <div class="card-header" style="background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: white; border-radius: 12px 12px 0 0; border: none;">
                <h5 class="mb-0" style="font-weight: 600;">🏗️ Solution Architect Performance</h5>
            </div>
            <div class="card-body" style="background: white; border-radius: 0 0 12px 12px;">
                <div class="table-responsive">
                    <table class="table table-striped sortable-table">
                        <thead>
                            <tr>
                                <th class="sortable" data-sort="name">Solution Architect ↕</th>
                                <th class="sortable" data-sort="projects">Projects ↕</th>
                                <th class="sortable" data-sort="budgetedHours">Budgeted Hours ↕</th>
                                <th class="sortable" data-sort="actualHours">Actual Hours ↕</th>
                                <th class="sortable" data-sort="successRate">Success Rate ↕</th>
                                <th class="sortable" data-sort="variance">Variance ↕</th>
                            </tr>
                        </thead>
                        <tbody>
    `;
    
    currentAnalysis.solutionArchitects.forEach(sa => {
        saHtml += `
            <tr>
                <td><strong><a href="#" class="sa-link" data-sa="${sa.name}">${sa.name}</a></strong></td>
                <td>${sa.projects}</td>
                <td>${parseFloat(sa.budgetedHours).toLocaleString()}</td>
                <td>${parseFloat(sa.actualHours).toLocaleString()}</td>
                <td>${sa.successRate}%</td>
                <td>${sa.variance > 0 ? '+' : ''}${sa.variance}%</td>
            </tr>
        `;
    });
    
    saHtml += `
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
    
    saTab.innerHTML = saHtml;
    
    // Add SA Performance Highlights
    const topSAs = currentAnalysis.solutionArchitects.slice(0, 3);
    const highlightsHtml = `
        <div class="alert alert-info mb-3">
            <small><i class="fas fa-info-circle"></i> Performance highlights are for engagement and fun. Data may contain errors - use detailed analysis for business decisions.</small>
        </div>
        <div class="row">
            ${topSAs.map((sa, index) => {
                const medal = index === 0 ? '🥇' : index === 1 ? '🥈' : '🥉';
                const title = index === 0 ? 'Top SA' : index === 1 ? 'Runner Up' : 'Third Place';
                return `
                    <div class="col-md-4 mb-3">
                        <div class="card border-primary">
                            <div class="card-body text-center">
                                <h5 class="card-title">${medal} ${title}</h5>
                                <h6 class="card-subtitle mb-2 text-muted">
                                    <a href="#" class="sa-highlight-link" data-sa="${sa.name}">${sa.name}</a>
                                </h6>
                                <p class="card-text">
                                    <strong>Success Rate:</strong> ${sa.successRate}%<br>
                                    <strong>Projects:</strong> ${sa.projects}<br>
                                    <strong>Budget Hours:</strong> ${parseFloat(sa.budgetedHours).toLocaleString()}
                                </p>
                            </div>
                        </div>
                    </div>
                `;
            }).join('')}
        </div>
    `;
    
    saTab.insertAdjacentHTML('beforeend', highlightsHtml);
    
    // Add event listeners
    saTab.querySelectorAll('.sa-link, .sa-highlight-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const saName = e.target.getAttribute('data-sa');
            const sa = currentAnalysis.solutionArchitects.find(s => s.name === saName);
            if (sa) showSAProjects(sa.name, sa.projectDetails);
        });
    });
    
    // Add table sorting
    addTableSorting();
}

function updateCustomersTable() {
    if (!currentAnalysis || !currentAnalysis.customers) return;
    
    const customersTab = document.getElementById('customers');
    if (!customersTab) return;
    
    let customersHtml = `
        <div class="card" style="border: none; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 12px; margin-top: 1rem;">
            <div class="card-header" style="background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: white; border-radius: 12px 12px 0 0; border: none;">
                <h5 class="mb-0" style="font-weight: 600;">🏢 Practice vs Company Customer Performance</h5>
            </div>
            <div class="card-body" style="background: white; border-radius: 0 0 12px 12px;">
                <ul class="nav nav-pills mb-3" id="customerTabs">
                    <li class="nav-item">
                        <a class="nav-link active" data-bs-toggle="pill" href="#practiceCustomers">Practice Customers</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" data-bs-toggle="pill" href="#companyCustomers">All Company Customers</a>
                    </li>
                </ul>
                <div class="tab-content">
                    <div class="tab-pane fade show active" id="practiceCustomers">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="customer">Customer ↕</th>
                                        <th class="sortable" data-sort="projects">Projects ↕</th>
                                        <th class="sortable" data-sort="successRate">Success Rate ↕</th>
                                        <th class="sortable" data-sort="avgVariance">Avg Variance ↕</th>
                                        <th class="sortable" data-sort="budgetHrs">Total Budget Hrs ↕</th>
                                        <th class="sortable" data-sort="riskLevel">Risk Level ↕</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    currentAnalysis.customers.practice.forEach(customer => {
        customersHtml += `
            <tr>
                <td><strong><a href="#" class="customer-link" data-customer="${customer.name}" data-type="practice">${customer.name}</a></strong></td>
                <td>${customer.projects}</td>
                <td>${customer.successRate}%</td>
                <td>${customer.avgVariance > 0 ? '+' : ''}${customer.avgVariance}%</td>
                <td>${parseFloat(customer.totalBudgetHrs).toLocaleString()}</td>
                <td><span class="badge bg-${customer.riskColor}">${customer.riskLevel}</span></td>
            </tr>
        `;
    });
    
    customersHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="tab-pane fade" id="companyCustomers">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="customer">Customer ↕</th>
                                        <th class="sortable" data-sort="projects">Projects ↕</th>
                                        <th class="sortable" data-sort="successRate">Success Rate ↕</th>
                                        <th class="sortable" data-sort="avgVariance">Avg Variance ↕</th>
                                        <th class="sortable" data-sort="budgetHrs">Total Budget Hrs ↕</th>
                                        <th class="sortable" data-sort="riskLevel">Risk Level ↕</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    currentAnalysis.customers.company.forEach(customer => {
        customersHtml += `
            <tr>
                <td><strong><a href="#" class="customer-link" data-customer="${customer.name}" data-type="company">${customer.name}</a></strong></td>
                <td>${customer.projects}</td>
                <td>${customer.successRate}%</td>
                <td>${customer.avgVariance > 0 ? '+' : ''}${customer.avgVariance}%</td>
                <td>${parseFloat(customer.totalBudgetHrs).toLocaleString()}</td>
                <td><span class="badge bg-${customer.riskColor}">${customer.riskLevel}</span></td>
            </tr>
        `;
    });
    
    customersHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    customersTab.innerHTML = customersHtml;
    
    // Add event listeners
    customersTab.querySelectorAll('.customer-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const customerName = e.target.getAttribute('data-customer');
            const type = e.target.getAttribute('data-type');
            const customer = currentAnalysis.customers[type].find(c => c.name === customerName);
            if (customer) showCustomerProjects(customer.name, customer.projectDetails, type === 'practice' ? 'Practice' : 'Company');
        });
    });
    
    // Add table sorting
    addTableSorting();
}

function updateDASTable() {
    if (!currentAnalysis || !currentAnalysis.dasAnalysis) return;
    
    const dasTab = document.getElementById('das');
    if (!dasTab) return;
    
    let dasHtml = `
        <div class="card" style="border: none; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 12px; margin-top: 1rem;">
            <div class="card-header" style="background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: white; border-radius: 12px 12px 0 0; border: none;">
                <h5 class="mb-0" style="font-weight: 600;">📊 DAS+ (Delivery Accuracy Score Plus) Analysis</h5>
            </div>
            <div class="card-body" style="background: white; border-radius: 0 0 12px 12px;">
                <div class="table-responsive">
                    <table class="table table-striped sortable-table">
                        <thead>
                            <tr>
                                <th class="sortable" data-sort="name">Consultant ↕</th>
                                <th class="sortable" data-sort="projects">Projects ↕</th>
                                <th class="sortable" data-sort="avgDAS">Avg DAS+ ↕</th>
                                <th class="sortable" data-sort="medianDAS">Median DAS+ ↕</th>
                                <th class="sortable" data-sort="highCount">High Performing (≥0.85) ↕</th>
                                <th class="sortable" data-sort="reviewCount">Needs Review (0.3-0.9) ↕</th>
                                <th class="sortable" data-sort="lowCount">Low Performing (<0.75) ↕</th>
                            </tr>
                        </thead>
                        <tbody>
    `;
    
    currentAnalysis.dasAnalysis.forEach(das => {
        dasHtml += `
            <tr>
                <td><strong>${das.name}</strong></td>
                <td>${das.projects}</td>
                <td>${das.avgDAS}</td>
                <td>${das.medianDAS}</td>
                <td><span class="badge bg-success">${das.highCount}</span></td>
                <td><span class="badge bg-warning">${das.reviewCount}</span></td>
                <td><span class="badge bg-danger">${das.lowCount}</span></td>
            </tr>
        `;
    });
    
    dasHtml += `
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
    
    dasTab.innerHTML = dasHtml;
    
    // Add table sorting
    addTableSorting();
}

function updateCombinationsTable() {
    if (!currentAnalysis) return;
    
    const combinationsTab = document.getElementById('combinations');
    if (!combinationsTab) return;
    
    let combinationsHtml = `
        <div class="card" style="border: none; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 12px; margin-top: 1rem;">
            <div class="card-header" style="background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: white; border-radius: 12px 12px 0 0; border: none;">
                <h5 class="mb-0" style="font-weight: 600;">🤝 Team Combination Analysis</h5>
            </div>
            <div class="card-body" style="background: white; border-radius: 0 0 12px 12px;">
                <ul class="nav nav-pills mb-3">
                    <li class="nav-item">
                        <a class="nav-link active" data-bs-toggle="pill" href="#saCombinations">SA-Consultant Combinations</a>
                    </li>
                    <li class="nav-item">
                        <a class="nav-link" data-bs-toggle="pill" href="#saCustomerCombinations">SA-Customer Analysis</a>
                    </li>
                </ul>
                <div class="tab-content">
                    <div class="tab-pane fade show active" id="saCombinations">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="sa">Solution Architect ↕</th>
                                        <th class="sortable" data-sort="consultant">Consultant ↕</th>
                                        <th class="sortable" data-sort="projects">Projects ↕</th>
                                        <th class="sortable" data-sort="successRate">Success Rate ↕</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    if (currentAnalysis.saCombinations) {
        currentAnalysis.saCombinations.forEach(combo => {
            combinationsHtml += `
                <tr>
                    <td><strong>${combo.sa}</strong></td>
                    <td><strong>${combo.consultant}</strong></td>
                    <td>${combo.projects}</td>
                    <td>${combo.successRate}%</td>
                </tr>
            `;
        });
    }
    
    combinationsHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="tab-pane fade" id="saCustomerCombinations">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="sa">Solution Architect ↕</th>
                                        <th class="sortable" data-sort="customer">Customer ↕</th>
                                        <th class="sortable" data-sort="projects">Projects ↕</th>
                                        <th class="sortable" data-sort="successRate">Success Rate ↕</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    if (currentAnalysis.saCustomerAnalysis) {
        currentAnalysis.saCustomerAnalysis.forEach(combo => {
            combinationsHtml += `
                <tr>
                    <td><strong>${combo.sa}</strong></td>
                    <td><strong>${combo.customer}</strong></td>
                    <td>${combo.projects}</td>
                    <td>${combo.successRate}%</td>
                </tr>
            `;
        });
    }
    
    combinationsHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    combinationsTab.innerHTML = combinationsHtml;
    
    // Add table sorting
    addTableSorting();
}

function updatePerformanceHighlights() {
    if (!currentAnalysis || !currentAnalysis.consultantOfQuarter) return;
    
    // Add performance highlights to dashboard
    const dashboardTab = document.getElementById('dashboard');
    if (!dashboardTab) return;
    
    let highlightsContainer = document.getElementById('performanceHighlights');
    if (!highlightsContainer) {
        highlightsContainer = document.createElement('div');
        highlightsContainer.id = 'performanceHighlights';
        highlightsContainer.className = 'mt-4';
        dashboardTab.appendChild(highlightsContainer);
    }
    
    const topPerformers = currentAnalysis.consultantOfQuarter.slice(0, 3);
    
    let highlightsHtml = `
        <div class="alert alert-info mb-3">
            <small><i class="fas fa-info-circle"></i> Performance highlights are for engagement and fun. Data may contain errors - use detailed analysis for business decisions.</small>
        </div>
        <div class="row">
    `;
    
    topPerformers.forEach((performer, index) => {
        const medal = index === 0 ? '🥇' : index === 1 ? '🥈' : '🥉';
        const title = index === 0 ? 'Consultant of the Quarter' : index === 1 ? 'Runner Up' : 'Third Place';
        
        highlightsHtml += `
            <div class="col-md-4 mb-3">
                <div class="card border-primary">
                    <div class="card-body text-center">
                        <h5 class="card-title">${medal} ${title}</h5>
                        <h6 class="card-subtitle mb-2 text-muted">
                            <a href="#" class="consultant-highlight-link" data-consultant="${performer.name}">${performer.name}</a>
                        </h6>
                        <p class="card-text">
                            <strong>Score:</strong> ${performer.compositeScore}<br>
                            <strong>Success Rate:</strong> ${performer.successRate}%<br>
                            <strong>Projects:</strong> ${performer.projects}<br>
                            <strong>Hours:</strong> ${parseFloat(performer.hours).toLocaleString()}
                        </p>
                    </div>
                </div>
            </div>
        `;
    });
    
    highlightsHtml += '</div>';
    highlightsContainer.innerHTML = highlightsHtml;
    
    // Add click handlers for consultant names
    document.querySelectorAll('.consultant-highlight-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const consultantName = e.target.getAttribute('data-consultant');
            const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
            if (consultant) {
                showConsultantProjects(consultant.name, consultant.projectDetails);
            }
        });
    });
}

function showSAProjects(saName, projects) {
    let projectsHtml = `
        <div class="modal fade" id="saProjectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">📋 Projects for SA: ${saName}</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Job Number</th>
                                        <th>Description</th>
                                        <th>Customer</th>
                                        <th>Budget Hrs</th>
                                        <th>Actual Hrs</th>
                                        <th>Variance %</th>
                                        <th>Status</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const rowClass = variance > 30 ? 'table-danger' : variance < -10 ? 'table-success' : 'table-warning';
        
        projectsHtml += `
            <tr class="${rowClass}">
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.customer}</td>
                <td>${project.budgetHrs.toLocaleString()}</td>
                <td>${project.actualHrs.toLocaleString()}</td>
                <td>${project.variance > 0 ? '+' : ''}${project.variance}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
            </tr>
        `;
    });
    
    projectsHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    const existingModal = document.getElementById('saProjectsModal');
    if (existingModal) existingModal.remove();
    
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('saProjectsModal'));
    modal.show();
}

function showCustomerProjects(customerName, projects, type) {
    let projectsHtml = `
        <div class="modal fade" id="customerProjectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">📋 ${type} Projects for: ${customerName}</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Job Number</th>
                                        <th>Description</th>
                                        <th>Budget Hrs</th>
                                        <th>Actual Hrs</th>
                                        <th>Variance %</th>
                                        <th>Resources</th>
                                        <th>Solution Architect</th>
                                        <th>Status</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const rowClass = variance > 30 ? 'table-danger' : variance < -10 ? 'table-success' : 'table-warning';
        
        projectsHtml += `
            <tr class="${rowClass}">
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.budgetHrs.toLocaleString()}</td>
                <td>${project.actualHrs.toLocaleString()}</td>
                <td>${project.variance > 0 ? '+' : ''}${project.variance}%</td>
                <td>${project.resources}</td>
                <td>${project.solutionArchitect}</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
            </tr>
        `;
    });
    
    projectsHtml += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    const existingModal = document.getElementById('customerProjectsModal');
    if (existingModal) existingModal.remove();
    
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('customerProjectsModal'));
    modal.show();
}

function addTableSorting() {
    document.querySelectorAll('.sortable-table').forEach(table => {
        const headers = table.querySelectorAll('th.sortable');
        headers.forEach(header => {
            header.addEventListener('click', () => {
                const column = header.getAttribute('data-sort');
                sortTable(table, column, header);
            });
        });
    });
}

function sortTable(table, column, header) {
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const isAscending = header.classList.contains('sort-asc');
    
    rows.sort((a, b) => {
        const aVal = a.cells[header.cellIndex].textContent.trim();
        const bVal = b.cells[header.cellIndex].textContent.trim();
        
        const aNum = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
        const bNum = parseFloat(bVal.replace(/[^0-9.-]/g, ''));
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return isAscending ? bNum - aNum : aNum - bNum;
        } else {
            return isAscending ? bVal.localeCompare(aVal) : aVal.localeCompare(bVal);
        }
    });
    
    // Update header classes
    table.querySelectorAll('th.sortable').forEach(h => {
        h.classList.remove('sort-asc', 'sort-desc');
        const text = h.textContent.replace(/ [↑↓↕]/g, '');
        h.innerHTML = text + ' ↕';
    });
    
    const headerText = header.textContent.replace(/ [↑↓↕]/g, '');
    if (isAscending) {
        header.classList.add('sort-desc');
        header.innerHTML = headerText + ' ↓';
    } else {
        header.classList.add('sort-asc');
        header.innerHTML = headerText + ' ↑';
    }
    
    // Reorder rows
    rows.forEach(row => tbody.appendChild(row));
}

function exportToCSV() {
    if (!currentAnalysis) {
        showMessage('No analysis data to export', 'error');
        return;
    }
    
    let csv = 'Consultant,Projects,Hours,Efficiency %,Success Rate %,Within Budget,Over Budget\n';
    currentAnalysis.consultants.forEach(c => {
        csv += `"${c.name}",${c.projects},${c.hours},${c.efficiencyScore},${c.successRate},${c.withinBudget},${c.overBudget}\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `consultant-analysis-${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
    
    showMessage('Analysis exported to CSV', 'success');
}

function updatePerformanceChart() {
    if (!currentAnalysis || !currentAnalysis.consultants) return;
    
    const ctx = document.getElementById('performanceChart');
    if (!ctx) return;
    
    if (performanceChart) {
        performanceChart.destroy();
    }
    
    const topConsultants = currentAnalysis.consultants.slice(0, 10);
    
    performanceChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topConsultants.map(c => c.name),
            datasets: [{
                label: 'Success Rate %',
                data: topConsultants.map(c => parseFloat(c.successRate)),
                backgroundColor: 'rgba(59, 130, 246, 0.8)',
                borderColor: 'rgba(59, 130, 246, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: 'Top 10 Consultant Performance'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100
                }
            }
        }
    });
}

function toggleDataSource() {
    const fileUpload = document.getElementById('fileUploadSection');
    const sqlConnection = document.getElementById('sqlConnectionSection');
    const selectedSource = document.querySelector('input[name="dataSource"]:checked').value;
    
    if (selectedSource === 'file') {
        fileUpload.style.display = 'block';
        sqlConnection.style.display = 'none';
    } else {
        fileUpload.style.display = 'none';
        sqlConnection.style.display = 'block';
    }
}

function handleConfigUpload() {
    const fileInput = document.getElementById('configFile');
    const file = fileInput.files[0];
    
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const config = JSON.parse(e.target.result);
                
                // Validate config structure
                if (!config || typeof config !== 'object') {
                    throw new Error('Invalid configuration format');
                }
                
                // Apply thresholds with validation
                if (config.thresholds) {
                    const t = config.thresholds;
                    if (t.efficiency_threshold >= 0 && t.efficiency_threshold <= 1) {
                        document.getElementById('efficiencyThreshold').value = t.efficiency_threshold;
                    }
                    if (t.success_threshold >= 0 && t.success_threshold <= 1) {
                        document.getElementById('successThreshold').value = t.success_threshold;
                    }
                    if (t.green_threshold >= -1 && t.green_threshold <= 1) {
                        document.getElementById('greenThreshold').value = t.green_threshold;
                    }
                    if (t.yellow_threshold >= 0 && t.yellow_threshold <= 1) {
                        document.getElementById('yellowThreshold').value = t.yellow_threshold;
                    }
                    if (t.red_threshold >= 0 && t.red_threshold <= 2) {
                        document.getElementById('redThreshold').value = t.red_threshold;
                    }
                }
                
                // Apply trending settings
                if (config.trending) {
                    const trending = document.getElementById('enableTrending');
                    if (trending) {
                        trending.checked = !!config.trending.enable_trending;
                        toggleTrending(); // Apply the trending setting
                    }
                }
                
                // Apply project filtering settings
                if (config.project_filtering) {
                    const pf = config.project_filtering;
                    const enableDateFilter = document.getElementById('enableDateFilter');
                    const daysFromToday = document.getElementById('daysFromToday');
                    const specificDate = document.getElementById('specificDate');
                    
                    if (enableDateFilter) {
                        enableDateFilter.checked = !!pf.enable_date_filter;
                        toggleDateFilter(); // Show/hide date options
                    }
                    
                    // Set filter type radio buttons
                    if (pf.filter_type) {
                        const filterTypeRadio = document.querySelector(`input[name="filterType"][value="${pf.filter_type}"]`);
                        if (filterTypeRadio) {
                            filterTypeRadio.checked = true;
                        }
                    }
                    
                    // Set days and specific date values
                    if (daysFromToday && pf.days_from_today && pf.days_from_today > 0) {
                        daysFromToday.value = pf.days_from_today;
                    }
                    if (specificDate && pf.specific_date) {
                        specificDate.value = pf.specific_date;
                    }
                }
                
                applyConfigToGUI(config);
                showMessage('Configuration loaded and applied successfully', 'success');
            } catch (error) {
                showMessage('Error loading configuration: ' + error.message, 'error');
            }
        };
        reader.readAsText(file);
    }
}

function connectToSQL() {
    showMessage('SQL connection not implemented yet', 'error');
}

async function updateSingleThreshold(key, value) {
    try {
        const response = await fetch('/config');
        let config = response.ok ? await response.json() : {};
        
        if (!config.thresholds) config.thresholds = {};
        config.thresholds[key] = value;
        
        await fetch('/config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(config)
        });
    } catch (error) {
        console.error('Error updating threshold:', error);
    }
}

async function updateTrendingConfig() {
    try {
        const response = await fetch('/config');
        let config = response.ok ? await response.json() : {};
        
        if (!config.trending) config.trending = {};
        config.trending.enable_trending = document.getElementById('enableTrending').checked;
        
        await fetch('/config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(config)
        });
    } catch (error) {
        console.error('Error updating trending config:', error);
    }
}

async function updateDateFilterConfig() {
    try {
        const response = await fetch('/config');
        let config = response.ok ? await response.json() : {};
        
        if (!config.project_filtering) config.project_filtering = {};
        
        const enableDateFilter = document.getElementById('enableDateFilter');
        const filterType = document.querySelector('input[name="filterType"]:checked');
        const daysFromToday = document.getElementById('daysFromToday');
        const specificDate = document.getElementById('specificDate');
        
        if (enableDateFilter) config.project_filtering.enable_date_filter = enableDateFilter.checked;
        if (filterType) config.project_filtering.filter_type = filterType.value;
        if (daysFromToday) config.project_filtering.days_from_today = parseInt(daysFromToday.value) || 365;
        if (specificDate) config.project_filtering.specific_date = specificDate.value;
        
        await fetch('/config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(config)
        });
    } catch (error) {
        console.error('Error updating date filter config:', error);
    }
}

function toggleDateFilter() {
    const enabled = document.getElementById('enableDateFilter').checked;
    const dateOptions = document.getElementById('dateFilterOptions');
    dateOptions.style.display = enabled ? 'block' : 'none';
    updateDateRangeDisplay();
}

function toggleTrending() {
    const enabled = document.getElementById('enableTrending').checked;
    const fileInput = document.getElementById('excelFile');
    
    if (enabled) {
        fileInput.setAttribute('multiple', 'multiple');
        showMessage('Multi-file trending enabled - you can now select multiple Excel files', 'success');
    } else {
        fileInput.removeAttribute('multiple');
        showMessage('Single file mode - select one Excel file', 'success');
    }
}

function updateDateRangeDisplay() {
    const enabled = document.getElementById('enableDateFilter').checked;
    const filterType = document.querySelector('input[name="filterType"]:checked')?.value;
    const daysInput = document.getElementById('daysFromToday');
    const dateInput = document.getElementById('specificDate');
    const dateRangeInfo = document.getElementById('dateRangeInfo');
    
    if (filterType === 'days') {
        daysInput.style.display = 'block';
        dateInput.style.display = 'none';
    } else {
        daysInput.style.display = 'none';
        dateInput.style.display = 'block';
    }
    
    if (enabled) {
        if (filterType === 'days') {
            const days = daysInput.value || 365;
            dateRangeInfo.textContent = `Last ${days} days from today`;
        } else {
            const date = dateInput.value || 'Not set';
            dateRangeInfo.textContent = `From ${date} onwards`;
        }
    } else {
        dateRangeInfo.textContent = 'No date filtering active';
    }
}

function toggleAutoRefresh() {
    const enabled = document.getElementById('enableAutoRefresh').checked;
    
    if (enabled && currentAnalysis) {
        autoRefreshInterval = setInterval(() => {
            if (document.getElementById('excelFile').files.length > 0) {
                runAnalysis();
            }
        }, 300000); // 5 minutes
        showMessage('Auto-refresh enabled (5 minutes)', 'success');
    } else {
        if (autoRefreshInterval) {
            clearInterval(autoRefreshInterval);
            autoRefreshInterval = null;
        }
        showMessage('Auto-refresh disabled', 'success');
    }
}

async function loadExistingConfig() {
    try {
        const response = await fetch('/config');
        if (response.ok) {
            const config = await response.json();
            applyConfigToGUI(config);
        }
    } catch (error) {
        console.log('No existing config found, using defaults');
    }
}

function applyConfigToGUI(config) {
    // Apply thresholds with display updates
    if (config.thresholds) {
        const t = config.thresholds;
        const elements = {
            efficiencyThreshold: { value: t.efficiency_threshold, display: (t.efficiency_threshold * 100).toFixed(0) + '%' },
            successThreshold: { value: t.success_threshold, display: (t.success_threshold * 100).toFixed(0) + '%' },
            greenThreshold: { value: t.green_threshold, display: (t.green_threshold * 100).toFixed(0) + '%' },
            yellowThreshold: { value: t.yellow_threshold, display: (t.yellow_threshold * 100).toFixed(0) + '%' },
            redThreshold: { value: t.red_threshold, display: (t.red_threshold * 100).toFixed(0) + '%' }
        };
        
        Object.entries(elements).forEach(([id, config]) => {
            const el = document.getElementById(id);
            if (el && config.value !== undefined) {
                el.value = config.value;
                // Update the display text
                const displayEl = el.parentElement.querySelector('small');
                if (displayEl) displayEl.textContent = config.display;
            }
        });
    }
    
    // Apply trending
    if (config.trending) {
        const trending = document.getElementById('enableTrending');
        if (trending) {
            trending.checked = !!config.trending.enable_trending;
            toggleTrending();
        }
    }
    
    // Apply project filtering
    if (config.project_filtering) {
        const pf = config.project_filtering;
        const enableDateFilter = document.getElementById('enableDateFilter');
        
        if (enableDateFilter) {
            enableDateFilter.checked = !!pf.enable_date_filter;
            toggleDateFilter();
        }
        
        if (pf.filter_type) {
            const filterTypeRadio = document.querySelector(`input[name="filterType"][value="${pf.filter_type}"]`);
            if (filterTypeRadio) filterTypeRadio.checked = true;
        }
        
        const daysFromToday = document.getElementById('daysFromToday');
        const specificDate = document.getElementById('specificDate');
        
        if (daysFromToday && pf.days_from_today) daysFromToday.value = pf.days_from_today;
        if (specificDate && pf.specific_date) specificDate.value = pf.specific_date;
    }
    
    updateDateRangeDisplay();
}

function showLoading(show) {
    const button = document.getElementById('runAnalysis');
    if (show) {
        button.innerHTML = '<span class="spinner-border spinner-border-sm" role="status"></span> Processing...';
        button.disabled = true;
    } else {
        button.innerHTML = '🚀 Run Analysis';
        button.disabled = false;
    }
}