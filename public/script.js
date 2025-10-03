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
    enableExportButtons(false); // Initially disable export buttons
    
    // Initialize slider displays with default values
    updateSliderDisplay('efficiencyThreshold', 0.15);
    updateSliderDisplay('successThreshold', 0.30);
    updateSliderDisplay('greenThreshold', -0.10);
    updateSliderDisplay('yellowThreshold', 0.10);
    updateSliderDisplay('redThreshold', 0.30);
    
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
        const viewDQBtn = document.getElementById('viewDQ');
        const connectSQLBtn = document.getElementById('connectSQL');
        const exportResultsBtn = document.getElementById('exportResults');
        const applyConfigBtn = document.getElementById('applyConfig');
        const refreshConfigBtn = document.getElementById('refreshConfig');
        const saveConfigBtn = document.getElementById('saveConfig');
        if (runAnalysisBtn) runAnalysisBtn.addEventListener('click', runAnalysis);
        if (viewExclusionsBtn) viewExclusionsBtn.addEventListener('click', viewExclusions);
        if (viewDQBtn) viewDQBtn.addEventListener('click', viewDisqualified);
        if (connectSQLBtn) connectSQLBtn.addEventListener('click', connectToSQL);
        if (exportResultsBtn) exportResultsBtn.addEventListener('click', exportToCSV);
        if (saveConfigBtn) saveConfigBtn.addEventListener('click', saveCurrentConfig);
        
        // Export buttons
        const exportButtons = {
            'exportConsultantsCSV': () => exportTabData('consultants', 'csv'),
            'exportConsultantsXLSX': () => exportTabData('consultants', 'xlsx'),
            'exportAllCSV': () => exportAllData('csv'),
            'exportAllXLSX': () => exportAllData('xlsx')
        };
        
        Object.entries(exportButtons).forEach(([id, handler]) => {
            const btn = document.getElementById(id);
            if (btn) btn.addEventListener('click', handler);
        });
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
            if (el) {
                el.addEventListener('input', () => {
                    const value = parseFloat(el.value);
                    updateSliderDisplay(id, value);
                });
            }
        });
        
        const enableDateFilter = document.getElementById('enableDateFilter');
        const enableTrending = document.getElementById('enableTrending');
        const enableAutoRefresh = document.getElementById('enableAutoRefresh');
        const clearExclusionsCache = document.getElementById('clearExclusionsCache');
        if (enableDateFilter) enableDateFilter.addEventListener('change', () => {
            toggleDateFilter();
            updateDateFilterConfig();
        });
        if (enableTrending) enableTrending.addEventListener('change', () => {
            toggleTrending();
            updateTrendingConfig();
        });
        if (enableAutoRefresh) enableAutoRefresh.addEventListener('change', toggleAutoRefresh);
        if (clearExclusionsCache) clearExclusionsCache.addEventListener('click', clearCache);
        
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
    
    // If no files selected but we have cached analysis, use direct analysis endpoint
    if (files.length === 0) {
        if (currentAnalysis) {
            try {
                showLoading(true);
                showMessage('Using cached data for analysis...', 'info');
                
                const currentThresholds = {
                    efficiency_threshold: parseFloat(document.getElementById('efficiencyThreshold').value),
                    success_threshold: parseFloat(document.getElementById('successThreshold').value),
                    green_threshold: parseFloat(document.getElementById('greenThreshold').value),
                    yellow_threshold: parseFloat(document.getElementById('yellowThreshold').value),
                    red_threshold: parseFloat(document.getElementById('redThreshold').value)
                };
                
                const response = await fetch('/analyze', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ thresholds: currentThresholds })
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
                    showMessage('Analysis completed using cached data!', 'success');
                    enableExportButtons(true);
                } else {
                    showMessage('Error: ' + result.error, 'error');
                }
            } catch (error) {
                showMessage('Error running analysis: ' + error.message, 'error');
            } finally {
                showLoading(false);
            }
            return;
        } else {
            showMessage('Please select Excel file(s) first', 'error');
            return;
        }
    }

    // Use file upload for new files
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
    
    // Add current threshold settings
    const currentThresholds = {
        efficiency_threshold: parseFloat(document.getElementById('efficiencyThreshold').value),
        success_threshold: parseFloat(document.getElementById('successThreshold').value),
        green_threshold: parseFloat(document.getElementById('greenThreshold').value),
        yellow_threshold: parseFloat(document.getElementById('yellowThreshold').value),
        red_threshold: parseFloat(document.getElementById('redThreshold').value)
    };
    formData.append('thresholds', JSON.stringify(currentThresholds));

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
            
            // Enable export buttons
            const exportBtn = document.getElementById('exportResults');
            if (exportBtn) exportBtn.disabled = false;
            
            enableExportButtons(true);
            
            showMessage('Analysis completed successfully!', 'success');
            enableExportButtons(true);
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
            <div class="modal-dialog" style="max-width: 95vw;">
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
                        
                        <!-- Export Projects -->
                        <div class="mt-4 mb-3">
                            <div class="btn-group" role="group">
                                <button class="btn btn-sm btn-success" id="exportProjectsCSV">📄 Export CSV</button>
                                <button class="btn btn-sm btn-primary" id="exportProjectsXLSX">📊 Export Excel</button>
                            </div>
                        </div>
                        
                        <!-- Random Projects for Review -->
                        <div class="mt-4">
                            <div class="d-flex justify-content-between align-items-center mb-2">
                                <h6 class="text-primary mb-0">📋 Random Projects for Review</h6>
                                <div class="btn-group" role="group">
                                    <button class="btn btn-sm btn-primary" id="randomFiltered">Filtered Data</button>
                                    <button class="btn btn-sm btn-outline-secondary" id="randomAll">All Data</button>
                                    <button class="btn btn-sm btn-outline-info" id="refreshRandom">🔄 New Random</button>
                                </div>
                            </div>
                            <div class="table-responsive">
                                <table class="table table-sm table-bordered">
                                    <thead class="table-light">
                                        <tr>
                                            <th>Job Number</th>
                                            <th>Description</th>
                                            <th>Customer</th>
                                            <th>Variance %</th>
                                            <th>Status</th>
                                            <th>Links</th>
                                            <th>Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="randomProjectsTable">
                                    </tbody>
                                </table>
                            </div>
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
        
        // Add event listeners for random project controls
        document.getElementById('randomFiltered').addEventListener('click', () => {
            loadRandomProjects(consultantName, projects, 'filtered');
        });
        document.getElementById('randomAll').addEventListener('click', () => {
            loadRandomProjects(consultantName, projects, 'all');
        });
        document.getElementById('refreshRandom').addEventListener('click', () => {
            const activeBtn = document.querySelector('#projectsModal .btn-group .btn-primary');
            const mode = activeBtn?.id === 'randomAll' ? 'all' : 'filtered';
            loadRandomProjects(consultantName, projects, mode);
        });
        
        // Load initial random projects
        loadRandomProjects(consultantName, projects, 'filtered');
        
        // Add export handlers
        document.getElementById('exportProjectsCSV').addEventListener('click', () => {
            const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
            exportConsultantProjects(consultantName, projects, 'csv', consultant);
        });
        document.getElementById('exportProjectsXLSX').addEventListener('click', () => {
            const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
            exportConsultantProjects(consultantName, projects, 'xlsx', consultant);
        });
    }, 100);
}

async function handleExclusion(event) {
    const consultant = event.target.getAttribute('data-consultant');
    const project = event.target.getAttribute('data-project');
    
    // Show reason dialog
    const reason = await showExclusionReasonDialog(consultant, project);
    if (!reason) return; // User cancelled
    
    try {
        const response = await fetch('/exclude', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ consultant, project, reason })
        });
        
        if (response.ok) {
            event.target.textContent = '✓ Excluded';
            event.target.disabled = true;
            event.target.classList.remove('btn-outline-danger');
            event.target.classList.add('btn-success');
            showMessage(`${consultant} excluded from ${project}`, 'success');
            
            // Auto-recalculate using direct analysis
            if (currentAnalysis) {
                showMessage('Recalculating metrics...', 'info');
                try {
                    const currentThresholds = {
                        efficiency_threshold: parseFloat(document.getElementById('efficiencyThreshold').value),
                        success_threshold: parseFloat(document.getElementById('successThreshold').value),
                        green_threshold: parseFloat(document.getElementById('greenThreshold').value),
                        yellow_threshold: parseFloat(document.getElementById('yellowThreshold').value),
                        red_threshold: parseFloat(document.getElementById('redThreshold').value)
                    };
                    
                    const analysisResponse = await fetch('/analyze', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ thresholds: currentThresholds })
                    });
                    if (analysisResponse.ok) {
                        const result = await analysisResponse.json();
                        currentAnalysis = result.analysis;
                        updateDashboard();
                        updateConsultantsTable();
                        updateSolutionArchitectsTable();
                        updateCustomersTable();
                        updateDASTable();
                        updateCombinationsTable();
                        showMessage('Metrics updated successfully!', 'success');
                    } else {
                        showMessage('Exclusion saved. Please re-run analysis manually.', 'warning');
                    }
                } catch (error) {
                    showMessage('Exclusion saved. Please re-run analysis manually.', 'warning');
                }
            }
        } else {
            const errorData = await response.json();
            showMessage('Error: ' + (errorData.error || 'Failed to save exclusion'), 'error');
        }
    } catch (error) {
        showMessage('Error saving exclusion: ' + error.message, 'error');
    }
}

async function viewExclusions() {
    try {
        const response = await fetch('/exclusions');
        const exclusions = await response.json();
        
        let exclusionsHtml = `
            <div class="modal fade" id="exclusionsModal" tabindex="-1">
                <div class="modal-dialog modal-xl">
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
                                            <th>Reason</th>
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
                    <td><small class="text-muted">${exclusion.reason || 'No reason provided'}</small></td>
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
                
                // Auto-recalculate using direct analysis if we have cached data
                if (currentAnalysis) {
                    showMessage('Recalculating metrics...', 'info');
                    try {
                        const analysisResponse = await fetch('/analyze', { method: 'POST' });
                        if (analysisResponse.ok) {
                            const result = await analysisResponse.json();
                            currentAnalysis = result.analysis;
                            updateDashboard();
                            updateConsultantsTable();
                            updateSolutionArchitectsTable();
                            updateCustomersTable();
                            updateDASTable();
                            updateCombinationsTable();
                            showMessage('Metrics updated successfully!', 'success');
                        } else {
                            showMessage('Exclusion saved. Please re-run analysis manually.', 'warning');
                        }
                    } catch (error) {
                        showMessage('Exclusion saved. Please re-run analysis manually.', 'warning');
                    }
                }
            }
        } catch (error) {
            showMessage('Error removing exclusion: ' + error.message, 'error');
        }
    }
}

function showExclusionReasonDialog(consultant, project) {
    return new Promise((resolve) => {
        const dialogHtml = `
            <div class="modal fade" id="exclusionReasonModal" tabindex="-1">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title">🚫 Exclusion Reason Required</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                        </div>
                        <div class="modal-body">
                            <p><strong>Excluding:</strong> ${consultant} from project ${project}</p>
                            <div class="mb-3">
                                <label for="exclusionReason" class="form-label">Reason for exclusion <span class="text-danger">*</span></label>
                                <textarea class="form-control" id="exclusionReason" rows="3" placeholder="Please provide a reason for this exclusion (minimum 2 characters)" maxlength="500"></textarea>
                                <div class="form-text">This reason will be saved and visible when viewing exclusions.</div>
                                <div id="reasonError" class="text-danger" style="display: none;">Reason must be at least 2 characters long.</div>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                            <button type="button" class="btn btn-danger" id="confirmExclusion">Exclude Project</button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        // Remove existing modal if any
        const existingModal = document.getElementById('exclusionReasonModal');
        if (existingModal) existingModal.remove();
        
        document.body.insertAdjacentHTML('beforeend', dialogHtml);
        const modal = new bootstrap.Modal(document.getElementById('exclusionReasonModal'));
        const reasonTextarea = document.getElementById('exclusionReason');
        const confirmBtn = document.getElementById('confirmExclusion');
        const errorDiv = document.getElementById('reasonError');
        
        // Focus on textarea when modal opens
        modal.show();
        setTimeout(() => reasonTextarea.focus(), 300);
        
        // Validate reason on input
        reasonTextarea.addEventListener('input', () => {
            const reason = reasonTextarea.value.trim();
            if (reason.length >= 2) {
                confirmBtn.disabled = false;
                errorDiv.style.display = 'none';
                reasonTextarea.classList.remove('is-invalid');
            } else {
                confirmBtn.disabled = true;
                if (reason.length > 0) {
                    errorDiv.style.display = 'block';
                    reasonTextarea.classList.add('is-invalid');
                }
            }
        });
        
        // Handle confirm button
        confirmBtn.addEventListener('click', () => {
            const reason = reasonTextarea.value.trim();
            if (reason.length >= 2) {
                modal.hide();
                resolve(reason);
            } else {
                errorDiv.style.display = 'block';
                reasonTextarea.classList.add('is-invalid');
                reasonTextarea.focus();
            }
        });
        
        // Handle Enter key in textarea
        reasonTextarea.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
                confirmBtn.click();
            }
        });
        
        // Handle modal close/cancel
        document.getElementById('exclusionReasonModal').addEventListener('hidden.bs.modal', () => {
            document.getElementById('exclusionReasonModal').remove();
            resolve(null); // User cancelled
        });
        
        // Initially disable confirm button
        confirmBtn.disabled = true;
    });
}

function showMessage(message, type) {
    const toast = document.createElement('div');
    const alertType = type === 'error' ? 'danger' : type === 'info' ? 'info' : 'success';
    toast.className = `alert alert-${alertType} alert-dismissible fade show position-fixed`;
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
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0" style="font-weight: 600;">🏗️ Solution Architect Performance</h5>
                    <div class="btn-group" role="group">
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('solutionArchitects', 'csv')">
                            <i class="fas fa-file-csv"></i> CSV
                        </button>
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('solutionArchitects', 'xlsx')">
                            <i class="fas fa-file-excel"></i> Excel
                        </button>
                    </div>
                </div>
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
    
    // Get minimum project count from slider (default to 2 if not set)
    const existingSlider = document.getElementById('minProjectCount');
    const currentSliderValue = existingSlider?.value || 2;
    const minProjectCount = parseInt(currentSliderValue);
    
    let customersHtml = `
        <div class="card" style="border: none; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 12px; margin-top: 1rem;">
            <div class="card-header" style="background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%); color: white; border-radius: 12px 12px 0 0; border: none;">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0" style="font-weight: 600;">🏢 Practice vs Company Customer Performance</h5>
                    <div class="btn-group" role="group">
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('customers', 'csv')">
                            <i class="fas fa-file-csv"></i> CSV
                        </button>
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('customers', 'xlsx')">
                            <i class="fas fa-file-excel"></i> Excel
                        </button>
                    </div>
                </div>
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
                        <div class="mb-3">
                            <label for="minProjectCount" class="form-label">Minimum Project Count: <span id="minProjectCountValue">${currentSliderValue}</span></label>
                            <input type="range" class="form-range" id="minProjectCount" min="1" max="20" step="1" value="${currentSliderValue}">
                        </div>
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
    
    currentAnalysis.customers.practice
        .filter(customer => parseInt(customer.projects) >= minProjectCount)
        .forEach(customer => {
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
    
    // Add slider event listener for minimum project count (only if not already added)
    const minProjectCountSlider = document.getElementById('minProjectCount');
    const minProjectCountValue = document.getElementById('minProjectCountValue');
    if (minProjectCountSlider && minProjectCountValue && !minProjectCountSlider.hasAttribute('data-listener-added')) {
        minProjectCountSlider.setAttribute('data-listener-added', 'true');
        
        minProjectCountSlider.addEventListener('input', (e) => {
            const value = e.target.value;
            document.getElementById('minProjectCountValue').textContent = value;
            // Auto-refresh the customers table
            setTimeout(() => updateCustomersTable(), 100);
        });
    }
    
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
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0" style="font-weight: 600;">📊 DAS+ (Delivery Accuracy Score Plus) Analysis</h5>
                    <div class="btn-group" role="group">
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('das', 'csv')">
                            <i class="fas fa-file-csv"></i> CSV
                        </button>
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('das', 'xlsx')">
                            <i class="fas fa-file-excel"></i> Excel
                        </button>
                    </div>
                </div>
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
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0" style="font-weight: 600;">🤝 Team Combination Analysis</h5>
                    <div class="btn-group" role="group">
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('combinations', 'csv')">
                            <i class="fas fa-file-csv"></i> CSV
                        </button>
                        <button class="btn btn-outline-light btn-sm" onclick="exportTabData('combinations', 'xlsx')">
                            <i class="fas fa-file-excel"></i> Excel
                        </button>
                    </div>
                </div>
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
                        <div class="mb-3">
                            <label for="minSAProjectCount" class="form-label">Minimum Project Count: <span id="minSAProjectCountValue">2</span></label>
                            <input type="range" class="form-range" id="minSAProjectCount" min="1" max="10" step="1" value="2">
                        </div>
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
    
    // Get minimum project count from slider
    const existingSASlider = document.getElementById('minSAProjectCount');
    const currentSASliderValue = existingSASlider?.value || 2;
    const minSAProjectCount = parseInt(currentSASliderValue);
    
    if (currentAnalysis.saCombinations) {
        currentAnalysis.saCombinations
            .filter(combo => parseInt(combo.projects) >= minSAProjectCount)
            .forEach(combo => {
                combinationsHtml += `
                    <tr class="clickable-row" data-sa="${combo.sa}" data-consultant="${combo.consultant}">
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
    
    // Add click handlers for SA-Consultant combinations
    combinationsTab.querySelectorAll('.clickable-row').forEach(row => {
        row.addEventListener('click', (e) => {
            const sa = e.currentTarget.getAttribute('data-sa');
            const consultant = e.currentTarget.getAttribute('data-consultant');
            showSAConsultantProjects(sa, consultant);
        });
    });
    
    // Add slider event listener for SA combinations
    const minSAProjectCountSlider = document.getElementById('minSAProjectCount');
    const minSAProjectCountValue = document.getElementById('minSAProjectCountValue');
    if (minSAProjectCountSlider && minSAProjectCountValue && !minSAProjectCountSlider.hasAttribute('data-listener-added')) {
        minSAProjectCountSlider.setAttribute('data-listener-added', 'true');
        
        minSAProjectCountSlider.addEventListener('input', (e) => {
            const value = e.target.value;
            document.getElementById('minSAProjectCountValue').textContent = value;
            setTimeout(() => updateCombinationsTable(), 100);
        });
    }
    
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
                        <button class="btn btn-sm btn-outline-danger dq-btn" data-consultant="${performer.name}">🚫 DQ</button>
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
    
    // Add DQ handlers
    document.querySelectorAll('.dq-btn').forEach(btn => {
        btn.addEventListener('click', handleDisqualification);
    });
}

function showSAProjects(saName, projects) {
    let projectsHtml = `
        <div class="modal fade" id="saProjectsModal" tabindex="-1">
            <div class="modal-dialog" style="max-width: 95vw;">
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
            <div class="modal-dialog" style="max-width: 95vw;">
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
    const currentSort = header.getAttribute('data-sort-direction') || 'none';
    const isAscending = currentSort === 'desc'; // Toggle: if currently desc, make asc
    
    rows.sort((a, b) => {
        const aVal = a.cells[header.cellIndex].textContent.trim();
        const bVal = b.cells[header.cellIndex].textContent.trim();
        
        const aNum = parseFloat(aVal.replace(/[^0-9.-]/g, ''));
        const bNum = parseFloat(bVal.replace(/[^0-9.-]/g, ''));
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return isAscending ? aNum - bNum : bNum - aNum;
        } else {
            return isAscending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        }
    });
    
    // Clear all headers
    table.querySelectorAll('th.sortable').forEach(h => {
        h.removeAttribute('data-sort-direction');
        const text = h.textContent.replace(/ [↑↓↕]/g, '');
        h.innerHTML = text + ' ↕';
    });
    
    // Set current header
    const headerText = header.textContent.replace(/ [↑↓↕]/g, '');
    const newDirection = isAscending ? 'asc' : 'desc';
    header.setAttribute('data-sort-direction', newDirection);
    header.innerHTML = headerText + (isAscending ? ' ↑' : ' ↓');
    
    // Reorder rows
    rows.forEach(row => tbody.appendChild(row));
}

function exportToCSV() {
    exportTabData('consultants', 'csv');
}

async function exportTabData(tab, format) {
    if (!currentAnalysis) {
        showMessage('No analysis data to export', 'error');
        return;
    }
    
    let data;
    switch (tab) {
        case 'consultants':
            data = currentAnalysis.consultants;
            break;
        case 'solutionArchitects':
            data = currentAnalysis.solutionArchitects || [];
            break;
        case 'customers':
            // For customers, we'll export practice customers by default
            data = currentAnalysis.customers?.practice || [];
            break;
        case 'das':
            data = currentAnalysis.dasAnalysis || [];
            break;
        case 'combinations':
            data = currentAnalysis.saCombinations || [];
            break;
        default:
            showMessage('Invalid tab for export', 'error');
            return;
    }
    
    try {
        showMessage(`Exporting ${tab} data as ${format.toUpperCase()}...`, 'info');
        
        const response = await fetch('/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ format, tab, data })
        });
        
        if (!response.ok) {
            throw new Error(`Export failed: ${response.statusText}`);
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = response.headers.get('Content-Disposition')?.split('filename="')[1]?.slice(0, -1) || `${tab}-export.${format}`;
        a.click();
        window.URL.revokeObjectURL(url);
        
        showMessage(`${tab} data exported successfully!`, 'success');
    } catch (error) {
        showMessage(`Export failed: ${error.message}`, 'error');
    }
}

async function exportAllData(format) {
    if (!currentAnalysis) {
        showMessage('No analysis data to export', 'error');
        return;
    }
    
    try {
        showMessage(`Exporting complete analysis as ${format.toUpperCase()}...`, 'info');
        
        const response = await fetch('/export-all', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ format, analysis: currentAnalysis })
        });
        
        if (!response.ok) {
            throw new Error(`Export failed: ${response.statusText}`);
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = response.headers.get('Content-Disposition')?.split('filename="')[1]?.slice(0, -1) || `complete-analysis.${format}`;
        a.click();
        window.URL.revokeObjectURL(url);
        
        showMessage(`Complete analysis exported successfully!`, 'success');
    } catch (error) {
        showMessage(`Export failed: ${error.message}`, 'error');
    }
}

function enableExportButtons(enabled) {
    const exportButtonIds = [
        'exportConsultantsCSV', 'exportConsultantsXLSX',
        'exportAllCSV', 'exportAllXLSX'
    ];
    
    exportButtonIds.forEach(id => {
        const btn = document.getElementById(id);
        if (btn) btn.disabled = !enabled;
    });
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
    const sqlConnection = document.getElementById('ssrsSection');
    const selectedSource = document.querySelector('input[name="dataSource"]:checked').value;
    
    if (selectedSource === 'upload') {
        if (fileUpload) fileUpload.style.display = 'block';
        if (sqlConnection) sqlConnection.style.display = 'none';
    } else {
        if (fileUpload) fileUpload.style.display = 'none';
        if (sqlConnection) sqlConnection.style.display = 'block';
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
                        updateSliderDisplay('efficiencyThreshold', t.efficiency_threshold);
                    }
                    if (t.success_threshold >= 0 && t.success_threshold <= 1) {
                        document.getElementById('successThreshold').value = t.success_threshold;
                        updateSliderDisplay('successThreshold', t.success_threshold);
                    }
                    if (t.green_threshold >= -1 && t.green_threshold <= 1) {
                        document.getElementById('greenThreshold').value = t.green_threshold;
                        updateSliderDisplay('greenThreshold', t.green_threshold);
                    }
                    if (t.yellow_threshold >= 0 && t.yellow_threshold <= 1) {
                        document.getElementById('yellowThreshold').value = t.yellow_threshold;
                        updateSliderDisplay('yellowThreshold', t.yellow_threshold);
                    }
                    if (t.red_threshold >= 0 && t.red_threshold <= 2) {
                        document.getElementById('redThreshold').value = t.red_threshold;
                        updateSliderDisplay('redThreshold', t.red_threshold);
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

function updateSliderDisplay(sliderId, value) {
    const slider = document.getElementById(sliderId);
    if (slider) {
        const displayEl = slider.nextElementSibling;
        if (displayEl && displayEl.tagName === 'SMALL') {
            displayEl.textContent = (value * 100).toFixed(0) + '%';
        }
    }
}

async function saveCurrentConfig() {
    try {
        const config = {
            thresholds: {
                efficiency_threshold: parseFloat(document.getElementById('efficiencyThreshold').value),
                success_threshold: parseFloat(document.getElementById('successThreshold').value),
                green_threshold: parseFloat(document.getElementById('greenThreshold').value),
                yellow_threshold: parseFloat(document.getElementById('yellowThreshold').value),
                red_threshold: parseFloat(document.getElementById('redThreshold').value)
            },
            trending: {
                enable_trending: document.getElementById('enableTrending').checked
            },
            project_filtering: {
                enable_date_filter: document.getElementById('enableDateFilter').checked,
                filter_type: document.querySelector('input[name="filterType"]:checked')?.value || 'days',
                days_from_today: parseInt(document.getElementById('daysFromToday').value) || 365,
                specific_date: document.getElementById('specificDate').value
            }
        };
        
        await fetch('/config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(config)
        });
        
        showMessage('Configuration saved successfully!', 'success');
    } catch (error) {
        showMessage('Error saving configuration: ' + error.message, 'error');
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

async function clearCache() {
    if (confirm('Clear exclusions cache? This will force reload from file on next analysis.')) {
        try {
            const response = await fetch('/exclusions/cache', { method: 'DELETE' });
            if (response.ok) {
                showMessage('Exclusions cache cleared successfully', 'success');
            } else {
                showMessage('Error clearing cache', 'error');
            }
        } catch (error) {
            showMessage('Error clearing cache: ' + error.message, 'error');
        }
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

async function handleDisqualification(event) {
    const consultant = event.target.getAttribute('data-consultant');
    
    if (confirm(`Disqualify ${consultant} from Consultant of the Quarter?`)) {
        try {
            const response = await fetch('/disqualify', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ consultant })
            });
            
            if (response.ok) {
                event.target.textContent = '✓ DQ\'d';
                event.target.disabled = true;
                event.target.classList.remove('btn-outline-danger');
                event.target.classList.add('btn-secondary');
                showMessage(`${consultant} disqualified from Consultant of the Quarter`, 'success');
                
                // Auto-recalculate to update rankings
                if (currentAnalysis) {
                    try {
                        const analysisResponse = await fetch('/analyze', { method: 'POST' });
                        if (analysisResponse.ok) {
                            const result = await analysisResponse.json();
                            currentAnalysis = result.analysis;
                            updatePerformanceHighlights();
                            showMessage('Rankings updated!', 'success');
                        }
                    } catch (error) {
                        showMessage('DQ saved. Refresh to see updated rankings.', 'warning');
                    }
                }
            }
        } catch (error) {
            showMessage('Error disqualifying consultant: ' + error.message, 'error');
        }
    }
}

async function viewDisqualified() {
    try {
        const response = await fetch('/disqualified');
        const disqualified = await response.json();
        
        let dqHtml = `
            <div class="modal fade" id="dqModal" tabindex="-1">
                <div class="modal-dialog modal-lg">
                    <div class="modal-content">
                        <div class="modal-header">
                            <h5 class="modal-title">🚫 Disqualified Consultants</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                        </div>
                        <div class="modal-body">
        `;
        
        if (disqualified.length === 0) {
            dqHtml += '<p class="text-muted">No consultants are currently disqualified.</p>';
        } else {
            dqHtml += `
                            <div class="table-responsive">
                                <table class="table table-striped">
                                    <thead>
                                        <tr>
                                            <th>Consultant</th>
                                            <th>Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody>
            `;
            
            disqualified.forEach(consultant => {
                dqHtml += `
                    <tr>
                        <td>${consultant}</td>
                        <td>
                            <button class="btn btn-sm btn-outline-success re-qualify-btn" data-consultant="${consultant}">✓ Re-qualify</button>
                        </td>
                    </tr>
                `;
            });
            
            dqHtml += `
                                    </tbody>
                                </table>
                            </div>
            `;
        }
        
        dqHtml += `
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        const existingModal = document.getElementById('dqModal');
        if (existingModal) existingModal.remove();
        
        document.body.insertAdjacentHTML('beforeend', dqHtml);
        const modal = new bootstrap.Modal(document.getElementById('dqModal'));
        modal.show();
        
        document.querySelectorAll('.re-qualify-btn').forEach(btn => {
            btn.addEventListener('click', reQualifyConsultant);
        });
        
    } catch (error) {
        showMessage('Error loading disqualified consultants: ' + error.message, 'error');
    }
}

async function reQualifyConsultant(event) {
    const consultant = event.target.getAttribute('data-consultant');
    
    if (confirm(`Re-qualify ${consultant} for Consultant of the Quarter?`)) {
        try {
            const response = await fetch('/disqualify', {
                method: 'DELETE',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ consultant })
            });
            
            if (response.ok) {
                event.target.closest('tr').remove();
                showMessage(`${consultant} re-qualified for Consultant of the Quarter`, 'success');
                
                // Auto-recalculate to update rankings
                if (currentAnalysis) {
                    try {
                        const analysisResponse = await fetch('/analyze', { method: 'POST' });
                        if (analysisResponse.ok) {
                            const result = await analysisResponse.json();
                            currentAnalysis = result.analysis;
                            updatePerformanceHighlights();
                            showMessage('Rankings updated!', 'success');
                        }
                    } catch (error) {
                        showMessage('Re-qualification saved. Refresh to see updated rankings.', 'warning');
                    }
                }
            }
        } catch (error) {
            showMessage('Error re-qualifying consultant: ' + error.message, 'error');
        }
    }
}

function getRandomProjects(projects, count = 2) {
    if (!projects || projects.length === 0) return [];
    
    // Separate closed and open projects
    const closedProjects = projects.filter(p => {
        const status = (p.status || '').toLowerCase();
        const completion = parseFloat(p.completion) || 0;
        return status.includes('closed') || status.includes('complete') || completion >= 0.95;
    });
    
    const openProjects = projects.filter(p => {
        const status = (p.status || '').toLowerCase();
        const completion = parseFloat(p.completion) || 0;
        return !(status.includes('closed') || status.includes('complete') || completion >= 0.95);
    });
    
    // Prioritize closed projects, then open projects
    const shuffledClosed = [...closedProjects].sort(() => 0.5 - Math.random());
    const shuffledOpen = [...openProjects].sort(() => 0.5 - Math.random());
    
    const selected = [...shuffledClosed, ...shuffledOpen].slice(0, Math.min(count, projects.length));
    return selected;
}

async function loadRandomProjects(consultantName, filteredProjects, mode = 'filtered') {
    let projectsToSample = filteredProjects;
    
    if (mode === 'all' && currentAnalysis) {
        // Get all projects for this consultant from the full dataset
        const allConsultant = currentAnalysis.consultants.find(c => c.name === consultantName);
        if (allConsultant && allConsultant.projectDetails) {
            projectsToSample = allConsultant.projectDetails;
        }
    }
    
    const randomProjects = getRandomProjects(projectsToSample, 2);
    const randomTableBody = document.getElementById('randomProjectsTable');
    
    if (randomTableBody) {
        randomTableBody.innerHTML = '';
        
        randomProjects.forEach(project => {
            const variance = parseFloat(project.variance);
            const rowClass = variance > 30 ? 'table-danger' : variance < -10 ? 'table-success' : 'table-warning';
            
            const savantUrl = `https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo=${project.jobNumber}&isPmo=true`;
            const ssrsUrl = `https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber=${project.jobNumber}`;
            
            const row = document.createElement('tr');
            row.className = rowClass;
            row.innerHTML = `
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.customer}</td>
                <td>${project.variance > 0 ? '+' : ''}${project.variance}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
                <td>
                    <a href="${savantUrl}" target="_blank" class="btn btn-xs btn-outline-primary me-1">🔗</a>
                    <a href="${ssrsUrl}" target="_blank" class="btn btn-xs btn-outline-secondary">📄</a>
                </td>
                <td>
                    <button class="btn btn-xs btn-outline-danger random-exclude-btn" data-consultant="${consultantName}" data-project="${project.jobNumber}" data-mode="${mode}">🚫</button>
                </td>
            `;
            randomTableBody.appendChild(row);
        });
        
        // Add exclude handlers for random projects
        document.querySelectorAll('.random-exclude-btn').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                await handleExclusion(e);
                // Refresh random projects after exclusion
                const mode = e.target.getAttribute('data-mode');
                loadRandomProjects(consultantName, filteredProjects, mode);
            });
        });
    }
    
    // Update button states
    document.getElementById('randomFiltered').className = mode === 'filtered' ? 'btn btn-sm btn-primary' : 'btn btn-sm btn-outline-primary';
    document.getElementById('randomAll').className = mode === 'all' ? 'btn btn-sm btn-primary' : 'btn btn-sm btn-outline-secondary';
}

function showSAConsultantProjects(sa, consultant) {
    if (!currentAnalysis) return;
    
    // Find projects for this SA-Consultant combination
    const projects = [];
    
    // Search through all data to find matching projects
    // This would need access to the raw data, so we'll use a simplified approach
    // by finding projects from both SA and consultant project details
    const saData = currentAnalysis.solutionArchitects.find(s => s.name === sa);
    const consultantData = currentAnalysis.consultants.find(c => c.name === consultant);
    
    if (saData && consultantData) {
        // Find common projects between SA and consultant
        saData.projectDetails.forEach(saProject => {
            const matchingProject = consultantData.projectDetails.find(cProject => 
                cProject.jobNumber === saProject.jobNumber
            );
            if (matchingProject) {
                projects.push(saProject);
            }
        });
    }
    
    let projectsHtml = `
        <div class="modal fade" id="saConsultantProjectsModal" tabindex="-1">
            <div class="modal-dialog" style="max-width: 95vw;">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">📋 Projects for ${sa} + ${consultant}</h5>
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
    
    const existingModal = document.getElementById('saConsultantProjectsModal');
    if (existingModal) existingModal.remove();
    
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('saConsultantProjectsModal'));
    modal.show();
}

async function exportConsultantProjects(consultantName, projects, format, consultant) {
    try {
        const data = [
            {
                'Job Number': 'SUMMARY',
                'Description': `${consultantName} - Overall Performance`,
                'Customer': `Success Rate: ${consultant?.successRate || 'N/A'}%`,
                'Budget Hours': `Efficiency Score: ${consultant?.efficiencyScore || 'N/A'}%`,
                'Actual Hours': `Total Projects: ${consultant?.projects || 'N/A'}`,
                'Variance %': `Total Hours: ${consultant?.hours || 'N/A'}`,
                'Complete %': '',
                'Status': ''
            },
            {
                'Job Number': '',
                'Description': '',
                'Customer': '',
                'Budget Hours': '',
                'Actual Hours': '',
                'Variance %': '',
                'Complete %': '',
                'Status': ''
            },
            {
                'Job Number': 'Job Number',
                'Description': 'Description',
                'Customer': 'Customer',
                'Budget Hours': 'Budget Hours',
                'Actual Hours': 'Actual Hours',
                'Variance %': 'Variance %',
                'Complete %': 'Complete %',
                'Status': 'Status'
            },
            ...projects.map(project => ({
                'Job Number': project.jobNumber,
                'Description': project.description,
                'Customer': project.customer,
                'Budget Hours': project.budgetHrs,
                'Actual Hours': project.actualHrs,
                'Variance %': project.variance,
                'Complete %': (project.completion * 100).toFixed(0),
                'Status': project.status
            }))
        ];
        
        const response = await fetch('/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ format, tab: 'consultant-projects', data, consultantName, consultant })
        });
        
        if (!response.ok) {
            throw new Error(`Export failed: ${response.statusText}`);
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${consultantName.replace(/[^a-zA-Z0-9]/g, '_')}-projects.${format}`;
        a.click();
        window.URL.revokeObjectURL(url);
        
        showMessage(`Projects exported successfully!`, 'success');
    } catch (error) {
        showMessage(`Export failed: ${error.message}`, 'error');
    }
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