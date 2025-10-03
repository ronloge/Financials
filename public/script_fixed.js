let currentAnalysis = null;
let performanceChart = null;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    loadConfiguration();
    setupEventListeners();
    updateDateRangeDisplay();
    
    // Configuration panel toggle
    document.getElementById('configPanel').addEventListener('shown.bs.collapse', () => {
        document.getElementById('configToggle').className = 'fas fa-chevron-up float-end';
    });
    document.getElementById('configPanel').addEventListener('hidden.bs.collapse', () => {
        document.getElementById('configToggle').className = 'fas fa-chevron-down float-end';
    });
});

function setupEventListeners() {
    // Data source selection
    document.querySelectorAll('input[name="dataSource"]').forEach(radio => {
        radio.addEventListener('change', toggleDataSource);
    });
    
    // File uploads
    document.getElementById('excelFile').addEventListener('change', handleFileUpload);
    document.getElementById('configFile').addEventListener('change', handleConfigUpload);
    
    // Run analysis button
    document.getElementById('runAnalysis').addEventListener('click', runAnalysis);
    
    // View exclusions button
    document.getElementById('viewExclusions').addEventListener('click', viewExclusions);
    
    // SQL connection
    document.getElementById('connectSQL').addEventListener('click', connectToSQL);
    
    // Configuration changes
    document.getElementById('efficiencyThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('successThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('greenThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('yellowThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('redThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('enableDateFilter').addEventListener('change', toggleDateFilter);
    document.getElementById('enableTrending').addEventListener('change', toggleTrending);
    
    // Date filter options
    document.querySelectorAll('input[name="filterType"]').forEach(radio => {
        radio.addEventListener('change', updateDateRangeDisplay);
    });
    document.getElementById('daysFromToday').addEventListener('input', updateDateRangeDisplay);
    document.getElementById('specificDate').addEventListener('change', updateDateRangeDisplay);
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

        const result = await response.json();
        
        if (result.success) {
            currentAnalysis = result.analysis;
            updateDashboard();
            updateConsultantsTable();
            updateSolutionArchitectsTable();
            updateCustomersTable();
            updateDASTable();
            updateCombinationsTable();
            
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