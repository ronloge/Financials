let currentAnalysis = null;
let performanceChart = null;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    loadConfiguration();
    setupEventListeners();
    updateDateRangeDisplay();
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
    
    // SQL connection
    document.getElementById('connectSQL').addEventListener('click', connectToSQL);
    
    // Configuration changes
    document.getElementById('efficiencyThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('successThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('enableDateFilter').addEventListener('change', toggleDateFilter);
    
    // Date filter options
    document.querySelectorAll('input[name="filterType"]').forEach(radio => {
        radio.addEventListener('change', updateDateRangeDisplay);
    });
    document.getElementById('daysFromToday').addEventListener('input', updateDateRangeDisplay);
    document.getElementById('specificDate').addEventListener('change', updateDateRangeDisplay);
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('runAnalysis').disabled = false;
        showMessage('File selected: ' + file.name, 'success');
    }
}

async function runAnalysis() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    
    if (!file) {
        showMessage('Please select an Excel file first', 'error');
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', file);
    
    // Add filter files if selected
    const engineersFile = document.getElementById('engineersFile').files[0];
    const excludeFile = document.getElementById('excludeFile').files[0];
    const saFile = document.getElementById('saFile').files[0];
    
    if (engineersFile) formData.append('engineersFile', engineersFile);
    if (excludeFile) formData.append('excludeFile', excludeFile);
    if (saFile) formData.append('saFile', saFile);

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

function updateDashboard() {
    if (!currentAnalysis) return;

    // Update metric cards
    document.getElementById('totalConsultants').textContent = currentAnalysis.consultants.length;
    document.getElementById('totalProjects').textContent = currentAnalysis.totalProjects;
    
    const totalHours = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.hours), 0);
    document.getElementById('totalHours').textContent = totalHours.toLocaleString();
    
    const avgEfficiency = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.efficiencyScore), 0) / currentAnalysis.consultants.length;
    document.getElementById('avgEfficiency').textContent = avgEfficiency.toFixed(1) + '%';

    // Update chart
    updatePerformanceChart();
}

function updateConsultantsTable() {
    if (!currentAnalysis) return;

    const tbody = document.querySelector('#consultantsTable tbody');
    tbody.innerHTML = '';

    currentAnalysis.consultants.forEach(consultant => {
        const row = document.createElement('tr');
        
        // Add performance class based on efficiency
        const efficiency = parseFloat(consultant.efficiencyScore);
        if (efficiency >= 80) {
            row.classList.add('performance-good');
        } else if (efficiency >= 60) {
            row.classList.add('performance-average');
        } else {
            row.classList.add('performance-poor');
        }

        row.innerHTML = `
            <td><strong><a href="#" class="consultant-link" data-consultant="${consultant.name}">${consultant.name}</a></strong></td>
            <td>${consultant.projects}</td>
            <td>${parseFloat(consultant.hours).toLocaleString()}</td>
            <td>${consultant.efficiencyScore}%</td>
            <td>${consultant.successRate}%</td>
            <td><span class="badge bg-success">${consultant.withinBudget}</span></td>
            <td><span class="badge bg-danger">${consultant.overBudget}</span></td>
        `;
        
        // Add click handler for consultant name
        const consultantLink = row.querySelector('.consultant-link');
        consultantLink.addEventListener('click', (e) => {
            e.preventDefault();
            showConsultantProjects(consultant.name, consultant.projectDetails);
        });
        
        tbody.appendChild(row);
    });
}

function updatePerformanceChart() {
    if (!currentAnalysis) return;

    const ctx = document.getElementById('performanceChart').getContext('2d');
    
    if (performanceChart) {
        performanceChart.destroy();
    }

    const top10Consultants = currentAnalysis.consultants.slice(0, 10);
    
    performanceChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: top10Consultants.map(c => c.name),
            datasets: [{
                label: 'Efficiency Score (%)',
                data: top10Consultants.map(c => parseFloat(c.efficiencyScore)),
                backgroundColor: 'rgba(54, 162, 235, 0.6)',
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }, {
                label: 'Goal Attainment (%)',
                data: top10Consultants.map(c => parseFloat(c.successRate)),
                backgroundColor: 'rgba(255, 99, 132, 0.6)',
                borderColor: 'rgba(255, 99, 132, 1)',
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

async function loadConfiguration() {
    try {
        const response = await fetch('/config');
        const config = await response.json();
        
        // Update UI with config values
        document.getElementById('efficiencyThreshold').value = config.thresholds.efficiency_threshold;
        document.getElementById('successThreshold').value = config.thresholds.success_threshold;
        
        if (config.project_filtering) {
            document.getElementById('enableDateFilter').checked = config.project_filtering.enable_date_filter;
            document.getElementById('daysFromToday').value = config.project_filtering.days_from_today;
            document.getElementById('specificDate').value = config.project_filtering.specific_date;
            
            if (config.project_filtering.filter_type === 'date') {
                document.getElementById('filterDate').checked = true;
            }
        }
        
        updateThresholdDisplay();
        toggleDateFilter();
        updateDateRangeDisplay();
    } catch (error) {
        console.error('Error loading configuration:', error);
    }
}

async function saveConfiguration() {
    const config = {
        thresholds: {
            efficiency_threshold: parseFloat(document.getElementById('efficiencyThreshold').value),
            success_threshold: parseFloat(document.getElementById('successThreshold').value)
        },
        project_filtering: {
            enable_date_filter: document.getElementById('enableDateFilter').checked,
            filter_type: document.querySelector('input[name="filterType"]:checked').value,
            days_from_today: parseInt(document.getElementById('daysFromToday').value),
            specific_date: document.getElementById('specificDate').value
        }
    };

    try {
        const response = await fetch('/config', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(config)
        });

        const result = await response.json();
        if (result.success) {
            showMessage('Configuration saved', 'success');
        }
    } catch (error) {
        showMessage('Error saving configuration: ' + error.message, 'error');
    }
}

function updateThresholdDisplay() {
    const efficiency = document.getElementById('efficiencyThreshold');
    const success = document.getElementById('successThreshold');
    
    efficiency.nextElementSibling.textContent = (efficiency.value * 100).toFixed(0) + '%';
    success.nextElementSibling.textContent = (success.value * 100).toFixed(0) + '%';
    
    saveConfiguration();
}

function toggleDateFilter() {
    const enabled = document.getElementById('enableDateFilter').checked;
    const options = document.getElementById('dateFilterOptions');
    options.style.display = enabled ? 'block' : 'none';
    updateDateRangeDisplay();
}

function updateDateRangeDisplay() {
    const enabled = document.getElementById('enableDateFilter').checked;
    const filterType = document.querySelector('input[name="filterType"]:checked')?.value;
    const infoElement = document.getElementById('dateRangeInfo');
    
    if (!enabled) {
        infoElement.textContent = 'All projects included (no date filtering)';
        return;
    }
    
    if (filterType === 'days') {
        const days = document.getElementById('daysFromToday').value;
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - days);
        infoElement.textContent = `Projects from ${cutoffDate.toISOString().split('T')[0]} onwards (${days} days)`;
    } else {
        const specificDate = document.getElementById('specificDate').value;
        infoElement.textContent = `Projects from ${specificDate} onwards`;
    }
    
    saveConfiguration();
}

function showMessage(message, type) {
    // Create toast notification
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
    
    // Auto remove after 5 seconds
    setTimeout(() => {
        if (toast.parentNode) {
            toast.parentNode.removeChild(toast);
        }
    }, 5000);
}

function toggleDataSource() {
    const dataSource = document.querySelector('input[name="dataSource"]:checked').value;
    const fileSection = document.getElementById('fileUploadSection');
    const ssrsSection = document.getElementById('ssrsSection');
    
    if (dataSource === 'upload') {
        fileSection.style.display = 'block';
        ssrsSection.style.display = 'none';
    } else {
        fileSection.style.display = 'none';
        ssrsSection.style.display = 'block';
    }
}

function handleConfigUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const config = JSON.parse(e.target.result);
                applyConfigToUI(config);
                showMessage('Config file loaded successfully', 'success');
            } catch (error) {
                showMessage('Error parsing config file: ' + error.message, 'error');
            }
        };
        reader.readAsText(file);
    }
}

function applyConfigToUI(config) {
    if (config.thresholds) {
        document.getElementById('efficiencyThreshold').value = config.thresholds.efficiency_threshold || 0.15;
        document.getElementById('successThreshold').value = config.thresholds.success_threshold || 0.3;
    }
    
    if (config.project_filtering) {
        document.getElementById('enableDateFilter').checked = config.project_filtering.enable_date_filter || false;
        document.getElementById('daysFromToday').value = config.project_filtering.days_from_today || 365;
        document.getElementById('specificDate').value = config.project_filtering.specific_date || '';
        
        if (config.project_filtering.filter_type === 'date') {
            document.getElementById('filterDate').checked = true;
        } else {
            document.getElementById('filterDays').checked = true;
        }
    }
    
    updateThresholdDisplay();
    toggleDateFilter();
    updateDateRangeDisplay();
}

async function connectToSQL() {
    const server = document.getElementById('sqlServer').value;
    const database = document.getElementById('sqlDatabase').value;
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    const authType = document.querySelector('input[name="authType"]:checked').value;
    
    if (!username || !password) {
        showMessage('Please enter username and password', 'error');
        return;
    }
    
    try {
        showMessage('Connecting to SQL Server...', 'info');
        // This would need server-side implementation
        showMessage('SQL Server connection not yet implemented in Node.js version', 'error');
    } catch (error) {
        showMessage('Connection failed: ' + error.message, 'error');
    }
}

function showConsultantProjects(consultantName, projects) {
    // Create modal or expand section to show projects
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
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Job Number</th>
                                        <th>Description</th>
                                        <th>Customer</th>
                                        <th>Budget Hrs</th>
                                        <th>Actual Hrs</th>
                                        <th>Variance %</th>
                                        <th>Complete %</th>
                                        <th>Status</th>
                                        <th>Links</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const greenThreshold = (parseFloat(document.getElementById('greenThreshold')?.value) || -0.1) * 100;
        const redThreshold = (parseFloat(document.getElementById('redThreshold')?.value) || 0.3) * 100;
        
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
                <td>${project.completion}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
                <td>
                    <a href="${savantUrl}" target="_blank" class="btn btn-sm btn-outline-primary me-1">🔗 Savant</a>
                    <a href="${ssrsUrl}" target="_blank" class="btn btn-sm btn-outline-secondary">📄 SSRS</a>
                </td>
            </tr>
        `;
    });
    
    projectsHtml += `
                                </tbody>
                            </table>
                        </div>
                        
                        <div class="mt-4">
                            <h6>🎯 Random Projects for Review</h6>
                            <div id="reviewProjects">
                                ${getRandomReviewProjects(projects)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('projectsModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body and show
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('projectsModal'));
    modal.show();
}

function getRandomReviewProjects(projects) {
    if (projects.length < 2) {
        return '<p class="text-muted">Not enough projects for review selection</p>';
    }
    
    // Categorize projects by performance
    const goodProjects = projects.filter(p => parseFloat(p.variance) <= 10);
    const poorProjects = projects.filter(p => parseFloat(p.variance) > 30);
    const averageProjects = projects.filter(p => parseFloat(p.variance) > 10 && parseFloat(p.variance) <= 30);
    
    let selectedProjects = [];
    
    // Try to get one good and one poor project
    if (goodProjects.length > 0 && poorProjects.length > 0) {
        selectedProjects.push(goodProjects[Math.floor(Math.random() * goodProjects.length)]);
        selectedProjects.push(poorProjects[Math.floor(Math.random() * poorProjects.length)]);
    } else if (projects.length >= 2) {
        // Fallback to any 2 random projects
        const shuffled = [...projects].sort(() => 0.5 - Math.random());
        selectedProjects = shuffled.slice(0, 2);
    }
    
    let reviewHtml = '';
    selectedProjects.forEach(project => {
        const varianceClass = parseFloat(project.variance) > 30 ? 'text-danger' : 
                             parseFloat(project.variance) < -10 ? 'text-success' : 'text-warning';
        const status = parseFloat(project.variance) <= 10 ? '🟢 Good' : 
                      parseFloat(project.variance) > 30 ? '🔴 Poor' : '🟡 Average';
        
        const savantUrl = `https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo=${project.jobNumber}&isPmo=true`;
        const ssrsUrl = `https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber=${project.jobNumber}`;
        
        reviewHtml += `
            <div class="card mb-2">
                <div class="card-body">
                    <div class="row align-items-center">
                        <div class="col-md-8">
                            <strong>${project.jobNumber}</strong> - ${project.description} 
                            <span class="${varianceClass}">(${status}, Variance: ${project.variance > 0 ? '+' : ''}${project.variance}%)</span>
                        </div>
                        <div class="col-md-4 text-end">
                            <a href="${savantUrl}" target="_blank" class="btn btn-sm btn-outline-primary me-1">🔗 Savant</a>
                            <a href="${ssrsUrl}" target="_blank" class="btn btn-sm btn-outline-secondary">📄 SSRS</a>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    return reviewHtml;
}

function updateSolutionArchitectsTable() {
    if (!currentAnalysis || !currentAnalysis.solutionArchitects) return;

    const saTableBody = document.querySelector('#solutionArchitects .card-body');
    
    let tableHtml = `
        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <th>Solution Architect</th>
                        <th>Projects</th>
                        <th>Goal Attainment %</th>
                        <th>Budgeted Hours</th>
                        <th>Actual Hours</th>
                        <th>Variance %</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    currentAnalysis.solutionArchitects.forEach(sa => {
        const varianceClass = parseFloat(sa.variance) > 30 ? 'text-danger' : 
                             parseFloat(sa.variance) < -10 ? 'text-success' : 'text-warning';
        
        tableHtml += `
            <tr>
                <td><strong><a href="#" class="sa-link" data-sa="${sa.name}">${sa.name}</a></strong></td>
                <td>${sa.projects}</td>
                <td>${sa.successRate}%</td>
                <td>${parseFloat(sa.budgetedHours).toLocaleString()}</td>
                <td>${parseFloat(sa.actualHours).toLocaleString()}</td>
                <td class="${varianceClass}">${sa.variance > 0 ? '+' : ''}${sa.variance}%</td>
            </tr>
        `;
    });
    
    tableHtml += `
                </tbody>
            </table>
        </div>
    `;
    
    saTableBody.innerHTML = tableHtml;
    
    // Add click handlers for SA names
    document.querySelectorAll('.sa-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const saName = e.target.getAttribute('data-sa');
            const sa = currentAnalysis.solutionArchitects.find(s => s.name === saName);
            if (sa) {
                showSAProjects(saName, sa.projectDetails);
            }
        });
    });
}

function showSAProjects(saName, projects) {
    // Reuse the same modal structure as consultant projects
    let projectsHtml = `
        <div class="modal fade" id="saProjectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">🏢 Projects for ${saName} (Solution Architect)</h5>
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
                                        <th>Complete %</th>
                                        <th>Status</th>
                                        <th>Links</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const greenThreshold = (parseFloat(document.getElementById('greenThreshold')?.value) || -0.1) * 100;
        const redThreshold = (parseFloat(document.getElementById('redThreshold')?.value) || 0.3) * 100;
        
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
                <td>${project.completion}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
                <td>
                    <a href="${savantUrl}" target="_blank" class="btn btn-sm btn-outline-primary me-1">🔗 Savant</a>
                    <a href="${ssrsUrl}" target="_blank" class="btn btn-sm btn-outline-secondary">📄 SSRS</a>
                </td>
            </tr>
        `;
    });
    
    projectsHtml += `
                                </tbody>
                            </table>
                        </div>
                        
                        <div class="mt-4">
                            <h6>🎯 Random Projects for Review</h6>
                            <div id="saReviewProjects">
                                ${getRandomReviewProjects(projects)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('saProjectsModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body and show
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('saProjectsModal'));
    modal.show();
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