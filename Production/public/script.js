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
    
    // SQL connection
    document.getElementById('connectSQL').addEventListener('click', connectToSQL);
    
    // Configuration changes
    document.getElementById('efficiencyThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('successThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('greenThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('yellowThreshold').addEventListener('input', updateThresholdDisplay);
    document.getElementById('redThreshold').addEventListener('input', updateThresholdDisplay);
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
            updateCustomersTable();
            updateDASTable();
            updateCombinationsTable();
            
            // Console highlights for quarterly review
            if (currentAnalysis.consultantOfQuarter?.[0]) {
                console.log('🏆 QUARTERLY PERFORMANCE HIGHLIGHTS 🏆');
                console.log('=====================================');
                console.log(`🥇 Consultant of the Quarter: ${currentAnalysis.consultantOfQuarter[0].name}`);
                console.log(`   Composite Score: ${currentAnalysis.consultantOfQuarter[0].compositeScore}`);
                console.log(`   Success Rate: ${currentAnalysis.consultantOfQuarter[0].successRate}%`);
                console.log(`   Projects: ${currentAnalysis.consultantOfQuarter[0].projects}`);
                console.log(`   Hours: ${parseFloat(currentAnalysis.consultantOfQuarter[0].hours).toLocaleString()}`);
                console.log('');
                
                const topPerformers = currentAnalysis.consultantOfQuarter.slice(0, 5);
                console.log('🌟 Top 5 Composite Performers:');
                topPerformers.forEach((c, i) => {
                    console.log(`   ${i + 1}. ${c.name} - Score: ${c.compositeScore}`);
                });
                console.log('');
                
                const needsAttention = currentAnalysis.consultants.filter(c => parseFloat(c.efficiencyScore) < 60);
                if (needsAttention.length > 0) {
                    console.log('⚠️  Consultants Needing Attention:');
                    needsAttention.forEach(c => {
                        console.log(`   • ${c.name} - ${c.efficiencyScore}% efficiency`);
                    });
                }
            }
            
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

    // Practice (filtered) metrics
    const practiceConsultants = currentAnalysis.consultants.length;
    const practiceProjects = currentAnalysis.consultants.reduce((sum, c) => sum + c.projects, 0);
    const practiceHours = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.hours), 0);
    const practiceAvgEfficiency = currentAnalysis.consultants.reduce((sum, c) => sum + parseFloat(c.efficiencyScore), 0) / currentAnalysis.consultants.length;
    
    // Company-wide metrics
    const companyProjects = currentAnalysis.totalProjects;
    const companyHours = currentAnalysis.customers.company.reduce((sum, c) => sum + parseFloat(c.totalActualHrs), 0);
    const companyAvgSuccess = currentAnalysis.customers.company.reduce((sum, c) => sum + parseFloat(c.successRate), 0) / currentAnalysis.customers.company.length;

    // Update metric cards with comparison
    document.getElementById('totalConsultants').innerHTML = `
        <div><strong>${practiceConsultants}</strong> <small style="color: rgba(255,255,255,0.8);">Practice</small></div>
        <div><small style="color: rgba(255,255,255,0.7);">${currentAnalysis.solutionArchitects.length} SAs</small></div>
    `;
    
    document.getElementById('totalProjects').innerHTML = `
        <div><strong>${practiceProjects.toLocaleString()}</strong> <small style="color: rgba(255,255,255,0.8);">Practice</small></div>
        <div><small style="color: rgba(255,255,255,0.7);">${companyProjects.toLocaleString()} Company</small></div>
    `;
    
    document.getElementById('totalHours').innerHTML = `
        <div><strong>${practiceHours.toLocaleString()}</strong> <small style="color: rgba(255,255,255,0.8);">Practice</small></div>
        <div><small style="color: rgba(255,255,255,0.7);">${companyHours.toLocaleString()} Company</small></div>
    `;
    
    document.getElementById('avgEfficiency').innerHTML = `
        <div><strong>${practiceAvgEfficiency.toFixed(1)}%</strong> <small style="color: rgba(255,255,255,0.8);">Practice</small></div>
        <div><small style="color: rgba(255,255,255,0.7);">${companyAvgSuccess.toFixed(1)}% Company</small></div>
    `;

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
    
    // Add Performance Highlights
    const consultantOfQuarter = currentAnalysis.consultantOfQuarter?.[0];
    const mostImproved = currentAnalysis.consultants.find(c => parseFloat(c.efficiencyScore) >= 75 && parseFloat(c.efficiencyScore) < 90);
    const highVolume = [...currentAnalysis.consultants].sort((a, b) => b.hours - a.hours)[0];
    const needsReview = currentAnalysis.consultants.find(c => parseFloat(c.efficiencyScore) < 60);
    
    const highlightsHtml = `
        <div class="mt-4">
            <h6>🏆 Quarterly Performance Highlights</h6>
            <div class="alert alert-info alert-sm mb-3">
                <small><strong>Note:</strong> These highlights are for fun and engagement purposes. Results may contain errors and should be verified before making business decisions. Enjoy the friendly competition! 😄</small>
            </div>
            <div class="row">
                <div class="col-md-3">
                    <div class="card bg-warning text-dark">
                        <div class="card-body">
                            <h6>🏆 Consultant of the Quarter</h6>
                            <p class="mb-0"><a href="#" class="text-dark consultant-highlight" data-consultant="${consultantOfQuarter?.name}">${consultantOfQuarter?.name || 'N/A'}</a><br><small>Score: ${consultantOfQuarter?.compositeScore || 'N/A'}</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card bg-success text-white">
                        <div class="card-body">
                            <h6>📈 Most Improved</h6>
                            <p class="mb-0"><a href="#" class="text-white consultant-highlight" data-consultant="${mostImproved?.name}">${mostImproved?.name || 'N/A'}</a><br><small>${mostImproved?.efficiencyScore || 'N/A'}% efficiency</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card bg-info text-white">
                        <div class="card-body">
                            <h6>📊 High Volume</h6>
                            <p class="mb-0"><a href="#" class="text-white consultant-highlight" data-consultant="${highVolume?.name}">${highVolume?.name || 'N/A'}</a><br><small>${highVolume ? parseFloat(highVolume.hours).toLocaleString() : 'N/A'} hrs</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card bg-danger text-white">
                        <div class="card-body">
                            <h6>⚠️ Needs Review</h6>
                            <p class="mb-0"><a href="#" class="text-white consultant-highlight" data-consultant="${needsReview?.name}">${needsReview?.name || 'None'}</a><br><small>${needsReview?.efficiencyScore || ''}${needsReview ? '% efficiency' : ''}</small></p>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="row mt-3">
                <div class="col-md-12">
                    <div class="card bg-light">
                        <div class="card-body">
                            <h6>📊 Composite Scoring Breakdown (Consultant of Quarter)</h6>
                            <div class="row">
                                <div class="col-md-3"><strong>Success Rate (40%):</strong> ${consultantOfQuarter?.successRate || 'N/A'}%</div>
                                <div class="col-md-3"><strong>Efficiency (30%):</strong> ${consultantOfQuarter?.efficiencyScore || 'N/A'}%</div>
                                <div class="col-md-3"><strong>Volume (20%):</strong> ${consultantOfQuarter?.volumeScore || 'N/A'}%</div>
                                <div class="col-md-3"><strong>Consistency (10%):</strong> ${consultantOfQuarter?.consistencyScore || 'N/A'}%</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    document.querySelector('#consultants .card-body').insertAdjacentHTML('beforeend', highlightsHtml);
    
    // Add click handlers for highlights
    document.querySelectorAll('.consultant-highlight').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const consultantName = e.target.getAttribute('data-consultant');
            if (consultantName && consultantName !== 'None') {
                const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
                if (consultant) {
                    showConsultantProjects(consultantName, consultant.projectDetails);
                }
            }
        });
    });
    
    // Add sorting functionality
    addTableSorting();
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
        document.getElementById('greenThreshold').value = config.thresholds.green_threshold || -0.1;
        document.getElementById('yellowThreshold').value = config.thresholds.yellow_threshold || 0.1;
        document.getElementById('redThreshold').value = config.thresholds.red_threshold || 0.3;
        
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
            success_threshold: parseFloat(document.getElementById('successThreshold').value),
            green_threshold: parseFloat(document.getElementById('greenThreshold').value),
            yellow_threshold: parseFloat(document.getElementById('yellowThreshold').value),
            red_threshold: parseFloat(document.getElementById('redThreshold').value)
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
    const green = document.getElementById('greenThreshold');
    const yellow = document.getElementById('yellowThreshold');
    const red = document.getElementById('redThreshold');
    
    efficiency.nextElementSibling.textContent = (efficiency.value * 100).toFixed(0) + '%';
    success.nextElementSibling.textContent = (success.value * 100).toFixed(0) + '%';
    green.nextElementSibling.textContent = (green.value * 100).toFixed(0) + '%';
    yellow.nextElementSibling.textContent = (yellow.value * 100).toFixed(0) + '%';
    red.nextElementSibling.textContent = (red.value * 100).toFixed(0) + '%';
    
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
        document.getElementById('greenThreshold').value = config.thresholds.green_threshold || -0.1;
        document.getElementById('yellowThreshold').value = config.thresholds.yellow_threshold || 0.1;
        document.getElementById('redThreshold').value = config.thresholds.red_threshold || 0.3;
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
    
    // Add sorting to modal table
    setTimeout(() => addTableSorting(), 100);
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
            <table class="table table-striped sortable-table">
                <thead>
                    <tr>
                        <th class="sortable" data-sort="name">Solution Architect ↕</th>
                        <th class="sortable" data-sort="projects">Projects ↕</th>
                        <th class="sortable" data-sort="successRate">Goal Attainment % ↕</th>
                        <th class="sortable" data-sort="budgetedHours">Budgeted Hours ↕</th>
                        <th class="sortable" data-sort="actualHours">Actual Hours ↕</th>
                        <th class="sortable" data-sort="variance">Variance % ↕</th>
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
    
    // Add Performance Highlights
    const bestSA = currentAnalysis.solutionArchitects[0];
    const needsReview = currentAnalysis.solutionArchitects.find(sa => parseFloat(sa.successRate) < 60);
    const mostActive = [...currentAnalysis.solutionArchitects].sort((a, b) => b.projects - a.projects)[0];
    
    const saHighlightsHtml = `
        <div class="mt-4">
            <h6>🏆 Quarterly SA Performance Highlights</h6>
            <div class="alert alert-info alert-sm mb-3">
                <small><strong>Note:</strong> These highlights are for fun and team engagement. Please verify data before making decisions. 😄</small>
            </div>
            <div class="row">
                <div class="col-md-4">
                    <div class="card bg-success text-white">
                        <div class="card-body">
                            <h6>📈 Best Success Rate</h6>
                            <p class="mb-0"><a href="#" class="text-white sa-highlight" data-sa="${bestSA?.name}">${bestSA?.name || 'N/A'}</a><br><small>${bestSA?.successRate || 'N/A'}% success</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-warning text-dark">
                        <div class="card-body">
                            <h6>⚠️ Needs Review</h6>
                            <p class="mb-0"><a href="#" class="text-dark sa-highlight" data-sa="${needsReview?.name}">${needsReview?.name || 'None'}</a></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-info text-white">
                        <div class="card-body">
                            <h6>📊 Most Projects</h6>
                            <p class="mb-0"><a href="#" class="text-white sa-highlight" data-sa="${mostActive?.name}">${mostActive?.name || 'N/A'}</a><br><small>${mostActive?.projects || 'N/A'} projects</small></p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    saTableBody.insertAdjacentHTML('beforeend', saHighlightsHtml);
    
    // Add click handlers for SA highlights
    document.querySelectorAll('.sa-highlight').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const saName = e.target.getAttribute('data-sa');
            if (saName && saName !== 'None') {
                const sa = currentAnalysis.solutionArchitects.find(s => s.name === saName);
                if (sa) {
                    showSAProjects(saName, sa.projectDetails);
                }
            }
        });
    });
    
    // Add sorting functionality
    addTableSorting();
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
    
    // Add sorting to modal table
    setTimeout(() => addTableSorting(), 100);
}

function updateCustomersTable() {
    if (!currentAnalysis || !currentAnalysis.customers) return;

    const customerTableBody = document.querySelector('#customers .card-body');
    
    let tableHtml = `
        <div class="mb-3">
            <div class="btn-group" role="group">
                <input type="radio" class="btn-check" name="customerView" id="practiceView" checked>
                <label class="btn btn-outline-primary" for="practiceView">👥 Practice Members Only</label>
                
                <input type="radio" class="btn-check" name="customerView" id="companyView">
                <label class="btn btn-outline-secondary" for="companyView">🏢 All Company Projects</label>
            </div>
        </div>
        
        <div id="practiceCustomers">
            <h6>📊 Customer Performance - Practice Members (${currentAnalysis.customers.practice.length} customers with 3+ projects)</h6>
            <div class="table-responsive">
                <table class="table table-striped sortable-table">
                    <thead>
                        <tr>
                            <th class="sortable" data-sort="name">Customer ↕</th>
                            <th class="sortable" data-sort="projects">Projects ↕</th>
                            <th class="sortable" data-sort="successRate">Success Rate % ↕</th>
                            <th class="sortable" data-sort="avgVariance">Avg Variance % ↕</th>
                            <th class="sortable" data-sort="totalBudgetHrs">Budget Hours ↕</th>
                            <th class="sortable" data-sort="totalActualHrs">Actual Hours ↕</th>
                            <th class="sortable" data-sort="withinBudget">Within Budget ↕</th>
                            <th class="sortable" data-sort="overBudget">Over Budget ↕</th>
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    currentAnalysis.customers.practice.forEach(customer => {
        const successRate = parseFloat(customer.successRate);
        const successThreshold = parseFloat(document.getElementById('successThreshold').value) * 100;
        const rowClass = successRate >= (successThreshold + 20) ? 'table-success' : 
                        successRate >= successThreshold ? 'table-warning' : 'table-danger';
        
        tableHtml += `
            <tr class="${rowClass}">
                <td><strong><a href="#" class="customer-link" data-customer="${customer.name}" data-view="practice">${customer.name}</a></strong></td>
                <td>${customer.projects}</td>
                <td>${customer.successRate}%</td>
                <td>${customer.avgVariance > 0 ? '+' : ''}${customer.avgVariance}%</td>
                <td>${parseFloat(customer.totalBudgetHrs).toLocaleString()}</td>
                <td>${parseFloat(customer.totalActualHrs).toLocaleString()}</td>
                <td><span class="badge bg-success">${customer.withinBudget}</span></td>
                <td><span class="badge bg-danger">${customer.overBudget}</span></td>
                <td><span class="badge bg-${customer.riskColor}">${customer.riskLevel}</span></td>
            </tr>
        `;
    });
    
    tableHtml += `
                    </tbody>
                </table>
            </div>
        </div>
        
        <div id="companyCustomers" style="display: none;">
            <h6>📊 Customer Performance - All Company Projects (${currentAnalysis.customers.company.length} customers with 3+ projects)</h6>
            <div class="table-responsive">
                <table class="table table-striped sortable-table">
                    <thead>
                        <tr>
                            <th class="sortable" data-sort="name">Customer ↕</th>
                            <th class="sortable" data-sort="projects">Projects ↕</th>
                            <th class="sortable" data-sort="successRate">Success Rate % ↕</th>
                            <th class="sortable" data-sort="avgVariance">Avg Variance % ↕</th>
                            <th class="sortable" data-sort="totalBudgetHrs">Budget Hours ↕</th>
                            <th class="sortable" data-sort="totalActualHrs">Actual Hours ↕</th>
                            <th class="sortable" data-sort="withinBudget">Within Budget ↕</th>
                            <th class="sortable" data-sort="overBudget">Over Budget ↕</th>
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    currentAnalysis.customers.company.forEach(customer => {
        const successRate = parseFloat(customer.successRate);
        const successThreshold = parseFloat(document.getElementById('successThreshold').value) * 100;
        const rowClass = successRate >= (successThreshold + 20) ? 'table-success' : 
                        successRate >= successThreshold ? 'table-warning' : 'table-danger';
        
        tableHtml += `
            <tr class="${rowClass}">
                <td><strong><a href="#" class="customer-link" data-customer="${customer.name}" data-view="company">${customer.name}</a></strong></td>
                <td>${customer.projects}</td>
                <td>${customer.successRate}%</td>
                <td>${customer.avgVariance > 0 ? '+' : ''}${customer.avgVariance}%</td>
                <td>${parseFloat(customer.totalBudgetHrs).toLocaleString()}</td>
                <td>${parseFloat(customer.totalActualHrs).toLocaleString()}</td>
                <td><span class="badge bg-success">${customer.withinBudget}</span></td>
                <td><span class="badge bg-danger">${customer.overBudget}</span></td>
                <td><span class="badge bg-${customer.riskColor}">${customer.riskLevel}</span></td>
            </tr>
        `;
    });
    
    tableHtml += `
                    </tbody>
                </table>
            </div>
        </div>
    `;
    
    customerTableBody.innerHTML = tableHtml;
    
    // Add event listeners for view toggle
    document.getElementById('practiceView').addEventListener('change', () => {
        document.getElementById('practiceCustomers').style.display = 'block';
        document.getElementById('companyCustomers').style.display = 'none';
    });
    
    document.getElementById('companyView').addEventListener('change', () => {
        document.getElementById('practiceCustomers').style.display = 'none';
        document.getElementById('companyCustomers').style.display = 'block';
    });
    
    // Add click handlers for customer names
    document.querySelectorAll('.customer-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const customerName = e.target.getAttribute('data-customer');
            const view = e.target.getAttribute('data-view');
            const customerData = view === 'practice' ? 
                currentAnalysis.customers.practice.find(c => c.name === customerName) :
                currentAnalysis.customers.company.find(c => c.name === customerName);
            if (customerData) {
                showCustomerProjects(customerName, customerData.projectDetails, view);
            }
        });
    });
    
    // Add Customer Performance Highlights
    const bestCustomer = currentAnalysis.customers.practice[0];
    const riskCustomer = currentAnalysis.customers.practice.find(c => c.riskLevel === 'High' || c.riskLevel === 'Critical');
    const biggestCustomer = [...currentAnalysis.customers.practice].sort((a, b) => b.totalBudgetHrs - a.totalBudgetHrs)[0];
    
    const customerHighlightsHtml = `
        <div class="mt-4">
            <h6>🏆 Quarterly Customer Highlights</h6>
            <div class="alert alert-info alert-sm mb-3">
                <small><strong>Note:</strong> These highlights are for engagement and may contain data anomalies. Use for discussion starters! 😄</small>
            </div>
            <div class="row">
                <div class="col-md-4">
                    <div class="card bg-success text-white">
                        <div class="card-body">
                            <h6>📈 Best Performer</h6>
                            <p class="mb-0"><a href="#" class="text-white customer-highlight" data-customer="${bestCustomer?.name}" data-view="practice">${bestCustomer?.name || 'N/A'}</a><br><small>${bestCustomer?.successRate || 'N/A'}% success</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-danger text-white">
                        <div class="card-body">
                            <h6>⚠️ High Risk</h6>
                            <p class="mb-0"><a href="#" class="text-white customer-highlight" data-customer="${riskCustomer?.name}" data-view="practice">${riskCustomer?.name || 'None'}</a><br><small>${riskCustomer?.riskLevel || ''} Risk</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-info text-white">
                        <div class="card-body">
                            <h6>📊 Largest Volume</h6>
                            <p class="mb-0"><a href="#" class="text-white customer-highlight" data-customer="${biggestCustomer?.name}" data-view="practice">${biggestCustomer?.name || 'N/A'}</a><br><small>${biggestCustomer ? parseFloat(biggestCustomer.totalBudgetHrs).toLocaleString() : 'N/A'} hrs</small></p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    customerTableBody.insertAdjacentHTML('beforeend', customerHighlightsHtml);
    
    // Add click handlers for customer highlights
    document.querySelectorAll('.customer-highlight').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const customerName = e.target.getAttribute('data-customer');
            const view = e.target.getAttribute('data-view');
            if (customerName && customerName !== 'None') {
                const customerData = currentAnalysis.customers.practice.find(c => c.name === customerName);
                if (customerData) {
                    showCustomerProjects(customerName, customerData.projectDetails, view);
                }
            }
        });
    });
    
    // Add sorting functionality
    addTableSorting();
}

function showCustomerProjects(customerName, projects, view) {
    const viewLabel = view === 'practice' ? 'Practice Members' : 'All Company';
    
    let projectsHtml = `
        <div class="modal fade" id="customerProjectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">🏢 Projects for ${customerName} (${viewLabel})</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="jobNumber">Job Number ↕</th>
                                        <th class="sortable" data-sort="description">Description ↕</th>
                                        <th class="sortable" data-sort="budgetHrs">Budget Hrs ↕</th>
                                        <th class="sortable" data-sort="actualHrs">Actual Hrs ↕</th>
                                        <th class="sortable" data-sort="variance">Variance % ↕</th>
                                        <th class="sortable" data-sort="completion">Complete % ↕</th>
                                        <th class="sortable" data-sort="status">Status ↕</th>
                                        <th class="sortable" data-sort="resources">Resources ↕</th>
                                        <th class="sortable" data-sort="solutionArchitect">Solution Architect ↕</th>
                                        <th>Links</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const variance = parseFloat(project.variance);
        const rowClass = variance > 30 ? 'table-danger' : 
                        variance < -10 ? 'table-success' : 'table-warning';
        
        const savantUrl = `https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo=${project.jobNumber}&isPmo=true`;
        const ssrsUrl = `https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber=${project.jobNumber}`;
        
        projectsHtml += `
            <tr class="${rowClass}">
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.budgetHrs.toLocaleString()}</td>
                <td>${project.actualHrs.toLocaleString()}</td>
                <td>${project.variance > 0 ? '+' : ''}${project.variance}%</td>
                <td>${(project.completion * 100).toFixed(0)}%</td>
                <td><span class="badge bg-secondary">${project.status}</span></td>
                <td><small>${project.resources}</small></td>
                <td><small>${project.solutionArchitect}</small></td>
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
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // Remove existing modal if any
    const existingModal = document.getElementById('customerProjectsModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Add modal to body and show
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('customerProjectsModal'));
    modal.show();
    
    // Add sorting to modal table
    setTimeout(() => addTableSorting(), 100);
}

function addTableSorting() {
    // Remove existing event listeners to prevent duplicates
    document.querySelectorAll('.sortable').forEach(header => {
        const newHeader = header.cloneNode(true);
        header.parentNode.replaceChild(newHeader, header);
    });
    
    document.querySelectorAll('.sortable').forEach(header => {
        header.style.cursor = 'pointer';
        header.addEventListener('click', () => {
            const table = header.closest('table');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const sortKey = header.getAttribute('data-sort');
            const isNumeric = ['projects', 'hours', 'efficiencyScore', 'successRate', 'withinBudget', 'overBudget', 'budgetedHours', 'actualHours', 'variance', 'totalBudgetHrs', 'totalActualHrs', 'avgVariance'].includes(sortKey);
            
            // Toggle sort direction
            const currentDir = header.getAttribute('data-dir');
            const newDir = currentDir === 'asc' ? 'desc' : 'asc';
            header.setAttribute('data-dir', newDir);
            
            // Clear other headers
            table.querySelectorAll('.sortable').forEach(h => {
                if (h !== header) {
                    h.removeAttribute('data-dir');
                }
            });
            
            // Update all header displays
            table.querySelectorAll('.sortable').forEach(h => {
                const baseText = h.innerHTML.replace(/ [↑↓↕]/g, '');
                if (h === header) {
                    h.innerHTML = baseText + (newDir === 'asc' ? ' ↑' : ' ↓');
                } else {
                    h.innerHTML = baseText + ' ↕';
                }
            });
            
            rows.sort((a, b) => {
                let aVal, bVal;
                
                if (sortKey === 'name') {
                    aVal = a.cells[0].textContent.trim();
                    bVal = b.cells[0].textContent.trim();
                } else {
                    const colIndex = Array.from(header.parentNode.children).indexOf(header);
                    const aText = a.cells[colIndex].textContent;
                    const bText = b.cells[colIndex].textContent;
                    
                    if (isNumeric) {
                        aVal = parseFloat(aText.replace(/[^\d.-]/g, '')) || 0;
                        bVal = parseFloat(bText.replace(/[^\d.-]/g, '')) || 0;
                    } else {
                        aVal = aText.trim();
                        bVal = bText.trim();
                    }
                }
                
                if (isNumeric) {
                    return newDir === 'asc' ? aVal - bVal : bVal - aVal;
                } else {
                    return newDir === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
                }
            });
            
            rows.forEach(row => tbody.appendChild(row));
        });
    });
}

function updateDASTable() {
    if (!currentAnalysis || !currentAnalysis.dasAnalysis) return;

    const dasTableBody = document.querySelector('#das .card-body');
    
    let tableHtml = `
        <div class="mb-3">
            <h6>📊 DAS+ Consultant Summary</h6>
            <p class="text-muted">DAS+ measures alignment between budget usage and delivery progress. Score of 1.0 = perfect alignment.</p>
        </div>
        <div class="table-responsive">
            <table class="table table-striped sortable-table">
                <thead>
                    <tr>
                        <th class="sortable" data-sort="name">Consultant ↕</th>
                        <th class="sortable" data-sort="projects">Projects ↕</th>
                        <th class="sortable" data-sort="avgDAS">Avg DAS+ ↕</th>
                        <th class="sortable" data-sort="medianDAS">Median DAS+ ↕</th>
                        <th class="sortable" data-sort="lowCount">Low (<0.75) ↕</th>
                        <th class="sortable" data-sort="highCount">High (≥0.85) ↕</th>
                        <th class="sortable" data-sort="reviewCount">Review Projects ↕</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    currentAnalysis.dasAnalysis.forEach(consultant => {
        const avgDAS = parseFloat(consultant.avgDAS);
        const rowClass = avgDAS >= 0.85 ? 'table-success' : 
                        avgDAS >= 0.75 ? 'table-warning' : 'table-danger';
        
        tableHtml += `
            <tr class="${rowClass}">
                <td><strong><a href="#" class="das-consultant-link" data-consultant="${consultant.name}">${consultant.name}</a></strong></td>
                <td>${consultant.projects}</td>
                <td>${consultant.avgDAS}</td>
                <td>${consultant.medianDAS}</td>
                <td><span class="badge bg-danger">${consultant.lowCount}</span></td>
                <td><span class="badge bg-success">${consultant.highCount}</span></td>
                <td><span class="badge bg-warning">${consultant.reviewCount}</span></td>
            </tr>
        `;
    });
    
    tableHtml += `
                </tbody>
            </table>
        </div>
        
        <div class="mt-4">
            <h6>🏆 DAS+ Performance Highlights</h6>
            <div class="alert alert-info alert-sm mb-3">
                <small><strong>Note:</strong> DAS+ highlights are for performance insights and engagement. Verify project details before decisions. 😄</small>
            </div>
            <div class="row">
                <div class="col-md-6">
                    <div class="card bg-success text-white">
                        <div class="card-body">
                            <h6>📈 Best DAS+ Performance</h6>
                            <p class="mb-0"><a href="#" class="text-white das-highlight" data-consultant="${currentAnalysis.dasAnalysis[0]?.name}">${currentAnalysis.dasAnalysis[0]?.name || 'N/A'}</a><br><small>Avg: ${currentAnalysis.dasAnalysis[0]?.avgDAS || 'N/A'}</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="card bg-warning text-dark">
                        <div class="card-body">
                            <h6>⚠️ Needs Attention</h6>
                            <p class="mb-0"><a href="#" class="text-dark das-highlight" data-consultant="${currentAnalysis.dasAnalysis.find(c => parseFloat(c.avgDAS) < 0.75)?.name}">${currentAnalysis.dasAnalysis.find(c => parseFloat(c.avgDAS) < 0.75)?.name || 'None'}</a><br><small>${currentAnalysis.dasAnalysis.find(c => parseFloat(c.avgDAS) < 0.75) ? `${currentAnalysis.dasAnalysis.find(c => parseFloat(c.avgDAS) < 0.75).lowCount} low projects` : ''}</small></p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    dasTableBody.innerHTML = tableHtml;
    
    // Add click handlers for consultant names
    document.querySelectorAll('.das-consultant-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const consultantName = e.target.getAttribute('data-consultant');
            const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
            if (consultant) {
                showDASProjects(consultantName, consultant.projectDetails);
            }
        });
    });
    
    // Add click handlers for DAS highlights
    document.querySelectorAll('.das-highlight').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const consultantName = e.target.getAttribute('data-consultant');
            if (consultantName && consultantName !== 'None') {
                const consultant = currentAnalysis.consultants.find(c => c.name === consultantName);
                if (consultant) {
                    showDASProjects(consultantName, consultant.projectDetails);
                }
            }
        });
    });
    
    // Add sorting functionality
    addTableSorting();
}

function showDASProjects(consultantName, projects) {
    let projectsHtml = `
        <div class="modal fade" id="dasProjectsModal" tabindex="-1">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">📊 DAS+ Projects for ${consultantName}</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-striped sortable-table">
                                <thead>
                                    <tr>
                                        <th class="sortable" data-sort="jobNumber">Job Number ↕</th>
                                        <th class="sortable" data-sort="description">Description ↕</th>
                                        <th class="sortable" data-sort="budgetHrs">Budget Hrs ↕</th>
                                        <th class="sortable" data-sort="actualHrs">Actual Hrs ↕</th>
                                        <th class="sortable" data-sort="completion">Complete % ↕</th>
                                        <th class="sortable" data-sort="dasScore">DAS+ Score ↕</th>
                                        <th class="sortable" data-sort="status">Status ↕</th>
                                        <th>Links</th>
                                    </tr>
                                </thead>
                                <tbody>
    `;
    
    projects.forEach(project => {
        const dasScore = calculateDASScore(project.budgetHrs, project.actualHrs, project.completion, project.status);
        const rowClass = dasScore >= 0.85 ? 'table-success' : 
                        dasScore >= 0.75 ? 'table-warning' : 'table-danger';
        
        const savantUrl = `https://savant.netsync.com/v2/pmo/projects/details/financial?jobNo=${project.jobNumber}&isPmo=true`;
        const ssrsUrl = `https://ns-hou-ssrs01.netsync.com/ReportServer/Pages/ReportViewer.aspx?/Service+Delivery/Project+Financial+Details&rs:Command=Render&JobNumber=${project.jobNumber}`;
        
        projectsHtml += `
            <tr class="${rowClass}">
                <td><strong>${project.jobNumber}</strong></td>
                <td>${project.description}</td>
                <td>${project.budgetHrs.toLocaleString()}</td>
                <td>${project.actualHrs.toLocaleString()}</td>
                <td>${(project.completion * 100).toFixed(0)}%</td>
                <td><strong>${dasScore.toFixed(3)}</strong></td>
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
                    </div>
                </div>
            </div>
        </div>
    `;
    
    const existingModal = document.getElementById('dasProjectsModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    document.body.insertAdjacentHTML('beforeend', projectsHtml);
    const modal = new bootstrap.Modal(document.getElementById('dasProjectsModal'));
    modal.show();
    
    setTimeout(() => addTableSorting(), 100);
}

function calculateDASScore(budgetHrs, actualHrs, completion, status) {
    if (!budgetHrs || budgetHrs === 0) return 0;
    
    const budgetRatio = actualHrs / budgetHrs;
    let completionRatio = completion;
    
    // Completion normalization
    if (status && (status.toLowerCase().includes('closed') || status.toLowerCase().includes('complete'))) {
        completionRatio = 1.0;
    } else if (completion >= 0.95) {
        completionRatio = 1.0;
    }
    
    const dasScore = 1 - Math.abs(budgetRatio - completionRatio);
    return Math.max(0, Math.min(1, dasScore));
}

function updateCombinationsTable() {
    if (!currentAnalysis) return;

    const combinationsTableBody = document.querySelector('#combinations .card-body');
    
    let tableHtml = `
        <div class="row">
            <div class="col-md-6">
                <h6>👥 SA-Consultant Combinations (2+ projects)</h6>
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
    
    (currentAnalysis.saCombinations || []).forEach(combo => {
        const successRate = parseFloat(combo.successRate);
        const rowClass = successRate >= 80 ? 'table-success' : 
                        successRate >= 60 ? 'table-warning' : 'table-danger';
        
        tableHtml += `
            <tr class="${rowClass}">
                <td><strong>${combo.sa}</strong></td>
                <td><strong>${combo.consultant}</strong></td>
                <td>${combo.projects}</td>
                <td>${combo.successRate}%</td>
            </tr>
        `;
    });
    
    tableHtml += `
                        </tbody>
                    </table>
                </div>
            </div>
            <div class="col-md-6">
                <h6>🏢 SA-Customer Success Rates (2+ projects)</h6>
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
    
    (currentAnalysis.saCustomerAnalysis || []).forEach(combo => {
        const successRate = parseFloat(combo.successRate);
        const rowClass = successRate >= 80 ? 'table-success' : 
                        successRate >= 60 ? 'table-warning' : 'table-danger';
        
        tableHtml += `
            <tr class="${rowClass}">
                <td><strong>${combo.sa}</strong></td>
                <td><strong>${combo.customer}</strong></td>
                <td>${combo.projects}</td>
                <td>${combo.successRate}%</td>
            </tr>
        `;
    });
    
    tableHtml += `
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <div class="mt-4">
            <div class="card bg-info text-white">
                <div class="card-body">
                    <h6>📊 Customer Risk Scoring Explanation</h6>
                    <p class="mb-2"><strong>Risk factors considered:</strong></p>
                    <ul class="mb-2">
                        <li><strong>Success Rate (40% weight):</strong> <50% = High Risk, 50-70% = Medium Risk, 70-85% = Low Risk, >85% = Minimal Risk</li>
                        <li><strong>Project Volume (20% weight):</strong> <5 projects = Higher Risk (limited data), 5-10 = Medium, >10 = Lower Risk</li>
                        <li><strong>Variance Consistency (25% weight):</strong> >50% avg variance = High Risk, 30-50% = Medium, 15-30% = Low, <15% = Minimal</li>
                        <li><strong>Budget Size (15% weight):</strong> Very large (>5000 hrs) or very small (<2000 hrs) projects carry additional risk</li>
                    </ul>
                    <p class="mb-0"><strong>Risk Levels:</strong> Low (0-2 points), Medium (3-5 points), High (6-8 points), Critical (9+ points)</p>
                </div>
            </div>
        </div>
    `;
    
    combinationsTableBody.innerHTML = tableHtml;
    
    // Add Team Combination Highlights
    const bestCombo = (currentAnalysis.saCombinations || [])[0];
    const bestSACustomer = (currentAnalysis.saCustomerAnalysis || [])[0];
    const needsAttention = (currentAnalysis.saCombinations || []).find(c => parseFloat(c.successRate) < 60);
    
    const teamHighlightsHtml = `
        <div class="mt-4">
            <h6>🏆 Quarterly Team Performance Highlights</h6>
            <div class="alert alert-info alert-sm mb-3">
                <small><strong>Note:</strong> Team highlights are for fun and motivation. Always validate findings before strategic decisions. 😄</small>
            </div>
            <div class="row">
                <div class="col-md-4">
                    <div class="card bg-success text-white">
                        <div class="card-body">
                            <h6>📈 Best SA-Consultant Team</h6>
                            <p class="mb-0">${bestCombo ? `${bestCombo.sa} + ${bestCombo.consultant}` : 'N/A'}<br><small>${bestCombo?.successRate || 'N/A'}% success</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-primary text-white">
                        <div class="card-body">
                            <h6>🏢 Best SA-Customer Match</h6>
                            <p class="mb-0">${bestSACustomer ? `${bestSACustomer.sa} + ${bestSACustomer.customer}` : 'N/A'}<br><small>${bestSACustomer?.successRate || 'N/A'}% success</small></p>
                        </div>
                    </div>
                </div>
                <div class="col-md-4">
                    <div class="card bg-warning text-dark">
                        <div class="card-body">
                            <h6>⚠️ Team Needs Review</h6>
                            <p class="mb-0">${needsAttention ? `${needsAttention.sa} + ${needsAttention.consultant}` : 'None'}<br><small>${needsAttention?.successRate || ''}${needsAttention ? '% success' : ''}</small></p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    combinationsTableBody.insertAdjacentHTML('beforeend', teamHighlightsHtml);
    
    // Add sorting functionality
    addTableSorting();
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