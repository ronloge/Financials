// Add this to the end of showConsultantProjects function before the closing brace
setTimeout(() => {
    addTableSorting();
    document.querySelectorAll('#projectsModal .exclude-btn').forEach(btn => {
        btn.addEventListener('click', handleExclusion);
    });
}, 100);

// Replace the project row HTML in showConsultantProjects with this:
/*
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
*/