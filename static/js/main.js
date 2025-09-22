document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const shopSelect = document.getElementById('shopSelect');
    const resultsContainer = document.getElementById('resultsContainer');
    const noResults = document.getElementById('noResults');
    const exportBtn = document.getElementById('exportBtn');
    const alertContainer = document.getElementById('alertContainer');
    const searchInput = document.getElementById('searchInput');
    const statusFilter = document.getElementById('statusFilter');
    let currentData = [];

    uploadForm.addEventListener('submit', async function(e) {
        e.preventDefault();
        const formData = new FormData(uploadForm);
        const uploadText = document.getElementById('uploadText');
        const uploadSpinner = document.getElementById('uploadSpinner');
        if (!shopSelect.value) {
            showAlert('warning', 'Select a shop first!');
            return;
        }
        uploadText.classList.add('d-none');
        uploadSpinner.classList.remove('d-none');
        try {
            const response = await fetch('/upload', { method: 'POST', body: formData });
            const result = await response.json();
            if (response.ok) {
                currentData = result.data;
                displayResults(currentData);
                showAlert('success', result.message);
                exportBtn.classList.remove('d-none');
            } else {
                showAlert('danger', result.error);
            }
        } catch (error) {
            showAlert('danger', 'Upload failed: ' + error.message);
        } finally {
            uploadText.classList.remove('d-none');
            uploadSpinner.classList.add('d-none');
        }
    });

    function displayResults(data) {
        const tbody = document.querySelector('#resultsTable tbody');
        tbody.innerHTML = '';
        data.forEach(item => {
            const row = createTableRow(item);
            tbody.appendChild(row);
        });
        resultsContainer.classList.remove('d-none');
        noResults.classList.add('d-none');
    }

    function createTableRow(item) {
        const row = document.createElement('tr');
        if (parseFloat(item['Stock']) <= 0) {
            row.classList.add('table-zero-stock');
        }
        row.innerHTML = `
            <td>${item['Item Name']}</td>
            <td>${item['Sales']}</td>
            <td>${parseFloat(item['Stock']) <=0 ? '<span class="zero-stock-badge">0</span>' : item['Stock']}</td>
            <td class="command-text">${item['Command']}</td>
            <td>${item['Order From Location']}</td>
            <td>${item['Available Qty']}</td>
        `;
        return row;
    }

function filterResults() {
    const search = searchInput.value.toLowerCase();
    const filter = statusFilter.value;
    let filtered = currentData.filter(item =>
        item['Item Name'].toLowerCase().includes(search)
    );
    
    if (filter && filter !== '') {
        if (filter === 'All Orders') {
            // Show all items regardless of order need
            // No extra filter applied here
        } else if (filter === 'No Orders') {
            filtered = filtered.filter(item =>
                item['Command'] === 0 || item['Order From Location'] === 'Not Available'
            );
        } else {
            // For other filters, show only items which need ordering (Command > 0)
            filtered = filtered.filter(item =>
                (item['Command'] !== 0 && item['Command'] !== '0') &&
                item['Order From Location'] !== 'Not Available'
            );
            if (filter === 'From Warehouse') {
                filtered = filtered.filter(item =>
                    item['Order From Location'].toLowerCase().includes('warehouse')
                );
            } else if (filter === 'From Shops') {
                filtered = filtered.filter(item =>
                    item['Order From Location'].toLowerCase().includes('shop') &&
                    !item['Order From Location'].toLowerCase().includes('warehouse')
                );
            } else if (filter === 'Insufficient Stock') {
                filtered = filtered.filter(item =>
                    item['Order From Location'].toLowerCase().includes('insufficient')
                );
            } else if (filter === 'Zero Stock') {
                filtered = filtered.filter(item =>
                    parseFloat(item['Stock']) <= 0
                );
            }
        }
    }
    
    displayResults(filtered);
    showFilterInfo(filtered.length, currentData.length, filter || 'All Orders');
}

    searchInput.addEventListener('input', filterResults);
    statusFilter.addEventListener('change', filterResults);

    function showAlert(type, message) {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type}`;
        alertDiv.textContent = message;
        alertContainer.innerHTML = '';
        alertContainer.appendChild(alertDiv);
        setTimeout(() => {
            alertDiv.remove();
        }, 4000);
    }

    function showFilterInfo(count, total, filterName) {
        let infoDiv = document.querySelector('.filter-info');
        if (!infoDiv) {
            infoDiv = document.createElement('div');
            infoDiv.className = 'filter-info alert alert-info';
            document.querySelector('#resultsContainer').insertBefore(infoDiv, document.querySelector('.table-responsive'));
        }
        infoDiv.textContent = `Showing ${count} of ${total} items (${filterName})`;
    }

    exportBtn.addEventListener('click', function() {
        const rows = [];
        document.querySelectorAll('#resultsTable tbody tr').forEach(tr => {
            const cells = tr.querySelectorAll('td');
            if (cells.length >= 6) {
                rows.push({
                    'Item Name': cells[0].innerText,
                    'Sales': cells[1].innerText,
                    'Stock': cells[2].innerText.replace(/\D/g, ''), // remove any badges
                    'Command': cells[3].innerText,
                    'Order From Location': cells[4].innerText,
                    'Available Qty': cells[5].innerText,
                });
            }
        });
        fetch('/export', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(rows)
        })
        .then(response => response.blob())
        .then(blob => {
            const a = document.createElement('a');
            a.href = window.URL.createObjectURL(blob);
            a.download = 'filtered_orders.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
        });
    });
});
