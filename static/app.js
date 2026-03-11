document.addEventListener('DOMContentLoaded', () => {
    const countySelect = document.getElementById('countySelect');
    const republicanToggle = document.getElementById('republicanToggle');
    const searchBtn = document.getElementById('searchBtn');
    const searchLoader = document.getElementById('searchLoader');
    const btnText = document.querySelector('.btn-text');
    const resultsCount = document.getElementById('resultsCount');
    const tableBody = document.getElementById('tableBody');
    const downloadBtn = document.getElementById('downloadBtn');
    const repubCols = document.querySelectorAll('.repub-col');

    let currentData = [];

    // Fetch initial counties
    axios.get('/api/counties')
        .then(res => {
            if (res.data.success) {
                const counties = res.data.counties;
                counties.forEach(county => {
                    const opt = document.createElement('option');
                    opt.value = county;
                    opt.textContent = county;
                    countySelect.appendChild(opt);
                });
            }
        })
        .catch(err => console.error("Error loading counties:", err));

    // Handle search click
    searchBtn.addEventListener('click', async () => {
        // Get selected counties
        const selectedOptions = Array.from(countySelect.selectedOptions).map(opt => opt.value);
        const onlyRepublicans = republicanToggle.checked;

        // UI Loading State
        searchLoader.classList.remove('hidden');
        btnText.textContent = "Processing...";
        searchBtn.disabled = true;
        downloadBtn.disabled = true;

        if (onlyRepublicans) {
            repubCols.forEach(col => col.classList.remove('hidden'));
        } else {
            repubCols.forEach(col => col.classList.add('hidden'));
        }

        try {
            const response = await axios.post('/api/filter', {
                counties: selectedOptions,
                onlyRepublicans: onlyRepublicans
            });

            if (response.data.success) {
                currentData = response.data.data;
                renderTable(currentData, onlyRepublicans);
            } else {
                tableBody.innerHTML = `<tr><td colspan="7" class="empty-state" style="color: var(--danger)">Error: ${response.data.error}</td></tr>`;
            }
        } catch (error) {
            tableBody.innerHTML = `<tr><td colspan="7" class="empty-state" style="color: var(--danger)">Network Error: Could not reach the server.</td></tr>`;
            console.error(error);
        } finally {
            // Restore UI
            searchLoader.classList.add('hidden');
            btnText.textContent = "Search Officials";
            searchBtn.disabled = false;
        }
    });

    function renderTable(data, showRepubCols) {
        resultsCount.textContent = `${data.length} Found`;
        tableBody.innerHTML = '';

        if (data.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="7" class="empty-state">No officials found for this filter.</td></tr>`;
            downloadBtn.disabled = true;
            return;
        }

        downloadBtn.disabled = false;

        data.forEach(row => {
            const tr = document.createElement('tr');
            
            let html = `
                <td>${row.County}</td>
                <td>${row.Municipality}</td>
                <td>${row.Committee}</td>
                <td style="font-weight: 500; color: white">${row.Name}</td>
                <td>${row.Role}</td>
            `;

            if (showRepubCols) {
                html += `
                    <td>${row.Address || 'N/A'}</td>
                    <td>${row.Phone || 'N/A'}</td>
                    <td>${row.StateVoterId || 'N/A'}</td>
                `;
            } else {
                html += `
                    <td class="hidden"></td>
                    <td class="hidden"></td>
                    <td class="hidden"></td>
                `;
            }

            tr.innerHTML = html;
            tableBody.appendChild(tr);
        });
    }

    // CSV Download Functionality
    downloadBtn.addEventListener('click', () => {
        if (!currentData || currentData.length === 0) return;

        const onlyRepublicans = republicanToggle.checked;
        
        let headers = ['County', 'Municipality', 'Committee', 'Name', 'Role'];
        if (onlyRepublicans) {
            headers.push('Address', 'Phone', 'StateVoterId');
        }

        let csvContent = "data:text/csv;charset=utf-8," 
            + headers.join(",") + "\n";

        currentData.forEach(row => {
            // Basic string escaping for CSV
            const escape = (str) => `"${String(str || '').replace(/"/g, '""')}"`;
            let rowArray = [
                escape(row.County),
                escape(row.Municipality),
                escape(row.Committee),
                escape(row.Name),
                escape(row.Role)
            ];

            if (onlyRepublicans) {
                rowArray.push(escape(row.Address), escape(row.Phone), escape(row.StateVoterId));
            }

            csvContent += rowArray.join(",") + "\n";
        });

        // Trigger download using Blob to handle large files
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        
        const prefix = onlyRepublicans ? "Republican_Officials" : "All_Officials";
        link.setAttribute("download", `${prefix}_Export.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url); // Clean up memory
    });
});
