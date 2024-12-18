document.addEventListener('DOMContentLoaded', () => {
    const dataTable = document.getElementById('dataTable').getElementsByTagName('tbody')[0];
    const form = document.getElementById('dataForm');

    // Load rows from local storage on page load
    loadRowsFromStorage();

    // Add row to the table
    document.getElementById('addRow').addEventListener('click', () => {
        const formData = new FormData(form);
        const row = {};

        // Collect data from the form
        formData.forEach((value, key) => {
            row[key] = value;
        });

        addRowToTable(row);
        saveRowToStorage(row);
        form.reset(); // Clear the form fields
    });

    // Export table data to Excel
    document.getElementById('exportExcel').addEventListener('click', () => {
        const tableData = [];
        const headers = Array.from(dataTable.closest('table').rows[0].cells).map(th => th.innerText);

        tableData.push(headers);

        Array.from(dataTable.rows).forEach(row => {
            const rowData = Array.from(row.cells).map(td => td.innerText);
            tableData.push(rowData);
        });

        const worksheet = XLSX.utils.aoa_to_sheet(tableData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, 'NorthEdgeSportsAcademy.xlsx');
    });

    // Clear table data
    document.getElementById('clearTable').addEventListener('click', () => {
        localStorage.removeItem('tableData');
        dataTable.innerHTML = '';
    });

    // Add row to the table
    function addRowToTable(rowData) {
        const row = dataTable.insertRow();
        for (const key in rowData) {
            const cell = row.insertCell();
            cell.innerText = rowData[key];
        }
    }

    // Save row to local storage
    function saveRowToStorage(rowData) {
        const rows = JSON.parse(localStorage.getItem('tableData')) || [];
        rows.push(rowData);
        localStorage.setItem('tableData', JSON.stringify(rows));
    }

    // Load rows from local storage
    function loadRowsFromStorage() {
        const rows = JSON.parse(localStorage.getItem('tableData')) || [];
        rows.forEach(addRowToTable);
    }
});
