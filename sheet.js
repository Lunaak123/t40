let excelData = []; // Placeholder for Excel data
let currentSheetName = ''; // Placeholder for the current sheet name
let data = [];
let filteredData = [];

// Load the Google Sheets file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelSheet(fileUrl);
    }
});

// Function to load Excel data
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Display Sheet
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Highlight data based on selections
function highlightData() {
    const primaryColumn = document.getElementById('primary-column').value.trim().toUpperCase();
    const operationColumns = document.getElementById('operation-columns').value.trim().toUpperCase().split(',');
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10);

    const table = document.querySelector('table');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    rows.forEach((row, rowIndex) => {
        if (rowIndex < rowRangeFrom - 1 || rowIndex > rowRangeTo - 1) {
            row.style.backgroundColor = ''; // Reset color if out of range
            return;
        }

        const primaryCell = row.cells[primaryColumn.charCodeAt(0) - 65]; // Convert column letter to index
        const shouldHighlight = checkOperation(rowIndex, primaryCell, operationColumns, operation, operationType);

        if (shouldHighlight) {
            row.style.backgroundColor = '#d1e7dd'; // Highlight color
        } else {
            row.style.backgroundColor = ''; // Reset color
        }
    });
}

// Check the operation condition
function checkOperation(rowIndex, primaryCell, operationColumns, operation, operationType) {
    const primaryValue = primaryCell.textContent.trim();

    if (operationType === 'and') {
        return operationColumns.every(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65]; // Get cell for operation
            const colValue = colCell.textContent.trim();
            return operation === 'null' ? !colValue : colValue !== '';
        });
    } else if (operationType === 'or') {
        return operationColumns.some(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65]; // Get cell for operation
            const colValue = colCell.textContent.trim();
            return operation === 'null' ? !colValue : colValue !== '';
        });
    }
    return false;
}

// Apply Operation
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10);

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());
    filteredData = data.filter((row, index) => {
        // Check if the current row index is within the specified range
        if (index < rowRangeFrom - 1 || index > rowRangeTo - 1) return false;

        const isPrimaryNull = row[primaryColumn] === null || row[primaryColumn] === "";
        const columnChecks = operationColumns.map(col => operation === 'null' ? row[col] === null || row[col] === "" : row[col] !== null && row[col] !== "");

        return operationType === 'and' ? !isPrimaryNull && columnChecks.every(Boolean) : !isPrimaryNull && columnChecks.some(Boolean);
    });

    filteredData = filteredData.map(row => {
        const filteredRow = {};
        filteredRow[primaryColumn] = row[primaryColumn];
        operationColumns.forEach(col => filteredRow[col] = row[col] === null || row[col] === "" ? 'NULL' : row[col]);
        return filteredRow;
    });

    displaySheet(filteredData);
}

// Download functionality
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Confirm download button
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'downloaded_file';
    const format = document.getElementById('file-format').value;

    // Download logic based on format
    if (format === 'xlsx') {
        const ws = XLSX.utils.json_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, currentSheetName);
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Close modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Attach event listeners to row and column input fields for highlighting
document.getElementById('row-range-from').addEventListener('input', highlightData);
document.getElementById('row-range-to').addEventListener('input', highlightData);
document.getElementById('primary-column').addEventListener('input', highlightData);
document.getElementById('operation-columns').addEventListener('input', highlightData);
document.getElementById('operation-type').addEventListener('change', highlightData);
document.getElementById('operation').addEventListener('change', highlightData);

document.getElementById('apply-operation').addEventListener('click', applyOperation);
