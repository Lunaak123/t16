let data = [];
let filteredData = [];
let subsheetNames = [];

async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        subsheetNames = workbook.SheetNames.filter(name => name !== sheetName);
        const subsheetSelect = document.getElementById('subsheet-select');
        subsheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            subsheetSelect.appendChild(option);
        });

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length > 0) {
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
            Object.values(row).forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData === null ? '' : cellData;
                tr.appendChild(td);
            });
            table.appendChild(tr);
        });

        sheetContentDiv.appendChild(table);
    }
}

function applyOperations() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumns = document.getElementById('operation-columns').value.split(',').map(col => col.trim());
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    filteredData = data.filter(row => {
        const primaryCondition = row[primaryColumn] !== null;
        const otherConditions = operationColumns.map(col => row[col] !== null);
        const finalCondition = operationType === 'and' ? otherConditions.every(Boolean) : otherConditions.some(Boolean);

        return operation === 'not-null' ? primaryCondition && finalCondition : !primaryCondition && finalCondition;
    });

    displaySheet(filteredData);
}

function handleSubsheetSelect(event) {
    const selectedSheetName = event.target.value;

    if (selectedSheetName) {
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const subsheet = workbook.Sheets[selectedSheetName];
        data = XLSX.utils.sheet_to_json(subsheet, { defval: null });
        filteredData = [...data];
        displaySheet(filteredData);
    }
}

function downloadFile() {
    const filename = document.getElementById('filename').value;
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx') {
        const ws = XLSX.utils.json_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csv], { type: 'text/csv' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${filename}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

document.getElementById('apply-operation').addEventListener('click', applyOperations);
document.getElementById('subsheet-select').addEventListener('change', handleSubsheetSelect);
document.getElementById('download-button').addEventListener('click', downloadFile);

document.addEventListener('DOMContentLoaded', () => {
    const params = new URLSearchParams(window.location.search);
    const fileUrl = params.get('fileUrl');
    if (fileUrl) loadExcelSheet(fileUrl);
});
