document.getElementById('fetch-sheets').addEventListener('click', async () => {
    const excelUrl = document.getElementById('excel-url').value;
    const loadingIndicator = document.getElementById('loading-indicator');

    if (!excelUrl) {
        alert("Please enter a valid Excel file URL.");
        return;
    }

    try {
        loadingIndicator.style.display = 'block';
        const response = await fetch(excelUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        window.location.href = `sheet.html?sheetName=${encodeURIComponent(firstSheetName)}&fileUrl=${encodeURIComponent(excelUrl)}`;
    } catch (error) {
        console.error("Error loading Excel file:", error);
        alert("Failed to load the Excel file. Please check the URL and try again.");
    } finally {
        loadingIndicator.style.display = 'none';
    }
});
