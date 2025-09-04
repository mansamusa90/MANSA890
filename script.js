document.addEventListener('DOMContentLoaded', () => {
    const fileUpload = document.getElementById('file-upload');
    const sheetSelect = document.getElementById('sheet-select');
    const columnSelect = document.getElementById('column-select');
    const dataPreviewSection = document.getElementById('data-preview-section');
    const dataPreviewTable = document.getElementById('data-preview-table');
    const analysisConfigSection = document.getElementById('analysis-config-section');
    const reportGenerationSection = document.getElementById('report-generation-section');
    const generateReportBtn = document.getElementById('generate-report-btn');
    const reportOutput = document.getElementById('report-output');
    const dashboard = document.getElementById('dashboard');

    let workbook;
    let selectedSheet;
    let selectedData;

    fileUpload.addEventListener('change', handleFileUpload);
    sheetSelect.addEventListener('change', handleSheetSelection);
    columnSelect.addEventListener('change', handleColumnSelection);
    generateReportBtn.addEventListener('click', generateReport);

    function handleFileUpload(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' });
                    populateSheetSelect();
                    dataPreviewSection.classList.remove('hidden');
                    analysisConfigSection.classList.remove('hidden');
                    reportGenerationSection.classList.remove('hidden');
                } catch (error) {
                    alert('Error reading or parsing the Excel file. Please ensure it is a valid .xls or .xlsx file.');
                    console.error(error);
                }
            };
            reader.onerror = () => {
                alert('Error reading the file.');
            };
            reader.readAsArrayBuffer(file);
        }
    }

    function populateSheetSelect() {
        sheetSelect.innerHTML = '<option value="">Select a sheet</option>';
        workbook.SheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            sheetSelect.appendChild(option);
        });
    }

    function handleSheetSelection() {
        const sheetName = sheetSelect.value;
        if (sheetName) {
            selectedSheet = workbook.Sheets[sheetName];
            selectedData = XLSX.utils.sheet_to_json(selectedSheet, { header: 1 });
            populateColumnSelect();
            previewData();
        } else {
            columnSelect.innerHTML = '';
            dataPreviewTable.innerHTML = '';
        }
    }

    function populateColumnSelect() {
        columnSelect.innerHTML = '';
        if (selectedData && selectedData.length > 0) {
            const headers = selectedData[0];
            headers.forEach((header, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = header;
                columnSelect.appendChild(option);
            });
        }
    }

    function handleColumnSelection() {
        previewData();
    }

    function previewData() {
        if (!selectedData) return;

        const selectedColumnIndexes = Array.from(columnSelect.selectedOptions).map(opt => parseInt(opt.value, 10));
        let dataToPreview = selectedData;

        if (selectedColumnIndexes.length > 0) {
            dataToPreview = selectedData.map(row =>
                selectedColumnIndexes.map(index => row[index])
            );
        }

        if (dataToPreview.length === 0) {
            dataPreviewTable.innerHTML = '<p>No data to display.</p>';
            return;
        }

        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        const headerRow = document.createElement('tr');

        const headers = (selectedColumnIndexes.length > 0)
            ? selectedColumnIndexes.map(index => selectedData[0][index])
            : selectedData[0];

        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        const rows = (dataToPreview.length > 1) ? dataToPreview.slice(1) : [];
        rows.forEach(rowData => {
            const tr = document.createElement('tr');
            rowData.forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        dataPreviewTable.innerHTML = '';
        dataPreviewTable.appendChild(table);
    }

    function generateReport() {
        const selectedColumnIndexes = Array.from(columnSelect.selectedOptions).map(opt => parseInt(opt.value, 10));
        if (selectedColumnIndexes.length < 2) {
            alert("Please select at least two columns to generate a report. The first selected column is used for labels, and the subsequent columns are for data.");
            return;
        }

        reportOutput.classList.remove('hidden');
        const desiredGraphs = Array.from(document.getElementById('desired-graphs').selectedOptions).map(opt => opt.value);

        document.getElementById('graph1').innerHTML = '';
        document.getElementById('graph2').innerHTML = '';

        if (desiredGraphs.length > 0) {
            createChart('graph1', desiredGraphs[0]);
        }
        if (desiredGraphs.length > 1) {
            createChart('graph2', desiredGraphs[1]);
        }

        document.querySelector('#summaryInsights p').textContent = '';
    }

    function createChart(canvasId, chartType) {
        if (!selectedData || selectedData.length < 2) {
            alert("Not enough data to generate a chart.");
            return;
        }

        const canvas = document.createElement('canvas');
        document.getElementById(canvasId).appendChild(canvas);
        const ctx = canvas.getContext('2d');
        const selectedColumnIndexes = Array.from(columnSelect.selectedOptions).map(opt => parseInt(opt.value, 10));

        const labels = selectedData.slice(1).map(row => row[selectedColumnIndexes[0]]);
        const datasets = [];

        for (let i = 1; i < selectedColumnIndexes.length; i++) {
            const data = selectedData.slice(1).map(row => {
                const val = parseFloat(row[selectedColumnIndexes[i]]);
                return isNaN(val) ? null : val;
            });

            if (data.some(d => d === null)) {
                alert(`Warning: Non-numeric data found in column "${selectedData[0][selectedColumnIndexes[i]]}". Non-numeric values will be ignored.`);
            }

            datasets.push({
                label: selectedData[0][selectedColumnIndexes[i]],
                data: data.filter(d => d !== null),
                backgroundColor: `rgba(${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, 0.5)`,
                borderColor: `rgba(${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, 1)`,
                borderWidth: 1
            });
        }

        let finalChartType = chartType.toLowerCase().replace(' ', '');
        if (finalChartType === 'barchart') finalChartType = 'bar';
        if (finalChartType === 'linechart') finalChartType = 'line';
        if (finalChartType === 'piechart') finalChartType = 'pie';
        if (finalChartType === 'scatterplot') finalChartType = 'scatter';
        if (finalChartType === 'boxplot') finalChartType = 'boxplot';
        if (finalChartType === 'histogram') {
            finalChartType = 'bar'; // Use bar chart for histogram
        }

        new Chart(ctx, {
            type: finalChartType,
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: `${chartType} of ${document.getElementById('report-title').value || 'Data'}`
                    }
                }
            }
        });
    }

    function getPreviewData() {
        if (!selectedData) return null;

        const selectedColumnIndexes = Array.from(columnSelect.selectedOptions).map(opt => parseInt(opt.value, 10));

        if (selectedColumnIndexes.length === 0) {
            return selectedData; // Return all data if no specific columns are selected
        }

        return selectedData.map(row =>
            selectedColumnIndexes.map(index => row[index])
        );
    }

    document.getElementById('export-pdf').addEventListener('click', () => {
        const { jsPDF } = window.jspdf;
        html2canvas(dashboard).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const pdf = new jsPDF();
            const imgProps= pdf.getImageProperties(imgData);
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
            pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
            pdf.save("report.pdf");
        });
    });

    document.getElementById('export-png').addEventListener('click', () => {
        html2canvas(dashboard).then(canvas => {
            const link = document.createElement('a');
            link.download = 'report.png';
            link.href = canvas.toDataURL('image/png');
            link.click();
        });
    });

    document.getElementById('export-jpeg').addEventListener('click', () => {
        html2canvas(dashboard).then(canvas => {
            const link = document.createElement('a');
            link.download = 'report.jpeg';
            link.href = canvas.toDataURL('image/jpeg');
            link.click();
        });
    });

    function exportToCsv(filename, rows) {
        const processRow = row => row.map(val => `"${(val || '').toString().replace(/"/g, '""')}"`).join(',');
        const csvContent = "data:text/csv;charset=utf-8," + rows.map(processRow).join('\n');
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.getElementById('export-csv').addEventListener('click', () => {
        const dataToExport = getPreviewData();
        if (dataToExport && dataToExport.length > 0) {
            exportToCsv('report.csv', dataToExport);
        } else {
            alert('No data to export.');
        }
    });

    document.getElementById('export-excel').addEventListener('click', () => {
        const dataToExport = getPreviewData();
        if (dataToExport && dataToExport.length > 0) {
            const newWorkbook = XLSX.utils.book_new();
            const newSheet = XLSX.utils.aoa_to_sheet(dataToExport);
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Preview');

            const wbout = XLSX.write(newWorkbook, {bookType:'xlsx', type: 'binary'});
            function s2ab(s) {
                const buf = new ArrayBuffer(s.length);
                const view = new Uint8Array(buf);
                for (let i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
                return buf;
            }
            const blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'report.xlsx';
            link.click();
        } else {
            alert('No data to export.');
        }
    });
});
