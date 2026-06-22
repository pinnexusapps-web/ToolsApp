// ============================================================
// MAIN APPLICATION - Excel Macro Converter v2.1
// Developed by DSU11425
// ============================================================

// === App State ===
const AppState = {
    file: null,
    workbook: null,
    data: null,
    selectedFunction: null,
    results: null,
    outputData: null,
    processing: false,
    currentTableData: null,
    modifiedReasons: {},
    paramsCollapsed: false
};

// === DOM Elements ===
const DOM = {
    uploadArea: document.getElementById('uploadArea'),
    fileInput: document.getElementById('fileInput'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileSize: document.getElementById('fileSize'),
    removeFile: document.getElementById('removeFile'),
    functionCards: document.querySelectorAll('.function-card'),
    paramsSection: document.getElementById('paramsSection'),
    paramsContent: document.getElementById('paramsContent'),
    paramsToggle: document.getElementById('paramsToggle'),
    paramToggleIcon: document.getElementById('paramToggleIcon'),
    paramsContentWrapper: document.getElementById('paramsContentWrapper'),
    actionSection: document.getElementById('actionSection'),
    runMacro: document.getElementById('runMacro'),
    downloadResult: document.getElementById('downloadResult'),
    resultsSection: document.getElementById('resultsSection'),
    resultsContent: document.getElementById('resultsContent'),
    copyResults: document.getElementById('copyResults'),
    closeResults: document.getElementById('closeResults'),
    outputSection: document.getElementById('outputSection'),
    outputHead: document.getElementById('outputHead'),
    outputBody: document.getElementById('outputBody'),
    outputTable: document.getElementById('outputTable'),
    rowCount: document.getElementById('rowCount'),
    resetParams: document.getElementById('resetParams'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    copyTable: document.getElementById('copyTable'),
    downloadCsv: document.getElementById('downloadCsv'),
    downloadPng: document.getElementById('downloadPng'),
    downloadPdf: document.getElementById('downloadPdf'),
    uploadSection: document.getElementById('uploadSection')
};

// === Toast System ===
function showToast(message, type = 'info') {
    const container = document.querySelector('.toast-container') || (() => {
        const c = document.createElement('div');
        c.className = 'toast-container';
        document.body.appendChild(c);
        return c;
    })();

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icons = {
        success: 'fa-check-circle',
        error: 'fa-exclamation-circle',
        warning: 'fa-exclamation-triangle',
        info: 'fa-info-circle'
    };
    toast.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i> ${message}`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(100%)';
        toast.style.transition = 'all 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// === Loading Overlay ===
function showLoading(message = 'Processing...') {
    const p = DOM.loadingOverlay.querySelector('p');
    if (p) p.textContent = message;
    DOM.loadingOverlay.classList.add('active');
}

function hideLoading() {
    DOM.loadingOverlay.classList.remove('active');
}

// === File Handling ===
DOM.uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    DOM.uploadArea.parentElement.classList.add('dragover');
});

DOM.uploadArea.addEventListener('dragleave', () => {
    DOM.uploadArea.parentElement.classList.remove('dragover');
});

DOM.uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    DOM.uploadArea.parentElement.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
    }
});

DOM.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
});

DOM.removeFile.addEventListener('click', () => {
    resetFile();
});

function handleFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/csv'
    ];

    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showToast('Please upload a valid Excel or CSV file.', 'error');
        return;
    }

    const maxSize = 20 * 1024 * 1024;
    if (file.size > maxSize) {
        showToast('File is too large. Maximum size is 20MB.', 'error');
        return;
    }

    AppState.file = file;
    updateFileInfo(file);
    readFile(file);
}

function updateFileInfo(file) {
    DOM.fileName.textContent = file.name;
    DOM.fileSize.textContent = (file.size / 1024).toFixed(1) + ' KB';
    DOM.fileInfo.style.display = 'flex';
    DOM.uploadArea.style.display = 'none';
}

function resetFile() {
    AppState.file = null;
    AppState.workbook = null;
    AppState.data = null;
    DOM.fileInfo.style.display = 'none';
    DOM.uploadArea.style.display = 'block';
    DOM.fileInput.value = '';
    DOM.resultsSection.style.display = 'none';
    DOM.outputSection.style.display = 'none';
    DOM.actionSection.style.display = 'none';
    DOM.paramsSection.style.display = 'none';
    document.querySelectorAll('.function-card').forEach(c => c.classList.remove('active'));
    AppState.selectedFunction = null;
}

function readFile(file) {
    showLoading('Reading file...');
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            AppState.workbook = workbook;
            AppState.data = workbook;

            showToast(`File loaded successfully! Found ${workbook.SheetNames.length} sheet(s).`, 'success');
            hideLoading();
            enableFunctionSelection();
        } catch (err) {
            hideLoading();
            showToast('Error reading file: ' + err.message, 'error');
            console.error(err);
        }
    };

    reader.onerror = () => {
        hideLoading();
        showToast('Error reading file.', 'error');
    };

    reader.readAsArrayBuffer(file);
}

// === Function Selection ===
function enableFunctionSelection() {
    DOM.functionCards.forEach(card => {
        card.style.opacity = '1';
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => selectFunction(card));
    });
}

function selectFunction(card) {
    DOM.functionCards.forEach(c => c.classList.remove('active'));
    card.classList.add('active');
    AppState.selectedFunction = parseInt(card.dataset.function);

    loadParameters(AppState.selectedFunction);
    DOM.paramsSection.style.display = 'block';
    DOM.actionSection.style.display = 'block';
    DOM.resultsSection.style.display = 'none';
    DOM.outputSection.style.display = 'none';
    
    AppState.paramsCollapsed = false;
    DOM.paramsContentWrapper.style.display = 'block';
    DOM.paramToggleIcon.innerHTML = '<i class="fas fa-chevron-down"></i>';
}

// === Load Parameters ===
function loadParameters(funcNum) {
    const container = DOM.paramsContent;
    let html = '';

    switch (funcNum) {
        case 1:
            html = `
                <div class="param-group">
                    <label>SKU Column Header</label>
                    <input type="text" id="skuHeader" value="SKU" />
                    <span class="hint">The column header that contains SKU values</span>
                </div>
                <div class="param-group">
                    <label>Data Sheet Name</label>
                    <input type="text" id="dataSheet" value="Sheet1" />
                    <span class="hint">Name of the sheet containing lookup data</span>
                </div>
                <div class="param-group" style="grid-column: 1 / -1;">
                    <label>CSV Data Paste Area</label>
                    <textarea id="csvData" placeholder="Paste your CSV data here..." style="min-height:100px;"></textarea>
                    <span class="hint">📋 Paste CSV data or upload via the file input above</span>
                </div>
            `;
            break;

        case 2:
            html = `
                <div class="param-group">
                    <label>Roster Row</label>
                    <input type="number" id="rosterRow" value="17" min="1" />
                    <span class="hint">Row number where roster data starts</span>
                </div>
                <div class="param-group">
                    <label>Date Row</label>
                    <input type="number" id="dateRow" value="16" min="1" />
                    <span class="hint">Row number containing date headers</span>
                </div>
                <div class="param-group">
                    <label>Employee ID Column</label>
                    <input type="text" id="empIdCol" value="A" />
                    <span class="hint">Column letter for Employee IDs</span>
                </div>
            `;
            break;

        case 3:
            html = `
                <div class="param-group">
                    <label>Data Start Row</label>
                    <input type="number" id="startRow" value="2" min="1" />
                    <span class="hint">First row containing data (after headers)</span>
                </div>
                <div class="param-group">
                    <label>Date Column</label>
                    <input type="text" id="dateCol" value="D" />
                    <span class="hint">Column letter for date values</span>
                </div>
                <div class="param-group">
                    <label>Picker Column</label>
                    <input type="text" id="pickerCol" value="G" />
                    <span class="hint">Column letter for picker names</span>
                </div>
                <div class="param-group">
                    <label>Status Column</label>
                    <input type="text" id="statusCol" value="C" />
                    <span class="hint">Column letter for order status</span>
                </div>
                <div class="param-group">
                    <label>Price Column</label>
                    <input type="text" id="priceCol" value="I" />
                    <span class="hint">Column letter for item price</span>
                </div>
                <div class="param-group">
                    <label>Final Price Column</label>
                    <input type="text" id="finalPriceCol" value="J" />
                    <span class="hint">Column letter for final price</span>
                </div>
                <div class="param-group">
                    <label>Order ID Column</label>
                    <input type="text" id="orderIdCol" value="A" />
                    <span class="hint">Column letter for Order ID</span>
                </div>
            `;
            break;
    }

    container.innerHTML = html;
}

// === Collapsible Parameters ===
if (DOM.paramsToggle) {
    DOM.paramsToggle.addEventListener('click', () => {
        AppState.paramsCollapsed = !AppState.paramsCollapsed;
        if (AppState.paramsCollapsed) {
            DOM.paramsContentWrapper.style.display = 'none';
            DOM.paramToggleIcon.innerHTML = '<i class="fas fa-chevron-right"></i>';
        } else {
            DOM.paramsContentWrapper.style.display = 'block';
            DOM.paramToggleIcon.innerHTML = '<i class="fas fa-chevron-down"></i>';
        }
    });
}

// === Get Parameter Values ===
function getParams() {
    const params = {};
    const inputs = DOM.paramsContent.querySelectorAll('input, select, textarea');
    inputs.forEach(input => {
        params[input.id] = input.value;
    });
    return params;
}

// === Run Macro ===
DOM.runMacro.addEventListener('click', async () => {
    if (!AppState.selectedFunction) {
        showToast('Please select a macro function first.', 'warning');
        return;
    }

    if (!AppState.workbook && AppState.selectedFunction !== 1) {
        showToast('Please upload an Excel file first.', 'warning');
        return;
    }

    AppState.processing = true;
    DOM.runMacro.disabled = true;
    showLoading('Processing...');

    try {
        const params = getParams();
        let result = null;

        switch (AppState.selectedFunction) {
            case 1:
                result = await runCalculateReceivings(AppState.workbook, params);
                break;
            case 2:
                result = await runGenerateBioTime(AppState.workbook, params);
                break;
            case 3:
                result = await runGenerateDailyReport(AppState.workbook, params);
                break;
        }

        if (result) {
            AppState.results = result;
            displayResults(result);
            
            if (AppState.selectedFunction === 3 && result.report) {
                displayReportTable(result);
            } else if (result.data) {
                displayOutputTable(result.data);
            } else if (result.report) {
                displayOutputTable(result.report);
            }
            showToast('Macro executed successfully!', 'success');
        }
    } catch (err) {
        showToast('Error: ' + err.message, 'error');
        console.error(err);
    } finally {
        AppState.processing = false;
        DOM.runMacro.disabled = false;
        hideLoading();
    }
});

// === Display Results ===
function displayResults(result) {
    DOM.resultsSection.style.display = 'block';
    DOM.resultsContent.textContent = result.message || JSON.stringify(result, null, 2);
}

// === Display Output Table (Generic) ===
function displayOutputTable(data) {
    if (!data || !data.length) {
        DOM.outputSection.style.display = 'none';
        return;
    }

    AppState.currentTableData = data;
    DOM.outputSection.style.display = 'block';

    // Check if outputHead and outputBody exist
    if (!DOM.outputHead || !DOM.outputBody) {
        console.error('Output table elements not found');
        return;
    }

    const headers = Object.keys(data[0]);
    DOM.outputHead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;
    DOM.outputBody.innerHTML = data.map(row =>
        `<tr>${headers.map(h => `<td>${row[h] !== undefined ? row[h] : ''}</td>`).join('')}</tr>`
    ).join('');

    DOM.rowCount.textContent = `${data.length} rows`;

    if (DOM.copyTable) DOM.copyTable.style.display = 'inline-flex';
    if (DOM.downloadCsv) DOM.downloadCsv.style.display = 'inline-flex';
    if (DOM.downloadPng) DOM.downloadPng.style.display = 'none';
    if (DOM.downloadPdf) DOM.downloadPdf.style.display = 'none';
}

// === Display Report Table (For Daily Report) ===
function displayReportTable(result) {
    const reportData = result.report || [];
    const modifiedData = result.modifiedOrders || [];
    
    if (!reportData.length) {
        DOM.outputSection.style.display = 'none';
        return;
    }

    AppState.currentTableData = reportData;
    AppState.modifiedReasons = {};

    modifiedData.forEach((row, index) => {
        const key = `${row.PICKER}_${row['ORDER ID']}`;
        AppState.modifiedReasons[key] = row.REASON || 'Not found';
    });

    DOM.outputSection.style.display = 'block';

    let html = '';
    
    const reportHeaders = Object.keys(reportData[0]);
    const modHeaders = modifiedData.length > 0 ? Object.keys(modifiedData[0]) : [];
    
    let maxNameLength = 15;
    reportData.forEach(row => {
        const nameLength = String(row.PICKER || '').length;
        if (nameLength > maxNameLength) maxNameLength = nameLength;
    });
    const nameColWidth = Math.min(maxNameLength + 2, 35);
    
    html += `<h3 style="margin: 0 0 12px 0; font-size: 16px; font-weight: 500; text-align: center;">📊 Picker Summary</h3>`;
    html += `<table class="output-table" style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;table-layout:fixed;">`;
    html += `<thead><tr>`;
    const colWidths = {
        'PICKER': nameColWidth,
        'CANCELLED': 15,
        'SUCCESS': 15,
        'PICKED UP': 15
    };
    reportHeaders.forEach(h => {
        const width = colWidths[h] || 20;
        html += `<th style="padding:8px 10px;text-align:center;background:#f8f9fa;border-bottom:2px solid #d1d5db;font-weight:600;font-size:13px;width:${width}%;">${h}</th>`;
    });
    html += `</tr></thead><tbody>`;

    reportData.forEach((row, idx) => {
        const isGrandTotal = row.PICKER === 'Grand Total';
        const bgColor = isGrandTotal ? '#f8f9fa' : (idx % 2 === 0 ? '#ffffff' : '#fafafa');
        const fontWeight = isGrandTotal ? 'bold' : 'normal';
        const borderTop = isGrandTotal ? 'border-top:2px solid #d1d5db;' : '';
        const fontSize = isGrandTotal ? '15px' : '14px';
        
        html += `<tr style="background:${bgColor};border-bottom:1px solid #f0f0f0;height:34px;">`;
        reportHeaders.forEach(h => {
            const value = row[h] !== undefined ? row[h] : '';
            html += `<td style="padding:4px 10px;text-align:center;font-weight:${fontWeight};${borderTop}font-size:${fontSize};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${value}</td>`;
        });
        html += `</tr>`;
    });
    html += `</tbody></table>`;

    // Modified Orders Section
    html += `<h3 style="margin: 14px 0 8px 0; font-size: 15px; font-weight: 500; text-align: center;">📝 Modified Orders List</h3>`;
    
    if (modifiedData.length > 0) {
        html += `<table class="output-table" style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed;">`;
        html += `<thead><tr>`;
        const modColWidths = {
            'PICKER': 25,
            'ORDER ID': 25,
            'REASON': 50
        };
        modHeaders.forEach(h => {
            const width = modColWidths[h] || 33;
            html += `<th style="padding:6px 10px;text-align:center;background:#f8f9fa;border-bottom:1px solid #d1d5db;font-weight:600;font-size:12px;width:${width}%;">${h}</th>`;
        });
        html += `</tr></thead><tbody>`;

        modifiedData.forEach((row, idx) => {
            const key = `${row.PICKER}_${row['ORDER ID']}`;
            const reasonValue = AppState.modifiedReasons[key] || row.REASON || 'Not found';
            const bgColor = idx % 2 === 0 ? '#ffffff' : '#fafafa';
            
            html += `<tr style="background:${bgColor};border-bottom:1px solid #f0f0f0;height:32px;">`;
            html += `<td style="padding:4px 10px;text-align:center;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${row.PICKER}</td>`;
            html += `<td style="padding:4px 10px;text-align:center;font-size:13px;">${row['ORDER ID']}</td>`;
            html += `<td style="padding:4px 10px;text-align:center;">
                <input type="text" 
                       class="reason-input" 
                       data-key="${key}"
                       value="${reasonValue}"
                       style="width:100%;padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:13px;font-family:inherit;text-align:center;background:${reasonValue === 'Not found' ? '#fff3cd' : '#ffffff'};"
                       placeholder="Enter reason..."
                       onchange="updateReason(this)" />
            </td>`;
            html += `</tr>`;
        });
        html += `</tbody></table>`;
        html += `<div style="margin-top:6px;font-size:11px;color:#6c757d;text-align:center;">💡 Click on the reason cell to edit. Changes will be included in the PNG download.</div>`;
    } else {
        html += `<div style="padding:10px;background:#fafafa;border-radius:6px;color:#6c757d;font-style:italic;text-align:center;border:1px solid #e9ecef;font-size:13px;">
            No modifications recorded today.
        </div>`;
    }

    // Set the HTML content directly in the outputBody
    if (DOM.outputBody) {
        DOM.outputBody.innerHTML = html;
        DOM.rowCount.textContent = `${reportData.length} rows`;
    }

    if (DOM.copyTable) DOM.copyTable.style.display = 'inline-flex';
    if (DOM.downloadCsv) DOM.downloadCsv.style.display = 'inline-flex';
    if (DOM.downloadPng) DOM.downloadPng.style.display = 'inline-flex';
    if (DOM.downloadPdf) DOM.downloadPdf.style.display = 'inline-flex';
}

// === Update Reason ===
window.updateReason = function(input) {
    const key = input.dataset.key;
    AppState.modifiedReasons[key] = input.value;
    if (input.value && input.value !== 'Not found' && input.value.trim() !== '') {
        input.style.background = '#e8f5e9';
    } else {
        input.style.background = '#fff3cd';
    }
};

// === Copy Table to Clipboard ===
function copyTableToClipboard() {
    const data = AppState.currentTableData;
    if (!data || !data.length) {
        showToast('No data to copy.', 'warning');
        return;
    }

    try {
        const headers = Object.keys(data[0]);
        let tsv = headers.join('\t') + '\n';
        
        data.forEach(row => {
            const rowData = headers.map(h => {
                let val = row[h] !== undefined ? row[h] : '';
                if (typeof val === 'string' && (val.includes('\t') || val.includes('\n'))) {
                    val = `"${val}"`;
                }
                return val;
            });
            tsv += rowData.join('\t') + '\n';
        });

        navigator.clipboard.writeText(tsv).then(() => {
            showToast(`✅ Copied ${data.length} rows to clipboard!`, 'success');
        }).catch(() => {
            const textarea = document.createElement('textarea');
            textarea.value = tsv;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
            showToast(`✅ Copied ${data.length} rows to clipboard!`, 'success');
        });
    } catch (err) {
        showToast('Error copying data: ' + err.message, 'error');
        console.error(err);
    }
}

// === Download as CSV ===
function downloadAsCsv() {
    const data = AppState.currentTableData;
    if (!data || !data.length) {
        showToast('No data to download.', 'warning');
        return;
    }

    try {
        const headers = Object.keys(data[0]);
        let csv = headers.join(',') + '\n';
        
        data.forEach(row => {
            const rowData = headers.map(h => {
                let val = row[h] !== undefined ? row[h] : '';
                if (typeof val === 'string' && (val.includes(',') || val.includes('"') || val.includes('\n'))) {
                    val = `"${val.replace(/"/g, '""')}"`;
                }
                return val;
            });
            csv += rowData.join(',') + '\n';
        });

        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = `daily_report_${new Date().toISOString().slice(0,10)}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        showToast(`✅ Downloaded ${data.length} rows as CSV!`, 'success');
    } catch (err) {
        showToast('Error downloading CSV: ' + err.message, 'error');
        console.error(err);
    }
}

// === Download as PNG ===
function downloadReportAsPNG() {
    const results = AppState.results;
    
    if (!results || !results.report) {
        showToast('No report data available. Run a macro first.', 'warning');
        return;
    }

    document.querySelectorAll('.reason-input').forEach(input => {
        const key = input.dataset.key;
        AppState.modifiedReasons[key] = input.value;
    });

    showLoading('Generating PNG...');

    try {
        const reportDiv = document.createElement('div');
        reportDiv.style.cssText = `
            position: fixed;
            left: -9999px;
            top: 0;
            width: 800px;
            background: white;
            padding: 40px;
            font-family: 'Inter', -apple-system, Arial, sans-serif;
            z-index: -1;
        `;
        
        reportDiv.innerHTML = buildReportHTML(results);
        document.body.appendChild(reportDiv);

        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        script.onload = () => {
            html2canvas(reportDiv, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff',
                logging: false,
                width: 800,
                height: reportDiv.scrollHeight + 50
            }).then(canvas => {
                const link = document.createElement('a');
                const date = results.summary.date || 'today';
                link.download = `daily_report_${date}.png`;
                link.href = canvas.toDataURL('image/png');
                link.click();
                document.body.removeChild(reportDiv);
                hideLoading();
                showToast('✅ Report downloaded as PNG!', 'success');
            }).catch(err => {
                document.body.removeChild(reportDiv);
                hideLoading();
                showToast('Error generating PNG: ' + err.message, 'error');
                console.error(err);
            });
        };
        script.onerror = () => {
            document.body.removeChild(reportDiv);
            hideLoading();
            showToast('Failed to load image library. Please check your internet connection.', 'error');
        };
        document.head.appendChild(script);

    } catch (err) {
        console.error('Error generating PNG:', err);
        hideLoading();
        showToast('Error generating PNG: ' + err.message, 'error');
    }
}

// === Build Report HTML (For PNG/PDF - NO SUMMARY) ===
function buildReportHTML(results) {
    const summary = results.summary;
    const reportRows = results.report || [];
    const modifiedRows = results.modifiedOrders || [];

    const reasons = {};
    document.querySelectorAll('.reason-input').forEach(input => {
        reasons[input.dataset.key] = input.value;
    });

    let maxNameLength = 0;
    reportRows.forEach(row => {
        const nameLength = String(row.PICKER || '').length;
        if (nameLength > maxNameLength) maxNameLength = nameLength;
    });
    if (maxNameLength < 10) maxNameLength = 15;
    const nameColWidth = Math.min(maxNameLength + 2, 30);

    let html = `
        <div style="margin-bottom: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                <h1 style="font-size: 22px; font-weight: 600; margin: 0; color: #1a1a2e; text-align: center; width: 100%;">📊 Daily Report</h1>
            </div>
            <div style="text-align: center; font-size: 16px; color: #6c757d; font-weight: 500; margin-bottom: 12px;">
                ${summary.date || 'Today'}
            </div>
            <hr style="border: none; border-top: 2px solid #e9ecef; margin: 4px 0 14px 0;" />
        </div>
    `;

    // Picker Table
    html += `
        <div style="margin-bottom: 16px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 14px; table-layout: fixed;">
                <thead>
                    <tr style="background: #f8f9fa; border-bottom: 2px solid #d1d5db;">
                        <th style="padding: 8px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 13px; width: ${nameColWidth}%;">PICKER</th>
                        <th style="padding: 8px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 13px; width: 15%;">CANCELLED</th>
                        <th style="padding: 8px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 13px; width: 15%;">SUCCESS</th>
                        <th style="padding: 8px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 13px; width: 15%;">PICKED UP</th>
                    </tr>
                </thead>
                <tbody>
    `;

    reportRows.forEach((row, index) => {
        const isGrandTotal = row.PICKER === 'Grand Total';
        const bgColor = isGrandTotal ? '#f8f9fa' : (index % 2 === 0 ? '#ffffff' : '#fafafa');
        const fontWeight = isGrandTotal ? 'bold' : 'normal';
        const borderTop = isGrandTotal ? 'border-top: 2px solid #d1d5db;' : '';
        const fontSize = isGrandTotal ? '15px' : '14px';
        
        html += `
            <tr style="background: ${bgColor}; border-bottom: 1px solid #f0f0f0; height: 32px;">
                <td style="padding: 4px 10px; text-align: center; font-weight: ${fontWeight}; ${borderTop} font-size: ${fontSize}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.PICKER}</td>
                <td style="padding: 4px 10px; text-align: center; ${borderTop} font-size: ${fontSize};">${row.CANCELLED}</td>
                <td style="padding: 4px 10px; text-align: center; ${borderTop} font-size: ${fontSize};">${row.SUCCESS}</td>
                <td style="padding: 4px 10px; text-align: center; font-weight: ${isGrandTotal ? 'bold' : 'normal'}; ${borderTop} font-size: ${fontSize};">${row['PICKED UP']}</td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
    `;

    // Modified Orders List
    html += `
        <div style="margin-bottom: 12px;">
            <h3 style="font-size: 15px; font-weight: 600; margin: 0 0 6px 0; color: #1a1a2e; text-align: center;">📝 Modified Orders List</h3>
    `;

    if (modifiedRows && modifiedRows.length > 0) {
        html += `
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; table-layout: fixed;">
                <thead>
                    <tr style="background: #f8f9fa; border-bottom: 1px solid #d1d5db;">
                        <th style="padding: 6px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 12px; width: 25%;">PICKER</th>
                        <th style="padding: 6px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 12px; width: 25%;">ORDER ID</th>
                        <th style="padding: 6px 10px; text-align: center; font-weight: 700; color: #1a1a2e; font-size: 12px; width: 50%;">REASON</th>
                    </tr>
                </thead>
                <tbody>
        `;

        modifiedRows.forEach((row, index) => {
            const key = `${row.PICKER}_${row['ORDER ID']}`;
            const reasonValue = reasons[key] || row.REASON || 'Not found';
            const bgColor = index % 2 === 0 ? '#ffffff' : '#fafafa';
            
            html += `
                <tr style="background: ${bgColor}; border-bottom: 1px solid #f0f0f0; height: 30px;">
                    <td style="padding: 4px 10px; text-align: center; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.PICKER}</td>
                    <td style="padding: 4px 10px; text-align: center; font-size: 13px;">${row['ORDER ID']}</td>
                    <td style="padding: 4px 10px; text-align: center; font-size: 13px; color: ${reasonValue === 'Not found' ? '#dc3545' : '#1a1a2e'};">${reasonValue}</td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        `;
    } else {
        html += `
            <div style="padding: 10px; background: #fafafa; border-radius: 4px; color: #6c757d; font-style: italic; text-align: center; border: 1px solid #e9ecef; font-size: 13px;">
                No modifications recorded today.
            </div>
        `;
    }

    html += `</div>`;

    // Footer only - NO SUMMARY
    html += `
        <div style="margin-top: 12px; padding-top: 6px; border-top: 1px solid #e9ecef; font-size: 9px; color: #adb5bd; text-align: center;">
            Generated by Excel Macro Converter v2.1 | DSU11425
        </div>
    `;

    return html;
}

// === Download as PDF ===
function downloadReportAsPDF() {
    const results = AppState.results;
    
    if (!results || !results.report) {
        showToast('No report data available. Run a macro first.', 'warning');
        return;
    }

    document.querySelectorAll('.reason-input').forEach(input => {
        const key = input.dataset.key;
        AppState.modifiedReasons[key] = input.value;
    });

    showLoading('Generating PDF...');

    try {
        const reportDiv = document.createElement('div');
        reportDiv.style.cssText = `
            position: fixed;
            left: -9999px;
            top: 0;
            width: 800px;
            background: white;
            padding: 40px;
            font-family: 'Inter', -apple-system, Arial, sans-serif;
            z-index: -1;
        `;
        
        reportDiv.innerHTML = buildReportHTML(results);
        document.body.appendChild(reportDiv);

        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
        script.onload = () => {
            const opt = {
                margin: [10, 10, 10, 10],
                filename: `daily_report_${results.summary.date || 'today'}.pdf`,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true, logging: false },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };
            html2pdf().set(opt).from(reportDiv).save().then(() => {
                document.body.removeChild(reportDiv);
                hideLoading();
                showToast('✅ Report downloaded as PDF!', 'success');
            }).catch(err => {
                document.body.removeChild(reportDiv);
                hideLoading();
                showToast('Error generating PDF: ' + err.message, 'error');
                console.error(err);
            });
        };
        script.onerror = () => {
            document.body.removeChild(reportDiv);
            hideLoading();
            showToast('Failed to load PDF library. Please check your internet connection.', 'error');
        };
        document.head.appendChild(script);

    } catch (err) {
        console.error('Error generating PDF:', err);
        hideLoading();
        showToast('Error generating PDF: ' + err.message, 'error');
    }
}

// === Copy Results ===
DOM.copyResults.addEventListener('click', () => {
    const text = DOM.resultsContent.textContent;
    navigator.clipboard.writeText(text)
        .then(() => showToast('Copied to clipboard!', 'success'))
        .catch(() => {
            const range = document.createRange();
            range.selectNode(DOM.resultsContent);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            document.execCommand('copy');
            window.getSelection().removeAllRanges();
            showToast('Copied to clipboard!', 'success');
        });
});

// === Close Results ===
DOM.closeResults.addEventListener('click', () => {
    DOM.resultsSection.style.display = 'none';
});

// === Reset Parameters ===
DOM.resetParams.addEventListener('click', (e) => {
    e.stopPropagation();
    if (AppState.selectedFunction) {
        loadParameters(AppState.selectedFunction);
        showToast('Parameters reset to defaults.', 'info');
    }
});

// === Download Result (Excel) ===
DOM.downloadResult.addEventListener('click', () => {
    if (!AppState.results) {
        showToast('No results to download. Run a macro first.', 'warning');
        return;
    }

    const data = AppState.results.report || AppState.results.data || AppState.results;
    let filename = `macro_result_${new Date().toISOString().slice(0,10)}`;

    if (Array.isArray(data) && data.length) {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Results');
        const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([buf], { type: 'application/octet-stream' });
        downloadBlob(blob, `${filename}.xlsx`);
        showToast('Download started!', 'success');
    } else {
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'text/plain' });
        downloadBlob(blob, `${filename}.txt`);
        showToast('Download started!', 'success');
    }
});

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// === Event Listeners ===
if (DOM.copyTable) {
    DOM.copyTable.addEventListener('click', copyTableToClipboard);
}

if (DOM.downloadCsv) {
    DOM.downloadCsv.addEventListener('click', downloadAsCsv);
}

if (DOM.downloadPng) {
    DOM.downloadPng.addEventListener('click', downloadReportAsPNG);
}

if (DOM.downloadPdf) {
    DOM.downloadPdf.addEventListener('click', downloadReportAsPDF);
}

// === Keyboard Shortcuts ===
document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        DOM.runMacro.click();
    }
});

// === Initial State ===
document.querySelectorAll('.function-card').forEach(card => {
    card.style.opacity = '0.5';
    card.style.cursor = 'default';
});

console.log('📊 Excel Macro Converter v2.1 loaded.');
console.log('👨‍💻 Developed by DSU11425');