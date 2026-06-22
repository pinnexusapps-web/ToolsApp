// ============================================================
// GENERATE BIO TIME - v1.1
// Converted from VBA to JavaScript
// Developed by DSU11425
// ============================================================

async function runGenerateBioTime(workbook, params) {
    return new Promise((resolve, reject) => {
        try {
            const rosterRow = parseInt(params.rosterRow) || 17;
            const dateRow = parseInt(params.dateRow) || 16;
            const empIdCol = params.empIdCol || 'A';

            if (!workbook) {
                reject(new Error('Please upload an Excel file with roster data.'));
                return;
            }

            const sheetNames = workbook.SheetNames;
            const ws = workbook.Sheets[sheetNames[0]];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

            console.log('📊 Total rows in sheet:', data.length);

            let rosterRowIndex = -1;
            let startColIndex = -1;
            let found = false;

            for (let r = 0; r < Math.min(data.length, 200); r++) {
                const row = data[r];
                if (!row) continue;
                for (let c = 0; c < Math.min(row.length, 50); c++) {
                    const val = String(row[c] || '').toUpperCase().trim();
                    if (val === 'MONDAY' || val === 'MON' || val.includes('MONDAY') || 
                        val === 'TUESDAY' || val === 'WEDNESDAY' || 
                        val === 'THURSDAY' || val === 'FRIDAY' ||
                        val === 'SATURDAY' || val === 'SUNDAY') {
                        rosterRowIndex = r;
                        startColIndex = c;
                        found = true;
                        console.log(`✅ Found day header at row ${r}, col ${c}: "${val}"`);
                        break;
                    }
                }
                if (found) break;
            }

            if (!found) {
                for (let r = 0; r < Math.min(data.length, 200); r++) {
                    const row = data[r];
                    if (!row) continue;
                    for (let c = 0; c < Math.min(row.length, 50); c++) {
                        const val = String(row[c] || '').toUpperCase().trim();
                        if (val === 'EMPLOYEE ID' || val === 'EMPLPYEE ID' || val === 'EMPLOYEE') {
                            rosterRowIndex = r;
                            startColIndex = c;
                            found = true;
                            console.log(`✅ Found Employee ID header at row ${r}, col ${c}: "${val}"`);
                            break;
                        }
                    }
                    if (found) break;
                }
            }

            if (!found) {
                reject(new Error("Could not find roster header row. Please check your sheet structure."));
                return;
            }

            const dates = [];
            let dateFound = false;

            const dateRowIndex = rosterRowIndex - 1;
            if (dateRowIndex >= 0 && data[dateRowIndex]) {
                const dateRowData = data[dateRowIndex];
                for (let i = 0; i < 7; i++) {
                    const colIndex = startColIndex + i;
                    if (colIndex < dateRowData.length) {
                        let dateVal = dateRowData[colIndex];
                        const parsedDate = new Date(dateVal);
                        if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() > 1900) {
                            dates.push(parsedDate);
                            dateFound = true;
                        }
                    }
                }
            }

            if (!dateFound && data[rosterRowIndex]) {
                const rowData = data[rosterRowIndex];
                for (let i = 0; i < 7; i++) {
                    const colIndex = startColIndex + i;
                    if (colIndex < rowData.length) {
                        let dateVal = rowData[colIndex];
                        const parsedDate = new Date(dateVal);
                        if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() > 1900) {
                            dates.push(parsedDate);
                            dateFound = true;
                        }
                    }
                }
            }

            if (!dateFound || dates.length < 7) {
                console.log('⚠️ Using fallback dates');
                for (let i = 0; i < 7; i++) {
                    const fallback = new Date();
                    fallback.setDate(fallback.getDate() + i);
                    dates.push(fallback);
                }
            }

            console.log('📅 Dates found:', dates.map(d => formatDate(d)));

            const empColIndex = empIdCol.toUpperCase().charCodeAt(0) - 65;

            const results = [];
            let processedEmployees = 0;
            let skippedRows = 0;

            for (let r = rosterRowIndex + 1; r < data.length; r++) {
                const row = data[r];
                if (!row || row.length === 0) {
                    skippedRows++;
                    continue;
                }

                const empId = String(row[empColIndex] || '').trim();
                const empIdUpper = empId.toUpperCase();

                if (!empId) {
                    skippedRows++;
                    continue;
                }

                if (empIdUpper === 'EMPLOYEE ID' || 
                    empIdUpper === 'EMPLPYEE ID' || 
                    empIdUpper === 'EMPLOYEE' || 
                    empIdUpper.includes('TOTAL')) {
                    skippedRows++;
                    console.log(`⏭️ Skipping header row: "${empId}"`);
                    continue;
                }

                processedEmployees++;
                console.log(`👤 Processing employee: "${empId}" (row ${r})`);

                let currentShift = '';
                let groupStartDate = null;

                for (let i = 0; i < 7; i++) {
                    const colIndex = startColIndex + i;
                    let shiftVal = '';
                    
                    if (colIndex < row.length) {
                        shiftVal = String(row[colIndex] || '').trim();
                    }

                    const parsedShift = advancedShiftParser(shiftVal);
                    const currentDate = dates[i] || new Date();

                    if (i === 0) {
                        currentShift = parsedShift;
                        groupStartDate = currentDate;
                    } else if (parsedShift !== currentShift) {
                        results.push({
                            'Employee ID': empId,
                            'Shift': currentShift,
                            'Start Date': formatDate(groupStartDate),
                            'End Date': formatDate(dates[i - 1])
                        });
                        console.log(`  📝 Saved shift: ${currentShift} (${formatDate(groupStartDate)} - ${formatDate(dates[i - 1])})`);

                        currentShift = parsedShift;
                        groupStartDate = currentDate;
                    }

                    if (i === 6) {
                        results.push({
                            'Employee ID': empId,
                            'Shift': currentShift,
                            'Start Date': formatDate(groupStartDate),
                            'End Date': formatDate(currentDate)
                        });
                        console.log(`  📝 Final shift: ${currentShift} (${formatDate(groupStartDate)} - ${formatDate(currentDate)})`);
                    }
                }
            }

            console.log(`📊 Processed ${processedEmployees} employees, ${skippedRows} rows skipped`);
            console.log(`📊 Total results: ${results.length}`);

            if (results.length === 0) {
                reject(new Error(`No employee data found. Please check:\n` +
                    `1. The roster row (currently ${rosterRow}) is correct\n` +
                    `2. Employee IDs are in column "${empIdCol}"\n` +
                    `3. Data starts after the header row`));
                return;
            }

            resolve({
                data: results,
                message: `✅ Data Successfully Converted!\n\n` +
                         `Total entries: ${results.length}\n` +
                         `Total employees: ${new Set(results.map(r => r['Employee ID'])).size}\n` +
                         `Rows processed: ${processedEmployees}`,
                summary: {
                    totalEntries: results.length,
                    uniqueEmployees: new Set(results.map(r => r['Employee ID'])).size,
                    processedEmployees: processedEmployees,
                    skippedRows: skippedRows
                }
            });

        } catch (err) {
            console.error('❌ Error:', err);
            reject(err);
        }
    });
}

function advancedShiftParser(rawShift) {
    let shift = String(rawShift || '').toUpperCase().trim();
    shift = shift.replace(/\u00A0/g, ' ').trim();

    const offPatterns = ['OFF', 'MONDAY', 'TUESDAY', 'TUESADY', 'WEDNESDAY', 
                        'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];
    if (offPatterns.includes(shift) || shift === '') {
        return 'OFF';
    }

    try {
        let cleanStr = '';
        for (let ch of shift) {
            if ('0123456789AMPM- ()'.includes(ch)) {
                cleanStr += ch;
            }
        }
        cleanStr = cleanStr.trim();

        const shiftMatch = cleanStr.match(/(AM|PM)\s*(\d+)\s*\((\d+)\)/i);
        if (shiftMatch) {
            const prefix = shiftMatch[1].toUpperCase();
            const hour = String(shiftMatch[2]).padStart(2, '0');
            const duration = String(shiftMatch[3]).padStart(2, '0');
            return `${prefix} ${hour} (${duration})`;
        }

        const simpleMatch = cleanStr.match(/(AM|PM)\s*(\d+)/i);
        if (simpleMatch) {
            const prefix = simpleMatch[1].toUpperCase();
            const hour = String(simpleMatch[2]).padStart(2, '0');
            return `${prefix} ${hour}`;
        }

        cleanStr = cleanStr.replace(/\s*-\s*/g, '-');

        if (!cleanStr.includes('-')) {
            const pmIdx = cleanStr.indexOf('PM');
            const amIdx = cleanStr.indexOf('AM');
            let cutPos = -1;

            if (pmIdx > 0 && amIdx > 0) {
                cutPos = Math.min(pmIdx, amIdx) + 1;
            } else if (pmIdx > 0) {
                cutPos = pmIdx + 1;
            } else if (amIdx > 0) {
                cutPos = amIdx + 1;
            }

            if (cutPos > 0 && cutPos < cleanStr.length) {
                cleanStr = cleanStr.slice(0, cutPos).trim() + '-' + cleanStr.slice(cutPos).trim();
            }
        }

        const parts = cleanStr.split('-').map(s => s.trim());
        if (parts.length < 2) {
            return shift;
        }

        let startStr = parts[0];
        let endStr = parts[1];

        startStr = startStr.replace(/(AM|PM)/g, ' $1').trim();
        endStr = endStr.replace(/(AM|PM)/g, ' $1').trim();

        const startTime = parseTime(startStr);
        const endTime = parseTime(endStr);

        if (!startTime || !endTime) {
            return shift;
        }

        let duration = (endTime - startTime) / (1000 * 60 * 60);
        if (duration < 0) {
            duration = (endTime - startTime + 24 * 60 * 60 * 1000) / (1000 * 60 * 60);
        }

        const startHour = startTime.getHours();
        const prefix = startHour >= 12 ? 'PM' : 'AM';
        const displayHour = startHour % 12 || 12;

        return `${prefix} ${String(displayHour).padStart(2, '0')} (${String(Math.round(duration)).padStart(2, '0')})`;

    } catch (err) {
        return shift;
    }
}

function parseTime(str) {
    try {
        str = str.trim().toUpperCase();
        let hours = 0, minutes = 0;
        let isPM = str.includes('PM');
        let isAM = str.includes('AM');

        let timeStr = str.replace(/[AP]M/g, '').trim();
        const parts = timeStr.split(':');
        if (parts.length === 2) {
            hours = parseInt(parts[0]) || 0;
            minutes = parseInt(parts[1]) || 0;
        } else if (parts.length === 1) {
            hours = parseInt(parts[0]) || 0;
            minutes = 0;
        } else {
            return null;
        }

        if (isPM && hours < 12) hours += 12;
        if (isAM && hours === 12) hours = 0;

        const date = new Date();
        date.setHours(hours, minutes, 0, 0);
        return date;
    } catch (err) {
        return null;
    }
}

function formatDate(date) {
    if (!date || isNaN(date.getTime())) return '';
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
}