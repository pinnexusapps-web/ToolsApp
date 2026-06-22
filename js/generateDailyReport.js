// ============================================================
// GENERATE DAILY REPORT - v1.3 (FIXED)
// Converted from VBA to JavaScript
// Developed by DSU11425
// ============================================================

async function runGenerateDailyReport(workbook, params) {
    return new Promise((resolve, reject) => {
        try {
            console.log('🚀 Starting Generate Daily Report...');

            const startRow = parseInt(params.startRow) || 2;
            const dateCol = params.dateCol || 'D';
            const pickerCol = params.pickerCol || 'G';
            const statusCol = params.statusCol || 'C';
            const priceCol = params.priceCol || 'I';
            const finalPriceCol = params.finalPriceCol || 'J';
            const orderIdCol = params.orderIdCol || 'A';

            if (!workbook) {
                reject(new Error('Please upload an Excel file with order data.'));
                return;
            }

            const sheetNames = workbook.SheetNames;
            const ws = workbook.Sheets[sheetNames[0]];
            const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

            const dateColIdx = columnToIndex(dateCol);
            const pickerColIdx = columnToIndex(pickerCol);
            const statusColIdx = columnToIndex(statusCol);
            const priceColIdx = columnToIndex(priceCol);
            const finalPriceColIdx = columnToIndex(finalPriceCol);
            const orderIdColIdx = columnToIndex(orderIdCol);

            // Parse Data
            const rows = [];
            let dateValue = '';

            for (let r = startRow - 1; r < data.length; r++) {
                const row = data[r];
                if (!row || row.length === 0) continue;

                const picker = String(row[pickerColIdx] || '').trim();
                if (!picker || picker === 'Picker' || picker === 'PICKER') continue;

                const status = String(row[statusColIdx] || '').trim().toUpperCase();
                const price = parseFloat(String(row[priceColIdx] || '0').replace(/[^0-9.]/g, '')) || 0;
                const finalPrice = parseFloat(String(row[finalPriceColIdx] || '0').replace(/[^0-9.]/g, '')) || 0;

                if (r === startRow - 1) {
                    const rawDate = String(row[dateColIdx] || '');
                    dateValue = rawDate.split(' ')[0] || rawDate;
                }

                rows.push({
                    rowIndex: r,
                    picker: picker,
                    status: status,
                    price: price,
                    finalPrice: finalPrice,
                    orderId: String(row[orderIdColIdx] || '').trim(),
                    date: String(row[dateColIdx] || '')
                });
            }

            // Calculate Picker Statistics
            const pickerStats = {};
            let totalCancelled = 0;
            let totalSuccess = 0;
            let totalPickedUp = 0;
            let totalSales = 0;
            let totalModifications = 0;

            rows.forEach(row => {
                if (!pickerStats[row.picker]) {
                    pickerStats[row.picker] = {
                        cancelled: 0,
                        success: 0,
                        pickedUp: 0,
                        total: 0,
                        modifications: 0,
                        sales: 0
                    };
                }

                if (row.status === 'CANCELLED') {
                    pickerStats[row.picker].cancelled++;
                    totalCancelled++;
                } else if (row.status === 'PICKED_UP') {
                    pickerStats[row.picker].success++;
                    pickerStats[row.picker].pickedUp++;
                    pickerStats[row.picker].sales += row.finalPrice;
                    totalSuccess++;
                    totalPickedUp++;
                    totalSales += row.finalPrice;

                    const diff = Math.abs(row.finalPrice - row.price);
                    if (diff >= 1) {
                        pickerStats[row.picker].modifications++;
                        totalModifications++;
                    }
                }

                pickerStats[row.picker].total++;
            });

            // Build Report Rows - KEEP TG PREFIX
            const reportRows = [];
            const sortedPickers = Object.keys(pickerStats).sort((a, b) => {
                return (pickerStats[b].pickedUp + pickerStats[b].cancelled) - 
                       (pickerStats[a].pickedUp + pickerStats[a].cancelled);
            });

            sortedPickers.forEach(picker => {
                const stats = pickerStats[picker];
                const totalOrders = stats.cancelled + stats.pickedUp;
                
                // KEEP FULL NAME WITH TG PREFIX
                reportRows.push({
                    'PICKER': picker,
                    'CANCELLED': stats.cancelled || '',
                    'SUCCESS': stats.pickedUp,
                    'PICKED UP': totalOrders
                });
            });

            // Add Grand Total
            reportRows.push({
                'PICKER': 'Grand Total',
                'CANCELLED': totalCancelled,
                'SUCCESS': totalSuccess,
                'PICKED UP': totalPickedUp
            });

            // Modified Orders List
            const modifiedOrders = rows
                .filter(row => row.status === 'PICKED_UP' && Math.abs(row.finalPrice - row.price) >= 1)
                .map(row => ({
                    'PICKER': row.picker,
                    'ORDER ID': row.orderId,
                    'REASON': 'Not found'
                }));

            // Calculate Summary
            const avgSales = totalPickedUp > 0 ? totalSales / totalPickedUp : 0;

            const summary = {
                date: dateValue,
                orders: totalPickedUp,
                cancelled: totalCancelled,
                sales: totalSales,
                average: avgSales,
                modifications: totalModifications
            };

            // Message
            const message = `✅ Daily Report Generated Successfully!\n\n` +
                           `══════════════════════════════════\n` +
                           `📊 Date: ${summary.date}\n` +
                           `📋 Orders: ${summary.orders}\n` +
                           `❌ Cancelled: ${summary.cancelled}\n` +
                           `💰 Sales: ${summary.sales.toFixed(2)}\n` +
                           `📈 Average: ${summary.average.toFixed(2)}\n` +
                           `🔄 Modifications: ${summary.modifications}\n` +
                           `══════════════════════════════════`;

            resolve({
                report: reportRows,
                modifiedOrders: modifiedOrders,
                summary: summary,
                message: message
            });

        } catch (err) {
            console.error('❌ Error:', err);
            reject(err);
        }
    });
}

function columnToIndex(col) {
    col = col.toUpperCase();
    let index = 0;
    for (let i = 0; i < col.length; i++) {
        index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index - 1;
}