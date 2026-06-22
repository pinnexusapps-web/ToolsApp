// ============================================================
// CALCULATE RECEIVINGS - v1.2.2 (FINAL FIXED)
// Converted from VBA to JavaScript
// Developed by DSU11425
// ============================================================

async function runCalculateReceivings(workbook, params) {
    return new Promise((resolve, reject) => {
        try {
            console.log('🚀 Starting Calculate Receivings...');
            
            const skuHeader = params.skuHeader || 'SKU';
            const dataSheetName = params.dataSheet || 'Sheet1';
            const csvData = params.csvData || '';

            let mainSheet, dataSheet;

            // === STEP 1: Load Data ===
            if (workbook) {
                console.log('📂 Using uploaded workbook');
                const sheetNames = workbook.SheetNames;
                mainSheet = workbook.Sheets[sheetNames[0]];
                dataSheet = workbook.Sheets[dataSheetName] || workbook.Sheets[sheetNames[0]];
                console.log(`📊 Main sheet: "${sheetNames[0]}", Data sheet: "${dataSheetName}"`);
            } else {
                if (!csvData) {
                    reject(new Error('Please provide either an Excel file or CSV data.'));
                    return;
                }
                console.log('📄 Using CSV data');
                const lines = csvData.split('\n').filter(line => line.trim());
                const headers = lines[0].split(',').map(h => h.trim());
                const rows = lines.slice(1).map(line => line.split(',').map(v => v.trim()));
                const wsData = [headers, ...rows];
                const ws = XLSX.utils.aoa_to_sheet(wsData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                mainSheet = ws;
                dataSheet = ws;
            }

            // Convert to array of arrays
            const mainData = XLSX.utils.sheet_to_json(mainSheet, { header: 1, defval: '' });
            const dataData = XLSX.utils.sheet_to_json(dataSheet, { header: 1, defval: '' });

            console.log(`📊 Main data rows: ${mainData.length}, Data sheet rows: ${dataData.length}`);

            // === STEP 2: Find SKU Column ===
            let skuCol = -1;
            let skuRow = -1;
            
            for (let r = 0; r < Math.min(mainData.length, 100); r++) {
                const row = mainData[r];
                if (!row) continue;
                for (let c = 0; c < row.length; c++) {
                    const cellValue = String(row[c] || '').trim().toUpperCase();
                    if (cellValue === skuHeader.toUpperCase()) {
                        skuCol = c;
                        skuRow = r;
                        console.log(`✅ Found SKU at row ${r}, column ${c}: "${row[c]}"`);
                        break;
                    }
                }
                if (skuCol >= 0) break;
            }

            if (skuCol < 0) {
                reject(new Error(`❌ Could not find '${skuHeader}' column header. Please check your file.`));
                return;
            }

            // === STEP 3: Map All Columns ===
            // Based on your CSV: SKU, Product Name, Supplier SKU, Units Per Case, 
            // Total Ordered Cases, Ordered Quantity, Confirmed Quantity, 
            // Total Received Cases, Received Quantity, Unit Cost, Internal Tax, 
            // VAT, Gross Unit Cost, Cost, Net Discounted Cost
            
            const colMap = {
                sku: skuCol,
                productName: skuCol + 1,
                supplierSku: skuCol + 2,
                unitsPerCase: skuCol + 3,
                totalOrderedCases: skuCol + 4,
                orderedQty: skuCol + 5,
                confirmedQty: skuCol + 6,
                totalReceivedCases: skuCol + 7,
                receivedQty: skuCol + 8,
                unitCost: skuCol + 9,
                internalTax: skuCol + 10,
                vat: skuCol + 11,
                grossUnitCost: skuCol + 12,
                cost: skuCol + 13,
                netDiscountedCost: skuCol + 14
            };

            console.log(`📋 Column mapping: SKU=${skuCol}, N.D.Cost=${colMap.netDiscountedCost}`);

            // === STEP 4: Find Last Row with Data ===
            let lastRow = skuRow;
            for (let r = skuRow + 1; r < mainData.length; r++) {
                const row = mainData[r];
                if (!row) continue;
                const val = row[skuCol];
                if (val !== undefined && String(val).trim() !== '' && 
                    val !== 'SKU' && val !== 'sku' && !String(val).includes('TOTAL')) {
                    lastRow = r;
                }
            }

            console.log(`📊 Data rows found: from row ${skuRow + 1} to ${lastRow}`);

            if (lastRow <= skuRow) {
                reject(new Error('❌ No data found under SKU column. Please check your file.'));
                return;
            }

            // === STEP 5: Build VLOOKUP Data from Data Sheet ===
            // Looking for SKU in column K (index 10) and returning column AA (index 26)
            // Also getting Barcode from column O (index 14)
            
            const lookupData = {};
            console.log('🔍 Building VLOOKUP data from data sheet...');
            
            // Find SKU column in data sheet (try column K = index 10)
            let dataSkuCol = 10; // Default: K
            let dataReturnCol = 26; // Default: AA (for VLOOKUP value)
            let dataBarcodeCol = 14; // Default: O (for Barcode)
            
            // Try to find the SKU column in data sheet
            for (let r = 0; r < Math.min(dataData.length, 5); r++) {
                const row = dataData[r];
                if (!row) continue;
                for (let c = 0; c < Math.min(row.length, 30); c++) {
                    const cellValue = String(row[c] || '').trim().toUpperCase();
                    if (cellValue === 'SKU') {
                        dataSkuCol = c;
                        console.log(`✅ Found SKU in data sheet at column ${c}`);
                        break;
                    }
                }
                if (dataSkuCol !== 10) break;
            }

            // Build lookup data
            for (let r = 1; r < dataData.length; r++) {
                const row = dataData[r];
                if (!row) continue;
                if (row[dataSkuCol] !== undefined && row[dataSkuCol] !== '' && row[dataSkuCol] !== null) {
                    const key = String(row[dataSkuCol]).trim();
                    if (key) {
                        lookupData[key] = {
                            row: row,
                            index: r,
                            // Return column AA (index 26) or whatever is at dataReturnCol
                            lookupVal: parseFloat(row[dataReturnCol]) || 0,
                            // Barcode from column O (index 14)
                            barcode: String(row[dataBarcodeCol] || '').trim()
                        };
                    }
                }
            }

            console.log(`🔍 Found ${Object.keys(lookupData).length} records in lookup data`);

            // === STEP 6: Process Each Row ===
            const results = [];
            let totalRCost = 0;
            let processedCount = 0;

            console.log('📝 Processing rows...');

            for (let r = skuRow + 1; r <= lastRow; r++) {
                const row = mainData[r];
                if (!row) continue;

                const sku = String(row[skuCol] || '').trim();
                if (!sku || sku.toUpperCase() === 'SKU' || sku.toUpperCase().includes('TOTAL')) {
                    continue;
                }

                // Get all values
                const productName = String(row[colMap.productName] || '').trim();
                const unitsPerCase = parseFloat(String(row[colMap.unitsPerCase] || '0').replace(/,/g, '')) || 0;
                const totalOrderedCases = parseFloat(String(row[colMap.totalOrderedCases] || '0').replace(/,/g, '')) || 0;
                const orderedQty = parseFloat(String(row[colMap.orderedQty] || '0').replace(/,/g, '')) || 0;
                const grossUnitCost = parseFloat(String(row[colMap.grossUnitCost] || '0').replace(/,/g, '')) || 0;
                const cost = parseFloat(String(row[colMap.cost] || '0').replace(/,/g, '')) || 0;
                const netDiscountedCost = parseFloat(String(row[colMap.netDiscountedCost] || '0').replace(/,/g, '')) || 0;

                // === VLOOKUP: Get value from data sheet ===
                // Formula: =N.D.Cost * 1.05 * VLOOKUP(SKU, K:AA, 17, 0)
                let lookupVal = 1; // Default to 1 if no lookup found
                let barcode = '';
                
                if (lookupData[sku]) {
                    lookupVal = lookupData[sku].lookupVal || 1;
                    barcode = lookupData[sku].barcode;
                    console.log(`  ✓ SKU ${sku}: lookupVal=${lookupVal}, barcode=${barcode}`);
                } else {
                    console.log(`  ⚠️ SKU ${sku}: No lookup data found, using default`);
                }

                // R.Cost = N.D. Cost × 1.05 × VLOOKUP
                const rCost = netDiscountedCost * 1.05 * lookupVal;

                // Format barcode as array
                let formattedBarcode = '';
                if (barcode) {
                    // Check if multiple barcodes (comma separated)
                    if (barcode.includes(',')) {
                        const parts = barcode.split(',').map(p => p.trim());
                        formattedBarcode = `['${parts.join("', '")}']`;
                    } else {
                        formattedBarcode = `['${barcode}']`;
                    }
                }

                // Build result row
                results.push({
                    'SKU': sku,
                    'Product Name': productName || sku,
                    'Units Per Case': unitsPerCase,
                    'Total Ordered Cases': totalOrderedCases,
                    'Ordered Quantity': orderedQty,
                    'G.U.Cost': grossUnitCost,
                    'Cost': cost,
                    'N.D.Cost': netDiscountedCost,
                    'R.Cost': parseFloat(rCost.toFixed(3)),
                    'Barcode': formattedBarcode
                });

                totalRCost += rCost;
                processedCount++;
            }

            console.log(`✅ Processed ${processedCount} rows`);
            console.log(`💰 Total R.Cost: ${totalRCost.toFixed(3)}`);

            // === STEP 7: Add Total Row ===
            if (results.length > 0) {
                // Add separator row
                const separatorRow = {};
                Object.keys(results[0]).forEach(key => {
                    separatorRow[key] = key === 'SKU' ? '═══════════' : '';
                });
                results.push(separatorRow);

                // Add total row
                const totalRow = {};
                Object.keys(results[0]).forEach(key => {
                    if (key === 'SKU') {
                        totalRow[key] = 'TOTAL SUM:';
                    } else if (key === 'R.Cost') {
                        totalRow[key] = parseFloat(totalRCost.toFixed(3));
                    } else {
                        totalRow[key] = '';
                    }
                });
                results.push(totalRow);
            }

            // === STEP 8: Return Results ===
            const finalTotal = parseFloat(totalRCost.toFixed(2));
            
            resolve({
                data: results,
                summary: {
                    'Total SUM (R.Cost)': finalTotal,
                    'Rows Processed': processedCount,
                    'Total SKUs': processedCount
                },
                message: `✅ Process completed successfully!\n\n` +
                         `══════════════════════════════════\n` +
                         ` FINAL TOTAL AMOUNT (R.Cost): ${finalTotal.toFixed(2)}\n` +
                         `══════════════════════════════════\n` +
                         `Rows processed: ${processedCount}\n` +
                         `Total SKUs: ${processedCount}\n` +
                         `Formula: N.D. Cost × 1.05 × VLOOKUP`,
                total: finalTotal
            });

            console.log('🎉 Calculate Receivings completed successfully!');

        } catch (err) {
            console.error('❌ Error in Calculate Receivings:', err);
            reject(err);
        }
    });
}