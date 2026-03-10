/**
 * River & Sediment Discharge Data Processor
 * Pure HTML/CSS/JavaScript - works on GitHub Pages (no backend)
 */

const METADATA_NAMES = ['Metadata - River Water Discharge', 'Metadata - Suspended Sediment', 'Metadata', 'Metadata - River Water Discharg', 'Metadata - Suspended Sediment_T'];

function findDataSheet(wb) {
    const dataSheet = wb.SheetNames.find(n => {
        const lower = n.toLowerCase();
        return !lower.startsWith('metadata') && !lower.includes('metadata -');
    });
    return dataSheet || wb.SheetNames[0];
}

function sheetToArray(ws, skipRows = 0) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const data = [];
    for (let R = skipRows; R <= range.e.r; R++) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
            row.push(cell ? cell.v : undefined);
        }
        data.push(row);
    }
    return data;
}

function findHeaderRow(ws) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let R = 0; R <= Math.min(20, range.e.r); R++) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
            row.push(cell ? String(cell.v || '').toLowerCase() : '');
        }
        const hasDataValue = row.some(c => c.includes('data value') || c.includes('datavalue'));
        const hasDateCol = row.some(c => c.includes('data date') || c.includes('data time') || (c.includes('date') && c.includes('time')));
        if (hasDataValue && hasDateCol) return R;
        if (hasDataValue && row.some(c => c.includes('date'))) return R;
    }
    return 0;
}

function arrayToSheet(data) {
    return XLSX.utils.aoa_to_sheet(data);
}

const BORDER_STYLE = { style: 'thin', color: { rgb: 'FF94A3B8' } };

function applyTableStyles(ws) {
    if (!ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const border = { top: BORDER_STYLE, bottom: BORDER_STYLE, left: BORDER_STYLE, right: BORDER_STYLE };
    for (let R = 0; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[addr]) continue;
            if (R === 0) {
                ws[addr].s = {
                    font: { bold: true, sz: 12, color: { rgb: 'FFFFFFFF' } },
                    fill: { patternType: 'solid', fgColor: { rgb: 'FF2563EB' }, bgColor: { indexed: 64 } },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border
                };
            } else {
                ws[addr].s = {
                    alignment: { vertical: 'center' },
                    border
                };
            }
        }
    }
    if (!ws['!cols']) ws['!cols'] = [];
    const colWidths = [14, 28, 22, 14, 10];
    for (let C = 0; C <= range.e.c; C++) {
        ws['!cols'][C] = { wch: colWidths[C] || 14 };
    }
}

function parseDateTime(val) {
    if (val == null || val === '') return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    if (typeof val === 'number') {
        if (val > 100000 || val < 1) return null;
        const ms = (val - 25569) * 86400 * 1000;
        return new Date(ms);
    }
    const s = String(val).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2}):(\d{2})/);
    if (m) {
        return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
    }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)/i);
    if (m) {
        let h = parseInt(m[4], 10);
        if (m[7].toUpperCase() === 'PM' && h !== 12) h += 12;
        if (m[7].toUpperCase() === 'AM' && h === 12) h = 0;
        return new Date(+m[3], parseInt(m[1], 10) - 1, parseInt(m[2], 10), h, +m[5], +m[6]);
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
}

function formatDateTime(dt) {
    if (!dt || !(dt instanceof Date) || isNaN(dt.getTime())) return '';
    const m = dt.getMonth() + 1, d = dt.getDate(), y = dt.getFullYear();
    let h = dt.getHours(), min = dt.getMinutes(), sec = dt.getSeconds();
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    const ts = `${h}:${String(min).padStart(2,'0')}:${String(sec).padStart(2,'0')} ${ampm}`;
    return `${m}/${d}/${y} ${ts}`;
}

function getDateOnly(dt) {
    if (!dt) return null;
    return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
}

function processExcelData(rawData) {
    if (!rawData || rawData.length < 1) return null;
    const headers = rawData[0];
    const rows = rawData.slice(1) || [];

    let dateColIdx = -1, valueColIdx = -1;
    for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i] || '').toLowerCase().replace(/\s/g, '');
        const h2 = String(headers[i] || '').toLowerCase();
        if (h2.includes('data date') || h2.includes('data time') || (h2.includes('date') && h2.includes('time')) || h.includes('datetime')) dateColIdx = i;
        if ((h2.includes('data') && h2.includes('value')) || h.includes('datavalue')) valueColIdx = i;
    }
    if (dateColIdx < 0) dateColIdx = headers.findIndex(h => String(h||'').toLowerCase().includes('date'));
    if (dateColIdx < 0) dateColIdx = 0;
    if (valueColIdx < 0) valueColIdx = headers.findIndex(h => String(h||'').toLowerCase().includes('value'));
    if (valueColIdx < 0) valueColIdx = headers.length - 1;

    let data = rows.map(row => {
        const r = [...row];
        while (r.length < headers.length) r.push(undefined);
        return r;
    });

    data = data.filter(row => {
        const dt = parseDateTime(row[dateColIdx]);
        return dt != null;
    });

    data = data.map(row => {
        const dt = parseDateTime(row[dateColIdx]);
        return { row, dt, dateKey: getDateOnly(dt) };
    });

    data.sort((a, b) => b.dt - a.dt);

    data = data.filter(d => {
        const v = parseFloat(d.row[valueColIdx]);
        return !isNaN(v) && v !== 0;
    });

    data.sort((a, b) => a.dt - b.dt);
    const groupedByDate = {};
    data.forEach(d => {
        if (!groupedByDate[d.dateKey]) groupedByDate[d.dateKey] = [];
        groupedByDate[d.dateKey].push(d);
    });
    const multipleReadings = [];
    Object.entries(groupedByDate).forEach(([dateKey, items]) => {
        if (items.length > 1) {
            items.sort((a, b) => a.dt - b.dt);
            multipleReadings.push({
                dateKey,
                readings: items.map((it, i) => ({
                    time: formatDateTime(it.dt),
                    value: it.row[valueColIdx],
                    isTaken: i === 0
                })),
                count: items.length
            });
        }
    });
    const seen = new Set();
    data = data.filter(d => {
        if (seen.has(d.dateKey)) return false;
        seen.add(d.dateKey);
        return true;
    });

    data.sort((a, b) => b.dt - a.dt);

    data = data.map(d => {
        const row = [...d.row];
        row[dateColIdx] = formatDateTime(d.dt);
        return row;
    });

    return {
        data: [headers, ...data],
        dateColIdx,
        valueColIdx,
        headers,
        multipleReadings
    };
}

function loadAndProcessFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
                const sheetName = findDataSheet(wb);
                const ws = wb.Sheets[sheetName];

                const headerRow = findHeaderRow(ws);
                const rawData = sheetToArray(ws, headerRow);
                const result = processExcelData(rawData);
                if (!result) {
                    reject('No valid data found in sheet');
                    return;
                }
                resolve({ result, workbook: wb, sheetName });
            } catch (err) {
                reject(err.message || String(err));
            }
        };
        reader.onerror = () => reject('Failed to read file');
        reader.readAsArrayBuffer(file);
    });
}

function matchAndRemove(riverResult, sedimentResult) {
    const riverData = riverResult.data;
    const riverHeaders = riverResult.headers;
    const riverDateIdx = riverResult.dateColIdx;
    const sedimentData = sedimentResult.data;
    const sedimentDateIdx = sedimentResult.dateColIdx;

    const riverRows = riverData.slice(1);
    const sedimentRows = sedimentData.slice(1);

    const riverDates = new Set(riverRows.map(r => getDateOnly(parseDateTime(r[riverDateIdx]))).filter(Boolean));
    const sedimentDates = new Set(sedimentRows.map(r => getDateOnly(parseDateTime(r[sedimentDateIdx]))).filter(Boolean));

    const commonDates = new Set([...riverDates].filter(d => sedimentDates.has(d)));
    const riverOnlyDates = [...riverDates].filter(d => !sedimentDates.has(d)).sort();
    const sedimentOnlyDates = [...sedimentDates].filter(d => !riverDates.has(d)).sort();

    const riverMatched = [riverHeaders, ...riverRows.filter(r => commonDates.has(getDateOnly(parseDateTime(r[riverDateIdx]))))];
    const sedimentMatched = [sedimentResult.headers, ...sedimentRows.filter(r => commonDates.has(getDateOnly(parseDateTime(r[sedimentDateIdx]))))];

    const riverDeleted = [riverHeaders, ...riverRows.filter(r => !commonDates.has(getDateOnly(parseDateTime(r[riverDateIdx]))))];
    const sedimentDeleted = [sedimentResult.headers, ...sedimentRows.filter(r => !commonDates.has(getDateOnly(parseDateTime(r[sedimentDateIdx]))))];

    return { riverMatched, sedimentMatched, riverDeleted, sedimentDeleted, riverOnlyDates, sedimentOnlyDates };
}

function formatDateDisplay(dateKey) {
    if (!dateKey) return '';
    const [y, m, d] = dateKey.split('-');
    return `${m}/${d}/${y}`;
}

function generatePDFReport(reportData) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    let y = 15;

    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    doc.text('Data Cleaning Report', 105, y, { align: 'center' });
    y += 12;

    doc.setFontSize(10);
    doc.setFont(undefined, 'normal');

    const tableHead = [['S.No', 'Date', 'Time', 'Data Value', 'Taken']];
    const riverMulti = reportData.riverMultipleReadings || [];
    if (riverMulti.length > 0) {
        doc.setFont(undefined, 'bold');
        doc.text('1. River Discharge - Multiple Readings on Same Day', 14, y);
        y += 6;
        doc.setFont(undefined, 'normal');
        riverMulti.forEach((item, idx) => {
            const body = item.readings.map((r, i) => [
                (i === 0 ? idx + 1 : ''),
                (i === 0 ? formatDateDisplay(item.dateKey) : ''),
                r.time,
                r.value,
                r.isTaken ? 'Yes' : ''
            ]);
            doc.autoTable({
                startY: y,
                head: tableHead,
                body,
                theme: 'grid',
                headStyles: { fillColor: [37, 99, 235], textColor: 255 },
                columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 28 }, 2: { cellWidth: 45 }, 3: { cellWidth: 30 }, 4: { cellWidth: 20 } },
                didParseCell: (data) => {
                    if (data.section === 'body' && body[data.row.index] && body[data.row.index][4] === 'Yes') {
                        data.cell.styles.textColor = [220, 38, 38];
                        data.cell.styles.fontStyle = 'bold';
                    }
                }
            });
            y = doc.lastAutoTable.finalY + 2;
            doc.setFontSize(9);
            doc.text(`Count of readings: ${item.count}`, 14, y);
            y += 8;
        });
        y += 5;
    }

    const sedimentMulti = reportData.sedimentMultipleReadings || [];
    if (sedimentMulti.length > 0) {
        if (y > 250) { doc.addPage(); y = 15; }
        doc.setFont(undefined, 'bold');
        doc.text('2. Sediment - Multiple Readings on Same Day', 14, y);
        y += 6;
        doc.setFont(undefined, 'normal');
        sedimentMulti.forEach((item, idx) => {
            const body = item.readings.map((r, i) => [
                (i === 0 ? idx + 1 : ''),
                (i === 0 ? formatDateDisplay(item.dateKey) : ''),
                r.time,
                r.value,
                r.isTaken ? 'Yes' : ''
            ]);
            doc.autoTable({
                startY: y,
                head: tableHead,
                body,
                theme: 'grid',
                headStyles: { fillColor: [37, 99, 235], textColor: 255 },
                columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 28 }, 2: { cellWidth: 45 }, 3: { cellWidth: 30 }, 4: { cellWidth: 20 } },
                didParseCell: (data) => {
                    if (data.section === 'body' && body[data.row.index] && body[data.row.index][4] === 'Yes') {
                        data.cell.styles.textColor = [220, 38, 38];
                        data.cell.styles.fontStyle = 'bold';
                    }
                }
            });
            y = doc.lastAutoTable.finalY + 2;
            doc.setFontSize(9);
            doc.text(`Count of readings: ${item.count}`, 14, y);
            y += 8;
        });
        y += 5;
    }

    const riverOnlyDates = reportData.riverOnlyDates || [];
    if (riverOnlyDates.length > 0) {
        if (y > 250) { doc.addPage(); y = 15; }
        doc.setFont(undefined, 'bold');
        doc.text('3. Dates: River Data Exists but Sediment Does Not', 14, y);
        y += 6;
        doc.setFont(undefined, 'normal');
        const body = riverOnlyDates.map((d, i) => [i + 1, formatDateDisplay(d)]);
        doc.autoTable({
            startY: y,
            head: [['S.No', 'Date']],
            body,
            theme: 'grid',
            headStyles: { fillColor: [37, 99, 235], textColor: 255 },
            columnStyles: { 0: { cellWidth: 25 }, 1: { cellWidth: 50 } }
        });
        y = doc.lastAutoTable.finalY + 10;
    }

    const sedimentOnlyDates = reportData.sedimentOnlyDates || [];
    if (sedimentOnlyDates.length > 0) {
        if (y > 250) { doc.addPage(); y = 15; }
        doc.setFont(undefined, 'bold');
        doc.text('4. Dates: Sediment Data Exists but River Does Not', 14, y);
        y += 6;
        doc.setFont(undefined, 'normal');
        const body = sedimentOnlyDates.map((d, i) => [i + 1, formatDateDisplay(d)]);
        doc.autoTable({
            startY: y,
            head: [['S.No', 'Date']],
            body,
            theme: 'grid',
            headStyles: { fillColor: [37, 99, 235], textColor: 255 },
            columnStyles: { 0: { cellWidth: 25 }, 1: { cellWidth: 50 } }
        });
    }

    if (riverMulti.length === 0 && sedimentMulti.length === 0 && riverOnlyDates.length === 0 && sedimentOnlyDates.length === 0) {
        doc.text('No multiple readings or unmatched dates found.', 14, y);
    }

    doc.save('Data Cleaning report.pdf');
}

function downloadExcel(data, filename, originalWb, sheetName) {
    const ws = arrayToSheet(data);
    applyTableStyles(ws);
    const sheetTitle = filename.replace('.xlsx', '').slice(0, 31);
    let wb;
    if (originalWb && sheetName) {
        wb = XLSX.utils.book_new();
        for (const name of originalWb.SheetNames) {
            const lower = name.toLowerCase();
            if (lower.startsWith('metadata') || lower.includes('metadata -')) continue;
            if (name === sheetName) {
                XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
            } else {
                XLSX.utils.book_append_sheet(wb, originalWb.Sheets[name], name);
            }
        }
        if (!wb.SheetNames.length) XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
    } else {
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
    }
    XLSX.writeFile(wb, filename);
}

document.addEventListener('DOMContentLoaded', () => {
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        const savedTheme = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
        document.documentElement.setAttribute('data-theme', savedTheme);
        themeToggle.addEventListener('click', () => {
            const current = document.documentElement.getAttribute('data-theme');
            const next = current === 'dark' ? 'light' : 'dark';
            document.documentElement.setAttribute('data-theme', next);
            localStorage.setItem('theme', next);
        });
    }

    const form = document.getElementById('uploadForm');
    const riverInput = document.getElementById('riverFile');
    const sedimentInput = document.getElementById('sedimentFile');
    const riverFileName = document.getElementById('riverFileName');
    const sedimentFileName = document.getElementById('sedimentFileName');
    const errorMessage = document.getElementById('errorMessage');
    const successMessage = document.getElementById('successMessage');
    const submitBtn = document.getElementById('submitBtn');
    const downloadSection = document.getElementById('downloadSection');

    riverInput.addEventListener('change', () => {
        riverFileName.textContent = riverInput.files[0]?.name || 'No file chosen';
    });
    sedimentInput.addEventListener('change', () => {
        sedimentFileName.textContent = sedimentInput.files[0]?.name || 'No file chosen';
    });

    let lastResults = null;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        errorMessage.style.display = 'none';
        successMessage.style.display = 'none';
        downloadSection.style.display = 'none';

        const riverFile = riverInput.files[0];
        const sedimentFile = sedimentInput.files[0];

        if (!riverFile || !sedimentFile) {
            errorMessage.textContent = 'Please select both River and Sediment files.';
            errorMessage.style.display = 'block';
            return;
        }

        submitBtn.disabled = true;
        submitBtn.textContent = 'Processing...';

        try {
            const [riverLoad, sedimentLoad] = await Promise.all([
                loadAndProcessFile(riverFile),
                loadAndProcessFile(sedimentFile)
            ]);

            const matchResult = matchAndRemove(riverLoad.result, sedimentLoad.result);
            const { riverMatched, sedimentMatched, riverDeleted, sedimentDeleted } = matchResult;

            lastResults = {
                riverMatched,
                sedimentMatched,
                riverDeleted,
                sedimentDeleted,
                riverWb: riverLoad.workbook,
                riverSheet: riverLoad.sheetName,
                sedimentWb: sedimentLoad.workbook,
                sedimentSheet: sedimentLoad.sheetName,
                riverMultipleReadings: riverLoad.result.multipleReadings || [],
                sedimentMultipleReadings: sedimentLoad.result.multipleReadings || [],
                riverOnlyDates: matchResult.riverOnlyDates || [],
                sedimentOnlyDates: matchResult.sedimentOnlyDates || []
            };

            successMessage.textContent = 'Processing complete!';
            successMessage.style.display = 'block';
            downloadSection.style.display = 'block';
        } catch (err) {
            errorMessage.textContent = err || 'An error occurred during processing.';
            errorMessage.style.display = 'block';
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = 'Process & Download';
        }
    });

    document.getElementById('downloadRiver').addEventListener('click', () => {
        if (lastResults) downloadExcel(lastResults.riverMatched, 'River_Water_Discharge_Updated.xlsx', lastResults.riverWb, lastResults.riverSheet);
    });
    document.getElementById('downloadSediment').addEventListener('click', () => {
        if (lastResults) downloadExcel(lastResults.sedimentMatched, 'Suspended_Sediment_Updated.xlsx', lastResults.sedimentWb, lastResults.sedimentSheet);
    });
    document.getElementById('downloadRiverDeleted').addEventListener('click', () => {
        if (lastResults) downloadExcel(lastResults.riverDeleted, 'River_Discharge_Deleted.xlsx', lastResults.riverWb, lastResults.riverSheet);
    });
    document.getElementById('downloadSedimentDeleted').addEventListener('click', () => {
        if (lastResults) downloadExcel(lastResults.sedimentDeleted, 'Sediment_Deleted.xlsx', lastResults.sedimentWb, lastResults.sedimentSheet);
    });
    document.getElementById('downloadReport').addEventListener('click', () => {
        if (lastResults) {
            generatePDFReport({
                riverMultipleReadings: lastResults.riverMultipleReadings,
                sedimentMultipleReadings: lastResults.sedimentMultipleReadings,
                riverOnlyDates: lastResults.riverOnlyDates,
                sedimentOnlyDates: lastResults.sedimentOnlyDates
            });
        }
    });
});
