let workbookData = [];
let imageMap = {}; 

const overlay = document.getElementById('loadingOverlay');

// Show/Hide Setting groups based on Content Type
document.querySelectorAll('input[name="contentType"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        const isText = e.target.value === 'text';
        document.getElementById('fontSizeGroup').style.display = isText ? 'block' : 'none';
        document.getElementById('alignmentSection').style.display = isText ? 'block' : 'none';
        updatePreview();
    });
});

document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    // 1. Show Loading Overlay
    overlay.style.display = 'flex';

    // Small delay to let the browser show the spinner before the heavy CPU work begins
    setTimeout(async () => {
        try {
            const buffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.worksheets[0];

            workbookData = [];
            imageMap = {};

            // 2. Map Columns
            const headerRow = worksheet.getRow(1);
            const select = document.getElementById('columnSelect');
            select.innerHTML = '';
            headerRow.eachCell((cell, colNum) => {
                select.add(new Option(cell.text || `Col ${colNum}`, colNum));
            });

            // 3. Extract Rows
            worksheet.eachRow((row, rowNum) => {
                if (rowNum === 1) return;
                let obj = { rNum: rowNum, vals: {} };
                row.eachCell({ includeEmpty: true }, (cell, colNum) => {
                    obj.vals[colNum] = cell.text;
                });
                workbookData.push(obj);
            });

            // 4. Extract Images
            worksheet.getImages().forEach(img => {
                const m = workbook.model.media.find(med => med.index === img.imageId);
                const r = Math.floor(img.range.tl.row) + 1;
                const c = Math.floor(img.range.tl.col) + 1;
                imageMap[`${r}-${c}`] = {
                    data: `data:image/${m.extension};base64,${m.buffer.toString('base64')}`,
                    ext: m.extension.toUpperCase()
                };
            });

            document.getElementById('rowEnd').value = workbookData.length;
            document.getElementById('configArea').style.display = 'block';
            updatePreview();

        } catch (err) {
            console.error(err);
            alert("Failed to load Excel file. Ensure it is a valid .xlsx file.");
        } finally {
            // 5. Hide Loading Overlay
            overlay.style.display = 'none';
        }
    }, 100);
});

function createPDF(previewMode = false) {
    const col = document.getElementById('columnSelect').value;
    const type = document.querySelector('input[name="contentType"]:checked').value;
    const start = parseInt(document.getElementById('rowStart').value) - 1;
    const end = parseInt(document.getElementById('rowEnd').value);
    
    const lW = parseFloat(document.getElementById('lWidth').value);
    const lH = parseFloat(document.getElementById('lHeight').value);
    const gH = parseFloat(document.getElementById('hGap').value);
    const gV = parseFloat(document.getElementById('vGap').value);
    const mLeft = parseFloat(document.getElementById('mLeft').value);
    const mTop = parseFloat(document.getElementById('mTop').value);
    const fS = parseFloat(document.getElementById('fSize').value);
    const reps = parseInt(document.getElementById('repeats').value);
    const hA = document.getElementById('hAlign').value;
    const vA = document.getElementById('vAlign').value;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    let curX = mLeft;
    let curY = mTop;

    const rows = workbookData.slice(Math.max(0, start), end);

    rows.forEach((row) => {
        for (let i = 0; i < reps; i++) {
            if (curX + lW > 210 - 2) { curX = mLeft; curY += (lH + gV); }
            if (curY + lH > 275) {
                if (previewMode) return; 
                doc.addPage(); curX = mLeft; curY = mTop;
            }

            doc.setLineWidth(0.05); doc.setDrawColor(200);
            doc.rect(curX, curY, lW, lH);

            if (type === 'text') {
                const txt = String(row.vals[col] || "");
                doc.setFontSize(fS);
                const lines = doc.splitTextToSize(txt, lW - 1);
                const tH = lines.length * (fS * 0.3528);
                let tx = curX + 0.5;
                if (hA === 'center') tx = curX + (lW / 2);
                if (hA === 'right') tx = curX + lW - 0.5;
                let ty = curY + (fS * 0.3528);
                if (vA === 'middle') ty = curY + (lH / 2) - (tH / 2) + (fS * 0.25);
                if (vA === 'bottom') ty = curY + lH - tH + (fS * 0.25);
                doc.text(lines, tx, ty, { align: hA });
            } else {
                const img = imageMap[`${row.rNum}-${col}`];
                if (img) {
                    const pad = 0.2;
                    doc.addImage(img.data, img.ext, curX + pad, curY + pad, lW - (pad*2), lH - (pad*2));
                }
            }
            curX += (lW + gH);
        }
    });

    // Calibration Line
    doc.setPage(1);
    doc.setDrawColor(0); doc.setLineWidth(0.3);
    doc.line(10, 285, 110, 285);
    doc.setFontSize(7);
    doc.text("TRUE SIZE CHECK: This line must be exactly 100mm long after printing.", 10, 283);

    return doc;
}

function updatePreview() {
    if (workbookData.length === 0) return;
    const doc = createPDF(true);
    document.getElementById('pdfPreview').src = doc.output('bloburl');
}

document.getElementById('previewBtn').addEventListener('click', updatePreview);
document.getElementById('generateBtn').addEventListener('click', () => {
    const doc = createPDF(false);
    doc.save('labels.pdf');
});