let workbookData = [];
let imageMap = {}; 
const overlay = document.getElementById('loadingOverlay');

document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    overlay.style.display = 'flex';

    setTimeout(async () => {
        try {
            const buffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.worksheets[0];

            workbookData = [];
            imageMap = {};

            const headerRow = worksheet.getRow(1);
            const selects = [document.getElementById('a1_col'), document.getElementById('a2_col'), document.getElementById('a3_col')];
            
            selects.forEach(s => {
                s.innerHTML = '<option value="">(None)</option>';
                headerRow.eachCell((cell, colNum) => {
                    s.add(new Option(cell.text || `Col ${colNum}`, colNum));
                });
            });

            worksheet.eachRow((row, rowNum) => {
                if (rowNum === 1) return;
                let obj = { rNum: rowNum, vals: {} };
                for(let i = 1; i <= headerRow.cellCount; i++) {
                    obj.vals[i] = row.getCell(i).text;
                }
                workbookData.push(obj);
            });

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
            alert("Error loading .xlsx file."); 
        }
        overlay.style.display = 'none';
    }, 100);
});

function drawAsset(doc, num, rowObj, curX, curY) {
    const col = document.getElementById(`a${num}_col`).value;
    if (!col) return;

    const type = document.querySelector(`input[name="a${num}_type"]:checked`).value;
    const xOff = parseFloat(document.getElementById(`a${num}_x`).value) || 0;
    const yOff = parseFloat(document.getElementById(`a${num}_y`).value) || 0;
    const w = parseFloat(document.getElementById(`a${num}_w`).value) || 0;
    const h = parseFloat(document.getElementById(`a${num}_h`).value) || 0;
    const fSize = parseFloat(document.getElementById(`a${num}_f`).value) || 10; // Independent font size

    const targetX = curX + xOff;
    const targetY = curY + yOff;

    if (type === 'text') {
        const txt = String(rowObj.vals[col] || "");
        doc.setFontSize(fSize);
        // Correct Y-offset calculation for text baseline
        doc.text(txt, targetX, targetY + (fSize * 0.35));
    } else {
        const img = imageMap[`${rowObj.rNum}-${col}`];
        if (img) {
            try {
                doc.addImage(img.data, img.ext, targetX, targetY, w, h);
            } catch (e) { console.warn("Image render failed", e); }
        }
    }
}

function createPDF(previewMode = false) {
    const start = parseInt(document.getElementById('rowStart').value) - 1;
    const end = parseInt(document.getElementById('rowEnd').value);
    const lW = parseFloat(document.getElementById('lWidth').value);
    const lH = parseFloat(document.getElementById('lHeight').value);
    const gH = parseFloat(document.getElementById('hGap').value);
    const gV = parseFloat(document.getElementById('vGap').value);
    const mLeft = parseFloat(document.getElementById('mLeft').value);
    const mTop = parseFloat(document.getElementById('mTop').value);
    const reps = parseInt(document.getElementById('repeats').value);

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    let curX = mLeft;
    let curY = mTop;

    const rows = workbookData.slice(Math.max(0, start), end);

    rows.forEach((row) => {
        for (let i = 0; i < reps; i++) {
            if (curX + lW > 208) { curX = mLeft; curY += (lH + gV); }
            if (curY + lH > 275) {
                if (previewMode) return;
                doc.addPage(); curX = mLeft; curY = mTop;
            }

            // Draw Sticker Border
            doc.setLineWidth(0.05); doc.setDrawColor(220);
            doc.rect(curX, curY, lW, lH);

            // Draw Assets
            drawAsset(doc, 1, row, curX, curY);
            drawAsset(doc, 2, row, curX, curY);
            drawAsset(doc, 3, row, curX, curY);

            curX += (lW + gH);
        }
    });

    // Calibration Line
    doc.setPage(1);
    doc.setDrawColor(0); doc.setLineWidth(0.3);
    doc.line(10, 285, 110, 285);
    doc.setFontSize(7);
    doc.text("TRUE SIZE CHECK: LINE MUST BE 100mm", 10, 283);

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
    doc.save('labels_custom_layout.pdf');
});
