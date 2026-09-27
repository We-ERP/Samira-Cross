// متغيرات تخزين الداتا
let ftthData = null;
let dslData = null;
let irData = null;
let finalCrossMatchData = [];

// تعريف الـ Event Listeners للملفات الثلاثة
document.getElementById('file-ftth').addEventListener('change', (e) => loadFile(e, 'ftth'));
document.getElementById('file-dsl').addEventListener('change', (e) => loadFile(e, 'dsl'));
document.getElementById('file-ir').addEventListener('change', (e) => loadFile(e, 'ir'));

function loadFile(event, type) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        
        const statusBadge = document.getElementById(`${type}-status`);
        statusBadge.innerText = `Loaded (${jsonData.length} rows)`;
        statusBadge.classList.add('loaded');

        if (type === 'ftth') ftthData = jsonData;
        if (type === 'dsl') dslData = jsonData;
        if (type === 'ir') irData = jsonData;
    };
    reader.readAsArrayBuffer(file);
}

// دالة لتنظيف السيرفيس نمبر من FBB و FV ومسافات الفراغ
function cleanServiceNumber(val) {
    if (!val) return "";
    let str = String(val).toUpperCase();
    return str.replace(/FBB/g, '').replace(/FV/g, '').trim();
}

// دالة المطابقة
document.getElementById('btn-run').addEventListener('click', () => {
    if (!ftthData || !dslData || !irData) {
        alert("Please upload all 3 sheets (FTTH, DSL, IR) before running.");
        return;
    }

    // بناء Maps لسرعة البحث الخارقة (O(1) Lookup) بدلاً من اللف داخل اللف
    const dslMap = new Map();
    dslData.forEach(row => {
        const num = cleanServiceNumber(row['Service Number']);
        if (num) dslMap.set(num, row);
    });

    const ftthMap = new Map();
    ftthData.forEach(row => {
        const num = cleanServiceNumber(row['Service Number']);
        if (num) ftthMap.set(num, row);
    });

    finalCrossMatchData = [];

    // عملية المطابقة على شيت IR
    irData.forEach(irRow => {
        const newRow = { ...irRow }; // نسخ صف الـ IR كما هو
        const irNum = cleanServiceNumber(irRow['dsl_number']);
        const irCategory = String(irRow['case_sub_type'] || '').trim().toLowerCase(); 

        // 1. مطابقة الـ DSL
        const dslRow = dslMap.get(irNum);
        if (dslRow) {
            const dslServiceOrigin = String(dslRow['Service Name'] || '').trim();
            const dslServiceLower = dslServiceOrigin.toLowerCase();
            newRow['Current Activity Dsl'] = dslServiceOrigin;
            newRow['DSLMatch'] = (irCategory === dslServiceLower) ? "The same" : "Changed";
        } else {
            newRow['Current Activity Dsl'] = "Not avail";
            newRow['DSLMatch'] = "Not avail";
        }

        // 2. مطابقة الـ FTTH
        const ftthRow = ftthMap.get(irNum);
        if (ftthRow) {
            const ftthServiceOrigin = String(ftthRow['Service Name'] || '').trim();
            const ftthServiceLower = ftthServiceOrigin.toLowerCase();
            newRow['Current Activity FTTH'] = ftthServiceOrigin;
            newRow['FTTH Match'] = (irCategory === ftthServiceLower) ? "The same" : "Changed";
        } else {
            newRow['Current Activity FTTH'] = "Not avail";
            newRow['FTTH Match'] = "Not avail";
        }

        finalCrossMatchData.push(newRow);
    });

    renderTablePreview(finalCrossMatchData);
    document.getElementById('btn-export').style.display = 'inline-block';
});

// دالة عرض جزء من الداتا في الـ HTML عشان المتصفح ميهنجش لو الداتا ضخمة
function renderTablePreview(data) {
    const thead = document.querySelector('#result-table thead');
    const tbody = document.querySelector('#result-table tbody');
    thead.innerHTML = ''; tbody.innerHTML = '';

    if (data.length === 0) return;

    const headers = Object.keys(data[0]);
    let headerHtml = '<tr>';
    headers.forEach(h => headerHtml += `<th>${h}</th>`);
    headerHtml += '</tr>';
    thead.innerHTML = headerHtml;

    const fragment = document.createDocumentFragment();
    // عرض أول 300 صف فقط للحفاظ على سرعة المتصفح
    const previewData = data.slice(0, 300); 

    previewData.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            td.innerText = row[h] !== undefined ? row[h] : '';
            
            // تلوين الخلايا حسب الماتش
            if (h === 'FTTH Match' || h === 'DSLMatch') {
                if (row[h] === 'The same') td.className = 'bg-same';
                else if (row[h] === 'Changed') td.className = 'bg-changed';
                else if (row[h] === 'Not avail') td.className = 'bg-notavail';
            }
            tr.appendChild(td);
        });
        fragment.appendChild(tr);
    });
    tbody.appendChild(fragment);

    const msgDiv = document.getElementById('render-msg');
    msgDiv.style.display = 'block';
    if(data.length > 300) {
        msgDiv.innerText = `Previewing first 300 rows out of ${data.length}. Export to Excel to get the full Data.`;
    } else {
        msgDiv.innerText = `Matched ${data.length} rows successfully.`;
    }
}

// دالة تصدير الملف لإكسيل باستخدام ExcelJS للاحتفاظ بالألوان 
document.getElementById('btn-export').addEventListener('click', async () => {
    if (finalCrossMatchData.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cross-Match Report');
    
    const headers = Object.keys(finalCrossMatchData[0]);
    
    // إنشاء الأعمدة
    worksheet.columns = headers.map(h => ({ header: h, key: h, width: 20 }));
    
    // إضافة الداتا
    worksheet.addRows(finalCrossMatchData);

    // ستايل صف العناوين (الكحلي)
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A192F' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // ستايل الألوان بناءً على الكلمات
    const ftthMatchCol = headers.indexOf('FTTH Match') + 1;
    const dslMatchCol = headers.indexOf('DSLMatch') + 1;

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // تجاوز صف العناوين
        
        [ftthMatchCol, dslMatchCol].forEach(colIndex => {
            if (colIndex > 0) {
                const cell = row.getCell(colIndex);
                if (cell.value === 'The same') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
                    cell.font = { color: { argb: 'FF065F46' }, bold: true };
                } else if (cell.value === 'Changed') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
                    cell.font = { color: { argb: 'FF92400E' }, bold: true };
                } else if (cell.value === 'Not avail') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
                    cell.font = { color: { argb: 'FF991B1B' }, bold: true };
                }
            }
        });
    });

    // استخراج الملف
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Samira_CrossMatch_Output.xlsx');
});