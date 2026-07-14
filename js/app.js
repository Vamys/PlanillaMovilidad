// CONSTANTS
const TEMPLATE_VENDEDOR = './1. Formato Planilla de movilidad - por trabajador.xlsx';
const TEMPLATE_MAESTRO = './2. PASAJES FFVV JULIO 2026 MODELO.xlsx';
const VENDORS_DNI_FILE = './vendors_dni.json';
const BANCOS_VALIDOS = ['BCP', 'BBVA', 'SCOTIABANK', 'INTERBANK', 'PICHINCHA', 'BANBIF'];
const MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

let vendorsDniData = {};

// DOM Elements
document.addEventListener('DOMContentLoaded', async () => {
    // Load vendors DNI
    try {
        const res = await fetch(VENDORS_DNI_FILE);
        vendorsDniData = await res.json();
    } catch (e) {
        console.warn("No se pudo cargar vendors_dni.json", e);
    }

    // --- VENDEDOR TAB ---
    const btnAgregarFila = document.getElementById('btnAgregarFila');
    const recorridosTable = document.getElementById('recorridosTable').getElementsByTagName('tbody')[0];
    const btnGenerarExcel = document.getElementById('btnGenerarExcel');
    
    // Add first row
    agregarFila(recorridosTable);

    btnAgregarFila.addEventListener('click', () => agregarFila(recorridosTable));

    btnGenerarExcel.addEventListener('click', async () => {
        const nombre = document.getElementById('nombreTrabajador').value.trim();
        const banco = document.getElementById('bancoTrabajador').value;
        if (!nombre) {
            Swal.fire('Error', 'Ingrese el nombre del trabajador', 'error');
            return;
        }

        const filas = getFilasData(recorridosTable);
        if (filas.length === 0) {
            Swal.fire('Error', 'Debe agregar al menos un recorrido con monto', 'error');
            return;
        }

        try {
            Swal.fire({ title: 'Generando...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
            
            const arrayBuffer = await fetchFile(TEMPLATE_VENDEDOR);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.getWorksheet(1);

            // DNI lookup
            let dni = "";
            const vData = vendorsDniData[nombre.toUpperCase()];
            if (vData) {
                dni = typeof vData === 'string' ? vData : (vData.dni || "");
            }

            // Fill basic info
            worksheet.getCell('F13').value = nombre.toUpperCase();
            worksheet.getCell('F14').value = dni;
            worksheet.getCell('K11').value = banco;
            worksheet.getCell('G11').value = `Rendir a la cuenta: ${banco}`;
            
            // Llenar mes actual
            const currentMonth = MESES[new Date().getMonth()];
            worksheet.getCell('K14').value = currentMonth;

            // Fill rows (starting at row 18)
            for (let i = 0; i < filas.length; i++) {
                if (i > 13) break; // Max 14 rows in template
                const rowObj = filas[i];
                const d = new Date(`${rowObj.fecha}T12:00:00`); // Evitar problemas de zona horaria
                const r = 18 + i;
                worksheet.getCell(`A${r}`).value = d.getDate();
                worksheet.getCell(`B${r}`).value = d.getMonth() + 1;
                worksheet.getCell(`C${r}`).value = d.getFullYear();
                worksheet.getCell(`F${r}`).value = rowObj.ruta;
                worksheet.getCell(`J${r}`).value = rowObj.lugar;
                worksheet.getCell(`K${r}`).value = parseFloat(rowObj.monto);
            }

            const buffer = await workbook.xlsx.writeBuffer();
            const fileName = `Planilla_${nombre.replace(/\s+/g, '_')}_${currentMonth}.xlsx`;
            saveAs(new Blob([buffer]), fileName);

            Swal.fire('Éxito', 'Archivo Excel generado correctamente', 'success');
        } catch (error) {
            console.error(error);
            Swal.fire('Error', 'Ocurrió un error al generar el archivo: ' + error.message, 'error');
        }
    });

    // --- RRHH TAB ---
    const dropArea = document.getElementById('dropArea');
    const fileInput = document.getElementById('fileArchivosVendedores');
    const fileListDiv = document.getElementById('fileList');
    const btnConsolidar = document.getElementById('btnConsolidar');
    let uploadedFiles = [];

    const handleFiles = (files) => {
        for (let file of files) {
            if (file.name.endsWith('.xlsx')) {
                uploadedFiles.push(file);
            }
        }
        renderFileList();
    };

    const renderFileList = () => {
        if (uploadedFiles.length === 0) {
            fileListDiv.innerHTML = '';
            return;
        }
        let html = '<ul class="list-group">';
        uploadedFiles.forEach((f, i) => {
            html += `<li class="list-group-item d-flex justify-content-between align-items-center">
                        <div><i class="fa-solid fa-file-excel text-success me-2"></i> ${f.name}</div>
                        <button class="btn btn-sm btn-danger" onclick="removeFile(${i})"><i class="fa-solid fa-trash"></i></button>
                     </li>`;
        });
        html += '</ul>';
        fileListDiv.innerHTML = html;
    };

    window.removeFile = (index) => {
        uploadedFiles.splice(index, 1);
        renderFileList();
    };

    dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.classList.add('border-primary', 'bg-light'); });
    dropArea.addEventListener('dragleave', () => { dropArea.classList.remove('border-primary'); });
    dropArea.addEventListener('drop', (e) => {
        e.preventDefault();
        dropArea.classList.remove('border-primary');
        handleFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        fileInput.value = '';
    });

    btnConsolidar.addEventListener('click', async () => {
        if (uploadedFiles.length === 0) {
            Swal.fire('Atención', 'Suba al menos un archivo Excel de vendedor', 'warning');
            return;
        }

        try {
            Swal.fire({ title: 'Consolidando...', html: 'Procesando archivos...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
            const logBox = document.getElementById('logConsole');
            logBox.innerHTML = '';
            document.getElementById('logCard').classList.remove('d-none');
            
            const log = (msg) => { logBox.innerHTML += `${msg}\n`; logBox.scrollTop = logBox.scrollHeight; };
            log('Iniciando consolidación...');

            // Load Master
            const masterBuffer = await fetchFile(TEMPLATE_MAESTRO);
            const masterWb = new ExcelJS.Workbook();
            await masterWb.xlsx.load(masterBuffer);

            // Clean previous records in master
            masterWb.eachSheet((sheet) => {
                const sheetName = sheet.name.toUpperCase();
                if (BANCOS_VALIDOS.includes(sheetName)) {
                    let totalRow = null;
                    sheet.eachRow((row, rowNumber) => {
                        const val = row.getCell(2).value;
                        if (val && String(val).trim().toUpperCase() === 'TOTAL') {
                            totalRow = rowNumber;
                        }
                    });
                    if (totalRow && totalRow > 3) {
                        // Unmerge before deleting (ExcelJS constraint)
                        const merges = Object.values(sheet._merges);
                        merges.forEach(merge => {
                            if (merge.top >= 3 && merge.bottom < totalRow) {
                                sheet.unMergeCells(merge.model.model); // Unmerge
                            }
                        });
                        sheet.spliceRows(3, totalRow - 3);
                    }
                }
            });

            let consolidatedCount = 0;
            const BLOCK_SIZE = 26;

            for (let file of uploadedFiles) {
                const fileBuffer = await file.arrayBuffer();
                const workerWb = new ExcelJS.Workbook();
                await workerWb.xlsx.load(fileBuffer);
                
                let vendorName = null;
                let bancoName = null;
                let records = [];

                workerWb.eachSheet((ws) => {
                    if (vendorName) return; // already found
                    const vName = ws.getCell('F13').value;
                    const vBanco = ws.getCell('K11').value;
                    
                    if (vName && typeof vName === 'string') vendorName = vName.trim();
                    if (vBanco && typeof vBanco === 'string') bancoName = vBanco.trim().toUpperCase();

                    for (let r = 18; r <= 31; r++) {
                        const d = ws.getCell(`A${r}`).value;
                        if (!d) continue;
                        
                        let diaSemana = "OTRO";
                        try {
                            const mo = ws.getCell(`B${r}`).value || 1;
                            const yr = ws.getCell(`C${r}`).value || 2026;
                            const dateObj = new Date(yr, mo - 1, d);
                            const wmap = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
                            diaSemana = wmap[dateObj.getDay()];
                        } catch(e) {}

                        const ruta = ws.getCell(`F${r}`).value;
                        const lugar = ws.getCell(`J${r}`).value;
                        const monto = ws.getCell(`K${r}`).value;

                        records.push({
                            dia_semana: diaSemana,
                            ruta: ruta ? String(ruta).trim() : "",
                            lugar: lugar ? String(lugar).trim() : "",
                            monto: monto ? parseFloat(monto) : 0
                        });
                    }
                });

                if (!vendorName) { log(`❌ Archivo ignorado: '${file.name}' sin nombre en F13.`); continue; }
                if (!bancoName || !BANCOS_VALIDOS.includes(bancoName)) { log(`❌ Archivo ignorado: '${file.name}', banco inválido (${bancoName}).`); continue; }
                if (records.length === 0) { log(`⚠️ Archivo ignorado: '${file.name}', no tiene recorridos.`); continue; }

                // Find or create sheet
                let sheet = masterWb.getWorksheet(bancoName);
                if (!sheet) {
                    sheet = masterWb.addWorksheet(bancoName);
                    sheet.mergeCells('B1:G1');
                    sheet.getCell('B1').value = `TABLA DE PASAJES - ${bancoName}`;
                    const hdrs = ['VENDEDOR', 'DIA', 'RUTA', 'LUGAR', 'MONTO', 'TOTAL'];
                    hdrs.forEach((h, idx) => sheet.getCell(2, idx + 2).value = h);
                    sheet.mergeCells('B3:F3');
                    sheet.getCell('B3').value = 'TOTAL';
                }

                let totalRow = null;
                sheet.eachRow((row, rowNumber) => {
                    const val = row.getCell(2).value;
                    if (val && String(val).trim().toUpperCase() === 'TOTAL') {
                        totalRow = rowNumber;
                    }
                });
                
                const insertRow = totalRow ? totalRow : sheet.rowCount + 1;
                
                // Insert blank rows
                sheet.spliceRows(insertRow, 0, ...Array(BLOCK_SIZE).fill([]));

                const refBorder = {
                    top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
                };
                const refFont = { name: 'Calibri', size: 10 };

                // Vendedor Cell
                const cVend = sheet.getCell(insertRow, 2);
                cVend.value = vendorName;
                cVend.font = { name: 'Calibri', size: 10, bold: true };
                cVend.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cVend.border = refBorder;
                sheet.mergeCells(insertRow, 2, insertRow + BLOCK_SIZE - 1, 2);

                // Total SubCell
                const cTotal = sheet.getCell(insertRow, 7);
                cTotal.value = { formula: `SUM(F${insertRow}:F${insertRow + BLOCK_SIZE - 1})` };
                cTotal.font = refFont;
                cTotal.alignment = { horizontal: 'right', vertical: 'middle' };
                cTotal.numFmt = '#,##0.00';
                cTotal.border = refBorder;
                sheet.mergeCells(insertRow, 7, insertRow + BLOCK_SIZE - 1, 7);

                // Fill data
                for (let idx = 0; idx < BLOCK_SIZE; idx++) {
                    const curr = insertRow + idx;
                    sheet.getCell(curr, 2).border = refBorder;
                    sheet.getCell(curr, 7).border = refBorder;
                    
                    if (idx < records.length) {
                        const rec = records[idx];
                        sheet.getCell(curr, 3).value = rec.dia_semana;
                        sheet.getCell(curr, 4).value = rec.ruta;
                        sheet.getCell(curr, 5).value = rec.lugar;
                        sheet.getCell(curr, 6).value = rec.monto;
                        
                        sheet.getCell(curr, 3).alignment = { horizontal: 'center', vertical: 'middle' };
                        sheet.getCell(curr, 4).alignment = { horizontal: 'left', vertical: 'middle' };
                        sheet.getCell(curr, 5).alignment = { horizontal: 'center', vertical: 'middle' };
                        sheet.getCell(curr, 6).alignment = { horizontal: 'right', vertical: 'middle' };
                        sheet.getCell(curr, 6).numFmt = '#,##0.00';
                    }
                    
                    [3, 4, 5, 6].forEach(col => {
                        const cell = sheet.getCell(curr, col);
                        cell.font = refFont;
                        cell.border = refBorder;
                    });
                }

                // Grand Total
                let gtRow = null;
                sheet.eachRow((r, rNum) => {
                    const val = r.getCell(2).value;
                    if (val && String(val).trim().toUpperCase() === 'TOTAL') {
                        gtRow = rNum;
                    }
                });
                if (!gtRow) {
                    gtRow = sheet.rowCount + 1;
                    sheet.getCell(gtRow, 2).value = "TOTAL";
                    sheet.mergeCells(gtRow, 2, gtRow, 6);
                }
                const cGt = sheet.getCell(gtRow, 7);
                cGt.value = { formula: `SUM(G3:G${gtRow - 1})` };
                cGt.font = { name: 'Calibri', size: 10, bold: true };
                cGt.border = refBorder;
                cGt.alignment = { horizontal: 'right', vertical: 'middle' };
                cGt.numFmt = '#,##0.00';

                log(`✅ Agregado: '${vendorName}' en hoja '${bancoName}' (${records.length} registros).`);
                consolidatedCount++;
            }

            if (consolidatedCount === 0) {
                Swal.fire('Error', 'No se consolidó ningún archivo válido.', 'error');
                return;
            }

            log('Guardando archivo maestro consolidado...');
            const outBuffer = await masterWb.xlsx.writeBuffer();
            const dateStr = new Date().toISOString().slice(0,10).replace(/-/g, '_');
            saveAs(new Blob([outBuffer]), `Consolidado_Pasajes_FFVV_${dateStr}.xlsx`);

            Swal.fire('Éxito', 'Consolidación completada correctamente.', 'success');
        } catch (error) {
            console.error(error);
            Swal.fire('Error', 'Error en la consolidación: ' + error.message, 'error');
        }
    });
});

// --- Funciones Auxiliares ---
async function fetchFile(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Error descargando ${url}`);
    return await res.arrayBuffer();
}

function agregarFila(tbody) {
    const tr = document.createElement('tr');
    
    // Auto-set today's date if empty
    const today = new Date().toISOString().split('T')[0];
    
    tr.innerHTML = `
        <td><input type="date" class="form-control d-fecha" value="${today}" required></td>
        <td><input type="text" class="form-control d-ruta" placeholder="Ruta" required></td>
        <td><input type="text" class="form-control d-lugar" placeholder="Lugar" required></td>
        <td>
            <div class="input-group">
                <span class="input-group-text">S/</span>
                <input type="number" class="form-control d-monto" min="0" step="0.10" value="0.00" onchange="calcularTotal()" onkeyup="calcularTotal()">
            </div>
        </td>
        <td class="text-center">
            <button class="btn btn-sm btn-outline-danger" onclick="eliminarFila(this)"><i class="fa-solid fa-xmark"></i></button>
        </td>
    `;
    tbody.appendChild(tr);
    calcularTotal();
}

function eliminarFila(btn) {
    const tr = btn.closest('tr');
    tr.remove();
    calcularTotal();
}

function calcularTotal() {
    let total = 0;
    const montos = document.querySelectorAll('.d-monto');
    montos.forEach(m => {
        const val = parseFloat(m.value) || 0;
        total += val;
    });
    document.getElementById('totalMonto').innerText = `S/ ${total.toFixed(2)}`;
}

function getFilasData(tbody) {
    const filas = [];
    const trs = tbody.querySelectorAll('tr');
    trs.forEach(tr => {
        const monto = parseFloat(tr.querySelector('.d-monto').value) || 0;
        if (monto > 0) {
            filas.push({
                fecha: tr.querySelector('.d-fecha').value,
                ruta: tr.querySelector('.d-ruta').value,
                lugar: tr.querySelector('.d-lugar').value,
                monto: monto
            });
        }
    });
    return filas;
}
