// CONSTANTS
const TEMPLATE_VENDEDOR = './1. Formato Planilla de movilidad - por trabajador.xlsx';
const TEMPLATE_MAESTRO = './2. PASAJES FFVV JULIO 2026 MODELO.xlsx';
const VENDORS_DNI_FILE = './vendors_dni.json';
const BANCOS_VALIDOS = ['BCP', 'BBVA', 'SCOTIABANK', 'INTERBANK', 'PICHINCHA', 'BANBIF', 'FINANCIERO'];
const MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

let vendorsDniData = {};

document.addEventListener('DOMContentLoaded', async () => {
    
    // --- LOAD VENDORS ---
    const selectTrabajador = document.getElementById('nombreTrabajador');
    const inputNuevoTrabajador = document.getElementById('inputNuevoTrabajador');
    const inputDni = document.getElementById('dniTrabajador');
    const selectBanco = document.getElementById('bancoTrabajador');
    
    try {
        const res = await fetch(VENDORS_DNI_FILE);
        vendorsDniData = await res.json();
        
        Object.keys(vendorsDniData).forEach(vendor => {
            const opt = document.createElement('option');
            opt.value = vendor;
            opt.textContent = vendor;
            selectTrabajador.appendChild(opt);
        });
    } catch (e) {
        console.warn("No se pudo cargar vendors_dni.json", e);
    }

    selectTrabajador.addEventListener('change', async (e) => {
        const vendor = e.target.value;
        
        if (vendor === "NUEVO") {
            inputNuevoTrabajador.classList.remove('d-none');
            inputNuevoTrabajador.focus();
            inputDni.value = "";
            selectBanco.value = "";
            return;
        }

        inputNuevoTrabajador.classList.add('d-none');

        if (vendor && vendorsDniData[vendor]) {
            const data = vendorsDniData[vendor];
            inputDni.value = typeof data === 'string' ? data : (data.dni || "");
            if (typeof data === 'object' && data.banco) {
                selectBanco.value = data.banco;
            } else {
                selectBanco.value = "";
            }
        } else {
            inputDni.value = "";
            selectBanco.value = "";
        }
    });

    // Default dates
    const today = new Date();
    document.getElementById('fechaEmision').value = today.toISOString().split('T')[0];
    // Periodo vacío por defecto para forzar selección
    const inputPeriodo = document.getElementById('periodoTrabajador');
    inputPeriodo.value = "";

    // --- VENDEDOR TAB ---
    const btnAgregarFila = document.getElementById('btnAgregarFila');
    const btnRestablecer = document.getElementById('btnRestablecer');
    const recorridosTable = document.getElementById('recorridosTable').getElementsByTagName('tbody')[0];
    const btnGenerarExcel = document.getElementById('btnGenerarExcel');
    
    // Auto-generar días cuando cambia el periodo
    inputPeriodo.addEventListener('change', (e) => {
        const pVal = e.target.value;
        if (!pVal) return;
        
        const emptyMsg = document.getElementById('emptyRowMessage');
        if (emptyMsg) emptyMsg.remove();
        
        // Limpiar tabla actual
        recorridosTable.innerHTML = '';
        
        const [yyyy, mm] = pVal.split('-');
        const year = parseInt(yyyy);
        const month = parseInt(mm) - 1; // 0-indexed
        
        // Días del mes
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        
        for (let i = 1; i <= daysInMonth; i++) {
            const currentDate = new Date(year, month, i);
            const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
            
            // Omitir domingos por defecto? O los mostramos todos. Vamos a mostrarlos todos.
            agregarFilaEspecifica(recorridosTable, i, dateStr);
        }
    });
    
    btnAgregarFila.addEventListener('click', () => {
        const emptyMsg = document.getElementById('emptyRowMessage');
        if (emptyMsg) emptyMsg.remove();
        agregarFila(recorridosTable);
    });
    
    btnRestablecer.addEventListener('click', () => {
        recorridosTable.innerHTML = '';
        agregarFila(recorridosTable);
    });

    btnGenerarExcel.addEventListener('click', async () => {
        let nombre = selectTrabajador.value;
        if (nombre === "NUEVO") {
            nombre = inputNuevoTrabajador.value.trim();
        }
        
        const banco = selectBanco.value;
        const dni = inputDni.value;
        const periodoStr = document.getElementById('periodoTrabajador').value;
        
        if (!nombre) {
            Swal.fire({icon: 'error', title: 'Atención', text: 'Seleccione un trabajador', background: '#1e293b', color: '#fff'});
            return;
        }

        const filas = getFilasData(recorridosTable);
        if (filas.length === 0) {
            Swal.fire({icon: 'error', title: 'Atención', text: 'No hay filas en la tabla', background: '#1e293b', color: '#fff'});
            return;
        }

        const hasMonto = filas.some(f => f.monto > 0);
        if (!hasMonto) {
            Swal.fire({icon: 'error', title: 'Atención', text: 'Debe agregar al menos un recorrido con monto mayor a 0', background: '#1e293b', color: '#fff'});
            return;
        }

        try {
            Swal.fire({ title: 'Generando...', allowOutsideClick: false, background: '#1e293b', color: '#fff', didOpen: () => Swal.showLoading() });
            
            const arrayBuffer = await fetchFile(TEMPLATE_VENDEDOR);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const templateSheet = workbook.getWorksheet(1);

            let mesNombre = MESES[new Date().getMonth()];
            if (periodoStr) {
                const parts = periodoStr.split('-');
                mesNombre = MESES[parseInt(parts[1]) - 1];
            }

            const chunk_size = 14;
            const row_chunks = [];
            for (let i = 0; i < filas.length; i += chunk_size) {
                row_chunks.push(filas.slice(i, i + chunk_size));
            }

            for (let p_idx = 0; p_idx < row_chunks.length; p_idx++) {
                const chunk = row_chunks[p_idx];
                let sheet;
                if (p_idx === 0) {
                    sheet = templateSheet;
                    sheet.name = `Planilla Pg ${p_idx + 1}`;
                } else {
                    sheet = workbook.addWorksheet(`Planilla Pg ${p_idx + 1}`);
                    copyWorksheet(templateSheet, sheet);
                }

                // Fill basic info
                sheet.getCell('F13').value = nombre.toUpperCase();
                sheet.getCell('F14').value = dni;
                sheet.getCell('K11').value = banco;
                sheet.getCell('G11').value = `Rendir a la cuenta: ${banco}`;
                sheet.getCell('K14').value = mesNombre;

                // N° Planilla
                const nroPlanillaNum = parseInt(document.getElementById('nroPlanilla').value) || 1;
                sheet.getCell('K1').value = String(nroPlanillaNum + p_idx).padStart(2, '0');

                // Periodo en D8
                if (periodoStr) {
                    const [yr, mo] = periodoStr.split('-');
                    const pDate = new Date(yr, mo - 1, 1);
                    const cellD8 = sheet.getCell('D8');
                    cellD8.value = pDate;
                    cellD8.numFmt = 'mmm-yy';
                }

                // Fecha Emisión en K5
                const fechaEmisionStr = document.getElementById('fechaEmision').value;
                if (fechaEmisionStr) {
                    const [yr, mo, dy] = fechaEmisionStr.split('-');
                    const feDate = new Date(yr, mo - 1, dy);
                    const cellK5 = sheet.getCell('K5');
                    cellK5.value = feDate;
                    cellK5.numFmt = 'dd/mm/yyyy';
                }

                // Fill rows (starting at row 18)
                for (let r_idx = 0; r_idx < chunk_size; r_idx++) {
                    const r = 18 + r_idx;
                    if (r_idx < chunk.length) {
                        const rowObj = chunk[r_idx];
                        let d = "", m = "", y = "";
                        if (rowObj.fecha) {
                            const dateObj = new Date(`${rowObj.fecha}T12:00:00`);
                            d = dateObj.getDate();
                            m = dateObj.getMonth() + 1;
                            y = dateObj.getFullYear();
                        } else {
                            d = parseInt(rowObj.dia_num) || "";
                        }
                        
                        sheet.getCell(`A${r}`).value = d;
                        sheet.getCell(`B${r}`).value = m;
                        sheet.getCell(`C${r}`).value = y;
                        sheet.getCell(`D${r}`).value = rowObj.motivo;
                        sheet.getCell(`F${r}`).value = rowObj.ruta;
                        sheet.getCell(`J${r}`).value = rowObj.lugar;
                        sheet.getCell(`K${r}`).value = rowObj.monto > 0 ? parseFloat(rowObj.monto) : null;
                    } else {
                        // Clear template values
                        sheet.getCell(`A${r}`).value = null;
                        sheet.getCell(`B${r}`).value = null;
                        sheet.getCell(`C${r}`).value = null;
                        sheet.getCell(`D${r}`).value = null;
                        sheet.getCell(`F${r}`).value = null;
                        sheet.getCell(`J${r}`).value = null;
                        sheet.getCell(`K${r}`).value = null;
                    }
                }
                sheet.getCell('K32').value = { formula: `SUM(K18:K31)` };
            }

            const buffer = await workbook.xlsx.writeBuffer();
            const fileName = `Planilla_${nombre.replace(/\s+/g, '_')}_${mesNombre}.xlsx`;
            saveAs(new Blob([buffer]), fileName);

            Swal.fire({icon: 'success', title: 'Éxito', text: 'Archivo Excel generado correctamente', background: '#1e293b', color: '#fff'});
        } catch (error) {
            console.error(error);
            Swal.fire({icon: 'error', title: 'Error', text: 'Ocurrió un error al generar el archivo: ' + error.message, background: '#1e293b', color: '#fff'});
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
        let html = '<ul class="list-group list-group-flush border border-secondary rounded overflow-hidden">';
        uploadedFiles.forEach((f, i) => {
            html += `<li class="list-group-item bg-dark text-white d-flex justify-content-between align-items-center border-secondary">
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

    dropArea.addEventListener('dragover', (e) => { e.preventDefault(); dropArea.style.borderColor = 'var(--primary-accent)'; });
    dropArea.addEventListener('dragleave', () => { dropArea.style.borderColor = 'var(--border-color)'; });
    dropArea.addEventListener('drop', (e) => {
        e.preventDefault();
        dropArea.style.borderColor = 'var(--border-color)';
        handleFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        fileInput.value = '';
    });

    btnConsolidar.addEventListener('click', async () => {
        if (uploadedFiles.length === 0) {
            Swal.fire({icon: 'warning', title: 'Atención', text: 'Suba al menos un archivo Excel', background: '#1e293b', color: '#fff'});
            return;
        }

        try {
            Swal.fire({ title: 'Consolidando...', allowOutsideClick: false, background: '#1e293b', color: '#fff', didOpen: () => Swal.showLoading() });
            const logBox = document.getElementById('logConsole');
            logBox.innerHTML = '';
            document.getElementById('logCard').classList.remove('d-none');
            
            const log = (msg) => { logBox.innerHTML += `${msg}\n`; logBox.scrollTop = logBox.scrollHeight; };
            log('Iniciando consolidación...');

            const masterBuffer = await fetchFile(TEMPLATE_MAESTRO);
            const masterWb = new ExcelJS.Workbook();
            await masterWb.xlsx.load(masterBuffer);

            // Clean master: unmerge everything below row 2 and clear data rows from row 3
            masterWb.eachSheet((sheet) => {
                const sheetName = sheet.name.toUpperCase();
                if (!BANCOS_VALIDOS.includes(sheetName)) return;

                // Collect all merges from the sheet safely
                const mergeRanges = [];
                if (sheet._merges) {
                    // ExcelJS stores merges in _merges keyed by top-left address
                    // The values are Range objects with .tl, .br, .model etc.
                    Object.keys(sheet._merges).forEach(key => {
                        const m = sheet._merges[key];
                        // Only top-left cell has a full Range, skip secondary cells
                        if (m && m.model && m.model.top && m.model.left &&
                            m.model.top === m.model.top && m.model.left === m.model.left) {
                            mergeRanges.push({
                                key: key,
                                top: m.model.top,
                                left: m.model.left,
                                bottom: m.model.bottom,
                                right: m.model.right
                            });
                        }
                    });
                }

                // Unmerge those at row 3 or below
                mergeRanges.filter(m => m.top >= 3).forEach(m => {
                    try { sheet.unMergeCells(m.top, m.left, m.bottom, m.right); } catch(e) {}
                });

                // Clear all data rows from row 3 down (leave header rows 1+2)
                const maxR = sheet.rowCount;
                for (let r = maxR; r >= 3; r--) {
                    const row = sheet.getRow(r);
                    let hasContent = false;
                    row.eachCell(cell => { if (cell.value !== null && cell.value !== undefined) hasContent = true; });
                    if (hasContent) {
                        row.eachCell(cell => { cell.value = null; });
                    }
                }
            });

            let consolidatedCount = 0;
            const BLOCK_SIZE = 26;
            // Track next write row per sheet
            const sheetNextRow = {};
            // Initialize with row 3 for all existing bank sheets
            masterWb.eachSheet(sheet => {
                if (BANCOS_VALIDOS.includes(sheet.name.toUpperCase())) {
                    sheetNextRow[sheet.name.toUpperCase()] = 3;
                }
            });

            for (let file of uploadedFiles) {
                try {
                    const fileBuffer = await file.arrayBuffer();
                    const workerWb = XLSX.read(fileBuffer, { type: 'array' });
                    
                    let vendorName = null;
                    let bancoName = null;
                    let records = [];

                    for (let sheetName of workerWb.SheetNames) {
                        const ws = workerWb.Sheets[sheetName];
                        if (!ws) continue;

                        const vName = ws['F13'] ? ws['F13'].v : null;
                        const vBanco = ws['K11'] ? ws['K11'].v : null;
                        if (vName && typeof vName === 'string' && !vendorName) vendorName = vName.trim();
                        if (vBanco && typeof vBanco === 'string' && !bancoName) bancoName = vBanco.trim().toUpperCase();

                        for (let r = 18; r <= 31; r++) {
                            const cellA = ws['A' + r];
                            const d = cellA ? cellA.v : null;
                            if (d === null || d === undefined) continue;

                            let diaSemana = "OTRO";
                            try {
                                const mo = ws['B' + r] ? ws['B' + r].v : 1;
                                const yr = ws['C' + r] ? ws['C' + r].v : 2026;
                                const wmap = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
                                diaSemana = wmap[new Date(yr, mo - 1, d).getDay()];
                            } catch(e) {}

                            const cellF = ws['F' + r];
                            const cellJ = ws['J' + r];
                            const cellK = ws['K' + r];
                            const monto = cellK ? parseFloat(cellK.v) || 0 : 0;
                            if (monto > 0) {
                                records.push({
                                    dia_semana: diaSemana,
                                    ruta: cellF ? String(cellF.v).trim() : "",
                                    lugar: cellJ ? String(cellJ.v).trim() : "",
                                    monto: monto
                                });
                            }
                        }
                    }

                    if (!vendorName) { log(`❌ Ignorado: '${file.name}' sin nombre.`); continue; }
                    if (!bancoName || !BANCOS_VALIDOS.includes(bancoName)) { log(`❌ Ignorado: '${file.name}', banco inválido.`); continue; }
                    if (records.length === 0) { log(`⚠️ Ignorado: '${file.name}', sin montos.`); continue; }

                    let sheet = masterWb.worksheets.find(s => s.name.toUpperCase().trim() === bancoName);
                    if (!sheet) {
                        sheet = masterWb.addWorksheet(bancoName);
                        ['VENDEDOR', 'DIA', 'RUTA', 'LUGAR', 'MONTO', 'TOTAL'].forEach((h, i) => sheet.getCell(2, i + 2).value = h);
                        sheetNextRow[bancoName] = 3;
                    }

                    if (!sheetNextRow[bancoName]) sheetNextRow[bancoName] = 3;
                    const insertRow = sheetNextRow[bancoName];

                    const refBorder = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                    const refFont = { name: 'Calibri', size: 10 };

                    // Write vendor name cell (col B) — will merge at end
                    const cVend = sheet.getCell(insertRow, 2);
                    cVend.value = vendorName;
                    cVend.font = { name: 'Calibri', size: 10, bold: true };
                    cVend.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    cVend.border = refBorder;

                    // Write vendor total cell (col G) — will merge at end
                    const cVendTotal = sheet.getCell(insertRow, 7);
                    cVendTotal.font = refFont;
                    cVendTotal.alignment = { horizontal: 'right', vertical: 'middle' };
                    cVendTotal.numFmt = '#,##0.00';
                    cVendTotal.border = refBorder;

                    // Write data rows
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
                            sheet.getCell(curr, col).font = refFont;
                            sheet.getCell(curr, col).border = refBorder;
                        });
                    }

                    // Now merge vendor name and vendor total columns for the block
                    const blockEnd = insertRow + BLOCK_SIZE - 1;
                    try { sheet.mergeCells(insertRow, 2, blockEnd, 2); } catch(e) { console.warn('merge B', e.message); }
                    cVendTotal.value = { formula: `SUM(F${insertRow}:F${blockEnd})` };
                    try { sheet.mergeCells(insertRow, 7, blockEnd, 7); } catch(e) { console.warn('merge G', e.message); }

                    sheetNextRow[bancoName] = blockEnd + 1;

                    log(`✅ Agregado: '${vendorName}' (${records.length} rutas)`);
                    consolidatedCount++;
                } catch (fileErr) {
                    console.error(`Error procesando ${file.name}:`, fileErr);
                    log(`❌ Error leyendo '${file.name}': ${fileErr.message}`);
                }
            }

            // Write TOTAL row at the bottom of each sheet that was written to
            const refBorder = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            masterWb.eachSheet(sheet => {
                const sheetNameUpper = sheet.name.toUpperCase();
                if (!BANCOS_VALIDOS.includes(sheetNameUpper)) return;
                if (!sheetNextRow[sheetNameUpper] || sheetNextRow[sheetNameUpper] === 3) return;

                const gtRow = sheetNextRow[sheetNameUpper];
                const cGtLabel = sheet.getCell(gtRow, 2);
                cGtLabel.value = 'TOTAL';
                cGtLabel.font = { name: 'Calibri', size: 10, bold: true };
                cGtLabel.alignment = { horizontal: 'center', vertical: 'middle' };
                cGtLabel.border = refBorder;
                try { sheet.mergeCells(gtRow, 2, gtRow, 6); } catch(e) {}

                const cGt = sheet.getCell(gtRow, 7);
                cGt.value = { formula: `SUM(G3:G${gtRow - 1})` };
                cGt.font = { name: 'Calibri', size: 10, bold: true };
                cGt.border = refBorder;
                cGt.alignment = { horizontal: 'right', vertical: 'middle' };
                cGt.numFmt = '#,##0.00';
            });

            if (consolidatedCount === 0) {
                Swal.fire({icon: 'error', title: 'Error', text: 'No se consolidó ningún archivo válido.', background: '#1e293b', color: '#fff'});
                return;
            }

            log('Guardando archivo...');
            const outBuffer = await masterWb.xlsx.writeBuffer();
            const dateStr = new Date().toISOString().slice(0,10).replace(/-/g, '_');
            saveAs(new Blob([outBuffer]), `Consolidado_Pasajes_FFVV_${dateStr}.xlsx`);
            Swal.fire({icon: 'success', title: 'Éxito', text: 'Consolidación completada correctamente.', background: '#1e293b', color: '#fff'});
        } catch (error) {
            Swal.fire({icon: 'error', title: 'Error', text: error.message, background: '#1e293b', color: '#fff'});
        }
    });
});

async function fetchFile(url) {
    // Agregar cache-buster para evitar que el navegador cachee plantillas viejas
    const res = await fetch(url + '?t=' + new Date().getTime());
    if (!res.ok) throw new Error(`Error descargando ${url}`);
    return await res.arrayBuffer();
}

function agregarFila(tbody) {
    const todayStr = new Date().toISOString().split('T')[0];
    const d = new Date().getDate();
    agregarFilaEspecifica(tbody, d, todayStr);
}

function agregarFilaEspecifica(tbody, diaNum, fechaStr) {
    const tr = document.createElement('tr');
    
    // Columns: DÍA (num), FECHA / DÍA SEMANA, MOTIVO, DESTINO/RUTA, VIAJE(LUGAR), MONTO, ACCIÓN
    tr.innerHTML = `
        <td style="width: 70px;"><input type="number" class="form-control dark-input d-dia" min="1" max="31" value="${diaNum}"></td>
        <td><input type="date" class="form-control dark-input d-fecha" value="${fechaStr}"></td>
        <td><input type="text" class="form-control dark-input d-motivo" placeholder="Visita / G. Admin"></td>
        <td><input type="text" class="form-control dark-input d-ruta" placeholder="Destino / Ruta" required></td>
        <td><input type="text" class="form-control dark-input d-lugar" placeholder="Viaje (Lugar)"></td>
        <td><input type="number" class="form-control dark-input d-monto text-end" min="0" step="0.10" value="0.00" onchange="calcularTotal()" onkeyup="calcularTotal()"></td>
        <td class="text-center">
            <button class="btn btn-sm btn-outline-danger" onclick="eliminarFila(this)"><i class="fa-solid fa-xmark"></i></button>
        </td>
    `;
    tbody.appendChild(tr);
    
    // When date changes, update "dia" automatically if possible
    tr.querySelector('.d-fecha').addEventListener('change', (e) => {
        if(e.target.value) {
            const d = new Date(`${e.target.value}T12:00:00`);
            tr.querySelector('.d-dia').value = d.getDate();
        }
    });

    calcularTotal();
}

function eliminarFila(btn) {
    const tr = btn.closest('tr');
    tr.remove();
    calcularTotal();
}

window.calcularTotal = function() {
    let total = 0;
    const montos = document.querySelectorAll('.d-monto');
    montos.forEach(m => {
        total += parseFloat(m.value) || 0;
    });
    document.getElementById('totalMonto').innerText = `S/. ${total.toFixed(2)}`;
}

function getFilasData(tbody) {
    const filas = [];
    const trs = tbody.querySelectorAll('tr');
    trs.forEach(tr => {
        if (!tr.querySelector('.d-monto')) return; // Ignore empty message row
        const monto = parseFloat(tr.querySelector('.d-monto').value) || 0;
        filas.push({
            dia_num: tr.querySelector('.d-dia').value,
            fecha: tr.querySelector('.d-fecha').value,
            motivo: tr.querySelector('.d-motivo').value || "",
            ruta: tr.querySelector('.d-ruta').value || "",
            lugar: tr.querySelector('.d-lugar').value || "",
            monto: monto
        });
    });
    return filas;
}

function safeInsertRows(sheet, insertRow, numRows) {
    const originalMerges = [];
    if (sheet._merges) {
        const mergesObj = sheet._merges.map || {};
        for (let range in mergesObj) {
            const mergeObj = mergesObj[range];
            if (mergeObj && mergeObj.model) {
                originalMerges.push({
                    range: range,
                    model: Object.assign({}, mergeObj.model)
                });
            }
        }
    }
    
    // Unmerge all merges using their exact range keys
    originalMerges.forEach(m => {
        try { sheet.unMergeCells(m.range); } catch(e) {}
    });
    
    // Actually insert the rows (ExcelJS shifts cell values and single cell styles)
    sheet.spliceRows(insertRow, 0, ...Array(numRows).fill([]));
    
    // Re-merge all merges, shifting those that are at or below insertRow
    originalMerges.forEach(m => {
        let { top, left, bottom, right } = m.model;
        if (top >= insertRow) {
            top += numRows;
            bottom += numRows;
        }
        try {
            sheet.mergeCells(top, left, bottom, right);
        } catch(e) {
            console.warn(`Failed to re-merge shifted range ${top},${left}:${bottom},${right}`, e);
        }
    });
}

function copyWorksheet(srcSheet, destSheet) {
    // Copy column widths
    if (srcSheet.columns) {
        destSheet.columns = srcSheet.columns.map(col => {
            return {
                key: col.key,
                header: col.header,
                width: col.width,
                style: col.style,
                hidden: col.hidden,
                outlineLevel: col.outlineLevel
            };
        });
    }
    
    // Copy rows and cells
    srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const destRow = destSheet.getRow(rowNumber);
        destRow.height = row.height;
        
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const destCell = destRow.getCell(colNumber);
            destCell.value = cell.value;
            // Copy styles
            if (cell.font) destCell.font = Object.assign({}, cell.font);
            if (cell.fill) destCell.fill = Object.assign({}, cell.fill);
            if (cell.border) destCell.border = Object.assign({}, cell.border);
            if (cell.alignment) destCell.alignment = Object.assign({}, cell.alignment);
            if (cell.numFmt) destCell.numFmt = cell.numFmt;
        });
    });
    
    // Copy merges
    const mergesObj = srcSheet._merges && srcSheet._merges.map ? srcSheet._merges.map : {};
    for (let range in mergesObj) {
        const mergeObj = mergesObj[range];
        if (mergeObj && mergeObj.model) {
            try {
                destSheet.mergeCells(
                    mergeObj.model.top,
                    mergeObj.model.left,
                    mergeObj.model.bottom,
                    mergeObj.model.right
                );
            } catch(e) {}
        }
    }
}
