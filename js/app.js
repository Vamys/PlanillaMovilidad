// CONSTANTS
const TEMPLATE_VENDEDOR = './1. Formato Planilla de movilidad - por trabajador.xlsx';
const TEMPLATE_MAESTRO = './2. PASAJES FFVV JULIO 2026 MODELO.xlsx';
const VENDORS_DNI_FILE = './vendors_dni.json';
const BANCOS_VALIDOS = ['BCP', 'BBVA', 'SCOTIABANK', 'INTERBANK', 'PICHINCHA', 'BANBIF'];
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
            Swal.fire({icon: 'error', title: 'Atención', text: 'Debe agregar al menos un recorrido con monto', background: '#1e293b', color: '#fff'});
            return;
        }

        try {
            Swal.fire({ title: 'Generando...', allowOutsideClick: false, background: '#1e293b', color: '#fff', didOpen: () => Swal.showLoading() });
            
            const arrayBuffer = await fetchFile(TEMPLATE_VENDEDOR);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.getWorksheet(1);

            // Fill basic info
            worksheet.getCell('F13').value = nombre.toUpperCase();
            worksheet.getCell('F14').value = dni;
            worksheet.getCell('K11').value = banco;
            worksheet.getCell('G11').value = `Rendir a la cuenta: ${banco}`;
            
            // Periodo
            let mesNombre = MESES[new Date().getMonth()];
            if (periodoStr) {
                const parts = periodoStr.split('-');
                mesNombre = MESES[parseInt(parts[1]) - 1];
            }
            worksheet.getCell('K14').value = mesNombre;

            // Fill rows (starting at row 18)
            for (let i = 0; i < filas.length; i++) {
                const rowObj = filas[i];
                let d, m, y;
                if (rowObj.fecha) {
                    const dateObj = new Date(`${rowObj.fecha}T12:00:00`);
                    d = dateObj.getDate();
                    m = dateObj.getMonth() + 1;
                    y = dateObj.getFullYear();
                } else {
                    d = rowObj.dia_num;
                    m = "";
                    y = "";
                }
                
                const r = 18 + i;
                worksheet.getCell(`A${r}`).value = d;
                worksheet.getCell(`B${r}`).value = m;
                worksheet.getCell(`C${r}`).value = y;
                worksheet.getCell(`F${r}`).value = rowObj.ruta;
                worksheet.getCell(`J${r}`).value = rowObj.lugar;
                worksheet.getCell(`K${r}`).value = parseFloat(rowObj.monto);
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

            // Clean master
            masterWb.eachSheet((sheet) => {
                const sheetName = sheet.name.toUpperCase();
                if (BANCOS_VALIDOS.includes(sheetName)) {
                    let totalRow = null;
                    sheet.eachRow((row, rowNumber) => {
                        const val = row.getCell(2).value;
                        if (val && String(val).trim().toUpperCase() === 'TOTAL') totalRow = rowNumber;
                    });
                    if (totalRow && totalRow > 3) {
                        if (sheet._merges) {
                            for (let range in sheet._merges) {
                                const mergeObj = sheet._merges[range];
                                if (!mergeObj || !mergeObj.model) continue;
                                if (mergeObj.model.top >= 3 && mergeObj.model.bottom < totalRow) {
                                    sheet.unMergeCells(range);
                                }
                            }
                        }
                        sheet.spliceRows(3, totalRow - 3);
                    }
                }
            });

            let consolidatedCount = 0;
            const BLOCK_SIZE = 26;

            for (let file of uploadedFiles) {
                try {
                    const fileBuffer = await file.arrayBuffer();
                    const workerWb = new ExcelJS.Workbook();
                    // Intentar cargar el archivo; algunos xlsx generan errores de ExcelJS
                    await workerWb.xlsx.load(fileBuffer);
                
                let vendorName = null;
                let bancoName = null;
                let records = [];

                workerWb.eachSheet((ws) => {
                    if (vendorName) return; 
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

                        records.push({
                            dia_semana: diaSemana,
                            ruta: ws.getCell(`F${r}`).value || "",
                            lugar: ws.getCell(`J${r}`).value || "",
                            monto: ws.getCell(`K${r}`).value || 0
                        });
                    }
                });

                if (!vendorName) { log(`❌ Ignorado: '${file.name}' sin nombre.`); continue; }
                if (!bancoName || !BANCOS_VALIDOS.includes(bancoName)) { log(`❌ Ignorado: '${file.name}', banco inválido.`); continue; }
                if (records.length === 0) { log(`⚠️ Ignorado: '${file.name}', sin rutas.`); continue; }

                let sheet = masterWb.getWorksheet(bancoName);
                if (!sheet) {
                    sheet = masterWb.addWorksheet(bancoName);
                    sheet.mergeCells('B1:G1');
                    sheet.getCell('B1').value = `TABLA DE PASAJES - ${bancoName}`;
                    ['VENDEDOR', 'DIA', 'RUTA', 'LUGAR', 'MONTO', 'TOTAL'].forEach((h, idx) => sheet.getCell(2, idx + 2).value = h);
                    sheet.mergeCells('B3:F3');
                    sheet.getCell('B3').value = 'TOTAL';
                }

                let totalRow = null;
                sheet.eachRow((row, rowNumber) => {
                    if (String(row.getCell(2).value || '').trim().toUpperCase() === 'TOTAL') totalRow = rowNumber;
                });
                const insertRow = totalRow ? totalRow : sheet.rowCount + 1;
                
                sheet.spliceRows(insertRow, 0, ...Array(BLOCK_SIZE).fill([]));

                const refBorder = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                const refFont = { name: 'Calibri', size: 10 };

                const cVend = sheet.getCell(insertRow, 2);
                cVend.value = vendorName;
                cVend.font = { name: 'Calibri', size: 10, bold: true };
                cVend.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cVend.border = refBorder;
                sheet.mergeCells(insertRow, 2, insertRow + BLOCK_SIZE - 1, 2);

                const cTotal = sheet.getCell(insertRow, 7);
                cTotal.value = { formula: `SUM(F${insertRow}:F${insertRow + BLOCK_SIZE - 1})` };
                cTotal.font = refFont;
                cTotal.alignment = { horizontal: 'right', vertical: 'middle' };
                cTotal.numFmt = '#,##0.00';
                cTotal.border = refBorder;
                sheet.mergeCells(insertRow, 7, insertRow + BLOCK_SIZE - 1, 7);

                for (let idx = 0; idx < BLOCK_SIZE; idx++) {
                    const curr = insertRow + idx;
                    sheet.getCell(curr, 2).border = refBorder;
                    sheet.getCell(curr, 7).border = refBorder;
                    
                    if (idx < records.length) {
                        const rec = records[idx];
                        sheet.getCell(curr, 3).value = rec.dia_semana;
                        sheet.getCell(curr, 4).value = rec.ruta;
                        sheet.getCell(curr, 5).value = rec.lugar;
                        sheet.getCell(curr, 6).value = parseFloat(rec.monto);
                        
                        sheet.getCell(curr, 3).alignment = { horizontal: 'center', vertical: 'middle' };
                        sheet.getCell(curr, 4).alignment = { horizontal: 'left', vertical: 'middle' };
                        sheet.getCell(curr, 5).alignment = { horizontal: 'center', vertical: 'middle' };
                        sheet.getCell(curr, 6).alignment = { horizontal: 'right', vertical: 'middle' };
                        sheet.getCell(curr, 6).numFmt = '#,##0.00';
                    }
                    [3, 4, 5, 6].forEach(col => { sheet.getCell(curr, col).font = refFont; sheet.getCell(curr, col).border = refBorder; });
                }

                let gtRow = null;
                sheet.eachRow((r, rNum) => {
                    if (String(r.getCell(2).value || '').trim().toUpperCase() === 'TOTAL') gtRow = rNum;
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

                log(`✅ Agregado: '${vendorName}' (${records.length} rutas)`);
                consolidatedCount++;
                } catch (fileErr) {
                    console.error(`Error procesando ${file.name}:`, fileErr);
                    log(`❌ Error leyendo '${file.name}': ${fileErr.message}. Prueba guardarlo como .xlsx desde Excel y reintenta.`);
                }
            }

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
    const res = await fetch(url);
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
        if (monto > 0) {
            filas.push({
                dia_num: tr.querySelector('.d-dia').value,
                fecha: tr.querySelector('.d-fecha').value,
                ruta: tr.querySelector('.d-ruta').value,
                lugar: tr.querySelector('.d-lugar').value,
                monto: monto
            });
        }
    });
    return filas;
}
