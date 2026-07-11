// Global state
let vendorsData = {};
let dnisData = {};
let bancosData = {};
let queueFiles = [];

document.addEventListener('DOMContentLoaded', () => {
    initCommon();
    if (document.getElementById('planillaForm')) {
        initVendedorForm();
    }
    if (document.getElementById('dropZone')) {
        initRrhhPanel();
    }
});

// Common utilities
function initCommon() {
    // Modal Close
    const modal = document.getElementById('alertModal');
    const closeBtn = document.querySelector('.close-btn');
    if (closeBtn && modal) {
        closeBtn.onclick = () => modal.style.display = 'none';
        window.onclick = (event) => {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        };
    }
}

function showAlert(message) {
    const modal = document.getElementById('alertModal');
    const msgEl = document.getElementById('modalMessage');
    if (modal && msgEl) {
        msgEl.textContent = message;
        modal.style.display = 'flex';
    } else {
        alert(message);
    }
}

// VENDEDOR FORM LOGIC
function initVendedorForm() {
    const form = document.getElementById('planillaForm');
    const vendedorSelect = document.getElementById('vendedorSelect');
    const nuevoVendedorInput = document.getElementById('nuevoVendedorInput');
    const dniInput = document.getElementById('dniInput');
    const periodoInput = document.getElementById('periodoInput');
    const fechaEmisionInput = document.getElementById('fechaEmisionInput');
    const btnAgregar = document.getElementById('btnAgregarFila');
    const btnRestablecer = document.getElementById('btnRestablecer');
    const calendarBody = document.getElementById('calendarBody');

    // Set default fecha emision as today
    const today = new Date().toISOString().split('T')[0];
    fechaEmisionInput.value = today;

    // Load Vendors list from API
    fetch('/api/get-vendors')
        .then(res => res.json())
        .then(data => {
            vendorsData = data.vendors || {};
            dnisData = data.dnis || {};
            bancosData = data.bancos || {};
            
            // Populate select dropdown
            Object.keys(vendorsData).sort().forEach(vendor => {
                const opt = document.createElement('option');
                opt.value = vendor;
                opt.textContent = vendor;
                vendedorSelect.appendChild(opt);
            });
        })
        .catch(err => {
            console.error('Error fetching vendors:', err);
            showAlert('Error al conectar con el servidor para obtener los vendedores.');
        });

    // Handle Vendor Selection
    vendedorSelect.addEventListener('change', () => {
        const val = vendedorSelect.value;
        const bancoSelect = document.getElementById('bancoSelect');
        if (val === 'NUEVO') {
            nuevoVendedorInput.style.display = 'block';
            nuevoVendedorInput.setAttribute('required', 'true');
            dniInput.value = '';
            bancoSelect.value = '';
        } else {
            nuevoVendedorInput.style.display = 'none';
            nuevoVendedorInput.removeAttribute('required');
            if (dnisData[val]) {
                dniInput.value = dnisData[val];
            } else {
                dniInput.value = '';
            }
            if (bancosData[val]) {
                bancoSelect.value = bancosData[val];
            } else {
                bancoSelect.value = '';
            }
        }
        generateCalendar();
    });

    // Handle Period Change
    periodoInput.addEventListener('change', generateCalendar);

    // Dynamic Row Totals Calculation
    calendarBody.addEventListener('input', (e) => {
        if (e.target.classList.contains('row-monto')) {
            calculateTotal();
        }
    });

    // Handle "Agregar Fila"
    btnAgregar.addEventListener('click', () => {
        const currentPeriod = periodoInput.value;
        if (!currentPeriod) {
            showAlert('Por favor, seleccione un Periodo primero.');
            return;
        }
        
        const [year, month] = currentPeriod.split('-').map(Number);
        
        // Append an empty editable row
        const row = document.createElement('tr');
        
        // Create Day Select dropdown for user to choose which day of the month
        const daysInMonth = new Date(year, month, 0).getDate();
        let daySelectHtml = `<select class="row-dia-select" style="width: 70px;">`;
        for (let d = 1; d <= daysInMonth; d++) {
            daySelectHtml += `<option value="${d}">${d}</option>`;
        }
        daySelectHtml += `</select>`;

        const weekDays = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];

        row.innerHTML = `
            <td>${daySelectHtml}</td>
            <td class="row-fecha-label">--</td>
            <td><input type="text" class="row-motivo" value="Visita a clientes" required></td>
            <td><input type="text" class="row-ruta" placeholder="Detalle la ruta" required></td>
            <td>
                <select class="row-lugar" required>
                    <option value="SULLANA">SULLANA</option>
                    <option value="PAITA">PAITA</option>
                    <option value="TALARA">TALARA</option>
                    <option value="PIURA">PIURA</option>
                    <option value="PERIFERIA">PERIFERIA</option>
                    <option value="AUTOVENTA">AUTOVENTA</option>
                    <option value="OTRO">OTRO</option>
                </select>
                <input type="text" class="row-lugar-custom hidden-input mt-2" placeholder="Escriba lugar">
            </td>
            <td><input type="number" step="0.1" min="0" class="row-monto" value="0.00" style="width: 90px;" required></td>
            <td class="text-center"><button type="button" class="btn btn-danger btn-sm btn-delete-row">🗑️</button></td>
        `;

        calendarBody.appendChild(row);
        
        // Event Listeners for new row
        const diaSelect = row.querySelector('.row-dia-select');
        const fechaLabel = row.querySelector('.row-fecha-label');
        
        const updateDateLabel = () => {
            const d = Number(diaSelect.value);
            const dateObj = new Date(year, month - 1, d);
            const dayName = weekDays[dateObj.getDay()];
            fechaLabel.textContent = `${dayName} (${String(d).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year})`;
            if (dayName === 'DOMINGO') {
                row.classList.add('domingo');
            } else {
                row.classList.remove('domingo');
            }
        };

        diaSelect.addEventListener('change', updateDateLabel);
        updateDateLabel();

        const lugarSelect = row.querySelector('.row-lugar');
        const customLugarInput = row.querySelector('.row-lugar-custom');
        lugarSelect.addEventListener('change', () => {
            if (lugarSelect.value === 'OTRO') {
                customLugarInput.style.display = 'block';
                customLugarInput.setAttribute('required', 'true');
            } else {
                customLugarInput.style.display = 'none';
                customLugarInput.removeAttribute('required');
            }
        });

        row.querySelector('.btn-delete-row').onclick = () => {
            row.remove();
            calculateTotal();
        };

        calculateTotal();
    });

    // Handle Restablecer Predeterminados
    btnRestablecer.addEventListener('click', generateCalendar);

    // Generate Calendar Table based on selected Period & Vendor
    function generateCalendar() {
        const vendor = vendedorSelect.value;
        const period = periodoInput.value;

        if (!vendor || !period) {
            calendarBody.innerHTML = `<tr><td colspan="7" class="text-center py-4">Seleccione un Vendedor y Periodo para cargar sus rutas...</td></tr>`;
            calculateTotal();
            return;
        }

        const [year, month] = period.split('-').map(Number);
        const daysInMonth = new Date(year, month, 0).getDate();
        const weekDays = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];

        // Get defaults for selected vendor
        const defaults = (vendorsData[vendor] && vendorsData[vendor].defaults) || {};

        calendarBody.innerHTML = '';

        for (let d = 1; d <= daysInMonth; d++) {
            const dateObj = new Date(year, month - 1, d);
            const dayOfWeek = dateObj.getDay();
            const dayName = weekDays[dayOfWeek];

            const row = document.createElement('tr');
            if (dayName === 'DOMINGO') {
                row.classList.add('domingo');
            }

            // Load default route & place & amount if available
            let defRoute = '';
            let defLugar = '';
            let defMonto = 0.0;
            let motivo = 'Visita a clientes';

            if (dayName !== 'DOMINGO') {
                const dayDef = defaults[dayName];
                if (dayDef) {
                    defRoute = dayDef.ruta || '';
                    defLugar = dayDef.lugar || '';
                    defMonto = dayDef.monto || 0.0;
                }
            } else {
                motivo = '-';
            }

            // Check if lugar matches our list
            const standardLugares = ['SULLANA', 'PAITA', 'TALARA', 'PIURA', 'PERIFERIA', 'AUTOVENTA'];
            let placeSelectValue = '';
            let customPlaceValue = '';

            if (defLugar) {
                const upperLugar = defLugar.toUpperCase();
                if (standardLugares.includes(upperLugar)) {
                    placeSelectValue = upperLugar;
                } else {
                    placeSelectValue = 'OTRO';
                    customPlaceValue = defLugar;
                }
            } else if (dayName !== 'DOMINGO') {
                // If there's no default but it's not Sunday, let's default to SULLANA or empty
                placeSelectValue = 'SULLANA';
            }

            row.innerHTML = `
                <td><strong class="row-dia">${d}</strong></td>
                <td>${dayName} (${String(d).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year})</td>
                <td><input type="text" class="row-motivo" value="${motivo}" required></td>
                <td><input type="text" class="row-ruta" value="${defRoute}" placeholder="Detalle la ruta" required></td>
                <td>
                    <select class="row-lugar" required>
                        <option value="" ${placeSelectValue === '' ? 'selected' : ''}>-- Seleccione --</option>
                        <option value="SULLANA" ${placeSelectValue === 'SULLANA' ? 'selected' : ''}>SULLANA</option>
                        <option value="PAITA" ${placeSelectValue === 'PAITA' ? 'selected' : ''}>PAITA</option>
                        <option value="TALARA" ${placeSelectValue === 'TALARA' ? 'selected' : ''}>TALARA</option>
                        <option value="PIURA" ${placeSelectValue === 'PIURA' ? 'selected' : ''}>PIURA</option>
                        <option value="PERIFERIA" ${placeSelectValue === 'PERIFERIA' ? 'selected' : ''}>PERIFERIA</option>
                        <option value="AUTOVENTA" ${placeSelectValue === 'AUTOVENTA' ? 'selected' : ''}>AUTOVENTA</option>
                        <option value="OTRO" ${placeSelectValue === 'OTRO' ? 'selected' : ''}>OTRO</option>
                    </select>
                    <input type="text" class="row-lugar-custom ${placeSelectValue === 'OTRO' ? '' : 'hidden-input'} mt-2" value="${customPlaceValue}" placeholder="Escriba lugar">
                </td>
                <td><input type="number" step="0.1" min="0" class="row-monto" value="${defMonto.toFixed(2)}" style="width: 90px;" required></td>
                <td class="text-center"><button type="button" class="btn btn-danger btn-sm btn-delete-row">🗑️</button></td>
            `;

            calendarBody.appendChild(row);

            // Connect change handler for Lugar select
            const lugarSelect = row.querySelector('.row-lugar');
            const customLugarInput = row.querySelector('.row-lugar-custom');
            lugarSelect.addEventListener('change', () => {
                if (lugarSelect.value === 'OTRO') {
                    customLugarInput.style.display = 'block';
                    customLugarInput.setAttribute('required', 'true');
                } else {
                    customLugarInput.style.display = 'none';
                    customLugarInput.removeAttribute('required');
                }
            });

            // Connect delete handler
            row.querySelector('.btn-delete-row').onclick = () => {
                row.remove();
                calculateTotal();
            };
        }

        calculateTotal();
    }

    // Sum all row amounts
    function calculateTotal() {
        let total = 0.0;
        const montoInputs = calendarBody.querySelectorAll('.row-monto');
        montoInputs.forEach(input => {
            const val = parseFloat(input.value);
            if (!isNaN(val)) {
                total += val;
            }
        });
        document.getElementById('montoTotalVal').textContent = `S/. ${total.toFixed(2)}`;
    }

    // Submit Planilla Form
    form.addEventListener('submit', (e) => {
        e.preventDefault();

        const vendedorVal = vendedorSelect.value;
        const name = vendedorVal === 'NUEVO' ? nuevoVendedorInput.value.trim() : vendedorVal;
        const dni = dniInput.value.trim();
        const bancoSelect = document.getElementById('bancoSelect');
        const bancoVal = bancoSelect.value;
        const periodStr = periodoInput.value; // e.g. "2026-06"
        const fechaEmision = fechaEmisionInput.value;
        const planillaNro = parseInt(document.getElementById('planillaNroInput').value);

        if (!name || !dni || !bancoVal || !periodStr || !fechaEmision) {
            showAlert('Por favor, complete todos los campos obligatorios.');
            return;
        }

        // Collect rows
        const rows = [];
        const tableRows = calendarBody.querySelectorAll('tr');
        if (tableRows.length === 0 || tableRows[0].querySelector('td[colspan]')) {
            showAlert('Debe haber al menos un día/fila registrado.');
            return;
        }

        let hasError = false;
        tableRows.forEach(tr => {
            if (hasError) return;

            let d;
            const diaSelect = tr.querySelector('.row-dia-select');
            if (diaSelect) {
                d = parseInt(diaSelect.value);
            } else {
                d = parseInt(tr.querySelector('.row-dia').textContent);
            }

            const [year, month] = periodStr.split('-').map(Number);
            const dateObj = new Date(year, month - 1, d);
            const weekDays = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
            const dayOfWeek = weekDays[dateObj.getDay()];

            const motivo = tr.querySelector('.row-motivo').value.trim();
            const ruta = tr.querySelector('.row-ruta').value.trim();
            
            const lugarSelect = tr.querySelector('.row-lugar');
            let lugar = lugarSelect.value;
            if (lugar === 'OTRO') {
                lugar = tr.querySelector('.row-lugar-custom').value.trim().toUpperCase();
            }

            const monto = parseFloat(tr.querySelector('.row-monto').value);

            if (!motivo || !ruta || !lugar || isNaN(monto)) {
                showAlert(`Fila del día ${d}: Por favor, complete todos los campos de ruta y monto.`);
                hasError = true;
                return;
            }

            rows.push({
                dia: d,
                mes: month,
                ano: year,
                dia_semana: dayOfWeek,
                motivo: motivo,
                ruta: ruta,
                lugar: lugar,
                monto: monto
            });
        });

        if (hasError) return;

        // POST request to generate planilla
        const payload = {
            vendedor: name,
            dni: dni,
            banco: bancoVal,
            periodo: periodStr,
            fecha_emision: fechaEmision,
            planilla_nro_inicial: planillaNro,
            rows: rows
        };

        // Show loading state
        const submitBtn = form.querySelector('button[type="submit"]');
        const origText = submitBtn.textContent;
        submitBtn.textContent = 'Generando Planilla...';
        submitBtn.disabled = true;

        fetch('/api/generate-planilla', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(errData => {
                    throw new Error(errData.message || 'Error en el servidor.');
                });
            }
            return response.blob();
        })
        .then(blob => {
            // Trigger download of the generated file
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            // File naming convention matching app.py
            a.download = `Planilla_Movilidad_${name.replace(/ /g, '_')}_${bancoVal}_${periodStr.replace('-', '_')}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            
            showAlert('¡Planilla generada y descargada con éxito!');
            
            // Reload vendors DNI/Names mapping in case a new vendor was added
            if (vendedorVal === 'NUEVO') {
                location.reload();
            }
        })
        .catch(err => {
            console.error('Error generating planilla:', err);
            showAlert('Error al generar planilla: ' + err.message);
        })
        .finally(() => {
            submitBtn.textContent = origText;
            submitBtn.disabled = false;
        });
    });
}

// RRHH PANEL LOGIC
function initRrhhPanel() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const queueBody = document.getElementById('queueBody');
    const btnClear = document.getElementById('btnClearQueue');
    const btnConsolidar = document.getElementById('btnConsolidar');
    const resultsSection = document.getElementById('resultsSection');
    const resultsDetails = document.getElementById('resultsDetails');

    // Drag-and-drop events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', () => {
        handleFiles(fileInput.files);
    });

    function handleFiles(files) {
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (file.name.endsWith('.xlsx')) {
                // Check if already in queue
                if (!queueFiles.some(f => f.name === file.name && f.size === file.size)) {
                    queueFiles.push(file);
                }
            } else {
                showAlert(`El archivo "${file.name}" no es un Excel válido (.xlsx)`);
            }
        }
        updateQueueUI();
    }

    function updateQueueUI() {
        const emptyRow = document.getElementById('emptyQueueRow');
        if (queueFiles.length === 0) {
            if (!emptyRow) {
                queueBody.innerHTML = `<tr id="emptyQueueRow"><td colspan="5" class="text-center py-4">No hay archivos en cola. Sube planillas para comenzar.</td></tr>`;
            }
            btnClear.disabled = true;
            btnConsolidar.disabled = true;
            return;
        }

        if (emptyRow) {
            queueBody.innerHTML = '';
        }

        // Render queue rows
        // We can't easily parse worker details on client-side Excel without library, 
        // so we'll show filename and let server parse it. We can extract details from name or write placeholders
        queueBody.innerHTML = '';
        queueFiles.forEach((file, index) => {
            const tr = document.createElement('tr');
            
            // Try to extract worker name, bank and period from filename
            // Planilla_Movilidad_[Vendedor]_[Banco]_[Periodo].xlsx
            let vendorName = 'Analizando...';
            let periodName = 'Analizando...';
            const nameMatch = file.name.match(/Planilla_Movilidad_(.+)_([A-Z]+)_(\d{4}_\d{2})\.xlsx/);
            if (nameMatch) {
                vendorName = nameMatch[1].replace(/_/g, ' ') + ` (${nameMatch[2]})`;
                periodName = nameMatch[3].replace('_', '-');
            }

            tr.innerHTML = `
                <td><strong>${file.name}</strong></td>
                <td>${vendorName}</td>
                <td>${periodName}</td>
                <td class="text-center">-</td>
                <td class="text-center">
                    <button type="button" class="btn btn-danger btn-sm" onclick="removeQueueItem(${index})">Eliminar</button>
                </td>
            `;
            queueBody.appendChild(tr);
        });

        btnClear.disabled = false;
        btnConsolidar.disabled = false;
    }

    // Global hook for item deletion
    window.removeQueueItem = function(index) {
        queueFiles.splice(index, 1);
        updateQueueUI();
    };

    btnClear.addEventListener('click', () => {
        queueFiles = [];
        updateQueueUI();
        resultsSection.classList.add('hidden-input');
    });

    // Submit files to Consolidation API — two-step: process → download
    btnConsolidar.addEventListener('click', () => {
        if (queueFiles.length === 0) return;

        const formData = new FormData();
        queueFiles.forEach(file => { formData.append('files', file); });

        // ── Mostrar barra de progreso ──
        resultsSection.classList.remove('hidden-input');
        document.getElementById('successAlert').style.display = 'none';
        resultsDetails.innerHTML = `
            <div id="progressContainer" style="margin: 12px 0;">
                <p id="progressStatus" style="margin-bottom: 8px; font-weight: 500;">
                    ⏳ Procesando ${queueFiles.length} planilla(s)...
                </p>
                <div style="background: #e2e8f0; border-radius: 8px; overflow: hidden; height: 12px;">
                    <div id="progressBar" style="height: 100%; width: 0%; background: linear-gradient(90deg, #6366f1, #8b5cf6); border-radius: 8px; transition: width 0.4s ease; animation: progressPulse 1.5s ease-in-out infinite;"></div>
                </div>
            </div>
        `;
        // Animar barra mientras espera respuesta
        let pct = 5;
        const progressBar = document.getElementById('progressBar');
        const progressTimer = setInterval(() => {
            if (pct < 85) { pct += Math.random() * 8; }
            if (progressBar) progressBar.style.width = Math.min(pct, 85) + '%';
        }, 400);

        btnConsolidar.textContent = 'Procesando...';
        btnConsolidar.disabled = true;
        btnClear.disabled = true;

        // Paso 1: Procesar y obtener token
        fetch('/api/generate-consolidado', {
            method: 'POST',
            body: formData
        })
        .then(res => res.json())
        .then(data => {
            clearInterval(progressTimer);
            if (progressBar) progressBar.style.width = '100%';

            if (data.status !== 'success') {
                throw { message: data.message || 'Error desconocido.', logs: data.logs || [] };
            }

            // Paso 2: Descargar el archivo por token
            document.getElementById('progressStatus').textContent = '📥 Descargando archivo...';
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = `/api/download-consolidado/${data.token}`;
            a.download = data.filename || 'Consolidado_Pasajes_FFVV.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            // Mostrar resultados
            document.getElementById('successAlert').style.display = 'block';
            document.getElementById('successAlert').innerHTML =
                `<strong>✅ ¡Consolidado generado!</strong> Se procesaron <strong>${data.consolidated_count}</strong> planilla(s). Descargando <em>${data.filename}</em>...`;

            let logHtml = '<br><strong>Detalles del procesamiento:</strong><br>';
            (data.logs || []).forEach(log => {
                const color = log.startsWith('[OK]') ? '#16a34a' : '#dc2626';
                logHtml += `<span style="color:${color}">&gt; ${log}</span><br>`;
            });
            resultsDetails.innerHTML = document.getElementById('successAlert').outerHTML + logHtml;
            document.getElementById('successAlert').style.display = 'none';

            queueFiles = [];
            updateQueueUI();
        })
        .catch(err => {
            clearInterval(progressTimer);
            console.error('Error generando consolidado:', err);
            let errHtml = `<span style="color:#ef4444"><strong>Error en el proceso:</strong><br>${err.message || String(err)}</span>`;
            if (err.logs && err.logs.length > 0) {
                errHtml += '<br><strong>Detalles:</strong><br>';
                err.logs.forEach(log => { errHtml += `&gt; ${log}<br>`; });
            }
            resultsDetails.innerHTML = errHtml;
            document.getElementById('successAlert').style.display = 'none';
        })
        .finally(() => {
            btnConsolidar.textContent = '📥 Generar y Descargar Consolidado';
            btnConsolidar.disabled = false;
            btnClear.disabled = false;
        });
    });
}
