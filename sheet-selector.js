// Gestión de Selector de Hojas del Excel
let loadedWorkbook = null;
let availableSheets = [];

// Variables globales para paginación
let currentTablePage = 1;
let itemsPerTablePage = 5;
let filteredTableData = [];

// Función para limpiar completamente todos los datos al cargar un nuevo archivo
function clearAllData() {
    console.log('🧹 Limpiando todos los datos anteriores...');
    
    // Resetear variables globales
    loadedWorkbook = null;
    availableSheets = [];
    filteredTableData = [];
    currentTablePage = 1;
    
    // Limpiar la tabla
    const tbody = document.getElementById('testTableBody');
    if (tbody) {
        tbody.innerHTML = '';
    }
    
    // Resetear filtros
    const filterAplicacion = document.getElementById('filterAplicacion');
    const filterVertical = document.getElementById('filterVertical');
    const filterEpica = document.getElementById('filterEpica');
    const filterHistoria = document.getElementById('filterHistoria');
    
    if (filterAplicacion) filterAplicacion.innerHTML = '<option value="">Todas las Aplicaciones</option>';
    if (filterVertical) filterVertical.innerHTML = '<option value="">Todos los Verticales</option>';
    if (filterEpica) filterEpica.innerHTML = '<option value="">Todas las Épicas</option>';
    if (filterHistoria) filterHistoria.innerHTML = '<option value="">Todas las Historias de Usuario</option>';
    
    // Resetear KPI cards a 0
    const kpiIds = ['plannedTests', 'successfulTests', 'reviewTests', 'pendingTests', 'approvedTests', 'dismissedTests'];
    kpiIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) element.textContent = '0';
    });
    
    // Resetear velocímetro y círculos de progreso
    updateProgressPercentage(0);
    
    const progressIds = [
        'designProgressPercentage',
        'reviewProgressPercentage', 
        'pendingProgressPercentage',
        'approvalProgressPercentage',
        'pendingApprovalProgressPercentage'
    ];
    
    progressIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) element.textContent = '0%';
        
        const circleId = id.replace('Percentage', 'Circle');
        const circle = document.getElementById(circleId);
        if (circle) {
            const progressCircle = circle.querySelector('circle:last-child');
            if (progressCircle) {
                progressCircle.style.strokeDashoffset = '264';
            }
        }
    });
    
    console.log('✅ Datos limpiados correctamente');
}

// Función para esperar a que XLSX esté disponible
function waitForXLSX(callback) {
    if (typeof XLSX !== 'undefined') {
        console.log('✅ XLSX está disponible');
        callback();
    } else {
        console.log('⏳ Esperando a que XLSX se cargue...');
        setTimeout(() => waitForXLSX(callback), 100);
    }
}

// Verificar que XLSX esté cargado al iniciar
waitForXLSX(() => {
    console.log('✅ XLSX cargado correctamente, versión:', XLSX.version);
});

// Modificar el manejador del archivo Excel
function handleExcelFileWithSheets(file) {
    // Limpiar todos los datos anteriores antes de cargar el nuevo archivo
    clearAllData();
    
    // Verificar que XLSX esté disponible antes de procesar
    if (typeof XLSX === 'undefined') {
        showNotification('Error: La librería XLSX no está cargada. Por favor, recarga la página.', 'error');
        console.error('[ERROR] XLSX no está disponible');
        return;
    }
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Guardar el workbook globalmente
            loadedWorkbook = workbook;
            availableSheets = workbook.SheetNames;
            
            // Si hay más de una hoja, mostrar selector
            if (availableSheets.length > 1) {
                showSheetSelector(availableSheets);
            } else if (availableSheets.length === 1) {
                // Si solo hay una hoja, cargarla directamente
                loadSheetData(availableSheets[0]);
            } else {
                showNotification('El archivo no contiene hojas válidas', 'error');
            }
            
        } catch (error) {
            console.error('Error al procesar el archivo Excel:', error);
            showNotification('Error al procesar el archivo: ' + error.message, 'error');
        }
    };
    
    reader.onerror = function() {
        showNotification('Error al leer el archivo', 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

// Mostrar modal de selección de hojas
function showSheetSelector(sheets) {
    const modal = document.getElementById('sheetSelectorModal');
    const body = document.getElementById('sheetSelectorBody');
    
    if (!modal || !body) {
        console.error('Modal de selección no encontrado');
        return;
    }
    
    // Limpiar contenido previo
    body.innerHTML = '';
    
    // Crear elementos para cada hoja
    sheets.forEach((sheetName, index) => {
        const sheetItem = document.createElement('div');
        sheetItem.className = 'sheet-item';
        sheetItem.onclick = () => selectSheet(sheetName);
        
        sheetItem.innerHTML = `
            <div class="sheet-item-icon">
                <i class="fas fa-file-excel"></i>
            </div>
            <div class="sheet-item-info">
                <div class="sheet-item-name">${sheetName}</div>
            </div>
            <div class="sheet-item-arrow">
                <i class="fas fa-chevron-right"></i>
            </div>
        `;
        
        body.appendChild(sheetItem);
    });
    
    // Mostrar modal
    modal.style.display = 'flex';
}

// Seleccionar y cargar una hoja específica
function selectSheet(sheetName) {
    // Ocultar modal
    const modal = document.getElementById('sheetSelectorModal');
    if (modal) {
        modal.style.display = 'none';
    }
    
    // Cargar datos de la hoja seleccionada
    loadSheetData(sheetName);
    
    // Mostrar notificación
    showNotification(`Proyecto "${sheetName}" cargado exitosamente`, 'success');
    
    // Mostrar botón para cambiar de proyecto
    const selectButton = document.getElementById('selectProjectButton');
    if (selectButton) {
        selectButton.style.display = 'flex';
    }
}

// Cargar datos de una hoja específica
function loadSheetData(sheetName) {
    if (!loadedWorkbook || !loadedWorkbook.Sheets[sheetName]) {
        showNotification('Error: Hoja no encontrada', 'error');
        return;
    }
    
    try {
        // Convertir la hoja a JSON
        const sheetData = XLSX.utils.sheet_to_json(loadedWorkbook.Sheets[sheetName]);
        
        console.log(`Datos cargados de la hoja "${sheetName}":`, sheetData);
        
        // Aquí debes procesar los datos según la estructura de tu Excel
        // Por ejemplo:
        processSheetData(sheetData, sheetName);
        
    } catch (error) {
        console.error('Error al procesar la hoja:', error);
        showNotification('Error al procesar los datos de la hoja', 'error');
    }
}

// Procesar datos de la hoja
function processSheetData(data, sheetName) {
    // Actualizar el nombre del proyecto en el header
    const projectNameElement = document.getElementById('projectNameDetail');
    if (projectNameElement) {
        projectNameElement.textContent = sheetName;
    }
    
    // Actualizar métricas dinámicas (KPIs, progreso general y porcentajes individuales)
    updateDynamicMetrics(data);
    
    // Actualizar tabla
    updateProgressTable(data);
}

// Actualizar un KPI
function updateKPI(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.textContent = value;
    }
}

// Actualizar porcentaje de progreso general
// Actualizar tabla de progreso
function updateProgressTable(data) {
    const tbody = document.getElementById('testTableBody');
    if (!tbody) return;
    
    // Guardar los datos globalmente para filtrado
    window.currentTableData = data;
    
    // Inicializar filteredTableData con todos los datos
    filteredTableData = [...data];
    
    // Resetear a la primera página
    currentTablePage = 1;
    
    // Actualizar todos los filtros
    updateAllFilters(data);
    
    // Configurar eventos de filtrado
    setupTableFilters();
    
    // Renderizar con paginación (respetará el itemsPerTablePage que por defecto es 5)
    renderTableWithPagination();
}

// Actualizar todos los filtros dinámicamente (solo al cargar datos iniciales)
function updateAllFilters(data) {
    updateAplicacionFilter(data);
    updateCascadeFilters(); // Actualizar filtros en cascada
}

// Actualizar filtro de aplicaciones (siempre muestra todos)
function updateAplicacionFilter(data) {
    const filterSelect = document.getElementById('filterAplicacion');
    if (!filterSelect) return;
    
    const currentValue = filterSelect.value; // Guardar selección actual
    
    // Crear mapa para normalizar (mantener la primera ocurrencia con su formato original)
    const aplicacionesMap = new Map();
    data.forEach(row => {
        const app = row['app'];
        if (app && app.toString().trim() !== '') {
            const normalizedKey = app.toString().trim().toLowerCase();
            if (!aplicacionesMap.has(normalizedKey)) {
                aplicacionesMap.set(normalizedKey, app.toString().trim());
            }
        }
    });
    
    const aplicaciones = Array.from(aplicacionesMap.values());
    aplicaciones.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
    
    filterSelect.innerHTML = '<option value="">Todas las Aplicaciones</option>';
    aplicaciones.forEach(app => {
        const option = document.createElement('option');
        option.value = app;
        option.textContent = app;
        filterSelect.appendChild(option);
    });
    
    // Restaurar selección si existe
    if (currentValue && aplicaciones.includes(currentValue)) {
        filterSelect.value = currentValue;
    }
}

// Actualizar filtros en cascada basándose en selecciones previas
function updateCascadeFilters() {
    if (!window.currentTableData) return;
    
    const filterAplicacion = document.getElementById('filterAplicacion');
    const filterVertical = document.getElementById('filterVertical');
    const filterEpica = document.getElementById('filterEpica');
    const filterHistoria = document.getElementById('filterHistoria');
    const filterEstadoHU = document.getElementById('filterEstadoHU');
    
    const selectedAplicacion = filterAplicacion ? filterAplicacion.value : '';
    const selectedVertical = filterVertical ? filterVertical.value : '';
    const selectedEpica = filterEpica ? filterEpica.value : '';
    const selectedHistoria = filterHistoria ? filterHistoria.value : '';
    
    // Guardar valores actuales
    const currentVertical = selectedVertical;
    const currentEpica = selectedEpica;
    const currentHistoria = selectedHistoria;
    const currentEstadoHU = filterEstadoHU ? filterEstadoHU.value : '';
    
    // Filtrar datos según las selecciones previas
    let filteredData = window.currentTableData;
    
    // Actualizar Vertical basándose en Aplicación
    if (selectedAplicacion) {
        filteredData = filteredData.filter(row => row['app'] === selectedAplicacion);
    }
    updateVerticalFilter(filteredData, currentVertical);
    
    // Actualizar Épica basándose en Aplicación + Vertical
    filteredData = window.currentTableData;
    if (selectedAplicacion) {
        filteredData = filteredData.filter(row => row['app'] === selectedAplicacion);
    }
    if (selectedVertical) {
        filteredData = filteredData.filter(row => row['vertical'] === selectedVertical);
    }
    updateEpicaFilter(filteredData, currentEpica);
    
    // Actualizar Historia basándose en Aplicación + Vertical + Épica
    filteredData = window.currentTableData;
    if (selectedAplicacion) {
        filteredData = filteredData.filter(row => row['app'] === selectedAplicacion);
    }
    if (selectedVertical) {
        filteredData = filteredData.filter(row => row['vertical'] === selectedVertical);
    }
    if (selectedEpica) {
        filteredData = filteredData.filter(row => row['epica'] === selectedEpica);
    }
    updateHistoriaFilter(filteredData, currentHistoria);
    
    // Actualizar Estado HU basándose en Aplicación + Vertical + Épica + Historia
    filteredData = window.currentTableData;
    if (selectedAplicacion) {
        filteredData = filteredData.filter(row => row['app'] === selectedAplicacion);
    }
    if (selectedVertical) {
        filteredData = filteredData.filter(row => row['vertical'] === selectedVertical);
    }
    if (selectedEpica) {
        filteredData = filteredData.filter(row => row['epica'] === selectedEpica);
    }
    if (selectedHistoria) {
        filteredData = filteredData.filter(row => row['hu'] === selectedHistoria);
    }
    updateEstadoHUFilter(filteredData, currentEstadoHU);
}

// Actualizar filtro de verticales
function updateVerticalFilter(data, preserveValue = null) {
    const filterSelect = document.getElementById('filterVertical');
    if (!filterSelect) return;
    
    // Crear mapa para normalizar (mantener la primera ocurrencia con su formato original)
    const verticalesMap = new Map();
    data.forEach(row => {
        const vertical = row['vertical'];
        if (vertical && vertical.toString().trim() !== '') {
            const normalizedKey = vertical.toString().trim().toLowerCase();
            if (!verticalesMap.has(normalizedKey)) {
                verticalesMap.set(normalizedKey, vertical.toString().trim());
            }
        }
    });
    
    const verticales = Array.from(verticalesMap.values());
    verticales.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
    
    filterSelect.innerHTML = '<option value="">Todos los Verticales</option>';
    verticales.forEach(vertical => {
        const option = document.createElement('option');
        option.value = vertical;
        option.textContent = vertical;
        filterSelect.appendChild(option);
    });
    
    // Restaurar valor si existe en las opciones disponibles
    if (preserveValue && verticales.includes(preserveValue)) {
        filterSelect.value = preserveValue;
    }
}

// Actualizar filtro de épicas
function updateEpicaFilter(data, preserveValue = null) {
    const filterSelect = document.getElementById('filterEpica');
    if (!filterSelect) return;
    
    // Crear mapa para normalizar (mantener la primera ocurrencia con su formato original)
    const epicasMap = new Map();
    data.forEach(row => {
        const epica = row['epica'];
        if (epica && epica.toString().trim() !== '') {
            const normalizedKey = epica.toString().trim().toLowerCase();
            if (!epicasMap.has(normalizedKey)) {
                epicasMap.set(normalizedKey, epica.toString().trim());
            }
        }
    });
    
    const epicas = Array.from(epicasMap.values());
    epicas.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
    
    filterSelect.innerHTML = '<option value="">Todas las Épicas</option>';
    epicas.forEach(epica => {
        const option = document.createElement('option');
        option.value = epica;
        option.textContent = epica;
        filterSelect.appendChild(option);
    });
    
    // Restaurar valor si existe en las opciones disponibles
    if (preserveValue && epicas.includes(preserveValue)) {
        filterSelect.value = preserveValue;
    }
}

// Actualizar filtro de historias de usuario
function updateHistoriaFilter(data, preserveValue = null) {
    const filterSelect = document.getElementById('filterHistoria');
    if (!filterSelect) return;
    
    // Crear mapa para normalizar (mantener la primera ocurrencia con su formato original)
    const historiasMap = new Map();
    data.forEach(row => {
        const historia = row['hu'];
        if (historia && historia.toString().trim() !== '') {
            const normalizedKey = historia.toString().trim().toLowerCase();
            if (!historiasMap.has(normalizedKey)) {
                historiasMap.set(normalizedKey, historia.toString().trim());
            }
        }
    });
    
    const historias = Array.from(historiasMap.values());
    historias.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
    
    filterSelect.innerHTML = '<option value="">Todas las Historias de Usuario</option>';
    historias.forEach(historia => {
        const option = document.createElement('option');
        option.value = historia;
        option.textContent = historia;
        filterSelect.appendChild(option);
    });
    
    // Restaurar valor si existe en las opciones disponibles
    if (preserveValue && historias.includes(preserveValue)) {
        filterSelect.value = preserveValue;
    }
}

// Actualizar filtro de Estado HU
function updateEstadoHUFilter(data, preserveValue = null) {
    const filterSelect = document.getElementById('filterEstadoHU');
    if (!filterSelect) return;
    
    // Crear mapa para normalizar (mantener la primera ocurrencia con su formato original)
    const estadosMap = new Map();
    data.forEach(row => {
        const estado = row['estado_hu'];
        if (estado && estado.toString().trim() !== '') {
            const normalizedKey = estado.toString().trim().toLowerCase();
            if (!estadosMap.has(normalizedKey)) {
                estadosMap.set(normalizedKey, estado.toString().trim());
            }
        }
    });
    
    const estados = Array.from(estadosMap.values());
    estados.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
    
    filterSelect.innerHTML = '<option value="">Todos los Estados HU</option>';
    estados.forEach(estado => {
        const option = document.createElement('option');
        option.value = estado;
        option.textContent = estado;
        filterSelect.appendChild(option);
    });
    
    // Restaurar valor si existe en las opciones disponibles
    if (preserveValue && estados.includes(preserveValue)) {
        filterSelect.value = preserveValue;
    }
}

// Configurar filtros de tabla
function setupTableFilters() {
    const filterAplicacion = document.getElementById('filterAplicacion');
    const filterVertical = document.getElementById('filterVertical');
    const filterEpica = document.getElementById('filterEpica');
    const filterHistoria = document.getElementById('filterHistoria');
    const filterEstadoHU = document.getElementById('filterEstadoHU');
    const clearFiltersBtn = document.getElementById('clearFiltersBtn');
    const itemsPerPageSelect = document.getElementById('itemsPerPageSelect');
    const prevPageBtn = document.getElementById('prevPageBtn');
    const nextPageBtn = document.getElementById('nextPageBtn');
    
    if (filterAplicacion) {
        filterAplicacion.removeEventListener('change', handleFilterChange);
        filterAplicacion.addEventListener('change', handleFilterChange);
    }
    
    if (filterVertical) {
        filterVertical.removeEventListener('change', handleFilterChange);
        filterVertical.addEventListener('change', handleFilterChange);
    }
    
    if (filterEpica) {
        filterEpica.removeEventListener('change', handleFilterChange);
        filterEpica.addEventListener('change', handleFilterChange);
    }
    
    if (filterHistoria) {
        filterHistoria.removeEventListener('change', handleFilterChange);
        filterHistoria.addEventListener('change', handleFilterChange);
    }
    
    if (filterEstadoHU) {
        filterEstadoHU.removeEventListener('change', handleFilterChange);
        filterEstadoHU.addEventListener('change', handleFilterChange);
    }
    
    if (clearFiltersBtn) {
        clearFiltersBtn.removeEventListener('click', clearAllFilters);
        clearFiltersBtn.addEventListener('click', clearAllFilters);
    }
    
    // Configurar paginación
    if (itemsPerPageSelect) {
        itemsPerPageSelect.removeEventListener('change', handleItemsPerPageChange);
        itemsPerPageSelect.addEventListener('change', handleItemsPerPageChange);
    }
    
    if (prevPageBtn) {
        prevPageBtn.removeEventListener('click', goToTablePrevPage);
        prevPageBtn.addEventListener('click', goToTablePrevPage);
    }
    
    if (nextPageBtn) {
        nextPageBtn.removeEventListener('click', goToTableNextPage);
        nextPageBtn.addEventListener('click', goToTableNextPage);
    }
}

// Manejar cambio en items por página
function handleItemsPerPageChange(event) {
    changeTableItemsPerPage(event.target.value);
}

// Limpiar todos los filtros
function clearAllFilters() {
    const filterAplicacion = document.getElementById('filterAplicacion');
    const filterVertical = document.getElementById('filterVertical');
    const filterEpica = document.getElementById('filterEpica');
    const filterHistoria = document.getElementById('filterHistoria');
    const filterEstadoHU = document.getElementById('filterEstadoHU');
    
    // Resetear todos los filtros a su valor por defecto
    if (filterAplicacion) filterAplicacion.value = '';
    if (filterVertical) filterVertical.value = '';
    if (filterEpica) filterEpica.value = '';
    if (filterHistoria) filterHistoria.value = '';
    if (filterEstadoHU) filterEstadoHU.value = '';
    
    // Actualizar todos los filtros con datos completos
    if (window.currentTableData) {
        updateAllFilters(window.currentTableData);
    }
    
    // Aplicar filtros (mostrar todos los datos)
    applyTableFilters();
}

// Manejar cambio en cualquier filtro
function handleFilterChange(event) {
    // Primero actualizar la cascada de filtros
    updateCascadeFilters();
    // Luego aplicar los filtros a la tabla
    applyTableFilters();
}

// Aplicar filtros a la tabla
function applyTableFilters() {
    if (!window.currentTableData) return;
    
    const filterAplicacion = document.getElementById('filterAplicacion');
    const filterVertical = document.getElementById('filterVertical');
    const filterEpica = document.getElementById('filterEpica');
    const filterHistoria = document.getElementById('filterHistoria');
    const filterEstadoHU = document.getElementById('filterEstadoHU');
    
    const selectedAplicacion = filterAplicacion ? filterAplicacion.value : '';
    const selectedVertical = filterVertical ? filterVertical.value : '';
    const selectedEpica = filterEpica ? filterEpica.value : '';
    const selectedHistoria = filterHistoria ? filterHistoria.value : '';
    const selectedEstadoHU = filterEstadoHU ? filterEstadoHU.value : '';
    
    // Filtrar datos
    let filteredData = window.currentTableData;
    
    // Filtrar por aplicación
    if (selectedAplicacion) {
        filteredData = filteredData.filter(row => row['app'] === selectedAplicacion);
    }
    
    // Filtrar por vertical
    if (selectedVertical) {
        filteredData = filteredData.filter(row => row['vertical'] === selectedVertical);
    }
    
    // Filtrar por épica
    if (selectedEpica) {
        filteredData = filteredData.filter(row => row['epica'] === selectedEpica);
    }
    
    // Filtrar por historia de usuario
    if (selectedHistoria) {
        filteredData = filteredData.filter(row => row['hu'] === selectedHistoria);
    }
    
    // Filtrar por estado HU
    if (selectedEstadoHU) {
        filteredData = filteredData.filter(row => row['estado_hu'] === selectedEstadoHU);
    }
    
    // Guardar datos filtrados globalmente
    filteredTableData = filteredData;
    
    // Renderizar tabla con paginación
    renderTableWithPagination();
    
    // Actualizar métricas dinámicamente basándose en los datos filtrados
    updateDynamicMetrics(filteredData);
}

// Renderizar tabla con paginación
function renderTableWithPagination() {
    const tbody = document.getElementById('testTableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    const totalItems = filteredTableData.length;
    
    if (totalItems === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="11" style="text-align: center; padding: 2rem; color: #999;">No se encontraron resultados</td>';
        tbody.appendChild(tr);
        updateTablePaginationControls(0, 0);
        return;
    }
    
    // Calcular paginación
    const totalPages = itemsPerTablePage === 100 ? 1 : Math.ceil(totalItems / itemsPerTablePage);
    
    // Ajustar página actual si excede el total
    if (currentTablePage > totalPages) {
        currentTablePage = totalPages;
    }
    if (currentTablePage < 1) {
        currentTablePage = 1;
    }
    
    // Calcular índices
    const startIndex = (currentTablePage - 1) * itemsPerTablePage;
    const endIndex = itemsPerTablePage === 100 ? totalItems : Math.min(startIndex + itemsPerTablePage, totalItems);
    
    // Renderizar solo los items de la página actual
    const pageData = filteredTableData.slice(startIndex, endIndex);
    
    pageData.forEach(row => {
        const tr = document.createElement('tr');
        
        // Obtener los valores directamente del Excel
        const cpDiseñados = parseInt(row['cp_diseñado_qa']) || 0;
        const cpAprobadoDev = parseInt(row['cp_aprobado_dev']) || 0;
        const cpAprobadoNeg = parseInt(row['cp_aprobado_neg']) || 0;
        const cpPendienteDev = parseInt(row['cp_pendiente_dev']) || 0;
        const cpPendienteNeg = parseInt(row['cp_pendiente_neg']) || 0;
        const estadoHU = row['estado_hu'] || '-';
        const link = row['link'] || '';
        
        // Crear el HTML del link
        let linkHTML = '<td style="text-align: center;">-</td>';
        if (link && link.trim() !== '') {
            linkHTML = `<td style="text-align: center;">
                <a href="${link.trim()}" target="_blank" rel="noopener noreferrer" class="link-icon" title="Ver enlace">
                    <i class="fas fa-external-link-alt"></i>
                </a>
            </td>`;
        }
        
        tr.innerHTML = `
            <td>${row['app'] || '-'}</td>
            <td>${row['vertical'] || '-'}</td>
            <td>${row['epica'] || '-'}</td>
            <td>${row['hu'] || '-'}</td>
            <td>${estadoHU}</td>
            <td>${cpDiseñados}</td>
            <td>${cpAprobadoDev}</td>
            <td>${cpPendienteDev}</td>
            <td>${cpAprobadoNeg}</td>
            <td>${cpPendienteNeg}</td>
            ${linkHTML}
        `;
        tbody.appendChild(tr);
    });
    
    // Actualizar controles de paginación
    updateTablePaginationControls(totalItems, totalPages);
}

// Actualizar controles de paginación
function updateTablePaginationControls(totalItems, totalPages) {
    const paginationInfo = document.getElementById('paginationInfo');
    const prevBtn = document.getElementById('prevPageBtn');
    const nextBtn = document.getElementById('nextPageBtn');
    const pageNumbers = document.getElementById('pageNumbers');
    
    if (!paginationInfo || !prevBtn || !nextBtn || !pageNumbers) return;
    
    // Actualizar información de paginación
    const startItem = totalItems === 0 ? 0 : (currentTablePage - 1) * itemsPerTablePage + 1;
    const endItem = Math.min(currentTablePage * itemsPerTablePage, totalItems);
    paginationInfo.textContent = `Mostrando ${startItem}-${endItem} de ${totalItems} registros`;
    
    // Actualizar botones
    prevBtn.disabled = currentTablePage === 1;
    nextBtn.disabled = currentTablePage === totalPages || totalPages === 0;
    
    // Actualizar números de página
    pageNumbers.innerHTML = '';
    
    if (totalPages <= 7) {
        // Mostrar todas las páginas
        for (let i = 1; i <= totalPages; i++) {
            addPageNumber(i, pageNumbers);
        }
    } else {
        // Mostrar páginas con elipsis
        addPageNumber(1, pageNumbers);
        
        if (currentTablePage > 3) {
            pageNumbers.appendChild(createEllipsis());
        }
        
        const startPage = Math.max(2, currentTablePage - 1);
        const endPage = Math.min(totalPages - 1, currentTablePage + 1);
        
        for (let i = startPage; i <= endPage; i++) {
            addPageNumber(i, pageNumbers);
        }
        
        if (currentTablePage < totalPages - 2) {
            pageNumbers.appendChild(createEllipsis());
        }
        
        addPageNumber(totalPages, pageNumbers);
    }
}

// Agregar número de página
function addPageNumber(pageNum, container) {
    const pageBtn = document.createElement('button');
    pageBtn.className = 'page-number' + (pageNum === currentTablePage ? ' active' : '');
    pageBtn.textContent = pageNum;
    pageBtn.onclick = () => goToTablePage(pageNum);
    container.appendChild(pageBtn);
}

// Crear elipsis
function createEllipsis() {
    const ellipsis = document.createElement('span');
    ellipsis.className = 'page-ellipsis';
    ellipsis.textContent = '...';
    return ellipsis;
}

// Navegar a página específica
function goToTablePage(page) {
    currentTablePage = page;
    renderTableWithPagination();
}

// Ir a página anterior
function goToTablePrevPage() {
    if (currentTablePage > 1) {
        currentTablePage--;
        renderTableWithPagination();
    }
}

// Ir a página siguiente
function goToTableNextPage() {
    const totalPages = itemsPerTablePage === 100 ? 1 : Math.ceil(filteredTableData.length / itemsPerTablePage);
    if (currentTablePage < totalPages) {
        currentTablePage++;
        renderTableWithPagination();
    }
}

// Cambiar items por página
function changeTableItemsPerPage(newItemsPerPage) {
    itemsPerTablePage = parseInt(newItemsPerPage);
    currentTablePage = 1; // Resetear a primera página
    renderTableWithPagination();
}

// Actualizar métricas dinámicamente según datos filtrados
function updateDynamicMetrics(data) {
    // Contadores usando Set para valores únicos
    const aplicacionesUnicas = new Set();
    const verticalesUnicas = new Set();
    const epicasUnicas = new Set();
    const historiasUnicas = new Set();
    
    // Contadores para sumas
    let totalCriterios = 0;
    let totalPlanificado = 0;
    let totalDiseñados = 0;
    let totalRevisados = 0;
    let totalAprobados = 0;
    let totalPendientes = 0;
    
    // Procesar datos filtrados
    data.forEach(row => {
        // Contar aplicaciones únicas
        if (row['app'] && row['app'].toString().trim() !== '') {
            aplicacionesUnicas.add(row['app']);
        }
        
        // Contar verticales únicas
        if (row['vertical'] && row['vertical'].toString().trim() !== '') {
            verticalesUnicas.add(row['vertical']);
        }
        
        // Contar épicas únicas
        if (row['epica'] && row['epica'].toString().trim() !== '') {
            epicasUnicas.add(row['epica']);
        }
        
        // Contar historias únicas (cada fila es una HU)
        if (row['hu'] && row['hu'].toString().trim() !== '') {
            historiasUnicas.add(row['hu']);
        }
        
        // Sumar casos de prueba
        const planificado = parseInt(row['cp_planificado']) || 0;
        const diseñados = parseInt(row['cp_diseñado_qa']) || 0;
        const aprobadoDev = parseInt(row['cp_aprobado_dev']) || 0;
        const aprobadoNeg = parseInt(row['cp_aprobado_neg']) || 0;
        const pendienteDev = parseInt(row['cp_pendiente_dev']) || 0;
        const pendienteNeg = parseInt(row['cp_pendiente_neg']) || 0;
        
        totalPlanificado += planificado;
        totalDiseñados += diseñados;
        totalRevisados += aprobadoDev;  // Ahora es aprobados por desarrollo
        totalAprobados += aprobadoNeg;  // Ahora es aprobados por negocio
        totalPendientes += pendienteNeg;  // Pendientes de negocio
    });
    
    // Calcular porcentajes
    const porcentajeAvance = totalDiseñados > 0 ? Math.round((totalAprobados / totalDiseñados) * 100) : 0;
    
    // Calcular porcentajes individuales según las nuevas columnas del Excel
    
    // % Avance Diseño CP (Equipo QA): cp_planificado es el 100% y cp_diseñado_qa su avance
    const porcentajeDiseño = totalPlanificado > 0 ? Math.round((totalDiseñados / totalPlanificado) * 100) : 0;
    
    // % Avance Aprobaciones CP (Equipo Desarrollo): cp_diseñado_qa es el 100% y cp_aprobado_dev su avance
    const porcentajeRevisiones = totalDiseñados > 0 ? Math.round((totalRevisados / totalDiseñados) * 100) : 0;
    
    // % Avance Aprobaciones CP (Equipo Negocio): cp_diseñado_qa es el 100% y cp_aprobado_neg su avance
    const porcentajeAprobaciones = totalDiseñados > 0 ? Math.round((totalAprobados / totalDiseñados) * 100) : 0;
    
    // % Pendiente Aprobación (Equipo Desarrollo): usar cp_pendiente_dev directamente del Excel
    // Calculamos el porcentaje: (suma de cp_pendiente_dev / suma de cp_diseñado_qa) * 100
    let totalPendientesDev = 0;
    data.forEach(row => {
        totalPendientesDev += parseInt(row['cp_pendiente_dev']) || 0;
    });
    const porcentajePendientes = totalDiseñados > 0 ? Math.round((totalPendientesDev / totalDiseñados) * 100) : 0;
    
    // % Pendiente Aprobación (Equipo de Negocio): usar cp_pendiente_neg directamente del Excel
    const porcentajePendientesAprobacion = totalDiseñados > 0 ? Math.round((totalPendientes / totalDiseñados) * 100) : 0;
    
    // Actualizar tarjetas KPI (los nombres deben coincidir con los <h3> del HTML)
    updateKPICard('APLICACIONES', aplicacionesUnicas.size);
    updateKPICard('VERTICALES', verticalesUnicas.size);
    updateKPICard('EPICAS', epicasUnicas.size);
    updateKPICard('HISTORIAS DE USUARIO', historiasUnicas.size);
    updateKPICard('CASOS DE PRUEBA DISEÑADOS', totalDiseñados);
    updateKPICard('CASOS DE PRUEBA APROBADOS', totalAprobados);
    
    // Actualizar porcentaje de avance general
    updateProgressPercentage(porcentajeAvance);
    
    // Actualizar porcentajes individuales
    updateIndividualProgress('designProgressPercentage', porcentajeDiseño);
    updateIndividualProgress('reviewProgressPercentage', porcentajeRevisiones);
    updateIndividualProgress('approvalProgressPercentage', porcentajeAprobaciones);
    updateIndividualProgress('pendingProgressPercentage', porcentajePendientes);
    updateIndividualProgress('pendingApprovalProgressPercentage', porcentajePendientesAprobacion);
}

// Actualizar tarjeta KPI individual
function updateKPICard(title, value) {
    // Buscar la tarjeta por su título
    const kpiCards = document.querySelectorAll('.kpi-card');
    
    kpiCards.forEach(card => {
        const titleElement = card.querySelector('h3');
        if (titleElement) {
            // Normalizar el texto eliminando saltos de línea y texto entre paréntesis
            const normalizedTitle = titleElement.textContent.trim().toUpperCase().replace(/\s+/g, ' ');
            const normalizedSearch = title.toUpperCase().trim();
            
            // Comprobar si el título comienza con el texto de búsqueda
            if (normalizedTitle.startsWith(normalizedSearch) || normalizedTitle.includes(normalizedSearch)) {
                const numberElement = card.querySelector('.kpi-value .number');
                if (numberElement) {
                    // Animar el cambio de valor
                    numberElement.style.transition = 'transform 0.3s ease, color 0.3s ease';
                    numberElement.style.transform = 'scale(1.2)';
                    numberElement.textContent = value;
                    
                    setTimeout(() => {
                        numberElement.style.transform = 'scale(1)';
                    }, 300);
                }
            }
        }
    });
}

// Actualizar porcentaje de avance general con velocímetro animado
function updateProgressPercentage(percentage) {
    const percentageElement = document.getElementById('generalProgressPercentage');
    if (percentageElement) {
        // Actualizar texto del porcentaje
        percentageElement.textContent = `${percentage}%`;
        
        // Actualizar velocímetro
        const progressPath = document.getElementById('speedometerProgress');
        const needle = document.getElementById('speedometerNeedle');
        
        if (progressPath && needle) {
            const circumference = 314; // Longitud del arco (π * radio * π)
            const offset = circumference - (percentage / 100) * circumference;
            
            // Animar el arco de progreso
            setTimeout(() => {
                progressPath.style.strokeDashoffset = offset;
                progressPath.style.stroke = getColorByPercentage(percentage);
            }, 50);
            
            // Animar la aguja (de -90° a 90° = 180° total)
            const angle = -90 + (percentage / 100) * 180;
            needle.style.transform = `rotate(${angle}deg)`;
            
            // Actualizar color del porcentaje
            const color = getColorByPercentage(percentage);
            percentageElement.style.color = color;
        }
    }
}

// Calcular color basado en porcentaje (0% = rojo, 100% = verde)
function getColorByPercentage(percentage) {
    // Asegurar que el porcentaje esté entre 0 y 100
    const percent = Math.max(0, Math.min(100, percentage));
    
    // Definir colores
    const red = { r: 220, g: 38, b: 38 };      // #dc2626 (rojo)
    const yellow = { r: 234, g: 179, b: 8 };   // #eab308 (amarillo)
    const green = { r: 34, g: 197, b: 94 };    // #22c55e (verde)
    
    let color;
    
    if (percent <= 50) {
        // Transición de rojo a amarillo (0-50%)
        const ratio = percent / 50;
        color = {
            r: Math.round(red.r + (yellow.r - red.r) * ratio),
            g: Math.round(red.g + (yellow.g - red.g) * ratio),
            b: Math.round(red.b + (yellow.b - red.b) * ratio)
        };
    } else {
        // Transición de amarillo a verde (50-100%)
        const ratio = (percent - 50) / 50;
        color = {
            r: Math.round(yellow.r + (green.r - yellow.r) * ratio),
            g: Math.round(yellow.g + (green.g - yellow.g) * ratio),
            b: Math.round(yellow.b + (green.b - yellow.b) * ratio)
        };
    }
    
    return `rgb(${color.r}, ${color.g}, ${color.b})`;
}

// Actualizar porcentajes individuales con SVG animado
function updateIndividualProgress(elementId, percentage) {
    const progressElement = document.getElementById(elementId);
    if (progressElement) {
        // Actualizar texto del porcentaje
        progressElement.textContent = `${percentage}%`;
        
        // Obtener el círculo SVG
        const circleId = elementId.replace('Percentage', 'Circle');
        const circleContainer = document.getElementById(circleId);
        
        if (circleContainer) {
            const svgCircle = circleContainer.querySelector('svg circle:last-child');
            
            if (svgCircle) {
                // Calcular el dashoffset basado en el porcentaje
                const radius = parseFloat(svgCircle.getAttribute('r'));
                const circumference = 2 * Math.PI * radius;
                
                // Inicializar strokeDasharray si no está configurado
                if (!svgCircle.style.strokeDasharray) {
                    svgCircle.style.strokeDasharray = circumference;
                    svgCircle.style.strokeDashoffset = circumference;
                }
                
                const offset = circumference - (percentage / 100) * circumference;
                
                // Aplicar animación después de un pequeño delay
                setTimeout(() => {
                    svgCircle.style.strokeDashoffset = offset;
                }, 50);
                
                // Aplicar color basado en porcentaje
                let color;
                if (elementId !== 'pendingProgressPercentage' && elementId !== 'pendingApprovalProgressPercentage') {
                    color = getColorByPercentage(percentage);
                } else {
                    // Para pendientes, invertir el color
                    const invertedPercent = 100 - percentage;
                    color = getColorByPercentage(invertedPercent);
                }
                
                svgCircle.style.stroke = color;
                progressElement.style.color = color;
            }
        }
    }
}

// Reabrir selector de proyectos
function openProjectSelector() {
    if (availableSheets && availableSheets.length > 0) {
        showSheetSelector(availableSheets);
    } else {
        showNotification('No hay hojas disponibles', 'warning');
    }
}

// Mostrar notificación (si no existe esta función, créala)
function showNotification(message, type) {
    // Implementación básica de notificación
    console.log(`[${type.toUpperCase()}] ${message}`);
    
    // Puedes implementar una notificación visual aquí
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        background: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : '#f59e0b'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        z-index: 10001;
        animation: slideInRight 0.3s ease;
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}
