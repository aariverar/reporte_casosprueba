// Configuración de colores consistente con CSS
const colors = {
    success: '#10b981',
    danger: '#ef4444',
    pending: '#8b5cf6',
    primary: '#2563eb',
    warning: '#f59e0b'
};

// Variable global para almacenar los datos cargados del Excel
let loadedExcelData = null;

// Datos de ejemplo para las pruebas
const testData = {
    projectInfo: {
        name: 'Building Block - Core',
        qaResponsible: 'María González',
        startDate: '01/09/2024',
        endDate: '30/09/2024',
        status: 'Pruebas QA',
        progress: 75
    },
    summary: {
        planned: 45,    // Pruebas aún por ejecutar
        successful: 85, // Pruebas completadas exitosamente
        failed: 15,     // Pruebas que fallaron
        pending: 3,     // Pruebas en revisión/pendientes
        blocked: 2,     // Pruebas bloqueadas
        dismissed: 0    // Pruebas desestimadas (se calcula desde Excel)
        // Total inicial sin dismissed fijo
    },
    trend: [
        { date: '01/09/2024', planned: 150, successful: 0, failed: 0, pending: 0, blocked: 0, dismissed: 0 },   // Día 1: Todo planificado
        { date: '02/09/2024', planned: 130, successful: 15, failed: 3, pending: 2, blocked: 0, dismissed: 0 },  // Día 2: Empezando
        { date: '03/09/2024', planned: 108, successful: 28, failed: 6, pending: 6, blocked: 2, dismissed: 0 },  // Día 3: Progresando, 2 bloqueadas
        { date: '04/09/2024', planned: 88, successful: 42, failed: 8, pending: 10, blocked: 2, dismissed: 0 },  // Día 4: Avanzando
        { date: '05/09/2024', planned: 72, successful: 55, failed: 12, pending: 8, blocked: 3, dismissed: 0 },  // Día 5: Más progreso, 1 nueva bloqueada
        { date: '06/09/2024', planned: 58, successful: 68, failed: 14, pending: 8, blocked: 2, dismissed: 0 },  // Día 6: Continuando, desbloqueada 1
        { date: '07/09/2024', planned: 50, successful: 75, failed: 15, pending: 8, blocked: 2 },  // Día 7: Más avance
        { date: '08/09/2024', planned: 45, successful: 80, failed: 15, pending: 8, blocked: 2 },  // Día 8: Estado actual
        { date: '09/09/2024', planned: 45, successful: 85, failed: 15, pending: 3, blocked: 2 },  // Día 9
        { date: '10/09/2024', planned: 40, successful: 90, failed: 16, pending: 2, blocked: 2 },  // Día 10
        { date: '11/09/2024', planned: 35, successful: 95, failed: 16, pending: 2, blocked: 2 },  // Día 11
        { date: '12/09/2024', planned: 30, successful: 100, failed: 17, pending: 1, blocked: 2 }, // Día 12
        { date: '13/09/2024', planned: 25, successful: 105, failed: 17, pending: 1, blocked: 2 }, // Día 13
        { date: '14/09/2024', planned: 20, successful: 110, failed: 18, pending: 0, blocked: 2 }, // Día 14
        { date: '15/09/2024', planned: 15, successful: 115, failed: 18, pending: 0, blocked: 2 }, // Día 15
        { date: '01/10/2024', planned: 12, successful: 118, failed: 18, pending: 0, blocked: 2 }, // Octubre día 1
        { date: '02/10/2024', planned: 10, successful: 120, failed: 18, pending: 0, blocked: 2 }, // Octubre día 2
        { date: '03/10/2024', planned: 8, successful: 122, failed: 18, pending: 0, blocked: 2 },  // Octubre día 3
        { date: '04/10/2024', planned: 6, successful: 124, failed: 18, pending: 0, blocked: 2 },  // Octubre día 4
        { date: '05/10/2024', planned: 4, successful: 126, failed: 18, pending: 0, blocked: 2 },  // Octubre día 5
        { date: '06/10/2024', planned: 2, successful: 128, failed: 18, pending: 0, blocked: 2 },  // Octubre día 6
        { date: '07/10/2024', planned: 1, successful: 129, failed: 18, pending: 0, blocked: 2 },  // Octubre día 7
        { date: '08/10/2024', planned: 0, successful: 130, failed: 18, pending: 0, blocked: 2 },  // Octubre día 8
        { date: '09/10/2024', planned: 0, successful: 132, failed: 18, pending: 0, blocked: 0 },  // Octubre día 9
        { date: '10/10/2024', planned: 0, successful: 134, failed: 16, pending: 0, blocked: 0 },  // Octubre día 10
        { date: '11/10/2024', planned: 0, successful: 136, failed: 14, pending: 0, blocked: 0 },  // Octubre día 11
        { date: '12/10/2024', planned: 0, successful: 138, failed: 12, pending: 0, blocked: 0 },  // Octubre día 12
        { date: '13/10/2024', planned: 0, successful: 140, failed: 10, pending: 0, blocked: 0 },  // Octubre día 13
        { date: '14/10/2024', planned: 0, successful: 142, failed: 8, pending: 0, blocked: 0 },   // Octubre día 14 (hoy)
    ],
    categories: [
        { date: '2024-09-28', planned: 128, successful: 102, failed: 11, pending: 5 },
        { date: '2024-09-29', planned: 130, successful: 105, failed: 8, pending: 7 },
        { date: '2024-09-30', planned: 132, successful: 108, failed: 12, pending: 4 }
    ],
    categories: [
        { escenario: 'API Tests', planned: 30, successful: 25, failed: 3, pending: 1 },
        { escenario: 'UI Tests', planned: 28, successful: 20, failed: 5, pending: 2 },
        { escenario: 'Integration', planned: 25, successful: 18, failed: 4, pending: 1 },
        { escenario: 'Performance', planned: 18, successful: 12, failed: 2, pending: 1 },
        { escenario: 'Security', planned: 12, successful: 10, failed: 1, pending: 0 }
    ],
    defects: {
        summary: { critical: 2, high: 5, medium: 8, low: 3 },
        details: [
            { id: 'DEF001', title: 'Sistema no responde después de login', severity: 'critical', status: 'Open', escenario: 'API Tests', assignee: 'Ana García', dateFound: '18/09/2024' },
            { id: 'DEF002', title: 'Error de validación en formulario de registro', severity: 'critical', status: 'Assigned', escenario: 'UI Tests', assignee: 'Carlos López', dateFound: '20/09/2024' },
            { id: 'DEF003', title: 'Timeout en procesamiento de pagos', severity: 'high', status: 'In Progress DEV', escenario: 'Integration', assignee: 'María Rodríguez', dateFound: '19/09/2024' },
            { id: 'DEF004', title: 'Interfaz no responsive en móviles', severity: 'high', status: 'Under Review', escenario: 'UI Tests', assignee: 'Pedro Ruiz', dateFound: '21/09/2024' },
            { id: 'DEF005', title: 'Performance lenta en búsquedas', severity: 'high', status: 'Resolved', escenario: 'Performance', assignee: 'Laura Sánchez', dateFound: '17/09/2024' },
            { id: 'DEF006', title: 'Mensajes de error poco claros', severity: 'medium', status: 'ReTesting', escenario: 'UI Tests', assignee: 'José Martín', dateFound: '22/09/2024' },
            { id: 'DEF007', title: 'Falta validación de campos', severity: 'medium', status: 'In Progress DEV', escenario: 'API Tests', assignee: 'Ana García', dateFound: '16/09/2024' },
            { id: 'DEF008', title: 'Inconsistencia en colores del tema', severity: 'low', status: 'Closed', escenario: 'UI Tests', assignee: 'Carlos López', dateFound: '23/09/2024' },
            { id: 'DEF009', title: 'Error en validación de email', severity: 'medium', status: 'Reject', escenario: 'API Tests', assignee: 'Ana García', dateFound: '24/09/2024' },
            { id: 'DEF010', title: 'Problema con scroll en móviles', severity: 'low', status: 'ReOpened', escenario: 'UI Tests', assignee: 'Pedro Ruiz', dateFound: '25/09/2024' }
        ]
    },
    testDetails: [
        { id: 'T001', name: 'Login API Validation', escenario: 'API Tests', status: 'success', dia_ejecutado: '08/10/2024', executor: 'Ana García' },
        { id: 'T002', name: 'User Registration Flow', escenario: 'UI Tests', status: 'success', dia_ejecutado: '08/10/2024', executor: 'Carlos López' },
        { id: 'T003', name: 'Payment Processing', escenario: 'Integration', status: 'failure', dia_ejecutado: '07/10/2024', executor: 'María Rodríguez' },
        { id: 'T004', name: 'Database Connection', escenario: 'Integration', status: 'success', dia_ejecutado: '08/10/2024', executor: 'José Martín' },
        { id: 'T005', name: 'Load Testing', escenario: 'Performance', status: 'pending', dia_ejecutado: '', executor: 'Laura Sánchez' },
        { id: 'T006', name: 'Search Functionality', escenario: 'UI Tests', status: 'failure', dia_ejecutado: '06/10/2024', executor: 'Pedro Ruiz' },
        { id: 'T007', name: 'Email Notifications', escenario: 'API Tests', status: 'success', dia_ejecutado: '08/10/2024', executor: 'Ana García' },
        { id: 'T008', name: 'Security Headers', escenario: 'Security', status: 'success', dia_ejecutado: '09/10/2024', executor: 'Carlos López' },
        { id: 'T009', name: 'Password Reset Flow', escenario: 'UI Tests', status: 'pending', dia_ejecutado: '', executor: 'Ana García' },
        { id: 'T010', name: 'API Rate Limiting', escenario: 'API Tests', status: 'blocked', dia_ejecutado: '', executor: 'José Martín' },
        { id: 'T011', name: 'Mobile Responsive Design', escenario: 'UI Tests', status: 'failure', dia_ejecutado: '05/10/2024', executor: 'Pedro Ruiz' },
        { id: 'T012', name: 'Database Backup Process', escenario: 'Integration', status: 'success', dia_ejecutado: '07/10/2024', executor: 'María Rodríguez' },
        { id: 'T013', name: 'Stress Testing CPU', escenario: 'Performance', status: 'success', dia_ejecutado: '04/10/2024', executor: 'Laura Sánchez' },
        { id: 'T014', name: 'HTTPS Certificate Validation', escenario: 'Security', status: 'success', dia_ejecutado: '09/10/2024', executor: 'Carlos López' },
        { id: 'T015', name: 'Shopping Cart Functionality', escenario: 'UI Tests', status: 'pending', dia_ejecutado: '', executor: 'Pedro Ruiz' },
        { id: 'T016', name: 'OAuth2 Authentication', escenario: 'API Tests', status: 'failure', dia_ejecutado: '03/10/2024', executor: 'Ana García' },
        { id: 'T017', name: 'Cross-browser Compatibility', escenario: 'UI Tests', status: 'planned', dia_ejecutado: '', executor: 'Ana García' },
        { id: 'T018', name: 'Memory Leak Testing', escenario: 'Performance', status: 'planned', dia_ejecutado: '', executor: 'Laura Sánchez' },
        { id: 'T019', name: 'SQL Injection Prevention', escenario: 'Security', status: 'blocked', dia_ejecutado: '', executor: 'Carlos López' },
        { id: 'T020', name: 'File Upload Validation', escenario: 'API Tests', status: 'planned', dia_ejecutado: '', executor: 'José Martín' },
        { id: 'T021', name: 'Legacy Browser Support', escenario: 'UI Tests', status: 'dismissed', dia_ejecutado: '', executor: 'Ana García' },
        { id: 'T022', name: 'Obsolete API Endpoints', escenario: 'API Tests', status: 'dismissed', dia_ejecutado: '', executor: 'Carlos López' },
        { id: 'T023', name: 'Deprecated Features Test', escenario: 'Integration', status: 'dismissed', dia_ejecutado: '', executor: 'María Rodríguez' },
        { id: 'T024', name: 'Old Version Compatibility', escenario: 'Performance', status: 'dismissed', dia_ejecutado: '', executor: 'Laura Sánchez' },
        { id: 'T025', name: 'Unused Security Protocols', escenario: 'Security', status: 'dismissed', dia_ejecutado: '', executor: 'José Martín' }
    ]
};

// Función para forzar actualización anti-caché
function forcePageRefresh() {
    // Limpiar todo el localStorage
    localStorage.clear();
    sessionStorage.clear();
    
    // Detectar si la página viene del caché del navegador
    if (performance.navigation.type === performance.navigation.TYPE_BACK_FORWARD) {
        console.log('🔄 Página cargada desde navegación hacia atrás/adelante - Forzando recarga...');
        location.reload(true);
        return;
    }
    
    // Detectar si hay datos previos en memoria
    if (typeof testData !== 'undefined' && testData.projectInfo && testData.projectInfo.name !== 'Proyecto de Ejemplo') {
        console.log('🔄 Datos previos detectados - Limpiando...');
        resetToDefaultState();
    }
    
    // Agregar timestamp único para evitar caché de GitHub Pages
    const currentTimestamp = new Date().getTime();
    if (!window.location.search.includes('_t=')) {
        const separator = window.location.search ? '&' : '?';
        const newUrl = window.location.href + separator + '_t=' + currentTimestamp;
        console.log('🔄 Agregando timestamp anti-caché:', newUrl);
        // Solo cambiar URL sin recargar si no hay parámetros de timestamp
        window.history.replaceState({}, '', newUrl);
    }
    
    console.log('✅ Limpieza anti-caché completada');
}

// Función para resetear a estado por defecto
function resetToDefaultState() {
    // Resetear testData a valores por defecto
    if (typeof testData !== 'undefined') {
        testData.projectInfo = {
            name: 'Proyecto de Ejemplo',
            qaResponsible: 'No asignado',
            startDate: '01/09/2024',
            endDate: '30/09/2024',
            status: 'En progreso',
            progress: 0
        };
        
        testData.summary = {
            planned: 0,
            successful: 0,
            failed: 0,
            pending: 0,
            blocked: 0,
            dismissed: 0
        };
    }
    
    // Limpiar todos los gráficos
    clearAllCharts();
    
    console.log('✅ Estado reseteado a valores por defecto');
}

//  Inicializar la aplicación
document.addEventListener('DOMContentLoaded', function() {
    // Service Worker deshabilitado temporalmente para evitar problemas de caché
    // if ('serviceWorker' in navigator) {
    //     navigator.serviceWorker.register('./sw.js?v=20251103')
    //         .then(function(registration) {
    //             console.log('✅ Service Worker registrado:', registration);
    //         })
    //         .catch(function(error) {
    //             console.log('❌ Error registrando Service Worker:', error);
    //         });
    // }
    
    // Fuerza limpieza agresiva del caché al cargar la página
    forcePageRefresh();
    
    initializeDashboard();
    createCharts();
    populateTestTable();
    populateDefectsTable();
    setupEventListeners();
    setupFileUpload();
    setupEditorIntegration();
});

// Calcular progreso del proyecto basado en tendencia histórica
function calculateProjectProgress() {
    if (!testData.trend || testData.trend.length === 0) {
        console.warn('⚠️ No hay datos de tendencia histórica disponibles');
        return 0.0;
    }
    
    // Función para parsear fechas en formato dd/MM/yyyy
    function parseDateDDMMYYYY(dateStr) {
        if (!dateStr || typeof dateStr !== 'string') return null;
        
        const parts = dateStr.split('/');
        if (parts.length !== 3) return null;
        
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // JavaScript usa meses 0-11
        const year = parseInt(parts[2], 10);
        
        if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
        
        const date = new Date(year, month, day);
        return isNaN(date.getTime()) ? null : date;
    }
    
    // Fecha objetivo: fecha actual del sistema
    const fechaObjetivo = new Date();
    console.log(`🗓️ CÁLCULO % REAL - Buscar % Real para fecha: ${fechaObjetivo.toISOString().split('T')[0]} (${fechaObjetivo.toLocaleDateString('es-ES', { weekday: 'long' })})`);
    
    // Ordenar tendencia por fecha
    const trendSorted = [...testData.trend].sort((a, b) => {
        const dateA = parseDateDDMMYYYY(a.date);
        const dateB = parseDateDDMMYYYY(b.date);
        return dateA && dateB ? dateA - dateB : 0;
    });
    
    // Buscar % Real directamente de la columna "% Real" de Tendencia_Historica
    let realProgress = 0.0;
    let ultimaFechaIncluida = null;
    
    // Buscar el último día que sea <= fecha actual y que tenga % Real
    trendSorted.forEach((dayData, index) => {
        const dayDate = parseDateDDMMYYYY(dayData.date);
        if (!dayDate) {
            console.warn(`⚠️ Saltando fila ${index + 1}: No se pudo parsear fecha "${dayData.date}"`);
            return;
        }
        
        const isIncluded = dayDate <= fechaObjetivo;
        const porcentajeReal = parseFloat(dayData.realProgress || dayData.real_progress || 0);
        const fechaFormateada = dayDate.toISOString().split('T')[0];
        const diaSemana = dayDate.toLocaleDateString('es-ES', { weekday: 'long' });
        
        console.log(`  Fila ${index + 1}: ${dayData.date} (${diaSemana}) -> % Real: ${porcentajeReal}%`);
        console.log(`    Comparación: ${fechaFormateada} <= ${fechaObjetivo.toISOString().split('T')[0]} = ${isIncluded} ${isIncluded ? '✅ INCLUIR' : '❌ EXCLUIR'}`);
        
        if (isIncluded && porcentajeReal > 0) {
            realProgress = porcentajeReal;
            ultimaFechaIncluida = dayData.date;
        }
    });
    
    console.log(`📊 % REAL DESDE COLUMNA TENDENCIA_HISTORICA:`);
    console.log(`   • FECHA OBJETIVO: ${fechaObjetivo.toISOString().split('T')[0]} (${fechaObjetivo.toLocaleDateString('es-ES', { weekday: 'long' })})`);
    console.log(`   • ÚLTIMA FECHA INCLUIDA: ${ultimaFechaIncluida}`);
    console.log(`   • % Real encontrado: ${realProgress}%`);
    console.log(`💡 LECTURA DIRECTA DESDE COLUMNA % REAL DEL EXCEL`);
    
    return realProgress;
}

// Función para calcular el progreso planificado basado en tendencia histórica
function calculatePlannedProgress() {
    if (!testData.trend || testData.trend.length === 0) {
        console.warn('⚠️ No hay datos de tendencia histórica para progreso planificado');
        return 0;
    }
    
    // Obtener la fecha actual para comparar con los días de la tendencia
    const today = new Date();
    console.log(`🗓️ CÁLCULO % PLANIFICADO - Fórmula Excel: SUMA($C$2:C_actual)/SUMA($C$2:$C_último)*100`);
    console.log(`🗓️ FECHA ACTUAL SISTEMA: ${today.toISOString().split('T')[0]} (${today.toLocaleDateString()})`);
    console.log(`🗓️ FECHA ACTUAL EN MILISEGUNDOS: ${today.getTime()}`);
    console.log(`🗓️ DÍA DE LA SEMANA: ${today.toLocaleDateString('es-ES', { weekday: 'long' })}`);
    
    // Ordenar tendencia por fecha para procesamiento secuencial
    const trendSorted = [...testData.trend].sort((a, b) => new Date(a.date) - new Date(b.date));
    
    console.log(`📋 DATOS DE TENDENCIA ENCONTRADOS (${trendSorted.length} registros):`);
    trendSorted.forEach((day, index) => {
        console.log(`  ${index + 1}: Fecha raw: "${day.date}" (tipo: ${typeof day.date}) | Planificadas: ${day.planned || 0}`);
        const dayDate = parseDateDDMMYYYY(day.date);
        if (!dayDate) {
            console.log(`    ❌ NO SE PUDO PARSEAR (formato esperado: dd/MM/yyyy)`);
        } else {
            console.log(`    ✅ Fecha parseada: ${dayDate.toLocaleDateString('es-ES', { weekday: 'long' })} -> ${dayDate.toISOString().split('T')[0]}`);
        }
    });
    
    // PASO 1: Calcular SUMA($C$2:$C_último) - Total planificadas de todo el proyecto
    let totalPlanificadasProyecto = 0;
    trendSorted.forEach(dayData => {
        totalPlanificadasProyecto += dayData.planned || 0;
    });
    
    // PASO 2: Calcular SUMA($C$2:C_actual) - Acumulado hasta fecha actual o más cercana
    let sumaPlanificadasAcumulada = 0;
    let ultimaFechaIncluida = null;
    let fechaObjetivo = new Date(); // CORRECCIÓN: Usar fecha actual del sistema
    
    console.log(`🎯 BUSCANDO DATOS HASTA: ${fechaObjetivo.toISOString().split('T')[0]} (${fechaObjetivo.toLocaleDateString('es-ES', { weekday: 'long' })})`);
    
    // Función para parsear fechas en formato dd/MM/yyyy
    function parseDateDDMMYYYY(dateStr) {
        if (!dateStr || typeof dateStr !== 'string') return null;
        
        const parts = dateStr.split('/');
        if (parts.length !== 3) return null;
        
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // JavaScript usa meses 0-11
        const year = parseInt(parts[2], 10);
        
        if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
        
        const date = new Date(year, month, day);
        return isNaN(date.getTime()) ? null : date;
    }
    
    // Buscar la fecha exacta o la más cercana anterior a la fecha actual
    let fechaEncontrada = null;
    const fechaObjetivoStr = fechaObjetivo.toISOString().split('T')[0];
    trendSorted.forEach((dayData) => {
        const dayDate = parseDateDDMMYYYY(dayData.date);
        if (dayDate && dayDate.toISOString().split('T')[0] === fechaObjetivoStr) {
            fechaEncontrada = dayData.date;
            console.log(`🎉 ENCONTRADA FECHA EXACTA: ${dayData.date}`);
        }
    });
    
    if (!fechaEncontrada) {
        console.log(`⚠️ No se encontró datos para ${fechaObjetivoStr}, buscando fecha más cercana anterior...`);
        // Buscar la fecha más cercana anterior a la fecha actual
        let fechasCercanas = trendSorted.filter(dayData => {
            const dayDate = parseDateDDMMYYYY(dayData.date);
            return dayDate && dayDate <= fechaObjetivo;
        });
        if (fechasCercanas.length > 0) {
            fechaEncontrada = fechasCercanas[fechasCercanas.length - 1].date;
            console.log(`📅 USANDO FECHA MÁS CERCANA: ${fechaEncontrada}`);
        }
    }
    
    trendSorted.forEach((dayData, index) => {
        const dayDate = parseDateDDMMYYYY(dayData.date);
        const planificadasDelDia = dayData.planned || 0;
        
        // Validar que la fecha sea válida
        if (!dayDate) {
            console.warn(`⚠️ Saltando fila ${index + 1}: No se pudo parsear fecha "${dayData.date}"`);
            return;
        }
        
        const isIncluded = dayDate <= fechaObjetivo;
        const fechaFormateada = dayDate.toISOString().split('T')[0];
        const diaSemana = dayDate.toLocaleDateString('es-ES', { weekday: 'long' });
        
        console.log(`  Fila ${index + 1}: ${dayData.date} (${diaSemana}) -> Planificadas: ${planificadasDelDia}`);
        console.log(`    Comparación: ${fechaFormateada} <= ${fechaObjetivoStr} = ${isIncluded} ${isIncluded ? '✅ INCLUIR' : '❌ EXCLUIR'}`);
        
        if (isIncluded) {
            sumaPlanificadasAcumulada += planificadasDelDia;
            ultimaFechaIncluida = dayData.date;
        }
    });
    
    if (totalPlanificadasProyecto === 0) {
        console.warn('⚠️ Total de planificadas del proyecto es 0');
        return 0;
    }
    
    // PASO 3: Aplicar fórmula Excel: SUMA($C$2:C_actual)/SUMA($C$2:$C_último)*100
    const plannedProgress = parseFloat(((sumaPlanificadasAcumulada / totalPlanificadasProyecto) * 100).toFixed(1));
    
    console.log(`� PROGRESO PLANIFICADO (TENDENCIA_HISTORICA):`);
    console.log(`   • FECHA OBJETIVO: ${fechaObjetivoStr} (${fechaObjetivo.toLocaleDateString('es-ES', { weekday: 'long' })})`);
    console.log(`   • ÚLTIMA FECHA INCLUIDA: ${ultimaFechaIncluida}`);
    console.log(`   • SUMA($C$2:C_actual) hasta ${ultimaFechaIncluida}: ${sumaPlanificadasAcumulada}`);
    console.log(`   • SUMA($C$2:$C_último) total proyecto: ${totalPlanificadasProyecto}`);
    console.log(`   • % Planificado: ${plannedProgress}% = (${sumaPlanificadasAcumulada}/${totalPlanificadasProyecto}) * 100`);
    console.log(`   • Filas procesadas: ${trendSorted.length}`);
    console.log(`   • RESULTADO ACTUAL: ${plannedProgress}%`);
    console.log(`💡 CÁLCULO DINÁMICO BASADO EN FECHA ACTUAL DEL SISTEMA`);
    
    return plannedProgress;
}

// Inicializar dashboard
function initializeDashboard() {
    // Mostrar fecha actual
    const currentDate = new Date().toLocaleDateString('es-ES', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    const currentDateElement = document.getElementById('currentDate');
    if (currentDateElement) {
        currentDateElement.textContent = currentDate;
    }

    // Actualizar información del proyecto solo si los elementos existen
    const projectName = document.getElementById('projectName');
    if (projectName) {
        projectName.textContent = `Proyecto: ${testData.projectInfo.name}`;
    }
    
    const projectNameDetail = document.getElementById('projectNameDetail');
    if (projectNameDetail) {
        projectNameDetail.textContent = testData.projectInfo.name;
    }
    
    const qaResponsible = document.getElementById('qaResponsible');
    if (qaResponsible) {
        qaResponsible.textContent = testData.projectInfo.qaResponsible;
    }
    
    const projectStartDate = document.getElementById('projectStartDate');
    if (projectStartDate) {
        projectStartDate.textContent = testData.projectInfo.startDate;
    }
    
    const projectEndDate = document.getElementById('projectEndDate');
    if (projectEndDate) {
        projectEndDate.textContent = testData.projectInfo.endDate;
    }
    
    const projectStatusDetail = document.getElementById('projectStatusDetail');
    if (projectStatusDetail) {
        projectStatusDetail.textContent = testData.projectInfo.status;
    }
    
    const reportStartDate = document.getElementById('reportStartDate');
    if (reportStartDate) {
        reportStartDate.textContent = testData.projectInfo.startDate;
    }
    
    const reportEndDate = document.getElementById('reportEndDate');
    if (reportEndDate) {
        reportEndDate.textContent = testData.projectInfo.endDate;
    }
    
    // Actualizar progreso del proyecto con barras duales
    const realProgress = calculateProjectProgress();
    const plannedProgress = calculatePlannedProgress();
    
    console.log(`🔄 Actualizando barras: Planificado=${plannedProgress}%, Actual=${realProgress}%`);
    
    // Actualizar progreso real
    const realProgressElement = document.getElementById('realProgress');
    const realProgressText = document.getElementById('realProgressText');
    
    if (realProgressElement && realProgressText) {
        realProgressElement.style.width = realProgress + '%';
        realProgressText.textContent = realProgress.toFixed(1) + '%';
        console.log(`✅ Progreso real actualizado: ${realProgress}%`);
    } else {
        console.warn('⚠️ Elementos de progreso real no encontrados');
    }
    
    // Actualizar progreso planificado
    const plannedProgressElement = document.getElementById('plannedProgress');
    const plannedProgressText = document.getElementById('plannedProgressText');
    
    if (plannedProgressElement && plannedProgressText) {
        plannedProgressElement.style.width = plannedProgress + '%';
        plannedProgressText.textContent = plannedProgress.toFixed(1) + '%';
        console.log(`✅ Progreso planificado actualizado: ${plannedProgress}%`);
    } else {
        console.warn('⚠️ Elementos de progreso planificado no encontrados');
    }
    
    // Actualizar el progreso en testData para consistencia
    testData.projectInfo.progress = realProgress;
    testData.projectInfo.plannedProgress = plannedProgress;
    
    // Actualizar estado del reporte
    const reportStatus = document.getElementById('reportStatus');
    if (reportStatus) {
        reportStatus.textContent = testData.projectInfo.status;
        
        // Aplicar estilos según el estado
        switch(testData.projectInfo.status) {
            case 'Finalizado':
                reportStatus.style.background = 'rgba(16, 185, 129, 0.2)';
                reportStatus.style.color = '#10b981';
                break;
            case 'Pruebas QA':
                reportStatus.style.background = 'rgba(59, 130, 246, 0.2)';
                reportStatus.style.color = '#3b82f6';
                break;
            case 'En Progreso':
            default:
                reportStatus.style.background = 'rgba(245, 158, 11, 0.2)';
                reportStatus.style.color = '#f59e0b';
                break;
        }
    }

    // Calcular el total de pruebas (todos los estados)
    const blockedCount = testData.summary.blocked || 0;
    const dismissedCount = testData.summary.dismissed || 0;
    const totalTests = testData.summary.planned + testData.summary.successful + 
                      testData.summary.failed + testData.summary.pending + blockedCount + dismissedCount;
    
    // Actualizar KPIs con la nueva lógica solo si los elementos existen
    const plannedTests = document.getElementById('plannedTests');
    if (plannedTests) {
        plannedTests.textContent = testData.summary.planned;
    }
    
    const successfulTests = document.getElementById('successfulTests');
    if (successfulTests) {
        successfulTests.textContent = testData.summary.successful;
    }
    
    const failedTests = document.getElementById('failedTests');
    if (failedTests) {
        failedTests.textContent = testData.summary.failed;
    }
    
    const pendingTests = document.getElementById('pendingTests');
    if (pendingTests) {
        pendingTests.textContent = testData.summary.pending;
        console.log(`🔍 DEBUGGING DOM - VALOR FINAL MOSTRADO EN PENDIENTES: ${testData.summary.pending}`);
    }
    
    // Actualizar pruebas bloqueadas
    const blockedTests = document.getElementById('blockedTests');
    if (blockedTests) {
        blockedTests.textContent = blockedCount;
    }
    
    // Actualizar pruebas desestimadas
    const dismissedTests = document.getElementById('dismissedTests');
    if (dismissedTests) {
        dismissedTests.textContent = dismissedCount;
    }
    
    console.log('🎯 ACTUALIZANDO ELEMENTOS DOM:');
    console.log(`   • Planificadas: ${testData.summary.planned}`);
    console.log(`   • Exitosas: ${testData.summary.successful}`);
    console.log(`   • Fallidas: ${testData.summary.failed}`);
    console.log(`   • Pendientes: ${testData.summary.pending}`);
    console.log(`   • Bloqueadas: ${blockedCount}`);
    console.log(`   • DESESTIMADAS: ${dismissedCount}`);
    console.log(`   • Total: ${totalTests}`);
    
    // Actualizar porcentajes basados en la nueva lógica de progreso
    updateKPIProgressPercentages(totalTests);
}

// Función para calcular y mostrar porcentajes de progreso basados en el total inicial
function updateKPIProgressPercentages(totalTests) {
    if (totalTests === 0) {
        hideAllPercentages();
        return;
    }
    
    // Calcular porcentajes del total para cada tipo
    const plannedPercentage = (testData.summary.planned / totalTests) * 100;
    const successfulPercentage = (testData.summary.successful / totalTests) * 100;
    const failedPercentage = (testData.summary.failed / totalTests) * 100;
    const pendingPercentage = (testData.summary.pending / totalTests) * 100;
    const blockedPercentage = ((testData.summary.blocked || 0) / totalTests) * 100;
    const dismissedPercentage = ((testData.summary.dismissed || 0) / totalTests) * 100;
    
    // Actualizar elementos en el DOM con la nueva lógica
    updateProgressPercentageElement('plannedTests', plannedPercentage, 'planned');
    updateProgressPercentageElement('successfulTests', successfulPercentage, 'successful');
    updateProgressPercentageElement('failedTests', failedPercentage, 'failed');
    updateProgressPercentageElement('pendingTests', pendingPercentage, 'pending');
    updateProgressPercentageElement('blockedTests', blockedPercentage, 'blocked');
    updateProgressPercentageElement('dismissedTests', dismissedPercentage, 'dismissed');
}

// Función para actualizar un elemento de porcentaje con lógica de progreso
function updateProgressPercentageElement(testType, percentage, category) {
    // Buscar el elemento de porcentaje correspondiente
    const numberElement = document.getElementById(testType);
    if (!numberElement) return;
    
    const kpiValue = numberElement.parentElement;
    let percentageElement = kpiValue.querySelector('.percentage');
    
    if (!percentageElement) {
        // Si no existe el elemento de porcentaje, crearlo
        percentageElement = document.createElement('span');
        percentageElement.className = 'percentage';
        kpiValue.appendChild(percentageElement);
    }
    
    // Formatear el porcentaje de progreso
    const formattedProgress = formatProgressPercentage(percentage, category);
    percentageElement.textContent = formattedProgress.text;
    
    // Actualizar clases para el color y tooltip
    percentageElement.className = 'percentage ' + formattedProgress.class;
    percentageElement.title = formattedProgress.tooltip;
    
    // Mostrar el elemento si estaba oculto
    percentageElement.style.display = '';
}

// Función para formatear porcentajes de progreso
function formatProgressPercentage(percentage, category) {
    const roundedPercentage = Math.round(percentage * 10) / 10; // Redondear a 1 decimal
    
    let text, cssClass, tooltip;
    
    switch(category) {
        case 'planned':
            text = `${roundedPercentage}%`;
            cssClass = percentage > 70 ? 'neutral' : percentage > 30 ? 'warning' : 'positive';
            tooltip = `${roundedPercentage}% del total aún por ejecutar`;
            break;
            
        case 'successful':
            text = `${roundedPercentage}%`;
            cssClass = percentage > 70 ? 'positive' : percentage > 30 ? 'neutral' : 'negative';
            tooltip = `${roundedPercentage}% del total completado exitosamente`;
            break;
            
        case 'failed':
            text = `${roundedPercentage}%`;
            cssClass = percentage > 15 ? 'negative' : percentage > 5 ? 'warning' : 'positive';
            tooltip = `${roundedPercentage}% del total con fallas`;
            break;
            
        case 'pending':
            text = `${roundedPercentage}%`;
            cssClass = percentage > 20 ? 'negative' : percentage > 10 ? 'warning' : 'neutral';
            tooltip = `${roundedPercentage}% del total en estado pendiente`;
            break;
            
        default:
            text = `${roundedPercentage}%`;
            cssClass = 'neutral';
            tooltip = `${roundedPercentage}% del total`;
    }
    
    return { text, class: cssClass, tooltip };
}

// Función para calcular y actualizar los porcentajes de cambio (función original mantenida para compatibilidad)
function updateKPIPercentages() {
    if (testData.trend && testData.trend.length >= 2) {
        // Obtener el día actual y el día anterior
        const currentDay = testData.trend[testData.trend.length - 1];
        const previousDay = testData.trend[testData.trend.length - 2];
        
        // Calcular cambios porcentuales
        const plannedChange = calculatePercentageChange(previousDay.planned, currentDay.planned);
        const successfulChange = calculatePercentageChange(previousDay.successful, currentDay.successful);
        const failedChange = calculatePercentageChange(previousDay.failed, currentDay.failed);
        const pendingChange = calculatePercentageChange(previousDay.pending, currentDay.pending);
        
        // Actualizar elementos en el DOM
        updatePercentageElement('plannedTests', plannedChange);
        updatePercentageElement('successfulTests', successfulChange);
        updatePercentageElement('failedTests', failedChange);
        updatePercentageElement('pendingTests', pendingChange);
    } else {
        // Si no hay datos suficientes, mostrar "N/A" o ocultar los porcentajes
        hideAllPercentages();
    }
}

// Función para calcular el cambio porcentual
function calculatePercentageChange(oldValue, newValue) {
    if (oldValue === 0) {
        return newValue > 0 ? 100 : 0; // Si el valor anterior era 0 y ahora hay algo, es 100% de incremento
    }
    return ((newValue - oldValue) / oldValue) * 100;
}

// Función para actualizar un elemento de porcentaje específico
function updatePercentageElement(testType, change) {
    // Buscar el elemento de porcentaje correspondiente
    const numberElement = document.getElementById(testType);
    if (!numberElement) return;
    
    const kpiValue = numberElement.parentElement;
    let percentageElement = kpiValue.querySelector('.percentage');
    
    if (!percentageElement) {
        // Si no existe el elemento de porcentaje, crearlo
        percentageElement = document.createElement('span');
        percentageElement.className = 'percentage';
        percentageElement.title = 'Cambio respecto al día anterior';
        kpiValue.appendChild(percentageElement);
    }
    
    // Formatear el porcentaje
    const formattedChange = formatPercentageChange(change, testType);
    percentageElement.textContent = formattedChange.text;
    
    // Actualizar clases para el color
    percentageElement.className = 'percentage ' + formattedChange.class;
    
    // Asegurar que el tooltip esté presente
    if (!percentageElement.title) {
        percentageElement.title = 'Cambio respecto al día anterior';
    }
    
    // Mostrar el elemento si estaba oculto
    percentageElement.style.display = '';
}

// Función para formatear el cambio porcentual
function formatPercentageChange(change, testType = '') {
    const absChange = Math.abs(change);
    const roundedChange = Math.round(absChange * 10) / 10; // Redondear a 1 decimal
    
    let text, cssClass;
    
    if (change > 0) {
        text = `+${roundedChange}%`;
        // Para pruebas fallidas, un aumento es malo (rojo)
        cssClass = (testType === 'failedTests') ? 'negative' : 'positive';
    } else if (change < 0) {
        text = `-${roundedChange}%`;
        // Para pruebas fallidas, una disminución es buena (verde)
        cssClass = (testType === 'failedTests') ? 'positive' : 'negative';
    } else {
        text = '0%';
        cssClass = 'neutral';
    }
    
    return { text, class: cssClass };
}

// Función para ocultar todos los porcentajes cuando no hay datos suficientes
function hideAllPercentages() {
    const percentageElements = document.querySelectorAll('.percentage');
    percentageElements.forEach(element => {
        element.textContent = '';
        element.style.display = 'none';
    });
}

// Crear todas las gráficas
function createCharts() {
    // Solo crear gráficos si los elementos existen en el DOM
    if (document.getElementById('pieChart')) {
        createPieChart();
    }
    if (document.getElementById('trendChart')) {
        createTrendChart();
    }
    if (document.getElementById('categoryChart')) {
        createCategoryChart();
    }
    if (document.getElementById('defectsChart')) {
        createDefectsChart();
    }
    if (document.getElementById('defectCycleTimeChart')) {
        updateDefectCycleTimeChart(testData.defects);
    }
    if (document.getElementById('coverageChart')) {
        updateCoverageChart(testData.testDetails);
    }
    if (document.getElementById('dailyExecutionChart')) {
        updateDailyExecutionChart(testData.testDetails);
    }
    if (document.getElementById('executorDistributionChart')) {
        updateExecutorDistributionChart(testData.testDetails);
    }
    if (document.getElementById('deliverablesChart')) {
        updateDeliverablesChart();
    }
}

// Gráfica de pastel - Resumen de resultados
function createPieChart() {
    const values = [testData.summary.successful, testData.summary.failed, testData.summary.pending];
    const labels = ['Exitosas', 'Fallidas', 'Pendientes'];
    const chartColors = [colors.success, colors.danger, colors.pending];
    
    // Agregar planificadas si existe
    if (testData.summary.planned && testData.summary.planned > 0) {
        values.push(testData.summary.planned);
        labels.push('Planificadas');
        chartColors.push('#9CA3AF');
    }
    
    // Agregar bloqueadas si existe
    if (testData.summary.blocked && testData.summary.blocked > 0) {
        values.push(testData.summary.blocked);
        labels.push('Bloqueadas');
        chartColors.push('#6B7280');
    }
    
    // Agregar desestimadas si existe
    if (testData.summary.dismissed && testData.summary.dismissed > 0) {
        values.push(testData.summary.dismissed);
        labels.push('Desestimadas');
        chartColors.push('#8B8B8B');
    }
    
    const data = [{
        values: values,
        labels: labels,
        type: 'pie',
        hole: 0.4,
        marker: {
            colors: chartColors
        },
        textinfo: 'label+percent',
        textposition: 'outside',
        hovertemplate: '<b>%{label}</b><br>Cantidad: %{value}<br>Porcentaje: %{percent}<extra></extra>'
    }];

    const layout = {
        margin: { t: 20, b: 20, l: 20, r: 20 },
        showlegend: false,
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: {
            family: 'Inter, sans-serif',
            size: 12,
            color: '#1e293b'
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('pieChart', data, layout, config);
}

// Gráfica de tendencia - Desde fecha de inicio hasta fecha de fin del proyecto
function createTrendChart() {
    // Convertir fechas del proyecto de DD/MM/YYYY a YYYY-MM-DD para comparación
    const projectStartDate = testData.projectInfo.startDate.split('/').reverse().join('-');
    const projectEndDate = testData.projectInfo.endDate.split('/').reverse().join('-');
    
    // Filtrar datos de tendencia según las fechas del proyecto
    const filteredTrendData = testData.trend.filter(d => {
        return d.date >= projectStartDate && d.date <= projectEndDate;
    });
    
    // Si no hay datos filtrados, usar todos los datos disponibles
    const trendData = filteredTrendData.length > 0 ? filteredTrendData : testData.trend;
    const dates = trendData.map(d => d.date);
    
    const plannedTrace = {
        x: dates,
        y: trendData.map(d => d.planned),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Planificadas',
        line: { color: colors.warning, width: 2, dash: 'dash' },
        marker: { size: 5 }
    };
    
    const successTrace = {
        x: dates,
        y: trendData.map(d => d.successful),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Exitosas',
        line: { color: colors.success, width: 3 },
        marker: { size: 6 }
    };

    const failureTrace = {
        x: dates,
        y: trendData.map(d => d.failed),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Fallidas',
        line: { color: colors.danger, width: 3 },
        marker: { size: 6 }
    };

    const pendingTrace = {
        x: dates,
        y: trendData.map(d => d.pending),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Pendientes',
        line: { color: colors.pending, width: 3 },
        marker: { size: 6 }
    };

    const blockedTrace = {
        x: dates,
        y: trendData.map(d => d.blocked || 0),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Bloqueadas',
        line: { color: '#6B7280', width: 3 },
        marker: { size: 6 }
    };

    const dismissedTrace = {
        x: dates,
        y: trendData.map(d => d.dismissed || 0),
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Desestimadas',
        line: { color: '#8B8B8B', width: 3 },
        marker: { size: 6 }
    };

    const data = [plannedTrace, successTrace, failureTrace, pendingTrace, blockedTrace, dismissedTrace];

    const layout = {
        margin: { t: 20, b: 40, l: 40, r: 20 },
        showlegend: false,
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        xaxis: {
            gridcolor: '#e2e8f0',
            linecolor: '#e2e8f0',
            tickformat: '%m/%d'
        },
        yaxis: {
            gridcolor: '#e2e8f0',
            linecolor: '#e2e8f0'
        },
        font: {
            family: 'Inter, sans-serif',
            size: 11,
            color: '#64748b'
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('trendChart', data, layout, config);
}

// Gráfica de barras - Pruebas por escenario
function createCategoryChart() {
    const categories = testData.categories.map(c => c.escenario);
    
    const plannedTrace = {
        x: categories,
        y: testData.categories.map(c => c.planned),
        type: 'bar',
        name: 'Planificadas',
        marker: { 
            color: colors.warning,
            opacity: 0.3
        }
    };
    
    const successTrace = {
        x: categories,
        y: testData.categories.map(c => c.successful),
        type: 'bar',
        name: 'Exitosas',
        marker: { color: colors.success }
    };

    const failureTrace = {
        x: categories,
        y: testData.categories.map(c => c.failed),
        type: 'bar',
        name: 'Fallidas',
        marker: { color: colors.danger }
    };

    const pendingTrace = {
        x: categories,
        y: testData.categories.map(c => c.pending),
        type: 'bar',
        name: 'Pendientes',
        marker: { color: colors.pending }
    };

    const blockedTrace = {
        x: categories,
        y: testData.categories.map(c => c.blocked || 0),
        type: 'bar',
        name: 'Bloqueadas',
        marker: { color: '#6B7280' }
    };

    const dismissedTrace = {
        x: categories,
        y: testData.categories.map(c => c.dismissed || 0),
        type: 'bar',
        name: 'Desestimadas',
        marker: { color: '#8B8B8B' }
    };

    const data = [plannedTrace, successTrace, failureTrace, pendingTrace, blockedTrace, dismissedTrace];

    const layout = {
        barmode: 'stack',
        margin: { t: 20, b: 80, l: 40, r: 20 },
        showlegend: false,
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        xaxis: {
            gridcolor: '#e2e8f0',
            linecolor: '#e2e8f0',
            tickangle: -45
        },
        yaxis: {
            gridcolor: '#e2e8f0',
            linecolor: '#e2e8f0'
        },
        font: {
            family: 'Inter, sans-serif',
            size: 11,
            color: '#64748b'
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('categoryChart', data, layout, config);
}

// Gráfica de defectos por severidad
function createDefectsChart() {
    const defectsData = testData.defects.summary;
    
    // Actualizar contadores en el HTML
    document.getElementById('criticalDefects').textContent = defectsData.critical;
    document.getElementById('highDefects').textContent = defectsData.high;
    document.getElementById('mediumDefects').textContent = defectsData.medium;
    document.getElementById('lowDefects').textContent = defectsData.low;
    
    const data = [{
        y: ['Críticos', 'Altos', 'Medios', 'Bajos'],
        x: [defectsData.critical, defectsData.high, defectsData.medium, defectsData.low],
        type: 'bar',
        orientation: 'h',
        marker: {
            color: ['#dc2626', '#fb923c', '#f59e0b', '#22c55e'],
            line: {
                color: '#ffffff',
                width: 1
            }
        },
        text: [defectsData.critical, defectsData.high, defectsData.medium, defectsData.low],
        textposition: 'inside',
        textfont: {
            color: 'white',
            weight: 'bold'
        },
        hovertemplate: '<b>%{y}</b><br>Cantidad: %{x}<extra></extra>'
    }];

    const layout = {
        margin: { t: 20, b: 40, l: 80, r: 40 },
        showlegend: false,
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: {
            family: 'Inter, sans-serif',
            size: 12,
            color: '#1e293b'
        },
        xaxis: {
            showgrid: true,
            gridcolor: 'rgba(148, 163, 184, 0.1)',
            zeroline: false,
            tickfont: { size: 11, color: '#64748b' },
            title: {
                text: 'Cantidad de Defectos',
                font: { size: 12, color: '#64748b' }
            }
        },
        yaxis: {
            showgrid: false,
            zeroline: false,
            tickfont: { size: 12, color: '#1e293b' },
            categoryorder: 'array',
            categoryarray: ['Bajos', 'Medios', 'Altos', 'Críticos']
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('defectsChart', data, layout, config);
}

// Función para actualizar Burndown Chart
function updateDefectCycleTimeChart(defectsData) {
    if (!defectsData || !defectsData.details || defectsData.details.length === 0) {
        console.log('No hay datos de defectos para el cycle time chart');
        // Mostrar gráfico vacío
        const layout = {
            title: {
                text: '',
                font: { size: 14, color: '#1e293b' }
            },
            xaxis: {
                title: 'Fecha de Resolución',
                gridcolor: '#f1f5f9'
            },
            yaxis: {
                title: 'Tiempo de Ciclo (días laborales)',
                gridcolor: '#f1f5f9'
            },
            plot_bgcolor: 'rgba(0,0,0,0)',
            paper_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
            margin: { t: 30, r: 30, b: 80, l: 60 },
            annotations: [{
                text: 'No hay datos de defectos disponibles',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('burndownChart', [], layout, config);
        return;
    }

    console.log('� Datos de defectos para cycle time:', defectsData);

    // Función helper para parsear fechas de manera más robusta
    function parseExcelDate(dateStr) {
        if (!dateStr || dateStr === '') return null;
        
        let date;
        
        // Si es un número (fecha serial de Excel)
        if (typeof dateStr === 'number') {
            const excelEpoch = new Date(1900, 0, 1);
            const millisecondsPerDay = 24 * 60 * 60 * 1000;
            let adjustedDays = dateStr;
            if (dateStr > 59) {
                adjustedDays = dateStr - 1;
            }
            date = new Date(excelEpoch.getTime() + (adjustedDays - 1) * millisecondsPerDay);
        }
        // Si es un string
        else if (typeof dateStr === 'string') {
            // Formato DD/MM/AAAA HH:MM
            if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}$/)) {
                const [datePart, timePart] = dateStr.split(/\s+/);
                const [day, month, year] = datePart.split('/').map(Number);
                const [hour, minute] = timePart.split(':').map(Number);
                date = new Date(year, month - 1, day, hour, minute);
            }
            // Formato DD/MM/AAAA
            else if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
                const [day, month, year] = dateStr.split('/').map(Number);
                date = new Date(year, month - 1, day);
            }
            else {
                // Intentar parsear directamente
                date = new Date(dateStr);
            }
        }
        else {
            date = new Date(dateStr);
        }
        
        // Validar que la fecha sea válida
        if (isNaN(date.getTime())) {
            console.warn('📊 No se pudo parsear la fecha:', dateStr);
            return null;
        }
        
        return date;
    }

    // Calcular cycle time para cada defecto resuelto
    const cycleTimeData = defectsData.details
        .filter(defect => {
            // Buscar tanto dateResolved como dateResolution para compatibilidad
            const resolvedDate = defect.dateResolved || defect.dateResolution;
            const foundDate = defect.dateFound;
            const hasResolved = resolvedDate && resolvedDate !== '';
            const hasFound = foundDate && foundDate !== '';
            
            console.log(`📊 Defecto ${defect.id}: dateFound="${foundDate}", dateResolved="${defect.dateResolved}", dateResolution="${defect.dateResolution}", incluir=${hasResolved && hasFound}`);
            return hasResolved && hasFound;
        }) // Solo defectos resueltos
        .map(defect => {
            const foundDate = parseExcelDate(defect.dateFound);
            // Usar dateResolved o dateResolution según esté disponible
            const resolvedDateStr = defect.dateResolved || defect.dateResolution;
            const resolvedDate = parseExcelDate(resolvedDateStr);
            
            // Validar que las fechas sean válidas antes de calcular
            if (!foundDate || !resolvedDate) {
                console.warn(`📊 Fechas inválidas para defecto ${defect.id}:`, {
                    dateFound: defect.dateFound,
                    dateResolvedUsed: resolvedDateStr,
                    foundDateValid: foundDate !== null,
                    resolvedDateValid: resolvedDate !== null
                });
                return null; // Retornar null para filtrar después
            }
            
            const cycleTimeDays = calculateBusinessDays(foundDate, resolvedDate);
            
            console.log(`📊 Cálculo cycle time para ${defect.id} (SOLO DÍAS LABORALES):`, {
                dateFound: defect.dateFound,
                dateResolvedUsed: resolvedDateStr,
                foundDate: foundDate.toISOString(),
                resolvedDate: resolvedDate.toISOString(),
                cycleTimeDays: cycleTimeDays,
                note: 'Calculado con días laborales (excluyendo sábados y domingos)'
            });
            
            return {
                id: defect.id,
                title: defect.title,
                severity: defect.severity,
                resolvedDate: resolvedDate,
                cycleTime: Math.max(0, cycleTimeDays), // Evitar valores negativos
                status: defect.status
            };
        })
        .filter(item => item !== null) // Filtrar defectos con fechas inválidas
        .sort((a, b) => a.resolvedDate - b.resolvedDate); // Ordenar por fecha de resolución

    console.log('� Datos de cycle time calculados:', cycleTimeData);

    if (cycleTimeData.length === 0) {
        // Mostrar gráfico indicando que no hay defectos resueltos
        const layout = {
            title: {
                text: '',
                font: { size: 14, color: '#1e293b' }
            },
            xaxis: {
                title: 'Fecha de Resolución',
                gridcolor: '#f1f5f9'
            },
            yaxis: {
                title: 'Tiempo de Ciclo (días laborales)',
                gridcolor: '#f1f5f9'
            },
            plot_bgcolor: 'rgba(0,0,0,0)',
            paper_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
            margin: { t: 30, r: 30, b: 80, l: 60 },
            annotations: [{
                text: 'No hay defectos resueltos para mostrar',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('burndownChart', [], layout, config);
        return;
    }

    // Preparar datos para barras verticales
    const defectIds = cycleTimeData.map(item => item.id);
    const cycleTimes = cycleTimeData.map(item => item.cycleTime);
    
    // Colores basados en si cumple o no el objetivo de 3 días
    const colors = cycleTimeData.map(item => {
        if (item.cycleTime <= 3) {
            // Verde - Cumple el objetivo
            return '#10b981';
        } else if (item.cycleTime <= 5) {
            // Amarillo - Aceptable pero por encima del objetivo
            return '#f59e0b';
        } else {
            // Rojo - Muy por encima del objetivo
            return '#ef4444';
        }
    });

    // Textos hover con información detallada
    const hoverTexts = cycleTimeData.map(item => 
        `ID: ${item.id}<br>` +
        `Título: ${item.title.length > 40 ? item.title.substring(0, 40) + '...' : item.title}<br>` +
        `Severidad: ${item.severity}<br>` +
        `Cycle Time: ${item.cycleTime} días laborales<br>` +
        `Estado: ${item.cycleTime <= 3 ? '✅ Dentro del objetivo' : '⚠️ Fuera del objetivo'}`
    );

    // Crear el trace para las barras
    const barTrace = {
        x: defectIds,
        y: cycleTimes,
        type: 'bar',
        name: 'Cycle Time',
        marker: {
            color: colors,
            line: {
                color: '#374151',
                width: 1
            }
        },
        text: hoverTexts,
        hovertemplate: '%{text}<extra></extra>',
        textposition: 'outside',
        texttemplate: '%{y}d'
    };

    // Línea horizontal para el objetivo de 3 días laborales
    const objectiveLine = {
        x: defectIds,
        y: Array(defectIds.length).fill(3),
        type: 'scatter',
        mode: 'lines',
        name: 'Objetivo (3 días laborales)',
        line: {
            color: '#dc2626',
            width: 2,
            dash: 'dash'
        },
        hovertemplate: 'Objetivo: 3 días laborales<extra></extra>'
    };

    // Calcular estadísticas
    const averageCycleTime = cycleTimes.reduce((sum, time) => sum + time, 0) / cycleTimes.length;
    const defectsWithinObjective = cycleTimes.filter(time => time <= 3).length;
    const complianceRate = (defectsWithinObjective / cycleTimes.length * 100).toFixed(1);

    const layout = {
        title: {
            text: '',
            font: { size: 14, color: '#1e293b' }
        },
        xaxis: {
            title: 'Defectos',
            gridcolor: '#f1f5f9',
            tickangle: -45,
            type: 'category'
        },
        
        plot_bgcolor: 'rgba(0,0,0,0)',
        paper_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
        margin: { t: 30, r: 30, b: 100, l: 80 },
        hovermode: 'closest',
        showlegend: true,
        legend: {
            x: 0,
            y: 1.1,
            orientation: 'h'
        },
        shapes: [{
            type: 'rect',
            xref: 'paper',
            yref: 'y',
            x0: 0,
            y0: 0,
            x1: 1,
            y1: 3,
            fillcolor: 'rgba(16, 185, 129, 0.1)',
            line: { width: 0 }
        }]
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    console.log('📊 Renderizando gráfico de barras Time To Fix');
    Plotly.newPlot('burndownChart', [barTrace, objectiveLine], layout, config);
}

// Función para actualizar Cobertura por Módulo
function updateCoverageChart(detailsData) {
    if (!detailsData || detailsData.length === 0) {
        console.log('⚠️ No hay datos de detalles para cobertura por escenario');
        return;
    }

    console.log(`🔍 DEBUGGING COBERTURA: Actualizando gráfico de cobertura con ${detailsData.length} pruebas`);
    console.log('🔍 DEBUGGING COBERTURA: Muestra de datos:', detailsData.slice(0, 3));

    // Agrupar por escenario/módulo
    const moduleStats = {};
    
    detailsData.forEach((test, index) => {
        const module = test.escenario || test.categoria || test.modulo || 'Sin Escenario';
        const status = test.status || '';
        
        if (!moduleStats[module]) {
            moduleStats[module] = {
                total: 0,
                completed: 0,
                planned: 0,
                exitosas: 0,
                desestimadas: 0,
                fallidas: 0,
                pendientes: 0,
                bloqueadas: 0
            };
        }
        
        moduleStats[module].total++;
        
        // Contar por tipo específico para debugging
        switch(status) {
            case 'success':
                moduleStats[module].exitosas++;
                moduleStats[module].completed++;
                break;
            case 'dismissed':
                moduleStats[module].desestimadas++;
                moduleStats[module].completed++;
                break;
            case 'failure':
                moduleStats[module].fallidas++;
                moduleStats[module].planned++;
                break;
            case 'pending':
                moduleStats[module].pendientes++;
                moduleStats[module].planned++;
                break;
            case 'blocked':
                moduleStats[module].bloqueadas++;
                moduleStats[module].planned++;
                break;
            default:
                console.warn(`⚠️ Estado desconocido en cobertura: "${status}" para prueba ${index + 1}`);
                moduleStats[module].planned++;
        }
        
        // Log detallado para las primeras pruebas de cada módulo
        if (moduleStats[module].total <= 3) {
            const cubierto = (status === 'success' || status === 'dismissed') ? 'CUBIERTO' : 'NO CUBIERTO';
            console.log(`📊 COBERTURA ${module}: Prueba "${test.name}" con status "${status}" → ${cubierto}`);
        }
    });

    console.log(`📈 DEBUGGING COBERTURA: Estadísticas detalladas por escenario:`);
    Object.keys(moduleStats).forEach(module => {
        const stats = moduleStats[module];
        const porcentaje = Math.round((stats.completed / stats.total) * 100);
        console.log(`   • ${module}:`);
        console.log(`     - Total: ${stats.total}`);
        console.log(`     - Exitosas: ${stats.exitosas}`);
        console.log(`     - Desestimadas: ${stats.desestimadas}`);
        console.log(`     - Completadas (exitosas + desestimadas): ${stats.completed}`);
        console.log(`     - Fallidas: ${stats.fallidas}`);
        console.log(`     - Pendientes: ${stats.pendientes}`);
        console.log(`     - Bloqueadas: ${stats.bloqueadas}`);
        console.log(`     - COBERTURA: ${porcentaje}%`);
    });

    // Calcular porcentajes de cobertura
    const modules = Object.keys(moduleStats);
    const coveragePercentages = modules.map(module => {
        const stats = moduleStats[module];
        return Math.round((stats.completed / stats.total) * 100);
    });

    console.log('📊 Porcentajes de cobertura final:', modules.map((m, i) => `${m}: ${coveragePercentages[i]}%`));
    
    // Mostrar módulos con 100% de éxito (barras verdes)
    const perfectModules = modules.filter((m, i) => coveragePercentages[i] === 100);
    if (perfectModules.length > 0) {
        console.log('🟢 Módulos con 100% exitosos (barras verdes):', perfectModules);
    }

    // Actualizar estadísticas generales - 100% = todos los casos exitosos
    const totalModules = modules.length;
    const completeModules = coveragePercentages.filter(pct => pct === 100).length; // Módulos con todos los casos exitosos
    const avgCoverage = coveragePercentages.length > 0 ? 
        Math.round(coveragePercentages.reduce((a, b) => a + b, 0) / coveragePercentages.length) : 0;

    // Actualizar elementos HTML
    const avgCoverageElement = document.getElementById('avgCoverage');
    const completeModulesElement = document.getElementById('completeModules');
    
    if (avgCoverageElement) {
        avgCoverageElement.textContent = `${avgCoverage}%`;
    }
    if (completeModulesElement) {
        completeModulesElement.textContent = completeModules;
    }

    console.log(`📊 Estadísticas generales: Promedio=${avgCoverage}%, Completos=${completeModules}/${totalModules}`);

    // Crear gráfico de barras horizontal
    const trace = {
        x: coveragePercentages,
        y: modules,
        type: 'bar',
        orientation: 'h',
        marker: {
            color: coveragePercentages.map(pct => {
                if (pct === 100) return '#22c55e'; // Verde brillante - TODOS EXITOSOS
                if (pct >= 80) return '#84cc16';   // Verde claro
                if (pct >= 60) return '#eab308';   // Amarillo
                if (pct >= 40) return '#f97316';   // Naranja
                if (pct > 0) return '#ef4444';     // Rojo - algunos exitosos
                return '#6b7280';                  // Gris - ningún exitoso
            }),
            line: {
                color: '#e2e8f0',
                width: 1
            }
        },
        text: coveragePercentages.map(pct => `${pct}%`),
        textposition: 'inside',
        insidetextanchor: 'middle',
        textfont: {
            color: 'white',
            size: 12,
            family: 'Inter, sans-serif'
        }
    };

    const layout = {
        title: {
            text: '',
            font: { size: 14, color: '#1e293b' }
        },
        xaxis: {
            title: 'Porcentaje de Cobertura (%)',
            gridcolor: '#f1f5f9',
            range: [0, 100]
        },
        yaxis: {
            title: '',
            gridcolor: '#f1f5f9',
            automargin: true,
            tickfont: {
                size: 11,
                family: 'Inter, sans-serif'
            }
        },
        
        plot_bgcolor: 'rgba(0,0,0,0)',
        paper_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
        margin: { t: 30, r: 30, b: 60, l: 200 },
        showlegend: false
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('coverageChart', [trace], layout, config);
}

// Función para actualizar Ejecución Diaria por Ejecutor
function updateDailyExecutionChart(detailsData) {
    // Verificar que el elemento DOM existe
    const chartElement = document.getElementById('dailyExecutionChart');
    if (!chartElement) {
        console.warn('Elemento dailyExecutionChart no encontrado en el DOM');
        return;
    }

    if (!detailsData || detailsData.length === 0) {
        console.log('No hay datos de pruebas para el gráfico de ejecución diaria');
        
        // Mostrar gráfico vacío
        const layout = {
            title: {
                text: '',
                font: { size: 14, color: '#1e293b' }
            },
            xaxis: {
                title: 'Ejecutores',
                gridcolor: '#f1f5f9'
            },
            yaxis: {
                title: 'Casos de Prueba Ejecutados',
                gridcolor: '#f1f5f9'
            },
            plot_bgcolor: 'rgba(0,0,0,0)',
            paper_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
            margin: { t: 30, r: 30, b: 80, l: 60 },
            annotations: [{
                text: 'No hay datos de ejecución disponibles',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('dailyExecutionChart', [], layout, config);
        return;
    }

    console.log('🔍 Actualizando gráfico de ejecución diaria con', detailsData.length, 'pruebas');

    // Filtrar solo las pruebas que tienen fecha de ejecución y ejecutor
    const executedTests = detailsData.filter(test => {
        return test.dia_ejecutado && test.dia_ejecutado.trim() !== '' && 
               test.executor && test.executor.trim() !== '';
    });

    console.log('📈 Pruebas con fecha y ejecutor:', executedTests.length);

    if (executedTests.length === 0) {
        // Mostrar gráfico indicando que no hay datos de ejecución
        const layout = {
            title: {
                text: '',
                font: { size: 14, color: '#1e293b' }
            },
            xaxis: {
                title: 'Ejecutores',
                gridcolor: '#f1f5f9'
            },
            yaxis: {
                title: 'Casos de Prueba Ejecutados',
                gridcolor: '#f1f5f9'
            },
            plot_bgcolor: 'rgba(0,0,0,0)',
            paper_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
            margin: { t: 30, r: 30, b: 80, l: 60 },
            annotations: [{
                text: 'No hay pruebas ejecutadas para mostrar',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('dailyExecutionChart', [], layout, config);
        return;
    }

    // Agrupar por ejecutor y fecha
    const executionStats = {};
    
    executedTests.forEach(test => {
        const executor = test.executor.trim();
        const date = test.dia_ejecutado.trim();
        
        if (!executionStats[executor]) {
            executionStats[executor] = {};
        }
        
        if (!executionStats[executor][date]) {
            executionStats[executor][date] = 0;
        }
        
        executionStats[executor][date]++;
    });

    console.log('📊 Estadísticas de ejecución por ejecutor:', executionStats);

    // Obtener todas las fechas únicas y ordenarlas
    const allDates = [...new Set(executedTests.map(test => test.dia_ejecutado.trim()))];
    allDates.sort((a, b) => {
        // Convertir fechas DD/MM/YYYY a Date para ordenar correctamente
        const parseDate = (dateStr) => {
            const [day, month, year] = dateStr.split('/');
            return new Date(year, month - 1, day);
        };
        return parseDate(a) - parseDate(b);
    });

    console.log('📅 Fechas encontradas (ordenadas):', allDates);

    // Crear traces para cada ejecutor
    const executors = Object.keys(executionStats);
    const traces = [];
    
    // Colores para diferentes ejecutores
    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#06b6d4', '#84cc16', '#f97316', '#ec4899', '#6366f1'
    ];

    executors.forEach((executor, index) => {
        const executorData = allDates.map(date => executionStats[executor][date] || 0);
        
        traces.push({
            x: allDates,
            y: executorData,
            type: 'bar',
            name: executor,
            marker: {
                color: colors[index % colors.length],
                line: {
                    color: '#ffffff',
                    width: 1
                }
            },
            hovertemplate: `<b>${executor}</b><br>Fecha: %{x}<br>Casos ejecutados: %{y}<extra></extra>`
        });
    });

    // Calcular estadísticas generales
    const totalExecutors = executors.length;
    const totalCasesExecuted = executedTests.length;
    const avgDailyExecution = allDates.length > 0 ? Math.round(totalCasesExecuted / allDates.length) : 0;

    // Actualizar elementos HTML
    const totalExecutorsElement = document.getElementById('totalExecutors');
    const avgDailyExecutionElement = document.getElementById('avgDailyExecution');
    
    if (totalExecutorsElement) {
        totalExecutorsElement.textContent = totalExecutors;
    }
    if (avgDailyExecutionElement) {
        avgDailyExecutionElement.textContent = avgDailyExecution;
    }

    console.log(`📊 Estadísticas generales: ${totalExecutors} ejecutores, ${avgDailyExecution} casos promedio por día`);

    const layout = {
        title: {
            text: '',
            font: { size: 14, color: '#1e293b' }
        },
        yaxis: {
            title: 'Casos de Prueba Ejecutados',
            gridcolor: '#f1f5f9'
        },
        plot_bgcolor: 'rgba(0,0,0,0)',
        paper_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Inter, sans-serif', size: 12, color: '#64748b' },
        margin: { t: 30, r: 30, b: 100, l: 60 },
        barmode: 'group',
        showlegend: true,
        legend: {
            x: 0,
            y: 1.02,
            orientation: 'h',
            bgcolor: 'rgba(255,255,255,0.8)',
            bordercolor: '#e2e8f0',
            borderwidth: 1
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('dailyExecutionChart', traces, layout, config);
}

// Función para actualizar Distribución Total por Ejecutor (Pie Chart)
function updateExecutorDistributionChart(detailsData) {
    // Verificar que el elemento DOM existe
    const chartElement = document.getElementById('executorDistributionChart');
    if (!chartElement) {
        console.warn('Elemento executorDistributionChart no encontrado en el DOM');
        return;
    }

    if (!detailsData || detailsData.length === 0) {
        console.log('No hay datos de pruebas para el gráfico de distribución por ejecutor');
        
        // Mostrar gráfico vacío
        const layout = {
            margin: { t: 20, b: 20, l: 20, r: 20 },
            showlegend: false,
            paper_bgcolor: 'transparent',
            plot_bgcolor: 'transparent',
            font: {
                family: 'Inter, sans-serif',
                size: 12,
                color: '#1e293b'
            },
            annotations: [{
                text: 'No hay datos de ejecución disponibles',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('executorDistributionChart', [], layout, config);
        return;
    }

    console.log('🍰 Actualizando gráfico de distribución por ejecutor con', detailsData.length, 'pruebas');

    // Filtrar solo las pruebas que tienen fecha de ejecución y ejecutor
    const executedTests = detailsData.filter(test => {
        return test.dia_ejecutado && test.dia_ejecutado.trim() !== '' && 
               test.executor && test.executor.trim() !== '';
    });

    console.log('📈 Pruebas ejecutadas con datos completos:', executedTests.length);

    if (executedTests.length === 0) {
        // Mostrar gráfico indicando que no hay datos de ejecución
        const layout = {
            margin: { t: 20, b: 20, l: 20, r: 20 },
            showlegend: false,
            paper_bgcolor: 'transparent',
            plot_bgcolor: 'transparent',
            font: {
                family: 'Inter, sans-serif',
                size: 12,
                color: '#1e293b'
            },
            annotations: [{
                text: 'No hay pruebas ejecutadas para mostrar',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        const config = {
            responsive: true,
            displayModeBar: false
        };

        Plotly.newPlot('executorDistributionChart', [], layout, config);
        return;
    }

    // Contar casos ejecutados por cada ejecutor
    const executorCounts = {};
    
    executedTests.forEach(test => {
        const executor = test.executor.trim();
        if (!executorCounts[executor]) {
            executorCounts[executor] = 0;
        }
        executorCounts[executor]++;
    });

    console.log('📊 Distribución por ejecutor:', executorCounts);

    // Ordenar ejecutores por cantidad de casos (descendente)
    const sortedExecutors = Object.entries(executorCounts)
        .sort((a, b) => b[1] - a[1]);

    const executors = sortedExecutors.map(entry => entry[0]);
    const counts = sortedExecutors.map(entry => entry[1]);

    // Colores para el gráfico circular
    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#06b6d4', '#84cc16', '#f97316', '#ec4899', '#6366f1',
        '#14b8a6', '#f472b6', '#a855f7', '#22d3ee', '#facc15'
    ];

    // Calcular estadísticas generales
    const totalExecutedCases = executedTests.length;
    const topExecutor = sortedExecutors.length > 0 ? 
        `${sortedExecutors[0][0]} (${sortedExecutors[0][1]} casos)` : '-';

    console.log(`📊 Total de casos ejecutados: ${totalExecutedCases}, ejecutor principal: ${topExecutor}`);

    // Crear el gráfico circular (pie chart sin hueco)
    const data = [{
        values: counts,
        labels: executors,
        type: 'pie',
        hole: 0, // Gráfico circular completo (no donut)
        marker: {
            colors: colors.slice(0, executors.length)
        },
        textinfo: 'label+percent+value',
        textposition: 'outside',
        hovertemplate: '<b>%{label}</b><br>' +
                      'Casos ejecutados: %{value}<br>' +
                      'Porcentaje: %{percent}<br>' +
                      '<extra></extra>',
        textfont: {
            size: 12,
            color: '#1e293b'
        }
    }];

    const layout = {
        margin: { t: 20, b: 80, l: 20, r: 20 },
        showlegend: true,
        legend: {
            orientation: 'h',
            x: 0.5,
            y: -0.1,
            xanchor: 'center',
            yanchor: 'top',
            font: {
                size: 11,
                color: '#1e293b'
            }
        },
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: {
            family: 'Inter, sans-serif',
            size: 12,
            color: '#1e293b'
        }
    };

    const config = {
        responsive: true,
        displayModeBar: false
    };

    Plotly.newPlot('executorDistributionChart', data, layout, config);
}

// Variables para paginación
let currentPage = 1;
let itemsPerPage = 5;
let filteredTests = [];

// Variables para paginación de defectos
let defectsCurrentPage = 1;
let defectsItemsPerPage = 5;
let filteredDefects = [];

// Función para resaltar términos de búsqueda
function highlightSearchTerm(text, searchTerm) {
    if (!searchTerm || searchTerm.trim() === '') {
        return text;
    }
    
    // Convertir a string si no lo es
    const textStr = String(text || '');
    
    const searchRegex = new RegExp(`(${searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
    return textStr.replace(searchRegex, '<mark style="background-color: #fff3cd; color: #856404; padding: 1px 2px; border-radius: 2px;">$1</mark>');
}

// Poblar tabla de detalles de pruebas con paginación mejorada
function populateTestTable() {
    const tableBody = document.getElementById('testTableBody');
    if (!tableBody) return;
    
    // Verificar si hay datos de pruebas
    if (!testData.testDetails || testData.testDetails.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="6" class="no-data-message">
                    <div class="no-data-content">
                        <i class="fas fa-vial"></i>
                        <p>No hay pruebas registradas</p>
                        <small>Las pruebas aparecerán aquí cuando se carguen datos</small>
                    </div>
                </td>
            </tr>
        `;
        updatePaginationControls(0, 0);
        return;
    }
    
    // Si no hay filtro aplicado, usar todas las pruebas
    if (filteredTests.length === 0) {
        // Verificar si realmente no hay filtros aplicados
        const searchInput = document.getElementById('searchInput');
        const statusFilter = document.getElementById('testStatusFilter');
        
        const searchTerm = searchInput ? searchInput.value.trim() : '';
        const statusValue = statusFilter ? statusFilter.value : '';
        
        const hasActiveFilters = searchTerm !== '' || statusValue !== '';
        
        if (!hasActiveFilters) {
            filteredTests = [...testData.testDetails];
        }
    }
    
    const totalItems = filteredTests.length;
    const totalPages = itemsPerPage === 100 ? 1 : Math.ceil(totalItems / itemsPerPage);
    
    // Limpiar tabla
    tableBody.innerHTML = '';
    
    // Si no hay resultados después del filtrado, mostrar mensaje
    if (totalItems === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="6" class="no-results-message">
                    <div class="no-results-content">
                        <i class="fas fa-search"></i>
                        <p>No se encontraron pruebas que coincidan con los filtros</p>
                        <small>Intente modificar o limpiar los filtros para ver más resultados</small>
                    </div>
                </td>
            </tr>
        `;
        updatePaginationControls(0, 0);
        return;
    }
    
    // Calcular índices para la página actual
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = itemsPerPage === 100 ? totalItems : Math.min(startIndex + itemsPerPage, totalItems);
    
    // Obtener término de búsqueda actual para resaltado
    const searchInput = document.getElementById('searchInput');
    const currentSearchTerm = searchInput ? searchInput.value.trim() : '';
    
    // Mostrar tests de la página actual
    const testsToShow = filteredTests.slice(startIndex, endIndex);
    
    testsToShow.forEach(test => {
        const row = document.createElement('tr');
        
        // Aplicar resaltado si hay término de búsqueda
        const highlightedId = highlightSearchTerm(test.id, currentSearchTerm);
        const highlightedName = highlightSearchTerm(test.name, currentSearchTerm);
        const highlightedEscenario = highlightSearchTerm(test.escenario, currentSearchTerm);
        const statusText = getStatusText(test.status);
        const highlightedStatus = highlightSearchTerm(statusText, currentSearchTerm);
        const highlightedExecutor = highlightSearchTerm(test.executor, currentSearchTerm);
        
        // Formatear fecha de ejecución
        const fechaEjecucion = test.dia_ejecutado ? formatFecha(test.dia_ejecutado) : '-';
        const highlightedFechaEjecucion = highlightSearchTerm(fechaEjecucion, currentSearchTerm);
        
        row.innerHTML = `
            <td>${highlightedId}</td>
            <td>${highlightedName}</td>
            <td>${highlightedEscenario}</td>
            <td><span class="status-badge ${test.status}">${highlightedStatus}</span></td>
            <td>${highlightedFechaEjecucion}</td>
            <td>${highlightedExecutor}</td>
        `;
        tableBody.appendChild(row);
    });
    
    // Actualizar controles de paginación
    updatePaginationControls(totalItems, totalPages);
}

// Actualizar controles de paginación
function updatePaginationControls(totalItems, totalPages) {
    const paginationInfo = document.getElementById('paginationInfo');
    const prevBtn = document.getElementById('prevPageBtn');
    const nextBtn = document.getElementById('nextPageBtn');
    const pageNumbers = document.getElementById('pageNumbers');
    
    // Calcular rango actual
    const startIndex = (currentPage - 1) * itemsPerPage + 1;
    const endIndex = Math.min(currentPage * itemsPerPage, totalItems);
    
    // Actualizar información
    if (totalItems === 0) {
        paginationInfo.textContent = 'No hay registros para mostrar';
    } else {
        paginationInfo.textContent = `Mostrando ${startIndex}-${endIndex} de ${totalItems} registros`;
    }
    
    // Botones anterior/siguiente
    prevBtn.disabled = currentPage === 1;
    nextBtn.disabled = currentPage === totalPages || totalPages === 0;
    
    // Generar números de página
    generatePageNumbers(pageNumbers, currentPage, totalPages);
}

// Generar números de página
function generatePageNumbers(container, current, total) {
    container.innerHTML = '';
    
    if (total <= 1) return;
    
    const maxVisiblePages = 5;
    let startPage = Math.max(1, current - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(total, startPage + maxVisiblePages - 1);
    
    // Ajustar si no hay suficientes páginas al final
    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }
    
    // Botón primera página
    if (startPage > 1) {
        createPageButton(container, 1, '1');
        if (startPage > 2) {
            createPageButton(container, null, '...');
        }
    }
    
    // Páginas visibles
    for (let i = startPage; i <= endPage; i++) {
        createPageButton(container, i, i.toString());
    }
    
    // Botón última página
    if (endPage < total) {
        if (endPage < total - 1) {
            createPageButton(container, null, '...');
        }
        createPageButton(container, total, total.toString());
    }
}

// Crear botón de página
function createPageButton(container, pageNum, text) {
    const button = document.createElement('button');
    button.className = 'page-number';
    button.textContent = text;
    
    if (pageNum === null) {
        button.classList.add('ellipsis');
    } else {
        if (pageNum === currentPage) {
            button.classList.add('active');
        }
        button.addEventListener('click', () => goToPage(pageNum));
    }
    
    container.appendChild(button);
}

// Ir a una página específica
function goToPage(page) {
    currentPage = page;
    populateTestTable();
}

// Ir a página anterior
function goToPrevPage() {
    if (currentPage > 1) {
        currentPage--;
        populateTestTable();
    }
}

// Ir a página siguiente
function goToNextPage() {
    const totalPages = itemsPerPage === 100 ? 1 : Math.ceil(filteredTests.length / itemsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        populateTestTable();
    }
}

// Cambiar elementos por página
function changeItemsPerPage(newItemsPerPage) {
    itemsPerPage = parseInt(newItemsPerPage);
    currentPage = 1; // Resetear a primera página
    populateTestTable();
}

// ===== FUNCIONES DE PAGINACIÓN PARA DEFECTOS =====

// Actualizar controles de paginación para defectos
function updateDefectsPaginationControls(totalItems, totalPages) {
    const paginationInfo = document.getElementById('defectsPaginationInfo');
    const prevBtn = document.getElementById('defectsPrevPageBtn');
    const nextBtn = document.getElementById('defectsNextPageBtn');
    const pageNumbers = document.getElementById('defectsPageNumbers');
    
    if (!paginationInfo || !prevBtn || !nextBtn || !pageNumbers) return;
    
    // Calcular rango actual
    const startIndex = (defectsCurrentPage - 1) * defectsItemsPerPage + 1;
    const endIndex = Math.min(defectsCurrentPage * defectsItemsPerPage, totalItems);
    
    // Actualizar información
    if (totalItems === 0) {
        paginationInfo.textContent = 'No hay registros para mostrar';
    } else {
        paginationInfo.textContent = `Mostrando ${startIndex}-${endIndex} de ${totalItems} registros`;
    }
    
    // Botones anterior/siguiente
    prevBtn.disabled = defectsCurrentPage === 1;
    nextBtn.disabled = defectsCurrentPage === totalPages || totalPages === 0;
    
    // Generar números de página
    generateDefectsPageNumbers(pageNumbers, defectsCurrentPage, totalPages);
}

// Generar números de página para defectos
function generateDefectsPageNumbers(container, current, total) {
    container.innerHTML = '';
    
    if (total <= 1) return;
    
    const maxVisiblePages = 5;
    let startPage = Math.max(1, current - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(total, startPage + maxVisiblePages - 1);
    
    // Ajustar si no hay suficientes páginas al final
    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }
    
    // Botón primera página
    if (startPage > 1) {
        createDefectsPageButton(container, 1, '1');
        if (startPage > 2) {
            createDefectsPageButton(container, null, '...');
        }
    }
    
    // Páginas visibles
    for (let i = startPage; i <= endPage; i++) {
        createDefectsPageButton(container, i, i.toString());
    }
    
    // Botón última página
    if (endPage < total) {
        if (endPage < total - 1) {
            createDefectsPageButton(container, null, '...');
        }
        createDefectsPageButton(container, total, total.toString());
    }
}

// Crear botón de página para defectos
function createDefectsPageButton(container, pageNum, text) {
    const button = document.createElement('button');
    button.className = 'page-number';
    button.textContent = text;
    
    if (pageNum === null) {
        button.classList.add('ellipsis');
    } else {
        if (pageNum === defectsCurrentPage) {
            button.classList.add('active');
        }
        button.addEventListener('click', () => goToDefectsPage(pageNum));
    }
    
    container.appendChild(button);
}

// Ir a una página específica de defectos
function goToDefectsPage(page) {
    defectsCurrentPage = page;
    populateDefectsTable();
}

// Ir a página anterior de defectos
function goToDefectsPrevPage() {
    if (defectsCurrentPage > 1) {
        defectsCurrentPage--;
        populateDefectsTable();
    }
}

// Ir a página siguiente de defectos
function goToDefectsNextPage() {
    const totalPages = defectsItemsPerPage === 100 ? 1 : Math.ceil(filteredDefects.length / defectsItemsPerPage);
    if (defectsCurrentPage < totalPages) {
        defectsCurrentPage++;
        populateDefectsTable();
    }
}

// Cambiar elementos por página para defectos
function changeDefectsItemsPerPage(newItemsPerPage) {
    defectsItemsPerPage = parseInt(newItemsPerPage);
    defectsCurrentPage = 1; // Resetear a primera página
    populateDefectsTable();
}

// Obtener texto del estado
function getStatusText(status) {
    switch(status) {
        case 'success': return 'Exitosa';
        case 'failure': return 'Fallida';
        case 'pending': return 'Pendiente';
        case 'blocked': return 'Bloqueada';
        case 'planned': return 'Planificada';
        case 'dismissed': return 'Desestimado';
        default: return status;
    }
}

// Función para formatear fechas (solo fecha, sin hora)
function formatFecha(fechaStr) {
    console.log('formatFecha recibe:', fechaStr, typeof fechaStr);
    
    if (!fechaStr || fechaStr === '') return '-';
    
    // Si ya está en formato DD/MM/AAAA, devolverlo tal como está
    if (typeof fechaStr === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(fechaStr)) {
        const [dia, mes, año] = fechaStr.split('/');
        return `${dia.padStart(2, '0')}/${mes.padStart(2, '0')}/${año}`;
    }
    
    // Si viene como string vacío o guion, mantener
    if (fechaStr === '-' || fechaStr.trim() === '') return '-';
    
    // Para cualquier otro caso, intentar usar formatDateForDashboard
    try {
        const fechaProcessada = formatDateForDashboard(fechaStr);
        return fechaProcessada || '-';
    } catch (error) {
        console.warn('Error formateando fecha:', fechaStr, error);
        return '-';
    }
}

// Poblar tabla de defectos
function populateDefectsTable() {
    const tableBody = document.getElementById('defectsTableBody');
    if (!tableBody) return;
    
    // Verificar si hay datos de defectos
    if (!testData.defects || !testData.defects.details || testData.defects.details.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="8" class="no-data-message">
                    <div class="no-data-content">
                        <i class="fas fa-bug"></i>
                        <p>No hay defectos registrados</p>
                        <small>Los defectos aparecerán aquí cuando se carguen datos</small>
                    </div>
                </td>
            </tr>
        `;
        updateDefectsPaginationControls(0, 0);
        return;
    }
    
    // Si no hay filtro aplicado, usar todos los defectos
    if (filteredDefects.length === 0) {
        // Verificar si realmente no hay filtros aplicados
        const searchTerm = document.getElementById('defectsSearchInput')?.value || '';
        const severityFilter = document.getElementById('severityFilter')?.value || '';
        const statusFilter = document.getElementById('statusFilter')?.value || '';
        
        const hasActiveFilters = searchTerm !== '' || severityFilter !== '' || statusFilter !== '';
        
        if (!hasActiveFilters) {
            filteredDefects = [...testData.defects.details];
        }
    }
    
    const totalItems = filteredDefects.length;
    const totalPages = defectsItemsPerPage === 100 ? 1 : Math.ceil(totalItems / defectsItemsPerPage);
    
    // Limpiar tabla
    tableBody.innerHTML = '';
    
    // Si no hay resultados después del filtrado, mostrar mensaje
    if (totalItems === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="8" class="no-results-message">
                    <div class="no-results-content">
                        <i class="fas fa-search"></i>
                        <p>No se encontraron defectos que coincidan con los filtros</p>
                        <small>Intente modificar o limpiar los filtros para ver más resultados</small>
                    </div>
                </td>
            </tr>
        `;
        updateDefectsPaginationControls(0, 0);
        return;
    }
    
    // Calcular índices para la página actual
    const startIndex = (defectsCurrentPage - 1) * defectsItemsPerPage;
    const endIndex = defectsItemsPerPage === 100 ? totalItems : Math.min(startIndex + defectsItemsPerPage, totalItems);
    
    // Mostrar defectos de la página actual
    const defectsToShow = filteredDefects.slice(startIndex, endIndex);
    
    // Obtener término de búsqueda actual para resaltado
    const searchInput = document.getElementById('defectsSearchInput');
    const currentSearchTerm = searchInput ? searchInput.value.trim() : '';
    
    defectsToShow.forEach(defect => {
        const row = document.createElement('tr');
        const daysOpen = calculateDaysOpen(defect.dateFound, defect.dateResolved);
        const severityClass = getSeverityClass(defect.severity);
        const statusClass = getStatusClass(defect.status);
        const daysOpenClass = getDaysOpenClass(daysOpen, defect.status);
        
        // Aplicar resaltado si hay término de búsqueda
        const highlightedId = highlightSearchTerm(defect.id, currentSearchTerm);
        const highlightedTitle = highlightSearchTerm(defect.title, currentSearchTerm);
        const highlightedAssignee = highlightSearchTerm(defect.assignee, currentSearchTerm);
        const highlightedCpImpactados = highlightSearchTerm(String(defect.cp_impactados || ''), currentSearchTerm);
        
        row.innerHTML = `
            <td><strong>${highlightedId}</strong></td>
            <td>${highlightedTitle}</td>
            <td><span class="defect-severity ${severityClass}">${defect.severity}</span></td>
            <td><span class="defect-status ${statusClass}">${defect.status}</span></td>
            <td>${highlightedCpImpactados}</td>
            <td>${highlightedAssignee}</td>
            <td>${formatDate(defect.dateFound)}</td>
            <td><span class="days-open ${daysOpenClass}">${daysOpen} días</span></td>
        `;
        tableBody.appendChild(row);
    });
    
    // Actualizar controles de paginación
    updateDefectsPaginationControls(totalItems, totalPages);
}

// Calcular días abierto (solo días laborales - lunes a viernes)
function calculateDaysOpen(dateFound, dateResolved = null) {
    // Si hay fecha de resolución, calcular entre fecha_encontrado y fecha_resolucion
    // Si no hay fecha de resolución, calcular desde fecha_encontrado hasta hoy
    const endDate = dateResolved && dateResolved.trim() !== '' ? parseDate(dateResolved) : new Date();
    
    // Convertir fecha encontrada a objeto Date
    const foundDate = parseDate(dateFound);
    
    if (isNaN(foundDate.getTime())) {
        console.warn('Fecha de encontrado inválida en calculateDaysOpen:', dateFound);
        return 0;
    }
    
    if (dateResolved && isNaN(endDate.getTime())) {
        console.warn('Fecha de resolución inválida en calculateDaysOpen:', dateResolved);
        // Si la fecha de resolución es inválida, usar fecha actual
        return calculateBusinessDays(foundDate, new Date());
    }
    
    // Calcular solo días laborales (lunes a viernes)
    const businessDays = calculateBusinessDays(foundDate, endDate);
    
    const endDateStr = dateResolved ? dateResolved : 'HOY';
    console.log(`Calculando días abiertos laborales: Fecha encontrado=${dateFound}, Fecha resolución=${endDateStr}, Días laborales=${businessDays}`);
    console.log(`💼 SOLO DÍAS LABORALES: Excluyendo sábados y domingos del cálculo`);
    
    return businessDays;
}

// Función auxiliar para parsear fechas en diferentes formatos
function parseDate(dateString) {
    if (!dateString) return new Date();
    
    // Convertir fecha a objeto Date
    let parsedDate;
    
    if (typeof dateString === 'string') {
        // Formato DD/MM/AAAA HH:MM
        if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}$/)) {
            const [datePart, timePart] = dateString.split(/\s+/);
            const [day, month, year] = datePart.split('/').map(Number);
            const [hour, minute] = timePart.split(':').map(Number);
            parsedDate = new Date(year, month - 1, day, hour, minute);
        }
        // Formato DD/MM/AAAA (solo fecha)
        else if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
            const [day, month, year] = dateString.split('/').map(Number);
            parsedDate = new Date(year, month - 1, day);
        }
        // Formato YYYY-MM-DD HH:MM
        else if (dateString.match(/^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}$/)) {
            const [datePart, timePart] = dateString.split(/\s+/);
            const [year, month, day] = datePart.split('-').map(Number);
            const [hour, minute] = timePart.split(':').map(Number);
            parsedDate = new Date(year, month - 1, day, hour, minute);
        }
        // Formato YYYY-MM-DD (solo fecha)
        else if (dateString.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
            const [year, month, day] = dateString.split('-').map(Number);
            parsedDate = new Date(year, month - 1, day);
        }
        else {
            // Intentar parsear directamente como último recurso
            parsedDate = new Date(dateString);
        }
    } else {
        // Si ya es un objeto Date o número
        parsedDate = new Date(dateString);
    }
    
    return parsedDate;
}

// Función auxiliar para calcular días laborales entre dos fechas (excluyendo sábados y domingos)
function calculateBusinessDays(startDate, endDate) {
    // Asegurar que las fechas estén en el orden correcto
    let start = new Date(startDate);
    let end = new Date(endDate);
    
    if (start > end) {
        [start, end] = [end, start];
    }
    
    // Ajustar fechas a medianoche para contar días completos
    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);
    
    let businessDays = 0;
    let currentDate = new Date(start);
    
    console.log(`📅 Calculando días laborales desde ${start.toLocaleDateString('es-ES')} hasta ${end.toLocaleDateString('es-ES')}`);
    
    while (currentDate <= end) {
        const dayOfWeek = currentDate.getDay(); // 0=Domingo, 1=Lunes, ..., 6=Sábado
        const dayName = currentDate.toLocaleDateString('es-ES', { weekday: 'long' });
        
        // Solo contar si es día laboral (lunes=1 a viernes=5)
        if (dayOfWeek >= 1 && dayOfWeek <= 5) {
            businessDays++;
            console.log(`  ✅ ${dayName} ${currentDate.toLocaleDateString('es-ES')} - DÍA LABORAL (Total: ${businessDays})`);
        } else {
            console.log(`  ❌ ${dayName} ${currentDate.toLocaleDateString('es-ES')} - FIN DE SEMANA (Excluido)`);
        }
        
        // Avanzar al siguiente día
        currentDate.setDate(currentDate.getDate() + 1);
    }
    
    console.log(`📊 TOTAL DÍAS LABORALES: ${businessDays}`);
    return businessDays;
}

// Obtener clase CSS de severidad
function getSeverityClass(severity) {
    switch(severity) {
        case 'Crítica': return 'critical';
        case 'Alta': return 'high';
        case 'Media': return 'medium';
        case 'Baja': return 'low';
        default: return 'medium';
    }
}

// Obtener clase CSS de estado
function getStatusClass(status) {
    switch(status) {
        case 'Open': return 'open';
        case 'Reject': return 'reject';
        case 'Assigned': return 'assigned';
        case 'Under Review': return 'under-review';
        case 'In Progress DEV': return 'in-progress-dev';
        case 'Resolved': return 'resolved';
        case 'ReTesting': return 'retesting';
        case 'ReOpened': return 'reopened';
        case 'Closed': return 'closed';
        default: return 'open';
    }
}

// Obtener clase CSS para días abierto
function getDaysOpenClass(days, status) {
    // Estados que se consideran cerrados/terminados
    if (status === 'Closed' || status === 'Resolved' || status === 'Reject') return 'normal';
    if (days > 30) return 'critical';
    if (days > 14) return 'warning';
    return 'normal';
}

// Formatear fecha
function formatDate(dateString) {
    if (!dateString) return '';
    
    let date;
    
    // Si está en formato DD/MM/AAAA HH:MM, extraer solo la fecha
    if (typeof dateString === 'string' && dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}$/)) {
        const datePart = dateString.split(' ')[0]; // Tomar solo la parte de fecha
        return datePart; // Ya está en formato DD/MM/AAAA
    }
    
    // Si ya está en formato DD/MM/AAAA, retornarlo tal como está
    if (typeof dateString === 'string' && dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        return dateString;
    }
    
    // Si está en formato YYYY-MM-DD, convertirlo
    if (typeof dateString === 'string' && dateString.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const [year, month, day] = dateString.split('-');
        return `${day}/${month}/${year}`;
    }
    
    // Si está en formato YYYY-MM-DD HH:MM, convertir solo la fecha
    if (typeof dateString === 'string' && dateString.match(/^\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}$/)) {
        const datePart = dateString.split(' ')[0]; // YYYY-MM-DD
        const [year, month, day] = datePart.split('-');
        return `${day}/${month}/${year}`;
    }
    
    // Intentar convertir si es un objeto Date o string parseeable
    try {
        date = new Date(dateString);
        if (isNaN(date.getTime())) {
            return dateString; // Si no se puede parsear, retornar el original
        }
        
        // Formatear como DD/MM/AAAA
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        
        return `${day}/${month}/${year}`;
    } catch (error) {
        console.warn('Error formateando fecha:', dateString, error);
        return dateString;
    }
}

// Filtrar defectos
function filterDefects() {
    const searchTerm = document.getElementById('defectsSearchInput')?.value.toLowerCase() || '';
    const severityFilter = document.getElementById('severityFilter')?.value || '';
    const statusFilter = document.getElementById('statusFilter')?.value || '';
    
    if (!testData.defects || !testData.defects.details) {
        filteredDefects = [];
        defectsCurrentPage = 1;
        populateDefectsTable();
        updateDefectsFilterInfo('', 0, 0);
        return;
    }
    
    const totalDefects = testData.defects.details.length;
    
    // Filtrar defectos basado en los criterios de búsqueda
    if (searchTerm === '' && severityFilter === '' && statusFilter === '') {
        filteredDefects = [...testData.defects.details];
    } else {
        filteredDefects = testData.defects.details.filter(defect => {
            const matchesSearch = defect.title.toLowerCase().includes(searchTerm) ||
                                defect.id.toLowerCase().includes(searchTerm) ||
                                defect.assignee.toLowerCase().includes(searchTerm) ||
                                (defect.cp_impactados || '').toLowerCase().includes(searchTerm);
            const matchesSeverity = !severityFilter || defect.severity === severityFilter;
            const matchesStatus = !statusFilter || defect.status === statusFilter;
            
            return matchesSearch && matchesSeverity && matchesStatus;
        });
    }
    
    // Actualizar información de filtrado
    const hasActiveFilters = searchTerm !== '' || severityFilter !== '' || statusFilter !== '';
    const filterDescription = buildFilterDescription(searchTerm, severityFilter, statusFilter);
    updateDefectsFilterInfo(filterDescription, filteredDefects.length, totalDefects, hasActiveFilters);
    
    // Resetear a la primera página y actualizar la tabla
    defectsCurrentPage = 1;
    populateDefectsTable();
}

// Construir descripción del filtro activo
function buildFilterDescription(searchTerm, severityFilter, statusFilter) {
    const filters = [];
    
    if (searchTerm) {
        filters.push(`texto: "${searchTerm}"`);
    }
    if (severityFilter) {
        filters.push(`severidad: ${severityFilter}`);
    }
    if (statusFilter) {
        filters.push(`estado: ${statusFilter}`);
    }
    
    return filters.length > 0 ? `Filtrado por ${filters.join(', ')}` : '';
}

// Actualizar información de filtrado de defectos
function updateDefectsFilterInfo(filterDescription, foundResults, totalResults, hasActiveFilters = false) {
    // Buscar si existe un elemento para mostrar info de filtrado
    let filterInfo = document.getElementById('defectsFilterInfo');
    
    if (!filterInfo) {
        // Crear elemento si no existe
        filterInfo = document.createElement('div');
        filterInfo.id = 'defectsFilterInfo';
        filterInfo.className = 'filter-results-info';
        
        // Insertar después de los controles de filtrado
        const tableHeader = document.querySelector('.defects-details .table-header');
        if (tableHeader) {
            const tableContainer = tableHeader.nextElementSibling;
            if (tableContainer) {
                tableHeader.parentNode.insertBefore(filterInfo, tableContainer);
            }
        }
    }
    
    // Actualizar contenido
    if (!hasActiveFilters) {
        filterInfo.style.display = 'none';
    } else {
        filterInfo.style.display = 'block';
        if (foundResults === 0) {
            filterInfo.innerHTML = `
                <i class="fas fa-exclamation-triangle"></i> 
                No se encontraron defectos. ${filterDescription}
                <button class="btn-clear-filters" onclick="clearDefectsFilters()">
                    <i class="fas fa-times"></i> Limpiar filtros
                </button>
            `;
            filterInfo.className = 'filter-results-info no-results';
        } else {
            filterInfo.innerHTML = `
                <i class="fas fa-filter"></i> 
                ${foundResults} de ${totalResults} defecto(s) mostrado(s). ${filterDescription}
                <button class="btn-clear-filters" onclick="clearDefectsFilters()">
                    <i class="fas fa-times"></i> Limpiar filtros
                </button>
            `;
            filterInfo.className = 'filter-results-info with-results';
        }
    }
    
    // Log en consola
    if (hasActiveFilters) {
        console.log(`🔍 Filtro de defectos: ${foundResults}/${totalResults} - ${filterDescription}`);
    }
}

// Limpiar filtros de defectos
function clearDefectsFilters() {
    // Limpiar todos los filtros
    const searchInput = document.getElementById('defectsSearchInput');
    const severityFilter = document.getElementById('severityFilter');
    const statusFilter = document.getElementById('statusFilter');
    
    if (searchInput) searchInput.value = '';
    if (severityFilter) severityFilter.value = '';
    if (statusFilter) statusFilter.value = '';
    
    // Volver a filtrar (que mostrará todos los defectos)
    filterDefects();
}

// Exportar defectos
function exportDefects() {
    if (!testData.defects || !testData.defects.details) return;
    
    const csvContent = [
        ['ID', 'Título', 'Severidad', 'Estado', 'Escenario', 'Asignado a', 'Fecha Encontrado', 'Días Abierto'],
        ...testData.defects.details.map(defect => [
            defect.id,
            defect.title,
            defect.severity,
            defect.status,
            defect.escenario,
            defect.assignee,
            defect.dateFound,
            calculateDaysOpen(defect.dateFound, defect.dateResolved)
        ])
    ].map(row => row.map(field => `"${field}"`).join(',')).join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `defectos_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
}

// Configurar event listeners
function setupEventListeners() {
    // Filtros de tiempo para el gráfico de pastel
    const filterButtons = document.querySelectorAll('.btn-filter');
    filterButtons.forEach(button => {
        button.addEventListener('click', function() {
            filterButtons.forEach(btn => btn.classList.remove('active'));
            this.classList.add('active');
            // Aquí podrías actualizar los datos según el filtro seleccionado
        });
    });

    // Búsqueda en la tabla de pruebas
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', function() {
            filterTable(this.value);
        });
    }

    // Filtro por estado en la tabla de pruebas
    const testStatusFilter = document.getElementById('testStatusFilter');
    if (testStatusFilter) {
        testStatusFilter.addEventListener('change', function() {
            filterTable();
        });
    }

    // Búsqueda y filtros en la tabla de defectos
    const defectsSearchInput = document.getElementById('defectsSearchInput');
    if (defectsSearchInput) {
        defectsSearchInput.addEventListener('input', filterDefects);
    }

    const severityFilter = document.getElementById('severityFilter');
    if (severityFilter) {
        severityFilter.addEventListener('change', filterDefects);
    }

    const statusFilter = document.getElementById('statusFilter');
    if (statusFilter) {
        statusFilter.addEventListener('change', filterDefects);
    }

    // Exportar datos
    const exportButton = document.querySelector('.btn-export');
    if (exportButton) {
        exportButton.addEventListener('click', exportData);
    }

    // Event listeners para paginación
    const prevPageBtn = document.getElementById('prevPageBtn');
    if (prevPageBtn) {
        prevPageBtn.addEventListener('click', goToPrevPage);
    }

    const nextPageBtn = document.getElementById('nextPageBtn');
    if (nextPageBtn) {
        nextPageBtn.addEventListener('click', goToNextPage);
    }

    const itemsPerPageSelect = document.getElementById('itemsPerPageSelect');
    if (itemsPerPageSelect) {
        itemsPerPageSelect.addEventListener('change', function() {
            changeItemsPerPage(this.value);
        });
    }

    // Event listeners para paginación de defectos
    const defectsPrevPageBtn = document.getElementById('defectsPrevPageBtn');
    if (defectsPrevPageBtn) {
        defectsPrevPageBtn.addEventListener('click', goToDefectsPrevPage);
    }

    const defectsNextPageBtn = document.getElementById('defectsNextPageBtn');
    if (defectsNextPageBtn) {
        defectsNextPageBtn.addEventListener('click', goToDefectsNextPage);
    }

    const defectsItemsPerPageSelect = document.getElementById('defectsItemsPerPageSelect');
    if (defectsItemsPerPageSelect) {
        defectsItemsPerPageSelect.addEventListener('change', function() {
            changeDefectsItemsPerPage(this.value);
        });
    }

    // Auto-actualización cada 5 minutos
    setInterval(updateDashboard, 300000);
}

// Filtrar tabla de pruebas con búsqueda y filtro por estado
function filterTable(searchTerm = null) {
    // Obtener valores de filtros
    const searchInput = document.getElementById('searchInput');
    const statusFilter = document.getElementById('testStatusFilter');
    
    const searchTermValue = searchTerm !== null ? searchTerm : (searchInput ? searchInput.value : '');
    const searchTermLower = searchTermValue.toLowerCase().trim();
    const statusFilterValue = statusFilter ? statusFilter.value : '';
    
    // Verificar si hay filtros activos
    const hasActiveFilters = searchTermLower !== '' || statusFilterValue !== '';
    
    // Filtrar los tests basado en búsqueda y filtro de estado
    if (!hasActiveFilters) {
        filteredTests = [...testData.testDetails];
    } else {
        filteredTests = testData.testDetails.filter(test => {
            // Filtro de búsqueda por texto
            let matchesSearch = true;
            if (searchTermLower !== '') {
                const searchableFields = [
                    test.id,                              // ID de la prueba
                    test.name,                            // Nombre de la prueba
                    test.escenario,                       // Escenario
                    getStatusText(test.status),           // Estado (traducido)
                    test.status,                          // Estado original
                    test.executor,                        // Ejecutor
                    test.dia_ejecutado,                   // Fecha de ejecución
                    formatFecha(test.dia_ejecutado)       // Fecha formateada
                ];
                
                matchesSearch = searchableFields.some(field => {
                    return field && field.toString().toLowerCase().includes(searchTermLower);
                });
            }
            
            // Filtro por estado
            let matchesStatus = true;
            if (statusFilterValue !== '') {
                matchesStatus = test.status === statusFilterValue;
            }
            
            return matchesSearch && matchesStatus;
        });
    }
    
    // Actualizar información de filtrado
    const filterDescription = buildTestFilterDescription(searchTermLower, statusFilterValue);
    updateTestFilterInfo(filterDescription, filteredTests.length, testData.testDetails.length, hasActiveFilters);
    
    // Resetear a la primera página y actualizar la tabla
    currentPage = 1;
    populateTestTable();
}

// Construir descripción del filtro activo para pruebas
function buildTestFilterDescription(searchTerm, statusFilter) {
    const filters = [];
    
    if (searchTerm) {
        filters.push(`texto: "${searchTerm}"`);
    }
    if (statusFilter) {
        const statusText = getStatusText(statusFilter);
        filters.push(`estado: ${statusText}`);
    }
    
    return filters.length > 0 ? `Filtrado por ${filters.join(', ')}` : '';
}

// Actualizar información de filtrado de pruebas
function updateTestFilterInfo(filterDescription, foundResults, totalResults, hasActiveFilters = false) {
    // Buscar si existe un elemento para mostrar info de filtrado
    let filterInfo = document.getElementById('testFilterInfo');
    
    if (!filterInfo) {
        // Crear elemento si no existe
        filterInfo = document.createElement('div');
        filterInfo.id = 'testFilterInfo';
        filterInfo.className = 'filter-results-info';
        
        // Insertar después de los controles de filtrado
        const tableHeader = document.querySelector('section.table-section:not(.defects-details) .table-header');
        if (tableHeader) {
            const tableContainer = tableHeader.nextElementSibling;
            if (tableContainer) {
                tableHeader.parentNode.insertBefore(filterInfo, tableContainer);
            }
        }
    }
    
    // Actualizar contenido
    if (!hasActiveFilters) {
        filterInfo.style.display = 'none';
    } else {
        filterInfo.style.display = 'block';
        if (foundResults === 0) {
            filterInfo.innerHTML = `
                <i class="fas fa-exclamation-triangle"></i> 
                No se encontraron pruebas. ${filterDescription}
                <button class="btn-clear-filters" onclick="clearTestFilters()">
                    <i class="fas fa-times"></i> Limpiar filtros
                </button>
            `;
            filterInfo.className = 'filter-results-info no-results';
        } else {
            filterInfo.innerHTML = `
                <i class="fas fa-filter"></i> 
                ${foundResults} de ${totalResults} prueba(s) mostrada(s). ${filterDescription}
                <button class="btn-clear-filters" onclick="clearTestFilters()">
                    <i class="fas fa-times"></i> Limpiar filtros
                </button>
            `;
            filterInfo.className = 'filter-results-info with-results';
        }
    }
    
    // Log en consola
    if (hasActiveFilters) {
        console.log(`🔍 Filtro de pruebas: ${foundResults}/${totalResults} - ${filterDescription}`);
    }
}

// Limpiar filtros de pruebas
function clearTestFilters() {
    // Limpiar todos los filtros
    const searchInput = document.getElementById('searchInput');
    const statusFilter = document.getElementById('testStatusFilter');
    
    if (searchInput) searchInput.value = '';
    if (statusFilter) statusFilter.value = '';
    
    // Volver a filtrar (que mostrará todas las pruebas)
    filterTable();
}

// Actualizar información de resultados de búsqueda
function updateSearchResultsInfo(searchTerm, foundResults, totalResults) {
    // Buscar si existe un elemento para mostrar info de búsqueda
    let searchInfo = document.getElementById('searchResultsInfo');
    
    if (!searchInfo) {
        // Crear elemento si no existe
        searchInfo = document.createElement('div');
        searchInfo.id = 'searchResultsInfo';
        searchInfo.className = 'search-results-info';
        
        // Insertar después del input de búsqueda
        const searchBox = document.querySelector('.search-box');
        if (searchBox) {
            searchBox.appendChild(searchInfo);
        }
    }
    
    // Actualizar contenido
    if (searchTerm.trim() === '') {
        searchInfo.textContent = '';
        searchInfo.style.display = 'none';
    } else {
        searchInfo.style.display = 'block';
        if (foundResults === 0) {
            searchInfo.innerHTML = `<i class="fas fa-exclamation-triangle"></i> No se encontraron resultados para "${searchTerm}"`;
            searchInfo.style.color = '#dc3545';
        } else {
            searchInfo.innerHTML = `<i class="fas fa-check-circle"></i> ${foundResults} de ${totalResults} prueba(s) encontrada(s)`;
            searchInfo.style.color = '#28a745';
        }
    }
    
    // Log en consola
    if (searchTerm !== '' && foundResults === 0) {
        console.log(`🔍 Búsqueda "${searchTerm}": No se encontraron resultados`);
    } else if (searchTerm !== '') {
        console.log(`🔍 Búsqueda "${searchTerm}": ${foundResults} resultado(s) encontrado(s)`);
    }
}

// Exportar datos
function exportData() {
    const data = testData.testDetails.map(test => ({
        ID: test.id,
        'Nombre de la Prueba': test.name,
        'Escenario': test.escenario,
        'Estado': getStatusText(test.status),
        'Fecha Ejecución': formatFecha(test.dia_ejecutado),
        'Ejecutor': test.executor
    }));

    const csv = convertToCSV(data);
    downloadCSV(csv, `reporte_qa_${new Date().toISOString().split('T')[0]}.csv`);
}

// Convertir a CSV
function convertToCSV(data) {
    const headers = Object.keys(data[0]).join(',');
    const rows = data.map(row => Object.values(row).join(',')).join('\n');
    return headers + '\n' + rows;
}

// Descargar CSV
function downloadCSV(csv, filename) {
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    window.URL.revokeObjectURL(url);
}

// Actualizar dashboard (simular datos en tiempo real)
function updateDashboard() {
    // Simular cambios menores en los datos
    testData.summary.planned += Math.floor(Math.random() * 5) - 2;
    testData.summary.successful += Math.floor(Math.random() * 3) - 1;
    testData.summary.failed += Math.floor(Math.random() * 2) - 1;
    testData.summary.pending += Math.floor(Math.random() * 2) - 1;
    
    // Asegurar que los números no sean negativos
    testData.summary.planned = Math.max(50, testData.summary.planned);
    testData.summary.successful = Math.max(0, testData.summary.successful);
    testData.summary.failed = Math.max(0, testData.summary.failed);
    testData.summary.pending = Math.max(0, testData.summary.pending);
    
    // Actualizar la interfaz
    initializeDashboard();
    createPieChart();
}

// Hacer las gráficas responsivas cuando cambie el tamaño de ventana
window.addEventListener('resize', function() {
    Plotly.Plots.resize('pieChart');
    Plotly.Plots.resize('trendChart');
    Plotly.Plots.resize('categoryChart');
    Plotly.Plots.resize('defectsChart');
    Plotly.Plots.resize('burndownChart');
    Plotly.Plots.resize('coverageChart');
    Plotly.Plots.resize('dailyExecutionChart');
    Plotly.Plots.resize('executorDistributionChart');
});

// Configurar carga de archivos
function setupFileUpload() {
    const fileInput = document.getElementById('excelFile');
    if (fileInput) {
        fileInput.addEventListener('change', handleFileUpload);
    }
}

// Manejar carga de archivos
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Limpiar estado anterior antes de cargar nuevo archivo
    resetDashboardState();

    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.csv')) {
        handleCSVFile(file);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        handleExcelFile(file);
    } else {
        showNotification('Formato de archivo no soportado. Use .xlsx, .xls o .csv', 'error');
        hideLoadingState(); // Ocultar loading en caso de error
        // Limpiar el input del archivo
        event.target.value = '';
        return;
    }
}

// Función para resetear el estado del dashboard
function resetDashboardState() {
    console.log('🔄 Limpiando estado del dashboard para nueva carga...');
    
    // Mostrar indicador de carga
    showLoadingState();
    
    // Limpiar TODA la memoria del objeto testData
    Object.keys(testData).forEach(key => {
        if (Array.isArray(testData[key])) {
            testData[key] = [];
        } else if (typeof testData[key] === 'object') {
            testData[key] = {};
        }
    });
    
    // Forzar limpieza de variables globales
    if (typeof filteredTests !== 'undefined') filteredTests = [];
    if (typeof currentPage !== 'undefined') currentPage = 1;
    if (typeof filteredDefects !== 'undefined') filteredDefects = [];
    if (typeof defectsCurrentPage !== 'undefined') defectsCurrentPage = 1;
    
    // Resetear variables de paginación
    filteredTests = [];
    currentPage = 1;
    
    // Resetear variables de paginación de defectos
    filteredDefects = [];
    defectsCurrentPage = 1;
    
    // Limpiar selectores de proyectos activos si existen
    const existingModal = document.getElementById('projectSelectorModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Resetear valores de selector de items por página
    const itemsPerPageSelect = document.getElementById('itemsPerPageSelect');
    if (itemsPerPageSelect) {
        itemsPerPageSelect.value = '5';
        itemsPerPage = 5;
    }
    
    // Resetear valores de selector de items por página para defectos
    const defectsItemsPerPageSelect = document.getElementById('defectsItemsPerPageSelect');
    if (defectsItemsPerPageSelect) {
        defectsItemsPerPageSelect.value = '5';
        defectsItemsPerPage = 5;
    }
    
    // Limpiar campo de búsqueda
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.value = '';
    }
    
    // Limpiar filtros de defectos
    const defectsSearchInput = document.getElementById('defectsSearchInput');
    if (defectsSearchInput) {
        defectsSearchInput.value = '';
    }
    
    // Limpiar gráficos existentes para evitar solapamiento visual
    clearAllCharts();
    
    // Limpiar localStorage para evitar datos cached
    localStorage.removeItem('qaEditorData');
    
    // Limpiar filtros adicionales
    const severityFilter = document.getElementById('severityFilter');
    if (severityFilter) {
        severityFilter.value = '';
    }
    
    console.log('✅ Estado del dashboard limpiado correctamente');
}

// Función para mostrar estado de carga
function showLoadingState() {
    // Agregar overlay de carga si no existe
    let loadingOverlay = document.getElementById('loadingOverlay');
    if (!loadingOverlay) {
        loadingOverlay = document.createElement('div');
        loadingOverlay.id = 'loadingOverlay';
        loadingOverlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255, 255, 255, 0.9);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            font-family: Inter, sans-serif;
        `;
        loadingOverlay.innerHTML = `
            <div style="text-align: center;">
                <div style="border: 4px solid #f3f3f3; border-top: 4px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px;"></div>
                <p style="font-size: 16px; color: #333;">Cargando nuevo archivo...</p>
            </div>
        `;
        document.body.appendChild(loadingOverlay);
        
        // Agregar animación CSS si no existe
        if (!document.getElementById('spinAnimation')) {
            const style = document.createElement('style');
            style.id = 'spinAnimation';
            style.textContent = `
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            `;
            document.head.appendChild(style);
        }
    }
    loadingOverlay.style.display = 'flex';
}

// Función para ocultar estado de carga
function hideLoadingState() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.style.display = 'none';
    }
}

// Función para limpiar todos los gráficos
function clearAllCharts() {
    const chartIds = ['pieChart', 'trendChart', 'categoryChart', 'defectsChart', 'coverageChart', 'dailyExecutionChart', 'executorDistributionChart'];
    chartIds.forEach(chartId => {
        const chartElement = document.getElementById(chartId);
        if (chartElement) {
            // Limpiar contenido del gráfico
            chartElement.innerHTML = '';
        }
    });
}

// Manejar archivos CSV
function handleCSVFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const csvData = e.target.result;
            const parsedData = parseCSV(csvData);
            
            if (isTestDetailsCSV(parsedData)) {
                updateDashboardWithCSVData(parsedData, 'details');
                showNotification('Datos de detalle de pruebas cargados exitosamente', 'success');
            } else if (isSummaryCSV(parsedData)) {
                updateDashboardWithCSVData(parsedData, 'summary');
                showNotification('Datos de resumen diario cargados exitosamente', 'success');
            } else if (isDefectsCSV(parsedData)) {
                updateDashboardWithCSVData(parsedData, 'defects');
                showNotification('Datos de defectos cargados exitosamente', 'success');
            } else {
                showNotification('Formato CSV no reconocido. Verifique las columnas.', 'error');
            }
        } catch (error) {
            console.error('Error procesando CSV:', error);
            showNotification('Error al procesar el archivo CSV: ' + error.message, 'error');
            // Resetear el input del archivo en caso de error
            const fileInput = document.getElementById('excelFile');
            if (fileInput) {
                fileInput.value = '';
            }
        }
    };
    
    reader.onerror = function() {
        showNotification('Error al leer el archivo CSV', 'error');
        const fileInput = document.getElementById('excelFile');
        if (fileInput) {
            fileInput.value = '';
        }
    };
    
    reader.readAsText(file);
}

// Manejar archivos Excel
function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Verificar si es formato multi-proyecto (sin necesidad de hojas Resumen_Pruebas ni Categorias_Pruebas)
            const expectedSheets = ['Proyectos', 'Tendencia_Historica', 'Detalle_Pruebas', 'Defectos'];
            const hasAllSheets = expectedSheets.every(sheet => workbook.SheetNames.includes(sheet));
            
            if (hasAllSheets) {
                handleMultiProjectExcel(workbook);
                showNotification('Archivo multi-proyecto cargado exitosamente', 'success');
                hideLoadingState(); // Ocultar loading después de carga exitosa
            } else {
                // Formato legacy - procesar hojas individuales
                handleLegacyExcel(workbook);
            }
            
        } catch (error) {
            console.error('Error procesando Excel:', error);
            showNotification('Error al procesar el archivo Excel: ' + error.message, 'error');
            hideLoadingState(); // Ocultar loading en caso de error
            // Resetear el input del archivo en caso de error
            const fileInput = document.getElementById('excelFile');
            if (fileInput) {
                fileInput.value = '';
            }
        }
    };
    
    reader.onerror = function() {
        showNotification('Error al leer el archivo', 'error');
        const fileInput = document.getElementById('excelFile');
        if (fileInput) {
            fileInput.value = '';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Manejar Excel multi-proyecto
function handleMultiProjectExcel(workbook) {
    try {
        // Leer todas las pestañas (sin Resumen_Pruebas ni Categorias_Pruebas, se calculan automáticamente)
        const proyectos = XLSX.utils.sheet_to_json(workbook.Sheets['Proyectos']);
        const tendenciaHistorica = XLSX.utils.sheet_to_json(workbook.Sheets['Tendencia_Historica']);
        const detallePruebas = XLSX.utils.sheet_to_json(workbook.Sheets['Detalle_Pruebas']);
        const defectos = XLSX.utils.sheet_to_json(workbook.Sheets['Defectos']);
        
        // Almacenar datos en variable global para uso posterior
        loadedExcelData = {
            proyectos,
            tendenciaHistorica,
            detallePruebas,
            defectos
        };
        
        // Permitir seleccionar proyecto
        showProjectSelector(proyectos, loadedExcelData);
        
    } catch (error) {
        console.error('Error procesando datos multi-proyecto:', error);
        showNotification('Error al procesar los datos del archivo', 'error');
    }
}

// Mostrar selector de proyecto
function showProjectSelector(proyectos, allData) {
    // Crear modal de selección
    const modal = document.createElement('div');
    modal.className = 'project-selector-modal';
    modal.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100%; height: 100%;
        background: rgba(0,0,0,0.5); z-index: 1000;
        display: flex; align-items: center; justify-content: center;
    `;
    
    const content = document.createElement('div');
    content.style.cssText = `
        background: white; padding: 2rem; border-radius: 12px; max-width: 500px; width: 90%;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
    `;
    
    content.innerHTML = `
        <h3 style="margin-bottom: 1rem; color: #1e293b;">Seleccionar Proyecto</h3>
        <p style="margin-bottom: 1.5rem; color: #64748b;">El archivo contiene ${proyectos.length} proyecto(s). Seleccione cuál desea cargar:</p>
        <div id="projectList" style="margin-bottom: 1.5rem; max-height: 300px; overflow-y: auto; border: 1px solid #e5e7eb; border-radius: 6px; padding: 0.5rem;"></div>
        <div style="display: flex; gap: 1rem; justify-content: flex-end;">
            <button id="cancelBtn" style="padding: 0.5rem 1rem; border: 1px solid #d1d5db; border-radius: 6px; background: white; color: #374151; cursor: pointer;">Cancelar</button>
        </div>
    `;
    
    const projectList = content.querySelector('#projectList');
    proyectos.forEach((proyecto, index) => {
        const projectDiv = document.createElement('div');
        projectDiv.style.cssText = `
            padding: 0.75rem; border: 1px solid #e5e7eb; border-radius: 6px; margin-bottom: 0.5rem; cursor: pointer;
            transition: all 0.2s ease; background: white;
        `;
        
        projectDiv.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div style="flex: 1;">
                    <div style="font-weight: bold; color: #1e293b; margin-bottom: 0.25rem;">${proyecto.nombre_proyecto}</div>
                    <div style="font-size: 0.875rem; color: #64748b;">
                        <span style="font-weight: 500;">ID:</span> ${proyecto.proyecto_id} | 
                        <span style="font-weight: 500;">Responsable:</span> ${proyecto.responsable_qa}
                    </div>
                </div>
            </div>
        `;
        
        projectDiv.addEventListener('click', () => {
            loadProjectData(proyecto.proyecto_id, allData);
            document.body.removeChild(modal);
        });
        
        projectDiv.addEventListener('mouseover', () => {
            projectDiv.style.borderColor = '#3b82f6';
            projectDiv.style.backgroundColor = '#f8fafc';
        });
        
        projectDiv.addEventListener('mouseout', () => {
            projectDiv.style.borderColor = '#e5e7eb';
            projectDiv.style.backgroundColor = 'white';
        });
        
        projectList.appendChild(projectDiv);
    });
    
    content.querySelector('#cancelBtn').addEventListener('click', () => {
        document.body.removeChild(modal);
    });
    
    modal.appendChild(content);
    document.body.appendChild(modal);
}

// Cargar datos del proyecto seleccionado
function loadProjectData(proyectoId, allData) {
    try {
        // Filtrar datos por proyecto
        const proyectoInfo = allData.proyectos.find(p => p.proyecto_id === proyectoId);
        const tendencia = allData.tendenciaHistorica.filter(t => t.proyecto_id === proyectoId);
        const pruebas = allData.detallePruebas.filter(p => p.proyecto_id === proyectoId);
        const defectos = allData.defectos.filter(d => d.proyecto_id === proyectoId);
        
        // DEBUG: Mostrar estructura de datos de Tendencia_Historica
        console.log('🔍 DEBUG TENDENCIA_HISTORICA DESDE EXCEL:');
        console.log('Total de registros de tendencia histórica:', allData.tendenciaHistorica.length);
        console.log('Registros filtrados para proyecto', proyectoId, ':', tendencia.length);
        if (tendencia.length > 0) {
            console.log('Estructura del primer registro de tendencia:', tendencia[0]);
            console.log('Todas las claves disponibles:', Object.keys(tendencia[0]));
        } else {
            console.log('⚠️ No se encontraron datos de tendencia para este proyecto');
        }
        
        // CALCULAR AUTOMÁTICAMENTE LOS RESÚMENES BASÁNDOSE EN DETALLE_PRUEBAS
        const resumenCalculado = calculateSummaryFromDetails(pruebas);
        const categoriasCalculadas = calculateCategoriesFromDetails(pruebas);
        
        // Transformar a formato del dashboard
        const dashboardData = transformMultiProjectData({
            proyecto: proyectoInfo,
            resumen: resumenCalculado,
            tendencia,
            categorias: categoriasCalculadas,
            pruebas,
            defectos
        });
        
        // Actualizar dashboard
        updateDashboardWithTransformedData(dashboardData);
        showNotification(`Proyecto "${proyectoInfo.nombre_proyecto}" cargado exitosamente`, 'success');
        hideLoadingState(); // Ocultar loading después de carga exitosa
        
    } catch (error) {
        console.error('Error cargando datos del proyecto:', error);
        showNotification('Error al cargar los datos del proyecto', 'error');
        hideLoadingState(); // Ocultar loading en caso de error
    }
}

// ========================================
// FUNCIONES PARA CALCULAR RESÚMENES AUTOMÁTICAMENTE
// ========================================

// Calcular resumen general basándose en los detalles de pruebas
function calculateSummaryFromDetails(pruebas) {
    const total = pruebas.length;
    let exitosas = 0;
    let fallidas = 0;
    let pendientes = 0;
    let bloqueadas = 0;
    let planificadas = 0;
    let desestimadas = 0;
    
    // Contar por estado
    pruebas.forEach(prueba => {
        // Aplicar mapStatus para mapear correctamente desde Excel
        const estadoOriginal = prueba.Estado || prueba.estado || '';
        const estadoMapeado = mapStatus(estadoOriginal);
        console.log(`DEBUG CONTEO: Estado original "${estadoOriginal}" -> mapeado "${estadoMapeado}"`);
        
        switch (estadoMapeado.toLowerCase()) {
            case 'exitosa':
            case 'success':
                exitosas++;
                console.log(`  ✅ Contando como EXITOSA (total exitosas: ${exitosas})`);
                break;
            case 'fallida':
            case 'failed':
            case 'failure':
                fallidas++;
                console.log(`  ❌ Contando como FALLIDA (total fallidas: ${fallidas})`);
                break;
            case 'pendiente':
            case 'pending':
                pendientes++;
                console.log(`  ⏳ Contando como PENDIENTE (total pendientes: ${pendientes})`);
                break;
            case 'bloqueada':
            case 'blocked':
                bloqueadas++;
                console.log(`  🚫 Contando como BLOQUEADA (total bloqueadas: ${bloqueadas})`);
                break;
            case 'desestimado':
            case 'dismissed':
                desestimadas++;
                console.log(`  🚯 Contando como DESESTIMADA (total desestimadas: ${desestimadas})`);
                break;
            case 'planificada':
            case 'planned':
            case '':
            case undefined:
            case null:
                planificadas++; // Solo celdas vacías o explícitamente planificadas
                console.log(`  📋 Contando como PLANIFICADA (total planificadas: ${planificadas})`);
                break;
            default:
                console.warn(`⚠️ Estado desconocido en calculateSummaryFromDetails: "${estadoOriginal}" -> "${estadoMapeado}"`);
                pendientes++; // Estados desconocidos se consideran pendientes, no planificadas
                console.log(`  ❓ Estado desconocido, contando como PENDIENTE (total pendientes: ${pendientes})`);
                break;
        }
    });
    
    console.log(`📊 RESUMEN CALCULADO DESDE EXCEL:`);
    console.log(`   • Total: ${total}`);
    console.log(`   • Planificadas: ${planificadas}`);
    console.log(`   • Exitosas: ${exitosas}`);
    console.log(`   • Fallidas: ${fallidas}`);
    console.log(`   • Pendientes: ${pendientes}`);
    console.log(`   • Bloqueadas: ${bloqueadas}`);
    console.log(`   • DESESTIMADAS: ${desestimadas}`);
    
    return {
        proyecto_id: pruebas[0]?.proyecto_id || '',
        pruebas_planificadas: planificadas,
        pruebas_exitosas: exitosas,
        pruebas_fallidas: fallidas,
        pruebas_pendientes: pendientes,
        pruebas_bloqueadas: bloqueadas,
        pruebas_desestimadas: desestimadas,
        fecha_actualizacion: new Date().toISOString().split('T')[0]
    };
}

// Calcular categorías/escenarios basándose en los detalles de pruebas
function calculateCategoriesFromDetails(pruebas) {
    const categorias = {};
    
    // Agrupar por escenario
    pruebas.forEach(prueba => {
        const escenario = prueba.escenario || prueba.categoria || 'Sin Categoría';
        
        if (!categorias[escenario]) {
            categorias[escenario] = {
                proyecto_id: prueba.proyecto_id,
                escenario: escenario,
                planificadas: 0,
                exitosas: 0,
                fallidas: 0,
                pendientes: 0,
                bloqueadas: 0,
                desestimadas: 0,
                prioridad: 'Media',
                responsable: 'Equipo QA'
            };
        }
        
        // Contar por estado - aplicar mapStatus para mapear correctamente desde Excel
        const estadoOriginal = prueba.Estado || prueba.estado || '';
        const estadoMapeado = mapStatus(estadoOriginal);
        
        switch (estadoMapeado.toLowerCase()) {
            case 'exitosa':
            case 'success':
                categorias[escenario].exitosas++;
                break;
            case 'fallida':
            case 'failed':
            case 'failure':
                categorias[escenario].fallidas++;
                break;
            case 'pendiente':
            case 'pending':
                categorias[escenario].pendientes++;
                break;
            case 'bloqueada':
            case 'blocked':
                categorias[escenario].bloqueadas++;
                break;
            case 'desestimado':
            case 'dismissed':
                categorias[escenario].desestimadas++;
                break;
            case '':
            case undefined:
            case null:
                categorias[escenario].planificadas++; // Celdas vacías se consideran planificadas
                break;
            default:
                categorias[escenario].planificadas++; // Estados desconocidos se consideran planificadas
                break;
        }
    });
    
    const resultado = Object.values(categorias);
    console.log('Categorías calculadas:', resultado);
    
    return resultado;
}

// Transformar datos multi-proyecto a formato dashboard
function transformMultiProjectData(data) {
    // Debug: mostrar fechas originales
    console.log('Fechas originales del Excel:', {
        fecha_inicio: data.proyecto.fecha_inicio,
        fecha_fin: data.proyecto.fecha_fin,
        tipo_inicio: typeof data.proyecto.fecha_inicio,
        tipo_fin: typeof data.proyecto.fecha_fin
    });
    
    console.log(`🔍 DEBUGGING TRANSFORM - VALORES DEL RESUMEN:`);
    console.log(`   • data.resumen.pruebas_planificadas: ${data.resumen.pruebas_planificadas}`);
    console.log(`   • data.resumen.pruebas_exitosas: ${data.resumen.pruebas_exitosas}`);
    console.log(`   • data.resumen.pruebas_fallidas: ${data.resumen.pruebas_fallidas}`);
    console.log(`   • data.resumen.pruebas_pendientes: ${data.resumen.pruebas_pendientes}`);
    console.log(`   • data.resumen.pruebas_bloqueadas: ${data.resumen.pruebas_bloqueadas}`);
    console.log(`   • data.resumen.pruebas_desestimadas: ${data.resumen.pruebas_desestimadas}`);
    
    console.log(`📋 DEBUGGING ENTREGABLES - Datos desde Excel:`);
    console.log(`   • data.proyecto.PP: "${data.proyecto.PP}"`);
    console.log(`   • data.proyecto.MT: "${data.proyecto.MT}"`);
    console.log(`   • data.proyecto.CP: "${data.proyecto.CP}"`);
    console.log(`   • data.proyecto.EE: "${data.proyecto.EE}"`);
    console.log(`   • data.proyecto.DF: "${data.proyecto.DF}"`);
    console.log(`   • data.proyecto.AC: "${data.proyecto.AC}"`);
    
    const transformed = {
        projectInfo: {
            name: data.proyecto.nombre_proyecto,
            qaResponsible: data.proyecto.responsable_qa,
            startDate: formatDateForDashboard(data.proyecto.fecha_inicio),
            endDate: formatDateForDashboard(data.proyecto.fecha_fin),
            status: data.proyecto.estado_proyecto,
            progress: 0,  // Se calculará automáticamente basado en pruebas exitosas
            // Campos de entregables (semáforo)
            PP: data.proyecto.PP || '',  // Plan de Pruebas
            MT: data.proyecto.MT || '',  // Matriz de Trazabilidad
            CP: data.proyecto.CP || '',  // Matriz de Casos de Prueba
            EE: data.proyecto.EE || '',  // Evidencias de Prueba
            DF: data.proyecto.DF || '',  // Matriz de Defectos
            AC: data.proyecto.AC || ''   // Informe Final (Acta de Cierre)
        },
        summary: {
            planned: data.resumen.pruebas_planificadas,
            successful: data.resumen.pruebas_exitosas,
            failed: data.resumen.pruebas_fallidas,
            pending: data.resumen.pruebas_pendientes,
            blocked: data.resumen.pruebas_bloqueadas || 0,
            dismissed: data.resumen.pruebas_desestimadas || 0
        },
        trend: (() => {
            // Mapear datos de tendencia desde Excel
            const trendData = data.tendencia.map(t => ({
                date: formatDateForDashboard(t.fecha),
                planned: t.planificadas,
                successful: t.exitosas,
                failed: t.fallidas,
                pending: t.pendientes,
                blocked: t.bloqueadas || 0,
                dismissed: t.desestimadas || 0,
                realProgress: t.real_progress || t.realProgress || t['% Real'] || 0
            }));
            
            return trendData;
        })(),
        categories: data.categorias.map(c => ({
            escenario: c.escenario,
            planned: c.planificadas,
            successful: c.exitosas,
            failed: c.fallidas,
            pending: c.pendientes,
            blocked: c.bloqueadas || 0,
            dismissed: c.desestimadas || 0,
            priority: c.prioridad,
            responsible: c.responsable
        })),
        testDetails: (() => {
            // Debug: mostrar datos de pruebas originales
            console.log('=== DEBUG DETALLES DE PRUEBAS ===');
            console.log('Cantidad de pruebas en Excel:', data.pruebas.length);
            console.log('Muestra de pruebas originales:', data.pruebas.slice(0, 3));
            
            const mappedTests = data.pruebas.map(p => {
                // Debug específico para la fecha de ejecución
                console.log(`Prueba ${p.test_id}: dia_ejecutado =`, p.dia_ejecutado, `(tipo: ${typeof p.dia_ejecutado})`);
                
                // Procesar la fecha de ejecución usando la función correcta
                const fechaFormateada = formatDateForDashboard(p.dia_ejecutado);
                console.log(`Fecha formateada para ${p.test_id}:`, fechaFormateada);
                
                return {
                    id: p.test_id,
                    name: p.nombre_prueba,
                    escenario: p.escenario,
                    status: mapExcelStatusToDashboard(p.estado),
                    dia_ejecutado: fechaFormateada,
                    executor: p.ejecutor,
                    environment: p.ambiente,
                    comments: p.comentarios
                };
            });
            
            console.log('Pruebas después de mapeo:', mappedTests.slice(0, 3));
            console.log('===============================');
            return mappedTests;
        })(),
        defects: {
            summary: calculateDefectsSummary(data.defectos),
            details: data.defectos.map(d => ({
                id: d.defecto_id,
                title: d.titulo,
                description: d.descripcion,
                severity: d.severidad,
                status: d.estado,
                cp_impactados: d.cp_impactados,
                assignee: d.asignado_a,
                reporter: d.reportado_por,
                dateFound: formatDateForDashboard(d.fecha_encontrado),
                dateAssigned: formatDateForDashboard(d.fecha_asignacion),
                dateResolved: formatDateForDashboard(d.fecha_resolucion),
                stepsToReproduce: d.pasos_reproducir,
                environment: d.ambiente_encontrado
            }))
        }
    };
    
    // Debug: mostrar fechas transformadas
    console.log('Fechas transformadas:', {
        startDate: transformed.projectInfo.startDate,
        endDate: transformed.projectInfo.endDate
    });
    
    console.log(`🔍 DEBUGGING TRANSFORM - SUMMARY FINAL TRANSFORMADO:`);
    console.log(`   • planned: ${transformed.summary.planned}`);
    console.log(`   • successful: ${transformed.summary.successful}`);
    console.log(`   • failed: ${transformed.summary.failed}`);
    console.log(`   • pending: ${transformed.summary.pending}`);
    console.log(`   • blocked: ${transformed.summary.blocked}`);
    console.log(`   • dismissed: ${transformed.summary.dismissed}`);
    
    return transformed;
}

// Calcular resumen de defectos
function calculateDefectsSummary(defects) {
    return {
        critical: defects.filter(d => d.severidad === 'Crítica').length,
        high: defects.filter(d => d.severidad === 'Alta').length,
        medium: defects.filter(d => d.severidad === 'Media').length,
        low: defects.filter(d => d.severidad === 'Baja').length
    };
}

// Formatear fecha para dashboard - maneja diferentes formatos de Excel
function formatDateForDashboard(dateValue) {
    if (!dateValue && dateValue !== 0) return '';
    
    console.log('formatDateForDashboard input:', dateValue, typeof dateValue);
    
    let date;
    
    // Si es un número (fecha serial de Excel)
    if (typeof dateValue === 'number') {
        // Excel fecha serial: días desde 1900-01-01
        // Excel considera 1900 como año bisiesto (bug histórico)
        const excelEpoch = new Date(1900, 0, 1);
        const millisecondsPerDay = 24 * 60 * 60 * 1000;
        
        // Ajustar por el bug del año bisiesto de Excel
        let adjustedDays = dateValue;
        if (dateValue > 59) {
            adjustedDays = dateValue - 1;
        }
        
        date = new Date(excelEpoch.getTime() + (adjustedDays - 1) * millisecondsPerDay);
        console.log('Excel serial convertido:', dateValue, '->', date);
    }
    // Si es un string
    else if (typeof dateValue === 'string') {
        // Limpiar string y corregir formatos corruptos
        let cleanedDate = dateValue.trim();
        
        // Corregir formato corrupto como "24/1072025" -> "24/10/2025"
        if (cleanedDate.match(/^\d{1,2}\/\d{6,8}$/)) {
            const parts = cleanedDate.split('/');
            if (parts.length === 2 && parts[1].length >= 6) {
                const day = parts[0];
                const monthYear = parts[1];
                // Asumir que los primeros 2 dígitos son el mes
                const month = monthYear.substring(0, 2);
                const year = monthYear.substring(2);
                cleanedDate = `${day}/${month}/${year}`;
                console.log('Formato corrupto corregido:', dateValue, '->', cleanedDate);
            }
        }
        
        // Formato DD/MM/AAAA HH:MM (con hora y minutos)
        if (cleanedDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}$/)) {
            const [datePart, timePart] = cleanedDate.split(/\s+/);
            const [day, month, year] = datePart.split('/').map(Number);
            
            date = new Date(year, month - 1, day);
            console.log('DD/MM/AAAA HH:MM convertido (ignorando hora):', dateValue, '->', date);
        }
        // Formato DD/MM/AAAA (solo fecha)
        else if (cleanedDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
            const [day, month, year] = cleanedDate.split('/').map(Number);
            date = new Date(year, month - 1, day);
            console.log('DD/MM/AAAA convertido:', cleanedDate, '->', date);
        }
        // Formato DD-MM-AAAA HH:MM
        else if (cleanedDate.match(/^\d{1,2}-\d{1,2}-\d{4}\s+\d{1,2}:\d{2}$/)) {
            const [datePart, timePart] = cleanedDate.split(/\s+/);
            const [day, month, year] = datePart.split('-').map(Number);
            
            date = new Date(year, month - 1, day);
        }
        // Formato DD-MM-AAAA (solo fecha)
        else if (cleanedDate.match(/^\d{1,2}-\d{1,2}-\d{4}$/)) {
            const [day, month, year] = cleanedDate.split('-').map(Number);
            date = new Date(year, month - 1, day);
        }
        // Formato YYYY-MM-DD HH:MM
        else if (cleanedDate.match(/^\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}$/)) {
            const [datePart, timePart] = cleanedDate.split(/\s+/);
            const [year, month, day] = datePart.split('-').map(Number);
            
            date = new Date(year, month - 1, day);
        }
        // Formato YYYY-MM-DD (solo fecha)
        else if (cleanedDate.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
            const [year, month, day] = cleanedDate.split('-').map(Number);
            date = new Date(year, month - 1, day);
        }
        else {
            // Intentar parsear directamente
            date = new Date(cleanedDate);
        }
        
        console.log('String convertido:', dateValue, '->', date);
    }
    // Si es ya un objeto Date
    else if (dateValue instanceof Date) {
        date = dateValue;
    }
    else {
        // Último intento
        date = new Date(dateValue);
    }
    
    // Validar que la fecha es válida
    if (isNaN(date.getTime())) {
        console.warn('Fecha inválida:', dateValue);
        return String(dateValue); // Devolver el valor original como string
    }
    
    // Formatear fecha (solo fecha, sin hora)
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    
    // Siempre devolver solo la fecha sin hora
    let formatted = `${day}/${month}/${year}`;
    
    console.log('Fecha final formateada (solo fecha):', formatted);
    return formatted;
}

// Mapear estado de Excel a dashboard
function mapExcelStatusToDashboard(status) {
    const statusMap = {
        'Exitosa': 'success',
        'Fallida': 'failure',
        'Pendiente': 'pending',
        'Bloqueada': 'blocked',
        'Desestimado': 'dismissed',
        'Planificada': 'planned'
    };
    
    console.log(`🔄 mapExcelStatusToDashboard: "${status}" → "${statusMap[status] || 'pending'}"`);
    return statusMap[status] || 'pending';
}

// Actualizar dashboard con datos transformados
function updateDashboardWithTransformedData(data) {
    try {
        // Actualizar datos globales
        Object.assign(testData, data);
        
        // Reiniciar variables de paginación
        filteredTests = [];
        currentPage = 1;
        
        // Reiniciar variables de paginación de defectos
        filteredDefects = [];
        defectsCurrentPage = 1;
        
        // Actualizar interfaz
        initializeDashboard();
        createCharts();
        populateTestTable();
        populateDefectsTable();
        
        // Guardar en localStorage para el editor
        localStorage.setItem('qaEditorData', JSON.stringify(data));
        
        // Guardar datos para funcionalidad de compartir enlace
        currentExcelData = data;
        
        // Mostrar el botón de seleccionar proyecto
        const selectProjectButton = document.getElementById('selectProjectButton');
        if (selectProjectButton) {
            selectProjectButton.style.display = 'flex';
        }
        
        console.log('Dashboard actualizado exitosamente');
    } catch (error) {
        console.error('Error actualizando dashboard:', error);
        showNotification('Error al actualizar el dashboard: ' + error.message, 'error');
    }
}

// Manejar Excel formato legacy
function handleLegacyExcel(workbook) {
    let hasDetailsSheet = false;
    let detailsData = [];
    
    // Procesar hoja de detalles primero si existe
    if (workbook.SheetNames.includes('Detalle_Pruebas')) {
        hasDetailsSheet = true;
        const worksheet = workbook.Sheets['Detalle_Pruebas'];
        detailsData = XLSX.utils.sheet_to_json(worksheet);
        updateDashboardWithExcelData(detailsData, 'details');
    }
    
    // Si existe la hoja de detalles, NO procesar la hoja de resumen (se calcula automáticamente)
    if (!hasDetailsSheet && workbook.SheetNames.includes('Resumen_Diario')) {
        const worksheet = workbook.Sheets['Resumen_Diario'];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        updateDashboardWithExcelData(jsonData, 'summary');
    }
    
    if (workbook.SheetNames.includes('Tendencia_Semanal')) {
        const worksheet = workbook.Sheets['Tendencia_Semanal'];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        updateDashboardWithExcelData(jsonData, 'trend');
    }
    
    // CALCULAR CATEGORÍAS AUTOMÁTICAMENTE desde Detalle_Pruebas en lugar de leer la hoja
    if (hasDetailsSheet && detailsData.length > 0) {
        const calculatedCategories = calculateCategoriesFromDetails(detailsData);
        updateDashboardWithExcelData(calculatedCategories, 'categories');
    } else if (workbook.SheetNames.includes('Categorias_Pruebas')) {
        // Solo usar la hoja si NO hay detalles (compatibilidad legacy)
        const worksheet = workbook.Sheets['Categorias_Pruebas'];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        updateDashboardWithExcelData(jsonData, 'categories');
    }
    
    showNotification('Archivo Excel cargado exitosamente', 'success');
}

// Parsear CSV
function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];
    
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim() === '') continue;
        
        const values = lines[i].split(',');
        const row = {};
        
        headers.forEach((header, index) => {
            row[header] = values[index] ? values[index].trim() : '';
        });
        
        data.push(row);
    }
    
    return data;
}

// Verificar si es CSV de detalles de pruebas
function isTestDetailsCSV(data) {
    if (data.length === 0) return false;
    const firstRow = data[0];
    return firstRow.hasOwnProperty('ID_Prueba') && firstRow.hasOwnProperty('Nombre_Prueba');
}

// Verificar si es CSV de resumen
function isSummaryCSV(data) {
    if (data.length === 0) return false;
    const firstRow = data[0];
    return (firstRow.hasOwnProperty('Fecha') && firstRow.hasOwnProperty('Pruebas_Exitosas')) ||
           (firstRow.hasOwnProperty('Fecha') && firstRow.hasOwnProperty('Pruebas_Planificadas'));
}

// Verificar si es CSV de defectos
function isDefectsCSV(data) {
    if (data.length === 0) return false;
    const firstRow = data[0];
    return firstRow.hasOwnProperty('defecto_id') && firstRow.hasOwnProperty('titulo') && firstRow.hasOwnProperty('severidad');
}

// Actualizar dashboard con datos CSV
function updateDashboardWithCSVData(data, type) {
    if (type === 'details') {
        // Actualizar detalles de pruebas
        testData.testDetails = data.map(row => {
            console.log('Fecha original del Excel:', row.dia_ejecutado, typeof row.dia_ejecutado);
            return {
                id: row.ID_Prueba || '',
                name: row.Nombre_Prueba || '',
                escenario: row.Escenario || '',
                status: mapStatus(row.Estado || ''),
                dia_ejecutado: row.dia_ejecutado || '',
                executor: row.Ejecutor || ''
            };
        });
        
        // COMENTADO: No recalcular resumen cuando viene de Excel multi-proyecto
        // porque ya se calculó correctamente en calculateSummaryFromDetails()
        // updateSummaryFromDetails();
        populateTestTable();
        
    } else if (type === 'summary') {
        // Actualizar datos de resumen
        const latestData = data[data.length - 1];
        if (latestData) {
            testData.summary = {
                planned: parseInt(latestData.Pruebas_Planificadas || testData.summary.planned || 100),
                successful: parseInt(latestData.Pruebas_Exitosas || 0),
                failed: parseInt(latestData.Pruebas_Fallidas || 0),
                pending: parseInt(latestData.Pruebas_Pendientes || 0)
            };
        }
        
        initializeDashboard();
    } else if (type === 'defects') {
        // Actualizar datos de defectos
        testData.defects = {
            summary: calculateDefectsSummary(data),
            details: data.map(row => ({
                id: row.defecto_id || '',
                title: row.titulo || '',
                description: row.descripcion || '',
                severity: row.severidad || '',
                status: row.estado || '',
                cp_impactados: row.cp_impactados || '',
                assignee: row.asignado_a || '',
                reporter: row.reportado_por || '',
                dateFound: formatDateForDashboard(row.fecha_encontrado || ''),
                dateAssigned: formatDateForDashboard(row.fecha_asignacion || ''),
                dateResolved: formatDateForDashboard(row.fecha_resolucion || ''),
                stepsToReproduce: row.pasos_reproducir || '',
                environment: row.ambiente_encontrado || ''
            }))
        };
        
        populateDefectsTable();
    }
    
    // Recrear gráficas
    createCharts();
}

// Actualizar dashboard con datos Excel
function updateDashboardWithExcelData(data, type) {
    switch (type) {
        case 'details':
            testData.testDetails = data.map(row => {
                console.log('Fecha Excel (updateDashboardWithExcelData):', row.dia_ejecutado, typeof row.dia_ejecutado);
                return {
                    id: row.ID_Prueba || '',
                    name: row.Nombre_Prueba || '',
                    escenario: row.Escenario || '',
                    status: mapStatus(row.Estado || ''),
                    dia_ejecutado: row.dia_ejecutado || '',
                    executor: row.Ejecutor || ''
                };
            });
            // COMENTADO: No recalcular resumen para evitar doble conteo
            // updateSummaryFromDetails();
            populateTestTable();
            break;
            
        case 'summary':
            const latestData = data[data.length - 1];
            if (latestData) {
                testData.summary = {
                    planned: parseInt(latestData.Pruebas_Planificadas || testData.summary.planned || 100),
                    successful: parseInt(latestData.Pruebas_Exitosas || 0),
                    failed: parseInt(latestData.Pruebas_Fallidas || 0),
                    pending: parseInt(latestData.Pruebas_Pendientes || 0)
                };
            }
            initializeDashboard();
            break;
            
        case 'trend':
            testData.trend = data.map(row => ({
                date: row.Fecha || '',
                planned: parseInt(row.Planificadas || row.Planned || 0),
                successful: parseInt(row.Exitosas || 0),
                failed: parseInt(row.Fallidas || 0),
                pending: parseInt(row.Pendientes || 0)
            }));
            break;
            
        case 'categories':
            testData.categories = data.map(row => ({
                escenario: row.Escenario || '',
                planned: parseInt(row.Planificadas || row.Planned || 0),
                successful: parseInt(row.Exitosas || 0),
                failed: parseInt(row.Fallidas || 0),
                pending: parseInt(row.Pendientes || 0)
            }));
            break;
            
        case 'defects':
            testData.defects = {
                summary: calculateDefectsSummary(data),
                details: data.map(row => ({
                    id: row.defecto_id || '',
                    title: row.titulo || '',
                    description: row.descripcion || '',
                    severity: row.severidad || '',
                    status: row.estado || '',
                    cp_impactados: row.cp_impactados || '',
                    assignee: row.asignado_a || '',
                    reporter: row.reportado_por || '',
                    dateFound: formatDateForDashboard(row.fecha_encontrado || ''),
                    dateAssigned: formatDateForDashboard(row.fecha_asignacion || ''),
                    dateResolved: formatDateForDashboard(row.fecha_resolucion || ''),
                    stepsToReproduce: row.pasos_reproducir || '',
                    environment: row.ambiente_encontrado || ''
                }))
            };
            populateDefectsTable();
            break;
    }
    
    createCharts();
}

// Mapear estados del archivo a formato interno
function mapStatus(status) {
    // Si el estado está vacío o es null/undefined, considerarlo como planificada
    if (!status || status.toString().trim() === '') {
        return 'planned';
    }
    
    // Normalizar el estado:
    // 1. Convertir a string por si viene como número
    // 2. Quitar espacios al inicio y final
    // 3. Convertir a minúsculas
    // 4. Quitar la 's' final si existe (para manejar plurales)
    const normalizedStatus = status.toString().trim().toLowerCase();
    const singularStatus = normalizedStatus.replace(/s+$/, ''); // Quitar 's' o 'ss' al final
    
    // Mapeo amplio que incluye múltiples variaciones
    const statusPatterns = {
        // Exitosa/Exitoso/Success
        'exitosa': 'success',
        'exitoso': 'success',
        'exito': 'success',
        'correcta': 'success',
        'correcto': 'success',
        'success': 'success',
        'successful': 'success',
        'passed': 'success',
        'pass': 'success',
        'ok': 'success',
        
        // Fallida/Failed
        'fallida': 'failure',
        'fallido': 'failure',
        'falla': 'failure',
        'fallo': 'failure',
        'failed': 'failure',
        'failure': 'failure',
        'error': 'failure',
        'fail': 'failure',
        
        // Pendiente/Pending
        'pendiente': 'pending',
        'pending': 'pending',
        'en espera': 'pending',
        'esperando': 'pending',
        'waiting': 'pending',
        'wait': 'pending',
        
        // Bloqueada/Blocked
        'bloqueada': 'blocked',
        'bloqueado': 'blocked',
        'bloqueo': 'blocked',
        'blocked': 'blocked',
        'block': 'blocked',
        
        // Planificada/Planned
        'planificada': 'planned',
        'planificado': 'planned',
        'planeada': 'planned',
        'planeado': 'planned',
        'planned': 'planned',
        'plan': 'planned',
        'por ejecutar': 'planned',
        
        // Desestimada/Dismissed
        'desestimada': 'dismissed',
        'desestimado': 'dismissed',
        'desestima': 'dismissed',
        'dismissed': 'dismissed',
        'dismiss': 'dismissed',
        'descartada': 'dismissed',
        'descartado': 'dismissed',
        'cancelada': 'dismissed',
        'cancelado': 'dismissed'
    };
    
    // Intentar mapear primero con el texto completo
    let mappedStatus = statusPatterns[normalizedStatus];
    
    // Si no encuentra coincidencia, intentar con la versión singular
    if (!mappedStatus) {
        mappedStatus = statusPatterns[singularStatus];
    }
    
    // Si aún no hay coincidencia, intentar buscar si alguna palabra clave está contenida
    if (!mappedStatus) {
        for (const [key, value] of Object.entries(statusPatterns)) {
            if (normalizedStatus.includes(key) || key.includes(normalizedStatus)) {
                mappedStatus = value;
                break;
            }
        }
    }
    
    if (!mappedStatus) {
        console.warn(`⚠️ ESTADO DESCONOCIDO: "${status}" - NO SE MAPEA A NINGÚN ESTADO CONOCIDO (será considerado como PENDING)`);
        return 'pending'; // Por defecto, estado desconocido se considera pendiente, no planificado
    }
    
    console.log(`🔄 MAPEO ESTADO: "${status}" → "${mappedStatus}"`);
    return mappedStatus;
}

// Actualizar resumen basado en detalles
function updateSummaryFromDetails() {
    const successful = testData.testDetails.filter(t => t.status === 'success').length;
    const failed = testData.testDetails.filter(t => t.status === 'failure').length;
    const pending = testData.testDetails.filter(t => t.status === 'pending').length;
    const blocked = testData.testDetails.filter(t => t.status === 'blocked').length;
    const planned = testData.testDetails.filter(t => t.status === 'planned').length;
    const dismissed = testData.testDetails.filter(t => t.status === 'dismissed').length;
    
    console.log('📊 ACTUALIZANDO SUMMARY DESDE DETALLES:');
    console.log(`   • Exitosas: ${successful}`);
    console.log(`   • Fallidas: ${failed}`);
    console.log(`   • Pendientes: ${pending}`);
    console.log(`   • Bloqueadas: ${blocked}`);
    console.log(`   • Planificadas: ${planned}`);
    console.log(`   • DESESTIMADAS: ${dismissed}`);
    
    testData.summary = {
        planned,
        successful,
        failed,
        pending,
        blocked,
        dismissed
    };
    
    console.log('🔄 Summary actualizado:', testData.summary);
    
    // Actualizar la información del proyecto con el nuevo progreso calculado
    initializeDashboard();
    
    // Actualizar gráfico de cobertura con los nuevos datos
    updateCoverageChart(testData.testDetails);
    
    // Actualizar gráfico de ejecución diaria
    updateDailyExecutionChart(testData.testDetails);
    
    // Actualizar gráfico de distribución por ejecutor
    updateExecutorDistributionChart(testData.testDetails);
    
    // Actualizar gráfico de entregables
    updateDeliverablesChart();
    
    console.log('✅ Resumen, cobertura, ejecución diaria, distribución por ejecutor y entregables actualizados desde detalles');
}

// Mostrar notificaciones
function showNotification(message, type = 'info') {
    // Crear elemento de notificación
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <i class="fas ${type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-exclamation-circle' : 'fa-info-circle'}"></i>
        <span>${message}</span>
        <button class="notification-close" onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    // Agregar estilos si no existen
    if (!document.getElementById('notification-styles')) {
        const style = document.createElement('style');
        style.id = 'notification-styles';
        style.textContent = `
            .notification {
                position: fixed;
                top: 20px;
                right: 20px;
                padding: 1rem 1.5rem;
                border-radius: 8px;
                display: flex;
                align-items: center;
                gap: 0.75rem;
                color: white;
                font-size: 0.875rem;
                z-index: 1000;
                box-shadow: var(--shadow-lg);
                animation: slideIn 0.3s ease-out;
            }
            .notification-success { background: #10b981; }
            .notification-error { background: #ef4444; }
            .notification-info { background: #2563eb; }
            .notification-close {
                background: none;
                border: none;
                color: white;
                cursor: pointer;
                padding: 0.25rem;
                opacity: 0.8;
            }
            .notification-close:hover { opacity: 1; }
            @keyframes slideIn {
                from { transform: translateX(100%); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        if (notification.parentElement) {
            notification.remove();
        }
    }, 5000);
}

// Configurar integración con el editor
function setupEditorIntegration() {
    // Escuchar mensajes del editor
    window.addEventListener('message', function(event) {
        if (event.data.type === 'QA_DATA_UPDATE') {
            updateDashboardWithEditorData(event.data.data);
        }
    });
    
    // Escuchar cambios en localStorage
    window.addEventListener('storage', function(event) {
        if (event.key === 'qaEditorData' && event.newValue) {
            try {
                const data = JSON.parse(event.newValue);
                updateDashboardWithEditorData(data);
            } catch (error) {
                console.error('Error procesando datos del localStorage:', error);
            }
        }
    });
    
    // Polling para detectar cambios en localStorage (fallback)
    setInterval(() => {
        loadEditorData();
    }, 2000);
    
    // Notificar que el dashboard está listo
    if (window.opener && !window.opener.closed) {
        window.opener.postMessage({ type: 'DASHBOARD_READY' }, '*');
    }
    
    // Cargar datos del editor si están disponibles
    loadEditorData();
}

// Variable para rastrear la última actualización
let lastDataHash = '';

// Cargar datos del editor desde localStorage
function loadEditorData() {
    try {
        const savedData = localStorage.getItem('qaEditorData');
        if (savedData) {
            // Crear hash simple usando la longitud y algunos caracteres del string
            const currentHash = simpleHash(savedData);
            if (currentHash !== lastDataHash) {
                const data = JSON.parse(savedData);
                updateDashboardWithEditorData(data);
                lastDataHash = currentHash;
                console.log('Dashboard actualizado desde el editor');
            }
        }
    } catch (error) {
        console.error('Error cargando datos del editor:', error);
    }
}

// Función helper para crear un hash simple que soporte Unicode
function simpleHash(str) {
    let hash = 0;
    if (str.length === 0) return hash.toString();
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convertir a 32-bit integer
    }
    return hash.toString();
}

// Actualizar dashboard con datos del editor
function updateDashboardWithEditorData(data) {
    try {
        // Actualizar datos
        if (data.summary) {
            testData.summary = data.summary;
        }
        
        if (data.trend) {
            testData.trend = data.trend;
        }
        
        if (data.categories) {
            testData.categories = data.categories;
        }
        
        if (data.testDetails) {
            testData.testDetails = data.testDetails;
        }
        
        if (data.defects) {
            testData.defects = data.defects;
        }
        
        // Actualizar interfaz
        initializeDashboard();
        createCharts();
        populateTestTable();
        populateDefectsTable();
        
        console.log('Dashboard actualizado desde el editor');
    } catch (error) {
        console.error('Error actualizando dashboard:', error);
        showNotification('Error al actualizar desde el editor', 'error');
    }
}

// Abrir editor de datos
function openDataEditor() {
    const editorWindow = window.open('template_excel_modelo.html', 'dataEditor', 'width=1400,height=900');
    
    if (editorWindow) {
        // Enviar datos actuales al editor cuando se cargue
        editorWindow.addEventListener('load', function() {
            const currentData = {
                summary: testData.summary,
                trend: testData.trend,
                categories: testData.categories,
                testDetails: testData.testDetails
            };
            
            editorWindow.postMessage({
                type: 'LOAD_CURRENT_DATA',
                data: currentData
            }, '*');
        });
        
        showNotification('Editor de datos abierto en nueva ventana', 'success');
    } else {
        showNotification('No se pudo abrir el editor. Verifica los pop-ups bloqueados.', 'error');
    }
}

// Función para editar información del proyecto
function editProjectInfo() {
    // Verificar que testData.projectInfo existe
    if (!testData.projectInfo) {
        console.error('No hay datos de proyecto disponibles');
        showNotification('Error: No se encontraron datos del proyecto', 'error');
        return;
    }
    
    const modal = document.createElement('div');
    modal.className = 'modal-overlay';
    modal.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100%; height: 100%;
        background: rgba(0,0,0,0.8); display: flex; align-items: center;
        justify-content: center; z-index: 1000; backdrop-filter: blur(4px);
    `;
    
    // Función para convertir fecha DD/MM/YYYY a YYYY-MM-DD
    function convertDateForInput(dateStr) {
        if (!dateStr || !dateStr.includes('/')) return '';
        const parts = dateStr.split('/');
        if (parts.length !== 3) return '';
        return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
    }
    
    const form = document.createElement('div');
    form.className = 'modal-content';
    form.style.cssText = `
        background: #ffffff; border-radius: 12px; padding: 2rem;
        max-width: 500px; width: 90%; box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        max-height: 80vh; overflow-y: auto; border: 1px solid #e2e8f0;
        color: #1e293b;
    `;
    
    // Detectar modo oscuro
    const isDarkMode = document.documentElement.getAttribute('data-theme') === 'dark' || 
                      window.matchMedia('(prefers-color-scheme: dark)').matches;
    
    if (isDarkMode) {
        form.style.background = '#1e293b';
        form.style.color = '#f8fafc';
        form.style.border = '1px solid #475569';
    }
    
    form.innerHTML = `
        <h3 style="margin: 0 0 1.5rem 0; font-size: 1.25rem; font-weight: 600; color: inherit;">
            Información del Proyecto
        </h3>
        
        <div style="margin-bottom: 1rem;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                Nombre del Proyecto:
            </label>
            <select id="editProjectName" style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                                               border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                                               color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                                               box-sizing: border-box; transition: border-color 0.3s ease;"
                    onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                    onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'">
                <option value="">Seleccionar proyecto...</option>
                <option value="Building Block - Aplicaciones de Soporte" ${testData.projectInfo.name === 'Building Block - Aplicaciones de Soporte' ? 'selected' : ''}>Building Block - Aplicaciones de Soporte</option>
                <option value="Building Block - Canales Asistidos" ${testData.projectInfo.name === 'Building Block - Canales Asistidos' ? 'selected' : ''}>Building Block - Canales Asistidos</option>
                <option value="Building Block - Canales Digitales" ${testData.projectInfo.name === 'Building Block - Canales Digitales' ? 'selected' : ''}>Building Block - Canales Digitales</option>
                <option value="Building Block - Cards" ${testData.projectInfo.name === 'Building Block - Cards' ? 'selected' : ''}>Building Block - Cards</option>
                <option value="Building Block - Ciberseguridad" ${testData.projectInfo.name === 'Building Block - Ciberseguridad' ? 'selected' : ''}>Building Block - Ciberseguridad</option>
                <option value="Building Block - Core" ${testData.projectInfo.name === 'Building Block - Core' ? 'selected' : ''}>Building Block - Core</option>
                <option value="Building Block - Data" ${testData.projectInfo.name === 'Building Block - Data' ? 'selected' : ''}>Building Block - Data</option>
                <option value="Building Block - Integración" ${testData.projectInfo.name === 'Building Block - Integración' ? 'selected' : ''}>Building Block - Integración</option>
                <option value="Building Block - Payments" ${testData.projectInfo.name === 'Building Block - Payments' ? 'selected' : ''}>Building Block - Payments</option>
                <option value="Building Block - Procesos" ${testData.projectInfo.name === 'Building Block - Procesos' ? 'selected' : ''}>Building Block - Procesos</option>
            </select>
        </div>
        
        <div style="margin-bottom: 1rem;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                Responsable QA:
            </label>
            <input type="text" id="editQAResponsible" value="${testData.projectInfo.qaResponsible || ''}" 
                   style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                          border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                          color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                          box-sizing: border-box; transition: border-color 0.3s ease;"
                   onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                   onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'" />
        </div>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem;">
            <div>
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                    Fecha Inicio:
                </label>
                <input type="date" id="editStartDate" value="${convertDateForInput(testData.projectInfo.startDate)}" 
                       style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                              border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                              color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                              box-sizing: border-box; transition: border-color 0.3s ease;"
                       onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                       onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'" />
            </div>
            <div>
                <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                    Fecha Fin:
                </label>
                <input type="date" id="editEndDate" value="${convertDateForInput(testData.projectInfo.endDate)}" 
                       style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                              border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                              color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                              box-sizing: border-box; transition: border-color 0.3s ease;"
                       onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                       onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'" />
            </div>
        </div>
        
        <div style="margin-bottom: 1rem;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                Progreso (%):
            </label>
            <input type="number" id="editProgress" value="${testData.projectInfo.progress || 0}" min="0" max="100"
                   style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                          border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                          color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                          box-sizing: border-box; transition: border-color 0.3s ease;"
                   onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                   onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'" />
        </div>
        
        <div style="margin-bottom: 1.5rem;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 500; font-size: 14px; color: inherit;">
                Estado:
            </label>
            <select id="editStatus" style="width: 100%; padding: 0.75rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                                           border-radius: 6px; background: ${isDarkMode ? '#0f172a' : '#ffffff'}; 
                                           color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; font-size: 14px; 
                                           box-sizing: border-box; transition: border-color 0.3s ease;"
                    onfocus="this.style.borderColor='#3b82f6'; this.style.boxShadow='0 0 0 3px rgba(59, 130, 246, 0.1)'"
                    onblur="this.style.borderColor='${isDarkMode ? '#475569' : '#d1d5db'}'; this.style.boxShadow='none'">
                <option value="Planificación" ${testData.projectInfo.status === 'Planificación' ? 'selected' : ''}>Planificación</option>
                <option value="Análisis" ${testData.projectInfo.status === 'Análisis' ? 'selected' : ''}>Análisis</option>
                <option value="Diseño" ${testData.projectInfo.status === 'Diseño' ? 'selected' : ''}>Diseño</option>
                <option value="Preparación" ${testData.projectInfo.status === 'Preparación' ? 'selected' : ''}>Preparación</option>
                <option value="Pruebas QA" ${testData.projectInfo.status === 'Pruebas QA' ? 'selected' : ''}>Pruebas QA</option>
                <option value="Cierre" ${testData.projectInfo.status === 'Cierre' ? 'selected' : ''}>Cierre</option>
            </select>
        </div>
        
        <div style="display: flex; gap: 1rem; justify-content: flex-end; margin-top: 1.5rem;">
            <button type="button" onclick="closeModal()" 
                    style="padding: 0.75rem 1.5rem; border: 1px solid ${isDarkMode ? '#475569' : '#d1d5db'}; 
                           background: transparent; color: ${isDarkMode ? '#f8fafc' : '#1e293b'}; 
                           border-radius: 6px; cursor: pointer; font-size: 14px; 
                           transition: all 0.3s ease;"
                    onmouseover="this.style.background='${isDarkMode ? '#374151' : '#f3f4f6'}'"
                    onmouseout="this.style.background='transparent'">
                Cancelar
            </button>
            <button type="button" onclick="saveProjectInfo()" 
                    style="padding: 0.75rem 1.5rem; border: none; background: #3b82f6; color: white; 
                           border-radius: 6px; cursor: pointer; font-size: 14px; 
                           transition: all 0.3s ease;"
                    onmouseover="this.style.background='#2563eb'; this.style.transform='translateY(-1px)'"
                    onmouseout="this.style.background='#3b82f6'; this.style.transform='translateY(0)'">
                Guardar
            </button>
        </div>
    `;
    
    modal.appendChild(form);
    document.body.appendChild(modal);
    
    // Funciones del modal
    window.closeModal = function() {
        modal.remove();
        delete window.closeModal;
        delete window.saveProjectInfo;
    };
    
    window.saveProjectInfo = function() {
        updateProjectInfo();
    };
}

// Función para actualizar información del proyecto
function updateProjectInfo() {
    try {
        // Función para convertir fecha YYYY-MM-DD a DD/MM/YYYY
        function convertDateFromInput(dateStr) {
            if (!dateStr) return '';
            const parts = dateStr.split('-');
            if (parts.length !== 3) return dateStr;
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
        
        // Validar campos requeridos
        const projectName = document.getElementById('editProjectName').value;
        const qaResponsible = document.getElementById('editQAResponsible').value.trim();
        
        if (!projectName || projectName === '' || !qaResponsible) {
            showNotification('Por favor complete todos los campos requeridos', 'error');
            return;
        }
        
        // Actualizar datos del proyecto
        testData.projectInfo = {
            name: projectName,
            qaResponsible: qaResponsible,
            startDate: convertDateFromInput(document.getElementById('editStartDate').value),
            endDate: convertDateFromInput(document.getElementById('editEndDate').value),
            progress: parseInt(document.getElementById('editProgress').value) || 0,
            status: document.getElementById('editStatus').value
        };
        
        // Guardar en localStorage
        localStorage.setItem('qaTestData', JSON.stringify(testData));
        
        // Actualizar la vista
        initializeDashboard();
        
        // Cerrar modal
        if (window.closeModal) {
            window.closeModal();
        }
        
        // Mostrar confirmación
        showNotification('Información del proyecto actualizada correctamente', 'success');
        
    } catch (error) {
        console.error('Error actualizando información del proyecto:', error);
        showNotification('Error al actualizar la información del proyecto', 'error');
    }
}

// ========================================
// FUNCIONES PARA AMPLIFICAR GRÁFICOS
// ========================================

// Variable para guardar el gráfico actual en el modal
let currentExpandedChart = null;

// Función para expandir un gráfico específico
function expandChart(chartType, title) {
    const modal = document.getElementById('chartModal');
    const modalTitle = document.getElementById('chartModalTitle');
    const modalContainer = document.getElementById('chartModalContainer');
    
    // Limpiar el contenedor del modal
    modalContainer.innerHTML = '';
    
    // Establecer el título
    modalTitle.textContent = title;
    
    // Crear el contenedor para el gráfico ampliado
    const expandedChartDiv = document.createElement('div');
    expandedChartDiv.id = `expanded${chartType}Chart`;
    expandedChartDiv.style.width = '100%';
    expandedChartDiv.style.height = '100%';
    modalContainer.appendChild(expandedChartDiv);
    
    // Mostrar el modal
    modal.style.display = 'flex';
    document.body.style.overflow = 'hidden';
    
    // Obtener el gráfico original y copiarlo ampliado
    setTimeout(() => {
        copyAndExpandExistingChart(chartType, expandedChartDiv.id);
    }, 100);
}

// Función para copiar y ampliar gráficos existentes
function copyAndExpandExistingChart(chartType, containerId) {
    const container = document.getElementById(containerId);
    let sourceChartId, newLayout;
    
    // Mapear tipos de gráfico a sus IDs originales
    const chartIdMap = {
        'pie': 'pieChart',
        'trend': 'trendChart', 
        'category': 'categoryChart',
        'defects': 'defectsChart',
        'burndown': 'burndownChart',
        'coverage': 'coverageChart',
        'dailyExecution': 'dailyExecutionChart',
        'executorDistribution': 'executorDistributionChart',
        'deliverables': 'deliverablesChart'
    };
    
    sourceChartId = chartIdMap[chartType];
    const sourceElement = document.getElementById(sourceChartId);
    
    if (!sourceElement || !sourceElement.data || !sourceElement.layout) {
        console.error('No se pudo encontrar el gráfico fuente:', sourceChartId);
        return;
    }
    
    // Copiar datos del gráfico original
    const originalData = JSON.parse(JSON.stringify(sourceElement.data));
    const originalLayout = JSON.parse(JSON.stringify(sourceElement.layout));
    
    // Ajustar layout para el modal ampliado
    newLayout = {
        ...originalLayout,
        width: undefined,
        height: undefined,
        autosize: true,
        margin: { t: 60, b: 60, l: 60, r: 60 },
        paper_bgcolor: 'white',
        plot_bgcolor: 'white',
        font: {
            ...originalLayout.font,
            size: 16, // Aumentar el tamaño de fuente para mejor legibilidad
            color: '#374151' // Color de texto más oscuro para mejor contraste
        }
    };
    
    // Configuración específica para gráfico circular
    if (chartType === 'pie') {
        newLayout.showlegend = true;
        newLayout.legend = {
            orientation: 'v',
            x: 1.05,
            y: 0.5,
            font: { size: 14 }
        };
        // Hacer texto más grande en el gráfico circular
        if (originalData[0]) {
            originalData[0].textfont = { size: 16 };
        }
    }
    
    // Configuración específica para gráfico de tendencia
    if (chartType === 'trend') {
        newLayout.xaxis = {
            ...originalLayout.xaxis,
            tickfont: { size: 14 }
        };
        newLayout.yaxis = {
            ...originalLayout.yaxis,
            tickfont: { size: 14 }
        };
    }
    
    // Configuración específica para gráfico burndown
    if (chartType === 'burndown') {
        newLayout.xaxis = {
            ...originalLayout.xaxis,
            tickfont: { size: 14 }
        };
        newLayout.yaxis = {
            ...originalLayout.yaxis,
            tickfont: { size: 14 }
        };
        newLayout.legend = {
            ...originalLayout.legend,
            font: { size: 14 }
        };
    }
    
    // Configuración específica para gráfico de cobertura
    if (chartType === 'coverage') {
        newLayout.xaxis = {
            ...originalLayout.xaxis,
            tickfont: { size: 14 }
        };
        newLayout.yaxis = {
            ...originalLayout.yaxis,
            tickfont: { size: 14 }
        };
        newLayout.margin = { t: 60, b: 60, l: 150, r: 60 }; // Más espacio para etiquetas de módulos
    }
    
    // Configuración específica para gráfico de entregables
    if (chartType === 'deliverables') {
        newLayout.margin = { t: 60, b: 60, l: 250, r: 60 }; // Más espacio para nombres de entregables
        newLayout.height = 500; // Altura fija para mejor visualización
        newLayout.yaxis = {
            ...originalLayout.yaxis,
            tickfont: { size: 14 }
        };
    }
    
    // Crear el gráfico ampliado con los datos originales
    Plotly.newPlot(containerId, originalData, newLayout, {
        displayModeBar: true,
        displaylogo: false,
        modeBarButtonsToRemove: ['pan2d', 'lasso2d'],
        responsive: true
    });
}

// Crear gráfico circular ampliado
function createExpandedPieChart(container) {
    const data = testData.summary;
    
    const values = [data.successful, data.failed, data.pending];
    const labels = ['Exitosas', 'Fallidas', 'Pendientes'];
    const colors = ['#10B981', '#EF4444', '#F59E0B'];
    
    // Agregar planificadas si existe
    if (data.planned && data.planned > 0) {
        values.push(data.planned);
        labels.push('Planificadas');
        colors.push('#9CA3AF');
    }
    
    // Agregar bloqueadas si existe
    if (data.blocked && data.blocked > 0) {
        values.push(data.blocked);
        labels.push('Bloqueadas');
        colors.push('#6B7280');
    }
    
    const chartData = [{
        values: values,
        labels: labels,
        type: 'pie',
        marker: {
            colors: colors
        },
        textinfo: 'label+percent+value',
        textfont: {
            size: 16,
            color: '#374151'
        },
        hovertemplate: '<b>%{label}</b><br>' +
                      'Cantidad: %{value}<br>' +
                      'Porcentaje: %{percent}<br>' +
                      '<extra></extra>'
    }];

    const layout = {
        title: {
            text: 'Resumen de Resultados de Pruebas',
            font: { size: 24, color: '#374151' }
        },
        font: { size: 14 },
        showlegend: true,
        legend: {
            orientation: "h",
            yanchor: "bottom",
            y: -0.1,
            xanchor: "center",
            x: 0.5,
            font: { size: 16 }
        },
        margin: { t: 80, b: 80, l: 40, r: 40 },
        height: null,
        autosize: true
    };

    const config = {
        responsive: true,
        displayModeBar: true,
        displaylogo: false,
        modeBarButtonsToAdd: ['downloadSVG'],
        modeBarButtonsToRemove: ['autoScale2d', 'resetScale2d']
    };

    currentExpandedChart = Plotly.newPlot(container, chartData, layout, config);
}

// Crear gráfico de tendencias ampliado
function createExpandedTrendChart(container) {
    const categories = testData.categories.map(c => c.escenario);
    
    const plannedTrace = {
        x: categories,
        y: testData.categories.map(c => c.planned),
        name: 'Planificadas',
        type: 'bar',
        marker: { color: '#6B7280' },
        hovertemplate: '<b>%{x}</b><br>Planificadas: %{y}<extra></extra>'
    };

    const successfulTrace = {
        x: categories,
        y: testData.categories.map(c => c.successful),
        name: 'Exitosas',
        type: 'bar',
        marker: { color: '#10B981' },
        hovertemplate: '<b>%{x}</b><br>Exitosas: %{y}<extra></extra>'
    };

    const failedTrace = {
        x: categories,
        y: testData.categories.map(c => c.failed),
        name: 'Fallidas',
        type: 'bar',
        marker: { color: '#EF4444' },
        hovertemplate: '<b>%{x}</b><br>Fallidas: %{y}<extra></extra>'
    };

    const pendingTrace = {
        x: categories,
        y: testData.categories.map(c => c.pending),
        name: 'Pendientes',
        type: 'bar',
        marker: { color: '#F59E0B' },
        hovertemplate: '<b>%{x}</b><br>Pendientes: %{y}<extra></extra>'
    };

    const blockedTrace = {
        x: categories,
        y: testData.categories.map(c => c.blocked || 0),
        name: 'Bloqueadas',
        type: 'bar',
        marker: { color: '#6B7280' },
        hovertemplate: '<b>%{x}</b><br>Bloqueadas: %{y}<extra></extra>'
    };

    const layout = {
        title: {
            text: 'Pruebas por Escenario - Vista Detallada',
            font: { size: 24, color: '#374151' }
        },
        xaxis: { 
            title: 'Escenarios de Prueba',
            titlefont: { size: 16 },
            tickfont: { size: 14 }
        },
        yaxis: { 
            title: 'Cantidad de Pruebas',
            titlefont: { size: 16 },
            tickfont: { size: 14 }
        },
        barmode: 'group',
        legend: {
            orientation: "h",
            yanchor: "bottom",
            y: -0.2,
            xanchor: "center",
            x: 0.5,
            font: { size: 16 }
        },
        margin: { t: 80, b: 120, l: 60, r: 40 },
        height: null,
        autosize: true
    };

    const config = {
        responsive: true,
        displayModeBar: true,
        displaylogo: false,
        modeBarButtonsToAdd: ['downloadSVG'],
        modeBarButtonsToRemove: ['autoScale2d', 'resetScale2d']
    };

    currentExpandedChart = Plotly.newPlot(container, [plannedTrace, successfulTrace, failedTrace, pendingTrace, blockedTrace], layout, config);
}

// Crear gráfico de categorías ampliado (mismo que trend pero con diferente título)
function createExpandedCategoryChart(container) {
    createExpandedTrendChart(container);
    // Actualizar el título después de crear el gráfico
    setTimeout(() => {
        Plotly.relayout(container, {
            'title.text': 'Distribución de Pruebas por Escenario'
        });
    }, 100);
}

// Crear gráfico de defectos ampliado
function createExpandedDefectsChart(container) {
    const defectsData = testData.defects.summary;
    
    const chartData = [{
        values: [defectsData.critical, defectsData.high, defectsData.medium, defectsData.low],
        labels: ['Críticos', 'Altos', 'Medios', 'Bajos'],
        type: 'pie',
        marker: {
            colors: ['#DC2626', '#F59E0B', '#3B82F6', '#10B981']
        },
        textinfo: 'label+percent+value',
        textfont: {
            size: 16,
            color: '#374151'
        },
        hovertemplate: '<b>%{label}</b><br>' +
                      'Cantidad: %{value}<br>' +
                      'Porcentaje: %{percent}<br>' +
                      '<extra></extra>'
    }];

    const layout = {
        title: {
            text: 'Distribución de Defectos por Severidad',
            font: { size: 24, color: '#374151' }
        },
        font: { size: 14 },
        showlegend: true,
        legend: {
            orientation: "h",
            yanchor: "bottom",
            y: -0.1,
            xanchor: "center",
            x: 0.5,
            font: { size: 16 }
        },
        margin: { t: 80, b: 80, l: 40, r: 40 },
        height: null,
        autosize: true
    };

    const config = {
        responsive: true,
        displayModeBar: true,
        displaylogo: false,
        modeBarButtonsToAdd: ['downloadSVG'],
        modeBarButtonsToRemove: ['autoScale2d', 'resetScale2d']
    };

    currentExpandedChart = Plotly.newPlot(container, chartData, layout, config);
}

// Función para cerrar el modal de gráficos
function closeChartModal(event) {
    // Si se hace click en el overlay (no en el contenido), cerrar el modal
    if (event && event.target.id !== 'chartModal') {
        return;
    }
    
    const modal = document.getElementById('chartModal');
    const modalContainer = document.getElementById('chartModalContainer');
    
    // Limpiar el gráfico
    if (currentExpandedChart) {
        currentExpandedChart = null;
    }
    
    // Limpiar el contenedor
    modalContainer.innerHTML = '';
    
    // Ocultar el modal
    modal.style.display = 'none';
    document.body.style.overflow = 'auto';
}

// Manejar el redimensionamiento de ventana para gráficos ampliados
window.addEventListener('resize', function() {
    if (currentExpandedChart) {
        setTimeout(() => {
            const modalContainer = document.getElementById('chartModalContainer');
            if (modalContainer.firstChild) {
                Plotly.Plots.resize(modalContainer.firstChild);
            }
        }, 100);
    }
});

// Manejar el cierre del modal con la tecla ESC
document.addEventListener('keydown', function(event) {
    if (event.key === 'Escape') {
        closeChartModal();
    }
});

// ========================================
// FUNCIONES PARA GENERAR PDF
// ========================================

// Función de manejo para verificar librerías antes de generar PDF
function handlePDFGeneration() {
    if (!checkPDFLibraries()) {
        showNotification('Las librerías necesarias para PDF no están disponibles. Verifique su conexión a internet.', 'error');
        return;
    }
    
    generatePDFReport();
}

// Función principal para generar el PDF exactamente como el dashboard
async function generatePDFReport() {
    // Mostrar mensaje de carga
    showNotification('Generando reporte PDF...', 'info');
    
    try {
        // Verificar si las librerías están disponibles
        if (!window.jspdf) {
            throw new Error('jsPDF no está cargado. Verifique la conexión a internet.');
        }
        
        if (!window.html2canvas) {
            throw new Error('html2canvas no está cargado. Verifique la conexión a internet.');
        }
        
        // Crear instancia de jsPDF
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        
        // Activar modo PDF (ocultar elementos innecesarios)
        document.body.classList.add('pdf-mode');
        
        // Esperar un momento para que se apliquen los estilos
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Configuración de la página
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 10;
        const contentWidth = pageWidth - (margin * 2);
        let currentY = margin;
        
        console.log('📄 Iniciando generación PDF del dashboard...');
        
        // === HEADER ===
        await addPDFHeader(pdf, currentY);
        currentY += 25;
        
        // 1. CAPTURAR AVANCE GENERAL (Información del Proyecto)
        console.log('📸 Capturando Avance General...');
        const projectSection = document.querySelector('.project-section');
        if (projectSection) {
            try {
                const projectCanvas = await html2canvas(projectSection, {
                    scale: 2,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const projectImgData = projectCanvas.toDataURL('image/png');
                const projectImgWidth = contentWidth;
                const projectImgHeight = (projectCanvas.height * projectImgWidth) / projectCanvas.width;
                
                // Verificar si cabe en la página actual
                if (currentY + projectImgHeight > pageHeight - margin) {
                    pdf.addPage();
                    currentY = margin;
                }
                
                pdf.addImage(projectImgData, 'PNG', margin, currentY, projectImgWidth, projectImgHeight);
                currentY += projectImgHeight + 5;
                console.log('✅ Avance General capturado');
            } catch (error) {
                console.warn('⚠️ Error capturando Avance General:', error);
            }
        }
        
        // 2. CAPTURAR TARJETAS KPI
        console.log('📸 Capturando tarjetas KPI...');
        const kpiSection = document.querySelector('.kpi-section');
        if (kpiSection) {
            try {
                const kpiCanvas = await html2canvas(kpiSection, {
                    scale: 2,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const kpiImgData = kpiCanvas.toDataURL('image/png');
                const kpiImgWidth = contentWidth;
                const kpiImgHeight = (kpiCanvas.height * kpiImgWidth) / kpiCanvas.width;
                
                // Verificar si cabe en la página actual
                if (currentY + kpiImgHeight > pageHeight - margin) {
                    pdf.addPage();
                    currentY = margin;
                }
                
                pdf.addImage(kpiImgData, 'PNG', margin, currentY, kpiImgWidth, kpiImgHeight);
                currentY += kpiImgHeight + 5;
                console.log('✅ Tarjetas KPI capturadas');
            } catch (error) {
                console.warn('⚠️ Error capturando tarjetas KPI:', error);
            }
        }
        
        // 3. CAPTURAR PORCENTAJES DE AVANCE INDIVIDUAL
        console.log('📸 Capturando Porcentajes de Avance Individual...');
        const individualProgressSection = document.querySelector('.individual-progress-section');
        if (individualProgressSection) {
            try {
                const progressCanvas = await html2canvas(individualProgressSection, {
                    scale: 2,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const progressImgData = progressCanvas.toDataURL('image/png');
                const progressImgWidth = contentWidth;
                const progressImgHeight = (progressCanvas.height * progressImgWidth) / progressCanvas.width;
                
                // Verificar si cabe en la página actual
                if (currentY + progressImgHeight > pageHeight - margin) {
                    pdf.addPage();
                    currentY = margin;
                }
                
                pdf.addImage(progressImgData, 'PNG', margin, currentY, progressImgWidth, progressImgHeight);
                currentY += progressImgHeight + 5;
                console.log('✅ Porcentajes de Avance Individual capturados');
            } catch (error) {
                console.warn('⚠️ Error capturando Porcentajes de Avance Individual:', error);
            }
        }
        
        // 4. CAPTURAR TABLA DE DETALLE DE PROGRESO
        console.log('📸 Capturando tabla de Detalle de Progreso...');
        const tableSection = document.querySelector('.table-section');
        if (tableSection) {
            try {
                // Guardar la paginación actual
                const paginationContainer = tableSection.querySelector('.pagination-container');
                const originalPaginationDisplay = paginationContainer ? paginationContainer.style.display : '';
                
                // Ocultar controles de paginación para el PDF
                if (paginationContainer) {
                    paginationContainer.style.display = 'none';
                }
                
                const tableCanvas = await html2canvas(tableSection, {
                    scale: 1.5,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                // Restaurar controles de paginación
                if (paginationContainer) {
                    paginationContainer.style.display = originalPaginationDisplay;
                }
                
                const tableImgData = tableCanvas.toDataURL('image/png');
                const tableImgWidth = contentWidth;
                const tableImgHeight = (tableCanvas.height * tableImgWidth) / tableCanvas.width;
                
                // Verificar si cabe en la página actual
                if (currentY + tableImgHeight > pageHeight - margin) {
                    pdf.addPage();
                    currentY = margin;
                }
                
                // Si la tabla es muy alta, dividirla en múltiples páginas
                if (tableImgHeight > pageHeight - margin * 2) {
                    const maxHeight = pageHeight - margin * 2;
                    let remainingHeight = tableImgHeight;
                    let sourceY = 0;
                    
                    while (remainingHeight > 0) {
                        const heightToAdd = Math.min(remainingHeight, maxHeight);
                        const sourceHeight = (heightToAdd * tableCanvas.height) / tableImgHeight;
                        
                        // Crear un canvas temporal para esta sección
                        const tempCanvas = document.createElement('canvas');
                        tempCanvas.width = tableCanvas.width;
                        tempCanvas.height = sourceHeight;
                        const tempCtx = tempCanvas.getContext('2d');
                        tempCtx.drawImage(tableCanvas, 0, sourceY, tableCanvas.width, sourceHeight, 0, 0, tableCanvas.width, sourceHeight);
                        
                        const tempImgData = tempCanvas.toDataURL('image/png');
                        pdf.addImage(tempImgData, 'PNG', margin, margin, tableImgWidth, heightToAdd);
                        
                        remainingHeight -= heightToAdd;
                        sourceY += sourceHeight;
                        
                        if (remainingHeight > 0) {
                            pdf.addPage();
                        }
                    }
                    
                    currentY = margin + (tableImgHeight % maxHeight);
                } else {
                    pdf.addImage(tableImgData, 'PNG', margin, currentY, tableImgWidth, tableImgHeight);
                    currentY += tableImgHeight;
                }
                
                console.log('✅ Tabla de Detalle de Progreso capturada');
            } catch (error) {
                console.warn('⚠️ Error capturando tabla de Detalle de Progreso:', error);
            }
        }
        
        // Agregar footer en todas las páginas
        addPDFFooter(pdf);
        
        // Desactivar modo PDF
        document.body.classList.remove('pdf-mode');
        
        // Generar nombre del archivo
        const currentDate = new Date().toLocaleDateString('es-ES').replace(/\//g, '-');
        const fileName = `Dashboard_QA_Diseño_Casos_Prueba_${currentDate}.pdf`;
        
        console.log('💾 Guardando PDF:', fileName);
        console.log('📊 PDF generado con todas las secciones del dashboard');
        
        // Descargar el PDF
        pdf.save(fileName);
        
        showNotification('PDF generado exitosamente', 'success');
        
    } catch (error) {
        console.error('Error generando PDF:', error);
        showNotification(`Error al generar PDF: ${error.message}`, 'error');
        document.body.classList.remove('pdf-mode');
    }
}

// Función auxiliar para capturar gráficos individuales de forma compacta
async function captureIndividualChartsCompact(pdf, margin, contentWidth, pageHeight, startY) {
    console.log('📸 Capturando gráficos individuales compactos...');
    
    const specificChartIds = ['pieChart', 'trendChart', 'categoryChart', 'defectsChart', 'burndownChart', 'coverageChart'];
    
    let chartsPerRow = 3; // 3 gráficos por fila para ser más compacto
    let chartIndex = 0;
    let currentY = startY;
    
    // Calcular tamaños más pequeños para que quepan todos
    const availableHeight = pageHeight - currentY - margin - 10;
    const chartWidth = (contentWidth - 10) / 3; // Dividir en 3 columnas
    const maxChartHeight = availableHeight / 2; // Máximo 2 filas
    
    for (const chartId of specificChartIds) {
        const chartElement = document.getElementById(chartId);
        
        if (chartElement) {
            try {
                const chartCanvas = await html2canvas(chartElement, {
                    scale: 1, // Escala reducida para ser más compacto
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const chartImgData = chartCanvas.toDataURL('image/png');
                const originalChartHeight = (chartCanvas.height * chartWidth) / chartCanvas.width;
                const chartHeight = Math.min(originalChartHeight, maxChartHeight);
                
                // Calcular posición en grid 3x2
                const col = chartIndex % chartsPerRow;
                const row = Math.floor(chartIndex / chartsPerRow);
                
                const x = margin + (col * (chartWidth + 5));
                const y = currentY + (row * (chartHeight + 8));
                
                pdf.addImage(chartImgData, 'PNG', x, y, chartWidth, chartHeight);
                chartIndex++;
                
            } catch (error) {
                console.warn(`⚠️ Error capturando gráfico ${chartId}:`, error);
            }
        }
    }
    
    console.log(`✅ ${chartIndex} gráficos capturados en formato compacto`);
}

// Función auxiliar para capturar gráficos individuales (fallback original)
async function captureIndividualCharts(pdf, margin, contentWidth, pageHeight) {
    console.log('📸 Capturando gráficos individuales...');
    
    const specificChartIds = ['pieChart', 'trendChart', 'categoryChart', 'defectsChart', 'burndownChart', 'coverageChart'];
    
    let chartsPerRow = 2;
    let chartIndex = 0;
    let currentY = margin + 40;
    
    for (const chartId of specificChartIds) {
        const chartElement = document.getElementById(chartId);
        
        if (chartElement) {
            try {
                const chartCanvas = await html2canvas(chartElement, {
                    scale: 1.5,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });
                
                const chartImgData = chartCanvas.toDataURL('image/png');
                const chartWidth = (contentWidth - 5) / 2;
                const chartHeight = (chartCanvas.height * chartWidth) / chartCanvas.width;
                
                const isLeftColumn = (chartIndex % chartsPerRow) === 0;
                const x = isLeftColumn ? margin : margin + chartWidth + 5;
                
                if (currentY + chartHeight > pageHeight - margin) {
                    pdf.addPage();
                    currentY = margin + 20;
                    chartIndex = 0;
                }
                
                const finalY = (chartIndex % chartsPerRow) === 0 ? currentY : currentY;
                const finalX = isLeftColumn ? margin : margin + chartWidth + 5;
                
                pdf.addImage(chartImgData, 'PNG', finalX, finalY, chartWidth, chartHeight);
                
                if ((chartIndex % chartsPerRow) === 1) {
                    currentY += chartHeight + 10;
                }
                
                chartIndex++;
                
            } catch (error) {
                console.warn(`⚠️ Error capturando gráfico ${chartId}:`, error);
            }
        }
    }
}

// Función para agregar el header del PDF
async function addPDFHeader(pdf, startY) {
    // Logo y título principal
    pdf.setFillColor(236, 0, 0); // Rojo Santander
    pdf.rect(0, 0, pdf.internal.pageSize.getWidth(), 30, 'F');
    
    // Título principal
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(18);
    pdf.setTextColor(255, 255, 255);
    pdf.text('Santander - Dashboard QA', 15, 15);
    
    // Subtítulo
    pdf.setFontSize(12);
    pdf.setFont('helvetica', 'normal');
    pdf.text('Reporte de Ejecución Diaria de Pruebas', 15, 22);
    
    // Fecha de generación
    const currentDate = new Date().toLocaleDateString('es-ES', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(10);
    pdf.text(`Generado: ${currentDate}`, pdf.internal.pageSize.getWidth() - 15, 22, { align: 'right' });
}

// Función para agregar información del proyecto
async function addProjectInfo(pdf, startY, contentWidth) {
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(14);
    pdf.text('Información del Proyecto', 15, startY);
    
    // Línea separadora
    pdf.setDrawColor(236, 0, 0);
    pdf.setLineWidth(0.5);
    pdf.line(15, startY + 2, 15 + contentWidth, startY + 2);
    
    startY += 15;
    
    // Layout más compacto - etiqueta:valor en la misma línea
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    
    // Dos columnas principales con separación mínima
    const leftColumnStart = 15;
    const rightColumnStart = 300; // Posición fija para segunda columna
    
    // Fila 1: Proyecto y Responsable
    pdf.setFont('helvetica', 'bold');
    pdf.text('Proyecto: ', leftColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    // Calcular posición después del label
    const proyectoLabelWidth = pdf.getTextWidth('Proyecto: ');
    const projectName = testData.projectInfo.name.length > 20 ? 
                       testData.projectInfo.name.substring(0, 20) + '...' : 
                       testData.projectInfo.name;
    pdf.text(projectName, leftColumnStart + proyectoLabelWidth, startY);
    
    pdf.setFont('helvetica', 'bold');
    pdf.text('Responsable: ', rightColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    const responsableLabelWidth = pdf.getTextWidth('Responsable: ');
    pdf.text(testData.projectInfo.qaResponsible, rightColumnStart + responsableLabelWidth, startY);
    
    startY += 12;
    
    // Fila 2: Fechas
    pdf.setFont('helvetica', 'bold');
    pdf.text('Inicio: ', leftColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    const inicioLabelWidth = pdf.getTextWidth('Inicio: ');
    pdf.text(testData.projectInfo.startDate, leftColumnStart + inicioLabelWidth, startY);
    
    pdf.setFont('helvetica', 'bold');
    pdf.text('Cierre: ', rightColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    const cierreLabelWidth = pdf.getTextWidth('Cierre: ');
    pdf.text(testData.projectInfo.endDate, rightColumnStart + cierreLabelWidth, startY);
    
    startY += 12;
    
    // Fila 3: Estado y Progreso compacto
    pdf.setFont('helvetica', 'bold');
    pdf.text('Estado: ', leftColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    const estadoLabelWidth = pdf.getTextWidth('Estado: ');
    pdf.text(testData.projectInfo.status, leftColumnStart + estadoLabelWidth, startY);
    
    const realProgress = calculateProjectProgress();
    const plannedProgress = calculatePlannedProgress();
    
    pdf.setFont('helvetica', 'bold');
    pdf.text('Progreso: ', rightColumnStart, startY);
    pdf.setFont('helvetica', 'normal');
    const progresoLabelWidth = pdf.getTextWidth('Progreso: ');
    pdf.text(`${plannedProgress.toFixed(1)}% / ${realProgress.toFixed(1)}%`, rightColumnStart + progresoLabelWidth, startY);
}

// Función para agregar resumen de KPIs
async function addKPISummary(pdf, startY, contentWidth) {
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(14);
    pdf.text('Resumen de Resultados', 15, startY);
    
    // Línea separadora
    pdf.setDrawColor(236, 0, 0);
    pdf.line(15, startY + 2, 15 + contentWidth, startY + 2);
    
    startY += 12; // Reducido de 15 a 12
    
    // KPIs más compactos para formato A4
    const totalKPIs = 5;
    const kpiSpacing = 3; // Reducido de 5 a 3
    const totalSpacing = kpiSpacing * (totalKPIs - 1);
    const kpiWidth = (contentWidth - totalSpacing) / totalKPIs;
    const kpiHeight = 18; // Reducido de 20 a 18
    
    const kpis = [
        { 
            label: 'Planificadas', 
            value: testData.summary.planned, 
            color: [245, 158, 11],
            icon: '◷'
        },
        { 
            label: 'Exitosas', 
            value: testData.summary.successful, 
            color: [16, 185, 129],
            icon: '✓'
        },
        { 
            label: 'Fallidas', 
            value: testData.summary.failed, 
            color: [239, 68, 68],
            icon: '✗'
        },
        { 
            label: 'Pendientes', 
            value: testData.summary.pending, 
            color: [139, 92, 246],
            icon: '○'
        },
        { 
            label: 'Bloqueadas', 
            value: testData.summary.blocked || 0, 
            color: [107, 114, 128],
            icon: '⬛'
        }
    ];
    
    let x = 15;
    const y = startY;
    
    kpis.forEach((kpi, index) => {
        // Fondo del KPI con color más claro
        const lightColor = kpi.color.map(c => Math.min(255, c + (255 - c) * 0.85));
        pdf.setFillColor(lightColor[0], lightColor[1], lightColor[2]);
        pdf.rect(x, y, kpiWidth, kpiHeight, 'F');
        
        // Borde
        pdf.setDrawColor(kpi.color[0], kpi.color[1], kpi.color[2]);
        pdf.setLineWidth(0.3);
        pdf.rect(x, y, kpiWidth, kpiHeight);
        
        // Texto del valor (número) - centrado y más compacto
        pdf.setTextColor(kpi.color[0], kpi.color[1], kpi.color[2]);
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(11); // Reducido de 12 a 11
        const valueText = kpi.value.toString();
        const valueWidth = pdf.getTextWidth(valueText);
        pdf.text(valueText, x + (kpiWidth - valueWidth) / 2, y + 8); // Ajustado de 9 a 8
        
        // Texto del label - centrado y más pequeño
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(7); // Reducido de 8 a 7
        pdf.setTextColor(0, 0, 0);
        const labelWidth = pdf.getTextWidth(kpi.label);
        pdf.text(kpi.label, x + (kpiWidth - labelWidth) / 2, y + 15); // Ajustado de 17 a 15
        
        x += kpiWidth + kpiSpacing;
    });
}

// Función para agregar gráficos al PDF
async function addChartsToPDF(pdf, startY, contentWidth) {
    try {
        // Verificar si Plotly está disponible
        if (!window.Plotly) {
            console.warn('Plotly no está disponible, omitiendo gráficos');
            pdf.setFont('helvetica', 'normal');
            pdf.setFontSize(12);
            pdf.setTextColor(100, 100, 100);
            pdf.text('Los gráficos no están disponibles en este reporte.', 15, startY + 20);
            return;
        }
        
        const chartIds = ['pieChart', 'trendChart', 'categoryChart', 'defectsChart', 'burndownChart'];
        const chartTitles = [
            'Distribución de Resultados',
            'Tendencia de Ejecución',
            'Pruebas por Escenario',
            'Gestión de Defectos',
            'Cycle Time de Defectos'
        ];
        const chartsPerRow = 2;
        const chartWidth = (contentWidth - 10) / chartsPerRow;
        const chartHeight = 60;
        const titleHeight = 8; // Espacio para títulos
        
        let currentX = 15;
        let currentY = startY;
        let chartsAdded = 0;
        
        console.log('📊 Capturando gráficos para PDF usando Plotly.toImage...');
        
        for (let i = 0; i < chartIds.length; i++) {
            const chartElement = document.getElementById(chartIds[i]);
            
            if (chartElement && chartElement.querySelector('.plotly-graph-div')) {
                try {
                    console.log(`📈 Capturando gráfico: ${chartIds[i]}`);
                    
                    // Agregar título del gráfico
                    pdf.setFont('helvetica', 'bold');
                    pdf.setFontSize(9);
                    pdf.setTextColor(0, 0, 0);
                    pdf.text(chartTitles[i], currentX + chartWidth/2, currentY + 5, { align: 'center' });
                    
                    // Usar Plotly.toImage para exportar directamente
                    const imgData = await Plotly.toImage(chartElement, {
                        format: 'png',
                        width: chartWidth * 4, // Multiplicar por factor de escala
                        height: chartHeight * 4,
                        scale: 1
                    });
                    
                    // Agregar al PDF (con offset para el título)
                    pdf.addImage(imgData, 'PNG', currentX, currentY + titleHeight, chartWidth, chartHeight);
                    chartsAdded++;
                    
                    console.log(`✅ Gráfico ${chartIds[i]} agregado al PDF`);
                    
                    // Calcular posición para el siguiente gráfico
                    if ((chartsAdded) % chartsPerRow === 0) {
                        // Nueva fila
                        currentX = 15;
                        currentY += chartHeight + titleHeight + 15; // Espacio para título + gráfico + separación
                        
                        // Verificar si necesitamos nueva página
                        if (currentY + chartHeight + titleHeight > pdf.internal.pageSize.getHeight() - 30) {
                            pdf.addPage();
                            currentY = 15;
                        }
                    } else {
                        // Siguiente columna
                        currentX += chartWidth + 5;
                    }
                    
                } catch (chartError) {
                    console.warn(`⚠️ Error capturando gráfico ${chartIds[i]}:`, chartError);
                    
                    // Agregar título incluso para placeholder
                    pdf.setFont('helvetica', 'bold');
                    pdf.setFontSize(9);
                    pdf.setTextColor(0, 0, 0);
                    pdf.text(chartTitles[i], currentX + chartWidth/2, currentY + 5, { align: 'center' });
                    
                    // Agregar placeholder para el gráfico fallido
                    pdf.setDrawColor(200, 200, 200);
                    pdf.setFillColor(245, 245, 245);
                    pdf.rect(currentX, currentY + titleHeight, chartWidth, chartHeight, 'FD');
                    
                    pdf.setFont('helvetica', 'normal');
                    pdf.setFontSize(10);
                    pdf.setTextColor(100, 100, 100);
                    pdf.text('Gráfico no disponible', currentX + chartWidth/2, currentY + titleHeight + chartHeight/2, { align: 'center' });
                    
                    // Continuar con el siguiente gráfico
                    if ((i + 1) % chartsPerRow === 0) {
                        currentX = 15;
                        currentY += chartHeight + titleHeight + 15;
                        if (currentY + chartHeight + titleHeight > pdf.internal.pageSize.getHeight() - 30) {
                            pdf.addPage();
                            currentY = 15;
                        }
                    } else {
                        currentX += chartWidth + 5;
                    }
                }
            } else {
                console.warn(`⚠️ Gráfico ${chartIds[i]} no encontrado o no tiene datos`);
            }
        }
        
        console.log(`📊 Total de gráficos agregados al PDF: ${chartsAdded}/${chartIds.length}`);
        
    } catch (error) {
        console.error('❌ Error general agregando gráficos al PDF:', error);
        // Agregar mensaje de error en el PDF
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(12);
        pdf.setTextColor(200, 0, 0);
        pdf.text('Error cargando gráficos. Consulte la consola para más detalles.', 15, startY + 20);
    }
}

// Función para agregar tablas al PDF
async function addTablesToPDF(pdf, startY, contentWidth) {
    // Tabla de defectos (simplificada)
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(12);
    pdf.text('Resumen de Defectos por Severidad', 15, startY);
    
    startY += 10;
    
    // Headers de la tabla
    const headers = ['Severidad', 'Cantidad', 'Estado'];
    const colWidths = [40, 30, 40];
    let tableX = 15;
    
    // Header
    pdf.setFillColor(240, 240, 240);
    pdf.rect(tableX, startY, colWidths.reduce((a, b) => a + b, 0), 8, 'F');
    
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(9);
    pdf.setTextColor(0, 0, 0);
    
    let currentX = tableX + 2;
    headers.forEach((header, index) => {
        pdf.text(header, currentX, startY + 5);
        currentX += colWidths[index];
    });
    
    // Datos de defectos
    const defectRows = [
        ['Crítica', testData.defects.summary.critical.toString(), 'Requiere atención'],
        ['Alta', testData.defects.summary.high.toString(), 'Importante'],
        ['Media', testData.defects.summary.medium.toString(), 'Moderada'],
        ['Baja', testData.defects.summary.low.toString(), 'Menor']
    ];
    
    startY += 8;
    pdf.setFont('helvetica', 'normal');
    
    defectRows.forEach((row, rowIndex) => {
        currentX = tableX + 2;
        row.forEach((cell, cellIndex) => {
            pdf.text(cell, currentX, startY + 5);
            currentX += colWidths[cellIndex];
        });
        
        // Línea separadora
        pdf.setDrawColor(200, 200, 200);
        pdf.line(tableX, startY + 8, tableX + colWidths.reduce((a, b) => a + b, 0), startY + 8);
        startY += 8;
    });
    
    // Estadísticas adicionales
    startY += 10;
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.text('Estadísticas del Proyecto:', 15, startY);
    
    startY += 8;
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(9);
    
    const totalTests = testData.summary.successful + testData.summary.failed + 
                      testData.summary.pending + testData.summary.planned + 
                      (testData.summary.blocked || 0);
    
    const stats = [
        `Total de Pruebas: ${totalTests}`,
        `Tasa de Éxito: ${((testData.summary.successful / totalTests) * 100).toFixed(1)}%`,
        `Tasa de Falla: ${((testData.summary.failed / totalTests) * 100).toFixed(1)}%`,
        `Defectos Críticos/Altos: ${testData.defects.summary.critical + testData.defects.summary.high}`
    ];
    
    stats.forEach(stat => {
        pdf.text(`• ${stat}`, 20, startY);
        startY += 6;
    });
}

// Función para agregar títulos de sección
async function addPDFSectionTitle(pdf, title, y) {
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(14);
    pdf.setTextColor(236, 0, 0);
    pdf.text(title, 15, y);
    
    // Línea decorativa
    pdf.setDrawColor(236, 0, 0);
    pdf.setLineWidth(0.5);
    pdf.line(15, y + 2, 15 + pdf.getTextWidth(title), y + 2);
}

// Función para agregar footer a todas las páginas
function addPDFFooter(pdf) {
    const totalPages = pdf.internal.getNumberOfPages();
    
    for (let i = 1; i <= totalPages; i++) {
        pdf.setPage(i);
        
        // Footer
        const footerY = pdf.internal.pageSize.getHeight() - 15;
        
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(8);
        pdf.setTextColor(100, 100, 100);
        
        // Copyright
        pdf.text('© Santander Consumer Bank - Aseguramiento de Calidad', 15, footerY);
        
        // Número de página
        pdf.text(`Página ${i} de ${totalPages}`, pdf.internal.pageSize.getWidth() - 15, footerY, { align: 'right' });
        
        // Información de contacto
        pdf.text('Powered by Equipo QA | Support: arivera_scb@santander.com.pe', 15, footerY + 5);
    }
}

// Verificar que las librerías PDF estén disponibles
function checkPDFLibraries() {
    const libraries = {
        jsPDF: window.jspdf,
        html2canvas: window.html2canvas
    };
    
    const missing = Object.keys(libraries).filter(lib => !libraries[lib]);
    
    if (missing.length > 0) {
        console.warn('⚠️ Librerías PDF faltantes:', missing);
        return false;
    }
    
    console.log('✅ Librerías PDF disponibles');
    return true;
}

// Verificar librerías cuando se carga la página
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(() => {
        checkPDFLibraries();
    }, 2000); // Verificar después de 2 segundos para asegurar que las librerías se carguen
    
    // Verificar si hay datos en la URL al cargar la página
    loadDataFromURL();
});

// ==========================================
// FUNCIONALIDAD DE COMPARTIR ENLACE
// ==========================================

// Variable global para almacenar los datos cargados
let currentExcelData = null;

// Función para generar enlace compartible
async function generateShareableLink() {
    if (!currentExcelData) {
        alert('No hay datos cargados para compartir. Por favor, carga un archivo Excel primero.');
        return;
    }
    
    try {
        const baseUrl = window.location.origin + window.location.pathname;
        
        // Opción 1: Intentar crear enlace corto local
        const shortId = storeShortLink(currentExcelData);
        let shareableUrl;
        let displayUrl;
        let isShortLink = false;
        
        if (shortId) {
            // Crear enlace corto usando ID local
            shareableUrl = `${baseUrl}?id=${shortId}`;
            displayUrl = shareableUrl;
            isShortLink = true;
            console.log('✅ Enlace corto generado con ID:', shortId);
        } else {
            // Fallback: usar método original comprimido
            const compressedData = compressData(currentExcelData);
            shareableUrl = `${baseUrl}?data=${encodeURIComponent(compressedData)}`;
            displayUrl = shortenUrl(shareableUrl);
            console.log('⚠️ Usando enlace largo como fallback');
        }
        
        console.log('📊 Enlace generado - Longitud:', shareableUrl.length);
        
        // Mostrar el modal con el enlace
        document.getElementById('shareableLink').value = shareableUrl;
        
        // Mostrar versión apropiada en el display
        const displayElement = document.getElementById('shareableLinkDisplay');
        displayElement.textContent = displayUrl;
        
        // Actualizar información sobre el enlace
        const urlNote = document.querySelector('.url-note');
        if (urlNote) {
            if (isShortLink) {
                urlNote.textContent = `(enlace corto - ${shareableUrl.length} caracteres)`;
                urlNote.style.color = '#16a34a'; // Verde para indicar éxito
            } else {
                urlNote.textContent = `(enlace largo acortado para visualización - ${shareableUrl.length} caracteres)`;
                urlNote.style.color = '#ea580c'; // Naranja para advertencia
            }
        }
        
        // Opción 2: Intentar acortar con servicio externo si el enlace sigue siendo muy largo
        if (!isShortLink && shareableUrl.length > 2000) {
            try {
                const externalShortUrl = await generateShortUrl(shareableUrl);
                if (externalShortUrl !== shareableUrl && externalShortUrl.length < shareableUrl.length) {
                    document.getElementById('shareableLink').value = externalShortUrl;
                    displayElement.textContent = externalShortUrl;
                    
                    if (urlNote) {
                        urlNote.textContent = `(enlace acortado externamente - ${externalShortUrl.length} caracteres)`;
                        urlNote.style.color = '#16a34a';
                    }
                    
                    console.log('✅ Enlace acortado externamente:', externalShortUrl);
                }
            } catch (error) {
                console.log('No se pudo acortar externamente:', error.message);
            }
        }
        
        document.getElementById('shareModal').style.display = 'flex';
        
    } catch (error) {
        console.error('❌ Error al generar enlace compartible:', error);
        alert('Error al generar el enlace compartible. Los datos pueden ser demasiado grandes.');
    }
}

// Función para acortar URLs largas para visualización
function shortenUrl(url) {
    const maxDisplayLength = 60;
    
    if (url.length <= maxDisplayLength) {
        return url;
    }
    
    // Para URLs muy largas, mostrar solo el dominio y una indicación de datos
    const urlObj = new URL(url);
    const baseUrl = `${urlObj.protocol}//${urlObj.hostname}${urlObj.pathname}`;
    
    if (urlObj.search) {
        // Si hay parámetros, mostrar una versión muy simplificada
        return `${baseUrl}?data=[${Math.round(url.length / 1000)}KB]`;
    }
    
    // Si la URL base es muy larga, acortarla también
    if (baseUrl.length > maxDisplayLength) {
        const start = baseUrl.substring(0, 30);
        const end = baseUrl.substring(baseUrl.length - 25);
        return `${start}...${end}`;
    }
    
    return baseUrl;
}

// Nueva función para generar enlace corto con servicio externo
async function generateShortUrl(longUrl) {
    try {
        // Intentar usar is.gd como servicio de acortado
        const response = await fetch(`https://is.gd/create.php?format=simple&url=${encodeURIComponent(longUrl)}`);
        
        if (response.ok) {
            const shortUrl = await response.text();
            if (shortUrl && !shortUrl.includes('Error')) {
                return shortUrl.trim();
            }
        }
    } catch (error) {
        console.log('No se pudo usar servicio de acortado externo:', error.message);
    }
    
    // Fallback: usar acortado local
    return shortenUrl(longUrl);
}

// Sistema de enlaces cortos locales
const SHORT_LINKS_STORAGE_KEY = 'dashboardShortLinks';

// Función para almacenar enlace largo y generar ID corto
function storeShortLink(data) {
    try {
        // Generar ID único corto
        const shortId = createShortHash(data);
        
        // Obtener enlaces almacenados
        const storedLinks = JSON.parse(localStorage.getItem(SHORT_LINKS_STORAGE_KEY) || '{}');
        
        // Almacenar el enlace con timestamp para limpieza automática
        storedLinks[shortId] = {
            data: data,
            created: Date.now(),
            accessed: Date.now()
        };
        
        // Limpiar enlaces antiguos (más de 30 días)
        const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
        Object.keys(storedLinks).forEach(key => {
            if (storedLinks[key].accessed < thirtyDaysAgo) {
                delete storedLinks[key];
            }
        });
        
        // Guardar en localStorage
        localStorage.setItem(SHORT_LINKS_STORAGE_KEY, JSON.stringify(storedLinks));
        
        return shortId;
    } catch (error) {
        console.error('Error storing short link:', error);
        return null;
    }
}

// Función para recuperar datos desde ID corto
function getDataFromShortId(shortId) {
    try {
        const storedLinks = JSON.parse(localStorage.getItem(SHORT_LINKS_STORAGE_KEY) || '{}');
        
        if (storedLinks[shortId]) {
            // Actualizar timestamp de acceso
            storedLinks[shortId].accessed = Date.now();
            localStorage.setItem(SHORT_LINKS_STORAGE_KEY, JSON.stringify(storedLinks));
            
            return storedLinks[shortId].data;
        }
        
        return null;
    } catch (error) {
        console.error('Error retrieving short link data:', error);
        return null;
    }
}

// Función para crear hash corto del enlace
function createShortHash(data) {
    // Crear un hash simple del contenido
    let hash = 0;
    const str = JSON.stringify(data);
    
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convertir a 32-bit integer
    }
    
    // Convertir a base36 para hacer más corto y añadir timestamp
    const shortHash = Math.abs(hash).toString(36);
    const timeStamp = Date.now().toString(36).slice(-4); // Últimos 4 chars del timestamp
    
    return (shortHash + timeStamp).substring(0, 8); // Usar solo 8 caracteres totales
}

// Función para copiar el enlace al portapapeles
function copyShareableLink() {
    const linkInput = document.getElementById('shareableLink');
    
    try {
        // Usar la Clipboard API moderna si está disponible
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(linkInput.value).then(() => {
                showCopySuccess();
            }).catch(() => {
                fallbackCopy();
            });
        } else {
            fallbackCopy();
        }
    } catch (err) {
        console.error('❌ Error al copiar enlace:', err);
        alert('No se pudo copiar el enlace. Puedes seleccionarlo manualmente.');
    }
}

// Función de fallback para copiar
function fallbackCopy() {
    const linkInput = document.getElementById('shareableLink');
    linkInput.style.display = 'block';
    linkInput.select();
    linkInput.setSelectionRange(0, 99999);
    
    document.execCommand('copy');
    linkInput.style.display = 'none';
    showCopySuccess();
}

// Función para mostrar feedback de copia exitosa
function showCopySuccess() {
    // Cambiar temporalmente el texto del botón para dar feedback
    const copyBtn = document.querySelector('.btn-copy');
    const originalText = copyBtn.innerHTML;
    copyBtn.innerHTML = '<i class="fas fa-check"></i> ¡Copiado!';
    copyBtn.style.background = '#28a745';
    
    setTimeout(() => {
        copyBtn.innerHTML = originalText;
        copyBtn.style.background = '';
    }, 2000);
    
    console.log('✅ Enlace completo copiado al portapapeles');
}

// Función para cerrar el modal de compartir
function closeShareModal(event) {
    if (event && event.target !== event.currentTarget) return;
    document.getElementById('shareModal').style.display = 'none';
}

// Función para comprimir datos (usando base64 simple)
function compressData(data) {
    const jsonString = JSON.stringify(data);
    return btoa(encodeURIComponent(jsonString));
}

// Función para descomprimir datos
function decompressData(compressedData) {
    try {
        console.log('🔧 Iniciando descompresión de datos...');
        
        // Validar que el parámetro no esté vacío
        if (!compressedData || compressedData.trim() === '') {
            console.error('❌ Datos comprimidos vacíos');
            return null;
        }
        
        // Intentar decodificar base64
        let decodedData;
        try {
            decodedData = atob(compressedData);
        } catch (base64Error) {
            console.error('❌ Error en decodificación base64:', base64Error);
            return null;
        }
        
        // Intentar decodificar URI
        let jsonString;
        try {
            jsonString = decodeURIComponent(decodedData);
        } catch (uriError) {
            console.error('❌ Error en decodificación URI:', uriError);
            return null;
        }
        
        // Intentar parsear JSON
        let parsedData;
        try {
            parsedData = JSON.parse(jsonString);
        } catch (jsonError) {
            console.error('❌ Error en parseo JSON:', jsonError);
            return null;
        }
        
        console.log('✅ Datos descomprimidos exitosamente');
        return parsedData;
        
    } catch (error) {
        console.error('❌ Error general al descomprimir datos:', error);
        return null;
    }
}

// Función para cargar datos desde URL
function loadDataFromURL() {
    const urlParams = new URLSearchParams(window.location.search);
    const dataParam = urlParams.get('data');
    const idParam = urlParams.get('id');
    
    // Primero intentar cargar desde ID corto (nuevo sistema)
    if (idParam) {
        console.log('🔄 Cargando datos desde ID corto:', idParam);
        
        const data = getDataFromShortId(idParam);
        
        if (data) {
            try {
                // Validar que los datos tengan la estructura esperada
                if (data.projectInfo && data.summary) {
                    // Actualizar dashboard con los datos cargados
                    updateDashboardWithTransformedData(data);
                    
                    // Mostrar mensaje de información
                    showLoadedFromLinkMessage('ID corto');
                    
                    console.log('✅ Datos cargados desde ID corto');
                    console.log('📊 Proyecto:', data.projectInfo.name);
                } else {
                    console.error('❌ Los datos del ID corto no tienen la estructura esperada');
                    showErrorMessage('Los datos del enlace no son válidos o están corruptos.');
                }
            } catch (error) {
                console.error('❌ Error procesando datos del ID corto:', error);
                showErrorMessage('Error al procesar los datos del enlace compartido.');
            }
        } else {
            console.error('❌ No se encontraron datos para el ID corto:', idParam);
            showErrorMessage('El enlace ha expirado o no es válido. Los enlaces cortos tienen una duración limitada.');
        }
        return;
    }
    
    // Fallback al sistema antiguo (parámetro data)
    if (dataParam) {
        console.log('🔄 Cargando datos desde URL (sistema legacy)...');
        console.log('📏 Longitud del parámetro data:', dataParam.length);
        
        try {
            const decompressedData = decompressData(dataParam);
            
            if (decompressedData) {
                // Validar que los datos tengan la estructura esperada
                if (decompressedData.projectInfo && decompressedData.summary) {
                    // Actualizar dashboard con los datos cargados
                    updateDashboardWithTransformedData(decompressedData);
                    
                    // Mostrar mensaje de información
                    showLoadedFromLinkMessage('enlace comprimido');
                    
                    console.log('✅ Datos cargados desde enlace compartido (legacy)');
                    console.log('📊 Proyecto:', decompressedData.projectInfo.name);
                } else {
                    console.error('❌ Los datos no tienen la estructura esperada');
                    showErrorMessage('Los datos del enlace no son válidos o están corruptos.');
                }
            } else {
                console.error('❌ No se pudieron descomprimir los datos del enlace');
                showErrorMessage('No se pudo decodificar el enlace. El enlace puede estar incompleto o dañado.');
            }
        } catch (error) {
            console.error('❌ Error procesando datos del enlace:', error);
            showErrorMessage('Error al procesar los datos del enlace compartido.');
        }
    }
}

// Función para mostrar mensaje cuando se carga desde enlace
function showLoadedFromLinkMessage(linkType = 'enlace compartido') {
    const message = document.createElement('div');
    message.className = 'loaded-from-link-message';
    message.innerHTML = `
        <i class="fas fa-info-circle"></i>
        Datos cargados desde ${linkType}
        <button onclick="this.parentElement.remove()" style="background: none; border: none; color: white; margin-left: 10px; cursor: pointer;">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    message.style.cssText = `
        position: fixed;
        top: 80px;
        right: 20px;
        background: var(--primary-color);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1000;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        font-weight: 500;
        animation: slideInRight 0.3s ease-out;
    `;
    
    document.body.appendChild(message);
    
    // Auto remover después de 5 segundos
    setTimeout(() => {
        if (message.parentElement) {
            message.remove();
        }
    }, 5000);
}

// Función para mostrar mensajes de error
function showErrorMessage(errorText) {
    const message = document.createElement('div');
    message.className = 'error-message';
    message.innerHTML = `
        <i class="fas fa-exclamation-triangle"></i>
        ${errorText}
        <button onclick="this.parentElement.remove()" style="background: none; border: none; color: white; margin-left: 10px; cursor: pointer;">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    message.style.cssText = `
        position: fixed;
        top: 80px;
        right: 20px;
        background: #dc2626;
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1000;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        font-weight: 500;
        animation: slideInRight 0.3s ease-out;
        max-width: 400px;
    `;
    
    document.body.appendChild(message);
    
    // Auto remover después de 8 segundos para errores
    setTimeout(() => {
        if (message.parentElement) {
            message.remove();
        }
    }, 8000);
}

// Agregar animación CSS para el mensaje
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    .loaded-from-link-message {
        animation: slideInRight 0.3s ease-out;
    }
`;
document.head.appendChild(style);

// Función para abrir el selector de proyecto con datos precargados
function openProjectSelector() {
    if (loadedExcelData && loadedExcelData.proyectos && loadedExcelData.proyectos.length > 0) {
        // Si hay datos cargados, mostrar el modal de selección
        showProjectSelector(loadedExcelData.proyectos, loadedExcelData);
    } else {
        // Si no hay datos cargados, abrir selector de archivos
        const fileInput = document.getElementById('excelFile');
        if (fileInput) {
            fileInput.click();
        } else {
            showNotification('No hay datos cargados. Por favor, cargue un archivo Excel primero.', 'warning');
        }
    }
}

// Función para renderizar el gráfico de estado de entregables
function updateDeliverablesChart() {
    const chartElement = document.getElementById('deliverablesChart');
    if (!chartElement) {
        console.warn('Elemento deliverablesChart no encontrado en el DOM');
        return;
    }

    // Verificar si hay datos del proyecto cargados
    if (!testData || !testData.projectInfo) {
        console.log('No hay datos de proyecto cargados para entregables');
        
        const layout = {
            margin: { t: 20, b: 20, l: 20, r: 20 },
            showlegend: false,
            paper_bgcolor: 'transparent',
            plot_bgcolor: 'transparent',
            annotations: [{
                text: 'No hay datos de entregables disponibles',
                x: 0.5,
                y: 0.5,
                xref: 'paper',
                yref: 'paper',
                showarrow: false,
                font: { size: 14, color: '#64748b' }
            }]
        };

        Plotly.newPlot('deliverablesChart', [], layout, { responsive: true, displayModeBar: false });
        return;
    }

    // Definición de los entregables con sus campos del Excel
    const deliverables = [
        { name: 'Plan de Pruebas', code: 'PP', field: 'PP' },
        { name: 'Matriz de Trazabilidad', code: 'MT', field: 'MT' },
        { name: 'Matriz de Casos de Prueba', code: 'CP', field: 'CP' },
        { name: 'Evidencias de Prueba', code: 'EE', field: 'EE' },
        { name: 'Matriz de Defectos', code: 'DF', field: 'DF' },
        { name: 'Informe Final', code: 'AC', field: 'AC' }
    ];

    // DEBUG: Mostrar datos de entregables
    console.log('🔍 DEBUG ENTREGABLES - testData.projectInfo:', testData.projectInfo);
    console.log('📋 Valores de entregables en projectInfo:');
    deliverables.forEach(d => {
        console.log(`   ${d.code} (${d.name}):`, testData.projectInfo[d.field]);
    });

    // Procesar los datos de entregables
    const labels = [];
    const colors = [];
    const xValues = [];
    const statusTexts = [];

    deliverables.forEach((deliverable, index) => {
        labels.push(deliverable.name);
        xValues.push(0.5); // Centrar los círculos
        
        let color = '#9ca3af'; // Gris por defecto (sin información)
        let statusText = 'Sin información';
        
        if (testData.projectInfo[deliverable.field]) {
            const value = testData.projectInfo[deliverable.field].toString().toUpperCase().trim();
            
            console.log(`🎨 Procesando ${deliverable.code}: valor original="${testData.projectInfo[deliverable.field]}", procesado="${value}"`);
            
            // Detectar color directamente del Excel
            if (value === 'V' || value === 'VERDE' || value === 'GREEN' || value === 'COMPLETADO') {
                color = '#10b981'; // Verde
                statusText = 'Completado';
                console.log(`   ✅ Color asignado: VERDE`);
            } else if (value === 'A' || value === 'AMARILLO' || value === 'YELLOW' || value === 'EN PROGRESO') {
                color = '#f59e0b'; // Amarillo/Naranja
                statusText = 'En progreso';
                console.log(`   ⚠️ Color asignado: AMARILLO`);
            } else if (value === 'R' || value === 'ROJO' || value === 'RED' || value === 'PENDIENTE') {
                color = '#ef4444'; // Rojo
                statusText = 'Pendiente';
                console.log(`   ❌ Color asignado: ROJO`);
            } else {
                // Si no coincide con ningún valor conocido, mantener gris
                statusText = 'Sin información';
                console.log(`   ⚪ Color asignado: GRIS (valor no reconocido: "${value}")`);
            }
        } else {
            console.log(`🎨 ${deliverable.code}: Campo vacío o undefined`);
        }
        
        colors.push(color);
        statusTexts.push(statusText);
    });

    // Invertir arrays para que se muestren en orden correcto (Plotly invierte el eje Y)
    labels.reverse();
    colors.reverse();
    xValues.reverse();
    statusTexts.reverse();

    // Crear el gráfico tipo scatter con círculos grandes
    const trace = {
        type: 'scatter',
        mode: 'markers+text',
        x: xValues,
        y: labels,
        marker: {
            size: 30,
            color: colors,
            line: {
                color: 'rgba(0,0,0,0.2)',
                width: 2
            },
            symbol: 'circle'
        },
        text: statusTexts,
        textposition: 'middle right',
        textfont: {
            size: 11,
            color: '#1e293b',
            family: 'Inter, sans-serif',
            weight: 500
        },
        hovertemplate: '<b>%{y}</b><br>Estado: %{text}<br><extra></extra>'
    };

    const layout = {
        margin: { t: 30, b: 40, l: 200, r: 150 },
        showlegend: false,
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        xaxis: {
            showticklabels: false,
            showgrid: false,
            zeroline: false,
            range: [0, 1],
            fixedrange: true
        },
        yaxis: {
            tickfont: {
                size: 11,
                color: '#1e293b',
                family: 'Inter, sans-serif',
                weight: 500
            },
            automargin: true,
            fixedrange: true
        },
        font: {
            family: 'Inter, sans-serif',
            size: 11,
            color: '#1e293b'
        },
        height: 320,
        hovermode: 'closest'
    };

    const config = {
        responsive: true,
        displayModeBar: false,
        staticPlot: false
    };

    Plotly.newPlot('deliverablesChart', [trace], layout, config);
}

// Sobrescribir la funci�n handleFileUpload para usar el nuevo selector de hojas
document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('excelFile');
    if (fileInput) {
        fileInput.removeEventListener('change', handleFileUpload);
        fileInput.addEventListener('change', function(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            const fileName = file.name.toLowerCase();
            
            if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                // Usar el nuevo manejador con selector de hojas
                handleExcelFileWithSheets(file);
            } else if (fileName.endsWith('.csv')) {
                handleCSVFile(file);
            } else {
                showNotification('Formato de archivo no soportado. Use .xlsx, .xls o .csv', 'error');
                event.target.value = '';
            }
        });
    }
});
