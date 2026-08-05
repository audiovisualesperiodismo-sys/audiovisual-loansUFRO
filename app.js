// ==========================================
// CONFIGURACIÓN Y ESTADO DE LA APLICACIÓN
// ==========================================

const CONFIG = {
    demoMode: false,
    scriptUrl: localStorage.getItem('audiolend_script_url') || 'https://script.google.com/macros/s/AKfycbywwlkVaHLmOA3c67V3PhPjcIK96a2IYfkFZ1Mgw9HdGRYp3Kj4twWJvoZVBDK7_Fu8/exec',
    sheetUrl: localStorage.getItem('audiolend_sheet_url') || 'https://docs.google.com/spreadsheets/d/1HTRMz15UmVqTHgGkU95OjK7wihZJgEm05G7QfgNqNeg/edit?pli=1&gid=1451542023#gid=1451542023',
    adminPassword: 'adminfelipe',
};

// Base de Datos en Memoria para el MODO DEMO (Columnas reales de la UFRO)
const DEMO_DATABASE = {
    inventory: [
        { id: 1, category: "Cámaras", name: "Sony Alpha 7 III", total: 5, available: 5, codes: ["SONY-01", "SONY-02", "SONY-03", "SONY-04", "SONY-05"], image: "https://images.unsplash.com/photo-1516035069371-29a1b244cc32?w=500&auto=format&fit=crop&q=60", description: "Cámara mirrorless Full Frame ideal para video y fotografía profesional." },
        { id: 2, category: "Cámaras", name: "Canon EOS R6", total: 3, available: 2, codes: ["CAN-01", "CAN-02", "CAN-03"], image: "https://images.unsplash.com/photo-1502982720700-bfff97f2ecac?w=500&auto=format&fit=crop&q=60", description: "Cámara mirrorless Full Frame con excelente estabilización e ideal para tomas de acción." },
        { id: 3, category: "Trípodes", name: "Manfrotto Pro", total: 6, available: 5, codes: ["MAN-01", "MAN-02", "MAN-03", "MAN-04", "MAN-05", "MAN-06"], image: "https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?w=500&auto=format&fit=crop&q=60", description: "Trípode de video profesional con cabezal fluido y gran capacidad de carga." },
        { id: 4, category: "Trípodes", name: "Benro Aero Video", total: 4, available: 4, codes: ["BEN-01", "BEN-02", "BEN-03", "BEN-04"], image: "https://images.unsplash.com/photo-1590233649088-e81e12a897b1?w=500&auto=format&fit=crop&q=60", description: "Trípode compacto para viajes y grabaciones rápidas con paneo suave." },
        { id: 5, category: "Audio", name: "Micrófono Røde Wireless Go", total: 8, available: 8, codes: ["RODE-01", "RODE-02", "RODE-03", "RODE-04", "RODE-05", "RODE-06", "RODE-07", "RODE-08"], image: "https://images.unsplash.com/photo-1590602847861-f357a9332bbc?w=500&auto=format&fit=crop&q=60", description: "Sistema de micrófono inalámbrico ultracompacto ideal para entrevistas rápidas." },
        { id: 6, category: "Audio", name: "Boom Sennheiser ME66", total: 3, available: 1, codes: ["SENN-01", "SENN-02", "SENN-03"], image: "https://images.unsplash.com/photo-1590602846989-e2458b330d3b?w=500&auto=format&fit=crop&q=60", description: "Micrófono de cañón profesional con excelente direccionalidad y rechazo de ruido ambiente." },
        { id: 7, category: "Luces", name: "Foco Led Aputure 300d II", total: 4, available: 4, codes: ["AP-01", "AP-02", "AP-03", "AP-04"], image: "https://images.unsplash.com/photo-1507679799987-c73779587ccf?w=500&auto=format&fit=crop&q=60", description: "Luz LED de alta potencia para iluminación de sets de grabación profesional." },
        { id: 8, category: "Luces", name: "Kit Led Neewer 480", total: 5, available: 5, codes: ["NW-01", "NW-02", "NW-03", "NW-04", "NW-05"], image: "https://images.unsplash.com/photo-1616440347437-b1c73416efc2?w=500&auto=format&fit=crop&q=60", description: "Kit de focos LED portátiles con temperatura de color regulable para estudio de retrato y entrevistas." }
    ],
    students: [],
    loans: []
};

// Estado Local
let appState = {
    inventory: [],
    students: [],
    loans: [],
    cart: [],
    subjects: [],
    selectedCategory: "Cámaras",
    validatedStudent: null,
    isAdminLoggedIn: false,
    activeAdminSubtab: "admin-tab-loans",
    activeAdminLoanFilter: "all",
    currentDeliveryLoanId: null,
    currentReturnLoanId: null,
    tempDeliveryItems: []
};

// Referencias DOM
const dom = {
    navButtons: document.querySelectorAll('.nav-btn'),
    tabContents: document.querySelectorAll('.tab-content'),
    currentSectionTitle: document.getElementById('current-section-title'),
    currentSectionSubtitle: document.getElementById('current-section-subtitle'),
    connectionStatus: document.getElementById('connection-status'),
    modeToggle: document.getElementById('mode-toggle'),
    modeSwitchContainer: document.getElementById('mode-switch-container'),
    btnRefresh: document.getElementById('btn-refresh'),
    categoryTabs: document.getElementById('category-tabs-list'),
    equipmentGrid: document.getElementById('equipment-grid-container'),
    selectedItems: document.getElementById('selected-items-container'),
    studentRut: document.getElementById('student-rut'),
    btnVerifyStudent: document.getElementById('btn-verify-student'),
    studentInfoBox: document.getElementById('student-info-box'),
    studentStatusBadge: document.getElementById('student-status-badge'),
    studentStatusText: document.getElementById('student-status-text'),
    displayStudentName: document.getElementById('display-student-name'),
    displayStudentEmail: document.getElementById('display-student-email'),
    displayStudentCareer: document.getElementById('display-student-career'), // Usado para mostrar Fono
    studentDebtWarning: document.getElementById('student-debt-warning'),
    displayStudentDebt: document.getElementById('display-student-debt'),
    inventoryCodesGroup: document.getElementById('inventory-codes-group'),
    inventoryInputsList: document.getElementById('inventory-inputs-list'),
    btnConfirmLoan: document.getElementById('btn-confirm-loan'),
    loanDetailsGroup: document.getElementById('loan-details-group'),
    loanDateOut: document.getElementById('loan-date-out'),
    loanDateIn: document.getElementById('loan-date-in'),
    loanSubject: document.getElementById('loan-subject'),
    adminLoginForm: document.getElementById('admin-login-form'),
    adminPassword: document.getElementById('admin-password'),
    adminLoginContainer: document.getElementById('admin-login-container'),
    adminDashboardContainer: document.getElementById('admin-dashboard-container'),
    adminSubnavButtons: document.querySelectorAll('.admin-subnav-btn'),
    adminTabContents: document.querySelectorAll('.admin-tab-content'),
    adminLoansTableBody: document.getElementById('admin-loans-table-body'),
    adminLoanFilterBtns: document.querySelectorAll('.btn-filter'),
    formAddEquipment: document.getElementById('form-add-equipment'),
    formAddStudent: document.getElementById('form-add-student'),
    stStatusSelect: document.getElementById('st-status'),
    stDebtGroup: document.getElementById('st-debt-group'),
    adminEquipmentList: document.getElementById('admin-equipment-list'),
    adminStudentsList: document.getElementById('admin-students-list'),
    btnSaveSettings: document.getElementById('btn-save-settings'),
    scriptUrlInput: document.getElementById('script-url-input'),
    sheetUrlInput: document.getElementById('sheet-url-input'),
    linkGoogleSheet: document.getElementById('link-google-sheet'),
    btnLogout: document.getElementById('btn-logout'),
    btnLogoutSubnav: document.getElementById('btn-logout-subnav'),
    btnExportExcel: document.getElementById('btn-export-excel'),
    toastContainer: document.getElementById('toast-container'),
    chartPopularCategory: document.getElementById('chart-popular-category'),
    
    // Modales
    scannerModal: document.getElementById('scanner-modal'),
    btnCloseScanner: document.getElementById('btn-close-scanner'),
    deliveryModal: document.getElementById('delivery-modal'),
    btnCloseDelivery: document.getElementById('btn-close-delivery'),
    deliveryInputsContainer: document.getElementById('delivery-inputs-container'),
    deliveryGlobalObs: document.getElementById('delivery-global-obs'),
    btnConfirmDelivery: document.getElementById('btn-confirm-delivery'),
    deliveryAddEquipmentSelect: document.getElementById('delivery-add-equipment-select'),
    btnDeliveryAddEquipment: document.getElementById('btn-delivery-add-equipment'),
    successModal: document.getElementById('success-modal'),
    btnCloseSuccess: document.getElementById('btn-close-success'),
    successLoanSummary: document.getElementById('success-loan-summary'),
    
    // Devolución Modal
    returnModal: document.getElementById('return-modal'),
    btnCloseReturn: document.getElementById('btn-close-return'),
    btnConfirmReturn: document.getElementById('btn-confirm-return'),
    returnItemsContainer: document.getElementById('return-items-container'),
    returnGlobalObs: document.getElementById('return-global-obs'),
    
    // Métricas
    metricTotalEquipos: document.getElementById('metric-total-equipos'),
    metricPrestamosActivos: document.getElementById('metric-prestamos-activos'),
    metricPrestamosAtrasados: document.getElementById('metric-prestamos-atrasados'),
    metricAlumnosBloqueados: document.getElementById('metric-alumnos-bloqueados'),
    metricTotalHistorico: document.getElementById('metric-total-historico'),
};

let charts = {
    popularItems: null,
    flowDays: null,
    topStudents: null,
    subjectDistribution: null,
    overdueStatus: null
};

let html5QrReader = null;

// ==========================================
// INICIALIZACIÓN
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    CONFIG.demoMode = dom.modeToggle.checked;
    dom.scriptUrlInput.value = CONFIG.scriptUrl;
    if (dom.sheetUrlInput) {
        dom.sheetUrlInput.value = CONFIG.sheetUrl.includes('1_YOUR_SHEET_ID') ? '' : CONFIG.sheetUrl;
    }
    
    // Forzar limpieza de base demo antigua para empezar de cero
    if (!localStorage.getItem('audiolend_clean_slate_v1')) {
        localStorage.removeItem('audiolend_demo_inventory');
        localStorage.removeItem('audiolend_demo_students');
        localStorage.removeItem('audiolend_demo_loans');
        localStorage.setItem('audiolend_clean_slate_v1', 'true');
    }
    
    // Configurar fecha mínima de retiro para hoy
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    if (dom.loanDateOut) {
        dom.loanDateOut.min = `${yyyy}-${mm}-${dd}`;
    }
    
    loadData();
    initEventListeners();
    lucide.createIcons();
    
    // Registro de Service Worker para PWA (Fase 23)
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('./sw.js')
                .then(reg => console.log('Service Worker registrado con éxito', reg))
                .catch(err => console.warn('Error al registrar Service Worker', err));
        });
    }
});

function getDirectImageUrl(url) {
    if (!url) return '';
    url = url.trim();
    if (url.includes('drive.google.com') || url.includes('docs.google.com')) {
        let fileId = '';
        const dMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (dMatch && dMatch[1]) {
            fileId = dMatch[1];
        } else {
            const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
            if (idMatch && idMatch[1]) {
                fileId = idMatch[1];
            }
        }
        if (fileId) {
            return `https://lh3.googleusercontent.com/d/${fileId}`;
        }
    }
    return url;
}

function groupInventoryItems(rawInventory) {
    const grouped = [];
    if (!rawInventory || !Array.isArray(rawInventory)) return [];
    
    rawInventory.forEach(item => {
        if (!item || !item.name) return;
        const normalizedName = item.name.trim();
        const existing = grouped.find(i => i.name.toLowerCase() === normalizedName.toLowerCase());
        
        if (existing) {
            existing.total += parseInt(item.total) || 0;
            existing.available += parseInt(item.available) || 0;
            if (item.codes && Array.isArray(item.codes)) {
                item.codes.forEach(c => {
                    const cleanCode = c.trim();
                    if (cleanCode && !existing.codes.includes(cleanCode)) {
                        existing.codes.push(cleanCode);
                    }
                });
            }
            if (!existing.image && item.image) {
                existing.image = getDirectImageUrl(item.image);
            }
            if (!existing.description && item.description) {
                existing.description = item.description;
            }
        } else {
            grouped.push({
                id: item.id,
                category: item.category,
                name: normalizedName,
                total: parseInt(item.total) || 0,
                available: parseInt(item.available) || 0,
                codes: item.codes ? [...item.codes].map(c => c.trim()) : [],
                image: item.image ? getDirectImageUrl(item.image) : "",
                description: item.description || ""
            });
        }
    });
    return grouped;
}

async function loadData() {
    if (CONFIG.demoMode) {
        if (appState.inventory.length === 0) {
            let rawInv = JSON.parse(localStorage.getItem('audiolend_demo_inventory'));
            if (rawInv && Array.isArray(rawInv) && rawInv.length > 0 && !rawInv.some(i => i.description || i.image)) {
                localStorage.removeItem('audiolend_demo_inventory');
                rawInv = null;
            }
            if (!rawInv) {
                rawInv = JSON.parse(JSON.stringify(DEMO_DATABASE.inventory));
            }
            appState.inventory = groupInventoryItems(rawInv);
            appState.students = JSON.parse(localStorage.getItem('audiolend_demo_students')) || JSON.parse(JSON.stringify(DEMO_DATABASE.students));
            appState.loans = JSON.parse(localStorage.getItem('audiolend_demo_loans')) || JSON.parse(JSON.stringify(DEMO_DATABASE.loans));
        }
        appState.subjects = ["Periodismo Escrito", "Periodismo Radial", "Periodismo Televisivo", "Fotoperiodismo", "Comunicación Digital", "Cine y Documental", "Proyecto de Título", "Ninguna (Proyecto Personal)"];
        
        recalculateDemoStock();
        updateConnectionStatus(true);
        renderLoansModule();
        renderSubjectsDropdown();
        
        if (appState.isAdminLoggedIn) {
            updateAdminDashboard();
            renderAdminLoans(appState.activeAdminLoanFilter);
        }
        showToast("Datos demo activos (AVP UFRO)", "info");
    } else {
        if (!CONFIG.scriptUrl) {
            showToast("Especifica la URL del Apps Script en el Panel Admin.", "warning");
            setDemoMode(true);
            return;
        }
        
        showToast("Conectando con base de datos UFRO...", "info");
        try {
            updateConnectionStatus(null);
            const response = await fetch(`${CONFIG.scriptUrl}?action=getInitData`);
            if (!response.ok) throw new Error("Error en red");
            
            const data = await response.json();
            if (data.status === "error") throw new Error(data.message);
            
            if (data.debugInfo) {
                console.log("[DEBUG AVP] Encabezados del Sheet e índices detectados:", data.debugInfo);
            }
            
            appState.inventory = groupInventoryItems(data.inventory);
            appState.students = data.students;
            appState.loans = data.loans;
            appState.subjects = data.subjects || [];
            
            if (data.sheetUrl) {
                CONFIG.sheetUrl = data.sheetUrl;
                localStorage.setItem('audiolend_sheet_url', data.sheetUrl);
                if (dom.sheetUrlInput) {
                    dom.sheetUrlInput.value = data.sheetUrl;
                }
            }
            
            updateConnectionStatus(false);
            renderLoansModule();
            renderSubjectsDropdown();
            
            if (appState.isAdminLoggedIn) {
                updateAdminDashboard();
                renderAdminLoans(appState.activeAdminLoanFilter);
            }
            showToast("Datos UFRO sincronizados correctamente", "success");
        } catch (error) {
            console.error(error);
            showToast(`Error de conexión: ${error.message || error}. Volviendo a Modo Demo.`, "danger");
            setDemoMode(true);
        }
    }
}

function saveDemoState() {
    if (CONFIG.demoMode) {
        localStorage.setItem('audiolend_demo_inventory', JSON.stringify(appState.inventory));
        localStorage.setItem('audiolend_demo_students', JSON.stringify(appState.students));
        localStorage.setItem('audiolend_demo_loans', JSON.stringify(appState.loans));
    }
}

function recalculateDemoStock() {
    appState.inventory.forEach(item => {
        item.available = item.total;
    });
    appState.loans.forEach(loan => {
        if (loan.status === "Solicitado" || loan.status === "Retirado") {
            const item = appState.inventory.find(i => i.name === loan.item);
            if (item && item.available > 0) {
                item.available--;
            }
        }
    });
}

function setDemoMode(active) {
    CONFIG.demoMode = active;
    dom.modeToggle.checked = active;
    updateConnectionStatus(active);
    loadData();
}

function updateConnectionStatus(isDemo) {
    if (isDemo === true) {
        dom.connectionStatus.className = "status-indicator demo-active";
        dom.connectionStatus.querySelector('.status-label').textContent = "Modo Demo Activo";
    } else if (isDemo === false) {
        dom.connectionStatus.className = "status-indicator live-active";
        dom.connectionStatus.querySelector('.status-label').textContent = "Conectado a Google Sheets";
    } else {
        dom.connectionStatus.className = "status-indicator";
        dom.connectionStatus.querySelector('.status-label').textContent = "Sincronizando...";
    }
}

// ==========================================
// REGISTRO DE EVENTOS
// ==========================================

function initEventListeners() {
    dom.modeToggle.addEventListener('change', (e) => {
        setDemoMode(e.target.checked);
    });
    
    dom.btnRefresh.addEventListener('click', () => {
        loadData();
    });
    
    dom.navButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const targetButton = e.currentTarget;
            const targetSection = targetButton.getAttribute('data-target');
            
            showSection(targetSection);
            dom.navButtons.forEach(b => b.classList.remove('active'));
            targetButton.classList.add('active');
            
            const titleMap = {
                'section-loans': { title: 'Formulario de Solicitud', subtitle: 'Selecciona los equipos y valida tu RUT para enviar tu solicitud de préstamo.' },
                'section-protocols': { title: 'Protocolos de Uso', subtitle: 'Revisa las condiciones y normativas de préstamo de equipos AVP UFRO.' },
                'section-admin': { title: 'Panel de Administración AVP', subtitle: 'Gestiona solicitudes, registra retiros físicos, administra catálogo y estadísticas.' }
            };
            
            dom.currentSectionTitle.textContent = titleMap[targetSection].title;
            dom.currentSectionSubtitle.textContent = titleMap[targetSection].subtitle;
        });
    });
    
    dom.adminSubnavButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const target = e.currentTarget.getAttribute('data-admin-target');
            appState.activeAdminSubtab = target;
            
            dom.adminSubnavButtons.forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            
            dom.adminTabContents.forEach(tab => {
                if (tab.id === target) tab.classList.add('active');
                else tab.classList.remove('active');
            });
            
            if (target === 'admin-tab-stats') {
                setTimeout(() => updateCharts(), 100);
            } else if (target === 'admin-tab-loans') {
                renderAdminLoans(appState.activeAdminLoanFilter);
            } else if (target === 'admin-tab-config') {
                renderAdminConfigLists();
            }
        });
    });
    
    dom.adminLoanFilterBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            dom.adminLoanFilterBtns.forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            const filterValue = e.currentTarget.getAttribute('data-filter');
            appState.activeAdminLoanFilter = filterValue;
            renderAdminLoans(filterValue);
        });
    });
    
    dom.formAddEquipment.addEventListener('submit', (e) => {
        e.preventDefault();
        saveEquipmentConfig();
    });
    
    dom.formAddStudent.addEventListener('submit', (e) => {
        e.preventDefault();
        saveStudentConfig();
    });
    
    dom.stStatusSelect.addEventListener('change', (e) => {
        if (e.target.value === 'Bloqueado') {
            dom.stDebtGroup.classList.remove('hidden');
            document.getElementById('st-debt').required = true;
        } else {
            dom.stDebtGroup.classList.add('hidden');
            document.getElementById('st-debt').required = false;
        }
    });
    
    dom.btnVerifyStudent.addEventListener('click', () => {
        verifyStudent();
    });
    
    dom.studentRut.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') verifyStudent();
    });
    
    dom.btnConfirmLoan.addEventListener('click', () => {
        processLoanCheckout();
    });

    dom.loanDateOut.addEventListener('change', () => {
        const dateOut = dom.loanDateOut.value;
        if (dateOut) {
            // Poner la fecha mínima de entrega igual a la de retiro
            dom.loanDateIn.min = dateOut;
            
            // Calcular la fecha máxima (retiro + 4 días)
            const parts = dateOut.split('-');
            const y = parseInt(parts[0], 10);
            const m = parseInt(parts[1], 10) - 1;
            const d = parseInt(parts[2], 10);
            
            const maxDate = new Date(y, m, d + 4);
            const maxYear = maxDate.getFullYear();
            const maxMonth = String(maxDate.getMonth() + 1).padStart(2, '0');
            const maxDay = String(maxDate.getDate()).padStart(2, '0');
            const maxDateStr = `${maxYear}-${maxMonth}-${maxDay}`;
            
            dom.loanDateIn.max = maxDateStr;
            
            // Si la fecha actual de entrega ya no calza con el rango, restablecerla
            const dateIn = dom.loanDateIn.value;
            if (dateIn) {
                if (dateIn < dateOut || dateIn > maxDateStr) {
                    dom.loanDateIn.value = "";
                    showToast("La fecha de entrega se ha restablecido por estar fuera del límite de 4 días.", "warning");
                }
            }
        } else {
            dom.loanDateIn.min = "";
            dom.loanDateIn.max = "";
        }
    });

    dom.loanDateIn.addEventListener('change', () => {
        const dateOut = dom.loanDateOut.value;
        const dateIn = dom.loanDateIn.value;
        if (!dateOut && dateIn) {
            showToast("Por favor, selecciona primero la fecha de retiro.", "warning");
            dom.loanDateIn.value = "";
            return;
        }
        if (dateOut && dateIn) {
            const partsOut = dateOut.split('-').map(Number);
            const partsIn = dateIn.split('-').map(Number);
            const dateOutObj = new Date(partsOut[0], partsOut[1] - 1, partsOut[2]);
            const dateInObj = new Date(partsIn[0], partsIn[1] - 1, partsIn[2]);
            
            if (dateOutObj >= dateInObj) {
                showToast("La fecha de retiro debe ser anterior a la de entrega.", "warning");
                dom.loanDateIn.value = "";
            } else {
                const diffTime = dateInObj - dateOutObj;
                const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
                if (diffDays > 4) {
                    showToast("El período de préstamo no puede ser superior a 4 días.", "warning");
                    dom.loanDateIn.value = "";
                }
            }
        }
    });
    
    dom.btnConfirmDelivery.addEventListener('click', () => {
        confirmPhysicalDelivery();
    });
    
    dom.btnCloseDelivery.addEventListener('click', () => {
        dom.deliveryModal.classList.add('hidden');
    });
    
    if (dom.btnDeliveryAddEquipment) {
        dom.btnDeliveryAddEquipment.addEventListener('click', (e) => {
            e.preventDefault();
            const loanId = appState.currentDeliveryLoanId;
            if (!loanId) return;
            const selectedName = dom.deliveryAddEquipmentSelect.value;
            if (!selectedName) {
                showToast("Por favor selecciona un equipo.", "warning");
                return;
            }
            
            const firstItem = appState.loans.find(l => l.id === loanId);
            if (!firstItem) return;
            
            const extraItem = {
                id: loanId,
                rut: firstItem.rut,
                name: firstItem.name,
                email: firstItem.email,
                item: selectedName,
                code: "",
                status: "Solicitado",
                dateOut: firstItem.dateOut,
                dateDeliver: "",
                dateIn: "",
                progRetiro: firstItem.progRetiro,
                progDevolucion: firstItem.progDevolucion,
                subject: firstItem.subject,
                obs: "",
                obsReturn: "",
                daysOverdue: 0,
                isNew: true
            };
            
            appState.tempDeliveryItems.push(extraItem);
            renderDeliveryModalInputs();
        });
    }
    
    dom.btnConfirmReturn.addEventListener('click', () => {
        confirmReturnCheckout();
    });
    
    dom.btnCloseReturn.addEventListener('click', () => {
        dom.returnModal.classList.add('hidden');
    });
    
    dom.adminLoginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        if (dom.adminPassword.value === CONFIG.adminPassword) {
            appState.isAdminLoggedIn = true;
            dom.adminLoginContainer.classList.add('hidden');
            dom.adminDashboardContainer.classList.remove('hidden');
            dom.adminPassword.value = '';
            
            if (dom.modeSwitchContainer) {
                dom.modeSwitchContainer.classList.remove('hidden');
            }
            if (dom.btnRefresh) {
                dom.btnRefresh.classList.remove('hidden');
            }
            
            document.querySelector(`.admin-subnav-btn[data-admin-target="${appState.activeAdminSubtab}"]`).click();
            updateAdminDashboard();
            showToast("Acceso Administrador AVP verificado", "success");
        } else {
            showToast("Contraseña incorrecta", "danger");
        }
    });
    
    [dom.btnLogout, dom.btnLogoutSubnav].forEach(btn => {
        if (btn) {
            btn.addEventListener('click', () => {
                appState.isAdminLoggedIn = false;
                dom.adminLoginContainer.classList.remove('hidden');
                dom.adminDashboardContainer.classList.add('hidden');
                
                if (dom.modeSwitchContainer) {
                    dom.modeSwitchContainer.classList.add('hidden');
                }
                if (dom.btnRefresh) {
                    dom.btnRefresh.classList.add('hidden');
                }
                
                dom.navButtons[0].click();
                showToast("Sesión Admin cerrada", "info");
            });
        }
    });
    
    if (dom.btnExportExcel) {
        dom.btnExportExcel.addEventListener('click', () => {
            exportLoansToExcel();
        });
    }
    
    if (dom.chartPopularCategory) {
        dom.chartPopularCategory.addEventListener('change', () => {
            updateCharts();
        });
    }
    
    dom.btnSaveSettings.addEventListener('click', () => {
        const url = dom.scriptUrlInput.value.trim();
        if (url) {
            if (url.includes('/edit')) {
                showToast("Error: Has copiado la URL de edición del script. Debes usar la URL de implementación que termina en '/exec'.", "danger");
                return;
            }
            if (!url.includes('/macros/s/') || (!url.includes('/exec') && !url.includes('/dev'))) {
                showToast("Advertencia: Asegúrate de que la URL corresponda a la Web App y termine en '/exec'.", "warning");
            }
        }
        CONFIG.scriptUrl = url;
        localStorage.setItem('audiolend_script_url', url);
        
        const sheetUrl = dom.sheetUrlInput ? dom.sheetUrlInput.value.trim() : '';
        if (sheetUrl) {
            CONFIG.sheetUrl = sheetUrl;
            localStorage.setItem('audiolend_sheet_url', sheetUrl);
        } else {
            localStorage.removeItem('audiolend_sheet_url');
            CONFIG.sheetUrl = 'https://docs.google.com/spreadsheets/d/1_YOUR_SHEET_ID/edit';
        }
        
        showToast("Configuración guardada", "success");
        if (url) setDemoMode(false);
    });
    
    const btnRunDiagnostics = document.getElementById('btn-run-diagnostics');
    const diagnosticsResult = document.getElementById('diagnostics-result');
    if (btnRunDiagnostics && diagnosticsResult) {
        btnRunDiagnostics.addEventListener('click', async () => {
            diagnosticsResult.style.display = "block";
            diagnosticsResult.textContent = "Ejecutando diagnóstico... Por favor espera.";
            
            if (CONFIG.demoMode) {
                diagnosticsResult.textContent = "Modo Demo Activo.\n\nPara ejecutar el diagnóstico, ingresa tu URL de Google Apps Script y desactiva el switch de Modo Demo.";
                return;
            }
            
            try {
                const response = await fetch(`${CONFIG.scriptUrl}?action=getInitData`);
                if (!response.ok) throw new Error("Error en la petición de red.");
                const data = await response.json();
                
                if (data.status === "error") {
                    diagnosticsResult.textContent = `Error del Servidor:\n${data.message}`;
                    return;
                }
                
                if (data.debugInfo) {
                    const info = data.debugInfo;
                    let text = `Columnas detectadas en 'Inventario':\n${JSON.stringify(info.detectedHeaders, null, 2)}\n\n`;
                    text += `Índices detectados:\n`;
                    text += `- Categoría: ${info.indices.categoryIdx !== -1 ? 'OK (col ' + (info.indices.categoryIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Nombre (Equipo): ${info.indices.nameIdx !== -1 ? 'OK (col ' + (info.indices.nameIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Cantidad Total: ${info.indices.totalIdx !== -1 ? 'OK (col ' + (info.indices.totalIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Disponible: ${info.indices.availableIdx !== -1 ? 'OK (col ' + (info.indices.availableIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Códigos: ${info.indices.codesIdx !== -1 ? 'OK (col ' + (info.indices.codesIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Imagen: ${info.indices.imageIdx !== -1 ? 'OK (col ' + (info.indices.imageIdx + 1) + ')' : 'No encontrado (-1)'}\n`;
                    text += `- Descripción: ${info.indices.descriptionIdx !== -1 ? 'OK (col ' + (info.indices.descriptionIdx + 1) + ')' : 'No encontrado (-1)'}\n\n`;
                    
                    if (info.indices.imageIdx === -1 || info.indices.descriptionIdx === -1) {
                        text += `⚠️ ATENCIÓN: El script no está encontrando las columnas de Imagen o Descripción. Asegúrate de agregarlas en la fila 1 de la pestaña 'Inventario' de tu Google Sheets con los nombres recomendados (Imagen, Descripción) y publicar una 'Nueva versión' del código en Apps Script.`;
                    } else {
                        text += `✅ TODO CORRECTO: Las columnas e índices han sido asignados correctamente. Si las imágenes siguen sin verse, comprueba que los enlaces pegados en las celdas sean válidos y públicos.`;
                    }
                    diagnosticsResult.textContent = text;
                } else {
                    diagnosticsResult.textContent = `Conectado, pero el script no retornó la información de diagnóstico.\nEsto confirma que el script web en Google Sheets no tiene la última versión del código. Realiza un despliegue de "Nueva versión" en Apps Script.`;
                }
            } catch (e) {
                diagnosticsResult.textContent = `Error de conexión:\n${e.message || e}\n\nVerifica que la URL del script sea la correcta y que esté configurada para permitir accesos externos (CORS).`;
            }
        });
    }
    
    dom.btnCloseScanner.addEventListener('click', () => {
        stopScanner();
    });
    
    if (dom.btnCloseSuccess) {
        dom.btnCloseSuccess.addEventListener('click', () => {
            dom.successModal.classList.add('hidden');
            resetCheckoutFlow();
        });
    }
}

function showSection(sectionId) {
    dom.tabContents.forEach(content => {
        if (content.id === sectionId) content.classList.add('active');
        else content.classList.remove('active');
    });
    
    if (sectionId === 'section-admin') {
        if (appState.isAdminLoggedIn) {
            dom.adminLoginContainer.classList.add('hidden');
            dom.adminDashboardContainer.classList.remove('hidden');
            document.querySelector(`.admin-subnav-btn[data-admin-target="${appState.activeAdminSubtab}"]`).click();
        } else {
            dom.adminLoginContainer.classList.remove('hidden');
            dom.adminDashboardContainer.classList.add('hidden');
        }
    }
}

// ==========================================
// SELECCIÓN DE EQUIPOS
// ==========================================

function renderLoansModule() {
    const categories = [...new Set(appState.inventory.map(item => item.category))];
    
    dom.categoryTabs.innerHTML = '';
    categories.forEach(cat => {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.className = `category-tab-btn ${cat === appState.selectedCategory ? 'active' : ''}`;
        btn.textContent = cat;
        btn.addEventListener('click', () => {
            appState.selectedCategory = cat;
            document.querySelectorAll('.category-tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderEquipmentGrid();
        });
        li.appendChild(btn);
        dom.categoryTabs.appendChild(li);
    });
    
    renderEquipmentGrid();
    renderCart();
}

function renderEquipmentGrid() {
    dom.equipmentGrid.innerHTML = '';
    const filtered = appState.inventory.filter(item => item.category === appState.selectedCategory);
    
    if (filtered.length === 0) {
        dom.equipmentGrid.innerHTML = '<div class="empty-state"><p>No hay equipos registrados.</p></div>';
        return;
    }
    
    filtered.forEach(item => {
        const isSelected = appState.cart.some(cartItem => cartItem.id === item.id);
        const card = document.createElement('div');
        card.className = `equipment-item-card ${isSelected ? 'selected' : ''}`;
        
        let iconName = 'help-circle';
        if (item.category === 'Cámaras') iconName = 'camera';
        else if (item.category === 'Trípodes') iconName = 'video';
        else if (item.category === 'Audio') iconName = 'mic';
        else if (item.category === 'Luces') iconName = 'sun';
        
        const isAvailable = item.available > 0;
        const imgStyle = isAvailable ? 'width:100%; height:100%; object-fit:contain; transition: var(--transition-smooth);' : 'width:100%; height:100%; object-fit:contain; filter: grayscale(100%); opacity: 0.6; transition: var(--transition-smooth);';
        
        const mediaHtml = item.image
            ? `<div class="item-image-wrapper" style="width:100%; height:120px; overflow:hidden; border-radius:var(--border-radius-sm); margin-bottom:8px; display:flex; align-items:center; justify-content:center; background:#ffffff; border:none;">
                   <img src="${item.image}" alt="${item.name}" style="${imgStyle}">
               </div>`
            : `<div class="item-icon-wrapper" style="margin-bottom:8px;${isAvailable ? '' : ' filter: grayscale(100%); opacity: 0.5;'}">
                   <i data-lucide="${iconName}"></i>
               </div>`;
        
        card.innerHTML = `
            <div class="item-card-header" style="position:relative; display:block; width:100%;">
                ${mediaHtml}
                <span class="item-status-badge ${isAvailable ? 'available' : 'out-of-stock'}" style="position:absolute; top:8px; right:8px; z-index:2;">
                    ${isAvailable ? `${item.available} disp.` : 'Agotado'}
                </span>
            </div>
            <div class="item-card-body" style="margin-top:6px;">
                <h4 style="font-size:0.9rem; font-weight:700; margin:0 0 4px 0; line-height:1.2;">${item.name}</h4>
                ${item.description 
                    ? `<p class="item-card-desc" style="font-size:0.75rem; color:var(--text-secondary); margin:4px 0 0 0; line-height:1.3; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;">${item.description}</p>`
                    : `<p style="font-size:0.75rem; color:var(--text-secondary); margin:4px 0 0 0;">Categoría: ${item.category}</p>`}
            </div>
        `;
        
        card.addEventListener('click', () => {
            if (!isAvailable && !isSelected) {
                showToast("Sin stock disponible en este momento.", "warning");
                return;
            }
            toggleCartItem(item);
        });
        
        dom.equipmentGrid.appendChild(card);
    });
    
    lucide.createIcons();
}

function toggleCartItem(item) {
    const index = appState.cart.findIndex(cartItem => cartItem.id === item.id);
    if (index > -1) {
        appState.cart.splice(index, 1);
        showToast(`${item.name} eliminado de la selección`, "info");
    } else {
        appState.cart.push(item);
        showToast(`${item.name} seleccionado`, "success");
    }
    renderEquipmentGrid();
    renderCart();
}

function renderCart() {
    dom.selectedItems.innerHTML = '';
    
    if (appState.cart.length === 0) {
        dom.selectedItems.innerHTML = `
            <div class="empty-state">
                <i data-lucide="plus-circle"></i>
                <p>No has seleccionado ningún equipo.</p>
            </div>
        `;
        dom.inventoryCodesGroup.classList.add('hidden');
        dom.btnConfirmLoan.disabled = true;
        lucide.createIcons();
        return;
    }
    
    appState.cart.forEach(item => {
        const row = document.createElement('div');
        row.className = 'selected-item-row';
        row.innerHTML = `
            <div class="selected-item-info">
                <h5>${item.name}</h5>
                <span>${item.category}</span>
            </div>
            <button class="btn-remove-item" data-id="${item.id}">
                <i data-lucide="trash-2"></i>
            </button>
        `;
        
        row.querySelector('.btn-remove-item').addEventListener('click', (e) => {
            e.stopPropagation();
            toggleCartItem(item);
        });
        
        dom.selectedItems.appendChild(row);
    });
    
    lucide.createIcons();
    renderInventoryCodeInputs();
}

// ==========================================
// VALIDACIÓN Y SOLICITUD (EMAIL 1/3)
// ==========================================

function verifyStudent() {
    const rawRut = dom.studentRut.value.trim();
    if (!rawRut) {
        showToast("Ingresa el RUT del alumno.", "warning");
        return;
    }
    
    const formattedRut = formatRut(rawRut);
    dom.studentRut.value = formattedRut;
    
    if (CONFIG.demoMode) {
        const student = appState.students.find(s => cleanRut(s.rut) === cleanRut(formattedRut));
        if (student) {
            processStudentVerificationResult(student);
        } else {
            showToast("Estudiante no registrado en la carrera de Periodismo.", "danger");
            resetStudentValidation();
        }
    } else {
        executeVerifyStudentApi(formattedRut);
    }
}

async function executeVerifyStudentApi(rut) {
    showToast("Verificando estudiante en AVP Sheets...", "info");
    dom.btnVerifyStudent.disabled = true;
    try {
        const response = await fetch(`${CONFIG.scriptUrl}?action=checkStudent&rut=${rut}`);
        const data = await response.json();
        
        if (data.status === "success" && data.student) {
            processStudentVerificationResult(data.student);
        } else {
            showToast(data.message || "RUT no registrado.", "danger");
            resetStudentValidation();
        }
    } catch (e) {
        console.error(e);
        showToast("Error de red.", "danger");
        resetStudentValidation();
    } finally {
        dom.btnVerifyStudent.disabled = false;
    }
}

function getDaysOverdue(loan) {
    if (loan.status !== "Retirado" || !loan.progDevolucion) return 0;
    
    // Parse progDevolucion (YYYY-MM-DD)
    const parts = loan.progDevolucion.split(' ')[0].split('-');
    if (parts.length < 3) return 0;
    const dueYear = parseInt(parts[0], 10);
    const dueMonth = parseInt(parts[1], 10) - 1;
    const dueDay = parseInt(parts[2], 10);
    
    const dueDate = new Date(dueYear, dueMonth, dueDay);
    
    // Current date (midnight to be fair)
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    const diffTime = today - dueDate;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays > 0 ? diffDays : 0;
}

function checkStudentOverdueLoans(studentRut) {
    if (!studentRut) return [];
    const cleanSearch = studentRut.replace(/[^0-9kK]/g, '').toLowerCase();
    
    const studentLoans = appState.loans.filter(l => l.rut.replace(/[^0-9kK]/g, '').toLowerCase() === cleanSearch);
    const overdueLoans = [];
    
    studentLoans.forEach(loan => {
        const days = getDaysOverdue(loan);
        if (days > 0) {
            overdueLoans.push({ loan, days });
        }
    });
    
    return overdueLoans;
}

function processStudentVerificationResult(student) {
    appState.validatedStudent = student;
    dom.studentInfoBox.classList.remove('hidden');
    
    // Verificar si el alumno tiene préstamos atrasados activos
    const overdue = checkStudentOverdueLoans(student.rut);
    if (overdue.length > 0) {
        student.status = "Bloqueado";
        const listMsg = overdue.map(o => `${o.loan.item} (Debió entregarse el ${o.loan.progDevolucion.split(' ')[0]}, ${o.days} días de atraso)`).join(', ');
        student.debt = `Bloqueo Automático: Préstamo atrasado de ${listMsg}`;
    }
    
    // Unir nombre y apellido para mostrar
    const fullName = student.name + " " + (student.lastname || "");
    dom.displayStudentName.textContent = fullName;
    dom.displayStudentEmail.textContent = student.email;
    dom.displayStudentCareer.textContent = "Fono: " + (student.fono || "-");
    
    // En las columnas UFRO, si Observaciones (debt) no está vacío, el alumno está bloqueado
    if (student.status === "Bloqueado" || (student.debt && student.debt.trim() !== "")) {
        student.status = "Bloqueado"; // Forzar estado si tiene observación
        dom.studentStatusBadge.className = "student-status-badge blocked";
        dom.studentStatusText.textContent = "Bloqueado";
        dom.studentDebtWarning.classList.remove('hidden');
        dom.displayStudentDebt.textContent = student.debt || "Posee observaciones de deuda activas.";
        dom.loanDetailsGroup.classList.add('hidden');
        dom.btnConfirmLoan.disabled = true;
        showToast("Estudiante bloqueado. Proceso de solicitud congelado.", "danger");
    } else {
        dom.studentStatusBadge.className = "student-status-badge success";
        dom.studentStatusText.textContent = "Estudiante Regular";
        dom.studentDebtWarning.classList.add('hidden');
        dom.loanDetailsGroup.classList.remove('hidden');
        showToast("Estudiante regular validado.", "success");
        checkCheckoutValidity();
    }
    
    lucide.createIcons();
}

function resetStudentValidation() {
    appState.validatedStudent = null;
    dom.studentInfoBox.classList.add('hidden');
    dom.loanDetailsGroup.classList.add('hidden');
    dom.loanDateOut.value = '';
    dom.loanDateIn.value = '';
    dom.loanSubject.value = '';
    dom.btnConfirmLoan.disabled = true;
}

function resetCheckoutFlow() {
    appState.cart = [];
    appState.validatedStudent = null;
    dom.studentRut.value = '';
    dom.studentInfoBox.classList.add('hidden');
    dom.loanDetailsGroup.classList.add('hidden');
    dom.loanDateOut.value = '';
    dom.loanDateIn.value = '';
    dom.loanSubject.value = '';
    renderLoansModule();
}

function showSuccessModal(details) {
    if (dom.successLoanSummary) {
        const itemsHtml = details.items.map(it => `<li>${it.name}</li>`).join('');
        dom.successLoanSummary.innerHTML = `
            <div><strong>ID Préstamo:</strong> <span style="color:var(--primary); font-weight:700;">${details.id}</span></div>
            <div><strong>Alumno:</strong> ${details.student}</div>
            <div><strong>E-mail:</strong> ${details.email}</div>
            <div style="margin-top: 4px;">
                <strong>Equipos Solicitados:</strong>
                <ul style="margin-left: 16px; margin-top: 2px; list-style-type: disc;">
                    ${itemsHtml}
                </ul>
            </div>
            <div style="margin-top: 4px; display:flex; flex-wrap:wrap; gap:8px; justify-content:space-between; font-size:0.75rem; border-top:1px solid #e2e8f0; padding-top:8px; color:var(--text-secondary);">
                <span><strong>Retiro Programado:</strong><br>${details.progRetiro}</span>
                <span style="text-align:right; min-width:120px;"><strong>Devolución Programada:</strong><br>${details.progDevolucion}</span>
            </div>
        `;
    }
    
    if (dom.successModal) {
        dom.successModal.classList.remove('hidden');
    }
    lucide.createIcons();
}

function renderInventoryCodeInputs() {
    dom.inventoryInputsList.innerHTML = '';
    dom.inventoryCodesGroup.classList.add('hidden');
    checkCheckoutValidity();
}

function checkCheckoutValidity() {
    if (appState.cart.length > 0 && appState.validatedStudent && appState.validatedStudent.status === "Activo") {
        dom.btnConfirmLoan.disabled = false;
    } else {
        dom.btnConfirmLoan.disabled = true;
    }
}

async function processLoanCheckout() {
    if (dom.btnConfirmLoan.disabled) return;
    
    const dateOut = dom.loanDateOut.value;
    const dateIn = dom.loanDateIn.value;
    const subject = dom.loanSubject.value;
    
    if (!dateOut || !dateIn || !subject) {
        showToast("Por favor completa las fechas y selecciona la asignatura.", "warning");
        return;
    }
    
    const partsOut = dateOut.split('-').map(Number);
    const partsIn = dateIn.split('-').map(Number);
    const dateOutObj = new Date(partsOut[0], partsOut[1] - 1, partsOut[2]);
    const dateInObj = new Date(partsIn[0], partsIn[1] - 1, partsIn[2]);
    
    if (dateOutObj >= dateInObj) {
        showToast("La fecha de retiro debe ser anterior a la de entrega.", "warning");
        return;
    }
    
    // Validar duración máxima de 4 días
    const diffTime = dateInObj - dateOutObj;
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
    if (diffDays > 4) {
        showToast("El período de préstamo no puede ser superior a 4 días.", "warning");
        return;
    }
    
    const loanItems = appState.cart.map(item => {
        return {
            itemId: item.id,
            name: item.name,
            category: item.category,
            code: "Pte. Entrega"
        };
    });
    
    dom.btnConfirmLoan.disabled = true;
    
    const formattedDateOut = dateOut.replace('T', ' ');
    const formattedDateIn = dateIn.replace('T', ' ');
    
    if (CONFIG.demoMode) {
        const dateStr = getNowFormatted();
        const studentFullName = appState.validatedStudent.name + " " + (appState.validatedStudent.lastname || "");
        
        let nextIdNumber = 1006;
        if (appState.loans.length > 0) {
            const lastIds = appState.loans.map(l => {
                const numMatch = l.id.match(/\d+/);
                return numMatch ? parseInt(numMatch[0]) : 1000;
            });
            nextIdNumber = Math.max(...lastIds) + 1;
        }
        const loanId = "L-" + nextIdNumber;
        
        loanItems.forEach(loanItem => {
            // Crea el préstamo con estado "Solicitado"
            appState.loans.push({
                id: loanId,
                rut: appState.validatedStudent.rut,
                name: studentFullName,
                email: appState.validatedStudent.email,
                item: loanItem.name,
                code: loanItem.code,
                dateOut: dateStr,
                dateIn: "",
                status: "Solicitado",
                progRetiro: formattedDateOut,
                progDevolucion: formattedDateIn,
                subject: subject
            });
            
            const invItem = appState.inventory.find(i => i.id === loanItem.itemId);
            if (invItem && invItem.available > 0) {
                invItem.available--;
            }
        });
        
        saveDemoState();
        
        // Simulación Email 1 (Solicitud agrupada)
        console.log(`[SIMULACIÓN GMAIL 1/3] Enviando Comprobante de Solicitud de Préstamo a ${appState.validatedStudent.email}:`, {
            alumno: studentFullName,
            equipos: loanItems.map(i => `${i.name} [${i.code}]`),
            fechaSolicitud: dateStr,
            progRetiro: formattedDateOut,
            progDevolucion: formattedDateIn,
            subject: subject
        });
        
        showToast(`Solicitud recibida. Comprobante enviado a ${appState.validatedStudent.email}`, "success");
        showSuccessModal({
            id: loanId,
            student: studentFullName,
            email: appState.validatedStudent.email,
            items: loanItems,
            progRetiro: formattedDateOut,
            progDevolucion: formattedDateIn
        });
        loadData();
    } else {
        executeCheckoutApi(loanItems, formattedDateOut, formattedDateIn, subject);
    }
}

async function executeCheckoutApi(loanItems, progRetiro, progDevolucion, subject) {
    showToast("Enviando solicitud a Google Sheets...", "info");
    try {
        const payload = {
            student: appState.validatedStudent,
            items: loanItems,
            timestamp: getNowFormatted(),
            progRetiro: progRetiro,
            progDevolucion: progDevolucion,
            subject: subject
        };
        
        const response = await fetch(`${CONFIG.scriptUrl}?action=createLoan`, {
            method: 'POST',
            mode: 'cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        
        const textData = await response.text();
        let data;
        try {
            data = JSON.parse(textData);
        } catch (parseErr) {
            console.error("La respuesta del script no es un JSON válido:", textData);
            if (textData.includes("html") || textData.includes("Error") || textData.includes("authorization")) {
                throw new Error("El script de Google Sheets arrojó un error de ejecución. Ejecuta la función testConnection en el editor de Apps Script para verificar que los permisos de envío estén otorgados.");
            }
            throw new Error("Respuesta inválida del servidor.");
        }
        
        if (data.status === "success") {
            showToast(data.message || `Solicitud registrada con éxito.`, "success");
            const studentFullName = appState.validatedStudent.name + " " + (appState.validatedStudent.lastname || "");
            showSuccessModal({
                id: data.loanId || "Solicitud registrada",
                student: studentFullName,
                email: appState.validatedStudent.email,
                items: loanItems,
                progRetiro: progRetiro,
                progDevolucion: progDevolucion
            });
            loadData();
        } else {
            showToast(data.message || "Error al solicitar préstamo.", "danger");
            dom.btnConfirmLoan.disabled = false;
        }
    } catch (e) {
        console.error(e);
        showToast(`Error de conexión: ${e.message || "Error de red."}`, "danger");
        dom.btnConfirmLoan.disabled = false;
    }
}

// ==========================================
// ENTREGA FÍSICA Y CAMBIO A RETIRADO (EMAIL 2/3)
// ==========================================

function openDeliveryModal(loanId) {
    appState.currentDeliveryLoanId = loanId;
    if (dom.deliveryGlobalObs) {
        dom.deliveryGlobalObs.value = '';
    }
    
    const loanItems = appState.loans.filter(l => l.id === loanId && l.status === "Solicitado");
    if (loanItems.length === 0) {
        showToast("No hay equipos pendientes de entrega para esta solicitud.", "warning");
        return;
    }
    
    appState.tempDeliveryItems = loanItems.map(item => ({ ...item, isNew: false }));
    renderDeliveryModalInputs();
    
    dom.deliveryModal.classList.remove('hidden');
}

function renderDeliveryModalInputs() {
    const loanId = appState.currentDeliveryLoanId;
    dom.deliveryInputsContainer.innerHTML = '';
    
    appState.tempDeliveryItems.forEach((loanItem, index) => {
        const itemInv = appState.inventory.find(i => i.name === loanItem.item);
        const allCodes = itemInv && itemInv.codes ? itemInv.codes : [];
        
        const rentedCodes = appState.loans
            .filter(l => l.item === loanItem.item && l.status === "Retirado")
            .map(l => l.code.toUpperCase().trim());
            
        const selectedInModal = [];
        for (let i = 0; i < index; i++) {
            const codeEl = document.getElementById(`delivery-code-input-${i}`);
            if (codeEl) {
                selectedInModal.push(codeEl.value.toUpperCase().trim());
            }
        }
            
        const availableCodes = allCodes.filter(c => !rentedCodes.includes(c.toUpperCase().trim()) && !selectedInModal.includes(c.toUpperCase().trim()));
        
        const row = document.createElement('div');
        row.className = 'inventory-code-row';
        row.style.marginBottom = '20px';
        row.style.padding = '12px';
        row.style.border = '1px solid var(--glass-border)';
        row.style.borderRadius = '8px';
        row.style.background = 'rgba(15, 23, 42, 0.01)';
        
        let selectOptionsHtml = '<option value="">-- Seleccionar código disponible --</option>';
        availableCodes.forEach(code => {
            selectOptionsHtml += `<option value="${code}">${code}</option>`;
        });
        selectOptionsHtml += `<option value="custom">Otro / Escanear código...</option>`;
        
        const itemLabel = loanItem.isNew ? `[EXTRA] ${loanItem.item}` : loanItem.item;
        
        row.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; border-bottom: 1px solid var(--glass-border); padding-bottom: 6px;">
                <label style="font-weight: 700; color: var(--text-primary); font-size: 0.95rem;">
                    ${index + 1}. ${itemLabel}
                </label>
                <button class="btn btn-danger btn-icon-only btn-remove-item-delivery" data-index="${index}" style="padding: 4px 8px; border-radius: 4px; font-size: 0.72rem; display: flex; align-items: center; gap: 4px; border: 1px solid rgba(239, 68, 68, 0.2);" title="Anular este equipo">
                    <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i> Anular
                </button>
            </div>
            
            <div style="margin-bottom: 12px;">
                <label style="font-size: 0.75rem; color: var(--text-secondary); font-weight: 600;">Seleccionar Código en Stock:</label>
                <select id="delivery-code-select-${index}" class="select-input" style="width: 100%; margin-top: 4px; padding: 8px; border-radius: 6px; border: 1px solid var(--glass-border);">
                    ${selectOptionsHtml}
                </select>
            </div>
            
            <div id="custom-code-container-${index}" style="display: none; margin-top: 8px;">
                <label style="font-size: 0.75rem; color: var(--text-secondary); font-weight: 600;">Escribir o Escanear Código:</label>
                <div class="input-with-button" style="margin-top: 4px;">
                    <input type="text" id="delivery-code-input-${index}" placeholder="Escribe o escanea el código" class="select-input" style="background:#fff;" autocomplete="off">
                    <button class="btn btn-secondary btn-icon-only" id="btn-scan-delivery-trigger-${index}">
                        <i data-lucide="scan-line"></i>
                    </button>
                </div>
            </div>
        `;
        
        const select = row.querySelector(`#delivery-code-select-${index}`);
        const customContainer = row.querySelector(`#custom-code-container-${index}`);
        const input = row.querySelector(`#delivery-code-input-${index}`);
        const scanBtn = row.querySelector(`#btn-scan-delivery-trigger-${index}`);
        const removeBtn = row.querySelector('.btn-remove-item-delivery');
        
        if (availableCodes.length > 0) {
            select.value = availableCodes[0];
            input.value = availableCodes[0];
        } else {
            select.value = "custom";
            customContainer.style.display = "block";
        }
        
        select.addEventListener('change', (e) => {
            if (e.target.value === 'custom') {
                customContainer.style.display = "block";
                input.value = '';
                input.focus();
            } else {
                customContainer.style.display = "none";
                input.value = e.target.value;
            }
        });
        
        scanBtn.addEventListener('click', (e) => {
            e.preventDefault();
            startScanner(`delivery-code-input-${index}`);
        });
        
        removeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const idx = parseInt(e.currentTarget.getAttribute('data-index'));
            const targetItem = appState.tempDeliveryItems[idx];
            if (targetItem.isNew) {
                appState.tempDeliveryItems.splice(idx, 1);
                renderDeliveryModalInputs();
            } else {
                removeDeliveryItem(loanId, targetItem.item);
            }
        });
        
        dom.deliveryInputsContainer.appendChild(row);
    });
    
    if (dom.deliveryAddEquipmentSelect) {
        dom.deliveryAddEquipmentSelect.innerHTML = '';
        
        const addable = appState.inventory.filter(item => {
            const tempAddedCount = appState.tempDeliveryItems.filter(temp => temp.item === item.name && temp.isNew).length;
            const availableCount = item.available - tempAddedCount;
            return availableCount > 0;
        });
        
        if (addable.length === 0) {
            const opt = document.createElement('option');
            opt.value = "";
            opt.textContent = "No hay equipos con stock disponible";
            opt.disabled = true;
            opt.selected = true;
            dom.deliveryAddEquipmentSelect.appendChild(opt);
        } else {
            const defOpt = document.createElement('option');
            defOpt.value = "";
            defOpt.textContent = "-- Seleccionar equipo a agregar --";
            defOpt.disabled = true;
            defOpt.selected = true;
            dom.deliveryAddEquipmentSelect.appendChild(defOpt);
            
            addable.forEach(item => {
                const tempAddedCount = appState.tempDeliveryItems.filter(temp => temp.item === item.name && temp.isNew).length;
                const availableCount = item.available - tempAddedCount;
                
                const opt = document.createElement('option');
                opt.value = item.name;
                opt.textContent = `${item.name} (Disponibles: ${availableCount})`;
                dom.deliveryAddEquipmentSelect.appendChild(opt);
            });
        }
    }
    
    lucide.createIcons();
}

async function confirmPhysicalDelivery() {
    const loanId = appState.currentDeliveryLoanId;
    if (!loanId || !appState.tempDeliveryItems || appState.tempDeliveryItems.length === 0) return;
    
    const globalObsVal = dom.deliveryGlobalObs ? dom.deliveryGlobalObs.value.trim() : '';
    
    const itemsPayload = [];
    const selectedCodes = new Set();
    
    for (let index = 0; index < appState.tempDeliveryItems.length; index++) {
        const inputVal = document.getElementById(`delivery-code-input-${index}`).value.trim().toUpperCase();
        if (!inputVal) {
            showToast(`Ingresa el código para el equipo: ${appState.tempDeliveryItems[index].item}`, "warning");
            return;
        }
        if (selectedCodes.has(inputVal)) {
            showToast(`El código de inventario "${inputVal}" está asignado más de una vez en esta entrega.`, "warning");
            return;
        }
        selectedCodes.add(inputVal);
        
        itemsPayload.push({
            name: appState.tempDeliveryItems[index].item,
            code: inputVal,
            obs: globalObsVal,
            isNew: !!appState.tempDeliveryItems[index].isNew
        });
    }
    
    dom.deliveryModal.classList.add('hidden');
    
    if (CONFIG.demoMode) {
        const dateStr = getNowFormatted();
        
        appState.tempDeliveryItems.forEach((tempItem, index) => {
            const assigned = itemsPayload[index];
            if (tempItem.isNew) {
                const newLoanRecord = {
                    id: loanId,
                    rut: tempItem.rut,
                    name: tempItem.name,
                    email: tempItem.email,
                    item: tempItem.item,
                    code: assigned.code,
                    status: "Retirado",
                    dateOut: tempItem.dateOut,
                    dateDeliver: dateStr,
                    dateIn: "",
                    progRetiro: tempItem.progRetiro,
                    progDevolucion: tempItem.progDevolucion,
                    subject: tempItem.subject,
                    obs: globalObsVal,
                    obsReturn: "",
                    daysOverdue: 0
                };
                appState.loans.push(newLoanRecord);
                
                const invItem = appState.inventory.find(i => i.name === tempItem.item);
                if (invItem) {
                    invItem.available = Math.max(0, invItem.available - 1);
                }
            } else {
                const loanItem = appState.loans.find(l => l.id === loanId && l.item === tempItem.item && l.status === "Solicitado");
                if (loanItem) {
                    loanItem.status = "Retirado";
                    loanItem.code = assigned.code;
                    loanItem.obs = globalObsVal;
                    loanItem.dateDeliver = dateStr;
                }
            }
        });
        
        saveDemoState();
        
        console.log(`[SIMULACIÓN GMAIL 2/3] Enviando Comprobante de Retiro Físico a ${appState.tempDeliveryItems[0].email}:`, {
            alumno: appState.tempDeliveryItems[0].name,
            equipos: itemsPayload,
            fechaRetiro: dateStr,
            loanId: loanId
        });
        
        showToast(`Retiro físico confirmado. Comprobante enviado a ${appState.tempDeliveryItems[0].email}`, "success");
        loadData();
    } else {
        executeDeliverLoanApi(loanId, itemsPayload);
    }
}

async function executeDeliverLoanApi(loanId, items) {
    showToast("Confirmando entrega en Google Sheets...", "info");
    try {
        const payload = {
            loanId: loanId,
            items: items,
            timestamp: getNowFormatted()
        };
        
        const response = await fetch(`${CONFIG.scriptUrl}?action=deliverLoan`, {
            method: 'POST',
            mode: 'cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        
        const data = await response.json();
        if (data.status === "success") {
            showToast("Retiro registrado y correo enviado correctamente", "success");
            loadData();
        } else {
            showToast(data.message || "Error al registrar entrega.", "danger");
        }
    } catch (e) {
        console.error(e);
        showToast("Error de conexión.", "danger");
    }
}

async function removeDeliveryItem(loanId, itemName) {
    if (!confirm(`¿Anular la solicitud del equipo "${itemName}" para este préstamo? El resto de los equipos se mantendrán en la solicitud.`)) {
        return;
    }
    
    if (CONFIG.demoMode) {
        const loanIndex = appState.loans.findIndex(l => l.id === loanId && l.item === itemName && l.status === "Solicitado");
        if (loanIndex > -1) {
            appState.loans[loanIndex].status = "Anulado";
            if (appState.loans[loanIndex].code === "Pte. Entrega" || !appState.loans[loanIndex].code) {
                appState.loans[loanIndex].code = "Anulado";
            }
            
            // Reincorporar stock
            const invItem = appState.inventory.find(i => i.name === itemName);
            if (invItem) {
                invItem.available = Math.min(invItem.total, invItem.available + 1);
            }
            
            saveDemoState();
            showToast(`Solicitud de "${itemName}" anulada correctamente.`, "info");
            
            await loadData();
            
            const remainingItems = appState.loans.filter(l => l.id === loanId && l.status === "Solicitado");
            if (remainingItems.length > 0) {
                openDeliveryModal(loanId);
            } else {
                dom.deliveryModal.classList.add('hidden');
            }
        }
    } else {
        showToast("Anulando equipo en Google Sheets...", "info");
        try {
            const payload = {
                loanId: loanId,
                itemName: itemName,
                timestamp: getNowFormatted()
            };
            
            const response = await fetch(`${CONFIG.scriptUrl}?action=cancelLoan`, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload)
            });
            
            const data = await response.json();
            if (data.status === "success") {
                showToast(`Equipo "${itemName}" anulado con éxito`, "success");
                await loadData();
                
                const remainingItems = appState.loans.filter(l => l.id === loanId && l.status === "Solicitado");
                if (remainingItems.length > 0) {
                    openDeliveryModal(loanId);
                } else {
                    dom.deliveryModal.classList.add('hidden');
                }
            } else {
                showToast(data.message || "Error al anular el equipo.", "danger");
            }
        } catch (e) {
            console.error(e);
            showToast("Error de conexión.", "danger");
        }
    }
}

// ==========================================
// DEVOLUCIÓN Y ANULACIÓN DE EQUIPOS
// ==========================================

function openReturnModal(loanId) {
    appState.currentReturnLoanId = loanId;
    const loanItems = appState.loans.filter(l => l.id === loanId && l.status === "Retirado");
    if (loanItems.length === 0) {
        showToast("No se encontraron equipos en tránsito para esta solicitud.", "warning");
        return;
    }
    
    if (dom.returnItemsContainer) {
        dom.returnItemsContainer.innerHTML = loanItems.map(item => `
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(0,0,0,0.05); padding-bottom:6px;">
                <span style="font-weight:600; font-size:0.85rem;">• ${item.item}</span>
                <span class="loan-code-badge" style="margin:0;">${item.code}</span>
            </div>
        `).join('');
    }
    
    if (dom.returnGlobalObs) {
        dom.returnGlobalObs.value = '';
    }
    
    if (dom.returnModal) {
        dom.returnModal.classList.remove('hidden');
        lucide.createIcons();
    }
}

async function confirmReturnCheckout() {
    const loanId = appState.currentReturnLoanId;
    if (!loanId) return;
    
    const obsReturnVal = dom.returnGlobalObs ? dom.returnGlobalObs.value.trim() : '';
    
    if (dom.returnModal) {
        dom.returnModal.classList.add('hidden');
    }
    
    if (CONFIG.demoMode) {
        const dateStr = getNowFormatted();
        const loanItems = appState.loans.filter(l => l.id === loanId && l.status === "Retirado");
        if (loanItems.length === 0) return;
        
        const returnedItems = [];
        loanItems.forEach(loan => {
            const days = getDaysOverdue(loan);
            loan.status = "Devuelto";
            loan.dateIn = dateStr;
            loan.daysOverdue = days;
            loan.obsReturn = obsReturnVal;
            
            const invItem = appState.inventory.find(i => i.name === loan.item);
            if (invItem) {
                invItem.available = Math.min(invItem.total, invItem.available + 1);
            }
            returnedItems.push({ name: loan.item, code: loan.code });
        });
        
        saveDemoState();
        
        // Simulación Email 3 (Devolución)
        console.log(`[SIMULACIÓN GMAIL 3/3] Enviando Comprobante de Devolución a ${loanItems[0].email}:`, {
            alumno: loanItems[0].name,
            equipos: returnedItems,
            fechaRetorno: dateStr,
            loanId: loanId,
            obsReturn: obsReturnVal
        });
        
        showToast(`Devolución exitosa. Comprobante enviado a ${loanItems[0].email}`, "success");
        await loadData();
        renderAdminLoans(appState.activeAdminLoanFilter);
    } else {
        executeReturnApi(loanId, obsReturnVal);
    }
}

async function executeReturnApi(loanId, obsReturn) {
    showToast("Procesando devolución en AVP Sheets...", "info");
    try {
        const payload = {
            loanId: loanId,
            timestamp: getNowFormatted(),
            obsReturn: obsReturn
        };
        
        const response = await fetch(`${CONFIG.scriptUrl}?action=returnLoan`, {
            method: 'POST',
            mode: 'cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        
        const data = await response.json();
        if (data.status === "success") {
            showToast("Devolución procesada e email enviado", "success");
            await loadData();
            renderAdminLoans(appState.activeAdminLoanFilter);
        } else {
            showToast(data.message || "Error al procesar devolución.", "danger");
        }
    } catch (e) {
        console.error(e);
        showToast("Error de conexión.", "danger");
    }
}

async function processCancelLoan(loanId) {
    if (!confirm(`¿Estás seguro de que deseas anular la solicitud/préstamo completa ${loanId}? Esta acción es irreversible y devolverá los equipos al stock.`)) {
        return;
    }
    
    if (CONFIG.demoMode) {
        const loanItems = appState.loans.filter(l => l.id === loanId);
        const activeItems = loanItems.filter(l => l.status === "Solicitado" || l.status === "Retirado");
        if (activeItems.length === 0) {
            showToast("No se encontraron ítems activos para anular.", "warning");
            return;
        }
        
        const cancelledItems = [];
        activeItems.forEach(loan => {
            const wasRetirado = loan.status === "Retirado";
            loan.status = "Anulado";
            if (loan.code === "Pte. Entrega" || !loan.code) {
                loan.code = "Anulado";
            }
            
            const invItem = appState.inventory.find(i => i.name === loan.item);
            if (invItem) {
                invItem.available = Math.min(invItem.total, invItem.available + 1);
            }
            cancelledItems.push({ name: loan.item, code: loan.code });
        });
        
        saveDemoState();
        
        // Simulación Email
        console.log(`[SIMULACIÓN GMAIL] Enviando Notificación de Anulación a ${activeItems[0].email}:`, {
            alumno: activeItems[0].name,
            equipos: cancelledItems,
            fecha: getNowFormatted(),
            loanId: loanId
        });
        
        showToast("Préstamo anulado con éxito en modo Demo.", "success");
        await loadData();
        renderAdminLoans(appState.activeAdminLoanFilter);
    } else {
        executeCancelApi(loanId);
    }
}

async function executeCancelApi(loanId) {
    showToast("Procesando anulación en AVP Sheets...", "info");
    try {
        const payload = {
            loanId: loanId,
            timestamp: getNowFormatted()
        };
        
        const response = await fetch(`${CONFIG.scriptUrl}?action=cancelLoan`, {
            method: 'POST',
            mode: 'cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        
        const data = await response.json();
        if (data.status === "success") {
            showToast("Préstamo anulado con éxito", "success");
            await loadData();
            renderAdminLoans(appState.activeAdminLoanFilter);
        } else {
            showToast(data.message || "Error al anular el préstamo.", "danger");
        }
    } catch (e) {
        console.error(e);
        showToast("Error de conexión.", "danger");
    }
}

// ==========================================
// PANEL DE ADMINISTRACIÓN Y SUBTABS
// ==========================================

function updateAdminDashboard() {
    if (!appState.isAdminLoggedIn) return;
    
    const totalEquipos = appState.inventory.reduce((sum, item) => sum + item.total, 0);
    const uniqueLoans = groupLoansById(appState.loans);
    const prestamosActivos = uniqueLoans.filter(l => l.status === "Retirado" || l.status === "Solicitado").length;
    const prestamosAtrasados = uniqueLoans.filter(l => l.status === "Retirado" && getDaysOverdue(l) > 0).length;
    // Si un alumno tiene observaciones (debt), se considera bloqueado
    const alumnosBloqueados = appState.students.filter(s => s.status === "Bloqueado" || (s.debt && s.debt.trim() !== "")).length;
    const totalHistorico = new Set(appState.loans.map(l => l.id)).size;
    
    dom.metricTotalEquipos.textContent = totalEquipos;
    dom.metricPrestamosActivos.textContent = prestamosActivos;
    if (dom.metricPrestamosAtrasados) {
        dom.metricPrestamosAtrasados.textContent = prestamosAtrasados;
    }
    dom.metricAlumnosBloqueados.textContent = alumnosBloqueados;
    dom.metricTotalHistorico.textContent = totalHistorico;
    
    const hasValidSheetUrl = CONFIG.sheetUrl && !CONFIG.sheetUrl.includes('1_YOUR_SHEET_ID');
    if (hasValidSheetUrl) {
        dom.linkGoogleSheet.href = CONFIG.sheetUrl;
        dom.linkGoogleSheet.classList.remove('hidden');
    } else {
        dom.linkGoogleSheet.classList.add('hidden');
    }
}

function getGroupStatus(items) {
    if (items.some(it => it.status === "Retirado")) return "Retirado";
    if (items.some(it => it.status === "Solicitado")) return "Solicitado";
    if (items.every(it => it.status === "Anulado")) return "Anulado";
    return "Devuelto";
}

function groupLoansById(loans) {
    const groups = {};
    loans.forEach(loan => {
        if (!groups[loan.id]) {
            groups[loan.id] = {
                id: loan.id,
                rut: loan.rut,
                name: loan.name,
                email: loan.email,
                dateOut: loan.dateOut,
                dateDeliver: loan.dateDeliver || "",
                dateIn: loan.dateIn,
                progRetiro: loan.progRetiro,
                progDevolucion: loan.progDevolucion,
                subject: loan.subject,
                obs: loan.obs || "",
                daysOverdue: parseInt(loan.daysOverdue) || 0,
                obsReturn: loan.obsReturn || "",
                items: []
            };
        } else if (loan.obs && !groups[loan.id].obs) {
            groups[loan.id].obs = loan.obs;
        }
        if (loan.obsReturn && !groups[loan.id].obsReturn) {
            groups[loan.id].obsReturn = loan.obsReturn;
        }
        groups[loan.id].items.push({
            name: loan.item,
            code: loan.code,
            status: loan.status,
            dateIn: loan.dateIn,
            obs: loan.obs || ""
        });
    });
    
    return Object.values(groups).map(group => {
        group.status = getGroupStatus(group.items);
        const returnedItems = group.items.filter(it => it.status === "Devuelto" && it.dateIn);
        if (returnedItems.length > 0) {
            group.dateIn = returnedItems[0].dateIn;
        }
        return group;
    });
}

function renderAdminLoans(filter = "all") {
    dom.adminLoansTableBody.innerHTML = '';
    
    const grouped = groupLoansById(appState.loans);
    
    let filtered = grouped;
    if (filter === "solicitado") {
        filtered = grouped.filter(l => l.status === "Solicitado");
    } else if (filter === "retirado") {
        filtered = grouped.filter(l => l.status === "Retirado");
    } else if (filter === "returned") {
        filtered = grouped.filter(l => l.status === "Devuelto" || l.status === "Anulado");
    }
    
    filtered = [...filtered].reverse();
    
    if (filtered.length === 0) {
        dom.adminLoansTableBody.innerHTML = `
            <tr>
                <td colspan="11" class="table-empty">
                    <i data-lucide="list"></i>
                    <p>No hay solicitudes registradas.</p>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }
    
    filtered.forEach(loan => {
        let actionBtn = '-';
        if (loan.status === 'Solicitado') {
            actionBtn = `
                <div class="action-buttons-cell" style="display:flex; gap:8px; justify-content:center;">
                    <button class="btn btn-primary btn-icon-only btn-deliver-loan" data-loan-id="${loan.id}" title="Confirmar Retiro Físico">
                        <i data-lucide="check-square"></i> Entregar
                    </button>
                    <button class="btn btn-danger btn-icon-only btn-cancel-loan" data-loan-id="${loan.id}" title="Anular Solicitud">
                        <i data-lucide="x-circle"></i> Anular
                    </button>
                </div>
            `;
        } else if (loan.status === 'Retirado') {
            actionBtn = `
                <div class="action-buttons-cell" style="display:flex; gap:8px; justify-content:center;">
                    <button class="btn btn-success btn-icon-only btn-return-loan" data-loan-id="${loan.id}" title="Registrar Devolución">
                        <i data-lucide="rotate-ccw"></i> Devolver
                    </button>
                    <button class="btn btn-danger btn-icon-only btn-cancel-loan" data-loan-id="${loan.id}" title="Anular Préstamo">
                        <i data-lucide="x-circle"></i> Anular
                    </button>
                </div>
            `;
        }
        
        const daysOverdue = getDaysOverdue(loan);
        const finalOverdue = loan.status === 'Retirado' ? daysOverdue : (loan.daysOverdue || 0);
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong>${loan.id}</strong></td>
            <td><code>${loan.rut}</code></td>
            <td>${loan.name}</td>
            <td>
                <div class="loan-items-list">
                    ${loan.items.map(it => `<span class="loan-item-entry">• ${it.name}</span>`).join('')}
                    ${loan.subject ? `<small style="color:var(--text-secondary); font-size:0.72rem; font-weight:500; margin-top:2px;">Asig: ${loan.subject}</small>` : ''}
                    ${loan.obs ? `<small style="color:var(--text-secondary); font-style:italic; font-size:0.72rem; font-weight:500; margin-top:2px; word-break:break-word; max-width:200px; line-height:1.2; display:block;">Obs: ${loan.obs}</small>` : ''}
                    ${loan.obsReturn ? `<small style="color:var(--danger); font-style:italic; font-size:0.72rem; font-weight:500; margin-top:2px; word-break:break-word; max-width:200px; line-height:1.2; display:block;">Obs. Dev: ${loan.obsReturn}</small>` : ''}
                </div>
            </td>
            <td>
                <div class="loan-codes-list">
                    ${loan.items.map(it => {
                        let displayCode = it.code;
                        let extraStyle = "";
                        if ((it.status === "Anulado" || loan.status === "Anulado") && (displayCode === "Pte. Entrega" || !displayCode)) {
                            displayCode = "Anulado";
                            extraStyle = "background: rgba(239, 68, 68, 0.05); border-color: rgba(239, 68, 68, 0.15); color: var(--danger);";
                        }
                        return `<span class="loan-code-badge" style="${extraStyle}">${displayCode}</span>`;
                    }).join('')}
                </div>
            </td>
            <td>
                <div style="display:flex; flex-direction:column; align-items:flex-start;">
                    <span>${loan.dateOut || '-'}</span>
                </div>
            </td>
            <td>
                <div style="display:flex; flex-direction:column; align-items:flex-start;">
                    <span>${loan.status === 'Solicitado' ? 'Pendiente' : (loan.dateDeliver || loan.dateOut || '-')}</span>
                </div>
            </td>
            <td>
                <div style="display:flex; flex-direction:column; align-items:flex-start;">
                    <span>${loan.status === 'Devuelto' ? loan.dateIn : (loan.status === 'Retirado' ? 'En Tránsito' : '-')}</span>
                    ${loan.status === 'Solicitado' && loan.progDevolucion ? `<small style="color:var(--text-secondary); font-size:0.72rem; font-weight:500;" title="Devolución programada">Dev: ${loan.progDevolucion}</small>` : ''}
                </div>
            </td>
            <td>
                <div style="display:flex; flex-direction:column; align-items:flex-start;">
                    ${(() => {
                        if (loan.status === 'Solicitado' || loan.status === 'Anulado') {
                            return '<span style="color:var(--text-secondary); font-weight:500;">-</span>';
                        }
                        if (loan.status === 'Retirado') {
                            if (daysOverdue > 0) {
                                return `
                                    <span class="table-badge danger" style="background: rgba(239, 68, 68, 0.08); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.2); font-size: 0.7rem; font-weight: 700; display:flex; align-items:center; gap:3px; text-align:left; justify-content:flex-start; margin-left:0; width:fit-content;">
                                        <i data-lucide="alert-triangle" style="width:11px; height:11px; flex-shrink:0;"></i> +${daysOverdue} d atraso
                                    </span>
                                `;
                            } else {
                                if (loan.progDevolucion) {
                                    const progParts = loan.progDevolucion.split(' ')[0].split('-');
                                    if (progParts.length >= 3) {
                                        const progDate = new Date(parseInt(progParts[0], 10), parseInt(progParts[1], 10) - 1, parseInt(progParts[2], 10));
                                        const now = new Date();
                                        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
                                        const diffTime = progDate - today;
                                        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                                        if (diffDays > 0) {
                                            return `
                                                <span class="table-badge success" style="background: rgba(16, 185, 129, 0.08); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.2); font-size: 0.7rem; font-weight: 700; display:flex; align-items:center; gap:3px; text-align:left; justify-content:flex-start; margin-left:0; width:fit-content;">
                                                    <i data-lucide="check" style="width:11px; height:11px; flex-shrink:0;"></i> Quedan ${diffDays} d
                                                </span>
                                            `;
                                        } else if (diffDays === 0) {
                                            return `
                                                <span class="table-badge warning" style="background: rgba(245, 158, 11, 0.08); color: #d97706; border: 1px solid rgba(245, 158, 11, 0.2); font-size: 0.7rem; font-weight: 700; display:flex; align-items:center; gap:3px; text-align:left; justify-content:flex-start; margin-left:0; width:fit-content;">
                                                    <i data-lucide="clock" style="width:11px; height:11px; flex-shrink:0;"></i> Vence hoy
                                                </span>
                                            `;
                                        }
                                    }
                                }
                                return '<span class="table-badge success" style="background: rgba(16, 185, 129, 0.08); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.2); font-size: 0.7rem; font-weight: 700;">A tiempo</span>';
                            }
                        }
                        if (loan.status === 'Devuelto') {
                            if (finalOverdue > 0) {
                                return `
                                    <span class="table-badge danger" style="background: rgba(239, 68, 68, 0.08); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.2); font-size: 0.7rem; font-weight: 700; display:flex; align-items:center; gap:3px; text-align:left; justify-content:flex-start; margin-left:0; width:fit-content;">
                                        <i data-lucide="alert-triangle" style="width:11px; height:11px; flex-shrink:0;"></i> Tarde (${finalOverdue} d)
                                    </span>
                                `;
                            } else {
                                return `
                                    <span class="table-badge success" style="background: rgba(16, 185, 129, 0.08); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.2); font-size: 0.7rem; font-weight: 700; display:flex; align-items:center; gap:3px; text-align:left; justify-content:flex-start; margin-left:0; width:fit-content;">
                                        <i data-lucide="check" style="width:11px; height:11px; flex-shrink:0;"></i> Al día (0 d)
                                    </span>
                                `;
                            }
                        }
                        return '-';
                    })()}
                </div>
            </td>
            <td>
                <div style="display:flex; flex-direction:column; align-items:center; gap:2px;">
                    <span class="table-badge ${loan.status.toLowerCase()}">${loan.status}</span>
                </div>
            </td>
            <td>${actionBtn}</td>
        `;
        
        if (loan.status === 'Solicitado') {
            row.querySelector('.btn-deliver-loan').addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-loan-id');
                openDeliveryModal(id);
            });
            row.querySelector('.btn-cancel-loan').addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-loan-id');
                processCancelLoan(id);
            });
        } else if (loan.status === 'Retirado') {
            row.querySelector('.btn-return-loan').addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-loan-id');
                openReturnModal(id);
            });
            row.querySelector('.btn-cancel-loan').addEventListener('click', (e) => {
                const id = e.currentTarget.getAttribute('data-loan-id');
                processCancelLoan(id);
            });
        }
        
        dom.adminLoansTableBody.appendChild(row);
    });
    
    lucide.createIcons();
}

function renderSubjectsDropdown() {
    const select = document.getElementById('loan-subject');
    if (!select) return;
    
    select.innerHTML = '<option value="" disabled selected>Selecciona una asignatura...</option>';
    
    const subjects = appState.subjects && appState.subjects.length > 0
        ? appState.subjects
        : ["Periodismo Escrito", "Periodismo Radial", "Periodismo Televisivo", "Fotoperiodismo", "Comunicación Digital", "Cine y Documental", "Proyecto de Título", "Ninguna (Proyecto Personal)"];
        
    subjects.forEach(sub => {
        const option = document.createElement('option');
        option.value = sub;
        option.textContent = sub;
        select.appendChild(option);
    });
}

function updateCharts() {
    if (!appState.isAdminLoggedIn) return;
    
    // Configurar y poblar dinámicamente el selector de categorías del gráfico popular
    if (dom.chartPopularCategory) {
        const currentSelected = dom.chartPopularCategory.value;
        const categories = [...new Set(appState.inventory.map(item => item.category))].filter(Boolean);
        
        // Reconstruir opciones si cambiaron las categorías de los equipos
        const optionValues = Array.from(dom.chartPopularCategory.options).map(o => o.value);
        const categoriesChanged = categories.length !== optionValues.length - 1 || 
                                  !categories.every(cat => optionValues.includes(cat));
                                  
        if (categoriesChanged) {
            dom.chartPopularCategory.innerHTML = '<option value="all">Todas las Categorías</option>';
            categories.forEach(cat => {
                const opt = document.createElement('option');
                opt.value = cat;
                opt.textContent = cat;
                dom.chartPopularCategory.appendChild(opt);
            });
            
            if (categories.includes(currentSelected)) {
                dom.chartPopularCategory.value = currentSelected;
            } else {
                dom.chartPopularCategory.value = "all";
            }
        }
    }
    
    const selectedCat = dom.chartPopularCategory ? dom.chartPopularCategory.value : "all";
    
    const itemCounts = {};
    appState.loans.forEach(loan => {
        const invItem = appState.inventory.find(i => i.name.toLowerCase().trim() === loan.item.toLowerCase().trim());
        const category = invItem ? invItem.category : "Otros";
        
        if (selectedCat === "all" || category === selectedCat) {
            itemCounts[loan.item] = (itemCounts[loan.item] || 0) + 1;
        }
    });
    
    const sortedItems = Object.entries(itemCounts)
        .sort((a, b) => b[1] - a[1]);
        
    const popularLabels = sortedItems.map(item => item[0]);
    const popularData = sortedItems.map(item => item[1]);
    renderBarChart('chart-popular-items', 'popularItems', popularLabels, popularData, 'Préstamos Totales', '#4f46e5');
    
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const flowDays = { 'Lunes': 0, 'Martes': 0, 'Miércoles': 0, 'Jueves': 0, 'Viernes': 0, 'Sábado': 0, 'Domingo': 0 };
    
    appState.loans.forEach(loan => {
        if (loan.dateOut) {
            const cleanDate = loan.dateOut.split(' ')[0];
            const d = new Date(cleanDate);
            if (!isNaN(d.getDay())) {
                const dayName = dayNames[d.getDay()];
                flowDays[dayName] = (flowDays[dayName] || 0) + 1;
            }
        }
    });
    renderLineChart('chart-flow-days', 'flowDays', Object.keys(flowDays), Object.values(flowDays), 'Préstamos', '#7c3aed');
    
    const studentCounts = {};
    appState.loans.forEach(loan => {
        studentCounts[loan.name] = (studentCounts[loan.name] || 0) + 1;
    });
    
    const sortedStudents = Object.entries(studentCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);
        
    const studentLabels = sortedStudents.map(s => s[0]);
    const studentData = sortedStudents.map(s => s[1]);
    renderHorizontalBarChart('chart-top-students', 'topStudents', studentLabels, studentData, 'Préstamos', '#10b981');

    // Préstamos por Asignatura (Doughnut)
    const subjectCounts = {};
    appState.loans.forEach(loan => {
        const sub = loan.subject || "No Especificada";
        subjectCounts[sub] = (subjectCounts[sub] || 0) + 1;
    });
    const sortedSubjects = Object.entries(subjectCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);
    const subjectLabels = sortedSubjects.map(s => s[0]);
    const subjectData = sortedSubjects.map(s => s[1]);
    const subjectColors = ['#4f46e5', '#7c3aed', '#10b981', '#f59e0b', '#ef4444', '#06b6d4', '#ec4899', '#64748b'];
    renderDoughnutChart('chart-subject-distribution', 'subjectDistribution', subjectLabels, subjectData, subjectColors);
    
    // Cumplimiento y Retrasos (Doughnut)
    let devueltoOK = 0;
    let devueltoLate = 0;
    let retiradoOK = 0;
    let retiradoLate = 0;
    appState.loans.forEach(loan => {
        if (loan.status === "Devuelto") {
            if (isLoanReturnedLate(loan)) devueltoLate++;
            else devueltoOK++;
        } else if (loan.status === "Retirado") {
            if (getDaysOverdue(loan) > 0) retiradoLate++;
            else retiradoOK++;
        }
    });
    const complianceLabels = ["Devuelto a Tiempo", "Devuelto con Atraso", "Activo (A Tiempo)", "Activo (Atrasado)"];
    const complianceData = [devueltoOK, devueltoLate, retiradoOK, retiradoLate];
    const complianceColors = ["#10b981", "#f59e0b", "#3b82f6", "#ef4444"];
    renderDoughnutChart('chart-overdue-status', 'overdueStatus', complianceLabels, complianceData, complianceColors);
    
    // Tabla de Detalles de Alumnos con Mayor Demanda
    const groupedLoans = groupLoansById(appState.loans);
    const studentStats = {};
    
    groupedLoans.forEach(loan => {
        const key = loan.name;
        if (!studentStats[key]) {
            studentStats[key] = {
                name: loan.name,
                rut: loan.rut,
                total: 0,
                onTime: 0,
                late: 0,
                activeLate: 0
            };
        }
        
        studentStats[key].total++;
        
        if (loan.status === "Devuelto") {
            const finalOverdue = loan.daysOverdue || 0;
            if (finalOverdue > 0) {
                studentStats[key].late++;
            } else {
                studentStats[key].onTime++;
            }
        } else if (loan.status === "Retirado") {
            const daysOverdue = getDaysOverdue(loan);
            if (daysOverdue > 0) {
                studentStats[key].activeLate++;
                studentStats[key].late++; // También cuenta como entrega con retraso para las tasas de mora
            } else {
                studentStats[key].onTime++;
            }
        } else if (loan.status === "Solicitado") {
            studentStats[key].onTime++; // Se asume a tiempo temporalmente
        }
    });
    
    const sortedStudentsStats = Object.values(studentStats)
        .sort((a, b) => b.total - a.total)
        .slice(0, 5); // Tomamos el Top 5
        
    const tbody = document.getElementById('admin-top-students-table-body');
    if (tbody) {
        tbody.innerHTML = '';
        if (sortedStudentsStats.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:20px; color:var(--text-secondary);">No hay registros de alumnos.</td></tr>';
        } else {
            sortedStudentsStats.forEach(st => {
                const totalEvaluados = st.onTime + st.late;
                const tasaRetraso = totalEvaluados > 0 ? Math.round((st.late / totalEvaluados) * 100) : 0;
                
                let tasaColor = 'var(--success)';
                if (tasaRetraso > 50) tasaColor = 'var(--danger)';
                else if (tasaRetraso > 20) tasaColor = 'var(--warning)';
                
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td><strong>${st.name}</strong><br><small style="color:var(--text-secondary); font-size:0.72rem;">RUT: ${st.rut}</small></td>
                    <td><span class="loan-code-badge" style="background:rgba(79, 70, 229, 0.05); color:var(--primary); font-weight:700; font-size:0.75rem;">${st.total} pedidos</span></td>
                    <td><span style="color:var(--success); font-weight:600;">${st.onTime - st.activeLate}</span></td>
                    <td><span style="color:var(--danger); font-weight:600;">${st.late - st.activeLate}</span></td>
                    <td>
                        ${st.activeLate > 0 ? `
                            <span class="table-badge danger" style="background: rgba(239, 68, 68, 0.08); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.2); font-size: 0.65rem; font-weight:700;">
                                ${st.activeLate} activo(s)
                            </span>
                        ` : '<span style="color:var(--text-secondary); font-size:0.8rem;">-</span>'}
                    </td>
                    <td>
                        <span style="color:${tasaColor}; font-weight:700;">${tasaRetraso}%</span>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }
    }
}

function renderBarChart(canvasId, chartKey, labels, data, label, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    
    charts[chartKey] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                backgroundColor: color + '22',
                borderColor: color,
                borderWidth: 2,
                borderRadius: 4,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', stepSize: 1 } },
                x: { grid: { display: false }, ticks: { color: '#64748b' } }
            }
        }
    });
}

function renderLineChart(canvasId, chartKey, labels, data, label, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    
    charts[chartKey] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                borderColor: color,
                backgroundColor: color + '11',
                fill: true,
                tension: 0.35,
                borderWidth: 3,
                pointBackgroundColor: color,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', stepSize: 1 } },
                x: { grid: { display: false }, ticks: { color: '#64748b' } }
            }
        }
    });
}

function renderHorizontalBarChart(canvasId, chartKey, labels, data, label, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    
    charts[chartKey] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                backgroundColor: color + '22',
                borderColor: color,
                borderWidth: 2,
                borderRadius: 4,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                x: { grid: { color: '#f1f5f9' }, ticks: { color: '#64748b', stepSize: 1 } },
                y: { grid: { display: false }, ticks: { color: '#64748b' } }
            }
        }
    });
}

function renderAdminConfigLists() {
    renderAdminEquipmentList();
    renderAdminStudentsList();
}

function renderAdminEquipmentList() {
    dom.adminEquipmentList.innerHTML = '';
    if (appState.inventory.length === 0) {
        dom.adminEquipmentList.innerHTML = '<li class="table-empty"><p>No hay equipos registrados.</p></li>';
        return;
    }
    
    appState.inventory.forEach(item => {
        const li = document.createElement('li');
        li.className = 'admin-list-item';
        li.innerHTML = `
            <div class="admin-item-info">
                <h5>${item.name}</h5>
                <p>Categoría: ${item.category} | Stock Total: ${item.total} (Disp: ${item.available})</p>
            </div>
            <button class="btn-delete-item" data-id="${item.id}" data-name="${item.name}">
                <i data-lucide="trash-2"></i>
            </button>
        `;
        li.querySelector('.btn-delete-item').addEventListener('click', (e) => {
            const id = parseInt(e.currentTarget.getAttribute('data-id'));
            const name = e.currentTarget.getAttribute('data-name');
            deleteEquipmentConfig(id, name);
        });
        dom.adminEquipmentList.appendChild(li);
    });
    lucide.createIcons();
}

function renderAdminStudentsList() {
    dom.adminStudentsList.innerHTML = '';
    if (appState.students.length === 0) {
        dom.adminStudentsList.innerHTML = '<li class="table-empty"><p>No hay estudiantes registrados.</p></li>';
        return;
    }
    
    appState.students.forEach(student => {
        const li = document.createElement('li');
        li.className = 'admin-list-item';
        
        // Determinar estado UFRO (si observaciones/debt no está vacío)
        const isBlocked = student.status === 'Bloqueado' || (student.debt && student.debt.trim() !== "");
        const badgeColor = isBlocked ? 'color: var(--danger); font-weight:700;' : 'color: var(--success); font-weight:700;';
        const displayStatus = isBlocked ? 'Bloqueado' : 'Activo';
        
        li.innerHTML = `
            <div class="admin-item-info">
                <h5>${student.name} ${student.lastname || ""}</h5>
                <p>RUT: ${student.rut} | Fono: ${student.fono || '-'}</p>
                <small style="${badgeColor}">Estado: ${displayStatus} ${isBlocked ? `(${student.debt})` : ''}</small>
            </div>
            <button class="btn-delete-item" data-rut="${student.rut}" data-name="${student.name}">
                <i data-lucide="trash-2"></i>
            </button>
        `;
        li.querySelector('.btn-delete-item').addEventListener('click', (e) => {
            const rut = e.currentTarget.getAttribute('data-rut');
            const name = e.currentTarget.getAttribute('data-name');
            deleteStudentConfig(rut, name);
        });
        dom.adminStudentsList.appendChild(li);
    });
    lucide.createIcons();
}

async function saveEquipmentConfig() {
    const name = document.getElementById('eq-name').value.trim();
    const category = document.getElementById('eq-category').value;
    const total = parseInt(document.getElementById('eq-qty').value);
    const codesRaw = document.getElementById('eq-codes').value.trim();
    const codes = codesRaw.split(',').map(c => c.trim().toUpperCase()).filter(c => c.length > 0);
    const imageRaw = document.getElementById('eq-image') ? document.getElementById('eq-image').value.trim() : '';
    const image = getDirectImageUrl(imageRaw);
    const description = document.getElementById('eq-description') ? document.getElementById('eq-description').value.trim() : '';
    
    if (!name || codes.length === 0) {
        showToast("Completa los campos requeridos.", "warning");
        return;
    }
    
    const payload = { name, category, total, codes, image, description };
    
    if (CONFIG.demoMode) {
        const newId = appState.inventory.length > 0 ? Math.max(...appState.inventory.map(i => i.id)) + 1 : 1;
        appState.inventory.push({ id: newId, category, name, total, available: total, codes, image, description });
        saveDemoState();
        showToast(`Equipo "${name}" guardado`, "success");
        dom.formAddEquipment.reset();
        renderAdminConfigLists();
        renderLoansModule();
        updateAdminDashboard();
    } else {
        showToast("Agregando equipo...", "info");
        try {
            const response = await fetch(`${CONFIG.scriptUrl}?action=addEquipment`, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload)
            });
            const data = await response.json();
            if (data.status === "success") {
                showToast(`Equipo "${name}" agregado.`, "success");
                dom.formAddEquipment.reset();
                loadData();
            } else {
                showToast(data.message || "Error al agregar.", "danger");
            }
        } catch (e) {
            console.error(e);
            showToast("Error de red.", "danger");
        }
    }
}

async function saveStudentConfig() {
    const rawRut = document.getElementById('st-rut').value.trim();
    const nameFull = document.getElementById('st-name').value.trim();
    const email = document.getElementById('st-email').value.trim();
    const fono = document.getElementById('st-fono').value.trim();
    const status = dom.stStatusSelect.value;
    const debt = status === 'Bloqueado' ? document.getElementById('st-debt').value.trim() : '';
    
    if (!rawRut || !nameFull || !email) {
        showToast("Rellena los campos obligatorios.", "warning");
        return;
    }
    
    // Dividir nombre completo en Nombre y Apellido
    const nameParts = nameFull.split(' ');
    const name = nameParts[0];
    const lastname = nameParts.slice(1).join(' ') || '';
    
    const rut = formatRut(rawRut);
    const payload = { rut, name, lastname, fono, email, status, debt };
    
    if (CONFIG.demoMode) {
        const exists = appState.students.some(s => cleanRut(s.rut) === cleanRut(rut));
        if (exists) {
            showToast("RUT ya registrado", "warning");
            return;
        }
        
        appState.students.push(payload);
        saveDemoState();
        showToast(`Estudiante "${nameFull}" registrado`, "success");
        dom.formAddStudent.reset();
        dom.stDebtGroup.classList.add('hidden');
        renderAdminConfigLists();
        updateAdminDashboard();
    } else {
        showToast("Registrando estudiante...", "info");
        try {
            const response = await fetch(`${CONFIG.scriptUrl}?action=addStudent`, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify(payload)
            });
            const data = await response.json();
            if (data.status === "success") {
                showToast(`Estudiante registrado correctamente`, "success");
                dom.formAddStudent.reset();
                dom.stDebtGroup.classList.add('hidden');
                loadData();
            } else {
                showToast(data.message || "Error al registrar.", "danger");
            }
        } catch (e) {
            console.error(e);
            showToast("Error de red.", "danger");
        }
    }
}

async function deleteEquipmentConfig(id, name) {
    if (!confirm(`¿Eliminar equipo "${name}"?`)) return;
    if (CONFIG.demoMode) {
        appState.inventory = appState.inventory.filter(i => i.id !== id);
        saveDemoState();
        showToast(`Equipo "${name}" eliminado.`, "info");
        renderAdminConfigLists();
        renderLoansModule();
        updateAdminDashboard();
    } else {
        showToast("Eliminando equipo...", "info");
        try {
            const response = await fetch(`${CONFIG.scriptUrl}?action=deleteEquipment`, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ name })
            });
            const data = await response.json();
            if (data.status === "success") {
                showToast("Equipo eliminado con éxito.", "success");
                loadData();
            } else {
                showToast(data.message || "Error.", "danger");
            }
        } catch (e) {
            console.error(e);
            showToast("Error.", "danger");
        }
    }
}

async function deleteStudentConfig(rut, name) {
    if (!confirm(`¿Eliminar alumno "${name}"?`)) return;
    if (CONFIG.demoMode) {
        appState.students = appState.students.filter(s => s.rut !== rut);
        saveDemoState();
        showToast(`Estudiante "${name}" eliminado.`, "info");
        renderAdminConfigLists();
        updateAdminDashboard();
    } else {
        showToast("Eliminando estudiante...", "info");
        try {
            const response = await fetch(`${CONFIG.scriptUrl}?action=deleteStudent`, {
                method: 'POST',
                mode: 'cors',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ rut })
            });
            const data = await response.json();
            if (data.status === "success") {
                showToast("Estudiante eliminado con éxito.", "success");
                loadData();
            } else {
                showToast(data.message || "Error.", "danger");
            }
        } catch (e) {
            console.error(e);
            showToast("Error.", "danger");
        }
    }
}

// ==========================================
// INTEGRACIÓN DE CÁMARA (WEBCAM SCANNER)
// ==========================================

function startScanner(targetInputId) {
    dom.scannerModal.classList.remove('hidden');
    html5QrReader = new Html5Qrcode("reader");
    
    const config = {
        fps: 10,
        qrbox: (width, height) => {
            const minSize = Math.min(width, height);
            const boxSize = Math.floor(minSize * 0.7);
            return { width: boxSize, height: boxSize / 2 };
        }
    };
    
    html5QrReader.start(
        { facingMode: "environment" },
        config,
        (decodedText) => {
            document.getElementById(targetInputId).value = decodedText;
            showToast(`Código escaneado: ${decodedText}`, "success");
            stopScanner();
        },
        (errorMessage) => {}
    ).catch(err => {
        console.error(err);
        showToast("Error: No se pudo acceder a la cámara.", "danger");
        stopScanner();
    });
}

function stopScanner() {
    dom.scannerModal.classList.add('hidden');
    if (html5QrReader) {
        html5QrReader.stop().then(() => {
            html5QrReader = null;
        }).catch(err => {
            console.error(err);
        });
    }
}

// ==========================================
// UTILIDADES Y FORMATOS (RUT, FECHA, TOAST)
// ==========================================

function formatRut(rut) {
    let clean = cleanRut(rut);
    if (clean.length < 2) return clean;
    
    let body = clean.slice(0, -1);
    let dv = clean.slice(-1).toUpperCase();
    
    let formatted = "";
    while (body.length > 3) {
        formatted = "." + body.slice(-3) + formatted;
        body = body.slice(0, -3);
    }
    formatted = body + formatted + "-" + dv;
    return formatted;
}

function cleanRut(rut) {
    return rut.replace(/[^0-9kK]/g, '');
}

function getNowFormatted() {
    const d = new Date();
    const pad = (n) => n.toString().padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function showToast(message, type = "info") {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let iconName = 'info';
    if (type === 'success') iconName = 'check-circle';
    else if (type === 'danger') iconName = 'alert-octagon';
    else if (type === 'warning') iconName = 'alert-triangle';
    
    toast.innerHTML = `
        <i data-lucide="${iconName}"></i>
        <div class="toast-message">${message}</div>
    `;
    
    dom.toastContainer.appendChild(toast);
    lucide.createIcons();
    
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => {
            toast.remove();
        }, 300);
    }, 4000);
}

// ==========================================
// NUEVAS FUNCIONES DE ESTADÍSTICAS Y EXPORTACIÓN
// ==========================================

function isLoanReturnedLate(loan) {
    if (loan.status !== "Devuelto" || !loan.dateIn || !loan.progDevolucion) return false;
    
    // Parse progDevolucion (YYYY-MM-DD)
    const progParts = loan.progDevolucion.split(' ')[0].split('-');
    if (progParts.length < 3) return false;
    const progDate = new Date(parseInt(progParts[0], 10), parseInt(progParts[1], 10) - 1, parseInt(progParts[2], 10));
    
    // Parse dateIn (YYYY-MM-DD)
    const realParts = loan.dateIn.split(' ')[0].split('-');
    if (realParts.length < 3) return false;
    const realDate = new Date(parseInt(realParts[0], 10), parseInt(realParts[1], 10) - 1, parseInt(realParts[2], 10));
    
    return realDate > progDate;
}

function renderDoughnutChart(canvasId, chartKey, labels, data, colors) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (charts[chartKey]) charts[chartKey].destroy();
    
    charts[chartKey] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors,
                borderWidth: 2,
                borderColor: '#ffffff',
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: 'var(--text-secondary)',
                        font: { family: "'Plus Jakarta Sans', sans-serif", size: 10, weight: '600' }
                    }
                }
            }
        }
    });
}

function exportLoansToExcel() {
    if (appState.loans.length === 0) {
        showToast("No hay datos de préstamos para exportar.", "warning");
        return;
    }
    
    // Cabeceras en español
    const headers = [
        "ID Prestamo",
        "RUT Alumno",
        "Nombre Alumno",
        "Email Alumno",
        "Equipo",
        "Codigo Inventario",
        "Asignatura",
        "Fecha Registro/Solicitud",
        "Fecha Retiro Real",
        "Fecha Devolucion Real",
        "Estado",
        "Observaciones",
        "Dias Atraso (Activo)",
        "Devuelto con Atraso"
    ];
    
    const rows = appState.loans.map(loan => {
        const daysOverdue = getDaysOverdue(loan);
        const wasReturnedLate = isLoanReturnedLate(loan);
        
        return [
            loan.id || "",
            loan.rut || "",
            loan.name || "",
            loan.email || "",
            loan.item || "",
            loan.code || "",
            loan.subject || "",
            loan.dateOut || "",
            loan.status === 'Solicitado' ? "Pendiente" : (loan.dateDeliver || loan.dateOut || ""),
            loan.status === 'Devuelto' ? (loan.dateIn || "") : (loan.status === 'Retirado' ? 'En Transito' : '-'),
            loan.status || "",
            loan.obs || "",
            daysOverdue > 0 ? `${daysOverdue}` : "0",
            wasReturnedLate ? "SI" : "NO"
        ];
    });
    
    // Generar formato CSV usando delimitador punto y coma (;) compatible con Excel en Español
    const csvContent = [headers, ...rows].map(e => e.map(val => `"${val.toString().replace(/"/g, '""')}"`).join(";")).join("\n");
    
    // Incorporar BOM UTF-8 para compatibilidad de acentos y caracteres latinos
    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    link.setAttribute("download", `AVP_UFRO_Prestamos_${dateStr}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast("Datos exportados exitosamente. Abre el archivo CSV con Excel.", "success");
}
