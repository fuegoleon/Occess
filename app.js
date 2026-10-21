/**
 * Sistema de Gestión y Filtrado de Servicios (SAR)
 * Lógica interactiva con filtros de columna tipo Excel y Criterios SAR
 */

// 1. Definición completa de columnas disponibles
// 1. Columnas originales de la consulta SQL (Visibles por defecto)
const COLUMNS_DEF = [
  { key: 'ticket', label: 'Nro Ticket', defaultVisible: true, isCore: true },
  { key: 'primer_nombre', label: 'Nombres Afiliado', defaultVisible: true, isCore: true },
  { key: 'tiempo_atencion', label: 'Tiempo Atención', defaultVisible: true, isCore: true, isNumber: true },
  { key: 'usuario', label: 'Nombre Usuario Registro', defaultVisible: true, isCore: true },
  { key: 'correo_afiliado', label: 'Correo Afiliado', defaultVisible: true, isCore: true },
  { key: 'servicio', label: 'Servicio', defaultVisible: true, isCore: true },
  { key: 'depto_origen', label: 'Departamento Origen', defaultVisible: true, isCore: true },
  { key: 'telefono', label: 'Teléfono de Contacto', defaultVisible: true, isCore: true },
  
  // Columnas adicionales opcionales de la base (ocultas por defecto)
  { key: 'afiliado', label: 'Nombre Completo Afiliado', defaultVisible: false, isCore: false },
  { key: 'depto_usuario', label: 'Departamento Usuario Registro', defaultVisible: false, isCore: false },
  { key: 'agrupamiento', label: 'Agrupamiento de Servicio', defaultVisible: false, isCore: false },
  { key: 'regulador', label: 'Regulador', defaultVisible: false, isCore: false },
  { key: 'rechazado', label: '¿Rechazado?', defaultVisible: false, isCore: false },
  { key: 'equipo_contacto', label: '¿Equipo Contacto?', defaultVisible: false, isCore: false },
  { key: 'estado', label: 'Estado Servicio', defaultVisible: false, isCore: false },
  { key: 'tipo_servicio', label: 'Tipo Servicio', defaultVisible: false, isCore: false },
  { key: 'fecha_registro', label: 'Fecha Registro', defaultVisible: false, isCore: false },
  { key: 'hora_registro', label: 'Hora Registro', defaultVisible: false, isCore: false }
];

// Columnas dinámicas activas (se actualizan si se carga otro Excel)
let currentColumns = [...COLUMNS_DEF];

// Días de la semana para evaluación automática y simulación
const DIAS_SEMANA_KEYS = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];

// Obtener la clave del día para una fecha dada (por defecto HOY en tiempo real)
function getTodayDayKey(date = new Date()) {
  return DIAS_SEMANA_KEYS[date.getDay()] || 'lunes';
}

// Obtener la fecha correspondiente a un día de la semana dentro de la semana de referencia (por defecto semana actual de HOY)
function getCurrentWeekDayDate(dayKey, baseDate = new Date()) {
  const dayKeyToOffset = {
    'lunes': 1,
    'martes': 2,
    'miercoles': 3,
    'jueves': 4,
    'viernes': 5,
    'sabado': 6,
    'domingo': 7
  };
  const base = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
  // getDay(): 0 = Domingo, 1 = Lunes, ..., 6 = Sábado. Convertimos Domingo a 7.
  const currentDayNum = base.getDay() === 0 ? 7 : base.getDay();
  const targetDayNum = dayKeyToOffset[String(dayKey).toLowerCase()] || currentDayNum;
  const diffDays = targetDayNum - currentDayNum;
  const targetDate = new Date(base);
  targetDate.setDate(base.getDate() + diffDays);
  return targetDate;
}

// Estado global de la aplicación (SIN FILTROS PREVIOS APLICADOS)
const state = {
  rawData: [],
  filteredData: [],
  originalData: null,
  originalColumns: null,
  visibleColumns: new Set(COLUMNS_DEF.filter(c => c.defaultVisible).map(c => c.key)),
  columnFilters: {}, // { [colKey]: Set<string> de valores seleccionados }
  quickFilters: {},  // { [colKey]: string texto }
  globalSearch: '',
  onlySQLFilter: false, // Filtro exclusivo por condiciones WHERE de la consulta SQL SAR
  filterFechaAnterior: false, // Filtro inteligente por fecha anterior (Corte 9:00 AM)
  simulatedDay: getTodayDayKey(), // Día evaluado ('lunes', 'martes', etc.) - SIEMPRE el día de HOY por defecto
  sort: { key: 'ticket', direction: 'asc' },
  currentPage: 1,
  pageSize: 50,
  activePopupCol: null
};

// Elementos del DOM
const elements = {
  tableHeaderRow: document.getElementById('tableHeaderRow'),
  quickFilterRow: document.getElementById('quickFilterRow'),
  tableBody: document.getElementById('tableBody'),
  globalSearchInput: document.getElementById('globalSearchInput'),
  pageSizeSelect: document.getElementById('pageSizeSelect'),
  paginationInfo: document.getElementById('paginationInfo'),
  paginationControls: document.getElementById('paginationControls'),
  activeFiltersBar: document.getElementById('activeFiltersBar'),
  filterBadgesContainer: document.getElementById('filterBadgesContainer'),
  btnClearAllFilters: document.getElementById('btnClearAllFilters'),
  lastUpdatedBadge: document.getElementById('lastUpdatedBadge'),
  currentFileBadge: document.getElementById('currentFileBadge'),
  
  // Botón Filtro SQL SAR
  btnFiltroSQL: document.getElementById('btnFiltroSQL'),
  filtroSQLLabel: document.getElementById('filtroSQLLabel'),

  // Botón Filtro Fecha Anterior (Corte 9:00 AM)
  btnFiltroFechaAnterior: document.getElementById('btnFiltroFechaAnterior'),
  filtroFechaAnteriorLabel: document.getElementById('filtroFechaAnteriorLabel'),
  simuladorDiaBox: document.getElementById('simuladorDiaBox'),
  simuladorDiaSelect: document.getElementById('simuladorDiaSelect'),

  // Carga de Excel & Pantalla de Carga
  btnLoadExcel: document.getElementById('btnLoadExcel'),
  excelFileInput: document.getElementById('excelFileInput'),
  dropOverlay: document.getElementById('dropOverlay'),
  toastContainer: document.getElementById('toastContainer'),
  excelLoadingOverlay: document.getElementById('excelLoadingOverlay'),
  loadingIconWrap: document.getElementById('loadingIconWrap'),
  loadingIconMain: document.getElementById('loadingIconMain'),
  loadingTitle: document.getElementById('loadingTitle'),
  loadingFileBadge: document.getElementById('loadingFileBadge'),
  loadingFileName: document.getElementById('loadingFileName'),
  loadingFileSize: document.getElementById('loadingFileSize'),
  loadingStatus: document.getElementById('loadingStatus'),
  loadingPercent: document.getElementById('loadingPercent'),
  loadingProgressBar: document.getElementById('loadingProgressBar'),
  loadingSubstatus: document.getElementById('loadingSubstatus'),

  // KPIs
  kpiTotalTickets: document.getElementById('kpiTotalTickets'),
  kpiFilteredTickets: document.getElementById('kpiFilteredTickets'),
  kpiAvgTiempo: document.getElementById('kpiAvgTiempo'),
  kpiDeptosCount: document.getElementById('kpiDeptosCount'),

  // Botones y modales
  btnColumnSelector: document.getElementById('btnColumnSelector'),
  columnModal: document.getElementById('columnModal'),
  btnCloseColumnModal: document.getElementById('btnCloseColumnModal'),
  columnTogglesContainer: document.getElementById('columnTogglesContainer'),
  colSearchInput: document.getElementById('colSearchInput'),
  btnSelectAllCols: document.getElementById('btnSelectAllCols'),
  btnDeselectAllCols: document.getElementById('btnDeselectAllCols'),
  btnResetDefaultCols: document.getElementById('btnResetDefaultCols'),
  btnApplyCols: document.getElementById('btnApplyCols'),
  
  btnExportExcel: document.getElementById('btnExportExcel'),
  btnCopyClipboard: document.getElementById('btnCopyClipboard'),
  btnThemeToggle: document.getElementById('btnThemeToggle'),
  themeIcon: document.getElementById('themeIcon'),
  
  detailModal: document.getElementById('detailModal'),
  btnCloseDetailModal: document.getElementById('btnCloseDetailModal'),
  btnCloseDetailBtn: document.getElementById('btnCloseDetailBtn'),
  detailContent: document.getElementById('detailContent'),
  detailModalTitle: document.getElementById('detailModalTitle'),

  // Botón Limpiar Data
  btnClearData: document.getElementById('btnClearData')
};

// ==========================================
// IndexedDB para persistencia de datos entre recargas
// ==========================================
const IDB_NAME = 'SAR_DataStore';
const IDB_VERSION = 1;
const IDB_STORE = 'excelData';

function openIDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

function idbSave(key, value) {
  return openIDB().then(db => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = (e) => reject(e.target.error);
    });
  });
}

function idbLoad(key) {
  return openIDB().then(db => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = (e) => reject(e.target.error);
    });
  });
}

function idbClear() {
  return openIDB().then(db => {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = (e) => reject(e.target.error);
    });
  });
}

// Guardar estado completo en IndexedDB
function saveStateToIDB(fileName, rawData, columns) {
  const payload = {
    fileName: fileName,
    rawData: rawData,
    columns: columns,
    savedAt: new Date().toLocaleString('es-PE')
  };
  return idbSave('currentSession', payload).catch(err => {
    console.warn('No se pudo guardar en IndexedDB:', err);
  });
}

// ==========================================
// ESTADO VACÍO / BIENVENIDA
// ==========================================
function showEmptyState() {
  elements.tableHeaderRow.innerHTML = '';
  elements.quickFilterRow.innerHTML = '';
  elements.tableBody.innerHTML = `
    <tr>
      <td colspan="10" style="text-align: center; padding: 80px 40px;">
        <div style="display: flex; flex-direction: column; align-items: center; gap: 16px;">
          <div style="font-size: 64px; opacity: 0.6; animation: float 3s ease-in-out infinite;">📊</div>
          <h2 style="color: var(--text-main); margin: 0; font-size: 22px; font-weight: 700;">Sin datos cargados</h2>
          <p style="color: var(--text-muted); margin: 0; max-width: 420px; line-height: 1.6; font-size: 14px;">
            Cargue un archivo Excel (.xlsx, .xls, .csv) usando el botón 
            <strong style="color: var(--primary);">📂 Cargar Excel</strong> 
            o arrastre y suelte el archivo directamente en esta página.
          </p>
          <button onclick="document.getElementById('excelFileInput').click()" 
                  class="btn btn-primary" 
                  style="margin-top: 8px; padding: 10px 24px; font-size: 15px; border-radius: 10px; cursor: pointer;">
            📂 Seleccionar archivo Excel
          </button>
        </div>
      </td>
    </tr>
  `;
  elements.paginationInfo.textContent = '0 registros';
  elements.paginationControls.innerHTML = '';
  elements.kpiTotalTickets.textContent = '0';
  elements.kpiFilteredTickets.textContent = '0';
  elements.kpiAvgTiempo.textContent = '0 min';
  if (elements.kpiDeptosCount) elements.kpiDeptosCount.textContent = '0';
  elements.activeFiltersBar.style.display = 'none';
  if (elements.currentFileBadge) elements.currentFileBadge.textContent = 'Sin archivo cargado';
  if (elements.lastUpdatedBadge) elements.lastUpdatedBadge.textContent = 'Sin datos';
}

// Inicialización de la aplicación
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  setupEventListeners();
  setupExcelLoader();
  updateSimuladorDiaOptions();
  // Intentar restaurar datos de sesión anterior desde IndexedDB
  restoreFromIDB();
});

// Validación y extracción de teléfono desde Observación Inicial (comienza con 9 y 8 números a la derecha: 9XXXXXXXX)
function ensurePhoneExtracted(row) {
  const keys = Object.keys(row);
  let telKey = keys.find(k => k === 'telefono' || k === 'Teléfono de Contacto' || k === 'Telefono de Contacto')
            || keys.find(k => k.trim().toLowerCase().includes('tel') || k.trim().toLowerCase().includes('fono'));
  
  let obsKey = keys.find(k => k === 'observacion' || k === 'Observación Inicial' || k === 'Observacion Inicial')
            || keys.find(k => k.trim().toLowerCase().includes('observac') || k.trim().toLowerCase().includes('obs'));

  let currentTel = '';
  if (telKey && row[telKey] !== undefined) {
    const v = String(row[telKey] || '').trim();
    if (v && v !== '-' && v !== '0' && v.toLowerCase() !== 'sin teléfono' && v.toLowerCase() !== 'sin telefono') {
      currentTel = v;
    }
  }

  // Si no hay teléfono válido en el campo, buscar en Observación Inicial
  if (!currentTel && obsKey && row[obsKey]) {
    const obsText = String(row[obsKey]);
    // Regla: número que empieza con 9 y 8 números a la derecha (9XXXXXXXX)
    const match = obsText.match(/(?:(?<=\D)|^)(9\d{8})(?=\D|$)/);
    if (match) {
      const foundPhone = match[1];
      if (telKey) {
        row[telKey] = foundPhone;
      }
      row.telefono = foundPhone;
      if (row['Teléfono de Contacto'] !== undefined) row['Teléfono de Contacto'] = foundPhone;
      if (row['Telefono de Contacto'] !== undefined) row['Telefono de Contacto'] = foundPhone;
    }
  }
}

// ==========================================
// FUNCIONES EXACTAS DE LA CONSULTA SQL SAR:
// Encontrar_Correo, Encontrar_Nombre, Primera_Palabra
// ==========================================
function Encontrar_Correo(texto) {
  if (!texto) return "Sin Correo";
  const str = String(texto);
  const posOpen = str.indexOf("<");
  if (posOpen >= 0) {
    const posClose = str.indexOf(">", posOpen + 1);
    if (posClose > posOpen) {
      const extracted = str.substring(posOpen + 1, posClose).trim();
      return extracted.length > 0 ? extracted : "Sin Correo";
    }
  }
  return "Sin Correo";
}

function Encontrar_Nombre(texto) {
  if (!texto) return "SinNombre";
  const str = String(texto);
  const posOpen = str.indexOf("{");
  if (posOpen >= 0) {
    const posClose = str.indexOf("}", posOpen + 1);
    if (posClose > posOpen) {
      const extracted = str.substring(posOpen + 1, posClose).trim();
      return extracted.length > 0 ? extracted : "SinNombre";
    }
  }
  return "SinNombre";
}

function Primera_Palabra(texto) {
  if (!texto) return "";
  const upper = String(texto).trim().toUpperCase();
  const spaceIdx = upper.indexOf(" ");
  if (spaceIdx > 0) {
    return upper.substring(0, spaceIdx);
  }
  return upper;
}

function resolveNombreAfiliado(row) {
  const contacto = String(row['Nombre de persona de Contacto'] || row.contacto || '').trim();
  if (contacto.length > 0) {
    return Primera_Palabra(contacto);
  }
  const afiliado = String(row['Nombres Afiliado'] || row.afiliado || '').trim();
  if (afiliado.length > 0) {
    return Primera_Palabra(afiliado);
  }
  const obs = String(row['Observación Inicial'] || row.observacion || '');
  const nombreObs = Encontrar_Nombre(obs);
  return Primera_Palabra(nombreObs);
}

function resolveCorreoAfiliado(row) {
  const correoRaw = String(row['Correo de Contacto'] || row['Correo Afiliado'] || row.correo_afiliado || '').trim();
  if (correoRaw && correoRaw.toLowerCase() !== 'sin correo' && correoRaw !== '-') {
    return correoRaw;
  }
  const obs = String(row['Observación Inicial'] || row.observacion || '');
  return Encontrar_Correo(obs);
}

// ==========================================
// TABLAS DE REFERENCIA SQL MAESTRAS (Fallback autónomo)
// ==========================================
const TABLAS_REFERENCIA_DEFAULT = {
  reguladores: {
    "CARTA AL GERENTE ONCOCENTER": "CARTA AL GERENTE",
    "CARTA AL GERENTE ONCOSALUD": "CARTA AL GERENTE",
    "CARTA NOTARIAL": "CARTA NOTARIAL",
    "INDECOPI - CARTA DE BUENOS OFICIOS": "INDECOPI",
    "INDECOPI - DENUNCIAS": "INDECOPI",
    "INDECOPI - RECLAMOS SAC": "INDECOPI",
    "INDECOPI - REQUERIMIENTOS": "INDECOPI",
    "INDECOPI - RESOLUCIONES": "INDECOPI",
    "ONCOCENTER - CLINICA ONCOSALUD": "LIBRO ONCOCENTER",
    "ONCOCENTER - SEDE MIRAFLORES": "LIBRO ONCOCENTER",
    "ONCOCENTER - SEDE SAN BORJA": "LIBRO ONCOCENTER",
    "ONCOCENTER - SEDE SAN ISIDRO": "LIBRO ONCOCENTER",
    "ONCOCENTER - SEDE SURCO": "LIBRO ONCOCENTER",
    "ONCOCENTER - CLÍNICA DELGADO": "LIBRO ONCOCENTER",
    "ONCOSALUD - Arequipa": "LIBRO ONCOSALUD",
    "ONCOSALUD - Chiclayo": "LIBRO ONCOSALUD",
    "ONCOSALUD - Cuzco": "LIBRO ONCOSALUD",
    "ONCOSALUD - Ica": "LIBRO ONCOSALUD",
    "ONCOSALUD - Piura": "LIBRO ONCOSALUD",
    "ONCOSALUD - SEDE MIRAFLORES": "LIBRO ONCOSALUD",
    "ONCOSALUD - SEDE SAN BORJA": "LIBRO ONCOSALUD",
    "ONCOSALUD - SEDE SURCO": "LIBRO ONCOSALUD",
    "ONCOSALUD - Trujillo": "LIBRO ONCOSALUD",
    "ONCOSALUD - SEDE SAN ISIDRO": "LIBRO ONCOSALUD",
    "VIRTUAL - ONCOCENTER": "LIBRO VIRTUAL",
    "VIRTUAL - ONCOSALUD": "LIBRO VIRTUAL",
    "VIRTUAL ONCOSALUD - DECLINADO": "LIBRO VIRTUAL",
    "SUSALUD - CECONAR": "SUSALUD",
    "SUSALUD - CONSULTA DE AFILIACIONES": "SUSALUD",
    "SUSALUD - PEDIDOS DE INTERVENCION": "SUSALUD",
    "SUSALUD - QUEJAS": "SUSALUD",
    "SUSALUD - RECLAMOS": "SUSALUD",
    "SUSALUD - REQUERIMIENTOS": "SUSALUD",
    "SUSALUD - RESOLUCIONES": "SUSALUD",
    "Canales Alternos": "TRANSACCIONES DIARIAS",
    "Derechos Arco": "TRANSACCIONES DIARIAS",
    "No Recurrentes": "TRANSACCIONES DIARIAS",
    "CASOS CRÍTICOS": "TRANSACCIONES DIARIAS",
    "CENTROS DE CONCILIACIÓN EXTRAJUDICIAL - CONCILIACIONES": "TRANSACCIONES DIARIAS",
    "Fiscalía - Ministerio Público": "TRANSACCIONES DIARIAS",
    "CASOS LEGALES RELEVANTES": "TRANSACCIONES DIARIAS",
    "FACEBOOK": "TRANSACCIONES DIARIAS",
    "ONCOSALUD - CLINICA ONCOSALUD": "LIBRO ONCOSALUD",
    "Poder Judicial": "TRANSACCIONES DIARIAS",
    "Resagos Oncocenter": "TRANSACCIONES DIARIAS",
    "CANAL DIGITAL": "TRANSACCIONES DIARIAS",
    "RECLAMO DECLINADO": "TRANSACCIONES DIARIAS",
    "SUSALUD - DENUNCIA": "SUSALUD",
    "VIRTUAL - ONCOSALUD - ANGELLA": "LIBRO VIRTUAL"
  },
  serviciosRechazados: [
    "Solicitud de desafiliación",
    "DEVOLUCIONES POS",
    "DESPACHO BACK EN PRIMERA LÍNEA"
  ],
  tiemposAtencion: {
    "Gestión documentaria Back office CC": 5,
    "Modificaciones de programa Back Office CC": 10,
    "Adelanto chequeo preventivo": 15,
    "Agendamiento de desafiliación": 15,
    "Atención Oncocenter (Back Office)": 15,
    "CAMBIO DE CONDICION DE FUMADOR": 15,
    "CAMBIO DE TITULAR / APORTANTE": 15,
    "Chequeo Preventivo Citas": 15,
    "CONDONACIÓN DE DEUDAS": 15,
    "CONSTANCIAS DE AFILIACION Y DESAFILIACION": 15,
    "CONSTANCIAS DE AFILIACION Y DESAFILIACION AUNA SALUD": 15,
    "Desafiliación Directa": 15,
    "Desbloqueo contacto clientes": 15,
    "DESPACHO BACK OFFICE": 15,
    "Despacho Boletas / Facturas": 15,
    "Despacho Cartas Recordatorias": 15,
    "Despacho Kit de Afiliación": 15,
    "DETRACCIÓN": 15,
    "DEVOLUCION  POST DESAFILIACIÓN": 15,
    "DEVOLUCION ERRORES DE COBRO": 15,
    "Devolución ERRORES de COBRO": 15,
    "DEVOLUCION POR ERROR EN PROCESO DE AFILIACION": 15,
    "DEVOLUCIONES POR CAMBIOS EN EL PROGRAMA": 15,
    "DEVOLUCIONES POR SOLICITUD DE COBRANZAS": 15,
    "ENVIO DOCUMENTO DEL AFILIADO PARA GESTIÓN": 15,
    "ESTADO DE DESAFILIACION": 15,
    "FACILIDADES DE PAGO": 15,
    "FRACCIONAMIENTO Y CONDONACIONES DE CUOTA": 15,
    "GESTIÓN BACK PLATAFORMA PRESENCIAL": 15,
    "INFORMACIÓN DE COBERTURA Y BENEFICIOS": 15,
    "INFORMACIÓN DE PROMOCIONES ESPECIALES": 15,
    "MIGRACION DE CANALES DE VENTA": 15,
    "MIGRACION DE PROGRAMAS AUNA SALUD": 15,
    "MIGRACION DE PROGRAMAS ONCOLOGICOS": 15,
    "MIGRACIÓN GRUPAL CONVENIO EMPRESAS Y ALIANZAS": 15,
    "Modificación de fecha de Pago": 15,
    "Modificación de FORMA de pago": 15,
    "REACTIVACION DE SEGURO": 15,
    "REGULARIZACION NO RECURRENTE": 15,
    "REGULARIZACION RECURRENTE": 15,
    "REPLICA - POR MAL ANÁLISIS": 15,
    "REPLICA - SOLICITUD ADICIONAL": 15,
    "REPLICA - SOLICITUD DE DOCUMENTO AL AFILIADO": 15,
    "REPLICA -INSATISFACCIÓN CON LA RESPUESTA BRINDADA": 15,
    "REPLICA POR RECONSIDERACIÓN O REEVALUACIÓN DEL CASO": 15,
    "RETIRO DE BASE DE CONTACTO": 15,
    "Separación y unificación de Programas": 15,
    "Solicitud de Cambio de Asesora": 15,
    "Solicitud de Canje Boleta - Factura": 15,
    "SOLICITUD DE CONTRATO DE AFILIACIÓN": 15,
    "Solicitud de Correción  Boleta - Factura": 15,
    "Solicitud de Estados de Cuenta": 15,
    "SOLICITUD ESTADO DE CUOTAS": 15,
    "Validación de Beneficios del afiliado": 15,
    "Validación de Cobertura Oncosalud": 15,
    "VALIDACIÓN DE INCREMENTO DE TARIFA": 15
  }
};

function getTablasReferencia() {
  return window.TABLAS_REFERENCIA || TABLAS_REFERENCIA_DEFAULT;
}


// Precalcula una cadena de búsqueda normalizada en minúsculas para cada fila
// Esto acelera la búsqueda global hasta 50 veces al evitar iterar 60 columnas por fila
function buildSearchStr(row) {
  if (!row) return '';
  const parts = [];
  for (const k in row) {
    if (k.charCodeAt(0) === 95) continue; // Ignorar claves internas con prefijo '_'
    const v = row[k];
    if (v !== undefined && v !== null && v !== '') {
      parts.push(String(v).toLowerCase());
    }
  }
  return parts.join(' ');
}

// Cargar datos desde data.js o desde datos restaurados (SIN NINGÚN FILTRO Y MOSTRANDO TODAS LAS COLUMNAS)
function loadData(restoredPayload) {
  let dataSource = null;
  let columnsSource = null;
  let metaSource = null;
  let fileNameSource = 'Sin archivo cargado';

  if (restoredPayload) {
    // Datos restaurados desde IndexedDB
    dataSource = restoredPayload.rawData;
    columnsSource = restoredPayload.columns;
    fileNameSource = restoredPayload.fileName || 'Archivo restaurado';
    metaSource = { fecha_actualizacion: restoredPayload.savedAt, archivo_base: fileNameSource };
  } else if (window.DATOS_SISTEMA && Array.isArray(window.DATOS_SISTEMA)) {
    // Datos desde data.js (modo compatible)
    dataSource = window.DATOS_SISTEMA;
    columnsSource = window.COLUMNAS_SISTEMA;
    metaSource = window.METADATOS_SISTEMA;
  }

  if (!dataSource || !Array.isArray(dataSource) || dataSource.length === 0) {
    showEmptyState();
    return;
  }

  state.rawData = [...dataSource];
  state.rawData.forEach(row => {
    ensurePhoneExtracted(row);
    row.primer_nombre = resolveNombreAfiliado(row);
    row.correo_afiliado = resolveCorreoAfiliado(row);
    row._origFechaRegistro = row._origFechaRegistro || row['Fecha Registro'] || row.fecha_registro || '';
    row._origMesRegistro = row._origMesRegistro || row['Mes Registro'] || row.mes_registro || '';
    row._origHoraRegistro = row._origHoraRegistro || row['Hora Registro'] || row.hora_registro || '';
    row._origDateYMD = getRowDateYMD(row);
    row._searchStr = buildSearchStr(row);
  });

  // Determinar TODAS las columnas del archivo original
  let headers = [];
  if (columnsSource && Array.isArray(columnsSource) && columnsSource.length > 0) {
    headers = columnsSource.map(h => typeof h === 'string' ? h.replace('Tiempo Atenci\u00c3\u00b3n', 'Tiempo Atenci\u00f3n') : (h.label || h.key));
  } else if (state.rawData.length > 0) {
    const metaKeys = new Set(['ticket', 'primer_nombre', 'tiempo_atencion', 'correo_afiliado', 'telefono', 'cumple_sar', 'es_merino_sin_int', 'criterio_motivo']);
    headers = Object.keys(state.rawData[0]).filter(k => !metaKeys.has(k));
  }

  // Configurar definiciones para TODAS las columnas
  currentColumns = headers.map(h => {
    const isNum = /tiempo|duraci|monto|cuota/i.test(h);
    return {
      key: h,
      label: h,
      defaultVisible: true,
      isCore: isCoreSQLCol(h, h),
      isNumber: isNum
    };
  });

  state.visibleColumns = new Set(currentColumns.map(c => c.key));

  if (!state.originalData) {
    state.originalData = [...state.rawData];
    state.originalColumns = [...currentColumns];
  }
  if (metaSource) {
    if (metaSource.fecha_actualizacion && elements.lastUpdatedBadge) {
      elements.lastUpdatedBadge.textContent = 'Actualizado: ' + metaSource.fecha_actualizacion;
    }
    if (metaSource.archivo_base && elements.currentFileBadge) {
      elements.currentFileBadge.textContent = metaSource.archivo_base;
    }
  }

  // Asegurar que todos los filtros estén limpios por defecto
  state.columnFilters = {};
  state.quickFilters = {};
  state.globalSearch = '';
  state.currentPage = 1;

  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

// Persistencia de filtros y estado de UI en localStorage
function saveUIState() {
  if (!state.rawData || state.rawData.length === 0) return;
  try {
    const ui = {
      onlySQLFilter: state.onlySQLFilter,
      globalSearch: state.globalSearch,
      quickFilters: state.quickFilters,
      currentPage: state.currentPage,
      columnOrder: currentColumns.map(c => c.key),
      columnFilters: Object.fromEntries(
        Object.entries(state.columnFilters).map(([k, set]) => [k, Array.from(set)])
      )
    };
    localStorage.setItem('SAR_UIState', JSON.stringify(ui));
  } catch (e) {
    console.warn('No se pudo guardar SAR_UIState:', e);
  }
}

function restoreUIState() {
  try {
    const raw = localStorage.getItem('SAR_UIState');
    if (!raw) return;
    const ui = JSON.parse(raw);
    if (!ui) return;

    if (Array.isArray(ui.columnOrder) && ui.columnOrder.length > 0) {
      const orderMap = new Map();
      ui.columnOrder.forEach((key, idx) => orderMap.set(key, idx));
      currentColumns.sort((a, b) => {
        const idxA = orderMap.has(a.key) ? orderMap.get(a.key) : 9999;
        const idxB = orderMap.has(b.key) ? orderMap.get(b.key) : 9999;
        return idxA - idxB;
      });
    }

    if (typeof ui.onlySQLFilter === 'boolean') {
      state.onlySQLFilter = ui.onlySQLFilter;
      if (elements.btnFiltroSQL) {
        if (state.onlySQLFilter) {
          elements.btnFiltroSQL.classList.remove('btn-outline');
          elements.btnFiltroSQL.classList.add('btn-primary');
          elements.filtroSQLLabel.textContent = '✓ Filtro SQL Activo (8 Columnas)';
        } else {
          elements.btnFiltroSQL.classList.remove('btn-primary');
          elements.btnFiltroSQL.classList.add('btn-outline');
          elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';
        }
      }
    }
    if (ui.globalSearch) {
      state.globalSearch = ui.globalSearch;
      if (elements.globalSearchInput) elements.globalSearchInput.value = ui.globalSearch;
    }
    if (ui.quickFilters && typeof ui.quickFilters === 'object') {
      state.quickFilters = ui.quickFilters;
    }
    if (ui.columnFilters && typeof ui.columnFilters === 'object') {
      state.columnFilters = {};
      Object.entries(ui.columnFilters).forEach(([k, arr]) => {
        if (Array.isArray(arr) && arr.length > 0) {
          state.columnFilters[k] = new Set(arr);
        }
      });
    }
    if (ui.currentPage) {
      state.currentPage = ui.currentPage;
    }
  } catch (e) {
    console.warn('Error al restaurar SAR_UIState:', e);
  }
}

// Restaurar datos desde IndexedDB al recargar la página
function restoreFromIDB() {
  idbLoad('currentSession').then(payload => {
    if (payload && payload.rawData && payload.rawData.length > 0) {
      loadData(payload);
      restoreUIState();
      renderHeaderAndQuickFilters();
      applyFiltersAndSort();
      showToast(`✅ Datos restaurados: ${payload.rawData.length.toLocaleString()} registros de "${payload.fileName}"`, 'success');
    } else {
      loadData();
    }
  }).catch(err => {
    console.warn('Error al restaurar desde IndexedDB:', err);
    loadData();
  });
}

// Limpiar toda la data y volver al estado vacío
function clearAllData() {
  runWithLoading({
    title: 'Limpiando Base de Datos',
    iconType: 'clear',
    badgeText: 'Eliminación completa',
    badgeIcon: '🧹',
    initialStatus: 'Eliminando registros de memoria...',
    initialSubstatus: 'Limpiando almacenamiento local IndexedDB...',
    successStatus: '¡Base de Datos Limpia!',
    successSubstatus: 'Todos los registros han sido removidos',
    holdSuccessMs: 350
  }, () => {
    state.rawData = [];
    state.filteredData = [];
    state.originalData = null;
    state.originalColumns = null;
    state.columnFilters = {};
    state.quickFilters = {};
    state.globalSearch = '';
    state.onlySQLFilter = false;
    state.filterFechaAnterior = false;
    state.simulatedDay = getTodayDayKey();
    state.currentPage = 1;
    currentColumns = [...COLUMNS_DEF];
    state.visibleColumns = new Set(COLUMNS_DEF.filter(c => c.defaultVisible).map(c => c.key));

    if (elements.globalSearchInput) elements.globalSearchInput.value = '';
    if (elements.btnFiltroSQL) {
      elements.btnFiltroSQL.classList.remove('btn-primary');
      elements.btnFiltroSQL.classList.add('btn-outline');
      elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';
    }
    if (elements.btnFiltroFechaAnterior) {
      elements.btnFiltroFechaAnterior.classList.remove('btn-fecha-activa');
      elements.btnFiltroFechaAnterior.classList.add('btn-outline');
      elements.filtroFechaAnteriorLabel.textContent = '📅 Filtro Fecha Anterior (9:00 AM)';
    }
    if (elements.simuladorDiaBox) {
      elements.simuladorDiaBox.style.display = 'none';
    }
    if (elements.simuladorDiaSelect) {
      updateSimuladorDiaOptions();
    }

    // Borrar IndexedDB y localStorage
    idbClear().catch(err => console.warn('Error al limpiar IndexedDB:', err));
    try {
      localStorage.removeItem('SAR_UIState');
    } catch (e) {}

    showEmptyState();
  }).then(() => {
    showToast('🧹 Todos los datos han sido eliminados', 'info');
  });
}

// Configurar eventos
function setupEventListeners() {
  // Toggle Filtro Consulta SQL SAR (muestra solo las 8 columnas especificadas)
  if (elements.btnFiltroSQL) {
    elements.btnFiltroSQL.addEventListener('click', () => {
      closeAllExcelPopups();
      const willActivate = !state.onlySQLFilter;

      if (willActivate) {
        const isChained = state.filterFechaAnterior;
        const prevCount = state.filteredData.length;

        runWithLoading({
          title: isChained ? 'Aplicando Filtro SQL sobre Fecha Anterior' : 'Aplicando Filtro SQL (SAR)',
          iconType: 'sql',
          badgeText: isChained ? `Filtro Encadenado (${prevCount.toLocaleString()} reg. base)` : '8 Columnas Principales',
          badgeIcon: '⚡',
          initialStatus: isChained 
            ? `Filtrando los ${prevCount.toLocaleString()} registros de Fecha Anterior con reglas SQL SAR...`
            : 'Evaluando condiciones WHERE y SLA...',
          initialSubstatus: 'Filtrando por Call Center y horarios de corte...',
          successStatus: isChained ? '¡Filtro SQL Encadenado Aplicado!' : '¡Filtro SQL SAR Activado!',
          successSubstatus: () => isChained 
            ? `${state.filteredData.length.toLocaleString()} registros cumplen ambas condiciones (8 columnas)`
            : 'Mostrando únicamente las 8 columnas principales',
          holdSuccessMs: 400
        }, () => {
          state.onlySQLFilter = true;
          elements.btnFiltroSQL.classList.remove('btn-outline');
          elements.btnFiltroSQL.classList.add('btn-primary');
          elements.filtroSQLLabel.textContent = isChained ? '✓ Filtro SQL (Encadenado)' : '✓ Filtro SQL Activo (8 Col)';

          state.previousVisibleColumns = new Set(state.visibleColumns);
          const sqlCols = getSQLTargetColumns();
          state.visibleColumns = new Set(sqlCols.map(c => c.key));

          state.currentPage = 1;
          renderHeaderAndQuickFilters();
          applyFiltersAndSort();
        }).then(() => {
          if (isChained) {
            showToast(`⚡🔗 Filtro SQL aplicado sobre Fecha Anterior: de ${prevCount.toLocaleString()} registros quedaron ${state.filteredData.length.toLocaleString()} (8 columnas principales)`, 'success');
          } else {
            showToast('Filtro SQL SAR activado: mostrando solo 8 columnas principales', 'info');
          }
        });
      } else {
        runWithLoading({
          title: 'Restaurando Vista Completa',
          iconType: 'columns',
          badgeText: 'Todas las Columnas',
          badgeIcon: '📊',
          initialStatus: 'Restaurando configuración de columnas...',
          initialSubstatus: 'Reactivando columnas visibles previas...',
          successStatus: '¡Vista Completa Restaurada!',
          successSubstatus: 'Mostrando todas las columnas del archivo',
          holdSuccessMs: 350
        }, () => {
          state.onlySQLFilter = false;
          elements.btnFiltroSQL.classList.remove('btn-primary');
          elements.btnFiltroSQL.classList.add('btn-outline');
          elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';

          if (state.previousVisibleColumns && state.previousVisibleColumns.size > 0) {
            state.visibleColumns = new Set(state.previousVisibleColumns);
          } else {
            state.visibleColumns = new Set(currentColumns.map(c => c.key));
          }
          state.previousVisibleColumns = null;

          if (state.filterFechaAnterior && elements.filtroFechaAnteriorLabel) {
            elements.filtroFechaAnteriorLabel.textContent = '✓ Fecha Anterior Activo (9:00 AM)';
          }

          state.currentPage = 1;
          renderHeaderAndQuickFilters();
          applyFiltersAndSort();
        }).then(() => {
          showToast('Filtro SQL desactivado: mostrando todas las columnas', 'info');
        });
      }
    });
  }

  // Toggle Filtro Inteligente Fecha Anterior (Corte 9:00 AM)
  if (elements.btnFiltroFechaAnterior) {
    elements.btnFiltroFechaAnterior.addEventListener('click', () => {
      closeAllExcelPopups();
      const willActivate = !state.filterFechaAnterior;

      if (willActivate) {
        // Asegurar que el día evaluado sea siempre el día de HOY por defecto
        state.simulatedDay = getTodayDayKey();
        updateSimuladorDiaOptions();
        const cfg = getTargetDatesConfig(state.simulatedDay);
        const isChained = state.onlySQLFilter;
        const prevCount = state.filteredData.length;

        const badgeDay = cfg.isToday ? `Hoy ${cfg.dayName}` : `Día: ${cfg.dayName}`;
        const targetDesc = cfg.isWeekend ? '0 días' : (cfg.dayOfWeek === 1 ? '3 días antes: Vie, Sáb, Dom' : `Ayer: ${cfg.targetDates[0]}`);

        runWithLoading({
          title: isChained ? 'Aplicando Fecha Anterior sobre Filtro SQL' : 'Aplicando Filtro de Fecha Anterior (9:00 AM)',
          iconType: 'calendar',
          badgeText: `${badgeDay} (${targetDesc})`,
          badgeIcon: '📅',
          initialStatus: isChained 
            ? `Aplicando corte de fecha anterior a los ${prevCount.toLocaleString()} registros de Filtro SQL...`
            : 'Evaluando reglas de corte SAR a las 9:00 AM...',
          initialSubstatus: 'Identificando rango de fechas previas y actualizando Fecha Registro...',
          successStatus: isChained ? '¡Filtros Encadenados Aplicados!' : '¡Filtro de Fecha Anterior Aplicado!',
          successSubstatus: () => `${state.filteredData.length.toLocaleString()} registros resultantes (Hoy 09:00:00)`,
          holdSuccessMs: 420
        }, () => {
          state.filterFechaAnterior = true;
          elements.btnFiltroFechaAnterior.classList.remove('btn-outline');
          elements.btnFiltroFechaAnterior.classList.add('btn-fecha-activa');
          elements.filtroFechaAnteriorLabel.textContent = isChained ? '✓ Fecha Anterior (Encadenado)' : '✓ Fecha Anterior Activo (9:00 AM)';
          if (isChained && elements.filtroSQLLabel) {
            elements.filtroSQLLabel.textContent = '✓ Filtro SQL (Encadenado)';
          }
          if (elements.simuladorDiaBox) elements.simuladorDiaBox.style.display = 'flex';

          const res = applyFechaAnteriorTransformation();
          state.currentPage = 1;
          renderHeaderAndQuickFilters();
          applyFiltersAndSort();
          return res;
        }).then((res) => {
          if (isChained) {
            showToast(`📅🔗 Filtro Fecha Anterior aplicado junto con Filtro SQL: ${state.filteredData.length.toLocaleString()} registros resultantes`, 'success');
          } else if (res) {
            showToast(`✅ ${res.cfg.summary}. ${res.matchCount.toLocaleString()} registros actualizados con Fecha Registro: ${res.hoyStr}, Mes: ${res.mesActual}, Hora: ${res.horaActual}`, 'success');
          }
        });
      } else {
        runWithLoading({
          title: 'Restaurando Fechas Originales',
          iconType: 'reset',
          badgeText: 'Restaurando Datos',
          badgeIcon: '🔄',
          initialStatus: 'Restableciendo valores de fecha y hora originales...',
          initialSubstatus: 'Deshaciendo corte de 9:00 AM...',
          successStatus: '¡Valores Originales Restaurados!',
          successSubstatus: 'Mostrando todos los registros sin filtro de fecha',
          holdSuccessMs: 350
        }, () => {
          state.filterFechaAnterior = false;
          elements.btnFiltroFechaAnterior.classList.remove('btn-fecha-activa');
          elements.btnFiltroFechaAnterior.classList.add('btn-outline');
          elements.filtroFechaAnteriorLabel.textContent = '📅 Filtro Fecha Anterior (9:00 AM)';
          if (elements.simuladorDiaBox) elements.simuladorDiaBox.style.display = 'none';

          if (state.onlySQLFilter && elements.filtroSQLLabel) {
            elements.filtroSQLLabel.textContent = '✓ Filtro SQL Activo (8 Col)';
          }

          restoreOriginalFechaValues();
          state.currentPage = 1;
          renderHeaderAndQuickFilters();
          applyFiltersAndSort();
        }).then(() => {
          showToast('Filtro de fecha anterior desactivado: valores originales restaurados', 'info');
        });
      }
    });
  }

  // Cambio de simulación de día de la semana
  if (elements.simuladorDiaSelect) {
    elements.simuladorDiaSelect.addEventListener('change', (e) => {
      const selectedDay = e.target.value;
      state.simulatedDay = selectedDay;
      const cfg = getTargetDatesConfig(selectedDay);
      const dayLabel = `${cfg.dayName}${cfg.isToday ? ' (Hoy)' : ''}`;
      const targetDesc = cfg.isWeekend ? '0 días' : (cfg.dayOfWeek === 1 ? '3 días antes' : `Ayer: ${cfg.targetDates[0]}`);

      if (state.filterFechaAnterior) {
        runWithLoading({
          title: `Evaluando Día: ${cfg.dayName}${cfg.isToday ? ' (Hoy)' : ''}`,
          iconType: 'calendar',
          badgeText: `${dayLabel} (${targetDesc})`,
          badgeIcon: '📅',
          initialStatus: 'Recalculando reglas de corte por día...',
          initialSubstatus: 'Filtrando registros de fecha anterior correspondiente...',
          successStatus: '¡Corte de Fecha Aplicado!',
          successSubstatus: (res) => `${(res?.matchCount || 0).toLocaleString()} registros encontrados`,
          holdSuccessMs: 400
        }, () => {
          const res = applyFechaAnteriorTransformation();
          state.currentPage = 1;
          applyFiltersAndSort();
          return res;
        }).then((res) => {
          if (res) {
            showToast(`Día evaluado: ${res.cfg.summary}. Registros: ${res.matchCount.toLocaleString()}`, res.cfg.isWeekend ? 'warning' : 'info');
          }
        });
      }
    });
  }

  // Búsqueda global con Debounce (250ms) para que el tipeo sea instantáneo y fluido
  let globalSearchDebounceTimer = null;
  elements.globalSearchInput.addEventListener('input', (e) => {
    const term = e.target.value.trim().toLowerCase();
    clearTimeout(globalSearchDebounceTimer);
    globalSearchDebounceTimer = setTimeout(() => {
      state.globalSearch = term;
      state.currentPage = 1;
      applyFiltersAndSort();
    }, 250);
  });

  // Ejecución inmediata al presionar Enter en el buscador
  elements.globalSearchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      clearTimeout(globalSearchDebounceTimer);
      state.globalSearch = e.target.value.trim().toLowerCase();
      state.currentPage = 1;
      applyFiltersAndSort();
    }
  });

  // Selector de tamaño de página
  elements.pageSizeSelect.addEventListener('change', (e) => {
    state.pageSize = parseInt(e.target.value, 10);
    state.currentPage = 1;
    renderTable();
  });

  // Limpiar todos los filtros
  elements.btnClearAllFilters.addEventListener('click', () => {
    runWithLoading({
      title: 'Restableciendo Todos los Filtros',
      iconType: 'reset',
      badgeText: 'Filtros Activos',
      badgeIcon: '🔄',
      initialStatus: 'Limpiando filtros de columnas y búsquedas...',
      initialSubstatus: 'Restaurando tabla a estado original...',
      successStatus: '¡Filtros Restablecidos!',
      successSubstatus: `${state.rawData.length.toLocaleString()} registros visibles`,
      holdSuccessMs: 350
    }, () => {
      state.columnFilters = {};
      state.quickFilters = {};
      state.globalSearch = '';
      state.onlySQLFilter = false;
      elements.globalSearchInput.value = '';

      // Reset botón Filtro SQL
      if (elements.btnFiltroSQL) {
        elements.btnFiltroSQL.classList.remove('btn-primary');
        elements.btnFiltroSQL.classList.add('btn-outline');
        elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';
      }

      // Restaurar columnas
      if (state.previousVisibleColumns && state.previousVisibleColumns.size > 0) {
        state.visibleColumns = new Set(state.previousVisibleColumns);
      } else {
        state.visibleColumns = new Set(currentColumns.map(c => c.key));
      }
      state.previousVisibleColumns = null;

      // Limpiar inputs rápidos
      document.querySelectorAll('.col-quick-search').forEach(input => input.value = '');
      
      state.currentPage = 1;
      renderHeaderAndQuickFilters();
      applyFiltersAndSort();
    });
  });

  // Modal de Columnas
  elements.btnColumnSelector.addEventListener('click', openColumnModal);
  elements.btnCloseColumnModal.addEventListener('click', () => elements.columnModal.classList.remove('open'));
  if (elements.btnSelectAllCols) {
    elements.btnSelectAllCols.addEventListener('click', () => {
      document.querySelectorAll('.col-toggle-checkbox').forEach(cb => cb.checked = true);
    });
  }
  if (elements.btnDeselectAllCols) {
    elements.btnDeselectAllCols.addEventListener('click', () => {
      document.querySelectorAll('.col-toggle-checkbox').forEach(cb => cb.checked = false);
    });
  }
  if (elements.btnResetDefaultCols) {
    elements.btnResetDefaultCols.addEventListener('click', () => {
      // 8 columnas principales solicitadas de la consulta SQL SAR
      const sqlCols = getSQLTargetColumns();
      const targetKeys = new Set(sqlCols.map(c => c.key));
      document.querySelectorAll('.col-toggle-checkbox').forEach(cb => {
        const colKey = cb.dataset.col;
        const colDef = currentColumns.find(c => c.key === colKey);
        const isCore = targetKeys.has(colKey) || (colDef && isCoreSQLCol(colDef.key, colDef.label));
        cb.checked = isCore;
      });
    });
  }
  if (elements.colSearchInput) {
    elements.colSearchInput.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      document.querySelectorAll('.col-checkbox-card').forEach(card => {
        const text = card.textContent.toLowerCase();
        card.style.display = text.includes(q) ? 'flex' : 'none';
      });
    });
  }
  elements.btnApplyCols.addEventListener('click', applyColumnVisibility);

  // Cerrar modales con fondo
  elements.columnModal.addEventListener('click', (e) => {
    if (e.target === elements.columnModal) elements.columnModal.classList.remove('open');
  });
  elements.detailModal.addEventListener('click', (e) => {
    if (e.target === elements.detailModal) elements.detailModal.classList.remove('open');
  });
  elements.btnCloseDetailModal.addEventListener('click', () => elements.detailModal.classList.remove('open'));
  elements.btnCloseDetailBtn.addEventListener('click', () => elements.detailModal.classList.remove('open'));

  // Exportar y copiar
  elements.btnExportExcel.addEventListener('click', exportToExcelCSV);
  elements.btnCopyClipboard.addEventListener('click', copyToClipboard);

  // Cerrar popups de filtro al hacer clic fuera
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.excel-popup') && !e.target.closest('.excel-filter-btn')) {
      closeAllExcelPopups();
    }
  });

  // Tema
  elements.btnThemeToggle.addEventListener('click', toggleTheme);

  // Botón Limpiar Data
  if (elements.btnClearData) {
    elements.btnClearData.addEventListener('click', () => {
      if (state.rawData.length === 0) {
        showToast('No hay datos cargados para limpiar', 'info');
        return;
      }
      clearAllData();
    });
  }
}

// Helper para determinar si una columna pertenece a las 8 principales de la consulta SQL SAR
function isCoreSQLCol(key, label) {
  const k = String(key || '').trim().toLowerCase();
  const l = String(label || '').trim().toLowerCase();
  const text = `${k} ${l}`;

  // 1. Nro Ticket
  if (k === 'ticket' || k === 'nro ticket' || k === 'nro. ticket' || l === 'nro ticket' || l === 'nro. ticket') return true;
  if (text.includes('ticket') && !text.includes('origen') && !text.includes('atenci')) return true;

  // 2. Nombres Afiliado
  if (k === 'primer_nombre' || k === 'nombres afiliado' || l === 'nombres afiliado') return true;
  if (text.includes('afiliado') && !text.includes('cod') && !text.includes('programa') && !text.includes('grupo') && !text.includes('completo')) return true;

  // 3. Tiempo Atención
  if (k === 'tiempo_atencion' || k === 'tiempo atención' || text.includes('tiempo atenci')) return true;

  // 4. Nombre Usuario Registro
  if (k === 'usuario' || k === 'nombre usuario registro' || l === 'nombre usuario registro') return true;
  if (text.includes('usuario') && text.includes('registro') && !text.includes('depto') && !text.includes('departamento')) return true;

  // 5. Correo Afiliado / Correo de Contacto
  if (k === 'correo_afiliado' || k === 'correo afiliado' || k === 'correo de contacto' || l === 'correo afiliado' || l === 'correo de contacto') return true;
  if (text.includes('correo') && !text.includes('fecha')) return true;

  // 6. Servicio (exacto)
  if (k === 'servicio' || l === 'servicio') return true;

  // 7. Departamento Origen
  if (k === 'depto_origen' || k === 'departamento origen' || l === 'departamento origen') return true;
  if (text.includes('departamento') && text.includes('origen')) return true;

  // 8. Teléfono de Contacto
  if (k === 'telefono' || k === 'teléfono de contacto' || k === 'telefono de contacto' || text.includes('teléfono') || text.includes('telefono')) return true;

  return false;
}

// Obtener las 8 columnas solicitadas para la vista con Filtro SQL SAR:
// Nro Ticket , Nombres Afiliado , Nombre Usuario Registro , Correo Afiliado , Servicio , Tiempo Atención , Departamento Origen , Teléfono de Contacto
function getSQLTargetColumns() {
  const findCol = (predicate) => currentColumns.find(predicate);
  
  const colTicket = findCol(c => c.key === 'Nro Ticket' || c.key === 'Nro. Ticket' || c.key === 'ticket') 
                 || findCol(c => (c.key.toLowerCase().includes('ticket') || c.label.toLowerCase().includes('ticket')) && !c.key.toLowerCase().includes('origen') && !c.key.toLowerCase().includes('atenci'));

  const colAfiliado = findCol(c => c.key === 'Nombres Afiliado' || c.key === 'primer_nombre') 
                   || findCol(c => (c.key.toLowerCase().includes('afiliado') || c.label.toLowerCase().includes('afiliado')) && !c.key.toLowerCase().includes('cod') && !c.key.toLowerCase().includes('programa'));

  const colUsuario = findCol(c => c.key === 'Nombre Usuario Registro' || c.key === 'usuario') 
                  || findCol(c => c.key.toLowerCase().includes('usuario') && c.key.toLowerCase().includes('registro') && !c.key.toLowerCase().includes('departamento'));

  const colCorreo = findCol(c => c.key === 'Correo Afiliado' || c.key === 'Correo de Contacto' || c.key === 'correo_afiliado') 
                 || findCol(c => c.key.toLowerCase().includes('correo') && !c.key.toLowerCase().includes('fecha'));

  const colServicio = findCol(c => c.key.trim().toLowerCase() === 'servicio' || c.label.trim().toLowerCase() === 'servicio');

  const colTiempo = findCol(c => c.key === 'Tiempo Atención' || c.key === 'tiempo_atencion' || c.key.includes('Tiempo Atenci')) 
                 || findCol(c => c.key.toLowerCase().includes('tiempo') && c.key.toLowerCase().includes('atenci'));

  const colDeptoOri = findCol(c => c.key === 'Departamento Origen' || c.key === 'depto_origen') 
                   || findCol(c => c.key.toLowerCase().includes('departamento') && c.key.toLowerCase().includes('origen'));

  const colTelefono = findCol(c => c.key === 'Teléfono de Contacto' || c.key === 'Telefono de Contacto' || c.key === 'telefono') 
                   || findCol(c => c.key.toLowerCase().includes('tel') || c.key.toLowerCase().includes('fono'));

  const list = [];
  if (colTicket) list.push({ ...colTicket, label: 'Nro Ticket' });
  if (colAfiliado) list.push({ ...colAfiliado, label: 'Nombres Afiliado' });
  if (colUsuario) list.push({ ...colUsuario, label: 'Nombre Usuario Registro' });
  if (colCorreo) list.push({ ...colCorreo, label: 'Correo Afiliado' });
  if (colServicio) list.push({ ...colServicio, label: 'Servicio' });
  if (colTiempo) list.push({ ...colTiempo, label: 'Tiempo Atención', isNumber: true });
  else list.push({ key: 'Tiempo Atención', label: 'Tiempo Atención', isNumber: true, isCore: true });
  if (colDeptoOri) list.push({ ...colDeptoOri, label: 'Departamento Origen' });
  if (colTelefono) list.push({ ...colTelefono, label: 'Teléfono de Contacto' });

  return list;
}

// Obtener las columnas visibles según el estado (state.visibleColumns)
function getVisibleColumnsList() {
  const visible = currentColumns.filter(col => state.visibleColumns.has(col.key));
  if (visible.length === 0) {
    const sqlCols = getSQLTargetColumns();
    if (sqlCols.length > 0) return sqlCols;
    return currentColumns.slice(0, 8);
  }
  return visible;
}

// ==========================================
// ARRASTRE Y COLOCACIÓN LIBRE DE COLUMNAS (DRAG & DROP)
// ==========================================
let draggedColKey = null;
let isDraggingCol = false;

function handleColDragStart(e, colKey) {
  isDraggingCol = true;
  draggedColKey = colKey;
  if (e.stopPropagation) e.stopPropagation();
  if (e.dataTransfer) {
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', colKey);
  }
  const th = e.currentTarget.closest('th');
  if (th) th.classList.add('th-dragging');
}

function handleColDragOver(e) {
  e.preventDefault();
  if (e.stopPropagation) e.stopPropagation();
  if (e.dataTransfer) {
    e.dataTransfer.dropEffect = 'move';
  }
  const th = e.currentTarget.closest('th');
  if (!th || th.dataset.col === draggedColKey) return;

  const rect = th.getBoundingClientRect();
  const isRightHalf = e.clientX > (rect.left + rect.width / 2);

  if (isRightHalf) {
    th.classList.remove('th-drag-left');
    th.classList.add('th-drag-right');
  } else {
    th.classList.remove('th-drag-right');
    th.classList.add('th-drag-left');
  }
}

function handleColDragLeave(e) {
  const th = e.currentTarget.closest('th');
  if (th) {
    th.classList.remove('th-drag-left', 'th-drag-right');
  }
}

function handleColDragEnd(e) {
  document.querySelectorAll('th').forEach(th => {
    th.classList.remove('th-dragging', 'th-drag-left', 'th-drag-right');
  });
  draggedColKey = null;
  setTimeout(() => { isDraggingCol = false; }, 150);
}

function handleColDrop(e, targetColKey) {
  e.preventDefault();
  if (e.stopPropagation) e.stopPropagation();
  const th = e.currentTarget.closest('th');
  let dropAfter = false;
  if (th) {
    const rect = th.getBoundingClientRect();
    dropAfter = e.clientX > (rect.left + rect.width / 2);
  }

  document.querySelectorAll('th').forEach(t => {
    t.classList.remove('th-dragging', 'th-drag-left', 'th-drag-right');
  });

  if (!draggedColKey || draggedColKey === targetColKey) {
    draggedColKey = null;
    return;
  }

  reorderColumns(draggedColKey, targetColKey, dropAfter);
  draggedColKey = null;
}

// Reordenar columnas colocando la columna exactamente donde el usuario la suelta (antes o después)
function reorderColumns(fromKey, targetKey, dropAfter = false) {
  if (fromKey === targetKey) return;
  const fromIdx = currentColumns.findIndex(c => c.key === fromKey);
  let toIdx = currentColumns.findIndex(c => c.key === targetKey);
  if (fromIdx === -1 || toIdx === -1) return;

  const [movedCol] = currentColumns.splice(fromIdx, 1);
  toIdx = currentColumns.findIndex(c => c.key === targetKey);
  const insertIdx = dropAfter ? toIdx + 1 : toIdx;
  currentColumns.splice(insertIdx, 0, movedCol);

  renderHeaderAndQuickFilters();
  renderTable();
  saveUIState();
  showToast(`Columna "${movedCol.label}" reubicada`, 'info');
}

// Renderizar cabecera y fila de filtro rápido
function renderHeaderAndQuickFilters() {
  const visibleCols = getVisibleColumnsList();

  // Cabecera principal
  elements.tableHeaderRow.innerHTML = visibleCols.map(col => {
    const isFiltered = (state.columnFilters[col.key] && state.columnFilters[col.key].size > 0) || (state.quickFilters[col.key]);
    const isSorted = state.sort.key === col.key;
    const sortArrow = isSorted ? (state.sort.direction === 'asc' ? '▲' : '▼') : '↕';

    return `
      <th data-col="${col.key}"
          draggable="true"
          ondragstart="handleColDragStart(event, '${col.key}')"
          ondragover="handleColDragOver(event)"
          ondragleave="handleColDragLeave(event)"
          ondrop="handleColDrop(event, '${col.key}')"
          ondragend="handleColDragEnd(event)"
          title="Arrastra esta columna para colocarla donde quieras">
        <div class="th-content">
          <div class="col-drag-handle" title="Arrastra para mover">⋮⋮</div>
          
          <div class="th-title" onclick="handleSort('${col.key}')" title="Clic para ordenar">
            <span>${col.label}</span>
            <span class="sort-icon ${isSorted ? 'active' : ''}">${sortArrow}</span>
          </div>

          <button type="button" class="excel-filter-btn ${isFiltered ? 'active' : ''}" 
                  onclick="openExcelFilterPopup(event, '${col.key}')" 
                  title="Filtrar por esta columna tipo Excel">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>
          </button>
        </div>
      </th>
    `;
  }).join('');

  // Fila de filtro rápido
  elements.quickFilterRow.innerHTML = visibleCols.map(col => {
    const curVal = state.quickFilters[col.key] || '';
    return `
      <th>
        <input type="text" class="col-quick-search" data-col="${col.key}" 
               placeholder="Buscar..." value="${curVal}"
               oninput="handleQuickFilterInput('${col.key}', this.value)">
      </th>
    `;
  }).join('');
}

let quickFilterDebounceTimer = null;
// Filtrado rápido por input con Debounce (250ms)
function handleQuickFilterInput(colKey, text) {
  text = text.trim();
  if (text) {
    state.quickFilters[colKey] = text.toLowerCase();
  } else {
    delete state.quickFilters[colKey];
  }
  clearTimeout(quickFilterDebounceTimer);
  quickFilterDebounceTimer = setTimeout(() => {
    state.currentPage = 1;
    applyFiltersAndSort();
  }, 250);
}

// Ordenamiento
function handleSort(colKey) {
  if (isDraggingCol) return;
  if (state.sort.key === colKey) {
    state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
  } else {
    state.sort.key = colKey;
    state.sort.direction = 'asc';
  }
  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

// =========================================================================
// FUNCIONES Y REGLAS DE NEGOCIO: FILTRO DE FECHA ANTERIOR (CORTE 9:00 AM)
// =========================================================================

// Extraer cadena YYYY-MM-DD normalizada directamente desde la columna Fecha Registro (igual que en Excel y BD)
function getRowDateYMD(row) {
  // 1. Columna Fecha Registro (Prioridad 1: Es la fecha oficial de registro en Excel / Base de Datos)
  const val = row._origFechaRegistro || row['Fecha Registro'] || row.fecha_registro || row['Fecha'] || '';
  if (val !== undefined && val !== null && val !== '') {
    if (val instanceof Date) {
      const y = val.getFullYear();
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const d = String(val.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    if (typeof val === 'number' && val > 30000) {
      const excelEpoch = new Date(1899, 11, 30);
      const dObj = new Date(excelEpoch.getTime() + val * 86400000);
      const y = dObj.getFullYear();
      const m = String(dObj.getMonth() + 1).padStart(2, '0');
      const d = String(dObj.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    const str = String(val).trim();
    // Formato ISO / SQL Server: YYYY-MM-DD o YYYY/MM/DD ... (ej: 2026-10-20 23:57:09.000)
    const mIso = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (mIso) {
      const y = mIso[1];
      const m = mIso[2].padStart(2, '0');
      const d = mIso[3].padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    // Formato latino / peruano: DD/MM/YYYY o DD-MM-YYYY ... (ej: 20/10/2026 23:57:09)
    const mLat = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (mLat) {
      const d = mLat[1].padStart(2, '0');
      const m = mLat[2].padStart(2, '0');
      const y = mLat[3];
      return `${y}-${m}-${d}`;
    }
  }

  // 2. Respaldo (Fallback): Solo si la columna Fecha Registro viene vacía, deducir desde el código del ticket
  const ticket = String(row.ticket || row['Nro. Ticket'] || row['Nro Ticket'] || '').trim();
  const mTicket = ticket.match(/^[A-Za-z]{2,4}-(\d{2})(\d{2})(\d{2})-/i);
  if (mTicket) {
    const y = '20' + mTicket[1];
    const m = mTicket[2];
    const d = mTicket[3];
    return `${y}-${m}-${d}`;
  }

  return '';
}

// Actualizar y poblar dinámicamente las opciones del selector de Día Evaluado según la fecha actual (HOY)
function updateSimuladorDiaOptions() {
  if (!elements.simuladorDiaSelect) return;

  const today = new Date();
  const todayKey = getTodayDayKey(today);
  const currentSelected = (state.simulatedDay && state.simulatedDay !== 'auto') ? state.simulatedDay : todayKey;

  const daysDef = [
    { key: 'lunes', label: 'Lunes' },
    { key: 'martes', label: 'Martes' },
    { key: 'miercoles', label: 'Miércoles' },
    { key: 'jueves', label: 'Jueves' },
    { key: 'viernes', label: 'Viernes' },
    { key: 'sabado', label: 'Sábado' },
    { key: 'domingo', label: 'Domingo' }
  ];

  elements.simuladorDiaSelect.innerHTML = daysDef.map(d => {
    const dateForDay = getCurrentWeekDayDate(d.key, today);
    const isToday = d.key === todayKey;
    const isSelected = d.key === currentSelected;

    let desc = '';
    if (d.key === 'lunes') {
      const v = new Date(dateForDay); v.setDate(dateForDay.getDate() - 3);
      const s = new Date(dateForDay); s.setDate(dateForDay.getDate() - 2);
      const dm = new Date(dateForDay); dm.setDate(dateForDay.getDate() - 1);
      desc = `Jala Viernes ${v.getDate()}, Sábado ${s.getDate()} y Domingo ${dm.getDate()}`;
    } else if (['martes', 'miercoles', 'jueves', 'viernes'].includes(d.key)) {
      const prev = new Date(dateForDay); prev.setDate(dateForDay.getDate() - 1);
      const prevNames = { martes: 'Lunes', miercoles: 'Martes', jueves: 'Miércoles', viernes: 'Jueves' };
      desc = `Jala ${prevNames[d.key]} ${prev.getDate()}`;
    } else {
      desc = 'No jala nada';
    }

    const optionText = isToday 
      ? `${d.label} (Hoy - ${desc})`
      : `${d.label} (${desc})`;

    return `<option value="${d.key}" ${isSelected ? 'selected' : ''}>${optionText}</option>`;
  }).join('');

  elements.simuladorDiaSelect.value = currentSelected;
}

// Configuración de fechas objetivo según el día de la semana (por defecto SIEMPRE el día de HOY)
function getTargetDatesConfig(simulatedDay) {
  const today = new Date();
  const todayKey = getTodayDayKey(today);

  // Si no se indica día o es 'auto', SIEMPRE evaluar el día de HOY
  const dayKey = (simulatedDay && simulatedDay !== 'auto') ? String(simulatedDay).toLowerCase() : todayKey;

  // Fecha de referencia: Si es el día de hoy, usar today directamente. De lo contrario, calcular el día dentro de la semana actual.
  let refDate;
  if (dayKey === todayKey) {
    refDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  } else {
    refDate = getCurrentWeekDayDate(dayKey, today);
  }

  const d = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());
  const day = d.getDay(); // 0 = Domingo, 1 = Lunes, 2 = Martes, 3 = Miércoles, 4 = Jueves, 5 = Viernes, 6 = Sábado

  const formatYMD = (date) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const dayNum = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${dayNum}`;
  };

  const addDays = (num) => {
    const res = new Date(d);
    res.setDate(res.getDate() + num);
    return res;
  };

  const isTodayEval = (dayKey === todayKey);
  const tagHoy = isTodayEval ? ' (Hoy)' : '';

  if (day === 1) {
    // Regla: Si es LUNES -> Jala 3 días antes: Domingo (-1), Sábado (-2) y Viernes (-3)
    const dom = formatYMD(addDays(-1));
    const sab = formatYMD(addDays(-2));
    const vie = formatYMD(addDays(-3));
    return {
      dayOfWeek: 1,
      dayName: 'Lunes',
      dayKey: 'lunes',
      isToday: isTodayEval,
      refDate: d,
      targetDates: [dom, sab, vie],
      isWeekend: false,
      summary: `Lunes${tagHoy} detectado: Jalando 3 días antes (Viernes ${vie}, Sábado ${sab} y Domingo ${dom})`
    };
  } else if (day >= 2 && day <= 5) {
    // Regla: Si es MARTES a VIERNES -> Jala 1 día antes (ayer)
    const dayNames = ['', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];
    const ayer = formatYMD(addDays(-1));
    return {
      dayOfWeek: day,
      dayName: dayNames[day],
      dayKey: DIAS_SEMANA_KEYS[day],
      isToday: isTodayEval,
      refDate: d,
      targetDates: [ayer],
      isWeekend: false,
      summary: `${dayNames[day]}${tagHoy} detectado: Jalando 1 día antes (Ayer: ${ayer})`
    };
  } else {
    // Regla: SÁBADO y DOMINGO no deberían jalar nada
    const dayNames = { 0: 'Domingo', 6: 'Sábado' };
    return {
      dayOfWeek: day,
      dayName: dayNames[day],
      dayKey: DIAS_SEMANA_KEYS[day],
      isToday: isTodayEval,
      refDate: d,
      targetDates: [],
      isWeekend: true,
      summary: `${dayNames[day]}${tagHoy} (Fin de semana): Según regla no se jala nada sábado ni domingo.`
    };
  }
}

// Aplicar transformación a las filas coincidentes:
// Fecha Registro -> Fecha de hoy a las 09:00:00
// Mes Registro   -> Mes actual
// Hora Registro  -> 09:00:00
function applyFechaAnteriorTransformation() {
  const cfg = getTargetDatesConfig(state.simulatedDay);
  const targetSet = new Set(cfg.targetDates);

  const refDate = cfg.refDate;
  const dStr = String(refDate.getDate()).padStart(2, '0');
  const mStr = String(refDate.getMonth() + 1).padStart(2, '0');
  const yStr = refDate.getFullYear();

  const hoyStr = `${dStr}/${mStr}/${yStr} 09:00:00`;
  const mesActual = String(refDate.getMonth() + 1);
  const horaActual = '09:00:00';

  let matchCount = 0;

  state.rawData.forEach(row => {
    // Asegurar respaldo de fecha original
    if (row._origFechaRegistro === undefined) {
      row._origFechaRegistro = row['Fecha Registro'] || row.fecha_registro || '';
      row._origMesRegistro = row['Mes Registro'] || row.mes_registro || '';
      row._origHoraRegistro = row['Hora Registro'] || row.hora_registro || '';
      row._origCumpleSar = row.cumple_sar;
    }
    row._origDateYMD = getRowDateYMD(row);

    if (!cfg.isWeekend && targetSet.has(row._origDateYMD)) {
      matchCount++;
      row._cumpleFechaAnterior = true;

      // Colocar Fecha Registro = Fecha de hoy 09:00:00
      row['Fecha Registro'] = hoyStr;
      if (row.fecha_registro !== undefined) row.fecha_registro = hoyStr;

      // Colocar Mes Registro = Mes actual
      row['Mes Registro'] = mesActual;
      if (row.mes_registro !== undefined) row.mes_registro = mesActual;

      // Colocar Hora Registro = 09:00:00
      row['Hora Registro'] = horaActual;
      if (row.hora_registro !== undefined) row.hora_registro = horaActual;

      // Reevaluar dinámicamente cumple_sar con la nueva fecha y hora (Hoy 09:00:00)
      row.cumple_sar = matchesSQLFilterDynamic(row) ? 'S' : 'No';
    } else {
      row._cumpleFechaAnterior = false;
      // Restaurar original si no coincide
      if (row._origFechaRegistro !== undefined) {
        row['Fecha Registro'] = row._origFechaRegistro;
        if (row.fecha_registro !== undefined) row.fecha_registro = row._origFechaRegistro;
        row['Mes Registro'] = row._origMesRegistro;
        if (row.mes_registro !== undefined) row.mes_registro = row._origMesRegistro;
        row['Hora Registro'] = row._origHoraRegistro;
        if (row.hora_registro !== undefined) row.hora_registro = row._origHoraRegistro;
        if (row._origCumpleSar !== undefined) row.cumple_sar = row._origCumpleSar;
      }
    }
  });

  return { cfg, matchCount, hoyStr, mesActual, horaActual };
}

// Restaurar valores originales cuando se apaga el filtro
function restoreOriginalFechaValues() {
  state.rawData.forEach(row => {
    if (row._origFechaRegistro !== undefined) {
      row['Fecha Registro'] = row._origFechaRegistro;
      if (row.fecha_registro !== undefined) row.fecha_registro = row._origFechaRegistro;
      row['Mes Registro'] = row._origMesRegistro;
      if (row.mes_registro !== undefined) row.mes_registro = row._origMesRegistro;
      row['Hora Registro'] = row._origHoraRegistro;
      if (row.hora_registro !== undefined) row.hora_registro = row._origHoraRegistro;
      if (row._origCumpleSar !== undefined) row.cumple_sar = row._origCumpleSar;
      row._cumpleFechaAnterior = false;
    }
  });
}

// Aplicar filtros globales, rápidos y por columna (SIN NINGÚN FILTRO POR DEFECTO)
function applyFiltersAndSort() {
  let result = state.rawData;

  // 0. Filtro inteligente por fecha anterior (Corte 9:00 AM)
  if (state.filterFechaAnterior) {
    result = result.filter(item => item._cumpleFechaAnterior === true);
  }

  // 1. Filtro consulta SQL SAR (si el botón está activo)
  if (state.onlySQLFilter) {
    result = result.filter(item => matchesSQLFilter(item));
  }

  // 2. Filtro global ultra-rápido mediante índice precomputado
  if (state.globalSearch) {
    const term = state.globalSearch;
    result = result.filter(item => {
      const searchStr = item._searchStr || (item._searchStr = buildSearchStr(item));
      return searchStr.includes(term);
    });
  }

  // 2. Filtros rápidos de columna (evaluados en una sola pasada)
  const quickKeys = Object.keys(state.quickFilters);
  if (quickKeys.length > 0) {
    result = result.filter(item => {
      for (let i = 0; i < quickKeys.length; i++) {
        const colKey = quickKeys[i];
        const query = state.quickFilters[colKey];
        if (query) {
          const val = String(item[colKey] ?? '').toLowerCase();
          if (!val.includes(query)) return false;
        }
      }
      return true;
    });
  }

  // 3. Filtros avanzados de columna tipo Excel (evaluados en una sola pasada)
  const activeColFilterKeys = Object.keys(state.columnFilters).filter(k => state.columnFilters[k] && state.columnFilters[k].size > 0);
  if (activeColFilterKeys.length > 0) {
    result = result.filter(item => {
      for (let i = 0; i < activeColFilterKeys.length; i++) {
        const colKey = activeColFilterKeys[i];
        const selectedSet = state.columnFilters[colKey];
        const itemVal = String(item[colKey] ?? '');
        if (!selectedSet.has(itemVal)) return false;
      }
      return true;
    });
  }

  // 4. Ordenamiento de alto rendimiento (evita el costo excesivo de localeCompare en bucles grandes)
  const sortKey = state.sort.key;
  if (sortKey && result.length > 1) {
    const sortDir = state.sort.direction === 'asc' ? 1 : -1;
    const colDef = currentColumns.find(c => c.key === sortKey);
    const isNumber = colDef && colDef.isNumber;

    result.sort((a, b) => {
      const valA = a[sortKey];
      const valB = b[sortKey];

      if (valA === valB) return 0;
      if (valA === undefined || valA === null || valA === '') return 1;
      if (valB === undefined || valB === null || valB === '') return -1;

      if (isNumber) {
        const numA = typeof valA === 'number' ? valA : (parseFloat(valA) || 0);
        const numB = typeof valB === 'number' ? valB : (parseFloat(valB) || 0);
        return (numA - numB) * sortDir;
      }

      // Comparación rápida de cadenas (50x más veloz que localeCompare en arrays grandes)
      const strA = typeof valA === 'string' ? valA.toLowerCase() : String(valA).toLowerCase();
      const strB = typeof valB === 'string' ? valB.toLowerCase() : String(valB).toLowerCase();

      if (strA < strB) return -1 * sortDir;
      if (strA > strB) return 1 * sortDir;
      return 0;
    });
  }

  state.filteredData = result;
  updateKPIs();
  renderActiveFilterBadges();
  renderTable();
  saveUIState();
}

// Renderizar la tabla con paginación
function renderTable() {
  const visibleCols = getVisibleColumnsList();
  const total = state.filteredData.length;
  
  if (total === 0) {
    elements.tableBody.innerHTML = `
      <tr>
        <td colspan="${visibleCols.length}" style="text-align: center; padding: 40px; color: var(--text-muted);">
          No se encontraron registros que coincidan con los filtros aplicados.
        </td>
      </tr>
    `;
    elements.paginationInfo.textContent = '0 de 0 registros';
    elements.paginationControls.innerHTML = '';
    return;
  }

  const startIdx = (state.currentPage - 1) * state.pageSize;
  const endIdx = Math.min(startIdx + state.pageSize, total);
  const pageRows = state.filteredData.slice(startIdx, endIdx);

  elements.tableBody.innerHTML = pageRows.map(row => {
    const cellsHtml = visibleCols.map(col => {
      let val = row[col.key];
      if ((val === undefined || val === '') && (col.key === 'Tiempo Atención' || col.key.includes('Tiempo Atenci'))) {
        val = row.tiempo_atencion;
      }
      val = val ?? '';

      // Formateo especial por columna
      if (col.key === 'ticket' || col.key === 'Nro. Ticket' || col.key === 'Nro Ticket') {
        const ticketVal = row.ticket || row['Nro. Ticket'] || row['Nro Ticket'] || val;
        return `<td class="ticket-cell" onclick="showTicketDetail('${ticketVal}')" style="cursor: pointer; font-weight: 600;" title="Ver detalle">${val}</td>`;
      }
      if (col.key === 'primer_nombre' || (state.onlySQLFilter && (col.label === 'Nombres Afiliado' || col.key === 'Nombres Afiliado'))) {
        const nomVal = state.onlySQLFilter ? resolveNombreAfiliado(row) : (row.primer_nombre || val);
        return `<td style="font-weight: 600; color: var(--text-main);">${nomVal}</td>`;
      }
      if (col.key === 'tiempo_atencion' || col.key === 'Tiempo Atención' || col.key.includes('Tiempo Atenci')) {
        const num = parseFloat(val) || 0;
        let badgeClass = 'medio';
        if (num <= 5) badgeClass = 'rapido';
        else if (num > 15) badgeClass = 'largo';
        return `<td><span class="tiempo-badge ${badgeClass}">${num} min</span></td>`;
      }
      if (col.key === 'cumple_sar' || col.key === 'rechazado' || col.key === 'equipo_contacto') {
        const isSi = val === 'S' || val === 'Sí' || val === 'SI';
        return `<td><span class="tag-pill ${isSi ? 'tag-si' : 'tag-no'}">${val}</span></td>`;
      }
      if (col.key === 'telefono' || col.key === 'Teléfono de Contacto' || col.key === 'Telefono de Contacto') {
        const telVal = row[col.key] || row.telefono || val;
        return `<td>${telVal && telVal !== '-' ? `<a href="tel:${telVal}" style="color: inherit; text-decoration: none; font-weight: 500;">📞 ${telVal}</a>` : '-'}</td>`;
      }
      if (col.key === 'correo_afiliado' || col.key === 'Correo de Contacto' || col.key === 'Correo Afiliado' || col.label === 'Correo Afiliado') {
        const mailVal = state.onlySQLFilter ? resolveCorreoAfiliado(row) : (row.correo_afiliado || row[col.key] || val);
        const isEmail = String(mailVal).includes('@');
        return `<td>${isEmail ? `<a href="mailto:${mailVal}" style="color: var(--primary); text-decoration: none;">✉️ ${mailVal}</a>` : `<span style="color: var(--text-muted);">${mailVal}</span>`}</td>`;
      }

      // Truncar textos largos (como Observación Inicial) para que la tabla sea legible
      const strVal = String(val);
      if (strVal.length > 70) {
        return `<td title="${strVal.replace(/"/g, '&quot;')}">${strVal.substring(0, 67)}...</td>`;
      }
      return `<td title="${strVal.replace(/"/g, '&quot;')}">${strVal}</td>`;
    }).join('');

    return `<tr>${cellsHtml}</tr>`;
  }).join('');

  // Actualizar barra de paginación
  elements.paginationInfo.textContent = `Mostrando ${startIdx + 1} - ${endIdx} de ${total.toLocaleString()} registros`;
  renderPaginationControls(total);
}

// Paginación
function renderPaginationControls(total) {
  const totalPages = Math.ceil(total / state.pageSize);
  if (totalPages <= 1) {
    elements.paginationControls.innerHTML = '';
    return;
  }

  let html = `
    <button class="page-btn" ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(1)" title="Primera página">«</button>
    <button class="page-btn" ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(${state.currentPage - 1})" title="Anterior">‹</button>
  `;

  let startP = Math.max(1, state.currentPage - 2);
  let endP = Math.min(totalPages, state.currentPage + 2);

  if (startP > 1) html += `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;

  for (let p = startP; p <= endP; p++) {
    html += `
      <button class="page-btn ${p === state.currentPage ? 'active' : ''}" onclick="changePage(${p})">${p}</button>
    `;
  }

  if (endP < totalPages) html += `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;

  html += `
    <button class="page-btn" ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${state.currentPage + 1})" title="Siguiente">›</button>
    <button class="page-btn" ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${totalPages})" title="Última página">»</button>
  `;

  elements.paginationControls.innerHTML = html;
}

function changePage(page) {
  state.currentPage = page;
  renderTable();
  document.querySelector('.table-responsive').scrollTop = 0;
}

// Actualizar tarjetas de KPI
function updateKPIs() {
  const totalRaw = state.rawData.length;
  const filtered = state.filteredData;
  const filteredCount = filtered.length;

  elements.kpiTotalTickets.textContent = totalRaw.toLocaleString();
  elements.kpiFilteredTickets.textContent = filteredCount.toLocaleString();


  // Tiempo promedio de atención
  if (filteredCount > 0) {
    const sumTiempo = filtered.reduce((acc, row) => acc + (parseFloat(row.tiempo_atencion) || 0), 0);
    const avg = (sumTiempo / filteredCount).toFixed(1);
    elements.kpiAvgTiempo.textContent = `${avg} min`;
  } else {
    elements.kpiAvgTiempo.textContent = '0 min';
  }

  // Departamentos únicos (opcional si existe el elemento en UI)
  if (elements.kpiDeptosCount) {
    const deptosSet = new Set(filtered.map(r => r.depto_origen).filter(Boolean));
    elements.kpiDeptosCount.textContent = deptosSet.size.toLocaleString();
  }
}

// Renderizar badges de filtros activos
function renderActiveFilterBadges() {
  const activeEntries = [];

  // Si ambos filtros principales están activos, agregar badge destacado de Filtro Encadenado
  if (state.onlySQLFilter && state.filterFechaAnterior) {
    const cfg = getTargetDatesConfig(state.simulatedDay);
    activeEntries.push({
      type: 'chained_filters',
      key: 'chained_filters',
      label: `🔗 FILTRO ENCADENADO: Fecha Anterior [${cfg.dayName}${cfg.isToday ? ' (Hoy)' : ''}] ➔ Filtro SQL SAR (${state.filteredData.length.toLocaleString()} registros)`
    });
  }

  // Filtro de consulta SQL SAR
  if (state.onlySQLFilter) {
    activeEntries.push({
      type: 'sql',
      key: 'sql',
      label: '⚡ Filtro SQL SAR: Call Center / Horarios SLA'
    });
  }

  // Filtro de Fecha Anterior (Corte 9:00 AM)
  if (state.filterFechaAnterior) {
    const cfg = getTargetDatesConfig(state.simulatedDay);
    const dateDesc = cfg.isWeekend ? 'Fin de semana: 0 registros' : `Fechas: ${cfg.targetDates.join(', ')}`;
    activeEntries.push({
      type: 'fecha_anterior',
      key: 'fecha_anterior',
      label: `📅 Fecha Anterior [${cfg.dayName}${cfg.isToday ? ' (Hoy)' : ''}]: ${dateDesc} → Asignado: Fecha Hoy 09:00:00, Mes ${cfg.refDate.getMonth() + 1}, Hora 09:00:00`
    });
  }

  // Filtros rápidos
  Object.keys(state.quickFilters).forEach(key => {
    const colDef = currentColumns.find(c => c.key === key);
    activeEntries.push({
      type: 'quick',
      key,
      label: `${colDef ? colDef.label : key}: contiene "${state.quickFilters[key]}"`
    });
  });

  // Filtros avanzados por columna
  Object.keys(state.columnFilters).forEach(key => {
    const selected = state.columnFilters[key];
    if (selected && selected.size > 0) {
      const colDef = currentColumns.find(c => c.key === key);
      const valStr = selected.size === 1 ? Array.from(selected)[0] : `${selected.size} valores`;
      activeEntries.push({
        type: 'column',
        key,
        label: `${colDef ? colDef.label : key}: ${valStr}`
      });
    }
  });

  // Búsqueda global
  if (state.globalSearch) {
    activeEntries.push({
      type: 'global',
      key: 'global',
      label: `Búsqueda global: "${state.globalSearch}"`
    });
  }

  if (activeEntries.length > 0) {
    elements.activeFiltersBar.style.display = 'flex';
    elements.btnClearAllFilters.style.display = 'inline-flex';
    elements.filterBadgesContainer.innerHTML = activeEntries.map(entry => `
      <div class="filter-badge">
        <span>${entry.label}</span>
        <span class="filter-badge-remove" onclick="removeActiveFilter('${entry.type}', '${entry.key}')" title="Quitar filtro">×</span>
      </div>
    `).join('');
  } else {
    elements.activeFiltersBar.style.display = 'none';
    elements.btnClearAllFilters.style.display = 'none';
  }
}

function removeActiveFilter(type, key) {
  if (type === 'chained_filters') {
    state.onlySQLFilter = false;
    state.filterFechaAnterior = false;
    restoreOriginalFechaValues();
    if (elements.btnFiltroSQL) {
      elements.btnFiltroSQL.classList.remove('btn-primary');
      elements.btnFiltroSQL.classList.add('btn-outline');
      elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';
    }
    if (elements.btnFiltroFechaAnterior) {
      elements.btnFiltroFechaAnterior.classList.remove('btn-fecha-activa');
      elements.btnFiltroFechaAnterior.classList.add('btn-outline');
      elements.filtroFechaAnteriorLabel.textContent = '📅 Filtro Fecha Anterior (9:00 AM)';
    }
    if (elements.simuladorDiaBox) elements.simuladorDiaBox.style.display = 'none';
    if (state.previousVisibleColumns && state.previousVisibleColumns.size > 0) {
      state.visibleColumns = new Set(state.previousVisibleColumns);
    } else {
      state.visibleColumns = new Set(currentColumns.map(c => c.key));
    }
    state.previousVisibleColumns = null;
    showToast('Filtros encadenados removidos: mostrando base completa', 'info');
  } else if (type === 'sql') {
    state.onlySQLFilter = false;
    if (elements.btnFiltroSQL) {
      elements.btnFiltroSQL.classList.remove('btn-primary');
      elements.btnFiltroSQL.classList.add('btn-outline');
      elements.filtroSQLLabel.textContent = '⚡ Filtro SQL (SAR)';
    }
    if (state.filterFechaAnterior && elements.filtroFechaAnteriorLabel) {
      elements.filtroFechaAnteriorLabel.textContent = '✓ Fecha Anterior Activo (9:00 AM)';
    }
    if (state.previousVisibleColumns && state.previousVisibleColumns.size > 0) {
      state.visibleColumns = new Set(state.previousVisibleColumns);
    } else {
      state.visibleColumns = new Set(currentColumns.map(c => c.key));
    }
    state.previousVisibleColumns = null;
    showToast('Filtro SQL removido', 'info');
  } else if (type === 'fecha_anterior') {
    state.filterFechaAnterior = false;
    restoreOriginalFechaValues();
    if (elements.btnFiltroFechaAnterior) {
      elements.btnFiltroFechaAnterior.classList.remove('btn-fecha-activa');
      elements.btnFiltroFechaAnterior.classList.add('btn-outline');
      elements.filtroFechaAnteriorLabel.textContent = '📅 Filtro Fecha Anterior (9:00 AM)';
    }
    if (elements.simuladorDiaBox) {
      elements.simuladorDiaBox.style.display = 'none';
    }
    if (state.onlySQLFilter && elements.filtroSQLLabel) {
      elements.filtroSQLLabel.textContent = '✓ Filtro SQL Activo (8 Col)';
    }
    showToast('Filtro de fecha anterior removido: valores originales restaurados', 'info');
  } else if (type === 'quick') {
    delete state.quickFilters[key];
    const inp = document.querySelector(`.col-quick-search[data-col="${key}"]`);
    if (inp) inp.value = '';
  } else if (type === 'column') {
    delete state.columnFilters[key];
  } else if (type === 'global') {
    state.globalSearch = '';
    elements.globalSearchInput.value = '';
  }
  
  state.currentPage = 1;
  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

// Evaluación fiel de las condiciones WHERE de la consulta SQL SAR
function matchesSQLFilter(item) {
  // Si no está activo el filtro de fecha anterior y la fila no fue transformada, usar valor precalculado si existe
  if (!state.filterFechaAnterior && !item._cumpleFechaAnterior) {
    if (item.cumple_sar === 'S') return true;
    if (item.cumple_sar === 'No') return false;
  }

  return matchesSQLFilterDynamic(item);
}

function matchesSQLFilterDynamic(item) {
  const ref = getTablasReferencia();

  const deptoOri = String(item.depto_origen || item['Departamento Origen'] || '').trim();
  const deptoOriLower = deptoOri.toLowerCase();
  const deptoUsu = String(item.depto_usuario || item['Departamento Usuario Registro'] || '').trim().toLowerCase();
  const ticket = String(item.ticket || item['Nro. Ticket'] || item['Nro Ticket'] || '').trim();
  const tipoServ = String(item.tipo_servicio || item['Tipo Servicio'] || '').trim().toLowerCase();
  const estado = String(item.estado || item['Estado Servicio'] || '').trim().toLowerCase();
  const servicio = String(item.servicio || item['Servicio'] || '').trim();
  const entidadReg = String(item['Entidades Reguladoras'] || item.entidad_reguladora || '').trim();

  // 1. Regulador: [Base SAR V1] LEFT JOIN Reguladores ON [Base SAR V1].[Entidades Reguladoras] = Reguladores.[ENTIDADES REGULADORAS]
  // WHERE ((Reguladores.Regulador)="TRANSACCIONES DIARIAS" Or (Reguladores.Regulador) Is Null)
  let regulador = item.regulador;
  if (!regulador && entidadReg) {
    regulador = ref.reguladores[entidadReg] || null;
  }
  const isRegOk = !regulador || regulador === 'TRANSACCIONES DIARIAS' || regulador === '(Sin regulador)';

  // 2. [Servicios Rechazados].[Servicio] Is Null
  const isNoRech = !ref.serviciosRechazados.includes(servicio);

  // 3. [Base SAR V1].[Tipo Servicio] IN ("requerimiento","reclamo")
  const isTipoOk = tipoServ === 'requerimiento' || tipoServ === 'reclamo';

  // 4. [Base SAR V1].[Estado Servicio] = "En curso"
  const isEstadoOk = estado === 'en curso';

  if (!isRegOk || !isNoRech || !isTipoOk || !isEstadoOk) return false;

  // Rama 1: Call Center / CCG
  const esCallCenter = deptoOriLower.includes('call center') || 
                       deptoUsu.includes('call center') || 
                       ticket.toLowerCase().startsWith('ccg-');
  if (esCallCenter) return true;

  // Rama 2: Horarios / SLA
  const isNotInterno = !deptoOriLower.includes('interno');
  if (!isNotInterno) return false;

  // [Tiempo de Atención Servicios].[Servicio] Is Not Null
  const tieneTiempo = ref.tiemposAtencion.hasOwnProperty(servicio);
  if (!tieneTiempo) return false;

  // Parsear Fecha y Hora Registro
  let fechaStr = String(item['Fecha Registro'] || item.fecha_registro || '').trim();
  let horaStr = String(item['Hora Registro'] || item.hora_registro || '00:00:00').trim();

  const tm = horaStr.match(/\d{2}:\d{2}:\d{2}/) || horaStr.match(/\d{2}:\d{2}/);
  if (tm) {
    horaStr = tm[0].length === 5 ? tm[0] + ':00' : tm[0];
  } else {
    horaStr = '00:00:00';
  }

  let fechaObj = null;
  const fMatch = fechaStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (fMatch) {
    fechaObj = new Date(parseInt(fMatch[3], 10), parseInt(fMatch[2], 10) - 1, parseInt(fMatch[1], 10));
  } else {
    const isoMatch = fechaStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      fechaObj = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
    }
  }

  if (!fechaObj) return false;
  fechaObj.setHours(0, 0, 0, 0);

  // Fecha de referencia: usar la del simulador/hoy para que coincida exactamente
  const cfg = (typeof getTargetDatesConfig === 'function') ? getTargetDatesConfig(state.simulatedDay) : null;
  const refDate = (cfg && cfg.refDate) ? cfg.refDate : new Date();
  const today = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());
  const isMonday = today.getDay() === 1;

  const oneDayMs = 24 * 60 * 60 * 1000;
  const diffDays = Math.round((today - fechaObj) / oneDayMs);

  if (!isMonday) {
    // Si NO es Lunes:
    // 1. Hoy <= 17:00:00
    if (diffDays === 0 && horaStr <= '17:00:00') return true;
    // 2. Ayer > 17:00:00
    if (diffDays === 1 && horaStr > '17:00:00') return true;
  } else {
    // Si ES Lunes:
    // 1. Viernes (Hoy - 3) > 17:00:00
    if (diffDays === 3 && horaStr > '17:00:00') return true;
    // 2. Sábado (Hoy - 2)
    if (diffDays === 2) return true;
    // 3. Domingo (Hoy - 1)
    if (diffDays === 1) return true;
    // 4. Lunes (Hoy) <= 17:00:00
    if (diffDays === 0 && horaStr <= '17:00:00') return true;
  }

  return false;
}

// Obtener dataset relevante para calcular valores y conteos del popup de filtro de una columna
function getDataForColumnFilter(colKey) {
  let dataset = state.rawData;

  // 0. Filtro inteligente por fecha anterior (Corte 9:00 AM)
  if (state.filterFechaAnterior) {
    dataset = dataset.filter(item => item._cumpleFechaAnterior === true);
  }

  // 1. Filtro consulta SQL SAR (si está activo, reduce a la cantidad filtrada)
  if (state.onlySQLFilter) {
    dataset = dataset.filter(item => matchesSQLFilter(item));
  }

  // 2. Búsqueda global
  if (state.globalSearch) {
    const term = state.globalSearch;
    dataset = dataset.filter(item => {
      return Object.values(item).some(val => 
        String(val).toLowerCase().includes(term)
      );
    });
  }

  // 3. Filtros rápidos de otras columnas
  Object.keys(state.quickFilters).forEach(k => {
    if (k !== colKey) {
      const query = state.quickFilters[k];
      if (query) {
        dataset = dataset.filter(item => {
          const val = String(item[k] || '').toLowerCase();
          return val.includes(query);
        });
      }
    }
  });

  // 4. Filtros tipo Excel de otras columnas
  Object.keys(state.columnFilters).forEach(k => {
    if (k !== colKey) {
      const selectedSet = state.columnFilters[k];
      if (selectedSet && selectedSet.size > 0) {
        dataset = dataset.filter(item => {
          const itemVal = String(item[k] ?? '');
          return selectedSet.has(itemVal);
        });
      }
    }
  });

  return dataset;
}

// ==========================================
// POPUP DE FILTRO TIPO EXCEL
// ==========================================
function openExcelFilterPopup(event, colKey) {
  event.stopPropagation();
  closeAllExcelPopups();

  state.activePopupCol = colKey;
  const button = event.currentTarget;
  const colDef = currentColumns.find(c => c.key === colKey);

  // Calcular valores únicos y sus frecuencias dependientes del filtro activo (cantidad menor)
  const relevantDataset = getDataForColumnFilter(colKey);
  const countsMap = new Map();
  relevantDataset.forEach(row => {
    const val = String(row[colKey] ?? '');
    countsMap.set(val, (countsMap.get(val) || 0) + 1);
  });

  const distinctValues = Array.from(countsMap.keys()).sort((a, b) => 
    a.localeCompare(b, 'es', { numeric: true, sensitivity: 'base' })
  );

  const currentlySelected = state.columnFilters[colKey]
    ? new Set([...state.columnFilters[colKey]].filter(v => countsMap.has(v)))
    : new Set(distinctValues);

  const isAllChecked = distinctValues.length > 0 && distinctValues.every(v => currentlySelected.has(v));

  // Crear popup
  const popup = document.createElement('div');
  popup.className = 'excel-popup';
  popup.id = 'activeExcelPopup';

  popup.innerHTML = `
    <div class="excel-popup-header">
      <span class="excel-popup-title">${colDef ? colDef.label : colKey}</span>
      <button class="excel-popup-close" onclick="closeAllExcelPopups()">&times;</button>
    </div>

    <div class="excel-popup-actions">
      <button class="excel-popup-btn" onclick="applySortAndClose('${colKey}', 'asc')">
        🔼 Ordenar de A a Z
      </button>
      <button class="excel-popup-btn" onclick="applySortAndClose('${colKey}', 'desc')">
        🔽 Ordenar de Z a A
      </button>
      <button class="excel-popup-btn" onclick="clearColumnFilterAndClose('${colKey}')" style="color: var(--danger);">
        ↺ Borrar filtro de esta columna
      </button>
    </div>

    <div class="excel-popup-search">
      <input type="text" id="excelPopupSearchInput" placeholder="Buscar en lista..." oninput="filterPopupCheckboxes(this.value)">
    </div>

    <div class="excel-popup-list" id="excelPopupCheckboxList">
      <label class="excel-checkbox-item" style="font-weight: 600; border-bottom: 1px solid var(--border); padding-bottom: 4px;">
        <span class="excel-checkbox-label">
          <input type="checkbox" id="excelSelectAllCb" onchange="toggleSelectAllPopupCheckboxes(this.checked)" ${isAllChecked ? 'checked' : ''}>
          (Seleccionar todo)
        </span>
        <span class="excel-checkbox-count">${distinctValues.length}</span>
      </label>

      ${distinctValues.map((val) => {
        const isChecked = currentlySelected.has(val);
        const count = countsMap.get(val);
        const displayVal = val === '' ? '(Vacíos)' : val;
        return `
          <label class="excel-checkbox-item popup-cb-item" data-val="${encodeURIComponent(val)}">
            <span class="excel-checkbox-label">
              <input type="checkbox" class="excel-val-cb" value="${encodeURIComponent(val)}" ${isChecked ? 'checked' : ''}>
              <span>${displayVal}</span>
            </span>
            <span class="excel-checkbox-count">${count}</span>
          </label>
        `;
      }).join('')}
    </div>

    <div class="excel-popup-footer">
      <button class="btn btn-outline btn-sm" onclick="closeAllExcelPopups()">Cancelar</button>
      <button class="btn btn-primary btn-sm" onclick="saveExcelColumnFilter('${colKey}')">Aceptar</button>
    </div>
  `;

  document.body.appendChild(popup);

  // Posicionar el popup debajo del botón
  const rect = button.getBoundingClientRect();
  const popupWidth = 300;
  let leftPos = rect.left + window.scrollX;
  if (leftPos + popupWidth > window.innerWidth - 20) {
    leftPos = window.innerWidth - popupWidth - 20;
  }

  popup.style.top = `${rect.bottom + window.scrollY + 6}px`;
  popup.style.left = `${Math.max(10, leftPos)}px`;
}

function closeAllExcelPopups() {
  const existing = document.getElementById('activeExcelPopup');
  if (existing) existing.remove();
  state.activePopupCol = null;
}

function filterPopupCheckboxes(query) {
  query = query.toLowerCase().trim();
  document.querySelectorAll('.popup-cb-item').forEach(item => {
    const val = decodeURIComponent(item.dataset.val).toLowerCase();
    item.style.display = val.includes(query) ? 'flex' : 'none';
  });
}

function toggleSelectAllPopupCheckboxes(checked) {
  document.querySelectorAll('.popup-cb-item').forEach(item => {
    if (item.style.display !== 'none') {
      const cb = item.querySelector('.excel-val-cb');
      if (cb) cb.checked = checked;
    }
  });
}

function saveExcelColumnFilter(colKey) {
  const selected = new Set();
  const checkboxes = document.querySelectorAll('.excel-val-cb');
  let allChecked = true;

  checkboxes.forEach(cb => {
    if (cb.checked) {
      selected.add(decodeURIComponent(cb.value));
    } else {
      allChecked = false;
    }
  });

  if (allChecked || selected.size === checkboxes.length) {
    delete state.columnFilters[colKey];
  } else {
    state.columnFilters[colKey] = selected;
  }

  closeAllExcelPopups();
  state.currentPage = 1;
  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

function clearColumnFilterAndClose(colKey) {
  delete state.columnFilters[colKey];
  closeAllExcelPopups();
  state.currentPage = 1;
  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

function applySortAndClose(colKey, direction) {
  state.sort.key = colKey;
  state.sort.direction = direction;
  closeAllExcelPopups();
  renderHeaderAndQuickFilters();
  applyFiltersAndSort();
}

// ==========================================
// MODAL SELECTOR DE COLUMNAS
// ==========================================
function openColumnModal() {
  if (elements.colSearchInput) elements.colSearchInput.value = '';

  // Si el filtro SQL está activo y visibleColumns aún tiene todas las columnas, sincronizar con las 8 SQL
  if (state.onlySQLFilter && (!state.visibleColumns || state.visibleColumns.size > 8)) {
    const sqlCols = getSQLTargetColumns();
    state.visibleColumns = new Set(sqlCols.map(c => c.key));
  }

  // Ordenar para mostrar las columnas SQL SAR principales primero si el filtro está activo, o el orden actual de currentColumns
  let displayCols = [...currentColumns];
  if (state.onlySQLFilter) {
    displayCols.sort((a, b) => {
      const aCore = isCoreSQLCol(a.key, a.label) ? 1 : 0;
      const bCore = isCoreSQLCol(b.key, b.label) ? 1 : 0;
      if (aCore !== bCore) return bCore - aCore;
      return 0;
    });
  }

  elements.columnTogglesContainer.innerHTML = displayCols.map(col => {
    const isCore = isCoreSQLCol(col.key, col.label);
    const isChecked = state.visibleColumns.has(col.key);

    return `
      <label class="col-checkbox-card ${isCore ? 'core-col' : ''}">
        <input type="checkbox" class="col-toggle-checkbox" data-col="${col.key}" ${isChecked ? 'checked' : ''}>
        <div>
          <div style="font-weight: 600; font-size: 13px; color: ${isCore ? '#1d4ed8' : 'inherit'};">
            ${col.label} ${isCore ? '★' : ''}
          </div>
          <div style="font-size: 11px; color: ${isCore ? '#2563eb' : 'var(--text-muted)'}; font-weight: ${isCore ? '600' : 'normal'};">
            ${isCore ? '<span class="core-badge">★ Columna SQL SAR</span>' : 'Columna del archivo'}
          </div>
        </div>
      </label>
    `;
  }).join('');

  elements.columnModal.classList.add('open');
}

function applyColumnVisibility() {
  const newSet = new Set();
  document.querySelectorAll('.col-toggle-checkbox:checked').forEach(cb => {
    newSet.add(cb.dataset.col);
  });

  if (newSet.size === 0) {
    showToast('Debe haber al menos una columna seleccionada.', 'warning');
    return;
  }

  elements.columnModal.classList.remove('open');

  runWithLoading({
    title: 'Actualizando Columnas Visibles',
    iconType: 'columns',
    badgeText: `${newSet.size} columnas seleccionadas`,
    badgeIcon: '📊',
    initialStatus: 'Reconfigurando estructura de la tabla...',
    initialSubstatus: 'Aplicando visibilidad de columnas...',
    successStatus: '¡Columnas Actualizadas!',
    successSubstatus: `Mostrando ${newSet.size} columnas seleccionadas`,
    holdSuccessMs: 350
  }, () => {
    state.visibleColumns = newSet;
    renderHeaderAndQuickFilters();
    renderTable();
    saveUIState();
  }).then(() => {
    showToast(`Columnas actualizadas (${newSet.size} visibles)`, 'info');
  });
}

// ==========================================
// DETALLE DEL REGISTRO / TICKET
// ==========================================
function showTicketDetail(ticketNumber) {
  const record = state.rawData.find(r => 
    (r.ticket && String(r.ticket) === String(ticketNumber)) || 
    (r['Nro. Ticket'] && String(r['Nro. Ticket']) === String(ticketNumber)) || 
    (r['Nro Ticket'] && String(r['Nro Ticket']) === String(ticketNumber)) ||
    (Object.values(r)[0] && String(Object.values(r)[0]) === String(ticketNumber))
  );
  if (!record) return;

  const displayTitle = record.ticket || record['Nro. Ticket'] || record['Nro Ticket'] || ticketNumber;
  elements.detailModalTitle.textContent = `Detalle: ${displayTitle}`;
  elements.detailContent.innerHTML = currentColumns.map(col => {
    const val = record[col.key] || '-';
    return `
      <div class="detail-item ${col.key === 'afiliado' || col.key === 'servicio' || String(val).length > 40 ? 'full' : ''}">
        <span class="detail-label">${col.label}</span>
        <span class="detail-value">${val}</span>
      </div>
    `;
  }).join('');

  elements.detailModal.classList.add('open');
}

// ==========================================
// ==========================================
// EXPORTAR A EXCEL (XLSX MULTI-HOJAS / CSV)
// ==========================================
function normalizeDepto(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function getRowTiempoAtencion(row) {
  let t = row['Tiempo Atención'] ?? row.tiempo_atencion;
  if (t === undefined || t === null || t === '') {
    const serv = String(row['Servicio'] || row.servicio || '').trim();
    const ref = getTablasReferencia();
    if (ref.tiemposAtencion && ref.tiemposAtencion.hasOwnProperty(serv)) {
      t = ref.tiemposAtencion[serv];
    } else {
      t = 15;
    }
  }
  const match = String(t ?? '').match(/\d+/);
  return match ? parseInt(match[0], 10) : 0;
}

// Clasificación fiel en las 5 hojas solicitadas
function categorizeRowSQL(row) {
  const dNorm = normalizeDepto(row['Departamento Origen'] || row.depto_origen);

  // 1. ONCOSALUD: Back Office - Correos y Plataforma - Módulo
  if (dNorm === 'back office - correos' || dNorm === 'plataforma - modulo' || dNorm.startsWith('plataforma - modulo')) {
    return 'ONCOSALUD';
  }

  // 2. AUNA SALUD: Back Office - Correo Auna Salud
  if (dNorm === 'back office - correo auna salud') {
    return 'AUNA SALUD';
  }

  // Departamentos Call Center
  const callDeptos = [
    'back office - callcenter',
    'back office - call center',
    'call center - auna salud',
    'call center - diamante',
    'call center - oro',
    'call center - plata',
    'call center - renovacion blindaje',
    'call center - retenciones'
  ];

  if (callDeptos.includes(dNorm)) {
    const tiempo = getRowTiempoAtencion(row);
    if (tiempo === 15) return 'CALL SOPORTE';
    if (tiempo === 5) return 'CALL DESPACHO';
    if (tiempo === 10) return 'CALL MODIFICACIONES';
  }

  return 'OTROS';
}

function formatRowForExport(row, cols) {
  return cols.map(col => {
    let val = row[col.key];
    if ((val === undefined || val === null || val === '') && (col.key === 'Tiempo Atención' || col.key === 'tiempo_atencion' || col.key.includes('Tiempo Atenci'))) {
      val = row.tiempo_atencion ?? row['Tiempo Atención'] ?? '';
    }
    if (state.onlySQLFilter && (col.label === 'Nombres Afiliado' || col.key === 'Nombres Afiliado')) {
      val = resolveNombreAfiliado(row);
    }
    if (state.onlySQLFilter && (col.label === 'Correo Afiliado' || col.key === 'Correo Afiliado')) {
      val = resolveCorreoAfiliado(row);
    }
    if ((col.label === 'Teléfono de Contacto' || col.key === 'Teléfono de Contacto') && row.telefono) {
      val = row.telefono;
    }
    return val ?? '';
  });
}

function exportToExcelCSV() {
  const visibleCols = getVisibleColumnsList();
  const rows = state.filteredData;

  if (rows.length === 0) {
    showToast('No hay datos para exportar.', 'warning');
    return;
  }

  const isSQL = state.onlySQLFilter;

  runWithLoading({
    title: isSQL ? 'Generando Excel en 5 Hojas' : 'Exportando Datos a Excel',
    iconType: 'export',
    badgeText: `${rows.length.toLocaleString()} filas`,
    badgeIcon: '📥',
    initialStatus: isSQL ? 'Distribuyendo registros en las 5 hojas...' : 'Generando archivo Excel (.xlsx)...',
    initialSubstatus: isSQL ? 'ONCOSALUD, AUNA SALUD, CALL SOPORTE, CALL DESPACHO, CALL MODIFICACIONES' : 'Preparando celdas y codificación...',
    successStatus: '¡Archivo Excel Generado Exitosamente!',
    successSubstatus: 'Descarga iniciada en su equipo',
    holdSuccessMs: 450
  }, () => {
    if (typeof XLSX !== 'undefined') {
      const wb = XLSX.utils.book_new();
      const headers = visibleCols.map(c => c.label);

      if (isSQL) {
        // Diccionario de las 5 hojas ordenadas
        const sheetsData = {
          'ONCOSALUD': [],
          'AUNA SALUD': [],
          'CALL SOPORTE': [],
          'CALL DESPACHO': [],
          'CALL MODIFICACIONES': [],
          'OTROS': []
        };

        rows.forEach(row => {
          const cat = categorizeRowSQL(row);
          if (sheetsData[cat]) {
            sheetsData[cat].push(row);
          } else {
            sheetsData['OTROS'].push(row);
          }
        });

        const sheetNames = ['ONCOSALUD', 'AUNA SALUD', 'CALL SOPORTE', 'CALL DESPACHO', 'CALL MODIFICACIONES'];
        if (sheetsData['OTROS'].length > 0) {
          sheetNames.push('OTROS');
        }

        sheetNames.forEach(sheetName => {
          const sheetRows = sheetsData[sheetName] || [];
          const aoa = [headers, ...sheetRows.map(r => formatRowForExport(r, visibleCols))];
          const ws = XLSX.utils.aoa_to_sheet(aoa);
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });

        const fechaHoy = new Date().toISOString().slice(0, 10);
        const fileName = `Occess_SAR_Filtrado_5Hojas_${fechaHoy}.xlsx`;
        XLSX.writeFile(wb, fileName);

        return {
          isSQL: true,
          counts: {
            onco: sheetsData['ONCOSALUD'].length,
            auna: sheetsData['AUNA SALUD'].length,
            soporte: sheetsData['CALL SOPORTE'].length,
            despacho: sheetsData['CALL DESPACHO'].length,
            modificaciones: sheetsData['CALL MODIFICACIONES'].length,
            otros: sheetsData['OTROS'].length
          }
        };
      } else {
        // Exportación normal en 1 sola hoja
        const aoa = [headers, ...rows.map(r => formatRowForExport(r, visibleCols))];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, 'Datos');
        const fileName = `Occess_Export_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        return { isSQL: false };
      }
    } else {
      // Fallback CSV si XLSX no estuviera cargado
      let csvContent = '\uFEFF';
      csvContent += visibleCols.map(c => `"${c.label.replace(/"/g, '""')}"`).join(';') + '\r\n';

      rows.forEach(row => {
        const line = formatRowForExport(row, visibleCols).map(val => `"${String(val).replace(/"/g, '""')}"`).join(';');
        csvContent += line + '\r\n';
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', `Datos_Exportados_${new Date().toISOString().slice(0, 10)}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      return { isSQL: false };
    }
  }).then((res) => {
    if (res && res.isSQL) {
      const c = res.counts;
      let msg = `✅ Excel 5 hojas generado: ONCOSALUD (${c.onco}), AUNA SALUD (${c.auna}), SOPORTE (${c.soporte}), DESPACHO (${c.despacho}), MODIFICACIONES (${c.modificaciones})`;
      if (c.otros > 0) msg += `, OTROS (${c.otros})`;
      showToast(msg, 'success');
    } else {
      showToast(`✅ ${rows.length.toLocaleString()} registros exportados correctamente a Excel`, 'success');
    }
  });
}

function copyToClipboard() {
  const visibleCols = getVisibleColumnsList();
  const rows = state.filteredData;

  if (rows.length === 0) {
    showToast('No hay datos para copiar.', 'warning');
    return;
  }

  runWithLoading({
    title: 'Copiando al Portapapeles',
    iconType: 'copy',
    badgeText: `${rows.length.toLocaleString()} registros`,
    badgeIcon: '📋',
    initialStatus: 'Formateando registros para hoja de cálculo...',
    initialSubstatus: 'Generando formato TSV compatible con Excel...',
    successStatus: '¡Copiado al Portapapeles!',
    successSubstatus: 'Listo para pegar en Excel con Ctrl + V',
    holdSuccessMs: 400
  }, () => {
    let tsvContent = visibleCols.map(c => c.label).join('\t') + '\n';
    rows.forEach(row => {
      tsvContent += visibleCols.map(col => {
        let val = row[col.key];
        if ((val === undefined || val === '') && (col.key === 'Tiempo Atención' || col.key.includes('Tiempo Atenci'))) {
          val = row.tiempo_atencion;
        }
        if (state.onlySQLFilter && (col.label === 'Nombres Afiliado' || col.key === 'Nombres Afiliado')) {
          val = resolveNombreAfiliado(row);
        }
        if (state.onlySQLFilter && (col.label === 'Correo Afiliado' || col.key === 'Correo Afiliado')) {
          val = resolveCorreoAfiliado(row);
        }
        if ((col.label === 'Teléfono de Contacto' || col.key === 'Teléfono de Contacto') && row.telefono) {
          val = row.telefono;
        }
        return String(val ?? '');
      }).join('\t') + '\n';
    });

    return navigator.clipboard.writeText(tsvContent);
  }).then(() => {
    showToast(`¡Copiado con éxito! ${rows.length.toLocaleString()} registros listos para pegar en Excel.`, 'success');
  }).catch(() => {
    showToast('No se pudo copiar automáticamente. Por favor use el botón Exportar a Excel.', 'error');
  });
}

// ==========================================
// CARGA DINÁMICA DE OTROS ARCHIVOS EXCEL
// ==========================================
function setupExcelLoader() {
  if (elements.btnLoadExcel && elements.excelFileInput) {
    elements.btnLoadExcel.addEventListener('click', () => {
      elements.excelFileInput.click();
    });

    elements.excelFileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        processUploadedExcel(file);
      }
      e.target.value = '';
    });
  }


  // Drag and drop sobre toda la ventana para archivos Excel externos
  if (elements.dropOverlay) {
    window.addEventListener('dragover', (e) => {
      // Si estamos arrastrando una columna de la tabla internamente, NO mostrar el overlay de Excel
      if (isDraggingCol || draggedColKey) {
        return;
      }
      // Verificar si lo que se arrastra es realmente un archivo externo de la computadora
      const isFile = e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
      if (!isFile) {
        return;
      }
      e.preventDefault();
      elements.dropOverlay.classList.add('active');
    });

    window.addEventListener('dragleave', (e) => {
      if (e.relatedTarget === null || e.clientX <= 0 || e.clientY <= 0) {
        elements.dropOverlay.classList.remove('active');
      }
    });

    window.addEventListener('drop', (e) => {
      elements.dropOverlay.classList.remove('active');
      if (isDraggingCol || draggedColKey) {
        return; // Evento de soltar columna gestionado por th
      }
      e.preventDefault();
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        const file = e.dataTransfer.files[0];
        if (file.name.match(/\.(xlsx|xls|csv)$/i)) {
          processUploadedExcel(file);
        } else {
          showToast('Por favor arrastre un archivo Excel (.xlsx, .xls) o CSV', 'error');
        }
      }
    });
  }
}

// ==========================================
// FUNCIONES DE CONTROL DE LA PANTALLA DE CARGA (LOADING MODAL)
// ==========================================
function formatBytes(bytes) {
  if (bytes === 0 || !bytes) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Configuración de iconos SVG e iluminación aura para cada tipo de acción
function getActionIconConfig(type) {
  switch (type) {
    case 'calendar':
      return {
        aura: 'radial-gradient(circle, rgba(245, 158, 11, 0.45) 0%, rgba(217, 119, 6, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(245, 158, 11, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="8" width="36" height="34" rx="8" fill="#f59e0b"/>
            <rect x="6" y="8" width="36" height="11" rx="8" fill="#d97706"/>
            <line x1="16" y1="4" x2="16" y2="12" stroke="white" stroke-width="3.5" stroke-linecap="round"/>
            <line x1="32" y1="4" x2="32" y2="12" stroke="white" stroke-width="3.5" stroke-linecap="round"/>
            <circle cx="16" cy="26" r="2.5" fill="white"/>
            <circle cx="24" cy="26" r="2.5" fill="white"/>
            <circle cx="32" cy="26" r="2.5" fill="white"/>
            <circle cx="16" cy="34" r="2.5" fill="white"/>
            <circle cx="24" cy="34" r="2.5" fill="white"/>
            <circle cx="32" cy="34" r="2.5" fill="white"/>
          </svg>
        `
      };
    case 'sql':
    case 'filter':
      return {
        aura: 'radial-gradient(circle, rgba(59, 130, 246, 0.45) 0%, rgba(99, 102, 241, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(59, 130, 246, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="6" width="36" height="36" rx="10" fill="#3b82f6"/>
            <polygon points="26,10 14,26 23,26 21,38 34,22 25,22" fill="white" stroke="white" stroke-width="1.5" stroke-linejoin="round"/>
          </svg>
        `
      };
    case 'export':
      return {
        aura: 'radial-gradient(circle, rgba(16, 185, 129, 0.45) 0%, rgba(6, 182, 212, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(16, 185, 129, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="6" width="36" height="36" rx="10" fill="#10b981"/>
            <path d="M24 14V29M24 29L17 22M24 29L31 22" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
            <line x1="14" y1="35" x2="34" y2="35" stroke="white" stroke-width="3.5" stroke-linecap="round"/>
          </svg>
        `
      };
    case 'copy':
      return {
        aura: 'radial-gradient(circle, rgba(6, 182, 212, 0.45) 0%, rgba(59, 130, 246, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(6, 182, 212, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="14" y="14" width="26" height="26" rx="8" fill="#06b6d4"/>
            <path d="M12 32H10A4 4 0 0 1 6 28V12A4 4 0 0 1 10 8H26A4 4 0 0 1 30 12V14" stroke="#0891b2" stroke-width="3.5" stroke-linecap="round"/>
            <path d="M21 27L25 31L33 22" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        `
      };
    case 'clear':
      return {
        aura: 'radial-gradient(circle, rgba(239, 68, 68, 0.45) 0%, rgba(244, 63, 94, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(239, 68, 68, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="6" width="36" height="36" rx="10" fill="#ef4444"/>
            <path d="M16 18H32M20 18V31A2 2 0 0 0 22 33H26A2 2 0 0 0 28 31V18M21 14H27" stroke="white" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        `
      };
    case 'columns':
      return {
        aura: 'radial-gradient(circle, rgba(99, 102, 241, 0.45) 0%, rgba(139, 92, 246, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(99, 102, 241, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="6" width="36" height="36" rx="10" fill="#6366f1"/>
            <line x1="18" y1="12" x2="18" y2="36" stroke="white" stroke-width="3" stroke-linecap="round"/>
            <line x1="30" y1="12" x2="30" y2="36" stroke="white" stroke-width="3" stroke-linecap="round"/>
            <line x1="10" y1="20" x2="38" y2="20" stroke="white" stroke-width="2.5"/>
          </svg>
        `
      };
    case 'reset':
      return {
        aura: 'radial-gradient(circle, rgba(139, 92, 246, 0.45) 0%, rgba(168, 85, 247, 0.2) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(139, 92, 246, 0.35))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <circle cx="24" cy="24" r="18" fill="#8b5cf6"/>
            <path d="M16 20A9 9 0 0 1 31 16L33 14M33 14V19M33 14H28" stroke="white" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M32 28A9 9 0 0 1 17 32L15 34M15 34V29M15 34H20" stroke="white" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        `
      };
    default:
      return {
        aura: 'radial-gradient(circle, rgba(16, 185, 129, 0.35) 0%, rgba(6, 182, 212, 0.15) 70%, transparent 100%)',
        filter: 'drop-shadow(0 8px 16px rgba(16, 185, 129, 0.3))',
        svg: `
          <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
            <rect x="6" y="4" width="36" height="40" rx="8" fill="#10b981"/>
            <path d="M14 18L22 30M22 18L14 30" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M28 18H34M28 24H34M28 30H34" stroke="white" stroke-width="2.5" stroke-linecap="round"/>
          </svg>
        `
      };
  }
}

/**
 * Ejecuta una operación con pantalla de carga estética, barra de progreso y animación
 * @param {Object} options Opciones de personalización
 * @param {Function} actionFn Función a ejecutar durante la carga
 * @returns {Promise}
 */
function runWithLoading(options, actionFn) {
  if (!elements.excelLoadingOverlay) {
    try {
      const res = actionFn ? actionFn() : null;
      return Promise.resolve(res);
    } catch (err) {
      return Promise.reject(err);
    }
  }

  const {
    title = 'Procesando información...',
    badgeText = null,
    badgeIcon = '⚡',
    badgeSize = null,
    iconType = 'default',
    initialStatus = 'Iniciando proceso...',
    initialSubstatus = 'Preparando entorno...',
    successStatus = '¡Operación completada!',
    successSubstatus = 'Datos procesados exitosamente',
    holdSuccessMs = 400
  } = options || {};

  // Configurar títulos y badges
  if (elements.loadingTitle) elements.loadingTitle.textContent = title;

  if (elements.loadingFileBadge) {
    if (badgeText) {
      elements.loadingFileBadge.style.display = 'inline-flex';
      const fileIconEl = elements.loadingFileBadge.querySelector('.file-icon');
      if (fileIconEl) fileIconEl.textContent = badgeIcon;
      if (elements.loadingFileName) elements.loadingFileName.textContent = badgeText;
      if (elements.loadingFileSize) {
        if (badgeSize) {
          elements.loadingFileSize.style.display = 'inline-block';
          elements.loadingFileSize.textContent = badgeSize;
        } else {
          elements.loadingFileSize.style.display = 'none';
        }
      }
    } else {
      elements.loadingFileBadge.style.display = 'none';
    }
  }

  // Configurar icono y aura
  const iconCfg = getActionIconConfig(iconType);
  const pulseEl = elements.excelLoadingOverlay.querySelector('.excel-icon-pulse');
  if (pulseEl) pulseEl.style.background = iconCfg.aura;
  if (elements.loadingIconMain) {
    elements.loadingIconMain.className = 'excel-loading-icon';
    elements.loadingIconMain.style.filter = iconCfg.filter;
    elements.loadingIconMain.innerHTML = iconCfg.svg;
  }

  // Reset de la barra de progreso
  const card = elements.excelLoadingOverlay.querySelector('.excel-loading-card');
  if (card) card.classList.remove('has-error');
  if (elements.loadingProgressBar) {
    elements.loadingProgressBar.classList.remove('completed');
    elements.loadingProgressBar.style.width = '0%';
  }
  if (elements.loadingPercent) elements.loadingPercent.textContent = '0%';
  if (elements.loadingStatus) elements.loadingStatus.textContent = initialStatus;
  if (elements.loadingSubstatus) elements.loadingSubstatus.textContent = initialSubstatus;

  // Mostrar modal con transición CSS
  elements.excelLoadingOverlay.classList.add('active');
  elements.excelLoadingOverlay.setAttribute('aria-hidden', 'false');

  return new Promise((resolve, reject) => {
    // Primer paso visible para que el usuario aprecie el inicio
    setTimeout(() => {
      updateExcelProgress(35, initialStatus, initialSubstatus);

      setTimeout(() => {
        let actionResult = null;
        try {
          if (typeof actionFn === 'function') {
            actionResult = actionFn();
          }
        } catch (err) {
          console.error('Error durante la ejecución:', err);
          errorExcelLoadingModal(err.message || 'Error en la operación.');
          reject(err);
          return;
        }

        // Si la función retorna una Promesa
        if (actionResult && typeof actionResult.then === 'function') {
          updateExcelProgress(70, 'Procesando respuesta asíncrona...', 'Sincronizando datos...');
          actionResult.then((asyncRes) => {
            finishAction(asyncRes);
          }).catch((err) => {
            errorExcelLoadingModal(err.message || 'Error en la operación.');
            reject(err);
          });
        } else {
          updateExcelProgress(75, 'Aplicando cambios en la tabla...', 'Actualizando métricas y filas...');
          setTimeout(() => {
            finishAction(actionResult);
          }, 150);
        }

        function finishAction(res) {
          const finalSub = (typeof successSubstatus === 'function') ? successSubstatus(res) : successSubstatus;
          updateExcelProgress(100, successStatus, finalSub);
          
          if (elements.loadingProgressBar) {
            elements.loadingProgressBar.classList.add('completed');
          }

          if (elements.loadingIconMain) {
            elements.loadingIconMain.className = 'excel-loading-icon success-state';
            elements.loadingIconMain.style.filter = 'drop-shadow(0 8px 16px rgba(16, 185, 129, 0.45))';
            elements.loadingIconMain.innerHTML = `
              <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
                <circle cx="24" cy="24" r="20" fill="#10b981"/>
                <path d="M14 24L21 31L34 17" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            `;
          }

          setTimeout(() => {
            hideExcelLoadingModal();
            resolve(res);
          }, holdSuccessMs);
        }
      }, 140);
    }, 60);
  });
}

function showExcelLoadingModal(fileName, fileSize) {
  if (!elements.excelLoadingOverlay) return;
  
  if (elements.loadingFileBadge) elements.loadingFileBadge.style.display = 'inline-flex';
  if (elements.loadingFileName) elements.loadingFileName.textContent = fileName || 'Archivo Excel';
  if (elements.loadingFileSize) {
    elements.loadingFileSize.style.display = 'inline-block';
    elements.loadingFileSize.textContent = formatBytes(fileSize);
  }
  if (elements.loadingTitle) elements.loadingTitle.textContent = 'Procesando archivo Excel...';
  
  // Icono SVG original de Excel
  const pulseEl = elements.excelLoadingOverlay.querySelector('.excel-icon-pulse');
  if (pulseEl) pulseEl.style.background = 'radial-gradient(circle, rgba(16, 185, 129, 0.35) 0%, rgba(6, 182, 212, 0.15) 70%, transparent 100%)';

  if (elements.loadingIconMain) {
    elements.loadingIconMain.className = 'excel-loading-icon';
    elements.loadingIconMain.style.filter = 'drop-shadow(0 8px 16px rgba(16, 185, 129, 0.3))';
    elements.loadingIconMain.innerHTML = `
      <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
        <rect x="6" y="4" width="36" height="40" rx="8" fill="#10b981"/>
        <path d="M14 18L22 30M22 18L14 30" stroke="white" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
        <path d="M28 18H34M28 24H34M28 30H34" stroke="white" stroke-width="2.5" stroke-linecap="round"/>
      </svg>
    `;
  }

  // Quitar estados de error o completado previo
  const card = elements.excelLoadingOverlay.querySelector('.excel-loading-card');
  if (card) card.classList.remove('has-error');
  if (elements.loadingProgressBar) {
    elements.loadingProgressBar.classList.remove('completed');
    elements.loadingProgressBar.style.width = '0%';
  }
  if (elements.loadingPercent) elements.loadingPercent.textContent = '0%';
  if (elements.loadingStatus) elements.loadingStatus.textContent = 'Preparando lectura...';
  if (elements.loadingSubstatus) elements.loadingSubstatus.textContent = 'Iniciando procesamiento del archivo...';

  elements.excelLoadingOverlay.classList.add('active');
  elements.excelLoadingOverlay.setAttribute('aria-hidden', 'false');
}

function updateExcelProgress(percent, statusText, substatusText) {
  const safePercent = Math.min(Math.max(percent, 0), 100);
  if (elements.loadingProgressBar) {
    elements.loadingProgressBar.style.width = `${safePercent}%`;
  }
  if (elements.loadingPercent) {
    elements.loadingPercent.textContent = `${Math.round(safePercent)}%`;
  }
  if (statusText && elements.loadingStatus) {
    elements.loadingStatus.textContent = statusText;
  }
  if (substatusText && elements.loadingSubstatus) {
    elements.loadingSubstatus.textContent = substatusText;
  }
}

function finishExcelProgress(totalRows, callback) {
  // 1. CARGAR TODA LA BARRA AL 100%
  updateExcelProgress(100, '¡Carga completada al 100%!', `${totalRows.toLocaleString()} registros procesados exitosamente`);
  
  if (elements.loadingProgressBar) {
    elements.loadingProgressBar.classList.add('completed');
  }

  // Cambiar icono a checkmark verde con animación de éxito
  if (elements.loadingIconMain) {
    elements.loadingIconMain.className = 'excel-loading-icon success-state';
    elements.loadingIconMain.innerHTML = `
      <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
        <circle cx="24" cy="24" r="20" fill="#10b981"/>
        <path d="M14 24L21 31L34 17" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    `;
  }

  // Mantener la barra totalmente llena al 100% para que el usuario la aprecie
  setTimeout(() => {
    hideExcelLoadingModal();
    if (typeof callback === 'function') {
      callback();
    }
  }, 750);
}

function hideExcelLoadingModal() {
  if (!elements.excelLoadingOverlay) return;
  elements.excelLoadingOverlay.classList.remove('active');
  elements.excelLoadingOverlay.setAttribute('aria-hidden', 'true');
}

function errorExcelLoadingModal(errorMessage) {
  if (!elements.excelLoadingOverlay) return;
  const card = elements.excelLoadingOverlay.querySelector('.excel-loading-card');
  if (card) card.classList.add('has-error');

  if (elements.loadingProgressBar) {
    elements.loadingProgressBar.style.width = '100%';
  }
  if (elements.loadingPercent) {
    elements.loadingPercent.textContent = 'Error';
  }
  if (elements.loadingStatus) {
    elements.loadingStatus.textContent = 'Error al procesar';
  }
  if (elements.loadingSubstatus) {
    elements.loadingSubstatus.textContent = errorMessage || 'No se pudo leer el archivo Excel.';
  }

  if (elements.loadingIconMain) {
    elements.loadingIconMain.className = 'excel-loading-icon';
    elements.loadingIconMain.innerHTML = `
      <svg viewBox="0 0 48 48" fill="none" class="excel-svg" xmlns="http://www.w3.org/2000/svg">
        <circle cx="24" cy="24" r="20" fill="#ef4444"/>
        <path d="M16 16L32 32M32 16L16 32" stroke="white" stroke-width="4" stroke-linecap="round"/>
      </svg>
    `;
  }

  setTimeout(() => {
    hideExcelLoadingModal();
  }, 3000);
}

// Función modular para normalizar una fila de Excel y precalcular campos SAR e índice de búsqueda
function normalizeSingleExcelRow(row, ref) {
  const norm = {};
  for (const [key, val] of Object.entries(row)) {
    let strVal = val;
    const isHoraKey = /hora/i.test(key);
    const isFechaKey = /fecha/i.test(key);

    if (val instanceof Date) {
      const y = val.getFullYear();
      const isBaseYear = y <= 1900;
      const hh = String(val.getHours()).padStart(2, '0');
      const mm = String(val.getMinutes()).padStart(2, '0');
      const ss = String(val.getSeconds()).padStart(2, '0');

      if (isHoraKey || isBaseYear) {
        strVal = `${hh}:${mm}:${ss}`;
      } else {
        const d = String(val.getDate()).padStart(2, '0');
        const m = String(val.getMonth() + 1).padStart(2, '0');
        if (val.getHours() > 0 || val.getMinutes() > 0 || val.getSeconds() > 0) {
          strVal = `${d}/${m}/${y} ${hh}:${mm}:${ss}`;
        } else {
          strVal = `${d}/${m}/${y}`;
        }
      }
    } else if (typeof val === 'number') {
      if (isHoraKey && val >= 0 && val < 1) {
        const totalSecs = Math.round(val * 86400);
        const hh = String(Math.floor(totalSecs / 3600)).padStart(2, '0');
        const mm = String(Math.floor((totalSecs % 3600) / 60)).padStart(2, '0');
        const ss = String(totalSecs % 60).padStart(2, '0');
        strVal = `${hh}:${mm}:${ss}`;
      } else if (isFechaKey && val > 30000) {
        const excelEpoch = new Date(1899, 11, 30);
        const dObj = new Date(excelEpoch.getTime() + val * 86400000);
        const d = String(dObj.getDate()).padStart(2, '0');
        const m = String(dObj.getMonth() + 1).padStart(2, '0');
        const y = dObj.getFullYear();
        const hh = String(dObj.getHours()).padStart(2, '0');
        const mm = String(dObj.getMinutes()).padStart(2, '0');
        const ss = String(dObj.getSeconds()).padStart(2, '0');
        if (dObj.getHours() > 0 || dObj.getMinutes() > 0 || dObj.getSeconds() > 0) {
          strVal = `${d}/${m}/${y} ${hh}:${mm}:${ss}`;
        } else {
          strVal = `${d}/${m}/${y}`;
        }
      }
    } else if (typeof val === 'string') {
      if (isHoraKey && val.includes('1899')) {
        const tm = val.match(/\d{2}:\d{2}:\d{2}/) || val.match(/\d{2}:\d{2}/);
        if (tm) strVal = tm[0].length === 5 ? tm[0] + ':00' : tm[0];
      }
    } else if (val === null || val === undefined) {
      strVal = '-';
    }
    norm[key] = strVal;
  }

  ensurePhoneExtracted(norm);
  norm.primer_nombre = resolveNombreAfiliado(norm);
  norm.correo_afiliado = resolveCorreoAfiliado(norm);

  norm._origFechaRegistro = norm['Fecha Registro'] || norm.fecha_registro || '';
  norm._origMesRegistro = norm['Mes Registro'] || norm.mes_registro || '';
  norm._origHoraRegistro = norm['Hora Registro'] || norm.hora_registro || '';
  norm._origDateYMD = getRowDateYMD(norm);

  const serv = String(norm['Servicio'] || norm.servicio || '').trim();
  if (ref.tiemposAtencion && ref.tiemposAtencion.hasOwnProperty(serv)) {
    norm.tiempo_atencion = ref.tiemposAtencion[serv];
    norm['Tiempo Atención'] = ref.tiemposAtencion[serv];
  } else {
    norm.tiempo_atencion = 15;
    norm['Tiempo Atención'] = 15;
  }

  const entidadReg = String(norm['Entidades Reguladoras'] || '').trim();
  norm.regulador = (ref.reguladores && ref.reguladores[entidadReg]) || '';
  norm.rechazado = (ref.serviciosRechazados && ref.serviciosRechazados.includes(serv)) ? 'S' : 'No';

  norm.cumple_sar = matchesSQLFilter(norm) ? 'S' : 'No';
  norm._origCumpleSar = norm.cumple_sar;

  // Precomputar índice para búsqueda global ultrarrápida
  norm._searchStr = buildSearchStr(norm);

  return norm;
}

function processUploadedExcel(file) {
  if (typeof XLSX === 'undefined') {
    showToast('Librería de Excel no disponible en el navegador', 'error');
    return;
  }

  // Mostrar modal de carga inmediatamente
  showExcelLoadingModal(file.name, file.size);
  updateExcelProgress(15, 'Leyendo archivo...', 'Cargando archivo en memoria del navegador...');

  const reader = new FileReader();

  reader.onprogress = function(e) {
    if (e.lengthComputable && e.total > 0) {
      const readPercent = 15 + Math.round((e.loaded / e.total) * 20);
      updateExcelProgress(readPercent, 'Leyendo archivo...', `${formatBytes(e.loaded)} de ${formatBytes(e.total)}`);
    }
  };

  reader.onload = function(e) {
    updateExcelProgress(35, 'Analizando libro Excel con SheetJS...', 'Leyendo hojas y celdas con modo rápido...');

    setTimeout(() => {
      try {
        const data = new Uint8Array(e.target.result);
        // Banderas de lectura de alto rendimiento:
        // dense: true (usa arrays 2D compactos en vez de miles de objetos de celda)
        // cellFormula: false, cellHTML: false, cellText: false (acelera la lectura hasta un 50%)
        const workbook = XLSX.read(data, {
          type: 'array',
          cellDates: true,
          dense: true,
          cellFormula: false,
          cellHTML: false,
          cellText: false
        });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        updateExcelProgress(55, 'Extrayendo filas y columnas...', `Hoja seleccionada: "${firstSheetName}"`);

        setTimeout(() => {
          try {
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (!jsonData || jsonData.length === 0) {
              errorExcelLoadingModal('El archivo Excel está vacío o no contiene filas con datos.');
              showToast('El archivo Excel está vacío o no contiene filas con datos.', 'error');
              return;
            }

            // Guardar respaldo de los datos iniciales
            if (!state.originalData) {
              state.originalData = [...state.rawData];
              state.originalColumns = [...COLUMNS_DEF];
            }

            // Extraer todas las columnas presentes en el Excel de forma rápida
            const allHeaders = new Set();
            const sampleLimit = Math.min(jsonData.length, 300);
            for (let r = 0; r < sampleLimit; r++) {
              Object.keys(jsonData[r]).forEach(k => allHeaders.add(k));
            }
            if (jsonData.length > 300) {
              for (let r = 300; r < jsonData.length; r += 100) {
                Object.keys(jsonData[r]).forEach(k => allHeaders.add(k));
              }
            }

            const totalRows = jsonData.length;
            const normalizedData = [];
            const ref = getTablasReferencia();
            const CHUNK_SIZE = 2500;
            let currentIdx = 0;

            function processNextChunk() {
              try {
                const nextLimit = Math.min(currentIdx + CHUNK_SIZE, totalRows);
                for (let i = currentIdx; i < nextLimit; i++) {
                  normalizedData.push(normalizeSingleExcelRow(jsonData[i], ref));
                }
                currentIdx = nextLimit;

                const progressPct = 65 + Math.round((currentIdx / totalRows) * 23); // 65% a 88%
                updateExcelProgress(progressPct, 'Normalizando registros y fechas...', `Procesando fila ${currentIdx.toLocaleString()} de ${totalRows.toLocaleString()}...`);

                if (currentIdx < totalRows) {
                  setTimeout(processNextChunk, 0);
                } else {
                  finalizeNormalizedData();
                }
              } catch (chunkErr) {
                console.error(chunkErr);
                errorExcelLoadingModal('Error al normalizar registros: ' + chunkErr.message);
                showToast('Error al normalizar registros: ' + chunkErr.message, 'error');
              }
            }

            function finalizeNormalizedData() {
              updateExcelProgress(90, 'Generando cabeceras y renderizando vista...', 'Configurando columnas y tablas...');

              setTimeout(() => {
                try {
                  allHeaders.add('Tiempo Atención');

                  const headersArray = Array.from(allHeaders);
                  const newColDefs = headersArray.map(header => {
                    return {
                      key: header,
                      label: header,
                      defaultVisible: true,
                      isCore: isCoreSQLCol(header, header),
                      isNumber: typeof jsonData[0]?.[header] === 'number' || /tiempo/i.test(header)
                    };
                  });

                  currentColumns = newColDefs;
                  state.rawData = normalizedData;
                  state.visibleColumns = new Set(newColDefs.map(c => c.key));
                  state.columnFilters = {};
                  state.quickFilters = {};
                  state.globalSearch = '';
                  if (elements.globalSearchInput) elements.globalSearchInput.value = '';

                  if (elements.currentFileBadge) {
                    elements.currentFileBadge.textContent = file.name;
                  }
                  if (elements.lastUpdatedBadge) {
                    elements.lastUpdatedBadge.textContent = 'Actualizado: ' + new Date().toLocaleString('es-PE');
                  }

                  saveStateToIDB(file.name, normalizedData, newColDefs);
                  try { localStorage.removeItem('SAR_UIState'); } catch (e) {}

                  renderHeaderAndQuickFilters();
                  applyFiltersAndSort();

                  finishExcelProgress(normalizedData.length, () => {
                    showToast(`¡Excel cargado! ${normalizedData.length.toLocaleString()} registros cargados sin filtros`, 'success');
                  });
                } catch (errRender) {
                  console.error(errRender);
                  errorExcelLoadingModal('Error al renderizar los datos: ' + errRender.message);
                  showToast('Error al renderizar los datos: ' + errRender.message, 'error');
                }
              }, 40);
            }

            updateExcelProgress(65, 'Normalizando registros y fechas...', `Iniciando ${totalRows.toLocaleString()} registros...`);
            setTimeout(processNextChunk, 0);

          } catch (errJson) {
            console.error(errJson);
            errorExcelLoadingModal('Error al procesar hoja: ' + errJson.message);
            showToast('Error al procesar la hoja de Excel: ' + errJson.message, 'error');
          }
        }, 40);

      } catch (errXlsx) {
        console.error(errXlsx);
        errorExcelLoadingModal('Error al interpretar el libro Excel: ' + errXlsx.message);
        showToast('Error al leer el archivo Excel: ' + errXlsx.message, 'error');
      }
    }, 40);
  };

  reader.onerror = function(err) {
    console.error(err);
    errorExcelLoadingModal('Error al leer el archivo físico.');
    showToast('Error al leer el archivo seleccionado', 'error');
  };

  reader.readAsArrayBuffer(file);
}

function showToast(message, type = 'info') {
  if (!elements.toastContainer) return;
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span>${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span>
    <span>${message}</span>
  `;
  elements.toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// ==========================================
// TEMA CLARO / OSCURO
// ==========================================
function initTheme() {
  const saved = localStorage.getItem('sar_theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
  updateThemeIcon(saved);
}

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || 'light';
  const next = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('sar_theme', next);
  updateThemeIcon(next);
}

function updateThemeIcon(theme) {
  elements.themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
}
