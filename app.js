// Estado principal de la aplicación.
// Aquí se guarda el archivo original, sus filas, la muestra generada y datos auxiliares.
let originalFile = null;
let originalFileName = "";
let originalRows = [];
let sampleRows = [];
let headers = [];
let reportDate = "";
let reportAuditor = "";
let workbookRef = null;
let currentSheetName = "";
let activeLicense = null;

const LICENSE_STORAGE_KEY = 'rsm_muestras_licencia_v1';

// Referencias a elementos del DOM para manipular la interfaz desde JavaScript.
const licenseGate = document.getElementById('licenseGate');
const appShell = document.getElementById('appShell');
const licenseFileInput = document.getElementById('licenseFileInput');
const activateLicenseBtn = document.getElementById('activateLicenseBtn');
const licenseMessage = document.getElementById('licenseMessage');
const licenseDetails = document.getElementById('licenseDetails');
const licenseStatusCard = document.getElementById('licenseStatusCard');
const licenseStatus = document.getElementById('licenseStatus');
const fileInput = document.getElementById('fileInput');
const sampleCountInput = document.getElementById('sampleCount');
const useSeedCheckbox = document.getElementById('useSeedCheckbox');
const seedInput = document.getElementById('seedInput');
const auditorInput = document.getElementById('auditorInput');
const sheetSelect = document.getElementById('sheetSelect');
const sheetSelectorWrap = document.getElementById('sheetSelectorWrap');
const generateBtn = document.getElementById('generateBtn');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');
const downloadPdfBtn = document.getElementById('downloadPdfBtn');
const resetBtn = document.getElementById('resetBtn');
const fileInfo = document.getElementById('fileInfo');
const message = document.getElementById('message');
const resultCard = document.getElementById('resultCard');
const resultTable = document.getElementById('resultTable');
const reportMeta = document.getElementById('reportMeta');

// Logo embebido para asegurar que el PDF siempre pueda renderizar la marca.
const EMBEDDED_LOGO_DATA_URL = 'data:image/png;base64,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lkNZzu3K3dZe9bPqbdltuX/yB3TbGFiOZNmzyPj9VT8r3pa/vY+0llvnQXwsB7PsGP398aE/7JsYs3ITRWu5cR7Ex3Iwy519kfHnqy5WzZnvCN/X6mrLk2aCfcNydMuO0d/fr7qYVhr5XyE8eC2Pzgk2loNZto/+/u9VP4fSi+Uv6Q+7X7nIXjPBHrAc3nKvXLX+4lWVmgWUF7flzncQH8vRLNtHf/96Vc3C+SNP+fsLqy03voP4WA5nebIuMr6/6udbmR2M5w9jesuzb4KN5WiWzaO/Z6+6CDhLbsrPXldveXQdxMdyOMvm0d+vXvV6XvHcCiU/eRjTW07yCfYJywewnNuVWyyvev08v33D2L59zFLJT15Wb7lzHcTHcjzL1tFfyoO8/t75z67XL1WPApZb3wQby+EsW0d/6atmtwKW//qWYzyIj+V4lq2jv7qWTyUsTz+/LbTbGFgOZ9k4+qtruSlhOXkO4mM5oGXj6K+q5UcRy73nID6WA1o2jv7Ststli+XWNcHGckDLttFf2na5bLHcuCbYWA5oOelhVba8NGUsT1/WHyWH5ZiWbaO/mpbvhSyPjoP4WI5o2Tb6q2n5Ushy7ziIj+WQlk2jv5qWm0KWO88EG8sRLZtGf2nbJYbJcvNl/VFyWA5qObcr99jY8qWY5dkxwcZySMuW0V/adolhszzaD+JjOaZly+gvbbvEsFlOjgk2lkNatoz+6lk+lbPc2Q/iYzmoZcPor5rlpSlnubX+KDksh7VsGP1Vs3wpaLm55i93wfKhLBtGf2nb27LR8jn/eHvH8qEsG0Z/tSzfi1oWfII3LB/Lsn6RUcvyUMey4dEPyzEt60d/advbstNyp3/zYDmoZf3oL217W3Za7i0fjeWYlketsbTtbdlpWX8QH8thLatHf1UsL0Mly4YJNpajWm61i4wqli9NJcuz5SWxHNOyevRXw/LSVLJsmWBjOaxl7a5cDcunWpY70wdjOajlTvlUVsHyvallOZm+FWA5qGXt6C9tusJwWjYcxMdyYMujbgGQNl1hOC2btjGwHNaycvRX3PKlqWfZ9qpYjmpZOfpLWy6WnZb1P0oOy7Et53blhpqWH01Fy73tY7Ec1rJu9FfW8tLUtGzbxsByXMu60V9Ry8tQ1fJkM4nlsJZzu3Jf1SwLKLssWw7iYzm05aRZZHTzppRdlo37gFiOa1k5+nuft1orOy2bDuJjObRl7ehP/psnXTsYbsu98W2E5cCWlaO/Ipr/21S3PBrfR1gObDm3K/fkScmpeTk19S3bJthYDm25Ma1sz/Z186f80hyWrXNzLEe2rBv9/e8p8GqTPDRbWLYdxMdycMuqXbmfvOhvzvLlhdNyZ/1ILG9e+/brtH/X23prN1IdZ9U9OXNl9n8f8wcOsDv6u+r9Q+T5+oN/KwrheXX1vHz+4KZGkZY95/PH9d+mr9czjin8ah7DRERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERES0Tb8BhdLuTF1oImIAAAAASUVORK5CYII=';

// Revisa si las librerías locales realmente cargaron en el navegador.
// Si alguna falta, la app se bloquea y se muestra un error claro.
function showInlineMessage(target, text, isError = false) {
  target.textContent = text;
  target.classList.add('show');
  target.style.borderColor = isError ? '#7f1d1d' : '#b8bec5';
  target.style.background = isError ? '#fee2e2' : '#f8fafc';
  target.style.color = isError ? '#7f1d1d' : '#1f2937';
}

function clearInlineMessage(target) {
  target.textContent = '';
  target.classList.remove('show');
}

function getBogotaIsoDate(date = new Date()) {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(date);
}

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: 'same-origin',
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    },
    ...options
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.message || 'No se pudo completar la operación.');
  }

  return data;
}

function persistValidatedLicense(license) {
  localStorage.setItem(LICENSE_STORAGE_KEY, JSON.stringify(license));
}

function readStoredLicense() {
  const raw = localStorage.getItem(LICENSE_STORAGE_KEY);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (error) {
    localStorage.removeItem(LICENSE_STORAGE_KEY);
    return null;
  }
}

async function validateLicenseWithServer(raw) {
  const data = await apiFetch('/api/licenses/validate', {
    method: 'POST',
    body: JSON.stringify(raw)
  });

  return data.license;
}

function renderLicenseBadge(license) {
  licenseStatusCard.style.display = 'block';
  licenseStatus.innerHTML = `
    <span class="pill">Licencia activa</span>
    <span class="pill">Empresa: ${escapeHtml(license.company)}</span>
    <span class="pill">Titular: ${escapeHtml(license.issuedTo)}</span>
    <span class="pill">Vence: ${escapeHtml(license.expiresAt)}</span>
    <a
      class="admin-link admin-link-icon"
      href="/generador-licencias"
      target="_blank"
      rel="noopener noreferrer"
      title="Generador de licencias"
      aria-label="Abrir generador de licencias"
    >
      <span aria-hidden="true">+</span>
    </a>
  `;
}

function lockApplication() {
  activeLicense = null;
  licenseGate.style.display = 'flex';
  appShell.style.display = 'none';
  licenseStatusCard.style.display = 'none';
  setAppEnabled(false);
}

function unlockApplication(license) {
  activeLicense = license;
  licenseGate.style.display = 'none';
  appShell.style.display = 'block';
  renderLicenseBadge(license);
  setAppEnabled(!missingLibraries.length);
}

function getMissingLibraries() {
  const missing = [];

  // XLSX procesa archivos CSV/XLS/XLSX.
  if (!window.XLSX) missing.push('xlsx.full.min.js');
  // jsPDF genera el PDF final.
  if (!window.jspdf || !window.jspdf.jsPDF) missing.push('jspdf.umd.min.js');

  // autoTable añade a jsPDF la capacidad de dibujar tablas.
  const hasAutoTable =
    window.jspdf &&
    window.jspdf.jsPDF &&
    window.jspdf.jsPDF.API &&
    typeof window.jspdf.jsPDF.API.autoTable === 'function';

  if (!hasAutoTable) missing.push('jspdf.plugin.autotable.min.js');

  return missing;
}

// Activa o desactiva los controles según el estado general de la app.
function setAppEnabled(enabled) {
  fileInput.disabled = !enabled;
  sampleCountInput.disabled = !enabled;
  useSeedCheckbox.disabled = !enabled;
  seedInput.disabled = !enabled || !useSeedCheckbox.checked;
  sheetSelect.disabled = !enabled;
  generateBtn.disabled = !enabled;
  downloadExcelBtn.disabled = true;
  downloadPdfBtn.disabled = true;
  resetBtn.disabled = !enabled;
}

// Muestra u oculta el campo de semilla segun el check del usuario.
function syncSeedFieldState() {
  const useSeed = useSeedCheckbox.checked;
  seedInput.disabled = !useSeed;
  seedInput.style.display = useSeed ? 'block' : 'none';

  if (!useSeed) {
    seedInput.value = '';
  }
}

// Muestra un mensaje en pantalla.
// Si isError es true, usa colores de error; si no, usa colores neutros.
function showMessage(text, isError = false) {
  message.textContent = text;
  message.classList.add('show');
  message.style.borderColor = isError ? '#7f1d1d' : '#b8bec5';
  message.style.background = isError ? '#fee2e2' : '#f8fafc';
  message.style.color = isError ? '#7f1d1d' : '#1f2937';
}

// Limpia el cuadro de mensajes.
function clearMessage() {
  message.textContent = '';
  message.classList.remove('show');
}

// Validación inicial: si faltan archivos de la carpeta vendor, la app se deshabilita.
const missingLibraries = getMissingLibraries();
if (missingLibraries.length) {
  setAppEnabled(false);
  showMessage(
    `No se pudieron cargar las librer\u00edas locales requeridas: ${missingLibraries.join(', ')}. ` +
    'Verifica que la carpeta "vendor" est\u00e9 completa junto a este archivo.',
    true
  );
} else {
  setAppEnabled(false);
}

const storedLicense = readStoredLicense();
if (storedLicense) {
  validateLicenseWithServer(storedLicense)
    .then(validatedLicense => {
      persistValidatedLicense(validatedLicense);
      unlockApplication(validatedLicense);
    })
    .catch(() => {
      localStorage.removeItem(LICENSE_STORAGE_KEY);
      lockApplication();
      showInlineMessage(licenseMessage, 'Carga una licencia válida para habilitar la herramienta.');
    });
} else {
  lockApplication();
  showInlineMessage(licenseMessage, 'Carga una licencia válida para habilitar la herramienta.');
}

// Devuelve la fecha y hora actual en formato de Bogotá.
// Se usa para marcar cuándo se generó el reporte.
function formatBogotaDate(date = new Date()) {
  return new Intl.DateTimeFormat('es-CO', {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  }).format(date).replace(',', '');
}

// Convierte una semilla cualquiera en un número entero reproducible.
// Así podemos generar siempre la misma muestra si el usuario repite la semilla.
function hashSeed(value) {
  const str = String(value);
  let hash = 2166136261;
  for (let i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }
  return Math.abs(hash >>> 0);
}

// Generador pseudoaleatorio basado en semilla.
// Cada vez que se usa la misma semilla, produce la misma secuencia.
function createSeededRandom(seed) {
  let state = hashSeed(seed) || 1;
  return function () {
    state = (1664525 * state + 1013904223) % 4294967296;
    return state / 4294967296;
  };
}

// Mezcla un arreglo usando Fisher-Yates.
// Si hay semilla, usa el generador reproducible; si no, usa Math.random.
function shuffle(array, seed = null) {
  const arr = [...array];
  const randomFn = seed !== null && seed !== '' ? createSeededRandom(seed) : Math.random;

  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(randomFn() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// Extrae todos los encabezados detectados en las filas leídas.
// Esto permite soportar filas con columnas faltantes sin romper la tabla.
function extractHeaders(rows) {
  const headerSet = new Set();
  rows.forEach(row => {
    Object.keys(row).forEach(key => headerSet.add(key));
  });
  return Array.from(headerSet);
}

// Normaliza todas las filas:
// 1. Descarta filas totalmente vacías.
// 2. Asegura que cada fila tenga todas las columnas detectadas.
// 3. Limpia el valor de cada celda antes de guardarlo.
function normalizeRows(rows) {
  return rows
    .filter(row => Object.values(row).some(v => v !== null && v !== undefined && String(v).trim() !== ''))
    .map(row => {
      const cleaned = {};
      headers.forEach(h => {
        const value = row[h];
        cleaned[h] = normalizeCellValue(value, h);
      });
      return cleaned;
    });
}

// Normaliza una celda individual.
// Si Excel devuelve una fecha con hora 00:00:00, se muestra solo la fecha.
function normalizeCellValue(value, headerName = '') {
  if (value === undefined || value === null) return '';

  if (typeof value === 'string') {
    const trimmed = value.trim();
    const normalizedHeader = String(headerName).toLowerCase();
    const isDateColumn = normalizedHeader.includes('fecha');
    const shortDateTimeMatch = trimmed.match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})\s+\d{1,2}:\d{2}(?::\d{2})?$/);
    const midnightMatch = trimmed.match(/^(\d{2}\/\d{2}\/\d{4}) 00:00:00$/);

    if (isDateColumn && shortDateTimeMatch) {
      return shortDateTimeMatch[1];
    }

    if (midnightMatch) {
      return midnightMatch[1];
    }
  }

  return value;
}

// Intenta detectar el separador de un CSV revisando las primeras líneas.
// Se usa para diferenciar archivos separados por coma, punto y coma o tabulación.
function detectCsvSeparator(text) {
  const firstLines = text.split(/\r?\n/).slice(0, 5).join('\n');
  const commaCount = (firstLines.match(/,/g) || []).length;
  const semicolonCount = (firstLines.match(/;/g) || []).length;
  const tabCount = (firstLines.match(/\t/g) || []).length;

  if (semicolonCount > commaCount && semicolonCount >= tabCount) return ';';
  if (tabCount > commaCount && tabCount > semicolonCount) return '\t';
  return ',';
}

// Lee el archivo cargado según su extensión.
// Soporta CSV, XLSX y XLS.
async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();

  // Si es CSV, se lee como texto y se intenta detectar su separador.
  if (ext === 'csv') {
    const text = await file.text();
    const separator = detectCsvSeparator(text);
    const workbook = XLSX.read(text, {
      type: 'string',
      raw: false,
      FS: separator
    });
    workbookRef = workbook;
    currentSheetName = workbook.SheetNames[0];
    sheetSelectorWrap.style.display = 'none';
    return loadSheetData(workbook, currentSheetName);
  }

  // Si es Excel, se lee como binario y se listan sus hojas.
  if (ext === 'xlsx' || ext === 'xls') {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    workbookRef = workbook;

    // Se llenan las opciones del selector de hojas incluyendo una opción inicial.
    sheetSelect.innerHTML = [
      '<option value="">Selecciona una hoja</option>',
      ...workbook.SheetNames.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    ].join('');

    // Para archivos Excel siempre dejamos visible el selector,
    // así el usuario puede confirmar la hoja detectada.
    sheetSelectorWrap.style.display = 'block';

    if (workbook.SheetNames.length > 1) {
      currentSheetName = '';
      originalRows = [];
      headers = [];
      return [];
    }

    // Si solo hay una hoja, se selecciona automáticamente,
    // pero el selector sigue visible para que el usuario la vea.
    currentSheetName = workbook.SheetNames[0];
    sheetSelect.value = currentSheetName;
    return loadSheetData(workbook, currentSheetName);
  }

  throw new Error('El archivo no tiene un formato soportado. Usa un archivo CSV, XLSX o XLS.');
}

// Carga una hoja específica del libro y la transforma en filas utilizables.
function loadSheetData(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`No se pudo leer la hoja "${sheetName}". Intenta seleccionar otra hoja del archivo.`);
  }

  // raw:false pide a SheetJS que respete formatos visibles, útil para fechas.
  // dateNF define el formato preferido cuando la celda es fecha.
  const parsed = XLSX.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false,
    dateNF: 'dd/mm/yyyy hh:mm:ss'
  });

  if (!parsed.length) {
    throw new Error(`La hoja "${sheetName}" est\u00e1 vac\u00eda o no tiene filas utilizables. Elige otra hoja o revisa el archivo.`);
  }

  headers = extractHeaders(parsed);
  originalRows = normalizeRows(parsed);

  if (!headers.length) {
    throw new Error('No se detectaron columnas v\u00e1lidas en el archivo. Verifica que la primera fila tenga encabezados.');
  }

  if (!originalRows.length) {
    throw new Error('No se encontraron filas utilizables en el archivo. Revisa si el documento tiene datos completos.');
  }

  return originalRows;
}

// Muestra un resumen del archivo actual: nombre, hoja, filas y columnas.
function renderFileInfo(totalRows) {
  const sheetInfo = currentSheetName ? `<span class="pill">Hoja: ${escapeHtml(currentSheetName)}</span>` : '';
  fileInfo.innerHTML = `
    <span class="pill">Archivo: ${escapeHtml(originalFileName)}</span>
    ${sheetInfo}
    <span class="pill">Filas detectadas: ${totalRows}</span>
    <span class="pill">Columnas: ${headers.length}</span>
  `;
}

// Construye la tabla HTML con la muestra generada.
function renderTable() {
  if (!sampleRows.length) return;

  const cols = ['N\u00daMERO', ...headers];
  const thead = `<thead><tr>${cols.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>`;
  const tbody = `<tbody>${sampleRows.map(row => `
    <tr>${cols.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('')}</tr>
  `).join('')}</tbody>`;

  resultTable.innerHTML = thead + tbody;
  resultCard.style.display = 'block';
  const auditorMeta = reportAuditor ? ` | Auditor responsable: ${reportAuditor}` : '';
  reportMeta.textContent = `Reporte generado el: ${reportDate} | Informaci\u00f3n obtenida de: ${originalFileName}${currentSheetName ? ` | Hoja: ${currentSheetName}` : ''}${auditorMeta}`;
}

// Escapa texto para evitar que caracteres especiales se interpreten como HTML.
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Limpia un texto para poder usarlo como parte de un nombre de archivo.
function slugifyFilePart(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
}

// Quita la extension del archivo original para reutilizar solo su nombre base.
function getBaseFileName(fileName) {
  return String(fileName || '').replace(/\.[^.]+$/, '');
}

// Convierte la fecha del reporte a un formato compacto para descargas.
function getReportDateStamp() {
  if (!reportDate) {
    return formatBogotaDate()
      .replace(/\//g, '-')
      .replace(/:/g, '-')
      .replace(/\s+/g, '_');
  }

  return reportDate
    .replace(/\//g, '-')
    .replace(/:/g, '-')
    .replace(/\s+/g, '_');
}

// Construye nombres de archivo mas descriptivos para las descargas.
function buildDownloadFileName(extension) {
  const baseName = slugifyFilePart(getBaseFileName(originalFileName)) || 'muestra';
  const sheetName = currentSheetName ? `_${slugifyFilePart(currentSheetName)}` : '';
  const dateStamp = getReportDateStamp();
  return `${baseName}${sheetName}_${dateStamp}.${extension}`;
}

// Genera la muestra aleatoria tomando las primeras filas del arreglo ya mezclado.
// También agrega una columna enumerada llamada NÚMERO.
function buildSample(count) {
  const seedValue = useSeedCheckbox.checked ? seedInput.value.trim() : '';
  const randomized = shuffle(originalRows, seedValue === '' ? null : seedValue).slice(0, count);

  sampleRows = randomized.map((row, index) => ({
    'N\u00daMERO': index + 1,
    ...row
  }));

  reportDate = formatBogotaDate();
  renderTable();
  downloadExcelBtn.disabled = false;
  downloadPdfBtn.disabled = false;
}

// Descarga cualquier Blob creando un enlace temporal en memoria.
function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

// Dibuja una version vectorial simple del logo RSM para asegurar su presencia en el PDF.
function drawPdfLogo(doc, x, y) {
  doc.setFillColor(154, 158, 163);
  doc.rect(x, y, 2.2, 2.2, 'F');
  doc.setFillColor(44, 160, 28);
  doc.rect(x + 3, y, 4.2, 2.2, 'F');
  doc.setFillColor(0, 156, 222);
  doc.rect(x + 8.2, y, 11.5, 2.2, 'F');

  doc.setTextColor(102, 106, 112);
  doc.setFontSize(17);
  doc.setFont(undefined, 'bold');
  doc.text('RSM', x, y + 7.4);
  doc.setFont(undefined, 'normal');
}

// Convierte la muestra actual a un archivo Excel (.xlsx).
function exportSampleExcelBlob() {
  const sheetName = currentSheetName || 'Muestra';
  const sourceInfo = `Informaci\u00f3n obtenida de: ${originalFileName}${currentSheetName ? ` | Hoja: ${currentSheetName}` : ''}`;
  const auditorInfo = reportAuditor ? `Auditor responsable: ${reportAuditor}` : 'Auditor responsable: No registrado';
  const seedInfo = useSeedCheckbox.checked && seedInput.value.trim()
    ? `Semilla utilizada: ${seedInput.value.trim()}`
    : 'Semilla utilizada: Aleatoria';
  const titleRows = [
    ['RSM Colombia | Generador de Muestras RSM'],
    [`Reporte generado el: ${reportDate}`],
    [sourceInfo],
    [auditorInfo],
    [seedInfo],
    []
  ];

  const ws = XLSX.utils.aoa_to_sheet(titleRows);
  XLSX.utils.sheet_add_json(ws, sampleRows, { origin: 'A7', skipHeader: false });

  const cols = ['N\u00daMERO', ...headers];
  ws['!autofilter'] = {
    ref: `A7:${XLSX.utils.encode_col(cols.length - 1)}7`
  };

  ws['!cols'] = cols.map((col, index) => {
    const values = sampleRows.map(row => String(row[col] ?? ''));
    const maxLength = Math.max(
      String(col).length,
      ...values.map(value => value.length)
    );

    return {
      wch: Math.min(Math.max(maxLength + 2, index === 0 ? 10 : 14), 40)
    };
  });

  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(cols.length - 1, 0) } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(cols.length - 1, 0) } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: Math.max(cols.length - 1, 0) } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: Math.max(cols.length - 1, 0) } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: Math.max(cols.length - 1, 0) } }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  const xlsxBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob(
    [xlsxBuffer],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  );
}

// Convierte la muestra actual a PDF en orientación horizontal.
function exportSamplePdfBlob() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape' });
  const cols = ['N\u00daMERO', ...headers];
  const body = sampleRows.map(row => cols.map(col => String(row[col] ?? '')));
  const pageWidth = doc.internal.pageSize.getWidth();

  // Encabezado corporativo para que el PDF se identifique claramente como documento RSM.
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(10, 8, pageWidth - 20, 26, 4, 4, 'F');
  doc.setFillColor(154, 158, 163);
  doc.roundedRect(14, 12, 24, 0.7, 0.35, 0.35, 'F');
  doc.setFillColor(44, 160, 28);
  doc.roundedRect(44, 12, 78, 0.7, 0.35, 0.35, 'F');
  doc.setFillColor(0, 156, 222);
  doc.roundedRect(128, 12, pageWidth - 142, 0.7, 0.35, 0.35, 'F');

  drawPdfLogo(doc, 14, 16);

  doc.setTextColor(102, 106, 112);
  doc.setFontSize(9);
  doc.text('RSM COLOMBIA | MUESTREO DOCUMENTAL', 40, 20);

  doc.setTextColor(23, 48, 77);
  doc.setFontSize(18);
  doc.text('Generador de Muestras RSM', 40, 29);

  doc.setTextColor(31, 41, 55);
  doc.setFontSize(10);
  doc.text(`Reporte generado el: ${reportDate}`, 14, 40);
  doc.text(`Informaci\u00f3n obtenida de: ${originalFileName}${currentSheetName ? ` | Hoja: ${currentSheetName}` : ''}`, 14, 47);
  if (reportAuditor) {
    doc.text(`Auditor responsable: ${reportAuditor}`, 14, 54);
  }

  doc.autoTable({
    head: [cols],
    body,
    startY: reportAuditor ? 61 : 54,
    styles: { fontSize: 8, cellPadding: 2 },
    headStyles: { fillColor: [102, 106, 112] }
  });

  const finalY = doc.lastAutoTable ? doc.lastAutoTable.finalY : 54;
  doc.setDrawColor(224, 231, 235);
  doc.line(14, finalY + 6, pageWidth - 14, finalY + 6);
  doc.setTextColor(95, 103, 112);
  doc.setFontSize(8);
  doc.text('Documento generado desde la herramienta interna de RSM. Desarrollado por Yeison Castro.', 14, finalY + 12);

  return doc.output('blob');
}

// Restaura el estado interno de la aplicación.
async function activateLicenseFromFile() {
  clearInlineMessage(licenseMessage);

  const file = licenseFileInput.files[0];
  if (!file) {
    showInlineMessage(licenseMessage, 'Primero selecciona un archivo de licencia en formato JSON.', true);
    return;
  }

  try {
    const rawText = await file.text();
    const parsedLicense = JSON.parse(rawText);
    const validatedLicense = await validateLicenseWithServer(parsedLicense);

    persistValidatedLicense(validatedLicense);
    licenseDetails.style.display = 'flex';
    licenseDetails.innerHTML = `
      <span class="pill">Empresa: ${escapeHtml(validatedLicense.company)}</span>
      <span class="pill">Titular: ${escapeHtml(validatedLicense.issuedTo)}</span>
      <span class="pill">Vence: ${escapeHtml(validatedLicense.expiresAt)}</span>
    `;
    showInlineMessage(licenseMessage, 'Licencia validada correctamente. La aplicaci\u00f3n ha sido habilitada.');
    unlockApplication(validatedLicense);
  } catch (error) {
    licenseDetails.style.display = 'none';
    licenseDetails.innerHTML = '';
    showInlineMessage(licenseMessage, error.message || 'No se pudo validar la licencia.', true);
  }
}

function resetState() {
  originalFile = null;
  originalFileName = '';
  originalRows = [];
  sampleRows = [];
  headers = [];
  reportDate = '';
  reportAuditor = '';
  workbookRef = null;
  currentSheetName = '';
}

// Cuando el usuario selecciona un archivo:
// 1. Limpia resultados anteriores.
// 2. Lee el nuevo archivo.
// 3. Muestra información básica.
fileInput.addEventListener('change', async (e) => {
  clearMessage();
  resultCard.style.display = 'none';
  resultTable.innerHTML = '';
  sampleRows = [];
  downloadExcelBtn.disabled = true;
  downloadPdfBtn.disabled = true;

  const file = e.target.files[0];
  if (!file) return;

  try {
    // Se guarda referencia al archivo original para mantener su nombre y contexto.
    originalFile = file;
    originalFileName = file.name;

    await parseFile(file);

    if (workbookRef && workbookRef.SheetNames.length > 1 && !currentSheetName) {
      fileInfo.innerHTML = `
        <span class="pill">Archivo: ${escapeHtml(originalFileName)}</span>
        <span class="pill">Hojas disponibles: ${workbookRef.SheetNames.length}</span>
      `;
      showMessage(`Archivo cargado correctamente. Se detectaron ${workbookRef.SheetNames.length} hojas. Selecciona la hoja que quieres usar para generar la muestra.`);
      return;
    }

    renderFileInfo(originalRows.length);

    const extraMessage = currentSheetName ? ` Hoja cargada: ${currentSheetName}.` : '';
    showMessage(`Archivo cargado correctamente.${extraMessage} Ahora indica cu\u00e1ntas filas aleatorias deseas generar.`);
  } catch (error) {
    resetState();
    fileInfo.innerHTML = '';
    sheetSelect.innerHTML = '';
    sheetSelectorWrap.style.display = 'none';
    showMessage(error.message || 'No se pudo procesar el archivo.', true);
  }
});

// Si el archivo Excel tiene varias hojas, este evento cambia de hoja activa.
sheetSelect.addEventListener('change', () => {
  if (!workbookRef || !sheetSelect.value) return;

  try {
    clearMessage();
    resultCard.style.display = 'none';
    resultTable.innerHTML = '';
    sampleRows = [];
    downloadExcelBtn.disabled = true;
    downloadPdfBtn.disabled = true;

    currentSheetName = sheetSelect.value;
    loadSheetData(workbookRef, currentSheetName);
    renderFileInfo(originalRows.length);
    showMessage(`Hoja "${currentSheetName}" cargada correctamente. Ahora escribe cu\u00e1ntas filas aleatorias deseas incluir en la muestra.`);
  } catch (error) {
    showMessage(error.message || 'No se pudo cambiar de hoja. Intenta seleccionar otra hoja del archivo.', true);
  }
});

// Genera la muestra cuando el usuario hace clic en el botón principal.
// Antes valida que exista archivo, que el número sea entero y que no exceda las filas disponibles.
generateBtn.addEventListener('click', () => {
  clearMessage();

  if (workbookRef && workbookRef.SheetNames.length > 1 && !currentSheetName) {
    showMessage('Antes de generar la muestra, selecciona una hoja de Excel en el listado.', true);
    return;
  }

  const auditorName = auditorInput.value.trim();
  if (!auditorName) {
    showMessage('Indica el nombre del auditor responsable antes de generar la muestra.', true);
    auditorInput.focus();
    return;
  }

  if (!originalRows.length) {
    showMessage('Primero carga un archivo con datos v\u00e1lidos para poder generar la muestra.', true);
    return;
  }

  if (useSeedCheckbox.checked && !seedInput.value.trim()) {
    showMessage('Activaste el uso de semilla, pero a\u00fan no has escrito un valor.', true);
    seedInput.focus();
    return;
  }

  const count = Number(sampleCountInput.value);

  if (!Number.isInteger(count) || count <= 0) {
    showMessage('La cantidad de muestras debe ser un n\u00famero entero mayor que 0.', true);
    return;
  }

  if (count > originalRows.length) {
    showMessage(`Solicitaste ${count} filas, pero solo hay ${originalRows.length} disponibles. Ingresa un n\u00famero menor o igual a ${originalRows.length}.`, true);
    return;
  }

  reportAuditor = auditorName;
  buildSample(count);

  const seedMessage = useSeedCheckbox.checked && seedInput.value.trim() !== ''
    ? ` Se utiliz\u00f3 la semilla: ${seedInput.value.trim()}.`
    : '';
  showMessage(`Muestra generada correctamente con ${count} filas.${seedMessage} Ya puedes descargar el resultado en Excel o PDF.`);
});

// Activa o desactiva el uso de semilla reproducible.
useSeedCheckbox.addEventListener('change', () => {
  syncSeedFieldState();
});

activateLicenseBtn.addEventListener('click', () => {
  activateLicenseFromFile();
});

licenseFileInput.addEventListener('change', () => {
  clearInlineMessage(licenseMessage);
  licenseDetails.style.display = 'none';
  licenseDetails.innerHTML = '';
});

// Descarga solo el archivo Excel de la muestra.
downloadExcelBtn.addEventListener('click', () => {
  downloadBlob(exportSampleExcelBlob(), buildDownloadFileName('xlsx'));
});

// Descarga solo el PDF de la muestra.
downloadPdfBtn.addEventListener('click', () => {
  downloadBlob(exportSamplePdfBlob(), buildDownloadFileName('pdf'));
});

// Reinicia tanto el estado interno como la interfaz visible.
resetBtn.addEventListener('click', () => {
  resetState();
  fileInput.value = '';
  sampleCountInput.value = '';
  useSeedCheckbox.checked = false;
  seedInput.value = '';
  syncSeedFieldState();
  auditorInput.value = '';
  sheetSelect.innerHTML = '';
  sheetSelectorWrap.style.display = 'none';
  fileInfo.innerHTML = '';
  resultTable.innerHTML = '';
  resultCard.style.display = 'none';
  downloadExcelBtn.disabled = true;
  downloadPdfBtn.disabled = true;
  clearMessage();
});
