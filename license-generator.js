const adminGate = document.getElementById('adminGate');
const generatorShell = document.getElementById('generatorShell');
const adminUsernameInput = document.getElementById('adminUsernameInput');
const adminPasswordInput = document.getElementById('adminPasswordInput');
const unlockGeneratorBtn = document.getElementById('unlockGeneratorBtn');
const logoutGeneratorBtn = document.getElementById('logoutGeneratorBtn');
const adminAccessMessage = document.getElementById('adminAccessMessage');
const adminSessionBadge = document.getElementById('adminSessionBadge');

const companyInput = document.getElementById('companyInput');
const issuedToInput = document.getElementById('issuedToInput');
const expiresAtInput = document.getElementById('expiresAtInput');
const productInput = document.getElementById('productInput');
const generateLicenseBtn = document.getElementById('generateLicenseBtn');
const downloadLicenseBtn = document.getElementById('downloadLicenseBtn');
const copyLicenseBtn = document.getElementById('copyLicenseBtn');
const licenseGeneratorMessage = document.getElementById('licenseGeneratorMessage');
const licenseResultCard = document.getElementById('licenseResultCard');
const licenseMeta = document.getElementById('licenseMeta');
const licenseKeyPreview = document.getElementById('licenseKeyPreview');
const licenseJsonOutput = document.getElementById('licenseJsonOutput');

const currentAdminPasswordInput = document.getElementById('currentAdminPasswordInput');
const newAdminPasswordInput = document.getElementById('newAdminPasswordInput');
const confirmAdminPasswordInput = document.getElementById('confirmAdminPasswordInput');
const changeAdminPasswordBtn = document.getElementById('changeAdminPasswordBtn');
const adminPasswordChangeMessage = document.getElementById('adminPasswordChangeMessage');

let currentLicensePayload = null;

function showBoxMessage(element, text, isError = false) {
  element.textContent = text;
  element.classList.add('show');
  element.style.borderColor = isError ? '#7f1d1d' : '#b8bec5';
  element.style.background = isError ? '#fee2e2' : '#f8fafc';
  element.style.color = isError ? '#7f1d1d' : '#1f2937';
}

function clearBoxMessage(element) {
  element.textContent = '';
  element.classList.remove('show');
}

function showAdminMessage(text, isError = false) {
  showBoxMessage(adminAccessMessage, text, isError);
}

function clearAdminMessage() {
  clearBoxMessage(adminAccessMessage);
}

function showLicenseMessage(text, isError = false) {
  showBoxMessage(licenseGeneratorMessage, text, isError);
}

function clearLicenseMessage() {
  clearBoxMessage(licenseGeneratorMessage);
}

function showPasswordChangeMessage(text, isError = false) {
  showBoxMessage(adminPasswordChangeMessage, text, isError);
}

function clearPasswordChangeMessage() {
  clearBoxMessage(adminPasswordChangeMessage);
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function slugifyFilePart(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
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

function lockGenerator() {
  adminGate.style.display = 'flex';
  generatorShell.style.display = 'none';
  adminPasswordInput.value = '';
}

function unlockGenerator(user) {
  adminGate.style.display = 'none';
  generatorShell.style.display = 'block';
  adminSessionBadge.textContent = `Modo administrador activo: ${user}`;
}

function resetPasswordChangeForm() {
  currentAdminPasswordInput.value = '';
  newAdminPasswordInput.value = '';
  confirmAdminPasswordInput.value = '';
}

function buildDownloadName(payload) {
  const company = slugifyFilePart(payload.empresa) || 'empresa';
  const user = slugifyFilePart(payload.titular) || 'titular';
  return `licencia_${company}_${user}_${payload.vence}.json`;
}

function downloadTextFile(content, fileName) {
  const blob = new Blob([content], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function renderLicense(payload) {
  currentLicensePayload = payload;
  licenseResultCard.style.display = 'block';
  licenseKeyPreview.textContent = `Clave generada: ${payload.licencia}`;
  licenseJsonOutput.value = JSON.stringify(payload, null, 2);
  licenseMeta.innerHTML = `
    <span class="pill">Producto: ${escapeHtml(payload.producto)}</span>
    <span class="pill">Empresa: ${escapeHtml(payload.empresa)}</span>
    <span class="pill">Titular: ${escapeHtml(payload.titular)}</span>
    <span class="pill">Vence: ${escapeHtml(payload.vence)}</span>
  `;
  downloadLicenseBtn.disabled = false;
  copyLicenseBtn.disabled = false;
}

async function checkSession() {
  try {
    const data = await apiFetch('/api/admin/session', { method: 'GET', headers: {} });
    if (data.authenticated) {
      unlockGenerator(data.user);
    } else {
      lockGenerator();
      showAdminMessage('Este módulo requiere autenticación administrativa.');
    }
  } catch (error) {
    lockGenerator();
    showAdminMessage('No se pudo conectar con el backend de administración. Verifica que el servidor esté en ejecución.', true);
  }
}

unlockGeneratorBtn.addEventListener('click', async () => {
  clearAdminMessage();

  const username = adminUsernameInput.value.trim();
  const password = adminPasswordInput.value.trim();

  if (!username || !password) {
    showAdminMessage('Ingresa usuario y clave de administrador.', true);
    return;
  }

  try {
    const data = await apiFetch('/api/admin/login', {
      method: 'POST',
      body: JSON.stringify({ username, password })
    });
    unlockGenerator(data.user);
  } catch (error) {
    showAdminMessage(error.message, true);
    adminPasswordInput.select();
  }
});

logoutGeneratorBtn.addEventListener('click', async () => {
  try {
    await apiFetch('/api/admin/logout', { method: 'POST', body: JSON.stringify({}) });
  } catch (error) {
    // Aunque falle la llamada, devolvemos la UI al estado bloqueado.
  }

  lockGenerator();
  showAdminMessage('La sesión administrativa se cerró correctamente.');
});

generateLicenseBtn.addEventListener('click', async () => {
  clearLicenseMessage();

  const company = companyInput.value.trim();
  const issuedTo = issuedToInput.value.trim();
  const expiresAt = expiresAtInput.value;

  if (!company || !issuedTo || !expiresAt) {
    showLicenseMessage('Completa empresa, titular y fecha de vencimiento antes de generar la licencia.', true);
    return;
  }

  try {
    const data = await apiFetch('/api/licenses/generate', {
      method: 'POST',
      body: JSON.stringify({ company, issuedTo, expiresAt })
    });
    renderLicense(data.license);
    showLicenseMessage(data.message || 'Licencia generada correctamente.');
  } catch (error) {
    currentLicensePayload = null;
    downloadLicenseBtn.disabled = true;
    copyLicenseBtn.disabled = true;
    licenseResultCard.style.display = 'none';
    showLicenseMessage(error.message, true);
  }
});

downloadLicenseBtn.addEventListener('click', () => {
  if (!currentLicensePayload) return;
  downloadTextFile(JSON.stringify(currentLicensePayload, null, 2), buildDownloadName(currentLicensePayload));
});

copyLicenseBtn.addEventListener('click', async () => {
  if (!currentLicensePayload) return;

  try {
    await navigator.clipboard.writeText(JSON.stringify(currentLicensePayload, null, 2));
    showLicenseMessage('El contenido JSON de la licencia se copió al portapapeles.');
  } catch (error) {
    showLicenseMessage('No se pudo copiar al portapapeles. Intenta usar la descarga del archivo.', true);
  }
});

changeAdminPasswordBtn.addEventListener('click', async () => {
  clearPasswordChangeMessage();

  const currentPassword = currentAdminPasswordInput.value.trim();
  const newPassword = newAdminPasswordInput.value.trim();
  const confirmPassword = confirmAdminPasswordInput.value.trim();

  if (!currentPassword || !newPassword || !confirmPassword) {
    showPasswordChangeMessage('Completa la clave actual, la nueva clave y la confirmación.', true);
    return;
  }

  try {
    const data = await apiFetch('/api/admin/change-password', {
      method: 'POST',
      body: JSON.stringify({ currentPassword, newPassword, confirmPassword })
    });
    resetPasswordChangeForm();
    showPasswordChangeMessage(data.message || 'La clave se actualizó correctamente.');
  } catch (error) {
    showPasswordChangeMessage(error.message, true);
  }
});

adminPasswordInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    unlockGeneratorBtn.click();
  }
});

expiresAtInput.value = getBogotaIsoDate();
checkSession();
