const AUTH_KEY = 'csv_dashboard_auth';
const AUTH_VERSION = 'v1';
let REQUIRED_PASSWORD = '';

function readAuthState() {
  try {
    const raw = sessionStorage.getItem(AUTH_KEY);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (parsed && parsed.version === AUTH_VERSION && parsed.value === '1' && typeof parsed.issuedAt === 'number') {
      return parsed;
    }

    return null;
  } catch (error) {
    return null;
  }
}

function writeAuthState() {
  sessionStorage.setItem(AUTH_KEY, JSON.stringify({ version: AUTH_VERSION, value: '1', issuedAt: Date.now() }));
}

function clearAuthState() {
  sessionStorage.removeItem(AUTH_KEY);
}

function isAuthenticated() {
  return !!readAuthState();
}

function redirectToDashboard() {
  window.location.href = '/dashboard';
}

async function base64ToBytes(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function deriveKey(passphrase, saltBytes) {
  const keyMaterial = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(passphrase),
    'PBKDF2',
    false,
    ['deriveKey']
  );

  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: saltBytes, iterations: 100000, hash: 'SHA-256' },
    keyMaterial,
    { name: 'AES-GCM', length: 256 },
    false,
    ['decrypt']
  );
}

async function decryptSecretPayload(rawPayload, passphrase) {
  const payload = JSON.parse(rawPayload);
  const saltBytes = await base64ToBytes(payload.salt);
  const ivBytes = await base64ToBytes(payload.iv);
  const ciphertextBytes = await base64ToBytes(payload.ct);
  const key = await deriveKey(passphrase, saltBytes);
  const decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: ivBytes }, key, ciphertextBytes);
  return new TextDecoder().decode(decrypted);
}

async function loadRequiredPassword() {
  try {
    const [secretResponse, keyResponse] = await Promise.all([
      fetch('/secret.txt', { cache: 'no-store' }),
      fetch('/key.txt', { cache: 'no-store' })
    ]);

    if (!secretResponse.ok || !keyResponse.ok) {
      throw new Error('Password files are unavailable.');
    }

    const secretText = await secretResponse.text();
    const passphrase = (await keyResponse.text()).trim();
    const decrypted = await decryptSecretPayload(secretText, passphrase);
    REQUIRED_PASSWORD = decrypted.trim();

    if (!REQUIRED_PASSWORD) {
      throw new Error('Password is empty.');
    }
  } catch (error) {
    REQUIRED_PASSWORD = '';
    throw error;
  }
}

async function initializeLogin() {
  try {
    await loadRequiredPassword();
  } catch (error) {
    const loginError = document.getElementById('loginError');
    if (loginError) {
      loginError.textContent = 'Unable to load the secure password.';
    }
    return;
  }

  if (isAuthenticated()) {
    redirectToDashboard();
    return;
  }

  const form = document.getElementById('loginForm');
  const passwordInput = document.getElementById('loginPassword');
  const loginError = document.getElementById('loginError');

  if (form && passwordInput && loginError) {
    form.addEventListener('submit', function (event) {
      event.preventDefault();
      const password = passwordInput.value;

      if (password === REQUIRED_PASSWORD) {
        writeAuthState();
        loginError.textContent = '';
        passwordInput.value = '';
        redirectToDashboard();
      } else {
        clearAuthState();
        loginError.textContent = 'Incorrect password. Please try again.';
        passwordInput.value = '';
        passwordInput.focus();
      }
    });
  } else {
    console.warn('Login form is not available on this page.');
  }
}

initializeLogin();
