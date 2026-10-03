const REQUIRED_PASSWORD = 'AltDshb543!';
const AUTH_KEY = 'csv_dashboard_auth';
const AUTH_VERSION = 'v1';

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
  window.location.href = 'dashboard.html';
}

if (isAuthenticated()) {
  redirectToDashboard();
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
