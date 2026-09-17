// auth.js — Auth real via API (JWT + cookie)
import { api, clearToken, getToken, setToken } from './api-client.js';

const SESSION_KEY = 'bi_cubo_enterprise_session_v1';

export function getSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function setSession(user) {
  try {
    localStorage.setItem(SESSION_KEY, JSON.stringify(user));
  } catch {}
}

export function clearSession() {
  try {
    localStorage.removeItem(SESSION_KEY);
  } catch {}
  clearToken();
}

function mapUser(user) {
  if (!user) return null;
  return {
    id: user.id,
    username: user.email,
    email: user.email,
    name: user.name,
    role: user.role || 'analyst',
    loginTime: Date.now(),
  };
}

export function initAuth(onUserChange) {
  const loginModal = document.getElementById('loginModal');
  const loginForm = document.getElementById('loginForm');
  const inputUser = document.getElementById('inputUser');
  const inputPass = document.getElementById('inputPass');
  const inputName = document.getElementById('inputName');
  const userNameLabel = document.getElementById('userNameLabel');
  const btnLogout = document.getElementById('btnLogout');
  const authModeHint = document.getElementById('authModeHint');
  const authToggleBtn = document.getElementById('authToggleMode');
  const submitLoginBtn = document.getElementById('submitLoginBtn');
  const authError = document.getElementById('authError');
  const nameGroup = document.getElementById('nameGroup');

  let mode = 'login'; // login | register

  function showError(msg) {
    if (!authError) return;
    authError.textContent = msg || '';
    authError.classList.toggle('hidden', !msg);
  }

  function setMode(next) {
    mode = next;
    const isRegister = mode === 'register';
    if (nameGroup) nameGroup.classList.toggle('hidden', !isRegister);
    if (authModeHint) {
      authModeHint.textContent = isRegister
        ? 'Crie sua conta para salvar cubos e datasets na nuvem.'
        : 'Acesse com e-mail e senha. Demo: admin / admin';
    }
    if (authToggleBtn) {
      authToggleBtn.textContent = isRegister
        ? 'Já tenho conta — entrar'
        : 'Criar nova conta';
    }
    if (submitLoginBtn) {
      const span = submitLoginBtn.querySelector('span');
      if (span) span.textContent = isRegister ? 'Criar conta' : 'Entrar no Sistema';
    }
    showError('');
  }

  function updateUI(session) {
    if (session && (session.username || session.email)) {
      loginModal.classList.add('hidden');
      if (userNameLabel) userNameLabel.textContent = session.name || session.email || session.username;
      if (onUserChange) onUserChange(session);
    } else {
      loginModal.classList.remove('hidden');
    }
  }

  async function restoreSession() {
    if (!getToken()) {
      clearSession();
      updateUI(null);
      return;
    }
    try {
      const data = await api('/me');
      const session = mapUser(data.user);
      setSession(session);
      updateUI(session);
    } catch {
      clearSession();
      updateUI(null);
    }
  }

  if (authToggleBtn) {
    authToggleBtn.addEventListener('click', (e) => {
      e.preventDefault();
      setMode(mode === 'login' ? 'register' : 'login');
    });
  }

  if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      showError('');

      const userVal = (inputUser?.value || '').trim();
      const passVal = (inputPass?.value || '').trim();
      const nameVal = (inputName?.value || '').trim();

      if (!userVal || !passVal) {
        showError('Preencha usuário/e-mail e senha.');
        return;
      }

      const submitBtn = submitLoginBtn;
      if (submitBtn) submitBtn.disabled = true;

      try {
        const endpoint = mode === 'register' ? '/register' : '/login';
        const body =
          mode === 'register'
            ? { email: userVal.includes('@') ? userVal : `${userVal}@bicubo.app`, password: passVal, name: nameVal || userVal }
            : { email: userVal, password: passVal };

        const data = await api(endpoint, { method: 'POST', body, auth: false });
        setToken(data.token);
        const session = mapUser(data.user);
        setSession(session);
        updateUI(session);
      } catch (err) {
        showError(err.message || 'Falha na autenticação');
      } finally {
        if (submitBtn) submitBtn.disabled = false;
      }
    });
  }

  if (btnLogout) {
    btnLogout.addEventListener('click', async () => {
      try {
        await api('/logout', { method: 'POST' });
      } catch {}
      clearSession();
      updateUI(null);
      window.location.reload();
    });
  }

  setMode('login');
  restoreSession();
}
