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
  const btnOpenLogin = document.getElementById('btnOpenLogin');
  const btnSkipLogin = document.getElementById('btnSkipLogin');
  const userAvatar = document.getElementById('userAvatar');
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
    const logged = !!(session && (session.username || session.email));
    if (logged) {
      loginModal?.classList.add('hidden');
      if (userNameLabel) userNameLabel.textContent = session.name || session.email || session.username;
      if (userAvatar) userAvatar.textContent = String(session.name || session.email || 'U').slice(0, 2).toUpperCase();
      btnLogout?.classList.remove('hidden');
      btnOpenLogin?.classList.add('hidden');
    } else {
      if (userNameLabel) userNameLabel.textContent = 'Visitante';
      if (userAvatar) userAvatar.textContent = 'VI';
      btnLogout?.classList.add('hidden');
      btnOpenLogin?.classList.remove('hidden');
    }
    if (onUserChange) onUserChange(session || null);
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
        const msg = err.name === 'TimeoutError'
          ? 'O servidor demorou para responder. Você pode continuar sem login.'
          : (err.message || 'Falha na autenticação');
        showError(msg);
      } finally {
        if (submitBtn) submitBtn.disabled = false;
      }
    });
  }

  if (btnOpenLogin) {
    btnOpenLogin.addEventListener('click', () => loginModal?.classList.remove('hidden'));
  }
  if (btnSkipLogin) {
    btnSkipLogin.addEventListener('click', () => loginModal?.classList.add('hidden'));
  }

  if (btnLogout) {
    btnLogout.addEventListener('click', async () => {
      try {
        await api('/logout', { method: 'POST' });
      } catch {}
      clearSession();
      updateUI(null);
    });
  }

  setMode('login');
  restoreSession();
}
