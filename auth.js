// auth.js — Enterprise Session & User Profile Manager

const SESSION_KEY = 'bi_cubo_enterprise_session_v1';

export function getSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

export function setSession(user) {
  try {
    localStorage.setItem(SESSION_KEY, JSON.stringify(user));
  } catch (e) {}
}

export function clearSession() {
  try {
    localStorage.removeItem(SESSION_KEY);
  } catch (e) {}
}

export function initAuth(onUserChange) {
  const loginModal = document.getElementById('loginModal');
  const loginForm = document.getElementById('loginForm');
  const inputUser = document.getElementById('inputUser');
  const inputPass = document.getElementById('inputPass');
  const userNameLabel = document.getElementById('userNameLabel');
  const btnLogout = document.getElementById('btnLogout');

  function updateUI() {
    const session = getSession();
    if (session && session.username) {
      loginModal.classList.add('hidden');
      if (userNameLabel) userNameLabel.textContent = session.name || session.username;
      if (onUserChange) onUserChange(session);
    } else {
      loginModal.classList.remove('hidden');
    }
  }

  if (loginForm) {
    loginForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const userVal = inputUser.value.trim();
      const passVal = inputPass.value.trim();

      if (!userVal || !passVal) {
        alert('Por favor, preencha o usuário e a senha.');
        return;
      }

      // Demo login validation
      const sessionObj = {
        username: userVal,
        name: userVal === 'admin' ? 'Admin Analista' : userVal,
        role: 'Director BI',
        loginTime: Date.now()
      };

      setSession(sessionObj);
      updateUI();
    });
  }

  if (btnLogout) {
    btnLogout.addEventListener('click', () => {
      clearSession();
      updateUI();
    });
  }

  updateUI();
}
