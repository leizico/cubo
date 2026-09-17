import { Hono } from 'hono';
import { cors } from 'hono/cors';
import bcrypt from 'bcryptjs';
import {
  clearAuthCookie,
  readAuthToken,
  setAuthCookie,
  signToken,
  verifyToken,
} from './auth-tokens.js';
import {
  createUser,
  deleteCube,
  deleteDataset,
  findUserByEmail,
  findUserById,
  getDataset,
  getStoreMeta,
  listCubes,
  listDatasets,
  sanitizeUser,
  saveCube,
  saveDataset,
} from './store.js';
import { runSqlOnRows } from './sql-engine.js';

const app = new Hono().basePath('/api');

app.use('*', cors({
  origin: (origin) => origin || '*',
  credentials: true,
  allowHeaders: ['Content-Type', 'Authorization'],
  allowMethods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
}));

async function requireUser(c) {
  const token = readAuthToken(c);
  const payload = await verifyToken(token);
  if (!payload?.id) return null;
  const user = await findUserById(payload.id);
  return user ? sanitizeUser(user) : null;
}

app.get('/health', async (c) => {
  const meta = await getStoreMeta();
  return c.json({ ok: true, ...meta, ts: Date.now() });
});

app.post('/register', async (c) => {
  try {
    const body = await c.req.json();
    const email = String(body.email || '').trim().toLowerCase();
    const password = String(body.password || '');
    const name = String(body.name || email.split('@')[0] || 'Analista').trim();

    if (!email || !email.includes('@')) return c.json({ error: 'E-mail inválido' }, 400);
    if (password.length < 6) return c.json({ error: 'Senha deve ter pelo menos 6 caracteres' }, 400);

    const passwordHash = await bcrypt.hash(password, 10);
    const user = await createUser({ email, passwordHash, name });
    const token = await signToken(user);
    setAuthCookie(c, token);
    return c.json({ user, token });
  } catch (e) {
    return c.json({ error: e.message || 'Falha no registro' }, e.status || 500);
  }
});

app.post('/login', async (c) => {
  try {
    const body = await c.req.json();
    const email = String(body.email || body.username || '').trim().toLowerCase();
    const password = String(body.password || '');
    if (!email || !password) return c.json({ error: 'Informe e-mail e senha' }, 400);

    let userRow = await findUserByEmail(email);

    if (!userRow && (email === 'admin' || email === 'admin@bicubo.app') && password === 'admin') {
      const passwordHash = await bcrypt.hash('admin', 10);
      const created = await createUser({
        email: 'admin@bicubo.app',
        passwordHash,
        name: 'Admin Analista',
      });
      userRow = await findUserById(created.id);
    }

    if (!userRow && email === 'admin') {
      userRow = await findUserByEmail('admin@bicubo.app');
    }

    if (!userRow) return c.json({ error: 'Credenciais inválidas' }, 401);
    const ok = await bcrypt.compare(password, userRow.passwordHash);
    if (!ok) return c.json({ error: 'Credenciais inválidas' }, 401);

    const user = sanitizeUser(userRow);
    const token = await signToken(user);
    setAuthCookie(c, token);
    return c.json({ user, token });
  } catch (e) {
    return c.json({ error: e.message || 'Falha no login' }, 500);
  }
});

app.post('/logout', async (c) => {
  clearAuthCookie(c);
  return c.json({ ok: true });
});

app.get('/me', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ user: null }, 401);
  return c.json({ user });
});

app.get('/cubes', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  return c.json({ cubes: await listCubes(user.id) });
});

app.post('/cubes', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  try {
    const body = await c.req.json();
    if (!body.name || !String(body.name).trim()) {
      return c.json({ error: 'Nome do cubo é obrigatório' }, 400);
    }
    return c.json({ cube: await saveCube(user.id, body) });
  } catch (e) {
    return c.json({ error: e.message || 'Falha ao salvar cubo' }, e.status || 500);
  }
});

app.delete('/cubes', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  const id = c.req.query('id');
  if (!id) return c.json({ error: 'id obrigatório' }, 400);
  const ok = await deleteCube(user.id, id);
  if (!ok) return c.json({ error: 'Cubo não encontrado' }, 404);
  return c.json({ ok: true });
});

app.get('/datasets', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  return c.json({ datasets: await listDatasets(user.id) });
});

app.post('/datasets', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  try {
    const body = await c.req.json();
    const meta = await saveDataset(user.id, { name: body.name, rows: body.rows });
    return c.json({ dataset: meta });
  } catch (e) {
    return c.json({ error: e.message || 'Falha ao salvar dataset' }, e.status || 500);
  }
});

app.get('/dataset', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  const id = c.req.query('id');
  if (!id) return c.json({ error: 'id obrigatório' }, 400);
  const dataset = await getDataset(user.id, id);
  if (!dataset) return c.json({ error: 'Dataset não encontrado' }, 404);
  return c.json({
    dataset: {
      id: dataset.id,
      name: dataset.name,
      fields: dataset.fields,
      rowCount: dataset.rowCount,
      createdAt: dataset.createdAt,
      rows: dataset.rows,
    },
  });
});

app.delete('/dataset', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  const id = c.req.query('id');
  if (!id) return c.json({ error: 'id obrigatório' }, 400);
  const ok = await deleteDataset(user.id, id);
  if (!ok) return c.json({ error: 'Dataset não encontrado' }, 404);
  return c.json({ ok: true });
});

app.post('/sql', async (c) => {
  const user = await requireUser(c);
  if (!user) return c.json({ error: 'Não autenticado' }, 401);
  try {
    const body = await c.req.json();
    let rows = Array.isArray(body.rows) ? body.rows : null;
    if (!rows && body.datasetId) {
      const dataset = await getDataset(user.id, body.datasetId);
      if (!dataset) return c.json({ error: 'Dataset não encontrado' }, 404);
      rows = dataset.rows;
    }
    if (!rows || !rows.length) {
      return c.json({ error: 'Nenhum dado disponível para consultar' }, 400);
    }
    const result = await runSqlOnRows(body.sql || '', rows);
    return c.json({ result });
  } catch (e) {
    return c.json({ error: e.message || 'Falha na consulta SQL' }, e.status || 500);
  }
});

app.notFound((c) => c.json({ error: 'Rota não encontrada' }, 404));

export default app;
