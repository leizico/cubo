import { randomUUID } from 'crypto';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, '..', 'data');
const STORE_PATH = path.join(DATA_DIR, 'store.json');

const emptyStore = () => ({
  users: [],
  datasets: [],
  cubes: [],
});

/** Cache em memória (útil no serverless da Vercel entre invocações quentes). */
let memoryStore = null;
let writeQueue = Promise.resolve();

async function ensureLoaded() {
  if (memoryStore) return memoryStore;

  try {
    const raw = await fs.readFile(STORE_PATH, 'utf8');
    memoryStore = JSON.parse(raw);
  } catch {
    memoryStore = emptyStore();
    try {
      await fs.mkdir(DATA_DIR, { recursive: true });
      await fs.writeFile(STORE_PATH, JSON.stringify(memoryStore, null, 2), 'utf8');
    } catch {
      // Em ambientes read-only (serverless) seguimos só em memória.
    }
  }

  if (!memoryStore.users) memoryStore.users = [];
  if (!memoryStore.datasets) memoryStore.datasets = [];
  if (!memoryStore.cubes) memoryStore.cubes = [];
  return memoryStore;
}

async function persist() {
  if (!memoryStore) return;
  writeQueue = writeQueue.then(async () => {
    try {
      await fs.mkdir(DATA_DIR, { recursive: true });
      await fs.writeFile(STORE_PATH, JSON.stringify(memoryStore, null, 2), 'utf8');
    } catch {
      // Sem disco persistente — ok em demo serverless.
    }
  });
  return writeQueue;
}

export async function getStoreMeta() {
  const s = await ensureLoaded();
  return {
    mode: process.env.VERCEL ? 'memory+ephemeral-disk' : 'file',
    users: s.users.length,
    datasets: s.datasets.length,
    cubes: s.cubes.length,
  };
}

export async function findUserByEmail(email) {
  const s = await ensureLoaded();
  return s.users.find((u) => u.email.toLowerCase() === email.toLowerCase()) || null;
}

export async function findUserById(id) {
  const s = await ensureLoaded();
  return s.users.find((u) => u.id === id) || null;
}

export async function createUser({ email, passwordHash, name }) {
  const s = await ensureLoaded();
  if (s.users.some((u) => u.email.toLowerCase() === email.toLowerCase())) {
    const err = new Error('E-mail já cadastrado');
    err.status = 409;
    throw err;
  }
  const user = {
    id: randomUUID(),
    email: email.trim().toLowerCase(),
    passwordHash,
    name: name.trim(),
    role: 'analyst',
    createdAt: Date.now(),
  };
  s.users.push(user);
  await persist();
  return sanitizeUser(user);
}

export function sanitizeUser(user) {
  if (!user) return null;
  return {
    id: user.id,
    email: user.email,
    name: user.name,
    role: user.role,
    createdAt: user.createdAt,
  };
}

export async function listCubes(userId) {
  const s = await ensureLoaded();
  return s.cubes
    .filter((c) => c.userId === userId)
    .sort((a, b) => b.savedAt - a.savedAt)
    .map((c) => ({
      id: c.id,
      name: c.name,
      datasetId: c.datasetId || null,
      savedAt: c.savedAt,
      pivot: c.pivot,
      fieldAliases: c.fieldAliases,
      fieldSortDirs: c.fieldSortDirs,
      filterSelected: c.filterSelected,
      derivedFields: c.derivedFields || [],
      chart: c.chart,
      chartType: c.chartType,
    }));
}

export async function getCube(userId, cubeId) {
  const s = await ensureLoaded();
  const cube = s.cubes.find((c) => c.id === cubeId && c.userId === userId);
  return cube || null;
}

export async function saveCube(userId, payload) {
  const s = await ensureLoaded();
  const now = Date.now();
  if (payload.id) {
    const idx = s.cubes.findIndex((c) => c.id === payload.id && c.userId === userId);
    if (idx >= 0) {
      s.cubes[idx] = {
        ...s.cubes[idx],
        ...payload,
        userId,
        savedAt: now,
      };
      await persist();
      return s.cubes[idx];
    }
  }
  const cube = {
    id: payload.id || randomUUID(),
    userId,
    name: payload.name,
    datasetId: payload.datasetId || null,
    pivot: payload.pivot || { rows: [], cols: [], filters: [], values: [] },
    fieldAliases: payload.fieldAliases || {},
    fieldSortDirs: payload.fieldSortDirs || {},
    filterSelected: payload.filterSelected || {},
    derivedFields: payload.derivedFields || [],
    chart: payload.chart || null,
    chartType: payload.chartType || 'bar',
    savedAt: now,
  };
  s.cubes.push(cube);
  await persist();
  return cube;
}

export async function deleteCube(userId, cubeId) {
  const s = await ensureLoaded();
  const before = s.cubes.length;
  s.cubes = s.cubes.filter((c) => !(c.id === cubeId && c.userId === userId));
  if (s.cubes.length === before) return false;
  await persist();
  return true;
}

export async function listDatasets(userId) {
  const s = await ensureLoaded();
  return s.datasets
    .filter((d) => d.userId === userId)
    .sort((a, b) => b.createdAt - a.createdAt)
    .map((d) => ({
      id: d.id,
      name: d.name,
      fields: d.fields,
      rowCount: d.rowCount,
      createdAt: d.createdAt,
    }));
}

export async function getDataset(userId, datasetId) {
  const s = await ensureLoaded();
  return s.datasets.find((d) => d.id === datasetId && d.userId === userId) || null;
}

export async function saveDataset(userId, { name, rows }) {
  if (!Array.isArray(rows) || !rows.length) {
    const err = new Error('Dataset vazio');
    err.status = 400;
    throw err;
  }
  const jsonSize = Buffer.byteLength(JSON.stringify(rows), 'utf8');
  const maxBytes = 3 * 1024 * 1024;
  if (jsonSize > maxBytes) {
    const err = new Error('Dataset excede o limite de 3 MB');
    err.status = 413;
    throw err;
  }

  const s = await ensureLoaded();
  const dataset = {
    id: randomUUID(),
    userId,
    name: name || 'Dataset',
    fields: Object.keys(rows[0] || {}),
    rowCount: rows.length,
    rows,
    createdAt: Date.now(),
  };
  s.datasets.push(dataset);

  // Mantém no máximo 20 datasets por usuário
  const mine = s.datasets.filter((d) => d.userId === userId).sort((a, b) => b.createdAt - a.createdAt);
  if (mine.length > 20) {
    const dropIds = new Set(mine.slice(20).map((d) => d.id));
    s.datasets = s.datasets.filter((d) => !dropIds.has(d.id));
  }

  await persist();
  return {
    id: dataset.id,
    name: dataset.name,
    fields: dataset.fields,
    rowCount: dataset.rowCount,
    createdAt: dataset.createdAt,
  };
}

export async function deleteDataset(userId, datasetId) {
  const s = await ensureLoaded();
  const before = s.datasets.length;
  s.datasets = s.datasets.filter((d) => !(d.id === datasetId && d.userId === userId));
  if (s.datasets.length === before) return false;
  await persist();
  return true;
}
