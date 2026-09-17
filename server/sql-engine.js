import alasql from 'alasql';

/**
 * Executa SQL SELECT sobre um array de objetos.
 * Tabela padrão: vendas (aliases: data, Vendas, Vendas_Empresariais).
 */
export function runSqlOnRows(sql, rows) {
  if (!sql || typeof sql !== 'string') {
    const err = new Error('SQL inválido');
    err.status = 400;
    throw err;
  }

  const trimmed = sql.trim().replace(/;+\s*$/, '');
  const forbidden = /\b(drop|delete|update|insert|alter|attach|outfile|truncate)\b/i;
  if (forbidden.test(trimmed)) {
    const err = new Error('Apenas consultas SELECT são permitidas');
    err.status = 400;
    throw err;
  }
  if (!/^\s*select\b/i.test(trimmed)) {
    const err = new Error('A query deve começar com SELECT');
    err.status = 400;
    throw err;
  }

  const sample = rows[0] || {};
  const originals = Object.keys(sample);
  const used = new Set();
  const fieldMap = {};

  originals.forEach((field) => {
    let safe = field
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .replace(/^[0-9]/, 'f_$&');
    if (!safe) safe = 'col';
    let candidate = safe;
    let i = 2;
    while (used.has(candidate.toLowerCase())) {
      candidate = `${safe}_${i++}`;
    }
    used.add(candidate.toLowerCase());
    fieldMap[field] = candidate;
  });

  const normalizedRows = rows.map((row) => {
    const out = {};
    originals.forEach((field) => {
      out[fieldMap[field]] = row[field];
    });
    return out;
  });

  let rewritten = trimmed;

  // Substitui nomes originais (mais longos primeiro) e formas com underscore
  const replacements = [];
  originals.forEach((field) => {
    const safe = fieldMap[field];
    replacements.push([field, safe]);
    const underscored = field.replace(/\s+/g, '_');
    if (underscored !== field) replacements.push([underscored, safe]);
    const noUnit = field.replace(/\s*\([^)]*\)\s*/g, '').trim();
    if (noUnit && noUnit !== field) {
      replacements.push([noUnit, safe]);
      replacements.push([noUnit.replace(/\s+/g, '_'), safe]);
    }
  });

  replacements.sort((a, b) => b[0].length - a[0].length);

  replacements.forEach(([from, to]) => {
    if (!from) return;
    const escaped = escapeRegExp(from);
    rewritten = rewritten.replace(new RegExp(`\\[${escaped}\\]`, 'gi'), to);
    rewritten = rewritten.replace(new RegExp(`"${escaped}"`, 'gi'), to);
    rewritten = rewritten.replace(new RegExp(`\`${escaped}\``, 'gi'), to);
    rewritten = rewritten.replace(new RegExp(`(?<![\\w])${escaped}(?![\\w])`, 'gi'), to);
  });

  rewritten = rewritten
    .replace(/\bFROM\s+Vendas_Empresariais\b/gi, 'FROM vendas')
    .replace(/\bFROM\s+Vendas\b/gi, 'FROM vendas')
    .replace(/\bFROM\s+data\b/gi, 'FROM vendas')
    .replace(/\bAS\s+Total\b/gi, 'AS total_valor')
    .replace(/\bAS\s+Total_Faturamento\b/gi, 'AS total_faturamento');

  const db = new alasql.Database();
  db.exec('CREATE TABLE vendas');
  db.tables.vendas.data = normalizedRows;

  try {
    const result = db.exec(rewritten);
    const list = Array.isArray(result) ? result : [];
    return {
      rows: list,
      rowCount: list.length,
      fields: list.length ? Object.keys(list[0]) : [],
      fieldMap,
      executedSql: rewritten,
    };
  } catch (e) {
    const err = new Error(`Erro SQL: ${e.message}`);
    err.status = 400;
    throw err;
  }
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
