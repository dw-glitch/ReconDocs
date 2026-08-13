import { norm, text } from "./analysis-engine.js";

/**
 * Leitura genérica de planilhas: descobre a linha de cabeçalho, a coluna do
 * documento, a coluna de status e a coluna de data em qualquer layout, sem
 * depender de um modelo de planilha específico. Todo mapeamento pode ser
 * revisto manualmente pelo usuário.
 */

export const HEADER_ALIASES = {
  document: [
    "DOCUMENTO", "DOCUMENTOS", "NOME DOCUMENTO", "NOME DO DOCUMENTO", "CODIGO",
    "CODIGO DO DOCUMENTO", "CODIGO DOCUMENTAL", "NUMERO DO DOCUMENTO", "NUMERO DOCUMENTO",
    "N DOCUMENTO", "N DO DOCUMENTO", "DOCUMENT NUMBER", "DOC NUMBER", "DOCUMENT ID",
    "DOC", "ARQUIVO", "NOME DO ARQUIVO", "NOME DO CHECKLIST", "TAG", "IDENTIFICADOR",
  ],
  status: [
    "STATUS", "STATUS SIGEM", "STATUS DO DOCUMENTO", "STATUS DA INSPECAO", "SITUACAO",
    "SITUACAO SIGEM", "SITUACAO DO DOCUMENTO", "ESTADO", "FASE", "ETAPA", "CONDICAO",
  ],
  date: [
    "DATA", "DATA DO DOCUMENTO", "DATA DE POSTAGEM", "DATA POSTAGEM", "DATA DE EMISSAO",
    "DATA EMISSAO", "DATA DE ENVIO", "DATA DE CADASTRO", "DATA DE ATUALIZACAO",
    "ULTIMA ATUALIZACAO", "DATE", "EMISSAO",
  ],
};

const DATE_TEXT = /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s+\d{1,2}:\d{2})?$/;
const PLAIN_NUMBER = [
  /^\d+$/,
  /^\d+[.,]\d+$/,
  /^\d{1,3}(\.\d{3})+(,\d+)?$/,
  /^\d{1,3}(,\d{3})+(\.\d+)?$/,
];

const SPREADSHEET_EXTENSION = /\.(xlsx|xls|xlsm|csv|tsv)$/i;

/**
 * Nome de exibição da planilha a partir do arquivo carregado: o próprio nome
 * do arquivo, sem a extensão. É esse nome que aparece na tela e em todos os
 * títulos, colunas e abas do relatório.
 */
export function labelFromFileName(fileName) {
  return text(fileName).replace(SPREADSHEET_EXTENSION, "").trim();
}

/**
 * Garante que duas planilhas não fiquem com o mesmo nome — dois arquivos de
 * mesmo nome gerariam colunas e abas indistinguíveis no relatório.
 */
export function uniqueLabelAmong(desired, taken = []) {
  const used = new Set(taken.map((label) => text(label).toLowerCase()));
  const base = text(desired);
  if (!used.has(base.toLowerCase())) return base;
  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${base} (${suffix})`;
    if (!used.has(candidate.toLowerCase())) return candidate;
  }
}

export function normalizedHeader(value) {
  return norm(value).replace(/[^A-Z0-9]+/g, " ").trim();
}

export function headerAliasScore(header, field) {
  const key = normalizedHeader(header);
  if (!key) return 0;
  const aliases = HEADER_ALIASES[field] || [];
  if (aliases.includes(key)) return 12;
  if (aliases.some((alias) => key === `${alias}S` || key.startsWith(`${alias} `))) return 8;
  if (aliases.some((alias) => key.includes(alias) || alias.includes(key))) return 4;
  return 0;
}

export function looksLikeIdentifier(value) {
  const key = norm(value);
  if (key.length < 4) return false;
  if (PLAIN_NUMBER.some((pattern) => pattern.test(key))) return false;
  if (DATE_TEXT.test(key)) return false;
  return /\d/.test(key) && /[A-Z0-9][-_.\s]?/.test(key);
}

export function looksLikeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  return DATE_TEXT.test(text(value));
}

/**
 * Converte o valor de uma célula de data em `Date`, para que o relatório
 * exporte data de verdade — ordenável e filtrável no Excel — em vez de texto.
 * Texto no formato dd/mm/aaaa também é reconhecido. Devolve `null` quando o
 * valor não é uma data.
 */
export function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const raw = text(value);
  const match = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?$/);
  if (!match) return null;
  const [, day, month, year, hour = "0", minute = "0"] = match;
  const fullYear = year.length === 2 ? 2000 + Number(year) : Number(year);
  const parsed = new Date(fullYear, Number(month) - 1, Number(day), Number(hour), Number(minute));
  if (Number.isNaN(parsed.getTime())) return null;
  if (parsed.getDate() !== Number(day) || parsed.getMonth() !== Number(month) - 1) return null;
  return parsed;
}

export function detectHeaderRow(rows, limit = 200) {
  let best = { index: 0, score: -Infinity };
  const scanned = Math.min(rows.length, limit);
  for (let index = 0; index < scanned; index += 1) {
    const cells = (rows[index] || []).map((cell) => text(cell));
    const filled = cells.filter(Boolean);
    if (filled.length < 2) continue;
    const textual = filled.filter((value) => /[A-Za-zÀ-ÿ]/.test(value)).length;
    const numeric = filled.filter((value) => /^[\d.,%]+$/.test(value)).length;
    const identifiers = filled.filter(looksLikeIdentifier).length;
    const filters = filled.filter((value) => /:\s*$/.test(value)).length;
    const populatedBelow = rows.slice(index + 1, index + 12).filter((row) => (row || []).some((cell) => text(cell))).length;
    const score = textual * 3 + filled.length - numeric * 4 - identifiers * 3 - filters * 5 + populatedBelow * 2 - index * 0.5;
    if (score > best.score) best = { index, score };
  }
  return best.score > 0 ? best.index : -1;
}

function columnSample(rows, column, dataStart, size = 300) {
  return rows.slice(dataStart, dataStart + size)
    .map((row) => (row || [])[column])
    .filter((value) => text(value));
}

export function detectColumns(rows, headers, dataStart) {
  const mapping = { document: -1, status: -1, date: -1, extras: [] };
  const columns = headers.length;

  let bestDocument = { index: -1, score: 0 };
  let bestStatus = { index: -1, score: 0 };
  let bestDate = { index: -1, score: 0 };

  for (let column = 0; column < columns; column += 1) {
    const sample = columnSample(rows, column, dataStart);
    const header = headers[column];
    const filled = sample.length;
    const distinct = new Set(sample.map((value) => norm(value))).size;

    const identifierRate = filled ? sample.filter(looksLikeIdentifier).length / filled : 0;
    const uniqueRate = filled ? distinct / filled : 0;
    const documentScore = headerAliasScore(header, "document") * 2 + identifierRate * 14 + uniqueRate * 6 + (filled ? 1 : 0);
    if (filled && documentScore > bestDocument.score) bestDocument = { index: column, score: documentScore };

    const repeatRate = filled ? 1 - uniqueRate : 0;
    const shortTexts = filled ? sample.filter((value) => text(value).length <= 40 && /[A-Za-zÀ-ÿ]/.test(text(value))).length / filled : 0;
    const statusScore = headerAliasScore(header, "status") * 3 + repeatRate * 5 + shortTexts * 3;
    if (filled && headerAliasScore(header, "status") > 0 && statusScore > bestStatus.score) bestStatus = { index: column, score: statusScore };

    const dateRate = filled ? sample.filter(looksLikeDate).length / filled : 0;
    const dateScore = headerAliasScore(header, "date") * 3 + dateRate * 10;
    if (filled && dateScore >= 6 && dateScore > bestDate.score) bestDate = { index: column, score: dateScore };
  }

  mapping.document = bestDocument.index;
  mapping.status = bestStatus.index;
  mapping.date = bestDate.index;
  return mapping;
}

/**
 * Perfila uma aba já convertida em matriz de linhas.
 * @param rows matriz `unknown[][]` da aba.
 * @param options `{ columnLabel }` — gerador do rótulo de colunas sem cabeçalho.
 */
export function profileRows(rows, options = {}) {
  const columnLabel = options.columnLabel || ((index) => `Coluna ${index + 1}`);
  const headerIndex = detectHeaderRow(rows);
  const hasHeader = headerIndex >= 0;
  const dataStart = hasHeader ? headerIndex + 1 : 0;
  let maxColumns = 1;
  for (const row of rows) maxColumns = Math.max(maxColumns, (row || []).length);
  const headers = Array.from({ length: maxColumns }, (_, index) => {
    const raw = hasHeader ? text(rows[headerIndex]?.[index]) : "";
    return raw || columnLabel(index);
  });
  const mapping = detectColumns(rows, headers, dataStart);
  const rowCount = rows.slice(dataStart).filter((row) => (row || []).some((cell) => text(cell))).length;
  return { headerIndex, dataStart, headers, mapping, rowCount };
}

/**
 * Até `limit` valores distintos e não vazios de uma coluna, na ordem em que
 * aparecem — usados para mostrar ao usuário o que uma coluna realmente
 * contém, em vez de ele escolher apenas pelo nome do cabeçalho.
 */
export function sampleColumnValues(rows, column, dataStart, limit = 3) {
  const seen = new Set();
  const samples = [];
  for (let index = dataStart; index < rows.length && samples.length < limit; index += 1) {
    const value = text((rows[index] || [])[column]);
    if (!value || seen.has(value)) continue;
    seen.add(value);
    samples.push(value);
  }
  return samples;
}

/**
 * Quantos registros o mapeamento atual da aba produziria — a mesma regra de
 * `recordFromRow`, mas contando em vez de montar cada registro. Serve para
 * avisar cedo quando uma coluna mapeada não gera nenhum documento, sem
 * esperar o cruzamento rodar.
 *
 * @param profile perfil completo da aba, incluindo `rows` (a matriz lida do
 *   arquivo) além de `headers`, `mapping`, `dataStart` e `startRow`.
 */
export function countMappedRecords(profile) {
  if (!profile || profile.mapping.document < 0) return 0;
  let count = 0;
  for (let index = profile.dataStart; index < profile.rows.length; index += 1) {
    if (recordFromRow(profile, profile.rows[index] || [], index)) count += 1;
  }
  return count;
}

/**
 * Converte uma linha da planilha no registro usado pelo cruzamento.
 * Retorna `null` quando a linha não tem documento ou apenas repete o cabeçalho.
 *
 * @param profile perfil da aba (`headers`, `mapping`, `startRow`).
 * @param row linha bruta da planilha.
 * @param index índice da linha dentro da matriz da aba.
 */
export function recordFromRow(profile, row = [], index = 0) {
  const { headers = [], mapping, startRow = 0 } = profile;
  if (!mapping || mapping.document < 0) return null;
  const document = text(row[mapping.document]);
  if (!document) return null;
  if (normalizedHeader(document) === normalizedHeader(headers[mapping.document])) return null;

  const extras = {};
  for (const column of mapping.extras || []) {
    const value = text(row[column]);
    if (value) extras[headers[column] || `Coluna ${column + 1}`] = value;
  }
  const rawDate = mapping.date >= 0 ? row[mapping.date] : "";
  return {
    document,
    status: mapping.status >= 0 ? text(row[mapping.status]) : "",
    date: text(rawDate),
    dateValue: parseDateValue(rawDate),
    extras,
    rowNumber: startRow + index + 1,
  };
}
