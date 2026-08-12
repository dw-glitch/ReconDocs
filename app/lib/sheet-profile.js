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

export function detectHeaderRow(rows, limit = 40) {
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
  const maxColumns = Math.max(1, ...rows.slice(0, 400).map((row) => (row || []).length));
  const headers = Array.from({ length: maxColumns }, (_, index) => {
    const raw = hasHeader ? text(rows[headerIndex]?.[index]) : "";
    return raw || columnLabel(index);
  });
  const mapping = detectColumns(rows, headers, dataStart);
  const rowCount = rows.slice(dataStart).filter((row) => (row || []).some((cell) => text(cell))).length;
  return { headerIndex, dataStart, headers, mapping, rowCount };
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
  return {
    document,
    status: mapping.status >= 0 ? text(row[mapping.status]) : "",
    date: mapping.date >= 0 ? text(row[mapping.date]) : "",
    extras,
    rowNumber: startRow + index + 1,
  };
}
