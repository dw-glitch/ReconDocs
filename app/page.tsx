"use client";

import { Fragment, useDeferredValue, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import {
  AlertTriangle,
  ArrowLeftRight,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  CircleCheck,
  CircleHelp,
  Download,
  FileSpreadsheet,
  Filter,
  LoaderCircle,
  LockKeyhole,
  RotateCcw,
  Search,
  Settings2,
  UploadCloud,
  X,
} from "lucide-react";
import {
  analyzeDatasets,
  canonicalId,
  documentKind,
  etDocumentSubtype,
  filterResults,
  norm,
  parseDocumentReference,
  SOURCE_LABELS,
  summarizeResults,
  text,
} from "./lib/analysis-engine";

type Role = "sgp" | "sigem" | "planned";
type DocumentScope = "all" | "et" | "cv" | "n1710";
type EtScope = "all" | "rir" | "cm";
type FieldName = "document" | "title" | "revision" | "status" | "allocationStatus" | "allocationNumber";
type Mapping = Record<FieldName, number>;

type SheetProfile = {
  name: string;
  rows: unknown[][];
  startRow: number;
  startCol: number;
  headerIndex: number;
  dataStart: number;
  headers: string[];
  mapping: Mapping;
  rowCount: number;
};

type UploadData = {
  role: Role;
  fileName: string;
  fileSize: number;
  sheets: SheetProfile[];
  selectedSheet: number;
};

type Progress = { value: number; label: string };

type SpreadsheetRecord = {
  code: string;
  title?: string;
  revision?: string;
  status?: string;
  allocationStatus?: string;
  allocationNumber?: string;
  allocationStatusHeader?: string;
  fileName: string;
  sheetName: string;
  rowNumber: number;
  referenceValue?: string;
  embeddedRevision?: string;
  mappedRevision?: string;
  revisionConflict?: boolean;
};

type SourceSummary = {
  present: boolean;
  count: number;
  codes: string[];
  ambiguous: boolean;
  record: SpreadsheetRecord | null;
};

type FieldComparison = {
  label: string;
  kind: "equal" | "missing" | "different";
  sgp: string;
  sigem: string;
};

type SourceFieldComparison = {
  label: string;
  kind: "equal" | "missing" | "different" | "not_comparable";
  values: { sgp: string; sigem: string; planned: string };
  detail: string;
};

type AnalysisItem = {
  id: string;
  displayCode: string;
  sgp: SourceSummary;
  sigem: SourceSummary;
  planned: SourceSummary;
  presence: { kind: string; label: string };
  allocation: { kind: string; label: string; reason: string; evidence: string };
  posting: { kind: string; label: string; detail: string; postedPlannedCodes: string[]; notPostedPlannedCodes: string[] };
  documentTag: string;
  comparisons: FieldComparison[];
  differences: FieldComparison[];
  fieldComparisons: Record<"code" | "revision" | "title" | "status", SourceFieldComparison>;
  documentKind: Exclude<DocumentScope, "all"> | "unknown";
  nt: { kind: string; label: string; explanation: string };
  ambiguous: boolean;
  needsReview: boolean;
};

type AnalysisMetrics = {
  total: number;
  inBoth: number;
  onlySgp: number;
  onlySigem: number;
  onlyPlanned: number;
  allocated: number;
  notAllocated: number;
  missingPlanned: number;
  review: number;
  ntAlternates: number;
  posted: number;
  notPosted: number;
  codeChanged: number;
  bothForms: number;
};

type AnalysisOutput = { results: AnalysisItem[]; metrics: AnalysisMetrics };

const EMPTY_MAPPING: Mapping = {
  document: -1,
  title: -1,
  revision: -1,
  status: -1,
  allocationStatus: -1,
  allocationNumber: -1,
};

const ALIASES: Record<FieldName, string[]> = {
  document: [
    "DOCUMENTO", "NOME DOCUMENTO", "NOMEDOCUMENTO", "CODIGO", "CODIGO DO DOCUMENTO",
    "CODIGO DOCUMENTAL", "NUMERO DO DOCUMENTO", "N DOCUMENTO", "DOCUMENT NUMBER", "DOC NUMBER",
    "DOCUMENT ID", "DOC", "ARQUIVO", "NOME DO ARQUIVO", "NOME DO CHECKLIST",
  ],
  title: ["TITULO", "TITULO DO DOCUMENTO", "DESCRICAO", "DESCRICAO DO DOCUMENTO", "DOCUMENT TITLE"],
  revision: ["REVISAO", "REV", "REVISAO ATUAL", "INDICE DE REVISAO"],
  status: ["STATUS SIGEM", "STATUS", "STATUS DA INSPECAO", "SITUACAO", "SITUACAO SIGEM", "ESTADO"],
  allocationStatus: [
    "CONFIRMACAO DE DOCUMENTOS PREVISTOS", "CONFIRMACAO DOCUMENTOS PREVISTOS",
    "STATUS DA ALOCACAO", "SITUACAO DA ALOCACAO", "ALOCADO", "ALOCADO?", "CONFIRMACAO",
  ],
  allocationNumber: ["ALOCACAO", "NUMERO DA ALOCACAO", "N DA ALOCACAO", "CODIGO DA ALOCACAO"],
};

const FILE_META: Record<Role, { number: string; title: string; subtitle: string }> = {
  sgp: { number: "01", title: "Lista do SGP", subtitle: "Relação de documentos cadastrados no SGP" },
  sigem: { number: "02", title: "Consulta Geral do SIGEM", subtitle: "Exportação completa da consulta geral" },
  planned: { number: "03", title: "Documentos previstos", subtitle: "Mostra como o tag foi alocado e se essa forma foi postada" },
};

const FIELD_LABELS: Record<FieldName, string> = {
  document: "Código do documento (sem revisão)",
  title: "Título",
  revision: "Revisão",
  status: "Status",
  allocationStatus: "Confirmação de alocação",
  allocationNumber: "Número da alocação",
};

const SCOPE_OPTIONS: { key: DocumentScope; label: string; description: string }[] = [
  { key: "all", label: "Todos", description: "ET, CV e N-1710" },
  { key: "et", label: "Somente ET", description: "Relatórios codificados pela ET" },
  { key: "cv", label: "Somente CV", description: "Currículos com grupo CV" },
  { key: "n1710", label: "Somente N-1710", description: "Documentos técnicos de projeto" },
];

const SCOPE_LABELS: Record<DocumentScope, string> = {
  all: "Todos os documentos",
  et: "Somente documentos ET",
  cv: "Somente documentos CV",
  n1710: "Somente documentos N-1710",
};

const ET_SCOPE_OPTIONS: { key: EtScope; label: string }[] = [
  { key: "all", label: "Todos os ET" },
  { key: "rir", label: "Doc RIR" },
  { key: "cm", label: "Doc de C&M" },
];

function scopeLabel(scope: DocumentScope, etScope: EtScope) {
  if (scope !== "et" || etScope === "all") return SCOPE_LABELS[scope];
  return etScope === "rir" ? "Somente documentos ET · Doc RIR" : "Somente documentos ET · Doc de C&M";
}

const FILTERS = [
  ["all", "Todos"],
  ["not_posted", "Não postados"],
  ["changed", "Código alterado"],
  ["not_allocated", "Não alocados"],
] as const;

function formatNumber(value: number) {
  return new Intl.NumberFormat("pt-BR").format(value);
}

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1).replace(".", ",")} MB`;
}

function normalizedHeader(value: unknown) {
  return norm(value).replace(/[^A-Z0-9]+/g, " ").trim();
}

function aliasScore(header: unknown, field: FieldName) {
  const key = normalizedHeader(header);
  if (!key) return 0;
  if (field === "document") {
    if (["CODIGO DO DOCUMENTO", "CODIGO DOCUMENTAL", "NUMERO DO DOCUMENTO", "DOCUMENT NUMBER", "DOCUMENT ID", "NOME DO CHECKLIST"].includes(key)) return 30;
    if (key.includes("DOCUMENTO DE REFERENCIA") || key.includes("REFERENCIA DO DOCUMENTO")) return 2;
  }
  const aliases = ALIASES[field];
  if (aliases.includes(key)) return 12;
  if (aliases.some((alias) => key.includes(alias) || alias.includes(key))) return 5;
  return 0;
}

function looksLikeDocument(value: unknown) {
  const key = canonicalId(value);
  if (key.length < 7) return false;
  return /_RNEST_|5290\.00|\d{4,}/.test(key) || (key.includes("_") && /[A-Z]/.test(key) && /\d/.test(key));
}

function detectMapping(headers: string[], role: Role, fallbackDocument = -1): Mapping {
  const mapping = { ...EMPTY_MAPPING };
  const wanted: FieldName[] = role === "planned"
    ? ["document"]
    : ["document", "title", "revision", "status"];
  for (const field of wanted) {
    let best = { index: -1, score: 0 };
    headers.forEach((header, index) => {
      const score = aliasScore(header, field);
      if (score > best.score) best = { index, score };
    });
    mapping[field] = best.index;
  }
  if (mapping.document < 0) mapping.document = fallbackDocument;
  return mapping;
}

function profileSheet(sheetName: string, sheet: XLSX.WorkSheet, role: Role): SheetProfile {
  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: true,
  }) as unknown[][];

  let bestHeader = { index: -1, score: 0 };
  const headerFields: FieldName[] = role === "planned"
    ? ["document"]
    : role === "sgp"
      ? ["document", "revision", "status"]
      : ["document", "title", "revision", "status"];
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 50); rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const documentScore = Math.max(0, ...row.map((cell) => aliasScore(cell, "document")));
    const fieldScore = headerFields
      .filter((field) => field !== "document")
      .reduce((sum, field) => sum + Math.max(0, ...row.map((cell) => aliasScore(cell, field))), 0);
    const nonEmptyBonus = Math.min(40, row.filter((cell) => text(cell)).length);
    const filterPenalty = row.filter((cell) => /%|:\s*$/.test(text(cell))).length * 8;
    const score = documentScore * 10 + fieldScore + nonEmptyBonus - filterPenalty;
    if (score > bestHeader.score) bestHeader = { index: rowIndex, score };
  }

  const hasHeader = bestHeader.score >= 100;
  const headerIndex = hasHeader ? bestHeader.index : -1;
  const dataStart = hasHeader ? headerIndex + 1 : 0;
  const maxColumns = Math.max(1, ...rows.slice(0, 250).map((row) => row.length));
  const headers = Array.from({ length: maxColumns }, (_, index) => {
    const rawHeader = hasHeader ? text(rows[headerIndex]?.[index]) : "";
    return rawHeader || `Coluna ${XLSX.utils.encode_col(range.s.c + index)}`;
  });

  let fallbackDocument = -1;
  let bestDocumentScore = -1;
  for (let column = 0; column < maxColumns; column += 1) {
    const sample = rows.slice(dataStart, dataStart + 300).map((row) => row[column]).filter((value) => text(value));
    if (!sample.length) continue;
    const documentHits = sample.filter(looksLikeDocument).length;
    const score = documentHits / sample.length + Math.min(sample.length, 50) / 10000;
    if (score > bestDocumentScore) {
      bestDocumentScore = score;
      fallbackDocument = column;
    }
  }

  return {
    name: sheetName,
    rows,
    startRow: range.s.r,
    startCol: range.s.c,
    headerIndex,
    dataStart,
    headers,
    mapping: detectMapping(headers, role, fallbackDocument),
    rowCount: Math.max(0, rows.length - dataStart),
  };
}

async function readSpreadsheet(file: File, role: Role): Promise<UploadData> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: false });
  const sheets = workbook.SheetNames.map((name) => profileSheet(name, workbook.Sheets[name], role));
  const selectedSheet = Math.max(0, sheets.findIndex((sheet) => sheet.mapping.document >= 0 && sheet.rowCount > 0));
  return { role, fileName: file.name, fileSize: file.size, sheets, selectedSheet };
}

function rowToRecord(upload: UploadData, row: unknown[], rowIndex: number) {
  const sheet = upload.sheets[upload.selectedSheet];
  const map = sheet.mapping;
  const value = (field: FieldName) => map[field] >= 0 ? row[map[field]] : "";
  const referenceValue = text(value("document"));
  const parsedReference = parseDocumentReference(referenceValue, value("revision"));
  const code = parsedReference.code;
  if (!code || normalizedHeader(code) === "DOCUMENTO" || norm(code) === "FIM" || !looksLikeDocument(code)) return null;
  return {
    code,
    title: text(value("title")),
    revision: parsedReference.revision,
    status: text(value("status")),
    allocationStatus: text(value("allocationStatus")),
    allocationNumber: text(value("allocationNumber")),
    allocationStatusHeader: map.allocationStatus >= 0 ? sheet.headers[map.allocationStatus] : "Confirmação de alocação",
    fileName: upload.fileName,
    sheetName: sheet.name,
    rowNumber: sheet.startRow + rowIndex + 1,
    referenceValue,
    embeddedRevision: parsedReference.embeddedRevision,
    mappedRevision: parsedReference.mappedRevision,
    revisionConflict: parsedReference.revisionConflict,
  };
}

async function recordsFromUpload(upload: UploadData, onProgress?: (done: number, total: number) => void) {
  const sheet = upload.sheets[upload.selectedSheet];
  const result = [];
  const rows = sheet.rows;
  for (let index = sheet.dataStart; index < rows.length; index += 1) {
    const record = rowToRecord(upload, rows[index] || [], index);
    if (record) result.push(record);
    if (index > 0 && index % 5000 === 0) {
      onProgress?.(index, rows.length);
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    }
  }
  return result;
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function createBrandLogoDataUrl() {
  const canvas = document.createElement("canvas");
  canvas.width = 820;
  canvas.height = 170;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Não foi possível preparar a identidade visual do relatório.");

  context.clearRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#153A5C";
  context.beginPath();
  context.roundRect(16, 18, 132, 132, 26);
  context.fill();

  context.strokeStyle = "#FFFFFF";
  context.lineWidth = 10;
  context.lineCap = "round";
  context.lineJoin = "round";
  context.beginPath();
  context.moveTo(45, 66);
  context.lineTo(116, 66);
  context.moveTo(45, 66);
  context.lineTo(61, 50);
  context.moveTo(45, 66);
  context.lineTo(61, 82);
  context.stroke();

  context.beginPath();
  context.moveTo(116, 101);
  context.lineTo(45, 101);
  context.moveTo(116, 101);
  context.lineTo(100, 85);
  context.moveTo(116, 101);
  context.lineTo(100, 117);
  context.stroke();

  context.fillStyle = "#0C7657";
  context.beginPath();
  context.arc(122, 126, 25, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = "#FFFFFF";
  context.lineWidth = 6;
  context.beginPath();
  context.moveTo(110, 126);
  context.lineTo(119, 135);
  context.lineTo(136, 115);
  context.stroke();

  context.fillStyle = "#153A5C";
  context.font = "700 64px Arial, sans-serif";
  context.fillText("RECON", 178, 94);
  const reconWidth = context.measureText("RECON").width;
  context.fillStyle = "#0C7657";
  context.fillText("DOCS", 178 + reconWidth, 94);
  context.fillStyle = "#657B8D";
  context.font = "600 22px Arial, sans-serif";
  context.letterSpacing = "3px";
  context.fillText("CONFERÊNCIA SGP × SIGEM", 182, 132);
  return canvas.toDataURL("image/png");
}

function addReportBranding(
  sheet: import("exceljs").Worksheet,
  logoImageId: number,
  endColumn: string,
  title: string,
) {
  sheet.addImage(logoImageId, {
    tl: { col: 0.15, row: 0.18 },
    ext: { width: 218, height: 45 },
  });
  sheet.mergeCells(`D1:${endColumn}1`);
  sheet.getCell("D1").value = title;
  sheet.getCell("D1").font = { bold: true, size: 16, color: { argb: "153A5C" } };
  sheet.getCell("D1").alignment = { vertical: "middle", horizontal: "left" };
  sheet.mergeCells(`D2:${endColumn}2`);
  sheet.getCell("D2").value = "Conferência entre SGP, Consulta Geral do SIGEM e Documentos Previstos";
  sheet.getCell("D2").font = { size: 10, color: { argb: "52687B" } };
  sheet.mergeCells(`D3:${endColumn}3`);
  sheet.getCell("D3").value = `Relatório gerado em ${new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(new Date())}`;
  sheet.getCell("D3").font = { size: 9, italic: true, color: { argb: "7B8F9D" } };
  sheet.getRow(1).height = 29;
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 18;
  sheet.getRow(4).height = 8;
}

function reportRow(item: AnalysisItem) {
  return {
    tag: item.documentTag,
    codigoSgp: item.sgp.codes.join(" | "),
    revisaoSgp: item.fieldComparisons.revision.values.sgp,
    codigoAlocado: item.planned.codes.join(" | "),
    alocado: item.allocation.label,
    codigoSigem: item.sigem.codes.join(" | "),
    revisaoSigem: item.fieldComparisons.revision.values.sigem,
    situacaoPostagem: `${item.posting.label}. ${item.posting.detail}`.trim(),
    diferenca: `${item.fieldComparisons.code.detail}; revisão: ${item.fieldComparisons.revision.detail}`,
    formaNt: item.nt.label,
  };
}

async function exportWorkbook(results: AnalysisItem[], metrics: AnalysisMetrics, scope: DocumentScope, etScope: EtScope) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ReconDocs — Conferência SGP x SIGEM";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject = "Conferência de documentos e presença em Documentos Previstos";
  workbook.title = "ReconDocs — Relatório de Conferência SGP x SIGEM";

  const navy = "153A5C";
  const teal = "0C7657";
  const amber = "A56812";
  const red = "A64035";
  const pale = "F4F7F9";
  const white = "FFFFFF";
  const logoImageId = workbook.addImage({ base64: createBrandLogoDataUrl(), extension: "png" });

  const summary = workbook.addWorksheet("Resumo", { views: [{ state: "frozen", ySplit: 5 }] });
  summary.columns = [24, 16, 24, 16, 24, 16].map((width) => ({ width }));
  addReportBranding(summary, logoImageId, "F", "RESUMO DA CONFERÊNCIA");
  summary.getRow(5).values = ["Indicador", "Quantidade", "Indicador", "Quantidade", "Indicador", "Quantidade"];
  summary.getRow(6).values = ["Tags analisados", metrics.total, "Postados no SIGEM", metrics.posted, "Alocados", metrics.allocated];
  summary.getRow(7).values = ["Não postados", metrics.notPosted, "Código alterado", metrics.codeChanged, "Duas formas postadas", metrics.bothForms];
  summary.getRow(8).values = ["Somente SGP", metrics.onlySgp, "Somente SIGEM", metrics.onlySigem, "Não alocados", metrics.notAllocated];
  summary.getRow(9).values = ["Encontrados nas duas formas de nt-", metrics.ntAlternates, "", "", "", ""];
  summary.getRow(10).values = ["Escopo analisado", scopeLabel(scope, etScope), "", "", "", ""];
  summary.mergeCells("B10:F10");
  summary.getCell("B10").alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  summary.getRow(10).height = 27;
  summary.getRow(5).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: white } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
  });
  summary.getRow(5).height = 25;

  const columns = [
    { header: "TAG DO DOCUMENTO", key: "tag", width: 34 },
    { header: "CÓDIGO NO SGP", key: "codigoSgp", width: 48 },
    { header: "REVISÃO SGP", key: "revisaoSgp", width: 14 },
    { header: "CÓDIGO ALOCADO (DOCUMENTOS PREVISTOS)", key: "codigoAlocado", width: 52 },
    { header: "ALOCADO?", key: "alocado", width: 16 },
    { header: "CÓDIGO POSTADO NO SIGEM", key: "codigoSigem", width: 48 },
    { header: "REVISÃO SIGEM", key: "revisaoSigem", width: 14 },
    { header: "SITUAÇÃO DA POSTAGEM", key: "situacaoPostagem", width: 72 },
    { header: "DIFERENÇA SGP × SIGEM", key: "diferenca", width: 72 },
    { header: "FORMA COM/SEM nt-", key: "formaNt", width: 28 },
  ];

  const addDataSheet = (name: string, items: AnalysisItem[]) => {
    const sheet = workbook.addWorksheet(name, { views: [{ state: "frozen", ySplit: 5, xSplit: 1 }] });
    sheet.columns = columns.map(({ key, width }) => ({ key, width }));
    addReportBranding(sheet, logoImageId, "J", name.toUpperCase());
    const header = sheet.getRow(5);
    header.values = columns.map((column) => column.header);
    header.height = 30;
    header.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: white } };
      const differenceColumns = new Set([8, 9]);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: differenceColumns.has(cell.col) ? amber : navy } };
      cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    });
    sheet.addRows(items.map(reportRow));
    sheet.autoFilter = { from: "A5", to: "J5" };
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 5) return;
      row.alignment = { vertical: "top", wrapText: true };
      const allocationCell = row.getCell(5);
      const postingCell = row.getCell(8);
      const allocationKey = norm(allocationCell.value);
      const postingKey = norm(postingCell.value);
      const allocationColor = allocationKey === "ALOCADO" ? teal : allocationKey === "NAO ALOCADO" ? red : amber;
      allocationCell.font = { bold: true, color: { argb: allocationColor } };
      postingCell.font = { bold: true, color: { argb: postingKey.startsWith("NAO POSTADO") ? red : postingKey.includes("ALTERADO") || postingKey.includes("DUAS FORMAS") ? amber : teal } };
      for (const columnIndex of [9]) {
        const differenceCell = row.getCell(columnIndex);
        const noDifference = norm(differenceCell.value) === "SEM DIFERENCA";
        differenceCell.font = { bold: !noDifference, color: { argb: noDifference ? teal : amber } };
      }
      if ((rowNumber - 5) % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: pale } };
    });
    return sheet;
  };

  addDataSheet("Conferência", results);

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `ReconDocs_Conferencia_SGP_SIGEM_${scope}${scope === "et" ? `-${etScope}` : ""}_${new Date().toISOString().slice(0, 10)}.xlsx`,
  );
}

function MappingSelect({ profile, field, onChange }: { profile: SheetProfile; field: FieldName; onChange: (index: number) => void }) {
  return (
    <label className="mapping-field">
      <span>{FIELD_LABELS[field]}{field === "document" ? " *" : ""}</span>
      <select value={profile.mapping[field]} onChange={(event) => onChange(Number(event.target.value))}>
        <option value={-1}>Não usar</option>
        {profile.headers.map((header, index) => <option key={`${header}-${index}`} value={index}>{header}</option>)}
      </select>
    </label>
  );
}

function UploadCard({ role, value, busy, onUpload, onRemove, onSheet, onMapping }: {
  role: Role;
  value?: UploadData;
  busy: boolean;
  onUpload: (file: File) => void;
  onRemove: () => void;
  onSheet: (sheet: number) => void;
  onMapping: (field: FieldName, index: number) => void;
}) {
  const [dragging, setDragging] = useState(false);
  const [mappingOpen, setMappingOpen] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const meta = FILE_META[role];
  const profile = value?.sheets[value.selectedSheet];
  const fields: FieldName[] = role === "planned"
    ? ["document"]
    : role === "sgp"
      ? ["document", "revision", "status"]
      : ["document", "title", "revision", "status"];

  const receive = (file?: File) => {
    if (file) onUpload(file);
  };

  return (
    <article className={`upload-card ${value ? "has-file" : ""}`}>
      <div className="upload-card-heading">
        <span className="step-number">{meta.number}</span>
        <div>
          <h3>{meta.title}</h3>
          <p>{meta.subtitle}</p>
        </div>
      </div>

      {!value ? (
        <button
          type="button"
          className={`dropzone ${dragging ? "is-dragging" : ""}`}
          onClick={() => input.current?.click()}
          onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => { event.preventDefault(); setDragging(false); receive(event.dataTransfer.files[0]); }}
          disabled={busy}
        >
          <UploadCloud size={23} strokeWidth={1.8} />
          <span>Arraste a planilha ou <strong>selecione</strong></span>
          <small>.xlsx, .xls ou .csv</small>
        </button>
      ) : (
        <div className="file-loaded">
          <div className="file-summary">
            <span className="file-icon"><FileSpreadsheet size={22} /></span>
            <div>
              <strong title={value.fileName}>{value.fileName}</strong>
              <span>{formatBytes(value.fileSize)} · {formatNumber(profile?.rowCount || 0)} linhas</span>
            </div>
            <button type="button" className="icon-button" onClick={onRemove} aria-label={`Remover ${meta.title}`}><X size={17} /></button>
          </div>
          <div className="detected-line">
            {profile?.mapping.document !== -1 ? <CircleCheck size={16} /> : <AlertTriangle size={16} />}
            <span>{profile?.mapping.document !== -1 ? `Código detectado: ${profile?.headers[profile.mapping.document]}` : "Selecione a coluna do código"}</span>
          </div>
          <button type="button" className="mapping-toggle" onClick={() => setMappingOpen((open) => !open)}>
            <Settings2 size={15} /> Conferir colunas <ChevronDown size={15} className={mappingOpen ? "rotate" : ""} />
          </button>
          {mappingOpen && profile && (
            <div className="mapping-panel">
              {value.sheets.length > 1 && (
                <label className="mapping-field">
                  <span>Aba analisada</span>
                  <select value={value.selectedSheet} onChange={(event) => onSheet(Number(event.target.value))}>
                    {value.sheets.map((sheet, index) => <option key={sheet.name} value={index}>{sheet.name} · {formatNumber(sheet.rowCount)} linhas</option>)}
                  </select>
                </label>
              )}
              {fields.map((field) => <MappingSelect key={field} profile={profile} field={field} onChange={(index) => onMapping(field, index)} />)}
            </div>
          )}
        </div>
      )}
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xls,.xlsm,.csv,.tsv"
        hidden
        onChange={(event) => receive(event.target.files?.[0])}
      />
    </article>
  );
}

function StatusBadge({ kind, label }: { kind: string; label: string }) {
  return <span className={`status-badge status-${kind}`}>{label}</span>;
}

function differenceCount(item: AnalysisItem) {
  return [item.fieldComparisons.code, item.fieldComparisons.revision]
    .filter((comparison) => comparison.kind === "different" || comparison.kind === "missing").length;
}

function DetailPanel({ item }: { item: AnalysisItem }) {
  return (
    <div className="detail-panel">
      <div className="detail-block">
        <span className="detail-label">Situação da postagem</span>
        <strong>{item.posting.label}</strong>
        <p>{item.posting.detail}</p>
      </div>
      <div className="detail-block detail-sources">
        <span className="detail-label">Códigos por fonte</span>
        <p><b>SGP:</b> {item.sgp.codes.join(" | ") || "Não localizado"}</p>
        <p><b>SIGEM:</b> {item.sigem.codes.join(" | ") || "Não localizado"}</p>
        <p><b>Alocado:</b> {item.planned.codes.join(" | ") || "Não localizado"}</p>
      </div>
      <div className="detail-block">
        <span className="detail-label">Revisão por fonte</span>
        <p><b>SGP:</b> {item.fieldComparisons.revision.values.sgp || "vazio"}</p>
        <p><b>SIGEM:</b> {item.fieldComparisons.revision.values.sigem || "vazio"}</p>
        <p className={item.fieldComparisons.revision.kind !== "equal" ? "difference-text" : ""}><b>Diferença:</b> {item.fieldComparisons.revision.detail}</p>
      </div>
      <div className="detail-block detail-wide">
        <span className="detail-label">Diferenças essenciais</span>
        {[item.fieldComparisons.code, item.fieldComparisons.revision].map((comparison) => (
          <p key={comparison.label} className={comparison.kind === "different" || comparison.kind === "missing" ? "difference-text" : ""}>
            <b>{comparison.label}:</b> {comparison.detail}
          </p>
        ))}
        <p><b>Forma nt-:</b> {item.nt.label}. {item.nt.explanation}</p>
      </div>
    </div>
  );
}

export default function Home() {
  const [uploads, setUploads] = useState<Partial<Record<Role, UploadData>>>({});
  const [busyRole, setBusyRole] = useState<Role | null>(null);
  const [error, setError] = useState("");
  const [progress, setProgress] = useState<Progress | null>(null);
  const [analysis, setAnalysis] = useState<AnalysisOutput | null>(null);
  const [documentScope, setDocumentScope] = useState<DocumentScope>("all");
  const [completedScope, setCompletedScope] = useState<DocumentScope>("all");
  const [etScope, setEtScope] = useState<EtScope>("all");
  const [completedEtScope, setCompletedEtScope] = useState<EtScope>("all");
  const [activeFilter, setActiveFilter] = useState("all");
  const [query, setQuery] = useState("");
  const deferredQuery = useDeferredValue(query);
  const [page, setPage] = useState(1);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const PAGE_SIZE = 50;

  const ready = (["sgp", "sigem", "planned"] as Role[]).every((role) => {
    const upload = uploads[role];
    return upload && upload.sheets[upload.selectedSheet]?.mapping.document >= 0;
  });

  const filtered = useMemo<AnalysisItem[]>(
    () => analysis ? filterResults(analysis.results, activeFilter, deferredQuery) as AnalysisItem[] : [],
    [analysis, activeFilter, deferredQuery],
  );
  const pageCount = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const visibleRows = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const changeFilter = (filter: string) => {
    setActiveFilter(filter);
    setPage(1);
    setExpanded(null);
  };

  const handleFile = async (role: Role, file: File) => {
    setError("");
    if (!/\.(xlsx|xls|xlsm|csv|tsv)$/i.test(file.name)) {
      setError("Use uma planilha .xlsx, .xls, .xlsm, .csv ou .tsv.");
      return;
    }
    setBusyRole(role);
    try {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const parsed = await readSpreadsheet(file, role);
      if (!parsed.sheets.length) throw new Error("A planilha não possui abas legíveis.");
      setUploads((current) => ({ ...current, [role]: parsed }));
      setAnalysis(null);
    } catch (cause) {
      setError(cause instanceof Error ? `Não foi possível ler ${file.name}: ${cause.message}` : `Não foi possível ler ${file.name}.`);
    } finally {
      setBusyRole(null);
    }
  };

  const updateUpload = (role: Role, transform: (upload: UploadData) => UploadData) => {
    setUploads((current) => current[role] ? { ...current, [role]: transform(current[role]!) } : current);
    setAnalysis(null);
  };

  const updateSheet = (role: Role, selectedSheet: number) => updateUpload(role, (upload) => ({ ...upload, selectedSheet }));
  const updateMapping = (role: Role, field: FieldName, index: number) => updateUpload(role, (upload) => ({
    ...upload,
    sheets: upload.sheets.map((sheet, sheetIndex) => sheetIndex === upload.selectedSheet
      ? { ...sheet, mapping: { ...sheet.mapping, [field]: index } }
      : sheet),
  }));

  const runAnalysis = async () => {
    if (!ready) return;
    setError("");
    setProgress({ value: 4, label: "Preparando as três bases" });
    setAnalysis(null);
    try {
      const datasets: Record<Role, SpreadsheetRecord[]> = { sgp: [], sigem: [], planned: [] };
      const roles: Role[] = ["sgp", "sigem", "planned"];
      for (let roleIndex = 0; roleIndex < roles.length; roleIndex += 1) {
        const role = roles[roleIndex];
        setProgress({ value: 8 + roleIndex * 25, label: `Lendo ${SOURCE_LABELS[role]}` });
        const records = await recordsFromUpload(uploads[role]!, (done, total) => {
          const base = 8 + roleIndex * 25;
          setProgress({ value: base + Math.round((done / total) * 20), label: `Lendo ${SOURCE_LABELS[role]} · ${formatNumber(done)} linhas` });
        });
        datasets[role] = records;
      }
      setProgress({ value: 88, label: "Ligando os códigos pelo mesmo tag" });
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const completeOutput = analyzeDatasets(datasets) as AnalysisOutput;
      const scopedResults = completeOutput.results.filter((item) => {
        if (documentScope === "all") return true;
        if (item.documentKind !== documentScope) return false;
        return documentScope !== "et" || etScope === "all" || etDocumentSubtype(item.displayCode, item.sgp.record?.title || item.sigem.record?.title) === etScope;
      });
      const output: AnalysisOutput = { results: scopedResults, metrics: summarizeResults(scopedResults) as AnalysisMetrics };
      setProgress({ value: 100, label: "Conferência concluída" });
      setAnalysis(output);
      setCompletedScope(documentScope);
      setCompletedEtScope(etScope);
      setActiveFilter("all");
      setQuery("");
      setPage(1);
      setExpanded(null);
      setTimeout(() => setProgress(null), 450);
      requestAnimationFrame(() => document.getElementById("results")?.scrollIntoView({ behavior: "smooth", block: "start" }));
    } catch (cause) {
      setProgress(null);
      setError(cause instanceof Error ? cause.message : "A análise não pôde ser concluída.");
    }
  };

  const exportReport = async () => {
    if (!analysis) return;
    setExporting(true);
    setError("");
    try {
      await exportWorkbook(analysis.results, analysis.metrics, completedScope, completedEtScope);
    } catch (cause) {
      setError(cause instanceof Error ? `Não foi possível exportar: ${cause.message}` : "Não foi possível exportar o relatório.");
    } finally {
      setExporting(false);
    }
  };

  const reset = () => {
    setUploads({});
    setAnalysis(null);
    setError("");
    setProgress(null);
    setQuery("");
    setPage(1);
    setDocumentScope("all");
    setCompletedScope("all");
    setEtScope("all");
    setCompletedEtScope("all");
  };

  const filterCount = (filter: string) => {
    if (!analysis) return 0;
    return filterResults(analysis.results, filter, "").length;
  };

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#top" aria-label="ReconDocs — início">
          <span className="brand-mark">
            <ArrowLeftRight size={21} />
            <span className="brand-check"><Check size={10} strokeWidth={3.2} /></span>
          </span>
          <span className="brand-copy"><b><span>RECON</span><em>DOCS</em></b><small>CONFERÊNCIA SGP × SIGEM</small></span>
        </a>
        <div className="privacy"><LockKeyhole size={15} /> Processamento local e seguro</div>
      </header>

      <section className="hero" id="top">
        <div className="hero-copy">
          <span className="eyebrow">RECONDOCS · CONFERÊNCIA DOCUMENTAL</span>
          <h1>Descubra como cada tag foi alocado e qual forma chegou ao SIGEM.</h1>
          <p>O ReconDocs liga códigos diferentes pelo mesmo <b>tag do documento</b>, preserva as revisões e mostra se nenhuma, uma ou as duas formas foram postadas.</p>
        </div>
        <aside className="method-card">
          <CircleCheck size={20} />
          <div><b>Correspondência pelo tag</b><span>Identifica alterações de código entre revisões, inclusive a diferença com/sem nt-.</span></div>
        </aside>
      </section>

      <section className="workspace-section" aria-labelledby="upload-title">
        <div className="section-heading">
          <div>
            <span className="section-kicker">ENTRADAS</span>
            <h2 id="upload-title">Carregue as três fontes</h2>
          </div>
          {Object.keys(uploads).length > 0 && <button type="button" className="quiet-button" onClick={reset}><RotateCcw size={15} /> Limpar tudo</button>}
        </div>

        <div className="upload-grid">
          {(["sgp", "sigem", "planned"] as Role[]).map((role) => (
            <UploadCard
              key={role}
              role={role}
              value={uploads[role]}
              busy={busyRole === role}
              onUpload={(file) => handleFile(role, file)}
              onRemove={() => { setUploads((current) => { const next = { ...current }; delete next[role]; return next; }); setAnalysis(null); }}
              onSheet={(sheet) => updateSheet(role, sheet)}
              onMapping={(field, index) => updateMapping(role, field, index)}
            />
          ))}
        </div>

        <div className="scope-selector" aria-labelledby="scope-title">
          <div className="scope-heading">
            <span className="section-kicker">ESCOPO DA BUSCA</span>
            <b id="scope-title">Quais documentos deseja comparar?</b>
          </div>
          <div className="scope-options">
            {SCOPE_OPTIONS.map((option) => (
              <button
                key={option.key}
                type="button"
                className={documentScope === option.key ? "active" : ""}
                aria-pressed={documentScope === option.key}
                onClick={() => {
                  setDocumentScope(option.key);
                  if (option.key !== "et") setEtScope("all");
                  setAnalysis(null);
                }}
              >
                <b>{option.label}</b>
                <span>{option.description}</span>
              </button>
            ))}
          </div>
          {documentScope === "et" && (
            <div className="et-scope" aria-label="Tipo de documento ET">
              <span>Dentro de ET:</span>
              {ET_SCOPE_OPTIONS.map((option) => (
                <button
                  key={option.key}
                  type="button"
                  className={etScope === option.key ? "active" : ""}
                  aria-pressed={etScope === option.key}
                  onClick={() => { setEtScope(option.key); setAnalysis(null); }}
                >
                  {option.label}
                </button>
              ))}
            </div>
          )}
        </div>

        {error && <div className="error-banner" role="alert"><AlertTriangle size={18} /><span>{error}</span></div>}

        <div className="action-bar">
          <div className="action-note"><LockKeyhole size={16} /><span>Seus arquivos não são enviados. Toda a leitura acontece neste navegador.</span></div>
          <button type="button" className="primary-button" onClick={runAnalysis} disabled={!ready || Boolean(progress) || Boolean(busyRole)}>
            {progress ? <LoaderCircle size={18} className="spin" /> : <ArrowLeftRight size={18} />}
            {progress ? progress.label : "Comparar as três planilhas"}
          </button>
        </div>
        {progress && <div className="progress-track" aria-label={progress.label}><span style={{ width: `${progress.value}%` }} /></div>}
      </section>

      {analysis && (
        <section className="results-section" id="results">
          <div className="results-heading">
            <div>
              <span className="section-kicker">RESULTADO DA ANÁLISE</span>
              <h2>Conferência concluída</h2>
              <p>{formatNumber(analysis.metrics.total)} tags consolidados · {scopeLabel(completedScope, completedEtScope)}.</p>
            </div>
            <button type="button" className="export-button" onClick={exportReport} disabled={exporting}>
              {exporting ? <LoaderCircle size={17} className="spin" /> : <Download size={17} />}
              {exporting ? "Gerando relatório" : "Exportar relatório Excel"}
            </button>
          </div>

          <div className="metrics-grid">
            <article className="metric-card"><span>TAGS ANALISADOS</span><strong>{formatNumber(analysis.metrics.total)}</strong><small>Universo consolidado</small></article>
            <article className="metric-card metric-good"><span>ALOCADOS</span><strong>{formatNumber(analysis.metrics.allocated)}</strong><small>Confirmados nos previstos</small></article>
            <article className="metric-card metric-good"><span>POSTADOS</span><strong>{formatNumber(analysis.metrics.posted)}</strong><small>Localizados no SIGEM</small></article>
            <article className="metric-card metric-bad"><span>NÃO POSTADOS</span><strong>{formatNumber(analysis.metrics.notPosted)}</strong><small>Ausentes da Consulta Geral</small></article>
            <article className="metric-card metric-warn"><span>CÓDIGO ALTERADO</span><strong>{formatNumber(analysis.metrics.codeChanged)}</strong><small>Mesmo tag, forma diferente</small></article>
          </div>

          <div className="insight-strip">
            <div><b>{formatNumber(analysis.metrics.bothForms)}</b><span>com as duas formas postadas</span></div>
            <div><b>{formatNumber(analysis.metrics.notAllocated)}</b><span>não alocados</span></div>
            <div><b>{formatNumber(analysis.metrics.onlySgp)}</b><span>somente no SGP</span></div>
            <div><b>{formatNumber(analysis.metrics.ntAlternates)}</b><span>com/sem nt- entre as fontes</span></div>
          </div>

          <div className="results-card">
            <div className="table-toolbar">
              <div className="filter-tabs" role="tablist" aria-label="Filtros do relatório">
                {FILTERS.map(([key, label]) => (
                  <button key={key} type="button" role="tab" aria-selected={activeFilter === key} className={activeFilter === key ? "active" : ""} onClick={() => changeFilter(key)}>
                    {label}<span>{formatNumber(filterCount(key))}</span>
                  </button>
                ))}
              </div>
              <label className="search-field"><Search size={17} /><input value={query} onChange={(event) => { setQuery(event.target.value); setPage(1); }} placeholder="Buscar documento" /><span className="sr-only">Buscar documento</span></label>
            </div>

            <div className="table-meta"><span><Filter size={14} /> Exibindo {formatNumber(filtered.length)} resultados</span><small>Clique em uma linha para ver as comparações</small></div>
            <div className="table-scroll">
              <table>
                <thead><tr><th>Tag do documento</th><th>SGP × SIGEM</th><th>Alocado?</th><th>Postagem</th><th>Forma nt-</th><th aria-label="Detalhes" /></tr></thead>
                <tbody>
                  {visibleRows.map((item) => (
                    <Fragment key={item.id}>
                      <tr key={item.id} className={expanded === item.id ? "expanded-row" : ""} onClick={() => setExpanded(expanded === item.id ? null : item.id)} tabIndex={0} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") setExpanded(expanded === item.id ? null : item.id); }}>
                        <td><strong className="document-code">{item.documentTag}</strong><small>{item.documentKind === "et" ? "ET" : item.documentKind === "cv" ? "CV" : "N-1710"} · Rev. SGP: {item.fieldComparisons.revision.values.sgp || "vazio"} · SIGEM: {item.fieldComparisons.revision.values.sigem || "vazio"}</small></td>
                        <td><StatusBadge kind={item.presence.kind} label={item.presence.label} /></td>
                        <td><StatusBadge kind={item.allocation.kind} label={item.allocation.label} /></td>
                        <td><StatusBadge kind={item.posting.kind} label={item.posting.label} /></td>
                        <td><span className={`nt-result nt-${item.nt.kind}`}>{item.nt.label}</span></td>
                        <td><ChevronDown size={17} className={expanded === item.id ? "rotate" : ""} /></td>
                      </tr>
                      {expanded === item.id && <tr key={`${item.id}-details`} className="details-row"><td colSpan={6}><DetailPanel item={item} /></td></tr>}
                    </Fragment>
                  ))}
                  {!visibleRows.length && <tr><td colSpan={6}><div className="empty-state"><CircleHelp size={28} /><b>Nenhum documento neste filtro</b><span>Altere o filtro ou a busca para visualizar outros resultados.</span></div></td></tr>}
                </tbody>
              </table>
            </div>
            <div className="pagination">
              <span>Página {page} de {pageCount}</span>
              <div>
                <button type="button" onClick={() => { setPage((current) => Math.max(1, current - 1)); setExpanded(null); }} disabled={page === 1} aria-label="Página anterior"><ChevronLeft size={17} /></button>
                <button type="button" onClick={() => { setPage((current) => Math.min(pageCount, current + 1)); setExpanded(null); }} disabled={page === pageCount} aria-label="Próxima página"><ChevronRight size={17} /></button>
              </div>
            </div>
          </div>
        </section>
      )}

      <footer><span><b>ReconDocs</b> · Conferência SGP × SIGEM</span><span>Comparação transparente. Decisões rastreáveis.</span></footer>
    </main>
  );
}
