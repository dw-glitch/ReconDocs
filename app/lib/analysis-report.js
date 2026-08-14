import ExcelJS from "exceljs";
import { addReportBranding, createBrandLogoDataUrl } from "./report-branding.js";
import { norm } from "./analysis-engine.js";

const COLORS = {
  navy: "153A5C", teal: "0C7657", amber: "A56812", red: "A64035",
  pale: "F4F7F9", white: "FFFFFF", muted: "657B8D",
};
const FALLBACK_LOGO = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

export const ANALYSIS_REPORT_COLUMNS = [
  { header: "TAG DO DOCUMENTO", key: "tag", width: 35 },
  { header: "CÓDIGO NO SGP", key: "codigoSgp", width: 50 },
  { header: "REVISÃO SGP", key: "revisaoSgp", width: 14 },
  { header: "CÓDIGO POSTADO NO SIGEM", key: "codigoSigem", width: 50 },
  { header: "REVISÃO SIGEM", key: "revisaoSigem", width: 14 },
  { header: "ALOCAÇÃO", key: "alocacao", width: 17 },
  { header: "SITUAÇÃO DA POSTAGEM", key: "postagem", width: 34 },
  { header: "DIFERENÇA IDENTIFICADA", key: "diferenca", width: 72 },
  { header: "FORMA nt-", key: "nt", width: 27 },
];

export function analysisReportRows(results = []) {
  return results.map((item) => ({
    tag: item.documentTag,
    codigoSgp: item.sgp.codes.join(" | "),
    revisaoSgp: item.revisionComparison?.sgp || item.fieldComparisons.revision.values.sgp,
    codigoSigem: item.sigem.codes.join(" | "),
    revisaoSigem: item.revisionComparison?.sigem || item.fieldComparisons.revision.values.sigem,
    alocacao: item.allocation.label,
    postagem: item.posting.label,
    diferenca: item.differenceDetail || `${item.fieldComparisons.code.detail}; revisão: ${item.fieldComparisons.revision.detail}`,
    nt: item.nt.label,
  }));
}

function styleHeader(row) {
  row.height = 31;
  row.eachCell((cell, column) => {
    cell.font = { bold: true, color: { argb: COLORS.white }, size: 10 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: column === 8 ? COLORS.amber : COLORS.navy } };
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  });
}

function styleRows(sheet) {
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= 5) return;
    row.alignment = { vertical: "top", wrapText: true };
    row.height = 30;
    if ((rowNumber - 5) % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.pale } };
    const allocation = row.getCell(6);
    const posting = row.getCell(7);
    const difference = row.getCell(8);
    allocation.font = { bold: true, color: { argb: norm(allocation.value) === "ALOCADO" ? COLORS.teal : COLORS.red } };
    const postingKey = norm(posting.value);
    posting.font = { bold: true, color: { argb: postingKey.startsWith("NAO POSTADO") ? COLORS.red : postingKey.includes("ALTERADO") || postingKey.includes("FORMAS") ? COLORS.amber : COLORS.teal } };
    difference.font = { color: { argb: norm(difference.value).includes("SEM DIFERENCA") ? COLORS.teal : COLORS.amber } };
  });
}

export async function buildAnalysisWorkbook(results, metrics, scopeLabel = "Todos os documentos") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ReconDocs";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "ReconDocs — Conferência SGP x SIGEM";
  workbook.subject = "Comparação documental com alocação por presença em Documentos Previstos";

  const logo = typeof document === "undefined" ? FALLBACK_LOGO : createBrandLogoDataUrl();
  const logoImageId = workbook.addImage({ base64: logo, extension: "png" });
  const summary = workbook.addWorksheet("Resumo", { views: [{ state: "frozen", ySplit: 5 }], properties: { showGridLines: false } });
  summary.columns = [26, 16, 26, 16, 26, 16].map((width) => ({ width }));
  addReportBranding(summary, logoImageId, "F", "RESUMO DA CONFERÊNCIA");
  summary.getRow(5).values = ["INDICADOR", "QUANTIDADE", "INDICADOR", "QUANTIDADE", "INDICADOR", "QUANTIDADE"];
  summary.getRow(6).values = ["Tags analisados", metrics.total, "Postados no SIGEM", metrics.posted, "Alocados", metrics.allocated];
  summary.getRow(7).values = ["Não postados", metrics.notPosted, "Código alterado", metrics.codeChanged, "Diferença de revisão", metrics.revisionChanged || 0];
  summary.getRow(8).values = ["Somente SGP", metrics.onlySgp, "Somente SIGEM", metrics.onlySigem, "Não alocados", metrics.notAllocated];
  summary.getRow(9).values = ["Duas formas nt-", metrics.ntAlternates, "Transição de EAP", metrics.eapChanged || 0, "Revisar", metrics.review];
  summary.getRow(10).values = ["Escopo", scopeLabel];
  summary.mergeCells("B10:F10");
  styleHeader(summary.getRow(5));
  for (let row = 6; row <= 10; row += 1) {
    summary.getRow(row).height = 25;
    summary.getRow(row).alignment = { vertical: "middle", wrapText: true };
  }

  const sheet = workbook.addWorksheet("Conferência", { views: [{ state: "frozen", ySplit: 5, xSplit: 1 }], properties: { showGridLines: false } });
  sheet.columns = ANALYSIS_REPORT_COLUMNS.map(({ key, width }) => ({ key, width }));
  addReportBranding(sheet, logoImageId, "I", "CONFERÊNCIA SGP × SIGEM");
  sheet.getRow(5).values = ANALYSIS_REPORT_COLUMNS.map((column) => column.header);
  styleHeader(sheet.getRow(5));
  sheet.addRows(analysisReportRows(results));
  sheet.autoFilter = { from: "A5", to: "I5" };
  styleRows(sheet);

  return workbook;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export async function exportAnalysisWorkbook(results, metrics, scopeLabel = "Todos os documentos", scopeKey = "all") {
  const workbook = await buildAnalysisWorkbook(results, metrics, scopeLabel);
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `ReconDocs_Conferencia_${scopeKey}_${new Date().toISOString().slice(0, 10)}.xlsx`,
  );
}
