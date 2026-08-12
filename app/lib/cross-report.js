import ExcelJS from "exceljs";
import { norm } from "./analysis-engine.js";
import { MATCH_MODES } from "./crosscheck-engine.js";
import { addReportBranding, REPORT_COLORS } from "./report-branding.js";

export function columnLetter(index) {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const rest = (value - 1) % 26;
    name = String.fromCharCode(65 + rest) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export const CROSS_REPORT_SHEETS = [
  "Resumo Executivo",
  "Resultado Consolidado",
  "Somente Consulta Geral",
  "Somente Documentos Previstos",
  "Divergências",
];

function safeSheetName(name) {
  return name.replace(/[[\]:*?/\\]/g, " ").slice(0, 31).trim() || "Planilha";
}

/**
 * Monta o relatório Excel do cruzamento: Resumo Executivo, Resultado
 * Consolidado, Somente Consulta Geral, Somente Documentos Previstos,
 * Divergências e, quando houver, Planilhas Adicionais.
 *
 * @param output resultado de `crossReference`.
 * @param options `{ logoBase64 }` — a marca é gerada no navegador e injetada aqui.
 */
export function buildCrossWorkbook(output, options = {}) {
  const { navy, teal, amber, red, pale, white } = REPORT_COLORS;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ReconDocs — Cruzamento de planilhas";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.subject = "Cruzamento dinâmico de planilhas de documentos";
  workbook.title = "ReconDocs — Relatório de Cruzamento";
  const subtitle = `Cruzamento de ${output.sources.length} planilha(s): ${output.sources.map((source) => source.label).join(", ")}`;
  const logoImageId = workbook.addImage({ base64: options.logoBase64 || "", extension: "png" });

  const general = output.sources.find((source) => source.id === output.generalId) || null;
  const planned = output.sources.find((source) => source.id === output.plannedId) || null;
  const others = output.sources.filter((source) => source.id !== output.generalId && source.id !== output.plannedId);

  const addDataSheet = (name, columns, rows, highlight) => {
    const sheet = workbook.addWorksheet(safeSheetName(name), { views: [{ state: "frozen", ySplit: 5, xSplit: 1 }] });
    const endColumn = columnLetter(Math.max(3, columns.length - 1));
    sheet.columns = columns.map(({ key, width }) => ({ key, width }));
    addReportBranding(sheet, logoImageId, endColumn, name.toUpperCase(), subtitle);
    const header = sheet.getRow(5);
    header.values = columns.map((column) => column.header);
    header.height = 30;
    header.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: white } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: navy } };
      cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    });
    sheet.addRows(rows);
    sheet.autoFilter = { from: "A5", to: `${endColumn}5` };
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 5) return;
      row.alignment = { vertical: "top", wrapText: true };
      if ((rowNumber - 5) % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: pale } };
      highlight?.(row);
    });
    return sheet;
  };

  // Aba 1 — Resumo Executivo
  const metrics = output.metrics;
  const summary = workbook.addWorksheet("Resumo Executivo", { views: [{ state: "frozen", ySplit: 5 }] });
  summary.columns = [{ width: 46 }, { width: 18 }, { width: 62 }];
  addReportBranding(summary, logoImageId, "C", "RESUMO EXECUTIVO", subtitle);
  summary.getRow(5).values = ["Indicador", "Quantidade", "Detalhe"];
  summary.getRow(5).height = 25;
  summary.getRow(5).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: white } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: teal } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
  });

  const summaryRows = [
    ["Total de documentos analisados", metrics.total, "Documentos distintos em todas as planilhas carregadas"],
    ["Total na Consulta Geral", metrics.totalGeneral, general ? general.label : "Planilha não carregada"],
    ["Total em Documentos Previstos", metrics.totalPlanned, planned ? planned.label : "Planilha não carregada"],
    ["Total de documentos em comum", metrics.inCommon, "Presentes na Consulta Geral e em Documentos Previstos"],
    ["Somente na Consulta Geral", metrics.onlyGeneral, "Postados e não encontrados nas demais planilhas"],
    ["Somente em Documentos Previstos", metrics.onlyPlanned, "Previstos e não encontrados nas demais planilhas"],
    ["Total de documentos alocados", metrics.allocated, "Encontrados em Documentos Previstos"],
    ["Total de documentos não alocados", metrics.notAllocated, "Não encontrados em Documentos Previstos"],
    ["Total de divergências de status", metrics.divergences, "Status diferente entre as planilhas"],
    ["Presentes em todas as planilhas", metrics.inAllSheets, "Documentos encontrados em todas as bases carregadas"],
    ["Presentes em apenas algumas planilhas", metrics.partial, "Documentos ausentes em pelo menos uma base"],
    ["Ausentes na Consulta Geral", metrics.missingGeneral, "Pendentes de postagem no SIGEM"],
  ];
  for (const source of output.sources) {
    summaryRows.push([`Encontrados em ${source.label}`, source.total, `Planilha ${source.label}`]);
  }
  summaryRows.push(["Critério de comparação", MATCH_MODES[output.matchMode], "Regra usada para casar os documentos"]);
  summary.addRows(summaryRows);
  summary.eachRow((row, rowNumber) => {
    if (rowNumber <= 5) return;
    row.alignment = { vertical: "middle", wrapText: true };
    row.getCell(1).font = { bold: true, color: { argb: navy } };
    row.getCell(2).alignment = { horizontal: "left" };
    row.getCell(3).font = { size: 10, color: { argb: "52687B" } };
    if ((rowNumber - 5) % 2 === 0) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: pale } };
  });

  // Aba 2 — Resultado Consolidado
  const consolidatedColumns = [
    { header: "DOCUMENTO", key: "documento", width: 52 },
    { header: "STATUS CONSULTA GERAL", key: "statusGeral", width: 28 },
    { header: "EXISTE CONSULTA GERAL", key: "existeGeral", width: 22 },
    { header: "EXISTE DOCUMENTOS PREVISTOS", key: "existePrevistos", width: 26 },
    { header: "ALOCADO", key: "alocado", width: 14 },
    { header: "STATUS OUTRAS PLANILHAS", key: "statusOutras", width: 46 },
    { header: "OBSERVAÇÕES", key: "observacoes", width: 72 },
    { header: "PRESENÇA", key: "presenca", width: 24 },
    { header: "DATA CONSULTA GERAL", key: "dataGeral", width: 18 },
    { header: "INFORMAÇÕES COMPLEMENTARES", key: "complementos", width: 60 },
  ];
  for (const source of output.sources) {
    consolidatedColumns.push({ header: `EXISTE ${source.label.toUpperCase()}`, key: `existe_${source.id}`, width: 20 });
    consolidatedColumns.push({ header: `STATUS ${source.label.toUpperCase()}`, key: `status_${source.id}`, width: 26 });
  }

  const consolidatedRow = (row) => {
    const values = {
      documento: row.document,
      statusGeral: row.generalStatus,
      existeGeral: general ? (row.existsGeneral ? "Sim" : "Não") : "—",
      existePrevistos: planned ? (row.existsPlanned ? "Sim" : "Não") : "—",
      alocado: row.allocated,
      statusOutras: row.otherStatuses,
      observacoes: row.observations,
      presenca: row.presence.label,
      dataGeral: row.generalDate,
      complementos: row.complements,
    };
    for (const source of output.sources) {
      values[`existe_${source.id}`] = row.sources[source.id]?.present ? "Sim" : "Não";
      values[`status_${source.id}`] = row.sources[source.id]?.status || "";
    }
    return values;
  };

  addDataSheet("Resultado Consolidado", consolidatedColumns, output.rows.map(consolidatedRow), (row) => {
    const paint = (index) => {
      const cell = row.getCell(index);
      const value = norm(cell.value);
      if (value === "SIM") cell.font = { bold: true, color: { argb: teal } };
      else if (value === "NAO") cell.font = { bold: true, color: { argb: red } };
    };
    paint(3);
    paint(4);
    paint(5);
    for (let index = 11; index <= consolidatedColumns.length; index += 2) paint(index);
  });

  // Abas 3 e 4 — exclusivos de cada base
  const exclusiveColumns = [
    { header: "DOCUMENTO", key: "documento", width: 52 },
    { header: "STATUS", key: "status", width: 30 },
    { header: "DATA", key: "data", width: 18 },
    { header: "LINHA DE ORIGEM", key: "linha", width: 16 },
    { header: "AUSENTE EM", key: "ausente", width: 46 },
    { header: "OBSERVAÇÕES", key: "observacoes", width: 72 },
  ];
  const exclusiveRow = (row, sourceId) => ({
    documento: row.document,
    status: row.sources[sourceId]?.status || "",
    data: row.sources[sourceId]?.date || "",
    linha: row.sources[sourceId]?.rows.join(", ") || "",
    ausente: row.missingIn.map((source) => source.label).join(", "),
    observacoes: row.observations,
  });

  addDataSheet(
    "Somente Consulta Geral",
    exclusiveColumns,
    general ? output.rows.filter((row) => row.onlyGeneral).map((row) => exclusiveRow(row, general.id)) : [],
  );
  addDataSheet(
    "Somente Documentos Previstos",
    exclusiveColumns,
    planned ? output.rows.filter((row) => row.onlyPlanned).map((row) => exclusiveRow(row, planned.id)) : [],
  );

  // Aba 5 — Divergências
  const divergenceColumns = [{ header: "DOCUMENTO", key: "documento", width: 52 }];
  for (const source of output.sources) {
    divergenceColumns.push({ header: `STATUS ${source.label.toUpperCase()}`, key: `status_${source.id}`, width: 30 });
  }
  divergenceColumns.push({ header: "DIVERGÊNCIA", key: "divergencia", width: 72 });

  const divergences = output.rows.filter((row) => row.statusDivergence).map((row) => {
    const values = {
      documento: row.document,
      divergencia: row.divergentStatuses.map((entry) => `${entry.label}: ${entry.status}`).join(" × "),
    };
    for (const source of output.sources) values[`status_${source.id}`] = row.sources[source.id]?.status || "";
    return values;
  });
  addDataSheet("Divergências", divergenceColumns, divergences, (row) => {
    row.getCell(divergenceColumns.length).font = { bold: true, color: { argb: amber } };
  });

  if (others.length) {
    const extraColumns = [
      { header: "PLANILHA", key: "planilha", width: 32 },
      { header: "DOCUMENTO", key: "documento", width: 52 },
      { header: "STATUS", key: "status", width: 30 },
      { header: "DATA", key: "data", width: 18 },
      { header: "INFORMAÇÕES COMPLEMENTARES", key: "complementos", width: 62 },
      { header: "TAMBÉM PRESENTE EM", key: "presente", width: 46 },
    ];
    const extraRows = others.flatMap((source) => output.rows
      .filter((row) => row.sources[source.id]?.present)
      .map((row) => ({
        planilha: source.label,
        documento: row.document,
        status: row.sources[source.id].status,
        data: row.sources[source.id].date,
        complementos: row.sources[source.id].extrasText,
        presente: row.presentIn.filter((item) => item.id !== source.id).map((item) => item.label).join(", "),
      })));
    addDataSheet("Planilhas Adicionais", extraColumns, extraRows);
  }

  return workbook;
}
