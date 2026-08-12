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


function safeSheetName(name, used) {
  const base = name.replace(/[[\]:*?/\\]/g, " ").slice(0, 31).trim() || "Planilha";
  if (!used || !used.has(base)) {
    used?.add(base);
    return base;
  }
  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${base.slice(0, 31 - String(suffix).length - 1)} ${suffix}`;
    if (!used.has(candidate)) {
      used.add(candidate);
      return candidate;
    }
  }
}

/**
 * Nomes das abas que o relatório terá para um dado cruzamento. As abas de
 * referência e de alocação só existem quando esses papéis foram atribuídos, e
 * levam o nome que o usuário deu à planilha.
 */
export function crossReportSheetNames(output) {
  const names = ["Resumo Executivo", "Resultado Consolidado"];
  if (output.generalLabel) names.push(`Somente ${output.generalLabel}`);
  if (output.plannedLabel) names.push(`Somente ${output.plannedLabel}`);
  names.push("Exclusivos por planilha", "Divergências");
  const extras = (output.sources || []).filter((source) => source.id !== output.generalId && source.id !== output.plannedId);
  if (extras.length) names.push("Planilhas Adicionais");
  const used = new Set();
  return names.map((name) => safeSheetName(name, used));
}

/**
 * Monta o relatório Excel do cruzamento. As colunas e as abas se ajustam às
 * planilhas carregadas: nada é fixo em uma base específica, e os títulos usam o
 * nome que o usuário deu a cada planilha.
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

  const usedNames = new Set();
  const addDataSheet = (name, columns, rows, highlight) => {
    const sheet = workbook.addWorksheet(safeSheetName(name, usedNames), { views: [{ state: "frozen", ySplit: 5, xSplit: 1 }] });
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
  const summary = workbook.addWorksheet(safeSheetName("Resumo Executivo", usedNames), { views: [{ state: "frozen", ySplit: 5 }] });
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
    ["Planilhas cruzadas", output.sources.length, output.sources.map((source) => source.label).join(", ")],
    ["Presentes em todas as planilhas", metrics.inAllSheets, "Encontrados em todas as bases carregadas"],
    ["Presentes em apenas algumas planilhas", metrics.partial, "Ausentes em pelo menos uma base"],
    ["Exclusivos de uma única planilha", metrics.exclusive, "Encontrados em uma base e em nenhuma outra"],
    ["Total de divergências de status", metrics.divergences, "Status diferente entre as planilhas"],
  ];
  if (general) {
    summaryRows.push(
      [`Total em ${general.label}`, metrics.totalGeneral, `Documentos encontrados em ${general.label}`],
      [`Ausentes em ${general.label}`, metrics.missingGeneral, `Não encontrados em ${general.label}`],
      [`Somente em ${general.label}`, metrics.onlyGeneral, "Sem correspondência nas demais planilhas"],
    );
  }
  if (planned) {
    summaryRows.push(
      [`Total em ${planned.label}`, metrics.totalPlanned, `Documentos encontrados em ${planned.label}`],
      [`Somente em ${planned.label}`, metrics.onlyPlanned, "Sem correspondência nas demais planilhas"],
      ["Total de documentos alocados", metrics.allocated, `Encontrados em ${planned.label}`],
      ["Total de documentos não alocados", metrics.notAllocated, `Não encontrados em ${planned.label}`],
    );
  }
  if (general && planned) {
    summaryRows.push(["Total de documentos em comum", metrics.inCommon, `Presentes em ${general.label} e em ${planned.label}`]);
  }
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
  const consolidatedColumns = [{ header: "DOCUMENTO", key: "documento", width: 52 }];
  if (general) {
    consolidatedColumns.push(
      { header: `STATUS ${general.label.toUpperCase()}`, key: "statusGeral", width: 28 },
      { header: `EXISTE ${general.label.toUpperCase()}`, key: "existeGeral", width: 22 },
    );
  }
  if (planned) {
    consolidatedColumns.push(
      { header: `EXISTE ${planned.label.toUpperCase()}`, key: "existePrevistos", width: 26 },
      { header: "ALOCADO", key: "alocado", width: 14 },
    );
  }
  consolidatedColumns.push(
    { header: general || planned ? "STATUS OUTRAS PLANILHAS" : "STATUS POR PLANILHA", key: "statusOutras", width: 46 },
    { header: "OBSERVAÇÕES", key: "observacoes", width: 72 },
    { header: "PRESENÇA", key: "presenca", width: 24 },
  );
  if (general) consolidatedColumns.push({ header: `DATA ${general.label.toUpperCase()}`, key: "dataGeral", width: 18 });
  consolidatedColumns.push({ header: "INFORMAÇÕES COMPLEMENTARES", key: "complementos", width: 60 });

  const perSourceStart = consolidatedColumns.length + 1;
  for (const source of output.sources) {
    consolidatedColumns.push({ header: `EXISTE ${source.label.toUpperCase()}`, key: `existe_${source.id}`, width: 20 });
    consolidatedColumns.push({ header: `STATUS ${source.label.toUpperCase()}`, key: `status_${source.id}`, width: 26 });
  }

  const consolidatedRow = (row) => {
    const values = {
      documento: row.document,
      statusGeral: row.generalStatus,
      existeGeral: row.existsGeneral ? "Sim" : "Não",
      existePrevistos: row.existsPlanned ? "Sim" : "Não",
      alocado: row.allocated,
      statusOutras: general || planned
        ? row.otherStatuses
        : row.statusEntries.map((entry) => `${entry.label}: ${entry.status}`).join("; "),
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

  const simNaoColumns = consolidatedColumns
    .map((column, index) => ({ column, position: index + 1 }))
    .filter(({ column, position }) => position >= perSourceStart
      ? column.key.startsWith("existe_")
      : ["existeGeral", "existePrevistos", "alocado"].includes(column.key))
    .map(({ position }) => position);

  addDataSheet("Resultado Consolidado", consolidatedColumns, output.rows.map(consolidatedRow), (row) => {
    for (const position of simNaoColumns) {
      const cell = row.getCell(position);
      const value = norm(cell.value);
      if (value === "SIM") cell.font = { bold: true, color: { argb: teal } };
      else if (value === "NAO") cell.font = { bold: true, color: { argb: red } };
    }
  });

  // Abas exclusivas: por papel, quando atribuído, e sempre a lista geral
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

  if (general) {
    addDataSheet(
      `Somente ${general.label}`,
      exclusiveColumns,
      output.rows.filter((row) => row.onlyGeneral).map((row) => exclusiveRow(row, general.id)),
    );
  }
  if (planned) {
    addDataSheet(
      `Somente ${planned.label}`,
      exclusiveColumns,
      output.rows.filter((row) => row.onlyPlanned).map((row) => exclusiveRow(row, planned.id)),
    );
  }

  addDataSheet(
    "Exclusivos por planilha",
    [{ header: "PLANILHA", key: "planilha", width: 32 }, ...exclusiveColumns],
    output.rows
      .filter((row) => row.exclusiveIn)
      .map((row) => ({ planilha: row.exclusiveIn.label, ...exclusiveRow(row, row.exclusiveIn.id) })),
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
