import test from "node:test";
import assert from "node:assert/strict";
import { crossReference } from "../app/lib/crosscheck-engine.js";
import { buildCrossWorkbook, columnLetter, crossReportSheetNames } from "../app/lib/cross-report.js";

// PNG 1x1 transparente — a marca real é desenhada no navegador e injetada aqui.
const LOGO = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

function sampleOutput() {
  const record = (document, status = "", extra = {}) => ({ document, status, date: "", extras: {}, rowNumber: 2, ...extra });
  return crossReference([
    {
      id: "general",
      role: "general",
      label: "Consulta Geral",
      records: [
        record("DOC-001", "Aprovado", { date: "10/02/2026", rowNumber: 6 }),
        record("DOC-002", "Em análise", { rowNumber: 7 }),
        record("DOC-003", "Aprovado", { rowNumber: 8 }),
      ],
    },
    {
      id: "planned",
      role: "planned",
      label: "Documentos Previstos",
      records: [record("DOC-001"), record("DOC-002"), record("DOC-900")],
    },
    {
      id: "extra-1",
      role: "extra",
      label: "Checklist de Campo",
      records: [record("DOC-001", "Reprovado", { extras: { Disciplina: "Elétrica" } })],
    },
  ]);
}

function values(sheet, rowNumber) {
  const row = sheet.getRow(rowNumber).values;
  return (Array.isArray(row) ? row.slice(1) : []).map((value) => (value === undefined || value === null ? "" : String(value)));
}

test("columnLetter cobre colunas de uma e duas letras", () => {
  assert.equal(columnLetter(0), "A");
  assert.equal(columnLetter(25), "Z");
  assert.equal(columnLetter(26), "AA");
  assert.equal(columnLetter(27), "AB");
});

test("o relatório traz as abas exigidas, na ordem, nomeadas pelas planilhas do usuário", () => {
  const output = sampleOutput();
  const workbook = buildCrossWorkbook(output, { logoBase64: LOGO });
  const names = workbook.worksheets.map((sheet) => sheet.name);
  assert.deepEqual(names, [
    "Resumo Executivo",
    "Resultado Consolidado",
    "Somente Consulta Geral",
    "Somente Documentos Previstos",
    "Exclusivos por planilha",
    "Divergências",
    "Planilhas Adicionais",
  ]);
  assert.deepEqual(crossReportSheetNames(output), names);
});

test("Aba 1 — Resumo Executivo lista os indicadores consolidados", () => {
  const workbook = buildCrossWorkbook(sampleOutput(), { logoBase64: LOGO });
  const summary = workbook.getWorksheet("Resumo Executivo");
  assert.deepEqual(values(summary, 5), ["Indicador", "Quantidade", "Detalhe"]);

  const indicators = new Map();
  summary.eachRow((row, rowNumber) => {
    if (rowNumber <= 5) return;
    indicators.set(String(row.getCell(1).value), row.getCell(2).value);
  });

  assert.equal(indicators.get("Total de documentos analisados"), 4);
  assert.equal(indicators.get("Total em Consulta Geral"), 3);
  assert.equal(indicators.get("Total em Documentos Previstos"), 3);
  assert.equal(indicators.get("Exclusivos de uma única planilha"), 2);
  assert.equal(indicators.get("Total de documentos em comum"), 2);
  assert.equal(indicators.get("Total de documentos alocados"), 3);
  assert.equal(indicators.get("Total de documentos não alocados"), 1);
  assert.equal(indicators.get("Total de divergências de status"), 1);
  assert.equal(indicators.get("Encontrados em Checklist de Campo"), 1);
});

test("Aba 2 — Resultado Consolidado usa as colunas pedidas e uma coluna por planilha", () => {
  const workbook = buildCrossWorkbook(sampleOutput(), { logoBase64: LOGO });
  const sheet = workbook.getWorksheet("Resultado Consolidado");
  const headers = values(sheet, 5);

  assert.deepEqual(headers.slice(0, 7), [
    "DOCUMENTO",
    "STATUS CONSULTA GERAL",
    "EXISTE CONSULTA GERAL",
    "EXISTE DOCUMENTOS PREVISTOS",
    "ALOCADO",
    "STATUS OUTRAS PLANILHAS",
    "OBSERVAÇÕES",
  ]);
  assert.ok(headers.includes("STATUS CHECKLIST DE CAMPO"));
  assert.ok(headers.includes("EXISTE CHECKLIST DE CAMPO"));
  assert.equal(sheet.rowCount, 9); // 5 de cabeçalho + 4 documentos

  const first = values(sheet, 6);
  assert.equal(first[0], "DOC-001");
  assert.equal(first[1], "Aprovado");
  assert.equal(first[2], "Sim");
  assert.equal(first[3], "Sim");
  assert.equal(first[4], "Sim");
  assert.equal(first[5], "Checklist de Campo: Reprovado");
  assert.match(first[6], /Divergência de status/);
  assert.equal(first[8], "10/02/2026");

  const missing = values(sheet, 9); // DOC-900, só nos previstos
  assert.equal(missing[0], "DOC-900");
  assert.equal(missing[2], "Não");
  assert.equal(missing[4], "Sim");
  assert.match(missing[6], /Não localizado em Consulta Geral/);
});

test("Abas 3 e 4 — listas exclusivas de cada base", () => {
  const workbook = buildCrossWorkbook(sampleOutput(), { logoBase64: LOGO });
  const onlyGeneral = workbook.getWorksheet("Somente Consulta Geral");
  const onlyPlanned = workbook.getWorksheet("Somente Documentos Previstos");

  assert.deepEqual(values(onlyGeneral, 5).slice(0, 4), ["DOCUMENTO", "STATUS", "DATA", "LINHA DE ORIGEM"]);
  assert.equal(onlyGeneral.rowCount, 6);
  assert.deepEqual(values(onlyGeneral, 6).slice(0, 4), ["DOC-003", "Aprovado", "", "8"]);

  assert.equal(onlyPlanned.rowCount, 6);
  assert.equal(values(onlyPlanned, 6)[0], "DOC-900");
});

test("Aba 5 — Divergências traz uma coluna de status por planilha", () => {
  const workbook = buildCrossWorkbook(sampleOutput(), { logoBase64: LOGO });
  const sheet = workbook.getWorksheet("Divergências");

  assert.deepEqual(values(sheet, 5), [
    "DOCUMENTO",
    "STATUS CONSULTA GERAL",
    "STATUS DOCUMENTOS PREVISTOS",
    "STATUS CHECKLIST DE CAMPO",
    "DIVERGÊNCIA",
  ]);
  assert.equal(sheet.rowCount, 6);
  assert.deepEqual(values(sheet, 6), [
    "DOC-001",
    "Aprovado",
    "",
    "Reprovado",
    "Consulta Geral: Aprovado × Checklist de Campo: Reprovado",
  ]);
});

test("o relatório é gravável como .xlsx", async () => {
  const workbook = buildCrossWorkbook(sampleOutput(), { logoBase64: LOGO });
  const buffer = await workbook.xlsx.writeBuffer();
  assert.ok(buffer.byteLength > 5000);
  assert.deepEqual([...new Uint8Array(buffer).slice(0, 2)], [0x50, 0x4b]); // assinatura ZIP
});

test("sem planilhas adicionais o relatório dispensa a aba extra", () => {
  const output = crossReference([
    { id: "general", role: "general", label: "Consulta Geral", records: [{ document: "DOC-001", status: "Aprovado", date: "", extras: {}, rowNumber: 2 }] },
    { id: "planned", role: "planned", label: "Documentos Previstos", records: [{ document: "DOC-002", status: "", date: "", extras: {}, rowNumber: 2 }] },
  ]);
  const workbook = buildCrossWorkbook(output, { logoBase64: LOGO });
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), [
    "Resumo Executivo",
    "Resultado Consolidado",
    "Somente Consulta Geral",
    "Somente Documentos Previstos",
    "Exclusivos por planilha",
    "Divergências",
  ]);
});

test("os títulos seguem o nome dado às planilhas, não um processo específico", () => {
  const record = (document, status = "") => ({ document, status, date: "", extras: {}, rowNumber: 2 });
  const output = crossReference([
    { id: "a", role: "general", label: "Base Comercial", records: [record("PED-1", "Aberto")] },
    { id: "b", role: "planned", label: "Carteira 2026", records: [record("PED-1"), record("PED-2")] },
  ]);
  const workbook = buildCrossWorkbook(output, { logoBase64: LOGO });

  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name).slice(2, 4), [
    "Somente Base Comercial",
    "Somente Carteira 2026",
  ]);
  const headers = values(workbook.getWorksheet("Resultado Consolidado"), 5);
  assert.deepEqual(headers.slice(0, 5), [
    "DOCUMENTO",
    "STATUS BASE COMERCIAL",
    "EXISTE BASE COMERCIAL",
    "EXISTE CARTEIRA 2026",
    "ALOCADO",
  ]);
  const indicators = [];
  workbook.getWorksheet("Resumo Executivo").eachRow((row, rowNumber) => {
    if (rowNumber > 5) indicators.push(String(row.getCell(1).value));
  });
  assert.ok(indicators.includes("Total em Base Comercial"));
  assert.ok(indicators.includes("Ausentes em Base Comercial"));
  assert.ok(indicators.includes("Total em Carteira 2026"));
  assert.ok(!indicators.some((label) => /Consulta Geral|Previstos|SIGEM|SGP/i.test(label)));
});

test("sem papéis atribuídos o relatório é puramente genérico", () => {
  const record = (document, status = "") => ({ document, status, date: "", extras: {}, rowNumber: 2 });
  const output = crossReference([
    { id: "a", role: "extra", label: "Planilha 1", records: [record("X-1", "Novo"), record("X-2", "Novo")] },
    { id: "b", role: "extra", label: "Planilha 2", records: [record("X-1", "Antigo")] },
    { id: "c", role: "extra", label: "Planilha 3", records: [record("X-3", "Novo")] },
  ]);
  const workbook = buildCrossWorkbook(output, { logoBase64: LOGO });

  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), [
    "Resumo Executivo",
    "Resultado Consolidado",
    "Exclusivos por planilha",
    "Divergências",
    "Planilhas Adicionais",
  ]);

  const headers = values(workbook.getWorksheet("Resultado Consolidado"), 5);
  assert.deepEqual(headers.slice(0, 4), ["DOCUMENTO", "STATUS POR PLANILHA", "OBSERVAÇÕES", "PRESENÇA"]);
  assert.ok(!headers.includes("ALOCADO"));
  assert.ok(!headers.some((header) => header.startsWith("EXISTE CONSULTA")));

  const exclusives = workbook.getWorksheet("Exclusivos por planilha");
  assert.equal(exclusives.rowCount, 7); // 5 de cabeçalho + X-2 e X-3
  assert.deepEqual(values(exclusives, 6).slice(0, 2), ["Planilha 1", "X-2"]);
  assert.deepEqual(values(exclusives, 7).slice(0, 2), ["Planilha 3", "X-3"]);

  assert.equal(output.metrics.exclusive, 2);
  assert.equal(output.metrics.allocated, 0);
  assert.equal(output.metrics.divergences, 1);
});
