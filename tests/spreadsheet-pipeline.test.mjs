import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx";
import { profileRows, recordFromRow } from "../app/lib/sheet-profile.js";
import { crossReference } from "../app/lib/crosscheck-engine.js";
import { buildCrossWorkbook } from "../app/lib/cross-report.js";

const LOGO = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

/** Grava um .xlsx real e o lê de volta como o navegador faz. */
function roundTrip(rows, sheetName = "Dados") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  const buffer = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  const parsed = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: false });
  const sheet = parsed.Sheets[parsed.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, blankrows: true });
  const profile = profileRows(matrix, { columnLabel: (index) => `Coluna ${XLSX.utils.encode_col(range.s.c + index)}` });
  return { name: parsed.SheetNames[0], rows: matrix, startRow: range.s.r, ...profile };
}

/** Repete o laço de leitura da tela sobre uma aba já perfilada. */
function recordsOf(profile) {
  const records = [];
  for (let index = profile.dataStart; index < profile.rows.length; index += 1) {
    const record = recordFromRow(profile, profile.rows[index] || [], index);
    if (record) records.push(record);
  }
  return records;
}

const CONSULTA_GERAL = [
  ["Consulta Geral — SIGEM"],
  ["Emitido em:", "12/08/2026"],
  [],
  ["Número do Documento", "Status", "Data de Postagem", "Disciplina"],
  ["ET-5290.00-22000-912-1LV-001", "Aprovado", "10/02/2026", "Elétrica"],
  ["ET-5290.00-22000-912-1LV-002", "Em análise", "11/02/2026", "Instrumentação"],
  ["ET-5290.00-22000-912-1LV-003", "Aprovado", "12/02/2026", "Mecânica"],
  ["", "", "", ""],
];

const DOCUMENTOS_PREVISTOS = [
  ["Código do Documento", "Responsável"],
  ["ET-5290.00-22000-912-1LV-001", "Equipe A"],
  ["ET-5290.00-22000-912-1LV-002", "Equipe B"],
  ["ET-5290.00-22000-912-1LV-900", "Equipe C"],
];

const CHECKLIST = [
  ["Arquivo", "Situação"],
  ["ET-5290.00-22000-912-1LV-001 - Rev. P.pdf", "Reprovado"],
  ["ET-5290.00-22000-912-1LV-003.pdf", "Aprovado"],
];

test("perfila um .xlsx real com título e linhas soltas antes do cabeçalho", () => {
  const profile = roundTrip(CONSULTA_GERAL);
  assert.equal(profile.headerIndex, 3);
  assert.equal(profile.dataStart, 4);
  assert.equal(profile.rowCount, 3);
  assert.equal(profile.headers[profile.mapping.document], "Número do Documento");
  assert.equal(profile.headers[profile.mapping.status], "Status");
  assert.equal(profile.headers[profile.mapping.date], "Data de Postagem");
});

test("extrai registros com número da linha de origem e ignora linhas vazias", () => {
  const profile = roundTrip(CONSULTA_GERAL);
  profile.mapping.extras = [3];
  const records = recordsOf(profile);

  assert.equal(records.length, 3);
  assert.deepEqual(records[0], {
    document: "ET-5290.00-22000-912-1LV-001",
    status: "Aprovado",
    date: "10/02/2026",
    extras: { Disciplina: "Elétrica" },
    rowNumber: 5,
  });
  assert.equal(records[2].rowNumber, 7);
});

test("cruza três planilhas reais e gera o relatório Excel completo", async () => {
  const geral = roundTrip(CONSULTA_GERAL);
  geral.mapping.extras = [3];
  const previstos = roundTrip(DOCUMENTOS_PREVISTOS);
  previstos.mapping.extras = [1];
  const checklist = roundTrip(CHECKLIST);

  assert.equal(previstos.headers[previstos.mapping.document], "Código do Documento");
  assert.equal(checklist.headers[checklist.mapping.document], "Arquivo");
  assert.equal(checklist.headers[checklist.mapping.status], "Situação");

  const output = crossReference([
    { id: "general", role: "general", label: "Consulta Geral", records: recordsOf(geral) },
    { id: "planned", role: "planned", label: "Documentos Previstos", records: recordsOf(previstos) },
    { id: "extra-1", role: "extra", label: "Checklist de Campo", records: recordsOf(checklist) },
  ]);

  // O modo inteligente casa "…-001 - Rev. P.pdf" com "…-001".
  assert.equal(output.metrics.total, 4);
  assert.equal(output.metrics.inCommon, 2);
  assert.equal(output.metrics.allocated, 3);
  assert.equal(output.metrics.notAllocated, 1);
  assert.equal(output.metrics.divergences, 1);
  assert.equal(output.metrics.onlyGeneral, 0);
  assert.equal(output.metrics.onlyPlanned, 1);

  const rows = new Map(output.rows.map((row) => [row.document, row]));
  const primeiro = rows.get("ET-5290.00-22000-912-1LV-001");
  assert.equal(primeiro.presence.kind, "all");
  assert.equal(primeiro.allocated, "Sim");
  assert.equal(primeiro.generalStatus, "Aprovado");
  assert.equal(primeiro.otherStatuses, "Checklist de Campo: Reprovado");
  assert.match(primeiro.complements, /Consulta Geral — Disciplina: Elétrica/);
  assert.match(primeiro.complements, /Documentos Previstos — Responsável: Equipe A/);
  assert.equal(rows.get("ET-5290.00-22000-912-1LV-003").allocated, "Não");
  assert.equal(rows.get("ET-5290.00-22000-912-1LV-900").existsGeneral, false);

  const workbook = buildCrossWorkbook(output, { logoBase64: LOGO });
  const consolidated = workbook.getWorksheet("Resultado Consolidado");
  assert.equal(consolidated.rowCount, 9);
  const buffer = await workbook.xlsx.writeBuffer();
  assert.deepEqual([...new Uint8Array(buffer).slice(0, 2)], [0x50, 0x4b]);
});

test("no modo exato a variação com revisão e extensão vira outro documento", () => {
  const geral = roundTrip(CONSULTA_GERAL);
  const checklist = roundTrip(CHECKLIST);
  const output = crossReference(
    [
      { id: "general", role: "general", label: "Consulta Geral", records: recordsOf(geral) },
      { id: "extra-1", role: "extra", label: "Checklist de Campo", records: recordsOf(checklist) },
    ],
    { matchMode: "exact" },
  );
  assert.equal(output.metrics.total, 5);
  assert.equal(output.metrics.inAllSheets, 0);
});
