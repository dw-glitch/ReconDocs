import test from "node:test";
import assert from "node:assert/strict";
import {
  crossKey,
  crossReference,
  filterCrossRows,
} from "../app/lib/crosscheck-engine.js";
import {
  detectHeaderRow,
  headerAliasScore,
  looksLikeDate,
  looksLikeIdentifier,
  profileRows,
} from "../app/lib/sheet-profile.js";

function source(id, role, label, records) {
  return {
    id,
    role,
    label,
    records: records.map((record, index) => ({
      status: "",
      date: "",
      extras: {},
      rowNumber: index + 2,
      ...(typeof record === "string" ? { document: record } : record),
    })),
  };
}

const general = (records) => source("general", "general", "Consulta Geral", records);
const planned = (records) => source("planned", "planned", "Documentos Previstos", records);
const extra = (records, id = "extra-1", label = "Planilha 3") => source(id, "extra", label, records);

test("crossKey aproxima o mesmo documento com extensão e revisão no modo inteligente", () => {
  assert.equal(
    crossKey("ET-5290.00-22000-912-1LV-001 - Rev. P.pdf"),
    crossKey("ET-5290.00-22000-912-1LV-001"),
  );
  assert.notEqual(
    crossKey("ET-5290.00-22000-912-1LV-001 - Rev. P.pdf", "exact"),
    crossKey("ET-5290.00-22000-912-1LV-001", "exact"),
  );
});

test("cruza presença, alocação e contagens entre Consulta Geral e Documentos Previstos", () => {
  const output = crossReference([
    general([
      { document: "DOC-001", status: "Aprovado" },
      { document: "DOC-002", status: "Em análise" },
      { document: "DOC-003", status: "Aprovado" },
    ]),
    planned([{ document: "DOC-001" }, { document: "DOC-002" }, { document: "DOC-900" }]),
  ]);

  assert.equal(output.metrics.total, 4);
  assert.equal(output.metrics.totalGeneral, 3);
  assert.equal(output.metrics.totalPlanned, 3);
  assert.equal(output.metrics.inCommon, 2);
  assert.equal(output.metrics.onlyGeneral, 1);
  assert.equal(output.metrics.onlyPlanned, 1);
  assert.equal(output.metrics.allocated, 3);
  assert.equal(output.metrics.notAllocated, 1);
  assert.equal(output.metrics.missingGeneral, 1);

  const rowByCode = new Map(output.rows.map((row) => [row.document, row]));
  assert.equal(rowByCode.get("DOC-001").allocated, "Sim");
  assert.equal(rowByCode.get("DOC-003").allocated, "Não");
  assert.equal(rowByCode.get("DOC-003").existsGeneral, true);
  assert.equal(rowByCode.get("DOC-900").existsGeneral, false);
  assert.equal(rowByCode.get("DOC-001").generalStatus, "Aprovado");
});

test("aponta divergência de status somente quando as planilhas discordam", () => {
  const output = crossReference([
    general([{ document: "DOC-001", status: "Aprovado" }, { document: "DOC-002", status: "Aprovado" }]),
    planned([{ document: "DOC-001" }, { document: "DOC-002" }]),
    extra([{ document: "DOC-001", status: "Reprovado" }, { document: "DOC-002", status: "aprovado " }]),
  ]);

  assert.equal(output.metrics.divergences, 1);
  const divergent = output.rows.find((row) => row.statusDivergence);
  assert.equal(divergent.document, "DOC-001");
  assert.deepEqual(divergent.divergentStatuses.map((entry) => entry.status), ["Aprovado", "Reprovado"]);
  assert.equal(divergent.otherStatuses, "Planilha 3: Reprovado");
  assert.match(divergent.observations, /Divergência de status/);
});

test("classifica presença total, parcial e única entre três planilhas", () => {
  const output = crossReference([
    general([{ document: "DOC-001" }, { document: "DOC-002" }]),
    planned([{ document: "DOC-001" }]),
    extra([{ document: "DOC-001" }, { document: "DOC-003" }]),
  ]);

  const rows = new Map(output.rows.map((row) => [row.document, row]));
  assert.equal(rows.get("DOC-001").presence.kind, "all");
  assert.equal(rows.get("DOC-002").presence.kind, "single");
  assert.equal(rows.get("DOC-003").presence.kind, "single");
  assert.deepEqual(rows.get("DOC-002").missingIn.map((item) => item.label), ["Documentos Previstos", "Planilha 3"]);
  assert.equal(output.metrics.inAllSheets, 1);
  assert.equal(output.metrics.partial, 2);
  assert.equal(rows.get("DOC-002").onlyGeneral, true);
  assert.equal(rows.get("DOC-003").onlyGeneral, false);
});

test("mantém informações complementares e linhas de origem por planilha", () => {
  const output = crossReference([
    general([{ document: "DOC-001", status: "Aprovado", date: "10/02/2026", extras: { Disciplina: "Elétrica" }, rowNumber: 7 }]),
    extra([{ document: "DOC-001", extras: { Responsável: "Equipe A" }, rowNumber: 12 }]),
  ]);

  const [row] = output.rows;
  assert.equal(row.generalDate, "10/02/2026");
  assert.deepEqual(row.sources.general.rows, [7]);
  assert.match(row.complements, /Consulta Geral — Disciplina: Elétrica/);
  assert.match(row.complements, /Planilha 3 — Responsável: Equipe A/);
});

test("agrupa linhas repetidas da mesma planilha e registra a duplicidade", () => {
  const output = crossReference([
    general([
      { document: "DOC-001", status: "Aprovado", rowNumber: 4 },
      { document: "doc-001", status: "Aprovado", rowNumber: 9 },
    ]),
    planned([{ document: "DOC-001" }]),
  ]);

  assert.equal(output.rows.length, 1);
  assert.equal(output.rows[0].sources.general.count, 2);
  assert.equal(output.rows[0].sources.general.duplicated, true);
  assert.match(output.rows[0].observations, /Linhas repetidas/);
});

test("sem planilha de previstos a alocação fica indefinida em vez de negativa", () => {
  const output = crossReference([general([{ document: "DOC-001" }]), extra([{ document: "DOC-001" }])]);
  assert.equal(output.rows[0].allocation.kind, "unknown");
  assert.equal(output.metrics.allocated, 0);
  assert.equal(output.metrics.notAllocated, 0);
});

test("filtros e busca selecionam os documentos esperados", () => {
  const output = crossReference([
    general([{ document: "DOC-001", status: "Aprovado" }, { document: "DOC-002", status: "Aprovado" }]),
    planned([{ document: "DOC-001" }, { document: "DOC-900" }]),
    extra([{ document: "DOC-001", status: "Reprovado" }]),
  ]);

  const count = (filter) => filterCrossRows(output.rows, filter, "").length;
  assert.equal(count("all"), 3);
  assert.equal(count("in_all"), 1);
  assert.equal(count("only_general"), 1);
  assert.equal(count("only_planned"), 1);
  assert.equal(count("not_allocated"), 1);
  assert.equal(count("missing_general"), 1);
  assert.equal(count("divergent"), 1);
  assert.equal(filterCrossRows(output.rows, "all", "doc-900").length, 1);
  assert.equal(filterCrossRows(output.rows, "all", "reprovado").length, 1);
});

test("detecta cabeçalho, coluna de documento, status e data em layout com títulos soltos", () => {
  const rows = [
    ["Relatório de documentos", "", ""],
    ["Filtro:", "Unidade 32", ""],
    [],
    ["Código do Documento", "Status", "Data de Postagem"],
    ["ET-5290.00-22000-912-1LV-001", "Aprovado", "10/02/2026"],
    ["ET-5290.00-22000-912-1LV-002", "Em análise", "11/02/2026"],
    ["ET-5290.00-22000-912-1LV-003", "Aprovado", "12/02/2026"],
  ];
  assert.equal(detectHeaderRow(rows), 3);
  const profile = profileRows(rows);
  assert.equal(profile.dataStart, 4);
  assert.equal(profile.rowCount, 3);
  assert.deepEqual(profile.mapping.document, 0);
  assert.deepEqual(profile.mapping.status, 1);
  assert.deepEqual(profile.mapping.date, 2);
  assert.deepEqual(profile.headers, ["Código do Documento", "Status", "Data de Postagem"]);
});

test("mapeia a coluna do documento mesmo sem cabeçalho reconhecido", () => {
  const rows = [
    ["Item", "Referência", "Situação"],
    [1, "5900.0018047.05.2-ABC-CV-GER-0001", "Aprovado"],
    [2, "5900.0018047.05.2-ABC-CV-GER-0002", "Aprovado"],
    [3, "5900.0018047.05.2-ABC-CV-GER-0003", "Pendente"],
  ];
  const profile = profileRows(rows);
  assert.equal(profile.mapping.document, 1);
  assert.equal(profile.mapping.status, 2);
});

test("heurísticas auxiliares de leitura de planilha", () => {
  assert.equal(headerAliasScore("Número do Documento", "document"), 12);
  assert.equal(headerAliasScore("Situação", "status"), 12);
  assert.equal(headerAliasScore("Peso", "document"), 0);
  assert.equal(looksLikeIdentifier("ET-5290.00-22000-912-1LV-001"), true);
  assert.equal(looksLikeIdentifier("1.234,56"), false);
  assert.equal(looksLikeIdentifier("10/02/2026"), false);
  assert.equal(looksLikeDate("10/02/2026"), true);
  assert.equal(looksLikeDate(new Date("2026-02-10T00:00:00Z")), true);
  assert.equal(looksLikeDate("Aprovado"), false);
});

test("o cruzamento não repete a leitura para bases grandes", () => {
  const size = 20000;
  const generalRecords = Array.from({ length: size }, (_, index) => ({ document: `DOC-${index}`, status: index % 2 ? "Aprovado" : "Pendente" }));
  const plannedRecords = Array.from({ length: size / 2 }, (_, index) => ({ document: `DOC-${index * 2}` }));
  const output = crossReference([general(generalRecords), planned(plannedRecords)]);
  assert.equal(output.metrics.total, size);
  assert.equal(output.metrics.allocated, size / 2);
  assert.equal(output.metrics.notAllocated, size / 2);
});
